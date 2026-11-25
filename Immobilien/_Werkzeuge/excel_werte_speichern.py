"""Berechnete Werte als Cache in eine openpyxl-Datei schreiben.

Damit zeigen Zellen und Diagramme auch ohne Neuberechnung Werte (Excel-Schutzansicht,
Handy-Vorschau, Numbers, Google Sheets). Die Formeln bleiben unverändert.

Aufruf: python excel_werte_speichern.py EINGABE.xlsx AUSGABE.xlsx
Benötigt: LibreOffice (soffice), openpyxl, lxml.
"""
import datetime
import glob
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

from lxml import etree
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries

SRC, OUT = sys.argv[1], sys.argv[2]
RECALC = (glob.glob("/root/.claude/skills/synced/*/xlsx/scripts/recalc.py") or [None])[0]

tmpdir = tempfile.mkdtemp()
calc = os.path.join(tmpdir, "calc.xlsx")
shutil.copy(SRC, calc)
if RECALC:
    res = subprocess.run(["python", RECALC, calc, "150"], capture_output=True, text=True)
    info = json.loads(res.stdout)
    assert info.get("status") == "success" and info["total_errors"] == 0, info
else:  # Fallback: direkt mit LibreOffice neu berechnen
    subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", tmpdir + "/out", calc], check=True)
    calc = os.path.join(tmpdir, "out", "calc.xlsx")
    info = {"total_formulas": "?"}
vals = load_workbook(calc, data_only=True)

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
M = "{%s}" % NS["m"]
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"

zin = zipfile.ZipFile(SRC)
files = {n: zin.read(n) for n in zin.namelist()}

wbx = etree.fromstring(files["xl/workbook.xml"])
rels = etree.fromstring(files["xl/_rels/workbook.xml.rels"])
rid2target = {r.get("Id"): r.get("Target") for r in rels}
sheet_files = {}
for s in wbx.find("m:sheets", NS):
    t = rid2target[s.get("{%s}id" % NS["r"])].lstrip("/")
    sheet_files[s.get("name")] = t if t.startswith("xl/") else "xl/" + t

n_cells = 0
for name, path in sheet_files.items():
    root = etree.fromstring(files[path])
    ws = vals[name]
    for c in root.iter(M + "c"):
        if c.find(M + "f") is None:
            continue
        v = ws[c.get("r")].value
        old = c.find(M + "v")
        if old is not None:
            c.remove(old)
        if v is None:
            continue
        ve = etree.SubElement(c, M + "v")
        if isinstance(v, bool):
            c.set("t", "b")
            ve.text = "1" if v else "0"
        elif isinstance(v, (int, float)):
            c.attrib.pop("t", None)
            ve.text = repr(float(v)) if isinstance(v, float) else str(v)
        elif isinstance(v, (datetime.datetime, datetime.date)):  # Datum als Excel-Seriennummer
            c.attrib.pop("t", None)
            if not isinstance(v, datetime.datetime):
                v = datetime.datetime(v.year, v.month, v.day)
            ve.text = repr((v - datetime.datetime(1899, 12, 30)).total_seconds() / 86400)
        else:
            c.set("t", "str")
            ve.text = str(v)
        n_cells += 1
    files[path] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def ref_values(ref):
    sheet, rng = ref.rsplit("!", 1)
    ws = vals[sheet.strip("'")]
    c1, r1, c2, r2 = range_boundaries(rng.replace("$", ""))
    return [ws.cell(row=r, column=c).value for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]


n_series = 0
for path in [n for n in files if re.match(r"xl/charts/chart\d+\.xml$", n)]:
    root = etree.fromstring(files[path])
    for tag, cache_tag in (("numRef", "numCache"), ("strRef", "strCache")):
        for ref in root.iter(C + tag):
            data = ref_values(ref.find(C + "f").text)
            old = ref.find(C + cache_tag)
            if old is not None:
                ref.remove(old)
            cache = etree.SubElement(ref, C + cache_tag)
            if cache_tag == "numCache":
                etree.SubElement(cache, C + "formatCode").text = "General"
            etree.SubElement(cache, C + "ptCount").set("val", str(len(data)))
            for i, v in enumerate(data):
                if v is None or (cache_tag == "numCache" and not isinstance(v, (int, float))):
                    continue
                pt = etree.SubElement(cache, C + "pt")
                pt.set("idx", str(i))
                etree.SubElement(pt, C + "v").text = repr(float(v)) if cache_tag == "numCache" else str(v)
            n_series += 1
    files[path] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    for n in zin.namelist():
        z.writestr(n, files[n])
shutil.rmtree(tmpdir, ignore_errors=True)
print(f"Zellen mit Wert: {n_cells}, Diagramm-Referenzen: {n_series}, Formeln: {info['total_formulas']}")
