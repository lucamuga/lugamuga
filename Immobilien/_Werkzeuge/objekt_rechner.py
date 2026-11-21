"""Objekt-Rechner für eine vermietete Immobilie mit bis zu zwei Wohnungen (Excel).

Jede Annahme hat drei Spalten:
  C = konservativer Wert (damit wird gerechnet – immer mit Puffer)
  D = Spanne als Text (zeigt Verhandlungs- bzw. Unsicherheitsspielraum)
  E = günstiger Wert (Best Case, zum Vergleich)
Ergebnisse werden für C (konservativ) und E (günstig) nebeneinander berechnet.

Aufruf: python3 objekt_rechner.py objekt.json AUSGABE_roh.xlsx
Danach: python3 excel_werte_speichern.py AUSGABE_roh.xlsx ZIEL.xlsx
"""
import json
import sys

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

CFG = json.load(open(sys.argv[1], encoding="utf-8"))
OUT = sys.argv[2]

F = "Arial"
BLK = Font(name=F, size=10)
BOLD = Font(name=F, size=10, bold=True)
BLUE = Font(name=F, size=10, color="0000FF")
GREY = Font(name=F, size=9, color="555555", italic=True)
HDR = Font(name=F, size=10, bold=True, color="FFFFFF")
HFILL = PatternFill("solid", fgColor="9E3B24")
SUB = PatternFill("solid", fgColor="EDE7E3")
YEL = PatternFill("solid", fgColor="FFFF00")
KEY = PatternFill("solid", fgColor="F5E6E0")
LINE = Border(bottom=Side(style="thin", color="D5D2CC"))
FMT = {"eur": '#,##0 €;-#,##0 €;"-"', "eur2": '#,##0.00 €;-#,##0.00 €;"-"', "pct": '0.0%;-0.0%;"-"',
       "pct2": '0.00%;-0.00%;"-"', "num": '0', "num1": '0.0'}

wb = Workbook()
ws = wb.active
ws.title = "Rechner"
ws.sheet_view.showGridLines = False
ws["B1"] = CFG["titel"]
ws["B1"].font = Font(name=F, size=15, bold=True, color="1C2430")
ws["B2"] = CFG["untertitel"]
ws["B2"].font = GREY
ws.merge_cells("B2:G2")
ws["B2"].alignment = Alignment(wrap_text=True, vertical="top")
ws.row_dimensions[2].height = 42

heads = ["Annahme", "Konservativ (gerechnet)", "Spanne", "Günstig (Vergleich)", "Einheit", "Quelle / Begründung"]
for i, h in enumerate(heads):
    c = ws.cell(row=4, column=2 + i, value=h)
    c.font = HDR
    c.fill = HFILL
    c.alignment = Alignment(wrap_text=True, vertical="center")
ws.row_dimensions[4].height = 30

R = {}  # Name -> Zeile
r = 5
for grp in CFG["eingaben"]:
    ws.cell(row=r, column=2, value=grp["gruppe"]).font = BOLD
    for col in range(2, 8):
        ws.cell(row=r, column=col).fill = SUB
    r += 1
    for e in grp["werte"]:
        R[e["name"]] = r
        ws.cell(row=r, column=2, value=e["label"]).font = BLK
        for col, key in ((3, "konservativ"), (5, "guenstig")):
            c = ws.cell(row=r, column=col, value=e[key])
            c.font = BLUE
            c.fill = YEL
            c.number_format = FMT[e["fmt"]]
            c.alignment = Alignment(horizontal="right", vertical="top")
        ws.cell(row=r, column=4, value=e.get("spanne", "–")).font = BLK
        ws.cell(row=r, column=6, value=e["einheit"]).font = GREY
        for col in (2, 6):
            ws.cell(row=r, column=col).alignment = Alignment(vertical="top", wrap_text=(col == 2))
        q = ws.cell(row=r, column=7, value=e["quelle"])
        q.font = GREY
        for col in (4, 7):
            ws.cell(row=r, column=col).alignment = Alignment(wrap_text=True, vertical="top")
        for col in range(2, 8):
            ws.cell(row=r, column=col).border = LINE
        if e.get("aus_expose"):
            ws.cell(row=r, column=7).font = Font(name=F, size=9, color="2E7A4B", bold=True)
        r += 1
last_input = r - 1


def v(name, col):
    return f"{col}{R[name]}"


# ---------------- Ergebnisse ----------------
r = last_input + 2
ws.cell(row=r, column=2, value="Ergebnis")
for i, h in enumerate(["Ergebnis", "Konservativ", "", "Günstig", "", "Erklärung"]):
    c = ws.cell(row=r, column=2 + i, value=h)
    c.font = HDR
    c.fill = HFILL
r += 1


def res(col, key):
    """Formel-Text für Ergebnis 'key' in Spalte col (C oder E)."""
    g = lambda n: v(n, col)
    O = lambda n: f"{col}{OUT_ROWS[n]}"
    return {
        "flaeche": f"={g('wf1')}+{g('wf2')}",
        "kalt": f"={g('wf1')}*{g('miete1')}+{g('wf2')}*{g('miete2')}",
        "nk": f"={g('preis')}*({g('grest')}+{g('notar')}+{g('makler')})",
        "gesamt": f"={g('preis')}+{O('nk')}+{g('sofort')}",
        "darlehen": f"=MAX(0,{O('gesamt')}-{g('ek')})",
        "rate": f"={O('darlehen')}*({g('zins')}+{g('tilgung')})/12",
        "tilgJ": f"=IF({O('darlehen')}=0,0,{O('darlehen')}-({O('darlehen')}*(1+{g('zins')}/12)^12-{O('rate')}*((1+{g('zins')}/12)^12-1)/({g('zins')}/12)))",
        "zinsJ": f"=12*{O('rate')}-{O('tilgJ')}",
        "m_kalt": f"={O('kalt')}",
        "m_ausfall": f"=-{O('kalt')}*{g('ausfall')}",
        "m_pv": f"={g('pv')}/12",
        "m_rate": f"=-{O('rate')}",
        "m_inst": f"=-{O('flaeche')}*{g('instand')}/12",
        "m_sonst": f"=-{g('sonst')}",
        "cfVor": f"=SUM({O('m_kalt')}:{O('m_sonst')})",
        "afa": f"=(({g('preis')}+{O('nk')})*{g('gebaeude')}+{g('sofort')})*{g('afa')}",
        "ergebnis": f"=({O('m_kalt')}+{O('m_ausfall')}+{O('m_pv')}+{O('m_inst')}+{O('m_sonst')})*12-{O('zinsJ')}-{O('afa')}",
        "steuer": f"={O('ergebnis')}*{g('steuersatz')}",
        "m_steuer": f"=-{O('steuer')}/12",
        "cfNach": f"={O('cfVor')}+{O('m_steuer')}",
        "faktor": f"=IFERROR({g('preis')}/({O('kalt')}*12),0)",
        "brutto": f"=IFERROR({O('kalt')}*12/{g('preis')},0)",
        "netto": f"=IFERROR(({O('m_kalt')}+{O('m_ausfall')}+{O('m_pv')}+{O('m_inst')}+{O('m_sonst')})*12/{O('gesamt')},0)",
        "anlauf": f"={g('leer_monate')}*({O('rate')}+{g('sonst')})",
        "ek_gesamt": f"=MIN({g('ek')},{O('gesamt')})+{O('anlauf')}",
        "ekr": f"=IFERROR(({O('cfNach')}*12+{O('tilgJ')})/{O('ek_gesamt')},0)",
        "steigung": f"=IF({O('darlehen')}=0,0,(1+{g('grest')}+{g('notar')}+{g('makler')})/12*(({g('zins')}+{g('tilgung')})-{g('steuersatz')}*({O('zinsJ')}/{O('darlehen')}+{g('gebaeude')}*{g('afa')})))",
        "maxPreis": f"=IF({O('steigung')}<=0,\"–\",MAX(0,{g('preis')}+{O('cfNach')}/{O('steigung')}))",
        "preisF20": f"=20*{O('kalt')}*12",
    }[key]


LAYOUT = [
    ("sec", "Investition"),
    ("flaeche", "Wohnfläche gesamt (m²)", "num", "Wohnung 1 + Wohnung 2"),
    ("nk", "Kaufnebenkosten", "eur", "Grunderwerbsteuer + Notar/Grundbuch + Makler"),
    ("gesamt", "Gesamtinvestition", "eur", "Kaufpreis + Nebenkosten + Sofort-Investitionen"),
    ("darlehen", "Darlehen", "eur", "Gesamtinvestition − Eigenkapital"),
    ("rate", "Kreditrate pro Monat", "eur", "Darlehen × (Zins + Tilgung) ÷ 12"),
    ("sec", "Monatliche Rechnung (1. Jahr, voll vermietet)"),
    ("m_kalt", "Kaltmiete beide Wohnungen", "eur", ""),
    ("m_ausfall", "Mietausfall-Reserve", "eur", ""),
    ("m_pv", "Einnahmen Photovoltaik", "eur", "Volleinspeisung laut Exposé, Höhe unbekannt"),
    ("m_rate", "Kreditrate", "eur", ""),
    ("m_inst", "Instandhaltungsrücklage", "eur", ""),
    ("m_sonst", "Nicht umlagefähige Kosten", "eur", ""),
    ("cfVor", "Cashflow vor Steuern", "eur", ""),
    ("m_steuer", "Steuereffekt (+ = Erstattung)", "eur", "vereinfacht, siehe Zeilen darunter"),
    ("cfNach", "CASHFLOW NACH STEUERN / MONAT", "eur", "Das, was jeden Monat übrig bleibt oder zugeschossen werden muss"),
    ("sec", "Kennzahlen"),
    ("kalt", "Kaltmiete pro Monat", "eur", ""),
    ("faktor", "Kaufpreisfaktor", "num1", "bis 18 günstig, 18–22 mittel, über 22 teuer"),
    ("brutto", "Bruttomietrendite", "pct2", "Jahreskaltmiete ÷ Kaufpreis"),
    ("netto", "Nettomietrendite", "pct2", "nach Ausfall, Instandhaltung, Kosten ÷ Gesamtinvestition"),
    ("ekr", "Eigenkapitalrendite (Cashflow + Tilgung)", "pct", "ohne Wertsteigerung"),
    ("sec", "Eigenkapital & Anlaufphase"),
    ("anlauf", "Kosten der Anlaufphase (Rate + Kosten ohne Miete)", "eur", "Monate bis zur Vermietung (Umbau, Mietersuche)"),
    ("ek_gesamt", "Eigenkapitalbedarf gesamt", "eur", "eingesetztes Eigenkapital + Anlaufphase"),
    ("sec", "Steuer 1. Jahr (vereinfacht)"),
    ("zinsJ", "Zinsen Jahr 1 (monatsgenau)", "eur", ""),
    ("tilgJ", "Tilgung Jahr 1 (Vermögensaufbau)", "eur", ""),
    ("afa", "Abschreibung (AfA) pro Jahr", "eur", "Gebäudeanteil von Kaufpreis + Nebenkosten + Sofort-Investitionen × AfA-Satz"),
    ("ergebnis", "Steuerliches Ergebnis", "eur", "negativ = Verlust, mindert deine Steuer"),
    ("steuer", "Steuer (+ zahlen / − Erstattung)", "eur", ""),
    ("sec", "Einkaufs-Check"),
    ("maxPreis", "MAX. KAUFPREIS FÜR CASHFLOW ≥ 0", "eur", "bei gleichem Eigenkapital und gleichen Annahmen"),
    ("preisF20", "Kaufpreis für Faktor 20", "eur", "Orientierung: Faktor 20 = 5 % Bruttorendite"),
    ("steigung", "Hilfswert: € Cashflow pro € Kaufpreis", "num1", "für die Verhandlungstabelle"),
]
OUT_ROWS = {}
rr = r
for item in LAYOUT:
    if item[0] != "sec":
        OUT_ROWS[item[0]] = rr
    rr += 1
for item in LAYOUT:
    if item[0] == "sec":
        ws.cell(row=r, column=2, value=item[1]).font = BOLD
        for col in range(2, 8):
            ws.cell(row=r, column=col).fill = SUB
        r += 1
        continue
    key, label, fmt, expl = item
    big = label.isupper()
    ws.cell(row=r, column=2, value=label).font = BOLD if big else BLK
    for col in ("C", "E"):
        c = ws[f"{col}{r}"]
        c.value = res(col, key)
        c.number_format = FMT[fmt] if key != "steigung" else "0.000000"
        c.font = BOLD if big else BLK
        c.alignment = Alignment(horizontal="right")
    ws.cell(row=r, column=7, value=expl).font = GREY
    ws.cell(row=r, column=7).alignment = Alignment(wrap_text=True, vertical="top")
    if big:
        for col in range(2, 8):
            ws.cell(row=r, column=col).fill = KEY
    if key == "steigung":
        for col in range(2, 8):
            ws.cell(row=r, column=col).font = GREY
    r += 1

green = Font(name=F, bold=True, color="2E7A4B")
red = Font(name=F, bold=True, color="B3362F")
for key in ("cfNach", "cfVor"):
    rw = OUT_ROWS[key]
    ws.conditional_formatting.add(f"C{rw}:E{rw}", FormulaRule(formula=[f"C{rw}>=0"], font=green))
    ws.conditional_formatting.add(f"C{rw}:E{rw}", FormulaRule(formula=[f"C{rw}<0"], font=red))

# ---------------- Verhandlungstabelle ----------------
r += 1
for i, h in enumerate(["Kaufpreis-Szenarien (konservative Annahmen)", "Kaufpreis", "Cashflow / Monat", "Kaufpreisfaktor", "", "Hinweis"]):
    c = ws.cell(row=r, column=2 + i, value=h)
    c.font = HDR
    c.fill = HFILL
r += 1
tab_first = r
for lab, preis, hint in CFG["verhandlung"]:
    ws.cell(row=r, column=2, value=lab).font = BLK
    p = ws.cell(row=r, column=3, value=preis)
    p.font = BLUE
    p.fill = YEL
    p.number_format = FMT["eur"]
    pr = R["preis"]
    cf = ws.cell(row=r, column=4, value=f"=$C${OUT_ROWS['cfNach']}-$C${OUT_ROWS['steigung']}*(C{r}-$C${pr})")
    cf.number_format = FMT["eur"]
    cf.font = BOLD
    fk = ws.cell(row=r, column=5, value=f"=IFERROR(C{r}/($C${OUT_ROWS['kalt']}*12),0)")
    fk.number_format = FMT["num1"]
    ws.cell(row=r, column=7, value=hint).font = GREY
    for col in range(2, 8):
        ws.cell(row=r, column=col).border = LINE
    r += 1
ws.conditional_formatting.add(f"D{tab_first}:D{r - 1}", FormulaRule(formula=[f"D{tab_first}>=0"], font=green))
ws.conditional_formatting.add(f"D{tab_first}:D{r - 1}", FormulaRule(formula=[f"D{tab_first}<0"], font=red))

r += 1
for i, n in enumerate(CFG["hinweise"]):
    c = ws.cell(row=r + i, column=2, value=n)
    c.font = BOLD if i == 0 else GREY
    ws.merge_cells(start_row=r + i, start_column=2, end_row=r + i, end_column=7)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r + i].height = 15 if i == 0 else 28

ws[f"B{OUT_ROWS['maxPreis']}"].comment = Comment(
    "Der Cashflow hängt linear vom Kaufpreis ab. Höchster Kaufpreis, bei dem der Cashflow nach Steuern noch 0 € ist (gleiches Eigenkapital, Darlehen > 0).", "Rechner")
for col, w in {"A": 2, "B": 44, "C": 17, "D": 26, "E": 17, "F": 14, "G": 58}.items():
    ws.column_dimensions[col].width = w
ws.freeze_panes = "C5"
ws.page_setup.orientation = "landscape"
ws.page_setup.paperSize = ws.PAPERSIZE_A4
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
json.dump({"inputs": R, "outputs": OUT_ROWS, "tabelle": [tab_first, r - 1 - len(CFG["hinweise"]) - 1]},
          open(OUT + ".map.json", "w"))
print("ok", OUT)
