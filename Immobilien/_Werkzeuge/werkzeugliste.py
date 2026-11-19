"""Baut die Werkzeug- und Materialliste für Fix-&-Flip-Sanierungen (Excel).

Aufruf: python3 werkzeugliste.py AUSGABE_roh.xlsx
Danach: python3 excel_werte_speichern.py AUSGABE_roh.xlsx ../Fix-und-Flip/04_Sanierungs-Handbuch/Werkzeug_und_Materialliste.xlsx
"""
import sys

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

OUT = sys.argv[1]
F = "Arial"
BLK = Font(name=F, size=10)
BOLD = Font(name=F, size=10, bold=True)
BLUE = Font(name=F, size=10, color="0000FF")
GREY = Font(name=F, size=9, color="555555", italic=True)
HDR = Font(name=F, size=10, bold=True, color="FFFFFF")
HFILL = PatternFill("solid", fgColor="9E3B24")
SUB = PatternFill("solid", fgColor="EDE7E3")
YEL = PatternFill("solid", fgColor="FFFF00")
LINE = Border(bottom=Side(style="thin", color="D5D2CC"))
EUR = '#,##0 €;-#,##0 €;"-"'

# Minimum-Ausstattung: nur was du für Eigenleistung (Rückbau, Maler, Boden, Montage) wirklich brauchst.
# (Kategorie, Artikel, Wofür, Kaufen/Leihen, Priorität, Preis von, Preis bis)
WERKZEUGE = [
    ("Akku-Geräte (alle von einer Marke, 18 V)", "Akku-Bohrschrauber-Set (2 Akkus 4–5 Ah + Ladegerät)", "Schrauben, Bohren in Holz/Metall, Möbel- und Türmontage", "Kaufen", "Start", 180, 350),
    ("Akku-Geräte (alle von einer Marke, 18 V)", "Akku-Bohrhammer SDS-plus + Meißel (Flach- und Spitzmeißel)", "Bohren in Beton/Mauerwerk, Fliesen und Putz abschlagen", "Kaufen", "Start", 170, 380),
    ("Akku-Geräte (alle von einer Marke, 18 V)", "Akku-Multitool (oszillierend) mit Sägeblättern", "Ersetzt Stichsäge und Feinsäge: Zargen kürzen, Ausschnitte, Fugen entfernen", "Kaufen", "Start", 80, 180),
    ("Akku-Geräte (alle von einer Marke, 18 V)", "Akku-Winkelschleifer 125 mm + Trennscheiben", "Metall, Fliesen, Stein trennen (nie Asbest!)", "Kaufen", "Start", 80, 180),
    ("Akku-Geräte (alle von einer Marke, 18 V)", "Rührquirl (passend für Bohrhammer/Bohrschrauber)", "Spachtel und Kleber anrühren – statt eigenem Rührwerk", "Kaufen", "Start", 10, 25),
    ("Staub", "Bausauger Staubklasse M", "Staub beim Bohren, Schleifen, Stemmen absaugen (Gesundheit!) – alternativ mieten", "Kaufen", "Start", 470, 700),
    ("Messen & Prüfen", "Mess-Set: Wasserwaage 120 cm, Zollstock, Maßband 8 m, Bleistift", "Aufmaß, Ausrichten, Zuschnitt", "Kaufen", "Start", 40, 90),
    ("Messen & Prüfen", "Leitungsortungsgerät + zweipoliger Spannungsprüfer", "Vor jedem Bohren Leitungen finden, Spannungsfreiheit prüfen", "Kaufen", "Start", 60, 200),
    ("Handwerkzeug", "Handwerkzeug-Set: Hammer, Kuhfuß, Cuttermesser, VDE-Schraubendreher, Kombizange", "Abbruch, Montage, Zuschneiden", "Kaufen", "Start", 70, 140),
    ("Handwerkzeug", "Spachtel-Set + Glätter + Kartuschenpresse", "Spachteln, Silikon/Acryl, Montagekleber", "Kaufen", "Start", 25, 55),
    ("Maler & Boden", "Maler-Set: Farbrollen, Teleskopstange, Pinsel, Abdeckvlies", "Wände und Decken streichen, Böden schützen", "Kaufen", "Start", 70, 140),
    ("Maler & Boden", "Boden-Set: Zugeisen, Schlagklotz, Abstandskeile + Gehrungslade mit Feinsäge", "Klick-Vinyl und Sockelleisten verlegen", "Kaufen", "Start", 35, 70),
    ("Baustelle", "Stehleiter (Alu, 6–8 Stufen)", "Arbeiten an Decken und oberen Wandbereichen", "Kaufen", "Start", 80, 200),
    ("Baustelle", "Eimer, Mörtelkübel, Besen, Schaufel, Kabeltrommel", "Anmischen, Aufräumen, Strom", "Kaufen", "Start", 70, 140),
    ("Leihen – nur bei Bedarf (Preis pro Tag)", "Abbruchhammer (schwer)", "Estrich, Fliesen, Mauerwerk großflächig abbrechen", "Leihen", "Bei Bedarf", 40, 80),
    ("Leihen – nur bei Bedarf (Preis pro Tag)", "Bautrockner / Luftentfeuchter", "Estrich, Putz und Wände trocknen", "Leihen", "Bei Bedarf", 25, 50),
    ("Leihen – nur bei Bedarf (Preis pro Tag)", "Fliesenschneider", "Nur wenn du selbst fliest – sonst bringt der Fliesenleger ihn mit", "Leihen", "Bei Bedarf", 40, 80),
    ("Leihen – nur bei Bedarf (Preis pro Tag)", "Hochdruckreiniger", "Pflaster, Fassade, Terrasse reinigen", "Leihen", "Bei Bedarf", 30, 60),
    ("Leihen – nur bei Bedarf (Preis pro Tag)", "Anhänger", "Material holen, Sperrmüll wegbringen", "Leihen", "Bei Bedarf", 20, 50),
]

SCHUTZ = [
    ("Pflicht – immer", "FFP3-Masken (10 Stück)", "Staub beim Schleifen, Stemmen, Abbruch", "Kaufen", "Start", 20, 40),
    ("Pflicht – immer", "Schutzbrille", "Splitter, Staub", "Kaufen", "Start", 10, 25),
    ("Pflicht – immer", "Gehörschutz (Kapsel)", "Bohrhammer, Winkelschleifer", "Kaufen", "Start", 15, 30),
    ("Pflicht – immer", "Arbeitshandschuhe (Schnittschutz)", "Abbruch, Glas, Bleche", "Kaufen", "Start", 15, 30),
    ("Pflicht – immer", "Sicherheitsschuhe S3", "Durchtritt- und Zehenschutz", "Kaufen", "Start", 60, 150),
    ("Pflicht – immer", "Erste-Hilfe-Kasten (DIN 13157)", "Gehört auf jede Baustelle", "Kaufen", "Start", 20, 35),
    ("Pflicht – immer", "Feuerlöscher 6 kg (ABC)", "Brandschutz bei Flex- und Elektroarbeiten", "Kaufen", "Start", 40, 80),
]

# Fachbetriebe (Elektro, Sanitär/Heizung, Fliesenleger, Fenster, Estrich) bringen ihr Material laut Angebot mit.
# Du kaufst nur für deine Eigenleistung und wählst die sichtbaren Standard-Materialien aus.
# (Gruppe, Material, Empfehlung, Faustregel Menge, Preis ca.)
MATERIAL = [
    ("Maler (Eigenleistung)", "Tiefengrund", "lösemittelfrei, nur auf saugenden Wänden", "ca. 1 l pro 5–10 m²", "15–30 € / 10 l"),
    ("Maler (Eigenleistung)", "Fertigspachtel", "Löcher und Risse schließen", "1–2 Eimer pro Haus", "15–30 € / Eimer"),
    ("Maler (Eigenleistung)", "Wandfarbe weiß (Dispersion, Nassabrieb Kl. 2)", "eine Farbe für alle Wände und Decken", "ca. 1 l pro 6–7 m² je Anstrich, 2 Anstriche", "60–120 € / 12,5 l"),
    ("Maler (Eigenleistung)", "Acryl + Malerkrepp + Abdeckfolie", "Anschlussfugen, sauberes Abkleben", "1 Kartusche pro Raum", "40–80 € pro Haus"),
    ("Boden (Eigenleistung)", "Klick-Vinyl (Rigid), helle Eiche, Trittschall integriert", "ein Boden für alle Wohnräume", "Bodenfläche + 10 % Verschnitt", "25–45 € / m²"),
    ("Boden (Eigenleistung)", "Sockelleisten weiß + Montagekleber", "einheitlich im ganzen Haus", "Raumumfang + 10 %", "3–8 € / lfm"),
    ("Bad (du wählst aus, Fachbetrieb liefert und baut ein)", "Fliese 60 × 60 cm matt, hellgrau/beige", "eine Fliese für Boden und Wand (Boden R10)", "Fläche + 10 % Verschnitt", "25–50 € / m²"),
    ("Bad (du wählst aus, Fachbetrieb liefert und baut ein)", "Sanitär-Paket: WC, Waschtisch, Armaturen, Duschwand", "weiß, solide Mittelklasse-Marke", "1 Paket pro Bad", "1.000–2.500 € pro Bad"),
    ("Bad (du wählst aus, Fachbetrieb liefert und baut ein)", "Sanitärsilikon (fungizid, weiß)", "Wartungsfugen nachziehen", "2–3 Kartuschen pro Bad", "6–10 € / Kartusche"),
    ("Kleinmaterial", "Schrauben- und Dübel-Sortiment", "Universal", "1 Sortiment", "30–60 €"),
    ("Kleinmaterial", "Big Bags + reißfeste Müllsäcke", "Bauschutt und Abfall getrennt sammeln", "nach Rückbau-Menge", "5–15 € / Big Bag"),
    ("Kleinmaterial", "Verbrauch: Bohrer, Bits, Schleifpapier, Trennscheiben", "Ersatz für die Maschinen", "laufend", "30–60 € pro Projekt"),
    ("Nur falls im Objekt nötig", "Dämmung oberste Geschossdecke + Dampfbremse + Klebeband", "günstigster Sprung in der Energieklasse; Stärke laut Energieberater", "Deckenfläche + 10 %", "10–25 € / m²"),
    ("Nur falls im Objekt nötig", "Gipskarton + Profile + Schrauben", "neue Wände, Vorwände, Deckenverkleidung", "nach Plan", "6–12 € / m²"),
    ("Nur falls im Objekt nötig", "Ausgleichsmasse", "unebene Böden vor dem Vinyl (bis ca. 10 mm)", "ca. 1,5 kg pro m² und mm", "15–30 € / 25 kg"),
]


def header(ws, row, heads, widths):
    for i, h in enumerate(heads):
        c = ws.cell(row=row, column=2 + i, value=h)
        c.font = HDR
        c.fill = HFILL
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for i, w in enumerate(widths):
        ws.column_dimensions[chr(66 + i)].width = w
    ws.column_dimensions["A"].width = 2
    ws.row_dimensions[row].height = 30


def tool_sheet(ws, title, rows, start=7):
    ws.sheet_view.showGridLines = False
    ws["B1"] = title
    ws["B1"].font = Font(name=F, size=14, bold=True)
    ws["B2"] = "Minimum-Ausstattung: nur was du für deine Eigenleistung brauchst. Preise sind grobe Richtwerte (Stand 2026). Spalte „Hab ich“ abhaken, gelbe Felder kannst du anpassen."
    ws["B2"].font = GREY
    header(ws, start - 1, ["Kategorie", "Artikel", "Wofür", "Kaufen / Leihen", "Priorität", "Preis von", "Preis bis", "Hab ich", "Notiz"],
           [24, 40, 46, 13, 12, 11, 11, 10, 26])
    dv = DataValidation(type="list", formula1='"Ja,Nein"', allow_blank=True)
    ws.add_data_validation(dv)
    cat = None
    r = start
    for row in rows:
        if row[0] != cat:
            cat = row[0]
            ws.cell(row=r, column=2, value=cat).font = BOLD
            for col in range(2, 11):
                ws.cell(row=r, column=col).fill = SUB
            r += 1
        for j, v in enumerate(row[1:]):
            c = ws.cell(row=r, column=3 + j, value=v)
            c.font = BLUE if j >= 4 else BLK
            c.alignment = Alignment(wrap_text=True, vertical="top")
            if j >= 4:
                c.number_format = EUR
                c.fill = YEL
        hab = ws.cell(row=r, column=9, value="Nein")
        hab.font = BLUE
        hab.fill = YEL
        hab.alignment = Alignment(horizontal="center", vertical="top")
        dv.add(hab)
        ws.cell(row=r, column=10).fill = YEL
        for col in range(2, 11):
            ws.cell(row=r, column=col).border = LINE
        r += 1
    last = r - 1
    ws.conditional_formatting.add(f"C{start}:J{last}", FormulaRule(formula=[f'$I{start}="Ja"'], font=Font(name=F, color="2E7A4B"), fill=PatternFill("solid", fgColor="DCEFE4")))
    ws.freeze_panes = ws.cell(row=start, column=3)
    return start, last


wb = Workbook()
ws = wb.active
ws.title = "Werkzeuge"
w0, w1 = tool_sheet(ws, "Werkzeuge – Minimum-Ausstattung", WERKZEUGE, start=9)
ps = wb.create_sheet("Schutzausrüstung")
p0, p1 = tool_sheet(ps, "Schutzausrüstung (Pflicht für jede Baustelle)", SCHUTZ)

# Zusammenfassung oben im Blatt Werkzeuge
ws["B3"] = "Übersicht"
ws["B3"].font = BOLD
ws["F3"] = "von"
ws["G3"] = "bis"
for a in ("F3", "G3"):
    ws[a].font = BOLD
    ws[a].alignment = Alignment(horizontal="right")
labels = ["Startausstattung kaufen (Werkzeuge + Schutzausrüstung)", "davon noch nicht vorhanden",
          "Leihgeräte (Summe der Tages-/Wochenpreise, zur Orientierung)"]


def sumifs(sheet, first, last, col, extra):
    pre = f"{sheet}!" if sheet else ""
    rng = lambda c: f"{pre}${c}${first}:${c}${last}"
    cond = "".join(f',{rng(c)},"{v}"' for c, v in extra)
    return f"SUMIFS({rng(col)}{cond})"


for i, (lab, extra, schutz) in enumerate([
    (labels[0], [("E", "Kaufen"), ("F", "Start")], True),
    (labels[1], [("E", "Kaufen"), ("F", "Start"), ("I", "<>Ja")], True),
    (labels[2], [("E", "Leihen")], False),
]):
    r = 4 + i
    ws[f"C{r}"] = lab
    ws[f"C{r}"].font = BLK
    for col, src in (("F", "G"), ("G", "H")):
        f = sumifs("", w0, w1, src, extra)
        if schutz:
            f += "+" + sumifs("Schutzausrüstung", p0, p1, src, extra)
        c = ws[f"{col}{r}"]
        c.value = "=" + f
        c.font = BOLD
        c.number_format = EUR

# Materialien
m = wb.create_sheet("Materialien")
m.sheet_view.showGridLines = False
m["B1"] = "Materialien – das Minimum"
m["B1"].font = Font(name=F, size=14, bold=True)
m["B2"] = ("Fachbetriebe (Elektro, Sanitär/Heizung, Fliesenleger, Fenster, Estrich) bringen ihr Material laut Angebot mit. Du kaufst nur für deine Eigenleistung "
            "und wählst die sichtbaren Materialien aus. Preise sind grobe Richtwerte (Stand 2026). Status per Auswahl setzen.")
m["B2"].font = GREY
m.merge_cells("B2:H2")
m["B2"].alignment = Alignment(wrap_text=True, vertical="top")
m.row_dimensions[2].height = 30
header(m, 4, ["Gewerk", "Material", "Empfehlung", "Faustregel Menge", "Preis ca.", "Status", "Notiz"], [22, 40, 44, 30, 22, 12, 26])
dv2 = DataValidation(type="list", formula1='"Offen,Bestellt,Geliefert"', allow_blank=True)
m.add_data_validation(dv2)
cat = None
r = 5
for row in MATERIAL:
    if row[0] != cat:
        cat = row[0]
        m.cell(row=r, column=2, value=cat).font = BOLD
        for col in range(2, 9):
            m.cell(row=r, column=col).fill = SUB
        r += 1
    for j, v in enumerate(row[1:]):
        c = m.cell(row=r, column=3 + j, value=v)
        c.font = BLK
        c.alignment = Alignment(wrap_text=True, vertical="top")
    st = m.cell(row=r, column=7, value="Offen")
    st.font = BLUE
    st.fill = YEL
    st.alignment = Alignment(horizontal="center", vertical="top")
    dv2.add(st)
    m.cell(row=r, column=8).fill = YEL
    for col in range(2, 9):
        m.cell(row=r, column=col).border = LINE
    r += 1
m.conditional_formatting.add(f"C5:H{r - 1}", FormulaRule(formula=['$G5="Geliefert"'], font=Font(name=F, color="2E7A4B"), fill=PatternFill("solid", fgColor="DCEFE4")))
m.conditional_formatting.add(f"C5:H{r - 1}", FormulaRule(formula=['$G5="Bestellt"'], fill=PatternFill("solid", fgColor="F8EDD8")))
m.freeze_panes = "C5"

for sh in (ws, ps, m):
    sh.page_setup.orientation = "landscape"
    sh.page_setup.paperSize = sh.PAPERSIZE_A4
    sh.page_setup.fitToWidth = 1
    sh.page_setup.fitToHeight = 0
    sh.sheet_properties.pageSetUpPr.fitToPage = True


# ---------------- Set-Empfehlungen ----------------
# (Paket, Artikel, Inhalt/Hinweis, Preis von, Preis bis, Quelle)
PAKETE = [
    ("Paket 1 – Sparpaket Einhell Power X-Change", "Für den Einstieg: alles Nötige zum kleinsten Preis. Gut für 1–2 Projekte pro Jahr.", [
        ("Einhell Akku-Set: Schlagbohrschrauber TE-CD 18/48, Bohrhammer TE-HD 18/20, Winkelschleifer TE-AG 18/115", "inkl. 2 × 3,0-Ah-Akku und Ladegerät", 300, 330, "amazon.de (B0DBLLGKS1)"),
        ("Einhell VARRITO Akku-Multitool (solo)", "passt auf dieselben Akkus, inkl. Zubehör", 59, 64, "geizhals.de / heise Preisvergleich"),
    ]),
    ("Paket 2 – Preis-Leistungs-Tipp Einhell mit Profi-Bohrhammer", "Empfehlung: gleiche Akkus, aber ein bürstenloser Bohrhammer (2,6 J), der Stemmen und Fliesen-Abschlagen besser schafft.", [
        ("Einhell Set TE-TK 18/2 Li: Bohrschrauber + Winkelschleifer", "inkl. Akku und Ladegerät", 120, 125, "obi.de"),
        ("Einhell Professional Bohrhammer TP-HD 18/26 Li BL (solo)", "bürstenlos, SDS-plus, mit Koffer", 150, 160, "geizhals.de"),
        ("Einhell VARRITO Akku-Multitool (solo)", "inkl. Zubehör", 59, 64, "geizhals.de / heise Preisvergleich"),
        ("Zweiter Power-X-Change-Akku 4,0 Ah", "Schätzung – Preis vor Kauf prüfen", 40, 60, "Schätzung"),
    ]),
    ("Paket 3 – Profi Bosch Professional 18 V", "Für Dauereinsatz und viele Projekte; robuster, teurer, größtes Profi-Zubehörprogramm.", [
        ("Bosch Professional Set GSR + GBH + GWS", "Bohrschrauber, Bohrhammer, Winkelschleifer, 2 Akkus, Ladegerät, Tasche – Generation vor Kauf prüfen", 459, 599, "gotools.de, haberl.at, bosch-professional.com (UVP 566 € netto)"),
        ("Bosch Professional Akku-Multitool (solo)", "Schätzung – Modell und Preis vor Kauf prüfen", 150, 220, "Schätzung"),
    ]),
]
EXTRA = [
    ("Bausauger Staubklasse M", [
        ("Metabo ASR 35 M ACP", "Preis-Leistungs-Tipp: 35 l, automatische Filterabreinigung", 523, 589, "geizhals.eu"),
        ("Bosch GAS 35 M AFC Professional", "35 l, automatische Filterreinigung", 470, 604, "werkstatt-ratgeber.de / amazon.de"),
        ("Kärcher NT 30/1 Tact Te M", "30 l, Tact-Filterabreinigung", 640, 666, "geizhals.de"),
    ]),
    ("Messen & Prüfen", [
        ("Bosch Truvo Ortungsgerät", "findet Strom- und Metallleitungen vor dem Bohren", 34, 59, "geizhals.de / bosch-diy.com"),
        ("Zweipoliger Spannungsprüfer (Markengerät)", "Schätzung – Preis vor Kauf prüfen", 20, 50, "Schätzung"),
    ]),
]

e = wb.create_sheet("Set-Empfehlungen", 0)
e.sheet_view.showGridLines = False
e["B1"] = "Werkzeug-Sets mit gutem Preis-Leistungs-Verhältnis"
e["B1"].font = Font(name=F, size=14, bold=True)
e["B2"] = ("Ein Akku-System wählen und dabei bleiben – die Akkus sind der teure Teil. Preise sind Momentaufnahmen aus Preisvergleichen "
           "und Händlerseiten (Stand Nov. 2026) und schwanken stark; „Schätzung“ = kein belastbarer Preis gefunden. Vor dem Kauf Artikelnummer und Inhalt prüfen.")
e["B2"].font = GREY
e.merge_cells("B2:G2")
e["B2"].alignment = Alignment(wrap_text=True, vertical="top")
e.row_dimensions[2].height = 42
for i, h in enumerate(["Paket / Artikel", "Inhalt / Hinweis", "Preis von", "Preis bis", "Quelle", "Gekauft"]):
    c = e.cell(row=4, column=2 + i, value=h)
    c.font = HDR
    c.fill = HFILL
    c.alignment = Alignment(wrap_text=True, vertical="center")
for col, w in {"A": 2, "B": 58, "C": 50, "D": 12, "E": 12, "F": 34, "G": 10}.items():
    e.column_dimensions[col].width = w
dv3 = DataValidation(type="list", formula1='"Ja,Nein"', allow_blank=True)
e.add_data_validation(dv3)
TOT = PatternFill("solid", fgColor="F5E6E0")
r = 5
paket_rows = []


def item_rows(items, r):
    first = r
    for art, inh, lo_, hi_, src in items:
        e.cell(row=r, column=2, value=art).font = BLK
        e.cell(row=r, column=3, value=inh).font = GREY if inh.startswith("Schätzung") else BLK
        for col, v in ((4, lo_), (5, hi_)):
            c = e.cell(row=r, column=col, value=v)
            c.font = BLUE
            c.fill = YEL
            c.number_format = EUR
        e.cell(row=r, column=6, value=src).font = GREY
        g = e.cell(row=r, column=7, value="Nein")
        g.font = BLUE
        g.fill = YEL
        g.alignment = Alignment(horizontal="center")
        dv3.add(g)
        for col in range(2, 8):
            e.cell(row=r, column=col).alignment = Alignment(wrap_text=True, vertical="top", horizontal="center" if col == 7 else None)
            e.cell(row=r, column=col).border = LINE
        r += 1
    return first, r - 1, r


for name, why, items in PAKETE:
    e.cell(row=r, column=2, value=name).font = BOLD
    for col in range(2, 8):
        e.cell(row=r, column=col).fill = SUB
    head = r
    r += 1
    e.cell(row=r, column=2, value=why).font = GREY
    e.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    e.cell(row=r, column=2).alignment = Alignment(wrap_text=True)
    e.row_dimensions[r].height = 28
    r += 1
    f0, f1, r = item_rows(items, r)
    e.cell(row=r, column=2, value="Summe Paket").font = BOLD
    for col, L in ((4, "D"), (5, "E")):
        c = e.cell(row=r, column=col, value=f"=SUM({L}{f0}:{L}{f1})")
        c.font = BOLD
        c.number_format = EUR
    for col in range(2, 8):
        e.cell(row=r, column=col).fill = TOT
    paket_rows.append((name, r, items))
    r += 2

for name, items in EXTRA:
    e.cell(row=r, column=2, value=name + " (für alle Pakete gleich – eines davon wählen)").font = BOLD
    for col in range(2, 8):
        e.cell(row=r, column=col).fill = SUB
    r += 1
    _, _, r = item_rows(items, r)
    r += 1

e.cell(row=r, column=2, value="Meine Empfehlung").font = BOLD
tips = [
    "• Paket 2 bietet das beste Preis-Leistungs-Verhältnis: günstige Einhell-Akkus, aber ein kräftiger bürstenloser Bohrhammer für Abbruch und Fliesen.",
    "• Paket 3 lohnt sich, wenn du mehrere Projekte pro Jahr machst oder Handwerker-Qualität willst – höherer Wiederverkaufswert.",
    "• Parkside (Lidl) ist als Hauptsystem nicht empfohlen: Das aktuelle Kombiset hat keinen Bohrhammer und Winkelschleifer, und ein älterer Parkside-Bohrhammer fiel im Dauertest durch.",
    "• Den Bausauger kannst du für das erste Projekt auch mieten – er ist das teuerste Einzelteil.",
]
for i, tip in enumerate(tips):
    c = e.cell(row=r + 1 + i, column=2, value=tip)
    c.font = GREY
    e.merge_cells(start_row=r + 1 + i, start_column=2, end_row=r + 1 + i, end_column=6)
e.freeze_panes = "C5"
e.page_setup.orientation = "landscape"
e.page_setup.paperSize = e.PAPERSIZE_A4
e.page_setup.fitToWidth = 1
e.page_setup.fitToHeight = 0
e.sheet_properties.pageSetUpPr.fitToPage = True
print("Pakete", [(n[:8], rr) for n, rr, _ in paket_rows])

wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print("Werkzeuge", w0, w1, "Schutz", p0, p1)
