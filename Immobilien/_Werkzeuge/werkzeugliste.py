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

# (Kategorie, Artikel, Wofür, Kaufen/Leihen, Priorität, Preis von, Preis bis)
WERKZEUGE = [
    ("Akku-System 18 V (eine Marke)", "Akku-Bohrschrauber (Set mit 2 Akkus + Ladegerät)", "Schrauben, Bohren in Holz/Metall, Trockenbau", "Kaufen", "Start", 150, 300),
    ("Akku-System 18 V (eine Marke)", "Zusatzakkus (2 × 4–5 Ah) + Schnellladegerät", "Ganztägig arbeiten ohne Pause", "Kaufen", "Start", 120, 250),
    ("Akku-System 18 V (eine Marke)", "Schlagschrauber", "Lange Schrauben, Dachlatten, Montage", "Kaufen", "Bei Bedarf", 100, 200),
    ("Akku-System 18 V (eine Marke)", "Bohrhammer SDS-plus", "Bohren in Beton/Mauerwerk, leichtes Stemmen, Fliesen abschlagen", "Kaufen", "Start", 150, 350),
    ("Akku-System 18 V (eine Marke)", "Winkelschleifer 125 mm", "Trennen von Metall, Fliesen, Stein (nie Asbest!)", "Kaufen", "Start", 80, 180),
    ("Akku-System 18 V (eine Marke)", "Stichsäge", "Ausschnitte in Holz, Platten, Arbeitsplatten", "Kaufen", "Start", 100, 200),
    ("Akku-System 18 V (eine Marke)", "Multitool (oszillierend)", "Türzargen kürzen, Ausschnitte, Silikon/Fugen entfernen", "Kaufen", "Start", 80, 180),
    ("Akku-System 18 V (eine Marke)", "Handkreissäge", "Gerade Schnitte in Platten, OSB, Holz", "Kaufen", "Bei Bedarf", 120, 250),
    ("Akku-System 18 V (eine Marke)", "LED-Baustrahler (2 Stück)", "Licht in Räumen ohne Strom/Leuchten", "Kaufen", "Start", 40, 100),
    ("Netzgeräte", "Bausauger Staubklasse M", "Staub beim Bohren, Schleifen, Stemmen absaugen", "Kaufen", "Start", 250, 600),
    ("Netzgeräte", "Rührwerk", "Fliesenkleber, Spachtel, Ausgleichsmasse anrühren", "Kaufen", "Start", 80, 200),
    ("Netzgeräte", "Exzenterschleifer", "Holz, Türen, Spachtelstellen schleifen", "Kaufen", "Bei Bedarf", 60, 150),
    ("Netzgeräte", "Kapp- und Gehrungssäge", "Sockelleisten, Vinyldielen, Latten exakt sägen", "Kaufen", "Start", 150, 350),
    ("Netzgeräte", "Heißluftföhn", "Alte Farbe/Kleber lösen, Folien", "Kaufen", "Bei Bedarf", 30, 60),
    ("Messen & Prüfen", "Kreuzlinienlaser", "Gerade Linien für Fliesen, Trockenbau, Bilder", "Kaufen", "Start", 80, 250),
    ("Messen & Prüfen", "Wasserwaagen 60 cm + 120 cm", "Waagerecht/senkrecht prüfen", "Kaufen", "Start", 30, 80),
    ("Messen & Prüfen", "Richtlatte 2 m (Alu)", "Ebenheit von Wänden und Böden prüfen, Abziehen", "Kaufen", "Start", 20, 40),
    ("Messen & Prüfen", "Zollstock + Maßband 8 m", "Aufmaß, Zuschnitt", "Kaufen", "Start", 15, 30),
    ("Messen & Prüfen", "Anschlagwinkel", "Rechte Winkel anzeichnen", "Kaufen", "Start", 10, 25),
    ("Messen & Prüfen", "Leitungsortungsgerät", "Strom-, Wasserleitungen und Metall vor dem Bohren finden", "Kaufen", "Start", 40, 150),
    ("Messen & Prüfen", "Zweipoliger Spannungsprüfer", "Spannungsfreiheit sicher prüfen", "Kaufen", "Start", 20, 50),
    ("Messen & Prüfen", "Feuchtemessgerät", "Feuchte in Wänden/Holz bei Besichtigung und Sanierung", "Kaufen", "Start", 25, 80),
    ("Handwerkzeug", "Hammer + Fäustel", "Nageln, Stemmen mit Meißel", "Kaufen", "Start", 25, 50),
    ("Handwerkzeug", "Meißel-Set (Hand und SDS-plus)", "Stemmen, Fliesen und Putz entfernen", "Kaufen", "Start", 20, 40),
    ("Handwerkzeug", "Kuhfuß / Brechstange", "Abbruch, Leisten und Verkleidungen lösen", "Kaufen", "Start", 15, 30),
    ("Handwerkzeug", "Cuttermesser + Ersatzklingen", "Gipskarton, Vinyl, Folien, Teppich", "Kaufen", "Start", 10, 20),
    ("Handwerkzeug", "VDE-Schraubendreher-Set", "Schalter, Steckdosen (nur spannungsfrei!), Montage", "Kaufen", "Start", 25, 60),
    ("Handwerkzeug", "Zangen-Set (Kombi-, Wasserpumpen-, Seitenschneider)", "Universal, Sanitär-Kleinteile", "Kaufen", "Start", 40, 80),
    ("Handwerkzeug", "Spachtel-Set + Glätter + Kelle", "Spachteln, Putzen, Kleber auftragen", "Kaufen", "Start", 20, 45),
    ("Handwerkzeug", "Kartuschenpresse", "Silikon, Acryl, Montagekleber", "Kaufen", "Start", 10, 25),
    ("Handwerkzeug", "Gummihammer + Zahnspachtel-Set", "Fliesen verlegen", "Kaufen", "Bei Bedarf", 20, 40),
    ("Handwerkzeug", "Bohrer-, Bit- und Sägeblatt-Set", "Verbrauchsmaterial für die Maschinen", "Kaufen", "Start", 30, 60),
    ("Maler & Boden", "Farbrollen-Set + Teleskopstange + Pinsel", "Wände und Decken streichen", "Kaufen", "Start", 40, 80),
    ("Maler & Boden", "Abdeckvlies (wiederverwendbar)", "Böden und Treppen schützen", "Kaufen", "Start", 30, 60),
    ("Maler & Boden", "Verlege-Set Boden (Zugeisen, Schlagklotz, Abstandskeile)", "Klick-Vinyl und Laminat verlegen", "Kaufen", "Start", 20, 40),
    ("Transport & Ordnung", "Schubkarre", "Bauschutt, Material", "Kaufen", "Start", 60, 120),
    ("Transport & Ordnung", "Eimer, Mörtelkübel, Besen, Schaufel", "Anmischen, Aufräumen", "Kaufen", "Start", 30, 60),
    ("Transport & Ordnung", "Stehleiter (Alu, 6–8 Stufen)", "Arbeiten an Decken und oberen Wandbereichen", "Kaufen", "Start", 80, 200),
    ("Transport & Ordnung", "Arbeitsböcke (Paar)", "Zuschnitt, Arbeitstisch", "Kaufen", "Start", 40, 100),
    ("Transport & Ordnung", "Kabeltrommel + Verlängerungen", "Strom auf der Baustelle", "Kaufen", "Start", 40, 80),
    ("Transport & Ordnung", "Werkzeugkisten / Systemkoffer", "Ordnung, Transport, Diebstahlschutz", "Kaufen", "Start", 30, 80),
    ("Außen", "Hochdruckreiniger", "Pflaster, Fassade, Terrasse reinigen", "Kaufen", "Bei Bedarf", 120, 300),
    ("Außen", "Heckenschere, Spaten, Rechen", "Garten herrichten", "Kaufen", "Bei Bedarf", 60, 150),
    ("Leihen (Preis pro Tag)", "Abbruchhammer (schwer)", "Estrich, Fliesen und Mauerwerk abbrechen", "Leihen", "Bei Bedarf", 40, 80),
    ("Leihen (Preis pro Tag)", "Wandschlitzfräse mit Absaugung", "Schlitze für Elektro/Sanitär", "Leihen", "Bei Bedarf", 50, 90),
    ("Leihen (Preis pro Tag)", "Fliesen- und Bodenstripper", "Alte Bodenbeläge und Fliesen großflächig entfernen", "Leihen", "Bei Bedarf", 60, 120),
    ("Leihen (Preis pro Tag)", "Nass-Fliesenschneider", "Große Feinsteinzeugfliesen schneiden", "Leihen", "Bei Bedarf", 40, 80),
    ("Leihen (Preis pro Tag)", "Bautrockner / Luftentfeuchter", "Estrich, Putz und Wände trocknen", "Leihen", "Bei Bedarf", 25, 50),
    ("Leihen (Preis pro Tag)", "Boden-/Estrichschleifer", "Untergründe vorbereiten, Kleberreste (nicht bei Asbest!)", "Leihen", "Bei Bedarf", 60, 120),
    ("Leihen (Preis pro Tag)", "Airless-Farbspritzgerät", "Große Wand- und Deckenflächen schnell streichen", "Leihen", "Bei Bedarf", 50, 100),
    ("Leihen (Preis pro Tag)", "Dampf-Tapetenablöser", "Alte Tapeten lösen", "Leihen", "Bei Bedarf", 15, 30),
    ("Leihen (Preis pro Tag)", "Anhänger", "Material, Sperrmüll", "Leihen", "Bei Bedarf", 20, 50),
    ("Leihen (Preis pro Woche)", "Gerüst (Fassade/Dach)", "Arbeiten in der Höhe (Aufbau durch Fachfirma empfohlen)", "Leihen", "Bei Bedarf", 150, 400),
]

SCHUTZ = [
    ("Atemschutz", "FFP3-Masken (10 Stück)", "Staub beim Schleifen, Stemmen, Abbruch", "Kaufen", "Start", 20, 40),
    ("Augen & Ohren", "Schutzbrille (2 Stück)", "Splitter, Staub", "Kaufen", "Start", 10, 25),
    ("Augen & Ohren", "Gehörschutz (Kapsel)", "Bohrhammer, Flex, Kappsäge", "Kaufen", "Start", 15, 30),
    ("Hände & Füße", "Arbeits- und Schnittschutzhandschuhe (Set)", "Abbruch, Glas, Bleche", "Kaufen", "Start", 15, 30),
    ("Hände & Füße", "Sicherheitsschuhe S3", "Durchtritt- und Zehenschutz", "Kaufen", "Start", 60, 150),
    ("Hände & Füße", "Knieschoner", "Fliesen, Boden, Trockenbau", "Kaufen", "Start", 20, 40),
    ("Kleidung", "Einweg-Overalls (Set)", "Staubige Arbeiten, Mineralwolle", "Kaufen", "Start", 20, 40),
    ("Notfall", "Erste-Hilfe-Kasten (DIN 13157)", "Pflicht auf jeder Baustelle", "Kaufen", "Start", 20, 35),
    ("Notfall", "Feuerlöscher 6 kg (ABC)", "Brandschutz bei Flex-, Heiß- und Elektroarbeiten", "Kaufen", "Start", 40, 80),
]

# (Gewerk, Material, Empfehlung, Faustregel Menge, Preis ca.)
MATERIAL = [
    ("Maler", "Tiefengrund", "lösemittelfrei, für saugende Untergründe", "ca. 1 l pro 5–10 m²", "15–30 € / 10 l"),
    ("Maler", "Spachtelmasse (Fertig- oder Pulverspachtel)", "für Q2/Q3-Oberflächen", "nach Zustand", "15–30 € / Eimer"),
    ("Maler", "Wandfarbe weiß, Dispersion", "Nassabrieb Klasse 2, gute Deckkraft", "ca. 1 l pro 6–7 m² je Anstrich", "60–120 € / 12,5 l"),
    ("Maler", "Malervlies", "nur bei rissigen, schlechten Wänden", "Wandfläche + 10 %", "1–3 € / m²"),
    ("Maler", "Acryl", "Anschlussfugen an Decken, Türzargen", "1 Kartusche pro Raum", "3–6 € / Kartusche"),
    ("Boden", "Rigid-/Klick-Vinyl, helle Eiche", "Nutzungsklasse 33, integrierte Trittschalldämmung", "Bodenfläche + 8–10 % Verschnitt", "25–45 € / m²"),
    ("Boden", "Sockelleisten weiß + Clips/Kleber", "einheitlich im ganzen Haus", "Raumumfang + 10 %", "3–8 € / lfm"),
    ("Boden", "Ausgleichsmasse", "für Unebenheiten bis ca. 10 mm", "ca. 1,5 kg pro m² und mm", "15–30 € / 25 kg"),
    ("Bad", "Feinsteinzeug 60 × 60 cm matt", "hellgrau/beige, rutschhemmend R10 am Boden", "Fläche + 10 % Verschnitt", "25–50 € / m²"),
    ("Bad", "Flexkleber C2 + Fugenmörtel", "für Feinsteinzeug und Fußbodenheizung", "ca. 4–6 kg Kleber pro m²", "15–30 € / 25 kg"),
    ("Bad", "Abdichtungs-Set", "Flüssigabdichtung/Dichtschlämme, Dichtband, Innen-/Außenecken, Manschetten", "alle Spritzwasserflächen + Boden Dusche", "100–250 € je Bad"),
    ("Bad", "Vorwandelemente WC + Waschtisch", "Markenware, mit Spülkasten", "1 je Objekt", "150–350 € / Stück"),
    ("Bad", "Sanitärkeramik weiß", "spülrandloses WC, Waschtisch 60–80 cm", "je Bad", "300–800 € je Bad"),
    ("Bad", "Armaturen + Duschsystem", "solide Mittelklasse-Marke", "je Bad", "300–800 € je Bad"),
    ("Bad", "Duschrinne + Glaswand", "bodengleiche Dusche", "je Bad", "500–1.200 €"),
    ("Bad", "LED-Spiegel + Handtuchheizkörper", "„neues Bad“-Effekt", "je Bad", "250–600 €"),
    ("Bad", "Sanitärsilikon (fungizid)", "Wartungsfugen", "2–3 Kartuschen je Bad", "6–10 € / Kartusche"),
    ("Trockenbau & Dämmung", "Gipskartonplatten (im Bad imprägniert)", "12,5 mm", "Wand-/Deckenfläche + 10 %", "4–8 € / m²"),
    ("Trockenbau & Dämmung", "CW/UW-Profile + Schnellbauschrauben", "Ständerwerk", "nach Wandlänge", "2–4 € / lfm"),
    ("Trockenbau & Dämmung", "Dämmung (Mineralwolle oder Holzfaser, WLS 035)", "oberste Geschossdecke/Dachschräge", "Fläche, Stärke laut Energieberater", "8–25 € / m²"),
    ("Trockenbau & Dämmung", "Dampfbremse + Klebeband", "luftdicht verkleben!", "Fläche + 15 %", "1–3 € / m²"),
    ("Elektro (über Elektriker)", "Schalterprogramm reinweiß", "ein Programm im ganzen Haus", "laut Elektroplan", "5–15 € / Stück"),
    ("Elektro (über Elektriker)", "LED-Einbauspots + Rauchmelder", "Küche, Flur, Bad; Rauchmelder in Schlafräumen und Fluren", "laut Plan", "10–25 € / Spot; 20–40 € / Melder"),
    ("Türen", "Innentür weiß CPL mit Zarge + Drücker", "einheitlich", "je Türöffnung", "250–500 € / Stück"),
    ("Kleinmaterial", "Schrauben- und Dübel-Sortiment", "Universal", "1–2 Sortimente", "30–60 €"),
    ("Kleinmaterial", "Bauschaum, Montagekleber", "Zargen, Fensterbänke", "nach Bedarf", "8–15 € / Dose"),
    ("Kleinmaterial", "Abdeckfolie, Malerkrepp", "Schutz während Arbeiten", "großzügig", "30–60 € pro Projekt"),
    ("Kleinmaterial", "Big Bags, Müllsäcke (reißfest)", "Bauschutt, Abfall getrennt", "nach Rückbau", "5–15 € / Big Bag"),
    ("Kleinmaterial", "Schleifpapier, Trennscheiben", "Verbrauch", "laufend", "20–50 € pro Projekt"),
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
    ws["B2"] = "Preise sind grobe Richtwerte (Stand 2026, Mittelklasse-Marken). Spalte „Hab ich“ per Auswahl abhaken, gelbe Felder kannst du anpassen."
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
w0, w1 = tool_sheet(ws, "Werkzeuge für die Flip-Sanierung", WERKZEUGE, start=9)
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
m["B1"] = "Materialien – dein Flip-Standard"
m["B1"].font = Font(name=F, size=14, bold=True)
m["B2"] = "Einmal festlegen, bei jedem Projekt gleich bestellen. Preise sind grobe Richtwerte (Stand 2026). Bestellstatus per Auswahl setzen."
m["B2"].font = GREY
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

wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print("Werkzeuge", w0, w1, "Schutz", p0, p1)
