from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.comments import Comment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

import os, sys
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(__file__), "flip_raw.xlsx")
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
EUR = '#,##0 €;-#,##0 €;"-"'
PCT = '0.0%;-0.0%;"-"'
PCT2 = '0.00%;-0.00%;"-"'

# ---------------- Deals ----------------
# name, ort, quelle, hinweis, angebot, kauf, wf, baujahr, grest, notar, makler, reno_qm, puffer, monate, ltv, zins, lfd, arv_qm, vk, vkfix, tax, ziel, miete
D = [
    ("Bungalow Damme-West", "Damme-West, 748 m² Grundstück", "immowelt / immobilo (Nov. 2026)",
     "Baujahr 1968/1979 (Angaben widersprüchlich), „Entwicklungspotenzial“",
     250000, 250000, 139.5, "1968/79", 0.05, 0.02, 0.0357, 600, 0.15, 8, 0.70, 0.045, 300, 2700, 0.0357, 1500, 0.35, 0.15, 0),
    ("Doppelhaushälfte Damme", "Damme-Glückauf, 581 m² Grundstück", "Immobilien Völker, ImmoNr 261698",
     "renovierungsbedürftig, Baujahr 1984, Käuferprovision 3,57 %",
     235000, 235000, 113, "1984", 0.05, 0.02, 0.0357, 450, 0.15, 6, 0.70, 0.045, 300, 2600, 0.0357, 1500, 0.35, 0.15, 0),
    ("Einfamilienhaus Lohne", "Lohne, 524 m² Grundstück, 10 Zimmer", "Haskamp Immobilien",
     "Baujahr 1946, Energieklasse H, Gasheizung von 2024, + 121 m² Nutzfläche",
     195000, 195000, 224, "1946", 0.05, 0.02, 0.0357, 750, 0.15, 10, 0.70, 0.045, 350, 2200, 0.0357, 1500, 0.35, 0.15, 0),
    ("Haus Lemförde", "Lemförde, Bahnhofstr. 76, 3.136 m² Grundstück", "immowelt (Lkr. Diepholz)",
     "Wasserschaden + Schimmel laut Inserat; Chance nur über Grundstücksteilung",
     236000, 236000, 149, "?", 0.05, 0.02, 0.0357, 900, 0.20, 12, 0.70, 0.045, 300, 2300, 0.0357, 1500, 0.35, 0.15, 0),
    ("Wohnung Damme-Ort", "Damme-Ort, 3 Zimmer", "Aggregator (immosuchmaschine)",
     "Baujahr ca. 1993, gepflegt – Gegenbeispiel: kein Sanierungsfall",
     190000, 190000, 83, "1993", 0.05, 0.02, 0.0357, 300, 0.15, 4, 0.70, 0.045, 250, 2700, 0.0357, 1500, 0.35, 0.15, 0),
    ("Live-in-Flip: DHH Damme", "wie Spalte D, du wohnst 2 Jahre selbst darin", "Szenario",
     "verhandelt auf 215.000 €, viel Eigenleistung, steuerfrei bei Eigennutzung, Verkauf privat",
     235000, 215000, 113, "1984", 0.05, 0.02, 0.0357, 350, 0.15, 24, 0.85, 0.036, 150, 2600, 0.0, 2500, 0.0, 0.15, 850),
]

wb = Workbook()
ws = wb.active
ws.title = "Deals"
ws.sheet_view.showGridLines = False
ws["B1"] = "Fix & Flip Rechner – Deals in Damme und Umgebung"
ws["B1"].font = Font(name=F, size=15, bold=True, color="1C2430")
ws["B2"] = ("Gelbe Felder mit blauer Schrift sind Annahmen und können geändert werden. Alle Werte sind Schätzungen auf Basis von Online-Inseraten "
            "(Stand Nov. 2026) – vor einem Gebot Exposé, Besichtigung mit Bausachverständigem und Handwerkerangebote einholen.")
ws["B2"].font = GREY
ws.merge_cells("B2:H2")
ws["B2"].alignment = Alignment(wrap_text=True)
ws.row_dimensions[2].height = 28

cols = [get_column_letter(3 + i) for i in range(len(D))]

def header_row(r, text):
    ws[f"B{r}"] = text
    ws[f"B{r}"].font = HDR
    for c in ["B"] + cols:
        ws[f"{c}{r}"].fill = HFILL

header_row(3, "Objekt")
for i, d in enumerate(D):
    c = cols[i]
    ws[f"{c}3"] = d[0]
    ws[f"{c}3"].font = HDR
    ws[f"{c}3"].alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
ws.row_dimensions[3].height = 30

text_rows = [(4, "Lage", 1), (5, "Quelle", 2), (6, "Hinweis", 3)]
for r, lab, idx in text_rows:
    ws[f"B{r}"] = lab
    ws[f"B{r}"].font = BOLD
    for i, d in enumerate(D):
        cc = ws[f"{cols[i]}{r}"]
        cc.value = d[idx]
        cc.font = GREY if r != 4 else BLK
        cc.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 40 if r == 6 else 28

inputs = [
    (8, None, "Kauf"),
    (9, "Angebotspreis", 4, EUR),
    (10, "Kaufpreis (verhandelt)", 5, EUR),
    (11, "Wohnfläche (m²)", 6, '0.0'),
    (12, "Baujahr", 7, '@'),
    (13, "Grunderwerbsteuer", 8, PCT),
    (14, "Notar & Grundbuch", 9, PCT),
    (15, "Maklerprovision Kauf", 10, PCT2),
    (16, None, "Sanierung & Haltedauer"),
    (17, "Sanierungskosten pro m² Wohnfläche", 11, EUR),
    (18, "Puffer für Unvorhergesehenes", 12, PCT),
    (19, "Haltedauer bis Verkauf (Monate)", 13, '0'),
    (20, "Finanzierungsanteil (Bank)", 14, PCT),
    (21, "Zins p. a.", 15, PCT2),
    (22, "Laufende Kosten pro Monat", 16, EUR),
    (23, None, "Verkauf & Steuer"),
    (24, "Verkaufspreis pro m² nach Sanierung", 17, EUR),
    (25, "Verkaufskosten (Makler) in %", 18, PCT2),
    (26, "Verkaufskosten pauschal (Energieausweis, Fotos, Anzeigen)", 19, EUR),
    (27, "Steuersatz auf den Gewinn", 20, PCT),
    (28, "Zielmarge (Gewinn vor Steuern / Verkaufspreis)", 21, PCT),
    (29, "Gesparte Kaltmiete pro Monat (nur Live-in)", 22, EUR),
]
for r, lab, idx, *fmt in inputs:
    if lab is None:
        ws[f"B{r}"] = idx
        ws[f"B{r}"].font = BOLD
        for c in ["B"] + cols:
            ws[f"{c}{r}"].fill = SUB
        continue
    ws[f"B{r}"] = lab
    ws[f"B{r}"].font = BLK
    for i, d in enumerate(D):
        cc = ws[f"{cols[i]}{r}"]
        cc.value = d[idx]
        cc.font = BLUE
        cc.fill = YEL
        cc.number_format = fmt[0]
        cc.alignment = Alignment(horizontal="right")

out = [
    (31, None, "Ergebnis"),
    (32, "Verkaufspreis nach Sanierung", "={c}11*{c}24", EUR),
    (33, "Kaufnebenkosten", "={c}10*({c}13+{c}14+{c}15)", EUR),
    (34, "Sanierung inkl. Puffer", "={c}11*{c}17*(1+{c}18)", EUR),
    (35, "Investition (Kauf + Nebenkosten + Sanierung)", "={c}10+{c}33+{c}34", EUR),
    (36, "Zinsen während der Haltedauer", "={c}35*{c}20*{c}21*{c}19/12", EUR),
    (37, "Laufende Kosten während der Haltedauer", "={c}22*{c}19", EUR),
    (38, "Verkaufskosten", "={c}32*{c}25+{c}26", EUR),
    (39, "Gesamtkosten", "={c}35+{c}36+{c}37+{c}38", EUR),
    (40, "GEWINN VOR STEUERN", "={c}32-{c}39", EUR),
    (41, "Steuer", "=MAX(0,{c}40)*{c}27", EUR),
    (42, "Gewinn nach Steuern", "={c}40-{c}41", EUR),
    (43, "Gesparte Miete (Live-in)", "={c}29*{c}19", EUR),
    (44, "ERGEBNIS FÜR DICH", "={c}42+{c}43", EUR),
    (45, "Marge vor Steuern inkl. gesparter Miete / Verkaufspreis", "=IFERROR(({c}40+{c}43)/{c}32,0)", PCT),
    (46, "Kostenquote (Gesamtkosten / Verkaufspreis)", "=IFERROR({c}39/{c}32,0)", PCT),
    (47, "Eigenkapitalbedarf", "={c}35*(1-{c}20)+{c}36+{c}37", EUR),
    (48, "Rendite auf dein Eigenkapital p. a.", "=IFERROR({c}44/{c}47*12/{c}19,0)", PCT),
    (49, None, "Einkaufs-Check"),
    (50, "Max. Kaufpreis für die Zielmarge", "=MAX(0,({c}32*(1-{c}25)-{c}26-{c}34*(1+{c}20*{c}21*{c}19/12)-{c}37+{c}43-{c}28*{c}32)/((1+{c}13+{c}14+{c}15)*(1+{c}20*{c}21*{c}19/12)))", EUR),
    (51, "Nötiger Abschlag auf den Angebotspreis", "=IFERROR(MAX(0,1-{c}50/{c}9),0)", PCT),
    (52, "Bewertung", '=IF({c}45>={c}28,"Deal",IF({c}45>=0.08,"Grenzwertig – nachverhandeln",IF({c}45>=0,"Zu knapp","Verlust zu diesem Preis")))', None),
]
for r, lab, f, *fmt in out:
    if lab is None:
        ws[f"B{r}"] = f
        ws[f"B{r}"].font = HDR
        for c in ["B"] + cols:
            ws[f"{c}{r}"].fill = HFILL
        continue
    ws[f"B{r}"] = lab
    ws[f"B{r}"].font = BOLD if lab.isupper() or r in (39, 50, 52) else BLK
    for c in cols:
        cc = ws[f"{c}{r}"]
        cc.value = f.format(c=c)
        cc.font = BOLD if lab.isupper() or r in (39, 50, 52) else BLK
        if fmt[0]:
            cc.number_format = fmt[0]
        cc.alignment = Alignment(horizontal="right" if r != 52 else "center", wrap_text=(r == 52))
for r in (40, 44, 50):
    for c in ["B"] + cols:
        ws[f"{c}{r}"].fill = KEY
ws.row_dimensions[52].height = 30

rng = f"{cols[0]}52:{cols[-1]}52"
first = cols[0]
ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{first}52="Deal"'], fill=PatternFill("solid", fgColor="DCEFE4"), font=Font(name=F, bold=True, color="2E7A4B")))
ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({first}52,5)="Grenz"'], fill=PatternFill("solid", fgColor="F8EDD8"), font=Font(name=F, bold=True, color="9A620F")))
ws.conditional_formatting.add(rng, FormulaRule(formula=[f'OR({first}52="Zu knapp",LEFT({first}52,7)="Verlust")'], fill=PatternFill("solid", fgColor="F6DCD9"), font=Font(name=F, bold=True, color="B3362F")))
for r in (40, 42, 44):
    ws.conditional_formatting.add(f"{cols[0]}{r}:{cols[-1]}{r}", FormulaRule(formula=[f"{first}{r}<0"], font=Font(name=F, bold=True, color="B3362F")))
    ws.conditional_formatting.add(f"{cols[0]}{r}:{cols[-1]}{r}", FormulaRule(formula=[f"{first}{r}>=0"], font=Font(name=F, bold=True, color="2E7A4B")))

ws["B17"].comment = Comment("Richtwerte 2026: leichte Renovierung 200–400 €/m², mittlere Sanierung 400–700 €/m², Komplett-/Kernsanierung 800–1.400 €/m². Siehe Blatt Sanierungskosten.", "Rechner")
ws["B24"].comment = Comment("Damme Ø Häuser ca. 2.420 €/m² (immowelt Okt. 2026), unsanierte DHH ca. 2.090 €/m², Neubau-Niveau A+ ca. 3.480 €/m². Für ein gut saniertes älteres Haus sind 2.600–2.800 €/m² eine vorsichtige Annahme. Mit dem Gutachterausschuss/Maklern vor Ort prüfen!", "Rechner")
ws["B27"].comment = Comment("Privat: Verkauf innerhalb von 10 Jahren = Einkommensteuer mit persönlichem Steuersatz (§ 23 EStG). Steuerfrei, wenn du im Verkaufsjahr und den beiden Jahren davor selbst darin wohnst. Ab mehr als 3 Objekten in 5 Jahren droht gewerblicher Grundstückshandel (zusätzlich Gewerbesteuer). Steuerberater fragen!", "Rechner")
ws["B50"].comment = Comment("Höchster Kaufpreis, bei dem der Gewinn vor Steuern noch die Zielmarge erreicht – bei gleichen Sanierungs-, Halte- und Verkaufsannahmen. Das ist dein Limit für die Verhandlung.", "Rechner")
ws["B20"].comment = Comment("Anteil von Kauf + Nebenkosten + Sanierung, den die Bank finanziert. Bei Flips meist 60–80 %; beim selbstgenutzten Haus (Live-in) bis 90 %.", "Rechner")

notes = ["So liest du das Blatt:",
         "• Bewertung „Deal“ = Marge ≥ Zielmarge. „Grenzwertig“ = 8 % bis Zielmarge. Darunter: nicht zu diesem Preis kaufen.",
         "• Wichtigste Zeile ist „Max. Kaufpreis für die Zielmarge“: Das ist dein Höchstgebot. Der Abschlag zeigt, wie weit du runterhandeln musst.",
         "• Die Spalte Live-in-Flip zeigt, warum Selbstnutzung für den Start so stark ist: keine Steuer, bessere Finanzierung, gesparte Miete, Eigenleistung.",
         "• Keine Steuer- oder Rechtsberatung. Annahmen zu Verkaufspreisen mit Gutachterausschuss Landkreis Vechta / Diepholz und Maklern vor Ort prüfen."]
for i, n in enumerate(notes):
    c = ws[f"B{54 + i}"]
    c.value = n
    c.font = BOLD if i == 0 else GREY
    ws.merge_cells(f"B{54 + i}:H{54 + i}")

ws.column_dimensions["A"].width = 2
ws.column_dimensions["B"].width = 46
for c in cols:
    ws.column_dimensions[c].width = 21
ws.freeze_panes = "C4"
dv = DataValidation(type="decimal", operator="between", formula1="0", formula2="1", showErrorMessage=True,
                    errorTitle="Prozentwert", error="Bitte einen Prozentwert zwischen 0 % und 100 % eingeben.")
ws.add_data_validation(dv)
for r in (13, 14, 15, 18, 20, 21, 25, 27, 28):
    dv.add(f"{cols[0]}{r}:{cols[-1]}{r}")
ws.page_setup.orientation = "landscape"
ws.page_setup.paperSize = ws.PAPERSIZE_A4
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True

# ---------------- Sanierungskosten ----------------
s = wb.create_sheet("Sanierungskosten")
s.sheet_view.showGridLines = False
s["B1"] = "Sanierungskosten & Wertwirkung – Richtwerte 2026"
s["B1"].font = Font(name=F, size=14, bold=True)
s["B2"] = ("Bundesweite Richtwerte aus Ratgebern/Kostenrechnern 2026 (Spannen weichen stark ab). Für jedes Objekt Angebote von Handwerkern aus der "
           "Region einholen. Priorität = wie stark die Maßnahme den Verkaufspreis im Verhältnis zu den Kosten hebt.")
s["B2"].font = GREY
s.merge_cells("B2:G2")
s["B2"].alignment = Alignment(wrap_text=True)
s.row_dimensions[2].height = 28
heads = ["Maßnahme", "Kosten (Richtwert)", "Einheit", "Wirkung auf Verkauf", "Priorität beim Flip", "Hinweis"]
for i, h in enumerate(heads):
    c = s.cell(row=4, column=2 + i, value=h)
    c.font = HDR
    c.fill = HFILL
    c.alignment = Alignment(wrap_text=True, vertical="center")
rows = [
    ("Streichen, Tapeten, Spachteln", "15–30 €", "pro m² Wohnfläche", "hoch (erster Eindruck)", "Muss", "Gut in Eigenleistung machbar."),
    ("Bodenbeläge (Vinyl/Laminat, Fliesen ausbessern)", "40–80 €", "pro m² Bodenfläche", "hoch", "Muss", "Einheitlicher, heller Boden wirkt wie neu."),
    ("Bad komplett erneuern", "12.000–25.000 €", "pro Bad", "sehr hoch", "Muss bei altem Bad", "Kein Luxus – solide, zeitlos, bodengleiche Dusche."),
    ("Gäste-WC", "5.000–9.000 €", "pauschal", "mittel", "Kann", "Nur wenn vorhanden oder günstig nachrüstbar."),
    ("Küche", "5.000–12.000 €", "pauschal", "gering bis mittel", "Meist weglassen", "Käufer wollen eigene Küche; nur Anschlüsse vorbereiten."),
    ("Elektrik erneuern", "90–140 €", "pro m² Wohnfläche", "mittel (Sicherheit)", "Muss bei alter Elektrik ohne FI", "Nur Fachbetrieb. Käufer & Gutachter fragen gezielt danach."),
    ("Fenster & Außentüren", "12.000–22.500 €", "für ca. 15 Fenster", "hoch (Energieklasse, Optik)", "Oft lohnend", "Klassen-Sprung im Energieausweis bringt Preisaufschlag."),
    ("Heizung: Wärmepumpe", "25.000–40.000 €", "pauschal", "hoch (Energieklasse)", "Je nach Objekt", "Förderung für Vermieter/Investoren nur 30 % Grundförderung (BEG seit 21.07.2026); Boni nur für Selbstnutzer."),
    ("Dämmung oberste Geschossdecke", "30–60 €", "pro m² Deckenfläche", "mittel bis hoch", "Fast immer", "Günstigste Maßnahme für eine bessere Energieklasse."),
    ("Dach neu eindecken + dämmen", "150–300 €", "pro m² Dachfläche", "hoch, aber teuer", "Nur wenn nötig", "Großer Kostenblock – vor Kauf prüfen lassen!"),
    ("Fassadendämmung (WDVS)", "150–250 €", "pro m² Fassade", "mittel", "Selten beim Flip", "Lohnt sich meist nur bei Klasse G/H und langer Haltedauer."),
    ("Kellerabdichtung / Feuchtigkeit", "15.000–40.000 €", "pauschal", "verhindert Abschläge", "Pflicht bei Feuchte", "Ursache klären lassen (Sachverständiger)."),
    ("Schimmel- / Wasserschadensanierung", "5.000–50.000 € +", "pauschal", "verhindert Abschläge", "Pflicht", "Sehr schwer kalkulierbar – Gutachten vor Kauf!"),
    ("Grundriss öffnen (nicht tragende Wand)", "2.000–6.000 €", "pro Wand", "mittel bis hoch", "Kann", "Offene Küche/Wohnen ist gefragt; tragende Wände mit Statiker (8.000–20.000 €)."),
    ("Garten & Eingangsbereich", "2.000–8.000 €", "pauschal", "hoch (erster Eindruck)", "Muss", "Rückschnitt, Rasen, Pflaster reinigen, Haustür."),
    ("Energieausweis (Bedarfsausweis)", "300–600 €", "pauschal", "Pflicht im Inserat", "Muss", "Nach der Sanierung neu ausstellen lassen."),
    ("Home Staging & Profi-Fotos", "1.000–3.000 €", "pauschal", "hoch", "Muss", "Leere Räume wirken kleiner; gute Fotos = mehr Anfragen."),
]
for i, row in enumerate(rows):
    r = 5 + i
    for j, v in enumerate(row):
        c = s.cell(row=r, column=2 + j, value=v)
        c.font = BOLD if j == 0 else BLK
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c.border = Border(bottom=Side(style="thin", color="D5D2CC"))
    s.row_dimensions[r].height = 30
r = 5 + len(rows) + 1
s[f"B{r}"] = "Energieklasse und Preis:"
s[f"B{r}"].font = BOLD
s[f"B{r+1}"] = ("Laut immowelt (2026) werden Häuser der Klasse G ca. 10 % und Klasse H ca. 17 % günstiger angeboten als Klasse D. Das IW (Wohnindex Q1/2026) "
                "misst gegenüber A/A+ Abschläge von rund 20–30 % für G/H. Ein Sprung von G/H auf D/C ist deshalb oft der stärkste Hebel beim Flip.")
s[f"B{r+1}"].font = GREY
s.merge_cells(f"B{r+1}:G{r+1}")
s[f"B{r+1}"].alignment = Alignment(wrap_text=True)
s.row_dimensions[r + 1].height = 42
s[f"B{r+3}"] = "Faustregeln für das Sanierungsbudget pro m² Wohnfläche: Auffrischung 200–400 € · mittlere Sanierung 400–700 € · Komplettsanierung 800–1.400 €"
s[f"B{r+3}"].font = BOLD
s.merge_cells(f"B{r+3}:G{r+3}")
for col, w in {"A": 2, "B": 34, "C": 18, "D": 20, "E": 22, "F": 22, "G": 52}.items():
    s.column_dimensions[col].width = w
s.freeze_panes = "B5"

wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print("ok")
