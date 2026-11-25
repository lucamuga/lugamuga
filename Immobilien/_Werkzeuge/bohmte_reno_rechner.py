"""Excel nur für die ETW Bohmte: Fix & Flip (Vorlage des Nutzers) + Renovierungs-Raster + Material + Gantt.

Aufruf: python3 bohmte_reno_rechner.py VORLAGE.xlsx AUSGABE_roh.xlsx
Danach: python3 excel_werte_speichern.py AUSGABE_roh.xlsx ZIEL.xlsx  (rechnet und speichert die Werte)
"""
import sys
from datetime import date
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter as L

VORLAGE, OUT = sys.argv[1], sys.argv[2]

BLAU = PatternFill("solid", fgColor="DDEBF7")      # Eingabe (wie in der Vorlage: hellblau)
KOPF = PatternFill("solid", fgColor="1F3A5F")
GRAU = PatternFill("solid", fgColor="F2F2F2")
SUMME = PatternFill("solid", fgColor="FCE4D6")
HF = Font(bold=True, color="FFFFFF")
B = Font(bold=True)
TITEL = Font(bold=True, size=14)
KLEIN = Font(italic=True, size=9, color="666666")
DUENN = Side(style="thin", color="BFBFBF")
RAHMEN = Border(top=DUENN, bottom=DUENN, left=DUENN, right=DUENN)
EUR = '#,##0 "€";[Red]-#,##0 "€"'
EUR2 = '#,##0.00 "€"'
DAT = "DD.MM.YY"

wb = openpyxl.load_workbook(VORLAGE)
for n in [s for s in wb.sheetnames if s not in ("Kalkulation", "DISCLAIMER")]:
    del wb[n]
ff = wb["Kalkulation"]
ff.title = "1 Fix & Flip"


def kopf(ws, row, cols, texte):
    for c, t in zip(cols, texte):
        z = ws.cell(row=row, column=c, value=t)
        z.font, z.fill, z.alignment, z.border = HF, KOPF, Alignment(wrap_text=True, vertical="center"), RAHMEN


def eingabe(z, wert, fmt=None):
    z.value, z.fill, z.border = wert, BLAU, RAHMEN
    if fmt:
        z.number_format = fmt


# =====================================================================
# 2 Bewertung & Maße
# =====================================================================
bw = wb.create_sheet("2 Bewertung & Maße")
bw["A1"] = "ETW Bohmte – Zustandsbewertung und Maße"; bw["A1"].font = TITEL
bw["A2"] = ("Eigene Einschätzung aus der Anzeige (Bj. 1964, „teilweise modernisiert“, Gas 2019, Klasse E, vermietet) – "
            "noch ohne Besichtigung. Hellblau = änderbar. Alle Mengen im Material-Blatt rechnen mit diesen Maßen.")
bw["A2"].font = KLEIN
kopf(bw, 4, range(1, 4), ["Maß", "Wert", "Herleitung"])
MASSE = [  # (Text, Wert/Formel, Format, Herleitung, eingabe?)
    ("Wohnfläche (m²)", 80, "0.0", "Anzeige", True),
    ("Raumhöhe (m)", 2.5, "0.00", "üblich Baujahr 1964", True),
    ("Bad: Bodenfläche (m²)", 6, "0.0", "Annahme – beim Aufmaß prüfen", True),
    ("Bad: Fliesenhöhe Wand (m)", 2.0, "0.00", "raumhoch wäre 2,50", True),
    ("Bad: Wandumfang (m)", 10, "0.0", "ca. 2,0 × 3,0 m Raum", True),
    ("Bad: Wandfliesen (m²)", "=B9*B8-2", "0.0", "Umfang × Höhe − Tür", False),
    ("Wandfläche gesamt (m²)", "=B5*2.6", "0", "Faustwert: Wohnfläche × 2,6", False),
    ("Wandfläche zu tapezieren/streichen (m²)", "=B11-B10", "0", "Wand gesamt − Bad-Fliesen", False),
    ("Deckenfläche (m²)", "=B5", "0", "gleich Wohnfläche", False),
    ("Bodenbelag Vinyl (m²)", "=B5-B7", "0.0", "Wohnfläche − Bad", False),
    ("Sockelleisten (lfm)", "=ROUND(B5*1.1,0)", "0", "Faustwert: Wohnfläche × 1,1", False),
    ("Innentüren (Stück)", 6, "0", "4 Zimmer + Bad + Küche", True),
    ("Fenster (Stück)", 7, "0", "Annahme", True),
    ("Steckdosen + Schalter (Stück)", 45, "0", "ca. 6–8 je Raum", True),
    ("Heizkörper (Stück)", 6, "0", "Annahme", True),
    ("Sperrmüll (m³)", "=ROUND(B5*0.08,1)", "0.0", "Faustwert leere Wohnung + Böden/Türen", False),
    ("Bauschutt Fliesen (m³)", "=ROUND((B10+B7)*0.03*1.5,1)", "0.0", "Fliesen + Kleber, aufgelockert", False),
]
for i, (t, v, f, h, inp) in enumerate(MASSE):
    r = 5 + i
    bw.cell(row=r, column=1, value=t).border = RAHMEN
    z = bw.cell(row=r, column=2)
    if inp:
        eingabe(z, v, f)
    else:
        z.value, z.number_format, z.border = v, f, RAHMEN
    bw.cell(row=r, column=3, value=h).font = KLEIN
M = {k: f"'2 Bewertung & Maße'!$B${5+i}" for i, k in enumerate(
    ["wfl", "hoehe", "badboden", "fliesh", "badumf", "badwand", "wand", "wandstreich", "decke", "boden",
     "sockel", "tueren", "fenster", "dosen", "hk", "sperr", "schutt"])}

r0 = 24
bw.cell(row=r0, column=1, value="Optionen").font = B
OPT = [("Fenster tauschen? (oft Gemeinschaftseigentum!)", "Nein"), ("Küchenzeile einbauen?", "Nein"),
       ("Küchenzeile Preis (€)", 3000)]
dv = DataValidation(type="list", formula1='"Ja,Nein"', allow_blank=False)
bw.add_data_validation(dv)
for i, (t, v) in enumerate(OPT):
    bw.cell(row=r0 + 1 + i, column=1, value=t)
    eingabe(bw.cell(row=r0 + 1 + i, column=2), v, EUR if isinstance(v, int) else None)
    if isinstance(v, str):
        dv.add(bw.cell(row=r0 + 1 + i, column=2))
O_FENSTER, O_KUECHE, O_KPREIS = (f"'2 Bewertung & Maße'!$B${r0+1+i}" for i in range(3))

r1 = 30
bw.cell(row=r1, column=1, value="Zustandsbewertung je Bauteil").font = B
kopf(bw, r1 + 1, range(1, 8), ["Bauteil", "Eigentum", "Zustand (geschätzt)", "Note 1–5", "Maßnahme", "Einstufung", "Wer"])
BEW = [
    ("Heizung (Gas 2019)", "Gemeinschaft", "neu, gut", 1, "keine – nur Heizkörper/Ventile in der Wohnung", "Optional", "Fremd (Heizung)"),
    ("Heizkörper / Thermostate", "Sondereigentum*", "vermutlich alt", 3, "Bad-Heizkörper + Thermostatköpfe tauschen", "Empfohlen", "Fremd (Heizung)"),
    ("Fenster", "meist Gemeinschaft*", "unbekannt, evtl. Kunststoff 80er/90er", 3, "nur prüfen; Tausch braucht WEG-Beschluss", "Optional", "Fremd (Fenster)"),
    ("Elektrik", "Sondereigentum", "teils original 1964 (ohne FI?)", 4, "neue Unterverteilung mit FI, Leitungen teils neu, neue Schalter/Dosen", "Pflicht (Sicherheit)", "Eigen (Elektriker im Team)"),
    ("Bad", "Sondereigentum", "vermutlich 80er/90er", 4, "komplett neu: Rohinstallation, Abdichtung, Fliesen, Objekte", "Empfohlen (Werttreiber)", "Sanitär fremd, Fliesen eigen"),
    ("Küche", "Sondereigentum", "Anschlüsse alt", 3, "Anschlüsse erneuern; Küchenzeile optional", "Empfohlen", "Sanitär fremd"),
    ("Böden", "Sondereigentum", "alt (Teppich/Laminat?)", 4, "raus, ausgleichen, Vinyl klick", "Empfohlen (Werttreiber)", "Eigen"),
    ("Wände / Decken", "Sondereigentum", "Tapete, Gebrauchsspuren", 3, "Tapete runter, spachteln, Vlies + 2× streichen", "Empfohlen", "Eigen"),
    ("Innentüren", "Sondereigentum", "Originaltüren 1964", 4, "neue Türen mit Zarge", "Empfohlen (günstig, große Wirkung)", "Eigen"),
    ("Wohnungseingangstür", "meist Gemeinschaft*", "unbekannt", 3, "bleibt", "–", "–"),
    ("Keller / Stellplätze", "Sondereigentum/Sondernutzung", "unbekannt", 2, "entrümpeln, reinigen", "Pflicht", "Eigen"),
]
for i, row in enumerate(BEW):
    for c, v in enumerate(row, 1):
        z = bw.cell(row=r1 + 2 + i, column=c, value=v)
        z.border, z.alignment = RAHMEN, Alignment(wrap_text=True, vertical="top")
rg = r1 + 2 + len(BEW) + 1
bw.cell(row=rg, column=1, value="Gesamturteil").font = B
bw.cell(row=rg, column=2, value=("Mittlerer Renovierungsbedarf: Bad, Elektrik, Böden, Wände und Türen erneuern. "
                                 "Heizung ist neu; Fenster und Dach betreffen die WEG. Wohnung ist vermietet – "
                                 "die Renovierung geht erst, wenn sie leer ist."))
bw.cell(row=rg + 1, column=1, value="* laut Teilungserklärung prüfen").font = KLEIN
bw.cell(row=rg + 2, column=1, value="Renovierung je m² (aus Blatt 3)")
bw.cell(row=rg + 2, column=2, value="='3 Renovierung'!C52/B5").number_format = EUR
for c, w in zip("ABCDEFG", [40, 22, 32, 9, 48, 26, 26]):
    bw.column_dimensions[c].width = w

# =====================================================================
# 4 Material (vor Blatt 3 angelegt, Reihenfolge später sortiert)
# =====================================================================
ma = wb.create_sheet("4 Material")
ma["A1"] = "Material- und Entsorgungsliste (Mengen automatisch aus den Maßen)"; ma["A1"].font = TITEL
ma["A2"] = ("Preise: Baumarkt/Online-Schätzwerte 2026, inkl. MwSt. Hellblau = änderbar. "
            "Nr = Arbeitsschritt im Blatt „3 Renovierung“."); ma["A2"].font = KLEIN
kopf(ma, 4, range(1, 13), ["Nr", "Position", "Bezug", "Menge netto", "Einheit", "Verschnitt", "Menge inkl. Verschnitt",
                           "Gebinde (Menge/Pack)", "Packungen", "Preis je Pack", "Summe", "Hinweis"])
MAT = [  # Nr, Position, Bezug, Menge(Formel), Einheit, Verschnitt, Gebinde, Preis, Hinweis
    (2, "Container Sperrmüll 7 m³ (inkl. Stellen/Abholen)", "Sperrmüll m³", f"={M['sperr']}", "m³", 0, 7, 520, "Holz/Böden/Möbel getrennt günstiger"),
    (2, "Container Bauschutt / Fliesen 3 m³", "Bauschutt m³", f"={M['schutt']}", "m³", 0, 3, 380, "nur sortenreiner Schutt!"),
    (2, "Big Bags / Säcke, Wertstoffhof (Elektro, Metall)", "pauschal", 1, "psch", 0, 1, 80, "Elektro-/Metallschrott kostenlos am Wertstoffhof"),
    (2, "Stellgenehmigung öffentliche Fläche", "pauschal", 0, "psch", 0, 1, 60, "nur nötig, wenn kein eigener Stellplatz"),
    (3, "Abdeckfolie, Krepp, Schuttsäcke", "pauschal", 1, "psch", 0, 1, 90, ""),
    (3, "Trennscheiben, Meißel (Verschleiß), Tapetenlöser", "pauschal", 1, "psch", 0, 1, 70, ""),
    (5, "Unterverteilung 3-reihig, bestückt (FI/LS)", "Stück", 1, "Stk", 0, 1, 450, "Abnahme durch Elektriker"),
    (5, "NYM-J 3×1,5 mm²", "Wohnfläche × 3 m", f"={M['wfl']}*3", "m", 0.1, 100, 95, "100-m-Ring"),
    (5, "NYM-J 5×2,5 mm² (Herd)", "Stück", 15, "m", 0.1, 25, 60, "25-m-Ring"),
    (5, "Unterputz-/Abzweigdosen", "Steckdosen + Schalter", f"={M['dosen']}+10", "Stk", 0.1, 25, 22, ""),
    (5, "Schnellputz/Gips für Schlitze", "Wohnfläche", f"={M['wfl']}/20", "Sack", 0, 1, 12, "20 kg"),
    (8, "Fertigspachtel", "Wand + Decke", f"={M['wandstreich']}+{M['decke']}", "m²", 0, 60, 26, "20-kg-Eimer ≈ 60 m² bei Teilspachtelung"),
    (8, "Tiefgrund", "Wand + Decke", f"={M['wandstreich']}+{M['decke']}", "m²", 0, 80, 32, "10 l ≈ 80 m²"),
    (8, "Schleifgitter, Gewebe, Kantenschutz", "pauschal", 1, "psch", 0, 1, 60, ""),
    (9, "Wandfliesen", "Bad-Wandfläche", f"={M['badwand']}", "m²", 0.1, 1, 25, "einfache Serie 30×60"),
    (9, "Bodenfliesen", "Bad-Boden", f"={M['badboden']}", "m²", 0.1, 1, 30, "rutschhemmend R10"),
    (9, "Flexkleber (ca. 4 kg/m²)", "Fliesenfläche", f"=({M['badwand']}+{M['badboden']})*4", "kg", 0.1, 25, 18, "25-kg-Sack"),
    (9, "Fugenmörtel", "Fliesenfläche", f"=({M['badwand']}+{M['badboden']})*0.5", "kg", 0.1, 5, 16, ""),
    (9, "Abdichtung (Dichtschlämme, Dichtband, Manschetten)", "Set", 1, "Set", 0, 1, 130, "Pflicht im Duschbereich"),
    (9, "Silikon, Fliesenschienen, Kreuze", "pauschal", 1, "psch", 0, 1, 70, ""),
    (11, "Malervlies", "Wand + Decke", f"={M['wandstreich']}+{M['decke']}", "m²", 0.1, 25, 32, "Rolle 25 m²"),
    (11, "Vlieskleber", "Wand + Decke", f"={M['wandstreich']}+{M['decke']}", "m²", 0.1, 80, 38, "16 kg ≈ 80 m²"),
    (11, "Wandfarbe weiß (2 Anstriche, 7 m²/l)", "Wand + Decke × 2", f"=({M['wandstreich']}+{M['decke']})*2/7", "l", 0.05, 12.5, 55, "12,5-l-Eimer"),
    (11, "Acryl, Abdeckmaterial, Rollen/Pinsel", "pauschal", 1, "psch", 0, 1, 90, ""),
    (12, "Vinyl-Klickboden mit Trittschall", "Bodenfläche", f"={M['boden']}", "m²", 0.08, 1, 24, "Nutzungsklasse 23/31"),
    (12, "Ausgleichsmasse (Teilflächen)", "Bodenfläche", f"={M['boden']}*1.5", "kg", 0, 25, 18, "≈ 1,5 kg/m² im Schnitt"),
    (12, "Sockelleisten 2,5 m", "lfm", f"={M['sockel']}", "m", 0.1, 2.5, 7, ""),
    (12, "Übergangsprofile", "Türen", f"={M['tueren']}", "Stk", 0, 1, 12, ""),
    (13, "Innentür komplett mit Zarge", "Türen", f"={M['tueren']}", "Stk", 0, 1, 190, "Weißlack, Röhrenspan"),
    (13, "Drückergarnitur", "Türen", f"={M['tueren']}", "Stk", 0, 1, 25, ""),
    (13, "Montageschaum, Keile", "pauschal", 1, "psch", 0, 1, 40, ""),
    (14, "Schalter/Steckdosen inkl. Rahmen", "Steckdosen + Schalter", f"={M['dosen']}", "Stk", 0.05, 1, 8, "einfache Serie"),
    (14, "Leuchten/Fassungen (Verkaufszustand)", "Räume", 8, "Stk", 0, 1, 15, ""),
    (14, "Prüfprotokoll / Abnahme Elektriker", "pauschal", 1, "psch", 0, 1, 350, "Dienstleistung, Pflicht"),
    (16, "Reinigungsmittel, Endreinigung", "pauschal", 1, "psch", 0, 1, 50, ""),
    (16, "Werkzeugmiete (z. B. Nassschneider, Bautrockner)", "pauschal", 1, "psch", 0, 1, 120, "falls nicht vorhanden"),
]
MA0 = 5
for i, (nr, pos, bez, menge, eh, vs, geb, preis, hin) in enumerate(MAT):
    r = MA0 + i
    ma.cell(row=r, column=1, value=nr)
    ma.cell(row=r, column=2, value=pos)
    ma.cell(row=r, column=3, value=bez).font = KLEIN
    z = ma.cell(row=r, column=4, value=menge); z.number_format = "0.0"
    if not (isinstance(menge, str) and menge.startswith("=")):
        z.fill = BLAU
    ma.cell(row=r, column=5, value=eh)
    eingabe(ma.cell(row=r, column=6), vs, "0%")
    ma.cell(row=r, column=7, value=f"=D{r}*(1+F{r})").number_format = "0.0"
    eingabe(ma.cell(row=r, column=8), geb, "0.0")
    ma.cell(row=r, column=9, value=f"=IF(D{r}<=0,0,IF(H{r}>0,ROUNDUP(G{r}/H{r},0),G{r}))").number_format = "0"
    eingabe(ma.cell(row=r, column=10), preis, EUR2)
    ma.cell(row=r, column=11, value=f"=I{r}*J{r}").number_format = EUR
    ma.cell(row=r, column=12, value=hin).font = KLEIN
    for c in range(1, 13):
        ma.cell(row=r, column=c).border = RAHMEN
MA1 = MA0 + len(MAT) - 1
ma.cell(row=MA1 + 2, column=2, value="Summe Material und Entsorgung").font = B
z = ma.cell(row=MA1 + 2, column=11, value=f"=SUM(K{MA0}:K{MA1})"); z.number_format, z.font, z.fill = EUR, B, SUMME
for c, w in zip("ABCDEFGHIJKL", [5, 46, 20, 11, 7, 10, 12, 12, 10, 12, 12, 40]):
    ma.column_dimensions[c].width = w
ma.freeze_panes = "C5"

# =====================================================================
# 3 Renovierung (Raster)
# =====================================================================
re = wb.create_sheet("3 Renovierung")
re["A1"] = "Renovierungs-Raster: Schritte, Team, Fahrten, Zeitplan"; re["A1"].font = TITEL
re["A2"] = ("Hellblau = änderbar. Dauer in Arbeitstagen (Mo–Fr). Start ergibt sich aus den Vorgängern; "
            "Gantt (Blatt 5) und Fix & Flip (Blatt 1) passen sich automatisch an."); re["A2"].font = KLEIN
EIN = [
    ("Projektstart (Kaufvertrag beim Notar)", date(2027, 1, 11), DAT, "Annahme"),
    ("Stundensatz Helfer/Monteure (€/h)", 20, EUR2, "was du deinen Leuten zahlst"),
    ("Arbeitsstunden pro Tag", 8, "0.0", ""),
    ("Entfernung Damme → Bohmte, einfach (km)", 25, "0", "ca.-Wert, bitte prüfen"),
    ("Kilometersatz (€/km)", 0.30, EUR2, "steuerliche Pauschale"),
    ("Fahrtzeit einfach (Minuten)", 30, "0", ""),
    ("Fahrtzeit bezahlen?", "Ja", None, "Ja = Helfer bekommen Fahrtzeit bezahlt"),
    ("Kostenpuffer auf Renovierung", 0.15, "0%", "geht in Blatt 1 als Sicherheitspuffer"),
    ("Zeitpuffer auf Renovierungsdauer", 0.15, "0%", ""),
    ("Kleinmaterial-Zuschlag auf Material", 0.05, "0%", "Schrauben, Dübel, Kleinkram"),
]
for i, (t, v, f, h) in enumerate(EIN):
    r = 4 + i
    re.cell(row=r, column=1, value=t)
    eingabe(re.cell(row=r, column=3), v, f)
    re.cell(row=r, column=4, value=h).font = KLEIN
dv2 = DataValidation(type="list", formula1='"Ja,Nein"'); re.add_data_validation(dv2); dv2.add("C10")
START, SATZ, STD, KM, KMS, FZ, FZB, KP, ZP, KLM = (f"$C${4+i}" for i in range(10))

H = 16
kopf(re, H, range(1, 22), ["Nr", "Arbeitsschritt", "Art", "Vorgänger 1", "Vorgänger 2", "Dauer (AT)", "Start", "Ende",
                           "Helfer bezahlt (Pers.)", "Ich selbst (0/1)", "Fahrzeuge je Tag", "Std. Helfer", "Lohn Helfer",
                           "Fahrtkosten", "Fahrtzeit-Lohn", "Material", "Fremdfirma (Angebot)", "Summe", "Eigene Std.",
                           "Fahrten (Tage)", "Hinweis"])
re.row_dimensions[H].height = 45
dv3 = DataValidation(type="list", formula1='"Eigen,Fremd,Kauf,Verkauf,Puffer"'); re.add_data_validation(dv3)
T0 = H + 1
TASKS = [  # Nr, Schritt, Art, V1, V2, Dauer, Helfer, ich, Fahrzeuge, Fremd, Hinweis
    (0, "Kauf: Notar, Finanzierung, Grundbuch → Übergabe (Wohnung leer)", "Kauf", None, None, 30, 0, 0, 0, 0, "Wohnung ist vermietet! Erst nach Auszug"),
    (1, "Übergabe, Aufmaß, Termine mit Fremdfirmen", "Eigen", 0, None, 2, 0, 1, 1, 0, ""),
    (2, "Entrümpelung inkl. Container und Trennung", "Eigen", 1, None, 2, 1, 1, 1, 0, "Container-Kosten im Material-Blatt"),
    (3, "Demontage & Abbruch (Fliesen, Böden, Tapete, Türen)", "Eigen", 2, None, 4, 2, 1, 1, 0, ""),
    (4, "Sanitär Rohinstallation (Bad, Küche)", "Fremd", 3, None, 3, 0, 0, 0, 3500, "Angebot einholen"),
    (5, "Elektro Rohinstallation (Schlitze, Leitungen, Verteilung)", "Eigen", 3, None, 5, 1, 1, 1, 0, "Elektriker aus deinem Team"),
    (6, "Heizung: Bad-Heizkörper, Thermostate, Abgleich", "Fremd", 3, None, 2, 0, 0, 0, 1500, "Angebot einholen"),
    (7, "Fenster (nur wenn Option = Ja)", "Fremd", 3, None, f"=IF({O_FENSTER}=\"Ja\",2,0)", 0, 0, 0, f"=IF({O_FENSTER}=\"Ja\",{M['fenster']}*850,0)", "WEG-Beschluss nötig"),
    (8, "Spachteln, Putz, Trockenbau", "Eigen", 4, 5, 4, 2, 1, 1, 0, ""),
    (9, "Bad: Abdichtung und Fliesen", "Eigen", 8, None, 5, 1, 1, 1, 0, ""),
    (10, "Sanitär Endmontage (WC, Waschtisch, Dusche, Armaturen)", "Fremd", 9, None, 2, 0, 0, 0, 4500, "inkl. Sanitärobjekte – Angebot"),
    (11, "Malerarbeiten (Vlies + 2× streichen)", "Eigen", 8, None, 6, 2, 1, 1, 0, ""),
    (12, "Bodenbelag Vinyl + Sockelleisten", "Eigen", 11, None, 4, 1, 1, 1, 0, ""),
    (13, "Innentüren einbauen", "Eigen", 12, None, 2, 1, 1, 1, 0, ""),
    (14, "Elektro Feininstallation + Prüfprotokoll", "Eigen", 11, None, 2, 1, 1, 1, 0, ""),
    (15, "Küchenzeile (nur wenn Option = Ja)", "Fremd", 12, None, f"=IF({O_KUECHE}=\"Ja\",1,0)", 0, 0, 0, f"=IF({O_KUECHE}=\"Ja\",{O_KPREIS},0)", "Lieferung + Montage"),
    (16, "Endreinigung, Verkaufsfotos", "Eigen", 13, 14, 1, 1, 1, 1, 0, ""),
    (17, "Zeitpuffer", "Puffer", 16, 10, None, 0, 0, 0, 0, "automatisch: Zeitpuffer × Renovierungsdauer"),
    (18, "Vermarktung, Besichtigungen, Verkauf bis Zahlung", "Verkauf", 17, None, 40, 0, 0, 0, 0, "ca. 8 Wochen"),
]
T1 = T0 + len(TASKS) - 1
RNG = f"$A${T0}:$H${T1}"
for i, (nr, txt, art, v1, v2, dauer, helfer, ich, fz, fremd, hin) in enumerate(TASKS):
    r = T0 + i
    re.cell(row=r, column=1, value=nr)
    re.cell(row=r, column=2, value=txt)
    eingabe(re.cell(row=r, column=3), art); dv3.add(re.cell(row=r, column=3))
    eingabe(re.cell(row=r, column=4), v1)
    eingabe(re.cell(row=r, column=5), v2)
    if nr == 17:  # Zeitpuffer automatisch
        re.cell(row=r, column=6, value=f"=ROUNDUP(NETWORKDAYS(G{T0+1},MAX(H{T0+1}:H{r-1}))*{ZP},0)")
    elif isinstance(dauer, str):
        re.cell(row=r, column=6, value=dauer)
    else:
        eingabe(re.cell(row=r, column=6), dauer)
    vor = (f"MAX(IF(D{r}=\"\",0,VLOOKUP(D{r},{RNG},8,0)),IF(E{r}=\"\",0,VLOOKUP(E{r},{RNG},8,0)))")
    re.cell(row=r, column=7, value=f"=IF(AND(D{r}=\"\",E{r}=\"\"),WORKDAY({START}-1,1),WORKDAY({vor},1))").number_format = DAT
    re.cell(row=r, column=8, value=f"=IF(F{r}>0,WORKDAY(G{r},F{r}-1),G{r}-1)").number_format = DAT
    for c, v in ((9, helfer), (10, ich), (11, fz)):
        eingabe(re.cell(row=r, column=c), v)
    eig = f'C{r}="Eigen"'
    re.cell(row=r, column=12, value=f"=IF({eig},F{r}*I{r}*{STD},0)").number_format = "0"
    re.cell(row=r, column=13, value=f"=L{r}*{SATZ}").number_format = EUR
    re.cell(row=r, column=14, value=f"=IF({eig},F{r}*K{r}*2*{KM}*{KMS},0)").number_format = EUR
    re.cell(row=r, column=15, value=f"=IF(AND({eig},{FZB}=\"Ja\"),F{r}*I{r}*2*{FZ}/60*{SATZ},0)").number_format = EUR
    re.cell(row=r, column=16, value=f"=SUMIF('4 Material'!$A${MA0}:$A${MA1},A{r},'4 Material'!$K${MA0}:$K${MA1})").number_format = EUR
    z = re.cell(row=r, column=17)
    if isinstance(fremd, str):
        z.value = fremd
    else:
        eingabe(z, fremd)
    z.number_format = EUR
    re.cell(row=r, column=18, value=f"=SUM(M{r}:Q{r})").number_format = EUR
    re.cell(row=r, column=19, value=f"=IF({eig},F{r}*J{r}*{STD},0)").number_format = "0"
    re.cell(row=r, column=20, value=f"=IF({eig},F{r},0)").number_format = "0"
    re.cell(row=r, column=21, value=hin).font = KLEIN
    for c in range(1, 22):
        re.cell(row=r, column=c).border = RAHMEN
    if art in ("Kauf", "Verkauf", "Puffer"):
        for c in (1, 2):
            re.cell(row=r, column=c).fill = GRAU
rs = T1 + 1
re.cell(row=rs, column=2, value="Summe").font = B
for c in range(12, 21):
    z = re.cell(row=rs, column=c, value=f"=SUM({L(c)}{T0}:{L(c)}{T1})")
    z.font, z.fill = B, SUMME
    z.number_format = "0" if c in (12, 19, 20) else EUR

# Zusammenfassung
S0 = 41
re.cell(row=S0, column=1, value="Zusammenfassung Kosten").font = TITEL
ZUS = [
    ("Material und Entsorgung", f"=P{rs}"),
    ("Kleinmaterial-Zuschlag", f"=P{rs}*{KLM}"),
    ("Fremdfirmen (Heizung, Sanitär, Fenster, Küche)", f"=Q{rs}"),
    ("Lohn Helfer/Monteure", f"=M{rs}"),
    ("Fahrtkosten (nur eigenes Team)", f"=N{rs}"),
    ("Fahrtzeit-Lohn (nur eigenes Team)", f"=O{rs}"),
    ("Zwischensumme Renovierung ohne Puffer", f"=SUM(C{S0+1}:C{S0+6})"),
    ("davon Entrümpelung (Schritt 2) → Blatt 1, Zeile Entrümpelung", f"=R{T0+2}*(1+0)"),
    ("davon Küchenzeile → Blatt 1, Zeile Küche", f"=R{T0+15}"),
    ("Kostenpuffer (ohne Entrümpelung und Küche)", f"=(C{S0+7}-C{S0+8}-C{S0+9})*{KP}"),
    ("GESAMT Renovierung inkl. Puffer", f"=C{S0+7}+C{S0+10}"),
]
for i, (t, f) in enumerate(ZUS):
    r = S0 + 1 + i
    re.cell(row=r, column=1, value=t)
    z = re.cell(row=r, column=3, value=f); z.number_format = EUR
    if t.startswith("GESAMT") or t.startswith("Zwischensumme"):
        re.cell(row=r, column=1).font = B; z.font = B; z.fill = SUMME
re.cell(row=S0 + 8, column=3).value = f"=R{T0+2}+P{T0+2}*{KLM}"
# C52: Renovierung ohne Entrümpelung, ohne Küche, ohne Puffer (für Blatt 1 Zeile „Renovierung“)
re.cell(row=S0 + 12, column=1, value="Renovierung für Blatt 1 (ohne Entrümpelung, Küche, Puffer)")
re.cell(row=S0 + 12, column=3, value=f"=C{S0+7}-C{S0+8}-C{S0+9}").number_format = EUR
assert S0 + 12 == 53

Z0 = 55
re.cell(row=Z0, column=1, value="Zeitplan").font = TITEL
ZEIT = [
    ("Projektstart (Notar)", f"=G{T0}", DAT),
    ("Renovierung von", f"=G{T0+1}", DAT),
    ("Renovierung bis (inkl. Zeitpuffer)", f"=H{T0+17}", DAT),
    ("Renovierung Arbeitstage inkl. Puffer", f"=NETWORKDAYS(C{Z0+2},C{Z0+3})", "0"),
    ("Verkauf abgeschlossen (Geld da)", f"=H{T1}", DAT),
    ("Projektdauer gesamt (Monate) → Blatt 1", f"=ROUNDUP((C{Z0+5}-C{Z0+1}+1)/30.4,0)", "0"),
    ("Deine eigenen Arbeitsstunden", f"=S{rs}", "0"),
    ("Helfer-Stunden bezahlt", f"=L{rs}", "0"),
]
for i, (t, f, fm) in enumerate(ZEIT):
    r = Z0 + 1 + i
    re.cell(row=r, column=1, value=t)
    z = re.cell(row=r, column=3, value=f); z.number_format = fm
MONATE = f"'3 Renovierung'!$C${Z0+6}"
for c, w in zip(range(1, 22), [5, 50, 9, 10, 10, 9, 10, 10, 10, 9, 9, 9, 11, 11, 11, 11, 13, 12, 9, 9, 34]):
    re.column_dimensions[L(c)].width = w
re.column_dimensions["A"].width = 44
re.freeze_panes = f"C{T0}"

# =====================================================================
# 5 Gantt (Wochenansicht)
# =====================================================================
ga = wb.create_sheet("5 Zeitplan (Gantt)")
ga["A1"] = "Zeitplan – passt sich automatisch an Blatt 3 an"; ga["A1"].font = TITEL
ga["A2"] = "Grün = Eigenleistung · Blau = Fremdfirma · Grau = Kauf/Verkauf · Orange = Puffer. Eine Spalte = eine Woche (Montag)."
ga["A2"].font = KLEIN
GH, WOCHEN = 4, 36
kopf(ga, GH, range(1, 7), ["Nr", "Arbeitsschritt", "Art", "Start", "Ende", "AT"])
for w in range(WOCHEN):
    c = 7 + w
    if w == 0:
        f = f"='3 Renovierung'!$G${T0}-WEEKDAY('3 Renovierung'!$G${T0},2)+1"
    else:
        f = f"={L(c-1)}{GH}+7"
    z = ga.cell(row=GH, column=c, value=f)
    z.number_format, z.font, z.fill = "DD.MM.", Font(bold=True, color="FFFFFF", size=8), KOPF
    z.alignment = Alignment(text_rotation=90, horizontal="center")
    ga.column_dimensions[L(c)].width = 3.2
ga.row_dimensions[GH].height = 42
G0 = GH + 1
for i in range(len(TASKS)):
    r, q = G0 + i, T0 + i
    for c, src in ((1, "A"), (2, "B"), (3, "C"), (4, "G"), (5, "H"), (6, "F")):
        z = ga.cell(row=r, column=c, value=f"='3 Renovierung'!{src}{q}")
        z.border = RAHMEN
        if c in (4, 5):
            z.number_format = DAT
    for w in range(WOCHEN):
        c = 7 + w
        col = L(c)
        z = ga.cell(row=r, column=c, value=f"=IF(AND($F{r}>0,$D{r}<={col}${GH}+4,$E{r}>={col}${GH}),1,\"\")")
        z.number_format = ";;;"
        z.border = RAHMEN
last = L(6 + WOCHEN)
bereich = f"G{G0}:{last}{G0+len(TASKS)-1}"
for art, farbe in (("Eigen", "70AD47"), ("Fremd", "4472C4"), ("Kauf", "A6A6A6"), ("Verkauf", "A6A6A6"), ("Puffer", "ED7D31")):
    ga.conditional_formatting.add(bereich, FormulaRule(formula=[f'AND(G{G0}=1,$C{G0}="{art}")'],
                                                       fill=PatternFill("solid", fgColor=farbe, bgColor=farbe)))
ga.cell(row=G0 + len(TASKS) + 1, column=2, value="Projektdauer gesamt (Monate)")
ga.cell(row=G0 + len(TASKS) + 1, column=6, value=f"={MONATE}")
for c, w in zip("ABCDEF", [4, 46, 8, 9, 9, 5]):
    ga.column_dimensions[c].width = w
ga.freeze_panes = f"G{G0}"

# =====================================================================
# 1 Fix & Flip: Vorlage befüllen und mit Blatt 2/3 verknüpfen
# =====================================================================
ff["G2"] = "ETW Bohmte, 1. OG, Bj. 1964, 80 m² – Kleinanzeigen 3563162731"
ff["D6"] = 102000
ff["D7"] = f"={M['wfl']}"
ff["C8"], ff["C9"], ff["C10"] = 0.02, 0.0357, 0.05
ff["I6"] = f"='3 Renovierung'!C{S0+8}"
ff["G7"] = "Renovierung (aus Blatt 3)"
ff["I7"] = "='3 Renovierung'!C53"
ff["H8"] = f"='3 Renovierung'!{KP}"
ff["I9"] = f"='3 Renovierung'!C{S0+9}"
ff["C15"] = f"={MONATE}"
ff["I16"], ff["I17"] = 280, 60
ff["O5"], ff["O6"] = 1800, 2100
ff["O10"] = "=O7*0.0357"
ff["O10"].number_format = EUR
ff["B43"] = 7.5
ff["E19"] = "=D19/D7"
for r in range(32, 40):  # Vorlagen-Fehler: Verkaufsmakler fehlte beim Gewinn, Co-Investor bei HOCH auf O7
    ff[f"E{r}"] = f"=$O$7-D{r}-($O$11*$O$7)-$O$12-$O$10"
    ff[f"L{r}"] = f"=$O$8-D{r}-($O$11*$O$8)-$O$12-$O$10"
ff["S38"] = ("Verknüpft: Wohnfläche (Blatt 2), Entrümpelung, Renovierung, Küche, Puffer und Projektdauer (Blatt 3). "
             "Annahmen: Hausgeld 280 €/Monat (erfragen), Strom+Heizung 60 €, Verkauf renoviert 1.800–2.100 €/m², "
             "Verkauf mit Makler 3,57 %, Grunderwerbsteuer 5 %. Wohnung ist vermietet – Rechnung gilt erst, wenn sie leer ist.")

ff.merge_cells("S38:W48")
ff["S38"].alignment = Alignment(wrap_text=True, vertical="top")
ff["S38"].font = Font(size=9)
order = ["1 Fix & Flip", "2 Bewertung & Maße", "3 Renovierung", "4 Material", "5 Zeitplan (Gantt)", "DISCLAIMER"]
wb._sheets = [wb[n] for n in order]
wb.active = 0
for ws in wb.worksheets:
    ws.sheet_view.tabSelected = ws.title == order[0]
wb.save(OUT)
print("ok", OUT)
