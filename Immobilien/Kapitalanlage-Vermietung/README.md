# Kapitalanlage – Vermietung

Häuser und Wohnungen kaufen, vermieten und langfristig halten (Buy & Hold).

## Ordner
| Ordner | Inhalt |
|---|---|
| `Objekte` | Ein Ordner pro geprüftem Objekt mit Steckbrief (Fazit, Chancen, Probleme) und eigenem Rechner |

## Geprüfte Objekte
| Objekt | Cashflow vorsichtig / günstig | Kurzfazit |
|---|---|---|
| `Damme-Glueckauf_Zweifamilienhaus_IS24-171741799` | −1.039 € / −93 € | zu teuer für Vermietung; max. ca. 178.000 € |
| `Damme-CAD_Mehrfamilienhaus_KA-3558261375` | −1.096 € / −150 € | Preis fehlt; erst Mietliste und Preis erfragen |
| `Damme-Neubau-DHH_allkauf_KA-3578260976` | −1.589 € / −117 € | Neubau-Ausbauhaus; Gesamtkosten weit über Anzeigepreis |

## So prüfst du ein neues Objekt
1. Ordner `Objekte/Ort_Objektart_Quelle-ID` anlegen.
2. `objekt.json` aus einem vorhandenen Objekt kopieren und die Werte anpassen (konservativ / Spanne / günstig).
3. Rechner erzeugen: `python3 _Werkzeuge/objekt_rechner.py objekt.json roh.xlsx`, danach `python3 _Werkzeuge/excel_werte_speichern.py roh.xlsx Rechner_….xlsx`.
4. Steckbrief schreiben.
