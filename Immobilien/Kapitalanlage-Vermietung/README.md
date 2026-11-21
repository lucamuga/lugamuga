# Kapitalanlage – Vermietung

Häuser und Wohnungen kaufen, vermieten und langfristig halten (Buy & Hold).

## Ordner
| Ordner | Inhalt |
|---|---|
| `Objekte` | Ein Ordner pro geprüftem Objekt mit Steckbrief (Fazit, Chancen, Probleme) und eigenem Rechner |

## So prüfst du ein neues Objekt
1. Ordner `Objekte/Ort_Objektart_Quelle-ID` anlegen.
2. `objekt.json` aus einem vorhandenen Objekt kopieren und die Werte anpassen (konservativ / Spanne / günstig).
3. Rechner erzeugen: `python3 _Werkzeuge/objekt_rechner.py objekt.json roh.xlsx`, danach `python3 _Werkzeuge/excel_werte_speichern.py roh.xlsx Rechner_….xlsx`.
4. Steckbrief schreiben.
