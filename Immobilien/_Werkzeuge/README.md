# Werkzeuge für die Immobilien-Dateien

Erzeugt bzw. prüft die Dateien in `Fix-und-Flip/` und hilft bei allen Excel-Dateien.

| Datei | Zweck | Aufruf |
|---|---|---|
| `flip_rechner.py` | Baut den Fix-&-Flip-Rechner (Excel, nur Formeln) | `python3 flip_rechner.py /pfad/flip_raw.xlsx`, danach mit `excel_werte_speichern.py` nach `../Fix-und-Flip/02_Rechner/` schreiben |
| `excel_werte_speichern.py` | Rechnet eine Excel-Datei mit LibreOffice durch und speichert die Werte in der Datei, damit sie überall (Handy, Schutzansicht) angezeigt werden | `python3 excel_werte_speichern.py EINGABE.xlsx AUSGABE.xlsx` |
| `flip_pruefung.py` | Unabhängige Nachrechnung der Flip-Formeln zum Abgleich | `python3 flip_pruefung.py` |
| `fix_flip_plan.js` | Erzeugt den Strategieplan (Word) aus den Ergebnissen des Rechners | `NODE_PATH=$(npm root -g) node fix_flip_plan.js deals.json ../Fix-und-Flip/01_Strategieplan/Fix_Flip_Plan_Damme.docx` |

Benötigt: Python 3 mit `openpyxl` und `lxml`, LibreOffice (`soffice`), Node.js mit dem Paket `docx`.
| `doc_helfer.js` | Gemeinsame Bausteine (Überschriften, Tabellen, Kästen) für die Word-Dokumente | wird von den `.js`-Skripten geladen |
| `sanierungs_handbuch.js` | Erzeugt das Sanierungs-Handbuch (Word) | `NODE_PATH=$(npm root -g) node sanierungs_handbuch.js ../Fix-und-Flip/04_Sanierungs-Handbuch/Sanierungs-Handbuch_Fix_Flip.docx` |
| `werkzeugliste.py` | Baut die Werkzeug- und Materialliste (Excel) | `python3 werkzeugliste.py /pfad/wz_raw.xlsx`, danach `excel_werte_speichern.py` |
| `objekt_rechner.py` | Objekt-Rechner für Vermietung mit bis zu 2 Wohnungen (Spalten: konservativ, Spanne, günstig) aus einer `objekt.json` | `python3 objekt_rechner.py objekt.json roh.xlsx`, danach `excel_werte_speichern.py` |
