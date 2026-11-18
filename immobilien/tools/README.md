# Werkzeuge für die Immobilien-Dateien

| Datei | Zweck | Aufruf |
|---|---|---|
| `flip_rechner.py` | Baut den Fix-&-Flip-Rechner (Excel, nur Formeln) | `python3 flip_rechner.py /pfad/flip_raw.xlsx` |
| `excel_werte_speichern.py` | Rechnet eine Excel-Datei mit LibreOffice durch und speichert die Werte in der Datei, damit sie überall (Handy, Schutzansicht) angezeigt werden | `python3 excel_werte_speichern.py EINGABE.xlsx AUSGABE.xlsx` |
| `flip_pruefung.py` | Unabhängige Nachrechnung der Flip-Formeln zum Abgleich | `python3 flip_pruefung.py` |
| `fix_flip_plan.js` | Erzeugt den Strategieplan (Word) aus den Ergebnissen des Rechners | `NODE_PATH=$(npm root -g) node fix_flip_plan.js deals.json Fix_Flip_Plan_Damme.docx` |

Benötigt: Python 3 mit `openpyxl` und `lxml`, LibreOffice (`soffice`), Node.js mit dem Paket `docx`.
