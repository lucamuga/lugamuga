# Kapitalanlage – Vermietung

Häuser und Wohnungen kaufen, vermieten und langfristig halten (Buy & Hold).

## Ordner
| Ordner | Inhalt |
|---|---|
| `Objekte` | Ein Ordner pro geprüftem Objekt mit Steckbrief (Fazit, Chancen, Probleme) und eigenem Rechner |

## Geprüfte Objekte
| Objekt | Cashflow vorsichtig / günstig | Kurzfazit |
|---|---|---|
| `Damme-Glueckauf_Zweifamilienhaus_IS24-171741799` | −896 € / −48 € | zu teuer für Vermietung; max. ca. 207.500 € |
| `Damme-CAD_Mehrfamilienhaus_KA-3558261375` | −1.056 € / −150 € | Preis fehlt; erst Mietliste und Preis erfragen |
| `Damme-Neubau-DHH_allkauf_KA-3578260976` | −1.465 € / −94 € | Neubau-Ausbauhaus; Gesamtkosten weit über Anzeigepreis |

## Regel: Nur einrechnen, was nötig ist
Sanierung und Umbau kommen nur in den Rechner, wenn sie **Pflicht** sind (Gesetz, Sicherheit) oder du **ohne sie nicht vermieten kannst**. Beispiele:
- **Pflicht:** GEG-Nachrüstung nach dem Kauf (Rohrdämmung, oberste Geschossdecke, Kessel älter als 30 Jahre ohne Brennwert/Niedertemperatur), Rauchmelder, getrennte Zähler und Heizkostenerfassung bei mehreren Mietparteien, Küchenanschlüsse, Reparatur von Defekten und Sicherheitsmängeln.
- **Nicht einrechnen:** Einbauküche, Maler und Böden bei gepflegtem Zustand, Bad-Modernisierung, Fenster/Fassade/Heizungstausch ohne Pflicht.
- **Laufende Folgen trotzdem rechnen:** z. B. CO₂-Kosten-Anteil des Vermieters bei schlechter Energieklasse.
Jeder Steckbrief hat dazu einen Abschnitt „Sanierungs-Check“.

## So prüfst du ein neues Objekt
1. Ordner `Objekte/Ort_Objektart_Quelle-ID` anlegen.
2. `objekt.json` aus einem vorhandenen Objekt kopieren und die Werte anpassen (konservativ / Spanne / günstig).
3. Rechner erzeugen: `python3 _Werkzeuge/objekt_rechner.py objekt.json roh.xlsx`, danach `python3 _Werkzeuge/excel_werte_speichern.py roh.xlsx Rechner_….xlsx`.
4. Steckbrief schreiben.
