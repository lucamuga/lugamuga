# Fix & Flip – Deals

Jedes Objekt hat einen eigenen Ordner mit Steckbrief. Neue Objekte: Ordner `Ort_Objektart` anlegen, Steckbrief kopieren, Zahlen im Rechner eintragen.

| Objekt | Angebot | Gewinn vor Steuern | Max. Kaufpreis | Status |
|---|---|---|---|---|
| [Einfamilienhaus Lohne](Lohne_Einfamilienhaus_1946/Steckbrief.md) | 195.000 € | 50.664 € | 174.505 € | Grenzwertig – nachverhandeln |
| [Bungalow Damme-West](Damme-West_Bungalow/Steckbrief.md) | 250.000 € | −21.203 € | 181.173 € | Verlust zu diesem Preis |
| [Doppelhaushälfte Damme](Damme-Glueckauf_Doppelhaushaelfte/Steckbrief.md) | 235.000 € | −43.319 € | 157.190 € | Verlust zu diesem Preis |
| [Haus Lemförde](Lemfoerde_Bahnhofstrasse_76/Steckbrief.md) | 236.000 € | −109.788 € | 94.668 € | Verlust zu diesem Preis |

**Verworfen:** Wohnung Damme-Ort (83 m², Baujahr 1993, 190.000 €) – gepflegt, kein Sanierungsfall, Ergebnis −27.625 €.

## Beobachtungsliste Zwangsversteigerungen
- Steinfeld, Handorfer Straße 81 – Verkehrswert 370.000 €, AG Vechta; Termin 24.04.2026 vorbei, Status erfragen
- Dinklage – Verkehrswert 495.000 €, laut Portal Termin 27.11.2026 (unbestätigt)
- AG Diepholz (Hüde/Lemförde) – wöchentlich auf zvg-portal.de prüfen

## Neue Kandidaten aus Leverage (Suche 25.11.2026, 25 km um Damme, bis 300.000 €)
338 Treffer. Herausgefiltert: Neubau-/Bausatz-Anzeigen, Baudenkmäler/Fachwerk, Ausland, Rohbau. Gerechnet mit `_Werkzeuge/flip_pruefung.py`, vorsichtig / günstig (Verkaufspreise nach Sanierung sind Schätzungen).

| Objekt | Preis | Gewinn vorsichtig / günstig | Max. Kaufpreis (15 % Marge) vorsichtig / günstig | In Leverage |
|---|---|---|---|---|
| [DHH Lohne, 1932, 148 m²](https://www.immobilienscout24.de/expose/172517721) | 169.000 € | −16.500 € / +98.700 € | 109.500 € / 204.600 € | ja |
| [EFH Ostercappeln, 1955, 155 m²](https://www.kleinanzeigen.de/s-anzeige/viel-platz-bei-bad-essen-immobilie-mit-potenzial-fuer-handwerker-und-kreative-koepfe/3453692508-208-2632) | 179.000 € | −35.600 € / +59.800 € | 106.500 € / 184.800 € | ja |
| [EFH Stemwede, 1900, 140 m²](https://www.immobilienscout24.de/expose/172795623) | 51.750 € | −30.800 € / +70.000 € | – / 79.300 € | nein (Monatslimit) |
| EFH Diepholz-Aschen, 1965, 124 m² | 159.000 € | −50.800 € / +25.000 € | 82.700 € / 144.300 € | nein |
| DHH Stemwede, 1961, 140 m² | 169.000 € | −58.500 € / +26.600 € | 83.700 € / 152.800 € | nein |
| REH Quakenbrück, 1936, 106 m² | 127.000 € | −63.800 € / +8.500 € | 45.400 € / 104.100 € | nein |

Hinweise:
- **Lohne:** rechtlich Wohnungseigentum (524/1.000 Miteigentum, Teilungserklärung prüfen), Gasheizung neuwertig, 2 Bäder, Doppelgarage; Käuferprovision 3,57 %.
- **Ostercappeln:** Energieklasse C, Heizung 2019 → Sanierung eher am unteren Ende; Käuferprovision 3,57 %; 139 Tage online.
- In Leverage ist die Maklerprovision bei beiden auf 0 % gesetzt – in der App auf 3,57 % korrigieren.

### Nachgerechnet mit deinem FixundFlipRechner (Datei `02_Rechner/FixundFlip_Rechner_Kandidaten_2026-11.xlsx`, ein Blatt je Objekt)
Annahmen: Grunderwerbsteuer 5 %, Notar 2 %, Makler 3,57 %, Zins 5,5 % auf 100 % Kaufpreis, Puffer 15 %, Verkauf mit Makler 3,57 %.

| Objekt | Gewinn GERING / HOCH | „Goldene Mitte“ | Kaufpreis für ca. 20 % Marge |
|---|---|---|---|
| DHH Lohne | −19.700 € / +39.500 € | +9.900 € (2,7 %) | ca. 123.000 € (−27 %) |
| EFH Ostercappeln | −38.700 € / +7.800 € | −15.500 € | ca. 117.000 € (−35 %) |
| EFH Stemwede 1900 | −30.100 € / +11.900 € | −9.100 € | nicht erreichbar |
| EFH Diepholz-Aschen | −54.300 € / −17.100 € | −35.700 € | ca. 88.000 € (−45 %) |
| DHH Stemwede 1961 | −62.000 € / −20.000 € | −41.000 € | ca. 93.000 € (−45 %) |
| REH Quakenbrück | −66.300 € / −34.500 € | −50.400 € | nicht erreichbar |
