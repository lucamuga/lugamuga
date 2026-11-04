# MKT Terminal

Markt-Dashboard im Trading-Terminal-Stil, gebaut mit dem Skill `.claude/skills/bloomberg-terminal`.
Eine einzige HTML-Datei, ohne Build: `index.html` im Browser öffnen.

**Alle Kurse und Meldungen sind simuliert** (Random Walk ab gerundeten Beispielwerten).

## Funktionen

- **MKT (1)**: 19 Weltindizes nach Region, 2-Tages-Sparkline, Tick-Highlight, High/Low, %Ytd
- **NEWS (2)**: aus den simulierten Bewegungen erzeugte Meldungen
- **GMOV (3)**: Top-Gewinner und -Verlierer
- **GVOL (4)**: Heatmap der Tagesspanne
- **Watchlist**: ☆ in der Tabelle klicken, mit `W` umschalten; wird im Browser (localStorage) gespeichert
- `Ctrl+L` Live an/aus, `Ctrl+R` einzelner Tick, `T` Theme, `?` Hilfe

## Echte Kurse anbinden

Kurs-APIs (Alpha Vantage, Finnhub, Polygon …) brauchen einen API-Schlüssel, der nicht im
Browser landen darf. Dafür eine kleine Server-Route bauen, die die API abfragt, und in
`tick()` statt des Random Walks diese Route per `fetch` abfragen.
