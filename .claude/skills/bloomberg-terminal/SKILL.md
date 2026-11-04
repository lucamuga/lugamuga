---
name: bloomberg-terminal
description: Baut Finanz-Dashboards und Marktdaten-Oberflächen im Stil eines Bloomberg-artigen Trading-Terminals (dicht, monospace, dunkel, Orange-Akzent, Kürzel-Buttons, Live-Ticks, Sparklines, Tastenkürzel). Verwenden bei "Bloomberg Terminal", "Trading-Terminal", "Börsen-Dashboard", "Marktübersicht", "Ticker-Tabelle", "Watchlist", "Finanz-UI im Terminal-Look" oder wenn eine React/Next.js/HTML-Oberfläche wie ein professionelles Markt-Terminal aussehen soll.
---

# Bloomberg-Terminal-Stil

Leitfaden für Oberflächen, die sich wie ein professionelles Markt-Terminal anfühlen.
Referenz-Implementierung (Next.js 15, React 19, Tailwind, shadcn/ui, Jotai, React Query):
https://github.com/feremabraz/bloomberg-terminal – nur als Inspiration lesen; das Repo
hat **keine Lizenz**, also keinen Code daraus 1:1 übernehmen.

> Nur den *Stil* nachbauen. Kein Bloomberg-Logo, keinen Markennamen als Produkttitel und
> nichts, was als offizielles Bloomberg-Produkt durchgehen könnte. Eigenen Namen wählen
> (z. B. "MKT TERMINAL").

## 1. Designprinzipien

- **Dichte vor Weißraum**: `text-xs` (11–12px), Zellen `px-2 py-1`, Zeilenhöhe ~20px.
- **Monospace überall**: `font-family: "Courier New", ui-monospace, monospace;`
  Zahlen mit `font-variant-numeric: tabular-nums` und rechtsbündig.
- **Eckig**: `rounded-none`, 1px-Ränder, keine Schatten, keine Verläufe.
- **Farbe trägt Bedeutung**, nie Dekoration (siehe Tokens).
- **Tastatur zuerst**: jede Ansicht/Aktion hat ein Kürzel; Hilfe-Overlay mit `?`.
- **Daten sind lebendig**: Ticks blinken kurz auf, Frische-Indikator zeigt das Alter.

## 2. Farb-Tokens

| Token            | Dark      | Light     | Verwendung                              |
|------------------|-----------|-----------|-----------------------------------------|
| `--bg`           | `#121212` | `#f0f0f0` | Seitenhintergrund                       |
| `--surface`      | `#1e1e1e` | `#e0e0e0` | Kopfzeile, Filterleiste, Tabellenkopf   |
| `--header`       | `#000000` | `#d0d0d0` | oberste Leiste                          |
| `--text`         | `#ffffff` | `#000000` | Standardtext                            |
| `--text-2`       | `#888888` | `#555555` | Zeilennummern, Labels, Zeitstempel      |
| `--accent`       | `#ff9900` | `#ff9900` | Ticker-Kürzel, Section-Titel, Links     |
| `--value`        | `#f5f5b8` | `#8b7500` | Kurs/Wert-Spalte                        |
| `--pos`          | `#4caf50` | `#2e7d32` | positive Veränderung                    |
| `--neg`          | `#f44336` | `#c62828` | negative Veränderung                    |
| `--border`       | `#333333` | `#cccccc` | Zeilen- und Panel-Trenner               |
| `--flash`        | `#713f12` | `#fde047` | Hintergrund-Blitz bei Tick (≈500ms)     |

Als CSS-Variablen auf `:root` definieren, Dark/Light über `[data-theme]` oder Tailwind
`dark:` umschalten. In Tailwind **keine** dynamischen Klassen wie `` `text-[${color}]` ``
bauen – der JIT-Compiler findet sie nicht. Stattdessen CSS-Variablen
(`text-[var(--accent)]`) oder `style={{ color }}` verwenden.

## 3. Layout-Bausteine

Von oben nach unten, volle Breite, kein Seiten-Padding:

1. **Funktionsleiste** (`--surface`): Kürzel-Buttons in Großbuchstaben, 4–5 Zeichen.
   - Rot = destruktiv/abbrechen (`CANCL`, `STOP`)
   - Grün = Aktion/Ansicht (`NEW`, `NEWS`, `GMOV`, `GVOL`, `LIVE`)
   - Amber mit schwarzem Text = Hilfsfunktionen (`HELP`, `REFR`, `DARK`)
   - Button: `h-6 px-2 text-xs font-mono rounded-none border border-neutral-700`,
     optional 12px-Icon (lucide) links.
   - Rechts (`ml-auto`): Datenquelle, Frische-Punkt, letzte Aktualisierung.
2. **Filterleiste**: Watchlist-Dropdown (fett, Akzentfarbe) + kleine eckige Checkboxen
   (`h-3 w-3 rounded-none`) für Spalten/Filter (Movers, Volatility, YTD, Währung …).
3. **Hauptbereich**: eine dichte Tabelle oder geteilte Panels (`react-resizable-panels`).
4. **Statuszeile** (optional): Tastenkürzel-Hinweise, Marktstatus, Uhrzeit.

## 4. Markttabelle

- Erste Spalte `sticky left-0` mit Hintergrund, fixe Breite (120–140px), enthält
  Zeilennummer (`--text-2`) + Kürzel (`--accent`).
- Gruppierung in **nummerierte Sections**: `1) Americas`, `2) EMEA`, `3) Asia/Pacific`
  als volle Zeile in Akzentfarbe, fett.
- Typische Spalten: `Market | RMI | 2Day | Value | Net Chg | %Chg | Δ AVAT | Time | %Ytd | %YtdCur`
- Vorzeichen immer ausschreiben (`+1.23`, `-0.45%`), 2 Nachkommastellen.
  Fehlende Werte als `—` statt `N/A`.
- Weniger wichtige Spalten responsiv ausblenden: `hidden sm:table-cell`, `hidden md:table-cell`.
- Tabelle mit `border-separate border-spacing-0`, damit der Sticky-Rand sauber bleibt.

## 5. Live-Daten & Tick-Highlight

- Server-State über **React Query** (`refetchInterval` für den LIVE-Modus,
  `staleTime` passend zur Quelle), UI-State (Ansicht, Filter, Theme) über **Jotai** oder
  `useState`/Context.
- Beim Update alten und neuen Wert je Zelle vergleichen; geänderte Zell-Keys
  (`${region}-${id}-${field}`) für ~500ms in einem Set halten und der Zelle dann
  `bg-[var(--flash)] transition-colors duration-500` geben.
- Frische-Indikator: Punkt `h-2 w-2 rounded-full`
  grün + `animate-pulse` (< 10s), grün (< 30s), gelb + pulse (< 60s), rot (älter).
- Ohne API-Schlüssel immer eine **Simulation** anbieten (Random Walk ±0.1–0.5 % pro Tick),
  damit die UI sofort lebendig ist. Echte Quellen: Alpha Vantage, Finnhub, Polygon usw.;
  Schlüssel nur serverseitig (API-Route), nie im Client.

## 6. Sparklines

- Canvas oder Inline-SVG, ca. **80×20px**, Linienbreite 1, keine Achsen.
- Zweigeteilt: linke Hälfte = Vortag in `#666`, rechte Hälfte = heute in `--pos`/`--neg`
  je nach Tagesrichtung.
- Bei Canvas `devicePixelRatio` berücksichtigen (Canvas-Größe × dpr, `ctx.scale(dpr, dpr)`).

## 7. Tastenkürzel

Ein globaler `keydown`-Listener mit Tabelle `{ key, ctrl?, alt?, shift?, action, description }`;
Treffer → `preventDefault()`. Nicht auslösen, wenn der Fokus in einem Eingabefeld liegt.

| Kürzel   | Aktion                     |
|----------|----------------------------|
| `1`–`4`  | Ansicht wechseln (Markt, News, Movers, Volatilität) |
| `Ctrl+N` | neue Watchlist             |
| `Ctrl+B` | leeren (mit Bestätigung)   |
| `Ctrl+R` | Daten aktualisieren        |
| `Ctrl+L` | Live-Updates an/aus        |
| `Esc`    | abbrechen / Dialog schließen |
| `?`      | Hilfe-Overlay              |

Das Hilfe-Overlay ist ein eckiger Dialog (`--surface`, 2px-Rand) mit einer zweispaltigen
Tabelle "Shortcut | Description".

## 8. Typische Ansichten

- **Market**: Regionen-Tabelle (oben).
- **News**: Liste mit Zeit (`--text-2`), Quelle (`--accent`), Headline; neueste oben.
- **Movers**: Top-Gewinner/-Verlierer nach `%Chg`, zwei Spalten nebeneinander.
- **Volatility**: Tabelle/Heatmap nach Volatilitätskennzahl.
- **Analyse (optional)**: KI-Zusammenfassung des Marktes über eine Server-Route mit
  Rate-Limit; Ergebnis als Monospace-Text in einem Panel.

## 9. Vorgehen

1. Stack klären: Für ein bestehendes React/Next.js-Projekt Tokens + Komponenten dort
   einbauen; für einen schnellen Prototyp die Vorlage
   `references/terminal-template.html` (eine Datei, ohne Build) kopieren und anpassen.
2. Tokens (Abschnitt 2) zuerst anlegen, dann Funktionsleiste → Filterleiste → Tabelle.
3. Simulation einbauen, bevor eine echte Datenquelle angebunden wird.
4. Tastenkürzel + Hilfe-Overlay ergänzen.
5. Prüfen: Dark und Light, schmale Breite (Sticky-Spalte + ausgeblendete Spalten),
   Tick-Blitz sichtbar, keine horizontale Seiten-Scrollbar außerhalb der Tabelle.
