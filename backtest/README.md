# Nasdaq 10:00-ET-Reversal: Backtest

Mit diesem Code lässt sich prüfen, ob das oft zitierte **10-Uhr-Reversal** im Nasdaq (NQ/MNQ-Futures oder QQQ) einen
messbaren Vorteil hat. Die Uhrzeiten beziehen sich auf New York (ET), in Deutschland ist das meist 16:00 MEZ/MESZ.
Rund um die Zeitumstellung (März/Oktober bzw. November) liegt es kurzzeitig bei 15:00.

> Kein Anlageratschlag. Ein Backtest zeigt nur, was in der Vergangenheit unter vereinfachten Annahmen passiert wäre.

## Die Idee

1. **Opening Drive (09:30–09:59 ET):** Nach dem Cash-Open läuft der Markt oft stark in eine Richtung. Dahinter stehen
   Overnight-Orders, Gap-Fills und Stops.
2. **10:00 ET:** Um diese Zeit kommen viele Makrodaten (ISM, JOLTS, Konsumentenvertrauen, Neubauverkäufe). An manchen Tagen
   sprechen auch Fed-Vertreter, daher der Spitzname „Powell 10 AM“. Der Opening Drive erschöpft sich dann häufig und der
   Markt dreht.
3. **Trade:** Gehandelt wird gegen den Opening Drive. Der Stop liegt hinter dem Hoch bzw. Tief des Drives, das Ziel ist ein
   Vielfaches des Risikos oder ein 50-%-Retracement. Spätestens um 11:30 ET wird glattgestellt.

### Regeln im Code (`strategy.py`)

| Parameter | Standard | Bedeutung |
|---|---|---|
| `--min-drive` | 0.20 | Mindestgröße des Opening Drive in % (Open 09:30 → Close 09:59) |
| `--mode` | `break` | `break`: Einstieg, wenn eine Bar nach 10:05 über/unter der 10:00–10:05-Range schließt, Fill zum Open der nächsten Bar. `fade`: Einstieg sofort zum 10:00-Open |
| `--setup-end` | 10:15 | Nach dieser Zeit gibt es keine neuen Einstiege mehr |
| Stop | – | Tageshoch (Short) bzw. Tagestief (Long) bis zum Einstieg, plus `--stop-buffer` % |
| `--target` / `--target-r` | `r` / 2.0 | Ziel bei 2R oder (`retrace`) beim 50-%-Retracement des Drives |
| `--exit-time` | 11:30 | Zeit-Exit |

Fills sind konservativ angesetzt: Liegen Stop und Ziel in derselben Bar, zählt der Stop. Ein Gap über den Stop wird zum Open
gefüllt. Slippage (Standard 1 Tick pro Seite) und Kommission (Standard $4 Round Trip) werden immer abgezogen.

## Daten

Yahoo Finance liefert nur wenige Wochen Minutendaten. Für belastbare Ergebnisse braucht man **mindestens 2–3 Jahre
1-Minuten-Daten**, z. B. aus einem TradingView-Export, von FirstRate Data, Databento oder Kibot.

CSV-Format (Spaltennamen sind case-insensitiv, `volume` ist optional):

```
datetime,open,high,low,close,volume
2024-03-05 09:30:00,18100.25,18112.00,18098.50,18110.75,5321
```

- Zeitstempel ohne Zeitzone werden als New-York-Zeit gelesen. Andernfalls `--source-tz UTC` (o. ä.) angeben.
  Unix-Timestamps werden automatisch erkannt.
- Overnight-Bars werden verworfen, es bleiben nur 09:30–16:00 ET übrig.
- Standardmäßig läuft die Strategie auf **5-Minuten-Kerzen** (`--timeframe 5m`). 1m-Daten werden automatisch
  zu 5m zusammengefasst. Mit `--timeframe 1m` oder `15m` lässt sich das ändern. Auf 5m ist die 10:00-Range genau die
  10:00-Kerze, der Einstieg erfolgt nach einem Schlusskurs über/unter dieser Kerze (bis zur 10:10-Kerze).

**News-Tage (optional):** Eine CSV, deren erste Spalte die Daten mit einem 10:00-ET-Release enthält, z. B. aus dem
Wirtschaftskalender von Investing.com oder ForexFactory. Mit `--news-dates news.csv` werden die Ergebnisse nach News-Tag
aufgeschlüsselt, mit `--news-only` wird nur an diesen Tagen gehandelt.

## Nutzung

```bash
pip install -r backtest/requirements.txt

# Funktionstest mit erfundenen Daten (Ergebnisse sind bedeutungslos!)
python backtest/run_backtest.py --synthetic

# Echte NQ-Daten (Zeitstempel in UTC)
python backtest/run_backtest.py --csv data/nq_1m.csv --source-tz UTC

# Nur der letzte Monat (30 Kalendertage bis zum letzten Tag in den Daten) oder ein fester Zeitraum
python backtest/run_backtest.py --csv data/nq_1m.csv --last-days 30
python backtest/run_backtest.py --csv data/nq_1m.csv --from 2026-09-01 --to 2026-09-30

# MNQ, Fade-Variante, 50-%-Ziel, nur an News-Tagen
python backtest/run_backtest.py --csv data/nq_1m.csv --tick-value 0.5 --commission 1.5 \
    --mode fade --target retrace --news-dates data/news.csv --news-only

# QQQ via yfinance (nur auf dem eigenen Rechner, 5m-Daten der letzten ~60 Tage), letzter Monat, 100 Aktien pro Trade
python backtest/run_backtest.py --yfinance QQQ --last-days 30 --tick-size 0.01 --tick-value 1 --commission 0

# Tests
python -m pytest backtest/tests
```

Ausgabe:

- Konsole: Zusammenfassung (Trefferquote, Ø R, Expectancy, Profit Factor, Max Drawdown, Sharpe) und Aufschlüsselung nach
  Monat, Wochentag, Richtung und News-Tag.
- `results/report.md`: Monatsübersicht mit Kennzahlen und allen Trades als Tabelle.
- `results/charts/JJJJ-MM-TT.png`: **5-Minuten-Kerzenchart je Trade** (09:30–12:00) mit Opening Drive, 10:00-Range,
  Einstieg, Stop, Ziel und Exit. Abschalten mit `--no-charts`.
- `results/equity.png`: kumulierter P&L mit den Tagesergebnissen als Balken.
- `results/trades.csv`: alle Trades als Rohdaten.

### Zufalls-Baseline

Für jeden Trade wird dieselbe Einstiegszeit mit demselben Risiko-Ziel-Verhältnis gehandelt, nur mit **zufälliger
Richtung**, und das 200-mal. Wenn die Strategie nicht klar besser abschneidet als die meisten dieser Läufe (`runs >= strategy`
sollte unter ca. 5 % liegen), kommt der „Vorteil“ eher vom Markttrend oder vom Zufall als vom 10-Uhr-Muster.

## Fallstricke

- **Overfitting:** Je mehr Parameter man optimiert, desto schöner sieht der Backtest aus und desto schlechter läuft es live.
  Am besten auf einem Zeitraum optimieren (z. B. 2021–2023) und auf einem anderen prüfen (2024–2025).
- **Regimewechsel:** Das Muster hängt von der Volatilität und vom Zinsumfeld ab. Die Ergebnisse sollten pro Jahr
  verglichen werden.
- **Kontraktrolle:** Bei Futures-Daten sollte man einen back-adjusted Continuous Contract oder den jeweils aktiven Kontrakt
  verwenden, sonst entstehen künstliche Gaps.
- **FOMC-Tage:** Die Fed-Entscheidung kommt um 14:00 ET, also nach dem Trade-Fenster. Der Vormittag ist an diesen Tagen aber
  oft untypisch ruhig. Ein Test mit und ohne diese Tage lohnt sich.
