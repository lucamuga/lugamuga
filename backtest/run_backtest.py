"""Command-line backtest of the Nasdaq 10:00 ET reversal.

Examples:
    python backtest/run_backtest.py --csv data/nq_1m.csv --source-tz UTC
    python backtest/run_backtest.py --synthetic --mode fade --target retrace
    python backtest/run_backtest.py --yfinance QQQ --interval 5m --tick-size 0.01 --tick-value 0.01
"""

from __future__ import annotations

import argparse
import os
import sys
from datetime import date, datetime, time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pandas as pd  # noqa: E402

from backtest import data_loader, engine, metrics, synthetic  # noqa: E402
from backtest.strategy import Params  # noqa: E402


def _hhmm(s: str) -> time:
    return datetime.strptime(s, "%H:%M").time()


def _load_news_dates(path: str | None) -> frozenset[date]:
    if not path:
        return frozenset()
    col = pd.read_csv(path).iloc[:, 0]
    return frozenset(pd.to_datetime(col).dt.date)


def parse_args(argv=None) -> argparse.Namespace:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    src = ap.add_mutually_exclusive_group(required=True)
    src.add_argument("--csv", help="CSV with datetime, open, high, low, close[, volume]")
    src.add_argument("--yfinance", metavar="SYMBOL", help="download via yfinance (short history only)")
    src.add_argument("--synthetic", action="store_true", help="made-up demo data")
    ap.add_argument("--source-tz", default=data_loader.NY_TZ, help="timezone of naive CSV timestamps")
    ap.add_argument("--interval", default="1m", help="yfinance interval")
    ap.add_argument("--resample", help="pandas rule, e.g. 5min")

    ap.add_argument("--mode", choices=["break", "fade"], default="break")
    ap.add_argument("--min-drive", type=float, default=0.20, help="min opening drive in percent")
    ap.add_argument("--range-end", type=_hhmm, default=time(10, 5))
    ap.add_argument("--setup-end", type=_hhmm, default=time(10, 15))
    ap.add_argument("--exit-time", type=_hhmm, default=time(11, 30))
    ap.add_argument("--stop-buffer", type=float, default=0.0, help="percent beyond the drive extreme")
    ap.add_argument("--target", choices=["r", "retrace"], default="r")
    ap.add_argument("--target-r", type=float, default=2.0)

    ap.add_argument("--tick-size", type=float, default=0.25)
    ap.add_argument("--tick-value", type=float, default=5.0, help="NQ 5.0, MNQ 0.5")
    ap.add_argument("--slippage-ticks", type=float, default=1.0)
    ap.add_argument("--commission", type=float, default=4.0, help="$ per round trip")
    ap.add_argument("--contracts", type=int, default=1)

    ap.add_argument("--news-dates", help="CSV whose first column lists dates with a 10:00 ET release")
    ap.add_argument("--news-only", action="store_true")
    ap.add_argument("--weekdays", default="0,1,2,3,4", help="Monday=0")
    ap.add_argument("--max-gap", type=float, help="skip days gapping more than this percent")

    ap.add_argument("--baseline-runs", type=int, default=200)
    ap.add_argument("--out", default="results")
    return ap.parse_args(argv)


def main(argv=None) -> int:
    a = parse_args(argv)
    if a.csv:
        bars = data_loader.load_csv(a.csv, source_tz=a.source_tz, resample=a.resample)
    elif a.yfinance:
        bars = data_loader.load_yfinance(a.yfinance, interval=a.interval)
    else:
        bars = synthetic.make_dataset()
        print("WARNING: synthetic data - results demonstrate the code, not the market.\n")
    if a.resample and not a.csv:
        bars = data_loader.normalize(bars, resample=a.resample)

    params = Params(
        mode=a.mode,
        min_drive_pct=a.min_drive,
        range_end=a.range_end,
        setup_end=a.setup_end,
        stop_buffer_pct=a.stop_buffer,
        target=a.target,
        target_r=a.target_r,
        exit_time=a.exit_time,
    )
    costs = engine.Costs(a.tick_size, a.tick_value, a.slippage_ticks, a.commission, a.contracts)
    filters = engine.Filters(
        weekdays=tuple(int(x) for x in a.weekdays.split(",")),
        news_dates=_load_news_dates(a.news_dates),
        news_only=a.news_only,
        max_gap_pct=a.max_gap,
    )

    n_days = len(set(bars.index.date))
    print(f"Data: {n_days} sessions, {bars.index[0]} -> {bars.index[-1]}")
    print(f"Params: {params}\n")

    trades = engine.run(bars, params, costs, filters)
    stats = metrics.summary(trades, trading_days=n_days)
    print("== Summary ==")
    for k, v in stats.items():
        print(f"  {k:<18} {v}")
    if trades.empty:
        return 0

    base = engine.random_baseline(bars, trades, params, costs, runs=a.baseline_runs)
    if base.size:
        pct = float((base >= stats["expectancy_usd"]).mean() * 100)
        print(
            f"\n== Random-direction baseline ({base.size} runs) ==\n"
            f"  mean expectancy  {base.mean():.2f} $/trade (5-95%: {pd.Series(base).quantile(.05):.2f} .. "
            f"{pd.Series(base).quantile(.95):.2f})\n"
            f"  runs >= strategy {pct:.1f}%  (low = edge unlikely to be luck)"
        )

    for name, table in metrics.breakdown(trades).items():
        print(f"\n== By {name} ==\n{table.to_string()}")

    os.makedirs(a.out, exist_ok=True)
    trades_path = os.path.join(a.out, "trades.csv")
    trades.to_csv(trades_path, index=False)
    print(f"\nTrades written to {trades_path}")
    try:
        import matplotlib

        matplotlib.use("Agg")
        import matplotlib.pyplot as plt

        fig, ax = plt.subplots(figsize=(10, 4))
        ax.plot(pd.to_datetime(trades["date"]), trades["pnl_usd"].cumsum())
        ax.set_title("10:00 ET reversal - cumulative P&L ($)")
        ax.grid(alpha=0.3)
        fig.tight_layout()
        chart = os.path.join(a.out, "equity.png")
        fig.savefig(chart, dpi=120)
        print(f"Equity curve written to {chart}")
    except ImportError:
        print("matplotlib not installed - skipping equity chart")
    return 0


if __name__ == "__main__":
    sys.exit(main())
