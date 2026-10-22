"""Walk-forward search for a robust variant of the 10 AM reversal.

Every parameter combination in GRID is backtested once over the whole history. Then, fold by fold,
the best combination on the previous `--train-months` is picked and traded on the following
`--test-months` it has never seen. Only those out-of-sample trades count for the verdict.

    python backtest/optimize.py --csv data/nq_1m.csv --source-tz UTC
    python backtest/optimize.py --synthetic            # smoke test only
    python backtest/optimize.py --random-walk          # control: must NOT pass
"""

from __future__ import annotations

import argparse
import itertools
import os
import sys
from concurrent.futures import ProcessPoolExecutor
from dataclasses import asdict
from datetime import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402

from backtest import data_loader, engine, metrics, synthetic  # noqa: E402
from backtest.strategy import Params  # noqa: E402

GRID = {
    "mode": ["break", "fade"],
    "min_drive_pct": [0.15, 0.30, 0.50],
    "target": [("r", 1.5), ("r", 2.0), ("r", 3.0), ("retrace", 0.0)],
    "vwap_confirm": [False, True],
    "breakeven_r": [0.0, 1.0],
    "exit_time": [time(11, 30), time(12, 30)],
    "drive_adr_min": [0.0, 0.25],
}

# A variant only passes if its out-of-sample record meets all of these.
CRITERIA = {
    "min_trades": 100,
    "min_profit_factor": 1.2,
    "min_avg_r": 0.05,
    "min_positive_folds_pct": 60.0,
    "max_baseline_pct": 5.0,  # share of random-direction runs that do at least as well
}


def combos() -> list[Params]:
    keys = list(GRID)
    out = []
    for values in itertools.product(*GRID.values()):
        kw = dict(zip(keys, values))
        target, target_r = kw.pop("target")
        if kw["mode"] == "fade" and kw["vwap_confirm"]:
            continue  # at 10:00 the drive is still on, VWAP lags behind it: the filter only removes trades
        out.append(Params(target=target, target_r=target_r or 2.0, **kw))
    return out


_BARS: pd.DataFrame | None = None
_COSTS: engine.Costs | None = None


def _init(bars: pd.DataFrame, costs: engine.Costs) -> None:
    global _BARS, _COSTS
    _BARS, _COSTS = bars, costs


def _run_one(p: Params) -> pd.DataFrame:
    return engine.run(_BARS, p, _COSTS)


def backtest_all(bars: pd.DataFrame, params: list[Params], costs: engine.Costs, workers: int) -> list[pd.DataFrame]:
    if workers <= 1:
        _init(bars, costs)
        return [_run_one(p) for p in params]
    with ProcessPoolExecutor(max_workers=workers, initializer=_init, initargs=(bars, costs)) as ex:
        return list(ex.map(_run_one, params, chunksize=4))


def _slice(trades: pd.DataFrame, start, end) -> pd.DataFrame:
    if trades.empty:
        return trades
    d = pd.to_datetime(trades["date"])
    return trades[(d >= start) & (d < end)]


def score(trades: pd.DataFrame, min_trades: int) -> float:
    """Average R per trade, shrunk towards zero for small samples."""
    n = len(trades)
    if n < min_trades:
        return -np.inf
    return float(trades["r_multiple"].mean() * n / (n + 20))


def walk_forward(
    bars: pd.DataFrame,
    params: list[Params],
    results: list[pd.DataFrame],
    costs: engine.Costs,
    train_months: int,
    test_months: int,
    min_train_trades: int,
    baseline_runs: int,
) -> tuple[pd.DataFrame, pd.DataFrame, np.ndarray]:
    first = pd.Timestamp(bars.index[0].date()).to_period("M").to_timestamp()
    last = pd.Timestamp(bars.index[-1].date())
    folds, oos, base_sum, base_n = [], [], np.zeros(baseline_runs), 0
    start = first
    while True:
        train_end = start + pd.DateOffset(months=train_months)
        test_end = train_end + pd.DateOffset(months=test_months)
        if train_end > last:
            break
        scores = [score(_slice(r, start, train_end), min_train_trades) for r in results]
        best = int(np.argmax(scores))
        test = _slice(results[best], train_end, test_end)
        if np.isfinite(scores[best]):
            if not test.empty:
                oos.append(test.assign(fold=len(folds)))
                test_bars = bars[(bars.index.date >= train_end.date()) & (bars.index.date < test_end.date())]
                b = engine.random_baseline(test_bars, test, params[best], costs, runs=baseline_runs, seed=len(folds))
                base_sum += b * len(test)
                base_n += len(test)
            folds.append(
                {
                    "train": f"{start.date()} - {(train_end - pd.Timedelta(days=1)).date()}",
                    "test": f"{train_end.date()} - {(min(test_end, last + pd.Timedelta(days=1)) - pd.Timedelta(days=1)).date()}",
                    "params": describe(params[best]),
                    "train_avg_r": round(float(_slice(results[best], start, train_end)["r_multiple"].mean()), 3),
                    "test_trades": len(test),
                    "test_avg_r": round(float(test["r_multiple"].mean()), 3) if len(test) else np.nan,
                    "test_usd": round(float(test["pnl_usd"].sum()), 0) if len(test) else 0.0,
                }
            )
        start += pd.DateOffset(months=test_months)
    oos_trades = pd.concat(oos, ignore_index=True) if oos else pd.DataFrame()
    baseline = base_sum / base_n if base_n else np.array([])
    return pd.DataFrame(folds), oos_trades, baseline


def describe(p: Params) -> str:
    d, base = asdict(p), asdict(Params())
    parts = [f"{k}={v.strftime('%H:%M') if isinstance(v, time) else v}" for k, v in d.items() if v != base[k]]
    return ", ".join(parts) or "Standard"


def verdict(oos: pd.DataFrame, folds: pd.DataFrame, baseline: np.ndarray) -> tuple[bool, list[str]]:
    if oos.empty:
        return False, ["Keine Out-of-Sample-Trades."]
    s = metrics.summary(oos)
    pf = s["profit_factor"]
    positive = float((folds.loc[folds["test_trades"] > 0, "test_usd"] > 0).mean() * 100)
    base_pct = float((baseline >= s["expectancy_usd"]).mean() * 100) if baseline.size else 100.0
    checks = [
        (s["trades"] >= CRITERIA["min_trades"], f"Trades {s['trades']} >= {CRITERIA['min_trades']}"),
        (pf >= CRITERIA["min_profit_factor"], f"Profit Factor {pf} >= {CRITERIA['min_profit_factor']}"),
        (s["avg_r"] >= CRITERIA["min_avg_r"], f"Ø R {s['avg_r']} >= {CRITERIA['min_avg_r']}"),
        (positive >= CRITERIA["min_positive_folds_pct"],
         f"profitable Testphasen {positive:.0f}% >= {CRITERIA['min_positive_folds_pct']:.0f}%"),
        (base_pct <= CRITERIA["max_baseline_pct"],
         f"Zufalls-Baseline gleich gut oder besser in {base_pct:.1f}% <= {CRITERIA['max_baseline_pct']}%"),
    ]
    return all(ok for ok, _ in checks), [("OK    " if ok else "FAIL  ") + text for ok, text in checks]


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    src = ap.add_mutually_exclusive_group(required=True)
    src.add_argument("--csv")
    src.add_argument("--synthetic", action="store_true", help="made-up data with a built-in reversal")
    src.add_argument("--random-walk", action="store_true", help="made-up data without any edge (control)")
    ap.add_argument("--source-tz", default=data_loader.NY_TZ)
    ap.add_argument("--timeframe", choices=["1m", "5m"], default="5m")
    ap.add_argument("--train-months", type=int, default=12)
    ap.add_argument("--test-months", type=int, default=3)
    ap.add_argument("--min-train-trades", type=int, default=40)
    ap.add_argument("--tick-size", type=float, default=0.25)
    ap.add_argument("--tick-value", type=float, default=5.0)
    ap.add_argument("--slippage-ticks", type=float, default=1.0)
    ap.add_argument("--commission", type=float, default=4.0)
    ap.add_argument("--baseline-runs", type=int, default=200)
    ap.add_argument("--workers", type=int, default=os.cpu_count() or 1)
    ap.add_argument("--out", default="results")
    a = ap.parse_args(argv)

    rule = None if a.timeframe == "1m" else "5min"
    if a.csv:
        bars = data_loader.load_csv(a.csv, source_tz=a.source_tz, resample=rule)
    else:
        raw = synthetic.make_dataset(days=500) if a.synthetic else synthetic.random_walk(days=750)
        bars = data_loader.normalize(raw, resample=rule)
        print("WARNING: synthetic data - this only checks that the optimizer runs.\n")
    costs = engine.Costs(a.tick_size, a.tick_value, a.slippage_ticks, a.commission)

    params = combos()
    n_days = len(set(bars.index.date))
    print(f"{n_days} sessions, {bars.index[0].date()} -> {bars.index[-1].date()}, {len(params)} variants")
    results = backtest_all(bars, params, costs, a.workers)

    folds, oos, baseline = walk_forward(
        bars, params, results, costs, a.train_months, a.test_months, a.min_train_trades, a.baseline_runs
    )
    if folds.empty:
        print(f"Not enough history: need more than {a.train_months} months.")
        return 1

    insample = pd.DataFrame(
        [{"params": describe(p), **{k: v for k, v in metrics.summary(r, n_days).items() if k != "exit_reasons"}}
         for p, r in zip(params, results)]
    ).sort_values("avg_r", ascending=False)

    passed, checks = verdict(oos, folds, baseline)
    oos_stats = metrics.summary(oos, None) if not oos.empty else {"trades": 0}
    pd.set_option("display.width", 200)
    pd.set_option("display.max_colwidth", 90)
    print("\n== Walk-forward folds ==")
    print(folds.to_string(index=False))
    print("\n== Out-of-sample (nur ungesehene Testphasen) ==")
    for k, v in oos_stats.items():
        print(f"  {k:<18} {v}")
    print("\n== Verdict ==")
    print("\n".join("  " + c for c in checks))
    print("  => " + ("ROBUST: besteht alle Kriterien" if passed else "NICHT robust profitabel"))
    print("\n== Top 10 in-sample (gesamte Historie - überoptimistisch!) ==")
    print(insample.head(10).to_string(index=False))

    os.makedirs(a.out, exist_ok=True)
    folds.to_csv(os.path.join(a.out, "walkforward_folds.csv"), index=False)
    insample.to_csv(os.path.join(a.out, "insample_variants.csv"), index=False)
    if not oos.empty:
        oos.to_csv(os.path.join(a.out, "walkforward_oos_trades.csv"), index=False)
    print(f"\nCSV files written to {a.out}/")
    return 0 if passed else 2


if __name__ == "__main__":
    sys.exit(main())
