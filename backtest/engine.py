"""Day-by-day simulation of the 10 AM reversal with conservative fills and trading costs."""

from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import pandas as pd

from backtest.strategy import Params, Signal, generate_signal


@dataclass(frozen=True)
class Costs:
    tick_size: float = 0.25  # NQ / MNQ
    tick_value: float = 5.0  # $ per tick per contract (NQ 5.0, MNQ 0.5)
    slippage_ticks: float = 1.0  # per side, always adverse
    commission: float = 4.0  # $ per contract round trip
    contracts: int = 1

    @property
    def point_value(self) -> float:
        return self.tick_value / self.tick_size


@dataclass(frozen=True)
class Filters:
    weekdays: tuple[int, ...] = (0, 1, 2, 3, 4)  # Monday = 0
    news_dates: frozenset = frozenset()  # datetime.date values with a 10:00 ET release
    news_only: bool = False
    max_gap_pct: float | None = None  # skip days whose open gaps more than this vs. prior close


def simulate_exit(day: pd.DataFrame, sig: Signal, p: Params) -> tuple[pd.Timestamp, float, str]:
    """Walk bars from the entry bar on. If stop and target fall in the same bar, assume the stop."""
    t = day.index.time
    window = day[(day.index >= sig.entry_time) & (t < p.exit_time)]
    for ts, bar in window.iterrows():
        if sig.direction == 1:
            if bar["low"] <= sig.stop:
                return ts, min(sig.stop, bar["open"]), "stop"
            if bar["high"] >= sig.target:
                return ts, max(sig.target, bar["open"]), "target"
        else:
            if bar["high"] >= sig.stop:
                return ts, max(sig.stop, bar["open"]), "stop"
            if bar["low"] <= sig.target:
                return ts, min(sig.target, bar["open"]), "target"
    last = window.iloc[-1]
    return window.index[-1], float(last["close"]), "time"


def _trade_record(day: pd.DataFrame, sig: Signal, p: Params, c: Costs) -> dict:
    exit_time, exit_price, reason = simulate_exit(day, sig, p)
    slip = c.slippage_ticks * c.tick_size
    gross_pts = (exit_price - sig.entry_price) * sig.direction
    net_pts = gross_pts - 2 * slip
    pnl = net_pts * c.point_value * c.contracts - c.commission * c.contracts
    return {
        "date": sig.entry_time.date(),
        "direction": "long" if sig.direction == 1 else "short",
        "drive_pct": round(sig.drive_pct, 4),
        "entry_time": sig.entry_time,
        "entry_price": sig.entry_price,
        "stop": sig.stop,
        "target": sig.target,
        "exit_time": exit_time,
        "exit_price": exit_price,
        "exit_reason": reason,
        "risk_pts": sig.risk,
        "pnl_pts": net_pts,
        # R includes costs: commission is converted to points of the same contract.
        "r_multiple": (pnl / (c.point_value * c.contracts)) / sig.risk,
        "pnl_usd": pnl,
    }


def iter_days(bars: pd.DataFrame):
    for date, day in bars.groupby(bars.index.date):
        yield date, day


def run(
    bars: pd.DataFrame,
    p: Params = Params(),
    c: Costs = Costs(),
    f: Filters = Filters(),
) -> pd.DataFrame:
    """Run the strategy over every session in `bars` and return one row per trade."""
    rows = []
    prev_close = None
    for date, day in iter_days(bars):
        day_open, day_close = day["open"].iloc[0], day["close"].iloc[-1]
        gap_pct = None if prev_close is None else (day_open - prev_close) / prev_close * 100.0
        prev_close = day_close

        is_news = date in f.news_dates
        if date.weekday() not in f.weekdays or (f.news_only and not is_news):
            continue
        if f.max_gap_pct is not None and gap_pct is not None and abs(gap_pct) > f.max_gap_pct:
            continue

        sig = generate_signal(day, p)
        if sig is None:
            continue
        row = _trade_record(day, sig, p, c)
        row["news_day"] = is_news
        row["gap_pct"] = gap_pct
        rows.append(row)

    return pd.DataFrame(rows)


def random_baseline(
    bars: pd.DataFrame,
    trades: pd.DataFrame,
    p: Params = Params(),
    c: Costs = Costs(),
    runs: int = 200,
    seed: int = 0,
) -> np.ndarray:
    """Expectancy ($/trade) of coin-flip direction on the same days, entry times and risk.

    If the strategy's expectancy is not clearly above this distribution, its edge is
    likely the market's drift or noise rather than the 10 AM reversal.
    """
    if trades.empty:
        return np.array([])
    days = dict(iter_days(bars))
    rng = np.random.default_rng(seed)
    results = []
    for _ in range(runs):
        pnls = []
        for tr in trades.itertuples():
            d = 1 if rng.random() < 0.5 else -1
            entry, risk = tr.entry_price, tr.risk_pts
            reward = abs(tr.target - entry)
            sig = Signal(d, tr.entry_time, entry, entry - d * risk, entry + d * reward, tr.drive_pct)
            pnls.append(_trade_record(days[tr.date], sig, p, c)["pnl_usd"])
        results.append(float(np.mean(pnls)))
    return np.array(results)
