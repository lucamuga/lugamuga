"""Synthetic 1-minute sessions with a known opening drive and 10:00 behaviour, for tests and demos.

These prices are made up. Results on them say nothing about the real market.
"""

from __future__ import annotations

from datetime import date

import numpy as np
import pandas as pd

from backtest.data_loader import NY_TZ


def make_day(
    day: date,
    start_price: float = 18000.0,
    drive_pct: float = 0.5,
    after_pct: float = -0.8,
    noise_pct: float = 0.0,
    rng: np.random.Generator | None = None,
) -> pd.DataFrame:
    """One RTH session: linear drive of `drive_pct` until 10:00, a 5-minute pause, then a move of
    `after_pct` (relative to the 10:00 price) until 11:00, then flat until the close."""
    idx = pd.date_range(f"{day} 09:30", f"{day} 15:59", freq="1min", tz=NY_TZ)
    n = len(idx)
    path = np.empty(n)
    peak = start_price * (1 + drive_pct / 100)
    end = peak * (1 + after_pct / 100)
    for i in range(n):
        if i < 30:  # 09:30-09:59
            path[i] = start_price + (peak - start_price) * (i + 1) / 30
        elif i < 35:  # 10:00-10:04: tight range around the 10:00 price
            path[i] = peak + (-1) ** i * abs(peak - start_price) * 0.02
        elif i < 90:  # 10:05-10:59
            path[i] = peak + (end - peak) * (i - 34) / 55
        else:
            path[i] = end
    if noise_pct and rng is not None:
        path = path * (1 + rng.normal(0, noise_pct / 100, n))

    close = path
    open_ = np.concatenate([[start_price], close[:-1]])
    wick = start_price * 0.0001
    high = np.maximum(open_, close) + wick
    low = np.minimum(open_, close) - wick
    return pd.DataFrame(
        {"open": open_, "high": high, "low": low, "close": close, "volume": 1000.0},
        index=pd.DatetimeIndex(idx, name="datetime"),
    )


def make_dataset(
    days: int = 250,
    reversal_prob: float = 0.6,
    seed: int = 42,
    start: str = "2024-01-02",
) -> pd.DataFrame:
    """Business days where the opening drive reverses with probability `reversal_prob`."""
    rng = np.random.default_rng(seed)
    price = 18000.0
    frames = []
    for d in pd.bdate_range(start, periods=days):
        drive = rng.choice([-1, 1]) * rng.uniform(0.05, 0.8)
        follow = rng.uniform(0.2, 1.0) * (-np.sign(drive) if rng.random() < reversal_prob else np.sign(drive))
        df = make_day(d.date(), price, drive, follow, noise_pct=0.02, rng=rng)
        frames.append(df)
        price = float(df["close"].iloc[-1])
    return pd.concat(frames)
