"""Nasdaq 10:00 ET reversal: fade the opening drive once the 10 AM data release turns the market."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import time

import pandas as pd

OPEN = time(9, 30)
TEN_AM = time(10, 0)


@dataclass(frozen=True)
class Params:
    mode: str = "break"  # "break": wait for a break of the 10:00 range; "fade": enter at the 10:00 open
    min_drive_pct: float = 0.20  # minimum |09:30 -> 09:59| move in percent
    range_end: time = time(10, 5)  # 10:00 range is built from bars in [10:00, range_end)
    setup_end: time = time(10, 15)  # no new entries at or after this time
    stop_buffer_pct: float = 0.0  # extra room beyond the drive extreme, in percent of price
    target: str = "r"  # "r": target_r multiples of risk; "retrace": 50% retracement of the drive
    target_r: float = 2.0
    exit_time: time = time(11, 30)  # flat at the close of the last bar before this time


@dataclass(frozen=True)
class Signal:
    direction: int  # +1 long, -1 short
    entry_time: pd.Timestamp
    entry_price: float
    stop: float
    target: float
    drive_pct: float

    @property
    def risk(self) -> float:
        return abs(self.entry_price - self.stop)


def opening_drive_pct(day: pd.DataFrame) -> float | None:
    """Percent move from the 09:30 open to the last close before 10:00."""
    t = day.index.time
    drive = day[(t >= OPEN) & (t < TEN_AM)]
    if drive.empty:
        return None
    first_open = drive["open"].iloc[0]
    return (drive["close"].iloc[-1] - first_open) / first_open * 100.0


def generate_signal(day: pd.DataFrame, p: Params = Params()) -> Signal | None:
    """Return the day's trade setup, using only bars up to and including the entry bar."""
    drive_pct = opening_drive_pct(day)
    if drive_pct is None or abs(drive_pct) < p.min_drive_pct:
        return None
    direction = -1 if drive_pct > 0 else 1

    t = day.index.time
    entry_idx = None
    if p.mode == "fade":
        after = (t >= TEN_AM) & (t < p.setup_end)
        if after.any():
            entry_idx = int(after.argmax())
    elif p.mode == "break":
        rng = day[(t >= TEN_AM) & (t < p.range_end)]
        if rng.empty:
            return None
        rng_high, rng_low = rng["high"].max(), rng["low"].min()
        for i in range(len(day) - 1):
            if t[i] < p.range_end:
                continue
            if t[i] >= p.setup_end:
                break
            close = day["close"].iloc[i]
            if (direction == 1 and close > rng_high) or (direction == -1 and close < rng_low):
                entry_idx = i + 1  # fill at the next bar's open
                break
    else:
        raise ValueError(f"Unknown mode: {p.mode!r}")

    if entry_idx is None or t[entry_idx] >= p.exit_time:
        return None

    entry_price = float(day["open"].iloc[entry_idx])
    before = day.iloc[:entry_idx]
    buffer = entry_price * p.stop_buffer_pct / 100.0
    if direction == 1:
        stop = float(before["low"].min()) - buffer
    else:
        stop = float(before["high"].max()) + buffer
    risk = (entry_price - stop) * direction
    if risk <= 0:
        return None

    if p.target == "r":
        target = entry_price + direction * p.target_r * risk
    elif p.target == "retrace":
        drive = day[(t >= OPEN) & (t < TEN_AM)]
        start = float(drive["open"].iloc[0])
        extreme = float(drive["low"].min() if direction == 1 else drive["high"].max())
        target = extreme + 0.5 * (start - extreme)
        if (target - entry_price) * direction <= 0:
            return None
    else:
        raise ValueError(f"Unknown target: {p.target!r}")

    return Signal(
        direction=direction,
        entry_time=day.index[entry_idx],
        entry_price=entry_price,
        stop=stop,
        target=target,
        drive_pct=drive_pct,
    )
