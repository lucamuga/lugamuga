"""Load intraday OHLCV bars and normalise them to New York regular trading hours."""

from __future__ import annotations

from datetime import time

import pandas as pd

NY_TZ = "America/New_York"
RTH_START = time(9, 30)
RTH_END = time(16, 0)

_DATETIME_CANDIDATES = ("datetime", "timestamp", "date", "time", "ts_event")


def _find_datetime_column(df: pd.DataFrame) -> str:
    for name in _DATETIME_CANDIDATES:
        if name in df.columns:
            return name
    raise ValueError(
        f"No datetime column found. Expected one of {_DATETIME_CANDIDATES}, got {list(df.columns)}"
    )


def normalize(df: pd.DataFrame, source_tz: str = NY_TZ, resample: str | None = None) -> pd.DataFrame:
    """Return bars indexed by NY time, restricted to 09:30-16:00, optionally resampled.

    Naive timestamps are interpreted as `source_tz`; aware ones are converted.
    """
    df = df.copy()
    df.columns = [c.strip().lower() for c in df.columns]
    missing = {"open", "high", "low", "close"} - set(df.columns)
    if missing:
        raise ValueError(f"Missing OHLC columns: {sorted(missing)}")
    if "volume" not in df.columns:
        df["volume"] = 0

    idx = df.index
    if not isinstance(idx, pd.DatetimeIndex):
        col = _find_datetime_column(df)
        raw = df.pop(col)
        if pd.api.types.is_numeric_dtype(raw):
            # Unix seconds (TradingView exports) or nanoseconds (Databento).
            unit = "ns" if raw.abs().max() > 1e12 else "s"
            idx = pd.DatetimeIndex(pd.to_datetime(raw, unit=unit, utc=True))
        else:
            idx = pd.DatetimeIndex(pd.to_datetime(raw, utc=False, format="mixed"))
    if idx.tz is None:
        idx = idx.tz_localize(source_tz, ambiguous="infer", nonexistent="shift_forward")
    df.index = idx.tz_convert(NY_TZ)
    df.index.name = "datetime"

    df = df[["open", "high", "low", "close", "volume"]].astype(float).sort_index()
    df = df[~df.index.duplicated(keep="first")]

    if resample:
        df = (
            df.resample(resample, label="left", closed="left")
            .agg({"open": "first", "high": "max", "low": "min", "close": "last", "volume": "sum"})
            .dropna(subset=["open"])
        )

    t = df.index.time
    return df[(t >= RTH_START) & (t < RTH_END)]


def load_csv(path: str, source_tz: str = NY_TZ, resample: str | None = None) -> pd.DataFrame:
    """Load a CSV with a datetime column plus open/high/low/close[/volume]."""
    return normalize(pd.read_csv(path), source_tz=source_tz, resample=resample)


def load_yfinance(symbol: str = "QQQ", interval: str = "1m", period: str | None = None) -> pd.DataFrame:
    """Download bars via yfinance (only ~30 days of 1m / ~60 days of 5m history)."""
    import yfinance as yf

    if period is None:
        period = "7d" if interval == "1m" else "60d"
    raw = yf.download(symbol, interval=interval, period=period, progress=False, auto_adjust=False)
    if isinstance(raw.columns, pd.MultiIndex):
        raw.columns = raw.columns.get_level_values(0)
    return normalize(raw, source_tz="UTC")
