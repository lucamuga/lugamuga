"""Performance statistics for a trade list produced by `engine.run`."""

from __future__ import annotations

import numpy as np
import pandas as pd

WEEKDAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]


def _max_drawdown(equity: pd.Series) -> float:
    peak = equity.cummax().clip(lower=0)
    return float((equity - peak).min()) if len(equity) else 0.0


def summary(trades: pd.DataFrame, trading_days: int | None = None) -> dict:
    if trades.empty:
        return {"trades": 0}
    pnl = trades["pnl_usd"]
    wins, losses = pnl[pnl > 0], pnl[pnl <= 0]
    equity = pnl.cumsum()

    # Sharpe on daily P&L, counting non-trading days as zero.
    n_days = max(trading_days or len(trades), len(trades))
    daily = np.concatenate([pnl.to_numpy(), np.zeros(n_days - len(trades))])
    sharpe = float(daily.mean() / daily.std(ddof=1) * np.sqrt(252)) if daily.std(ddof=1) > 0 else float("nan")

    return {
        "trades": int(len(trades)),
        "trade_days_pct": round(100.0 * len(trades) / n_days, 1),
        "win_rate_pct": round(100.0 * len(wins) / len(trades), 1),
        "avg_r": round(float(trades["r_multiple"].mean()), 3),
        "expectancy_usd": round(float(pnl.mean()), 2),
        "profit_factor": round(float(wins.sum() / -losses.sum()), 2) if losses.sum() < 0 else float("inf"),
        "total_usd": round(float(pnl.sum()), 2),
        "max_drawdown_usd": round(_max_drawdown(equity), 2),
        "sharpe_daily_ann": round(sharpe, 2),
        "exit_reasons": trades["exit_reason"].value_counts().to_dict(),
    }


def breakdown(trades: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Win rate, average R and P&L grouped by weekday, direction and news day."""
    if trades.empty:
        return {}
    df = trades.assign(
        weekday=pd.to_datetime(trades["date"]).dt.weekday.map(lambda i: WEEKDAYS[i]),
        win=trades["pnl_usd"] > 0,
    )
    out = {}
    for key in ("weekday", "direction", "news_day"):
        if key not in df:
            continue
        g = df.groupby(key)
        out[key] = pd.DataFrame(
            {
                "trades": g.size(),
                "win_rate_pct": (g["win"].mean() * 100).round(1),
                "avg_r": g["r_multiple"].mean().round(3),
                "total_usd": g["pnl_usd"].sum().round(2),
            }
        )
    if "weekday" in out:
        out["weekday"] = out["weekday"].reindex([d for d in WEEKDAYS if d in out["weekday"].index])
    return out
