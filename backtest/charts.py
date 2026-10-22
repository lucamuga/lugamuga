"""Candlestick charts of each trade day and an equity curve, drawn with plain matplotlib."""

from __future__ import annotations

import os
from datetime import time

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import pandas as pd  # noqa: E402
from matplotlib.patches import Rectangle  # noqa: E402

from backtest.strategy import TEN_AM, Params  # noqa: E402

UP, DOWN = "#1a9850", "#d73027"


def _candles(ax, bars: pd.DataFrame) -> None:
    width = 0.7
    for i, (o, h, l, c) in enumerate(bars[["open", "high", "low", "close"]].itertuples(index=False)):
        color = UP if c >= o else DOWN
        ax.vlines(i, l, h, color=color, linewidth=0.8)
        ax.add_patch(Rectangle((i - width / 2, min(o, c)), width, max(abs(c - o), 1e-9), color=color))


def trade_chart(day: pd.DataFrame, trade, p: Params, path: str, chart_end: time = time(12, 0)) -> None:
    """Chart 09:30 to `chart_end` with the opening drive, the 10:00 range, entry, stop, target and exit."""
    t = day.index.time
    bars = day[t < max(chart_end, p.exit_time)]
    x_of = {ts: i for i, ts in enumerate(bars.index)}
    labels = [ts.strftime("%H:%M") for ts in bars.index]

    fig, ax = plt.subplots(figsize=(12, 5.5))
    _candles(ax, bars)

    ten = [i for i, ts in enumerate(bars.index) if ts.time() >= TEN_AM]
    if ten:
        ax.axvspan(-0.5, ten[0] - 0.5, color="#4575b4", alpha=0.07, label="Opening Drive 09:30-10:00")
        ax.axvline(ten[0] - 0.5, color="#4575b4", linestyle=":", linewidth=1)
        rng = bars[(bars.index.time >= TEN_AM) & (bars.index.time < p.range_end)]
        if not rng.empty and p.mode == "break":
            ax.add_patch(
                Rectangle(
                    (ten[0] - 0.5, rng["low"].min()),
                    len(rng),
                    rng["high"].max() - rng["low"].min(),
                    fill=False,
                    edgecolor="#4575b4",
                    linewidth=1.2,
                    label="10:00-Range",
                )
            )

    xe, xx = x_of.get(trade.entry_time, 0), x_of.get(trade.exit_time, len(bars) - 1)
    ax.hlines(trade.stop, xe, xx, color=DOWN, linestyle="--", linewidth=1, label=f"Stop {trade.stop:.2f}")
    ax.hlines(trade.target, xe, xx, color=UP, linestyle="--", linewidth=1, label=f"Ziel {trade.target:.2f}")
    marker = "^" if trade.direction == "long" else "v"
    ax.plot(xe, trade.entry_price, marker=marker, color="black", markersize=11, linestyle="none",
            label=f"Einstieg {trade.direction} {trade.entry_price:.2f}")
    ax.plot(xx, trade.exit_price, marker="X", color="#762a83", markersize=10, linestyle="none",
            label=f"Exit ({trade.exit_reason}) {trade.exit_price:.2f}")

    step = max(1, len(bars) // 16)
    ax.set_xticks(range(0, len(bars), step))
    ax.set_xticklabels(labels[::step], rotation=45, fontsize=8)
    ax.set_xlim(-1, len(bars))
    result = "Gewinn" if trade.pnl_usd > 0 else "Verlust"
    ax.set_title(
        f"{trade.date}  |  Drive {trade.drive_pct:+.2f}%  |  {result} {trade.pnl_usd:+.0f} $  ({trade.r_multiple:+.2f}R)"
    )
    ax.grid(alpha=0.25)
    ax.legend(loc="best", fontsize=8)
    fig.tight_layout()
    fig.savefig(path, dpi=110)
    plt.close(fig)


def equity_chart(trades: pd.DataFrame, path: str, title: str) -> None:
    fig, ax = plt.subplots(figsize=(10, 4))
    dates = pd.to_datetime(trades["date"])
    equity = trades["pnl_usd"].cumsum()
    ax.plot(dates, equity, marker="o", markersize=3)
    ax.bar(dates, trades["pnl_usd"], color=[UP if v > 0 else DOWN for v in trades["pnl_usd"]], alpha=0.35)
    ax.axhline(0, color="grey", linewidth=0.8)
    ax.set_title(title)
    ax.set_ylabel("$")
    ax.grid(alpha=0.3)
    fig.autofmt_xdate()
    fig.tight_layout()
    fig.savefig(path, dpi=120)
    plt.close(fig)


def write_all(bars: pd.DataFrame, trades: pd.DataFrame, p: Params, out_dir: str) -> list[str]:
    """One chart per trade in `out_dir/charts/`; returns the written paths."""
    chart_dir = os.path.join(out_dir, "charts")
    os.makedirs(chart_dir, exist_ok=True)
    days = {d: g for d, g in bars.groupby(bars.index.date)}
    paths = []
    for tr in trades.itertuples():
        path = os.path.join(chart_dir, f"{tr.date}.png")
        trade_chart(days[tr.date], tr, p, path)
        paths.append(path)
    return paths
