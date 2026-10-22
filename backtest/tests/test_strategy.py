from datetime import date, time

import pandas as pd
import pytest

from backtest import data_loader, engine
from backtest.strategy import Params, generate_signal
from backtest.synthetic import make_day

DAY = date(2024, 3, 5)  # Tuesday
NO_COSTS = engine.Costs(slippage_ticks=0, commission=0)


def test_reversal_after_up_drive_goes_short_and_hits_target():
    day = make_day(DAY, drive_pct=0.5, after_pct=-1.0)
    sig = generate_signal(day, Params())
    assert sig is not None and sig.direction == -1
    assert sig.entry_time.time() > time(10, 5)
    assert sig.stop == pytest.approx(day.loc[: sig.entry_time].iloc[:-1]["high"].max())

    trades = engine.run(day, Params(), NO_COSTS)
    assert len(trades) == 1
    tr = trades.iloc[0]
    assert tr.exit_reason == "target"
    assert tr.r_multiple == pytest.approx(2.0)


def test_down_drive_gives_long():
    sig = generate_signal(make_day(DAY, drive_pct=-0.5, after_pct=1.0), Params())
    assert sig is not None and sig.direction == 1 and sig.target > sig.entry_price > sig.stop


def test_continuation_day_has_no_break_signal():
    assert generate_signal(make_day(DAY, drive_pct=0.5, after_pct=0.8), Params()) is None


def test_small_drive_is_skipped():
    assert generate_signal(make_day(DAY, drive_pct=0.1, after_pct=-1.0), Params(min_drive_pct=0.2)) is None


def test_fade_mode_enters_at_ten_open():
    sig = generate_signal(make_day(DAY, drive_pct=0.5, after_pct=-1.0), Params(mode="fade"))
    assert sig.entry_time.time() == time(10, 0)


def test_signal_uses_no_future_bars():
    day = make_day(DAY, drive_pct=0.5, after_pct=-1.0)
    full = generate_signal(day, Params())
    truncated = generate_signal(day.loc[: full.entry_time], Params())
    assert truncated == full


def test_stop_wins_when_stop_and_target_in_same_bar():
    day = make_day(DAY, drive_pct=0.5, after_pct=-1.0)
    sig = generate_signal(day, Params())
    i = day.index.get_loc(sig.entry_time) + 1
    day.iloc[i, day.columns.get_loc("high")] = sig.stop + 10
    day.iloc[i, day.columns.get_loc("low")] = sig.target - 10
    tr = engine.run(day, Params(), NO_COSTS).iloc[0]
    assert tr.exit_reason == "stop"
    assert tr.r_multiple == pytest.approx(-1.0)


def test_time_exit_when_market_goes_flat():
    # Target far away: the reversal stalls after 11:00 and the position is closed on time.
    day = make_day(DAY, drive_pct=0.5, after_pct=-0.15)
    tr = engine.run(day, Params(target_r=100), NO_COSTS).iloc[0]
    assert tr.exit_reason == "time"
    assert tr.exit_time.time() < time(11, 30)


def test_costs_reduce_pnl():
    day = make_day(DAY, drive_pct=0.5, after_pct=-1.0)
    free = engine.run(day, Params(), NO_COSTS).iloc[0]
    paid = engine.run(day, Params(), engine.Costs(slippage_ticks=1, commission=4)).iloc[0]
    assert paid.pnl_usd == pytest.approx(free.pnl_usd - 2 * 5.0 - 4.0)


def test_filters_skip_weekday_and_non_news_days():
    day = make_day(DAY, drive_pct=0.5, after_pct=-1.0)
    assert engine.run(day, Params(), f=engine.Filters(weekdays=(0,))).empty
    assert engine.run(day, Params(), f=engine.Filters(news_only=True)).empty
    assert len(engine.run(day, Params(), f=engine.Filters(news_only=True, news_dates=frozenset({DAY})))) == 1


def test_loader_converts_utc_and_keeps_rth_only(tmp_path):
    idx = pd.date_range("2024-03-05 14:00", "2024-03-05 21:30", freq="1min", tz="UTC")
    df = pd.DataFrame({"Datetime": idx.tz_localize(None), "Open": 1.0, "High": 2.0, "Low": 0.5, "Close": 1.5})
    path = tmp_path / "bars.csv"
    df.to_csv(path, index=False)
    bars = data_loader.load_csv(str(path), source_tz="UTC", resample="5min")
    assert bars.index[0].time() == time(9, 30)  # 14:30 UTC = 09:30 EST
    assert bars.index[-1].time() == time(15, 55)
    assert (bars.index.minute % 5 == 0).all()


def test_five_minute_run_writes_report_and_charts(tmp_path):
    from backtest import run_backtest

    assert run_backtest.main(["--synthetic", "--last-days", "10", "--baseline-runs", "5", "--out", str(tmp_path)]) == 0
    assert (tmp_path / "report.md").read_text().startswith("# 10:00-ET-Reversal")
    assert (tmp_path / "equity.png").exists()
    assert list((tmp_path / "charts").glob("*.png"))


def test_breakeven_moves_stop_to_entry():
    # Short reaches 1.3R, then the market chops around just below the entry and pokes above it.
    day = make_day(DAY, drive_pct=0.5, after_pct=-1.0)
    sig = generate_signal(day, Params())
    i = day.index.get_loc(sig.entry_time)
    e, r = sig.entry_price, sig.risk
    later = day.index[i + 2:]
    day.loc[later, ["open", "close"]] = e - 0.2 * r
    day.loc[later, "high"], day.loc[later, "low"] = e + 0.5 * r, e - 0.3 * r
    day.iloc[i + 1, day.columns.get_loc("low")] = e - 1.3 * r
    plain = engine.run(day, Params(), NO_COSTS).iloc[0]
    be = engine.run(day, Params(breakeven_r=1.0), NO_COSTS).iloc[0]
    assert be.exit_reason == "breakeven" and be.r_multiple == pytest.approx(0.0)
    assert plain.exit_reason == "time" and plain.r_multiple == pytest.approx(0.2)


def test_max_risk_filter_skips_wide_stops():
    day = make_day(DAY, drive_pct=0.5, after_pct=-1.0)
    sig = generate_signal(day, Params())
    risk_pct = sig.risk / sig.entry_price * 100
    assert generate_signal(day, Params(max_risk_pct=risk_pct / 2)) is None
    assert generate_signal(day, Params(max_risk_pct=risk_pct * 2)) is not None


def test_adr_filter_needs_history_and_size():
    day = make_day(DAY, drive_pct=0.5, after_pct=-1.0)
    p = Params(drive_adr_min=0.5)
    assert generate_signal(day, p, adr_pct=None) is None
    assert generate_signal(day, p, adr_pct=0.8) is not None  # drive 0.5% >= 0.5 * 0.8%
    assert generate_signal(day, p, adr_pct=1.2) is None  # drive 0.5% < 0.5 * 1.2%


def test_walk_forward_rejects_random_walk():
    from backtest import optimize, synthetic

    bars = data_loader.normalize(synthetic.random_walk(days=330, seed=3), resample="5min")
    params = [Params(), Params(mode="fade", target="retrace"), Params(min_drive_pct=0.5, target_r=3.0)]
    costs = engine.Costs()
    results = optimize.backtest_all(bars, params, costs, workers=1)
    folds, oos, base = optimize.walk_forward(bars, params, results, costs, 6, 3, 10, 20)
    assert len(folds) >= 3
    passed, _ = optimize.verdict(oos, folds, base)
    assert not passed
