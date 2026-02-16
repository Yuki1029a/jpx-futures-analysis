"""Market snapshot analysis — option OI positioning + price context.

Outputs a structured summary of:
1. NK225 option OI distribution by strike (PUT/CALL)
2. Key participants' positioning
3. Price context from yfinance
4. P/C ratio dynamics
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pandas as pd
import numpy as np
from datetime import date, timedelta

# ── Market data via yfinance ──
try:
    import yfinance as yf
except ImportError:
    print("yfinance not installed. Run: pip install yfinance")
    sys.exit(1)

# ── Internal imports ──
from data.aggregator import (
    load_option_weekly_data,
    load_weekly_data,
    load_daily_futures_oi,
    build_available_weeks,
    SESSION_MODES,
)

# ── Parameters ──
END_DATE = date(2026, 2, 14)
LOOKBACK_DAYS = 90
START_DATE = END_DATE - timedelta(days=LOOKBACK_DAYS)

print("=" * 70)
print(f"  市場スナップショット分析  {END_DATE}")
print("=" * 70)

# ────────────────────────────────────────────
# 1. Price data
# ────────────────────────────────────────────
print("\n[1] 市場価格データ取得中...")

tickers = {
    "^N225": "日経225",
    "^TPX": "TOPIX",
    "JPY=X": "USD/JPY",
    "^VIX": "VIX",
}

for ticker, name in tickers.items():
    try:
        data = yf.download(ticker, start=str(START_DATE), end=str(END_DATE + timedelta(days=1)),
                           progress=False, auto_adjust=True)
        if data.empty:
            print(f"  {name}: データなし")
            continue

        last = data.iloc[-1]
        prev = data.iloc[-2] if len(data) > 1 else last

        close = float(last["Close"].iloc[0]) if hasattr(last["Close"], "iloc") else float(last["Close"])
        prev_close = float(prev["Close"].iloc[0]) if hasattr(prev["Close"], "iloc") else float(prev["Close"])
        chg = close - prev_close
        chg_pct = chg / prev_close * 100

        # 20d stats
        recent = data.tail(20)
        high_20d = float(recent["High"].max())
        low_20d = float(recent["Low"].min())
        avg_20d = float(recent["Close"].mean())

        # 60d stats
        recent_60 = data.tail(60)
        high_60d = float(recent_60["High"].max()) if len(recent_60) > 0 else None
        low_60d = float(recent_60["Low"].min()) if len(recent_60) > 0 else None

        print(f"\n  {name} ({ticker})")
        print(f"    終値: {close:,.2f}  ({chg:+,.2f} / {chg_pct:+.2f}%)")
        print(f"    20日: 高値 {high_20d:,.2f} / 安値 {low_20d:,.2f} / 平均 {avg_20d:,.2f}")
        if high_60d:
            print(f"    60日: 高値 {high_60d:,.2f} / 安値 {low_60d:,.2f}")
    except Exception as e:
        print(f"  {name}: エラー - {e}")

# ────────────────────────────────────────────
# 2. NK225 Option OI analysis (latest week)
# ────────────────────────────────────────────
print("\n\n[2] オプション建玉分析...")

weeks = build_available_weeks()
if not weeks:
    print("  週データなし")
    sys.exit(0)

latest_week = weeks[0]  # most recent
print(f"  対象週: {latest_week.label}")
print(f"  営業日: {', '.join(str(d) for d in latest_week.trading_days)}")

# Load option data (全セッション合計)
session_keys = SESSION_MODES["全セッション合計"]

# Try 2026年02月限 first, then 03月限
for cm in ["2602", "2603"]:
    opt_rows = load_option_weekly_data(
        latest_week,
        contract_month=cm,
        session_keys=session_keys,
        participant_ids=None,
    )
    if opt_rows:
        print(f"  限月: {cm}")
        break

if not opt_rows:
    print("  オプションデータなし")
else:
    # Aggregate PUT/CALL totals
    put_vol_total = sum(r.put_week_total or 0 for r in opt_rows)
    call_vol_total = sum(r.call_week_total or 0 for r in opt_rows)
    pcr_vol = put_vol_total / call_vol_total if call_vol_total > 0 else 0

    print(f"\n  === P/C Ratio (出来高) ===")
    print(f"    PUT出来高: {int(put_vol_total):,}")
    print(f"    CALL出来高: {int(call_vol_total):,}")
    print(f"    P/C比率: {pcr_vol:.3f}")

    # OI at end of period (per strike)
    latest_td = latest_week.trading_days[-1] if latest_week.trading_days else None

    if latest_td:
        print(f"\n  === 建玉残高 (最終日: {latest_td}) ===")

        put_oi_total = sum(r.put_daily_oi.get(latest_td, 0) for r in opt_rows)
        call_oi_total = sum(r.call_daily_oi.get(latest_td, 0) for r in opt_rows)
        pcr_oi = put_oi_total / call_oi_total if call_oi_total > 0 else 0

        print(f"    PUT建玉: {int(put_oi_total):,}")
        print(f"    CALL建玉: {int(call_oi_total):,}")
        print(f"    P/C比率(建玉): {pcr_oi:.3f}")

    # Top strikes by OI
    print(f"\n  === PUT建玉上位ストライク ===")
    put_oi_by_strike = []
    for r in opt_rows:
        if latest_td:
            oi = r.put_daily_oi.get(latest_td, 0)
            chg = r.put_daily_oi_change.get(latest_td, 0)
            if oi > 0:
                put_oi_by_strike.append((r.strike_price, oi, chg))
    put_oi_by_strike.sort(key=lambda x: x[1], reverse=True)
    for strike, oi, chg in put_oi_by_strike[:10]:
        print(f"    {strike:>7,}  OI: {oi:>8,}  変化: {chg:>+7,}")

    print(f"\n  === CALL建玉上位ストライク ===")
    call_oi_by_strike = []
    for r in opt_rows:
        if latest_td:
            oi = r.call_daily_oi.get(latest_td, 0)
            chg = r.call_daily_oi_change.get(latest_td, 0)
            if oi > 0:
                call_oi_by_strike.append((r.strike_price, oi, chg))
    call_oi_by_strike.sort(key=lambda x: x[1], reverse=True)
    for strike, oi, chg in call_oi_by_strike[:10]:
        print(f"    {strike:>7,}  OI: {oi:>8,}  変化: {chg:>+7,}")

    # Volume concentration
    print(f"\n  === 出来高上位ストライク (PUT) ===")
    put_vol_strikes = [(r.strike_price, r.put_week_total or 0) for r in opt_rows if (r.put_week_total or 0) > 0]
    put_vol_strikes.sort(key=lambda x: x[1], reverse=True)
    for strike, vol in put_vol_strikes[:10]:
        print(f"    {strike:>7,}  出来高: {int(vol):>8,}")

    print(f"\n  === 出来高上位ストライク (CALL) ===")
    call_vol_strikes = [(r.strike_price, r.call_week_total or 0) for r in opt_rows if (r.call_week_total or 0) > 0]
    call_vol_strikes.sort(key=lambda x: x[1], reverse=True)
    for strike, vol in call_vol_strikes[:10]:
        print(f"    {strike:>7,}  出来高: {int(vol):>8,}")

    # Participant breakdown for top strike
    if latest_td and put_oi_by_strike:
        top_put_strike = put_oi_by_strike[0][0]
        for r in opt_rows:
            if r.strike_price == top_put_strike:
                bd = r.put_daily_breakdown.get(latest_td, [])
                if bd:
                    print(f"\n  === PUT {top_put_strike:,} 参加者内訳 ({latest_td}) ===")
                    for name, vol in bd[:8]:
                        print(f"    {name:<40s} {int(vol):>8,}")
                break

    if latest_td and call_oi_by_strike:
        top_call_strike = call_oi_by_strike[0][0]
        for r in opt_rows:
            if r.strike_price == top_call_strike:
                bd = r.call_daily_breakdown.get(latest_td, [])
                if bd:
                    print(f"\n  === CALL {top_call_strike:,} 参加者内訳 ({latest_td}) ===")
                    for name, vol in bd[:8]:
                        print(f"    {name:<40s} {int(vol):>8,}")
                break

# ────────────────────────────────────────────
# 3. Futures positioning
# ────────────────────────────────────────────
print("\n\n[3] 先物ポジショニング...")

fut_rows = load_weekly_data(
    latest_week,
    product="NK225F",
    contract_month="2603",
    session_keys=session_keys,
    include_oi=True,
)

if fut_rows:
    print(f"\n  === 主要参加者ポジション変動 ===")
    # Sort by absolute net change
    sorted_rows = sorted(fut_rows, key=lambda r: abs(r.oi_net_change or 0), reverse=True)
    for r in sorted_rows[:12]:
        direction = r.inferred_direction or ""
        net_chg = r.oi_net_change or 0
        total_vol = sum(r.daily_volumes.values())
        print(f"    {r.participant_name:<40s}  "
              f"NetΔ: {net_chg:>+8,.0f}  "
              f"Vol: {int(total_vol):>8,}  "
              f"{direction}")
else:
    print("  先物データなし")

print("\n" + "=" * 70)
print("  分析完了")
print("=" * 70)
