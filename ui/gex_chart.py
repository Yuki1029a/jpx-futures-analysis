"""GEX (Gamma Exposure) chart rendering.

Display unit: hedge amount per 100-yen move in the underlying (in 億円).
Also shows equivalent NK225 futures (large) contracts.
"""
from __future__ import annotations

import streamlit as st
import pandas as pd
from datetime import date

from models import OptionStrikeRow, WeekDefinition
import plotly.graph_objects as go

from utils.gex import calc_gex_profile, calc_gex_surface, get_sq_date

_DOW_JP = ["月", "火", "水", "木", "金", "土", "日"]
_OKU = 1e8          # 1億円
_MOVE = 100          # 表示基準: 原資産100円変動
_FUTURES_MULT = 1000 # 先物ラージ1枚 = 日経225 x 1000


def render_gex_section(
    rows: list[OptionStrikeRow],
    week: WeekDefinition,
    contract_month: str,
) -> None:
    """Render the GEX analysis section."""
    st.subheader(f"ガンマエクスポージャー ({week.label})")

    if not rows:
        st.warning("オプションデータがありません。")
        return

    # --- Controls ---
    c1, c2, c3 = st.columns(3)
    with c1:
        spot = st.number_input(
            "原資産価格 (日経225)",
            value=38500, min_value=10000, max_value=60000, step=100,
            key="gex_spot",
        )
    with c2:
        iv_pct = st.slider("IV (%)", 5, 60, 20, key="gex_iv")
    with c3:
        sq = get_sq_date(contract_month)
        st.metric("SQ日", sq.strftime("%Y/%m/%d"))

    sigma = iv_pct / 100.0

    # --- Extract OI from latest available date ---
    as_of, put_oi, call_oi, all_strikes = _extract_latest_oi(rows)

    if not all_strikes:
        st.info("建玉データがありません。")
        return

    dow = _DOW_JP[as_of.weekday()]
    st.caption(f"建玉基準日: {as_of.strftime('%Y/%m/%d')}({dow})  |  "
               f"PUT OI計: {sum(put_oi.values()):,}枚  CALL OI計: {sum(call_oi.values()):,}枚")

    # --- Calculate ---
    profile = calc_gex_profile(
        strikes=sorted(all_strikes),
        put_oi=put_oi,
        call_oi=call_oi,
        spot=float(spot),
        expiry_date=sq,
        as_of=as_of,
        sigma=sigma,
    )

    # Scale: per 100-yen move
    net_100 = profile.total_net_gex * _MOVE / _OKU
    call_100 = profile.total_call_gex * _MOVE / _OKU
    put_100 = profile.total_put_gex * _MOVE / _OKU

    # Futures equivalent: GEX per 1 yen / (spot * futures_mult) = contracts per 1 yen
    # Per 100 yen: contracts * 100
    futures_per_100 = profile.total_net_gex * _MOVE / (spot * _FUTURES_MULT)

    # --- Summary metrics ---
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.metric("Net GEX /100円", f"{net_100:+,.0f} 億円")
    with m2:
        st.metric("CALL GEX /100円", f"{call_100:+,.0f} 億円")
    with m3:
        st.metric("PUT GEX /100円", f"{put_100:+,.0f} 億円")
    with m4:
        if profile.flip_point:
            st.metric("フリップ", f"{profile.flip_point:,.0f}")
        else:
            st.metric("フリップ", "N/A")

    # Sub-metrics
    s1, s2 = st.columns(2)
    with s1:
        st.caption(f"先物ラージ換算 (100円変動時): {futures_per_100:+,.0f} 枚")
    with s2:
        days_to_sq = max((sq - as_of).days, 0)
        st.caption(f"SQ残: {days_to_sq}日  |  IV: {iv_pct}%")

    # --- Chart ---
    _render_gex_bar_chart(profile.df, spot)

    # --- Detail table ---
    with st.expander("GEX詳細テーブル"):
        _render_gex_table(profile.df, spot)

    # --- 3D Surface ---
    st.markdown("---")
    _render_gex_3d_surface(
        sorted(all_strikes), put_oi, call_oi,
        float(spot), sq, as_of, sigma,
    )


def _extract_latest_oi(
    rows: list[OptionStrikeRow],
) -> tuple[date, dict[int, int], dict[int, int], set[int]]:
    """Extract the latest available daily OI across all strikes."""
    all_dates: set[date] = set()
    for r in rows:
        all_dates.update(r.put_daily_oi.keys())
        all_dates.update(r.call_daily_oi.keys())

    if not all_dates:
        return date.today(), {}, {}, set()

    latest = max(all_dates)
    put_oi: dict[int, int] = {}
    call_oi: dict[int, int] = {}
    all_strikes: set[int] = set()

    for r in rows:
        p = r.put_daily_oi.get(latest, 0)
        c = r.call_daily_oi.get(latest, 0)
        if p > 0 or c > 0:
            all_strikes.add(r.strike_price)
            if p > 0:
                put_oi[r.strike_price] = p
            if c > 0:
                call_oi[r.strike_price] = c

    return latest, put_oi, call_oi, all_strikes


def _render_gex_bar_chart(df: pd.DataFrame, spot: float) -> None:
    """Render GEX bar chart. Values scaled to per-100-yen-move in 億円."""
    if df.empty:
        return

    chart_df = pd.DataFrame({
        "CALL GEX": df["call_gex"].values * _MOVE / _OKU,
        "PUT GEX": df["put_gex"].values * _MOVE / _OKU,
    }, index=df["strike"].apply(lambda x: f"{int(x):,}"))

    st.bar_chart(
        chart_df,
        color=["#2ecc71", "#e74c3c"],
        height=450,
    )

    st.caption(f"原資産: {spot:,.0f}  |  単位: 億円 / 原資産100円変動  |  "
               f"緑=CALL (正ガンマ), 赤=PUT (負ガンマ)")


def _render_gex_table(df: pd.DataFrame, spot: float) -> None:
    """Show detailed GEX values per strike (per 100 yen move)."""
    display = df.copy()
    display["call_gex"] = (display["call_gex"] * _MOVE / _OKU).round(1)
    display["put_gex"] = (display["put_gex"] * _MOVE / _OKU).round(1)
    display["net_gex"] = (display["net_gex"] * _MOVE / _OKU).round(1)
    # Futures equivalent per strike
    display["futures"] = (display["net_gex"] * _OKU / (spot * _FUTURES_MULT)).round(0).astype(int)
    display = display.rename(columns={
        "strike": "行使価格",
        "call_gex": "CALL(億/100円)",
        "put_gex": "PUT(億/100円)",
        "net_gex": "Net(億/100円)",
        "futures": "先物換算(枚)",
    })
    display["行使価格"] = display["行使価格"].astype(int)

    # Filter to strikes with meaningful GEX
    display = display[
        (display["CALL(億/100円)"].abs() >= 0.1) |
        (display["PUT(億/100円)"].abs() >= 0.1)
    ]

    st.dataframe(
        display,
        use_container_width=True,
        hide_index=True,
        height=min(len(display) * 35 + 40, 600),
    )


def _render_gex_3d_surface(
    strikes: list[int],
    put_oi: dict[int, int],
    call_oi: dict[int, int],
    spot: float,
    sq: date,
    as_of: date,
    sigma: float,
) -> None:
    """Render GEX curve: X=spot range, Y=total Net GEX (all strikes summed)."""
    st.subheader("GEXカーブ (原資産変動シミュレーション)")

    spot_range = st.slider(
        "原資産レンジ (±円)", 1000, 5000, 3000, step=500,
        key="gex_3d_range",
    )

    spots, strike_arr, surface = calc_gex_surface(
        strikes=strikes,
        put_oi=put_oi,
        call_oi=call_oi,
        spot_center=spot,
        spot_range=float(spot_range),
        spot_step=100.0,
        expiry_date=sq,
        as_of=as_of,
        sigma=sigma,
    )

    # Sum across all strikes → total net GEX per spot level
    total_gex = surface.sum(axis=1) * _MOVE / _OKU  # 億円/100円

    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=spots,
        y=total_gex,
        mode="lines",
        line=dict(color="#2ecc71", width=2),
        fill="tozeroy",
        name="Net GEX",
    ))

    # Current spot marker
    fig.add_vline(x=spot, line_dash="dash", line_color="blue",
                  annotation_text=f"現在値 {spot:,.0f}")

    # Zero line
    fig.add_hline(y=0, line_color="gray", line_width=0.5)

    fig.update_layout(
        xaxis_title="原資産価格",
        yaxis_title="Net GEX (億円/100円変動)",
        height=450,
        margin=dict(l=0, r=0, t=30, b=0),
    )

    st.plotly_chart(fig, use_container_width=True)
    st.caption("原資産が動いたときの全strikeの合算Net GEX  |  "
               "正=ディーラーロングガンマ(安定化), 負=ショートガンマ(不安定化)")
