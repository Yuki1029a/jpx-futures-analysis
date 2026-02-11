"""GEX (Gamma Exposure) chart rendering."""
from __future__ import annotations

import streamlit as st
import pandas as pd
from datetime import date

from models import OptionStrikeRow, WeekDefinition
from utils.gex import calc_gex_profile, get_sq_date

_DOW_JP = ["月", "火", "水", "木", "金", "土", "日"]
_OKU = 1e8  # 1億円


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
    st.caption(f"建玉基準日: {as_of.strftime('%Y/%m/%d')}({dow})")

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

    # --- Summary metrics ---
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.metric("Net GEX", f"{profile.total_net_gex / _OKU:+,.1f} 億円")
    with m2:
        st.metric("CALL GEX", f"{profile.total_call_gex / _OKU:+,.1f} 億円")
    with m3:
        st.metric("PUT GEX", f"{profile.total_put_gex / _OKU:+,.1f} 億円")
    with m4:
        if profile.flip_point:
            st.metric("フリップ", f"{profile.flip_point:,.0f}")
        else:
            st.metric("フリップ", "N/A")

    # --- Chart ---
    _render_gex_bar_chart(profile.df, spot)

    # --- Detail table ---
    with st.expander("GEX詳細テーブル"):
        _render_gex_table(profile.df)


def _extract_latest_oi(
    rows: list[OptionStrikeRow],
) -> tuple[date, dict[int, int], dict[int, int], set[int]]:
    """Extract the latest available daily OI across all strikes.

    Returns (as_of_date, put_oi_dict, call_oi_dict, all_strikes).
    """
    # Find the latest date that has OI data
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
    """Render GEX bar chart using st.bar_chart."""
    if df.empty:
        return

    # Convert to 億円 for display
    chart_df = pd.DataFrame({
        "CALL GEX": df["call_gex"].values / _OKU,
        "PUT GEX": df["put_gex"].values / _OKU,
    }, index=df["strike"].apply(lambda x: f"{int(x):,}"))

    st.bar_chart(
        chart_df,
        color=["#2ecc71", "#e74c3c"],
        height=450,
    )

    # Show spot line info
    st.caption(f"原資産: {spot:,.0f}  |  緑=CALL GEX (正), 赤=PUT GEX (負)")


def _render_gex_table(df: pd.DataFrame) -> None:
    """Show detailed GEX values per strike."""
    display = df.copy()
    display["call_gex"] = (display["call_gex"] / _OKU).round(3)
    display["put_gex"] = (display["put_gex"] / _OKU).round(3)
    display["net_gex"] = (display["net_gex"] / _OKU).round(3)
    display.columns = ["行使価格", "CALL GEX(億)", "PUT GEX(億)", "Net GEX(億)"]
    display["行使価格"] = display["行使価格"].astype(int)

    st.dataframe(
        display,
        use_container_width=True,
        hide_index=True,
        height=min(len(display) * 35 + 40, 600),
    )
