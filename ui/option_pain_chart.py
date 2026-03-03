"""Option Pain (Max Pain) chart rendering.

Computes the settlement price that minimizes total option payout
(i.e., maximum pain for option holders / minimum payout for writers).
Includes time-series view of Max Pain vs NK225 closing prices
and 3D pain surface for selected contract month.
"""
from __future__ import annotations

import streamlit as st
import numpy as np
import plotly.graph_objects as go
from datetime import date, datetime
from collections import defaultdict

from models import OptionStrikeRow, WeekDefinition, DailyOIBalance

_DOW_JP = ["月", "火", "水", "木", "金", "土", "日"]


def render_option_pain_section(
    all_month_rows: dict[str, list[OptionStrikeRow]],
    week: WeekDefinition,
) -> None:
    """Render option pain charts for multiple contract months."""
    st.subheader("オプションペイン分析")

    if not all_month_rows:
        st.info("オプションデータなし")
        return

    # Time series chart (all months)
    _render_maxpain_timeseries()

    st.markdown("---")

    # 3D pain surface (selected month)
    _render_pain_3d_section(all_month_rows, week)

    st.markdown("---")

    # Per-month pain charts
    for cm in sorted(all_month_rows.keys()):
        rows = all_month_rows[cm]
        if not rows:
            continue
        _render_single_pain(rows, week, cm)


def _format_cm(cm: str) -> str:
    if not cm:
        return "-"
    return f"20{cm[:2]}年{cm[2:]}月限"


# =====================================================================
# Max Pain time series
# =====================================================================

@st.cache_data(ttl=600, show_spinner=False)
def _load_maxpain_timeseries_data() -> tuple[
    dict[date, dict[str, int | None]],
    list[str],
    list[date],
    list[float],
]:
    """Load and compute max pain time series data (cached 10 min)."""
    from data import fetcher
    from data.aggregator import _load_daily_oi_for_date

    months = fetcher.get_available_volume_months()
    all_dates: list[date] = []
    for m in months[:4]:
        try:
            entries = fetcher.get_volume_index(m)
        except Exception:
            continue
        for entry in entries:
            d = datetime.strptime(entry["TradeDate"], "%Y%m%d").date()
            all_dates.append(d)
    all_dates = sorted(set(all_dates))[-30:]

    if not all_dates:
        return {}, [], [], []

    maxpain_data: dict[date, dict[str, int | None]] = {}
    all_cms: set[str] = set()

    for td in all_dates:
        try:
            records = _load_daily_oi_for_date(td)
        except Exception:
            continue
        if not records:
            continue
        mp = _compute_max_pain(records)
        if mp:
            maxpain_data[td] = mp
            all_cms.update(mp.keys())

    contract_months = []
    for cm in sorted(all_cms):
        count = sum(1 for d in maxpain_data if maxpain_data[d].get(cm) is not None)
        if count >= 3:
            contract_months.append(cm)

    nk_dates: list[date] = []
    nk_prices: list[float] = []
    try:
        import yfinance as yf
        hist = yf.Ticker("^N225").history(period="3mo")
        if not hist.empty:
            nk_dates = [d.date() for d in hist.index]
            nk_prices = hist["Close"].tolist()
    except Exception:
        pass

    return maxpain_data, contract_months, nk_dates, nk_prices


def _compute_max_pain(records: list[DailyOIBalance]) -> dict[str, int | None]:
    """Compute Max Pain for each contract month from DailyOIBalance records."""
    cm_data: dict[str, dict[int, dict[str, int]]] = defaultdict(
        lambda: defaultdict(lambda: {"CALL": 0, "PUT": 0})
    )
    for r in records:
        if r.current_oi > 0:
            cm_data[r.contract_month][r.strike_price][r.option_type] += r.current_oi

    results: dict[str, int | None] = {}
    for cm, strike_map in cm_data.items():
        if not strike_map:
            results[cm] = None
            continue

        strikes = sorted(strike_map.keys())
        call_oi = {k: strike_map[k]["CALL"] for k in strikes}
        put_oi = {k: strike_map[k]["PUT"] for k in strikes}

        total_oi = sum(call_oi.values()) + sum(put_oi.values())
        if total_oi == 0:
            results[cm] = None
            continue

        best_strike = None
        min_pain = float("inf")
        for S in strikes:
            pain = sum(max(0, S - K) * oi for K, oi in call_oi.items())
            pain += sum(max(0, K - S) * oi for K, oi in put_oi.items())
            if pain < min_pain:
                min_pain = pain
                best_strike = S

        results[cm] = best_strike

    return results


def _render_maxpain_timeseries() -> None:
    """Render the Max Pain time series chart with NK225 overlay."""
    with st.spinner("Max Pain時系列データ読み込み中..."):
        maxpain_data, contract_months, nk_dates, nk_prices = (
            _load_maxpain_timeseries_data()
        )

    if not maxpain_data or not contract_months:
        st.info("Max Pain時系列データなし")
        return

    fig = go.Figure()

    if nk_dates:
        fig.add_trace(go.Scatter(
            x=nk_dates, y=nk_prices,
            mode="lines", name="NK225終値",
            line=dict(color="black", width=2.5),
        ))

    colors = ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd", "#8c564b"]
    sorted_dates = sorted(maxpain_data.keys())

    for i, cm in enumerate(contract_months):
        cm_dates = [d for d in sorted_dates if maxpain_data[d].get(cm) is not None]
        cm_values = [maxpain_data[d][cm] for d in cm_dates]
        color = colors[i % len(colors)]
        fig.add_trace(go.Scatter(
            x=cm_dates, y=cm_values,
            mode="lines+markers",
            name=f"MaxPain {_format_cm(cm)}",
            line=dict(color=color, width=1.5, dash="dot"),
            marker=dict(size=6, color=color),
        ))

    fig.update_layout(
        title="Max Pain 時系列 vs NK225終値",
        xaxis=dict(title="日付", type="date"),
        yaxis=dict(title="価格", tickformat=","),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
            bgcolor="rgba(255,255,255,0.8)",
        ),
        hovermode="x unified",
        template="plotly_white",
        height=500,
        margin=dict(l=0, r=0, t=40, b=0),
    )

    st.plotly_chart(fig, use_container_width=True)

    # Summary metrics
    nk_lookup = dict(zip(nk_dates, nk_prices))
    latest_date = sorted_dates[-1] if sorted_dates else None
    if latest_date:
        cols = st.columns(len(contract_months) + 1)
        nk_close = nk_lookup.get(latest_date)
        with cols[0]:
            if nk_close:
                st.metric("NK225終値", f"{nk_close:,.0f}")
        for i, cm in enumerate(contract_months):
            mp = maxpain_data[latest_date].get(cm)
            with cols[i + 1]:
                if mp:
                    delta = None
                    if nk_close:
                        delta = f"{mp - nk_close:+,.0f}"
                    st.metric(f"MaxPain {cm[:2]}/{cm[2:]}", f"{mp:,}", delta=delta)


# =====================================================================
# 3D Pain Surface (selected contract month, last 5 days)
# =====================================================================

def _render_pain_3d_section(
    all_month_rows: dict[str, list[OptionStrikeRow]],
    week: WeekDefinition,
) -> None:
    """Render 3D option pain surface for a selected contract month."""
    st.subheader("ペインプロファイル 3D (直近5営業日)")

    sorted_cms = sorted(all_month_rows.keys())
    if not sorted_cms:
        return

    selected_cm = st.selectbox(
        "限月選択",
        sorted_cms,
        format_func=_format_cm,
        key="pain_3d_cm",
    )

    rows = all_month_rows.get(selected_cm, [])
    if not rows:
        st.info("データなし")
        return

    # Collect all dates with OI data
    all_dates: set[date] = set()
    for r in rows:
        all_dates.update(r.put_daily_oi.keys())
        all_dates.update(r.call_daily_oi.keys())

    if not all_dates:
        st.info("建玉データなし")
        return

    sorted_dates = sorted(all_dates)[-5:]  # last 5 days

    # Collect strikes with meaningful OI on any of the dates
    active_strikes: set[int] = set()
    for r in rows:
        for td in sorted_dates:
            if r.put_daily_oi.get(td, 0) > 0 or r.call_daily_oi.get(td, 0) > 0:
                active_strikes.add(r.strike_price)
    strikes = sorted(active_strikes)

    if len(strikes) < 3:
        st.info("有効行使価格が不足")
        return

    # Filter strikes to a reasonable range around the pain center
    # (too many strikes makes the 3D chart unreadable)
    # Compute rough center from latest day
    latest = sorted_dates[-1]
    put_oi_latest = {r.strike_price: r.put_daily_oi.get(latest, 0) for r in rows}
    call_oi_latest = {r.strike_price: r.call_daily_oi.get(latest, 0) for r in rows}
    total_oi_by_strike = {
        s: put_oi_latest.get(s, 0) + call_oi_latest.get(s, 0) for s in strikes
    }
    # Weighted center
    total_w = sum(total_oi_by_strike.values())
    if total_w > 0:
        center = sum(s * oi for s, oi in total_oi_by_strike.items()) / total_w
    else:
        center = strikes[len(strikes) // 2]

    # Keep strikes within +/- 8000 of center (covers meaningful range)
    filtered_strikes = [s for s in strikes if abs(s - center) <= 8000]
    if len(filtered_strikes) < 3:
        filtered_strikes = strikes  # fallback

    OKU = 1e8
    MULT = 1000
    scale = MULT / OKU

    # Build pain matrix: rows=dates, cols=strikes
    z_data = []
    date_labels = []
    max_pain_per_day = []

    for td in sorted_dates:
        put_oi = {}
        call_oi = {}
        for r in rows:
            p = r.put_daily_oi.get(td, 0)
            c = r.call_daily_oi.get(td, 0)
            if p > 0:
                put_oi[r.strike_price] = p
            if c > 0:
                call_oi[r.strike_price] = c

        pain_row = []
        min_pain = float("inf")
        mp_strike = None
        for S in filtered_strikes:
            pain = sum(max(0, S - K) * oi for K, oi in call_oi.items())
            pain += sum(max(0, K - S) * oi for K, oi in put_oi.items())
            pain_scaled = pain * scale
            pain_row.append(pain_scaled)
            if pain < min_pain:
                min_pain = pain
                mp_strike = S

        z_data.append(pain_row)
        dow = _DOW_JP[td.weekday()]
        date_labels.append(f"{td.strftime('%m/%d')}({dow})")
        max_pain_per_day.append(mp_strike)

    z_array = np.array(z_data)

    # 3D surface chart
    fig = go.Figure(data=[go.Surface(
        x=filtered_strikes,
        y=list(range(len(sorted_dates))),
        z=z_array,
        colorscale="YlOrRd",
        colorbar=dict(title="払出額(億円)"),
        hovertemplate=(
            "行使価格: %{x:,}<br>"
            "日付: %{customdata}<br>"
            "払出額: %{z:,.0f}億円<extra></extra>"
        ),
        customdata=np.array([[dl] * len(filtered_strikes) for dl in date_labels]),
    )])

    # Add max pain markers
    mp_x = max_pain_per_day
    mp_y = list(range(len(sorted_dates)))
    mp_z = [z_data[i][filtered_strikes.index(mp)] if mp in filtered_strikes else 0
            for i, mp in enumerate(max_pain_per_day)]

    fig.add_trace(go.Scatter3d(
        x=mp_x, y=mp_y, z=mp_z,
        mode="markers+text",
        marker=dict(size=6, color="green", symbol="diamond"),
        text=[f"MP:{s:,}" for s in max_pain_per_day],
        textposition="top center",
        textfont=dict(size=10, color="green"),
        name="Max Pain",
        hovertemplate="Max Pain: %{x:,}<br>%{text}<extra></extra>",
    ))

    fig.update_layout(
        title=f"Option Pain 3D Surface - {_format_cm(selected_cm)}",
        scene=dict(
            xaxis=dict(title="行使価格", tickformat=","),
            yaxis=dict(
                title="日付",
                tickvals=list(range(len(sorted_dates))),
                ticktext=date_labels,
            ),
            zaxis=dict(title="払出額 (億円)", tickformat=","),
            camera=dict(eye=dict(x=1.5, y=-1.8, z=0.8)),
        ),
        height=600,
        margin=dict(l=0, r=0, t=40, b=0),
    )

    st.plotly_chart(fig, use_container_width=True)

    # Max Pain per day summary
    cols = st.columns(len(sorted_dates))
    for i, (td, mp) in enumerate(zip(sorted_dates, max_pain_per_day)):
        dow = _DOW_JP[td.weekday()]
        with cols[i]:
            prev_mp = max_pain_per_day[i - 1] if i > 0 else None
            delta = None
            if prev_mp and mp:
                diff = mp - prev_mp
                if diff != 0:
                    delta = f"{diff:+,}"
            st.metric(
                f"{td.strftime('%m/%d')}({dow})",
                f"{mp:,}" if mp else "N/A",
                delta=delta,
            )


# =====================================================================
# Single contract month pain chart
# =====================================================================

def _render_single_pain(
    rows: list[OptionStrikeRow],
    week: WeekDefinition,
    contract_month: str,
) -> None:
    """Compute and render option pain chart for one contract month."""
    latest_date = _find_latest_oi_date(rows)
    if latest_date is None:
        return

    strikes: list[int] = []
    put_oi: dict[int, int] = {}
    call_oi: dict[int, int] = {}

    for row in rows:
        p = row.put_daily_oi.get(latest_date, 0)
        c = row.call_daily_oi.get(latest_date, 0)
        if p > 0 or c > 0:
            strikes.append(row.strike_price)
            put_oi[row.strike_price] = p
            call_oi[row.strike_price] = c

    if len(strikes) < 3:
        return

    strikes.sort()
    total_oi = sum(put_oi.values()) + sum(call_oi.values())
    if total_oi == 0:
        return

    settlement_prices = strikes
    call_pain = []
    put_pain = []
    total_pain = []

    for S in settlement_prices:
        cp = sum(max(0, S - K) * call_oi.get(K, 0) for K in strikes)
        pp = sum(max(0, K - S) * put_oi.get(K, 0) for K in strikes)
        call_pain.append(cp)
        put_pain.append(pp)
        total_pain.append(cp + pp)

    max_pain_idx = total_pain.index(min(total_pain))
    max_pain_strike = settlement_prices[max_pain_idx]

    OKU = 1e8
    MULT = 1000
    scale = MULT / OKU
    call_pain_oku = [v * scale for v in call_pain]
    put_pain_oku = [v * scale for v in put_pain]
    total_pain_oku = [v * scale for v in total_pain]

    fig = go.Figure()

    fig.add_trace(go.Bar(
        name="CALL Payout",
        x=[f"{s:,}" for s in settlement_prices],
        y=call_pain_oku,
        marker_color="rgba(74, 144, 217, 0.7)",
    ))
    fig.add_trace(go.Bar(
        name="PUT Payout",
        x=[f"{s:,}" for s in settlement_prices],
        y=put_pain_oku,
        marker_color="rgba(217, 74, 74, 0.7)",
    ))
    fig.add_trace(go.Scatter(
        name="Total",
        x=[f"{s:,}" for s in settlement_prices],
        y=total_pain_oku,
        mode="lines+markers",
        line=dict(color="black", width=2),
        marker=dict(size=3),
    ))

    fig.add_vline(x=max_pain_idx, line_dash="dash", line_color="green", line_width=2)
    fig.add_annotation(
        x=f"{max_pain_strike:,}",
        y=max(total_pain_oku) * 0.95,
        text=f"Max Pain: {max_pain_strike:,}",
        showarrow=True, arrowhead=2, arrowcolor="green",
        font=dict(color="green", size=12),
    )

    dow = _DOW_JP[latest_date.weekday()]
    date_str = f"{latest_date.strftime('%m/%d')}({dow})"

    fig.update_layout(
        title=f"Option Pain - {_format_cm(contract_month)}  (建玉基準: {date_str})",
        xaxis_title="SQ決済価格",
        yaxis_title="オプション払出額 (億円)",
        barmode="stack",
        height=420,
        margin=dict(l=0, r=0, t=40, b=0),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )

    st.plotly_chart(fig, use_container_width=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("Max Pain", f"{max_pain_strike:,}")
    with c2:
        st.metric("PUT OI計", f"{sum(put_oi.values()):,}")
    with c3:
        st.metric("CALL OI計", f"{sum(call_oi.values()):,}")

    st.markdown("---")


def _find_latest_oi_date(rows: list[OptionStrikeRow]) -> date | None:
    """Find the latest date with OI data across all strikes."""
    all_dates: set[date] = set()
    for r in rows:
        all_dates.update(r.put_daily_oi.keys())
        all_dates.update(r.call_daily_oi.keys())
    return max(all_dates) if all_dates else None
