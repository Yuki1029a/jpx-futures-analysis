"""Option Pain (Max Pain) chart rendering.

Computes the settlement price that minimizes total option payout
(i.e., maximum pain for option holders / minimum payout for writers).
"""
from __future__ import annotations

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import date

from models import OptionStrikeRow, WeekDefinition


def render_option_pain_section(
    all_month_rows: dict[str, list[OptionStrikeRow]],
    week: WeekDefinition,
) -> None:
    """Render option pain charts for multiple contract months.

    Args:
        all_month_rows: {contract_month: [OptionStrikeRow, ...]}
        week: Current week definition.
    """
    st.subheader("オプションペイン分析")

    if not all_month_rows:
        st.info("オプションデータなし")
        return

    for cm in sorted(all_month_rows.keys()):
        rows = all_month_rows[cm]
        if not rows:
            continue
        _render_single_pain(rows, week, cm)


def _format_cm(cm: str) -> str:
    if not cm:
        return "-"
    return f"20{cm[:2]}年{cm[2:]}月限"


def _render_single_pain(
    rows: list[OptionStrikeRow],
    week: WeekDefinition,
    contract_month: str,
) -> None:
    """Compute and render option pain chart for one contract month."""
    # Extract latest OI per strike
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

    # Filter to strikes with meaningful OI for cleaner chart
    total_oi = sum(put_oi.values()) + sum(call_oi.values())
    if total_oi == 0:
        return

    # Calculate option pain for each possible settlement price
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
    max_pain_value = total_pain[max_pain_idx]

    # Scale to 億円 (x1000 multiplier for NK225 options)
    OKU = 1e8
    MULT = 1000  # NK225 option multiplier
    scale = MULT / OKU
    call_pain_oku = [v * scale for v in call_pain]
    put_pain_oku = [v * scale for v in put_pain]
    total_pain_oku = [v * scale for v in total_pain]

    # Build chart
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

    # Max pain annotation
    fig.add_vline(
        x=max_pain_idx,
        line_dash="dash",
        line_color="green",
        line_width=2,
    )
    fig.add_annotation(
        x=f"{max_pain_strike:,}",
        y=max(total_pain_oku) * 0.95,
        text=f"Max Pain: {max_pain_strike:,}",
        showarrow=True,
        arrowhead=2,
        arrowcolor="green",
        font=dict(color="green", size=12),
    )

    dow_jp = ["月", "火", "水", "木", "金", "土", "日"]
    date_str = f"{latest_date.strftime('%m/%d')}({dow_jp[latest_date.weekday()]})"

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

    # Summary metrics
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
