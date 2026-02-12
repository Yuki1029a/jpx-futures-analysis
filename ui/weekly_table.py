"""Main weekly table visualization component."""
from __future__ import annotations

import streamlit as st
import pandas as pd
import numpy as np
from models import WeeklyParticipantRow, WeekDefinition, DailyFuturesOI
import config


# Day-of-week labels in Japanese
_DOW_JP = ["月", "火", "水", "木", "金", "土", "日"]


def render_weekly_table(
    rows: list[WeeklyParticipantRow],
    week: WeekDefinition,
    product: str,
    contract_month: str,
    show_oi: bool = True,
    tab_label: str = "",
    stats_20d: dict | None = None,
    daily_futures_oi: dict | None = None,
) -> None:
    """Render the main weekly analysis table.

    Args:
        show_oi: If True, include OI columns (前週L/S/Net, 今週L/S/Net, 増減, 推定買/売, 方向).
                 If False, show only daily volumes and weekly total (for session-specific tabs).
        tab_label: Label shown in subheader for session context.
        stats_20d: {participant_id: (avg, max)} - 20-day stats passed separately to avoid mutating cached rows.
        daily_futures_oi: {date: DailyFuturesOI} - aggregate daily OI balance per trading day.
    """
    cm_label = f"20{contract_month[:2]}年{contract_month[2:]}月限" if contract_month else ""
    title = f"{config.PRODUCT_DISPLAY_NAMES.get(product, product)} {cm_label}  ({week.label})"
    if tab_label and tab_label != "全セッション合計":
        title += f"  [{tab_label}]"
    st.subheader(title)

    if not rows:
        st.warning("選択された条件のデータがありません。")
        return

    # Daily futures OI bar above the table
    if daily_futures_oi:
        _render_daily_oi_bar(week, daily_futures_oi)

    df = _build_display_dataframe(rows, week, show_oi, stats_20d)
    styled = _apply_table_styling(df, week, show_oi)

    st.dataframe(
        styled,
        use_container_width=True,
        height=min(len(rows) * 35 + 60, 900),
    )

    if show_oi:
        _render_summary_stats(rows)


def _build_display_dataframe(
    rows: list[WeeklyParticipantRow],
    week: WeekDefinition,
    show_oi: bool,
    stats_20d: dict | None = None,
) -> pd.DataFrame:
    """Build the display DataFrame with numeric columns for proper sorting."""
    records = []
    for row in rows:
        rec = {"参加者": row.participant_name}

        if show_oi:
            rec["前週L"] = row.start_oi_long
            rec["前週S"] = row.start_oi_short
            rec["前週Net"] = row.start_oi_net

        # Daily volume columns
        weekly_total = 0.0
        for td in week.trading_days:
            dow = _DOW_JP[td.weekday()]
            col_name = f"{td.strftime('%m/%d')}({dow})"
            vol = row.daily_volumes.get(td)
            rec[col_name] = vol if vol else None
            if vol:
                weekly_total += vol

        rec["週間計"] = weekly_total if weekly_total > 0 else None

        # 20-day stats from separate dict (not from row object)
        avg_20d = None
        max_20d = None
        if stats_20d and row.participant_id in stats_20d:
            avg_20d, max_20d = stats_20d[row.participant_id]
        rec["20日平均"] = round(avg_20d) if avg_20d is not None else None
        rec["20日最大"] = round(max_20d) if max_20d is not None else None

        if show_oi:
            rec["今週L"] = row.end_oi_long
            rec["今週S"] = row.end_oi_short
            rec["今週Net"] = row.end_oi_net
            rec["増減"] = row.oi_net_change

            # Estimate buy/sell breakdown
            est_buy = None
            est_sell = None
            if row.oi_net_change is not None and weekly_total > 0:
                est_buy = (weekly_total + row.oi_net_change) / 2
                est_sell = (weekly_total - row.oi_net_change) / 2
            rec["推定買"] = est_buy
            rec["推定売"] = est_sell
            rec["方向"] = _direction_label(row.inferred_direction)

        records.append(rec)

    return pd.DataFrame(records)


def _apply_table_styling(df: pd.DataFrame, week: WeekDefinition, show_oi: bool):
    """Apply conditional formatting and number formatting."""

    # Identify column groups
    day_cols = [f"{td.strftime('%m/%d')}({_DOW_JP[td.weekday()]})" for td in week.trading_days]
    int_cols = list(day_cols) + ["週間計", "20日平均", "20日最大"]

    net_cols = []
    if show_oi:
        oi_cols = ["前週L", "前週S", "今週L", "今週S"]
        net_cols = ["前週Net", "今週Net", "増減"]
        est_cols = ["推定買", "推定売"]
        int_cols += oi_cols + net_cols + est_cols

    # Color functions
    def _color_signed(val):
        if pd.isna(val):
            return ""
        try:
            n = float(val)
            if n > 0:
                return "background-color: #c6efce; color: #006100"
            elif n < 0:
                return "background-color: #ffc7ce; color: #9c0006"
        except (ValueError, TypeError):
            pass
        return ""

    def _color_direction(val):
        if val == "BUY":
            return "background-color: #c6efce; color: #006100; font-weight: bold"
        elif val == "SELL":
            return "background-color: #ffc7ce; color: #9c0006; font-weight: bold"
        return ""

    styled = df.style

    # Apply sign-based coloring to net/change columns
    for col in net_cols:
        if col in df.columns:
            styled = styled.map(_color_signed, subset=[col])

    # Color direction column
    if show_oi and "方向" in df.columns:
        styled = styled.map(_color_direction, subset=["方向"])

    # Number formatting
    fmt_int = lambda v: f"{int(v):,}" if pd.notna(v) else "-"
    fmt_signed = lambda v: f"{int(v):+,}" if pd.notna(v) else "-"

    for col in df.columns:
        if col in net_cols:
            styled = styled.format(fmt_signed, subset=[col])
        elif col in int_cols:
            styled = styled.format(fmt_int, subset=[col])

    return styled


def _render_summary_stats(rows: list[WeeklyParticipantRow]) -> None:
    """Render summary metrics below the table."""
    st.markdown("---")
    cols = st.columns(4)

    buyers = sum(1 for r in rows if r.inferred_direction == "BUY")
    sellers = sum(1 for r in rows if r.inferred_direction == "SELL")
    total_vol = sum(sum(r.daily_volumes.values()) for r in rows)
    total_net = sum(r.oi_net_change for r in rows if r.oi_net_change is not None)

    with cols[0]:
        st.metric("買い方", buyers)
    with cols[1]:
        st.metric("売り方", sellers)
    with cols[2]:
        st.metric("週間出来高計", f"{int(total_vol):,}")
    with cols[3]:
        delta_color = "normal" if total_net >= 0 else "inverse"
        st.metric("全体Net増減", f"{int(total_net):+,}", delta_color=delta_color)


def _render_daily_oi_bar(
    week: WeekDefinition,
    daily_futures_oi: dict,
) -> None:
    """Render daily OI as a progress bar row + change metrics aligned with trading days."""
    days = week.trading_days
    if not days:
        return

    # Collect all OI values to determine max for progress bar scaling
    oi_values: list[int] = []
    for td in days:
        rec = daily_futures_oi.get(td)
        if rec:
            oi_values.append(rec.current_oi)
            if rec.previous_oi > 0:
                oi_values.append(rec.previous_oi)

    if not oi_values:
        return

    max_oi = max(oi_values) if oi_values else 1

    # OI bar row (ProgressColumn)
    rec_oi: dict[str, int | None] = {}
    for td in days:
        dow = _DOW_JP[td.weekday()]
        col = f"{td.strftime('%m/%d')}({dow})"
        oi_rec = daily_futures_oi.get(td)
        rec_oi[col] = oi_rec.current_oi if oi_rec else None

    df_oi = pd.DataFrame([rec_oi])
    day_cols = list(rec_oi.keys())

    col_config = {}
    for col in day_cols:
        col_config[col] = st.column_config.ProgressColumn(
            col,
            min_value=0,
            max_value=int(max_oi * 1.05),
            format=" ",
        )

    st.caption("建玉残高")
    st.dataframe(
        df_oi,
        column_config=col_config,
        use_container_width=True,
        hide_index=True,
        height=50,
    )

    # Change row as colored metrics
    cols = st.columns(len(days))
    for col_ui, td in zip(cols, days):
        oi_rec = daily_futures_oi.get(td)
        with col_ui:
            if oi_rec:
                chg = oi_rec.net_change
                color = "#006100" if chg >= 0 else "#9c0006"
                st.markdown(
                    f"<div style='text-align:center;font-size:0.85em;'>"
                    f"<b>{oi_rec.current_oi:,}</b><br>"
                    f"<span style='color:{color}'>{chg:+,}</span>"
                    f"</div>",
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    "<div style='text-align:center;font-size:0.85em;'>-</div>",
                    unsafe_allow_html=True,
                )


def _direction_label(direction) -> str:
    if direction == "BUY":
        return "BUY"
    elif direction == "SELL":
        return "SELL"
    elif direction == "NEUTRAL":
        return "-"
    return ""
