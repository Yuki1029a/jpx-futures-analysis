"""Option strike-price table with single-row layout, OI bar charts, and breakdown panel.

Each strike = one DataFrame row showing:
  - Participant-filtered daily volumes (NumberColumn)
  - Daily OI bar charts [current_oi, previous_oi] (BarChartColumn)
  - Weekly OI long/short

Layout: left = main table (clickable, sortable), right = detail panel.
"""
from __future__ import annotations

import streamlit as st
import pandas as pd
from datetime import date
from models import OptionStrikeRow, WeekDefinition

_DOW_JP = ["月", "火", "水", "木", "金", "土", "日"]


def render_option_strike_table(
    rows: list[OptionStrikeRow],
    week: WeekDefinition,
    tab_label: str = "",
) -> None:
    """Render option table (left) + detail panel (right)."""
    title = f"日経225オプション ({week.label})"
    if tab_label and tab_label != "全セッション合計":
        title += f"  [{tab_label}]"
    st.subheader(title)

    if not rows:
        st.warning("オプションデータがありません。")
        return

    # OI bar toggle
    show_oi_bars = st.checkbox("建玉バー表示", value=True, key=f"oi_bar_{tab_label}")

    # Compute max OI for consistent bar scaling
    max_oi = 1
    if show_oi_bars:
        for row in rows:
            for v in row.put_daily_oi.values():
                if v and v > max_oi:
                    max_oi = v
            for v in row.call_daily_oi.values():
                if v and v > max_oi:
                    max_oi = v

    df = _build_display_dataframe(rows, week, show_oi_bars)
    put_day_cols = [_day_col(td, "P") for td in week.trading_days]
    call_day_cols = [_day_col(td, "C") for td in reversed(week.trading_days)]
    put_oi_cols = [_oi_col(td, "P") for td in week.trading_days] if show_oi_bars else []
    call_oi_cols = [_oi_col(td, "C") for td in reversed(week.trading_days)] if show_oi_bars else []

    # Column config
    col_config = _build_column_config(
        df, put_day_cols, call_day_cols, put_oi_cols, call_oi_cols, max_oi,
    )

    # Hide metadata column
    col_config["_strike_idx"] = None

    # Layout: table (left 75%) | detail (right 25%)
    left_col, right_col = st.columns([3, 1])

    with left_col:
        table_key = f"opt_table_{tab_label}"
        event = st.dataframe(
            df,
            use_container_width=True,
            height=min(len(df) * 35 + 60, 900),
            on_select="rerun",
            selection_mode="single-cell",
            key=table_key,
            column_config=col_config,
        )

    # Parse cell selection
    selected_strike_idx = None
    selected_date = None
    selected_type = None

    all_put_cols = set(put_day_cols) | set(put_oi_cols)
    all_call_cols = set(call_day_cols) | set(call_oi_cols)

    if event and event.selection and event.selection.cells:
        cell = event.selection.cells[0]
        df_row_idx = cell[0]
        col_name = cell[1]

        if 0 <= df_row_idx < len(df):
            strike_idx_val = df.iloc[df_row_idx].get("_strike_idx")
            if strike_idx_val is not None:
                selected_strike_idx = int(strike_idx_val)

            if col_name in all_put_cols:
                selected_type = "PUT"
                selected_date = _any_col_to_date(col_name, week)
            elif col_name in all_call_cols:
                selected_type = "CALL"
                selected_date = _any_col_to_date(col_name, week)
            elif col_name.startswith("P"):
                selected_type = "PUT"
            elif col_name.startswith("C"):
                selected_type = "CALL"

    with right_col:
        _render_detail_panel(
            rows, week,
            selected_strike_idx, selected_date, selected_type,
            tab_label,
        )

    _render_option_summary(rows)


# --- Helpers ---

def _day_col(td: date, prefix: str) -> str:
    dow = _DOW_JP[td.weekday()]
    return f"{prefix}{td.strftime('%m/%d')}({dow})"


def _oi_col(td: date, prefix: str) -> str:
    return f"{prefix}建{td.strftime('%d')}"


def _any_col_to_date(col_name: str, week: WeekDefinition) -> date | None:
    """Resolve a volume or OI column name to a date."""
    for td in week.trading_days:
        if (_day_col(td, "P") == col_name or _day_col(td, "C") == col_name
                or _oi_col(td, "P") == col_name or _oi_col(td, "C") == col_name):
            return td
    return None


def _build_column_config(
    df: pd.DataFrame,
    put_day_cols: list[str],
    call_day_cols: list[str],
    put_oi_cols: list[str],
    call_oi_cols: list[str],
    max_oi: int,
) -> dict:
    """Build column_config for formatting."""
    oi_cols = ["P前週L", "P前週S", "P今週L", "P今週S",
               "C前週L", "C前週S", "C今週L", "C今週S"]
    total_cols = ["P計", "C計"]
    all_num = put_day_cols + call_day_cols + oi_cols + total_cols

    col_config = {}
    for col in all_num:
        if col in df.columns:
            col_config[col] = st.column_config.NumberColumn(col, format="%d")

    col_config["行使価格"] = st.column_config.NumberColumn("行使価格", format="%d")

    # BarChartColumn for OI bars
    for col in put_oi_cols + call_oi_cols:
        if col in df.columns:
            col_config[col] = st.column_config.BarChartColumn(
                col,
                width="small",
                y_min=0,
                y_max=max_oi,
            )

    return col_config


def _build_display_dataframe(
    rows: list[OptionStrikeRow],
    week: WeekDefinition,
    show_oi_bars: bool = True,
) -> pd.DataFrame:
    """Build DataFrame with one row per strike.

    Returns DataFrame with _strike_idx metadata column.
    """
    records = []

    for idx, row in enumerate(rows):
        rec = _build_volume_row(row, week, show_oi_bars)
        rec["行使価格"] = row.strike_price
        rec["_strike_idx"] = idx
        records.append(rec)

    return pd.DataFrame(records)


def _build_volume_row(
    row: OptionStrikeRow,
    week: WeekDefinition,
    show_oi_bars: bool = True,
) -> dict:
    """Build the volume row for a strike, optionally with OI bar data."""
    rec = {}
    rec["P前週L"] = row.put_start_oi_long
    rec["P前週S"] = row.put_start_oi_short

    for td in week.trading_days:
        rec[_day_col(td, "P")] = row.put_daily_volumes.get(td) or None
        if show_oi_bars:
            curr = row.put_daily_oi.get(td)
            if curr is not None:
                chg = row.put_daily_oi_change.get(td, 0)
                prev = curr - chg if chg else curr
                rec[_oi_col(td, "P")] = [curr, prev]
            else:
                rec[_oi_col(td, "P")] = None

    rec["P計"] = row.put_week_total
    rec["P今週L"] = row.put_end_oi_long
    rec["P今週S"] = row.put_end_oi_short

    rec["C今週L"] = row.call_end_oi_long
    rec["C今週S"] = row.call_end_oi_short
    rec["C計"] = row.call_week_total

    for td in reversed(week.trading_days):
        rec[_day_col(td, "C")] = row.call_daily_volumes.get(td) or None
        if show_oi_bars:
            curr = row.call_daily_oi.get(td)
            if curr is not None:
                chg = row.call_daily_oi_change.get(td, 0)
                prev = curr - chg if chg else curr
                rec[_oi_col(td, "C")] = [curr, prev]
            else:
                rec[_oi_col(td, "C")] = None

    rec["C前週L"] = row.call_start_oi_long
    rec["C前週S"] = row.call_start_oi_short

    return rec


def _render_detail_panel(
    rows: list[OptionStrikeRow],
    week: WeekDefinition,
    strike_idx: int | None,
    selected_date: date | None,
    selected_type: str | None,
    tab_label: str,
) -> None:
    """Right panel: show participant breakdown + daily OI for selected cell."""
    st.markdown("**詳細パネル**")

    if strike_idx is None or strike_idx >= len(rows):
        st.caption("テーブルのセルをクリック")
        return

    target_row = rows[strike_idx]

    if selected_type is None:
        selected_type = "CALL"

    # If no specific date, let user pick
    if selected_date is None:
        st.markdown(f"**{selected_type} {target_row.strike_price:,}**")
        day_labels = [f"{td.strftime('%m/%d')}({_DOW_JP[td.weekday()]})"
                      for td in week.trading_days]
        prefix = f"bd_{tab_label}"
        day_choice = st.selectbox("日付", day_labels, key=f"{prefix}_day_r")
        if day_choice is None:
            return
        selected_date = week.trading_days[day_labels.index(day_choice)]

    dow = _DOW_JP[selected_date.weekday()]
    date_str = f"{selected_date.strftime('%m/%d')}({dow})"

    # 1) Participant breakdown
    _render_participant_breakdown(target_row, selected_type, selected_date, date_str)

    # 2) Daily OI balance (always shown below breakdown if data exists)
    _render_oi_detail(target_row, selected_type, selected_date)


def _render_participant_breakdown(
    row: OptionStrikeRow,
    option_type: str,
    td: date,
    date_str: str,
) -> None:
    """Show per-participant volume breakdown."""
    is_put = option_type == "PUT"
    breakdown = (row.put_daily_breakdown if is_put else row.call_daily_breakdown).get(td, [])

    header = f"{option_type} {row.strike_price:,}  {date_str}"

    if not breakdown:
        st.markdown(f"**{header}**")
        st.caption("出来高データなし")
        return

    total = sum(v for _, v in breakdown)
    st.markdown(f"**{header}**")
    st.markdown(f"出来高: **{int(total):,}**枚")

    bd_df = pd.DataFrame(breakdown, columns=["参加者", "枚数"])
    bd_df["枚数"] = bd_df["枚数"].astype(int)
    bd_df["構成比"] = (bd_df["枚数"] / total * 100).round(1).astype(str) + "%"
    st.dataframe(
        bd_df,
        use_container_width=True,
        hide_index=True,
        height=min(len(breakdown) * 35 + 40, 500),
    )


def _render_oi_detail(
    row: OptionStrikeRow,
    option_type: str,
    td: date,
) -> None:
    """Show daily OI balance detail below participant breakdown."""
    is_put = option_type == "PUT"
    oi = (row.put_daily_oi if is_put else row.call_daily_oi).get(td)
    chg = (row.put_daily_oi_change if is_put else row.call_daily_oi_change).get(td)

    if oi is None:
        return

    st.markdown("---")
    prev = oi - chg if chg is not None else None
    chg_display = f"+{chg:,}" if chg and chg > 0 else f"{chg:,}" if chg else "0"

    st.metric("建玉残高", f"{oi:,}", delta=chg_display)
    if prev is not None:
        st.caption(f"前日残高: {prev:,}")


def _render_option_summary(rows: list[OptionStrikeRow]) -> None:
    """Summary stats below the table."""
    st.markdown("---")
    cols = st.columns(4)

    total_put_vol = sum(r.put_week_total or 0 for r in rows)
    total_call_vol = sum(r.call_week_total or 0 for r in rows)
    active_strikes = sum(1 for r in rows
                         if (r.put_week_total or 0) > 0 or (r.call_week_total or 0) > 0)
    pcr = total_put_vol / total_call_vol if total_call_vol > 0 else 0

    with cols[0]:
        st.metric("PUT出来高計", f"{int(total_put_vol):,}")
    with cols[1]:
        st.metric("CALL出来高計", f"{int(total_call_vol):,}")
    with cols[2]:
        st.metric("P/C比率", f"{pcr:.2f}")
    with cols[3]:
        st.metric("有効行使価格数", active_strikes)
