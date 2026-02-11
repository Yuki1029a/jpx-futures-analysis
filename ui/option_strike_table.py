"""Option strike-price table with cell-click breakdown panel.

Layout: left = main table (clickable), right = participant breakdown.
PUT=light red header, CALL=light blue header via CSS injection.
Cell click on daily volume column triggers breakdown in right panel.
"""
from __future__ import annotations

import streamlit as st
import pandas as pd
from datetime import date
from models import OptionStrikeRow, WeekDefinition

_DOW_JP = ["月", "火", "水", "木", "金", "土", "日"]

# CSS to color-code PUT/CALL column headers and cells
_TABLE_CSS = """
<style>
/* Strike price column - bold grey */
div[data-testid="stDataFrame"] th:has(div[title="行使価格"]),
div[data-testid="stDataFrame"] th:has(div[title="行使価格"]) * {
    font-weight: bold !important;
}
</style>
"""


def render_option_strike_table(
    rows: list[OptionStrikeRow],
    week: WeekDefinition,
    tab_label: str = "",
) -> None:
    """Render option table (left) + breakdown panel (right)."""
    title = f"日経225オプション ({week.label})"
    if tab_label and tab_label != "全セッション合計":
        title += f"  [{tab_label}]"
    st.subheader(title)

    if not rows:
        st.warning("オプションデータがありません。")
        return

    df = _build_display_dataframe(rows, week)
    put_day_cols = [_day_col(td, "P") for td in week.trading_days]
    call_day_cols = [_day_col(td, "C") for td in reversed(week.trading_days)]

    # Column config for integer formatting
    col_config = _build_column_config(df, put_day_cols, call_day_cols)

    # Layout: table (left 75%) | breakdown (right 25%)
    left_col, right_col = st.columns([3, 1])

    with left_col:
        st.markdown(_TABLE_CSS, unsafe_allow_html=True)
        table_key = f"opt_table_{tab_label}"

        event = st.dataframe(
            df,
            use_container_width=True,
            height=min(len(rows) * 35 + 60, 900),
            on_select="rerun",
            selection_mode="single-cell",
            key=table_key,
            column_config=col_config,
        )

    # Parse cell selection
    selected_strike = None
    selected_date = None
    selected_type = None

    if event and event.selection and event.selection.cells:
        cell = event.selection.cells[0]
        # cells format: tuple(row_idx, col_name)
        row_idx = cell[0]
        col_name = cell[1]

        if 0 <= row_idx < len(rows):
            selected_strike = rows[row_idx].strike_price

            if col_name in put_day_cols:
                selected_type = "PUT"
                selected_date = _col_to_date(col_name, week)
            elif col_name in call_day_cols:
                selected_type = "CALL"
                selected_date = _col_to_date(col_name, week)
            # OI / total columns: show strike info but no daily breakdown
            elif col_name.startswith("P"):
                selected_type = "PUT"
            elif col_name.startswith("C"):
                selected_type = "CALL"

    with right_col:
        _render_breakdown_panel(
            rows, week, selected_strike, selected_date, selected_type,
            put_day_cols, call_day_cols, tab_label,
        )

    _render_option_summary(rows)


def _day_col(td, prefix: str) -> str:
    dow = _DOW_JP[td.weekday()]
    return f"{prefix}{td.strftime('%m/%d')}({dow})"


def _col_to_date(col_name: str, week: WeekDefinition) -> date | None:
    for td in week.trading_days:
        if _day_col(td, "P") == col_name or _day_col(td, "C") == col_name:
            return td
    return None


def _build_column_config(
    df: pd.DataFrame,
    put_day_cols: list[str],
    call_day_cols: list[str],
) -> dict:
    """Build column_config for integer formatting."""
    oi_cols = ["P前週L", "P前週S", "P今週L", "P今週S",
               "C前週L", "C前週S", "C今週L", "C今週S"]
    total_cols = ["P計", "C計"]
    all_num = put_day_cols + call_day_cols + oi_cols + total_cols

    col_config = {}
    for col in all_num:
        if col in df.columns:
            col_config[col] = st.column_config.NumberColumn(col, format="%d")
    col_config["行使価格"] = st.column_config.NumberColumn("行使価格", format="%d")
    return col_config


def _render_breakdown_panel(
    rows: list[OptionStrikeRow],
    week: WeekDefinition,
    selected_strike: int | None,
    selected_date: date | None,
    selected_type: str | None,
    put_day_cols: list[str],
    call_day_cols: list[str],
    tab_label: str,
) -> None:
    """Right panel: participant breakdown for selected cell."""
    st.markdown("**参加者別内訳**")

    if selected_strike is None:
        st.caption("左のテーブルの出来高セルをクリック")
        return

    # Find the row
    target_row = None
    for r in rows:
        if r.strike_price == selected_strike:
            target_row = r
            break
    if target_row is None:
        return

    if selected_type is None:
        selected_type = "CALL"

    # If no specific date selected (e.g. clicked OI column), let user pick
    if selected_date is None:
        st.markdown(f"**{selected_type} {selected_strike:,}**")
        day_labels = [f"{td.strftime('%m/%d')}({_DOW_JP[td.weekday()]})" for td in week.trading_days]
        prefix = f"bd_{tab_label}"
        day_choice = st.selectbox("日付", day_labels, key=f"{prefix}_day_r")
        if day_choice is None:
            return
        day_idx = day_labels.index(day_choice)
        selected_date = week.trading_days[day_idx]

    # Get breakdown
    is_put = selected_type == "PUT"
    breakdown = (target_row.put_daily_breakdown if is_put else target_row.call_daily_breakdown).get(selected_date, [])

    dow = _DOW_JP[selected_date.weekday()]
    header = f"{selected_type} {selected_strike:,}  {selected_date.strftime('%m/%d')}({dow})"

    if not breakdown:
        st.markdown(f"**{header}**")
        st.caption("データなし")
        return

    total = sum(v for _, v in breakdown)
    st.markdown(f"**{header}**")
    st.markdown(f"合計: **{int(total):,}**枚")

    bd_df = pd.DataFrame(breakdown, columns=["参加者", "枚数"])
    bd_df["枚数"] = bd_df["枚数"].astype(int)
    bd_df["構成比"] = (bd_df["枚数"] / total * 100).round(1).astype(str) + "%"
    st.dataframe(
        bd_df,
        use_container_width=True,
        hide_index=True,
        height=min(len(breakdown) * 35 + 40, 500),
    )


def _build_display_dataframe(
    rows: list[OptionStrikeRow],
    week: WeekDefinition,
) -> pd.DataFrame:
    """Build DataFrame with PUT left | strike center | CALL right layout."""
    records = []

    for row in rows:
        rec = {}

        rec["P前週L"] = row.put_start_oi_long
        rec["P前週S"] = row.put_start_oi_short

        for td in week.trading_days:
            col = _day_col(td, "P")
            rec[col] = row.put_daily_volumes.get(td) or None

        rec["P計"] = row.put_week_total
        rec["P今週L"] = row.put_end_oi_long
        rec["P今週S"] = row.put_end_oi_short

        rec["行使価格"] = row.strike_price

        rec["C今週L"] = row.call_end_oi_long
        rec["C今週S"] = row.call_end_oi_short
        rec["C計"] = row.call_week_total

        for td in reversed(week.trading_days):
            col = _day_col(td, "C")
            rec[col] = row.call_daily_volumes.get(td) or None

        rec["C前週L"] = row.call_start_oi_long
        rec["C前週S"] = row.call_start_oi_short

        records.append(rec)

    return pd.DataFrame(records)


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
