"""Option strike-price table visualization with cell click breakdown.

Layout:
  PUT側（左）                                   | CALL側（右）
  前週L|前週S|1/6 1/7 ... |合計|今週L|今週S| 行使価格 |今週L|今週S|合計| ... 1/7 1/6|前週L|前週S

- Strike prices sorted descending (high to low)
- PUT daily columns: left-to-right (old -> new)
- CALL daily columns: right-to-left (new -> old)
- Click a daily volume cell to see per-participant breakdown
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
    """Render the option strike-price table with cell-click breakdown."""
    title = f"日経225オプション ({week.label})"
    if tab_label and tab_label != "全セッション合計":
        title += f"  [{tab_label}]"
    st.subheader(title)

    if not rows:
        st.warning("オプションデータがありません。")
        return

    df = _build_display_dataframe(rows, week)

    # Build column config for integer formatting
    col_config = {}
    put_day_cols = [_day_col(td, "P") for td in week.trading_days]
    call_day_cols = [_day_col(td, "C") for td in reversed(week.trading_days)]
    num_cols = (
        put_day_cols + call_day_cols
        + ["P前週L", "P前週S", "P今週L", "P今週S",
           "C前週L", "C前週S", "C今週L", "C今週S",
           "P計", "C計"]
    )
    for col in num_cols:
        if col in df.columns:
            col_config[col] = st.column_config.NumberColumn(
                col, format="%d",
            )
    col_config["行使価格"] = st.column_config.NumberColumn(
        "行使価格", format="%d",
    )

    # Unique key per tab to avoid widget conflicts
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

    # Handle cell selection -> show breakdown dialog
    if event and event.selection and event.selection.cells:
        cell = event.selection.cells[0]
        row_idx = cell["row"]
        col_name = cell["column"]

        # Only respond to daily volume columns (P02/05(月) or C02/05(月))
        if col_name in put_day_cols or col_name in call_day_cols:
            _show_breakdown(rows, week, row_idx, col_name, put_day_cols, call_day_cols)

    _render_option_summary(rows)


def _day_col(td, prefix: str) -> str:
    dow = _DOW_JP[td.weekday()]
    return f"{prefix}{td.strftime('%m/%d')}({dow})"


def _col_to_date(col_name: str, week: WeekDefinition) -> date | None:
    """Extract date from column name like 'P02/05(月)' or 'C02/05(月)'."""
    for td in week.trading_days:
        if _day_col(td, "P") == col_name or _day_col(td, "C") == col_name:
            return td
    return None


def _show_breakdown(
    rows: list[OptionStrikeRow],
    week: WeekDefinition,
    row_idx: int,
    col_name: str,
    put_day_cols: list[str],
    call_day_cols: list[str],
):
    """Show per-participant breakdown for selected cell."""
    if row_idx < 0 or row_idx >= len(rows):
        return

    row = rows[row_idx]
    td = _col_to_date(col_name, week)
    if td is None:
        return

    is_put = col_name in put_day_cols
    option_type = "PUT" if is_put else "CALL"
    breakdown = (row.put_daily_breakdown if is_put else row.call_daily_breakdown).get(td, [])

    if not breakdown:
        return

    dow = _DOW_JP[td.weekday()]
    total = sum(v for _, v in breakdown)

    st.markdown("---")
    st.markdown(
        f"**{option_type} {row.strike_price:,}  "
        f"{td.strftime('%m/%d')}({dow})  "
        f"合計: {int(total):,}枚**"
    )

    bd_df = pd.DataFrame(breakdown, columns=["参加者", "枚数"])
    bd_df["枚数"] = bd_df["枚数"].astype(int)
    bd_df["構成比"] = (bd_df["枚数"] / total * 100).round(1).astype(str) + "%"
    st.dataframe(
        bd_df,
        use_container_width=False,
        hide_index=True,
        height=min(len(breakdown) * 35 + 40, 400),
    )


def _build_display_dataframe(
    rows: list[OptionStrikeRow],
    week: WeekDefinition,
) -> pd.DataFrame:
    """Build DataFrame with PUT left | strike center | CALL right layout."""
    records = []

    for row in rows:
        rec = {}

        # PUT OI (start)
        rec["P前週L"] = row.put_start_oi_long
        rec["P前週S"] = row.put_start_oi_short

        # PUT daily volumes (left to right: old -> new)
        for td in week.trading_days:
            col = _day_col(td, "P")
            rec[col] = row.put_daily_volumes.get(td) or None

        rec["P計"] = row.put_week_total

        # PUT OI (end)
        rec["P今週L"] = row.put_end_oi_long
        rec["P今週S"] = row.put_end_oi_short

        # Strike price (center)
        rec["行使価格"] = row.strike_price

        # CALL OI (end = 今週, left side near strike)
        rec["C今週L"] = row.call_end_oi_long
        rec["C今週S"] = row.call_end_oi_short

        # CALL weekly total
        rec["C計"] = row.call_week_total

        # CALL daily volumes (right to left: new -> old)
        for td in reversed(week.trading_days):
            col = _day_col(td, "C")
            rec[col] = row.call_daily_volumes.get(td) or None

        # CALL OI (start = 前週, right edge)
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
