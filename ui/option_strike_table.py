"""Option strike-price table visualization.

Layout:
  PUT側（左）                                   | CALL側（右）
  前週L|前週S|1/6 1/7 ... |合計|今週L|今週S| 行使価格 |今週L|今週S|合計| ... 1/7 1/6|前週L|前週S

- Strike prices sorted descending (high to low)
- PUT daily columns: left-to-right (old → new)
- CALL daily columns: right-to-left (new → old)
"""
from __future__ import annotations

import streamlit as st
import pandas as pd
from models import OptionStrikeRow, WeekDefinition

_DOW_JP = ["月", "火", "水", "木", "金", "土", "日"]


def render_option_strike_table(
    rows: list[OptionStrikeRow],
    week: WeekDefinition,
    tab_label: str = "",
) -> None:
    """Render the option strike-price table."""
    title = f"日経225オプション ({week.label})"
    if tab_label and tab_label != "全セッション合計":
        title += f"  [{tab_label}]"
    st.subheader(title)

    if not rows:
        st.warning("オプションデータがありません。")
        return

    df = _build_display_dataframe(rows, week)
    styled = _apply_styling(df, week)

    st.dataframe(
        styled,
        use_container_width=True,
        height=min(len(rows) * 35 + 60, 900),
    )

    _render_option_summary(rows)


def _day_col(td, prefix: str) -> str:
    dow = _DOW_JP[td.weekday()]
    return f"{prefix}{td.strftime('%m/%d')}({dow})"


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

        # PUT daily volumes (left to right: old → new)
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

        # CALL daily volumes (right to left: new → old)
        for td in reversed(week.trading_days):
            col = _day_col(td, "C")
            rec[col] = row.call_daily_volumes.get(td) or None

        # CALL OI (start = 前週, right edge)
        rec["C前週L"] = row.call_start_oi_long
        rec["C前週S"] = row.call_start_oi_short

        records.append(rec)

    return pd.DataFrame(records)


def _apply_styling(df: pd.DataFrame, week: WeekDefinition):
    """Apply formatting and colors."""
    put_day_cols = [_day_col(td, "P") for td in week.trading_days]
    call_day_cols = [_day_col(td, "C") for td in reversed(week.trading_days)]
    oi_cols = ["P前週L", "P前週S", "P今週L", "P今週S",
               "C前週L", "C前週S", "C今週L", "C今週S"]
    total_cols = ["P計", "C計"]
    num_cols = put_day_cols + call_day_cols + oi_cols + total_cols

    fmt_int = lambda v: f"{int(v):,}" if pd.notna(v) else ""

    styled = df.style

    for col in df.columns:
        if col in num_cols:
            styled = styled.format(fmt_int, subset=[col])
        elif col == "行使価格":
            styled = styled.format(lambda v: f"{int(v):,}", subset=[col])

    # Highlight strike column
    def _strike_style(val):
        return "font-weight: bold; background-color: #e8e8e8"

    if "行使価格" in df.columns:
        styled = styled.map(_strike_style, subset=["行使価格"])

    # Color PUT columns with light red tint, CALL with light blue tint
    def _put_bg(val):
        if pd.notna(val) and val > 0:
            return "background-color: #fff0f0"
        return ""

    def _call_bg(val):
        if pd.notna(val) and val > 0:
            return "background-color: #f0f0ff"
        return ""

    for col in put_day_cols + total_cols[:1]:
        if col in df.columns:
            styled = styled.map(_put_bg, subset=[col])

    for col in call_day_cols + total_cols[1:]:
        if col in df.columns:
            styled = styled.map(_call_bg, subset=[col])

    return styled


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
