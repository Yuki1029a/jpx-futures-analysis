"""Option strike-price table — unified single table with cell click → detail.

Single st.dataframe with on_select: click any cell to see participant
breakdown in the detail panel. 行使価格 as index (pinned left on scroll).
column_config for number formatting. No separate navigator needed.
"""
from __future__ import annotations

import streamlit as st
import pandas as pd
from datetime import date
from models import OptionStrikeRow, WeekDefinition

_DOW_JP = ["月", "火", "水", "木", "金", "土", "日"]

_SUMMARY_ROWS = 2


def render_option_strike_table(
    rows: list[OptionStrikeRow],
    week: WeekDefinition,
    tab_label: str = "",
) -> None:
    """Render single interactive option table + detail panel."""
    title = f"日経225オプション ({week.label})"
    if tab_label and tab_label != "全セッション合計":
        title += f"  [{tab_label}]"
    st.subheader(title)

    if not rows:
        st.warning("オプションデータがありません。")
        return

    # Build DataFrame
    ordered_cols = _build_column_order(week)
    df = _build_display_dataframe(rows, week, ordered_cols)

    # Column config for formatting
    col_config = _build_column_config(df, week)

    # Hide metadata column
    col_config["_strike_idx"] = None

    # Build column order for display (exclude hidden cols)
    display_cols = [c for c in ordered_cols if c != "行使価格"]

    # Layout: table (left) | detail (right)
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
            column_order=display_cols,
        )

    # Parse cell selection
    selected_strike_idx = None
    selected_date = None
    selected_type = None

    # Classify columns
    put_cols, call_cols = _classify_columns(week)

    if event and event.selection and event.selection.cells:
        cell = event.selection.cells[0]
        df_row_idx = cell[0]
        col_name = cell[1]

        if 0 <= df_row_idx < len(df):
            strike_idx_val = df.iloc[df_row_idx].get("_strike_idx")
            if strike_idx_val is not None and not pd.isna(strike_idx_val):
                selected_strike_idx = int(strike_idx_val)

            if col_name in put_cols:
                selected_type = "PUT"
                selected_date = _col_to_date(col_name, week)
            elif col_name in call_cols:
                selected_type = "CALL"
                selected_date = _col_to_date(col_name, week)
            elif col_name and col_name.startswith("P"):
                selected_type = "PUT"
            elif col_name and col_name.startswith("C"):
                selected_type = "CALL"

    with right_col:
        _render_detail_panel(
            rows, week,
            selected_strike_idx, selected_date, selected_type,
            tab_label,
        )

    _render_option_summary(rows)


# =====================================================================
# Column name helpers
# =====================================================================

def _day_col(td: date, prefix: str) -> str:
    dow = _DOW_JP[td.weekday()]
    return f"{prefix}{td.strftime('%m/%d')}({dow})"


def _jpx_vol_col(td: date, prefix: str) -> str:
    return f"{prefix}出{td.strftime('%d')}"


def _oi_col(td: date, prefix: str) -> str:
    return f"{prefix}建{td.strftime('%d')}"


def _oi_chg_col(td: date, prefix: str) -> str:
    return f"{prefix}増{td.strftime('%d')}"


def _classify_columns(week: WeekDefinition) -> tuple[set[str], set[str]]:
    """Return (put_cols, call_cols) sets for all per-date columns."""
    put_cols = set()
    call_cols = set()
    for td in week.trading_days:
        put_cols |= {_day_col(td, "P"), _jpx_vol_col(td, "P"),
                     _oi_col(td, "P"), _oi_chg_col(td, "P")}
        call_cols |= {_day_col(td, "C"), _jpx_vol_col(td, "C"),
                      _oi_col(td, "C"), _oi_chg_col(td, "C")}
    return put_cols, call_cols


def _col_to_date(col_name: str, week: WeekDefinition) -> date | None:
    """Resolve any per-date column name to a date."""
    for td in week.trading_days:
        for prefix in ("P", "C"):
            if col_name in (_day_col(td, prefix), _jpx_vol_col(td, prefix),
                            _oi_col(td, prefix), _oi_chg_col(td, prefix)):
                return td
    return None


# =====================================================================
# Column order
# =====================================================================

def _build_column_order(week: WeekDefinition) -> list[str]:
    """PUT side | 行使価格 | CALL side."""
    cols = []

    cols.append("P前週L")
    cols.append("P前週S")
    for td in week.trading_days:
        cols.append(_day_col(td, "P"))
        cols.append(_jpx_vol_col(td, "P"))
        cols.append(_oi_col(td, "P"))
        cols.append(_oi_chg_col(td, "P"))
    cols.append("P計")
    cols.append("P今週L")
    cols.append("P今週S")

    cols.append("行使価格")

    cols.append("C今週L")
    cols.append("C今週S")
    cols.append("C計")
    for td in reversed(week.trading_days):
        cols.append(_oi_chg_col(td, "C"))
        cols.append(_oi_col(td, "C"))
        cols.append(_jpx_vol_col(td, "C"))
        cols.append(_day_col(td, "C"))
    cols.append("C前週L")
    cols.append("C前週S")

    return cols


# =====================================================================
# Column config
# =====================================================================

def _build_column_config(df: pd.DataFrame, week: WeekDefinition) -> dict:
    """Build column_config for number formatting."""
    cfg = {}

    # 行使価格 — pinned text column (index)
    cfg["行使価格"] = st.column_config.TextColumn("行使価格", pinned=True)

    # All numeric columns get NumberColumn with comma formatting
    for col in df.columns:
        if col in ("行使価格", "_strike_idx"):
            continue
        cfg[col] = st.column_config.NumberColumn(col, format="%d")

    return cfg


# =====================================================================
# DataFrame building
# =====================================================================

def _build_display_dataframe(
    rows: list[OptionStrikeRow],
    week: WeekDefinition,
    ordered_cols: list[str],
) -> pd.DataFrame:
    """Build DataFrame with summary rows + one row per strike.

    Row 0: PUT合計, Row 1: CALL合計, Row 2+: individual strikes.
    """
    summary_rows = _build_summary_rows(rows, week)
    records = []
    for idx, row in enumerate(rows):
        rec = _build_volume_row(row, week)
        rec["_strike_idx"] = idx
        records.append(rec)

    all_cols = ordered_cols + ["_strike_idx"]
    return pd.DataFrame(summary_rows + records, columns=all_cols)


def _build_summary_rows(rows, week):
    put_rec = {"行使価格": "PUT合計", "_strike_idx": None}
    call_rec = {"行使価格": "CALL合計", "_strike_idx": None}

    for col in ("P前週L", "P前週S", "P今週L", "P今週S",
                "C前週L", "C前週S", "C今週L", "C今週S"):
        put_rec[col] = None
        call_rec[col] = None

    put_total = 0.0
    call_total = 0.0

    for td in week.trading_days:
        p_vol = sum(r.put_daily_volumes.get(td, 0) for r in rows)
        c_vol = sum(r.call_daily_volumes.get(td, 0) for r in rows)
        p_jpx = sum(r.put_daily_jpx_volume.get(td, 0) for r in rows)
        c_jpx = sum(r.call_daily_jpx_volume.get(td, 0) for r in rows)
        p_oi = sum(r.put_daily_oi.get(td, 0) for r in rows)
        c_oi = sum(r.call_daily_oi.get(td, 0) for r in rows)
        p_chg = sum(r.put_daily_oi_change.get(td, 0) for r in rows)
        c_chg = sum(r.call_daily_oi_change.get(td, 0) for r in rows)

        put_rec[_day_col(td, "P")] = p_vol or None
        put_rec[_jpx_vol_col(td, "P")] = p_jpx or None
        put_rec[_oi_col(td, "P")] = p_oi or None
        put_rec[_oi_chg_col(td, "P")] = p_chg or None
        put_rec[_day_col(td, "C")] = None
        put_rec[_jpx_vol_col(td, "C")] = None
        put_rec[_oi_col(td, "C")] = None
        put_rec[_oi_chg_col(td, "C")] = None

        call_rec[_day_col(td, "C")] = c_vol or None
        call_rec[_jpx_vol_col(td, "C")] = c_jpx or None
        call_rec[_oi_col(td, "C")] = c_oi or None
        call_rec[_oi_chg_col(td, "C")] = c_chg or None
        call_rec[_day_col(td, "P")] = None
        call_rec[_jpx_vol_col(td, "P")] = None
        call_rec[_oi_col(td, "P")] = None
        call_rec[_oi_chg_col(td, "P")] = None

        put_total += p_vol
        call_total += c_vol

    put_rec["P計"] = put_total or None
    put_rec["C計"] = None
    call_rec["P計"] = None
    call_rec["C計"] = call_total or None

    return [put_rec, call_rec]


def _build_volume_row(row, week):
    rec = {}
    rec["P前週L"] = row.put_start_oi_long
    rec["P前週S"] = row.put_start_oi_short

    for td in week.trading_days:
        rec[_day_col(td, "P")] = row.put_daily_volumes.get(td) or None
        rec[_jpx_vol_col(td, "P")] = row.put_daily_jpx_volume.get(td) or None
        rec[_oi_col(td, "P")] = row.put_daily_oi.get(td) or None
        rec[_oi_chg_col(td, "P")] = row.put_daily_oi_change.get(td) or None

    rec["P計"] = row.put_week_total
    rec["P今週L"] = row.put_end_oi_long
    rec["P今週S"] = row.put_end_oi_short

    rec["行使価格"] = f"{row.strike_price:,}"

    rec["C今週L"] = row.call_end_oi_long
    rec["C今週S"] = row.call_end_oi_short
    rec["C計"] = row.call_week_total

    for td in reversed(week.trading_days):
        rec[_oi_chg_col(td, "C")] = row.call_daily_oi_change.get(td) or None
        rec[_oi_col(td, "C")] = row.call_daily_oi.get(td) or None
        rec[_jpx_vol_col(td, "C")] = row.call_daily_jpx_volume.get(td) or None
        rec[_day_col(td, "C")] = row.call_daily_volumes.get(td) or None

    rec["C前週L"] = row.call_start_oi_long
    rec["C前週S"] = row.call_start_oi_short

    return rec


# =====================================================================
# Detail panel
# =====================================================================

def _render_detail_panel(
    rows: list[OptionStrikeRow],
    week: WeekDefinition,
    strike_idx: int | None,
    selected_date: date | None,
    selected_type: str | None,
    tab_label: str,
) -> None:
    st.markdown("**詳細パネル**")

    if strike_idx is None or strike_idx >= len(rows):
        st.caption("テーブルのセルをクリック")
        return

    target_row = rows[strike_idx]

    if selected_type is None:
        selected_type = "CALL"

    if selected_date is None:
        day_labels = [f"{td.strftime('%m/%d')}({_DOW_JP[td.weekday()]})"
                      for td in week.trading_days]
        prefix = f"bd_{tab_label}"
        day_choice = st.selectbox("日付", day_labels, key=f"{prefix}_day_r")
        if day_choice is None:
            return
        selected_date = week.trading_days[day_labels.index(day_choice)]

    dow = _DOW_JP[selected_date.weekday()]
    date_str = f"{selected_date.strftime('%m/%d')}({dow})"

    _render_participant_breakdown(target_row, selected_type, selected_date, date_str)
    _render_oi_detail(target_row, selected_type, selected_date)


def _render_participant_breakdown(row, option_type, td, date_str):
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


def _render_oi_detail(row, option_type, td):
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

    jpx_vol = (row.put_daily_jpx_volume if is_put else row.call_daily_jpx_volume).get(td)
    if jpx_vol:
        st.caption(f"JPX出来高: {jpx_vol:,}")


def _render_option_summary(rows: list[OptionStrikeRow]) -> None:
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
