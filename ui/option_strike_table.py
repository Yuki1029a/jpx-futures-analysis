"""Option strike-price table with Styler-based rendering for readability.

Each strike = one DataFrame row showing:
  - PUT side: weekly OI, daily volumes, JPX volume, daily OI
  - 行使価格 (center, highlighted)
  - CALL side: daily OI, JPX volume, daily volumes, weekly OI

Uses pandas Styler for conditional formatting, color-coded columns,
and number formatting consistent with the futures weekly_table.py.
"""
from __future__ import annotations

import streamlit as st
import pandas as pd
from datetime import date
from models import OptionStrikeRow, WeekDefinition

_DOW_JP = ["月", "火", "水", "木", "金", "土", "日"]

# Summary header row count
_SUMMARY_ROWS = 2

# --- Color palette ---
_PUT_BG = "#fff0f0"         # Very light red for PUT columns
_CALL_BG = "#f0f4ff"        # Very light blue for CALL columns
_STRIKE_BG = "#fffde7"      # Light yellow for strike price column
_SUMMARY_PUT_BG = "#f8d7da" # Pink for PUT summary row
_SUMMARY_CALL_BG = "#cfe2ff" # Light blue for CALL summary row
_OI_BG_P = "#fce4ec"        # PUT OI column slightly darker
_OI_BG_C = "#e3f2fd"        # CALL OI column slightly darker
_JPX_BG_P = "#fff3e0"       # PUT JPX volume warm tint
_JPX_BG_C = "#e8f5e9"       # CALL JPX volume cool tint
_HEADER_BG = "#f5f5f5"      # Neutral for week OI headers


def render_option_strike_table(
    rows: list[OptionStrikeRow],
    week: WeekDefinition,
    tab_label: str = "",
) -> None:
    """Render option table (styled) + detail panel (right)."""
    title = f"日経225オプション ({week.label})"
    if tab_label and tab_label != "全セッション合計":
        title += f"  [{tab_label}]"
    st.subheader(title)

    if not rows:
        st.warning("オプションデータがありません。")
        return

    # Build ordered column list
    ordered_cols = _build_column_order(week)
    df = _build_display_dataframe(rows, week, ordered_cols)

    # Apply styling
    styled = _apply_styling(df, week, ordered_cols)

    # Layout: table (left 75%) | detail (right 25%)
    left_col, right_col = st.columns([3, 1])

    with left_col:
        st.dataframe(
            styled,
            use_container_width=True,
            height=min(len(df) * 35 + 60, 900),
        )

    with right_col:
        _render_detail_panel_selectbox(rows, week, tab_label)

    _render_option_summary(rows)


# --- Column name helpers ---

def _day_col(td: date, prefix: str) -> str:
    dow = _DOW_JP[td.weekday()]
    return f"{prefix}{td.strftime('%m/%d')}({dow})"


def _jpx_vol_col(td: date, prefix: str) -> str:
    return f"{prefix}出{td.strftime('%d')}"


def _oi_col(td: date, prefix: str) -> str:
    return f"{prefix}建{td.strftime('%d')}"


def _oi_chg_col(td: date, prefix: str) -> str:
    return f"{prefix}増{td.strftime('%d')}"


# --- Column order ---

def _build_column_order(week: WeekDefinition) -> list[str]:
    """Build explicit column order: PUT side | 行使価格 | CALL side."""
    cols = []

    # PUT side (left to right)
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

    # Center
    cols.append("行使価格")

    # CALL side (mirror)
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


# --- Styling ---

def _apply_styling(
    df: pd.DataFrame,
    week: WeekDefinition,
    ordered_cols: list[str],
) -> pd.io.formats.style.Styler:
    """Apply comprehensive styling: background colors, number formatting, alignment."""

    # Column classification
    put_day_cols = set(_day_col(td, "P") for td in week.trading_days)
    call_day_cols = set(_day_col(td, "C") for td in week.trading_days)
    put_jpx_cols = set(_jpx_vol_col(td, "P") for td in week.trading_days)
    call_jpx_cols = set(_jpx_vol_col(td, "C") for td in week.trading_days)
    put_oi_cols = set(_oi_col(td, "P") for td in week.trading_days)
    call_oi_cols = set(_oi_col(td, "C") for td in week.trading_days)
    put_chg_cols = set(_oi_chg_col(td, "P") for td in week.trading_days)
    call_chg_cols = set(_oi_chg_col(td, "C") for td in week.trading_days)
    put_week_oi = {"P前週L", "P前週S", "P今週L", "P今週S", "P計"}
    call_week_oi = {"C前週L", "C前週S", "C今週L", "C今週S", "C計"}

    signed_cols = put_chg_cols | call_chg_cols

    def _cell_style(row_idx, col):
        """Determine background color for a cell."""
        # Summary rows
        if row_idx == 0:  # PUT合計
            return f"background-color: {_SUMMARY_PUT_BG}; font-weight: bold"
        if row_idx == 1:  # CALL合計
            return f"background-color: {_SUMMARY_CALL_BG}; font-weight: bold"

        # Strike price column
        if col == "行使価格":
            return f"background-color: {_STRIKE_BG}; font-weight: bold; text-align: center"

        # PUT side
        if col in put_day_cols or col in put_week_oi:
            return f"background-color: {_PUT_BG}"
        if col in put_jpx_cols:
            return f"background-color: {_JPX_BG_P}"
        if col in put_oi_cols:
            return f"background-color: {_OI_BG_P}"
        if col in put_chg_cols:
            return f"background-color: {_OI_BG_P}"

        # CALL side
        if col in call_day_cols or col in call_week_oi:
            return f"background-color: {_CALL_BG}"
        if col in call_jpx_cols:
            return f"background-color: {_JPX_BG_C}"
        if col in call_oi_cols:
            return f"background-color: {_OI_BG_C}"
        if col in call_chg_cols:
            return f"background-color: {_OI_BG_C}"

        return ""

    def _apply_cell_colors(s):
        """Apply per-cell background colors."""
        row_idx = s.name
        return [_cell_style(row_idx, col) for col in s.index]

    def _color_signed(val):
        """Green for positive, red for negative OI changes."""
        if pd.isna(val):
            return ""
        try:
            n = float(val)
            if n > 0:
                return "color: #006100"
            elif n < 0:
                return "color: #9c0006"
        except (ValueError, TypeError):
            pass
        return ""

    styled = df.style.apply(_apply_cell_colors, axis=1)

    # Signed coloring for OI change columns (participant rows only)
    all_chg = list(put_chg_cols | call_chg_cols)
    valid_chg = [c for c in all_chg if c in df.columns]
    if valid_chg:
        participant_idx = list(range(_SUMMARY_ROWS, len(df)))
        if participant_idx:
            styled = styled.map(
                _color_signed,
                subset=(participant_idx, valid_chg),
            )

    # Number formatting
    fmt_int = lambda v: f"{int(v):,}" if pd.notna(v) and v != "" else "-"
    fmt_signed = lambda v: f"{int(v):+,}" if pd.notna(v) and v != "" else "-"

    for col in df.columns:
        if col == "行使価格":
            continue
        if col in signed_cols:
            styled = styled.format(fmt_signed, subset=[col])
        else:
            styled = styled.format(fmt_int, subset=[col])

    return styled


# --- DataFrame building ---

def _build_display_dataframe(
    rows: list[OptionStrikeRow],
    week: WeekDefinition,
    ordered_cols: list[str],
) -> pd.DataFrame:
    """Build DataFrame with summary rows + one row per strike.

    Row 0: PUT合計 (all-strike totals for PUT side)
    Row 1: CALL合計 (all-strike totals for CALL side)
    Row 2+: individual strikes (sorted by strike price)
    """
    summary_rows = _build_summary_rows(rows, week)
    records = []
    for row in rows:
        rec = _build_volume_row(row, week)
        records.append(rec)

    df = pd.DataFrame(summary_rows + records, columns=ordered_cols)
    return df


def _build_summary_rows(
    rows: list[OptionStrikeRow],
    week: WeekDefinition,
) -> list[dict]:
    """Build PUT合計 and CALL合計 summary rows."""
    put_rec = {"行使価格": "PUT合計"}
    call_rec = {"行使価格": "CALL合計"}

    # Initialize all week-level OI cols to None
    for col in ("P前週L", "P前週S", "P今週L", "P今週S",
                "C前週L", "C前週S", "C今週L", "C今週S"):
        put_rec[col] = None
        call_rec[col] = None

    put_total_week = 0.0
    call_total_week = 0.0

    for td in week.trading_days:
        p_vol = sum(r.put_daily_volumes.get(td, 0) for r in rows)
        c_vol = sum(r.call_daily_volumes.get(td, 0) for r in rows)
        p_jpx = sum(r.put_daily_jpx_volume.get(td, 0) for r in rows)
        c_jpx = sum(r.call_daily_jpx_volume.get(td, 0) for r in rows)
        p_oi = sum(r.put_daily_oi.get(td, 0) for r in rows)
        c_oi = sum(r.call_daily_oi.get(td, 0) for r in rows)
        p_chg = sum(r.put_daily_oi_change.get(td, 0) for r in rows)
        c_chg = sum(r.call_daily_oi_change.get(td, 0) for r in rows)

        # PUT summary row: only PUT columns populated
        put_rec[_day_col(td, "P")] = p_vol or None
        put_rec[_jpx_vol_col(td, "P")] = p_jpx or None
        put_rec[_oi_col(td, "P")] = p_oi or None
        put_rec[_oi_chg_col(td, "P")] = p_chg or None
        put_rec[_day_col(td, "C")] = None
        put_rec[_jpx_vol_col(td, "C")] = None
        put_rec[_oi_col(td, "C")] = None
        put_rec[_oi_chg_col(td, "C")] = None

        # CALL summary row: only CALL columns populated
        call_rec[_day_col(td, "C")] = c_vol or None
        call_rec[_jpx_vol_col(td, "C")] = c_jpx or None
        call_rec[_oi_col(td, "C")] = c_oi or None
        call_rec[_oi_chg_col(td, "C")] = c_chg or None
        call_rec[_day_col(td, "P")] = None
        call_rec[_jpx_vol_col(td, "P")] = None
        call_rec[_oi_col(td, "P")] = None
        call_rec[_oi_chg_col(td, "P")] = None

        put_total_week += p_vol
        call_total_week += c_vol

    put_rec["P計"] = put_total_week or None
    put_rec["C計"] = None
    call_rec["P計"] = None
    call_rec["C計"] = call_total_week or None

    return [put_rec, call_rec]


def _build_volume_row(
    row: OptionStrikeRow,
    week: WeekDefinition,
) -> dict:
    """Build the data row for a single strike price."""
    rec = {}

    # --- PUT side ---
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

    # --- Strike price (center) ---
    rec["行使価格"] = f"{row.strike_price:,}"

    # --- CALL side ---
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


# --- Detail panel (selectbox-based, replaces click-based) ---

def _render_detail_panel_selectbox(
    rows: list[OptionStrikeRow],
    week: WeekDefinition,
    tab_label: str,
) -> None:
    """Right panel: user selects strike/date/type via selectbox."""
    st.markdown("**詳細パネル**")

    if not rows:
        return

    prefix = f"det_{tab_label}"

    # Option type
    opt_type = st.radio(
        "タイプ", ["PUT", "CALL"],
        horizontal=True, key=f"{prefix}_type",
    )

    # Strike selection
    strike_labels = [f"{r.strike_price:,}" for r in rows]
    # Default to ATM-ish (middle)
    default_idx = len(rows) // 2
    strike_choice = st.selectbox(
        "行使価格", strike_labels,
        index=default_idx, key=f"{prefix}_strike",
    )
    if strike_choice is None:
        return
    strike_idx = strike_labels.index(strike_choice)
    target_row = rows[strike_idx]

    # Date selection
    day_labels = [f"{td.strftime('%m/%d')}({_DOW_JP[td.weekday()]})"
                  for td in week.trading_days]
    day_choice = st.selectbox(
        "日付", day_labels,
        index=len(day_labels) - 1,  # Default to latest day
        key=f"{prefix}_day",
    )
    if day_choice is None:
        return
    selected_date = week.trading_days[day_labels.index(day_choice)]

    dow = _DOW_JP[selected_date.weekday()]
    date_str = f"{selected_date.strftime('%m/%d')}({dow})"

    st.markdown("---")
    _render_participant_breakdown(target_row, opt_type, selected_date, date_str)
    _render_oi_detail(target_row, opt_type, selected_date)


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

    jpx_vol = (row.put_daily_jpx_volume if is_put else row.call_daily_jpx_volume).get(td)
    if jpx_vol:
        st.caption(f"JPX出来高: {jpx_vol:,}")


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
