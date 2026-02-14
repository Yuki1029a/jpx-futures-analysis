"""Option strike-price table with Styler-based rendering + clickable navigator.

Architecture:
  1. Main table (Styler → HTML) — full data with color coding, number formatting,
     sticky 行使価格 column that stays visible during horizontal scroll.
  2. Navigator grid (st.dataframe on_select) — compact strike x date grid
     showing daily volumes + OI per strike. Click a cell → detail panel.
"""
from __future__ import annotations

import streamlit as st
import pandas as pd
from datetime import date
from models import OptionStrikeRow, WeekDefinition

_DOW_JP = ["月", "火", "水", "木", "金", "土", "日"]

_SUMMARY_ROWS = 2

# --- Color palette ---
_PUT_BG = "#fff0f0"
_CALL_BG = "#f0f4ff"
_STRIKE_BG = "#fffde7"
_SUMMARY_PUT_BG = "#f8d7da"
_SUMMARY_CALL_BG = "#cfe2ff"
_OI_BG_P = "#fce4ec"
_OI_BG_C = "#e3f2fd"
_JPX_BG_P = "#fff3e0"
_JPX_BG_C = "#e8f5e9"


def render_option_strike_table(
    rows: list[OptionStrikeRow],
    week: WeekDefinition,
    tab_label: str = "",
) -> None:
    """Render option table (styled) + navigator + detail panel."""
    title = f"日経225オプション ({week.label})"
    if tab_label and tab_label != "全セッション合計":
        title += f"  [{tab_label}]"
    st.subheader(title)

    if not rows:
        st.warning("オプションデータがありません。")
        return

    # ====== 1. Main Styler table (HTML with sticky strike column) ======
    ordered_cols = _build_column_order(week)
    df = _build_display_dataframe(rows, week, ordered_cols)
    styled = _apply_styling(df, week)

    # Find strike column index (0-based among data columns)
    strike_col_idx = ordered_cols.index("行使価格")

    _render_sticky_table(styled, strike_col_idx, tab_label)

    # ====== 2. Navigator grid + Detail panel (side by side) ======
    st.markdown("---")
    st.caption("行使価格ナビゲーター — セルをクリックで詳細表示")

    nav_left, nav_right = st.columns([2, 1])

    nav_df, nav_meta = _build_navigator_df(rows, week)

    with nav_left:
        nav_key = f"nav_{tab_label}"
        event = st.dataframe(
            nav_df,
            use_container_width=True,
            height=min(len(nav_df) * 35 + 50, 500),
            on_select="rerun",
            selection_mode="single-cell",
            key=nav_key,
            column_config=_nav_column_config(nav_df),
        )

    selected_strike_idx = None
    selected_date = None
    selected_type = None

    if event and event.selection and event.selection.cells:
        cell = event.selection.cells[0]
        df_row_idx = cell[0]
        col_name = cell[1]

        if 0 <= df_row_idx < len(nav_meta):
            selected_strike_idx = nav_meta[df_row_idx]

        if col_name and col_name != "行使価格":
            selected_date, selected_type = _parse_nav_col(col_name, week)

    with nav_right:
        _render_detail_panel(
            rows, week,
            selected_strike_idx, selected_date, selected_type,
            tab_label,
        )

    _render_option_summary(rows)


# =====================================================================
# Sticky HTML table rendering
# =====================================================================

def _render_sticky_table(
    styled: pd.io.formats.style.Styler,
    strike_col_idx: int,
    tab_label: str,
) -> None:
    """Render Styler as HTML with sticky strike price column.

    The strike price column (and header) stays fixed while scrolling horizontally.
    Uses CSS position:sticky with appropriate left offset.
    """
    # th/td are 1-indexed in nth-child (col 0 = row header if present)
    # Styler hides index by default when we set hide(axis="index")
    styled = styled.hide(axis="index")
    html = styled.to_html()

    # CSS nth-child is 1-based, and since index is hidden,
    # column 0 in data = nth-child(1) in rendered table
    nth = strike_col_idx + 1

    table_id = f"sticky_table_{tab_label}".replace(" ", "_")

    css = f"""
    <style>
    #{table_id} {{
        overflow-x: auto;
        max-height: 800px;
        overflow-y: auto;
        border: 1px solid #ddd;
        position: relative;
    }}
    #{table_id} table {{
        border-collapse: separate;
        border-spacing: 0;
        font-size: 12px;
        white-space: nowrap;
    }}
    #{table_id} th, #{table_id} td {{
        padding: 4px 8px;
        border: 1px solid #e0e0e0;
        text-align: right;
    }}
    /* Sticky strike price column */
    #{table_id} td:nth-child({nth}),
    #{table_id} th:nth-child({nth}) {{
        position: sticky;
        left: 0;
        z-index: 2;
        background-color: {_STRIKE_BG} !important;
        font-weight: bold;
        text-align: center;
        border-right: 2px solid #bbb;
        min-width: 80px;
    }}
    /* Sticky header row */
    #{table_id} thead th {{
        position: sticky;
        top: 0;
        z-index: 3;
        background-color: #f8f9fa;
        border-bottom: 2px solid #999;
    }}
    /* Corner cell: both sticky directions */
    #{table_id} thead th:nth-child({nth}) {{
        z-index: 4;
        background-color: {_STRIKE_BG} !important;
    }}
    </style>
    """

    full_html = f'{css}<div id="{table_id}">{html}</div>'
    st.markdown(full_html, unsafe_allow_html=True)


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


# =====================================================================
# Main table column order
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
# Styler for main table
# =====================================================================

def _apply_styling(
    df: pd.DataFrame,
    week: WeekDefinition,
) -> pd.io.formats.style.Styler:
    """Color coding + number formatting."""

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
        if row_idx == 0:
            return f"background-color: {_SUMMARY_PUT_BG}; font-weight: bold"
        if row_idx == 1:
            return f"background-color: {_SUMMARY_CALL_BG}; font-weight: bold"
        if col == "行使価格":
            return f"background-color: {_STRIKE_BG}; font-weight: bold; text-align: center"
        if col in put_day_cols or col in put_week_oi:
            return f"background-color: {_PUT_BG}"
        if col in put_jpx_cols:
            return f"background-color: {_JPX_BG_P}"
        if col in put_oi_cols or col in put_chg_cols:
            return f"background-color: {_OI_BG_P}"
        if col in call_day_cols or col in call_week_oi:
            return f"background-color: {_CALL_BG}"
        if col in call_jpx_cols:
            return f"background-color: {_JPX_BG_C}"
        if col in call_oi_cols or col in call_chg_cols:
            return f"background-color: {_OI_BG_C}"
        return ""

    def _apply_cell_colors(s):
        return [_cell_style(s.name, col) for col in s.index]

    def _color_signed(val):
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

    valid_chg = [c for c in signed_cols if c in df.columns]
    if valid_chg:
        participant_idx = list(range(_SUMMARY_ROWS, len(df)))
        if participant_idx:
            styled = styled.map(_color_signed, subset=(participant_idx, valid_chg))

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


# =====================================================================
# Main table DataFrame
# =====================================================================

def _build_display_dataframe(
    rows: list[OptionStrikeRow],
    week: WeekDefinition,
    ordered_cols: list[str],
) -> pd.DataFrame:
    summary_rows = _build_summary_rows(rows, week)
    records = [_build_volume_row(row, week) for row in rows]
    return pd.DataFrame(summary_rows + records, columns=ordered_cols)


def _build_summary_rows(rows, week):
    put_rec = {"行使価格": "PUT合計"}
    call_rec = {"行使価格": "CALL合計"}

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
# Navigator grid (compact, clickable) — daily volume + OI
# =====================================================================

def _nav_day_label(td: date) -> str:
    return f"{td.strftime('%m/%d')}({_DOW_JP[td.weekday()]})"


def _build_navigator_df(
    rows: list[OptionStrikeRow],
    week: WeekDefinition,
) -> tuple[pd.DataFrame, list[int]]:
    """Build compact navigator: strike x (P出来高, P建玉, C出来高, C建玉) per day.

    Columns per day: P高{dd} P残{dd} | C残{dd} C高{dd}
    Returns (DataFrame, row_index_map).
    """
    col_order = ["行使価格"]
    for td in week.trading_days:
        dd = td.strftime("%d")
        col_order.append(f"P高{dd}")
        col_order.append(f"P残{dd}")
        col_order.append(f"C残{dd}")
        col_order.append(f"C高{dd}")

    records = []
    meta = []

    for idx, row in enumerate(rows):
        has_activity = (
            any(row.put_daily_volumes.get(td, 0) > 0 for td in week.trading_days)
            or any(row.call_daily_volumes.get(td, 0) > 0 for td in week.trading_days)
            or any(row.put_daily_oi.get(td, 0) > 0 for td in week.trading_days)
            or any(row.call_daily_oi.get(td, 0) > 0 for td in week.trading_days)
        )
        if not has_activity:
            continue

        rec = {"行使価格": f"{row.strike_price:,}"}
        for td in week.trading_days:
            dd = td.strftime("%d")
            p_vol = row.put_daily_volumes.get(td, 0)
            p_oi = row.put_daily_oi.get(td, 0)
            c_vol = row.call_daily_volumes.get(td, 0)
            c_oi = row.call_daily_oi.get(td, 0)
            rec[f"P高{dd}"] = p_vol if p_vol else None
            rec[f"P残{dd}"] = p_oi if p_oi else None
            rec[f"C残{dd}"] = c_oi if c_oi else None
            rec[f"C高{dd}"] = c_vol if c_vol else None
        records.append(rec)
        meta.append(idx)

    df = pd.DataFrame(records, columns=col_order)
    return df, meta


def _nav_column_config(df: pd.DataFrame) -> dict:
    cfg = {}
    cfg["行使価格"] = st.column_config.TextColumn("行使価格", width="small")
    for col in df.columns:
        if col == "行使価格":
            continue
        cfg[col] = st.column_config.NumberColumn(col, format="%d", width="small")
    return cfg


def _parse_nav_col(
    col_name: str,
    week: WeekDefinition,
) -> tuple[date | None, str | None]:
    """Parse navigator column name to (date, 'PUT'/'CALL')."""
    for td in week.trading_days:
        dd = td.strftime("%d")
        if col_name in (f"P高{dd}", f"P残{dd}"):
            return td, "PUT"
        if col_name in (f"C高{dd}", f"C残{dd}"):
            return td, "CALL"
    return None, None


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
        st.caption("ナビゲーターのセルをクリック")
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
