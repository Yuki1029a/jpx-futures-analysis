"""Option strike-price table with Styler rendering + clickable navigator.

Architecture:
  1. Main table (Styler via st.dataframe) — color-coded, formatted, with
     行使価格 as DataFrame index (pinned left in Streamlit's grid).
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

    # ====== 1. Main table with 行使価格 as index (pinned left) ======
    ordered_cols = _build_column_order(week)
    df = _build_display_dataframe(rows, week, ordered_cols)

    # Set 行使価格 as index → Streamlit pins index columns on left
    df = df.set_index("行使価格")

    styled = _apply_styling(df, week)
    st.dataframe(
        styled,
        use_container_width=True,
        height=min(len(df) * 35 + 60, 900),
    )

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
# Main table column order (excluding 行使価格 which becomes index)
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
    """Color coding + number formatting.

    Note: st.dataframe renders Styler with background colors intact.
    """
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

    # Index labels for summary rows
    summary_put_label = df.index[0]   # "PUT合計"
    summary_call_label = df.index[1]  # "CALL合計"

    def _cell_style(row_label, col, val):
        """Combined background + text color for a single cell."""
        parts = []

        # Background color
        if row_label == summary_put_label:
            parts.append(f"background-color: {_SUMMARY_PUT_BG}; font-weight: bold")
        elif row_label == summary_call_label:
            parts.append(f"background-color: {_SUMMARY_CALL_BG}; font-weight: bold")
        elif col in put_day_cols or col in put_week_oi:
            parts.append(f"background-color: {_PUT_BG}")
        elif col in put_jpx_cols:
            parts.append(f"background-color: {_JPX_BG_P}")
        elif col in put_oi_cols or col in put_chg_cols:
            parts.append(f"background-color: {_OI_BG_P}")
        elif col in call_day_cols or col in call_week_oi:
            parts.append(f"background-color: {_CALL_BG}")
        elif col in call_jpx_cols:
            parts.append(f"background-color: {_JPX_BG_C}")
        elif col in call_oi_cols or col in call_chg_cols:
            parts.append(f"background-color: {_OI_BG_C}")

        # Signed text color for OI change columns (non-summary rows)
        if (col in signed_cols
                and row_label != summary_put_label
                and row_label != summary_call_label):
            try:
                if pd.notna(val):
                    n = float(val)
                    if n > 0:
                        parts.append("color: #006100")
                    elif n < 0:
                        parts.append("color: #9c0006")
            except (ValueError, TypeError):
                pass

        return "; ".join(parts)

    def _apply_all_styles(s):
        """Apply background + text color for each cell in a row."""
        return [_cell_style(s.name, col, s[col]) for col in s.index]

    styled = df.style.apply(_apply_all_styles, axis=1)

    fmt_int = lambda v: f"{int(v):,}" if pd.notna(v) and v != "" else "-"
    fmt_signed = lambda v: f"{int(v):+,}" if pd.notna(v) and v != "" else "-"

    for col in df.columns:
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
    """Build compact navigator: strike x (P volume, P OI, C OI, C volume) per day.

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
