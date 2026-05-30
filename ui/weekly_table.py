"""Main weekly table visualization component."""
from __future__ import annotations

import streamlit as st
import pandas as pd
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
    extra_oi_series: list | None = None,
) -> None:
    """Render the main weekly analysis table.

    Args:
        show_oi: If True, include OI columns.
        tab_label: Label shown in subheader for session context.
        stats_20d: {participant_id: (avg, max)}.
        daily_futures_oi: {date: DailyFuturesOI} - primary contract (期近 in SQ week).
        extra_oi_series: list of (label_suffix, {date: DailyFuturesOI}) -
                         additional contracts (e.g. 期先 in SQ week).
                         Each adds 2 header rows (建玉残高 + 前日比).
    """
    cm_label = f"20{contract_month[:2]}年{contract_month[2:]}月限" if contract_month else ""
    title = f"{config.PRODUCT_DISPLAY_NAMES.get(product, product)} {cm_label}  ({week.label})"
    if tab_label and tab_label != "全セッション合計":
        title += f"  [{tab_label}]"
    st.subheader(title)

    if not rows:
        st.warning("選択された条件のデータがありません。")
        return

    # Build list of (label, oi_dict) ─ primary first, then extras
    oi_series: list = []
    if daily_futures_oi:
        primary_label = "期近" if extra_oi_series else ""
        oi_series.append((primary_label, daily_futures_oi))
    if extra_oi_series:
        for lbl, d in extra_oi_series:
            if d:
                oi_series.append((lbl, d))

    df = _build_display_dataframe(rows, week, show_oi, stats_20d, oi_series)

    oi_header_rows = len(oi_series) * 2  # 2 rows per contract (建玉残高 + 前日比)

    styled = _apply_table_styling(df, week, show_oi, oi_header_rows)

    st.dataframe(
        styled,
        use_container_width=True,
        height=min((len(rows) + oi_header_rows) * 35 + 60, 900),
    )

    if show_oi:
        _render_summary_stats(rows)


def _build_display_dataframe(
    rows: list[WeeklyParticipantRow],
    week: WeekDefinition,
    show_oi: bool,
    stats_20d: dict | None = None,
    oi_series: list | None = None,
) -> pd.DataFrame:
    """Build the display DataFrame.

    oi_series: list of (label_suffix, {date: DailyFuturesOI}). For each entry,
    two header rows are emitted: 建玉残高[label] and 前日比[label].
    """
    day_col_names = []
    for td in week.trading_days:
        dow = _DOW_JP[td.weekday()]
        day_col_names.append(f"{td.strftime('%m/%d')}({dow})")

    def _blank_row(label: str) -> dict:
        rec = {"参加者": label}
        if show_oi:
            rec["前週L"] = None
            rec["前週S"] = None
        for col in day_col_names:
            rec[col] = None
        rec["週間計"] = None
        rec["20日平均"] = None
        rec["20日最大"] = None
        if show_oi:
            rec["今週L"] = None
            rec["今週S"] = None
            rec["増減"] = None
            rec["推定買"] = None
            rec["推定売"] = None
            rec["方向"] = ""
        return rec

    # --- OI header rows (2 per contract) ---
    oi_rows = []
    if oi_series:
        for label_suffix, oi_dict in oi_series:
            suffix = f"（{label_suffix}）" if label_suffix else ""

            # 建玉残高 row
            rec_oi = _blank_row(f"建玉残高{suffix}")
            for td, col_name in zip(week.trading_days, day_col_names):
                rec = oi_dict.get(td)
                rec_oi[col_name] = rec.current_oi if rec else None
            oi_rows.append(rec_oi)

            # 前日比 row
            rec_chg = _blank_row(f"前日比{suffix}")
            for td, col_name in zip(week.trading_days, day_col_names):
                rec = oi_dict.get(td)
                rec_chg[col_name] = rec.net_change if rec else None
            oi_rows.append(rec_chg)

    # --- Participant rows ---
    records = []
    for row in rows:
        rec = {"参加者": row.participant_name}

        if show_oi:
            rec["前週L"] = row.start_oi_long
            rec["前週S"] = row.start_oi_short

        weekly_total = 0.0
        for td, col_name in zip(week.trading_days, day_col_names):
            vol = row.daily_volumes.get(td)
            rec[col_name] = vol if vol else None
            if vol:
                weekly_total += vol

        rec["週間計"] = weekly_total if weekly_total > 0 else None

        avg_20d = None
        max_20d = None
        if stats_20d and row.participant_id in stats_20d:
            avg_20d, max_20d = stats_20d[row.participant_id]
        rec["20日平均"] = round(avg_20d) if avg_20d is not None else None
        rec["20日最大"] = round(max_20d) if max_20d is not None else None

        if show_oi:
            rec["今週L"] = row.end_oi_long
            rec["今週S"] = row.end_oi_short
            rec["増減"] = row.oi_net_change

            est_buy = None
            est_sell = None
            if row.oi_net_change is not None and weekly_total > 0:
                est_buy = (weekly_total + row.oi_net_change) / 2
                est_sell = (weekly_total - row.oi_net_change) / 2
            rec["推定買"] = est_buy
            rec["推定売"] = est_sell
            rec["方向"] = _direction_label(row.inferred_direction)

        records.append(rec)

    return pd.DataFrame(oi_rows + records)


def _apply_table_styling(
    df: pd.DataFrame, week: WeekDefinition, show_oi: bool,
    oi_header_rows: int = 0,
):
    """Apply conditional formatting and number formatting."""
    day_cols = [f"{td.strftime('%m/%d')}({_DOW_JP[td.weekday()]})" for td in week.trading_days]
    int_cols = list(day_cols) + ["週間計", "20日平均", "20日最大"]

    signed_cols = []  # columns that should show +/- sign
    long_cols = set()
    short_cols = set()
    if show_oi:
        oi_cols = ["前週L", "前週S", "今週L", "今週S"]
        signed_cols = ["増減"]
        est_cols = ["推定買", "推定売"]
        int_cols += oi_cols + signed_cols + est_cols
        long_cols = {"前週L", "今週L"}
        short_cols = {"前週S", "今週S"}

    # -- Color palette for L/S --
    _LONG_BG = "#dceefb"   # soft blue for Long
    _SHORT_BG = "#fde8e8"  # soft red for Short

    # Color functions — only apply to participant rows (skip OI header rows)
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

    # Style OI header rows. Each contract contributes 2 consecutive rows:
    # even-offset (0,2,...) = 建玉残高 (bold blue bg), odd-offset = 前日比 (signed color)
    def _style_oi_header(row_idx, val, col):
        offset = row_idx % 2
        if offset == 0:
            # 建玉残高
            if col in day_cols and pd.notna(val):
                return "background-color: #dae8fc; color: #1a3c5e; font-weight: bold"
            return "background-color: #dae8fc; color: #1a3c5e"
        else:
            # 前日比
            if col in day_cols:
                return _color_signed(val)
            return ""

    styled = df.style

    # Apply OI header row styling
    if oi_header_rows > 0:
        def _apply_oi_style(s):
            styles = [""] * len(s)
            row_idx = s.name
            if row_idx < oi_header_rows:
                for i, (col, val) in enumerate(s.items()):
                    styles[i] = _style_oi_header(row_idx, val, col)
            return styles
        styled = styled.apply(_apply_oi_style, axis=1)

    # Apply L/S background color to participant rows
    if long_cols or short_cols:
        participant_idx = list(range(oi_header_rows, len(df)))
        if participant_idx:
            for col in long_cols:
                if col in df.columns:
                    styled = styled.map(
                        lambda v: f"background-color: {_LONG_BG}" if pd.notna(v) and v != "" else "",
                        subset=(participant_idx, [col]),
                    )
            for col in short_cols:
                if col in df.columns:
                    styled = styled.map(
                        lambda v: f"background-color: {_SHORT_BG}" if pd.notna(v) and v != "" else "",
                        subset=(participant_idx, [col]),
                    )

    # Apply sign-based coloring to signed columns (participant rows only)
    if signed_cols:
        participant_idx = list(range(oi_header_rows, len(df)))
        for col in signed_cols:
            if col in df.columns and participant_idx:
                styled = styled.map(
                    _color_signed,
                    subset=(participant_idx, [col]),
                )

    # Color direction column (participant rows only)
    if show_oi and "方向" in df.columns:
        participant_idx = list(range(oi_header_rows, len(df)))
        if participant_idx:
            styled = styled.map(
                _color_direction,
                subset=(participant_idx, ["方向"]),
            )

    # Number formatting
    fmt_int = lambda v: f"{int(v):,}" if pd.notna(v) and v != "" else "-"
    fmt_signed = lambda v: f"{int(v):+,}" if pd.notna(v) and v != "" else "-"

    for col in df.columns:
        if col == "参加者" or col == "方向":
            continue
        if col in signed_cols:
            styled = styled.format(fmt_signed, subset=[col])
        elif col in int_cols:
            styled = styled.format(fmt_int, subset=[col])

    return styled


def _render_summary_stats(rows: list[WeeklyParticipantRow]) -> None:
    """Render summary metrics below the table."""
    st.markdown("---")
    cols = st.columns(4)

    total_vol = sum(sum(r.daily_volumes.values()) for r in rows)
    oi_available = any(r.oi_net_change is not None for r in rows)

    if oi_available:
        buyers = sum(1 for r in rows if r.inferred_direction == "BUY")
        sellers = sum(1 for r in rows if r.inferred_direction == "SELL")
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
    else:
        with cols[0]:
            st.metric("週間出来高計", f"{int(total_vol):,}")
        with cols[1]:
            st.metric("買い方", "未発表")
        with cols[2]:
            st.metric("売り方", "未発表")
        with cols[3]:
            st.metric("全体Net増減", "未発表")


def _direction_label(direction) -> str:
    if direction == "BUY":
        return "BUY"
    elif direction == "SELL":
        return "SELL"
    elif direction == "NEUTRAL":
        return "-"
    return ""
