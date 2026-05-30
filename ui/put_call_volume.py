"""PUT/CALL daily volume tracking — simple aggregate per-day table.

Data source: open_interest_e.xlsx Attachment1 (parse_daily_oi_excel),
trading_volume per strike summed by option_type and contract month.
"""
from __future__ import annotations

import streamlit as st
import pandas as pd
from models import WeekDefinition
from data.cached_loaders import (
    wk_key,
    cached_option_contract_months,
    cached_put_call_daily_volumes,
)


_DOW_JP = ["月", "火", "水", "木", "金", "土", "日"]


def render_put_call_volume_section(week: WeekDefinition) -> None:
    """Render daily PUT/CALL aggregate volume table.

    Layout:
      - Selector: All contract months / specific contract month
      - Table: rows = [PUT vol, CALL vol, PUT/CALL, PUT OI, CALL OI],
               columns = trading days, plus 週間計
    """
    st.subheader(f"PUT/CALL 日次出来高  ({week.label})")

    wk = wk_key(week)
    contract_months = cached_option_contract_months(wk)
    if not contract_months:
        st.info("オプションデータなし")
        return

    options = ["全限月合算"] + [
        f"20{cm[:2]}年{cm[2:]}月限" for cm in contract_months
    ]
    selected = st.selectbox("限月", options, key="pcv_cm")

    if selected == "全限月合算":
        target_cm = ""
        target_label = "全限月合算"
    else:
        idx = options.index(selected) - 1
        target_cm = contract_months[idx]
        target_label = selected

    with st.spinner("PUT/CALL データ集計中..."):
        daily = cached_put_call_daily_volumes(wk, target_cm)

    if not daily:
        st.info("該当データなし")
        return

    # Build columns
    day_col_names = []
    for td in week.trading_days:
        dow = _DOW_JP[td.weekday()]
        day_col_names.append(f"{td.strftime('%m/%d')}({dow})")

    # Build rows
    put_vol_row = {"指標": "PUT 出来高"}
    call_vol_row = {"指標": "CALL 出来高"}
    pc_ratio_row = {"指標": "P/C 比"}
    put_oi_row = {"指標": "PUT 建玉"}
    call_oi_row = {"指標": "CALL 建玉"}

    week_put_vol = 0
    week_call_vol = 0

    for td, col_name in zip(week.trading_days, day_col_names):
        d = daily.get(td)
        if d is None:
            put_vol_row[col_name] = None
            call_vol_row[col_name] = None
            pc_ratio_row[col_name] = None
            put_oi_row[col_name] = None
            call_oi_row[col_name] = None
            continue
        pv, cv = d["PUT"], d["CALL"]
        put_vol_row[col_name] = pv if pv else None
        call_vol_row[col_name] = cv if cv else None
        pc_ratio_row[col_name] = (pv / cv) if cv > 0 else None
        put_oi_row[col_name] = d["PUT_OI"] if d["PUT_OI"] else None
        call_oi_row[col_name] = d["CALL_OI"] if d["CALL_OI"] else None
        week_put_vol += pv
        week_call_vol += cv

    # Week total column
    put_vol_row["週間計"] = week_put_vol if week_put_vol else None
    call_vol_row["週間計"] = week_call_vol if week_call_vol else None
    pc_ratio_row["週間計"] = (
        week_put_vol / week_call_vol if week_call_vol > 0 else None
    )
    put_oi_row["週間計"] = None
    call_oi_row["週間計"] = None

    df = pd.DataFrame(
        [put_vol_row, call_vol_row, pc_ratio_row, put_oi_row, call_oi_row]
    )

    # Cell-by-cell formatting: P/C 比 row uses 2-decimal float, others use comma-int
    def _fmt_int(v):
        return f"{int(v):,}" if pd.notna(v) else "-"

    def _fmt_ratio(v):
        return f"{float(v):.2f}" if pd.notna(v) else "-"

    display_df = df.copy()
    for c in display_df.columns:
        if c == "指標":
            continue
        display_df[c] = df.apply(
            lambda r, col=c: (
                _fmt_ratio(r[col]) if r["指標"] == "P/C 比" else _fmt_int(r[col])
            ),
            axis=1,
        )

    # Color P/C row by sentiment (>1.2 put-heavy red, <0.83 call-heavy blue)
    def _style_pc(row):
        if row["指標"] != "P/C 比":
            return [""] * len(row)
        styles = []
        for c, v_str in row.items():
            if c == "指標":
                styles.append("font-weight: bold")
                continue
            try:
                v = float(v_str)
                if v > 1.2:
                    styles.append("background-color: #ffe0e0")
                elif v < 0.83:
                    styles.append("background-color: #e0e8ff")
                else:
                    styles.append("")
            except (ValueError, TypeError):
                styles.append("")
        return styles

    styled = display_df.style.apply(_style_pc, axis=1)
    st.dataframe(styled, use_container_width=True, height=240)

    # Summary metrics
    if week_put_vol > 0 or week_call_vol > 0:
        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric("週間 PUT 計", f"{week_put_vol:,}")
        with c2:
            st.metric("週間 CALL 計", f"{week_call_vol:,}")
        with c3:
            if week_call_vol > 0:
                st.metric("週間 P/C 比", f"{week_put_vol / week_call_vol:.2f}")

    st.caption(f"対象: {target_label}  /  ソース: open_interest_e.xlsx Attachment1")
