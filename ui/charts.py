"""Supplementary chart visualizations."""

import streamlit as st
import pandas as pd
from models import WeeklyParticipantRow, WeekDefinition


_DOW_JP = ["月", "火", "水", "木", "金", "土", "日"]


def render_net_change_bar_chart(
    rows: list[WeeklyParticipantRow],
    top_n: int = 15,
) -> None:
    """Horizontal bar chart: net OI change per participant."""
    data = [
        (r.participant_name, r.oi_net_change)
        for r in rows
        if r.oi_net_change is not None and r.oi_net_change != 0
    ]
    if not data:
        return

    data.sort(key=lambda x: x[1])  # ascending: sellers at top, buyers at bottom
    data = data[:top_n // 2] + data[-(top_n // 2):]  # top sellers + top buyers

    st.subheader("建玉増減ランキング")
    df = pd.DataFrame(data, columns=["参加者", "Net増減"])
    df = df.set_index("参加者")
    st.bar_chart(df)


def render_daily_volume_stacked(
    rows: list[WeeklyParticipantRow],
    week: WeekDefinition,
    top_n: int = 10,
) -> None:
    """Stacked bar chart: daily volumes by top participants."""
    top_rows = sorted(rows, key=lambda r: sum(r.daily_volumes.values()), reverse=True)[:top_n]

    if not top_rows or not week.trading_days:
        return

    st.subheader("日次売買高推移 (上位)")

    records = []
    for td in week.trading_days:
        dow = _DOW_JP[td.weekday()]
        date_label = f"{td.strftime('%m/%d')}({dow})"
        for r in top_rows:
            vol = r.daily_volumes.get(td, 0)
            if vol:
                # Shorten name for readability
                name = r.participant_name
                if len(name) > 20:
                    name = name[:18] + ".."
                records.append({
                    "日付": date_label,
                    "参加者": name,
                    "売買高": vol,
                })

    if records:
        df = pd.DataFrame(records)
        pivot = df.pivot_table(index="日付", columns="参加者", values="売買高", fill_value=0)
        st.bar_chart(pivot)
