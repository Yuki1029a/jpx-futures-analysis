"""JPX Futures Participant Analysis - Main Streamlit App.

Run with: streamlit run app.py
"""

import streamlit as st
from data.cache import ensure_cache_dirs
from data.aggregator import load_weekly_data, SESSION_MODES
from ui.sidebar import render_sidebar
from ui.weekly_table import render_weekly_table
from ui.charts import render_net_change_bar_chart, render_daily_volume_stacked

st.set_page_config(
    page_title="先物手口分析",
    layout="wide",
    initial_sidebar_state="expanded",
)

ensure_cache_dirs()


def main():
    selections = render_sidebar()

    product = selections["product"]
    week = selections["week"]
    contract_month = selections["contract_month"]

    # Session tabs
    tab_labels = list(SESSION_MODES.keys())
    tabs = st.tabs(tab_labels)

    for tab, label in zip(tabs, tab_labels):
        with tab:
            session_keys = SESSION_MODES[label]
            is_total = label == "全セッション合計"

            with st.spinner("データ読み込み中..."):
                rows = load_weekly_data(
                    week, product, contract_month,
                    session_keys=session_keys,
                    include_oi=True,
                )

            if not rows:
                st.info("該当データなし")
                continue

            render_weekly_table(
                rows, week, product, contract_month,
                show_oi=True,
                tab_label=label,
            )

            if is_total:
                st.markdown("---")
                col1, col2 = st.columns(2)
                with col1:
                    render_net_change_bar_chart(rows)
                with col2:
                    render_daily_volume_stacked(rows, week)


if __name__ == "__main__":
    main()
