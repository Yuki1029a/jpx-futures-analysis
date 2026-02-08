"""JPX Futures Participant Analysis - Main Streamlit App.

Run with: streamlit run app.py
"""

import streamlit as st
from data.cache import ensure_cache_dirs
from data.aggregator import load_weekly_data, compute_20d_stats, SESSION_MODES
from ui.sidebar import render_sidebar
from ui.weekly_table import render_weekly_table
from ui.charts import render_net_change_bar_chart, render_daily_volume_stacked
from models import WeekDefinition

st.set_page_config(
    page_title="先物手口分析",
    layout="wide",
    initial_sidebar_state="expanded",
)

ensure_cache_dirs()


@st.cache_data(ttl=1800, show_spinner=False)
def _cached_weekly_data(
    start_oi_str: str, end_oi_str: str, trading_days_str: str,
    product: str, contract_month: str, session_key_str: str,
):
    """Cache wrapper for load_weekly_data. Keys are plain strings for hashability."""
    from datetime import datetime, date
    start_oi = datetime.strptime(start_oi_str, "%Y%m%d").date()
    end_oi = datetime.strptime(end_oi_str, "%Y%m%d").date() if end_oi_str else None
    tdays = [datetime.strptime(s, "%Y%m%d").date() for s in trading_days_str.split(",") if s]
    session_keys = "ALL" if session_key_str == "ALL" else tuple(session_key_str.split(","))

    week = WeekDefinition(
        start_oi_date=start_oi, end_oi_date=end_oi,
        trading_days=tdays, label="",
    )
    return load_weekly_data(week, product, contract_month,
                            session_keys=session_keys, include_oi=True)


@st.cache_data(ttl=1800, show_spinner=False)
def _cached_20d_stats(
    start_oi_str: str, end_oi_str: str, trading_days_str: str,
    product: str, contract_month: str, session_key_str: str,
):
    """Cache wrapper for compute_20d_stats."""
    from datetime import datetime
    start_oi = datetime.strptime(start_oi_str, "%Y%m%d").date()
    end_oi = datetime.strptime(end_oi_str, "%Y%m%d").date() if end_oi_str else None
    tdays = [datetime.strptime(s, "%Y%m%d").date() for s in trading_days_str.split(",") if s]
    session_keys = "ALL" if session_key_str == "ALL" else tuple(session_key_str.split(","))

    week = WeekDefinition(
        start_oi_date=start_oi, end_oi_date=end_oi,
        trading_days=tdays, label="",
    )
    return compute_20d_stats(week, product, contract_month,
                             session_keys=session_keys)


def _week_to_cache_keys(week: WeekDefinition):
    """Convert WeekDefinition to hashable strings for cache key."""
    return (
        week.start_oi_date.strftime("%Y%m%d"),
        week.end_oi_date.strftime("%Y%m%d") if week.end_oi_date else "",
        ",".join(d.strftime("%Y%m%d") for d in week.trading_days),
    )


def _session_to_str(session_keys):
    if session_keys == "ALL":
        return "ALL"
    return ",".join(session_keys)


def main():
    selections = render_sidebar()

    product = selections["product"]
    week = selections["week"]
    contract_month = selections["contract_month"]

    wk_start, wk_end, wk_tdays = _week_to_cache_keys(week)

    # Session tabs
    tab_labels = list(SESSION_MODES.keys())
    tabs = st.tabs(tab_labels)

    for tab, label in zip(tabs, tab_labels):
        with tab:
            session_keys = SESSION_MODES[label]
            is_total = label == "全セッション合計"
            sk_str = _session_to_str(session_keys)

            with st.spinner("データ読み込み中..."):
                rows = _cached_weekly_data(
                    wk_start, wk_end, wk_tdays,
                    product, contract_month, sk_str,
                )
                stats_20d = _cached_20d_stats(
                    wk_start, wk_end, wk_tdays,
                    product, contract_month, sk_str,
                )

            if not rows:
                st.info("該当データなし")
                continue

            render_weekly_table(
                rows, week, product, contract_month,
                show_oi=True,
                tab_label=label,
                stats_20d=stats_20d,
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
