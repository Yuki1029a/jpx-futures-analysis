"""JPX Futures & Options Participant Analysis - Main Streamlit App.

Run with: streamlit run app.py
"""
from __future__ import annotations

import streamlit as st
from data.cache import ensure_cache_dirs
from data.aggregator import (
    load_weekly_data, compute_20d_stats,
    load_option_weekly_data, load_daily_futures_oi, SESSION_MODES,
)
from ui.sidebar import render_sidebar
from ui.weekly_table import render_weekly_table
from ui.charts import render_net_change_bar_chart, render_daily_volume_stacked
from ui.option_strike_table import render_option_strike_table
from ui.gex_chart import render_gex_section

st.set_page_config(
    page_title="先物手口分析",
    layout="wide",
    initial_sidebar_state="expanded",
)

ensure_cache_dirs()

# Debug: deployment version
_APP_VERSION = "2026-02-13b"


def _make_cache_key(product, contract_month, wk_label, sk_str, kind):
    """Build a unique string key for session_state caching."""
    return f"{kind}|{product}|{contract_month}|{wk_label}|{sk_str}"


def _get_or_load(product, contract_month, week, sk_str, session_keys):
    """Load futures data using session_state as cache."""
    key_rows = _make_cache_key(product, contract_month, week.label, sk_str, "rows")
    key_stats = _make_cache_key(product, contract_month, week.label, sk_str, "stats")

    if key_rows not in st.session_state:
        st.session_state[key_rows] = load_weekly_data(
            week, product, contract_month,
            session_keys=session_keys, include_oi=True,
        )
    if key_stats not in st.session_state:
        st.session_state[key_stats] = compute_20d_stats(
            week, product, contract_month,
            session_keys=session_keys,
        )
    return st.session_state[key_rows], st.session_state[key_stats]


def _get_or_load_options(week, contract_month, sk_str, session_keys, participant_ids):
    """Load option data using session_state as cache."""
    pid_str = ",".join(sorted(participant_ids)) if participant_ids is not None else "ALL"
    key = f"opt_rows_v2|{week.label}|{contract_month}|{sk_str}|{pid_str}"
    if key not in st.session_state:
        st.session_state[key] = load_option_weekly_data(
            week,
            contract_month=contract_month,
            session_keys=session_keys,
            participant_ids=participant_ids,
        )
    return st.session_state[key]


def main():
    st.sidebar.caption(f"v{_APP_VERSION}")
    selections = render_sidebar()

    product = selections["product"]
    week = selections["week"]
    contract_month = selections["contract_month"]
    opt_cm = selections["option_contract_month"]
    opt_pids = selections["option_participant_ids"]

    # Top-level tabs: Futures, Options, GEX
    main_tab1, main_tab2, main_tab3 = st.tabs(["先物分析", "オプション分析", "GEX分析"])

    with main_tab1:
        _render_futures_section(product, week, contract_month)

    with main_tab2:
        _render_options_section(week, opt_cm, opt_pids)

    with main_tab3:
        _render_gex_section(week, opt_cm, opt_pids)


def _render_futures_section(product, week, contract_month):
    """Render the futures analysis tabs."""
    tab_labels = list(SESSION_MODES.keys())
    tabs = st.tabs(tab_labels)

    # Load daily futures OI once (same data for all session tabs)
    daily_fut_oi = load_daily_futures_oi(week, product, contract_month)

    for tab, label in zip(tabs, tab_labels):
        with tab:
            session_keys = SESSION_MODES[label]
            is_total = label == "全セッション合計"
            sk_str = label

            with st.spinner("データ読み込み中..."):
                rows, stats_20d = _get_or_load(
                    product, contract_month, week, sk_str, session_keys,
                )

            if not rows:
                st.info("該当データなし")
                continue

            render_weekly_table(
                rows, week, product, contract_month,
                show_oi=True,
                tab_label=label,
                stats_20d=stats_20d,
                daily_futures_oi=daily_fut_oi,
            )

            if is_total:
                st.markdown("---")
                col1, col2 = st.columns(2)
                with col1:
                    render_net_change_bar_chart(rows)
                with col2:
                    render_daily_volume_stacked(rows, week)


def _render_options_section(week, opt_cm, opt_pids):
    """Render the options analysis tabs."""
    if not opt_cm:
        st.info("オプション限月を選択してください")
        return

    tab_labels = list(SESSION_MODES.keys())
    tabs = st.tabs(tab_labels)

    for tab, label in zip(tabs, tab_labels):
        with tab:
            session_keys = SESSION_MODES[label]
            sk_str = label

            with st.spinner("オプションデータ読み込み中..."):
                opt_rows = _get_or_load_options(
                    week, opt_cm, sk_str, session_keys, opt_pids,
                )

            if not opt_rows:
                st.info("オプションデータなし")
                continue

            render_option_strike_table(opt_rows, week, tab_label=label)


def _render_gex_section(week, opt_cm, opt_pids):
    """Render GEX analysis tab."""
    if not opt_cm:
        st.info("サイドバーでオプション限月を選択してください")
        return

    # Load option data (全セッション合計, no participant filter for aggregate OI)
    with st.spinner("GEXデータ読み込み中..."):
        opt_rows = _get_or_load_options(
            week, opt_cm, "全セッション合計",
            SESSION_MODES["全セッション合計"], opt_pids,
        )

    if not opt_rows:
        st.info("オプションデータなし")
        return

    render_gex_section(opt_rows, week, opt_cm)


if __name__ == "__main__":
    main()
