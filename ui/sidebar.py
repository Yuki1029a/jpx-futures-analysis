"""Streamlit sidebar controls for filtering and navigation."""
from __future__ import annotations

import streamlit as st
from data.aggregator import (
    build_available_weeks,
    get_available_contract_months,
    get_available_option_contract_months,
    get_option_participants,
)
import config


def _format_contract_month(cm: str) -> str:
    if not cm:
        return "-"
    return f"20{cm[:2]}年{cm[2:]}月限"


def render_sidebar() -> dict:
    """Render sidebar controls and return selections dict.

    Keys: product, week, contract_month,
          option_contract_month, option_participant_ids
    """
    st.sidebar.title("先物手口分析")

    # Product selector
    product = st.sidebar.selectbox(
        "商品",
        config.TARGET_PRODUCTS,
        format_func=lambda p: config.PRODUCT_DISPLAY_NAMES.get(p, p),
    )

    # Week selector
    weeks = build_available_weeks(max_weeks=26)
    if not weeks:
        st.sidebar.error("データが見つかりません")
        st.stop()

    week = st.sidebar.selectbox(
        "分析週",
        weeks,
        format_func=lambda w: w.label,
    )

    # Futures contract month selector
    contract_months = get_available_contract_months(week, product)
    if not contract_months or contract_months == [""]:
        st.sidebar.warning("限月データなし")
        st.stop()

    contract_month = st.sidebar.selectbox(
        "限月",
        contract_months,
        format_func=_format_contract_month,
    )

    # --- Option-specific controls ---
    st.sidebar.markdown("---")
    st.sidebar.subheader("オプション設定")

    # Option contract month
    opt_months = get_available_option_contract_months(week)
    option_contract_month = ""
    if opt_months:
        option_contract_month = st.sidebar.selectbox(
            "オプション限月",
            opt_months,
            format_func=_format_contract_month,
        )
    else:
        st.sidebar.info("オプション限月データなし")

    # Option participant filter (individual checkboxes)
    option_participant_ids = None  # None = all participants
    if option_contract_month:
        participants = get_option_participants(week, option_contract_month)
        if participants:
            with st.sidebar.expander("参加者フィルター", expanded=False):
                # Select all / Deselect all buttons
                btn_col1, btn_col2 = st.columns(2)
                with btn_col1:
                    if st.button("全選択", key="opt_sel_all"):
                        for pid, _ in participants:
                            st.session_state[f"opt_pid_{pid}"] = True
                with btn_col2:
                    if st.button("全解除", key="opt_desel_all"):
                        for pid, _ in participants:
                            st.session_state[f"opt_pid_{pid}"] = False

                # Individual checkboxes
                selected_pids = []
                for pid, name in participants:
                    key = f"opt_pid_{pid}"
                    if key not in st.session_state:
                        st.session_state[key] = True
                    checked = st.checkbox(name, key=key)
                    if checked:
                        selected_pids.append(pid)

                # Return None (all) if everyone is checked, else the list
                if len(selected_pids) == len(participants):
                    option_participant_ids = None
                elif selected_pids:
                    option_participant_ids = selected_pids
                else:
                    option_participant_ids = []  # empty = no data

    return {
        "product": product,
        "week": week,
        "contract_month": contract_month,
        "option_contract_month": option_contract_month,
        "option_participant_ids": option_participant_ids,
    }
