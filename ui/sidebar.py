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

    # Option participant filter
    option_participant_ids = None  # None = all participants
    if option_contract_month:
        participants = get_option_participants(week, option_contract_month)
        if participants:
            with st.sidebar.expander("参加者フィルター", expanded=False):
                select_all = st.checkbox("全員選択", value=True, key="opt_select_all")
                if select_all:
                    option_participant_ids = None  # all
                else:
                    selected = st.multiselect(
                        "参加者を選択",
                        options=[pid for pid, _ in participants],
                        default=[pid for pid, _ in participants],
                        format_func=lambda pid: dict(participants).get(pid, pid),
                    )
                    option_participant_ids = selected if selected else None

    return {
        "product": product,
        "week": week,
        "contract_month": contract_month,
        "option_contract_month": option_contract_month,
        "option_participant_ids": option_participant_ids,
    }
