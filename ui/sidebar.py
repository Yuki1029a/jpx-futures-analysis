"""Streamlit sidebar controls for filtering and navigation."""

import streamlit as st
from data.aggregator import build_available_weeks, get_available_contract_months
import config


def render_sidebar() -> dict:
    """Render sidebar controls and return selections dict.

    Keys: product, week, contract_month
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

    # Contract month selector
    contract_months = get_available_contract_months(week, product)
    if not contract_months or contract_months == [""]:
        st.sidebar.warning("限月データなし")
        st.stop()

    contract_month = st.sidebar.selectbox(
        "限月",
        contract_months,
        format_func=lambda cm: f"20{cm[:2]}年{cm[2:]}月限" if cm else "-",
    )

    return {
        "product": product,
        "week": week,
        "contract_month": contract_month,
    }
