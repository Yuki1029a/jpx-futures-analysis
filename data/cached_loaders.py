"""Streamlit-cached wrappers around aggregator functions.

Provides @st.cache_data layers on top of pure aggregator functions so that
re-running Streamlit (widget change, multi-user, etc.) doesn't re-parse the
same Excel files. TTL = 5 minutes.

WeekDefinition is not directly hashable (contains list[date]); callers should
pass a stable key obtained via `wk_key(week)`.
"""
from __future__ import annotations

import streamlit as st

from models import WeekDefinition
from data.aggregator import (
    load_weekly_data,
    compute_20d_stats,
    load_daily_futures_oi,
    load_option_weekly_data,
    load_put_call_daily_volumes,
    get_available_contract_months,
    get_available_option_contract_months,
    get_option_participants,
    SESSION_MODES,
)


_TTL = 300  # seconds (5 min). JPX intraday data refreshes a few times a day.


def wk_key(week: WeekDefinition):
    """Convert WeekDefinition into a hashable key for cache lookups."""
    return (week.start_oi_date, week.end_oi_date, tuple(week.trading_days))


def _reconstruct(week_key) -> WeekDefinition:
    start_oi, end_oi, td_tuple = week_key
    return WeekDefinition(
        start_oi_date=start_oi,
        end_oi_date=end_oi,
        trading_days=list(td_tuple),
        label="",
    )


# ---------------------------------------------------------------------------
# Sidebar dropdowns (cheap functions, but called on every rerun)
# ---------------------------------------------------------------------------

@st.cache_data(ttl=_TTL, show_spinner=False)
def cached_contract_months(week_key, product):
    return get_available_contract_months(_reconstruct(week_key), product)


@st.cache_data(ttl=_TTL, show_spinner=False)
def cached_option_contract_months(week_key):
    return get_available_option_contract_months(_reconstruct(week_key))


@st.cache_data(ttl=_TTL, show_spinner=False)
def cached_option_participants(week_key, contract_month):
    return get_option_participants(_reconstruct(week_key), contract_month)


# ---------------------------------------------------------------------------
# Heavy loaders (these are the real bottleneck on week change)
# ---------------------------------------------------------------------------

@st.cache_data(ttl=_TTL, show_spinner=False)
def cached_weekly_data(week_key, product, contract_month, sk_str):
    """Cached load_weekly_data. sk_str is the SESSION_MODES key (e.g. '全セッション合計')."""
    return load_weekly_data(
        _reconstruct(week_key), product, contract_month,
        session_keys=SESSION_MODES[sk_str], include_oi=True,
    )


@st.cache_data(ttl=_TTL, show_spinner=False)
def cached_20d_stats(week_key, product, contract_month, sk_str):
    return compute_20d_stats(
        _reconstruct(week_key), product, contract_month,
        session_keys=SESSION_MODES[sk_str],
    )


@st.cache_data(ttl=_TTL, show_spinner=False)
def cached_daily_futures_oi(week_key, product, contract_month):
    return load_daily_futures_oi(_reconstruct(week_key), product, contract_month)


@st.cache_data(ttl=_TTL, show_spinner=False)
def cached_option_weekly_data(week_key, contract_month, sk_str, pid_str):
    """pid_str is either 'ALL', '' (empty list), or comma-separated participant ids."""
    if pid_str == "ALL":
        pids = None
    elif pid_str == "":
        pids = []
    else:
        pids = pid_str.split(",")
    return load_option_weekly_data(
        _reconstruct(week_key),
        contract_month=contract_month,
        session_keys=SESSION_MODES[sk_str],
        participant_ids=pids,
    )


@st.cache_data(ttl=_TTL, show_spinner=False)
def cached_put_call_daily_volumes(week_key, contract_month):
    return load_put_call_daily_volumes(_reconstruct(week_key), contract_month)
