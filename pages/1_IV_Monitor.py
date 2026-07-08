"""IV Monitor — QRI NK225 option IV snapshots (15-min delayed) viewer.

Data: R2 qri_iv/raw/YYYYMMDD/*.parquet (scripts/fetch_qri_iv.py が収集)
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from data import iv_views
from data.qri_iv_loader import load_snapshot

st.set_page_config(page_title="IVモニター", layout="wide")

_TTL = 600  # R2は15分毎更新のため10分キャッシュ


@st.cache_data(ttl=_TTL, show_spinner=False)
def _days():
    return iv_views.list_available_days()


@st.cache_data(ttl=_TTL, show_spinner=False)
def _keys(day_str: str):
    return iv_views.snapshot_keys(day_str)


@st.cache_data(ttl=_TTL, show_spinner=False)
def _snapshot(key: str) -> pd.DataFrame:
    df = load_snapshot(key)
    return pd.DataFrame() if df is None else df


@st.cache_data(ttl=_TTL, show_spinner=False)
def _intraday(day_str: str) -> pd.DataFrame:
    return iv_views.intraday_atm(day_str)


@st.cache_data(ttl=1800, show_spinner=False)
def _daily(days: tuple) -> pd.DataFrame:
    return iv_views.daily_atm_series(list(days))


_MONTH_COLORS = ["#14285a", "#ba7517", "#1d9e75", "#993556", "#534ab7", "#888780"]


def _fmt_cm(cm: str) -> str:
    return f"20{cm[:2]}年{cm[2:]}月限" if len(cm) == 4 else cm


def main() -> None:
    st.title("IVモニター（日経225オプション）")
    st.caption("出所: QRI（15分ディレイ）スナップショット / IV=気配ミッド優先"
               "（両気配欠損時のみ約定・清算IV） / %表示")

    days = _days()
    if not days:
        st.error("qri_iv データが見つかりません（R2/ローカルとも空）")
        st.stop()

    # ---------------- sidebar ----------------
    st.sidebar.title("IVモニター")
    day = st.sidebar.selectbox("日付", list(reversed(days)),
                               format_func=lambda d: f"{d[:4]}/{d[4:6]}/{d[6:8]}")
    keys = _keys(day)
    if not keys:
        st.warning("この日のスナップショットがありません")
        st.stop()
    key = st.sidebar.selectbox("スナップショット", list(reversed(keys)),
                               format_func=iv_views.key_time_label)

    df = _snapshot(key)
    if df.empty:
        st.warning("スナップショットを読み込めませんでした")
        st.stop()

    all_months = sorted(df.contract_month.unique())
    months = st.sidebar.multiselect("限月", all_months, default=all_months,
                                    format_func=_fmt_cm)
    side = st.sidebar.radio("スマイル表示", ["CALL", "PUT", "両方"], horizontal=True)
    n_hist = st.sidebar.slider("日次推移の対象日数", 5, len(days), min(20, len(days)))

    upd = df.source_update_time.dropna()
    st.caption(f"スナップショット: {iv_views.key_time_label(key)}  |  "
               f"ソース更新: {upd.iloc[0] if len(upd) else '-'}  |  行数: {len(df):,}")

    # ---------------- summary ----------------
    summ = iv_views.month_summary(df)
    cols = st.columns(max(len(summ), 1))
    for c, (_, r) in zip(cols, summ.iterrows()):
        with c:
            atm = "-" if pd.isna(r.atm_iv) else f"{r.atm_iv * 100:.1f}%"
            sk = "-" if pd.isna(r.skew25) else f"{r.skew25 * 100:+.1f}%"
            st.metric(_fmt_cm(r.contract_month),
                      f"ATM {atm}",
                      f"25Δスキュー {sk}", delta_color="off")
            st.caption(f"ATM行使 {r.atm_strike:,.0f} | OI計 {r.oi_total:,.0f}"
                       if pd.notna(r.atm_strike) else "")

    # ---------------- smile ----------------
    st.subheader("IVスマイル")
    sm = iv_views.smile_frame(df, months)
    fig = go.Figure()
    for mi, cm in enumerate(months):
        col = _MONTH_COLORS[mi % len(_MONTH_COLORS)]
        for ot, dashv in (("CALL", "solid"), ("PUT", "dot")):
            if side != "両方" and ot != side:
                continue
            s = sm[(sm.contract_month == cm) & (sm.option_type == ot)]
            if not len(s):
                continue
            fig.add_trace(go.Scatter(
                x=s.strike, y=s.iv_pct, mode="lines+markers",
                name=f"{_fmt_cm(cm)} {ot}",
                line=dict(color=col, dash=dashv, width=1.6),
                marker=dict(size=4),
            ))
        summ_row = summ[summ.contract_month == cm]
        if len(summ_row) and pd.notna(summ_row.iloc[0].atm_strike):
            fig.add_vline(x=float(summ_row.iloc[0].atm_strike),
                          line=dict(color=col, width=0.8, dash="dash"))
    fig.update_layout(height=420, template="plotly_white",
                      xaxis_title="行使価格", yaxis_title="IV (%)",
                      legend=dict(orientation="h", y=-0.18),
                      margin=dict(l=0, r=0, t=10, b=0))
    st.plotly_chart(fig, use_container_width=True)
    st.caption("実線=CALL / 点線=PUT / 縦破線=ATM行使価格")

    # ---------------- intraday ATM ----------------
    st.subheader("日内 ATM IV 推移")
    intr = _intraday(day)
    if len(intr):
        fig2 = go.Figure()
        for mi, cm in enumerate([m for m in months if m in set(intr.contract_month)]):
            s = intr[intr.contract_month == cm]
            fig2.add_trace(go.Scatter(
                x=s.time, y=s.atm_iv_pct, mode="lines+markers",
                name=_fmt_cm(cm),
                line=dict(color=_MONTH_COLORS[mi % len(_MONTH_COLORS)], width=1.8),
            ))
        fig2.update_layout(height=300, template="plotly_white",
                           xaxis_title="スナップショット時刻", yaxis_title="ATM IV (%)",
                           legend=dict(orientation="h", y=-0.3),
                           margin=dict(l=0, r=0, t=10, b=0))
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.info("日内データなし")

    # ---------------- daily history ----------------
    st.subheader("日次推移（各日最終スナップショット）")
    hist = _daily(tuple(days[-n_hist:]))
    if len(hist):
        c1, c2 = st.columns(2)
        with c1:
            fig3 = go.Figure()
            for mi, cm in enumerate([m for m in months if m in set(hist.contract_month)]):
                s = hist[hist.contract_month == cm]
                fig3.add_trace(go.Scatter(
                    x=s.day, y=s.atm_iv_pct, mode="lines+markers",
                    name=_fmt_cm(cm),
                    line=dict(color=_MONTH_COLORS[mi % len(_MONTH_COLORS)], width=1.8),
                ))
            fig3.update_layout(height=300, template="plotly_white",
                               title="ATM IV", yaxis_title="%",
                               legend=dict(orientation="h", y=-0.3),
                               margin=dict(l=0, r=0, t=30, b=0))
            st.plotly_chart(fig3, use_container_width=True)
        with c2:
            fig4 = go.Figure()
            for mi, cm in enumerate([m for m in months if m in set(hist.contract_month)]):
                s = hist[hist.contract_month == cm]
                fig4.add_trace(go.Scatter(
                    x=s.day, y=s.skew25_pct, mode="lines+markers",
                    name=_fmt_cm(cm),
                    line=dict(color=_MONTH_COLORS[mi % len(_MONTH_COLORS)], width=1.8),
                ))
            fig4.update_layout(height=300, template="plotly_white",
                               title="25Δスキュー (PUT−CALL)", yaxis_title="%",
                               legend=dict(orientation="h", y=-0.3),
                               margin=dict(l=0, r=0, t=30, b=0))
            st.plotly_chart(fig4, use_container_width=True)
    else:
        st.info("日次データなし")

    # ---------------- raw table ----------------
    with st.expander("スナップショット生データ"):
        show = df[df.contract_month.isin(months)].copy()
        for c in ("iv", "ask_iv", "bid_iv"):
            show[c] = (show[c] * 100).round(2)
        st.dataframe(show, use_container_width=True, height=420)


main()
