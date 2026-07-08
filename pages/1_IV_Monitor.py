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


@st.cache_data(ttl=1800, show_spinner=False)
def _strike_daily(days: tuple, cm: str, ot: str, strikes: tuple) -> pd.DataFrame:
    return iv_views.strike_daily_series(list(days), cm, ot, list(strikes))


@st.cache_data(ttl=_TTL, show_spinner=False)
def _strike_intra(day_str: str, cm: str, ot: str, strikes: tuple) -> pd.DataFrame:
    return iv_views.strike_intraday_series(day_str, cm, ot, list(strikes))


@st.cache_data(ttl=_TTL, show_spinner=False)
def _flow(days: tuple, cm: str) -> pd.DataFrame:
    return iv_views.flow_judgement(list(days), cm)


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

    # ---------------- per-strike IV x OI ----------------
    st.subheader("行使価格別 IV×建玉（買われたか・売られたか）")
    st.caption("判定の定石: 建玉増×IV上昇=新規買い / 建玉増×IV低下=新規売り / "
               "建玉減×IV上昇=買い戻し / 建玉減×IV低下=手仕舞い。"
               "建玉はJPX T+1公表（QRIは前日残）のため、Δ建玉は概ね前営業日の売買を反映。")

    sc1, sc2, sc3 = st.columns([1, 1, 3])
    with sc1:
        cm_sel = st.selectbox("限月", all_months, format_func=_fmt_cm, key="ps_cm")
    with sc2:
        ot_sel = st.radio("タイプ", ["PUT", "CALL"], horizontal=True, key="ps_ot")

    g_now = df[(df.contract_month == cm_sel) & (df.option_type == ot_sel)].copy()
    g_now = g_now[g_now.oi.notna()].sort_values("oi", ascending=False)
    strike_opts = sorted(g_now.strike.astype(int).unique().tolist())
    default_strikes = sorted(g_now.head(5).strike.astype(int).tolist())
    with sc3:
        strikes_sel = st.multiselect("行使価格（既定=建玉上位5本）", strike_opts,
                                     default=default_strikes, key="ps_strikes",
                                     format_func=lambda x: f"{x:,}")

    if strikes_sel:
        hist_days = tuple(days[-n_hist:])
        sd = _strike_daily(hist_days, cm_sel, ot_sel, tuple(sorted(strikes_sel)))
        c1, c2 = st.columns(2)
        with c1:
            fig5 = go.Figure()
            for mi, k in enumerate(sorted(strikes_sel)):
                s = sd[sd.strike == k]
                col = _MONTH_COLORS[mi % len(_MONTH_COLORS)]
                fig5.add_trace(go.Scatter(
                    x=s.day, y=s.iv_pct, mode="lines+markers",
                    name=f"{k:,}", line=dict(color=col, width=1.8)))
            fig5.update_layout(height=320, template="plotly_white",
                               title=f"{_fmt_cm(cm_sel)} {ot_sel} IV（日次）",
                               yaxis_title="IV (%)",
                               legend=dict(orientation="h", y=-0.25),
                               margin=dict(l=0, r=0, t=30, b=0))
            st.plotly_chart(fig5, use_container_width=True)
        with c2:
            fig6 = go.Figure()
            for mi, k in enumerate(sorted(strikes_sel)):
                s = sd[sd.strike == k]
                col = _MONTH_COLORS[mi % len(_MONTH_COLORS)]
                fig6.add_trace(go.Scatter(
                    x=s.day, y=s.oi, mode="lines+markers",
                    name=f"{k:,}", line=dict(color=col, width=1.8, dash="dot")))
            fig6.update_layout(height=320, template="plotly_white",
                               title=f"{_fmt_cm(cm_sel)} {ot_sel} 建玉残（日次）",
                               yaxis_title="枚",
                               legend=dict(orientation="h", y=-0.25),
                               margin=dict(l=0, r=0, t=30, b=0))
            st.plotly_chart(fig6, use_container_width=True)

        intr_s = _strike_intra(day, cm_sel, ot_sel, tuple(sorted(strikes_sel)))
        if len(intr_s):
            fig7 = go.Figure()
            for mi, k in enumerate(sorted(strikes_sel)):
                s = intr_s[intr_s.strike == k]
                fig7.add_trace(go.Scatter(
                    x=s.time, y=s.iv_pct, mode="lines+markers",
                    name=f"{k:,}",
                    line=dict(color=_MONTH_COLORS[mi % len(_MONTH_COLORS)], width=1.6)))
            fig7.update_layout(height=280, template="plotly_white",
                               title=f"{_fmt_cm(cm_sel)} {ot_sel} IV 日内（{day[:4]}/{day[4:6]}/{day[6:8]}）",
                               yaxis_title="IV (%)",
                               legend=dict(orientation="h", y=-0.3),
                               margin=dict(l=0, r=0, t=30, b=0))
            st.plotly_chart(fig7, use_container_width=True)

    # ---- ΔOI x ΔIV 判定（直近2データ日） ----
    fl = _flow(tuple(days), cm_sel)
    if len(fl):
        fl2 = fl[fl.d_oi.notna() & fl.d_iv_pct.notna() & (fl.d_oi != 0)].copy()
        jc1, jc2 = st.columns([3, 2])
        with jc1:
            colmap = {"新規買い": "#1d9e75", "新規売り": "#c83c3c",
                      "買い戻し": "#378add", "手仕舞い": "#888780"}
            fig8 = go.Figure()
            for jname, jcol in colmap.items():
                s = fl2[fl2.judge == jname]
                if not len(s):
                    continue
                fig8.add_trace(go.Scatter(
                    x=s.d_oi, y=s.d_iv_pct, mode="markers+text", name=jname,
                    text=[f"{'P' if t == 'PUT' else 'C'}{k:,}" for t, k in zip(s.option_type, s.strike)],
                    textposition="top center", textfont=dict(size=9),
                    marker=dict(color=jcol, size=10, opacity=0.8)))
            fig8.add_vline(x=0, line=dict(color="gray", width=0.5))
            fig8.add_hline(y=0, line=dict(color="gray", width=0.5))
            fig8.update_layout(height=380, template="plotly_white",
                               title="Δ建玉 × ΔIV（直近2データ日・全ストライク）",
                               xaxis_title="Δ建玉 (枚)", yaxis_title="ΔIV (%pt)",
                               legend=dict(orientation="h", y=-0.2),
                               margin=dict(l=0, r=0, t=30, b=0))
            st.plotly_chart(fig8, use_container_width=True)
        with jc2:
            show_fl = fl2.reindex(fl2.d_oi.abs().sort_values(ascending=False).index).head(15)
            show_fl = show_fl[["option_type", "strike", "oi", "d_oi", "iv_pct", "d_iv_pct", "judge"]]
            show_fl.columns = ["タイプ", "行使", "建玉", "Δ建玉", "IV%", "ΔIV%pt", "判定"]
            show_fl["IV%"] = show_fl["IV%"].round(1)
            show_fl["ΔIV%pt"] = show_fl["ΔIV%pt"].round(2)
            st.dataframe(show_fl, hide_index=True, use_container_width=True, height=380)
    else:
        st.info("Δ建玉×ΔIV: 比較可能な2日分のデータがありません")

    # ---------------- raw table ----------------
    with st.expander("スナップショット生データ"):
        show = df[df.contract_month.isin(months)].copy()
        for c in ("iv", "ask_iv", "bid_iv"):
            show[c] = (show[c] * 100).round(2)
        st.dataframe(show, use_container_width=True, height=420)


main()
