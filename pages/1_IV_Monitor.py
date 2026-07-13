"""IV Monitor — QRI NK225 option IV snapshots (15-min delayed) viewer.

Data: R2 qri_iv/raw/YYYYMMDD/*.parquet (scripts/fetch_qri_iv.py が収集)
主目的: 行使価格別の Δ建玉×ΔIV からフロー（買われたか・売られたか）を判定する。
スマイル・ATM時系列は補助情報としてexpanderに格納。
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


@st.cache_data(ttl=1800, show_spinner=False)
def _daily(days: tuple) -> pd.DataFrame:
    return iv_views.daily_atm_series(list(days))


@st.cache_data(ttl=1800, show_spinner=False)
def _strike_daily(days: tuple, cm: str, ot: str, strikes: tuple) -> pd.DataFrame:
    return iv_views.strike_daily_series(list(days), cm, ot, list(strikes))


@st.cache_data(ttl=_TTL, show_spinner=False)
def _strike_intra(days: tuple, n_days: int, cm: str, ot: str, strikes: tuple) -> pd.DataFrame:
    return iv_views.strike_intraday_series(list(days), n_days, cm, ot, list(strikes))


@st.cache_data(ttl=1800, show_spinner=False)
def _strike_oi_jpx(days: tuple, cm: str, ot: str, strikes: tuple) -> pd.DataFrame:
    return iv_views.strike_daily_oi_jpx(list(days), cm, ot, list(strikes))


@st.cache_data(ttl=_TTL, show_spinner=False)
def _flow(days: tuple, cm: str):
    return iv_views.flow_judgement(list(days), cm)


_COLORS = ["#14285a", "#ba7517", "#1d9e75", "#993556", "#534ab7", "#888780"]
_JUDGE_COLORS = {"新規買い": "#1d9e75", "新規売り": "#c83c3c",
                 "買い戻し": "#378add", "手仕舞い": "#888780", "中立": "#c9c5bc"}


def _fmt_cm(cm: str) -> str:
    return f"20{cm[:2]}年{cm[2:]}月限" if len(cm) == 4 else cm


def _fmt_d(d: str) -> str:
    return f"{int(d[4:6])}/{int(d[6:8])}"


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
    key = keys[-1]  # 最終スナップショット固定（日内推移は下のチャートで確認）
    df = _snapshot(key)
    if df.empty:
        st.warning("スナップショットを読み込めませんでした")
        st.stop()
    if len(days) > 5:
        n_hist = st.sidebar.slider("日次チャートの日数", 5, len(days), min(20, len(days)))
    else:
        n_hist = len(days)
    intra_n = st.sidebar.slider("日内チャートの日数", 1, 5, 3)

    upd = df.source_update_time.dropna()
    st.caption(f"最終スナップショット: {iv_views.key_time_label(key)}  |  "
               f"ソース更新: {upd.iloc[0] if len(upd) else '-'}  |  行数: {len(df):,}")

    all_months = sorted(df.contract_month.unique())
    days_sel = [d for d in days if d <= day]  # 選択日をアンカーに過去方向のみ使用

    # ---------------- フロー判定（主機能） ----------------
    st.subheader("行使価格別 IV×建玉（買われたか・売られたか）")

    sc1, sc2, sc3 = st.columns([1, 1, 3])
    with sc1:
        cm_sel = st.selectbox("限月", all_months, format_func=_fmt_cm, key="ps_cm")
    with sc2:
        ot_sel = st.radio("タイプ", ["PUT", "CALL"], horizontal=True, key="ps_ot")

    fl, fmeta = _flow(tuple(days_sel), cm_sel)

    g_now = df[(df.contract_month == cm_sel) & (df.option_type == ot_sel)].copy()
    g_now = g_now[g_now.oi.notna()]
    strike_opts = sorted(g_now.strike.astype(int).unique().tolist())

    # 既定の行使 = Δ建玉の絶対値上位5本（フロー未算出時は建玉上位5本）
    default_strikes: list[int] = []
    if len(fl):
        f_ot = fl[(fl.option_type == ot_sel) & fl.d_oi.notna() & (fl.d_oi != 0)]
        f_ot = f_ot.reindex(f_ot.d_oi.abs().sort_values(ascending=False).index)
        default_strikes = [int(k) for k in f_ot.head(5).strike if int(k) in strike_opts]
    if not default_strikes:
        default_strikes = sorted(
            g_now.sort_values("oi", ascending=False).head(5).strike.astype(int).tolist())

    with sc3:
        sel_mode = st.radio("行使の選び方", ["Δ建玉上位5", "範囲指定", "個別指定"],
                            horizontal=True, key="ps_mode")
    if sel_mode == "範囲指定" and strike_opts:
        lo0 = min(default_strikes) if default_strikes else strike_opts[len(strike_opts) // 3]
        hi0 = max(default_strikes) if default_strikes else strike_opts[2 * len(strike_opts) // 3]
        lo, hi = st.select_slider(
            "行使価格の範囲", options=strike_opts, value=(lo0, hi0),
            key=f"ps_range_{cm_sel}_{ot_sel}", format_func=lambda x: f"{x:,}")
        strikes_sel = [k for k in strike_opts if lo <= k <= hi]
    elif sel_mode == "個別指定":
        strikes_sel = st.multiselect(
            "行使価格", strike_opts, default=sorted(default_strikes),
            key=f"ps_strikes_{day}_{cm_sel}_{ot_sel}",
            format_func=lambda x: f"{x:,}")
    else:
        strikes_sel = sorted(default_strikes)
        if strikes_sel:
            st.caption("対象行使: " + " / ".join(f"{k:,}" for k in strikes_sel))

    # チャートは線が多すぎると判読不能になるため上限16本（建玉上位を優先）
    chart_strikes = sorted(strikes_sel)
    if len(chart_strikes) > 16:
        top_oi = g_now[g_now.strike.isin(chart_strikes)].sort_values("oi", ascending=False)
        chart_strikes = sorted(top_oi.head(16).strike.astype(int).tolist())
        st.caption(f"※チャートは選択{len(strikes_sel)}本のうち建玉上位16本のみ描画"
                   "（生データ表は全選択行使を表示）")

    # 文脈情報: ATM・地合い・観測窓
    summ = iv_views.month_summary(df)
    srow = summ[summ.contract_month == cm_sel]
    atm_txt = "-"
    if len(srow) and pd.notna(srow.iloc[0].atm_iv):
        atm_txt = f"ATM {srow.iloc[0].atm_strike:,.0f} / {srow.iloc[0].atm_iv * 100:.1f}%"
    if fmeta.get("source") == "jpx":
        iv0, iv1 = fmeta["iv_days"]
        win_txt = (f"対象取引日: {_fmt_d(fmeta['trade_day'])}"
                   f"（Δ建玉 = JPX建玉残高表〔当日20:00頃公表〕の増減 / "
                   f"ΔIV = {_fmt_d(iv1)}→{_fmt_d(iv0)} 最終気配の変化）")
    elif fmeta.get("source") == "qri":
        od0, od1 = fmeta["oi_days"]
        iv0, iv1 = fmeta["iv_days"]
        win_txt = (f"Δ建玉 = {_fmt_d(od1)}→{_fmt_d(od0)} のOI差"
                   f"（QRI表示値・翌取引日朝反映。2026-07-12以前収集分は"
                   f"4桁以上の建玉に桁欠落の既知欠陥あり・参考値） / "
                   f"ΔIV = {_fmt_d(iv1)}→{_fmt_d(iv0)} の気配変化"
                   + ("（OIの取引日窓に整合）" if fmeta["aligned"]
                      else "（整合用データ不足のため同時点差で近似）"))
    else:
        win_txt = ""
    if win_txt:
        lvl_txt = f"地合い（限月内ΔIV中央値）: {fmeta['level_pct']:+.2f}%pt（n={fmeta['n_level']}）"
        st.caption(f"{atm_txt}  |  {lvl_txt}  |  {win_txt}")
    else:
        st.caption(atm_txt)
    st.caption("判定 = Δ建玉 × 超過ΔIV（ΔIVから地合いを控除）: "
               "建玉増×IV上昇=新規買い / 建玉増×IV低下=新規売り / "
               "建玉減×IV上昇=買い戻し / 建玉減×IV低下=手仕舞い / "
               f"|超過ΔIV|<{iv_views._NEUTRAL_BAND}%ptは中立")

    # ---- Δ建玉 × 超過ΔIV 判定 ----
    if len(fl):
        fl2 = fl[fl.judge != ""].copy()
        fig8 = go.Figure()
        for jname, jcol in _JUDGE_COLORS.items():
            s = fl2[fl2.judge == jname]
            if not len(s):
                continue
            fig8.add_trace(go.Scatter(
                x=s.d_oi, y=s.d_iv_ex_pct, mode="markers+text", name=jname,
                text=[f"{'P' if t == 'PUT' else 'C'}{k:,}"
                      for t, k in zip(s.option_type, s.strike)],
                textposition="top center", textfont=dict(size=9),
                marker=dict(color=jcol, size=10, opacity=0.8)))
        fig8.add_vline(x=0, line=dict(color="gray", width=0.5))
        fig8.add_hline(y=0, line=dict(color="gray", width=0.5))
        fig8.update_layout(height=420, template="plotly_white",
                           title="Δ建玉 × 超過ΔIV（全ストライク）",
                           xaxis_title="Δ建玉 (枚)",
                           yaxis_title="超過ΔIV (%pt, 地合い控除後)",
                           legend=dict(orientation="h", y=-0.2),
                           margin=dict(l=0, r=0, t=30, b=0))
        st.plotly_chart(fig8, use_container_width=True)

        # テーブルはΔ建玉が動いた全系列が母集団（IV欠損=判定不能も含める）
        tbl = fl[fl.d_oi.notna() & (fl.d_oi != 0)].copy()
        tbl["judge"] = tbl.judge.where(tbl.judge != "", "判定不能(IV欠損)")
        show_fl = tbl.reindex(tbl.d_oi.abs().sort_values(ascending=False).index).head(15)
        show_fl = show_fl[["option_type", "strike", "oi", "d_oi", "volume",
                           "iv_pct", "d_iv_pct", "d_iv_ex_pct", "judge"]]
        show_fl.columns = ["タイプ", "行使", "建玉", "Δ建玉", "出来高",
                           "IV%", "ΔIV%pt", "超過ΔIV", "判定"]
        show_fl["IV%"] = show_fl["IV%"].round(1)
        show_fl["ΔIV%pt"] = show_fl["ΔIV%pt"].round(2)
        show_fl["超過ΔIV"] = show_fl["超過ΔIV"].round(2)
        st.dataframe(show_fl, hide_index=True, use_container_width=True, height=420)
    else:
        st.info("Δ建玉×ΔIV: 比較可能な2日分のデータがありません")

    # ---- 選択行使の時系列 ----
    if strikes_sel:
        hist_days = tuple(days_sel[-n_hist:])
        ck = tuple(chart_strikes)
        sd = _strike_daily(hist_days, cm_sel, ot_sel, ck)
        c1, c2 = st.columns(2)
        with c1:
            fig5 = go.Figure()
            for mi, k in enumerate(chart_strikes):
                s = sd[sd.strike == k]
                fig5.add_trace(go.Scatter(
                    x=s.day, y=s.iv_pct, mode="lines+markers",
                    name=f"{k:,}",
                    line=dict(color=_COLORS[mi % len(_COLORS)], width=1.8)))
            fig5.update_layout(height=320, template="plotly_white",
                               title=f"{_fmt_cm(cm_sel)} {ot_sel} IV（日次）",
                               yaxis_title="IV (%)",
                               legend=dict(orientation="h", y=-0.25),
                               margin=dict(l=0, r=0, t=30, b=0))
            st.plotly_chart(fig5, use_container_width=True)
        od = _strike_oi_jpx(hist_days, cm_sel, ot_sel, ck)
        oi_src, oi_df = ("JPX建玉残高表", od) if len(od) else ("QRI表示値", sd)
        with c2:
            fig6 = go.Figure()
            for mi, k in enumerate(chart_strikes):
                s = oi_df[oi_df.strike == k]
                fig6.add_trace(go.Scatter(
                    x=s.day, y=s.oi, mode="lines+markers",
                    name=f"{k:,}",
                    line=dict(color=_COLORS[mi % len(_COLORS)], width=1.8, dash="dot")))
            fig6.update_layout(height=320, template="plotly_white",
                               title=f"{_fmt_cm(cm_sel)} {ot_sel} 建玉残（日次・{oi_src}）",
                               yaxis_title="枚",
                               legend=dict(orientation="h", y=-0.25),
                               margin=dict(l=0, r=0, t=30, b=0))
            st.plotly_chart(fig6, use_container_width=True)
            if oi_src == "QRI表示値":
                st.caption("※この限月はJPX建玉残高表（別紙1）非掲載のためQRI表示値。"
                           "2026-07-12以前の収集分は4桁以上の建玉が下位桁欠落（既知欠陥）")

        intr_s = _strike_intra(tuple(days_sel), intra_n, cm_sel, ot_sel, ck)
        if len(intr_s):
            fig7 = go.Figure()
            for mi, k in enumerate(chart_strikes):
                s = intr_s[intr_s.strike == k]
                col = _COLORS[mi % len(_COLORS)]
                fig7.add_trace(go.Scatter(
                    x=s.ts, y=s.iv_pct, mode="lines+markers",
                    name=f"{k:,}", line=dict(color=col, width=1.6)))
                # 出来高はスナップショット間の増分（取引日切替のリセット時は当該値）
                dv = s.volume.diff()
                dv = dv.where(dv >= 0, s.volume)
                if len(dv):
                    dv.iloc[0] = s.volume.iloc[0]
                fig7.add_trace(go.Bar(
                    x=s.ts, y=dv, name=f"{k:,} 出来高", width=8 * 60 * 1000,
                    marker_color=col, opacity=0.35, yaxis="y2", showlegend=False))
            # 取引のない時間帯を詰める（昼: 6:00-8:45 / 夕: 15:45-17:00 / 週末）
            breaks = [dict(bounds=[6, 8.75], pattern="hour"),
                      dict(bounds=[15.75, 17], pattern="hour")]
            if all(pd.Timestamp(f"{d[:4]}-{d[4:6]}-{d[6:8]}").weekday() < 5
                   for d in intr_s.day.unique()):
                breaks.append(dict(bounds=["sat", "mon"]))
            fig7.update_layout(height=380, template="plotly_white",
                               title=f"{_fmt_cm(cm_sel)} {ot_sel} IV 日内（直近{intra_n}取引日）"
                                     "　棒=出来高（スナップショット間増分・右軸）",
                               yaxis=dict(title="IV (%)"),
                               yaxis2=dict(title="出来高 (枚)", overlaying="y",
                                           side="right", showgrid=False,
                                           rangemode="tozero"),
                               barmode="overlay",
                               xaxis=dict(type="date", tickformat="%m/%d %H:%M",
                                          tickangle=-45, rangebreaks=breaks),
                               legend=dict(orientation="h", y=-0.45),
                               margin=dict(l=0, r=0, t=30, b=0))
            st.plotly_chart(fig7, use_container_width=True)

    # ---------------- 補助情報（expander） ----------------
    with st.expander("IVスマイル（選択限月・最終スナップショット）"):
        sm = iv_views.smile_frame(df, [cm_sel])
        fig = go.Figure()
        for ot, dashv, col in (("CALL", "solid", _COLORS[0]), ("PUT", "dot", _COLORS[1])):
            s = sm[sm.option_type == ot]
            if not len(s):
                continue
            fig.add_trace(go.Scatter(
                x=s.strike, y=s.iv_pct, mode="lines+markers",
                name=f"{_fmt_cm(cm_sel)} {ot}",
                line=dict(color=col, dash=dashv, width=1.6),
                marker=dict(size=4)))
        if len(srow) and pd.notna(srow.iloc[0].atm_strike):
            fig.add_vline(x=float(srow.iloc[0].atm_strike),
                          line=dict(color="gray", width=0.8, dash="dash"))
        fig.update_layout(height=380, template="plotly_white",
                          xaxis_title="行使価格", yaxis_title="IV (%)",
                          legend=dict(orientation="h", y=-0.18),
                          margin=dict(l=0, r=0, t=10, b=0))
        st.plotly_chart(fig, use_container_width=True)
        st.caption("実線=CALL / 点線=PUT / 縦破線=ATM行使価格。"
                   "特定行使の局所的な盛り上がり・凹みはフロー集中の痕跡")

    with st.expander("ATM IV 日次推移（全限月）"):
        hist = _daily(tuple(days_sel[-n_hist:]))
        if len(hist):
            fig3 = go.Figure()
            for mi, cm in enumerate(sorted(hist.contract_month.unique())):
                s = hist[hist.contract_month == cm]
                fig3.add_trace(go.Scatter(
                    x=s.day, y=s.atm_iv_pct, mode="lines+markers",
                    name=_fmt_cm(cm),
                    line=dict(color=_COLORS[mi % len(_COLORS)], width=1.8)))
            fig3.update_layout(height=300, template="plotly_white",
                               yaxis_title="ATM IV (%)",
                               legend=dict(orientation="h", y=-0.3),
                               margin=dict(l=0, r=0, t=10, b=0))
            st.plotly_chart(fig3, use_container_width=True)
        else:
            st.info("日次データなし")

    with st.expander("スナップショット生データ（選択限月）"):
        only_sel = st.checkbox("選択した行使価格のみ", value=True, key="raw_only_sel")
        show = df[df.contract_month == cm_sel].copy()
        if only_sel and strikes_sel:
            show = show[show.strike.astype(int).isin(strikes_sel)]
        show = show.sort_values(["option_type", "strike"])
        for c in ("iv", "ask_iv", "bid_iv"):
            show[c] = (show[c] * 100).round(2)
        st.dataframe(show, use_container_width=True, height=420)


main()
