"""IV モニターの「行使価格別 IV×建玉」セクション（判定表・日次チャート・日内チャート）。

日経225オプション（pages/1_IV_Monitor.py）と日経225ミニオプション＝週次満期（pages/2_Weekly_Options.py）で共用する。
データ側の切替は data.iv_views の mini フラグ（スナップショットの保存先、建玉残高表の別紙1/別紙2、両気配しきい値）。
"""
from __future__ import annotations

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from data import iv_views

_TTL = 300  # JPX公表・QRI収集とも時刻が揺れるため5分で再取得
_COLORS = ["#14285a", "#ba7517", "#1d9e75", "#993556", "#534ab7", "#888780"]


@st.cache_data(ttl=600, show_spinner=False)
def _strike_daily(days: tuple, cm: str, ot: str, strikes: tuple, mini: bool) -> pd.DataFrame:
    return iv_views.strike_daily_series(list(days), cm, ot, list(strikes), mini=mini)


@st.cache_data(ttl=_TTL, show_spinner=False)
def _strike_ohlc(days: tuple, n_days: int, cm: str, ot: str, strikes: tuple, freq: str, mini: bool) -> pd.DataFrame:
    return iv_views.strike_intraday_ohlc(list(days), n_days, cm, ot, list(strikes), freq, mini=mini)


@st.cache_data(ttl=600, show_spinner=False)
def _strike_oi_jpx(days: tuple, cm: str, ot: str, strikes: tuple, mini: bool) -> pd.DataFrame:
    return iv_views.strike_daily_oi_jpx(list(days), cm, ot, list(strikes), mini=mini)


@st.cache_data(ttl=_TTL, show_spinner=False)
def _flow(days: tuple, cm: str, mini: bool):
    return iv_views.flow_judgement(list(days), cm, mini=mini)


def _fmt_d(d: str) -> str:
    return f"{int(d[4:6])}/{int(d[6:8])}"


def render_strike_section(days_sel: list[str], cm_sel: str, df: pd.DataFrame, n_hist: int, intra_n: int, *,
                          mini: bool = False, fmt_cm=None, key_prefix: str = "ps") -> None:
    """選択限月（満期）の Δ建玉×超過ΔIV 判定表と、選択行使の日次・日内チャートを描く。

    days_sel : スナップショットのある日（YYYYMMDD、昇順、末尾 = 表示日）
    cm_sel   : 限月（月次 'YYMM'）または満期コード（ミニ 'YYMMDD'）
    df       : 行使一覧・ATM・建玉フォールバックに使うスナップショット（表示日の最終など）
    """
    fmt_cm = fmt_cm or (lambda c: c)
    term = "満期" if mini else "限月"
    oi_src_lbl = "JPX建玉残高表 別紙2" if mini else "JPX建玉残高表"
    day = days_sel[-1] if days_sel else ""

    sc2, sc3 = st.columns([1, 3])
    with sc2:
        ot_sel = st.radio("タイプ", ["PUT", "CALL"], horizontal=True, key=f"{key_prefix}_ot")

    fl, fmeta = _flow(tuple(days_sel), cm_sel, mini)

    # 行使リストはスナップショットの全掲載行使から作る。
    # （夕場スナップショットはOI列が全欠損になるため、OIの有無で絞ると
    # 夜間閲覧時にセクションごと消える）
    g_now = df[(df.contract_month == cm_sel) & (df.option_type == ot_sel)].copy()
    g_now["oi"] = pd.to_numeric(g_now.oi, errors="coerce") if "oi" in g_now.columns else None
    strike_opts = sorted(g_now.strike.astype(int).unique().tolist())

    # 既定の行使 = Δ建玉上位5 → 建玉上位5（JPX） → 中位5本 の順でフォールバック
    default_strikes: list[int] = []
    if len(fl):
        f_ot = fl[fl.option_type == ot_sel]
        m1 = f_ot[f_ot.d_oi.notna() & (f_ot.d_oi != 0)]
        m1 = m1.reindex(m1.d_oi.abs().sort_values(ascending=False).index)
        default_strikes = [int(k) for k in m1.head(5).strike if int(k) in strike_opts]
        if not default_strikes:
            m2 = f_ot[f_ot.oi.notna()].sort_values("oi", ascending=False)
            default_strikes = [int(k) for k in m2.head(5).strike if int(k) in strike_opts]
    if not default_strikes:
        g_oi = g_now[g_now.oi.notna()].sort_values("oi", ascending=False)
        default_strikes = sorted(g_oi.head(5).strike.astype(int).tolist())
    if not default_strikes and strike_opts:
        mid = len(strike_opts) // 2
        default_strikes = strike_opts[max(0, mid - 2):mid + 3]

    with sc3:
        sel_mode = st.radio("行使の選び方", ["Δ建玉上位5", "範囲指定", "個別指定"],
                            horizontal=True, key=f"{key_prefix}_mode")
    if sel_mode == "範囲指定" and strike_opts:
        lo0 = min(default_strikes) if default_strikes else strike_opts[len(strike_opts) // 3]
        hi0 = max(default_strikes) if default_strikes else strike_opts[2 * len(strike_opts) // 3]
        lo, hi = st.select_slider(
            "行使価格の範囲", options=strike_opts, value=(lo0, hi0),
            key=f"{key_prefix}_range_{cm_sel}_{ot_sel}", format_func=lambda x: f"{x:,}")
        strikes_sel = [k for k in strike_opts if lo <= k <= hi]
    elif sel_mode == "個別指定":
        strikes_sel = st.multiselect(
            "行使価格", strike_opts, default=sorted(default_strikes),
            key=f"{key_prefix}_strikes_{day}_{cm_sel}_{ot_sel}",
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
                   f"（Δ建玉 = {oi_src_lbl}〔当日20:00頃公表〕の増減 / "
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
        lvl_txt = f"地合い（{term}内ΔIV中央値）: {fmeta['level_pct']:+.2f}%pt（n={fmeta['n_level']}）"
        st.caption(f"{atm_txt}  |  {lvl_txt}  |  {win_txt}")
    else:
        st.caption(atm_txt)
    st.caption("判定 = Δ建玉 × 超過ΔIV（ΔIVから地合いを控除）: "
               "建玉増×IV上昇=新規買い / 建玉増×IV低下=新規売り / "
               "建玉減×IV上昇=買い戻し / 建玉減×IV低下=手仕舞い / "
               f"|超過ΔIV|<{iv_views._NEUTRAL_BAND}%ptは中立")
    st.caption("出来高加重超過ΔIV = Σ e(t)·Δ出来高(t) ÷ Σ Δ出来高(t)。e(t) = 直前スナップショットからの"
               f"ΔIV − 同時点の地合い（{term}内ΔIV中央値）。窓は超過ΔIVと同じで、売買が入った区間のIVの動きのみを集計"
               "（判定には未使用）")

    # ---- Δ建玉 × 超過ΔIV 判定 ----
    if len(fl):
        # テーブルはΔ建玉が動いた全系列が母集団（IV欠損=判定不能も含める）
        tbl = fl[fl.d_oi.notna() & (fl.d_oi != 0)].copy()
        if "vw_ex_pct" not in tbl.columns:  # QRIフォールバック時は未算出
            tbl["vw_ex_pct"] = None
        tbl["judge"] = tbl.judge.where(tbl.judge != "", "判定不能(IV欠損)")
        show_fl = tbl.reindex(tbl.d_oi.abs().sort_values(ascending=False).index).head(15)
        show_fl = show_fl[["option_type", "strike", "oi", "d_oi", "volume",
                           "iv_pct", "d_iv_pct", "d_iv_ex_pct", "vw_ex_pct", "judge"]]
        show_fl.columns = ["タイプ", "行使", "建玉", "Δ建玉", "出来高",
                           "IV%", "ΔIV%pt", "超過ΔIV", "出来高加重超過ΔIV", "判定"]
        show_fl["出来高加重超過ΔIV"] = pd.to_numeric(show_fl["出来高加重超過ΔIV"], errors="coerce").round(2)
        show_fl["IV%"] = show_fl["IV%"].round(1)
        show_fl["ΔIV%pt"] = show_fl["ΔIV%pt"].round(2)
        show_fl["超過ΔIV"] = show_fl["超過ΔIV"].round(2)
        st.dataframe(show_fl, hide_index=True, use_container_width=True, height=420)
    else:
        st.info("Δ建玉×ΔIV: 比較可能な2日分のデータがありません")

    if not strikes_sel:
        return

    # ---- 選択行使の時系列 ----
    hist_days = tuple(days_sel[-n_hist:])
    ck = tuple(chart_strikes)
    sd = _strike_daily(hist_days, cm_sel, ot_sel, ck, mini)
    c1, c2 = st.columns(2)
    with c1:
        fig5 = go.Figure()
        for mi, k in enumerate(chart_strikes):
            s = sd[sd.strike == k] if len(sd) else sd
            fig5.add_trace(go.Scatter(
                x=s.day if len(s) else [], y=s.iv_pct if len(s) else [], mode="lines+markers",
                name=f"{k:,}",
                line=dict(color=_COLORS[mi % len(_COLORS)], width=1.8)))
        fig5.update_layout(height=320, template="plotly_white",
                           title=f"{fmt_cm(cm_sel)} {ot_sel} IV（日次）",
                           yaxis_title="IV (%)",
                           legend=dict(orientation="h", y=-0.25),
                           margin=dict(l=0, r=0, t=30, b=0))
        st.plotly_chart(fig5, use_container_width=True)
    od = _strike_oi_jpx(hist_days, cm_sel, ot_sel, ck, mini)
    oi_src, oi_df = (oi_src_lbl, od) if len(od) else ("QRI表示値", sd)
    with c2:
        fig6 = go.Figure()
        for mi, k in enumerate(chart_strikes):
            s = oi_df[oi_df.strike == k] if len(oi_df) else oi_df
            fig6.add_trace(go.Scatter(
                x=s.day if len(s) else [], y=s.oi if len(s) else [], mode="lines+markers",
                name=f"{k:,}",
                line=dict(color=_COLORS[mi % len(_COLORS)], width=1.8, dash="dot")))
        fig6.update_layout(height=320, template="plotly_white",
                           title=f"{fmt_cm(cm_sel)} {ot_sel} 建玉残（日次・{oi_src}）",
                           yaxis_title="枚",
                           legend=dict(orientation="h", y=-0.25),
                           margin=dict(l=0, r=0, t=30, b=0))
        st.plotly_chart(fig6, use_container_width=True)
        if oi_src == "QRI表示値":
            st.caption(f"※この{term}は{oi_src_lbl}非掲載のためQRI表示値。"
                       "2026-07-12以前の収集分は4桁以上の建玉が下位桁欠落（既知欠陥）")

    c1, c2 = st.columns([3, 2])
    mode = c1.radio("日内チャート", ["IV", "累積超過ΔIV"], horizontal=True,
                    key=f"{key_prefix}_intra_mode")
    freq_lab = c2.radio("足", ["30分", "1時間", "2時間"], index=1, horizontal=True,
                        key=f"{key_prefix}_intra_freq")
    freq = {"30分": "30min", "1時間": "1h", "2時間": "2h"}[freq_lab]
    ohlc = _strike_ohlc(tuple(days_sel), intra_n, cm_sel, ot_sel, ck, freq, mini)
    if not len(ohlc):
        st.caption("日内チャート: この期間のスナップショットがありません")
        return
    pre = "iv" if mode == "IV" else "e"
    for ci in range(0, len(chart_strikes), 3):  # 1グラフ3行使まで
        grp = chart_strikes[ci:ci + 3]
        fig7 = go.Figure()
        for mi, k in enumerate(grp, start=ci):
            s_ = ohlc[ohlc.strike == k]
            s_ = s_.dropna(subset=["iv_c"]) if pre == "iv" else s_
            if s_.empty:
                continue
            col = _COLORS[mi % len(_COLORS)]
            if pre == "iv":
                # 陽線=塗り、陰線=白抜き。色は行使価格ごとに固定
                fig7.add_trace(go.Candlestick(
                    x=s_.bar, open=s_.iv_o, high=s_.iv_h, low=s_.iv_l, close=s_.iv_c,
                    name=f"{k:,}",
                    increasing=dict(line=dict(color=col, width=1), fillcolor=col),
                    decreasing=dict(line=dict(color=col, width=1), fillcolor="white")))
            else:
                # 取引日ごとに累積を0から始めるため、日の境目で線を切る（つなぐと全行使に同じ段差が出る）
                brk = s_.tdate.ne(s_.tdate.shift()) & (s_.index != s_.index[0])
                s_ = pd.concat([s_, s_[brk].assign(e_cum=None, e_sum=None,
                                                   bar=s_.bar[brk] - pd.Timedelta(seconds=1))]
                               ).sort_values("bar")
                fig7.add_trace(go.Scatter(
                    x=s_.bar, y=s_.e_cum, mode="lines", name=f"{k:,}", connectgaps=False,
                    line=dict(color=col, width=1.6, shape="hv"),
                    customdata=s_.e_sum,
                    hovertemplate="%{x|%m/%d %H:%M} 累積 %{y:.2f} / 足内 %{customdata:.2f}"))
            fig7.add_trace(go.Bar(
                x=s_.bar, y=s_.dv, name=f"{k:,} 出来高", marker_color=col,
                opacity=0.35, yaxis="y2", showlegend=False))
        if pre == "e":
            fig7.add_hline(y=0, line=dict(color="#888780", width=1, dash="dot"))
            for b0 in ohlc.groupby("tdate").bar.min().sort_values().iloc[1:]:
                fig7.add_vline(x=b0, line=dict(color="#b4b2a9", width=1, dash="dot"))  # 取引日の境目（累積を0に戻す）
        breaks = [dict(bounds=[6.25, 8.75], pattern="hour"),
                  dict(bounds=[15.75, 16.5], pattern="hour")]
        if all(pd.Timestamp(f"{d[:4]}-{d[4:6]}-{d[6:8]}").weekday() < 5
               for d in ohlc.day.unique()):
            breaks.append(dict(bounds=["sat", "mon"]))
        ytitle = "IV (%)" if pre == "iv" else "累積超過ΔIV (%pt)"
        fig7.update_layout(height=340, template="plotly_white",
                           title=f"{fmt_cm(cm_sel)} {ot_sel} {mode} {freq_lab}足"
                                 f"（直近{intra_n}取引日）{' / '.join(f'{k:,}' for k in grp)}",
                           yaxis=dict(title=ytitle),
                           yaxis2=dict(title="出来高 (枚)", overlaying="y",
                                       side="right", showgrid=False,
                                       rangemode="tozero"),
                           barmode="overlay",
                           xaxis=dict(type="date", tickformat="%m/%d %H:%M",
                                      dtick=3 * 60 * 60 * 1000, tickangle=-45,
                                      rangebreaks=breaks,
                                      rangeslider=dict(visible=False)),
                           legend=dict(orientation="h", y=-0.45),
                           margin=dict(l=0, r=0, t=30, b=0))
        st.plotly_chart(fig7, use_container_width=True)
    st.caption("棒=足内出来高（右軸）。IV: 足内スナップショット（約15分間隔）の始値・高値・安値・終値。"
               "累積超過ΔIV: e(t) = ΔIV_k(t, t−1) − L(t) を窓の起点から累積。"
               f"L(t) = 同{term}の全行使（PUT/CALL、両気配）の固定行使ΔIV(t, t−1)の中央値で、表の超過ΔIVと同じ定義。"
               "累積は取引日（夜間16:30〜日中15:15）ごとに前取引日の最終値を0として開始。"
               f"両気配{iv_views._min_two_sided(mini)}銘柄未満のスナップショット（引け後・夜間寄り前後）は除外。")
