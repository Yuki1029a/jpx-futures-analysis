"""週次オプション（日経225ミニオプション、水曜・金曜 SQ）の建玉モニター。

出所: JPX 建玉残高表 別紙2（当日 20:00 頃公表）。取引単位は指数 × 100。
"""
from __future__ import annotations

import sys
from datetime import date, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from data import iv_views
from data import weekly_options as wo
from ui.iv_monitor_section import render_strike_section

st.set_page_config(page_title="週次オプション", layout="wide")
_TTL = 300
_COL = {"put": "#eb6834", "call": "#2a78d6", "ink": "#0b0b0b", "mut": "#888780", "mp": "#1baf7a"}


@st.cache_data(ttl=_TTL, show_spinner=False)
def _latest(start: date):
    return wo.latest_available_date(start)


@st.cache_data(ttl=_TTL, show_spinner=False)
def _expiries(d: date):
    return wo.list_expiries(d)


@st.cache_data(ttl=_TTL, show_spinner=False)
def _frame(d: date, exp: str):
    return wo.strike_frame(d, exp)


@st.cache_data(ttl=_TTL, show_spinner=False)
def _hist(exp: str, d: date, n: int):
    return wo.history(exp, d, n)


@st.cache_data(ttl=_TTL, show_spinner=False)
def _iv_keys(d: date):
    from data.qri_iv_loader import _R2_PREFIX_MINI, list_snapshot_keys
    return list_snapshot_keys(d, _R2_PREFIX_MINI)


@st.cache_data(ttl=_TTL, show_spinner=False)
def _iv_snapshot(key: str):
    from data.qri_iv_loader import load_snapshot
    df = load_snapshot(key)
    return df if df is not None else pd.DataFrame()


@st.cache_data(ttl=_TTL, show_spinner=False)
def _mini_days():
    return iv_views.list_available_days(mini=True)


def _pick_close(keys: list[str]):
    """日中引けのスナップショット = 気配時刻（source_update_time）が 08:45〜15:30:59 の最終。15:30 以降はマーケットメイカーが気配を引くため使わない。"""
    for k in reversed(keys):
        df = _iv_snapshot(k)
        if df.empty or "source_update_time" not in df.columns:
            continue
        s_ = df.source_update_time.dropna().astype(str); t = s_.iloc[0][11:19] if len(s_) else ""
        if "08:45:00" <= t <= "15:30:59":
            return k, df
    return None, None


def main():
    st.sidebar.title("週次オプション")
    if st.sidebar.button("最新データに更新", use_container_width=True):
        st.cache_data.clear(); st.rerun()
    pick = st.sidebar.date_input("基準日", value=date.today())
    d = _latest(pick)
    if d is None:
        st.warning("直近 10 日に建玉残高表がありません"); st.stop()
    if d != pick:
        st.sidebar.caption(f"{pick:%m/%d} の建玉残高表が無いため {d:%m/%d} を表示")
    exps = _expiries(d)
    if not exps:
        st.warning(f"{d:%Y/%m/%d} の別紙2にミニオプションの建玉がありません"); st.stop()
    labels = {e: wo.expiry_info(e)["label"] for e in exps}
    live = [e for e in exps if wo.expiry_info(e)["sq"] > d] or exps
    default = max(live, key=lambda e: float(_frame(d, e)[["put_oi", "call_oi"]].sum().sum()))  # 建玉が最大の未満期
    exp = st.sidebar.selectbox("満期", exps, index=exps.index(default), format_func=lambda e: labels[e])
    n_hist = st.sidebar.slider("推移の日数", 5, 30, 15)
    intra_n = st.sidebar.slider("日内チャートの日数", 1, 5, 2)
    info = wo.expiry_info(exp)

    st.title("週次オプション（日経225ミニオプション）建玉モニター")
    st.caption(f"建玉残高表 別紙2（{d:%Y/%m/%d} 引け、20:00 公表）。満期コード {exp} = 最終取引日 {info['last_trading']:%m/%d}、SQ 算出 {info['sq']:%m/%d}。"
               "取引単位は指数 × 100。IV は QRI（15 分遅延）を 2026-10-20 から収集。手口（参加者別）は JPX 未公表。")
    f = _frame(d, exp)
    if f.empty:
        st.info("この満期の建玉がありません"); st.stop()
    mp, pain = wo.max_pain(f)
    put_oi, call_oi = int(f.put_oi.sum()), int(f.call_oi.sum())
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("プット建玉", f"{put_oi:,}", f"{int(f.put_chg.sum()):+,}")
    c2.metric("コール建玉", f"{call_oi:,}", f"{int(f.call_chg.sum()):+,}")
    c3.metric("P/C 比（建玉）", f"{put_oi / call_oi:.2f}" if call_oi else "-")
    c4.metric("マックスペイン", f"{mp:,} 円" if mp else "-")
    c5.metric("取引高（P/C）", f"{int(f.put_vol.sum()):,} / {int(f.call_vol.sum()):,}")

    # ---- 行使価格別チャート
    lo = f[(f.put_oi > 0) | (f.call_oi > 0)]
    fig = go.Figure()
    fig.add_trace(go.Bar(x=lo.strike, y=lo.call_oi, name="コール建玉", marker_color=_COL["call"], width=90))
    fig.add_trace(go.Bar(x=lo.strike, y=-lo.put_oi, name="プット建玉（下向き）", marker_color=_COL["put"], width=90))
    if mp:
        fig.add_vline(x=mp, line=dict(color=_COL["mp"], dash="dash"), annotation_text=f"マックスペイン {mp:,}", annotation_position="top")
    fig.update_layout(height=380, template="plotly_white", barmode="overlay", title=f"{labels[exp]}：行使価格別の建玉（{d:%m/%d}）",
                      xaxis=dict(title="行使価格（円）"), yaxis=dict(title="建玉（枚）"), legend=dict(orientation="h", y=-0.25), margin=dict(l=0, r=0, t=40, b=0))
    st.plotly_chart(fig, use_container_width=True)

    # ---- 表
    show = f[(f.put_oi > 0) | (f.call_oi > 0) | (f.put_vol > 0) | (f.call_vol > 0)].copy()
    show = show[["strike", "put_vol", "put_oi", "put_chg", "call_vol", "call_oi", "call_chg"]]
    show.columns = ["行使価格", "PUT 取引高", "PUT 建玉", "PUT Δ建玉", "CALL 取引高", "CALL 建玉", "CALL Δ建玉"]
    st.dataframe(show.style.format("{:,.0f}").format("{:+,.0f}", subset=["PUT Δ建玉", "CALL Δ建玉"]), hide_index=True, use_container_width=True, height=min(600, 40 + 35 * len(show)))
    top = f.assign(tot=f.put_chg.abs() + f.call_chg.abs()).nlargest(8, "tot")
    st.caption("Δ建玉の大きい行使価格: " + "、".join(f"{int(r.strike):,}（P {int(r.put_chg):+,} / C {int(r.call_chg):+,}）" for r in top.itertuples()))

    # ---- 推移
    h = _hist(exp, d, n_hist)
    if len(h) >= 2:
        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(x=h.date, y=h.put_oi, name="プット建玉", line=dict(color=_COL["put"], width=2)))
        fig2.add_trace(go.Scatter(x=h.date, y=h.call_oi, name="コール建玉", line=dict(color=_COL["call"], width=2)))
        fig2.update_layout(height=300, template="plotly_white", title="建玉の推移", yaxis=dict(title="枚"), legend=dict(orientation="h", y=-0.3), margin=dict(l=0, r=0, t=40, b=0))
        fig3 = go.Figure()
        fig3.add_trace(go.Scatter(x=h.date, y=h.max_pain, name="マックスペイン", line=dict(color=_COL["mp"], width=2), mode="lines+markers"))
        fig3.update_layout(height=260, template="plotly_white", title="マックスペインの推移", yaxis=dict(title="円"), margin=dict(l=0, r=0, t=40, b=0))
        cA, cB = st.columns(2)
        cA.plotly_chart(fig2, use_container_width=True); cB.plotly_chart(fig3, use_container_width=True)
        hh = h.copy(); hh["date"] = hh.date.map(lambda x: x.strftime("%m/%d"))
        hh.columns = ["日付", "PUT 建玉", "CALL 建玉", "PUT 取引高", "CALL 取引高", "PUT 新規建玉", "CALL 新規建玉", "マックスペイン"]
        st.dataframe(hh, hide_index=True, use_container_width=True)
    else:
        st.info("推移を描くには 2 日以上の建玉データが必要です")

    # ---- IV（QRI ミニオプション、15 分遅延）
    st.subheader("IV（QRI 日経225ミニオプション、15 分遅延）")
    dstr = d.strftime("%Y%m%d")
    days_sel = [x for x in _mini_days() if x <= dstr]  # 判定・チャートの窓（表示日をアンカーに過去方向）
    ivk = _iv_keys(d); ivk_next = _iv_keys(d + timedelta(days=1))  # 夜間立会分は翌暦日のフォルダに入る
    if not (ivk or ivk_next):
        st.caption("この日の IV スナップショットは未収集です（ミニオプションの QRI 収集は 2026-10-20 開始）。")
    else:
        when = st.radio("表示時点", ["日中引け（気配 15:30 まで）", "最新（夜間を含む）"], horizontal=True, key="wk_when")
        key_sel, ivdf = _pick_close(ivk) if when.startswith("日中") else (None, None)
        if key_sel is None:
            key_sel = (ivk + ivk_next)[-1]; ivdf = _iv_snapshot(key_sel)
            if when.startswith("日中"):
                st.caption("日中引け（気配 08:45〜15:30）のスナップショットが無いため最新を表示")
        sub = ivdf[ivdf.contract_month == exp].copy() if "contract_month" in ivdf.columns else ivdf.iloc[0:0]
        if sub.empty:
            st.caption(f"スナップショット {key_sel.split('/')[-1][9:15]} にこの満期（{exp}）の行がありません")
        else:
            from data.iv_views import with_eff_iv
            sub = with_eff_iv(sub); sub["delta"] = pd.to_numeric(sub.delta, errors="coerce")
            pv = sub.pivot_table(index="strike", columns="option_type", values=["eff_iv", "delta", "oi", "volume"], aggfunc="first")
            tbl = pd.DataFrame({"行使価格": pv.index})
            for ot, lab in (("PUT", "PUT"), ("CALL", "CALL")):
                tbl[f"{lab} IV%"] = (pv[("eff_iv", ot)].values * 100).round(1) if ("eff_iv", ot) in pv.columns else None
                tbl[f"{lab} Δ"] = pv[("delta", ot)].values.round(2) if ("delta", ot) in pv.columns else None
            tbl = tbl[(tbl["PUT IV%"].notna()) | (tbl["CALL IV%"].notna())]
            fn = key_sel.split("/")[-1]; src = ivdf.source_update_time.dropna().astype(str)
            st.caption(f"スナップショット {fn[4:6]}/{fn[6:8]} {fn[9:11]}:{fn[11:13]} 取得（気配時刻 {src.iloc[0][11:16] if len(src) else '-'}、配信は 15 分遅延）。両気配の仲値 IV、無ければ約定・清算ベース。")
            figiv = go.Figure()
            for ot, col in (("PUT", _COL["put"]), ("CALL", _COL["call"])):
                s_ = sub[(sub.option_type == ot) & sub.eff_iv.notna()].sort_values("strike")
                figiv.add_trace(go.Scatter(x=s_.strike, y=s_.eff_iv * 100, mode="lines+markers", name=f"{ot} IV", line=dict(color=col, width=1.6)))
            figiv.update_layout(height=300, template="plotly_white", title=f"{labels[exp]}：IV スマイル", yaxis=dict(title="IV (%)"), xaxis=dict(title="行使価格（円）"), legend=dict(orientation="h", y=-0.3), margin=dict(l=0, r=0, t=40, b=0))
            st.plotly_chart(figiv, use_container_width=True)
            st.dataframe(tbl, hide_index=True, use_container_width=True, height=min(500, 40 + 35 * len(tbl)))
        # ---- IV モニターの流用（Δ建玉×超過ΔIV 判定・日次チャート・日内ローソク足／累積超過ΔIV）
        if days_sel:
            st.subheader("行使価格別 IV×建玉（買われたか・売られたか）")
            st.caption("日経225オプションの IV モニターと同じ判定・チャート。Δ建玉は別紙2、IV はミニオプションの QRI スナップショット。")
            render_strike_section(days_sel, exp, ivdf, n_hist, intra_n, mini=True, fmt_cm=lambda c: labels.get(c, c), key_prefix="wk")

    # ---- 満期一覧
    st.subheader(f"{d:%m/%d} 時点の満期一覧")
    rows = []
    for e in exps:
        fe = _frame(d, e); mpe, _ = wo.max_pain(fe); ie = wo.expiry_info(e)
        rows.append({"満期": ie["label"], "PUT 建玉": int(fe.put_oi.sum()), "CALL 建玉": int(fe.call_oi.sum()),
                     "P/C": round(fe.put_oi.sum() / fe.call_oi.sum(), 2) if fe.call_oi.sum() else None, "マックスペイン": mpe, "取引高": int(fe.put_vol.sum() + fe.call_vol.sum())})
    st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)


main()
