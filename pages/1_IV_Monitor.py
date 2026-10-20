"""IV Monitor — QRI NK225 option IV snapshots (15-min delayed) viewer.

Data: R2 qri_iv/raw/YYYYMMDD/*.parquet (scripts/fetch_qri_iv.py が収集)
主目的: 行使価格別の Δ建玉×ΔIV からフロー（買われたか・売られたか）を判定する。
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
from ui.iv_monitor_section import render_strike_section

st.set_page_config(page_title="IVモニター", layout="wide")

_TTL = 300  # JPX公表・QRI収集とも時刻が揺れるため5分で再取得


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
    key = keys[-1]  # 最終スナップショット固定（日内推移は下のチャートで確認）
    df = _snapshot(key)
    # SQ前日（最終取引日）の夜間以降、QRIは期近を一覧から外す。16:59以前の最終スナップショットも
    # 読み、限月候補はその和集合、表示は選択限月を含むスナップショットを使う
    day_keys = [k for k in keys if iv_views.key_time_label(k) <= "16:59:59"]
    df_day = _snapshot(day_keys[-1]) if day_keys and day_keys[-1] != key else df
    if df.empty:
        st.warning("スナップショットを読み込めませんでした")
        st.stop()
    if len(days) > 5:
        n_hist = st.sidebar.slider("日次チャートの日数", 5, len(days), min(20, len(days)))
    else:
        n_hist = len(days)
    intra_n = st.sidebar.slider("日内チャートの日数", 1, 5, 3)
    # JPXの建玉残高表は公表時刻が日によってずれるため、手動で即時再取得できるようにする
    if st.sidebar.button("最新データに更新", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    upd = df.source_update_time.dropna()
    st.caption(f"最終スナップショット: {iv_views.key_time_label(key)}  |  "
               f"ソース更新: {upd.iloc[0] if len(upd) else '-'}  |  行数: {len(df):,}")

    all_months = sorted(set(df.contract_month.unique()) | set(df_day.contract_month.unique()))
    days_sel = [d for d in days if d <= day]  # 選択日をアンカーに過去方向のみ使用

    # ---------------- フロー判定（主機能） ----------------
    st.subheader("行使価格別 IV×建玉（買われたか・売られたか）")
    sc1, _ = st.columns([1, 4])
    with sc1:
        cm_sel = st.selectbox("限月", all_months, format_func=_fmt_cm, key="ps_cm")
    if cm_sel not in set(df.contract_month.unique()):
        df = df_day
        st.caption(f"{_fmt_cm(cm_sel)}は最終スナップショットに無いため、"
                   f"{iv_views.key_time_label(day_keys[-1])}時点のスナップショットを表示")
    # 判定表・日次チャート・日内チャートは週次（ミニ）ページと共通（ui/iv_monitor_section.py）
    render_strike_section(days_sel, cm_sel, df, n_hist, intra_n, mini=False, fmt_cm=_fmt_cm, key_prefix="ps")


main()
