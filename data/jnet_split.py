"""行使価格別の立会内／立会外（J-NET）取引高。

出所（優先）: JPX 日報 siop_dyr PDF（Daily_Report_OSE_{d}.zip、翌営業日 09:00 頃公表）の「競争売買市場」「J-NET市場」節。
補完      : J-Quants 日経225オプション日次（VoOA = 立会内、Vo − VoOA = 立会外。翌 03:00 頃更新、月次のみ）。
日報が出るまでの朝（03:00〜09:00）は J-Quants、以降は日報が使われる。
両者は 2026-10-08 に全行使で一致を確認済み。

保存: R2 jnet/siop/YYYYMMDD.parquet、ローカル cache/jnet/siop/（scripts/build_jnet_split.py が作成）
列  : product('op225'|'mini'), cm（月次 'YYMM'、ミニ 'YYMMDD' = 最終取引日 = 別紙2 の満期コード）,
      option_type, strike, auction_vol（立会内、日中＋夜間）, jnet_vol（立会外、日中＋夜間）
行使価格別のセッション（日中／夜間）内訳は日報に無い。
"""
from __future__ import annotations

import io
import logging
from datetime import date

import pandas as pd

import config
from data.r2_storage import r2_get

logger = logging.getLogger(__name__)
R2_PREFIX = "jnet/siop"
LOCAL = config.CACHE_DIR / "jnet" / "siop"
PRODUCTS = {"日経225オプション": "op225", "日経225ミニオプション": "mini"}
_CACHE: dict[str, pd.DataFrame] = {}


def _key(d: date | str) -> str:
    return d if isinstance(d, str) else d.strftime("%Y%m%d")


def build_from_siop(content: bytes) -> tuple[date, pd.DataFrame]:
    """siop_dyr PDF → (日付, 行使価格別の立会内／立会外取引高)。"""
    from data.parser_daily_report_options import parse_siop_pdf
    d, rows = parse_siop_pdf(content, tuple(PRODUCTS))
    df = pd.DataFrame(rows)
    cols = ["product", "cm", "option_type", "strike", "auction_vol", "jnet_vol"]
    if df.empty:
        return d, pd.DataFrame(columns=cols)
    df["product"] = df["product"].map(PRODUCTS)

    def code(r) -> str:
        if r["product"] != "mini":
            return r["contract_month"]                      # 'YYMM'
        sq = r["contract_month"]                            # ミニは SQ 日 'YYMMDD'
        lm, ld = (int(x) for x in str(r["last_day"]).split("."))
        yy = int(sq[:2]) - (1 if lm > int(sq[2:4]) else 0)  # 年跨ぎ（12 月最終取引 → 1 月 SQ）
        return f"{yy:02d}{lm:02d}{ld:02d}"

    df["cm"] = df.apply(code, axis=1)
    g = df.pivot_table(index=["product", "cm", "option_type", "strike_price"], columns="market",
                       values="trading_volume", aggfunc="sum", fill_value=0).reset_index()
    g = g.rename(columns={"strike_price": "strike", "auction": "auction_vol", "jnet": "jnet_vol"})
    for c in ("auction_vol", "jnet_vol"):
        if c not in g.columns:
            g[c] = 0
    g["strike"] = g.strike.astype(int)
    return d, g[cols].reset_index(drop=True)


def load(d: date | str) -> pd.DataFrame | None:
    """取引日 d の日報由来の内訳。未作成なら None（None はキャッシュしない）。"""
    k = _key(d)
    if k in _CACHE:
        return _CACHE[k]
    df = None
    try:
        b = r2_get(f"{R2_PREFIX}/{k}.parquet")
        if b is None and (LOCAL / f"{k}.parquet").exists():
            b = (LOCAL / f"{k}.parquet").read_bytes()
        if b is not None:
            df = pd.read_parquet(io.BytesIO(b))
    except Exception as e:  # noqa: BLE001
        logger.warning("jnet split load failed %s: %s", k, e)
    if df is not None:
        if len(_CACHE) > 64:
            _CACHE.clear()
        _CACHE[k] = df
    return df


def volume_split(d: date | str, cm: str, mini: bool = False) -> tuple[pd.DataFrame | None, str | None]:
    """限月（満期）cm の (option_type, strike) ごとの立会外取引高 vo_off と出所。

    日報があれば日報（J-NET 節に無い行使は立会外 0）、無ければ J-Quants（月次のみ）。どちらも無ければ (None, None)。
    立会内は呼び出し側で「建玉残高表の取引高 − vo_off」とする（両者の取引高合計は一致）。
    """
    df = load(d)
    if df is not None:
        sub = df[(df["product"] == ("mini" if mini else "op225")) & (df.cm == cm)]
        return sub[["option_type", "strike", "jnet_vol"]].rename(columns={"jnet_vol": "vo_off"}).reset_index(drop=True), "日報"
    if mini:
        return None, None
    from data.jquants_options import volume_split as jq_split
    jq = jq_split(d, cm)
    if jq is None:
        return None, None
    return jq[["option_type", "strike", "vo_off"]].reset_index(drop=True), "J-Quants"
