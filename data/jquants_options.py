"""J-Quants 日経225オプション日次データ（scripts/fetch_jquants_options.py が保存した parquet）の読み込み。

R2 jquants/op225/YYYYMMDD.parquet → ローカル cache/jquants/op225/ の順。無ければ None。
列: date, cm('YYMM'), option_type, strike, vo, vooa, vo_off, oi, settle, theo, iv, basevol, underpx, ir, ltd, sqd, code
"""
from __future__ import annotations

import io
import logging
from datetime import date

import pandas as pd

import config
from data.r2_storage import r2_get

logger = logging.getLogger(__name__)
R2_PREFIX = "jquants/op225"
LOCAL = config.CACHE_DIR / "jquants" / "op225"
_CACHE: dict[str, pd.DataFrame | None] = {}


def load_day(d: date | str) -> pd.DataFrame | None:
    """取引日 d の全行使・全限月。未取得なら None。"""
    key = d if isinstance(d, str) else d.strftime("%Y%m%d")
    if key in _CACHE:
        return _CACHE[key]
    df = None
    try:
        b = r2_get(f"{R2_PREFIX}/{key}.parquet")
        if b is None and (LOCAL / f"{key}.parquet").exists():
            b = (LOCAL / f"{key}.parquet").read_bytes()
        if b is not None:
            df = pd.read_parquet(io.BytesIO(b))
    except Exception as e:  # noqa: BLE001
        logger.warning("jquants load failed %s: %s", key, e)
    if len(_CACHE) > 64:
        _CACHE.clear()
    _CACHE[key] = df
    return df


def volume_split(d: date | str, cm: str) -> pd.DataFrame | None:
    """限月 cm の (option_type, strike) ごとの 立会内 vooa / 立会外 vo_off / 合計 vo。"""
    df = load_day(d)
    if df is None or df.empty:
        return None
    g = df[df.cm == cm][["option_type", "strike", "vo", "vooa", "vo_off"]].copy()
    g["strike"] = g.strike.astype(int)
    return g.reset_index(drop=True)
