"""Pure data-prep helpers for the IV Monitor page (no streamlit imports).

Source: QRI NK225 option IV snapshots collected by scripts/fetch_qri_iv.py
  R2:    qri_iv/raw/YYYYMMDD/YYYYMMDD_HHMMSS.parquet
  Local: cache/qri_iv/raw/...

Each snapshot: one row per (contract_month, strike, option_type) with
settle/oi/volume/iv/ask_iv/bid_iv/greeks. iv columns are fractions (0.25 = 25%).
"""
from __future__ import annotations

import logging
from datetime import date, datetime

import pandas as pd

from data import r2_storage
from data.qri_iv_loader import _R2_PREFIX, _LOCAL_ROOT, list_snapshot_keys, load_snapshot

logger = logging.getLogger(__name__)


def list_available_days(limit: int = 60) -> list[str]:
    """Days (YYYYMMDD) that have at least one snapshot. Ascending. R2 + local union."""
    days: set[str] = set()
    r2_storage._init_client()
    if r2_storage._client is not None:
        try:
            resp = r2_storage._client.list_objects_v2(
                Bucket=r2_storage._bucket, Prefix=f"{_R2_PREFIX}/", Delimiter="/"
            )
            for p in resp.get("CommonPrefixes", []):
                d = p["Prefix"].rstrip("/").split("/")[-1]
                if len(d) == 8 and d.isdigit():
                    days.add(d)
        except Exception as e:
            logger.warning("list_available_days R2 failed: %s", e)
    if _LOCAL_ROOT.exists():
        for p in _LOCAL_ROOT.iterdir():
            if p.is_dir() and len(p.name) == 8 and p.name.isdigit():
                days.add(p.name)
    return sorted(days)[-limit:]


def snapshot_keys(day_str: str) -> list[str]:
    """Snapshot keys of a day, ascending (oldest first)."""
    d = datetime.strptime(day_str, "%Y%m%d").date()
    keys = list_snapshot_keys(d)
    if keys:
        return keys
    local_dir = _LOCAL_ROOT / day_str
    if local_dir.exists():
        return [f"{_R2_PREFIX}/{day_str}/{p.name}" for p in sorted(local_dir.glob("*.parquet"))]
    return []


def key_time_label(key: str) -> str:
    """'qri_iv/raw/20260709/20260709_051536.parquet' -> '05:15:36'"""
    stem = key.rsplit("/", 1)[-1].split(".")[0]
    t = stem.split("_")[-1]
    return f"{t[:2]}:{t[2:4]}:{t[4:6]}" if len(t) == 6 else t


# ---------------------------------------------------------------- metrics ----
# QRIの 'iv' 列は約定/清算ベースで更新が疎（夜間・閑散ストライクで据置き）。
# ザラ場の実勢は気配IV(ask_iv/bid_iv)に出るため、指標は「気配ミッド優先、
# 両気配欠損時のみ 'iv'」の eff_iv で計算する。

def with_eff_iv(df: pd.DataFrame) -> pd.DataFrame:
    """Add eff_iv = mid(ask_iv, bid_iv) when both quoted, else traded/settle iv."""
    m = df.copy()
    mid = (m.ask_iv + m.bid_iv) / 2.0
    m["eff_iv"] = mid.where(m.ask_iv.notna() & m.bid_iv.notna(), m["iv"])
    return m


def _atm_row(g: pd.DataFrame) -> pd.Series | None:
    """ATM row of one (snapshot x month): CALL delta nearest 0.5, fallback |c_iv - p_iv| min.

    g must carry eff_iv (see with_eff_iv). Returned row's eff_iv is the ATM IV.
    """
    c = g[(g.option_type == "CALL") & g.delta.notna() & g.eff_iv.notna()]
    if len(c):
        return c.loc[(c.delta - 0.5).abs().idxmin()]
    piv = g[g.eff_iv.notna()].pivot_table(
        index="strike", columns="option_type", values="eff_iv", aggfunc="first")
    if {"CALL", "PUT"} <= set(piv.columns):
        both = piv.dropna()
        if len(both):
            k = (both.CALL - both.PUT).abs().idxmin()
            c2 = g[(g.option_type == "CALL") & (g.strike == k) & g.eff_iv.notna()]
            if len(c2):
                return c2.iloc[0]
    return None


def _delta_iv(g: pd.DataFrame, side: str, target: float) -> float | None:
    """eff_iv at the strike whose delta is nearest to target (side='CALL'/'PUT')."""
    s = g[(g.option_type == side) & g.delta.notna() & g.eff_iv.notna()]
    if not len(s):
        return None
    return float(s.loc[(s.delta - target).abs().idxmin(), "eff_iv"])


def month_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Per contract_month: ATM strike/IV, 25d skew, OI/volume totals.

    skew25 = IV(PUT delta -0.25) - IV(CALL delta 0.25)  (fraction)
    """
    rows = []
    for cm, g in with_eff_iv(df).groupby("contract_month"):
        atm = _atm_row(g)
        c25 = _delta_iv(g, "CALL", 0.25)
        p25 = _delta_iv(g, "PUT", -0.25)
        rows.append({
            "contract_month": cm,
            "atm_strike": None if atm is None else atm.strike,
            "atm_iv": None if atm is None else atm.eff_iv,
            "skew25": None if (c25 is None or p25 is None) else p25 - c25,
            "oi_total": g.oi.sum(),
            "vol_total": g.volume.sum(),
        })
    return pd.DataFrame(rows).sort_values("contract_month").reset_index(drop=True)


def smile_frame(df: pd.DataFrame, months: list[str]) -> pd.DataFrame:
    """Long frame for smile plot: contract_month, option_type, strike, iv(%), oi."""
    m = with_eff_iv(df)
    m = m[m.contract_month.isin(months) & m.eff_iv.notna()]
    m["iv_pct"] = m.eff_iv * 100.0
    return m[["contract_month", "option_type", "strike", "iv_pct", "oi"]].sort_values("strike")


def intraday_atm(day_str: str) -> pd.DataFrame:
    """ATM IV per month across all snapshots of a day.

    Returns columns: time (label), contract_month, atm_iv_pct.
    """
    rows = []
    for key in snapshot_keys(day_str):
        df = load_snapshot(key)
        if df is None or df.empty:
            continue
        tl = key_time_label(key)
        for cm, g in with_eff_iv(df).groupby("contract_month"):
            atm = _atm_row(g)
            if atm is not None and pd.notna(atm.eff_iv):
                rows.append({"time": tl, "contract_month": cm, "atm_iv_pct": atm.eff_iv * 100.0})
    return pd.DataFrame(rows)


def daily_atm_series(days: list[str]) -> pd.DataFrame:
    """Last-snapshot ATM IV / skew per month for each day.

    Returns columns: day (date), contract_month, atm_iv_pct, skew25_pct.
    """
    rows = []
    for d in days:
        keys = snapshot_keys(d)
        if not keys:
            continue
        df = load_snapshot(keys[-1])
        if df is None or df.empty:
            continue
        dt = date(int(d[:4]), int(d[4:6]), int(d[6:8]))
        for cm, g in with_eff_iv(df).groupby("contract_month"):
            atm = _atm_row(g)
            c25 = _delta_iv(g, "CALL", 0.25)
            p25 = _delta_iv(g, "PUT", -0.25)
            rows.append({
                "day": dt,
                "contract_month": cm,
                "atm_iv_pct": None if atm is None or pd.isna(atm.eff_iv) else atm.eff_iv * 100.0,
                "skew25_pct": None if (c25 is None or p25 is None) else (p25 - c25) * 100.0,
            })
    return pd.DataFrame(rows)
