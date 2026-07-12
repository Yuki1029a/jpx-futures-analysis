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


def strike_daily_series(days: list[str], cm: str, option_type: str,
                        strikes: list[int] | None = None) -> pd.DataFrame:
    """Per-strike daily series (last snapshot per day) for one month x type.

    Returns: day(date), strike, iv_pct(eff_iv), oi, volume.
    """
    rows = []
    for d in days:
        keys = snapshot_keys(d)
        if not keys:
            continue
        df = load_snapshot(keys[-1])
        if df is None or df.empty:
            continue
        m = with_eff_iv(df)
        g = m[(m.contract_month == cm) & (m.option_type == option_type)]
        if strikes:
            g = g[g.strike.isin(strikes)]
        dt = date(int(d[:4]), int(d[4:6]), int(d[6:8]))
        for _, r in g.iterrows():
            rows.append({
                "day": dt, "strike": int(r.strike),
                "iv_pct": None if pd.isna(r.eff_iv) else r.eff_iv * 100.0,
                "oi": None if pd.isna(r.oi) else float(r.oi),
                "volume": None if pd.isna(r.volume) else float(r.volume),
            })
    return pd.DataFrame(rows)


def strike_intraday_series(day_str: str, cm: str, option_type: str,
                           strikes: list[int]) -> pd.DataFrame:
    """Per-strike IV across all snapshots of one day. Returns: time, strike, iv_pct."""
    rows = []
    for key in snapshot_keys(day_str):
        df = load_snapshot(key)
        if df is None or df.empty:
            continue
        m = with_eff_iv(df)
        g = m[(m.contract_month == cm) & (m.option_type == option_type) & m.strike.isin(strikes)]
        tl = key_time_label(key)
        for _, r in g.iterrows():
            if pd.notna(r.eff_iv):
                rows.append({"time": tl, "strike": int(r.strike), "iv_pct": r.eff_iv * 100.0})
    return pd.DataFrame(rows)


_QUAD = {(1, 1): "新規買い", (1, -1): "新規売り", (-1, 1): "買い戻し", (-1, -1): "手仕舞い"}
_NEUTRAL_BAND = 0.05  # %pt: |超過ΔIV| がこれ未満は方向を断定しない


def flow_judgement(days: list[str], cm: str) -> tuple[pd.DataFrame, dict]:
    """Δ建玉×ΔIV quadrant per (strike, type), IV窓をOIの取引日窓に整合させる。

    QRIのoiはJPX清算後残高のT+1反映のため、データ日 d の最終スナップショットの
    OIは「dより前の直近取引日」引け残。各日最終スナップを新しい順に snaps とし、
      k  = OIが snaps[0] から最初に動いたデータ日 index
      k2 = snaps[k] からさらに動いた index
    とすると Δ建玉 = OI[0]−OI[k] は取引日窓 (day[k2], day[k]] の売買を反映する。
    同じ取引日窓の気配変化は ΔIV = eff_iv[k]−eff_iv[k2]（1データ日前へシフト）。
    地合い（サーフェス平行移動）は限月内・両気配ストライクのΔIV中央値 level で
    控除し、超過ΔIV = ΔIV−level の符号で判定する。
    Returns (df, meta):
      df   : option_type, strike, oi, d_oi, iv_pct, d_iv_pct, d_iv_ex_pct,
             volume, judge
      meta : oi_days=(当日,比較日), iv_days=(IV当日,IV比較日),
             level_pct, n_level, aligned(bool: IV窓シフト成立)
    """
    snaps: list[tuple[str, pd.DataFrame]] = []
    for d in reversed(days):
        keys = snapshot_keys(d)
        if not keys:
            continue
        df = load_snapshot(keys[-1])
        if df is not None and not df.empty:
            snaps.append((d, with_eff_iv(df)))
        if len(snaps) >= 10:
            break
    meta = {"oi_days": None, "iv_days": None,
            "level_pct": 0.0, "n_level": 0, "aligned": True}
    if len(snaps) < 2:
        return pd.DataFrame(), meta

    def _slice(df):
        g = df[df.contract_month == cm]
        g = g[~g.duplicated(subset=["option_type", "strike"])]
        return g.set_index(["option_type", "strike"])

    sl = [(d, _slice(df)) for d, df in snaps]

    def _oi_moved(a, b):
        common = a.index.intersection(b.index)
        if not len(common):
            return False
        return (a.oi.loc[common].fillna(0) - b.oi.loc[common].fillna(0)).abs().sum() > 0

    k = next((i for i in range(1, len(sl)) if _oi_moved(sl[0][1], sl[i][1])), None)
    if k is None:
        return pd.DataFrame(), meta
    k2 = next((j for j in range(k + 1, len(sl)) if _oi_moved(sl[k][1], sl[j][1])), None)
    if k2 is None:
        ivc_i, ivp_i = 0, k  # シフト先データ不足 → 未整合のまま近似
        meta["aligned"] = False
    else:
        ivc_i, ivp_i = k, k2

    cur, prv = sl[0][1], sl[k][1]
    ivc, ivp = sl[ivc_i][1], sl[ivp_i][1]
    meta["oi_days"] = (sl[0][0], sl[k][0])
    meta["iv_days"] = (sl[ivc_i][0], sl[ivp_i][0])

    idx = cur.index.intersection(prv.index)
    rows = []
    lvl_q, lvl_all = [], []
    for key in idx:
        rc, rp = cur.loc[key], prv.loc[key]
        d_oi = None
        if pd.notna(rc.oi) and pd.notna(rp.oi):
            d_oi = float(rc.oi - rp.oi)
        c_iv = ivc.eff_iv.get(key)
        p_iv = ivp.eff_iv.get(key)
        d_iv = None
        if c_iv is not None and p_iv is not None and pd.notna(c_iv) and pd.notna(p_iv):
            d_iv = float(c_iv - p_iv) * 100.0
            lvl_all.append(d_iv)
            if (pd.notna(ivc.ask_iv.get(key)) and pd.notna(ivc.bid_iv.get(key))
                    and pd.notna(ivp.ask_iv.get(key)) and pd.notna(ivp.bid_iv.get(key))):
                lvl_q.append(d_iv)
        rows.append({
            "option_type": key[0], "strike": int(key[1]),
            "oi": None if pd.isna(rc.oi) else float(rc.oi),
            "d_oi": d_oi,
            "iv_pct": None if pd.isna(rc.eff_iv) else float(rc.eff_iv) * 100.0,
            "d_iv_pct": d_iv,
            "volume": None if pd.isna(rc.volume) else float(rc.volume),
        })
    out = pd.DataFrame(rows)
    if not len(out):
        return out, meta

    pool = lvl_q if len(lvl_q) >= 5 else lvl_all
    level = float(pd.Series(pool).median()) if pool else 0.0
    meta["level_pct"] = level
    meta["n_level"] = len(pool)
    out["d_iv_ex_pct"] = out.d_iv_pct - level

    def _judge(r):
        if pd.isna(r.d_oi) or r.d_oi == 0 or pd.isna(r.d_iv_ex_pct):
            return ""
        if abs(r.d_iv_ex_pct) < _NEUTRAL_BAND:
            return "中立"
        return _QUAD[(1 if r.d_oi > 0 else -1, 1 if r.d_iv_ex_pct > 0 else -1)]

    out["judge"] = out.apply(_judge, axis=1)
    return out, meta


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
