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
    """Add eff_iv = mid(ask_iv, bid_iv) when both quoted, else traded/settle iv.

    全行欠損の列はparquetでnull型→pandasでobject dtypeになり算術がTypeError
    になるため、参照する数値列はここで一括して強制float化する。
    """
    m = df.copy()
    for c in ("ask_iv", "bid_iv", "iv", "oi", "volume"):
        if c in m.columns:
            m[c] = pd.to_numeric(m[c], errors="coerce")
    m["eff_iv"] = ((m.ask_iv + m.bid_iv) / 2.0).where(
        m.ask_iv.notna() & m.bid_iv.notna(), m["iv"])
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


def strike_daily_oi_jpx(days: list[str], cm: str, option_type: str,
                        strikes: list[int]) -> pd.DataFrame:
    """JPX建玉残高表ベースの行使価格別・日次建玉残（取引日のみ）。

    QRIのoi列は2026-07-12以前の収集分に「4桁以上は先頭3桁のみ」の破損が
    あるため、建玉はJPX公表値を正とする。別紙1に無い限月（遠い四半期限月）
    や未公表日は行が生成されない。
    Returns: day(date), strike, oi.
    """
    rows = []
    for d in days:
        dt = date(int(d[:4]), int(d[4:6]), int(d[6:8]))
        jpx = _jpx_daily_oi_frame(dt, cm)
        if jpx is None or not len(jpx):
            continue
        g = jpx[(jpx.option_type == option_type) & jpx.strike.isin(strikes)]
        for _, r in g.iterrows():
            rows.append({"day": dt, "strike": int(r.strike), "oi": r.oi})
    return pd.DataFrame(rows)


def strike_daily_series(days: list[str], cm: str, option_type: str,
                        strikes: list[int] | None = None) -> pd.DataFrame:
    """Per-strike daily series (16:59以前・有効IVありの最終スナップショット).

    Returns: day(date), strike, iv_pct(eff_iv), oi, volume.
    """
    rows = []
    for d in days:
        kd = _day_iv_snapshot(d, cm)
        if kd is None or kd[1].empty:
            continue
        m = with_eff_iv(kd[1])
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


def strike_intraday_series(days: list[str], n_days: int, cm: str, option_type: str,
                           strikes: list[int]) -> pd.DataFrame:
    """直近n取引日の全スナップショット横断の行使別IV・累積出来高。

    days（昇順、末尾=アンカー日）から取引日を n_days 日選ぶ（アンカー日は
    JPXファイル未公表の当日でも常に含める）。volume はQRI表示の累積出来高
    （取引日ごとにリセット）。
    Returns: time("M/D HH:MM"), day, strike, iv_pct, volume.
    """
    if not days:
        return pd.DataFrame()
    anchor = days[-1]
    tdays = [d for d in days if d == anchor or _is_trading_day(d)][-n_days:]
    rows = []
    for d in tdays:
        for key in snapshot_keys(d):
            df = load_snapshot(key)
            if df is None or df.empty:
                continue
            m = with_eff_iv(df)
            g = m[(m.contract_month == cm) & (m.option_type == option_type)
                  & m.strike.isin(strikes)]
            tl = f"{int(d[4:6])}/{int(d[6:8])} {key_time_label(key)[:5]}"
            ts = pd.Timestamp(f"{d[:4]}-{d[4:6]}-{d[6:8]} {key_time_label(key)}")
            for _, r in g.iterrows():
                if pd.isna(r.eff_iv) and pd.isna(r.volume):
                    continue
                rows.append({
                    "time": tl, "ts": ts, "day": d, "strike": int(r.strike),
                    "iv_pct": None if pd.isna(r.eff_iv) else r.eff_iv * 100.0,
                    "volume": None if pd.isna(r.volume) else float(r.volume),
                })
    return pd.DataFrame(rows)


_QUAD = {(1, 1): "新規買い", (1, -1): "新規売り", (-1, 1): "買い戻し", (-1, -1): "手仕舞い"}
_NEUTRAL_BAND = 0.05  # %pt: |超過ΔIV| がこれ未満は方向を断定しない


def _iv_slice(df: pd.DataFrame, cm: str) -> pd.DataFrame:
    """1スナップショットを (option_type, strike[int]) index に整形（eff_iv付与）。"""
    g = with_eff_iv(df)
    g = g[g.contract_month == cm]
    g = g[~g.duplicated(subset=["option_type", "strike"])].copy()
    g["strike"] = g["strike"].astype(int)
    return g.set_index(["option_type", "strike"])


def _jpx_daily_oi_frame(trade_day: date, cm: str) -> pd.DataFrame | None:
    """JPX建玉残高表（当日20:00頃公表・系列別）の1限月分。未公表・休日はNone。"""
    from data.aggregator import _load_daily_oi_for_date
    try:
        recs = _load_daily_oi_for_date(trade_day, cm)
    except Exception as e:
        logger.warning("JPX daily OI load failed for %s: %s", trade_day, e)
        return None
    if not recs:
        return None
    return pd.DataFrame([{
        "option_type": r.option_type, "strike": int(r.strike_price),
        "oi": float(r.current_oi), "d_oi": float(r.net_change),
        "volume": float(r.trading_volume),
    } for r in recs])


def _day_last_key_1659(day_str: str) -> str | None:
    """その日の「16:59:59以前で最後」のスナップショットキー（無ければ当日最終）。

    17:00以降の夕場スナップショットは翌取引日分の値動きを含むため、
    日次系列・IV窓の端からは除外する（取引日＝前日ナイト＋当日日中に対応）。
    """
    keys = snapshot_keys(day_str)
    if not keys:
        return None
    pre = [k for k in keys
           if k.rsplit("_", 1)[-1].split(".")[0] <= "165959"]
    return pre[-1] if pre else keys[-1]


def _day_iv_snapshot(day_str: str, cm: str, min_rows: int = 5,
                     max_walk: int = 6) -> tuple[str, pd.DataFrame] | None:
    """16:59以前で「当該限月のIVが実際に入っている」最後の (key, df)。

    引け後スナップショットは板クリアでIVが全欠損の日がある（例: 7/7 16:53）。
    その場合は最大 max_walk 本さかのぼり、有効IVが min_rows 以上ある
    スナップショットを使う。全滅なら16:59以前の最終（空でも）を返す。
    """
    keys = snapshot_keys(day_str)
    if not keys:
        return None
    pre = [k for k in keys
           if k.rsplit("_", 1)[-1].split(".")[0] <= "165959"] or keys
    last: tuple[str, pd.DataFrame] | None = None
    for k in reversed(pre[-max_walk:]):
        df = load_snapshot(k)
        if df is None or df.empty:
            continue
        if last is None:
            last = (k, df)
        g = with_eff_iv(df)
        g = g[g.contract_month == cm]
        if len(g) and int(g.eff_iv.notna().sum()) >= min_rows:
            return k, df
    return last


def _is_trading_day(day_str: str) -> bool:
    """JPX建玉残高表の有無で取引日判定（休日・週末はファイルなし）。

    IV窓の起点を直前「取引日」の最終スナップショットに限定するために使う。
    週末スナップショット（ナイトセッション終了後の静止気配）を起点にすると
    金曜ナイト分の変動が窓から漏れるため。
    """
    from data import fetcher
    try:
        return fetcher.download_daily_oi_excel(
            date(int(day_str[:4]), int(day_str[4:6]), int(day_str[6:8]))) is not None
    except Exception:
        return False


def _level_judge(out: pd.DataFrame, lvl_q: list, lvl_all: list,
                 meta: dict) -> tuple[pd.DataFrame, dict]:
    """地合い（限月内・両気配ストライクのΔIV中央値）を控除し4象限判定を付与。"""
    if not len(out):
        return out, meta
    for c in ("oi", "d_oi", "iv_pct", "d_iv_pct", "volume"):
        out[c] = pd.to_numeric(out[c], errors="coerce")  # 全None列のobject化を防ぐ
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


def flow_judgement(days: list[str], cm: str) -> tuple[pd.DataFrame, dict]:
    """Δ建玉×ΔIV quadrant per (strike, type)。

    第1候補（source="jpx"）: JPX建玉残高表（当日20:00頃公表・系列別）。
      取引日Tの増減(net change)をΔ建玉とし、ΔIVは直前取引日→Tの
      「各日16:59以前の最終スナップショット」差（夕場＝翌取引日分を除外）。
      取引日Tのセッション（T-1ナイト＋T日中）と窓が定義から一致する。
    フォールバック（source="qri"）: QRIのOI欄は翌取引日朝反映のため、
      OIが実際に動いた2時点ペアを探し、IV窓を1段前へシフトして整合させる。
    地合い（サーフェス平行移動）は限月内・両気配ストライクのΔIV中央値 level
    で控除し、超過ΔIV = ΔIV−level の符号で判定する。
    Returns (df, meta):
      df   : option_type, strike, oi, d_oi, iv_pct, d_iv_pct, d_iv_ex_pct,
             volume, judge
      meta : source("jpx"/"qri"/None), trade_day(jpx時), oi_days(qri時),
             iv_days, level_pct, n_level, aligned
    """
    meta = {"source": None, "trade_day": None, "oi_days": None, "iv_days": None,
            "level_pct": 0.0, "n_level": 0, "aligned": True}
    desc = [d for d in reversed(days)]

    # ---- 1) JPX建玉残高表（当日分・系列別） ----
    for i, d in enumerate(desc[:6]):
        jpx = _jpx_daily_oi_frame(date(int(d[:4]), int(d[4:6]), int(d[6:8])), cm)
        if jpx is None or not len(jpx):
            continue
        cur_kd = _day_iv_snapshot(d, cm)
        prv_day = next((x for x in desc[i + 1:i + 7]
                        if snapshot_keys(x) and _is_trading_day(x)), None)
        if cur_kd is None or prv_day is None:
            continue
        prv_kd = _day_iv_snapshot(prv_day, cm)
        if prv_kd is None:
            continue
        cur_df, prv_df = cur_kd[1], prv_kd[1]
        if cur_df.empty or prv_df.empty:
            continue
        ci, pi = _iv_slice(cur_df, cm), _iv_slice(prv_df, cm)
        rows, lvl_q, lvl_all = [], [], []
        for _, r in jpx.iterrows():
            key = (r.option_type, int(r.strike))
            cv, pv = ci.eff_iv.get(key), pi.eff_iv.get(key)
            d_iv = None
            if cv is not None and pv is not None and pd.notna(cv) and pd.notna(pv):
                d_iv = float(cv - pv) * 100.0
                lvl_all.append(d_iv)
                if (pd.notna(ci.ask_iv.get(key)) and pd.notna(ci.bid_iv.get(key))
                        and pd.notna(pi.ask_iv.get(key)) and pd.notna(pi.bid_iv.get(key))):
                    lvl_q.append(d_iv)
            rows.append({
                "option_type": r.option_type, "strike": int(r.strike),
                "oi": r.oi, "d_oi": r.d_oi,
                "iv_pct": None if cv is None or pd.isna(cv) else float(cv) * 100.0,
                "d_iv_pct": d_iv,
                "volume": r.volume,
            })
        meta.update(source="jpx", trade_day=d, iv_days=(d, prv_day))
        return _level_judge(pd.DataFrame(rows), lvl_q, lvl_all, meta)

    # ---- 2) フォールバック: QRIのOI欄（翌取引日朝反映）ペア方式 ----
    snaps: list[tuple[str, pd.DataFrame]] = []
    for d in desc:
        kd = _day_iv_snapshot(d, cm)
        if kd is None or kd[1].empty:
            continue
        snaps.append((d, kd[1]))
        if len(snaps) >= 10:
            break
    if len(snaps) < 2:
        return pd.DataFrame(), meta

    sl = [(d, _iv_slice(df, cm)) for d, df in snaps]

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
    meta.update(source="qri", oi_days=(sl[0][0], sl[k][0]),
                iv_days=(sl[ivc_i][0], sl[ivp_i][0]))

    idx = cur.index.intersection(prv.index)
    rows, lvl_q, lvl_all = [], [], []
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
    return _level_judge(pd.DataFrame(rows), lvl_q, lvl_all, meta)


def daily_atm_series(days: list[str]) -> pd.DataFrame:
    """ATM IV / skew per month for each day (16:59以前の最終スナップショット).

    Returns columns: day (date), contract_month, atm_iv_pct, skew25_pct.
    """
    rows = []
    for d in days:
        key = _day_last_key_1659(d)
        if key is None:
            continue
        df = load_snapshot(key)
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
