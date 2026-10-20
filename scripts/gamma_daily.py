"""日次の γ 蓄積ジョブ。指定日（既定: 直近の営業日）の日中引けスナップショットから、

  リスク中立分布の要約（F, T, ATM IV, RR25, 分位点, SVI パラメータ）
  べき効用で傾けた実確率の要約（γ = 0, 2, 4, 6）
  市場が織り込む γ（方法 I カーネル法とその変種、曲率、方法 II モーメント法）
  実確率側の材料（RV20, RV60, 20 日リターン）

を 1 レコードにして R2 の qri_iv/derived/gamma_daily.jsonl に追記する（同じ日は置き換え）。
使い方: python scripts/gamma_daily.py [--day YYYYMMDD] [--since YYYYMMDD] [--no-upload]
GitHub Actions（fetch_daily_oi.yml、JST 18:00 以降の実行）から毎営業日呼ばれる。
"""
from __future__ import annotations

import argparse
import io
import json
import logging
import sys
import warnings
from datetime import date, timedelta
from pathlib import Path

warnings.filterwarnings("ignore")
sys.path.insert(0, str(Path(__file__).resolve().parent)); sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import numpy as np
import pandas as pd

import config
from data import iv_views
from data.r2_storage import r2_get, r2_put
from gamma_lib import Physical, front_cm, gamma_kernel, gamma_moment, p_moments, q_moments
from rnd_lib import day_close_snap, density, fit_svi, prep, rr_delta, skew25, stats, svi, tilt

logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s"); log = logging.getLogger("gamma_daily")
R2_KEY = "qri_iv/derived/gamma_daily.jsonl"; LOCAL = config.CACHE_DIR / "derived" / "gamma_daily.jsonl"
XS = np.linspace(-0.6, 0.6, 6001)


def nikkei_close() -> pd.Series:
    """日経平均の日次終値。yfinance が取れた日は cache/derived/n225_close.csv に保存し、レート制限時はそのキャッシュを使う。"""
    import yfinance as yf
    cache = Path(__file__).resolve().parent.parent / "cache" / "derived" / "n225_close.csv"
    try:
        h = yf.download("^N225", start="1999-06-01", progress=False, auto_adjust=False); h.columns = [c[0] if isinstance(c, tuple) else c for c in h.columns]
        s = h.Close.dropna()
        if len(s) > 1000:
            cache.parent.mkdir(parents=True, exist_ok=True); s.to_csv(cache, header=["Close"]); return s
    except Exception as e:  # noqa: BLE001
        print("yfinance failed:", e)
    try:  # 代替ソース: stooq（^NKX = 日経平均）
        import requests
        r = requests.get("https://stooq.com/q/d/l/?s=^nkx&i=d", timeout=30, headers={"User-Agent": "Mozilla/5.0"}); r.raise_for_status()
        h = pd.read_csv(io.StringIO(r.text), index_col=0, parse_dates=True); s = h["Close"].dropna(); s = s[s.index >= "1999-06-01"]
        if len(s) > 1000:
            cache.parent.mkdir(parents=True, exist_ok=True); s.to_csv(cache, header=["Close"]); print(f"using stooq N225 close through {s.index[-1].date()}"); return s
    except Exception as e:  # noqa: BLE001
        print("stooq failed:", e)
    if cache.exists():
        s = pd.read_csv(cache, index_col=0, parse_dates=True)["Close"]; print(f"using cached N225 close through {s.index[-1].date()}"); return s
    raise RuntimeError("N225 close unavailable (yfinance failed, no cache)")


def compute(day: str, phys: Physical) -> dict | None:
    cm = front_cm(day)
    if cm is None:
        return None
    k = day_close_snap(day, cm, 120)
    if k is None:
        log.info("%s: no day-close snapshot", day); return None
    ts, F, T, pts, g = prep(k, cm)
    if not np.isfinite(F) or len(pts) < 30:
        log.info("%s: insufficient quotes", day); return None
    p, rmse = fit_svi(pts, T); K, fK, gg, ivf = density(p, F, T, XS); st = stats(K, fK)
    if gg.min() < -0.05 or rmse > 0.03:
        log.info("%s: poor fit rmse=%.3f gmin=%.3f", day, rmse, gg.min()); return None
    sa, s25p, s25c = skew25(p, T, XS, ivf)
    d = pd.Timestamp(day); n_days = max(int(round(T * 250)), 3); rv20 = phys.rv_at(d, 20); rv60 = phys.rv_at(d, 60)
    xP = phys.sample(n_days, rv20); res = gamma_kernel(K, fK, F, xP)
    rec = dict(day=day, cm=cm, ts=str(ts), key=k, F=F, T=T, n_days=n_days, n_pts=int(len(pts)), rmse=rmse, gmin=float(gg.min()), p=p.tolist(),
               atm=sa, s25p=s25p, s25c=s25c, rr25_delta=rr_delta(k, cm), mean=st["mean"], sd=st["sd"], skew=st["skew"], kurt=st["kurt"], mode=st["mode"],
               q={str(q): v for q, v in st["q"].items()}, p_lt60=st["cdf"](60000), p_lt65=st["cdf"](65000), p_ge70=1 - st["cdf"](70000),
               rv20=rv20, rv60=rv60, ret20=phys.ret(d, 20),
               gamma_I=res["gamma"], gamma_I_quad=res["gamma_quadfit"], quad=res["quad"], r2=res["r2"],
               gamma_rv60=gamma_kernel(K, fK, F, phys.sample(n_days, rv60))["gamma"], gamma_c80=gamma_kernel(K, fK, F, xP, 0.10, 0.90)["gamma"], gamma_bw=gamma_kernel(K, fK, F, xP, bw=0.12)["gamma"],
               gamma_II=gamma_moment(K, fK, F, xP), tilt={})
    mq, sq_, skq, kuq = q_moments(K, fK, F); mp_, sp_, skp, kup = p_moments(xP)
    rec.update(skew_q=skq, kurt_q=kuq, skew_p=skp, kurt_p=kup, s_q=sq_, s_p=sp_)
    for gm in (0.0, 2.0, 4.0, 6.0):
        stp = stats(K, tilt(K, fK, gm)) if gm else st
        rec["tilt"][str(gm)] = dict(mean=stp["mean"], median=stp["q"][0.5], mode=stp["mode"], q05=stp["q"][0.05], q95=stp["q"][0.95], p_lt65=stp["cdf"](65000), p_ge70=1 - stp["cdf"](70000))
    return rec


def load_store(upload: bool) -> list[dict]:
    raw = r2_get(R2_KEY) if upload else None
    if raw is None and LOCAL.exists():
        raw = LOCAL.read_bytes()
    return [json.loads(l) for l in raw.decode("utf-8").splitlines() if l.strip()] if raw else []


def save_store(rows: list[dict], upload: bool) -> None:
    rows = sorted({r["day"]: r for r in rows}.values(), key=lambda r: r["day"])
    body = ("\n".join(json.dumps(r, ensure_ascii=False) for r in rows) + "\n").encode("utf-8")
    LOCAL.parent.mkdir(parents=True, exist_ok=True); LOCAL.write_bytes(body)
    if upload:
        r2_put(R2_KEY, body)


def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--day"); ap.add_argument("--since"); ap.add_argument("--no-upload", action="store_true"); ap.add_argument("--force", action="store_true")
    a = ap.parse_args(); upload = not a.no_upload
    phys = Physical(nikkei_close()); store = load_store(upload); have = {r["day"] for r in store}
    if a.since:
        days = [d for d in sorted(iv_views.list_available_days(400)) if d >= a.since and pd.Timestamp(d).weekday() < 5]
    elif a.day:
        days = [a.day]
    else:
        today = date.today(); days = [(today - timedelta(days=i)).strftime("%Y%m%d") for i in range(0, 6) if (today - timedelta(days=i)).weekday() < 5][:1]
        # 当日の引けがまだ無ければ直前の営業日を対象にする
        if not iv_views.snapshot_keys(days[0]):
            days = [(today - timedelta(days=i)).strftime("%Y%m%d") for i in range(1, 7) if (today - timedelta(days=i)).weekday() < 5][:1]
    n_new = 0
    for d in days:
        if d in have and not a.force:
            log.info("%s: already stored", d); continue
        try:
            rec = compute(d, phys)
        except Exception as e:
            log.warning("%s: failed: %s", d, e); continue
        if rec is None:
            continue
        store = [r for r in store if r["day"] != d] + [rec]; n_new += 1
        log.info("%s %s F=%.0f atm=%.1f%% gamma_I=%.2f gamma_II=%.2f rv20=%.1f%%", d, rec["cm"], rec["F"], rec["atm"] * 100, rec["gamma_I"], rec["gamma_II"], rec["rv20"] * 100)
        if n_new % 10 == 0:
            save_store(store, upload)
    if n_new:
        save_store(store, upload)
    log.info("done: %d new, %d total", n_new, len({r['day'] for r in store}))


if __name__ == "__main__":
    main()
