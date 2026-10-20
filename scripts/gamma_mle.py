"""方法 III（最尤法）: 実現 SQ 値の尤度で γ を推定する。格子は負側（−6）から +12。

対象:
  月次 … 日経225オプション各限月。予測時点は残存 28 日に最も近い営業日の引け（gamma_daily / gamma_calc の系列を再利用）。
  週次 … 日経225ミニオプション（水・金 SQ）。予測時点は、その満期の両気配が min_q 本以上ある最も早い営業日の引け
         （気配時刻 15:30:59 以前の最終、残存 1.5 日以上）。QRI のミニ IV は 2026-10-20 収集開始。
実現 SQ 値: JPX sq_his.pdf（月次）、sq_his(mini,weekly).pdf（週次）を取得して読む。未実現の満期は pending として記録。
使い方: python scripts/gamma_mle.py <出力ディレクトリ>
"""
from __future__ import annotations

import io
import json
import re
import sys
import warnings
from datetime import date, timedelta
from pathlib import Path

warnings.filterwarnings("ignore")
sys.path.insert(0, str(Path(__file__).resolve().parent)); sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import numpy as np
import pandas as pd
import pdfplumber
import requests

import rnd_lib
from data import iv_views
from data.qri_iv_loader import _R2_PREFIX_MINI, list_snapshot_keys, load_snapshot
from rnd_lib import density, fit_svi, prep, source_time, stats, tilt

OUT = Path(sys.argv[1]); OUT.mkdir(parents=True, exist_ok=True)
GRID = np.arange(-6, 12.01, 0.25); XS = np.linspace(-0.6, 0.6, 6001)
PDF_M = "https://www.jpx.co.jp/markets/derivatives/special-quotation/mklp7700000028jz-att/sq_his.pdf"
PDF_W = "https://www.jpx.co.jp/markets/derivatives/special-quotation/mklp7700000028jz-att/sq_his(mini,weekly).pdf"
UA = {"User-Agent": "Mozilla/5.0"}


def pdf_text(url: str) -> str:
    r = requests.get(url, headers=UA, timeout=60); r.raise_for_status()
    with pdfplumber.open(io.BytesIO(r.content)) as pdf:
        return "\n".join((p.extract_text() or "") for p in pdf.pages)


def monthly_sq() -> dict[str, float]:
    """'2026年1月 51,525.23 …' / '2月 57,045.65 …' → {'2601': 51525.23, ...}"""
    out, year = {}, None
    for l in pdf_text(PDF_M).split("\n"):
        m = re.match(r"^(?:(\d{4})年)?(\d{1,2})月\s+([\d,]+\.\d+)", l)
        if m:
            if m.group(1):
                year = int(m.group(1))
            if year:
                out[f"{year % 100:02d}{int(m.group(2)):02d}"] = float(m.group(3).replace(",", ""))
    return out


def weekly_sq() -> dict[str, float]:
    """'2026年10月16日限 64,665.18' → {'20261016': 64665.18}（SQ 算出日）"""
    out = {}
    for l in pdf_text(PDF_W).split("\n"):
        m = re.match(r"^(\d{4})年(\d{1,2})月(\d{1,2})日限\s+([\d,]+\.\d+)", l)
        if m:
            out[f"{int(m.group(1))}{int(m.group(2)):02d}{int(m.group(3)):02d}"] = float(m.group(4).replace(",", ""))
    return out


def loglik_grid(K, fK, sq: float) -> np.ndarray:
    return np.array([np.log(max(float(np.interp(sq, K, tilt(K, fK, g))), 1e-15)) for g in GRID])


def profile(ll_sum: np.ndarray, n: int) -> dict:
    i = int(np.argmax(ll_sum)); ci = GRID[ll_sum >= ll_sum[i] - 1.92]
    return dict(n=n, gamma_hat=float(GRID[i]), ci_lo=float(ci.min()), ci_hi=float(ci.max()), ll=ll_sum.tolist(),
                ll_at_0=float(ll_sum[int(np.argmin(np.abs(GRID)))]), llmax=float(ll_sum[i]))


# ---------------------------------------------------------------- 月次
SQM, SQW = monthly_sq(), weekly_sq()
print("monthly SQ:", {k: v for k, v in SQM.items() if k >= "2606"}); print("weekly SQ (latest 6):", dict(list(SQW.items())[-6:]))
from gamma_daily import LOCAL as STORE_LOCAL, R2_KEY as STORE_KEY  # noqa: E402
from data.r2_storage import r2_get, r2_put  # noqa: E402
_raw = r2_get(STORE_KEY) or (STORE_LOCAL.read_bytes() if STORE_LOCAL.exists() else b"")
ser = pd.DataFrame([json.loads(l) for l in _raw.decode("utf-8").splitlines() if l.strip()])  # 日次蓄積（gamma_daily.py）を月次の材料にする
ser["day"] = ser.day.astype(str); ser["cm"] = ser.cm.astype(str)
monthly, ll_m = [], np.zeros(len(GRID))
for cm in sorted(ser.cm.unique()):
    if cm not in SQM:
        continue
    c = ser[ser.cm == cm].copy(); c["dd"] = (c["T"] * 365 - 28).abs(); r = c.sort_values("dd").iloc[0]
    K, fK, _, _ = density(np.array(r.p), r.F, r["T"], XS); ll = loglik_grid(K, fK, SQM[cm]); ll_m += ll
    monthly.append(dict(cm=cm, day=r.day, T_days=float(r["T"] * 365), F=float(r.F), atm=float(r.atm), sq=SQM[cm], u=float(stats(K, fK)["cdf"](SQM[cm])), gamma_single=float(GRID[np.argmax(ll)])))
    print("M", cm, r.day, f"T={r['T']*365:.1f}d F={r.F:,.0f} SQ={SQM[cm]:,.0f} u={monthly[-1]['u']:.3f} single={GRID[np.argmax(ll)]}")

# ---------------------------------------------------------------- 週次（ミニ）
weekly, pending, ll_w = [], [], np.zeros(len(GRID))
days = [d for d in sorted(iv_views.list_available_days(60)) if d >= "20261020" and pd.Timestamp(d).weekday() < 5]
cands: dict[str, list] = {}
for d in days:
    keys = list_snapshot_keys(date(int(d[:4]), int(d[4:6]), int(d[6:8])), _R2_PREFIX_MINI)
    best = None
    for k in reversed(keys):
        st = source_time(k)
        if st is None or st > rnd_lib.DAY_CLOSE or st < "08:45:00":
            continue
        best = k; break
    if best is None:
        continue
    df = load_snapshot(best)
    if df is None or df.empty:
        continue
    two = df[df.ask_iv.notna() & df.bid_iv.notna()].groupby("contract_month").size()
    for code, n in two.items():
        if n >= 60:
            cands.setdefault(str(code), []).append((d, best, int(n)))
for code, lst in sorted(cands.items()):
    sq_day = (date(2000 + int(code[:2]), int(code[2:4]), int(code[4:6])) + timedelta(days=1)); sq_key = sq_day.strftime("%Y%m%d")
    rnd_lib.SQ_DATES[code] = pd.Timestamp(sq_day.strftime("%Y-%m-%d") + " 09:00")
    picked = None
    for d, k, n in lst:  # 最も早い予測日（残存 1.5 日以上）
        ts, F, T, pts, g = prep(k, code)
        if T * 365 >= 1.5 and np.isfinite(F) and len(pts) >= 25:
            picked = (d, k, n, ts, F, T, pts); break
    if picked is None:
        continue
    d, k, n, ts, F, T, pts = picked
    p, rmse = fit_svi(pts, T); K, fK, gg, _ = density(p, F, T, XS)
    rec = dict(code=code, sq_day=sq_key, forecast_day=d, ts=str(ts), T_days=float(T * 365), F=float(F), n_two_sided=n, rmse=float(rmse), gmin=float(gg.min()), p=p.tolist())
    if sq_key in SQW and rmse <= 0.03 and gg.min() > -0.05:
        ll = loglik_grid(K, fK, SQW[sq_key]); ll_w += ll
        rec.update(sq=SQW[sq_key], u=float(stats(K, fK)["cdf"](SQW[sq_key])), gamma_single=float(GRID[np.argmax(ll)])); weekly.append(rec)
        print("W", code, d, f"T={T*365:.1f}d F={F:,.0f} SQ={SQW[sq_key]:,.0f} u={rec['u']:.3f} single={rec['gamma_single']}")
    else:
        rec["status"] = "pending" if sq_key not in SQW else "poor_fit"; pending.append(rec)
        print("W", code, d, f"T={T*365:.1f}d F={F:,.0f} rmse={rmse*100:.2f}pt", rec["status"])
res = dict(grid=GRID.tolist(), monthly=monthly, weekly=weekly, pending=pending,
           profile_monthly=profile(ll_m, len(monthly)), profile_weekly=profile(ll_w, len(weekly)) if weekly else None,
           profile_all=profile(ll_m + ll_w, len(monthly) + len(weekly)))
body = json.dumps(res, ensure_ascii=False, default=float).encode("utf-8"); (OUT / "gamma_mle.json").write_bytes(body)
if "--upload" in sys.argv:
    r2_put("qri_iv/derived/gamma_mle.json", body)
pm, pa = res["profile_monthly"], res["profile_all"]
print(f"monthly: n={pm['n']} gamma_hat={pm['gamma_hat']} CI=[{pm['ci_lo']}, {pm['ci_hi']}] lnL(0)-max={pm['ll_at_0']-pm['llmax']:.2f}")
print(f"all: n={pa['n']} gamma_hat={pa['gamma_hat']} CI=[{pa['ci_lo']}, {pa['ci_hi']}]  pending weekly: {[(x['code'], x['status']) for x in pending]}")
