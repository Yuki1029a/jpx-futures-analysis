"""市場が織り込むリスク回避度 γ の推定（日次系列・最尤法・原資産との関係・精度）。

方法 I  （カーネル法）: ln[f_Q(x)/f_P(x)] = a − γ·x を中央部で回帰。f_P は歴史的 n 日リターン（開始時 RV20 で標準化）を現在の RV で再スケールした経験分布（KDE）。
方法 II （モーメント法）: γ = (Skew_P − Skew_Q) / (s_P (Kurt_P − 3))（Bakshi–Kapadia–Madan 2003 の近似）。
方法 III（最尤法）     : 実現 SQ 値の対数尤度 Σ ln f_P(SQ_i; γ) を γ について最大化（Bliss–Panigirtzoglou 2004）。
使い方: python scripts/gamma_calc.py <出力ディレクトリ>
"""
import io
import json
import os
import sys
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")
sys.path.insert(0, str(Path(__file__).resolve().parent)); sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import numpy as np
import pandas as pd
import yfinance as yf
from scipy.stats import gaussian_kde, norm

import rnd_lib
from data import iv_views
from rnd_lib import SQ_DATES, day_close_snap, density, fit_svi, prep, tilt

OUT = Path(sys.argv[1]); OUT.mkdir(parents=True, exist_ok=True)
SQ_REALIZED = {"2606": 66698.04, "2607": 69171.55, "2608": 69702.13, "2609": 63039.49, "2610": 68465.23}  # JPX sq_his.pdf
XS = np.linspace(-0.6, 0.6, 6001)

# ---------------------------------------------------------------- 実確率の材料: 日経平均の日次リターン
h = yf.download("^N225", start="1999-06-01", end="2026-10-21", progress=False, auto_adjust=False); h.columns = [c[0] if isinstance(c, tuple) else c for c in h.columns]
lp = np.log(h.Close.dropna()); r1 = lp.diff().dropna(); rv20 = r1.rolling(20).std() * np.sqrt(250); rv60 = r1.rolling(60).std() * np.sqrt(250)
fwd = {hh: (lp.shift(-hh) - lp) for hh in (5, 10, 20)}


def physical_sample(n_days: int, rv_now: float, start="2000-01-01"):
    """歴史的 n 日対数リターンを開始時点の RV20 で標準化し、現在の RV で再スケールした標本 x。"""
    rn = (lp.shift(-n_days) - lp); y = (rn / rv20).dropna(); y = y[y.index >= start]
    return (y * rv_now).values


def q_moments(K, fK, F):
    dK = np.gradient(K); m = fK * dK; m = m / m.sum(); x = np.log(K / F)
    mx = (x * m).sum(); s = np.sqrt(((x - mx) ** 2 * m).sum())
    return mx, s, (((x - mx) / s) ** 3 * m).sum(), (((x - mx) / s) ** 4 * m).sum()


def p_moments(xs):
    s = xs.std(); z = (xs - xs.mean()) / s
    return xs.mean(), s, (z ** 3).mean(), (z ** 4).mean()


def gamma_kernel(K, fK, F, xP, lo=0.05, hi=0.95, bw=None):
    """方法 I。ln(f_Q/f_P) を x に回帰。戻り値: γ, 切片, 2次項, 決定係数, 中央範囲."""
    dK = np.gradient(K); mQ = fK * dK; mQ = mQ / mQ.sum(); cdf = np.cumsum(mQ); x = np.log(K / F)
    kde = gaussian_kde(xP, bw_method=bw); fP = kde(x)  # x の密度
    fQx = mQ / np.gradient(x)
    sel = (cdf >= lo) & (cdf <= hi) & (fP > 1e-8) & (fQx > 1e-8)
    lnM = np.log(fQx[sel]) - np.log(fP[sel]); w = mQ[sel]; xx = x[sel]
    A = np.vstack([np.ones(sel.sum()), xx]).T; Wm = np.sqrt(w)[:, None]
    coef, *_ = np.linalg.lstsq(A * Wm, lnM * Wm[:, 0], rcond=None)
    pred = A @ coef; r2 = 1 - np.sum(w * (lnM - pred) ** 2) / np.sum(w * (lnM - np.average(lnM, weights=w)) ** 2)
    A2 = np.vstack([np.ones(sel.sum()), xx, xx ** 2]).T; coef2, *_ = np.linalg.lstsq(A2 * Wm, lnM * Wm[:, 0], rcond=None)
    return dict(gamma=float(-coef[1]), a=float(coef[0]), quad=float(coef2[2]), gamma_quadfit=float(-coef2[1]), r2=float(r2), xlo=float(xx.min()), xhi=float(xx.max()), lnM=lnM, x=xx, fP=fP, fQ=fQx, xgrid=x)


def front_cm(day):
    d = pd.Timestamp(day)
    for cm in sorted(SQ_DATES, key=lambda c: SQ_DATES[c]):
        if (SQ_DATES[cm] - d).days >= 8:
            return cm
    return None


# ---------------------------------------------------------------- 日次系列
days = [d for d in sorted(iv_views.list_available_days(200)) if pd.Timestamp(d).weekday() < 5]
rows = []; latest = None
for d in days:
    cm = front_cm(d)
    if cm is None:
        continue
    try:
        k = day_close_snap(d, cm, 120)
        if k is None:
            continue
        ts, F, T, pts, g = prep(k, cm)
        if not np.isfinite(F) or len(pts) < 30:
            continue
        p, rmse = fit_svi(pts, T); K, fK, gg, ivf = density(p, F, T, XS)
        if gg.min() < -0.05 or rmse > 0.03:
            print(d, "skip fit", round(rmse * 100, 2), round(float(gg.min()), 3)); continue
        n_days = max(int(round(T * 250)), 3); dt = pd.Timestamp(d)
        rv_now = float(rv20[rv20.index <= dt].iloc[-1]); rv60_now = float(rv60[rv60.index <= dt].iloc[-1])
        xP = physical_sample(n_days, rv_now); res = gamma_kernel(K, fK, F, xP)
        var = {"rv60": gamma_kernel(K, fK, F, physical_sample(n_days, rv60_now))["gamma"], "c80": gamma_kernel(K, fK, F, xP, 0.10, 0.90)["gamma"],
               "bw_narrow": gamma_kernel(K, fK, F, xP, bw=0.12)["gamma"]}
        mq, sq_, skq, kuq = q_moments(K, fK, F); mp, sp_, skp, kup = p_moments(xP)
        g2 = (skp - skq) / (sp_ * (kup - 3)) if kup > 3.2 else np.nan
        atm = float(np.sqrt(rnd_lib.svi(p, np.array([0.0]))[0][0] / T))
        rows.append(dict(day=d, cm=cm, ts=str(ts), F=F, T=T, n_days=n_days, atm=atm, rv20=rv_now, rv60=rv60_now, rmse=rmse, p=p.tolist(),
                         gamma_q=res["gamma_quadfit"], beta2=res["quad"], ivrv=atm - rv_now,  # 主指標
                         gamma_I=res["gamma"], gamma_I_quad=res["gamma_quadfit"], quad=res["quad"], r2=res["r2"],
                         gamma_rv60=var["rv60"], gamma_c80=var["c80"], gamma_bw=var["bw_narrow"], gamma_II=float(g2), skew_q=float(skq), kurt_q=float(kuq), skew_p=float(skp), kurt_p=float(kup), s_q=float(sq_), s_p=float(sp_),
                         ret20=float(lp[lp.index <= dt].iloc[-1] - lp[lp.index <= dt].iloc[-21]) if (lp.index <= dt).sum() > 21 else np.nan))
        latest = dict(day=d, K=K, fQ=res["fQ"], fP=res["fP"], x=res["xgrid"], lnM=res["lnM"], xsel=res["x"], a=res["a"], gamma=res["gamma"], gamma_q=res["gamma_quadfit"], beta2=res["quad"], F=F, T=T, fK=fK, xP=xP)
        print(d, cm, round(F), f"T={T*365:.0f}d atm={atm*100:.1f} rv20={rv_now*100:.1f} gI={res['gamma']:.2f} (rv60 {var['rv60']:.2f}, c80 {var['c80']:.2f}) gII={g2:.2f} r2={res['r2']:.2f} quad={res['quad']:.1f}", flush=True)
    except Exception as e:
        print(d, "error", e)
ser = pd.DataFrame(rows); ser.to_json(OUT / "gamma_series.json", orient="records")
# 原資産との関係
ser["date"] = pd.to_datetime(ser.day); ser = ser.set_index("date")
for hh in (5, 10, 20):
    ser[f"fwd{hh}"] = fwd[hh].reindex(ser.index, method="ffill").values
def summ(col):
    """系列の要約と、原資産・ボラ・先行リターンとの相関。"""
    s = ser[col]
    return dict(mean=float(s.mean()), sd=float(s.std()), median=float(s.median()), q10=float(s.quantile(0.1)), q90=float(s.quantile(0.9)), autocorr1=float(s.autocorr(1)),
                corr=dict(lnF=float(np.corrcoef(s, np.log(ser.F))[0, 1]), ret20=float(ser[[col, "ret20"]].corr().iloc[0, 1]), rv20=float(np.corrcoef(s, ser.rv20)[0, 1]),
                          atm=float(np.corrcoef(s, ser.atm)[0, 1]), ivrv=float(np.corrcoef(s, ser.ivrv)[0, 1]), gII=float(ser[[col, "gamma_II"]].corr().iloc[0, 1])),
                fwd_corr={str(hh): float(ser[[col, f"fwd{hh}"]].dropna().corr().iloc[0, 1]) for hh in (5, 10, 20)})


# 主指標 = gamma_q（2 次式の傾き）と beta2（曲率）。直線の gamma_I は参考（rel["lin"]）。
rel = dict(n=int(len(ser)), primary="gamma_q", **summ("gamma_q"), b2=summ("beta2"), lin=summ("gamma_I"), mean_II=float(ser.gamma_II.mean()),
           fwd_n={str(hh): int(ser[f"fwd{hh}"].notna().sum()) for hh in (5, 10, 20)},
           variant_spread=float((ser[["gamma_I", "gamma_rv60", "gamma_c80", "gamma_bw"]].max(axis=1) - ser[["gamma_I", "gamma_rv60", "gamma_c80", "gamma_bw"]].min(axis=1)).mean()),
           corr_q_lin=float(ser[["gamma_q", "gamma_I"]].corr().iloc[0, 1]))
X = np.vstack([np.ones(len(ser)), ser.ret20.fillna(0).values, ser.rv20.values]).T; b, *_ = np.linalg.lstsq(X, ser.gamma_q.values, rcond=None)
rel["reg"] = dict(const=float(b[0]), ret20=float(b[1]), rv20=float(b[2]))
Xb = np.vstack([np.ones(len(ser)), ser.ivrv.values * 100]).T; bb, *_ = np.linalg.lstsq(Xb, ser.beta2.values, rcond=None)
rel["reg_b2"] = dict(const=float(bb[0]), ivrv=float(bb[1]))
# ---------------------------------------------------------------- ブートストラップ（最終日）
rng = np.random.default_rng(0); nb = 200; L = latest; n_days = int(round(L["T"] * 250)); xP = L["xP"]; blk = max(n_days, 5)
bs = {"gamma_q": [], "beta2": [], "gamma_I": []}
nblocks = len(xP) // blk
for _ in range(nb):
    idx = rng.integers(0, len(xP) - blk, size=nblocks); samp = np.concatenate([xP[i:i + blk] for i in idx])
    rb = gamma_kernel(L["K"], L["fK"], L["F"], samp); bs["gamma_q"].append(rb["gamma_quadfit"]); bs["beta2"].append(rb["quad"]); bs["gamma_I"].append(rb["gamma"])
_bsum = lambda v: dict(mean=float(np.mean(v)), sd=float(np.std(v)), lo=float(np.percentile(v, 2.5)), hi=float(np.percentile(v, 97.5)))
boot = dict(**_bsum(bs["gamma_q"]), b2=_bsum(bs["beta2"]), lin=_bsum(bs["gamma_I"]), n=nb, block=blk)
# ---------------------------------------------------------------- 方法 III: 最尤法（実現 SQ）
grid = np.arange(0, 12.01, 0.25); mle = {"expiries": [], "profile_28d": None, "profile_all": None}
ll28 = np.zeros_like(grid); llall = np.zeros_like(grid); n_all = 0
for cm, sq in SQ_REALIZED.items():
    cands = [r for r in rows if r["cm"] == cm]
    if not cands:
        continue
    # 各日の密度を再計算（保存していないため）
    per_day = []
    for r in cands:
        F, T = r["F"], r["T"]; K, fK, gg, ivf = density(np.array(r["p"]), F, T, XS)
        dK = np.gradient(K); vals = []
        for gm in grid:
            fP = tilt(K, fK, gm); vals.append(np.log(max(float(np.interp(sq, K, fP)), 1e-12)))
        per_day.append(dict(day=r["day"], T=T, F=F, ll=np.array(vals), u0=float(np.interp(sq, K, np.cumsum(fK * dK) / (fK * dK).sum()))))
    if not per_day:
        continue
    best = min(per_day, key=lambda z: abs(z["T"] * 365 - 28)); ll28 += best["ll"]
    avg = np.mean([z["ll"] for z in per_day], axis=0); llall += avg; n_all += 1
    mle["expiries"].append(dict(cm=cm, sq=sq, day28=best["day"], T28=best["T"] * 365, F28=best["F"], u0=best["u0"], gamma_hat_single=float(grid[np.argmax(best["ll"])]), n_days=len(per_day)))
    print("MLE", cm, best["day"], round(best["T"] * 365), round(best["F"]), "u0", round(best["u0"], 3), "single-gamma", grid[np.argmax(best["ll"])])
i28 = int(np.argmax(ll28)); ci = grid[ll28 >= ll28[i28] - 1.92]
mle["profile_28d"] = dict(grid=grid.tolist(), ll=ll28.tolist(), gamma_hat=float(grid[i28]), ci_lo=float(ci.min()), ci_hi=float(ci.max()), n=len(mle["expiries"]))
ia = int(np.argmax(llall)); mle["profile_all"] = dict(grid=grid.tolist(), ll=llall.tolist(), gamma_hat=float(grid[ia]), n=n_all)
json.dump(dict(series_summary=rel, boot=boot, mle=mle, latest=dict(day=L["day"], gamma=L["gamma"], gamma_q=L["gamma_q"], beta2=L["beta2"], a=L["a"], F=L["F"], T=L["T"]), sq_realized=SQ_REALIZED), open(OUT / "gamma_calc.json", "w"), default=float)
np.savez(OUT / "gamma_latest.npz", K=L["K"], fQ=L["fQ"], fP=L["fP"], x=L["x"], lnM=L["lnM"], xsel=L["xsel"], fK=L["fK"])
print("summary", json.dumps(rel, default=float)); print("boot", boot); print("mle28", mle["profile_28d"]["gamma_hat"], mle["profile_28d"]["ci_lo"], mle["profile_28d"]["ci_hi"], "mle_all", mle["profile_all"]["gamma_hat"])
