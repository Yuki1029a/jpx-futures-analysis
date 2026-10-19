"""リスク中立密度（SVI + Breeden–Litzenberger）の共通関数。analyze_rnd_sq.py と同じ手順を関数化したもの。

時刻の扱い: QRI の配信は 15 分遅延で、スナップショットの source_update_time が気配の時刻。
日中立会は 08:45-15:45、夜間立会は 17:00-翌 06:00。「引け」のスナップショットは
source_update_time が 15:45:59 以前の最終（両気配が十分あるもの）とし、夜間の気配を混ぜない。
"""
from __future__ import annotations

import sys
from pathlib import Path

import numpy as np
import pandas as pd
from scipy.optimize import minimize
from scipy.stats import norm

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from data import iv_views  # noqa: E402
from data.qri_iv_loader import load_snapshot, prefetch_snapshots  # noqa: E402

SQ_DATES = {"2606": pd.Timestamp("2026-06-12 09:00"), "2607": pd.Timestamp("2026-07-10 09:00"), "2608": pd.Timestamp("2026-08-14 09:00"),
            "2609": pd.Timestamp("2026-09-11 09:00"), "2610": pd.Timestamp("2026-10-09 09:00"), "2611": pd.Timestamp("2026-11-13 09:00"),
            "2612": pd.Timestamp("2026-12-11 09:00"), "2701": pd.Timestamp("2027-01-08 09:00"), "2703": pd.Timestamp("2027-03-12 09:00")}
DAY_CLOSE, NIGHT_CLOSE = "15:45:59", "06:00:59"
PRICE_OK_FROM = "20260713"  # これより前の収集分は 4 桁以上の価格が下位桁欠落（IV・デルタは無事）


def source_time(k: str) -> str | None:
    """スナップショットの気配時刻（source_update_time）を 'HH:MM:SS' で返す。無ければ None。"""
    df = load_snapshot(k)
    if df is None or df.empty or "source_update_time" not in df.columns:
        return None
    s = df.source_update_time.dropna().astype(str)
    return s.iloc[0][11:19] if len(s) else None


def snap_for(day: str, max_t: str, cm: str, min_two_sided: int = 150, by_source: bool = True, min_t: str = "00:00:00") -> str | None:
    """その日の [min_t, max_t]（既定: 気配時刻 source_update_time で判定）で、両気配が min_two_sided 以上ある最終スナップショットのキー。"""
    ks = [k for k in iv_views.snapshot_keys(day) if iv_views.key_time_label(k) <= ("23:59:59" if by_source else max_t)]
    prefetch_snapshots(ks[-16:])
    for k in reversed(ks):
        st = source_time(k) if by_source else iv_views.key_time_label(k)
        if st is None or st > max_t or st < min_t:
            continue
        g = iv_views._month_reduced(k, cm)
        if g is not None and int(g.q.sum()) >= min_two_sided:
            return k
    return None


def day_close_snap(day: str, cm: str, min_two_sided: int = 150) -> str | None:
    """日中引け（気配時刻 08:45〜15:45:59）の最終スナップショット。夜間立会（17:00〜翌06:00）の気配は使わない。"""
    return snap_for(day, DAY_CLOSE, cm, min_two_sided, by_source=True, min_t="08:45:00")


def _forward_from_delta(g: pd.DataFrame, T: float) -> float:
    """コールの QRI デルタと IV から先物価格を逆算（価格欠落期の代替）: F = K·exp(σ√T·N⁻¹(Δ) − σ²T/2)。"""
    c = g[(g.option_type == "CALL") & g.delta.between(0.3, 0.7) & g.eff_iv.notna()]
    if len(c) < 3:
        return np.nan
    Fq = c.strike * np.exp(c.eff_iv * np.sqrt(T) * norm.ppf(c.delta) - 0.5 * c.eff_iv ** 2 * T)
    return float(np.median(Fq))


def prep(k: str, cm: str):
    """スナップショット → (気配時刻, 先物推定値F, 残存T, OTM側の両気配IV点, 限月の全行)。"""
    m = iv_views.with_eff_iv(load_snapshot(k)); g = m[m.contract_month == cm].copy()
    for c in ["ask_price", "bid_price", "ask_iv", "bid_iv", "delta"]:
        g[c] = pd.to_numeric(g[c], errors="coerce")
    g["mid"] = ((g.ask_price + g.bid_price) / 2).where(g.ask_price.notna() & g.bid_price.notna())
    g["q"] = g.ask_iv.notna() & g.bid_iv.notna()
    d = k.split("/")[-2]
    st = source_time(k) or iv_views.key_time_label(k)
    ts = pd.Timestamp(f"{d[:4]}-{d[4:6]}-{d[6:8]} {st}")
    if st[:2] <= "06" and iv_views.key_time_label(k)[:2] <= "06":
        pass  # 夜間明けのスナップショットは日付がファイルの日付と同じ
    T = (SQ_DATES[cm] - ts).total_seconds() / (365 * 86400)
    pv = g.pivot_table(index="strike", columns="option_type", values="mid").dropna()
    F = np.nan
    if d >= PRICE_OK_FROM and len(pv) >= 5:
        par = pv.CALL - pv.PUT; near = par.abs().nsmallest(5).index; F = float(np.median(near + par[near]))
    if not np.isfinite(F):
        F = _forward_from_delta(g, T)
    q = g[g.q].copy(); q["x"] = np.log(q.strike / F); q["spr"] = q.ask_iv - q.bid_iv
    otm = q[((q.option_type == "PUT") & (q.x < 0)) | ((q.option_type == "CALL") & (q.x >= 0))]
    pts = otm.groupby("strike").agg(x=("x", "first"), iv=("eff_iv", "mean"), spr=("spr", "mean")).reset_index()
    pts = pts[pts.iv.notna() & (pts.iv > 0.03)]
    return ts, F, T, pts, g


def svi(p, x):
    a, b, rho, m, s = p; u = x - m; R = np.sqrt(u * u + s * s)
    return a + b * (rho * u + R), b * (rho + u / R), b * s * s / R ** 3


def fit_svi(pts, T):
    x, wobs, wt = pts.x.values, (pts.iv.values ** 2) * T, 1.0 / (pts.spr.values + 0.005) ** 2

    def unpack(z):
        return np.array([z[0], np.exp(z[1]), np.tanh(z[2]), z[3], np.exp(z[4])])

    def loss(z):
        p = unpack(z); w, _, _ = svi(p, x)
        pen = 1e3 * max(0.0, -(p[0] + p[1] * p[4] * np.sqrt(1 - p[2] ** 2)))
        return float(np.sum(wt * (w - wobs) ** 2)) + pen
    w0 = float(np.interp(0, x, wobs)); best = None
    for rho0 in (-0.7, -0.3):
        for b0 in (0.05, 0.2):
            z0 = [w0 * 0.6, np.log(b0), np.arctanh(rho0), 0.0, np.log(0.1)]
            r = minimize(loss, z0, method="Nelder-Mead", options={"maxiter": 20000, "xatol": 1e-9, "fatol": 1e-12})
            if best is None or r.fun < best.fun:
                best = r
    p = unpack(best.x); w, _, _ = svi(p, x)
    return p, float(np.sqrt(np.mean((np.sqrt(w / T) - pts.iv.values) ** 2)))


def density(p, F, T, xs):
    w, w1, w2 = svi(p, xs); sw = np.sqrt(w)
    g = (1 - xs * w1 / (2 * w)) ** 2 - (w1 ** 2 / 4) * (1 / w + 0.25) + w2 / 2
    dm = -xs / sw - sw / 2; px = g / np.sqrt(2 * np.pi * w) * np.exp(-dm ** 2 / 2)
    K = F * np.exp(xs)
    return K, px / K, g, np.sqrt(w / T)


def stats(K, fK):
    dK = np.gradient(K); mass = fK * dK; Z = mass.sum(); mass = mass / Z; cdf = np.cumsum(mass)
    mean = float((K * mass).sum()); sd = float(np.sqrt(((K - mean) ** 2 * mass).sum()))
    sk = float((((K - mean) / sd) ** 3 * mass).sum()); ku = float((((K - mean) / sd) ** 4 * mass).sum()) - 3
    qs = {q: float(np.interp(q, cdf, K)) for q in (0.01, 0.05, 0.10, 0.25, 0.50, 0.75, 0.90, 0.95, 0.99)}
    return dict(Z=float(Z), mean=mean, sd=sd, skew=sk, kurt=ku, mode=float(K[np.argmax(fK)]), q=qs,
                cdf=lambda k: float(np.interp(k, K, cdf)), P=lambda lo, hi: float(mass[(K >= lo) & (K < hi)].sum()))


def skew25(p, T, xs, ivf):
    sa = float(np.sqrt(svi(p, np.array([0.0]))[0][0] / T))
    xg = xs[np.abs(xs) < 0.3]; sg = ivf[np.abs(xs) < 0.3]; dc = norm.cdf((-xg + 0.5 * sg * sg * T) / (sg * np.sqrt(T)))
    s25c = float(np.interp(0.25, dc[::-1], sg[::-1])); s25p = float(np.interp(-0.25, dc - 1, sg))
    return sa, s25p, s25c


def rr_delta(k: str, cm: str) -> float:
    """QRI のデルタから 25 デルタの IV を直接補間した RR25（%pt）。SVI を介さない確認用。"""
    m = iv_views.with_eff_iv(load_snapshot(k)); g = m[m.contract_month == cm].copy()
    g["delta"] = pd.to_numeric(g.delta, errors="coerce"); g = g[g.ask_iv.notna() & g.bid_iv.notna() & g.delta.notna()]
    c = g[g.option_type == "CALL"].sort_values("delta"); p = g[g.option_type == "PUT"].sort_values("delta")
    return float((np.interp(-0.25, p.delta, p.eff_iv) - np.interp(0.25, c.delta, c.eff_iv)) * 100)


def tilt(K, fK, gamma):
    """べき効用（相対的リスク回避度 gamma）でリスク中立密度を傾けた実確率密度: f_P ∝ f_Q · K^gamma。"""
    dK = np.gradient(K); m = fK * dK * K ** gamma; m = m / m.sum()
    return m / dK


def analyze_key(k: str, cm: str, gammas=(2.0, 4.0, 6.0)) -> dict:
    """1 スナップショットのリスク中立分布の要約（べき効用で傾けた分布の要約も含む）。"""
    ts, F, T, pts, g = prep(k, cm); p, rmse = fit_svi(pts, T)
    xs = np.linspace(-0.6, 0.6, 6001); K, fK, gg, ivf = density(p, F, T, xs); st = stats(K, fK)
    sa, s25p, s25c = skew25(p, T, xs, ivf)
    out = dict(key=k, ts=str(ts), F=F, T=T, n_pts=int(len(pts)), rmse=rmse, gmin=float(gg.min()), atm=sa, s25p=s25p, s25c=s25c,
               rr25_svi=(s25p - s25c) * 100, rr25_delta=rr_delta(k, cm), mean=st["mean"], sd=st["sd"], skew=st["skew"], kurt=st["kurt"], mode=st["mode"],
               q={str(q): v for q, v in st["q"].items()}, cdf={str(kv): st["cdf"](kv) for kv in range(50000, 80001, 500)}, p=p.tolist(), tilt={})
    for gm in gammas:
        stp = stats(K, tilt(K, fK, gm))
        out["tilt"][str(gm)] = dict(mean=stp["mean"], median=stp["q"][0.5], mode=stp["mode"], q05=stp["q"][0.05], q95=stp["q"][0.95], p_lt65=stp["cdf"](65000), p_ge70=1 - stp["cdf"](70000), p_lt60=stp["cdf"](60000))
    return out
