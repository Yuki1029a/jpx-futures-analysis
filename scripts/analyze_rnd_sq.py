"""11月限SQ値のリスク中立密度（SVI + Breeden-Litzenberger）と建玉の分析。出力: $TMP/rnd.json, rnd.npz。続けて analyze_rnd_sq_pdf.py で docs/rnd_sq_2611.pdf を生成。"""
import sys, os, json, warnings; warnings.filterwarnings("ignore"); sys.path.insert(0, ".")
import numpy as np, pandas as pd
from scipy.stats import norm
from scipy.optimize import minimize
from datetime import date
from data import iv_views
from data.qri_iv_loader import load_snapshot, prefetch_snapshots
from data.aggregator import _load_daily_oi_for_date
OUT = os.environ["TMP"]; CM = "2611"; SQ = pd.Timestamp("2026-11-13 09:00")


def snap_for(day, max_t="15:15:59"):
    ks = [k for k in iv_views.snapshot_keys(day) if iv_views.key_time_label(k) <= max_t]
    prefetch_snapshots(ks[-10:])
    for k in reversed(ks):
        g = iv_views._month_reduced(k, CM)
        if g is not None and int(g.q.sum()) >= 150:
            return k
    return None


def prep(k):
    m = iv_views.with_eff_iv(load_snapshot(k)); g = m[m.contract_month == CM].copy()
    for c in ["ask_price", "bid_price", "ask_iv", "bid_iv", "delta"]:
        g[c] = pd.to_numeric(g[c], errors="coerce")
    g["mid"] = ((g.ask_price + g.bid_price) / 2).where(g.ask_price.notna() & g.bid_price.notna())
    g["q"] = g.ask_iv.notna() & g.bid_iv.notna()
    pv = g.pivot_table(index="strike", columns="option_type", values="mid").dropna()
    par = pv.CALL - pv.PUT; near = par.abs().nsmallest(5).index; F = float(np.median(near + par[near]))
    d = k.split("/")[-2]; ts = pd.Timestamp(f"{d[:4]}-{d[4:6]}-{d[6:8]} {iv_views.key_time_label(k)}")
    T = (SQ - ts).total_seconds() / (365 * 86400)
    q = g[g.q].copy(); q["x"] = np.log(q.strike / F); q["spr"] = q.ask_iv - q.bid_iv
    otm = q[((q.option_type == "PUT") & (q.x < 0)) | ((q.option_type == "CALL") & (q.x >= 0))]
    pts = otm.groupby("strike").agg(x=("x", "first"), iv=("eff_iv", "mean"), spr=("spr", "mean")).reset_index()
    pts = pts[pts.iv.notna() & (pts.iv > 0.03)]
    return ts, F, T, pts, g


def svi(p, x):
    a, b, rho, m, s = p; u = x - m; R = np.sqrt(u * u + s * s)
    w = a + b * (rho * u + R); w1 = b * (rho + u / R); w2 = b * s * s / R ** 3
    return w, w1, w2


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
    rmse = float(np.sqrt(np.mean((np.sqrt(w / T) - pts.iv.values) ** 2)))
    return p, rmse


def density(p, F, T, xs):
    w, w1, w2 = svi(p, xs); sw = np.sqrt(w)
    g = (1 - xs * w1 / (2 * w)) ** 2 - (w1 ** 2 / 4) * (1 / w + 0.25) + w2 / 2
    dm = -xs / sw - sw / 2
    px = g / np.sqrt(2 * np.pi * w) * np.exp(-dm ** 2 / 2)
    K = F * np.exp(xs)
    return K, px / K, g, np.sqrt(w / T)


def stats(K, fK):
    dK = np.gradient(K); mass = fK * dK; Z = mass.sum(); mass = mass / Z
    cdf = np.cumsum(mass); mean = float((K * mass).sum()); sd = float(np.sqrt(((K - mean) ** 2 * mass).sum()))
    sk = float((((K - mean) / sd) ** 3 * mass).sum()); ku = float((((K - mean) / sd) ** 4 * mass).sum()) - 3
    qs = {q: float(np.interp(q, cdf, K)) for q in (0.01, 0.05, 0.10, 0.25, 0.50, 0.75, 0.90, 0.95, 0.99)}
    return dict(Z=float(Z), mean=mean, sd=sd, skew=sk, kurt=ku, mode=float(K[np.argmax(fK)]), q=qs,
                cdf=lambda k: float(np.interp(k, K, cdf)), P=lambda lo, hi: float(mass[(K >= lo) & (K < hi)].sum()))


def lognormal(F, T, s_atm, xs):
    w0 = s_atm ** 2 * T; px = 1 / np.sqrt(2 * np.pi * w0) * np.exp(-(xs + w0 / 2) ** 2 / (2 * w0))
    K = F * np.exp(xs); return K, px / K


def d1f(x, s, T):
    return (-x + 0.5 * s * s * T) / (s * np.sqrt(T))


def skew25(p, T, xs, ivf):
    sa = float(np.sqrt(svi(p, np.array([0.0]))[0][0] / T))
    xg = xs[np.abs(xs) < 0.3]; sg = ivf[np.abs(xs) < 0.3]; dc = norm.cdf(d1f(xg, sg, T))
    s25c = float(np.interp(0.25, dc[::-1], sg[::-1])); s25p = float(np.interp(-0.25, dc - 1, sg))
    return sa, s25p, s25c


# ---------------- latest ----------------
sys.path.insert(0, str(__import__("pathlib").Path(__file__).resolve().parent))
from rnd_lib import day_close_snap  # 気配時刻 15:45:59 以前の最終スナップショット（夜間の気配は使わない）
ASOF_DAY, OI_DATE = "20261019", date(2026, 10, 19)  # 基準時点: 日中引け、建玉は同日 20:00 公表分
k0 = day_close_snap(ASOF_DAY, CM)
ts, F, T, pts, g = prep(k0)
p, rmse = fit_svi(pts, T)
xs = np.linspace(-0.6, 0.6, 6001)
K, fK, gg, ivf = density(p, F, T, xs); st = stats(K, fK)
s_atm, s25p, s25c = skew25(p, T, xs, ivf)
KL, fL = lognormal(F, T, s_atm, xs); sL = stats(KL, fL)
print(f"snapshot {ts}  F={F:,.0f}  T={T*365:.1f}日  n_pts={len(pts)}  SVI={np.round(p,5).tolist()}  rmse_iv={rmse*100:.3f}pt  g_min={gg.min():.4f}  Z={st['Z']:.4f}")
print(f"ATM IV={s_atm*100:.2f}%  25dP={s25p*100:.2f} 25dC={s25c*100:.2f} RR25={(s25p-s25c)*100:+.2f}  BF25={((s25p+s25c)/2-s_atm)*100:+.2f}")
print(f"RND: mean={st['mean']:,.0f} (F={F:,.0f}) median={st['q'][0.5]:,.0f} mode={st['mode']:,.0f} sd={st['sd']:,.0f} skew={st['skew']:+.3f} exkurt={st['kurt']:+.3f}")
print("quantiles RND:", {k: round(v) for k, v in st['q'].items()})
print("quantiles LN :", {k: round(v) for k, v in sL['q'].items()})
for k in range(62000, 76001, 1000):
    print(f"P(S_T<{k:,}) RND={st['cdf'](k)*100:5.1f}%  LN={sL['cdf'](k)*100:5.1f}%   P({k:,}<=S<{k+1000:,}) RND={st['P'](k,k+1000)*100:4.1f}%")
# ---------------- OI (10/15) ----------------
oi = pd.DataFrame([vars(r) for r in _load_daily_oi_for_date(OI_DATE, CM)])
oi = oi[oi.current_oi > 0]
S = np.arange(60000, 78001, 125)
c, q = oi[oi.option_type == "CALL"], oi[oi.option_type == "PUT"]
pain = np.array([(c.current_oi * np.maximum(s - c.strike_price, 0)).sum() + (q.current_oi * np.maximum(q.strike_price - s, 0)).sum() for s in S])
maxpain = int(S[np.argmin(pain)])
print(f"OI 10/15: CALL={c.current_oi.sum():,} PUT={q.current_oi.sum():,} P/C={q.current_oi.sum()/c.current_oi.sum():.2f}  maxpain={maxpain:,}  OIweighted call={np.average(c.strike_price, weights=c.current_oi):,.0f} put={np.average(q.strike_price, weights=q.current_oi):,.0f}")
print("top CALL OI:", c.nlargest(6, "current_oi")[["strike_price", "current_oi", "net_change"]].values.tolist())
print("top PUT  OI:", q.nlargest(6, "current_oi")[["strike_price", "current_oi", "net_change"]].values.tolist())
print("top dOI 10/15:", oi.reindex(oi.net_change.abs().sort_values(ascending=False).index).head(8)[["option_type", "strike_price", "net_change", "current_oi"]].values.tolist())
print(f"P(|S_T-maxpain|<=500)={st['P'](maxpain-500, maxpain+500)*100:.1f}%  P(S_T<maxpain)={st['cdf'](maxpain)*100:.1f}%")


def gex(strike, oi_, typ):
    x = np.log(strike / F); s = float(np.sqrt(svi(p, np.array([x]))[0][0] / T)); gam = norm.pdf(d1f(x, s, T)) / (F * s * np.sqrt(T))
    return gam * oi_ * 1000 * F ** 2 * 0.01 * (1 if typ == "CALL" else -1)


oi["gex"] = [gex(r.strike_price, r.current_oi, r.option_type) for r in oi.itertuples()]
gx = oi.groupby("strike_price").gex.sum().sort_index()
print("GEX total(okuyen/1%)=", round(gx.sum() / 1e8, 1), " by strike top:", (gx.abs().nlargest(6) / 1e8).round(1).to_dict())
import yfinance as yf
h = yf.download("^N225", start="2026-07-01", end="2026-10-17", progress=False, auto_adjust=False); h.columns = [c_[0] if isinstance(c_, tuple) else c_ for c_ in h.columns]
lr = np.log(h.Close).diff().dropna(); rv20 = float(lr[-20:].std() * np.sqrt(250)); rv60 = float(lr[-60:].std() * np.sqrt(250))
print(f"RV20={rv20*100:.1f}% RV60={rv60*100:.1f}%  ATM IV={s_atm*100:.1f}%  ratio20={rv20/s_atm:.2f}")
lam = (rv20 / s_atm) ** 2; pP = p.copy(); pP[0] *= lam; pP[1] *= lam
KP, fP, gP, _ = density(pP, F, T, xs); sP = stats(KP, fP)
print("P-proxy quantiles:", {k: round(v) for k, v in sP['q'].items()}, "g_min", round(float(gP.min()), 4))
# ---------------- evolution ----------------
ev = []
for d in ["20261009", "20261013", "20261014", "20261015", "20261016", "20261019"]:
    k = day_close_snap(d, CM)
    if k is None:
        continue
    ts_, F_, T_, pts_, _ = prep(k); p_, rm_ = fit_svi(pts_, T_); K_, f_, g_, ivf_ = density(p_, F_, T_, xs); s_ = stats(K_, f_)
    sa, a25p, a25c = skew25(p_, T_, xs, ivf_)
    ev.append(dict(day=ts_.strftime("%m/%d %H:%M"), F=round(F_), atm=round(sa * 100, 2), rr25=round((a25p - a25c) * 100, 2), q05=round(s_['q'][0.05]), q25=round(s_['q'][0.25]), q50=round(s_['q'][0.5]), q75=round(s_['q'][0.75]), q95=round(s_['q'][0.95]), sd=round(s_['sd']), skew=round(s_['skew'], 3), p_gt70k=round((1 - s_['cdf'](70000)) * 100, 1), p_lt65k=round(s_['cdf'](65000) * 100, 1), rmse=round(rm_ * 100, 3), gmin=round(float(g_.min()), 4)))
ev = pd.DataFrame(ev); print(ev.to_string(index=False))
np.savez(OUT + "/rnd.npz", K=K, fK=fK, fL=fL, fP=fP, xs=xs, ivf=ivf, pts_x=pts.x.values, pts_iv=pts.iv.values, pts_K=pts.strike.values, S=S, pain=pain,
         oi_strike=oi.strike_price.values, oi_type=oi.option_type.values.astype(str), oi_cur=oi.current_oi.values, oi_chg=oi.net_change.values, gx_k=gx.index.values, gx_v=gx.values)
json.dump(dict(ts=str(ts), F=F, T=T, p=p.tolist(), rmse=rmse, gmin=float(gg.min()), s_atm=s_atm, s25p=s25p, s25c=s25c, st={k_: v for k_, v in st.items() if k_ in ("mean", "sd", "skew", "kurt", "mode")}, q=st['q'], qL=sL['q'], qP=sP['q'],
               maxpain=maxpain, rv20=rv20, rv60=rv60, lam=lam, cdf={k_: st['cdf'](k_) for k_ in range(60000, 78001, 500)}, cdfL={k_: sL['cdf'](k_) for k_ in range(60000, 78001, 500)}, cdfP={k_: sP['cdf'](k_) for k_ in range(60000, 78001, 500)},
               ev=ev.to_dict("records"), gex_total=float(gx.sum()), call_oi=int(c.current_oi.sum()), put_oi=int(q.current_oi.sum()), n_pts=int(len(pts))), open(OUT + "/rnd.json", "w"))
