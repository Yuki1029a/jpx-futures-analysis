"""系統A（オプション価格から見通しを抽出する視点）の解説資料用の数値計算。結果を JSON/NPZ に保存する。

使い方: python scripts/outlook_a_calc.py <出力ディレクトリ>
入力: %TMP%/rnd.json, rnd.npz（10/16 引けの密度）、%TMP%/rnd_series_2611.json、QRI スナップショット（R2）、JPX 建玉、日経平均（yfinance）
"""
import io
import json
import os
import sys
import warnings
from datetime import date
from pathlib import Path

warnings.filterwarnings("ignore")
sys.path.insert(0, str(Path(__file__).resolve().parent)); sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import numpy as np
import pandas as pd
import yfinance as yf
from scipy.stats import norm

import rnd_lib
from data import iv_views
from data.qri_iv_loader import load_snapshot, prefetch_snapshots
from rnd_lib import analyze_key, snap_for, svi

OUT = Path(sys.argv[1]); OUT.mkdir(parents=True, exist_ok=True); TMP = Path(os.environ["TMP"])
rnd_lib.SQ_DATES.update({"2607": pd.Timestamp("2026-07-10 09:00"), "2608": pd.Timestamp("2026-08-14 09:00"), "2609": pd.Timestamp("2026-09-11 09:00"), "2610": pd.Timestamp("2026-10-09 09:00")})
SQD = rnd_lib.SQ_DATES
J = json.load(open(TMP / "rnd.json")); Z = np.load(TMP / "rnd.npz", allow_pickle=True)
K, fK, p16, F16, T16 = Z["K"], Z["fK"], np.array(J["p"]), J["F"], J["T"]
res = {}


def black(F, Kv, T, s, call=True):
    d1 = (np.log(F / Kv) + 0.5 * s * s * T) / (s * np.sqrt(T)); d2 = d1 - s * np.sqrt(T)
    return F * norm.cdf(d1) - Kv * norm.cdf(d2) if call else Kv * norm.cdf(-d2) - F * norm.cdf(-d1)


# ---------------------------------------------------------------- (1) 日次系列: ATM IV, モデルフリー IV, 実現ボラ
def front_cm(day):
    d = pd.Timestamp(day)
    for cm in sorted(SQD, key=lambda c: SQD[c]):
        if (SQD[cm] - d).days >= 8:
            return cm
    return None


def mfiv_from_quotes(g, F, T):
    """VIX 型のモデルフリー分散。OTM 両気配の IV からブラック価格を作り、離散ストライクで台形積分。"""
    q = g[g.ask_iv.notna() & g.bid_iv.notna() & g.eff_iv.notna()].copy()
    otm = q[((q.option_type == "PUT") & (q.strike < F)) | ((q.option_type == "CALL") & (q.strike >= F))].drop_duplicates("strike").sort_values("strike")
    if len(otm) < 10:
        return np.nan, 0
    Ks = otm.strike.values.astype(float); Qs = np.array([black(F, k, T, s, k >= F) for k, s in zip(Ks, otm.eff_iv.values)])
    dK = np.gradient(Ks); K0 = Ks[Ks <= F].max()
    var = (2 / T) * np.sum(dK / Ks ** 2 * Qs) - (1 / T) * (F / K0 - 1) ** 2
    return float(np.sqrt(max(var, 0))), int(len(Ks))


days = sorted(iv_views.list_available_days(200))
series = []
for d in days:
    if pd.Timestamp(d).weekday() >= 5:
        continue
    cm = front_cm(d)
    if cm is None:
        continue
    try:
        k = snap_for(d, "15:15:59", cm, 120)
    except Exception:
        k = None
    if k is None:
        continue
    try:
        ts, F, T, pts, g = rnd_lib.prep(k, cm)
    except Exception:
        continue
    pts = pts.sort_values("x"); atm = float(np.interp(0.0, pts.x.values, pts.iv.values))
    mf, nk = mfiv_from_quotes(g, F, T)
    series.append(dict(day=d, cm=cm, ts=str(ts), F=F, T=T, atm=atm, mfiv=mf, n=nk, kmin=float(pts.strike.min()), kmax=float(pts.strike.max())))
    print(d, cm, round(F), f"atm={atm*100:.1f} mfiv={mf*100:.1f} n={nk}", flush=True)
ser = pd.DataFrame(series)
h = yf.download("^N225", start="2000-01-01", end="2026-10-17", progress=False, auto_adjust=False); h.columns = [c[0] if isinstance(c, tuple) else c for c in h.columns]
lr = np.log(h.Close).diff().dropna()
rv20 = (lr.rolling(20).std() * np.sqrt(250)); rv60 = (lr.rolling(60).std() * np.sqrt(250))
fwd20 = (lr[::-1].rolling(20).std()[::-1].shift(-1) * np.sqrt(250))  # 翌日から 20 日の実現ボラ
ser["date"] = pd.to_datetime(ser.day); ser = ser.set_index("date")
ser["rv20"] = rv20.reindex(ser.index, method="ffill").values; ser["rv60"] = rv60.reindex(ser.index, method="ffill").values; ser["fwd20"] = fwd20.reindex(ser.index, method="ffill").values
ser["vrp_atm"] = ser.atm ** 2 - ser.rv20 ** 2; ser["vrp_mf"] = ser.mfiv ** 2 - ser.rv20 ** 2; ser["vrp_expost"] = ser.mfiv ** 2 - ser.fwd20 ** 2
ser.reset_index().to_json(OUT / "a3_series.json", orient="records", date_format="iso")
res["a3"] = dict(n=int(len(ser)), start=ser.day.iloc[0], end=ser.day.iloc[-1], mean_atm=float(ser.atm.mean()), mean_mfiv=float(ser.mfiv.mean()), mean_rv20=float(ser.rv20.mean()),
                 mean_vrp_mf=float(ser.vrp_mf.mean()), share_pos=float((ser.vrp_mf > 0).mean()), mean_expost=float(ser.vrp_expost.dropna().mean()), n_expost=int(ser.vrp_expost.notna().sum()),
                 last=ser.iloc[-1][["day", "cm", "F", "atm", "mfiv", "rv20", "rv60", "vrp_mf", "vrp_atm", "n", "kmin", "kmax"]].to_dict(),
                 corr_mf_atm=float(ser[["mfiv", "atm"]].corr().iloc[0, 1]))
# ---------------------------------------------------------------- (2) 実確率への変換（べき効用で傾ける）
dK = np.gradient(K); massQ = fK * dK; massQ = massQ / massQ.sum()


def summarize(mass):
    cdf = np.cumsum(mass); mean = float((K * mass).sum())
    return dict(mean=mean, median=float(np.interp(0.5, cdf, K)), mode=float(K[np.argmax(mass / dK)]), q05=float(np.interp(0.05, cdf, K)), q95=float(np.interp(0.95, cdf, K)),
                p_lt65=float(np.interp(65000, K, cdf)), p_ge70=float(1 - np.interp(70000, K, cdf)), p_lt_09F=float(np.interp(0.9 * F16, K, cdf)), p_lt60=float(np.interp(60000, K, cdf)))


tilt = {"0": summarize(massQ)}
for gam in (2, 4, 6):
    mP = massQ * K ** gam; mP = mP / mP.sum(); tilt[str(gam)] = summarize(mP)
res["a2"] = dict(F=F16, T=T16, tilt=tilt, lognormal_shift={str(g_): float(np.exp(g_ * J["s_atm"] ** 2 * T16)) for g_ in (2, 4, 6)}, s_atm=J["s_atm"])
np.savez(OUT / "a2_tilt.npz", K=K, fQ=massQ / dK, **{f"f{g_}": (massQ * K ** g_ / (massQ * K ** g_).sum()) / dK for g_ in (2, 4, 6)})
# ---------------------------------------------------------------- (3) BKM モデルフリー・モーメント（10/16 引け、SVI 価格）
xs = np.linspace(-0.8, 0.8, 8001); Kg = F16 * np.exp(xs); sig = np.sqrt(svi(p16, xs)[0] / T16)
C = np.array([black(F16, k, T16, s, True) for k, s in zip(Kg, sig)]); P = np.array([black(F16, k, T16, s, False) for k, s in zip(Kg, sig)])
up, dn = Kg >= F16, Kg < F16; lnu = np.log(Kg[up] / F16); lnd = np.log(F16 / Kg[dn]); dKg = np.gradient(Kg)
V = np.sum(2 * (1 - lnu) / Kg[up] ** 2 * C[up] * dKg[up]) + np.sum(2 * (1 + lnd) / Kg[dn] ** 2 * P[dn] * dKg[dn])
W = np.sum((6 * lnu - 3 * lnu ** 2) / Kg[up] ** 2 * C[up] * dKg[up]) - np.sum((6 * lnd + 3 * lnd ** 2) / Kg[dn] ** 2 * P[dn] * dKg[dn])
X = np.sum((12 * lnu ** 2 - 4 * lnu ** 3) / Kg[up] ** 2 * C[up] * dKg[up]) + np.sum((12 * lnd ** 2 + 4 * lnd ** 3) / Kg[dn] ** 2 * P[dn] * dKg[dn])
mu = -V / 2 - W / 6 - X / 24
skew_bkm = (W - 3 * mu * V + 2 * mu ** 3) / (V - mu ** 2) ** 1.5; kurt_bkm = (X - 4 * mu * W + 6 * mu ** 2 * V - 3 * mu ** 4) / (V - mu ** 2) ** 2
# 密度からの対数リターンのモーメント
x_ = np.log(K / F16); mx = float((x_ * massQ).sum()); sx = float(np.sqrt(((x_ - mx) ** 2 * massQ).sum()))
skew_x = float((((x_ - mx) / sx) ** 3 * massQ).sum()); kurt_x = float((((x_ - mx) / sx) ** 4 * massQ).sum())
res["a1"] = dict(V=float(V), W=float(W), X=float(X), mu=float(mu), vol_mf=float(np.sqrt(V / T16)), skew_bkm=float(skew_bkm), kurt_bkm=float(kurt_bkm), skew_x=skew_x, kurt_x=kurt_x,
                 skew_level=J["st"]["skew"], kurt_level=J["st"]["kurt"], atm=J["s_atm"], rr25=J["s25p"] - J["s25c"], bf25=(J["s25p"] + J["s25c"]) / 2 - J["s_atm"])
# ---------------------------------------------------------------- (4) 実確率の裾: 20 営業日リターンの歴史的頻度
r20 = np.log(h.Close).diff(20).dropna(); rv_at = rv20.reindex(r20.index)
hist = dict(n=int(len(r20)), p_lt_m10=float((r20 < np.log(0.9)).mean()), p_lt_m5=float((r20 < np.log(0.95)).mean()), p_gt_p5=float((r20 > np.log(1.05)).mean()), p_gt_p7=float((r20 > np.log(1.07)).mean()),
            hi=dict(n=int((rv_at >= 0.25).sum()), p_lt_m10=float((r20[rv_at >= 0.25] < np.log(0.9)).mean()), p_lt_m5=float((r20[rv_at >= 0.25] < np.log(0.95)).mean()), p_gt_p5=float((r20[rv_at >= 0.25] > np.log(1.05)).mean())),
            since="2000-01")
cdfQ = np.cumsum(massQ)
res["a5"] = dict(hist=hist, q=dict(p_lt_m10=float(np.interp(0.9 * F16, K, cdfQ)), p_lt_m5=float(np.interp(0.95 * F16, K, cdfQ)), p_gt_p5=float(1 - np.interp(1.05 * F16, K, cdfQ)), p_gt_p7=float(1 - np.interp(1.07 * F16, K, cdfQ))),
                 rv20_now=float(ser.rv20.iloc[-1]))
# ---------------------------------------------------------------- (6) 校正: 10月限の分布と実現 SQ 68,465.23
SQ_OCT = 68465.23; pit = []
for d in ["20260911", "20260918", "20260925", "20261001", "20261008"]:
    k = snap_for(d, "15:15:59", "2610", 120)
    if k is None:
        continue
    r = analyze_key(k, "2610"); Kr, fr = rnd_lib.density(np.array(r["p"]), r["F"], r["T"], np.linspace(-0.6, 0.6, 6001))[:2]
    m_ = fr * np.gradient(Kr); m_ = m_ / m_.sum(); u = float(np.interp(SQ_OCT, Kr, np.cumsum(m_)))
    pit.append(dict(day=d, ts=r["ts"], F=r["F"], T_days=r["T"] * 365, atm=r["atm"], q05=r["q"]["0.05"], q50=r["q"]["0.5"], q95=r["q"]["0.95"], u=u, z=float(norm.ppf(min(max(u, 1e-6), 1 - 1e-6)))))
    if d == "20260911":
        np.savez(OUT / "a6_pit.npz", K=Kr, f=fr)
    print("PIT", d, round(r["F"]), round(u, 3))
res["a6"] = dict(sq=SQ_OCT, rows=pit)
# ---------------------------------------------------------------- (7) 期間構造（10/16 引け）
k16 = snap_for("20261016", "15:15:59", "2611"); ts_rows = []
for cm in ["2611", "2612", "2703"]:
    try:
        r = analyze_key(k16, cm)
        ts_rows.append(dict(cm=cm, T=r["T"], atm=r["atm"], rr25=r["rr25_delta"], rr25_svi=r["rr25_svi"], F=r["F"], rmse=r["rmse"], skew=r["skew"], q05=r["q"]["0.05"], q95=r["q"]["0.95"]))
    except Exception as e:
        print("term", cm, e)
fw = []
for a, b in zip(ts_rows, ts_rows[1:]):
    fw.append(dict(pair=f"{a['cm']}→{b['cm']}", sig_f=float(np.sqrt((b["atm"] ** 2 * b["T"] - a["atm"] ** 2 * a["T"]) / (b["T"] - a["T"]))), days=(b["T"] - a["T"]) * 365))
res["a7"] = dict(rows=ts_rows, forward=fw)
# ---------------------------------------------------------------- (8) フロー代理指標（IV モニターの判定、10/16）
dd = [d for d in days if d <= "20261016"]
fl, meta = iv_views.flow_judgement(dd, "2611")
fl = fl[fl.d_oi.notna() & (fl.d_oi != 0)].copy()
cnt = fl.groupby(["option_type", "judge"]).size().unstack(fill_value=0)
newoi = fl[fl.d_oi > 0].groupby("option_type").d_oi.sum()
vw = fl.groupby("option_type").apply(lambda z: float(np.nansum(z.d_iv_ex_pct * z.d_oi.abs()) / np.nansum(z.d_oi.abs()))) if "d_iv_ex_pct" in fl.columns else None
res["a8"] = dict(meta={k_: (str(v) if not isinstance(v, (int, float)) else v) for k_, v in meta.items()}, counts=cnt.to_dict(), new_oi=newoi.to_dict(),
                 pc_new=float(newoi.get("PUT", 0) / max(newoi.get("CALL", 1), 1)), top=fl.reindex(fl.d_oi.abs().sort_values(ascending=False).index).head(10)[["option_type", "strike", "d_oi", "oi", "d_iv_pct", "d_iv_ex_pct", "judge"]].to_dict("records"),
                 oi_weighted_ex=vw.to_dict() if vw is not None else {})
# ---------------------------------------------------------------- (9) GEX（10/16 建玉）
gk, gv = Z["gx_k"], Z["gx_v"]; cum = np.cumsum(gv); flip = None
for i in range(1, len(gk)):
    if cum[i - 1] < 0 <= cum[i] or cum[i - 1] >= 0 > cum[i]:
        flip = float(gk[i]); break
res["a9"] = dict(total=float(gv.sum()), flip=flip, top=[(int(gk[i]), float(gv[i])) for i in np.argsort(-np.abs(gv))[:8]], F=F16, night_move=(63658 - 65720) / 65720)
np.savez(OUT / "a9_gex.npz", k=gk, v=gv)
json.dump(res, open(OUT / "a_calc.json", "w"), ensure_ascii=False, default=float)
print(json.dumps({k_: (v if k_ not in ("a3",) else {kk: vv for kk, vv in v.items() if kk != "last"}) for k_, v in res.items()}, ensure_ascii=False, default=float)[:3000])
