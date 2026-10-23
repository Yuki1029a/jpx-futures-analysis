"""成行で入った後の逆行（MAE）・指値の約定確率・追随エントリーの損切り先行確率・分散比を、日経 VI 水準別に定量化する。1 分足（2026-04-17〜05-14）。"""
import os, numpy as np, pandas as pd, matplotlib; matplotlib.use("Agg"); import matplotlib.pyplot as plt
from matplotlib import font_manager
for f in ("C:/Windows/Fonts/YuGothR.ttc", "C:/Windows/Fonts/meiryo.ttc"):
    if os.path.exists(f):
        font_manager.fontManager.addfont(f); plt.rcParams["font.family"] = font_manager.FontProperties(fname=f).get_name(); break
OUT = "H:/マイドライブ/bank/nk225op_sq_rnd_2611/"
bar = pd.read_parquet("C:/Users/kawai/AppData/Local/Temp/tick/bars_1min.parquet").sort_values(["sess_id", "min"])
vi = pd.read_csv("data/nikkei_vi_daily.csv", parse_dates=["date"]).set_index("date").close
def vi_for(date, sess):
    d = pd.Timestamp(date); return vi[vi.index < d].iloc[-1] if sess == "D" else (vi[d] if d in vi.index else vi[vi.index < d].iloc[-1])
N_LIST = (15, 30, 60); KS = (50, 100, 150, 200, 300); DELTA = 60  # 追随判定: 直前 5 分で 60 円以上動いた後に同方向へ成行
rows, chase, vr_rows = [], [], []
for sid, g in bar.groupby("sess_id"):
    date, sess = sid.split("_"); v = vi_for(date, sess); grid = np.arange(g["min"].min(), g["min"].max() + 1)
    c = g.set_index("min").c.reindex(grid).ffill().values; h = g.set_index("min").h.reindex(grid).fillna(pd.Series(c, index=grid)).values; l = g.set_index("min").l.reindex(grid).fillna(pd.Series(c, index=grid)).values
    n = len(c); r1 = np.diff(np.log(c))
    for q in (5, 15, 30):
        rq = np.log(c[q:]) - np.log(c[:-q]); vr_rows.append(dict(sess_id=sid, date=date, sess=sess, vi=v, q=q, vr=float(rq.var() / (q * r1.var())) if r1.var() > 0 else np.nan))
    for N in N_LIST:
        for t in range(5, n - N):
            fl, fh = l[t + 1:t + N + 1].min(), h[t + 1:t + N + 1].max(); e = c[t]
            rows.append(dict(sess_id=sid, date=date, sess=sess, vi=v, N=N, t=grid[t], px=e, mae_buy=e - fl, mae_sell=fh - e, mfe_buy=fh - e, mfe_sell=e - fl))
            if N == 30:
                m5 = c[t] - c[t - 5]
                for side, cond in (("buy", m5 >= DELTA), ("sell", m5 <= -DELTA), ("flat", abs(m5) < DELTA)):
                    if not cond:
                        continue
                    for k in (100, 200):
                        # 追随買い: 損切り e−k が利確 e+k より先に付くか（平坦時は買いで評価）
                        if side in ("buy", "flat"):
                            stop = np.where(l[t + 1:t + N + 1] <= e - k)[0]; targ = np.where(h[t + 1:t + N + 1] >= e + k)[0]
                        else:
                            stop = np.where(h[t + 1:t + N + 1] >= e + k)[0]; targ = np.where(l[t + 1:t + N + 1] <= e - k)[0]
                        s0 = stop[0] if len(stop) else 10 ** 9; t0 = targ[0] if len(targ) else 10 ** 9
                        chase.append(dict(sess_id=sid, vi=v, side=side, k=k, outcome="stop" if s0 < t0 else ("target" if t0 < s0 else "neither")))
df = pd.DataFrame(rows); ch = pd.DataFrame(chase); vr = pd.DataFrame(vr_rows)
df["mae"] = np.concatenate([df.mae_buy.values, df.mae_sell.values])[:len(df)] if False else (df.mae_buy + df.mae_sell) / 2  # 両方向平均（対称化）
pooled = pd.concat([df.assign(mae=df.mae_buy, mfe=df.mfe_buy, side="buy"), df.assign(mae=df.mae_sell, mfe=df.mfe_sell, side="sell")], ignore_index=True)
q_lo, q_hi = vi.loc[pd.Timestamp("2026-04-16"):pd.Timestamp("2026-05-14")].quantile([1 / 3, 2 / 3]).values
def tert(v):
    return "低 VI" if v <= q_lo else ("高 VI" if v > q_hi else "中 VI")
pooled["tier"] = pooled.vi.map(tert); ch["tier"] = ch.vi.map(tert); vr["tier"] = vr.vi.map(tert)
print(f"VI 三分位の境界: {q_lo:.1f} / {q_hi:.1f}（期間の VI 範囲 {pooled.vi.min():.1f}〜{pooled.vi.max():.1f}）  セッション数: " + str(pooled.groupby('tier').sess_id.nunique().to_dict()))
print("\n[1] 成行エントリー後 N 分以内の最大逆行（MAE、円）: 分位点と、k 円以上逆行する確率")
tab1 = []
for N in N_LIST:
    for tier in ("低 VI", "中 VI", "高 VI"):
        s = pooled[(pooled.N == N) & (pooled.tier == tier)]; vi_m = s.vi.mean(); px = s.px.mean()
        scale = px * vi_m / 100 * np.sqrt(N / (252 * 1170))  # VI 換算の N 分標準偏差（取引分ベース）
        tab1.append(dict(N=N, tier=tier, vi=round(vi_m, 1), n=len(s), med=s.mae.median(), q75=s.mae.quantile(0.75), q90=s.mae.quantile(0.9), p100=(s.mae >= 100).mean(), p200=(s.mae >= 200).mean(), p300=(s.mae >= 300).mean(), vi_sd=scale, ratio_med=s.mae.median() / scale))
t1 = pd.DataFrame(tab1); print(t1.round({"med": 0, "q75": 0, "q90": 0, "p100": 2, "p200": 2, "p300": 2, "vi_sd": 0, "ratio_med": 2}).to_string(index=False))
print("\n[2] k 円引いた指値の N 分以内約定確率（＝成行より k 円有利に入れた確率）と、指値にした場合の期待改善（円）")
tab2 = []
for N in (15, 30, 60):
    for tier in ("低 VI", "中 VI", "高 VI"):
        s = pooled[(pooled.N == N) & (pooled.tier == tier)]; rec = dict(N=N, tier=tier)
        for k in KS:
            rec[f"fill{k}"] = (s.mae >= k).mean()
        rec["E_improve"] = s.mae.mean(); tab2.append(rec)
t2 = pd.DataFrame(tab2); print(t2.round(2).to_string(index=False))
print("\n[3] 直前 5 分で 60 円以上動いた方向へ成行で追随（30 分、±k 円のどちらが先か）")
t3 = ch.groupby(["tier", "side", "k"]).outcome.value_counts(normalize=True).unstack().fillna(0); t3["n"] = ch.groupby(["tier", "side", "k"]).size(); print(t3.round(2).to_string())
print("\n[4] 分散比 VR(q) = Var(q 分リターン)/(q·Var(1 分))  <1 なら平均回帰（往復で削られる）")
t4 = vr.groupby(["tier", "q"]).vr.agg(["mean", "median", "count"]).round(2); print(t4.to_string())
import statsmodels.formula.api as smf
sm = pooled[pooled.N == 30].groupby("sess_id").agg(mae_med=("mae", "median"), vi=("vi", "first"), px=("px", "mean")).reset_index(); sm["lm"] = np.log(sm.mae_med / sm.px); sm["lvi"] = np.log(sm.vi)
m = smf.ols("lm ~ lvi", data=sm).fit(cov_type="HC1"); print(f"\n[5] セッション別 log(MAE 中央値/価格) ~ log VI: 弾性 {m.params['lvi']:+.2f} (t={m.tvalues['lvi']:+.1f}, R2={m.rsquared:.2f}, n={len(sm)})  ※VI に比例なら 1.0")
vr15 = vr[vr.q == 15].merge(sm[["sess_id", "mae_med"]], on="sess_id"); m2 = smf.ols("vr ~ lvi", data=vr15.assign(lvi=np.log(vr15.vi))).fit(cov_type="HC1"); print(f"    分散比 VR(15) ~ log VI: 係数 {m2.params['lvi']:+.2f} (t={m2.tvalues['lvi']:+.1f})")
pd.concat([t1.assign(table=1), t2.assign(table=2)], ignore_index=True).to_csv(OUT + "tick_entry_mae_vi.csv", index=False)
# ---- 図
fig, ax = plt.subplots(2, 2, figsize=(13, 8.5)); cols = {"低 VI": "#1d9e75", "中 VI": "#ba7517", "高 VI": "#993556"}
for tier in ("低 VI", "中 VI", "高 VI"):
    s = pooled[(pooled.N == 30) & (pooled.tier == tier)]; xs = np.arange(0, 601, 10); ax[0, 0].plot(xs, [(s.mae >= k).mean() * 100 for k in xs], color=cols[tier], lw=2, label=f"{tier}（VI {s.vi.mean():.0f}）")
ax[0, 0].set_title("成行で入って 30 分以内に k 円以上逆行する確率 ＝ k 円引いた指値が 30 分以内に約定する確率", loc="left", fontsize=10); ax[0, 0].set_xlabel("k（円）"); ax[0, 0].set_ylabel("%"); ax[0, 0].grid(alpha=0.3); ax[0, 0].legend(fontsize=8)
for tier in ("低 VI", "中 VI", "高 VI"):
    s = t1[t1.tier == tier]; ax[0, 1].plot(s.N, s.med, marker="o", color=cols[tier], lw=2, label=f"{tier} 中央値"); ax[0, 1].plot(s.N, s.q90, marker="s", ls="--", color=cols[tier], lw=1.2, label=f"{tier} 90% 点")
ax[0, 1].set_title("最大逆行 MAE（円）と保有時間", loc="left", fontsize=10); ax[0, 1].set_xlabel("N（分）"); ax[0, 1].set_ylabel("円"); ax[0, 1].grid(alpha=0.3); ax[0, 1].legend(fontsize=7, ncol=2)
cs = ch[ch.side != "flat"].groupby(["tier", "k"]).outcome.value_counts(normalize=True).unstack().fillna(0).reset_index(); cf = ch[ch.side == "flat"].groupby(["tier", "k"]).outcome.value_counts(normalize=True).unstack().fillna(0).reset_index()
w = 0.35; x = np.arange(3)
for j, k in enumerate((100, 200)):
    a = cs[cs.k == k].set_index("tier").reindex(["低 VI", "中 VI", "高 VI"]); b = cf[cf.k == k].set_index("tier").reindex(["低 VI", "中 VI", "高 VI"])
    ax[1, 0].bar(x + (j - 0.5) * w, a.stop * 100, w, color=["#1d9e75", "#ba7517", "#993556"], alpha=0.9 if k == 100 else 0.5, label=f"追随後 ±{k} 円: 損切りが先")
    ax[1, 0].plot(x + (j - 0.5) * w, b.stop * 100, "k_", ms=14, mew=2, label="（参考）追随なし" if j == 0 else None)
ax[1, 0].axhline(50, color="#888780", ls=":", lw=1); ax[1, 0].set_xticks(x); ax[1, 0].set_xticklabels(["低 VI", "中 VI", "高 VI"]); ax[1, 0].set_ylabel("%"); ax[1, 0].set_title("直前 5 分で 60 円以上動いた方向へ成行で追随 → 30 分以内に ±k 円のどちらが先に付くか（損切り先行の割合）", loc="left", fontsize=9); ax[1, 0].grid(alpha=0.3, axis="y"); ax[1, 0].legend(fontsize=8)
for tier in ("低 VI", "中 VI", "高 VI"):
    s = t4.loc[tier]; ax[1, 1].plot(s.index, s["mean"], marker="o", color=cols[tier], lw=2, label=tier)
ax[1, 1].axhline(1, color="#888780", ls=":", lw=1); ax[1, 1].set_title("分散比 VR(q)（1 なら ランダムウォーク、<1 なら平均回帰）", loc="left", fontsize=10); ax[1, 1].set_xlabel("q（分）"); ax[1, 1].grid(alpha=0.3); ax[1, 1].legend(fontsize=8)
fig.suptitle("日経225先物 1 分足（2026-04-17〜05-14、31 セッション）: VI 水準別の「成行で入るとやられる」度合い", x=0.01, ha="left", fontsize=12); fig.tight_layout(); fig.savefig(OUT + "tick_entry_mae_vi.png", dpi=150); print("saved")
