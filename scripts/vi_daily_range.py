"""日次・3 年半: 前日の日経 VI と当日の日経平均の値動き（日中レンジ、終値変化率）の関係。"""
import io, os, requests, numpy as np, pandas as pd, matplotlib; matplotlib.use("Agg"); import matplotlib.pyplot as plt
from matplotlib import font_manager
import statsmodels.formula.api as smf
for f in ("C:/Windows/Fonts/YuGothR.ttc", "C:/Windows/Fonts/meiryo.ttc"):
    if os.path.exists(f): font_manager.fontManager.addfont(f); plt.rcParams["font.family"] = font_manager.FontProperties(fname=f).get_name(); break
r = requests.get("https://indexes.nikkei.co.jp/nkave/historical/nikkei_stock_average_daily_jp.csv", timeout=60, headers={"User-Agent": "Mozilla/5.0"})
px = pd.read_csv(io.BytesIO(r.content), encoding="shift_jis"); px.columns = ["Date", "Close", "Open", "High", "Low"][:len(px.columns)]
px = px[px.Date.astype(str).str.match(r"\d{4}/\d{2}/\d{2}")].copy(); px["Date"] = pd.to_datetime(px.Date)
for c in ("Close", "Open", "High", "Low"): px[c] = pd.to_numeric(px[c], errors="coerce")
px = px.dropna().set_index("Date").sort_index()
px = px[px.index >= "2022-12-01"]; px.to_csv("data/nikkei225_daily_ohlc.csv")
vi = pd.read_csv("data/nikkei_vi_daily.csv", parse_dates=["date"]).set_index("date").sort_index()
d = px.join(vi.close.rename("vi"), how="inner"); d["vi_prev"] = d.vi.shift(1); d["dvi"] = d.vi - d.vi_prev
d["range"] = (d.High - d.Low) / d.Close.shift(1) * 100; d["absret"] = np.log(d.Close / d.Close.shift(1)).abs() * 100; d["gap"] = np.log(d.Open / d.Close.shift(1)).abs() * 100; d = d.dropna()
d["exp_range"] = 1.596 * d.vi_prev / np.sqrt(252)  # ブラウン運動の期待レンジ 1.596·σ√T（日次）
print(f"標本 {d.index.min().date()}〜{d.index.max().date()} n={len(d)}  VI 範囲 {d.vi_prev.min():.1f}〜{d.vi_prev.max():.1f}")
d["bin"] = pd.cut(d.vi_prev, [0, 16, 18, 20, 22, 25, 30, 40, 100], right=False)
t = d.groupby("bin", observed=True).agg(n=("range", "size"), vi=("vi_prev", "mean"), range_mean=("range", "mean"), range_med=("range", "median"), range_q90=("range", lambda s: s.quantile(0.9)), absret_mean=("absret", "mean"), exp_range=("exp_range", "mean"))
t["実現/VI換算"] = t.range_mean / t.exp_range; print("\n前日 VI 水準別: 当日の日中レンジ（%）・終値変化率の絶対値（%）"); print(t.round(2).to_string())
m = smf.ols("np.log(range) ~ np.log(vi_prev)", data=d).fit(cov_type="HAC", cov_kwds={"maxlags": 5}); print(f"\nlog レンジ ~ log 前日VI: 弾性 {m.params.iloc[1]:+.2f} (t={m.tvalues.iloc[1]:+.1f}) R2={m.rsquared:.2f}   ※VI 比例なら 1.0")
m2 = smf.ols("np.log(range) ~ np.log(vi_prev) + dvi", data=d).fit(cov_type="HAC", cov_kwds={"maxlags": 5}); print(f"+ 当日の VI 変化: log VI {m2.params.iloc[1]:+.2f} (t={m2.tvalues.iloc[1]:+.1f}), ΔVI {m2.params['dvi']:+.3f}/pt (t={m2.tvalues['dvi']:+.1f}) R2={m2.rsquared:.2f}")
print("\n同じ VI 水準でのばらつき: VI 20〜25 の日のレンジ 10%〜90% 点 =", np.round(d[(d.vi_prev >= 20) & (d.vi_prev < 25)].range.quantile([0.1, 0.5, 0.9]).values, 2))
up = d[d.dvi >= 2]; dn = d[d.dvi <= -2]; fl = d[d.dvi.abs() < 1]
print(f"当日 VI が +2pt 以上上昇した日のレンジ平均 {up.range.mean():.2f}%（n={len(up)}）, −2pt 以上低下 {dn.range.mean():.2f}%（n={len(dn)}）, ±1pt 未満 {fl.range.mean():.2f}%（n={len(fl)}）")
sub = d[(d.index >= "2026-04-17") & (d.index <= "2026-05-14")]; print(f"ティック期間 4/17〜5/14: VI 平均 {sub.vi_prev.mean():.1f}, レンジ平均 {sub.range.mean():.2f}% vs 全期間で同じ VI 帯（30〜40）の平均 {d[(d.vi_prev>=30)&(d.vi_prev<40)].range.mean():.2f}%")
fig, ax = plt.subplots(1, 2, figsize=(13, 5))
ax[0].scatter(d.vi_prev, d.range, s=8, alpha=0.35, color="#14285a"); ax[0].plot(t.vi, t.range_mean, "o-", color="#eb6834", lw=2, label="VI 帯ごとの平均レンジ"); xs = np.linspace(d.vi_prev.min(), d.vi_prev.max(), 50); ax[0].plot(xs, 1.596 * xs / np.sqrt(252), "--", color="#888780", label="VI 換算の期待レンジ 1.596·VI/√252")
ax[0].set_xlabel("前日の日経VI"); ax[0].set_ylabel("当日の日中レンジ（%）"); ax[0].set_title(f"日次 {d.index.min():%Y-%m}〜{d.index.max():%Y-%m}（n={len(d)}）: 前日 VI と当日レンジ", loc="left", fontsize=10); ax[0].grid(alpha=0.3); ax[0].legend(fontsize=8); ax[0].set_ylim(0, 8)
ax[1].scatter(d.dvi, d.range, s=8, alpha=0.35, color="#14285a"); ax[1].set_xlabel("当日の VI 変化（pt）"); ax[1].set_ylabel("当日の日中レンジ（%）"); ax[1].set_title("当日の VI 変化と当日レンジ（VI が跳ねる日に動く）", loc="left", fontsize=10); ax[1].grid(alpha=0.3); ax[1].set_ylim(0, 8); ax[1].set_xlim(-10, 15)
fig.tight_layout(); fig.savefig("H:/マイドライブ/bank/nk225op_sq_rnd_2611/vi_daily_range.png", dpi=150); print("saved")
