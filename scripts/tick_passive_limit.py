"""売り・買い別に「どれだけ引いて指値を置くべきか」: 30 分以内の約定確率が 50%/70% になる引き幅（MAE 分位点）を、VI 水準・VI 変化・前セッション実現ボラ・時間帯・直前の動き別に集計。"""
import os, numpy as np, pandas as pd, matplotlib; matplotlib.use("Agg"); import matplotlib.pyplot as plt
from matplotlib import font_manager
import statsmodels.formula.api as smf
for f in ("C:/Windows/Fonts/YuGothR.ttc", "C:/Windows/Fonts/meiryo.ttc"):
    if os.path.exists(f):
        font_manager.fontManager.addfont(f); plt.rcParams["font.family"] = font_manager.FontProperties(fname=f).get_name(); break
OUT = "H:/マイドライブ/bank/nk225op_sq_rnd_2611/"
bar = pd.read_parquet("C:/Users/kawai/AppData/Local/Temp/tick/bars_1min.parquet").sort_values(["sess_id", "min"])
vi = pd.read_csv("data/nikkei_vi_daily.csv", parse_dates=["date"]).set_index("date").close
N = 30
# 前セッションの実現ボラ（1 分足、年率）
rv = {}
for sid, g in bar.groupby("sess_id"):
    r1 = np.log(g.c).diff().dropna(); rv[sid] = float(np.sqrt((r1 ** 2).sum() / len(r1) * 252 * 1170) * 100) if len(r1) > 10 else np.nan
sess_order = sorted(bar.sess_id.unique(), key=lambda s: (s[:8], 0 if s.endswith("D") else 1)); prev_rv = {s: rv.get(sess_order[i - 1]) if i else np.nan for i, s in enumerate(sess_order)}
rows = []
for sid, g in bar.groupby("sess_id"):
    date, sess = sid.split("_"); d = pd.Timestamp(date); v_prev = vi[vi.index < d].iloc[-1]; v_same = vi[d] if d in vi.index else v_prev; v_prev2 = vi[vi.index < d].iloc[-2]
    v = v_prev if sess == "D" else v_same; dv = (v_prev - v_prev2) if sess == "D" else (v_same - v_prev)
    grid = np.arange(g["min"].min(), g["min"].max() + 1); c = g.set_index("min").c.reindex(grid).ffill().values
    h = g.set_index("min").h.reindex(grid).fillna(pd.Series(c, index=grid)).values; l = g.set_index("min").l.reindex(grid).fillna(pd.Series(c, index=grid)).values; n = len(c)
    for t in range(5, n - N):
        e = c[t]; fl, fh = l[t + 1:t + N + 1].min(), h[t + 1:t + N + 1].max(); m = int(grid[t]) % 1440
        tod = ("寄り 08:45-09:45" if m < 585 else ("引け 14:45-15:45" if m >= 885 else "日中 09:45-14:45")) if sess == "D" else ("夜 17-22" if 17 * 60 <= m < 22 * 60 else ("深夜 22-01" if (m >= 22 * 60 or m < 60) else "早朝 01-06"))
        rows.append(dict(sess_id=sid, sess=sess, vi=v, dvi=dv, rv_prev=prev_rv[sid], tod=tod, m5=c[t] - c[t - 5], e=e, mae_buy=e - fl, mae_sell=fh - e, run_buy=fh - e, run_sell=e - fl))
df = pd.DataFrame(rows); df = df.dropna(subset=["rv_prev", "dvi", "vi"]).reset_index(drop=True)
q_lo, q_hi = np.percentile(df.drop_duplicates("sess_id").vi, [33, 67]); df["vi_t"] = np.where(df.vi <= q_lo, "低 VI", np.where(df.vi > q_hi, "高 VI", "中 VI"))
r_lo, r_hi = np.percentile(df.drop_duplicates("sess_id").rv_prev, [33, 67]); df["rv_t"] = np.where(df.rv_prev <= r_lo, "低 RV", np.where(df.rv_prev > r_hi, "高 RV", "中 RV"))
df["dvi_t"] = np.where(df.dvi > 1.0, "VI 上昇(+1超)", np.where(df.dvi < -1.0, "VI 低下(−1超)", "VI 横ばい"))
df["mv"] = np.where(df.m5 >= 60, "直前 5 分 +60 円超", np.where(df.m5 <= -60, "直前 5 分 −60 円超", "直前 横ばい"))


def summ(g, side):
    mae = g[f"mae_{side}"]; run = g[f"run_{side}"]; nf = mae < 100
    return pd.Series({"n": len(g), "k50": mae.quantile(0.5), "k70": mae.quantile(0.3), "fill100": (mae >= 100).mean() * 100, "fill200": (mae >= 200).mean() * 100,
                      "逃げ幅(100円未約定時の順行中央値)": run[nf].median()})


print(f"30 分以内に約定する確率が 50% / 70% になる引き幅 k50 / k70（円）。VI 境界 {q_lo:.1f}/{q_hi:.1f}、前セッション RV 境界 {r_lo:.0f}/{r_hi:.0f}%")
tabs = {}
for key, lab in (("vi_t", "VI 水準"), ("dvi_t", "VI の前日比"), ("rv_t", "前セッション実現ボラ"), ("tod", "時間帯"), ("mv", "直前の動き")):
    parts = []
    for side, nm in (("buy", "買い（下に指値）"), ("sell", "売り（上に指値）")):
        t = df.groupby(key)[["mae_buy", "mae_sell", "run_buy", "run_sell"]].apply(lambda g: summ(g, side)); t.insert(0, "方向", nm); parts.append(t)
    tab = pd.concat(parts).reset_index().rename(columns={key: lab}); tabs[lab] = tab; print(f"\n== {lab} =="); print(tab.round(0).to_string(index=False))
# 回帰: log k（MAE+5）を状態変数と時間帯に回帰（売り・買い別）
df["lvi"] = np.log(df.vi); df["lrv"] = np.log(df.rv_prev)
d2 = df.replace([np.inf, -np.inf], np.nan).dropna(subset=["lvi", "lrv", "dvi", "mae_buy", "mae_sell"]).reset_index(drop=True)
for side in ("buy", "sell"):
    m = smf.ols(f"np.log(mae_{side} + 5) ~ lvi + lrv + dvi + C(tod) + C(mv)", data=d2).fit(cov_type="HC1")
    print(f"\n[{side}] log(MAE+5) ~ log VI + log RV_prev + ΔVI + 時間帯 + 直前の動き（HC1 SE、時系列相関で t は過大）: "
          f"log VI {m.params['lvi']:+.2f} (t={m.tvalues['lvi']:+.1f})  log RV_prev {m.params['lrv']:+.2f} (t={m.tvalues['lrv']:+.1f})  ΔVI {m.params['dvi']:+.3f}/pt (t={m.tvalues['dvi']:+.1f})  R2={m.rsquared:.2f}")
pd.concat([t.assign(表=k) for k, t in tabs.items()]).to_csv(OUT + "tick_passive_limit.csv", index=False)
# 図
fig, ax = plt.subplots(1, 3, figsize=(15, 4.8)); ks = np.arange(0, 401, 10)
for side, ls, nm in (("buy", "-", "買い"), ("sell", "--", "売り")):
    for tier, col in (("低 VI", "#1d9e75"), ("中 VI", "#ba7517"), ("高 VI", "#993556")):
        s = df[df.vi_t == tier]; ax[0].plot(ks, [(s[f"mae_{side}"] >= k).mean() * 100 for k in ks], ls=ls, color=col, lw=1.8, label=f"{nm}・{tier}")
ax[0].set_title("引き幅 k と 30 分以内の約定確率（VI 水準別、実線=買い・破線=売り）", loc="left", fontsize=9); ax[0].set_xlabel("k（円）"); ax[0].set_ylabel("%"); ax[0].grid(alpha=0.3); ax[0].legend(fontsize=7, ncol=2)
for i, (key, lab) in enumerate((("tod", "時間帯"), ("rv_t", "前セッション実現ボラ"))):
    t = tabs[lab]; cats = list(t[lab].unique()); x = np.arange(len(cats)); w = 0.38
    for j, (nm, col) in enumerate((("買い（下に指値）", "#2a78d6"), ("売り（上に指値）", "#eb6834"))):
        s = t[t["方向"] == nm].set_index(lab).reindex(cats); ax[i + 1].bar(x + (j - 0.5) * w, s.k70, w, color=col, alpha=0.85, label=f"{nm} k70"); ax[i + 1].plot(x + (j - 0.5) * w, s.k50, "k_", ms=12, mew=2)
    ax[i + 1].set_xticks(x); ax[i + 1].set_xticklabels(cats, fontsize=8, rotation=15); ax[i + 1].set_ylabel("円"); ax[i + 1].grid(alpha=0.3, axis="y"); ax[i + 1].set_title(f"{lab}別: 約定 70% の引き幅 k70（棒）と 50% の k50（横線）", loc="left", fontsize=9); ax[i + 1].legend(fontsize=7)
fig.suptitle("日経225先物 1 分足（2026-04-17〜05-14）: 売り・買い別 パッシブ指値の引き幅統計（30 分窓）", x=0.01, ha="left", fontsize=11); fig.tight_layout(); fig.savefig(OUT + "tick_passive_limit.png", dpi=150); print("saved")
