"""(A) 時間帯別セッション値幅確定確率、(B) 日経VI × 時間当たり出来高 × 時間当たり先物変動。ティック 1 分足（2026-04-17〜05-14）。"""
import os, io, requests, numpy as np, pandas as pd, matplotlib; matplotlib.use("Agg"); import matplotlib.pyplot as plt
from matplotlib import font_manager
import statsmodels.formula.api as smf
for f in ("C:/Windows/Fonts/YuGothR.ttc", "C:/Windows/Fonts/meiryo.ttc"):
    if os.path.exists(f):
        font_manager.fontManager.addfont(f); plt.rcParams["font.family"] = font_manager.FontProperties(fname=f).get_name(); break
OUT = "H:/マイドライブ/bank/nk225op_sq_rnd_2611/"
bar = pd.read_parquet("C:/Users/kawai/AppData/Local/Temp/tick/bars_1min.parquet")
# ---- 日経VI（日経平均プロフィル、日次）
r = requests.get("https://indexes.nikkei.co.jp/nkave/historical/nikkei_stock_average_vi_daily_jp.csv", headers={"User-Agent": "Mozilla/5.0"}, timeout=60)
vi = pd.read_csv(io.BytesIO(r.content), encoding="shift_jis"); vi.columns = ["date", "close", "open", "high", "low"]
vi = vi[vi.date.astype(str).str.match(r"\d{4}/\d{2}/\d{2}")].copy(); vi["date"] = pd.to_datetime(vi.date)
for c in ("close", "open", "high", "low"):
    vi[c] = pd.to_numeric(vi[c], errors="coerce")
vi = vi.dropna(subset=["close"]).sort_values("date"); vi.to_csv("data/nikkei_vi_daily.csv", index=False)
print("VI:", vi.date.min().date(), "..", vi.date.max().date(), len(vi))


def lbl(m):
    m = int(m) % 1440
    return f"{m // 60:02d}:{m % 60:02d}"


# ================= (A) 値幅確定確率
def range_fix(sess):
    res = []
    for sid, g in bar[bar.session == sess].groupby("sess_id"):
        g = g.sort_values("min"); grid = np.arange(g["min"].min(), g["min"].max() + 1)
        h = g.set_index("min").h.reindex(grid).ffill(); l = g.set_index("min").l.reindex(grid).ffill()
        rh, rl = h.cummax(), l.cummin(); H, L = h.max(), l.min()
        res.append(pd.DataFrame({"sess_id": sid, "min": grid, "high_set": (rh >= H).values, "low_set": (rl <= L).values,
                                 "frac": ((rh - rl) / (H - L)).values, "t_high": int(h.idxmax()), "t_low": int(l.idxmin()), "range": H - L}))
    return pd.concat(res)


figA, axA = plt.subplots(2, 2, figsize=(13, 8.5)); tabA = {}
for i, (sess, name) in enumerate((("D", "日中（08:45〜15:45）"), ("E", "ナイト（17:00〜翌06:00）"))):
    rf = range_fix(sess); n_s = rf.sess_id.nunique(); grp = rf.groupby("min")
    prob = pd.DataFrame({"high": grp.high_set.mean(), "low": grp.low_set.mean(), "both": rf.assign(b=rf.high_set & rf.low_set).groupby("min").b.mean(),
                         "frac_med": grp.frac.median(), "frac_q25": grp.frac.quantile(0.25), "frac_q75": grp.frac.quantile(0.75)})
    ax = axA[i, 0]; x = prob.index.values
    ax.plot(x, prob.high * 100, color="#2a78d6", lw=1.8, label="高値が確定済み"); ax.plot(x, prob.low * 100, color="#eb6834", lw=1.8, label="安値が確定済み")
    ax.plot(x, prob.both * 100, color="#0b0b0b", lw=2.4, label="値幅（高値・安値とも）確定済み")
    ax.fill_between(x, prob.frac_q25 * 100, prob.frac_q75 * 100, color="#1baf7a", alpha=0.15)
    ax.plot(x, prob.frac_med * 100, color="#1baf7a", lw=1.4, ls="--", label="最終値幅に対する既達成率（中央値、帯は四分位）")
    ticks = x[::60] if sess == "D" else x[::120]; ax.set_xticks(ticks); ax.set_xticklabels([lbl(t) for t in ticks], fontsize=8)
    ax.set_ylim(0, 100); ax.set_ylabel("%"); ax.grid(alpha=0.3); ax.set_title(f"{name}  セッション値幅の確定確率（n = {n_s} セッション）", loc="left", fontsize=10); ax.legend(fontsize=8, loc="lower right")
    th = rf.groupby("sess_id").t_high.first(); tl = rf.groupby("sess_id").t_low.first(); bins = np.arange(x.min(), x.max() + 31, 30)
    ax2 = axA[i, 1]; ax2.hist([th, tl], bins=bins, label=["高値の時刻", "安値の時刻"], color=["#2a78d6", "#eb6834"], alpha=0.8)
    tk = bins[::2] if sess == "D" else bins[::4]; ax2.set_xticks(tk); ax2.set_xticklabels([lbl(t) for t in tk], fontsize=8)
    ax2.set_ylabel("セッション数"); ax2.grid(alpha=0.3); ax2.legend(fontsize=8); ax2.set_title("セッション高値・安値が付いた時刻（30 分刻み）", loc="left", fontsize=10)
    chk = [x.min() + 30, x.min() + 60, x.min() + 120, x.min() + 180, x.min() + 240, x.max() - 60, x.max() - 30] if sess == "D" else [x.min() + 60, x.min() + 120, x.min() + 240, x.min() + 360, x.min() + 480, x.min() + 600, x.max() - 60]
    tabA[sess] = pd.DataFrame({"時刻": [lbl(t) for t in chk], "高値確定%": [prob.high.get(t, np.nan) * 100 for t in chk], "安値確定%": [prob.low.get(t, np.nan) * 100 for t in chk],
                               "値幅確定%": [prob.both.get(t, np.nan) * 100 for t in chk], "既達成率中央値%": [prob.frac_med.get(t, np.nan) * 100 for t in chk]})
    print(f"\n[{name}] n={n_s}  平均値幅 {rf.groupby('sess_id').range.first().mean():.0f} 円"); print(tabA[sess].round(0).to_string(index=False))
figA.suptitle("日経225先物ラージ 1 分足（2026-04-17〜05-14）: 時間帯別セッション値幅確定確率", x=0.01, ha="left", fontsize=12); figA.tight_layout()
figA.savefig(OUT + "tick_session_range_fix.png", dpi=150); print("saved A")

# ================= (B) VI × 時間当たり出来高 × 時間当たり変動
bar = bar.sort_values(["sess_id", "min"]).reset_index(drop=True); bar["ret"] = bar.groupby("sess_id").c.transform(lambda s: np.log(s).diff())
bar["hour"] = np.where(bar.session == "D", ((bar["min"] - 525) // 60).clip(0, 6), ((bar["min"] - 1020) // 60).clip(0, 12)).astype(int)
hb = bar.groupby(["sess_id", "date", "session", "hour"]).agg(v=("v", "sum"), o=("o", "first"), c=("c", "last"), h=("h", "max"), l=("l", "min"),
                                                            rv=("ret", lambda s: float(np.nansum(s ** 2))), nmin=("min", "size")).reset_index()
hb["absret"] = (np.log(hb.c / hb.o)).abs() * 100; hb["rng"] = (hb.h - hb.l) / hb.o * 100
hb["rv_ann"] = np.sqrt(hb.rv * (60 / hb.nmin.clip(lower=1)) * 19.5 * 252) * 100  # 1 分 RV → 年率（取引時間 19.5h/日）
hb["dt"] = pd.to_datetime(hb.date); vis = vi.set_index("date").close
hb["vi"] = [vis[vis.index < d].iloc[-1] if s == "D" else (vis[d] if d in vis.index else vis[vis.index < d].iloc[-1]) for d, s in zip(hb.dt, hb.session)]
hb["hour_lbl"] = [lbl(525 + h * 60) if s == "D" else lbl(1020 + h * 60) for h, s in zip(hb.hour, hb.session)]
hb.to_csv(OUT + "tick_hourly_stats.csv", index=False)
hb["lv"] = np.log(hb.v.clip(lower=1)); hb["lrv"] = np.log(hb.rv.clip(lower=1e-12)); hb["lvi"] = np.log(hb.vi); hb["sh"] = hb.session + hb.hour.astype(str)
m1 = smf.ols("lrv ~ lvi + C(sh)", data=hb).fit(cov_type="HC1"); m2 = smf.ols("lv ~ lvi + C(sh)", data=hb).fit(cov_type="HC1")
m3 = smf.ols("lrv ~ lv + C(sh)", data=hb).fit(cov_type="HC1"); m4 = smf.ols("lrv ~ lv + lvi + C(sh)", data=hb).fit(cov_type="HC1")
print(f"\n[B] 時間足 n={len(hb)}（日中 {int((hb.session == 'D').sum())}、ナイト {int((hb.session == 'E').sum())}）、VI 範囲 {hb.vi.min():.1f}〜{hb.vi.max():.1f}")
print(f"  log RV ~ log VI + 時間帯:  弾性 {m1.params['lvi']:+.2f} (t={m1.tvalues['lvi']:+.1f}, R2={m1.rsquared:.2f})  ※VI^2 比例なら 2.0")
print(f"  log 出来高 ~ log VI + 時間帯: 弾性 {m2.params['lvi']:+.2f} (t={m2.tvalues['lvi']:+.1f}, R2={m2.rsquared:.2f})")
print(f"  log RV ~ log 出来高 + 時間帯: 弾性 {m3.params['lv']:+.2f} (t={m3.tvalues['lv']:+.1f}, R2={m3.rsquared:.2f})")
print(f"  log RV ~ log 出来高 + log VI + 時間帯: 出来高 {m4.params['lv']:+.2f} (t={m4.tvalues['lv']:+.1f}), VI {m4.params['lvi']:+.2f} (t={m4.tvalues['lvi']:+.1f}), R2={m4.rsquared:.2f}")
prof = hb.groupby(["session", "hour", "hour_lbl"]).agg(v=("v", "mean"), rv_ann=("rv_ann", "mean"), absret=("absret", "mean"), rng=("rng", "mean")).reset_index()
print("\n時間帯プロファイル（平均）:"); print(prof.round(2).to_string(index=False))
day = hb.groupby(["date", "session"]).agg(rv=("rv", "sum"), v=("v", "sum"), vi=("vi", "first")).reset_index()
dd = day.pivot(index="date", columns="session", values="rv").fillna(0); dd["rv_day"] = dd.sum(axis=1); dd["vi"] = day[day.session == "D"].set_index("date").vi; dd["rv_ann"] = np.sqrt(dd.rv_day * 252) * 100
print("\n日次: 実現ボラ（年率、日中+ナイト） vs 前日 VI"); print(dd[["vi", "rv_ann"]].round(1).to_string()); print(f"  平均 実現ボラ {dd.rv_ann.mean():.1f}% vs VI 平均 {dd.vi.mean():.1f}%  比 {(dd.rv_ann / dd.vi).mean():.2f}")
figB, ax = plt.subplots(2, 2, figsize=(13, 8.5))
for sess, col, nm in (("D", "#2a78d6", "日中"), ("E", "#eb6834", "ナイト")):
    p = prof[prof.session == sess]; ax[0, 0].bar(p.hour_lbl, p.v, color=col, alpha=0.7, label=nm); ax[0, 1].plot(p.hour_lbl, p.rv_ann, marker="o", color=col, label=nm)
ax[0, 0].set_title("時間帯別 平均出来高（枚/時間）", loc="left", fontsize=10); ax[0, 0].tick_params(axis="x", rotation=45, labelsize=8); ax[0, 0].legend(fontsize=8); ax[0, 0].grid(alpha=0.3, axis="y")
ax[0, 1].axhline(hb.vi.mean(), color="#888780", ls="--", lw=1, label=f"VI 平均 {hb.vi.mean():.1f}"); ax[0, 1].set_title("時間帯別 実現ボラ（1 分足、年率 %）", loc="left", fontsize=10)
ax[0, 1].tick_params(axis="x", rotation=45, labelsize=8); ax[0, 1].legend(fontsize=8); ax[0, 1].grid(alpha=0.3)
res_v = smf.ols("lv ~ C(sh)", data=hb).fit().resid; res_rv = smf.ols("lrv ~ C(sh)", data=hb).fit().resid
ax[1, 0].scatter(res_v, res_rv, s=14, c=np.where(hb.session == "D", "#2a78d6", "#eb6834"), alpha=0.7); b = np.polyfit(res_v, res_rv, 1); xg = np.linspace(res_v.min(), res_v.max(), 50)
ax[1, 0].plot(xg, np.polyval(b, xg), color="#0b0b0b", lw=1.2); ax[1, 0].set_title(f"時間帯効果を除いた log 出来高 vs log 実現分散（傾き {b[0]:.2f}）", loc="left", fontsize=10)
ax[1, 0].set_xlabel("log 出来高（時間帯平均からの乖離）"); ax[1, 0].set_ylabel("log 実現分散（同）"); ax[1, 0].grid(alpha=0.3)
ax[1, 1].scatter(dd.vi, dd.rv_ann, s=30, color="#14285a")
for d_, r_ in dd.iterrows():
    ax[1, 1].annotate(d_[4:6] + "/" + d_[6:], (r_.vi, r_.rv_ann), fontsize=7, xytext=(3, 3), textcoords="offset points")
lim = [min(dd.vi.min(), dd.rv_ann.min()) * 0.9, max(dd.vi.max(), dd.rv_ann.max()) * 1.1]; ax[1, 1].plot(lim, lim, color="#888780", ls="--", lw=1)
ax[1, 1].set_xlabel("前日 日経VI（%）"); ax[1, 1].set_ylabel("当日 実現ボラ（年率 %、日中+ナイト）"); ax[1, 1].set_title("日次: VI と実現ボラ", loc="left", fontsize=10); ax[1, 1].grid(alpha=0.3)
figB.suptitle("日経225先物ラージ 1 分足（2026-04-17〜05-14）: 日経VI × 時間当たり出来高 × 時間当たり変動", x=0.01, ha="left", fontsize=12); figB.tight_layout()
figB.savefig(OUT + "tick_vi_volume_variation.png", dpi=150); print("saved B")
