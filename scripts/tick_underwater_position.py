"""セッション（デイ／イブ）を 1 クッションとして、エントリー価格がそのセッションの値幅のどの位置（有利側の端からのパーセンタイル）にあったかと、
次のセッション終了までに含み損を抱えている時間の割合の関係を、VI 水準別に集計する。買いは安値側から、売りは高値側からのパーセンタイルで合算。"""
import os, numpy as np, pandas as pd, matplotlib; matplotlib.use("Agg"); import matplotlib.pyplot as plt
from matplotlib import font_manager
for f in ("C:/Windows/Fonts/YuGothR.ttc", "C:/Windows/Fonts/meiryo.ttc"):
    if os.path.exists(f):
        font_manager.fontManager.addfont(f); plt.rcParams["font.family"] = font_manager.FontProperties(fname=f).get_name(); break
OUT = "H:/マイドライブ/bank/nk225op_sq_rnd_2611/"
bar = pd.read_parquet("C:/Users/kawai/AppData/Local/Temp/tick/bars_1min.parquet").sort_values(["sess_id", "min"])
vi = pd.read_csv("data/nikkei_vi_daily.csv", parse_dates=["date"]).set_index("date").close
order = sorted(bar.sess_id.unique(), key=lambda s: (s[:8], 0 if s.endswith("D") else 1))
series = {}
for sid, g in bar.groupby("sess_id"):
    grid = np.arange(g["min"].min(), g["min"].max() + 1); series[sid] = g.set_index("min").c.reindex(grid).ffill().values
rows = []
for i, sid in enumerate(order[:-1]):
    nxt = order[i + 1]; date, sess = sid.split("_"); d = pd.Timestamp(date)
    # 連続性: D→同日 E、E→翌営業日 D のみ採用
    if not ((sess == "D" and nxt == f"{date}_E") or (sess == "E" and nxt.endswith("_D") and nxt[:8] > date)):
        continue
    v = vi[vi.index < d].iloc[-1] if sess == "D" else (vi[d] if d in vi.index else vi[vi.index < d].iloc[-1])
    c = series[sid]; cn = series[nxt]; H, L = c.max(), c.min(); n = len(c)
    if H - L <= 0:
        continue
    runmax, runmin = np.maximum.accumulate(c), np.minimum.accumulate(c)
    for t in range(30, n - 1):  # 寄り後 30 分以降をエントリー候補
        e = c[t]; fut = np.concatenate([c[t + 1:], cn]); nf = len(fut)
        loc_post = (e - L) / (H - L); rr = runmax[t] - runmin[t]; loc_pre = (e - runmin[t]) / rr if rr > 0 else np.nan
        uw_long = (fut < e).mean(); uw_short = (fut > e).mean()
        rows.append(dict(sess_id=sid, sess=sess, vi=v, side="買い", pct_post=loc_post, pct_pre=loc_pre, uw=uw_long, final=(fut[-1] - e)))
        rows.append(dict(sess_id=sid, sess=sess, vi=v, side="売り", pct_post=1 - loc_post, pct_pre=(1 - loc_pre) if rr > 0 else np.nan, uw=uw_short, final=(e - fut[-1])))
df = pd.DataFrame(rows); q_lo, q_hi = np.percentile(df.drop_duplicates("sess_id").vi, [33, 67])
df["tier"] = np.where(df.vi <= q_lo, f"低 VI(≤{q_lo:.0f})", np.where(df.vi > q_hi, f"高 VI(>{q_hi:.0f})", "中 VI"))
edges = [0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.7, 1.0001]; labels = ["0-10%", "10-20%", "20-30%", "30-40%", "40-50%", "50-70%", "70-100%"]
print(f"セッション数 {df.sess_id.nunique()}（次セッションが連続するもの）、VI 境界 {q_lo:.1f}/{q_hi:.1f}")
for col, nm in (("pct_post", "セッション確定後の値幅に対する位置（事後）"), ("pct_pre", "エントリー時点までの値幅に対する位置（事前、寄り 30 分以降）")):
    d2 = df.dropna(subset=[col]).copy(); d2["bucket"] = pd.cut(d2[col], edges, labels=labels, right=False)
    print(f"\n== {nm}: 有利側の端からの位置 × VI → 次セッション終了までに含み損の時間割合（平均 %）／含み損時間が 25% 以下で済む確率 ==")
    for sess, snm in (("D", "デイでエントリー → イブ終了まで"), ("E", "イブでエントリー → 翌デイ終了まで")):
        s = d2[d2.sess == sess]
        t1 = s.groupby(["bucket", "tier"], observed=True).uw.mean().unstack() * 100; t2 = s.groupby(["bucket", "tier"], observed=True).uw.apply(lambda x: (x <= 0.25).mean()).unstack() * 100
        t = pd.concat({"含み損時間 平均%": t1, "P(含み損時間≤25%)": t2}, axis=1); print(f"-- {snm} --"); print(t.round(0).to_string())
    # 逆引き: 含み損時間の平均を Y% 以下にするには端から何 % 以内で入る必要があるか
    print("逆引き（買い・売り合算、デイ＋イブ）: 含み損時間の平均 ≤ Y% を満たす最大の位置パーセンタイル")
    inv = []
    for tier in sorted(d2.tier.unique()):
        s = d2[d2.tier == tier].sort_values(col); cum = s.uw.expanding().mean(); rec = {"VI": tier}
        for Y in (0.10, 0.25, 0.40):
            ok = s[col][cum <= Y]; rec[f"≤{int(Y*100)}%"] = f"{ok.max()*100:.0f}%" if len(ok) else "なし"
        inv.append(rec)
    print(pd.DataFrame(inv).to_string(index=False))
# 図: 事前位置 × 含み損時間（VI 別、デイ／イブ）
fig, ax = plt.subplots(1, 2, figsize=(13, 5)); cols = {}
for a, sess, snm in ((ax[0], "D", "デイでエントリー → イブ終了まで"), (ax[1], "E", "イブでエントリー → 翌デイ終了まで")):
    s = df.dropna(subset=["pct_pre"]); s = s[s.sess == sess]
    for tier, col_ in zip(sorted(s.tier.unique()), ("#1d9e75", "#ba7517", "#993556")):
        g = s[s.tier == tier]; b = pd.cut(g.pct_pre, np.linspace(0, 1, 11), include_lowest=True); m = g.groupby(b, observed=True).uw.mean() * 100
        a.plot([iv.mid * 100 for iv in m.index], m.values, marker="o", lw=2, color=col_, label=tier)
    a.axhline(25, color="#888780", ls=":", lw=1); a.axhline(50, color="#888780", ls=":", lw=1); a.set_xlabel("エントリー時点の値幅内の位置（有利側の端から %）"); a.set_ylabel("含み損の時間割合（%）"); a.set_title(snm, loc="left", fontsize=10); a.grid(alpha=0.3); a.legend(fontsize=8)
fig.suptitle("日経225先物 1 分足（2026-04-17〜05-14）: セッション値幅内の位置と、次セッション終了までの含み損時間（事前の値幅、買い・売り合算）", x=0.01, ha="left", fontsize=10); fig.tight_layout(); fig.savefig(OUT + "tick_underwater_position.png", dpi=150); print("saved")
df.to_parquet(OUT.replace("H:/", "C:/Users/kawai/AppData/Local/Temp/tick/").replace("マイドライブ/bank/nk225op_sq_rnd_2611/", "") + "underwater_rows.parquet") if False else None
