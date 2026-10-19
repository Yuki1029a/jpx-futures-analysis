"""日経225オプション 11月限 P65,500 ロング（顧客ショートの受け）の先物ヘッジ設計レポート（PDF）。

基準時点: 2026-10-16 日中引け近傍の QRI スナップショット。夜間（10/16 16:30〜10/17 06:15）で検証。
使い方: python scripts/build_hedge_report.py "H:/マイドライブ/bank/nk225op_sq_rnd_2611"
"""
from __future__ import annotations

import os
import sys
import warnings
from datetime import date
from pathlib import Path

warnings.filterwarnings("ignore")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import CondPageBreak, PageBreak, SimpleDocTemplate
from scipy.stats import norm

from data import iv_views
from data.qri_iv_loader import load_snapshot, prefetch_snapshots
from report_kit import A, B, BODY, CAP, EQ, H1, H2, INK, MUT, O, SMALL, SUB, TBL, TITLE, GRID, Pg, fig_to_img, fmt, make_footer, sp, style_ax

OUT_DIR = Path(sys.argv[1] if len(sys.argv) > 1 else "H:/マイドライブ/bank/nk225op_sq_rnd_2611"); OUT_DIR.mkdir(parents=True, exist_ok=True)
PDF_PATH = OUT_DIR / "nk225op_2611_p65500_hedge_report.pdf"
CM, K0, SQ, N, MULT = "2611", 65500, pd.Timestamp("2026-11-13 09:00"), 100, 1000


# ------------------------------------------------------------------ 計算
def prep(k):
    m_ = iv_views.with_eff_iv(load_snapshot(k)); g = m_[m_.contract_month == CM].copy()
    for c in ["ask_price", "bid_price", "ask_iv", "bid_iv", "delta", "gamma", "theta", "vega"]:
        g[c] = pd.to_numeric(g[c], errors="coerce")
    g["mid"] = ((g.ask_price + g.bid_price) / 2).where(g.ask_price.notna() & g.bid_price.notna())
    pv = g.pivot_table(index="strike", columns="option_type", values="mid").dropna(); par = pv.CALL - pv.PUT
    near = par.abs().nsmallest(5).index; F = float(np.median(near + par[near]))
    d = k.split("/")[-2]; ts = pd.Timestamp(f"{d[:4]}-{d[4:6]}-{d[6:8]} {iv_views.key_time_label(k)}")
    return ts, F, (SQ - ts).total_seconds() / (365 * 86400), g


def greeks(F, Kv, T, s):
    d1 = (np.log(F / Kv) + 0.5 * s * s * T) / (s * np.sqrt(T)); d2 = d1 - s * np.sqrt(T)
    return dict(price=Kv * norm.cdf(-d2) - F * norm.cdf(-d1), delta=norm.cdf(d1) - 1, gamma=norm.pdf(d1) / (F * s * np.sqrt(T)),
                vega=F * norm.pdf(d1) * np.sqrt(T) / 100, theta=-F * norm.pdf(d1) * s / (2 * np.sqrt(T)) / 365, d1=d1, d2=d2)


ks = [k for k in iv_views.snapshot_keys("20261016") if iv_views.key_time_label(k) <= "15:15:59"]
ts0, F0, T0, g0 = prep(ks[-1]); row0 = g0[(g0.strike == K0) & (g0.option_type == "PUT")].iloc[0]
sig0 = float(row0.eff_iv); G0 = greeks(F0, K0, T0, sig0); GK = greeks(K0, K0, T0, sig0)
# 12月限の先物価格（パリティ）
m12 = iv_views.with_eff_iv(load_snapshot(ks[-1])); g12 = m12[m12.contract_month == "2612"].copy()
for c in ["ask_price", "bid_price"]:
    g12[c] = pd.to_numeric(g12[c], errors="coerce")
g12["mid"] = ((g12.ask_price + g12.bid_price) / 2).where(g12.ask_price.notna() & g12.bid_price.notna())
pv12 = g12.pivot_table(index="strike", columns="option_type", values="mid").dropna(); par12 = pv12.CALL - pv12.PUT; near12 = par12.abs().nsmallest(5).index
F12 = float(np.median(near12 + par12[near12]))
# ∂σ/∂F の回帰（ATM 近傍のプット、15 分差分）
days = [d for d in sorted(iv_views.list_available_days(15)) if "20261009" <= d <= "20261016"]
keys = [k for d in days for k in iv_views.snapshot_keys(d)]; prefetch_snapshots(keys)
rows = []
for k in keys:
    gg = iv_views._month_reduced(k, CM)
    if gg is None or int(gg.q.sum()) < 150:
        continue
    ts_, F_, T_, g_ = prep(k)
    q_ = g_[g_.ask_iv.notna() & g_.bid_iv.notna()].copy(); q_["m"] = np.log(q_.strike / F_) / (sig0 * np.sqrt(T_))
    q_ = q_[(q_.m.abs() < 0.5) & (q_.option_type == "PUT")]
    rows.append(pd.DataFrame({"ts": ts_, "F": F_, "strike": q_.strike, "iv": q_.eff_iv}))
r = pd.concat(rows).sort_values(["strike", "ts"]); gpr = r.groupby("strike")
r["dF"] = gpr.F.diff(); r["ds"] = gpr.iv.diff(); r["dt"] = gpr.ts.diff().dt.total_seconds() / 60
r = r[(r.dt <= 20) & (r.dF.abs() > 5)]


def beta_of(rr):
    return float(np.cov(rr.ds * 100, rr.dF)[0, 1] / np.var(rr.dF, ddof=1) * 1000), len(rr)


REG = [("10/9〜10/16 全期間", r), ("10/9〜10/14（急落前）", r[r.ts <= "2026-10-14 23:59"]), ("10/15〜10/16（急落以降）", r[r.ts >= "2026-10-15"]), ("10/16 のみ", r[r.ts >= "2026-10-16"])]
betas = {lab: beta_of(rr) for lab, rr in REG}
h_black = -G0["delta"] * N * MULT / 1000
h_mv = {lab: -(G0["delta"] + G0["vega"] * b / 1000) * N * MULT / 1000 for lab, (b, n_) in betas.items()}
BETA_REF = betas["10/9〜10/14（急落前）"][0]; H_REF = h_mv["10/9〜10/14（急落前）"]
# ラダー
ladder = []
for Fv in range(63500, 67501, 500):
    gv = greeks(Fv, K0, T0, sig0)
    ladder.append((Fv, gv["delta"], -gv["delta"] * N, -(gv["delta"] + gv["vega"] * BETA_REF / 1000) * N))
# 夜間の経路
nk = [k for k in iv_views.snapshot_keys("20261016") if iv_views.key_time_label(k) >= "16:30:00"] + [k for k in iv_views.snapshot_keys("20261017") if iv_views.key_time_label(k) <= "06:15:59"]
prefetch_snapshots(nk); path = []
for k in nk:
    gg = iv_views._month_reduced(k, CM)
    if gg is None or int(gg.q.sum()) < 150:
        continue
    ts_, F_, T_, g_ = prep(k); rw = g_[(g_.strike == K0) & (g_.option_type == "PUT")]
    if rw.empty or pd.isna(rw.mid.iloc[0]):
        continue
    path.append(dict(ts=ts_, F=F_, P=float(rw.mid.iloc[0]), iv=float(rw.eff_iv.iloc[0])))
path = pd.DataFrame(path)
HS = {"Black 47枚": round(h_black), "基準 53枚": round(H_REF), "急落以降 64枚": round(h_mv["10/15〜10/16（急落以降）"])}
for lab, h in HS.items():
    path[lab] = ((path.P - row0.mid) * N * MULT + h * MULT * (path.F - F0)) / 1e4
end = path.iloc[-1]; dF, dsig, ddays = end.F - F0, end.iv - sig0, (end.ts - ts0).total_seconds() / 86400
gamma_term = 0.5 * G0["gamma"] * dF ** 2 * N * MULT / 1e4; vega_term = G0["vega"] * dsig * 100 * N * MULT / 1e4
theta_term = G0["theta"] * ddays * N * MULT / 1e4; delta_term = G0["delta"] * dF * N * MULT / 1e4
opt_change = (end.P - row0.mid) * N * MULT / 1e4; resid = opt_change - (delta_term + gamma_term + vega_term + theta_term)
be = sig0 * F0 * np.sqrt(1 / 365)


# ------------------------------------------------------------------ 図
def fig_convex():
    Fs = np.linspace(F0 - 3000, F0 + 3000, 200); Pv = np.array([greeks(f, K0, T0, sig0)["price"] for f in Fs])
    fig, a = plt.subplots(figsize=(7.2, 3.3))
    a.plot(Fs, (Pv - G0["price"]) * N * MULT / 1e4, color=B, lw=2, label="P65,500 ×100枚の価値変化（IV 一定）")
    a.plot(Fs, G0["delta"] * (Fs - F0) * N * MULT / 1e4, color=MUT, lw=1.2, ls="--", label="接線（デルタ分 = 先物ヘッジが打ち消す部分）")
    a.plot(Fs, ((Pv - G0["price"]) + (-G0["delta"]) * (Fs - F0)) * N * MULT / 1e4, color=O, lw=2, label="ヘッジ後（オプション + 先物 47 枚買い）")
    a.axvline(F0, color=INK, lw=0.8); a.axhline(0, color=MUT, lw=0.8)
    style_ax(a, "図1  先物価格に対するプット価値と、デルタヘッジ後の価値（100枚、万円、時間経過なし）", "先物価格 F（円）", "価値の変化（万円）")
    a.legend(fontsize=8, frameon=False, loc="upper center")
    return fig_to_img(fig)


def fig_ladder():
    Fs = np.linspace(62000, 69000, 300); hb = np.array([-greeks(f, K0, T0, sig0)["delta"] * N for f in Fs])
    hm = np.array([-(greeks(f, K0, T0, sig0)["delta"] + greeks(f, K0, T0, sig0)["vega"] * BETA_REF / 1000) * N for f in Fs])
    fig, a = plt.subplots(figsize=(7.2, 3.2))
    a.plot(Fs, hb, color=B, lw=2, label="Black デルタ（IV 一定）"); a.plot(Fs, hm, color=O, lw=2, label=f"最小分散デルタ（∂σ/∂F = {BETA_REF:+.2f}pt/千円）")
    for Fv, _, hb_, hm_ in ladder:
        a.scatter([Fv], [hb_], color=B, s=14, zorder=3)
    a.axvline(F0, color=INK, lw=0.8); a.text(F0, a.get_ylim()[1] if False else hb.max() * 0.98, f" F={F0:,.0f}", fontsize=8, color=INK, va="top")
    style_ax(a, "図2  先物価格ごとに保有すべき先物ラージ枚数（P65,500 ×100枚、買い）", "先物価格 F（円）", "先物ラージ（枚）")
    a.legend(fontsize=8, frameon=False, loc="upper right")
    return fig_to_img(fig)


def fig_pnl():
    x = np.linspace(-3000, 3000, 300); dt = 1 / 365
    base = 0.5 * G0["gamma"] * x ** 2 + G0["theta"] * 1.0
    dsig_resp = BETA_REF / 1000 * x  # %pt
    pb = base + G0["vega"] * dsig_resp + (G0["delta"] + h_black / N) * x          # Black ヘッジ、IV が反応
    pm = base + G0["vega"] * dsig_resp + (G0["delta"] + H_REF / N) * x            # 最小分散ヘッジ
    fig, a = plt.subplots(figsize=(7.2, 3.2))
    a.plot(x, base * N * MULT / 1e4, color=MUT, lw=1.2, ls="--", label="IV 一定・Black ヘッジ（ガンマ − セータ）")
    a.plot(x, pb * N * MULT / 1e4, color=B, lw=2, label="IV が反応・Black ヘッジ 47 枚")
    a.plot(x, pm * N * MULT / 1e4, color=O, lw=2, label="IV が反応・最小分散ヘッジ 53 枚")
    a.axhline(0, color=MUT, lw=0.8)
    for s_ in (-1, 1):
        a.axvline(s_ * be, color=A, lw=1, ls=":")
    a.text(be, a.get_ylim()[1] * 0.9 if False else (0.5 * G0["gamma"] * 3000 ** 2 + G0["theta"]) * N * MULT / 1e4 * 0.9, f" 損益分岐 ±{be:,.0f}円", fontsize=8, color=A)
    style_ax(a, f"図3  1日後の損益（100枚、万円）。IV の反応は ∂σ/∂F = {BETA_REF:+.2f}pt/千円 を仮定", "先物の変化 ΔF（円）", "損益（万円）")
    a.legend(fontsize=8, frameon=False, loc="upper center")
    return fig_to_img(fig)


def fig_night():
    fig, ax = plt.subplots(3, 1, figsize=(7.2, 7.2), sharex=True)
    t = path.ts
    ax[0].plot(t, path.F, color=INK, lw=1.6); ax[0].axhline(F0, color=MUT, lw=0.8, ls=":"); style_ax(ax[0], "図4  夜間立会（10/16 16:30 〜 10/17 06:00）の推移", "", "先物価格 F（円）")
    ax[1].plot(t, path.P, color=B, lw=1.6); ax[1].axhline(row0.mid, color=MUT, lw=0.8, ls=":"); style_ax(ax[1], "", "", "P65,500 仲値（円）")
    for (lab, h), col in zip(HS.items(), (B, O, A)):
        ax[2].plot(t, path[lab], color=col, lw=1.6, label=lab)
    ax[2].axhline(0, color=MUT, lw=0.8); style_ax(ax[2], "", "時刻", "ヘッジ後損益（万円）"); ax[2].legend(fontsize=8, frameon=False, loc="upper left")
    fig.autofmt_xdate()
    return fig_to_img(fig)


# ------------------------------------------------------------------ 文書
S = []
S += [sp(10), Pg("日経225オプション 2026年11月限 P65,500<br/>ロング・ポジションの先物ヘッジ設計", TITLE),
      Pg("10/16 引け直前の顧客ショートを受けた場合の数量と運用、夜間での検証", SUB),
      TBL([["項目", "内容"],
           ["ポジション", f"11月限 プット 行使価格 65,500 円を顧客から買い取り（こちらがロング）。例示数量 {N} 枚（1 枚 = 指数 × 1,000）"],
           ["基準時点", f"QRI スナップショット {ts0:%Y-%m-%d %H:%M}（日中立会の引け近傍。配信は 15 分遅延）"],
           ["先物価格の推定値", f"F = {fmt(F0)} 円（11月 SQ 決済ベース。プット・コール・パリティから推定）"],
           ["P65,500 の気配", f"買い {row0.bid_price:,.0f} / 売り {row0.ask_price:,.0f} 円、仲値 {row0.mid:,.0f} 円、IV {sig0*100:.2f}%"],
           ["残存期間", f"{T0*365:.1f} 暦日（SQ算出 11/13 9:00 まで）"],
           ["ヘッジ手段", "日経225先物 12月限ラージ（端数はミニ）"],
           ["検証", "夜間立会 10/16 16:30 〜 10/17 06:02 の実データ"]], [40, 138]),
      sp(6), Pg("要約", H1),
      Pg(f"・買い持ちのプットはデルタが負なので、ヘッジは先物の買いである。Black–Scholes のデルタ {G0['delta']:.3f} に基づく数量は {N} 枚あたり先物ラージ {h_black:.1f} 枚（F がちょうど 65,500 円なら {-GK['delta']*N:.1f} 枚）。"),
      Pg(f"・ただし本相場では先物が下がるときに IV が上がる。この連動を織り込んだ最小分散デルタは、急落前の平常時の反応（∂σ/∂F = {BETA_REF:+.2f}pt/千円）で {H_REF:.0f} 枚、急落日を含む反応では {h_mv['10/15〜10/16（急落以降）']:.0f} 枚になる。"
         f"推奨は {H_REF:.0f} 枚を基準とし、Black の {h_black:.0f} 枚を下限、{h_mv['10/15〜10/16（急落以降）']:.0f} 枚を上限とする。"),
      Pg(f"・買い持ちはガンマが正で、先物が 1,000 円下がるごとに {G0['gamma']*1000*N:.1f} 枚分の買い増し、上がれば同量の売りで調整する。対価はセータ 1 日 {-G0['theta']*N*MULT/1e4:,.0f} 万円で、1 日の値幅が {be:,.0f} 円を超えればガンマが上回る。"),
      Pg(f"・執行は引け（15:15）の板寄せで行う。15:15 から夜間立会の 16:30 までは先物が取引できず、無ヘッジの時間を残さないためである。12月限とのベーシスは {F12-F0:+,.0f} 円で、数量の補正は不要。"),
      Pg(f"・夜間の検証: 先物は {dF:+,.0f} 円動き、IV は {dsig*100:+.1f}pt 上がった。ヘッジ後損益は Black 47 枚で {path['Black 47枚'].iloc[-1]:+,.0f} 万円、基準 53 枚で {path['基準 53枚'].iloc[-1]:+,.0f} 万円。値幅が損益分岐を超えたガンマ収益と IV 上昇のベガ収益が源泉で、最小分散ヘッジは設計どおり損益のぶれを小さくした。"),
      sp(3), Pg("構成", H2),
      Pg("1. 設定と前提　2. 記号　3. 理論：デルタヘッジと IV 連動の修正　4. 数量の設計　5. 運用（調整・執行）　6. 夜間の検証　7. 留意点　付録", SMALL),
      PageBreak()]

S += [Pg("1. 設定と前提", H1),
      Pg("「P65,500 の顧客ショートを受けた」を、顧客がプットを売り、こちらがその相手方として買い持ち（ロング）になった、と解釈する。逆にこちらが売り持ちなら、本書の先物の売買方向と損益の符号をすべて反転すればよい。"),
      Pg(f"時点は 10/16 の日中立会の引け直前とし、QRI の {ts0:%H:%M} スナップショットを用いる。先物の気配は収集していないため、先物価格はオプションのパリティ F = K + C − P を ATM 近傍 5 行使で計算した中央値 {fmt(F0)} 円を使う。"
         f"ご指定の「65,500 円付近」との差は {F0-K0:+,.0f} 円で、デルタの差は {G0['delta']-GK['delta']:+.3f} である。両方の値を示す。"),
      Pg("オプションは 11月 SQ 値で決済されるが、ヘッジに使う先物は 12月限（期近）である。12月限オプションのパリティから見た 12 月の先物価格は "
         f"{fmt(F12)} 円で、11 月の {fmt(F0)} 円との比は {F0/F12:.4f}。先物が 1 円動くとき 11 月の先物価格もほぼ 1 円動くとみなせ、数量の補正（0.1%）は無視できる。"),
      Pg("金利はゼロとする（残存 28 日・金利 0.5% で価格への影響 0.04%）。IV は QRI の両気配の仲値を用いる。"),
      Pg("2. 記号", H1),
      TBL([["記号", "意味", "本書での値"],
           ["F", "先物価格（11月 SQ 決済ベース）", f"{fmt(F0)} 円"],
           ["K", "行使価格", "65,500 円"],
           ["T", "残存期間（年、暦日 ÷ 365）", f"{T0:.5f}（{T0*365:.1f} 日）"],
           ["σ", "P65,500 の IV（年率）", f"{sig0*100:.2f}%"],
           ["P(F, σ, t)", "プット 1 単位の理論価格", f"{G0['price']:,.1f} 円（仲値 {row0.mid:,.0f} 円）"],
           ["Δ（デルタ）", "F が 1 円動いたときの P の変化", f"{G0['delta']:.4f}"],
           ["Γ（ガンマ）", "F が 1 円動いたときの Δ の変化", f"{G0['gamma']:.7f} /円"],
           ["Vega（ベガ）", "σ が 1%pt 動いたときの P の変化", f"{G0['vega']:.1f} 円"],
           ["Θ（セータ）", "1 日経過したときの P の変化", f"{G0['theta']:.1f} 円/日"],
           ["h", "保有する先物ラージの枚数（買いが正）", "設計対象"],
           ["N", "オプションの枚数", f"{N} 枚（1 枚 = 1,000 単位）"],
           ["∂σ/∂F", "F が 1,000 円動いたときの固定行使 IV の変化（%pt）", "第3.3節で推定"]], [30, 90, 58]),
      Pg("3. 理論：デルタヘッジと IV 連動の修正", H1), Pg("3.1 なぜ先物を買うのか", H2),
      Pg("ポジションの価値は、オプション N 枚と先物 h 枚の合計である。先物は値動きがそのまま損益になる（価格は F そのもの）。"),
      EQ(r"V = N\cdot 1000\cdot P(F,\sigma,t) + h\cdot 1000\cdot F"),
      Pg("F が ΔF、σ が Δσ、時間が Δt だけ変わったときの価値の変化を、2 次までの近似で書くと次のようになる。"),
      EQ(r"\Delta V \approx 1000\left[(N\Delta + h)\,\Delta F + \frac{1}{2}N\Gamma(\Delta F)^{2} + N\,\mathrm{Vega}\,\Delta\sigma + N\,\Theta\,\Delta t\right]"),
      Pg("第 1 項は F の変化に比例する部分で、符号が分からない ΔF に対して損益が線形に出る。ヘッジとは、この項を消す h を選ぶことである。"),
      EQ(r"h_{Black} = -N\,\Delta"),
      Pg(f"プットの Δ は負なので h は正、すなわち先物の買いになる。残る項は ΔF の 2 乗に比例する部分（ガンマ、買い持ちでは常に利益）、IV の変化に比例する部分（ベガ）、時間経過（セータ、買い持ちでは常に損失）である。図1 はこの関係を示す。"),
      fig_convex(),
      Pg("3.2 ブラック式のグリークス", H2),
      Pg("先物を原資産とするプットの理論価格と感応度は次のとおりである（Black 1976。N は標準正規分布の累積分布関数、φ は密度関数）。"),
      EQ(r"P = K\,N(-d_{2}) - F\,N(-d_{1}),\qquad d_{1} = \frac{\ln(F/K) + \sigma^{2}T/2}{\sigma\sqrt{T}},\quad d_{2} = d_{1} - \sigma\sqrt{T}"),
      EQ(r"\Delta = N(d_{1}) - 1,\qquad \Gamma = \frac{\varphi(d_{1})}{F\sigma\sqrt{T}},\qquad \mathrm{Vega} = \frac{F\varphi(d_{1})\sqrt{T}}{100},\qquad \Theta = -\frac{F\varphi(d_{1})\sigma}{2\sqrt{T}\cdot 365}"),
      Pg(f"<b>数値。</b>F = {fmt(F0)}、K = 65,500、T = {T0:.5f}、σ = {sig0:.4f} を代入すると d₁ = {G0['d1']:.4f}、d₂ = {G0['d2']:.4f}、N(d₁) = {norm.cdf(G0['d1']):.4f}、φ(d₁) = {norm.pdf(G0['d1']):.4f}。"
         f"よって P = {G0['price']:,.1f} 円（仲値 {row0.mid:,.0f} 円と一致）、Δ = {G0['delta']:.4f}、Γ = {G0['gamma']:.7f}、Vega = {G0['vega']:.1f} 円、Θ = {G0['theta']:.1f} 円/日。"
         f"QRI 配信の値（Δ {row0.delta:.4f}、Vega {row0.vega:.1f}、Θ {row0.theta:.1f}）と一致する。"),
      Pg("3.3 IV が先物に連動するときの修正（最小分散デルタ）", H2),
      Pg("3.1 の式は σ が F と独立に動く前提である。実際の株価指数では、F が下がると同じ行使価格の IV が上がる（第1章の相場でも観測される）。IV の変化のうち F に連動する部分を"),
      EQ(r"\Delta\sigma \approx \frac{\partial\sigma}{\partial F}\,\Delta F + \varepsilon"),
      Pg("と書くと（ε は F と無関係な部分）、価値の変化のうち ΔF に比例する項は (NΔ + h)ΔF + N·Vega·(∂σ/∂F)·ΔF となる。これを消す h は"),
      EQ(r"h_{mv} = -N\left(\Delta + \mathrm{Vega}\cdot\frac{\partial\sigma}{\partial F}\right)"),
      Pg("である（最小分散デルタ。損益の分散を最小にするという意味で、Hull and White 2017 の定式化と同じ）。プットの買い持ちでは Vega > 0、∂σ/∂F < 0 なので修正項は負、すなわち Black のデルタより多く先物を買う必要がある。"
         "理由は直観的である。F が下がるとプットは Δ の分だけ値上がりするが、同時に IV が上がる分も値上がりする。Black のデルタだけヘッジしていると、下落時にオプションの利益が先物の損失を上回り、上昇時はその逆になる。"),
      Pg("<b>∂σ/∂F の推定。</b>11月限の ATM 近傍（標準化マネーネス |ln(K/F)/(σ√T)| < 0.5）のプットについて、15 分おきのスナップショット間の IV 変化 Δσ と先物変化 ΔF を集め、最小二乗の傾き β = Cov(Δσ, ΔF)/Var(ΔF) を求めた（単位: %pt / 1,000 円）。相場局面で値が大きく変わるため、期間を分けて示す。"),
      TBL([["推定期間", "∂σ/∂F（%pt / 1,000円）", "標本数", "h_mv（100 枚あたり、ラージ）"]] +
          [[lab, f"{b:+.3f}", f"{n_:,}", f"{h_mv[lab]:.1f} 枚"] for lab, (b, n_) in betas.items()], [52, 40, 26, 60]),
      Pg(f"表3  Black のデルタでは {h_black:.1f} 枚。例: 急落前の β = {BETA_REF:+.3f} なら h_mv = −100 × ({G0['delta']:.4f} + {G0['vega']:.1f} × {BETA_REF:+.3f}/1000) = {H_REF:.1f} 枚。", CAP),
      Pg("3.4 ガンマとセータの釣り合い", H2),
      Pg("ヘッジ後に残る損益は ½NΓ(ΔF)² + N·Vega·ε + NΘΔt である。IV が動かなければ、1 日の損益がゼロになる値幅は Θ·1 = ½Γ(ΔF)² を解いて"),
      EQ(r"|\Delta F|_{be} = \sqrt{\frac{-2\Theta}{\Gamma}} = \sigma F\sqrt{1/365}"),
      Pg(f"となる。数値では |ΔF| = {sig0:.4f} × {fmt(F0)} × √(1/365) = {be:,.0f} 円。1 日の値幅がこれより大きければガンマの利益がセータを上回り、小さければ下回る。"
         f"ヘッジした買い持ちの 1 日の期待損益は ½NΓF²(RV² − IV²)Δt と書け、実現ボラティリティ RV が IV を上回る日だけ正になる。直近 20 日の RV は 28.0% で IV {sig0*100:.1f}% とほぼ等しく、現時点でこの取引に期待値の偏りはない。"),
      PageBreak()]

S += [Pg("4. 数量の設計", H1),
      TBL([["方式", "前提", "100 枚あたりの先物ラージ", "ミニ換算"],
           ["Black デルタ", f"IV 一定。F = {fmt(F0)}", f"{h_black:.1f} 枚 買い", f"{h_black*10:.0f} 枚"],
           ["Black デルタ", "IV 一定。F = 65,500 ちょうど", f"{-GK['delta']*N:.1f} 枚 買い", f"{-GK['delta']*N*10:.0f} 枚"],
           ["最小分散（基準）", f"∂σ/∂F = {BETA_REF:+.2f}pt/千円（急落前の平常時）", f"{H_REF:.1f} 枚 買い", f"{H_REF*10:.0f} 枚"],
           ["最小分散（上限）", f"∂σ/∂F = {betas['10/15〜10/16（急落以降）'][0]:+.2f}pt/千円（急落日を含む）", f"{h_mv['10/15〜10/16（急落以降）']:.1f} 枚 買い", f"{h_mv['10/15〜10/16（急落以降）']*10:.0f} 枚"]], [34, 70, 44, 30]),
      Pg(f"推奨は基準の {H_REF:.0f} 枚である。Black の {h_black:.0f} 枚は IV の連動を無視した下限で、下落局面で過小ヘッジになる。上限の {h_mv['10/15〜10/16（急落以降）']:.0f} 枚は急落日の反応に引きずられた値で、上昇局面では先物の損失が先行する。"
         "IV の連動は相場局面で変わるため、翌日以降は直近の反応を見て 50〜60 枚の範囲で見直す。"),
      fig_ladder(),
      TBL([["先物価格 F（円）", "Δ（Black）", "保有枚数 Black", "保有枚数 最小分散（基準）"]] +
          [[fmt(Fv), f"{dl:.3f}", f"{hb_:.1f}", f"{hm_:.1f}"] for Fv, dl, hb_, hm_ in ladder], [40, 36, 44, 56]),
      Pg(f"表5  F の水準ごとに保有すべき先物ラージ（100 枚あたり）。隣り合う行の差が、F が 500 円動いたときの調整枚数（ATM 付近で約 {G0['gamma']*500*N:.1f} 枚）。IV は {sig0*100:.1f}% で固定。", CAP),
      Pg("5. 運用（調整・執行）", H1),
      Pg("<b>調整の向き。</b>買い持ちはガンマが正である。F が下がると Δ はより負になる（表5）ので先物を買い増し、F が上がれば売る。「下がったら買い、上がったら売る」を機械的に繰り返すことで、値動きの往復がガンマの利益として実現する。"),
      Pg(f"<b>調整の頻度。</b>目安は F が 500 円動くごと（約 {G0['gamma']*500*N:.0f} 枚）、または立会の節目（寄り・引け・夜間の寄り）ごとである。細かく刻むほど損益のぶれは小さくなるが、売買コストが増える。"
         f"先物ラージの気配幅 10 円 × 1,000 = 1 万円を 1 往復ごとに失うので、{G0['gamma']*500*N:.0f} 枚の調整で約 {G0['gamma']*500*N:.0f} 万円のコストになる。"),
      Pg("<b>執行。</b>引け直前に受けた取引なら、引け（15:15）の板寄せで大半を執行する。15:15 から夜間立会の 16:30 までは先物を売買できず、その間は無ヘッジになるからである。端数はミニで合わせる（ラージ 1 枚 = ミニ 10 枚）。"),
      fig_pnl(),
      Pg(f"図3 は 1 日後の損益を ΔF の関数として描いたものである。IV が一定なら損益は ΔF について対称な放物線で、±{be:,.0f} 円でゼロを横切る。IV が ∂σ/∂F = {BETA_REF:+.2f}pt/千円 で反応すると、Black ヘッジの損益は下落側に傾き（下落で利益、上昇で損失）、最小分散ヘッジではこの傾きが消えて対称に戻る。", CAP),
      PageBreak()]

S += [Pg("6. 夜間の検証（10/16 16:30 〜 10/17 06:02）", H1),
      Pg(f"引けで建てた 3 通りのヘッジ（Black {HS['Black 47枚']} 枚、基準 {HS['基準 53枚']} 枚、上限 {HS['急落以降 64枚']} 枚）を調整せずに保有したと仮定し、夜間立会の 15 分おきの気配で損益を追った。"),
      fig_night(),
      TBL([["項目", "値"],
           ["先物の変化 ΔF", f"{fmt(F0)} → {fmt(end.F)} 円（{dF:+,.0f} 円）"],
           ["P65,500 仲値", f"{row0.mid:,.0f} → {end.P:,.0f} 円（{end.P-row0.mid:+,.0f} 円）"],
           ["IV", f"{sig0*100:.1f}% → {end.iv*100:.1f}%（{dsig*100:+.1f}pt）"],
           ["経過時間", f"{ddays:.2f} 日"]], [50, 128]),
      CondPageBreak(110 * mm),
      Pg("オプション 100 枚の価値変化を、第3.1節の近似式の項ごとに分けると次のようになる。"),
      TBL([["項目", "式", "万円"],
           ["デルタ項", "N·Δ·ΔF", f"{delta_term:+,.0f}"],
           ["ガンマ項", "½·N·Γ·(ΔF)²", f"{gamma_term:+,.0f}"],
           ["ベガ項", "N·Vega·Δσ", f"{vega_term:+,.0f}"],
           ["セータ項", "N·Θ·Δt", f"{theta_term:+,.0f}"],
           ["残差（高次項・気配の変化）", "", f"{resid:+,.0f}"],
           ["オプションの価値変化（実績）", "", f"{opt_change:+,.0f}"]], [60, 60, 58]),
      TBL([["ヘッジ", "先物の損益", "ヘッジ後の損益"]] + [[lab, f"{h*MULT*dF/1e4:+,.0f} 万円", f"{path[lab].iloc[-1]:+,.0f} 万円"] for lab, h in HS.items()], [50, 60, 68]),
      Pg(f"値幅 {abs(dF):,.0f} 円は損益分岐 {be:,.0f} 円の {abs(dF)/be:.1f} 倍で、ガンマ項が {gamma_term:+,.0f} 万円、IV 上昇のベガ項が {vega_term:+,.0f} 万円と、どちらも買い持ちに有利に働いた。"
         f"Black ヘッジの損益が最も大きいのは、デルタ項の一部（IV 連動分）を打ち消していないためで、上昇していれば同じ理由で損益が最も小さくなっていた。最小分散ヘッジ（基準 53 枚）の損益 {path['基準 53枚'].iloc[-1]:+,.0f} 万円は、ガンマとベガの利益からセータを引いた水準に近く、設計どおり「方向に賭けない」結果になっている。"),
      Pg("7. 留意点", H1),
      Pg("・∂σ/∂F は直近 6 営業日の推定で、相場局面により −0.4 から −2.4pt/千円 まで変わる。数量は日々、直近の反応で見直す。"),
      Pg("・IV が実現ボラティリティとほぼ等しい現在、セータの持ち出しに対するガンマの期待収益は見合っているが、超過はない。夜間の結果は 1 標本であり、繰り返せば平均してゼロに近づく。"),
      Pg("・残存 28 日では SQ 当日のピン・リスクや SQ 値と先物の乖離（SQ は構成銘柄の始値から算出）はまだ小さいが、満期週には別途の管理が必要になる。"),
      Pg("・QRI の気配は 15 分遅延で、引け直前の実勢より古い可能性がある。実際の執行価格では数量を再計算する。"),
      sp(4), Pg("付録 A. 再現方法", H1),
      Pg("scripts/build_hedge_report.py <出力フォルダ> を実行する。QRI スナップショット（R2）から 10/16 引け近傍と夜間の気配を読み、グリークス・回帰・損益を計算して本書を生成する。行使価格・限月・数量はスクリプト冒頭の定数で指定する。", BODY),
      Pg("付録 B. 参考文献", H1),
      Pg("Black, F. (1976), The Pricing of Commodity Contracts, Journal of Financial Economics 3.", SMALL),
      Pg("Hull, J. and A. White (2017), Optimal Delta Hedging for Options, Journal of Banking and Finance 82.", SMALL),
      Pg("Derman, E. (1999), Regimes of Volatility, Risk 12(4).", SMALL)]

doc = SimpleDocTemplate(str(PDF_PATH), pagesize=A4, leftMargin=16 * mm, rightMargin=16 * mm, topMargin=16 * mm, bottomMargin=18 * mm,
                        title="日経225オプション 2026年11月限 P65,500 ロングの先物ヘッジ設計", author="先物手口分析")
ft = make_footer(f"11月限 P65,500 ロングの先物ヘッジ設計（基準 {ts0:%Y-%m-%d %H:%M}、作成 {date.today():%Y-%m-%d}）")
doc.build(S, onFirstPage=ft, onLaterPages=ft)
print("written", PDF_PATH)
print({k: round(v, 1) for k, v in h_mv.items()}, "black", round(h_black, 1), "night pnl", {k: round(path[k].iloc[-1]) for k in HS})
