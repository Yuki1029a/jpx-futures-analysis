"""解説ノート: プライシングカーネル・パズル、状態依存 γ の文献、行使価格の出来高・建玉による重み付け。

使い方: python scripts/build_pk_puzzle_note.py <重み付け実験の出力txt> <出力フォルダ>
日次蓄積 cache/derived/gamma_daily.jsonl と最新日の気配から図表を作る。
"""
from __future__ import annotations

import json
import re
import sys
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")
sys.path.insert(0, str(Path(__file__).resolve().parent)); sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import statsmodels.api as sm
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import CondPageBreak, PageBreak, SimpleDocTemplate

from gamma_daily import nikkei_close
from gamma_lib import Physical, gamma_kernel, q_moments
from report_kit import A, B, BODY, CAP, EQ, GRID, H1, H2, INK, MUT, O, SMALL, SUB, TBL, TITLE, Y, Pg, fig_to_img, fmt, make_footer, sp, style_ax
from rnd_lib import density, prep

GW = Path(sys.argv[1]).read_text(encoding="utf-8"); OUT_DIR = Path(sys.argv[2]); OUT_DIR.mkdir(parents=True, exist_ok=True)
PDF_PATH = OUT_DIR / "nk225op_pricing_kernel_puzzle_state_gamma_note.pdf"
pct = lambda v: f"{v*100:.2f}%"

# ------------------------------------------------------------------ 計算
rows = [json.loads(l) for l in open("cache/derived/gamma_daily.jsonl", encoding="utf-8") if l.strip()]
df = pd.DataFrame(rows); df["date"] = pd.to_datetime(df.day.astype(str))
r = rows[-1]; DAY = r["day"]; ts, F, T, pts, g = prep(r["key"], r["cm"]); XS = np.linspace(-0.6, 0.6, 6001)
K, fK, gg, _ = density(np.array(r["p"]), F, T, XS); muQ, sQ, skQ, kuQ = q_moments(K, fK, F)
phys = Physical(nikkei_close()); xP = phys.sample(r["n_days"], r["rv20"]); muP, sP = float(xP.mean()), float(xP.std())
gk = gamma_kernel(K, fK, F, xP); beta2_gauss = 0.5 * (1 / sP ** 2 - 1 / sQ ** 2); gamma_gauss = (muP - muQ) / sQ ** 2
shift_tbl = []
for s_ in (-0.01, -0.005, 0.0, 0.005, 0.01):
    g2 = gamma_kernel(K, fK, F, xP - xP.mean() + muQ + s_); shift_tbl.append((s_, g2["gamma"], g2["gamma_quadfit"], g2["quad"]))
# 状態回帰
X = pd.DataFrame({"atm": df.atm * 100, "rv20": df.rv20 * 100, "ret20": df.ret20 * 100, "T_days": df["T"] * 365}); X["ivrv"] = X.atm - X.rv20
Xc = sm.add_constant(X[["atm", "rv20", "ret20", "T_days"]])
reg = {lab: sm.OLS(df[col], Xc).fit(cov_type="HAC", cov_kwds={"maxlags": 5}) for lab, col in [("gamma_I", "gamma_I"), ("quad", "quad")]}
corr = pd.concat([df.gamma_I, df.quad.rename("curv"), X], axis=1).corr()
# 重み付け実験の読み取り
wrows = [(m.group(1).strip(), float(m.group(2)), float(m.group(3)), float(m.group(4))) for m in re.finditer(r"^(.+?)\s+γ̂\(直線\)=([+-][\d.]+)\s+γ̂\(2次の1次項\)=([+-][\d.]+)\s+曲率=([+-][\d.]+)", GW, re.M)]
atm_line = re.search(r"ATM IV: 現行 ([\d.]+)% 出来高重み ([\d.]+)% 建玉重み ([\d.]+)%", GW).groups()
vol_top = re.search(r"出来高上位: (\[\[.*?\]\])", GW).group(1)
# 重みのプロファイル（行使価格ごと）
g["volume"] = pd.to_numeric(g.volume, errors="coerce").fillna(0); g["oi"] = pd.to_numeric(g.oi, errors="coerce").fillna(0)
otm = g[g.q & (((g.option_type == "PUT") & (g.strike < F)) | ((g.option_type == "CALL") & (g.strike >= F)))]
wv = otm.groupby("strike").volume.sum(); wo = otm.groupby("strike").oi.sum()
dK = np.gradient(K); mQ = fK * dK; mQ = mQ / mQ.sum()


# ------------------------------------------------------------------ 図
def fig_kernel():
    fig, ax = plt.subplots(1, 2, figsize=(7.2, 3.1)); x = gk["xgrid"]
    ax[0].plot(x, gk["fQ"], color=B, lw=1.8, label=f"f_Q（SVI、s_Q = {sQ*100:.2f}%）"); ax[0].plot(x, gk["fP"], color=O, lw=1.8, label=f"f_P（歴史的、s_P = {sP*100:.2f}%）")
    ax[0].axvspan(gk["xlo"], gk["xhi"], color=GRID, alpha=0.5, lw=0); ax[0].set_xlim(-0.3, 0.3); style_ax(ax[0], f"(a) 密度（{DAY[:4]}/{DAY[4:6]}/{DAY[6:]} 引け、残存 {T*365:.0f} 日）", "x = ln(S_T / F)", "密度"); ax[0].legend(fontsize=7.5, frameon=False)
    xx, lnM = gk["x"], gk["lnM"]; A1 = np.vstack([np.ones_like(xx), xx]).T; A2 = np.vstack([np.ones_like(xx), xx, xx ** 2]).T
    w = np.interp(xx, x, mQ); W = np.sqrt(w)[:, None]
    c1, *_ = np.linalg.lstsq(A1 * W, lnM * W[:, 0], rcond=None); c2, *_ = np.linalg.lstsq(A2 * W, lnM * W[:, 0], rcond=None)
    ax[1].plot(xx, lnM, color=INK, lw=2, label="ln(f_Q / f_P)（観測）"); ax[1].plot(xx, A1 @ c1, color=B, lw=1.2, ls="--", label=f"直線: γ̂ = {-c1[1]:+.2f}")
    ax[1].plot(xx, A2 @ c2, color=O, lw=1.4, label=f"2次: 傾き {-c2[1]:+.2f}、曲率 β₂ = {c2[2]:+.1f}")
    xg = np.linspace(xx.min(), xx.max(), 200); gauss = c2[0] + c2[1] * xg + beta2_gauss * xg ** 2
    ax[1].plot(xg, gauss, color=MUT, lw=1, ls=":", label=f"2次の切片・傾きのまま曲率を正規近似値 {beta2_gauss:+.1f} に")
    ax[1].axhline(0, color=MUT, lw=0.6); style_ax(ax[1], "(b) 対数プライシングカーネル（中央 5〜95%）", "x = ln(S_T / F)", "ln M + 定数"); ax[1].legend(fontsize=7, frameon=False)
    fig.tight_layout(); return fig_to_img(fig)


def fig_schematic():
    fig, ax = plt.subplots(1, 2, figsize=(7.2, 2.8)); x = np.linspace(-0.2, 0.2, 300); sp_ = 0.085
    for ratio, col, lab in ((1.00, B, "s_Q = s_P（べき効用: 直線）"), (1.08, O, "s_Q = 1.08 s_P（10/20 と同程度）"), (1.20, A, "s_Q = 1.20 s_P（高いボラ・プレミアム）")):
        sq_ = sp_ * ratio; ax[0].plot(x, -3 * x + 0.5 * (1 / sp_ ** 2 - 1 / sq_ ** 2) * x ** 2, color=col, lw=1.6, label=lab)
    ax[0].axhline(0, color=MUT, lw=0.6); style_ax(ax[0], "(a) 正規近似: ln M(x) = c − γx + β₂x²（γ = 3）", "x", "ln M"); ax[0].legend(fontsize=7, frameon=False)
    for ratio, col in ((1.00, B), (1.08, O), (1.20, A)):
        sq_ = sp_ * ratio; ax[1].plot(x, 3 - 2 * 0.5 * (1 / sp_ ** 2 - 1 / sq_ ** 2) * x, color=col, lw=1.6)
    ax[1].axhline(0, color=MUT, lw=0.6); style_ax(ax[1], "(b) 局所リスク回避度 γ(x) = −d ln M / dx", "x", "γ(x)")
    ax[1].text(0.19, -0.3, "右側で負（パズル）", ha="right", fontsize=8, color=A); fig.tight_layout(); return fig_to_img(fig)


def fig_state():
    fig, ax = plt.subplots(1, 3, figsize=(7.2, 2.7))
    ax[0].scatter(X.ivrv, df.quad, s=14, color=B, alpha=0.75); b = np.polyfit(X.ivrv, df.quad, 1); xg = np.linspace(X.ivrv.min(), X.ivrv.max(), 50); ax[0].plot(xg, np.polyval(b, xg), color=O, lw=1.2)
    style_ax(ax[0], f"(a) 曲率 β₂ と IV−RV（相関 {corr.loc['curv','ivrv']:+.2f}）", "ATM IV − RV₂₀（%pt）", "β₂")
    ax[1].scatter(X.atm, df.gamma_I, s=14, color=B, alpha=0.75); b = np.polyfit(X.atm, df.gamma_I, 1); xg = np.linspace(X.atm.min(), X.atm.max(), 50); ax[1].plot(xg, np.polyval(b, xg), color=O, lw=1.2)
    style_ax(ax[1], f"(b) γ̂（直線）と ATM IV（相関 {corr.loc['gamma_I','atm']:+.2f}）", "ATM IV（%）", "γ̂")
    ax[2].plot(df.date, df.quad, color=B, lw=1.4, label="β₂（左軸）"); a2 = ax[2].twinx(); a2.plot(df.date, X.ivrv, color=O, lw=1.2, label="IV−RV（右軸）"); a2.tick_params(colors=MUT, labelsize=8)
    style_ax(ax[2], "(c) 推移", "", "β₂"); ax[2].tick_params(axis="x", rotation=45); h1, l1 = ax[2].get_legend_handles_labels(); h2, l2 = a2.get_legend_handles_labels(); ax[2].legend(h1 + h2, l1 + l2, fontsize=7, frameon=False)
    fig.tight_layout(); return fig_to_img(fig)


def fig_weights():
    fig, a = plt.subplots(figsize=(7.2, 2.6)); x = np.log(K / F); sel = (gk["xgrid"] >= gk["xlo"]) & (gk["xgrid"] <= gk["xhi"])
    a.plot(x, mQ / mQ.max(), color=B, lw=1.8, label="現行: f_Q の確率質量")
    xv = np.log(wv.index.values / F); a.bar(xv, wv.values / wv.values.max(), width=0.006, color=O, alpha=0.7, label="出来高（当日）")
    xo = np.log(wo.index.values / F); a.plot(xo, wo.values / wo.values.max(), color=A, lw=1.2, ls="--", label="建玉")
    a.axvspan(gk["xlo"], gk["xhi"], color=GRID, alpha=0.5, lw=0); a.set_xlim(-0.4, 0.3); style_ax(a, "図4  回帰の重みの形（最大値 = 1 に正規化、帯は中央 5〜95% の推定区間）", "x = ln(K / F)", "重み"); a.legend(fontsize=7.5, frameon=False)
    return fig_to_img(fig)


# ------------------------------------------------------------------ 文書
doc = SimpleDocTemplate(str(PDF_PATH), pagesize=A4, leftMargin=16 * mm, rightMargin=16 * mm, topMargin=16 * mm, bottomMargin=16 * mm, title="プライシングカーネル・パズルと状態依存 γ")
D = f"{DAY[:4]}-{DAY[4:6]}-{DAY[6:]}"; rg, rq = reg["gamma_I"], reg["quad"]
story = [
    Pg("プライシングカーネル・パズルと状態依存 γ", TITLE),
    Pg(f"当方の γ 推定（方法 I〜III）への含意、関連研究の整理、行使価格の出来高・建玉による重み付けの検討（データ: 2026-06-03〜{D} 引け、{len(df)} 営業日）", SUB),
    Pg("要旨", H1),
    Pg("1. 「市場が織り込む γ」は、リスク中立密度 f_Q と実確率密度 f_P の比（プライシングカーネル M）の傾きである。べき効用なら ln M は x = ln(S_T/F) の一次式で、傾きの符号を反転したものが γ となる。"),
    Pg("2. 実際の ln M は一次式ではなく U 字である（プライシングカーネル・パズル）。正規近似で書くと U 字の曲率は β₂ = ½(1/s_P² − 1/s_Q²) であり、「リスク中立の標準偏差 s_Q が実確率の s_P より大きい」こと、すなわちボラティリティ・リスクプレミアムそのものである。"
       f"{D} 引けでは s_Q = {sQ*100:.2f}%、s_P = {sP*100:.2f}% で予測曲率 {beta2_gauss:+.1f}、観測曲率 {gk['quad']:+.1f}。100 営業日の曲率は ATM IV − RV₂₀ と相関 {corr.loc['curv','ivrv']:+.2f} で動く。"),
    Pg("3. 従って一次式の γ̂（方法 I の直線）は、U 字のどこを重く見るかで値が変わる（重み依存）。同時に γ̂ は「予測期間中の f_P と f_Q の平均の差 ÷ 分散」に近く、f_P の平均を 1% 動かすだけで 2 前後動く。1 日分の気配から γ の水準を決めることはできず、水準は実現 SQ 値による最尤法（方法 III）で、形状（傾きと曲率）は日次で蓄積する、という分担が妥当である。"),
    Pg("4. γ が状態変数（ボラティリティ、過去リターン、景気）に依存するという研究は多い（Campbell–Cochrane 1999 の習慣形成、Rosenberg–Engle 2002、Bliss–Panigirtzoglou 2004、Bollerslev–Gibson–Zhou 2011、Christoffersen–Heston–Jacobs 2013、Song–Xiu 2016 ほか）。共通の結論は「γ は一定ではなく、ボラティリティ（分散プレミアム）を第二の状態変数として扱うと U 字が消える」である。"),
    Pg("5. 出来高・建玉による重み付けは、SVI 当てはめでは ATM IV を 0.1%pt しか変えず、回帰の重みとしては直線の γ̂ を −0.64〜+0.04 の範囲で動かすが、2 次式の傾き（−0.82〜−0.97）と曲率（+12.9〜+15.1）はほぼ不変である。重みは γ の定義に入らないため「情報の重み」としての理論的根拠はなく、推定量の分散を下げる（測定精度の代理）としての使い方に限るのが筋である。"),
    sp(2),
    Pg("1. 準備：記号と基本式", H1),
    Pg("満期 T の先物清算値（SQ 値）を S_T、現在の先物価格を F とし、対数リターンを x = ln(S_T / F) と書く。オプション価格から復元した x の密度（リスク中立密度）を f_Q(x)、実際に起こる x の密度（実確率密度）を f_P(x) とする。"
       "満期に 1 円を支払う証券の価格を状態ごとに並べたものがプライシングカーネル M(x) で、無リスク金利 r のもとで次の関係が成り立つ。"),
    EQ(r"M(x) = e^{-rT}\,\frac{f_Q(x)}{f_P(x)} \qquad\Longleftrightarrow\qquad \ln f_Q(x) - \ln f_P(x) = rT + \ln M(x)"),
    Pg("代表的投資家が富 W = S_T のべき効用 U(W) = W^(1−γ)/(1−γ) を持つなら、M は限界効用 U′(W) = W^(−γ) に比例する。両辺の対数を取ると、"),
    EQ(r"\ln M(x) = a - \gamma\,x \qquad (\mathrm{式\,1})\qquad\quad \gamma(x) \equiv -\frac{d \ln M(x)}{dx}"),
    Pg("式 1 の右の定義は「局所的な相対的リスク回避度」である。べき効用なら γ(x) は定数 γ に等しい。方法 I はこの一次式を中央 5〜95% の x で加重回帰し、傾きの符号を反転して γ̂ とする。"),
    Pg("<b>正規近似で γ̂ の正体を見る。</b>f_P と f_Q をともに正規分布（平均 μ_P, μ_Q、標準偏差 s_P, s_Q）で近似すると、対数密度の差は x の 2 次式になる。"),
    EQ(r"\ln\frac{f_Q(x)}{f_P(x)} = c + \left(\frac{\mu_Q}{s_Q^2} - \frac{\mu_P}{s_P^2}\right)x + \frac{1}{2}\left(\frac{1}{s_P^2} - \frac{1}{s_Q^2}\right)x^2 \qquad (\mathrm{式\,2})"),
    Pg("s_P = s_Q = s のとき 2 次項は消え、一次の係数は (μ_Q − μ_P)/s² となる。式 1 と比べると、"),
    EQ(r"\gamma = \frac{\mu_P - \mu_Q}{s^2} \qquad (\mathrm{式\,3})"),
    Pg(f"つまり一次式の γ は「予測期間の期待リターン差（リスクプレミアム）÷ 分散」である。{D} 引けの数値を入れると、μ_P = {muP*100:+.2f}%、μ_Q = {muQ*100:+.2f}%、s_Q = {sQ*100:.2f}% から γ = {gamma_gauss:+.2f} となる。"
       f"ここで分母 s² = {sQ**2:.4f} は小さいため、μ_P を 1%pt 動かすと γ は約 {1/sQ**2/100:.1f} 動く。実際の f_P（正規ではない）で f_P の平均だけをずらした結果を表 1 に示す。"),
    TBL([["f_P の平均（μ_Q からのずれ）", "γ̂（直線）", "γ̂（2 次式の 1 次項）", "曲率 β₂"]] + [[f"{s_*100:+.1f}%pt", f"{g1:+.2f}", f"{g2:+.2f}", f"{q:+.1f}"] for s_, g1, g2, q in shift_tbl], [58, 36, 44, 34]),
    Pg(f"表1  {D} 引け。f_P の形を変えずに平均だけを動かした感応度。基準設定（歴史的標本の平均 μ_P = {muP*100:+.2f}%）の γ̂ は {gk['gamma']:+.2f}。", CAP),
    Pg("表 1 が示す通り、1 日分の気配から得られる γ̂ の水準は f_P の平均（＝将来 24 日間の期待リターンの仮定）に支配される。歴史的平均は 100 日どころか 20 年の標本でも標準誤差が大きく、ここが方法 I の本質的な限界である。曲率は平均のずれで 12〜26 と動くが符号は変わらない。"),
    PageBreak(),
    Pg("2. プライシングカーネル・パズル", H1),
    Pg("2.1 何が観測されたか", H2),
    Pg("Jackwerth (2000, RFS) は S&P 500 オプションから推定した f_Q と、実現リターンから推定した f_P の比を取り、1987 年の暴落以降は ln M が x の単調減少にならず、x = 0 付近で局所的に増加する（γ(x) が負になる）ことを示した。Aït-Sahalia–Lo (2000, J. Econometrics) と Rosenberg–Engle (2002, JFE) も同時期に同じ形状を報告し、以降「プライシングカーネル・パズル」と呼ばれる。"
       "後続研究の多くは U 字（左右で上昇、中央で下降）を報告しており、Beare–Schmidt (2016, J. Applied Econometrics) は単調性の統計的検定で棄却、Cuesdeanu–Jackwerth (2018, Annals of Finance) が包括的な展望論文である。"),
    Pg("なぜ「パズル」かというと、式 1 の定義で γ(x) < 0 となる区間は、投資家が「より豊かになる状態」の 1 円をより高く買っていることを意味し、富に対して単調増加・凹の効用関数を持つ代表的投資家では説明できないからである。"),
    Pg("2.2 正規近似で見る U 字の正体", H2),
    Pg("式 2 に戻ると、2 次の係数は β₂ = ½(1/s_P² − 1/s_Q²) であり、s_Q > s_P のとき正になる。すなわち <b>U 字は「リスク中立分布がヒストリカルより広い」ことと同値</b>である。リスク中立分布の幅はインプライド・ボラティリティ、実確率分布の幅は実現ボラティリティで決まるので、β₂ > 0 は「IV が RV を上回る」＝負のボラティリティ・リスクプレミアム（投資家が分散の買いにプレミアムを払う）に他ならない。"),
    fig_schematic(),
    Pg("図1  正規近似での ln M と局所リスク回避度。s_Q/s_P が 1 を超えるほど U 字が深くなり、右側で γ(x) が負になる。γ = 3 の例。", CAP),
    Pg(f"{D} 引けの実測は図 2 の通りである。観測した ln(f_Q/f_P) は中央で下がり両端で上がる U 字で、2 次式の曲率 β₂ = {gk['quad']:+.1f} に対し、正規近似の予測 ½(1/{sP:.4f}² − 1/{sQ:.4f}²) = {beta2_gauss:+.1f} である。差の {gk['quad']-beta2_gauss:+.1f} は f_Q の左歪み（RR25 で正、プット側が高い）と裾の厚さによる。"),
    fig_kernel(),
    Pg(f"図2  {D} 引け（気配時刻 {str(ts)[11:16]}）、11月限。(a) 密度、(b) 対数カーネルと直線・2 次式の当てはめ。点線は 2 次式の切片・傾きを保ったまま曲率だけを正規近似の予測値に置き換えたもの。", CAP),
    CondPageBreak(60 * mm),
    Pg("2.3 当方データでの検証：曲率はボラティリティ・プレミアムで動く", H2),
    Pg(f"2.2 の解釈が正しければ、日次の曲率 β₂ は「ATM IV − RV₂₀」と同方向に動くはずである。{len(df)} 営業日について確認した（図 3）。"),
    fig_state(),
    Pg(f"図3  日次蓄積（2026-06-03〜{D}）。(a) 曲率と IV−RV の散布図、(b) 直線 γ̂ と ATM IV、(c) 推移。", CAP),
    TBL([["被説明変数", "定数", "ATM IV（%）", "RV₂₀（%）", "20 日リターン（%）", "残存（日）", "R²"],
         ["γ̂（直線）", f"{rg.params['const']:+.2f} ({rg.tvalues['const']:+.1f})", f"{rg.params['atm']:+.3f} ({rg.tvalues['atm']:+.1f})", f"{rg.params['rv20']:+.3f} ({rg.tvalues['rv20']:+.1f})", f"{rg.params['ret20']:+.3f} ({rg.tvalues['ret20']:+.1f})", f"{rg.params['T_days']:+.3f} ({rg.tvalues['T_days']:+.1f})", f"{rg.rsquared:.2f}"],
         ["曲率 β₂", f"{rq.params['const']:+.1f} ({rq.tvalues['const']:+.1f})", f"{rq.params['atm']:+.2f} ({rq.tvalues['atm']:+.1f})", f"{rq.params['rv20']:+.2f} ({rq.tvalues['rv20']:+.1f})", f"{rq.params['ret20']:+.2f} ({rq.tvalues['ret20']:+.1f})", f"{rq.params['T_days']:+.2f} ({rq.tvalues['T_days']:+.1f})", f"{rq.rsquared:.2f}"]],
        [26, 26, 28, 28, 30, 24, 14]),
    Pg(f"表2  状態変数への回帰（n = {int(rg.nobs)}、括弧内は Newey–West t 値、ラグ 5）。", CAP),
    Pg(f"曲率は IV に正、RV に負に反応し（係数はほぼ同じ大きさで逆符号）、IV − RV との相関は {corr.loc['curv','ivrv']:+.2f} である。これは式 2 の β₂ が 1/s_P² − 1/s_Q² であることの直接の帰結で、当方の U 字はボラティリティ・プレミアムで説明できる。"
       f"直線 γ̂ は ATM IV に負（IV 1%pt 上昇で {rg.params['atm']:+.2f}）で、Bliss–Panigirtzoglou (2004) の「ボラティリティが高いとき γ は低い」と同方向だが、当方の f_P は RV₂₀ で幅を決めているため、IV ≫ RV の局面では f_Q の方が広くなり直線の傾きが機械的に寝る成分を含む。方向は文献と整合するが、独立の証拠とは言えない。"),
    Pg("2.4 説明仮説の整理", H2),
    TBL([["系統", "主張", "代表的研究", "当方データでの位置づけ"],
         ["統計的", "f_P の推定誤差（短い標本、まれな暴落の未観測＝ペソ問題、無条件密度の使用）が U 字を作る。条件付きの f_P を使えば単調に戻る。", "Linn–Shive–Shumway (2018, RFS)", "f_P を RV₂₀ で条件付けしても U 字は残る。平均のずれは傾きを動かすが曲率の符号は変えない（表 1）。"],
         ["第二状態変数", "M は価格 S だけでなくボラティリティ σ² にも依存する。負の分散プレミアム（ξ > 0）を持つ M を S だけに射影すると U 字が現れる。VIX で条件付けると U 字は消える。", "Christoffersen–Heston–Jacobs (2013, RFS)、Song–Xiu (2016, J. Econometrics)、Bakshi–Madan–Panayotov (2010, JFE)", "曲率と IV−RV の相関 0.79 はこの系統の予測そのもの。"],
         ["選好", "確率ウエイト（ランク依存効用）、異質な信念と空売り制約、センチメント。", "Polkovnichenko–Zhao (2013, JFE)、Bakshi–Madan–Panayotov (2010)", "日経ではコール側の裾の需要（10/20 出来高上位 75,000・80,000 円）が右側の上昇に寄与しうる。"],
         ["需給・市場構造", "ディーラーの在庫制約により OTM の需要圧力が IV を押し上げ、f_Q が f_P から乖離する。", "Bollen–Whaley (2004, JF)、Gârleanu–Pedersen–Poteshman (2009, RFS)", "建玉・出来高の集中（§4）と併せて検証可能。"]],
        [20, 60, 48, 48]),
    Pg("表3  パズルの説明仮説。", CAP),
    CondPageBreak(50 * mm),
    Pg("2.5 2 状態変数カーネルと当方の γ の関係", H2),
    Pg("第二状態変数の系統を式で書く。満期時点の分散 σ_T² を追加の状態変数とし、カーネルを"),
    EQ(r"M \propto S_T^{-\gamma}\,\exp\left(\xi\,(\sigma_T^2 - \mathrm{E}[\sigma_T^2])\right)"),
    Pg("とする。ξ > 0 は「分散が高い状態の 1 円を高く買う」＝負の分散プレミアムである。これを x だけの関数に射影する（x を与えたときの期待値を取る）と、"),
    EQ(r"\ln M(x) \approx c - \gamma\,x + \xi\,\mathrm{E}[\sigma_T^2 \mid x] \qquad (\mathrm{式\,4})"),
    Pg("x を与えたときの期待分散 E[σ_T² | x] は、大きな変動（|x| が大）のときに高いので x の凸関数である。従って式 4 の第 3 項が U 字を作り、曲率は ξ × E[σ_T² | x] の曲率になる。"
       "さらにレバレッジ効果（下落時に分散が上がる）があれば E[σ_T² | x] は x = 0 で右下がりの傾きを持ち、式 4 の x = 0 における傾きは −γ + ξ × (負の値) となる。すなわち<b>一次式で推定した γ̂ は真の γ と ξ の混合</b>であり、この分離には (i) ボラティリティで条件付ける（Song–Xiu は VIX オプションを使った）、(ii) γ と ξ を実現値で同時推定する、のいずれかが要る。日経では VI 先物はあるが VI オプションは無く、(ii) すなわち方法 III の蓄積が現実的である。"),
    Pg("2.6 当方の推定方法への含意", H2),
    Pg("・方法 I（直線）: 式 3 の通り「期待リターン差 ÷ 分散」であり、f_P の平均の仮定に支配される。単独で水準を論じない。<br/>"
       "・方法 I（2 次式）: 傾き（ATM での局所 γ）と曲率 β₂（分散プレミアム）に分解して記録する。曲率は IV−RV の別表現であり、独立の情報ではないが「市場が分散にいくら払っているか」を密度の言葉で記録した値になる。<br/>"
       "・方法 II（モーメント法）: 歪度と尖度の差から γ を出す。式 4 の ξ 項は歪度にも入るため同じ混合の問題を持つ。<br/>"
       "・方法 III（最尤法）: 実現 SQ 値で f_P を直接評価するため f_P の仮定が不要で、U 字のもとでも「予測密度として最良のべき近似」を与える。格子は −6〜12 に拡張済み、週次満期（ミニ）を取り込み年 100 点規模にして初めて区間が縮む。"),
    PageBreak(),
    Pg("3. γ は状態変数の関数か：関連研究", H1),
    Pg("3.1 理論：なぜ γ が動くのか", H2),
    Pg("Campbell–Cochrane (1999, JPE) の習慣形成モデルでは、効用が消費 C そのものではなく「習慣水準 X を超えた消費」C − X に依存する。余剰消費比率 S_t = (C_t − X_t)/C_t を定義すると、相対的リスク回避度は"),
    EQ(r"\mathrm{RRA}_t = \frac{\gamma}{S_t}"),
    Pg("となり、消費が習慣水準に近づく不況期（S_t が小）にリスク回避度が上がる。株価下落後に γ が上がる「反景気循環的なリスク回避度」はこのモデルの予測である。"
       "Barberis–Huang–Santos (2001, QJE) はプロスペクト理論で逆の経路（直前の利益の後はリスクを取りやすい「ハウスマネー効果」）も示しており、過去リターンの符号で γ が動く根拠は双方向にある。"),
    Pg("3.2 実証研究の整理", H2),
    TBL([["研究", "データ・方法", "γ を動かす変数と結果"],
         ["Rosenberg–Engle (2002, JFE)", "S&P 500 オプション、月次で経験的プライシングカーネルを推定（1991–95）", "γ は反景気循環。信用スプレッド拡大と正、イールドカーブのスティープ化と負の相関。時変カーネルのヘッジ比率が時不変を上回る。"],
         ["Bliss–Panigirtzoglou (2004, JF)", "S&P 500・FTSE 100、f_Q をべき効用で f_P に変換し予測力を最大化（1983–2001）", "γ は予測期間で減少（1 週間で約 8、6 週間で約 2）。ボラティリティが高いほど γ は低い。"],
         ["Brandt–Wang (2003, JME)", "消費と債券データ、習慣形成の拡張", "γ はインフレ・ニュースにも反応する。"],
         ["Bollerslev–Gibson–Zhou (2011, J. Econometrics)", "S&P 500 の IV と高頻度 RV から分散プレミアムを推定", "分散プレミアムから逆算した時変 γ_t はマクロ指標（生産、P/E、信用スプレッド）と連動。"],
         ["Christoffersen–Heston–Jacobs (2013, RFS)", "Heston 型 GARCH と分散依存カーネル", "カーネルの分散依存（ξ）が U 字、インプライド分散のプレミアム、OTM プットの割高を同時に説明。"],
         ["Song–Xiu (2016, J. Econometrics)", "S&P 500 オプションと VIX オプション、ノンパラメトリック", "VIX で条件付けると株価リターンに対するカーネルは単調減少、無条件では U 字。ボラティリティに対するカーネルは U 字。"],
         ["Kiesel–Rahe (2017, JBF)", "Heston 型モデルで γ を VIX の関数として推定", "γ は VIX の水準で変動、時変 γ でリスク予測が改善。"],
         ["Linn–Shive–Shumway (2018, RFS)", "条件付き f_P（GARCH 系）で再推定", "無条件 f_P の使用が U 字の主因と主張。条件付けると単調。"],
         ["Cuesdeanu–Jackwerth (2018, Annals of Finance)", "展望論文", "推定法・説明仮説の体系的整理。"]],
        [40, 58, 78]),
    Pg("表4  γ の状態依存とプライシングカーネルに関する主要研究。", CAP),
    Pg("<b>共通する結論。</b>(1) γ は時間を通じて一定ではない。(2) 最も頑健な説明変数はボラティリティ（またはその代理である分散プレミアム）である。(3) ボラティリティを第二の状態変数として明示すると U 字は消え、残る「価格に対する γ」は正の単調減少に戻る。(4) 予測期間（残存日数）にも依存する。"),
    CondPageBreak(60 * mm),
    Pg("3.3 当方の蓄積設計への反映", H2),
    Pg("表 2 の回帰は f_P の構成（RV₂₀ でスケール）に由来する機械的な成分を含むため、状態依存の証拠としては弱い。実現値に基づく状態依存 γ の推定は、方法 III を次のように拡張する。満期 t ごとに予測時点の状態変数 z_t（IV−RV、20 日リターン、残存日数）を記録し、"),
    EQ(r"\gamma_t = \gamma_0 + \gamma_1\,(\mathrm{IV}-\mathrm{RV})_t + \gamma_2\,\mathrm{ret20}_t,\qquad \max_{\gamma_0,\gamma_1,\gamma_2}\ \sum_t \ln f_{P,t}(S_{T_t};\,\gamma_t)"),
    Pg("を最大化する。3 母数の識別には実現点が 50〜100 本必要で、月次（年 12 本）では不可能だが、ミニの週次満期（年 100 本前後）を取り込めば 1 年で到達する。現在の日次蓄積（gamma_daily.jsonl）には day, T, F, atm, rv20, ret20, 2 次式の傾きと曲率が入っており、gamma_mle.json の weekly 配列に実現値が加わり次第この回帰に移れる。"),
    Pg("記録しておくべき項目を整理する。"),
    TBL([["項目", "用途", "現状"],
         ["F, T, SVI 母数 p", "f_Q の再構成（任意の γ で傾け直せる）", "蓄積済み"],
         ["s_Q, s_P, IV−RV", "曲率（分散プレミアム）の分解、状態変数", "s_Q は再計算可、IV−RV は atm/rv20 から"],
         ["2 次式の傾き・曲率、直線 γ̂", "形状の日次記録", "蓄積済み"],
         ["実現 SQ 値（月次・週次）", "方法 III、状態依存 γ の最尤推定", "月次 5 本、週次は 10/23 以降に蓄積開始"],
         ["行使価格別の出来高・建玉", "需給仮説の検証（§4）", "気配スナップショットに含まれる"]],
        [48, 70, 58]),
    Pg("表5  蓄積項目。", CAP),
    CondPageBreak(70 * mm),
    Pg("4. 行使価格の出来高・建玉による重み付け", H1),
    Pg("4.1 重みが入りうる場所", H2),
    Pg("γ 推定の計算は 3 段階あり、重みは前 2 段階に入りうる。(i) SVI の当てはめ：各行使価格の IV の観測値を 2 乗誤差で当てはめる際の重み。現行は気配幅の逆 2 乗 1/(spread + 0.5%)²。(ii) カーネル回帰：ln(f_Q/f_P) を x の一次式（または 2 次式）に当てはめる際の重み。現行は f_Q の確率質量。(iii) 最尤法：実現値 1 点の密度を評価するだけなので、行使価格の重みは入らない。"),
    Pg("4.2 理論的な位置づけ", H2),
    Pg("M = f_Q/f_P は 2 つの密度の比であり、定義のどこにも出来高や建玉は現れない。従って「出来高の多い行使価格ほど情報が多いから重く見る」という発想に、母集団の量としての根拠はない。重みが正当化されるのは推定量の性質を通じてのみで、次の 2 つに分かれる。"),
    Pg("<b>(a) 正しく特定されている場合（ln M が本当に一次式）。</b>重みは推定の分散だけに影響し、最適重みは測定誤差の分散の逆数である（加重最小 2 乗の定理）。測定誤差の源は気配幅と離散的な刻みなので、気配幅ベースの現行重みが理論に沿う。出来高・建玉は「流動性が高いほど気配が真の価格に近い」という意味で測定精度の代理になりうるが、気配幅を直接使う方が近い。"),
    Pg("<b>(b) 特定を誤っている場合（ln M が U 字）。</b>一次式の傾きは、局所的な傾き γ(x) を重みで平均したものになる。"),
    EQ(r"\hat\gamma_{\mathrm{lin}} \approx \frac{\sum_i w_i\,(x_i-\bar x)\,\gamma(x_i)\,(x_i - \bar x)}{\sum_i w_i\,(x_i-\bar x)^2}"),
    Pg("U 字では γ(x) が左で正、右で負なので、どの x を重く見るかで γ̂ の値が変わる。これが重み依存の正体であり、重みを工夫して解決する問題ではなく、一次式という特定が原因である。出来高で重み付けた γ̂ は「取引が集中する行使価格での局所 γ」という別の推定対象を指す。"),
    Pg("4.3 実験結果", H2),
    Pg(f"{D} 引け（11月限、OTM 両気配 212 本、当日出来高 {re.search(r'出来高合計=([\d,]+)', GW).group(1)} 枚、建玉 {re.search(r'建玉合計=([\d,]+)', GW).group(1)} 枚）で、(i)(ii) の重みを入れ替えた結果を表 6 に示す。出来高上位の行使価格は {vol_top} で、45,000・75,000・80,000 円は中央 5〜95% の推定区間の外にある。"),
    fig_weights(),
    TBL([["設定（SVI の重み / 回帰の重み）", "γ̂（直線）", "γ̂（2 次式の 1 次項）", "曲率 β₂"]] + [[lab, f"{g1:+.3f}", f"{g2:+.3f}", f"{q:+.1f}"] for lab, g1, g2, q in wrows], [82, 30, 36, 24]),
    Pg(f"表6  重み付けの比較。SVI の重みを出来高・建玉で変えたときの ATM IV は 現行 {atm_line[0]}% → 出来高 {atm_line[1]}% → 建玉 {atm_line[2]}%。", CAP),
    Pg("読み取れることは 3 点である。(1) SVI 側の重み変更は ATM IV を 0.1%pt しか動かさず、密度の形もほぼ変わらない（曲率 14.4 → 15.1）。気配幅の重みが既に流動性を反映しているためである。(2) 回帰側の重み変更は直線の γ̂ を −0.64 から +0.04 まで動かす。出来高・建玉は OTM（特にプット側 55,000〜62,000 円）に偏っており、4.2(b) の通り U 字の左側（局所 γ が正の領域）を重く見るため γ̂ が上がる。(3) 2 次式の傾き（−0.82〜−0.97）と曲率（+12.9〜+15.1）は重みにほとんど依らない。U 字を 2 次式で捉えれば重み依存は消える。"),
    Pg("4.4 推奨", H2),
    Pg("・SVI の重みは気配幅ベースを維持する（測定誤差の理論に沿い、結果も変わらない）。<br/>"
       "・日次記録の主指標は 2 次式の傾き（ATM での局所 γ）と曲率 β₂ とし、直線 γ̂ は参考値に落とす。<br/>"
       "・出来高重みの直線 γ̂ は「取引の中心での局所 γ」として補助列に加える価値はあるが、建玉重みは古い建玉が残るだけで情報の流れを表さないため採用しない。<br/>"
       "・重み付けで解けない部分（γ と ξ の分離、水準の決定）は、式 4 の 2 母数カーネルを週次満期の実現値で最尤推定することで対処する。"),
    sp(2),
    Pg("5. 結論", H1),
    Pg("当方が日次で観測している ln(f_Q/f_P) の U 字はプライシングカーネル・パズルそのものであり、その曲率は IV − RV（分散プレミアム）で説明できる。従って「市場が織り込む γ」を 1 つの数字で表すことは特定の誤りであり、(傾き, 曲率) の 2 次元で記録し、水準は実現 SQ 値による最尤法で決める、という分担に改める。文献はいずれも γ の状態依存（ボラティリティ、過去リターン、残存期間）を支持しており、週次満期の実現値が溜まり次第、γ_t = γ₀ + γ₁(IV−RV) + γ₂ ret20 の最尤推定に進む。出来高・建玉の重み付けは γ の定義に根拠を持たず、推定精度の代理としての限定的な使い方に留める。"),
    Pg("参考文献", H2),
    Pg("Aït-Sahalia, Y., Lo, A. W. (2000) Nonparametric risk management and implied risk aversion. J. Econometrics 94. / Bakshi, G., Madan, D., Panayotov, G. (2010) Returns of claims on the upside and the viability of U-shaped pricing kernels. JFE 97. / Barberis, N., Huang, M., Santos, T. (2001) Prospect theory and asset prices. QJE 116. / "
       "Beare, B., Schmidt, L. (2016) An empirical test of pricing kernel monotonicity. J. Applied Econometrics 31. / Bliss, R. R., Panigirtzoglou, N. (2004) Option-implied risk aversion estimates. JF 59. / Bollen, N., Whaley, R. (2004) Does net buying pressure affect the shape of implied volatility functions? JF 59. / "
       "Bollerslev, T., Gibson, M., Zhou, H. (2011) Dynamic estimation of volatility risk premia and investor risk aversion from option-implied and realized volatilities. J. Econometrics 160. / Brandt, M., Wang, K. (2003) Time-varying risk aversion and unexpected inflation. JME 50. / Campbell, J. Y., Cochrane, J. H. (1999) By force of habit. JPE 107. / "
       "Christoffersen, P., Heston, S., Jacobs, K. (2013) Capturing option anomalies with a variance-dependent pricing kernel. RFS 26. / Cuesdeanu, H., Jackwerth, J. C. (2018) The pricing kernel puzzle: survey and outlook. Annals of Finance 14. / Gârleanu, N., Pedersen, L. H., Poteshman, A. (2009) Demand-based option pricing. RFS 22. / "
       "Jackwerth, J. C. (2000) Recovering risk aversion from option prices and realized returns. RFS 13. / Kiesel, R., Rahe, F. (2017) Option pricing under time-varying risk-aversion with applications to risk forecasting. JBF 76. / Linn, M., Shive, S., Shumway, T. (2018) Pricing kernel monotonicity and conditional information. RFS 31. / "
       "Polkovnichenko, V., Zhao, F. (2013) Probability weighting functions implied in options prices. JFE 107. / Rosenberg, J., Engle, R. (2002) Empirical pricing kernels. JFE 64. / Song, Z., Xiu, D. (2016) A tale of two option markets: Pricing kernels and volatility risk. J. Econometrics 190.", SMALL),
]
doc.build(story, onFirstPage=make_footer(f"プライシングカーネル・パズルと状態依存 γ（データ 〜{D} 引け、作成 2026-10-21）"), onLaterPages=make_footer(f"プライシングカーネル・パズルと状態依存 γ（データ 〜{D} 引け、作成 2026-10-21）"))
print("written", PDF_PATH)
