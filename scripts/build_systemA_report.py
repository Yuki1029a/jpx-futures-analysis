"""系統A（オプション価格から先物の先行き・投資家の見通しを抽出する視点 A1〜A9）の詳細解説資料（PDF）。

入力: scripts/outlook_a_calc.py の出力（a_calc.json ほか）、%TMP%/rnd.json・rnd.npz（10/16 引けの密度）、%TMP%/rnd_series_2611.json
使い方: python scripts/build_systemA_report.py <計算結果ディレクトリ> <出力フォルダ>
"""
from __future__ import annotations

import json
import os
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import CondPageBreak, PageBreak, SimpleDocTemplate
from scipy.stats import norm

from report_kit import A, B, BODY, CAP, EQ, GRID, H1, H2, INK, MUT, O, SMALL, SUB, TBL, TITLE, Y, Pg, fig_to_img, fmt, make_footer, sp, style_ax
from rnd_lib import svi

CALC = Path(sys.argv[1]); OUT_DIR = Path(sys.argv[2]); OUT_DIR.mkdir(parents=True, exist_ok=True); TMP = Path(os.environ["TMP"])
PDF_PATH = OUT_DIR / "nk225op_systemA_option_implied_outlook_guide.pdf"
R = json.load(open(CALC / "a_calc.json")); J = json.load(open(TMP / "rnd.json")); SER = json.load(open(TMP / "rnd_series_2611.json"))
A3 = pd.read_json(CALC / "a3_series.json"); A3["date"] = pd.to_datetime(A3.day.astype(str))
TL = np.load(CALC / "a2_tilt.npz"); PIT = np.load(CALC / "a6_pit.npz"); GX = np.load(CALC / "a9_gex.npz")
F16, T16, s_atm = J["F"], J["T"], J["s_atm"]; a1, a2, a3, a5, a6, a7, a8, a9 = R["a1"], R["a2"], R["a3"], R["a5"], R["a6"], R["a7"], R["a8"], R["a9"]
pct = lambda v: f"{v*100:.1f}%"


# ------------------------------------------------------------------ 図
def fig_qp():
    K = TL["K"]; sel = (K >= 50000) & (K <= 80000)
    fig, a = plt.subplots(figsize=(7.2, 3.4))
    a.plot(K[sel], TL["fQ"][sel] * 1e5, color=B, lw=2, label="リスク中立 f_Q（γ = 0）")
    for g_, col in ((2, A), (4, O), (6, Y)):
        a.plot(K[sel], TL[f"f{g_}"][sel] * 1e5, color=col, lw=1.6, ls="--", label=f"実確率の近似 f_P（γ = {g_}）")
    a.axvline(F16, color=INK, lw=0.8)
    style_ax(a, "図1  べき効用でリスク中立密度を傾けた実確率密度（11月限、10/16 引け）", "SQ値（円）", "確率密度（%/1,000円）"); a.legend(fontsize=8, frameon=False)
    return fig_to_img(fig)


def fig_smiles():
    fig, a = plt.subplots(figsize=(7.2, 3.3)); xs = np.linspace(-0.25, 0.2, 300)
    for r, col in zip([s for s in SER if s["label"] in ("10/09 引け", "10/15 引け", "10/16 引け")], (MUT, O, B)):
        p = np.array(r["p"]); a.plot(np.exp(xs) * 100 - 100, np.sqrt(svi(p, xs)[0] / r["T"]) * 100, color=col, lw=2, label=f"{r['label']}（F {r['F']:,.0f}）")
    a.axvline(0, color=INK, lw=0.8)
    style_ax(a, "図2  スマイルの形の変化（横軸 = 行使価格 ÷ F − 1）", "F からの乖離（%）", "IV（%）"); a.legend(fontsize=8, frameon=False)
    return fig_to_img(fig)


def fig_vrp():
    fig, ax = plt.subplots(2, 1, figsize=(7.2, 6.0), sharex=True); d = A3.date
    ax[0].plot(d, A3.mfiv * 100, color=B, lw=1.8, label="モデルフリー IV（VIX 型、期近）"); ax[0].plot(d, A3.atm * 100, color=A, lw=1.4, label="ATM IV")
    ax[0].plot(d, A3.rv20 * 100, color=O, lw=1.8, label="実現ボラティリティ（直近 20 日）"); ax[0].plot(d, A3.fwd20 * 100, color=MUT, lw=1.2, ls="--", label="その後 20 日の実現ボラティリティ")
    style_ax(ax[0], "図3  インプライドと実現のボラティリティ（年率 %、2026-06〜10）", "", "%"); ax[0].legend(fontsize=7.5, frameon=False, ncol=2)
    ax[1].plot(d, A3.vrp_mf * 1e4, color=B, lw=1.8, label="VRP（事前）= MFIV² − RV₂₀²"); ax[1].plot(d, A3.vrp_expost * 1e4, color=MUT, lw=1.2, ls="--", label="VRP（事後）= MFIV² − その後 20 日の RV²")
    ax[1].axhline(0, color=MUT, lw=0.8); style_ax(ax[1], "", "", "分散（%²）"); ax[1].legend(fontsize=7.5, frameon=False); fig.autofmt_xdate()
    return fig_to_img(fig)


def fig_rr():
    df = pd.DataFrame([{"label": r["label"], "F": r["F"], "rr": r["rr25_delta"], "atm": r["atm"] * 100} for r in SER])
    fig, ax = plt.subplots(2, 1, figsize=(7.2, 4.6), sharex=True); xi = np.arange(len(df))
    ax[0].plot(xi, df.F, color=INK, lw=1.6, marker="s", ms=4); style_ax(ax[0], "図4  先物推定値と RR25 の推移（11月限）", "", "F（円）")
    ax[1].plot(xi, df.rr, color=O, lw=2, marker="o", ms=4, label="RR25（%pt）"); ax[1].plot(xi, df.atm, color=B, lw=1.4, marker="o", ms=3, label="ATM IV（%）")
    ax[1].set_xticks(xi); ax[1].set_xticklabels(df.label.str.replace(" ", "\n"), fontsize=8); style_ax(ax[1], "", "", "%"); ax[1].legend(fontsize=8, frameon=False)
    return fig_to_img(fig)


def fig_tail():
    labs = ["−10% 未満", "−5% 未満", "+5% 超", "+7% 超"]; keys = ["p_lt_m10", "p_lt_m5", "p_gt_p5", "p_gt_p7"]
    q = [a5["q"][k] * 100 for k in keys]; hu = [a5["hist"][k] * 100 for k in keys]; hh = [a5["hist"]["hi"].get(k, np.nan) * 100 for k in keys]
    fig, a = plt.subplots(figsize=(7.2, 3.0)); x = np.arange(4); w = 0.26
    a.bar(x - w, q, w, color=B, label="リスク中立（10/16 引け、28 日先）"); a.bar(x, hu, w, color=MUT, label="実績頻度 2000 年〜（20 営業日）"); a.bar(x + w, hh, w, color=O, label="実績頻度、開始時 RV₂₀ ≥ 25% のみ")
    for i, (q_, u_, h_) in enumerate(zip(q, hu, hh)):
        a.text(i - w, q_ + 0.5, f"{q_:.1f}", ha="center", fontsize=7.5); a.text(i, u_ + 0.5, f"{u_:.1f}", ha="center", fontsize=7.5)
        if not np.isnan(h_):
            a.text(i + w, h_ + 0.5, f"{h_:.1f}", ha="center", fontsize=7.5)
    a.set_xticks(x); a.set_xticklabels(labs); style_ax(a, "図5  裾の確率：リスク中立と歴史的頻度（%）", "28 日間の変化率", "確率（%）"); a.legend(fontsize=7.5, frameon=False)
    return fig_to_img(fig)


def fig_pit():
    K, f = PIT["K"], PIT["f"]; sel = (K >= 50000) & (K <= 80000); r0 = a6["rows"][0]
    fig, a = plt.subplots(figsize=(7.2, 3.0)); a.plot(K[sel], f[sel] * 1e5, color=B, lw=2, label=f"10月限の分布（{r0['day'][4:6]}/{r0['day'][6:]} 引け、残存 {r0['T_days']:.0f} 日）")
    a.axvline(a6["sq"], color=O, lw=1.6, ls="--"); a.text(a6["sq"], (f[sel] * 1e5).max() * 0.9, f" 実現 SQ {a6['sq']:,.0f}\n u = {r0['u']:.2f}", fontsize=8.5, color=O)
    a.axvline(r0["F"], color=INK, lw=0.8); style_ax(a, "図6  校正の例：予測時の分布と実現した SQ 値", "SQ値（円）", "確率密度（%/1,000円）"); a.legend(fontsize=8, frameon=False, loc="upper left")
    return fig_to_img(fig)


def fig_term():
    rows = a7["rows"]; Td = np.array([r["T"] * 365 for r in rows]); atm = np.array([r["atm"] * 100 for r in rows]); rr = np.array([r["rr25"] for r in rows])
    fig, ax = plt.subplots(1, 2, figsize=(7.2, 2.9))
    ax[0].plot(Td, atm, color=B, lw=2, marker="o"); [ax[0].text(t, v + 0.15, f"{v:.1f}", ha="center", fontsize=8) for t, v in zip(Td, atm)]
    style_ax(ax[0], "図7a  ATM IV の期間構造", "残存日数", "IV（%）")
    ax[1].plot(Td, rr, color=O, lw=2, marker="o", label="RR25 実測"); ax[1].plot(Td, rr[0] * np.sqrt(Td[0] / Td), color=MUT, lw=1.2, ls="--", label="1/√T で減衰した場合")
    style_ax(ax[1], "図7b  RR25 の期間構造", "残存日数", "%pt"); ax[1].legend(fontsize=7.5, frameon=False)
    fig.tight_layout()
    return fig_to_img(fig)


def fig_gex():
    k, v = GX["k"], GX["v"]; sel = (k >= 55000) & (k <= 80000)
    fig, ax = plt.subplots(2, 1, figsize=(7.2, 5.0), sharex=True)
    ax[0].bar(k[sel], v[sel] / 1e8, width=400, color=[B if x >= 0 else O for x in v[sel]]); ax[0].axhline(0, color=MUT, lw=0.8); ax[0].axvline(F16, color=INK, lw=0.8)
    style_ax(ax[0], "図8  行使価格別 GEX と累積（億円/1%、慣例の符号）", "", "GEX（億円/1%）")
    ax[1].plot(k, np.cumsum(v) / 1e8, color=INK, lw=1.8); ax[1].axhline(0, color=MUT, lw=0.8); ax[1].set_xlim(55000, 80000); ax[1].axvline(F16, color=INK, lw=0.8)
    style_ax(ax[1], "", "行使価格（円）", "累積 GEX（億円/1%）")
    return fig_to_img(fig)


# ------------------------------------------------------------------ 文書
S = []
S += [sp(8), Pg("オプション価格から先物の先行きと<br/>投資家の見通しを抽出する方法", TITLE), Pg("系統 A（A1〜A9）の詳細解説：理論・式・数値例・実装", SUB),
      TBL([["項目", "内容"],
           ["位置づけ", "前稿「推移と見通し抽出の視点」第 II 部 6.1 節（系統 A：オプション価格から）の各項目を、独立に読める形で展開したもの"],
           ["数値例の基準", f"日経225オプション 2026年11月限、10/16 引け（QRI {J['ts'][:16]}）。F = {fmt(F16)} 円、残存 {T16*365:.1f} 日。日次系列は 2026-06-03〜10-16"],
           ["前提知識", "リスク中立分布の推定手順（SVI、Breeden–Litzenberger）は前々稿「SQ値の織り込みに関する確率分析」第 3〜4 章。本稿は結果を用いるが、必要な定義は再掲する"],
           ["作成", f"{date.today():%Y-%m-%d}、先物手口分析ワークスペース（scripts/outlook_a_calc.py, scripts/build_systemA_report.py）"]], [34, 144]),
      sp(5), Pg("構成", H2),
      TBL([["章", "内容", "答える問い"],
           ["0", "共通の土台：記号、リスク中立確率と実確率、三つのプレミアム", "なぜ価格から読んだ確率は「見通し」そのものではないのか"],
           ["A1", "リスク中立分布の形状の時系列", "分布の形から何が読めるか。モデルフリー・モーメントとは何か"],
           ["A2", "実確率への変換", "保険料を取り除いて実確率に近づけるには"],
           ["A3", "分散リスク・プレミアム", "IV と実現ボラの差は何を表し、何を予測するか"],
           ["A4", "スキューとリスクリバーサル", "プットの割高さは何で決まり、リターンと関係するか"],
           ["A5", "テール（ジャンプ）の確率", "大幅安の確率はどこまで「恐怖」でどこまで「予想」か"],
           ["A6", "校正検証", "分布が正しかったかをどう検定するか"],
           ["A7", "期間構造", "限月間の差からイベントと期待の時間配分を読む"],
           ["A8", "オプションのフロー", "出来高・建玉の変化から情報を持つ投資家の動きを読む"],
           ["A9", "ディーラーのガンマ", "見通しではなく値動きの経路を読む"],
           ["10", "統合・実装・参考文献", "どの順で実装し、どう検証するか"]], [12, 90, 76]),
      PageBreak()]

# ================================================================== 0
S += [Pg("0. 共通の土台", H1), Pg("0.1 記号", H2),
      TBL([["記号", "意味", "備考"],
           ["S_T", "満期 T の SQ値（確率変数）", "11月限は 2026-11-13 算出"],
           ["F", "満期 T で決済される先物の価格（リスク中立分布の平均）", f"パリティ推定、{fmt(F16)} 円"],
           ["R = ln(S_T / F)", "満期までの対数リターン", "モーメントの計算対象"],
           ["x = ln(K / F)", "行使価格 K の対数マネーネス", ""],
           ["σ(K), σ_ATM", "行使価格 K の IV、ATM（K = F）の IV", f"σ_ATM = {s_atm*100:.2f}%"],
           ["f_Q, Φ_Q", "リスク中立の密度と累積分布", "オプション価格から一意に決まる"],
           ["f_P, Φ_P", "実確率（物理測度）の密度と累積分布", "直接は観測できない"],
           ["M(S)", "確率的割引因子（プライシング・カーネル）", "f_Q = f_P · M / E_P[M]"],
           ["γ", "相対的リスク回避度", "べき効用 U(S) = S^(1−γ)/(1−γ)"],
           ["RV", "実現ボラティリティ（年率）", "日次対数リターンの標準偏差 × √250"],
           ["IV_MF", "モデルフリー・インプライド・ボラティリティ", "VIX 型の計算（A3）"],
           ["Δ, Γ, Vega", "デルタ、ガンマ、ベガ", "ブラック式"]], [40, 80, 58]),
      Pg("0.2 リスク中立確率と実確率", H2),
      Pg("資産の価格は、将来の受取 X を「確率で重み付けした期待値」そのものではない。投資家は損失が出る状態での受取を高く評価するため、価格は"),
      EQ(r"\mathrm{Price} = E_{P}[M\cdot X] = E_{Q}[X]"),
      Pg("と書ける。M は状態ごとの評価の重み（確率的割引因子）、E_P は実確率での期待値である。右辺は、実確率を M で歪めた確率 Q（リスク中立確率）での単純な期待値として同じ価格を表す。両者の関係は"),
      EQ(r"f_{Q}(S) = \frac{M(S)\, f_{P}(S)}{E_{P}[M]}"),
      Pg("である。株価指数では、下落時（S が小さい）ほど M が大きい（投資家が困る状態ほど重みが大きい）。したがって f_Q は f_P より下落側に厚く、上昇側に薄い。"
         "オプション価格から取り出せるのは f_Q であり、投資家の見通し（f_P）を知るには M、すなわち「保険料」を取り除く必要がある。"),
      Pg("0.3 三つのプレミアム", H2),
      TBL([["名称", "定義", "現れ方", "扱う章"],
           ["分散リスク・プレミアム", "IV² − E_P[RV²]。分散の価格と期待値の差", "全行使価格の IV が実現ボラより高い", "A3"],
           ["スキュー・プレミアム", "下落側の保険（低い行使価格のプット）に上乗せされる対価", "RR25 が正、f_Q の左裾が f_P より厚い", "A2, A4"],
           ["ジャンプ（テール）・プレミアム", "急落リスクへの対価。分散プレミアムのうち裾に帰属する部分", "深い OTM プットの価格が歴史的頻度より高い", "A5"]], [40, 60, 54, 24]),
      Pg("これらは互いに重なり合い、価格だけでは完全には分離できない。本稿の各章は、どの部分をどの仮定で切り出すかの違いとして読める。"),
      PageBreak()]

# ================================================================== A1
S += [Pg("A1. リスク中立分布の形状の時系列", H1), Pg("A1.1 何を測るか、なぜ形に情報があるか", H2),
      Pg("リスク中立分布の平均は無裁定によって F に固定される。したがって「市場がどこを見ているか」は、平均ではなく、幅（標準偏差）、非対称（歪度）、裾の厚さ（尖度）、および下落側と上昇側の保険の相対価格（RR25）に現れる。"
         "これらを日次で追うと、(a) 水準の変化を伴わない「恐怖の増減」、(b) 下落と上昇のどちらに保険需要が偏っているか、(c) 分布の幅が実現ボラと比べて過大か、が分かる。"),
      Pg("A1.2 密度から計算するモーメント", H2),
      Pg("推定した密度 f_Q（前々稿）から、対数リターン R = ln(S_T/F) のモーメントを数値積分で求める。"),
      EQ(r"m = \int R\, f_{Q}\, dS,\quad s^{2} = \int (R-m)^{2} f_{Q}\, dS,\quad \mathrm{Skew} = \frac{\int (R-m)^{3} f_{Q}\, dS}{s^{3}},\quad \mathrm{Kurt} = \frac{\int (R-m)^{4} f_{Q}\, dS}{s^{4}} - 3"),
      Pg(f"<b>数値例（10/16 引け）。</b>対数リターンの歪度 {a1['skew_x']:+.2f}、超過尖度 {a1['kurt_x']:+.2f}（積分範囲 x ∈ [−0.6, 0.6]）。SQ値そのもの（水準）の歪度は {a1['skew_level']:+.2f}、超過尖度 {a1['kurt_level']:+.2f} で、対数を取ると左裾がより強調される。"),
      Pg("A1.3 モデルフリー・モーメント（Bakshi–Kapadia–Madan 2003）", H2),
      Pg("密度を経由せず、OTM オプション価格の積分として対数リターンのモーメントを直接計算する方法がある。任意の受取関数 H(S) の価格は、OTM オプションの価格 C(K)、P(K) を H の 2 階微分で重み付けした積分で複製できる（Carr–Madan のスパニング）。R²、R³、R⁴ を受取とする契約の価格を V、W、X とすると、"),
      EQ(r"V = \int_{F}^{\infty} \frac{2\left(1-\ln\frac{K}{F}\right)}{K^{2}} C(K)\,dK + \int_{0}^{F} \frac{2\left(1+\ln\frac{F}{K}\right)}{K^{2}} P(K)\,dK"),
      EQ(r"W = \int_{F}^{\infty} \frac{6\ln\frac{K}{F} - 3\left(\ln\frac{K}{F}\right)^{2}}{K^{2}} C(K)\,dK - \int_{0}^{F} \frac{6\ln\frac{F}{K} + 3\left(\ln\frac{F}{K}\right)^{2}}{K^{2}} P(K)\,dK"),
      EQ(r"X = \int_{F}^{\infty} \frac{12\left(\ln\frac{K}{F}\right)^{2} - 4\left(\ln\frac{K}{F}\right)^{3}}{K^{2}} C(K)\,dK + \int_{0}^{F} \frac{12\left(\ln\frac{F}{K}\right)^{2} + 4\left(\ln\frac{F}{K}\right)^{3}}{K^{2}} P(K)\,dK"),
      Pg("金利ゼロのとき μ = −V/2 − W/6 − X/24 を対数リターンの平均の近似として、"),
      EQ(r"\mathrm{Skew}_{Q} = \frac{W - 3\mu V + 2\mu^{3}}{(V-\mu^{2})^{3/2}},\qquad \mathrm{Kurt}_{Q} = \frac{X - 4\mu W + 6\mu^{2}V - 3\mu^{4}}{(V-\mu^{2})^{2}}"),
      Pg("となる。利点は、スマイルの関数形（SVI など）を仮定しないことである。欠点は、積分が深い OTM（気配の無い領域）まで及ぶため、両翼をどう補外するかに結果が強く依存することである（Jiang and Tian 2005 が打ち切り誤差を分析）。"),
      Pg(f"<b>数値例。</b>SVI の価格を使い、積分範囲を変えて計算すると、歪度は x ∈ [−0.6, 0.6] で −2.11、[−0.8, 0.8] で {a1['skew_bkm']:.2f}、気配のある範囲（22,500〜85,000 円）で −2.54 と変わる。尖度は 13.7〜20.8。"
         f"密度の数値積分（A1.2、{a1['skew_x']:+.2f}）より大きく出るのは、1/K² の重みが深い OTM プットの価格を強く拾うためである。"
         "したがって、モデルフリー・モーメントは「水準」よりも「同じ規約で計算した系列の変化」として使う。"),
      CondPageBreak(110 * mm), Pg("A1.4 形状の時系列と読み方", H2), fig_smiles(),
      TBL([["時点", "F", "ATM IV", "RR25", "歪度（水準）", "超過尖度", "5% 点", "95% 点"]] +
          [[r["label"], fmt(r["F"]), f"{r['atm']*100:.1f}", f"{r['rr25_delta']:+.1f}", f"{r['skew']:+.2f}", f"{r['kurt']:+.2f}", fmt(r["q"]["0.05"]), fmt(r["q"]["0.95"])] for r in SER], [26, 20, 20, 18, 26, 22, 22, 22]),
      Pg("表 A1  形状指標の日次系列（11月限）。", CAP),
      Pg("図2 は 10/9、10/15（急落日）、10/16 のスマイルを F からの乖離で並べたものである。急落日はスマイル全体が持ち上がり（分散プレミアムの拡大）、翌日は水準が戻る一方で左側の傾きが残った（スキューの拡大）。"
         "形状指標は、この二つの動き（全体の上下＝幅、傾き＝非対称）を分けて記録する。"),
      Pg("<b>文献の示唆。</b>日経225 では Shiratsuka (2001) が 1989〜96 年のデータで、リスク中立分布は将来の価格変化について情報を持つが、歴史的分布を安定して上回る予測力は無く、結果が標本期間に依存すると報告した。"
         "米国では Birru and Figlewski (2012) が 2008 年の危機時に、分布の形が市場の下落と同時に歪み、落ち着くと速やかに戻ることを示した。形状は「予測」というより「市場が今どこに保険を掛けているか」の記述として用いるのが安全である。"),
      Pg("<b>実装。</b>日次で (F, σ_ATM, RR25, BF25, 歪度, 尖度, 5%/95% 点, 当てはめ誤差) を保存し、各指標を直近 1 年の分布に対する分位（または z 値）に直して監視する。急落日の値は当てはめ誤差が大きいので、誤差 1%pt 超の日は印を付ける。"),
      PageBreak()]

# ================================================================== A2
t0, t2, t4, t6 = a2["tilt"]["0"], a2["tilt"]["2"], a2["tilt"]["4"], a2["tilt"]["6"]
S += [Pg("A2. 実確率への変換", H1), Pg("A2.1 なぜ必要か", H2),
      Pg("0.2 節のとおり、f_Q は保険料込みの分布である。「投資家が実際にどの確率で何を見込んでいるか」に近づけるには、M を推定して f_P = f_Q · E[M] / M を計算する。M の形をどう置くかで三つの方法に分かれる。"),
      Pg("A2.2 方法 1：べき効用（パラメトリック）", H2),
      Pg("代表的投資家が富 S についてべき効用 U(S) = S^(1−γ)/(1−γ) を持つと仮定する。限界効用は U′(S) = S^(−γ) で、M はこれに比例する。したがって"),
      EQ(r"f_{P}(S) = \frac{f_{Q}(S)\, S^{\gamma}}{\int f_{Q}(u)\, u^{\gamma}\, du}"),
      Pg("となる。γ が大きいほど下落状態の重みが大きく評価されていたことになり、取り除いたときの f_P は右へ寄る。対数正規の場合は閉じた形で、平均が exp(γσ²T) 倍になる。"),
      Pg(f"<b>数値例（10/16 引け、図1）。</b>対数正規の近似では γ = 4 で平均は exp(4 × {s_atm:.3f}² × {T16:.4f}) = {a2['lognormal_shift']['4']:.4f} 倍、すなわち約 {(a2['lognormal_shift']['4']-1)*F16:+,.0f} 円の上方シフトになる。推定した密度を実際に傾けた結果は表 A2 のとおりで、γ = 4 で平均が {t4['mean']-t0['mean']:+,.0f} 円、P(S_T < 65,000) が {pct(t0['p_lt65'])} から {pct(t4['p_lt65'])} へ、P(S_T ≥ 70,000) が {pct(t0['p_ge70'])} から {pct(t4['p_ge70'])} へ動く。"),
      fig_qp(),
      TBL([["γ", "平均", "中央値", "最頻値", "5% 点", "95% 点", "P(<60,000)", "P(<65,000)", "P(≥70,000)"]] +
          [[g_, fmt(t["mean"]), fmt(t["median"]), fmt(t["mode"]), fmt(t["q05"]), fmt(t["q95"]), pct(t["p_lt60"]), pct(t["p_lt65"]), pct(t["p_ge70"])] for g_, t in (("0（Q）", t0), ("2", t2), ("4", t4), ("6", t6))], [18, 20, 20, 20, 20, 20, 20, 20, 20]),
      Pg("表 A2  リスク回避度 γ による実確率の近似（11月限、10/16 引け）。", CAP),
      Pg("<b>γ の決め方。</b>Bliss and Panigirtzoglou (2004) は、過去の分布予測と実現値から γ を最尤推定し、S&P 500 と FTSE 100 の 4 週間ホライズンで概ね 2〜8 の範囲、代表値として 4 前後を報告している。"
         "日本株についても同様の推定（Hitotsubashi Journal of Economics 2013 掲載の日経225 を用いた研究など）があり、値は期間と手法で大きく変わる。A6 の校正ログを蓄積すれば、自分のデータで γ を推定できる（対数尤度 Σ ln f_P(S_T,i; γ) を γ について最大化）。"),
      Pg("A2.3 方法 2：ノンパラメトリック（経験分布との比）", H2),
      Pg("f_P を過去の実現リターンの分布（カーネル密度など）で置き、M ∝ f_Q / f_P を直接推定する（Jackwerth 2000; Aït-Sahalia and Lo 2000）。長い時系列が要るうえ、推定された M が S について単調減少にならない「プライシング・カーネル・パズル」が知られている。見通し抽出の実務では、方法 1 の γ を校正する方が安定する。"),
      Pg("A2.4 方法 3：Ross の回復定理", H2),
      Pg("Ross (2015) は、状態間の遷移が時間的に一様で状態空間が有界なら、オプション価格から M と f_P を同時に（γ の仮定なしに）回復できると示した。しかし Borovička, Hansen and Scheinkman (2016) は、回復されるのは M の「長期成分」を除いた部分に過ぎず、一般にはリスク・プレミアムが残ることを示し、"
         "Jackwerth and Menner (2020) は S&P 500 で回復した分布が歴史的分布より予測力が劣ることを報告した。実装の優先度は低い。"),
      Pg("A2.5 読み方", H2),
      Pg("γ を一つに決めるより、γ = 0（Q）と γ = 4 程度の二つを並べ、「保険料の幅」として読むのが実用的である。表 A2 では、保険料を除いた実確率の見立てでも P(S_T < 65,000) は 28% 前後あり、下落側の厚みは保険料だけでは説明できない。"),
      PageBreak()]

# ================================================================== A3
lst = a3["last"]
S += [Pg("A3. 分散リスク・プレミアム（VRP）", H1), Pg("A3.1 定義", H2),
      Pg("分散リスク・プレミアムは、オプション価格が織り込む分散（IV²）と、実際に実現すると期待される分散の差である。"),
      EQ(r"\mathrm{VRP}_{t} = \mathrm{IV}_{t}^{2} - E_{t}^{P}\left[\mathrm{RV}_{t,t+\tau}^{2}\right]"),
      Pg("実現分散は観測できるので、事前（ex ante）には直近の実現分散で、事後（ex post）には実際に実現した分散で代用する。本稿では年率ボラティリティの 2 乗（分散）で表し、単位は %²（例: IV 30% なら 900）とする。"),
      Pg("A3.2 なぜプレミアムが生じ、なぜ予測力を持つか", H2),
      Pg("分散を売る側（オプションの売り手）は、急落時に損失が膨らむ。投資家はその状態での損失を嫌うため、分散の価格には対価が上乗せされる。VRP は「分散リスクに対する市場の嫌悪の強さ」の指標であり、"
         "Bollerslev, Tauchen and Zhou (2009) は、VRP が高い（恐怖が強い）ときほどその後 1〜2 四半期の株式リターンが高いことを示した。リスク回避が強まると要求リターンが上がり、価格が下がって将来のリターンが高くなる、という経路である。"
         "Bollerslev, Marrone, Xu and Zhou (2014) は日本を含む国際データで同様の関係を報告した（予測力は国により差がある）。"),
      Pg("A3.3 インプライド分散の計算：ATM IV とモデルフリー IV", H2),
      Pg("IV² に何を使うかで二通りある。ATM IV の 2 乗は簡単だが、スキューの情報を含まない。VIX と同じモデルフリーの定義は、OTM オプションの価格を 1/K² で重み付けして足し合わせる。"),
      EQ(r"\mathrm{IV}_{MF}^{2} = \frac{2}{T}\sum_{i} \frac{\Delta K_{i}}{K_{i}^{2}}\, Q(K_{i}) - \frac{1}{T}\left(\frac{F}{K_{0}} - 1\right)^{2}"),
      Pg("Q(K) は行使価格 K の OTM オプションの価格（K < F ならプット、K ≥ F ならコール）、ΔK は隣の行使価格との間隔の平均、K₀ は F の直下の行使価格である。この式は対数契約（受取 −2 ln(S_T/F)）の複製から導かれ、A1.3 の V と同じ量である（V/T = IV_MF²）。"),
      Pg(f"<b>数値例（10/16 引け）。</b>両気配のある {lst['n']} 行使価格（{lst['kmin']:,.0f}〜{lst['kmax']:,.0f} 円）で IV_MF = {lst['mfiv']*100:.1f}%。ATM IV {lst['atm']*100:.1f}% より高いのは、スキューと裾の厚みが分散に加算されるためである（2026-06 以降の日次で両者の相関は {a3['corr_mf_atm']:.2f}、IV_MF が平均 {(a3['mean_mfiv']-a3['mean_atm'])*100:.1f}pt 高い）。"
         "日経VI は JPX が同じ考え方で 30 日一定残存に補間した指数であり、QRI データから自前で再現できる。"),
      Pg("A3.4 実現ボラティリティ", H2),
      EQ(r"\mathrm{RV}_{t} = \sqrt{250\cdot \frac{1}{n}\sum_{i=1}^{n} r_{t-i}^{2}},\qquad r_{t} = \ln\frac{S_{t}}{S_{t-1}}"),
      Pg(f"日次終値から n = 20 日で計算した。5 分足を使えば推定精度が上がる（Bollerslev らは 5 分の実現分散を用いる）。10/16 時点で RV₂₀ = {lst['rv20']*100:.1f}%、RV₆₀ = {lst['rv60']*100:.1f}%。"),
      CondPageBreak(150 * mm), Pg("A3.5 系列と読み方", H2), fig_vrp(),
      TBL([["期間 2026-06-03〜10-16（n = %d 日）" % a3["n"], "値"],
           ["IV_MF の平均 / ATM IV の平均 / RV₂₀ の平均", f"{a3['mean_mfiv']*100:.1f}% / {a3['mean_atm']*100:.1f}% / {a3['mean_rv20']*100:.1f}%"],
           ["VRP（事前）= IV_MF² − RV₂₀² の平均", f"{a3['mean_vrp_mf']*1e4:,.0f} %²（正の日の割合 {a3['share_pos']*100:.0f}%）"],
           ["VRP（事後）= IV_MF² − その後 20 日の RV² の平均", f"{a3['mean_expost']*1e4:,.0f} %²（n = {a3['n_expost']}）"],
           ["10/16 の VRP（事前）", f"{lst['vrp_mf']*1e4:,.0f} %²（IV_MF {lst['mfiv']*100:.1f}%、RV₂₀ {lst['rv20']*100:.1f}%）"]], [100, 78]),
      Pg("表 A3  分散リスク・プレミアムの要約。", CAP),
      Pg("事前の VRP は 4 か月の 78% の日で正、事後でも平均は正であり、分散の売り手に平均して対価が支払われている。急落日（8/19、10/15）には実現が追いつき VRP が縮むか負になり、その後再び開く。これは分散プレミアムの典型的な動きである。"
         "10/16 は ATM IV と RV₂₀ がほぼ等しく（VRP_ATM ≈ 0）、スキュー分を含む IV_MF ベースでは正である。"),
      Pg("<b>予測への使い方。</b>Bollerslev らの回帰は r_{t→t+h} = a + b·VRP_t + ε で、h = 3 か月前後に b が最大になる。重複標本のため標準誤差は Hodrick (1992) 型で補正する。本ワークスペースでは 2026-06 からの 4 か月分しかなく、検証には 2 年以上の蓄積か、日経VI と日経平均の長期系列（外部データ）の利用が要る。"),
      PageBreak()]

# ================================================================== A4
S += [Pg("A4. スキューとリスクリバーサル", H1), Pg("A4.1 指標の定義", H2),
      Pg("スキューは「同じ満期で、低い行使価格ほど IV が高い」傾きである。測り方は三つある。"),
      TBL([["指標", "定義", "特徴"],
           ["RR25", "σ(Δ = −0.25) − σ(Δ = +0.25)。25 デルタのプット IV とコール IV の差", "デルタで位置を決めるため、F や満期が変わっても比較できる。最も一般的"],
           ["BF25", "[σ(Δ = −0.25) + σ(Δ = +0.25)] / 2 − σ_ATM", "曲がり（両翼の高さ）。尖度に対応"],
           ["リスク中立歪度", "A1 のモーメント。密度またはモデルフリー", "分布全体の非対称。両翼の補外に敏感"],
           ["CBOE SKEW 型", "100 − 10 × Skew_Q（30 日）", "米国で公表される指数の定義。歪度 −1 なら 110"]], [30, 80, 68]),
      Pg(f"10/16 引け: RR25 = {a1['rr25']*100:+.2f}pt（SVI）、QRI デルタからの直接補間で +{a7['rows'][0]['rr25']:.2f}pt、BF25 = {a1['bf25']*100:+.2f}pt。"),
      Pg("A4.2 スキューは何で決まるか", H2),
      Pg("Bakshi, Kapadia and Madan (2003) は、べき効用の代表的投資家のもとで、リスク中立歪度が実確率の歪度・尖度とリスク回避度から次のように決まることを示した（近似）。"),
      EQ(r"\mathrm{Skew}_{Q} \approx \mathrm{Skew}_{P} - \gamma\, s_{P}\, (\mathrm{Kurt}_{P} - 3)"),
      Pg("s_P は実確率の標準偏差である。実確率の歪度がゼロでも、裾が厚く（Kurt_P > 3）リスク回避が強い（γ > 0）なら、リスク中立歪度は負になる。つまり観測されるスキューは「実際に下落が起こりやすい」のか「下落への嫌悪が強い」のかを区別せずに含む。"
         "RR25 の上昇は、この二つのどちらか、または両方の上昇として読む。"),
      Pg("A4.3 実証の示唆", H2),
      Pg("・個別株では、スマイルの傾きが急な銘柄ほどその後のリターンが低い（Xing, Zhang and Zhao 2010：情報を持つ投資家がプットで悪材料を先取りする）。事前のリスク中立歪度が低い（負が大きい）銘柄ほど期待リターンが高いという逆の関係も報告されている（Conrad, Dittmar and Ghysels 2013）。"
         "指数の時系列では、スキューの水準そのものの予測力は弱く、変化（急拡大）がリスク回避の上昇を示す指標として使われることが多い。"),
      Pg("・An, Ang, Bali and Cakici (2014) は、コール IV の上昇がその後の株価上昇、プット IV の上昇が下落と関係することを示した。IV の「変化」が投資家の情報を映すという解釈で、IV モニターの超過ΔIV の考え方と同じ方向である。"),
      CondPageBreak(130 * mm), Pg("A4.4 系列の読み方", H2), fig_rr(),
      Pg(f"図4 では、F が下がる局面で RR25 が {SER[0]['rr25_delta']:+.1f} から {SER[-1]['rr25_delta']:+.1f}pt へ拡大している。これは A4.2 の式で γ·s_P·(Kurt_P − 3) が大きくなった（下落時の嫌悪が強まった）ことと整合する。"
         "逆に、F が下がっているのに RR25 が縮む局面があれば、下落の保険需要が後退した（売り疲れ、またはヘッジの手仕舞い）と読める。スキューを使うときは「水準」より「F の変化に対する相対的な変化」を見る。"),
      PageBreak()]

# ================================================================== A5
hq, hh = a5["q"], a5["hist"]
S += [Pg("A5. テール（ジャンプ）の確率", H1), Pg("A5.1 問題", H2),
      Pg("深い OTM プットの価格から得られる「大幅安の確率」は、実際の発生頻度より高い。差はテール・リスク・プレミアムである。Bollerslev and Todorov (2011) は、高頻度データで検出した実際のジャンプの分布（実確率の裾）と、OTM オプションから推定したリスク中立の裾を比べ、"
         "株式のリスク・プレミアムのかなりの部分がジャンプ・テールへの対価であること、そのプレミアム（「投資家の恐怖」指数）がリターンの予測力を持つことを示した。"),
      Pg("A5.2 方法", H2),
      Pg("リスク中立の裾は、A1 の密度から P_Q(S_T < (1 − a)F) を直接読む。実確率の裾は、(i) 過去の同じ期間長のリターンの頻度、(ii) 条件付きの頻度（開始時のボラティリティが同程度の期間に限る）、(iii) 高頻度データのジャンプ検出（二乗変動と二乗偏差変動の差）から推定する。本稿は (i)(ii) を示す。"),
      fig_tail(),
      TBL([["28 日間（20 営業日）の変化率", "リスク中立 P_Q（10/16）", "実績頻度 2000 年〜", "実績頻度、開始時 RV₂₀ ≥ 25%"],
           ["−10% 未満", pct(hq["p_lt_m10"]), pct(hh["p_lt_m10"]), pct(hh["hi"]["p_lt_m10"])],
           ["−5% 未満", pct(hq["p_lt_m5"]), pct(hh["p_lt_m5"]), pct(hh["hi"]["p_lt_m5"])],
           ["+5% 超", pct(hq["p_gt_p5"]), pct(hh["p_gt_p5"]), pct(hh["hi"]["p_gt_p5"])],
           ["+7% 超", pct(hq["p_gt_p7"]), pct(hh["p_gt_p7"]), "—"]], [50, 42, 42, 44]),
      Pg(f"表 A5  裾の確率の比較。実績は日経平均の日次終値（2000-01〜2026-10、n = {hh['n']:,}）の 20 営業日対数リターン。条件付きは開始時点の RV₂₀ が 25% 以上の期間（n = {hh['hi']['n']:,}）。", CAP),
      Pg(f"<b>読み方。</b>−10% 未満の確率は、リスク中立で {pct(hq['p_lt_m10'])}、無条件の実績では {pct(hh['p_lt_m10'])} と、価格は頻度の約 2 倍の保険料を含む。しかし現在のように RV₂₀ が {a5['rv20_now']*100:.0f}% と高い局面に限ると実績頻度は {pct(hh['hi']['p_lt_m10'])} に上がり、リスク中立確率を上回る。"
         "つまり「割高な保険」に見える左裾も、ボラティリティの高い局面での経験頻度と比べると割高ではない。テール確率の評価は、条件付けの仕方で結論が反転し得る。"),
      Pg("<b>限界。</b>深い OTM の気配は幅が広く、SVI の両翼は直線的な補外である。歴史的頻度は 2008 年や 2020 年の急落を含むかどうかで大きく変わる。確率の絶対水準よりも、同じ規約で計算した「価格の裾 ÷ 実績の裾」の比率の推移を見るのが実用的である。"),
      PageBreak()]

# ================================================================== A6
rows6 = a6["rows"]
S += [Pg("A6. 校正検証", H1), Pg("A6.1 なぜ必要か", H2),
      Pg("分布の形を読む前に、その分布が「当たっている」かを確かめる必要がある。確率予測の正しさは、一つの実現値では評価できず、多数の予測と実現の対を集めて初めて検定できる。これが校正（calibration）検証である。A2 の γ の推定も、この蓄積の上で行う。"),
      Pg("A6.2 方法：確率積分変換と Berkowitz 検定", H2),
      Pg("予測時点 i の累積分布 Φ_i で、実現した SQ値 S_i を変換した値を u_i とする。"),
      EQ(r"u_{i} = \Phi_{i}(S_{i}),\qquad z_{i} = N^{-1}(u_{i})"),
      Pg("分布が正しければ u_i は区間 [0, 1] の一様分布に、z_i は標準正規分布に従う。Berkowitz (2001) は z_i に 1 次の自己回帰 z_i = μ + ρ z_{i−1} + ε_i（ε の分散 σ²）を当てはめ、(μ, σ², ρ) = (0, 1, 0) を尤度比で検定する。"),
      EQ(r"LR = -2\left[\,\ell(0,1,0) - \ell(\hat\mu,\hat\sigma^{2},\hat\rho)\,\right] \sim \chi^{2}(3)"),
      Pg("u_i が 0 や 1 に偏れば「幅が狭すぎる」、0.5 付近に集まれば「幅が広すぎる」、片側に偏れば「中心がずれている（プレミアムの符号）」と読める。一様性だけならコルモゴロフ–スミルノフ検定も使える。"),
      Pg("A6.3 数値例：10月限", H2), fig_pit(),
      TBL([["予測時点", "F", "残存", "5% 点", "50% 点", "95% 点", "実現 SQ", "u = Φ(SQ)", "z"]] +
          [[f"{r['day'][4:6]}/{r['day'][6:]}", fmt(r["F"]), f"{r['T_days']:.0f} 日", fmt(r["q05"]), fmt(r["q50"]), fmt(r["q95"]), fmt(a6["sq"]), f"{r['u']:.3f}", f"{r['z']:+.2f}"] for r in rows6], [22, 20, 16, 20, 20, 20, 22, 22, 16]),
      Pg(f"表 A6  10月限（SQ 2026-10-09、実現値 {a6['sq']:,.2f} 円）の校正。", CAP),
      Pg(f"9 月中旬の予測では実現値が分布の上位 13〜19%（u = 0.81〜0.87）に入り、10/1 の予測では中央付近（u = 0.43）だった。1 限月分の例なので検定はできないが、同じ満期に対する複数時点の u は互いに独立でない点に注意が要る（検定には満期ごとに 1 点を使う）。"),
      Pg("<b>標本数の確保。</b>月次限月では年 12 点しか得られない。日経225ミニオプション（週次満期、月曜・水曜・金曜）を加えれば年 100 点を超え、1 年で検定可能になる。ワークスペースの日報 PDF 読み取りは週次銘柄に対応済みである。"),
      PageBreak()]

# ================================================================== A7
r11, r12, r03 = a7["rows"]; fw = a7["forward"]
S += [Pg("A7. 期間構造", H1), Pg("A7.1 何が読めるか", H2),
      Pg("複数の限月の IV を並べると、(a) 市場が想定するボラティリティの時間的な配分、(b) 特定の期間（イベントを含む期間）に集中した分散、(c) スキューが満期とともにどう減るか、が分かる。"),
      Pg("A7.2 フォワード分散", H2),
      Pg("総分散は加法的である。満期 T₁ と T₂（T₁ < T₂）の ATM IV を σ₁、σ₂ とすると、T₁ から T₂ の間の分散（フォワード分散）は"),
      EQ(r"\sigma_{f}^{2}(T_{1},T_{2}) = \frac{\sigma_{2}^{2}T_{2} - \sigma_{1}^{2}T_{1}}{T_{2} - T_{1}}"),
      Pg("である。無裁定なら σ_f² ≥ 0（カレンダー・スプレッドの条件）。イベント（日銀会合、決算、米国指標）が T₁ と T₂ の間にあれば、σ_f² はそのイベントの分散を含む。平常時の分散 σ_base² を別途見積もれば、イベントの寄与は (σ_f² − σ_base²)(T₂ − T₁) として取り出せる。"),
      fig_term(),
      TBL([["限月", "残存日数", "ATM IV", "RR25（Δ基準）", "リスク中立 5% 点", "95% 点", "当てはめ誤差"]] +
          [[r["cm"][:2] + "年" + r["cm"][2:] + "月限", f"{r['T']*365:.0f}", f"{r['atm']*100:.1f}%", f"{r['rr25']:+.1f}pt", fmt(r["q05"]), fmt(r["q95"]), f"{r['rmse']*100:.2f}pt"] for r in a7["rows"]], [34, 22, 22, 28, 28, 24, 24]),
      Pg("表 A7  10/16 引けの期間構造。", CAP),
      Pg(f"<b>数値例。</b>11月→12月のフォワード IV は √[(σ₁₂²T₁₂ − σ₁₁²T₁₁)/(T₁₂ − T₁₁)] = {fw[0]['sig_f']*100:.1f}%、12月→3月は {fw[1]['sig_f']*100:.1f}%。期近 {r11['atm']*100:.1f}% から遠い限月ほど低い右下がり（バックワーデーション）で、市場は「直近の 1 か月が最も荒く、その後は落ち着く」と織り込んでいる。"
         "平時の日経225 は期先ほど高いコンタンゴが普通であり、右下がりはストレス局面の形である。"),
      Pg(f"<b>スキューの期間構造。</b>RR25 は {r11['rr25']:+.1f}pt（28 日）、{r12['rr25']:+.1f}pt（56 日）、{r03['rr25']:+.1f}pt（147 日）。独立な日次リターンなら歪度は 1/√T で減るので RR25 も同程度に減るはずだが（図7b の破線）、実測は緩やかにしか減らない。"
         "ボラティリティの持続とレバレッジ効果により、長い満期でも非対称が残ることを示す。"),
      Pg("<b>読み方。</b>期間構造の傾き（期近 − 期先）は、リスク選好の短期的な悪化の指標として VIX の期間構造と同様に使える。日経VI 先物の限月間価格差も同じ情報を持つ。"),
      PageBreak()]

# ================================================================== A8
cnt = a8["counts"]; nb = cnt.get("新規買い", {}); ns = cnt.get("新規売り", {}); cl = cnt.get("手仕舞い", {}); bb = cnt.get("買い戻し", {})
S += [Pg("A8. オプションのフロー", H1), Pg("A8.1 理論と実証", H2),
      Pg("情報を持つ投資家は、レバレッジが高く空売り制約の無いオプション市場で先に動く、という仮説がある（Easley, O'Hara and Srinivas 1998）。Pan and Poteshman (2006) は、取引所の建て／落ち区分付きの出来高を使い、"
         "「新規建ての買い注文」に限ったプット・コール比が翌日〜翌週の株価を予測すること（比率が高い銘柄ほど低いリターン）、公開情報である通常の出来高には予測力がほとんど無いことを示した。情報は「誰が、新規に、買ったか」に宿る。"),
      EQ(r"\mathrm{PC}_{t}^{\mathrm{open\,buy}} = \frac{\mathrm{Put\ open\!-\!buy\ volume}_{t}}{\mathrm{Put\ open\!-\!buy\ volume}_{t} + \mathrm{Call\ open\!-\!buy\ volume}_{t}}"),
      Pg("A8.2 JPX データでの代理指標", H2),
      Pg("JPX の公表データには建て／落ちの区分も、買い／売りの区分も無い。代理として三つを使う。"),
      Pg("・<b>新規建て数量</b>：建玉の増加分 max(ΔOI, 0) は、新規の買いと売りが同数成立したことを示す（方向は不明）。"),
      Pg("・<b>方向の推定</b>：IV モニターの判定（Δ建玉 × 超過ΔIV）。建玉が増え、同時にその行使価格の IV が地合いを超えて上がれば「新規買い」、下がれば「新規売り」とみなす。"),
      Pg("・<b>出来高加重の超過ΔIV</b>：売買が入った時間帯の IV の動きだけを集計した値（IV モニターの列）。買い主導の約定か売り主導の約定かの強さを示す。"),
      TBL([["判定（10/16、11月限）", "コール（銘柄数）", "プット（銘柄数）"],
           ["新規買い（建玉増 × IV 上昇）", str(nb.get("CALL", 0)), str(nb.get("PUT", 0))],
           ["新規売り（建玉増 × IV 低下）", str(ns.get("CALL", 0)), str(ns.get("PUT", 0))],
           ["買い戻し（建玉減 × IV 上昇）", str(bb.get("CALL", 0)), str(bb.get("PUT", 0))],
           ["手仕舞い（建玉減 × IV 低下）", str(cl.get("CALL", 0)), str(cl.get("PUT", 0))],
           ["新規建て数量 Σ max(ΔOI, 0)（枚）", f"{a8['new_oi'].get('CALL', 0):,.0f}", f"{a8['new_oi'].get('PUT', 0):,.0f}"]], [80, 49, 49]),
      Pg(f"表 A8  判定の集計。新規建ての P/C 比 = {a8['pc_new']:.2f}。地合い（限月内ΔIV 中央値）{a8['meta']['level_pct']:+.2f}pt。", CAP),
      TBL([["タイプ", "行使", "Δ建玉", "建玉", "ΔIV", "超過ΔIV", "判定"]] + [[t["option_type"], fmt(t["strike"]), f"{t['d_oi']:+,.0f}", fmt(t["oi"]), f"{t['d_iv_pct']:+.1f}", f"{t['d_iv_ex_pct']:+.1f}", t["judge"]] for t in a8["top"][:8]], [22, 22, 24, 24, 22, 24, 40]),
      Pg("表 A9  Δ建玉の大きい銘柄（10/16）。", CAP),
      Pg("10/16 は、新規売りの判定がコール・プットとも新規買いを上回った。急落翌日に IV が全体として低下する中で（地合い −1.35pt）、68,000 円コールと 65,500 円プットの新規建てが売り主導と判定され、深い OTM（50,000 円プット）と 71,000 円コールの新規建てが買い主導と判定されている。"
         "Pan–Poteshman の比率に相当するのは「新規買いと判定されたプット ÷（同プット + 同コール）」であり、この日は 17/(17+12) = 0.59 である。"),
      Pg("<b>限界。</b>判定は IV の動きで方向を推定する代理であり、ディーラーのヘッジや裁定の売買を顧客の方向性注文と区別できない。J-NET（立会外）の大口は別途集計し、立会の判定と分けて扱う。予測力の検証には判定の日次集計を蓄積し、翌週のリターンとの回帰を行う。"),
      PageBreak()]

# ================================================================== A9
S += [Pg("A9. ディーラーのガンマ（GEX）", H1), Pg("A9.1 見通しではなく経路", H2),
      Pg("オプションのマーケットメイカー（ディーラー）は、顧客の売買の反対側を持ち、先物でデルタをヘッジする。ディーラーのポートフォリオのガンマが負なら、先物が下がるほど売り、上がるほど買うヘッジが機械的に出て、値動きを増幅する。正なら逆で、値動きを吸収する。"
         "これは投資家の見通しではないが、「同じニュースに対して相場がどれだけ動きやすいか」を決める経路の情報である。"),
      Pg("A9.2 定義", H2),
      EQ(r"\mathrm{GEX}(K) = \pm\,\Gamma(K)\cdot \mathrm{OI}(K)\cdot 1000\cdot F^{2}\cdot 0.01,\qquad \mathrm{GEX} = \sum_{K}\mathrm{GEX}(K)"),
      Pg("符号は「ディーラーがコールの買い手・プットの売り手」という慣例的仮定で、コール＋・プット−とする。GEX は先物が 1% 動いたときにディーラーが出すヘッジ売買の金額（円）である。"
         "累積 GEX（低い行使価格から足し上げた値）の符号が変わる水準を「ガンマ・フリップ」と呼び、そこを境に値動きの性質が変わるとされる。"),
      fig_gex(),
      Pg(f"<b>数値例（10/16 建玉）。</b>合計 {a9['total']/1e8:+,.0f} 億円/1%。累積はどの水準でも負で、ガンマ・フリップは表示範囲に無い（プット建玉がコールを大きく上回るため）。"
         f"10/16 夜間の下落（{a9['night_move']*100:+.1f}%）に対し、符号の仮定が正しければディーラーのヘッジ売りは約 {-a9['total']*abs(a9['night_move'])*100/1e8:,.0f} 億円と見積もられる。"),
      Pg("A9.3 実証", H2),
      Pg("Barbon and Buraschi (2020) は、ディーラーのガンマが負のときに日中の値動きの自己相関（モメンタム）が強まり、実現ボラティリティが上がることを示した。Baltussen, Da, Lammers and Martens (2021) は、レバレッジ ETF やオプションのヘッジ需要が引けにかけての値動きの持続（日中モメンタム）を生むことを示した。"
         "Ni, Pearson and Poteshman (2005) は満期日の価格が行使価格に引き寄せられる現象をディーラーのヘッジで説明した。"),
      Pg("<b>日本市場での注意。</b>日経225オプションではプットの売り手に個人投資家や事業法人が多いとされ、「ディーラーがプットの売り手」という慣例の符号が当てはまらない可能性がある。符号を観測できない以上、GEX は大きさと集中する行使価格の情報として使い、符号は手口（取引参加者別建玉）から補う。"),
      PageBreak()]

# ================================================================== 10
S += [Pg("10. 統合・実装・参考文献", H1), Pg("10.1 各指標の性質", H2),
      TBL([["章", "指標", "種類", "頻度", "状態", "検証に必要なもの"],
           ["A1", "形状指標（歪度・尖度・RR25・裾確率）", "記述（Q）", "日次", "実装済み", "1 年以上の系列と A6 の校正"],
           ["A2", "γ による実確率の近似", "補正", "日次", "本稿で試算", "A6 の蓄積で γ を推定"],
           ["A3", "VRP（IV_MF² − RV²）", "予測（リターン）", "日次", "本稿で系列化", "2 年以上、または日経VI の長期系列"],
           ["A4", "RR25 の水準と変化", "記述・予測", "日次", "実装済み", "翌週リターンとの回帰"],
           ["A5", "裾確率 ÷ 実績頻度", "補正・記述", "日次", "本稿で試算", "高頻度データでのジャンプ検出"],
           ["A6", "PIT・Berkowitz 検定", "検証", "満期ごと", "本稿で例示", "週次満期の取り込み"],
           ["A7", "フォワード分散・スキューの期間構造", "記述", "日次", "本稿で試算", "イベント日程との対応付け"],
           ["A8", "新規建ての P/C、判定の集計", "予測（情報）", "日次", "実装済み（判定）", "日次集計の蓄積と回帰"],
           ["A9", "GEX・ガンマ・フリップ", "経路", "日次", "実装済み", "手口による符号の補正"]], [12, 54, 28, 16, 28, 40]),
      Pg("10.2 実装の順序", H2),
      Pg("(1) A1・A3・A4・A7 の日次指標を一つの表に保存する（既存の rnd_series.py と outlook_a_calc.py を日次収集に組み込む）。(2) A6 の校正ログを満期ごとに記録し、週次満期を加える。(3) 1 年後に A2 の γ と A3・A4・A8 の予測回帰を行う。"
         "(4) A9 は手口データと結合して符号を推定する。指標の読み方は、絶対水準ではなく直近 1 年の分位で統一する。"),
      Pg("10.3 検証の設計", H2),
      Pg("予測対象は「1〜4 週先の先物リターン」と「SQ値の分布」の二つに固定する。前者は予測回帰（重複標本の標準誤差補正）、後者は A6 の校正検定で評価する。サンプル期間で結果が変わること（Shiratsuka 2001）を前提に、期間を分けた安定性の確認を必ず行う。"),
      sp(4), Pg("参考文献", H2),
      Pg("Aït-Sahalia, Y. and A. Lo (2000), Nonparametric Risk Management and Implied Risk Aversion, Journal of Econometrics 94.", SMALL),
      Pg("An, B.-J., A. Ang, T. Bali and N. Cakici (2014), The Joint Cross Section of Stocks and Options, Journal of Finance 69.", SMALL),
      Pg("Bakshi, G., N. Kapadia and D. Madan (2003), Stock Return Characteristics, Skew Laws, and the Differential Pricing of Individual Equity Options, Review of Financial Studies 16.", SMALL),
      Pg("Baltussen, G., Z. Da, S. Lammers and M. Martens (2021), Hedging Demand and Market Intraday Momentum, Journal of Financial Economics 142.", SMALL),
      Pg("Barbon, A. and A. Buraschi (2020), Gamma Fragility, working paper (SSRN).", SMALL),
      Pg("Berkowitz, J. (2001), Testing Density Forecasts, with Applications to Risk Management, Journal of Business and Economic Statistics 19.", SMALL),
      Pg("Birru, J. and S. Figlewski (2012), Anatomy of a Meltdown: The Risk Neutral Density for the S&P 500 in the Fall of 2008, Journal of Financial Markets 15.", SMALL),
      Pg("Bliss, R. and N. Panigirtzoglou (2004), Option-Implied Risk Aversion Estimates, Journal of Finance 59.", SMALL),
      Pg("Bollerslev, T., G. Tauchen and H. Zhou (2009), Expected Stock Returns and Variance Risk Premia, Review of Financial Studies 22.", SMALL),
      Pg("Bollerslev, T., J. Marrone, L. Xu and H. Zhou (2014), Stock Return Predictability and Variance Risk Premia: Statistical Inference and International Evidence, Journal of Financial and Quantitative Analysis 49.", SMALL),
      Pg("Bollerslev, T. and V. Todorov (2011), Tails, Fears, and Risk Premia, Journal of Finance 66.", SMALL),
      Pg("Borovička, J., L. P. Hansen and J. Scheinkman (2016), Misspecified Recovery, Journal of Finance 71.", SMALL),
      Pg("Breeden, D. and R. Litzenberger (1978), Prices of State-Contingent Claims Implicit in Option Prices, Journal of Business 51.", SMALL),
      Pg("Carr, P. and D. Madan (2001), Optimal Positioning in Derivative Securities, Quantitative Finance 1.", SMALL),
      Pg("Conrad, J., R. Dittmar and E. Ghysels (2013), Ex Ante Skewness and Expected Stock Returns, Journal of Finance 68.", SMALL),
      Pg("Easley, D., M. O'Hara and P. Srinivas (1998), Option Volume and Stock Prices: Evidence on Where Informed Traders Trade, Journal of Finance 53.", SMALL),
      Pg("Hodrick, R. (1992), Dividend Yields and Expected Stock Returns: Alternative Procedures for Inference and Measurement, Review of Financial Studies 5.", SMALL),
      Pg("Jackwerth, J. (2000), Recovering Risk Aversion from Option Prices and Realized Returns, Review of Financial Studies 13.", SMALL),
      Pg("Jackwerth, J. and M. Menner (2020), Does the Ross Recovery Theorem Work Empirically?, Journal of Financial Economics 137.", SMALL),
      Pg("Jiang, G. and Y. Tian (2005), The Model-Free Implied Volatility and Its Information Content, Review of Financial Studies 18.", SMALL),
      Pg("Ni, S., N. Pearson and A. Poteshman (2005), Stock Price Clustering on Option Expiration Dates, Journal of Financial Economics 78.", SMALL),
      Pg("Pan, J. and A. Poteshman (2006), The Information in Option Volume for Future Stock Prices, Review of Financial Studies 19.", SMALL),
      Pg("Ross, S. (2015), The Recovery Theorem, Journal of Finance 70.", SMALL),
      Pg("Shiratsuka, S. (2001), Information Content of Implied Probability Distributions: Empirical Studies on Japanese Stock Price Index Options, Monetary and Economic Studies 19(3), Bank of Japan.", SMALL),
      Pg("Xing, Y., X. Zhang and R. Zhao (2010), What Does the Individual Option Volatility Smirk Tell Us About Future Equity Returns?, Journal of Financial and Quantitative Analysis 45.", SMALL)]

doc = SimpleDocTemplate(str(PDF_PATH), pagesize=A4, leftMargin=16 * mm, rightMargin=16 * mm, topMargin=16 * mm, bottomMargin=18 * mm,
                        title="オプション価格から先物の先行きと投資家の見通しを抽出する方法（系統A 詳細解説）", author="先物手口分析")
ft = make_footer(f"系統A 詳細解説：オプション価格からの見通し抽出（数値例 2026-10-16 引け、作成 {date.today():%Y-%m-%d}）")
doc.build(S, onFirstPage=ft, onLaterPages=ft)
print("written", PDF_PATH)
