"""「市場が織り込むリスク回避度 γ」の推定レポート（PDF）。gamma_calc.py の出力を用いる。

使い方: python scripts/build_gamma_report.py <計算結果ディレクトリ> <出力フォルダ>
"""
from __future__ import annotations

import json
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

from report_kit import A, B, BODY, CAP, EQ, GRID, H1, H2, INK, MUT, O, SMALL, SUB, TBL, TITLE, Y, Pg, fig_to_img, fmt, make_footer, sp, style_ax

CALC = Path(sys.argv[1]); OUT_DIR = Path(sys.argv[2]); OUT_DIR.mkdir(parents=True, exist_ok=True)
PDF_PATH = OUT_DIR / "nk225op_market_implied_gamma_report.pdf"
G = json.load(open(CALC / "gamma_calc.json")); ML = json.load(open(CALC / "gamma_mle.json")); S_ = pd.read_json(CALC / "gamma_series.json"); L = np.load(CALC / "gamma_latest.npz")
S_["day"] = S_.day.astype(str); S_["date"] = pd.to_datetime(S_.day); S_ = S_[S_.ts.str[11:13].astype(int) >= 8].reset_index(drop=True)  # 夜間のみの日を除外
rel, boot, mle, lat = G["series_summary"], G["boot"], G["mle"], G["latest"]
last = S_.iloc[-1]; VAR = ["gamma_I", "gamma_rv60", "gamma_c80", "gamma_bw"]
if "gamma_q" not in S_.columns:
    S_["gamma_q"] = S_.gamma_I_quad; S_["beta2"] = S_.quad; S_["ivrv"] = S_.atm - S_.rv20
pct = lambda v: f"{v*100:.1f}%"


# ------------------------------------------------------------------ 図
def fig_series():
    fig, ax = plt.subplots(4, 1, figsize=(7.2, 9.2), sharex=True); d = S_.date
    lo, hi = S_[VAR].min(axis=1), S_[VAR].max(axis=1)
    ax[0].fill_between(d, lo, hi, color=MUT, alpha=0.12, label="参考: 直線 γ̂ の変種（RV60・中央 80%・狭い帯域）の範囲")
    ax[0].plot(d, S_.gamma_q, color=B, lw=2.0, label="γ̂_q（主指標：2 次式の傾き、ATM 近傍の局所 γ）")
    ax[0].plot(d, S_.gamma_I, color=MUT, lw=1.0, label="参考: 直線の γ̂（方法 I 基準設定）"); ax[0].plot(d, S_.gamma_II, color=O, lw=1.0, ls="--", label="参考: 方法 II（モーメント法）")
    ax[0].axhline(0, color=MUT, lw=0.8); ax[0].set_ylim(min(-2, float(S_.gamma_q.min()) - 0.3), max(6, float(S_.gamma_II.quantile(0.95)) + 0.5))
    style_ax(ax[0], "図1  市場が織り込む γ の日次推定（期近限月、引け）", "", "γ"); ax[0].legend(fontsize=7, frameon=False)
    ax[1].plot(d, S_.beta2, color=B, lw=2.0, label="β₂（主指標：曲率、左軸）"); a2 = ax[1].twinx(); a2.plot(d, S_.ivrv * 100, color=O, lw=1.1, label="ATM IV − RV₂₀（%pt、右軸）"); a2.tick_params(colors=MUT, labelsize=8)
    ax[1].axhline(0, color=MUT, lw=0.8); style_ax(ax[1], "", "", "β₂"); h1, l1 = ax[1].get_legend_handles_labels(); h2, l2 = a2.get_legend_handles_labels(); ax[1].legend(h1 + h2, l1 + l2, fontsize=7, frameon=False)
    ax[2].plot(d, S_.F, color=INK, lw=1.6); style_ax(ax[2], "", "", "先物推定値 F（円）")
    ax[3].plot(d, S_.atm * 100, color=B, lw=1.4, label="ATM IV"); ax[3].plot(d, S_.rv20 * 100, color=O, lw=1.4, label="RV₂₀"); style_ax(ax[3], "", "", "%"); ax[3].legend(fontsize=7.5, frameon=False)
    fig.autofmt_xdate()
    return fig_to_img(fig)


def fig_kernel():
    x, lnM = L["xsel"], L["lnM"]; a_, g_ = lat["a"], lat["gamma"]
    w = np.interp(x, L["x"], L["fQ"] * np.gradient(L["x"])); c2 = np.polyfit(x, lnM, 2, w=np.sqrt(np.clip(w, 1e-12, None)))  # 2 次式（重み = f_Q 質量）
    fig, ax = plt.subplots(1, 2, figsize=(7.2, 3.0))
    ax[0].plot(x * 100, lnM, color=INK, lw=1.8, label="ln[f_Q(x)/f_P(x)]（観測）"); ax[0].plot(x * 100, np.polyval(c2, x), color=B, lw=1.6, label=f"2 次式（主指標）: 傾き γ̂_q = {-c2[1]:+.2f}、曲率 β₂ = {c2[0]:+.1f}")
    ax[0].plot(x * 100, a_ - g_ * x, color=MUT, lw=1.1, ls="--", label=f"参考: 直線（γ̂ = {g_:+.2f}）")
    style_ax(ax[0], f"図2a  対数プライシング・カーネル（{lat['day'][4:6]}/{lat['day'][6:]} 引け）", "x = ln(K/F)（%）", "ln M + 定数"); ax[0].legend(fontsize=7, frameon=False)
    K = L["K"]; sel = (K >= 0.78 * lat["F"]) & (K <= 1.2 * lat["F"]); xg = L["x"]
    ax[1].plot(K[sel], (L["fQ"] / K)[sel] * 1e5, color=B, lw=1.8, label="f_Q（オプション）"); ax[1].plot(K[sel], (L["fP"] / K)[sel] * 1e5, color=MUT, lw=1.4, ls="--", label="f_P（歴史的リターン、現在の RV で再スケール）")
    ax[1].axvline(lat["F"], color=INK, lw=0.8); style_ax(ax[1], "図2b  二つの密度", "SQ値（円）", "確率密度（%/1,000円）"); ax[1].legend(fontsize=7, frameon=False)
    fig.tight_layout()
    return fig_to_img(fig)


def fig_relation():
    fig, ax = plt.subplots(2, 2, figsize=(7.2, 5.6)); ax = ax.ravel()
    ax[0].scatter(S_.ret20 * 100, S_.gamma_q, s=14, color=B); style_ax(ax[0], "図3a  γ̂_q と直近 20 日リターン", "20 日対数リターン（%）", "γ̂_q")
    ax[1].scatter(S_.rv20 * 100, S_.gamma_q, s=14, color=B); style_ax(ax[1], "図3b  γ̂_q と実現ボラティリティ RV₂₀", "RV₂₀（%）", "γ̂_q")
    ax[2].scatter(S_.ivrv * 100, S_.beta2, s=14, color=O); style_ax(ax[2], "図3c  曲率 β₂ と ATM IV − RV₂₀", "ATM IV − RV₂₀（%pt）", "β₂")
    ax[3].scatter(S_.ret20 * 100, S_.beta2, s=14, color=O); style_ax(ax[3], "図3d  曲率 β₂ と直近 20 日リターン", "20 日対数リターン（%）", "β₂")
    for a_ in ax:
        a_.axhline(0, color=MUT, lw=0.8)
    fig.tight_layout()
    return fig_to_img(fig)


def fig_mle():
    g = np.array(ML["grid"]); pm, pa, pw = ML["profile_monthly"], ML["profile_all"], ML["profile_weekly"]
    fig, a = plt.subplots(figsize=(7.2, 3.0)); ll = np.array(pm["ll"]); a.plot(g, ll - ll.max(), color=B, lw=2, label=f"月次限月（{pm['n']} 点、残存 28 日前後の予測）")
    if pw:
        lw_ = np.array(pw["ll"]); a.plot(g, lw_ - lw_.max(), color=O, lw=1.6, ls="--", label=f"週次満期（{pw['n']} 点）")
        la = np.array(pa["ll"]); a.plot(g, la - la.max(), color=INK, lw=1.2, label=f"合計（{pa['n']} 点）")
    a.axhline(-1.92, color=MUT, lw=0.8, ls=":"); a.text(g[-1], -1.8, "95% 区間の閾値（−1.92）", ha="right", fontsize=7.5, color=MUT)
    a.axvline(pa["gamma_hat"], color=B, lw=0.8); a.axvline(0, color=MUT, lw=0.6); a.set_ylim(-4, 0.3)
    style_ax(a, "図4  最尤法（方法 III）のプロファイル対数尤度（最大値からの差、格子 −6〜12）", "γ", "対数尤度の差"); a.legend(fontsize=7.5, frameon=False, loc="lower left")
    return fig_to_img(fig)


# ------------------------------------------------------------------ 文書
S = []
S += [sp(8), Pg("市場が織り込むリスク回避度 γ の推定", TITLE), Pg("日経225オプションのリスク中立分布と歴史的リターン分布から：方法・日次系列・精度・蓄積の設計", SUB),
      TBL([["項目", "内容"],
           ["対象", f"日経225オプション期近限月、各営業日の日中引け（気配時刻 15:30:59 以前の最終）。{S_.day.iloc[0]} 〜 {S_.day.iloc[-1]}（{len(S_)} 営業日）"],
           ["γ の意味", "代表的投資家のべき効用 U(S) = S^(1−γ)/(1−γ) の相対的リスク回避度。オプション価格が内包する保険料の大きさを一つの数に要約する"],
           ["主指標", "対数プライシング・カーネル ln(f_Q/f_P) の 2 次式当てはめ（中央 5〜95%、重み = f_Q 質量）の <b>傾き γ̂_q</b>（ATM 近傍の局所リスク回避度）と <b>曲率 β₂</b>（両裾の保険料 = ボラティリティ・プレミアム）。直線の γ̂ は参考値"],
           ["推定法", "I. カーネル法（リスク中立密度 ÷ 歴史的密度の対数を x に回帰。主指標は 2 次式）、II. モーメント法（歪度と尖度の関係）、III. 最尤法（実現 SQ 値の尤度）"],
           ["データ", "QRI スナップショット（R2、2026-06〜）、日経平均日次終値（2000-01〜、Yahoo Finance）、JPX 最終清算数値（sq_his.pdf）"],
           ["作成", f"{date.today():%Y-%m-%d}、先物手口分析ワークスペース（scripts/gamma_calc.py, scripts/build_gamma_report.py）"]], [30, 148]),
      sp(5), Pg("要約", H1),
      Pg(f"・主指標（2 次式の傾き γ̂_q）の日次系列は平均 {rel['mean']:+.2f}、中央値 {rel['median']:+.2f}、10〜90% 点 {rel['q10']:+.2f}〜{rel['q90']:+.2f}、直近（{last.day[4:6]}/{last.day[6:]}）{last.gamma_q:+.2f}。"
         f"文献の代表値（Bliss–Panigirtzoglou の 4 前後）より小さい。参考の直線 γ̂ は平均 {rel['lin']['mean']:+.2f}（主指標との相関 {rel['corr_q_lin']:+.2f}）、方法 II の平均は {rel['mean_II']:.2f}。"),
      Pg(f"・曲率 β₂ は平均 {rel['b2']['mean']:+.1f}、中央値 {rel['b2']['median']:+.1f}（10〜90% 点 {rel['b2']['q10']:+.1f}〜{rel['b2']['q90']:+.1f}、直近 {last.beta2:+.1f}）。符号は局面で変わり、ATM IV > RV₂₀ のとき正（U 字：中心付近では保険料が小さく両裾で急に大きくなる）、急落直後に RV₂₀ が IV を上回ると負（山型）になる。"
         f"β₂ は ATM IV − RV₂₀ と相関 {rel['b2']['corr']['ivrv']:+.2f} で動き、正規近似の β₂ = ½(1/s_P² − 1/s_Q²) が示す通り「リスク中立分布がヒストリカルより広い」＝ボラティリティ・プレミアムの密度表現である。"
         "直線の γ̂ は U 字のどこを重く見るかで値が変わる（重み依存）ため参考値に格下げした。"),
      Pg(f"・原資産との関係：γ̂_q は直近 20 日リターンと相関 {rel['corr']['ret20']:+.2f}、実現ボラティリティと {rel['corr']['rv20']:+.2f}、ATM IV と {rel['corr']['atm']:+.2f}。曲率 β₂ は 20 日リターンと {rel['b2']['corr']['ret20']:+.2f}、RV₂₀ と {rel['b2']['corr']['rv20']:+.2f}。"
         f"先行性（γ̂_q と次の 5・10・20 営業日のリターンの相関）は {rel['fwd_corr']['5']:+.2f} / {rel['fwd_corr']['10']:+.2f} / {rel['fwd_corr']['20']:+.2f}、β₂ では {rel['b2']['fwd_corr']['5']:+.2f} / {rel['b2']['fwd_corr']['10']:+.2f} / {rel['b2']['fwd_corr']['20']:+.2f} で、標本が少なく結論は出せない。"),
      Pg(f"・精度：最終日のブロック・ブートストラップでは γ̂_q の標準誤差 {boot['sd']:.2f}（95% 区間 {boot['lo']:+.2f}〜{boot['hi']:+.2f}）、β₂ の標準誤差 {boot['b2']['sd']:.1f}（{boot['b2']['lo']:+.1f}〜{boot['b2']['hi']:+.1f}）。参考の直線 γ̂ は設定の違い（RV60・中央 80%・狭い帯域）で 1 日あたり平均 {rel['variant_spread']:.2f} 動く。"
         f"最尤法（方法 III、月次 {ML['profile_monthly']['n']} 限月、格子 −6〜12）の γ̂ は {ML['profile_all']['gamma_hat']:+.2f} だが、95% 区間は {ML['profile_all']['ci_lo']:+.1f}〜{ML['profile_all']['ci_hi']:+.1f} と格子全体に及び、現時点では情報量が乏しい。週次満期（ミニ）の取り込みを開始し、実現待ち {len(ML['pending'])} 本。"),
      Pg("・目下の方針：主指標 γ̂_q・β₂ を筆頭に、参考値（直線 γ̂ とその変種、方法 II）、F、RV、IV、IV−RV を日次で蓄積し、満期ごとに実現 SQ の尤度（方法 III、週次満期を含む）を積み上げる。γ の水準は方法 III で、形状は γ̂_q・β₂ で決める分担とする（第 8 章）。"),
      sp(3), Pg("構成", H2), Pg("1. 問いの定義　2. 理論　3. データと実装　4. 結果 I：日次系列　5. 結果 II：最尤法　6. 原資産価格との関係　7. 精度の評価と向上策　8. 蓄積の設計　9. 留意点・参考文献", SMALL),
      PageBreak()]

S += [Pg("1. 問いの定義", H1),
      Pg("「市場が織り込む γ」とは、オプション価格から得られるリスク中立分布 f_Q と、実際に起こりうる分布 f_P の差を、代表的投資家のリスク回避度として一つの数に要約したものである。"
         "γ が大きいほど、下落状態の受取が高く評価されており、f_Q は f_P より左に厚い。γ = 0 なら保険料は無く、f_Q = f_P である。"),
      Pg("γ は観測できない f_P に依存するため、推定には f_P を何で近似するかの仮定が要る。本稿は三つの独立な近似（歴史的分布、モーメントの関係、実現値の尤度）で推定し、互いの整合と精度を評価する。"),
      Pg("2. 理論", H1), Pg("2.1 べき効用とプライシング・カーネル", H2),
      Pg("確率的割引因子 M(S) を用いると f_Q(S) = f_P(S)·M(S)/E_P[M] である。べき効用 U(S) = S^(1−γ)/(1−γ) では M(S) ∝ U′(S) = S^(−γ) なので、対数を取ると"),
      EQ(r"\ln\frac{f_{Q}(S)}{f_{P}(S)} = a - \gamma\,\ln S = a' - \gamma\,x,\qquad x = \ln\frac{S}{F}"),
      Pg("となる。すなわち、二つの密度の比の対数は x の直線で、傾きが −γ である。これが方法 I の根拠である。直線から外れる（曲がる）なら、べき効用の仮定が成り立っていない。"),
      Pg("2.2 モーメルの関係（方法 II）", H2),
      Pg("Bakshi, Kapadia and Madan (2003) は、べき効用のもとでリスク中立歪度と実確率のモーメントが近似的に次の関係を満たすことを示した。"),
      EQ(r"\mathrm{Skew}_{Q} \approx \mathrm{Skew}_{P} - \gamma\, s_{P}\,(\mathrm{Kurt}_{P} - 3)\quad\Rightarrow\quad \hat\gamma_{II} = \frac{\mathrm{Skew}_{P} - \mathrm{Skew}_{Q}}{s_{P}\,(\mathrm{Kurt}_{P}-3)}"),
      Pg("s_P、Kurt_P は実確率の対数リターンの標準偏差と尖度である。密度全体ではなく 3 つのモーメントだけを使うため、両翼の推定誤差の影響を受けにくい一方、尖度の推定誤差に敏感である。"),
      Pg("2.3 最尤法（方法 III）", H2),
      Pg("満期 i の予測時点で得たリスク中立密度 f_Q,i を γ で傾けた密度 f_P,i(S; γ) ∝ f_Q,i(S)·S^γ で、実現した SQ 値 S_i の尤度を評価し、"),
      EQ(r"\hat\gamma_{III} = \arg\max_{\gamma}\ \sum_{i}\ln f_{P,i}(S_{i};\gamma)"),
      Pg("とする（Bliss and Panigirtzoglou 2004）。f_P の近似に歴史的分布を使わない点で方法 I・II と独立である。満期ごとに 1 点しか得られないため、精度は標本数（満期の数）で決まる。95% 区間は対数尤度が最大値から 1.92 以内の γ の範囲（尤度比検定）とする。"),
      Pg("3. データと実装", H1), Pg("3.1 リスク中立密度 f_Q", H2),
      Pg("各営業日の期近限月（SQ まで 8 日以上残るもの）について、日中引け（気配時刻 08:45〜15:30:59 の最終スナップショット、両気配 120 銘柄以上）から SVI を当てはめ、密度を x ∈ [−0.6, 0.6] の 6,001 点で計算した（手順は前々稿）。"
         "当てはめ誤差が 3%pt を超える日、密度が負になる日は除外した。2026-07-13 より前の収集分は価格の桁が欠落しているため、先物価格はコールのデルタと IV から逆算した。"),
      Pg("3.2 実確率密度 f_P（方法 I・II）", H2),
      Pg("日経平均の日次終値（2000-01〜）から、残存期間に対応する n 営業日（n = T×250）の対数リターン r_n を作り、各期間の開始時点の実現ボラティリティ RV₂₀ で割って標準化する。これを現在の RV₂₀ で再スケールした標本 x_j = (r_n,j / RV₂₀,j) × RV₂₀,now を、"
         "ガウス・カーネル密度（Scott の帯域）で平滑化して f_P(x) とする。標準化により、ボラティリティ水準の違う過去の期間を現在の局面に揃えつつ、多日リターンの歪みと厚い裾（経験分布）を保つ。"),
      Pg("3.3 回帰（方法 I）", H2),
      Pg("リスク中立の累積確率が 5〜95% の範囲で ln f_Q(x) − ln f_P(x) を x の 2 次式 c − γ_q x + β₂ x² に回帰する（重み = f_Q の質量）。主指標は <b>γ̂_q</b>（1 次項の符号反転、x = 0 すなわち ATM 近傍の局所リスク回避度）と <b>β₂</b>（曲率）である。"
         "直線 c − γx の回帰も行い、参考値 γ̂ として記録する。直線の傾きは U 字のどの範囲を重く見るかで値が変わるため主指標にしない。"),
      Pg("3.4 精度の設計", H2),
      Pg("設定の違いによる変動（RV₆₀ での再スケール、中央 80% の範囲、狭い帯域）を毎日計算し、幅を記録する。最終日についてはブロック・ブートストラップ（ブロック長 = n 日、200 回）で標準誤差を求める。"),
      PageBreak()]

S += [Pg("4. 結果 I：日次系列", H1), fig_series(),
      TBL([["統計量", "主指標 γ̂_q（2 次式の傾き）", "主指標 β₂（曲率）", "参考: 直線 γ̂"],
           ["標本", f"{len(S_)} 営業日（{S_.day.iloc[0][:4]}-{S_.day.iloc[0][4:6]}-{S_.day.iloc[0][6:]} 〜 {S_.day.iloc[-1][:4]}-{S_.day.iloc[-1][4:6]}-{S_.day.iloc[-1][6:]}）", "", ""],
           ["平均 / 中央値 / 標準偏差", f"{rel['mean']:+.2f} / {rel['median']:+.2f} / {rel['sd']:.2f}", f"{rel['b2']['mean']:+.1f} / {rel['b2']['median']:+.1f} / {rel['b2']['sd']:.1f}", f"{rel['lin']['mean']:+.2f} / {rel['lin']['median']:+.2f} / {rel['lin']['sd']:.2f}"],
           ["10% 点 / 90% 点", f"{rel['q10']:+.2f} / {rel['q90']:+.2f}", f"{rel['b2']['q10']:+.1f} / {rel['b2']['q90']:+.1f}", f"{rel['lin']['q10']:+.2f} / {rel['lin']['q90']:+.2f}"],
           ["1 日の自己相関", f"{rel['autocorr1']:.2f}", f"{rel['b2']['autocorr1']:.2f}", f"{rel['lin']['autocorr1']:.2f}"],
           ["直線当てはめの決定係数（平均）", "", "", f"{S_.r2.mean():.2f}"],
           ["方法 II の平均 / 相関", f"{rel['mean_II']:.2f} / {rel['corr']['gII']:+.2f}", f"— / {rel['b2']['corr']['gII']:+.2f}", f"— / {rel['lin']['corr']['gII']:+.2f}"],
           [f"直近 {last.day[4:6]}/{last.day[6:]}", f"{last.gamma_q:+.2f}", f"{last.beta2:+.1f}", f"{last.gamma_I:+.2f}（変種 {min(last[VAR]):+.2f}〜{max(last[VAR]):+.2f}）"],
           ["同日の市場", f"F = {fmt(last.F)}、ATM IV {last.atm*100:.1f}%、RV₂₀ {last.rv20*100:.1f}%、IV−RV {last.ivrv*100:+.1f}%pt、γ̂_II = {last.gamma_II:.2f}", "", ""]], [40, 50, 40, 48]),
      Pg("表1  日次系列の要約。", CAP), fig_kernel(),
      Pg(f"図2a は最終日の対数カーネルである。中央部では傾きが緩く、両端で急に立ち上がる。2 次式の傾き γ̂_q = {lat['gamma_q']:+.2f} が ATM 近傍の局所リスク回避度、曲率 β₂ = {lat['beta2']:+.1f} が両裾の保険料である。直線で近似した傾き（γ̂ = {lat['gamma']:+.2f}）は、この曲線の「中央 90% の平均的な傾き」に過ぎない。"
         "この形は米国株でも繰り返し報告されており（Jackwerth 2000、Aït-Sahalia and Lo 2000 の「プライシング・カーネル・パズル」）、べき効用では説明できない。"
         "実務上の含意は、保険料を一つの γ で取り除くと、中央部では取り過ぎ、両裾では取り足りない、ということである。"),
      Pg(f"方法 II（モーメント法）の γ̂ は平均 {rel['mean_II']:.2f} と γ̂_q より{'大きい' if rel['mean_II'] > rel['mean'] else '小さい'}。モーメント法は両裾の厚さ（尖度）の差を γ に押し込めるため、裾の保険料が大きい局面で高く出る。"
         "二つの方法の差そのものが「べき効用で表せない部分」の大きさの目安になる。"),
      PageBreak(),
      Pg("5. 結果 II：最尤法（実現 SQ 値による検証）", H1), fig_mle(),
      TBL([["限月", "予測時点", "残存", "F（予測時）", "実現 SQ", "u = Φ_Q(SQ)", "単独の γ̂"]] +
          [[e["cm"][:2] + "年" + e["cm"][2:] + "月限", f"{e['day'][4:6]}/{e['day'][6:]}", f"{e['T_days']:.0f} 日", fmt(e["F"]), fmt(e["sq"]), f"{e['u']:.2f}", f"{e['gamma_single']:+.1f}"] for e in ML["monthly"]] +
          [[f"週次 SQ {e['sq_day'][4:6]}/{e['sq_day'][6:]}", f"{e['forecast_day'][4:6]}/{e['forecast_day'][6:]}", f"{e['T_days']:.1f} 日", fmt(e["F"]), fmt(e["sq"]), f"{e['u']:.2f}", f"{e['gamma_single']:+.1f}"] for e in ML["weekly"]], [30, 24, 20, 28, 28, 26, 22]),
      Pg(f"表2  満期ごとの実現値。u はリスク中立分布における実現値の分位（0.5 が中央）。単独の γ̂ は格子の端（−6 または +12）に張り付く。"
         f"合計 {ML['profile_all']['n']} 点の γ̂ = {ML['profile_all']['gamma_hat']:+.2f}、95% 区間 {ML['profile_all']['ci_lo']:+.1f}〜{ML['profile_all']['ci_hi']:+.1f}。γ = 0 と最大値の対数尤度差は {ML['profile_monthly']['ll_at_0']-ML['profile_monthly']['llmax']:+.2f}。", CAP),
      Pg("格子を負側（−6）まで広げても結論は変わらない。上で実現した満期は γ を上げたがり、下で実現した満期は下げたがるため、合計の尤度は平坦で、γ = 0 も区別できない。"
         "方法 III は蓄積によって初めて精度が出る指標であり、現時点では方法 I・II の外部検証として位置づける。"),
      Pg(f"<b>週次満期の取り込み。</b>日経225ミニオプション（水・金 SQ）は 2026-10-20 から IV を収集しており、満期ごとに「その満期の両気配が 60 本以上ある最も早い営業日の引け（残存 1.5 日以上）」を予測時点とする。"
         f"実現値は JPX の週次 SQ 履歴 PDF から取る。現在、予測を作成済みで実現待ちの満期は {len(ML['pending'])} 本（{', '.join('SQ ' + x['sq_day'][4:6] + '/' + x['sq_day'][6:] for x in ML['pending'])}）。"
         "月次の年 12 点に対し週次は年 100 点前後になるため、区間の縮小は週次の蓄積で決まる。日次ジョブ（gamma_mle.py）が毎営業日に更新する。"),
      Pg("6. 原資産価格との関係", H1), fig_relation(),
      TBL([["相関", "γ̂_q（2 次式の傾き）", "β₂（曲率）", "参考: 直線 γ̂"],
           ["ln F（水準）", f"{rel['corr']['lnF']:+.2f}", f"{rel['b2']['corr']['lnF']:+.2f}", f"{rel['lin']['corr']['lnF']:+.2f}"],
           ["直近 20 日の対数リターン", f"{rel['corr']['ret20']:+.2f}", f"{rel['b2']['corr']['ret20']:+.2f}", f"{rel['lin']['corr']['ret20']:+.2f}"],
           ["RV₂₀", f"{rel['corr']['rv20']:+.2f}", f"{rel['b2']['corr']['rv20']:+.2f}", f"{rel['lin']['corr']['rv20']:+.2f}"],
           ["ATM IV", f"{rel['corr']['atm']:+.2f}", f"{rel['b2']['corr']['atm']:+.2f}", f"{rel['lin']['corr']['atm']:+.2f}"],
           ["ATM IV − RV₂₀", f"{rel['corr']['ivrv']:+.2f}", f"{rel['b2']['corr']['ivrv']:+.2f}", f"{rel['lin']['corr']['ivrv']:+.2f}"],
           ["次の 5 / 10 / 20 営業日のリターン（先行性）", f"{rel['fwd_corr']['5']:+.2f} / {rel['fwd_corr']['10']:+.2f} / {rel['fwd_corr']['20']:+.2f}", f"{rel['b2']['fwd_corr']['5']:+.2f} / {rel['b2']['fwd_corr']['10']:+.2f} / {rel['b2']['fwd_corr']['20']:+.2f}", f"{rel['lin']['fwd_corr']['5']:+.2f} / {rel['lin']['fwd_corr']['10']:+.2f} / {rel['lin']['fwd_corr']['20']:+.2f}"],
           ["回帰", f"γ̂_q = {rel['reg']['const']:+.2f} {rel['reg']['ret20']:+.2f}·ret20 {rel['reg']['rv20']:+.2f}·RV₂₀", f"β₂ = {rel['reg_b2']['const']:+.1f} {rel['reg_b2']['ivrv']:+.2f}·(IV−RV)[%pt]", ""]], [52, 48, 44, 34]),
      Pg(f"表3  主指標と原資産・ボラティリティの関係（2026-06〜10、n = {rel['fwd_n']['5']} / {rel['fwd_n']['10']} / {rel['fwd_n']['20']}。重複標本のため相関の有意性は割り引いて読む）。", CAP),
      Pg("理論的には、習慣形成モデル（Campbell and Cochrane 1999）のように下落後にリスク回避が高まるなら、γ̂_q は 20 日リターンと負、ボラティリティと正の相関を持つ。"
         f"標本では ret20 との相関が {rel['corr']['ret20']:+.2f}、RV₂₀ との相関が {rel['corr']['rv20']:+.2f} である。ただし f_P を現在の RV で再スケールしているため、ボラティリティが急上昇した直後は f_P が広がって γ̂_q が機械的に下がる側面がある（第 7 章）。"
         f"曲率 β₂ は IV − RV と相関 {rel['b2']['corr']['ivrv']:+.2f} で、正規近似 β₂ = ½(1/s_P² − 1/s_Q²) の通りボラティリティ・プレミアムの密度表現である。価格との関係を蓄積する際は、主指標と同時に RV₂₀、ATM IV、IV−RV を保存し、ボラティリティ要因を制御した上で評価する。"),
      PageBreak()]

S += [Pg("7. 精度の評価と向上策", H1), Pg("7.1 誤差の源", H2),
      TBL([["源", "影響", "対策"],
           ["f_P のスケール（どのボラティリティで再スケールするか）", f"RV₂₀ と RV₆₀ の違いだけで γ̂ は 1 日平均 {abs(S_.gamma_I - S_.gamma_rv60).mean():.2f} 動く。急変動直後に最も大きい", "GARCH 型の条件付き分散予測、または IV の一定割合（IV×RV/IV の長期平均）でスケールする"],
           ["f_P の形（歴史的分布の期間・標本）", "2008 年・2020 年を含むかで裾が変わる", "期間を分けた感度分析。将来は週次満期の実現値で f_P を直接検証"],
           ["中央範囲の選び方", f"90% と 80% で平均 {abs(S_.gamma_I - S_.gamma_c80).mean():.2f} の差", "範囲を固定して「変化」を見る。曲率（2 次項）を併記"],
           ["f_Q の両翼（SVI の補外・気配の幅）", "5% 点より外で密度の誤差が大きい", "中央範囲に限定（実施済み）。深い OTM の気配幅で重み付け"],
           ["気配時刻", "引け前後で IV が動く。夜間の気配を混ぜると不連続", "気配時刻 15:30:59 以前の最終に統一（実施済み）。収集を 16:15 まで延長（実施済み）"],
           ["標本数（方法 III）", f"{ML['profile_monthly']['n']} 限月では区間が格子全体に及ぶ", "月次限月の蓄積（年 12 点）。週次満期の IV を収集できれば年 100 点超"]], [52, 62, 64]),
      Pg("7.2 ブートストラップによる標準誤差（最終日）", H2),
      Pg(f"f_P の標本（歴史的 n 日リターン）をブロック長 {boot['block']} 日で {boot['n']} 回再抽出して主指標を再計算した結果、γ̂_q は平均 {boot['mean']:+.2f}、標準誤差 {boot['sd']:.2f}、95% 区間 {boot['lo']:+.2f}〜{boot['hi']:+.2f}。"
         f"β₂ は平均 {boot['b2']['mean']:+.1f}、標準誤差 {boot['b2']['sd']:.1f}、95% 区間 {boot['b2']['lo']:+.1f}〜{boot['b2']['hi']:+.1f}（参考の直線 γ̂ は {boot['lin']['lo']:+.2f}〜{boot['lin']['hi']:+.2f}）。"
         "歴史的分布の標本誤差だけでこの幅があり、設定の違い（7.1）を加えると不確実性はさらに大きい。日次の γ̂_q は小数第 1 位までの精度と考える。"),
      Pg("7.3 向上策の優先順位", H2),
      Pg("(1) 条件付き分散の予測で f_P をスケールする（RV₂₀ の機械的な遅れを除く）。(2) 方法 I・II・III を同じ日付で並べ、三者の乖離を監視する。(3) γ を「ATM の傾き γ̂_q」と「裾の増分 β₂」の二つの数に分けて記録する（2026-10-21 実施、主指標化）。"
         "(4) 満期ごとの実現値を蓄積し、方法 III の区間を縮める。(5) 週次満期（ミニ）の IV 収集（2026-10-20 開始）で方法 III の点数を年 100 本規模にする。"),
      Pg("8. 蓄積の設計", H1),
      TBL([["項目", "内容"],
           ["頻度・時点", "各営業日の日中引け（気配時刻 15:30:59 以前の最終スナップショット）。夜間立会の値は別列に参考保存"],
           ["主指標の列", "gamma_q（2 次式の傾き = ATM 近傍の局所 γ）、beta2（曲率 = 両裾の保険料）、ivrv（ATM IV − RV₂₀）"],
           ["参考・材料の列", "day, cm, ts, key, F, T, atm, rr25_delta, 分位点, p（SVI 5 パラメータ）, rmse, gamma_I, gamma_rv60, gamma_c80, gamma_bw, r2, gamma_II, s_q, s_p, skew_q, kurt_q, skew_p, kurt_p, rv20, rv60, ret20, tilt（γ = 0, 2, 4, 6 の要約）"],
           ["満期ごとの記録", "cm, 予測日, T, F, 分位点, 実現 SQ, u = Φ_Q(SQ), ln f_P(SQ; γ) の γ 格子（方法 III の積み上げ用）"],
           ["保存先", "R2 qri_iv/derived/gamma_daily.jsonl（scripts/gamma_daily.py、GitHub Actions の日次ジョブ）。方法 III は gamma_mle.json（scripts/gamma_mle.py）。本稿の全系列再計算は scripts/gamma_calc.py"],
           ["検証計画", "3 か月ごとに：γ̂ の分布、価格・ボラとの相関、方法 III のプロファイル尤度、PIT の一様性（Berkowitz 検定）を更新"]], [34, 144]),
      CondPageBreak(60 * mm), Pg("9. 留意点", H1),
      Pg("・γ̂_q は「ATM 近傍での局所的な傾き」、β₂ は「両裾での保険料の増分」であり、いずれも投資家のリスク回避度そのものではない。β₂ が大きい局面では一つの γ で表すこと自体に無理があるため、直線の γ̂ は参考値に留める。"),
      Pg("・f_P は歴史的分布の再スケールという近似で、構造変化（市場制度、投資家構成）を反映しない。"),
      Pg("・4 か月の標本で、急落（8/19、10/15）が 2 回含まれる。統計量はこれらの日に引きずられる。"),
      sp(3), Pg("参考文献", H2),
      Pg("Aït-Sahalia, Y. and A. Lo (2000), Nonparametric Risk Management and Implied Risk Aversion, Journal of Econometrics 94.", SMALL),
      Pg("Bakshi, G., N. Kapadia and D. Madan (2003), Stock Return Characteristics, Skew Laws, and the Differential Pricing of Individual Equity Options, Review of Financial Studies 16.", SMALL),
      Pg("Bliss, R. and N. Panigirtzoglou (2004), Option-Implied Risk Aversion Estimates, Journal of Finance 59.", SMALL),
      Pg("Campbell, J. and J. Cochrane (1999), By Force of Habit: A Consumption-Based Explanation of Aggregate Stock Market Behavior, Journal of Political Economy 107.", SMALL),
      Pg("Jackwerth, J. (2000), Recovering Risk Aversion from Option Prices and Realized Returns, Review of Financial Studies 13.", SMALL),
      Pg("Rosenberg, J. and R. Engle (2002), Empirical Pricing Kernels, Journal of Financial Economics 64.", SMALL),
      Pg("Shiratsuka, S. (2001), Information Content of Implied Probability Distributions: Empirical Studies on Japanese Stock Price Index Options, Monetary and Economic Studies 19(3), Bank of Japan.", SMALL)]

doc = SimpleDocTemplate(str(PDF_PATH), pagesize=A4, leftMargin=16 * mm, rightMargin=16 * mm, topMargin=16 * mm, bottomMargin=18 * mm, title="市場が織り込むリスク回避度 γ の推定", author="先物手口分析")
ft = make_footer(f"市場が織り込むリスク回避度 γ の推定（データ 〜{last.day[:4]}-{last.day[4:6]}-{last.day[6:]} 引け、作成 {date.today():%Y-%m-%d}）")
doc.build(S, onFirstPage=ft, onLaterPages=ft)
print("written", PDF_PATH)
