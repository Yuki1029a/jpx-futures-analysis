"""「市場が織り込むリスク回避度 γ」レポート（PDF）: 実現 SQ 値による最尤推定（方法 III）の蓄積状況と精度の見通し。

2026-10-22 の方針変更: 日次横断面の γ̂（カーネル法・モーメント法）は f_P の仮定とボラ水準に支配されるため掲載しない。
入力: gamma_mle.py の出力 gamma_mle.json（<計算結果ディレクトリ>）と日次蓄積 cache/derived/gamma_daily.jsonl。
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
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import CondPageBreak, PageBreak, SimpleDocTemplate

from report_kit import B, CAP, EQ, H1, H2, INK, MUT, O, SMALL, SUB, TBL, TITLE, Pg, fig_to_img, fmt, make_footer, sp, style_ax
from rnd_lib import svi

CALC = Path(sys.argv[1]); OUT_DIR = Path(sys.argv[2]); OUT_DIR.mkdir(parents=True, exist_ok=True)
PDF_PATH = OUT_DIR / "nk225op_market_implied_gamma_report.pdf"
ML = json.load(open(CALC / "gamma_mle.json"))
STORE = [json.loads(l) for l in open(Path(__file__).resolve().parent.parent / "cache" / "derived" / "gamma_daily.jsonl", encoding="utf-8") if l.strip()]
pm, pw, pa = ML["profile_monthly"], ML["profile_weekly"], ML["profile_all"]
LAST = STORE[-1]["day"]; FIRST = STORE[0]["day"]
D = lambda d: f"{d[:4]}-{d[4:6]}-{d[6:]}"

# ---------------------------------------------------------------- 情報量（フィッシャー情報 ≈ Var_Q(x) ≈ s²）
def _atm_from_p(p, T):
    return float(np.sqrt(max(svi(np.array(p), np.array([0.0]))[0][0], 1e-10) / T))

info_m = [(e["cm"], e["atm"] ** 2 * e["T_days"] / 365) for e in ML["monthly"]]
info_w = {e["code"]: _atm_from_p(e["p"], e["T_days"] / 365) ** 2 * e["T_days"] / 365 for e in ML["weekly"] + ML["pending"]}
I_now = sum(v for _, v in info_m) + sum(info_w[e["code"]] for e in ML["weekly"])
I_pend = sum(info_w[e["code"]] for e in ML["pending"])
SIG = float(np.mean([r["atm"] for r in STORE]))  # 期近 ATM IV の平均（年率）
se_now = 1 / np.sqrt(I_now) if I_now > 0 else float("inf")
YEARS = [1, 3, 5, 10, 15, 20, 25]
se_years = {y: 1 / np.sqrt(SIG ** 2 * y) for y in YEARS}


def fig_mle():
    g = np.array(ML["grid"]); fig, a = plt.subplots(figsize=(7.2, 3.0)); ll = np.array(pm["ll"]); a.plot(g, ll - ll.max(), color=B, lw=2, label=f"月次限月（{pm['n']} 点、残存 28 日前後の予測）")
    if pw:
        lw_ = np.array(pw["ll"]); a.plot(g, lw_ - lw_.max(), color=O, lw=1.6, ls="--", label=f"週次満期（{pw['n']} 点）")
        la = np.array(pa["ll"]); a.plot(g, la - la.max(), color=INK, lw=1.2, label=f"合計（{pa['n']} 点）")
    a.axhline(-1.92, color=MUT, lw=0.8, ls=":"); a.text(g[-1], -1.8, "95% 区間の閾値（−1.92）", ha="right", fontsize=7.5, color=MUT)
    a.axvline(pa["gamma_hat"], color=B, lw=0.8); a.axvline(0, color=MUT, lw=0.6); a.set_ylim(-4, 0.3)
    style_ax(a, "図1  方法 III のプロファイル対数尤度（最大値からの差、格子 −6〜12）", "γ", "対数尤度の差"); a.legend(fontsize=7.5, frameon=False, loc="lower left")
    return fig_to_img(fig)


def fig_se():
    fig, a = plt.subplots(figsize=(7.2, 2.6)); ys = np.linspace(0.5, 25, 200); a.plot(ys, 1 / np.sqrt(SIG ** 2 * ys), color=B, lw=2, label=f"γ₀ 単独（σ = {SIG*100:.0f}%）")
    a.plot(ys, np.sqrt(3) / np.sqrt(SIG ** 2 * ys), color=O, lw=1.4, ls="--", label="状態依存 3 母数（目安 √3 倍）")
    a.axhline(1, color=MUT, lw=0.8, ls=":"); a.axvline(I_now / SIG ** 2, color=INK, lw=0.8); a.text(I_now / SIG ** 2 + 0.3, 5.5, f"現在 ≈ {I_now/SIG**2:.2f} 年分", fontsize=8, color=INK)
    a.set_ylim(0, 6); style_ax(a, "図2  実現値の暦年数と γ̂ の標準誤差（SE ≈ 1/√(σ²·年数)）", "予測期間が覆う暦年数", "標準誤差"); a.legend(fontsize=7.5, frameon=False)
    return fig_to_img(fig)


# ---------------------------------------------------------------- 文書
S = []
S += [sp(8), Pg("市場が織り込むリスク回避度 γ", TITLE), Pg("実現 SQ 値による最尤推定（方法 III）の蓄積状況と、精度の見通し", SUB),
      TBL([["項目", "内容"],
           ["対象", f"日経225オプション（月次限月、残存 28 日前後の引け）と日経225ミニオプション（週次満期）。f_Q の日次蓄積 {D(FIRST)} 〜 {D(LAST)}（{len(STORE)} 営業日）"],
           ["γ の意味", "代表的投資家のべき効用 U(S) = S^(1−γ)/(1−γ) の相対的リスク回避度。リスク中立密度 f_Q を S^γ で傾けた密度が実確率密度 f_P になる、という関係の γ"],
           ["推定法", "方法 III（最尤法）のみ。各満期の予測時点の f_Q を γ で傾けた密度で実現 SQ 値の尤度を評価し、満期をまたいで足し合わせる。日次横断面の γ̂（カーネル法・モーメント法）は検証材料として蓄積するが掲載しない（第 6 章）"],
           ["データ", "QRI スナップショット（気配時刻 15:30:59 以前の最終、R2）、JPX 最終清算数値（sq_his.pdf、sq_his(mini,weekly).pdf）"],
           ["作成", f"{date.today():%Y-%m-%d}、先物手口分析ワークスペース（scripts/gamma_mle.py, scripts/build_gamma_report.py）"]], [30, 148]),
      sp(5), Pg("要約", H1),
      Pg(f"・実現済みの満期は月次 {pm['n']} 本（週次 {pw['n'] if pw else 0} 本）。γ̂ = {pa['gamma_hat']:+.2f}、95% 区間 {pa['ci_lo']:+.1f}〜{pa['ci_hi']:+.1f} は格子全体に及び、γ = 0 との対数尤度差は {pm['ll_at_0']-pm['llmax']:+.2f}。現時点で水準について言えることはない。"),
      Pg(f"・精度は点数ではなく暦年数で決まる。1 満期の γ に関する情報は実現値の分散 σ²T に等しく、満期をまたいで足すと σ² × 覆った暦年数になる。現在の情報量は {I_now:.3f}（{I_now/SIG**2:.2f} 年分）で標準誤差 ≈ {se_now:.1f}。"
         f"σ ≈ {SIG*100:.0f}% では 1 年で ≈ {se_years[1]:.1f}、10 年で ≈ {se_years[10]:.1f}、20 年で ≈ {se_years[20]:.2f}。週次満期は同じ暦期間を月次と重複して覆うため、点数の割に情報は増えない。"),
      Pg("・従って前向きの蓄積だけでは実用精度（標準誤差 1 以下）に 10 年以上かかる。JPX のヒストリカル・データ（日経225オプションの日次清算値、2000 年以降）で方法 III を遡及計算すれば、購入直後に 20 年分（標準誤差 ≈ 0.7）が得られる（第 5 章）。"),
      Pg("・日次横断面の γ̂（リスク中立密度 ÷ 歴史的密度の傾き）は、歴史的密度の平均・幅の仮定とボラティリティ水準に支配され、選好が不変でも ±1 動く。市場のリスク回避度の水準としては報告せず、方法 III の検証材料として列を保持する（別稿「プライシングカーネル・パズルと状態依存 γ」§2.3）。"),
      sp(3), Pg("構成", H2), Pg("1. 問いの定義　2. 方法 III の理論と情報量　3. データ　4. 結果　5. 精度の見通しと遡及データ　6. 日次蓄積の位置づけ　7. 留意点・参考文献", SMALL),
      PageBreak()]

S += [Pg("1. 問いの定義", H1),
      Pg("「市場が織り込む γ」とは、オプション価格から得られるリスク中立分布 f_Q と、実際に起こる分布 f_P の差を、代表的投資家のリスク回避度として一つの数に要約したものである。γ が大きいほど下落状態の受取が高く評価されており、f_Q は f_P より左に厚い。γ = 0 なら f_Q = f_P である。"),
      Pg("f_P は観測できない。文献（Bliss and Panigirtzoglou 2004 ほか）が用いる唯一の直接的な方法は、f_Q を γ で傾けた密度が実現値をどれだけ当てたかを長期の実現系列で評価することであり、本稿の方法 III はこれに当たる。1 日の横断面（f_Q と歴史的分布の比）から γ を出す方法は、歴史的分布の仮定に支配されるため採らない。"),
      Pg("2. 方法 III の理論と情報量", H1), Pg("2.1 べき効用による傾け", H2),
      Pg("確率的割引因子 M(S) を用いると f_Q(S) = f_P(S)·M(S)/E_P[M] である。べき効用では M(S) ∝ S^(−γ) なので、x = ln(S/F) を使って"),
      EQ(r"f_P(x;\gamma) = \frac{f_Q(x)\,e^{\gamma x}}{\int f_Q(u)\,e^{\gamma u}\,du}"),
      Pg("と書ける。γ は指数型の傾け母数で、x がその十分統計量である。"),
      Pg("2.2 最尤推定と区間", H2),
      Pg("満期 i の予測時点で得た f_Q,i を γ で傾け、実現した SQ 値 x_i の対数尤度を満期をまたいで足す。"),
      EQ(r"\hat\gamma_{III} = \mathrm{arg\,max}_{\gamma}\ \sum_{i}\ln f_{P,i}(x_{i};\gamma),\qquad \ln f_{P,i}(x;\gamma) = \ln f_{Q,i}(x) + \gamma x - \ln\int f_{Q,i}(u)e^{\gamma u}du"),
      Pg("95% 区間は対数尤度が最大値から 1.92 以内の γ の範囲（尤度比検定）。格子は −6〜12 の 0.25 刻みである。"),
      Pg("2.3 1 満期が持つ情報量", H2),
      Pg("対数尤度を γ で微分したスコアは x_i − E_γ[x]、その分散がフィッシャー情報である。指数型なので"),
      EQ(r"I_i(\gamma) = \mathrm{Var}_{\gamma}(x) \approx \mathrm{Var}_{Q}(x) = s_i^{2} \approx \sigma_i^{2}\,T_i,\qquad \mathrm{SE}(\hat\gamma) \approx \frac{1}{\sqrt{\sum_i s_i^2}}"),
      Pg("すなわち 1 満期の情報は、その満期の予測期間 T_i にリスク中立分散 σ_i² を掛けたものに等しい。満期が暦の上で重ならないように並べると Σσ_i²T_i ≈ σ² × 覆った暦年数となる。"),
      Pg(f"数値例: 残存 28 日、ATM IV 30% の月次 1 本は s² = 0.30² × 28/365 = {0.09*28/365:.4f}。年 12 本で {0.09*28/365*12:.3f}、標準誤差 1/√{0.09*28/365*12:.3f} ≈ {1/np.sqrt(0.09*28/365*12):.1f}。"
         f"残存 3 日の週次 1 本は s² = {0.09*3/365:.5f} で、月次 1 本の 1/9 に過ぎず、しかも同じ暦期間を月次と重複して覆う（実現パスが共通）ため情報は加算されない。"),
      Pg("<b>含意。</b>γ の精度を決めるのは実現値の点数ではなく、予測期間が覆う暦年数である。週次満期の取り込みは短い残存での検証（ボラのスケール依存）には役立つが、水準の精度にはほぼ寄与しない。"),
      PageBreak(),
      Pg("3. データ", H1),
      Pg("・f_Q：各営業日の日中引け（気配時刻 08:45〜15:30:59 の最終スナップショット。15:30 以降はマーケットメイカーが気配を引く）から SVI を当てはめ、Breeden–Litzenberger で密度化。当てはめ誤差 3%pt 超・密度が負の日は除外。<br/>"
         "・月次の予測時点：各限月について残存 28 日に最も近い営業日の引け。<br/>"
         "・週次（ミニ）の予測時点：その満期の両気配が 60 本以上ある最も早い営業日の引け（残存 1.5 日以上）。2026-10-20 収集開始。<br/>"
         "・実現値：JPX 最終清算数値（月次 sq_his.pdf、週次 sq_his(mini,weekly).pdf）。<br/>"
         "・日次ジョブ gamma_daily.py が f_Q（SVI 母数・F・T・分位点）を蓄積し、gamma_mle.py が実現値の追加を検知して尤度を更新する（GitHub Actions、JST 18:00 以降）。"),
      Pg("4. 結果", H1), fig_mle(),
      TBL([["満期", "予測時点", "残存", "F（予測時）", "実現 SQ", "u = Φ_Q(SQ)", "単独の γ̂", "情報 s²"]] +
          [[e["cm"][:2] + "年" + e["cm"][2:] + "月限", f"{e['day'][4:6]}/{e['day'][6:]}", f"{e['T_days']:.0f} 日", fmt(e["F"]), fmt(e["sq"]), f"{e['u']:.2f}", f"{e['gamma_single']:+.1f}", f"{v:.4f}"] for e, (_, v) in zip(ML["monthly"], info_m)] +
          [[f"週次 SQ {e['sq_day'][4:6]}/{e['sq_day'][6:]}", f"{e['forecast_day'][4:6]}/{e['forecast_day'][6:]}", f"{e['T_days']:.1f} 日", fmt(e["F"]), fmt(e["sq"]), f"{e['u']:.2f}", f"{e['gamma_single']:+.1f}", f"{info_w[e['code']]:.4f}"] for e in ML["weekly"]] +
          [[f"週次 SQ {e['sq_day'][4:6]}/{e['sq_day'][6:]}（実現待ち）", f"{e['forecast_day'][4:6]}/{e['forecast_day'][6:]}", f"{e['T_days']:.1f} 日", fmt(e["F"]), "—", "—", "—", f"{info_w[e['code']]:.4f}"] for e in ML["pending"]],
          [36, 20, 18, 26, 26, 24, 20, 20]),
      Pg(f"表1  満期ごとの実現値と情報量。u はリスク中立分布における実現値の分位（0.5 が中央）。単独の γ̂ は格子の端に張り付く。合計 n = {pa['n']}、γ̂ = {pa['gamma_hat']:+.2f}、95% 区間 {pa['ci_lo']:+.1f}〜{pa['ci_hi']:+.1f}、情報量 Σs² = {I_now:.3f}（標準誤差 ≈ {se_now:.1f}）。実現待ちが加わると +{I_pend:.4f}。", CAP),
      Pg("上で実現した満期は γ を上げたがり、下で実現した満期は下げたがるため、合計の尤度は平坦で γ = 0 も区別できない。これは標本の少なさそのものであり、方法の欠陥ではない。"),
      CondPageBreak(70 * mm),
      Pg("5. 精度の見通しと遡及データ", H1), fig_se(),
      TBL([["覆う暦年数", "1", "3", "5", "10", "15", "20", "25"], ["γ₀ 単独の標準誤差"] + [f"{se_years[y]:.2f}" for y in YEARS], ["状態依存 3 母数（目安）"] + [f"{se_years[y]*np.sqrt(3):.2f}" for y in YEARS]], [44, 19, 19, 19, 19, 19, 19, 19]),
      Pg(f"表2  σ = {SIG*100:.0f}%（期近 ATM IV の平均）のときの標準誤差。3 母数は説明変数の分散次第で目安。", CAP),
      Pg("前向きの蓄積では、標準誤差 1 に達するまで 10 年以上を要する。実用的な経路は JPX のヒストリカル・データによる遡及で、方法 III に必要なものは次の通りである。"),
      TBL([["項目", "必要性", "内容"],
           ["日経225オプションの日次価格（全限月・全行使）", "必須", "各営業日の清算値（できれば引けの最良気配 bid/ask）。残存 28 日前後の 1 断面を限月ごとに取り、SVI → f_Q を作る。2000 年以降（約 300 限月）が理想、最低でも 10 年"],
           ["日経225先物の日次清算値（同一限月）", "必須", "F の決定（パリティでも代替可だが先物があれば直接）"],
           ["最終清算数値（SQ）", "入手済み", "JPX 公開 PDF（sq_his.pdf）。無料"],
           ["ウィークリーオプション（2015-05〜）の日次価格", "任意", "短い残存での検証用。水準の精度にはほぼ寄与しない（2.3 節）"],
           ["板情報（ティック・気配）", "不要", "引けの断面だけで足りる。15:30 の気配時刻ルールは清算値には適用できないため、清算値が引けの実勢とどれだけ乖離するかを当方の 2026-06〜 の気配データで検証してから使う"]],
          [52, 20, 106]),
      Pg("表3  遡及計算に必要なデータ。清算値はモデル値で埋められる行使価格（取引の無い OTM）を含むため、密度の裾は気配ベースより滑らかに出る。中央部（5〜95%）で評価する方法 III への影響は小さいが、取得後に 2026 年分で気配ベースと突き合わせる。", CAP),
      Pg("遡及で得られるもの: (1) γ₀ の水準（20 年で標準誤差 ≈ 0.7）、(2) 局面別の γ（高ボラ・低ボラ、上昇後・下落後）、(3) 2008 年・2020 年を含む裾の検証（PIT・Berkowitz 検定）。(2)(3) は前向きの蓄積では事実上得られない。"),
      Pg("6. 日次蓄積の位置づけ", H1),
      Pg("日次ジョブは f_Q（SVI 母数、F、T、ATM IV、分位点、γ = 0/2/4/6 で傾けた要約）を蓄積する。これが方法 III の予測時点の材料であり、定例レポートの実確率の帯（γ = 0 と 4 の併記）の根拠でもある。"
         "同時に計算している横断面の γ̂（gamma_q, beta2, gamma_I, gamma_II）は、f_P を歴史的リターンの RV スケールで作るため、選好が不変でもボラ水準で ±1 動く（別稿 §2.3 の実験）。これらは市場のリスク回避度として報告せず、方法 III の γ̂ が得られた時点で「横断面の γ̂ がどの程度ずれるか」を検証する材料として列を保持する。"),
      CondPageBreak(50 * mm), Pg("7. 留意点", H1),
      Pg("・方法 III の γ は「べき効用が正しいとしたときの最良の 1 母数近似」である。対数カーネルが U 字（別稿）なら、γ̂_III は裾と中央の保険料を平均した値になる。"),
      Pg("・予測時点を残存 28 日に固定しているため、γ̂_III は 1 か月先の分布に関する値である。残存依存（Bliss–Panigirtzoglou は 1 週間で約 8、6 週間で約 2）は遡及データで初めて検証できる。"),
      Pg("・清算値ベースの遡及と気配ベースの前向き蓄積を混ぜるときは、2026 年分で両者の f_Q を比較して系統差を確認する。"),
      sp(3), Pg("参考文献", H2),
      Pg("Bliss, R. and N. Panigirtzoglou (2004), Option-Implied Risk Aversion Estimates, Journal of Finance 59. / Berkowitz, J. (2001), Testing Density Forecasts, with Applications to Risk Management, Journal of Business & Economic Statistics 19. / "
         "Jackwerth, J. (2000), Recovering Risk Aversion from Option Prices and Realized Returns, Review of Financial Studies 13. / Rosenberg, J. and R. Engle (2002), Empirical Pricing Kernels, Journal of Financial Economics 64. / "
         "Cuesdeanu, H. and J. Jackwerth (2018), The Pricing Kernel Puzzle: Survey and Outlook, Annals of Finance 14.", SMALL)]

doc = SimpleDocTemplate(str(PDF_PATH), pagesize=A4, leftMargin=16 * mm, rightMargin=16 * mm, topMargin=16 * mm, bottomMargin=16 * mm, title="市場が織り込むリスク回避度 γ（方法 III）")
foot = make_footer(f"市場が織り込むリスク回避度 γ：方法 III の蓄積状況（データ 〜{D(LAST)} 引け、作成 {date.today():%Y-%m-%d}）")
doc.build(S, onFirstPage=foot, onLaterPages=foot); print("written", PDF_PATH)
