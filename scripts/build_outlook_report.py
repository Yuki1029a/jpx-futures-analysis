"""11月限 SQ値の織り込みの推移（リスク中立分布の日次系列）と、先物の先行き・投資家見通しを抽出する視点のリサーチをまとめた PDF。

入力: scripts/rnd_series.py の出力 JSON。
使い方: python scripts/build_outlook_report.py <series.json> <出力フォルダ>
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

SER = json.load(open(sys.argv[1])); OUT_DIR = Path(sys.argv[2]); OUT_DIR.mkdir(parents=True, exist_ok=True)
PDF_PATH = OUT_DIR / "nk225op_2611_rnd_evolution_outlook_research.pdf"
df = pd.DataFrame([{"label": r["label"], "ts": pd.Timestamp(r["ts"]), "F": r["F"], "atm": r["atm"] * 100, "rr": r["rr25_delta"], "mode": r["mode"], "mean": r["mean"],
                    "q05": r["q"]["0.05"], "q25": r["q"]["0.25"], "q50": r["q"]["0.5"], "q75": r["q"]["0.75"], "q95": r["q"]["0.95"], "sd": r["sd"], "sk": r["skew"], "kurt": r["kurt"],
                    "p60": r["cdf"]["60000"] * 100, "p65": r["cdf"]["65000"] * 100, "p70": (1 - r["cdf"]["70000"]) * 100, "p72": (1 - r["cdf"]["72000"]) * 100,
                    "rmse": r["rmse"] * 100, "n": r["n_pts"]} for r in SER])
prev = df[df.label == "10/16 引け"].iloc[0]; last = df.iloc[-1]; first = df.iloc[0]


# ------------------------------------------------------------------ 図
def fig_fan():
    fig, a = plt.subplots(figsize=(7.2, 3.6)); xi = np.arange(len(df))
    a.fill_between(xi, df.q05, df.q95, color=B, alpha=0.15, label="5〜95% 点"); a.fill_between(xi, df.q25, df.q75, color=B, alpha=0.35, label="25〜75% 点")
    a.plot(xi, df.q50, color=B, lw=2, marker="o", ms=4, label="中央値"); a.plot(xi, df["mode"], color=O, lw=1.6, marker="^", ms=4, label="最頻値")
    a.plot(xi, df.F, color=INK, lw=1.2, marker="s", ms=4, label="F（先物推定値 = 分布の平均）")
    for lv in (60000, 65000, 70000):
        a.axhline(lv, color=GRID, lw=0.8)
    a.set_xticks(xi); a.set_xticklabels(df.label.str.replace(" ", "\n"), fontsize=8)
    style_ax(a, "図1  リスク中立分位点の推移（11月限、各営業日の引け近傍）", "", "SQ値（円）"); a.legend(fontsize=8, frameon=False, loc="lower left")
    return fig_to_img(fig)


def fig_tails():
    fig, ax = plt.subplots(2, 1, figsize=(7.2, 5.2), sharex=True); xi = np.arange(len(df))
    ax[0].plot(xi, df.p65, color=O, lw=2, marker="o", ms=4, label="P(S_T < 65,000)"); ax[0].plot(xi, df.p60, color=O, lw=1.4, ls="--", marker="o", ms=3, label="P(S_T < 60,000)")
    ax[0].plot(xi, df.p70, color=B, lw=2, marker="s", ms=4, label="P(S_T ≥ 70,000)"); ax[0].plot(xi, df.p72, color=B, lw=1.4, ls="--", marker="s", ms=3, label="P(S_T ≥ 72,000)")
    style_ax(ax[0], "図2  水準別の確率の推移", "", "確率（%）"); ax[0].legend(fontsize=8, frameon=False, ncol=2)
    ax[1].plot(xi, df.atm, color=B, lw=2, marker="o", ms=4, label="ATM IV（%）"); ax[1].plot(xi, df.rr, color=O, lw=2, marker="s", ms=4, label="RR25（%pt、QRI デルタ基準）")
    ax[1].set_xticks(xi); ax[1].set_xticklabels(df.label.str.replace(" ", "\n"), fontsize=8)
    style_ax(ax[1], "", "", "%"); ax[1].legend(fontsize=8, frameon=False)
    return fig_to_img(fig)


def fig_shape():
    fig, a = plt.subplots(figsize=(7.2, 3.0)); xi = np.arange(len(df))
    a.plot(xi, df["sk"], color=B, lw=2, marker="o", ms=4, label="歪度"); a.plot(xi, df.sd / df.F * 100, color=O, lw=2, marker="s", ms=4, label="標準偏差 ÷ F（%）")
    a.axhline(0, color=MUT, lw=0.8); a.set_xticks(xi); a.set_xticklabels(df.label.str.replace(" ", "\n"), fontsize=8)
    style_ax(a, "図3  分布の形：歪度と相対的な幅", "", "値"); a.legend(fontsize=8, frameon=False)
    return fig_to_img(fig)


# ------------------------------------------------------------------ 文書
S = []
d_prev = {c: last[c] - prev[c] for c in ("F", "q05", "q25", "q50", "q75", "q95", "mode", "p65", "p70", "p60", "rr", "atm", "sk")}
d_all = {c: last[c] - first[c] for c in ("F", "q05", "q50", "q95", "mode", "p65", "p70", "rr", "atm", "sk")}
S += [sp(10), Pg("11月限 SQ値の織り込みの推移と、<br/>先物の先行き・投資家見通しを抽出する視点", TITLE),
      Pg("リスク中立分布の日次系列（10/9〜10/17 夜間引け）の更新と、見通し抽出手法のリサーチ", SUB),
      TBL([["項目", "内容"],
           ["対象", "日経225オプション 2026年11月限（SQ 2026-11-13）"],
           ["系列", f"10/9（10月SQ 日）から {last.label}（{last.ts:%m/%d %H:%M}）まで、各営業日の 15:15 以前の最終スナップショットで同一手順により推定"],
           ["前版", "10/16 引け基準の分析レポート（nk225op_2611_sq_rnd_report.pdf）。本書はその推移の更新と、見通し抽出の視点の整理"],
           ["データ", "QRI スナップショット（15 分遅延）、JPX 建玉残高表、日経平均終値。見通し抽出の視点は文献と公表データの調査による"]], [34, 144]),
      sp(6), Pg("要約", H1),
      Pg(f"・前版（10/16 引け）から {last.label}にかけて、先物推定値 F は {fmt(prev.F)} 円から {fmt(last.F)} 円へ {d_prev['F']:+,.0f} 円下がり、分布の中央値は {fmt(prev.q50)} → {fmt(last.q50)} 円（{d_prev['q50']:+,.0f} 円）、"
         f"最頻値は {fmt(prev['mode'])} → {fmt(last['mode'])} 円（{d_prev['mode']:+,.0f} 円）に移った。5% 点は {d_prev['q05']:+,.0f} 円、95% 点は {d_prev['q95']:+,.0f} 円の移動で、下側の裾が上側より大きく広がった。"),
      Pg(f"・P(S_T < 65,000) は {prev.p65:.0f}% → {last.p65:.0f}%、P(S_T ≥ 70,000) は {prev.p70:.0f}% → {last.p70:.0f}%。RR25（25 デルタのプット IV − コール IV）は {prev.rr:+.1f} → {last.rr:+.1f}pt、ATM IV は {prev.atm:.1f}% → {last.atm:.1f}% で、水準の下方修正に加えて下落側の裾の厚みが増した。"),
      Pg(f"・10/9 から通算では F が {d_all['F']:+,.0f} 円（{d_all['F']/first.F*100:+.1f}%）、中央値が {d_all['q50']:+,.0f} 円、5% 点が {d_all['q05']:+,.0f} 円、95% 点が {d_all['q95']:+,.0f} 円動いた。分布全体の下方移動と、左裾の拡大（歪度 {first["sk"]:+.2f} → {last["sk"]:+.2f}）が同時に進んだ 1 週間である。"),
      Pg("・「リスク中立確率から導いた先物の到達価格」は、分布の平均で見る限り常に先物価格 F に等しい（無裁定の帰結）。情報は平均ではなく形にある。最頻値が F より高い一方で 5% 点が大きく下にあるのは、「最も起こりやすいのは小幅高、ただし大幅安の保険が高く買われている」状態を示す。"),
      Pg("・投資家の見通しを抽出する視点は、(A) オプション価格の形状と実確率への補正、(B) ポジション・フロー、(C) サーベイの 3 系統に整理できる。本ワークスペースの既存データで直ちに実装できるのは、リスク中立分布の指標の日次系列と検証（校正）、投資部門別の先物ネット売買、裁定取引残高、CFTC 建玉報告（CME 日経先物）の取り込みである（第 6〜7 章）。"),
      sp(3), Pg("構成", H2),
      Pg("第 I 部 推移の更新：1. 方法　2. 系列　3. 前版との比較と読み取り　4. 「到達価格」の意味／第 II 部 見通し抽出の視点：5. 枠組み　6. 三つの情報源の整理　7. 統合と実装の優先順位　付録", SMALL),
      PageBreak()]

# ---------------- Part I
S += [Pg("第 I 部　推移の更新", H1), Pg("1. 方法", H2),
      Pg("前版と同じ手順を各営業日に適用した。各日の 15:15 以前の最終スナップショット（両気配 150 銘柄以上）から、プット・コール・パリティで先物推定値 F を求め、"
         "OTM 側の両気配 IV に SVI を当てはめ、Gatheral–Jacquier の解析式でリスク中立密度を得て、分位点・最頻値・水準別の確率を計算した。"
         "RR25 は SVI を介さず QRI のデルタから直接補間した値を用いる（SVI 由来の値は当てはめ誤差の影響を受ける日があるため）。"),
      Pg(f"最終点の「{last.label}」は 10/16 の夜間立会の最終スナップショット（10/17 {last.ts:%H:%M}）である。夜間の気配は幅が広く、SVI の当てはめ誤差が {last.rmse:.1f}%pt と日中（0.3〜0.8pt）より大きい。"
         "同夜の 04:00〜06:00 の 10 スナップショットで分位点はおおむね安定している（中央値 64,500〜64,800 円、5% 点 53,000〜54,100 円）が、両翼（5% 点より外）の精度は日中より劣る。10/12 は祝日取引（現物市場は休場）である。"),
      Pg("2. 系列", H2), fig_fan(), CondPageBreak(95 * mm),
      TBL([["時点", "F", "ATM IV", "RR25", "最頻値", "5% 点", "25% 点", "50% 点", "75% 点", "95% 点", "歪度", "P(<65k)", "P(≥70k)", "誤差"]] +
          [[r.label.replace(" ", "<br/>"), fmt(r.F), f"{r.atm:.1f}", f"{r.rr:+.1f}", fmt(r["mode"]), fmt(r.q05), fmt(r.q25), fmt(r.q50), fmt(r.q75), fmt(r.q95), f"{r["sk"]:+.2f}", f"{r.p65:.0f}%", f"{r.p70:.0f}%", f"{r.rmse:.1f}"] for _, r in df.iterrows()],
          [17, 13, 11, 10, 13, 13, 13, 13, 13, 13, 11, 12, 12, 10]),
      Pg("表1  日次のリスク中立分布の要約（円、%、%pt）。誤差 = SVI の IV 二乗平均誤差（%pt）。", CAP),
      fig_tails(), fig_shape(),
      CondPageBreak(60 * mm), Pg("3. 前版との比較と読み取り", H2),
      TBL([["指標", "10/9 引け", "10/16 引け（前版）", f"{last.label}", "前版比", "10/9 比"],
           ["F（円）", fmt(first.F), fmt(prev.F), fmt(last.F), f"{d_prev['F']:+,.0f}", f"{d_all['F']:+,.0f}"],
           ["最頻値（円）", fmt(first["mode"]), fmt(prev["mode"]), fmt(last["mode"]), f"{d_prev['mode']:+,.0f}", f"{d_all['mode']:+,.0f}"],
           ["中央値（円）", fmt(first.q50), fmt(prev.q50), fmt(last.q50), f"{d_prev['q50']:+,.0f}", f"{d_all['q50']:+,.0f}"],
           ["5% 点（円）", fmt(first.q05), fmt(prev.q05), fmt(last.q05), f"{d_prev['q05']:+,.0f}", f"{d_all['q05']:+,.0f}"],
           ["95% 点（円）", fmt(first.q95), fmt(prev.q95), fmt(last.q95), f"{d_prev['q95']:+,.0f}", f"{d_all['q95']:+,.0f}"],
           ["ATM IV（%）", f"{first.atm:.1f}", f"{prev.atm:.1f}", f"{last.atm:.1f}", f"{d_prev['atm']:+.1f}", f"{d_all['atm']:+.1f}"],
           ["RR25（pt）", f"{first.rr:+.1f}", f"{prev.rr:+.1f}", f"{last.rr:+.1f}", f"{d_prev['rr']:+.1f}", f"{d_all['rr']:+.1f}"],
           ["歪度", f"{first["sk"]:+.2f}", f"{prev["sk"]:+.2f}", f"{last["sk"]:+.2f}", f"{d_prev['sk']:+.2f}", f"{d_all['sk']:+.2f}"],
           ["P(S_T < 60,000)", f"{first.p60:.1f}%", f"{prev.p60:.1f}%", f"{last.p60:.1f}%", f"{d_prev['p60']:+.1f}pt", f"{last.p60-first.p60:+.1f}pt"],
           ["P(S_T < 65,000)", f"{first.p65:.1f}%", f"{prev.p65:.1f}%", f"{last.p65:.1f}%", f"{d_prev['p65']:+.1f}pt", f"{d_all['p65']:+.1f}pt"],
           ["P(S_T ≥ 70,000)", f"{first.p70:.1f}%", f"{prev.p70:.1f}%", f"{last.p70:.1f}%", f"{d_prev['p70']:+.1f}pt", f"{d_all['p70']:+.1f}pt"]], [34, 24, 30, 30, 28, 28]),
      Pg(f"<b>前版から。</b>10/16 の夜間立会で先物は {d_prev['F']:+,.0f} 円下落し、分布は全体として下方に移った。移動幅は 5% 点（{d_prev['q05']:+,.0f} 円）が 95% 点（{d_prev['q95']:+,.0f} 円）より大きく、"
         f"RR25 が {d_prev['rr']:+.1f}pt 拡大したことと整合する。下落の保険（低い行使価格のプット）が相対的に高く買われ、左裾が厚くなった。最頻値と F の差は {prev['mode']-prev.F:+,.0f} 円から {last['mode']-last.F:+,.0f} 円へ広がっており、"
         "「中心は F よりやや上、裾は下」という形が強まっている。"),
      Pg(f"<b>10/9 から。</b>1 週間で F は {d_all['F']/first.F*100:+.1f}% 下落した。分布の幅（標準偏差 ÷ F）は {first.sd/first.F*100:.1f}% から {last.sd/last.F*100:.1f}% へ広がり、歪度は {first["sk"]:+.2f} から {last["sk"]:+.2f} へ悪化した。"
         f"P(S_T ≥ 70,000) は {first.p70:.0f}% から {last.p70:.0f}% へ、P(S_T < 65,000) は {first.p65:.0f}% から {last.p65:.0f}% へ入れ替わった。10/15 の急落日に幅が最大になり、翌日に幅は縮んだが裾の厚みは戻っていない。"),
      Pg("4. 「先物の到達価格」をリスク中立分布で語るときの注意", H2),
      Pg("リスク中立分布の平均は、無裁定である限り先物価格 F に一致する。したがって「分布から導いた到達価格の期待値」は常に F であり、新しい情報を含まない。意味のある量は次の三つである。"),
      Pg("・<b>最頻値</b>（最も密度の高い水準）。本系列では一貫して F より 1,100〜1,900 円高い。左に長い裾を持つ分布では、裾の分だけ平均が引き下げられ、山は右に残るためである。"),
      Pg("・<b>分位点</b>（5%・25%・50%・75%・95% 点）。到達価格を「点」ではなく「確率付きの帯」として示す。"),
      Pg("・<b>水準別の確率</b>。特定の行使価格やサポート水準を SQ 値が下回る確率として読む。"),
      Pg("いずれもリスク中立確率であり、実際の発生確率（実確率）とは分散リスク・プレミアムとスキュー・プレミアムの分だけ異なる。第 II 部では、この差を補正して「投資家の見通し」に近づける方法を含めて整理する。"),
      PageBreak()]

# ---------------- Part II
S += [Pg("第 II 部　先物の先行き・投資家の見通しを抽出する視点", H1), Pg("5. 枠組み", H2),
      Pg("「投資家の見通し」は一つの量ではない。(a) 将来の価格分布についての主観確率、(b) 実際に取られているポジション、(c) 口頭・調査で表明される意見、の三つは互いに異なり、それぞれ別の情報源から取り出す。"),
      TBL([["系統", "情報源", "取り出せるもの", "歪みの原因"],
           ["A. 価格", "オプション価格（スマイル、期間構造、出来高・建玉）、先物価格", "リスク中立分布、分散リスク・プレミアム、スキュー、テール確率", "リスク・プレミアム（保険料）が混入する。実確率ではない"],
           ["B. ポジション・フロー", "投資部門別売買、取引参加者別建玉、裁定残高、CFTC 建玉報告、信用残、ETF 口数", "誰がどちらを持っているか（主体別のネット・ポジション）", "ヘッジ目的の売買と方向性の売買が区別できない"],
           ["C. サーベイ", "QUICK 月次調査、各種 DI、グローバル FMS", "表明された予想（水準・方向）", "予想は外挿的で、将来リターンと逆相関になりやすい"]], [26, 52, 52, 48]),
      Pg("価格から得る分布は「保険料込み」であり、ポジションは「意図を含まない数量」であり、サーベイは「行動を伴わない意見」である。三つを突き合わせて初めて、見通しとして解釈できる。"),
      Pg("6. 三つの情報源の整理", H2), Pg("6.1 系統 A：オプション価格から", H2),
      TBL([["視点", "何を測るか", "根拠・文献", "本ワークスペースでの実装可否"],
           ["A1. リスク中立分布の形状の時系列", "歪度・尖度・RR25・テール確率（本書 第 I 部）", "Breeden–Litzenberger (1978)。日経225 では Shiratsuka (2001, BOJ) が 1989〜96 年で検証：情報はあるが過去分布を上回る予測力は安定せず、サンプル期間に強く依存", "実装済み。日次で蓄積し、SQ ごとに校正検証（A6）を行うのが次の段階"],
           ["A2. 実確率への変換", "リスク中立密度 f_Q を効用関数で傾けて実確率密度 f_P を得る。べき効用なら f_P(S) ∝ f_Q(S)·S^γ（γ は相対的リスク回避度）", "Jackwerth (2000)、Aït-Sahalia and Lo (2000)、Bliss and Panigirtzoglou (2004)。γ は過去の実現値で校正する。Ross (2015) の回復定理は理論上プレミアムを分離できるが、実証では機能しない（Borovička, Hansen and Scheinkman 2016; Jackwerth and Menner 2020）", "可能。f_Q の日次系列と SQ 実現値の蓄積が必要（最低 1〜2 年）。当面は γ を文献値（2〜4）で置いた感度分析"],
           ["A3. 分散リスク・プレミアム VRP", "VRP = IV² − E[RV²]。IV が実現ボラティリティの期待値をどれだけ上回るか", "Bollerslev, Tauchen and Zhou (2009)：VRP が 1〜2 四半期先のリターンを予測。Bollerslev, Marrone, Xu and Zhou (2014)：日本を含む国際データで検証", "可能。ATM IV（QRI）と日経平均の実現ボラティリティから日次で計算できる"],
           ["A4. スキュー・リスクリバーサル", "RR25、25 デルタの BF、リスク中立歪度", "Bakshi, Kapadia and Madan (2003)：歪度は投資家のリスク回避と分布の非対称から決まる。Conrad, Dittmar and Ghysels (2013)：事前歪度と将来リターン。Xing, Zhang and Zhao (2010)：個別株のスマイルの傾きが将来リターンを予測", "実装済み（RR25 の系列）。予測力の検証は未実施"],
           ["A5. テール（ジャンプ）確率", "深い OTM プットの価格から急落確率とその保険料を分離", "Bollerslev and Todorov (2011)：テール部分のプレミアムが「恐怖」を表し、リターン予測力を持つ", "部分的に可能。両翼の気配幅が広く、推定誤差が大きい（本書 第 1 章）"],
           ["A6. 校正検証", "SQ 実現値が分布のどの分位にあったかを記録し（確率積分変換）、一様分布から外れるかを検定", "Berkowitz (2001)。Shiratsuka (2001) も同趣旨の検証を実施", "可能。SQ ごとに 1 点なので年 12 点。月次限月に加え週次（ミニ）を使えば年 100 点超"],
           ["A7. 期間構造", "11月・12月・3月限の ATM IV と RR25 の差。イベント（決算、日銀会合、米国指標）の織り込み時期", "分散の期間構造は将来の分散の期待とプレミアムの和。限月間差分でイベント分散を抽出", "可能。3 限月の QRI データがある"],
           ["A8. オプションのフロー", "新規買い・新規売りの別に集計した P/C 比", "Pan and Poteshman (2006)：新規建ての買い注文の P/C 比が将来リターンを予測（情報を持つ投資家の取引）。JPX には建て/落ちの区分がないため、Δ建玉 × 超過ΔIV の判定（IV モニター）が代替", "代替指標は実装済み。検証は未実施"],
           ["A9. ディーラーのガンマ", "GEX の符号と大きさ", "Barbon and Buraschi (2020)：ディーラーの負のガンマが値動きを増幅。Ni, Pearson and Poteshman (2005)：満期日の価格の引き寄せ", "実装済み。見通しではなく経路（値動きの増幅・減衰）の情報"]], [30, 40, 70, 38]),
      PageBreak(),
      Pg("6.2 系統 B：ポジション・フローから", H2),
      TBL([["視点", "何を測るか", "根拠・文献・データ", "本ワークスペースでの実装可否"],
           ["B1. 投資部門別取引状況（JPX 週次）", "海外投資家・個人・信託銀行等の先物の売買差引（枚・金額）", "Kamesaka, Nofsinger and Kawakita (2003)：日本では海外投資家の買い越しが将来リターンと正、個人は逆。週次 CSV（商品コード 301 = 日経225先物ラージ、313 = ミニ）", "データあり。週次 VWAP と累積ネット建玉の計算は提案済み・未実装"],
           ["B2. 取引参加者別建玉・手口（JPX 日次）", "証券会社別の先物ネット建玉と日々の売買", "本ワークスペースの主機能。参加者をヘッジファンド系・国内機関系などに分類すれば主体別の代理指標になる", "実装済み（分類は未実装）"],
           ["B3. CFTC 建玉報告（CME 日経225先物）", "レバレッジド・ファンド（ヘッジファンド・CTA）とアセット・マネージャーのネット建玉。円建て（コード 240743）とドル建て（240741）", "毎週金曜公表、火曜時点。建玉は約 3.5 万枚と OSE より小さいが、海外ファンドの方向性ポジションの直接観測値。COT の予測力は商品によりまちまち", "未実装。CFTC の公開 CSV を週次で取得すれば可能"],
           ["B4. 裁定取引残高（JPX 週次・日次）", "裁定買い残（先物売り・現物買い）と裁定売り残", "裁定業者が吸収した先物の買い越し・売り越しを映す日本固有の指標。JPX「プログラム売買・裁定取引の状況」に週次（前週末比）と日次（前日比）の PDF", "未実装。PDF の表を読む必要がある（日報と同じ方式で可能）"],
           ["B5. 個人の指標：信用残・評価損益率・レバレッジ ETF の口数", "個人投資家のポジションと心理", "Kaniel, Saar and Titman (2008)：個人の売買は逆張りで、短期のリターンと正の関係。日本では 1570/1357 の口数増減が個人の方向を示す", "未実装。JPX・運用会社の公表値から取得可能"],
           ["B6. CTA ポジション推定", "トレンド追随モデルから機械的な売買を推定", "本ワークスペースの CTA分析プロジェクト。Baltas and Kosowski (2013) 型のモデル", "別プロジェクトに実装あり。先物の需給見通しとして接続可能"],
           ["B7. 日経VI 先物・期間構造", "VI 先物のコンタンゴ／バックワーデーション", "ボラティリティの期待と分散プレミアムの期間構造。リスク選好の代理", "未実装。JPX の VI 先物データで可能"]], [32, 42, 66, 38]),
      Pg("6.3 系統 C：サーベイから", H2),
      TBL([["視点", "何を測るか", "根拠・文献・データ", "実装可否"],
           ["C1. QUICK 月次調査（株式）", "市場関係者の月末・3 か月後の日経平均予想平均と見通し DI。1994 年から蓄積、毎月第 1 木曜に調査", "Greenwood and Shleifer (2014)：サーベイの期待は過去リターンの外挿であり、将来リターンと負の相関。コンセンサスとの乖離が逆張り指標になる", "公表値を手入力または PDF 取得。月次"],
           ["C2. グローバル機関投資家調査", "日本株への配分の過不足（オーバーウエート比率）", "BofA グローバル・ファンドマネージャー調査など。極端な値が転換点と一致する例が多い", "月次。要約値の収集のみ"],
           ["C3. 証券会社・運用会社のレポート", "目標水準の分布（予想の散らばり）", "予想の散らばりは不確実性の指標（Diether, Malloy and Scherbina 2002 は個別株で不一致と低リターンの関係を示す）", "収集コストが高い"]], [32, 50, 62, 34]),
      PageBreak(),
      Pg("7. 統合と実装の優先順位", H2),
      Pg("三系統の指標は、頻度も性質も異なる。統合のための設計を次のとおり提案する。"),
      Pg("<b>(1) 指標の分類。</b>各指標を「水準」（分布の中心や予想水準）、「方向」（ネット・ポジションの符号）、「補正」（プレミアムの大きさ）の三種に分け、同じ種類の指標だけを比較する。価格から得た分布は「補正」を施した後に初めて「水準」として扱う。"),
      Pg("<b>(2) 検証の設計。</b>予測対象を「1〜4 週先の先物リターン」と「SQ 値の分布」の二つに固定し、予測回帰（指標 → リターン）と校正検定（Berkowitz 2001）で評価する。指標ごとに 1 年以上の蓄積が要るため、蓄積の開始が最優先である。"),
      Pg("<b>(3) 文献から期待できる効果。</b>最も再現性が報告されているのは VRP のリターン予測（四半期）、海外投資家フローとの正の関係（日本）、サーベイ期待の逆相関である。リスク中立分布の形状そのものの予測力は弱いか不安定であり（Shiratsuka 2001）、分布は「予測」より「市場が何に保険を掛けているか」の記述に向く。"),
      TBL([["優先", "項目", "必要な作業", "頻度", "データ"],
           ["1", "リスク中立分布の指標の日次蓄積と校正ログ", "rnd_series.py を日次収集に組み込み、SQ ごとに実現値の分位を記録", "日次", "既存（QRI）"],
           ["2", "分散リスク・プレミアム", "ATM IV² − 実現分散（20 日）の日次系列。四半期先リターンとの関係を検証", "日次", "既存（QRI・日経平均）"],
           ["3", "投資部門別の先物ネット売買と裁定残高", "週次 CSV と裁定残高 PDF の取得、海外投資家ネットの累積と週次 VWAP", "週次", "JPX 公表"],
           ["4", "CFTC 建玉報告（240743 / 240741）", "CFTC の週次 CSV からレバレッジド・ファンドとアセット・マネージャーのネットを抽出", "週次", "CFTC 公表"],
           ["5", "QUICK 月次調査の予想平均・DI", "公表値の収集と、コンセンサス乖離の系列化", "月次", "QUICK・日経"],
           ["6", "実確率への変換（効用関数による傾け）", "γ の校正。1〜2 の蓄積後", "日次", "1 の蓄積後"]], [12, 54, 66, 16, 30]),
      Pg("優先 1〜2 は既存データで即日着手でき、3〜4 は公表データの取り込みのみで実装できる。6 は蓄積が前提になる。"),
      sp(4), Pg("付録 A. 本書の計算条件", H1),
      Pg("各日の推定は前版（nk225op_2611_sq_rnd_report.pdf 第 3〜4 章）と同一の手順・パラメータである。スナップショットは各営業日 15:15 以前の最終（両気配 150 銘柄以上）、最終点は 10/16 夜間立会の最終（10/17 06:02）。"
         "RR25 は QRI のデルタから 25 デルタのプット・コール IV を補間した差。分位点は x = ln(K/F) ∈ [−0.6, 0.6] の 6,001 点で数値積分した累積分布から補間。再現: scripts/rnd_series.py と scripts/build_outlook_report.py。", BODY),
      Pg("付録 B. 参考文献・資料", H1),
      Pg("Aït-Sahalia, Y. and A. Lo (2000), Nonparametric Risk Management and Implied Risk Aversion, Journal of Econometrics 94.", SMALL),
      Pg("Bakshi, G., N. Kapadia and D. Madan (2003), Stock Return Characteristics, Skew Laws, and the Differential Pricing of Individual Equity Options, Review of Financial Studies 16.", SMALL),
      Pg("Baltas, N. and R. Kosowski (2013), Momentum Strategies in Futures Markets and Trend-Following Funds, working paper.", SMALL),
      Pg("Barbon, A. and A. Buraschi (2020), Gamma Fragility, working paper (SSRN).", SMALL),
      Pg("Berkowitz, J. (2001), Testing Density Forecasts, with Applications to Risk Management, Journal of Business and Economic Statistics 19.", SMALL),
      Pg("Bliss, R. and N. Panigirtzoglou (2004), Option-Implied Risk Aversion Estimates, Journal of Finance 59.", SMALL),
      Pg("Bollerslev, T., G. Tauchen and H. Zhou (2009), Expected Stock Returns and Variance Risk Premia, Review of Financial Studies 22.", SMALL),
      Pg("Bollerslev, T., J. Marrone, L. Xu and H. Zhou (2014), Stock Return Predictability and Variance Risk Premia: Statistical Inference and International Evidence, Journal of Financial and Quantitative Analysis 49.", SMALL),
      Pg("Bollerslev, T. and V. Todorov (2011), Tails, Fears, and Risk Premia, Journal of Finance 66.", SMALL),
      Pg("Borovička, J., L. P. Hansen and J. Scheinkman (2016), Misspecified Recovery, Journal of Finance 71.", SMALL),
      Pg("Breeden, D. and R. Litzenberger (1978), Prices of State-Contingent Claims Implicit in Option Prices, Journal of Business 51.", SMALL),
      Pg("Conrad, J., R. Dittmar and E. Ghysels (2013), Ex Ante Skewness and Expected Stock Returns, Journal of Finance 68.", SMALL),
      Pg("Diether, K., C. Malloy and A. Scherbina (2002), Differences of Opinion and the Cross Section of Stock Returns, Journal of Finance 57.", SMALL),
      Pg("Greenwood, R. and A. Shleifer (2014), Expectations of Returns and Expected Returns, Review of Financial Studies 27.", SMALL),
      Pg("Jackwerth, J. (2000), Recovering Risk Aversion from Option Prices and Realized Returns, Review of Financial Studies 13.", SMALL),
      Pg("Jackwerth, J. and M. Menner (2020), Does the Ross Recovery Theorem Work Empirically?, Journal of Financial Economics 137.", SMALL),
      Pg("Kamesaka, A., J. Nofsinger and H. Kawakita (2003), Investment Patterns and Performance of Investor Groups in Japan, Pacific-Basin Finance Journal 11.", SMALL),
      Pg("Kaniel, R., G. Saar and S. Titman (2008), Individual Investor Trading and Stock Returns, Journal of Finance 63.", SMALL),
      Pg("Nakamura, H. and S. Shiratsuka (1999), Extracting Market Expectations from Option Prices: Case Studies in Japanese Option Markets, Monetary and Economic Studies 17 (Bank of Japan).", SMALL),
      Pg("Ni, S., N. Pearson and A. Poteshman (2005), Stock Price Clustering on Option Expiration Dates, Journal of Financial Economics 78.", SMALL),
      Pg("Pan, J. and A. Poteshman (2006), The Information in Option Volume for Future Stock Prices, Review of Financial Studies 19.", SMALL),
      Pg("Ross, S. (2015), The Recovery Theorem, Journal of Finance 70.", SMALL),
      Pg("Shiratsuka, S. (2001), Information Content of Implied Probability Distributions: Empirical Studies on Japanese Stock Price Index Options, Monetary and Economic Studies 19(3) (Bank of Japan).", SMALL),
      Pg("Xing, Y., X. Zhang and R. Zhao (2010), What Does the Individual Option Volatility Smirk Tell Us About Future Equity Returns?, Journal of Financial and Quantitative Analysis 45.", SMALL),
      Pg("データ資料: JPX「プログラム売買・裁定取引の状況」（週次・日次 PDF）、JPX「投資部門別取引状況（先物・オプション）」（週次 CSV）、CFTC Commitments of Traders / Traders in Financial Futures（CME Nikkei Stock Average Yen Denom = 240743）、QUICK 月次調査（株式、1994 年〜、毎月第 1 木曜調査）。", SMALL)]

doc = SimpleDocTemplate(str(PDF_PATH), pagesize=A4, leftMargin=16 * mm, rightMargin=16 * mm, topMargin=16 * mm, bottomMargin=18 * mm,
                        title="11月限 SQ値の織り込みの推移と見通し抽出の視点", author="先物手口分析")
ft = make_footer(f"11月限 SQ値の織り込みの推移と見通し抽出の視点（最終データ {last.ts:%Y-%m-%d %H:%M}、作成 {date.today():%Y-%m-%d}）")
doc.build(S, onFirstPage=ft, onLaterPages=ft)
print("written", PDF_PATH)
