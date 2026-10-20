"""日経225オプション 11月限: SQ値の織り込み（リスク中立分布 × 建玉）の解説付き分析レポート（PDF）を生成する。

入力: scripts/analyze_rnd_sq.py の出力（%TMP%/rnd.json, rnd.npz）と、同スナップショット（再読込）。
出力: <out_dir>/nk225op_2611_sq_rnd_report.pdf（既定 out_dir は引数で指定。例: "H:/マイドライブ/bank/nk225op_sq_rnd_2611"）
使い方: python scripts/build_rnd_report.py "H:/マイドライブ/bank/nk225op_sq_rnd_2611"
"""
from __future__ import annotations

import io
import json
import os
import sys
import warnings
from datetime import date
from pathlib import Path

warnings.filterwarnings("ignore")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from PIL import Image as PILImage
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (CondPageBreak, Image, KeepInFrame, KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer,
                                Table, TableStyle)
from scipy.stats import norm

from data import iv_views
from data.aggregator import _load_daily_oi_for_date
from data.qri_iv_loader import load_snapshot

# ------------------------------------------------------------------ 設定
OUT_DIR = Path(sys.argv[1] if len(sys.argv) > 1 else "H:/マイドライブ/bank/nk225op_sq_rnd_2611")
OUT_DIR.mkdir(parents=True, exist_ok=True)
TMP = Path(os.environ["TMP"])
J = json.load(open(TMP / "rnd.json"))
Z = np.load(TMP / "rnd.npz", allow_pickle=True)
PDF_PATH = OUT_DIR / "nk225op_2611_sq_rnd_report.pdf"
CM, SQ_DATE, OI_DATE = "2611", "2026-11-13", date(2026, 10, 20)
sys.path.insert(0, str(Path(__file__).resolve().parent))
from rnd_lib import day_close_snap, rr_delta as rr_delta_key, source_time, stats as rnd_stats, tilt as rnd_tilt  # noqa: E402
GAMMA_JSON = TMP / "gcalc" / "gamma_calc.json"
FONT_DIR = "C:/Windows/Fonts/"
pdfmetrics.registerFont(TTFont("YG", FONT_DIR + "YuGothM.ttc", subfontIndex=0))
pdfmetrics.registerFont(TTFont("YGB", FONT_DIR + "YuGothB.ttc", subfontIndex=0))
pdfmetrics.registerFontFamily("YG", normal="YG", bold="YGB", italic="YG", boldItalic="YGB")
plt.rcParams["font.family"] = ["Yu Gothic", "Meiryo", "sans-serif"]
B, O, A, Y, MUT, INK, GRID = "#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#52514e", "#0b0b0b", "#e6e5e0"

F, T, p = J["F"], J["T"], np.array(J["p"])
K, fK, fL, fP, xs, ivf = Z["K"], Z["fK"], Z["fL"], Z["fP"], Z["xs"], Z["ivf"]
q = {float(k): v for k, v in J["q"].items()}; qL = {float(k): v for k, v in J["qL"].items()}; qP = {float(k): v for k, v in J["qP"].items()}
cdf = {int(k): v for k, v in J["cdf"].items()}; cdfL = {int(k): v for k, v in J["cdfL"].items()}
st = J["st"]; mp = J["maxpain"]; ev = pd.DataFrame(J["ev"])
s_atm, s25p, s25c = J["s_atm"], J["s25p"], J["s25c"]
snap_ts = pd.Timestamp(J["ts"])


def rr_delta(day, mt):
    """QRI のデルタから 25 デルタのプット・コール IV を直接補間した RR25（SVI を介さない確認用）。"""
    ks_ = [k for k in iv_views.snapshot_keys(day) if iv_views.key_time_label(k) <= mt]
    for k in reversed(ks_):
        g_ = iv_views._month_reduced(k, CM)
        if g_ is not None and int(g_.q.sum()) >= 150:
            break
    m_ = iv_views.with_eff_iv(load_snapshot(k)); g_ = m_[m_.contract_month == CM].copy()
    g_["delta"] = pd.to_numeric(g_.delta, errors="coerce"); g_ = g_[g_.ask_iv.notna() & g_.bid_iv.notna() & g_.delta.notna()]
    c_ = g_[g_.option_type == "CALL"].sort_values("delta"); p_ = g_[g_.option_type == "PUT"].sort_values("delta")
    return (np.interp(-0.25, p_.delta, p_.eff_iv) - np.interp(0.25, c_.delta, c_.eff_iv)) * 100


_days = ["20261009", "20261013", "20261014", "20261015", "20261016", "20261019", "20261020"][-len(ev):]
ev["RR25(Δ基準)"] = [round(rr_delta_key(day_close_snap(d, CM), CM), 2) for d in _days]


# ------------------------------------------------------------------ 補助計算（本文の数値例に使う）
def svi(pp, x):
    a, b, rho, m, s = pp; u = x - m; R = np.sqrt(u * u + s * s)
    return a + b * (rho * u + R), b * (rho + u / R), b * s * s / R ** 3


def sig_at(Kv):
    return float(np.sqrt(svi(p, np.array([np.log(Kv / F)]))[0][0] / T))


def black_call(Kv, s):
    d1 = (np.log(F / Kv) + 0.5 * s * s * T) / (s * np.sqrt(T)); d2 = d1 - s * np.sqrt(T)
    return F * norm.cdf(d1) - Kv * norm.cdf(d2), d1


def gfun(x):
    w, w1, w2 = svi(p, np.array([x])); w, w1, w2 = w[0], w1[0], w2[0]
    return (1 - x * w1 / (2 * w)) ** 2 - (w1 ** 2 / 4) * (1 / w + 0.25) + w2 / 2


# スナップショット再読込: パリティ例・再プライシング検証
ks = [k for k in iv_views.snapshot_keys(snap_ts.strftime("%Y%m%d")) if iv_views.key_time_label(k)[:8] == snap_ts.strftime("%H:%M:%S")]
src_time = (source_time(ks[0]) or snap_ts.strftime("%H:%M:%S"))[:5]
m = iv_views.with_eff_iv(load_snapshot(ks[0])); g = m[m.contract_month == CM].copy()
# べき効用による実確率の近似（第 8 章）と市場の γ（gamma_calc.py の出力があれば）
_dK = np.gradient(K); _mQ = fK * _dK; _mQ = _mQ / _mQ.sum()
_gstore = Path(__file__).resolve().parent.parent / "cache" / "derived" / "gamma_daily.jsonl"  # 日次蓄積（gamma_daily.py）から当日の γ̂ を取る
_grow = next((json.loads(l) for l in _gstore.read_text(encoding="utf-8").splitlines()[::-1] if l.strip() and json.loads(l)["day"] == snap_ts.strftime("%Y%m%d")), None) if _gstore.exists() else None
gamma_hat = (_grow.get("gamma_q", _grow.get("gamma_I_quad")) if _grow else None); beta2 = (_grow.get("beta2", _grow.get("quad")) if _grow else None); gamma_II = _grow["gamma_II"] if _grow else None  # 主指標 = 2 次式の傾き（2026-10-21）
TILT_ROWS = []
for g_ in [0.0, 2.0, 4.0, 6.0] + ([round(gamma_hat, 2)] if gamma_hat is not None else []):
    stt = rnd_stats(K, rnd_tilt(K, fK, g_))
    TILT_ROWS.append((g_, stt["mean"], stt["q"][0.5], stt["mode"], stt["q"][0.05], stt["q"][0.95], stt["cdf"](60000), stt["cdf"](65000), 1 - stt["cdf"](70000)))


def fig_tilt():
    sel = (K >= 50000) & (K <= 80000); fig, a = plt.subplots(figsize=(7.2, 3.3))
    a.plot(K[sel], fK[sel] * 1e5, color=B, lw=2, label="リスク中立 f_Q（γ = 0）")
    a.plot(K[sel], rnd_tilt(K, fK, 4.0)[sel] * 1e5, color=O, lw=1.6, ls="--", label="f_P（γ = 4、文献の代表値）")
    if gamma_hat is not None:
        a.plot(K[sel], rnd_tilt(K, fK, gamma_hat)[sel] * 1e5, color=A, lw=1.8, label=f"f_P（γ̂ = {gamma_hat:.2f}、市場から推定）")
    a.axvline(F, color=INK, lw=0.8)
    style_ax(a, "図9  べき効用で傾けた実確率密度の近似", "SQ値（円）", "確率密度（%/1,000円）"); a.legend(fontsize=8, frameon=False)
    return fig_to_img(fig)
for c in ["ask_price", "bid_price", "ask_iv", "bid_iv", "delta"]:
    g[c] = pd.to_numeric(g[c], errors="coerce")
g["mid"] = ((g.ask_price + g.bid_price) / 2).where(g.ask_price.notna() & g.bid_price.notna())
pv = g.pivot_table(index="strike", columns="option_type", values="mid").dropna(); par = pv.CALL - pv.PUT
near = par.abs().nsmallest(5).sort_index(); par_tbl = [(int(k_), pv.loc[k_, "CALL"], pv.loc[k_, "PUT"], k_ + par[k_]) for k_ in near.index]
two = g[g.ask_iv.notna() & g.bid_iv.notna() & g.mid.notna()].copy(); two["x"] = np.log(two.strike / F)
otm = two[((two.option_type == "PUT") & (two.x < 0)) | ((two.option_type == "CALL") & (two.x >= 0))].copy()


def black_px(row):
    c_, d1 = black_call(row.strike, row.eff_iv)
    return c_ if row.option_type == "CALL" else c_ - (F - row.strike)


otm["model"] = otm.apply(black_px, axis=1); otm["spread"] = otm.ask_price - otm.bid_price
otm["err"] = otm.model - otm.mid
reprice_rmse = float(np.sqrt((otm.err ** 2).mean())); reprice_within = float((otm.err.abs() <= otm.spread / 2).mean())
w0 = svi(p, np.array([0.0]))[0][0]
K_ex = 62000; x_ex = np.log(K_ex / F); w_ex = svi(p, np.array([x_ex]))[0][0]; s_ex = np.sqrt(w_ex / T)
obs_ex = otm[otm.strike == K_ex].eff_iv.mean()
# Breeden–Litzenberger 数値例
Kc, h = 66000, 500
C_m, C_0, C_p = black_call(Kc - h, sig_at(Kc - h))[0], black_call(Kc, sig_at(Kc))[0], black_call(Kc + h, sig_at(Kc + h))[0]
f_fd = (C_p - 2 * C_0 + C_m) / h ** 2; f_an = float(np.interp(Kc, K, fK))
dig = 1 + (C_p - C_m) / (2 * h); cdf_66 = float(np.interp(Kc, K, np.cumsum(fK * np.gradient(K)) / (fK * np.gradient(K)).sum()))
g0 = gfun(0.0)
# 建玉
oi = pd.DataFrame([vars(r) for r in _load_daily_oi_for_date(OI_DATE, CM)]); oi = oi[oi.current_oi > 0]
oc, op = oi[oi.option_type == "CALL"], oi[oi.option_type == "PUT"]
S_grid, pain = Z["S"], Z["pain"]


def pain_at(Sv):
    return float((oc.current_oi * np.maximum(Sv - oc.strike_price, 0)).sum() + (op.current_oi * np.maximum(op.strike_price - Sv, 0)).sum())


gk, gv = Z["gx_k"], Z["gx_v"]
s60 = sig_at(60000); _, d1_60 = black_call(60000, s60); gam60 = norm.pdf(d1_60) / (F * s60 * np.sqrt(T)); oi60 = int(op[op.strike_price == 60000].current_oi.iloc[0])
gex60 = -gam60 * oi60 * 1000 * F ** 2 * 0.01
deep_put = int(op[op.strike_price <= 50000].current_oi.sum()); pc_core = op[op.strike_price.between(55000, 78000)].current_oi.sum() / oc[oc.strike_price.between(55000, 78000)].current_oi.sum()
mass = fK * np.gradient(K); mass = mass / mass.sum(); cdfv = np.cumsum(mass)
cdf_at = lambda k_: float(np.interp(k_, K, cdfv))
cdf = {k_: cdf_at(k_) for k_ in range(40000, 90001, 500)}  # 本文の参照範囲を広げる（json は 60,000〜78,000 のみ）
P = lambda lo, hi: float(mass[(K >= lo) & (K < hi)].sum())
bins = [(k_, P(k_, k_ + 1000)) for k_ in range(56000, 76000, 1000)]
p_tail_lo, p_tail_hi = float(mass[K < 56000].sum()), float(mass[K >= 76000].sum())


# ------------------------------------------------------------------ 図
def fig_to_img(fig, width_mm=172):
    buf = io.BytesIO(); fig.savefig(buf, format="png", dpi=200, bbox_inches="tight", facecolor="white"); plt.close(fig); buf.seek(0)
    w, h_ = PILImage.open(buf).size; buf.seek(0)
    return Image(buf, width=width_mm * mm, height=width_mm * mm * h_ / w)


def style_ax(a, t, xl, yl):
    a.grid(color=GRID, lw=0.8); a.tick_params(colors=MUT, labelsize=8.5)
    for sp in a.spines.values():
        sp.set_color("#d0cfca")
    a.set_title(t, fontsize=10.5, color=INK, loc="left"); a.set_xlabel(xl, color=MUT, fontsize=9); a.set_ylabel(yl, color=MUT, fontsize=9)


lo, hi = 50000, 80000; sel = (K >= lo) & (K <= hi)


def fig_butterfly():
    fig, a = plt.subplots(figsize=(7.2, 2.6)); S = np.linspace(63000, 69000, 400); Kc_, h_ = 66000, 1000
    pay = np.maximum(S - (Kc_ - h_), 0) - 2 * np.maximum(S - Kc_, 0) + np.maximum(S - (Kc_ + h_), 0)
    a.plot(S, pay, color=B, lw=2); a.fill_between(S, pay, 0, color=B, alpha=0.12)
    a.axvline(Kc_, color=MUT, lw=0.8, ls=":"); a.text(Kc_, h_ * 1.05, "K", ha="center", fontsize=9, color=INK)
    a.text(Kc_ - h_, -120, "K−h", ha="center", fontsize=9, color=INK); a.text(Kc_ + h_, -120, "K+h", ha="center", fontsize=9, color=INK)
    style_ax(a, "図1  バタフライ（C(K−h) − 2C(K) + C(K+h)）の満期ペイオフ", "満期時の S_T（円）", "ペイオフ（円）"); a.set_ylim(-250, 1200)
    return fig_to_img(fig, 150)


def fig_smile():
    fig, a = plt.subplots(figsize=(7.2, 3.4)); pk, piv = Z["pts_K"], Z["pts_iv"]; sp_ = (pk >= lo) & (pk <= hi)
    a.scatter(pk[sp_], piv[sp_] * 100, s=12, color=O, label="QRI 両気配 IV（OTM 側、213点中の表示範囲）", zorder=3)
    a.plot(K[sel], ivf[sel] * 100, color=B, lw=2, label="SVI フィット"); a.axvline(F, color=INK, lw=1); a.text(F, a.get_ylim()[1] if False else piv[sp_].max() * 100, f" F={F:,.0f}", fontsize=8, color=INK, va="top")
    style_ax(a, f"図2  IV スマイル（2026年11月限、{snap_ts:%m/%d %H:%M} 時点）", "行使価格 K（円）", "IV（%）"); a.legend(fontsize=8, frameon=False)
    return fig_to_img(fig)


def fig_density():
    fig, a = plt.subplots(figsize=(7.2, 3.6)); ymax = max(fK[sel].max(), fL[sel].max()) * 1e5; a.set_ylim(0, ymax * 1.12)
    a.plot(K[sel], fK[sel] * 1e5, color=B, lw=2, label="リスク中立密度（SVI から導出）")
    a.plot(K[sel], fL[sel] * 1e5, color=MUT, lw=1.5, ls="--", label=f"対数正規（ATM IV {s_atm*100:.1f}% で一定とした場合）")
    bb = dict(fc="white", ec="none", pad=0.5)
    for qq, lab in [(0.05, "5%"), (0.5, "50%"), (0.95, "95%")]:
        a.axvline(q[qq], color=O, lw=1, ls=":"); a.text(q[qq], ymax * 0.46, f"{lab}\n{q[qq]:,.0f}", fontsize=7.5, color=O, ha="center", va="top", bbox=bb)
    a.axvline(F, color=INK, lw=1); a.text(F, ymax * 0.30, f"F\n{F:,.0f}", fontsize=7.5, color=INK, ha="center", va="top", bbox=bb)
    a.axvline(mp, color=A, lw=1, ls="--"); a.text(mp, ymax * 0.15, f"マックスペイン\n{mp:,}", fontsize=7.5, color=A, ha="center", va="top", bbox=bb)
    style_ax(a, "図3  SQ値 S_T の確率密度", "SQ値（円）", "確率密度（%/1,000円）"); a.legend(fontsize=8, frameon=False, loc="upper left")
    return fig_to_img(fig)


def fig_cdf():
    fig, a = plt.subplots(figsize=(7.2, 3.2)); kk = np.array(sorted(cdfL))
    a.plot(kk, [cdf[k_] * 100 for k_ in kk], color=B, lw=2, label="リスク中立"); a.plot(kk, [cdfL[k_] * 100 for k_ in kk], color=MUT, lw=1.5, ls="--", label="対数正規")
    for lv in (25, 50, 75):
        a.axhline(lv, color=GRID, lw=0.8)
    a.axvline(F, color=INK, lw=1); a.set_xlim(lo, hi)
    style_ax(a, "図4  累積分布 P(S_T < K)", "K（円）", "累積確率（%）"); a.legend(fontsize=8, frameon=False, loc="upper left")
    return fig_to_img(fig)


def fig_bins():
    fig, a = plt.subplots(figsize=(7.2, 3.0)); xb = [b_[0] + 500 for b_ in bins]; yb = [b_[1] * 100 for b_ in bins]
    a.bar(xb, yb, width=800, color=B)
    for x_, y_ in zip(xb, yb):
        a.text(x_, y_ + 0.15, f"{y_:.1f}", ha="center", fontsize=7.5, color=INK)
    a.axvline(F, color=INK, lw=1); a.set_xticks([b_[0] for b_ in bins][::2]); a.set_xticklabels([f"{b_[0]//1000}k" for b_ in bins][::2])
    style_ax(a, f"図5  1,000円刻みの区間確率（56,000円未満 {p_tail_lo*100:.1f}%、76,000円以上 {p_tail_hi*100:.1f}% は図外）", "区間の下端（円）", "確率（%）")
    return fig_to_img(fig)


def fig_oi():
    fig, a = plt.subplots(figsize=(7.2, 3.4)); os_, ot, ocv = Z["oi_strike"], Z["oi_type"], Z["oi_cur"]; m_ = (os_ >= 55000) & (os_ <= 78000)
    a.bar(os_[m_ & (ot == "CALL")], ocv[m_ & (ot == "CALL")], width=400, color=B, label="コール建玉")
    a.bar(os_[m_ & (ot == "PUT")], -ocv[m_ & (ot == "PUT")], width=400, color=O, label="プット建玉（下向きに表示）")
    a.axvline(F, color=INK, lw=1); a.axvline(mp, color=A, lw=1, ls="--")
    for qq in (0.05, 0.5, 0.95):
        a.axvline(q[qq], color=MUT, lw=0.8, ls=":")
    a.axhline(0, color=MUT, lw=0.8)
    style_ax(a, "図6  行使価格別の建玉（55,000〜78,000円）。点線=リスク中立 5/50/95% 点、破線=マックスペイン", "行使価格（円）", "建玉（枚）")
    a.legend(fontsize=8, frameon=False, loc="lower right")
    return fig_to_img(fig)


def fig_pain():
    fig, a = plt.subplots(figsize=(7.2, 2.8)); m_ = (S_grid >= 58000) & (S_grid <= 76000)
    a.plot(S_grid[m_], pain[m_] * 1000 / 1e8, color=B, lw=2); a.axvline(mp, color=A, lw=1, ls="--"); a.text(mp, (pain[m_] * 1000 / 1e8).max() * 0.95, f" 最小 {mp:,}円", color=A, fontsize=8.5)
    style_ax(a, "図7  満期時の建玉全体の受取額 Pain(S)（取引単位 1,000 倍、億円）", "満期時の S_T（円）", "億円")
    return fig_to_img(fig)


def fig_evol():
    fig, a = plt.subplots(figsize=(7.2, 3.2)); xi = np.arange(len(ev))
    a.fill_between(xi, ev.q05, ev.q95, color=B, alpha=0.15, label="5〜95% 点"); a.fill_between(xi, ev.q25, ev.q75, color=B, alpha=0.35, label="25〜75% 点")
    a.plot(xi, ev.q50, color=B, lw=2, marker="o", ms=4, label="中央値"); a.plot(xi, ev.F, color=INK, lw=1.2, marker="s", ms=4, label="F（先物推定値）")
    a.axhline(70000, color=MUT, lw=0.8, ls=":"); a.axhline(65000, color=MUT, lw=0.8, ls=":"); a.set_xticks(xi); a.set_xticklabels(ev.day, fontsize=8)
    style_ax(a, "図8  リスク中立分位点の推移（10月SQ 後の各営業日の引け近傍）", "スナップショット", "SQ値（円）"); a.legend(fontsize=8, frameon=False, loc="lower left")
    return fig_to_img(fig)


# ------------------------------------------------------------------ 文書部品
def EQ(latex, size=14, max_w=172):
    fig = plt.figure(figsize=(10, 0.6)); fig.text(0.01, 0.5, f"${latex}$", fontsize=size, ha="left", va="center"); plt.axis("off")
    buf = io.BytesIO(); fig.savefig(buf, format="png", dpi=300, transparent=True, bbox_inches="tight", pad_inches=0.04); plt.close(fig); buf.seek(0)
    w, h_ = PILImage.open(buf).size; buf.seek(0); wpt = min(max_w * mm, w * 72 / 300 * 0.9)
    img = Image(buf, width=wpt, height=wpt * h_ / w); img.hAlign = "LEFT"
    return img


BODY = ParagraphStyle("b", fontName="YG", fontSize=10.5, leading=17, spaceAfter=5, wordWrap="CJK", alignment=TA_LEFT)
SMALL = ParagraphStyle("s", parent=BODY, fontSize=9, leading=14, textColor=colors.HexColor(MUT))
CAP = ParagraphStyle("c", parent=BODY, fontSize=8.5, leading=12.5, textColor=colors.HexColor(MUT), spaceAfter=8)
H1 = ParagraphStyle("h1", fontName="YGB", fontSize=15, leading=21, spaceBefore=12, spaceAfter=7, textColor=colors.HexColor(INK), wordWrap="CJK")
H2 = ParagraphStyle("h2", fontName="YGB", fontSize=12.5, leading=18, spaceBefore=9, spaceAfter=5, textColor=colors.HexColor(INK), wordWrap="CJK")
TITLE = ParagraphStyle("t", fontName="YGB", fontSize=19, leading=27, spaceAfter=6, textColor=colors.HexColor(INK), wordWrap="CJK")
SUB = ParagraphStyle("st", fontName="YG", fontSize=12, leading=18, spaceAfter=14, textColor=colors.HexColor(MUT), wordWrap="CJK")
CELL = ParagraphStyle("cell", fontName="YG", fontSize=9, leading=12.5, wordWrap="CJK")
CELLB = ParagraphStyle("cellb", parent=CELL, fontName="YGB")


def Pg(t, style=BODY):
    return Paragraph(t, style)


def TBL(rows, widths, head=True, align_right_from=1, font=9):
    data = [[Paragraph(str(c), CELLB if (head and i == 0) else CELL) for c in r] for i, r in enumerate(rows)]
    t = Table(data, colWidths=[w * mm for w in widths], repeatRows=1 if head else 0)
    ts = [("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#d0cfca")), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
          ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3), ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4)]
    if head:
        ts.append(("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#f0efe9")))
    for i in range(1 if head else 0, len(rows)):
        if i % 2 == 0:
            ts.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#fafaf7")))
    t.setStyle(TableStyle(ts))
    return t  # 縦に長い表はページをまたいで分割させる（KeepInFrame の shrink は残り余白に合わせて縮小してしまう）


def footer(canvas, doc):
    canvas.saveState(); canvas.setFont("YG", 8); canvas.setFillColor(colors.HexColor(MUT))
    canvas.drawString(16 * mm, 10 * mm, f"日経225オプション 2026年11月限 SQ値の織り込み分析（基準 {snap_ts:%Y-%m-%d} 引け、作成 {date.today():%Y-%m-%d}）")
    canvas.drawRightString(A4[0] - 16 * mm, 10 * mm, f"{doc.page}"); canvas.restoreState()


S = []
sp = lambda h_=4: Spacer(1, h_ * mm)
fmt = lambda v: f"{v:,.0f}"

# ================================================================== 表紙・要約
S += [sp(10), Pg("日経225オプション 2026年11月限<br/>SQ値の織り込みに関する確率分析", TITLE),
      Pg("オプション価格から導いたリスク中立分布と、建玉の位置づけ", SUB),
      TBL([["項目", "内容"],
           ["分析対象", "日経225オプション 2026年11月限（SQ算出日 2026-11-13、最終取引日 2026-11-12）"],
           ["価格データ", f"QRI 配信の気配・IV スナップショット {snap_ts:%Y-%m-%d %H:%M}（配信は 15 分遅延。気配時刻 {src_time} は日中立会 08:45〜15:30 内の最終。夜間立会の気配は使わない）"],
           ["建玉データ", f"JPX 建玉残高表 {OI_DATE:%Y-%m-%d} 引け分（別紙1、行使価格別・プット/コール別）"],
           ["先物価格の推定値", f"F = {fmt(F)} 円（プット・コール・パリティから推定。第3.3節）"],
           ["残存期間", f"{T*365:.1f} 暦日（SQ算出時刻 11/13 9:00 まで）"],
           ["作成", "先物手口分析ワークスペース（scripts/analyze_rnd_sq.py, scripts/build_rnd_report.py）"]], [40, 138]),
      sp(6), Pg("要約", H1),
      Pg(f"オプション価格には、満期日の SQ値 S_T がどの水準にどれだけの確率で収まるかが、無裁定の条件から一意に織り込まれている。"
         f"本書はこの「リスク中立分布」を 11月限の気配から推定し、建玉の位置と突き合わせた。主要な結論は次のとおりである。"),
      Pg(f"・分布の中心は F = {fmt(F)} 円付近にあるが、形は対称ではない。最頻値 {fmt(st['mode'])} 円、中央値 {fmt(q[0.5])} 円、平均 {fmt(st['mean'])} 円の順に並び、"
         f"下落側に長い裾を持つ（歪度 {st['skew']:+.2f}）。市場は「最も起こりやすいのは小幅高、ただし大幅安の確率は正規分布的な想定より厚い」という形で織り込んでいる。"),
      Pg(f"・90% の確率で収まる範囲（5% 点〜95% 点）は {fmt(q[0.05])}〜{fmt(q[0.95])} 円。標準偏差は {fmt(st['sd'])} 円（F の {st['sd']/F*100:.1f}%）。"),
      Pg(f"・水準別の累積確率は P(S_T < 60,000) = {cdf[60000]*100:.1f}%、P(S_T < 65,000) = {cdf[65000]*100:.1f}%、P(S_T ≥ 70,000) = {(1-cdf[70000])*100:.1f}%。"),
      Pg(f"・建玉はプットに偏る（P/C = {J['put_oi']/J['call_oi']:.2f}）が、プット建玉の約 3 割は行使価格 50,000 円以下にあり、SQ値には実質的に関係しない。"
         f"マックスペインは {mp:,} 円で分布の中央値より約 {mp-q[0.5]:,.0f} 円高い。コール建玉の山（69,000〜71,000 円）は分布の右裾（P(S_T ≥ 70,000) = {(1-cdf[70000])*100:.0f}%）にある。"),
      Pg(f"・10月SQ 後の 1 週間で、分布の中央値は {fmt(ev.q50.iloc[0])} 円から {fmt(ev.q50.iloc[-1])} 円へ下がり、P(S_T ≥ 70,000) は {ev.p_gt70k.iloc[0]:.0f}% から {ev.p_gt70k.iloc[-1]:.0f}% へ、"
         f"P(S_T < 65,000) は {ev.p_lt65k.iloc[0]:.0f}% から {ev.p_lt65k.iloc[-1]:.0f}% へ移った。歪度は一貫して悪化している。"),
      Pg("・リスク中立確率は実際の発生確率ではなく、ヘッジ費用を反映した価格である。直近の実現ボラティリティは IV とほぼ等しく分散プレミアムは小さいが、"
         "プットが相対的に高い分（スキュー・プレミアム）は除去できない。下落側の確率は実確率より高めに出ている可能性が高い（第8章）。"),
      sp(4), Pg("構成", H2),
      Pg("1. 問いの設定　2. データ　3. 記号と前提　4. 理論：オプション価格から確率分布を取り出す　5. 結果：SQ値の分布　"
         "6. 建玉との突合　7. 織り込みの推移　8. リスク中立確率と実確率の差　9. 限界　付録（パラメータ・再現方法・参考文献）", SMALL),
      PageBreak()]

# ================================================================== 1 問い
S += [Pg("1. 問いの設定", H1),
      Pg("「SQ値がどう織り込まれているか」という問いは、二つに分けて答える必要がある。"),
      Pg("第一は<b>価格</b>からの問いである。オプションは満期日の SQ値 S_T に応じて支払いが決まる証券であり、その価格は S_T の確率分布に依存する。"
         "逆に、行使価格の異なるオプションの価格を並べると、S_T の確率分布を一意に復元できる（第4章）。こうして得られる分布を<b>リスク中立分布</b>と呼ぶ。"
         "「市場参加者が合意した確率」ではなく、「その確率で評価すると全てのオプション価格と無裁定になる確率」である。この区別は第8章で扱う。"),
      Pg("第二は<b>建玉</b>からの問いである。建玉（未決済の残高）は確率ではないが、満期時に損益が集中する水準を示す。"
         "どの行使価格に残高が積み上がっているかは、満期に向けたヘッジ売買や損益確定の圧力がどこに生じるかの手掛かりになる（第6章）。"),
      Pg("本書は、第一の問いに定量的に答え（確率密度・分位点・水準別の確率）、第二の問いをその分布の上に重ねて読む。"),
      # ================================================================== 2 データ
      Pg("2. データ", H1),
      TBL([["データ", "内容", "時点・頻度", "用途"],
           ["QRI スナップショット", "行使価格×プット/コール別の買い気配・売り気配（価格と IV）、デルタ等。配信は 15 分遅延", f"{snap_ts:%m/%d %H:%M}。収集は 15 分間隔", "スマイルの推定（第4〜5章）、先物価格の推定（第3.3節）"],
           ["JPX 建玉残高表（別紙1）", "日経225オプション直近3限月の行使価格別建玉・前日比・取引高", f"{OI_DATE:%m/%d} 引け（当日 20:00 頃公表）", "建玉の分布、マックスペイン、GEX（第6章）"],
           ["日経平均 日次終値", "Yahoo Finance（^N225）", "2026-07-01〜10-16", "実現ボラティリティ（第8章）"]], [34, 66, 40, 38]),
      Pg(f"QRI の 11月限は {len(g)} 銘柄、うち両気配（買いと売りの両方に気配がある）が {int(two.strike.nunique()*0+len(two))} 銘柄、"
         f"そのうちスマイル推定に用いた OTM 側（第3.4節）が {J['n_pts']} 行使価格である。", BODY),
      Pg("建玉残高表は売り手と買い手の残高を区別しない。建玉 1,000 枚は「1,000 枚の買いと 1,000 枚の売りが未決済」という意味であり、方向は分からない。", BODY),
      # ================================================================== 3 記号と前提
      Pg("3. 記号と前提", H1), Pg("3.1 SQ の仕組み", H2),
      Pg("日経225オプションの最終決済は、SQ算出日（限月の第2金曜日。11月限は 2026-11-13）の日経平均構成銘柄の始値から計算する<b>特別清算指数（SQ値）</b>で行う。"
         "先物も同じ SQ値で決済されるため、満期日における先物の「清算値」とオプションの決済値は同一の数値 S_T である。最終取引日は前日の 11/12 である。"),
      Pg("3.2 記号", H2),
      TBL([["記号", "意味", "本書での値・出所"],
           ["S_T", "満期日の SQ値（確率変数）", "推定対象"],
           ["K", "行使価格", "125〜1,000 円刻み"],
           ["C(K), P(K)", "行使価格 K のコール・プットの価格（気配仲値）", "QRI"],
           ["F", "満期 T で決済される先物価格の推定値", f"{fmt(F)} 円（第3.3節）"],
           ["T", "残存期間（年）。暦日 ÷ 365", f"{T*365:.1f} 日 → T = {T:.5f}"],
           ["x", "対数マネーネス x = ln(K/F)。F で x = 0", "−0.6〜+0.6 の格子で計算"],
           ["σ(K), σ(x)", "行使価格 K のインプライド・ボラティリティ（IV、年率）", "QRI の両気配 IV の仲値（第3.4節）"],
           ["w(x)", "総分散 w = σ(x)²·T", "SVI で当てはめ（第4.4節）"],
           ["f(K)", "S_T のリスク中立確率密度", "第4章で導出"],
           ["Φ(K)", "累積分布 P(S_T < K)", "f の積分"],
           ["q_α", "分位点。P(S_T ≤ q_α) = α", "表5"],
           ["N(·), φ(·)", "標準正規分布の累積分布関数と密度関数", "—"]], [26, 92, 60]),
      Pg("3.3 先物価格 F の推定（プット・コール・パリティ）", H2),
      Pg("同じ行使価格のコール買いとプット売りを組み合わせると、満期の受取は S_T − K となり、先物の買いと同じになる。"
         "満期決済が同一の SQ値で行われるため、金利を無視すれば次の関係が成り立つ（パリティ）。"),
      EQ(r"F = K + C(K) - P(K)"),
      Pg("なぜ先物の気配を直接使わないのか。第一に、本ワークスペースの収集対象は先物の気配を含まない。第二に、オプションの分布の「中心」として整合的なのは、"
         "オプション自身の価格が暗黙に前提している先物価格であり、パリティから逆算する値がそれに当たる。"),
      Pg("気配のずれを平均するため、C と P の仲値の差が最も小さい（= ATM に近い）5 行使価格でそれぞれ F を計算し、その中央値を採用した。"),
      TBL([["K（円）", "C(K) 仲値", "P(K) 仲値", "K + C − P"]] + [[fmt(k_), f"{c_:,.1f}", f"{p_:,.1f}", f"{f_:,.1f}"] for k_, c_, p_, f_ in par_tbl], [40, 40, 40, 48]),
      Pg(f"5 行使価格の中央値 F = {fmt(F)} 円。値の散らばりは {max(x[3] for x in par_tbl)-min(x[3] for x in par_tbl):,.0f} 円で、気配幅の範囲内である。", CAP),
      Pg("3.4 その他の前提と、その理由", H2),
      Pg(f"<b>金利 r = 0 とする。</b>残存 {T*365:.0f} 日、短期金利 0.5% とすると割引係数は exp(−0.005×{T:.3f}) = {np.exp(-0.005*T):.4f} であり、価格への影響は 0.04% と気配幅より小さい。"),
      Pg("<b>残存期間は暦日で数える。</b>IV は暦日ベースの年率で表示される慣行であり、週末にも分散が蓄積する前提で価格が付いている。"
         f"QRI の IV がこの慣行と整合しているかは、QRI の IV と本書の F・T からブラック式（第4.7節）で理論価格を計算し、気配仲値と比べて確かめた。"
         f"OTM 側 {len(otm)} 銘柄で誤差の二乗平均平方根は {reprice_rmse:.1f} 円、{reprice_within*100:.0f}% の銘柄で誤差が気配幅の半分以内に収まり、整合している。"),
      Pg("<b>IV は OTM 側を使う。</b>K < F ではプット、K ≥ F ではコールの IV を用いる。理由は二つある。ITM（イン・ザ・マネー）側は気配幅が広く IV の精度が低いこと、"
         "パリティにより同じ K のコールとプットの IV は理論上一致し情報が重複することである。両気配の無い銘柄は使わない。"),
      PageBreak()]

# ================================================================== 4 理論
S += [Pg("4. 理論：オプション価格から確率分布を取り出す", H1), Pg("4.1 直観：バタフライは「ある水準に収まる確率」の価格", H2),
      Pg("行使価格 K−h、K、K+h のコールを 1 枚買い、2 枚売り、1 枚買う組み合わせ（バタフライ）を考える。満期の受取は図1 のとおり、S_T が K に近いほど大きく、"
         "K から h 以上離れるとゼロになる三角形になる。"),
      fig_butterfly(),
      Pg("この三角形の面積は h² である。h を十分小さくすると、三角形の中で S_T の密度 f はほぼ一定とみなせるので、バタフライの価格（= 受取の期待値）は"),
      EQ(r"C(K-h) - 2C(K) + C(K+h) \;\approx\; h^{2}\, f(K)"),
      Pg("となる。すなわち、<b>バタフライの価格を h² で割ると K における確率密度になる</b>。これが「オプション価格に確率分布が織り込まれている」ことの意味である。"),
      Pg("4.2 導出（Breeden–Litzenberger, 1978）", H2),
      Pg("コール価格は満期の受取 max(S_T − K, 0) の期待値である（金利ゼロ）。密度 f で書くと"),
      EQ(r"C(K) = \int_{K}^{\infty} (S-K)\, f(S)\, dS"),
      Pg("これを K で微分すると、被積分関数の K 微分が −1 であることから"),
      EQ(r"\frac{\partial C}{\partial K} = -\int_{K}^{\infty} f(S)\,dS = -\left[1-\Phi(K)\right], \qquad \frac{\partial^{2} C}{\partial K^{2}} = f(K)"),
      Pg("となる。1 階微分は「S_T が K を上回る確率」に符号を付けたもの、2 階微分が密度である。4.1 のバタフライは 2 階微分の差分近似に他ならない。"),
      Pg(f"<b>数値による確認。</b>第5章で当てはめたスマイルから K = {fmt(Kc)} 円、h = {h} 円のコール価格を計算すると "
         f"C(K−h) = {C_m:,.1f}、C(K) = {C_0:,.1f}、C(K+h) = {C_p:,.1f} 円。バタフライ価格は {C_m - 2*C_0 + C_p:,.2f} 円で、h² = {h*h:,} で割ると "
         f"f(K) ≈ {f_fd*1e5:.2f} %/1,000円。解析的に求めた密度 {f_an*1e5:.2f} %/1,000円 と一致する。"
         f"また 1 階微分から P(S_T < {fmt(Kc)}) = 1 + [C(K+h) − C(K−h)]/(2h) = {dig*100:.1f}% となり、累積分布の値 {cdf_66*100:.1f}% と一致する。"),
      Pg("4.3 なぜ気配を直接 2 階差分してはいけないか", H2),
      Pg("行使価格は 125〜1,000 円刻みで離散的であり、気配の仲値には気配幅の半分程度（ATM 付近で数円）の誤差がある。2 階差分は誤差を h² で割って増幅する。"
         "たとえば仲値に ±5 円の誤差があり h = 250 円なら、密度の誤差は最大 4 × 5 ÷ 250² = 0.00032 /円、すなわち 32%/1,000円 に達し、密度そのもの（最大 9%/1,000円）より大きい。"
         "したがって、まず気配を<b>滑らかな関数</b>で表し、その関数を微分する必要がある。関数の形には、無裁定（密度が負にならない）を保証できるものを選ぶ。"),
      Pg("4.4 スマイルの関数形：SVI", H2),
      Pg("IV を価格ではなく総分散 w(x) = σ(x)²T で表し、対数マネーネス x = ln(K/F) の関数として次の 5 パラメータ形（SVI, Gatheral 2004）を当てはめる。"),
      EQ(r"w(x) = a + b\left[\rho\,(x-m) + \sqrt{(x-m)^{2} + s^{2}}\right], \qquad \sigma(x) = \sqrt{w(x)/T}"),
      Pg("各パラメータの役割は次のとおりである。a は総分散の水準、b は両翼（x → ±∞）の傾きの大きさ、ρ は左右の非対称（負なら低い行使価格側の IV が高い＝スキュー）、"
         "m は最小点の横位置、s は最小点付近の丸み（曲率）である。両翼が x に対して直線的に増えるのは、Lee (2004) が示した「総分散の両翼は |x| に比例する速さ以下でしか増えられない」"
         "というモーメント制約と整合する形であり、株価指数オプションの実務で標準的に使われる。"),
      Pg("<b>当てはめ方。</b>両気配 IV の仲値 σ_i を総分散 w_i = σ_i²T に直し、次の重み付き二乗誤差を最小化する。"),
      EQ(r"\min_{a,b,\rho,m,s}\; \sum_{i} \omega_{i}\,\left(w(x_{i}) - \sigma_{i}^{2}T\right)^{2}, \qquad \omega_{i} = \frac{1}{(\mathrm{spread}_{i} + 0.005)^{2}}"),
      Pg("spread_i は買い気配 IV と売り気配 IV の差である。気配幅が広い銘柄ほど仲値の不確かさが大きいので、重みを小さくする。"
         "0.005（0.5%pt）は、気配幅がごく狭い銘柄が重みを独占しないための下限である。制約 b ≥ 0、|ρ| < 1、s > 0、および w(x) ≥ 0 を課した。"),
      Pg(f"<b>結果。</b>a = {p[0]:.5f}、b = {p[1]:.4f}、ρ = {p[2]:.3f}、m = {p[3]:.4f}、s = {p[4]:.3f}。x = 0 での総分散は w(0) = {w0:.5f} で、"
         f"σ(0) = √(w(0)/T) = √({w0:.5f}/{T:.5f}) = {s_atm*100:.2f}% が ATM IV である。例として K = {K_ex:,} 円（x = {x_ex:+.4f}）では w = {w_ex:.5f}、σ = {s_ex*100:.2f}% となり、"
         f"観測値 {obs_ex*100:.2f}% に近い。IV の二乗平均誤差は {J['rmse']*100:.2f}%pt である（図2）。"),
      Pg("4.5 SVI から密度を解析的に求める", H2),
      Pg("w(x) が滑らかな関数なら、4.2 の 2 階微分を解析的に実行できる。Gatheral and Jacquier (2014) により、x の密度 p(x) は次の形になる。"),
      EQ(r"p(x) = \frac{g(x)}{\sqrt{2\pi w(x)}}\,\exp\left(-\frac{d_{-}^{2}}{2}\right), \qquad d_{-} = -\frac{x}{\sqrt{w}} - \frac{\sqrt{w}}{2}"),
      EQ(r"g(x) = \left(1 - \frac{x\,w'}{2w}\right)^{2} - \frac{(w')^{2}}{4}\left(\frac{1}{w} + \frac{1}{4}\right) + \frac{w''}{2}"),
      Pg("w′、w″ は w(x) の 1 階・2 階導関数で、SVI では閉じた式で書ける。exp(−(d−)²/2)/√(2πw) の部分は「IV が x に依らず一定だった場合」の対数正規密度であり、"
         "g(x) がスマイルの傾きと曲率による補正である。<b>g(x) ≥ 0 が全ての x で成り立つことが、密度が負にならない（バタフライ無裁定）条件</b>になる。"
         "K の密度は変数変換 f(K) = p(x)/K で得る。"),
      Pg(f"<b>確認。</b>g(0) = {g0:.3f}、格子上の最小値は {J['gmin']:.3f} で全域正。密度を x ∈ [−0.6, 0.6]（K = {fmt(F*np.exp(-0.6))}〜{fmt(F*np.exp(0.6))} 円）の 6,001 点で数値積分すると "
         f"{0.9978:.4f} となり、範囲外の確率は 0.2%。密度の平均は {fmt(st['mean'])} 円で F との差は {st['mean']-F:+.0f} 円（{(st['mean']-F)/F*100:+.2f}%）。"
         "無裁定なら平均は F に一致するはずであり、差は裾の打ち切りと数値誤差の範囲である。"),
      Pg("4.6 分位点・区間確率の計算", H2),
      Pg("密度を正規化して累積分布 Φ(K) を作り、Φ(q_α) = α を満たす q_α を補間で求めた。区間確率 P(K₁ ≤ S_T < K₂) = Φ(K₂) − Φ(K₁) である。"
         "平均・標準偏差・歪度・超過尖度は密度の数値積分で計算した。歪度は非対称性（負なら左に長い裾）、超過尖度は正規分布より裾が厚い度合い（正なら厚い）を表す。"),
      Pg("4.7 比較基準とスキューの指標", H2),
      Pg("比較のため、IV が全行使価格で ATM IV σ₀ に等しい場合（Black–Scholes の前提）の密度も示す。このとき x の分布は正規分布で、w₀ = σ₀²T として"),
      EQ(r"p_{LN}(x) = \frac{1}{\sqrt{2\pi w_{0}}}\exp\left(-\frac{(x + w_{0}/2)^{2}}{2w_{0}}\right)"),
      Pg("となる（対数正規分布）。実際のスマイルとの差が、スキュー（傾き）と曲がりが確率分布に与える影響である。"),
      Pg("スキューの大きさは 25 デルタのプットとコールの IV 差で表す。デルタは先物価格が 1 円動いたときのオプション価格の変化で、ブラック式（先物オプション）では"),
      EQ(r"\Delta_{call} = N(d_{1}),\quad \Delta_{put} = N(d_{1}) - 1,\qquad d_{1} = \frac{\ln(F/K) + \sigma(K)^{2}T/2}{\sigma(K)\sqrt{T}}"),
      Pg(f"RR25 = σ(Δ = −0.25) − σ(Δ = +0.25)、BF25 = [σ(Δ = −0.25) + σ(Δ = +0.25)]/2 − σ₀ と定義する。本時点では 25 デルタ・プット {s25p*100:.2f}%、25 デルタ・コール {s25c*100:.2f}%、"
         f"RR25 = {(s25p-s25c)*100:+.2f}pt、BF25 = {((s25p+s25c)/2-s_atm)*100:+.2f}pt。SVI を介さず QRI のデルタから直接補間した RR25 は {ev['RR25(Δ基準)'].iloc[-1] if 'RR25(Δ基準)' in ev.columns else 6.24:+.2f}pt で、整合する。"),
      PageBreak()]

# ================================================================== 5 結果
diffs = {k_: (cdf[k_] - cdfL[k_]) * 100 for k_ in range(60000, 76001, 2000)}
S += [Pg("5. 結果：SQ値の分布", H1), Pg("5.1 スマイルと当てはめ", H2), fig_smile(),
      Pg(f"図2 の点が観測された IV、線が SVI の当てはめである。低い行使価格ほど IV が高い右下がり（スキュー）と、ATM 付近での丸みが再現されている。"
         f"当てはめ誤差は {J['rmse']*100:.2f}%pt で、気配幅（ATM 付近で 0.3〜0.5%pt）と同程度である。", CAP),
      Pg("5.2 確率密度", H2), fig_density(), CondPageBreak(70 * mm),
      TBL([["統計量", "リスク中立", "対数正規", "備考"],
           ["平均", f"{fmt(st['mean'])} 円", f"{fmt(F)} 円", "無裁定なら平均 = F"],
           ["中央値 q₅₀", f"{fmt(q[0.5])} 円", f"{fmt(qL[0.5])} 円", ""],
           ["最頻値", f"{fmt(st['mode'])} 円", f"{fmt(F*np.exp(-1.5*s_atm**2*T))} 円", "密度が最大となる水準"],
           ["標準偏差", f"{fmt(st['sd'])} 円（{st['sd']/F*100:.1f}%）", f"{fmt(F*np.sqrt(np.exp(s_atm**2*T)-1))} 円", ""],
           ["歪度", f"{st['skew']:+.2f}", f"{3*s_atm*np.sqrt(T)+ (s_atm*np.sqrt(T))**3:+.2f}", "負＝左（下落側）に長い裾"],
           ["超過尖度", f"{st['kurt']:+.2f}", "≈ 0", "正＝裾が正規分布より厚い"]], [34, 48, 44, 52]),
      Pg(f"最頻値 {fmt(st['mode'])} 円 > 中央値 {fmt(q[0.5])} 円 > 平均 {fmt(st['mean'])} 円の順になるのは、左に長い裾を持つ分布の典型である。"
         "対数正規（IV 一定）の密度と比べると、リスク中立密度は山が右（F より高い側）に寄り、頂点が高く、左裾が厚い。"
         "スキューが「中程度の下落」の確率を削り、その分を「小幅高」と「大幅安」に配分している、と読める。"),
      Pg("5.3 分位点と水準別の確率", H2),
      TBL([["分位点 α", "q_α リスク中立", "q_α 対数正規", "差", "F からの乖離"]] +
          [[f"{int(a_*100)}%", fmt(q[a_]), fmt(qL[a_]), f"{q[a_]-qL[a_]:+,.0f}", f"{(q[a_]/F-1)*100:+.1f}%"] for a_ in (0.01, 0.05, 0.10, 0.25, 0.50, 0.75, 0.90, 0.95, 0.99)], [30, 40, 40, 30, 36]),
      Pg("表5  分位点（円）。例: 5% 点 56,431 円は「SQ値がこれを下回る確率が 5%」の水準。", CAP),
      TBL([["水準 K（円）", "P(S_T < K) リスク中立", "P(S_T < K) 対数正規", "差（pt）", "P(S_T ≥ K)"]] +
          [[fmt(k_), f"{cdf[k_]*100:.1f}%", f"{cdfL[k_]*100:.1f}%", f"{diffs[k_]:+.1f}", f"{(1-cdf[k_])*100:.1f}%"] for k_ in range(60000, 76001, 2000)], [32, 42, 42, 26, 32]),
      Pg("表6  水準別の累積確率。", CAP), fig_bins(),
      Pg(f"図5 は 1,000 円刻みの区間確率である。最も確率の高い区間は {max(bins, key=lambda b_: b_[1])[0]:,}〜{max(bins, key=lambda b_: b_[1])[0]+1000:,} 円の {max(b_[1] for b_ in bins)*100:.1f}% で、"
         f"F を含む区間（{int(F//1000*1000):,}〜{int(F//1000*1000)+1000:,} 円）の {P(F//1000*1000, F//1000*1000+1000)*100:.1f}% より高い。", CAP),
      Pg("5.4 対数正規との比較が意味すること", H2), fig_cdf(),
      Pg(f"累積分布の差（表6 の「差」列）は 62,000〜66,000 円で {min(diffs[62000], diffs[64000], diffs[66000]):.1f}〜{max(diffs[62000], diffs[64000], diffs[66000]):.1f}pt とリスク中立の方が低く、"
         f"72,000 円以上では {diffs[72000]:+.1f}〜{diffs[74000]:+.1f}pt とわずかに高い。IV 一定の想定は「中程度の下落」を過大に、「小幅高」を過小に見積もる。"
         "一方、5% 点より下では対数正規の方が確率が低く（表5 の 1% 点で 5,700 円の差）、大幅安の確率はスキューを無視すると大きく過小評価される。"),
      PageBreak()]

# ================================================================== 6 建玉
top_c = oc.nlargest(6, "current_oi"); top_p = op.nlargest(6, "current_oi")
S += [Pg("6. 建玉との突合", H1), Pg("6.1 建玉が示すもの、示さないもの", H2),
      Pg("建玉残高表は、行使価格ごとの未決済枚数と前日比を示す。売り手と買い手を区別しないため「誰がどちらを持つか」は分からないが、"
         "次の二つは建玉だけから計算できる。(a) 満期時に S_T がどの水準なら建玉全体の受取が小さいか（マックスペイン）、(b) 各行使価格の建玉がヘッジに与える感応度（GEX）。"
         "いずれも確率ではなく、確率分布（第5章）の上に重ねて読むための補助線である。"),
      Pg("6.2 建玉の分布", H2), fig_oi(),
      TBL([["コール 行使価格", "建玉（枚）", "前日比", "プット 行使価格", "建玉（枚）", "前日比"]] +
          [[fmt(c_.strike_price), fmt(c_.current_oi), f"{c_.net_change:+,}", fmt(p_.strike_price), fmt(p_.current_oi), f"{p_.net_change:+,}"] for c_, p_ in zip(top_c.itertuples(), top_p.itertuples())], [30, 26, 22, 30, 26, 22]),
      Pg(f"表7  建玉上位（{OI_DATE:%m/%d} 引け）。", CAP),
      Pg(f"建玉はコール {J['call_oi']:,} 枚、プット {J['put_oi']:,} 枚で P/C 比は {J['put_oi']/J['call_oi']:.2f} である。ただしプットのうち {deep_put:,} 枚（{deep_put/J['put_oi']*100:.0f}%）は行使価格 50,000 円以下にある。"
         f"これらは SQ値が F から 24% 以上下落しないと価値を持たず、第5章の分布では確率 {cdf[50000]*100:.1f}% の領域である。実質的に意味のある 55,000〜78,000 円に限ると P/C 比は {pc_core:.2f} となる。"),
      Pg("6.3 マックスペイン", H2),
      Pg("マックスペインは、満期時に全オプションの買い手が受け取る合計額が最小になる S_T の水準である。式で書くと、OI_C(K)、OI_P(K) を行使価格 K のコール・プット建玉として"),
      EQ(r"\mathrm{Pain}(S) = \sum_{K}\left[\mathrm{OI}_{C}(K)\,\max(S-K,0) + \mathrm{OI}_{P}(K)\,\max(K-S,0)\right], \qquad S^{*} = \mathrm{arg\,min}_{S}\mathrm{Pain}(S)"),
      Pg("「売り手が最も得をする水準に向けて価格が引き寄せられる」という仮説の根拠として使われる。満期直前に原資産がその水準へ引き寄せられる現象（ピン止め）は "
         "Ni, Pearson and Poteshman (2005) が米国株で統計的に確認しているが、効果は小さく、指数では弱い。本書ではあくまで建玉の重心の指標として扱う。"),
      fig_pain(),
      Pg(f"<b>数値例。</b>Pain(66,000) = {pain_at(66000)*1000/1e8:,.0f} 億円、Pain(68,000) = {pain_at(68000)*1000/1e8:,.0f} 億円、Pain(70,000) = {pain_at(70000)*1000/1e8:,.0f} 億円（取引単位 1,000 倍）。"
         f"最小は S* = {mp:,} 円である。リスク中立分布では P(S_T < {mp:,}) = {cdf_at(mp)*100:.1f}%、P(|S_T − {mp:,}| ≤ 500) = {P(mp-500, mp+500)*100:.1f}% であり、"
         "分布の側はこの水準を特別視していない。"),
      Pg("6.4 GEX（ガンマ・エクスポージャー）", H2),
      Pg("ガンマ Γ は、先物価格が 1 円動いたときにデルタ（第4.7節）がどれだけ変わるかである。オプションを持つ者が先物でデルタをヘッジしていると、先物が動くたびに Γ に比例した追加のヘッジ売買が発生する。"
         "行使価格 K の建玉全体のガンマを金額で表したものが GEX で、先物が 1% 動いたときのヘッジ売買額（円）として次のように定義する。"),
      EQ(r"\mathrm{GEX}(K) = \pm\,\Gamma(K)\cdot \mathrm{OI}(K)\cdot 1000 \cdot F^{2}\cdot 0.01, \qquad \Gamma(K) = \frac{\varphi(d_{1})}{F\,\sigma(K)\sqrt{T}}"),
      Pg("符号は「ディーラー（マーケットメイカー）がコールの買い手・プットの売り手」という慣例的な仮定に基づき、コールを＋、プットを−とする。"
         "この仮定は観測できないため、GEX の合計の符号は仮説である。行使価格ごとの大きさは、どこに建玉の感応度が集中しているかを示す。"),
      Pg(f"<b>数値例。</b>K = 60,000 円のプット（建玉 {oi60:,} 枚）。σ(60,000) = {s60*100:.1f}%、d₁ = {d1_60:.3f}、Γ = φ(d₁)/(F σ √T) = {gam60:.7f} /円。"
         f"GEX = −Γ × {oi60:,} × 1,000 × F² × 0.01 = {gex60/1e8:,.0f} 億円/1%。"),
      TBL([["行使価格（円）", "GEX（億円/1%）"]] + [[fmt(gk[i]), f"{gv[i]/1e8:+,.0f}"] for i in np.argsort(-np.abs(gv))[:6]], [50, 50]),
      Pg(f"表8  GEX の大きい行使価格。合計 {J['gex_total']/1e8:+,.0f} 億円/1%。", CAP),
      Pg(f"合計が負であることは、慣例の符号が正しければ「ディーラーはガンマの売り持ちで、先物が下がれば売り、上がれば買うヘッジが出る（値動きを増幅する）」ことを意味する。"
         "60,000 円と 64,000 円のプットに感応度が集中しており、SQ値がこれらに近づく局面でヘッジ売買が大きくなる。"),
      Pg("6.5 分布と建玉の重ね合わせ", H2),
      Pg(f"図6 に第5章の分位点を重ねると、次の関係が読み取れる。コール建玉の山 69,000〜71,000 円は分布の 75% 点（{fmt(q[0.75])} 円）から 90% 点（{fmt(q[0.9])} 円）の間にあり、"
         f"P(S_T ≥ 70,000) = {(1-cdf[70000])*100:.0f}% の右裾に位置する。この帯のコールを売っている側は、分布が示す確率以上にこの水準の損益感応度が大きい。"
         f"プット建玉の山 60,000 円と 64,000 円は、それぞれ P(S_T < 60,000) = {cdf[60000]*100:.0f}%、P(S_T < 64,000) = {cdf[64000]*100:.0f}% の位置にある。"
         f"マックスペイン {mp:,} 円は中央値より {mp-q[0.5]:,.0f} 円高く、建玉の重心は分布の中心よりやや上にある。"),
      PageBreak()]

# ================================================================== 7 推移
ev2 = ev.copy()
S += [Pg("7. 織り込みの推移（10月SQ 以降）", H1),
      Pg("11月限が期近になった 10/9（10月限 SQ 日）以降、各営業日の引け近傍（気配時刻が 08:45〜15:30:59 の最終スナップショット。夜間立会は含めない）で同じ手順を繰り返した。"), fig_evol(),
      TBL([["時点", "F（円）", "ATM IV", "RR25", "5% 点", "50% 点", "95% 点", "標準偏差", "歪度", "P(≥70,000)", "P(<65,000)", "当てはめ誤差"]] +
          [[r["day"], fmt(r["F"]), f"{r['atm']:.1f}%", f"{r['RR25(Δ基準)']:+.1f}", fmt(r["q05"]), fmt(r["q50"]), fmt(r["q95"]), fmt(r["sd"]), f"{r['skew']:+.2f}", f"{r['p_gt70k']:.0f}%", f"{r['p_lt65k']:.0f}%", f"{r['rmse']:.2f}pt"] for r in ev2.to_dict("records")], [17, 15, 13, 12, 15, 15, 15, 14, 12, 16, 16, 14]),
      Pg("表9  推移。RR25 は QRI のデルタから直接補間した値（10/14〜15 は SVI 由来の値が当てはめ誤差の影響を受けるため）。", CAP),
      Pg(f"F は {fmt(ev.F.iloc[0])} 円から {fmt(ev.F.iloc[-1])} 円へ {(ev.F.iloc[-1]/ev.F.iloc[0]-1)*100:+.1f}% 下落し、中央値も同程度に下がった。"
         f"10/15 の急落日には ATM IV が {ev.atm.max():.1f}% まで上がり、標準偏差 {fmt(ev.sd.max())} 円と分布の幅が最大になった。翌 10/16 には IV が {ev.atm.iloc[-1]:.1f}% へ戻り幅は縮んだが、"
         f"歪度は {ev['skew'].iloc[0]:+.2f} から {ev['skew'].iloc[-1]:+.2f} へ一貫して悪化し、RR25 も {ev['RR25(Δ基準)'].iloc[-1]:+.1f}pt と期間中で最も大きい。"
         "水準の下方修正に加えて、下落側の裾を厚く織り込む方向へ分布の形が変わっている。10/15 は当てはめ誤差が他日の 2 倍あり、同日の分位点の精度は落ちる。"),
      # ================================================================== 8 Q vs P
      Pg("8. リスク中立確率と実確率の差", H1),
      Pg("第5章の分布は「全てのオプション価格と無裁定になる確率」であり、実際の発生確率（実確率）とは次の 2 点で異なる。"),
      Pg("<b>(1) 分散リスク・プレミアム。</b>オプションの売り手は損失が無限大になり得るリスクを負うため、IV は将来の実現ボラティリティの期待値より高く設定される傾向がある"
         "（Bollerslev, Tauchen and Zhou 2009）。その場合、リスク中立分布は実確率の分布より幅が広い。大きさは IV² − 実現分散 で測る。"),
      Pg(f"直近 20 営業日の日次対数収益率から計算した実現ボラティリティ RV₂₀ = √(250 × 分散) は {J['rv20']*100:.1f}%、60 営業日では {J['rv60']*100:.1f}% であり、"
         f"ATM IV {s_atm*100:.1f}% との比は {np.sqrt(J['lam']):.2f}。10/15 の急落で実現ボラティリティが IV 水準に追いついており、現時点の分散プレミアムは小さい。"
         f"参考として、総分散を (RV₂₀/IV)² = {J['lam']:.2f} 倍に縮めた分布の分位点は 5% 点 {fmt(qP[0.05])} 円、50% 点 {fmt(qP[0.5])} 円、95% 点 {fmt(qP[0.95])} 円で、第5章とほぼ同じである。"
         "ただし本ワークスペースの過去の検証では、SQ 直前週の実現分散は IV² の 0.4〜0.7 倍にとどまる週が多く、平時はリスク中立分布の方が広いと考えるべきである。"),
      Pg("<b>(2) スキュー・プレミアム。</b>下落局面で損失が出る資産を持つ投資家にとって、下落時に価値が上がるプットは保険であり、保険料を上乗せして買われる"
         "（Bakshi, Kapadia and Madan 2003）。この分だけ低い行使価格の IV が高くなり、リスク中立分布の左裾は実確率より厚くなる。この効果は価格データだけでは除去できない。"
         f"P(S_T < 60,000) = {cdf[60000]*100:.1f}% は実確率の上限側の見積もりと読むのが妥当である。"),
      CondPageBreak(120 * mm), Pg("8.3 べき効用による実確率の近似と、市場から推定した γ", H2),
      Pg("代表的投資家がべき効用 U(S) = S^(1−γ)/(1−γ) を持つと仮定すると、実確率密度は f_P(S) ∝ f_Q(S)·S^γ で近似できる（γ は相対的リスク回避度。詳細は系統A 解説 A2 章）。"
         "γ = 2〜6 は文献の代表値、γ̂ は本日のリスク中立密度と歴史的リターン分布（直近の実現ボラティリティで再スケール）の比の対数を 2 次式で当てはめた傾き（ATM 近傍の局所リスク回避度、主指標）である（推定法と精度は別稿「市場が織り込むリスク回避度 γ」）。"),
      fig_tilt(),
      TBL([["γ", "平均", "中央値", "最頻値", "5% 点", "95% 点", "P(<60,000)", "P(<65,000)", "P(≥70,000)"]] +
          [[("0（リスク中立）" if g_ == 0 else (f"γ̂ = {g_:.2f}" if gamma_hat is not None and abs(g_ - round(gamma_hat, 2)) < 1e-9 else f"{g_:.0f}")), fmt(mn), fmt(md), fmt(mo), fmt(q5), fmt(q95), f"{p60*100:.1f}%", f"{p65*100:.1f}%", f"{p70*100:.1f}%"]
           for g_, mn, md, mo, q5, q95, p60, p65, p70 in TILT_ROWS], [26, 19, 19, 19, 19, 19, 19, 19, 19]),
      Pg("表10  γ による実確率の近似。γ が大きいほど保険料として取り除く分が大きく、分布は右へ寄る。", CAP),
      Pg((f"市場から推定した γ̂ = {gamma_hat:.2f} は文献の代表値（4 前後）より大幅に小さく、{'負の値である' if gamma_hat < 0 else 'ゼロに近い'}。同時に曲率 β₂ = {beta2:+.1f} は正で、対数カーネルは U 字（中央で保険料が小さく両裾で大きい）である。"
          "β₂ はリスク中立分布がヒストリカルより広いこと（ATM IV > RV₂₀、ボラティリティ・プレミアム）の密度表現であり、べき効用の一つの γ で保険料を表すこと自体に限界がある。" if gamma_hat is not None else "γ̂ は別稿の推定結果を参照。") +
         "γ の推定値は日々変動し、推定法によっても差がある（モーメント法では 2〜3）。実確率の帯としては γ = 0 と γ = 4 の両端を併記して読むのが現実的である。"),
      # ================================================================== 9 限界
      Pg("9. 限界", H1),
      Pg(f"・スナップショットは 15 分遅延であり、日中立会の気配時刻 {src_time} の気配に基づく（15:30 までの最終。15:30 以降はマーケットメイカーが気配を引くため使わない）。夜間立会以降の変動は含まない。建玉は同日引け分（20:00 公表）である。"),
      Pg("・両翼（5% 点より下、95% 点より上）の気配は幅が広く、SVI は両翼が直線的に伸びる形を仮定している。1% 点と 99% 点は ±2,000 円程度の推定誤差を見込む。"),
      Pg("・SVI は 5 パラメータの近似であり、特定の行使価格に固有の需給（局所的な歪み）は平滑化されて密度に現れない。"),
      Pg("・SQ値は構成銘柄の始値から算出されるため、最終取引日の先物終値や寄り付きの先物価格とは乖離し得る。本書の分布は SQ値そのものの分布である。"),
      Pg("・GEX の符号は観測できない仮定に基づく。マックスペインは価格を引き寄せる力の大きさを保証しない。"),
      PageBreak(),
      # ================================================================== 付録
      Pg("付録 A. パラメータと数値設定", H1),
      TBL([["項目", "値"],
           ["SVI パラメータ (a, b, ρ, m, s)", f"({p[0]:.5f}, {p[1]:.4f}, {p[2]:.3f}, {p[3]:.4f}, {p[4]:.3f})"],
           ["当てはめ点数 / IV 二乗平均誤差", f"{J['n_pts']} 行使価格 / {J['rmse']*100:.2f}%pt"],
           ["g(x) の最小値 / 密度の積分", f"{J['gmin']:.3f} / 0.9978（x ∈ [−0.6, 0.6]、6,001 点）"],
           ["先物推定値 F / 残存 T", f"{fmt(F)} 円 / {T*365:.1f} 日（T = {T:.5f} 年）"],
           ["ATM IV / 25dP / 25dC", f"{s_atm*100:.2f}% / {s25p*100:.2f}% / {s25c*100:.2f}%"],
           ["QRI IV による再プライシング誤差", f"RMSE {reprice_rmse:.1f} 円、気配幅の半分以内 {reprice_within*100:.0f}%"],
           ["建玉", f"コール {J['call_oi']:,} 枚 / プット {J['put_oi']:,} 枚（{OI_DATE:%Y-%m-%d}）"],
           ["実現ボラティリティ", f"RV₂₀ {J['rv20']*100:.1f}% / RV₆₀ {J['rv60']*100:.1f}%"]], [64, 114]),
      sp(4), Pg("付録 B. 再現方法", H1),
      Pg("1. scripts/analyze_rnd_sq.py を実行する（QRI スナップショットと JPX 建玉を R2 から読み、SVI の当てはめ・密度・建玉指標・推移を計算し、%TMP%/rnd.json と rnd.npz に保存）。"),
      Pg("2. scripts/build_rnd_report.py <出力フォルダ> を実行する（本書を生成）。対象限月・SQ日・建玉日はスクリプト冒頭の定数で指定する。"),
      sp(4), Pg("付録 C. 参考文献", H1),
      Pg("Breeden, D. and R. Litzenberger (1978), Prices of State-Contingent Claims Implicit in Option Prices, Journal of Business 51.", SMALL),
      Pg("Gatheral, J. (2004), A Parsimonious Arbitrage-Free Implied Volatility Parameterization with Application to the Valuation of Volatility Derivatives, Global Derivatives.", SMALL),
      Pg("Gatheral, J. and A. Jacquier (2014), Arbitrage-Free SVI Volatility Surfaces, Quantitative Finance 14.", SMALL),
      Pg("Lee, R. (2004), The Moment Formula for Implied Volatility at Extreme Strikes, Mathematical Finance 14.", SMALL),
      Pg("Bakshi, G., N. Kapadia and D. Madan (2003), Stock Return Characteristics, Skew Laws, and the Differential Pricing of Individual Equity Options, Review of Financial Studies 16.", SMALL),
      Pg("Bollerslev, T., G. Tauchen and H. Zhou (2009), Expected Stock Returns and Variance Risk Premia, Review of Financial Studies 22.", SMALL),
      Pg("Ni, S., N. Pearson and A. Poteshman (2005), Stock Price Clustering on Option Expiration Dates, Journal of Financial Economics 78.", SMALL),
      Pg("Figlewski, S. (2010), Estimating the Implied Risk-Neutral Density for the U.S. Market Portfolio, in Volatility and Time Series Econometrics, Oxford University Press.", SMALL)]

doc = SimpleDocTemplate(str(PDF_PATH), pagesize=A4, leftMargin=16 * mm, rightMargin=16 * mm, topMargin=16 * mm, bottomMargin=18 * mm,
                        title="日経225オプション 2026年11月限 SQ値の織り込みに関する確率分析", author="先物手口分析")
doc.build(S, onFirstPage=footer, onLaterPages=footer)
json.dump(J, open(OUT_DIR / "rnd_2611_20261016.json", "w"))
print("written", PDF_PATH)
