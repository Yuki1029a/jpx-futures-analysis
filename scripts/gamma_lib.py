"""市場が織り込むリスク回避度 γ の推定に使う関数（gamma_calc.py / gamma_daily.py 共通）。

方法 I  カーネル法 : ln[f_Q(x)/f_P(x)] = a − γ x を中央部で回帰。f_P は歴史的 n 日リターン（開始時 RV20 で標準化）を現在の RV で再スケールした KDE。
方法 II モーメント法: γ = (Skew_P − Skew_Q) / (s_P (Kurt_P − 3))。
"""
from __future__ import annotations

import numpy as np
import pandas as pd
from scipy.stats import gaussian_kde

from rnd_lib import SQ_DATES


class Physical:
    """日経平均の日次終値から、実確率側の材料（RV、標準化 n 日リターン）を作る。"""

    def __init__(self, close: pd.Series):
        self.lp = np.log(close.dropna()); self.r1 = self.lp.diff().dropna()
        self.rv20 = self.r1.rolling(20).std() * np.sqrt(250); self.rv60 = self.r1.rolling(60).std() * np.sqrt(250)

    def rv_at(self, d: pd.Timestamp, n: int = 20) -> float:
        s = self.rv20 if n == 20 else self.rv60
        return float(s[s.index <= d].iloc[-1])

    def ret(self, d: pd.Timestamp, n: int = 20) -> float:
        x = self.lp[self.lp.index <= d]
        return float(x.iloc[-1] - x.iloc[-1 - n]) if len(x) > n else np.nan

    def sample(self, n_days: int, rv_now: float, start="2000-01-01") -> np.ndarray:
        """歴史的 n 日対数リターンを開始時点の RV20 で標準化し、rv_now で再スケールした標本。"""
        rn = (self.lp.shift(-n_days) - self.lp); y = (rn / self.rv20).dropna(); y = y[y.index >= start]
        return (y * rv_now).values


def q_moments(K, fK, F):
    dK = np.gradient(K); m = fK * dK; m = m / m.sum(); x = np.log(K / F)
    mx = (x * m).sum(); s = np.sqrt(((x - mx) ** 2 * m).sum())
    return float(mx), float(s), float((((x - mx) / s) ** 3 * m).sum()), float((((x - mx) / s) ** 4 * m).sum())


def p_moments(xs):
    s = xs.std(); z = (xs - xs.mean()) / s
    return float(xs.mean()), float(s), float((z ** 3).mean()), float((z ** 4).mean())


def gamma_kernel(K, fK, F, xP, lo=0.05, hi=0.95, bw=None) -> dict:
    """方法 I。戻り値: gamma, a（切片）, quad（2 次項）, gamma_quadfit, r2, xlo, xhi と曲線（lnM, x, fP, fQ, xgrid）。"""
    dK = np.gradient(K); mQ = fK * dK; mQ = mQ / mQ.sum(); cdf = np.cumsum(mQ); x = np.log(K / F)
    fP = gaussian_kde(xP, bw_method=bw)(x); fQx = mQ / np.gradient(x)
    sel = (cdf >= lo) & (cdf <= hi) & (fP > 1e-8) & (fQx > 1e-8)
    lnM = np.log(fQx[sel]) - np.log(fP[sel]); w = mQ[sel]; xx = x[sel]; Wm = np.sqrt(w)[:, None]
    A = np.vstack([np.ones(sel.sum()), xx]).T; coef, *_ = np.linalg.lstsq(A * Wm, lnM * Wm[:, 0], rcond=None)
    pred = A @ coef; r2 = 1 - np.sum(w * (lnM - pred) ** 2) / np.sum(w * (lnM - np.average(lnM, weights=w)) ** 2)
    A2 = np.vstack([np.ones(sel.sum()), xx, xx ** 2]).T; coef2, *_ = np.linalg.lstsq(A2 * Wm, lnM * Wm[:, 0], rcond=None)
    return dict(gamma=float(-coef[1]), a=float(coef[0]), quad=float(coef2[2]), gamma_quadfit=float(-coef2[1]), r2=float(r2), xlo=float(xx.min()), xhi=float(xx.max()),
                lnM=lnM, x=xx, fP=fP, fQ=fQx, xgrid=x)


def gamma_moment(K, fK, F, xP) -> float:
    _, _, skq, kuq = q_moments(K, fK, F); _, sp_, skp, kup = p_moments(xP)
    return float((skp - skq) / (sp_ * (kup - 3))) if kup > 3.2 else float("nan")


def front_cm(day: str, min_days: int = 8) -> str | None:
    d = pd.Timestamp(day)
    for cm in sorted(SQ_DATES, key=lambda c: SQ_DATES[c]):
        if (SQ_DATES[cm] - d).days >= min_days:
            return cm
    return None
