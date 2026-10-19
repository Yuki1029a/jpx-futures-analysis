"""日経225ミニオプション（週次満期: 水曜・金曜 SQ）の建玉データ。

出所: JPX 建玉残高表 別紙2（{yyyymmdd}open_interest.xlsx の 3 番目のシート）。行使価格別の
取引高・当日建玉・前日比・前日建玉がプット/コール別に載る。銘柄コードは
'NK225 MINI P261022-64000' の形式で、6 桁の日付は最終取引日（SQ 算出日の前営業日）。
水曜 SQ の銘柄は火曜、金曜 SQ の銘柄は木曜の日付を持つ。取引単位は指数 × 100。
取引参加者別取引高（手口）にはミニオプションが含まれない。IV（QRI）は月次限月のみ。
"""
from __future__ import annotations

import logging
from datetime import date, timedelta

import numpy as np
import pandas as pd

from data import fetcher
from data.parser_daily_oi import parse_daily_mini_oi_excel
from models import DailyOIBalance

logger = logging.getLogger(__name__)
MULT = 100
_cache: dict[date, list[DailyOIBalance]] = {}
_WD = "月火水木金土日"


def load_mini_oi(trade_date: date) -> list[DailyOIBalance]:
    """その日の別紙2の全レコード（キャッシュ/R2 から。未公表・未収集なら空）。"""
    if trade_date in _cache:
        return _cache[trade_date]
    content = fetcher.download_daily_oi_excel(trade_date)
    recs = parse_daily_mini_oi_excel(content) if content else []
    if recs:
        _cache[trade_date] = recs
    return recs


def latest_available_date(start: date | None = None, lookback: int = 10) -> date | None:
    """建玉残高表が存在する直近の営業日。"""
    d = start or date.today()
    for i in range(lookback):
        x = d - timedelta(days=i)
        if x.weekday() < 5 and load_mini_oi(x):
            return x
    return None


def expiry_info(code: str) -> dict:
    """'261022' → 最終取引日・SQ日（翌暦日）・曜日ラベル。"""
    last = date(2000 + int(code[:2]), int(code[2:4]), int(code[4:6])); sq = last + timedelta(days=1)
    return dict(code=code, last_trading=last, sq=sq, label=f"SQ {sq:%m/%d}({_WD[sq.weekday()]})・最終取引 {last:%m/%d}({_WD[last.weekday()]})")


def list_expiries(trade_date: date, min_oi: int = 1) -> list[str]:
    """その日の別紙2にある満期コード（建玉合計が min_oi 以上）を昇順で返す。"""
    recs = load_mini_oi(trade_date)
    tot: dict[str, int] = {}
    for r in recs:
        tot[r.contract_month] = tot.get(r.contract_month, 0) + r.current_oi
    return sorted(k for k, v in tot.items() if v >= min_oi)


def strike_frame(trade_date: date, expiry: str) -> pd.DataFrame:
    """満期 expiry の行使価格別の表: strike, put_oi, put_chg, put_vol, call_oi, call_chg, call_vol。"""
    recs = [r for r in load_mini_oi(trade_date) if r.contract_month == expiry]
    if not recs:
        return pd.DataFrame(columns=["strike", "put_oi", "put_chg", "put_vol", "call_oi", "call_chg", "call_vol"])
    df = pd.DataFrame([{"strike": r.strike_price, "type": r.option_type, "oi": r.current_oi, "chg": r.net_change, "vol": r.trading_volume} for r in recs])
    pv = df.pivot_table(index="strike", columns="type", values=["oi", "chg", "vol"], aggfunc="sum", fill_value=0)
    out = pd.DataFrame({"strike": pv.index})
    for side, lab in (("PUT", "put"), ("CALL", "call")):
        for col in ("oi", "chg", "vol"):
            out[f"{lab}_{col}"] = pv[(col, side)].values if (col, side) in pv.columns else 0
    return out.sort_values("strike").reset_index(drop=True)


def max_pain(frame: pd.DataFrame) -> tuple[int | None, pd.DataFrame]:
    """満期時に全建玉の受取額が最小になる水準と、水準ごとの受取額（指数 1 単位あたり × 建玉）。"""
    if frame.empty or (frame.put_oi.sum() + frame.call_oi.sum()) == 0:
        return None, pd.DataFrame()
    S = np.arange(frame.strike.min() - 500, frame.strike.max() + 501, 125)
    pain = np.array([(frame.call_oi * np.maximum(s - frame.strike, 0)).sum() + (frame.put_oi * np.maximum(frame.strike - s, 0)).sum() for s in S])
    return int(S[int(np.argmin(pain))]), pd.DataFrame({"S": S, "pain": pain})


def history(expiry: str, end: date, n_days: int = 15) -> pd.DataFrame:
    """満期 expiry の建玉合計（プット/コール）・新規建玉・マックスペインの日次推移（直近 n 営業日）。"""
    rows = []; d = end; tried = 0
    while len(rows) < n_days and tried < n_days * 2 + 10:
        if d.weekday() < 5:
            f = strike_frame(d, expiry)
            if not f.empty:
                mp, _ = max_pain(f)
                rows.append(dict(date=d, put_oi=int(f.put_oi.sum()), call_oi=int(f.call_oi.sum()), put_vol=int(f.put_vol.sum()), call_vol=int(f.call_vol.sum()),
                                 put_new=int(f.put_chg.clip(lower=0).sum()), call_new=int(f.call_chg.clip(lower=0).sum()), max_pain=mp))
        d -= timedelta(days=1); tried += 1
    return pd.DataFrame(rows).sort_values("date").reset_index(drop=True) if rows else pd.DataFrame()
