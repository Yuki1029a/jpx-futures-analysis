"""J-Quants の日経225オプション日次データ（/v2/derivatives/bars/daily/options/225）を取得し、R2 とローカルに parquet で保存する。

用途: 立会内取引高（VoOA）と立会外（Vo − VoOA）の内訳、清算値・理論価格・IV を IV モニター等で使う。
保存先: R2 jquants/op225/YYYYMMDD.parquet、ローカル cache/jquants/op225/YYYYMMDD.parquet
列: date, cm('YYMM'), option_type('PUT'/'CALL'), strike(int), vo, vooa, vo_off, oi, settle, theo, iv, basevol, underpx, ir, ltd, sqd, code
API キー: 環境変数 JQUANTS_API_KEY → <repo>/.env → ../amova_クオンツファンド解析/.env の順に探す。
使い方: python scripts/fetch_jquants_options.py [--day YYYYMMDD | --since YYYYMMDD] [--upload] [--force]
既定: 直近 5 平日のうち未保存の日を取得（J-Quants は翌 03:00 頃更新）。
"""
from __future__ import annotations

import argparse
import io
import logging
import os
import re
import sys
import time
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import requests

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from data.r2_storage import r2_get, r2_put  # noqa: E402

logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s"); log = logging.getLogger("jq")
BASE = "https://api.jquants.com/v2/derivatives/bars/daily/options/225"
R2_PREFIX = "jquants/op225"; LOCAL = ROOT / "cache" / "jquants" / "op225"


def api_key() -> str | None:
    k = os.environ.get("JQUANTS_API_KEY")
    if k:
        return k.strip()
    for p in (ROOT / ".env", ROOT.parent / "amova_クオンツファンド解析" / ".env"):
        try:
            m = re.search(r"JQUANTS_API_KEY\s*=\s*\"?([A-Za-z0-9_\-]+)", p.read_text(encoding="utf-8"))
            if m:
                return m.group(1)
        except Exception:  # noqa: BLE001
            continue
    return None


def fetch_day(day: str, key: str) -> pd.DataFrame | None:
    """1 日分を全ページ取得。契約範囲外・休日は None。"""
    rows, pk = [], None
    for _ in range(50):
        params = {"date": day}
        if pk:
            params["pagination_key"] = pk
        r = requests.get(BASE, headers={"x-api-key": key}, params=params, timeout=90)
        if r.status_code == 429:
            time.sleep(10); continue
        if r.status_code != 200:
            log.warning("%s: HTTP %s %s", day, r.status_code, r.text[:120]); return None
        j = r.json(); rows += j.get("data", []); pk = j.get("pagination_key")
        if not pk:
            break
    if not rows:
        return None
    df = pd.DataFrame(rows); df = df[df.EmMrgnTrgDiv == "002"].copy()
    out = pd.DataFrame({
        "date": df.Date.str.replace("-", ""), "cm": df.CM.str[2:4] + df.CM.str[5:7], "option_type": df.PCDiv.map({"1": "PUT", "2": "CALL"}),
        "strike": pd.to_numeric(df.Strike, errors="coerce").astype("Int64"), "vo": pd.to_numeric(df.Vo, errors="coerce"), "vooa": pd.to_numeric(df.VoOA, errors="coerce"),
        "oi": pd.to_numeric(df.OI, errors="coerce"), "settle": pd.to_numeric(df.Settle, errors="coerce"), "theo": pd.to_numeric(df.Theo, errors="coerce"),
        "iv": pd.to_numeric(df.IV, errors="coerce"), "basevol": pd.to_numeric(df.BaseVol, errors="coerce"), "underpx": pd.to_numeric(df.UnderPx, errors="coerce"),
        "ir": pd.to_numeric(df.IR, errors="coerce"), "ltd": df.LTD, "sqd": df.SQD, "code": df.Code})
    out["vo_off"] = out.vo - out.vooa
    return out.reset_index(drop=True)


def stored(day: str, upload: bool) -> bool:
    if (LOCAL / f"{day}.parquet").exists():
        return True
    return upload and r2_get(f"{R2_PREFIX}/{day}.parquet") is not None


def save(day: str, df: pd.DataFrame, upload: bool) -> None:
    LOCAL.mkdir(parents=True, exist_ok=True); buf = io.BytesIO(); df.to_parquet(buf, index=False); body = buf.getvalue()
    (LOCAL / f"{day}.parquet").write_bytes(body)
    if upload:
        r2_put(f"{R2_PREFIX}/{day}.parquet", body)


def main() -> int:
    ap = argparse.ArgumentParser(); ap.add_argument("--day"); ap.add_argument("--since"); ap.add_argument("--upload", action="store_true"); ap.add_argument("--force", action="store_true")
    a = ap.parse_args(); key = api_key()
    if not key:
        log.warning("JQUANTS_API_KEY not found; skip"); return 0
    today = date.today()
    if a.day:
        days = [a.day]
    elif a.since:
        d0 = date(int(a.since[:4]), int(a.since[4:6]), int(a.since[6:8])); days = [(d0 + timedelta(days=i)).strftime("%Y%m%d") for i in range((today - d0).days + 1) if (d0 + timedelta(days=i)).weekday() < 5]
    else:
        days = [(today - timedelta(days=i)).strftime("%Y%m%d") for i in range(1, 8) if (today - timedelta(days=i)).weekday() < 5][:5][::-1]
    n = 0
    for d in days:
        if not a.force and stored(d, a.upload):
            continue
        df = fetch_day(d, key)
        if df is None or df.empty:
            log.info("%s: no data", d); continue
        save(d, df, a.upload); n += 1; log.info("%s: %d rows (vo %d, 立会内 %d)", d, len(df), int(df.vo.sum()), int(df.vooa.sum()))
        time.sleep(0.6)  # Standard プランは 120 回/分
    log.info("done: %d days", n); return 0


if __name__ == "__main__":
    sys.exit(main())
