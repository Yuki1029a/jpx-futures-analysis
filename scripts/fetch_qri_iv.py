"""Scrape NK225 Option IV/Greeks from QRI (15-min delayed) and save as Parquet.

Source: https://svc.qri.jp/jpx/nkopm/[index]
  index=0 (or empty) → 期近、index=1, 2, ... → 順に次の限月。
  Page header explicitly states minimum 15-minute delay.
  Requires Referer: https://www.jpx.co.jp/

Output:
  cache/qri_iv/raw/YYYYMMDD/YYYYMMDD_HHMM.parquet
    One file per fetch run, all months × all strikes in one table.
    Append-friendly: each row carries fetch_time + source_update_time so
    downstream code can build time-series easily.

Usage:
  python scripts/fetch_qri_iv.py             # one-shot fetch
  python scripts/fetch_qri_iv.py --loop      # run every 15 min until Ctrl+C
  python scripts/fetch_qri_iv.py --max 5     # probe up to /jpx/nkopm/4 (default 8)
"""
from __future__ import annotations

import argparse
import logging
import re
import sys
import time
from datetime import datetime, date
from pathlib import Path
from typing import Optional

import pandas as pd
import requests
from bs4 import BeautifulSoup


_BASE = "https://svc.qri.jp/jpx/nkopm/"
_REFERER = "https://www.jpx.co.jp/"
_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0 Safari/537.36"
)
_DEFAULT_MAX_INDEX = 8  # probe /0../7 by default; stop on first failure

_PROJECT_ROOT = Path(__file__).resolve().parent.parent
_OUT_ROOT = _PROJECT_ROOT / "cache" / "qri_iv" / "raw"

logger = logging.getLogger("fetch_qri_iv")


# ---------------------------------------------------------------------------
# HTTP
# ---------------------------------------------------------------------------

def _fetch_html(index: int, timeout: int = 30) -> Optional[str]:
    """GET the per-month HTML. Returns text or None on failure / not-found."""
    suffix = "" if index == 0 else str(index)
    url = _BASE + suffix
    try:
        r = requests.get(
            url,
            headers={"User-Agent": _UA, "Referer": _REFERER},
            timeout=timeout,
        )
    except requests.RequestException as e:
        logger.warning("HTTP error for %s: %s", url, e)
        return None
    if r.status_code != 200:
        logger.warning("Non-200 (%d) for %s", r.status_code, url)
        return None
    # The error page is short — sanity-check for actual content
    if "price-info-scroll" not in r.text:
        logger.info("No price-info-scroll in %s → treat as no contract", url)
        return None
    return r.text


# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------

# Numeric helpers
_NUM_RE = re.compile(r"-?\d{1,3}(?:,\d{3})*(?:\.\d+)?|-?\d+(?:\.\d+)?")
_PCT_RE = re.compile(r"-?\d+(?:\.\d+)?")


def _num(s: Optional[str]) -> Optional[float]:
    if s is None:
        return None
    s = s.replace(",", "").strip()
    if s in ("", "-", "−"):
        return None
    m = _NUM_RE.match(s)
    if not m:
        return None
    try:
        return float(m.group())
    except ValueError:
        return None


def _pct(s: Optional[str]) -> Optional[float]:
    """Parse '31.58%' → 0.3158 (float). Returns None for missing."""
    if s is None:
        return None
    s = s.replace(",", "").strip()
    if s in ("", "-", "−"):
        return None
    m = _PCT_RE.search(s)
    if not m:
        return None
    try:
        return float(m.group()) / 100.0
    except ValueError:
        return None


def _two_lines(cell) -> tuple[str, str]:
    """Split a <td> with <br/> into two halves (sell-side / buy-side)."""
    if cell is None:
        return "", ""
    text = cell.get_text(separator="\n", strip=False).strip()
    parts = [p.strip() for p in text.split("\n") if p.strip()]
    if len(parts) >= 2:
        return parts[0], parts[1]
    if len(parts) == 1:
        return parts[0], ""
    return "", ""


def _parse_qty_bracket(s: str) -> tuple[Optional[float], Optional[float]]:
    """Parse '3 (44)' → (price=3, qty=44)."""
    if not s or s in ("-", "−"):
        return None, None
    m = re.match(r"\s*([-0-9,\.]+)\s*\(\s*([-0-9,\.]+)\s*\)", s)
    if not m:
        return _num(s), None
    return _num(m.group(1)), _num(m.group(2))


def _parse_contract_month_label(soup: BeautifulSoup) -> str:
    """Extract the OPTION contract month (YYMM) from the active tab.

    Page layout:
      <div id="futuresContractTab">
        <li class="active"><a>...>7月限月</a></li>  ← THIS is the option month
        <li>...</li>
      </div>
    The "(26年06月限)" line in the header refers to the underlying NK225
    futures (front month), NOT the option's expiration — do not use it.

    Year is inferred: if the parsed month >= current month, use current year,
    else assume next year (handles year boundary cases like Dec→Mar listings).
    """
    tab_div = soup.select_one("#futuresContractTab")
    if not tab_div:
        return ""
    active = tab_div.select_one("li.active")
    if not active:
        return ""
    txt = active.get_text(strip=True)
    m = re.search(r"(\d{1,2})月限月", txt)
    if not m:
        return ""
    month = int(m.group(1))
    today = date.today()
    year = today.year if month >= today.month else today.year + 1
    return f"{year % 100:02d}{month:02d}"


def _parse_update_time(soup: BeautifulSoup) -> Optional[datetime]:
    dd = soup.select_one(".update-time dd")
    if not dd:
        return None
    txt = dd.get_text(strip=True)
    # e.g. "2026/06/02 23:48"
    try:
        return datetime.strptime(txt, "%Y/%m/%d %H:%M")
    except ValueError:
        return None


def _iter_strike_rows(soup: BeautifulSoup):
    """Yield (row_num_tr, greek_tr_or_None) pairs."""
    body = soup.select_one("tbody.price-info-scroll")
    if body is None:
        return
    rows = body.find_all("tr", recursive=False)
    i = 0
    while i < len(rows):
        tr = rows[i]
        cls = tr.get("class", [])
        if "row-num" in cls:
            greek = None
            if i + 1 < len(rows) and "greek" in rows[i + 1].get("class", []):
                greek = rows[i + 1]
                i += 2
            else:
                i += 1
            yield tr, greek
        else:
            i += 1


def _extract_greeks(greek_tr) -> tuple[dict, dict]:
    """Return ({put delta/gamma/theta/vega}, {call ...})."""
    out_put, out_call = {}, {}
    if greek_tr is None:
        return out_put, out_call
    tables = greek_tr.find_all("table", class_="greek-value-table")
    keys = ["delta", "gamma", "theta", "vega"]
    for side, tbl in zip(("put", "call"), tables):
        rows = tbl.find_all("tr")
        if len(rows) < 2:
            continue
        cells = rows[1].find_all("td")
        if len(cells) < 4:
            continue
        d = {keys[i]: _num(cells[i].get_text(strip=True)) for i in range(4)}
        if side == "put":
            out_put = d
        else:
            out_call = d
    return out_put, out_call


def parse_page(html: str, *, fetch_time: datetime) -> list[dict]:
    """Parse one contract month page into a list of per-strike records.

    Each <tr.row-num> yields TWO records (one PUT, one CALL).
    """
    soup = BeautifulSoup(html, "html.parser")
    cm_label = _parse_contract_month_label(soup)
    update_time = _parse_update_time(soup)

    records: list[dict] = []
    for row_tr, greek_tr in _iter_strike_rows(soup):
        tds = row_tr.find_all("td", recursive=False)
        if len(tds) < 17:
            continue

        # Columns (per HTML inspection — CALL is LEFT, PUT is RIGHT):
        # 0  CALL 清算値          1  CALL 建玉残    2  CALL 取引高
        # 3  CALL 売気配IV/買気配IV (1行目=売,2行目=買)
        # 4  CALL 売気配(数量)/買気配(数量)
        # 5  CALL IV               6  CALL 前日比+%  7  CALL 現在値 / 時刻
        # 8  権利行使価格 (strike)
        # 9  PUT 現在値 / 時刻   10  PUT 前日比+%
        # 11 PUT IV              12 PUT 売気配(数量)/買気配(数量)
        # 13 PUT 売気配IV/買気配IV
        # 14 PUT 取引高         15 PUT 建玉残  16 PUT 清算値

        strike_txt = tds[8].get_text(separator=" ", strip=True)
        strike_m = re.search(r"\d{2,3},\d{3}|\d{4,6}", strike_txt)
        if not strike_m:
            continue
        strike = int(strike_m.group().replace(",", ""))

        call_greeks, put_greeks = _extract_greeks(greek_tr)

        # CALL side (LEFT, tds[0..7])
        c_settle = _num(tds[0].get_text(strip=True))
        c_oi = _num(tds[1].get_text(strip=True))
        c_volume = _num(tds[2].get_text(strip=True))
        c_ask_iv_s, c_bid_iv_s = _two_lines(tds[3])
        c_ask_q_s, c_bid_q_s = _two_lines(tds[4])
        c_iv = _pct(tds[5].get_text(strip=True))
        c_chg_s = tds[6].get_text(separator="\n", strip=True)
        c_chg_parts = [x.strip() for x in c_chg_s.split("\n") if x.strip()]
        c_chg = _num(c_chg_parts[0]) if c_chg_parts else None
        c_chg_pct = _pct(c_chg_parts[1]) if len(c_chg_parts) >= 2 else None
        c_last_s = tds[7].get_text(separator="\n", strip=True)
        c_last_parts = [x.strip() for x in c_last_s.split("\n") if x.strip()]
        c_last = _num(c_last_parts[0]) if c_last_parts else None
        c_ask_price, c_ask_qty = _parse_qty_bracket(c_ask_q_s)
        c_bid_price, c_bid_qty = _parse_qty_bracket(c_bid_q_s)

        records.append({
            "fetch_time": fetch_time,
            "source_update_time": update_time,
            "contract_month": cm_label,
            "strike": strike,
            "option_type": "CALL",
            "settle": c_settle,
            "oi": c_oi,
            "volume": c_volume,
            "ask_iv": _pct(c_ask_iv_s),
            "bid_iv": _pct(c_bid_iv_s),
            "iv": c_iv,
            "ask_price": c_ask_price,
            "ask_qty": c_ask_qty,
            "bid_price": c_bid_price,
            "bid_qty": c_bid_qty,
            "last": c_last,
            "change": c_chg,
            "change_pct": c_chg_pct,
            "delta": call_greeks.get("delta"),
            "gamma": call_greeks.get("gamma"),
            "theta": call_greeks.get("theta"),
            "vega": call_greeks.get("vega"),
        })

        # PUT side (RIGHT, tds[9..16])
        p_last_s = tds[9].get_text(separator="\n", strip=True)
        p_last_parts = [x.strip() for x in p_last_s.split("\n") if x.strip()]
        p_last = _num(p_last_parts[0]) if p_last_parts else None
        p_chg_s = tds[10].get_text(separator="\n", strip=True)
        p_chg_parts = [x.strip() for x in p_chg_s.split("\n") if x.strip()]
        p_chg = _num(p_chg_parts[0]) if p_chg_parts else None
        p_chg_pct = _pct(p_chg_parts[1]) if len(p_chg_parts) >= 2 else None
        p_iv = _pct(tds[11].get_text(strip=True))
        p_ask_q_s, p_bid_q_s = _two_lines(tds[12])
        p_ask_iv_s, p_bid_iv_s = _two_lines(tds[13])
        p_volume = _num(tds[14].get_text(strip=True))
        p_oi = _num(tds[15].get_text(strip=True))
        p_settle = _num(tds[16].get_text(strip=True))
        p_ask_price, p_ask_qty = _parse_qty_bracket(p_ask_q_s)
        p_bid_price, p_bid_qty = _parse_qty_bracket(p_bid_q_s)

        records.append({
            "fetch_time": fetch_time,
            "source_update_time": update_time,
            "contract_month": cm_label,
            "strike": strike,
            "option_type": "PUT",
            "settle": p_settle,
            "oi": p_oi,
            "volume": p_volume,
            "ask_iv": _pct(p_ask_iv_s),
            "bid_iv": _pct(p_bid_iv_s),
            "iv": p_iv,
            "ask_price": p_ask_price,
            "ask_qty": p_ask_qty,
            "bid_price": p_bid_price,
            "bid_qty": p_bid_qty,
            "last": p_last,
            "change": p_chg,
            "change_pct": p_chg_pct,
            "delta": put_greeks.get("delta"),
            "gamma": put_greeks.get("gamma"),
            "theta": put_greeks.get("theta"),
            "vega": put_greeks.get("vega"),
        })

    return records


# ---------------------------------------------------------------------------
# Driver
# ---------------------------------------------------------------------------

def run_once(max_index: int = _DEFAULT_MAX_INDEX, out_root: Path = _OUT_ROOT) -> Optional[Path]:
    """Single fetch over all available months. Returns output Parquet path."""
    fetch_time = datetime.now()
    all_records: list[dict] = []
    consecutive_misses = 0
    for idx in range(max_index):
        html = _fetch_html(idx)
        if html is None:
            consecutive_misses += 1
            # 2 misses in a row → assume no more months to probe
            if consecutive_misses >= 2 and idx > 0:
                break
            continue
        consecutive_misses = 0
        recs = parse_page(html, fetch_time=fetch_time)
        logger.info("idx=%d → %d records", idx, len(recs))
        all_records.extend(recs)
        time.sleep(1.0)  # gentle pacing

    if not all_records:
        logger.warning("No records fetched")
        return None

    df = pd.DataFrame(all_records)
    day_dir = out_root / fetch_time.strftime("%Y%m%d")
    day_dir.mkdir(parents=True, exist_ok=True)
    fname = day_dir / fetch_time.strftime("%Y%m%d_%H%M%S.parquet")
    df.to_parquet(fname, index=False)
    logger.info("Wrote %d records → %s", len(df), fname)
    return fname


def run_loop(interval_min: int, max_index: int, out_root: Path):
    while True:
        start = time.time()
        try:
            run_once(max_index=max_index, out_root=out_root)
        except Exception as e:
            logger.exception("run_once failed: %s", e)
        elapsed = time.time() - start
        sleep_for = max(60.0, interval_min * 60 - elapsed)
        logger.info("Sleeping %.0fs", sleep_for)
        time.sleep(sleep_for)


def main():
    p = argparse.ArgumentParser(description="QRI NK225 Option IV scraper")
    p.add_argument("--loop", action="store_true", help="Run indefinitely on interval")
    p.add_argument("--interval", type=int, default=15, help="Loop interval minutes (default 15)")
    p.add_argument("--max", type=int, default=_DEFAULT_MAX_INDEX,
                   help="Highest /jpx/nkopm/<i> index to probe (default 8)")
    p.add_argument("--out", type=Path, default=_OUT_ROOT, help="Output root dir")
    p.add_argument("-v", "--verbose", action="store_true")
    args = p.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )

    if args.loop:
        run_loop(args.interval, args.max, args.out)
    else:
        path = run_once(args.max, args.out)
        if path is None:
            sys.exit(1)


if __name__ == "__main__":
    main()
