"""Fetch JSON indexes and Excel files from JPX website."""
from __future__ import annotations

import logging
import re
import time
import requests
from datetime import date
from pathlib import Path
import config

logger = logging.getLogger(__name__)
from data.cache import (
    ensure_cache_dirs, get_cached_bytes, save_to_cache,
    get_cached_json, save_json_to_cache,
)


def fetch_json(url: str, cache_hours: float = 1.0) -> dict:
    """Fetch a JSON endpoint with caching."""
    cached = get_cached_json(url, max_age_hours=cache_hours)
    if cached is not None:
        return cached
    response = requests.get(url, timeout=30)
    response.raise_for_status()
    data = response.json()
    save_json_to_cache(url, data)
    return data


def fetch_excel(url: str, subdir: Path, cache_hours: float = 168.0) -> bytes:
    """Fetch an Excel file with caching. Default cache: 7 days."""
    cached = get_cached_bytes(url, subdir, max_age_hours=cache_hours)
    if cached is not None:
        return cached
    response = requests.get(url, timeout=60)
    response.raise_for_status()
    save_to_cache(url, subdir, response.content)
    return response.content


def get_available_volume_months() -> list[str]:
    """Return list of available months (YYYYMM) for daily volume data."""
    data = fetch_json(config.VOLUME_MONTHLY_LIST_URL)
    return [entry["Month"] for entry in data["TableDatas"]]


def get_volume_index(yyyymm: str) -> list[dict]:
    """Return list of daily volume file entries for a given month.

    Each entry: {TradeDate, Night, NightJNet, WholeDay, WholeDayJNet}
    Returned in chronological order (oldest first).
    """
    url = config.VOLUME_INDEX_URL_TEMPLATE.replace("{yyyymm}", yyyymm)
    data = fetch_json(url)
    return list(reversed(data["TableDatas"]))


def get_available_oi_years() -> list[dict]:
    """Return [{Year, Jsonfile}, ...]."""
    data = fetch_json(config.OI_YEAR_LIST_URL)
    return data["TableDatas"]


def get_oi_index(year: str) -> list[dict]:
    """Return list of weekly OI file entries for a given year.

    Each entry: {TradeDate, IndexFutures, IndexOptions, SecuritiesOptions}
    Returned in chronological order.
    """
    years = get_available_oi_years()
    json_path = None
    for y in years:
        if y["Year"] == year:
            json_path = y["Jsonfile"]
            break
    if json_path is None:
        raise ValueError(f"No OI data available for year {year}")
    url = config.JPX_BASE_URL + json_path
    data = fetch_json(url)
    return list(reversed(data["TableDatas"]))


def download_volume_excel(file_path: str) -> bytes:
    """Download a daily volume Excel file given its relative path."""
    url = config.JPX_BASE_URL + file_path
    return fetch_excel(url, config.CACHE_VOLUME_DIR)


def download_oi_excel(file_path: str) -> bytes:
    """Download a weekly OI Excel file given its relative path."""
    url = config.JPX_BASE_URL + file_path
    return fetch_excel(url, config.CACHE_OI_DIR)


_att_dir_cache: dict[str, tuple[float, str | None]] = {}
_UA = {"User-Agent": "Mozilla/5.0"}


def discover_att_dir(filename_pattern: str, cache_hours: float = 1.0) -> str | None:
    """「当日取引高等」ページHTMLから filename_pattern（正規表現）に合う添付の
    ディレクトリID（xxxx-att）を発見する。結果はプロセス内で cache_hours 保持。

    JPXの添付IDは予告なく変わる（2026-09-30 実例）ためテンプレート固定にしない。
    """
    now = time.time()
    hit = _att_dir_cache.get(filename_pattern)
    if hit and now - hit[0] < cache_hours * 3600:
        return hit[1]
    att = None
    try:
        resp = requests.get(config.TRADING_VOLUME_INDEX_URL, timeout=30, headers=_UA)
        resp.raise_for_status()
        m = re.search(r'href="/markets/derivatives/trading-volume/([^"/]+-att)/'
                      + filename_pattern, resp.text)
        att = m.group(1) if m else None
    except Exception:
        logger.warning("discover_att_dir failed for %s", filename_pattern, exc_info=True)
    _att_dir_cache[filename_pattern] = (now, att)
    return att


def _canonical_url(filename: str) -> str:
    """添付ID非依存のキャッシュキー（実在URLではない。ファイル名が同じなら同じ鍵）。"""
    return config.TRADING_VOLUME_ATT_BASE + "_canonical_/" + filename


def _candidate_urls(filename: str, discover_pattern: str) -> list[str]:
    dirs: list[str] = []
    found = discover_att_dir(discover_pattern)
    if found:
        dirs.append(found)
    dirs += [a for a in config.KNOWN_ATT_DIRS if a not in dirs]
    return [config.TRADING_VOLUME_ATT_BASE + a + "/" + filename for a in dirs]


def _cached_any(filename: str, subdir: Path, cache_hours: float) -> bytes | None:
    """キャッシュ（L1/R2）のみ参照。JPXへは行かない。

    cache.py のキーはURL末尾のファイル名なので、添付IDが違っても同じ鍵になる。
    正規キーと既知IDのURLは同一ファイルを指すが、将来キー方式が変わっても
    拾えるよう全候補を順に見る。
    """
    urls = [_canonical_url(filename)] + [
        config.TRADING_VOLUME_ATT_BASE + a + "/" + filename for a in config.KNOWN_ATT_DIRS]
    seen: set[str] = set()
    for url in urls:
        if url in seen:
            continue
        seen.add(url)
        cached = get_cached_bytes(url, subdir, max_age_hours=cache_hours)
        if cached is not None:
            return cached
    return None


def _fetch_attachment(filename: str, discover_pattern: str, subdir: Path,
                      cache_hours: float) -> bytes | None:
    """キャッシュ（全候補）→ライブ取得（発見ID→既知IDを順に試行、404は次へ）。"""
    cached = _cached_any(filename, subdir, cache_hours)
    if cached is not None:
        return cached
    for url in _candidate_urls(filename, discover_pattern):
        try:
            resp = requests.get(url, timeout=60, headers=_UA)
            if resp.status_code == 404:
                logger.debug("404: %s", url)
                continue
            resp.raise_for_status()
            save_to_cache(url, subdir, resp.content)
            return resp.content
        except Exception:
            logger.warning("fetch failed: %s", url, exc_info=True)
    return None


def download_market_data_excel(trade_date: date, session: str = "whole_day") -> bytes | None:
    """取引概況Excel（P/C別売買代金入り）。最新営業日以外はJPXに無く未取得なら None。

    キャッシュ（L1/R2）にあれば過去日でもそこから返る。確定値で不変のため実質無期限。
    """
    fn = config.MARKET_DATA_FILENAME.format(
        yyyymmdd=trade_date.strftime("%Y%m%d"), session=session)
    return _fetch_attachment(fn, r"\d{8}_derivatives_market_data_" + session + r"\.xlsx",
                             config.CACHE_MARKET_DATA_DIR, cache_hours=24 * 3650)


def cached_market_data_excel(trade_date: date, session: str = "whole_day") -> bytes | None:
    """取引概況Excelをキャッシュのみから返す（過去日表示用。JPXへは行かない）。"""
    fn = config.MARKET_DATA_FILENAME.format(
        yyyymmdd=trade_date.strftime("%Y%m%d"), session=session)
    return _cached_any(fn, config.CACHE_MARKET_DATA_DIR, cache_hours=24 * 3650)


def download_daily_oi_excel(trade_date: date) -> bytes | None:
    """建玉残高表Excel。キャッシュ優先、無ければ現在の添付IDと既知IDで取得。未存在は None。"""
    fn = config.DAILY_OI_FILENAME.format(yyyymmdd=trade_date.strftime("%Y%m%d"))
    return _fetch_attachment(fn, r"\d{8}open_interest\.xlsx",
                             config.CACHE_DAILY_OI_DIR, cache_hours=168.0)
