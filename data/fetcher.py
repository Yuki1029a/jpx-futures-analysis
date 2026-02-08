"""Fetch JSON indexes and Excel files from JPX website."""
from __future__ import annotations

import requests
from pathlib import Path
import config
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
