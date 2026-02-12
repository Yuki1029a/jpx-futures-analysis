"""Fetch today's daily OI balance AND participant volume from JPX, cache to R2.

Designed to run as a GitHub Actions job at ~20:15 JST daily (weekdays).
Uses the existing fetcher/cache/R2 pipeline.
"""
import sys
import logging
from datetime import date, timezone, timedelta
from pathlib import Path

# Add project root to path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from data.cache import ensure_cache_dirs
from data.fetcher import (
    download_daily_oi_excel,
    get_volume_index,
    download_volume_excel,
)
import config

logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")
logger = logging.getLogger(__name__)


def main():
    jst = timezone(timedelta(hours=9))
    today = date.today()
    yyyymm = today.strftime("%Y%m")
    date_str = today.strftime("%Y%m%d")

    logger.info("Fetching data for %s", today)
    ensure_cache_dirs()

    # --- 1. Daily OI balance (open_interest.xlsx) ---
    content = download_daily_oi_excel(today)
    if content is None:
        logger.warning("No daily OI file for %s (holiday or not yet published)", today)
    else:
        logger.info("Daily OI: %d bytes for %s", len(content), today)

    # --- 2. Participant volume (手口 Excel) ---
    try:
        entries = get_volume_index(yyyymm)
    except Exception:
        logger.warning("Failed to get volume index for %s", yyyymm, exc_info=True)
        entries = []

    today_entries = [e for e in entries if e["TradeDate"] == date_str]
    if not today_entries:
        logger.warning("No volume entries for %s in index", today)
    else:
        fetched = 0
        for entry in today_entries:
            for key in config.VOLUME_SESSION_KEYS:
                file_path = entry.get(key)
                if file_path:
                    try:
                        data = download_volume_excel(file_path)
                        logger.info("Volume %s: %d bytes", key, len(data))
                        fetched += 1
                    except Exception:
                        logger.warning("Failed to fetch %s: %s", key, file_path, exc_info=True)
                # Also fetch JNet (Japanese net) versions if available
                jnet_key = key + "JNet"
                file_path_jnet = entry.get(jnet_key)
                if file_path_jnet:
                    try:
                        data = download_volume_excel(file_path_jnet)
                        logger.info("Volume %s: %d bytes", jnet_key, len(data))
                        fetched += 1
                    except Exception:
                        pass
        logger.info("Fetched %d volume files for %s", fetched, today)

    logger.info("Done.")


if __name__ == "__main__":
    main()
