"""Fetch today's daily OI balance from JPX and cache to R2.

Designed to run as a GitHub Actions job at ~20:15 JST daily.
Uses the existing fetcher/cache/R2 pipeline.
"""
import sys
import logging
from datetime import date, timezone, timedelta
from pathlib import Path

# Add project root to path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from data.cache import ensure_cache_dirs
from data.fetcher import download_daily_oi_excel

logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")
logger = logging.getLogger(__name__)


def main():
    # At 11:15 UTC it's 20:15 JST -- same calendar day
    jst = timezone(timedelta(hours=9))
    today = date.today()

    logger.info("Fetching daily OI for %s", today)
    ensure_cache_dirs()

    content = download_daily_oi_excel(today)
    if content is None:
        logger.warning("No daily OI file available for %s (holiday or not yet published)", today)
        sys.exit(0)  # Not an error

    logger.info("Successfully fetched and cached %d bytes for %s", len(content), today)


if __name__ == "__main__":
    main()
