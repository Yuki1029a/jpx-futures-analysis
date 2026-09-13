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
    download_market_data_excel,
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
    # 公表時刻は日によってずれる（定時実行時に未公表のことがある）ため、
    # 当日だけでなく直近3日分を毎回試行し、取りこぼしを次回実行で回収する。
    # 取得済みの日はローカル/R2キャッシュが効くので再ダウンロードは発生しない。
    for back in range(3, -1, -1):
        d = today - timedelta(days=back)
        if d.weekday() >= 5:
            continue
        content = download_daily_oi_excel(d)
        if content is None:
            logger.warning("No daily OI file for %s (holiday or not yet published)", d)
        else:
            logger.info("Daily OI: %d bytes for %s", len(content), d)

    # --- 2. Participant volume (手口 Excel) ---
    try:
        entries = get_volume_index(yyyymm)
    except Exception:
        logger.warning("Failed to get volume index for %s", yyyymm, exc_info=True)
        entries = []

    # 当日だけでなく直近3営業日分のエントリを毎回試行する。
    # 夜間ファイル(Night/NightJNet)は「翌取引日」のTradeDateで公表されるため、
    # 当日限定だと公表タイミング次第で取りこぼす。取得済み分はR2キャッシュが
    # 効くので再ダウンロードは発生しない。月初は前月indexも遡る。
    recent_entries = [e for e in entries if e["TradeDate"] <= date_str][-3:]
    if today.day <= 4:
        prev_ym = (today.replace(day=1) - timedelta(days=1)).strftime("%Y%m")
        try:
            prev_entries = get_volume_index(prev_ym)
            recent_entries = (prev_entries + recent_entries)[-3:]
        except Exception:
            logger.warning("Failed to get volume index for %s", prev_ym, exc_info=True)
    today_entries = recent_entries
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

    # --- 3. 取引概況 (P/C売買代金) ---
    # JPXは最新営業日分しか掲載しないため、毎日R2へ保全して履歴を蓄積する。
    # 直近3日分を試行（過去日はキャッシュ済みなら即スキップ、未収集なら404でNone）。
    for back in range(2, -1, -1):
        d = today - timedelta(days=back)
        if d.weekday() >= 5:
            continue
        for sess in ("whole_day", "night"):
            content = download_market_data_excel(d, sess)
            if content is not None:
                logger.info("Market data %s %s: %d bytes", d, sess, len(content))

    logger.info("Done.")


if __name__ == "__main__":
    main()
