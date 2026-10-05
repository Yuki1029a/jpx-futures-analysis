"""Application-wide constants and configuration."""

from pathlib import Path

# --- Paths ---
PROJECT_ROOT = Path(__file__).parent
CACHE_DIR = PROJECT_ROOT / "cache"
CACHE_VOLUME_DIR = CACHE_DIR / "volume"
CACHE_OI_DIR = CACHE_DIR / "oi"
CACHE_INDEX_DIR = CACHE_DIR / "index"
CACHE_DAILY_OI_DIR = CACHE_DIR / "daily_oi"
CACHE_MARKET_DATA_DIR = CACHE_DIR / "market_data"

# --- JPX API Base ---
JPX_BASE_URL = "https://www.jpx.co.jp"

# --- JPX「当日取引高等」ページの添付ファイル（建玉残高表・取引概況） ---
# 添付ディレクトリID（xxxx-att）はJPX側のCMS都合で予告なく変わる（2026-09-30に
# tvdivq00000014nn-att → t13vrt0000026aes-att に変わり3営業日分を取りこぼした）。
# 取得時はページHTMLから現在のIDを動的に発見し、過去IDは候補として残す
# （旧URLでキャッシュ済みのファイルを引き続き参照するため）。
TRADING_VOLUME_INDEX_URL = JPX_BASE_URL + "/markets/derivatives/trading-volume/index.html"
TRADING_VOLUME_ATT_BASE = JPX_BASE_URL + "/markets/derivatives/trading-volume/"
KNOWN_ATT_DIRS = ["t13vrt0000026aes-att", "tvdivq00000014nn-att"]  # 新しい順
# 先物・オプション取引概況（P/C別売買代金）。session: "whole_day" / "night"。
# JPXは最新営業日分しか掲載しないため日次収集でR2に履歴を蓄積する。
MARKET_DATA_FILENAME = "{yyyymmdd}_derivatives_market_data_{session}.xlsx"
# 2026-09-30以降、取引概況はページ内JSON経由で以下に掲載（過去分も約1か月残る）。
# 索引: /automation/markets/derivatives/trading-volume/json/derivatives_market_data_{mor,eve}.json
MARKET_DATA_FILES_BASE = JPX_BASE_URL + "/automation/markets/derivatives/trading-volume/files/"

# --- Daily Volume (売買高) ---
VOLUME_MONTHLY_LIST_URL = (
    f"{JPX_BASE_URL}/automation/markets/derivatives/"
    "participant-volume/json/participant-volume_monthlylist.json"
)
VOLUME_INDEX_URL_TEMPLATE = (
    JPX_BASE_URL + "/automation/markets/derivatives/"
    "participant-volume/json/participant_volume_{yyyymm}.json"
)

# --- Weekly Open Interest (建玉) ---
OI_YEAR_LIST_URL = (
    f"{JPX_BASE_URL}/automation/markets/derivatives/"
    "open-interest/json/open_interest_yearlist.json"
)
# OI year index URL is taken from the year list JSON (Jsonfile field)

# --- Excel Parsing: Daily Volume ---
VOLUME_DATA_START_ROW = 9
VOLUME_COLUMNS = {
    "product": 1,        # A
    "issue_code": 2,     # B
    "contract": 3,       # C
    "rank": 4,           # D
    "participant_id": 5, # E
    "name_jp": 6,        # F
    "name_en": 7,        # G
    "volume": 8,         # H
}

# --- Excel Parsing: Weekly Open Interest ---
OI_DATA_OFFSET = 2        # Data rows start 2 rows after section header
OI_ROWS_PER_SECTION = 15  # 15 participants per product per side

# Near month columns (left half)
# C-E = 売超参加者 (short/net sellers), F-H = 買超参加者 (long/net buyers)
OI_NEAR_COLUMNS = {
    "rank": 1,            # A
    "contract_month": 2,  # B (merged, only in first data row)
    "short_pid": 3,       # C  売超
    "short_name_jp": 4,   # D  売超
    "short_volume": 5,    # E  売超
    "long_pid": 6,        # F  買超
    "long_name_jp": 7,    # G  買超
    "long_volume": 8,     # H  買超
}

# Far month columns (right half)
# M-O = 売超参加者 (short/net sellers), P-R = 買超参加者 (long/net buyers)
OI_FAR_COLUMNS = {
    "rank": 11,           # K
    "contract_month": 12, # L (merged, only in first data row)
    "short_pid": 13,      # M  売超
    "short_name_jp": 14,  # N  売超
    "short_volume": 15,   # O  売超
    "long_pid": 16,       # P  買超
    "long_name_jp": 17,   # Q  買超
    "long_volume": 18,    # R  買超
}

# --- Daily OI Balance（建玉残高表。最新営業日分のみ掲載） ---
DAILY_OI_FILENAME = "{yyyymmdd}open_interest.xlsx"

# --- Target Products ---
TARGET_PRODUCTS = ["NK225F", "TOPIXF"]

# Volume file session keys to download (sum both for true daily total)
VOLUME_SESSION_KEYS = ["WholeDay", "Night"]

# Display names
PRODUCT_DISPLAY_NAMES = {
    "NK225F": "日経225先物",
    "TOPIXF": "TOPIX先物",
    "NK225MF": "日経225mini",
    "NK225OP": "日経225オプション",
}

# --- Options Configuration ---
OPTION_OI_SECTION_KEYWORDS = {
    "PUT": ["put", "プット", "PUT"],
    "CALL": ["call", "コール", "CALL"],
}

OPTION_STRIKE_DISPLAY_RANGE = {
    "ATM±5": 5,
    "ATM±10": 10,
    "ATM±20": 20,
    "全て": 9999,
}
