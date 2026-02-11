"""Application-wide constants and configuration."""

from pathlib import Path

# --- Paths ---
PROJECT_ROOT = Path(__file__).parent
CACHE_DIR = PROJECT_ROOT / "cache"
CACHE_VOLUME_DIR = CACHE_DIR / "volume"
CACHE_OI_DIR = CACHE_DIR / "oi"
CACHE_INDEX_DIR = CACHE_DIR / "index"
CACHE_DAILY_OI_DIR = CACHE_DIR / "daily_oi"

# --- JPX API Base ---
JPX_BASE_URL = "https://www.jpx.co.jp"

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

# --- Daily OI Balance ---
DAILY_OI_URL_TEMPLATE = (
    "https://www.jpx.co.jp/markets/derivatives/"
    "trading-volume/tvdivq00000014nn-att/"
    "{yyyymmdd}open_interest.xlsx"
)

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
