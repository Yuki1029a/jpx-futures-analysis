"""Aggregate data into weekly views combining OI and daily volumes.

Night session date handling:
    JPX publishes Night session data under the NEXT business day's trade date.
    e.g., Friday 1/30's night session appears in the 2/2 (Monday) file.
    We shift Night data back to the actual market date (previous business day).
"""

from datetime import date, datetime
from typing import Optional
from models import (
    ParticipantVolume, ParticipantOI,
    WeekDefinition, WeeklyParticipantRow,
)
from data import fetcher
from data.parser_volume import parse_volume_excel, merge_volume_records
from data.parser_oi import parse_oi_excel
import config


# --- Trading date index (cached) ---

_trading_dates_cache: list[date] | None = None
_next_td_map: dict[date, date] | None = None  # market_date -> next trading date
_volume_parse_cache: dict[str, list[ParticipantVolume]] = {}  # file_path -> parsed records
_oi_parse_cache: dict[str, list[ParticipantOI]] = {}  # file_path -> parsed records


def _ensure_trading_date_index():
    """Build and cache the sorted list of all trading dates and
    a mapping from each trading date to the next one."""
    global _trading_dates_cache, _next_td_map
    if _trading_dates_cache is not None:
        return
    _trading_dates_cache = get_all_trading_dates()
    _next_td_map = {}
    for i in range(len(_trading_dates_cache) - 1):
        _next_td_map[_trading_dates_cache[i]] = _trading_dates_cache[i + 1]


def _get_next_trading_date(d: date) -> date | None:
    """Return the next trading date after d."""
    _ensure_trading_date_index()
    return _next_td_map.get(d)


def _get_prev_trading_date(d: date) -> date | None:
    """Return the previous trading date before d."""
    _ensure_trading_date_index()
    idx = None
    for i, td in enumerate(_trading_dates_cache):
        if td == d:
            idx = i
            break
    if idx is not None and idx > 0:
        return _trading_dates_cache[idx - 1]
    return None


# --- Public API ---

def get_all_oi_dates() -> list[date]:
    """Collect all OI report dates from available years, sorted ascending."""
    years_info = fetcher.get_available_oi_years()
    all_dates = []
    for y_info in years_info:
        year = y_info["Year"]
        try:
            entries = fetcher.get_oi_index(year)
        except Exception:
            continue
        for entry in entries:
            d = datetime.strptime(entry["TradeDate"], "%Y%m%d").date()
            all_dates.append(d)
    all_dates.sort()
    return all_dates


def get_all_trading_dates() -> list[date]:
    """Collect all trading dates from available months, sorted ascending."""
    months = fetcher.get_available_volume_months()
    all_dates = []
    for m in months:
        try:
            entries = fetcher.get_volume_index(m)
        except Exception:
            continue
        for entry in entries:
            d = datetime.strptime(entry["TradeDate"], "%Y%m%d").date()
            all_dates.append(d)
    all_dates.sort()
    return all_dates


def build_available_weeks(max_weeks: int = 26) -> list[WeekDefinition]:
    """Build a list of available weeks based on OI publication dates.

    A 'week' = period between two consecutive OI dates.
    If trading data exists after the latest OI date, a "current week"
    entry is prepended with end_oi_date=None (OI not yet published).
    Returns weeks in reverse chronological order (most recent first).
    """
    oi_dates = get_all_oi_dates()
    trading_dates = get_all_trading_dates()

    weeks = []

    # Check for in-progress week (trading data after latest OI)
    if oi_dates and trading_dates:
        latest_oi = oi_dates[-1]
        future_trades = [d for d in trading_dates if d > latest_oi]
        if future_trades:
            future_trades.sort()
            weeks.append(WeekDefinition(
                start_oi_date=latest_oi,
                end_oi_date=None,  # OI not yet published
                trading_days=future_trades,
                label=f"{latest_oi.strftime('%m/%d')} - (進行中)",
            ))

    for i in range(len(oi_dates) - 1, 0, -1):
        if len(weeks) >= max_weeks:
            break
        end_date = oi_dates[i]
        start_date = oi_dates[i - 1]

        # Trading days between start (exclusive) and end (inclusive)
        t_days = [d for d in trading_dates if start_date < d <= end_date]
        t_days.sort()

        weeks.append(WeekDefinition(
            start_oi_date=start_date,
            end_oi_date=end_date,
            trading_days=t_days,
            label=f"{start_date.strftime('%m/%d')} - {end_date.strftime('%m/%d')}",
        ))

    return weeks


def get_available_contract_months(
    week: WeekDefinition,
    product: str,
) -> list[str]:
    """Return available contract months (YYMM) for a given week and product."""
    oi_records = []
    if week.end_oi_date:
        oi_records = _load_oi_for_date(week.end_oi_date, product)
    if not oi_records:
        oi_records = _load_oi_for_date(week.start_oi_date, product)

    months = sorted(set(r.contract_month for r in oi_records))
    return months if months else [""]


"""Session filter keys."""
# Day-type sessions: file trade_date == actual market date
SESSION_DAY_KEYS = ("WholeDay", "WholeDayJNet")
# Night-type sessions: file trade_date == NEXT business day (actual = prev day)
SESSION_NIGHT_KEYS = ("Night", "NightJNet")

SESSION_ALL = "ALL"
SESSION_AUCTION_DAY = ("WholeDay",)
SESSION_AUCTION_NIGHT = ("Night",)
SESSION_JNET_DAY = ("WholeDayJNet",)
SESSION_JNET_NIGHT = ("NightJNet",)

# Map display label -> session mode
SESSION_MODES = {
    "全セッション合計": SESSION_ALL,
    "立会内(日中)": SESSION_AUCTION_DAY,
    "立会内(夜間)": SESSION_AUCTION_NIGHT,
    "立会外(日中)": SESSION_JNET_DAY,
    "立会外(夜間)": SESSION_JNET_NIGHT,
}


def load_weekly_data(
    week: WeekDefinition,
    product: str,
    contract_month: str,
    session_keys=SESSION_ALL,
    include_oi: bool = True,
) -> list[WeeklyParticipantRow]:
    """Load and aggregate all data for a specific week/product/contract.

    Night session data is shifted to the previous business day to match
    actual market timing.
    """
    # 1. Load OI
    start_oi: dict[str, ParticipantOI] = {}
    end_oi: dict[str, ParticipantOI] = {}
    if include_oi:
        start_oi_all = _load_oi_for_date(week.start_oi_date, product)
        start_oi = {r.participant_id: r for r in start_oi_all
                    if r.contract_month == contract_month}
        if week.end_oi_date:
            end_oi_all = _load_oi_for_date(week.end_oi_date, product)
            end_oi = {r.participant_id: r for r in end_oi_all
                      if r.contract_month == contract_month}

    # 2. Load daily volumes with proper Night session shifting
    daily_volumes: dict[date, dict[str, ParticipantVolume]] = {}
    for td in week.trading_days:
        records = _load_volume_for_market_date(
            td, product, contract_month, session_keys
        )
        daily_volumes[td] = {r.participant_id: r for r in records}

    # 3. Collect all participant IDs
    all_pids = set()
    all_pids.update(start_oi.keys())
    all_pids.update(end_oi.keys())
    for day_data in daily_volumes.values():
        all_pids.update(day_data.keys())

    # 4. Build name lookup
    name_lookup = _build_name_lookup(daily_volumes, start_oi, end_oi)

    # 5. Assemble rows
    rows = []
    for pid in all_pids:
        s_oi = start_oi.get(pid)
        e_oi = end_oi.get(pid)

        s_long = s_oi.long_volume if s_oi and s_oi.long_volume else 0.0
        s_short = s_oi.short_volume if s_oi and s_oi.short_volume else 0.0
        s_net = s_long - s_short

        e_long = e_oi.long_volume if e_oi and e_oi.long_volume else 0.0
        e_short = e_oi.short_volume if e_oi and e_oi.short_volume else 0.0
        e_net = e_long - e_short

        dvols = {}
        for td in week.trading_days:
            day_data = daily_volumes.get(td, {})
            pv = day_data.get(pid)
            if pv:
                dvols[td] = pv.volume

        oi_net_change = e_net - s_net

        has_oi = s_oi is not None or e_oi is not None
        direction = None
        if has_oi:
            if oi_net_change > 0:
                direction = "BUY"
            elif oi_net_change < 0:
                direction = "SELL"
            else:
                direction = "NEUTRAL"

        rows.append(WeeklyParticipantRow(
            participant_id=pid,
            participant_name=name_lookup.get(pid, pid),
            start_oi_long=s_long if s_oi else None,
            start_oi_short=s_short if s_oi else None,
            start_oi_net=s_net if s_oi else None,
            daily_volumes=dvols,
            end_oi_long=e_long if e_oi else None,
            end_oi_short=e_short if e_oi else None,
            end_oi_net=e_net if e_oi else None,
            oi_net_change=oi_net_change if has_oi else None,
            inferred_direction=direction,
        ))

    # Sort by total weekly volume descending
    rows.sort(key=lambda r: sum(r.daily_volumes.values()), reverse=True)
    return rows


# --- Private helpers ---

def _load_oi_for_date(d: date, product: str) -> list[ParticipantOI]:
    """Load OI data for a specific report date and product."""
    year = str(d.year)
    date_str = d.strftime("%Y%m%d")

    try:
        entries = fetcher.get_oi_index(year)
    except Exception:
        return []

    for entry in entries:
        if entry["TradeDate"] == date_str:
            file_path = entry.get("IndexFutures")
            if not file_path:
                return []
            try:
                if file_path in _oi_parse_cache:
                    return _oi_parse_cache[file_path]
                content = fetcher.download_oi_excel(file_path)
                records = parse_oi_excel(content, [product])
                _oi_parse_cache[file_path] = records
                return records
            except Exception:
                return []

    return []


def _load_raw_session(
    jpx_trade_date: date,
    product: str,
    contract_month: str,
    file_keys: tuple[str, ...],
) -> list[ParticipantVolume]:
    """Load specific session files for a given JPX trade date (raw, no shifting)."""
    yyyymm = jpx_trade_date.strftime("%Y%m")
    date_str = jpx_trade_date.strftime("%Y%m%d")

    try:
        entries = fetcher.get_volume_index(yyyymm)
    except Exception:
        return []

    for entry in entries:
        if entry["TradeDate"] == date_str:
            all_records = []
            for key in file_keys:
                path = entry.get(key)
                if path:
                    try:
                        if path in _volume_parse_cache:
                            records = _volume_parse_cache[path]
                        else:
                            content = fetcher.download_volume_excel(path)
                            records = parse_volume_excel(content, [product])
                            _volume_parse_cache[path] = records
                        all_records.append(records)
                    except Exception:
                        pass
            merged = merge_volume_records(*all_records)
            return [r for r in merged if r.contract_month == contract_month]

    return []


def _redate_records(records: list[ParticipantVolume], new_date: date) -> list[ParticipantVolume]:
    """Change trade_date on all records so merge keys align."""
    for r in records:
        r.trade_date = new_date
    return records


def _load_volume_for_market_date(
    market_date: date,
    product: str,
    contract_month: str,
    session_mode,
) -> list[ParticipantVolume]:
    """Load volume data for an actual market date with Night session shifting.

    For a given market_date:
      - Day-type files (WholeDay, WholeDayJNet):
          Found in JPX trade_date == market_date
      - Night-type files (Night, NightJNet):
          Found in JPX trade_date == NEXT business day after market_date
          (because JPX labels Night under the next day's trade date)
          Night records are re-dated to market_date before merging.

    session_mode can be:
      - SESSION_ALL: combine all 4 files with proper shifting
      - A tuple of specific keys like ("WholeDay",) or ("Night",)
    """
    _ensure_trading_date_index()

    if session_mode == SESSION_ALL:
        # Day files from market_date
        day_records = _load_raw_session(
            market_date, product, contract_month, SESSION_DAY_KEYS
        )
        # Night files from next trading date, re-dated to market_date
        night_records = []
        next_td = _get_next_trading_date(market_date)
        if next_td:
            night_records = _load_raw_session(
                next_td, product, contract_month, SESSION_NIGHT_KEYS
            )
            _redate_records(night_records, market_date)
        return merge_volume_records(day_records, night_records)

    # Single session mode
    requested_keys = session_mode

    # Determine if the requested keys are Night-type
    is_night = all(k in SESSION_NIGHT_KEYS for k in requested_keys)

    if is_night:
        # Night data for market_date lives in next trading day's file
        next_td = _get_next_trading_date(market_date)
        if not next_td:
            return []
        records = _load_raw_session(next_td, product, contract_month, requested_keys)
        _redate_records(records, market_date)
        return records
    else:
        # Day data is straightforward
        return _load_raw_session(market_date, product, contract_month, requested_keys)


def compute_20d_stats(
    week: WeekDefinition,
    product: str,
    contract_month: str,
    session_keys=SESSION_ALL,
) -> dict[str, tuple[float, float]]:
    """Compute 20-business-day average and max daily volume per participant.

    Looks back 20 trading dates from the start of the given week.
    Returns: {participant_id: (avg, max)}
    """
    _ensure_trading_date_index()

    # Find the 20 trading dates ending at the day before the week starts
    # Use all trading dates up to and including the last trading day before week.trading_days[0]
    if not week.trading_days:
        return {}

    week_start = week.trading_days[0]
    lookback_dates = [d for d in _trading_dates_cache if d < week_start]
    lookback_dates = lookback_dates[-20:]  # last 20

    if not lookback_dates:
        return {}

    # Load daily volumes for each lookback date
    # pid -> list of daily volumes
    pid_daily: dict[str, list[float]] = {}

    for td in lookback_dates:
        records = _load_volume_for_market_date(
            td, product, contract_month, session_keys
        )
        for r in records:
            pid_daily.setdefault(r.participant_id, []).append(r.volume)

    # Compute stats
    result = {}
    for pid, volumes in pid_daily.items():
        avg = sum(volumes) / len(volumes)
        mx = max(volumes)
        result[pid] = (avg, mx)

    return result


def _build_name_lookup(
    daily_volumes: dict,
    start_oi: dict[str, ParticipantOI],
    end_oi: dict[str, ParticipantOI],
) -> dict[str, str]:
    """Build participant_id -> display_name lookup.

    Priority: English name from daily volume > Japanese name from OI.
    """
    lookup: dict[str, str] = {}

    # From OI (Japanese names, lower priority)
    for pid, rec in {**start_oi, **end_oi}.items():
        if rec.participant_name_jp:
            lookup[pid] = rec.participant_name_jp

    # From daily volume (English names, higher priority)
    for day_data in daily_volumes.values():
        for pid, pv in day_data.items():
            if pv.participant_name_en:
                lookup[pid] = pv.participant_name_en

    return lookup
