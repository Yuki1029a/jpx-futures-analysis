"""Aggregate data into weekly views combining OI and daily volumes.

Night session date handling:
    JPX publishes Night session data under the NEXT business day's trade date.
    e.g., Friday 1/30's night session appears in the 2/2 (Monday) file.
    We shift Night data back to the actual market date (previous business day).
"""
from __future__ import annotations

from datetime import date, datetime, timedelta
from typing import Optional
from models import (
    ParticipantVolume, ParticipantOI,
    WeekDefinition, WeeklyParticipantRow,
    OptionParticipantOI, OptionParticipantVolume, OptionStrikeRow,
    DailyOIBalance, DailyFuturesOI,
)
from data import fetcher
from data.parser_volume import (
    parse_volume_excel, merge_volume_records,
    parse_option_volume_excel, merge_option_volume_records,
)
from data.parser_oi import parse_oi_excel
from data.parser_option_oi import parse_option_oi_excel
from data.parser_daily_oi import parse_daily_oi_excel, parse_daily_futures_oi_excel
import config


# --- Trading date index (cached) ---

_trading_dates_cache: list[date] | None = None
_next_td_map: dict[date, date] | None = None  # market_date -> next trading date
_volume_parse_cache: dict[str, list[ParticipantVolume]] = {}  # file_path -> parsed records
_oi_parse_cache: dict[str, list[ParticipantOI]] = {}  # file_path -> parsed records
_option_volume_parse_cache: dict[str, list[OptionParticipantVolume]] = {}
_option_oi_parse_cache: dict[str, list[OptionParticipantOI]] = {}


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
            # Extend to today: add weekdays between last known trade and today
            # Only add dates where daily OI data actually exists (skip holidays)
            today = date.today()
            last_known = future_trades[-1]
            d = last_known + timedelta(days=1)
            while d <= today:
                if d.weekday() < 5 and d not in future_trades:  # Mon-Fri
                    # Probe: check if daily OI file exists for this date
                    content = fetcher.download_daily_oi_excel(d)
                    if content is not None:
                        future_trades.append(d)
                d += timedelta(days=1)
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

        # OI net change and direction require BOTH start and end OI.
        # For in-progress weeks (end_oi_date=None), end_oi is empty,
        # so these fields must be None — the data hasn't been published yet.
        has_both_oi = s_oi is not None and e_oi is not None
        oi_net_change = (e_net - s_net) if has_both_oi else None

        direction = None
        if has_both_oi:
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
            oi_net_change=oi_net_change,
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
                    records = _oi_parse_cache[file_path]
                else:
                    content = fetcher.download_oi_excel(file_path)
                    records = parse_oi_excel(content, None)
                    _oi_parse_cache[file_path] = records
                return [r for r in records if r.product == product]
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
                            records = parse_volume_excel(content, None)
                            _volume_parse_cache[path] = records
                        # Filter by product after cache lookup
                        filtered = [r for r in records if r.product == product]
                        all_records.append(filtered)
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

    Uses the most recent 20 trading dates up to and including the last
    trading day of the given week (i.e. the current week's data IS included).
    Returns: {participant_id: (avg, max)}
    """
    _ensure_trading_date_index()

    if not week.trading_days:
        return {}

    week_end = week.trading_days[-1]
    # All trading dates up to and including the last day of this week
    candidates = [d for d in _trading_dates_cache if d <= week_end]
    lookback_dates = candidates[-20:]  # last 20

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

    # Compute stats (avg over full 20-day window, not just days with activity)
    n_days = len(lookback_dates)
    result = {}
    for pid, volumes in pid_daily.items():
        avg = sum(volumes) / n_days
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


# =====================================================================
# Option aggregation
# =====================================================================

def get_available_option_contract_months(
    week: WeekDefinition,
) -> list[str]:
    """Return available option contract months (YYMM) for a given week.

    Checks OI data first, falls back to volume data.
    """
    # Try OI files
    oi_months: set[str] = set()
    for d in [week.end_oi_date, week.start_oi_date]:
        if d is None:
            continue
        records = _load_option_oi_raw(d)
        for r in records:
            if r.contract_month:
                oi_months.add(r.contract_month)
        if oi_months:
            break

    # Also check daily OI balance files for additional contract months
    if week.trading_days:
        daily_records = _load_daily_oi_for_date(week.trading_days[-1])
        for r in daily_records:
            if r.contract_month:
                oi_months.add(r.contract_month)

    if oi_months:
        return sorted(oi_months)

    # Fallback: check volume data from the first trading day
    if week.trading_days:
        vol_records = _load_option_volume_for_market_date(
            week.trading_days[0], SESSION_ALL
        )
        vol_months = set(r.contract_month for r in vol_records if r.contract_month)
        if vol_months:
            return sorted(vol_months)

    return []


def get_option_participants(
    week: WeekDefinition,
    contract_month: str,
    session_keys=SESSION_ALL,
) -> list[tuple[str, str]]:
    """Return list of (participant_id, display_name) for option data.

    Combines participants from OI and volume data.
    """
    pid_names: dict[str, str] = {}

    # From OI
    for d in [week.end_oi_date, week.start_oi_date]:
        if d is None:
            continue
        records = _load_option_oi_raw(d)
        for r in records:
            if r.contract_month == contract_month and r.participant_id:
                if r.participant_id not in pid_names:
                    pid_names[r.participant_id] = r.participant_name_jp or r.participant_id

    # From volume (first few trading days)
    for td in week.trading_days[:3]:
        vol_records = _load_option_volume_for_market_date(td, session_keys)
        for r in vol_records:
            if r.contract_month == contract_month and r.participant_id:
                name = r.participant_name_en or r.participant_name_jp or r.participant_id
                pid_names[r.participant_id] = name

    return sorted(pid_names.items(), key=lambda x: x[1])


def load_option_weekly_data(
    week: WeekDefinition,
    contract_month: str = "",
    session_keys=SESSION_ALL,
    participant_ids: list[str] | None = None,
) -> list[OptionStrikeRow]:
    """Load option data for a week: OI + daily volumes, aggregated by strike.

    Args:
        week: Week definition.
        contract_month: YYMM filter (empty = all).
        session_keys: Session filter.
        participant_ids: If provided, only include these participants.

    Returns OptionStrikeRow per strike price with filtered participant totals.
    """
    # 1. Load OI
    start_oi = _load_option_oi_for_date(
        week.start_oi_date, contract_month, participant_ids
    )
    end_oi = {}
    if week.end_oi_date:
        end_oi = _load_option_oi_for_date(
            week.end_oi_date, contract_month, participant_ids
        )

    # 2. Load daily option volumes
    daily_vols: dict[date, list[OptionParticipantVolume]] = {}
    for td in week.trading_days:
        records = _load_option_volume_for_market_date(td, session_keys)
        # Filter by contract_month
        if contract_month:
            records = [r for r in records if r.contract_month == contract_month]
        # Filter by participant
        if participant_ids is not None:
            pid_set = set(participant_ids)
            records = [r for r in records if r.participant_id in pid_set]
        daily_vols[td] = records

    # 2.5 Load daily OI balance (aggregate, not per-participant)
    daily_oi: dict[date, list[DailyOIBalance]] = {}
    for td in week.trading_days:
        oi_records = _load_daily_oi_for_date(td, contract_month)
        daily_oi[td] = oi_records

    # 3. Aggregate by strike
    return _aggregate_by_strike(start_oi, end_oi, daily_vols, week, daily_oi)


_daily_oi_parse_cache: dict[str, list[DailyOIBalance]] = {}


def _load_daily_oi_for_date(
    trade_date: date,
    contract_month: str = "",
) -> list[DailyOIBalance]:
    """Load daily OI balance records for a specific date.

    Aggregate per-strike data (not per-participant).
    No night session shifting needed (single daily file).
    """
    date_str = trade_date.strftime("%Y%m%d")
    cache_key = f"daily_oi_{date_str}"

    if cache_key in _daily_oi_parse_cache:
        records = _daily_oi_parse_cache[cache_key]
    else:
        content = fetcher.download_daily_oi_excel(trade_date)
        if content is None:
            _daily_oi_parse_cache[cache_key] = []
            return []
        records = parse_daily_oi_excel(content)
        _daily_oi_parse_cache[cache_key] = records

    if contract_month:
        records = [r for r in records if r.contract_month == contract_month]
    return records


_daily_futures_oi_cache: dict[str, list[DailyFuturesOI]] = {}


def load_daily_futures_oi(
    week: WeekDefinition,
    product: str,
    contract_month: str,
) -> dict[date, DailyFuturesOI]:
    """Load daily futures OI balance for each trading day in the week.

    Returns {date: DailyFuturesOI} for the matching product and contract month.
    Also derives previous day's OI from next day's previous_oi field
    (e.g. Tuesday's previous_oi = Monday's current_oi).
    """
    result: dict[date, DailyFuturesOI] = {}
    for td in week.trading_days:
        date_str = td.strftime("%Y%m%d")
        cache_key = f"daily_futures_oi_{date_str}"

        if cache_key in _daily_futures_oi_cache:
            records = _daily_futures_oi_cache[cache_key]
        else:
            content = fetcher.download_daily_oi_excel(td)
            if content is None:
                _daily_futures_oi_cache[cache_key] = []
                continue
            records = parse_daily_futures_oi_excel(content)
            _daily_futures_oi_cache[cache_key] = records

        for r in records:
            if r.product == product and r.contract_month == contract_month:
                result[td] = r
                break

    # Derive previous day's OI from next day's previous_oi field
    trading_day_set = set(week.trading_days)
    for td in list(result.keys()):
        rec = result[td]
        if rec.previous_oi > 0:
            prev_td = _get_prev_trading_date(td)
            if prev_td is not None and prev_td in trading_day_set and prev_td not in result:
                # Reconstruct previous day's record from current day's previous_oi
                # We know: prev day current_oi = this day previous_oi
                # prev day net_change is unknown, approximate from OI difference
                # Look for the day before prev_td to get prev_prev_oi
                prev_prev_td = _get_prev_trading_date(prev_td)
                prev_prev_oi = 0
                if prev_prev_td and prev_prev_td in result:
                    prev_prev_oi = result[prev_prev_td].current_oi

                derived_change = rec.previous_oi - prev_prev_oi if prev_prev_oi > 0 else 0
                result[prev_td] = DailyFuturesOI(
                    report_date=prev_td,
                    product=product,
                    contract_month=contract_month,
                    trading_volume=0,
                    current_oi=rec.previous_oi,
                    net_change=derived_change,
                    previous_oi=prev_prev_oi,
                )

    return result


def _load_option_oi_raw(d: date) -> list[OptionParticipantOI]:
    """Load raw option OI records for a date (cached)."""
    year = str(d.year)
    date_str = d.strftime("%Y%m%d")

    try:
        entries = fetcher.get_oi_index(year)
    except Exception:
        return []

    for entry in entries:
        if entry["TradeDate"] == date_str:
            file_path = entry.get("IndexOptions")
            if not file_path:
                return []
            try:
                if file_path in _option_oi_parse_cache:
                    return _option_oi_parse_cache[file_path]
                content = fetcher.download_oi_excel(file_path)
                records = parse_option_oi_excel(content)
                _option_oi_parse_cache[file_path] = records
                return records
            except Exception:
                return []

    return []


def _load_option_oi_for_date(
    d: date,
    contract_month: str = "",
    participant_ids: list[str] | None = None,
) -> dict[tuple[str, int], tuple[float, float]]:
    """Load option OI for a date.

    Returns {(option_type, strike_price): (total_long, total_short)}
    summed across filtered participants.
    """
    records = _load_option_oi_raw(d)

    # Apply filters
    filtered = records
    if contract_month:
        filtered = [r for r in filtered if r.contract_month == contract_month]
    if participant_ids is not None:
        pid_set = set(participant_ids)
        filtered = [r for r in filtered if r.participant_id in pid_set]

    # Aggregate long/short per (type, strike)
    agg: dict[tuple[str, int], list[float]] = {}
    for r in filtered:
        key = (r.option_type, r.strike_price)
        if key not in agg:
            agg[key] = [0.0, 0.0]
        agg[key][0] += (r.long_volume or 0)
        agg[key][1] += (r.short_volume or 0)
    return {k: (v[0], v[1]) for k, v in agg.items()}


def _load_option_volume_raw_session(
    jpx_trade_date: date,
    file_keys: tuple[str, ...],
) -> list[OptionParticipantVolume]:
    """Load option volume records for specific session files."""
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
                        if path in _option_volume_parse_cache:
                            records = _option_volume_parse_cache[path]
                        else:
                            content = fetcher.download_volume_excel(path)
                            records = parse_option_volume_excel(content)
                            _option_volume_parse_cache[path] = records
                        all_records.append(records)
                    except Exception:
                        pass
            return merge_option_volume_records(*all_records)

    return []


def _redate_option_records(
    records: list[OptionParticipantVolume], new_date: date,
) -> list[OptionParticipantVolume]:
    for r in records:
        r.trade_date = new_date
    return records


def _load_option_volume_for_market_date(
    market_date: date,
    session_mode,
) -> list[OptionParticipantVolume]:
    """Load option volume with Night session shifting (same logic as futures)."""
    _ensure_trading_date_index()

    if session_mode == SESSION_ALL:
        day_records = _load_option_volume_raw_session(
            market_date, SESSION_DAY_KEYS
        )
        night_records = []
        next_td = _get_next_trading_date(market_date)
        if next_td:
            night_records = _load_option_volume_raw_session(
                next_td, SESSION_NIGHT_KEYS
            )
            _redate_option_records(night_records, market_date)
        return merge_option_volume_records(day_records, night_records)

    requested_keys = session_mode
    is_night = all(k in SESSION_NIGHT_KEYS for k in requested_keys)

    if is_night:
        next_td = _get_next_trading_date(market_date)
        if not next_td:
            return []
        records = _load_option_volume_raw_session(next_td, requested_keys)
        _redate_option_records(records, market_date)
        return records
    else:
        return _load_option_volume_raw_session(market_date, requested_keys)


def _aggregate_by_strike(
    start_oi: dict[tuple[str, int], tuple[float, float]],
    end_oi: dict[tuple[str, int], tuple[float, float]],
    daily_vols: dict[date, list[OptionParticipantVolume]],
    week: WeekDefinition,
    daily_oi: dict[date, list[DailyOIBalance]] | None = None,
) -> list[OptionStrikeRow]:
    """Aggregate all data into OptionStrikeRow per strike price.

    Each row contains all-participant totals for PUT and CALL.
    OI values are (total_long, total_short) tuples.
    daily_oi: per-date aggregate OI balance (not per-participant).
    """
    # Collect all strikes
    all_strikes: set[int] = set()
    for (ot, s) in start_oi:
        all_strikes.add(s)
    for (ot, s) in end_oi:
        all_strikes.add(s)
    for day_records in daily_vols.values():
        for r in day_records:
            all_strikes.add(r.strike_price)
    if daily_oi:
        for day_records in daily_oi.values():
            for r in day_records:
                all_strikes.add(r.strike_price)

    # Build volume aggregation: (date, type, strike) -> total volume
    vol_agg: dict[tuple[date, str, int], float] = {}
    # Build per-participant breakdown: (date, type, strike) -> [(name, vol), ...]
    vol_detail: dict[tuple[date, str, int], list[tuple[str, float]]] = {}
    for td, records in daily_vols.items():
        for r in records:
            key = (td, r.option_type, r.strike_price)
            vol_agg[key] = vol_agg.get(key, 0) + r.volume
            name = r.participant_name_en or r.participant_name_jp or r.participant_id
            vol_detail.setdefault(key, []).append((name, r.volume))

    # Sort breakdowns by volume descending
    for key in vol_detail:
        vol_detail[key].sort(key=lambda x: -x[1])

    # Build daily OI balance lookup: (date, type, strike) -> DailyOIBalance
    oi_bal_lookup: dict[tuple[date, str, int], DailyOIBalance] = {}
    if daily_oi:
        for td, records in daily_oi.items():
            for r in records:
                oi_bal_lookup[(td, r.option_type, r.strike_price)] = r

        # Derive previous day's OI from current day's previous_oi field
        trading_day_set = set(week.trading_days)
        for td, records in daily_oi.items():
            prev_td = _get_prev_trading_date(td)
            if prev_td is None or prev_td not in trading_day_set:
                continue
            for r in records:
                prev_key = (prev_td, r.option_type, r.strike_price)
                if prev_key not in oi_bal_lookup:
                    oi_bal_lookup[prev_key] = DailyOIBalance(
                        report_date=prev_td,
                        contract_month=r.contract_month,
                        option_type=r.option_type,
                        strike_price=r.strike_price,
                        trading_volume=0,
                        current_oi=r.previous_oi,
                        net_change=0,
                        previous_oi=0,
                    )

    rows = []
    for strike in sorted(all_strikes, reverse=True):
        put_daily = {}
        put_total = 0.0
        call_daily = {}
        call_total = 0.0
        put_breakdown = {}
        call_breakdown = {}
        put_doi = {}
        put_doi_chg = {}
        call_doi = {}
        call_doi_chg = {}
        put_jpx_vol = {}
        call_jpx_vol = {}

        for td in week.trading_days:
            pv = vol_agg.get((td, "PUT", strike), 0)
            cv = vol_agg.get((td, "CALL", strike), 0)
            if pv > 0:
                put_daily[td] = pv
                put_total += pv
                put_breakdown[td] = vol_detail.get((td, "PUT", strike), [])
            if cv > 0:
                call_daily[td] = cv
                call_total += cv
                call_breakdown[td] = vol_detail.get((td, "CALL", strike), [])

            # Daily OI balance + JPX aggregate volume
            p_bal = oi_bal_lookup.get((td, "PUT", strike))
            if p_bal:
                put_doi[td] = p_bal.current_oi
                put_doi_chg[td] = p_bal.net_change
                if p_bal.trading_volume > 0:
                    put_jpx_vol[td] = p_bal.trading_volume
            c_bal = oi_bal_lookup.get((td, "CALL", strike))
            if c_bal:
                call_doi[td] = c_bal.current_oi
                call_doi_chg[td] = c_bal.net_change
                if c_bal.trading_volume > 0:
                    call_jpx_vol[td] = c_bal.trading_volume

        ps = start_oi.get(("PUT", strike))  # (long, short) or None
        pe = end_oi.get(("PUT", strike))
        cs = start_oi.get(("CALL", strike))
        ce = end_oi.get(("CALL", strike))

        rows.append(OptionStrikeRow(
            strike_price=strike,
            put_start_oi_long=ps[0] if ps else None,
            put_start_oi_short=ps[1] if ps else None,
            put_end_oi_long=pe[0] if pe else None,
            put_end_oi_short=pe[1] if pe else None,
            put_daily_volumes=put_daily,
            put_week_total=put_total if put_total > 0 else None,
            call_start_oi_long=cs[0] if cs else None,
            call_start_oi_short=cs[1] if cs else None,
            call_end_oi_long=ce[0] if ce else None,
            call_end_oi_short=ce[1] if ce else None,
            call_daily_volumes=call_daily,
            call_week_total=call_total if call_total > 0 else None,
            put_daily_breakdown=put_breakdown,
            call_daily_breakdown=call_breakdown,
            put_daily_oi=put_doi,
            put_daily_oi_change=put_doi_chg,
            call_daily_oi=call_doi,
            call_daily_oi_change=call_doi_chg,
            put_daily_jpx_volume=put_jpx_vol,
            call_daily_jpx_volume=call_jpx_vol,
        ))

    return rows
