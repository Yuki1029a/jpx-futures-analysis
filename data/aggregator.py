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
    OptionParticipantOI, OptionParticipantVolume, OptionStrikeRow,
)
from data import fetcher
from data.parser_volume import (
    parse_volume_excel, merge_volume_records,
    parse_option_volume_excel, merge_option_volume_records,
)
from data.parser_oi import parse_oi_excel
from data.parser_option_oi import parse_option_oi_excel
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


# =====================================================================
# Option aggregation
# =====================================================================

def load_option_weekly_data(
    week: WeekDefinition,
    session_keys=SESSION_ALL,
) -> list[OptionStrikeRow]:
    """Load option data for a week: OI + daily volumes, aggregated by strike.

    Returns OptionStrikeRow per strike price with all-participant totals.
    """
    # 1. Load OI
    start_oi = _load_option_oi_for_date(week.start_oi_date)
    end_oi = {}
    if week.end_oi_date:
        end_oi = _load_option_oi_for_date(week.end_oi_date)

    # 2. Load daily option volumes
    daily_vols: dict[date, list[OptionParticipantVolume]] = {}
    for td in week.trading_days:
        records = _load_option_volume_for_market_date(td, session_keys)
        daily_vols[td] = records

    # 3. Aggregate by strike
    return _aggregate_by_strike(start_oi, end_oi, daily_vols, week)


def _load_option_oi_for_date(d: date) -> dict[tuple[str, int], tuple[float, float]]:
    """Load option OI for a date.

    Returns {(option_type, strike_price): (total_long, total_short)}
    summed across all participants.
    """
    year = str(d.year)
    date_str = d.strftime("%Y%m%d")

    try:
        entries = fetcher.get_oi_index(year)
    except Exception:
        return {}

    for entry in entries:
        if entry["TradeDate"] == date_str:
            file_path = entry.get("IndexOptions")
            if not file_path:
                return {}
            try:
                if file_path in _option_oi_parse_cache:
                    records = _option_oi_parse_cache[file_path]
                else:
                    content = fetcher.download_oi_excel(file_path)
                    records = parse_option_oi_excel(content)
                    _option_oi_parse_cache[file_path] = records
                # Aggregate long/short per (type, strike)
                agg: dict[tuple[str, int], list[float]] = {}
                for r in records:
                    key = (r.option_type, r.strike_price)
                    if key not in agg:
                        agg[key] = [0.0, 0.0]
                    agg[key][0] += (r.long_volume or 0)
                    agg[key][1] += (r.short_volume or 0)
                return {k: (v[0], v[1]) for k, v in agg.items()}
            except Exception:
                return {}

    return {}


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
) -> list[OptionStrikeRow]:
    """Aggregate all data into OptionStrikeRow per strike price.

    Each row contains all-participant totals for PUT and CALL.
    OI values are (total_long, total_short) tuples.
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

    # Build volume aggregation: (date, type, strike) -> total volume
    vol_agg: dict[tuple[date, str, int], float] = {}
    for td, records in daily_vols.items():
        for r in records:
            key = (td, r.option_type, r.strike_price)
            vol_agg[key] = vol_agg.get(key, 0) + r.volume

    rows = []
    for strike in sorted(all_strikes, reverse=True):
        put_daily = {}
        put_total = 0.0
        call_daily = {}
        call_total = 0.0

        for td in week.trading_days:
            pv = vol_agg.get((td, "PUT", strike), 0)
            cv = vol_agg.get((td, "CALL", strike), 0)
            if pv > 0:
                put_daily[td] = pv
                put_total += pv
            if cv > 0:
                call_daily[td] = cv
                call_total += cv

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
        ))

    return rows
