"""Data models for the application."""

from dataclasses import dataclass, field
from datetime import date
from typing import Optional


@dataclass
class ParticipantVolume:
    """One participant's daily trading volume for a specific product/contract."""
    trade_date: date
    product: str                # "NK225F", "TOPIXF"
    contract_month: str         # YYMM format, e.g., "2603"
    participant_id: str         # 5-digit string
    participant_name_en: str
    participant_name_jp: str
    rank: int
    volume: float               # Combined buy+sell total (Day+Night sum)
    volume_day: float           # Day session only
    volume_night: float         # Night session only


@dataclass
class ParticipantOI:
    """One participant's open interest position."""
    report_date: date
    product: str
    contract_month: str         # YYMM format
    participant_id: str
    participant_name_jp: str
    long_volume: Optional[float]
    short_volume: Optional[float]


@dataclass
class WeekDefinition:
    """A trading week bounded by two OI report dates."""
    start_oi_date: date         # Previous Friday's OI date
    end_oi_date: date           # Current Friday's OI date
    trading_days: list[date] = field(default_factory=list)
    label: str = ""


@dataclass
class WeeklyParticipantRow:
    """Aggregated weekly data for one participant, used for display."""
    participant_id: str
    participant_name: str
    start_oi_long: Optional[float] = None
    start_oi_short: Optional[float] = None
    start_oi_net: Optional[float] = None
    daily_volumes: dict = field(default_factory=dict)  # date -> volume
    end_oi_long: Optional[float] = None
    end_oi_short: Optional[float] = None
    end_oi_net: Optional[float] = None
    oi_net_change: Optional[float] = None
    inferred_direction: Optional[str] = None  # "BUY", "SELL", "NEUTRAL"
    avg_20d: Optional[float] = None           # 20-day average daily volume
    max_20d: Optional[float] = None           # 20-day max daily volume


@dataclass
class OptionParticipantOI:
    """One participant's option open interest position."""
    report_date: date
    option_type: str            # "PUT" or "CALL"
    strike_price: int           # Strike price (e.g., 38000)
    participant_id: str
    participant_name_jp: str
    long_volume: Optional[float]
    short_volume: Optional[float]


@dataclass
class OptionParticipantVolume:
    """One participant's daily option trading volume."""
    trade_date: date
    option_type: str            # "PUT" or "CALL"
    strike_price: int           # Strike price
    participant_id: str
    participant_name_en: str
    participant_name_jp: str
    rank: int
    volume: float               # Combined buy+sell total (Day+Night sum)
    volume_day: float           # Day session only
    volume_night: float         # Night session only


@dataclass
class OptionStrikeRow:
    """Aggregated weekly data for one strike price, used for display."""
    strike_price: int
    # PUT side data
    put_start_oi_long: Optional[float] = None
    put_start_oi_short: Optional[float] = None
    put_end_oi_long: Optional[float] = None
    put_end_oi_short: Optional[float] = None
    put_daily_volumes: dict = field(default_factory=dict)  # date -> total_volume
    put_week_total: Optional[float] = None
    # CALL side data
    call_start_oi_long: Optional[float] = None
    call_start_oi_short: Optional[float] = None
    call_end_oi_long: Optional[float] = None
    call_end_oi_short: Optional[float] = None
    call_daily_volumes: dict = field(default_factory=dict)  # date -> total_volume
    call_week_total: Optional[float] = None
