"""Gamma Exposure (GEX) calculation using simplified Black-Scholes.

Assumptions:
  - Flat IV across all strikes (user-adjustable)
  - Fixed risk-free rate
  - SQ date = 2nd Friday of the contract month
  - Contract multiplier = 1000 (NK225 options)

Convention (dealer perspective):
  - CALL GEX > 0: dealers are long gamma (hedging stabilises price)
  - PUT GEX < 0: dealers are short gamma (hedging amplifies moves)
"""
from __future__ import annotations

import math
from datetime import date
from typing import NamedTuple

import numpy as np
import pandas as pd
from scipy.stats import norm


class GEXProfile(NamedTuple):
    df: pd.DataFrame          # columns: strike, call_gex, put_gex, net_gex
    flip_point: float | None  # strike where net GEX crosses zero
    total_call_gex: float
    total_put_gex: float
    total_net_gex: float


def calc_gex_profile(
    strikes: list[int],
    put_oi: dict[int, int],
    call_oi: dict[int, int],
    spot: float,
    expiry_date: date,
    as_of: date,
    sigma: float = 0.20,
    r: float = 0.005,
    contract_multiplier: float = 1000.0,
) -> GEXProfile:
    """Compute GEX profile across strikes.

    Parameters
    ----------
    strikes : sorted list of strike prices
    put_oi / call_oi : {strike: open_interest}
    spot : underlying price
    expiry_date : SQ date
    as_of : calculation date
    sigma : implied volatility (flat)
    r : risk-free rate
    contract_multiplier : 1000 for NK225 options
    """
    T = max((expiry_date - as_of).days, 0) / 365.0

    records = []
    for K in sorted(strikes):
        gamma = _bs_gamma(spot, K, T, sigma, r) if T > 0 else 0.0

        c_oi = call_oi.get(K, 0)
        p_oi = put_oi.get(K, 0)

        # GEX in notional terms (yen)
        # Dealer is long gamma on calls sold, short gamma on puts sold
        call_gex = gamma * c_oi * spot * contract_multiplier
        put_gex = -gamma * p_oi * spot * contract_multiplier

        records.append({
            "strike": K,
            "call_gex": call_gex,
            "put_gex": put_gex,
            "net_gex": call_gex + put_gex,
        })

    df = pd.DataFrame(records)

    total_call = df["call_gex"].sum()
    total_put = df["put_gex"].sum()
    total_net = df["net_gex"].sum()

    # Find flip point (net_gex sign change, nearest to spot)
    flip = _find_flip_point(df, spot)

    return GEXProfile(
        df=df,
        flip_point=flip,
        total_call_gex=total_call,
        total_put_gex=total_put,
        total_net_gex=total_net,
    )


def get_sq_date(contract_month: str) -> date:
    """Derive SQ date (2nd Friday) from YYMM contract month string.

    e.g. "2603" -> 2026-03-13 (2nd Friday of March 2026)
    """
    yy = int(contract_month[:2])
    mm = int(contract_month[2:])
    year = 2000 + yy

    # Find 2nd Friday: first day of month, advance to first Friday, then +7
    first_day = date(year, mm, 1)
    # weekday: 0=Mon ... 4=Fri
    days_to_fri = (4 - first_day.weekday()) % 7
    first_friday = first_day.day + days_to_fri
    if first_friday == 0:
        first_friday = 7
    second_friday = first_friday + 7

    return date(year, mm, second_friday)


def calc_gex_surface(
    strikes: list[int],
    put_oi: dict[int, int],
    call_oi: dict[int, int],
    spot_center: float,
    spot_range: float,
    spot_step: float,
    expiry_date: date,
    as_of: date,
    sigma: float = 0.20,
    r: float = 0.005,
    contract_multiplier: float = 1000.0,
) -> tuple[np.ndarray, np.ndarray, np.ndarray]:
    """Compute Net GEX surface over spot range x strikes.

    Returns
    -------
    spots : 1-D array of spot values (Y axis)
    strike_arr : 1-D array of strike prices (X axis)
    surface : 2-D array shape (len(spots), len(strikes)), net_gex per cell
    """
    T = max((expiry_date - as_of).days, 0) / 365.0
    sorted_strikes = sorted(strikes)
    strike_arr = np.array(sorted_strikes, dtype=float)

    spots = np.arange(
        spot_center - spot_range,
        spot_center + spot_range + spot_step,
        spot_step,
    )

    c_oi_arr = np.array([call_oi.get(K, 0) for K in sorted_strikes], dtype=float)
    p_oi_arr = np.array([put_oi.get(K, 0) for K in sorted_strikes], dtype=float)

    surface = np.zeros((len(spots), len(sorted_strikes)))

    for i, S in enumerate(spots):
        for j, K in enumerate(sorted_strikes):
            gamma = _bs_gamma(S, K, T, sigma, r) if T > 0 else 0.0
            call_gex = gamma * c_oi_arr[j] * S * contract_multiplier
            put_gex = -gamma * p_oi_arr[j] * S * contract_multiplier
            surface[i, j] = call_gex + put_gex

    return spots, strike_arr, surface


# --- Internal ---

def _bs_gamma(S: float, K: float, T: float, sigma: float, r: float) -> float:
    """Black-Scholes gamma (same for call and put)."""
    if T <= 0 or sigma <= 0 or S <= 0 or K <= 0:
        return 0.0

    sqrt_T = math.sqrt(T)
    d1 = (math.log(S / K) + (r + 0.5 * sigma ** 2) * T) / (sigma * sqrt_T)

    return norm.pdf(d1) / (S * sigma * sqrt_T)


def _find_flip_point(df: pd.DataFrame, spot: float) -> float | None:
    """Find the strike nearest to spot where net_gex crosses zero."""
    if df.empty:
        return None

    net = df["net_gex"].values
    strikes = df["strike"].values

    best_flip = None
    best_dist = float("inf")

    for i in range(len(net) - 1):
        if net[i] * net[i + 1] < 0:  # sign change
            # Linear interpolation
            frac = abs(net[i]) / (abs(net[i]) + abs(net[i + 1]))
            flip_strike = strikes[i] + frac * (strikes[i + 1] - strikes[i])
            dist = abs(flip_strike - spot)
            if dist < best_dist:
                best_dist = dist
                best_flip = flip_strike

    return best_flip
