"""Option Pain (Max Pain) chart rendering.

Computes the settlement price that minimizes total option payout.
Includes:
- Max Pain time series vs NK225
- Full-range OI distribution profile with Max Pain overlay
- Per-month pain profile charts
"""
from __future__ import annotations

import streamlit as st
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import date, datetime
from collections import defaultdict

from models import OptionStrikeRow, WeekDefinition, DailyOIBalance

_DOW_JP = ["月", "火", "水", "木", "金", "土", "日"]


def render_option_pain_section(
    all_month_rows: dict[str, list[OptionStrikeRow]],
    week: WeekDefinition,
) -> None:
    """Render option pain charts for multiple contract months."""
    st.subheader("オプションペイン分析")

    if not all_month_rows:
        st.info("オプションデータなし")
        return

    # 1. Max Pain time series (all months)
    _render_maxpain_timeseries()

    st.markdown("---")

    # 2. OI distribution profile + Max Pain driver
    _render_oi_profile_section(sorted(all_month_rows.keys()))



def _format_cm(cm: str) -> str:
    if not cm:
        return "-"
    return f"20{cm[:2]}年{cm[2:]}月限"


def _date_label(td: date) -> str:
    return f"{td.strftime('%m/%d')}({_DOW_JP[td.weekday()]})"


# =====================================================================
# Max Pain time series
# =====================================================================

@st.cache_data(ttl=600, show_spinner=False)
def _load_maxpain_timeseries_data() -> tuple[
    dict[date, dict[str, int | None]],
    list[str],
    list[date],
    list[float],
]:
    """Load and compute max pain time series data (cached 10 min)."""
    from data import fetcher
    from data.aggregator import _load_daily_oi_for_date

    months = fetcher.get_available_volume_months()
    all_dates: list[date] = []
    for m in months[:4]:
        try:
            entries = fetcher.get_volume_index(m)
        except Exception:
            continue
        for entry in entries:
            d = datetime.strptime(entry["TradeDate"], "%Y%m%d").date()
            all_dates.append(d)
    all_dates = sorted(set(all_dates))[-30:]

    if not all_dates:
        return {}, [], [], []

    maxpain_data: dict[date, dict[str, int | None]] = {}
    all_cms: set[str] = set()

    for td in all_dates:
        try:
            records = _load_daily_oi_for_date(td)
        except Exception:
            continue
        if not records:
            continue
        mp = _compute_max_pain_from_balance(records)
        if mp:
            maxpain_data[td] = mp
            all_cms.update(mp.keys())

    contract_months = []
    for cm in sorted(all_cms):
        count = sum(1 for d in maxpain_data if maxpain_data[d].get(cm) is not None)
        if count >= 3:
            contract_months.append(cm)

    nk_dates: list[date] = []
    nk_prices: list[float] = []
    try:
        import yfinance as yf
        hist = yf.Ticker("^N225").history(period="3mo")
        if not hist.empty:
            nk_dates = [d.date() for d in hist.index]
            nk_prices = hist["Close"].tolist()
    except Exception:
        pass

    return maxpain_data, contract_months, nk_dates, nk_prices


def _compute_max_pain_from_balance(
    records: list[DailyOIBalance],
) -> dict[str, int | None]:
    """Compute Max Pain per contract month from DailyOIBalance records."""
    cm_data: dict[str, dict[int, dict[str, int]]] = defaultdict(
        lambda: defaultdict(lambda: {"CALL": 0, "PUT": 0})
    )
    for r in records:
        if r.current_oi > 0:
            cm_data[r.contract_month][r.strike_price][r.option_type] += r.current_oi

    results: dict[str, int | None] = {}
    for cm, strike_map in cm_data.items():
        results[cm] = _calc_max_pain(strike_map)
    return results


def _calc_max_pain(
    strike_map: dict[int, dict[str, int]],
) -> int | None:
    """Find strike that minimizes total option payout."""
    if not strike_map:
        return None

    strikes = sorted(strike_map.keys())
    call_oi = {k: strike_map[k].get("CALL", 0) for k in strikes}
    put_oi = {k: strike_map[k].get("PUT", 0) for k in strikes}

    if sum(call_oi.values()) + sum(put_oi.values()) == 0:
        return None

    best_strike = None
    min_pain = float("inf")
    for S in strikes:
        pain = sum(max(0, S - K) * oi for K, oi in call_oi.items())
        pain += sum(max(0, K - S) * oi for K, oi in put_oi.items())
        if pain < min_pain:
            min_pain = pain
            best_strike = S
    return best_strike


def _render_maxpain_timeseries() -> None:
    """Render the Max Pain time series chart with NK225 overlay."""
    with st.spinner("Max Pain時系列データ読み込み中..."):
        maxpain_data, contract_months, nk_dates, nk_prices = (
            _load_maxpain_timeseries_data()
        )

    if not maxpain_data or not contract_months:
        st.info("Max Pain時系列データなし")
        return

    fig = go.Figure()

    if nk_dates:
        fig.add_trace(go.Scatter(
            x=nk_dates, y=nk_prices,
            mode="lines", name="NK225終値",
            line=dict(color="black", width=2.5),
        ))

    colors = ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd", "#8c564b"]
    sorted_dates = sorted(maxpain_data.keys())

    for i, cm in enumerate(contract_months):
        cm_dates = [d for d in sorted_dates if maxpain_data[d].get(cm) is not None]
        cm_values = [maxpain_data[d][cm] for d in cm_dates]
        color = colors[i % len(colors)]
        fig.add_trace(go.Scatter(
            x=cm_dates, y=cm_values,
            mode="lines+markers",
            name=f"MaxPain {_format_cm(cm)}",
            line=dict(color=color, width=1.5, dash="dot"),
            marker=dict(size=6, color=color),
        ))

    fig.update_layout(
        title="Max Pain 時系列 vs NK225終値",
        xaxis=dict(title="日付", type="date"),
        yaxis=dict(title="価格", tickformat=","),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
            bgcolor="rgba(255,255,255,0.8)",
        ),
        hovermode="x unified",
        template="plotly_white",
        height=500,
        margin=dict(l=0, r=0, t=40, b=0),
    )

    st.plotly_chart(fig, use_container_width=True)

    # Summary metrics
    nk_lookup = dict(zip(nk_dates, nk_prices))
    latest_date = sorted_dates[-1] if sorted_dates else None
    if latest_date:
        cols = st.columns(len(contract_months) + 1)
        nk_close = nk_lookup.get(latest_date)
        with cols[0]:
            if nk_close:
                st.metric("NK225終値", f"{nk_close:,.0f}")
        for i, cm in enumerate(contract_months):
            mp = maxpain_data[latest_date].get(cm)
            with cols[i + 1]:
                if mp:
                    delta = None
                    if nk_close:
                        delta = f"{mp - nk_close:+,.0f}"
                    st.metric(f"MaxPain {cm[:2]}/{cm[2:]}", f"{mp:,}", delta=delta)


# =====================================================================
# OI Distribution Profile + Max Pain Driver
# =====================================================================

@st.cache_data(ttl=600, show_spinner=False)
def _load_oi_snapshots(
    contract_month: str, n_days: int = 5,
) -> tuple[
    list[date],
    list[dict[int, int]],   # put_oi per day
    list[dict[int, int]],   # call_oi per day
    list[int | None],       # max_pain per day
]:
    """Load daily OI snapshots for last n_days trading dates (cached)."""
    from data import fetcher
    from data.aggregator import _load_daily_oi_for_date

    # Get recent trading dates
    months = fetcher.get_available_volume_months()
    all_dates: list[date] = []
    for m in months[:4]:
        try:
            entries = fetcher.get_volume_index(m)
        except Exception:
            continue
        for entry in entries:
            d = datetime.strptime(entry["TradeDate"], "%Y%m%d").date()
            all_dates.append(d)
    all_dates = sorted(set(all_dates))

    # Find dates with actual OI data
    valid_dates = []
    put_oi_list = []
    call_oi_list = []
    mp_list = []

    for td in reversed(all_dates):
        if len(valid_dates) >= n_days:
            break
        try:
            records = _load_daily_oi_for_date(td, contract_month)
        except Exception:
            continue
        if not records:
            continue

        put_oi: dict[int, int] = {}
        call_oi: dict[int, int] = {}
        sm: dict[int, dict[str, int]] = defaultdict(lambda: {"CALL": 0, "PUT": 0})

        for r in records:
            if r.current_oi > 0:
                if r.option_type == "PUT":
                    put_oi[r.strike_price] = r.current_oi
                else:
                    call_oi[r.strike_price] = r.current_oi
                sm[r.strike_price][r.option_type] += r.current_oi

        if not sm:
            continue

        valid_dates.append(td)
        put_oi_list.append(put_oi)
        call_oi_list.append(call_oi)
        mp_list.append(_calc_max_pain(sm))

    # Reverse to chronological order
    valid_dates.reverse()
    put_oi_list.reverse()
    call_oi_list.reverse()
    mp_list.reverse()

    return valid_dates, put_oi_list, call_oi_list, mp_list


def _render_oi_profile_section(available_cms: list[str]) -> None:
    """Render OI change heatmap and distribution with Max Pain."""
    st.subheader("建玉変動 vs Max Pain")

    if not available_cms:
        return

    selected_cm = st.selectbox(
        "限月選択", available_cms,
        format_func=_format_cm,
        key="oi_profile_cm",
    )

    with st.spinner("建玉データ読み込み中..."):
        dates, put_oi_list, call_oi_list, mp_list = _load_oi_snapshots(
            selected_cm, n_days=5,
        )

    if len(dates) < 2:
        st.info("データ不足（2日以上必要）")
        return

    # Collect all strikes across all days
    all_strikes: set[int] = set()
    for po, co in zip(put_oi_list, call_oi_list):
        all_strikes.update(po.keys())
        all_strikes.update(co.keys())
    strikes = sorted(all_strikes)

    if len(strikes) < 3:
        st.info("有効行使価格が不足")
        return

    # Load NK225 prices for overlay
    _, _, nk_dates, nk_prices = _load_maxpain_timeseries_data()
    nk_lookup = dict(zip(nk_dates, nk_prices))

    # --- Primary: OI change heatmap ---
    _render_oi_change_heatmap(
        dates, strikes, put_oi_list, call_oi_list, mp_list, selected_cm,
        nk_lookup=nk_lookup,
    )

    # --- Max Pain summary metrics ---
    cols = st.columns(len(dates))
    for i, (td, mp) in enumerate(zip(dates, mp_list)):
        with cols[i]:
            prev_mp = mp_list[i - 1] if i > 0 else None
            delta = None
            if prev_mp and mp:
                diff = mp - prev_mp
                if diff != 0:
                    delta = f"{diff:+,}"
            st.metric(
                _date_label(td),
                f"{mp:,}" if mp else "N/A",
                delta=delta,
            )

    # --- Secondary: OI distribution overlay (collapsible) ---
    with st.expander("建玉分布オーバーレイ (全行使価格)"):
        _render_oi_distribution_chart(
            dates, strikes, put_oi_list, call_oi_list, mp_list, selected_cm,
        )


def _render_oi_distribution_chart(
    dates: list[date],
    strikes: list[int],
    put_oi_list: list[dict[int, int]],
    call_oi_list: list[dict[int, int]],
    mp_list: list[int | None],
    contract_month: str,
) -> None:
    """OI distribution: latest vs previous day, bars offset side-by-side."""
    if len(dates) < 2:
        return

    fig = go.Figure()

    # Use latest 2 days only
    curr_idx = len(dates) - 1
    prev_idx = len(dates) - 2
    curr_label = _date_label(dates[curr_idx])
    prev_label = _date_label(dates[prev_idx])

    curr_po = put_oi_list[curr_idx]
    prev_po = put_oi_list[prev_idx]
    curr_co = call_oi_list[curr_idx]
    prev_co = call_oi_list[prev_idx]

    # Compute offset: ~40% of typical strike interval
    intervals = [strikes[i + 1] - strikes[i] for i in range(min(20, len(strikes) - 1))]
    offset = int(np.median(intervals) * 0.35) if intervals else 50

    # Previous day bars (lighter, offset left)
    prev_x = [s - offset for s in strikes]
    fig.add_trace(go.Bar(
        x=prev_x,
        y=[-prev_po.get(s, 0) for s in strikes],
        name=f"PUT {prev_label}",
        marker_color="rgba(220, 60, 60, 0.25)",
        width=offset * 1.8,
        legendgroup="prev_put",
        hovertemplate=f"PUT {prev_label}<br>行使価格: %{{customdata:,}}<br>OI: %{{meta:,}}<extra></extra>",
        customdata=strikes,
        meta=[prev_po.get(s, 0) for s in strikes],
    ))
    fig.add_trace(go.Bar(
        x=prev_x,
        y=[prev_co.get(s, 0) for s in strikes],
        name=f"CALL {prev_label}",
        marker_color="rgba(60, 120, 220, 0.25)",
        width=offset * 1.8,
        legendgroup="prev_call",
        hovertemplate=f"CALL {prev_label}<br>行使価格: %{{customdata:,}}<br>OI: %{{y:,}}<extra></extra>",
        customdata=strikes,
    ))

    # Current day bars (darker, offset right)
    curr_x = [s + offset for s in strikes]
    fig.add_trace(go.Bar(
        x=curr_x,
        y=[-curr_po.get(s, 0) for s in strikes],
        name=f"PUT {curr_label}",
        marker_color="rgba(220, 60, 60, 0.8)",
        width=offset * 1.8,
        legendgroup="curr_put",
        hovertemplate=f"PUT {curr_label}<br>行使価格: %{{customdata:,}}<br>OI: %{{meta:,}}<extra></extra>",
        customdata=strikes,
        meta=[curr_po.get(s, 0) for s in strikes],
    ))
    fig.add_trace(go.Bar(
        x=curr_x,
        y=[curr_co.get(s, 0) for s in strikes],
        name=f"CALL {curr_label}",
        marker_color="rgba(60, 120, 220, 0.8)",
        width=offset * 1.8,
        legendgroup="curr_call",
        hovertemplate=f"CALL {curr_label}<br>行使価格: %{{customdata:,}}<br>OI: %{{y:,}}<extra></extra>",
        customdata=strikes,
    ))

    # Max Pain lines: previous (bottom label) and current (top label)
    mp_prev = mp_list[prev_idx]
    mp_curr = mp_list[curr_idx]
    if mp_prev is not None:
        fig.add_vline(
            x=mp_prev, line_dash="dash",
            line_color="rgba(0, 180, 0, 0.4)", line_width=1.5,
        )
        fig.add_annotation(
            x=mp_prev, y=0, yref="y domain",
            text=f"MP {prev_label}: {mp_prev:,}",
            showarrow=False,
            font=dict(color="rgba(0,150,0,0.7)", size=10),
            bgcolor="rgba(255,255,255,0.8)",
            yanchor="bottom",
        )
    if mp_curr is not None:
        fig.add_vline(
            x=mp_curr, line_dash="solid",
            line_color="green", line_width=2,
        )
        fig.add_annotation(
            x=mp_curr, y=1, yref="y domain",
            text=f"MP {curr_label}: {mp_curr:,}",
            showarrow=False,
            font=dict(color="green", size=11, weight="bold"),
            bgcolor="rgba(255,255,255,0.8)",
            yanchor="top",
        )

    fig.update_layout(
        title=f"建玉分布 - {_format_cm(contract_month)}  ({prev_label} vs {curr_label})",
        xaxis=dict(title="行使価格", tickformat=","),
        yaxis=dict(title="建玉枚数 (上=CALL / 下=PUT)"),
        barmode="overlay",
        template="plotly_white",
        height=500,
        margin=dict(l=0, r=0, t=40, b=0),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
        ),
    )

    st.plotly_chart(fig, use_container_width=True)
    st.caption(
        f"薄色={prev_label}(前日) / 濃色={curr_label}(当日) / "
        "緑実線=当日MP / 緑破線=前日MP / 上=CALL / 下=PUT"
    )


def _render_oi_change_heatmap(
    dates: list[date],
    strikes: list[int],
    put_oi_list: list[dict[int, int]],
    call_oi_list: list[dict[int, int]],
    mp_list: list[int | None],
    contract_month: str,
    nk_lookup: dict[date, float] | None = None,
) -> None:
    """Vertical subplots of daily OI changes, filtered to active strikes.

    Each row = one day transition (full chart width).
    PUT changes shown as red bars (upward), CALL as blue (downward/inverted).
    Only strikes with non-zero changes are displayed.
    """
    if len(dates) < 2:
        return

    n_trans = len(dates) - 1
    n_strikes = len(strikes)

    # Build change dicts per transition
    put_chg_dicts: list[dict[int, int]] = []
    call_chg_dicts: list[dict[int, int]] = []
    for t in range(n_trans):
        pc: dict[int, int] = {}
        cc: dict[int, int] = {}
        for s in strikes:
            pv = put_oi_list[t + 1].get(s, 0) - put_oi_list[t].get(s, 0)
            cv = call_oi_list[t + 1].get(s, 0) - call_oi_list[t].get(s, 0)
            if pv != 0:
                pc[s] = pv
            if cv != 0:
                cc[s] = cv
        put_chg_dicts.append(pc)
        call_chg_dicts.append(cc)

    # Rank strikes by cumulative absolute change, show top N
    _MAX_BARS = 40
    cum_score: dict[int, int] = defaultdict(int)
    for pc, cc in zip(put_chg_dicts, call_chg_dicts):
        for s, v in pc.items():
            cum_score[s] += abs(v)
        for s, v in cc.items():
            cum_score[s] += abs(v)
    ranked = sorted(cum_score, key=cum_score.get, reverse=True)
    active_strikes = sorted(ranked[:_MAX_BARS])

    if not active_strikes:
        st.info("建玉変動なし")
        return

    show_n = min(4, n_trans)
    start = n_trans - show_n

    titles = []
    for t in range(start, n_trans):
        lbl = f"{_date_label(dates[t])}→{_date_label(dates[t+1])}"
        if mp_list[t + 1]:
            lbl += f"  MP:{mp_list[t+1]:,}"
            if mp_list[t]:
                diff = mp_list[t + 1] - mp_list[t]
                lbl += f"({diff:+,})"
        titles.append(lbl)

    fig = make_subplots(
        rows=show_n, cols=1,
        subplot_titles=titles,
        shared_xaxes=True,
        vertical_spacing=0.08,
    )

    for row_idx, t in enumerate(range(start, n_trans)):
        row = row_idx + 1
        show_legend = (row_idx == 0)
        pc = put_chg_dicts[t]
        cc = call_chg_dicts[t]

        put_y = [pc.get(s, 0) for s in active_strikes]
        call_y = [-cc.get(s, 0) for s in active_strikes]
        call_raw = [cc.get(s, 0) for s in active_strikes]

        # PUT bars (red, upward)
        fig.add_trace(go.Bar(
            x=active_strikes, y=put_y,
            name="PUT",
            marker_color="rgba(220, 60, 60, 0.7)",
            showlegend=show_legend,
            legendgroup="put",
            hovertemplate="行使価格: %{x:,}<br>PUT OI変化: %{y:+,}<extra></extra>",
        ), row=row, col=1)

        # CALL bars (blue, inverted downward)
        fig.add_trace(go.Bar(
            x=active_strikes, y=call_y,
            name="CALL",
            marker_color="rgba(60, 100, 220, 0.7)",
            showlegend=show_legend,
            legendgroup="call",
            hovertemplate="行使価格: %{x:,}<br>CALL OI変化: %{customdata:+,}<extra></extra>",
            customdata=call_raw,
        ), row=row, col=1)

        # Max Pain vertical line
        mp_val = mp_list[t + 1]
        if mp_val:
            fig.add_vline(
                x=mp_val,
                line=dict(color="green", width=2, dash="dash"),
                row=row, col=1,
            )

        # NK225 closing price vertical line
        if nk_lookup:
            nk_close = nk_lookup.get(dates[t + 1])
            if nk_close:
                fig.add_vline(
                    x=nk_close,
                    line=dict(color="black", width=1.5, dash="dot"),
                    row=row, col=1,
                )
                if row_idx == 0:
                    fig.add_annotation(
                        x=nk_close, y=1,
                        yref="y domain",
                        text=f"NK {nk_close:,.0f}",
                        showarrow=False,
                        font=dict(color="black", size=9),
                        bgcolor="rgba(255,255,255,0.8)",
                        yanchor="top",
                    )

    fig.update_layout(
        title=f"建玉変動 (変動ありストライクのみ) - {_format_cm(contract_month)}",
        height=280 * show_n + 60,
        template="plotly_white",
        barmode="relative",
        margin=dict(l=0, r=0, t=60, b=0),
        legend=dict(orientation="h", y=-0.05),
    )

    # Axis formatting
    for i in range(1, show_n + 1):
        yaxis_key = "yaxis" if i == 1 else f"yaxis{i}"
        xaxis_key = "xaxis" if i == 1 else f"xaxis{i}"
        fig.update_layout(**{
            yaxis_key: dict(zeroline=True, zerolinewidth=1, zerolinecolor="black"),
            xaxis_key: dict(tickformat=","),
        })

    st.plotly_chart(fig, use_container_width=True)
    st.caption(
        f"赤=PUT OI変化(上向き) / 青=CALL OI変化(下向き反転) / "
        f"緑破線=Max Pain / 黒点線=NK225終値 / "
        f"表示ストライク: {len(active_strikes)}本 "
        f"({min(active_strikes):,}~{max(active_strikes):,})"
    )


