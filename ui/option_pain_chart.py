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

    st.markdown("---")

    # 3. Per-month pain profile (latest day snapshot)
    for cm in sorted(all_month_rows.keys()):
        rows = all_month_rows[cm]
        if not rows:
            continue
        _render_single_pain(rows, week, cm)


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

    # --- Primary: OI change heatmap ---
    _render_oi_change_heatmap(
        dates, strikes, put_oi_list, call_oi_list, mp_list, selected_cm,
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
    """OI distribution: PUT as negative bars, CALL as positive, overlaid per day."""
    fig = go.Figure()

    n = len(dates)
    # Color gradient: oldest=light, newest=dark
    put_alphas = [0.15 + 0.85 * i / max(n - 1, 1) for i in range(n)]
    call_alphas = [0.15 + 0.85 * i / max(n - 1, 1) for i in range(n)]

    for i, (td, po, co) in enumerate(zip(dates, put_oi_list, call_oi_list)):
        label = _date_label(td)
        pa = put_alphas[i]
        ca = call_alphas[i]
        is_latest = (i == n - 1)

        # PUT OI (negative side)
        put_vals = [-po.get(s, 0) for s in strikes]
        fig.add_trace(go.Bar(
            x=strikes, y=put_vals,
            name=f"PUT {label}",
            marker_color=f"rgba(220, 60, 60, {pa})",
            width=200,
            showlegend=is_latest,
            legendgroup="PUT",
            hovertemplate=f"PUT {label}<br>行使価格: %{{x:,}}<br>OI: %{{customdata:,}}<extra></extra>",
            customdata=[po.get(s, 0) for s in strikes],
        ))

        # CALL OI (positive side)
        call_vals = [co.get(s, 0) for s in strikes]
        fig.add_trace(go.Bar(
            x=strikes, y=call_vals,
            name=f"CALL {label}",
            marker_color=f"rgba(60, 120, 220, {ca})",
            width=200,
            showlegend=is_latest,
            legendgroup="CALL",
            hovertemplate=f"CALL {label}<br>行使価格: %{{x:,}}<br>OI: %{{y:,}}<extra></extra>",
        ))

    # Max Pain vertical lines
    mp_colors = [f"rgba(0, 180, 0, {0.3 + 0.7 * i / max(n - 1, 1)})" for i in range(n)]
    for i, (td, mp) in enumerate(zip(dates, mp_list)):
        if mp is None:
            continue
        is_latest = (i == n - 1)
        fig.add_vline(
            x=mp,
            line_dash="dash" if not is_latest else "solid",
            line_color=mp_colors[i],
            line_width=2 if is_latest else 1,
            annotation_text=f"MP {_date_label(td)}" if is_latest else None,
            annotation_position="top",
            annotation_font_color="green",
        )

    fig.update_layout(
        title=f"建玉分布 (全行使価格) - {_format_cm(contract_month)}",
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
        "濃い色=直近 / 薄い色=過去 / 緑縦線=Max Pain位置 / "
        "上半分=CALL建玉 / 下半分=PUT建玉（反転表示）"
    )


def _render_oi_change_heatmap(
    dates: list[date],
    strikes: list[int],
    put_oi_list: list[dict[int, int]],
    call_oi_list: list[dict[int, int]],
    mp_list: list[int | None],
    contract_month: str,
) -> None:
    """Heatmap of OI changes: X=strike, Y=date transition, color=OI change."""
    if len(dates) < 2:
        return

    n_trans = len(dates) - 1
    n_strikes = len(strikes)

    # Build change matrices [n_trans x n_strikes]
    put_chg = np.zeros((n_trans, n_strikes))
    call_chg = np.zeros((n_trans, n_strikes))

    for t in range(n_trans):
        for s_idx, s in enumerate(strikes):
            put_chg[t, s_idx] = (
                put_oi_list[t + 1].get(s, 0) - put_oi_list[t].get(s, 0)
            )
            call_chg[t, s_idx] = (
                call_oi_list[t + 1].get(s, 0) - call_oi_list[t].get(s, 0)
            )

    # Y-axis labels with Max Pain info
    y_labels = []
    for i in range(n_trans):
        lbl = f"{_date_label(dates[i])}→{_date_label(dates[i+1])}"
        if mp_list[i] and mp_list[i + 1]:
            diff = mp_list[i + 1] - mp_list[i]
            lbl += f"  MP:{mp_list[i+1]:,}({diff:+,})"
        y_labels.append(lbl)

    # Symmetric color range
    max_abs = float(max(
        np.abs(put_chg).max() if put_chg.size else 1,
        np.abs(call_chg).max() if call_chg.size else 1,
        1,
    ))

    # Custom hover text
    put_hover = []
    for t in range(n_trans):
        row = []
        for s_idx, s in enumerate(strikes):
            val = int(put_chg[t, s_idx])
            row.append(
                f"行使価格: {s:,}<br>{y_labels[t]}<br>PUT OI変動: {val:+,}"
            )
        put_hover.append(row)

    call_hover = []
    for t in range(n_trans):
        row = []
        for s_idx, s in enumerate(strikes):
            val = int(call_chg[t, s_idx])
            row.append(
                f"行使価格: {s:,}<br>{y_labels[t]}<br>CALL OI変動: {val:+,}"
            )
        call_hover.append(row)

    fig = make_subplots(
        rows=2, cols=1,
        subplot_titles=["PUT OI 前日比", "CALL OI 前日比"],
        vertical_spacing=0.15,
        shared_xaxes=True,
    )

    fig.add_trace(go.Heatmap(
        z=put_chg.tolist(),
        x=strikes,
        y=y_labels,
        colorscale="RdBu",
        zmid=0,
        zmin=-max_abs,
        zmax=max_abs,
        text=put_hover,
        hoverinfo="text",
        colorbar=dict(title="枚数", len=0.4, y=0.8),
    ), row=1, col=1)

    fig.add_trace(go.Heatmap(
        z=call_chg.tolist(),
        x=strikes,
        y=y_labels,
        colorscale="RdBu",
        zmid=0,
        zmin=-max_abs,
        zmax=max_abs,
        text=call_hover,
        hoverinfo="text",
        colorbar=dict(title="枚数", len=0.4, y=0.2),
    ), row=2, col=1)

    # Max Pain markers (green diamonds at after-MP position per row)
    for subplot_row in [1, 2]:
        mp_x = [mp_list[t + 1] for t in range(n_trans) if mp_list[t + 1]]
        mp_y = [y_labels[t] for t in range(n_trans) if mp_list[t + 1]]
        fig.add_trace(go.Scatter(
            x=mp_x,
            y=mp_y,
            mode="markers",
            marker=dict(
                symbol="diamond",
                size=10,
                color="lime",
                line=dict(color="black", width=1),
            ),
            name="Max Pain",
            showlegend=(subplot_row == 1),
            legendgroup="mp",
            hovertemplate="Max Pain: %{x:,}<extra></extra>",
        ), row=subplot_row, col=1)

    fig.update_layout(
        title=f"OI 前日比ヒートマップ - {_format_cm(contract_month)}",
        height=max(350, 120 * n_trans + 200),
        template="plotly_white",
        margin=dict(l=0, r=0, t=40, b=0),
    )

    fig.update_xaxes(tickformat=",", row=2, col=1)

    st.plotly_chart(fig, use_container_width=True)
    st.caption(
        "青=OI増加 / 赤=OI減少 / 緑◆=Max Pain位置"
    )


# =====================================================================
# Single contract month pain chart
# =====================================================================

def _render_single_pain(
    rows: list[OptionStrikeRow],
    week: WeekDefinition,
    contract_month: str,
) -> None:
    """Compute and render option pain chart for one contract month."""
    latest_date = _find_latest_oi_date(rows)
    if latest_date is None:
        return

    strikes: list[int] = []
    put_oi: dict[int, int] = {}
    call_oi: dict[int, int] = {}

    for row in rows:
        p = row.put_daily_oi.get(latest_date, 0)
        c = row.call_daily_oi.get(latest_date, 0)
        if p > 0 or c > 0:
            strikes.append(row.strike_price)
            put_oi[row.strike_price] = p
            call_oi[row.strike_price] = c

    if len(strikes) < 3:
        return

    strikes.sort()
    total_oi = sum(put_oi.values()) + sum(call_oi.values())
    if total_oi == 0:
        return

    settlement_prices = strikes
    call_pain = []
    put_pain = []
    total_pain = []

    for S in settlement_prices:
        cp = sum(max(0, S - K) * call_oi.get(K, 0) for K in strikes)
        pp = sum(max(0, K - S) * put_oi.get(K, 0) for K in strikes)
        call_pain.append(cp)
        put_pain.append(pp)
        total_pain.append(cp + pp)

    max_pain_idx = total_pain.index(min(total_pain))
    max_pain_strike = settlement_prices[max_pain_idx]

    OKU = 1e8
    MULT = 1000
    scale = MULT / OKU
    call_pain_oku = [v * scale for v in call_pain]
    put_pain_oku = [v * scale for v in put_pain]
    total_pain_oku = [v * scale for v in total_pain]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="CALL Payout",
        x=[f"{s:,}" for s in settlement_prices],
        y=call_pain_oku,
        marker_color="rgba(74, 144, 217, 0.7)",
    ))
    fig.add_trace(go.Bar(
        name="PUT Payout",
        x=[f"{s:,}" for s in settlement_prices],
        y=put_pain_oku,
        marker_color="rgba(217, 74, 74, 0.7)",
    ))
    fig.add_trace(go.Scatter(
        name="Total",
        x=[f"{s:,}" for s in settlement_prices],
        y=total_pain_oku,
        mode="lines+markers",
        line=dict(color="black", width=2),
        marker=dict(size=3),
    ))
    fig.add_vline(x=max_pain_idx, line_dash="dash", line_color="green", line_width=2)
    fig.add_annotation(
        x=f"{max_pain_strike:,}",
        y=max(total_pain_oku) * 0.95,
        text=f"Max Pain: {max_pain_strike:,}",
        showarrow=True, arrowhead=2, arrowcolor="green",
        font=dict(color="green", size=12),
    )

    dow = _DOW_JP[latest_date.weekday()]
    date_str = f"{latest_date.strftime('%m/%d')}({dow})"

    fig.update_layout(
        title=f"Option Pain - {_format_cm(contract_month)}  (建玉基準: {date_str})",
        xaxis_title="SQ決済価格",
        yaxis_title="オプション払出額 (億円)",
        barmode="stack",
        height=420,
        margin=dict(l=0, r=0, t=40, b=0),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )

    st.plotly_chart(fig, use_container_width=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("Max Pain", f"{max_pain_strike:,}")
    with c2:
        st.metric("PUT OI計", f"{sum(put_oi.values()):,}")
    with c3:
        st.metric("CALL OI計", f"{sum(call_oi.values()):,}")

    st.markdown("---")


def _find_latest_oi_date(rows: list[OptionStrikeRow]) -> date | None:
    """Find the latest date with OI data across all strikes."""
    all_dates: set[date] = set()
    for r in rows:
        all_dates.update(r.put_daily_oi.keys())
        all_dates.update(r.call_daily_oi.keys())
    return max(all_dates) if all_dates else None
