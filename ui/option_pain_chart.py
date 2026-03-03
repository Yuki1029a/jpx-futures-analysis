"""Option Pain (Max Pain) chart rendering.

Computes the settlement price that minimizes total option payout.
Includes:
- Max Pain time series vs NK225
- OI change heatmap showing which strikes drive Max Pain movement
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

    # 2. OI change heatmap + Max Pain driver analysis
    _render_pain_driver_section(all_month_rows, week)

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
# OI Change Heatmap + Max Pain Driver Analysis
# =====================================================================

def _render_pain_driver_section(
    all_month_rows: dict[str, list[OptionStrikeRow]],
    week: WeekDefinition,
) -> None:
    """Render OI change heatmap showing which strikes drive Max Pain."""
    st.subheader("Max Pain 変動要因 (OI変動ヒートマップ)")

    sorted_cms = sorted(all_month_rows.keys())
    if not sorted_cms:
        return

    selected_cm = st.selectbox(
        "限月選択", sorted_cms,
        format_func=_format_cm,
        key="pain_driver_cm",
    )

    rows = all_month_rows.get(selected_cm, [])
    if not rows:
        st.info("データなし")
        return

    # Collect dates with OI data
    all_dates: set[date] = set()
    for r in rows:
        all_dates.update(r.put_daily_oi.keys())
        all_dates.update(r.call_daily_oi.keys())

    if not all_dates:
        st.info("建玉データなし")
        return

    sorted_dates = sorted(all_dates)[-5:]
    if len(sorted_dates) < 2:
        st.info("日数不足（2日以上必要）")
        return

    # Build strike lookup
    strike_lookup: dict[int, OptionStrikeRow] = {r.strike_price: r for r in rows}

    # Determine strike range: filter to strikes with meaningful OI
    active_strikes: set[int] = set()
    for r in rows:
        for td in sorted_dates:
            total = r.put_daily_oi.get(td, 0) + r.call_daily_oi.get(td, 0)
            if total >= 100:  # filter noise
                active_strikes.add(r.strike_price)

    if len(active_strikes) < 3:
        st.info("有効行使価格が不足")
        return

    # Filter to +/- 6000 from OI-weighted center
    latest = sorted_dates[-1]
    oi_weights = {
        s: (strike_lookup[s].put_daily_oi.get(latest, 0) +
            strike_lookup[s].call_daily_oi.get(latest, 0))
        for s in active_strikes
    }
    total_w = sum(oi_weights.values())
    center = (sum(s * w for s, w in oi_weights.items()) / total_w
              if total_w > 0 else sorted(active_strikes)[len(active_strikes) // 2])

    filtered_strikes = sorted(s for s in active_strikes if abs(s - center) <= 6000)
    if len(filtered_strikes) < 3:
        filtered_strikes = sorted(active_strikes)

    # Use 500-yen-step standard strikes only for cleaner heatmap
    standard_strikes = [s for s in filtered_strikes if s % 500 == 0]
    if len(standard_strikes) >= 5:
        filtered_strikes = standard_strikes

    date_labels = [_date_label(td) for td in sorted_dates]

    # Build OI change matrices (Y=strikes ascending, X=dates)
    put_chg = []
    call_chg = []
    for strike in filtered_strikes:
        r = strike_lookup.get(strike)
        p_row = []
        c_row = []
        for td in sorted_dates:
            p_row.append(r.put_daily_oi_change.get(td, 0) if r else 0)
            c_row.append(r.call_daily_oi_change.get(td, 0) if r else 0)
        put_chg.append(p_row)
        call_chg.append(c_row)

    put_chg_arr = np.array(put_chg)
    call_chg_arr = np.array(call_chg)

    # Compute Max Pain per day
    max_pain_per_day = []
    for td in sorted_dates:
        sm: dict[int, dict[str, int]] = defaultdict(lambda: {"CALL": 0, "PUT": 0})
        for r in rows:
            p = r.put_daily_oi.get(td, 0)
            c = r.call_daily_oi.get(td, 0)
            if p > 0:
                sm[r.strike_price]["PUT"] = p
            if c > 0:
                sm[r.strike_price]["CALL"] = c
        mp = _calc_max_pain(sm)
        max_pain_per_day.append(mp)

    # Symmetric color range
    abs_max = max(
        np.abs(put_chg_arr).max() if put_chg_arr.size else 0,
        np.abs(call_chg_arr).max() if call_chg_arr.size else 0,
        1,
    )

    strike_labels = [f"{s:,}" for s in filtered_strikes]

    # Build figure: PUT heatmap (left) | CALL heatmap (right)
    fig = make_subplots(
        rows=1, cols=2,
        subplot_titles=["PUT OI 日次変動", "CALL OI 日次変動"],
        shared_yaxes=True,
        horizontal_spacing=0.08,
    )

    # PUT heatmap
    fig.add_trace(go.Heatmap(
        x=date_labels,
        y=strike_labels,
        z=put_chg_arr.tolist(),
        colorscale="RdBu_r",  # red=negative(unwind), blue=positive(build)
        zmid=0,
        zmin=-abs_max,
        zmax=abs_max,
        colorbar=dict(title="枚数", x=0.46, len=0.8),
        hovertemplate="行使価格: %{y}<br>日付: %{x}<br>PUT OI変動: %{z:+,}<extra></extra>",
    ), row=1, col=1)

    # CALL heatmap
    fig.add_trace(go.Heatmap(
        x=date_labels,
        y=strike_labels,
        z=call_chg_arr.tolist(),
        colorscale="RdBu_r",
        zmid=0,
        zmin=-abs_max,
        zmax=abs_max,
        colorbar=dict(title="枚数", x=1.02, len=0.8),
        hovertemplate="行使価格: %{y}<br>日付: %{x}<br>CALL OI変動: %{z:+,}<extra></extra>",
    ), row=1, col=2)

    # Max Pain overlay on both panels
    mp_labels = [f"{mp:,}" if mp else "" for mp in max_pain_per_day]
    for col_idx in [1, 2]:
        fig.add_trace(go.Scatter(
            x=date_labels,
            y=[f"{mp:,}" if mp else None for mp in max_pain_per_day],
            mode="lines+markers+text",
            marker=dict(size=10, color="lime", symbol="diamond",
                        line=dict(color="black", width=1.5)),
            line=dict(color="lime", width=2),
            text=mp_labels,
            textposition="middle right",
            textfont=dict(size=10, color="green"),
            name="Max Pain" if col_idx == 1 else None,
            showlegend=(col_idx == 1),
            hovertemplate="Max Pain: %{y}<extra></extra>",
        ), row=1, col=col_idx)

    fig.update_layout(
        title=f"OI変動 vs Max Pain推移 - {_format_cm(selected_cm)}",
        height=max(len(filtered_strikes) * 22 + 120, 450),
        margin=dict(l=0, r=0, t=60, b=0),
        template="plotly_white",
    )
    fig.update_yaxes(autorange="reversed", row=1, col=1)
    fig.update_yaxes(autorange="reversed", row=1, col=2)

    st.plotly_chart(fig, use_container_width=True)

    st.caption(
        "青=OI増加（新規建て）/ 赤=OI減少（解消）/ "
        "緑◆=Max Pain / ヒートマップはPUT・CALL別の日次OI変動量"
    )

    # Top OI changes driving Max Pain shifts
    _render_pain_driver_table(
        rows, filtered_strikes, sorted_dates, max_pain_per_day,
    )


def _render_pain_driver_table(
    rows: list[OptionStrikeRow],
    strikes: list[int],
    dates: list[date],
    max_pain_per_day: list[int | None],
) -> None:
    """Show table of top OI changes that likely drove Max Pain shifts."""
    strike_lookup = {r.strike_price: r for r in rows}

    for i in range(1, len(dates)):
        prev_mp = max_pain_per_day[i - 1]
        curr_mp = max_pain_per_day[i]
        if prev_mp is None or curr_mp is None:
            continue

        mp_change = curr_mp - prev_mp
        if mp_change == 0:
            continue

        td = dates[i]
        direction = "上昇" if mp_change > 0 else "下落"

        # Collect OI changes for this day
        changes = []
        for s in strikes:
            r = strike_lookup.get(s)
            if not r:
                continue
            p_chg = r.put_daily_oi_change.get(td, 0)
            c_chg = r.call_daily_oi_change.get(td, 0)
            if abs(p_chg) >= 50 or abs(c_chg) >= 50:
                changes.append((s, p_chg, c_chg))

        if not changes:
            continue

        # Sort by absolute total change
        changes.sort(key=lambda x: abs(x[1]) + abs(x[2]), reverse=True)

        with st.expander(
            f"{_date_label(td)}: Max Pain {prev_mp:,} → {curr_mp:,} "
            f"({mp_change:+,} {direction})",
            expanded=(i == len(dates) - 1),  # latest day expanded
        ):
            rows_display = []
            for s, p_chg, c_chg in changes[:8]:
                impact = ""
                # PUT OI increase below MP → pulls MP down
                # PUT OI increase above MP → pulls MP up
                # CALL OI increase above MP → pulls MP down
                # CALL OI increase below MP → pulls MP up
                if p_chg > 0 and s < curr_mp:
                    impact = "MP引下げ"
                elif p_chg > 0 and s > curr_mp:
                    impact = "MP引上げ"
                elif p_chg < 0 and s < curr_mp:
                    impact = "MP引上げ"
                elif c_chg > 0 and s > curr_mp:
                    impact = "MP引下げ"
                elif c_chg > 0 and s < curr_mp:
                    impact = "MP引上げ"
                elif c_chg < 0 and s > curr_mp:
                    impact = "MP引上げ"
                elif c_chg < 0 and s < curr_mp:
                    impact = "MP引下げ"

                rows_display.append({
                    "行使価格": f"{s:,}",
                    "PUT OI変動": f"{p_chg:+,}" if p_chg else "-",
                    "CALL OI変動": f"{c_chg:+,}" if c_chg else "-",
                    "影響方向": impact,
                })

            import pandas as pd
            st.dataframe(
                pd.DataFrame(rows_display),
                use_container_width=True,
                hide_index=True,
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
