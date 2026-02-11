"""Parse daily OI balance Excel (open_interest_e.xlsx Attachment1).

Sheet structure (Attachment1 = NK225 Options per strike):
  Row 2: Report date (datetime)
  Row 5: Section headers -- PUT left, CALL right
  Row 6: Column headers
  Row 7+: Data rows

  PUT cols: A=contract code, B=volume, C=current_oi, D=change, E=prev_oi
  CALL cols: G=contract code, H=volume, I=current_oi, J=change, K=prev_oi

  Contract code format: "NIKKEI 225 P2603-38000"
  PUT and CALL sides are NOT row-aligned.
  Blocks separated by "Total for Contract Month" rows.
"""
from __future__ import annotations

import io
import re
import logging
from datetime import date, datetime

import openpyxl

from models import DailyOIBalance

logger = logging.getLogger(__name__)

_CONTRACT_RE = re.compile(r'NIKKEI 225 ([PC])(\d{4})-(\d+)')


def parse_daily_oi_excel(content: bytes) -> list[DailyOIBalance]:
    """Parse Attachment1 sheet of daily OI balance Excel.

    Returns all PUT and CALL records across all contract months.
    """
    wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)

    # Use second sheet (Attachment1); index-based since JP version has garbled names
    if len(wb.worksheets) < 2:
        logger.warning("Daily OI Excel has fewer than 2 sheets")
        wb.close()
        return []

    ws = wb.worksheets[1]
    report_date = _extract_date(ws)
    results: list[DailyOIBalance] = []

    for row_idx in range(7, ws.max_row + 1):
        # PUT side (col A=1)
        put_code = ws.cell(row=row_idx, column=1).value
        if put_code:
            m = _CONTRACT_RE.search(str(put_code))
            if m and m.group(1) == "P":
                results.append(DailyOIBalance(
                    report_date=report_date,
                    contract_month=m.group(2),
                    option_type="PUT",
                    strike_price=int(m.group(3)),
                    trading_volume=_safe_int(ws.cell(row=row_idx, column=2).value),
                    current_oi=_safe_int(ws.cell(row=row_idx, column=3).value),
                    net_change=_safe_int(ws.cell(row=row_idx, column=4).value),
                    previous_oi=_safe_int(ws.cell(row=row_idx, column=5).value),
                ))

        # CALL side (col G=7)
        call_code = ws.cell(row=row_idx, column=7).value
        if call_code:
            m = _CONTRACT_RE.search(str(call_code))
            if m and m.group(1) == "C":
                results.append(DailyOIBalance(
                    report_date=report_date,
                    contract_month=m.group(2),
                    option_type="CALL",
                    strike_price=int(m.group(3)),
                    trading_volume=_safe_int(ws.cell(row=row_idx, column=8).value),
                    current_oi=_safe_int(ws.cell(row=row_idx, column=9).value),
                    net_change=_safe_int(ws.cell(row=row_idx, column=10).value),
                    previous_oi=_safe_int(ws.cell(row=row_idx, column=11).value),
                ))

    wb.close()
    return results


def _extract_date(ws) -> date:
    """Extract report date from row 2, column A."""
    val = ws.cell(row=2, column=1).value
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    # Fallback: extract digits
    digits = re.findall(r'\d+', str(val))
    if len(digits) >= 3:
        return date(int(digits[0]), int(digits[1]), int(digits[2]))
    raise ValueError(f"Cannot extract date from daily OI Excel: {val}")


def _safe_int(val) -> int:
    """Convert cell value to int, defaulting to 0."""
    if val is None:
        return 0
    try:
        return int(val)
    except (ValueError, TypeError):
        return 0
