"""Shared interval math for the WFM workbook generator and tests."""

from __future__ import annotations

from datetime import datetime
from typing import Iterable


def interval_productive_fraction(
    shift_start: datetime,
    shift_end: datetime,
    interval_start: datetime,
    interval_end: datetime,
) -> float:
    """Return the fraction of an interval covered by a shift window."""
    if shift_end <= shift_start or interval_end <= interval_start:
        return 0.0

    overlap_start = max(shift_start, interval_start)
    overlap_end = min(shift_end, interval_end)
    overlap_seconds = max(0.0, (overlap_end - overlap_start).total_seconds())
    interval_seconds = (interval_end - interval_start).total_seconds()

    if interval_seconds <= 0:
        return 0.0

    return overlap_seconds / interval_seconds


def row_interval_value(
    *,
    employee_id: str | None,
    activity_code: str | None,
    shift_start: datetime | None,
    shift_end: datetime | None,
    interval_start: datetime,
    interval_end: datetime,
    productive_codes: Iterable[str],
    nonproductive_codes: Iterable[str],
    inactive_codes: Iterable[str],
) -> float:
    """Return numeric staffed coverage for one roster row and interval."""
    if not employee_id or not activity_code or not shift_start or not shift_end:
        return 0.0

    code = activity_code.strip().upper()
    productive = {value.upper() for value in productive_codes}
    nonproductive = {value.upper() for value in nonproductive_codes}
    inactive = {value.upper() for value in inactive_codes}

    if not code or code in nonproductive or code in inactive or code not in productive:
        return 0.0

    return interval_productive_fraction(
        shift_start,
        shift_end,
        interval_start,
        interval_end,
    )
