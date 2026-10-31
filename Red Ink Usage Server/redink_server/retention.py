"""Deterministic retention boundary; no application-specific classification."""
from __future__ import annotations

import calendar
import re
from datetime import date


def validate_cutoff_date(value: str) -> str:
    if not isinstance(value, str) or (value and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", value)):
        raise ValueError("retention.cutoff_date must be blank or YYYY-MM-DD.")
    if value:
        try:
            date.fromisoformat(value)
        except ValueError as exc:
            raise ValueError("retention.cutoff_date must be a valid calendar date.") from exc
    return value


def effective_cutoff(configured: str, keep_months: int, today: date) -> str:
    """Latest of the explicit date and calendar-month rolling boundary."""
    cutoff = validate_cutoff_date(configured)
    if keep_months:
        index = today.year * 12 + today.month - 1 - keep_months
        year, month_zero = divmod(index, 12)
        if year < 1:
            rolling = "0001-01-01"
        else:
            month = month_zero + 1
            rolling = date(year, month, min(today.day,calendar.monthrange(year, month)[1])).isoformat()
        cutoff = max(cutoff, rolling)
    return cutoff
