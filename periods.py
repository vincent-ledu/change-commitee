from __future__ import annotations
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta, MO


def week_bounds_splus1(ref_date: datetime) -> tuple[datetime, datetime, datetime]:
    """Return (monday_current, monday_next, sunday_next_end_of_day) for S+1."""
    monday_current = (ref_date + relativedelta(weekday=MO(-1))).replace(hour=0, minute=0, second=0, microsecond=0)
    monday_next = monday_current + timedelta(weeks=1)
    # End of next Sunday = start of the following Monday - 1 microsecond
    sunday_next = (monday_next + timedelta(days=7)) - timedelta(microseconds=1)
    return monday_current, monday_next, sunday_next


def week_bounds_sminus1(ref_date: datetime) -> tuple[datetime, datetime]:
    """Return (monday_prev, sunday_prev_end_of_day) for S-1."""
    monday_current = (ref_date + relativedelta(weekday=MO(-1))).replace(hour=0, minute=0, second=0, microsecond=0)
    monday_prev = monday_current - timedelta(weeks=1)
    sunday_prev = (monday_prev + timedelta(days=7)) - timedelta(microseconds=1)
    return monday_prev, sunday_prev


def week_bounds_current(ref_date: datetime) -> tuple[datetime, datetime]:
    """Return (monday_current, sunday_current_end_of_day) for the current week (S)."""
    monday_current = (ref_date + relativedelta(weekday=MO(-1))).replace(hour=0, minute=0, second=0, microsecond=0)
    sunday_current = (monday_current + timedelta(days=7)) - timedelta(microseconds=1)
    return monday_current, sunday_current
