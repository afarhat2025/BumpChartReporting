"""Utility functions for date and format handling."""

from datetime import datetime


def compute_effective_month_start(part_date) -> tuple:
    """
    Compute the effective month start date for a part.
    
    Args:
        part_date: Date from part object
        
    Returns:
        Tuple of (datetime_object, iso_format_string)
    """
    # Parse date
    try:
        import pandas as pd
        ts = pd.to_datetime(part_date)
        if isinstance(ts, pd.Timestamp):
            if ts.tzinfo is not None:
                ts = ts.tz_convert(None)
            effective_date = ts.to_pydatetime()
        elif isinstance(ts, datetime):
            effective_date = ts.replace(tzinfo=None)
        else:
            effective_date = datetime.min
    except Exception:
        effective_date = datetime.min
    
    if effective_date == datetime.min:
        effective_date = datetime.today()
    
    month_start = effective_date.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    return month_start, month_start.strftime("%Y-%m-%dT%H:%M:%SZ")


def format_price(value: float) -> str:
    """Format a value as currency."""
    try:
        return f"{float(value):.2f}"
    except Exception:
        return str(value)
