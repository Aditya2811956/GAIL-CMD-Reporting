"""
Utility functions for extraction.

Period formatting and other helpers.
"""

from datetime import datetime


def get_month_period(report_month: str) -> str:
    """Convert YYYY-MM to readable period like 'November 2024'."""
    dt = datetime.strptime(report_month, "%Y-%m")
    return dt.strftime("%B %Y")


def get_quarter_period(report_month: str) -> str:
    """Convert YYYY-MM to quarter like 'Q2 of FY 2024-25'."""
    dt = datetime.strptime(report_month, "%Y-%m")
    month = dt.month
    year = dt.year

    if month in [4, 5, 6]:
        q = 1
        fy_start = year
    elif month in [7, 8, 9]:
        q = 2
        fy_start = year
    elif month in [10, 11, 12]:
        q = 3
        fy_start = year
    else:  # 1, 2, 3
        q = 4
        fy_start = year - 1

    fy_end = fy_start + 1
    return f"Q{q} of FY {fy_start}-{str(fy_end)[-2:]}"


def get_year_period(report_month: str) -> str:
    """Convert YYYY-MM to fiscal year like 'FY 2024-25'."""
    dt = datetime.strptime(report_month, "%Y-%m")
    month = dt.month
    year = dt.year

    if month >= 4:
        fy_start = year
    else:
        fy_start = year - 1

    fy_end = fy_start + 1
    return f"FY {fy_start}-{str(fy_end)[-2:]}"


def get_period(report_month: str, period_type: str) -> str:
    """Get formatted period string based on type."""
    if period_type == "month":
        return get_month_period(report_month)
    elif period_type == "quarter":
        return get_quarter_period(report_month)
    elif period_type == "year":
        return get_year_period(report_month)
    else:
        return report_month
