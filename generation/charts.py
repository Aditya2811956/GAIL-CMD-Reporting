"""
Chart helpers for CMD reports.

The generation approach has moved from matplotlib PNG images to native OOXML
chart data updates (see chart_updater.py). This module re-exports the relevant
functions for backward compatibility.
"""

from .chart_updater import (
    make_inflation_update,
    make_interest_rates_update,
    make_gas_yearly_update,
    make_gas_monthly_update,
    make_share_price_update,
    update_docx_charts,
)

__all__ = [
    "make_inflation_update",
    "make_interest_rates_update",
    "make_gas_yearly_update",
    "make_gas_monthly_update",
    "make_share_price_update",
    "update_docx_charts",
]
