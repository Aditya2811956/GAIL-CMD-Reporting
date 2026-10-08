"""
Extraction module for CMD Report Automation.

Table-first metric extraction from ChromaDB & Web Scrapers.
"""

from .metrics import METRIC_CONFIGS, MetricConfig
from .utils import get_period, get_month_period, get_quarter_period, get_year_period
from .web_scrapers import (
    fetch_repo_rate,
    fetch_mclr_1y,
    fetch_cpi,
    fetch_iip,
    fetch_gdp,
    fetch_wpi,
    scrape_all_macro_metrics,
)

try:
    from .service import MetricExtractor, ExtractionResult
except ImportError:
    MetricExtractor = None
    ExtractionResult = None

__all__ = [
    "MetricExtractor",
    "ExtractionResult",
    "METRIC_CONFIGS",
    "MetricConfig",
    "get_period",
    "get_month_period",
    "get_quarter_period",
    "get_year_period",
    "fetch_repo_rate",
    "fetch_mclr_1y",
    "fetch_cpi",
    "fetch_iip",
    "fetch_gdp",
    "fetch_wpi",
    "scrape_all_macro_metrics",
]
