"""
Scrapers package for automated macroeconomic data extraction.
"""

from .mospi_scraper import fetch_cpi_data, fetch_iip_data, fetch_gdp_data
from .wpi_scraper import fetch_wpi_data

__all__ = [
    "fetch_cpi_data",
    "fetch_iip_data",
    "fetch_gdp_data",
    "fetch_wpi_data",
]
