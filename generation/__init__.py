"""
Generation module for CMD report automation.

Handles:
- Data validation
- Chart generation
- Table updates
- DOCX report generation
"""

from .schemas import (
    InflationData,
    GDPData,
    GDPForecastData,
    GasStatsMonthly,
    GasStatsYearly,
    InterestRateData,
    SharePriceData,
    SatatData,
    CGDData,
    CGDEntityData,
    ValidationResult,
    ValidationError,
    GeneratedReportMetadata,
)

from .validation import (
    validate_extraction_results,
    generate_validation_report,
    validate_inflation_data,
    validate_gdp_data,
    validate_gas_stats_monthly,
    validate_gas_stats_yearly,
    validate_interest_rates,
    validate_share_price,
)

from .chart_updater import (
    update_docx_charts,
    make_inflation_update,
    make_interest_rates_update,
    make_gas_yearly_update,
    make_gas_monthly_update,
    make_share_price_update,
)

from .tables import (
    snapshot_table_styles,
    restore_table_styles,
    force_black_grid,
    markdown_to_docx_table,
    update_table_from_markdown,
    get_table_markdown,
    add_update_note_after_table,
)

from .table_updater import (
    TableUpdater,
    update_gdp_table,
    update_table_from_chromadb,
)

from .service import (
    ReportGenerator,
    GenerationConfig,
    GenerationResult,
    HistoricalDataManager,
    generate_report,
)

__all__ = [
    # Schemas
    "InflationData",
    "GDPData",
    "GDPForecastData",
    "GasStatsMonthly",
    "GasStatsYearly",
    "InterestRateData",
    "SharePriceData",
    "SatatData",
    "CGDData",
    "CGDEntityData",
    "ValidationResult",
    "ValidationError",
    "GeneratedReportMetadata",
    # Validation
    "validate_extraction_results",
    "generate_validation_report",
    "validate_inflation_data",
    "validate_gdp_data",
    "validate_gas_stats_monthly",
    "validate_gas_stats_yearly",
    "validate_interest_rates",
    "validate_share_price",
    # Charts (native OOXML updater)
    "update_docx_charts",
    "make_inflation_update",
    "make_interest_rates_update",
    "make_gas_yearly_update",
    "make_gas_monthly_update",
    "make_share_price_update",
    # Tables
    "snapshot_table_styles",
    "restore_table_styles",
    "force_black_grid",
    "markdown_to_docx_table",
    "update_table_from_markdown",
    "get_table_markdown",
    "add_update_note_after_table",
    # Table Updater
    "TableUpdater",
    "update_gdp_table",
    "update_table_from_chromadb",
    # Service
    "ReportGenerator",
    "GenerationConfig",
    "GenerationResult",
    "HistoricalDataManager",
    "generate_report",
]
