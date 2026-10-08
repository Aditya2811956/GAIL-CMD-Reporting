"""
Data validation for extraction results.

Three-level validation:
1. Schema validation (Pydantic)
2. Range validation (sanity checks)
3. Consistency validation (cross-field checks)
"""

import logging
from typing import Any, Dict, List, Optional

from pydantic import ValidationError as PydanticValidationError

try:
    from .schemas import (
        InflationData,
        GDPData,
        GasStatsMonthly,
        GasStatsYearly,
        InterestRateData,
        SharePriceData,
        ValidationResult,
        ValidationError,
    )
except ImportError:
    from generation.schemas import (
        InflationData,
        GDPData,
        GasStatsMonthly,
        GasStatsYearly,
        InterestRateData,
        SharePriceData,
        ValidationResult,
        ValidationError,
    )

logger = logging.getLogger(__name__)

# Valid ranges for sanity checks
VALID_RANGES = {
    "WPI": (-15, 25),        # WPI inflation typically -15% to 25%
    "CPI": (-10, 30),        # CPI inflation typically -10% to 30%
    "IIP": (-50, 50),        # IIP growth can be volatile
    "repo_rate": (3, 12),    # Repo rate typically 3% to 12%
    "mclr_1y": (5, 15),      # MCLR typically 5% to 15%
    "gdp_growth": (-30, 30), # GDP growth -30% to 30%
    "gas_production_monthly": (1000, 10000),  # MMSCM monthly
    "gas_production_yearly": (50, 200),       # MMSCMD yearly
    "market_cap": (50000, 500000),            # Rs Cr
    "share_price": (50, 500),                 # Rs
}


def _check_range(
    value: Optional[float],
    field: str,
    range_key: str,
    result: ValidationResult
) -> None:
    """Check if value is within expected range."""
    if value is None:
        return

    range_tuple = VALID_RANGES.get(range_key)
    if not range_tuple:
        return

    min_val, max_val = range_tuple
    if value < min_val or value > max_val:
        result.add_warning(
            field,
            f"{field}={value} is outside expected range [{min_val}, {max_val}]"
        )


def validate_inflation_data(
    data: Dict[str, Any],
    result: ValidationResult
) -> Optional[InflationData]:
    """Validate inflation extraction result."""
    try:
        validated = InflationData(**data)

        # Range checks
        _check_range(validated.WPI, "WPI", "WPI", result)
        _check_range(validated.CPI, "CPI", "CPI", result)
        _check_range(validated.IIP, "IIP", "IIP", result)

        # Check for missing values
        missing = []
        if validated.WPI is None:
            missing.append("WPI")
        if validated.CPI is None:
            missing.append("CPI")
        if validated.IIP is None:
            missing.append("IIP")

        if missing:
            result.add_warning("inflation", f"Missing values: {', '.join(missing)}")

        return validated

    except PydanticValidationError as e:
        for error in e.errors():
            result.add_error(
                f"inflation.{error['loc'][0]}",
                error['msg']
            )
        return None


def validate_gdp_data(
    data: Dict[str, Any],
    result: ValidationResult
) -> Optional[GDPData]:
    """Validate GDP extraction result."""
    try:
        validated = GDPData(**data)

        # Range check for growth
        _check_range(validated.growth_pct, "gdp_growth", "gdp_growth", result)

        # Sanity check: GDP value should be positive (skip if null)
        if validated.gdp_rs_cr is not None and validated.gdp_rs_cr <= 0:
            result.add_error("gdp.gdp_rs_cr", "GDP value must be positive")

        return validated

    except PydanticValidationError as e:
        for error in e.errors():
            result.add_error(
                f"gdp.{error['loc'][0]}",
                error['msg']
            )
        return None


def validate_gas_stats_monthly(
    data: Dict[str, Any],
    result: ValidationResult
) -> Optional[GasStatsMonthly]:
    """Validate monthly gas stats."""
    try:
        validated = GasStatsMonthly(**data)

        # Range checks
        _check_range(validated.net_production, "gas.net", "gas_production_monthly", result)
        _check_range(validated.lng_import, "gas.lng", "gas_production_monthly", result)

        # Consistency check: total should be approximately net + lng
        expected_total = validated.net_production + validated.lng_import
        if abs(validated.total_consumption - expected_total) > expected_total * 0.1:
            result.add_warning(
                "gas.total",
                f"Total ({validated.total_consumption}) differs from net + lng ({expected_total})"
            )

        return validated

    except PydanticValidationError as e:
        for error in e.errors():
            result.add_error(
                f"gas_monthly.{error['loc'][0]}",
                error['msg']
            )
        return None


def validate_gas_stats_yearly(
    data: Dict[str, Any],
    result: ValidationResult
) -> Optional[GasStatsYearly]:
    """Validate yearly gas stats."""
    try:
        validated = GasStatsYearly(**data)

        # Range checks
        _check_range(validated.net_production, "gas.net", "gas_production_yearly", result)
        _check_range(validated.lng_import, "gas.lng", "gas_production_yearly", result)

        # Consistency check
        expected_total = validated.net_production + validated.lng_import
        if abs(validated.total_consumption - expected_total) > expected_total * 0.1:
            result.add_warning(
                "gas.total",
                f"Total ({validated.total_consumption}) differs from net + lng ({expected_total})"
            )

        return validated

    except PydanticValidationError as e:
        for error in e.errors():
            result.add_error(
                f"gas_yearly.{error['loc'][0]}",
                error['msg']
            )
        return None


def validate_interest_rates(
    data: Dict[str, Any],
    result: ValidationResult
) -> Optional[InterestRateData]:
    """Validate interest rate data."""
    try:
        validated = InterestRateData(**data)

        # Warn if both values are missing (extraction couldn't find data)
        if validated.repo_rate is None and validated.mclr_1y is None:
            result.add_warning(
                "interest_rates",
                "Both repo_rate and mclr_1y are missing; will use historical data"
            )
            return validated

        # Range checks
        _check_range(validated.repo_rate, "repo_rate", "repo_rate", result)
        if validated.mclr_1y:
            _check_range(validated.mclr_1y, "mclr_1y", "mclr_1y", result)

        # Consistency: MCLR should be higher than repo rate
        if validated.mclr_1y and validated.repo_rate and validated.mclr_1y < validated.repo_rate:
            result.add_warning(
                "interest_rates",
                f"MCLR ({validated.mclr_1y}) should not be lower than Repo ({validated.repo_rate})"
            )

        return validated

    except PydanticValidationError as e:
        for error in e.errors():
            result.add_error(
                f"interest_rates.{error['loc'][0]}",
                error['msg']
            )
        return None


def validate_share_price(
    data: Dict[str, Any],
    result: ValidationResult
) -> Optional[SharePriceData]:
    """Validate share price data."""
    try:
        validated = SharePriceData(**data)

        # Range checks
        _check_range(validated.market_cap, "market_cap", "market_cap", result)
        _check_range(validated.share_price, "share_price", "share_price", result)

        return validated

    except PydanticValidationError as e:
        for error in e.errors():
            result.add_error(
                f"share_price.{error['loc'][0]}",
                error['msg']
            )
        return None


def validate_extraction_results(
    results: Dict[str, Any],
    report_month: str,
    strict: bool = False
) -> ValidationResult:
    """
    Validate all extraction results before generation.

    Args:
        results: Dictionary of extraction results keyed by metric name
        report_month: The report month in YYYY-MM format
        strict: If True, warnings become errors

    Returns:
        ValidationResult with valid flag, errors, warnings, and cleaned data
    """
    validation_result = ValidationResult(valid=True, validated_data={})

    # Validate each metric type
    metric_validators = {
        "wpi": lambda d: validate_inflation_data({"month": d.get("month", report_month), "WPI": d.get("WPI")}, validation_result),
        "cpi": lambda d: validate_inflation_data({"month": d.get("month", report_month), "CPI": d.get("CPI")}, validation_result),
        "iip": lambda d: validate_inflation_data({"month": d.get("month", report_month), "IIP": d.get("IIP")}, validation_result),
        "gdp": lambda d: validate_gdp_data(d, validation_result),
        "gas_monthly": lambda d: validate_gas_stats_monthly(d, validation_result),
        "gas_yearly": lambda d: validate_gas_stats_yearly(d, validation_result),
        "gas_stats_monthly": lambda d: validate_gas_stats_monthly(d, validation_result),
        "gas_stats_yearly": lambda d: validate_gas_stats_yearly(d, validation_result),
        "interest_repo": lambda d: validate_interest_rates(d, validation_result),
        "share_price": lambda d: validate_share_price(d, validation_result),
    }

    # Metrics that pass through without strict validation
    passthrough_metrics = {"gdp_forecast", "satat_data", "cgd_data", "cgd_cumulative"}

    validated_data = {}

    for metric_name, extraction_result in results.items():
        # Handle ExtractionResult objects from extraction module
        if hasattr(extraction_result, "data"):
            data = extraction_result.data
            if not extraction_result.success:
                validation_result.add_error(
                    metric_name,
                    f"Extraction failed: {extraction_result.error}"
                )
                continue
        else:
            data = extraction_result

        if data is None:
            validation_result.add_warning(metric_name, "No data extracted")
            continue

        # Check if extraction returned all-null values (effectively no data)
        id_fields = {"month", "year", "quarter", "period"}
        value_fields = {k: v for k, v in data.items() if k not in id_fields} if isinstance(data, dict) else {}
        if isinstance(data, dict) and value_fields and all(v is None for v in value_fields.values()):
            validation_result.add_warning(metric_name, "Extraction returned all null values")
            continue

        # Check for passthrough metrics first (no strict validation)
        if metric_name in passthrough_metrics:
            validated_data[metric_name] = data
            continue

        # Find matching validator — exact key match first, then substring
        validator = metric_validators.get(metric_name)
        if not validator:
            for key, val_func in metric_validators.items():
                if key in metric_name.lower():
                    validator = val_func
                    break

        if validator:
            validated = validator(data)
            if validated:
                validated_data[metric_name] = validated.model_dump()
        else:
            # No specific validator, store raw data
            validated_data[metric_name] = data

    # Update validation result
    validation_result.validated_data = validated_data

    # In strict mode, warnings become errors
    if strict and validation_result.has_warnings:
        for warning in validation_result.warnings:
            validation_result.add_error(warning.field, warning.message)
        validation_result.warnings = []

    # Set valid flag
    validation_result.valid = not validation_result.has_errors

    return validation_result


def generate_validation_report(result: ValidationResult) -> str:
    """
    Generate a human-readable validation report.

    Args:
        result: ValidationResult from validate_extraction_results

    Returns:
        Formatted string report
    """
    lines = []
    lines.append("=" * 60)
    lines.append("EXTRACTION VALIDATION REPORT")
    lines.append("=" * 60)

    if result.valid:
        lines.append("\n[PASS] All validations passed")
    else:
        lines.append("\n[FAIL] Validation errors found")

    if result.errors:
        lines.append("\n--- ERRORS ---")
        for error in result.errors:
            lines.append(f"  [ERROR] {error.field}: {error.message}")

    if result.warnings:
        lines.append("\n--- WARNINGS ---")
        for warning in result.warnings:
            lines.append(f"  [WARN] {warning.field}: {warning.message}")

    if result.validated_data:
        lines.append("\n--- VALIDATED METRICS ---")
        for metric, data in result.validated_data.items():
            lines.append(f"  {metric}: {data}")

    lines.append("\n" + "=" * 60)

    return "\n".join(lines)
