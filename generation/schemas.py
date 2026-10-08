"""
Pydantic models for validated generation data.

These models ensure type safety and provide validation
for data coming from the extraction module.
"""

from datetime import datetime
from typing import Any, Dict, List, Optional, Union
from pydantic import BaseModel, Field, field_validator


class InflationData(BaseModel):
    """Monthly inflation metrics (WPI, CPI, IIP)."""
    month: str = Field(..., description="Month in YYYY-MM format")
    WPI: Optional[float] = Field(None, description="Wholesale Price Index (%)")
    CPI: Optional[float] = Field(None, description="Consumer Price Index (%)")
    IIP: Optional[float] = Field(None, description="Index of Industrial Production (%)")

    @field_validator("month")
    @classmethod
    def validate_month_format(cls, v: str) -> str:
        try:
            datetime.strptime(v, "%Y-%m")
        except ValueError:
            raise ValueError(f"month must be in YYYY-MM format, got: {v}")
        return v


class GDPData(BaseModel):
    """Quarterly GDP data."""
    quarter: str = Field(..., description="Quarter label (e.g., Q1-FY25)")
    gdp_rs_cr: Optional[int] = Field(None, description="GDP in Rs Crores")
    gdp_usd_bn: Optional[float] = Field(None, description="GDP in USD Billions")
    growth_pct: Optional[float] = Field(None, description="Y-o-Y growth percentage")

    @field_validator("quarter")
    @classmethod
    def validate_quarter_format(cls, v: str) -> str:
        import re
        if not re.match(r"Q[1-4]-FY\d{2}", v):
            raise ValueError(f"quarter must be in Q1-FY25 format, got: {v}")
        return v


class GDPForecastData(BaseModel):
    """GDP forecast from various agencies."""
    agency: str = Field(..., description="Forecasting agency (RBI, World Bank, IMF)")
    forecast: str = Field(..., description="Forecast value (e.g., '7.2% (FY25)')")
    release_date: Optional[str] = Field(None, description="When forecast was released")


class GasStatsMonthly(BaseModel):
    """Monthly natural gas statistics."""
    month: str = Field(..., description="Month in YYYY-MM format")
    net_production: float = Field(..., alias="net", description="Net production (MMSCM)")
    lng_import: float = Field(..., alias="lng", description="LNG import (MMSCM)")
    total_consumption: float = Field(..., alias="total", description="Total consumption (MMSCM)")

    class Config:
        populate_by_name = True


class GasStatsYearly(BaseModel):
    """Yearly natural gas statistics."""
    year: str = Field(..., description="Fiscal year (e.g., 2023-24)")
    net_production: float = Field(..., alias="net", description="Net production (MMSCMD)")
    lng_import: float = Field(..., alias="lng", description="LNG import (MMSCMD)")
    total_consumption: float = Field(..., alias="total", description="Total consumption (MMSCMD)")

    class Config:
        populate_by_name = True


class InterestRateData(BaseModel):
    """Monthly interest rate data."""
    month: str = Field(..., description="Month in YYYY-MM format")
    repo_rate: Optional[float] = Field(None, alias="Repo", description="RBI Repo Rate (%)")
    mclr_1y: Optional[float] = Field(None, alias="MCLR1Y", description="MCLR 1-Year Rate (%)")

    class Config:
        populate_by_name = True


class SharePriceData(BaseModel):
    """Monthly share price and market cap data."""
    month: str = Field(..., description="Month in YYYY-MM format")
    market_cap: float = Field(..., alias="Market Cap (Rs Cr.)", description="Market Cap in Rs Cr")
    share_price: float = Field(..., alias="GAIL Share", description="GAIL Share Price (Rs)")

    class Config:
        populate_by_name = True


class SatatData(BaseModel):
    """SATAT scheme data."""
    eois_received: Optional[int] = Field(None, description="EOIs Received (No.)")
    lois_issued: Optional[int] = Field(None, description="LOIs Issued (No.)")
    financial_closure: Optional[int] = Field(None, description="Financial Closure count")
    cbg_plants_commissioned: Optional[int] = Field(None, description="CBG Plants commissioned")
    cumulative_cbg_sale_ros: Optional[str] = Field(None, description="Cumulative CBG sale from ROs")
    cbg_sale_current_month: Optional[float] = Field(None, description="CBG sale current month (Tonnes)")
    cbg_sale_fy: Optional[float] = Field(None, description="CBG sale current FY (Tonnes)")


class CGDEntityData(BaseModel):
    """CGD entity data row."""
    entity: str = Field(..., description="CGD entity name")
    target_dpng: Optional[int] = Field(None, description="Internal Target (DPNG) Nos")
    actual_dpng: Optional[int] = Field(None, description="Actual (DPNG) Nos")
    target_cng: Optional[int] = Field(None, description="Internal Target (CNG) Nos")
    actual_cng: Optional[int] = Field(None, description="Actual (CNG) Nos")
    industrial_png: Optional[int] = Field(None, description="Industrial PNG (Nos)")
    commercial_png: Optional[int] = Field(None, description="Commercial PNG (Nos)")


class CGDData(BaseModel):
    """CGD entities data."""
    entities: List[CGDEntityData] = Field(default_factory=list)


class ValidationError(BaseModel):
    """Single validation error."""
    field: str
    message: str
    severity: str = Field(default="error", description="error or warning")


class ValidationResult(BaseModel):
    """Result of data validation."""
    valid: bool = Field(..., description="Whether validation passed")
    errors: List[ValidationError] = Field(default_factory=list, description="Validation errors")
    warnings: List[ValidationError] = Field(default_factory=list, description="Validation warnings")
    validated_data: Optional[Dict[str, Any]] = Field(None, description="Cleaned/validated data")

    @property
    def has_errors(self) -> bool:
        return len(self.errors) > 0

    @property
    def has_warnings(self) -> bool:
        return len(self.warnings) > 0

    def add_error(self, field: str, message: str) -> None:
        self.errors.append(ValidationError(field=field, message=message, severity="error"))

    def add_warning(self, field: str, message: str) -> None:
        self.warnings.append(ValidationError(field=field, message=message, severity="warning"))


class GeneratedReportMetadata(BaseModel):
    """Metadata for a generated report."""
    report_month: str
    template_type: str = Field(default="12", description="6 or 12 month template")
    generated_at: datetime = Field(default_factory=datetime.utcnow)
    validation_result: Optional[ValidationResult] = None
    source_documents: List[str] = Field(default_factory=list)
    extraction_results: Dict[str, Any] = Field(default_factory=dict)
