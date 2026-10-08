"""
Metric configurations for extraction.

Defines all metrics that can be extracted from documents.
"""

from dataclasses import dataclass
from typing import Any, Dict, Optional


@dataclass
class MetricConfig:
    """Configuration for extracting a specific metric."""
    name: str
    metric_type: str  # Maps to ingestion metric_types tag
    query: str  # Semantic query to find relevant tables
    extraction_prompt: str  # Prompt template for LLM extraction
    output_schema: Dict[str, Any]  # Expected JSON structure
    period_type: str = "month"  # "month", "quarter", "year"
    document_hint: Optional[str] = None  # Optional document name pattern
    chunk_type_filter: Optional[str] = None  # Filter by chunk_type (e.g. "table", "row")
    bundle_chunks: bool = False  # Concatenate all retrieved chunks into one before LLM call
    n_results_override: Optional[int] = None  # Override n_results for retrieval (e.g. for row-based metrics)
    append_period_to_query: bool = False  # Append the data period to the query at extraction time


# Define all metrics we need to extract
METRIC_CONFIGS: Dict[str, MetricConfig] = {
    "wpi": MetricConfig(
        name="Wholesale Price Index",
        metric_type="inflation_wpi",
        query="annual rate of inflation WPI All India Wholesale Price Index provisional percent",
        document_hint="wpi",
        chunk_type_filter="paragraph",
        extraction_prompt="""Extract the overall WPI (All India Wholesale Price Index) headline inflation rate from this content.
The report period is {period} but WPI data may lag by 1-2 months — extract the LATEST available month.

Content:
{table}

Instructions:
- Look for a sentence like "The annual rate of inflation based on All India Wholesale Price Index (WPI) number is X.XX% (provisional) for the month of <Month>, <Year>"
- Return that percentage as a number (e.g. -1.21 not "-1.21%")
- This is the ALL COMMODITIES overall rate — do NOT use Food Index, Primary Articles, or any sub-group rate
- For the month field, use the data month in YYYY-MM format (e.g. October 2025 → 2025-10)

Return JSON only:""",
        output_schema={"month": "YYYY-MM", "WPI": 0.0},
    ),

    "cpi": MetricConfig(
        name="Consumer Price Index",
        metric_type="inflation_cpi",
        query="Headline Inflation year-on-year CPI General Provisional percentage",
        document_hint="cpi",
        chunk_type_filter="paragraph",
        extraction_prompt="""Extract the CPI headline inflation rate from this content. The report period is {period} but CPI data may lag — extract the LATEST available month.

Content:
{table}

Instructions:
- Look for a sentence like "Headline Inflation: Year-on-year inflation rate ... is X.XX% (Provisional)"
- Return the Combined (All India) CPI (General) year-on-year inflation percentage as a number (e.g. 5.48 not "5.48%")
- For the month field, use the data month in YYYY-MM format (e.g. September 2025 → 2025-09)
- If only Rural/Urban values are present, return null

Return JSON only:""",
        output_schema={"month": "YYYY-MM", "CPI": 0.0},
    ),

    "iip": MetricConfig(
        name="Index of Industrial Production",
        metric_type="iip",
        query="IIP index industrial production growth rate percent General sector latest month",
        document_hint="iip",
        chunk_type_filter="paragraph",
        bundle_chunks=True,
        n_results_override=5,
        extraction_prompt="""Extract the IIP (Index of Industrial Production) overall/General growth rate for the latest available month.

Content:
{table}

Report period: {period}. IIP data lags ~2 months — extract the LATEST month available.

The content is from an IIP press release. Look for:
- Sentences like "IIP growth rate for the month of December 2025 is 7.8 percent"
- Or "Index of industrial production records growth of X% in <Month Year>"
- The overall/General rate — NOT sector-specific rates (Mining, Manufacturing, Electricity)

Instructions:
- Return the growth rate as a number (e.g. 7.8, not "7.8%")
- For the month field, return the data month in YYYY-MM format (e.g. December 2025 → 2025-12)
- If no overall IIP growth rate is found, return null for IIP

Return JSON only:""",
        output_schema={"month": "YYYY-MM", "IIP": 0.0},
    ),

    "gdp": MetricConfig(
        name="Gross Domestic Product",
        metric_type="gdp",
        query="GDP gross domestic product quarterly Rs Crore",
        period_type="quarter",
        append_period_to_query=True,
        document_hint="gdp",
        extraction_prompt="""Extract the latest available GDP value for the period around {period}.

Content:
{table}

Instructions:
- Look for GDP at constant prices (Real GDP) or GDP (Rs. Cr) or "GDP at 2011-12 prices"
- Find the LATEST AVAILABLE quarter. Prefer the most recent fiscal year over older years.
  Indian fiscal years run April–March: Q1=Apr-Jun, Q2=Jul-Sep, Q3=Oct-Dec, Q4=Jan-Mar.
  Example notation: Q2-FY26 = July–September 2025; Q1-FY25 = April–June 2024.
- If the table shows multiple fiscal years, pick the row with the highest FY number and latest quarter.
- IMPORTANT: If the most recent quarter available in the table is more than 4 quarters before the
  target period {period}, the data is too old — return null for ALL fields. Do not extract stale data.
- Return the value in Rs. Crore (remove commas; Indian lakh notation: 47.88 lakh crore = 4788000).
- For growth_pct: Look for a "% change" or "Growth (%)" or "YoY growth" column in the same row.
  The growth column may be labelled "Percentage change over previous year", "Growth Rate", or similar.
  If the value is shown as a ratio (e.g. 1.078), convert to percentage (7.8).
  If no growth_pct column exists in the table, return null — do NOT compute it yourself.

Return JSON only:""",
        output_schema={"quarter": "Q2-FY26", "gdp_rs_cr": 0, "growth_pct": 0.0},
    ),

"gdp_forecast": MetricConfig(
    name="GDP Forecast",
    metric_type="gdp_forecast",
    query="GDP forecast projection RBI World Bank IMF BMI Fitch S&P growth",
    period_type="year",
    bundle_chunks=True,
    n_results_override=5,
    extraction_prompt="""Extract GDP growth forecasts from RBI, World Bank, IMF, BMI Fitch, and S&P from this content.

Content:
{table}

Instructions:
- Find forecasts from any of these institutions that are present:
  1. RBI
  2. World Bank
  3. IMF
  4. BMI Fitch
  5. S&P
- "World Bank" may appear as "WB".
- "BMI Fitch" may appear as "BMI", "Fitch Solutions", or similar.
- "S&P" may appear as "S&P Global", "Standard & Poor's", or similar.
- RBI forecasts may appear in monetary policy summaries.
- Extract the forecast percentage and release date if available.
- Preserve the forecast period/year exactly.
- Indian fiscal years must use FY notation:
  FY25 = 2024-25
  FY26 = 2025-26
  FY27 = 2026-27
- Calendar years must use CY notation:
  CY26 = 2026
  CY27 = 2027
- If an institution provides forecasts for multiple years, include ALL available forecast years.
- For multiple years, put each forecast on a separate line using \\n.
- Example:
  "7.2% (FY26)\\n6.5% (FY27)"
- For a single-year forecast, use:
  "7.3% (FY26)"
- Keep the percentage and year together. Do NOT return only a raw number such as 6.5.
- Extract the release date if available.
- If a particular institution's forecast is not found in the content, return null for that institution.
- Do not invent or estimate missing forecasts.

Return JSON only:""",
    output_schema={
        "RBI": {
            "forecast": "X.X% (FY26)",
            "release_date": ""
        },
        "World_Bank": {
            "forecast": "X.X% (FY26)",
            "release_date": ""
        },
        "IMF": {
            "forecast": "X.X% (CY26)\\nX.X% (CY27)",
            "release_date": ""
        },
        "BMI_Fitch": {
            "forecast": "X.X% (FY26)",
            "release_date": ""
        },
        "S&P": {
            "forecast": "X.X% (FY26)",
            "release_date": ""
        }
    },
),

    "gas_stats_monthly": MetricConfig(
        name="Natural Gas Monthly Stats",
        metric_type="gas_production",
        query="natural gas net production LNG import total consumption MMSCM monthly",
        extraction_prompt="""Extract Natural Gas statistics for {period} from this table.

Table:
{table}

Instructions:
- Find "Net Production" row and get the value for {period} in MMSCM
- Find "LNG Import" row and get the value for {period} in MMSCM
- Find "Total Consumption" or compute Net Production + LNG Import
- All values should be in MMSCM (Million Standard Cubic Meters) — monthly totals, NOT daily averages
- IMPORTANT: If the table shows LPG production, LPG data (in MMT), quarterly/yearly gas data,
  OR any data that is NOT monthly natural gas production in MMSCM, return null for all values.

Return JSON only:""",
        output_schema={"month": "YYYY-MM", "net": 0, "lng": 0, "total": 0},
    ),

    "gas_stats_yearly": MetricConfig(
        name="Natural Gas Yearly Stats",
        metric_type="gas_production",
        period_type="year",
        query="natural gas net production LNG import total consumption MMSCMD yearly annual",
        append_period_to_query=True,
        extraction_prompt="""Extract Natural Gas yearly statistics for {period} from this table.

Table:
{table}

Instructions:
- Find "Net Production" row; find "LNG Import" row; find "Total Consumption" row
- The fiscal year columns may be labelled like "2024-25", "FY25", "2024-2025", etc.
- Find the column for {period}. If {period} is not present, use the MOST RECENT fiscal year
  (the rightmost/latest column with data) that is 2023-24 or later.
- Values should be in MMSCMD (Million Standard Cubic Meters per Day) — NOT annual totals
- IMPORTANT: If the table is a "Historical" table with data only up to 2020-21 or earlier,
  OR if the data appears to be annual totals in MMSCM rather than daily averages in MMSCMD,
  return null for all numeric fields — do NOT compute or estimate daily averages.
- For the year field, return the fiscal year label from the table (e.g. "2024-25")

Return JSON only:""",
        output_schema={"year": "2024-25", "net": 0, "lng": 0, "total": 0},
    ),

    "interest_repo": MetricConfig(
        name="RBI Repo Rate",
        metric_type="interest_repo",
        query="RBI repo rate MCLR 1-Year interest rate",
        document_hint="cmd information",
        chunk_type_filter="table",
        bundle_chunks=True,
        n_results_override=3,
        extraction_prompt="""Extract the RBI Policy Repo Rate and MCLR 1-Year Rate for {period} from this table.

Table:
{table}

Instructions:
- The table has columns including "Effective Date" (or "Month"), "RBI Repo Rate", and "MCLR-1Y" (or similar)
- Find the row where the date/month matches {period} (e.g. September 2025, Oct-25, Oct 2025, or 2025-09)
  or use the MOST RECENT row with non-null values at or before that date
- Rate values may be in decimal form (0.055 = 5.5%) or percentage form (5.5) or string ("5.50%") — always convert to plain percentage number
- Repo rate is typically 5.5–7.0%; MCLR 1-year is typically 8.0–9.5%
- If the row for the target month is blank/empty, use the nearest prior row with actual values
- For the month field, return the actual date of the row used, in YYYY-MM format
- Return rates as percentage numbers (e.g. 5.5 not "5.5%")

Return JSON only:""",
        output_schema={"month": "YYYY-MM", "repo_rate": 0.0, "mclr_1y": 0.0},
    ),

    "cgd_data": MetricConfig(
        name="CGD Entity Data",
        metric_type="cgd",
        query="CGD entity Internal Target DPNG Actual CNG Industrial Commercial PNG",
        append_period_to_query=True,
        document_hint="cgd",
        chunk_type_filter="email_body",
        extraction_prompt="""Extract CGD (City Gas Distribution) entity-wise data for {period} from this email content.
The email is a reply chain — it contains tables for multiple months. Extract ONLY the table
for FY 2025-26 cumulative data up to {period} (e.g. "April - August" for August 2025).

Content:
{table}

Instructions:
- Find the table with heading like "FY 2025-26 (April - <month>)" matching {period}
- Extract columns: Internal Target (DPNG), Actual (DPNG), Internal Target (CNG), Actual (CNG), Industrial PNG, Commercial PNG
- Extract ALL entities (IGL, MNGL, CUGL, GGL, AGL, BGL, MGL + UEPL, TNGCL, GAIL Gas + JVCs, GAIL 6 Cities / Six Cities, BGCL, Total)
- Remove commas from numbers (3,70,000 → 370000)
- dpng_target / dpng_actual = DPNG (domestic PNG) connections
- cng_target / cng_actual = CNG stations
- industrial_png = Industrial PNG connections
- commercial_png = Commercial PNG connections

Return JSON array with one object per entity:""",
        output_schema=[{"entity": "IGL", "dpng_target": 0, "dpng_actual": 0, "cng_target": 0, "cng_actual": 0, "industrial_png": 0, "commercial_png": 0}],
    ),

    "cgd_cumulative": MetricConfig(
        name="CGD Cumulative Connections",
        metric_type="cgd",
        query="CGD cumulative connections DPNG CNG IGL MNGL CUGL since inception total",
        document_hint="cgd",
        chunk_type_filter="table",
        extraction_prompt="""Extract CGD cumulative (since inception) DPNG and CNG connections for {period}.
The email is a reply chain — find the cumulative table that corresponds most closely to {period}.

Table:
{table}

Instructions:
- Look for a table explicitly labelled "Cumulative" or "DPNG (Lakhs) - Cumulative" with DPNG values in lakhs
- The cumulative table has DPNG in lakhs (e.g. 26.28 lakhs for IGL), not absolute counts
- Do NOT use the FY annual actuals table (which has absolute DPNG counts like 83,689)
- If multiple cumulative tables are present (one per month in the chain), use the one closest to {period}
- If no cumulative table is found, return null for all entities

Return JSON array:""",
        output_schema=[{"entity": "IGL", "dpng_cumulative": 0, "cng_cumulative": 0}],
    ),

    "retail_lng": MetricConfig(
        name="Retail LNG / LCNG Stations",
        metric_type="retail_lng",
        query="GAIL retail LNG LCNG stations entities locations",
        document_hint="cgd",
        chunk_type_filter="email_body",
        extraction_prompt="""Extract the Retail LNG and LCNG station details for {period} from this email content.
The email is a reply chain — it may have multiple versions of this table (empty templates and filled-in ones).

Content:
{table}

Instructions:
- Find the FIRST occurrence of the LNG/LCNG station table that has actual data filled in
  (i.e. entity names like MGL, VGL, IGL are present alongside location names)
- Ignore empty template rows where the entity and location columns are blank
- The table has 3 station types: "Only LNG stations", "LNG + LCNG stations", "Only LCNG stations"
- Each row has: station type, JV/Subsidiary entity names, Locations (city + state)

Return JSON array:""",
        output_schema=[{
            "type": "Only LNG stations",
            "entities": ["MGL", "VGL"],
            "locations": ["Savroli (Maharashtra)"]
        }],
    ),

    "satat_data": MetricConfig(
        name="SATAT Scheme Data",
        metric_type="satat",
        query="SATAT CBG scheme EOI LOI plants commissioned CBG sale ROS tonnes",
        document_hint="cbg",
        chunk_type_filter="email_body",
        extraction_prompt="""Extract SATAT/CBG scheme data as of {period} from this email content.
The email may be a reply chain — extract only the data relevant to {period}.

Content:
{table}

Instructions:
- Find EOIs Received (No.), LOIs Issued (No.)
- Find Financial Closure count (No.)
- Find CBG Plants Commissioned (No.)
- Find Cumulative CBG sale from ROs (in Tonnes — may be in thousands, keep as-is)
- Find CBG Sale for the current month closest to {period} (Tonnes)
- Find CBG Sale for the current financial year as of {period} (Tonnes)
- If a value has a * suffix (e.g. "11.18*"), include the number without the asterisk
- Plants Under Construction and CBG Offtake may not be present — use null if missing

Return JSON only:""",
        output_schema={
            "eoi_received": 0,
            "loi_issued": 0,
            "financial_closure": 0,
            "plants_commissioned": 0,
            "plants_under_construction": 0,
            "cbg_offtake_tpd": 0.0,
            "cumulative_cbg_sale_ros": 0.0,
            "cbg_sale_current_month": 0.0,
            "cbg_sale_fy": 0.0,
        },
    ),
}
