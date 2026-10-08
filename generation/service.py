"""
Main report generation service.

Orchestrates:
1. Validation of extraction results
2. Transformation to generation data models
3. Table updates
4. Chart generation
5. Document saving
"""

import io
import json
import calendar
import logging
import re
import shutil
import subprocess
import tempfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union
from datetime import datetime

from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openai import AzureOpenAI, OpenAI

try:
    from ..config import AzureConfig, OpenAIConfig, LLMConfig
    from .news_scraper import scrape_news_for_month, summarize_news
    from .validation import validate_extraction_results, generate_validation_report
    from .schemas import ValidationResult
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
        set_table_column_widths,
        markdown_to_docx_table,
        update_table_from_markdown,
        add_update_note_after_table,
    )
    from .table_updater import (
        TableUpdater,
        parse_quarter,
        sort_gdp_history,
        filter_gdp_by_cutoff,
        build_gdp_table_markdown,
        _fill_computed_fields,
        get_query_embedding,
        extract_table_from_chromadb,
    )
    from .msg_table_extractor import MsgTableExtractor
except ImportError:
    from config import AzureConfig, OpenAIConfig, LLMConfig
    from generation.news_scraper import scrape_news_for_month, summarize_news
    from generation.validation import validate_extraction_results, generate_validation_report
    from generation.schemas import ValidationResult
    from generation.msg_table_extractor import MsgTableExtractor
    from generation.chart_updater import (
        update_docx_charts,
        make_inflation_update,
        make_interest_rates_update,
        make_gas_yearly_update,
        make_gas_monthly_update,
        make_share_price_update,
    )
    from generation.tables import (
        snapshot_table_styles,
        restore_table_styles,
        force_black_grid,
        set_table_column_widths,
        markdown_to_docx_table,
        update_table_from_markdown,
        add_update_note_after_table,
    )
    from generation.table_updater import (
        TableUpdater,
        parse_quarter,
        sort_gdp_history,
        filter_gdp_by_cutoff,
        build_gdp_table_markdown,
        _fill_computed_fields,
        get_query_embedding,
        extract_table_from_chromadb,
    )

logger = logging.getLogger(__name__)

# XML 1.0 illegal characters — strip before embedding text in DOCX
_XML_ILLEGAL = re.compile(
    r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F\uD800-\uDFFF\uFFFE\uFFFF]"
)


def _sanitize_xml(text: str) -> str:
    """Remove characters that are illegal in XML 1.0 to prevent DOCX corruption."""
    return _XML_ILLEGAL.sub("", text)


# =============================================================================
# PDF Conversion Utilities (LibreOffice headless)
# =============================================================================

def _find_soffice() -> Optional[str]:
    """
    Locate the LibreOffice soffice binary on Mac, Linux, or Windows.

    Install:
      Linux  : sudo apt-get install -y libreoffice-writer
      Mac    : brew install --cask libreoffice
      Windows: download from https://www.libreoffice.org
    """
    candidates = [
        "soffice",
        "libreoffice",
        "/usr/bin/soffice",
        "/usr/lib/libreoffice/program/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",  # macOS
        r"C:\Program Files\LibreOffice\program\soffice.exe",      # Windows
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for c in candidates:
        if shutil.which(c) or Path(c).exists():
            return c
    return None


def convert_docx_to_pdf(docx_path: Path, pdf_path: Optional[Path] = None) -> Optional[Path]:
    """
    Convert a DOCX file to PDF using docx2pdf (MS Word) or LibreOffice headless.

    Args:
        docx_path: Path to the source DOCX file.
        pdf_path:  Desired output PDF path. Defaults to same name as DOCX with .pdf extension.

    Returns:
        Path to the generated PDF, or None if conversion tools are not available.
    """
    if pdf_path is None:
        pdf_path = docx_path.with_suffix(".pdf")

    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    # 1. Try docx2pdf (Native Word automation on Windows / macOS)
    try:
        from docx2pdf import convert
        import pythoncom
        pythoncom.CoInitialize()
        logger.info(f"Converting DOCX to PDF using docx2pdf: {docx_path} -> {pdf_path}")
        convert(str(docx_path), str(pdf_path))
        if pdf_path.exists() and pdf_path.stat().st_size > 0:
            logger.info(f"PDF saved via docx2pdf: {pdf_path} ({pdf_path.stat().st_size:,} bytes)")
            return pdf_path
    except Exception as d2p_err:
        logger.debug(f"docx2pdf conversion failed/not applicable: {d2p_err}")

    # 2. Try LibreOffice headless
    soffice = _find_soffice()
    if soffice is not None:
        try:
            logger.info(f"Converting DOCX to PDF using LibreOffice: {soffice}")
            subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf",
                 "--outdir", str(pdf_path.parent), str(docx_path)],
                check=True,
                capture_output=True,
            )
            lo_output = pdf_path.parent / (docx_path.stem + ".pdf")
            if lo_output != pdf_path and lo_output.exists():
                lo_output.rename(pdf_path)

            if pdf_path.exists():
                logger.info(f"PDF saved via LibreOffice: {pdf_path} ({pdf_path.stat().st_size:,} bytes)")
                return pdf_path
        except Exception as lo_err:
            logger.error(f"LibreOffice conversion failed: {lo_err}")

    logger.warning("PDF conversion skipped: neither docx2pdf (Word) nor LibreOffice is available.")
    return None


def fetch_usd_inr_rate(month: str) -> Optional[float]:
    """
    Fetch USD/INR reference rate from RBI for the last trading day of the given month.

    Scrapes https://www.rbi.org.in/scripts/ReferenceRateArchive.aspx
    Queries a 5-day window ending on the last calendar day so that holidays are handled
    (returns the most recent available rate within the window).
    Returns INR per 1 USD (e.g., 85.62), or None if fetch fails.

    Args:
        month: YYYY-MM format (e.g., "2025-11")
    """
    import calendar
    import datetime
    import requests
    from bs4 import BeautifulSoup

    try:
        year, mon = int(month[:4]), int(month[5:7])
        last_day = calendar.monthrange(year, mon)[1]
        end_date = datetime.date(year, mon, last_day)
        # Query 5-day window so weekends/holidays are covered
        start_date = end_date - datetime.timedelta(days=5)

        to_rbi = lambda d: d.strftime("%d/%m/%Y")

        url = "https://www.rbi.org.in/scripts/ReferenceRateArchive.aspx"
        session = requests.Session()
        session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        })

        # Step 1: GET page to extract ASP.NET hidden tokens
        resp = session.get(url, timeout=10)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")

        def _hidden(field_id):
            tag = soup.find("input", {"id": field_id})
            return tag["value"] if tag else ""

        # Step 2: POST — RBI form requires __EVENTTARGET/__EVENTARGUMENT + btnSubmit
        payload = {
            "__EVENTTARGET": "",
            "__EVENTARGUMENT": "",
            "__VIEWSTATE": _hidden("__VIEWSTATE"),
            "__VIEWSTATEGENERATOR": _hidden("__VIEWSTATEGENERATOR"),
            "__EVENTVALIDATION": _hidden("__EVENTVALIDATION"),
            "txtFromDate": to_rbi(start_date),
            "txtToDate": to_rbi(end_date),
            "chkUSD": "on",
            "btnSubmit": " GO ",
        }

        resp2 = session.post(url, data=payload, timeout=10)
        resp2.raise_for_status()

        # Step 3: Find the table whose header contains "USD", take first data row
        soup2 = BeautifulSoup(resp2.text, "html.parser")
        for table in soup2.find_all("table"):
            rows = table.find_all("tr")
            if len(rows) < 2:
                continue
            headers = [th.get_text(strip=True).upper() for th in rows[0].find_all(["th", "td"])]
            usd_col_idx = next((i for i, h in enumerate(headers) if "USD" in h), None)
            if usd_col_idx is None:
                continue
            # Rows are newest-first; first data row = most recent trading day
            for row in rows[1:]:
                cells = row.find_all(["td", "th"])
                if usd_col_idx < len(cells):
                    raw = cells[usd_col_idx].get_text(strip=True).replace(",", "")
                    try:
                        rate = float(raw)
                        if 50 < rate < 200:  # sanity check: plausible INR/USD range
                            date_cell = cells[0].get_text(strip=True)
                            logger.info(f"RBI USD/INR rate for {month} (date: {date_cell}): {rate}")
                            return rate
                    except ValueError:
                        continue

        logger.warning(f"Could not parse USD/INR rate from RBI response for {month}")
        return None

    except Exception as e:
        logger.warning(f"Could not fetch RBI USD/INR rate for {month}: {e}")
        return None


def fetch_gail_share_month(month: str) -> Optional[dict]:
    """
    Fetch GAIL (NSE: GAIL.NS) closing price and market cap for a single month.

    Uses the last available trading day's closing price within that month.
    Market cap (Rs Cr.) = close_price × shares_outstanding_crore, rounded to 2dp.
    GAIL Share price is rounded to 2 decimal places.

    Args:
        month: "YYYY-MM" format (e.g. "2025-11")

    Returns:
        {"month": "YYYY-MM", "Market Cap (Rs Cr.)": float, "GAIL Share": float}
        or None if the fetch fails.
    """
    try:
        import calendar
        import yfinance as yf

        year, mon = int(month[:4]), int(month[5:7])
        # Build a [first, last+1) date window for the month
        last_day = calendar.monthrange(year, mon)[1]
        start = f"{year}-{mon:02d}-01"
        if mon == 12:
            end = f"{year + 1}-01-01"
        else:
            end = f"{year}-{mon + 1:02d}-01"

        ticker = yf.Ticker("GAIL.NS")
        hist = ticker.history(start=start, end=end)

        if hist.empty:
            logger.warning(f"Yahoo Finance returned no data for GAIL.NS in {month}")
            return None

        # Last trading day of the month
        close = round(float(hist["Close"].iloc[-1]), 2)

        # Shares outstanding → crore shares for market cap
        info = ticker.info
        shares_abs = info.get("sharesOutstanding") or info.get("impliedSharesOutstanding")
        if shares_abs and shares_abs > 1e6:
            shares_crore = shares_abs / 1e7
        else:
            shares_crore = 657.0  # GAIL fallback constant (crore shares, circa 2024-25)
            logger.warning(f"GAIL shares outstanding not available from Yahoo, using {shares_crore}Cr")

        market_cap_cr = round(close * shares_crore, 2)

        logger.info(f"Fetched GAIL share data for {month}: ₹{close}, Mkt Cap ₹{market_cap_cr}Cr")
        return {
            "month": month,
            "Market Cap (Rs Cr.)": market_cap_cr,
            "GAIL Share": close,
        }

    except Exception as e:
        logger.warning(f"Could not fetch GAIL share data for {month}: {e}")
        return None


@dataclass
class GenerationConfig:
    """Configuration for report generation."""
    template_path: Path
    output_path: Optional[Path] = None
    report_month: str = ""  # YYYY-MM format
    data_month: str = ""  # YYYY-MM format, actual month of source data (T-1)
    template_type: str = "12"  # "6" or "12" months
    strict_validation: bool = False
    temp_dir: Optional[Path] = None
    generate_pdf: bool = True  # Also convert output DOCX → PDF via LibreOffice
    usd_exchange_rate: Optional[float] = None  # INR per 1 USD (e.g., 84.5). If None, gdp_usd_bn not calculated.
    history_cutoff_month: Optional[str] = None
    # Filter month-keyed history to entries <= this value before chart generation.
    # Defaults to data_month inside _transform_data if not set.
    persist_history: bool = True
    # Set False to suppress all writes to output/data/*.json (safe for test runs).

    def __post_init__(self):
        self.template_path = Path(self.template_path)
        if self.output_path:
            self.output_path = Path(self.output_path)
        if self.temp_dir:
            self.temp_dir = Path(self.temp_dir)
        else:
            self.temp_dir = Path(tempfile.gettempdir())


@dataclass
class GenerationResult:
    """Result of report generation."""
    success: bool
    output_path: Optional[Path] = None
    pdf_path: Optional[Path] = None
    validation_result: Optional[ValidationResult] = None
    error: Optional[str] = None
    charts_generated: List[str] = field(default_factory=list)
    tables_updated: List[str] = field(default_factory=list)
    tables_not_updated: List[str] = field(default_factory=list)
    # Maps metric → actual data source used in the report (may differ from extraction
    # attempt when generation falls back to history files).
    sources_used: Dict[str, str] = field(default_factory=dict)


class HistoricalDataManager:
    """
    Manages historical data for chart generation.

    Loads, updates, and saves historical time-series data.
    """

    def __init__(self, data_dir: Path):
        self.data_dir = Path(data_dir)
        self.data_dir.mkdir(parents=True, exist_ok=True)

    def _get_file_path(self, data_type: str) -> Path:
        return self.data_dir / f"{data_type}_history.json"

    def load(self, data_type: str, default: List[dict] = None) -> List[dict]:
        """Load historical data from JSON file."""
        file_path = self._get_file_path(data_type)
        if file_path.exists():
            with open(file_path, "r") as f:
                return json.load(f)
        return default or []

    def save(self, data_type: str, data: List[dict]) -> None:
        """Save historical data to JSON file."""
        file_path = self._get_file_path(data_type)
        with open(file_path, "w") as f:
            json.dump(data, f, indent=2)

    def load_dict(self, data_type: str, default: Optional[dict] = None) -> dict:
        """Load a single dict from JSON (for non-list data like GDP forecast baseline)."""
        file_path = self._get_file_path(data_type)
        if file_path.exists():
            with open(file_path, "r") as f:
                return json.load(f)
        return default if default is not None else {}

    def save_dict(self, data_type: str, data: dict) -> None:
        """Save a single dict to JSON."""
        file_path = self._get_file_path(data_type)
        with open(file_path, "w") as f:
            json.dump(data, f, indent=2)

    def update_and_trim(
        self,
        data_type: str,
        new_entry: dict,
        key_field: str,
        max_entries: int = 12,
        cutoff_key: Optional[str] = None,
        persist: bool = True,
    ) -> List[dict]:
        """
        Update historical data with a new entry and trim to max entries.

        Args:
            data_type: Type of data (inflation, interest_rates, etc.)
            new_entry: New data entry to add
            key_field: Field to use for deduplication (e.g., "month")
            max_entries: Maximum number of entries to keep
            cutoff_key: If set, filter history to entries where key_field <= cutoff_key
                        before adding the new entry. Use to prevent future data appearing
                        in charts when generating reports for older months.
            persist: When False, skip writing to disk (safe for historical test runs).

        Returns:
            Updated data list
        """
        data = self.load(data_type)

        # Skip if key field is None — don't pollute history with bad entries
        if new_entry.get(key_field) is None:
            logger.warning(f"Skipping history update for '{data_type}': {key_field} is None")
            return data

        # Filter to entries at or before the cutoff
        if cutoff_key is not None:
            data = [d for d in data if (d.get(key_field) or "") <= cutoff_key]

        # Remove existing entry with same key
        key_value = new_entry.get(key_field)
        data = [d for d in data if d.get(key_field) != key_value]

        # Add new entry
        data.append(new_entry)

        # Sort and trim
        data = sorted(data, key=lambda x: x.get(key_field) or "")[-max_entries:]

        if persist:
            self.save(data_type, data)

        return data

    def load_filtered(
        self,
        data_type: str,
        key_field: str,
        cutoff: str,
        default: Optional[List[dict]] = None,
    ) -> List[dict]:
        """Load history and filter to entries where key_field <= cutoff."""
        return [
            d for d in self.load(data_type, default or [])
            if d.get(key_field, "") <= cutoff
        ]


class ReportGenerator:
    """
    Main report generator class.

    Orchestrates the full pipeline:
    1. Validate extraction results
    2. Transform to generation data models
    3. Update tables
    4. Generate charts
    5. Save document
    """

    # Table indices in template
    # template1.docx ("12"): 5 tables at indices 0-4
    #   Table 0: GDP (13 columns: label + 12 quarters)
    #   Table 1: GDP Forecast (3 columns: Agency, Forecast, Date)
    #   Table 2: SATAT (2 columns: Particulars, Details)
    #   Table 3: CGD Targets (7 columns: Entity + 6 metrics)
    #   Table 4: CGD Cumulative (13 columns: Entity + 12 months)
    # template2.docx ("6"): 6 tables at indices 0-5
    #   Table 0: GDP (same as template1)
    #   Table 1: GDP Forecast (same as template1)
    #   Table 2: SATAT (same as template1)
    #   Table 3: Retail LNG (5 rows × 4 cols: Stations, JV/Subsidiary, Count, Locations)
    #   Table 4: CGD Targets (shifted from index 3)
    #   Table 5: CGD Cumulative (shifted from index 4)
    TABLE_INDICES = {
        "12": {
            "gdp": 0,
            "gdp_forecast": 1,
            "satat": 2,
            "cgd": 3,
            "cgd_cumulative": 4,
        },
        "6": {
            "gdp": 0,
            "gdp_forecast": 1,
            "satat": 2,
            "retail_lng": 3,
            "cgd": 4,
            "cgd_cumulative": 5,
        }
    }

    # Chart indices in template
    CHART_INDICES = {
        "12": {
            "inflation": 0,
            "interest_rates": 1,
            "gas_yearly": 2,
            "gas_monthly": 3,
            "share_price": 4,
        },
        "6": {
            "inflation": 0,
            "interest_rates": 1,
            "share_price": 2,
        }
    }

    def __init__(
        self,
        config: GenerationConfig,
        data_dir: Optional[Path] = None,
        table_updater: Optional["TableUpdater"] = None,
        msg_extractor: Optional["MsgTableExtractor"] = None,
    ):
        """
        Initialize the report generator.

        Args:
            config: Generation configuration
            data_dir: Directory for historical data (optional)
            table_updater: Optional TableUpdater instance for ChromaDB fallback.
                           Used when extraction results are missing for a table.
            msg_extractor: Optional MsgTableExtractor for direct .msg email table
                           extraction. Takes priority over table_updater for
                           SATAT / CGD / LNG tables.
        """
        self.config = config
        self.doc: Optional[Document] = None
        self.validation_result: Optional[ValidationResult] = None
        self.table_updater = table_updater
        self.msg_extractor = msg_extractor

        # Historical data manager
        if data_dir:
            self.history = HistoricalDataManager(data_dir)
        else:
            self.history = None

    def validate(
        self,
        extraction_results: Dict[str, Any],
        strict: Optional[bool] = None
    ) -> ValidationResult:
        """
        Validate extraction results before generation.

        Args:
            extraction_results: Dictionary of extraction results
            strict: Override config strict mode

        Returns:
            ValidationResult with validation status and cleaned data
        """
        strict_mode = strict if strict is not None else self.config.strict_validation

        self.validation_result = validate_extraction_results(
            results=extraction_results,
            report_month=self.config.report_month,
            strict=strict_mode
        )

        logger.info(generate_validation_report(self.validation_result))

        return self.validation_result
    def _update_cover_page(self) -> None:
        """Dynamically populate Issue Date, Issue No, and Version on the cover page."""

        report_month = self.config.report_month or datetime.now().strftime("%Y-%m")

        try:
            dt = datetime.strptime(report_month, "%Y-%m")
            month_name = dt.strftime("%B %Y")   # e.g. August 2026
            month_num = dt.strftime("%m")       # e.g. 08
        except Exception:
            month_name = report_month
            month_num = "01"

        version_str = "Final"

        for t in self.doc.element.xpath('//w:t'):
            if t.text:

                if "Issue No -" in t.text and not any(c.isdigit() for c in t.text):
                    t.text = t.text.replace(
                        "Issue No -",
                        f"Issue No - {month_num}"
                    )

                elif "Issue Date:" in t.text and not any(c.isdigit() for c in t.text):
                    t.text = t.text.replace(
                        "Issue Date:",
                        f"Issue Date: {month_name}"
                    )

                elif "Version-" in t.text and "Final" not in t.text:
                    t.text = t.text.replace(
                        "Version-",
                        f"Version - {version_str}"
                    )

        logger.info(
            f"Cover page updated: Issue Date={month_name}, "
            f"Issue No={month_num}, Version={version_str}"
        )
    def generate(
        self,
        extraction_results: Dict[str, Any],
        skip_validation: bool = False
    ) -> GenerationResult:
        """
        Generate the full report.

        Args:
            extraction_results: Dictionary of extraction results
            skip_validation: Skip validation step (not recommended)

        Returns:
            GenerationResult with output path and status
        """
        result = GenerationResult(success=False)

        try:
            # Step 1: Validation
            if not skip_validation:
                validation = self.validate(extraction_results)
                result.validation_result = validation

                if not validation.valid:
                    result.error = f"Validation failed: {len(validation.errors)} errors"
                    return result

            # Step 2: Transform data (before template load so charts can be baked in)
            sources_used: Dict[str, str] = {}
            transformed = self._transform_data(extraction_results, sources_used)
            result.sources_used = sources_used

            # Step 3: Apply native OOXML chart updates to template bytes
            template_bytes = self.config.template_path.read_bytes()
            template_bytes, result.charts_generated = self._apply_chart_updates(
                transformed, template_bytes
            )

            # Step 4: Load template with charts already updated
            self.doc = Document(io.BytesIO(template_bytes))
            logger.info(f"Loaded template: {self.config.template_path}")

            # Step 4.5: Update cover page metadata
            self._update_cover_page()

            # Step 5: Update tables
            result.tables_updated, result.tables_not_updated = self._update_tables(transformed)

            # Step 5.5: Append news section
            self._update_news_section()

            # Step 6: Save document (DOCX + optional PDF)
            output_path, pdf_path = self._save_document()
            result.output_path = output_path
            result.pdf_path = pdf_path
            result.success = True

            logger.info(f"Report generated successfully: {output_path}")

        except Exception as e:
            logger.error(f"Report generation failed: {e}", exc_info=True)
            result.error = str(e)

        return result

    def _transform_data(self, extraction_results: Dict[str, Any], sources_used: Dict[str, str]) -> Dict[str, Any]:
        """
        Transform extraction results into chart/table data formats.

        This merges extracted data with historical data for time-series.
        Populates sources_used with the actual data source used for each metric.
        """
        transformed = {}

        # Get validated data for normal report fields.
        # For CPI/WPI/IIP time-series, preserve the original extraction results
        # because validation may reduce list-based historical rows to a single value.
        if self.validation_result and self.validation_result.validated_data:
            data = dict(self.validation_result.validated_data)

            for metric in ("cpi", "wpi", "iip"):
                if metric in extraction_results:
                    data[metric] = extraction_results[metric]
        else:
            data = extraction_results   

        # Use data_month for chart entries (actual month of source data)
        entry_month = self.config.data_month or self.config.report_month

        # Cutoff for month-keyed history filtering and persist flag
        cutoff = self.config.history_cutoff_month or self.config.report_month
        persist = self.config.persist_history
        if not persist or self.config.history_cutoff_month:
            logger.info(
                f"History settings: history_cutoff_month={cutoff}, persist_history={persist}"
            )

        # Fetch USD/INR exchange rate if not already set (for GDP USD Bn computation)
        if self.config.usd_exchange_rate is None:
            rate = fetch_usd_inr_rate(entry_month)
            if rate:
                self.config.usd_exchange_rate = rate

        # ============================================================
        # TRANSFORM INFLATION DATA
        # ============================================================
        #
        # IMPORTANT:
        # Fresh extraction may contain only the latest month.
        # Historical inflation data must therefore be merged with
        # the fresh extraction before building the chart.
        #
        # Fresh values take precedence over historical values.
        # Missing fresh values DO NOT overwrite historical values.
        # ============================================================

        if any(k in data for k in ["wpi", "cpi", "iip"]):

            def _normalise_metric(metric_data, metric_name):
                """
                Convert either:

                    {"month": "2026-07", "CPI": 4.45}

                or:

                    [
                        {"month": "2026-06", "CPI": 2.10},
                        {"month": "2026-07", "CPI": 4.45},
                    ]

                or a scalar value

                into:

                    {
                        "2026-06": 2.10,
                        "2026-07": 4.45,
                    }
                """

                result = {}

                if isinstance(metric_data, list):

                    for item in metric_data:

                        if not isinstance(item, dict):
                            continue

                        month = item.get("month")

                        if not month:
                            continue

                        value = item.get(metric_name)

                        if value is not None:
                            result[month] = value

                elif isinstance(metric_data, dict):

                    month = (
                        metric_data.get("month")
                        or entry_month
                    )

                    value = metric_data.get(metric_name)

                    if value is not None:
                        result[month] = value

                elif metric_data is not None:

                    result[entry_month] = metric_data

                return result

            # --------------------------------------------------------
            # 1. Get freshly extracted values
            # --------------------------------------------------------

            fresh_cpi = _normalise_metric(
                data.get("cpi"),
                "CPI",
            )

            fresh_iip = _normalise_metric(
                data.get("iip"),
                "IIP",
            )

            fresh_wpi = _normalise_metric(
                data.get("wpi"),
                "WPI",
            )

            logger.info(
                "Fresh inflation extraction: "
                "CPI=%d months, WPI=%d months, IIP=%d months",
                len(fresh_cpi),
                len(fresh_wpi),
                len(fresh_iip),
            )

            logger.info(
                "Fresh CPI months: %s",
                sorted(fresh_cpi.keys()),
            )

            logger.info(
                "Fresh WPI months: %s",
                sorted(fresh_wpi.keys()),
            )

            logger.info(
                "Fresh IIP months: %s",
                sorted(fresh_iip.keys()),
            )

            # --------------------------------------------------------
            # 2. Load historical inflation data
            # --------------------------------------------------------

            historical = []

            if self.history:

                try:

                    historical = self.history.load_filtered(
                        "inflation",
                        "month",
                        cutoff,
                    ) or []

                except Exception as exc:

                    logger.warning(
                        "Could not load historical inflation data: %s",
                        exc,
                    )

                    historical = []

            logger.info(
                "Historical inflation rows loaded: %d",
                len(historical),
            )

            # --------------------------------------------------------
            # 3. Start with historical data
            # --------------------------------------------------------

            merged_by_month = {}

            for row in historical:

                if not isinstance(row, dict):
                    continue

                month = row.get("month")

                if not month:
                    continue

                if cutoff and month > cutoff:
                    continue

                merged_by_month[month] = {
                    "month": month,
                    "WPI": row.get("WPI"),
                    "CPI": row.get("CPI"),
                    "IIP": row.get("IIP"),
                }

            # --------------------------------------------------------
            # 4. Merge fresh CPI
            #
            # Fresh non-null values override history.
            # --------------------------------------------------------

            for month, value in fresh_cpi.items():

                if cutoff and month > cutoff:
                    continue

                if month not in merged_by_month:

                    merged_by_month[month] = {
                        "month": month,
                        "WPI": None,
                        "CPI": None,
                        "IIP": None,
                    }

                merged_by_month[month]["CPI"] = value

            # --------------------------------------------------------
            # 5. Merge fresh WPI
            # --------------------------------------------------------

            for month, value in fresh_wpi.items():

                if cutoff and month > cutoff:
                    continue

                if month not in merged_by_month:

                    merged_by_month[month] = {
                        "month": month,
                        "WPI": None,
                        "CPI": None,
                        "IIP": None,
                    }

                merged_by_month[month]["WPI"] = value

            # --------------------------------------------------------
            # 6. Merge fresh IIP
            # --------------------------------------------------------

            for month, value in fresh_iip.items():

                if cutoff and month > cutoff:
                    continue

                if month not in merged_by_month:

                    merged_by_month[month] = {
                        "month": month,
                        "WPI": None,
                        "CPI": None,
                        "IIP": None,
                    }

                merged_by_month[month]["IIP"] = value

            # --------------------------------------------------------
            # 7. Build chronological inflation data
            # --------------------------------------------------------

            inflation_data = [
                merged_by_month[month]
                for month in sorted(merged_by_month)
            ]

            # --------------------------------------------------------
            # 8. Keep latest 12 months
            # --------------------------------------------------------

            inflation_data = inflation_data[-18:]

            transformed["inflation"] = inflation_data

            # --------------------------------------------------------
            # 9. Source tracking
            # --------------------------------------------------------

            if fresh_cpi:
                sources_used["cpi"] = "MoSPI CPI API"

            elif historical:
                sources_used["cpi"] = (
                    "inflation_history.json (carried forward)"
                )

            if fresh_iip:
                sources_used["iip"] = "MoSPI IIP API"

            elif historical:
                sources_used["iip"] = (
                    "inflation_history.json (carried forward)"
                )

            if fresh_wpi:
                sources_used["wpi"] = "extraction"

            elif historical:
                sources_used["wpi"] = (
                    "inflation_history.json (carried forward)"
                )

            # --------------------------------------------------------
            # 10. Detailed logging
            # --------------------------------------------------------

            logger.info(
                "================================================"
            )

            logger.info(
                "FINAL INFLATION DATA FOR CHART"
            )

            logger.info(
                "Inflation data transformed: %d months",
                len(inflation_data),
            )

            logger.info(
                "Inflation months: %s",
                [x["month"] for x in inflation_data],
            )

            for row in inflation_data:

                logger.info(
                    "Inflation row: "
                    "month=%s CPI=%s WPI=%s IIP=%s",
                    row.get("month"),
                    row.get("CPI"),
                    row.get("WPI"),
                    row.get("IIP"),
                )

            logger.info(
                "================================================"
            )

        elif self.history:

            existing = self.history.load_filtered(
                "inflation",
                "month",
                cutoff,
            )

            if existing:

                transformed["inflation"] = existing

                sources_used["wpi"] = (
                    "inflation_history.json (carried forward)"
                )

                sources_used["cpi"] = (
                    "inflation_history.json (carried forward)"
                )

                sources_used["iip"] = (
                    "inflation_history.json (carried forward)"
                )

        # Transform interest rates
        if "interest_repo" in data:
            repo = data["interest_repo"].get("repo_rate") or data["interest_repo"].get("Repo")
            mclr = data["interest_repo"].get("mclr_1y") or data["interest_repo"].get("MCLR1Y")

            if repo is None and mclr is None:
                # Extraction found nothing - just use existing historical data
                logger.warning("Interest rate extraction returned no values, using historical data only")
                if self.history:
                    transformed["interest_rates"] = self.history.load_filtered(
                        "interest_rates", "month", cutoff
                    )
                sources_used["interest_repo"] = "interest_rates_history.json (carried forward)"
            else:
                interest_entry = {
                    "month": entry_month,
                    "Repo": repo,
                    "MCLR1Y": mclr,
                }

                if self.history:
                    interest_data = self.history.load("interest_rates", [])

                    # Update existing month or add new month.
                    interest_data = [
                        d for d in interest_data
                        if d.get("month") != interest_entry.get("month")
                    ]
                    interest_data.append(interest_entry)

                    # Keep complete historical history on disk.
                    interest_data = sorted(
                        interest_data,
                        key=lambda x: x.get("month") or ""
                    )

                    if persist:
                        self.history.save("interest_rates", interest_data)

                    # Only the report/chart gets the cutoff.
                    interest_data = [
                        d for d in interest_data
                        if (d.get("month") or "") <= cutoff
                    ][-22:]
                else:
                    interest_data = [interest_entry]

                transformed["interest_rates"] = interest_data
                sources_used["interest_repo"] = "extraction"
        elif self.history:
            existing = self.history.load_filtered("interest_rates", "month", cutoff)
            if existing:
                transformed["interest_rates"] = existing
                sources_used["interest_repo"] = "interest_rates_history.json (no extraction, carried forward)"

        # Transform GDP — update history only if value changed
        if "gdp" in data and isinstance(data["gdp"], dict) and self.history:
            gdp_entry = data["gdp"]
            # Skip if extraction returned all nulls (stale/missing data) — fall through to history
            if gdp_entry.get("quarter") is None and gdp_entry.get("gdp_rs_cr") is None:
                logger.info("GDP extraction returned all nulls — skipping, will use existing history")
                sources_used["gdp"] = "gdp_history.json (extraction returned nulls, carried forward)"
            else:
                existing_history = self.history.load("gdp")
                existing = next(
                    (e for e in existing_history if e.get("quarter") == gdp_entry.get("quarter")),
                    None
                )
                if existing is None:
                    logger.info(f"GDP {gdp_entry.get('quarter')}: new quarter, adding to history")
                    gdp_history = self.history.update_and_trim(
                        "gdp", gdp_entry, "quarter", 20, persist=persist
                    )
                    sources_used["gdp"] = "extraction"
                elif (
                    existing.get("gdp_rs_cr") != gdp_entry.get("gdp_rs_cr")
                    or existing.get("growth_pct") != gdp_entry.get("growth_pct")
                ):
                    # Don't overwrite a good existing value with null from extraction
                    if gdp_entry.get("gdp_rs_cr") is None and existing.get("gdp_rs_cr") is not None:
                        logger.info(
                            f"GDP {gdp_entry.get('quarter')}: extraction returned null gdp_rs_cr, "
                            f"preserving existing value {existing.get('gdp_rs_cr')}"
                        )
                        gdp_history = existing_history
                        sources_used["gdp"] = (
                            f"gdp_history.json (extraction null for {gdp_entry.get('quarter')}, "
                            f"preserved existing {existing.get('gdp_rs_cr')})"
                        )
                    else:
                        logger.info(
                            f"GDP {gdp_entry.get('quarter')}: value changed "
                            f"{existing.get('gdp_rs_cr')} → {gdp_entry.get('gdp_rs_cr')}, updating"
                        )
                        gdp_history = self.history.update_and_trim(
                            "gdp", gdp_entry, "quarter", 20, persist=persist
                        )
                        sources_used["gdp"] = "extraction"
                else:
                    logger.info(
                        f"GDP {gdp_entry.get('quarter')}: value unchanged "
                        f"({gdp_entry.get('gdp_rs_cr')}), using existing history"
                    )
                    gdp_history = existing_history
                    sources_used["gdp"] = (
                        f"gdp_history.json (value unchanged {gdp_entry.get('gdp_rs_cr')})"
                    )
                transformed["gdp"] = gdp_history
        elif "gdp" in data:
            transformed["gdp"] = data["gdp"]
            sources_used["gdp"] = "extraction"

        # Apply quarter-based cutoff to GDP history when running benchmarks.
        # The monthly history_cutoff_month is converted to a quarter boundary so
        # that future quarters (not yet published at the report date) are excluded.
        # Rule: exclude quarters whose end month >= cutoff_month.
        if "gdp" in transformed and cutoff and isinstance(transformed["gdp"], list):
            filtered = filter_gdp_by_cutoff(transformed["gdp"], cutoff)
            if len(filtered) != len(transformed["gdp"]):
                excluded = [
                    q["quarter"] for q in transformed["gdp"]
                    if q not in filtered
                ]
                logger.info(
                    f"GDP history: excluded {excluded} (quarter end >= cutoff {cutoff})"
                )
            transformed["gdp"] = filtered

        # Load GDP history from file even if no new extraction (for table population)
        if "gdp" not in transformed and self.history:
            existing = self.history.load("gdp")
            if existing:
                if cutoff:
                    existing = filter_gdp_by_cutoff(existing, cutoff)
                transformed["gdp"] = existing
                sources_used.setdefault("gdp", "gdp_history.json (no extraction, loaded from file)")

        # Transform gas stats — unified monthly/yearly
        if "gas_stats_monthly" in data:
            gm = data["gas_stats_monthly"]
            if isinstance(gm, dict) and self.history:
                gas_monthly = self.history.update_and_trim(
                    "gas_monthly", gm, "month", 12,
                    cutoff_key=cutoff, persist=persist,
                )
                transformed["gas_monthly"] = gas_monthly
            elif isinstance(gm, list):
                transformed["gas_monthly"] = gm
        elif self.history:
            existing = self.history.load_filtered("gas_monthly", "month", cutoff)
            if existing:
                transformed["gas_monthly"] = existing

        if "gas_stats_yearly" in data:
            gy = data["gas_stats_yearly"]
            if isinstance(gy, dict) and self.history:
                # Round float values to 1 decimal place before storing
                for key in ("net", "lng", "total"):
                    if isinstance(gy.get(key), float):
                        gy[key] = round(gy[key], 1)
                gas_yearly = self.history.update_and_trim(
                    "gas_yearly", gy, "year", 12, persist=persist
                )
                transformed["gas_yearly"] = gas_yearly
            elif isinstance(gy, list):
                transformed["gas_yearly"] = gy
        elif self.history:
            existing = self.history.load("gas_yearly")
            if existing:
                transformed["gas_yearly"] = existing

        # Legacy gas_production key (backward compat)
        if "gas_production" in data and "gas_monthly" not in transformed:
            gp = data["gas_production"]
            if isinstance(gp, list) and gp and any(
                v is not None for d in gp for k, v in d.items() if k in ("net", "lng", "total")
            ):
                transformed["gas_monthly"] = gp

        # Transform share price — stored history for past months, Yahoo Finance
        # for the report month (not data_month — share price is real-time, no data lag).
        # Falls back to stored history on failure.
        share_price_month = self.config.report_month or entry_month
        if "share_price" in data and data["share_price"]:
            transformed["share_price"] = data["share_price"]
        elif self.history:
            # Build exactly 12 consecutive calendar months ending at report_month.
            # Fetch missing months from Yahoo Finance so gaps in stored history
            # do not shorten or distort the chart.
            report_year, report_mon = map(int, share_price_month.split("-"))
            required_months = []

            for offset in range(22, -1, -1):
                total = report_year * 12 + (report_mon - 1) - offset
                year = total // 12
                mon = total % 12 + 1
                required_months.append(f"{year:04d}-{mon:02d}")

            existing = self.history.load("gail_share_data", default=[])
            existing_by_month = {e["month"]: e for e in existing}

            for month in required_months:
                if month not in existing_by_month:
                    fresh_entry = fetch_gail_share_month(month)
                    if fresh_entry:
                        existing_by_month[month] = fresh_entry
                        logger.info(f"Fetched missing GAIL share history for {month}")

            merged = [
                existing_by_month[month]
                for month in required_months
                if month in existing_by_month
            ]

            if persist:
                self.history.save(
                    "gail_share_data",
                    sorted(
                        existing_by_month.values(),
                        key=lambda x: x["month"]
                    )[-24:]
                )

            if merged:
                transformed["share_price"] = merged

        # GDP Forecast — persist per-agency baseline; fall back to it if extraction is empty
        if "gdp_forecast" in data and isinstance(data.get("gdp_forecast"), dict) and self.history:
            new_forecast = data["gdp_forecast"]
            baseline = self.history.load_dict("gdp_forecast", default={})
            changed = []
            for agency in ("RBI", "World_Bank", "IMF", "BMI_Fitch", "S&P"):
                agency_data = new_forecast.get(agency)
                if agency_data and any(v for v in agency_data.values()):  # non-empty
                    if baseline.get(agency) != agency_data:
                        baseline[agency] = agency_data
                        changed.append(agency)
            if changed:
                if persist:
                    self.history.save_dict("gdp_forecast", baseline)
                logger.info(f"GDP forecast baseline updated for agencies: {changed}")
            transformed["gdp_forecast"] = baseline
        elif "gdp_forecast" in data:
            # history manager not configured — pass through as before
            transformed["gdp_forecast"] = data["gdp_forecast"]

        # Fallback: no new extraction — load stored baseline
        if "gdp_forecast" not in transformed and self.history:
            baseline = self.history.load_dict("gdp_forecast", default={})
            if baseline:
                logger.info("No new GDP forecast extraction data — using stored baseline")
                transformed["gdp_forecast"] = baseline

        # Pass through remaining table data
        for key in ["satat_data", "cgd_data", "cgd_cumulative", "retail_lng"]:
            if key in data:
                transformed[key] = data[key]

        return transformed

    def _clear_table_data_cells(self, table_index: int, skip_col_0: bool = True) -> None:
        """Clear all data cells, preserving header row (row 0) and optionally label col (col 0)."""
        if table_index >= len(self.doc.tables):
            return
        table = self.doc.tables[table_index]
        for row_idx, row in enumerate(table.rows):
            if row_idx == 0:
                continue  # keep header row
            for col_idx, cell in enumerate(row.cells):
                if skip_col_0 and col_idx == 0:
                    continue  # keep label column (agency/entity/metric names)
                cell.text = ""

    def _update_tables(self, data: Dict[str, Any]) -> Tuple[List[str], List[str]]:
        """Update tables in the document. Returns (updated, not_updated)."""
        updated = []
        not_updated = []
        template_type = self.config.template_type
        table_indices = self.TABLE_INDICES.get(template_type, self.TABLE_INDICES["12"])

        logger.info(f"Updating tables (template type: {template_type})")

        # GDP table
        if "gdp" in data:
            logger.info("Updating GDP table from extraction data")
            try:
                if isinstance(data["gdp"], str):
                    update_table_from_markdown(
                        self.doc,
                        table_indices["gdp"],
                        data["gdp"],
                    )
                    updated.append("gdp")
                elif isinstance(data["gdp"], list):
                    self._update_gdp_table_from_history(
                        table_indices["gdp"], data["gdp"]
                    )
                    updated.append("gdp")
                elif isinstance(data["gdp"], dict):
                    self._update_gdp_table_from_dict(
                        table_indices["gdp"], data["gdp"]
                    )
                    updated.append("gdp")
            except Exception as e:
                logger.error(f"Failed to update GDP table: {e}", exc_info=True)
        elif self.table_updater:
            logger.warning("No GDP extraction data, trying TableUpdater fallback")
            try:
                if self.table_updater.update_gdp_table(self.doc):
                    updated.append("gdp")
            except Exception as e:
                logger.error(f"GDP TableUpdater fallback failed: {e}", exc_info=True)

        # GDP forecast table
        if "gdp_forecast" in data:
            logger.info("Updating GDP forecast table from extraction data")
            try:
                if isinstance(data["gdp_forecast"], str):
                    update_table_from_markdown(
                        self.doc,
                        table_indices["gdp_forecast"],
                        data["gdp_forecast"],
                    )
                    updated.append("gdp_forecast")
                elif isinstance(data["gdp_forecast"], dict):
                    self._update_gdp_forecast_table(
                        table_indices["gdp_forecast"], data["gdp_forecast"]
                    )
                    updated.append("gdp_forecast")
            except Exception as e:
                logger.error(f"Failed to update GDP forecast table: {e}", exc_info=True)
        else:
            logger.warning("No GDP forecast data — clearing table cells")
            self._clear_table_data_cells(table_indices["gdp_forecast"])
            not_updated.append("gdp_forecast")

        # SATAT table
        # Priority: MsgTableExtractor (direct email HTML → most reliable) >
        #           extraction data (ChromaDB + LLM → may return wrong values) >
        #           ChromaDB TableUpdater > clear
        if self.msg_extractor:
            logger.info("Updating SATAT table — trying direct email extraction first")
            md = self.msg_extractor.extract_satat_table()
            if md:
                update_table_from_markdown(self.doc, table_indices["satat"], md)
                updated.append("satat")
            elif "satat_data" in data and data["satat_data"]:
                logger.warning("SATAT msg extraction returned nothing — falling back to extraction data")
                try:
                    if isinstance(data["satat_data"], str):
                        update_table_from_markdown(
                            self.doc,
                            table_indices["satat"],
                            data["satat_data"],
                        )
                        updated.append("satat")
                    elif isinstance(data["satat_data"], dict):
                        self._update_satat_table(
                            table_indices["satat"], data["satat_data"]
                        )
                        updated.append("satat")
                except Exception as e:
                    logger.error(f"Failed to update SATAT table from extraction: {e}", exc_info=True)
            elif self.table_updater:
                logger.warning("SATAT msg extraction returned nothing, trying ChromaDB TableUpdater")
                try:
                    if self.table_updater.update_satat_table(self.doc):
                        updated.append("satat")
                    else:
                        logger.warning("SATAT TableUpdater returned False — clearing table cells")
                        self._clear_table_data_cells(table_indices["satat"])
                        not_updated.append("satat")
                except Exception as e:
                    logger.error(f"SATAT TableUpdater fallback failed: {e}", exc_info=True)
                    self._clear_table_data_cells(table_indices["satat"])
                    not_updated.append("satat")
            else:
                logger.warning("SATAT msg extraction returned nothing — clearing table cells")
                self._clear_table_data_cells(table_indices["satat"])
                not_updated.append("satat")
        elif "satat_data" in data:
            logger.info("Updating SATAT table from extraction data")
            try:
                if isinstance(data["satat_data"], str):
                    update_table_from_markdown(
                        self.doc,
                        table_indices["satat"],
                        data["satat_data"],
                    )
                    updated.append("satat")
                elif isinstance(data["satat_data"], dict):
                    if not data["satat_data"]:
                        logger.warning("SATAT data dict is empty, skipping")
                    else:
                        self._update_satat_table(
                            table_indices["satat"], data["satat_data"]
                        )
                        updated.append("satat")
            except Exception as e:
                logger.error(f"Failed to update SATAT table: {e}", exc_info=True)
        elif self.table_updater:
            logger.warning("No SATAT data — trying ChromaDB TableUpdater")
            try:
                if self.table_updater.update_satat_table(self.doc):
                    updated.append("satat")
                else:
                    logger.warning("SATAT TableUpdater returned False — clearing table cells")
                    self._clear_table_data_cells(table_indices["satat"])
                    not_updated.append("satat")
            except Exception as e:
                logger.error(f"SATAT TableUpdater failed: {e}", exc_info=True)
                self._clear_table_data_cells(table_indices["satat"])
                not_updated.append("satat")
        else:
            logger.warning("No SATAT data and no TableUpdater — clearing table cells")
            self._clear_table_data_cells(table_indices["satat"])
            not_updated.append("satat")

        # CGD targets table
        # Priority: MsgTableExtractor (direct email HTML — entity rows with full values) >
        #           extraction data (ChromaDB + LLM — usually only partial entity data) >
        #           ChromaDB TableUpdater > clear
        if self.msg_extractor:
            logger.info("Updating CGD targets table — trying direct email extraction first")
            md = self.msg_extractor.extract_cgd_targets_table()
            if md:
                update_table_from_markdown(self.doc, table_indices["cgd"], md)
                updated.append("cgd")
            elif "cgd_data" in data and data["cgd_data"]:
                logger.info("CGD targets msg extraction returned nothing — falling back to extraction data")
                try:
                    if isinstance(data["cgd_data"], str):
                        update_table_from_markdown(self.doc, table_indices["cgd"], data["cgd_data"])
                    else:
                        self._update_cgd_table(table_indices["cgd"], data["cgd_data"])
                    updated.append("cgd")
                except Exception as e:
                    logger.error(f"CGD table extraction data fallback failed: {e}", exc_info=True)
                    self._clear_table_data_cells(table_indices["cgd"])
                    not_updated.append("cgd")
            elif self.table_updater:
                logger.warning("CGD targets msg extraction returned nothing — trying ChromaDB TableUpdater")
                try:
                    if self.table_updater.update_cgd_targets_table(self.doc):
                        updated.append("cgd")
                    else:
                        self._clear_table_data_cells(table_indices["cgd"])
                        not_updated.append("cgd")
                except Exception as e:
                    logger.error(f"CGD TableUpdater fallback failed: {e}", exc_info=True)
                    self._clear_table_data_cells(table_indices["cgd"])
                    not_updated.append("cgd")
            else:
                logger.warning("CGD targets: all extraction methods failed — clearing")
                self._clear_table_data_cells(table_indices["cgd"])
                not_updated.append("cgd")
        elif "cgd_data" in data and data["cgd_data"]:
            logger.info("No msg_extractor — updating CGD table from extraction data")
            try:
                if isinstance(data["cgd_data"], str):
                    update_table_from_markdown(self.doc, table_indices["cgd"], data["cgd_data"])
                else:
                    self._update_cgd_table(table_indices["cgd"], data["cgd_data"])
                updated.append("cgd")
            except Exception as e:
                logger.error(f"Failed to update CGD table: {e}", exc_info=True)
                self._clear_table_data_cells(table_indices["cgd"])
                not_updated.append("cgd")
        elif self.table_updater:
            logger.warning("No CGD data, trying ChromaDB direct query via TableUpdater")
            try:
                if self.table_updater.update_cgd_targets_table(self.doc):
                    updated.append("cgd")
                else:
                    self._clear_table_data_cells(table_indices["cgd"])
                    not_updated.append("cgd")
            except Exception as e:
                logger.error(f"CGD TableUpdater fallback failed: {e}", exc_info=True)
                self._clear_table_data_cells(table_indices["cgd"])
                not_updated.append("cgd")
        else:
            logger.warning("No CGD data and no TableUpdater — clearing table cells")
            self._clear_table_data_cells(table_indices["cgd"])
            not_updated.append("cgd")

        # Retail LNG table — template2 only (index 3)
        if "retail_lng" in table_indices:
            if "retail_lng" in data:
                logger.info("Updating Retail LNG table from extraction data")
                try:
                    self._update_retail_lng_table(
                        table_indices["retail_lng"], data["retail_lng"]
                    )
                    updated.append("retail_lng")
                except Exception as e:
                    logger.error(f"Failed to update Retail LNG table: {e}", exc_info=True)
            elif self.msg_extractor:
                logger.info("No Retail LNG extraction data — trying direct email extraction")
                md = self.msg_extractor.extract_lng_stations_table()
                if md:
                    self._update_retail_lng_table(
                        table_indices["retail_lng"],
                        md,
                    )
                    updated.append("retail_lng")
                elif self.table_updater:
                    logger.warning("LNG stations msg extraction returned nothing, trying ChromaDB TableUpdater")
                    try:
                        if self.table_updater.update_retail_lng_table(self.doc):
                            updated.append("retail_lng")
                        else:
                            logger.warning("Retail LNG TableUpdater returned False — clearing table cells")
                            self._clear_table_data_cells(table_indices["retail_lng"])
                            not_updated.append("retail_lng")
                    except Exception as e:
                        logger.error(f"Retail LNG TableUpdater fallback failed: {e}", exc_info=True)
                        self._clear_table_data_cells(table_indices["retail_lng"])
                        not_updated.append("retail_lng")
                else:
                    logger.warning("LNG stations msg extraction returned nothing and no TableUpdater — clearing")
                    self._clear_table_data_cells(table_indices["retail_lng"])
                    not_updated.append("retail_lng")
            elif self.table_updater:
                logger.warning("No Retail LNG data, trying TableUpdater fallback")
                try:
                    if self.table_updater.update_retail_lng_table(self.doc):
                        updated.append("retail_lng")
                    else:
                        logger.warning("Retail LNG TableUpdater returned False — clearing table cells")
                        self._clear_table_data_cells(table_indices["retail_lng"])
                        not_updated.append("retail_lng")
                except Exception as e:
                    logger.error(f"Retail LNG TableUpdater fallback failed: {e}", exc_info=True)
                    self._clear_table_data_cells(table_indices["retail_lng"])
                    not_updated.append("retail_lng")
            else:
                logger.warning("No Retail LNG data and no TableUpdater — clearing table cells")
                self._clear_table_data_cells(table_indices["retail_lng"])
                not_updated.append("retail_lng")

        # CGD cumulative table
        # Priority: MsgTableExtractor (metric rows, entity columns — matches GT structure) >
        #           extraction data (partial entity list) > ChromaDB TableUpdater > clear
        if self.msg_extractor:
            logger.info("Updating CGD cumulative table — trying direct email extraction first")
            md = self.msg_extractor.extract_cgd_cumulative_table()
            if md:
                update_table_from_markdown(self.doc, table_indices["cgd_cumulative"], md)
                updated.append("cgd_cumulative")
            elif "cgd_cumulative" in data and data["cgd_cumulative"]:
                # Email had no cumulative table — fall back to extraction data
                logger.info("CGD cumulative: no email table, falling back to extraction data")
                try:
                    self._update_cgd_cumulative_table(table_indices["cgd_cumulative"], data["cgd_cumulative"])
                    updated.append("cgd_cumulative")
                except Exception as e:
                    logger.error(f"CGD cumulative fallback failed: {e}", exc_info=True)
                    self._clear_table_data_cells(table_indices["cgd_cumulative"])
                    not_updated.append("cgd_cumulative")
            else:
                logger.warning("CGD cumulative: no email table and no extraction data — clearing")
                self._clear_table_data_cells(table_indices["cgd_cumulative"])
                not_updated.append("cgd_cumulative")
        elif "cgd_cumulative" in data and data["cgd_cumulative"]:
            logger.info("No msg_extractor — updating CGD cumulative from extraction data")
            try:
                if isinstance(data["cgd_cumulative"], str):
                    update_table_from_markdown(self.doc, table_indices["cgd_cumulative"], data["cgd_cumulative"])
                else:
                    self._update_cgd_cumulative_table(table_indices["cgd_cumulative"], data["cgd_cumulative"])
                updated.append("cgd_cumulative")
            except Exception as e:
                logger.error(f"Failed to update CGD cumulative table: {e}", exc_info=True)
                self._clear_table_data_cells(table_indices["cgd_cumulative"])
                not_updated.append("cgd_cumulative")
        elif self.table_updater:
            logger.warning("No CGD cumulative data, trying ChromaDB direct query via TableUpdater")
            try:
                if self.table_updater.update_cgd_cumulative_table(self.doc):
                    updated.append("cgd_cumulative")
                else:
                    self._clear_table_data_cells(table_indices["cgd_cumulative"])
                    not_updated.append("cgd_cumulative")
            except Exception as e:
                logger.error(f"CGD cumulative TableUpdater fallback failed: {e}", exc_info=True)
                self._clear_table_data_cells(table_indices["cgd_cumulative"])
                not_updated.append("cgd_cumulative")
        else:
            logger.warning("No CGD cumulative data and no TableUpdater — clearing table cells")
            self._clear_table_data_cells(table_indices["cgd_cumulative"])
            not_updated.append("cgd_cumulative")

        # CBG Plant production paragraphs
        if self.msg_extractor:
            if self._update_cbg_plant_paragraphs():
                updated.append("cbg_plant_production")
            else:
                not_updated.append("cbg_plant_production")

        self._clean_section_titles()

        logger.info(f"Table updates complete. Updated: {updated}. Not updated (cleared): {not_updated}")
        return updated, not_updated

    def _update_cbg_plant_paragraphs(self) -> bool:
        """Update CBG Plant production paragraphs from MsgTableExtractor."""
        import re as _re
        if not self.msg_extractor:
            return False
        data = self.msg_extractor.extract_cbg_production()
        if not data:
            logger.warning("CBG Plant: no production data found in .msg files")
            return False
        updated_any = False
        for para in self.doc.paragraphs:
            txt = para.text
            if _re.search(r"Production of CBG in the Month", txt, _re.IGNORECASE):
                if data.get("month_production_kg") and data.get("month_str"):
                    for run in para.runs:
                        run.text = ""
                    if para.runs:
                        para.runs[0].text = (
                            f"Production of CBG in the Month of {data['month_str']} : "
                            f"{data['month_production_kg']} Kg"
                        )
                    updated_any = True
            elif _re.search(r"Cumulative Production till", txt, _re.IGNORECASE):
                if data.get("cumulative_production_kg") and data.get("month_str"):
                    for run in para.runs:
                        run.text = ""
                    if para.runs:
                        para.runs[0].text = (
                            f"Cumulative Production till {data['month_str']} "
                            f"(since commissioning) : {data['cumulative_production_kg']} Kg"
                        )
                    updated_any = True
        if updated_any:
            logger.info("Updated CBG Plant production paragraphs")
        return updated_any

    def _clean_section_titles(self) -> None:
        """
        Strip hardcoded month references and location names from section headings.

        Patterns removed:
          - " (Till Jun'25)"  / " (Till August'25)"  etc.
          - " as on Jun'25"   / " as on August'25)"  etc.
          - " (updated quarterly)"
          - " Ranchi"         (CBG Plant Ranchi → CBG Plant)
          - ", FY 2025-26"    (or any FY YYYY-YY suffix)
        """
        import re

        # Match both ASCII apostrophe (') and Unicode right single quote (')
        apos = r"['\u2019]"
        # Title suffix patterns to strip
        title_patterns = [
            re.compile(r"\s*\(Till\s+[A-Za-z]+" + apos + r"\d{2}\)", re.IGNORECASE),
            re.compile(r"\s+as\s+on\s+[A-Za-z]+" + apos + r"\d{2}\)?", re.IGNORECASE),
            re.compile(r"\s*\(updated\s+quarterly\)", re.IGNORECASE),
            re.compile(r"\s+Ranchi\b", re.IGNORECASE),
            re.compile(r",\s*FY\s+\d{4}-\d{2,4}", re.IGNORECASE),
        ]

        # CBG Plant production data patterns — skip (handled by _update_cbg_plant_paragraphs)
        cbg_plant_content_patterns = [
            re.compile(r"Production of CBG in the Month", re.IGNORECASE),
            re.compile(r"Cumulative Production till", re.IGNORECASE),
        ]

        for para in self.doc.paragraphs:
            full_text = para.text
            if not full_text.strip():
                continue
            # Format the Retail LNG introductory sentence.
            if re.search(
                r"GAIL along with its JVs and Subsidiaries have installed the following stations",
                full_text,
                re.IGNORECASE,
            ):
                para.paragraph_format.space_before = Pt(0)
                para.paragraph_format.space_after = Pt(4)
                para.paragraph_format.line_spacing = 1.0

            # Skip CBG Plant production lines — _update_cbg_plant_paragraphs handles them
            if any(pat.search(full_text) for pat in cbg_plant_content_patterns):
                continue

            cleaned = full_text

            # Retail LNG title should show the actual month of the source data.
            if re.search(r"GAIL Retail LNG data", full_text, re.IGNORECASE):
                source_month = None

                if self.msg_extractor:
                    source_month = getattr(self.msg_extractor, "report_month", None)

                source_month = (
                    source_month
                    or self.config.data_month
                    or self.config.report_month
                )

                try:
                    dt = datetime.strptime(source_month, "%Y-%m")
                    month_name = dt.strftime("%B %Y")
                    cleaned = re.sub(
                        r"\s*\(Till.*?\)",
                        "",
                        full_text,
                        flags=re.IGNORECASE,
                    ).strip()
                    cleaned = f"{cleaned} (Till {month_name})"
                except Exception:
                    cleaned = full_text

            else:
                for pat in title_patterns:
                    cleaned = pat.sub("", cleaned)

            if cleaned != full_text:
                # Preserve first run's formatting; clear all runs then set text on first
                if para.runs:
                    para.runs[0].text = cleaned
                    for run in para.runs[1:]:
                        run.text = ""
                logger.info(f"Title cleaned: {full_text!r} → {cleaned!r}")

    def _update_gdp_table_from_history(self, table_index: int, gdp_history: List[Dict[str, Any]]) -> None:
        """
        Update the GDP table from full history list using rolling window.

        Uses sort_gdp_history() for chronological ordering and
        build_gdp_table_markdown() for proper value formatting:
          - Comma-separated GDP values (e.g., "4,218,520")
          - Percentage formatting for growth (e.g., "6.7%")
          - Decimal formatting for USD values

        Template GDP table layout (matches ground-truth CMD Dashboard):
          Row 0: headers   [GDP, Q1-FY23, Q2-FY23, ...]
          Row 1: Rs. Cr    [GDP (₹ Cr), val, val, ...]
          Row 2: USD Bn    [GDP (USD Bn), val, val, ...]
          Row 3: % Growth  [% Growth, val, val, ...]
        """
        if table_index >= len(self.doc.tables):
            logger.warning(f"GDP table index {table_index} out of range")
            return

        # Sort chronologically
        sorted_history = sort_gdp_history(gdp_history)

        # Fill computed fields BEFORE selecting the latest 9 quarters.
        # This allows YoY growth to use the previous year's quarter.
        sorted_history = _fill_computed_fields(
            sorted_history,
            self.config.usd_exchange_rate
        )

        # Keep only the latest 9 quarters for the report table
        entries = sorted_history[-9:]

        logger.info(
            f"Updating GDP table with {len(entries)} quarters: "
            f"{[e.get('quarter') for e in entries]}"
        )

        # Build markdown and update using style-pr  eserving utility
        markdown = build_gdp_table_markdown(entries)
        update_table_from_markdown(
            doc=self.doc,
            table_index=table_index,
            markdown_table=markdown,
            preserve_styles=True,
        )

        logger.info(f"Updated GDP table with {len(entries)} quarters from history")

    def _update_gdp_forecast_table(self, table_index: int, forecast: Dict[str, Any]) -> None:
        """
        Update GDP forecast table (Table 3).

        Expected layout: rows for each agency (RBI, World Bank, IMF)
        with columns: Agency | Forecast | Release Date
        """
        if table_index >= len(self.doc.tables):
            logger.warning(f"GDP forecast table index {table_index} out of range")
            return

        table = self.doc.tables[table_index]

        # Map agency keys to rows — skip header row (row 0)
        agency_map = {
            "RBI": 1,
            "World_Bank": 2,
            "IMF": 3,
            "BMI_Fitch": 4,
            "S&P": 5,
        }

        for agency_key, row_idx in agency_map.items():
            agency_data = forecast.get(agency_key, {})
            if not agency_data or row_idx >= len(table.rows):
                continue

            row = table.rows[row_idx]
            # Col 1: forecast value — template already has the correct formatted text
            # (e.g. "6.5% (FY26)"). Extraction returns a raw number which loses the
            # year context, so we only overwrite if the cell is currently blank.
            if len(row.cells) > 1 and agency_data.get("forecast") is not None:
                existing = row.cells[1].text.strip()
                if not existing:
                    row.cells[1].text = str(agency_data["forecast"])
            # Col 2: release date — always update; this changes every month
            if len(row.cells) > 2 and agency_data.get("release_date"):
                row.cells[2].text = str(agency_data["release_date"])

        logger.info("Updated GDP forecast table")

    def _update_satat_table(self, table_index: int, satat: Dict[str, Any]) -> None:
        """
        Update SATAT table (Table 4).

        Expected layout: 2-column table (Particulars | Details).
        Maps dict keys to row labels using specific pattern matching.
        """
        import re as _re

        if table_index >= len(self.doc.tables):
            logger.warning(f"SATAT table index {table_index} out of range")
            return

        table = self.doc.tables[table_index]

        for row in table.rows[1:]:
            if len(row.cells) < 2:
                continue
            label = row.cells[0].text.strip().lower()
            val = None

            if "eoi" in label:
                val = satat.get("eoi_received") or satat.get("eois_received")
            elif "loi" in label:
                val = satat.get("loi_issued") or satat.get("lois_issued")
            elif "financial closure" in label:
                val = satat.get("financial_closure")
            elif "commissioned" in label and "under construction" not in label:
                val = satat.get("plants_commissioned") or satat.get("cbg_plants_commissioned")
            elif "under construction" in label:
                val = satat.get("plants_under_construction")
            elif "offtake" in label:
                val = satat.get("cbg_offtake_tpd")
            elif "from ros" in label or "from r.o" in label:
                # Uniquely matches "Cumulative CBG sale from ROs" row only
                val = satat.get("cumulative_cbg_sale_ros")
            elif "synchronization" in label and _re.search(r'\b20\d\d-\d\d\b', label):
                # "During 2025-26" pattern → fiscal year row
                val = satat.get("cbg_sale_fy")
            elif "synchronization" in label:
                # "During August'25" pattern → current month row
                val = satat.get("cbg_sale_current_month")

            if val is not None:
                row.cells[1].text = str(val)

        logger.info("Updated SATAT table")

    def _update_retail_lng_table(self, table_index: int, retail_lng_data: Any) -> None:
        """
        Update Retail LNG table (template2.docx Table 3).

        Table layout:
          Row 0: header  [Stations, JV/Subsidiary, No. of Stations, Locations]
          Row 1-3: data  rows per station type
          Row 4: total   [empty, empty, Total: 12, empty]

        Accepts:
          - str: markdown table → update_table_from_markdown directly
          - list of dicts: row-by-row fill
            Each dict: {station_type, jv, count, locations}
        """
        if table_index >= len(self.doc.tables):
            logger.warning(f"Retail LNG table index {table_index} out of range")
            return

        if isinstance(retail_lng_data, str):
            update_table_from_markdown(
                self.doc,
                table_index,
                retail_lng_data,
                preserve_styles=True,
            )
            set_table_column_widths(
                self.doc.tables[table_index],
                [1.8, 1.5, 0.8, 5.1],
            )
            for row in self.doc.tables[table_index].rows:
                row.height = None
                row.height_rule = None
            logger.info("Updated Retail LNG table from markdown string")
            return

        if not isinstance(retail_lng_data, list):
            logger.warning(f"Unexpected retail_lng_data type: {type(retail_lng_data)}, skipping")
            return

        table = self.doc.tables[table_index]

        # Fill data rows starting at row 1 (row 0 is header)
        data_row_idx = 1
        for entry in retail_lng_data:
            if data_row_idx >= len(table.rows):
                break
            row = table.rows[data_row_idx]
            if len(row.cells) < 4:
                data_row_idx += 1
                continue

            # Support both extraction schema (type/entities/locations) and
            # legacy schema (station_type/jv/count/locations)
            station_type = entry.get("type") or entry.get("station_type") or ""
            entities = entry.get("entities") or ([entry["jv"]] if "jv" in entry else [])
            if isinstance(entities, str):
                entities = [entities]
            locations = entry.get("locations") or ([entry.get("locations", "")] if "locations" in entry else [])
            if isinstance(locations, str):
                locations = [locations]

            row.cells[0].text = str(station_type)
            row.cells[1].text = ", ".join(entities) if entities else ""
            row.cells[2].text = str(len(entities)) if entities else (str(entry.get("count", "")) if "count" in entry else "")
            row.cells[3].text = "\n".join(locations) if locations else ""
            data_row_idx += 1

        logger.info(f"Updated Retail LNG table with {len(retail_lng_data)} rows")

    def _update_cgd_table(self, table_index: int, cgd_data: Any) -> None:
        """
        Update CGD entity table (Table 5).

        Accepts either a list of entity dicts or a dict with 'entities' key.
        """
        if table_index >= len(self.doc.tables):
            logger.warning(f"CGD table index {table_index} out of range")
            return

        table = self.doc.tables[table_index]

        # Normalize to list of entities
        if isinstance(cgd_data, dict):
            entities = cgd_data.get("entities", [])
        elif isinstance(cgd_data, list):
            entities = cgd_data
        else:
            logger.warning(f"Unexpected cgd_data type: {type(cgd_data)}")
            return

        # Match entities to rows by entity name
        for entity in entities:
            entity_name = entity.get("entity", "").strip().upper()
            if not entity_name:
                continue

            for row in table.rows:
                if len(row.cells) < 2:
                    continue
                row_label = row.cells[0].text.strip().upper()
                if entity_name in row_label or row_label in entity_name:
                    # Fill available columns — keys match extraction schema
                    col_keys = ["dpng_target", "dpng_actual", "cng_target",
                                "cng_actual", "industrial_png", "commercial_png"]
                    for col_idx, key in enumerate(col_keys, start=1):
                        if col_idx < len(row.cells) and key in entity:
                            val = entity[key]
                            row.cells[col_idx].text = str(val) if val is not None else ""
                    break

        logger.info(f"Updated CGD table with {len(entities)} entities")

    def _update_cgd_cumulative_table(self, table_index: int, cgd_cumulative_data: Any) -> None:
        """
        Update CGD cumulative table (metric-row, entity-column layout).

        Uses dpng_cumulative and cng_cumulative keys (distinct from the targets table
        which uses dpng_target/actual, cng_target/actual).
        """
        if table_index >= len(self.doc.tables):
            logger.warning(f"CGD cumulative table index {table_index} out of range")
            return

        table = self.doc.tables[table_index]

        if isinstance(cgd_cumulative_data, dict):
            entities = cgd_cumulative_data.get("entities", [])
        elif isinstance(cgd_cumulative_data, list):
            entities = cgd_cumulative_data
        else:
            logger.warning(f"Unexpected cgd_cumulative_data type: {type(cgd_cumulative_data)}")
            return

        col_keys = ["dpng_cumulative", "cng_cumulative"]
        for entity in entities:
            entity_name = entity.get("entity", "").strip().upper()
            if not entity_name:
                continue
            for row in table.rows:
                if len(row.cells) < 2:
                    continue
                row_label = row.cells[0].text.strip().upper()
                if entity_name in row_label or row_label in entity_name:
                    for col_idx, key in enumerate(col_keys, start=1):
                        if col_idx < len(row.cells) and key in entity:
                            val = entity[key]
                            row.cells[col_idx].text = str(val) if val is not None else ""
                    break

        logger.info(f"Updated CGD cumulative table with {len(entities)} entities")

    def _update_gdp_table_from_dict(self, table_index: int, gdp: Dict[str, Any]) -> None:
        """
        Update the GDP table by appending a new quarter column (shifting oldest out).

        The template GDP table layout:
          Row 0: headers   [GDP, Q1-FY23, Q2-FY23, ...]
          Row 1: Rs. Cr    [GDP (Rs. Cr), val, val, ...]
          Row 2: USD Bn    [GDP (USD Billion), val, val, ...]
          Row 3: % Growth  [% Growth (y-o-y), val, val, ...]
        """
        if table_index >= len(self.doc.tables):
            logger.warning(f"GDP table index {table_index} out of range")
            return

        table = self.doc.tables[table_index]
        quarter = gdp.get("quarter", "")
        gdp_rs_cr = gdp.get("gdp_rs_cr")
        gdp_usd_bn = gdp.get("gdp_usd_bn")
        growth_pct = gdp.get("growth_pct")

        if not quarter:
            logger.warning("No quarter in GDP data, skipping table update")
            return

        # Check if this quarter already exists in headers
        header_row = table.rows[0]
        existing_headers = [cell.text.strip() for cell in header_row.cells]

        if quarter in existing_headers:
            # Update existing column
            col_idx = existing_headers.index(quarter)
        else:
            # Shift left: drop column 1 (oldest), append new column at end
            # For simplicity, just update the last column with the new data
            col_idx = len(header_row.cells) - 1
            header_row.cells[col_idx].text = quarter

        # Update values
        def fmt_num(val, fmt_str="{:,.0f}"):
            return fmt_str.format(val) if val is not None else ""

        if len(table.rows) > 1 and gdp_rs_cr is not None:
            table.rows[1].cells[col_idx].text = fmt_num(gdp_rs_cr)
        if len(table.rows) > 2 and growth_pct is not None:
            table.rows[2].cells[col_idx].text = f"{round(float(growth_pct), 2)}%"
        if len(table.rows) > 3 and gdp_usd_bn is not None:
            table.rows[3].cells[col_idx].text = fmt_num(gdp_usd_bn, "{:.2f}")

        logger.info(f"Updated GDP table column {col_idx} with {quarter} data")

    def _apply_chart_updates(
        self, data: Dict[str, Any], template_bytes: bytes
    ) -> Tuple[bytes, List[str]]:
        """
        Apply native OOXML chart data updates to the template bytes.

        Operates at the ZIP byte level: parses each chart XML, rebuilds the
        data caches with new values, and disables the external data link.
        Returns the modified DOCX bytes and a list of chart names updated.
        """
        chart_updates: List[dict] = []
        generated: List[str] = []
        template_type = self.config.template_type
        chart_indices = self.CHART_INDICES.get(template_type, self.CHART_INDICES["12"])
        last_n = 23

        # Inflation chart (chart1.xml)
        if "inflation" in data and data["inflation"]:
            try:
                chart_updates.append(make_inflation_update(data["inflation"], last_n=18))
                generated.append("inflation")
            except Exception as e:
                logger.error(f"Failed to build inflation chart update: {e}")

        # Interest rates chart (chart2.xml)
        if "interest_rates" in data and data["interest_rates"]:
            try:
                chart_updates.append(make_interest_rates_update(data["interest_rates"], last_n=22))
                generated.append("interest_rates")
            except Exception as e:
                logger.error(f"Failed to build interest rates chart update: {e}")

        # Gas yearly chart (chart3.xml) — only in 12-month template
        if "gas_yearly" in chart_indices and "gas_yearly" in data and data["gas_yearly"]:
            try:
                chart_updates.append(make_gas_yearly_update(data["gas_yearly"]))
                generated.append("gas_yearly")
            except Exception as e:
                logger.error(f"Failed to build gas yearly chart update: {e}")

        # Gas monthly chart (chart4.xml) — only in 12-month template
        if "gas_monthly" in chart_indices and "gas_monthly" in data and data["gas_monthly"]:
            try:
                chart_updates.append(make_gas_monthly_update(data["gas_monthly"]))
                generated.append("gas_monthly")
            except Exception as e:
                logger.error(f"Failed to build gas monthly chart update: {e}")

        # Share price chart — file depends on template type (index+1 → chart{n}.xml)
        if "share_price" in chart_indices and "share_price" in data and data["share_price"]:
            try:
                share_chart_file = f"word/charts/chart{chart_indices['share_price'] + 1}.xml"
                chart_updates.append(make_share_price_update(data["share_price"], last_n=last_n, chart_file=share_chart_file))
                generated.append("share_price")
            except Exception as e:
                logger.error(f"Failed to build share price chart update: {e}")

        if chart_updates:
            try:
                template_bytes = update_docx_charts(template_bytes, chart_updates)
                logger.info(f"Applied {len(chart_updates)} native chart update(s): {generated}")
            except Exception as e:
                logger.error(f"Failed to apply chart updates: {e}", exc_info=True)
                generated = []

        return template_bytes, generated

    def _update_news_section(self) -> bool:
        """
        Scrape energy news for the report month, summarize via LLM, and append
        a 'News' section to the document.

        Returns True if at least one news item was added.
        Failures are logged and silently skipped — report generation continues regardless.
        """
        current_month = datetime.now().strftime("%Y-%m")
        try:
            raw_items = scrape_news_for_month(current_month)
        except Exception as e:
            logger.warning("News scraping failed (skipping news section): %s", e)
            return False

        if not raw_items:
            logger.warning("No news items scraped for %s — skipping news section", current_month)
            return False

        try:
            llm_provider = LLMConfig.PROVIDER
            if llm_provider == "openai":
                llm_client = OpenAI(api_key=OpenAIConfig.API_KEY)
                llm_model = OpenAIConfig.LLM_MODEL
            else:
                llm_client = AzureOpenAI(
                    azure_endpoint=AzureConfig.OPENAI_ENDPOINT,
                    api_key=AzureConfig.OPENAI_API_KEY,
                    api_version=AzureConfig.OPENAI_API_VERSION,
                )
                llm_model = AzureConfig.LLM_DEPLOYMENT

            news_items = summarize_news(raw_items, current_month, llm_client, llm_model)
        except Exception as e:
            logger.warning("News summarization failed (skipping news section): %s", e)
            return False

        if not news_items:
            logger.warning("LLM returned no news items for %s — skipping news section", current_month)
            return False

        # Heading — Times New Roman, bold, underlined, 12pt (matches all other sections)
        heading = self.doc.add_paragraph()
        heading.style = self.doc.styles["Normal"]
        h_run = heading.add_run("News Items:")
        h_run.bold = True
        h_run.underline = True
        h_run.font.name = "Times New Roman"
        h_run.font.size = Pt(12)

        # 2-column table: narrow number col (0.5") + wide text col (5.5")
        from docx.oxml.ns import qn as _qn
        from docx.oxml import OxmlElement as _OxmlElement

        COL_WIDTHS = [720, 7920]  # twips: 0.5" + 5.5"

        def _apply_table_widths(tbl, widths):
            """Set fixed column widths via tblGrid + tblLayout + per-cell tcW."""
            tblPr = tbl.find(_qn("w:tblPr"))
            if tblPr is None:
                tblPr = _OxmlElement("w:tblPr")
                tbl.insert(0, tblPr)
            # Fixed layout — disables Word's autofit
            tblLayout = tblPr.find(_qn("w:tblLayout"))
            if tblLayout is None:
                tblLayout = _OxmlElement("w:tblLayout")
                tblPr.append(tblLayout)
            tblLayout.set(_qn("w:type"), "fixed")
            # Total table width
            tblW = tblPr.find(_qn("w:tblW"))
            if tblW is None:
                tblW = _OxmlElement("w:tblW")
                tblPr.append(tblW)
            tblW.set(_qn("w:w"), str(sum(widths)))
            tblW.set(_qn("w:type"), "dxa")
            # tblGrid — defines column widths
            old_grid = tbl.find(_qn("w:tblGrid"))
            if old_grid is not None:
                tbl.remove(old_grid)
            tblGrid = _OxmlElement("w:tblGrid")
            for w in widths:
                gc = _OxmlElement("w:gridCol")
                gc.set(_qn("w:w"), str(w))
                tblGrid.append(gc)
            tbl.insert(list(tbl).index(tblPr) + 1, tblGrid)
            # Per-cell tcW
            for row in tbl.findall(_qn("w:tr")):
                for i, tc in enumerate(row.findall(_qn("w:tc"))):
                    tcPr = tc.find(_qn("w:tcPr"))
                    if tcPr is None:
                        tcPr = _OxmlElement("w:tcPr")
                        tc.insert(0, tcPr)
                    for old in tcPr.findall(_qn("w:tcW")):
                        tcPr.remove(old)
                    tcW = _OxmlElement("w:tcW")
                    tcW.set(_qn("w:w"), str(widths[i]))
                    tcW.set(_qn("w:type"), "dxa")
                    tcPr.append(tcW)

        rows = [item for item in news_items if item.get("text", "").strip()]
        table = self.doc.add_table(rows=len(rows), cols=2)
        table.style = "Table Grid"
        _apply_table_widths(table._tbl, COL_WIDTHS)

        for row_idx, item in enumerate(rows):
            cells = table.rows[row_idx].cells
            cells[0].text = str(row_idx + 1)
            cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            # Text + source on a new line in the same cell
            text = _sanitize_xml(item.get("text", "").strip())
            source = _sanitize_xml(item.get("source", "").strip())
            cells[1].text = text
            if source:
                src_para = cells[1].add_paragraph()
                src_run = src_para.add_run(f"Source: {source}")
                src_run.italic = True
                src_run.font.name = "Times New Roman"
                src_run.font.size = Pt(10)
            for cell in cells:
                for para in cell.paragraphs:
                    for r in para.runs:
                        if not r.font.name:
                            r.font.name = "Times New Roman"
                        if not r.font.size:
                            r.font.size = Pt(12)

        logger.info("Added News section with %d items to report", len(news_items))
        return True

    def _save_document(self) -> Tuple[Path, Optional[Path]]:
        """Save the document as DOCX (and optionally PDF). Returns (docx_path, pdf_path)."""
        if self.config.output_path:
            output_path = self.config.output_path
        else:
            today = datetime.now().strftime("%Y_%m_%d")
            filename = f"CMD_Report_{self.config.report_month}_{today}.docx"
            output_path = self.config.temp_dir / filename

        self.doc.save(output_path)
        logger.info(f"Saved report to: {output_path}")

        pdf_path: Optional[Path] = None
        if self.config.generate_pdf:
            pdf_out = output_path.with_suffix(".pdf")
            pdf_path = convert_docx_to_pdf(output_path, pdf_out)

        return output_path, pdf_path


def generate_report(
    template_path: Union[str, Path],
    extraction_results: Dict[str, Any],
    report_month: str,
    template_type: str = "12",
    output_path: Optional[Union[str, Path]] = None,
    data_dir: Optional[Union[str, Path]] = None,
    strict_validation: bool = False,
    skip_validation: bool = False,
    table_updater: Optional["TableUpdater"] = None,
) -> GenerationResult:
    """
    Convenience function to generate a report.

    Args:
        template_path: Path to the DOCX template
        extraction_results: Dictionary of extraction results
        report_month: Report month in YYYY-MM format
        template_type: "6" or "12" months template
        output_path: Output path for generated report
        data_dir: Directory for historical data
        strict_validation: Treat warnings as errors
        skip_validation: Skip validation step
        table_updater: Optional TableUpdater for ChromaDB fallback when
                       extraction data is missing

    Returns:
        GenerationResult with status and output path
    """
    config = GenerationConfig(
        template_path=Path(template_path),
        output_path=Path(output_path) if output_path else None,
        report_month=report_month,
        template_type=template_type,
        strict_validation=strict_validation,
    )

    generator = ReportGenerator(
        config=config,
        data_dir=Path(data_dir) if data_dir else None,
        table_updater=table_updater,
    )

    return generator.generate(extraction_results, skip_validation=skip_validation)
