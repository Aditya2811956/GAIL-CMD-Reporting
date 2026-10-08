"""
mospi_scraper.py — Official MoSPI API & Multi-Tier Fallback Scraper.

Data Sources:

1. Primary:
   Official MoSPI e-Sankhyiki API
   https://api.mospi.gov.in

   CPI Warehouse:
   https://cpi.mospi.gov.in

2. Secondary:
   Official MoSPI Press Release PDF extraction
   using pdfplumber / pypdfium2 / pypdf

3. Tertiary:
   Persistent local historical time series, where handled
   by the calling generation/service layer.

Scrapes:

    - CPI headline inflation rate (%)
    - IIP General growth rate (%)
    - GDP quarterly growth rate (%)

IMPORTANT:

    CPI and IIP API functions return LISTS of monthly records.

    This is required by the report chart, which needs historical
    monthly observations rather than only the latest value.
"""

import io
import json
import logging
import os
import re
import ssl
from bs4 import BeautifulSoup

from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional
from urllib.parse import urljoin

import requests
import urllib3

from requests.adapters import HTTPAdapter
from urllib3.poolmanager import PoolManager

from dotenv import load_dotenv
import esankhyiki


# ============================================================
# ENVIRONMENT
# ============================================================

load_dotenv()

urllib3.disable_warnings()

logger = logging.getLogger(__name__)


# ============================================================
# MOSPI URLS
# ============================================================

MOSPI_BASE_URL = "https://www.mospi.gov.in/"

MOSPI_HOME_API = (
    "https://www.mospi.gov.in/"
    "api/main-site/get-home-main-site-data?lang=en"
)

# Current MoSPI e-Sankhyiki API base.
#
# CPI:
#   /api/cpi/getCPIIndex
#
# IIP:
#   /api/iip/getIipData
#
MOSPI_API_BASE = "https://api.mospi.gov.in/api"

CPI_WAREHOUSE_BASE = "https://cpi.mospi.gov.in"

# MoSPI revised series transition dates.
# CPI: 2012=100 through December 2025; 2024=100 from January 2026.
# IIP: 2011-12 through March 2026; 2022-23 from April 2026.
CPI_LEGACY_BASE_YEAR = "2012"
CPI_CURRENT_BASE_YEAR = "2024"
CPI_CURRENT_SERIES_START = "2026-01"

IIP_LEGACY_BASE_YEAR = "2011-12"
IIP_CURRENT_BASE_YEAR = "2022-23"
IIP_CURRENT_SERIES_START = "2026-04"


# ============================================================
# MOSPI API TOKEN
# ============================================================

MOSPI_API_TOKEN = os.getenv(
    "MOSPI_API_TOKEN",
    "",
).strip()


# ============================================================
# HTTP HEADERS
# ============================================================

_HEADERS = {
    "User-Agent": (
        "CMD-Backend/1.0 "
        "(MoSPI official data extraction)"
    ),
    "Accept": "application/json",
}

# IMPORTANT:
# Add the token BEFORE creating the API session.
#
# Your previous version created the session before adding
# Authorization, which meant the session would not inherit
# the token.
if MOSPI_API_TOKEN:

    _HEADERS["Authorization"] = (
        f"Bearer {MOSPI_API_TOKEN}"
    )


logger.info(
    "MoSPI API token configured: %s",
    bool(MOSPI_API_TOKEN),
)


# ============================================================
# LEGACY TLS ADAPTER
# ============================================================

class LegacyTLSAdapter(HTTPAdapter):
    """
    HTTPS adapter for MoSPI API servers that require
    legacy TLS server compatibility.

    Some Windows/Python 3.11 + OpenSSL 3 environments
    reject the MoSPI API during the TLS handshake with:

        SSL:
        UNSAFE_LEGACY_RENEGOTIATION_DISABLED

    This adapter enables the OpenSSL legacy-server-connect
    compatibility option where available.
    """

    def init_poolmanager(
        self,
        connections,
        maxsize,
        block=False,
        **pool_kwargs,
    ):
        context = ssl.create_default_context()

        # ----------------------------------------------------
        # Enable legacy server connection compatibility.
        #
        # Python/OpenSSL normally exposes this as:
        #
        #     ssl.OP_LEGACY_SERVER_CONNECT
        #
        # A fallback numeric value is included because
        # OpenSSL defines this option as 0x4.
        # ----------------------------------------------------

        legacy_option = getattr(
            ssl,
            "OP_LEGACY_SERVER_CONNECT",
            0x4,
        )

        try:

            context.options |= legacy_option

        except Exception as exc:

            logger.warning(
                "Could not enable "
                "OP_LEGACY_SERVER_CONNECT: %s",
                exc,
            )

        # ----------------------------------------------------
        # MoSPI's API is being contacted with verify=False
        # in this project because the Windows TLS handshake
        # itself is the compatibility problem.
        # ----------------------------------------------------

        context.check_hostname = False

        context.verify_mode = (
            ssl.CERT_NONE
        )

        # ----------------------------------------------------
        # Create urllib3 pool.
        # ----------------------------------------------------

        self.poolmanager = PoolManager(
            num_pools=connections,
            maxsize=maxsize,
            block=block,
            ssl_context=context,
            **pool_kwargs,
        )


# ============================================================
# MOSPI API SESSION
# ============================================================

def _create_mospi_session() -> requests.Session:
    """
    Create a requests session specifically for the
    api.mospi.gov.in API.

    The session uses the LegacyTLSAdapter so that
    Windows/OpenSSL 3 does not reject the MoSPI server
    during TLS negotiation.
    """

    session = requests.Session()

    session.mount(
        "https://api.mospi.gov.in/",
        LegacyTLSAdapter(),
    )

    session.headers.update(
        _HEADERS
    )

    return session


_MOSPI_API_SESSION = _create_mospi_session()


# ============================================================
# MONTH MAP
# ============================================================

_MONTH_MAP = {

    "january": "01",
    "jan": "01",

    "february": "02",
    "feb": "02",

    "march": "03",
    "mar": "03",

    "april": "04",
    "apr": "04",

    "may": "05",

    "june": "06",
    "jun": "06",

    "july": "07",
    "jul": "07",

    "august": "08",
    "aug": "08",

    "september": "09",
    "sep": "09",
    "sept": "09",

    "october": "10",
    "oct": "10",

    "november": "11",
    "nov": "11",

    "december": "12",
    "dec": "12",
}


# ============================================================
# HELPERS
# ============================================================

def _normalise_month(
    value: Any,
) -> Optional[str]:
    """
    Convert a month representation into MM.

    Handles:

        January
        Jan
        january
        1
        01
    """

    if value is None:
        return None

    value_str = (
        str(value)
        .strip()
        .lower()
    )

    # Month name.
    if value_str in _MONTH_MAP:

        return _MONTH_MAP[
            value_str
        ]

    # Numeric month.
    try:

        month_number = int(
            float(value_str)
        )

        if 1 <= month_number <= 12:

            return f"{month_number:02d}"

    except (
        ValueError,
        TypeError,
    ):

        pass

    return None


def _normalise_year(
    value: Any,
) -> Optional[str]:
    """
    Extract a 4-digit calendar year.
    """

    if value is None:
        return None

    value_str = (
        str(value)
        .strip()
    )

    match = re.search(
        r"\b(20\d{2})\b",
        value_str,
    )

    if match:

        return match.group(1)

    return None


def _extract_numeric(
    row: Dict[str, Any],
    possible_fields: List[str],
) -> Optional[float]:
    """
    Try multiple possible field names and return a float.

    This makes the scraper tolerant of small response-schema
    differences between MoSPI API versions.
    """

    for field in possible_fields:

        value = row.get(
            field
        )

        if value is None:
            continue

        if isinstance(
            value,
            str,
        ):

            value = (
                value
                .strip()
            )

            if not value:
                continue

            value = (
                value
                .replace(
                    "%",
                    "",
                )
                .strip()
            )

        try:

            return float(
                value
            )

        except (
            ValueError,
            TypeError,
        ):

            continue

    return None


# ============================================================
# PDF TEXT EXTRACTION
# ============================================================

def _extract_text_from_pdf_bytes(
    pdf_bytes: bytes,
    max_pages: int = 5,
) -> str:
    """
    Safely extract text from in-memory PDF bytes.

    Tries:

        1. pdfplumber
        2. pypdfium2
        3. pypdf
    """

    # --------------------------------------------------------
    # pdfplumber
    # --------------------------------------------------------

    try:

        import pdfplumber

        with pdfplumber.open(
            io.BytesIO(
                pdf_bytes
            )
        ) as pdf:

            pages = [
                page.extract_text()
                or ""
                for page in pdf.pages[
                    :max_pages
                ]
            ]

            text = "\n".join(
                pages
            )

            if text.strip():

                return text

    except Exception as exc:

        logger.debug(
            "pdfplumber extraction failed: %s",
            exc,
        )

    # --------------------------------------------------------
    # pypdfium2
    # --------------------------------------------------------

    try:

        import pypdfium2

        pdf = pypdfium2.PdfDocument(
            pdf_bytes
        )

        pages = []

        for i in range(
            min(
                max_pages,
                len(pdf),
            )
        ):

            text_page = (
                pdf[i]
                .get_textpage()
                .get_text_range()
            )

            pages.append(
                text_page
            )

        text = "\n".join(
            pages
        )

        if text.strip():

            return text

    except Exception as exc:

        logger.debug(
            "pypdfium2 extraction failed: %s",
            exc,
        )

    # --------------------------------------------------------
    # pypdf
    # --------------------------------------------------------

    try:

        import pypdf

        reader = pypdf.PdfReader(
            io.BytesIO(
                pdf_bytes
            )
        )

        pages_text = []

        for i, page in enumerate(
            reader.pages
        ):

            if i >= max_pages:
                break

            pages_text.append(
                page.extract_text()
                or ""
            )

        text = "\n".join(
            pages_text
        )

        if text.strip():

            return text

    except Exception as exc:

        logger.debug(
            "pypdf extraction failed: %s",
            exc,
        )

    return ""


# ============================================================
# MOSPI HOME / RELEASE API
# ============================================================

def _fetch_mospi_releases_list() -> List[Dict[str, Any]]:
    """
    Query MoSPI home REST API to get latest release metadata.
    """

    try:

        resp = requests.get(
            MOSPI_HOME_API,
            headers=_HEADERS,
            timeout=12,
            verify=False,
        )

        logger.info(
            "MoSPI home API status: %s",
            resp.status_code,
        )

        if resp.status_code != 200:

            logger.warning(
                "MoSPI home API returned "
                "status %s",
                resp.status_code,
            )

            return []

        data = resp.json()

        releases = (
            data.get(
                "latestReleasesData"
            )
            or []
        )

        announcements = (
            data.get(
                "announcementData"
            )
            or []
        )

        return (
            releases
            + announcements
        )

    except Exception:

        logger.exception(
            "Failed to query MoSPI home API"
        )

        # Fall through to latest published month when exact report month is unavailable.


# ============================================================
# IIP API
# ============================================================

def fetch_iip_api(
    year: Optional[str] = None,
    month_code: Optional[str] = None,
    base_year: str = "2022-23",
) -> List[Dict[str, Any]]:
    """
    Fetch monthly General IIP growth from MoSPI.

    Default base year:
        2022-23

    Returns:

        [
            {
                "month": "2025-08",
                "IIP": 4.3,
                "source": "MoSPI IIP API",
                "base_year": "2022-23"
            },
            ...
        ]

    IMPORTANT:

        This function preserves all returned monthly records.
        It does NOT use rows[-1].
    """

    url = (
        f"{MOSPI_API_BASE}"
        "/iip/getIipData"
    )

    params = {

        "base_year":
            base_year,

        "frequency":
            "Monthly",

        "type":
            "General",

        "Format":
            "JSON",

        "limit":
            100,
    }

    if year:

        params["year"] = year

    if month_code:

        params[
            "month_code"
        ] = month_code

    logger.info(
        "========== IIP API REQUEST =========="
    )

    logger.info(
        "IIP API parameters: %s",
        params,
    )

    try:

        # IMPORTANT:
        # Use the special MoSPI session rather than
        # requests.get() so the TLS compatibility
        # adapter is actually used.
        logger.info(
    "MoSPI API request headers: %s",
    {
        key: (
            "***"
            if key.lower() == "authorization"
            else value
        )
        for key, value
        in _MOSPI_API_SESSION.headers.items()
    },
)

        resp = _MOSPI_API_SESSION.get(
            url,
            params=params,
            timeout=20,
            verify=False,
        )

        logger.info(
            "IIP API URL: %s",
            resp.url,
        )

        logger.info(
            "IIP API status: %s",
            resp.status_code,
        )

        if resp.status_code != 200:

            logger.error(
                "IIP API response: %s",
                resp.text[:3000],
            )

            return []

        data = resp.json()

        rows = (
            data.get("data")
            or []
        )

        logger.info(
            "IIP API returned %d raw rows",
            len(rows),
        )

        logger.info(
            "IIP raw sample: %s",
            rows[:3],
        )

        results = []

        for row in rows:

            # ------------------------------------------------
            # Find growth value.
            # ------------------------------------------------

            growth = _extract_numeric(
                row,
                [
                    "growth_rate",
                    "growth",
                    "growthRate",
                    "iip_growth",
                    "growth_yoy",
                    "growth_rate_yoy",
                ],
            )

            if growth is None:

                logger.debug(
                    "Skipping IIP row with no "
                    "growth value: %s",
                    row,
                )

                continue

            # ------------------------------------------------
            # Find year.
            # ------------------------------------------------

            row_year = (
                _normalise_year(
                    row.get("year")
                )
                or _normalise_year(
                    row.get("Year")
                )
                or _normalise_year(
                    year
                )
            )

            # ------------------------------------------------
            # Find month.
            # ------------------------------------------------

            row_month = _normalise_month(
                row.get("month")
                or row.get("month_name")
                or row.get("month_code")
                or row.get("Month")
                or row.get("Month_Name")
            )

            if (
                not row_year
                or not row_month
            ):

                logger.debug(
                    "Skipping IIP row with "
                    "invalid period: %s",
                    row,
                )

                continue

            # ------------------------------------------------
            # Normalized result.
            # ------------------------------------------------

            results.append(
                {
                    "month":
                        f"{row_year}-{row_month}",

                    "IIP":
                        growth,

                    "source":
                        "MoSPI IIP API",

                    "base_year":
                        base_year,
                }
            )

        # ----------------------------------------------------
        # Remove duplicate months.
        # ----------------------------------------------------

        unique = {}

        for row in results:

            unique[
                row["month"]
            ] = row

        results = list(
            unique.values()
        )

        # ----------------------------------------------------
        # Chronological order.
        # ----------------------------------------------------

        results.sort(
            key=lambda x:
                x["month"]
        )

        logger.info(
            "Normalized IIP rows: %d",
            len(results),
        )

        logger.info(
            "Normalized IIP data: %s",
            results,
        )

        return results

    except requests.exceptions.SSLError:

        logger.exception(
            "IIP API SSL/TLS extraction failed"
        )

        # Fall through to latest published month when exact report month is unavailable.

    except requests.exceptions.RequestException:

        logger.exception(
            "IIP API network request failed"
        )

        # Fall through to latest published month when exact report month is unavailable.

    except ValueError:

        logger.exception(
            "IIP API returned invalid JSON"
        )

        # Fall through to latest published month when exact report month is unavailable.

    except Exception:

        logger.exception(
            "IIP API extraction failed"
        )

        # Fall through to latest published month when exact report month is unavailable.


def _normalise_target_month(value: Optional[str]) -> Optional[str]:
    """Return a validated YYYY-MM month string or None."""
    if not value:
        return None
    try:
        return datetime.strptime(str(value).strip(), "%Y-%m").strftime("%Y-%m")
    except (TypeError, ValueError):
        return None


def _get_cpi_base_year(month: str) -> str:
    """Return the CPI series base applicable to a reference month."""
    return (
        CPI_CURRENT_BASE_YEAR
        if str(month) >= CPI_CURRENT_SERIES_START
        else CPI_LEGACY_BASE_YEAR
    )


def _get_iip_base_year(month: str) -> str:
    """Return the IIP series base applicable to a reference month."""
    return (
        IIP_CURRENT_BASE_YEAR
        if str(month) >= IIP_CURRENT_SERIES_START
        else IIP_LEGACY_BASE_YEAR
    )


def _extract_release_pdf_url(item: Dict[str, Any]) -> Optional[str]:
    """Extract a PDF URL from a MoSPI release/announcement object."""
    file_info = item.get("file_one") or {}
    file_path = file_info.get("path")
    if file_path:
        return urljoin(MOSPI_BASE_URL, str(file_path))

    for key in ("file", "pdf", "pdf_url", "document", "url", "link"):
        value = item.get(key)
        if isinstance(value, dict):
            value = value.get("path") or value.get("url") or value.get("href")
        if value and ".pdf" in str(value).lower():
            return urljoin(MOSPI_BASE_URL, str(value))

    return None


def _find_cpi_release_pdf(target_year: str, target_month: str) -> Optional[Dict[str, Any]]:
    """
    Find the official MoSPI CPI release for the EXACT requested month.

    The homepage release feed is only a latest-release feed, so it is not
    treated as a historical database. We first search its metadata for an
    exact month match and then use the deterministic MoSPI press-release
    page slug as the historical fallback.
    """
    target = f"{target_year}-{int(target_month):02d}"
    month_name = datetime.strptime(target, "%Y-%m").strftime("%B")

    # 1. Search the official home/release API, but only accept an exact month.
    try:
        releases = _fetch_mospi_releases_list()
        candidates = []
        for item in releases:
            title = str(item.get("title") or item.get("heading") or "")
            haystack = title.lower()
            if "cpi" not in haystack and "consumer price index" not in haystack:
                continue
            if month_name.lower() not in haystack or str(target_year) not in haystack:
                continue
            pdf_url = _extract_release_pdf_url(item)
            if pdf_url:
                candidates.append({"title": title, "pdf_url": pdf_url, "source": "MoSPI release API"})
        if candidates:
            logger.info("Found exact CPI release in MoSPI release feed: %s", candidates[0])
            return candidates[0]
    except Exception:
        logger.exception("MoSPI release-feed CPI lookup failed")

    # 2. Historical press-release page. The current MoSPI site uses this
    # deterministic pattern for CPI monthly release pages.
    slug = f"press-release-cpi-{month_name.lower()}-{target_year}"
    page_url = urljoin(MOSPI_BASE_URL, f"press-release/{slug}")

    try:
        response = requests.get(
            page_url,
            headers=_HEADERS,
            timeout=20,
            verify=False,
        )
        logger.info("CPI release page status=%s url=%s", response.status_code, page_url)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, "html.parser")
            for anchor in soup.find_all("a", href=True):
                href = str(anchor.get("href"))
                text = anchor.get_text(" ", strip=True)
                haystack = f"{text} {href}".lower()
                if ".pdf" in href.lower() and (
                    "cpi" in haystack
                    or "consumer price" in haystack
                    or "press" in haystack
                    or month_name.lower() in haystack
                ):
                    pdf_url = urljoin(page_url, href)
                    return {
                        "title": text or f"CPI Press Release {month_name} {target_year}",
                        "pdf_url": pdf_url,
                        "source": "MoSPI press-release page",
                    }
    except Exception:
        logger.exception("Historical CPI press-release page lookup failed: %s", page_url)

    # 3. CPI warehouse press-PDF naming convention. This is deliberately a
    # candidate only; it is accepted only after an HTTP 200 response.
    warehouse_pdf = (
        f"{CPI_WAREHOUSE_BASE}/PDFile/Press/"
        f"PR%20{month_name}%20{target_year}.pdf"
    )
    try:
        response = requests.get(
            warehouse_pdf,
            headers=_HEADERS,
            timeout=20,
            verify=False,
            allow_redirects=True,
        )
        content_type = response.headers.get("Content-Type", "").lower()
        if response.status_code == 200 and (
            response.content[:4] == b"%PDF" or "pdf" in content_type
        ):
            return {
                "title": f"CPI Press Release {month_name} {target_year}",
                "pdf_url": response.url,
                "source": "MoSPI CPI Warehouse",
            }
    except Exception:
        logger.exception("CPI warehouse PDF lookup failed: %s", warehouse_pdf)

    logger.warning("No exact official CPI release PDF found for %s", target)
    return None


def _fetch_cpi_release_pdf_value(
    pdf_url: str,
    target_year: str,
    target_month: str,
    base_year: Optional[str] = None,
) -> Optional[Dict[str, Any]]:
    """
    Download an official MoSPI CPI press-release PDF and extract:

        All India
        Combined
        CPI (General)
        YoY inflation %

    This is used as a historical fallback when the 2024-base
    CPI API returns the index but no YoY inflation value.
    """

    logger.info(
        "Downloading official MoSPI CPI release: %s",
        pdf_url,
    )

    try:
        response = _MOSPI_API_SESSION.get(
            pdf_url,
            timeout=30,
            verify=False,
        )

        logger.info(
            "CPI PDF HTTP status: %s",
            response.status_code,
        )

        if response.status_code != 200:
            logger.warning(
                "Unable to download CPI PDF for %s-%s",
                target_year,
                target_month,
            )
            return None

        pdf_bytes = response.content

        if not pdf_bytes:
            logger.warning(
                "CPI PDF was empty for %s-%s",
                target_year,
                target_month,
            )
            return None

        # Extract enough pages to cover the headline/table.
        text = _extract_text_from_pdf_bytes(
            pdf_bytes,
            max_pages=5,
        )

        if not text:
            logger.warning(
                "Could not extract text from CPI PDF "
                "for %s-%s",
                target_year,
                target_month,
            )
            return None

        logger.debug(
            "Extracted CPI PDF text length for %s-%s: %d",
            target_year,
            target_month,
            len(text),
        )

        # Normalize whitespace while preserving the text.
        normalized = re.sub(
            r"[ \t]+",
            " ",
            text,
        )

        normalized = re.sub(
            r"\n+",
            "\n",
            normalized,
        )

        # --------------------------------------------------------
        # Extract headline All-India Combined CPI inflation.
        #
        # MoSPI CPI press releases normally contain two months:
        #
        #   July, 2026 (Provisional) June, 2026 (Final)
        #
        # followed by:
        #
        #   Rural Urban Combined Rural Urban Combined
        #
        # and then:
        #
        #   CPI (General) 4.84 3.96 4.45 4.74 3.93 4.38
        #   Inflation (%)
        #
        # Therefore:
        #
        #   July 2026 -> 4.84 3.96 4.45 -> Combined = 4.45
        #   June 2026 -> 4.74 3.93 4.38 -> Combined = 4.38
        #
        # IMPORTANT:
        # We only return a value when the requested month is
        # actually one of the two months represented by the table.
        #
        # We NEVER return the first three values for an unrelated
        # month. This prevents:
        #
        #   May 2026 -> 4.45   <-- WRONG
        #
        # --------------------------------------------------------

        inflation_value = None

        requested_year_int = int(target_year)
        requested_month_int = int(target_month)

        # --------------------------------------------------------
        # Find the CPI General inflation row.
        # --------------------------------------------------------

        cpi_row_match = re.search(
            r"CPI\s*\(\s*General\s*\)"
            r"\s*"
            r"(?P<values>"
            r"(?:[+-]?\d+(?:\.\d+)?\s+){2,12}"
            r")"
            r"Inflation\s*\(\s*%\s*\)",
            normalized,
            re.IGNORECASE | re.DOTALL,
        )

        if cpi_row_match:

            values_text = cpi_row_match.group(
                "values"
            )

            numeric_values = re.findall(
                r"[+-]?\d+(?:\.\d+)?",
                values_text,
            )

            logger.debug(
                "CPI General row values for %s-%s: %s",
                target_year,
                target_month,
                numeric_values,
            )

            # ----------------------------------------------------
            # We need at least six values:
            #
            # current/release month:
            #   Rural, Urban, Combined
            #
            # previous month:
            #   Rural, Urban, Combined
            # ----------------------------------------------------

            if len(numeric_values) >= 6:

                try:

                    numeric_values = [
                        float(value)
                        for value in numeric_values[:6]
                    ]

                    # ------------------------------------------------
                    # Locate the month headings immediately before
                    # the CPI table.
                    #
                    # Example:
                    #
                    # July, 2026 (Provisional) June, 2026 (Final)
                    # Rural Urban Combined Rural Urban Combined
                    # CPI (General) ...
                    # ------------------------------------------------

                    heading_start = max(
                        0,
                        cpi_row_match.start() - 1500,
                    )

                    heading_text = normalized[
                        heading_start:
                        cpi_row_match.start()
                    ]

                    month_matches = list(
                        re.finditer(
                            r"\b"
                            r"(January|February|March|April|May|June|"
                            r"July|August|September|October|November|December)"
                            r"\s*,?\s*"
                            r"(\d{4})"
                            r"(?:\s*\([^)]*\))?",
                            heading_text,
                            re.IGNORECASE,
                        )
                    )

                    # We only care about the last two month headings
                    # before the CPI table.
                    month_matches = month_matches[-2:]

                    table_months = []

                    for month_match in month_matches:

                        month_name = month_match.group(1)
                        year_value = int(
                            month_match.group(2)
                        )

                        month_number = datetime.strptime(
                            month_name[:3],
                            "%b",
                        ).month

                        table_months.append(
                            (
                                year_value,
                                month_number,
                            )
                        )

                    logger.debug(
                        "CPI table months detected: %s",
                        table_months,
                    )

                    # ------------------------------------------------
                    # Map requested month to the correct group.
                    #
                    # Group 0:
                    #   values[0], values[1], values[2]
                    #
                    # Group 1:
                    #   values[3], values[4], values[5]
                    # ------------------------------------------------

                    selected_group = None

                    for index, (
                        table_year,
                        table_month,
                    ) in enumerate(table_months):

                        if (
                            table_year
                            == requested_year_int
                            and table_month
                            == requested_month_int
                        ):

                            selected_group = index
                            break

                    # ------------------------------------------------
                    # IMPORTANT:
                    #
                    # If the requested month is not actually present
                    # in the table, do NOT guess.
                    # ------------------------------------------------

                    if (
                        selected_group is not None
                        and selected_group < 2
                    ):

                        value_index = (
                            selected_group * 3
                            + 2
                        )

                        candidate = numeric_values[
                            value_index
                        ]

                        if -50 <= candidate <= 100:

                            inflation_value = candidate

                            logger.info(
                                "CPI table SUCCESS: "
                                "%s-%02d -> Combined=%s "
                                "(group=%s)",
                                requested_year_int,
                                requested_month_int,
                                inflation_value,
                                selected_group,
                            )

                    else:

                        logger.debug(
                            "Requested CPI month %s-%02d "
                            "is not represented by this release table",
                            requested_year_int,
                            requested_month_int,
                        )

                except (
                    ValueError,
                    TypeError,
                ):

                    inflation_value = None

        # --------------------------------------------------------
        # Fallback for the release month's explicit headline.
        #
        # Example:
        #
        # Retail inflation based on Consumer Price Index in
        # July, 2026 is 4.45%
        #
        # This fallback is only allowed when the requested month
        # explicitly appears in the headline text.
        # --------------------------------------------------------

        if inflation_value is None:

            headline_pattern = re.compile(
                r"(?:retail\s+inflation|inflation)"
                r".{0,500}?"
                r"(?:in|for)\s+"
                r"(January|February|March|April|May|June|"
                r"July|August|September|October|November|December)"
                r"\s*,?\s*"
                r"(\d{4})"
                r".{0,100}?"
                r"\bis\s+"
                r"([+-]?\d+(?:\.\d+)?)"
                r"\s*%",
                re.IGNORECASE | re.DOTALL,
            )

            for match in headline_pattern.finditer(
                normalized
            ):

                month_name = match.group(1)
                year_value = int(
                    match.group(2)
                )

                month_number = datetime.strptime(
                    month_name[:3],
                    "%b",
                ).month

                if (
                    year_value == requested_year_int
                    and month_number == requested_month_int
                ):

                    try:

                        candidate = float(
                            match.group(3)
                        )

                        if -50 <= candidate <= 100:

                            inflation_value = candidate

                            logger.info(
                                "CPI headline SUCCESS: "
                                "%s-%02d -> %s",
                                requested_year_int,
                                requested_month_int,
                                candidate,
                            )

                            break

                    except (
                        ValueError,
                        TypeError,
                    ):

                        pass

        if inflation_value is None:

            logger.warning(
                "Could not identify Combined CPI inflation "
                "inside PDF for %s-%s",
                target_year,
                target_month,
            )

            # Print useful debugging information.
            logger.debug(
                "CPI PDF beginning for %s-%s:\n%s",
                target_year,
                target_month,
                normalized[:5000],
            )

            return None

        result = {
            "month": (
                f"{target_year}-"
                f"{int(target_month):02d}"
            ),
            "CPI": inflation_value,
            "source": (
                "MoSPI CPI Press Release"
            ),
            "base_year": base_year or _get_cpi_base_year(
                f"{target_year}-{int(target_month):02d}"
            ),
        }

        logger.info(
            "CPI PDF SUCCESS: %s",
            result,
        )

        return result

    except requests.exceptions.RequestException:

        logger.exception(
            "Network error downloading CPI PDF "
            "for %s-%s",
            target_year,
            target_month,
        )

    except Exception:

        logger.exception(
            "Unexpected CPI PDF extraction error "
            "for %s-%s",
            target_year,
            target_month,
        )

    return None


def fetch_cpi_api(
    year: Optional[str] = None,
    month_code: Optional[str] = None,
    base_year: str = CPI_LEGACY_BASE_YEAR,
) -> List[Dict[str, Any]]:
    """
    Fetch CPI inflation from the official MoSPI eSankhyiki API.

    Supports:
    - 2012 base: getCPIData + group/subgroup schema
    - 2024 base: getCPIData + division schema
    """

    try:
        import esankhyiki

        if not year or not month_code:
            return []

        params = {
            "base_year": str(base_year),
            "year": str(year),
            "month_code": int(month_code),
            "series": "Current",
        }

        if str(base_year) == CPI_LEGACY_BASE_YEAR:
            params["limit"] = 100

        data = esankhyiki.get_data(
            "CPI",
            params,
            format="df",
        )

        if data is None or data.empty:
            logger.warning(
                "No CPI data returned for %s-%02d (base=%s)",
                year,
                int(month_code),
                base_year,
            )
            return []

        # --------------------------------------------------------
        # 2024 BASE — NEW CPI SCHEMA
        # --------------------------------------------------------
        if str(base_year) == CPI_CURRENT_BASE_YEAR:
            target = data[
                (data["state"].astype(str).str.strip().str.lower() == "all india")
                & (
                    data["sector"].astype(str).str.strip().str.lower()
                    == "combined"
                )
                & (
                    data["division"].astype(str).str.strip().str.lower()
                    == "cpi (general)"
                )
            ]

        # --------------------------------------------------------
        # 2012 BASE — LEGACY CPI SCHEMA
        # --------------------------------------------------------
        else:
            target = data[
                (data["state"].astype(str).str.strip().str.lower() == "all india")
                & (
                    data["sector"].astype(str).str.strip().str.lower()
                    == "combined"
                )
                & (
                    data["group"].astype(str).str.strip().str.lower()
                    == "general"
                )
                & (
                    data["subgroup"].astype(str).str.strip().str.lower()
                    == "general-overall"
                )
            ]

        if target.empty:
            logger.warning(
                "No headline CPI row found for %s-%02d (base=%s)",
                year,
                int(month_code),
                base_year,
            )
            return []

        row = target.iloc[0]

        inflation_value = float(row["inflation"])

        if inflation_value is None:
            logger.warning(
                "No CPI inflation value found for %s-%02d (base=%s)",
                year,
                int(month_code),
                base_year,
            )
            return []

        month_padded = f"{int(month_code):02d}"

        return [
            {
                "month": f"{year}-{month_padded}",
                "CPI": inflation_value,
                "source": "MoSPI CPI API",
                "base_year": str(base_year),
            }
        ]

    except Exception:
        logger.exception(
            "CPI API extraction failed for %s-%s (base=%s)",
            year,
            month_code,
            base_year,
        )
        # Fall through to latest published month when exact report month is unavailable.

def fetch_cpi_from_homepage() -> Optional[Dict[str, Any]]:
    """
    Fetch the latest CPI inflation value displayed on the official
    MoSPI homepage banner.

    Returns:
        {
            "month": "YYYY-MM",
            "CPI": float,
            "source": "MoSPI Homepage API"
        }
        or None if unavailable.
    """
    try:
        url = "https://www.mospi.gov.in/api/main-site/get-home-main-site-data?lang=en"

        resp = requests.get(
            url,
            headers=_HEADERS,
            timeout=20,
            verify=False,
        )

        if resp.status_code != 200:
            logger.warning(
                "MoSPI homepage API returned status %s",
                resp.status_code,
            )
            return None

        data = resp.json()

        for item in data.get("banner", []):
            heading = BeautifulSoup(
                item.get("stats_heading", ""),
                "html.parser",
            ).get_text(" ", strip=True)

            # Identify CPI banner
            if (
                "inflation rate" in heading.lower()
                and "consumer price index" in (
                    item.get("tool_tip", "").lower()
                )
            ):
                value_text = str(
                    item.get("stats_percentage", "")
                ).replace("%", "").strip()

                value = float(value_text)

                match = re.search(
                    r"\((January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})\)",
                    heading,
                    re.IGNORECASE,
                )

                if not match:
                    logger.warning(
                        "Could not determine CPI month from homepage banner: %s",
                        heading,
                    )
                    return None

                month_name = match.group(1).lower()
                year = match.group(2)

                month_map = {
                    "january": "01",
                    "february": "02",
                    "march": "03",
                    "april": "04",
                    "may": "05",
                    "june": "06",
                    "july": "07",
                    "august": "08",
                    "september": "09",
                    "october": "10",
                    "november": "11",
                    "december": "12",
                }

                month = f"{year}-{month_map[month_name]}"

                result = {
                    "month": month,
                    "CPI": value,
                    "source": "MoSPI Homepage API",
                    "base_year": "2012",
                }

                logger.info(
                    "Successfully extracted current CPI from MoSPI homepage: %s",
                    result,
                )

                return result

        logger.warning(
            "CPI banner not found on MoSPI homepage"
        )
        return None

    except Exception as exc:
        logger.warning(
            "MoSPI homepage CPI extraction failed: %s",
            exc,
        )
        return None


def fetch_iip_from_homepage() -> Optional[Dict[str, Any]]:
    """
    Fetch the latest IIP growth value displayed on the official
    MoSPI homepage banner.

    Returns:
        {
            "month": "YYYY-MM",
            "IIP": float,
            "source": "MoSPI Homepage API",
            "base_year": "2022-23"
        }
        or None if unavailable.
    """
    try:
        resp = requests.get(
            MOSPI_HOME_API,
            headers=_HEADERS,
            timeout=20,
            verify=False,
        )

        if resp.status_code != 200:
            logger.warning(
                "MoSPI homepage API returned status %s while fetching IIP",
                resp.status_code,
            )
            return None

        data = resp.json()

        for item in data.get("banner", []):
            heading = BeautifulSoup(
                item.get("stats_heading", ""),
                "html.parser",
            ).get_text(" ", strip=True)

            # Identify the IIP banner by its heading/link/tooltip.
            haystack = " ".join(
                [
                    heading,
                    str(item.get("tool_tip", "")),
                    str(item.get("redirectional_link", "")),
                ]
            ).lower()

            if "index of industrial production" not in haystack:
                continue

            value_text = str(
                item.get("stats_percentage", "")
            ).replace("%", "").strip()

            try:
                value = float(value_text)
            except (TypeError, ValueError):
                logger.warning(
                    "Invalid IIP value on homepage banner: %r",
                    item.get("stats_percentage"),
                )
                continue

            match = re.search(
                r"\((January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})\)",
                heading,
                re.IGNORECASE,
            )

            if not match:
                logger.warning(
                    "Could not determine IIP month from homepage banner: %s",
                    heading,
                )
                continue

            month_name = match.group(1).lower()
            year = match.group(2)
            month = f"{year}-{_MONTH_MAP[month_name]}"

            result = {
                "month": month,
                "IIP": value,
                "source": "MoSPI Homepage API",
                "base_year": "2022-23",
            }

            logger.info(
                "Successfully extracted current IIP from MoSPI homepage: %s",
                result,
            )
            return result

        logger.warning("IIP banner not found on MoSPI homepage")
        return None

    except Exception as exc:
        logger.warning(
            "MoSPI homepage IIP extraction failed: %s",
            exc,
        )
        return None

def _log_monthly_coverage(
    rows: List[Dict[str, Any]],
    start_month: Optional[str] = None,
    end_month: Optional[str] = None,
    label: str = "series",
) -> None:
    """Log missing months without inventing/interpolating observations."""
    months = sorted(
        str(row.get("month"))
        for row in rows
        if re.fullmatch(r"20\d{2}-(0[1-9]|1[0-2])", str(row.get("month", "")))
    )
    if not months:
        logger.warning("%s coverage: no valid YYYY-MM rows", label)
        return

    if not start_month:
        start_month = months[0]
    if not end_month:
        end_month = months[-1]

    try:
        start = datetime.strptime(start_month, "%Y-%m")
        end = datetime.strptime(end_month, "%Y-%m")
    except ValueError:
        return

    expected = []
    cursor = start
    while cursor <= end:
        expected.append(cursor.strftime("%Y-%m"))
        if cursor.month == 12:
            cursor = cursor.replace(year=cursor.year + 1, month=1)
        else:
            cursor = cursor.replace(month=cursor.month + 1)

    actual = set(months)
    missing = [month for month in expected if month not in actual]
    if missing:
        logger.warning(
            "%s coverage is incomplete: expected=%d received=%d missing=%s",
            label,
            len(expected),
            len(actual),
            missing,
        )
    else:
        logger.info(
            "%s coverage complete: %d monthly observations (%s -> %s)",
            label,
            len(expected),
            expected[0],
            expected[-1],
        )


# ============================================================
# CPI SCRAPING
# ============================================================

def fetch_cpi_data(
    report_month: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """
    Fetch All-India Combined General CPI YoY inflation.

    CPI series transition:
        - 2012=100 through December 2025
        - 2024=100 from January 2026

    The CPI graph requires historical monthly inflation values.
    Therefore, every required month is fetched directly from the
    official MoSPI e-Sankhyiki API using the correct base year.

    For example:
        2024-11 -> 2012 base
        2024-12 -> 2012 base
        2025-01 -> 2012 base
        ...
        2025-12 -> 2012 base
        2026-01 -> 2024 base
        2026-02 -> 2024 base
        ...
        2026-08 -> 2024 base
    """

    logger.info(
        "========== CPI EXTRACTION START =========="
    )

    requested_month = _normalise_target_month(
        report_month
    )

    # ============================================================
    # 1. EXACT REPORT MONTH
    # ============================================================

    if requested_month:

        logger.info(
            "Fetching CPI for exact report month: %s",
            requested_month,
        )

        try:
            dt = datetime.strptime(
                requested_month,
                "%Y-%m",
            )

            base_year = _get_cpi_base_year(
                requested_month
            )

            result = fetch_cpi_api(
                year=str(dt.year),
                month_code=str(dt.month),
                base_year=base_year,
            )

            if result:
                logger.info(
                    "CPI exact month successfully fetched: %s",
                    result,
                )
                return result

        except Exception:
            logger.exception(
                "CPI API failed for %s",
                requested_month,
            )

        logger.warning(
            "CPI extraction failed for %s",
            requested_month,
        )

        # Fall through to latest published month when exact report month is unavailable.

    # ============================================================
    # 2. DETERMINE GRAPH RANGE
    # ============================================================

    # Current report date.
    today = datetime.now()

    # The graph should end at the latest published month.
    #
    # The homepage is used only to discover that latest month.
    # We do NOT use its CPI value because the API is the source
    # used for the actual graph observations.

    latest_published_month = None

    try:
        homepage_cpi = fetch_cpi_from_homepage()

        if homepage_cpi:
            latest_published_month = (
                str(
                    homepage_cpi.get(
                        "month",
                        "",
                    )
                )
            )

    except Exception:
        logger.exception(
            "Could not determine latest CPI month from homepage"
        )

    # If homepage is unavailable, use the previous month as a
    # conservative fallback.
    if not latest_published_month:

        if today.month == 1:
            latest_date = today.replace(
                year=today.year - 1,
                month=12,
                day=1,
            )
        else:
            latest_date = today.replace(
                month=today.month - 1,
                day=1,
            )

        latest_published_month = (
            latest_date.strftime("%Y-%m")
        )

    logger.info(
        "Latest CPI month for graph: %s",
        latest_published_month,
    )

    # ============================================================
    # 3. GRAPH START MONTH
    # ============================================================

    # 22 monthly observations, matching the historical graph
    # window used elsewhere in the report.
    end_date = datetime.strptime(
        latest_published_month,
        "%Y-%m",
    )

    months_needed = []

    cursor = end_date

    for _ in range(22):
        months_needed.append(
            cursor.strftime("%Y-%m")
        )

        if cursor.month == 1:
            cursor = cursor.replace(
                year=cursor.year - 1,
                month=12,
            )
        else:
            cursor = cursor.replace(
                month=cursor.month - 1
            )

    months_needed.reverse()

    logger.info(
        "CPI graph months required: %s",
        months_needed,
    )

    # ============================================================
    # 4. FETCH EVERY MONTH DIRECTLY FROM CPI API
    # ============================================================

    results: List[Dict[str, Any]] = []

    for month in months_needed:

        try:

            dt = datetime.strptime(
                month,
                "%Y-%m",
            )

            # Automatically select the correct CPI series.
            #
            # Before Jan 2026:
            #     2012 base
            #
            # Jan 2026 onward:
            #     2024 base

            base_year = _get_cpi_base_year(
                month
            )

            logger.info(
                "Fetching CPI graph month: %s | base=%s",
                month,
                base_year,
            )

            month_result = fetch_cpi_api(
                year=str(dt.year),
                month_code=str(dt.month),
                base_year=base_year,
            )

            # fetch_cpi_api() should return the exact requested
            # month because year + month_code are supplied.

            matched = None

            for row in month_result:

                if (
                    str(row.get("month"))
                    == month
                    and row.get("CPI") is not None
                ):
                    matched = row
                    break

            if matched:

                results.append(
                    matched
                )

                logger.info(
                    "CPI graph value: %s",
                    matched,
                )

            else:

                logger.warning(
                    "No CPI value returned for %s",
                    month,
                )

        except Exception:

            logger.exception(
                "Failed to fetch CPI for %s",
                month,
            )

    # ============================================================
    # 5. SORT AND DEDUPLICATE
    # ============================================================

    unique: Dict[
        str,
        Dict[str, Any],
    ] = {}

    for row in results:

        month = str(
            row.get(
                "month",
                "",
            )
        )

        if month:
            unique[month] = row

    result = sorted(
        unique.values(),
        key=lambda row:
            str(
                row.get(
                    "month",
                    "",
                )
            ),
    )

    # ============================================================
    # 6. COVERAGE CHECK
    # ============================================================

    _log_monthly_coverage(
        result,
        start_month=(
            months_needed[0]
            if months_needed
            else None
        ),
        end_month=(
            months_needed[-1]
            if months_needed
            else None
        ),
        label="CPI",
    )

    logger.info(
        "Final CPI rows: %d",
        len(result),
    )

    logger.info(
        "Final CPI data: %s",
        result,
    )

    return result


# ============================================================
# IIP SCRAPING
# ============================================================

def fetch_iip_data(
    report_month: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """
    Fetch IIP General YoY growth using the correct MoSPI series.

    Series transition:
        - 2011-12 through March 2026
        - 2022-23 from April 2026

    For a requested month the scraper queries that exact month with the
    correct base year. The homepage is used only when it represents the
    exact requested month. No future or unrelated homepage observation is
    silently substituted for a historical request.
    """
    logger.info("========== IIP EXTRACTION START ==========")

    requested_month = _normalise_target_month(report_month)
    api_res: List[Dict[str, Any]] = []

    # ------------------------------------------------------------
    # TIER 1 — OFFICIAL e-SANKHYIKI IIP API
    # ------------------------------------------------------------
    try:
        if requested_month:
            base_year = _get_iip_base_year(requested_month)
            dt = datetime.strptime(requested_month, "%Y-%m")
            api_res = fetch_iip_api(
                year=str(dt.year),
                month_code=str(dt.month),
                base_year=base_year,
            )
        else:
            current_year = datetime.now().year
            previous_year = current_year - 1

            # Never mix the old and new IIP bases in one request.
            api_res.extend(
                fetch_iip_api(
                    year=str(previous_year),
                    month_code="1,2,3,4,5,6,7,8,9,10,11,12",
                    base_year=_get_iip_base_year(f"{previous_year}-12"),
                )
            )
            api_res.extend(
                fetch_iip_api(
                    year=str(current_year),
                    month_code="1,2,3,4,5,6,7,8,9,10,11,12",
                    base_year=IIP_CURRENT_BASE_YEAR,
                )
            )
            api_res.extend(
                fetch_iip_api(
                    year=str(current_year),
                    month_code="1,2,3",
                    base_year=IIP_LEGACY_BASE_YEAR,
                )
            )

        logger.info("IIP API returned %d normalized rows", len(api_res))
    except Exception:
        logger.exception("IIP API extraction failed")
        api_res = []

    # ------------------------------------------------------------
    # TIER 2 — OFFICIAL HOMEPAGE, EXACT MONTH ONLY
    # ------------------------------------------------------------
    try:
        homepage_iip = fetch_iip_from_homepage()
        if homepage_iip:
            homepage_month = str(homepage_iip.get("month", ""))

            if requested_month:
                if homepage_month == requested_month or homepage_month < requested_month:
                    existing = {str(row.get("month")) for row in api_res}
                    if homepage_month not in existing:
                        api_res.append(homepage_iip)
            else:
                existing = {str(row.get("month")) for row in api_res}
                if homepage_month not in existing:
                    api_res.append(homepage_iip)
    except Exception:
        logger.exception("IIP homepage extraction failed")

    unique: Dict[str, Dict[str, Any]] = {}
    for row in api_res:
        month = row.get("month")
        if month:
            unique[str(month)] = row

    result = sorted(
        unique.values(),
        key=lambda row: str(row.get("month", "")),
    )

    if requested_month:
        exact = [row for row in result if row.get("month") == requested_month]
        if exact:
            logger.info("IIP exact month %s obtained successfully", requested_month)
            return exact
        logger.warning("IIP exact month %s was not returned by API/homepage", requested_month)
        # Fall through to latest published month when exact report month is unavailable.

    if result:
        logger.info("IIP combined result: %d rows", len(result))
        logger.info("IIP latest result: %s", result[-1])
        _log_monthly_coverage(result, label="IIP")
        return result

    logger.warning("IIP API/homepage returned no usable data.")
    return []


# ============================================================
# GDP SCRAPING
# ============================================================

def fetch_gdp_data(
    report_month: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """
    Fetch quarterly GDP data from MoSPI eSankhyiki.

    Uses the new GDP series with base year 2022-23.
    Returns constant-price GDP values in ₹ crore.
    """

    logger.info("Fetching GDP data from MoSPI eSankhyiki...")

    try:
        df = esankhyiki.get_data(
            "NAS",
            {
                "indicator_code": 5,
                "base_year": "2022-23",
                "series": "Current",
                "frequency_code": 2,
                "limit": 100,
            },
            format="df",
        )

        if df is None or df.empty:
            logger.warning("No GDP data returned from eSankhyiki")
            return []

        results: List[Dict[str, Any]] = []

        for _, row in df.iterrows():
            year = str(row["year"]).strip()
            quarter = str(row["quarter"]).strip()

            constant_price = row.get("constant_price")

            if constant_price is None:
                continue

            try:
                gdp_rs_cr = int(float(constant_price))
            except (TypeError, ValueError):
                logger.warning(
                    f"Invalid GDP value for {year} {quarter}: "
                    f"{constant_price}"
                )
                continue

            # Convert API financial year + quarter
            # Example: 2025-26 + Q1 -> Q1-FY26
            fy_end = year.split("-")[-1]

            if len(fy_end) == 2:
                quarter_label = f"{quarter}-FY{fy_end}"
            else:
                quarter_label = f"{quarter}-FY{fy_end[-2:]}"

            results.append(
                {
                    "quarter": quarter_label,
                    "gdp_rs_cr": gdp_rs_cr,
                    "growth_pct": None,
                    "source": (
                        "MoSPI eSankhyiki NAS "
                        "(Base Year 2022-23)"
                    ),
                }
            )

        logger.info(
            f"Successfully extracted {len(results)} GDP records "
            f"from MoSPI eSankhyiki"
        )

        return results

    except Exception as e:
        logger.exception(
            f"Failed to fetch GDP data from eSankhyiki: {e}"
        )
        # Fall through to latest published month when exact report month is unavailable.

    return None
