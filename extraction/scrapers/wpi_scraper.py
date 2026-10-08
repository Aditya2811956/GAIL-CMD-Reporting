"""
wpi_scraper.py — Direct Homepage & In-Memory PDF scraper for Wholesale Price Index (WPI).

Source:
  Office of the Economic Adviser, Ministry of Commerce & Industry
  (https://eaindustry.nic.in/)

Approach:
  1. Queries the official portal homepage to discover the latest published press release PDF link.
  2. Downloads the bulletin PDF directly into memory.
  3. Extracts the provisional headline All-India Wholesale Price Index (WPI) inflation rate.
  4. Robust regex parsing supporting both '%' and 'per cent'.
  5. 100% fail-safe: returns None gracefully on any network or parsing error without crashing.
"""

import io
import logging
import re
from datetime import datetime
from typing import Any, Dict, List, Optional
from urllib.parse import urljoin

import requests
import urllib3
from bs4 import BeautifulSoup

urllib3.disable_warnings()

logger = logging.getLogger(__name__)

_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

WPI_BASE_URL = "https://eaindustry.nic.in/"

_MONTH_MAP = {
    "january": "01", "jan": "01",
    "february": "02", "feb": "02",
    "march": "03", "mar": "03",
    "april": "04", "apr": "04",
    "may": "05",
    "june": "06", "jun": "06",
    "july": "07", "jul": "07",
    "august": "08", "aug": "08",
    "september": "09", "sep": "09", "sept": "09",
    "october": "10", "oct": "10",
    "november": "11", "nov": "11",
    "december": "12", "dec": "12",
}


def _extract_text_from_pdf_bytes(pdf_bytes: bytes, max_pages: int = 5) -> str:
    """Safely extract text from in-memory PDF bytes using pdfplumber/pypdfium2/pypdf."""
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            pages = [p.extract_text() or "" for p in pdf.pages[:max_pages]]
            return "\n".join(pages)
    except Exception as e:
        logger.debug(f"pdfplumber extraction failed: {e}")

    try:
        import pypdfium2
        pdf = pypdfium2.PdfDocument(pdf_bytes)
        pages = [pdf[i].get_textpage().get_text_range() for i in range(min(max_pages, len(pdf)))]
        return "\n".join(pages)
    except Exception as e:
        logger.debug(f"pypdfium2 extraction failed: {e}")

    try:
        import pypdf
        reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
        pages_text = [page.extract_text() or "" for i, page in enumerate(reader.pages) if i < max_pages]
        return "\n".join(pages_text)
    except Exception as e:
        logger.debug(f"pypdf extraction failed: {e}")

    return ""


def fetch_wpi_data(
    report_month: Optional[str] = None,
) -> Optional[List[Dict[str, Any]]]:
    """
    Fetch historical monthly All-India WPI inflation from the
    official Office of Economic Adviser Excel dataset.

    Source:
        https://eaindustry.nic.in/

    Returns:
        [
            {
                "month": "YYYY-MM",
                "WPI": float,
                "source": "Office of Economic Adviser WPI Excel"
            },
            ...
        ]

    The Excel contains WPI index values. We calculate YoY inflation:

        WPI inflation = ((current index - index 12 months ago)
                         / index 12 months ago) * 100
    """

    logger.info(
        "Attempting WPI extraction from Office of Economic Adviser Excel..."
    )

    try:
        import io
        import openpyxl

        url = (
            "https://eaindustry.nic.in/"
            "indx_download_2223/wpi_monthly_index_202609.xlsx"
        )

        resp = requests.get(
            url,
            headers=_HEADERS,
            timeout=30,
            verify=False,
        )

        if resp.status_code != 200:
            logger.warning(
                "WPI Excel download failed: status=%s",
                resp.status_code,
            )
            return None

        content_type = (
            resp.headers.get("Content-Type") or ""
        ).lower()

        if "spreadsheet" not in content_type and not resp.content.startswith(b"PK"):
            logger.warning(
                "WPI download is not an Excel file: content-type=%s",
                content_type,
            )
            return None

        logger.info(
            "Downloaded WPI Excel successfully: %d bytes",
            len(resp.content),
        )

        workbook = openpyxl.load_workbook(
            io.BytesIO(resp.content),
            data_only=True,
            read_only=True,
        )

        sheet = workbook[workbook.sheetnames[0]]

        rows = list(
            sheet.iter_rows(
                values_only=True
            )
        )

        if not rows:
            logger.warning(
                "WPI Excel contains no rows"
            )
            return None

        header = rows[0]
        data_rows = rows[1:]

        # Columns:
        # 0 = Level
        # 1 = Commodity Name
        # 2 = Commodity Code
        # 3 = Weight
        # 4+ = monthly index values

        months = list(header[4:])

        all_row = next(
            (
                row
                for row in data_rows
                if row[0] == "ALL"
                and str(row[2]) == "1000000000"
            ),
            None,
        )

        if not all_row:
            logger.warning(
                "Could not find ALL Commodities WPI row"
            )
            return None

        index_values = list(all_row[4:])

        if len(months) != len(index_values):
            logger.warning(
                "WPI header/value length mismatch: months=%d values=%d",
                len(months),
                len(index_values),
            )
            return None

        # Convert Excel month labels such as:
        # Apr-24 -> 2024-04
        parsed = []

        for month_label, value in zip(
            months,
            index_values,
        ):

            if month_label is None or value is None:
                continue

            try:
                month_text = str(
                    month_label
                ).strip()

                month_dt = datetime.strptime(
                    month_text,
                    "%b-%y",
                )

                index_value = float(value)

                parsed.append(
                    (
                        month_dt,
                        index_value,
                    )
                )

            except Exception:
                continue

        parsed.sort(
            key=lambda x: x[0]
        )

        if len(parsed) < 13:
            logger.warning(
                "Not enough WPI history to calculate YoY inflation"
            )
            return None

        results = []

        # Need 12 months of history to calculate YoY.
        for i in range(
            12,
            len(parsed),
        ):

            current_month, current_index = (
                parsed[i]
            )

            previous_month, previous_index = (
                parsed[i - 12]
            )

            if previous_index == 0:
                continue

            wpi_value = (
                (current_index - previous_index)
                / previous_index
                * 100
            )

            results.append(
                {
                    "month": current_month.strftime(
                        "%Y-%m"
                    ),
                    "WPI": round(
                        wpi_value,
                        2,
                    ),
                    "source":
                        "Office of Economic Adviser WPI Excel",
                }
            )

        if not results:
            logger.warning(
                "WPI Excel produced no YoY results"
            )
            return None

        # If a specific month was requested,
        # return only that month if available.
        if report_month:

            matching = [
                row
                for row in results
                if row["month"] == report_month
            ]

            if matching:
                return matching

            logger.warning(
                "Requested WPI month %s not found",
                report_month,
            )

            return None

        logger.info(
            "Successfully extracted %d WPI monthly records",
            len(results),
        )

        logger.info(
            "Latest WPI record: %s",
            results[-1],
        )

        return results

    except Exception as e:

        logger.warning(
            "WPI Excel extraction failed safely: %s",
            e,
        )

        return None