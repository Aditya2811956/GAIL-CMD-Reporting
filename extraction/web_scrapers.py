"""
Web scrapers for metrics not available in ingested documents.

Provides automated scrapers for:
  - RBI Policy Repo Rate
  - SBI 1-Year MCLR Rate
  - MoSPI CPI (Consumer Price Index)
  - MoSPI IIP (Index of Industrial Production)
  - MoSPI GDP (Gross Domestic Product)
  - WPI (Wholesale Price Index)

Used for live report generation with automatic fallback to historical JSON / ChromaDB.
"""

import logging
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Any, Dict, Optional

import requests
from bs4 import BeautifulSoup

try:
    from .scrapers.mospi_scraper import fetch_cpi_data, fetch_iip_data, fetch_gdp_data
    from .scrapers.wpi_scraper import fetch_wpi_data
except ImportError:
    from extraction.scrapers.mospi_scraper import fetch_cpi_data, fetch_iip_data, fetch_gdp_data
    from extraction.scrapers.wpi_scraper import fetch_wpi_data

logger = logging.getLogger(__name__)

_HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}


def fetch_repo_rate() -> Optional[float]:
    """
    Scrape RBI Policy Repo Rate (most recent week).

    Source: https://www.rbi.org.in/Scripts/BS_NSDPDisplay.aspx?param=4

    Returns:
        Repo rate as float (e.g. 6.25), or None on failure.
    """
    url = "https://www.rbi.org.in/Scripts/BS_NSDPDisplay.aspx?param=4"
    try:
        r = requests.get(url, headers=_HEADERS, timeout=15)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")

        for table in soup.find_all("table"):
            for row in table.find_all("tr"):
                cells = [c.get_text(strip=True) for c in row.find_all(["td", "th"])]
                if not cells or cells[0] != "Policy Repo Rate":
                    continue
                values = [v for v in cells[1:] if v]
                if values:
                    # Rightmost column = most recent week
                    rate = float(values[-1])
                    logger.info(f"RBI Repo Rate scraped: {rate}%")
                    return rate

        logger.warning("Policy Repo Rate row not found in RBI table")
        return None

    except Exception as e:
        logger.error(f"Failed to scrape RBI repo rate: {e}")
        return None

def fetch_repo_rate_history(start_month: str = "2024-11", end_month: str = "2026-08") -> list[dict]:
    """
    Scrape RBI historical Policy Repo Rate data.

    Returns one Repo Rate observation per month using the latest
    available RBI weekly observation within that month.

    Example:
        [
            {"month": "2024-11", "Repo": 6.50},
            {"month": "2024-12", "Repo": 6.50},
            ...
        ]
    """

    from calendar import monthrange
    from datetime import datetime

    history = []

    try:
        start = datetime.strptime(start_month, "%Y-%m")
        end = datetime.strptime(end_month, "%Y-%m")

        current = start

        while current <= end:

            year = current.year
            month = current.month

            # Find all Saturdays in the month.
            last_day = monthrange(year, month)[1]

            for day in range(last_day, 0, -1):
                date_obj = datetime(year, month, day)

                # RBI WSS is published for Saturdays.
                if date_obj.weekday() != 5:
                    continue

                selected_date = f"{month}/{day}/{year}"

                detail_url = (
                    "https://www.rbi.org.in/Scripts/"
                    f"WSSViewDetail.aspx?TYPE=Basic&PARAM1={selected_date}"
                )

                try:
                    r = requests.get(
                        detail_url,
                        headers=_HEADERS,
                        timeout=15,
                    )
                    r.raise_for_status()

                    soup = BeautifulSoup(
                        r.text,
                        "html.parser",
                    )

                    ratios_link = None

                    for a in soup.find_all("a", href=True):
                        if a.get_text(" ", strip=True) == "Ratios and Rates":
                            ratios_link = a["href"]
                            break

                    if not ratios_link:
                        continue

                    ratios_url = (
                        "https://www.rbi.org.in/Scripts/"
                        + ratios_link.lstrip("./")
                    )

                    r2 = requests.get(
                        ratios_url,
                        headers=_HEADERS,
                        timeout=15,
                    )
                    r2.raise_for_status()

                    soup2 = BeautifulSoup(
                        r2.text,
                        "html.parser",
                    )

                    repo_rate = None

                    for row in soup2.find_all("tr"):
                        cells = [
                            c.get_text(" ", strip=True)
                            for c in row.find_all(["td", "th"])
                        ]

                        if not cells or cells[0] != "Policy Repo Rate":
                            continue

                        # The final cell is the most recent weekly Repo Rate
                        try:
                            repo_rate = float(cells[-1])
                        except (ValueError, TypeError):
                            repo_rate = None

                        break

                    if repo_rate is None:
                        continue

                    history.append(
                        {
                            "month": current.strftime("%Y-%m"),
                            "Repo": repo_rate,
                        }
                    )

                    logger.info(
                        "RBI Repo Rate: %s -> %.2f%%",
                        current.strftime("%Y-%m"),
                        repo_rate,
                    )

                    # We only need one observation per month.
                    break

                except Exception as e:
                    logger.warning(
                        "Failed RBI WSS lookup for %s: %s",
                        selected_date,
                        e,
                    )

            # Move to next month.
            if month == 12:
                current = datetime(year + 1, 1, 1)
            else:
                current = datetime(year, month + 1, 1)

        # Remove duplicates and sort.
        by_month = {}

        for item in history:
            by_month[item["month"]] = item

        history = sorted(
            by_month.values(),
            key=lambda x: x["month"],
        )

        logger.info(
            "RBI Repo historical data scraped: %d months",
            len(history),
        )

        return history

    except Exception as e:
        logger.error(
            "Failed to scrape RBI Repo history: %s",
            e,
        )
        return []


def fetch_mclr_1y() -> Optional[float]:
    """
    Scrape SBI MCLR 1-Year rate (current).

    Source: https://sbi.bank.in/web/interest-rates/interest-rates/mclr

    Returns:
        MCLR 1-Year rate as float (e.g. 8.70), or None on failure.
    """
    url = "https://sbi.bank.in/web/interest-rates/interest-rates/mclr"
    try:
        r = requests.get(url, headers=_HEADERS, timeout=15)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")

        for table in soup.find_all("table"):
            for row in table.find_all("tr"):
                cells = [c.get_text(strip=True) for c in row.find_all(["td", "th"])]
                if not cells:
                    continue
                # Match "One Year" row
                if any("one year" in c.lower() for c in cells):
                    values = [c for c in cells if _is_rate(c)]
                    if values:
                        # Use "Revised MCLR" if two columns, else first value
                        rate = float(values[-1])
                        logger.info(f"SBI MCLR 1Y scraped: {rate}%")
                        return rate

        logger.warning("One Year MCLR row not found in SBI table")
        return None

    except Exception as e:
        logger.error(f"Failed to scrape SBI MCLR: {e}")
        return None

def fetch_mclr_1y_history() -> list[dict]:
    """
    Scrape SBI 1-Year MCLR historical data.

    Returns:
        [
            {"month": "2026-04", "MCLR1Y": 8.70},
            ...
        ]
    """

    url = (
        "https://sbi.bank.in/web/interest-rates/"
        "interest-rates/mclr-historical-data"
    )

    history = []

    try:
        r = requests.get(
            url,
            headers=_HEADERS,
            timeout=15,
        )
        r.raise_for_status()

        soup = BeautifulSoup(
            r.text,
            "html.parser",
        )

        for table in soup.find_all("table"):

            rows = table.find_all("tr")

            for row in rows:

                cells = [
                    c.get_text(" ", strip=True)
                    for c in row.find_all(["td", "th"])
                ]

                # Data rows have:
                # Date + ON + 1M + 3M + 6M + 1Y + 2Y + 3Y
                if len(cells) < 8:
                    continue

                try:
                    effective_date = datetime.strptime(
                        cells[0],
                        "%d.%m.%Y",
                    )
                except ValueError:
                    continue

                try:
                    mclr_1y = float(
                        cells[5].replace("%", "").strip()
                    )
                except (ValueError, IndexError):
                    continue

                history.append(
                    {
                        "month": effective_date.strftime("%Y-%m"),
                        "MCLR1Y": mclr_1y,
                    }
                )

            if history:
                break

        # Keep one record per month.
        # If multiple observations exist for a month,
        # the later row encountered replaces the earlier one.
        by_month = {}

        for item in history:
            by_month[item["month"]] = item

        history = sorted(
            by_month.values(),
            key=lambda x: x["month"],
        )

        logger.info(
            "SBI MCLR historical data scraped: %d months",
            len(history),
        )

        return history

    except Exception as e:
        logger.error(
            "Failed to scrape SBI MCLR history: %s",
            e,
        )
        return []


def fetch_cpi(report_month: Optional[str] = None) -> Optional[Dict[str, Any]]:
    """Scrape latest CPI data."""
    return fetch_cpi_data(report_month)


def fetch_iip(report_month: Optional[str] = None) -> Optional[Dict[str, Any]]:
    """Scrape latest IIP data."""
    return fetch_iip_data(report_month)


def fetch_gdp(report_month: Optional[str] = None) -> Optional[Dict[str, Any]]:
    """Scrape latest GDP data."""
    return fetch_gdp_data(report_month)


def fetch_wpi(report_month: Optional[str] = None) -> Optional[Dict[str, Any]]:
    """Scrape latest WPI data."""
    return fetch_wpi_data(report_month)


def _is_rate(text: str) -> bool:
    """Check if a string looks like an interest rate (e.g. '8.70')."""
    try:
        val = float(text.replace("%", "").strip())
        return 1.0 <= val <= 25.0
    except ValueError:
        return False


def scrape_all_macro_metrics(report_month: Optional[str] = None) -> Dict[str, Any]:
    """
    Concurrently execute all live web scrapers.

    Returns:
        Dictionary containing extracted metrics (e.g., {'cpi': ..., 'iip': ..., 'wpi': ..., 'repo_rate': ...}).
        Missing/failed scrapers simply omit the key, ensuring zero application crashes.
    """
    logger.info(f"Running full macro web scraping suite for report_month={report_month}...")
    results: Dict[str, Any] = {}

    tasks = {
        "repo_rate": lambda: fetch_repo_rate(),
        "mclr_1y": lambda: fetch_mclr_1y(),
        "cpi": lambda: fetch_cpi_data(report_month),
        "iip": lambda: fetch_iip_data(report_month),
        "gdp": lambda: fetch_gdp_data(report_month),
        "wpi": lambda: fetch_wpi_data(report_month),
    }

    with ThreadPoolExecutor(max_workers=5) as executor:
        future_to_metric = {executor.submit(fn): metric for metric, fn in tasks.items()}
        for future in as_completed(future_to_metric):
            metric_name = future_to_metric[future]
            try:
                val = future.result()
                if val is not None:
                    results[metric_name] = val
                    logger.info(f"Scraper '{metric_name}' succeeded: {val}")
                else:
                    logger.info(f"Scraper '{metric_name}' returned None (fallback will be used)")
            except Exception as e:
                logger.warning(f"Scraper '{metric_name}' raised exception: {e} (skipping gracefully)")

    # Auto-persist into output/data/ if available
    #
    # IMPORTANT:
    # Scrapers may return either:
    #   {"month": "2026-07", "CPI": 4.45}
    #
    # or:
    #   [{"month": "2026-07", "CPI": 4.45}, ...]
    #
    # Therefore we normalise both forms before updating history.
    try:
        from pathlib import Path
        import json

        data_dir = (
            Path(__file__).resolve().parent.parent
            / "output"
            / "data"
        )

        if data_dir.exists():

            # ============================================================
            # INFLATION HISTORY
            # ============================================================

            if any(
                metric in results
                for metric in ("cpi", "wpi", "iip")
            ):

                inf_path = (
                    data_dir
                    / "inflation_history.json"
                )

                if inf_path.exists():

                    inf_history = json.loads(
                        inf_path.read_text(
                            encoding="utf-8"
                        )
                    )

                    if not isinstance(
                        inf_history,
                        list,
                    ):
                        inf_history = []

                    # ----------------------------------------------------
                    # Normalise scraper output to a list of rows.
                    # ----------------------------------------------------

                    def _normalise_metric_result(
                        value,
                        metric_name,
                    ):
                        rows = []

                        if isinstance(
                            value,
                            dict,
                        ):
                            month = value.get(
                                "month"
                            )

                            if month:
                                rows.append(
                                    {
                                        "month": month,
                                        metric_name: value.get(
                                            metric_name
                                        ),
                                    }
                                )

                        elif isinstance(
                            value,
                            list,
                        ):
                            for item in value:

                                if not isinstance(
                                    item,
                                    dict,
                                ):
                                    continue

                                month = item.get(
                                    "month"
                                )

                                if not month:
                                    continue

                                rows.append(
                                    {
                                        "month": month,
                                        metric_name: item.get(
                                            metric_name
                                        ),
                                    }
                                )

                        return rows

                    # ----------------------------------------------------
                    # Build a month -> row map from existing history.
                    # ----------------------------------------------------

                    history_by_month = {}

                    for item in inf_history:

                        if not isinstance(
                            item,
                            dict,
                        ):
                            continue

                        month = item.get(
                            "month"
                        )

                        if not month:
                            continue

                        history_by_month[
                            month
                        ] = dict(item)

                    # ----------------------------------------------------
                    # Merge each scraper result.
                    # ----------------------------------------------------

                    for result_key, metric_name in (
                        ("cpi", "CPI"),
                        ("wpi", "WPI"),
                        ("iip", "IIP"),
                    ):

                        if result_key not in results:
                            continue

                        rows = _normalise_metric_result(
                            results[result_key],
                            metric_name,
                        )

                        for row in rows:

                            month = row.get(
                                "month"
                            )

                            if not month:
                                continue

                            if month not in history_by_month:
                                history_by_month[
                                    month
                                ] = {
                                    "month": month
                                }

                            value = row.get(
                                metric_name
                            )

                            # Do NOT overwrite a valid historical
                            # value with None.
                            if value is not None:

                                history_by_month[
                                    month
                                ][
                                    metric_name
                                ] = value

                    # ----------------------------------------------------
                    # Keep history sorted chronologically.
                    # ----------------------------------------------------

                    inf_history = sorted(
                        history_by_month.values(),
                        key=lambda x: x["month"],
                    )

                    inf_path.write_text(
                        json.dumps(
                            inf_history,
                            indent=2,
                            allow_nan=False,
                        ),
                        encoding="utf-8",
                    )

                    logger.info(
                        "Persisted scraped inflation history: "
                        "%d months",
                        len(inf_history),
                    )

            # ============================================================
            # INTEREST RATE HISTORY
            # ============================================================

            if (
                "repo_rate" in results
                or "mclr_1y" in results
            ):

                ir_path = (
                    data_dir
                    / "interest_rates_history.json"
                )

                if ir_path.exists():

                    ir_history = json.loads(
                        ir_path.read_text(
                            encoding="utf-8"
                        )
                    )

                    if not isinstance(
                        ir_history,
                        list,
                    ):
                        ir_history = []

                    target_m = (
                        report_month
                        or "2026-08"
                    )

                    existing = next(
                        (
                            item
                            for item in ir_history
                            if item.get("month")
                            == target_m
                        ),
                        None,
                    )

                    if not existing:

                        existing = {
                            "month": target_m
                        }

                        ir_history.append(
                            existing
                        )

                    if (
                        "repo_rate" in results
                        and results["repo_rate"]
                    ):
                        existing[
                            "Repo"
                        ] = results[
                            "repo_rate"
                        ]

                    if (
                        "mclr_1y" in results
                        and results["mclr_1y"]
                    ):
                        existing[
                            "MCLR1Y"
                        ] = results[
                            "mclr_1y"
                        ]

                    ir_history = sorted(
                        ir_history,
                        key=lambda x: x["month"],
                    )

                    ir_path.write_text(
                        json.dumps(
                            ir_history,
                            indent=2,
                            allow_nan=False,
                        ),
                        encoding="utf-8",
                    )

                    logger.info(
                        "Persisted scraped interest rates "
                        "to %s for month %s",
                        ir_path.name,
                        target_m,
                    )

    except Exception as persist_err:

        logger.warning(
            "Failed to auto-persist scraped metrics "
            "(non-fatal): %s",
            persist_err,
        )

    return results
