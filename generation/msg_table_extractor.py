"""
Direct HTML-table extractor from .msg email files.

Bypasses ChromaDB/LLM for SATAT, CGD, and LNG tables by reading the email
body directly and pulling the HTML table at the known fixed index position.
"""
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Dict, List, Optional

logger = logging.getLogger(__name__)


def _open_msg(path: Path):
    import extract_msg
    return extract_msg.openMsg(str(path))

class MsgTableExtractor:
    """
    Extract tables directly from monthly .msg email files.

    documents_dir: folder containing the month's .msg files
                   (e.g. revamp/documents/dec_2025/)
    """

    def __init__(self, documents_dir: Path, report_month: Optional[str] = None):
        self.documents_dir = Path(documents_dir)
        self.report_month = report_month
        self._msg_cache: Dict[Path, object] = {}
        self._preload_msgs()

    def _preload_msgs(self) -> None:
        """Open all .msg files in documents_dir concurrently to parallelise RTF conversion."""
        if not self.documents_dir.exists():
            return
        paths = list(self.documents_dir.glob("*.msg"))
        if not paths:
            return
        logger.info("Pre-loading %d .msg file(s) concurrently from %s", len(paths), self.documents_dir)
        with ThreadPoolExecutor(max_workers=len(paths)) as executor:
            futures = {executor.submit(_open_msg, p): p for p in paths}
            for future in as_completed(futures):
                path = futures[future]
                try:
                    self._msg_cache[path] = future.result()
                    logger.info("Loaded %s", path.name)
                except Exception as e:
                    logger.warning("Failed to pre-load %s: %s", path.name, e)

    # ── public API ─────────────────────────────────────────────────────────

    def extract_satat_table(self) -> Optional[str]:
        """Extract Table 0 from the SATAT/CBG .msg file → markdown string."""
        # Prefer the dedicated CBG.msg file because CBG Production.msg
        # contains production data and may not contain HTML tables.
        msg_path = self.documents_dir / "CBG.msg"

        if not msg_path.exists():
            # Fallback for months where the dedicated CBG.msg is not present.
            msg_path = self._find_msg(keywords=["satat", "cbg"])

        if not msg_path:
            msg_path = self._find_msg(keywords=["cbg"])

        return self._extract_table_at(msg_path, table_index=0, min_fill=0.4)

    def extract_cgd_targets_table(self) -> Optional[str]:
        """
        Entity-row CGD targets table from the CGD+LNG .msg file → markdown string.

        The table has entities (IGL, MNGL, CUGL…) as rows and metrics (DPNG target,
        actual, CNG target, actual…) as columns.  Its position within the email varies
        by month, so we detect it by content rather than by fixed index.
        """
        msg_path = self._find_msg(keywords=["cgd"])
        return self._extract_cgd_table(msg_path, want_targets=True)

    def extract_cgd_cumulative_table(self) -> Optional[str]:
        """
        Metric-row CGD cumulative table from the CGD+LNG .msg file → markdown string.

        The table has metrics (DPNG Lakhs, Industrial, Commercial, CNG Stations) as rows
        and entities (IGL, MGL, MNGL…) as columns.  Its position within the email varies
        by month, so we detect it by content rather than by fixed index.
        """
        msg_path = self._find_msg(keywords=["cgd"])
        return self._extract_cgd_table(msg_path, want_targets=False)

    def extract_lng_stations_table(self) -> Optional[str]:
        """Extract Retail LNG station table and add station counts + total."""
        msg_path = self._find_msg(keywords=["cgd"])

        markdown = self._extract_table_at(
            msg_path,
            table_index=2,
            min_fill=0.3,
        )

        if not markdown:
            return None

        lines = [line.strip() for line in markdown.splitlines() if line.strip()]

        if len(lines) < 4:
            return markdown

        # Keep the markdown separator/header structure, but add the
        # "No. of Stations" column.
        header = lines[0]
        separator = lines[1]

        data_lines = lines[2:]

        output = [
            "| Stations | JV/Subsidiary | No. of Stations | Locations of stations + States |",
            "|:---------|:---------------|:----------------|:--------------------------------|",
        ]

        total = 0

        for line in data_lines:
            if not line.startswith("|"):
                continue

            columns = [col.strip() for col in line.strip("|").split("|")]

            if len(columns) < 3:
                continue

            station_type = columns[0]
            jv = columns[1]
            locations = columns[2]

            # Count actual station locations listed in the source table.
            count = len([
                location
                for location in locations.split(",")
                if location.strip()
            ])

            total += count

            output.append(
                f"| {station_type} | {jv} | {count:02d} | {locations} |"
            )


        output.append(
            f"| | | Total: {total} | |"
        )

        return "\n".join(output)

    def extract_cbg_production(self) -> Optional[dict]:
        """
        Extract CBG plant production data for the requested report month.

        Priority:
        1. Exact requested month/year (self.report_month)
        2. If unavailable, the most recent production month found
        """
        import re as _re
        from datetime import datetime as _datetime

        try:
            import extract_msg as _em
        except ImportError as e:
            logger.error(f"extract_msg not installed: {e}")
            return None

        if not self.documents_dir.exists():
            return None

        candidate_paths = sorted(
            self.documents_dir.glob("*.msg"),
            key=lambda p: p.name.lower()
        )

        cbg_paths = [
            p for p in candidate_paths
            if "cbg" in p.name.lower() and "satat" not in p.name.lower()
        ]

        # Fallback to any CBG file if no non-SATAT CBG files exist.
        if not cbg_paths:
            cbg_paths = [
                p for p in candidate_paths
                if "cbg" in p.name.lower()
            ]

        prod_pat = _re.compile(
            r"Production of CBG in the Month of (\w+ \d{4}) was\s+([\d,]+)\s*Kg",
            _re.IGNORECASE
        )
        cum_pat = _re.compile(
            r"Cumulative Production till (\w+ \d{4}) was\s+([\d,]+)\s*Kg",
            _re.IGNORECASE
        )

        # Collect every production month available in the source files.
        available = []

        for msg_path in cbg_paths:
            try:
                msg = self._msg_cache.get(msg_path) or _em.openMsg(str(msg_path))
                body = msg.body or ""

                prod_matches = list(prod_pat.finditer(body))
                cum_matches = list(cum_pat.finditer(body))

                for m_prod in prod_matches:
                    month_str = m_prod.group(1)

                    # Find cumulative value for the same month.
                    cumulative_value = None
                    for m_cum in cum_matches:
                        if m_cum.group(1).lower() == month_str.lower():
                            cumulative_value = m_cum.group(2)
                            break

                    try:
                        month_date = _datetime.strptime(month_str, "%B %Y")
                    except ValueError:
                        continue

                    available.append({
                        "month_str": month_str,
                        "month_date": month_date,
                        "month_production_kg": m_prod.group(2),
                        "cumulative_production_kg": cumulative_value,
                        "source": msg_path.name,
                    })

            except Exception as e:
                logger.warning(f"Failed to read {msg_path.name}: {e}")

        if not available:
            logger.warning(
                f"No CBG production data found in {self.documents_dir}"
            )
            return None

        # 1. Prefer the exact requested report month.
        requested_date = None

        if self.report_month:
            try:
                requested_date = _datetime.strptime(
                    self.report_month,
                    "%Y-%m"
                )
            except ValueError:
                logger.warning(
                    f"Invalid report_month format: {self.report_month}"
                )

        if requested_date:
            exact_matches = [
                item for item in available
                if item["month_date"].year == requested_date.year
                and item["month_date"].month == requested_date.month
            ]

            if exact_matches:
                selected = exact_matches[-1]
                logger.info(
                    f"CBG production data found for requested month "
                    f"{self.report_month}: {selected['month_str']} "
                    f"from {selected['source']}"
                )
                return {
                    "month_str": selected["month_str"],
                    "month_production_kg": selected["month_production_kg"],
                    "cumulative_production_kg": selected["cumulative_production_kg"],
                }

            logger.warning(
                f"No CBG production data found for requested month "
                f"{self.report_month}; falling back to most recent available month"
            )

        # 2. Fall back to the most recent available production month.
        selected = max(
            available,
            key=lambda item: item["month_date"]
        )

        logger.info(
            f"Using most recent available CBG production data: "
            f"{selected['month_str']} from {selected['source']}"
        )

        return {
            "month_str": selected["month_str"],
            "month_production_kg": selected["month_production_kg"],
            "cumulative_production_kg": selected["cumulative_production_kg"],
        }

    # ── private helpers ─────────────────────────────────────────────────────

    # Known entity names that appear as row labels in the targets table
    # or as column headers in the cumulative table.
    _CGD_ENTITY_NAMES = {"IGL", "MGL", "MNGL", "CUGL", "GGL", "AGL", "BGL", "TNGCL", "BGCL"}

    def _is_targets_table(self, tbl) -> bool:
        """
        Return True if this is the entity-row targets table (rows = entities, cols = metrics).

        Detection: the first data row's first non-empty cell is a known entity name.
        Cumulative table has entities as column 2+ headers in row 0.
        """
        rows = tbl.find_all("tr")
        if len(rows) < 2:
            return False
        # Check second row (first data row after header) — col 0 should be entity name
        first_data_cells = [td.get_text(strip=True) for td in rows[1].find_all(["td", "th"])]
        if first_data_cells:
            label = first_data_cells[0].strip().upper()
            for entity in self._CGD_ENTITY_NAMES:
                if entity in label:
                    return True
        return False

    def _extract_cgd_table(self, msg_path: Optional[Path], want_targets: bool) -> Optional[str]:
        """
        Scan CGD email tables and return the one matching the requested type.

        want_targets=True  → entity-row table (DPNG/CNG target/actual per entity)
        want_targets=False → metric-row table (DPNG Lakhs, Industrial… per entity column)
        """
        if msg_path is None:
            return None
        tables = self._get_html_tables(msg_path)
        for tbl_idx, tbl in enumerate(tables[:3]):  # only inspect first 3 tables
            if not self._is_filled(tbl, min_fill=0.4):
                continue
            is_targets = self._is_targets_table(tbl)
            if is_targets != want_targets:
                continue
            # For cumulative (want_targets=False): validate at least one CGD entity appears
            # in the header row to avoid false-positives from unrelated tables (e.g. LNG stations)
            if not want_targets and not self._has_entity_header(tbl):
                continue
            markdown = self._table_to_markdown(tbl)
            logger.info(
                f"Extracted {'targets' if want_targets else 'cumulative'} CGD table "
                f"from {msg_path.name} (email table[{tbl_idx}])"
            )
            return markdown if markdown.strip() else None
        kind = "targets" if want_targets else "cumulative"
        logger.warning(f"{msg_path.name}: no {kind} CGD table found in first 3 tables")
        return None

    def _has_entity_header(self, tbl) -> bool:
        """Return True if any CGD entity name appears in the first row of the table."""
        rows = tbl.find_all("tr")
        if not rows:
            return False
        header_cells = [td.get_text(strip=True).upper() for td in rows[0].find_all(["td", "th"])]
        for cell in header_cells:
            for entity in self._CGD_ENTITY_NAMES:
                if entity in cell:
                    return True
        return False

    def _find_msg(self, keywords: List[str]) -> Optional[Path]:
        """
        Find the first .msg file whose name (lowercased) contains ALL keywords.
        Prefers files where ALL keywords match over partial matches.
        """
        if not self.documents_dir.exists():
            logger.warning(f"documents_dir does not exist: {self.documents_dir}")
            return None
        for path in self.documents_dir.glob("*.msg"):
            name_lower = path.name.lower()
            if all(kw in name_lower for kw in keywords):
                logger.debug(f"Matched {keywords} → {path.name}")
                return path
        logger.warning(f"No .msg file found matching {keywords} in {self.documents_dir}")
        return None

    def _get_html_tables(self, msg_path: Path):
        """Return list of <table> BeautifulSoup tags from the cached .msg HTML body."""
        try:
            from bs4 import BeautifulSoup
        except ImportError as e:
            logger.error(f"Missing dependency for .msg parsing: {e}")
            return []

        try:
            msg = self._msg_cache.get(msg_path)
            if msg is None:
                logger.warning(f"{msg_path.name} not in cache — opening now")
                msg = _open_msg(msg_path)
            html = getattr(msg, "htmlBody", None) or b""
            if isinstance(html, str):
                html = html.encode()
            if not html:
                logger.warning(f"No HTML body in {msg_path.name}")
                return []
            soup = BeautifulSoup(html, "html.parser")
            return soup.find_all("table")
        except Exception as e:
            logger.error(f"Failed to parse {msg_path}: {e}")
            return []

    def _table_to_markdown(self, tbl) -> str:
        """Convert a BS4 <table> tag to a pipe-delimited markdown string.

        Rows where all data columns (col 1 onwards) are empty are dropped.
        These are section-header separators (e.g. a "GAIL" row with no value
        in the second column) and would shift data rows if kept.
        """
        import re
        from tabulate import tabulate
        rows = []
        for tr in tbl.find_all("tr"):
            cells = []
            for td in tr.find_all(["td", "th"]):
                text = td.get_text(" ", strip=True)
                # Collapse spaces between digits caused by inline HTML sub-elements
                # e.g. "1 2" → "12", "4,5 6 3,7 2 0" → "4,563,720"
                text = re.sub(r"(\d)\s+(\d)", r"\1\2", text)
                cells.append(text)
            rows.append(cells)
        if not rows:
            return ""
        # Skip body rows whose data columns (col 1+) are all empty
        body = [r for r in rows[1:] if any(c.strip() for c in r[1:])]
        return tabulate(body, headers=rows[0], tablefmt="pipe")

    def _is_filled(self, tbl, min_fill: float) -> bool:
        """Return True if at least min_fill fraction of cells are non-empty."""
        cells = tbl.find_all(["td", "th"])
        if not cells:
            return False
        filled = sum(1 for c in cells if c.get_text(strip=True))
        return filled / len(cells) >= min_fill

    def _extract_table_at(
        self,
        msg_path: Optional[Path],
        table_index: int,
        min_fill: float,
    ) -> Optional[str]:
        """
        Extract the HTML table at table_index from msg_path and return as markdown.
        Returns None if the file is missing, table_index is out of range, or the
        table is below min_fill threshold.
        """
        if msg_path is None:
            return None
        tables = self._get_html_tables(msg_path)
        if table_index >= len(tables):
            logger.warning(
                f"{msg_path.name}: only {len(tables)} tables found, "
                f"requested index {table_index}"
            )
            return None
        tbl = tables[table_index]
        if not self._is_filled(tbl, min_fill):
            logger.warning(
                f"{msg_path.name} table[{table_index}] is below fill threshold "
                f"({min_fill:.0%}) — skipping"
            )
            return None
        markdown = self._table_to_markdown(tbl)
        logger.info(
            f"Extracted table[{table_index}] from {msg_path.name} "
            f"({len(tbl.find_all('tr'))} rows)"
        )
        return markdown if markdown.strip() else None
