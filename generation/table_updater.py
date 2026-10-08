"""
Table update functions for CMD report generation.

Handles direct table manipulation for:
- GDP table (rolling 12 quarters from historical JSON)
- GDP Forecast table (from ChromaDB extraction)
- SATAT table (from email data in ChromaDB)
- CGD tables (from email data in ChromaDB)
- Retail LNG table (from email data in ChromaDB)
"""

import json
import logging
from pathlib import Path
from typing import Dict, List, Optional, Any

from docx import Document
from openai import OpenAI, AzureOpenAI
import chromadb

try:
    from .tables import update_table_from_markdown
    from ..config import ChromaConfig, AzureConfig, OpenAIConfig, LLMConfig
except ImportError:
    from generation.tables import update_table_from_markdown
    from config import ChromaConfig, AzureConfig, OpenAIConfig, LLMConfig

logger = logging.getLogger(__name__)


# =============================================================================
# Table Indices (template1.docx)
# =============================================================================
TABLE_INDICES = {
    "gdp": 0,
    "gdp_forecast": 1,  # Not in ChromaDB yet - manual update needed
    "satat": 2,
    "retail_lng": 3,    # Template 2 only — Retail LNG stations
    "cgd_targets": 3,   # Template 1: index 3 | Template 2: index 4 (use service.py TABLE_INDICES for dispatch)
    "cgd_cumulative": 4,  # Template 1: index 4 | Template 2: index 5
}


# =============================================================================
# GDP Table Update (from historical JSON)
# =============================================================================
GDP_ROW_MAPPING = {
    "GDP (Rs. Cr)":       "gdp_rs_cr",
    "GDP (USD Billion)":  "gdp_usd_bn",
    "% Growth (y-o-y)":   "growth_pct",
}


def format_indian(num: int) -> str:
    """Format number using Indian number system (e.g. 3891251 → 38,91,251)."""
    s = str(abs(int(num)))
    if len(s) <= 3:
        return ("-" if num < 0 else "") + s
    last3 = s[-3:]
    rest = s[:-3]
    groups: List[str] = []
    while len(rest) > 2:
        groups.insert(0, rest[-2:])
        rest = rest[:-2]
    if rest:
        groups.insert(0, rest)
    return ("-" if num < 0 else "") + ",".join(groups) + "," + last3


def parse_quarter(quarter_str: str) -> tuple:
    """
    Parse quarter string to sortable tuple.

    Args:
        quarter_str: Quarter in format "Q1-FY25"

    Returns:
        Tuple of (fy_year, quarter_num) for sorting

    Example:
        "Q1-FY25" → (25, 1)
        "Q4-FY24" → (24, 4)
    """
    q, fy = quarter_str.split('-')
    q_num = int(q[1])  # "Q1" → 1
    fy_year = int(fy[2:])  # "FY25" → 25
    return (fy_year, q_num)


def sort_gdp_history(history: List[Dict]) -> List[Dict]:
    """Sort GDP history by quarter chronologically."""
    return sorted(history, key=lambda x: parse_quarter(x['quarter']))


def quarter_end_month(quarter_str: str) -> str:
    """
    Return the last calendar month of a quarter as "YYYY-MM".

    Indian fiscal year quarters:
      Q1 = April–June     → end month = 06
      Q2 = July–September → end month = 09
      Q3 = October–December → end month = 12
      Q4 = January–March  → end month = 03 (next calendar year)

    Examples:
      "Q1-FY26" → "2025-06"
      "Q2-FY26" → "2025-09"
      "Q4-FY26" → "2026-03"
    """
    q_num, fy_year = parse_quarter(quarter_str)  # (fy_year, q_num)
    # parse_quarter returns (fy_year, q_num) but the tuple is (fy_year, q_num)
    fy_year, q_num = parse_quarter(quarter_str)
    # fy_year is 2-digit (e.g. 26 for FY2025-26)
    full_fy = fy_year + 2000  # 26 → 2026 (end year of fiscal year)
    start_year = full_fy - 1   # FY26 starts April 2025

    if q_num == 1:   # April–June
        return f"{start_year}-06"
    elif q_num == 2: # July–September
        return f"{start_year}-09"
    elif q_num == 3: # October–December
        return f"{start_year}-12"
    else:            # Q4: January–March (next calendar year = full_fy)
        return f"{full_fy}-03"


def filter_gdp_by_cutoff(history: List[Dict], cutoff_month: str) -> List[Dict]:
    """
    Exclude GDP quarters whose end month >= cutoff_month.

    Used in benchmark runs to prevent future quarters (not yet published at the
    time of the report) from appearing in the generated GDP table.

    Rule: include quarter if quarter_end_month(q) < cutoff_month.

    Example: cutoff_month="2025-09" → excludes Q2-FY26 (ends 2025-09) and later;
             includes Q1-FY26 (ends 2025-06).
    """
    return [
        q for q in history
        if quarter_end_month(q["quarter"]) < cutoff_month
    ]


def _fill_computed_fields(sorted_quarters: List[Dict], exchange_rate: Optional[float] = None) -> List[Dict]:
    """
    Fill growth_pct (YoY) and gdp_usd_bn (from exchange rate) where missing.

    Args:
        sorted_quarters: List of quarter dicts, sorted chronologically
        exchange_rate: INR per 1 USD (e.g., 84.5). If None, gdp_usd_bn not computed.

    Returns:
        Same list with missing fields filled in-place.
    """
    # Build lookup: (fy_year, q_num) → rs_cr for YoY comparison
    lookup = {parse_quarter(q["quarter"]): q.get("gdp_rs_cr") for q in sorted_quarters}

    for q in sorted_quarters:
        fy, qn = parse_quarter(q["quarter"])

        # growth_pct: YoY = (current quarter / same quarter previous fiscal year - 1) * 100
        if q.get("growth_pct") is None:
            prev_key = (fy - 1, qn)  # same quarter, one fiscal year earlier
            prev_rs_cr = lookup.get(prev_key)
            cur_rs_cr = q.get("gdp_rs_cr")
            if prev_rs_cr and cur_rs_cr:
                q["growth_pct"] = round((cur_rs_cr / prev_rs_cr - 1) * 100, 1)

        # gdp_usd_bn: Rs Cr / (exchange_rate × 100) = USD Billion
        if q.get("gdp_usd_bn") is None and exchange_rate and q.get("gdp_rs_cr"):
            q["gdp_usd_bn"] = round(q["gdp_rs_cr"] / (exchange_rate * 100), 1)

    return sorted_quarters


def build_gdp_table_markdown(quarters_data: List[Dict]) -> str:
    """
    Build markdown table from GDP quarters data.

    Args:
        quarters_data: List of dicts with keys: quarter, gdp_rs_cr, growth_pct, gdp_usd_bn

    Returns:
        Markdown string representation of GDP table
    """
    # Build header row
    header = [""] + [q['quarter'] for q in quarters_data]

    # Build data rows
    rows = [header]
    for row_label, data_key in GDP_ROW_MAPPING.items():
        row = [row_label]
        for quarter_data in quarters_data:
            value = quarter_data.get(data_key)

            if value is not None:
                if data_key == "growth_pct":
                    row.append(f"{value}%")
                elif data_key == "gdp_rs_cr":
                    row.append(format_indian(value))
                else:
                    row.append(str(value))
            else:
                row.append("")
        rows.append(row)

    # Convert to markdown
    markdown_lines = []
    for row in rows:
        markdown_lines.append("| " + " | ".join(row) + " |")

    # Add separator after header
    separator = "| " + " | ".join(["---"] * len(header)) + " |"
    markdown_lines.insert(1, separator)

    return "\n".join(markdown_lines)


def update_gdp_table(doc: Document, gdp_history_path: Path, num_quarters: int = 12) -> bool:
    """
    Update GDP table with rolling quarters from historical JSON.

    Args:
        doc: python-docx Document object
        gdp_history_path: Path to gdp_history.json
        num_quarters: Number of most recent quarters to include (default: 12)

    Returns:
        True if successful, False otherwise
    """
    try:
        # Load GDP history
        with open(gdp_history_path) as f:
            gdp_history = json.load(f)

        logger.info(f"Loaded {len(gdp_history)} GDP quarters from {gdp_history_path}")

        # Sort and take last N quarters
        sorted_gdp = sort_gdp_history(gdp_history)
        last_quarters = sorted_gdp[-num_quarters:]

        logger.info(f"Selected last {len(last_quarters)} quarters: {[q['quarter'] for q in last_quarters]}")

        # Build markdown table
        markdown = build_gdp_table_markdown(last_quarters)

        # Update table using existing utility
        update_table_from_markdown(
            doc=doc,
            table_index=TABLE_INDICES["gdp"],
            markdown_table=markdown,
            preserve_styles=True
        )

        logger.info(f"✅ GDP table updated successfully with {len(last_quarters)} quarters")
        return True

    except Exception as e:
        logger.error(f"Failed to update GDP table: {e}")
        return False


# =============================================================================
# ChromaDB Table Extraction (for email tables)
# =============================================================================

def get_query_embedding(text: str) -> List[float]:
    """
    Generate embedding using configured provider (matches ingestion).

    Args:
        text: Query text to embed

    Returns:
        Embedding vector as list of floats
    """
    if LLMConfig.PROVIDER == "openai":
        # Use OpenAI
        client = OpenAI(api_key=OpenAIConfig.API_KEY)
        model = OpenAIConfig.EMBED_MODEL
    else:
        # Use Azure OpenAI
        client = AzureOpenAI(
            azure_endpoint=AzureConfig.OPENAI_ENDPOINT,
            api_key=AzureConfig.OPENAI_API_KEY,
            api_version=AzureConfig.OPENAI_API_VERSION,
        )
        model = AzureConfig.EMBED_DEPLOYMENT

    response = client.embeddings.create(
        model=model,
        input=[text]
    )
    return response.data[0].embedding


def extract_table_from_chromadb(
    collection,
    metric_type: str,
    query_text: str,
    n_results: int = 10
) -> Optional[Dict[str, Any]]:
    """
    Extract table from ChromaDB using metric_types filter.

    Args:
        collection: ChromaDB collection object
        metric_type: Metric type to filter by (e.g., "satat", "cgd")
        query_text: Query text for semantic search
        n_results: Number of results to retrieve for filtering

    Returns:
        Dict with 'table_markdown' and 'metadata', or None if not found
    """
    if collection is None:
        logger.warning(f"ChromaDB collection not provided, skipping {metric_type}")
        return None

    try:
        # Get embedding using configured provider
        query_embedding = get_query_embedding(query_text)

        # Query with embedding and where filter
        results = collection.query(
            query_embeddings=[query_embedding],
            where={"chunk_type": "table"},
            n_results=n_results
        )

        if not results['documents'] or not results['documents'][0]:
            logger.warning(f"No tables found for {metric_type}")
            return None

        # Filter by metric_types and source (post-query)
        for doc, meta in zip(results['documents'][0], results['metadatas'][0]):
            metric_types = meta.get('metric_types', '')
            source = meta.get('source', '')

            if metric_type in metric_types and source == "email_body":
                logger.info(f"✅ Found {metric_type} table from email: {meta.get('email_subject', 'N/A')[:60]}")
                return {
                    'table_markdown': doc,
                    'metadata': meta
                }

        logger.warning(f"No tables tagged with metric_type='{metric_type}' from email_body")
        return None

    except Exception as e:
        logger.error(f"Error querying ChromaDB for {metric_type}: {e}")
        return None


def update_table_from_chromadb(
    doc: Document,
    table_name: str,
    collection,
    metric_type: str,
    query_text: str
) -> bool:
    """
    Update a table by extracting data from ChromaDB.

    Args:
        doc: python-docx Document object
        table_name: Name of table (key in TABLE_INDICES)
        collection: ChromaDB collection object
        metric_type: Metric type to filter by
        query_text: Query text for semantic search

    Returns:
        True if successful, False otherwise
    """
    try:
        # Extract table from ChromaDB
        result = extract_table_from_chromadb(
            collection=collection,
            metric_type=metric_type,
            query_text=query_text
        )

        if not result:
            logger.warning(f"Could not extract {table_name} table from ChromaDB")
            return False

        # Update table
        table_idx = TABLE_INDICES[table_name]
        update_table_from_markdown(
            doc=doc,
            table_index=table_idx,
            markdown_table=result['table_markdown'],
            preserve_styles=True
        )

        logger.info(f"✅ {table_name} table updated successfully")
        return True

    except Exception as e:
        logger.error(f"Failed to update {table_name} table: {e}")
        return False


# =============================================================================
# Main Table Updater Class
# =============================================================================

class TableUpdater:
    """
    Main class for updating all tables in CMD report.

    Handles:
    - GDP table from historical JSON
    - Email tables (SATAT, CGD, Retail LNG) from ChromaDB
    """

    def __init__(
        self,
        data_dir: Path,
        collection_name: Optional[str] = None,
        chroma_host: Optional[str] = None,
        chroma_port: Optional[int] = None
    ):
        """
        Initialize table updater.

        Args:
            data_dir: Directory containing historical JSON files
            collection_name: ChromaDB collection name
            chroma_host: ChromaDB host
            chroma_port: ChromaDB port
        """
        self.data_dir = Path(data_dir)

        # Initialize ChromaDB connection
        try:
            self.chroma_client = chromadb.HttpClient(
                host=chroma_host or ChromaConfig.HOST,
                port=chroma_port or ChromaConfig.PORT
            )
            self.collection = self.chroma_client.get_collection(
                collection_name or ChromaConfig.COLLECTION
            )
            logger.info(f"Connected to ChromaDB collection: {collection_name or ChromaConfig.COLLECTION}")
        except Exception as e:
            logger.warning(f"ChromaDB connection failed: {e}. Email table extraction will be skipped.")
            self.collection = None

    def update_gdp_table(self, doc: Document) -> bool:
        """Update GDP table from gdp_history.json."""
        gdp_history_path = self.data_dir / "gdp_history.json"
        return update_gdp_table(doc, gdp_history_path)

    def update_gdp_forecast_table(self, doc: Document) -> bool:
        """
        Update GDP Forecast table from ChromaDB.

        Note: Currently not in ChromaDB - data needs to be ingested from PDFs.
        Returns False for now.
        """
        logger.warning("GDP Forecast data not yet in ChromaDB - skipping")
        return False

    def update_satat_table(self, doc: Document) -> bool:
        """Update SATAT table from ChromaDB."""
        return update_table_from_chromadb(
            doc=doc,
            table_name="satat",
            collection=self.collection,
            metric_type="satat",
            query_text="SATAT CBG plants compressed biogas scheme data"
        )

    def update_retail_lng_table(self, doc: Document) -> bool:
        """Update Retail LNG table (template2 only, Table 3) from ChromaDB."""
        return update_table_from_chromadb(
            doc=doc,
            table_name="retail_lng",
            collection=self.collection,
            metric_type="retail_lng",
            query_text="GAIL Retail LNG stations MGL VGL IGL LCNG"
        )

    def update_cgd_targets_table(self, doc: Document) -> bool:
        """Update CGD Targets table (Table 3) from ChromaDB."""
        return update_table_from_chromadb(
            doc=doc,
            table_name="cgd_targets",
            collection=self.collection,
            metric_type="cgd",
            query_text="CGD city gas distribution targets achievements DPNG CNG"
        )

    def update_cgd_cumulative_table(self, doc: Document) -> bool:
        """Update CGD Cumulative table (Table 4) from ChromaDB."""
        return update_table_from_chromadb(
            doc=doc,
            table_name="cgd_cumulative",
            collection=self.collection,
            metric_type="cgd",
            query_text="CGD city gas distribution IGL MGL MNGL CUGL cumulative"
        )

    def update_all_tables(self, doc: Document) -> Dict[str, bool]:
        """
        Update all tables in the document.

        Args:
            doc: python-docx Document object

        Returns:
            Dict mapping table name to success status
        """
        results = {}

        logger.info("Starting table updates...")

        # GDP table (from JSON)
        results['gdp'] = self.update_gdp_table(doc)

        # Tables from ChromaDB
        if self.collection:
            results['gdp_forecast'] = self.update_gdp_forecast_table(doc)  # Currently not available
            results['satat'] = self.update_satat_table(doc)
            results['retail_lng'] = self.update_retail_lng_table(doc)
            results['cgd_targets'] = self.update_cgd_targets_table(doc)
            results['cgd_cumulative'] = self.update_cgd_cumulative_table(doc)
        else:
            logger.warning("ChromaDB not connected, skipping email table updates")
            results['gdp_forecast'] = False
            results['satat'] = False
            results['retail_lng'] = False
            results['cgd_targets'] = False
            results['cgd_cumulative'] = False

        # Log summary
        successful = sum(results.values())
        total = len(results)
        logger.info(f"Table update complete: {successful}/{total} successful")

        return results
