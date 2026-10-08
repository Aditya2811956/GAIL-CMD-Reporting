"""
Excel document ingestion with smart header detection.

Handles government Excel files with irregular layouts:
- Empty rows at top
- Multi-row headers
- Metadata mixed with data
"""

import logging
from datetime import datetime
from io import BytesIO
from typing import Any, Dict, Generator, Optional, Tuple

import pandas as pd

try:
    from ..config import detect_metric_types
except ImportError:
    from config import detect_metric_types

logger = logging.getLogger(__name__)


def detect_header_row(df: pd.DataFrame, min_filled_ratio: float = 0.4) -> int:
    """
    Detect the header row in a DataFrame with irregular layout.

    Heuristics:
    1. Skip rows that are mostly empty
    2. Header row typically has text (not numbers) in most cells
    3. Header row usually comes before data rows with numbers
    """
    for idx in range(min(20, len(df))):
        row = df.iloc[idx]
        non_empty = row.dropna()

        if len(non_empty) < len(df.columns) * min_filled_ratio:
            continue

        string_count = sum(1 for v in non_empty
                          if isinstance(v, str) and not _is_numeric_string(str(v)))
        if string_count >= len(non_empty) * 0.5:
            return idx

    return 0


def _is_numeric_string(s: str) -> bool:
    """Check if string represents a number."""
    try:
        s = s.replace(',', '').strip()
        float(s)
        return True
    except (ValueError, AttributeError):
        return False


def extract_sheet_title(df_raw: pd.DataFrame, header_row: int) -> str:
    """Extract title/description from rows above header."""
    titles = []
    for idx in range(header_row):
        row = df_raw.iloc[idx]
        non_empty = [str(v).strip() for v in row.dropna()
                    if str(v).strip() and str(v).strip().lower() != 'nan']
        if non_empty:
            longest = max(non_empty, key=len)
            if len(longest) > 10:
                titles.append(longest)
    return " | ".join(titles[:2]) if titles else ""


def iter_excel_chunks(
    excel_bytes: bytes,
    document_name: str,
    report_date: Optional[str] = None,
) -> Generator[Tuple[str, Dict[str, Any]], None, None]:
    """
    Yield chunks from Excel file with smart header detection.

    Yields table-level chunks (full table as markdown for LLM context).
    """
    xl = pd.ExcelFile(BytesIO(excel_bytes))

    for sheet_idx, sheet_name in enumerate(xl.sheet_names, 1):
        logger.info(f"Processing sheet: {sheet_name}")

        df_raw = xl.parse(sheet_name, header=None)

        if df_raw.empty:
            continue

        header_row = detect_header_row(df_raw)
        logger.info(f"  Detected header at row {header_row}")

        sheet_title = extract_sheet_title(df_raw, header_row)

        df = xl.parse(sheet_name, header=header_row)
        df = df.dropna(how='all')

        df.columns = [str(c).strip() if pd.notna(c) else f"Col_{i}"
                      for i, c in enumerate(df.columns)]

        df = df.loc[:, ~df.columns.str.startswith('Unnamed')]

        if df.empty or len(df.columns) < 2:
            continue

        # Build full table as markdown
        table_md = df.to_markdown(index=False)
        table_text = f"Document: {document_name}\nSheet: {sheet_name}\n"
        if sheet_title:
            table_text += f"Title: {sheet_title}\n"
        table_text += f"\n{table_md}"

        metric_types = detect_metric_types(table_text)

        # Yield TABLE-level chunk
        table_meta = {
            "document_name": document_name,
            "chunk_type": "table",
            "sheet_name": sheet_name,
            "sheet_index": sheet_idx,
            "sheet_title": sheet_title,
            "report_date": report_date,
            "metric_types": ",".join(metric_types) if metric_types else "",
            "row_count": len(df),
            "columns": ",".join(df.columns.tolist()[:10]),
            "ingested_at": datetime.utcnow().isoformat(),
        }
        yield table_text, table_meta

        # Yield ROW-level chunks (needed for row-indexed sheets like IIP)
        for row_idx, row in df.iterrows():
            row_label = str(row.iloc[0]) if pd.notna(row.iloc[0]) else f"Row_{row_idx}"
            label_lower = row_label.lower()
            if any(p in label_lower for p in ["note:", "source:", "disclaimer", "updated on", "provisional"]):
                continue
            row_parts = []
            for col_idx, col in enumerate(df.columns):
                val = row.iloc[col_idx]
                if isinstance(val, pd.Series):
                    val = val.iloc[0] if len(val) > 0 else ""
                if pd.notna(val) and str(val).strip():
                    row_parts.append(f"{col}: {val}")
            if len(row_parts) < 2:
                continue
            row_text = f"[{sheet_name}] " + " | ".join(row_parts)
            yield row_text, {
                "document_name": document_name,
                "chunk_type": "row",
                "sheet_name": sheet_name,
                "sheet_title": sheet_title,
                "row_label": row_label[:100],
                "report_date": report_date,
                "metric_types": ",".join(metric_types) if metric_types else "",
                "ingested_at": datetime.utcnow().isoformat(),
            }
