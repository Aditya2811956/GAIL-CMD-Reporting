"""
PDF document ingestion using Azure Document Intelligence.

Extracts tables and paragraphs from PDFs with layout analysis.
"""

import logging
from datetime import datetime
from typing import Any, Dict, Generator, Optional, Tuple

from azure.ai.documentintelligence import DocumentIntelligenceClient
from azure.ai.documentintelligence.models import AnalyzeDocumentRequest
from azure.core.credentials import AzureKeyCredential
from azure.core.pipeline.transport import RequestsTransport

try:
    from ..config import AzureConfig, detect_metric_types
except ImportError:
    from config import AzureConfig, detect_metric_types

logger = logging.getLogger(__name__)


class InsecureHttpTransport(RequestsTransport):
    """Transport that disables SSL verification. Only used when DISABLE_SSL_VERIFY=1."""
    def __init__(self):
        super().__init__(connection_verify=False)


_di_client: Optional[DocumentIntelligenceClient] = None


def get_di_client() -> DocumentIntelligenceClient:
    """Get or create Document Intelligence client."""
    import os
    global _di_client
    if _di_client is None:
        transport = InsecureHttpTransport() if os.getenv("DISABLE_SSL_VERIFY") else None
        _di_client = DocumentIntelligenceClient(
            AzureConfig.DI_ENDPOINT,
            AzureKeyCredential(AzureConfig.DI_KEY),
            transport=transport,
        )
    return _di_client


def iter_pdf_chunks(
    pdf_bytes: bytes,
    document_name: str,
    report_date: Optional[str] = None,
) -> Generator[Tuple[str, Dict[str, Any]], None, None]:
    """
    Yield chunks from PDF using Azure Document Intelligence.

    Yields FULL TABLES (not just rows) for better LLM context.
    """
    di_client = get_di_client()

    logger.info(f"Analyzing PDF with Document Intelligence: {document_name}")
    poller = di_client.begin_analyze_document(
        model_id="prebuilt-layout",
        body=AnalyzeDocumentRequest(bytes_source=pdf_bytes),
    )
    result = poller.result()

    pages = getattr(result, "pages", []) or []
    tables = getattr(result, "tables", []) or []

    logger.info(f"  Found {len(pages)} pages, {len(tables)} tables")

    # Group tables by page
    tables_by_page: Dict[int, list] = {}
    for tbl in tables:
        regions = getattr(tbl, "bounding_regions", []) or []
        if regions:
            page_num = regions[0].page_number
            tables_by_page.setdefault(page_num, []).append(tbl)

    # Process each table
    for page_num in sorted(tables_by_page.keys()):
        for tbl_idx, tbl in enumerate(tables_by_page[page_num], 1):
            # Extract headers
            headers = {}
            for cell in tbl.cells:
                if cell.row_index == 0:
                    headers[cell.column_index] = cell.content.strip()

            header_list = [headers.get(i, f"Col_{i}") for i in range(tbl.column_count)]

            # Build table as markdown
            rows_md = []
            rows_md.append("| " + " | ".join(header_list) + " |")
            rows_md.append("| " + " | ".join(["---"] * len(header_list)) + " |")

            row_data: Dict[int, Dict[int, str]] = {}
            for cell in tbl.cells:
                if cell.row_index > 0:
                    row_data.setdefault(cell.row_index, {})[cell.column_index] = cell.content.strip()

            for ri in sorted(row_data.keys()):
                row_cells = [row_data[ri].get(ci, "") for ci in range(tbl.column_count)]
                rows_md.append("| " + " | ".join(row_cells) + " |")

            table_md = "\n".join(rows_md)
            table_text = f"Document: {document_name}\nPage: {page_num}\nTable {tbl_idx}\n\n{table_md}"

            metric_types = detect_metric_types(table_text)

            # Yield TABLE-level chunk
            table_meta = {
                "document_name": document_name,
                "chunk_type": "table",
                "page": page_num,
                "table_index": tbl_idx,
                "report_date": report_date,
                "metric_types": ",".join(metric_types) if metric_types else "",
                "row_count": tbl.row_count - 1,
                "columns": ",".join(header_list[:10]),
                "ingested_at": datetime.utcnow().isoformat(),
            }
            yield table_text, table_meta

    # Extract paragraphs — keep all paragraphs >= 50 chars.
    # metric_types is still stored for debugging but is not used as a gate.
    paragraphs = getattr(result, "paragraphs", []) or []
    for para_idx, para in enumerate(paragraphs):
        content = (getattr(para, "content", "") or "").strip()

        if len(content) < 50:
            continue

        metric_types = detect_metric_types(content)

        regions = getattr(para, "bounding_regions", []) or []
        page_num = regions[0].page_number if regions else 0

        yield content, {
            "document_name": document_name,
            "chunk_type": "paragraph",
            "page": page_num,
            "paragraph_index": para_idx,
            "report_date": report_date,
            "metric_types": ",".join(metric_types),
            "ingested_at": datetime.utcnow().isoformat(),
        }
