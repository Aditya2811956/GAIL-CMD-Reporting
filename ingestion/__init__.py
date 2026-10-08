"""
Ingestion module for CMD Report Automation.

Handles PDF, Excel, and Email (.msg) document ingestion into ChromaDB.
"""

from .service import ingest_document, IngestionResult
from .excel import iter_excel_chunks
from .pdf import iter_pdf_chunks
from .email import iter_email_chunks

__all__ = [
    "ingest_document",
    "IngestionResult",
    "iter_excel_chunks",
    "iter_pdf_chunks",
    "iter_email_chunks",
]
