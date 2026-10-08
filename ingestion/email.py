"""
Email (.msg) document ingestion.

Extracts:
- Email body text
- HTML tables embedded in email
- File attachments (recursively processed)
"""

import logging
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Generator, List, Optional, Tuple

import pandas as pd

try:
    from ..config import detect_metric_types
except ImportError:
    from config import detect_metric_types

logger = logging.getLogger(__name__)

# Optional imports
try:
    import extract_msg
    HAS_MSG_SUPPORT = True
except ImportError:
    HAS_MSG_SUPPORT = False
    extract_msg = None

try:
    from bs4 import BeautifulSoup
    HAS_BS4 = True
except ImportError:
    HAS_BS4 = False
    BeautifulSoup = None


def extract_tables_from_html(html_content: str) -> List[pd.DataFrame]:
    """Extract tables from HTML content."""
    if not HAS_BS4:
        logger.warning("BeautifulSoup not installed, skipping HTML table extraction")
        return []

    tables = []
    soup = BeautifulSoup(html_content, "html.parser")

    for table in soup.find_all("table"):
        rows = []
        for tr in table.find_all("tr"):
            cells = [td.get_text(strip=True) for td in tr.find_all(["td", "th"])]
            if cells and any(c for c in cells):
                rows.append(cells)

        if len(rows) > 1:
            try:
                # Deduplicate column names
                headers = rows[0]
                seen = {}
                unique_headers = []
                for h in headers:
                    if h in seen:
                        seen[h] += 1
                        unique_headers.append(f"{h}_{seen[h]}")
                    else:
                        seen[h] = 0
                        unique_headers.append(h)

                df = pd.DataFrame(rows[1:], columns=unique_headers)
                if len(df) > 0 and len(df.columns) > 1:
                    tables.append(df)
            except Exception as e:
                logger.debug(f"Could not parse table: {e}")
                continue

    return tables


def iter_email_chunks(
    msg_bytes: bytes,
    document_name: str,
    report_date: Optional[str] = None,
) -> Generator[Tuple[str, Dict[str, Any]], None, None]:
    """
    Yield chunks from Outlook .msg email file.

    Extracts:
    1. Email body text (with tables)
    2. HTML tables as structured data
    3. Attachments (recursively ingested)
    """
    if not HAS_MSG_SUPPORT:
        raise RuntimeError("extract-msg not installed. Run: pip install extract-msg")

    # Import here to avoid circular imports
    try:
        from .excel import iter_excel_chunks
        from .pdf import iter_pdf_chunks
    except ImportError:
        from ingestion.excel import iter_excel_chunks
        from ingestion.pdf import iter_pdf_chunks

    # Write to temp file (extract_msg needs file path)
    with tempfile.NamedTemporaryFile(suffix=".msg", delete=False) as tmp:
        tmp.write(msg_bytes)
        tmp_path = tmp.name

    try:
        msg = extract_msg.Message(tmp_path)

        subject = msg.subject or "No Subject"
        sender = msg.sender or "Unknown"
        date = str(msg.date) if msg.date else ""
        body = msg.body or ""
        html_body = msg.htmlBody or ""

        logger.info(f"Processing email: {subject}")
        logger.info(f"  From: {sender}, Date: {date}")
        logger.info(f"  Attachments: {len(msg.attachments)}")

        # Yield email metadata chunk
        email_summary = f"""Email: {subject}
From: {sender}
Date: {date}

Body:
{body[:5000]}"""

        metric_types = detect_metric_types(email_summary)

        yield email_summary, {
            "document_name": document_name,
            "chunk_type": "email_body",
            "email_subject": subject[:200],
            "email_sender": sender[:100],
            "email_date": date,
            "report_date": report_date,
            "metric_types": ",".join(metric_types) if metric_types else "",
            "ingested_at": datetime.utcnow().isoformat(),
        }

        # Extract tables from HTML body
        if html_body and HAS_BS4:
            tables = extract_tables_from_html(html_body)
            logger.info(f"  Found {len(tables)} tables in email body")

            for tbl_idx, df in enumerate(tables, 1):
                table_md = df.to_markdown(index=False)
                table_text = f"Email: {subject}\nTable {tbl_idx} from email body:\n\n{table_md}"

                tbl_metrics = detect_metric_types(table_text)

                yield table_text, {
                    "document_name": document_name,
                    "chunk_type": "table",
                    "source": "email_body",
                    "email_subject": subject[:200],
                    "table_index": tbl_idx,
                    "report_date": report_date,
                    "metric_types": ",".join(tbl_metrics) if tbl_metrics else "",
                    "row_count": len(df),
                    "columns": ",".join(df.columns.tolist()[:10]),
                    "ingested_at": datetime.utcnow().isoformat(),
                }

        # Process attachments
        for att in msg.attachments:
            att_name = att.longFilename or att.shortFilename or "unknown"
            att_data = att.data

            if not att_data or len(att_data) < 100:
                continue

            att_ext = Path(att_name).suffix.lower()

            # Skip images
            if att_ext in [".png", ".jpg", ".jpeg", ".gif", ".bmp"]:
                continue

            logger.info(f"  Processing attachment: {att_name} ({len(att_data)} bytes)")

            # Recursively process supported attachments
            if att_ext in [".xlsx", ".xls"]:
                for chunk_text, chunk_meta in iter_excel_chunks(att_data, att_name, report_date):
                    chunk_meta["source_email"] = document_name
                    chunk_meta["email_subject"] = subject[:200]
                    yield chunk_text, chunk_meta

            elif att_ext == ".pdf":
                try:
                    for chunk_text, chunk_meta in iter_pdf_chunks(att_data, att_name, report_date):
                        chunk_meta["source_email"] = document_name
                        chunk_meta["email_subject"] = subject[:200]
                        yield chunk_text, chunk_meta
                except Exception as e:
                    logger.error(f"Failed to process PDF attachment {att_name}: {e}")

        msg.close()

    finally:
        try:
            Path(tmp_path).unlink()
        except Exception:
            pass
