"""
CMD Report API — revamp-2

Exposes the same endpoint contract as the deployed POC while using the full
revamp pipeline for ingestion, extraction, and generation. All logic lives in
revamp/; this file is purely the API layer.

Run:
    uvicorn app:app --reload --port 8000

Endpoints (matching POC contract):
    POST /ingest-document           Upload a file → save + ingest into ChromaDB
    GET  /get_reference_pdfs        List uploaded reference files as base64
    POST /selecting_reference_files Ingest pre-uploaded files for a report month
    POST /generate-report           Extract metrics + generate DOCX/PDF
    GET  /get_templates             List template PDFs as base64
    GET  /get_report_as_pdf         Latest generated report as base64
    GET  /get_latest_reports        All generated reports as base64
"""
import base64
import logging
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional
from fastapi import FastAPI, File, Form, HTTPException, Query, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel
from config import ChromaConfig
from run_pipeline import write_citations
from ingestion.service import ingest_document as _ingest, IngestionResult, get_collection
from extraction.service import MetricExtractor, ExtractionResult
from extraction.web_scrapers import (
    fetch_repo_rate,
    fetch_mclr_1y,
    fetch_cpi,
    fetch_iip,
    fetch_gdp,
    fetch_wpi,
    scrape_all_macro_metrics,
)
from generation.service import ReportGenerator, GenerationConfig, GenerationResult
from generation.table_updater import TableUpdater
from generation.msg_table_extractor import MsgTableExtractor
from auth import router as auth_router
# ── directory constants ───────────────────────────────────────────────────────
_HERE = Path(__file__).parent.resolve()
# Dynamic REFERENCE_FILES_DIR: Check project data/reference_files, then /app/data/reference_files
_LOCAL_REF_DIR = _HERE / "data" / "reference_files"
_FALLBACK_REF_DIR = _HERE / "reference_files"
if _LOCAL_REF_DIR.exists():
    REFERENCE_FILES_DIR = _LOCAL_REF_DIR
elif _FALLBACK_REF_DIR.exists():
    REFERENCE_FILES_DIR = _FALLBACK_REF_DIR
else:
    REFERENCE_FILES_DIR = Path("/app/data/reference_files")
REFERENCE_FILES_DIR.mkdir(parents=True, exist_ok=True)

TEMPLATES_DIR = _HERE / "templates"
OUTPUT_DIR    = _HERE / "output"
OUTPUT_DIR.mkdir(exist_ok=True)
DATA_DIR      = OUTPUT_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)

# ── logging ───────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
)
logger = logging.getLogger(__name__)

# ── app ───────────────────────────────────────────────────────────────────────
app = FastAPI(
    title="GAIL CMD",
    redoc_url=None,
    root_path="/gail-cmd",
    docs_url="/docs",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(auth_router)

# ── helpers ───────────────────────────────────────────────────────────────────
_MONTH_NAMES = ["jan", "feb", "mar", "apr", "may", "jun",
                "jul", "aug", "sept", "oct", "nov", "dec"]


def _current_month() -> str:
    today = date.today()
    return f"{today.year}-{today.month:02d}"


def _prev_month(yyyymm: str) -> str:
    y, m = int(yyyymm[:4]), int(yyyymm[5:7])
    return f"{y - 1}-12" if m == 1 else f"{y}-{m - 1:02d}"


def _month_to_docs_folder(report_month: str) -> str:
    """Convert '2026-01' → 'jan_2026' (documents/ subfolder convention)."""
    y, m = int(report_month[:4]), int(report_month[5:7])
    return f"{_MONTH_NAMES[m - 1]}_{y}"


def _latest_ingested_month() -> str:
    """Return the latest YYYY-MM that has a subfolder in documents/, else current month."""
    docs_root = _HERE / "documents"
    candidates = []
    if docs_root.exists():
        for folder in docs_root.iterdir():
            if not folder.is_dir():
                continue
            parts = folder.name.split("_")
            if len(parts) == 2:
                month_name, year = parts
                if month_name in _MONTH_NAMES and year.isdigit():
                    m = _MONTH_NAMES.index(month_name) + 1
                    candidates.append(f"{year}-{m:02d}")
    return max(candidates) if candidates else _current_month()


def _to_b64(path: Path) -> str:
    return base64.b64encode(path.read_bytes()).decode("ascii")


def _latest_pdf() -> Optional[Path]:
    pdfs = sorted(OUTPUT_DIR.glob("*.pdf"), key=lambda p: p.stat().st_mtime, reverse=True)
    return pdfs[0] if pdfs else None


def _trim_old_reports(keep: int = 5) -> None:
    """Delete oldest generated reports beyond `keep` most recent, freeing VM disk space."""
    docx_files = sorted(OUTPUT_DIR.glob("CMD_Report_*.docx"), key=lambda p: p.stat().st_mtime, reverse=True)
    for old in docx_files[keep:]:
        old.unlink(missing_ok=True)
        old.with_suffix(".pdf").unlink(missing_ok=True)
        logger.info(f"Trimmed old report: {old.name}")


# ── response models ───────────────────────────────────────────────────────────
class IngestResult(BaseModel):
    status: str
    doc_type: str
    collection: str
    counts: Dict[str, int]
    details: Optional[Dict[str, Any]] = None


class PdfInline(BaseModel):
    name: str
    size_bytes: int
    modified_iso: str
    file_b64: str


class SelectionIn(BaseModel):
    filenames: List[str]


class SelectionOut(BaseModel):
    accepted: List[str]
    rejected: List[str]
    total: int


class TemplatePdf(BaseModel):
    name: str
    size_bytes: int
    modified_iso: str
    file_b64: str


class DeleteReferenceResult(BaseModel):
    filename: str
    deleted_from_disk: bool
    deleted_from_chroma: int


class DeleteReferencesIn(BaseModel):
    filenames: List[str]


class DeleteReferencesOut(BaseModel):
    deleted: List[DeleteReferenceResult]
    not_found: List[str]


# ── endpoints ─────────────────────────────────────────────────────────────────

@app.post(
    "/ingest-document",
    response_model=IngestResult,
    summary="Ingest a PDF, Excel, or .msg into Chroma",
)
async def ingest_document(
    file: UploadFile = File(..., description="PDF, Excel (.xlsx/.xls), or Outlook email (.msg)"),
    report_date: Optional[str] = Form(None),
    doc_link: Optional[str] = Form(None),
    batch_size: int = Form(128),
):
    """
    Upload a file → save to reference_files/ → ingest into ChromaDB.

    `report_date` (YYYY-MM) tags the document so extraction can filter by month.
    The file is saved to reference_files/ so it also appears in /get_reference_pdfs.
    """
    suffix = Path(file.filename or "").suffix.lower()
    supported = {".pdf", ".xlsx", ".xls", ".msg" ,".docx"}
    if suffix not in supported:
        raise HTTPException(
            status_code=400,
            detail=f"Unsupported file type '{suffix}'. Supported: {', '.join(sorted(supported))}",
        )

    content = await file.read()

    # Persist to reference_files/ so it shows up in /get_reference_pdfs
    dest = REFERENCE_FILES_DIR / file.filename
    dest.write_bytes(content)
    logger.info(f"Saved {file.filename} to reference_files/")

    # For .msg files with a report_date: also save to documents/{month}/ so
    # MsgTableExtractor can find them at generation time with correct month isolation.
    if suffix == ".msg" and report_date:
        month_folder = _HERE / "documents" / _month_to_docs_folder(report_date)
        month_folder.mkdir(parents=True, exist_ok=True)
        (month_folder / file.filename).write_bytes(content)
        logger.info(f"Also saved {file.filename} to documents/{month_folder.name}/")

    # Ingest into ChromaDB
    try:
        result: IngestionResult = _ingest(
            content=content,
            document_name=file.filename,
            file_type=suffix,
            collection_name=None,
            report_date=report_date,
        )
    except Exception as e:
        logger.error(f"Ingestion failed for {file.filename}: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))

    if suffix == ".pdf":
        doc_type = "pdf"
    elif suffix in {".xlsx", ".xls"}:
        doc_type = "excel"
    else:
        doc_type = "email"

    return IngestResult(
        status="ok" if result.success else "error",
        doc_type=doc_type,
        collection=ChromaConfig.COLLECTION,
        counts=result.counts,
        details={
            "report_date": report_date,
            "doc_link": doc_link,
            "error": result.error,
        },
    )


@app.get(
    "/get_reference_pdfs",
    response_model=List[PdfInline],
    summary="Return reference files (name + base64 bytes)",
)
def get_reference_pdfs(
    offset: int = Query(0, ge=0),
    limit: int = Query(20, ge=1, le=200),
    max_file_bytes: int = Query(20 * 1024 * 1024, ge=1),
    max_return_bytes: int = Query(80 * 1024 * 1024, ge=1),
):
    """
    List all files in reference_files/ as base64 payloads, newest first.
    Files arrive here via /ingest-document uploads.
    """
    files = sorted(
        (f for f in REFERENCE_FILES_DIR.iterdir() if f.is_file() and not f.name.startswith(".")),
        key=lambda f: f.stat().st_mtime,
        reverse=True,
    )
    files = files[offset: offset + limit]

    total_bytes = 0
    out: List[PdfInline] = []
    for f in files:
        st = f.stat()
        if st.st_size > max_file_bytes:
            continue
        b = f.read_bytes()
        total_bytes += len(b)
        if total_bytes > max_return_bytes:
            break
        out.append(PdfInline(
            name=f.name,
            size_bytes=st.st_size,
            modified_iso=datetime.fromtimestamp(st.st_mtime).isoformat(),
            file_b64=base64.b64encode(b).decode("ascii"),
        ))
    return out


@app.post(
    "/selecting_reference_files",
    response_model=SelectionOut,
    summary="Select reference files and ingest them for a report month",
)
def selecting_reference_files(
    payload: SelectionIn,
    report_date: Optional[str] = Query(
        None,
        description="Report month YYYY-MM to tag documents with. Defaults to current month.",
    ),
):
    """
    Validate that selected filenames exist in reference_files/, then ingest each
    one into ChromaDB tagged with `report_date`.

    Use this when files were already uploaded (via /ingest-document) and you want
    to re-ingest them for a specific report month.

    `report_date` is an optional query param so the existing frontend body contract
    (just `filenames`) is unchanged.
    """
    effective_date = report_date or _current_month()
    existing = {f.name: f for f in REFERENCE_FILES_DIR.iterdir() if f.is_file()}

    accepted, rejected = [], []

    for fn in payload.filenames:
        safe_name = Path(fn).name
        if safe_name not in existing:
            logger.warning(f"File not found in reference_files/: {safe_name}")
            rejected.append(fn)
            continue

        file_path = existing[safe_name]
        suffix = file_path.suffix.lower()
        try:
            result: IngestionResult = _ingest(
                content=file_path.read_bytes(),
                document_name=safe_name,
                file_type=suffix,
                collection_name=None,
                report_date=effective_date,
            )
            if result.success:
                logger.info(f"Ingested {safe_name} for report_date={effective_date}: {result.counts}")
                accepted.append(safe_name)
            else:
                logger.warning(f"Ingestion failed for {safe_name}: {result.error}")
                rejected.append(fn)
        except Exception as e:
            logger.error(f"Ingestion error for {safe_name}: {e}", exc_info=True)
            rejected.append(fn)

    return SelectionOut(accepted=accepted, rejected=rejected, total=len(payload.filenames))


@app.delete(
    "/reference-documents",
    response_model=DeleteReferencesOut,
    summary="Delete uploaded reference documents from disk and ChromaDB",
)
def delete_reference_documents(payload: DeleteReferencesIn):
    """
    Hard-delete one or more reference documents by filename.

    For each filename: removes the file from reference_files/ and deletes all
    its chunks from ChromaDB. For .msg files also removes the copy in
    documents/<month>/.

    Files not found on disk and with no ChromaDB chunks are reported in
    `not_found` (no error raised).
    """
    collection = get_collection()
    deleted: List[DeleteReferenceResult] = []
    not_found: List[str] = []

    for filename in payload.filenames:
        safe_name = Path(filename).name
        if safe_name != filename or not safe_name:
            not_found.append(filename)
            continue

        disk_path = REFERENCE_FILES_DIR / safe_name
        disk_exists = disk_path.exists()

        results = collection.get(where={"document_name": safe_name})
        chroma_count = len(results["ids"])

        if not disk_exists and chroma_count == 0:
            not_found.append(filename)
            continue

        if chroma_count > 0:
            collection.delete(where={"document_name": safe_name})
            logger.info(f"Deleted {chroma_count} ChromaDB chunks for '{safe_name}'")

        if disk_exists:
            disk_path.unlink()
            logger.info(f"Deleted reference file: {safe_name}")

        if safe_name.endswith(".msg"):
            docs_root = _HERE / "documents"
            if docs_root.exists():
                for month_dir in docs_root.iterdir():
                    msg_copy = month_dir / safe_name
                    if msg_copy.exists():
                        msg_copy.unlink()
                        logger.info(f"Deleted .msg copy from {month_dir.name}/")

        deleted.append(DeleteReferenceResult(
            filename=safe_name,
            deleted_from_disk=disk_exists,
            deleted_from_chroma=chroma_count,
        ))

    return DeleteReferencesOut(deleted=deleted, not_found=not_found)


@app.post("/generate-report", summary="Extract metrics and generate CMD report")
async def generate_report(
    template: str = Query("6", description="Template type: '6' or '12' months"),
    report_month: Optional[str] = Query(
        None,
        description="Report month in YYYY-MM format. Defaults to current month.",
    ),
):
    """
    Full pipeline for a given report month:
    1. Extract all metrics from ChromaDB
    2. Generate DOCX from template
    3. Convert to PDF

    Returns the DOCX as a file download.

    `template` defaults to '6' (6-month template). Adding `report_month` is
    additive and does not change the existing frontend integration.
    """
    effective_month = report_month or _latest_ingested_month()
    data_month = _prev_month(effective_month)

    template_file = "template2.docx"  # always use 6-month template
    template_path = TEMPLATES_DIR / template_file
    if not template_path.exists():
        raise HTTPException(status_code=400, detail=f"Template not found: {template_path}")

    try:
        # Step 1: Extract all metrics
        logger.info(f"Extracting metrics for report_month={effective_month}")
        extractor = MetricExtractor(report_date=effective_month)
        extraction_results_raw = extractor.extract_all(effective_month)

        extraction_results: Dict[str, Any] = {}
        for name, result in extraction_results_raw.items():
            if result.success and result.value:
                extraction_results[name] = result.value
            else:
                logger.warning(f"Extraction skipped for '{name}': {result.error}")

        logger.info(f"Extracted {len(extraction_results)}/{len(extraction_results_raw)} metrics")

        # Step 1b: Live Web Scraping Fallback for missing macroeconomic indicators
        missing_macro = [m for m in ["cpi", "iip", "wpi", "gdp"] if m not in extraction_results]
        if missing_macro:
            logger.info(f"Macro metrics missing from uploaded files: {missing_macro}. Running live web scrapers...")
            try:
                scraped_macro = scrape_all_macro_metrics(effective_month)
                for m in missing_macro:
                    if m in scraped_macro and scraped_macro[m]:
                        extraction_results[m] = scraped_macro[m]
                        logger.info(f"Populated missing metric '{m}' from live scraper")
            except Exception as scrape_err:
                logger.warning(f"Live scraping fallback encountered non-fatal error: {scrape_err}")

        # Step 2: TableUpdater (GDP history + ChromaDB fallback for SATAT/CGD)
        table_updater = TableUpdater(data_dir=DATA_DIR)

        # Step 3: MsgTableExtractor — use reference_files/ (uploaded docs are already there)
        # Fallback: documents/{month}/ for local dev with pre-placed files
        msg_extractor = None
        docs_dir = _HERE / "documents" / _month_to_docs_folder(effective_month)
        msg_dir = docs_dir if docs_dir.exists() else REFERENCE_FILES_DIR
        if any(msg_dir.glob("*.msg")):
            msg_extractor = MsgTableExtractor(msg_dir, report_month=effective_month)
            logger.info(f"MsgTableExtractor initialised from {msg_dir}")
        else:
            logger.info(f"No .msg files found in {msg_dir}, skipping MsgTableExtractor")

        # Step 4: Generate report
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = OUTPUT_DIR / f"CMD_Report_{effective_month}_{timestamp}.docx"
        config = GenerationConfig(
            template_path=template_path,
            output_path=output_path,
            report_month=effective_month,
            data_month=data_month,
            template_type="6",  # always template2.docx (6-month); ignore incoming template param
            generate_pdf=True,
            history_cutoff_month=None,
            persist_history=True,
        )

        generator = ReportGenerator(
            config=config,
            data_dir=DATA_DIR,
            table_updater=table_updater,
            msg_extractor=msg_extractor,
        )

        logger.info(f"DEBUG EXTRACTION RESULTS: {extraction_results}")
        result: GenerationResult = generator.generate(extraction_results, skip_validation=True)

        if not result.success:
            raise HTTPException(status_code=500, detail=result.error or "Generation failed")

        # Write citations JSON alongside the report
        write_citations(extraction_results_raw, effective_month, OUTPUT_DIR, generation_result=result)

        # Keep VM disk lean — trim oldest reports beyond last 5
        _trim_old_reports(keep=5)

        output_path = Path(result.output_path)
        return FileResponse(
            path=str(output_path),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=output_path.name,
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Report generation failed: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))


@app.get(
    "/get_templates",
    response_model=List[TemplatePdf],
    summary="List template PDFs (inline bytes)",
)
def get_templates(
    offset: int = Query(0, ge=0),
    limit: int = Query(20, ge=1, le=200),
    max_file_bytes: int = Query(20 * 1024 * 1024, ge=1),
    max_return_bytes: int = Query(80 * 1024 * 1024, ge=1),
):
    """Return template PDF previews from revamp/templates/ as base64."""
    if not TEMPLATES_DIR.exists():
        raise HTTPException(status_code=500, detail=f"Templates directory not found: {TEMPLATES_DIR}")

    files = sorted(TEMPLATES_DIR.glob("*.pdf"))[offset: offset + limit]

    total_bytes = 0
    out: List[TemplatePdf] = []
    for f in files:
        st = f.stat()
        if st.st_size > max_file_bytes:
            continue
        b = f.read_bytes()
        total_bytes += len(b)
        if total_bytes > max_return_bytes:
            break
        out.append(TemplatePdf(
            name=f.name,
            size_bytes=st.st_size,
            modified_iso=datetime.fromtimestamp(st.st_mtime).isoformat(),
            file_b64=base64.b64encode(b).decode("ascii"),
        ))
    return out


@app.get("/get_report_as_pdf", summary="Return the latest generated report as base64 PDF")
def get_report_as_pdf(
    template: str = Query("6", description="Ignored — always returns the latest generated report"),
):
    """
    Return the most recently generated PDF from output/ as a base64 JSON payload.
    The `template` param is accepted for API compatibility but not used.
    """
    pdf = _latest_pdf()
    if not pdf:
        raise HTTPException(status_code=404, detail="No generated report found in output/")

    today = date.today().strftime("%Y_%m_%d")
    pdf_name = f"{pdf.stem}_{today}.pdf"

    pdf_b64 = _to_b64(pdf)
    return JSONResponse({
        "pdf_filename": pdf_name,
        "pdf_b64": pdf_b64,
        "pdf_data_uri": f"data:application/pdf;base64,{pdf_b64}",
    })


@app.get("/get_latest_reports", summary="Return all generated reports as base64 PDFs")
def get_latest_reports(
    limit: Optional[int] = None,
    delete_after: bool = False,
):
    """
    Return all PDFs in output/, newest first, as base64 payloads.
    Optionally limit to the N most recent, or delete each file after encoding.
    """
    pdfs = sorted(OUTPUT_DIR.glob("*.pdf"), key=lambda p: p.stat().st_mtime, reverse=True)
    if limit is not None:
        pdfs = pdfs[:int(limit)]

    reports = []
    for p in pdfs:
        b64 = _to_b64(p)
        reports.append({
            "pdf_filename": p.name,
            "pdf_b64": b64,
            "pdf_data_uri": f"data:application/pdf;base64,{b64}",
        })
        if delete_after:
            p.unlink(missing_ok=True)

    return {"reports": reports}


@app.get("/download_report", summary="Download a generated report as a PDF file")
def download_report(
    month: Optional[str] = Query(None, description="Report month YYYY-MM (e.g. 2025-09). Defaults to latest."),
):
    """
    Serve a generated PDF directly — open this URL in the browser.
    If `month` is provided, returns the most recent PDF for that month.
    If omitted, returns the overall latest PDF.
    """
    if month:
        matches = sorted(
            OUTPUT_DIR.glob(f"CMD_Report_{month}_*.pdf"),
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )
        if not matches:
            raise HTTPException(status_code=404, detail=f"No report found for {month}. Generate it first via /generate-report?report_month={month}")
        pdf = matches[0]
    else:
        pdf = _latest_pdf()
        if not pdf:
            raise HTTPException(status_code=404, detail="No generated reports found in output/")

    return FileResponse(path=str(pdf), media_type="application/pdf", filename=pdf.name)


# ── history management ────────────────────────────────────────────────────────

_HISTORY_FILES = {
    "gdp":            "gdp_history.json",
    "gdp_forecast":   "gdp_forecast_history.json",
    "inflation":      "inflation_history.json",
    "interest_rates": "interest_rates_history.json",
    "gas_monthly":    "gas_monthly_history.json",
    "gas_yearly":     "gas_yearly_history.json",
    "gail_share":     "gail_share_data_history.json",
}


@app.get("/history/{metric}", summary="Read a metric history JSON")
def get_history(metric: str):
    """
    Return the full history JSON for a metric.

    Available metrics: gdp, gdp_forecast, inflation, interest_rates,
    gas_monthly, gas_yearly, gail_share
    """
    import json
    if metric not in _HISTORY_FILES:
        raise HTTPException(
            status_code=400,
            detail=f"Unknown metric '{metric}'. Available: {list(_HISTORY_FILES)}",
        )
    path = DATA_DIR / _HISTORY_FILES[metric]
    if not path.exists():
        raise HTTPException(status_code=404, detail=f"History file not found for '{metric}'")
    return JSONResponse(json.loads(path.read_text()))


@app.put("/history/{metric}", summary="Replace a metric history JSON")
async def put_history(metric: str, request: Request):
    """
    Replace the full history JSON for a metric.

    Send the complete array (or dict for gdp_forecast) as the request body.
    The existing file is overwritten atomically.

    Example — update inflation history:
        PUT /history/inflation
        Body: [{"month": "2025-01", "CPI": 4.26, "WPI": 2.31, "IIP": 5.1}, ...]
    """
    import json
    if metric not in _HISTORY_FILES:
        raise HTTPException(
            status_code=400,
            detail=f"Unknown metric '{metric}'. Available: {list(_HISTORY_FILES)}",
        )
    try:
        body = await request.json()
    except Exception:
        raise HTTPException(status_code=400, detail="Request body must be valid JSON")

    path = DATA_DIR / _HISTORY_FILES[metric]
    path.write_text(json.dumps(body, indent=2, ensure_ascii=False))
    logger.info(f"History updated: {_HISTORY_FILES[metric]}")
    return {"status": "ok", "metric": metric, "file": _HISTORY_FILES[metric]}


@app.get("/history", summary="List available metric history files")
def list_history():
    """Return the available metric names and whether their history file exists."""
    import json
    result = {}
    for metric, filename in _HISTORY_FILES.items():
        path = DATA_DIR / filename
        if path.exists():
            data = json.loads(path.read_text())
            result[metric] = {
                "file": filename,
                "exists": True,
                "entries": len(data) if isinstance(data, list) else len(data.keys()),
            }
        else:
            result[metric] = {"file": filename, "exists": False, "entries": 0}
    return result


# ── citations ─────────────────────────────────────────────────────────────────

@app.get("/citations/{month}", summary="View extraction citations for a report month")
def get_citations(month: str):
    """
    Return the citations JSON for a given report month (e.g. 2025-09).
    Shows which source document and table each extracted value came from.
    """
    import json
    citations_file = OUTPUT_DIR / f"citations_{month}.json"
    if not citations_file.exists():
        raise HTTPException(
            status_code=404,
            detail=f"No citations found for {month}. Run /generate-report?report_month={month} first.",
        )
    return JSONResponse(json.loads(citations_file.read_text()))


# ── live web scraping ──────────────────────────────────────────────────────────

@app.post("/api/scrape-live-data", summary="Trigger live web scrapers for MoSPI, WPI, and Interest Rates")
async def scrape_live_data(
    report_month: Optional[str] = Query(None, description="Report month in YYYY-MM format. Defaults to current month."),
):
    """
    Execute live web scrapers for:
      - MoSPI CPI (Consumer Price Index)
      - MoSPI IIP (Index of Industrial Production)
      - MoSPI GDP (Quarterly Constant Prices GDP)
      - WPI (Wholesale Price Index)
      - RBI Policy Repo Rate
      - SBI 1-Year MCLR Rate

    Returns all successfully scraped values. Missing metrics return None gracefully without failing.
    """
    effective_month = report_month or _current_month()
    logger.info(f"Triggering live web scrapers for month={effective_month} via API...")
    scraped_data = scrape_all_macro_metrics(effective_month)

    return {
        "status": "ok",
        "report_month": effective_month,
        "scraped_metrics": scraped_data,
        "total_scraped": len(scraped_data),
    }


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
