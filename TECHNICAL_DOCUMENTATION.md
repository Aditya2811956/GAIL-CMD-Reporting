# CMD Report Automation — Technical Documentation

**Version:** revamp-2
**Audience:** Deployment teams
**Last Updated:** June 2026

---

## Table of Contents

1. [Use Case & Purpose](#1-use-case--purpose)
2. [System Architecture](#2-system-architecture)
3. [Component Deep Dives](#3-component-deep-dives)
   - 3.1 [Ingestion Module](#31-ingestion-module)
   - 3.2 [Extraction Module](#32-extraction-module)
   - 3.3 [Generation Module](#33-generation-module)
   - 3.4 [REST API (app.py)](#34-rest-api-apppy)
4. [Data Flow](#4-data-flow)
5. [Historical Data Management](#5-historical-data-management)
6. [Configuration Reference](#6-configuration-reference)
7. [Deployment Guide](#7-deployment-guide)
8. [Directory Structure](#8-directory-structure)
9. [Known Limitations & Edge Cases](#9-known-limitations--edge-cases)

---

## 1. Use Case & Purpose

GAIL (Gas Authority of India Limited) publishes a monthly **Committee Meeting Dashboard (CMD)** report that summarises key macroeconomic and sector-specific indicators: inflation indices (WPI, CPI, IIP), GDP growth, natural gas statistics, interest rates, GAIL's share price, SATAT/CBG scheme progress, CGD (City Gas Distribution) rollout figures, and latest industry news.

Assembling this report manually requires an analyst to collect data from a dozen government sources (PDFs from MoSPI, PPAC, RBI, MoPNG), extract the relevant numbers, update tables and charts in a Word document, and publish a PDF — a process that takes 2–3 person-days per month.

This system **fully automates** that workflow. Each month, the operator:

1. Uploads the source documents (PDFs, Excel files, Outlook `.msg` emails) through a REST API.
2. Calls a single generation endpoint with the target month.
3. Downloads a finished DOCX/PDF report with all tables, charts, and a news digest updated.

The system handles document parsing, vector-based retrieval, LLM-driven value extraction, chart and table updates inside the Word document, and PDF conversion — with no human intervention in between.

---

## 2. System Architecture

The system is a Python FastAPI application deployed as a Docker container on a Linux VM. It communicates with a separately-hosted ChromaDB vector database and with Azure cloud services (Document Intelligence for PDF OCR, Azure OpenAI for LLM inference and embeddings).

```
┌──────────────────────────────────────────────────────────────────┐
│                          Client / Operator                       │
└────────────────────────────────┬─────────────────────────────────┘
                                 │  HTTP REST API (port 6067 prod)
                         ┌───────▼────────┐
                         │   FastAPI app  │  (Docker container)
                         │    (app.py)    │
                         └──┬──────┬──────┘
                            │      │
               ┌────────────▼──┐  ┌▼───────────────────┐
               │  Ingestion    │  │    Generation       │
               │  Module       │  │    Module           │
               └────────┬──────┘  └──────────┬──────────┘
                        │                    │
               ┌────────▼──────┐  ┌──────────▼──────────┐
               │  Extraction   │  │  Template DOCX /    │
               │  Module       │  │  History JSON files │
               └────────┬──────┘  └─────────────────────┘
                        │
          ┌─────────────┼──────────────────┐
          │             │                  │
   ┌──────▼──────┐  ┌───▼──────┐  ┌───────▼──────────┐
   │  ChromaDB   │  │ Azure    │  │  Web Scrapers    │
   │  (vectors)  │  │ OpenAI   │  │  (RBI, SBI,      │
   │             │  │ (LLM +   │  │   Google News    │
   └─────────────┘  │ Embed)   │  │   x17, PIB, PSU  │
                    └──────────┘  │   Watch, etc.)   │
                                   └──────────────────┘

   Azure Document Intelligence (PDF OCR) used during ingestion
```

### Key Architectural Decisions

**Table-First Retrieval.** Government documents contain highly structured tabular data. Rather than chunking tables row-by-row (which strips headers and context), this system stores each table as a single markdown-formatted chunk. The LLM receives a complete table and can navigate column headers, row labels, and all values simultaneously. This is the single biggest reliability improvement over naive RAG approaches.

**Single Collection, Multi-Month.** All months' documents live in one ChromaDB collection. Every chunk is tagged with a `report_date` metadata field (`YYYY-MM`). Extraction queries always filter on this field, ensuring a September query never retrieves October documents.

**T-1 Data Offset.** Government statistics (WPI, CPI, IIP, GDP) are published one month after the reference period ends. The September 2025 CMD report therefore contains August 2025 figures. The pipeline models this explicitly: `data_month = report_month - 1`. Documents for September are uploaded under `documents/aug_2025/` but tagged `report_date=2025-09`.

**No Hardcoded Fallbacks.** If extraction fails for a metric, the system records `ExtractionResult(success=False)` and logs a clear error. The caller (generation module) decides whether to skip the field or abort. This makes failures visible rather than masking them with stale data.

**Native OOXML Chart Updates.** The DOCX file format is a ZIP archive. Charts are stored as XML files inside it. Rather than regenerating charts with matplotlib and embedding images, the system opens the ZIP, edits the XML data cache directly, and repackages the file. This preserves the chart's native interactivity and formatting.

---

## 3. Component Deep Dives

### 3.1 Ingestion Module

**Location:** `ingestion/`
**Entry point:** `ingestion/service.py → ingest_document()`

The ingestion module receives raw file bytes, parses them based on file type, and writes structured vector embeddings into ChromaDB.

#### Document Types and Parsers

| File Type | Parser | Notes |
|-----------|--------|-------|
| `.pdf` | Azure Document Intelligence | Extracts tables + paragraphs with page numbers |
| `.xlsx` / `.xls` | pandas with smart header detection | Handles irregular government Excel layouts |
| `.msg` | extract-msg + BeautifulSoup | Extracts email body, HTML tables, and attachments |

**PDF Parsing.** Azure Document Intelligence returns a structured JSON response with tables (as cell grids) and paragraphs. The ingestion module converts each detected table to markdown format (column headers as the first row, pipe-delimited), then yields it as a single chunk. Paragraphs are similarly yielded individually. Each chunk is tagged with page number, document name, and inferred metric types (based on keyword matching against `config.METRIC_KEYWORDS`).

**Excel Parsing.** Government Excel files often have irregular structure: multi-row merged headers, blank rows, units on a separate row, titles embedded above the data. The parser uses a heuristic: it scans the first 15 rows of each sheet looking for the row that maximises the number of non-empty, non-numeric cells — that row is treated as the column header. Rows above the header are concatenated into a sheet title string. Each sheet becomes one or more table chunks.

**Email Parsing.** Outlook `.msg` files are processed with the `extract-msg` library. The email body (HTML) is parsed for embedded tables using BeautifulSoup, which converts them to DataFrames, then to markdown. Attachments (PDFs, Excel files) are extracted and recursively processed through the same ingestion pipeline. This handles the common case where data arrives as an Excel attachment inside a forwarded email chain.

#### Chunking and Metadata

Every chunk stored in ChromaDB carries these metadata fields:

| Field | Description | Example |
|-------|-------------|---------|
| `document_name` | Source filename | `WPI.pdf` |
| `chunk_type` | Type of chunk | `table`, `paragraph`, `email_body` |
| `metric_types` | List of metric tags | `["inflation_wpi"]` |
| `report_date` | Report month filter | `2025-09` |
| `page` | Page number (PDFs only) | `3` |
| `sheet_name` | Sheet name (Excel only) | `Sheet1` |

Metric type tags are assigned by matching chunk text against `METRIC_KEYWORDS` in `config.py`. For example, a chunk containing the word "WPI" or "wholesale price" receives the tag `inflation_wpi`. These tags are the primary filter used during extraction, narrowing the ChromaDB search before semantic comparison.

Chunks are upserted (not inserted) using deterministic IDs derived from document name + chunk index, so re-ingesting a document is idempotent.

---

### 3.2 Extraction Module

**Location:** `extraction/`
**Entry point:** `extraction/service.py → MetricExtractor`

The extraction module retrieves relevant chunks from ChromaDB and uses an LLM to parse out structured metric values from them.

#### Metric Configurations

Twelve metrics are defined in `extraction/metrics.py`. Each `MetricConfig` object specifies:

- **query** — a semantic search string (e.g., `"WPI all commodities index value"`)
- **extraction_prompt** — an LLM prompt template with `{table}` and `{period}` placeholders
- **output_schema** — the expected JSON structure the LLM must return
- **period_type** — `"month"`, `"quarter"`, or `"year"` (controls how `{period}` is formatted)
- **document_hint** — optional filename pattern to filter results (e.g., `"WPI"`)
- **chunk_type_filter** — restricts to `"table"`, `"paragraph"`, or `"email_body"`
- **bundle_chunks** — if true, concatenates multiple retrieved chunks before the LLM call

The twelve metrics are: `wpi`, `cpi`, `iip`, `gdp`, `gdp_forecast`, `gas_stats_monthly`, `gas_stats_yearly`, `interest_repo`, `satat_data`, `cgd_data`, `cgd_cumulative`, `retail_lng`.

#### Extraction Flow

```
MetricExtractor.extract(metric_name, report_month)
        │
        ▼
_retrieve_tables()
  - Embeds the metric's query string
  - Queries ChromaDB: filter by metric_types + report_date
  - Applies document_hint filter (soft: skip if no results, retry without)
  - Returns top-N table/paragraph chunks
        │
        ▼
_extract_with_llm(table_text, period_string)
  - Constructs full prompt: system role + extraction instructions + table markdown
  - Calls Azure OpenAI gpt-4o (temperature=0)
  - Parses JSON from response
  - Validates against output_schema
        │
        ▼
ExtractionResult(success, value, source_doc, source_page, error)
```

If the LLM returns malformed JSON or fails validation, the extractor tries the next retrieved chunk. If all chunks fail, `ExtractionResult(success=False)` is returned.

#### Special Cases

**interest_repo** (Repo Rate & MCLR): The primary source is a live web scrape — the RBI monetary policy page for the repo rate and the SBI website for MCLR-1Y. ChromaDB/LLM extraction is used only as fallback. This metric is unlikely to need document upload.

**gdp_forecast**: Extracted from the CMD master Excel if present, otherwise left empty. No hard failure.

**retail_lng**: Extracted from CGD email tables. If unavailable, the column is left blank.

---

### 3.3 Generation Module

**Location:** `generation/`
**Entry point:** `generation/service.py → ReportGenerator.generate()`

The generation module receives a dict of `ExtractionResult` objects, loads a template DOCX file, and produces a fully updated report.

#### Generation Steps

```
ReportGenerator.generate(extraction_results, report_month, template_type)
        │
        ├─► Validate extracted values (three-tier validation)
        │       Schema validation (Pydantic)
        │       Range validation (plausible bounds per metric)
        │       Consistency validation (cross-field sanity)
        │
        ├─► Load / update historical JSON files
        │       Merge new values into gdp_history.json, inflation_history.json, etc.
        │       Deduplicate by period key; trim to last 12 entries
        │
        ├─► Update tables in DOCX
        │       GDP table (12 quarters of history)
        │       SATAT/CBG scheme data table
        │       CGD targets table (entity rows)
        │       CGD cumulative table
        │       Retail LNG stations table
        │
        ├─► Update charts in DOCX (native OOXML)
        │       Inflation chart (WPI, CPI, IIP lines)
        │       Interest rates chart (Repo Rate, MCLR-1Y lines)
        │       Gas monthly chart (bar)
        │       Gas yearly chart (bar)
        │       GAIL share price chart (line + bar combo)
        │
        ├─► Scrape and append news section (non-blocking)
        │       ~22 RSS feeds (Google News topic queries, PIB, PSU Watch,
        │       Renewable Watch, Economic Times, etc.) fetched concurrently
        │       LLM summarisation → numbered table at end of report
        │
        ├─► Save DOCX to output/
        │
        └─► Convert DOCX → PDF via LibreOffice (headless)
```

#### Table Updates

Table update functions locate the target table in the DOCX by scanning for a distinctive header or caption string, then iterate over rows writing new values cell by cell. The `tables.py` helper preserves the original cell shading, borders, and font styles by copying only the text content — not the XML run formatting. Number formatting follows Indian convention (lakh/crore notation: `4,788,000` → `47,88,000`).

`template2.docx` includes a static **Abbreviations** section (heading + two-section table: CGD entity names and general abbreviations) at table index 6. This table is never touched by the generation pipeline and passes through unchanged in every generated report.

**SATAT and CGD tables** have a dual source: the primary path uses `MsgTableExtractor`, which reads the raw `.msg` email files directly from `documents/{month}/` without going through ChromaDB. This bypasses LLM processing entirely for these tables and is faster and more reliable. If the `.msg` files are unavailable (e.g., documents directory not mounted in Docker), the fallback path queries ChromaDB and uses the LLM.

On construction, `MsgTableExtractor` pre-loads every `.msg` file in `documents/{month}/` concurrently via a `ThreadPoolExecutor`, one thread per file, and caches the opened message objects. The pre-load also touches `htmlBody` on each message so the (slow, sequential-by-default) RTF→HTML conversion happens during this parallel phase rather than later, one-at-a-time, during extraction.

#### Chart Updates (OOXML)

A DOCX file is a ZIP archive. Charts inside it are stored at paths like `word/charts/chart1.xml`. The chart XML contains a data cache (`<c:numCache>`, `<c:strCache>`) that stores the last-known values from the linked data source. The chart updater:

1. Opens the DOCX as a ZIP in memory.
2. Iterates over all `word/charts/chart*.xml` files.
3. Identifies which chart to update by matching series names or axis labels.
4. Replaces the numeric and string caches with new data.
5. Removes the `<c:externalData>` element entirely from every chart XML (updated or not) — the template charts link to a Windows network share (`\\Dc-vm-fs\CSPA\...`) that is unreachable outside GAIL's intranet, and setting `autoUpdate="0"` was not enough to stop Word from erroring on open.
6. Strips OLE-object relationships (`http://schemas.../oleObject`) from every `word/charts/_rels/*.rels` file for the same reason.
7. Re-packages the ZIP and saves.

Chart logic lives in `chart_updater.py`; `charts.py` provides the public interface for the generation module.

#### News Table XML Sanitization

Text inserted into the news table (article summaries and source names) is passed through a sanitizer (`_sanitize_xml` in `service.py`) that strips characters illegal in XML 1.0 (control characters, surrogate pairs) before being written into the DOCX. Untreated, these characters corrupt the DOCX and make it unreadable by Word.

#### News Section

`news_scraper.py` (v2, the original `news_scraper.py` is preserved as `news_scraper_v1.py`) covers 9 CMD topic categories — Natural Gas/CBG/CGD, M&A Deals, New Projects, LNG, Renewables & Green Hydrogen, Petrochemicals, Crude, Global Macroeconomy, and Commodity Prices — via ~17 Google News topic-specific RSS queries plus direct feeds (PIB MoPNG/MNRE/MoP, PSU Watch, Renewable Watch, Economic Times, Hindu BusinessLine, Financial Express). The PNGRB HTML scrape and other fragile/blocked sources (Business Standard, Moneycontrol) have been removed.

All RSS sources are fetched **concurrently** via a `ThreadPoolExecutor`, each with retry and exponential backoff. Articles are filtered to those published in the report month, then deduplicated in two passes: exact match on normalized title, then near-duplicate detection via `difflib.SequenceMatcher`.

The LLM receives up to `target_count` (9–18, scaled by how many items were scraped) candidate items and is asked to select the most relevant, covering all 9 topic categories, and tag each with its topic. Each selected item is summarised to a 3–4 sentence paragraph (what happened, key figures/entities, why it matters for GAIL/the energy sector). `max_tokens` is capped at 4096 to stay under the model limit. The LLM's selected items are deduplicated again (on summary text, threshold 0.85) to remove near-duplicate stories the LLM may have reworded. The digest is appended as a formatted table at the end of the report. If scraping or summarisation fails for any reason, this step is logged and skipped — it does not block report generation.

#### Validation

Three validation tiers run before table and chart updates:

1. **Schema validation** — Pydantic models ensure the extracted JSON has the correct field names and types.
2. **Range validation** — each metric has defined plausible bounds (e.g., WPI between −15% and +25%). Values outside these bounds trigger a warning; the report is still generated.
3. **Consistency validation** — cross-field checks (e.g., total gas = domestic + LNG). Significant inconsistencies are logged.

Failed validation does not stop generation. The report is produced with whatever valid data is available; the failure is recorded in the citation/metadata output.

---

### 3.4 REST API (app.py)

The FastAPI application exposes the following endpoints over HTTP. Access control is handled at the infrastructure and network layer — ensure the VM's firewall and network policies are configured before exposing the API externally.

#### Ingestion Endpoints

**`POST /ingest-document`**
Accepts a multipart file upload (PDF, xlsx, or .msg) and an optional `report_date` parameter (`YYYY-MM`). Saves the file to `reference_files/` (or `documents/{month}/` for `.msg` files), then ingests it into ChromaDB. Returns ingestion statistics (chunk counts).

**`GET /get_reference_pdfs`**
Returns a paginated list of all files previously uploaded to `reference_files/`, as base64-encoded content. Used by the UI to browse available documents.

**`POST /selecting_reference_files?report_date=2026-02`**
Re-ingests a specific set of already-uploaded files for a given month. Useful when you want to associate a previously uploaded document with a new `report_date` without re-uploading it.

**`DELETE /reference-documents`**
Accepts a JSON body `{"filenames": [...]}` and hard-deletes each file from `reference_files/` and all of its chunks from ChromaDB. For `.msg` files, also removes the corresponding copy under `documents/{month}/`. Returns `deleted` (per-file disk/Chroma deletion counts) and `not_found` (filenames with no disk file and no ChromaDB chunks).

#### Generation Endpoints

**`POST /generate-report?report_month=2026-01&template=6`**
The main generation endpoint. Runs the full extraction + generation pipeline for the specified month. Returns the generated DOCX file as a download. Also saves a PDF to `output/` as a side effect. Automatically trims old reports, keeping the five most recent.

**`GET /download_report?month=2026-01`**
Serves the most recent PDF generated for the specified month. If `month` is omitted, returns the latest report overall.

#### History Endpoints

**`GET /history`**
Lists available metric history files (`gdp`, `inflation`, `interest_rates`, `gas_monthly`, `gas_yearly`, `gail_share`).

**`GET /history/{metric}`**
Returns the full time-series JSON array for a metric.

**`PUT /history/{metric}`**
Replaces a history file atomically with the provided JSON array. Used to correct historical values without re-running the pipeline.

#### Citation Endpoint

**`GET /citations/{month}`**
Returns extraction metadata for a generated report: for each metric, the source document, table/page, confidence, and whether the value came from live extraction or historical fallback. Useful for audit and debugging.

---

## 4. Data Flow

The end-to-end flow for a monthly report, from raw documents to finished PDF:

```
Operator uploads documents
        │
        ▼
POST /ingest-document (× N files)
        │
        ├─► File saved to reference_files/ (or documents/{month}/ for .msg)
        │
        └─► ingest_document()
                │
                ├─ PDF → Azure Document Intelligence → tables + paragraphs
                ├─ Excel → pandas → sheets as markdown tables
                └─ .msg → extract-msg → body + HTML tables + attachments
                        │
                        ▼
                Chunks tagged with metric_types + report_date
                        │
                        ▼
                Azure OpenAI embeddings (text-embedding-ada-002 or LH-embedding)
                        │
                        ▼
                ChromaDB upsert (collection: cmd_reports_v2)

        ─────────────────────────────────────────────

POST /generate-report?report_month=2026-01
        │
        ▼
MetricExtractor.extract_all()
        │
        ├─► For each of 12 metrics:
        │       ChromaDB semantic search (filtered by metric_types + report_date)
        │       → Full table/paragraph text retrieved
        │       → Sent to gpt-4o with extraction prompt
        │       → Structured JSON returned
        │
        ├─► interest_repo: live web scrape of RBI + SBI pages (primary)
        │
        └─► ExtractionResult dict assembled
                │
                ▼
        ReportGenerator.generate()
                │
                ├─► Validation (schema + range + consistency)
                ├─► History JSON update and persistence
                ├─► Table updates in DOCX (python-docx)
                ├─► Chart updates (OOXML ZIP manipulation)
                ├─► News section (RSS + PNGRB + LLM)
                ├─► Save DOCX to output/
                └─► LibreOffice: DOCX → PDF
                        │
                        ▼
        Generated DOCX returned as HTTP download
        PDF available at GET /download_report
```

---

## 5. Historical Data Management

Seven time-series JSON files are maintained in `output/data/`. These files are not generated fresh each run — they accumulate across all pipeline runs and serve as the long-term store for chart data.

| File | Key | Content |
|------|-----|---------|
| `gdp_history.json` | `quarter` | GDP at current prices, growth % per quarter |
| `gdp_forecast_history.json` | — | Dict: institution → forecast string |
| `inflation_history.json` | `month` | WPI, CPI, IIP per month |
| `interest_rates_history.json` | `month` | Repo rate, MCLR-1Y per month |
| `gas_monthly_history.json` | `month` | Domestic gas, LNG, total per month |
| `gas_yearly_history.json` | `year` | Same fields aggregated annually |
| `gail_share_data_history.json` | `month` | GAIL share price, market cap per month |

Each pipeline run reads the existing file, merges in the new extracted value (deduplicating by key), and writes it back. Charts are always rendered from the last 12 entries of the relevant file.

**Important caveat for backdated runs.** If you run the pipeline for September 2025 *after* already running October and November, the history files will already contain October and November data. The GDP table in the September report will then include future quarters. To avoid this, pass `history_cutoff_month` to `ReportGenerator`, which trims history files to entries on or before the report month before rendering.

---

## 6. Configuration Reference

All configuration is centralised in `config.py`. Runtime values are loaded from environment variables; default values cover the standard VM deployment.

### Environment Variables

| Variable | Required | Default | Description |
|----------|----------|---------|-------------|
| `DI_KEY` | Yes | — | Azure Document Intelligence API key |
| `AZURE_DI_ENDPOINT` | Yes | — | Azure Document Intelligence endpoint URL |
| `LLM_PROVIDER` | No | `azure` | LLM provider: `azure` or `openai` |
| `EMBED_PROVIDER` | No | `azure` | Embedding provider: `azure` or `openai` |
| `AZURE_OPENAI_API_KEY` | Yes (if azure) | — | Azure OpenAI API key |
| `AZURE_OPENAI_ENDPOINT` | Yes (if azure) | — | Azure OpenAI endpoint URL |
| `AZURE_OPENAI_API_VERSION` | No | `2024-02-01` | Azure OpenAI API version |
| `AZURE_OPENAI_EMBED_DEPLOYMENT` | Yes (if azure) | — | Embedding model deployment name |
| `AZURE_OPENAI_LLM_DEPLOYMENT` | No | `gpt-4o` | LLM deployment name |
| `OPENAI_API_KEY` | Yes (if openai) | — | OpenAI API key (direct) |
| `CHROMA_HOST` | No | `127.0.0.1` | ChromaDB hostname |
| `CHROMA_PORT` | No | `8001` | ChromaDB port |
| `CHROMA_COLLECTION` | No | `cmd_reports_v2` | ChromaDB collection name |

Place these in a `.env` file in the project root. The application loads `.env` automatically via python-dotenv.

### ChromaDB Configuration

> **Placeholder for vendor:** The ChromaDB server configuration (host, port, persistence path, and whether it runs as a separate container or process) will be provided separately by the client. Update `CHROMA_HOST` and `CHROMA_PORT` in `.env` to match.

The application requires ChromaDB to be reachable before startup. If ChromaDB is co-located on the same host, `127.0.0.1` works. If it runs in a separate container without host networking, use `host.docker.internal` (Mac/Windows Docker) or the container's service name (Docker Compose).

The collection `cmd_reports_v2` is created automatically on first use if it does not exist.

---

## 7. Deployment Guide

### Prerequisites

- Docker Engine 24+ and Docker Compose v2
- Access to Azure Document Intelligence and Azure OpenAI (or OpenAI direct) API keys
- A running ChromaDB instance (configuration provided separately)
- SSH access to the target Linux VM

### Step 1 — Transfer Code to the VM

From your local machine:

```bash
rsync -avz \
  --exclude='data/' \
  --exclude='.env' \
  --exclude='__pycache__/' \
  --exclude='*.pyc' \
  --exclude='test_reports/' \
  ./revamp-2/ \
  <user>@<vm-ip>:/home/<user>/cmd/revamp-2/
```

The `data/` directory (generated reports, reference files, ChromaDB state) is excluded to preserve existing VM data on subsequent deploys.

### Step 2 — Create Data Directories on the VM

On the VM (first deployment only):

```bash
cd /home/<user>/cmd/revamp-2
mkdir -p data/output/data data/reference_files data/documents
```

### Step 3 — Create the .env File on the VM

SSH into the VM and create `/home/<user>/cmd/revamp-2/.env`:

```bash
# Azure Document Intelligence
DI_KEY=<your-key>
AZURE_DI_ENDPOINT=https://<your-resource>.cognitiveservices.azure.com/

# LLM
LLM_PROVIDER=azure
AZURE_OPENAI_API_KEY=<your-key>
AZURE_OPENAI_ENDPOINT=https://<your-resource>.openai.azure.com
AZURE_OPENAI_API_VERSION=2024-02-01
AZURE_OPENAI_EMBED_DEPLOYMENT=<deployment-name>
AZURE_OPENAI_LLM_DEPLOYMENT=gpt-4o

# ChromaDB — fill in as per client configuration
CHROMA_HOST=127.0.0.1
CHROMA_PORT=<port>
CHROMA_COLLECTION=cmd_reports_v2
```

### Step 4 — Build and Start the Container

```bash
cd /home/<user>/cmd/revamp-2
docker compose -f docker-compose.prod.yml up --build -d
```

The container is named `gail_cmd`, listens on host port `6067`, and restarts automatically on failure.

### Step 5 — Verify the Application

```bash
curl http://localhost:6067/docs
```

The FastAPI auto-generated Swagger UI should respond. Alternatively:

```bash
docker logs gail_cmd -f
```

Look for `Uvicorn running on http://0.0.0.0:8001`.

### Step 6 — Subsequent Deployments

After code updates, re-sync and rebuild:

```bash
# From local machine
rsync -avz --exclude='data/' --exclude='.env' ... revamp-2/ <user>@<vm-ip>:/home/<user>/cmd/revamp-2/

# On VM
cd /home/<user>/cmd/revamp-2
docker compose -f docker-compose.prod.yml up --build -d
```

### Monthly Operation Workflow

Each month, the operator follows this sequence via the REST API (or a UI layer on top):

1. **Upload documents** — `POST /ingest-document` for each source file (WPI.pdf, CPI.pdf, IIP.pdf, GDP.pdf, PPAC Excel, SATAT email `.msg`, CGD email `.msg`, CMD master Excel). Set `report_date=YYYY-MM` to the report month (e.g., `2026-02` for the February report).

2. **Generate report** — `POST /generate-report?report_month=2026-02&template=6`. Wait for the response (typically 2–5 minutes depending on LLM latency).

3. **Download report** — The DOCX is returned directly by the generation endpoint, or retrieve the PDF via `GET /download_report?month=2026-02`.

4. **Review citations** — `GET /citations/2026-02` shows which source document each metric value came from, for audit purposes.

### Local Development (Mac/Windows)

For development without a VM:

```bash
cd revamp-2/
docker compose up --build
```

This uses `docker-compose.yml` (not `docker-compose.prod.yml`), which overrides `CHROMA_HOST` to `host.docker.internal` so the container can reach a ChromaDB instance running on the host machine.

Alternatively, run without Docker:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app:app --reload --port 8001
```

---

## 8. Directory Structure

```
revamp-2/
├── app.py                      # FastAPI application — REST endpoints
├── config.py                   # Centralised config: Azure, OpenAI, ChromaDB, file paths
├── run_pipeline.py             # CLI: ingest → extract → generate (for local/batch use)
├── migrate_collection.py       # Utility: migrate ChromaDB collection between hosts
├── requirements.txt            # Python dependencies
├── Dockerfile                  # Container build (Python 3.11, LibreOffice, fonts)
├── docker-compose.yml          # Local dev (Mac Docker)
├── docker-compose.prod.yml     # VM production deployment
├── .env.example                # Environment variable template
│
├── ingestion/
│   ├── service.py              # ingest_document() — routes to parser, embeds, writes ChromaDB
│   ├── pdf.py                  # Azure Document Intelligence PDF parser
│   ├── excel.py                # Smart-header Excel parser (handles irregular govt layouts)
│   └── email.py                # .msg email parser (body + tables + attachments)
│
├── extraction/
│   ├── service.py              # MetricExtractor class — retrieve + LLM extract
│   ├── metrics.py              # 12 MetricConfig definitions (query, prompt, schema)
│   ├── utils.py                # Period formatting: month/quarter/year → human string
│   └── web_scrapers.py         # Live scrapers for RBI Repo Rate and SBI MCLR-1Y
│
├── generation/
│   ├── service.py              # ReportGenerator — main orchestration
│   ├── schemas.py              # Pydantic data models for validated metric values
│   ├── validation.py           # Three-tier validation (schema, range, consistency)
│   ├── table_updater.py        # GDP history, SATAT, CGD table updates
│   ├── msg_table_extractor.py  # Direct .msg-to-table extraction (bypasses ChromaDB)
│   ├── chart_updater.py        # Native OOXML XML chart data updater
│   ├── tables.py               # Low-level python-docx table helpers
│   ├── news_scraper.py         # Google News RSS + PNGRB scraper + LLM summariser
│   ├── charts.py               # Chart generation utilities
│   └── evaluator.py            # Cell-by-cell DOCX comparison vs ground truth
│
├── templates/
│   ├── template1.docx                       # 12-month report template
│   ├── template2.docx                       # 6-month report template (production default)
│   └── template2_backup_pre_abbrev.docx     # Pre-Abbreviations-section backup of template2.docx
│
├── data/                       # VM-persistent storage (NOT in version control)
│   ├── output/
│   │   ├── data/               # History JSON files (accumulated time series)
│   │   └── CMD_Report_*.docx/.pdf
│   ├── reference_files/        # All uploaded source documents (flat, all months)
│   └── documents/              # .msg files organised by month (auto-created)
│       └── {month_year}/
```

---

## 9. Known Limitations & Edge Cases

### Document Availability

The pipeline is designed to be resilient to missing documents, but missing documents degrade accuracy:

- If the PPAC Excel is not uploaded, `gas_stats_monthly` and `gas_stats_yearly` will fail extraction. The corresponding chart will not update and a warning is logged.
- If the CGD/SATAT `.msg` emails are not available as files (e.g., only uploaded via the API without the Docker volume mount), `MsgTableExtractor` falls back to ChromaDB+LLM extraction for those tables, which is slower and may produce lower-confidence results depending on document structure.
- GDP is a quarterly metric. If no GDP PDF is uploaded for a given month (GDP is not published monthly), the system falls back to the last value in `gdp_history.json`. This is by design.

### Backdated Reports and History Accumulation

As described in Section 5, running a report for an earlier month after later months have already been processed will cause the history files to include future data. Always pass the correct `history_cutoff_month` when regenerating past reports, or manually edit the relevant history JSON via `PUT /history/{metric}` before regenerating.

### Email Reply Chains

CGD emails are often forwarded reply chains. The most recent cumulative table is embedded among older forwarded snapshots. `MsgTableExtractor` uses entity name detection (IGL, MNGL, CUGL, etc.) to identify the correct table. In edge cases where the table structure changes significantly, manual verification against the ground truth is recommended.

### LibreOffice PDF Conversion

PDF generation depends on LibreOffice being installed in the container. The Dockerfile installs it via apt. If LibreOffice is unavailable (e.g., stripped-down base image), PDF generation fails silently — the DOCX is still saved and returned. The `GET /download_report` endpoint serves the PDF; if only DOCX is available, callers should use the generation endpoint's direct DOCX response instead.

### ChromaDB Warm-Up

On first use after deploying the container, ChromaDB may take a few seconds to load persisted data. If the first `/ingest-document` call fails with a connection error, retry after a brief pause. Subsequent calls are stable.

### Interest Rate Web Scrapers

The `interest_repo` metric is scraped live from RBI and SBI websites. These scrapers are sensitive to HTML structure changes on those sites. If scraping fails, the metric falls back to ChromaDB extraction. If both fail, the interest rate chart will not update. Monitor the citations endpoint for `"source": "fallback"` indicators on this metric.

---

*End of documentation.*
