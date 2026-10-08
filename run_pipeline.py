"""
End-to-end pipeline for CMD report generation.

Usage:
    python run_pipeline.py --month 2025-11 --dir dec_2025 --template-type 12
"""

import argparse
import json
import logging
import os
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional

from dateutil.relativedelta import relativedelta

# Add parent to path
sys.path.insert(0, str(Path(__file__).parent))

logging.basicConfig(
    level=getattr(logging, os.getenv("LOG_LEVEL", "INFO")),
    format="%(asctime)s | %(levelname)s | %(message)s"
)
logger = logging.getLogger(__name__)

# Document base path
DOCS_BASE = Path(__file__).parent / "documents"
TEMPLATES_BASE = Path(__file__).parent / "templates"


def ingest_documents(docs_dir: Path, collection_name: str, report_date: str):
    """Ingest all documents from a directory into ChromaDB."""
    from ingestion import ingest_document

    logger.info(f"=" * 60)
    logger.info(f"PHASE 1: INGESTION")
    logger.info(f"=" * 60)
    logger.info(f"Documents directory: {docs_dir}")
    logger.info(f"Collection: {collection_name}")

    # Find all supported files
    files = []
    for ext in ["*.pdf", "*.xlsx", "*.xls", "*.msg"]:
        files.extend(docs_dir.glob(ext))

    logger.info(f"Found {len(files)} documents to ingest")

    results = []
    total_chunks = 0

    for file_path in files:
        logger.info(f"\nIngesting: {file_path.name}")

        with open(file_path, "rb") as f:
            content = f.read()

        result = ingest_document(
            content=content,
            document_name=file_path.name,
            file_type=file_path.suffix.lower(),
            collection_name=collection_name,
            report_date=report_date,
        )

        if result.success:
            chunk_count = sum(result.counts.values())
            total_chunks += chunk_count
            logger.info(f"  SUCCESS: {result.counts}")
        else:
            logger.error(f"  FAILED: {result.error}")

        results.append(result)

    logger.info(f"\nIngestion complete: {total_chunks} total chunks")
    return results


def extract_metrics(report_month: str, collection_name: str, report_date: str = ""):
    """Extract all metrics for the report month."""
    from extraction import MetricExtractor, METRIC_CONFIGS

    logger.info(f"\n" + "=" * 60)
    logger.info(f"PHASE 2: EXTRACTION")
    logger.info(f"=" * 60)
    logger.info(f"Report month: {report_month}")
    logger.info(f"Collection: {collection_name}")
    if report_date:
        logger.info(f"Filtering ChromaDB by report_date: {report_date}")

    extractor = MetricExtractor(collection_name=collection_name, report_date=report_date or None)

    # Extract all configured metrics
    metrics_to_extract = [
        "wpi", "cpi", "iip", "gdp", "gdp_forecast",
        "interest_repo", "gas_stats_monthly", "gas_stats_yearly",
        "satat_data", "cgd_data", "cgd_cumulative",
    ]

    results = {}
    for metric in metrics_to_extract:
        if metric not in METRIC_CONFIGS:
            logger.warning(f"Metric config not found: {metric}")
            continue

        logger.info(f"\nExtracting: {metric}")
        result = extractor.extract(metric, report_month)

        if result.success:
            logger.info(f"  SUCCESS: {json.dumps(result.value, indent=2)[:200]}...")
        else:
            logger.warning(f"  FAILED: {result.error}")

        results[metric] = result

    return results


def write_citations(extraction_results: dict, report_month: str, output_dir: Path, generation_result=None):
    """Write a citations JSON summarising the source and location for each metric.

    Each citation has two fields:
      - extraction_attempted: where the LLM tried to get the value (ChromaDB / web scrape)
      - value_used_in_report:  what actually ended up in the report (may be history fallback)
    """
    from extraction import METRIC_CONFIGS

    # sources_used from generation tells us what was actually used (overrides extraction for fallbacks)
    gen_sources = generation_result.sources_used if generation_result else {}

    citations = []
    for metric, result in extraction_results.items():
        cfg = METRIC_CONFIGS.get(metric)
        friendly_name = cfg.name if cfg else metric

        # --- extraction_attempted ---
        if not result.success:
            extraction_attempted = "failed"
            extraction_location = None
        else:
            extraction_attempted = result.source_document or "unknown"
            meta = result.source_chunk_meta or {}
            extraction_location = None
            if extraction_attempted.startswith("web:"):
                extraction_location = "live web scrape"
            elif "page" in meta:
                extraction_location = f"Page {meta['page']}"
            elif "sheet_name" in meta:
                loc = f"Sheet: {meta['sheet_name']}"
                if meta.get("sheet_title"):
                    loc += f" ({meta['sheet_title']})"
                extraction_location = loc

        # --- value_used_in_report ---
        # If generation recorded a fallback/override for this metric, show that.
        # "extraction" means generation used the extracted value directly.
        gen_source = gen_sources.get(metric)
        if gen_source and gen_source != "extraction":
            value_used = gen_source
        elif result.success:
            value_used = f"{extraction_attempted}" + (f" [{extraction_location}]" if extraction_location else "")
        else:
            value_used = "not used (extraction failed)"

        citations.append({
            "metric": metric,
            "name": friendly_name,
            "extraction_attempted": extraction_attempted,
            "extraction_location": extraction_location,
            "extracted_value": result.value if result.success else None,
            "value_used_in_report": value_used,
        })

    output = {
        "report_month": report_month,
        "generated_at": datetime.now().isoformat(),
        "citations": citations,
    }

    citations_path = output_dir / f"citations_{report_month}.json"
    with open(citations_path, "w") as f:
        json.dump(output, f, indent=2)

    logger.info(f"\nCitations written to: {citations_path}")

    # Pretty-print to console
    logger.info("\n--- METRIC CITATIONS ---")
    logger.info(f"  {'Metric':<35} {'Extracted From':<45} {'Used In Report'}")
    logger.info(f"  {'-'*35} {'-'*45} {'-'*50}")
    for c in citations:
        ext = c["extraction_attempted"] or "—"
        if c["extraction_location"]:
            ext += f" [{c['extraction_location']}]"
        used = c["value_used_in_report"] or "—"
        logger.info(f"  {c['name']:<35} {ext:<45} {used}")


def validate_and_generate(
    extraction_results: dict,
    report_month: str,
    template_type: str,
    output_dir: Path,
    data_month: str = "",
    docs_dir: Optional[Path] = None,
):
    """Validate extraction results and generate report."""
    from generation import (
        validate_extraction_results,
        generate_validation_report,
        GenerationConfig,
        ReportGenerator,
    )

    logger.info(f"\n" + "=" * 60)
    logger.info(f"PHASE 3: VALIDATION & GENERATION")
    logger.info(f"=" * 60)

    # Convert ExtractionResult objects to data
    data_for_validation = {}
    for metric, result in extraction_results.items():
        if hasattr(result, "value") and result.value:
            data_for_validation[metric] = result.value

    # Validate
    logger.info("\nValidating extraction results...")
    validation = validate_extraction_results(
        data_for_validation,
        report_month,
        strict=False
    )

    logger.info(generate_validation_report(validation))

    if not validation.valid:
        logger.error("Validation failed. Stopping.")
        return None

    # Find template
    template_name = f"template{1 if template_type == '12' else 2}.docx"
    template_path = TEMPLATES_BASE / template_name

    # Alternative template locations
    if not template_path.exists():
        template_path = Path(__file__).parent / "templates" / template_name
    if not template_path.exists():
        template_path = Path(__file__).parent.parent / f"CMD Dashboard_Nov'25.docx"
    if not template_path.exists():
        # Use the document from dec_2025 as template
        template_path = DOCS_BASE / "dec_2025" / "CMD Dashboard_Nov'25.docx"

    if not template_path.exists():
        logger.error(f"Template not found: {template_path}")
        logger.info("Skipping report generation - validation results available")
        return validation

    # Generate report
    logger.info(f"\nGenerating report from template: {template_path}")

    output_path = output_dir / f"CMD_Report_{report_month}_{datetime.now():%Y%m%d_%H%M%S}.docx"

    config = GenerationConfig(
        template_path=template_path,
        output_path=output_path,
        report_month=report_month,
        template_type=template_type,
        data_month=data_month,
    )

    # Wire up direct email table extractor if docs_dir provided
    msg_extractor = None
    if docs_dir and docs_dir.exists():
        from generation.msg_table_extractor import MsgTableExtractor
        msg_extractor = MsgTableExtractor(docs_dir)
        logger.info(f"MsgTableExtractor initialised for {docs_dir.name}")

    generator = ReportGenerator(config, data_dir=output_dir / "data", msg_extractor=msg_extractor)

    # Skip validation since we already did it
    result = generator.generate(data_for_validation, skip_validation=True)

    if result.success:
        logger.info(f"\nReport generated: {result.output_path}")
        logger.info(f"  Tables updated: {result.tables_updated}")
        logger.info(f"  Charts generated: {result.charts_generated}")
    else:
        logger.error(f"Generation failed: {result.error}")

    return result


def main():
    parser = argparse.ArgumentParser(description="CMD Report Pipeline")
    parser.add_argument("--month", required=True, help="Report month (YYYY-MM)")
    parser.add_argument("--dir", required=True, help="Documents directory name")
    parser.add_argument("--template-type", default="12", choices=["6", "12"])
    parser.add_argument("--collection", default="cmd_reports_v2", help="ChromaDB collection")
    parser.add_argument("--skip-ingest", action="store_true", help="Skip ingestion")
    parser.add_argument("--output-dir", default="./output", help="Output directory")
    args = parser.parse_args()

    docs_dir = DOCS_BASE / args.dir
    if not docs_dir.exists():
        logger.error(f"Documents directory not found: {docs_dir}")
        return 1

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    # Compute data_month (T-1): source documents contain previous month's data
    report_dt = datetime.strptime(args.month, "%Y-%m")
    data_dt = report_dt - relativedelta(months=1)
    data_month = data_dt.strftime("%Y-%m")
    logger.info(f"Report month: {args.month}, Data month (T-1): {data_month}")

    # Phase 1: Ingestion
    if not args.skip_ingest:
        ingest_results = ingest_documents(docs_dir, args.collection, args.month)
        if not any(r.success for r in ingest_results):
            logger.error("All ingestion failed!")
            return 1

    # Phase 2: Extraction (use data_month since documents contain T-1 data,
    # but filter ChromaDB by report_date=args.month since that's how ingestion tagged them)
    extraction_results = extract_metrics(data_month, args.collection, report_date=args.month)
    if not extraction_results:
        logger.error("Extraction failed!")
        return 1

    # Phase 3: Validation & Generation (report_month for output naming, data_month for chart entries)
    result = validate_and_generate(
        extraction_results,
        args.month,
        args.template_type,
        output_dir,
        data_month=data_month,
        docs_dir=docs_dir,
    )

    # Write citations after generation so sources_used (fallbacks) is available
    write_citations(extraction_results, args.month, output_dir, generation_result=result)

    logger.info("\n" + "=" * 60)
    logger.info("PIPELINE COMPLETE")
    logger.info("=" * 60)

    return 0


if __name__ == "__main__":
    sys.exit(main())
