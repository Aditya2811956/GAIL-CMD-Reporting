"""
Main extraction service.

Table-first metric extraction from ChromaDB.
"""

import json
import logging
import re
from dataclasses import dataclass
from datetime import datetime
from dateutil.relativedelta import relativedelta
from typing import Any, Dict, List, Optional, Tuple

import chromadb
from openai import AzureOpenAI, OpenAI

try:
    from ..config import AzureConfig, OpenAIConfig, LLMConfig, ChromaConfig
except ImportError:
    from config import AzureConfig, OpenAIConfig, LLMConfig, ChromaConfig

try:
    from .metrics import METRIC_CONFIGS
    from .utils import get_period
    from .web_scrapers import fetch_repo_rate, fetch_mclr_1y
except ImportError:
    from extraction.metrics import METRIC_CONFIGS
    from extraction.utils import get_period
    from extraction.web_scrapers import fetch_repo_rate, fetch_mclr_1y

logger = logging.getLogger(__name__)


@dataclass
class ExtractionResult:
    """Result of a metric extraction."""
    metric: str
    success: bool
    value: Optional[Dict[str, Any]] = None
    error: Optional[str] = None
    source_document: Optional[str] = None
    source_chunk_meta: Optional[Dict[str, Any]] = None
    source_table_preview: Optional[str] = None
    raw_llm_response: Optional[str] = None


class MetricExtractor:
    """
    Table-First metric extractor.

    Retrieves full tables filtered by metric type, then uses LLM to extract values.
    """

    def __init__(
        self,
        collection_name: Optional[str] = None,
        chroma_host: Optional[str] = None,
        chroma_port: Optional[int] = None,
        report_date: Optional[str] = None,
    ):
        self.collection_name = collection_name or ChromaConfig.COLLECTION
        self.report_date = report_date

        # Initialize ChromaDB
        try:
            self.chroma_client = chromadb.HttpClient(
                host=chroma_host or ChromaConfig.HOST,
                port=chroma_port or ChromaConfig.PORT,
            )
            self.collection = self.chroma_client.get_or_create_collection(self.collection_name)
        except Exception as e:
            logger.warning(f"ChromaDB connection unavailable ({e}). Falling back to live web scrapers and history files.")
            self.chroma_client = None
            self.collection = None

        # Initialize LLM client
        self.llm_provider = LLMConfig.PROVIDER
        if self.llm_provider == "openai":
            self.llm = OpenAI(api_key=OpenAIConfig.API_KEY)
            self.llm_model = OpenAIConfig.LLM_MODEL
            logger.info(f"Using OpenAI LLM (model: {self.llm_model})")
        else:
            self.llm = AzureOpenAI(
                azure_endpoint=AzureConfig.OPENAI_ENDPOINT,
                api_key=AzureConfig.OPENAI_API_KEY,
                api_version=AzureConfig.OPENAI_API_VERSION,
            )
            self.llm_model = AzureConfig.LLM_DEPLOYMENT
            logger.info(f"Using Azure OpenAI LLM (deployment: {self.llm_model})")

        # Initialize separate embed client if EMBED_PROVIDER differs from LLM_PROVIDER
        self.embed_provider = LLMConfig.EMBED_PROVIDER
        if self.embed_provider == self.llm_provider:
            self.embed_client = self.llm
        elif self.embed_provider == "openai":
            self.embed_client = OpenAI(api_key=OpenAIConfig.API_KEY)
            logger.info("Using OpenAI for embeddings (separate from LLM provider)")
        else:
            self.embed_client = AzureOpenAI(
                azure_endpoint=AzureConfig.OPENAI_ENDPOINT,
                api_key=AzureConfig.OPENAI_API_KEY,
                api_version=AzureConfig.OPENAI_API_VERSION,
            )
            logger.info("Using Azure OpenAI for embeddings (separate from LLM provider)")

        logger.info(f"MetricExtractor initialized with collection: {self.collection_name}")

    def _get_query_embedding(self, text: str) -> List[float]:
        """Generate embedding for query text using the embed client (may differ from LLM client)."""
        model = OpenAIConfig.EMBED_MODEL if self.embed_provider == "openai" else AzureConfig.EMBED_DEPLOYMENT
        response = self.embed_client.embeddings.create(model=model, input=[text])
        return response.data[0].embedding

    def _retrieve_tables(
        self,
        query: str,
        metric_type: Optional[str] = None,
        n_results: int = 3,
        document_hint: Optional[str] = None,
        chunk_type_filter: Optional[str] = None,
    ) -> List[Tuple[str, Dict]]:
        """
        Retrieve relevant tables from ChromaDB.

        Returns list of (document_text, metadata) tuples.
        Both document_hint and metric_type are soft filters: each is applied only
        when it narrows results; otherwise it is skipped with a warning.
        """
        if self.collection is None:
            return []

        # Generate query embedding using same model as ingestion
        query_embedding = self._get_query_embedding(query)

        # Build where clause — chunk_type_filter is applied at query time (hard filter)
        conditions = []
        if self.report_date:
            conditions.append({"report_date": {"$eq": self.report_date}})
        if chunk_type_filter:
            conditions.append({"chunk_type": {"$eq": chunk_type_filter}})
        if len(conditions) == 1:
            where = conditions[0]
        elif len(conditions) > 1:
            where = {"$and": conditions}
        else:
            where = None

        results = self.collection.query(
            query_embeddings=[query_embedding],
            n_results=n_results * 4,
            where=where,
            include=["documents", "metadatas"],
        )

        candidates = list(zip(results["documents"][0], results["metadatas"][0]))

        # Log what ChromaDB returned before filtering
        for doc, meta in candidates:
            logger.debug(
                f"  candidate: {meta.get('document_name','?')} | "
                f"metric_types={meta.get('metric_types','?')} | "
                f"chunk_type={meta.get('chunk_type','?')} | "
                f"preview={doc[:80].replace(chr(10),' ')!r}"
            )

        # Soft document_hint filter
        if document_hint:
            hinted = [(d, m) for d, m in candidates
                      if document_hint in m.get("document_name", "").lower()]
            if hinted:
                candidates = hinted
            else:
                logger.warning(f"document_hint='{document_hint}' matched no filenames — skipping hint")

        # Soft metric_type filter
        if metric_type:
            typed = [(d, m) for d, m in candidates
                     if metric_type in m.get("metric_types", "")]
            if typed:
                candidates = typed
            else:
                logger.warning(f"metric_type='{metric_type}' matched no chunks — skipping type filter")

        return candidates[:n_results]

    def _extract_with_llm(
        self,
        table_content: str,
        prompt_template: str,
        period: str,
        output_schema: Dict[str, Any],
    ) -> Tuple[Optional[Dict], str]:
        """
        Use LLM to extract structured data from table.

        Returns (extracted_dict, raw_response).
        """
        prompt = prompt_template.format(table=table_content, period=period)

        schema_hint = json.dumps(output_schema, indent=2)
        full_prompt = f"""{prompt}
Expected format:
{schema_hint}

Important:
- Return ONLY valid JSON, no markdown code blocks
- Use null for missing values
- Numbers should be actual numbers, not strings"""

        logger.debug(f"Full prompt sent to LLM:\n{full_prompt}")

        raw_response = ""
        try:
            response = self.llm.chat.completions.create(
                model=self.llm_model,
                messages=[
                    {"role": "system", "content": "You are a data extraction assistant. Extract structured data from tables and return valid JSON only."},
                    {"role": "user", "content": full_prompt},
                ],
                temperature=0,
                max_tokens=1000,
            )

            raw_response = response.choices[0].message.content.strip()
            logger.debug(f"Raw LLM response:\n{raw_response}")

            # Clean up response
            cleaned = raw_response
            if cleaned.startswith("```"):
                cleaned = re.sub(r"^```(?:json)?\n?", "", cleaned)
                cleaned = re.sub(r"\n?```$", "", cleaned)

            extracted = json.loads(cleaned)
            return extracted, raw_response

        except json.JSONDecodeError as e:
            logger.error(f"JSON parse error: {e}\nRaw response: {raw_response}")
            return None, raw_response
        except Exception as e:
            logger.error(f"LLM extraction error: {e}")
            return None, str(e)

    def extract(
        self,
        metric: str,
        report_month: str,
        n_tables: int = 2,
    ) -> ExtractionResult:
        """
        Extract a metric value for the given report month.

        Args:
            metric: Metric name (e.g., "wpi", "cpi", "gdp")
            report_month: Report month in YYYY-MM format
            n_tables: Number of tables to retrieve and try

        Returns:
            ExtractionResult with extracted value or error
        """
        config = METRIC_CONFIGS.get(metric)
        if not config:
            return ExtractionResult(
                metric=metric,
                success=False,
                error=f"Unknown metric: {metric}. Available: {list(METRIC_CONFIGS.keys())}"
            )

        # Data month is T-1: documents contain last month's published stats
        data_month = (datetime.strptime(report_month, "%Y-%m") - relativedelta(months=1)).strftime("%Y-%m")
        period = get_period(data_month, config.period_type)
        logger.info(f"Extracting {config.name} for {period} (report_month={report_month})")

        # Web scraper path: Repo Rate and MCLR fetched live, not from ChromaDB
        if metric == "interest_repo":
            repo_rate = fetch_repo_rate()
            mclr_1y = fetch_mclr_1y()
            if repo_rate is not None and mclr_1y is not None:
                return ExtractionResult(
                    metric=metric,
                    success=True,
                    value={"month": data_month, "repo_rate": repo_rate, "mclr_1y": mclr_1y},
                    source_document="web:rbi.org.in + sbi.bank.in",
                )
            logger.warning("Web scrape failed for interest_repo — falling back to ChromaDB")

        effective_n = config.n_results_override if config.n_results_override else n_tables
        query = f"{config.query} {period}" if config.append_period_to_query else config.query
        tables = self._retrieve_tables(
            query=query,
            metric_type=config.metric_type if config.metric_type else None,
            n_results=effective_n,
            document_hint=config.document_hint,
            chunk_type_filter=config.chunk_type_filter,
        )

        if not tables:
            return ExtractionResult(
                metric=metric,
                success=False,
                error=f"No tables found for metric_type={config.metric_type}"
            )

        logger.info(f"Found {len(tables)} relevant chunks")

        # Bundle mode: concatenate all chunks into one LLM call
        if config.bundle_chunks and len(tables) > 1:
            combined = "\n\n---\n\n".join(doc for doc, _ in tables)
            meta = tables[0][1]
            tables = [(combined, meta)]

        last_error = None
        for table_content, metadata in tables:
            doc_name = metadata.get("document_name", "unknown")
            logger.info(f"Trying extraction from: {doc_name}")

            extracted, raw_response = self._extract_with_llm(
                table_content=table_content,
                prompt_template=config.extraction_prompt,
                period=period,
                output_schema=config.output_schema,
            )

            if extracted:
                logger.debug(f"Extracted value: {extracted}")
                return ExtractionResult(
                    metric=metric,
                    success=True,
                    value=extracted,
                    source_document=doc_name,
                    source_chunk_meta=metadata,
                    source_table_preview=table_content[:500] + "..." if len(table_content) > 500 else table_content,
                    raw_llm_response=raw_response,
                )
            else:
                last_error = raw_response

        return ExtractionResult(
            metric=metric,
            success=False,
            error=f"Extraction failed from all {len(tables)} tables. Last error: {last_error}",
        )

    def extract_all(
        self,
        report_month: str,
        metrics: Optional[List[str]] = None,
    ) -> Dict[str, ExtractionResult]:
        """
        Extract all (or specified) metrics for a report month.

        Args:
            report_month: Report month in YYYY-MM format
            metrics: List of metrics to extract (default: all)

        Returns:
            Dict mapping metric name to ExtractionResult
        """
        if metrics is None:
            metrics = list(METRIC_CONFIGS.keys())

        results = {}
        for metric in metrics:
            logger.info(f"\n{'='*50}")
            logger.info(f"Extracting: {metric}")
            logger.info(f"{'='*50}")

            result = self.extract(metric, report_month)
            results[metric] = result

            if result.success:
                logger.info(f"✓ {metric}: {result.value}")
            else:
                logger.warning(f"✗ {metric}: {result.error}")

        return results
