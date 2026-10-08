"""
Main ingestion service.

Orchestrates document ingestion into ChromaDB.
"""

import hashlib
import logging
from dataclasses import dataclass
from typing import Dict, List, Optional, Set

import chromadb
from openai import AzureOpenAI, OpenAI

try:
    from ..config import AzureConfig, OpenAIConfig, LLMConfig, ChromaConfig
except ImportError:
    from config import AzureConfig, OpenAIConfig, LLMConfig, ChromaConfig

try:
    from .excel import iter_excel_chunks
    from .pdf import iter_pdf_chunks
    from .email import iter_email_chunks
except ImportError:
    from ingestion.excel import iter_excel_chunks
    from ingestion.pdf import iter_pdf_chunks
    from ingestion.email import iter_email_chunks

logger = logging.getLogger(__name__)


# Singleton clients
_azure_openai_client: Optional[AzureOpenAI] = None
_openai_client: Optional[OpenAI] = None
_chroma_client: Optional[chromadb.HttpClient] = None


def get_azure_openai_client() -> AzureOpenAI:
    """Get or create Azure OpenAI client."""
    global _azure_openai_client
    if _azure_openai_client is None:
        _azure_openai_client = AzureOpenAI(
            azure_endpoint=AzureConfig.OPENAI_ENDPOINT,
            api_key=AzureConfig.OPENAI_API_KEY,
            api_version=AzureConfig.OPENAI_API_VERSION
        )
    return _azure_openai_client


def get_openai_client() -> OpenAI:
    """Get or create OpenAI client (non-Azure)."""
    global _openai_client
    if _openai_client is None:
        _openai_client = OpenAI(api_key=OpenAIConfig.API_KEY)
    return _openai_client


def get_chroma_client() -> chromadb.HttpClient:
    """Get or create ChromaDB client."""
    global _chroma_client
    if _chroma_client is None:
        _chroma_client = chromadb.HttpClient(
            host=ChromaConfig.HOST,
            port=ChromaConfig.PORT
        )
    return _chroma_client


def get_collection(name: Optional[str] = None):
    """Get or create ChromaDB collection."""
    client = get_chroma_client()
    return client.get_or_create_collection(name or ChromaConfig.COLLECTION)


def get_embeddings(texts: List[str], batch_size: int = 16) -> List[List[float]]:
    """Generate embeddings using EMBED_PROVIDER (may differ from LLM_PROVIDER)."""
    vectors = []

    if LLMConfig.EMBED_PROVIDER == "openai":
        client = get_openai_client()
        model = OpenAIConfig.EMBED_MODEL
        logger.info(f"Using OpenAI for embeddings (model: {model})")
    else:
        client = get_azure_openai_client()
        model = AzureConfig.EMBED_DEPLOYMENT
        logger.info(f"Using Azure OpenAI for embeddings (deployment: {model})")

    for i in range(0, len(texts), batch_size):
        batch = texts[i:i + batch_size]
        response = client.embeddings.create(model=model, input=batch)
        vectors.extend([r.embedding for r in response.data])
    return vectors


def make_chunk_id(doc_name: str, chunk_type: str, text: str) -> str:
    """Create deterministic chunk ID. Note: caller passes text[:200], so two chunks sharing the same first 200 chars will collide."""
    raw = f"{doc_name}|{chunk_type}|{text}"
    return hashlib.sha256(raw.encode()).hexdigest()[:32]


@dataclass
class IngestionResult:
    """Result of document ingestion."""
    success: bool
    document_name: str
    counts: Dict[str, int]
    error: Optional[str] = None


def ingest_document(
    content: bytes,
    document_name: str,
    file_type: str,
    collection_name: Optional[str] = None,
    report_date: Optional[str] = None,
    batch_size: int = 50,
) -> IngestionResult:
    """
    Ingest a document into ChromaDB.

    Args:
        content: File content as bytes
        document_name: Name of the document
        file_type: File extension (e.g., ".pdf", ".xlsx", ".msg")
        collection_name: ChromaDB collection name (optional)
        report_date: Report date in YYYY-MM format (optional)
        batch_size: Batch size for embedding generation

    Returns:
        IngestionResult with counts of ingested chunks
    """
    try:
        collection = get_collection(collection_name)

        # Select appropriate chunk generator
        if file_type in [".xlsx", ".xls"]:
            chunk_gen = iter_excel_chunks(content, document_name, report_date)
        elif file_type == ".pdf":
            chunk_gen = iter_pdf_chunks(content, document_name, report_date)
        elif file_type == ".msg":
            chunk_gen = iter_email_chunks(content, document_name, report_date)
        elif file_type == ".docx":
                chunk_gen = iter_docx_chunks(content, document_name, report_date)

        else:
            return IngestionResult(
                success=False,
                document_name=document_name,
                counts={},
                error=f"Unsupported file type: {file_type}"
            )

        # Batch processing
        texts: List[str] = []
        metas: List[Dict] = []
        ids: List[str] = []
        seen_ids: Set[str] = set()
        counts = {"table": 0, "paragraph": 0, "email_body": 0}

        for text, meta in chunk_gen:
            if not text or not text.strip():
                continue

            chunk_type = meta.get("chunk_type", "unknown")
            chunk_id = make_chunk_id(document_name, chunk_type, text[:200])

            if chunk_id in seen_ids:
                continue
            seen_ids.add(chunk_id)

            texts.append(text)
            metas.append(meta)
            ids.append(chunk_id)
            counts[chunk_type] = counts.get(chunk_type, 0) + 1

            if len(texts) >= batch_size:
                logger.info(f"Storing batch of {len(texts)} chunks...")
                embeddings = get_embeddings(texts)
                collection.upsert(ids=ids, documents=texts, metadatas=metas, embeddings=embeddings)
                texts, metas, ids = [], [], []

        # Final batch
        if texts:
            logger.info(f"Storing final batch of {len(texts)} chunks...")
            embeddings = get_embeddings(texts)
            collection.upsert(ids=ids, documents=texts, metadatas=metas, embeddings=embeddings)

        # Remove zero counts
        counts = {k: v for k, v in counts.items() if v > 0}

        logger.info(f"Ingestion complete: {counts}")
        return IngestionResult(
            success=True,
            document_name=document_name,
            counts=counts
        )

    except Exception as e:
        logger.error(f"Ingestion failed: {e}", exc_info=True)
        return IngestionResult(
            success=False,
            document_name=document_name,
            counts={},
            error=str(e)
        )
