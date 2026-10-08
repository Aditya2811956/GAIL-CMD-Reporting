"""
Shared configuration for CMD Report Automation Pipeline.

All environment variables and constants are defined here.
"""

import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()


class AzureConfig:
    """Azure service configuration."""
    # Document Intelligence
    DI_ENDPOINT: str = os.getenv("AZURE_DI_ENDPOINT", "https://ikegaidocint.cognitiveservices.azure.com/")
    DI_KEY: str = os.getenv("DI_KEY", "")

    # Azure OpenAI
    OPENAI_ENDPOINT: str = os.getenv("AZURE_OPENAI_ENDPOINT", "https://openai-lh.openai.azure.com")
    OPENAI_API_KEY: str = os.getenv("AZURE_OPENAI_API_KEY", "")
    OPENAI_API_VERSION: str = os.getenv("AZURE_OPENAI_API_VERSION", "2024-02-01")
    EMBED_DEPLOYMENT: str = os.getenv("AZURE_OPENAI_EMBED_DEPLOYMENT", "LH-embedding")
    LLM_DEPLOYMENT: str = os.getenv("AZURE_OPENAI_LLM_DEPLOYMENT", "gpt-4o")


class OpenAIConfig:
    """OpenAI (non-Azure) configuration - fallback option."""
    API_KEY: str = os.getenv("OPENAI_API_KEY", "")
    EMBED_MODEL: str = os.getenv("OPENAI_EMBED_MODEL", "text-embedding-3-small")
    LLM_MODEL: str = os.getenv("OPENAI_LLM_MODEL", "gpt-4o")


class LLMConfig:
    """LLM provider selection."""
    # Set to "openai" or "azure"
    PROVIDER: str = os.getenv("LLM_PROVIDER", "azure")
    # Embedding provider — defaults to PROVIDER but can be set independently.
    # Important: must match the provider used when documents were ingested into ChromaDB.
    EMBED_PROVIDER: str = os.getenv("EMBED_PROVIDER", os.getenv("LLM_PROVIDER", "azure"))


class ChromaConfig:
    """ChromaDB configuration."""
    HOST: str = os.getenv("CHROMA_HOST", "127.0.0.1")
    PORT: int = int(os.getenv("CHROMA_PORT", "8001"))
    COLLECTION: str = os.getenv("CHROMA_COLLECTION", "cmd_reports_v2")


class PathConfig:
    """File path configuration."""
    DOCUMENTS_DIR: str = os.getenv("DOCUMENTS_DIR", str(Path(__file__).parent / "documents"))
    TEMPLATES_DIR: str = os.getenv("TEMPLATES_DIR", str(Path(__file__).parent / "templates"))
    OUTPUT_DIR: str = os.getenv("OUTPUT_DIR", "output")


# Metric type keywords for tagging during ingestion
METRIC_KEYWORDS = {
    "gdp": ["gdp", "gross domestic product", "gva", "gross value added"],
    "inflation_wpi": ["wpi", "wholesale price index", "all commodities"],
    "inflation_cpi": ["cpi", "consumer price index", "cpi general"],
    "iip": ["iip", "industrial production", "index of industrial"],
    "gas_production": ["net production", "mmscm", "mmscmd", "natural gas consumption"],
    "gas_lng": ["lng import", "lng consumption"],
    "interest_repo": ["repo rate", "policy repo", "rbi repo"],
    "interest_mclr": ["mclr", "marginal cost"],
    "gdp_forecast": ["forecast", "projection", "rbi bulletin", "world bank", "imf"],
    "satat": ["satat", "cbg", "compressed biogas", "eoi", "loi"],
    "cgd": ["cgd", "city gas", "igl", "mngl", "cugl", "dpng", "cng stations"],
    "retail_lng": ["retail lng", "lcng", "lng station", "lng dispensing"],
}


def detect_metric_types(text: str) -> list[str]:
    """Detect potential metric types from text content."""
    text_lower = text.lower()
    return [k for k, keywords in METRIC_KEYWORDS.items()
            if any(kw in text_lower for kw in keywords)]
