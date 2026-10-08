"""
news_scraper_v2.py — improved news scraper for CMD Dashboard.

Changes from v1:
  - More reliable RSS sources (PIB, Economic Times, Business Standard, MoPNG)
  - Retry with exponential backoff on all fetches
  - PNGRB removed (fragile HTML scraping replaced by RSS alternatives)
  - Two-pass deduplication (exact + near-duplicate via difflib)
  - Hardened LLM JSON parsing with retries and item-level fallback
  - Standalone runnable: python news_scraper_v2.py [YYYY-MM]

Changes from v2.0:
  - Added sources: PSU Watch, Renewable Watch, ORF, CII, Indian Petrochem, Hydrogen Insight
  - Added PIB feeds for MNRE and Ministry of Power
  - Added topic-specific Google News queries for all CMD topic areas:
    CBG/CGD, M&A deals, LNG, Green Hydrogen, Petrochemicals, Crude, Macro, Commodities
  - Updated LLM prompt to explicitly cover all 9 CMD news topic categories

Same public API as news_scraper.py:
  scrape_news_for_month(report_month) -> list
  summarize_news(raw_items, report_month, llm_client, llm_model) -> list
"""

import json
import logging
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from difflib import SequenceMatcher
from typing import Any, Optional

import requests
from bs4 import BeautifulSoup

logger = logging.getLogger(__name__)

_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (X11; Linux x86_64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}

_RETRY_ATTEMPTS = 3
_RETRY_BACKOFF = 2  # seconds, doubled each retry

_RSS_SOURCES = [
    # --- Google News: Natural Gas / GAIL / CGD / CBG ---
    {
        "name": "Google News – GAIL Natural Gas",
        "url": "https://news.google.com/rss/search?q=GAIL+natural+gas+india&hl=en-IN&gl=IN&ceid=IN:en",
    },
    {
        "name": "Google News – CGD CBG City Gas",
        "url": "https://news.google.com/rss/search?q=india+CGD+CBG+city+gas+distribution+compressed+biogas&hl=en-IN&gl=IN&ceid=IN:en",
    },
    # --- Google News: LNG ---
    {
        "name": "Google News – LNG India",
        "url": "https://news.google.com/rss/search?q=india+LNG+liquefied+natural+gas+sourcing+shipping+terminal&hl=en-IN&gl=IN&ceid=IN:en",
    },
    # --- Google News: M&A and New Projects ---
    {
        "name": "Google News – Energy MA Deals India",
        "url": "https://news.google.com/rss/search?q=india+energy+oil+gas+merger+acquisition+deal+MA&hl=en-IN&gl=IN&ceid=IN:en",
    },
    {
        "name": "Google News – Energy Infrastructure Projects India",
        "url": "https://news.google.com/rss/search?q=india+energy+infrastructure+project+approval+investment+pipeline&hl=en-IN&gl=IN&ceid=IN:en",
    },
    # --- Google News: Renewables and Green Hydrogen ---
    {
        "name": "Google News – India Renewables MNRE",
        "url": "https://news.google.com/rss/search?q=india+renewable+energy+solar+wind+MNRE&hl=en-IN&gl=IN&ceid=IN:en",
    },
    {
        "name": "Google News – Green Hydrogen India",
        "url": "https://news.google.com/rss/search?q=india+green+hydrogen+H2+electrolyzer+national+hydrogen+mission&hl=en-IN&gl=IN&ceid=IN:en",
    },
    # --- Google News: Petrochemicals ---
    {
        "name": "Google News – Petrochemicals India",
        "url": "https://news.google.com/rss/search?q=india+petrochemicals+polymer+petchem+refinery&hl=en-IN&gl=IN&ceid=IN:en",
    },
    # --- Google News: Crude Oil ---
    {
        "name": "Google News – Crude Oil India",
        "url": "https://news.google.com/rss/search?q=crude+oil+price+india+brent+opec+WTI&hl=en-IN&gl=IN&ceid=IN:en",
    },
    # --- Google News: Global Macroeconomy ---
    {
        "name": "Google News – India Macroeconomy",
        "url": "https://news.google.com/rss/search?q=india+GDP+inflation+RBI+fiscal+monetary+policy+macroeconomy&hl=en-IN&gl=IN&ceid=IN:en",
    },
    # --- Google News: Commodity Prices ---
    {
        "name": "Google News – Commodity Prices",
        "url": "https://news.google.com/rss/search?q=commodity+prices+coal+steel+fertilizer+india+global&hl=en-IN&gl=IN&ceid=IN:en",
    },
    # --- Google News: MoPNG / Oil Gas Policy ---
    {
        "name": "Google News – MoPNG Oil Gas Policy",
        "url": "https://news.google.com/rss/search?q=india+petroleum+ministry+MoPNG+oil+gas+policy&hl=en-IN&gl=IN&ceid=IN:en",
    },
    # --- Google News: PSU / Industry bodies (CII, ORF) ---
    {
        "name": "Google News – PSU India Energy",
        "url": "https://news.google.com/rss/search?q=PSU+india+energy+gas+petroleum+ONGC+IOC+BPCL&hl=en-IN&gl=IN&ceid=IN:en",
    },
    {
        "name": "Google News – CII India Energy Policy",
        "url": "https://news.google.com/rss/search?q=CII+confederation+indian+industry+energy+policy&hl=en-IN&gl=IN&ceid=IN:en",
    },
    {
        "name": "Google News – ORF Energy Research",
        "url": "https://news.google.com/rss/search?q=observer+research+foundation+ORF+energy+india&hl=en-IN&gl=IN&ceid=IN:en",
    },
    {
        "name": "Google News – Indian Petrochem",
        "url": "https://news.google.com/rss/search?q=indian+petrochem+petrochemical+industry+india&hl=en-IN&gl=IN&ceid=IN:en",
    },
    {
        "name": "Google News – Hydrogen Insight",
        "url": "https://news.google.com/rss/search?q=hydrogen+insight+green+hydrogen+india&hl=en-IN&gl=IN&ceid=IN:en",
    },
    # --- PIB: Ministry-specific feeds (Government of India) ---
    {
        "name": "PIB – Petroleum & Natural Gas (MoPNG)",
        "url": "https://www.pib.gov.in/RssMain.aspx?ModId=6&Lang=1&Regid=3",
    },
    {
        "name": "PIB – New & Renewable Energy (MNRE)",
        "url": "https://www.pib.gov.in/RssMain.aspx?ModId=36&Lang=1&Regid=3",
    },
    {
        "name": "PIB – Power (MoP)",
        "url": "https://www.pib.gov.in/RssMain.aspx?ModId=7&Lang=1&Regid=3",
    },
    # --- Sector-specific publications ---
    {
        "name": "Economic Times – Energy",
        "url": "https://economictimes.indiatimes.com/industry/energy/rssfeeds/13358575.cms",
    },
    {
        "name": "PSU Watch",
        "url": "https://psuwatch.com/feed/",
    },
    {
        "name": "Renewable Watch",
        "url": "https://renewablewatch.in/feed/",
    },
    {
        "name": "Hindu BusinessLine – Economy",
        "url": "https://www.thehindubusinessline.com/economy/feeder/default.rss",
    },
    {
        "name": "Financial Express – Economy",
        "url": "https://www.financialexpress.com/economy/feed/",
    },
]


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _parse_entry_date(entry) -> Optional[datetime]:
    for attr in ("published_parsed", "updated_parsed"):
        val = getattr(entry, attr, None)
        if val:
            try:
                return datetime(*val[:6])
            except Exception:
                pass
    return None


def _fetch_with_retry(url: str, timeout: int = 12) -> Optional[requests.Response]:
    """GET with up to _RETRY_ATTEMPTS retries and exponential backoff."""
    delay = _RETRY_BACKOFF
    for attempt in range(1, _RETRY_ATTEMPTS + 1):
        try:
            resp = requests.get(url, headers=_HEADERS, timeout=timeout)
            resp.raise_for_status()
            return resp
        except Exception as e:
            if attempt == _RETRY_ATTEMPTS:
                logger.warning("All %d attempts failed for %s: %s", _RETRY_ATTEMPTS, url, e)
                return None
            logger.debug("Attempt %d failed for %s (%s), retrying in %ds", attempt, url, e, delay)
            time.sleep(delay)
            delay *= 2
    return None


def _fetch_rss(url: str, source_name: str, report_month: str) -> list:
    """Fetch an RSS feed and return items published in report_month (YYYY-MM)."""
    try:
        import feedparser
    except ImportError:
        logger.warning("feedparser not installed — skipping: %s", source_name)
        return []

    resp = _fetch_with_retry(url)
    if resp is None:
        return []

    try:
        feed = feedparser.parse(resp.content)
        year, month = int(report_month[:4]), int(report_month[5:7])
        items = []
        for entry in feed.entries:
            pub = _parse_entry_date(entry)
            if pub is None or pub.year != year or pub.month != month:
                continue
            title = getattr(entry, "title", "").strip()
            if not title:
                continue
            raw_summary = getattr(entry, "summary", "") or getattr(entry, "description", "")
            summary = (
                BeautifulSoup(raw_summary, "html.parser").get_text(" ", strip=True)
                if raw_summary else ""
            )
            items.append({
                "title": title,
                "summary": summary[:500],
                "source": source_name,
                "date": pub.strftime("%Y-%m-%d"),
            })
        logger.info("RSS %s: %d items for %s", source_name, len(items), report_month)
        return items
    except Exception as e:
        logger.warning("RSS parse failed for %s: %s", source_name, e)
        return []


# ---------------------------------------------------------------------------
# Deduplication
# ---------------------------------------------------------------------------

def _normalize_title(title: str) -> str:
    """Lowercase, strip punctuation and extra whitespace for comparison."""
    t = title.lower()
    t = re.sub(r"[^\w\s]", "", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t


def _deduplicate(items: list, similarity_threshold: float = 0.82, field: str = "title") -> list:
    """
    Two-pass dedup:
      1. Exact match on normalized field value (hash-based, O(n))
      2. Near-duplicate via SequenceMatcher (O(n^2) but n is small ~100)
    """
    # Pass 1 — exact
    seen_exact: set = set()
    pass1: list = []
    for item in items:
        key = _normalize_title(item.get(field, ""))
        if key and key not in seen_exact:
            seen_exact.add(key)
            pass1.append(item)

    # Pass 2 — near-duplicate
    unique: list = []
    normalized_kept: list = []
    for item in pass1:
        norm = _normalize_title(item.get(field, ""))
        is_dup = any(
            SequenceMatcher(None, norm, kept).ratio() >= similarity_threshold
            for kept in normalized_kept
        )
        if not is_dup:
            unique.append(item)
            normalized_kept.append(norm)

    removed = len(items) - len(unique)
    if removed:
        logger.info("Deduplication removed %d duplicate items (%d → %d)", removed, len(items), len(unique))
    return unique


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def scrape_news_for_month(report_month: str) -> list:
    """
    Scrape all configured RSS feeds concurrently for articles published in
    report_month, then deduplicate.

    Args:
        report_month: YYYY-MM format (e.g. "2025-09")

    Returns:
        List of dicts: {title, summary, source, date}
    """
    all_items: list = []
    with ThreadPoolExecutor(max_workers=len(_RSS_SOURCES)) as executor:
        futures = {
            executor.submit(_fetch_rss, src["url"], src["name"], report_month): src["name"]
            for src in _RSS_SOURCES
        }
        for future in as_completed(futures):
            all_items.extend(future.result())

    unique_items = _deduplicate(all_items)
    logger.info("Total unique items for %s: %d (from %d raw)", report_month, len(unique_items), len(all_items))
    return unique_items


def summarize_news(
    raw_items: list,
    report_month: str,
    llm_client: Any,
    llm_model: str,
) -> list:
    """
    Use LLM to select and summarize the most relevant news items in CMD style.
    Retries up to 3 times on JSON parse failure. Falls back to raw titles if
    all retries fail so the news section is never completely empty.

    Args:
        raw_items: From scrape_news_for_month()
        report_month: YYYY-MM format
        llm_client: OpenAI or AzureOpenAI client
        llm_model: Model/deployment name

    Returns:
        List of dicts: {text: str, source: str}
    """
    if not raw_items:
        logger.warning("No raw news items to summarize")
        return []

    lines = []
    for i, item in enumerate(raw_items, 1):
        title = item.get("title", "")
        summary = item.get("summary", "")
        source = item.get("source", "")
        date = item.get("date", "")
        text = f"{title}. {summary}".strip(". ")
        lines.append(f"{i}. [{source}, {date}] {text}")

    numbered_list = "\n".join(lines)

    target_count = min(18, max(9, len(raw_items)))

    system_prompt = (
        "You are an energy sector analyst summarizing news for GAIL's CMD (Committee Meeting Dashboard). "
        "Prioritize items that fall into these topic categories:\n"
        "  1. Natural Gas (including CBG/CGD, city gas distribution, compressed biogas)\n"
        "  2. Deals – M&A (mergers, acquisitions, JVs in the energy sector)\n"
        "  3. New Projects (approvals, infrastructure investment in energy)\n"
        "  4. LNG (sourcing, shipping, regasification, consumption)\n"
        "  5. Renewables including H2 (solar, wind, green hydrogen, MNRE)\n"
        "  6. Petrochemicals (polymers, refining, petchem industry)\n"
        "  7. Crude (crude oil prices, OPEC, upstream exploration)\n"
        "  8. Global Macroeconomy (India GDP, inflation, RBI, fiscal policy)\n"
        "  9. Commodity Prices (coal, steel, fertilizers, energy commodities)\n"
        "Also include significant GAIL-specific news and government policy (MoPNG, MoP, MNRE, PIB). "
        "Summarize each selected item in 3–4 sentences (a full short paragraph) covering what happened, "
        "key figures/entities involved, and why it matters for GAIL/the energy sector, in a professional, "
        "neutral tone. "
        "You MUST respond with a valid JSON array only — no prose, no markdown fences, no explanation. "
        'Example: [{"text": "India approved... This is significant because... The move is expected to...", "source": "PIB", "topic": "New Projects"}]'
    )

    user_prompt = (
        f"Report month: {report_month}\n\n"
        f"Below are {len(raw_items)} news items. Select up to {target_count} most relevant, "
        "ensuring broad coverage across the 9 topic categories. "
        "Each selected item must cover a DIFFERENT underlying story — do not select "
        "the same news item twice, even with reworded text.\n"
        "Return a JSON array where each object has exactly:\n"
        '  "text": 3–4 sentence summary (a full short paragraph, ~3-4 lines)\n'
        '  "source": source name\n'
        '  "topic": one of the 9 categories above\n\n'
        "Return ONLY the JSON array, nothing else.\n\n"
        f"{numbered_list}"
    )

    def _parse_llm_response(raw: str) -> list:
        """Strip fences, parse JSON, validate items."""
        raw = re.sub(r"^```(?:json)?\s*", "", raw.strip())
        raw = re.sub(r"\s*```$", "", raw)
        # Handle object wrapper e.g. {"items": [...]}
        parsed = json.loads(raw)
        if isinstance(parsed, dict):
            for val in parsed.values():
                if isinstance(val, list):
                    parsed = val
                    break
        if not isinstance(parsed, list):
            raise ValueError("LLM did not return a JSON array")
        # Item-level validation — keep only valid items
        valid = [
            item for item in parsed
            if isinstance(item, dict)
            and isinstance(item.get("text"), str)
            and item["text"].strip()
        ]
        if not valid:
            raise ValueError("No valid items in LLM response")
        return valid

    last_error: Exception = Exception("unknown")
    for attempt in range(1, _RETRY_ATTEMPTS + 1):
        try:
            response = llm_client.chat.completions.create(
                model=llm_model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
                temperature=0,
                max_tokens=4096,
            )
            raw = response.choices[0].message.content.strip()
            items = _parse_llm_response(raw)
            items = _deduplicate(items, similarity_threshold=0.85, field="text")
            logger.info("LLM selected %d unique news items (attempt %d)", len(items), attempt)
            return items
        except Exception as e:
            last_error = e
            logger.warning("LLM summarization attempt %d/%d failed: %s", attempt, _RETRY_ATTEMPTS, e)
            if attempt < _RETRY_ATTEMPTS:
                time.sleep(_RETRY_BACKOFF * attempt)

    # Fallback — use raw scraped titles directly so section is never empty
    logger.error(
        "All LLM summarization attempts failed (%s). Falling back to raw titles.", last_error
    )
    fallback = [
        {"text": item.get("title", "").strip(), "source": item.get("source", "")}
        for item in raw_items[:12]
        if item.get("title", "").strip()
    ]
    logger.info("Fallback: returning %d raw titles as news items", len(fallback))
    return fallback


# ---------------------------------------------------------------------------
# Standalone test runner
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys
    import os
    from pathlib import Path

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )

    # Load .env from cmd/ directory
    env_path = Path(__file__).parent.parent / ".env"
    if env_path.exists():
        from dotenv import load_dotenv
        load_dotenv(env_path)
        print(f"Loaded .env from {env_path}")
    else:
        print(f"Warning: .env not found at {env_path}")

    report_month = sys.argv[1] if len(sys.argv) > 1 else datetime.now().strftime("%Y-%m")
    print(f"\n=== Scraping news for: {report_month} ===\n")

    # Step 1 — scrape
    items = scrape_news_for_month(report_month)
    print(f"\nScraped {len(items)} unique items:")
    for i, item in enumerate(items, 1):
        print(f"  {i:2}. [{item['date']}] [{item['source']}] {item['title'][:80]}")

    if not items:
        print("\nNo items scraped — check sources or try a different month.")
        sys.exit(0)

    # Step 2 — summarize via LLM
    llm_provider = os.getenv("LLM_PROVIDER", "azure")
    print(f"\n=== Summarizing via LLM (provider: {llm_provider}) ===\n")

    if llm_provider == "openai":
        from openai import OpenAI
        client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
        model = os.getenv("OPENAI_LLM_MODEL", "gpt-4o")
    else:
        from openai import AzureOpenAI
        client = AzureOpenAI(
            azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT", ""),
            api_key=os.getenv("AZURE_OPENAI_API_KEY", ""),
            api_version=os.getenv("AZURE_OPENAI_API_VERSION", "2024-02-01"),
        )
        model = os.getenv("AZURE_OPENAI_LLM_DEPLOYMENT", "gpt4o")

    summarized = summarize_news(items, report_month, client, model)

    print(f"\nFinal news items ({len(summarized)}):\n")
    for i, item in enumerate(summarized, 1):
        print(f"  {i:2}. {item.get('text', '')}")
        print(f"      Source: {item.get('source', '')}\n")
