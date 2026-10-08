"""
News scraper for CMD Dashboard news section.

Scrapes free public sources for energy sector news published in the report month,
then uses LLM to select and summarize items in CMD style.
"""

import json
import logging
import re
from datetime import datetime
from typing import Any, Optional

import requests
from bs4 import BeautifulSoup

logger = logging.getLogger(__name__)

_HEADERS = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)"}

_RSS_SOURCES = [
    {
        "name": "Google News – India Natural Gas",
        "url": "https://news.google.com/rss/search?q=india+natural+gas+GAIL&hl=en-IN&gl=IN&ceid=IN:en",
    },
    {
        "name": "Google News – India Oil Gas Policy",
        "url": "https://news.google.com/rss/search?q=india+oil+gas+petroleum+ministry+MoPNG&hl=en-IN&gl=IN&ceid=IN:en",
    },
    {
        "name": "Google News – India Energy Sector",
        "url": "https://news.google.com/rss/search?q=india+energy+sector+LNG+CNG+CGD&hl=en-IN&gl=IN&ceid=IN:en",
    },
    {
        "name": "Google News – India Renewable Energy",
        "url": "https://news.google.com/rss/search?q=india+renewable+energy+MNRE+solar+wind&hl=en-IN&gl=IN&ceid=IN:en",
    },
]


def _parse_entry_date(entry) -> Optional[datetime]:
    """Extract published datetime from a feedparser entry."""
    for attr in ("published_parsed", "updated_parsed"):
        val = getattr(entry, attr, None)
        if val:
            try:
                return datetime(*val[:6])
            except Exception:
                pass
    return None


def _fetch_rss(url: str, source_name: str, report_month: str) -> list:
    """
    Fetch an RSS feed and return items published in report_month (YYYY-MM).
    Each item: {title, summary, source, date}
    """
    try:
        import feedparser
    except ImportError:
        logger.warning("feedparser not installed — skipping RSS source: %s", source_name)
        return []

    try:
        resp = requests.get(url, headers=_HEADERS, timeout=10)
        resp.raise_for_status()
        feed = feedparser.parse(resp.content)
        year, month = int(report_month[:4]), int(report_month[5:7])
        items = []
        for entry in feed.entries:
            pub = _parse_entry_date(entry)
            if pub is None or pub.year != year or pub.month != month:
                continue
            title = getattr(entry, "title", "").strip()
            raw_summary = getattr(entry, "summary", "") or getattr(entry, "description", "")
            summary = BeautifulSoup(raw_summary, "html.parser").get_text(" ", strip=True) if raw_summary else ""
            items.append({
                "title": title,
                "summary": summary[:500],
                "source": source_name,
                "date": pub.strftime("%Y-%m-%d"),
            })
        logger.info("RSS %s: %d items for %s", source_name, len(items), report_month)
        return items
    except Exception as e:
        logger.warning("RSS fetch failed for %s (%s): %s", source_name, url, e)
        return []


def _fetch_pngrb(report_month: str) -> list:
    """
    Scrape PNGRB announcements page for items in report_month.
    Each item: {title, summary, source, date}
    """
    url = "https://www.pngrb.gov.in/English/Announcements.aspx"
    try:
        r = requests.get(url, headers=_HEADERS, timeout=15)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        year, month = int(report_month[:4]), int(report_month[5:7])
        items = []
        for row in soup.find_all(["tr", "li"]):
            text = row.get_text(" ", strip=True)
            date_match = re.search(r"(\d{2})/(\d{2})/(\d{4})", text)
            if not date_match:
                continue
            d, m, y = int(date_match.group(1)), int(date_match.group(2)), int(date_match.group(3))
            if y != year or m != month:
                continue
            link = row.find("a")
            title = link.get_text(strip=True) if link else text[:120]
            if title:
                items.append({
                    "title": title,
                    "summary": "",
                    "source": "PNGRB",
                    "date": f"{y}-{m:02d}-{d:02d}",
                })
        logger.info("PNGRB: %d items for %s", len(items), report_month)
        return items
    except Exception as e:
        logger.warning("PNGRB scrape failed: %s", e)
        return []


def scrape_news_for_month(report_month: str) -> list:
    """
    Scrape all configured RSS feeds and PNGRB for articles published in report_month.

    Args:
        report_month: YYYY-MM format (e.g. "2025-09")

    Returns:
        List of dicts: {title, summary, source, date}
    """
    all_items: list = []

    for src in _RSS_SOURCES:
        all_items.extend(_fetch_rss(src["url"], src["name"], report_month))

    all_items.extend(_fetch_pngrb(report_month))

    logger.info("Total scraped items for %s: %d", report_month, len(all_items))
    return all_items


def summarize_news(
    raw_items: list,
    report_month: str,
    llm_client: Any,
    llm_model: str,
) -> list:
    """
    Use LLM to select and summarize the most relevant news items in CMD style.

    Args:
        raw_items: Scraped items from scrape_news_for_month()
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

    system_prompt = (
        "You are an energy sector analyst summarizing news for GAIL's CMD (Committee Meeting Dashboard). "
        "Select the most relevant items related to the Indian energy sector, natural gas, GAIL operations, "
        "petroleum, renewable energy, government policy, or macroeconomic indicators. "
        "Summarize each selected item to 1–3 concise sentences in a professional, neutral tone."
    )

    user_prompt = (
        f"Report month: {report_month}\n\n"
        f"Below are {len(raw_items)} news items scraped from public sources.\n"
        "Select 10–12 most relevant items and return them as a JSON array.\n"
        "Each object must have:\n"
        '  "text": 1–3 sentence summary\n'
        '  "source": source name\n\n'
        "Return ONLY the JSON array, no other text.\n\n"
        f"{numbered_list}"
    )

    try:
        response = llm_client.chat.completions.create(
            model=llm_model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            temperature=0,
            max_tokens=2000,
        )
        raw = response.choices[0].message.content.strip()

        # Strip markdown code fences if present
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)

        items = json.loads(raw)
        if not isinstance(items, list):
            raise ValueError("LLM did not return a JSON array")

        logger.info("LLM selected %d news items for %s", len(items), report_month)
        return items

    except Exception as e:
        logger.error("News summarization failed: %s", e)
        return []
