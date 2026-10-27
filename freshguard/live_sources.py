"""Live knowledge sources (FR7 live retrieval) — Tavily web search and xAI Grok.

When a query's answer is **not in the local knowledge base**, FreshGuard fetches fresh
evidence from the web (Tavily) and/or from Grok, ingests it into the vector store and
knowledge base, then retrieves and verifies it exactly like any seeded document.

Both sources are optional and keyed by environment variables. If a key (or the network)
is missing the source simply returns nothing, so the system stays fully offline and the
test-suite remains deterministic:

  * Tavily web search  -> TAVILY_API_KEY
  * xAI Grok           -> XAI_API_KEY  (or GROK_API_KEY)

All network access goes through `requests` with short timeouts and is wrapped in
try/except; a failed call degrades gracefully to "no documents".
"""
from __future__ import annotations

import os
import re
from datetime import date

import requests

TAVILY_URL = "https://api.tavily.com/search"
GROK_URL = "https://api.x.ai/v1/chat/completions"
_TIMEOUT = 20  # seconds


# -- key / availability --------------------------------------------------------
def tavily_key() -> str | None:
    return os.environ.get("TAVILY_API_KEY")


def grok_key() -> str | None:
    return os.environ.get("XAI_API_KEY") or os.environ.get("GROK_API_KEY")


def tavily_available() -> bool:
    return bool(tavily_key())


def grok_available() -> bool:
    return bool(grok_key())


# -- helpers -------------------------------------------------------------------
def _slug(text: str, n: int = 40) -> str:
    s = re.sub(r"[^a-z0-9]+", "-", text.lower()).strip("-")
    return s[:n] or "query"


def _today() -> str:
    return date.today().isoformat()


# -- Tavily web search (FR7) ---------------------------------------------------
def tavily_search(query: str, max_results: int = 3) -> list[dict]:
    """Query the Tavily web-search API and return KB-shaped documents.

    Each web result becomes one document; Tavily's synthesised ``answer`` (when present)
    becomes an extra, high-signal document. Returns ``[]`` on any failure."""
    key = tavily_key()
    if not key:
        return []
    try:
        resp = requests.post(
            TAVILY_URL,
            json={
                "api_key": key,
                "query": query,
                "max_results": max(1, int(max_results)),
                "search_depth": "basic",
                "include_answer": True,
            },
            timeout=_TIMEOUT,
        )
        resp.raise_for_status()
        data = resp.json()
    except Exception:
        return []

    slug = _slug(query)
    docs: list[dict] = []

    answer = (data.get("answer") or "").strip()
    if answer:
        docs.append({
            "doc_id": f"tavily-{slug}-answer",
            "title": f"Web answer: {query}".strip()[:120],
            "url": "https://tavily.com/",
            "date": _today(),
            "text": answer,
        })

    for i, res in enumerate(data.get("results") or []):
        text = (res.get("content") or "").strip()
        if not text:
            continue
        docs.append({
            "doc_id": f"tavily-{slug}-{i}",
            "title": (res.get("title") or f"Web result {i + 1}")[:120],
            "url": res.get("url") or "https://tavily.com/",
            "date": (res.get("published_date") or _today())[:10],
            "text": text,
        })
    return docs


# -- xAI Grok (FR7 / FR9 live knowledge) ---------------------------------------
_GROK_SYSTEM = (
    "You are a precise factual knowledge source. Given a question, reply with a short "
    "factual passage of 2-4 sentences containing concrete facts, names, numbers and dates. "
    "Do not add commentary, do not refer to yourself, and do not hedge. If the event is in "
    "the future or has not happened yet, state that explicitly and give only what is "
    "currently known or scheduled."
)


def grok_knowledge(query: str, model: str = "grok-2-latest") -> list[dict]:
    """Ask Grok (xAI, OpenAI-compatible endpoint) for a factual passage about ``query``
    and return it as a single KB document. Returns ``[]`` on any failure."""
    key = grok_key()
    if not key:
        return []
    headers = {"Authorization": f"Bearer {key}", "Content-Type": "application/json"}
    base = {
        "model": model,
        "temperature": 0,
        "messages": [
            {"role": "system", "content": _GROK_SYSTEM},
            {"role": "user", "content": query},
        ],
    }
    # First try with xAI Live Search enabled; fall back to a plain completion if the
    # account/endpoint does not support that parameter.
    for body in ({**base, "search_parameters": {"mode": "auto"}}, base):
        try:
            resp = requests.post(GROK_URL, json=body, headers=headers, timeout=_TIMEOUT)
            if resp.status_code >= 400:
                continue
            data = resp.json()
            text = (data["choices"][0]["message"]["content"] or "").strip()
        except Exception:
            continue
        if text:
            return [{
                "doc_id": f"grok-{_slug(query)}",
                "title": f"Grok: {query}".strip()[:120],
                "url": "https://x.ai/grok",
                "date": _today(),
                "text": text,
            }]
    return []


# -- orchestration -------------------------------------------------------------
def gather_live_docs(query: str, settings) -> tuple[list[dict], list[str]]:
    """Collect live documents from every enabled, available source.

    Returns ``(docs, sources_used)`` where ``sources_used`` names the sources that
    actually returned content (e.g. ``["tavily", "grok"]``)."""
    docs: list[dict] = []
    used: list[str] = []
    max_docs = getattr(settings, "max_live_docs", 3)

    if getattr(settings, "web_search_provider", "tavily") == "tavily":
        tv = tavily_search(query, max_docs)
        if tv:
            docs.extend(tv)
            used.append("tavily")

    if getattr(settings, "grok_ingest_enabled", True):
        gk = grok_knowledge(query, getattr(settings, "grok_model", "grok-2-latest"))
        if gk:
            docs.extend(gk)
            used.append("grok")

    return docs, used
