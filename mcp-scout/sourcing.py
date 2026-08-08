# © 2026 Iryna Subbotina. All Rights Reserved. Proprietary — not open source. See LICENSE/NOTICE. Unauthorized copying, redistribution, or claim of authorship is prohibited.
"""LinkedIn X-Ray candidate sourcing via DuckDuckGo."""
from __future__ import annotations

import logging
import re

log = logging.getLogger(__name__)

_LOCATION_RE = re.compile(r"[-–|·•]\s*([^·•\-–]+(?:,\s*[^·•\-–]+)*)\s*$")
_STOPWORDS = {"of", "a", "an", "the", "and", "or", "for", "with", "in", "at", "to"}


def _build_queries(position: str, location: str, keywords: list[str] | None = None) -> list[str]:
    base = f'site:linkedin.com/in "{position}" "{location}"'
    if not keywords:
        return [base]
    quoted = [f'"{kw}"' for kw in keywords[:6]]
    # Most specific first: every keyword required in the same query.
    combined = f'{base} {" ".join(quoted)}'
    # Fallbacks in case the fully-combined query is too narrow to return anything —
    # still validated by _matches_required below, so they can't loosen precision.
    per_keyword = [f'{base} {kw}' for kw in quoted[:3]]
    return [combined] + per_keyword + [base]


def _ddg_search(query: str, max_results: int = 10) -> list[dict]:
    try:
        from ddgs import DDGS
    except ImportError:
        from duckduckgo_search import DDGS
    with DDGS() as ddgs:
        return list(ddgs.text(query, max_results=max_results))


def _parse_item(r: dict, position: str, location: str) -> dict | None:
    url: str = r.get("href", "") or r.get("link", "") or r.get("url", "")
    if "linkedin.com/in/" not in url:
        return None

    title: str = r.get("title", "")
    snippet: str = r.get("body", "") or r.get("snippet", "")

    parts = [p.strip() for p in re.split(r"\s[-–|]\s", title)]
    full_name = re.sub(r"\s*\|\s*LinkedIn$", "", parts[0], flags=re.IGNORECASE).strip()
    headline = parts[1] if len(parts) > 1 else ""
    company = re.sub(r"\s*\|\s*LinkedIn$", "", parts[2], flags=re.IGNORECASE).strip() if len(parts) > 2 else ""

    loc_match = _LOCATION_RE.search(snippet)
    inferred_location = loc_match.group(1).strip() if loc_match else location

    return {
        "full_name": full_name,
        "url": url,
        "headline": headline,
        "company": company,
        "location": inferred_location,
        "snippet": snippet[:300],
        "_searchable": f"{title} {headline} {company} {snippet}".lower(),
    }


def _tokens(text: str) -> list[str]:
    words = re.findall(r"[a-zA-ZЀ-ӿ؀-ۿ]+", text.lower())
    return [w for w in words if len(w) > 2 and w not in _STOPWORDS]


def _matches_required(candidate: dict, position: str, location: str, keywords: list[str] | None) -> bool:
    haystack = candidate["_searchable"]

    for tok in _tokens(position):
        if tok not in haystack:
            return False
    for part in location.split(","):
        part = part.strip().lower()
        if part and part not in haystack:
            return False
    for kw in (keywords or []):
        if kw.strip().lower() not in haystack:
            return False
    return True


def find_linkedin_candidates(
    position: str,
    location: str,
    keywords: list[str] | None = None,
    max_results: int = 20,
) -> list[dict]:
    """Search LinkedIn profiles via DuckDuckGo X-Ray.

    Only returns candidates whose title/snippet actually mentions every required
    term (position, location, and all keywords) — narrow searches may return few
    or no results rather than loosely-related profiles.
    """
    queries = _build_queries(position, location, keywords)
    seen: set[str] = set()
    candidates: list[dict] = []

    for query in queries:
        if len(candidates) >= max_results:
            break
        log.info("DDG X-Ray: %s", query)
        try:
            results = _ddg_search(query, max_results=10)
        except Exception as exc:
            log.warning("DDG search failed: %s", exc)
            continue

        for r in results:
            c = _parse_item(r, position, location)
            if not c or c["url"] in seen:
                continue
            if not _matches_required(c, position, location, keywords):
                continue
            seen.add(c["url"])
            c.pop("_searchable", None)
            candidates.append(c)

    log.info("X-Ray complete: %d matching candidates for '%s' in '%s'", len(candidates), position, location)
    return candidates[:max_results]
