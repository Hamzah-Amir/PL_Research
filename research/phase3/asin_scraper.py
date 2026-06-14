"""
Amazon search-results ASIN scraper — pure script (no API, no OpenClaw, no
browser). Feeds the Keepa pipeline: keyword in -> ordered ASIN list out.

Given a search keyword (or a full search/listing URL), it walks the paginated
results via `curl_cffi` with a real Chrome TLS fingerprint and extracts every
ASIN in ranked order, flagging sponsored placements separately from organic
ones. `AMAZON_COOKIE` from the project `.env` is used when present (optional —
search pages are not login-gated, but a session cookie reduces bot friction).

Anti-bot handling: CAPTCHA/robot-check pages are detected and retried with
backoff; a polite jitter delay runs between page fetches. Nothing is fabricated
— a blocked or empty page is reported in `note`, not papered over (Rule 0).

Deps: curl_cffi, beautifulsoup4 (in the project venv).
"""

from __future__ import annotations

import os
import random
import re
import time
from pathlib import Path
from typing import Dict, List, Optional
from urllib.parse import quote_plus

try:
    from curl_cffi import requests as _http
    _CURL = True
except ImportError:  # pragma: no cover
    import requests as _http  # type: ignore
    _CURL = False

try:
    from bs4 import BeautifulSoup
except ImportError as _e:  # pragma: no cover
    BeautifulSoup = None  # type: ignore
    _BS4_ERR = _e
else:
    _BS4_ERR = None

_ROOT = Path(__file__).resolve().parent.parent
_ENV_PATH = _ROOT / ".env"

_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-GB,en;q=0.9",
    "Upgrade-Insecure-Requests": "1",
}

_ASIN_RE = re.compile(r"^[A-Z0-9]{10}$")
_CAPTCHA_MARKERS = (
    "api-services-support@amazon.com",
    "Enter the characters you see below",
    "Type the characters you see in this image",
)


def _cookie() -> str:
    try:
        from dotenv import load_dotenv
        load_dotenv(dotenv_path=_ENV_PATH)
    except Exception:  # noqa: BLE001
        pass
    return (os.environ.get("AMAZON_COOKIE") or "").strip()


def _get(url: str, cookie: str, timeout: float = 30):
    headers = dict(_HEADERS)
    if cookie:
        headers["Cookie"] = cookie
    if _CURL:
        return _http.get(url, headers=headers, timeout=timeout, impersonate="chrome")
    return _http.get(url, headers=headers, timeout=timeout)


def _is_captcha(html: str) -> bool:
    return any(m in html for m in _CAPTCHA_MARKERS)


def _fetch_page(url: str, cookie: str, retries: int = 3) -> Optional[str]:
    """Fetch one results page; retry on CAPTCHA/5xx with backoff. None = blocked."""
    for attempt in range(retries):
        try:
            r = _get(url, cookie)
        except Exception:  # noqa: BLE001 — network hiccup, retry
            time.sleep(2.0 * (attempt + 1))
            continue
        if r.status_code == 200 and not _is_captcha(r.text):
            return r.text
        if r.status_code == 404:
            return None
        time.sleep(3.0 * (attempt + 1) + random.uniform(0, 2))
    return None


def _is_sponsored(block) -> bool:
    if block.select_one(
        '[data-component-type="sp-sponsored-result"], '
        ".puis-sponsored-label-text, .s-sponsored-label-info-icon"
    ):
        return True
    for el in block.select("a[aria-label], span[aria-label]"):
        if "sponsored" in (el.get("aria-label") or "").lower():
            return True
    return False


def _parse_results(html: str) -> List[Dict]:
    """Extract {asin, title, sponsored} blocks from one search page, in order."""
    soup = BeautifulSoup(html, "html.parser")
    out, seen = [], set()
    for block in soup.select('div[data-component-type="s-search-result"]'):
        asin = (block.get("data-asin") or "").strip().upper()
        if not _ASIN_RE.fullmatch(asin) or asin in seen:
            continue
        seen.add(asin)
        title_el = block.select_one("h2 span") or block.select_one("h2")
        out.append({
            "asin": asin,
            "title": title_el.get_text(" ", strip=True) if title_el else "",
            "sponsored": _is_sponsored(block),
        })
    return out


def _has_next_page(html: str) -> bool:
    soup = BeautifulSoup(html, "html.parser")
    nxt = soup.select_one("a.s-pagination-next")
    if nxt is not None:
        return True
    # Disabled "Next" renders as a span, an explicit signal of the last page.
    return False


def _search_url(keyword: str, marketplace: str, page: int) -> str:
    base = f"https://www.amazon.{marketplace}/s?k={quote_plus(keyword)}"
    return f"{base}&page={page}" if page > 1 else base


def _page_url(url: str, page: int) -> str:
    stripped = re.sub(r"([&?])page=\d+", r"\1", url).rstrip("&?")
    if page == 1:
        return stripped
    sep = "&" if "?" in stripped else "?"
    return f"{stripped}{sep}page={page}"


def scrape_asins(
    keyword: Optional[str] = None,
    url: Optional[str] = None,
    marketplace: str = "co.uk",
    max_pages: int = 1,
    include_sponsored: bool = True,
    delay: float = 1.5,
) -> Dict:
    """Scrape every ASIN from Amazon search results for `keyword` (or a full
    search/listing `url`), walking pages 1..max_pages in ranked order
    (default: page 1 only).

    Returns {keyword, url, marketplace, pages_fetched, pages_blocked, count,
    results[{asin, title, page, position, sponsored}], note}. Duplicates across
    pages keep their first (highest-ranked) occurrence. Never raises for normal
    failures — reports via note.
    """
    if BeautifulSoup is None:
        return {"error": f"beautifulsoup4 not installed ({_BS4_ERR})", "results": [], "count": 0}
    if not keyword and not url:
        return {"error": "Provide a keyword or a url.", "results": [], "count": 0}

    cookie = _cookie()
    results: List[Dict] = []
    seen: set = set()
    pages_fetched, pages_blocked = 0, 0
    position = 0

    for page in range(1, max_pages + 1):
        page_url = _page_url(url, page) if url else _search_url(keyword, marketplace, page)
        html = _fetch_page(page_url, cookie)
        if html is None:
            pages_blocked += 1
            break
        pages_fetched += 1
        blocks = _parse_results(html)
        if not blocks:
            break
        for b in blocks:
            if b["asin"] in seen:
                continue
            if not include_sponsored and b["sponsored"]:
                continue
            seen.add(b["asin"])
            position += 1
            results.append({**b, "page": page, "position": position})
        if not _has_next_page(html):
            break
        time.sleep(delay + random.uniform(0, 1))

    note = ""
    if pages_blocked:
        note = (
            f"Page {pages_fetched + 1} was blocked (CAPTCHA or repeated errors) — "
            f"results cover the first {pages_fetched} page(s) only."
        )
    elif pages_fetched == 0:
        note = "No pages could be fetched."
    elif not results:
        note = "Pages fetched but no search results found — selector mismatch or empty search."

    return {
        "keyword": keyword,
        "url": url,
        "marketplace": marketplace,
        "pages_fetched": pages_fetched,
        "pages_blocked": pages_blocked,
        "count": len(results),
        "sponsored_count": sum(1 for r in results if r["sponsored"]),
        "results": results,
        "note": note,
        "scraper": "curl_cffi" if _CURL else "requests",
    }


def asin_list(result: Dict, organic_only: bool = False) -> List[str]:
    """Flatten a scrape result into the ordered ASIN list consumed by the
    Keepa stage (empty if nothing usable)."""
    rows = result.get("results") or []
    if organic_only:
        rows = [r for r in rows if not r.get("sponsored")]
    return [r["asin"] for r in rows]


if __name__ == "__main__":
    import argparse
    import csv
    import json
    import sys
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    ap = argparse.ArgumentParser(description="Amazon search-results ASIN scraper (direct, no API).")
    ap.add_argument("keyword", nargs="?", help="Search keyword (omit if using --url)")
    ap.add_argument("--url", help="Full Amazon search/listing URL instead of a keyword")
    ap.add_argument("--marketplace", default="co.uk")
    ap.add_argument("--max-pages", type=int, default=1)
    ap.add_argument("--organic-only", action="store_true", help="Skip sponsored placements")
    ap.add_argument("--delay", type=float, default=1.5, help="Base seconds between page fetches")
    ap.add_argument("--csv", metavar="PATH", help="Write results to a CSV file")
    ap.add_argument("--asins-only", action="store_true", help="Print one ASIN per line")
    a = ap.parse_args()
    res = scrape_asins(
        keyword=a.keyword, url=a.url, marketplace=a.marketplace,
        max_pages=a.max_pages, include_sponsored=not a.organic_only, delay=a.delay,
    )
    if a.csv:
        with open(a.csv, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=["position", "page", "asin", "sponsored", "title"])
            w.writeheader()
            for row in res.get("results", []):
                w.writerow({k: row[k] for k in w.fieldnames})
        print(f"Wrote {res.get('count', 0)} ASINs to {a.csv}")
        if res.get("note"):
            print(f"Note: {res['note']}")
    elif a.asins_only:
        print("\n".join(asin_list(res, organic_only=a.organic_only)) or res.get("note") or res.get("error", ""))
    else:
        print(json.dumps(res, ensure_ascii=False, indent=2))
