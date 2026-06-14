"""
Competitor reviews for Phase 3 via the FlyBy / RapidAPI **Real-Time Amazon
Data** API (key in `.env` as RAPIDAPI_KEY) — replaces the cookie/TLS HTTP
scraper and the OpenClaw browser route entirely.

GATED ACCESS: Amazon login-gates the full review list, and the API's
`star_rating` / `page` parameters only work when a logged-in session cookie is
passed via the `cookie` query param. `AMAZON_COOKIE` from `.env` is sent on
every call; with a fresh cookie the full 1-2 star list is paginated, without
one (or once it expires) the API silently returns the ~8-10 public
detail-page reviews — detected and reported in `note`, never papered over
(Rule 0).

Selection rule (user decision 2026-06-12):
  * weaknesses — ALL 1-star and 2-star reviews (paginated, deduped).
  * strengths  — at most FIVE 4-5 star reviews (one TOP_REVIEWS page).
  * 3-star reviews are EXCLUDED.

Cost: ~3-15 requests/ASIN depending on negative-review volume
(free tier: 2,000/month). Deps: requests, python-dotenv.
"""

from __future__ import annotations

import os
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import requests

_ROOT = Path(__file__).resolve().parent.parent
_ENV_PATH = _ROOT / ".env"

HOST = "real-time-amazon-data.p.rapidapi.com"
MAX_STRENGTHS = 5
MAX_PAGES_PER_STAR = 10  # 10 reviews/page -> up to ~100 reviews per star band

# Amazon marketplace suffix -> API country code.
_COUNTRY = {"co.uk": "GB", "com": "US", "de": "DE", "fr": "FR", "it": "IT",
            "es": "ES", "ca": "CA", "com.au": "AU", "co.jp": "JP", "in": "IN"}


class FlybyError(RuntimeError):
    """Raised when the reviews API cannot complete a request."""


def _env(name: str) -> str:
    try:
        from dotenv import load_dotenv
        load_dotenv(dotenv_path=_ENV_PATH)
    except Exception:  # noqa: BLE001
        pass
    return (os.environ.get(name) or "").strip()


def _get(path: str, key: str, params: Dict, timeout: float) -> Dict:
    r = requests.get(
        f"https://{HOST}{path}",
        headers={"X-RapidAPI-Key": key, "X-RapidAPI-Host": HOST},
        params=params, timeout=timeout,
    )
    if r.status_code != 200:
        raise FlybyError(f"{path} HTTP {r.status_code}: {r.text[:200]}")
    return (r.json() or {}).get("data") or {}


def _norm(rev: Dict) -> Tuple[str, Dict]:
    """Normalise one API review into (id, {stars,title,body,verified,date})."""
    rid = rev.get("review_id") or f"{rev.get('review_title')}|{str(rev.get('review_comment'))[:40]}"
    try:
        stars = int(round(float(rev.get("review_star_rating"))))
    except (TypeError, ValueError):
        stars = None
    return rid, {
        "stars": stars,
        "title": (rev.get("review_title") or "").strip(),
        "body": (rev.get("review_comment") or "").strip(),
        "verified": rev.get("is_verified_purchase"),
        "date": rev.get("review_date"),
    }


def fetch_reviews(asin: str, marketplace: str = "co.uk", timeout: float = 60) -> Dict:
    """Fetch + select reviews for one ASIN per the project rule
    (all 1-2★ / max five 4-5★ / no 3★).

    Returns {asin, marketplace, gated_access, total_ratings, distribution,
    strengths[...], weaknesses[...], weaknesses_count, note, source}.
    Never raises for normal failures — reports via the "error" key instead.
    """
    asin = (asin or "").strip().upper()
    if not re.fullmatch(r"[A-Z0-9]{10}", asin):
        return {"asin": asin, "error": f"Not a valid ASIN: {asin!r}",
                "weaknesses": [], "weaknesses_count": 0}
    key = _env("RAPIDAPI_KEY")
    if not key:
        return {"asin": asin, "error": f"RAPIDAPI_KEY not found. Add it to {_ENV_PATH}.",
                "weaknesses": [], "weaknesses_count": 0}
    cookie = _env("AMAZON_COOKIE")
    country = _COUNTRY.get(marketplace.lower(), marketplace.upper())

    dist: Dict[str, int] = {}
    total_ratings = None
    weaknesses: Dict[str, Dict] = {}
    strengths: Dict[str, Dict] = {}
    errors: List[str] = []
    gated = False
    cookie_dead = False

    def _page(params: Dict) -> List[Dict]:
        nonlocal dist, total_ratings
        data = _get("/product-reviews", key, params, timeout)
        dist = {str(k): v for k, v in (data.get("rating_distribution") or {}).items()} or dist
        total_ratings = data.get("total_ratings") or total_ratings
        return data.get("reviews") or []

    # ── 1. Weaknesses: ALL 1★ + 2★, paginated (needs the session cookie) ─────
    if cookie:
        for star_val, tag in (("1_STARS", 1), ("2_STARS", 2)):
            if cookie_dead:
                break
            for page in range(1, MAX_PAGES_PER_STAR + 1):
                try:
                    revs = _page({"asin": asin, "country": country, "star_rating": star_val,
                                  "sort_by": "MOST_RECENT", "page": page, "cookie": cookie})
                except Exception as e:  # noqa: BLE001
                    errors.append(str(e))
                    break
                if not revs:
                    break
                got = dict(_norm(r) for r in revs)
                # Filter honoured? Off-band stars on page 1 mean the cookie is
                # dead and Amazon served the unfiltered public sample.
                if any(v["stars"] not in (tag, None) for v in got.values()):
                    cookie_dead = True
                    break
                new = 0
                for rid, v in got.items():
                    if rid not in weaknesses and (v["title"] or v["body"]):
                        weaknesses[rid] = v
                        new += 1
                gated = True
                if new == 0:  # repeated page -> end of the list
                    break

    # ── 2. Strengths: one TOP_REVIEWS page, keep max five 4-5★ ───────────────
    try:
        revs = _page({"asin": asin, "country": country, "star_rating": "ALL",
                      "sort_by": "TOP_REVIEWS", "page": 1,
                      **({"cookie": cookie} if cookie and not cookie_dead else {})})
    except Exception as e:  # noqa: BLE001
        errors.append(str(e))
        revs = []
    sample_negatives = 0
    for r in revs:
        rid, v = _norm(r)
        if not (v["title"] or v["body"]):
            continue
        if v["stars"] in (4, 5) and len(strengths) < MAX_STRENGTHS and rid not in strengths:
            strengths[rid] = v
        # Cookieless fallback: salvage any 1-2★ present in the public sample.
        if v["stars"] in (1, 2) and rid not in weaknesses:
            weaknesses[rid] = v
            sample_negatives += 1
    # 3★ are excluded by design (user rule).

    if not weaknesses and not strengths and errors:
        return {"asin": asin, "error": "; ".join(errors), "weaknesses": [], "weaknesses_count": 0}

    # ── 3. Honest note ────────────────────────────────────────────────────────
    note = ""
    try:
        neg_pct = int(dist.get("1") or 0) + int(dist.get("2") or 0)
    except (TypeError, ValueError):
        neg_pct = 0
    if cookie_dead:
        note = ("AMAZON_COOKIE appears expired/invalid — Amazon ignored the star filter, "
                "so only the public detail-page sample was available. Refresh the cookie "
                "in .env for the full 1-2 star list.")
    elif not cookie:
        note = "No AMAZON_COOKIE in .env — only the public detail-page sample was available."
    if not gated and not weaknesses and neg_pct:
        note += (" " if note else "") + (
            f"{neg_pct}% of all ratings are 1-2 star but none were retrievable.")

    return {
        "asin": asin,
        "marketplace": marketplace,
        "gated_access": gated,
        "total_ratings": total_ratings,
        "distribution": dist,
        "strengths": list(strengths.values()),
        "weaknesses": list(weaknesses.values()),
        "weaknesses_count": len(weaknesses),
        "note": note,
        "source": ("flyby/rapidapi product-reviews — full gated 1-2★ list via session cookie"
                   if gated else
                   "flyby/rapidapi product-reviews — public detail-page sample only"),
    }


def reviews_text(result: Optional[Dict]) -> str:
    """Render a fetch result into the STRENGTHS/WEAKNESSES block consumed by
    `claude_client.analyze_reviews` (empty if nothing usable)."""
    if not result or result.get("error"):
        return ""
    out: List[str] = []
    dist = result.get("distribution") or {}
    if dist:
        out.append("=== RATING DISTRIBUTION (all ratings) ===")
        line = ", ".join(f"{s}*: {dist.get(s, 0)}%" for s in ("5", "4", "3", "2", "1"))
        if result.get("total_ratings"):
            line += f"  (of {result['total_ratings']} total ratings)"
        out.append(line)
    if result.get("strengths"):
        out.append("\n=== STRENGTHS (4-5 star reviews, max 5) ===")
        for i, r in enumerate(result["strengths"], start=1):
            head = (r.get("title", "") + " — ") if r.get("title") else ""
            out.append(f"Review {i}: [{r.get('stars')}*] {head}{r.get('body', '')}")
    if result.get("weaknesses"):
        src = "full 1-2 star list" if result.get("gated_access") else "detail-page sample"
        out.append(f"\n=== WEAKNESSES (1-2 star reviews — {src}) ===")
        for i, r in enumerate(result["weaknesses"], start=1):
            head = (r.get("title", "") + " — ") if r.get("title") else ""
            out.append(f"Review {i}: [{r.get('stars')}*] {head}{r.get('body', '')}")
    if result.get("note"):
        out.append(f"\nNote: {result['note']}")
    return "\n".join(out)


if __name__ == "__main__":
    import argparse
    import json
    import sys
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    ap = argparse.ArgumentParser(description="Phase 3 competitor reviews via the FlyBy/RapidAPI API.")
    ap.add_argument("asin")
    ap.add_argument("--marketplace", default="co.uk")
    ap.add_argument("--text", action="store_true")
    a = ap.parse_args()
    res = fetch_reviews(a.asin, a.marketplace)
    if a.text:
        print(reviews_text(res) or res.get("error") or "(no reviews)")
    else:
        print(json.dumps(res, ensure_ascii=False, indent=2))
