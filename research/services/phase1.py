"""
PL Research Tool — Phase 1 (web / library form).

A single, self-contained module that turns the user's filters (budget,
category, seasonal) into a ranked product shortlist, reading the stored dataset
through the Django ORM (`products.Product`). No psycopg2 connection, no Excel,
no Black Box ingest, no prompts — every entry point returns plain Python data
for the frontend to render.

Public API
----------
    run_query(budget, category, include_seasonal, limit=10, offset=0) -> dict
        The high-level call. Returns:
        {
          "products": [...],   # the requested page (limit/offset) of the pool
          "pool":     [...],   # the full ranked pool (for client-side paging)
          "pool_size": int,
          "funnel":   [...],   # per-filter exclusion counts
          "meta":     {...},   # budget / category / seasonal / bands echoed back
        }
    fetch_pool(...)    -> list[dict]   # the full ranked pool
    filter_funnel(...) -> list[dict]   # how many each filter removed

Optional Claude judgement adjustments (±10 pts) are applied to the top of the
pool when USE_CLAUDE_PHASE1=true and ANTHROPIC_API_KEY is set; otherwise the
deterministic score stands alone.
"""

from __future__ import annotations

import json
import os
import random
from dataclasses import dataclass
from decimal import Decimal
from hashlib import sha1
from pathlib import Path
from typing import Dict, List, Optional

from django.conf import settings
from django.db.models import F, Q

from products.models import Product

# ── Tunables (ported verbatim from the original query.py) ────────────────────
BUDGET_REVENUE_MULTIPLIER = 1.52
REVENUE_BAND_WIDTH = 1500.0       # band = [cap - 1500, cap]
REVENUE_PROXIMITY_MAX = 30.0

MIN_RATING = 3.8                  # strictly greater than
MIN_LISTING_AGE_MONTHS = 6
MIN_MONTHLY_SALES = 50
MAX_DECLINE_PCT = -25.0           # remove if 90-day trend declines more than 25%

# Budget-tiered review window (review_min, review_max). max=None => no upper cap.
REVIEW_TIERS = [
    (20000, 200, None),     # Premium    : >= £20,000  -> 200+
    (10000, 200, 12000),    # Very High  : £10k-£20k   -> 200-12000
    (5000,  200, 8000),     # High       : £5k-£10k    -> 200-8000
    (2000,  200, 5000),     # Medium     : £2k-£5k     -> 200-5000
    (0,     200, 3000),     # Low        : < £2,000    -> 200-3000
]
DEFAULT_REVIEW_BAND = (200, 3000)

DEFAULT_POOL = 100
CLAUDE_TOP_N = 20
CLAUDE_MODEL = "claude-opus-4-8"


# ── Pure scoring helpers ─────────────────────────────────────────────────────
def review_band_for_budget(budget: Optional[float]):
    """(review_min, review_max) for the budget tier; max None = no cap."""
    if budget is None:
        return DEFAULT_REVIEW_BAND
    for lower, rmin, rmax in REVIEW_TIERS:
        if budget >= lower:
            return rmin, rmax
    return DEFAULT_REVIEW_BAND


def max_revenue_for_budget(budget: Optional[float]) -> Optional[float]:
    """£ budget -> max monthly revenue cap (budget × 1.52). None if no budget."""
    return None if budget is None else budget * BUDGET_REVENUE_MULTIPLIER


def revenue_band(max_revenue: Optional[float]):
    """(min_revenue, max_revenue) band, or (None, None) if no cap."""
    if max_revenue is None:
        return None, None
    return max_revenue - REVENUE_BAND_WIDTH, max_revenue


def revenue_proximity(rev: Optional[float], max_revenue: Optional[float]) -> float:
    """0–30: 0 at (cap − 1,500), 30 at the cap. 0 if no budget or no revenue."""
    if rev is None or max_revenue is None:
        return 0.0
    lo = max_revenue - REVENUE_BAND_WIDTH
    frac = (rev - lo) / REVENUE_BAND_WIDTH
    return round(REVENUE_PROXIMITY_MAX * max(0.0, min(1.0, frac)), 1)


def _to_num(v) -> Optional[float]:
    if v is None:
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


def budget_review_points(reviews: Optional[float], budget: Optional[float]) -> float:
    """0-15 review points: 15 at the band's rmin, decreasing linearly to 0 at rmax."""
    r = _to_num(reviews)
    if r is None or budget is None:
        return 0.0
    rmin, rmax = review_band_for_budget(budget)
    if rmin is None:
        return 0.0
    if r <= rmin:
        return 15.0
    if rmax is None or rmax <= rmin:
        rmax = rmin * 8.0 if rmin > 0 else rmin + 1000.0
    frac = (r - rmin) / (rmax - rmin)
    pts = max(0.0, 1.0 - min(1.0, frac)) * 15.0
    return round(pts, 1)


# ── ORM filter steps ─────────────────────────────────────────────────────────
def _filter_steps(budget, category, include_seasonal):
    """Ordered [(label, Q), ...] mirroring the original SQL funnel exactly.

    Boolean exclusions use the SQL "IS NOT TRUE" semantics (keep rows that are
    FALSE *or* NULL), expressed as `Q(field=False) | Q(field__isnull=True)`.
    """
    rmin, rmax = review_band_for_budget(budget)
    if rmax is None:
        review_label = f"Review count ({rmin}+)"
        review_q = Q(review_count__gte=rmin)
    else:
        review_label = f"Review count ({rmin}-{rmax})"
        review_q = Q(review_count__gte=rmin) & Q(review_count__lte=rmax)

    steps = []
    if category:
        steps.append((f"Category = {category}", Q(category__iexact=category)))
    steps += [
        (f"Min rating (> {MIN_RATING})", Q(rating__gt=MIN_RATING)),
        (f"Min listing age ({MIN_LISTING_AGE_MONTHS}mo)", Q(listing_age_months__gte=MIN_LISTING_AGE_MONTHS)),
        (f"Min monthly sales ({MIN_MONTHLY_SALES})", Q(estimated_sales__gte=MIN_MONTHLY_SALES)),
        (f"Sales trend (> {MAX_DECLINE_PCT:.0f}%)",
         Q(sales_trend_90d__isnull=True) | Q(sales_trend_90d__gt=MAX_DECLINE_PCT)),
        (review_label, review_q),
        ("Amazon-as-seller", Q(is_amazon_seller=False) | Q(is_amazon_seller__isnull=True)),
        ("Compatibility products", Q(is_compatibility=False) | Q(is_compatibility__isnull=True)),
        ("Global brands", Q(is_global_brand=False) | Q(is_global_brand__isnull=True)),
    ]
    if not include_seasonal:
        steps.append(("Seasonal products", Q(is_seasonal=False) | Q(is_seasonal__isnull=True)))
    max_revenue = max_revenue_for_budget(budget)
    if max_revenue is not None:
        lo, hi = revenue_band(max_revenue)
        steps.append((f"Revenue band (£{lo:,.0f}-£{hi:,.0f})",
                      Q(estimated_revenue__gte=lo) & Q(estimated_revenue__lte=hi)))
    return steps


def _normalise(row: Dict) -> Dict:
    """Cast ORM Decimals to float so the rows are JSON-/template-friendly."""
    return {k: (float(v) if isinstance(v, Decimal) else v) for k, v in row.items()}


# ── Funnel ───────────────────────────────────────────────────────────────────
def filter_funnel(budget: Optional[float] = None, *, category: Optional[str] = None,
                  include_seasonal: bool = True) -> List[Dict]:
    """Sequential exclusion report: for each filter, how many products it removed
    and how many remain. Returns ordered dicts {filter, removed, remaining}."""
    steps = _filter_steps(budget, category, include_seasonal)
    qs = Product.objects.all()
    remaining = qs.count()
    report = [{"filter": "All stored products", "removed": 0, "remaining": remaining}]
    for label, q in steps:
        qs = qs.filter(q)
        now = qs.count()
        report.append({"filter": label, "removed": remaining - now, "remaining": now})
        remaining = now
    return report


# ── Pool ─────────────────────────────────────────────────────────────────────
def _soft_shuffle(rows: List[Dict], jitter: float = 3.0) -> List[Dict]:
    def key(r):
        return float(r.get("opportunity_final") or 0.0) + random.uniform(-jitter, jitter)
    return sorted(rows, key=key, reverse=True)


def fetch_pool(budget: Optional[float] = None, *, category: Optional[str] = None,
               include_seasonal: bool = True, pool_size: Optional[int] = None,
               shuffle: bool = False) -> List[Dict]:
    """Return ALL products (dicts) passing the filters, ranked by final score
    (base 0–85 + revenue proximity 0–30 + budget review points 0–15), after the
    built-in quality filters + revenue band + seasonal option.

    `pool_size=None` (default) returns every matching product (could be 1,000+);
    pass an int to cap it. `shuffle=False` (default) gives a deterministic
    ranking so the web pages stay stable across requests; pass True for the
    original soft-shuffle variety.
    """
    steps = _filter_steps(budget, category, include_seasonal)
    qs = Product.objects.all()
    for _, q in steps:
        qs = qs.filter(q)
    qs = qs.order_by(F("opportunity_score").desc(nulls_last=True))
    rows = [_normalise(r) for r in qs.values()]

    max_revenue = max_revenue_for_budget(budget)
    for r in rows:
        base = float(r.get("opportunity_score") or 0.0)
        prox = revenue_proximity(_to_num(r.get("estimated_revenue")), max_revenue)
        r["revenue_proximity"] = prox
        review_pts = budget_review_points(r.get("review_count"), budget)
        r["budget_review_points"] = review_pts
        r["opportunity_final"] = round(base + prox + review_pts, 1)

    rows = _soft_shuffle(rows) if shuffle else sorted(
        rows, key=lambda r: r.get("opportunity_final") or 0.0, reverse=True
    )

    _apply_claude_scoring(rows, budget, top_n=CLAUDE_TOP_N)
    rows = sorted(rows, key=lambda r: r.get("opportunity_final") or 0.0, reverse=True)
    return rows[:pool_size] if pool_size else rows


# ── High-level entry point ───────────────────────────────────────────────────
def run_query(budget=None, category=None, include_seasonal=True, *, limit: int = 10,
              offset: int = 0, pool_size: Optional[int] = None, shuffle: bool = False,
              with_funnel: bool = True) -> Dict:
    """Run Phase 1 end-to-end and return everything the frontend needs.

    `budget` may be a number or a string ("£2,000" / "2000" / "") — it is parsed
    leniently. `category` "" is treated as no filter. `limit`/`offset` page the
    ranked pool (limit=0 returns the whole pool from `offset`).
    """
    budget = _parse_budget(budget)
    category = (str(category).strip() if category is not None else "") or None

    pool = fetch_pool(budget, category=category, include_seasonal=include_seasonal,
                      pool_size=pool_size, shuffle=shuffle)

    page = pool[offset: offset + limit] if limit else pool[offset:]

    rmin, rmax = review_band_for_budget(budget)
    max_rev = max_revenue_for_budget(budget)
    lo, hi = revenue_band(max_rev)
    meta = {
        "budget": budget,
        "category": category,
        "include_seasonal": include_seasonal,
        "pool_size": len(pool),
        "limit": limit,
        "offset": offset,
        "review_band": {"min": rmin, "max": rmax},
        "revenue_band": None if max_rev is None else {"low": round(lo, 2), "high": round(hi, 2)},
    }
    return {
        "products": page,
        "pool": pool,
        "pool_size": len(pool),
        "funnel": filter_funnel(budget, category=category, include_seasonal=include_seasonal)
                  if with_funnel else [],
        "meta": meta,
    }


def _parse_budget(raw) -> Optional[float]:
    if raw is None:
        return None
    if isinstance(raw, (int, float)):
        return float(raw)
    s = str(raw).replace("£", "").replace(",", "").replace("$", "").strip()
    try:
        return float(s) if s else None
    except ValueError:
        return None


# ── Optional Claude judgement scoring (opt-in via USE_CLAUDE_PHASE1) ──────────
_CACHE_DIR = Path(settings.BASE_DIR) / "output" / "llm_cache"


class Phase1ApiError(RuntimeError):
    """Raised when the Claude client cannot be constructed or a call fails."""


@dataclass
class ScoreResult:
    asin: str
    llm_adj: float
    llm_reason: str
    llm_tags: List[str]


def load_client():
    """Construct an Anthropic client from the project .env. Raises Phase1ApiError
    when the key or package is missing."""
    try:
        from dotenv import load_dotenv
    except Exception as e:  # pragma: no cover
        raise Phase1ApiError("python-dotenv is required to load .env") from e
    try:
        import anthropic
    except Exception as e:  # pragma: no cover
        raise Phase1ApiError("The 'anthropic' package is required but not installed.") from e

    load_dotenv(dotenv_path=Path(settings.BASE_DIR) / ".env")
    if not os.environ.get("ANTHROPIC_API_KEY"):
        raise Phase1ApiError("ANTHROPIC_API_KEY not found in .env")
    try:
        return anthropic.Anthropic()
    except Exception as e:
        raise Phase1ApiError(f"Could not initialise the Anthropic client: {e}") from e


def _cache_path_for(asins: List[str], budget: Optional[float]) -> Path:
    key = "|".join(sorted(str(a).upper() for a in asins)) + f"|{budget}"
    return _CACHE_DIR / f"phase1_claude_batch_{sha1(key.encode('utf-8')).hexdigest()}.json"


def _build_system_and_user(products: List[Dict], budget: Optional[float]):
    system_text = (
        "You are an expert Amazon UK private-label analyst.\n"
        "Given a small set of product metrics and the user's scoring rules, produce a concise\n"
        "numeric adjustment (`llm_adj`) between -10 and +10 to add to the product's numeric\n"
        "opportunity score. Also return a brief, specific reason explaining your adjustment.\n\n"
        "Rules (apply deterministically in Python; you are NOT replacing them):\n"
        " - The final score is made of two pillars: Market Entry (50) and Growth (50).\n"
        " - Review counts, BSR stability, fulfillment (FBA/FBM), and price band are primary factors.\n\n"
        "For each product, analyze: market saturation, pricing opportunity, demand signals,\n"
        "seasonality risk, and brand risk.\n\n"
        "Your output must be a JSON array of objects exactly matching the schema: `asin` (string),\n"
        "`llm_adj` (float, -10 to +10), `llm_reason` (string, brief), `llm_tags` (array of short tags).\n"
    )
    lines = []
    for p in products:
        asin = str(p.get("asin") or "").upper()
        lines.append(
            f"ASIN: {asin}\n  Title: {str(p.get('title') or '')[:120]}\n"
            f"  Price: {p.get('price') or ''}\n  Reviews: {p.get('review_count') or ''}\n"
            f"  Rating: {p.get('rating') or ''}\n  BSR: {p.get('bsr') or ''}\n"
            f"  Sales/mo: {p.get('estimated_sales') or ''}\n  Revenue/mo: {p.get('estimated_revenue') or ''}\n"
            f"  Trend90d: {p.get('sales_trend_90d') or ''}\n  Fulfillment: {p.get('fulfillment') or ''}\n"
            f"  Weight_kg: {p.get('weight_kg') or ''}\n  Seasonal: {p.get('is_seasonal') or False}\n"
        )
    user_block = (
        f"BATCH: {len(products)} products. Budget: £{budget}\n\n"
        "For each product block below, output one JSON object with `asin`, `llm_adj`,\n"
        "`llm_reason` (why — be specific), and `llm_tags` (quick labels).\n\nPRODUCTS:\n\n"
        + "\n---\n".join(lines)
    )
    system = [{"type": "text", "text": system_text, "cache_control": {"type": "ephemeral"}}]
    return system, user_block


def score_products_with_claude(client, products: List[Dict], budget: Optional[float],
                               *, max_batch: int = CLAUDE_TOP_N) -> List[ScoreResult]:
    """Score up to `max_batch` products in one batched Claude call (cached on disk
    to avoid repeated token cost). Raises Phase1ApiError on failure."""
    if not products:
        return []
    batch = products[:max_batch]
    asins = [str(p.get("asin") or "").upper() for p in batch]
    cache_path = _cache_path_for(asins, budget)
    if cache_path.exists():
        try:
            return [ScoreResult(**d) for d in json.loads(cache_path.read_text(encoding="utf-8"))]
        except Exception:
            pass

    system, user = _build_system_and_user(batch, budget)
    try:
        resp = client.messages.parse(
            model=CLAUDE_MODEL, max_tokens=16000,
            thinking={"type": "adaptive"}, output_config={"effort": "medium"},
            system=system, messages=[{"role": "user", "content": user}],
        )
    except Exception as e:
        raise Phase1ApiError(f"Claude Phase 1 batch scoring failed: {e}") from e

    text = getattr(resp, "response", None) or getattr(resp, "text", None) or str(resp)
    try:
        j = json.loads(text[text.index("["): text.rindex("]") + 1])
    except Exception:
        try:
            j = json.loads(text)
        except Exception as e:
            raise Phase1ApiError(f"Could not parse Claude response as JSON: {e}") from e

    results, out_data = [], []
    for o in j:
        try:
            adj = float(o.get("llm_adj") or 0.0)
        except Exception:
            adj = 0.0
        tags = o.get("llm_tags") or []
        if not isinstance(tags, list):
            tags = [str(tags)]
        sr = ScoreResult(asin=str(o.get("asin") or "").upper(), llm_adj=adj,
                         llm_reason=str(o.get("llm_reason") or "").strip(),
                         llm_tags=[str(t) for t in tags])
        results.append(sr)
        out_data.append(sr.__dict__)

    try:
        _CACHE_DIR.mkdir(parents=True, exist_ok=True)
        cache_path.write_text(json.dumps(out_data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass
    return results


def _apply_claude_scoring(rows: List[Dict], budget: Optional[float], top_n: int = CLAUDE_TOP_N) -> None:
    """If USE_CLAUDE_PHASE1 is truthy and Claude is available, adjust the top
    `top_n` rows' `opportunity_final` in place (adds llm_adj/llm_reason/llm_tags).
    Silent no-op when disabled or unavailable — never raises into the query path."""
    if os.environ.get("USE_CLAUDE_PHASE1", "").lower() not in ("true", "1", "yes"):
        return
    try:
        client = load_client()
        results = score_products_with_claude(client, rows[:top_n], budget, max_batch=top_n)
    except Phase1ApiError:
        return

    by_asin = {r.asin.upper(): r for r in results}
    for row in rows:
        res = by_asin.get(str(row.get("asin") or "").upper())
        if res:
            row["llm_adj"] = res.llm_adj
            row["llm_reason"] = res.llm_reason
            row["llm_tags"] = res.llm_tags
            row["opportunity_final"] = max(0.0, round((row.get("opportunity_final") or 0.0) + res.llm_adj, 1))
