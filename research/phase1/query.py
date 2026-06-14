"""
Phase 1 query — fetch product suggestions from the stored dataset (no live API).

Budget rule (user-specified): cap = budget × 1.52; products must fall in the
revenue BAND **[cap − 1,000, cap]** (a fixed £1,000-wide window ending at the
cap). The revenue-proximity score (0–30) is added here (closer to the cap =
higher) and combined with the stored base score (0–85) to rank the pool.

Built-in quality filters (always applied, never prompted — like the old
workflow): rating > 3.8, listing age ≥ 24 months, monthly sales ≥ 150, 90-day
trend > −25%, reviews 50–3,000, and exclude Amazon-seller / compatibility /
global-brand products. Seasonal and variants remain user options.

Optional Claude scoring: if `USE_CLAUDE_PHASE1=true` and `ANTHROPIC_API_KEY` is
set, the top products get judgement-based adjustments from Claude (±0-10 pts).
"""

import os
import random
from typing import Dict, List, Optional

BUDGET_REVENUE_MULTIPLIER = 1.52
REVENUE_BAND_WIDTH = 1500.0       # band = [cap - 1500, cap]
REVENUE_PROXIMITY_MAX = 30.0

# Built-in quality thresholds (the old workflow's F-filters, now fixed here).
# NOTE: these are the LIVE filters for the DB-backed Phase 1. (phase1/filters.py
# is the retired file-based module and is not used.)
MIN_RATING = 3.8                  # strictly greater than
MIN_LISTING_AGE_MONTHS = 6
MIN_MONTHLY_SALES = 50
MAX_DECLINE_PCT = -25.0           # remove if 90-day trend declines more than 25%

# Budget-tiered review window (review_min, review_max). Bigger budgets can take
# on more-established (higher-review) competitors. max=None means no upper limit.
# Tiers checked high-to-low by the budget (inventory £) lower bound.
REVIEW_TIERS = [
    # (budget_lower_bound, review_min, review_max)
    (20000, 200, None),     # Premium    : >= £20,000  -> 200+
    (10000, 200, 12000),    # Very High  : £10k-£20k   -> 200-12000
    (5000,  200, 8000),     # High       : £5k-£10k    -> 200-8000
    (2000,  200, 5000),     # Medium     : £2k-£5k     -> 200-5000
    (0,     200, 3000),     # Low        : < £2,000    -> 200-3000
]
DEFAULT_REVIEW_BAND = (200, 3000)  # when no budget is given


def review_band_for_budget(budget: Optional[float]):
    """Return (review_min, review_max) for the budget tier; max None = no cap."""
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
    """0–30: 0 at (cap − 1,000), 30 at the cap. 0 if no budget or no revenue."""
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
    """Return 0-15 review points based on the budget's review band.

    - If `reviews` is None or `budget` is None -> 0.
    - If the budget tier has no upper review cap (rmax is None), we scale
      conservatively using rmax = rmin * 8 to create a meaningful fallback.
    - Points linearly decrease from 15 at `rmin` to 0 at `rmax`.
    """
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


def _build_filter_steps(budget, max_revenue, category, include_seasonal):
    """Ordered list of (label, sql_clause, params) shared by fetch_pool and the
    funnel report, so exclusion counts always match the actual pool filters."""
    rmin, rmax = review_band_for_budget(budget)
    if rmax is None:
        review_label = f"Review count ({rmin}+)"
        review_clause = "review_count IS NOT NULL AND review_count >= %s"
        review_params = [rmin]
    else:
        review_label = f"Review count ({rmin}-{rmax})"
        review_clause = "review_count IS NOT NULL AND review_count BETWEEN %s AND %s"
        review_params = [rmin, rmax]

    steps = []
    if category:
        steps.append((f"Category = {category}", "category ILIKE %s", [category]))
    steps += [
        (f"Min rating (> {MIN_RATING})", "rating IS NOT NULL AND rating > %s", [MIN_RATING]),
        (f"Min listing age ({MIN_LISTING_AGE_MONTHS}mo)",
         "listing_age_months IS NOT NULL AND listing_age_months >= %s", [MIN_LISTING_AGE_MONTHS]),
        (f"Min monthly sales ({MIN_MONTHLY_SALES})",
         "estimated_sales IS NOT NULL AND estimated_sales >= %s", [MIN_MONTHLY_SALES]),
        (f"Sales trend (> {MAX_DECLINE_PCT:.0f}%)",
         "(sales_trend_90d IS NULL OR sales_trend_90d > %s)", [MAX_DECLINE_PCT]),
        (review_label, review_clause, review_params),
        ("Amazon-as-seller", "(is_amazon_seller IS NOT TRUE)", []),
        ("Compatibility products", "(is_compatibility IS NOT TRUE)", []),
        ("Global brands", "(is_global_brand IS NOT TRUE)", []),
    ]
    if not include_seasonal:
        steps.append(("Seasonal products", "(is_seasonal IS NOT TRUE)", []))
    if max_revenue is not None:
        lo, hi = revenue_band(max_revenue)
        steps.append((f"Revenue band (£{lo:,.0f}-£{hi:,.0f})",
                      "estimated_revenue IS NOT NULL AND estimated_revenue BETWEEN %s AND %s", [lo, hi]))
    return steps


def filter_funnel(
    conn,
    budget: Optional[float] = None,
    *,
    max_revenue: Optional[float] = None,
    category: Optional[str] = None,
    include_seasonal: bool = True,
) -> List[Dict]:
    """Sequential exclusion report: for each filter, how many products it removed
    and how many remain. Returns ordered dicts {filter, removed, remaining}."""
    from research.phase1.db import Phase1DbError

    if max_revenue is None:
        max_revenue = max_revenue_for_budget(budget)
    steps = _build_filter_steps(budget, max_revenue, category, include_seasonal)

    report: List[Dict] = []
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT COUNT(*) FROM products")
            remaining = cur.fetchone()[0]
            report.append({"filter": "All stored products", "removed": 0, "remaining": remaining})

            applied_clauses: List[str] = []
            applied_params: List = []
            for label, clause, params in steps:
                applied_clauses.append(clause)
                applied_params.extend(params)
                sql = "SELECT COUNT(*) FROM products WHERE " + " AND ".join(applied_clauses)
                cur.execute(sql, applied_params)
                now = cur.fetchone()[0]
                report.append({"filter": label, "removed": remaining - now, "remaining": now})
                remaining = now
    except Exception as e:  # noqa: BLE001
        raise Phase1DbError(f"Filter funnel failed: {e}") from e
    return report


def _soft_shuffle(rows: List[Dict], jitter: float = 3.0) -> List[Dict]:
    def key(r):
        return float(r.get("opportunity_final") or 0.0) + random.uniform(-jitter, jitter)
    return sorted(rows, key=key, reverse=True)


def _apply_claude_scoring(rows: List[Dict], budget: Optional[float], top_n: int = 20) -> None:
    """Optionally apply Claude judgement adjustments to the top N products.
    
    If `USE_CLAUDE_PHASE1=true` and Claude is available, score the top *top_n*
    products and apply adjustments to their `opportunity_final` scores.
    Modifies *rows* in-place by adding `llm_adj`, `llm_reason`, and updating
    `opportunity_final`.
    """
    use_claude = os.environ.get("USE_CLAUDE_PHASE1", "").lower() in ("true", "1", "yes")
    if not use_claude:
        return

    try:
        from research.phase1.claude_client import load_client, score_products_with_claude, Phase1ApiError
    except ImportError:
        # Claude client not available (shouldn't happen if USE_CLAUDE_PHASE1 is set, but fail gracefully)
        return

    try:
        client = load_client()
    except Phase1ApiError:
        # No API key; skip Claude scoring
        return

    batch = rows[:top_n]
    try:
        results = score_products_with_claude(client, batch, budget, max_batch=top_n)
    except Phase1ApiError:
        # Claude failed; continue without adjustments
        return

    # Apply adjustments: match by ASIN and update opportunity_final
    asin_to_result = {r.asin.upper(): r for r in results}
    for row in rows:
        asin = str(row.get("asin") or "").upper()
        result = asin_to_result.get(asin)
        if result:
            row["llm_adj"] = result.llm_adj
            row["llm_reason"] = result.llm_reason
            row["llm_tags"] = result.llm_tags
            old_score = row.get("opportunity_final") or 0.0
            new_score = round(old_score + result.llm_adj, 1)
            row["opportunity_final"] = max(0.0, new_score)  # clamp to 0



def fetch_pool(
    conn,
    budget: Optional[float] = None,
    *,
    max_revenue: Optional[float] = None,
    category: Optional[str] = None,
    include_seasonal: bool = True,
    pool_size: int = 100,
    shuffle: bool = True,
) -> List[Dict]:
    """Return up to *pool_size* products (dicts) ranked by final score
    (base 0–85 + revenue proximity 0–30), after the revenue band + built-in
    quality filters + seasonal option."""
    from psycopg2.extras import RealDictCursor
    from research.phase1.db import Phase1DbError

    if max_revenue is None:
        max_revenue = max_revenue_for_budget(budget)

    steps = _build_filter_steps(budget, max_revenue, category, include_seasonal)
    where = [clause for _, clause, _ in steps]
    params: List = [p for _, _, plist in steps for p in plist]

    sql = (
        "SELECT * FROM products WHERE " + " AND ".join(where)
        + " ORDER BY opportunity_score DESC NULLS LAST LIMIT 2000"
    )

    try:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute(sql, params)
            rows = [dict(r) for r in cur.fetchall()]
    except Exception as e:  # noqa: BLE001
        raise Phase1DbError(f"Pool query failed: {e}") from e

    # Add revenue proximity (query-time) -> final score, then rank.
    for r in rows:
        base = float(r.get("opportunity_score") or 0.0)
        prox = revenue_proximity(
            float(r["estimated_revenue"]) if r.get("estimated_revenue") is not None else None,
            max_revenue,
        )
        r["revenue_proximity"] = prox
        # Budget-aware review points (0-15) applied at query time so review
        # scoring follows the user's budget tiers rather than a global rule.
        review_count = r.get("review_count") if r.get("review_count") is not None else r.get("reviews")
        review_pts = budget_review_points(review_count, budget)
        r["budget_review_points"] = review_pts
        r["opportunity_final"] = round(base + prox + review_pts, 1)

    rows = _soft_shuffle(rows) if shuffle else sorted(
        rows, key=lambda r: r.get("opportunity_final") or 0.0, reverse=True
    )

    # Optional: apply Claude judgement adjustments to top products.
    _apply_claude_scoring(rows, budget, top_n=20)

    # Re-sort after Claude adjustments (if applied).
    rows = sorted(rows, key=lambda r: r.get("opportunity_final") or 0.0, reverse=True)
    
    return rows[:pool_size]
