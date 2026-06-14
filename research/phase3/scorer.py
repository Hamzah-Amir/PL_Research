"""
Competitor rubric scoring — the modernised 184-point rubric (PDF Step 10C).

The rubric is the classic Just One Dime competitor sheet (max 159) extended with
four data-driven elements the user approved (+25 -> 184):

    Review Velocity (5) · Product Variations (5) · Sales/Revenue Strength (10) ·
    Enhanced Content A+/Video (5)

Direction (unchanged from the template): a HIGHER score = a STRONGER competitor
= harder for a new seller to enter. Bands (re-based to 184 at the original
47% / 73% cut points): 0-86 easy · 87-134 mid-challenge · 135-184 difficult.

Scoring split:
  - Deterministic (Python, here): price tier, sponsored products / sponsored
    brands, reviews (category-relative), rating, age, bestseller, seller
    feedback, FBA/FBM, review velocity, variations, sales/revenue strength,
    enhanced content.
  - Judgement (Claude, `claude_client.score_listing`): product images, title,
    bullets/description, unique design, marketing images, pricing strategy.

Anything that cannot be verified from a source is scored `None` (Unknown —
needs user confirmation, per Rule 0). `None` scores are NOT counted toward the
total; they are listed in `unknown_elements` so the caller can surface them and
the sheet can leave that cell blank for the user.
"""

from datetime import datetime
from typing import Dict, List, Optional

import numpy as np

MAX_TOTAL = 184
BANDS = [
    (0, 86, "Easy to compete — high chance of success for a new seller"),
    (87, 134, "Mid-range challenge — need a strong plan to beat weaknesses"),
    (135, 184, "Difficult — established competition, high chance of failure"),
]

# Element order + max points. Marketing title/bullets reuse the product-page
# judgement (same listing); pricing strategy is Claude's.
ELEMENTS = [
    # key, label, max, section
    ("price", "Price", 25, "PRODUCT PAGE"),
    ("sponsored_products", "Sponsored Products", 20, "PRODUCT PAGE"),
    ("sponsored_brands", "Sponsored Brands (ex-AMS)", 10, "PRODUCT PAGE"),
    ("reviews", "# of Product Reviews", 10, "PRODUCT PAGE"),
    ("rating", "Avg. Product Rating", 5, "PRODUCT PAGE"),
    ("age", "Age of Product", 5, "PRODUCT PAGE"),
    ("product_images", "Product Images", 20, "PRODUCT PAGE"),
    ("product_title", "Product Title", 5, "PRODUCT PAGE"),
    ("bullets_description", "Product Description/Bullet Points", 5, "PRODUCT PAGE"),
    ("bestseller", "Bestseller Badge", 10, "PRODUCT PAGE"),
    ("seller", "Amazon or 3P Seller", 5, "PRODUCT PAGE"),
    ("unique_design", "Unique Product Design", 2, "PRODUCT PAGE"),
    ("fba_fbm", "FBA or FBM", 2, "PRODUCT PAGE"),
    ("mkt_images", "Marketing — Product Images", 15, "MARKETING STRATEGY"),
    ("mkt_title", "Marketing — Product Title", 5, "MARKETING STRATEGY"),
    ("mkt_bullets", "Marketing — Description/Bullets", 5, "MARKETING STRATEGY"),
    ("pricing_strategy", "Pricing Strategy", 10, "MARKETING STRATEGY"),
    ("review_velocity", "Review Velocity", 5, "MARKET MOMENTUM"),
    ("variations", "Product Variations", 5, "MARKET MOMENTUM"),
    ("sales_strength", "Sales / Revenue Strength", 10, "MARKET MOMENTUM"),
    ("enhanced_content", "Enhanced Content (A+/Video)", 5, "MARKET MOMENTUM"),
]
ELEMENT_MAX = {k: m for k, _, m, _ in ELEMENTS}


def _num(v) -> Optional[float]:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


def _parse_date(value) -> Optional[datetime]:
    if value is None:
        return None
    s = str(value).strip()
    if not s or s.lower() in ("nan", "n/a", "-"):
        return None
    # JS / Xray numeric dates are US mm/dd/yyyy (e.g. 12/14/2025, 03/30/2023),
    # so try that before dd/mm/yyyy (which still catches true UK-format dates).
    for fmt in ("%m/%d/%Y", "%d/%m/%Y", "%b %d, %Y", "%B %d, %Y", "%Y-%m-%d", "%d %b %Y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def _percentile(value: Optional[float], pool: List[float]) -> Optional[float]:
    """Fraction of *pool* values <= *value* (0..1). None if no data."""
    vals = [v for v in pool if v is not None]
    if value is None or not vals:
        return None
    return sum(1 for v in vals if v <= value) / len(vals)


def _rel_low_is_strong(value, values) -> Optional[float]:
    """0..1 where the LOWEST value scores 1.0 (used for price)."""
    vals = [v for v in values if v is not None]
    if value is None or len(vals) < 2:
        return None if value is None else 0.5
    lo, hi = min(vals), max(vals)
    if hi == lo:
        return 0.5
    return 1.0 - (value - lo) / (hi - lo)


def _e(score, mx, reason, info=""):
    """Build an element result; score None = Unknown (not counted)."""
    if score is not None:
        score = max(0, min(int(round(score)), mx))
    return {"score": score, "max": mx, "reason": reason, "info": info}


def score_competitor(
    comp: Dict,
    peers: List[Dict],
    pool_stats: Dict,
    listing,
) -> Dict:
    """Score one competitor across all 21 elements.

    *comp* — merged record + Keepa deep fields + sponsored counts + review_velocity.
    *peers* — the selected competitor set (for relative price/sales among them).
    *pool_stats* — {'review_counts': [...], 'sales': [...]} from the same-type
        pool (for category-relative reviews / sales strength).
    *listing* — the Claude `ListingScores` for this competitor (judgement elements).
    """
    el: Dict[str, Dict] = {}
    peer_prices = [_num(p.get("price")) for p in peers]
    peer_sales = [_num(p.get("monthly_sales")) for p in peers]
    total_kw = comp.get("total_keywords") or 0

    # 1. Price (lower = stronger)
    pos = _rel_low_is_strong(_num(comp.get("price")), peer_prices)
    if pos is None:
        el["price"] = _e(None, 25, "No price available.", "Unknown — need user confirmation")
    else:
        el["price"] = _e(25 * pos, 25, f"Price £{comp.get('price')} vs peers (lower=stronger).",
                         f"£{comp.get('price')}")

    # 2. Sponsored Products (presence across launch keywords)
    spc = comp.get("sp_count") or 0
    if total_kw:
        el["sponsored_products"] = _e(20 * spc / total_kw, 20,
            f"Sponsored Products on {spc}/{total_kw} launch keywords.",
            f"SP on {spc}/{total_kw} kws")
    else:
        el["sponsored_products"] = _e(None, 20, "No Xray keyword data.", "Unknown — need user confirmation")

    # 3. Sponsored Brands (ex-AMS)
    sbc = comp.get("sb_count") or 0
    if total_kw:
        el["sponsored_brands"] = _e(10 * sbc / total_kw, 10,
            f"Sponsored Brand ads on {sbc}/{total_kw} launch keywords.",
            f"SB on {sbc}/{total_kw} kws")
    else:
        el["sponsored_brands"] = _e(None, 10, "No Xray keyword data.", "Unknown — need user confirmation")

    # 4. Reviews (category-relative percentile)
    rc = _num(comp.get("review_count"))
    pct = _percentile(rc, pool_stats.get("review_counts", []))
    if pct is None:
        el["reviews"] = _e(None, 10, "No review count.", "Unknown — need user confirmation")
    else:
        el["reviews"] = _e(10 * pct, 10,
            f"{int(rc):,} reviews — {pct*100:.0f}th percentile in category.", f"{int(rc):,}")

    # 5. Rating
    rt = _num(comp.get("rating"))
    if rt is None:
        el["rating"] = _e(None, 5, "No rating.", "Unknown — need user confirmation")
    else:
        sc = 5 if rt >= 4.7 else 4 if rt >= 4.5 else 3 if rt >= 4.3 else 2 if rt >= 4.0 else 1 if rt >= 3.5 else 0
        el["rating"] = _e(sc, 5, f"Rating {rt}.", f"{rt}")

    # 6. Age of product (Date First Available / Creation Date)
    dt = _parse_date(comp.get("date_first_available"))
    if dt is None:
        el["age"] = _e(None, 5, "No reliable date first available.", "Unknown — need user confirmation")
    else:
        yrs = (datetime.now() - dt).days / 365.25
        sc = 5 if yrs >= 3 else 4 if yrs >= 2 else 3 if yrs >= 1.5 else 2 if yrs >= 1 else 1 if yrs >= 0.5 else 0
        el["age"] = _e(sc, 5, f"~{yrs:.1f} years old (older=more established).", dt.strftime("%Y-%m-%d"))

    # 7-9, 12, 14-17. Judgement elements from Claude
    def _claude(attr, mx, reason_attr):
        if listing is None:
            return _e(None, mx, "Claude listing scoring unavailable.", "Unknown — need user confirmation")
        val = getattr(listing, attr, None)
        if val is None or val < 0:  # -1 sentinel = cannot judge
            return _e(None, mx, getattr(listing, reason_attr, ""), "Unknown — need user confirmation")
        return _e(val, mx, getattr(listing, reason_attr, ""))

    el["product_images"] = _claude("product_images", 20, "product_images_reason")
    el["product_title"] = _claude("product_title", 5, "product_title_reason")
    el["bullets_description"] = _claude("bullets_description", 5, "bullets_description_reason")
    el["unique_design"] = _claude("unique_design", 2, "unique_design_reason")
    el["mkt_images"] = _claude("marketing_images", 15, "marketing_images_reason")
    # Marketing title/bullets reuse the product-page judgement (same listing).
    el["mkt_title"] = _claude("product_title", 5, "product_title_reason")
    el["mkt_bullets"] = _claude("bullets_description", 5, "bullets_description_reason")
    el["pricing_strategy"] = _claude("pricing_strategy", 10, "pricing_strategy_reason")

    # 10. Bestseller badge (flag + BSR proxy)
    bs = str(comp.get("best_seller") or "").strip().lower()
    bsr = _num(comp.get("bsr"))
    if bs in ("yes", "true", "1"):
        el["bestseller"] = _e(9, 10, "Holds a Best Seller badge.", "Best Seller: Yes")
    elif bsr is not None:
        sc = 5 if bsr <= 20 else 3 if bsr <= 100 else 2 if bsr <= 500 else 1
        el["bestseller"] = _e(sc, 10, f"No badge; BSR #{int(bsr)} (rank proxy).", f"No badge / BSR #{int(bsr)}")
    else:
        el["bestseller"] = _e(None, 10, "No badge/BSR data.", "Unknown — need user confirmation")

    # 11. Amazon or 3P seller (all-time feedback, deep pull)
    fb = comp.get("seller_feedback") or {}
    fb_count = _num(fb.get("rating_count"))
    if fb_count is not None:
        sc = 5 if fb_count >= 10000 else 4 if fb_count >= 2000 else 3 if fb_count >= 500 else 2 if fb_count >= 100 else 1
        el["seller"] = _e(sc, 5, f"Seller all-time feedback {int(fb_count):,}.", f"{int(fb_count):,} feedback")
    else:
        el["seller"] = _e(None, 5, "All-time seller feedback unavailable (needs offers pull / 3P buy box).",
                          "Unknown — need user confirmation")

    # 13. FBA or FBM
    ff = str(comp.get("fulfillment") or "").upper()
    if ff in ("FBA", "AMZ"):
        el["fba_fbm"] = _e(2, 2, "Fulfilled by Amazon (FBA/AMZ).", ff)
    elif ff == "FBM":
        el["fba_fbm"] = _e(0, 2, "Fulfilled by Merchant.", "FBM")
    else:
        el["fba_fbm"] = _e(None, 2, "Unknown fulfilment.", "Unknown — need user confirmation")

    # 18. Review velocity
    rv = _num(comp.get("review_velocity"))
    if rv is None:
        el["review_velocity"] = _e(None, 5, "No review-velocity data.", "Unknown — need user confirmation")
    else:
        sc = 5 if rv >= 30 else 3 if rv >= 10 else 1 if rv > 0 else 0
        el["review_velocity"] = _e(sc, 5, f"~{rv:.0f} reviews/month.", f"{rv:.0f}/mo")

    # 19. Product variations
    vc = _num(comp.get("variations_count"))
    if vc is None:
        el["variations"] = _e(None, 5, "No variation data.", "Unknown — need user confirmation")
    else:
        sc = 5 if vc >= 5 else 3 if vc >= 2 else 1 if vc >= 1 else 0
        el["variations"] = _e(sc, 5, f"{int(vc)} variation(s).", f"{int(vc)}")

    # 20. Sales / revenue strength (category-relative)
    sp = _percentile(_num(comp.get("monthly_sales")), pool_stats.get("sales", []))
    if sp is None:
        el["sales_strength"] = _e(None, 10, "No sales data.", "Unknown — need user confirmation")
    else:
        ms = _num(comp.get("monthly_sales"))
        el["sales_strength"] = _e(10 * sp, 10,
            f"{int(ms):,} units/mo — {sp*100:.0f}th percentile in category.", f"{int(ms):,}/mo")

    # 21. Enhanced content (A+ / video / image count)
    ap = comp.get("a_plus")
    vid = comp.get("has_video")
    imgs = _num(comp.get("images_count"))
    if ap is None and vid is None and imgs is None:
        el["enhanced_content"] = _e(None, 5, "No content data.", "Unknown — need user confirmation")
    else:
        if ap and vid:
            sc, why = 5, "A+ content and video present."
        elif ap or vid:
            sc, why = 4, "A+ content or video present."
        elif imgs is not None and imgs >= 6:
            sc, why = 3, f"{int(imgs)} images, no A+/video detected."
        else:
            sc, why = 1, "Minimal media."
        el["enhanced_content"] = _e(sc, 5, why, f"A+={ap} video={vid} imgs={imgs}")

    # Totals
    total = sum(v["score"] for v in el.values() if v["score"] is not None)
    unknown = [k for k, v in el.items() if v["score"] is None]
    band = next((d for lo, hi, d in BANDS if lo <= total <= hi), "")

    return {
        "asin": comp.get("asin"),
        "elements": el,
        "total": total,
        "max": MAX_TOTAL,
        "band": band,
        "unknown_elements": unknown,
    }


def pool_statistics(same_type_records) -> Dict:
    """Build {'review_counts','sales'} distributions from the same-type pool
    (a DataFrame or list of dicts) for category-relative scoring."""
    def _col(name):
        if hasattr(same_type_records, "columns"):
            return [_num(v) for v in same_type_records.get(name, [])]
        return [_num(r.get(name)) for r in same_type_records]
    return {
        "review_counts": [v for v in _col("review_count") if v is not None],
        "sales": [v for v in _col("monthly_sales") if v is not None],
    }
