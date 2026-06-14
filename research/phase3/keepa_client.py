"""
Keepa API client for Phase 3 competitor analysis.

Replaces the manual Keepa "Product Viewer" export with live API extraction
(key in `.env` as KEEPA_API_KEY). Hits the **Product Request** endpoint via the
`keepa` library's `query()` — 1 token per ASIN, batched up to 100 per call,
domain `GB` (Amazon UK).

Token discipline (the user's balance is small, ~60):
  - SHALLOW pool pull — `history=1, stats=180`, nothing else → ~1 token/ASIN.
    `stats` and `history` add no extra token cost and give current price, BSR,
    variations, bullets, images, FBA fee, sub-category, coupon, dimensions.
  - DEEP pull (final 3 competitors only) — adds `offers`, `buybox`, `rating`,
    `aplus`, `videos` for the all-time seller feedback (via `seller_query`),
    A+ flag and video presence. `offers` is the expensive part (6 tokens per
    offer page), so it is reserved for the 3 selected winners.

Every response is cached to disk per (ASIN, depth) so re-runs during development
cost ZERO tokens. RATING/COUNT_REVIEWS are intentionally NOT requested in the
shallow pull (Xray already provides them — saves 1 token/ASIN).

No heuristic fallback: if the API or key is unavailable, Phase 3 surfaces the
error (Rule 0) rather than guessing.
"""

import os
import pickle
import re
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd

_ROOT = Path(__file__).resolve().parent.parent
_ENV_PATH = _ROOT / ".env"
_CACHE_DIR = _ROOT / "output" / "keepa_cache"

DOMAIN = "GB"  # Amazon UK
_ASIN_RE = re.compile(r"^B0[0-9A-Z]{8}$")
_IMG_BASE = "https://m.media-amazon.com/images/I/"


class Phase3KeepaError(RuntimeError):
    """Raised when the Keepa API cannot complete a Phase 3 request."""


# ──────────────────────────────────────────────────────────────────────────────
# ASIN pool
# ──────────────────────────────────────────────────────────────────────────────

def build_asin_pool(
    js_df: Optional[pd.DataFrame],
    xray_dfs: Optional[List[pd.DataFrame]],
) -> List[str]:
    """Deduplicated, order-preserving parent-ASIN pool from JS + all Xray files.

    JS ASINs come first (the workflow's stated source), then any Xray ASINs not
    already present — so a strong competitor surfaced only by Xray is still
    covered. Parents only (variation children come from Keepa later, Phase 4).
    """
    pool: List[str] = []
    seen = set()

    def _add(series):
        for v in series:
            if v is None:
                continue
            s = str(v).strip().upper()
            if _ASIN_RE.match(s) and s not in seen:
                seen.add(s)
                pool.append(s)

    if js_df is not None and "asin" in js_df.columns:
        _add(js_df["asin"])
    for df in (xray_dfs or []):
        if df is not None and "asin" in df.columns:
            _add(df["asin"])
    return pool


# ──────────────────────────────────────────────────────────────────────────────
# Client + disk cache
# ──────────────────────────────────────────────────────────────────────────────

def load_keepa():
    """Build a keepa.Keepa client from the .env KEEPA_API_KEY. No fallback."""
    try:
        from dotenv import load_dotenv
    except ImportError as e:
        raise Phase3KeepaError(
            "python-dotenv is not installed. Run setup again."
        ) from e
    try:
        import keepa  # noqa: F401
    except ImportError as e:
        raise Phase3KeepaError(
            "The 'keepa' package is not installed. Run: pip install keepa."
        ) from e

    load_dotenv(dotenv_path=_ENV_PATH)
    key = os.environ.get("KEEPA_API_KEY")
    if not key:
        raise Phase3KeepaError(
            f"KEEPA_API_KEY not found. Add it to {_ENV_PATH} "
            f"(KEEPA_API_KEY=...). The file is git-ignored."
        )
    import keepa
    try:
        return keepa.Keepa(key)
    except Exception as e:  # noqa: BLE001
        raise Phase3KeepaError(f"Could not initialise the Keepa client: {e}") from e


def _cache_path(depth: str) -> Path:
    _CACHE_DIR.mkdir(parents=True, exist_ok=True)
    return _CACHE_DIR / f"products_{depth}.pkl"


def _load_cache(depth: str) -> Dict[str, dict]:
    path = _cache_path(depth)
    if path.exists():
        try:
            with open(path, "rb") as fh:
                return pickle.load(fh)
        except Exception:  # noqa: BLE001 — corrupt cache should not be fatal
            return {}
    return {}


def _save_cache(depth: str, cache: Dict[str, dict]) -> None:
    with open(_cache_path(depth), "wb") as fh:
        pickle.dump(cache, fh)


# ──────────────────────────────────────────────────────────────────────────────
# Product query (token-aware, cached)
# ──────────────────────────────────────────────────────────────────────────────

def query_products(
    api,
    asins: List[str],
    *,
    deep: bool = False,
    use_cache: bool = True,
    limit: Optional[int] = None,
) -> Dict[str, dict]:
    """
    Query Keepa for *asins* and return ``{asin: raw_product_dict}``.

    *deep* selects the richer (more expensive) parameter set for the final 3
    competitors. *limit* caps how many *uncached* ASINs are sent to the API in
    this call (token guard — e.g. ``limit=25``). Cached ASINs are always
    returned and never count against the limit.

    Raises Phase3KeepaError on API failure (no fallback).
    """
    depth = "deep" if deep else "shallow"
    cache = _load_cache(depth) if use_cache else {}

    wanted = [a for a in asins if _ASIN_RE.match(a or "")]
    have = {a: cache[a] for a in wanted if a in cache}
    missing = [a for a in wanted if a not in cache]

    if limit is not None and len(missing) > limit:
        missing = missing[:limit]

    if missing:
        kwargs = dict(domain=DOMAIN, history=True, stats=180, progress_bar=False)
        if deep:
            kwargs.update(offers=20, buybox=True, rating=True, aplus=True, videos=True)
        try:
            products = api.query(missing, **kwargs)
        except Exception as e:  # noqa: BLE001
            raise Phase3KeepaError(f"Keepa product query failed: {e}") from e

        for prod in products:
            asin = prod.get("asin")
            if asin:
                cache[asin] = prod
                have[asin] = prod
        if use_cache:
            _save_cache(depth, cache)

    # Preserve the requested order.
    return {a: have[a] for a in wanted if a in have}


def get_seller_feedback(api, seller_id: str) -> Optional[Dict]:
    """All-time seller feedback for *seller_id* via the Keepa Seller endpoint.

    Returns ``{"name", "rating_count", "rating_pct"}`` or None. Used only for
    the final 3 competitors (per the user's all-time-feedback decision).
    """
    if not seller_id or seller_id in ("-1", "-2"):
        return None
    try:
        result = api.seller_query(seller_id, domain=DOMAIN)
    except Exception:  # noqa: BLE001 — seller data is best-effort
        return None
    rec = result.get(seller_id) if isinstance(result, dict) else None
    if not rec:
        return None
    return {
        "name": rec.get("sellerName"),
        # Keepa exposes lifetime counts as currentRatingCount / current rating %.
        "rating_count": rec.get("currentRatingCount") or rec.get("ratingCount"),
        "rating_pct": rec.get("currentRating") or rec.get("rating"),
    }


# ──────────────────────────────────────────────────────────────────────────────
# Field extraction (Keepa-only canonical fields)
# ──────────────────────────────────────────────────────────────────────────────

def _img_urls(product: dict, limit: int = 9) -> List[str]:
    out = []
    for im in (product.get("images") or [])[:limit]:
        name = im.get("l") or im.get("m") if isinstance(im, dict) else None
        if name:
            out.append(_IMG_BASE + name)
    return out


def _listed_since(product: dict) -> Optional[str]:
    raw = product.get("listedSince")
    if not raw or raw <= 0:
        return None
    try:
        import keepa
        return keepa.keepa_minutes_to_time(raw).strftime("%Y-%m-%d")
    except Exception:  # noqa: BLE001
        return None


def extract_keepa_fields(product: dict) -> dict:
    """Pull the canonical Keepa fields Phase 3 needs from a raw product dict.

    Everything here is available in the cheap shallow pull except `a_plus`,
    `has_video`, `buybox_seller_id` and `seller_feedback`, which are populated
    only on the deep (final-3) pull.
    """
    stats = product.get("stats_parsed") or {}
    current = stats.get("current") or {}
    fba = product.get("fbaFees") or {}
    cat_tree = product.get("categoryTree") or []

    pick_pack = fba.get("pickAndPackFee")
    pick_pack = round(pick_pack / 100.0, 2) if isinstance(pick_pack, (int, float)) else None

    variations = product.get("variations") or []

    # A+ / video flags (deep pull). Keepa exposes A+ under aPlusContent when
    # aplus=1 is requested; treat any non-empty value as present.
    aplus_val = product.get("aPlusContent") or product.get("aplus")
    a_plus = bool(aplus_val) if aplus_val not in (None, "", [], {}) else None
    videos = product.get("videos")
    has_video = bool(videos) if videos not in (None, "", [], {}) else None

    bb_hist = product.get("buyBoxSellerIdHistory") or []
    buybox_seller_id = bb_hist[-1] if bb_hist else None

    return {
        "asin": product.get("asin"),
        "parent_asin": product.get("parentAsin"),
        "title": product.get("title"),
        "brand": product.get("brand") or product.get("manufacturer"),
        "price": current.get("NEW") if current.get("NEW", -1) not in (-1, None) else current.get("AMAZON"),
        "list_price": current.get("LISTPRICE") if current.get("LISTPRICE", -1) not in (-1, None) else None,
        "bsr": current.get("SALES") if current.get("SALES", -1) not in (-1, None) else None,
        "bsr_90d": (stats.get("avg90") or {}).get("SALES")
                   if (stats.get("avg90") or {}).get("SALES", -1) not in (-1, None) else None,
        "bullets": list(product.get("features") or []),
        "description": product.get("description"),
        "image_urls": _img_urls(product),
        "images_count": len(product.get("images") or []),
        "variations_count": len(variations),
        "variation_attrs": variations,
        "fba_pick_pack_fee": pick_pack,
        "referral_pct": product.get("referralFeePercent") or product.get("referralFeePercentage"),
        "category": cat_tree[0]["name"] if cat_tree else None,
        "subcategory": cat_tree[-1]["name"] if cat_tree else None,
        "coupon": product.get("coupon"),
        "monthly_sold": product.get("monthlySold"),
        "brand_store_url": product.get("brandStoreUrl"),
        "listed_since": _listed_since(product),
        "item_weight_g": product.get("itemWeight"),
        "package_weight_g": product.get("packageWeight"),
        # deep-only:
        "a_plus": a_plus,
        "has_video": has_video,
        "buybox_seller_id": buybox_seller_id,
        "seller_feedback": None,  # filled by main after get_seller_feedback()
    }
