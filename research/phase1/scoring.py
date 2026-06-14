"""
Phase 1 scoring + segmentation + exclusion flags (pure functions — no DB).

Scoring is the user's point-based formula (~115 max). The **revenue-proximity
factor (0-30) depends on the user's budget**, so it is computed at QUERY time
(see `query.revenue_proximity`). Everything else — the **base score (0-85)** —
is product-intrinsic and computed/stored at ingest:

    90-day trend (20)  +1 per 50% growth, capped 20
    YoY growth   (10)  +1 per 5%  growth, capped 10
    review count (15)  <200=15 · <500=10 · <1000=5
    rating       (10)  >=4.5=10 · >=4.2=7 · >=4.0=4
    weight       (15)  <0.3kg=15 · <0.5kg=10 · <1.0kg=5   (FBA cost proxy)
    price £8-£20 (10)  full 10 in range
    not Amazon   ( 5)  +5 if seller is not Amazon
    variation    (-5)  penalty if >10 variations
                      ----  base_score (stored as opportunity_score)
    revenue prox (30)  added at query: linear, 0 at (cap-1000) -> 30 at cap

Also computes segmentation buckets, `is_seasonal`, and the three exclusion
flags `is_amazon_seller` / `is_compatibility` / `is_global_brand`.
"""

import json
import re
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd

_ROOT = Path(__file__).resolve().parent.parent
_BLOCKLIST_PATH = Path(__file__).resolve().parent / "brand_blocklist.json"

# ── Segmentation thresholds ─────────────────────────────────────────────────
PRICE_LOW_MAX, PRICE_MID_MAX = 15.0, 50.0
REVIEW_LOW_MAX, REVIEW_MED_MAX = 100, 1000
SALES_LOW_MAX, SALES_MED_MAX = 100, 1000

# ── Scoring constants ───────────────────────────────────────────────────────
PRICE_SWEET_MIN, PRICE_SWEET_MAX = 8.0, 20.0
VARIATION_PENALTY_THRESHOLD = 10
REVENUE_PROXIMITY_MAX = 30.0   # used at query time

_SEASONAL_PEAK_MONTHS = {11, 12}
_SEASONAL_KEYWORDS = (
    "christmas", "xmas", "halloween", "valentine", "easter", "advent",
    "santa", "festive", "holiday", "thanksgiving", "summer", "winter",
    "gift set", "stocking",
)

# Compatibility / device-dependent product cues (old F9), title-based.
_COMPAT_KEYWORDS = (
    "compatible with", "replacement for", "replacement filter", "replacement part",
    "spare for", "for use with", "fits ", "case for", "cover for", "sleeve for",
    "strap for", "watch band", "watchband", "screen protector", "adapter for",
    "designed for", "suitable for",
)


def _load_blocklist() -> set:
    try:
        data = json.loads(_BLOCKLIST_PATH.read_text(encoding="utf-8"))
        return {str(b).strip().lower() for b in data.get("global_brands", []) if str(b).strip()}
    except Exception:  # noqa: BLE001 — missing/bad file -> no brand exclusions
        return set()


_BLOCKLIST = _load_blocklist()


def _num(v) -> Optional[float]:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


# ── Segmentation ────────────────────────────────────────────────────────────

def price_range(price):
    p = _num(price)
    if p is None or p < PRICE_LOW_MAX:
        return "low"
    return "mid" if p <= PRICE_MID_MAX else "high"


def review_bucket(reviews):
    r = _num(reviews) or 0
    return "low" if r < REVIEW_LOW_MAX else ("medium" if r <= REVIEW_MED_MAX else "high")


def sales_bucket(sales):
    s = _num(sales) or 0
    return "low" if s < SALES_LOW_MAX else ("medium" if s <= SALES_MED_MAX else "high")


def is_seasonal(best_sales_period, title="") -> bool:
    t = str(title or "").lower()
    if any(k in t for k in _SEASONAL_KEYWORDS):
        return True
    m = re.match(r"^(\d{1,2})[/\-]", str(best_sales_period or "").strip())
    return bool(m) and int(m.group(1)) in _SEASONAL_PEAK_MONTHS


def parse_weight_kg(value) -> Optional[float]:
    """Parse a weight cell to kilograms. Assumes kg unless 'g' (not 'kg') is present."""
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().lower()
    m = re.search(r"(\d+(?:\.\d+)?)", s)
    if not m:
        return None
    val = float(m.group(1))
    if "kg" in s:
        return val
    if re.search(r"\bg\b|gram|\bg$|\bg ", s) or s.endswith("g"):
        return val / 1000.0
    return val  # default kg


def is_amazon_seller(seller, fulfillment) -> bool:
    if "amazon" in str(seller or "").lower():
        return True
    return str(fulfillment or "").strip().upper() in ("AMZ", "AMAZON")


def is_compatibility(title) -> bool:
    t = str(title or "").lower()
    return any(k in t for k in _COMPAT_KEYWORDS)


def is_global_brand(brand) -> bool:
    return str(brand or "").strip().lower() in _BLOCKLIST


# ── Scoring (base, 0-85; revenue proximity added at query) ───────────────────

def _trend_points(trend_pct) -> float:
    t = _num(trend_pct)
    return 0.0 if t is None else max(0.0, min(20.0, t / 50.0))


def _yoy_points(yoy_pct) -> float:
    y = _num(yoy_pct)
    return 0.0 if y is None else max(0.0, min(10.0, y / 5.0))



def _rating_points(rating) -> float:
    r = _num(rating)
    if r is None:
        return 0.0
    return 10.0 if r >= 4.5 else 7.0 if r >= 4.2 else 4.0 if r >= 4.0 else 0.0


def _weight_points(weight_kg) -> float:
    w = _num(weight_kg)
    if w is None:
        return 0.0
    return 15.0 if w < 0.3 else 10.0 if w < 0.5 else 5.0 if w < 1.0 else 0.0


def _price_points(price) -> float:
    p = _num(price)
    return 10.0 if p is not None and PRICE_SWEET_MIN <= p <= PRICE_SWEET_MAX else 0.0


def base_score(row: dict) -> float:
    """The 0-85 product-intrinsic score (everything except revenue proximity)."""
    pts = (
        _trend_points(row.get("sales_trend_90d"))
        + _yoy_points(row.get("yoy_growth"))
        # Review points are now budget-dependent and applied at query time
        + _rating_points(row.get("rating"))
        + _weight_points(row.get("weight_kg"))
        + _price_points(row.get("price"))
        + (0.0 if row.get("is_amazon_seller") else 5.0)
    )
    if (_num(row.get("variation_count")) or 0) > VARIATION_PENALTY_THRESHOLD:
        pts -= 5.0
    return round(pts, 1)


def compute_scores(df: pd.DataFrame) -> pd.DataFrame:
    """Add base score, segmentation, seasonal, and exclusion flags to a frame."""
    out = df.copy()
    n = len(out)

    def col(name, default=None):
        return out[name] if name in out.columns else pd.Series([default] * n)

    out["weight_kg"] = col("weight").map(parse_weight_kg)
    out["sales_trend_90d"] = pd.to_numeric(col("sales_trend"), errors="coerce")
    out["yoy_growth"] = pd.to_numeric(col("yoy_growth"), errors="coerce")

    out["is_amazon_seller"] = [
        is_amazon_seller(s, f) for s, f in zip(col("seller"), col("fulfillment"))
    ]
    out["is_compatibility"] = col("title").map(is_compatibility)
    out["is_global_brand"] = col("brand").map(is_global_brand)

    vc = pd.to_numeric(col("variations"), errors="coerce").fillna(0)
    out["variation_count"] = vc.astype(int)
    out["has_variants"] = vc > 0

    out["price_range"] = col("price").map(price_range)
    out["review_bucket"] = col("reviews").map(review_bucket)
    out["sales_bucket"] = col("asin_sales").map(sales_bucket)
    out["is_seasonal"] = [is_seasonal(bp, t) for bp, t in zip(col("best_sales_period"), col("title"))]

    out["opportunity_score"] = [
        base_score({
            "sales_trend_90d": r.get("sales_trend_90d"),
            "yoy_growth": r.get("yoy_growth"),
            "review_count": r.get("reviews"),
            "rating": r.get("rating"),
            "weight_kg": r.get("weight_kg"),
            "price": r.get("price"),
            "is_amazon_seller": r.get("is_amazon_seller"),
            "variation_count": r.get("variation_count"),
        })
        for _, r in out.iterrows()
    ]
    return out
