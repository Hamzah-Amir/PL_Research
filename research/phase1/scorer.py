"""
Product scoring and ranking for Phase 1.

Selection:
  Score EVERY filtered product with a composite score (no pre-cap), rank them
  all, and return the best N (default 10).

Scoring pipeline (see module constants for the exact brackets):
  1. Log-normalise revenue and sales across the pool (not linear min-max), so
     products with similar revenue do not collapse to identical scores.
  2. Opportunity = sales / (reviews + 1), then min-max normalised.
  3. Rating mapped 3.8→0, 5.0→1 (clamped).
  4. Review-count score from a hard bracket (fewer reviews = better PL gap).
  5. Weighted base composite × 100.
  6. Seller-concentration multiplier (fewer active sellers = better).
  7. Trend multiplier (rising sales = better).
  8. Final clamp at 100.
"""

import math
from typing import Optional

import numpy as np
import pandas as pd

# ──────────────────────────────────────────────────────────────────────────────
# Composite score weights  (must sum to 1.0)
# ──────────────────────────────────────────────────────────────────────────────
WEIGHTS = {
    "revenue":     0.35,   # Monthly revenue potential
    "opportunity": 0.25,   # High sales relative to reviews (market gap signal)
    "sales":       0.20,   # Monthly sales volume
    "rating":      0.12,   # Star-rating quality
    "review_spot": 0.08,   # Review-count bracket (proven but not saturated)
}


def _log_norm(series: pd.Series) -> pd.Series:
    """
    Log-normalise a series to [0, 1].

    log_value   = log(1 + raw_value)
    normalised  = (log_value - min) / (max - min)

    Returns 0.5 for every row if all log-values are equal.
    """
    vals = pd.to_numeric(series, errors="coerce").fillna(0).clip(lower=0)
    log_vals = np.log1p(vals)
    mn, mx = log_vals.min(), log_vals.max()
    if mx == mn:
        return pd.Series([0.5] * len(series), index=series.index)
    return (log_vals - mn) / (mx - mn)


def _norm(series: pd.Series) -> pd.Series:
    """Min-max normalise a series to [0, 1]. Returns 0.5 if all values equal."""
    mn, mx = series.min(), series.max()
    if mx == mn:
        return pd.Series([0.5] * len(series), index=series.index)
    return (series - mn) / (mx - mn)


def _rating_score(rating: Optional[float]) -> float:
    """Map a star rating to [0, 1] where 5.0 → 1.0 and 3.8 → 0.0."""
    if rating is None or math.isnan(rating):
        return 0.0
    floor = 3.8
    ceiling = 5.0
    return max(0.0, min(1.0, (rating - floor) / (ceiling - floor)))


def _review_bracket(reviews: Optional[float]) -> float:
    """
    Hard review-count bracket — fewer reviews score higher (bigger PL gap).

      <= 50        → 1.00
      51 – 150     → 0.80
      151 – 300    → 0.60
      301 – 600    → 0.35
      601 – 1000   → 0.15
      > 1000       → 0.05
      missing/zero → 0.10
    """
    if reviews is None or math.isnan(reviews) or reviews <= 0:
        return 0.10
    r = float(reviews)
    if r <= 50:
        return 1.00
    if r <= 150:
        return 0.80
    if r <= 300:
        return 0.60
    if r <= 600:
        return 0.35
    if r <= 1000:
        return 0.15
    return 0.05


def _opportunity_score(sales: Optional[float], reviews: Optional[float]) -> float:
    """
    High-sales / low-review ratio signals under-served demand.
    Raw value is sales / (reviews + 1); normalised across the pool by the caller.
    """
    if sales is None or math.isnan(sales):
        return 0.0
    rev = reviews if (reviews is not None and not math.isnan(reviews)) else 0.0
    return float(sales) / (rev + 1.0)


def _seller_multiplier(active_sellers: Optional[float]) -> float:
    """
    Seller-concentration multiplier — fewer active sellers = more room.

      1     → 1.12
      2 – 3 → 1.06
      4 – 6 → 1.00
      > 6   → 0.92
      missing → 1.00
    """
    if active_sellers is None or math.isnan(active_sellers):
        return 1.00
    n = float(active_sellers)
    if n <= 0:
        return 1.00
    if n == 1:
        return 1.12
    if n <= 3:
        return 1.06
    if n <= 6:
        return 1.00
    return 0.92


def _trend_multiplier(trend: Optional[float]) -> float:
    """
    90-day sales-trend multiplier — rising sales = better.

      >= +50%       → 1.10
      +20% .. +49%  → 1.05
      -10% .. +19%  → 1.00
      -30% .. -11%  → 0.90
      < -30%        → 0.75
      missing       → 1.00
    """
    if trend is None or math.isnan(trend):
        return 1.00
    t = float(trend)
    if t >= 50:
        return 1.10
    if t >= 20:
        return 1.05
    if t >= -10:
        return 1.00
    if t >= -30:
        return 0.90
    return 0.75


def _build_risk_flags(row: pd.Series) -> str:
    """Return a comma-separated string of risk flags for a product row."""
    flags = []

    if row.get("seasonal_flag"):
        reason = row.get("seasonal_reason")
        if reason and str(reason).strip():
            flags.append(f"Seasonal — {reason}")
        else:
            flags.append("Seasonal")

    reviews = row.get("reviews")
    if reviews is not None and not math.isnan(float(reviews)):
        if float(reviews) < 50:
            flags.append("Low reviews (<50)")
        elif float(reviews) > 1000:
            flags.append("High reviews (>1000)")

    rating = row.get("rating")
    if rating is not None and not math.isnan(float(rating)):
        if float(rating) < 4.0:
            flags.append("Low rating (<4.0)")

    age = row.get("listing_age_months")
    if age is not None and not math.isnan(float(age)):
        if float(age) < 12:
            flags.append("Young listing (<12 mo)")

    trend = row.get("sales_trend")
    if trend is not None and not math.isnan(float(trend)):
        if float(trend) < -10:
            flags.append(f"Declining trend ({trend:+.0f}%)")

    fba = str(row.get("fulfillment", "")).upper()
    if "FBM" in fba:
        flags.append("FBM seller")

    sellers = row.get("active_sellers")
    if sellers is not None and not math.isnan(float(sellers)) and float(sellers) == 1:
        flags.append("Single seller")

    return ", ".join(flags) if flags else "None"


def score_all_products(df: pd.DataFrame) -> pd.DataFrame:
    """
    Compute a composite score for EVERY filtered product (no pre-cap) and
    return them all ranked by score descending.

    Adds columns: composite_score (0–100), risk_flags.
    """
    if df.empty:
        return df.copy()

    sales_col = "monthly_sales"
    rev_col = "monthly_revenue"

    # Score the entire filtered set — sub-scores are normalised across all of it.
    pool = df.copy().reset_index(drop=True)

    # ── STEP 1: revenue & sales sub-scores (LOG normalisation) ────────────────
    if rev_col in pool.columns:
        pool["_s_revenue"] = _log_norm(pool[rev_col])
    else:
        pool["_s_revenue"] = 0.5

    if sales_col in pool.columns:
        pool["_s_sales"] = _log_norm(pool[sales_col])
    else:
        pool["_s_sales"] = 0.5

    # ── STEP 1: rating sub-score ──────────────────────────────────────────────
    if "rating" in pool.columns:
        pool["_s_rating"] = pool["rating"].apply(
            lambda v: _rating_score(float(v) if pd.notna(v) else None)
        )
    else:
        pool["_s_rating"] = 0.5

    # ── STEP 2: review-bracket sub-score ──────────────────────────────────────
    if "reviews" in pool.columns:
        pool["_s_review"] = pool["reviews"].apply(
            lambda v: _review_bracket(float(v) if pd.notna(v) else None)
        )
    else:
        pool["_s_review"] = 0.10

    # ── STEP 1: opportunity sub-score (raw ratio, then min-max) ───────────────
    sales_series = pool.get(sales_col, pd.Series([0.0] * len(pool)))
    review_series = pool.get("reviews", pd.Series([0.0] * len(pool)))

    raw_opp = pd.Series(
        [
            _opportunity_score(
                float(s) if pd.notna(s) else None,
                float(r) if pd.notna(r) else None,
            )
            for s, r in zip(sales_series, review_series)
        ],
        index=pool.index,
    )
    pool["_s_opp"] = _norm(raw_opp)

    # ── STEP 3: weighted base composite (0–100) ───────────────────────────────
    base = (
        pool["_s_revenue"] * WEIGHTS["revenue"]
        + pool["_s_opp"]    * WEIGHTS["opportunity"]
        + pool["_s_sales"]  * WEIGHTS["sales"]
        + pool["_s_rating"] * WEIGHTS["rating"]
        + pool["_s_review"] * WEIGHTS["review_spot"]
    ) * 100

    # ── STEP 4: seller-concentration multiplier ───────────────────────────────
    seller_series = pool.get("active_sellers", pd.Series([np.nan] * len(pool)))
    seller_mult = pd.Series(
        [
            _seller_multiplier(float(s) if pd.notna(s) else None)
            for s in seller_series
        ],
        index=pool.index,
    )

    # ── STEP 5: trend multiplier ──────────────────────────────────────────────
    trend_series = pool.get("sales_trend", pd.Series([np.nan] * len(pool)))
    trend_mult = pd.Series(
        [
            _trend_multiplier(float(t) if pd.notna(t) else None)
            for t in trend_series
        ],
        index=pool.index,
    )

    # ── STEP 6: apply multipliers and clamp at 100 ────────────────────────────
    pool["composite_score"] = (base * seller_mult * trend_mult).clip(upper=100.0)

    # Risk flags
    pool["risk_flags"] = pool.apply(_build_risk_flags, axis=1)

    # Drop internal sub-score columns
    pool.drop(
        columns=[c for c in pool.columns if c.startswith("_s_")],
        inplace=True,
    )

    return pool.sort_values("composite_score", ascending=False).reset_index(drop=True)


def select_top_products(df: pd.DataFrame, top_n: int = 10) -> pd.DataFrame:
    """Score all products and return the best *top_n* (default 10)."""
    return score_all_products(df).head(top_n).reset_index(drop=True)
