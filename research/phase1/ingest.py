"""
Phase 1 ingestion — Helium 10 Black Box XLSX -> scored rows -> PostgreSQL.

Flow (weekly full refresh):
  1. Read one or more Black Box exports (reusing `file_reader.read_blackbox_file`)
     and concatenate; de-duplicate by ASIN (the same product can appear across
     revenue-band exports).
  2. Pull price history to compute each ASIN's `price_stability_factor`
     (neutral on first ingest — no history yet).
  3. Score + segment every row (`scoring.compute_scores`).
  4. Write a `product_history` snapshot row per product (the newly observed
     price/BSR/sales/reviews), then UPSERT into `products` by ASIN — existing
     rows are updated in place (keeping `created_at`, bumping `last_updated`),
     new rows inserted, missing rows kept (history is never deleted).

Revenue and sales are taken at **ASIN level** (`asin_sales` / `asin_revenue`),
never parent level, per the project rule.

`read_and_score()` and `build_product_rows()` are pure (no DB) and unit-testable.
"""

from typing import Dict, List, Optional

import numpy as np
import pandas as pd

from research.phase1.file_reader import read_blackbox_file
from research.phase1.scoring import compute_scores
from research.phase1.db import Phase1DbError

# products columns written on insert (created_at/last_updated handled separately).
_PRODUCT_COLS = [
    "asin", "title", "brand", "category", "subcategory", "price", "bsr",
    "estimated_sales", "estimated_revenue", "review_count", "rating",
    "weight_kg", "yoy_growth", "is_seasonal", "variation_count", "has_variants",
    "is_amazon_seller", "is_compatibility", "is_global_brand", "fulfillment",
    "image_url", "url", "listing_age_months", "sales_trend_90d",
    "price_trend_90d", "price_range", "review_bucket", "sales_bucket",
    "competition_score", "demand_score", "opportunity_score",
]


def _v(value):
    """NaN/NaT -> None, numpy scalars -> python scalars (psycopg2-friendly)."""
    if value is None:
        return None
    if isinstance(value, float) and np.isnan(value):
        return None
    if isinstance(value, (np.integer,)):
        return int(value)
    if isinstance(value, (np.floating,)):
        return None if np.isnan(value) else float(value)
    if isinstance(value, (np.bool_,)):
        return bool(value)
    if isinstance(value, str) and value.strip() == "":
        return None
    return value


def _find_price_trend_col(df: pd.DataFrame) -> Optional[str]:
    for c in df.columns:
        cl = str(c).lower()
        if "price" in cl and "trend" in cl:
            return c
    return None


def read_and_score(paths: List[str]) -> pd.DataFrame:
    """Read + concat + de-dup Black Box files, then score. Pure (no DB)."""
    frames = []
    for p in paths:
        df, _ = read_blackbox_file(p)
        frames.append(df)
    if not frames:
        raise ValueError("No Black Box files supplied to ingest.")
    combined = pd.concat(frames, ignore_index=True)
    if "asin" in combined.columns:
        combined = combined.dropna(subset=["asin"]).drop_duplicates(
            subset=["asin"], keep="first"
        ).reset_index(drop=True)
    return compute_scores(combined)


def build_product_rows(scored: pd.DataFrame) -> List[dict]:
    """Map a scored frame to a list of `products`-row dicts. Pure (no DB)."""
    price_trend_col = _find_price_trend_col(scored)
    rows: List[dict] = []
    for _, r in scored.iterrows():
        asin = _v(r.get("asin"))
        if not asin:
            continue
        rows.append({
            "asin": asin,
            "title": _v(r.get("title")),
            "brand": _v(r.get("brand")),
            "category": _v(r.get("category")),
            "subcategory": _v(r.get("subcategory")),
            "price": _v(r.get("price")),
            "bsr": _v(r.get("bsr")),
            "estimated_sales": _v(r.get("asin_sales")),       # ASIN-level
            "estimated_revenue": _v(r.get("asin_revenue")),   # ASIN-level
            "review_count": _v(r.get("reviews")),
            "rating": _v(r.get("rating")),
            "weight_kg": _v(r.get("weight_kg")),
            "yoy_growth": _v(r.get("yoy_growth")),
            "is_seasonal": _v(r.get("is_seasonal")),
            "variation_count": _v(r.get("variation_count")),
            "has_variants": _v(r.get("has_variants")),
            "is_amazon_seller": _v(r.get("is_amazon_seller")),
            "is_compatibility": _v(r.get("is_compatibility")),
            "is_global_brand": _v(r.get("is_global_brand")),
            "fulfillment": _v(r.get("fulfillment")),
            "image_url": _v(r.get("image_url")),
            "url": f"https://www.amazon.co.uk/dp/{asin}",
            "listing_age_months": _v(r.get("listing_age_months")),
            "sales_trend_90d": _v(r.get("sales_trend_90d")),
            "price_trend_90d": _v(r.get(price_trend_col)) if price_trend_col else None,
            "price_range": _v(r.get("price_range")),
            "review_bucket": _v(r.get("review_bucket")),
            "sales_bucket": _v(r.get("sales_bucket")),
            "competition_score": _v(r.get("competition_score")),
            "demand_score": _v(r.get("demand_score")),
            "opportunity_score": _v(r.get("opportunity_score")),
        })
    return rows


# ── DB operations ───────────────────────────────────────────────────────────

def _snapshot_and_upsert(conn, rows: List[dict]) -> int:
    from psycopg2.extras import execute_values

    hist = [
        (r["asin"], r["price"], r["bsr"], r["estimated_sales"], r["review_count"])
        for r in rows
    ]
    cols = _PRODUCT_COLS
    values = [tuple(r[c] for c in cols) for r in rows]
    update_set = ", ".join(
        f"{c} = EXCLUDED.{c}" for c in cols if c != "asin"
    ) + ", last_updated = NOW()"

    try:
        with conn.cursor() as cur:
            execute_values(
                cur,
                "INSERT INTO product_history (asin, price, bsr, estimated_sales, review_count) VALUES %s",
                hist,
            )
            execute_values(
                cur,
                f"INSERT INTO products ({', '.join(cols)}) VALUES %s "
                f"ON CONFLICT (asin) DO UPDATE SET {update_set}",
                values,
            )
        conn.commit()
    except Exception as e:  # noqa: BLE001
        conn.rollback()
        raise Phase1DbError(f"Ingest write failed: {e}") from e
    return len(rows)


def ingest_files(conn, paths: List[str]) -> int:
    """Full ingest: read + score + write (history snapshot + UPSERT).
    Returns the number of products written."""
    scored = read_and_score(paths)
    rows = build_product_rows(scored)
    n = _snapshot_and_upsert(conn, rows)
    print(f"  Ingested {n} products")
    return n
