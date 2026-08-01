"""
Add products to the database from a Helium 10 Black Box Excel export (or any
sheet with the same headers). Upserts by ASIN, so re-uploading updates rows.

This is the single re-added ingest path (the original bulk Phase-1 ingest was
removed when Phase 1 moved to the ORM). It fills the columns a Black Box export
carries; the derived opportunity/bucket scores are left NULL (they were produced
by the old scoring pipeline) — a product still works in Phase 2/3 and appears in
the shortlist (unscored rows sort last).
"""
from __future__ import annotations

import re
from typing import Optional

from products.models import Product


def _num(v) -> Optional[float]:
    """Parse a money/number-ish cell ('£12.99', '1,234', '4.5%') to float."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d.\-]", "", str(v))
    try:
        return float(s) if s not in ("", "-", ".") else None
    except ValueError:
        return None


def _int(v) -> Optional[int]:
    n = _num(v)
    return int(round(n)) if n is not None else None


def _weight_kg(v) -> Optional[float]:
    """'0.34 kg' -> 0.34, '340 g' -> 0.34, plain number assumed kg."""
    if v is None:
        return None
    s = str(v).lower()
    n = _num(s)
    if n is None:
        return None
    if "g" in s and "kg" not in s:  # grams
        return round(n / 1000.0, 4)
    return n


def _text(v) -> Optional[str]:
    if v is None:
        return None
    s = str(v).strip()
    return s or None


# Black Box header (lower-cased) -> (Product field, parser)
_MAP = {
    "title": ("title", _text),
    "brand": ("brand", _text),
    "category": ("category", _text),
    "subcategory": ("subcategory", _text),
    "price": ("price", _num),
    "bsr": ("bsr", _int),
    "asin sales": ("estimated_sales", _num),
    "asin revenue": ("estimated_revenue", _num),
    "review count": ("review_count", _int),
    "reviews rating": ("rating", _num),
    "sales year over year (%)": ("yoy_growth", _num),
    "sales trend (90 days) (%)": ("sales_trend_90d", _num),
    "price trend (90 days) (%)": ("price_trend_90d", _num),
    "listing age (months)": ("listing_age_months", _num),
    "weight": ("weight_kg", _weight_kg),
    "variation count": ("variation_count", _int),
    "fulfillment": ("fulfillment", _text),
    "url": ("url", _text),
    "image url": ("image_url", _text),
}

_ASIN_RE = re.compile(r"^[A-Z0-9]{10}$")


def ingest_products_excel(file) -> dict:
    """Read a Black Box-style .xlsx and upsert each row into `Product` by ASIN.
    Returns {added, updated, skipped, error}. Never raises for bad rows."""
    import openpyxl

    try:
        wb = openpyxl.load_workbook(file, read_only=True, data_only=True)
    except Exception as e:  # noqa: BLE001
        return {"error": f"Could not read the Excel file: {e}"}
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    try:
        header = [str(h).strip().lower() if h is not None else "" for h in next(rows)]
    except StopIteration:
        return {"error": "The sheet is empty."}
    if "asin" not in header:
        return {"error": "No 'ASIN' column found — upload a Black Box (Products) export."}

    added = updated = skipped = 0
    for raw in rows:
        rec = dict(zip(header, raw))
        asin = str(rec.get("asin") or "").strip().upper()
        if not _ASIN_RE.match(asin):
            skipped += 1
            continue
        defaults = {}
        for h, (field, parse) in _MAP.items():
            if h in rec:
                val = parse(rec.get(h))
                if val is not None:
                    defaults[field] = val
        # Derived flags a Black Box export implies.
        ful = str(rec.get("fulfillment") or "").strip().lower()
        seller = str(rec.get("seller") or "").strip().lower()
        defaults["is_amazon_seller"] = ful == "amazon" or seller == "amazon"
        vc = defaults.get("variation_count")
        if vc is not None:
            defaults["has_variants"] = vc > 1
        _, created = Product.objects.update_or_create(asin=asin, defaults=defaults)
        added += created
        updated += not created
    return {"added": added, "updated": updated, "skipped": skipped, "error": None}
