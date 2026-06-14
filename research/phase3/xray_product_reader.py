"""
Reads and parses a Helium 10 Xray **Products** export (Phase 3 input).

This is the product/listing search Xray view (one export per keyword, run on
Amazon UK) — distinct from `phase2/xray_reader.py`, which reads the Xray
*Keywords* view. It is the richest of the three Phase 3 sources and carries
several fields no other source gives us:

  - `Sponsored` — the ad-placement flag per keyword. Values seen in real
    exports: "Sponsored" (Sponsored Products), "Sponsored Brand" /
    "Sponsored Brand Video" (Sponsored Brands, the ex-"AMS" banner ads), or
    blank (organic). This drives the Sponsored Products and Sponsored Brands
    scoring once aggregated across all 6 keyword exports.
  - `Review velocity`, `Best Seller` (Yes/No), `Creation Date`, `Seller Age`,
    rating + review count, sales/revenue, fulfillment (FBA / AMZ / MFN).

Sponsored status is KEYWORD-SPECIFIC, so Phase 3 reads one Xray Products export
per launch keyword and aggregates per ASIN (see `aggregate_sponsored`).

openpyxl cannot open these files directly (embedded thumbnail drawings trip its
reader), but `pandas.read_excel` reads them fine — so we go through pandas.
Column names are normalised via `COLUMN_MAPPINGS` (broad alias fallback +
numeric parsing), the same approach as the other readers. Note the real `Price`
and `Fees` headers carry a trailing currency glyph (e.g. "Price  �"),
handled by prefix matching.
"""

import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd

COLUMN_MAPPINGS: Dict[str, List[str]] = {
    "asin": ["ASIN", "Asin"],
    "title": ["Product Details", "Title", "Product Name", "Product Title"],
    "url": ["URL", "Product URL", "Link"],
    "image_url": ["Image URL", "Image Url", "Main Image"],
    "brand": ["Brand", "Brand Name"],
    "price": ["Price"],
    "parent_sales": ["Parent Level Sales", "Parent Sales"],
    "asin_sales": ["ASIN Sales", "Asin Sales"],
    "recent_purchases": ["Recent Purchases", "Bought in past month"],
    "parent_revenue": ["Parent Level Revenue", "Parent Revenue"],
    "asin_revenue": ["ASIN Revenue", "Asin Revenue"],
    "title_char_count": ["Title Char. Count", "Title Character Count", "Title Length"],
    "bsr": ["BSR", "Best Seller Rank", "Sales Rank"],
    "seller_country": ["Seller Country/Region", "Seller Country", "Country"],
    "fees": ["Fees"],
    "active_sellers": ["Active Sellers", "Sellers", "No. of Sellers"],
    "rating": ["Ratings", "Rating", "Star Rating", "Stars"],
    "review_count": ["Review Count", "Reviews", "# of Reviews"],
    "images": ["Images", "Image Count", "# of Images"],
    "review_velocity": ["Review velocity", "Review Velocity", "Reviews/Month"],
    "buy_box": ["Buy Box", "Buy Box Seller", "BuyBox"],
    "category": ["Category", "Product Category"],
    "size_tier": ["Size Tier", "Product Tier", "Tier"],
    "fulfillment": ["Fulfillment", "Fulfilment", "Fulfillment Type"],
    "dimensions": ["Dimensions", "Product Dimensions"],
    "weight": ["Weight", "Item Weight"],
    "aba_most_clicked": ["ABA Most Clicked", "ABA Click Share"],
    "creation_date": ["Creation Date", "Date First Available", "Listed Since"],
    "sponsored": ["Sponsored", "Sponsored Type", "Ad Type"],
    "best_seller": ["Best Seller", "Bestseller", "Best Seller Badge"],
    "seller_age_mo": ["Seller Age (mo)", "Seller Age", "Seller Age (months)"],
    "seller": ["Seller", "Seller Name"],
}

# Parsed as numeric (strip currency/commas/glyph, H10 "-" -> None).
NUMERIC_COLUMNS = [
    "price", "parent_sales", "asin_sales", "recent_purchases", "parent_revenue",
    "asin_revenue", "title_char_count", "bsr", "fees", "active_sellers", "rating",
    "review_count", "images", "review_velocity", "weight", "seller_age_mo",
]

_ASIN_RE = re.compile(r"^B0[0-9A-Z]{8}$")


def _find_column(df: pd.DataFrame, canonical: str) -> Optional[str]:
    aliases = COLUMN_MAPPINGS.get(canonical, [])
    df_lower: Dict[str, str] = {str(c).lower().strip(): c for c in df.columns}
    for alias in aliases:
        if alias in df.columns:
            return alias
        if alias.lower().strip() in df_lower:
            return df_lower[alias.lower().strip()]
    # Prefix match handles trailing glyphs/units (e.g. "Price  �").
    for col in df.columns:
        col_strip = str(col).lower().strip()
        for alias in aliases:
            if col_strip == alias.lower() or col_strip.startswith(alias.lower()):
                return col
    return None


def _parse_numeric(value) -> Optional[float]:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if s in ("-", "–", "—", "N/A", "n/a", ""):
        return None
    s = (
        s.replace(",", "").replace("$", "").replace("£", "").replace("€", "")
        .replace("%", "").replace("�", "").strip()
    )
    m = re.match(r"^-?\d+(\.\d+)?", s)
    if m:
        try:
            return float(m.group(0))
        except (ValueError, TypeError):
            return None
    return None


def _clean_asin(value) -> Optional[str]:
    if value is None:
        return None
    s = str(value).strip().upper()
    return s if _ASIN_RE.match(s) else None


def classify_sponsored(value) -> Optional[str]:
    """Map an Xray `Sponsored` cell to a canonical ad class.

    Returns 'sponsored_brand' for Sponsored Brand / Sponsored Brand Video
    (the ex-"AMS" banner ads), 'sponsored_product' for plain "Sponsored", or
    None for organic / blank.
    """
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return None
    s = str(value).strip().lower()
    if not s or s in ("-", "n/a", "none", "organic"):
        return None
    if s.startswith("sponsored brand"):
        return "sponsored_brand"
    if s.startswith("sponsored"):
        return "sponsored_product"
    return None


def read_xray_products(file_path: str) -> Tuple[pd.DataFrame, int]:
    """
    Read one H10 Xray Products XLSX export and return a normalised DataFrame.

    Adds a canonical `sponsored_class` column ('sponsored_brand' /
    'sponsored_product' / None) derived from the raw `Sponsored` cell. Rows
    without a valid ASIN are dropped. Unmapped columns are kept with a 'raw_'
    prefix.
    """
    path = Path(file_path.strip().strip('"').strip("'"))
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")

    suffix = path.suffix.lower()
    if suffix in (".xlsx", ".xls"):
        df_raw = pd.read_excel(path)  # pandas handles the embedded drawings
    elif suffix == ".csv":
        df_raw = pd.read_csv(path)
    else:
        raise ValueError(f"Unsupported format '{suffix}'. Supply .xlsx, .xls, or .csv")

    cols_used: Dict[str, str] = {}
    data: Dict[str, pd.Series] = {}
    for canonical in COLUMN_MAPPINGS:
        actual = _find_column(df_raw, canonical)
        if actual and actual not in cols_used.values():
            cols_used[canonical] = actual
            data[canonical] = df_raw[actual].reset_index(drop=True)

    df = pd.DataFrame(data)

    mapped_actuals = set(cols_used.values())
    for col in df_raw.columns:
        if col not in mapped_actuals:
            safe_name = f"raw_{re.sub(r'[^a-zA-Z0-9_]', '_', str(col))}"
            df[safe_name] = df_raw[col].reset_index(drop=True)

    for col in NUMERIC_COLUMNS:
        if col in df.columns:
            df[col] = df[col].apply(_parse_numeric)

    if "sponsored" in df.columns:
        df["sponsored_class"] = df["sponsored"].apply(classify_sponsored)
    else:
        df["sponsored_class"] = None

    if "asin" in df.columns:
        df["asin"] = df["asin"].apply(_clean_asin)
        df = df[df["asin"].notna()].reset_index(drop=True)

    row_count = len(df)

    print(f"\n  Xray products loaded : {row_count}")
    detected = [k for k in COLUMN_MAPPINGS if k in cols_used]
    print(f"  Columns mapped       : {', '.join(detected) or '(none — check headers)'}")
    if "sponsored" not in cols_used:
        print("  WARNING: no 'Sponsored' column detected — Sponsored Products / "
              "Sponsored Brands scoring will be unavailable from this file.")
    return df, row_count


def aggregate_sponsored(per_keyword: Dict[str, pd.DataFrame]) -> Dict[str, Dict]:
    """
    Aggregate sponsored presence per ASIN across multiple keyword exports.

    *per_keyword* maps a keyword label -> its read_xray_products() DataFrame.
    Returns ``{asin: {"sp_keywords": [...], "sb_keywords": [...],
    "sp_count": int, "sb_count": int, "total_keywords": int}}`` where sp = a
    Sponsored Product placement and sb = a Sponsored Brand placement, listing
    the keywords on which the ASIN advertised. This feeds the "ads on many
    high-volume keywords" rubric criteria.
    """
    agg: Dict[str, Dict] = {}
    total_keywords = len(per_keyword)
    for kw, df in per_keyword.items():
        if "asin" not in df.columns:
            continue
        for _, row in df.iterrows():
            asin = row.get("asin")
            if not asin:
                continue
            rec = agg.setdefault(
                asin, {"sp_keywords": [], "sb_keywords": []}
            )
            cls = row.get("sponsored_class")
            if cls == "sponsored_product" and kw not in rec["sp_keywords"]:
                rec["sp_keywords"].append(kw)
            elif cls == "sponsored_brand" and kw not in rec["sb_keywords"]:
                rec["sb_keywords"].append(kw)

    for asin, rec in agg.items():
        rec["sp_count"] = len(rec["sp_keywords"])
        rec["sb_count"] = len(rec["sb_keywords"])
        rec["total_keywords"] = total_keywords
    return agg
