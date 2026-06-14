"""
Reads and parses a Jungle Scout **keyword/product search** CSV export
(Phase 3, Step 7 input).

The JS export is the keyword-search product list for the main product keyword.
It carries PARENT ASINs only (per workflow Rule 7 — variation/child ASINs come
from Keepa). For Phase 3 it provides the sales/revenue estimates (the workflow's
chosen sales source) and the parent-ASIN universe that seeds the competitor
pool, plus Date First Available, dimensions and weight.

Column names are normalised to a canonical set via `COLUMN_MAPPINGS`, the same
approach as `phase1/file_reader.py` and `phase2/xray_reader.py`. Mappings are
confirmed against a real JS UK export whose header row carries an
"Export time: ..." banner in the first cell (the real column headers — ASIN,
Product Name, Brand, Price, Monthly Units Sold, ... — sit on that same row, so a
plain header=0 read works; the banner/empty columns are ignored).
"""

import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd

# Maps internal (canonical) names to the JS export headers, with broad alias
# fallback for minor version/locale differences.
COLUMN_MAPPINGS: Dict[str, List[str]] = {
    "asin": ["ASIN", "Asin", "Product ASIN"],
    "title": ["Product Name", "Title", "Product Title", "Name"],
    "brand": ["Brand", "Brand Name"],
    "price": ["Price", "Current Price", "Buy Box Price"],
    "monthly_sales": [
        "Monthly Units Sold", "Est. Monthly Sales", "Monthly Sales",
        "Estimated Monthly Sales", "Units Sold",
    ],
    "daily_sales": ["Daily Units Sold", "Est. Daily Sales", "Daily Sales"],
    "monthly_revenue": [
        "Monthly Revenue", "Est. Monthly Revenue", "Revenue", "Estimated Revenue",
    ],
    "net_revenue": ["Net Revenue", "Net Sales"],
    "date_first_available": [
        "Date First Available", "First Available", "Listing Created",
        "Date Available",
    ],
    "rating": ["Star Rating", "Rating", "Average Rating", "Stars"],
    "reviews": ["Reviews", "Review Count", "Number of Reviews", "# of Reviews"],
    "amazon_fees": ["Amazon Fees", "FBA Fees", "Fees", "Estimated Fees"],
    "bsr": ["BSR", "Best Seller Rank", "Sales Rank", "Rank"],
    "lqs": ["LQS", "Listing Quality Score"],
    "fulfillment": ["Fulfillment", "Fulfilment", "Fulfillment Type", "Seller Type"],
    "num_sellers": ["No. of Sellers", "Number of Sellers", "Sellers", "# of Sellers"],
    "category": ["Category", "Product Category", "Parent Category"],
    "product_tier": ["Product Tier", "Tier", "Size Tier"],
    "dimensions": ["Dimensions", "Product Dimensions", "Size"],
    "weight": ["Weight", "Item Weight", "Product Weight"],
    "link": ["Link", "URL", "Product URL", "Amazon URL"],
}

# Canonical columns parsed as numeric (strip currency/commas, JS "-" -> None).
NUMERIC_COLUMNS = [
    "price", "monthly_sales", "daily_sales", "monthly_revenue", "net_revenue",
    "rating", "reviews", "amazon_fees", "bsr", "lqs", "num_sellers",
]

_ASIN_RE = re.compile(r"^B0[0-9A-Z]{8}$")


def _find_column(df: pd.DataFrame, canonical: str) -> Optional[str]:
    """Return the first df column that matches any known alias for *canonical*."""
    aliases = COLUMN_MAPPINGS.get(canonical, [])
    df_lower: Dict[str, str] = {str(c).lower().strip(): c for c in df.columns}

    for alias in aliases:
        if alias in df.columns:
            return alias
        if alias.lower().strip() in df_lower:
            return df_lower[alias.lower().strip()]

    # Broad partial-match fallback, but guard against over-eager substring hits
    # (e.g. "Net Revenue" should not be picked up for "monthly_revenue").
    for col in df.columns:
        col_strip = str(col).lower().strip()
        for alias in aliases:
            if col_strip == alias.lower() or col_strip.startswith(alias.lower()):
                return col
    return None


def _parse_numeric(value) -> Optional[float]:
    """Convert a cell value (possibly a formatted string or JS dash) to float."""
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if s in ("-", "–", "—", "N/A", "n/a", ""):
        return None
    s = (
        s.replace(",", "").replace("$", "").replace("£", "").replace("€", "")
        .replace("%", "").replace("�", "")
    )
    # "1.2 kg", "14.3 x 31 cm" etc. — take the leading number if present.
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


def read_js_products(file_path: str) -> Tuple[pd.DataFrame, int]:
    """
    Read a Jungle Scout keyword-search CSV/XLSX export and return a normalised
    DataFrame.

    Returns
    -------
    df : pd.DataFrame
        Normalised product data with canonical column names. Rows without a
        valid ASIN are dropped. Unmapped columns are preserved with a 'raw_'
        prefix (never discarded).
    row_count : int
        Number of product rows read.
    """
    path = Path(file_path.strip().strip('"').strip("'"))
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")

    suffix = path.suffix.lower()
    if suffix == ".csv":
        df_raw = pd.read_csv(path)
    elif suffix in (".xlsx", ".xls"):
        df_raw = pd.read_excel(path, engine="openpyxl")
    else:
        raise ValueError(f"Unsupported format '{suffix}'. Supply .csv, .xlsx, or .xls")

    cols_used: Dict[str, str] = {}
    data: Dict[str, pd.Series] = {}
    for canonical in COLUMN_MAPPINGS:
        actual = _find_column(df_raw, canonical)
        if actual and actual not in cols_used.values():
            cols_used[canonical] = actual
            data[canonical] = df_raw[actual].reset_index(drop=True)

    df = pd.DataFrame(data)

    # Append unmapped raw columns for reference (prefixed with 'raw_')
    mapped_actuals = set(cols_used.values())
    for col in df_raw.columns:
        if col not in mapped_actuals:
            safe_name = f"raw_{re.sub(r'[^a-zA-Z0-9_]', '_', str(col))}"
            df[safe_name] = df_raw[col].reset_index(drop=True)

    for col in NUMERIC_COLUMNS:
        if col in df.columns:
            df[col] = df[col].apply(_parse_numeric)

    # Keep only rows with a valid ASIN.
    if "asin" in df.columns:
        df["asin"] = df["asin"].apply(_clean_asin)
        df = df[df["asin"].notna()].reset_index(drop=True)

    row_count = len(df)

    print(f"\n  JS products loaded : {row_count}")
    detected = [k for k in COLUMN_MAPPINGS if k in cols_used]
    print(f"  Columns mapped     : {', '.join(detected) or '(none — check headers)'}")
    if "asin" not in cols_used:
        print("  WARNING: no ASIN column detected — verify this is a JS product export.")
    if "monthly_sales" not in cols_used:
        print("  WARNING: no monthly-sales column detected — sales ranking will be limited.")

    return df, row_count
