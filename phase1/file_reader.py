"""
Reads and parses Helium 10 Black Box XLSX export files.
Normalises the variety of column names that different H10 versions produce
into a single canonical set used throughout Phase 1.
"""

import re
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd

# Maps our internal (canonical) column names to all known H10 header variants.
COLUMN_MAPPINGS: Dict[str, List[str]] = {
    "asin": ["ASIN", "Asin", "asin"],
    "title": [
        "Product Name", "Product Title", "Title", "Name",
        "Item Name", "product name",
    ],
    "brand": ["Brand", "Brand Name", "brand"],
    "category": [
        "Category", "Main Category", "Category / Subcategory",
        "Product Category",
    ],
    "subcategory": [
        "Subcategory", "Sub-Category", "Sub Category",
        "Category > Sub Category", "Sub",
    ],
    "price": [
        "Price", "Avg Price", "Average Price", "Buy Box Price",
        "Current Price", "Selling Price",
    ],
    "monthly_revenue": [
        "Monthly Revenue", "Est. Monthly Revenue",
        "Estimated Monthly Revenue", "Revenue", "Est. Revenue",
        "Monthly Rev.", "Revenue/Month", "Estimated Revenue",
        "Est Monthly Revenue",
    ],
    # Kept as separate canonical fields so the revenue filter can choose the
    # correct source depending on the user's variation-products preference.
    "parent_revenue": [
        "Parent Level Revenue", "Parent Revenue", "Parent Level Rev",
    ],
    "asin_revenue": [
        "ASIN Revenue", "Asin Revenue", "Child Revenue", "ASIN Level Revenue",
    ],
    "monthly_sales": [
        "Monthly Sales", "Est. Monthly Sales",
        "Estimated Monthly Sales", "Sales", "Units Sold",
        "Est. Sales/Month", "Avg Monthly Sales",
        "Est. Units Sold", "Units/Month", "Estimated Sales",
        "Est Monthly Sales", "Monthly Unit Sales",
    ],
    # Kept as separate canonical fields so the sales figure can follow the same
    # variation-products preference as revenue (parent vs ASIN level).
    "parent_sales": [
        "Parent Level Sales", "Parent Sales", "Parent Level Units",
    ],
    "asin_sales": [
        "ASIN Sales", "Asin Sales", "Child Sales", "ASIN Level Sales",
    ],
    "bsr": [
        "BSR", "Best Seller Rank", "Avg BSR", "Average BSR",
        "Rank", "Best Sellers Rank", "Avg. BSR", "Sales Rank",
    ],
    "rating": [
        "Rating", "Avg Rating", "Average Rating", "Star Rating",
        "Stars", "Avg. Rating", "Customer Rating",
        # H10 Black Box GB
        "Reviews Rating",
    ],
    "reviews": [
        "Reviews", "Review Count", "# Reviews", "# of Reviews",
        "Number of Reviews", "Total Reviews", "Ratings",
        "Review #", "No. Reviews",
        # H10 Black Box GB
        "Review Count",
    ],
    "date_first_available": [
        "Date First Available", "First Available", "Launch Date",
        "Created Date", "Date Added", "First Date Available",
    ],
    "listing_age_months": [
        "Listing Age (Months)", "Age (Months)", "Listing Age",
        "Age", "Age in Months", "Months Listed",
    ],
    "fulfillment": [
        "Fulfillment", "Fulfillment Type", "FBA/FBM",
        "Fulfilled By", "Fulfillment Method",
    ],
    "seller": [
        "Seller", "Seller Name", "Sold By", "Merchant",
        "Brand/Seller",
    ],
    "seller_country": [
        "Seller Country/Region", "Seller Country", "Country",
        "Country/Region", "Seller Location", "Origin Country",
    ],
    "best_sales_period": [
        "Best Sales Period", "Best Sales Month", "Peak Sales Period",
    ],
    "net_margin": ["Net Margin", "Net Margin %", "Margin", "Net Profit Margin"],
    "fba_fees": ["FBA Fees", "Amazon FBA Fees", "Fees", "FBA Fee"],
    "sales_trend": [
        "Sales Trend", "30-Day Trend", "Sales Change", "Trend",
        "Monthly Sales Trend", "30 Day Trend",
        # H10 Black Box GB
        "Sales Trend (90 days) (%)",
    ],
    "weight": ["Weight", "Item Weight", "Product Weight"],
    "dimensions": ["Dimensions", "Product Dimensions", "Size"],
    "lqs": ["LQS", "Listing Quality Score", "Quality Score"],
    "images": ["Images", "Image Count", "# Images", "# of Images", "Number of Images"],
    "image_url": [
        "Image URL", "Image Url", "ImageURL", "Image Link",
        "Main Image URL", "Product Image URL", "Main Image",
    ],
    "variations": ["Variations", "Variation Count", "# Variations"],
    "active_sellers": [
        "Number of Active Sellers", "Active Sellers", "# Sellers",
        "Seller Count", "No. of Sellers", "Number of Sellers",
    ],
    "net_profit": ["Net Profit", "Profit", "Profit/Unit", "Net Profit Per Unit"],
}

NUMERIC_COLUMNS = [
    "price", "monthly_revenue", "parent_revenue", "asin_revenue",
    "monthly_sales", "parent_sales", "asin_sales", "bsr", "rating",
    "reviews", "net_margin", "fba_fees", "lqs", "images", "variations",
    "listing_age_months", "net_profit",
]


def _find_column(df: pd.DataFrame, canonical: str) -> Optional[str]:
    """Return the first df column that matches any known alias for *canonical*."""
    aliases = COLUMN_MAPPINGS.get(canonical, [])
    df_lower: Dict[str, str] = {c.lower().strip(): c for c in df.columns}

    for alias in aliases:
        if alias in df.columns:
            return alias
        if alias.lower().strip() in df_lower:
            return df_lower[alias.lower().strip()]

    # Broad partial-match fallback
    for col in df.columns:
        col_strip = col.lower().strip()
        for alias in aliases:
            if alias.lower() in col_strip:
                return col
    return None


def _parse_numeric(value) -> Optional[float]:
    """Convert a cell value (possibly a currency string or H10 dash) to float."""
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    # H10 uses '-' to represent missing / zero numeric values
    if s in ("-", "–", "—", "N/A", "n/a", ""):
        return None
    s = (
        s.replace(",", "")
        .replace("$", "")
        .replace("£", "")
        .replace("€", "")
        .replace("%", "")
        .replace("K", "000")
        .replace("k", "000")
    )
    try:
        return float(s)
    except (ValueError, TypeError):
        return None


def _compute_listing_age(value) -> Optional[float]:
    """Return listing age in months from a date string or existing numeric value."""
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    if isinstance(value, (datetime, pd.Timestamp)):
        dt = value if isinstance(value, datetime) else value.to_pydatetime()
    else:
        s = str(value).strip()
        parsed = False
        for fmt in [
            "%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y",
            "%B %d, %Y", "%b %d, %Y", "%d-%m-%Y",
            "%Y/%m/%d", "%d %B %Y",
        ]:
            try:
                dt = datetime.strptime(s, fmt)
                parsed = True
                break
            except ValueError:
                continue
        if not parsed:
            return None

    now = datetime.now()
    months = (now.year - dt.year) * 12 + (now.month - dt.month)
    return max(0.0, float(months))


def _detect_category(df: pd.DataFrame) -> str:
    """Return the most common value in the category column."""
    for col in ("category", "subcategory"):
        if col in df.columns:
            series = df[col].dropna()
            if len(series) > 0:
                mode = series.mode()
                if len(mode) > 0:
                    return str(mode.iloc[0])
    return "Unknown Category"


def read_blackbox_file(file_path: str) -> Tuple[pd.DataFrame, str]:
    """
    Read an H10 Black Box XLSX/CSV export and return a normalised DataFrame.

    Returns
    -------
    df : pd.DataFrame
        Normalised product data with canonical column names.
    category : str
        Most common product category detected in the file.
    """
    path = Path(file_path.strip().strip('"').strip("'"))
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")

    suffix = path.suffix.lower()

    if suffix in (".xlsx", ".xls"):
        df_raw: Optional[pd.DataFrame] = None
        for skip in (0, 1, 2):
            try:
                tmp = pd.read_excel(path, engine="openpyxl", skiprows=skip)
                # Accept this sheet if at least one ASIN-like column exists
                cols_lower = [c.lower() for c in tmp.columns]
                if any("asin" in c for c in cols_lower) or any(
                    re.match(r"^b[0-9a-z]{9}$", str(tmp.iloc[0, 0]).strip(), re.I)
                    for _ in [None]
                    if len(tmp) > 0
                ):
                    df_raw = tmp
                    break
            except Exception:
                pass
        if df_raw is None:
            df_raw = pd.read_excel(path, engine="openpyxl")
    elif suffix == ".csv":
        df_raw = pd.read_csv(path)
    else:
        raise ValueError(f"Unsupported format '{suffix}'. Supply .xlsx, .xls, or .csv")

    # Build normalised frame
    cols_used: Dict[str, str] = {}
    data: Dict[str, pd.Series] = {}
    for canonical in COLUMN_MAPPINGS:
        actual = _find_column(df_raw, canonical)
        if actual:
            cols_used[canonical] = actual
            data[canonical] = df_raw[actual].reset_index(drop=True)

    df = pd.DataFrame(data)

    # Append unmapped raw columns for reference (prefixed with 'raw_')
    mapped_actuals = set(cols_used.values())
    for col in df_raw.columns:
        if col not in mapped_actuals:
            safe_name = f"raw_{re.sub(r'[^a-zA-Z0-9_]', '_', str(col))}"
            df[safe_name] = df_raw[col].reset_index(drop=True)

    # Parse numeric columns
    for col in NUMERIC_COLUMNS:
        if col in df.columns and col != "listing_age_months":
            df[col] = df[col].apply(_parse_numeric)

    # Listing age: prefer existing column, fall back to computed from date
    if "listing_age_months" in df.columns:
        df["listing_age_months"] = df["listing_age_months"].apply(_parse_numeric)

    if (
        "listing_age_months" not in df.columns
        or df["listing_age_months"].isna().all()
    ) and "date_first_available" in df.columns:
        df["listing_age_months"] = df["date_first_available"].apply(
            _compute_listing_age
        )

    # Remove rows with no ASIN
    if "asin" in df.columns:
        df = df[df["asin"].notna()]
        df = df[df["asin"].astype(str).str.strip() != ""]
        df = df.reset_index(drop=True)

    category = _detect_category(df)

    print(f"\n  Products loaded    : {len(df)}")
    print(f"  Detected category  : {category}")
    detected_keys = [k for k in COLUMN_MAPPINGS if k in cols_used]
    print(f"  Columns mapped     : {', '.join(detected_keys)}")

    return df, category
