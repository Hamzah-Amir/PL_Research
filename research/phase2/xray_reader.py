"""
Reads and parses Helium 10 Xray **Keyword** export files (Phase 2 input).

In this project's workflow the Phase 2 keyword universe comes from the H10 Xray
Keywords view (exported for the target product's main keyword), not Cerebro —
though the export carries the same keyword metrics (Search Volume, Cerebro IQ
Score, Competing Products, Title Density, Keyword Sales, etc.).

Column names are normalised to a canonical set via `COLUMN_MAPPINGS`, the same
approach as Phase 1's `file_reader.py`. Mappings below are confirmed against a
real Xray keyword export (header on the first row):

    Keyword Phrase | Cerebro IQ Score | Search Volume | Search Volume Trend |
    Suggested PPC Bid | Keyword Sales | Competing Products | Title Density |
    Competitor Rank (avg)
"""

import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd

# Maps internal (canonical) names to the H10 Xray keyword-export headers, with
# broad alias fallback for minor version/locale differences.
COLUMN_MAPPINGS: Dict[str, List[str]] = {
    # The keyword phrase itself.
    "keyword": [
        "Keyword Phrase", "Keyword", "Keywords", "Phrase", "Search Term",
    ],
    # Cerebro IQ Score — H10's blended opportunity/relevancy metric.
    "cerebro_iq": [
        "Cerebro IQ Score", "Cerebro IQ", "IQ Score",
    ],
    # Monthly search volume — the primary selection signal alongside relevancy.
    "search_volume": [
        "Search Volume", "Search Vol", "SV", "Monthly Search Volume",
    ],
    # Search-volume trend (percentage).
    "search_volume_trend": [
        "Search Volume Trend", "Search Vol Trend", "SV Trend",
    ],
    # Suggested PPC bid — a formatted string range (e.g. "£0.46 (£0.32 – £0.70)").
    # Kept as text (not parsed numeric).
    "suggested_ppc_bid": [
        "Suggested PPC Bid", "Sponsored PPC Bid", "PPC Bid", "Suggested Bid",
    ],
    # Estimated units sold via the keyword.
    "keyword_sales": [
        "Keyword Sales", "Est. Keyword Sales", "Keyword Sales (est)",
    ],
    # Number of competing products for the keyword (lower = easier).
    "competing_products": [
        "Competing Products", "Competing Products Count", "# Competing Products",
    ],
    # How many top listings use the keyword in the title.
    "title_density": [
        "Title Density", "Title Density Exact", "Title Density (exact)",
    ],
    # Average competitor rank for the keyword.
    "competitor_rank": [
        "Competitor Rank (avg)", "Competitor Rank Avg", "Competitor Rank",
        "Avg Competitor Rank",
    ],
}

# Canonical columns parsed as numeric (strip currency/commas, H10 "-" -> None).
# suggested_ppc_bid is intentionally excluded — it is a formatted range string.
NUMERIC_COLUMNS = [
    "cerebro_iq", "search_volume", "search_volume_trend", "keyword_sales",
    "competing_products", "title_density", "competitor_rank",
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

    # Broad partial-match fallback (e.g. "Search Volume (UK)" -> search_volume)
    for col in df.columns:
        col_strip = str(col).lower().strip()
        for alias in aliases:
            if alias.lower() in col_strip:
                return col
    return None


def _parse_numeric(value) -> Optional[float]:
    """Convert a cell value (possibly a formatted string or H10 dash) to float."""
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    # H10 uses '-' to represent missing / not-ranked numeric values
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


def read_xray_keywords(file_path: str) -> Tuple[pd.DataFrame, int]:
    """
    Read an H10 Xray keyword XLSX/CSV export and return a normalised DataFrame.

    Returns
    -------
    df : pd.DataFrame
        Normalised keyword data with canonical column names. Unmapped columns
        are preserved with a 'raw_' prefix (never discarded).
    raw_row_count : int
        Number of keyword rows read.
    """
    path = Path(file_path.strip().strip('"').strip("'"))
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")

    suffix = path.suffix.lower()

    if suffix in (".xlsx", ".xls"):
        df_raw: Optional[pd.DataFrame] = None
        # Accept the first read where a keyword-like column appears (handles any
        # metadata rows above the header).
        for skip in (0, 1, 2):
            try:
                tmp = pd.read_excel(path, engine="openpyxl", skiprows=skip)
                cols_lower = [str(c).lower() for c in tmp.columns]
                if any("keyword" in c or "phrase" in c or "search term" in c
                       for c in cols_lower):
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

    # Parse numeric columns
    for col in NUMERIC_COLUMNS:
        if col in df.columns:
            df[col] = df[col].apply(_parse_numeric)

    # Drop rows with no keyword phrase
    if "keyword" in df.columns:
        df = df[df["keyword"].notna()]
        df = df[df["keyword"].astype(str).str.strip() != ""]
        df = df.reset_index(drop=True)

    raw_row_count = len(df)

    print(f"\n  Keywords loaded    : {raw_row_count}")
    detected_keys = [k for k in COLUMN_MAPPINGS if k in cols_used]
    print(f"  Columns mapped     : {', '.join(detected_keys) or '(none — check headers)'}")
    if "keyword" not in cols_used:
        print("  WARNING: no keyword-phrase column detected — verify the file is "
              "an H10 Xray keyword export.")
    if "search_volume" not in cols_used:
        print("  WARNING: no search-volume column detected — selection prep will "
              "be limited.")

    return df, raw_row_count
