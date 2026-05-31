"""
Reads and parses Helium 10 Cerebro keyword export files (Phase 2 input).

Cerebro is run on the single target ASIN chosen at the end of Phase 1 and
exports the keyword universe for that ASIN. This module normalises the variety
of column names different H10 versions / locales produce into a single
canonical set used throughout Phase 2 — the same approach as Phase 1's
`file_reader.py`.

NOTE: the canonical mappings below are built from the standard documented
Cerebro export headers plus broad alias fallbacks. They have not yet been
verified against a real export from this account; once a sample Cerebro XLSX is
available, confirm the header names and tighten `COLUMN_MAPPINGS` accordingly.
"""

import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd

# Maps our internal (canonical) column names to all known H10 Cerebro header
# variants. Order within each list is preference order for exact matching.
COLUMN_MAPPINGS: Dict[str, List[str]] = {
    # The keyword phrase itself.
    "keyword": [
        "Keyword Phrase", "Keyword", "Keywords", "Phrase",
        "Search Term", "keyword phrase",
    ],
    # Monthly search volume (Amazon). Primary selection signal alongside relevancy.
    "search_volume": [
        "Search Volume", "Search Vol", "SV", "Monthly Search Volume",
        "Est. Search Volume", "search volume",
    ],
    # 30/90-day search-volume trend (percentage).
    "search_volume_trend": [
        "Search Volume Trend", "Search Vol Trend", "SV Trend",
        "Search Volume Trend (90 days)", "Search Volume Trend (30 days)",
    ],
    # Cerebro IQ Score — H10's blended opportunity/relevancy metric.
    "cerebro_iq": [
        "Cerebro IQ Score", "Cerebro IQ", "IQ Score", "CerebroIQ",
    ],
    # Organic rank of the queried (target) ASIN for this keyword.
    "organic_rank": [
        "Organic Rank", "Org. Rank", "Organic Position",
        "Average Organic Rank", "Organic Rank (avg)",
    ],
    # Sponsored rank of the queried ASIN for this keyword.
    "sponsored_rank": [
        "Sponsored Rank", "Sponsored Position", "Sponsored Rank (avg)",
    ],
    # Position (Rank) — used heavily in later phases (Cerebro reverse-ASIN).
    "position_rank": [
        "Position (Rank)", "Position Rank", "Position", "Rank",
    ],
    # Number of competing products for the keyword (lower = easier).
    "competing_products": [
        "Competing Products", "Competing Products Count",
        "# Competing Products", "Competing", "Products",
    ],
    # CPR — H10 "Cerebro Product Rank" 8-day giveaway estimate.
    "cpr": [
        "CPR", "CPR 8-Day Giveaways", "Cerebro Product Rank",
        "CPR (8-Day Giveaways)", "8-Day Giveaways",
    ],
    # Title density — how many top listings use the keyword in the title.
    "title_density": [
        "Title Density", "Title Density Exact", "Title Density (exact)",
    ],
    # Number of sponsored ASINs competing on the keyword.
    "sponsored_asins": [
        "Sponsored ASINs", "Sponsored ASIN Count", "# Sponsored ASINs",
    ],
    # Keyword sales (estimated units sold via the keyword).
    "keyword_sales": [
        "Keyword Sales", "Est. Keyword Sales", "Keyword Sales (est)",
    ],
    # Word count of the phrase (1 = single word -> often too broad).
    "word_count": [
        "Word Count", "Words", "# Words", "Number of Words",
    ],
    # Match type / relevancy descriptors when present.
    "match_type": [
        "Match Type", "Keyword Type", "Type",
    ],
    # Amazon-recommended rank, when present.
    "amazon_recommended_rank": [
        "Amazon Recommended Rank", "Amazon Recommended", "Amazon Rec. Rank",
    ],
}

# Canonical columns parsed as numeric (strip currency/commas, H10 "-" -> None).
NUMERIC_COLUMNS = [
    "search_volume", "search_volume_trend", "cerebro_iq", "organic_rank",
    "sponsored_rank", "position_rank", "competing_products", "cpr",
    "title_density", "sponsored_asins", "keyword_sales", "word_count",
    "amazon_recommended_rank",
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
        col_strip = col.lower().strip()
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


def read_cerebro_file(file_path: str) -> Tuple[pd.DataFrame, int]:
    """
    Read an H10 Cerebro XLSX/CSV export and return a normalised DataFrame.

    Returns
    -------
    df : pd.DataFrame
        Normalised keyword data with canonical column names. Unmapped columns
        are preserved with a 'raw_' prefix (never discarded).
    raw_row_count : int
        Number of keyword rows read before any Phase 2 preparation.
    """
    path = Path(file_path.strip().strip('"').strip("'"))
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")

    suffix = path.suffix.lower()

    if suffix in (".xlsx", ".xls"):
        df_raw: Optional[pd.DataFrame] = None
        # Cerebro exports sometimes carry one or two metadata rows above the
        # header. Accept the first read where a keyword-like column appears.
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
              "a Cerebro export.")
    if "search_volume" not in cols_used:
        print("  WARNING: no search-volume column detected — selection prep will "
              "be limited.")

    return df, raw_row_count
