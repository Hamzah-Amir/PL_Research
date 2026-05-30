"""
Filter configuration and application for Phase 1 product hunting.

Most filters are built-in with hardcoded values. The external inputs are
budget (sets the revenue cap), the seasonal-products preference, and the
product category (used only for display).

Revenue and monthly sales are read from ASIN-level fields. Revenue is capped
at budget × 1.52 (no minimum). There is no variation filter — products are
included regardless of whether they have variations.

Built-in filter values
----------------------
F1  Remove Amazon-as-seller         : always ON
F2  Seasonal products               : user choice (keep+flag vs exclude; Christmas/Q4 always removed)
F3  Minimum listing age             : 6 months
F4  Minimum star rating             : > 3.8  (strictly greater)
F5  Minimum unit price              : GBP 5 (built-in, no upper cap)
    Maximum monthly revenue         : budget × 1.52  (ASIN-level, no minimum)
    Monthly sales are preserved from ASIN-level data for scoring and output, but are not used as a selection filter
F6  90-day sales trend decline      : remove if decline > 30 %
F7  Maximum review count            : 3 000
F8  Regulated category exclusions   : none (applied silently with empty list)
F9  Remove compatibility products   : always ON
"""

import re
from dataclasses import dataclass, field
from typing import List, Optional, Tuple

import pandas as pd

# ──────────────────────────────────────────────────────────────────────────────
# Hardcoded filter values — change here to update the backend rules
# ──────────────────────────────────────────────────────────────────────────────

_F3_MIN_LISTING_AGE   = 6      # months
_F4_MIN_RATING        = 3.8    # stars (strict: rating must be > this value)
_F5_MIN_PRICE         = 5.0    # GBP — built-in minimum unit price (no upper cap)
_F6_MAX_DECLINE_PCT   = -30.0  # 90-day sales trend threshold (negative = decline)
_F7_MAX_REVIEWS       = 3_000  # maximum review count

# ──────────────────────────────────────────────────────────────────────────────
# F9 — compatibility / accessory detection patterns
# ──────────────────────────────────────────────────────────────────────────────

COMPATIBILITY_PATTERNS = [
    r"\bfor\s+(apple|iphone|ipad|samsung|android|huawei|xiaomi|macbook|laptop|ps[45]|xbox|nintendo)\b",
    r"\bcompatible\s+with\b",
    r"\breplacement\s+for\b",
    r"\bfits?\s+(all|most|standard|your)\b",
    r"\baccessory\s+for\b",
    r"\battachment\s+for\b",
    r"\b(case|cover|holster|charger|cable|stand|mount|holder)\s+for\b",
    r"\bdesigned\s+for\b",
    r"\bworks\s+with\b",
]

# ──────────────────────────────────────────────────────────────────────────────
# F2 — seasonality classification
#
# Each product is classified REMOVE / FLAG / PASS (priority REMOVE > FLAG > PASS):
#   1. Best Sales Period month (MM/DD/YYYY):
#        10/11/12 -> REMOVE (Christmas/Q4)
#        06/07/08 -> FLAG   (summer)
#        missing/"-" -> FLAG
#        other    -> PASS
#   2. Title keywords (case-insensitive, word boundary):
#        christmas/xmas/halloween/easter/valentine -> REMOVE
#        summer/beach/pool/bbq/camping/swim        -> FLAG
#
# REMOVE products are always dropped. FLAG products are kept + noted when the
# user includes seasonal products, otherwise dropped. PASS is always kept.
# ──────────────────────────────────────────────────────────────────────────────

_SEASONAL_REMOVE_KEYWORDS = ["christmas", "xmas", "halloween", "easter", "valentine"]
_SEASONAL_FLAG_KEYWORDS   = ["summer", "beach", "pool", "bbq", "camping", "swim"]

MONTH_NAMES = {
    1: "Jan", 2: "Feb", 3: "Mar",  4: "Apr",  5: "May",  6: "Jun",
    7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec",
}


def _parse_best_sales_month(value) -> Optional[int]:
    """Extract the month (1–12) from a Best Sales Period 'MM/DD/YYYY' value.
    Returns None when the value is missing or '-'."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    s = str(value).strip()
    if s in ("", "-", "–", "—", "nan", "NaN", "N/A", "n/a", "None"):
        return None
    head = s.split("/")[0].strip()
    try:
        m = int(head)
        return m if 1 <= m <= 12 else None
    except ValueError:
        return None


def _classify_seasonal(row) -> Tuple[str, str]:
    """Classify one product's seasonality.

    Returns (verdict, reason) where verdict is 'remove', 'flag', or 'pass'.
    REMOVE outranks FLAG outranks PASS. The reason cites the Best Sales Period
    value or the matched title keyword(s) for transparency.
    """
    title = str(row.get("title", "") or "").lower()
    remove_reasons: List[str] = []
    flag_reasons: List[str] = []

    # Title keywords (word boundary at the start of the keyword)
    for kw in _SEASONAL_REMOVE_KEYWORDS:
        if re.search(r"\b" + re.escape(kw), title):
            remove_reasons.append(f"title keyword '{kw}'")
    for kw in _SEASONAL_FLAG_KEYWORDS:
        if re.search(r"\b" + re.escape(kw), title):
            flag_reasons.append(f"title keyword '{kw}'")

    # Best Sales Period month
    month = _parse_best_sales_month(row.get("best_sales_period"))
    if month in (10, 11, 12):
        remove_reasons.append(f"Best Sales Period month {month:02d} (Q4/Christmas)")
    elif month in (6, 7, 8):
        flag_reasons.append(f"Best Sales Period month {month:02d} (summer)")
    elif month is None:
        flag_reasons.append("Best Sales Period missing")

    if remove_reasons:
        return "remove", "; ".join(remove_reasons)
    if flag_reasons:
        return "flag", "; ".join(flag_reasons)
    return "pass", ""


# ──────────────────────────────────────────────────────────────────────────────
# FilterConfig dataclass
# ──────────────────────────────────────────────────────────────────────────────

@dataclass
class FilterConfig:
    """All filter settings for one Phase 1 run."""

    # Revenue filter (parent-level revenue). Keep rows where revenue <= max_revenue.
    # min_revenue is an optional floor (0 = no minimum); only applied when > 0.
    max_revenue: float = 0.0          # cap = budget × 1.52
    min_revenue: float = 0.0          # optional floor (0 = none)

    # F1
    f1_remove_amazon: bool = True

    # F2 — seasonality. want_seasonal=False excludes all seasonal products;
    # True keeps FLAG (summer/unknown) products with a note (REMOVE/Christmas
    # products are always dropped either way).
    want_seasonal: bool = False

    # F3
    f3_min_listing_age_months: int = _F3_MIN_LISTING_AGE

    # F4
    f4_min_rating: float = _F4_MIN_RATING

    # F5 — built-in minimum unit price (GBP, no upper cap).
    # Monthly sales are collected from ASIN-level fields but not filtered.
    f5_min_price: float = _F5_MIN_PRICE

    # F6
    f6_apply: bool = True
    f6_max_decline_pct: float = _F6_MAX_DECLINE_PCT

    # F7
    f7_max_reviews: int = _F7_MAX_REVIEWS

    # F8
    f8_excluded_categories: List[str] = field(default_factory=list)

    # F9
    f9_remove_compatibility: bool = True


# ──────────────────────────────────────────────────────────────────────────────
# Silent configuration — called once, no user prompts
# ──────────────────────────────────────────────────────────────────────────────

def build_filter_config(
    category: str,
    max_revenue: float,
    want_seasonal: bool = False,
) -> FilterConfig:
    """
    Build a FilterConfig with built-in values plus the seasonal-products
    preference. F5 (minimum unit price) is built-in, not user-supplied.

    Revenue and monthly sales are always read from ASIN-level fields.
    Revenue is capped at budget × 1.52 with no minimum. There is no variation filter.

    When *want_seasonal* is True, FLAG (summer/unknown) products are kept and
    noted; otherwise all seasonal products are removed. Hard-seasonal
    (Christmas/Q4) products are removed either way.
    """
    return FilterConfig(
        max_revenue              = max_revenue,
        min_revenue              = 0.0,
        want_seasonal            = want_seasonal,
        f1_remove_amazon         = True,
        f3_min_listing_age_months= _F3_MIN_LISTING_AGE,
        f4_min_rating            = _F4_MIN_RATING,
        f5_min_price             = _F5_MIN_PRICE,
        f6_apply                 = True,
        f6_max_decline_pct       = _F6_MAX_DECLINE_PCT,
        f7_max_reviews           = _F7_MAX_REVIEWS,
        f8_excluded_categories   = [],
        f9_remove_compatibility  = True,
    )


def _price_range_str(config: FilterConfig) -> str:
    """Human-readable description of the F5 minimum unit price (no upper cap)."""
    return f">= GBP {config.f5_min_price:,.2f}  (no upper limit)"


def print_active_filters(config: FilterConfig) -> None:
    """Print the filters that will be applied (informational only, no prompts)."""
    if config.min_revenue > 0:
        rev_desc = (
            f"GBP {config.min_revenue:,.0f} - GBP {config.max_revenue:,.0f}/mo"
            f"  (source: ASIN level revenue)"
        )
    else:
        rev_desc = f"<= GBP {config.max_revenue:,.0f}/mo  (source: ASIN level revenue)"

    seasonal_desc = (
        "Included (summer/unknown kept & flagged; Christmas/Q4 still removed)"
        if config.want_seasonal
        else "Excluded (all seasonal removed)"
    )

    rows = [
        ("Revenue filter",            rev_desc),
        ("F1  Amazon-as-seller",      "Excluded"),
        ("F2  Seasonal products",     seasonal_desc),
        ("F3  Min listing age",       f"{config.f3_min_listing_age_months} months"),
        ("F4  Min rating",            f"> {config.f4_min_rating} stars"),
        ("F5  Min unit price",        _price_range_str(config)),
        ("F6  90-day sales trend",    f"Remove if decline > {abs(config.f6_max_decline_pct):.0f}%"),
        ("F7  Max reviews",           f"{config.f7_max_reviews:,}"),
        ("F8  Regulated categories",  "None excluded"),
        ("F9  Compatibility prods",   "Excluded"),
    ]

    print()
    print("  Built-in filters (applied automatically):")
    print(f"  {'─'*58}")
    for label, value in rows:
        print(f"  {label:<30} {value}")
    print(f"  {'─'*58}")


def _resolve_revenue_source(df: pd.DataFrame) -> Tuple[Optional[str], str]:
    """
    Pick the revenue column. Revenue is read from the ASIN-level field,
    falling back to the generic 'monthly_revenue' then 'parent_revenue' if the
    ASIN-level column is absent. Returns (column_name, human_label).
    """
    preference = ["asin_revenue", "monthly_revenue", "parent_revenue"]
    ideal_label = "ASIN level revenue"
    for col in preference:
        if col in df.columns and df[col].notna().any():
            label = ideal_label if col == preference[0] else f"{ideal_label} -> {col}"
            return col, label
    return None, f"{ideal_label} (no revenue column found)"


def _resolve_sales_source(df: pd.DataFrame) -> Tuple[Optional[str], str]:
    """
    Pick the monthly-sales column. Sales are read from the ASIN-level field,
    falling back to the generic 'monthly_sales' then 'parent_sales' if the
    ASIN-level column is absent. Returns (column_name, human_label).
    """
    preference = ["asin_sales", "monthly_sales", "parent_sales"]
    ideal_label = "ASIN level sales"
    for col in preference:
        if col in df.columns and df[col].notna().any():
            label = ideal_label if col == preference[0] else f"{ideal_label} -> {col}"
            return col, label
    return None, f"{ideal_label} (no sales column found)"


# ──────────────────────────────────────────────────────────────────────────────
# Filter application
# ──────────────────────────────────────────────────────────────────────────────

def apply_filters(
    df: pd.DataFrame,
    config: FilterConfig,
) -> Tuple[pd.DataFrame, List[str]]:
    """
    Apply all built-in filters to *df*.

    Returns
    -------
    filtered_df : pd.DataFrame
    report      : list[str]  — one line per filter showing how many rows removed
    """
    report: List[str] = []
    current = df.copy()
    start_total = len(current)

    def _log(label: str, before: int, after: int) -> None:
        removed = before - after
        report.append(
            f"  {label:<45} {removed:>4} removed  ->  {after:>4} remain"
        )

    # ── Revenue filter (ASIN-level) ──────────────────────────────────────────
    # Revenue is read from the ASIN-level field and copied into 'monthly_revenue'
    # so scoring and output use the same figure. Keep rows where revenue <=
    # budget × 1.52 (optional floor applied only if min_revenue > 0).
    before = len(current)
    src_col, src_label = _resolve_revenue_source(current)
    if src_col is not None:
        current["monthly_revenue"] = current[src_col]

    if config.max_revenue > 0 and "monthly_revenue" in current.columns:
        rev = current["monthly_revenue"]
        mask = rev.notna() & (rev <= config.max_revenue)
        if config.min_revenue > 0:
            mask &= rev >= config.min_revenue
        current = current[mask].reset_index(drop=True)

    floor_note = (
        f"GBP{config.min_revenue:,.0f}-" if config.min_revenue > 0 else "<= "
    )
    _log(
        f"Revenue {floor_note}GBP{config.max_revenue:,.0f}/mo  [{src_label}]",
        before, len(current),
    )

    # ── Sales source (ASIN-level) ────────────────────────────────────────────
    # Monthly sales are read from the ASIN-level field and copied into
    # 'monthly_sales' for scoring and output. No sales floor filter is applied.
    sales_col, _ = _resolve_sales_source(current)
    if sales_col is not None:
        current["monthly_sales"] = current[sales_col]

    # ── F1: Amazon as seller ─────────────────────────────────────────────────
    before = len(current)
    if config.f1_remove_amazon:
        for col in ("seller", "fulfillment"):
            if col in current.columns:
                mask_amazon = current[col].astype(str).str.lower().str.contains(
                    r"\bamazon\b", na=False, regex=True
                )
                current = current[~mask_amazon].reset_index(drop=True)
    _log("F1  Remove Amazon-as-seller", before, len(current))

    # ── F3: Minimum listing age ──────────────────────────────────────────────
    before = len(current)
    if "listing_age_months" in current.columns:
        mask = current["listing_age_months"].notna() & (
            current["listing_age_months"] >= config.f3_min_listing_age_months
        )
        current = current[mask].reset_index(drop=True)
    _log(f"F3  Listing age >= {config.f3_min_listing_age_months} months", before, len(current))

    # ── F4: Minimum rating (strictly greater than threshold) ─────────────────
    before = len(current)
    if "rating" in current.columns:
        mask = current["rating"].notna() & (current["rating"] > config.f4_min_rating)
        current = current[mask].reset_index(drop=True)
    _log(f"F4  Rating > {config.f4_min_rating}", before, len(current))

    # ── F5a: Minimum unit price (built-in, no upper cap) ─────────────────────
    before = len(current)
    if "price" in current.columns:
        price = current["price"]
        mask = price.notna() & (price >= config.f5_min_price)
        current = current[mask].reset_index(drop=True)
    _log(f"F5  Unit price >= GBP {config.f5_min_price:,.2f}", before, len(current))

    # ── F6: 90-day sales trend decline ──────────────────────────────────────
    before = len(current)
    if config.f6_apply and "sales_trend" in current.columns:
        # Keep rows where trend is null (unknown = not declining) OR above threshold
        mask = current["sales_trend"].isna() | (
            current["sales_trend"] >= config.f6_max_decline_pct
        )
        current = current[mask].reset_index(drop=True)
    _log(
        f"F6  90-day trend >= {config.f6_max_decline_pct:.0f}%"
        if config.f6_apply else "F6  Sales trend (column absent, skipped)",
        before, len(current),
    )

    # ── F7: Maximum review count ─────────────────────────────────────────────
    before = len(current)
    if "reviews" in current.columns:
        mask = current["reviews"].notna() & (
            current["reviews"] <= config.f7_max_reviews
        )
        current = current[mask].reset_index(drop=True)
    _log(f"F7  Reviews <= {config.f7_max_reviews:,}", before, len(current))

    # ── F8: Regulated category exclusions (empty list = nothing excluded) ────
    before = len(current)
    if config.f8_excluded_categories and "category" in current.columns:
        excl_lower = [c.lower() for c in config.f8_excluded_categories]
        mask = ~current["category"].astype(str).str.lower().apply(
            lambda v: any(ex in v for ex in excl_lower)
        )
        current = current[mask].reset_index(drop=True)
    _log("F8  Regulated category exclusions", before, len(current))

    # ── F9: Compatibility / accessory products ───────────────────────────────
    before = len(current)
    if config.f9_remove_compatibility and "title" in current.columns:
        compiled = re.compile("|".join(COMPATIBILITY_PATTERNS), re.IGNORECASE)
        mask = ~current["title"].astype(str).apply(lambda t: bool(compiled.search(t)))
        current = current[mask].reset_index(drop=True)
    _log("F9  Remove compatibility products", before, len(current))

    # ── F2: Seasonality (REMOVE / FLAG / PASS) ───────────────────────────────
    # REMOVE (Christmas/Q4) is always dropped. FLAG (summer/unknown) is kept and
    # noted when want_seasonal is True, otherwise dropped. PASS is always kept.
    before = len(current)
    if len(current) > 0:
        verdicts = current.apply(_classify_seasonal, axis=1)
        current["_seasonal_verdict"] = [v for v, _ in verdicts]
        current["seasonal_reason"] = [r for _, r in verdicts]

        if config.want_seasonal:
            keep = current["_seasonal_verdict"] != "remove"
        else:
            keep = current["_seasonal_verdict"] == "pass"
        current = current[keep].reset_index(drop=True)

        # Only the kept FLAG products get highlighted/noted as seasonal.
        current["seasonal_flag"] = current["_seasonal_verdict"] == "flag"
        current["seasonal_reason"] = current["seasonal_reason"].where(
            current["seasonal_flag"], ""
        )
        current.drop(columns=["_seasonal_verdict"], inplace=True)
    else:
        current["seasonal_flag"] = pd.Series(dtype=bool)
        current["seasonal_reason"] = pd.Series(dtype=str)

    _log(
        "F2  Seasonal: drop Christmas/Q4 only (summer/unknown kept & flagged)"
        if config.want_seasonal
        else "F2  Seasonal: drop all seasonal (Christmas/Q4 + summer/unknown)",
        before, len(current),
    )

    report.append(f"  {'─'*58}")
    report.append(
        f"  {'TOTAL':<45} "
        f"{start_total - len(current):>4} removed  ->  {len(current):>4} remain"
    )

    return current, report
