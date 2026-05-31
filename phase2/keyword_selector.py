"""
Deterministic keyword preparation for Phase 2 (Step 5, objective part only).

By design, ALL keyword judgement — relevancy to the product, plus the Step-5
exclusions (branded competitor names, misspellings, wrong-category terms,
different-product-type terms, overly broad terms, question keywords) and the
final pick of 6 — is performed PURELY by Claude via the API
(`claude_client.select_keywords`). There is no deterministic relevancy logic
and no heuristic fallback anywhere in this module.

What stays deterministic here is only objective data hygiene that needs no
judgement:
  * de-duplicate identical phrases (keep the highest-SV instance), and
  * drop SV < 100 (the one explicit numeric rule in the PDF).

The surviving, SV-ranked pool is the candidate set handed to Claude.
"""

import re
from typing import List, Tuple

import pandas as pd

# The single objective numeric rule from the PDF Step 5.
MIN_SEARCH_VOLUME = 100
# The PDF locks a final set of this many keywords.
TARGET_KEYWORD_COUNT = 6


def _normalize_phrase(value) -> str:
    """Lower-case + collapse whitespace — for de-duplication only."""
    s = str(value or "").strip().lower()
    return re.sub(r"\s+", " ", s)


def prepare_candidates(
    df: pd.DataFrame,
    min_search_volume: int = MIN_SEARCH_VOLUME,
) -> Tuple[pd.DataFrame, List[str]]:
    """
    Objective Step-5 preparation of the Cerebro keyword universe.

    Steps (no judgement, no relevancy):
      1. Drop blank phrases.
      2. De-duplicate on the normalised phrase (keep the highest-SV instance).
      3. Drop SV < *min_search_volume* (skipped only if no SV column exists).
      4. Rank by search volume descending.

    Returns
    -------
    candidates : pd.DataFrame
        The prepared, SV-ranked pool handed to Claude for relevancy + selection.
    report : list[str]
        One human-readable line per step (mirrors Phase 1's report style).
    """
    report: List[str] = []

    def _log(label: str, before: int, after: int) -> None:
        report.append(f"  {label:<45} {before - after:>4} dropped  ->  {after:>4} remain")

    if df.empty:
        report.append("  (no keywords to prepare)")
        return df.copy(), report

    current = df.copy().reset_index(drop=True)
    start_total = len(current)

    # ── 1. Drop blank phrases ─────────────────────────────────────────────────
    before = len(current)
    if "keyword" in current.columns:
        current = current[current["keyword"].notna()]
        current["_norm"] = current["keyword"].apply(_normalize_phrase)
        current = current[current["_norm"] != ""].reset_index(drop=True)
    else:
        current["_norm"] = ""
    _log("Drop blank phrases", before, len(current))

    # ── 2. De-duplicate (keep highest SV) ─────────────────────────────────────
    before = len(current)
    if "search_volume" in current.columns:
        current = current.sort_values("search_volume", ascending=False, na_position="last")
    current = current.drop_duplicates("_norm", keep="first").reset_index(drop=True)
    _log("De-duplicate phrases", before, len(current))

    # ── 3. SV floor (the one objective PDF rule) ──────────────────────────────
    before = len(current)
    if "search_volume" in current.columns:
        sv = current["search_volume"]
        current = current[sv.notna() & (sv >= min_search_volume)].reset_index(drop=True)
        _log(f"Search volume >= {min_search_volume:,}", before, len(current))
    else:
        report.append("  Search volume floor (column absent, skipped)")

    # ── 4. Rank by SV descending ──────────────────────────────────────────────
    if "search_volume" in current.columns and not current.empty:
        current = current.sort_values(
            "search_volume", ascending=False, na_position="last"
        ).reset_index(drop=True)

    current.drop(columns=["_norm"], inplace=True, errors="ignore")

    report.append(f"  {'─'*58}")
    report.append(
        f"  {'PREPARED':<45} {start_total - len(current):>4} dropped  ->  "
        f"{len(current):>4} remain"
    )

    return current, report
