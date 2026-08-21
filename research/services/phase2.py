"""
PL Research Tool — Phase 2 (web / library form).

A single, self-contained module for keyword verification, mirroring the Phase 1
treatment (research/services/phase1.py): the target product comes from the
Django ORM (`products.Product`) by ASIN, all judgement is done by Claude (folded
in here), and every entry point returns plain data — no CLI, no prompts, no
interactive approve/swap loop. The web layer drives approval by re-calling with
`exclude`.

Keyword candidates come from a manually uploaded H10 Xray keyword export
(sample format: test_file/xray_keywords_tent_pegs.xlsx) — passed as a file path
or a file-like object (e.g. a Django UploadedFile).

Public API
----------
    run_phase2(asin, xray_file, *, target_count=6, exclude=None) -> dict
        High-level call. Returns:
        {
          "target_asin", "target_product",      # from the ORM
          "profile",            # Claude product profile (dict)
          "main_search_term",   # precise title-derived term — drives Phase 3 search
          "candidates",         # prepared SV-ranked keyword pool (list of dicts)
          "prep_report",        # human-readable prep steps (list of str)
          "selection",          # Claude's chosen keywords (dict) or None
          "js_autogen",         # in-flight Phase 3 JS-sheet job handle or None
          "error",              # set (str) if a step failed; else None
        }

Lower-level helpers (`derive_product_profile`, `select_keywords`,
`prepare_candidates`, `read_xray_keywords`) are exposed too; Phase 3 reuses
`ProductProfile` / `derive_product_profile` from here.
"""

from __future__ import annotations

import os
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
from django.conf import settings
from pydantic import BaseModel, Field

from products.models import Product

MODEL = "claude-opus-4-8"
MAX_CANDIDATES_TO_MODEL = 200   # SV-ranked slice sent to Claude
MIN_SEARCH_VOLUME = 100         # the one objective PDF Step-5 numeric rule
TARGET_KEYWORD_COUNT = 6


class Phase2ApiError(RuntimeError):
    """Raised when the Claude API cannot complete a Phase 2 judgement.
    Phase 2 surfaces this rather than guessing (no heuristic fallback)."""


# ── Structured-output schemas ────────────────────────────────────────────────
class ProductProfile(BaseModel):
    """Structured product identity used as the keyword-relevancy anchor."""
    product_name: str = Field(description="Short, search-style product name (how a shopper would search), 2-6 words.")
    product_type: str = Field(description="Core product type/category in 1-4 words (general, e.g. 'tent pegs').")
    main_search_term: str = Field(description="The SINGLE most specific high-intent search phrase a buyer would type to find THIS exact product — the core product noun PLUS its strongest defining/quality modifier, taken mainly from the title. Precise, not generic: 'heavy duty tent pegs' NOT 'tent pegs'; 'electric tyre pump' NOT 'pump'. 2-4 words, lowercase, no brand. Searching it on Amazon must surface the products most specific to this one.")
    key_attributes: List[str] = Field(description="Distinguishing attributes that decide relevancy (mechanism, material, size, what's included, manual vs electric, etc.).")
    use_cases: List[str] = Field(description="Primary uses or compatible items a shopper would search for.")
    not_this: List[str] = Field(description="Adjacent or look-alike product types this is NOT — used to reject wrong-product keywords.")
    brand: str = Field(description="The brand name to exclude from keyword picks; empty string if none is identifiable.")


class SelectedKeyword(BaseModel):
    keyword: str = Field(description="The exact keyword phrase, copied verbatim from the candidate list.")
    search_volume: int = Field(description="The keyword's monthly search volume from the candidate list.")
    relevancy: str = Field(description="Relevancy to the product: one of 'high', 'medium', or 'low'.")
    reason: str = Field(description="One sentence: why this keyword was chosen (relevancy + search volume).")


class KeywordSelection(BaseModel):
    keywords: List[SelectedKeyword]
    note: str = Field(description="One or two sentences on the overall selection rationale or any caveat.")


# ── Claude client ────────────────────────────────────────────────────────────
def load_client():
    """Build an Anthropic client from the project-root .env. Raises Phase2ApiError
    (with actionable guidance) if the SDK or key is missing — no fallback path."""
    try:
        from dotenv import load_dotenv
    except ImportError as e:
        raise Phase2ApiError("python-dotenv is not installed (pip install -r requirements.txt).") from e
    try:
        import anthropic
    except ImportError as e:
        raise Phase2ApiError("The 'anthropic' package is not installed (pip install -r requirements.txt).") from e

    load_dotenv(dotenv_path=Path(settings.BASE_DIR) / ".env")
    if not os.environ.get("ANTHROPIC_API_KEY"):
        raise Phase2ApiError("ANTHROPIC_API_KEY not found in .env (the file is git-ignored).")
    try:
        return anthropic.Anthropic()
    except Exception as e:  # noqa: BLE001
        raise Phase2ApiError(f"Could not initialise the Anthropic client: {e}") from e


# ── Call 1 — build the product profile (text only) ──────────────────────────
_PROFILE_SYSTEM = (
  "You are an Amazon product analyst. You are given a product's listing title "
"and its category. Read them carefully — the title reveals the product type, "
"mechanism, contents, and materials.\n\n"
"Build a structured product profile used to judge whether search keywords are "
"relevant to THIS product:\n"
"  - product_name: how a shopper would actually search for it (no brand, no marketing adjectives).\n"
"  - product_type: the core type/category (general, e.g. 'tent pegs').\n"
"  - main_search_term: the SINGLE most specific high-intent keyword a buyer would "
"search to find THIS exact product — the core product noun PLUS its strongest "
"defining/quality modifier, taken mainly from the title. Precise, never generic: "
"'heavy duty tent pegs' NOT 'tent pegs'; 'electric tyre pump' NOT 'pump'. 2-4 words, "
"lowercase, no brand. It must be specific enough that searching it on Amazon returns "
"the products most specific to this one.\n"
"  - key_attributes: distinguishing features that separate this product from look-alikes.\n"
"  - use_cases: the main uses or compatible items shoppers search for.\n"
"  - not_this: adjacent or look-alike product types this is NOT.\n"
"  - brand: the brand name (to exclude from keyword picks); empty if none.\n\n"
"Ground every field in the title and category text. Do not invent attributes. "
"If something cannot be determined, leave it out rather than guessing."
)


def derive_product_profile(client, title: str, category: str = "", subcategory: str = "") -> ProductProfile:
    """Build a structured product profile from the title + category (text only).
    Raises Phase2ApiError on any API failure — no heuristic fallback."""
    title = (title or "").strip()
    if not title:
        raise Phase2ApiError("No product title was found for the target ASIN — cannot build a profile.")

    cat_line = " > ".join(c for c in (category, subcategory) if c and str(c).strip())
    text = f"Listing title:\n{title}" + (f"\n\nCategory: {cat_line}" if cat_line else "")

    try:
        response = client.messages.parse(
            model=MODEL, max_tokens=2500,
            thinking={"type": "adaptive"}, output_config={"effort": "medium"},
            system=_PROFILE_SYSTEM,
            messages=[{"role": "user", "content": text}],
            output_format=ProductProfile,
        )
    except Exception as e:  # noqa: BLE001
        raise Phase2ApiError(f"Claude call failed while building the product profile: {e}") from e

    parsed = response.parsed_output
    if parsed is None:
        raise Phase2ApiError("Claude returned no parseable product profile.")
    return parsed


# ── Call 2 — select the top keywords ─────────────────────────────────────────
_SELECTION_RULES = (
    "You are an Amazon Private Label keyword strategist. From a list of candidate "
    "keywords for a specific product, select the best keywords to launch with.\n\n"
    "Selection criteria (PDF workflow Step 5):\n"
    "  - Pick by a combination of HIGH RELEVANCY to the product and HIGH search volume.\n"
    "  - Relevancy is judged against the product PROFILE provided — a keyword must "
    "describe THIS product (matching its type and key attributes), not merely a "
    "related, adjacent, or look-alike one.\n\n"
    "EXCLUDE a keyword if it is any of the following:\n"
    "  - the product's own brand name, or a branded competitor name\n"
    "  - a misspelling\n"
    "  - wrong-category (does not belong to this product's category)\n"
    "  - describes a DIFFERENT product type (see the profile's 'NOT this' list)\n"
    "  - overly broad / generic (would attract untargeted traffic)\n"
    "  - a question keyword (how/what/why/where/when/which/can/does/is/are ...)\n\n"
    "Rules:\n"
    "  - Copy each chosen keyword phrase and its search volume VERBATIM from the candidate list. Never invent a keyword or a number.\n"
    "  - If you genuinely cannot find enough relevant keywords, return fewer and explain why in the note rather than padding with weak picks.\n"
    "  - Give a one-sentence reason per pick citing relevancy and search volume."
)


def _format_profile(p: ProductProfile) -> str:
    def _join(items: List[str]) -> str:
        return ", ".join(i for i in items if i and str(i).strip()) or "(none)"
    return "\n".join([
        "PRODUCT PROFILE (judge keyword relevancy against THIS):",
        f"  Name: {p.product_name}",
        f"  Type: {p.product_type}",
        f"  Key attributes: {_join(p.key_attributes)}",
        f"  Use cases: {_join(p.use_cases)}",
        f"  NOT this product: {_join(p.not_this)}",
        f"  Brand to exclude from keywords: {p.brand or '(none)'}",
    ])


def _format_candidates(candidates: pd.DataFrame, limit: int) -> str:
    cols = [c for c in ("keyword", "search_volume", "competing_products", "cerebro_iq")
            if c in candidates.columns]
    lines = []
    for _, row in candidates.head(limit).iterrows():
        parts = [f"{row.get('keyword', '')}"]
        if "search_volume" in cols and pd.notna(row.get("search_volume")):
            parts.append(f"SV={int(row.get('search_volume'))}")
        if "competing_products" in cols and pd.notna(row.get("competing_products")):
            parts.append(f"competing={int(row.get('competing_products'))}")
        if "cerebro_iq" in cols and pd.notna(row.get("cerebro_iq")):
            parts.append(f"IQ={int(row.get('cerebro_iq'))}")
        lines.append("  - " + "  ".join(parts))
    return "\n".join(lines)


def select_keywords(client, profile: ProductProfile, candidates: pd.DataFrame,
                    target_count: int = TARGET_KEYWORD_COUNT, exclude: Optional[List[str]] = None,
                    max_candidates: int = MAX_CANDIDATES_TO_MODEL) -> KeywordSelection:
    """Have Claude pick the top *target_count* launch keywords from *candidates*,
    judged against *profile*. *exclude* = phrases the user rejected (the web swap
    step), so Claude picks replacements. Raises Phase2ApiError on API failure."""
    if candidates.empty:
        raise Phase2ApiError("No candidate keywords remain to select from.")

    candidate_block = _format_candidates(candidates, max_candidates)
    system = [
        {"type": "text", "text": _SELECTION_RULES, "cache_control": {"type": "ephemeral"}},
        {"type": "text",
         "text": (f"{_format_profile(profile)}\n\nCANDIDATE KEYWORDS (already de-duplicated and "
                  f"filtered to SV >= {MIN_SEARCH_VOLUME}, ranked by search volume):\n{candidate_block}"),
         "cache_control": {"type": "ephemeral"}},
    ]
    instruction = (f"Select the {target_count} best launch keywords for this product, "
                   f"following the selection criteria and exclusions.")
    if exclude:
        rejected = ", ".join(f'"{k}"' for k in exclude)
        instruction += (f"\n\nThe user rejected these previously-suggested keywords — do NOT select "
                        f"them again; choose the next-best relevant replacements: {rejected}.")

    try:
        response = client.messages.parse(
            model=MODEL, max_tokens=8000,
            thinking={"type": "adaptive"}, output_config={"effort": "high"},
            system=system,
            messages=[{"role": "user", "content": instruction}],
            output_format=KeywordSelection,
        )
    except Exception as e:  # noqa: BLE001
        raise Phase2ApiError(f"Claude call failed while selecting keywords: {e}") from e

    parsed = response.parsed_output
    if parsed is None:
        raise Phase2ApiError("Claude returned no parseable keyword selection.")
    return parsed


# ── Xray keyword file reader (path or file-like) ─────────────────────────────
_COLUMN_MAPPINGS: Dict[str, List[str]] = {
    "keyword": ["Keyword Phrase", "Keyword", "Keywords", "Phrase", "Search Term"],
    "cerebro_iq": ["Cerebro IQ Score", "Cerebro IQ", "IQ Score"],
    "search_volume": ["Search Volume", "Search Vol", "SV", "Monthly Search Volume"],
    "search_volume_trend": ["Search Volume Trend", "Search Vol Trend", "SV Trend"],
    "suggested_ppc_bid": ["Suggested PPC Bid", "Sponsored PPC Bid", "PPC Bid", "Suggested Bid"],
    "keyword_sales": ["Keyword Sales", "Est. Keyword Sales", "Keyword Sales (est)"],
    "competing_products": ["Competing Products", "Competing Products Count", "# Competing Products"],
    "title_density": ["Title Density", "Title Density Exact", "Title Density (exact)"],
    "competitor_rank": ["Competitor Rank (avg)", "Competitor Rank Avg", "Competitor Rank", "Avg Competitor Rank"],
}
_NUMERIC_COLUMNS = ["cerebro_iq", "search_volume", "search_volume_trend", "keyword_sales",
                    "competing_products", "title_density", "competitor_rank"]


def _find_column(df: pd.DataFrame, canonical: str) -> Optional[str]:
    aliases = _COLUMN_MAPPINGS.get(canonical, [])
    df_lower = {c.lower().strip(): c for c in df.columns}
    for alias in aliases:
        if alias in df.columns:
            return alias
        if alias.lower().strip() in df_lower:
            return df_lower[alias.lower().strip()]
    for col in df.columns:
        col_strip = str(col).lower().strip()
        for alias in aliases:
            if alias.lower() in col_strip:
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
    s = (s.replace(",", "").replace("$", "").replace("£", "").replace("€", "")
          .replace("%", "").replace("K", "000").replace("k", "000"))
    try:
        return float(s)
    except (ValueError, TypeError):
        return None


def _read_excel_any(file, skiprows=0):
    """pd.read_excel for a path or a file-like object (rewinds file-likes)."""
    if hasattr(file, "read"):
        try:
            file.seek(0)
        except Exception:  # noqa: BLE001
            pass
        return pd.read_excel(file, engine="openpyxl", skiprows=skiprows)
    return pd.read_excel(file, engine="openpyxl", skiprows=skiprows)


def read_xray_keywords(xray_file) -> Tuple[pd.DataFrame, int]:
    """Read an H10 Xray keyword XLSX/CSV (path str/Path or file-like) into a
    normalised DataFrame with canonical column names. Returns (df, row_count)."""
    is_path = not hasattr(xray_file, "read")
    suffix = Path(str(xray_file)).suffix.lower() if is_path else \
        Path(str(getattr(xray_file, "name", ""))).suffix.lower()

    if is_path:
        path = Path(str(xray_file).strip().strip('"').strip("'"))
        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")
        xray_file = path
        suffix = path.suffix.lower()

    if suffix in (".xlsx", ".xls", ""):
        df_raw = None
        for skip in (0, 1, 2):
            try:
                tmp = _read_excel_any(xray_file, skiprows=skip)
                if any("keyword" in str(c).lower() or "phrase" in str(c).lower()
                       or "search term" in str(c).lower() for c in tmp.columns):
                    df_raw = tmp
                    break
            except Exception:  # noqa: BLE001
                pass
        if df_raw is None:
            df_raw = _read_excel_any(xray_file)
    elif suffix == ".csv":
        df_raw = pd.read_csv(xray_file)
    else:
        raise ValueError(f"Unsupported format '{suffix}'. Supply .xlsx, .xls, or .csv")

    cols_used: Dict[str, str] = {}
    data: Dict[str, pd.Series] = {}
    for canonical in _COLUMN_MAPPINGS:
        actual = _find_column(df_raw, canonical)
        if actual and actual not in cols_used.values():
            cols_used[canonical] = actual
            data[canonical] = df_raw[actual].reset_index(drop=True)

    df = pd.DataFrame(data)
    mapped = set(cols_used.values())
    for col in df_raw.columns:
        if col not in mapped:
            df[f"raw_{re.sub(r'[^a-zA-Z0-9_]', '_', str(col))}"] = df_raw[col].reset_index(drop=True)

    for col in _NUMERIC_COLUMNS:
        if col in df.columns:
            df[col] = df[col].apply(_parse_numeric)

    if "keyword" in df.columns:
        df = df[df["keyword"].notna()]
        df = df[df["keyword"].astype(str).str.strip() != ""].reset_index(drop=True)

    return df, len(df)


# ── Deterministic candidate prep (de-dupe + SV floor) ────────────────────────
def _normalize_phrase(value) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip().lower())


def prepare_candidates(df: pd.DataFrame, min_search_volume: int = MIN_SEARCH_VOLUME) -> Tuple[pd.DataFrame, List[str]]:
    """Objective Step-5 prep (no judgement): drop blanks, de-dup (keep highest SV),
    drop SV < floor, rank by SV. Returns (candidates, human-readable report)."""
    report: List[str] = []

    def _log(label, before, after):
        report.append(f"{label}: {before - after} dropped -> {after} remain")

    if df.empty:
        return df.copy(), ["(no keywords to prepare)"]

    current = df.copy().reset_index(drop=True)
    start_total = len(current)

    before = len(current)
    if "keyword" in current.columns:
        current = current[current["keyword"].notna()]
        current["_norm"] = current["keyword"].apply(_normalize_phrase)
        current = current[current["_norm"] != ""].reset_index(drop=True)
    else:
        current["_norm"] = ""
    _log("Drop blank phrases", before, len(current))

    before = len(current)
    if "search_volume" in current.columns:
        current = current.sort_values("search_volume", ascending=False, na_position="last")
    current = current.drop_duplicates("_norm", keep="first").reset_index(drop=True)
    _log("De-duplicate phrases", before, len(current))

    before = len(current)
    if "search_volume" in current.columns:
        sv = current["search_volume"]
        current = current[sv.notna() & (sv >= min_search_volume)].reset_index(drop=True)
        _log(f"Search volume >= {min_search_volume}", before, len(current))
    else:
        report.append("Search volume floor (column absent, skipped)")

    if "search_volume" in current.columns and not current.empty:
        current = current.sort_values("search_volume", ascending=False, na_position="last").reset_index(drop=True)
    current.drop(columns=["_norm"], inplace=True, errors="ignore")

    report.append(f"PREPARED: {start_total - len(current)} dropped -> {len(current)} remain")
    return current, report


def _df_records(df: pd.DataFrame) -> List[Dict]:
    """DataFrame -> list of dicts with NaN replaced by None (JSON/template safe)."""
    if df.empty:
        return []
    return df.where(pd.notna(df), None).to_dict(orient="records")


# ── High-level entry point ───────────────────────────────────────────────────
def run_phase2(asin: str, xray_file, *, target_count: int = TARGET_KEYWORD_COUNT,
               exclude: Optional[List[str]] = None, autostart_phase3: bool = True) -> Dict:
    """Run Phase 2 end-to-end for one selected ASIN and return data for the
    frontend. The target's title/category come from the Product ORM; keyword
    candidates from the uploaded Xray file. Re-call with `exclude` to swap out
    rejected keywords. Errors are returned in the `error` key (never raised)."""
    asin = (asin or "").strip().upper()
    out: Dict = {
        "target_asin": asin, "target_product": None, "profile": None,
        "main_search_term": None, "candidates": [], "prep_report": [],
        "selection": None, "js_autogen": None, "error": None,
    }

    # Product details come from the ORM (the normal path). DEV FALLBACK: if the
    # ASIN isn't in the DB (e.g. a directly-entered ASIN that was never hunted),
    # pull title/category from Keepa (1 token) so it still reaches keyword
    # research instead of dead-ending on "not in the product database".
    product = Product.objects.filter(asin=asin).first()
    if product:
        title = product.title or ""
        category = product.category or ""
        subcategory = product.subcategory or ""
        source = "db"
    else:
        try:
            from research.services.phase3 import (
                load_keepa, query_products, extract_keepa_fields)
            kp = query_products(load_keepa(), [asin], deep=False).get(asin)
        except Exception as e:  # noqa: BLE001
            out["error"] = (f"ASIN {asin} is not in the product database, and the "
                            f"Keepa fallback failed: {e}")
            return out
        if not kp:
            out["error"] = (f"ASIN {asin} is not in the product database, and Keepa "
                            f"returned no product for it.")
            return out
        kf = extract_keepa_fields(kp)
        title = kf.get("title") or ""
        category = kf.get("category") or ""
        subcategory = kf.get("subcategory") or ""
        source = "keepa"

    if not title:
        out["error"] = f"No title available for ASIN {asin} — cannot build the keyword profile."
        return out
    out["target_product"] = {
        "asin": asin, "title": title, "category": category,
        "subcategory": subcategory, "source": source,
    }

    try:
        client = load_client()
        profile = derive_product_profile(
            client, title=title, category=category, subcategory=subcategory,
        )
    except Phase2ApiError as e:
        out["error"] = str(e)
        return out

    out["profile"] = profile.model_dump()
    # The precise, title-derived term (falls back to the general type) — this is
    # what generates the Phase 3 JS sheet, so it must map to the most specific
    # competitors, not a generic category.
    main_term = (getattr(profile, "main_search_term", "") or profile.product_type or "").strip()
    out["main_search_term"] = main_term

    # Kick off the Phase 3 competitor-sheet generation in the background now that
    # the main search term is known (no waiting later). Best-effort.
    if autostart_phase3 and main_term:
        try:
            from research.services.phase3 import start_background
            out["js_autogen"] = {"keyword": main_term, "job": start_background(main_term)}
        except Exception:  # noqa: BLE001 — Phase 3 can fall back to manual upload
            out["js_autogen"] = None

    # Read + prepare the uploaded Xray keyword file.
    try:
        kw_df, _ = read_xray_keywords(xray_file)
    except (FileNotFoundError, ValueError) as e:
        out["error"] = f"Could not read the Xray keyword file: {e}"
        return out
    if kw_df.empty:
        out["error"] = "No keywords could be read from the Xray keyword file."
        return out

    candidates, report = prepare_candidates(kw_df)
    out["candidates"] = _df_records(candidates)
    out["prep_report"] = report
    if candidates.empty:
        out["error"] = f"No candidate keywords remain after SV >= {MIN_SEARCH_VOLUME} filtering."
        return out

    try:
        selection = select_keywords(client, profile, candidates, target_count=target_count, exclude=exclude)
    except Phase2ApiError as e:
        out["error"] = str(e)
        return out

    out["selection"] = selection.model_dump()
    return out
