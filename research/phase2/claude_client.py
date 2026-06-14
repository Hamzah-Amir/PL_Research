"""
Claude API client for Phase 2 keyword verification.

All judgement in Phase 2 is done PURELY by Claude — there is no deterministic
relevancy logic and no heuristic fallback (per the project rule). This module
owns the two Claude calls:

  1. `derive_product_profile()` — read the target ASIN's Black Box title +
     category and build a structured product PROFILE (search name, type, key
     attributes, use cases, what-it-is-NOT, brand). This is the relevancy
     anchor. (Text only — the product image is not sent.)
  2. `select_keywords()`        — judge each candidate keyword's relevancy
     against that profile, apply the PDF Step-5 exclusions (branded /
     misspelling / wrong-category / different-product-type / overly-broad /
     question terms), and pick the top 6 by relevancy + search volume, with a
     reason per pick.

Both calls use structured outputs (`messages.parse`) so results are validated
Pydantic objects. Model: Claude Opus 4.8 with adaptive thinking. The stable
profile + candidate-keyword list are prompt-cached so the (future) Step-6 swap
loop reuses the cached prefix.

The API key is loaded from a `.env` file in the project root (git-ignored).
If the key is missing or the API errors, Phase 2 stops and reports it — it
never falls back to a heuristic, because relevancy must come from Claude.
"""

from pathlib import Path
from typing import List, Optional

import pandas as pd
from pydantic import BaseModel, Field

MODEL = "claude-opus-4-8"

# Project root = parent of the phase2/ package directory.
_ROOT = Path(__file__).resolve().parent.parent
_ENV_PATH = _ROOT / ".env"

# How many candidate keywords (SV-ranked) to send to Claude. The full universe
# can be large; the top slice by SV keeps the request bounded while still giving
# Claude far more than the 6 it must choose. Raise if you want a wider net.
MAX_CANDIDATES_TO_MODEL = 200


class Phase2ApiError(RuntimeError):
    """Raised when the Claude API cannot complete a Phase 2 judgement.

    Phase 2 surfaces this to the user rather than guessing a result.
    """


# ──────────────────────────────────────────────────────────────────────────────
# Structured-output schemas
# ──────────────────────────────────────────────────────────────────────────────

class ProductProfile(BaseModel):
    """Structured product identity used as the keyword-relevancy anchor."""
    product_name: str = Field(description="Short, search-style product name (how a shopper would search), 2-6 words.")
    product_type: str = Field(description="Core product type/category in 1-4 words.")
    key_attributes: List[str] = Field(description="Distinguishing attributes that decide relevancy (mechanism, material, size, what's included, manual vs electric, etc.).")
    use_cases: List[str] = Field(description="Primary uses or compatible items a shopper would search for.")
    not_this: List[str] = Field(description="Adjacent or look-alike product types this is NOT — used to reject wrong-product keywords.")
    brand: str = Field(description="The brand name to exclude from keyword picks; empty string if none is identifiable.")


class SelectedKeyword(BaseModel):
    """One chosen launch keyword with Claude's reasoning."""
    keyword: str = Field(description="The exact keyword phrase, copied verbatim from the candidate list.")
    search_volume: int = Field(description="The keyword's monthly search volume from the candidate list.")
    relevancy: str = Field(description="Relevancy to the product: one of 'high', 'medium', or 'low'.")
    reason: str = Field(description="One sentence: why this keyword was chosen (relevancy + search volume).")


class KeywordSelection(BaseModel):
    """The final picked set plus a short overall note."""
    keywords: List[SelectedKeyword]
    note: str = Field(description="One or two sentences on the overall selection rationale or any caveat.")


# ──────────────────────────────────────────────────────────────────────────────
# Client construction (.env key loading)
# ──────────────────────────────────────────────────────────────────────────────

def load_client():
    """
    Build an Anthropic client, loading the API key from the project-root `.env`.

    Raises Phase2ApiError with actionable guidance if the SDK or key is missing
    — there is no fallback path.
    """
    try:
        from dotenv import load_dotenv  # python-dotenv
    except ImportError as e:
        raise Phase2ApiError(
            "python-dotenv is not installed. Run setup again (pip install -r "
            "requirements.txt)."
        ) from e

    try:
        import anthropic
    except ImportError as e:
        raise Phase2ApiError(
            "The 'anthropic' package is not installed. Run setup again "
            "(pip install -r requirements.txt)."
        ) from e

    # Load .env from the project root (does not override an already-set env var).
    load_dotenv(dotenv_path=_ENV_PATH)

    import os
    if not os.environ.get("ANTHROPIC_API_KEY"):
        raise Phase2ApiError(
            f"ANTHROPIC_API_KEY not found. Create a .env file at {_ENV_PATH} "
            f"containing:  ANTHROPIC_API_KEY=sk-ant-...   (the file is git-ignored)."
        )

    try:
        return anthropic.Anthropic()
    except Exception as e:  # noqa: BLE001 - surface any construction failure
        raise Phase2ApiError(f"Could not initialise the Anthropic client: {e}") from e


# ──────────────────────────────────────────────────────────────────────────────
# Call 1 — build the product profile (text only)
# ──────────────────────────────────────────────────────────────────────────────

_PROFILE_SYSTEM = (
    "You are an Amazon product analyst. You are given a product's listing title "
    "and its category. Read them carefully — the title reveals the product type, "
    "mechanism, contents, and materials.\n\n"
    "Build a structured product profile that will be used to judge whether search "
    "keywords are relevant to THIS product:\n"
    "  - product_name: how a shopper would actually search for it (no brand, no "
    "marketing adjectives).\n"
    "  - product_type: the core type/category.\n"
    "  - key_attributes: the distinguishing features that separate this product "
    "from look-alikes (mechanism such as manual vs electric, material, size, "
    "what is included, form factor).\n"
    "  - use_cases: the main uses or compatible items shoppers search for.\n"
    "  - not_this: adjacent or look-alike product types this is NOT, so wrong-"
    "product keywords can be rejected.\n"
    "  - brand: the brand name (to exclude from keyword picks); empty if none.\n\n"
    "Do not invent attributes that are not supported by the text. If something "
    "cannot be determined, leave it out rather than guessing."
)


def derive_product_profile(
    client,
    title: str,
    category: str = "",
    subcategory: str = "",
) -> ProductProfile:
    """
    Build a structured product profile from the title + category (text only).

    Raises Phase2ApiError on any API failure — no heuristic fallback.
    """
    title = (title or "").strip()
    if not title:
        raise Phase2ApiError(
            "No product title was found for the target ASIN in the Black Box "
            "file — cannot build a product profile without it."
        )

    cat_line = " > ".join([c for c in (category, subcategory) if c and str(c).strip()])
    text = f"Listing title:\n{title}"
    if cat_line:
        text += f"\n\nCategory: {cat_line}"

    try:
        response = client.messages.parse(
            model=MODEL,
            max_tokens=2500,
            thinking={"type": "adaptive"},
            output_config={"effort": "medium"},
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


# ──────────────────────────────────────────────────────────────────────────────
# Call 2 — select the top keywords
# ──────────────────────────────────────────────────────────────────────────────

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
    "  - Copy each chosen keyword phrase and its search volume VERBATIM from the "
    "candidate list. Never invent a keyword or a number.\n"
    "  - If you genuinely cannot find enough relevant keywords, return fewer and "
    "explain why in the note rather than padding with weak picks.\n"
    "  - Give a one-sentence reason per pick citing relevancy and search volume."
)


def _format_profile(p: ProductProfile) -> str:
    """Render the product profile as the relevancy anchor block."""
    def _join(items: List[str]) -> str:
        return ", ".join(i for i in items if i and str(i).strip()) or "(none)"

    lines = [
        "PRODUCT PROFILE (judge keyword relevancy against THIS):",
        f"  Name: {p.product_name}",
        f"  Type: {p.product_type}",
        f"  Key attributes: {_join(p.key_attributes)}",
        f"  Use cases: {_join(p.use_cases)}",
        f"  NOT this product: {_join(p.not_this)}",
        f"  Brand to exclude from keywords: {p.brand or '(none)'}",
    ]
    return "\n".join(lines)


def _format_candidates(candidates: pd.DataFrame, limit: int) -> str:
    """Render the SV-ranked candidate pool as a compact text table for the model."""
    cols = [c for c in ("keyword", "search_volume", "competing_products", "cerebro_iq")
            if c in candidates.columns]
    head = candidates.head(limit)
    lines = []
    for _, row in head.iterrows():
        parts = [f"{row.get('keyword', '')}"]
        if "search_volume" in cols:
            sv = row.get("search_volume")
            parts.append(f"SV={int(sv) if pd.notna(sv) else 'NA'}")
        if "competing_products" in cols:
            cp = row.get("competing_products")
            if pd.notna(cp):
                parts.append(f"competing={int(cp)}")
        if "cerebro_iq" in cols:
            iq = row.get("cerebro_iq")
            if pd.notna(iq):
                parts.append(f"IQ={int(iq)}")
        lines.append("  - " + "  ".join(parts))
    return "\n".join(lines)


def select_keywords(
    client,
    profile: ProductProfile,
    candidates: pd.DataFrame,
    target_count: int = 6,
    exclude: Optional[List[str]] = None,
    max_candidates: int = MAX_CANDIDATES_TO_MODEL,
) -> KeywordSelection:
    """
    Have Claude pick the top *target_count* launch keywords from *candidates*,
    judged against the product *profile*.

    *exclude* is the list of phrases the user rejected in the (future) Step-6
    approval loop; Claude is told to pick replacements and avoid them. The
    profile + candidate list are prompt-cached so repeated swap rounds re-use the
    cached prefix.

    Raises Phase2ApiError on any API failure — no heuristic fallback.
    """
    if candidates.empty:
        raise Phase2ApiError("No candidate keywords remain to select from.")

    candidate_block = _format_candidates(candidates, max_candidates)

    # Stable, cacheable context: rules + product profile + candidate list.
    system = [
        {"type": "text", "text": _SELECTION_RULES, "cache_control": {"type": "ephemeral"}},
        {
            "type": "text",
            "text": (
                f"{_format_profile(profile)}\n\n"
                f"CANDIDATE KEYWORDS (already de-duplicated and filtered to "
                f"SV >= 100, ranked by search volume):\n{candidate_block}"
            ),
            "cache_control": {"type": "ephemeral"},
        },
    ]

    # Volatile instruction (target count + any rejected phrases) goes last so it
    # does not invalidate the cached prefix across swap rounds.
    instruction = (
        f"Select the {target_count} best launch keywords for this product, "
        f"following the selection criteria and exclusions."
    )
    if exclude:
        rejected = ", ".join(f'"{k}"' for k in exclude)
        instruction += (
            f"\n\nThe user rejected these previously-suggested keywords — do NOT "
            f"select them again; choose the next-best relevant replacements: {rejected}."
        )

    try:
        response = client.messages.parse(
            model=MODEL,
            max_tokens=8000,
            thinking={"type": "adaptive"},
            output_config={"effort": "high"},
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
