"""
Claude API client for Phase 3 competitor analysis.

All judgement is Claude's; deterministic data ops stay in Python (Rule 0 — no
heuristic fallback). This module owns three calls:

  1. `filter_same_product_type()` — given the target product profile and the
     candidate ASIN pool (titles), decide which competitors are the SAME product
     type as the target (PDF Step 8 same-product-type filter). Returns the kept
     ASINs with a short reason.
  2. `score_listing()` — the rubric elements that need judgement, scored per
     competitor: product images (vision on the listing image URLs), title,
     bullets/description, unique design, marketing images, and pricing strategy.
     Returns 0-N scores capped at each element's max, with a reason each.
  3. `analyze_reviews()` — top-3 weaknesses (from 1-3★) + solutions and top-3
     strengths (from 4-5★). Needs review data; if none is supplied it returns an
     "Unknown — need user confirmation" result rather than inventing complaints.

The product profile is built by `phase2.claude_client.derive_product_profile`
(reused — same anchor concept). Model: Claude Opus 4.8, adaptive thinking,
structured outputs via `messages.parse`. Key from the project-root `.env`.
"""

from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
from pydantic import BaseModel, Field

MODEL = "claude-opus-4-8"
_ROOT = Path(__file__).resolve().parent.parent
_ENV_PATH = _ROOT / ".env"

# Reuse Phase 2's profile builder + schema (same relevancy-anchor concept).
from research.phase2.claude_client import ProductProfile, derive_product_profile  # noqa: E402,F401

# Max points per judgement-scored rubric element (the 184-point rubric).
LISTING_ELEMENT_MAX = {
    "product_images": 20,
    "product_title": 5,
    "bullets_description": 5,
    "unique_design": 2,
    "marketing_images": 15,
    "pricing_strategy": 10,
}


class Phase3ApiError(RuntimeError):
    """Raised when the Claude API cannot complete a Phase 3 judgement."""


# ──────────────────────────────────────────────────────────────────────────────
# Client construction (.env key)
# ──────────────────────────────────────────────────────────────────────────────

def load_client():
    """Build an Anthropic client from the project-root `.env`. No fallback."""
    try:
        from dotenv import load_dotenv
    except ImportError as e:
        raise Phase3ApiError("python-dotenv is not installed. Run setup again.") from e
    try:
        import anthropic
    except ImportError as e:
        raise Phase3ApiError("The 'anthropic' package is not installed. Run setup again.") from e

    load_dotenv(dotenv_path=_ENV_PATH)
    import os
    if not os.environ.get("ANTHROPIC_API_KEY"):
        raise Phase3ApiError(
            f"ANTHROPIC_API_KEY not found. Add it to {_ENV_PATH} (git-ignored)."
        )
    try:
        return anthropic.Anthropic()
    except Exception as e:  # noqa: BLE001
        raise Phase3ApiError(f"Could not initialise the Anthropic client: {e}") from e


# ──────────────────────────────────────────────────────────────────────────────
# Call 1 — same-product-type filter (Step 8)
# ──────────────────────────────────────────────────────────────────────────────

class SameTypeResult(BaseModel):
    same_type_asins: List[str] = Field(
        description="ASINs from the candidate list that are the SAME core product type as the target. Copy ASINs verbatim."
    )
    note: str = Field(description="One or two sentences on what was kept vs excluded and why.")


_SAME_TYPE_RULES = (
    "You are an Amazon competitor analyst. You are given a TARGET product profile "
    "and a list of candidate products (ASIN + title) returned by keyword searches.\n\n"
    "Decide which candidates are the SAME CORE PRODUCT TYPE as the target — a real "
    "competitor a shopper would cross-shop against this exact product. Use the "
    "target profile's type, key attributes and 'NOT this' list.\n\n"
    "EXCLUDE: accessories, spare parts, look-alikes of a different type, bundles of "
    "a different product, and items from the profile's 'NOT this' list, even if they "
    "appear in the same search results. Keep both FBA and FBM sellers — fulfilment "
    "is irrelevant to product type.\n\n"
    "Return only the ASINs that ARE the same product type. Copy ASINs verbatim; never "
    "invent one. If unsure about a candidate, exclude it and say so in the note."
)


def filter_same_product_type(
    client,
    profile: ProductProfile,
    candidates: pd.DataFrame,
    max_candidates: int = 250,
) -> SameTypeResult:
    """Return the subset of *candidates* that are the same product type as the
    target. *candidates* needs 'asin' and 'title' columns."""
    if candidates.empty:
        raise Phase3ApiError("No candidates to classify for same-product-type.")

    rows = candidates.head(max_candidates)
    listing = "\n".join(
        f"  - {r['asin']}: {str(r.get('title') or '')[:140]}" for _, r in rows.iterrows()
    )
    profile_block = (
        f"TARGET PRODUCT PROFILE:\n  Name: {profile.product_name}\n"
        f"  Type: {profile.product_type}\n"
        f"  Key attributes: {', '.join(profile.key_attributes) or '(none)'}\n"
        f"  NOT this: {', '.join(profile.not_this) or '(none)'}"
    )
    system = [
        {"type": "text", "text": _SAME_TYPE_RULES, "cache_control": {"type": "ephemeral"}},
        {"type": "text", "text": profile_block, "cache_control": {"type": "ephemeral"}},
    ]
    user = f"CANDIDATES (ASIN: title):\n{listing}\n\nReturn the ASINs that are the same product type."
    try:
        resp = client.messages.parse(
            model=MODEL, max_tokens=16000, thinking={"type": "adaptive"},
            output_config={"effort": "medium"}, system=system,
            messages=[{"role": "user", "content": user}], output_format=SameTypeResult,
        )
    except Exception as e:  # noqa: BLE001
        raise Phase3ApiError(f"Claude same-product-type filter failed: {e}") from e
    if resp.parsed_output is None:
        raise Phase3ApiError("Claude returned no parseable same-product-type result.")
    return resp.parsed_output


# ──────────────────────────────────────────────────────────────────────────────
# Call 2 — listing quality scoring (vision)
# ──────────────────────────────────────────────────────────────────────────────

class ListingScores(BaseModel):
    product_images: int = Field(description="Product-page image quality score 0-20: professional/lifestyle/infographics/count.")
    product_images_reason: str
    product_title: int = Field(description="Title quality score 0-5: relevant, clear, well-structured (NOT keyword-stuffed).")
    product_title_reason: str
    bullets_description: int = Field(description="Bullets/description quality score 0-5: benefits, keywords, addresses pain points.")
    bullets_description_reason: str
    unique_design: int = Field(description="Design uniqueness score 0-2: 2 = visibly unique/differentiated design, 0 = generic/saturated. Ignore patents.")
    unique_design_reason: str
    marketing_images: int = Field(description="Overall brand/marketing imagery score 0-15 from the listing gallery + A+ visible (storefront not available).")
    marketing_images_reason: str
    pricing_strategy: int = Field(description="Pricing-strategy strength 0-10 from price positioning vs the market; use -1 if it cannot be judged without the seller's full catalog.")
    pricing_strategy_reason: str


_LISTING_RULES = (
    "You are scoring ONE Amazon competitor's listing for a competitor-analysis "
    "rubric. Higher score = STRONGER competitor (harder for a new seller to beat).\n\n"
    "Score each element within its max. Judge images from the provided listing "
    "images (form, professionalism, lifestyle shots, infographics, count). Judge "
    "title and bullets from the text. For 'unique_design', 2 = a visibly "
    "differentiated/distinctive design, 0 = a generic design sold by many; ignore "
    "patents (the user verifies those separately). For 'marketing_images', judge "
    "the overall brand imagery quality from the gallery and any A+ visible — note "
    "you do NOT have the brand storefront. For 'pricing_strategy', judge from the "
    "price positioning given; if you genuinely cannot assess it without the "
    "seller's full catalogue, return -1 (do not guess).\n\n"
    "Give a one-sentence reason per element. Base every judgement only on the "
    "evidence provided — never invent details you cannot see."
)


def score_listing(
    client,
    competitor: Dict,
    image_urls: Optional[List[str]] = None,
    max_images: int = 5,
) -> ListingScores:
    """Score the judgement-based rubric elements for one competitor.

    *competitor* should carry: title, bullets (list), price, market_price_context
    (a short string e.g. 'cheapest of 3' or price list), a_plus, has_video,
    images_count. *image_urls* are the listing image URLs (sent to vision).
    """
    bullets = competitor.get("bullets") or []
    text = (
        f"COMPETITOR LISTING\n"
        f"Title: {competitor.get('title') or '(none)'}\n"
        f"Price: {competitor.get('price')}\n"
        f"Price context vs market: {competitor.get('market_price_context') or '(n/a)'}\n"
        f"A+ content present: {competitor.get('a_plus')}\n"
        f"Video present: {competitor.get('has_video')}\n"
        f"Image count: {competitor.get('images_count')}\n"
        f"Bullets:\n" + "\n".join(f"  - {b}" for b in bullets[:8])
    )
    content: List[Dict] = [{"type": "text", "text": text}]
    for u in (image_urls or [])[:max_images]:
        content.append({"type": "image", "source": {"type": "url", "url": u}})

    try:
        resp = client.messages.parse(
            model=MODEL, max_tokens=3000, thinking={"type": "adaptive"},
            output_config={"effort": "medium"},
            system=[{"type": "text", "text": _LISTING_RULES, "cache_control": {"type": "ephemeral"}}],
            messages=[{"role": "user", "content": content}],
            output_format=ListingScores,
        )
    except Exception as e:  # noqa: BLE001
        raise Phase3ApiError(f"Claude listing-quality scoring failed: {e}") from e
    if resp.parsed_output is None:
        raise Phase3ApiError("Claude returned no parseable listing scores.")

    # Clamp each score to its element max (defensive; -1 sentinel kept as Unknown).
    s = resp.parsed_output
    for key, mx in LISTING_ELEMENT_MAX.items():
        val = getattr(s, key)
        if val is not None and val >= 0:
            setattr(s, key, min(val, mx))
    return s


# ──────────────────────────────────────────────────────────────────────────────
# Call 3 — review analysis (weaknesses / solutions / strengths)
# ──────────────────────────────────────────────────────────────────────────────

class ReviewAnalysis(BaseModel):
    weaknesses: List[str] = Field(description="Top 3 recurring complaints/pain points from 1-3 star reviews. Empty if no review data.")
    solutions: List[str] = Field(description="A concrete product solution for each weakness (materials/specs/design), aligned by index.")
    strengths: List[str] = Field(description="Top 3 features/materials/design aspects customers praise in 4-5 star reviews. Empty if no review data.")
    note: str = Field(description="Caveats; if no review data was provided, say 'Unknown — need user confirmation'.")


_REVIEW_RULES = (
    "You analyse Amazon customer feedback for a competitor. The input has two "
    "labelled sections:\n"
    "  - 'STRENGTHS (Amazon AI insight)': Amazon's 'Customers say' AI summary of "
    "what customers like. Derive the top 3 praised strengths from THIS section.\n"
    "  - 'WEAKNESSES (1-2 star reviews)': the actual 1-2 star reviews. Extract the "
    "top 3 recurring complaints from THESE; for each, propose a concrete product "
    "solution (material, specification, or design change).\n\n"
    "Base everything strictly on the supplied text. If a section is missing, leave "
    "its list empty and note it as 'Unknown — need user confirmation' — never "
    "invent complaints or praise (Rule 0)."
)


def analyze_reviews(client, competitor_label: str, reviews_text: str = "") -> ReviewAnalysis:
    """Analyse a competitor's reviews. With no review text, returns an explicit
    Unknown result (no fabrication)."""
    if not (reviews_text or "").strip():
        return ReviewAnalysis(
            weaknesses=[], solutions=[], strengths=[],
            note="Unknown — need user confirmation (no review data supplied).",
        )
    user = f"Competitor: {competitor_label}\n\nREVIEWS:\n{reviews_text[:15000]}"
    try:
        resp = client.messages.parse(
            model=MODEL, max_tokens=2500, thinking={"type": "adaptive"},
            output_config={"effort": "medium"},
            system=[{"type": "text", "text": _REVIEW_RULES, "cache_control": {"type": "ephemeral"}}],
            messages=[{"role": "user", "content": user}], output_format=ReviewAnalysis,
        )
    except Exception as e:  # noqa: BLE001
        raise Phase3ApiError(f"Claude review analysis failed: {e}") from e
    if resp.parsed_output is None:
        raise Phase3ApiError("Claude returned no parseable review analysis.")
    return resp.parsed_output
