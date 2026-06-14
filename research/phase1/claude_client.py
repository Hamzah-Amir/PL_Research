"""
Claude (Anthropic) client for Phase 1 scoring adjustments and explanations.

This module provides a small, safe wrapper to request judgement-based
adjustments from Claude for a batch of products. It is designed to be
opt-in: callers should check an environment toggle before invoking it.

Main entrypoints:
 - `load_client()` -> returns an `anthropic.Anthropic()` client or raises
 - `score_products_with_claude(client, products, budget)` -> list of
    `ScoreResult` entries (one per product). Uses local cache in
    `output/llm_cache/` to avoid repeated token costs.

The prompt asks Claude to produce small numeric adjustments (`llm_adj`) in
the range [-10, +10] together with a short reason and tags. The deterministic
scoring rules remain authoritative; Claude only suggests adjustments.
"""

from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha1
from pathlib import Path
from typing import Dict, List, Optional
import json
import os
import time

_ROOT = Path(__file__).resolve().parent.parent
_ENV_PATH = _ROOT / ".env"
_CACHE_DIR = _ROOT / "output" / "llm_cache"
_CACHE_DIR.mkdir(parents=True, exist_ok=True)

MODEL = "claude-opus-4-8"


class Phase1ApiError(RuntimeError):
    """Raised when the Claude client cannot be constructed or call fails."""


@dataclass
class ScoreResult:
    asin: str
    llm_adj: float
    llm_reason: str
    llm_tags: List[str]


def load_client():
    """Load an Anthropic client using the project `.env` file.

    Raises `Phase1ApiError` when the key/package is missing.
    """
    try:
        from dotenv import load_dotenv
    except Exception as e:  # pragma: no cover - runtime dependency
        raise Phase1ApiError("python-dotenv is required to load .env") from e

    try:
        import anthropic
    except Exception as e:  # pragma: no cover - runtime dependency
        raise Phase1ApiError("The 'anthropic' package is required but not installed.") from e

    load_dotenv(dotenv_path=_ENV_PATH)
    if not os.environ.get("ANTHROPIC_API_KEY"):
        raise Phase1ApiError(f"ANTHROPIC_API_KEY not found in {_ENV_PATH}")
    try:
        return anthropic.Anthropic()
    except Exception as e:
        raise Phase1ApiError(f"Could not initialise the Anthropic client: {e}") from e


def _cache_path_for(asins: List[str], budget: Optional[float]) -> Path:
    key = "|".join(sorted([str(a).upper() for a in asins])) + f"|{budget}"
    h = sha1(key.encode("utf-8")).hexdigest()
    return _CACHE_DIR / f"phase1_claude_batch_{h}.json"


def _build_system_and_user(products: List[Dict], budget: Optional[float]) -> (List[Dict], str):
    """Return system block (list) and user string content for the batch prompt."""
    system_text = (
        "You are an expert Amazon UK private-label analyst.\n"
        "Given a small set of product metrics and the user's scoring rules, produce a concise\n"
        "numeric adjustment (`llm_adj`) between -10 and +10 to add to the product's numeric\n"
        "opportunity score. Also return a brief, specific reason explaining your adjustment.\n\n"
        "Rules (apply deterministically in Python; you are NOT replacing them):\n"
        " - The final score is made of two pillars: Market Entry (50) and Growth (50).\n"
        " - Review counts, BSR stability, fulfillment (FBA/FBM), and price band are primary factors.\n\n"
        "For each product, analyze:\n"
        " 1. Market saturation: Is this category competitive (high reviews) or underexploited?\n"
        " 2. Pricing opportunity: Is price positioned well vs. competitors in this niche?\n"
        " 3. Demand signals: Does the sales trend, velocity, and revenue suggest growth potential?\n"
        " 4. Seasonality: Does seasonal volatility pose execution risk?\n"
        " 5. Brand risk: Are there red flags (e.g., niche-specific challenges)?\n\n"
        "Provide reasoning that explains what drove your adjustment, e.g.:\n"
        " - Positive adjustments: 'Market underexploited (low reviews), strong margin, growing trend'\n"
        " - Negative adjustments: 'High competition (saturated reviews), thin margin, erratic trend'\n"
        " - Zero adjustment: 'Deterministic score already reflects all factors; no additional insight'\n\n"
        "Your output must be a JSON array of objects exactly matching the schema: `asin` (string), `llm_adj` (float, -10 to +10), `llm_reason` (string, brief and specific), `llm_tags` (array of short tags).\n"
    )

    listing_lines = []
    for p in products:
        asin = str(p.get("asin") or p.get("ASIN") or "").upper()
        title = p.get("title") or p.get("name") or ""
        price = p.get("price") or p.get("price_£") or p.get("price_gbp") or ""
        reviews = p.get("review_count") or p.get("reviews") or ""
        rating = p.get("rating") or ""
        bsr = p.get("bsr") or p.get("BSR") or ""
        sales = p.get("asin_sales") or p.get("sales") or p.get("estimated_sales") or ""
        revenue = p.get("asin_revenue") or p.get("revenue") or p.get("estimated_revenue") or ""
        trend = p.get("sales_trend_90d") or p.get("trend90") or ""
        fulfillment = p.get("fulfillment") or p.get("fulfilment") or ""
        weight = p.get("weight_kg") or p.get("weight") or ""
        seasonal = p.get("is_seasonal") or p.get("seasonal") or False

        listing_lines.append(
            f"ASIN: {asin}\n  Title: {title[:120]}\n  Price: {price}\n  Reviews: {reviews}\n  Rating: {rating}\n  BSR: {bsr}\n  Sales/mo: {sales}\n  Revenue/mo: {revenue}\n  Trend90d: {trend}\n  Fulfillment: {fulfillment}\n  Weight_kg: {weight}\n  Seasonal: {seasonal}\n"
        )

    user_block = (
        f"BATCH: {len(products)} products. Budget: £{budget}\n\n"
        "For each product block below, output one JSON object with:\n"
        "  - `asin`: Product ASIN\n"
        "  - `llm_adj`: Your adjustment (-10 to +10 points)\n"
        "  - `llm_reason`: Why you made this adjustment. Be specific about which factors drove it (e.g., 'Market underexploited', 'Strong velocity but thin margin', 'Seasonal risk')\n"
        "  - `llm_tags`: Quick labels (e.g., ['underexploited', 'good_velocity'], ['oversaturated'], ['risk_seasonal'])\n\n"
        "PRODUCTS:\n\n"
        + "\n---\n".join(listing_lines)
    )
    system = [{"type": "text", "text": system_text, "cache_control": {"type": "ephemeral"}}]
    return system, user_block


def score_products_with_claude(client, products: List[Dict], budget: Optional[float], *, max_batch: int = 20) -> List[ScoreResult]:
    """Score up to `max_batch` products using a single batched Claude call.

    Returns a list of `ScoreResult` in the same order as the supplied products.
    Uses local cache to avoid repeated token costs. If Claude is unavailable,
    raises `Phase1ApiError`.
    """
    if not products:
        return []
    batch = products[:max_batch]
    asins = [str(p.get("asin") or p.get("ASIN") or "").upper() for p in batch]
    cache_path = _cache_path_for(asins, budget)
    if cache_path.exists():
        try:
            data = json.loads(cache_path.read_text(encoding="utf-8"))
            return [ScoreResult(**d) for d in data]
        except Exception:
            # If cache is corrupt, ignore and continue
            pass

    system, user = _build_system_and_user(batch, budget)

    try:
        resp = client.messages.parse(
            model=MODEL,
            max_tokens=16000,
            thinking={"type": "adaptive"},
            output_config={"effort": "medium"},
            system=system,
            messages=[{"role": "user", "content": user}],
        )
    except Exception as e:
        raise Phase1ApiError(f"Claude Phase1 batch scoring failed: {e}") from e

    # `resp` may carry `response` textual content — try to parse JSON array out of it
    text = getattr(resp, "response", None) or getattr(resp, "text", None) or str(resp)
    # attempt to extract the first JSON array substring
    try:
        start = text.index("[")
        end = text.rindex("]") + 1
        j = json.loads(text[start:end])
    except Exception:
        # Last resort: try to parse entire text
        try:
            j = json.loads(text)
        except Exception as e:
            raise Phase1ApiError(f"Could not parse Claude response as JSON: {e}\nResponse: {text[:2000]}") from e

    results: List[ScoreResult] = []
    out_data: List[Dict] = []
    for o in j:
        asin = str(o.get("asin") or "").upper()
        try:
            adj = float(o.get("llm_adj") or 0.0)
        except Exception:
            adj = 0.0
        reason = str(o.get("llm_reason") or "").strip()
        tags = o.get("llm_tags") or []
        if not isinstance(tags, list):
            tags = [str(tags)]
        sr = ScoreResult(asin=asin, llm_adj=adj, llm_reason=reason, llm_tags=[str(t) for t in tags])
        results.append(sr)
        out_data.append({"asin": sr.asin, "llm_adj": sr.llm_adj, "llm_reason": sr.llm_reason, "llm_tags": sr.llm_tags})
        # Debug: print reasoning for verification
        print(f"  [{asin}] adj={adj:+.1f} reason='{reason}'")

    # write cache (best-effort)
    try:
        cache_path.write_text(json.dumps(out_data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass

    return results
