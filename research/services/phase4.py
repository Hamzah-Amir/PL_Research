"""
PL Research Tool — Phase 4 (Critical Sheet, web / library form).

Builds the Critical Sheet (v9.6 workflow Phase 4) adapted to this Django tool:
Keepa is a live API here (not a manual bulk-export upload), so the PDF's manual
"variation ASIN extraction → run a Keepa bulk lookup → re-upload" pause is done
automatically. What remains is a single approval pause — the controlled
vocabulary + category-specific Design attributes (Step 12/13) — surfaced in the
web UI. After approval, Claude classifies every parent (Design + Material / Size
/ Color / Special Features / Packaging using image + text, Rule 8), variation
children are pulled from Keepa and added as their own rows (Rule 7), and a
fresh Critical Sheet workbook is written.

Rule 0 (Zero Assumption): any value with no verified source is written as
"Unknown — need user confirmation", never guessed.

Output is a COPY of the real PES workbook (`test_file/PES UK.xlsx`) with the
'Critical sheet' tab filled — the original is never modified, and all its other
tabs, formulas, charts and pivot tables are preserved byte-for-byte (we edit
only the Critical sheet's worksheet XML, never round-tripping through openpyxl).
All Keepa/Anthropic plumbing is reused from phase3 (cached, token-aware).
"""
from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Dict, List, Optional

import re

from pydantic import BaseModel, Field

from research.services.phase3 import (
    DOMAIN,
    MODEL,
    Phase3ApiError,
    Phase3KeepaError,
    _OUT_DIR,
    _ROOT,
    _ASIN_RE,
    _build_per_kw,
    _load_cache,
    _save_cache,
    extract_keepa_fields,
    filter_same_product_type,
    load_client,
    load_keepa,
    read_js_products,
)
from research.services.phase2 import derive_product_profile

UNKNOWN = "Unknown — need user confirmation"

# The Critical Sheet is filled into a COPY of the full PES workbook so the
# shipped file keeps all its tabs, formulas, charts and pivot tables. Layout
# below reflects the template's 'Critical sheet' tab: headers on row 8, data
# from row 9, date B6, keyword B7, Design legend in the AB(code)/AC(description)
# block above the header. We never load the workbook through openpyxl (which
# strips charts/pivots) — we copy the .xlsx and edit only the Critical sheet's
# worksheet XML in place (see write_critical_sheet).
PES_SOURCE = _ROOT / "test_file" / "PES UK.xlsx"
CRITICAL_SHEET = "Critical sheet"

HEADER_ROW = 8
DATA_START = 9
LEGEND_START = 2  # first legend row (above the header)
LEGEND_END = HEADER_ROW - 1  # last legend row (7)
LEGEND_CODE_COL = 28  # AB — Design legend code
LEGEND_DESC_COL = 29  # AC — Design legend description
ATTR_ROW = 7          # design-attribute names row (AD7…AG7)
ATTR_COL_START = 30   # AD
DATE_CELL = "B6"
KW_CELL = "B7"

# row-dict key -> 1-based column index in the template's 'Critical sheet'.
# Columns left unmapped (Min/Consistent/Max Price, SnL/Prime, Pack size,
# Comments, Screenshot) are user/derived fields — left as the template has them.
TEMPLATE_COL_MAP: Dict[str, int] = {
    "asin": 1,                  # A  ASIN
    "title": 2,                 # B  Product Name
    "brand": 3,                 # C  Brand
    "price": 4,                 # D  Price
    "min_price": 5,             # E  Min. Price
    "consistent_price": 6,      # F  Consistent Price (90-day avg)
    "max_price": 7,             # G  Max. Price
    "monthly_sales": 8,         # H  Mo. Sales
    "daily_sales": 9,           # I  D. Sales
    "monthly_revenue": 10,      # J  Mo. Revenue
    "date_first_available": 11, # K  Date First Available
    "reviews": 12,              # L  Rating Number
    "rating": 13,               # M  Rating
    "fees": 14,                 # N  Fees
    "bsr": 15,                  # O  Rank
    "fulfillment": 16,          # P  Seller type
    "sellers": 17,              # Q  Sellers
    "subcategory": 18,          # R  Sub-Category
    "category": 19,             # S  Category
    "tier": 20,                 # T  Tier
    "dimensions": 21,           # U  Dimensions
    "weight": 22,               # V  Weight
    "link": 23,                 # W  Link
    "snl_prime": 24,            # X  SnL/Prime
    "material": 25,             # Y  Material
    "pack_size": 26,            # Z  Pack size
    "color": 27,                # AA Color
    "size": 28,                 # AB Size
    "design": 29,               # AC Design
    "sf1": 30,                  # AD Special F1
    "sf2": 31,                  # AE Special F2
    "sf3": 32,                  # AF Special F3
    "packaging": 33,            # AG Packaging
}

# (key, header). The Design column carries the legend block above the header.
COLUMNS: List = [
    ("asin", "ASIN"),
    ("parent_asin", "Parent ASIN"),
    ("row_type", "Row Type"),
    ("brand", "Brand"),
    ("title", "Product Name"),
    ("design", "Design"),
    ("variation", "Variation"),
    ("subcategory", "Sub-Category"),
    ("category", "Category"),
    ("tier", "Tier"),
    ("monthly_sales", "Monthly Sales"),
    ("daily_sales", "Daily Sales"),
    ("monthly_revenue", "Monthly Revenue"),
    ("price", "Price"),
    ("rating", "Rating"),
    ("reviews", "Rating Count"),
    ("bsr", "BSR"),
    ("date_first_available", "Date First Available"),
    ("material", "Material"),
    ("size", "Size"),
    ("color", "Color"),
    ("sf1", "Special Feature 1"),
    ("sf2", "Special Feature 2"),
    ("sf3", "Special Feature 3"),
    ("packaging", "Packaging"),
    ("dimensions", "Dimensions"),
    ("weight", "Weight"),
    ("fees", "FBA Fees"),
    ("fulfillment", "FBA/FBM"),
    ("link", "Link"),
]

# Rule 7 — for a variation child these MUST come from variation-level Keepa data
# (never copied from the parent); Unknown if the child's own data is absent.
VARIATION_LEVEL = {
    "monthly_sales", "daily_sales", "monthly_revenue", "price", "rating",
    "reviews", "bsr", "date_first_available", "variation", "link",
}
# Rule 7 — static attributes a child may inherit from its parent.
PARENT_COPYABLE = {
    "brand", "title", "subcategory", "category", "tier", "dimensions",
    "weight", "fees", "material", "design", "sf1", "sf2", "sf3", "packaging",
    "fulfillment",
}

DEFAULT_PARENT_LIMIT = 20  # token guard — cap classified parents per run


# ──────────────────────────────────────────────────────────────────────────────
# Keepa enrichment — SELF-CONTAINED to Phase 4 (no changes to phase3.py). We
# call the Keepa client directly with the richer parameter set Phase 4 needs
# (stats + history + rating + buy-box) and cache it in a Phase-4-owned namespace.
# ──────────────────────────────────────────────────────────────────────────────

# Amazon "seller id" that means Amazon itself is the seller (all GB marketplaces).
_AMAZON_SELLER_IDS = {"A3P5ROKL5A1OLE", "ATVPDKIKX0DER", "A1F83G8C2ARO7P"}


def keepa_full_query(api, asins: List[str], *, use_cache: bool = True,
                     limit: Optional[int] = None) -> Dict[str, dict]:
    """Phase-4 Keepa pull: ``stats=180 & history & rating & buybox`` so every
    ASIN yields min/max/90-day prices, rating + review count, buy-box seller
    type and offer counts. Cached under the ``phase4`` namespace (separate from
    Phase 3's shallow/deep caches). ~4 tokens per uncached ASIN. Returns
    ``{asin: raw_product}``. Raises Phase3KeepaError on API failure."""
    cache = _load_cache("phase4") if use_cache else {}
    wanted = [a for a in asins if _ASIN_RE.match(a or "")]
    have = {a: cache[a] for a in wanted if a in cache}
    missing = [a for a in wanted if a not in cache]
    if limit is not None and len(missing) > limit:
        missing = missing[:limit]
    if missing:
        try:
            products = api.query(missing, domain=DOMAIN, history=True, stats=180,
                                 rating=True, buybox=True, progress_bar=False)
        except Exception as e:  # noqa: BLE001
            raise Phase3KeepaError(f"Keepa Phase-4 query failed: {e}") from e
        for prod in products:
            a = prod.get("asin")
            if a:
                cache[a] = prod
                have[a] = prod
        if use_cache:
            _save_cache("phase4", cache)
    return {a: have[a] for a in wanted if a in have}


def _stat_value(sp: dict, band: str, key: str):
    """Read a value from a Keepa stats_parsed band. min/max are (date, value)
    tuples; current/avg* are scalars. Returns a positive float or None."""
    v = (sp.get(band) or {}).get(key)
    if isinstance(v, (tuple, list)):
        v = v[1] if len(v) > 1 else None
    return round(float(v), 2) if isinstance(v, (int, float)) and v > 0 else None


def _stat_price(sp: dict, band: str):
    """Best available price for a stats band, trying NEW → AMAZON → buy box."""
    for key in ("NEW", "AMAZON", "BUY_BOX_SHIPPING"):
        v = _stat_value(sp, band, key)
        if v is not None:
            return v
    return None


def _seller_type(product: dict) -> Optional[str]:
    """Buy-box seller type: 'FBA' / 'FBM'. An Amazon-sold product counts as FBA
    (fulfilled by Amazon), per the user's rule."""
    st = product.get("stats") or {}
    sp = product.get("stats_parsed") or {}
    is_amazon = st.get("buyBoxIsAmazon")
    if is_amazon is None:
        is_amazon = sp.get("buyBoxIsAmazon")
    sid = st.get("buyBoxSellerId") or sp.get("buyBoxSellerId")
    fba = st.get("buyBoxIsFBA")
    if fba is None:
        fba = sp.get("buyBoxIsFBA")
    if is_amazon or (sid and sid in _AMAZON_SELLER_IDS) or fba is True:
        return "FBA"
    if fba is False:
        return "FBM"
    return None


def _snl_prime(price) -> Optional[str]:
    """SnL/Prime rule: product under £10 → 'SNL' (Small & Light), else 'Prime'."""
    if isinstance(price, (int, float)):
        return "SNL" if price < 10 else "Prime"
    return None


# Actual colour words, used to pull a real colour out of a variation attribute
# string (Amazon often crams size/pack into the "colour" theme).
_COLOR_WORDS = [
    "black", "silver", "grey", "gray", "white", "orange", "green", "blue",
    "red", "yellow", "pink", "purple", "gold", "brown", "chrome", "beige",
    "navy", "teal", "clear", "transparent", "multicolour", "multicolor",
]


def _parse_variation(s: Optional[str]):
    """Split a variation attribute string (e.g. 'Colour: 20cm - Black; Size: 20
    Pack') into (colour, product_size, pack_size) using heuristics, since Amazon
    labels these inconsistently."""
    if not s:
        return None, None, None
    text = re.sub(r"\b(colou?r|size|style|pattern)\s*:", " ", s, flags=re.I)
    low = text.lower()
    pm = re.search(r"(\d+\s*[-\s]?pack|pack\s*of\s*\d+|single)", low)
    pack = pm.group(0).strip() if pm else None
    zm = re.search(
        r"(\d+(?:\.\d+)?\s*(?:mm|cm|inch|in|ml|l|\")\b(?:\s*[x×]\s*\d+(?:\.\d+)?\s*(?:mm|cm|inch|in)?\b)*)",
        low)
    size = zm.group(0).strip() if zm else None
    color = None
    for w in _COLOR_WORDS:
        if re.search(r"\b" + re.escape(w) + r"\b", low):
            color = w.title()
            break
    return color, size, pack


def _keepa_full(product: dict) -> dict:
    """All the per-ASIN fields Phase 4 sources from Keepa (Rule-0 safe: None when
    the value is genuinely absent). Works for a parent OR a variation ASIN."""
    sp = product.get("stats_parsed") or {}
    cur = sp.get("current") or {}

    rating = cur.get("RATING")
    if isinstance(rating, (int, float)) and rating not in (-1, None):
        rating = round(rating / 10.0, 1) if rating > 5 else round(rating, 1)  # 0-50 → 0-5
    else:
        rating = None
    reviews = cur.get("COUNT_REVIEWS")
    reviews = int(reviews) if isinstance(reviews, (int, float)) and reviews >= 0 else None
    count_new = cur.get("COUNT_NEW")
    num_sellers = int(count_new) if isinstance(count_new, (int, float)) and count_new > 0 else None
    monthly = product.get("monthlySold")
    monthly = int(monthly) if isinstance(monthly, (int, float)) and monthly > 0 else None
    bsr = cur.get("SALES") if cur.get("SALES", -1) not in (-1, None) else None

    # Date First Available: no exact Keepa field — use trackingSince (first seen).
    dfa = None
    ts = product.get("trackingSince")
    if isinstance(ts, (int, float)) and ts > 0:
        try:
            import keepa
            dfa = keepa.keepa_minutes_to_time(ts).strftime("%Y-%m-%d")
        except Exception:  # noqa: BLE001
            dfa = None

    price = _stat_price(sp, "current")
    return {
        "price": price,
        "min_price": _stat_price(sp, "min"),
        "max_price": _stat_price(sp, "max"),
        "consistent_price": _stat_price(sp, "avg90"),  # 90-day weighted average
        "rating": rating,
        "reviews": reviews,
        "num_sellers": num_sellers,
        "monthly_sold": monthly,
        "seller_type": _seller_type(product),
        "date_first_available": dfa,
        "color": (product.get("color") or "").strip() or None,
        "bsr": bsr,
        "dimensions": _fmt_keepa_dims(product),
        "weight": _fmt_keepa_weight(product),
    }


def _fmt_keepa_dims(product: dict) -> Optional[str]:
    """Keepa package dimensions (mm) → 'L x W x H cm', or None if incomplete."""
    dims = [product.get("packageLength"), product.get("packageWidth"), product.get("packageHeight")]
    if all(isinstance(d, (int, float)) and d > 0 for d in dims):
        return " x ".join(f"{d / 10:.1f}" for d in dims) + " cm"
    return None


def _fmt_keepa_weight(product: dict) -> Optional[str]:
    """Keepa weight (g) → 'X.XX kg' (package weight preferred), or None."""
    g = product.get("packageWeight") or product.get("itemWeight")
    if isinstance(g, (int, float)) and g > 0:
        return f"{g / 1000:.2f} kg"
    return None


def _variation_children(fields: dict, exclude: set) -> List[dict]:
    """Parse a parent's Keepa `variation_attrs` into child records
    ``[{"asin", "attributes_str"}]``, skipping the parent itself and any ASIN
    in *exclude* (e.g. other JS parents already in the sheet)."""
    out = []
    seen = set()
    for v in (fields.get("variation_attrs") or []):
        if not isinstance(v, dict):
            continue
        casin = str(v.get("asin") or "").strip().upper()
        if not _ASIN_RE.match(casin) or casin in exclude or casin in seen:
            continue
        seen.add(casin)
        attrs = v.get("attributes") or []
        parts = []
        for a in attrs:
            if isinstance(a, dict):
                dim = a.get("dimension") or a.get("name")
                val = a.get("value")
                if val:
                    parts.append(f"{dim}: {val}" if dim else str(val))
        out.append({"asin": casin, "attributes_str": "; ".join(parts) or None})
    return out


# ──────────────────────────────────────────────────────────────────────────────
# Claude — Step 12: controlled vocabulary + category-specific Design attributes
# ──────────────────────────────────────────────────────────────────────────────

class VocabularyProposal(BaseModel):
    material_values: List[str] = Field(description="6-8 standard Material values for this category (main product body material). Synonyms collapsed.")
    size_guidance: str = Field(description="How to express Size for this category (the meaningful measurement, e.g. capacity/length/dimensions). One sentence.")
    color_values: List[str] = Field(description="6-8 standard Color values seen across these products.")
    packaging_values: List[str] = Field(description="Packaging types seen (e.g. retail box, polybag, pouch, tube).")
    special_feature_guidance: str = Field(description="What SF1 (primary differentiator), SF2 and SF3 mean for THIS category. One or two sentences.")
    design_attributes: List[str] = Field(description="The category-specific distinguishing attributes that actually vary across these products and should decide Design codes (per Rule 8). 3-6 attributes.")
    note: str = Field(description="One or two sentences on the category and any ambiguity.")


_VOCAB_RULES = (
    "You are preparing a controlled vocabulary for an Amazon private-label "
    "Critical Sheet (product-comparison spreadsheet). You are given a sample of "
    "same-category products (title + bullets + description + main image).\n\n"
    "Propose, per Rule 8 and product-agnostic rules:\n"
    " - Material: 6-8 standard values describing the MAIN product body material "
    "(not a minor component). Collapse synonyms.\n"
    " - Size: how to express the meaningful measurement for THIS category "
    "(capacity, length, dimensions…), avoiding generic Large/Mini when a "
    "measurable spec exists.\n"
    " - Color: 6-8 standard colour values seen.\n"
    " - Packaging: the packaging types actually seen.\n"
    " - Special Features: what the primary (SF1), secondary (SF2) and tertiary "
    "(SF3) differentiators mean for this category.\n"
    " - Design attributes: the category-specific attributes that actually vary "
    "across these products and must decide the Design grouping (form/shape, "
    "mechanism, head type, etc.). These will be approved by the user before any "
    "Design codes are assigned.\n\n"
    "Base everything strictly on the supplied products. Do not invent values for "
    "attributes you cannot see (Rule 0)."
)


def _sample_content(samples: List[dict], max_images: int = 8) -> List[dict]:
    """Build a Claude content list (text + a few main images) from product
    samples ``[{asin,title,bullets,description,category,image_url}]``."""
    lines = []
    for i, s in enumerate(samples, 1):
        bullets = "; ".join((s.get("bullets") or [])[:4])
        desc = (s.get("description") or "")[:300]
        lines.append(
            f"[{i}] ASIN {s.get('asin')} | Category: {s.get('category') or '(unknown)'}\n"
            f"    Title: {s.get('title') or '(none)'}\n"
            f"    Bullets: {bullets or '(none)'}\n"
            f"    Description: {desc or '(none)'}"
        )
    content: List[dict] = [{"type": "text", "text": "PRODUCT SAMPLE\n\n" + "\n\n".join(lines)}]
    imgs = 0
    for s in samples:
        url = s.get("image_url")
        if url and imgs < max_images:
            content.append({"type": "image", "source": {"type": "url", "url": url}})
            imgs += 1
    return content


def propose_vocabulary(client, samples: List[dict]) -> VocabularyProposal:
    """Step 12 — Claude proposes the controlled vocabulary + Design attributes."""
    if not samples:
        raise Phase3ApiError("No product samples to propose vocabulary from.")
    try:
        resp = client.messages.parse(
            model=MODEL, max_tokens=2500, thinking={"type": "adaptive"},
            output_config={"effort": "medium"},
            system=[{"type": "text", "text": _VOCAB_RULES, "cache_control": {"type": "ephemeral"}}],
            messages=[{"role": "user", "content": _sample_content(samples)}],
            output_format=VocabularyProposal,
        )
    except Exception as e:  # noqa: BLE001
        raise Phase3ApiError(f"Claude vocabulary proposal failed: {e}") from e
    if resp.parsed_output is None:
        raise Phase3ApiError("Claude returned no parseable vocabulary proposal.")
    return resp.parsed_output


# ──────────────────────────────────────────────────────────────────────────────
# Claude — Step 14B: per-product Design + attribute classification (Rule 8)
# ──────────────────────────────────────────────────────────────────────────────

class ProductClassification(BaseModel):
    asin: str = Field(description="The product ASIN, copied verbatim.")
    design_code: int = Field(description="Design group number (1,2,3…). Use -1 if it cannot be determined from image + text.")
    design_reason: str = Field(description="Which distinguishing attributes put it in this group.")
    material: str = Field(description="Main body material from the approved list, or 'Unknown — need user confirmation'.")
    size: str = Field(description="The product's physical SIZE measurement only (e.g. length/dimensions/capacity like '9 inch' or '4mm x 23cm') — NOT the pack count. 'Unknown — need user confirmation' if not stated.")
    pack_size: str = Field(description="The multipack/quantity only (e.g. 'Pack of 20', '10 Pack', 'Single'). NOT the product dimensions. 'Unknown — need user confirmation' if not stated.")
    color: str = Field(description="An actual COLOUR only (e.g. Black, Silver, Orange). Never a size, pack count or material. 'Unknown — need user confirmation' if no colour is stated.")
    sf1: str = Field(description="Primary differentiator, or 'Unknown — need user confirmation'.")
    sf2: str = Field(description="Secondary feature, or 'Unknown — need user confirmation'.")
    sf3: str = Field(description="Tertiary feature/bonus, or 'Unknown — need user confirmation'.")
    packaging: str = Field(description="Packaging type, or 'Unknown — need user confirmation'.")


class DesignLegendItem(BaseModel):
    code: int = Field(description="Design code number.")
    description: str = Field(description="Short description of what defines this Design group.")


class ClassificationBatch(BaseModel):
    items: List[ProductClassification] = Field(description="One entry per product, ASINs copied verbatim.")
    legend: List[DesignLegendItem] = Field(description="Legend: one entry per distinct Design code used.")


def _classify_rules(vocab: dict) -> str:
    return (
        "You classify Amazon products for a Critical Sheet using the combined "
        "reading of main image + title + bullets + description (Rule 8 — image "
        "shows form, text reveals mechanism/function; neither alone is enough).\n\n"
        "APPROVED VOCABULARY (use these; mark 'Unknown — need user confirmation' "
        "if a value cannot be verified from image + text — never guess, Rule 0):\n"
        f" - Material values: {', '.join(vocab.get('material_values') or []) or '(none given)'}\n"
        f" - Size guidance: {vocab.get('size_guidance') or '(none)'}\n"
        f" - Color values: {', '.join(vocab.get('color_values') or []) or '(none given)'}\n"
        f" - Packaging values: {', '.join(vocab.get('packaging_values') or []) or '(none given)'}\n"
        f" - Special-feature guidance: {vocab.get('special_feature_guidance') or '(none)'}\n"
        f" - Design distinguishing attributes: {', '.join(vocab.get('design_attributes') or []) or '(none given)'}\n\n"
        "COLOUR / SIZE / PACK — keep these three STRICTLY separate: 'color' is an "
        "actual colour only (never a size, pack count or material); 'size' is the "
        "product's physical measurement only (length/dimensions/capacity); "
        "'pack_size' is the multipack count only (e.g. 'Pack of 20'). Mark any of "
        "them 'Unknown — need user confirmation' if not stated.\n\n"
        "DESIGN RULE: two products share a Design code ONLY IF ALL the approved "
        "distinguishing attributes match. Assign codes 1,2,3… Then RE-EXAMINE "
        "every group and check whether a further attribute splits it into "
        "sub-designs; if so, use a new code rather than lumping them together "
        "(the sub-design corrective-action rule). Return a legend describing "
        "each distinct code you used. Copy every ASIN verbatim."
    )


def _classify_content(rows: List[dict], max_images: int) -> List[dict]:
    lines = []
    for r in rows:
        bullets = "; ".join((r.get("bullets") or [])[:5])
        desc = (r.get("description") or "")[:400]
        lines.append(
            f"ASIN {r.get('asin')}\n"
            f"  Title: {r.get('title') or '(none)'}\n"
            f"  Bullets: {bullets or '(none)'}\n"
            f"  Description: {desc or '(none)'}"
        )
    content: List[dict] = [{"type": "text", "text": "PRODUCTS TO CLASSIFY\n\n" + "\n\n".join(lines)}]
    imgs = 0
    for r in rows:
        url = (r.get("image_urls") or [None])[0]
        if url and imgs < max_images:
            content.append({"type": "image", "source": {"type": "url", "url": f"{url}"}})
            content.append({"type": "text", "text": f"(image above: ASIN {r.get('asin')})"})
            imgs += 1
    return content


def classify_products(client, rows: List[dict], vocab: dict, batch: int = 6) -> Dict:
    """Step 14B — classify parents' Design + structured attributes. Returns
    ``{"by_asin": {asin: {...}}, "legend": [{code,description}]}``. Batched to
    bound vision tokens; a batch whose response is truncated (too many products
    for the token budget → invalid JSON) is auto-split and retried, down to a
    single product, so a large parent set never fails the whole run."""
    by_asin: Dict[str, dict] = {}
    legend: Dict[int, str] = {}
    rules = _classify_rules(vocab)

    def _absorb(parsed) -> None:
        for it in parsed.items:
            code = it.design_code if it.design_code and it.design_code > 0 else None
            by_asin[it.asin.strip().upper()] = {
                "design": code,
                "material": it.material, "size": it.size, "pack_size": it.pack_size,
                "color": it.color, "sf1": it.sf1, "sf2": it.sf2, "sf3": it.sf3,
                "packaging": it.packaging,
            }
        for li in parsed.legend:
            if li.code and li.code > 0:
                legend.setdefault(li.code, li.description)

    def _process(chunk: List[dict]) -> None:
        try:
            resp = client.messages.parse(
                model=MODEL, max_tokens=8000, thinking={"type": "adaptive"},
                output_config={"effort": "medium"},
                system=[{"type": "text", "text": rules, "cache_control": {"type": "ephemeral"}}],
                messages=[{"role": "user", "content": _classify_content(chunk, max_images=len(chunk))}],
                output_format=ClassificationBatch,
            )
            parsed = resp.parsed_output
            if parsed is None:
                raise ValueError("no parseable classification")
        except Exception as e:  # noqa: BLE001 — likely a truncated/invalid JSON
            if len(chunk) > 1:  # split and retry the smaller halves
                mid = len(chunk) // 2
                _process(chunk[:mid])
                _process(chunk[mid:])
                return
            raise Phase3ApiError(f"Claude classification failed: {e}") from e
        _absorb(parsed)

    for i in range(0, len(rows), batch):
        _process(rows[i:i + batch])
    return {"by_asin": by_asin, "legend": [{"code": c, "description": legend[c]} for c in sorted(legend)]}


# ──────────────────────────────────────────────────────────────────────────────
# Row assembly (Rule 7 enforcement + Rule 0 Unknown)
# ──────────────────────────────────────────────────────────────────────────────

def _v(x):
    """Rule 0 display: a real value, else the Unknown marker."""
    if x is None:
        return UNKNOWN
    if isinstance(x, str) and not x.strip():
        return UNKNOWN
    return x


def _fulfillment(js_row: Optional[dict]) -> Optional[str]:
    if not js_row:
        return None
    ful = js_row.get("fulfillment")
    return str(ful) if ful is not None and str(ful).strip() else None


def _parent_row(asin: str, js_row: Optional[dict], kfields: dict, kfull: dict,
                cls: Optional[dict]) -> dict:
    """Parent row. Per the user's decision, the market data (prices, sales,
    revenue, rating/reviews, date, seller type, sellers, colour, dims/weight)
    is sourced from Keepa (`kfull`), with the JS sheet as a fallback; Design /
    Material / Size / SF / Packaging come from Claude (`cls`)."""
    js = js_row or {}
    cls = cls or {}
    kf = kfull or {}
    monthly = kf.get("monthly_sold")
    if monthly is None:
        monthly = js.get("monthly_sales")
    daily = round(monthly / 30.0, 1) if isinstance(monthly, (int, float)) else js.get("daily_sales")
    price = kf.get("price") if kf.get("price") is not None else js.get("price")
    revenue = (round(price * monthly, 2)
               if isinstance(price, (int, float)) and isinstance(monthly, (int, float))
               else js.get("monthly_revenue"))
    return {
        "asin": asin,
        "parent_asin": kfields.get("parent_asin") or "(parent)",
        "row_type": "Parent",
        "brand": _v(js.get("brand") or kfields.get("brand")),
        "title": _v(kfields.get("title") or js.get("title")),
        "design": _v(cls.get("design")),
        "variation": "(parent)",
        "subcategory": _v(kfields.get("subcategory") or js.get("category")),
        "category": _v(kfields.get("category") or js.get("category")),
        "tier": _v(js.get("product_tier")),
        "min_price": _v(kf.get("min_price")),
        "consistent_price": _v(kf.get("consistent_price")),
        "max_price": _v(kf.get("max_price")),
        "monthly_sales": _v(monthly),
        "daily_sales": _v(daily),
        "monthly_revenue": _v(revenue),
        "price": _v(price),
        "rating": _v(kf.get("rating") if kf.get("rating") is not None else js.get("rating")),
        "reviews": _v(kf.get("reviews") if kf.get("reviews") is not None else js.get("reviews")),
        "bsr": _v(kf.get("bsr") or js.get("bsr")),
        "sellers": _v(kf.get("num_sellers") if kf.get("num_sellers") is not None else js.get("num_sellers")),
        "date_first_available": _v(kf.get("date_first_available") or js.get("date_first_available")),
        "material": _v(cls.get("material")),
        "size": _v(cls.get("size")),          # product size only (Claude)
        "pack_size": _v(cls.get("pack_size")),  # multipack count only (Claude)
        "color": _v(cls.get("color")),        # a real colour only (Claude)
        "sf1": _v(cls.get("sf1")),
        "sf2": _v(cls.get("sf2")),
        "sf3": _v(cls.get("sf3")),
        "packaging": _v(cls.get("packaging")),
        "dimensions": _v(kf.get("dimensions") or js.get("dimensions")),
        "weight": _v(kf.get("weight") or js.get("weight")),
        "fees": _v(js.get("amazon_fees") or kfields.get("fba_pick_pack_fee")),
        "fulfillment": _v(kf.get("seller_type") or _fulfillment(js_row)),
        "snl_prime": _v(_snl_prime(price)),
        "link": _v(js.get("link") or (f"https://www.amazon.co.uk/dp/{asin}")),
    }


def _child_row(child: dict, kfields: dict, kfull: dict, parent_row: dict,
               js_row: Optional[dict] = None) -> dict:
    """Rule 7: variation-level fields from the child's OWN data — its own Keepa
    pull (`kfull`) first, then its JS-sheet row if it ranks for the keyword,
    else Unknown. Static attributes (Design/Material/SF/Packaging/brand/…) are
    inherited from the parent row."""
    js = js_row or {}
    kf = kfull or {}
    # Split the variation attribute into a real colour / product size / pack count.
    v_color, v_size, v_pack = _parse_variation(child.get("attributes_str"))
    monthly = kf.get("monthly_sold")
    if monthly is None:
        monthly = js.get("monthly_sales")
    daily = round(monthly / 30.0, 1) if isinstance(monthly, (int, float)) else js.get("daily_sales")
    price = kf.get("price") if kf.get("price") is not None else js.get("price")
    revenue = (round(price * monthly, 2)
               if isinstance(price, (int, float)) and isinstance(monthly, (int, float))
               else js.get("monthly_revenue"))
    row = {
        "asin": child["asin"],
        "parent_asin": parent_row["asin"],
        "row_type": "Variation",
        # Static — inherited from parent (Rule 7 allows).
        "brand": parent_row["brand"],
        "title": (f"{parent_row['title']} — {child['attributes_str']}"
                  if parent_row.get("title") not in (None, UNKNOWN) and child.get("attributes_str")
                  else parent_row["title"]),
        "design": parent_row["design"],
        "subcategory": parent_row["subcategory"],
        "category": parent_row["category"],
        "tier": _v(js.get("product_tier")) if js.get("product_tier") else parent_row["tier"],
        "material": parent_row["material"],
        "sf1": parent_row["sf1"], "sf2": parent_row["sf2"], "sf3": parent_row["sf3"],
        "packaging": parent_row["packaging"],
        # Physical dims/weight: the child's own Keepa/JS data, else inherit.
        "dimensions": _v(kf.get("dimensions") or js.get("dimensions") or parent_row["dimensions"]),
        "weight": _v(kf.get("weight") or js.get("weight") or parent_row["weight"]),
        "fees": _v(js.get("amazon_fees") or kfields.get("fba_pick_pack_fee") or None)
                if (js.get("amazon_fees") or kfields.get("fba_pick_pack_fee")) else parent_row["fees"],
        "fulfillment": _v(kf.get("seller_type") or _fulfillment(js_row) or parent_row["fulfillment"]),
        "snl_prime": _v(_snl_prime(price)),
        # Variation-level — child's own data (Keepa → JS), else Unknown.
        "variation": _v(child.get("attributes_str")),
        "min_price": _v(kf.get("min_price")),
        "consistent_price": _v(kf.get("consistent_price")),
        "max_price": _v(kf.get("max_price")),
        "monthly_sales": _v(monthly),
        "daily_sales": _v(daily),
        "monthly_revenue": _v(revenue),
        "price": _v(price),
        "rating": _v(kf.get("rating") if kf.get("rating") is not None else js.get("rating")),
        "reviews": _v(kf.get("reviews") if kf.get("reviews") is not None else js.get("reviews")),
        "bsr": _v(kf.get("bsr") or js.get("bsr")),
        "sellers": _v(kf.get("num_sellers") if kf.get("num_sellers") is not None else js.get("num_sellers")),
        "date_first_available": _v(kf.get("date_first_available") or js.get("date_first_available")),
        # Colour / product size / pack size: the variation's own value, else the
        # parent's classified value. Each column stays its own kind of data.
        "color": _v(v_color or parent_row.get("color")),
        "size": _v(v_size or parent_row.get("size")),
        "pack_size": _v(v_pack or parent_row.get("pack_size")),
        "link": f"https://www.amazon.co.uk/dp/{child['asin']}",
    }
    return row


def _attr_pick(attrs_str: Optional[str], keys) -> Optional[str]:
    """Pull a value like 'Colour: Red' out of the joined attributes string."""
    if not attrs_str:
        return None
    for part in attrs_str.split(";"):
        if ":" in part:
            k, v = part.split(":", 1)
            if any(key in k.strip().lower() for key in keys):
                return v.strip()
    return None


# ──────────────────────────────────────────────────────────────────────────────
# Workbook writer — fills the real PES template's 'Critical sheet' tab in place
# so the output matches PES UK.xlsx exactly (all other tabs preserved).
# ──────────────────────────────────────────────────────────────────────────────

def _cell_val(v):
    """Coerce numpy scalars to native Python for openpyxl; pass through else."""
    if v is None:
        return None
    item = getattr(v, "item", None)
    if callable(item):
        try:
            return v.item()
        except Exception:  # noqa: BLE001
            return v
    return v


def _col(idx: int) -> str:
    from openpyxl.utils import get_column_letter
    return get_column_letter(idx)


def _locate_sheet_part(zf, name: str = CRITICAL_SHEET, required: bool = True):
    """Return the zip part path of the worksheet XML named *name* (None when it
    isn't present and not *required*)."""
    wbxml = zf.read("xl/workbook.xml").decode("utf-8")
    relsxml = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    m = (re.search(r'<sheet[^>]*name="' + re.escape(name) + r'"[^>]*r:id="([^"]+)"', wbxml)
         or re.search(r'<sheet[^>]*r:id="([^"]+)"[^>]*name="' + re.escape(name) + r'"', wbxml))
    if not m:
        if required:
            raise Phase3ApiError(f"'{name}' tab not found in the PES workbook.")
        return None
    tm = re.search(r'Id="' + re.escape(m.group(1)) + r'"[^>]*Target="([^"]+)"', relsxml)
    if not tm:
        if required:
            raise Phase3ApiError(f"Could not resolve the '{name}' worksheet part.")
        return None
    return "xl/" + tm.group(1).replace("../", "").lstrip("/")


def _upsert_row_cells(xml: str, rownum: int, new_cells: Dict[str, str]) -> str:
    """Insert/replace cells (given as full ``<c …>…</c>`` strings keyed by ref)
    into row *rownum*, keeping the row's cells in column order and preserving all
    existing cells. Used to write ASIN keys into cells the template never created."""
    from openpyxl.utils import column_index_from_string
    m = re.search(rf'(<row r="{rownum}"[^>]*>)(.*?)(</row>)', xml, re.S)
    if not m:
        return xml
    open_tag, inner, close = m.groups()
    cells: Dict[str, str] = {}
    for cm in re.finditer(r'<c r="([A-Z]+\d+)"(?:[^>]*?/>|[^>]*?>.*?</c>)', inner, re.S):
        cells[cm.group(1)] = cm.group(0)
    cells.update(new_cells)
    ordered = "".join(cells[r] for r in sorted(
        cells, key=lambda ref: column_index_from_string(re.match(r'([A-Z]+)', ref).group(1))))
    return xml[:m.start()] + open_tag + ordered + close + xml[m.end():]


def _fill_sheet_xml(xml: str, cells: Dict[str, object], clears: set) -> str:
    """Single-pass edit of a worksheet's XML: set the value of each cell in
    *cells* and blank each cell in *clears*, PRESERVING every cell's existing
    style (`s="…"`). Numbers are written as `<v>`, strings as inline strings
    (so sharedStrings.xml is never touched). Only cells that already exist in the
    template are affected — its data region is fully pre-styled, so that covers
    everything Phase 4 writes."""
    from xml.sax.saxutils import escape

    def repl(m):
        ref, attrs = m.group("ref"), m.group("attrs")
        if ref not in cells and ref not in clears:
            return m.group(0)
        sm = re.search(r'\ss="(\d+)"', attrs)
        style = f' s="{sm.group(1)}"' if sm else ""
        if ref in clears:
            return f'<c r="{ref}"{style}/>'
        val = cells[ref]
        if isinstance(val, bool):
            val = int(val)
        if isinstance(val, (int, float)):
            return f'<c r="{ref}"{style}><v>{val}</v></c>'
        return (f'<c r="{ref}"{style} t="inlineStr"><is>'
                f'<t xml:space="preserve">{escape(str(val))}</t></is></c>')

    pat = re.compile(r'<c r="(?P<ref>[A-Z]+\d+)"(?P<attrs>[^>/]*)(?:/>|>.*?</c>)', re.DOTALL)
    return pat.sub(repl, xml)


H10_SHEET = "H10 basic data"
# H10 column key -> 1-based column (A..O). Sourced from the Phase-3 Xray export.
_H10_COLS = [
    ("title", 1), ("asin", 2), ("brand", 3), ("price", 4), ("asin_sales", 5),
    ("asin_revenue", 6), ("bsr", 7), ("fees", 8), ("active_sellers", 9),
    ("rating", 10), ("review_count", 11), ("images", 12), ("review_velocity", 13),
    ("buy_box", 14), ("category", 15),
]
# (date cell, keyword cell, Amazon-URL cell, first data row, last data row) per KW.
_H10_SECTIONS = [
    ("B5", "B6", "M3", 8, 64), ("B68", "B69", "M4", 71, 127),
    ("B129", "B130", "M5", 132, 188),
]


def _h10_cells(xray_by_kw: dict) -> Dict[str, object]:
    """Build the H10 Basic Data cell writes from up to 3 keywords' Xray exports:
    date + keyword + Amazon URL per section, then the top-57 products (by sales)
    filling the pre-styled data rows."""
    import pandas as pd
    from urllib.parse import quote_plus

    cells: Dict[str, object] = {}
    today = date.today().strftime("%Y-%m-%d")
    for (kw, df), (dcell, kcell, ucell, r0, r1) in zip(list(xray_by_kw.items())[:3], _H10_SECTIONS):
        cells[dcell] = today
        cells[kcell] = kw
        cells[ucell] = "https://www.amazon.co.uk/s?k=" + quote_plus(str(kw))
        d = df.copy()
        if "asin_sales" in d.columns:
            d = d.sort_values("asin_sales", ascending=False, na_position="last")
        for i, (_, prow) in enumerate(d.head(r1 - r0 + 1).iterrows()):
            rr = r0 + i
            for key, ci in _H10_COLS:
                v = prow.get(key)
                if v is None or (isinstance(v, float) and pd.isna(v)):
                    continue
                cells[f"{_col(ci)}{rr}"] = _cell_val(v)
    return cells


def _filter_h10_same_type(client, target_asin, uni, xray_by_kw: dict, log: list) -> dict:
    """Phase-5 Step 16.1: for each keyword's Xray products, keep only those that
    are the SAME product type as the target (Claude), before the top-57 pick.
    Derives the target profile once from its title/category (JS sheet → Xray).
    Falls back to the unfiltered set if the profile or a filter call fails, so
    H10 is never left empty."""
    # Target title/category for the profile: JS sheet first, else any Xray row.
    tj = uni["js_by_asin"].get(target_asin, {}) if uni else {}
    title = tj.get("title")
    category = tj.get("category")
    if not title:
        for df in xray_by_kw.values():
            if {"asin", "title"}.issubset(df.columns):
                hit = df[df["asin"] == target_asin]
                if not hit.empty:
                    title = hit.iloc[0].get("title")
                    category = category or hit.iloc[0].get("category")
                    break
    if not title:
        log.append("H10 same-type filter skipped — no target title to build a profile from.")
        return xray_by_kw
    try:
        profile = derive_product_profile(client, title=title, category=category or "")
    except Exception as e:  # noqa: BLE001
        log.append(f"H10 same-type filter skipped — profile build failed: {e}")
        return xray_by_kw

    out = {}
    for kw, df in xray_by_kw.items():
        if not {"asin", "title"}.issubset(df.columns) or df.empty:
            out[kw] = df
            continue
        try:
            kept = set(filter_same_product_type(client, profile, df[["asin", "title"]]).same_type_asins)
            fdf = df[df["asin"].isin(kept)]
            log.append(f"H10 '{kw}': {len(fdf)}/{len(df)} same product type"
                       + (f" (only {len(fdf)}<57)" if 0 < len(fdf) < 57 else ""))
            out[kw] = fdf if not fdf.empty else df  # never blank a section
        except Exception as e:  # noqa: BLE001
            log.append(f"H10 '{kw}' same-type filter failed ({e}) — kept all rows.")
            out[kw] = df
    return out


# ── Phase 6: KWs Complete/Filtered Data from the multi-ASIN Cerebro export ────
KWS_COMPLETE_SHEET = "KWs Complete Data"
KWS_FILTERED_SHEET = "KWs Filtered Data"
KWS_HEADER_ROW = 4
KWS_DATA_START = 5
KWS_ASIN_ROW = 3          # ASINs go in G3:P3 (Seller 1..10)
KWS_SELLER1_COL = 7       # G
_KWS_MAX_ROWS = 400       # cap KWs Complete Data (keywords) for a sane file size
# Key Evaluation reuses the same KW matrix (the launch keywords), offset lower:
# ASIN row 49, headers row 50, data from row 51 (feeds SUMIF/COUNTIF in rows 40-47).
KEY_EVAL_SHEET = "key Evaluation"
KEY_EVAL_ASIN_ROW = 49
KEY_EVAL_DATA_START = 51


def add_kws_to_workbook(workbook_path, cerebro_file, target_asin, launch_keywords,
                        key_eval_images=None, log: Optional[list] = None) -> str:
    """Phase 6, run AFTER the Critical Sheet (so the top-10 ASINs are known):
    fill the KWs Complete/Filtered + Key Evaluation tabs of an already-built
    Phase-4 workbook from a Cerebro export, in place (and, if given, the Key
    Evaluation G1:P1 product images). Returns the workbook path."""
    import os
    import zipfile

    src = Path(workbook_path)
    if not src.exists():
        raise Phase3ApiError(f"Phase-4 workbook not found: {workbook_path}")
    df = _read_cerebro(cerebro_file)

    edits = {}
    with zipfile.ZipFile(src) as zf:
        for name, arow, drow, filt in (
            (KWS_COMPLETE_SHEET, KWS_ASIN_ROW, KWS_DATA_START, False),
            (KWS_FILTERED_SHEET, KWS_ASIN_ROW, KWS_DATA_START, True),
            (KEY_EVAL_SHEET, KEY_EVAL_ASIN_ROW, KEY_EVAL_DATA_START, True),
        ):
            part = _locate_sheet_part(zf, name, required=False)
            if not part:
                continue
            xml = zf.read(part).decode("utf-8")
            new_xml, n = _kws_fill_xml(xml, df, target_asin, launch_keywords,
                                       filtered=filt, asin_row=arow, data_start=drow)
            if name == KEY_EVAL_SHEET and key_eval_images:
                new_xml = _upsert_row_cells(new_xml, 1, _key_eval_image_cells(key_eval_images))
            edits[part] = new_xml
            if log is not None:
                log.append(f"{name}: {n} keyword(s).")

    tmp = src.with_suffix(".kwstmp.xlsx")
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in edits:
                data = edits[item.filename].encode("utf-8")
            zout.writestr(item, data)
    os.replace(tmp, src)  # overwrite the deliverable in place
    return str(src)


def _read_cerebro(path: str):
    """Read a Helium 10 Cerebro (reverse-ASIN) export; normalise headers."""
    import pandas as pd
    p = str(path)
    df = pd.read_excel(p) if p.lower().endswith((".xlsx", ".xls")) else pd.read_csv(p)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _cerebro_asin_cols(df) -> List[str]:
    """The per-ASIN rank columns (headers that are ASINs) in the Cerebro export."""
    return [c for c in df.columns if _ASIN_RE.match(str(c).strip().upper())]


def _rank_num(v):
    """A rank cell → positive int, else None (dash / blank / non-numeric)."""
    import pandas as pd
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    try:
        n = int(round(float(re.sub(r"[^\d.\-]", "", str(v)) or "x")))
    except (ValueError, TypeError):
        return None
    return n if n > 0 else None


def _kws_fill_xml(xml: str, df, target_asin: str, launch_keywords, *, filtered: bool,
                  asin_row: int = KWS_ASIN_ROW, data_start: int = KWS_DATA_START):
    """Fill one KW-matrix tab: write the 10 ASINs into the seller row (G..P of
    *asin_row*) and one row per keyword from *data_start* (A=phrase, B=SV,
    C=competing, D=sponsored ASINs, E=CPS, F=ranking competitors, G=main-ASIN
    rank [Position (Rank)], H:P=the 9 competitor ASIN ranks). Missing competitor
    ranks are filled with the CPS value (v9.6 rule). Used for KWs Complete
    (any ASIN ≤30, top by SV), KWs Filtered and Key Evaluation rows 51+ (both =
    the approved launch keywords)."""
    from xml.sax.saxutils import escape

    def col(name, *alts):
        for c in (name, *alts):
            if c in df.columns:
                return c
        return None

    c_phrase = col("Keyword Phrase", "Keyword")
    c_sv = col("Search Volume")
    c_comp = col("Competing Products")
    c_spon = col("Sponsored ASINs")
    c_cps = col("Competitor Performance Score")
    c_rankc = col("Ranking Competitors (count)", "Ranking Competitors")
    c_pos = col("Position (Rank)")
    asin_cols = _cerebro_asin_cols(df)[:9]  # up to 9 competitors → H..P
    if not c_phrase:
        return xml, 0

    rows = df.to_dict("records")
    picked = []
    for r in rows:
        phrase = str(r.get(c_phrase) or "").strip()
        if not phrase:
            continue
        ranks = [_rank_num(r.get(c_pos))] + [_rank_num(r.get(a)) for a in asin_cols]
        if not any(x is not None and x <= 30 for x in ranks):
            continue  # base (KWs Complete): at least one ASIN ranks ≤30
        if filtered:  # KWs Filtered / Key Eval: keep only Search Volume > 350
            sv = _rank_num(r.get(c_sv)) if c_sv else None
            if sv is None or sv <= 350:
                continue
        picked.append(r)
    picked.sort(key=lambda r: _rank_num(r.get(c_sv)) or 0, reverse=True)
    if not filtered:
        picked = picked[:_KWS_MAX_ROWS]

    # Seller row (G..P of asin_row) — G=main (target), H..P = competitor ASINs.
    asin_slots = {KWS_SELLER1_COL: target_asin}
    for i, a in enumerate(asin_cols):
        asin_slots[KWS_SELLER1_COL + 1 + i] = a
    asin_cells = {f"{_col(ci)}{asin_row}":
                  f'<c r="{_col(ci)}{asin_row}" t="inlineStr"><is><t>{escape(str(a))}</t></is></c>'
                  for ci, a in asin_slots.items()}
    xml = _upsert_row_cells(xml, asin_row, asin_cells)

    def cell(ref, v):
        if v is None:
            return ""
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return f'<c r="{ref}"><v>{v}</v></c>'
        return f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{escape(str(v))}</t></is></c>'

    rowmap = {}   # rownum -> populated cell XML (replaces the template's empty row)
    for i, r in enumerate(picked):
        rr = data_start + i
        cps = str(r.get(c_cps)).strip() if c_cps and r.get(c_cps) is not None else None
        vals = {
            1: str(r.get(c_phrase)).strip(),
            2: _rank_num(r.get(c_sv)) if c_sv else None,
            3: _rank_num(r.get(c_comp)) if c_comp else None,
            4: _rank_num(r.get(c_spon)) if c_spon else None,
            5: cps,
            6: _rank_num(r.get(c_rankc)) if c_rankc else None,
            7: _rank_num(r.get(c_pos)) if c_pos else None,  # G = main ASIN rank
        }
        for j, a in enumerate(asin_cols):  # H..P competitor ranks (else CPS string)
            rk = _rank_num(r.get(a))
            vals[KWS_SELLER1_COL + 1 + j] = rk if rk is not None else cps
        if vals[7] is None:
            vals[7] = cps
        rowmap[rr] = "".join(cell(f"{_col(ci)}{rr}", v) for ci, v in sorted(vals.items()) if v is not None)

    xml = _populate_rows(xml, rowmap)
    return xml, len(picked)


def _key_eval_image_cells(image_urls) -> Dict[str, str]:
    """`=IMAGE(url)` cells for Key Evaluation G1:P1 — one product image per
    top-10 ASIN column (rendered by Excel 365 / Google Sheets)."""
    from xml.sax.saxutils import escape
    cells = {}
    for i, url in enumerate((image_urls or [])[:10]):
        if not url:
            continue
        ref = f"{_col(KWS_SELLER1_COL + i)}1"  # G1..P1
        cells[ref] = f'<c r="{ref}"><f>IMAGE("{escape(str(url))}",1)</f></c>'
    return cells


def _populate_rows(xml: str, rowmap: Dict[int, str]) -> str:
    """Replace the *content* of each existing ``<row r="N">`` whose number is in
    *rowmap* with the given cell XML (the KWs template pre-defines thousands of
    empty rows — we fill them in place rather than inserting duplicates)."""
    def repl(m):
        rn = int(m.group("r"))
        return f'<row r="{rn}">{rowmap[rn]}</row>' if rn in rowmap else m.group(0)
    return re.sub(r'<row r="(?P<r>\d+)"[^>]*?(?:/>|>.*?</row>)', repl, xml, flags=re.S)


def _patch_full_calc(wbxml: str) -> str:
    """Force Excel to recalculate on open so the PES formulas (Margin Calc,
    Price Analysis, …) that reference the Critical sheet pick up the new data."""
    if "<calcPr" in wbxml:
        def r(m):
            tag = m.group(0)
            if "fullCalcOnLoad" in tag:
                return re.sub(r'fullCalcOnLoad="[^"]*"', 'fullCalcOnLoad="1"', tag)
            return tag[:-2] + ' fullCalcOnLoad="1"/>' if tag.endswith("/>") else tag
        return re.sub(r'<calcPr[^>]*/?>', r, wbxml, count=1)
    return wbxml.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')


def write_critical_sheet(rows: List[dict], legend: List[dict], *, target_asin: str,
                         keyword: Optional[str], vocab: dict,
                         design_attributes: Optional[List[str]] = None,
                         xray_by_kw: Optional[dict] = None,
                         cerebro_df=None, launch_keywords=None, key_eval_images=None,
                         log: Optional[list] = None) -> str:
    """Ship a COPY of the full PES workbook with the 'Critical sheet' tab filled
    (and, when *xray_by_kw* is given, the 'H10 basic data' tab from the Phase-3
    Xray exports).

    The original `PES UK.xlsx` is never modified. Rather than round-tripping the
    workbook through openpyxl (which would strip its 14 charts / 42 pivot tables
    / drawings), we copy the file and surgically edit only the affected worksheet
    XML — so every other tab, formula, chart and pivot is preserved byte-for-byte.
    Returns the output path."""
    import zipfile

    if not PES_SOURCE.exists():
        raise Phase3ApiError(
            f"PES template not found at {PES_SOURCE}. Place 'PES UK.xlsx' in "
            f"{PES_SOURCE.parent}."
        )
    with zipfile.ZipFile(PES_SOURCE) as zf:
        part = _locate_sheet_part(zf)
        xml = zf.read(part).decode("utf-8")
        top_part = _locate_sheet_part(zf, "Top Relevant ASINs Data", required=False)
        top_xml = zf.read(top_part).decode("utf-8") if top_part else None
        h10_part = _locate_sheet_part(zf, H10_SHEET, required=False) if xray_by_kw else None
        h10_xml = zf.read(h10_part).decode("utf-8") if h10_part else None
        kc_part = kf_part = ke_part = None
        kc_xml = kf_xml = ke_xml = None
        if cerebro_df is not None:
            kc_part = _locate_sheet_part(zf, KWS_COMPLETE_SHEET, required=False)
            kc_xml = zf.read(kc_part).decode("utf-8") if kc_part else None
            kf_part = _locate_sheet_part(zf, KWS_FILTERED_SHEET, required=False)
            kf_xml = zf.read(kf_part).decode("utf-8") if kf_part else None
        if cerebro_df is not None or key_eval_images:
            ke_part = _locate_sheet_part(zf, KEY_EVAL_SHEET, required=False)
            ke_xml = zf.read(ke_part).decode("utf-8") if ke_part else None

    # Build the set of cell writes (ref -> value) + clears (blank scratch).
    cells: Dict[str, object] = {
        DATE_CELL: date.today().strftime("%Y-%m-%d"),
        KW_CELL: keyword or UNKNOWN,
    }
    clears = {f"{_col(c)}{r}" for r in range(LEGEND_START, LEGEND_END + 1)
              for c in range(26, 34)}  # sample legend scratch (Z..AG, rows 2-7)
    for i, item in enumerate(legend[: LEGEND_END - LEGEND_START + 1]):
        cells[f"{_col(LEGEND_CODE_COL)}{LEGEND_START + i}"] = item["code"]
        cells[f"{_col(LEGEND_DESC_COL)}{LEGEND_START + i}"] = item["description"]
    for j, attr in enumerate((design_attributes or [])[:4]):
        cells[f"{_col(ATTR_COL_START + j)}{ATTR_ROW}"] = attr
    for ri, row in enumerate(rows, DATA_START):
        for key, col in TEMPLATE_COL_MAP.items():
            val = _cell_val(row.get(key))
            if val is not None:
                cells[f"{_col(col)}{ri}"] = val
    clears -= set(cells)  # a written cell isn't a clear

    new_xml = _fill_sheet_xml(xml, cells, clears)

    # 'Top Relevant ASINs Data' VLOOKUPs everything from the Critical sheet keyed
    # on ASINs in B7:K7 — write our top-10 parent ASINs there so that tab (and
    # its downstream) fills on recalc.
    new_top_xml = None
    if top_xml is not None:
        parents = [str(r.get("asin")) for r in rows if r.get("row_type") == "Parent"][:10]
        key_cells = {f"{_col(2 + i)}7": f'<c r="{_col(2 + i)}7" t="inlineStr">'
                                        f'<is><t>{a}</t></is></c>'
                     for i, a in enumerate(parents)}
        if key_cells:
            new_top_xml = _upsert_row_cells(top_xml, 7, key_cells)

    # 'H10 basic data' — date/keyword/URL + top-57 products per keyword section.
    new_h10_xml = None
    if h10_xml is not None:
        new_h10_xml = _fill_sheet_xml(h10_xml, _h10_cells(xray_by_kw), set())

    # KWs Complete / Filtered Data — from the multi-ASIN Cerebro export (Phase 6).
    new_kc_xml = new_kf_xml = new_ke_xml = None
    if cerebro_df is not None:
        if kc_xml is not None:
            new_kc_xml, nkc = _kws_fill_xml(kc_xml, cerebro_df, target_asin, launch_keywords, filtered=False)
            if log is not None:
                log.append(f"KWs Complete Data: {nkc} keywords (any ASIN rank ≤30).")
        if kf_xml is not None:
            new_kf_xml, nkf = _kws_fill_xml(kf_xml, cerebro_df, target_asin, launch_keywords, filtered=True)
            if log is not None:
                log.append(f"KWs Filtered Data: {nkf} keyword(s) (SV > 350).")
    if ke_xml is not None:
        new_ke_xml = ke_xml
        if cerebro_df is not None:
            new_ke_xml, nke = _kws_fill_xml(new_ke_xml, cerebro_df, target_asin, launch_keywords,
                                            filtered=True, asin_row=KEY_EVAL_ASIN_ROW,
                                            data_start=KEY_EVAL_DATA_START)
            if log is not None:
                log.append(f"Key Evaluation rows 51+: {nke} keyword(s) (SV > 350).")
        if key_eval_images:
            new_ke_xml = _upsert_row_cells(new_ke_xml, 1, _key_eval_image_cells(key_eval_images))
            if log is not None:
                log.append(f"Key Evaluation images: {sum(1 for u in key_eval_images if u)} in G1:P1.")

    _OUT_DIR.mkdir(parents=True, exist_ok=True)
    ts = date.today().strftime("%Y%m%d")
    dest = _OUT_DIR / f"Phase4_CriticalSheet_{target_asin}_{ts}.xlsx"
    n = 1
    while dest.exists():
        dest = _OUT_DIR / f"Phase4_CriticalSheet_{target_asin}_{ts}_{n}.xlsx"
        n += 1

    # Map of zip-part -> replacement XML for every worksheet we edited.
    edited = {part: new_xml}
    if top_part and new_top_xml:
        edited[top_part] = new_top_xml
    if h10_part and new_h10_xml:
        edited[h10_part] = new_h10_xml
    if kc_part and new_kc_xml:
        edited[kc_part] = new_kc_xml
    if kf_part and new_kf_xml:
        edited[kf_part] = new_kf_xml
    if ke_part and new_ke_xml:
        edited[ke_part] = new_ke_xml

    # Rebuild the workbook: copy every part verbatim except the edited worksheets
    # and workbook.xml (force recalc). Charts/pivots/formulas untouched.
    with zipfile.ZipFile(PES_SOURCE) as zin, \
            zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in edited:
                data = edited[item.filename].encode("utf-8")
            elif item.filename == "xl/workbook.xml":
                data = _patch_full_calc(data.decode("utf-8")).encode("utf-8")
            elif re.match(r"xl/pivotCache/pivotCacheDefinition\d+\.xml$", item.filename):
                # Refresh pivots (e.g. Market Segmentation) from the new data on open.
                txt = data.decode("utf-8")
                if "refreshOnLoad" not in txt:
                    txt = re.sub(r"<pivotCacheDefinition\b",
                                 '<pivotCacheDefinition refreshOnLoad="1"', txt, count=1)
                data = txt.encode("utf-8")
            zout.writestr(item, data)
    if log is not None:
        log.append(f"Critical Sheet written into a copy of PES UK.xlsx: {dest.name}")
    return str(dest)


# ──────────────────────────────────────────────────────────────────────────────
# Orchestrator
# ──────────────────────────────────────────────────────────────────────────────

def _load_universe(js_file: str, parent_limit: int, log: list,
                   max_asins: Optional[int] = None, keepa_file: Optional[str] = None):
    """Read the JS sheet, pick the top parents by monthly sales, Keepa-pull them
    and their variation children, and return everything the propose/fill stages
    share. *max_asins* (testing only) caps the TOTAL ASINs (parents + variations)
    pulled + rendered, to conserve Keepa tokens."""
    js_df, _ = read_js_products(str(js_file))
    if "asin" not in js_df.columns:
        raise Phase3ApiError("The Jungle Scout sheet has no ASIN column.")

    df = js_df.copy()
    if "monthly_sales" in df.columns:
        df = df.sort_values("monthly_sales", ascending=False, na_position="last")
    parents = [a for a in df["asin"].tolist() if _ASIN_RE.match(str(a or ""))][:parent_limit]
    # A parent is ≥1 ASIN, so never pull more parents than the total cap.
    if max_asins:
        parents = parents[:max_asins]
    js_by_asin = {str(r["asin"]).upper(): {k: (None if (isinstance(v, float) and v != v) else v)
                                           for k, v in r.items()}
                  for _, r in df.iterrows() if _ASIN_RE.match(str(r.get("asin") or ""))}

    # ── Keepa source: uploaded export file (no-API path) OR the live Keepa API.
    # The file path is preferred when supplied; the API branch is left intact but
    # only runs when no file is given (see run_phase4 / the Phase-4 view). ──────
    parent_set = set(parents)
    recs: Dict[str, dict] = {}
    api = None
    if keepa_file:
        from .keepa_file import read_keepa_export, keepa_file_fields, keepa_file_full
        recs = read_keepa_export(keepa_file)
        pfields = {a: keepa_file_fields(recs[a], recs) for a in parents if a in recs}
        pfull = {a: keepa_file_full(recs[a]) for a in parents if a in recs}
        miss_p = [a for a in parents if a not in recs]
        log.append(f"Keepa file: {len(pfields)}/{len(parents)} parents matched"
                   + (f" — {len(miss_p)} not in the file → those cells show Unknown" if miss_p else ""))
    else:
        api = load_keepa()
        pprods = keepa_full_query(api, parents)
        log.append(f"Keepa API: {len(pprods)}/{len(parents)} parents (tokens left: {api.tokens_left})")
        pfields = {a: extract_keepa_fields(pprods[a]) for a in parents if a in pprods}
        pfull = {a: _keepa_full(pprods[a]) for a in parents if a in pprods}

    children_map: Dict[str, List[dict]] = {}
    for a in parents:
        children_map[a] = _variation_children(pfields.get(a, {}), exclude=parent_set)

    # Testing cap: keep parents + variations in order until the total hits
    # max_asins, so a run touches at most that many ASINs.
    if max_asins:
        ordered, keep = [], set()
        for a in parents:
            for x in [a] + [k["asin"] for k in children_map[a]]:
                if len(keep) < max_asins:
                    keep.add(x); ordered.append(x)
        parents = [a for a in parents if a in keep]
        children_map = {a: [k for k in children_map.get(a, []) if k["asin"] in keep]
                        for a in parents}

    all_children = [k["asin"] for a in parents for k in children_map[a]]

    if keepa_file:
        cfields = {ca: keepa_file_fields(recs[ca], recs) for ca in all_children if ca in recs}
        cfull = {ca: keepa_file_full(recs[ca]) for ca in all_children if ca in recs}
        if all_children:
            miss_c = [ca for ca in all_children if ca not in recs]
            log.append(f"Keepa file: {len(cfields)}/{len(all_children)} variation children matched"
                       + (f" — {len(miss_c)} variation ASINs are not in the file (run a Keepa "
                          f"lookup on them and re-upload a combined export to fill them)" if miss_c else ""))
    else:
        cprods = {}
        if all_children:
            cprods = keepa_full_query(api, all_children)
            log.append(f"Keepa API: {len(cprods)}/{len(all_children)} variation children "
                       f"(tokens left: {api.tokens_left})")
        cfields = {a: extract_keepa_fields(p) for a, p in cprods.items()}
        cfull = {a: _keepa_full(p) for a, p in cprods.items()}

    return {
        "parents": parents, "js_by_asin": js_by_asin,
        "pfields": pfields, "pfull": pfull,
        "children_map": children_map, "cfields": cfields, "cfull": cfull,
    }


def run_phase4(target_asin, js_file=None, *, keyword=None, approved=None,
               parent_limit=DEFAULT_PARENT_LIMIT, max_asins=None, xray_files=None,
               cerebro_file=None, launch_keywords=None, keepa_file=None):
    """Run Phase 4. Two stages:

    - ``approved is None`` → stage 'propose': Keepa-pull the universe and return
      Claude's controlled-vocabulary + Design-attribute proposal plus a variation
      summary for the user to review/edit.
    - ``approved`` (the edited vocab dict) → stage 'fill': classify, assemble
      rows (Rule 7 + Rule 0) and write the Critical Sheet.

    *max_asins* (testing only) caps the total ASINs pulled from Keepa + written
    to the sheet, to conserve tokens. Leave None in production (all ASINs).

    Errors are returned in ``error`` (never raised)."""
    log: list = []
    out = {
        "target_asin": (target_asin or "").strip().upper(), "keyword": keyword,
        "stage": "propose" if approved is None else "fill",
        "vocabulary": None, "variation_summary": None, "rows": [], "row_count": 0,
        "parent_count": 0, "variation_count": 0, "legend": [], "top10_asins": [],
        "workbook_path": None, "log": log, "error": None,
    }
    asin = out["target_asin"]
    if not asin:
        out["error"] = "No target ASIN supplied."
        return out
    if not js_file:
        out["error"] = "No Jungle Scout sheet available — run Phase 3 first."
        return out

    try:
        uni = _load_universe(js_file, parent_limit, log, max_asins=max_asins,
                             keepa_file=keepa_file)
    except (Phase3KeepaError, Phase3ApiError, FileNotFoundError, ValueError) as e:
        out["error"] = str(e)
        return out

    parents = uni["parents"]
    out["parent_count"] = len(parents)
    out["variation_count"] = sum(len(v) for v in uni["children_map"].values())
    if not parents:
        out["error"] = "No valid parent ASINs in the Jungle Scout sheet."
        return out

    try:
        client = load_client()
    except Phase3ApiError as e:
        out["error"] = str(e)
        return out

    # ── Stage 1: propose vocabulary + Design attributes ──────────────────────
    if approved is None:
        samples = []
        for a in parents[:25]:
            f = uni["pfields"].get(a, {})
            samples.append({
                "asin": a, "title": f.get("title") or uni["js_by_asin"].get(a, {}).get("title"),
                "bullets": f.get("bullets"), "description": f.get("description"),
                "category": f.get("category") or uni["js_by_asin"].get(a, {}).get("category"),
                "image_url": (f.get("image_urls") or [None])[0],
            })
        try:
            vocab = propose_vocabulary(client, samples)
        except Phase3ApiError as e:
            out["error"] = str(e)
            return out
        out["vocabulary"] = vocab.model_dump()
        out["variation_summary"] = [
            {"parent": a, "children": len(uni["children_map"].get(a, []))}
            for a in parents
        ]
        return out

    # ── Stage 2: classify + assemble + write ─────────────────────────────────
    parent_rows_for_cls = []
    for a in parents:
        f = uni["pfields"].get(a, {})
        parent_rows_for_cls.append({
            "asin": a, "title": f.get("title") or uni["js_by_asin"].get(a, {}).get("title"),
            "bullets": f.get("bullets"), "description": f.get("description"),
            "image_urls": f.get("image_urls"),
        })
    try:
        classified = classify_products(client, parent_rows_for_cls, approved)
    except Phase3ApiError as e:
        out["error"] = str(e)
        return out
    by_asin = classified["by_asin"]
    out["legend"] = classified["legend"]

    rows: List[dict] = []
    for a in parents:
        prow = _parent_row(a, uni["js_by_asin"].get(a), uni["pfields"].get(a, {}),
                           uni["pfull"].get(a, {}), by_asin.get(a))
        rows.append(prow)
        for child in uni["children_map"].get(a, []):
            ca = child["asin"]
            rows.append(_child_row(child, uni["cfields"].get(ca, {}),
                                   uni["cfull"].get(ca, {}), prow,
                                   js_row=uni["js_by_asin"].get(ca)))

    # Sort the whole Critical sheet by monthly sales, descending (user request).
    rows.sort(key=lambda r: r["monthly_sales"] if isinstance(r.get("monthly_sales"), (int, float)) else -1,
              reverse=True)

    # H10 Basic Data from the Phase-3 Xray exports (top 3 keywords), if provided.
    # Phase 5: keep only same-product-type products per keyword before top-57.
    xray_by_kw = None
    if xray_files:
        try:
            xray_by_kw = _build_per_kw(xray_files)
        except Exception as e:  # noqa: BLE001
            log.append(f"H10 Basic Data skipped — could not read Xray exports: {e}")
        if xray_by_kw:
            xray_by_kw = _filter_h10_same_type(client, asin, uni, xray_by_kw, log)

    # Cerebro export (Phase 6) → KWs Complete / Filtered Data.
    cerebro_df = None
    if cerebro_file:
        try:
            cerebro_df = _read_cerebro(cerebro_file)
        except Exception as e:  # noqa: BLE001
            log.append(f"KWs tabs skipped — could not read Cerebro export: {e}")

    out["row_count"] = len(rows)
    out["top10_asins"] = parents[:10]  # Cerebro-ready list for Phase 6
    out["rows"] = rows[:60]  # preview cap for the web view

    # Key Evaluation G1:P1 images — first Keepa image per top-10 row (sorted order).
    img_by_asin = {}
    for src in (uni.get("pfields", {}), uni.get("cfields", {})):
        for a, f in src.items():
            u = (f.get("image_urls") or [None])[0]
            if u:
                img_by_asin[a] = u
    key_eval_images = [img_by_asin.get(str(r.get("asin"))) for r in rows[:10]]
    try:
        out["workbook_path"] = write_critical_sheet(
            rows, classified["legend"], target_asin=asin, keyword=keyword, vocab=approved,
            design_attributes=approved.get("design_attributes"), xray_by_kw=xray_by_kw,
            cerebro_df=cerebro_df, launch_keywords=launch_keywords,
            key_eval_images=key_eval_images, log=log)
        log.append(f"Critical Sheet written: {len(rows)} rows "
                   f"({len(parents)} parents + {len(rows) - len(parents)} variations)"
                   + (f"; H10 Basic Data filled from {len(list(xray_by_kw)[:3])} keyword(s)"
                      if xray_by_kw else ""))
    except Exception as e:  # noqa: BLE001
        out["error"] = f"Could not write the Critical Sheet: {e}"
    return out
