"""
Keepa **file** reader — the no-API path for Phase 4.

The user exports product data from the Keepa website (Product Viewer / Product
Finder → "Export" → CSV/XLSX) and uploads it. This module parses that export
into the same per-ASIN shapes Phase 4 already consumes from the live Keepa API:

  * ``keepa_file_fields(rec, recs)`` → the ``extract_keepa_fields`` shape
    (title, brand, bullets, description, category, variation_attrs, …).
  * ``keepa_file_full(rec)``          → the ``_keepa_full`` shape
    (price/min/max/consistent, rating, reviews, seller_type, colour, dims …).

Keepa's export column names vary by UI version and locale (e.g.
``Buy Box 🚚: Current`` vs ``New: Current``), so matching is keyword-based on a
normalised header rather than exact strings. Anything genuinely absent is left
as ``None`` (Rule 0 — the caller marks it "Unknown"). The live-API path in
``phase4.py`` is untouched; this is only used when a Keepa file is supplied.
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd

_ASIN_RE = re.compile(r"^[A-Z0-9]{10}$")
_AMAZON_SELLER_HINTS = ("amazon", "amazon.co.uk", "amazon.com", "amazon eu")


def _norm(s) -> str:
    """Lower-case, strip emoji/punctuation → single-spaced tokens for matching."""
    return re.sub(r"[^a-z0-9]+", " ", str(s).lower()).strip()


def _num(v) -> Optional[float]:
    """Parse a Keepa numeric cell (currency/commas/'-'/blank) → float or None."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return None if (isinstance(v, float) and v != v) else float(v)
    s = str(v).strip()
    if not s or s.lower() in {"-", "–", "—", "n/a", "na", "nan", "none"}:
        return None
    s = re.sub(r"[^0-9.\-]", "", s.replace(",", ""))
    if s in {"", "-", "."}:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _int_kmb(v) -> Optional[int]:
    """Parse counts like '1K+ bought', '2,300', '1.2M' → int, or None."""
    if v is None:
        return None
    if isinstance(v, (int, float)) and not (isinstance(v, float) and v != v):
        return int(v)
    s = str(v).lower().replace(",", "").strip()
    m = re.search(r"(\d+(?:\.\d+)?)\s*([km])?", s)
    if not m:
        return None
    n = float(m.group(1))
    mult = {"k": 1_000, "m": 1_000_000}.get(m.group(2) or "", 1)
    return int(n * mult)


class _Cols:
    """Keyword matcher over a DataFrame's headers (normalised)."""

    def __init__(self, df: pd.DataFrame):
        self._map = {c: _norm(c) for c in df.columns}

    def find(self, all_of=(), any_of=(), none_of=()) -> Optional[str]:
        for col, n in self._map.items():
            if all(t in n for t in all_of) \
               and (not any_of or any(t in n for t in any_of)) \
               and not any(t in n for t in none_of):
                return col
        return None

    def find_all(self, all_of=(), none_of=()) -> List[str]:
        out = []
        for col, n in self._map.items():
            if all(t in n for t in all_of) and not any(t in n for t in none_of):
                out.append(col)
        return out


def _price(row, cols: _Cols, band: str) -> Optional[float]:
    """Best price for a band ('current'/'lowest'/'highest'/'avg') from Buy Box,
    then New, then Amazon columns."""
    band_terms = {
        "current": (["current"], ["avg", "lowest", "highest", "min", "max", "count", "days"]),
        "lowest":  (["lowest"], ["count"]),
        "highest": (["highest"], ["count"]),
        "avg":     (["avg"], ["count"]),
    }[band]
    want, block = band_terms
    for base in (["buy", "box"], ["new"], ["amazon"]):
        # avoid 'used'/'warehouse'/'collectible' price streams
        col = cols.find(all_of=base + want,
                        none_of=list(block) + ["used", "warehouse", "collectible", "offer"])
        if col is not None:
            v = _num(row.get(col))
            if v is not None:
                return round(v, 2)
    if band == "lowest":
        return _price(row, cols, "current")
    if band == "highest":
        return _price(row, cols, "current")
    if band == "avg":
        return _price(row, cols, "current")
    return None


def _seller_type(row, cols: _Cols) -> Optional[str]:
    """FBA/FBM from Keepa columns. Amazon-as-seller → FBA (fulfilled by Amazon)."""
    isfba = cols.find(all_of=["buy", "box"], any_of=["fba"])
    if isfba:
        raw = str(row.get(isfba) or "").strip().lower()
        if raw in {"true", "yes", "1", "fba"}:
            return "FBA"
        if raw in {"false", "no", "0", "fbm"}:
            return "FBM"
    seller = cols.find(all_of=["buy", "box", "seller"]) or cols.find(all_of=["seller"], none_of=["count", "rank", "feedback"])
    if seller:
        raw = str(row.get(seller) or "").strip().lower()
        if any(h in raw for h in _AMAZON_SELLER_HINTS):
            return "FBA"
    ful = cols.find(all_of=["fulfil"])  # fulfillment / fulfilment
    if ful:
        raw = str(row.get(ful) or "").strip().upper()
        if "FBA" in raw or "AMAZON" in raw:
            return "FBA"
        if "FBM" in raw or "MERCHANT" in raw or "SELLER" in raw:
            return "FBM"
    return None


def _dims(row, cols: _Cols) -> Optional[str]:
    """'L x W x H cm' from package dimension columns, else a single dimension
    string if that's all Keepa exported."""
    parts = []
    for axis in ("length", "width", "height"):
        c = cols.find(all_of=["package", axis]) or cols.find(all_of=[axis], none_of=["title"])
        if not c:
            return _passthrough(row, cols.find(all_of=["dimension"]))
        v = _num(row.get(c))
        if v is None:
            return _passthrough(row, cols.find(all_of=["dimension"]))
        # header unit → cm
        n = _norm(c)
        if "mm" in n:
            v /= 10.0
        elif "inch" in n or " in " in f" {n} ":
            v *= 2.54
        parts.append(v)
    return " x ".join(f"{p:.1f}" for p in parts) + " cm"


def _passthrough(row, col) -> Optional[str]:
    if not col:
        return None
    s = str(row.get(col) or "").strip()
    if not s or s.lower() in {"nan", "none", "na", "-"}:
        return None
    return s


def _weight(row, cols: _Cols) -> Optional[str]:
    c = cols.find(all_of=["package", "weight"]) or cols.find(all_of=["item", "weight"]) \
        or cols.find(all_of=["weight"], none_of=["title"])
    if not c:
        return None
    v = _num(row.get(c))
    if v is None:
        return None
    n = _norm(c)
    if "kg" in n:
        kg = v
    elif " g" in f" {n}" or "gram" in n:
        kg = v / 1000.0
    elif "oz" in n:
        kg = v * 0.0283495
    elif "lb" in n or "pound" in n:
        kg = v * 0.453592
    else:
        kg = v / 1000.0 if v > 100 else v  # heuristic: big number ⇒ grams
    return f"{kg:.2f} kg"


def _rating(row, cols: _Cols) -> Optional[float]:
    c = cols.find(all_of=["review", "rating"]) or cols.find(all_of=["rating"], none_of=["count"])
    v = _num(row.get(c)) if c else None
    if v is None:
        return None
    return round(v / 10.0, 1) if v > 5 else round(v, 1)  # Keepa 0-50 → 0-5


def _variation_asins(row, cols: _Cols, self_asin: str) -> List[str]:
    c = cols.find(all_of=["variation"], any_of=["asin", "asins"]) \
        or cols.find(all_of=["variations"], none_of=["count", "number"])
    if not c:
        return []
    raw = str(row.get(c) or "")
    out, seen = [], set()
    for tok in re.split(r"[,\s;|]+", raw.upper()):
        tok = tok.strip()
        if _ASIN_RE.match(tok) and tok != self_asin and tok not in seen:
            seen.add(tok)
            out.append(tok)
    return out


def _bullets(row, cols: _Cols) -> List[str]:
    out = []
    for c in cols.find_all(all_of=["feature"], none_of=["description features count"]):
        s = str(row.get(c) or "").strip()
        if s and s.lower() not in {"nan", "-"}:
            out.append(s)
    if not out:
        c = cols.find(all_of=["description", "features"]) or cols.find(all_of=["bullet"])
        if c:
            s = str(row.get(c) or "").strip()
            if s:
                out = [p.strip() for p in re.split(r"[\r\n]+|•", s) if p.strip()]
    return out


def _first_url(row, cols: _Cols) -> Optional[str]:
    c = cols.find(all_of=["image"], none_of=["count", "number"])
    if not c:
        return None
    s = str(row.get(c) or "").strip()
    if not s:
        return None
    return s.split(";")[0].split(",")[0].strip() or None


def read_keepa_export(path: str) -> Dict[str, dict]:
    """Parse a Keepa CSV/XLSX export → ``{ASIN: normalised_record}``. Rows without
    a valid ASIN are dropped. Every record carries the parsed fields both Phase-4
    adapters need; missing values are ``None``."""
    p = Path(str(path).strip().strip('"').strip("'"))
    if not p.exists():
        raise FileNotFoundError(f"Keepa file not found: {p}")
    if p.suffix.lower() == ".csv":
        df = pd.read_csv(p, dtype=str, keep_default_na=False)
    elif p.suffix.lower() in (".xlsx", ".xls"):
        df = pd.read_excel(p, dtype=str)
    else:
        raise ValueError(f"Unsupported Keepa file '{p.suffix}'. Use .csv or .xlsx.")

    cols = _Cols(df)
    asin_col = cols.find(all_of=["asin"], none_of=["parent", "variation"]) or cols.find(all_of=["asin"])
    if not asin_col:
        raise ValueError("No ASIN column found — is this a Keepa product export?")

    cat_root = cols.find(all_of=["categor", "root"])
    cat_sub = cols.find(all_of=["categor", "sub"])
    cat_tree = cols.find(all_of=["categor", "tree"]) or cols.find(all_of=["category"], none_of=["sub", "root", "id", "bsr", "rank"])
    title_c = cols.find(all_of=["title"]) or cols.find(all_of=["product", "name"])
    brand_c = cols.find(all_of=["brand"]) or cols.find(all_of=["manufacturer"])
    desc_c = cols.find(all_of=["description"], none_of=["features"])
    color_c = cols.find(all_of=["colour"]) or cols.find(all_of=["color"])
    size_c = cols.find(all_of=["size"], none_of=["package", "tier", "count", "file"])
    bsr_c = cols.find(all_of=["sales", "rank", "current"]) or cols.find(all_of=["sales", "rank"], none_of=["sub", "drops", "avg", "90"])
    reviews_c = cols.find(all_of=["review", "count"]) or cols.find(all_of=["ratings", "count"]) or cols.find(all_of=["review"], any_of=["count"])
    sellers_c = cols.find(all_of=["offer", "count", "current"]) or cols.find(all_of=["offer", "count"]) or cols.find(all_of=["seller", "count"])
    monthly_c = cols.find(all_of=["bought", "past", "month"]) or cols.find(all_of=["monthly", "sold"]) or cols.find(all_of=["bought"])
    date_c = cols.find(all_of=["listed", "since"]) or cols.find(all_of=["date", "first"]) or cols.find(all_of=["tracking", "since"])
    parent_c = cols.find(all_of=["parent", "asin"])
    fee_c = cols.find(all_of=["fba"], any_of=["fee", "pick"]) or cols.find(all_of=["pick", "pack"])

    recs: Dict[str, dict] = {}
    for _, row in df.iterrows():
        asin = str(row.get(asin_col) or "").strip().upper()
        if not _ASIN_RE.match(asin):
            continue
        rec = {
            "asin": asin,
            "parent_asin": (str(row.get(parent_c)).strip().upper() if parent_c and row.get(parent_c) else None),
            "title": _passthrough(row, title_c),
            "brand": _passthrough(row, brand_c),
            "description": _passthrough(row, desc_c),
            "bullets": _bullets(row, cols),
            "image_url": _first_url(row, cols),
            "color": _passthrough(row, color_c),
            "size": _passthrough(row, size_c),
            "category": _passthrough(row, cat_root) or _passthrough(row, cat_tree),
            "subcategory": _passthrough(row, cat_sub) or _passthrough(row, cat_tree),
            "price": _price(row, cols, "current"),
            "min_price": _price(row, cols, "lowest"),
            "max_price": _price(row, cols, "highest"),
            "consistent_price": _price(row, cols, "avg"),
            "rating": _rating(row, cols),
            "reviews": _int_kmb(row.get(reviews_c)) if reviews_c else None,
            "num_sellers": _int_kmb(row.get(sellers_c)) if sellers_c else None,
            "monthly_sold": _int_kmb(row.get(monthly_c)) if monthly_c else None,
            "bsr": int(_num(row.get(bsr_c))) if bsr_c and _num(row.get(bsr_c)) else None,
            "date_first_available": _passthrough(row, date_c),
            "seller_type": _seller_type(row, cols),
            "dimensions": _dims(row, cols),
            "weight": _weight(row, cols),
            "fba_pick_pack_fee": _num(row.get(fee_c)) if fee_c else None,
            "variation_asins": _variation_asins(row, cols, asin),
        }
        recs[asin] = rec
    return recs


def keepa_file_fields(rec: dict, recs: Dict[str, dict]) -> dict:
    """``extract_keepa_fields``-shaped dict from a Keepa-file record. Variation
    attributes are synthesised from each child's own Colour/Size cells so Rule 7
    (variation uses its own data) still holds."""
    variation_attrs = []
    for ca in rec.get("variation_asins", []):
        child = recs.get(ca) or {}
        attrs = []
        if child.get("color"):
            attrs.append({"dimension": "Colour", "value": child["color"]})
        if child.get("size"):
            attrs.append({"dimension": "Size", "value": child["size"]})
        variation_attrs.append({"asin": ca, "attributes": attrs})
    return {
        "asin": rec.get("asin"),
        "parent_asin": rec.get("parent_asin"),
        "title": rec.get("title"),
        "brand": rec.get("brand"),
        "price": rec.get("price"),
        "list_price": None,
        "bsr": rec.get("bsr"),
        "bsr_90d": None,
        "bullets": list(rec.get("bullets") or []),
        "description": rec.get("description"),
        "image_urls": [rec["image_url"]] if rec.get("image_url") else [],
        "images_count": 1 if rec.get("image_url") else 0,
        "variations_count": len(variation_attrs),
        "variation_attrs": variation_attrs,
        "fba_pick_pack_fee": rec.get("fba_pick_pack_fee"),
        "referral_pct": None,
        "category": rec.get("category"),
        "subcategory": rec.get("subcategory"),
        "coupon": None,
        "monthly_sold": rec.get("monthly_sold"),
        "brand_store_url": None,
        "listed_since": rec.get("date_first_available"),
        "item_weight_g": None,
        "package_weight_g": None,
        "a_plus": None,
        "has_video": None,
        "buybox_seller_id": None,
        "seller_feedback": None,
    }


def keepa_file_full(rec: dict) -> dict:
    """``_keepa_full``-shaped dict (market data) from a Keepa-file record."""
    return {
        "price": rec.get("price"),
        "min_price": rec.get("min_price"),
        "max_price": rec.get("max_price"),
        "consistent_price": rec.get("consistent_price"),
        "rating": rec.get("rating"),
        "reviews": rec.get("reviews"),
        "num_sellers": rec.get("num_sellers"),
        "monthly_sold": rec.get("monthly_sold"),
        "seller_type": rec.get("seller_type"),
        "date_first_available": rec.get("date_first_available"),
        "color": rec.get("color"),
        "bsr": rec.get("bsr"),
        "dimensions": rec.get("dimensions"),
        "weight": rec.get("weight"),
    }
