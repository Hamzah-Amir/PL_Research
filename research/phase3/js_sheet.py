"""
JS-sheet builder — replicates the Jungle Scout keyword-search export from our
own data: ASINs scraped from Amazon page 1 (`asin_scraper`) -> Keepa **Product
Request** endpoint (batched, ≤100 ASINs/call) -> CSV sheet in `output/`
(summary averages are returned/printed, not written into the CSV).

Column derivations (verified against a real JS export + Keepa export for the
same ASIN, 2026-06-12):
  Price            stats current NEW -> AMAZON -> BUY_BOX_SHIPPING
  Monthly Units    `monthlySold` (Amazon's own "bought in past month" badge —
                   bucketed floor, NOT JS's BSR-model estimate; numbers differ).
                   Fallback: 30-day sales-rank drops, flagged in `Sales Source`.
  Amazon Fees      FBA pick&pack fee + price x referral % (both from Keepa)
  Net Revenue      price - fees (per unit, like JS)
  Product Tier     computed from package dims/weight via Amazon UK size tiers
  LQS (est)        our own 1-10 listing-quality estimate (images, bullets,
                   title, description, rating, reviews) — JS's LQS is
                   proprietary and cannot be reproduced exactly.
Rule 0: anything Keepa doesn't supply stays blank — never guessed.

Token cost: ~2 tokens/ASIN (base 1 + rating 1; `--buybox` adds ~2 more for
buy-box price/fulfilment). Responses are cached (depth "sheet") in
`output/keepa_cache/`; entries already in the phase-3 deep cache are borrowed
for free. `--limit N` caps uncached ASINs sent to the API (token guard).
"""

from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Dict, List, Optional

from research.phase3.keepa_client import (
    DOMAIN,
    Phase3KeepaError,
    _load_cache,
    _save_cache,
    load_keepa,
)

_ROOT = Path(__file__).resolve().parent.parent.parent  # project root (above research/)
_OUT_DIR = _ROOT / "output"
_ASIN_RE = re.compile(r"^[A-Z0-9]{10}$")

# Final sheet columns — the JS export's columns, plus three honest extras.
COLUMNS = [
    "Position", "ASIN", "Product Name", "Brand", "Price", "Monthly Units Sold",
    "Daily Units Sold", "Monthly Revenue", "Date First Available", "Net Revenue",
    "Star Rating", "Reviews", "Amazon Fees", "BSR", "LQS (est)", "Fulfillment",
    "No. of Sellers", "Category", "Product Tier", "Dimensions", "Weight (kg)",
    "Link", "Sponsored", "Sales Source",
]

# Amazon UK FBA size tiers: (name, max_kg, dims_cm sorted descending).
_UK_TIERS = [
    ("Light envelope", 0.10, (33, 23, 2.5)),
    ("Standard envelope", 0.46, (33, 23, 2.5)),
    ("Large envelope", 0.96, (33, 23, 4)),
    ("Extra-Large envelope", 0.96, (33, 23, 6)),
    ("Small parcel", 3.9, (35, 25, 12)),
    ("Standard parcel", 11.9, (45, 34, 26)),
]


def uk_size_tier(l_cm, w_cm, h_cm, kg) -> Optional[str]:
    """Amazon UK size tier from package dims (cm) + weight (kg). None if unknown."""
    if not all(isinstance(v, (int, float)) and v > 0 for v in (l_cm, w_cm, h_cm, kg)):
        return None
    dims = sorted((l_cm, w_cm, h_cm), reverse=True)
    for name, max_kg, lims in _UK_TIERS:
        if kg <= max_kg and all(d <= m for d, m in zip(dims, lims)):
            return name
    return "Oversize"


def _lqs_estimate(images, bullets, title, description, rating, reviews) -> int:
    """Our own 1-10 listing-quality score (JS's LQS is proprietary)."""
    pts = 0
    pts += 3 if images >= 7 else 2 if images >= 5 else 1 if images >= 1 else 0
    pts += 2 if bullets >= 5 else 1 if bullets >= 3 else 0
    tl = len(title or "")
    pts += 2 if tl >= 80 else 1 if tl >= 40 else 0
    pts += 1 if (description or "").strip() else 0
    pts += 1 if isinstance(rating, (int, float)) and rating >= 4.3 else 0
    pts += 1 if isinstance(reviews, (int, float)) and reviews >= 100 else 0
    return min(pts, 10)


def _cur(current: dict, key: str):
    v = current.get(key)
    return v if isinstance(v, (int, float)) and v != -1 else None


def _listed_since(product: dict) -> Optional[str]:
    raw = product.get("listedSince")
    if not isinstance(raw, (int, float)) or raw <= 0:
        return None
    try:
        import keepa
        return keepa.keepa_minutes_to_time(raw).strftime("%Y-%m-%d")
    except Exception:  # noqa: BLE001
        return None


def _fulfillment(product: dict, stats: dict, current: dict) -> Optional[str]:
    bb_fba = stats.get("buyBoxIsFBA")  # only present with buybox=True
    if product.get("availabilityAmazon") == 0:
        return "AMZ"
    if bb_fba is True:
        return "FBA"
    if bb_fba is False:
        return "FBM"
    # Without buy-box data: a current FBA/FBM offer price is the next-best signal.
    if _cur(current, "NEW_FBA") is not None:
        return "FBA"
    if _cur(current, "NEW_FBM_SHIPPING") is not None:
        return "FBM"
    return None


def product_to_row(product: dict, marketplace: str = "co.uk") -> Dict:
    """Derive one JS-sheet row from a raw Keepa product dict (blanks = unknown)."""
    stats = product.get("stats_parsed") or {}
    current = stats.get("current") or {}

    price = _cur(current, "NEW")
    if price is None:
        price = _cur(current, "AMAZON")
    if price is None:
        price = _cur(current, "BUY_BOX_SHIPPING")

    monthly = product.get("monthlySold")
    source = "Amazon badge" if isinstance(monthly, (int, float)) and monthly > 0 else None
    if source is None:
        drops = stats.get("salesRankDrops30")
        if isinstance(drops, (int, float)) and drops > 0:
            monthly, source = drops, "BSR drops (est.)"
        else:
            monthly = None

    pick_pack = (product.get("fbaFees") or {}).get("pickAndPackFee")
    pick_pack = pick_pack / 100.0 if isinstance(pick_pack, (int, float)) else None
    ref_pct = product.get("referralFeePercentage") or product.get("referralFeePercent")
    fees = None
    if pick_pack is not None and isinstance(ref_pct, (int, float)) and price is not None:
        fees = round(pick_pack + price * ref_pct / 100.0, 2)

    rating = _cur(current, "RATING")
    reviews = _cur(current, "COUNT_REVIEWS")

    dims_mm = [product.get(k) for k in ("packageLength", "packageWidth", "packageHeight")]
    dims_cm = [round(v / 10.0, 1) if isinstance(v, (int, float)) and v > 0 else None for v in dims_mm]
    weight_g = product.get("packageWeight") or product.get("itemWeight")
    kg = round(weight_g / 1000.0, 2) if isinstance(weight_g, (int, float)) and weight_g > 0 else None

    cat_tree = product.get("categoryTree") or []
    asin = product.get("asin")
    title = product.get("title")

    return {
        "ASIN": asin,
        "Product Name": title,
        "Brand": product.get("brand") or product.get("manufacturer"),
        "Price": price,
        "Monthly Units Sold": monthly,
        "Daily Units Sold": round(monthly / 30.0, 1) if monthly is not None else None,
        "Monthly Revenue": round(monthly * price, 2) if monthly is not None and price is not None else None,
        "Date First Available": _listed_since(product),
        "Net Revenue": round(price - fees, 2) if price is not None and fees is not None else None,
        "Star Rating": rating,
        "Reviews": reviews,
        "Amazon Fees": fees,
        "BSR": _cur(current, "SALES"),
        "LQS (est)": _lqs_estimate(
            len(product.get("images") or []), len(product.get("features") or []),
            title, product.get("description"), rating, reviews,
        ),
        "Fulfillment": _fulfillment(product, stats, current),
        "No. of Sellers": _cur(current, "COUNT_NEW"),
        "Category": cat_tree[0]["name"] if cat_tree else None,
        "Product Tier": uk_size_tier(*dims_cm, kg),
        "Dimensions": (" x ".join(str(d) for d in dims_cm) + " cm") if all(d is not None for d in dims_cm) else None,
        "Weight (kg)": kg,
        "Link": f"https://www.amazon.{marketplace}/dp/{asin}",
        "Sales Source": source,
    }


def _query_sheet(api, asins: List[str], *, buybox: bool, use_cache: bool,
                 limit: Optional[int]):
    """Batch Keepa Product Request for the sheet (history+stats+rating).

    Cached per ASIN at depth "sheet"; missing ASINs are borrowed from the
    phase-3 deep cache when present (its params are a superset). `limit` caps
    uncached ASINs sent to the API.
    """
    cache = _load_cache("sheet") if use_cache else {}
    deep = _load_cache("deep") if use_cache else {}

    wanted = [a for a in asins if _ASIN_RE.fullmatch((a or "").strip().upper())]
    have, missing = {}, []
    for a in wanted:
        if a in cache:
            have[a] = cache[a]
        elif a in deep:
            have[a] = cache[a] = deep[a]
        else:
            missing.append(a)

    skipped = 0
    if limit is not None and len(missing) > limit:
        skipped = len(missing) - limit
        missing = missing[:limit]

    if missing:
        kwargs = dict(domain=DOMAIN, history=True, stats=180, rating=True, progress_bar=False)
        if buybox:
            kwargs.update(buybox=True)
        for i in range(0, len(missing), 100):  # endpoint cap: 100 ASINs/call
            try:
                products = api.query(missing[i:i + 100], **kwargs)
            except Exception as e:  # noqa: BLE001
                raise Phase3KeepaError(f"Keepa product query failed: {e}") from e
            for prod in products:
                if prod.get("asin"):
                    cache[prod["asin"]] = have[prod["asin"]] = prod
        if use_cache:
            _save_cache("sheet", cache)

    return {a: have[a] for a in wanted if a in have}, skipped


def build_js_sheet(
    keyword: Optional[str] = None,
    asins: Optional[List[str]] = None,
    *,
    marketplace: str = "co.uk",
    max_pages: int = 1,
    out_path: Optional[str] = None,
    buybox: bool = False,
    use_cache: bool = True,
    limit: Optional[int] = None,
) -> Dict:
    """Keyword (or explicit ASIN list) -> Keepa -> JS-style Excel sheet.

    Returns {out_path, count, skipped, tokens_left, summary, note}.
    """
    import pandas as pd

    scrape_meta = {}
    if asins:
        # Accept plain ASIN strings or scraper-result dicts ({asin, sponsored, position}).
        rows_meta = []
        for i, a in enumerate(asins):
            if isinstance(a, dict):
                rows_meta.append({"asin": (a.get("asin") or "").strip().upper(),
                                  "sponsored": a.get("sponsored"),
                                  "position": a.get("position") or i + 1})
            else:
                rows_meta.append({"asin": a.strip().upper(), "sponsored": None, "position": i + 1})
    else:
        if not keyword:
            raise ValueError("Provide a keyword or an ASIN list.")
        from research.phase3.asin_scraper import scrape_asins
        scraped = scrape_asins(keyword=keyword, marketplace=marketplace, max_pages=max_pages)
        if scraped.get("error") or not scraped.get("results"):
            raise Phase3KeepaError(
                f"ASIN scrape failed: {scraped.get('error') or scraped.get('note') or 'no results'}"
            )
        rows_meta = scraped["results"]
        scrape_meta = {"pages_fetched": scraped["pages_fetched"], "note": scraped.get("note", "")}

    api = load_keepa()
    products, skipped = _query_sheet(
        api, [r["asin"] for r in rows_meta], buybox=buybox, use_cache=use_cache, limit=limit,
    )

    rows = []
    for meta in rows_meta:
        prod = products.get(meta["asin"])
        if prod is None:
            continue  # not pulled (token limit) or unknown to Keepa
        row = product_to_row(prod, marketplace)
        row["Position"] = meta.get("position")
        row["Sponsored"] = meta.get("sponsored")
        rows.append(row)

    df = pd.DataFrame(rows, columns=COLUMNS)

    def _avg(col):
        s = pd.to_numeric(df[col], errors="coerce").dropna()
        return round(float(s.mean()), 2) if len(s) else None

    summary = {
        "Keyword": keyword or "(explicit ASIN list)",
        "Marketplace": f"amazon.{marketplace}",
        "Export date": date.today().isoformat(),
        "Products": len(df),
        "Avg Monthly Sales": _avg("Monthly Units Sold"),
        "Avg Monthly Revenue": _avg("Monthly Revenue"),
        "Avg Price": _avg("Price"),
        "Avg Reviews": _avg("Reviews"),
        "Avg BSR": _avg("BSR"),
        "Avg Rating": _avg("Star Rating"),
    }

    if out_path is None:
        slug = re.sub(r"[^a-z0-9]+", "_", (keyword or "asins").lower()).strip("_")
        _OUT_DIR.mkdir(parents=True, exist_ok=True)
        out_path = str(_OUT_DIR / f"js_sheet_{slug}_{date.today().isoformat()}.csv")

    # utf-8-sig so Excel opens the £/unicode columns correctly.
    df.to_csv(out_path, index=False, encoding="utf-8-sig")

    note = scrape_meta.get("note", "")
    if skipped:
        note = (note + " " if note else "") + (
            f"{skipped} ASIN(s) skipped by --limit token guard (re-run to fetch; cache keeps what's paid for)."
        )
    return {
        "out_path": out_path,
        "count": len(df),
        "skipped": skipped,
        "tokens_left": getattr(api, "tokens_left", None),
        "summary": summary,
        "note": note,
    }


def start_background(keyword: str):
    """Kick off `build_js_sheet(keyword)` in a daemon thread.

    Called by Phase 2 the moment Claude extracts the main search term from the
    target's title, so the sheet generates while the user works through the
    keyword steps. Returns ``(thread, holder)``; after ``thread.join()``,
    ``holder`` has "res" (build_js_sheet result) or "err" (the exception).
    """
    import threading

    holder: Dict = {}

    def _gen():
        try:
            holder["res"] = build_js_sheet(keyword=keyword)
        except Exception as e:  # noqa: BLE001 — surfaced by the consumer after join()
            holder["err"] = e

    t = threading.Thread(target=_gen, daemon=True)
    t.start()
    return t, holder


if __name__ == "__main__":
    import argparse
    import json
    import sys
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    ap = argparse.ArgumentParser(
        description="Build a Jungle Scout-style keyword sheet from Amazon page 1 + Keepa.",
    )
    ap.add_argument("keyword", nargs="?", help="Search keyword (omit if using --asins-file)")
    ap.add_argument("--asins-file",
                    help="ASIN list file (one per line, or an asin_scraper CSV) — skips the scrape")
    ap.add_argument("--marketplace", default="co.uk")
    ap.add_argument("--max-pages", type=int, default=1)
    ap.add_argument("--out", help="Output .csv path (default: output/js_sheet_<kw>_<date>.csv)")
    ap.add_argument("--buybox", action="store_true",
                    help="Also pull buy-box data (better Price/Fulfillment, ~+2 tokens/ASIN)")
    ap.add_argument("--limit", type=int, help="Max uncached ASINs to send to Keepa (token guard)")
    ap.add_argument("--no-cache", action="store_true")
    a = ap.parse_args()

    asins = None
    if a.asins_file:
        text = Path(a.asins_file).read_text(encoding="utf-8-sig")
        if a.asins_file.lower().endswith(".csv"):
            import csv as _csv
            import io
            asins = [{"asin": r.get("asin"), "position": int(r["position"]) if r.get("position") else None,
                      "sponsored": r.get("sponsored") in ("True", "true", "1")}
                     for r in _csv.DictReader(io.StringIO(text)) if r.get("asin")]
        else:
            asins = [ln.strip() for ln in text.splitlines() if ln.strip()]
    res = build_js_sheet(
        keyword=a.keyword, asins=asins, marketplace=a.marketplace, max_pages=a.max_pages,
        out_path=a.out, buybox=a.buybox, use_cache=not a.no_cache, limit=a.limit,
    )
    print(json.dumps(res, ensure_ascii=False, indent=2))
