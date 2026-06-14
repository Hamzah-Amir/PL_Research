"""
PL Research Tool — Phase 3: Competitor Identification + Analysis
Entry point (PDF Steps 7-11).

Flow:
  1. Target ASIN + locked keywords carry in from Phase 2 (or are prompted when
     run standalone).
  2. Step 7 inputs: the JS-style keyword sheet — AUTO-GENERATED via Amazon
     page-1 scrape + Keepa from the MAIN SEARCH TERM (Claude extracts it from
     the target's title in Phase 2, which starts the generation in the
     background right then, so the user never waits; manual JS CSV upload
     remains the fallback) — and one H10 Xray *Products* export per launch
     keyword (the Sponsored column drives the Sponsored Products / Sponsored
     Brands scoring; it is keyword-specific).
  3. Step 8: build the candidate pool (JS + Xray), Claude filters to the same
     product type, rank survivors by monthly sales -> top 3 (FBA + FBM eligible).
     A ranked table is printed (Rule 3).
  4. Keepa API (deep pull, final 3 only) for variations, A+/video, all-time
     seller feedback, fees, bullets, images.
  5. Step 10C: Claude scores the judgement elements (vision on listing images);
     the deterministic elements are scored in Python -> 184-point rubric.
  6. Deliverables + Step 11 Decision Gate (A workable / B re-pick / C reject).

Pure-deterministic data ops; all judgement via Claude (Rule 0 — no heuristic
fallback). Keepa/Anthropic keys from the git-ignored .env.
"""

import sys
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd

_ROOT = Path(__file__).resolve().parent.parent.parent  # project root (above research/)
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

# Windows consoles default to cp1252; the UI uses box-drawing glyphs. Make
# stdout UTF-8 (replace on failure) so output never crashes the pipeline.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:  # noqa: BLE001
    pass

from research.phase3.js_reader import read_js_products
from research.phase3.xray_product_reader import read_xray_products, aggregate_sponsored
from research.phase3 import keepa_client as kc
from research.phase3 import scorer as scoring
from research.phase3.claude_client import (
    Phase3ApiError, load_client, derive_product_profile,
    filter_same_product_type, score_listing, analyze_reviews,
)
from research.phase3.competitor_selector import build_candidate_records, select_top_competitors
from research.phase3.reviews_flyby import fetch_reviews, reviews_text


def _banner() -> None:
    print("\n" + "=" * 65)
    print("  PL Research Tool  |  Phase 3: Competitor Analysis  |  v9.6")
    print("=" * 65)


def _section(title: str) -> None:
    pad = max(0, 59 - len(title))
    print(f"\n--- {title} {'─' * pad}")


def _get_path(prompt: str, label: str) -> str:
    _section(label)
    print(f"  {prompt}\n")
    while True:
        p = input("  Path: ").strip().strip('"').strip("'")
        if Path(p).exists():
            return p
        print(f"  Not found: {p}")


def _get_js_df() -> Optional[pd.DataFrame]:
    """Prompt for the JS sheet: a file path, or `kw:<keyword>` to auto-generate
    it from Amazon page 1 + Keepa (no Jungle Scout subscription needed)."""
    _section("JUNGLE SCOUT FILE")
    print("  Paste the JS keyword-search CSV path, or type kw:<keyword> to")
    print("  auto-generate the sheet (Amazon page-1 scrape + Keepa, ~2 tokens/product).\n")
    while True:
        raw = input("  Path or kw:<keyword>: ").strip().strip('"').strip("'")
        if raw.lower().startswith("kw:"):
            kw = raw[3:].strip()
            if not kw:
                print("  Empty keyword — try again.")
                continue
            try:
                from research.phase3.js_sheet import build_js_sheet
                print(f"  Generating the sheet for '{kw}'...")
                res = build_js_sheet(keyword=kw)
                print(f"  Generated: {res['out_path']} ({res['count']} products, "
                      f"Keepa tokens left: {res['tokens_left']})")
                return read_js_products(res["out_path"])[0]
            except Exception as e:  # noqa: BLE001 — fall back to manual file
                print(f"  Auto-generation failed: {e}")
                continue
        if Path(raw).exists():
            return read_js_products(raw)[0]
        print(f"  Not found: {raw}")


def _collect_xray_files() -> Dict[str, pd.DataFrame]:
    """Prompt for a folder (globs Xray_products_*.xlsx) or comma-separated files."""
    _section("XRAY PRODUCT EXPORTS (one per launch keyword)")
    print("  Enter a folder containing the Xray product exports, or a comma-")
    print("  separated list of file paths.\n")
    raw = input("  Path(s): ").strip().strip('"').strip("'")
    paths: List[Path] = []
    p = Path(raw)
    if p.is_dir():
        paths = sorted(p.glob("Xray_products_*.xlsx")) or sorted(p.glob("*.xlsx"))
    else:
        paths = [Path(x.strip().strip('"').strip("'")) for x in raw.split(",") if x.strip()]
    per_kw: Dict[str, pd.DataFrame] = {}
    for fp in paths:
        if not fp.exists():
            print(f"  Skipping missing file: {fp}")
            continue
        kw = fp.stem.replace("Xray_products_Keyword_", "").replace("_", " ").strip()
        df, _ = read_xray_products(str(fp))
        per_kw[kw] = df
    return per_kw


def _lookup_title(asin: str, js_df: pd.DataFrame, per_kw: Dict[str, pd.DataFrame],
                  blackbox_df: Optional[pd.DataFrame]) -> Dict[str, str]:
    """Find a title + category for the target ASIN across the available sources."""
    for df in [blackbox_df, js_df, *per_kw.values()]:
        if df is None or "asin" not in df.columns:
            continue
        m = df[df["asin"].astype(str).str.upper() == asin]
        if not m.empty:
            row = m.iloc[0]
            return {
                "title": str(row.get("title") or "").strip(),
                "category": str(row.get("category") or "").strip(),
            }
    return {}


def _print_selection_table(top: pd.DataFrame, main_asin: str) -> None:
    """Step 8 / Rule 3 ranked table."""
    print(f"\n  {'Rank':<5}{'ASIN':<13}{'Brand':<18}{'Sales/mo':>9}  {'Rating':>6}  "
          f"{'FBA/FBM':<8}{'BSR':>7}  Main")
    print(f"  {'─'*78}")
    for i, r in enumerate(top.itertuples(), start=1):
        mark = "  <== MAIN" if r.asin == main_asin else ""
        sales = f"{int(r.monthly_sales):,}" if pd.notna(r.monthly_sales) else "N/A"
        rating = f"{r.rating:.1f}" if pd.notna(r.rating) else "N/A"
        bsr = f"{int(r.bsr):,}" if pd.notna(r.bsr) else "N/A"
        print(f"  {i:<5}{str(r.asin):<13}{str(r.brand or '')[:17]:<18}{sales:>9}  "
              f"{rating:>6}  {str(r.fulfillment):<8}{bsr:>7}{mark}")


def _print_scores(scored: List[Dict], records_by_asin: Dict[str, Dict]) -> None:
    for s in scored:
        rec = records_by_asin.get(s["asin"], {})
        print(f"\n  ── {rec.get('brand','')} ({s['asin']}) — "
              f"Score {s['total']}/{s['max']}  [{s['band']}]")
        for key, label, mx, _ in scoring.ELEMENTS:
            e = s["elements"][key]
            sc = e["score"] if e["score"] is not None else "—(Unknown)"
            print(f"     {label:<34} {str(sc):>10} / {mx}")
        if s["unknown_elements"]:
            print(f"     Unknown (need user input): {', '.join(s['unknown_elements'])}")


def run_phase3(
    blackbox_df: Optional[pd.DataFrame] = None,
    target_asin: Optional[str] = None,
    locked_keywords: Optional[list] = None,
    *,
    js_autogen: Optional[Dict] = None,
    dev_inputs: Optional[Dict] = None,
) -> Optional[Dict]:
    """Run Phase 3. Returns a result dict (selection + scores) or None.

    *js_autogen*: {'keyword': str, 'job': (thread, holder)} — the JS-sheet
    generation Phase 2 started when Claude extracted the main search term from
    the target's title. Phase 3 just joins it (the work is usually done by now).

    *dev_inputs* (testing only): {'js': path, 'xray_dir': path, 'target_asin':...,
    'deep_limit': int} to bypass interactive prompts.
    """
    _banner()

    # ── Target ASIN ──────────────────────────────────────────────────────────
    asin = (target_asin or (dev_inputs or {}).get("target_asin") or "").strip().upper()
    if not asin and not dev_inputs:
        _section("TARGET ASIN")
        asin = input("  Target ASIN (from Phase 1/2): ").strip().upper()
    if not asin:
        print("\n  No target ASIN — exiting Phase 3.")
        return None

    # ── Step 7 inputs ────────────────────────────────────────────────────────
    # The JS-style sheet is auto-generated (Amazon page-1 scrape -> Keepa ->
    # CSV) from the MAIN SEARCH TERM — extracted by Claude from the target's
    # title in Phase 2, which started the generation in the background back
    # then (js_autogen). Phase 3 joins that job here; by the time the user has
    # entered the Xray files it is normally already finished. Manual JS upload
    # remains the fallback.
    if dev_inputs:
        js_df, _ = read_js_products(dev_inputs["js"])
        import glob
        per_kw = {}
        for fp in sorted(glob.glob(str(Path(dev_inputs["xray_dir"]) / "Xray_products_*.xlsx"))):
            kw = Path(fp).stem.replace("Xray_products_Keyword_", "").replace("_", " ")
            per_kw[kw], _ = read_xray_products(fp)
    else:
        if js_autogen:
            _section("JUNGLE SCOUT SHEET (auto-generated from main search term)")
            print(f"  '{js_autogen['keyword']}' — started in Phase 2 when Claude")
            print("  extracted the search term from the target's title.")

        per_kw = _collect_xray_files()

        js_df = None
        if js_autogen:
            thread, holder = js_autogen["job"]
            if thread.is_alive():
                print("\n  Waiting for the JS-sheet generation to finish...")
            thread.join()
            if "res" in holder:
                res = holder["res"]
                print(f"  JS sheet ready: {res['out_path']} ({res['count']} products, "
                      f"Keepa tokens left: {res['tokens_left']})"
                      + (f"\n  Note: {res['note']}" if res.get("note") else ""))
                js_df, _ = read_js_products(res["out_path"])
            else:
                print(f"  JS-sheet generation failed: {holder.get('err')}")
                print("  Falling back to a manual Jungle Scout file.")
        if js_df is None:
            js_df = _get_js_df()
        if js_df is None:
            print("\n  No Jungle Scout data — cannot proceed.")
            return None

    if not per_kw:
        print("\n  No Xray product exports loaded — cannot proceed.")
        return None

    # ── Claude client ────────────────────────────────────────────────────────
    try:
        client = load_client()
    except Phase3ApiError as e:
        print(f"\n  {e}")
        return None

    # ── Target profile (relevancy anchor) ────────────────────────────────────
    info = _lookup_title(asin, js_df, per_kw, blackbox_df)
    if not info.get("title"):
        print(f"\n  Could not find a title for {asin} in the provided files.")
        return None
    _section("TARGET PRODUCT PROFILE (Claude)")
    print(f"  ASIN : {asin}\n  Title: {info['title'][:90]}")
    try:
        profile = derive_product_profile(client, title=info["title"], category=info.get("category", ""))
    except Phase3ApiError as e:
        print(f"\n  {e}")
        return None
    print(f"  Type : {profile.product_type}  |  NOT: {', '.join(profile.not_this) or '(none)'}")

    # ── Step 8: pool -> same-type -> top 3 ───────────────────────────────────
    _section("STEP 8 — COMPETITOR SELECTION")
    agg = aggregate_sponsored(per_kw)
    pool = kc.build_asin_pool(js_df, list(per_kw.values()))
    records = build_candidate_records(js_df, per_kw, agg, pool)
    try:
        top, same_type, report = select_top_competitors(client, profile, records, n=3)
    except (Phase3ApiError, ValueError) as e:
        print(f"\n  {e}")
        return None
    for line in report:
        print(line)
    if top.empty:
        print("\n  No same-product-type competitors found.")
        return None

    main_asin = top.iloc[0]["asin"]  # highest-sales same-type; main launch variant cross-check is Phase 6
    _print_selection_table(top, main_asin)

    # ── Step 10A: Keepa deep pull (final 3 only) ─────────────────────────────
    _section("KEEPA DEEP PULL (top 3)")
    top_asins = list(top["asin"])
    records_by_asin: Dict[str, Dict] = {r["asin"]: dict(r) for _, r in top.iterrows()}
    try:
        api = kc.load_keepa()
        deep_limit = (dev_inputs or {}).get("deep_limit")
        deep = (dev_inputs or {}).get("deep", True)
        products = kc.query_products(api, top_asins, deep=deep, limit=deep_limit)
        print(f"  Keepa returned {len(products)} of {len(top_asins)} (tokens left: {api.tokens_left})")
        for a in top_asins:
            if a in products:
                f = kc.extract_keepa_fields(products[a])
                if f.get("buybox_seller_id"):
                    f["seller_feedback"] = kc.get_seller_feedback(api, f["buybox_seller_id"])
                records_by_asin[a].update({k: v for k, v in f.items() if v is not None or k in
                                           ("a_plus", "has_video", "seller_feedback")})
    except kc.Phase3KeepaError as e:
        print(f"  Keepa unavailable: {e}\n  Variations/A+/seller-feedback will be Unknown.")

    # ── Step 10C: Claude listing scoring + reviews, then rubric ──────────────
    _section("STEP 10C — SCORING (Claude + rubric)")
    pool_stats = scoring.pool_statistics(same_type)
    scored: List[Dict] = []
    for a in top_asins:
        rec = records_by_asin[a]
        try:
            listing = score_listing(client, {
                "title": rec.get("title"), "bullets": rec.get("bullets"),
                "price": rec.get("price"), "a_plus": rec.get("a_plus"),
                "has_video": rec.get("has_video"), "images_count": rec.get("images_count"),
                "market_price_context": f"prices: {sorted([p for p in top['price'] if pd.notna(p)])}",
            }, image_urls=rec.get("image_urls"))
        except Phase3ApiError as e:
            print(f"  Listing scoring failed for {a}: {e}")
            listing = None
        # Reviews for this ASIN via the FlyBy/RapidAPI Real-Time Amazon Data
        # API (RAPIDAPI_KEY): ALL 1-2 star reviews found + max five 4-5 star;
        # 3 star excluded. The rating-distribution histogram supplies the
        # honest negative share (the full review list is login-gated on
        # Amazon — no source can reach it). dev_inputs["reviews_by_asin"][asin]
        # overrides with pre-supplied text.
        rev_text = (dev_inputs or {}).get("reviews_by_asin", {}).get(a, "")
        if not rev_text:
            rev = fetch_reviews(a)  # never raises; reports via "error"
            if rev.get("error"):
                print(f"  Reviews {a} (FlyBy) failed: {rev['error']}")
            else:
                print(f"  Reviews {a} (FlyBy): weaknesses={rev.get('weaknesses_count')} "
                      f"strengths={len(rev.get('strengths') or [])} "
                      f"(total ratings {rev.get('total_ratings')})"
                      + (f" [{rev['note']}]" if rev.get("note") else ""))
            rec["review_counts"] = {
                "total_reviews": rev.get("total_ratings"),
                "distribution": rev.get("distribution"),
                "weaknesses_count": rev.get("weaknesses_count"),
            }
            rev_text = reviews_text(rev)
        rec["review_analysis"] = analyze_reviews(client, rec.get("brand") or a, rev_text)
        s = scoring.score_competitor(rec, [dict(r) for _, r in top.iterrows()], pool_stats, listing)
        scored.append(s)

    _print_scores(scored, records_by_asin)

    # ── Step 10E: write the Competitor Analysis workbook ─────────────────────
    _section("WRITING DELIVERABLE (Competitor Analysis workbook)")
    template = (dev_inputs or {}).get(
        "template", "test_file/competitor_analysis/Compatitor Analysis.xlsx")
    scored_by_asin = {s["asin"]: s for s in scored}
    reviews_by_asin = {a: records_by_asin[a].get("review_analysis") for a in top_asins}
    kw_list = [getattr(k, "keyword", k) for k in (locked_keywords or [])]
    try:
        from research.phase3.output import write_competitor_workbook
        out_path = write_competitor_workbook(
            template, [records_by_asin[a] for a in top_asins],
            scored_by_asin, reviews_by_asin, kw_list,
            target_title=info.get("title"),
        )
        print(f"  Saved: {out_path}")
    except Exception as e:  # noqa: BLE001
        print(f"  Could not write workbook: {e}")

    # ── Step 11: Decision Gate ───────────────────────────────────────────────
    _section("STEP 11 — DECISION GATE")
    print("  (A) Workable  -> proceed to Phase 4")
    print("  (B) Change competitors -> re-pick")
    print("  (C) Not workable -> back to Phase 1")
    print("\n  Excel + PDF deliverables: pending sheet-layout confirmation for the")
    print("  4 new rubric rows (Review Velocity / Variations / Sales Strength /")
    print("  Enhanced Content). Scores above are final and ready to write.")

    return {
        "target_asin": asin, "profile": profile, "top": top,
        "same_type": same_type, "scored": scored, "records": records_by_asin,
    }


if __name__ == "__main__":
    run_phase3()
