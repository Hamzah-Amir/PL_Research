"""
Standalone Phase 4 (Critical Sheet) test — feeds a MANUALLY-PROVIDED Jungle
Scout sheet straight into Phase 4, bypassing the Amazon page-1 scraper (which
Amazon currently blocks). Nothing in the real pipeline (run_phase4 / views) is
changed — we just hand run_phase4 a `js_file` from disk instead of an
auto-generated one.

    python test_phase4.py [TARGET_ASIN] ["keyword"] [--js PATH] [--limit N]

Defaults:
    TARGET_ASIN  B0CBSHWRGJ   (present in the sample JS sheet)
    keyword      "sample keyword"   (only written into the sheet's B7 cell)
    --js         test_file/js-keyword-search.csv
    --limit      5   (parents to classify — keep small; each costs Keepa tokens
                      + a Claude vision call)

To test with your own JS export, drop it in test_file/ and pass --js <path>
(CSV or XLSX in Jungle Scout keyword-search format).

Stage 1 (propose): Keepa-pulls the parents + variations from the sheet and
prints Claude's proposed vocabulary + Design attributes.
Stage 2 (fill): auto-approves that proposal (edit `approved` below to change it)
and writes output/Phase4_CriticalSheet_<asin>_<date>.xlsx (filled into the PES
template's 'Critical sheet' tab).
"""
import argparse
import os
import sys

import django

os.environ.setdefault("DJANGO_SETTINGS_MODULE", "plresearch.settings")
django.setup()

from research.services.phase4 import run_phase4  # noqa: E402

DEFAULT_JS = "test_file/js-keyword-search.csv"
DEFAULT_ASIN = "B0CBSHWRGJ"


def main():
    ap = argparse.ArgumentParser(description="Test Phase 4 with a manual JS sheet.")
    ap.add_argument("asin", nargs="?", default=DEFAULT_ASIN, help="target ASIN")
    ap.add_argument("keyword", nargs="?", default="sample keyword",
                    help="keyword label written to the sheet (no scraping)")
    ap.add_argument("--js", default=DEFAULT_JS, help="path to the JS sheet (CSV/XLSX)")
    ap.add_argument("--limit", type=int, default=5, help="parents to classify")
    ap.add_argument("--max-asins", type=int, default=3,
                    help="TOTAL ASIN cap for testing (parents+variations); 0 = no cap")
    ap.add_argument("--cerebro", default="test_file/cerebro.xlsx", help="Cerebro export -> KWs tabs")
    ap.add_argument("--xray", default="test_file/xray-products",
                    help="folder of Helium 10 Xray exports -> H10 Basic Data tab")
    args = ap.parse_args()
    max_asins = args.max_asins or None
    xray = args.xray if args.xray and os.path.isdir(args.xray) else None
    cerebro = args.cerebro if args.cerebro and os.path.exists(args.cerebro) else None

    if not os.path.exists(args.js):
        print(f"JS sheet not found: {args.js}")
        sys.exit(1)
    asin = args.asin.strip().upper()
    print(f"Manual JS sheet: {args.js}\nTarget ASIN: {asin} | keyword: {args.keyword!r} | "
          f"parent_limit: {args.limit}\n")

    # ── Stage 1 — propose (manual JS sheet, no scraping) ─────────────────────
    r1 = run_phase4(asin, js_file=args.js, keyword=args.keyword, parent_limit=args.limit,
                    max_asins=max_asins)
    if r1.get("error"):
        print("PROPOSE ERROR:", r1["error"])
        sys.exit(1)
    v = r1["vocabulary"]
    print(f"Parents: {r1['parent_count']} | variation children: {r1['variation_count']}")
    print("\nProposed vocabulary:")
    for k, val in v.items():
        print(f"  {k}: {val}")
    for line in r1.get("log", []):
        print("  [log]", line)

    # ── Stage 2 — fill (auto-approve the proposal; edit here to change it) ────
    approved = {
        "material_values": v["material_values"],
        "size_guidance": v["size_guidance"],
        "color_values": v["color_values"],
        "packaging_values": v["packaging_values"],
        "special_feature_guidance": v["special_feature_guidance"],
        "design_attributes": v["design_attributes"],
    }
    r2 = run_phase4(asin, js_file=args.js, keyword=args.keyword, approved=approved,
                    parent_limit=args.limit, max_asins=max_asins, xray_files=xray,
                    cerebro_file=cerebro, launch_keywords=["tent pegs","metal tent pegs","ground stakes","net pegs","awning pegs","camping pegs metal"])
    if r2.get("error"):
        print("\nFILL ERROR:", r2["error"])
        sys.exit(1)
    print(f"\nBuilt: {r2['row_count']} rows "
          f"({r2['parent_count']} parents + {r2['variation_count']} variations)")
    print("Design legend:", r2.get("legend"))
    print("Workbook:", r2.get("workbook_path"))
    for line in r2.get("log", []):
        print("  [log]", line)


if __name__ == "__main__":
    main()
