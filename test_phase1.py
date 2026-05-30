"""
Non-interactive test of the full Phase 1 flow using the real sample file.

Budget : GBP 5,000  =>  min revenue GBP 7,600/month
All filters applied via build_filter_config (same as production).
All batches exported automatically (no user prompt in test mode).
"""

import sys
import io
import warnings
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.path.insert(0, str(Path(__file__).parent))
warnings.filterwarnings("ignore")

from phase1.file_reader import read_blackbox_file
from phase1.filters import build_filter_config, apply_filters, print_active_filters
from phase1.scorer import select_top_products
from phase1.output import export_to_excel, print_product_table

# ── Settings ──────────────────────────────────────────────────────────────────
TEST_FILE   = "test_file/GB_AMAZON_blackBoxProducts_1_2026-05-28.xlsx"
BUDGET      = 5_000.0
MAX_REVENUE = BUDGET * 1.52      # revenue cap — products above this are excluded
CATEGORY    = ""          # leave blank to auto-detect from file


def divider(title=""):
    if title:
        pad = max(0, 60 - len(title))
        print(f"\n--- {title} {'─'*pad}")
    else:
        print("─" * 65)


def run_test():
    print("=" * 65)
    print("  Phase 1 Test Run — Real Sample File")
    print("=" * 65)
    print(f"  File   : {TEST_FILE}")
    print(f"  Budget : GBP {BUDGET:,.0f}  =>  min revenue GBP {MAX_REVENUE:,.0f}/month")

    # ── Read file ─────────────────────────────────────────────────────────────
    divider("READ FILE")
    df, detected_category = read_blackbox_file(TEST_FILE)
    category = CATEGORY if CATEGORY else detected_category
    print(f"\n  Category in use : {category}")
    print(f"  Total products  : {len(df)}")

    # ── Build filters (built-in, no prompts) ──────────────────────────────────
    divider("FILTERS APPLIED")
    config = build_filter_config(category, MAX_REVENUE)
    print_active_filters(config)

    # ── Apply filters ─────────────────────────────────────────────────────────
    divider("FILTER RESULTS")
    filtered_df, report = apply_filters(df, config)
    print()
    for line in report:
        print(line)

    if filtered_df.empty:
        print("\n  No products passed all filters.")
        return

    # ── Score & rank ──────────────────────────────────────────────────────────
    divider("SCORING")
    ranked = select_top_products(filtered_df, top_n=100)
    pool_size = len(ranked)
    scores = ranked["composite_score"]
    print(f"\n  Pool (max 100)  : {pool_size} products")
    print(f"  Score range     : {scores.min():.1f} – {scores.max():.1f}")
    print(f"  >= 70 (green)   : {(scores >= 70).sum()}")
    print(f"  50-69 (amber)   : {((scores >= 50) & (scores < 70)).sum()}")
    print(f"  < 50  (red)     : {(scores < 50).sum()}")

    # ── Export all batches ────────────────────────────────────────────────────
    divider("BATCHES")
    exported = []
    for b in range((pool_size + 9) // 10):
        start = b * 10
        batch = ranked.iloc[start: start + 10].copy()
        print_product_table(batch, b + 1)
        path = export_to_excel(
            batch        = batch,
            batch_num    = b + 1,
            category     = category,
            budget       = BUDGET,
            max_revenue  = MAX_REVENUE,
            filter_config= config,
        )
        exported.append(path)
        print(f"\n  Saved ({len(batch)} products): {path}")

    # ── Summary ───────────────────────────────────────────────────────────────
    divider("SUMMARY")
    print(f"\n  Products in file         : {len(df)}")
    print(f"  After all filters        : {len(filtered_df)}")
    print(f"  Ranked pool              : {pool_size}")
    print(f"  Excel files created      : {len(exported)}")
    for p in exported:
        print(f"    {p}")
    print()
    print("=" * 65)


if __name__ == "__main__":
    run_test()
