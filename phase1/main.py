"""
PL Research Tool — Phase 1: Product Hunting
Entry point.

User is asked:
  1. Budget (GBP) — required
  2. Seasonal products? (Y/N) — keep & flag seasonal items, or exclude them
  3. Product category — optional (auto-detected from file if left blank)

Revenue and monthly sales are both read from ASIN-level fields
(capped at budget × 1.52, no minimum). There is no variation filter —
products are included regardless of variations.

All other filters run silently on the backend with hardcoded values (including
the F5 minimum unit price of GBP 5). Monthly sales are collected from
ASIN-level fields for scoring and output — no sales floor is applied.
All qualifying products are scored and ranked, then split by variation count
(the 'Variation Count' column: 0 / missing = no variants). Each batch produces
TWO Excel files — top 10 WITH variants and top 10 WITHOUT variants. If more
remain in either pool the user is asked whether to see the next 10 of each,
and so on until satisfied or both pools are exhausted.
"""

import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from phase1.file_reader import read_blackbox_file
from phase1.filters import build_filter_config, apply_filters, print_active_filters
from phase1.output import export_to_excel, print_product_table
from phase1.scorer import score_all_products


# ──────────────────────────────────────────────────────────────────────────────
# Input helpers
# ──────────────────────────────────────────────────────────────────────────────

def _banner() -> None:
    print()
    print("=" * 65)
    print("  PL Research Tool  |  Phase 1: Product Hunting  |  v9.6")
    print("=" * 65)


def _section(title: str) -> None:
    pad = max(0, 59 - len(title))
    print(f"\n--- {title} {'─' * pad}")


def _get_budget() -> float:
    _section("BUDGET")
    print("  Your budget sets the minimum monthly revenue a product must earn.")
    print("  Formula: Revenue Cap = Budget x 1.52")
    print("  Products earning MORE than the cap are too competitive — excluded.")
    print()
    while True:
        raw = input("  Enter your budget (GBP): ").strip().replace(",", "").replace("£", "")
        try:
            budget = float(raw)
            if budget <= 0:
                print("  Budget must be greater than zero. Try again.")
                continue
            min_rev = budget * 1.52
            print(f"\n  Budget              : GBP {budget:,.2f}")
            print(f"  Revenue cap (x1.52) : GBP {min_rev:,.2f} / month  (max allowed)")
            return budget
        except ValueError:
            print("  Invalid number. Enter a value like 5000.")


def _get_seasonal_choice() -> bool:
    """Ask whether the user wants seasonal products. Returns True for yes."""
    _section("SEASONAL PRODUCTS")
    print("  Do you want seasonal products included in your results?")
    print()
    print("    N (default) : all seasonal products are removed.")
    print("    Y           : seasonal products are kept and flagged with a note")
    print("                  (e.g. summer items). Christmas / Q4-only products")
    print("                  are always removed either way.")
    print()
    ans = input("  Include seasonal products?  Y = yes,  Enter/N = no: ").strip().lower()
    want = ans in ("y", "yes")
    print(f"\n  Seasonal products : {'INCLUDED (flagged)' if want else 'EXCLUDED'}")
    return want


def _get_category() -> str:
    _section("PRODUCT CATEGORY  (optional)")
    print("  Enter your target product category or niche.")
    print("  Press Enter to skip — category will be detected from the file.")
    print()
    return input("  Category (or Enter to skip): ").strip()


def _get_file_path() -> str:
    _section("BLACK BOX FILE")
    print("  Paste the full path to your Helium 10 Black Box XLSX export.")
    print()
    while True:
        path = input("  File path: ").strip().strip('"').strip("'")
        if Path(path).exists():
            return path
        print(f"  File not found: {path}")
        print("  Check the path and try again.")


# ──────────────────────────────────────────────────────────────────────────────
# Main flow
# ──────────────────────────────────────────────────────────────────────────────

def run_phase1() -> None:
    _banner()

    # ── 1. Budget ─────────────────────────────────────────────────────────────
    budget = _get_budget()
    max_revenue = budget * 1.52

    # ── 2. Seasonal-products preference ───────────────────────────────────────
    want_seasonal = _get_seasonal_choice()

    # ── 3. Optional category ──────────────────────────────────────────────────
    user_category = _get_category()

    # ── 4. File path ──────────────────────────────────────────────────────────
    file_path = _get_file_path()

    # ── 5. Read file ──────────────────────────────────────────────────────────
    _section("READING FILE")
    df, detected_category = read_blackbox_file(file_path)

    if df.empty:
        print("\n  No products could be read from the file.")
        print("  Verify it is an H10 Black Box export and try again.")
        return

    # Prefer user-supplied category; fall back to auto-detected
    category = user_category if user_category else detected_category
    print(f"\n  Category in use : {category}")
    print(f"  Products loaded : {len(df)}")

    # ── Build & display filters ───────────────────────────────────────────────
    _section("FILTERS  (applied automatically)")
    print(f"  Current month   : {datetime.now().strftime('%B %Y')}  "
          f"(seasonality is judged relative to now)")
    config = build_filter_config(category, max_revenue, want_seasonal)
    print_active_filters(config)

    # ── 7. Apply all filters ──────────────────────────────────────────────────
    _section("APPLYING FILTERS")
    filtered_df, report = apply_filters(df, config)

    print()
    for line in report:
        print(line)

    if filtered_df.empty:
        print()
        print("  No products passed all filters.")
        print("  Suggestions:")
        print("    - Raise your budget                (raises the revenue cap)")
        print("    - Toggle the variation-products choice")
        print("    - Check the file covers a broad enough category")
        print("    (built-in filters: price >= GBP 5, rating > 3.8,")
        print("     listing age >= 6 mo, reviews <= 3,000)")
        return

    # ── Score every product, then split by variation count ───────────────────
    _section("SCORING & RANKING")
    scored_all = score_all_products(filtered_df)

    # Split on the 'variations' (Variation Count) column: > 0 has variants,
    # 0 / missing has none. Each side stays ranked by composite score.
    if "variations" in scored_all.columns:
        var_counts = pd.to_numeric(scored_all["variations"], errors="coerce").fillna(0)
        with_var = scored_all[var_counts > 0].reset_index(drop=True)
        without_var = scored_all[var_counts <= 0].reset_index(drop=True)
    else:
        # No variation column → treat every product as having no variants.
        with_var = scored_all.iloc[0:0].reset_index(drop=True)
        without_var = scored_all.reset_index(drop=True)

    total_with = len(with_var)
    total_without = len(without_var)
    print(f"  Products scored  : {len(scored_all)}")
    print(f"    With variants  : {total_with}")
    print(f"    Without variants: {total_without}")

    # Each batch produces TWO Excel files (with-variants + without-variants).
    # The user is asked after each batch whether to see the next 10 of each.
    _POOLS = [
        (with_var,    total_with,    "With Variants",    "WithVariants"),
        (without_var, total_without, "Without Variants", "NoVariants"),
    ]

    _BATCH_SIZE = 10
    batch_num = 1
    offset = 0

    while offset < total_with or offset < total_without:
        _section(f"BATCH {batch_num}")

        for pool, total, label, tag in _POOLS:
            batch = pool.iloc[offset:offset + _BATCH_SIZE].reset_index(drop=True)
            if batch.empty:
                continue

            print_product_table(batch, batch_num, label)
            excel_path = export_to_excel(
                batch        = batch,
                batch_num    = batch_num,
                category     = category,
                budget       = budget,
                max_revenue  = max_revenue,
                filter_config= config,
                variant_label= label,
                filename_tag = tag,
            )
            print(f"\n  Excel saved — {label} ({len(batch)} products): {excel_path}")

        offset += _BATCH_SIZE
        batch_num += 1

        rem_with = max(0, total_with - offset)
        rem_without = max(0, total_without - offset)
        if rem_with == 0 and rem_without == 0:
            print("\n  No more products remaining.")
            break

        print(
            f"\n  Remaining — with variants: {rem_with}, "
            f"without variants: {rem_without}."
        )
        ans = input(
            "  Would you like to see the next 10 of each?  Y = yes,  Enter/N = no (exit): "
        ).strip().lower()
        if ans not in ("y", "yes"):
            break

    # ── Done ──────────────────────────────────────────────────────────────────
    print()
    print("=" * 65)
    print("  Phase 1 complete.")
    print("=" * 65)

    # ── Hand off to Phase 2 (keyword verification) — automatic ────────────────
    # The Black Box data is already in memory (df); Phase 2 reuses it so the user
    # is not asked to re-upload the file. The transition is automatic: Phase 2
    # opens and asks for the target ASIN picked from the Excel output (press
    # Enter at that prompt to exit if Phase 2 is not wanted).
    print("\n  Pick one ASIN from the Excel output to take into Phase 2.")
    print("  Moving on to Phase 2 (keyword verification)...")
    from phase2.main import run_phase2
    run_phase2(blackbox_df=df)


if __name__ == "__main__":
    run_phase1()
