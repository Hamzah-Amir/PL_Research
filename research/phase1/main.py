"""
PL Research Tool — Phase 1: Product Hunting (DB-backed)
Entry point.

Phase 1 is now a stored-dataset model:
  1. Helium 10 Black Box exports are ingested into PostgreSQL (scored + segmented;
     history preserved). Weekly full refresh.
  2. The user gives a budget (+ optional category / seasonal / variants choice);
     the tool queries a pool of the highest-opportunity products from the DB
     (no live API) and exports the top 10 to Excel.
  3. Not satisfied? It returns the next 10 from the same pool (no re-query).
  4. The user picks one ASIN, which hands off to Phase 2 (keyword verification).

Run:
    python run.py                 # connect DB, optional refresh, then shortlist
    python -m phase1.main ingest  # ingest/refresh only (prompts for files)
"""

import sys
from pathlib import Path
from typing import Dict, List, Optional

_ROOT = Path(__file__).resolve().parent.parent.parent  # project root (above research/)
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:  # noqa: BLE001
    pass

from research.phase1.db import get_connection, init_db, Phase1DbError
from research.phase1.ingest import ingest_files
from research.phase1.query import (
    fetch_pool, filter_funnel, max_revenue_for_budget, revenue_band,
    review_band_for_budget,
)
from research.phase1.output import export_shortlist, print_shortlist_table

BATCH_SIZE = 10
DEFAULT_POOL = 100


def _banner() -> None:
    print("\n" + "=" * 65)
    print("  PL Research Tool  |  Phase 1: Product Hunting  |  v9.6 (DB)")
    print("=" * 65)


def _section(title: str) -> None:
    print(f"\n--- {title} {'─' * max(0, 59 - len(title))}")


def _collect_paths(raw: str) -> List[str]:
    p = Path(raw.strip().strip('"').strip("'"))
    if p.is_dir():
        return [str(x) for x in sorted(p.glob("*.xlsx")) if "blackbox" in x.name.lower()] \
            or [str(x) for x in sorted(p.glob("*.xlsx"))]
    return [s.strip().strip('"').strip("'") for s in raw.split(",") if s.strip()]


def _product_count(conn) -> int:
    with conn.cursor() as cur:
        cur.execute("SELECT COUNT(*) FROM products")
        return cur.fetchone()[0]


def _run_ingest(conn) -> None:
    _section("INGEST / REFRESH")
    print("  Enter a folder of Black Box exports, or comma-separated file paths.\n")
    raw = input("  Black Box file(s): ").strip()
    paths = [p for p in _collect_paths(raw) if Path(p).exists()]
    if not paths:
        print("  No valid files found — skipping ingest.")
        return
    print(f"  Ingesting {len(paths)} file(s)...")
    ingest_files(conn, paths)


def _print_funnel(funnel: List[Dict]) -> None:
    """Console-only filter funnel: each filter's exclusions + remaining."""
    _section("FILTER FUNNEL (how many each filter removed)")
    print(f"  {'Filter':<34}{'Removed':>9}{'Remaining':>11}")
    print(f"  {'─' * 54}")
    for step in funnel:
        removed = "—" if step["removed"] == 0 and step["filter"].startswith("All") else f"{step['removed']:,}"
        print(f"  {step['filter'][:33]:<34}{removed:>9}{step['remaining']:>11,}")
    total_removed = sum(s["removed"] for s in funnel)
    print(f"  {'─' * 54}")
    print(f"  {'Total removed / qualifying':<34}{total_removed:>9,}{funnel[-1]['remaining']:>11,}")


def _parse_budget(raw: str) -> Optional[float]:
    s = raw.replace("£", "").replace(",", "").replace("$", "").strip()
    try:
        return float(s) if s else None
    except ValueError:
        return None


def _ask_query_params() -> Dict:
    _section("FIND PRODUCTS")
    budget = _parse_budget(input("  Inventory budget in £ (e.g. 4000, or Enter to skip): "))
    max_revenue = max_revenue_for_budget(budget)
    if max_revenue is not None:
        lo, hi = revenue_band(max_revenue)
        print(f"  -> revenue band: £{lo:,.0f} – £{hi:,.0f}/mo  (budget × 1.52, −£1,500 width)")
    rmin, rmax = review_band_for_budget(budget)
    review_txt = f"{rmin}+" if rmax is None else f"{rmin}-{rmax}"
    print("  -> built-in filters: rating>3.8, age>=6mo, sales>=50, 90d-trend>-25%,")
    print(f"                       reviews {review_txt} (by budget), exclude Amazon/compatibility/global-brand")
    category = input("  Category filter (Enter to skip): ").strip() or None
    seasonal = input("  Include seasonal products? (Y/n): ").strip().lower() not in ("n", "no")
    return {"budget": budget, "max_revenue": max_revenue, "category": category,
            "include_seasonal": seasonal}


def _shortlist_loop(conn, params: Dict) -> Optional[Dict]:
    """Show the pool 10 at a time; return the chosen product dict or None."""
    # Filter funnel (console only) — how many products each filter excluded.
    funnel = filter_funnel(
        conn, budget=params["budget"], max_revenue=params.get("max_revenue"),
        category=params["category"], include_seasonal=params["include_seasonal"],
    )
    _print_funnel(funnel)

    pool = fetch_pool(
        conn, budget=params["budget"], max_revenue=params.get("max_revenue"),
        category=params["category"], include_seasonal=params["include_seasonal"],
        pool_size=DEFAULT_POOL,
    )
    if not pool:
        print("\n  No products matched those filters. Try a different budget/category,")
        print("  or ingest more data.")
        return None

    print(f"\n  Pool: {len(pool)} products (best opportunity first).")
    batch_num, idx = 1, 0
    while idx < len(pool):
        batch = pool[idx: idx + BATCH_SIZE]
        print_shortlist_table(batch, batch_num)
        path = export_shortlist(
            batch, batch_num, budget=params["budget"], category=params["category"],
            include_seasonal=params["include_seasonal"], pool_total=len(pool),
        )
        print(f"  Excel: {path}")

        idx += BATCH_SIZE
        more = idx < len(pool)
        print("\n  Pick a product or see more:")
        print("    <ASIN>  = select that product and continue to Phase 2")
        if more:
            print("    n       = see the next 10 from the pool")
        print("    q       = quit")
        choice = input("  Your choice: ").strip()
        cl = choice.lower()
        if cl in ("q", "quit"):
            return None
        if cl in ("n", "next", "") and more:
            batch_num += 1
            continue
        # treat as ASIN
        asin = choice.upper()
        match = next((p for p in pool if str(p.get("asin", "")).upper() == asin), None)
        if match:
            return match
        print(f"  '{choice}' is not an ASIN in this pool — try again.")
        # re-show same batch
        idx -= BATCH_SIZE
    print("\n  End of pool — no more products.")
    return None


def run_phase1() -> None:
    _banner()
    try:
        conn = get_connection()
        init_db(conn)
    except Phase1DbError as e:
        print(f"\n  {e}")
        return

    try:
        count = _product_count(conn)
        print(f"\n  Products currently stored: {count:,}")
        if count == 0:
            print("  Database is empty — ingest a Black Box export to begin.")
            _run_ingest(conn)
        else:
            if input("  Refresh data from new Black Box files? (y/N): ").strip().lower() in ("y", "yes"):
                _run_ingest(conn)

        if _product_count(conn) == 0:
            print("\n  No products available — exiting.")
            return

        params = _ask_query_params()
        selected = _shortlist_loop(conn, params)
    finally:
        pass  # keep conn open until after handoff lookup below

    if not selected:
        print("\n  No product selected — Phase 1 ends here.")
        conn.close()
        return

    print("\n" + "=" * 65)
    print(f"  Selected target ASIN: {selected['asin']}  ({selected.get('brand') or ''})")
    print(f"  {str(selected.get('title') or '')[:80]}")
    print("=" * 65)
    target_product = {
        "asin": selected["asin"],
        "title": selected.get("title") or "",
        "category": selected.get("category") or "",
        "subcategory": selected.get("subcategory") or "",
    }
    conn.close()

    # ── Hand off to Phase 2 (keyword verification) ────────────────────────────
    print("\n  Moving on to Phase 2 (keyword verification)...")
    from research.phase2.main import run_phase2
    p2 = run_phase2(target_product=target_product)

    # ── Hand off to Phase 3 (competitor analysis) ─────────────────────────────
    if p2 and p2.get("keywords"):
        print("\n  Moving on to Phase 3 (competitor analysis)...")
        from research.phase3.main import run_phase3
        run_phase3(target_asin=p2.get("target_asin"), locked_keywords=p2.get("keywords"),
                   js_autogen=p2.get("js_autogen"))


def _main() -> None:
    if len(sys.argv) > 1 and sys.argv[1].lower() == "ingest":
        _banner()
        try:
            conn = get_connection()
            init_db(conn)
            _run_ingest(conn)
            print(f"\n  Stored products: {_product_count(conn):,}")
            conn.close()
        except Phase1DbError as e:
            print(f"\n  {e}")
        return
    run_phase1()


if __name__ == "__main__":
    _main()
