"""
PL Research Tool — Phase 2: Keyword Verification
Entry point.

Scope (current build): runs the pipeline up to and including the point where
Claude determines the launch keywords from the product's condensed title.

Flow:
  1. User types the target ASIN chosen at the end of Phase 1.
  2. Black Box product data: reused from Phase 1 in memory when chained from it
     (no re-upload); otherwise the XLSX path is prompted for. The ASIN is looked
     up to read its title + category.
  3. Claude builds a structured product profile (name, type, key attributes,
     use cases, what-it-is-NOT, brand) from the title + category — the relevancy
     anchor. Text only; the product image is not sent.
  4. User provides the H10 Xray keyword export for the target product.
  5. Deterministic prep: de-duplicate + drop SV < 100, rank by SV.
  6. Claude judges relevancy against the product profile and proposes the
     keywords (PDF Step 5), with a reason per pick.
  7. Approval/swap loop (PDF Step 6): the user approves the set or rejects
     specific keywords; rejected ones are excluded and Claude picks relevant
     replacements (approved keywords are kept) until the set is locked.
  8. The locked keywords are printed and returned (no Excel export — they are
     consumed directly by Phase 3's competitor analysis sheet).

All keyword judgement is performed purely by Claude — no deterministic
relevancy logic, no heuristic fallback. The API key is read from a git-ignored
.env in the project root.
"""

import re
import sys
from pathlib import Path
from typing import List, Optional

import pandas as pd

_ROOT = Path(__file__).resolve().parent.parent.parent  # project root (above research/)
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from research.phase1.file_reader import read_blackbox_file
from research.phase2.xray_reader import read_xray_keywords
from research.phase2.keyword_selector import prepare_candidates
from research.phase2.claude_client import (
    Phase2ApiError,
    load_client,
    derive_product_profile,
    select_keywords,
)


# ──────────────────────────────────────────────────────────────────────────────
# Input helpers
# ──────────────────────────────────────────────────────────────────────────────

def _banner() -> None:
    print()
    print("=" * 65)
    print("  PL Research Tool  |  Phase 2: Keyword Verification  |  v9.6")
    print("=" * 65)


def _section(title: str) -> None:
    pad = max(0, 59 - len(title))
    print(f"\n--- {title} {'─' * pad}")


def _get_target_asin() -> str:
    _section("TARGET ASIN")
    print("  Enter the single target ASIN you selected at the end of Phase 1.")
    print("  (press Enter on an empty line to cancel)")
    print()
    return input("  Target ASIN: ").strip().upper()


def _get_file_path(prompt: str, label: str) -> str:
    _section(label)
    print(f"  {prompt}")
    print()
    while True:
        path = input("  File path: ").strip().strip('"').strip("'")
        if Path(path).exists():
            return path
        print(f"  File not found: {path}")
        print("  Check the path and try again.")


def _lookup_product(df: pd.DataFrame, asin: str) -> dict:
    """Return {title, category, subcategory} for *asin* from a Black Box frame.

    Returns an empty dict if the ASIN (with a title) is not found.
    """
    if "asin" not in df.columns or "title" not in df.columns:
        return {}
    mask = df["asin"].astype(str).str.strip().str.upper() == asin
    matches = df[mask]
    if matches.empty:
        return {}
    row = matches.iloc[0]

    def _val(key: str) -> str:
        v = row.get(key)
        return str(v).strip() if v is not None and pd.notna(v) else ""

    title = _val("title")
    if not title:
        return {}
    return {
        "title": title,
        "category": _val("category"),
        "subcategory": _val("subcategory"),
    }


# ──────────────────────────────────────────────────────────────────────────────
# Step 6 — approval / swap loop
# ──────────────────────────────────────────────────────────────────────────────

def _print_keywords(keywords: list, header: str) -> None:
    """Print a numbered keyword table (SV, relevancy, reason)."""
    print(f"\n  {header}")
    print(f"  {'#':>2}  {'Keyword':<34}  {'SV':>7}  {'Relevancy':<9}  Reason")
    print(f"  {'─'*100}")
    for i, kw in enumerate(keywords, start=1):
        print(
            f"  {i:>2}.  {kw.keyword[:34]:<34}  {kw.search_volume:>7,}  "
            f"{kw.relevancy:<9}  {kw.reason}"
        )


def _parse_reject_indices(raw: str, n: int) -> List[int]:
    """Parse '2,5' / '2 5' into zero-based indices within [0, n)."""
    out: List[int] = []
    for tok in re.split(r"[,\s]+", raw):
        if tok.isdigit():
            i = int(tok) - 1
            if 0 <= i < n and i not in out:
                out.append(i)
    return out


def _run_approval_loop(client, profile, candidates, selection, target_count=6):
    """
    Step 6: let the user approve the set or swap out specific keywords.

    Approved keywords are kept; rejected ones are excluded and Claude is asked
    for relevant replacements (never re-suggesting a rejected or already-kept
    phrase). Loops until the user approves. Returns the locked list, or None if
    the user cancels.
    """
    current = list(selection.keywords)
    note = selection.note
    rejected: List[str] = []

    while True:
        _print_keywords(current, "PROPOSED KEYWORDS")
        if note:
            print(f"\n  Note: {note}")
        print()
        print("  Approve these, or swap some out:")
        print("    Enter              = approve all and lock")
        print("    numbers (e.g. 2,5) = reject those; Claude picks replacements")
        print("    q                  = cancel Phase 2")
        raw = input("  Your choice: ").strip().lower()

        if raw in ("q", "quit"):
            return None
        if raw in ("", "y", "yes", "approve"):
            return current

        idxs = _parse_reject_indices(raw, len(current))
        if not idxs:
            print("  Could not read any valid numbers — try again (e.g. 1,3).")
            continue

        rejected_now = [current[i].keyword for i in idxs]
        rejected.extend(rejected_now)
        keep = [kw for i, kw in enumerate(current) if i not in idxs]
        need = target_count - len(keep)

        print(f"\n  Swapping out: {', '.join(rejected_now)}")
        print(f"  Asking Claude for {need} replacement(s)...")
        try:
            repl = select_keywords(
                client, profile, candidates,
                target_count=need,
                exclude=rejected + [k.keyword for k in keep],
            )
        except Phase2ApiError as e:
            print(f"\n  {e}")
            print("  Keeping the current set; you can try again or approve.")
            continue

        current = keep + list(repl.keywords)[:need]
        note = repl.note
        if len(current) < target_count:
            print(f"\n  Only {len(current)} keyword(s) available after swaps "
                  f"(wanted {target_count}).")


# ──────────────────────────────────────────────────────────────────────────────
# Main flow
# ──────────────────────────────────────────────────────────────────────────────

def run_phase2(
    blackbox_df: Optional[pd.DataFrame] = None,
    target_product: Optional[dict] = None,
) -> Optional[dict]:
    """Run Phase 2.

    Three entry modes for obtaining the target product's title/category:
      - *target_product* (preferred): a dict ``{asin, title, category,
        subcategory}`` handed in by the DB-backed Phase 1 — no ASIN prompt, no
        Black Box file needed.
      - *blackbox_df*: an in-memory Black Box frame (legacy chaining) — prompts
        for the ASIN and looks it up.
      - neither (``python -m phase2.main``): prompts for the ASIN and the Black
        Box file path.

    Returns a dict ``{"target_asin", "keywords", "target_product"}`` consumed by
    Phase 3, or ``None`` if cancelled / unable to complete.
    """
    _banner()

    # ── 1-2. Target product (ASIN + title/category) ───────────────────────────
    if target_product and target_product.get("asin") and target_product.get("title"):
        asin = str(target_product["asin"]).strip().upper()
        product_row = {
            "title": target_product.get("title", ""),
            "category": target_product.get("category", ""),
            "subcategory": target_product.get("subcategory", ""),
        }
        _section("USING PHASE 1 TARGET PRODUCT")
        print(f"  Carried from Phase 1 — no ASIN entry or file needed.")
    else:
        asin = _get_target_asin()
        if not asin:
            print("\n  No ASIN entered — exiting Phase 2.")
            return
        if blackbox_df is not None:
            _section("USING PHASE 1 PRODUCT DATA")
            print(f"  Reusing {len(blackbox_df)} products from Phase 1 — no re-upload needed.")
            bb_df = blackbox_df
        else:
            bb_path = _get_file_path(
                "Paste the full path to your Phase 1 Helium 10 Black Box XLSX export.",
                "BLACK BOX FILE",
            )
            _section("READING BLACK BOX")
            bb_df, _ = read_blackbox_file(bb_path)

        product_row = _lookup_product(bb_df, asin)
        if not product_row:
            print(f"\n  Could not find ASIN {asin} (with a title) in the Black Box file.")
            print("  Verify the ASIN matches one in that export and try again.")
            return
    print(f"\n  Target ASIN  : {asin}")
    print(f"  Title        : {product_row['title']}")
    if product_row.get("category") or product_row.get("subcategory"):
        cat = " > ".join(c for c in (product_row.get("category"),
                                     product_row.get("subcategory")) if c)
        print(f"  Category     : {cat}")

    # ── Claude client (.env key) ──────────────────────────────────────────────
    try:
        client = load_client()
    except Phase2ApiError as e:
        print(f"\n  {e}")
        return

    # ── 3. Build the product profile from title + category (Claude) ───────────
    _section("BUILDING PRODUCT PROFILE (Claude)")
    try:
        product = derive_product_profile(
            client,
            title=product_row["title"],
            category=product_row.get("category", ""),
            subcategory=product_row.get("subcategory", ""),
        )
    except Phase2ApiError as e:
        print(f"\n  {e}")
        return
    print(f"\n  Product name   : {product.product_name}")
    print(f"  Product type   : {product.product_type}")
    print(f"  Key attributes : {', '.join(product.key_attributes) or '(none)'}")
    print(f"  Use cases      : {', '.join(product.use_cases) or '(none)'}")
    print(f"  NOT this       : {', '.join(product.not_this) or '(none)'}")
    print(f"  Brand          : {product.brand or '(none)'}")

    # ── Start the Phase 3 JS-sheet generation NOW (main search term known) ────
    # Claude just extracted the product type from the target's title — that is
    # the main search term. The Amazon page-1 scrape + Keepa pull run in a
    # background thread through the rest of Phase 2, so the sheet is ready the
    # moment Phase 3 needs it (no waiting, no Jungle Scout subscription).
    js_autogen = None
    main_term = (product.product_type or "").strip()
    if main_term:
        try:
            from research.phase3.js_sheet import start_background
            print(f"\n  Main search term: '{main_term}' — generating the Phase 3")
            print("  competitor sheet in the background (Amazon page 1 + Keepa,")
            print("  ~2 Keepa tokens per product)...")
            js_autogen = {"keyword": main_term, "job": start_background(main_term)}
        except Exception as e:  # noqa: BLE001 — Phase 3 falls back to manual upload
            print(f"\n  Could not start the JS-sheet generation: {e}")
            print("  Phase 3 will offer auto-generation or a manual file instead.")

    # ── 4. Xray keyword file → candidate keywords ─────────────────────────────
    kw_path = _get_file_path(
        "Paste the full path to your H10 Xray keyword export for this product.",
        "XRAY KEYWORD FILE",
    )
    _section("READING KEYWORDS")
    kw_df, _ = read_xray_keywords(kw_path)
    if kw_df.empty:
        print("\n  No keywords could be read from the Xray keyword file.")
        return

    # ── 5. Deterministic prep (dedupe + SV >= 100) ────────────────────────────
    _section("PREPARING CANDIDATES")
    candidates, report = prepare_candidates(kw_df)
    print()
    for line in report:
        print(line)
    if candidates.empty:
        print("\n  No candidate keywords remain after SV >= 100 filtering.")
        return

    # ── 6. Claude proposes the keywords (Step 5) ──────────────────────────────
    _section("SELECTING KEYWORDS (Claude)")
    print("  Judging relevancy to the product and applying Step-5 exclusions...")
    try:
        selection = select_keywords(client, product, candidates)
    except Phase2ApiError as e:
        print(f"\n  {e}")
        return

    # ── 7. Approval / swap loop (Step 6) ──────────────────────────────────────
    _section("APPROVE KEYWORDS (Step 6)")
    locked = _run_approval_loop(client, product, candidates, selection)
    if locked is None:
        print("\n  Phase 2 cancelled — no keywords locked.")
        return

    # ── 8. Final locked set ───────────────────────────────────────────────────
    _section("FINAL LOCKED KEYWORDS")
    _print_keywords(locked, "Locked launch keywords")

    print()
    print("=" * 65)
    print(f"  Phase 2 complete — {len(locked)} keyword(s) locked.")
    print("  These keywords carry into Phase 3 (competitor analysis sheet).")
    print("=" * 65)
    print()

    # Returned so Phase 3 can consume the target + locked keywords directly.
    # js_autogen carries the in-flight JS-sheet generation (started at profile
    # time); Phase 3 joins it instead of making the user wait or upload a file.
    return {
        "target_asin": asin,
        "keywords": locked,
        "js_autogen": js_autogen,
        "target_product": target_product or {
            "asin": asin, "title": product_row.get("title", ""),
            "category": product_row.get("category", ""),
            "subcategory": product_row.get("subcategory", ""),
        },
    }


if __name__ == "__main__":
    run_phase2()
