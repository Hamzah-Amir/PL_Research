# PL Research Tool — Programmer Spec

Phase 1 of a multi-phase Amazon UK PL research tool. Reads an H10 Black Box XLSX, applies layered filters, scores all survivors, splits them by variation count, and presents results in batches of 10 with a user satisfaction loop. Each batch exports two Excel files (with-variants + without-variants) with embedded product images. Phases 2+ are not yet implemented.

---

## Tech Stack

| Item | Spec |
|------|------|
| Language | Python 3.10+ |
| Environment | `venv/` — all deps isolated |
| Dependencies | `pandas`, `openpyxl`, `numpy`, `Pillow` |
| Input | H10 Black Box XLSX/CSV (one category) |
| Output | `openpyxl` Excel workbook, embedded product images |
| Interface | Interactive CLI (stdin/stdout) |
| Platform | Windows (PowerShell launch scripts) |

> `Pillow` is required for thumbnail generation. Image downloads use stdlib `urllib`.

---

## File Structure

```
PL_Research/
├── phase1/
│   ├── file_reader.py      # XLSX ingestion + column normalisation
│   ├── filters.py          # FilterConfig dataclass + apply_filters
│   ├── scorer.py           # Composite scoring; score_all_products / select_top_products
│   ├── output.py           # openpyxl Excel writer (embedded images)
│   └── main.py             # Entry point — user prompts + batch loop
├── run.py / run.ps1        # Launch scripts
├── setup.ps1               # One-time venv + pip install
├── output/                 # All Excel outputs written here
└── test_file/              # Sample input files
```

---

## User Prompts (only interactive inputs)

| # | Input | Notes |
|---|-------|-------|
| 1 | Budget (GBP, float > 0) | Sets `max_revenue = budget × 1.52`; re-prompt on invalid |
| 2 | Seasonal? Y/N (default N) | Y keeps FLAG products with a note; N removes all seasonal |
| 3 | Category (optional string) | Blank = auto-detect from file (display only) |
| — | File path | Strip surrounding quotes; re-prompt if not found |

No variation filter. Both revenue and monthly sales are read from **ASIN-level** columns.

---

## Module Responsibilities

### `file_reader.py`

**`read_blackbox_file(path) → (df, category)`**

- Accepts `.xlsx`, `.xls`, `.csv`. For XLSX, tries `skiprows=0/1/2` to skip H10 metadata rows (accepts when an ASIN column is found).
- Maps known H10 column name variants to canonical names via `COLUMN_MAPPINGS`. Key mappings:
  - `"Parent Level Revenue"` → `parent_revenue`; `"ASIN Revenue"` → `asin_revenue`
  - `"Parent Level Sales"` → `parent_sales`; `"ASIN Sales"` → `asin_sales`
  - `"Reviews Rating"` → `rating`; `"Sales Trend (90 days) (%)"` → `sales_trend`
  - `"Image URL"` → `image_url`; `"Best Sales Period"` → `best_sales_period`
- Parses all numeric columns: strips `£$€,%-`; H10 `"-"` / `"N/A"` / `""` → `None` (never raises).
- Computes `listing_age_months` from `date_first_available` when a direct age column is absent.
- Appends unmapped raw columns as `raw_{col}` — never discards data.
- Drops rows where ASIN is null or empty.

---

### `filters.py`

**`build_filter_config(category, max_revenue, want_seasonal) → FilterConfig`**  
**`apply_filters(df, config) → (filtered_df, report)`**

**Revenue source:** `_resolve_revenue_source` prefers `[asin_revenue, monthly_revenue, parent_revenue]` — copies chosen series into `monthly_revenue`.  
**Sales source:** `_resolve_sales_source` prefers `[asin_sales, monthly_sales, parent_sales]` — copies into `monthly_sales` for scoring and output. No sales floor filter is applied.

**Filter application order:** revenue cap → F1 → F3 → F4 → F5 → F6 → F7 → F8 → F9 → F2

| Filter | Rule |
|--------|------|
| Revenue cap | `monthly_revenue ≤ max_revenue` (ASIN-level); no minimum |
| F1 Amazon seller | Remove if `seller` or `fulfillment` contains `\bamazon\b` |
| F3 Listing age | `listing_age_months ≥ 6` |
| F4 Rating | `rating > 3.8` (strict greater-than) |
| F5 Price | `price ≥ 5.0` (built-in `_F5_MIN_PRICE`); no upper cap |
| F6 Trend | Remove if `sales_trend < -30`; null = keep |
| F7 Reviews | `reviews ≤ 3000` |
| F8 Categories | `f8_excluded_categories` list (default `[]`) |
| F9 Compatibility | Regex on title: `"for <device>"`, `"compatible with"`, etc. |
| F2 Seasonal | REMOVE: Christmas/Q4 keywords or BSP month 10–12. FLAG: summer/unknown — kept+noted if `want_seasonal`, else removed. PASS = always kept. Applied last. |

`report` format per line: `"{label:<45} {n_removed:>4} removed  ->  {n_remain:>4} remain"`. Absent column = skip silently (0 removed). Each filter resets the DataFrame index.

**`FilterConfig` hardcoded constants:** `_F3_MIN_LISTING_AGE=6`, `_F4_MIN_RATING=3.8`, `_F5_MIN_PRICE=5.0`, `_F6_MAX_DECLINE_PCT=-30.0`, `_F7_MAX_REVIEWS=3000`.

---

### `scorer.py`

**`score_all_products(df) → df`** — scores every row, returns full pool sorted by `composite_score` descending.  
**`select_top_products(df, top_n=10) → df`** — thin wrapper: `score_all_products(df).head(top_n)`.

Sub-scores are normalised across the full input set. If a normalisation range is degenerate (`min == max`), returns `0.5`.

**Sub-scores (weighted base composite):**

| Sub-score | Weight | Formula |
|-----------|--------|---------|
| `_s_revenue` | 35% | **log-norm** of `monthly_revenue`: `log1p(x)` then min-max |
| `_s_opp` | 25% | `sales / (reviews + 1)`, then min-max |
| `_s_sales` | 20% | **log-norm** of `monthly_sales`: `log1p(x)` then min-max |
| `_s_rating` | 12% | `(rating − 3.8) / 1.2` clamped [0, 1] |
| `_s_review` | 8% | hard review bracket (below) |

Log normalisation (not linear min-max) on revenue/sales prevents a single high-revenue outlier from compressing the rest of the pool into a near-zero band.

**Review bracket** (`_review_bracket`): ≤50→1.00, 51–150→0.80, 151–300→0.60, 301–600→0.35, 601–1000→0.15, >1000→0.05, missing/zero→0.10.

**Pipeline:** `base = (weighted sum) × 100`, then two multipliers, then clamp:

1. **Seller multiplier** (`_seller_multiplier`, from `active_sellers`): 1→×1.12, 2–3→×1.06, 4–6→×1.00, >6→×0.92, missing→×1.00.
2. **Trend multiplier** (`_trend_multiplier`, from `sales_trend`): ≥50→×1.10, 20–49→×1.05, −10–19→×1.00, −30–−11→×0.90, <−30→×0.75, missing→×1.00.
3. `composite_score = clip(base × seller_mult × trend_mult, max=100)`.

Internal `_s_*` columns dropped before returning.

**`risk_flags` column** (comma-separated string added per row): `Seasonal — <reason>`, `Low reviews (<50)`, `High reviews (>1000)`, `Low rating (<4.0)`, `Young listing (<12 mo)`, `Declining trend (X%)` (when trend < −10), `FBM seller`, `Single seller` (when `active_sellers == 1`).

---

### `output.py`

**`export_to_excel(batch, batch_num, category, budget, max_revenue, filter_config, variant_label="", filename_tag="") → path`**  
**`print_product_table(batch, batch_num, variant_label="") → None`**

`variant_label` (e.g. `"With Variants"`) appears in the sheet title + Row 1 header and the terminal table; `filename_tag` (e.g. `"WithVariants"`) is embedded in the file name.

**Sheet 1 — `B{N} — {variant_label}`** (title capped at 31 chars)
- Row 1: merged title bar (`… | Batch N — {variant_label}`); Row 2: metadata strip (category, budget, cap, date); Row 3: frozen column headers; Rows 4+: product data; trailing legend rows.
- Image embedding: `_make_thumbnail` downloads via `urllib` (10s timeout, UA header), resizes to 80×80px PNG via Pillow, cached per URL. Returns `None` on any failure. Cell shows `"no image"` if embedding fails — export never crashes.
- Row height: `max(38, 80×0.75 + 2)` to fit thumbnails.

**Column order:** `ASIN, Image, Product Name, Brand, Category, Price (£), Monthly Revenue (£), Monthly Sales, BSR, Rating, Reviews, Listing Age (mo), Fulfillment, Score, Notes / Risk Flags`

**Colour rules:** Score — Green `70AD47` ≥70, Amber `FFC000` 50–69, Red `FF7C80` <50. Notes cell — Yellow `FFE699` if non-empty. Data rows alternate `EFF3FF` / `FFFFFF`.

**Sheet 2 — `Filter Settings`:** two-column table of all filter values used in the run.

**Output path:** `output/Phase1_Batch{N}_{filename_tag}_{YYYYMMDD_HHMMSS}.xlsx` (`output/` created if absent). Each batch writes two files — `…_WithVariants_…` and `…_NoVariants_…`.

---

### `main.py`

```
1.  Banner
2.  _get_budget()           → budget; max_revenue = budget × 1.52
3.  _get_seasonal_choice()  → want_seasonal
4.  _get_category()         → category string (optional)
5.  _get_file_path()        → validated path (re-prompt if not found)
6.  read_blackbox_file()    → df, detected_category
7.  build_filter_config + print_active_filters
8.  apply_filters()         → filtered_df, report
9.  score_all_products()    → scored_all (full ranked pool)
10. Split scored_all by `variations` (Variation Count) into
    with_var (>0) and without_var (≤0/missing); each stays score-ranked
11. Batch loop (BATCH_SIZE = 10), shared offset across both pools:
      a. For each pool (with-variants, without-variants), slice next 10
      b. Skip empties; else print_product_table + export_to_excel
         (variant_label + filename_tag) → save one file per non-empty pool
      c. If either pool has more: prompt "See next 10 of each? Y / Enter=exit"
      d. Stop when satisfied or both pools exhausted
12. Completion message
```

Empty `filtered_df` → print diagnostic (suggestions: raise budget, allow seasonal, broader file) and `return`.

---

## Error Handling

| Scenario | Behaviour |
|----------|-----------|
| File not found | Re-prompt |
| Column absent | Filter skipped silently (0 removed) |
| All products filtered | Diagnostic message + `return` |
| Numeric parse failure | `None` — never raises |
| `min == max` in normalisation | Returns `0.5` |
| Image download / embed fails | Cell shows `"no image"`, export continues |

---

## Known H10 GB Black Box Column Names (Reference)

```
ASIN, Title, Brand, Fulfillment, Category, BSR, Price,
Parent Level Sales, ASIN Sales, Sales Trend (90 days) (%),
Parent Level Revenue, ASIN Revenue, Review Count, Reviews Rating,
Seller, Seller Country/Region, Number of Active Sellers,
Best Sales Period, Listing Age (Months), Image URL,
Number of Images, Variation Count, Sales to Reviews
```

---

## Phase Boundary

Phase 1 ends when the user selects one ASIN from the output Excel. That ASIN is the input for Phase 2.

---

# Phase 2 — Keyword Verification

Input: the single target ASIN from Phase 1 + the Phase 1 Black Box XLSX (for the product title) + an H10 **Xray keyword export** for that product (the Xray Keywords view; it carries the same keyword metrics as Cerebro — Search Volume, Cerebro IQ Score, Competing Products, Title Density, Keyword Sales). (Keepa/JS belong to Phase 3, not here.) Output: the launch keywords Claude selects + the user locks.

**Design rule:** all keyword *judgement* — relevancy to the product and every Step-5 exclusion — is performed **purely by Claude via the API. No deterministic relevancy logic, no heuristic fallback.** If the API key is missing or a call fails, Phase 2 stops and reports it (never guesses).

## Current build scope

Pipeline runs through **Step 5 (Claude proposes keywords)** and **Step 6 (user approval/swap loop, set locked)**, then prints and **returns** the locked set. No Excel export — the locked keywords are consumed directly by Phase 3's competitor analysis sheet.

## File Structure (Phase 2)

```
phase2/
├── xray_reader.py      # H10 Xray keyword XLSX/CSV -> normalised DataFrame (COLUMN_MAPPINGS, numeric parsing)
├── keyword_selector.py # prepare_candidates: dedupe + SV>=100 + SV rank (objective only)
├── claude_client.py    # .env key load; derive_product_profile + select_keywords (structured outputs)
└── main.py             # CLI orchestration: profile -> Step 5 propose -> Step 6 approve/swap -> lock
```

## Config / dependencies

- API key read from a git-ignored **`.env`** in the project root: `ANTHROPIC_API_KEY=sk-ant-...` (loaded via `python-dotenv`).
- New deps: `anthropic`, `python-dotenv` (added to `requirements.txt`).
- Launcher: `run.py` / `run.ps1` **always start at Phase 1**, which hands off into Phase 2 at the end. There is no phase menu. Phase 2 can be exercised in isolation during development with `python -m phase2.main` (prompts for the Black Box file since there's no Phase 1 data).

## Module Responsibilities (Phase 2)

### `xray_reader.py`
**`read_xray_keywords(path) -> (df, raw_row_count)`** — mirrors Phase 1's reader. Tries `skiprows=0/1/2` to skip metadata (accepts when a keyword/phrase column appears). `COLUMN_MAPPINGS` is **confirmed against a real Xray keyword export** and covers its columns: `keyword` (Keyword Phrase), `cerebro_iq`, `search_volume`, `search_volume_trend`, `suggested_ppc_bid` (text range, not parsed), `keyword_sales`, `competing_products`, `title_density`, `competitor_rank`. Broad alias fallback; unmapped columns kept as `raw_*`. Numeric cells parsed (currency/commas stripped, H10 `-` -> None). Drops blank-keyword rows.

### `keyword_selector.py`
**`prepare_candidates(df, min_search_volume=100) -> (candidates, report)`** — objective prep only: drop blank phrases, de-duplicate on normalised phrase (keep highest SV), drop SV < 100, rank by SV descending. `MIN_SEARCH_VOLUME=100`, `TARGET_KEYWORD_COUNT=6`. No relevancy, no flags, no fallback.

### `claude_client.py`
- **`load_client()`** — loads `.env`, requires `ANTHROPIC_API_KEY`, returns `anthropic.Anthropic()`. Raises `Phase2ApiError` (no fallback) on missing dep/key.
- **`derive_product_profile(client, title, category="", subcategory="") -> ProductProfile`** — Claude reads title + category and builds the relevancy anchor: `{product_name, product_type, key_attributes[], use_cases[], not_this[], brand}`. **Text only — the product image is NOT sent.** Model `claude-opus-4-8`, adaptive thinking, `effort=medium`, structured output via `messages.parse`.
- **`select_keywords(client, profile, candidates, target_count=6, exclude=None) -> KeywordSelection`** — Claude judges relevancy against the full `profile` (type + attributes + `not_this` + brand-to-exclude) and applies Step-5 exclusions, returns `keywords[]` (`keyword`, `search_volume`, `relevancy`, `reason`) + `note`. `effort=high`. Rules + rendered profile + candidate list sent as **prompt-cached** system blocks (so the future swap loop reuses the prefix); volatile instruction (`target_count`, `exclude`) sent last. Top `MAX_CANDIDATES_TO_MODEL=200` SV-ranked candidates sent.
- Pydantic schemas: `ProductProfile`, `SelectedKeyword`, `KeywordSelection`. `Phase2ApiError` surfaces all failures.

### `main.py`
**`run_phase2(blackbox_df=None) -> Optional[list]`**: prompt ASIN -> obtain Black Box data -> look up `{title, category, subcategory}` by ASIN -> `load_client` -> `derive_product_profile` (print profile) -> prompt Xray keyword path -> `read_xray_keywords` -> `prepare_candidates` -> `select_keywords` (Step 5 proposal) -> **`_run_approval_loop` (Step 6)** -> print and **return** the locked `SelectedKeyword` list (consumed by Phase 3; no export). Any `Phase2ApiError` or missing data prints a message and returns `None`; empty ASIN cancels.

**Step 6 — `_run_approval_loop(client, profile, candidates, selection, target_count=6)`:** prints the proposed keywords and reads a choice — `Enter`/`y` locks the set; comma/space-separated numbers reject those keywords; `q` cancels. Rejected phrases accumulate into an `exclude` list and approved keywords are kept; it re-calls `select_keywords(target_count=need, exclude=rejected+kept)` for replacements and merges, looping until approval. Returns the locked list or `None` (cancel). Helpers: `_print_keywords`, `_parse_reject_indices`.

**Phase 1 → Phase 2 handoff (automatic):** `phase1.main.run_phase1()` ends by calling `run_phase2(blackbox_df=df)` directly — no confirm prompt — passing the **already-loaded** normalised Black Box frame, so Phase 2 reuses it (`USING PHASE 1 PRODUCT DATA`) and does **not** re-prompt for the file. (Pressing Enter at the Phase 2 ASIN prompt exits if Phase 2 isn't wanted.) Running Phase 2 standalone (`python -m phase2.main`) leaves `blackbox_df=None`, which prompts for the Black Box path.

## Phase Boundary (Phase 2)

Phase 2 ends after the user locks the keyword set via the Step-6 approval/swap loop; the locked set is printed and returned for Phase 3. There is no Excel export — the keywords are an input to Phase 3's competitor analysis sheet. Phase 3 is not yet implemented (see `PHASE3_HANDOFF.md`).
