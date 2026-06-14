# PL Research Tool — Programmer Spec

Multi-phase Amazon UK PL research tool. **Phase 1 is a stored-dataset model:** Helium 10 Black Box exports are ingested into PostgreSQL (scored + segmented, history preserved), and product suggestions are served by **querying that DB by budget — no live API at query time**. The user picks one target ASIN, which flows into Phase 2 (keyword verification) and Phase 3 (competitor analysis).

---

## Tech Stack

| Item | Spec |
|------|------|
| Language | Python 3.10+ |
| Environment | `venv/` — all deps isolated |
| Dependencies | `pandas`, `openpyxl`, `numpy`, `Pillow`, **`psycopg2-binary`** (P1 DB); `anthropic`, `python-dotenv`, `keepa` (P2/P3) |
| Storage | **PostgreSQL** (local in dev, managed in prod) via `DATABASE_URL` in `.env` |
| Input | H10 Black Box XLSX/CSV (one category per export), ingested into Postgres |
| Output | `openpyxl` Excel shortlist, embedded product images |
| Interface | Interactive CLI (stdin/stdout) |
| Platform | Windows (PowerShell launch scripts) |

> Raw SQL via `psycopg2` (no ORM) — simple, fast, denormalised. `Pillow` for thumbnails; image downloads via stdlib `urllib`.

---

## Architecture (Phase 1)

1. **Ingest** — H10 Black Box exports are read (reusing `file_reader.read_blackbox_file`), scored + segmented, and **UPSERTed** into `products` by ASIN. Every ingest also writes a `product_history` snapshot row per product. **Weekly full refresh** (re-export → re-ingest); existing rows are updated in place (keeping `created_at`), new ones inserted, missing ones kept — **history is never deleted**.
2. **Query** — given a budget (+ optional category / seasonal toggle / variants choice) the tool pulls a pool of the highest-`opportunity_score` products and returns the top 10 to Excel; "next 10" pages the same pool with no re-query.
3. **Pick** — the user selects one ASIN; its `{asin, title, category, subcategory}` is handed to Phase 2.

Revenue and sales are taken at **ASIN level** (`ASIN Sales` / `ASIN Revenue`), never parent level.

---

## File Structure

```
PL_Research/
├── phase1/
│   ├── db.py            # get_connection (DATABASE_URL) + init_db (schema/indexes)
│   ├── file_reader.py   # H10 Black Box XLSX -> normalised DataFrame (reused for ingest)
│   ├── scoring.py       # competition/demand/opportunity + buckets + is_seasonal (pure)
│   ├── ingest.py        # read_and_score + build_product_rows (pure) + UPSERT/history (DB)
│   ├── query.py         # price_tiers_for_budget + fetch_pool (filters + soft shuffle)
│   ├── output.py        # export_shortlist (Excel, embedded images) + print_shortlist_table
│   └── main.py          # CLI: ingest + shortlist loop (top10 -> next10) -> pick ASIN -> P2
├── run.py / run.ps1     # Launch scripts (run.py -> phase1.main.run_phase1)
├── setup.ps1            # One-time venv + pip install
├── output/             # Excel outputs + keepa_cache/ (P3)
└── test_file/          # Sample input files
```

(`filters.py` and the old composite `scorer.py` from the pre-DB build are superseded: filtering now happens at query time and scoring lives in `scoring.py`.)

---

## Config

- **`.env`** (git-ignored): `DATABASE_URL=postgresql://user:password@localhost:5432/pl_research` (same var for managed prod). `ANTHROPIC_API_KEY` + `KEEPA_API_KEY` are used by P2/P3.
- `setup.ps1` runs `pip install -r requirements.txt` (now includes `psycopg2-binary`).

---

## Database Schema (PostgreSQL)

```
products(
  id PK, asin UNIQUE, title, brand, category, subcategory,
  price, bsr, estimated_sales, estimated_revenue,        -- ASIN-level
  review_count, rating, weight_kg, yoy_growth, is_seasonal,
  variation_count, has_variants,                          -- variants filter
  is_amazon_seller, is_compatibility, is_global_brand,    -- exclusion flags
  fulfillment, image_url, url, listing_age_months,
  sales_trend_90d, price_trend_90d,
  price_range, review_bucket, sales_bucket,               -- segmentation
  competition_score, demand_score,                        -- legacy, now NULL
  opportunity_score,                                      -- base 0-85 (see Scoring)
  created_at, last_updated )
product_history(id PK, asin, price, bsr, estimated_sales, review_count, captured_at)
product_reviews_insights(id PK, asin, common_complaints, negative_keywords JSONB, created_at)
indexes: products(asin UNIQUE), (category), (price_range), (opportunity_score DESC),
         (has_variants), (is_seasonal); product_history(asin); insights(asin)
```

The columns beyond the bare spec (`weight_kg`, `yoy_growth`, the three `is_*` flags, `variation_count`/`has_variants`, `brand`, `fulfillment`, `image_url`, `url`, `listing_age_months`, `estimated_revenue`, trends) feed the new scoring, the built-in filters, the variants filter, the Excel shortlist and the Phase 2 handoff. Schema creation is idempotent; new columns are added to an existing table via `ALTER TABLE ADD COLUMN IF NOT EXISTS`. `competition_score`/`demand_score` are retained as columns but no longer populated.

---

## Scoring (`scoring.py`, point-based, pure)

User-specified point formula (~115 max). The **base score (0–85) is product-intrinsic and stored** at ingest as `opportunity_score`; the **revenue-proximity factor (0–30) is budget-dependent and added at query time** (`query.revenue_proximity`), so the shortlist's displayed `opportunity_final` = base + proximity.

Base factors: **90-day trend** (+1 per 50%, cap 20) · **YoY** (+1 per 5%, cap 10) · **reviews** (<200=15, <500=10, <1000=5) · **rating** (≥4.5=10, ≥4.2=7, ≥4.0=4) · **weight** (<0.3kg=15, <0.5kg=10, <1.0kg=5) · **price £8–£20** (=10) · **not Amazon seller** (+5) · **variation penalty** (−5 if >10 variations). Revenue proximity: linear 0 at `(cap−1000)` → 30 at `cap`.

Exclusion flags (stored): `is_amazon_seller` (seller contains "amazon" or fulfilment AMZ), `is_compatibility` (title cues — "compatible with / replacement for / fits / case for …", the old F9), `is_global_brand` (brand ∈ `phase1/brand_blocklist.json`, an editable list).

## Segmentation (fixed thresholds)

- `price_range`: `<£15` low · `£15–£50` mid · `>£50` high.
- `review_bucket`: `<100` low · `100–1000` medium · `>1000` high.
- `sales_bucket`: `<100` low · `100–1000` medium · `>1000` high.
- `is_seasonal`: True if `Best Sales Period` month ∈ {Nov, Dec} or the title contains seasonal keywords (christmas/halloween/summer/gift set/…).
- `has_variants` = `Variation Count > 0`.

---

## User Prompts

| # | Input | Notes |
|---|-------|-------|
| 1 | Refresh data? (y/N) | If yes (or DB empty), prompts for Black Box file(s) — a folder or comma-separated paths |
| 2 | Inventory budget (£, optional) | cap = budget × 1.52; pool restricted to the **revenue band `[cap−1500, cap]`** (fixed £1,500 window ending at the cap) |
| 3 | Category (optional) | `ILIKE` filter on stored `category` |
| 4 | Include seasonal? (Y/n) | N excludes `is_seasonal` rows |
| — | Pick ASIN / `n` next 10 / `q` quit | Pagination over the in-memory pool, no re-query |

A **filter funnel** is printed to the console (only) before the shortlist — each filter's removed/remaining counts (shared `_build_filter_steps`, so it matches the pool exactly). There is no variants filter (the variation penalty still applies in scoring).

**Built-in quality filters (always applied, never prompted):** rating > 3.8, listing age ≥ 6 months, monthly sales ≥ 50, 90-day trend > −25% (missing trend kept), a **budget-tiered review window**, and exclude `is_amazon_seller` / `is_compatibility` / `is_global_brand`. Thresholds are constants in **`query.py`** (the legacy `filters.py` is unused).

**Budget-tiered review window** (`review_band_for_budget`, `REVIEW_TIERS`) — bigger budgets can compete with more-reviewed listings; `max=None` = no cap:

| Budget (£) | review min – max |
|---|---|
| `< 2,000` | 50 – 3,000 |
| `2,000–4,999` | 500 – 5,000 |
| `5,000–9,999` | 1,500 – 8,000 |
| `10,000–19,999` | 1,500 – 12,000 |
| `≥ 20,000` | 1,500 – (no cap) |
| no budget | 50 – 3,000 |

---

## Module Responsibilities (Phase 1)

### `db.py`
`get_connection()` opens psycopg2 from `DATABASE_URL` (raises `Phase1DbError`, no fallback). `init_db(conn)` runs the idempotent schema SQL. `SCHEMA_SQL` holds the 3 tables + indexes.

### `file_reader.py`
`read_blackbox_file(path) -> (df, category)` — accepts `.xlsx/.xls/.csv`, skips H10 metadata rows, normalises columns via `COLUMN_MAPPINGS` (incl. `asin_sales`, `asin_revenue`, `variations`, `best_sales_period`, `image_url`, `bsr`, `reviews`, `rating`, `listing_age_months`, `fulfillment`, `sales_trend`). Reused unchanged for ingestion.

### `scoring.py`
`compute_scores(df) -> df` adds the **base `opportunity_score` (0–85)**, segmentation buckets, `is_seasonal`, `variation_count/has_variants`, `weight_kg`, `yoy_growth`, `sales_trend_90d`, and the three exclusion flags. Pure, individually-testable helpers: `base_score`, `price_range`, `review_bucket`, `sales_bucket`, `is_seasonal`, `parse_weight_kg`, `is_amazon_seller`, `is_compatibility`, `is_global_brand` (loads `brand_blocklist.json`).

### `ingest.py`
- **Pure:** `read_and_score(paths)` reads + concatenates + de-dups by ASIN, then scores. `build_product_rows(scored)` maps to `products`-row dicts (NaN→None, numpy→python; `estimated_sales/revenue` from `asin_sales/asin_revenue`; `url` from ASIN).
- **DB:** `ingest_files(conn, paths)` = score → snapshot `product_history` → UPSERT `products` (`ON CONFLICT (asin) DO UPDATE … last_updated=NOW()`), via `execute_values`.

### `query.py`
`max_revenue_for_budget(budget)` = `budget × 1.52`; `revenue_band` = `[cap−1500, cap]`; `revenue_proximity(rev, cap)` = 0–30 (0 at band floor → 30 at cap). `_build_filter_steps(max_revenue, category, include_seasonal)` is the single source of filter clauses (used by both `fetch_pool` and `filter_funnel`). `fetch_pool(conn, budget, max_revenue=None, category, include_seasonal, pool_size=100, shuffle=True)` applies the band + built-in quality filters + seasonal, fetches matches (`LIMIT 2000`), then in Python adds `revenue_proximity` → `opportunity_final = base + proximity`, `_soft_shuffle`-ranks, returns top `pool_size`. `filter_funnel(...)` re-runs the steps cumulatively for the console exclusion report.

### `output.py`
`export_shortlist(products, batch_num, budget, category, variants, include_seasonal, pool_total)` writes a styled, image-embedded Excel batch (rank, ASIN, image, title, brand, sub-category, price, ASIN revenue/sales, BSR, rating, reviews, variants, opportunity/competition/demand, seasonal) with an opportunity colour scale + legend, timestamped to `output/`. `print_shortlist_table` prints the console summary.

### `main.py`
`run_phase1()` — connect + `init_db`; show stored count; optional ingest/refresh; ask query params; `_shortlist_loop` pages the pool (top 10 → Excel + console, then `n` for next 10) until the user enters an ASIN; hands the selected `{asin,title,category,subcategory}` to `run_phase2(target_product=…)`, then on to `run_phase3`. `python -m phase1.main ingest` runs ingest-only.

---

## Phase Boundary

Phase 1 ends when the user selects one ASIN from the shortlist. The DB-backed Phase 1 hands Phase 2 a `target_product` dict (`{asin, title, category, subcategory}`) directly — Phase 2 no longer needs the Black Box file.

---

# Phase 2 — Keyword Verification

Input: the target product (ASIN + title/category) carried from Phase 1 + an H10 **Xray keyword export** for that product (the Xray Keywords view; it carries the same keyword metrics as Cerebro — Search Volume, Cerebro IQ Score, Competing Products, Title Density, Keyword Sales). (Keepa/JS belong to Phase 3, not here.) Output: the launch keywords Claude selects + the user locks.

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
- Deps: `anthropic`, `python-dotenv`.
- Launcher: `run.py` / `run.ps1` **always start at Phase 1**, which hands off into Phase 2 at the end. Phase 2 can be exercised in isolation with `python -m phase2.main` (prompts for the ASIN + Black Box file since there's no Phase 1 data).

## Module Responsibilities (Phase 2)

### `xray_reader.py`
**`read_xray_keywords(path) -> (df, raw_row_count)`** — tries `skiprows=0/1/2` to skip metadata (accepts when a keyword/phrase column appears). `COLUMN_MAPPINGS` confirmed against a real Xray keyword export: `keyword`, `cerebro_iq`, `search_volume`, `search_volume_trend`, `suggested_ppc_bid` (text), `keyword_sales`, `competing_products`, `title_density`, `competitor_rank`. Broad alias fallback; unmapped columns kept as `raw_*`; numeric parsing (currency/commas stripped, H10 `-` -> None).

### `keyword_selector.py`
**`prepare_candidates(df, min_search_volume=100) -> (candidates, report)`** — objective prep only: drop blanks, de-duplicate on normalised phrase (keep highest SV), drop SV < 100, rank by SV descending. No relevancy, no fallback.

### `claude_client.py`
- **`load_client()`** — loads `.env`, requires `ANTHROPIC_API_KEY`, returns `anthropic.Anthropic()`; raises `Phase2ApiError` on missing dep/key.
- **`derive_product_profile(client, title, category="", subcategory="") -> ProductProfile`** — Claude builds the relevancy anchor `{product_name, product_type, key_attributes[], use_cases[], not_this[], brand}`. Text only. `claude-opus-4-8`, adaptive thinking, `effort=medium`, `messages.parse`. **Reused by Phase 3.**
- **`select_keywords(client, profile, candidates, target_count=6, exclude=None) -> KeywordSelection`** — Claude judges relevancy + Step-5 exclusions; rules + profile + candidate list sent as **prompt-cached** system blocks; volatile instruction last. `effort=high`.
- Pydantic schemas: `ProductProfile`, `SelectedKeyword`, `KeywordSelection`. `Phase2ApiError` surfaces all failures.

### `main.py`
**`run_phase2(blackbox_df=None, target_product=None) -> Optional[dict]`**: obtain the target product — preferred via **`target_product`** (the DB-backed Phase 1 handoff: `{asin,title,category,subcategory}`, no ASIN prompt / no file); else via the legacy `blackbox_df` lookup; else prompt for ASIN + Black Box path. Then `load_client` → `derive_product_profile` → prompt Xray keyword path → `read_xray_keywords` → `prepare_candidates` → `select_keywords` (Step 5) → `_run_approval_loop` (Step 6) → **return** `{"target_asin", "keywords", "target_product"}`. Any `Phase2ApiError`/missing data returns `None`.

**Step 6 — `_run_approval_loop(...)`:** prints proposed keywords; `Enter`/`y` locks, numbers reject (re-pick replacements via `select_keywords(exclude=…)`), `q` cancels.

**Phase 1 → Phase 2 handoff (automatic):** `phase1.main.run_phase1()` ends by calling `run_phase2(target_product=selected)` with the chosen product's `{asin,title,category,subcategory}` from the DB — no ASIN prompt, no Black Box file. Standalone `python -m phase2.main` prompts for both.

## Phase Boundary (Phase 2)

Phase 2 ends after the user locks the keyword set via the Step-6 loop; the locked set is returned (in a dict) for Phase 3. `run_phase2()` returns `{"target_asin", "keywords", "target_product"}`.

---

# Phase 3 — Competitor Identification + Analysis

Input: the target ASIN + locked keywords from Phase 2, plus the Step-7 inputs — the **JS-style keyword sheet, auto-generated** (no Jungle Scout needed: Phase 2 starts `js_sheet.start_background(main_term)` the moment Claude's profile yields the main search term from the target's **title**; Phase 3 joins the job via the `js_autogen` param — manual JS CSV upload is the fallback), **one H10 Xray *Products* export per launch keyword** (the `Sponsored` column is keyword-specific, so 5–6 files), and the **Keepa API** (live, key in `.env` as `KEEPA_API_KEY`). Output: the filled **Competitor Analysis workbook** (Competitor Analysis + Pricing Analysis + Sponsored Products tabs) and (pending) a Competitor PDF.

**Design rule (unchanged):** deterministic data ops in Python; all *judgement* via Claude (same-product-type, listing/image quality, review analysis) — no heuristic fallback (Rule 0). Keepa/Anthropic failures surface; never guessed.

## Data-source map (who supplies what)

- **JS-style sheet (auto-generated by `js_sheet.py`; manual JS CSV = fallback)** — monthly sales/revenue (the sales source; Amazon "bought in past month" badge, not JS's estimate), dimensions/weight, parent-ASIN universe. Product-level → main search term only.
- **Xray Products XLSX × (5–6 keywords)** — `Sponsored`/`Sponsored Brand`/`Sponsored Brand Video` flag (per keyword), review velocity, Best-Seller flag, rating, review count, creation date, sales/revenue, fulfilment (FBA/AMZ/MFN), seller, seller age.
- **Keepa API (deep pull, final 3 only)** — variations (count + attrs), bullets (`features`), images, FBA pick&pack fee, referral %, sub-category, coupon, price/BSR (`stats_parsed.current`), A+ flag (`aplus`), video (`videos`), buy-box seller → all-time feedback (`seller_query`).

## File Structure (Phase 3)

```
phase3/
  js_reader.py            read_js_products(path) -> (df, n)        — JS CSV, COLUMN_MAPPINGS
  xray_product_reader.py  read_xray_products(path) + aggregate_sponsored(per_kw)
  keepa_client.py         build_asin_pool · load_keepa · query_products(deep, cache, limit)
                          · get_seller_feedback · extract_keepa_fields
  claude_client.py        filter_same_product_type · score_listing (vision) · analyze_reviews
                          (reuses phase2 derive_product_profile / ProductProfile)
  competitor_selector.py  build_candidate_records · select_top_competitors -> top 3
  scorer.py               score_competitor (184-pt rubric) · pool_statistics
  output.py               write_competitor_workbook (label-driven fill)
  main.py                 run_phase3(target_asin, locked_keywords, dev_inputs)
  reviews_flyby.py        fetch_reviews(asin) via FlyBy/RapidAPI Real-Time Amazon Data
                          (RAPIDAPI_KEY) — ALL 1-2★ + max five 4-5★, 3★ excluded
  asin_scraper.py         scrape_asins(keyword|url) -> ranked ASIN list (page 1 default)
  js_sheet.py             build_js_sheet(keyword|asins) -> JS-style CSV via Keepa
                          · start_background(term) — launched by Phase 2 at profile time
```

## Token discipline (Keepa)

Hits the **Product Request** endpoint via `keepa.query()` (domain `GB`), 1 token/ASIN, ≤100 per call. **Selection uses 0 Keepa tokens** (JS+Xray only). Only the **final 3** get a Keepa pull. Shallow pull = `history=1, stats=180` (no extra cost; skips `rating` since Xray supplies it). Deep pull (final 3) adds `offers/buybox/rating/aplus/videos` for seller feedback + A+/video. Every response is **pickle-cached per (ASIN, depth)** in `output/keepa_cache/` so dev never re-spends.

## The 184-point rubric (`scorer.py`)

Classic Just One Dime sheet (159) + 4 approved data-driven elements (+25). Higher score = stronger competitor. Bands: **0–86 easy · 87–134 mid · 135–184 difficult**.

- **Deterministic (127 pts):** Price (25, relative), Sponsored Products (20, Xray), Sponsored Brands/ex-AMS (10, Xray), Reviews (10, category-percentile), Rating (5), Age (5, Date First Available), Bestseller (10, flag+BSR), Seller (5, all-time feedback), FBA/FBM (2), Review Velocity (5), Variations (5), Sales/Revenue Strength (10, percentile), Enhanced Content A+/Video (5).
- **Claude judgement (57 pts):** Product Images (20, vision), Title (5), Bullets/Description (5), Unique Design (2), Marketing Images (15), Pricing Strategy (10). Marketing title/bullets reuse the product-page judgement.
- **Unknown (Rule 0):** any element with no verified source scores `None` (not counted) and is listed in `unknown_elements`; the sheet leaves Score blank and writes "Unknown — need user confirmation" in Info. Common Unknowns: seller feedback (needs deep `offers` pull / 3P buy box), review velocity (when absent from Xray).

## Module notes

- **`competitor_selector.py`** — `build_candidate_records` merges JS + per-keyword Xray into one record/ASIN (sales = JS first, else Xray ASIN sales; `keyword_presence` = count of DISTINCT keyword files the ASIN appears in). `select_top_competitors` calls Claude's same-type filter, then ranks survivors by monthly sales (tie-break presence, then reviews) → top 3. FBA and FBM both eligible.
- **`scorer.py`** — relative elements (price, sales strength) computed across the set/pool; `_parse_date` tries US `mm/dd/yyyy` first (JS/Xray are US-format), then `dd/mm`, then text.
- **`reviews_flyby.py`** — competitor reviews via the FlyBy/RapidAPI **Real-Time Amazon Data** API (`RAPIDAPI_KEY` in `.env`; host `real-time-amazon-data.p.rapidapi.com`). **Gated access:** the `/product-reviews` endpoint takes the logged-in session cookie (`AMAZON_COOKIE` from `.env`) as a `cookie` query param — with a fresh cookie the star filters + pagination work and the **full 1-2★ list is pulled** (1_STARS then 2_STARS, ≤10 pages each, deduped by review_id; verified live: 20 weaknesses for B0C94Q85JV). Strengths come from one TOP_REVIEWS page. Selection rule (user decision): **all 1-2★, max five 4-5★, 3★ excluded**. A dead/expired cookie is auto-detected (off-band stars on page 1) → falls back to the ~8-10 public detail-page reviews and says so in `note` (Rule 0); the `rating_distribution` histogram always supplies the true negative share. ~3-15 requests/ASIN (free tier 2,000/mo). Replaces the deleted `review_scraper.py` and the OpenClaw route. **Ops note: refresh `AMAZON_COOKIE` when runs start reporting "cookie appears expired".**
- **`output.py`** — `write_competitor_workbook` loads the template and fills each of the 3 blocks by **locating grading rows via their column-E label** (section-aware for the duplicated Images/Title/Bullets), so it survives template edits (adding the 4 new "MARKET MOMENTUM" rows, renaming "AMS Ads"). Writes Info (col F) + Score (col J); leaves the `Total Score` `SUM` formula to pick scores up. The tall merged box under each "N Competitor" label (B11:C17 / B57:C63 / B104:C110) gets the **product title** (centred, wrapped — user decision: no image there). In both grading sections: the **Product Images** Info cell gets the **main product image embedded** (downloaded from the Keepa image URL, scaled to the row height; text "N listing image(s), video present" as fallback if the download fails), the **Product Description/Bullet Points** Info cell gets the first 5 bullets (description as fallback), and the **Product Title** Info cell gets the actual title — all only when the scorer left no info text of its own. Price Info is written numeric (feeds Pricing Analysis math). In Pricing Analysis: clears Rule-0 cost placeholders (COGS/shipping/HTS), replaces the 'Electric Balloon Pump' placeholder (merged D11:F11) with the **target product's title** from Phase 1 (`target_title`), writes the **90-day average BSR** (Keepa `avg90` via `bsr_90d`, current-BSR fallback) in col G, repairs the Top-3 row's broken `#REF!` refs (D/E/F/I/L15 → block-3 cells C115/C112/C111/F106/C114), rewrites the FBA-payout formula (col J) with each competitor's **actual Keepa referral % + pick&pack fee** (template's +0.5 supplier-shipping placeholder kept), and writes the per-unit freight formulas (M/N = `AG×1.1×3` sea / `×6` air per the header note) plus dims/weight; writes the 6 keywords into Sponsored Products (bids blank). Saves a timestamped file in `output/` — never overwrites the template.

## Standalone: `asin_scraper.py` (search-page ASIN scraper)

Groundwork for replicating the JS keyword sheet without Jungle Scout: scrape the ranked ASIN list straight from Amazon search results, then feed it to Keepa. **Not wired into the Phase 3 pipeline yet.** Same HTTP stack as `review_scraper.py` (`curl_cffi` Chrome-impersonation, optional `AMAZON_COOKIE`). `scrape_asins(keyword=… | url=…, marketplace="co.uk", max_pages=1)` scrapes **page 1 by default** (raise `--max-pages` to walk `…/s?k=<kw>&page=N`), parses only the real product grid (`div[data-component-type="s-search-result"]` → `data-asin`) so sponsored-brand carousels/banners (`AdHolder`, empty `data-asin`) are excluded; inline sponsored tiles are flagged via their sponsored label and can be dropped with `--organic-only`. CAPTCHA pages are detected + retried with backoff; jitter delay between pages; dedupe keeps the first (highest-ranked) hit. CLI: `python -m phase3.asin_scraper "tent pegs" --csv out.csv` (or `--asins-only`, `--url`). Validated 2026-06-12: "tent pegs" page 1 → 48 ASINs; full 7-page walk → 301 ASINs covering 53/55 ASINs in the reference JS export (the 2 misses were dead listings, BSR 53k+).

## `js_sheet.py` (JS keyword-sheet replica via Keepa — wired into the pipeline)

**Pipeline wiring:** Phase 2 calls `start_background(product.product_type)` immediately after `derive_product_profile` (the main search term IS the Claude-extracted product type from the target's title — NOT one of the 6 locked keywords) and returns the in-flight job as `js_autogen` in its result dict; `phase1.main` passes it to `run_phase3(js_autogen=…)`, which joins the thread after the Xray prompt — by then it has normally finished, so the user never waits. Fallbacks: if the job failed (or standalone Phase 3), the JUNGLE SCOUT FILE prompt accepts a path **or `kw:<keyword>`** to generate on the spot. A cancelled Phase 2 still leaves the Keepa pulls cached for the next run.

`build_js_sheet(keyword=… | asins=…)` chains `asin_scraper` (page 1) → Keepa **Product Request** (batched ≤100/call, `history+stats+rating`, ~2 tokens/ASIN; `--buybox` adds buy-box price/fulfilment for ~+2) → **CSV** in `output/js_sheet_<kw>_<date>.csv` mirroring the JS export's columns plus `Position`, `Sponsored`, `Sales Source` (summary averages are printed, not written into the CSV). `read_js_products` maps all 21 canonical columns from the generated CSV (verified). Key derivations: Monthly Units = Keepa `monthlySold` (Amazon's bucketed "bought in past month" badge — NOT JS's BSR-model estimate; fallback = 30-day rank drops, labelled in `Sales Source`); Fees = pick&pack + price×referral%; Net = price − fees; Product Tier = UK size-tier ladder from package dims/weight; `LQS (est)` = own 1-10 heuristic (JS's is proprietary). Rule 0: missing Keepa data stays blank. Caching: depth `"sheet"` in `output/keepa_cache/`, borrows free from the phase-3 `deep` cache; `--limit N` caps uncached ASINs per run (token guard — re-runs resume from cache). `--asins-file` takes a plain list or an `asin_scraper` CSV (keeps position/sponsored, avoids re-scraping). CLI: `python -m phase3.js_sheet "tent pegs"`. Validated 2026-06-12: full 48-ASIN page-1 sheet built; 41 badge + 6 rank-drop estimates + 1 honest blank; spot-on vs JS for price/reviews/BSR/sellers.

## Template requirement (4 new rows) — DONE

**`test_file/competitor_analysis/Compatitor Analysis.xlsx`** (the pipeline default in `main.py`; the user renamed the upgraded "184" file over the original) carries the 4 extra grading rows per competitor block, inserted after each Pricing Strategy row with exact column-E labels / column-K maxes: `Review Velocity` (5), `Product Variations` (5), `Sales / Revenue Strength` (10), `Enhanced Content (A+/Video)` (5) at rows 33-36 / 79-82 / 126-129, with Total+Max Score SUMs extended (`J13:J36` / `J59:J82` / `J106:J129`) and the Pricing Analysis cross-sheet refs re-pointed at the shifted block-2 rows. (Build note: openpyxl `insert_rows` moves only cells — merges, row heights and cross-sheet formulas were re-applied at shifted positions manually.)

## Phase 2 → Phase 3 handoff (automatic)

`phase1.main.run_phase1()` captures `run_phase2()`'s return and, if keywords were locked, calls `run_phase3(target_asin, locked_keywords, js_autogen=p2["js_autogen"])`. The JS sheet arrives via the `js_autogen` background job (started in Phase 2 at profile time); Phase 3 only prompts for the Xray files (and for a JS file only as fallback). Standalone dev: `run_phase3(dev_inputs={"js":..., "xray_dir":..., "target_asin":..., "deep":False})` bypasses prompts.

## Phase Boundary (Phase 3)

Phase 3 ends at the **Step 11 Decision Gate**: A workable → Phase 4; B change competitors → re-pick (loop to Step 8); C not workable → Phase 1. Pending items: the Competitor **PDF** report, the **deep `offers`** seller-feedback pull on the final 3, and a first **live multi-keyword run** once all 6 Xray exports are supplied. Phases 4–7 remain PDF-only.
