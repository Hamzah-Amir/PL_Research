# PL Research Tool — Client Guide

A multi-phase Amazon UK Private Label research tool. **Phase 1** keeps a stored database of products (from your Helium 10 Black Box exports), scores them for opportunity, and — based on your budget — hands you a ranked Excel shortlist to pick from. The ASIN you choose flows into **Phase 2** (keywords) and **Phase 3** (competitor analysis).

---

# Phase 1 — Product Hunting (database-backed)

Instead of processing a spreadsheet each time, Phase 1 now **stores your products in a database**. You load Helium 10 Black Box exports once (and refresh weekly); after that, you just ask for products by budget and the tool answers instantly from the stored, pre-scored data — no waiting on live lookups.

## Before You Start

| Item | What it is |
|---|---|
| PostgreSQL | A running PostgreSQL database (local on your machine in development) |
| `DATABASE_URL` | In a file named **`.env`** in the project folder: `DATABASE_URL=postgresql://user:password@localhost:5432/pl_research` (never committed to git) |
| Black Box export(s) | Helium 10 → Black Box → export one category as XLSX (you can load several at once) |

> You only need to load Black Box data when setting up or doing the **weekly refresh**. Day to day, you just query.

## How to Run

```
.\run.ps1
```

The tool connects to your database and:
1. **Tells you how many products are stored.** If it's empty (or you say yes to "refresh?"), it asks for your Black Box file(s) — point it at a folder or paste the paths — and loads them.
2. **Asks what you're looking for:** your inventory **budget** (£), an optional **category**, and whether to **include seasonal** products. It then prints a **filter funnel** (how many products each filter removed) before the shortlist.
3. **Shows the top 10** highest-opportunity matches as a ranked Excel sheet (with product photos) in the `output/` folder, and a summary in the window.
4. **Not satisfied?** Type `n` to see the **next 10** from the same pool (it does not re-query — it just shows more). Repeat until you find one.
5. **Pick a product** by typing its ASIN. That product is carried into Phase 2 automatically — you won't need to re-enter anything.

## How products are scored

Every product gets an **Opportunity Score (out of ~115)** — higher is better. Points are awarded for:

- **Revenue fit (30)** — how close the product's monthly revenue is to your budget cap (closer = better).
- **90-day sales trend (20)** and **year-over-year growth (10)** — rising sales score higher.
- **Reviews (15)** — fewer is better (<200 = full marks); easier to compete.
- **Rating (10)** — ≥4.5 best.
- **Weight (15)** — lighter = lower FBA cost (<0.3kg = full marks).
- **Price £8–£20 (10)**, **not sold by Amazon (5)**, minus a small penalty for products with **more than 10 variations**.

Products are also tagged into buckets (price / reviews / sales) and a **seasonal** flag.

## Built-in quality filters (always on)

To keep results realistic for a new seller, every shortlist automatically excludes/limits:

- **Rating > 3.8**, **listing age ≥ 6 months**, **monthly sales ≥ 50**, **sales not declining more than 25%**, and a **budget-tiered review window** (see below).
- **No Amazon-as-seller** products, **no compatibility/spare-part** products (e.g. "replacement for…", "fits…"), and **no global/major brands** (Dylon, Duracell, Le Creuset, etc.).

The big-brand list lives in `phase1/brand_blocklist.json` — open it to add or remove names (currently tuned for Home & Kitchen).

**Review window scales with your budget** — a bigger budget lets you take on more-established (higher-review) competitors:

| Budget | Reviews allowed |
|---|---|
| under £2,000 | 50 – 3,000 |
| £2,000 – £5,000 | 500 – 5,000 |
| £5,000 – £10,000 | 1,500 – 8,000 |
| £10,000 – £20,000 | 1,500 – 12,000 |
| over £20,000 | 1,500 and up (no cap) |
| no budget entered | 50 – 3,000 |

The funnel prints the exact window used for your budget (e.g. `Review count (500-5,000)`).

## How budget maps to products

Your budget sets a **revenue band**: cap = budget × 1.52, and you only see products earning **between (cap − £1,500) and the cap** per month. Example: £3,000 → cap £4,560 → products doing **£3,060–£4,560/month**; £10,000 → £13,700–£15,200. This keeps the shortlist to products realistically matched to your capital (not the giant sellers). Revenue/sales are always **ASIN-level**, not parent.

> Tip: a tight band + the quality filters can make the pool small — widen the budget or include seasonal to see more.

## Weekly refresh

Each week, export fresh Black Box data and load it (answer **y** to "refresh?"). The tool updates existing products in place and **keeps the history** — nothing is lost.

## FAQ

**"DATABASE_URL not found"** — Add it to your `.env` file (see Before You Start). Phase 1 needs the database.

**Do I upload a file every time?** No. You load Black Box data on setup and once a week; the rest of the time you just query the stored data.

**Can I get the same 10 every time?** The shortlist is lightly shuffled so repeated runs vary a little, while the strongest products stay near the top.

---

# Phase 2 — Keyword Verification

Phase 2 takes the product you chose in Phase 1 and finds the best keywords to launch with. **Claude itself decides which keywords are relevant to your product** — there is no rule-of-thumb shortcut.

**Production model (database-backed, no paid data APIs):** keyword data lives in your **database**, exactly like products in Phase 1. You download the keyword export for **every product in your database** and load them all in, keyed by ASIN. When Phase 1 hands an ASIN to Phase 2, the tool simply **queries that ASIN's keywords from the database** — nothing to upload at run time. The expensive Helium 10 / Jungle Scout *keyword APIs* are deliberately **not** used, and scraping search volume off Amazon isn't possible (it's H10's proprietary estimate, not published on Amazon).

## Before You Start

| Item | Source |
|------|--------|
| Target product | Carried over automatically from your Phase 1 selection (no need to re-enter the ASIN, title, or category) |
| Keyword data in the DB | Helium 10 **Cerebro** (or **Xray → Keywords**) export, XLSX, for each product — **loaded into the database** keyed by ASIN, alongside your Phase 1 product data |
| Anthropic API key | In your **`.env`**: `ANTHROPIC_API_KEY=sk-ant-...` (Claude does the keyword judgement) |

> Keyword data is loaded the same way as products: a bulk ingest you run when setting up / refreshing. By the time you reach Phase 2, the data is already in the database — no per-run upload. Keepa and Jungle Scout belong to Phase 3.

## How to Run

Phase 2 runs as a continuation of Phase 1 (`.\run.ps1`). When Phase 1 finishes and you've picked an ASIN, it **moves into Phase 2 automatically** — your selected product (and its title/category) carries over, and the tool **looks up that ASIN's keywords in the database**. You are not asked for the ASIN or any file. If an ASIN has no keyword data stored, it tells you to load that product's Cerebro/Xray export, then continues.

## What the Tool Does

1. **Uses your selected product** (title + category carried from Phase 1).
2. **Builds a product profile** — Claude works out the product's name, type, key distinguishing features, what it's used for, what it is *not*, and the brand to exclude. This is what keyword relevance is judged against. *(Text only — the product image is not sent.)*
3. **Reads the ASIN's keywords from the database** and tidies the list: removes duplicates and anything under 100 monthly searches.
4. **Claude proposes the launch keywords** — judging relevance against the profile and removing brand names, misspellings, wrong-category, different-product, overly broad, and question terms. Each pick has a one-line reason.
5. **You approve or swap** (Step 6): press **Enter** to lock the set, type **numbers** (e.g. `2,5`) to swap some out for the next-best, or **q** to cancel. Repeat until locked.
6. **Prints the locked keywords** with search volume, relevancy, and reason. These carry into **Phase 3** — no separate file to export.

## Loading keyword data into the DB

Download Cerebro (or Xray → Keywords) for your products as XLSX and run the keyword ingest — it stores each keyword row keyed by ASIN (history kept, same idea as products). Refresh it on the same cadence as your Black Box product data. After ingest, every Phase 2 run is fileless.

## FAQ

**"ANTHROPIC_API_KEY not found"** — Add `ANTHROPIC_API_KEY=sk-ant-...` to your `.env`. Phase 2 will not guess keywords without a working key.

**"No keyword data stored for this ASIN"** — That product's keyword export hasn't been ingested yet. Load its Cerebro/Xray Keywords export and re-run.

**Why not call an API or scrape it?** Helium 10 / Jungle Scout keyword APIs are too costly, and search volume can't be scraped from Amazon (it's H10's estimate, not shown on Amazon). The download-into-DB model is free, stable, and reusable.

**Do I still need the Black Box file in Phase 2?** No. Your product's title and category come straight from the Phase 1 database now.

---

# Phase 3 — Competitor Analysis

Phase 3 takes your target product + the locked keywords from Phase 2, finds your **top 3 real competitors**, scores each one against a 184-point rubric, and fills the **Competitor Analysis workbook** for you. It starts automatically after Phase 2.

## Before You Start

| You provide | What it is |
|---|---|
| ~~Jungle Scout CSV~~ | **No longer needed** — the tool builds this sheet itself. As soon as Claude reads your target product's title in Phase 2, it extracts the main search term and starts generating the sheet in the background (Amazon page 1 + Keepa). By the time Phase 3 needs it, it's ready — no waiting. **If scraping fails** (Amazon block / no working proxy), Phase 3 automatically falls back to a manually-provided JS sheet — `test_file/js-keyword-search.csv` by default, or set `PHASE3_FALLBACK_JS` in `.env` to your own file — and shows an amber note that the fallback was used. |
| Xray Products exports | One H10 Xray **Products** export per launch keyword (5–6 files) — the "Sponsored" column tells us who advertises on each keyword |
| Keepa API key | In your `.env` as `KEEPA_API_KEY=...` — the tool pulls Keepa data live, so **no manual Keepa export is needed** |
| RapidAPI key (FlyBy) | In your `.env` as `RAPIDAPI_KEY=...` — used to pull competitor **reviews** (ALL 1–2★ + top five 4–5★; 3★ ignored). Free tier covers ~2,000 requests/month; a Phase 3 run uses ~10–40. **If the subscription lapses** the Phase 3 page shows an amber warning and the Top-3 Strengths / Weaknesses / Solutions cells say "Unknown — need user confirmation" — resubscribe to *Real-Time Amazon Data* on rapidapi.com and re-run |
| Amazon session cookie | In your `.env` as the four split values `AMAZON_SESSION_ID=...`, `AMAZON_UBID_ACBUK=...`, `AMAZON_X_ACBUK=...`, `AMAZON_AT_ACBUK=...` (copy each from your browser's cookie panel for amazon.co.uk; a single full `AMAZON_COOKIE=...` string still works as a fallback). Passed to the reviews API to unlock the **full 1–2 star list** (Amazon login-gates it). When it expires the tool says so and falls back to the small public sample — just refresh the values from your browser |

> The auto-generated sheet covers the **main search term only** (sales numbers are the same regardless of keyword). Xray must be run on **every** launch keyword, because ad presence changes per keyword. Generating the sheet costs about 2 Keepa tokens per product (~100 for a page); re-runs on the same term are cached and free.

## What the Tool Does

1. **Builds the competitor pool** from the auto-generated keyword sheet + your Xray files.
2. **Claude keeps only same-product-type listings** (drops accessories, look-alikes, wrong types) — both FBA and FBM sellers.
3. **Picks the top 3** by monthly sales and prints a ranked table (ASIN, brand, sales, rating, FBA/FBM, BSR).
4. **Pulls deep data from Keepa** for just those 3 (variations, A+, fees, all-time seller feedback) — this keeps Keepa token use tiny.
5. **Scores each competitor out of 184.** A **higher score = a stronger competitor** (harder to beat). Bands: **0–86 easy to enter · 87–134 mid-challenge · 135–184 difficult.**
6. **Fills the workbook** — Competitor Analysis (all three blocks), Pricing Analysis (BSR/dimensions/weight; costs left blank until you have supplier quotes), and Sponsored Products (your 6 keywords). Saved to `output/`.
7. **Decision Gate:** you decide — (A) workable → Phase 4, (B) change competitors, or (C) not workable → back to Phase 1.

# Phase 4 — Critical Sheet

Phase 4 builds the **Critical Sheet** — the full product-data spreadsheet covering the top parent products in your niche **plus every variation** (colours/sizes/pack counts), one row each. It reuses the Jungle Scout universe from Phase 3 and pulls the variation ASINs **live from Keepa** — you don't run or upload any Keepa export.

**How it works (one approval step):**

1. On the Phase 4 page, click **"Propose vocabulary"**. The tool pulls the parents + their variations from Keepa, then Claude proposes a **controlled vocabulary** (Material / Size / Color / Packaging / Special Features) and the **category-specific Design attributes** that decide how products get grouped into Designs (Rule 8).
2. **You review and edit** the proposed lists in the form (add/remove values, tweak the Design attributes), then click **"Approve & build"**. Nothing is written until you approve.
3. Claude classifies each parent's **Design code** + attributes (image + text), variations inherit the parent's static attributes but keep their **own** sales/price/rating/BSR/link (never copied from the parent — Rule 7), and the tool writes **`Phase4_CriticalSheet_<ASIN>_<date>.xlsx`** to download. A **Design legend** sits above the header.

> **Output format:** the Critical Sheet is filled into your real **`PES UK.xlsx`** template (the `Critical sheet` tab — same columns, styling, date/keyword cells and Design-legend block). Keep `PES UK.xlsx` in the `test_file/` folder; if you update it, the tool picks up the change automatically on the next run. The current Phase-4 download contains the Critical sheet only (the other PES tabs belong to phases 5–7, which aren't built yet). Anything the tool can't verify shows **"Unknown — need user confirmation"** (Rule 0) — e.g. a variation's monthly sales when Keepa has no estimate for it.

## What "Unknown" means in the sheet

Some things can't be read from data (e.g. all-time seller feedback without a deeper Keepa pull, or AMS visibility that needs your screenshots). The tool writes **"Unknown — need user confirmation"** and leaves the score blank rather than guessing. Fill those cells in yourself.

## One-time template setup — done for you

The scoring sheet's 4 extra rows per competitor block (**Review Velocity, Product Variations, Sales / Revenue Strength, Enhanced Content (A+/Video)**, after "Pricing Strategy") are already in **`test_file/competitor_analysis/Compatitor Analysis.xlsx`**, which the tool uses by default. All 21 grading elements get filled.

## FAQ

**"KEEPA_API_KEY not found"** — Add `KEEPA_API_KEY=...` to your `.env` (same file as the Anthropic key). Phase 3 pulls Keepa live and won't guess that data.

**Why only the top 3 get Keepa data?** Keepa charges tokens per product. Picking the 3 from the free JS/Xray data first, then pulling Keepa only for those 3, keeps your token use to a handful per product.

**Do I still upload a Keepa file?** No — the tool calls the Keepa API itself. You only upload the Jungle Scout CSV and the Xray product exports.

# Extra Tool — Your Own Jungle Scout Sheet (standalone)

Two standalone helpers that together replace the Jungle Scout keyword export — no Jungle Scout subscription needed, only your Keepa key.

**One command does it all** (from the project folder):

```
venv\Scripts\python.exe -m phase3.js_sheet "tent pegs"
```

It scrapes **page 1 of Amazon search results** for your keyword, pulls every product through the **Keepa API**, and writes a **CSV** to `output/` with the familiar JS columns: ASIN, name, brand, price, monthly/daily units, monthly revenue, net revenue, rating, reviews, fees, BSR, listing quality (our own estimate), fulfilment, sellers, category, size tier, dimensions, weight, link — plus search position and a sponsored flag. Averages (price, sales, reviews, BSR, rating) are printed on screen.

**Costs & honesty notes:**
- Keepa charges about **2 tokens per product** (~100 tokens for a full page 1). Results are cached, so re-runs are free. Use `--limit 20` to cap spending per run — the next run resumes where you left off.
- **Monthly Units** comes from Amazon's own "bought in past month" badge (a rounded floor like 600+), not Jungle Scout's estimate — so the numbers will read lower than JS. The `Sales Source` column tells you where each number came from; missing data stays blank, never guessed.

You can also run the ASIN scraper alone: `venv\Scripts\python.exe -m phase3.asin_scraper "tent pegs" --csv output\asins.csv` (add `--max-pages 7` for deeper pages), then feed it back with `--asins-file`.

> These same tools now run **automatically inside the pipeline**: Phase 2 starts generating the sheet the moment Claude extracts your product's main search term, and Phase 3 picks it up — see "Phase 3 — Before You Start". If auto-generation ever fails (e.g. Amazon blocks the scrape), Phase 3 asks you for a file path instead — there you can also type `kw:<your keyword>` to retry generation.
