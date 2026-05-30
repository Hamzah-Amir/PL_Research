# PL Research Tool — Client Guide

This is Phase 1 of a multi-phase Amazon UK Private Label research tool. It takes a Helium 10 Black Box XLSX export, filters out poor candidates automatically, scores everything that survives, and gives you the best products in ranked batches — with product photos embedded in the Excel output. The single ASIN you pick here feeds into the next phases.

---

## Before You Start

| Item | Source |
|------|--------|
| Launch budget (£) | You decide |
| H10 Black Box XLSX | Helium 10 → Black Box → Products tab, **one category only** |

> Export one category per file. Include the **Image URL** column so product photos can be embedded.

---

## How to Run

**First-time setup (run once):**
```
.\setup.ps1
```
**Every run:**
```
.\run.ps1
```

---

## What the Tool Asks

**1. Budget (£)**
Sets the revenue cap: `Budget × 1.52`. Products earning more than the cap are too competitive and removed. Both revenue and monthly sales are taken from the **ASIN-level** columns in your export. There is no minimum revenue.

**2. Seasonal products?** (Y / Enter = No)
- **No (default):** all seasonal products removed.
- **Yes:** summer/unknown products are kept and flagged with a note explaining why. Christmas/Q4 products (Best Sales Period Oct–Dec, or titles containing christmas/xmas/halloween/easter/valentine) are **always removed** regardless.

**3. Category** (optional)
Press Enter to auto-detect from the file.

**4. File path**
Paste the full path to your XLSX export.

---

## Filters (all automatic)

You do not set these — they run in the background with proven values.

| Filter | Rule |
|--------|------|
| Revenue cap | ASIN-level revenue ≤ budget × 1.52 |
| Amazon as seller | Always removed |
| Seasonal products | See above |
| Listing age | ≥ 6 months |
| Star rating | > 3.8 (strict — 3.8 itself is removed) |
| Unit price | ≥ £5 (no upper limit) |
| 90-day sales trend | Removed if decline > 30% |
| Review count | ≤ 3,000 |
| Regulated categories | None excluded by default |
| Compatibility products | Always removed (e.g. "Case for iPhone 15") |

---

## Results

After filtering, every surviving product is scored and then **split into two groups by variation count** (the "Variation Count" column in your export — 0 means the product has no variants):

- **Top 10 With Variants** — products that have colour/size/pack variations
- **Top 10 Without Variants** — single-variant products

Each batch produces **two Excel files in one go**, both saved to `output/`:

- `Phase1_Batch1_WithVariants_<timestamp>.xlsx`
- `Phase1_Batch1_NoVariants_<timestamp>.xlsx`

If more products remain in either group you'll be asked **"Would you like to see the next 10 of each?"** — press Enter or N to exit, Y to continue. Each "yes" produces the next batch of both files (Batch 2, Batch 3, …).

### Scoring (0–100, higher is better)

| Signal | Weight |
|--------|--------|
| Monthly revenue | 35% |
| Sales-to-reviews ratio (market gap) | 25% |
| Monthly unit sales | 20% |
| Star rating | 12% |
| Review count (fewer is better) | 8% |

Revenue and sales are scored on a log scale so genuinely higher-earning products rise to the top instead of bunching together. The base score is then nudged by two factors: products with **fewer active sellers** get a small boost (less competition), and products with **rising 90-day sales** get a boost while sharply declining ones are penalised. Final scores are capped at 100.

Review count rewards under-served products: ≤50 reviews scores highest, anything over 1,000 scores lowest.

Score colours: **Green ≥ 70** (strong), **Amber 50–69** (moderate), **Red < 50** (weaker).

### Risk Flags (Notes column)

Products that pass all filters but still carry a warning get a note. Yellow-highlighted Notes cell = flag present.

| Flag | What it means |
|------|--------------|
| FBM seller | Competitor ships themselves, not via Amazon warehouse |
| Low reviews (<50) | Very thin social proof |
| High reviews (>1000) | Saturated — hard for a new listing to compete |
| Low rating (<4.0) | Just above the filter threshold |
| Young listing (<12 mo) | Less than a year of sales history |
| Declining trend | Sales dropping more than 10% over 90 days |
| Single seller | Only one active seller on the listing |
| Seasonal — reason | Kept seasonal product (if you chose Yes) |

---

## What to Do with the Results

1. Open the Excel file(s) in `output/`
2. Look at the green-scored products first; use the photos and Notes column to sense-check each one
3. **Select exactly one ASIN** to carry forward to Phase 2
4. Before selecting: check for trademark conflicts and active patents — the tool does not do this

---

## FAQ

**"No products passed all filters"** — Re-run with a higher budget (raises the cap), say Yes to seasonal products, or use a file covering a broader category.

**"no image" in the image cell** — Image URL was missing from your export or couldn't be downloaded. Make sure the Image URL column is included in your Black Box export.

**Can I re-run on the same file?** Yes. Each run saves timestamped files (e.g. `Phase1_Batch1_WithVariants_20260529_143012.xlsx`) — nothing is overwritten.

**A product scored red — should I ignore it?** The score is a ranking tool, not a pass/fail gate. Red means it ranked lower than the others in your filtered pool, not that it's bad.
