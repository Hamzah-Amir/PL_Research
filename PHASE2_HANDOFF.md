# PL Research Tool — Phase 2 Handoff

> Context file for starting **Phase 2 (Keyword Verification)** in a new chat.
> Phase 1 is complete and frozen. Read this first, then read
> `workflow_programmer.md` (technical spec) and `workflow_client.md` (user guide).
> The full multi-phase blueprint is `PL_Tool_Complete_Workflow_v9_6_Final (3).pdf`.

---

## 1. Where the project stands

**Phase 1 — Product Hunting — DONE.** A Python CLI that reads a Helium 10
Black Box XLSX export, filters it, scores every survivor, and outputs ranked
Excel shortlists. The end product of Phase 1 is the user **picking one target
ASIN** from the output to carry into Phase 2.

**Phase 2 — Keyword Verification — NOT STARTED.** This is the next build.

Phases 3–7 exist only in the PDF blueprint; do not start them.

---

## 2. How to run (unchanged for Phase 2)

```powershell
.\setup.ps1     # one-time: creates venv/, pip installs deps
.\run.ps1       # runs phase1.main.run_phase1()
```

- venv lives at `.\venv\` (Python 3.10+). Deps: `pandas`, `openpyxl`, `numpy`, `Pillow`.
- Sample input: `test_file\GB_AMAZON_blackBoxProducts_1_2026-05-30.xlsx` (500 products, Sports & Outdoors, Amazon UK).
- Outputs land in `output\`.

---

## 3. Phase 1 final state (so Phase 2 matches conventions)

```
phase1/
├── file_reader.py   # XLSX → normalised DataFrame (COLUMN_MAPPINGS, numeric parsing)
├── filters.py       # FilterConfig + apply_filters (F1–F9; no sales floor)
├── scorer.py        # score_all_products / select_top_products
├── output.py        # export_to_excel (embedded images), print_product_table
└── main.py          # CLI orchestration + batch loop
```

Key conventions Phase 2 should follow:
- **Pure deterministic Python**, no LLM calls yet. (The user plans to add an
  Anthropic API key later for the judgement-heavy steps; build the structural
  pipeline first so it can be swapped in.)
- Interactive CLI (stdin/stdout), Windows/PowerShell launch scripts.
- H10 column-name variants are normalised via a `COLUMN_MAPPINGS` dict in the
  reader — Cerebro exports will need the same treatment.
- Excel output via `openpyxl` with a styled header, frozen panes, and a
  settings/reference sheet. Reuse the palette + helpers in `output.py`.
- Output files are timestamped and never overwritten.

Phase 1 scoring (for reference, not reused in P2): weights
revenue 35 / opportunity 25 / sales 20 / rating 12 / review 8; log-normalised
revenue & sales; review-count bracket; seller-concentration and 90-day-trend
multipliers; clamp at 100. Each batch outputs two Excel files
(With Variants / Without Variants).

---

## 4. What Phase 2 must do (from the workflow PDF)

**Phase 2 — Keyword Verification.** Input is the **single target ASIN** chosen
at the end of Phase 1.

1. **Step 4 (USER):** Run H10 **Cerebro** on the target ASIN → export the full
   keyword XLSX → provide it to the tool.
2. **Step 5 (TOOL):** Select the **top 6 keywords** by highest relevancy +
   highest search volume. **Exclude:**
   - branded competitor names
   - misspellings
   - wrong-category keywords
   - overly broad keywords
   - question keywords
   - search volume < 100
   - keywords describing a *different* product type
   - Output the 6 with SV, relevancy score, and a reason per pick.
3. **Step 6 (USER):** User approves the 6, or requests swaps; tool substitutes
   rejected keywords until all 6 are approved. **Final 6 keywords are locked**
   before Phase 3.

> Note: Steps 5's exclusion rules ("wrong category", "different product type")
> are judgement calls the PDF assigns to Claude. Without an API key, Phase 2 can
> do the deterministic parts (SV ≥ 100, dedupe, sort by SV × relevancy, strip
> obvious question words / branded terms via a list) and surface the rest for
> the user to confirm. Decide with the user how much to automate now vs. defer
> to the API integration.

---

## 5. Open decisions to confirm at the start of Phase 2

1. **Module shape:** create a `phase2/` package mirroring `phase1/`
   (`cerebro_reader.py`, `keyword_selector.py`, `output.py`, `main.py`)? Or a
   single module? Recommend mirroring Phase 1.
2. **Input handoff:** how does Phase 2 receive the target ASIN — typed in by the
   user, or read from a Phase 1 output file? (Phase 1 currently does not persist
   a "selected ASIN" anywhere.)
3. **Cerebro column names:** the exact headers in an H10 Cerebro export are not
   yet captured in the codebase. First task in Phase 2 is to inspect a real
   Cerebro XLSX and build its `COLUMN_MAPPINGS` (Keyword, Search Volume,
   relevancy/Cerebro IQ score, etc.).
4. **Automation level:** how much of the Step-5 keyword judgement to hardcode vs.
   leave for the user / future Anthropic API call.
5. **Output:** Excel shortlist of 6 keywords (matching Phase 1 styling) plus the
   reason column? Confirm format.

---

## 6. Standing rules for this project

- Keep **both** `workflow_client.md` and `workflow_programmer.md` in sync with
  every code change (this rule is already in memory).
- Don't start Phase 3+ until Phase 2 is approved complete.
- Trademark / patent checks on the chosen product remain the user's
  responsibility — the tool does not do them.

---

## 7. First message to send in the new chat

> "Starting Phase 2 (Keyword Verification). Read PHASE2_HANDOFF.md,
> workflow_programmer.md, and workflow_client.md. I'll provide a sample H10
> Cerebro keyword export. Let's first inspect its columns and agree on the
> module structure before writing code."
