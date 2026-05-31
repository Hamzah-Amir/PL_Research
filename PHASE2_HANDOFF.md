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

**Phase 2 — Keyword Verification — FUNCTIONALLY COMPLETE.** Steps 4–6 are built:
Claude proposes the launch keywords (Step 5) and the user approves/swaps until
the set is locked (Step 6). The locked set is printed and **returned** from
`run_phase2()` for Phase 3 — no Excel export (the keywords feed Phase 3's
competitor analysis sheet). Remaining: a first live API run + verifying the
Cerebro column mappings against a real export.

Phases 3–7 exist only in the PDF blueprint; do not start them.

### Phase 2 resolved decisions (so future work matches)
- **Inputs:** target ASIN (typed) + Phase 1 Black Box XLSX (for the title) +
  H10 **Cerebro** export. Cerebro only — Keepa/JS are Phase 3.
- **All keyword judgement is purely Claude via API** — relevancy + every Step-5
  exclusion. **No deterministic relevancy logic, no heuristic fallback.** Only
  objective prep (dedupe + SV ≥ 100) is deterministic.
- **Product context:** ASIN → look up `{title, category, subcategory}` in Black
  Box → **Claude builds a structured product profile** (`derive_product_profile`):
  `product_name`, `product_type`, `key_attributes`, `use_cases`, `not_this`,
  `brand`. That profile (not a lossy short name) anchors the relevancy judgement
  (`select_keywords`). **Text only — the product image is NOT sent** (decided
  against sending the image).
- **API key:** git-ignored `.env` in project root (`ANTHROPIC_API_KEY`), loaded
  via `python-dotenv`. Model: `claude-opus-4-8`, adaptive thinking, structured
  outputs (`messages.parse`), prompt caching on the product + candidate prefix.
- **Modules:** `phase2/cerebro_reader.py`, `keyword_selector.py`,
  `claude_client.py`, `main.py`.
- **Phase 1 → Phase 2 connected (automatic):** `run.py`/`run.ps1` always start
  at Phase 1 (no menu). Phase 1 ends by calling `run_phase2(blackbox_df=df)`
  directly (no confirm) with the loaded Black Box frame — Phase 2 reuses it and
  does not re-prompt for the file. Press Enter at the Phase 2 ASIN prompt to
  exit if Phase 2 isn't wanted. For isolated dev testing, `python -m phase2.main`
  runs `run_phase2()` (no arg), which prompts for the Black Box path.
- **Step 6 loop:** `_run_approval_loop` keeps approved keywords, excludes
  rejected ones, and re-calls `select_keywords(target_count=need, exclude=...)`
  for replacements until the user locks the set.

### Remaining for Phase 2
1. Verify `cerebro_reader.COLUMN_MAPPINGS` against a **real** Cerebro export.
2. First live end-to-end API run (the two Claude calls are unexercised against a real key/files).

(No Excel export — locked keywords are returned from `run_phase2()` and consumed by Phase 3.)

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
