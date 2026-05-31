# PL Research Tool — Phase 3 Handoff

> Context file for starting **Phase 3 (Competitor Identification + Analysis)**.
> Phases 1 and 2 are built. Read this first, then `workflow_programmer.md`
> (technical spec) and `workflow_client.md` (user guide). The full multi-phase
> blueprint is `PL_Tool_Complete_Workflow_v9_6_Final (3).pdf` (Phase 3 =
> Steps 7–11). Phase 2 background is in `PHASE2_HANDOFF.md`.

---

## 1. Where the project stands

- **Phase 1 — Product Hunting — DONE.** Reads an H10 Black Box XLSX, filters,
  scores, splits by variation count, exports ranked Excel shortlists. User picks
  one **target ASIN**.
- **Phase 2 — Keyword Verification — FUNCTIONALLY COMPLETE.** From the target
  ASIN + an **H10 Xray keyword export**, Claude builds a product profile and
  selects the launch keywords; the user approves/swaps until the set is
  **locked**. `run_phase2()` **returns** the locked keywords. (Reader column
  mappings are verified against a real Xray export; remaining: a first live API
  run.)
- **Phase 3 — Competitor Identification + Analysis — NOT STARTED.** This is the
  next build. Phases 4–7 remain PDF-only; do not start them.

---

## 2. How to run (unchanged)

```powershell
.\setup.ps1     # one-time: venv + pip install (now includes anthropic, python-dotenv)
.\run.ps1       # always starts Phase 1, then flows into Phase 2 automatically
```

- venv at `.\venv\` (Python 3.10+). Deps: `pandas`, `openpyxl`, `numpy`,
  `Pillow`, `anthropic`, `python-dotenv`.
- Claude API key: git-ignored **`.env`** in project root (`ANTHROPIC_API_KEY=...`).
- Outputs land in `output\` (git-ignored). Sample input in `test_file\`.

---

## 3. Phase 1 + 2 combined state (conventions Phase 3 should follow)

```
phase1/  file_reader.py · filters.py · scorer.py · output.py · main.py
phase2/  xray_reader.py · keyword_selector.py · claude_client.py · main.py
run.py   always Phase 1 -> Phase 2 (no menu)
```

**Established conventions to reuse in Phase 3:**
- **Pure deterministic Python for data ops; Claude (API) for all judgement.**
  Relevancy / classification / "same product type" decisions go to Claude — no
  heuristic fallback (Rule 0: never guess; mark unknown / surface to user).
- **Column normalisation:** every tool export is read through a `COLUMN_MAPPINGS`
  dict with broad alias fallback + numeric parsing (strip currency/commas, H10
  `-` -> None). See `phase1/file_reader.py` and `phase2/xray_reader.py`. New
  Phase 3 readers (JS / Keepa / Xray) need the same treatment.
- **Claude client pattern** (`phase2/claude_client.py`): `.env` key via
  `python-dotenv`; model `claude-opus-4-8`; adaptive thinking; structured
  outputs via `client.messages.parse(output_format=PydanticModel)`; prompt
  caching on the stable context prefix; `Phase2ApiError`-style surfaced errors,
  no fallback.
- **Excel output** via `openpyxl` reusing the palette/helpers in
  `phase1/output.py` (styled header, frozen panes, embedded images, timestamped
  filenames, never overwritten).
- Interactive CLI (stdin/stdout); Windows/PowerShell launch.
- Keep **both** `workflow_client.md` and `workflow_programmer.md` in sync with
  every code change.

**Data available to hand into Phase 3:**
- The **target ASIN** (typed at the start of Phase 2).
- The **locked launch keywords** — `run_phase2()` returns a list of
  `SelectedKeyword` (`keyword`, `search_volume`, `relevancy`, `reason`). NOTE:
  `phase1/main.py` currently calls `run_phase2(blackbox_df=df)` and ignores the
  return. When wiring Phase 3, capture that return (and the target ASIN) so they
  flow forward — the keywords drive competitor keyword-ranking relevancy and the
  later keyword sheets.
- The in-memory **Black Box DataFrame** (already passed P1->P2).

---

## 4. What Phase 3 must do (PDF Steps 7–11)

**Phase 3 — Competitor Identification + Analysis.** Inputs: the target ASIN +
locked keywords from Phase 2, plus three new uploads.

1. **Step 7 (USER, squeezed):** run **Jungle Scout** (keyword search, CSV),
   **Keepa** (bulk ASIN lookup, XLSX), and **H10 Xray** (same keyword on Amazon
   UK, XLSX) — upload all three together.
2. **Step 8 (TOOL):** pick the **top 3 competitors**. Include **both FBA and
   FBM**. Apply a **same-product-type filter** across all three sources
   (judgement → Claude). Score each ASIN by **keyword-ranking relevancy +
   monthly sales**; select the top 3 by sales volume within the same product
   type. Output the 3 with FBA/FBM flag, full data, and a reason each.
3. **Step 9:** competitor **review data** (1–3★ and 4–5★ per competitor).
   Target path is the **FlyByAPIs** review API (ASIN + star filter -> JSON);
   current fallback is user-uploaded screenshot PDFs.
4. **Step 10 (TOOL, one pass):** fill the competitor sheets —
   - 10A Keepa **Competitor Analysis** tab: every Col F field populated (brand,
     URL, seller, price, BSR, fees, images, title, all bullets, A+ flag,
     FBA/FBM, listing age, rating, review count, dimensions, weight).
   - 10B AMS ad visibility (user screenshots — cannot be automated).
   - 10C full **rubric scoring** (max 159/competitor) + top-3 weaknesses (from
     1–3★) + top-3 strengths (from 4–5★).
   - 10D Sponsored Products (6 keywords × Exact/Phrase/Broad; PPC bids blank);
     **Pricing Analysis** product name always updated; BSR/price/fees from
     Keepa + JS; COGS/shipping blank until sourcing.
   - 10E **Deliverables:** Competitor Analysis Excel (all tabs, Col F full) +
     Competitor PDF (3 scored, gap analysis, market-entry verdict).
5. **Step 11 — DECISION GATE (mandatory pause):** user reviews the deliverables.
   Option A workable -> Phase 4; B change competitors -> re-pick (loop to Step 8);
   C not workable -> back to Phase 1.

---

## 5. Open decisions to confirm at the start of Phase 3

1. **Module shape:** a `phase3/` package mirroring prior phases —
   `js_reader.py`, `keepa_reader.py`, `xray_product_reader.py`,
   `competitor_selector.py` (Claude same-product-type + scoring), `output.py`,
   `main.py`? Recommend mirroring. (NB: Phase 3's Xray export is the
   product/listing search — distinct from Phase 2's `xray_reader.py`, which
   reads the Xray *Keywords* export. Name the Phase 3 one accordingly.)
2. **Real exports needed first:** capture actual **Jungle Scout CSV**, **Keepa
   XLSX**, and **Xray product XLSX** headers to build their `COLUMN_MAPPINGS`
   (the same way `phase2/xray_reader.py` was confirmed against a real Xray
   keyword export).
3. **The PES Excel template:** Step 10 fills a specific multi-tab Competitor
   Analysis workbook. The real template is required to map exact cells/tabs;
   without it, only the data model can be built. Obtain the template.
4. **Review data source:** FlyByAPIs API now, or start with screenshot-PDF
   ingestion? (Affects Step 9/10C.)
5. **Automation split for Step 8:** deterministic = merge/dedupe across 3
   sources, sales-volume ranking; Claude = same-product-type judgement + the
   reason text. Confirm how much to route to Claude vs leave deterministic.
6. **Deliverables:** confirm the two outputs (Competitor Analysis Excel +
   Competitor PDF) and how the PDF is generated.

---

## 6. Standing rules for this project

- Pure-deterministic data ops; **all judgement via Claude API, no heuristic
  fallback** (Rule 0 — never guess; mark unknown / ask).
- Keep **both** workflow docs in sync with every code change.
- Don't start Phase 4+ until Phase 3 is approved complete.
- Trademark / patent checks remain the user's responsibility.

---

## 7. First message to send in the new chat

> "Starting Phase 3 (Competitor Identification + Analysis). Read
> PHASE3_HANDOFF.md, workflow_programmer.md, and workflow_client.md. I'll
> provide sample Jungle Scout, Keepa, and Xray exports plus the PES Competitor
> Analysis template. Let's inspect their columns and agree on the module
> structure before writing code."
