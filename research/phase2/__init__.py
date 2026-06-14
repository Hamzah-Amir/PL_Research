"""Phase 2 — Keyword Verification.

Input is the single target ASIN chosen at the end of Phase 1 plus an H10
Cerebro keyword export for that ASIN. Phase 2 reads the Cerebro XLSX, prepares
a deterministic candidate keyword pool, and (once the Anthropic API is wired
in) selects the final 6 launch keywords.

Phase 1's conventions are mirrored throughout: canonical column normalisation,
openpyxl Excel output, interactive CLI, timestamped output files.
"""
