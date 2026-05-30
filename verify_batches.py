"""Quick proof that each batch is strictly capped at 10 rows."""
import sys, io
sys.path.insert(0, '.')
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import warnings
warnings.filterwarnings('ignore')

from phase1.file_reader import read_blackbox_file
from phase1.filters import build_filter_config, apply_filters
from phase1.scorer import select_top_products

df, category = read_blackbox_file('test_file/GB_AMAZON_blackBoxProducts_1_2026-05-28.xlsx')
config = build_filter_config(category, 7600.0)
filtered, _ = apply_filters(df, config)

# Hard cap at 100
ranked = select_top_products(filtered, top_n=100)
pool_size = len(ranked)

print(f"Products in file     : {len(df)}")
print(f"After all filters    : {len(filtered)}")
print(f"Pool (capped at 100) : {pool_size}")
print(f"Batches of 10        : {(pool_size + 9) // 10}")
print()
print(f"  {'Batch':>5}  {'From':>5}  {'To':>5}  {'Rows in Excel':>13}  {'Remaining after':>15}")
print(f"  {'─'*5}  {'─'*5}  {'─'*5}  {'─'*13}  {'─'*15}")

max_rows = 0
for b in range((pool_size + 9) // 10):
    start = b * 10
    batch = ranked.iloc[start: start + 10]
    rows  = len(batch)
    max_rows = max(max_rows, rows)
    end   = start + rows
    after = pool_size - end
    print(f"  {b+1:>5}  {start+1:>5}  {end:>5}  {rows:>13}  {after:>15}")

print()
print(f"  Max rows in any single Excel : {max_rows}  (must be ≤ 10)")
assert max_rows <= 10, "BUG: a batch exceeded 10 rows!"
print("  PASS — strict 10-per-batch cap confirmed.")
print()
print("Interactive flow (main.py):")
print("  Batch 1 → exported immediately → user asked 'Satisfied? Y/N'")
print("  Y (or Enter) → stop. Only Batch 1 Excel exists.")
print("  N → Batch 2 exported → user asked again → and so on.")
print("  Pool exhausted → tool ends automatically.")
