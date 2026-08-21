"""
Run: python debug_banner.py <path_to_banner.xlsx>
Prints column detection and sheet structure for diagnosis.
"""
import sys, openpyxl
sys.path.insert(0, '/tmp/crosstab-studio')
import engine as _engine
import manifest_builder as mb

if len(sys.argv) < 2:
    print("Usage: python debug_banner.py <banner.xlsx>")
    sys.exit(1)

path = sys.argv[1]
profile_name = "Corporate Reputation"

with open(path, 'rb') as f:
    file_bytes = f.read()

print("=" * 60)
print("COLUMN DETECTION")
print("=" * 60)
cols = _engine.get_columns(file_bytes, profile_name)
for idx, name, sub in cols:
    print(f"  col {idx:3d}: '{name}'  sublabel='{sub}'")

print()
print("=" * 60)
print("AUDIENCE MATCHING (for manifest audience names)")
print("=" * 60)
for aud in ["Gen Pop", "Tech Elites", "Tech Elite", "AI Fans", "AI Fan", "Gen Z"]:
    matched = mb._match_col_for_audience(cols, aud)
    print(f"  '{aud}' → col {matched}")

print()
print("=" * 60)
print("SHEET INVENTORY (first 30 sheets with Grid_GO16 or first 20 overall)")
print("=" * 60)
wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
for i, name in enumerate(wb.sheetnames):
    if 'GO16' in name or i < 20:
        print(f"  [{i:4d}] {name}")
    if i > 200 and 'GO16' not in name:
        break
wb.close()

print()
print("=" * 60)
print("BANNER CACHE — entries for Grid_GO16")
print("=" * 60)
cache = mb.BannerCache(file_bytes, profile_name, list(range(5)))  # first 5 col indices
entries = cache.entries_for("Grid_GO16")
for si, word, stype in entries[:20]:
    print(f"  sheet {si:4d}: stype='{stype}'  wording='{word[:80]}'")

if entries:
    print()
    print("=" * 60)
    print("PARSED CONTENT — first Grid_GO16 sheet (first 5 rows)")
    print("=" * 60)
    p = cache.parsed_sheet(entries[0][0])
    if p:
        print(f"  answers[:5]: {p.get('answers', [])[:5]}")
        print(f"  values[:2]:  {p.get('values', [])[:2]}")
    else:
        print("  (parse returned None)")
