#!/usr/bin/env python3
import json, re
from pathlib import Path
import pandas as pd

p = Path('.').resolve()
cache = p / 'Input' / 'DB_Cache' / 'mpdt_mapping_cache.json'
if not cache.exists():
    print('mapping cache not found:', cache)
    raise SystemExit(1)

payload = json.loads(cache.read_text(encoding='utf-8'))
mapping = payload.get('mapping', {}) if isinstance(payload, dict) else {}
cols = set()
for expr in mapping.values():
    if isinstance(expr, str):
        for m in re.finditer(r"join[12]\s*\[\s*['\"]?([^'\"]+)['\"]?\s*\]", expr, re.IGNORECASE):
            cols.add(m.group(1).strip())

print('discovered_columns_count=', len(cols))
print(sorted(list(cols))[:200])

# Read SmartForms sheet headers only
sf_file = p / 'Input' / 'SmartForms_RAW_MPDT_L2&L3.xlsx'
if not sf_file.exists():
    print('SmartForms fallback Excel not found:', sf_file)
    raise SystemExit(1)

xl = pd.ExcelFile(sf_file, engine='openpyxl')
l2_sheet = next((s for s in xl.sheet_names if 'l2' in s.lower()), xl.sheet_names[0])
l3_sheet = next((s for s in xl.sheet_names if 'l3' in s.lower()), xl.sheet_names[-1])

l2_cols = list(pd.read_excel(sf_file, sheet_name=l2_sheet, nrows=0, engine='openpyxl').columns)
l3_cols = list(pd.read_excel(sf_file, sheet_name=l3_sheet, nrows=0, engine='openpyxl').columns)

print('l2_header_count=', len(l2_cols), 'l3_header_count=', len(l3_cols))
print('l2_example=', l2_cols[:40])
print('l3_example=', l3_cols[:40])

# Show intersection of discovered cols with sheet headers (case-insensitive)
lc = {c.lower(): c for c in l2_cols}
rc = {c.lower(): c for c in l3_cols}
matched_l2 = [lc[c.lower()] for c in cols if c.lower() in lc]
matched_l3 = [rc[c.lower()] for c in cols if c.lower() in rc]
print('matched_in_l2=', matched_l2[:200])
print('matched_in_l3=', matched_l3[:200])
