# UpdatePW_ACBOS_MPDT_V2

Current implemented module: `step2b` (local mode).

## Run

```powershell
cd "C:\Users\seethapn1094\OneDrive - ARCADIS\Documents\DataScience\HS2_L3Database\UpdatePW_ACBOS_MPDT_V2"

# Full pipeline — always refreshes local caches from source Excel/DB files first
python main.py --step all --target-uaid2 HS2-000001403

# Partial re-run using existing cache (fast, no re-fetch)
python main.py --step step4 --target-uaid2 HS2-000001403

# Force re-fetch even on partial steps
python main.py --step step3 --refresh-sources --target-uaid2 HS2-000001403
```

## Data sourcing

| Mode | When | Behaviour |
|---|---|---|
| `--step all` | Full pipeline | Step 1 always re-reads source Excel files and re-caches. Tries DB/API first; falls back to local Excel if unavailable. |
| `--step stepN` | Partial run | Uses existing cache. Pass `--refresh-sources` to force re-fetch. |
| `--refresh-sources` | Any step | Forces re-fetch from DB/API and rebuilds cache. |

## Current behavior

- Uses `config/pipeline_config.json`
- Loads PW input from configured local file (`paths.pw_input`)
- Keeps selected PW columns
- Guarantees both `UAID_2` and `ASSET_ID` columns are present
- Adds reference flags:
  - `In_Works_Tracker`
  - `In_L2_UAID_ACBOS`
- Writes single sheet output to `Output/asset_deliverables.xlsx`

## Not yet implemented

- `data_source_mode = source` connectors
- Classifier module
- MPDT generator module
