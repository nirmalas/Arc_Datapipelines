HS2 MPDT/ACBOS Pipeline — README

Overview

This repository implements a pipeline to build MPDT (Master Project Data Table) files and ACBOS deliverables, consolidate tracker data, and (optionally) upload/download MPDTs to/from ProjectWise.

This README documents:
- Configuration options in `config/pipeline_config.json`
- Expected input files and their purpose
- End-to-end step-by-step run instructions
- How MPDT files are created (template, LoDM, attribute applicability, BO+ trimming)
- How ACBOS deliverables are produced and how filename mapping works
- ProjectWise download / upload scripts and credential notes
- Troubleshooting and common checks

Prerequisites

- Python 3.10+ (this repo used Python 3.12). Ensure required packages from `requirement_plan_updated.txt` are installed into your chosen venv.
- PowerShell (pwsh) for running the ProjectWise helper scripts.
- Optional: `pwps_dab` and ProjectWise client access if you want automated downloads/uploads.
- Files in the `Input/` folder (see below).

Configuration (`config/pipeline_config.json`)

Key fields and meaning (most commonly used keys):

- `output_directory`: Directory where generated outputs are written (default `Output`).
- `log_level`: Logging level (INFO/DEBUG).
- `batch_size`: How many UAIDs to process per batch when running in batch mode.
- `target_uaid2`: Array of UAID_2 values (strings) to process. CLI `--target-uaid2` overrides this.
- `file_type`: `AUTO` | `MPDT` | `ACBOS` — controls output file type detection/fallback.
- `missing_in_pw_file_type`: What file type to assume when a UAID is not present in ProjectWise metadata (usually `MPDT`).
- `fetch_external`: Flags to fetch fresh external sources automatically (db, smartforms, pw_extract). When `true` these can overwrite `Input/` files.
- `allow_scope3_fallback_when_pw_missing`: boolean, when ProjectWise extract has no record for a UAID, whether to fall back to DB-scope3 data for attribute values.
- `uaid_type_map`: map UAID -> preferred file type (optional overrides).
- `not_in_pw_type_map`: map of UAIDs that aren't in PW, used for test runs.
- `db`: connection info if `fetch_external.db=true` (ODBC connection string and table names).
- `smartforms_api`: SmartForms API credentials/config (if fetching from Bentley API).
- `pw`: ProjectWise settings:
  - `datasource`: ProjectWise datasource string used by PS scripts.
  - `username`: automation username (if used by PS scripts).
  - `ps1_download`: path to the PS script used to download latest MPDT (`Scripts/PWPS_Download_MPDT.ps1`).
  - `ps1_extract`: path to the PS script used to extract ProjectWise metadata (`Scripts/PWPS_Data_Extract.ps1`).
  - `ps1_upload`: path to the PS script used to upload generated MPDTs (`Scripts/PWPS_Upload_PW.ps1`).
- `paths`: local file paths used by the pipeline (relative to workspace). Important keys:
  - `pw_extract`: `Input/ACBOS MPDT.xlsx` (ProjectWise metadata extract).
  - `pw_extract_full`: `Input/ACBOS MPDT_FULL.xlsx` (full extract history, optional).
  - `l3_assets_scope_data`: cached DB scope data (`Input/l3_assets_scope_data.xlsx`).
  - `l2_uaid_acbos`: L2 mapping file `Input/L2 UAID-ACBOS_260129.xlsx`.
  - `lodm`: LoDM master file `Input/1MC06-..._lodm.xlsx` (CURRENT LoDM sheet used by generator).
  - `control_file`: control file used for mapping and defaults.
  - `mpdt_template`: MPDT template workbook `Input/C2-MPDT-Template-Mapping.xlsm`.
  - `sample_mpdt`: an example MPDT used to infer column layout.
  - `target_uaid2_file`: path to an optional spreadsheet providing the UAID_2 list (e.g., `Input/L2 UAID_Planting.xlsx`).
  - tracker inputs: `tracker_sep2025`, `tracker_2025`, `tracker_2024` (these are consolidated by the pipeline).

Input Files (what they contain)

- `Input/ACBOS MPDT.xlsx` — ProjectWise metadata extract (document names, paths, UAID columns). Used to determine latest existing MPDTs and some override values.
- `Input/ACBOS MPDT_FULL.xlsx` — full history extract (optional).
- `Input/SmartForms_RAW_MPDT_L2&L3.xlsx` — SmartForms raw export (if using SmartForms fallback).
- `Input/L2 UAID-ACBOS_260129.xlsx` — mapping between UAID_2 and ACBOS template names or stems.
- `Input/1MC06-..._lodm.xlsx` — LoDM sheet, used to decide which AU columns are applicable and counts for BO+.
- `Input/C2-MPDT-Template-Mapping.xlsm` — MPDT template with columns and formatting used when creating MPDT per UAID.
- `Input/1MC07-...xlsm` (sample MPDT) — example MPDT used to extract header columns and column indexes.
- `Input/JSON Works Tracker_*.xlsx` — multiple tracker sheets that are consolidated into `works_tracker.xlsx` and `consolidated_works_tracker.xlsx` during pipeline runs.
- `Input/l3_assets_scope_data.xlsx`, `Input/DB_Cache/*` — scope2 and scope3 data exports (DB) used for attribute fallback and OSGB coordinates.

High-level Steps (what each main script/step does)

The pipeline is orchestrated by `main.py` which exposes steps; run with `python main.py --step stepN` or use the CLI flags.

Typical end-to-end flow:
1. Step 1 — (optional) fetch sources
   - Reads/refreshes `Input/` files according to `fetch_external` flags. May run `Scripts/PWPS_Data_Extract.ps1` if `pw_extract=true`.
2. Step 2 — Build `asset_deliverables` (ACBOS mapping)
   - Uses L2 mapping and classifiers to assign each Asset a deliverable filename candidate.
   - Writes `Output/asset_deliverables.xlsx`.
3. Step 3 — Classify targets / Consolidate trackers
   - Runs classifier logic to choose final deliverable filename using priority rules: exact `asset_id` match, then `tb_level_2_id` tokenized match, then exact `description` match, then fallback to consolidated works tracker.
   - Writes `Output/classification_output.xlsx` and stamps `Deliverable_Mapped_By` into `Output/asset_deliverables.xlsx` (provenance).
4. Step 4 — MPDT generation
   - Builds the MPDT rows from mapping expressions for columns until Column AR, join1/join2 rows (scope2/scope3 and control files), and the MPDT template.
   - The columns after AU should be added based on the att_matrix (row1 =col C, row2 = ColE+' '+ColB) and the Runs round-2 processing to apply LoDM-driven applicability trimming.
   - The columns after AU+ the first 15 columns are retained as permanant.
   - After that all the attributes from att_matric are applied only if they are part of the attributes of the classcode(assethierarchycategory) column from the LoDM file. 
   - The columns after AU+15 should be present only if it is an attribute of atleast one of the UAID_3 in that MPDT output file.
   - If a column attribute is part of a different UAID_3 but not part of this UAID_3, the cell should be blackened. 
   - Optionally downloads the latest existing MPDT from ProjectWise (via `Scripts/PWPS_Download_MPDT.ps1`) in order to source the `Software Model Part ID` or other values.
   - Writes MPDT files under `Output/<timestamp>/MPDT/MPDT_<UAID>.xlsm`.
5. Step 5 — AC BOS generation
   - Produces ACBOS deliverable workbooks (naming parity with `mpdt_creation_v2.ipynb` where appropriate).
6. Step 6 — Upload to ProjectWise (optional)
   - Uses `Scripts/PWPS_Upload_PW.ps1` to upload created MPDTs back to ProjectWise, typically using `pwps_dab` cmdlets.

Detailed: MPDT creation internals

- Column population
  - The generator uses a `mapping_dict` and `populate_row()` to evaluate mapping expressions for each MPDT column using join1/ join2 context rows.
  - Values can come from scope2/scope3, control file, or computed expressions.

- LoDM-driven applicability and BO+ trimming
  - The LoDM (`lodm.xlsx`) is normalized and used to determine which AU (attribute) columns are applicable to a given classcode / UAID combination.
  - If an attribute is not present for the UAID's LoDM class, the corresponding MPDT AU cell is black-filled (or left blank) according to `apply_applicability()` behavior.
  - BO+ trimming: BO+ columns (construct items after BO in the MPDT) are trimmed based on counts inferred from LoDM attribute counts for the class; trailing BO+ columns beyond the determined keep range are removed to avoid empty/incorrect extra columns.

- Software Model Part ID sourcing
  - If a UAID has an existing MPDT document in ProjectWise, the pipeline will attempt to read the latest revision's workbook (local cache or via `Scripts/PWPS_Download_MPDT.ps1`) and extract the `Software Model Part ID` value from the known column (e.g., `Unnamed: 43` or `Software Model Part ID no.` depending on template).
  - If no existing PW MPDT is found, the field is left blank.
  - Note: automated PW download requires `pwps_dab` and appropriate credentials on the machine running the script.

ACBOS and filename mapping rules

Filename mapping priority (implemented in `classifier/type_classifier.py`):
- 1) `asset_id` exact match (highest priority)
- 2) `tb_level_2_id` tokenized match (split by comma, match any token)
- 3) `Asset Description` exact match
- 4) Fallback to consolidated works tracker (consolidates multiple tracker sheets)

The classifier writes `deliverable_mapped_by` for each row describing which rule matched. The `excel_generator/writer.py` stamps that into `Output/asset_deliverables.xlsx`.

ProjectWise integration

Scripts of interest (under `Scripts/`):
- `Scripts/PWPS_Data_Extract.ps1` — extract ProjectWise metadata to `Input/ACBOS MPDT.xlsx` (existing; used by step1 when `pw_extract=true`).
- `Scripts/PWPS_Download_MPDT.ps1` — (new) searches ProjectWise for the latest MPDT document for a UAID and downloads the latest revision to a local cache; the script prints `DownloadedFile=<path>` so the Python generator can pick it up.
- `Scripts/PWPS_Upload_PW.ps1` — uploads a generated MPDT back into ProjectWise (attachments, checkin metadata etc.).

Credential and environment notes

- These PowerShell scripts typically use `pwps_dab` cmdlets which require the ProjectWise CLI/PowerShell module and network access to the ProjectWise datasource.
- Running these scripts in CI or a headless server will require the service account to have the correct ProjectWise permissions.
- If you cannot run `pwps_dab` on your machine, you can still run the pipeline and skip automated download/upload (set `pw.ps1_download`/`ps1_upload` but set `fetch_external.pw_extract=false` and run those steps manually).

Running the pipeline (examples)

Run step 3 (classification) for a single UAID:

```powershell
python main.py --step step3 --target-uaid2 HS2-00002WUML
```

Run MPDT generation for a specific UAID (will create MPDT under `Output/`):

```powershell
python -m mpdt_generator.generator --workspace . --target-uaid2 HS2-00002WUML
```

Run the full pipeline for all UAIDs listed in `config/pipeline_config.json`:

```powershell
python main.py --step all
```

Upload generated MPDTs to ProjectWise (example call inside PS; the PS script handles datasource login):

```powershell
pwsh ./Scripts/PWPS_Upload_PW.ps1 -InputDir .\Output\<timestamp>\MPDT -Datasource "arcadis-uk-pw.bentley.com:arcadis-uk-07" -Username _asc_user_automation
```

Troubleshooting checklist

- If `Deliverable_Mapped_By` is empty in `Output/asset_deliverables.xlsx`:
  - Confirm `Output/asset_deliverables.xlsx` contains the expected `ASSET_ID` values. Writer matches by `ASSET_ID` first, then by `PW_UAID`.
  - Check classifier debug CSVs in the debug folder (e.g., `debug_<UAID>/mapping_eval.csv`) to see why the rules didn't match.

- If `Software Model Part ID` is missing:
  - Ensure `Input/ACBOS MPDT.xlsx` contains the UAID row with a valid document name. If not present, the pipeline attempts a ProjectWise download if `pw.ps1_download` is configured.
  - If PW download fails, run `Scripts/PWPS_Download_MPDT.ps1` manually and confirm it prints `DownloadedFile=<path>`.

- If ProjectWise scripts fail with missing cmdlets:
  - Confirm `pwps_dab` is installed and your user context can run its cmdlets.
  - Run the PS script interactively and check for credential prompts.

- If BO+ trimming or attribute applicability appears wrong:
  - Verify the `lodm.xlsx` file used by the pipeline is the correct, up-to-date LoDM and that the `CURRENT` sheet is the one intended.
  - Check `debug_<UAID>/scope3_rows.csv` and `mapping_eval.csv` to see attribute presence and counts.

Next steps / recommended checks before running in production

- Confirm `Input/` files are the correct, current canonical sources.
- If you want automated ProjectWise downloads/uploads, test `Scripts/PWPS_Download_MPDT.ps1` and `Scripts/PWPS_Upload_PW.ps1` interactively on a machine with PW access.
- Run a small batch of UAIDs end-to-end and inspect both the generated MPDT workbook and the debug CSVs.

Contacts / maintenance

- Code areas to inspect for future changes:
  - `classifier/type_classifier.py` — mapping rules and priority.
  - `mpdt_generator/generator.py` — MPDT population, round-2 trimming and applicability logic.
  - `excel_generator/writer.py` — write-back of `Deliverable_Mapped_By` and final output formatting.
  - `Scripts/*.ps1` — ProjectWise interaction wrappers (may need tuning per environment).

If you'd like, I can:
- Replace the existing `README.md` with this file or add it to the repo as `README_GENERATED.md` (I have saved this as `README_GENERATED.md`).
- Run a batch audit that counts how many deliverables used each mapping method and produce a short CSV for review.
- Attempt to run the ProjectWise downloader in your environment if you run the PS script and paste back any error logs.



---
Generated by GitHub Copilot (assistant). If you want edits or more/less detail, tell me which sections to expand or compress.
