"""
data_loader/local_loader.py — Load all local input files and consolidate trackers.

Each loader function can be called independently for testing/debugging.
"""
from __future__ import annotations

import json
import logging
from pathlib import Path
from datetime import datetime

import pandas as pd

from utils.common import (
    normalize_text,
    pick_first_existing_column,
    as_upper_key,
    read_table_any,
    write_table_cache,
    cache_signature,
    read_df_cache,
    write_df_cache,
)

# Column name variants for ID matching
UAID2_CANDIDATES = ["UAID_2", "UAID2", "Uaid_2", "Uaid", "UAID", "ASSET_ID", "ASSET ID", "SID"]

# Tracker column normalisation map
_TRACKER_COL_MAP = {
    "uaid_2": "UAID_2",
    "uaid2": "UAID_2",
    "asset id": "UAID_2",
    "level 2 asset name": "Asset_Name",
    "asset name": "Asset_Name",
    "assetname": "Asset_Name",
    "pw document name": "PW_Document_Name",
    "document name": "PW_Document_Name",
    "documentname": "PW_Document_Name",
    "revision": "Revision",
    "rev": "Revision",
    "status": "Status",
    "file type": "File_Type",
    "type": "File_Type",
}


def _normalise_cols(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [_TRACKER_COL_MAP.get(normalize_text(c), c) for c in df.columns]
    return df


def _read_excel_with_cache(
    path: Path,
    cache_dir: Path,
    cache_name: str,
    logger: logging.Logger,
    **read_excel_kwargs,
) -> pd.DataFrame:
    """Read an Excel sheet once, then reuse a source-validated pickle cache."""
    sig = cache_signature(path)
    cache_path = cache_dir / f"{cache_name}_v2_preserve_none.pkl"
    cached = read_df_cache(cache_path, sig, logger)
    if cached is not None:
        logger.info("%s loaded from fast cache: %d rows", cache_name, len(cached))
        return cached
    # Preserve literal strings such as "None" from source Excel files.
    # Pandas' default NA parser can otherwise convert them to NaN before the
    # MPDT/ACBOS generators get a chance to write the required text value.
    read_excel_kwargs.setdefault("keep_default_na", False)
    df = pd.read_excel(path, **read_excel_kwargs)
    write_df_cache(df, cache_path, sig, logger)
    return df



# Attribute codes that must never be propagated from LoDM/L3 into generated files.
# Normalisation removes spaces, underscores, hyphens, colons, dots and ampersands,
# so all variants of Com_AssetRef are caught.
_ATTR_CODE_SEP_RE = __import__("re").compile(r"[\s_\-\.:&]+")
_OMITTED_ATTRIBUTE_CODES = {"comassetref", "assetref"}

def _norm_attr_code(value) -> str:
    return _ATTR_CODE_SEP_RE.sub("", str(value or "")).lower()

def _is_omitted_attribute(value) -> bool:
    code = _norm_attr_code(value)
    return code in _OMITTED_ATTRIBUTE_CODES or code.endswith("assetref")

def _drop_omitted_attribute_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Drop deprecated attribute columns such as Com_AssetRef from source data."""
    if df is None or df.empty:
        return df
    drop_cols = [c for c in df.columns if _is_omitted_attribute(c)]
    return df.drop(columns=drop_cols, errors="ignore") if drop_cols else df

def _drop_omitted_lodm_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Remove deprecated attributes from the LoDM itself before downstream logic."""
    if df is None or df.empty:
        return df
    candidate_cols = [
        c for c in df.columns
        if normalize_text(c) in {
            "atttypename", "atttypedescription", "attrtypedisplayname",
            "attribute", "attributename"
        }
    ]
    if not candidate_cols:
        return df
    mask = pd.Series(False, index=df.index)
    for c in candidate_cols:
        mask = mask | df[c].map(_is_omitted_attribute)
    return df.loc[~mask].copy()

def _normalise_l3_source_values(df: pd.DataFrame) -> pd.DataFrame:
    """Clean L3 source data while preserving literal string 'None'.

    Existing pickle/parquet caches created before keep_default_na=False may have
    already converted Excel text 'None' to NaN.  This function cannot infer text
    from genuinely blank cells, but it prevents future conversion and ensures
    deprecated columns are removed at source.
    """
    if df is None or df.empty:
        return df
    df = _drop_omitted_attribute_columns(df.copy())
    # Avoid pandas nullable values leaking into later string comparisons.  Do not
    # convert actual missing values to the text 'None'.
    return df

# ---------------------------------------------------------------------------
# Individual loaders
# ---------------------------------------------------------------------------

def load_pw_extract(workspace: Path, cfg: dict, logger: logging.Logger) -> pd.DataFrame:
    """Load ProjectWise extract, preferring the full-columns source of truth."""
    paths = cfg.get("paths", {})
    # Priority: full-columns PW extract first, then older configured extracts,
    # then the cached extract as a last resort.
    candidates = [
        ("pw_extract_full_columns", paths.get("pw_extract_full_columns", "")),
        ("pw_extract_full_columns_default", "Input/ACBOS MPDT_FULLColumns.xlsx"),
        ("pw_extract_full", paths.get("pw_extract_full", "")),
        ("pw_extract", paths.get("pw_extract", "")),
    ]
    seen: set[str] = set()
    for key, rel in candidates:
        if not rel:
            continue
        path = (workspace / rel).resolve()
        path_key = str(path).lower()
        if path_key in seen:
            continue
        seen.add(path_key)
        if path.exists():
            try:
                cache_dir = workspace / paths.get("db_cache_dir", "Input/DB_Cache")
                df = _read_excel_with_cache(path, cache_dir, f"excel_{path.stem}", logger, dtype=str, engine="openpyxl")
                logger.info("PW extract loaded from '%s' (%s): %d rows", path.name, key, len(df))
                return df
            except PermissionError:
                logger.warning("PW extract '%s' is locked; trying fallback source...", path.name)
            except Exception as exc:
                logger.warning("Could not read PW extract '%s': %s", path.name, exc)
    cache_base = workspace / paths.get("db_cache_dir", "Input/DB_Cache") / "PW_Extract"
    cached = read_table_any(cache_base, logger)
    if not cached.empty:
        logger.info("PW extract loaded from cache: %d rows", len(cached))
        return cached

    logger.warning("No PW extract found.")
    return pd.DataFrame()


def load_l2_mapping(workspace: Path, cfg: dict, logger: logging.Logger) -> pd.DataFrame:
    """Load L2 UAID-ACBOS mapping file."""
    rel = cfg.get("paths", {}).get("l2_uaid_acbos", "Input/L2 UAID-ACBOS_260129.xlsx")
    path = (workspace / rel).resolve()
    if not path.exists():
        logger.warning("L2 mapping not found: %s", path)
        return pd.DataFrame()

    cache_dir = workspace / cfg.get("paths", {}).get("db_cache_dir", "Input/DB_Cache")
    df = _read_excel_with_cache(path, cache_dir, "l2_uaid_acbos", logger, dtype=str, engine="openpyxl")
    # Normalise key columns
    rename = {}
    for c in df.columns:
        nc = normalize_text(c)
        if nc in ("uaid_2", "uaid2", "uaid 2"):
            rename[c] = "UAID_2"
        elif nc in ("level 2 asset name", "asset name", "assetname", "l2 asset name"):
            rename[c] = "Asset_Name"
        elif nc in ("acbos", "acbos document", "acbos doc"):
            rename[c] = "ACBOS_Doc"
    df.rename(columns=rename, inplace=True)
    logger.info("L2 mapping loaded: %d rows", len(df))
    return df


def load_tracker(path: Path, source_label: str, sheet_name: str, logger: logging.Logger) -> pd.DataFrame | None:
    """Load a single JSON Works Tracker file."""
    if not path.exists():
        logger.warning("Tracker not found: %s", path)
        return None
    try:
        xl = pd.ExcelFile(path, engine="openpyxl")
        sheet = sheet_name if sheet_name in xl.sheet_names else xl.sheet_names[0]
        if sheet != sheet_name:
            logger.debug("  '%s' sheet not found in %s — using '%s'", sheet_name, path.name, sheet)
        df = pd.read_excel(path, sheet_name=sheet, dtype=str, engine="openpyxl", keep_default_na=False)
        df = _normalise_cols(df)
        df["_source"] = source_label
        logger.info("  Loaded tracker '%s' (sheet='%s'): %d rows", path.name, sheet, len(df))
        return df
    except Exception as exc:
        logger.warning("Could not read tracker '%s': %s", path.name, exc)
        return None


def consolidate_trackers(workspace: Path, cfg: dict, logger: logging.Logger) -> pd.DataFrame:
    """Load and consolidate all JSON Works Tracker files."""
    paths_cfg = cfg.get("paths", {})
    sheet_name = cfg.get("tracker_sheet_name", "Central_Team")

    tracker_keys = ["tracker_sep2025", "tracker_2025", "tracker_2024"]
    frames = []
    for key in tracker_keys:
        rel = paths_cfg.get(key)
        if not rel:
            continue
        path = (workspace / rel).resolve()
        df = load_tracker(path, key, sheet_name, logger)
        if df is not None:
            frames.append(df)

    if not frames:
        logger.warning("No tracker files loaded.")
        return pd.DataFrame()

    combined = pd.concat(frames, ignore_index=True)
    if "UAID_2" in combined.columns:
        combined = combined.dropna(subset=["UAID_2"])
        combined["UAID_2"] = combined["UAID_2"].str.strip()
        combined = combined.drop_duplicates(subset="UAID_2", keep="first")

    logger.info("Consolidated tracker: %d unique UAIDs", len(combined))

    # Cache
    cache_dir = workspace / cfg.get("paths", {}).get("db_cache_dir", "Input/DB_Cache")
    write_table_cache(combined, cache_dir / "json_works_tracker_consolidated", logger)

    return combined


def load_scope2(workspace: Path, cfg: dict, logger: logging.Logger) -> pd.DataFrame:
    """Load AssetsScope2 from cache."""
    cache_dir = workspace / cfg.get("paths", {}).get("db_cache_dir", "Input/DB_Cache")
    df = read_table_any(cache_dir / "AssetsScope2_full", logger)
    if df.empty:
        logger.warning("AssetsScope2 not found in cache.")
    else:
        logger.info("AssetsScope2 loaded: %d rows", len(df))
    return df


def load_scope3(workspace: Path, cfg: dict, logger: logging.Logger) -> pd.DataFrame:
    """Load AssetsScope3.

    Preferred source: user-updated Input/l3_assets_scope_data (parquet > xlsx).
    Fallback: DB_Cache/AssetsScope3_full (parquet > xlsx).
    """
    paths = cfg.get("paths", {})
    cache_dir = workspace / paths.get("db_cache_dir", "Input/DB_Cache")

    # Check pickle/parquet cache in DB_Cache first (fast). Pickle does not
    # require pyarrow/fastparquet and is usually fastest for local reruns.
    pkl_cache = cache_dir / "l3_assets_scope_data_v2_preserve_none.pkl"
    if pkl_cache.exists():
        try:
            df = pd.read_pickle(pkl_cache)
            df = _normalise_l3_source_values(df)
            logger.info("AssetsScope3 loaded from pickle cache '%s': %d rows", pkl_cache.name, len(df))
            return df
        except Exception as exc:
            logger.warning("Could not read scope3 pickle cache '%s': %s", pkl_cache.name, exc)

    parquet_cache = cache_dir / "l3_assets_scope_data_v2_preserve_none.parquet"
    if parquet_cache.exists():
        try:
            df = pd.read_parquet(parquet_cache)
            df = _normalise_l3_source_values(df)
            logger.info("AssetsScope3 loaded from parquet cache '%s': %d rows", parquet_cache.name, len(df))
            try:
                df.to_pickle(pkl_cache)
            except Exception:
                pass
            return df
        except Exception as exc:
            logger.warning("Could not read scope3 parquet cache '%s': %s", parquet_cache.name, exc)

    # Prefer explicit updated local file — xlsx (slow for large files)
    rel = paths.get("l3_assets_scope_data", "Input/l3_assets_scope_data.xlsx")
    src_base = (workspace / rel).resolve().with_suffix("")  # strip .xlsx to get base path
    xlsx_src = src_base.with_suffix(".xlsx")

    if xlsx_src.exists():
        try:
            df = _read_excel_with_cache(xlsx_src, cache_dir, "l3_assets_scope_data", logger, dtype=str, engine="openpyxl")
            df = _normalise_l3_source_values(df)
            logger.info("AssetsScope3 loaded from updated input '%s': %d rows", xlsx_src.name, len(df))
            # Save parquet cache to DB_Cache for fast future loads
            cache_dir = workspace / cfg.get("paths", {}).get("db_cache_dir", "Input/DB_Cache")
            parquet_cache = cache_dir / "l3_assets_scope_data_v2_preserve_none.parquet"
            try:
                cache_dir.mkdir(parents=True, exist_ok=True)
                df.to_pickle(cache_dir / "l3_assets_scope_data_v2_preserve_none.pkl")
                df.to_parquet(parquet_cache, index=False)
                logger.info("Saved scope3 pickle/parquet cache: %s", parquet_cache.name)
            except Exception as pexc:
                logger.warning("Could not save scope3 parquet cache: %s", pexc)
            return df
        except Exception as exc:
            logger.warning("Could not read updated AssetsScope3 file '%s': %s", xlsx_src, exc)

    # Fallback to cached AssetsScope3 snapshot (parquet preferred via read_table_any)
    cache_dir = workspace / cfg.get("paths", {}).get("db_cache_dir", "Input/DB_Cache")
    df = read_table_any(cache_dir / "AssetsScope3_full", logger)
    if df.empty:
        logger.warning("AssetsScope3 not found in cache.")
    else:
        df = _normalise_l3_source_values(df)
        logger.info("AssetsScope3 loaded: %d rows", len(df))
    return df


def _export_join_results(
    workspace: Path,
    cfg: dict,
    scope3_df: pd.DataFrame,
    sf_l3: pd.DataFrame,
    logger: logging.Logger,
) -> pd.DataFrame:
    """Build and export temporary join_results.xlsx (SmartForms L3 x Scope3)."""
    if scope3_df.empty or sf_l3.empty:
        logger.info("join_results skipped: Scope3 or SmartForms L3 is empty.")
        return pd.DataFrame()

    uaid3_col_s3 = pick_first_existing_column(scope3_df, ["UAID_3", "Uaid_3", "Uaid", "UAID"])
    uaid3_col_sf = pick_first_existing_column(sf_l3, ["UAID_3", "Uaid_3", "Uaid", "UAID", "Asset_ID"])
    if not uaid3_col_s3 or not uaid3_col_sf:
        logger.warning("join_results skipped: UAID_3 key column not found in Scope3 or SmartForms L3.")
        return pd.DataFrame()

    left = sf_l3.add_suffix("_sform3").copy()
    right = scope3_df.add_suffix("_l3").copy()
    left_key = f"{uaid3_col_sf}_sform3"
    right_key = f"{uaid3_col_s3}_l3"

    left["__join_key__"] = left[left_key].fillna("").astype(str).str.strip().str.upper()
    right["__join_key__"] = right[right_key].fillna("").astype(str).str.strip().str.upper()

    joined = pd.merge(left, right, on="__join_key__", how="inner")
    if "__join_key__" in joined.columns:
        joined.drop(columns=["__join_key__"], inplace=True)

    # Writing this merged diagnostic file to Excel is extremely expensive for
    # large SmartForms/Scope3 inputs. Keep it disabled by default; enable with
    # config: {"export_join_results": true} only when debugging the join.
    if not bool(cfg.get("export_join_results", False)):
        logger.info("join_results built in memory only (%d rows, %d cols); Excel export disabled.", len(joined), len(joined.columns))
        return joined

    out_rel = cfg.get("paths", {}).get("join_results", "Input/join_results.xlsx")
    out_path = (workspace / out_rel).resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        # Prefer pickle cache for speed; Excel export is just for manual inspection.
        write_table_cache(joined, out_path.with_suffix(""), logger)
        logger.info("join_results exported: %s (%d rows, %d cols)", out_path, len(joined), len(joined.columns))
    except Exception as exc:
        logger.warning("Could not write join_results '%s': %r", out_path, exc)

    return joined


def load_smartforms(workspace: Path, cfg: dict, logger: logging.Logger) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Load SmartForms L2 and L3 from cache."""
    cache_dir = workspace / cfg.get("paths", {}).get("db_cache_dir", "Input/DB_Cache")
    sf_l2 = read_table_any(cache_dir / "SmartForms_L2", logger)
    sf_l3 = read_table_any(cache_dir / "SmartForms_L3", logger)
    logger.info("SmartForms loaded — L2: %d rows, L3: %d rows", len(sf_l2), len(sf_l3))
    return sf_l2, sf_l3


def load_lodm(workspace: Path, cfg: dict, logger: logging.Logger) -> pd.DataFrame:
    """Load LoDM (attribute applicability per UAID_2)."""
    rel = cfg.get("paths", {}).get("lodm", "Input/1MC06-ASC-IM-SCH-C002-000009_lodm.xlsx")
    path = (workspace / rel).resolve()
    if not path.exists():
        logger.warning("LoDM file not found: %s", path)
        return pd.DataFrame()
    cache_dir = workspace / cfg.get("paths", {}).get("db_cache_dir", "Input/DB_Cache")
    df = _read_excel_with_cache(path, cache_dir, "lodm_current", logger, sheet_name="CURRENT LoDM", dtype=str, engine="openpyxl")
    df = _drop_omitted_lodm_rows(df)
    logger.info("LoDM loaded: %d rows, %d cols", len(df), len(df.columns))
    return df


def load_control_file(workspace: Path, cfg: dict, logger: logging.Logger) -> pd.DataFrame:
    """Load control file (multiple sheets combined)."""
    rel = cfg.get("paths", {}).get("control_file", "Input/1MC06-ASC-IM-GDE-C002-000090_controlfile.xlsx")
    path = (workspace / rel).resolve()
    if not path.exists():
        logger.warning("Control file not found: %s", path)
        return pd.DataFrame()

    cache_dir = workspace / cfg.get("paths", {}).get("db_cache_dir", "Input/DB_Cache")
    sig = cache_signature(path)
    cached = read_df_cache(cache_dir / "control_file_combined.pkl", sig, logger)
    if cached is not None:
        logger.info("Control file loaded from fast cache: %d rows", len(cached))
        return cached

    SKIP_SHEETS = {"cover page", "pr products", "calcs"}
    try:
        xl = pd.ExcelFile(path, engine="openpyxl")
        sheets_to_read = [s for s in xl.sheet_names if str(s).strip().lower() not in SKIP_SHEETS]
        dfs = []
        for sheet in sheets_to_read:
            try:
                df_s = pd.read_excel(path, sheet_name=sheet, header=1, dtype=str, engine="openpyxl", keep_default_na=False)
                df_s["__sheetname__"] = sheet
                dfs.append(df_s)
            except Exception:
                pass
        if dfs:
            combined = pd.concat(dfs, ignore_index=True)
            write_df_cache(combined, cache_dir / "control_file_combined.pkl", sig, logger)
            logger.info("Control file loaded: %d rows from %d sheets", len(combined), len(dfs))
            return combined
    except Exception as e:
        logger.warning("Control file read error: %s", e)
    return pd.DataFrame()


def load_mpdt_mapping(workspace: Path, cfg: dict, logger: logging.Logger) -> dict[str, str]:
    """Load MPDT column mapping from template mapping file."""
    rel = cfg.get("paths", {}).get("mpdt_template", "Input/C2-MPDT-Template-Mapping.xlsm")
    path = (workspace / rel).resolve()
    if not path.exists():
        logger.warning("MPDT mapping file not found: %s", path)
        return {}

    cache_dir = workspace / cfg.get("paths", {}).get("db_cache_dir", "Input/DB_Cache")
    cache_dir.mkdir(parents=True, exist_ok=True)
    cache_json = cache_dir / "mpdt_mapping_cache.json"
    cache_csv = cache_dir / "mpdt_mapping_cache.csv"

    src_sig = {
        "path": str(path),
        "mtime_ns": path.stat().st_mtime_ns,
        "size": path.stat().st_size,
    }

    if cache_json.exists():
        try:
            payload = json.loads(cache_json.read_text(encoding="utf-8"))
            if payload.get("source") == src_sig and isinstance(payload.get("mapping"), dict):
                mapping = {str(k): str(v) for k, v in payload["mapping"].items() if str(v).strip()}
                logger.info("MPDT mapping loaded from cache: %d column mappings", len(mapping))
                return mapping
        except Exception as exc:
            logger.debug("Could not read MPDT mapping cache (%s). Rebuilding cache.", exc)

    # The mapping row lives near the top of the sheet. Read only a few rows to
    # avoid loading the full sheet which can be very slow for large workbooks.
    try:
        df = pd.read_excel(path, sheet_name="MPDT Element of Asset", header=1, engine="openpyxl", nrows=8)
    except Exception:
        # Fallback to full read if the optimized read fails
        df = pd.read_excel(path, sheet_name="MPDT Element of Asset", header=1, engine="openpyxl")

    if len(df) < 1:
        return {}
    # mapping row is usually the 3rd row in original sheet (index 2), but if
    # the small read returned fewer rows pick the last available row.
    row_idx = 2 if len(df) > 2 else (len(df) - 1)
    mapping_row = df.iloc[row_idx]
    mapping = {}
    for col, val in mapping_row.items():
        if pd.notna(val) and str(val).strip():
            mapping[col] = str(val)

    try:
        cache_payload = {"source": src_sig, "mapping": mapping}
        cache_json.write_text(json.dumps(cache_payload, indent=2), encoding="utf-8")
        pd.DataFrame(
            [{"target_column": k, "mapping_expression": v} for k, v in mapping.items()]
        ).to_csv(cache_csv, index=False, encoding="utf-8")
    except Exception as exc:
        logger.debug("Could not write MPDT mapping cache (%s)", exc)

    logger.info("MPDT mapping loaded: %d column mappings (from row %d)", len(mapping), row_idx)
    return mapping


def load_sample_mpdt_columns(workspace: Path, cfg: dict, logger: logging.Logger) -> list[str]:
    """Load column headers from sample MPDT file."""
    rel = cfg.get("paths", {}).get("sample_mpdt", "Input/1MC07-CEK-BR-FRM-CS07_CL26-000005.xlsm")
    path = (workspace / rel).resolve()
    if not path.exists():
        logger.warning("Sample MPDT file not found: %s", path)
        return []
    df = pd.read_excel(path, sheet_name="MPDT Element of Asset", engine="openpyxl", header=1, nrows=0)
    logger.info("Sample MPDT columns loaded: %d columns", len(df.columns))
    return list(df.columns)


def load_sample_mpdt_row1_codes(workspace: Path, cfg: dict, logger: logging.Logger) -> dict[str, str]:
    """Return {row2_header: row1_short_code} for all columns in the sample MPDT.

    Row 1 contains AttTypeName short codes (e.g. 'NmnlSrfcAr', 'Mtrl') for
    LoDM attribute columns, and DB field names (e.g. 'UAID_2', 'AssetName_2')
    or section labels for governance/metadata columns.  Columns with no Row 1
    code (empty string) are treated as governance columns — never blackened and
    never physically removed from the output MPDT.
    """
    rel = cfg.get("paths", {}).get("sample_mpdt", "Input/1MC07-CEK-BR-FRM-CS07_CL26-000005.xlsm")
    path = (workspace / rel).resolve()
    if not path.exists():
        logger.warning("Sample MPDT not found for row1 codes: %s", path)
        return {}

    # Read Row 1 values without any header (gives raw cell values by position)
    df_r1 = pd.read_excel(
        path, sheet_name="MPDT Element of Asset", engine="openpyxl", header=None, nrows=1
    )
    # Read Row 2 as column names — exactly as load_sample_mpdt_columns does —
    # so the dict keys will always match what the rest of the pipeline uses.
    df_r2 = pd.read_excel(
        path, sheet_name="MPDT Element of Asset", engine="openpyxl", header=1, nrows=0
    )
    if df_r1.empty:
        return {}

    row1_vals = [str(v).strip() if pd.notna(v) else "" for v in df_r1.iloc[0]]
    row2_cols = list(df_r2.columns)  # identical to load_sample_mpdt_columns() output

    result: dict[str, str] = {}
    for r1, r2 in zip(row1_vals, row2_cols):
        result[r2] = r1  # r2 key matches exactly; r1 may be empty for governance cols

    logger.info(
        "MPDT row1 codes loaded: %d columns (%d with short code)",
        len(result), sum(1 for v in result.values() if v),
    )
    return result




def load_midp_navigator(workspace: Path, cfg: dict, logger: logging.Logger) -> pd.DataFrame:
    """Load MIDP_Navigator.xlsx used to resolve MPDT Model Container IDs.

    The file can be large, so it uses the same source-validated pickle cache as
    the other Excel inputs. Missing file is allowed: MPDT generation will leave
    the Model Container ID column blank.
    """
    rel = cfg.get("paths", {}).get("midp_navigator", "Input/MIDP_Navigator.xlsx")
    path = (workspace / rel).resolve()
    if not path.exists():
        logger.warning("MIDP Navigator file not found: %s", path)
        return pd.DataFrame()

    cache_dir = workspace / cfg.get("paths", {}).get("db_cache_dir", "Input/DB_Cache")
    try:
        df = _read_excel_with_cache(path, cache_dir, f"excel_{path.stem}", logger, dtype=str, engine="openpyxl")
        logger.info("MIDP Navigator loaded: %d rows", len(df))
        return df
    except Exception as exc:
        logger.warning("Could not load MIDP Navigator '%s': %s", path, exc)
        return pd.DataFrame()


def load_all_current_dm3_files(workspace: Path, cfg: dict, logger: logging.Logger) -> pd.DataFrame:
    """Load All_Current_DM3_files.xlsx for a first-pass MPDT Model Container lookup."""
    rel = cfg.get("paths", {}).get("all_current_dm3_files", "Input/All_Current_DM3_files.xlsx")
    path = (workspace / rel).resolve()
    if not path.exists():
        logger.warning("All Current DM3 files input not found: %s", path)
        return pd.DataFrame()

    cache_dir = workspace / cfg.get("paths", {}).get("db_cache_dir", "Input/DB_Cache")
    try:
        df = _read_excel_with_cache(path, cache_dir, f"excel_{path.stem}", logger, dtype=str, engine="openpyxl")
        logger.info("All Current DM3 files loaded: %d rows", len(df))
        return df
    except Exception as exc:
        logger.warning("Could not load All Current DM3 files '%s': %s", path, exc)
        return pd.DataFrame()


# ---------------------------------------------------------------------------
# All-in-one loader (used by main pipeline)
# ---------------------------------------------------------------------------

def load_all_sources(workspace: Path, cfg: dict, logger: logging.Logger, target_uaids: list[str] | None = None) -> dict:
    """Load all required data sources and return as a dict.

    target_uaids is accepted for future source filtering; current implementation
    still loads complete canonical tables but avoids unnecessary Excel diagnostics.
    """
    logger.info("Loading all local data sources...")

    pw_df = load_pw_extract(workspace, cfg, logger)
    l2_df = load_l2_mapping(workspace, cfg, logger)
    scope2_df = load_scope2(workspace, cfg, logger)
    scope3_df = load_scope3(workspace, cfg, logger)
    sf_l2, sf_l3 = load_smartforms(workspace, cfg, logger)
    join_results_df = _export_join_results(workspace, cfg, scope3_df, sf_l3, logger) if bool(cfg.get("build_join_results", False)) else pd.DataFrame()
    if join_results_df.empty:
        logger.info("join_results merge skipped during source load (set build_join_results=true to enable).")
    lodm_df = load_lodm(workspace, cfg, logger)
    control_df = load_control_file(workspace, cfg, logger)
    mapping_dict = load_mpdt_mapping(workspace, cfg, logger)
    midp_df = load_midp_navigator(workspace, cfg, logger)
    dm3_df = load_all_current_dm3_files(workspace, cfg, logger)
    mpdt_columns = load_sample_mpdt_columns(workspace, cfg, logger)
    mpdt_row1_codes = load_sample_mpdt_row1_codes(workspace, cfg, logger)

    # Extract UAID_2 sets for quick lookup
    pw_id_col = pick_first_existing_column(pw_df, UAID2_CANDIDATES) if not pw_df.empty else None
    l2_id_col = pick_first_existing_column(l2_df, UAID2_CANDIDATES) if not l2_df.empty else None

    return {
        "pw_df": pw_df,
        "l2_df": l2_df,
        "scope2_df": scope2_df,
        "scope3_df": scope3_df,
        "sf_l2": sf_l2,
        "sf_l3": sf_l3,
        "join_results_df": join_results_df,
        "lodm_df": lodm_df,
        "control_df": control_df,
        "mapping_dict": mapping_dict,
        "midp_df": midp_df,
        "dm3_df": dm3_df,
        "mpdt_columns": mpdt_columns,
        "row1_codes": mpdt_row1_codes,
        "pw_id_col": pw_id_col,
        "l2_id_col": l2_id_col,
    }
