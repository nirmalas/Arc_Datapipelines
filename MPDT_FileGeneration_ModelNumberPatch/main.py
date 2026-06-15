"""
main.py — End-to-end orchestrator for ACBOS/MPDT Pipeline V2.

Pipeline steps (each runnable individually):
  step1   — Data caching (DB + SmartForms + Trackers)
  step2   — Build asset deliverables (PW + L2 + Tracker join)
  step3   — Classify targets as MPDT or ACBOS
  step4   — Generate MPDT files (batch)
  step5   — Generate ACBOS files (batch)
  step6   — Stage files for PW upload

Data sourcing:
  --refresh-sources   Fetch fresh data from external sources (DB, SmartForms API)
                      and update the local Excel caches. Without this flag, all
                      steps use the existing local Excel files only.

Usage:
  python main.py --refresh-sources                    # just update inputs
  python main.py --target-uaid2 HS2-000001416 HS2-00002NSXW
  python main.py --target-uaid2 HS2-000001416 --file-type MPDT
  python main.py --step step3 --target-uaid2 HS2-000001416
  python main.py --step all --target-uaid2 HS2-000001416,HS2-00002NSXW
  python main.py --refresh-sources --step all --target-uaid2 HS2-000001416
"""
from __future__ import annotations

import argparse
import sys
import time
from datetime import datetime
from pathlib import Path

import pandas as pd

from utils.common import (
    load_config,
    resolve_workspace,
    setup_logger,
    timestamped_dir,
    write_json,
)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="ACBOS/MPDT Pipeline V2 — end-to-end orchestrator",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--target-uaid2", nargs="+", metavar="UAID_2",
        help="One or more UAID_2 identifiers (comma or space separated).",
    )
    parser.add_argument(
        "--file-type", choices=["AUTO", "MPDT", "ACBOS"], default=None,
        help="Force output type for all targets. Default: AUTO.",
    )
    parser.add_argument(
        "--step",
        choices=["all", "step1", "step2", "step3", "step4", "step5", "step6"],
        default="all",
        help="Which step to run. Default: all.",
    )
    parser.add_argument(
        "--refresh-sources", action="store_true",
        help="Fetch fresh data from external sources (DB, SmartForms API) "
             "and update local Excel caches. Without this flag, all steps "
             "use existing local files only.",
    )
    parser.add_argument("--no-cache", action="store_true", help="Force DB re-fetch (same as --refresh-sources).")
    parser.add_argument("--workspace", default=None)
    parser.add_argument("--log-level", default=None, choices=["DEBUG", "INFO", "WARNING", "ERROR"])
    parser.add_argument("--batch-size", type=int, default=None, help="Override batch size (default: 5).")
    return parser.parse_args()


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _load_uaids_from_file(path: Path) -> list[str]:
    """Read target UAID_2 values from an Excel file (UAIDs only, no filename map)."""
    uaids, _mpdt_map, _acbos_map = _load_matchup_file_full(path)
    return uaids


def _compact_col_name(value: object) -> str:
    """Normalize an input column name for robust matchup-file detection."""
    return str(value or "").strip().lower().replace(" ", "").replace("_", "").replace("-", "")


def _clean_matchup_docname(value: object, default_ext: str | None = None, strip_extension: bool = True) -> str:
    """Return a clean document/file name from a matchup cell.

    MPDT names are stored as stems because generator.py appends .xlsm.
    ACBOS names are stored as filenames with .ACBOS, because acbos_generator.py
    writes the package directly using the supplied filename.
    """
    raw = str(value or "").strip()
    if not raw or raw.lower() in ("nan", "none", "null"):
        return ""
    from pathlib import Path as _P
    name = _P(raw).name
    if strip_extension:
        name = _P(name).stem
    if default_ext and name and not name.upper().endswith(default_ext.upper()):
        name = f"{name}{default_ext}"
    return name


def _load_matchup_file_full(path: Path) -> tuple[list[str], dict[str, str], dict[str, str]]:
    """Read UAID_2 plus optional MPDT and ACBOS filenames from matchup Excel.

    Supported columns include:
      - Level 2 UAID / UAID_2 / UAID
      - MPDT: target MPDT document/file name
      - ACBOS: target ACBOS package filename

    Returns:
      (uaids, mpdt_map, acbos_map)
      where mpdt_map stores filename stems and acbos_map stores .ACBOS filenames.
    """
    if not path.exists():
        print(f"[WARNING] matchup_file not found: {path}")
        return [], {}, {}
    try:
        df = pd.read_excel(path, dtype=str, engine="openpyxl")
        if df.empty or len(df.columns) == 0:
            return [], {}, {}

        _UAID_HINTS = {"level2uaid", "uaid2", "uaid"}
        uaid_col = next((c for c in df.columns if _compact_col_name(c) in _UAID_HINTS), df.columns[0])

        # Prefer exact MPDT/ACBOS columns.  Generic document/file columns remain an
        # MPDT fallback only, preserving the older three-column matchup behavior.
        _MPDT_HINTS = {"mpdt", "mpdtfilename", "mpdtfile", "mpdtdocument", "mpdtdocumentname"}
        _MPDT_FALLBACK_HINTS = {"documentnumber", "documentname", "document", "filename", "file"}
        _ACBOS_HINTS = {"acbos", "acbosfilename", "acbosfile", "acbosdocument", "acbosdocumentname", "acbosdoc"}

        mpdt_col = next((c for c in df.columns if c != uaid_col and _compact_col_name(c) in _MPDT_HINTS), None)
        if mpdt_col is None:
            mpdt_col = next((c for c in df.columns if c != uaid_col and _compact_col_name(c) in _MPDT_FALLBACK_HINTS), None)
        acbos_col = next((c for c in df.columns if c != uaid_col and _compact_col_name(c) in _ACBOS_HINTS), None)

        df = df.dropna(subset=[uaid_col]).copy()
        df[uaid_col] = df[uaid_col].astype(str).str.strip()
        df = df[df[uaid_col].str.strip() != ""]
        df = df[~df[uaid_col].str.startswith("#")]

        uaids: list[str] = df[uaid_col].tolist()
        mpdt_map: dict[str, str] = {}
        acbos_map: dict[str, str] = {}
        for _, row in df.iterrows():
            u = str(row[uaid_col]).strip()
            if not u:
                continue
            if mpdt_col:
                mpdt_name = _clean_matchup_docname(row.get(mpdt_col, ""), strip_extension=True)
                if mpdt_name:
                    mpdt_map[u] = mpdt_name
            if acbos_col:
                acbos_name = _clean_matchup_docname(row.get(acbos_col, ""), default_ext=".ACBOS", strip_extension=False)
                if acbos_name:
                    acbos_map[u] = acbos_name

        print(
            f"[INFO] Loaded {len(uaids)} target UAIDs from {path.name} "
            f"(UAID column: '{uaid_col}', MPDT column: '{mpdt_col or ''}', ACBOS column: '{acbos_col or ''}', "
            f"{len(mpdt_map)} MPDT filenames, {len(acbos_map)} ACBOS filenames)"
        )
        return uaids, mpdt_map, acbos_map
    except Exception as exc:
        print(f"[WARNING] Could not read matchup_file '{path}': {exc}")
        return [], {}, {}


def _load_matchup_file(path: Path) -> tuple[list[str], dict[str, str]]:
    """Backward-compatible reader returning UAIDs and MPDT filename map only."""
    uaids, mpdt_map, _acbos_map = _load_matchup_file_full(path)
    return uaids, mpdt_map


def _apply_matchup_filenames_to_classification(
    classification: dict,
    mpdt_map: dict[str, str] | None,
    acbos_map: dict[str, str] | None,
) -> dict:
    """Attach explicit matchup filenames to classified targets.

    The classifier decides MPDT vs ACBOS; this function only sets the concrete
    filename to use for generation.  MPDT targets use the MPDT column.  ACBOS
    targets use the ACBOS column.  Existing target fields are preserved when the
    relevant matchup column is blank.
    """
    mpdt_map = mpdt_map or {}
    acbos_map = acbos_map or {}
    for target in classification.get("mpdt_targets", []) or []:
        uaid = str(target.get("uaid", "")).strip()
        if uaid in mpdt_map:
            target["file"] = mpdt_map[uaid]
            target["mpdt_file"] = mpdt_map[uaid]
    for target in classification.get("acbos_targets", []) or []:
        uaid = str(target.get("uaid", "")).strip()
        if uaid in acbos_map:
            target["file"] = acbos_map[uaid]
            target["acbos_file"] = acbos_map[uaid]
    return classification


# ---------------------------------------------------------------------------
# Step runners
# ---------------------------------------------------------------------------

def run_step1(workspace: Path, cfg: dict, logger, force_refresh: bool = False) -> dict:
    """Step 1: Optionally fetch fresh data from external sources and update local Input/ files.

    Control via config fetch_external flags:
      fetch_external.db          → query L3 SQL Server → overwrite Input/l3_assets_scope_data.xlsx
      fetch_external.smartforms  → fetch SmartForms API → overwrite Input/SmartForms_RAW_MPDT_L2&L3.xlsx
      fetch_external.pw_extract  → run PWPS_Data_Extract.ps1 → overwrites Input/ACBOS MPDT.xlsx

    When all flags are False (default) Step 1 simply verifies that the local Input/ files exist
    and reports their row counts. Steps 2+ always read from the same Input/ files.

    force_refresh=True (set by --refresh-sources CLI flag) overrides all flags to True.
    """
    logger.info("=" * 50)
    logger.info("STEP 1: Data Sources — L3 DB / SmartForms / ProjectWise")
    logger.info("=" * 50)

    fetch_cfg = cfg.get("fetch_external", {})
    do_db = force_refresh or bool(fetch_cfg.get("db", False))
    do_sf = force_refresh or bool(fetch_cfg.get("smartforms", False))
    do_pw = force_refresh or bool(fetch_cfg.get("pw_extract", False))

    paths = cfg.get("paths", {})
    scope3_path = (workspace / paths.get("l3_assets_scope_data", "Input/l3_assets_scope_data.xlsx")).resolve()
    sf_path = (workspace / paths.get("smartforms_fallback", "Input/SmartForms_RAW_MPDT_L2&L3.xlsx")).resolve()
    pw_path = (workspace / paths.get("pw_extract", "Input/ACBOS MPDT.xlsx")).resolve()

    results: dict = {}

    # -----------------------------------------------------------------------
    # L3 Database → Input/l3_assets_scope_data.xlsx
    # -----------------------------------------------------------------------
    if do_db:
        logger.info("--- Fetching L3 Database → %s ---", scope3_path.name)
        from data_fetcher.db_fetcher import fetch_all_db_tables
        try:
            db_results = fetch_all_db_tables(workspace, cfg, logger)
            # AssetsScope3 is the canonical scope3 input — write it to the Input/ file.
            scope3_df = db_results.get("AssetsScope3")
            if scope3_df is not None and not scope3_df.empty:
                scope3_path.parent.mkdir(parents=True, exist_ok=True)
                scope3_df.to_excel(str(scope3_path), index=False)
                logger.info("  AssetsScope3 written to %s (%d rows)", scope3_path.name, len(scope3_df))
                results["scope3_rows"] = len(scope3_df)
            scope2_df = db_results.get("AssetsScope2")
            if scope2_df is not None:
                results["scope2_rows"] = len(scope2_df)
        except Exception as exc:
            logger.warning("DB fetch failed (non-fatal): %s", exc)
            results["db_error"] = str(exc)
    else:
        if scope3_path.exists():
            import pandas as _pd
            _rows = len(_pd.read_excel(str(scope3_path), dtype=str, engine="openpyxl", usecols=[0]))
            logger.info("  L3 DB   — using local: %s (%d rows)", scope3_path.name, _rows)
            results["scope3_rows"] = _rows
        else:
            logger.warning("  L3 DB   — local file missing: %s  (set fetch_external.db=true to fetch)", scope3_path.name)

    # -----------------------------------------------------------------------
    # SmartForms → Input/SmartForms_RAW_MPDT_L2&L3.xlsx
    # -----------------------------------------------------------------------
    if do_sf:
        logger.info("--- Fetching SmartForms API → %s ---", sf_path.name)
        from data_fetcher.smartforms_fetcher import fetch_smartforms_and_save
        try:
            l2_rows, l3_rows = fetch_smartforms_and_save(workspace, cfg, logger, sf_path)
            logger.info("  SmartForms written — L2: %d, L3: %d", l2_rows, l3_rows)
            results["smartforms_l2_rows"] = l2_rows
            results["smartforms_l3_rows"] = l3_rows
        except Exception as exc:
            logger.warning("SmartForms fetch failed (non-fatal): %s", exc)
            results["smartforms_error"] = str(exc)
    else:
        if sf_path.exists():
            logger.info("  SmartForms — using local: %s", sf_path.name)
            results["smartforms_local"] = str(sf_path.name)
        else:
            logger.warning("  SmartForms — local file missing: %s  (set fetch_external.smartforms=true to fetch)", sf_path.name)

    # -----------------------------------------------------------------------
    # ProjectWise extract → Input/ACBOS MPDT.xlsx  (via PowerShell script)
    # -----------------------------------------------------------------------
    if do_pw:
        ps1_rel = cfg.get("pw", {}).get("ps1_extract", "Scripts/PWPS_Data_Extract.ps1")
        ps1_path = (workspace / ps1_rel).resolve()
        logger.info("--- Running PW extract script → %s ---", pw_path.name)
        if not ps1_path.exists():
            logger.warning("  PW PS1 script not found: %s", ps1_path)
            results["pw_error"] = "ps1_not_found"
        else:
            import subprocess
            try:
                proc = subprocess.run(
                    ["pwsh", "-File", str(ps1_path), "-OutputPath", str(pw_path)],
                    capture_output=True, text=True, timeout=600,
                )
                if proc.returncode == 0:
                    logger.info("  PW extract script completed OK — output: %s", pw_path.name)
                    results["pw_extract_ok"] = True
                else:
                    logger.warning("  PW extract script exited %d:\n%s", proc.returncode, proc.stderr[-2000:])
                    results["pw_error"] = f"exit_code={proc.returncode}"
            except Exception as exc:
                logger.warning("  PW extract script failed: %s", exc)
                results["pw_error"] = str(exc)
    else:
        if pw_path.exists():
            logger.info("  PW extract — using local: %s", pw_path.name)
            results["pw_local"] = str(pw_path.name)
        else:
            logger.warning("  PW extract — local file missing: %s  (set fetch_external.pw_extract=true to fetch)", pw_path.name)

    results["status"] = "success"
    logger.info("Step 1 complete.")
    return results


def run_step2(
    workspace: Path,
    cfg: dict,
    logger,
    pw_df: "pd.DataFrame | None" = None,
    l2_df: "pd.DataFrame | None" = None,
) -> dict:
    """Step 2: Build asset deliverables summary."""
    logger.info("=" * 50)
    logger.info("STEP 2: Build Asset Deliverables")
    logger.info("=" * 50)

    from data_loader.local_loader import load_pw_extract, load_l2_mapping
    from pw_processor.step2b_builder import build_asset_deliverables
    from excel_generator.writer import write_excel

    if pw_df is None:
        pw_df = load_pw_extract(workspace, cfg, logger)
    if l2_df is None:
        l2_df = load_l2_mapping(workspace, cfg, logger)

    # Augment PW extract with matchup file rows (one row per UAID → DocumentName pair).
    # The matchup file provides the ASSET_ID link that the raw PW extract lacks.
    matchup_path_rel = cfg.get("paths", {}).get("matchup_file", "")
    matchup_extra_rows: list[dict] = []
    if matchup_path_rel:
        _mu_path = (workspace / matchup_path_rel).resolve()
        _mu_uaids, _mu_mpdt_map, _mu_acbos_map = _load_matchup_file_full(_mu_path)
        _all_doc_map: dict[tuple[str, str], str] = {}
        for _uaid, _doc in (_mu_mpdt_map or {}).items():
            _all_doc_map[(_uaid, _doc)] = _doc
        for _uaid, _doc in (_mu_acbos_map or {}).items():
            _all_doc_map[(_uaid, Path(_doc).stem)] = Path(_doc).stem
        if _all_doc_map:
            # Build a DocumentName→PW-row lookup for fast metadata enrichment.
            pw_by_docname: dict[str, dict] = {}
            if pw_df is not None and not pw_df.empty and "DocumentName" in pw_df.columns:
                for _, pw_row in pw_df.iterrows():
                    dn = str(pw_row.get("DocumentName", "")).strip().lower()
                    if dn and dn not in pw_by_docname:
                        pw_by_docname[dn] = pw_row.to_dict()
            for (uaid, doc_stem), _doc in _all_doc_map.items():
                base = pw_by_docname.get(doc_stem.lower(), {})
                row: dict = dict(base)
                row["ASSET_ID"] = uaid
                row["DocumentName"] = doc_stem
                matchup_extra_rows.append(row)
            logger.info("  Matchup file added %d UAID->DocumentName rows for step2.", len(matchup_extra_rows))

    if matchup_extra_rows:
        extra_df = pd.DataFrame(matchup_extra_rows)
        pw_augmented = pd.concat(
            [pw_df if pw_df is not None else pd.DataFrame(), extra_df],
            ignore_index=True,
        )
    else:
        pw_augmented = pw_df if pw_df is not None else pd.DataFrame()

    deliverables_df = build_asset_deliverables(pw_augmented, l2_df, logger)

    out_path = workspace / "Output" / "asset_deliverables.xlsx"
    write_excel(deliverables_df, out_path, "Asset_Deliverables", logger)

    result = {"status": "success", "rows": len(deliverables_df), "output": str(out_path),
              "_pw_df": pw_df, "_l2_df": l2_df}
    logger.info("Step 2 complete: %d deliverables", len(deliverables_df))
    return result


def run_step3(
    workspace: Path,
    cfg: dict,
    target_uaid2: list[str],
    logger,
    pw_df: "pd.DataFrame | None" = None,
    l2_df: "pd.DataFrame | None" = None,
    matchup_map: "dict[str, str] | None" = None,
    acbos_matchup_map: "dict[str, str] | None" = None,
) -> dict:
    """Step 3: Classify each target UAID_2 as MPDT or ACBOS."""
    logger.info("=" * 50)
    logger.info("STEP 3: Classify Targets")
    logger.info("=" * 50)

    from data_loader.local_loader import load_pw_extract, load_l2_mapping, load_scope3
    from classifier.type_classifier import classify_targets
    from excel_generator.writer import (
        write_classification_output,
        add_skip_comments_to_asset_deliverables,
        add_mapping_method_to_asset_deliverables,
    )

    if pw_df is None:
        pw_df = load_pw_extract(workspace, cfg, logger)
    if l2_df is None:
        l2_df = load_l2_mapping(workspace, cfg, logger)
    scope3_df = load_scope3(workspace, cfg, logger)

    result = classify_targets(
        workspace, target_uaid2, pw_df, l2_df, scope3_df, cfg, logger,
        matchup_map=matchup_map or {},
    )
    result = _apply_matchup_filenames_to_classification(result, matchup_map, acbos_matchup_map)

    # Write outputs
    write_classification_output(
        result,
        workspace / "Output" / "classification_output.xlsx",
        logger,
    )
    
    # Add comments to asset_deliverables for skipped UAIDs
    skipped_uaids = result.get("skipped", [])
    deliverables_path = workspace / "Output" / "asset_deliverables.xlsx"
    add_mapping_method_to_asset_deliverables(deliverables_path, result.get("csv_records", []), logger)
    if skipped_uaids:
        add_skip_comments_to_asset_deliverables(deliverables_path, skipped_uaids, logger)
    
    write_json(workspace / "Output" / "classification_plan.json", result)

    logger.info(
        "Step 3 complete — MPDT:%d ACBOS:%d conflicts:%d skipped:%d",
        len(result["mpdt_targets"]),
        len(result["acbos_targets"]),
        len(result["conflicts"]),
        len(result["skipped"]),
    )
    return result


def run_step4(workspace: Path, cfg: dict, classification: dict, output_dir: Path, sources: dict, logger) -> dict:
    """Step 4: Generate MPDT files in batches."""
    logger.info("=" * 50)
    logger.info("STEP 4: Generate MPDT Files")
    logger.info("=" * 50)

    from mpdt_generator.generator import generate_mpdt_batch

    mpdt_targets = classification.get("mpdt_targets", [])
    if not mpdt_targets:
        logger.info("No MPDT targets — skipping.")
        return {"status": "success", "generated": [], "errors": []}

    batch_size = cfg.get("batch_size", 5)
    all_generated, all_errors = [], []

    for i in range(0, len(mpdt_targets), batch_size):
        batch = mpdt_targets[i:i + batch_size]
        batch_num = (i // batch_size) + 1
        total_batches = (len(mpdt_targets) + batch_size - 1) // batch_size
        logger.info("--- MPDT Batch %d/%d (%d targets) ---", batch_num, total_batches, len(batch))

        result = generate_mpdt_batch(workspace, cfg, batch, output_dir, sources, logger)
        all_generated.extend(result["generated"])
        all_errors.extend(result["errors"])

    final = {"status": "success" if not all_errors else "partial", "generated": all_generated, "errors": all_errors}
    write_json(workspace / "Output" / "mpdt_result.json", final)
    logger.info("Step 4 complete — generated:%d errors:%d", len(all_generated), len(all_errors))
    return final


def run_step5(workspace: Path, cfg: dict, classification: dict, output_dir: Path, sources: dict, logger) -> dict:
    """Step 5: Generate ACBOS files in batches."""
    logger.info("=" * 50)
    logger.info("STEP 5: Generate ACBOS Files")
    logger.info("=" * 50)

    from mpdt_generator.acbos_generator import generate_acbos_batch

    acbos_targets = classification.get("acbos_targets", [])
    if not acbos_targets:
        logger.info("No ACBOS targets — skipping.")
        return {"status": "success", "generated": [], "errors": []}

    batch_size = cfg.get("batch_size", 5)
    all_generated, all_errors = [], []

    for i in range(0, len(acbos_targets), batch_size):
        batch = acbos_targets[i:i + batch_size]
        batch_num = (i // batch_size) + 1
        total_batches = (len(acbos_targets) + batch_size - 1) // batch_size
        logger.info("--- ACBOS Batch %d/%d (%d targets) ---", batch_num, total_batches, len(batch))

        result = generate_acbos_batch(workspace, cfg, batch, output_dir, sources, logger)
        all_generated.extend(result["generated"])
        all_errors.extend(result["errors"])

    final = {"status": "success" if not all_errors else "partial", "generated": all_generated, "errors": all_errors}
    write_json(workspace / "Output" / "acbos_result.json", final)
    logger.info("Step 5 complete — generated:%d errors:%d", len(all_generated), len(all_errors))
    return final


def run_step6(workspace: Path, cfg: dict, logger) -> dict:
    """Step 6: Stage files for PW upload."""
    logger.info("=" * 50)
    logger.info("STEP 6: Stage/Upload to ProjectWise")
    logger.info("=" * 50)

    from pw_processor.pw_uploader import run_upload

    result = run_upload(workspace, cfg, logger, check_version=cfg.get("check_version", True))
    write_json(workspace / "Output" / "upload_result.json", result)
    logger.info("Step 6 complete: %s", result.get("status"))
    return result


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    args = parse_args()

    workspace = resolve_workspace(args.workspace)
    cfg = load_config(workspace)

    # Apply CLI overrides
    _matchup_map: dict[str, str] = {}
    _acbos_matchup_map: dict[str, str] = {}
    _config_target_uaids = [
        u.strip()
        for u in cfg.get("target_uaid2", [])
        if str(u).strip()
    ]
    if args.target_uaid2:
        uaids = []
        for v in args.target_uaid2:
            uaids.extend(u.strip() for u in v.split(",") if u.strip())
        cfg["target_uaid2"] = uaids
        # Still load filename mapping from matchup file if configured, so
        # individual UAIDs tested via CLI also get the correct document names.
        if cfg.get("paths", {}).get("matchup_file"):
            _mf = (workspace / cfg["paths"]["matchup_file"]).resolve()
            _, _matchup_map, _acbos_matchup_map = _load_matchup_file_full(_mf)
    elif cfg.get("paths", {}).get("matchup_file"):
        # Matchup file is the primary source: provides both UAIDs and MPDT filenames.
        _mf = (workspace / cfg["paths"]["matchup_file"]).resolve()
        _loaded_uaids, _matchup_map, _acbos_matchup_map = _load_matchup_file_full(_mf)
        _merged_uaids = list(_loaded_uaids)
        for _uaid in _config_target_uaids:
            if _uaid not in _merged_uaids:
                _merged_uaids.append(_uaid)
        cfg["target_uaid2"] = _merged_uaids
    elif cfg.get("paths", {}).get("target_uaid2_file"):
        # Legacy: UAIDs-only file with no filename mapping.
        _uaid_file = (workspace / cfg["paths"]["target_uaid2_file"]).resolve()
        _loaded_uaids = _load_uaids_from_file(_uaid_file)
        _merged_uaids = list(_loaded_uaids)
        for _uaid in _config_target_uaids:
            if _uaid not in _merged_uaids:
                _merged_uaids.append(_uaid)
        cfg["target_uaid2"] = _merged_uaids

    if args.file_type:
        cfg["file_type"] = args.file_type
    # --refresh-sources / --no-cache CLI flags force Step 1 to fetch from external sources.
    # Otherwise fetch_external.* config flags in pipeline_config.json control it.
    refresh_sources = args.refresh_sources or args.no_cache
    if refresh_sources:
        cfg["data_source_mode"] = "external"
    if args.log_level:
        cfg["log_level"] = args.log_level
    if args.batch_size:
        cfg["batch_size"] = args.batch_size

    # Setup
    (workspace / "Output").mkdir(parents=True, exist_ok=True)
    logger = setup_logger(workspace, "pipeline", cfg.get("log_level", "INFO"))

    logger.info("=" * 60)
    logger.info("ACBOS/MPDT Pipeline V2 — %s", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    logger.info("Workspace: %s", workspace)
    logger.info("Step: %s | Batch size: %d", args.step, cfg.get("batch_size", 5))
    logger.info("Data source: %s", "FORCED REFRESH (--refresh-sources)" if refresh_sources else "local Input/ files (set fetch_external.* in config to fetch from external)")
    logger.info("Target UAIDs: %s", cfg.get("target_uaid2", []))
    logger.info("=" * 60)

    target_uaid2 = cfg.get("target_uaid2", [])
    step = args.step
    pipeline_start = time.time()
    summary: dict = {"steps": {}}

    # If only --refresh-sources is used without --step, run step1 only
    if refresh_sources and step == "all" and not target_uaid2:
        step = "step1"

    # Validate we have targets for steps that need them
    if step in ("all", "step3", "step4", "step5") and not target_uaid2:
        logger.error("No target_uaid2 specified. Use --target-uaid2.")
        sys.exit(1)

    # Create output directory for generated files
    output_dir = timestamped_dir(workspace, "Output") if step in ("all", "step4", "step5") else workspace / "Output"

    try:
        # Shared DataFrames — loaded once when step2 precedes step3 to avoid double I/O.
        _shared_pw_df: pd.DataFrame | None = None
        _shared_l2_df: pd.DataFrame | None = None

        # Step 1: Data caching / refresh
        if step in ("all", "step1"):
            t0 = time.time()
            result = run_step1(workspace, cfg, logger.getChild("step1"), force_refresh=refresh_sources)
            summary["steps"]["step1"] = result
            logger.info("  [step1: %.1fs]", time.time() - t0)

        # Step 2: Asset deliverables
        if step in ("all", "step2"):
            t0 = time.time()
            result = run_step2(workspace, cfg, logger.getChild("step2"))
            # Stash loaded DataFrames so step3 can reuse them without re-reading files.
            _shared_pw_df = result.pop("_pw_df", None)
            _shared_l2_df = result.pop("_l2_df", None)
            summary["steps"]["step2"] = result
            logger.info("  [step2: %.1fs]", time.time() - t0)

        # Step 3: Classification
        classification = {}
        if step in ("all", "step3"):
            t0 = time.time()
            classification = run_step3(
                workspace, cfg, target_uaid2,
                logger.getChild("step3"),
                pw_df=_shared_pw_df,
                l2_df=_shared_l2_df,
                matchup_map=_matchup_map,
                acbos_matchup_map=_acbos_matchup_map,
            )
            summary["steps"]["step3"] = {
                "mpdt_count": len(classification.get("mpdt_targets", [])),
                "acbos_count": len(classification.get("acbos_targets", [])),
                "conflicts": len(classification.get("conflicts", [])),
                "skipped": len(classification.get("skipped", [])),
            }
            logger.info("  [step3: %.1fs]", time.time() - t0)
        elif step in ("step4", "step5"):
            # Load previous classification
            from utils.common import read_json
            classification = read_json(workspace / "Output" / "classification_plan.json", {})
            if not classification:
                logger.error("No classification_plan.json found. Run step3 first.")
                sys.exit(1)

        # Load sources for step 4/5
        sources = {}
        if step in ("all", "step4", "step5"):
            from data_loader.local_loader import load_all_sources
            sources = load_all_sources(workspace, cfg, logger.getChild("loader"), target_uaids=target_uaid2)

        # Step 4: Generate MPDT
        if step in ("all", "step4"):
            t0 = time.time()
            result = run_step4(workspace, cfg, classification, output_dir, sources, logger.getChild("step4"))
            summary["steps"]["step4"] = {
                "generated": len(result.get("generated", [])),
                "errors": len(result.get("errors", [])),
            }
            logger.info("  [step4: %.1fs]", time.time() - t0)

        # Step 5: Generate ACBOS
        if step in ("all", "step5"):
            t0 = time.time()
            result = run_step5(workspace, cfg, classification, output_dir, sources, logger.getChild("step5"))
            summary["steps"]["step5"] = {
                "generated": len(result.get("generated", [])),
                "errors": len(result.get("errors", [])),
            }
            logger.info("  [step5: %.1fs]", time.time() - t0)

        # Step 6: Upload
        if step in ("all", "step6"):
            t0 = time.time()
            result = run_step6(workspace, cfg, logger.getChild("step6"))
            summary["steps"]["step6"] = result
            logger.info("  [step6: %.1fs]", time.time() - t0)

    except Exception as exc:
        logger.error("Pipeline FAILED: %s", exc, exc_info=True)
        summary["error"] = str(exc)
        summary["status"] = "failed"
        write_json(workspace / "Output" / "pipeline_summary.json", summary)
        sys.exit(1)

    # Final summary
    elapsed = time.time() - pipeline_start
    summary["elapsed_seconds"] = round(elapsed, 1)
    summary["status"] = "success"
    summary["output_dir"] = str(output_dir)
    write_json(workspace / "Output" / "pipeline_summary.json", summary)

    logger.info("=" * 60)
    logger.info("Pipeline complete in %.1fs", elapsed)
    logger.info("Output: %s", output_dir)
    logger.info("Summary: Output/pipeline_summary.json")
    logger.info("=" * 60)


if __name__ == "__main__":
    main()
