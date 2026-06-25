"""
classifier/type_classifier.py — Classify each target UAID_2 as MPDT or ACBOS.

Classification logic (priority order):
  1. Manual override from config (uaid_type_map)
  2. Global --file-type override
  3. AUTO classification:
     a. PW extract — finds latest-version document; extension decides type
        (.acbos → ACBOS, .xlsm/.xlsx → MPDT)
     b. L2 UAID-ACBOS mapping (ACBOS column populated → ACBOS)
     If both agree → use that type (status: OK)
     If disagree  → flag as CONFLICT (do not assume)
     If only one  → use it (status: PARTIAL)
     If neither   → NEEDS_REVIEW

Can be run standalone:
  python -m classifier.type_classifier --target-uaid2 HS2-000001416 HS2-00002NSXW
"""
from __future__ import annotations

import argparse
import csv
import logging
import re
from pathlib import Path

import pandas as pd

from utils.common import (
    load_config,
    normalize_text,
    resolve_workspace,
    setup_logger,
    strip_acbos_suffix,
    write_json,
)


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

_ACBOS_EXT = {".acbos"}
_MPDT_EXT = {".xlsm", ".xls", ".xlsx"}


def _pick_col(df: pd.DataFrame, *candidates: str) -> str | None:
    norm_map = {normalize_text(c): c for c in df.columns}
    for c in candidates:
        k = normalize_text(c)
        if k in norm_map:
            return norm_map[k]
    return None


def _split_list_cell(v: str) -> list[str]:
    if v is None:
        return []
    s = str(v)
    if not s.strip():
        return []
    parts = re.split(r"[;,]", s)
    return [p.strip() for p in parts if p and p.strip()]


def build_consolidated_works_tracker(workspace: Path, cfg: dict, logger: logging.Logger) -> pd.DataFrame:
    """Combine tracker files into Input/works_tracker.xlsx and Input/consolidated_works_tracker.xlsx."""
    paths_cfg = cfg.get("paths", {})
    tracker_keys = ["tracker_sep2025", "tracker_2025", "tracker_2024"]
    sheet_norm_hints = ["central team sep2025", "central team", "central_team", "central"]

    frames: list[pd.DataFrame] = []
    for key in tracker_keys:
        rel = paths_cfg.get(key)
        if not rel:
            continue
        p = (workspace / rel).resolve()
        if not p.exists():
            continue
        try:
            xl = pd.ExcelFile(p, engine="openpyxl")
            pick = None
            for s in xl.sheet_names:
                ns = normalize_text(s)
                if any(h in ns for h in sheet_norm_hints):
                    pick = s
                    break
            if pick is None:
                pick = xl.sheet_names[0]
            df = pd.read_excel(p, sheet_name=pick, dtype=str, engine="openpyxl")
            # Normalize key column names for later matching.
            rename = {}
            for c in df.columns:
                nc = normalize_text(c)
                if nc in ("uaid 2", "uaid2", "asset id"):
                    rename[c] = "UAID_2"
                elif nc in ("description", "asset name", "assetname"):
                    rename[c] = "Description"
                elif nc in ("pw document name", "document name", "documentname", "deliverable name", "deliverablename"):
                    rename[c] = "PW_Document_Name"
            if rename:
                df = df.rename(columns=rename)
            df["_tracker_source"] = key
            frames.append(df)
            logger.info("  Tracker loaded: %s (sheet=%s, rows=%d)", p.name, pick, len(df))
        except Exception as exc:
            logger.warning("Could not read tracker '%s': %s", p.name, exc)

    if not frames:
        return pd.DataFrame()

    combined = pd.concat(frames, ignore_index=True)
    if "UAID_2" in combined.columns:
        combined["UAID_2"] = combined["UAID_2"].fillna("").astype(str).str.strip()

    out1 = workspace / "Input" / "works_tracker.xlsx"
    out2 = workspace / "Input" / "consolidated_works_tracker.xlsx"
    out1.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out1, engine="openpyxl") as w:
        combined.to_excel(w, sheet_name="works_tracker", index=False)
    with pd.ExcelWriter(out2, engine="openpyxl") as w:
        combined.to_excel(w, sheet_name="consolidated_works_tracker", index=False)
    logger.info("Consolidated works tracker written: %s and %s", out1.name, out2.name)
    return combined


# ---------------------------------------------------------------------------
# PW document lookup
# ---------------------------------------------------------------------------

def _parse_filedate(series: pd.Series) -> pd.Series:
    """Parse FileUpdated — handles Excel serial integers and ISO date strings."""
    def _parse_one(v):
        if pd.isna(v) or str(v).strip() in ("", "nan", "None"):
            return pd.NaT
        s = str(v).strip()
        if s.replace(".", "", 1).isdigit():
            try:
                return pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(s))
            except Exception:
                return pd.NaT
        return pd.to_datetime(s, errors="coerce")
    return series.apply(_parse_one)


def find_pw_best(
    uaid: str,
    pw_df: pd.DataFrame,
    input_description: str,
    acbos_doc: str | None,
    logger: logging.Logger,
) -> dict | None:
    """Find the controlling PW document for a UAID.

    Candidate rows come from direct UAID matches plus the ACBOS document listed
    in the L2 mapping.  The winning row is selected by TB_REV_DATE first, then
    Version/Rev as a tie-breaker.  The winning file extension determines MPDT
    versus ACBOS.
    """
    if pw_df.empty:
        return None

    pw_df = pw_df.copy()
    if "FileName_lower" not in pw_df.columns:
        pw_df["FileName_lower"] = pw_df["FileName"].fillna("").astype(str).str.lower() if "FileName" in pw_df.columns else ""

    def _add_candidates(frames: list[pd.DataFrame], hit: pd.DataFrame, method: str) -> None:
        if hit is None or hit.empty:
            return
        tmp = hit.copy()
        tmp["_match_method"] = method
        frames.append(tmp)

    def _stem(value: str) -> str:
        return Path(str(value or "").strip()).stem.upper()

    candidate_frames: list[pd.DataFrame] = []
    ukey = str(uaid).strip().upper()

    asset_id_col = _pick_col(pw_df, "ASSET_ID", "asset_id")
    if asset_id_col:
        hit = pw_df[pw_df[asset_id_col].fillna("").astype(str).str.strip().str.upper() == ukey]
        _add_candidates(candidate_frames, hit, "asset_id")

    tb_l2_col = _pick_col(pw_df, "TB_LEVEL_2_ID", "tb_level_2_id", "tb level 2 id", "ZZ_LBL_UAIDL2")
    if tb_l2_col:
        mask = pw_df[tb_l2_col].fillna("").astype(str).apply(
            lambda v: ukey in {x.strip().upper() for x in _split_list_cell(v)}
        )
        _add_candidates(candidate_frames, pw_df[mask], "tb_level_2_id")

    if input_description:
        desc_col = _pick_col(pw_df, "Description", "description")
        if desc_col:
            dkey = str(input_description).strip().upper()
            hit = pw_df[pw_df[desc_col].fillna("").astype(str).str.strip().str.upper() == dkey]
            _add_candidates(candidate_frames, hit, "description_exact")

    if acbos_doc:
        doc_stem = _stem(acbos_doc)
        doc_cols = [c for c in ("DocumentName", "FileName") if c in pw_df.columns]
        if doc_stem and doc_cols:
            mask = pd.Series(False, index=pw_df.index)
            for col in doc_cols:
                mask = mask | (pw_df[col].fillna("").astype(str).map(_stem) == doc_stem)
            _add_candidates(candidate_frames, pw_df[mask], "l2_acbos_doc")

    if not candidate_frames:
        return None

    all_rows = pd.concat(candidate_frames, ignore_index=True)
    if all_rows.empty:
        return None

    # Preserve the first match method for duplicate document/version rows.
    dedupe_cols = [c for c in ("DocumentName", "FileName", "Version", "FullPath") if c in all_rows.columns]
    all_rows = all_rows.drop_duplicates(subset=dedupe_cols or None, keep="first")

    fn_col = "FileName_lower" if "FileName_lower" in all_rows.columns else None
    if fn_col is None:
        return None

    valid_mask = all_rows[fn_col].str.endswith(".acbos") | all_rows[fn_col].apply(
        lambda x: any(str(x).endswith(ext) for ext in (".xlsm", ".xls", ".xlsx"))
    )
    all_rows = all_rows[valid_mask].copy()
    if all_rows.empty:
        return None

    if "TB_REV_DATE" in all_rows.columns:
        all_rows["_sort_date"] = _parse_filedate(all_rows["TB_REV_DATE"].astype(str))
    else:
        all_rows["_sort_date"] = pd.NaT
    if "FileUpdated" in all_rows.columns:
        all_rows["_file_updated_date"] = _parse_filedate(all_rows["FileUpdated"].astype(str))
        all_rows["_sort_date"] = all_rows["_sort_date"].fillna(all_rows["_file_updated_date"])
    else:
        all_rows["_file_updated_date"] = pd.NaT

    def _version_parts(value: object) -> tuple[int, int]:
        text = str(value or "").upper().strip()
        nums = re.findall(r"\d+", text)
        major = int(nums[0]) if nums else -1
        minor = int(nums[1]) if len(nums) > 1 else 0
        return major, minor

    versions = all_rows.get("Version", pd.Series([""] * len(all_rows))).map(_version_parts)
    all_rows["_version_major"] = versions.map(lambda x: x[0])
    all_rows["_version_minor"] = versions.map(lambda x: x[1])

    all_rows = all_rows.sort_values(
        ["_sort_date", "_version_major", "_version_minor"],
        ascending=[False, False, False],
        na_position="last",
    )
    best = all_rows.iloc[0]
    filename_lower = str(best.get("FileName_lower", "")).lower()
    winner = "ACBOS" if filename_lower.endswith(".acbos") else "MPDT"

    acbos_count = int(all_rows[fn_col].str.endswith(".acbos").sum())
    mpdt_count = int(all_rows[fn_col].apply(lambda x: any(str(x).endswith(ext) for ext in (".xlsm", ".xls", ".xlsx"))).sum())
    methods_used = [m for m in all_rows.get("_match_method", pd.Series(dtype=str)).dropna().astype(str).unique().tolist() if m]

    return {
        "pw_type": winner,
        "filename": str(best.get("FileName", "")).strip(),
        "filepath": str(best.get("FullPath", "")).strip(),
        "version": str(best.get("Version", "")).strip(),
        "file_date": str(best.get("FileUpdated", "")).strip(),
        "rev_date": str(best.get("TB_REV_DATE", "")).strip(),
        "match_methods": methods_used,
        "deliverable_mapped_by": str(best.get("_match_method", "")).strip() or (methods_used[0] if methods_used else ""),
        "acbos_count": acbos_count,
        "mpdt_count": mpdt_count,
    }


# ---------------------------------------------------------------------------
# L2 mapping lookup
# ---------------------------------------------------------------------------

def get_l2_info(uaid: str, l2_df: pd.DataFrame) -> dict:
    """
    Check L2 UAID-ACBOS mapping.
    ACBOS column populated → ACBOS
    Empty → NEEDS_REVIEW
    Not in file → NEEDS_REVIEW
    """
    if l2_df.empty or "UAID_2" not in l2_df.columns:
        return {"l2_type": "NEEDS_REVIEW", "asset_name": "", "acbos_doc": None, "in_l2": False}

    row = l2_df[l2_df["UAID_2"].fillna("").str.strip() == uaid]
    if row.empty:
        return {"l2_type": "NEEDS_REVIEW", "asset_name": "", "acbos_doc": None, "in_l2": False}

    r = row.iloc[0]
    asset_name = str(r.get("Asset_Name", "")).strip()
    raw_doc = str(r.get("ACBOS_Doc", "")).strip()
    acbos_doc = raw_doc if raw_doc and raw_doc.lower() not in ("nan", "none", "") else None

    return {
        "l2_type": "ACBOS" if acbos_doc else "NEEDS_REVIEW",
        "asset_name": asset_name,
        "acbos_doc": acbos_doc,
        "in_l2": True,
    }


# ---------------------------------------------------------------------------
# Conflict resolution
# ---------------------------------------------------------------------------

def resolve_type(
    pw_result: dict | None,
    l2_info: dict,
    uaid: str,
    logger: logging.Logger,
) -> tuple[str, str, str]:
    """
    Resolve final type. Returns (final_type, status, notes).
    final_type: ACBOS | MPDT | CONFLICT | NEEDS_REVIEW
    status: OK | PARTIAL | CONFLICT | NEEDS_REVIEW
    """
    pw_type = pw_result["pw_type"] if pw_result else None
    l2_type = l2_info["l2_type"]
    notes_parts: list[str] = []

    if pw_result:
        notes_parts.append(f"PW={pw_type}(methods={pw_result.get('match_methods',[])})")
    else:
        notes_parts.append("PW=not_found")

    if l2_type == "ACBOS":
        notes_parts.append(f"L2=ACBOS(doc='{l2_info.get('acbos_doc','')}')")
    elif l2_info.get("in_l2"):
        notes_parts.append("L2=in_file_no_acbos_doc")
    else:
        notes_parts.append("L2=not_in_mapping")

    notes = "; ".join(notes_parts)

    # PW found something
    if pw_type == "ACBOS":
        if l2_type == "ACBOS":
            return "ACBOS", "OK", notes
        return "ACBOS", "PARTIAL", notes

    if pw_type == "MPDT":
        # PW has absolute priority: MPDT wins regardless of what L2 says
        if l2_type == "ACBOS":
            logger.info("  PW=MPDT overrides L2=ACBOS (PW has priority) for %s", uaid)
        return "MPDT", "OK", notes

    # PW not found
    if l2_type == "ACBOS":
        return "ACBOS", "PARTIAL", notes

    return "NEEDS_REVIEW", "NEEDS_REVIEW", notes


# ---------------------------------------------------------------------------
# Main classification entry point
# ---------------------------------------------------------------------------

def classify_targets(
    workspace: Path,
    target_uaid2: list[str],
    pw_df: pd.DataFrame,
    l2_df: pd.DataFrame,
    scope3_df: pd.DataFrame,
    cfg: dict,
    logger: logging.Logger,
    matchup_map: "dict[str, str] | None" = None,
) -> dict:
    """
    Classify each target UAID_2 as MPDT or ACBOS.

    matchup_map (optional): retained for backward-compatible call sites. Matchup
    files provide output filenames after classification; they do not decide the
    MPDT/ACBOS type here. AUTO classification priority is:
      1. ProjectWise extract, preferably Input/ACBOS MPDT_FULLColumns.xlsx.
      2. missing_in_pw_file_type when the UAID is not present in PW.
      3. filename enrichment from matchup files in main.py.

    Returns dict with:
      mpdt_targets: [{uaid, file, status, notes}]
      acbos_targets: [{uaid, file, status, notes}]
      conflicts: [{uaid, status, notes}]
      skipped: [{uaid, status, notes}]
    """
    # Matchup filenames are applied after type classification in main.py.
    _ = matchup_map
    manual_map: dict[str, str] = cfg.get("uaid_type_map", {})
    not_in_pw_map: dict[str, str] = {
        k: str(v).strip().upper()
        for k, v in cfg.get("not_in_pw_type_map", {}).items()
        if str(v).strip().upper() in {"ACBOS", "MPDT"}
    }
    fallback_rows: list[dict] = cfg.get("fallback_asset_type_map", [])
    fallback_map: dict[str, dict] = {}
    for row in fallback_rows:
        if not isinstance(row, dict):
            continue
        asset_id = str(row.get("asset_id", "")).strip()
        if not asset_id:
            continue
        fallback_map[asset_id] = {
            "type": str(row.get("type", "")).strip().upper(),
            "deliverablename": str(row.get("deliverablename", "")).strip(),
        }
    file_type_override: str = cfg.get("file_type", "AUTO").upper()
    missing_in_pw_type: str = str(cfg.get("missing_in_pw_file_type", "AUTO")).strip().upper()

    allow_scope3_fallback = bool(cfg.get("allow_scope3_fallback_when_pw_missing", True))

    # Build set of available UAID_2 values in scope3 for fallback classification.
    scope3_uaid2_set: set[str] = set()
    if not scope3_df.empty:
        _u2_norms = {normalize_text(x) for x in ("UAID_2", "uaid_2", "uaid2", "ParentUaid", "Parent_UAID")}
        scope3_u2_col = next(
            (c for c in scope3_df.columns if normalize_text(c) in _u2_norms),
            None,
        )
        if scope3_u2_col:
            scope3_uaid2_set = set(
                scope3_df[scope3_u2_col]
                .fillna("")
                .astype(str)
                .str.strip()
                .tolist()
            )

    mpdt_targets: list[dict] = []
    acbos_targets: list[dict] = []
    conflicts: list[dict] = []
    skipped: list[dict] = []
    csv_records: list[dict] = []
    tracker_df = build_consolidated_works_tracker(workspace, cfg, logger)

    for uaid in target_uaid2:
        uaid = uaid.strip()
        logger.info("--- Classifying UAID_2: %s ---", uaid)

        l2_info = get_l2_info(uaid, l2_df)
        asset_name = l2_info["asset_name"] or uaid
        acbos_doc = l2_info.get("acbos_doc")

        pw_rec = find_pw_best(uaid, pw_df, asset_name, acbos_doc, logger)

        filename = pw_rec["filename"] if pw_rec else ""
        filepath = pw_rec["filepath"] if pw_rec else ""
        version = pw_rec["version"] if pw_rec else ""
        file_date = pw_rec["file_date"] if pw_rec else ""
        deliverable_mapped_by = pw_rec.get("deliverable_mapped_by", "") if pw_rec else ""

        # If PW matching failed, use consolidated works tracker.
        if not filename and not tracker_df.empty:
            tr_hit = pd.DataFrame()
            if "UAID_2" in tracker_df.columns:
                tr_hit = tracker_df[tracker_df["UAID_2"].fillna("").astype(str).str.strip().str.upper() == uaid.upper()]
                if not tr_hit.empty:
                    deliverable_mapped_by = "works_tracker_uaid2"
            if tr_hit.empty and asset_name and "Description" in tracker_df.columns:
                tr_hit = tracker_df[
                    tracker_df["Description"].fillna("").astype(str).str.strip().str.upper() == asset_name.upper()
                ]
                if not tr_hit.empty:
                    deliverable_mapped_by = "works_tracker_description"
            if not tr_hit.empty:
                pick = tr_hit.iloc[0]
                cand = str(pick.get("PW_Document_Name", "")).strip()
                if cand:
                    filename = cand if re.search(r"\.(xlsm|xlsx|xls)$", cand, re.IGNORECASE) else f"{cand}.xlsm"

        # Priority 1: Manual override
        if uaid in manual_map:
            ft = manual_map[uaid].upper()
            status = "OVERRIDE"
            notes = f"manual_override={ft}"
            logger.info("  Manual override: %s", ft)
        # Priority 2: Global type override
        elif file_type_override != "AUTO":
            ft = file_type_override
            status = "OVERRIDE"
            notes = f"global_override={ft}"
            logger.info("  Global override: %s", ft)
        # Priority 3: Auto-detect
        else:
            # Strict rule: if no PW ID match, do not infer from L2 alone.
            if not pw_rec:
                not_in_pw_ft = not_in_pw_map.get(uaid)
                fallback = fallback_map.get(uaid)
                if not_in_pw_ft:
                    ft = not_in_pw_ft
                    status = "NOT_IN_PW_MAP"
                    notes = f"Not in PW extract; classified via not_in_pw_type_map={ft}"
                    logger.warning(
                        "  No PW match for %s, using not_in_pw_type_map: %s", uaid, ft
                    )
                elif fallback and fallback.get("type") in {"ACBOS", "MPDT"}:
                    ft = fallback["type"]
                    status = "FALLBACK"
                    notes = (
                        f"PW=not_found; fallback_map={ft}; "
                        f"deliverablename='{fallback.get('deliverablename','')}'"
                    )
                    if not filename and fallback.get("deliverablename"):
                        filename = fallback["deliverablename"]
                    logger.warning("  No PW match for %s, using fallback map: %s", uaid, ft)
                elif missing_in_pw_type in {"ACBOS", "MPDT"}:
                    # Configurable fallback type when UAID is not present in PW extract.
                    ft = missing_in_pw_type
                    status = "FALLBACK_MISSING_PW_TYPE"
                    notes = f"Not in PW extract; forced by config missing_in_pw_file_type={ft}"
                    logger.warning(
                        "  No PW match for %s, using config missing_in_pw_file_type: %s",
                        uaid,
                        ft,
                    )
                elif allow_scope3_fallback and uaid in scope3_uaid2_set:
                    # Notebook parity: when UAID has scope3 data, still generate output
                    # even if the UAID is missing from current PW extract.
                    ft = "MPDT"
                    status = "FALLBACK_SCOPE3"
                    notes = "Not in PW extract; classified from Scope3 presence"
                    if not filename:
                        filename = f"MPDT_{uaid}.xlsm"
                    logger.warning(
                        "  No PW match for %s, but Scope3 rows exist -> fallback classify as MPDT",
                        uaid,
                    )
                else:
                    ft = "NEEDS_REVIEW"
                    status = "NOT_IN_PW_EXTRACT"
                    notes = "Not found in Input/ACBOS MPDT.xlsx — generation skipped"
                    logger.warning("  SKIP: %s not found in PW extract (Input/ACBOS MPDT.xlsx)", uaid)
            else:
                ft, status, notes = resolve_type(pw_rec, l2_info, uaid, logger)

        record = {
            "uaid": uaid,
            "outputfiletype": ft,
            "filename": filename,
            "filepath": filepath,
            "version": version,
            "file_date": file_date,
            "asset_name": asset_name,
            "acbos_doc": acbos_doc or "",
            "deliverable_mapped_by": deliverable_mapped_by,
            "status": status,
            "notes": notes,
        }
        csv_records.append(record)

        if ft == "MPDT":
            mpdt_targets.append({"uaid": uaid, "file": filename, "status": status, "notes": notes, "deliverable_mapped_by": deliverable_mapped_by})
            logger.info("  => MPDT (%s)", status)
        elif ft == "ACBOS":
            acbos_targets.append({"uaid": uaid, "file": filename, "status": status, "notes": notes, "deliverable_mapped_by": deliverable_mapped_by})
            logger.info("  => ACBOS (%s)", status)
        elif ft == "CONFLICT":
            conflicts.append({"uaid": uaid, "status": status, "notes": notes})
            logger.warning("  => CONFLICT -- requires manual review")
        else:  # NEEDS_REVIEW or NOT_IN_PW_EXTRACT
            skipped.append({"uaid": uaid, "status": status, "notes": notes})
            logger.warning("  => SKIPPED (%s)", status)

    result = {
        "mpdt_targets": mpdt_targets,
        "acbos_targets": acbos_targets,
        "conflicts": conflicts,
        "skipped": skipped,
        "csv_records": csv_records,
    }

    logger.info(
        "Classification complete — MPDT:%d ACBOS:%d conflicts:%d skipped:%d",
        len(mpdt_targets), len(acbos_targets), len(conflicts), len(skipped),
    )
    return result


def write_targets_csv(workspace: Path, records: list[dict], logger: logging.Logger) -> Path:
    """Write classification results to CSV."""
    out_dir = workspace / "Output"
    out_dir.mkdir(parents=True, exist_ok=True)
    csv_path = out_dir / "classification_targets.csv"
    fieldnames = [
        "uaid", "outputfiletype", "filename", "filepath",
        "version", "file_date", "asset_name", "acbos_doc", "deliverable_mapped_by", "status", "notes",
    ]
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(records)
    logger.info("Targets CSV written: %s", csv_path)
    return csv_path


# ---------------------------------------------------------------------------
# Standalone entry point
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description="Step 3 — Classify UAID_2 as MPDT or ACBOS")
    parser.add_argument("--workspace", default=None)
    parser.add_argument("--target-uaid2", nargs="+", required=True)
    args = parser.parse_args()

    workspace = resolve_workspace(args.workspace)
    cfg = load_config(workspace)
    logger = setup_logger(workspace, "classifier", cfg.get("log_level", "INFO"))

    # Load required data
    from data_loader.local_loader import load_pw_extract, load_l2_mapping, load_scope3
    pw_df = load_pw_extract(workspace, cfg, logger)
    l2_df = load_l2_mapping(workspace, cfg, logger)
    scope3_df = load_scope3(workspace, cfg, logger)

    uaids = []
    for v in args.target_uaid2:
        uaids.extend(u.strip() for u in v.split(",") if u.strip())

    result = classify_targets(workspace, uaids, pw_df, l2_df, scope3_df, cfg, logger)
    write_targets_csv(workspace, result["csv_records"], logger)
    write_json(workspace / "Output" / "classification_plan.json", result)
    logger.info("=== Classification complete ===")


if __name__ == "__main__":
    main()
