"""
pw_processor/pw_uploader.py - Stage generated files and upload to ProjectWise.

Can be run standalone:
  python -m pw_processor.pw_uploader --workspace .
"""
from __future__ import annotations

import argparse
import json
import logging
import os
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd

from utils.common import (
    load_config,
    normalize_text,
    read_json,
    resolve_workspace,
    setup_logger,
    stream_subprocess,
    strip_acbos_suffix,
    write_json,
)


_CORE_METADATA_COLUMNS = {
    "documentname",
    "description",
    "version",
    "revision",
    "rev",
    "filename",
    "file name",
    "filetype",
    "file type",
    "fullpath",
    "urn",
    "pwfolderpath",
    "pw folder path",
    "folderpath",
    "folder path",
    "localfilepath",
    "local file path",
    "filepath",
    "file path",
    "stagedfile",
    "sourcefile",
    "extension",
    "currentpwrevision",
    "currentrevision",
    "previousversion",
    "nextversion",
    "expectedrevision",
}


def _as_text(value: Any) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    return "" if s.lower() == "nan" else s


def _norm(value: Any) -> str:
    return normalize_text(value).replace("_", " ").strip().lower()


def _doc_key(name: Any) -> str:
    s = _as_text(name)
    s = Path(s).stem if s else ""
    return strip_acbos_suffix(s).strip().lower()


def _next_revision(current: str | None) -> str:
    from pw_processor.pw_metadata_generator import next_revision
    return next_revision(current)


def _version_sort_key(value: Any) -> tuple[int, str, int, int, str]:
    from pw_processor.pw_metadata_generator import version_sort_key
    return version_sort_key(value)


def _find_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    by_norm = {_norm(c): c for c in df.columns}
    for candidate in candidates:
        hit = by_norm.get(_norm(candidate))
        if hit:
            return hit
    return None


def _latest_revision(rows: pd.DataFrame, rev_col: str) -> str:
    values = [_as_text(v) for v in rows[rev_col].tolist() if _as_text(v)]
    if not values:
        return ""
    return max(values, key=_version_sort_key)


def check_versions(
    files_to_upload: list[dict],
    pw_df: pd.DataFrame,
    logger: logging.Logger,
) -> list[dict]:
    """Compare expected revision with PW. Adds revision metadata."""
    if pw_df.empty:
        logger.warning("PW extract not available - skipping version check.")
        for record in files_to_upload:
            record["revision_ok"] = True
        return files_to_upload

    doc_col = _find_column(pw_df, ["DocumentName", "Document Name", "Document", "Name"])
    rev_col = _find_column(pw_df, ["Version", "Revision", "Rev"])

    for record in files_to_upload:
        file_name = Path(record["file"]).stem
        norm_name = _doc_key(file_name)
        if doc_col and rev_col:
            pw_row = pw_df[pw_df[doc_col].fillna("").apply(_doc_key) == norm_name]
            if not pw_row.empty:
                latest_rev = _latest_revision(pw_row, rev_col)
                expected = _next_revision(latest_rev)
                record["current_pw_revision"] = latest_rev
                record["expected_revision"] = expected
                record["revision_ok"] = True
                logger.info("  %s - current: %s -> next: %s", file_name, latest_rev, expected)
            else:
                record["current_pw_revision"] = None
                record["expected_revision"] = "P01"
                record["revision_ok"] = True
                logger.warning("  %s - not found in PW extract; treating as new document (P01)", file_name)
        else:
            record["revision_ok"] = True
    return files_to_upload


def stage_files(files: list[dict], staging_dir: Path, logger: logging.Logger) -> list[dict]:
    """Copy generated files to staging directory for PW upload."""
    staging_dir.mkdir(parents=True, exist_ok=True)
    staged = []
    for record in files:
        src = Path(record["file"])
        if not src.exists():
            logger.warning("  File not found, skipping: %s", src)
            continue
        dest = staging_dir / src.name
        counter = 1
        while dest.exists():
            dest = staging_dir / f"{src.stem}_{counter}{src.suffix}"
            counter += 1
        shutil.copy2(src, dest)
        record["staged_file"] = str(dest)
        staged.append(record)
        logger.info("  Staged: %s", dest.name)
    return staged


def _add_pw_arg(cmd: list[str], name: str, value: object) -> None:
    text = _as_text(value)
    if text:
        cmd.extend([name, text])


def _read_pw_extract_file(path: Path, logger: logging.Logger) -> pd.DataFrame:
    df = pd.read_excel(path, dtype=str, engine="openpyxl", keep_default_na=False)
    logger.info("PW upload metadata extract loaded from '%s': %d rows, %d columns", path.name, len(df), len(df.columns))
    return df


def load_pw_upload_extract(workspace: Path, cfg: dict, logger: logging.Logger) -> pd.DataFrame:
    """Load the richest PW extract available for version checks and upload metadata."""
    paths = cfg.get("paths", {})
    candidates = [
        paths.get("pw_extract_full_columns", ""),
        "Input/ACBOS MPDT_FULLColumns.xlsx",
        paths.get("pw_extract_full", ""),
        paths.get("pw_extract", ""),
    ]
    seen: set[str] = set()
    for rel in candidates:
        if not rel:
            continue
        path = (workspace / rel).resolve()
        key = str(path).lower()
        if key in seen:
            continue
        seen.add(key)
        if not path.exists():
            continue
        try:
            return _read_pw_extract_file(path, logger)
        except Exception as exc:
            logger.warning("Could not read PW upload metadata extract '%s': %s", path.name, exc)

    from data_loader.local_loader import load_pw_extract
    return load_pw_extract(workspace, cfg, logger)


def _metadata_row_value(row: dict[str, Any], names: list[str]) -> str:
    norm_map = {_norm(k): k for k in row.keys()}
    for name in names:
        col = norm_map.get(_norm(name))
        if col:
            return _as_text(row.get(col, ""))
    return ""


def _derive_pw_folder_from_metadata(row: dict[str, Any], document_name: str) -> str:
    explicit = _metadata_row_value(row, ["PWFolderPath", "PW Folder Path", "FolderPath", "Folder Path"])
    if explicit:
        return explicit
    full_path = _metadata_row_value(row, ["FullPath"])
    if not full_path:
        return ""
    parts = [p for p in full_path.replace("/", "\\").split("\\") if p]
    if parts and _doc_key(parts[-1]) == _doc_key(document_name):
        return "\\".join(parts[:-1])
    return full_path


def _clean_json_value(value: Any) -> str:
    if value is None:
        return ""
    if pd.isna(value):
        return ""
    return str(value)


def _metadata_attributes(row: dict[str, Any]) -> dict[str, str]:
    attrs: dict[str, str] = {}
    for key, value in row.items():
        norm_key = _norm(key)
        if norm_key in _CORE_METADATA_COLUMNS:
            continue
        attrs[str(key)] = _clean_json_value(value)
    return attrs


def _load_metadata_rows(metadata_path: Path | None, logger: logging.Logger) -> dict[str, dict[str, Any]]:
    if not metadata_path or not metadata_path.exists():
        return {}
    df = pd.read_excel(metadata_path, dtype=str, engine="openpyxl", keep_default_na=False)
    rows_by_file: dict[str, dict[str, Any]] = {}
    for _, row in df.iterrows():
        rec = row.to_dict()
        file_name = _metadata_row_value(rec, ["FileName", "File Name", "LocalFileName"])
        local_path = _metadata_row_value(rec, ["LocalFilePath", "Local File Path", "FilePath", "File Path"])
        for key in (file_name, Path(local_path).name if local_path else ""):
            if key:
                rows_by_file[key.lower()] = rec
    logger.info("Loaded staged PW upload metadata rows from %s: %d", metadata_path.name, len(rows_by_file))
    return rows_by_file


def _write_attributes_sidecar(file_path: Path, metadata_row: dict[str, Any], logger: logging.Logger) -> Path:
    attrs = _metadata_attributes(metadata_row)
    out_path = file_path.with_suffix(file_path.suffix + ".metadata.json")
    out_path.write_text(json.dumps(attrs, indent=2), encoding="utf-8")
    logger.info("  Metadata JSON written: %s", out_path.name)
    return out_path


def upload_files(
    staged: list[dict],
    workspace: Path,
    cfg: dict,
    logger: logging.Logger,
    metadata_path: Path | None = None,
) -> tuple[list[dict], list[dict]]:
    """Upload staged files to ProjectWise via PowerShell script."""
    pw_cfg = cfg.get("pw", {})
    upload_ps1 = workspace / pw_cfg.get("ps1_upload", "Scripts/PWPS_Upload_PW.ps1")
    if not upload_ps1.exists():
        logger.error("Upload PS1 not found: %s", upload_ps1)
        return [], [{"error": f"Upload script not found: {upload_ps1}", "file": ""}]

    metadata_rows = _load_metadata_rows(metadata_path, logger)
    datasource = pw_cfg.get("datasource", "")
    username = pw_cfg.get("username", "")
    password = pw_cfg.get("password", "") or os.environ.get("PW_PASSWORD", "")
    pw_shell = pw_cfg.get("powershell", "powershell")
    default_folder = (
        pw_cfg.get("default_folder")
        or pw_cfg.get("upload_folder")
        or cfg.get("pw_folder_path")
        or ""
    )
    default_application = pw_cfg.get("application", "")

    succeeded, failed = [], []
    for record in staged:
        file_path = Path(record.get("staged_file", record.get("file", "")))
        file_stem = file_path.stem
        metadata_row = metadata_rows.get(file_path.name.lower(), {})
        document_name = _metadata_row_value(metadata_row, ["DocumentName", "Document Name", "Document"]) or record.get("document_name") or record.get("doc_name") or file_stem
        description = _metadata_row_value(metadata_row, ["Description", "Document Description"]) or record.get("description") or document_name
        version = _metadata_row_value(metadata_row, ["Version", "Revision", "Rev"]) or record.get("expected_revision") or record.get("version") or ""
        folder_path = _derive_pw_folder_from_metadata(metadata_row, document_name) or record.get("pw_folder_path") or default_folder
        application = record.get("application") or default_application
        attributes_json = _write_attributes_sidecar(file_path, metadata_row, logger) if metadata_row else None

        logger.info("  Uploading: %s", file_path.name)
        cmd = [
            pw_shell, "-NoProfile", "-ExecutionPolicy", "Bypass",
            "-File", str(upload_ps1),
            "-FilePath", str(file_path),
            "-DatasourceName", datasource,
            "-UserName", username,
        ]
        _add_pw_arg(cmd, "-Password", password)
        _add_pw_arg(cmd, "-PWFolderPath", folder_path)
        _add_pw_arg(cmd, "-DocumentName", document_name)
        _add_pw_arg(cmd, "-Description", description)
        _add_pw_arg(cmd, "-Version", version)
        _add_pw_arg(cmd, "-Application", application)
        _add_pw_arg(cmd, "-ProjectWiseBin", pw_cfg.get("projectwise_bin", ""))
        _add_pw_arg(cmd, "-AttributesJson", str(attributes_json) if attributes_json else "")
        rc = stream_subprocess(cmd, cwd=workspace, logger=logger)
        if rc == 0:
            succeeded.append(record)
            logger.info("  Upload OK: %s", file_path.name)
        else:
            record["upload_error"] = f"PowerShell exited with code {rc}"
            failed.append(record)
            logger.error("  Upload FAILED (rc=%d): %s", rc, file_path.name)

    return succeeded, failed


def run_upload(
    workspace: Path,
    cfg: dict,
    logger: logging.Logger,
    check_version: bool = True,
) -> dict:
    """Stage and upload all generated files from step results."""
    mpdt_result = read_json(workspace / "Output" / "mpdt_result.json", {})
    acbos_result = read_json(workspace / "Output" / "acbos_result.json", {})

    files_to_upload: list[dict] = []
    for rec in mpdt_result.get("generated", []):
        files_to_upload.append({**rec, "file_type": "MPDT"})
    for rec in acbos_result.get("generated", []):
        files_to_upload.append({**rec, "file_type": "ACBOS"})

    if not files_to_upload:
        logger.warning("No generated files to upload.")
        return {"status": "success", "uploaded": [], "failed": [], "staged": []}

    existing_files = []
    for record in files_to_upload:
        if Path(record.get("file", "")).exists():
            existing_files.append(record)
        else:
            logger.warning("Generated result points to missing file, skipping: %s", record.get("file", ""))
    files_to_upload = existing_files
    if not files_to_upload:
        logger.warning("No existing generated files to upload.")
        return {"status": "success", "uploaded": [], "failed": [], "staged": []}

    logger.info("Files to upload: %d", len(files_to_upload))

    pw_df = pd.DataFrame()
    if check_version or cfg.get("pw", {}).get("generate_upload_metadata", True):
        pw_df = load_pw_upload_extract(workspace, cfg, logger)

    if check_version:
        files_to_upload = check_versions(files_to_upload, pw_df, logger)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    staging_dir = workspace / f"output_pw_{ts}"
    staged = stage_files(files_to_upload, staging_dir, logger)
    logger.info("Staged %d files to %s", len(staged), staging_dir)

    metadata_path = None
    if cfg.get("pw", {}).get("generate_upload_metadata", True):
        from pw_processor.pw_metadata_generator import write_pw_upload_metadata
        metadata_path = write_pw_upload_metadata(
            files_to_upload=files_to_upload,
            staged=staged,
            pw_df=pw_df,
            cfg=cfg,
            output_dir=staging_dir,
            logger=logger,
        )

    if cfg.get("publish", False):
        succeeded, failed = upload_files(staged, workspace, cfg, logger, metadata_path=metadata_path)
        return {
            "status": "success" if not failed else "partial",
            "uploaded": succeeded,
            "failed": failed,
            "staged": [r["staged_file"] for r in staged],
            "staging_dir": str(staging_dir),
            "metadata_workbook": str(metadata_path) if metadata_path else None,
        }

    logger.info("publish=false in config - files staged but NOT uploaded.")
    return {
        "status": "staged_only",
        "staged": [r["staged_file"] for r in staged],
        "staging_dir": str(staging_dir),
        "metadata_workbook": str(metadata_path) if metadata_path else None,
    }


# ---------------------------------------------------------------------------
# Standalone
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description="Stage & upload files to ProjectWise")
    parser.add_argument("--workspace", default=None)
    parser.add_argument("--no-version-check", action="store_true")
    args = parser.parse_args()

    workspace = resolve_workspace(args.workspace)
    cfg = load_config(workspace)
    logger = setup_logger(workspace, "pw_uploader", cfg.get("log_level", "INFO"))

    logger.info("=== PW Uploader ===")
    result = run_upload(workspace, cfg, logger, check_version=not args.no_version_check)
    write_json(workspace / "Output" / "upload_result.json", result)
    logger.info("=== Upload complete: %s ===", result.get("status"))


if __name__ == "__main__":
    main()


