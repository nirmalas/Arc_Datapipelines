"""
pw_processor/pw_uploader.py — Stage generated files and upload to ProjectWise.

Can be run standalone:
  python -m pw_processor.pw_uploader --workspace .
"""
from __future__ import annotations

import argparse
import logging
import re
import os
import shutil
from datetime import datetime
from pathlib import Path

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


def _next_revision(current: str | None) -> str:
    """Increment revision: 'P01' -> 'P02', 'C03' -> 'C04'."""
    from pw_processor.pw_metadata_generator import next_revision
    return next_revision(current)


def check_versions(
    files_to_upload: list[dict],
    pw_df: pd.DataFrame,
    logger: logging.Logger,
) -> list[dict]:
    """Compare expected revision with PW. Adds revision metadata."""
    if pw_df.empty:
        logger.warning("PW extract not available — skipping version check.")
        for r in files_to_upload:
            r["revision_ok"] = True
        return files_to_upload

    doc_col = next((c for c in pw_df.columns if normalize_text(c) == "documentname"), None)
    rev_col = next((c for c in pw_df.columns if normalize_text(c) in ("version", "revision")), None)

    for record in files_to_upload:
        file_name = Path(record["file"]).stem
        norm_name = strip_acbos_suffix(file_name)
        if doc_col and rev_col:
            pw_row = pw_df[
                pw_df[doc_col].fillna("").apply(strip_acbos_suffix).str.lower() == norm_name.lower()
            ]
            if not pw_row.empty:
                latest_rev = pw_row[rev_col].dropna().max()
                expected = _next_revision(latest_rev)
                record["current_pw_revision"] = str(latest_rev)
                record["expected_revision"] = expected
                record["revision_ok"] = True
                logger.info("  %s — current: %s → next: %s", file_name, latest_rev, expected)
            else:
                record["current_pw_revision"] = None
                record["expected_revision"] = "P01"
                record["revision_ok"] = True
                logger.info("  %s — new document (P01)", file_name)
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


def upload_files(
    staged: list[dict],
    workspace: Path,
    cfg: dict,
    logger: logging.Logger,
) -> tuple[list[dict], list[dict]]:
    """Upload staged files to ProjectWise via PowerShell script."""
    pw_cfg = cfg.get("pw", {})
    upload_ps1 = workspace / pw_cfg.get("ps1_upload", "Scripts/PWPS_Upload_PW.ps1")
    if not upload_ps1.exists():
        logger.error("Upload PS1 not found: %s", upload_ps1)
        return [], [{"error": f"Upload script not found: {upload_ps1}", "file": ""}]

    datasource = pw_cfg.get("datasource", "")
    username = pw_cfg.get("username", "")
    password = pw_cfg.get("password", "") or os.environ.get("PW_PASSWORD", "")
    pw_shell = pw_cfg.get("powershell", "powershell")

    succeeded, failed = [], []
    for record in staged:
        file_path = record.get("staged_file", record.get("file", ""))
        logger.info("  Uploading: %s", Path(file_path).name)
        cmd = [
            pw_shell, "-NoProfile", "-ExecutionPolicy", "Bypass",
            "-File", str(upload_ps1),
            "-FilePath", file_path,
            "-DatasourceName", datasource,
            "-UserName", username,
        ]
        if password:
            cmd.extend(["-Password", password])
        rc = stream_subprocess(cmd, cwd=workspace, logger=logger)
        if rc == 0:
            succeeded.append(record)
            logger.info("  Upload OK: %s", Path(file_path).name)
        else:
            record["upload_error"] = f"PowerShell exited with code {rc}"
            failed.append(record)
            logger.error("  Upload FAILED (rc=%d): %s", rc, Path(file_path).name)

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

    logger.info("Files to upload: %d", len(files_to_upload))

    pw_df = pd.DataFrame()
    if check_version or cfg.get("pw", {}).get("generate_upload_metadata", True):
        from data_loader.local_loader import load_pw_extract
        pw_df = load_pw_extract(workspace, cfg, logger)

    # Version check
    if check_version:
        files_to_upload = check_versions(files_to_upload, pw_df, logger)

    # Stage
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

    # Upload
    if cfg.get("publish", False):
        succeeded, failed = upload_files(staged, workspace, cfg, logger)
        return {
            "status": "success" if not failed else "partial",
            "uploaded": succeeded,
            "failed": failed,
            "staged": [r["staged_file"] for r in staged],
            "staging_dir": str(staging_dir),
            "metadata_workbook": str(metadata_path) if metadata_path else None,
        }
    else:
        logger.info("publish=false in config — files staged but NOT uploaded.")
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
