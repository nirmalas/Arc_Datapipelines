"""
pw_processor/pw_metadata_generator.py - Build a ProjectWise upload metadata workbook.

The workbook is generated during step6 next to the staged files in output_pw_<timestamp>.
It mirrors the ProjectWise extract columns, copies existing PW metadata where possible,
sets the next revision, and shifts ProjectWise revision-history columns for upload.
"""
from __future__ import annotations

import logging
import re
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd

from utils.common import normalize_text, strip_acbos_suffix


AUTOMATION_REVISION_NOTE = "updated from L3 database through automation"

DEFAULT_METADATA_COLUMNS = [
    "DocumentName",
    "Description",
    "Version",
    "FileName",
    "FileType",
    "ASSET_ID",
    "UAID_2",
    "PWFolderPath",
    "LocalFilePath",
]

_DOCUMENT_COLUMNS = [
    "DocumentName", "Document Name", "Document", "Name", "Doc Name", "Deliverable Number",
]
_DESCRIPTION_COLUMNS = ["Description", "Document Description", "PW Description"]
_VERSION_COLUMNS = ["Version", "Revision", "Rev"]
_FILENAME_COLUMNS = ["FileName", "File Name", "Original File Name", "LocalFileName"]
_FILETYPE_COLUMNS = ["FileType", "File Type", "Type", "Document Type"]
_UAID_COLUMNS = ["ASSET_ID", "Asset ID", "UAID_2", "UAID2", "Level 2 UAID", "UAID"]
_FOLDER_COLUMNS = ["PWFolderPath", "PW Folder Path", "FolderPath", "Folder Path"]
_LOCAL_PATH_COLUMNS = ["LocalFilePath", "Local File Path", "FilePath", "File Path"]
_NOTE_COLUMNS = ["Note", "Notes", "Revision Note", "Revision Notes"]
_REVISION_HISTORY_RE = re.compile(r"^(RV_.+)_([0-9]+)$", re.IGNORECASE)


def _norm_key(value: Any) -> str:
    return normalize_text(value).replace(" ", "")


def _first_col(columns: list[str], candidates: list[str]) -> str | None:
    by_norm = {_norm_key(c): c for c in columns}
    for cand in candidates:
        hit = by_norm.get(_norm_key(cand))
        if hit:
            return hit
    return None


def _append_missing_columns(columns: list[str], extras: list[str]) -> list[str]:
    out = list(columns)
    existing = {_norm_key(c) for c in out}
    for col in extras:
        if _norm_key(col) not in existing:
            out.append(col)
            existing.add(_norm_key(col))
    return out


def _as_text(value: Any) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    return "" if s.lower() == "nan" else s


def _doc_key(name: Any) -> str:
    """Document key used for PW matching. ACBOS suffix is ignored for version lookup."""
    s = _as_text(name)
    s = Path(s).stem if s else ""
    return strip_acbos_suffix(s).strip().lower()


def version_sort_key(value: Any) -> tuple[int, str, int, int, str]:
    """Sort ProjectWise versions/revisions naturally: P02 > P01, P03.1 > P03."""
    s = _as_text(value)
    if not s:
        return (0, "", -1, -1, "")
    m = re.match(r"^([A-Za-z]+)\s*0*(\d+)(?:\.(\d+))?$", s)
    if m:
        suffix = int(m.group(3)) if m.group(3) is not None else -1
        return (2, m.group(1).upper(), int(m.group(2)), suffix, s)
    m = re.search(r"(\d+)(?:\.(\d+))?", s)
    if m:
        suffix = int(m.group(2)) if m.group(2) is not None else -1
        return (1, "", int(m.group(1)), suffix, s)
    return (0, "", -1, -1, s)


def next_revision(current: str | None) -> str:
    """Increment standard PW revisions such as P01 -> P02, P03.1 -> P03.2."""
    s = _as_text(current)
    if not s:
        return "P01"
    m = re.match(r"^([A-Za-z]+)\s*0*(\d+)(?:\.(\d+))?$", s)
    if m:
        prefix = m.group(1)
        width = max(2, len(m.group(2)))
        if m.group(3) is not None:
            return f"{prefix}{int(m.group(2)):0{width}d}.{int(m.group(3)) + 1}"
        return f"{prefix}{int(m.group(2)) + 1:0{width}d}"
    return s


def get_metadata_columns(cfg: dict, pw_df: pd.DataFrame) -> list[str]:
    """Return metadata columns, preferring the PW extract layout for ProjectWise uploads."""
    pw_cfg = cfg.get("pw", {}) if isinstance(cfg, dict) else {}
    prefer_config = bool(pw_cfg.get("prefer_config_metadata_columns", False))
    configured = (
        cfg.get("pw_upload_metadata_columns")
        or pw_cfg.get("upload_metadata_columns")
        or pw_cfg.get("metadata_columns")
    )

    if pw_df is not None and not pw_df.empty and not prefer_config:
        return _append_missing_columns([str(c) for c in pw_df.columns], DEFAULT_METADATA_COLUMNS)

    if isinstance(configured, str):
        columns = [c.strip() for c in configured.split(",") if c.strip()]
    elif isinstance(configured, list):
        columns = [str(c).strip() for c in configured if str(c).strip()]
    else:
        columns = []
    if columns:
        return _append_missing_columns(columns, DEFAULT_METADATA_COLUMNS)
    if pw_df is not None and not pw_df.empty:
        return _append_missing_columns([str(c) for c in pw_df.columns], DEFAULT_METADATA_COLUMNS)
    return DEFAULT_METADATA_COLUMNS.copy()


def _build_pw_latest_index(pw_df: pd.DataFrame) -> tuple[dict[str, dict], str | None, str | None]:
    """Index PW extract by exact document key, keeping the row with the highest version."""
    if pw_df is None or pw_df.empty:
        return {}, None, None
    columns = [str(c) for c in pw_df.columns]
    doc_col = _first_col(columns, _DOCUMENT_COLUMNS)
    rev_col = _first_col(columns, _VERSION_COLUMNS)
    if not doc_col:
        return {}, None, rev_col

    latest: dict[str, dict] = {}
    latest_key: dict[str, tuple[int, str, int, int, str]] = {}
    for _, row in pw_df.iterrows():
        key = _doc_key(row.get(doc_col, ""))
        if not key:
            continue
        sort_key = version_sort_key(row.get(rev_col, "")) if rev_col else (0, "", -1, -1, "")
        if key not in latest or sort_key > latest_key[key]:
            latest[key] = row.to_dict()
            latest_key[key] = sort_key
    return latest, doc_col, rev_col


def _set_if_col(row: dict[str, Any], columns: list[str], candidates: list[str], value: Any) -> None:
    col = _first_col(columns, candidates)
    if col:
        row[col] = value


def _set_all_matching_cols(row: dict[str, Any], columns: list[str], candidates: list[str], value: Any) -> None:
    cand_norms = {_norm_key(c) for c in candidates}
    for col in columns:
        if _norm_key(col) in cand_norms:
            row[col] = value


def _excel_date_serial(dt: datetime) -> str:
    return str((dt.date() - datetime(1899, 12, 30).date()).days)


def _derive_pw_folder_path(row: dict[str, Any], doc_name: str, default_folder: str) -> str:
    explicit = _as_text(row.get("PWFolderPath") or row.get("FolderPath") or row.get("PW Folder Path"))
    if explicit:
        return explicit
    full_path = _as_text(row.get("FullPath"))
    if full_path:
        parts = [p for p in re.split(r"[\\/]+", full_path) if p]
        if parts and _doc_key(parts[-1]) == _doc_key(doc_name):
            return "\\".join(parts[:-1])
        return full_path
    return default_folder


def _revision_history_new_values(next_rev: str, note: str, cfg: dict) -> dict[str, Any]:
    today_serial = _excel_date_serial(datetime.now())
    values = {
        "RV_V": next_rev,
        "RV_N": note,
        "RV_N2": note,
        "RV_RD": today_serial,
        "RV_DD": today_serial,
        "RV_CD": today_serial,
        "RV_AD": today_serial,
    }
    configured = cfg.get("pw", {}).get("revision_history_new_values", {}) if isinstance(cfg, dict) else {}
    if isinstance(configured, dict):
        for key, value in configured.items():
            values[str(key).strip().upper()] = value
    return values


def _shift_revision_history(row: dict[str, Any], columns: list[str], new_values: dict[str, Any]) -> None:
    groups: dict[str, dict[int, str]] = {}
    for col in columns:
        match = _REVISION_HISTORY_RE.match(str(col))
        if not match:
            continue
        prefix = match.group(1).upper()
        idx = int(match.group(2))
        groups.setdefault(prefix, {})[idx] = col

    for prefix, numbered_cols in groups.items():
        for idx in sorted(numbered_cols, reverse=True):
            col = numbered_cols[idx]
            if idx == 1:
                row[col] = new_values.get(prefix, "")
            else:
                previous_col = numbered_cols.get(idx - 1)
                row[col] = row.get(previous_col, "") if previous_col else ""


def _apply_upload_metadata_updates(
    row: dict[str, Any],
    columns: list[str],
    record: dict,
    source_file: Path,
    doc_name: str,
    current_rev: str,
    next_rev: str,
    default_folder: str,
    cfg: dict,
) -> None:
    note = cfg.get("pw", {}).get("automation_revision_note", AUTOMATION_REVISION_NOTE)
    file_name = source_file.name
    folder_path = record.get("pw_folder_path") or _derive_pw_folder_path(row, doc_name, default_folder)

    _set_all_matching_cols(row, columns, _DOCUMENT_COLUMNS, doc_name)
    _set_all_matching_cols(row, columns, _DESCRIPTION_COLUMNS, record.get("description") or row.get("Description") or doc_name)
    _set_all_matching_cols(row, columns, _VERSION_COLUMNS, next_rev)
    _set_if_col(row, columns, _FILENAME_COLUMNS, file_name)
    _set_if_col(row, columns, _FILETYPE_COLUMNS, record.get("file_type", source_file.suffix.lstrip(".").upper()))
    _set_all_matching_cols(row, columns, _UAID_COLUMNS, record.get("uaid", ""))
    _set_if_col(row, columns, _FOLDER_COLUMNS, folder_path)
    _set_if_col(row, columns, _LOCAL_PATH_COLUMNS, str(source_file))
    _set_if_col(row, columns, _NOTE_COLUMNS, note)

    new_history_values = _revision_history_new_values(next_rev, note, cfg)
    _shift_revision_history(row, columns, new_history_values)

    optional_values = {
        "currentpwrevision": current_rev or "",
        "currentrevision": current_rev or "",
        "previousversion": current_rev or "",
        "nextversion": next_rev,
        "expectedrevision": next_rev,
        "stagedfile": str(source_file),
        "sourcefile": record.get("file", ""),
        "extension": source_file.suffix.lstrip("."),
        "rvnew": next_rev,
        "rvoldpub": current_rev or "",
        "rvoldwip": current_rev or "",
    }
    for col in columns:
        key = _norm_key(col).lower()
        if key in optional_values:
            row[col] = optional_values[key]


def build_pw_upload_metadata(
    files_to_upload: list[dict],
    staged: list[dict],
    pw_df: pd.DataFrame,
    cfg: dict,
    logger: logging.Logger,
) -> pd.DataFrame:
    """Create one ProjectWise metadata row per staged generated file."""
    columns = get_metadata_columns(cfg, pw_df)
    latest_by_doc, _pw_doc_col, pw_rev_col = _build_pw_latest_index(pw_df)
    default_folder = (
        cfg.get("pw", {}).get("default_folder")
        or cfg.get("pw", {}).get("upload_folder")
        or cfg.get("pw_folder_path")
        or ""
    )

    rows: list[dict[str, Any]] = []
    for record in staged or files_to_upload:
        source_file = Path(record.get("staged_file") or record.get("file") or "")
        doc_name = source_file.stem if source_file.name else _as_text(record.get("document_name") or record.get("doc_name"))
        doc_key = _doc_key(doc_name)
        base = dict(latest_by_doc.get(doc_key, {}))
        row = {col: base.get(col, "") for col in columns}

        current_rev = record.get("current_pw_revision")
        if not _as_text(current_rev) and pw_rev_col and base:
            current_rev = base.get(pw_rev_col, "")
        current_rev = _as_text(current_rev)
        next_rev = record.get("expected_revision") or next_revision(current_rev)

        _apply_upload_metadata_updates(
            row=row,
            columns=columns,
            record=record,
            source_file=source_file,
            doc_name=doc_name,
            current_rev=current_rev,
            next_rev=next_rev,
            default_folder=default_folder,
            cfg=cfg,
        )
        if base:
            logger.info("  Metadata source: existing PW row for %s (%s -> %s)", doc_name, current_rev, next_rev)
        else:
            logger.warning("  Metadata source: no existing PW row for %s; using defaults (%s)", doc_name, next_rev)
        rows.append(row)

    logger.info("Built PW upload metadata rows: %d", len(rows))
    return pd.DataFrame(rows, columns=columns)


def write_pw_upload_metadata(
    files_to_upload: list[dict],
    staged: list[dict],
    pw_df: pd.DataFrame,
    cfg: dict,
    output_dir: Path,
    logger: logging.Logger,
) -> Path | None:
    """Write the ProjectWise upload metadata workbook in the staging folder."""
    if not staged and not files_to_upload:
        return None
    metadata_df = build_pw_upload_metadata(files_to_upload, staged, pw_df, cfg, logger)
    output_dir.mkdir(parents=True, exist_ok=True)
    filename = cfg.get("pw", {}).get("upload_metadata_filename", "PW_Upload_Metadata.xlsx")
    out_path = output_dir / filename
    metadata_df.to_excel(out_path, index=False, engine="openpyxl")
    logger.info("PW upload metadata workbook written: %s", out_path)
    return out_path
