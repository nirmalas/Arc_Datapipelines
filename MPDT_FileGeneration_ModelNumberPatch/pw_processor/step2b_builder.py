"""
pw_processor/step2b_builder.py — Build asset deliverables summary from PW + L2 + Tracker.

This builds a consolidated view of all deliverables with their latest versions,
classification flags (In_Works_Tracker, In_L2_UAID_ACBOS), used for classification.
"""
from __future__ import annotations

import logging
import re
from pathlib import Path

import pandas as pd

from utils.common import normalize_text


UAID2_CANDIDATES = ["UAID_2", "UAID2", "Uaid_2", "ASSET_ID", "ASSET ID", "PW_UAID", "SID"]

REQUIRED_OUTPUT_COLUMNS = [
    "ASSET_ID",
    "DocumentName",
    "FileUpdated",
    "Version",
    "Description",
    "FileName",
    "FullPath",
    "WorkflowState",
    "URN",
    "PW_PROJECT_NAME",
    "PW_TYPE_DESC",
    "PW_UAID",
    "TB_MIDP_STATUS",
]


def _parse_filedate(series: pd.Series) -> pd.Series:
    """Parse FileUpdated — handles Excel serial and date strings."""
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
    return pd.to_datetime(series.apply(_parse_one), errors="coerce")


def _parse_version_number(series: pd.Series) -> pd.Series:
    """Parse version strings like P07.1 -> 7.1 and return float for sorting."""

    def _parse_one(v):
        if pd.isna(v):
            return float("-inf")
        s = str(v).strip()
        if not s or s.lower() in ("nan", "none"):
            return float("-inf")

        # Keep only numeric component (e.g., P07.1 -> 07.1)
        m = re.search(r"(\d+(?:\.\d+)?)", s)
        if not m:
            return float("-inf")
        try:
            return float(m.group(1))
        except Exception:
            return float("-inf")

    return series.apply(_parse_one)


def _pick_column(df: pd.DataFrame, target: str, aliases: list[str]) -> str | None:
    """Find a matching column name using normalized matching and aliases."""
    norm_map = {normalize_text(c): c for c in df.columns}
    for name in [target] + aliases:
        key = normalize_text(name)
        if key in norm_map:
            return norm_map[key]
    return None


def build_asset_deliverables(
    pw_df: pd.DataFrame,
    l2_df: pd.DataFrame,
    logger: logging.Logger,
) -> pd.DataFrame:
    """
    Build asset_deliverables DataFrame from PW extract only.

    Rules:
    - Keep only required columns in required order.
    - Primary key: (DocumentName, ASSET_ID).
    - Ignore rows with null/blank values in any primary key column.
    - Deduplicate by highest Version number (numeric part only, e.g. P07.1 -> 7.1).
    - If still duplicated, keep row with latest FileUpdated.
    """
    if pw_df.empty:
        logger.warning("PW extract is empty, cannot build deliverables.")
        return pd.DataFrame()

    pw = pw_df.copy()

    # Resolve required columns by normalized names where possible.
    alias_map: dict[str, list[str]] = {
        "DocumentName": ["document name", "pw document name"],
        "Description": ["document description", "desc"],
        "Version": ["revision", "rev"],
        "FileName": ["file name", "filename"],
        "FullPath": ["full path", "path"],
        "WorkflowState": ["workflow state", "state"],
        "URN": ["urn"],
        "FileUpdated": ["file updated", "updated", "last updated"],
        "PW_PROJECT_NAME": ["pw project name", "project name"],
        "ASSET_ID": ["asset id", "uaid_2", "uaid2", "pw_uaid", "sid"],
        "PW_TYPE_DESC": ["pw type desc", "type desc", "type"],
        "PW_UAID": ["pw uaid", "uaid", "uaid_2"],
        "TB_MIDP_STATUS": ["tb midp status", "midp status", "status"],
    }

    selected = pd.DataFrame()
    missing_cols: list[str] = []
    for out_col in REQUIRED_OUTPUT_COLUMNS:
        source_col = _pick_column(pw, out_col, alias_map.get(out_col, []))
        if source_col:
            selected[out_col] = pw[source_col]
        else:
            selected[out_col] = ""
            missing_cols.append(out_col)

    if missing_cols:
        logger.warning("PW extract missing expected columns (filled blank): %s", missing_cols)

    # Normalize PK columns and drop rows with null/blank PK values.
    selected["DocumentName"] = selected["DocumentName"].fillna("").astype(str).str.strip()
    selected["ASSET_ID"] = selected["ASSET_ID"].fillna("").astype(str).str.strip()
    before_pk_filter = len(selected)
    selected = selected[(selected["DocumentName"] != "") & (selected["ASSET_ID"] != "")].copy()

    # Filter ASSET_ID values to those beginning with 'HS2-'
    before_hs2 = len(selected)
    selected = selected[selected["ASSET_ID"].str.upper().str.startswith("HS2-")].copy()
    dropped_nonhs2 = before_hs2 - len(selected)

    # Dedupe by (DocumentName, ASSET_ID) using version then FileUpdated.
    selected["_version_num"] = _parse_version_number(selected["Version"])
    selected["_file_date"] = _parse_filedate(selected["FileUpdated"])

    selected = selected.sort_values(
        by=["DocumentName", "ASSET_ID", "_version_num", "_file_date"],
        ascending=[True, True, False, False],
        na_position="last",
    ).drop_duplicates(subset=["DocumentName", "ASSET_ID"], keep="first")

    # Write FileUpdated as readable datetime string where parseable.
    selected["FileUpdated"] = selected["_file_date"].dt.strftime("%Y-%m-%d %H:%M:%S").fillna(
        selected["FileUpdated"].fillna("").astype(str)
    )

    # Keep only required output columns.
    selected = selected[REQUIRED_OUTPUT_COLUMNS].copy()

    logger.info(
        "Asset deliverables built: %d rows (dropped %d null-key rows, %d non-HS2 rows)",
        len(selected),
        before_pk_filter - before_hs2,
        dropped_nonhs2,
    )
    return selected
