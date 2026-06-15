"""
data_fetcher/smartforms_fetcher.py — Fetch SmartForms data via API or local fallback.

Can be run standalone:
  python -m data_fetcher.smartforms_fetcher
"""
from __future__ import annotations

import argparse
import logging
from pathlib import Path

import pandas as pd

from utils.common import (
    load_config,
    resolve_workspace,
    setup_logger,
    write_table_cache,
    read_table_any,
)


def fetch_smartforms_api(
    cfg: dict, logger: logging.Logger
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Fetch SmartForms L2 and L3 data from MIDP API using MSAL."""
    import msal
    import requests

    api_cfg = cfg.get("smartforms_api", {})
    tenant_id = api_cfg["tenant_id"]
    client_id = api_cfg.get("client_id", "")
    scope = api_cfg.get("scope", "https://api.bentley.com/.default")
    base_url = api_cfg.get("base_url", "https://api.bentley.com/assetregistry")

    if not client_id:
        raise ValueError("SmartForms API client_id is not configured.")

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.PublicClientApplication(client_id=client_id, authority=authority)
    token_result = app.acquire_token_interactive(scopes=[scope])
    if "access_token" not in token_result:
        raise RuntimeError(
            f"Token acquisition failed: {token_result.get('error_description')}"
        )

    headers = {"Authorization": f"Bearer {token_result['access_token']}"}

    def _get_all(endpoint: str) -> list[dict]:
        records, skip = [], 0
        while True:
            resp = requests.get(
                f"{base_url}/{endpoint}",
                headers=headers,
                params={"$top": 1000, "$skip": skip},
                timeout=60,
            )
            resp.raise_for_status()
            batch = resp.json().get("value", [])
            if not batch:
                break
            records.extend(batch)
            skip += len(batch)
        return records

    logger.info("Fetching SmartForms L2 assets via API...")
    l2_records = _get_all("level2assets")
    logger.info("Fetching SmartForms L3 assets via API...")
    l3_records = _get_all("level3assets")

    l2_df = pd.DataFrame(l2_records) if l2_records else pd.DataFrame()
    l3_df = pd.DataFrame(l3_records) if l3_records else pd.DataFrame()

    logger.info("SmartForms API — L2: %d, L3: %d", len(l2_df), len(l3_df))
    return l2_df, l3_df


def load_smartforms_fallback(
    workspace: Path, cfg: dict, logger: logging.Logger
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Load SmartForms from local Excel fallback file."""
    fallback_rel = cfg.get("paths", {}).get(
        "smartforms_fallback", "Input/SmartForms_RAW_MPDT_L2&L3.xlsx"
    )
    fallback = (workspace / fallback_rel).resolve()
    if not fallback.exists():
        logger.warning("SmartForms fallback Excel not found: %s", fallback)
        return pd.DataFrame(), pd.DataFrame()

    logger.info("Loading SmartForms from fallback Excel: %s", fallback.name)
    # Try fast path: if a parquet cache exists, use it via read_table_any elsewhere.
    # Otherwise attempt to read only necessary columns (usecols) to speed up large files.
    try:
        xl = pd.ExcelFile(fallback, engine="openpyxl")
        l2_sheet = next((s for s in xl.sheet_names if "l2" in s.lower()), xl.sheet_names[0])
        l3_sheet = next((s for s in xl.sheet_names if "l3" in s.lower()), xl.sheet_names[-1])

        # Define commonly-used columns (picked from downstream code usage).
        l2_usecols = [
            # UAID_2 variants
            'UAID_2', 'Uaid_2', 'Uaid', 'ASSET_ID',
            # names and classification
            'AssetName', 'Asset_Name', 'Asset Name', 'ClassCode', 'ClassName',
        ]
        l3_usecols = [
            'UAID_3', 'Uaid_3', 'Uaid', 'UAID', 'Asset_ID',
            'ParentUaid', 'Parent_UAID',
            'ClassCode', 'ClassName', 'AttributeTypeId', 'AttTypeName',
        ]

        # Augment usecols from MPDT mapping cache if available
        try:
            cache_dir = workspace / cfg.get('paths', {}).get('db_cache_dir', 'Input/DB_Cache')
            mapping_cache = cache_dir / 'mpdt_mapping_cache.json'
            if mapping_cache.exists():
                import json, re

                payload = json.loads(mapping_cache.read_text(encoding='utf-8'))
                mapping = payload.get('mapping', {}) if isinstance(payload, dict) else {}
                # Extract join1/join2 column names from mapping expressions like "join2[UAID_3]"
                col_names = set()
                for expr in mapping.values():
                    if not isinstance(expr, str):
                        continue
                    for m in re.finditer(r"join[12]\s*\[\s*['\"]?([^'\"]+)['\"]?\s*\]", expr, re.IGNORECASE):
                        col_names.add(m.group(1).strip())
                # Also include mapping target headers that may indicate L3/L2 fields
                for target in mapping.keys():
                    if isinstance(target, str) and ('_3' in target or '_2' in target or 'AssetName' in target):
                        col_names.add(target)

                # Heuristically add discovered names to L2/L3 candidate lists
                for cn in col_names:
                    low = cn.lower()
                    if '3' in low or 'uaid_3' in low or 'parent' in low or 'asset_id' in low:
                        l3_usecols.append(cn)
                    else:
                        l2_usecols.append(cn)
        except Exception:
            # Non-fatal — proceed with default candidates
            pass

        # Helper to pick available usecols present in sheet
        def _available_usecols(sheet_name, candidate_cols):
            cols = [c for c in pd.read_excel(fallback, sheet_name=sheet_name, nrows=0, engine='openpyxl').columns]
            # Return the actual sheet column names that match any candidate (case-insensitive)
            matched = []
            for d in cols:
                for c in candidate_cols:
                    if c.lower() == str(d).lower():
                        matched.append(d)
                        break
            return matched

        l2_use = _available_usecols(l2_sheet, l2_usecols)
        l3_use = _available_usecols(l3_sheet, l3_usecols)

        if l2_use:
            l2 = pd.read_excel(fallback, sheet_name=l2_sheet, dtype=str, usecols=l2_use, engine="openpyxl")
        else:
            l2 = pd.read_excel(fallback, sheet_name=l2_sheet, dtype=str, engine="openpyxl")

        if l3_use:
            l3 = pd.read_excel(fallback, sheet_name=l3_sheet, dtype=str, usecols=l3_use, engine="openpyxl")
        else:
            l3 = pd.read_excel(fallback, sheet_name=l3_sheet, dtype=str, engine="openpyxl")

        # If the fallback file is in normalized attribute format (rows of AttributeTypeId/AttributeValue),
        # pivot to a wide table so downstream mapping expressions like join2['startChainage'] work.
        def _pivot_attributes(df):
            if df is None or df.empty:
                return df
            cols = [str(c).strip() for c in df.columns]
            lower = [c.lower() for c in cols]
            # identify key columns
            attr_id_col = None
            for cand in ('AttributeTypeId', 'AttrTypeCode', 'AttrTypeDisplayName'):
                if any(c.lower() == cand.lower() for c in cols):
                    attr_id_col = next(c for c in cols if c.lower() == cand.lower())
                    break
            val_col = next((c for c in cols if c.lower() == 'attributevalue'), None)
            uaid_col = None
            for cand in ('Uaid', 'Uaid_3', 'Uaid_2', 'UAID', 'Asset_ID'):
                for c in cols:
                    if c.lower() == cand.lower():
                        uaid_col = c
                        break
                if uaid_col:
                    break

            if not (attr_id_col and val_col and uaid_col):
                return df

            try:
                wide = df[[uaid_col, attr_id_col, val_col]].copy()
                wide.columns = [uaid_col, 'attr_id', 'attr_val']
                # pivot
                pivoted = wide.pivot_table(index=uaid_col, columns='attr_id', values='attr_val', aggfunc='first')
                pivoted = pivoted.reset_index()
                # Ensure string dtype
                pivoted = pivoted.astype(str).replace('nan', '')
                return pivoted
            except Exception:
                return df

        l2_wide = _pivot_attributes(l2)
        l3_wide = _pivot_attributes(l3)

        # If pivot succeeded (i.e., produced more columns), use wide versions
        if not l2_wide.empty and len(l2_wide.columns) > len(l2.columns):
            l2 = l2_wide
        if not l3_wide.empty and len(l3_wide.columns) > len(l3.columns):
            l3 = l3_wide

        # Cache pivoted results for faster subsequent loads
        try:
            cache_dir = workspace / cfg.get('paths', {}).get('db_cache_dir', 'Input/DB_Cache')
            write_table_cache(l2.astype(str), cache_dir / 'SmartForms_L2', logger)
            write_table_cache(l3.astype(str), cache_dir / 'SmartForms_L3', logger)
        except Exception:
            pass

        logger.info("  SmartForms L2: %d rows | L3: %d rows (used cols L2=%d, L3=%d)", len(l2), len(l3), len(l2.columns), len(l3.columns))
        return l2, l3
    except Exception as exc:
        logger.warning("Fallback SmartForms read failed (attempting full read): %s", exc)
        # Try full read as last resort
        l2 = pd.read_excel(fallback, dtype=str, engine="openpyxl")
        l3 = pd.read_excel(fallback, dtype=str, engine="openpyxl")
        logger.info("  SmartForms L2: %d rows | L3: %d rows", len(l2), len(l3))
        return l2, l3


def fetch_and_cache_smartforms(
    workspace: Path, cfg: dict, logger: logging.Logger
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Try API first (if enabled + configured), fallback to local Excel.
    Caches results to DB_Cache/.
    """
    cache_dir = workspace / cfg.get("paths", {}).get("db_cache_dir", "Input/DB_Cache")
    use_cache = cfg.get("use_db_cache", True)

    # Check cache first
    if use_cache:
        l2_cached = read_table_any(cache_dir / "SmartForms_L2", logger)
        l3_cached = read_table_any(cache_dir / "SmartForms_L3", logger)
        if not l2_cached.empty or not l3_cached.empty:
            logger.info("SmartForms loaded from cache — L2: %d, L3: %d", len(l2_cached), len(l3_cached))
            return l2_cached, l3_cached

    # Try API
    api_cfg = cfg.get("smartforms_api", {})
    if api_cfg.get("enabled", False) and api_cfg.get("client_id"):
        try:
            l2_df, l3_df = fetch_smartforms_api(cfg, logger)
            # Cache results
            if not l2_df.empty:
                write_table_cache(l2_df, cache_dir / "SmartForms_L2", logger)
            if not l3_df.empty:
                write_table_cache(l3_df, cache_dir / "SmartForms_L3", logger)
            return l2_df, l3_df
        except Exception as exc:
            logger.warning("SmartForms API failed (%s). Falling back to Excel.", exc)

    # Fallback to Excel
    l2_df, l3_df = load_smartforms_fallback(workspace, cfg, logger)
    # Cache
    if not l2_df.empty:
        write_table_cache(l2_df, cache_dir / "SmartForms_L2", logger)
    if not l3_df.empty:
        write_table_cache(l3_df, cache_dir / "SmartForms_L3", logger)
    return l2_df, l3_df


def fetch_smartforms_and_save(
    workspace: Path,
    cfg: dict,
    logger: logging.Logger,
    output_path: Path,
) -> tuple[int, int]:
    """Fetch SmartForms from API and overwrite the canonical local Excel file.

    Used by Step 1 when fetch_external.smartforms=True.
    Returns (l2_row_count, l3_row_count).
    Raises if the API is not enabled/configured (caller should catch).
    """
    api_cfg = cfg.get("smartforms_api", {})
    if not api_cfg.get("enabled", False) or not api_cfg.get("client_id", ""):
        raise ValueError(
            "SmartForms API is not enabled or client_id is missing. "
            "Set smartforms_api.enabled=true and configure client_id in pipeline_config.json, "
            "OR set fetch_external.smartforms=false to use the local file."
        )

    l2_df, l3_df = fetch_smartforms_api(cfg, logger)

    if l2_df.empty and l3_df.empty:
        raise RuntimeError("SmartForms API returned no data.")

    # Write both sheets to the canonical local file.
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(str(output_path), engine="openpyxl") as xl:
        l2_df.to_excel(xl, sheet_name="SmartForms_L2", index=False)
        l3_df.to_excel(xl, sheet_name="SmartForms_L3", index=False)
    logger.info("SmartForms written to %s — L2: %d, L3: %d", output_path.name, len(l2_df), len(l3_df))
    return len(l2_df), len(l3_df)


# ---------------------------------------------------------------------------
# Standalone entry point
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description="Fetch SmartForms data")
    parser.add_argument("--workspace", default=None)
    parser.add_argument("--no-cache", action="store_true")
    args = parser.parse_args()

    workspace = resolve_workspace(args.workspace)
    cfg = load_config(workspace)
    if args.no_cache:
        cfg["use_db_cache"] = False
    logger = setup_logger(workspace, "smartforms_fetcher", cfg.get("log_level", "INFO"))

    logger.info("=== SmartForms Fetcher ===")
    l2, l3 = fetch_and_cache_smartforms(workspace, cfg, logger)
    logger.info("=== Complete — L2: %d, L3: %d ===", len(l2), len(l3))


if __name__ == "__main__":
    main()
