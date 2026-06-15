#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
mcid_mapper.py — Build Asset (UAID_2) -> DM3 Model Container mapping using MIDP and optional PW extract.

Enhancements:
- Normalize/expand the MIDP "Assets" column when it contains multiple assets per cell.
- Split each asset into two fields: UAID_2 and AssetDescription.
- Prefer exact UAID_2 matching on the exploded column during resolution, falling back to substring search when needed.

Business rules (from your refs):
1) Assets column contains UAID_2 [3.1]
2) Keep only DM3 deliverables
3) If several remain, prefer rows whose description contains "solid"
4) If several still remain, cross-check against ACBOS/PW document tokens
5) Return one unique deliverable number, otherwise blank
"""

from __future__ import annotations

import argparse
import logging
import re
from pathlib import Path
from typing import Any, Iterable, Optional, Set, Tuple

import pandas as pd


# ----------------------------
# Helpers (mirroring your refs)
# ----------------------------

def _is_empty(value: Any) -> bool:
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except Exception:
        pass
    return str(value).strip() in ("", "nan", "None", "NaN", "NaT")


def _normalize_text(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).strip()).lower()


def _pick_column(df: pd.DataFrame, candidates: tuple[str, ...]) -> Optional[str]:
    """Pick the first column whose normalized name matches one of the candidate names."""
    if df is None or df.empty:
        return None
    wanted = {_normalize_text(c) for c in candidates}
    for col in df.columns:
        if _normalize_text(col) in wanted:
            return col
    return None


def _document_tokens_from_pw_extract(pw_df: pd.DataFrame | None) -> Set[str]:
    """
    Return normalized document identifiers from ACBOS MPDT / PW extract:
    collects both full names and stems (upper-cased).
    """
    if pw_df is None or pw_df.empty:
        return set()
    cols = []
    for name in ("DocumentName", "Document Name", "FileName", "File Name", "Deliverable No", "Deliverable Number"):
        col = _pick_column(pw_df, (name,))
        if col and col not in cols:
            cols.append(col)
    tokens: set[str] = set()
    for col in cols:
        for raw in pw_df[col].dropna().astype(str):
            val = raw.strip()
            if not val:
                continue
            tokens.add(val.upper())
            tokens.add(Path(val).stem.upper())
    return tokens


# ----------------------------
# Asset cell normalization
# ----------------------------

_ASSET_ID_RE = re.compile(r"(HS2-[A-Za-z0-9]+)")

def _split_asset_item(item: str) -> Tuple[str, str]:
    """
    Split a single asset item into (UAID_2, AssetDescription).
    Expected formats:
      - "HS2-000001016 - South Heath Cutting"
      - "HS2-00002CT27 - Grims Ditch North Retaining Wall"
    Fallback: if " - " is missing, try regex HS2-... and treat the rest as description.
    """
    s = str(item).strip()
    # Remove stray header string that sometimes appears in data cells
    if s.strip().upper() == "ASSETS":
        return "", ""
    # Preferred split by " - "
    if " - " in s:
        left, right = s.split(" - ", 1)
        return left.strip(), right.strip()
    # Fallback: regex
    m = _ASSET_ID_RE.search(s)
    if m:
        uaid = m.group(1).strip()
        # Take everything after the matched UAID (strip leading separators)
        after = s[m.end():].lstrip(" -:–—\t")
        return uaid, after.strip()
    # No identifiable UAID
    return "", ""


def _explode_assets_column(midp_df: pd.DataFrame) -> pd.DataFrame:
    """
    Expand MIDP rows where the 'Assets' cell holds multiple assets separated by commas.
    For each asset token, create a row copy and add columns:
      - __AssetID (UAID_2)
      - __AssetDesc (the text after ' - ')
    Keeps original columns (Deliverable No, Description, etc.).
    Rows without a parsable UAID are dropped.
    """
    if midp_df is None or midp_df.empty:
        return pd.DataFrame()

    assets_col = _pick_column(midp_df, ("Assets", "Asset", "Asset ID", "Asset_ID", "UAID", "UAID_2"))
    if not assets_col:
        # No assets column — return as-is with placeholders
        out = midp_df.copy()
        out["__AssetID"] = ""
        out["__AssetDesc"] = ""
        return out

    rows = []
    for _, r in midp_df.iterrows():
        raw = str(r.get(assets_col, "") or "")
        if not raw.strip():
            # still append one row with blanks for completeness (optional)
            nr = r.copy()
            nr["__AssetID"] = ""
            nr["__AssetDesc"] = ""
            rows.append(nr)
            continue
        # Split by commas/newlines; assets are comma-separated in your example
        parts = re.split(r"[,\n]+", raw)
        for part in parts:
            part = part.strip()
            if not part:
                continue
            aid, desc = _split_asset_item(part)
            if not aid:
                continue
            nr = r.copy()
            nr["__AssetID"] = aid
            nr["__AssetDesc"] = desc
            rows.append(nr)

    if not rows:
        return pd.DataFrame(columns=list(midp_df.columns) + ["__AssetID", "__AssetDesc"])

    out = pd.DataFrame(rows)
    # Normalize strings
    for c in out.select_dtypes(include=["object"]).columns:
        out[c] = out[c].astype(str).str.strip()
    return out


# ----------------------------
# Resolver
# ----------------------------

class ModelContainerResolver:
    """
    Fast UAID_2 -> DM3 model-container lookup on exploded assets.
    Behavior mirrors your referenced business rules [3.1, 3.2].
    """

    def __init__(self, midp_df: pd.DataFrame | None, pw_df: pd.DataFrame | None = None, logger: logging.Logger | None = None):
        self.logger = logger or logging.getLogger(self.__class__.__name__)
        # Normalize MIDP: explode assets and split into id/desc
        base_df = midp_df.copy() if midp_df is not None and not midp_df.empty else pd.DataFrame()
        self.midp_df = _explode_assets_column(base_df) if not base_df.empty else base_df

        # Columns
        self.asset_id_col = "__AssetID" if "__AssetID" in self.midp_df.columns else None
        self.asset_desc_col = "__AssetDesc" if "__AssetDesc" in self.midp_df.columns else None
        self.deliverable_col = _pick_column(self.midp_df, ("Deliverable No", "Deliverable Number", "DocumentName", "Document Name")) if not self.midp_df.empty else None
        self.description_col = _pick_column(self.midp_df, ("Description", "Deliverable Name", "Deliverable Description", "Name")) if not self.midp_df.empty else None

        self.pw_document_tokens = _document_tokens_from_pw_extract(pw_df)

        if not self.midp_df.empty and (not self.asset_id_col or not self.deliverable_col):
            self.logger.warning(
                "MIDP loaded but required columns were not found (asset_id_col=%s, deliverable_col=%s). "
                "Model Container ID resolution may return blanks.",
                self.asset_id_col,
                self.deliverable_col,
            )

    def _unique_deliverable(self, df: pd.DataFrame) -> str:
        if df is None or df.empty or not self.deliverable_col:
            return ""
        vals: list[str] = []
        seen: set[str] = set()
        for raw in df[self.deliverable_col].dropna().astype(str):
            val = raw.strip()
            if not val or val.lower() == "nan":
                continue
            key = val.upper()
            if key not in seen:
                vals.append(val)
                seen.add(key)
        return vals[0] if len(vals) == 1 else ""

    def resolve(self, uaid2: str) -> str:
        """
        Return the unique DM3 deliverable number for UAID_2, or blank.
        Uses exploded __AssetID when available, else falls back to substring search.
        """
        if self.midp_df.empty or not self.deliverable_col or _is_empty(uaid2):
            return ""
        uaid = str(uaid2).strip()
        if not uaid:
            return ""

        candidates = self.midp_df
        # Prefer exact match on exploded __AssetID
        if self.asset_id_col:
            candidates = candidates[candidates[self.asset_id_col].fillna("").astype(str).str.strip().str.upper() == uaid.upper()]
        else:
            # Fallback: substring search in original assets column (if present)
            assets_col = _pick_column(self.midp_df, ("Assets", "Asset", "Asset ID", "Asset_ID", "UAID", "UAID_2"))
            if not assets_col:
                return ""
            asset_text = self.midp_df[assets_col].fillna("").astype(str)
            candidates = self.midp_df[asset_text.str.contains(re.escape(uaid), case=False, na=False)]

        if candidates.empty:
            return ""

        # Keep only DM3 deliverables
        deliverable_text = candidates[self.deliverable_col].fillna("").astype(str)
        candidates = candidates[deliverable_text.str.contains("DM3", case=False, na=False)]
        if candidates.empty:
            return ""

        one = self._unique_deliverable(candidates)
        if one:
            return one

        # Prefer 'solid' in the description if ambiguous
        if self.description_col:
            desc_text = candidates[self.description_col].fillna("").astype(str)
            solid = candidates[desc_text.str.contains("solid", case=False, na=False)]
            one = self._unique_deliverable(solid)
            if one:
                return one
            if not solid.empty:
                candidates = solid

        # Cross-check against PW tokens (stem or full)
        if self.pw_document_tokens:
            matched = candidates[
                candidates[self.deliverable_col]
                .fillna("")
                .astype(str)
                .map(lambda v: Path(v.strip()).stem.upper() in self.pw_document_tokens or v.strip().upper() in self.pw_document_tokens)
            ]
            one = self._unique_deliverable(matched)
            if one:
                return one
        
            
        if not matched.empty:
            return ", ".join(matched[self.deliverable_col].dropna().astype(str).unique())
            

        return ""


# ----------------------------
# IO helpers
# ----------------------------

def read_any(path: Path) -> pd.DataFrame:
    ext = path.suffix.lower()
    if ext == ".csv":
        return pd.read_csv(path)
    if ext in {".xlsx", ".xls"}:
        return pd.read_excel(path)
    raise ValueError(f"Unsupported input type: {ext} (use CSV/XLSX/XLS)")


# ----------------------------
# CLI
# ----------------------------

def parse_args():
    ap = argparse.ArgumentParser(description="Resolve DM3 Model Container per UAID_2 using MIDP (and optional PW) data, with asset expansion.")
    ap.add_argument("--midp", required=True, help="Path to MIDP Navigator export (CSV or Excel)")
    ap.add_argument("--pw", help="Optional ProjectWise extract (CSV or Excel) for cross-check")
    ap.add_argument("--uaids", help="Optional text file with UAID_2 values (one per line). If omitted, derive UAIDs from exploded MIDP.")
    ap.add_argument("--out", required=True, help="Output CSV path for Asset -> ModelContainer mapping")
    ap.add_argument("-v", "--verbose", action="count", default=1, help="Increase verbosity (-v, -vv)")
    return ap.parse_args()


def main():
    args = parse_args()
    level = logging.WARNING if args.verbose == 0 else (logging.INFO if args.verbose == 1 else logging.DEBUG)
    logging.basicConfig(level=level, format="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
    log = logging.getLogger("mcid_mapper")

    midp_path = Path(args.midp)
    pw_path = Path(args.pw) if args.pw else None
    out_path = Path(args.out)

    midp_df = read_any(midp_path)
    pw_df = read_any(pw_path) if pw_path and pw_path.exists() else None

    resolver = ModelContainerResolver(midp_df=midp_df, pw_df=pw_df, logger=log)

    # Establish UAID_2 list
    if args.uaids:
        with open(args.uaids, "r", encoding="utf-8") as f:
            uaids = [ln.strip() for ln in f if ln.strip()]
    else:
        # Use exploded __AssetID column if present
        if resolver.asset_id_col and not resolver.midp_df.empty:
            uaids = sorted({str(v).strip() for v in resolver.midp_df[resolver.asset_id_col].dropna() if str(v).strip()})
        else:
            # Fallback: derive heuristically from any Assets-like column
            assets_col = _pick_column(resolver.midp_df, ("Assets", "Asset", "Asset ID", "Asset_ID", "UAID", "UAID_2"))
            if not assets_col:
                raise RuntimeError("Could not derive UAIDs (no __AssetID and no Assets-like column). Provide --uaids.")
            tokens: set[str] = set()
            for a in resolver.midp_df[assets_col].dropna().astype(str):
                parts = re.split(r"[,\n]+", a)
                for p in parts:
                    aid, _ = _split_asset_item(p)
                    if aid:
                        tokens.add(aid)
            uaids = sorted(tokens)
        log.info(f"Resolved {len(uaids)} UAID candidates")

    rows = []
    for u in uaids:
        mc = resolver.resolve(u)
        rows.append({"Asset": u, "ModelContainer": mc})

    out_df = pd.DataFrame(rows).drop_duplicates().sort_values(["Asset"]).reset_index(drop=True)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_df.to_csv(out_path, index=False)
    log.info(f"Wrote {len(out_df)} mappings to {out_path}")

if __name__ == "__main__":
    main()