"""Resolve MPDT Model Container IDs from MIDP Navigator data.

The resolver implements the business rule used by MPDT generation:
  1. Start from MIDP_Navigator.xlsx rows whose Assets column contains UAID_2.
  2. Keep only rows where that Assets cell contains one asset.
  3. Keep only deliverables whose final 6-digit sequence starts with 0.
  4. Keep only DM3 model deliverables.
  5. If several remain, prefer rows whose description/name contains "solid".
  6. If several still remain, cross-check against ACBOS MPDT / ProjectWise extract.
  7. Return one unique deliverable number, otherwise blank.
"""
from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Any

import pandas as pd

from utils.common import normalize_text


def _is_empty(value: Any) -> bool:
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except Exception:
        pass
    return str(value).strip() in ("", "nan", "NaN", "NaT")


def _pick_column(df: pd.DataFrame, candidates: tuple[str, ...]) -> str | None:
    wanted = {normalize_text(c) for c in candidates}
    for col in df.columns:
        if normalize_text(col) in wanted:
            return col
    return None


def _document_tokens_from_pw_extract(pw_df: pd.DataFrame | None) -> set[str]:
    """Return normalized document identifiers from ACBOS MPDT / PW extract."""
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


class ModelContainerResolver:
    """Fast UAID_2 -> DM3 model-container lookup built once per MPDT batch.

    The resolver is intentionally conservative: it returns a Model Container ID
    only when the filtering rules leave exactly one unique deliverable number.
    """

    def __init__(
        self,
        midp_df: pd.DataFrame | None,
        pw_df: pd.DataFrame | None = None,
        dm3_df: pd.DataFrame | None = None,
        logger: logging.Logger | None = None,
    ):
        self.logger = logger
        self.midp_df = midp_df.copy() if midp_df is not None and not midp_df.empty else pd.DataFrame()
        self.asset_col = _pick_column(self.midp_df, ("Assets", "Asset", "Asset ID", "Asset_ID", "UAID", "UAID_2")) if not self.midp_df.empty else None
        self.deliverable_col = _pick_column(self.midp_df, ("Deliverable No", "Deliverable Number", "DocumentName", "Document Name")) if not self.midp_df.empty else None
        self.deliverable_type_col = _pick_column(self.midp_df, ("Deliverable Type", "DeliverableType", "Document Type", "Doc Type", "Type")) if not self.midp_df.empty else None
        self.description_col = _pick_column(self.midp_df, ("Description", "Deliverable Name", "Deliverable Description", "Name")) if not self.midp_df.empty else None
        self.discipline_col = _pick_column(self.midp_df, ("Discipline", "Disc", "Discipline Code", "Discipline_Code")) if not self.midp_df.empty else None
        self.pw_document_tokens = _document_tokens_from_pw_extract(pw_df)
        self.dm3_document_tokens, self.dm3_solid_tokens = self._document_tokens_from_dm3_extract(dm3_df)

        if not self.midp_df.empty and (not self.asset_col or not self.deliverable_col):
            if self.logger:
                self.logger.warning(
                    "MIDP Navigator loaded but required columns were not found (asset_col=%s, deliverable_col=%s). "
                    "Model Container ID will be left blank.",
                    self.asset_col,
                    self.deliverable_col,
                )

    @staticmethod
    def discipline_from_deliverable(deliverable_name: str | None) -> str:
        """Return discipline from third hyphen-separated part of deliverable/file name.

        Example: 1MC06-CEK-BR-FRM-CS06_CL09-000006 -> BR.
        """
        if _is_empty(deliverable_name):
            return ""
        stem = Path(str(deliverable_name).strip()).stem
        parts = [p.strip().upper() for p in stem.split("-") if p.strip()]
        return parts[2] if len(parts) >= 3 else ""

    @staticmethod
    def _normalize_discipline_code(value: Any) -> str:
        """Normalize MIDP/MPDT discipline values to a 2-letter code.

        Examples:
          AR - Architecture -> AR
          ARchitecture -> AR
          BR -> BR
          1MC06-CEK-BR-DM3-CS06_CL09-000007 -> BR
        """
        if _is_empty(value):
            return ""
        text = str(value).strip().upper()
        # A full deliverable number carries the discipline in the third part.
        parts = [p.strip().upper() for p in Path(text).stem.split("-") if p.strip()]
        if len(parts) >= 3 and re.fullmatch(r"[A-Z]{2,3}", parts[2]):
            return parts[2][:2]
        letters = re.sub(r"[^A-Z]", "", text)
        return letters[:2] if len(letters) >= 2 else ""

    @staticmethod
    def _discipline_from_midp_deliverable(value: Any) -> str:
        if _is_empty(value):
            return ""
        parts = [p.strip().upper() for p in Path(str(value).strip()).stem.split("-") if p.strip()]
        return parts[2][:2] if len(parts) >= 3 else ""

    @staticmethod
    def _document_tokens_from_dm3_extract(dm3_df: pd.DataFrame | None) -> tuple[set[str], set[str]]:
        """Return all current DM3 document tokens and tokens that look like solids."""
        if dm3_df is None or dm3_df.empty:
            return set(), set()

        name_col = _pick_column(dm3_df, ("Name", "Doc. Reference", "DocumentName", "Document Name", "Deliverable No", "Deliverable Number"))
        file_col = _pick_column(dm3_df, ("File Name", "FileName"))
        desc_col = _pick_column(dm3_df, ("Description", "Title", "Folder Description"))
        cols = [c for c in (name_col, file_col) if c]
        if not cols:
            return set(), set()

        all_tokens: set[str] = set()
        solid_tokens: set[str] = set()
        for _, row in dm3_df.iterrows():
            row_tokens: set[str] = set()
            for col in cols:
                raw = row.get(col, "")
                if _is_empty(raw):
                    continue
                text = str(raw).strip().upper()
                stem = Path(text).stem.upper()
                if stem:
                    row_tokens.add(stem)
                if text:
                    row_tokens.add(text)
            if not row_tokens:
                continue

            all_tokens.update(row_tokens)
            desc = str(row.get(desc_col, "")).strip().lower() if desc_col else ""
            if "solid" in desc and not re.search(r"withdrawn|not in use|civil 3d native", desc):
                solid_tokens.update(row_tokens)
        return all_tokens, solid_tokens

    @staticmethod
    def _asset_items(value: Any) -> list[str]:
        """Return distinct asset identifiers/items listed in a MIDP Assets cell."""
        if _is_empty(value):
            return []
        text = str(value).strip()
        hs2_ids = re.findall(r"HS2-[A-Za-z0-9]+", text, flags=re.IGNORECASE)
        if hs2_ids:
            seen: set[str] = set()
            out: list[str] = []
            for item in hs2_ids:
                key = item.upper()
                if key not in seen:
                    seen.add(key)
                    out.append(item)
            return out
        return [p.strip() for p in re.split(r"[,;|\n\r]+", text) if p.strip()]

    @classmethod
    def _has_one_asset(cls, value: Any) -> bool:
        return len(cls._asset_items(value)) == 1

    @staticmethod
    def _has_zero_prefixed_final_sequence(value: Any) -> bool:
        """True when the deliverable stem ends with a 6-digit segment starting with 0."""
        if _is_empty(value):
            return False
        stem = Path(str(value).strip()).stem.upper()
        return bool(re.search(r"-0\d{5}$", stem))

    def _filter_dm3_models(self, df: pd.DataFrame) -> pd.DataFrame:
        if df is None or df.empty or not self.deliverable_col:
            return df

        deliverable_text = df[self.deliverable_col].fillna("").astype(str)
        mask = deliverable_text.str.contains("DM3", case=False, na=False)

        if self.deliverable_type_col:
            type_text = df[self.deliverable_type_col].fillna("").astype(str)
            if type_text.str.strip().astype(bool).any():
                type_mask = (
                    type_text.str.contains("DM3", case=False, na=False)
                    | type_text.str.contains(r"\b3D\b.*\bmodel\b|\bmodel\b", case=False, na=False, regex=True)
                )
                mask = mask & type_mask

        return df[mask]

    def _base_candidates(self, uaid2: str) -> pd.DataFrame:
        if self.midp_df.empty or not self.asset_col or not self.deliverable_col or _is_empty(uaid2):
            return pd.DataFrame()
        uaid = str(uaid2).strip()
        if not uaid:
            return pd.DataFrame()

        asset_text = self.midp_df[self.asset_col].fillna("").astype(str)
        candidates = self.midp_df[asset_text.str.contains(re.escape(uaid), case=False, na=False)]
        if candidates.empty:
            return candidates

        single_asset = candidates[self.asset_col].map(self._has_one_asset)
        candidates = candidates[single_asset]
        if candidates.empty:
            return candidates

        deliverable_text = candidates[self.deliverable_col].fillna("").astype(str)
        candidates = candidates[deliverable_text.map(self._has_zero_prefixed_final_sequence)]
        if candidates.empty:
            return candidates

        return self._filter_dm3_models(candidates)

    def _filter_by_discipline(self, df: pd.DataFrame, discipline: str | None) -> pd.DataFrame:
        disc = self._normalize_discipline_code(discipline)
        if not disc or df is None or df.empty or not self.deliverable_col:
            return df

        def row_disc(row: pd.Series) -> str:
            if self.discipline_col and not _is_empty(row.get(self.discipline_col)):
                return self._normalize_discipline_code(row.get(self.discipline_col))
            return self._discipline_from_midp_deliverable(row.get(self.deliverable_col))

        mask = df.apply(lambda row: row_disc(row) == disc, axis=1)
        return df[mask]

    def _filter_by_token_set(self, df: pd.DataFrame, tokens: set[str]) -> pd.DataFrame:
        if df is None or df.empty or not self.deliverable_col or not tokens:
            return pd.DataFrame()
        return df[
            df[self.deliverable_col]
            .fillna("")
            .astype(str)
            .map(lambda v: Path(v.strip()).stem.upper() in tokens or v.strip().upper() in tokens)
        ]

    def resolve(self, uaid2: str, discipline: str | None = None) -> str:
        """Return the unique DM3 deliverable number for UAID_2, or blank."""
        candidates = self._base_candidates(uaid2)
        if candidates.empty:
            return ""

        # New rule: when the MPDT filename/discipline is available, keep only
        # DM3 models with the same discipline.  This removes cases such as AR
        # architectural DM3 models when the MPDT is BR/bridges.
        discipline_candidates = self._filter_by_discipline(candidates, discipline)
        if not discipline_candidates.empty:
            candidates = discipline_candidates

        if self.dm3_document_tokens:
            dm3_matched = self._filter_by_token_set(candidates, self.dm3_document_tokens)
            if not dm3_matched.empty:
                dm3_solid = self._filter_by_token_set(dm3_matched, self.dm3_solid_tokens)
                one = self._unique_deliverable(dm3_solid)
                if one:
                    return one
                one = self._unique_deliverable(dm3_matched)
                if one:
                    return one
                candidates = dm3_matched

        one = self._unique_deliverable(candidates)
        if one:
            return one

        # Prefer rows whose description/name contains "solid".
        if self.description_col:
            desc_text = candidates[self.description_col].fillna("").astype(str)
            solid = candidates[desc_text.str.contains("solid", case=False, na=False)]
            one = self._unique_deliverable(solid)
            if one:
                return one
            if not solid.empty:
                candidates = solid

        # Cross-check against ACBOS MPDT / ProjectWise extract.
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

        return ""

    def candidate_records(self, uaid2: str, discipline: str | None = None) -> list[dict[str, Any]]:
        """Return diagnostic candidate records for MCID matching output."""
        candidates = self._base_candidates(uaid2)
        if candidates.empty or not self.deliverable_col:
            return [{
                "UAID_2": uaid2,
                "MPDT_Discipline": discipline or "",
                "MIDP_Deliverable": "",
                "MIDP_Discipline": "",
                "Description": "",
                "MIDP_Deliverable_Type": "",
                "MIDP_Asset_Count": 0,
                "Has_Zero_Prefixed_Final_Sequence": False,
                "Present_In_All_Current_DM3": False,
                "Solid_In_All_Current_DM3": False,
                "Matched_Discipline": False,
                "Contains_Solid": False,
                "Resolved_Model_Container_ID": "",
                "Resolution_Status": "No MIDP match after single-asset, final-number and DM3 filters",
            }]

        disc = self._normalize_discipline_code(discipline)
        # Diagnostic output should show only candidates whose MIDP_Discipline
        # agrees with MPDT_Discipline.  MIDP values such as "AR - Architecture"
        # or "ARchitecture" are normalized to "AR" before comparison.
        discipline_candidates = self._filter_by_discipline(candidates, disc)
        if disc:
            candidates = discipline_candidates
        if candidates.empty:
            return [{
                "UAID_2": uaid2,
                "MPDT_Discipline": disc,
                "MIDP_Deliverable": "",
                "MIDP_Discipline": "",
                "Description": "",
                "MIDP_Deliverable_Type": "",
                "MIDP_Asset_Count": 0,
                "Has_Zero_Prefixed_Final_Sequence": False,
                "Present_In_All_Current_DM3": False,
                "Solid_In_All_Current_DM3": False,
                "Matched_Discipline": False,
                "Contains_Solid": False,
                "Resolved_Model_Container_ID": "",
                "Resolution_Status": "No MIDP DM3 match with matching discipline",
            }]
        resolved = self.resolve(uaid2, discipline)
        records: list[dict[str, Any]] = []
        for _, row in candidates.iterrows():
            deliverable = str(row.get(self.deliverable_col, "")).strip()
            midp_disc = ""
            if self.discipline_col and not _is_empty(row.get(self.discipline_col)):
                midp_disc = self._normalize_discipline_code(row.get(self.discipline_col))
            if not midp_disc:
                midp_disc = self._discipline_from_midp_deliverable(deliverable)
            desc = str(row.get(self.description_col, "")).strip() if self.description_col else ""
            deliverable_type = str(row.get(self.deliverable_type_col, "")).strip() if self.deliverable_type_col else ""
            deliverable_token = Path(deliverable).stem.upper()
            deliverable_upper = deliverable.upper()
            present_in_dm3 = deliverable_token in self.dm3_document_tokens or deliverable_upper in self.dm3_document_tokens
            solid_in_dm3 = deliverable_token in self.dm3_solid_tokens or deliverable_upper in self.dm3_solid_tokens
            records.append({
                "UAID_2": uaid2,
                "MPDT_Discipline": disc,
                "MIDP_Deliverable": deliverable,
                "MIDP_Discipline": midp_disc,
                "Description": desc,
                "MIDP_Deliverable_Type": deliverable_type,
                "MIDP_Asset_Count": len(self._asset_items(row.get(self.asset_col, ""))) if self.asset_col else 0,
                "Has_Zero_Prefixed_Final_Sequence": self._has_zero_prefixed_final_sequence(deliverable),
                "Present_In_All_Current_DM3": present_in_dm3,
                "Solid_In_All_Current_DM3": solid_in_dm3,
                "Matched_Discipline": bool(disc and midp_disc == disc),
                "Contains_Solid": "solid" in desc.lower(),
                "Resolved_Model_Container_ID": resolved,
                "Resolution_Status": "Unique" if resolved else "Ambiguous or no unique model after rules",
            })
        return records

    def write_match_report(self, targets: list[dict[str, Any]], output_dir: Path) -> Path | None:
        """Write temporary MCID matching diagnostics into the Output folder."""
        rows: list[dict[str, Any]] = []
        for target in targets:
            uaid = str(target.get("uaid", "")).strip()
            mpdt_file = target.get("mpdt_file") or target.get("file") or ""
            discipline = self.discipline_from_deliverable(mpdt_file)
            rows.extend(self.candidate_records(uaid, discipline))
        if not rows:
            return None
        output_dir.mkdir(parents=True, exist_ok=True)
        report_path = output_dir / "MCID_MIDP_Matches.xlsx"
        df = pd.DataFrame(rows)
        df = df.loc[:, ~df.columns.duplicated()]
        try:
            df.to_excel(report_path, index=False)
        except Exception:
            report_path = output_dir / "MCID_MIDP_Matches.csv"
            df.to_csv(report_path, index=False)
        return report_path

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


class AllCurrentDM3Resolver:
    """Resolve Model Container IDs from All_Current_DM3_files.xlsx.

    This file is a ProjectWise-style DM3 extract and normally does not carry
    UAID values.  Matching is therefore based on the MPDT deliverable context:
    project, discipline, and spatial CS/CL token from the filename.
    """

    def __init__(self, dm3_df: pd.DataFrame | None, logger: logging.Logger | None = None):
        self.logger = logger
        self.dm3_df = dm3_df.copy() if dm3_df is not None and not dm3_df.empty else pd.DataFrame()
        self.name_col = _pick_column(self.dm3_df, ("Name", "Doc. Reference", "DocumentName", "Document Name", "Deliverable No", "Deliverable Number")) if not self.dm3_df.empty else None
        self.file_col = _pick_column(self.dm3_df, ("File Name", "FileName")) if not self.dm3_df.empty else None
        self.description_col = _pick_column(self.dm3_df, ("Description", "Title", "Folder Description")) if not self.dm3_df.empty else None
        self.updated_col = _pick_column(self.dm3_df, ("File Updated", "Updated", "Created")) if not self.dm3_df.empty else None

        if not self.dm3_df.empty and not (self.name_col or self.file_col):
            if self.logger:
                self.logger.warning(
                    "All Current DM3 file loaded but no Name/File Name column was found. "
                    "Model Container ID will fall back to MIDP resolver."
                )
        if not self.dm3_df.empty and (self.name_col or self.file_col):
            self._prepare_lookup_columns()

    @staticmethod
    def _stem(value: Any) -> str:
        if _is_empty(value):
            return ""
        return Path(str(value).strip()).stem.upper()

    @staticmethod
    def _parts(value: Any) -> list[str]:
        stem = AllCurrentDM3Resolver._stem(value)
        return [p.strip().upper() for p in stem.split("-") if p.strip()]

    @staticmethod
    def _context(deliverable_name: str | None) -> dict[str, str]:
        parts = AllCurrentDM3Resolver._parts(deliverable_name)
        return {
            "project": parts[0] if len(parts) > 0 else "",
            "originator": parts[1] if len(parts) > 1 else "",
            "discipline": parts[2][:2] if len(parts) > 2 else "",
            "spatial": parts[4] if len(parts) > 4 else "",
        }

    def _row_stem(self, row: pd.Series) -> str:
        for col in (self.name_col, self.file_col):
            if col and not _is_empty(row.get(col)):
                return self._stem(row.get(col))
        return ""

    def _prepare_lookup_columns(self) -> None:
        def row_stem(row: pd.Series) -> str:
            return self._row_stem(row)

        self.dm3_df["_dm3_stem"] = self.dm3_df.apply(row_stem, axis=1)
        parts = self.dm3_df["_dm3_stem"].map(self._parts)
        self.dm3_df["_dm3_project"] = parts.map(lambda p: p[0] if len(p) > 0 else "")
        self.dm3_df["_dm3_discipline"] = parts.map(lambda p: p[2][:2] if len(p) > 2 else "")
        self.dm3_df["_dm3_spatial"] = parts.map(lambda p: p[4] if len(p) > 4 else "")
        self.dm3_df["_dm3_is_dm3"] = parts.map(lambda p: "DM3" in p)

    def _candidate_rows(self, deliverable_name: str | None) -> pd.DataFrame:
        if self.dm3_df.empty or not (self.name_col or self.file_col):
            return pd.DataFrame()
        ctx = self._context(deliverable_name)
        if not ctx["project"] or not ctx["discipline"] or not ctx["spatial"]:
            return pd.DataFrame()

        required_cols = {"_dm3_is_dm3", "_dm3_project", "_dm3_discipline", "_dm3_spatial"}
        if not required_cols.issubset(self.dm3_df.columns):
            self._prepare_lookup_columns()
        mask = (
            self.dm3_df["_dm3_is_dm3"]
            & (self.dm3_df["_dm3_project"] == ctx["project"])
            & (self.dm3_df["_dm3_discipline"] == ctx["discipline"])
            & (self.dm3_df["_dm3_spatial"] == ctx["spatial"])
        )
        return self.dm3_df[mask].copy()

    def _unique_name(self, df: pd.DataFrame) -> str:
        if df is None or df.empty:
            return ""
        vals: list[str] = []
        seen: set[str] = set()
        for _, row in df.iterrows():
            stem = self._row_stem(row)
            if not stem:
                continue
            key = stem.upper()
            if key not in seen:
                vals.append(stem)
                seen.add(key)
        return vals[0] if len(vals) == 1 else ""

    @staticmethod
    def _match_tokens(value: Any) -> list[str]:
        if _is_empty(value):
            return []
        stop_words = {
            "asset", "assets", "line", "model", "models", "civil", "native",
            "and", "the", "for", "with", "from", "into", "level", "scope",
        }
        tokens = [t.lower() for t in re.split(r"[^A-Za-z0-9]+", str(value)) if len(t) >= 3]
        return [t for t in tokens if t not in stop_words]

    def _filter_by_asset_name(self, df: pd.DataFrame, asset_name: str | None) -> pd.DataFrame:
        if df is None or df.empty or not self.description_col or _is_empty(asset_name):
            return df
        tokens = self._match_tokens(asset_name)
        if not tokens:
            return df
        desc = df[self.description_col].fillna("").astype(str).str.lower()
        mask = pd.Series(True, index=df.index)
        for token in tokens:
            mask = mask & desc.str.contains(re.escape(token), na=False)
        filtered = df[mask]
        return filtered if not filtered.empty else df

    def _prefer_active_non_native(self, df: pd.DataFrame) -> pd.DataFrame:
        if df is None or df.empty or not self.description_col:
            return df
        desc = df[self.description_col].fillna("").astype(str).str.lower()
        active = df[~desc.str.contains("withdrawn|not in use", case=False, na=False)]
        if not active.empty:
            df = active
            desc = df[self.description_col].fillna("").astype(str).str.lower()
        non_native = df[~desc.str.contains("civil 3d native", case=False, na=False)]
        return non_native if not non_native.empty else df

    def resolve(self, deliverable_name: str | None, asset_name: str | None = None) -> str:
        """Return a unique DM3 container ID for the MPDT deliverable, or blank."""
        candidates = self._candidate_rows(deliverable_name)
        candidates = self._filter_by_asset_name(candidates, asset_name)

        one = self._unique_name(candidates)
        if one:
            return one

        if not candidates.empty and self.description_col:
            desc = candidates[self.description_col].fillna("").astype(str)
            solid = candidates[desc.str.contains("solid", case=False, na=False)]
            one = self._unique_name(solid)
            if one:
                return one

            preferred = self._prefer_active_non_native(candidates)
            one = self._unique_name(preferred)
            if one:
                return one

        return ""


def build_all_current_dm3_resolver(dm3_df: pd.DataFrame | None, logger: logging.Logger | None = None) -> AllCurrentDM3Resolver:
    return AllCurrentDM3Resolver(dm3_df, logger)

def build_model_container_resolver(
    midp_df: pd.DataFrame | None,
    pw_df: pd.DataFrame | None,
    dm3_df: pd.DataFrame | None = None,
    logger: logging.Logger | None = None,
) -> ModelContainerResolver:
    return ModelContainerResolver(midp_df, pw_df, dm3_df, logger)
