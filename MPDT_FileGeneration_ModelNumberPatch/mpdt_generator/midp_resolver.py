"""Resolve MPDT Model Container IDs from MIDP Navigator data.

The resolver implements the business rule used by MPDT generation:
  1. Start from MIDP_Navigator.xlsx rows whose Assets column contains UAID_2.
  2. Keep only DM3 deliverables.
  3. If several remain, prefer rows whose description/name contains "solid".
  4. If several still remain, cross-check against ACBOS MPDT / ProjectWise extract.
  5. Return one unique deliverable number, otherwise blank.
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

    def __init__(self, midp_df: pd.DataFrame | None, pw_df: pd.DataFrame | None = None, logger: logging.Logger | None = None):
        self.logger = logger
        self.midp_df = midp_df.copy() if midp_df is not None and not midp_df.empty else pd.DataFrame()
        self.asset_col = _pick_column(self.midp_df, ("Assets", "Asset", "Asset ID", "Asset_ID", "UAID", "UAID_2")) if not self.midp_df.empty else None
        self.deliverable_col = _pick_column(self.midp_df, ("Deliverable No", "Deliverable Number", "DocumentName", "Document Name")) if not self.midp_df.empty else None
        self.description_col = _pick_column(self.midp_df, ("Description", "Deliverable Name", "Deliverable Description", "Name")) if not self.midp_df.empty else None
        self.discipline_col = _pick_column(self.midp_df, ("Discipline", "Disc", "Discipline Code", "Discipline_Code")) if not self.midp_df.empty else None
        self.pw_document_tokens = _document_tokens_from_pw_extract(pw_df)

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
        deliverable_text = candidates[self.deliverable_col].fillna("").astype(str)
        return candidates[deliverable_text.str.contains("DM3", case=False, na=False)]

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
                "Matched_Discipline": False,
                "Contains_Solid": False,
                "Resolved_Model_Container_ID": "",
                "Resolution_Status": "No DM3 MIDP match",
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
            records.append({
                "UAID_2": uaid2,
                "MPDT_Discipline": disc,
                "MIDP_Deliverable": deliverable,
                "MIDP_Discipline": midp_disc,
                "Description": desc,
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

def build_model_container_resolver(midp_df: pd.DataFrame | None, pw_df: pd.DataFrame | None, logger: logging.Logger | None = None) -> ModelContainerResolver:
    return ModelContainerResolver(midp_df, pw_df, logger)
