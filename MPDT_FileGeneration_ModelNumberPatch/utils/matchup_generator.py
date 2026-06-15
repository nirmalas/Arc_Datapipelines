"""
utils/matchup_generator.py — Auto-generate a UAID-to-MPDT matchup file from the PW extract.

Automates the manual process:
  Step 1 : Filter PW extract by discipline (3rd segment of DocumentName or DC_ROLE_CODE)
           and Description starting with 'MPDT' / 'MPDT -'.
  Step 2 : Strip 'MPDT ', 'MPDT - ' prefix (or suffix) from Description → comparable name.
  Step 3 : Exact-match that stripped name against 'Level 2 Asset Name' in the L2 UAID file.
  Step 4 : Fuzzy fallback for common differences (dashes, missing 'Cutting'/'Embankment', etc.).
  Step 5 : Repeat steps 1–4 for Descriptions *ending* with ' MPDT'.
  Step 6 : Repeat steps 1–4 for the 'MDPT' typo variant.

Output columns: Level 2 UAID | Level 2 Asset Name | MPDT (= DocumentName)

Usage (standalone):
  python -m utils.matchup_generator --discipline EV --output "Input/My Matchup.xlsx"
  python -m utils.matchup_generator --discipline EV CS BR --output "Input/My Matchup.xlsx"

Can also be imported and called from main.py / pipeline steps.
"""
from __future__ import annotations

import argparse
import difflib
import logging
import re
from pathlib import Path
from typing import Any

import pandas as pd

from utils.common import load_config, normalize_text, resolve_workspace, setup_logger


# ---------------------------------------------------------------------------
# Name normalisation helpers
# ---------------------------------------------------------------------------

# Words that may appear in PW descriptions but not in L2 names (or vice-versa).
_VARIANT_WORDS = re.compile(
    r"\b(cutting|embankment|fill|bridge|viaduct|overbridge|underbridge|culvert|tunnel)\b",
    re.IGNORECASE,
)


def _norm_for_match(text: str) -> str:
    """Normalise a name for matching: lowercase, collapse spaces, drop dashes."""
    s = str(text or "").strip()
    s = s.replace("-", " ").replace("_", " ")
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s


def _norm_no_variants(text: str) -> str:
    """Like _norm_for_match but also strips structural variant words."""
    s = _norm_for_match(text)
    s = _VARIANT_WORDS.sub("", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


# ---------------------------------------------------------------------------
# Strip MPDT prefix / suffix from a Description value
# ---------------------------------------------------------------------------

_MPDT_PREFIX = re.compile(r"^(MDPT|MPDT)\s*[-–]\s*|^(MDPT|MPDT)\s+", re.IGNORECASE)
_MPDT_SUFFIX = re.compile(r"\s*[-–]\s*(MDPT|MPDT)$|\s+(MDPT|MPDT)$", re.IGNORECASE)
_ACBOS_PREFIX = re.compile(r"^(ACBOS)\s*[-–]\s*|^(ACBOS)\s+", re.IGNORECASE)
_ACBOS_SUFFIX = re.compile(r"\s*[-–]\s*(ACBOS)$|\s+(ACBOS)$", re.IGNORECASE)
# Trailing document-type code (e.g. ' - FRM', ' - REP', ' - BRG') — common in ACBOS descriptions.
_DOC_TYPE_SUFFIX = re.compile(r"\s*[-–]\s*[A-Z]{2,5}$")


def _strip_mpdt(description: str) -> str:
    """Remove MPDT/MDPT prefix or suffix from a PW Description string."""
    s = str(description or "").strip()
    s = _MPDT_PREFIX.sub("", s).strip()
    s = _MPDT_SUFFIX.sub("", s).strip()
    return s


def _strip_acbos(description: str) -> str:
    """Remove ACBOS prefix/suffix and trailing doc-type code from a PW Description string.

    'ACBOS - Glebe House Listed Building - FRM' → 'Glebe House Listed Building'
    """
    s = str(description or "").strip()
    s = _ACBOS_PREFIX.sub("", s).strip()
    s = _ACBOS_SUFFIX.sub("", s).strip()
    s = _DOC_TYPE_SUFFIX.sub("", s).strip()
    return s


def _is_mpdt_description(description: str, variant: str = "prefix") -> bool:
    """
    Return True if the description is an MPDT document name.

    variant='prefix' : Description starts with MPDT/MDPT
    variant='suffix' : Description ends with MPDT/MDPT
    variant='any'    : either
    """
    desc = str(description or "").strip()
    has_prefix = bool(_MPDT_PREFIX.match(desc)) or desc.upper().startswith("MPDT") or desc.upper().startswith("MDPT")
    has_suffix = bool(_MPDT_SUFFIX.search(desc))
    if variant == "prefix":
        return has_prefix
    if variant == "suffix":
        return has_suffix
    return has_prefix or has_suffix


def _is_acbos_description(description: str, variant: str = "prefix") -> bool:
    """
    Return True if the description is an ACBOS document name.

    variant='prefix' : Description starts with ACBOS
    variant='suffix' : Description ends with ACBOS
    variant='any'    : either
    """
    desc = str(description or "").strip()
    has_prefix = bool(_ACBOS_PREFIX.match(desc)) or desc.upper().startswith("ACBOS")
    has_suffix = bool(_ACBOS_SUFFIX.search(desc))
    if variant == "prefix":
        return has_prefix
    if variant == "suffix":
        return has_suffix
    return has_prefix or has_suffix


# ---------------------------------------------------------------------------
# Main matching logic
# ---------------------------------------------------------------------------

def _best_doc_for_name(pw_candidates: pd.DataFrame) -> pd.Series:
    """
    Given multiple PW rows for the same stripped name, return the row with the
    latest version (sorted by Version then FileUpdated, descending).
    """
    if len(pw_candidates) == 1:
        return pw_candidates.iloc[0]

    tmp = pw_candidates.copy()
    if "FileUpdated" in tmp.columns:
        tmp["__date__"] = pd.to_datetime(tmp["FileUpdated"], errors="coerce")
    else:
        tmp["__date__"] = pd.NaT

    # Prefer rows with a numeric version like P02, P01 — parse version digit.
    def _ver_num(v: Any) -> int:
        m = re.search(r"\d+", str(v or ""))
        return int(m.group()) if m else 0

    if "Version" in tmp.columns:
        tmp["__ver__"] = tmp["Version"].apply(_ver_num)
    else:
        tmp["__ver__"] = 0

    tmp = tmp.sort_values(["__ver__", "__date__"], ascending=False, na_position="last")
    return tmp.iloc[0]


def _build_fuzzy_lookup(l2_names: list[str]) -> tuple[dict[str, str], list[str], dict[str, str], list[str]]:
    """
    Pre-compute normalised versions of all L2 names for fast matching.

    Returns two lookup dicts and two sorted norm lists:
      (norm_to_raw, norm_names, novar_to_raw, novar_names)
    """
    norm_to_raw: dict[str, str] = {}
    novar_to_raw: dict[str, str] = {}
    for n in l2_names:
        norm_to_raw[_norm_for_match(n)] = n
        novar_to_raw[_norm_no_variants(n)] = n
    return norm_to_raw, list(norm_to_raw.keys()), novar_to_raw, list(novar_to_raw.keys())


def _fuzzy_match(
    stripped: str,
    l2_names: list[str],
    threshold: float = 0.80,
    *,
    _cache: dict | None = None,
) -> tuple[str, float]:
    """
    Return the best-matching L2 asset name and its score.
    Uses difflib.get_close_matches (O(n) heap-based) instead of brute-force.
    Tries progressively looser normalisation:
      1. Full normalised match (dashes removed)
      2. Drop structural variant words and retry
    """
    # Pre-compute normalised L2 lists once per call batch (stored in _cache).
    if _cache is None:
        _cache = {}
    if "norm_to_raw" not in _cache:
        nt, nl, nv, nvl = _build_fuzzy_lookup(l2_names)
        _cache["norm_to_raw"] = nt
        _cache["norm_names"] = nl
        _cache["novar_to_raw"] = nv
        _cache["novar_names"] = nvl

    norm_stripped = _norm_for_match(stripped)
    novar_stripped = _norm_no_variants(stripped)

    # Pass 1: full normalisation
    hits1 = difflib.get_close_matches(norm_stripped, _cache["norm_names"], n=3, cutoff=threshold)
    if hits1:
        best_norm = hits1[0]
        score = difflib.SequenceMatcher(None, norm_stripped, best_norm).ratio()
        return _cache["norm_to_raw"][best_norm], round(score, 3)

    # Pass 2: drop variant words (allows "Cutting Embankment" mismatches)
    hits2 = difflib.get_close_matches(novar_stripped, _cache["novar_names"], n=3, cutoff=threshold)
    if hits2:
        best_norm = hits2[0]
        score = difflib.SequenceMatcher(None, novar_stripped, best_norm).ratio()
        return _cache["novar_to_raw"][best_norm], round(score, 3)

    # Nothing above threshold — return best scoring candidate for diagnostic logging
    best_name, best_score = "", 0.0
    for nm in _cache["norm_names"]:
        s = difflib.SequenceMatcher(None, norm_stripped, nm).ratio()
        if s > best_score:
            best_score, best_name = s, _cache["norm_to_raw"][nm]
    return "", round(best_score, 3)


def generate_matchup(
    pw_df: pd.DataFrame,
    l2_df: pd.DataFrame,
    disciplines: list[str],
    fuzzy_threshold: float = 0.80,
    extra_uaid_types: dict[str, str] | None = None,
    logger: logging.Logger | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Core matching function for both MPDT and ACBOS documents.

    Parameters
    ----------
    pw_df            : PW extract DataFrame (DocumentName, Description, Version, FileUpdated, …)
    l2_df            : L2 UAID file (UAID_2, Level 2 Asset Name)
    disciplines      : list of discipline codes to filter (e.g. ['EV', 'EV CS'])
    fuzzy_threshold  : minimum fuzzy match ratio to accept (0–1)
    extra_uaid_types : optional {UAID_2: "MPDT"/"ACBOS"} for UAIDs from config target_uaid2
                       that are not found via PW description matching. They are appended to the
                       output so the pipeline always processes them.

    Returns
    -------
    (matched_df, unmatched_df)
    matched_df columns: Level 2 UAID | Level 2 Asset Name | MPDT | file_type | match_score | match_type
    """
    if logger is None:
        logger = logging.getLogger(__name__)

    # ------------------------------------------------------------------ #
    # 1. Normalise L2 names → lookup dict                                  #
    # ------------------------------------------------------------------ #
    uaid_col = next((c for c in l2_df.columns if normalize_text(c) in ("uaid 2", "uaid2", "uaid_2")), None)
    name_col = next((c for c in l2_df.columns if normalize_text(c) in ("level 2 asset name", "asset name", "assetname")), None)

    if not uaid_col or not name_col:
        raise ValueError(f"L2 file must have UAID_2 and Level 2 Asset Name columns. Got: {l2_df.columns.tolist()}")

    l2_clean = l2_df[[uaid_col, name_col]].dropna(subset=[uaid_col, name_col]).copy()
    l2_clean[uaid_col] = l2_clean[uaid_col].astype(str).str.strip()
    l2_clean[name_col] = l2_clean[name_col].astype(str).str.strip()
    l2_names = l2_clean[name_col].tolist()
    l2_lookup_exact: dict[str, str] = {}   # normalised_name → UAID_2
    l2_lookup_raw: dict[str, str] = {}     # raw_name → UAID_2
    l2_uaid_to_name: dict[str, str] = {}   # UAID_2 → raw_name
    for _, r in l2_clean.iterrows():
        l2_lookup_raw[r[name_col]] = r[uaid_col]
        l2_lookup_exact[_norm_for_match(r[name_col])] = r[uaid_col]
        l2_uaid_to_name[r[uaid_col]] = r[name_col]

    logger.info("  L2 file: %d unique asset names", len(l2_names))

    # Shared fuzzy cache — normalised L2 lists computed only once.
    fuzzy_cache: dict = {}

    # ------------------------------------------------------------------ #
    # 2. Identify discipline from 3rd segment of DocumentName              #
    # ------------------------------------------------------------------ #
    disc_upper = {d.strip().upper() for d in disciplines}

    def _doc_discipline(doc_name: str) -> str:
        parts = str(doc_name or "").split("-")
        return parts[2].strip().upper() if len(parts) > 2 else ""

    has_role_col = "DC_ROLE_CODE" in pw_df.columns

    def _is_target_discipline(row: pd.Series) -> bool:
        d = _doc_discipline(str(row.get("DocumentName", "")))
        if d in disc_upper:
            return True
        if has_role_col:
            rc = str(row.get("DC_ROLE_CODE", "") or "").strip().upper()
            if rc in disc_upper:
                return True
        return False

    pw_disc = pw_df[pw_df.apply(_is_target_discipline, axis=1)].copy()
    logger.info("  PW rows for discipline %s: %d", disciplines, len(pw_disc))

    # ------------------------------------------------------------------ #
    # 3. Inner helper: match a set of candidates (MPDT or ACBOS)           #
    # ------------------------------------------------------------------ #
    def _match_doc_type(
        rows_prefix: pd.DataFrame,
        rows_suffix: pd.DataFrame,
        strip_fn,
        doc_file_type: str,
        type_label: str,
    ) -> tuple[list[dict], list[dict]]:
        rows_combined = pd.concat(
            [rows_prefix.assign(_variant="prefix"), rows_suffix.assign(_variant="suffix")],
            ignore_index=True,
        )
        if rows_combined.empty:
            logger.info("  %s candidate rows: 0", type_label)
            return [], []

        logger.info("  %s candidate rows: prefix=%d  suffix=%d", type_label, len(rows_prefix), len(rows_suffix))
        rows_combined["_stripped"] = rows_combined["Description"].apply(strip_fn)
        stripped_groups = rows_combined.groupby("_stripped", sort=False)
        logger.info("  %s unique stripped names: %d", type_label, len(stripped_groups))

        _matched: list[dict] = []
        _unmatched: list[dict] = []

        for stripped, group in stripped_groups:
            if not stripped:
                continue
            best_row = _best_doc_for_name(group)
            doc_name = str(best_row.get("DocumentName", "")).strip()

            norm = _norm_for_match(stripped)
            uaid = l2_lookup_exact.get(norm, "")
            if uaid:
                l2_name = l2_uaid_to_name.get(uaid, stripped)
                _matched.append({
                    "Level 2 UAID": uaid,
                    "Level 2 Asset Name": l2_name,
                    "MPDT": doc_name,
                    "file_type": doc_file_type,
                    "match_score": 1.0,
                    "match_type": "exact",
                })
                continue

            best_l2, score = _fuzzy_match(stripped, l2_names, threshold=fuzzy_threshold, _cache=fuzzy_cache)
            if best_l2:
                uaid = l2_lookup_raw.get(best_l2, "")
                _matched.append({
                    "Level 2 UAID": uaid,
                    "Level 2 Asset Name": best_l2,
                    "MPDT": doc_name,
                    "file_type": doc_file_type,
                    "match_score": round(score, 3),
                    "match_type": "fuzzy",
                })
            else:
                _unmatched.append({
                    "stripped_name": stripped,
                    "DocumentName": doc_name,
                    "file_type": doc_file_type,
                    "best_score": round(score, 3),
                })

        logger.info(
            "  %s Matched: %d  Unmatched: %d  (fuzzy threshold=%.0f%%)",
            type_label, len(_matched), len(_unmatched), fuzzy_threshold * 100,
        )
        return _matched, _unmatched

    # ------------------------------------------------------------------ #
    # 4. MPDT candidates                                                    #
    # ------------------------------------------------------------------ #
    mpdt_prefix = pw_disc[pw_disc["Description"].fillna("").apply(lambda d: _is_mpdt_description(d, "prefix"))]
    mpdt_suffix = pw_disc[
        pw_disc["Description"].fillna("").apply(lambda d: _is_mpdt_description(d, "suffix"))
        & ~pw_disc["Description"].fillna("").apply(lambda d: _is_mpdt_description(d, "prefix"))
    ]
    mpdt_matched, mpdt_unmatched = _match_doc_type(mpdt_prefix, mpdt_suffix, _strip_mpdt, "MPDT", "MPDT")

    # ------------------------------------------------------------------ #
    # 5. ACBOS candidates                                                   #
    # ------------------------------------------------------------------ #
    acbos_prefix = pw_disc[pw_disc["Description"].fillna("").apply(lambda d: _is_acbos_description(d, "prefix"))]
    acbos_suffix = pw_disc[
        pw_disc["Description"].fillna("").apply(lambda d: _is_acbos_description(d, "suffix"))
        & ~pw_disc["Description"].fillna("").apply(lambda d: _is_acbos_description(d, "prefix"))
    ]
    acbos_matched, acbos_unmatched = _match_doc_type(acbos_prefix, acbos_suffix, _strip_acbos, "ACBOS", "ACBOS")

    # ------------------------------------------------------------------ #
    # 6. Combine — prefer ACBOS over MPDT when same UAID appears in both   #
    # ------------------------------------------------------------------ #
    all_matched = mpdt_matched + acbos_matched
    all_unmatched = mpdt_unmatched + acbos_unmatched

    # If same UAID matched both MPDT and ACBOS, keep the ACBOS entry (ACBOS wins).
    seen_uaids: set[str] = set()
    deduped: list[dict] = []
    # Process ACBOS first so it takes priority.
    for row in sorted(all_matched, key=lambda r: 0 if r["file_type"] == "ACBOS" else 1):
        u = row["Level 2 UAID"]
        if u not in seen_uaids:
            seen_uaids.add(u)
            deduped.append(row)

    # ------------------------------------------------------------------ #
    # 7. Extra UAIDs from config target_uaid2 not yet in results           #
    # ------------------------------------------------------------------ #
    if extra_uaid_types:
        for uaid, ftype in extra_uaid_types.items():
            if uaid in seen_uaids:
                continue
            l2_name = l2_uaid_to_name.get(uaid, "")
            # Try to find a PW document for this UAID by matching its L2 name against PW descriptions.
            doc_name = ""
            if l2_name:
                strip_fn = _strip_acbos if ftype.upper() == "ACBOS" else _strip_mpdt
                is_fn = _is_acbos_description if ftype.upper() == "ACBOS" else _is_mpdt_description
                desc_col = pw_disc["Description"].fillna("")
                candidates_mask = desc_col.apply(lambda d: is_fn(d, "any"))
                pw_type_rows = pw_disc[candidates_mask].copy()
                if not pw_type_rows.empty:
                    pw_type_rows["_stripped"] = pw_type_rows["Description"].apply(strip_fn)
                    best_l2, score = _fuzzy_match(l2_name, pw_type_rows["_stripped"].tolist(), threshold=0.75)
                    if best_l2:
                        hit = pw_type_rows[pw_type_rows["_stripped"] == best_l2]
                        if not hit.empty:
                            doc_name = str(_best_doc_for_name(hit).get("DocumentName", "")).strip()
            deduped.append({
                "Level 2 UAID": uaid,
                "Level 2 Asset Name": l2_name,
                "MPDT": doc_name,
                "file_type": ftype.upper(),
                "match_score": 0.0,
                "match_type": "config",
            })
            seen_uaids.add(uaid)
            logger.info("  Extra UAID from config: %s  type=%s  doc=%s", uaid, ftype, doc_name or "(not found in PW)")

    result_df = pd.DataFrame(deduped)
    if not result_df.empty:
        result_df = result_df.sort_values(["file_type", "Level 2 Asset Name"]).reset_index(drop=True)

    logger.info(
        "  Total matched: %d (MPDT=%d ACBOS=%d)  Unmatched: %d",
        len(deduped),
        sum(1 for r in deduped if r["file_type"] == "MPDT"),
        sum(1 for r in deduped if r["file_type"] == "ACBOS"),
        len(all_unmatched),
    )
    if all_unmatched:
        logger.warning("  Unmatched PW descriptions (no L2 asset found):")
        for u in all_unmatched[:20]:
            logger.warning("    [%s] score=%.2f  '%s'  doc=%s",
                           u["file_type"], u["best_score"], u["stripped_name"], u["DocumentName"])
        if len(all_unmatched) > 20:
            logger.warning("    … and %d more", len(all_unmatched) - 20)

    return result_df, pd.DataFrame(all_unmatched)


# ---------------------------------------------------------------------------
# File I/O helpers
# ---------------------------------------------------------------------------

def load_pw_extract(pw_path: Path, logger: logging.Logger) -> pd.DataFrame:
    logger.info("  Loading PW extract: %s", pw_path.name)
    df = pd.read_excel(str(pw_path), dtype=str)
    # Keep latest version per DocumentName (same doc may appear multiple times).
    if "Version" in df.columns and "FileUpdated" in df.columns:
        df["__date__"] = pd.to_datetime(df["FileUpdated"], errors="coerce")
        df["__ver__"] = df["Version"].apply(
            lambda v: int(m.group()) if (m := re.search(r"\d+", str(v or ""))) else 0
        )
        df = df.sort_values(["__ver__", "__date__"], ascending=False).drop_duplicates(
            subset=["DocumentName"], keep="first"
        ).drop(columns=["__date__", "__ver__"])
    logger.info("  PW extract: %d unique documents", len(df))
    return df.reset_index(drop=True)


def load_l2_file(l2_path: Path, logger: logging.Logger) -> pd.DataFrame:
    logger.info("  Loading L2 UAID file: %s", l2_path.name)
    df = pd.read_excel(str(l2_path), dtype=str)
    logger.info("  L2 file: %d rows", len(df))
    return df


def save_matchup(result_df: pd.DataFrame, output_path: Path, logger: logging.Logger) -> None:
    # Output columns: drop internal scoring columns for the clean matchup file.
    out_cols = ["Level 2 UAID", "Level 2 Asset Name", "MPDT"]
    clean = result_df[[c for c in out_cols if c in result_df.columns]]
    clean.to_excel(str(output_path), index=False)
    logger.info("  Matchup file written: %s  (%d rows)", output_path.name, len(clean))


# ---------------------------------------------------------------------------
# Pipeline-callable entry point
# ---------------------------------------------------------------------------

def run_matchup_generator(
    workspace: Path,
    cfg: dict,
    disciplines: list[str],
    output_path: Path | None = None,
    fuzzy_threshold: float = 0.80,
    logger: logging.Logger | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Called from main.py or standalone.

    Returns (matched_df, unmatched_df).
    matched_df has columns: Level 2 UAID | Level 2 Asset Name | MPDT | match_score | match_type
    """
    if logger is None:
        logger = logging.getLogger(__name__)

    paths = cfg.get("paths", {})
    pw_path = (workspace / paths.get("pw_extract", "Input/ACBOS MPDT.xlsx")).resolve()
    l2_path = (workspace / paths.get("l2_uaid_acbos", "Input/L2 UAID-ACBOS_260129.xlsx")).resolve()

    pw_df = load_pw_extract(pw_path, logger)
    l2_df = load_l2_file(l2_path, logger)

    matched_df, unmatched_df = generate_matchup(
        pw_df, l2_df, disciplines,
        fuzzy_threshold=fuzzy_threshold,
        logger=logger,
    )

    if output_path is not None:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        save_matchup(matched_df, output_path, logger)

        if not unmatched_df.empty:
            unmatched_path = output_path.with_name(output_path.stem + "_unmatched.xlsx")
            unmatched_df.to_excel(str(unmatched_path), index=False)
            logger.info("  Unmatched written: %s", unmatched_path.name)

    return matched_df, unmatched_df


# ---------------------------------------------------------------------------
# Standalone CLI
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generate a UAID-to-MPDT matchup file from the PW extract."
    )
    parser.add_argument(
        "--workspace", default=None,
        help="Path to workspace root (default: current directory)",
    )
    parser.add_argument(
        "--discipline", nargs="+", default=["EV"],
        metavar="DISC",
        help="Discipline code(s) to filter (3rd segment of DocumentName). Default: EV",
    )
    parser.add_argument(
        "--output", default=None,
        metavar="PATH",
        help="Output Excel path. Default: Input/<Discipline> MPDT Matchup.xlsx",
    )
    parser.add_argument(
        "--fuzzy-threshold", type=float, default=0.80,
        metavar="0-1",
        help="Minimum fuzzy match ratio (0–1). Default: 0.80",
    )
    args = parser.parse_args()

    workspace = resolve_workspace(args.workspace)
    cfg = load_config(workspace)
    logger = setup_logger(workspace, "matchup_generator", cfg.get("log_level", "INFO"))

    disciplines = args.discipline
    disc_label = "_".join(disciplines)

    output_path = Path(args.output) if args.output else (workspace / "Input" / f"{disc_label} MPDT Matchup.xlsx")

    logger.info("=== Matchup Generator ===")
    logger.info("  Disciplines : %s", disciplines)
    logger.info("  Output      : %s", output_path)
    logger.info("  Fuzzy thresh: %.0f%%", args.fuzzy_threshold * 100)

    matched_df, unmatched_df = run_matchup_generator(
        workspace, cfg, disciplines,
        output_path=output_path,
        fuzzy_threshold=args.fuzzy_threshold,
        logger=logger,
    )

    logger.info("=== Done: %d matched, %d unmatched ===", len(matched_df), len(unmatched_df))


if __name__ == "__main__":
    main()
