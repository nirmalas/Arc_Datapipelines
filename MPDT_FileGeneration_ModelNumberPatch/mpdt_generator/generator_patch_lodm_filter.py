# generator_patch_lodm_filter.py
import re
from typing import Set
import pandas as pd

def _norm_code_variants(code: str) -> set[str]:
    s = str(code or "").strip()
    if not s:
        return set()
    return {s.lower(), s.replace("_", "-").lower(), s.replace("-", "_").lower()}

_RE_ATT_SEP = re.compile(r'[\s_\-\.:]+')
def _norm_att_code(s: str) -> str:
    return _RE_ATT_SEP.sub('', str(s or "")).lower()

def _lodm_attnorms_all(lodm_df: pd.DataFrame) -> Set[str]:
    if lodm_df is None or lodm_df.empty:
        return set()
    att_col = next((c for c in lodm_df.columns if str(c).strip().lower() == "atttypename".lower()), None)
    if not att_col:
        return set()
    return {_norm_att_code(v) for v in lodm_df[att_col].dropna().astype(str) if str(v).strip()}

def filter_dynamic_au_by_lodm(att_row1_dynamic: list[str], lodm_df: pd.DataFrame) -> list[str]:
    """
    Keep only dynamic AU Row1 entries that map to at least one LoDM AttTypeName.
    Matching uses normalized AttTypeName codes (strip separators and lowercase).
    """
    lodm_norms = _lodm_attnorms_all(lodm_df)
    if not lodm_norms:
        # If No LoDM present, keep dynamic as-is (or return empty to be conservative)
        return att_row1_dynamic
    out: list[str] = []
    for label in att_row1_dynamic:
        if not label:
            continue
        if _norm_att_code(label) in lodm_norms:
            out.append(label)
    return out