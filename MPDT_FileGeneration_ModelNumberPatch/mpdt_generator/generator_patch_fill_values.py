# generator_patch_fill_values.py
import pandas as pd
from typing import Dict, List

def _safe_get(df: pd.DataFrame, row_key: str, key_col: str, value_col: str) -> str:
    if df is None or df.empty or not key_col or not value_col:
        return ""
    m = df[df[key_col].fillna("").astype(str).str.strip() == str(row_key).strip()]
    if m.empty:
        return ""
    v = str(m.iloc[0].get(value_col, "")).strip()
    return v

def fill_au_values_for_row(ws, row_idx: int, start_col_idx: int, row1_headers: List[str], join2_wide: pd.DataFrame, join2_key_col: str, join2_key_value: str):
    """
    For each AU+ column, take Row1 header (att short code) as a column in join2_wide
    and write the cell value for the current row's UAID_3 (or key).
    - ws: openpyxl worksheet
    - row_idx: row index where to write values (data row)
    - start_col_idx: AU column index
    - row1_headers: list of AU Row1 short codes
    - join2_wide: a wide table keyed by UAID_3 (or similar) with att short codes as columns
    - join2_key_col: name of the key column in join2_wide (e.g., 'UAID_3')
    - join2_key_value: the key value for the current row to fetch
    """
    # Fast slice
    row_df = join2_wide[join2_wide[join2_key_col].fillna("").astype(str).str.strip() == str(join2_key_value).strip()]
    has_row = not row_df.empty
    for i, code in enumerate(row1_headers, start=start_col_idx):
        val = ""
        if has_row and code in row_df.columns:
            raw = row_df.iloc[0][code]
            val = "" if pd.isna(raw) else str(raw).strip()
        ws.cell(row=row_idx, column=i, value=val)