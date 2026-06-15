# generator_patch_headers_apply.py
from typing import List, Tuple

def _build_final_au_headers_from_att_matrix(template_path: str, lodm_df: pd.DataFrame, first15_permanent: int = 15) -> Tuple[List[str], List[str]]:
    """
    Returns (row1_headers, row2_headers) for AU+ after applying rules.
    """
    parts = build_au_headers_from_att_matrix(template_path, first15_permanent=first15_permanent)
    row1_perm = parts["row1_perm"]
    row2_perm = parts["row2_perm"]

    # Filter dynamic by LoDM
    row1_dyn_all = parts["row1_dynamic"]
    row2_dyn_all = parts["row2_dynamic"]
    row1_dyn_kept = filter_dynamic_au_by_lodm(row1_dyn_all, lodm_df)

    # Rebuild row2 dynamic for the kept row1 entries (positions preserved)
    row2_dyn_kept = []
    for r1 in row1_dyn_kept:
        idx = row1_dyn_all.index(r1)
        row2_dyn_kept.append(row2_dyn_all[idx])

    row1_final = row1_perm + row1_dyn_kept
    row2_final = row2_perm + row2_dyn_kept
    return row1_final, row2_final

def write_au_headers(ws, start_col_idx: int, row1_headers: List[str], row2_headers: List[str]):
    """
    Writes AU+ headers to the openpyxl worksheet ws:
      - Row 1 (template row for AttTypeName abbreviations)
      - Row 2 (display names)
    The caller must provide where AU is (start_col_idx).
    """
    # Assuming MPDT template uses headers in rows 1 and 2. Adjust if your template differs.
    r1 = 1
    r2 = 2
    for i, (h1, h2) in enumerate(zip(row1_headers, row2_headers), start=start_col_idx):
        ws.cell(row=r1, column=i, value=h1)
        ws.cell(row=r2, column=i, value=h2)