from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Dict, Optional

import pandas as pd


REV_RE = re.compile(r"^\s*P(\d{2,})\s*$", re.IGNORECASE)


def norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


def pick_col(df: pd.DataFrame, candidates: list[str]) -> Optional[str]:
    nmap = {norm(c): c for c in df.columns}
    for cand in candidates:
        c = nmap.get(norm(cand))
        if c:
            return c
    return None


def next_revision(existing_versions: list[str]) -> str:
    max_rev = 0
    for v in existing_versions:
        if not isinstance(v, str):
            continue
        m = REV_RE.match(v)
        if m:
            max_rev = max(max_rev, int(m.group(1)))
    if max_rev <= 0:
        return "P01"
    return f"P{max_rev + 1:02d}"


def build_existing_index(acbos_extract: Optional[Path]) -> Dict[str, dict]:
    if not acbos_extract or not acbos_extract.exists():
        return {}

    try:
        df = pd.read_excel(acbos_extract, dtype=str)
    except PermissionError:
        # If the extract is open/locked, continue with empty existing index.
        return {}
    if df.empty:
        return {}

    name_col = pick_col(df, ["Name", "DocumentName", "Document Name", "FileName", "File Name"])
    version_col = pick_col(df, ["Version", "Revision"])
    folder_col = pick_col(df, ["FolderPath", "Folder Path", "Path"])
    urn_col = pick_col(df, ["URN", "DocumentURN", "Document URN"])
    guid_col = pick_col(df, ["DocumentGUID", "Document GUID"])

    if not name_col:
        return {}

    idx: Dict[str, dict] = {}
    for _, row in df.iterrows():
        name = str(row.get(name_col, "") or "").strip()
        if not name:
            continue
        key = name.lower()
        version = str(row.get(version_col, "") or "").strip() if version_col else ""
        folder = str(row.get(folder_col, "") or "").strip() if folder_col else ""
        urn = str(row.get(urn_col, "") or "").strip() if urn_col else ""
        guid = str(row.get(guid_col, "") or "").strip() if guid_col else ""

        item = idx.setdefault(
            key,
            {
                "versions": [],
                "folder_path": folder,
                "document_urn": urn,
                "document_guid": guid,
            },
        )
        if version:
            item["versions"].append(version)
        if folder and not item["folder_path"]:
            item["folder_path"] = folder
        if urn and not item["document_urn"]:
            item["document_urn"] = urn
        if guid and not item["document_guid"]:
            item["document_guid"] = guid

    return idx


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate ProjectWise upload manifest for MPDT/ACBOS files.")
    parser.add_argument("--output-dir", default="Output", help="Directory containing generated files to upload.")
    parser.add_argument(
        "--acbos-extract",
        default="Input/ACBOS MPDT.xlsx",
        help="Path to ACBOS MPDT extract for existing document metadata.",
    )
    parser.add_argument(
        "--default-folder",
        default="",
        help="Default ProjectWise folder for new files when no existing folder is known.",
    )
    parser.add_argument(
        "--manifest-name",
        default="pw_upload_manifest.xlsx",
        help="Manifest filename written into output directory.",
    )
    args = parser.parse_args()

    out_dir = Path(args.output_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    acbos_extract = Path(args.acbos_extract).resolve()

    files = sorted([p for p in out_dir.rglob("*") if p.suffix.lower() in {".xlsm", ".xlsx", ".xls", ".acbos"}])
    if not files:
        print(f"No uploadable files found in {out_dir}")
        return 1

    existing = build_existing_index(acbos_extract if acbos_extract.exists() else None)

    rows = []
    for f in files:
        name = f.name
        key = name.lower()
        ex = existing.get(key, {})
        versions = ex.get("versions", [])
        is_existing = bool(ex)
        target_version = next_revision(versions) if is_existing else "P01"

        folder_path = ex.get("folder_path") or args.default_folder
        rows.append(
            {
                "local_file_path": str(f),
                "document_name": name,
                "description": name,
                "version": target_version,
                "application": "",
                "pw_folder_path": folder_path,
                "is_existing": is_existing,
                "existing_document_guid": ex.get("document_guid", ""),
                "existing_document_urn": ex.get("document_urn", ""),
                "existing_versions_found": ", ".join(versions),
                "attributes_json": "",
                "upload_enabled": True,
                "notes": "",
            }
        )

    manifest = pd.DataFrame(rows)
    manifest_path = out_dir / args.manifest_name
    manifest.to_excel(manifest_path, index=False)
    print(f"Wrote manifest: {manifest_path}")
    print(f"Rows: {len(manifest)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
