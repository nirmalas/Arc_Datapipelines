"""
mpdt_generator/acbos_generator.py — Generate .ACBOS files for target UAID_2 assets.

An .ACBOS file is a ZIP archive containing:
  Setup.BIN   — copied verbatim from template
  Key.BIN     — copied verbatim from template
  Data.XML    — generated from AssetsScope3 rows for the target UAID_2

Can be run standalone:
  python -m mpdt_generator.acbos_generator --workspace . --target-uaid2 HS2-00002NSXW
"""
from __future__ import annotations

import argparse
import logging
import re
import tempfile
import uuid
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Any, NamedTuple
from xml.dom import minidom

import pandas as pd

from utils.common import (
    get_available_path,
    load_config,
    normalize_text,
    read_table_any,
    resolve_workspace,
    sanitize_filename,
    setup_logger,
    strip_acbos_suffix,
    timestamped_dir,
    write_json,
    read_json,
)


# Columns NOT written as <Attribute> elements — structural only
_STRUCTURAL_COLS = {
    "id", "objectid", "ClassificationItemId", "ClsfItemId", "Alignment", "AlgnmPln",
    "Deleted", "ObjectState", "WsGuid", "AssociatedModelFile",
    "source_file", "data_source", "_source", "_source_sheet",
}
_STRUCTURAL_NORMS = {normalize_text(c) for c in _STRUCTURAL_COLS}


def new_guid() -> str:
    return str(uuid.uuid4()).lower()


# ---------------------------------------------------------------------------
# Template extraction
# ---------------------------------------------------------------------------

class AcbosTemplate(NamedTuple):
    """Parsed pieces of a sample/template .ACBOS package."""
    path: Path
    bins: list[tuple[str, Path]]
    attr_metadata: dict[str, dict[str, str]]
    attr_order: list[str]
    classes: set[str]


def extract_template_data(
    template_path: Path,
    extract_dir: Path,
    logger: logging.Logger,
) -> AcbosTemplate:
    """
    Open template ACBOS (ZIP) and extract the parts that must be preserved.

    Important details:
      * Setup.BIN/Key.BIN archive names are preserved exactly as in the sample.
      * Attribute order and DefinitionId/DefinitionWsGuid values are copied from
        the sample Data.XML.
      * Only attributes present in the sample template are later written, so
        generated ACBOS files do not gain extra columns/attributes from source
        DataFrame housekeeping columns.
    """
    bins: list[tuple[str, Path]] = []
    metadata: dict[str, dict[str, str]] = {}
    attr_order: list[str] = []
    classes: set[str] = set()

    template_extract_dir = extract_dir / sanitize_filename(template_path.stem)
    template_extract_dir.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(template_path, "r") as zf:
        for member in zf.namelist():
            name = Path(member).name
            name_up = name.upper()
            if name_up in {"SETUP.BIN", "KEY.BIN"}:
                dest = template_extract_dir / name
                with zf.open(member) as src, open(dest, "wb") as dst:
                    dst.write(src.read())
                bins.append((member, dest))
            elif name_up.endswith(".XML"):
                with zf.open(member) as f:
                    try:
                        tree = ET.parse(f)
                        xml_root = tree.getroot()
                        for obj in xml_root.iter("Object"):
                            cls = (obj.get("ClassificationItemId") or "").strip()
                            if cls:
                                classes.add(cls)
                            for attr in obj.findall("Attribute"):
                                name = attr.get("Name", "").strip()
                                if not name:
                                    continue
                                norm_name = normalize_text(name)
                                if norm_name not in metadata:
                                    attr_order.append(name)
                                    metadata[norm_name] = {
                                        "Name": name,
                                        "DefinitionId": attr.get("DefinitionId", ""),
                                        "DefinitionWsGuid": attr.get("DefinitionWsGuid", ""),
                                    }
                    except ET.ParseError as exc:
                        logger.warning("Could not parse template XML %s: %s", template_path.name, exc)

    logger.info(
        "  Template %s: %d BIN files, %d attribute definitions, %d classes",
        template_path.name, len(bins), len(metadata), len(classes),
    )
    return AcbosTemplate(template_path, bins, metadata, attr_order, classes)


def find_templates(workspace: Path, cfg: dict, logger: logging.Logger) -> list[Path]:
    """Find ACBOS template files, sorted deterministically."""
    tpl_dir = workspace / cfg.get("paths", {}).get("acbos_template_dir", "Input/ACBOS_Templates")
    if tpl_dir.exists():
        candidates = sorted(list(tpl_dir.glob("*.ACBOS")) + list(tpl_dir.glob("*.acbos")))
        if candidates:
            logger.info("Found %d ACBOS template(s) in %s", len(candidates), tpl_dir)
            return candidates
    logger.warning("No ACBOS template found in %s", tpl_dir)
    return []


def _row_classes(rows: pd.DataFrame) -> set[str]:
    """Return non-empty classification values present in rows."""
    classes: set[str] = set()
    for candidate in ("ClassificationItemId", "ClsfItemId", "HS2_Class"):
        norm = normalize_text(candidate)
        col = next((c for c in rows.columns if normalize_text(c) == norm), None)
        if col:
            classes.update(str(v).strip() for v in rows[col].dropna().unique() if str(v).strip())
    return classes


def choose_template_for_rows(
    rows: pd.DataFrame,
    templates: list[AcbosTemplate],
    deliverable_file: str,
    logger: logging.Logger,
) -> AcbosTemplate | None:
    """Choose the closest sample ACBOS template for the target rows."""
    if not templates:
        return None
    if len(templates) == 1:
        return templates[0]

    deliverable_upper = Path(deliverable_file or "").name.upper()
    row_classes = _row_classes(rows)

    def score(tpl: AcbosTemplate) -> tuple[int, int, int]:
        # Prefer templates whose discipline/document token appears in the deliverable name.
        # Example: template "...-BR-..." should beat "...-CV-..." for a BR deliverable.
        name_tokens = re.split(r"[-_ .]+", tpl.path.stem.upper())
        token_score = sum(1 for token in name_tokens if token and token in deliverable_upper)
        class_score = len(row_classes & tpl.classes)
        attr_score = len(tpl.attr_order)
        return (token_score, class_score, attr_score)

    chosen = max(templates, key=score)
    logger.info(
        "  Using ACBOS template %s for classes=%s",
        chosen.path.name, ",".join(sorted(row_classes)) or "<none>",
    )
    return chosen

# ---------------------------------------------------------------------------
# XML builder
# ---------------------------------------------------------------------------

def _is_empty(v) -> bool:
    if v is None:
        return True
    try:
        if pd.isna(v):
            return True
    except Exception:
        pass
    s = str(v).strip()
    return s == "" or s.lower() in ("nan", "none", "nat")


def _value_from_row_by_normalized_name(row: pd.Series, norm_name: str) -> Any:
    for col, val in row.items():
        if normalize_text(col) == norm_name:
            return val
    return None


def build_data_xml(rows: pd.DataFrame, template: AcbosTemplate | None) -> str:
    """
    Serialise rows to Data.XML (ACBOSProjectData format).

    The sample ACBOS Data.XML acts as the schema: generated Object attributes are
    written in the same order and use the same DefinitionId/DefinitionWsGuid.
    Source-only/helper columns are therefore not emitted as extra ACBOS
    attributes.
    """
    root = ET.Element("ACBOSProjectData")
    ET.SubElement(root, "ObjectState").text = "Unchanged"
    ET.SubElement(root, "WsGuid").text = "00000000-0000-0000-0000-000000000000"

    attr_order = template.attr_order if template else []
    attr_metadata = template.attr_metadata if template else {}

    for _, row in rows.iterrows():
        obj_elem = ET.SubElement(root, "Object")

        # Id
        obj_id = ""
        for id_col in ("ObjectId", "objectid", "Id"):
            norm = normalize_text(id_col)
            val = _value_from_row_by_normalized_name(row, norm)
            if not _is_empty(val):
                obj_id = str(val).strip()
                break
        obj_elem.set("Id", obj_id)

        # ClassificationItemId — check ClassificationItemId, then ClsfItemId, then HS2_Class
        cls_id = ""
        for candidate in ("ClassificationItemId", "ClsfItemId", "HS2_Class"):
            val = _value_from_row_by_normalized_name(row, normalize_text(candidate))
            if not _is_empty(val):
                cls_id = str(val).strip()
                break
        obj_elem.set("ClassificationItemId", cls_id)

        # Alignment — check Alignment, then AlgnmPln
        alignment = ""
        for candidate in ("Alignment", "AlgnmPln"):
            val = _value_from_row_by_normalized_name(row, normalize_text(candidate))
            if not _is_empty(val):
                alignment = str(val).strip()
                break
        obj_elem.set("Alignment", alignment)
        obj_elem.set("Deleted", "false")

        ET.SubElement(obj_elem, "ObjectState").text = "Unchanged"
        ET.SubElement(obj_elem, "WsGuid").text = new_guid()

        if attr_order:
            attribute_names = attr_order
        else:
            # Safe fallback if no template XML is available.
            attribute_names = [c for c in rows.columns if normalize_text(c) not in _STRUCTURAL_NORMS]

        for attr_name in attribute_names:
            norm_name = normalize_text(attr_name)
            if norm_name in _STRUCTURAL_NORMS:
                continue
            val = _value_from_row_by_normalized_name(row, norm_name)
            if _is_empty(val):
                continue
            str_val = str(val).strip()
            if not str_val:
                continue

            meta = attr_metadata.get(norm_name, {"Name": attr_name})
            attr_elem = ET.SubElement(obj_elem, "Attribute")
            attr_elem.set("DefinitionId", meta.get("DefinitionId", ""))
            attr_elem.set("Name", meta.get("Name", attr_name))
            attr_elem.set("Value", str_val)
            attr_elem.set("DefinitionWsGuid", meta.get("DefinitionWsGuid", ""))
            ET.SubElement(attr_elem, "ObjectState").text = "Unchanged"
            ET.SubElement(attr_elem, "WsGuid").text = new_guid()

    raw = ET.tostring(root, encoding="unicode")
    dom = minidom.parseString(raw)
    pretty = dom.toprettyxml(indent="  ", encoding="utf-8").decode("utf-8")
    pretty = "\n".join(line for line in pretty.splitlines() if line.strip())
    return pretty


# ---------------------------------------------------------------------------
# Filename resolution
# ---------------------------------------------------------------------------

def resolve_output_filename(
    uaid2: str,
    l2_df: pd.DataFrame,
    pw_df: pd.DataFrame,
    logger: logging.Logger,
) -> str:
    """Determine output .ACBOS filename for a given UAID_2."""
    # From L2 mapping
    if not l2_df.empty and "UAID_2" in l2_df.columns:
        row = l2_df[l2_df["UAID_2"].fillna("").str.strip() == uaid2]
        if not row.empty:
            for col in ("ACBOS_Document", "ACBOS_Doc", "ACBOS", "Document_Name"):
                if col in l2_df.columns:
                    val = str(row.iloc[0].get(col, "")).strip()
                    if val and val.lower() != "nan":
                        fn = val if val.upper().endswith(".ACBOS") else val + ".ACBOS"
                        return fn

    # From PW extract
    if not pw_df.empty:
        desc_col = next((c for c in pw_df.columns if normalize_text(c) == "description"), None)
        fn_col = next((c for c in pw_df.columns if normalize_text(c) == "filename"), None)
        if desc_col and fn_col:
            mask = (
                pw_df[desc_col].fillna("").str.upper().str.contains(uaid2, regex=False)
                & pw_df[fn_col].fillna("").str.upper().str.endswith(".ACBOS")
            )
            subset = pw_df[mask]
            if not subset.empty:
                return str(subset.iloc[0][fn_col]).strip()

    # Fallback
    return f"{uaid2}-ACBOS.ACBOS"


# ---------------------------------------------------------------------------
# Per-UAID generation
# ---------------------------------------------------------------------------

def generate_single_acbos(
    uaid2: str,
    scope3: pd.DataFrame,
    l2_df: pd.DataFrame,
    pw_df: pd.DataFrame,
    output_dir: Path,
    templates: list[AcbosTemplate],
    logger: logging.Logger,
    deliverable_file: str = "",
) -> Path | None:
    """Generate a single ACBOS file for a UAID_2."""
    # Find UAID_2 column in scope3
    _uaid2_norms = {normalize_text(x) for x in ("UAID_2", "uaid_2", "uaid2", "ParentUaid", "Parent_UAID")}
    uaid2_col = next(
        (c for c in scope3.columns if normalize_text(c) in _uaid2_norms), None
    )
    if not uaid2_col:
        logger.warning("  Cannot find UAID_2 column in AssetsScope3 — skipping %s", uaid2)
        return None

    rows = scope3[scope3[uaid2_col].fillna("").str.strip() == uaid2].copy()
    if rows.empty:
        logger.warning("  No AssetsScope3 rows for UAID_2=%s — skipping", uaid2)
        return None

    logger.info("  %s: %d scope3 rows", uaid2, len(rows))

    # Resolve filename
    # Use deliverable file from PW extract if available, otherwise resolve from mapping
    if deliverable_file:
        # Use the deliverable name directly as the output filename
        filename = Path(deliverable_file).name
    else:
        filename = resolve_output_filename(uaid2, l2_df, pw_df, logger)
    
    filename = sanitize_filename(filename)
    if not filename.upper().endswith(".ACBOS"):
        filename += ".ACBOS"

    template = choose_template_for_rows(rows, templates, filename, logger)
    xml_str = build_data_xml(rows, template)

    uaid_dir = output_dir / sanitize_filename(uaid2)
    uaid_dir.mkdir(parents=True, exist_ok=True)
    output_path = get_available_path(uaid_dir / filename)

    # Build ZIP package
    with zipfile.ZipFile(str(output_path), "w", compression=zipfile.ZIP_DEFLATED) as zf:
        if template:
            for archive_name, bin_path in template.bins:
                zf.write(str(bin_path), arcname=archive_name)
        zf.writestr("Data.XML", xml_str)

    logger.info("  ACBOS written: %s", output_path.name)
    return output_path


# ---------------------------------------------------------------------------
# Batch generation
# ---------------------------------------------------------------------------

def generate_acbos_batch(
    workspace: Path,
    cfg: dict,
    targets: list[dict],
    output_dir: Path,
    sources: dict,
    logger: logging.Logger,
) -> dict:
    """Generate ACBOS files for a batch of targets."""
    scope3 = sources["scope3_df"]
    l2_df = sources.get("l2_df", pd.DataFrame())
    pw_df = sources.get("pw_df", pd.DataFrame())

    if scope3.empty:
        raise RuntimeError("AssetsScope3 data not found. Run data cache step first.")

    # Find and extract templates
    tpl_paths = find_templates(workspace, cfg, logger)
    templates: list[AcbosTemplate] = []

    acbos_dir = output_dir / "ACBOS"
    acbos_dir.mkdir(parents=True, exist_ok=True)

    generated, errors = [], []

    with tempfile.TemporaryDirectory() as tmpdir:
        for tpl_path in tpl_paths:
            templates.append(extract_template_data(tpl_path, Path(tmpdir), logger))

        for target in targets:
            uaid2 = target["uaid"]
            deliverable_file = target.get("file", "")
            try:
                out = generate_single_acbos(
                    uaid2, scope3, l2_df, pw_df, acbos_dir, templates, logger,
                    deliverable_file,
                )
                if out:
                    generated.append({"uaid": uaid2, "file": str(out)})
            except Exception as exc:
                logger.error("ACBOS generation failed for %s: %s", uaid2, exc, exc_info=True)
                errors.append({"uaid": uaid2, "error": str(exc)})

    return {"generated": generated, "errors": errors}


# ---------------------------------------------------------------------------
# Standalone entry point
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description="Generate ACBOS files")
    parser.add_argument("--workspace", default=None)
    parser.add_argument("--target-uaid2", nargs="+", required=True)
    args = parser.parse_args()

    workspace = resolve_workspace(args.workspace)
    cfg = load_config(workspace)
    logger = setup_logger(workspace, "acbos_generator", cfg.get("log_level", "INFO"))

    logger.info("=== ACBOS Generator ===")

    from data_loader.local_loader import load_all_sources
    sources = load_all_sources(workspace, cfg, logger)

    uaids = []
    for v in args.target_uaid2:
        uaids.extend(u.strip() for u in v.split(",") if u.strip())

    targets = [{"uaid": u} for u in uaids]
    output_dir = timestamped_dir(workspace, "Output")

    result = generate_acbos_batch(workspace, cfg, targets, output_dir, sources, logger)
    write_json(workspace / "Output" / "acbos_result.json", result)
    logger.info("Generated: %d, Errors: %d", len(result["generated"]), len(result["errors"]))


if __name__ == "__main__":
    main()
