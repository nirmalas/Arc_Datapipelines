"""
mpdt_generator/acbos_generator.py â€” Generate .ACBOS files for target UAID_2 assets.

An .ACBOS file is a ZIP archive containing:
  Setup.BIN   â€” copied verbatim from template
  Key.BIN     â€” copied verbatim from template
  Data.XML    â€” generated from AssetsScope3 rows for the target UAID_2

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


# Columns NOT written as <Attribute> elements â€” structural only
_STRUCTURAL_COLS = {
    "id", "objectid", "ClassificationItemId", "ClsfItemId", "Alignment", "AlgnmPln",
    "Deleted", "ObjectState", "WsGuid", "AssociatedModelFile",
    "source_file", "data_source", "_source", "_source_sheet",
}
_STRUCTURAL_NORMS = {normalize_text(c) for c in _STRUCTURAL_COLS}

# ASC no longer requires Com_AssetRef.  Omit it from every generated ACBOS even
# if it appears in older templates, LoDM, or L3 DB extracts.
_OMITTED_ATTRIBUTE_CODES = {"comassetref", "assetref"}
_REQUIRED_EVEN_WHEN_BLANK_ATTRIBUTE_CODES = {"omasststs"}

# ACBOS must always include these core asset/location/quantity attributes when
# values exist in the L3/Scope3 row.  The sample ACBOS file is only a package
# template; it must not be treated as the complete list of allowed attributes.
_ACBOS_PERMANENT_ATTRIBUTE_NAMES = [
    "UAID_1",
    "UAID_2",
    "UAID_3",
    "HS2_Class",
    "UniClassID",
    "endChainage",
    "Mtrl",
    "startChainage",
    "osgbEasting",
    "osgbNorthing",
    "snakeGridEasting",
    "snakeGridNorthing",
    "NetVolume",
    # Required ACBOS commercial attributes.  They are present in the L3 DB as
    # Com_Dsgnr / Com_Cntrctr and may be listed in LoDM/template as
    # Com:Dsgnr / Com:Cntrctr.  Keep them as core ACBOS fields when values
    # exist, while flexible lookup below handles underscore/colon variants.
    "Com_Dsgnr",
    "Com_Cntrctr",
    "O&M_AsstSts",
]

_RE_ATT_SEP = re.compile(r'[\s_\-\.:&]+')
_ZERO_GUID = '00000000-0000-0000-0000-000000000000'
_XML_ILLEGAL_CHARS_RE = re.compile(
    '['
    '\x00-\x08'
    '\x0B\x0C'
    '\x0E-\x1F'
    '\uD800-\uDFFF'
    '\uFFFE\uFFFF'
    ']'
)



def _clean_xml_attr_text(value: Any, default: str = '') -> str:
    """Return text safe for XML attribute values and .NET ACBOS reader."""
    if value is None:
        return default
    text = str(value)
    if not text.strip():
        return default
    return _XML_ILLEGAL_CHARS_RE.sub('', text).strip()


def _valid_definition_id(value: Any) -> str:
    """DefinitionId must not be blank; ACBOS/.NET readers reject ''."""
    text = _clean_xml_attr_text(value, '')
    return text if text else '0'


def _valid_definition_guid(value: Any) -> str:
    """DefinitionWsGuid must not be blank; use zero GUID when template has no metadata."""
    text = _clean_xml_attr_text(value, '')
    return text if text else _ZERO_GUID


def _norm_att_code(value: Any) -> str:
    """Normalise attribute codes across LoDM/DB/template variants.

    Examples:
      O&M_AsstSts == O&M:AsstSts
      Com_Dsgnr   == Com:Dsgnr
    """
    return _RE_ATT_SEP.sub('', str(value or '')).lower()


def _is_omitted_attribute(value: Any) -> bool:
    code = _norm_att_code(value)
    return code in _OMITTED_ATTRIBUTE_CODES or code.endswith('assetref')


def _drop_omitted_attribute_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    drop_cols = [c for c in df.columns if _is_omitted_attribute(c)]
    return df.drop(columns=drop_cols, errors='ignore') if drop_cols else df


def _value_from_row_by_attr_code(row: pd.Series, attr_name: str) -> Any:
    target = _norm_att_code(attr_name)
    for col, val in row.items():
        if _norm_att_code(col) == target:
            return val
    return None


def _lodm_attr_names_by_class(lodm_df: pd.DataFrame | None) -> dict[str, list[str]]:
    """Return LoDM AttTypeName values keyed by normalised ClassCode.

    Attribute codes are preserved in LoDM order for deterministic ACBOS XML, but
    matching uses _norm_att_code so colon/underscore differences do not make
    attributes appear for the wrong classes.
    """
    out: dict[str, list[str]] = {}
    if lodm_df is None or lodm_df.empty:
        return out
    class_col_keys = {'classcode', 'assethierarchycategory', 'assethierachycategory', 'hs2class'}
    attr_col_keys = {'atttypename', 'attributecode', 'attributename', 'attribute'}
    cc_col = next((c for c in lodm_df.columns if _norm_att_code(c) in class_col_keys), None)
    att_col = next((c for c in lodm_df.columns if _norm_att_code(c) in attr_col_keys), None)
    if not cc_col or not att_col:
        return out
    seen: dict[str, set[str]] = {}
    for _, r in lodm_df.iterrows():
        code = str(r.get(cc_col, '')).strip()
        att = str(r.get(att_col, '')).strip()
        if not code or not att or _is_omitted_attribute(att):
            continue
        keys = {normalize_text(code), _norm_att_code(code)}
        keys = {k for k in keys if k}
        att_key = _norm_att_code(att)
        if not att_key:
            continue
        for key in keys:
            if att_key not in seen.setdefault(key, set()):
                out.setdefault(key, []).append(att)
                seen[key].add(att_key)
    return out


def _class_code_from_row(row: pd.Series) -> str:
    for candidate in ('HS2_Class', 'AssetHierarchyCategory', 'ClassCode', 'ClassificationItemId', 'ClsfItemId'):
        val = _value_from_row_by_normalized_name(row, normalize_text(candidate))
        if not _is_empty(val):
            return str(val).strip()
    return ''


def _upper_key(value: Any) -> str:
    if value is None or str(value).strip().lower() in {"", "nan", "nat"}:
        return ""
    return str(value).strip().upper()


def _group_df_by_upper_key(df: pd.DataFrame, key_col: str | None) -> dict[str, pd.DataFrame]:
    """Group rows by a normalised UAID key once per batch."""
    if df is None or df.empty or not key_col or key_col not in df.columns:
        return {}
    work = df.copy()
    work["__acbos_lookup_key__"] = work[key_col].map(_upper_key)
    work = work[work["__acbos_lookup_key__"] != ""]
    return {str(k): g.drop(columns=["__acbos_lookup_key__"]) for k, g in work.groupby("__acbos_lookup_key__", sort=False)}


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
    allowed_attr_norms: set[str]


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
      * Attribute metadata from the sample is reused where available.  The
        sample is not treated as the complete attribute allow-list; generated
        ACBOS attributes are driven by permanent asset fields plus LoDM for the
        row classcode.
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
    return AcbosTemplate(template_path, bins, metadata, attr_order, classes, set(metadata.keys()))


def find_templates(workspace: Path, cfg: dict, logger: logging.Logger) -> list[Path]:
    """Find ACBOS template files, sorted deterministically."""
    configured_file = cfg.get("paths", {}).get("acbos_template_file", "")
    if configured_file:
        template_path = (workspace / configured_file).resolve()
        if template_path.exists():
            logger.info("Using configured ACBOS template: %s", template_path)
            return [template_path]
        logger.warning("Configured ACBOS template not found: %s", template_path)

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
    # Literal string "None" is a valid L3 DB value and must be written to ACBOS.
    return s == "" or s.lower() in ("nan", "nat")


def _clean_acbos_output_value(v: Any) -> Any:
    """Return a value ready for ACBOS XML, preserving literal 'None'."""
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except Exception:
        pass
    return v


def _value_from_row_by_normalized_name(row: pd.Series, norm_name: str) -> Any:
    for col, val in row.items():
        if normalize_text(col) == norm_name:
            return val
    return None


def _metadata_for_attribute(attr_metadata: dict[str, dict[str, str]], attr_name: str) -> dict[str, str]:
    """Return template metadata with XML/schema-safe defaults.

    The sample ACBOS template may not contain metadata for every permanent or
    LoDM-driven attribute we now emit.  Writing DefinitionId="" or
    DefinitionWsGuid="" can make the ACBOS client fail with messages like
    "There is an error in XML document (line, position)" because those
    attributes are parsed as numeric/Guid values.  For template-missing
    attributes, use deterministic safe defaults instead of blank strings.
    """
    norm_name = normalize_text(attr_name)
    meta = attr_metadata.get(norm_name)
    if meta is None:
        target = _norm_att_code(attr_name)
        meta = next((m for m in attr_metadata.values() if _norm_att_code(m.get('Name', '')) == target), None)
    if meta is None:
        meta = {"Name": attr_name, "DefinitionId": "0", "DefinitionWsGuid": _ZERO_GUID}
    return {
        "Name": _clean_xml_attr_text(meta.get("Name") or attr_name, str(attr_name)),
        "DefinitionId": _valid_definition_id(meta.get("DefinitionId")),
        "DefinitionWsGuid": _valid_definition_guid(meta.get("DefinitionWsGuid")),
    }




def _remove_omitted_attributes_from_xml(root: ET.Element) -> None:
    """Hard-stop deprecated attributes such as Com_AssetRef in generated XML.

    This is intentionally applied at the XML element level as a final safety net,
    because older templates/LoDM/L3 extracts may contain casing, underscore,
    colon, or space variants that survive earlier DataFrame filtering.
    """
    for obj in list(root.iter("Object")):
        for attr in list(obj.findall("Attribute")):
            if _is_omitted_attribute(attr.get("Name", "")):
                obj.remove(attr)


def _append_attribute_once(attribute_names: list[str], seen_attr_codes: set[str], attr_name: Any) -> None:
    name = str(attr_name or '').strip()
    code = _norm_att_code(name)
    if not name or not code or code in seen_attr_codes or _is_omitted_attribute(name):
        return
    attribute_names.append(name)
    seen_attr_codes.add(code)


def build_data_xml(rows: pd.DataFrame, template: AcbosTemplate | None, lodm_attr_by_class: dict[str, list[str]] | None = None) -> str:
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

        # ClassificationItemId â€” check ClassificationItemId, then ClsfItemId, then HS2_Class
        cls_id = ""
        for candidate in ("ClassificationItemId", "ClsfItemId", "HS2_Class"):
            val = _value_from_row_by_normalized_name(row, normalize_text(candidate))
            if not _is_empty(val):
                cls_id = str(val).strip()
                break
        obj_elem.set("ClassificationItemId", cls_id)

        # Alignment â€” check Alignment, then AlgnmPln
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

        # Per-row ACBOS attributes are constructed like MPDT AU+:
        #   1) permanent asset/location/quantity attributes are always eligible
        #      when a value exists in the row;
        #   2) remaining attributes come from LoDM for this row's classcode.
        # The sample ACBOS template only supplies package structure and optional
        # DefinitionId/DefinitionWsGuid metadata.  It is not an allow-list.
        class_code = _class_code_from_row(row)
        lodm_attrs = (lodm_attr_by_class or {}).get(normalize_text(class_code), []) or (lodm_attr_by_class or {}).get(_norm_att_code(class_code), [])

        attribute_names: list[str] = []
        seen_attr_codes: set[str] = set()
        for attr_name in _ACBOS_PERMANENT_ATTRIBUTE_NAMES:
            _append_attribute_once(attribute_names, seen_attr_codes, attr_name)
        for attr_name in lodm_attrs:
            _append_attribute_once(attribute_names, seen_attr_codes, attr_name)

        for attr_name in attribute_names:
            norm_name = normalize_text(attr_name)
            attr_code = _norm_att_code(attr_name)
            required_even_when_blank = attr_code in _REQUIRED_EVEN_WHEN_BLANK_ATTRIBUTE_CODES
            if norm_name in _STRUCTURAL_NORMS or _is_omitted_attribute(attr_name):
                continue
            # First try exact normalised name, then flexible attribute-code matching
            # so DB columns such as O&M_AsstSts populate LoDM/template
            # O&M:AsstSts and do not leak onto classes where LoDM disallows it.
            val = _value_from_row_by_normalized_name(row, norm_name)
            val = _clean_acbos_output_value(val)
            if _is_empty(val):
                val = _clean_acbos_output_value(_value_from_row_by_attr_code(row, attr_name))

            # Literal string 'None' is a valid value for Com:RfrncNmbr and any
            # other L3 attribute.  Only true Python/Excel missing values are skipped.
            if _is_empty(val) and not required_even_when_blank:
                continue
            str_val = "" if _is_empty(val) else _clean_xml_attr_text(val)
            if not str_val and not required_even_when_blank:
                continue

            meta = _metadata_for_attribute(attr_metadata, attr_name)
            attr_elem = ET.SubElement(obj_elem, "Attribute")
            attr_elem.set("DefinitionId", meta["DefinitionId"])
            attr_elem.set("Name", meta["Name"])
            attr_elem.set("Value", str_val)
            attr_elem.set("DefinitionWsGuid", meta["DefinitionWsGuid"])
            ET.SubElement(attr_elem, "ObjectState").text = "Unchanged"
            ET.SubElement(attr_elem, "WsGuid").text = new_guid()

    _remove_omitted_attributes_from_xml(root)

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
    rows_by_uaid2: dict[str, pd.DataFrame] | None = None,
    lodm_attr_by_class: dict[str, list[str]] | None = None,
) -> Path | None:
    """Generate a single ACBOS file for a UAID_2."""
    rows = pd.DataFrame()
    if rows_by_uaid2 is not None:
        rows = rows_by_uaid2.get(_upper_key(uaid2), pd.DataFrame()).copy()
    else:
        # Find UAID_2 column in scope3
        _uaid2_norms = {normalize_text(x) for x in ("UAID_2", "uaid_2", "uaid2", "ParentUaid", "Parent_UAID")}
        uaid2_col = next((c for c in scope3.columns if normalize_text(c) in _uaid2_norms), None)
        if not uaid2_col:
            logger.warning("  Cannot find UAID_2 column in AssetsScope3 â€” skipping %s", uaid2)
            return None
        rows = scope3[scope3[uaid2_col].map(_upper_key) == _upper_key(uaid2)].copy()
    if rows.empty:
        logger.warning("  No AssetsScope3 rows for UAID_2=%s â€” skipping", uaid2)
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
    xml_str = build_data_xml(rows, template, lodm_attr_by_class=lodm_attr_by_class)

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
    scope3 = _drop_omitted_attribute_columns(sources["scope3_df"])
    l2_df = sources.get("l2_df", pd.DataFrame())
    pw_df = sources.get("pw_df", pd.DataFrame())
    lodm_df = sources.get("lodm_df", pd.DataFrame())
    lodm_attr_by_class = _lodm_attr_names_by_class(lodm_df)
    logger.info("Indexed ACBOS LoDM class attributes=%d", len(lodm_attr_by_class))

    if scope3.empty:
        raise RuntimeError("AssetsScope3 data not found. Run data cache step first.")

    # Index Scope3 once instead of filtering the full DataFrame for every target.
    _uaid2_norms = {normalize_text(x) for x in ("UAID_2", "uaid_2", "uaid2", "ParentUaid", "Parent_UAID")}
    uaid2_col = next((c for c in scope3.columns if normalize_text(c) in _uaid2_norms), None)
    rows_by_uaid2 = _group_df_by_upper_key(scope3, uaid2_col)
    logger.info("Indexed ACBOS Scope3 parents=%d", len(rows_by_uaid2))

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
                    deliverable_file, rows_by_uaid2=rows_by_uaid2,
                    lodm_attr_by_class=lodm_attr_by_class,
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
