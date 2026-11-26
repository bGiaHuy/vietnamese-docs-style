"""Independent 3-layer OPC and ECMA-376 WordprocessingML package validator."""

from __future__ import annotations

import os
import posixpath
import re
import subprocess
import zipfile
from collections import Counter
from pathlib import Path
from typing import Any

from lxml import etree

from eval_harness.ooxml_reader import MC, W, parse_xml, read_zip_parts

_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

_VALID_W_TYPES = {"auto", "dxa", "nil", "pct"}
_VALID_LINE_RULES = {"auto", "exact", "atLeast"}
_VALID_JC = {"left", "center", "right", "both", "start", "end", "distribute", "numTab", "highKashida", "lowKashida", "mediumKashida"}
_HEX_COLOR_RE = re.compile(r"^([0-9A-Fa-f]{6}|auto)$")

# Normative ECMA-376 Part 1 (§17) relative child element ordering per container.
_ECMA_SEQUENCES: dict[str, list[str]] = {
    "tblPr": [
        "tblStyle", "tblpPr", "tblOverlap", "bidiVisual",
        "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc",
        "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout",
        "tblCellMar", "tblLook", "tblCaption", "tblDescription", "tblPrChange",
    ],
    "tblGrid": ["gridCol", "tblGridChange"],
    "trPr": [
        "cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter",
        "cantSplit", "trHeight", "tblHeader", "tblCellSpacing", "jc", "hidden",
        "ins", "del", "trPrChange",
    ],
    "tcPr": [
        "cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders",
        "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign",
        "hideMark", "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange",
    ],
    "tblBorders": ["top", "left", "start", "bottom", "right", "end", "insideH", "insideV", "tl2br", "tr2bl"],
    "tcBorders": ["top", "left", "start", "bottom", "right", "end", "insideH", "insideV", "tl2br", "tr2bl"],
    "tblCellMar": ["top", "left", "start", "bottom", "right", "end"],
    "tcMar": ["top", "left", "start", "bottom", "right", "end"],
    "pPr": [
        "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
        "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
        "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
        "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
        "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
        "suppressOverlap", "jc", "textDirection", "textAlignment",
        "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
        "sectPr", "pPrChange",
    ],
    "rPr": [
        "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps",
        "strike", "dstrike", "outline", "shadow", "emboss", "imprint",
        "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
        "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
        "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang",
        "eastAsianLayout", "specVanish", "oMath", "rPrChange",
    ],
    "sectPr": [
        "headerReference", "footerReference", "footnotePr", "endnotePr",
        "type", "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType",
        "pgNumType", "cols", "formProt", "vAlign", "noEndnote", "titlePg",
        "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings",
        "sectPrChange",
    ],
    "style": [
        "name", "aliases", "basedOn", "next", "link", "autoRedefine",
        "hidden", "uiPriority", "semiHidden", "unhideWhenUsed", "qFormat",
        "locked", "personal", "personalCompose", "personalReply", "rsid",
        "pPr", "rPr", "tblPr", "trPr", "tcPr", "tblCellSpacing", "tblStylePr",
    ],
    "docDefaults": ["rPrDefault", "pPrDefault"],
}

_FIRST_PR_RULES: dict[str, tuple[str, set[str]]] = {
    "tbl": ("tblPr", {"tblGrid", "tr"}),
    "tr": ("trPr", {"tc"}),
    "tc": ("tcPr", {"p", "tbl"}),
    "p": ("pPr", {"r", "hyperlink", "fldSimple", "ins", "del"}),
    "r": ("rPr", {"t", "br", "tab", "drawing", "pict", "instrText", "fldChar"}),
}


def validate_opc_container(docx_path: str | Path) -> list[dict[str, str]]:
    """Validate ZIP container health, [Content_Types].xml, XML well-formedness, and .rels targets."""
    errors: list[dict[str, str]] = []
    try:
        meta, parts = read_zip_parts(docx_path)
    except (zipfile.BadZipFile, OSError) as exc:
        return [{"layer": "opc", "part": "", "rule": "zip_open", "message": str(exc)}]

    if meta["has_dir_entries"]:
        errors.append({"layer": "opc", "part": "", "rule": "zip_dir_entries", "message": "ZIP contains directory entries"})
    if meta["duplicate_names"]:
        errors.append({
            "layer": "opc",
            "part": "",
            "rule": "zip_duplicate_entries",
            "message": f"Duplicate ZIP entries: {meta['duplicate_names']}",
        })

    if "[Content_Types].xml" not in parts:
        errors.append({"layer": "opc", "part": "[Content_Types].xml", "rule": "missing_content_types", "message": "Missing [Content_Types].xml"})
        return errors

    default_exts: set[str] = set()
    overrides: set[str] = set()
    try:
        ct_root = parse_xml(parts["[Content_Types].xml"])
        for el in ct_root:
            if not isinstance(el.tag, str) or etree.QName(el).namespace != _CT_NS:
                continue
            ln = etree.QName(el).localname
            if ln == "Default":
                ext = (el.get("Extension") or "").lower()
                if ext:
                    default_exts.add(ext)
            elif ln == "Override":
                pn = (el.get("PartName") or "").lstrip("/").lower()
                if pn:
                    overrides.add(pn)
    except etree.XMLSyntaxError as exc:
        errors.append({"layer": "opc", "part": "[Content_Types].xml", "rule": "xml_parse", "message": str(exc)})
        return errors

    lower_names = {n.lower(): n for n in parts}
    for name in sorted(parts):
        if name == "[Content_Types].xml":
            continue
        basename = posixpath.basename(name)
        ext = basename.rsplit(".", 1)[1].lower() if "." in basename else ""
        if name.lower() not in overrides and ext not in default_exts:
            errors.append({
                "layer": "opc",
                "part": name,
                "rule": "content_type_coverage",
                "message": f"Part {name} missing from [Content_Types].xml",
            })
        if name.endswith((".xml", ".rels")):
            try:
                parse_xml(parts[name])
            except etree.XMLSyntaxError as exc:
                errors.append({"layer": "opc", "part": name, "rule": "xml_parse", "message": str(exc)})

        if name.endswith(".rels"):
            try:
                rel_root = parse_xml(parts[name])
            except etree.XMLSyntaxError:
                continue
            # Determine source directory per OPC part relationship naming rules.
            if name == "_rels/.rels":
                base_dir = ""
            else:
                rels_dir = posixpath.dirname(name)
                base_dir = posixpath.dirname(rels_dir)
            for rel in rel_root.findall(f"{{{_REL_NS}}}Relationship"):
                tmode = (rel.get("TargetMode") or "Internal").strip()
                if tmode.lower() == "external":
                    continue
                target = (rel.get("Target") or "").strip()
                if not target or target.upper() == "NULL":
                    errors.append({
                        "layer": "opc",
                        "part": name,
                        "rule": "rels_target_invalid",
                        "message": f"Invalid internal relationship target {target!r}",
                    })
                    continue
                resolved = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(base_dir, target))
                if resolved not in parts and resolved.lower() not in lower_names:
                    errors.append({
                        "layer": "opc",
                        "part": name,
                        "rule": "rels_dangling_target",
                        "message": f"Relationship target missing from ZIP: {resolved}",
                    })

    return errors


def _check_wml_tree(root: etree._Element, part_name: str) -> list[dict[str, str]]:
    errors: list[dict[str, str]] = []
    for el in root.iter():
        if not isinstance(el.tag, str):
            continue
        q = etree.QName(el)
        if q.namespace == MC and q.localname == "AlternateContent":
            choices = el.findall(f"{{{MC}}}Choice")
            fallbacks = el.findall(f"{{{MC}}}Fallback")
            if not choices:
                errors.append({"layer": "ecma376", "part": part_name, "rule": "mc_choice_missing", "message": "mc:AlternateContent lacks mc:Choice"})
            for ch in choices:
                if not (ch.get("Requires") or "").strip():
                    errors.append({"layer": "ecma376", "part": part_name, "rule": "mc_requires_missing", "message": "mc:Choice lacks @Requires"})
            if len(fallbacks) > 1:
                errors.append({"layer": "ecma376", "part": part_name, "rule": "mc_multiple_fallback", "message": "mc:AlternateContent has multiple mc:Fallback elements"})
            continue

        if q.namespace != W:
            continue
        ln = q.localname

        if ln in _FIRST_PR_RULES:
            pr_name, after_names = _FIRST_PR_RULES[ln]
            saw_content = False
            for ch in el:
                if not isinstance(ch.tag, str):
                    continue
                cq = etree.QName(ch)
                if cq.namespace != W:
                    continue
                if cq.localname in after_names:
                    saw_content = True
                elif cq.localname == pr_name and saw_content:
                    errors.append({
                        "layer": "ecma376",
                        "part": part_name,
                        "rule": "pr_after_content",
                        "message": f"{pr_name} appears after content child in {ln}",
                    })

        if ln in _ECMA_SEQUENCES:
            seq = _ECMA_SEQUENCES[ln]
            rank = {name: idx for idx, name in enumerate(seq)}
            # Allow headerReference and footerReference in any relative order per ECMA-376.
            if ln == "sectPr":
                rank["footerReference"] = rank["headerReference"]
            max_idx = -1
            max_name = ""
            for ch in el:
                if not isinstance(ch.tag, str):
                    continue
                cq = etree.QName(ch)
                if cq.namespace != W or cq.localname not in rank:
                    continue
                pos = rank[cq.localname]
                if pos < max_idx:
                    errors.append({
                        "layer": "ecma376",
                        "part": part_name,
                        "rule": "child_order",
                        "message": f"{cq.localname} out of order after {max_name} inside {ln}",
                    })
                elif pos > max_idx:
                    max_idx = pos
                    max_name = cq.localname

        # Check critical attribute enums & types.
        if ln in ("tblW", "tcW"):
            wtype = el.get(f"{{{W}}}type")
            if wtype is not None and wtype not in _VALID_W_TYPES:
                errors.append({"layer": "ecma376", "part": part_name, "rule": "invalid_w_type", "message": f"Invalid {ln}/@w:type={wtype!r}"})
        elif ln == "spacing":
            lr = el.get(f"{{{W}}}lineRule")
            if lr is not None and lr not in _VALID_LINE_RULES:
                errors.append({"layer": "ecma376", "part": part_name, "rule": "invalid_line_rule", "message": f"Invalid spacing/@w:lineRule={lr!r}"})
        elif ln == "jc":
            jv = el.get(f"{{{W}}}val")
            if jv is not None and jv not in _VALID_JC:
                errors.append({"layer": "ecma376", "part": part_name, "rule": "invalid_jc", "message": f"Invalid jc/@w:val={jv!r}"})
        elif ln == "color":
            cv = el.get(f"{{{W}}}val")
            if cv is not None and not _HEX_COLOR_RE.match(cv):
                errors.append({"layer": "ecma376", "part": part_name, "rule": "invalid_color", "message": f"Invalid color/@w:val={cv!r}"})

    return errors


def validate_ecma376_structure(docx_path: str | Path) -> list[dict[str, str]]:
    """Validate WordprocessingML element ordering, enums, and mc:AlternateContent structure."""
    try:
        _, parts = read_zip_parts(docx_path)
    except (zipfile.BadZipFile, OSError):
        return []

    errors: list[dict[str, str]] = []
    target_parts = [
        name
        for name in sorted(parts)
        if name == "word/document.xml"
        or name == "word/styles.xml"
        or name in ("word/footnotes.xml", "word/endnotes.xml", "word/comments.xml")
        or (name.startswith(("word/header", "word/footer")) and name.endswith(".xml"))
    ]
    for pname in target_parts:
        try:
            root = parse_xml(parts[pname])
        except etree.XMLSyntaxError:
            continue
        errors.extend(_check_wml_tree(root, pname))
    return errors


def validate_package(
    docx_path: str | Path,
    *,
    baseline_docx: str | Path | None = None,
    ecma376_xsd_path: str | Path | None = None,
    openxml_validator_cmd: str | None = None,
) -> dict[str, Any]:
    """Run independent multi-layer validation and return backend provenance + error deltas."""
    backends = ["standalone_opc", "standalone_ecma376_subset"]
    opc_errors = validate_opc_container(docx_path)
    wml_errors = validate_ecma376_structure(docx_path)
    ext_errors: list[dict[str, str]] = []

    xsd_candidate = Path(ecma376_xsd_path) if ecma376_xsd_path else (Path(__file__).resolve().parent / "schemas" / "ecma376" / "wml.xsd")
    if xsd_candidate.is_file():
        backends.append("ecma376_xsd")
        try:
            schema = etree.XMLSchema(etree.parse(str(xsd_candidate)))
            _, parts = read_zip_parts(docx_path)
            if "word/document.xml" in parts:
                doc_root = parse_xml(parts["word/document.xml"])
                if not schema.validate(doc_root):
                    for err in schema.error_log:
                        ext_errors.append({
                            "layer": "ecma376_xsd",
                            "part": "word/document.xml",
                            "rule": "xsd_validation",
                            "message": str(err.message),
                        })
        except Exception as exc:
            ext_errors.append({"layer": "ecma376_xsd", "part": "word/document.xml", "rule": "xsd_error", "message": str(exc)})

    sdk_cmd = openxml_validator_cmd or os.environ.get("OPENXML_VALIDATOR_CMD")
    if sdk_cmd:
        backends.append("openxml_sdk")
        try:
            proc = subprocess.run(
                [sdk_cmd, str(docx_path)],
                capture_output=True,
                text=True,
                timeout=60,
                check=False,
            )
            if proc.returncode != 0:
                ext_errors.append({
                    "layer": "openxml_sdk",
                    "part": "",
                    "rule": "openxml_sdk_nonzero",
                    "message": (proc.stderr or proc.stdout or f"exit={proc.returncode}").strip()[:500],
                })
        except Exception as exc:
            ext_errors.append({"layer": "openxml_sdk", "part": "", "rule": "openxml_sdk_exec", "message": str(exc)})

    all_errors = opc_errors + wml_errors + ext_errors
    new_errors = list(all_errors)

    if baseline_docx is not None and Path(baseline_docx).is_file():
        base_all = validate_opc_container(baseline_docx) + validate_ecma376_structure(baseline_docx)
        base_counts = Counter((e["layer"], e["part"], e["rule"], e["message"]) for e in base_all)
        new_errors = []
        seen_counts: Counter[tuple[str, str, str, str]] = Counter()
        for err in all_errors:
            key = (err["layer"], err["part"], err["rule"], err["message"])
            seen_counts[key] += 1
            if seen_counts[key] > base_counts.get(key, 0):
                new_errors.append(err)

    return {
        "valid": len(new_errors) == 0,
        "strict_valid": len(all_errors) == 0,
        "validator_backends": backends,
        "all_error_count": len(all_errors),
        "new_error_count": len(new_errors),
        "all_errors": all_errors,
        "new_errors": new_errors,
    }
