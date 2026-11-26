"""Completely decoupled independent grader and dev-v2.0 cross-checker."""

from __future__ import annotations

import argparse
import hashlib
import json
import tempfile
from pathlib import Path
from typing import Any

import pymupdf
from lxml import etree

from eval_harness.diff_ntpr import audit_namespaces, check_non_target_preservation
from eval_harness.ooxml_reader import (
    W,
    effective_run_format,
    find_style_record,
    inspect_docx,
    nfc,
    parse_xml,
    read_zip_parts,
    text_of_element,
)
from eval_harness.render_inspector import (
    detect_orphan_signature_pdf,
    detect_pagination_ripple,
    inspect_pdf_layout,
    render_to_pdf,
)
from eval_harness.schema_validator import validate_package


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def check_compliance_profile(docx_info: dict[str, Any], profile: str) -> dict[str, Any]:
    """Check Vietnamese document convention compliance independently of scripts/validate_docx.py."""
    issues: list[str] = []
    sections = docx_info.get("sections", [])
    if not sections:
        return {"status": "fail", "issues": ["missing sections"]}

    sec = sections[0]
    pw, ph = sec["page_width_cm"], sec["page_height_cm"]
    mar = sec["margins_cm"]
    # Allow portrait A4 (21.0 x 29.7) or landscape A4 (29.7 x 21.0).
    a4_ok = (abs(pw - 21.0) <= 0.25 and abs(ph - 29.7) <= 0.25) or (abs(pw - 29.7) <= 0.25 and abs(ph - 21.0) <= 0.25)
    if not a4_ok:
        issues.append(f"page size {pw}x{ph} cm is not A4")

    if profile == "administrative":
        if not (1.95 <= mar["top"] <= 2.55 and 1.95 <= mar["bottom"] <= 2.55):
            issues.append(f"ND30 top/bottom margin out of [2.0, 2.5]cm: {mar['top']}/{mar['bottom']}")
        if not (2.95 <= mar["left"] <= 3.55 and 1.45 <= mar["right"] <= 2.05):
            issues.append(f"ND30 left/right margin out of [3.0-3.5 / 1.5-2.0]cm: {mar['left']}/{mar['right']}")
        for tbl in docx_info.get("tables", []):
            if tbl.get("is_layout_table"):
                continue
            bad_sz = [s for s in tbl.get("data_font_sizes_pt", []) if s < 11.9]
            if bad_sz:
                issues.append(f"table {tbl['ordinal']} data font < 12pt: {bad_sz[0]}pt")
    elif profile == "party":
        if (
            abs(mar["top"] - 2.0) > 0.08
            or abs(mar["bottom"] - 2.0) > 0.08
            or abs(mar["left"] - 3.0) > 0.08
            or abs(mar["right"] - 1.5) > 0.08
        ):
            issues.append(f"Party HD05 margins expected 2.0/2.0/3.0/1.5cm, got {mar}")
        for p in docx_info.get("paragraphs", []):
            txt = p["text_nfc"].strip()
            if len(txt) < 60:
                continue
            fmt = p["effective_format"]
            if fmt.get("align") != "both":
                issues.append(f"Party body para {p['ordinal']} align={fmt.get('align')!r} != 'both'")
                break
            if fmt.get("line_rule") != "exact" or not (18.0 <= (fmt.get("line_pt") or 0.0) <= 22.0):
                issues.append(f"Party body para {p['ordinal']} line spacing not exact 18-22pt")
                break
    elif profile == "legal_normative":
        if abs(mar["top"] - 2.0) > 0.08 or abs(mar["bottom"] - 2.0) > 0.08:
            issues.append(f"ND187 top/bottom margins must be 2.0cm, got {mar['top']}/{mar['bottom']}")
        if not (2.95 <= mar["left"] <= 3.55 and 1.45 <= mar["right"] <= 2.05):
            issues.append(f"ND187 left/right margins out of range: {mar['left']}/{mar['right']}")
    elif profile == "academic":
        if any(mar[k] < 1.45 for k in ("top", "bottom", "left", "right")):
            issues.append(f"academic margin < 1.5cm: {mar}")

    return {"status": "pass" if not issues else "fail", "issues": issues}


def _check_table_assertions(
    tbl_before: dict[str, Any],
    tbl_after: dict[str, Any],
    grading: dict[str, Any],
) -> tuple[bool, bool]:
    """Check target table OOXML width/cols/alignment and styling assertions."""
    if "expected_width_min_cm" in grading and "expected_width_max_cm" in grading:
        w_ok = grading["expected_width_min_cm"] <= tbl_after["total_width_cm"] <= grading["expected_width_max_cm"]
    elif "expected_total_width_cm" in grading:
        tol = grading.get("width_tolerance_cm", 0.35)
        w_ok = abs(tbl_after["total_width_cm"] - grading["expected_total_width_cm"]) <= tol
    elif "expected_width_max_drift_cm" in grading:
        w_ok = abs(tbl_after["total_width_cm"] - tbl_before["total_width_cm"]) <= grading["expected_width_max_drift_cm"]
    else:
        w_ok = True

    cols_ok = True
    if "expected_col_widths_cm" in grading:
        ctol = grading.get("col_tolerance_cm", 0.15)
        exp_cols = grading["expected_col_widths_cm"]
        cols_ok = (
            len(tbl_after["grid_widths_cm"]) == len(exp_cols)
            and all(abs(a - b) <= ctol for a, b in zip(tbl_after["grid_widths_cm"], exp_cols))
        )
    if "expected_col_widths_pct" in grading:
        ptol = grading.get("pct_tolerance", 2.5)
        exp_pcts = grading["expected_col_widths_pct"]
        cols_ok = cols_ok and (
            len(tbl_after["grid_widths_pct"]) == len(exp_pcts)
            and all(abs(a - b) <= ptol for a, b in zip(tbl_after["grid_widths_pct"], exp_pcts))
        )
    if grading.get("check_grid_span_tcw"):
        cols_ok = cols_ok and bool(tbl_after.get("grid_span_tcw_consistent"))

    align_ok = (
        True
        if "expected_alignment" not in grading
        else tbl_after.get("alignment") == grading["expected_alignment"]
    )
    target_width_ok = bool(w_ok and cols_ok and align_ok)

    style_ok = True
    if "expected_cell_margins_cm" in grading:
        mtol = grading.get("margin_tolerance_cm", 0.03)
        got_mar = tbl_after.get("cell_margins_cm") or {}
        for side, want_cm in grading["expected_cell_margins_cm"].items():
            if side not in got_mar or abs(got_mar[side] - want_cm) > mtol:
                style_ok = False
                break
    if grading.get("require_no_cell_margin_overrides") and tbl_after.get("has_cell_margin_overrides"):
        style_ok = False

    if "expected_header_font_pt" in grading:
        want_hp = grading["expected_header_font_pt"]
        hdr_sizes = tbl_after.get("header_font_sizes_pt", [])
        if not hdr_sizes or any(abs(s - want_hp) > 0.1 for s in hdr_sizes):
            style_ok = False
    if grading.get("require_header_bold") and not tbl_after.get("header_all_bold"):
        style_ok = False
    if "expected_data_font_pt" in grading:
        want_dp = grading["expected_data_font_pt"]
        data_sizes = tbl_after.get("data_font_sizes_pt", [])
        if not data_sizes or any(abs(s - want_dp) > 0.1 for s in data_sizes):
            style_ok = False

    if "expected_repeat_header" in grading:
        if bool(tbl_after.get("repeat_header")) is not bool(grading["expected_repeat_header"]):
            style_ok = False
    if "expected_cant_split" in grading:
        if bool(tbl_after.get("cant_split_all_rows")) is not bool(grading["expected_cant_split"]):
            style_ok = False

    if "expected_border_style" in grading or "expected_border_color" in grading:
        borders = tbl_after.get("borders", {})
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            binfo = borders.get(side)
            if not binfo:
                style_ok = False
                break
            if "expected_border_style" in grading and binfo.get("val") != grading["expected_border_style"]:
                style_ok = False
                break
            if "expected_border_color" in grading and (binfo.get("color") or "").upper() != grading["expected_border_color"].upper():
                style_ok = False
                break

    if "expected_header_shading_fill" in grading:
        want_fill = grading["expected_header_shading_fill"].upper()
        hdr_fills = tbl_after.get("header_shading_fills", [])
        data_fills = tbl_after.get("data_shading_fills", [])
        if not hdr_fills or any((f or "").upper() != want_fill for f in hdr_fills):
            style_ok = False
        if any((f or "").upper() == want_fill for f in data_fills):
            style_ok = False

    return target_width_ok, style_ok


def _check_image_assertions(
    img_before: dict[str, Any],
    img_after: dict[str, Any],
    grading: dict[str, Any],
) -> bool:
    if not img_after.get("xfrm_consistent"):
        return False
    w_ok = True
    if "expected_image_width_cm" in grading:
        tol = grading.get("image_width_tolerance_cm", 0.25)
        w_ok = abs(img_after["width_cm"] - grading["expected_image_width_cm"]) <= tol
    ar_ok = True
    ar_b = img_before.get("aspect_ratio")
    ar_a = img_after.get("aspect_ratio")
    if ar_b and ar_a and "max_aspect_ratio_delta_pct" in grading:
        delta_pct = abs(ar_a - ar_b) / ar_b * 100.0
        ar_ok = delta_pct <= grading["max_aspect_ratio_delta_pct"]
    align_ok = True
    if "expected_image_alignment" in grading:
        align_ok = img_after.get("alignment") == grading["expected_image_alignment"]
    return bool(w_ok and ar_ok and align_ok)


def _check_paragraph_assertions(p_after: dict[str, Any], grading: dict[str, Any]) -> bool:
    fmt = p_after["effective_format"]
    if "expected_para_align" in grading:
        want_al = "both" if grading["expected_para_align"] == "justify" else grading["expected_para_align"]
        if fmt.get("align") != want_al:
            return False
    if "expected_first_line_cm_range" in grading:
        lo, hi = grading["expected_first_line_cm_range"]
        fl = fmt.get("first_line_cm")
        if fl is None or not (lo <= fl <= hi):
            return False
    if "expected_hanging_indent_cm_range" in grading:
        lo_h, hi_h = grading["expected_hanging_indent_cm_range"]
        hg = fmt.get("hanging_cm")
        if hg is None or not (lo_h <= hg <= hi_h):
            return False
    if "expected_line_rule" in grading:
        if fmt.get("line_rule") != grading["expected_line_rule"]:
            return False
    if "expected_line_pt_range" in grading:
        lo_pt, hi_pt = grading["expected_line_pt_range"]
        lp = fmt.get("line_pt")
        if lp is None or not (lo_pt <= lp <= hi_pt):
            return False
    if "min_para_spacing_pt" in grading:
        sa = fmt.get("space_after_pt") or 0.0
        if sa < grading["min_para_spacing_pt"] - 0.05:
            return False
    if "expected_keep_with_next" in grading:
        if bool(fmt.get("keep_with_next")) is not bool(grading["expected_keep_with_next"]):
            return False
    return True


def _check_style_assertions(after_info: dict[str, Any], grading: dict[str, Any]) -> bool:
    st = find_style_record(grading.get("target_style_name"), after_info["styles_ctx"])
    if st is None:
        return False
    r_pr = st.get("rPr")
    p_pr = st.get("pPr")

    if "expected_style_font" in grading:
        rf = r_pr.find(f"{{{W}}}rFonts") if r_pr is not None else None
        want_f = grading["expected_style_font"]
        if rf is None or rf.get(f"{{{W}}}ascii") != want_f or rf.get(f"{{{W}}}hAnsi") != want_f:
            return False
    if "expected_style_size_pt" in grading:
        sz = r_pr.find(f"{{{W}}}sz") if r_pr is not None else None
        val = sz.get(f"{{{W}}}val") if sz is not None else None
        if val is None or abs(int(val) / 2.0 - grading["expected_style_size_pt"]) > 0.05:
            return False
    if "expected_style_color" in grading:
        col = r_pr.find(f"{{{W}}}color") if r_pr is not None else None
        if col is None or (col.get(f"{{{W}}}val") or "").upper() != grading["expected_style_color"].upper():
            return False

    sp = p_pr.find(f"{{{W}}}spacing") if p_pr is not None else None
    if "expected_style_line_rule" in grading:
        if sp is None or sp.get(f"{{{W}}}lineRule") != grading["expected_style_line_rule"]:
            return False
    if "expected_style_line_pt" in grading:
        lval = sp.get(f"{{{W}}}line") if sp is not None else None
        if lval is None or abs(int(lval) / 20.0 - grading["expected_style_line_pt"]) > 0.1:
            return False
    if "expected_style_space_after_pt" in grading:
        aval = sp.get(f"{{{W}}}after") if sp is not None else None
        if aval is None or abs(int(aval) / 20.0 - grading["expected_style_space_after_pt"]) > 0.1:
            return False
    return True


def _check_section_assertions(
    before_info: dict[str, Any],
    after_info: dict[str, Any],
    grading: dict[str, Any],
) -> bool:
    sec_idx = grading.get("target_section_index", 1)
    sec_a = next((s for s in after_info["sections"] if s["section_index"] == sec_idx), None)
    if sec_a is None:
        return False
    if "expected_orientation" in grading and sec_a["orientation"] != grading["expected_orientation"]:
        return False
    for side, exp_cm in grading.get("expected_section_margins_cm", {}).items():
        if abs(sec_a["margins_cm"].get(side, -999) - exp_cm) > 0.05:
            return False
    if grading.get("keep_section1_portrait"):
        b1, a1 = before_info["sections"][0], after_info["sections"][0]
        if a1["orientation"] != "portrait" or any(
            abs(b1["margins_cm"][k] - a1["margins_cm"][k]) > 0.05 for k in ("top", "bottom", "left", "right")
        ):
            return False
    return True


def _check_merge_runs(p_before: dict[str, Any], p_after: dict[str, Any], grading: dict[str, Any]) -> bool:
    if p_after["run_count"] > grading.get("max_run_count_after", 100):
        return False
    if p_before["text_nfc"] != p_after["text_nfc"]:
        return False
    return p_before["char_bold_seq"] == p_after["char_bold_seq"]


def _check_out_of_scope_assertions(
    after_info: dict[str, Any],
    after_parts: dict[str, bytes],
    grading: dict[str, Any],
    target_para_ordinals: list[int],
) -> bool:
    """Check Group E out-of-scope assertions (footers, headers, footnotes, nested tables, hanging indents)."""
    checked_any = False

    if "expected_footer_text_contains" in grading:
        checked_any = True
        part_filter = grading.get("target_footer_part", "any")
        hf = after_info.get("headers_footers", {})
        candidates = [
            v["text_nfc"]
            for k, v in hf.items()
            if k.startswith("word/footer") and (part_filter == "any" or k == part_filter)
        ]
        if not candidates:
            return False
        joined = "\n".join(candidates)
        for sub in grading["expected_footer_text_contains"]:
            if nfc(sub) not in joined:
                return False

    if "expected_header_text_contains" in grading:
        checked_any = True
        part_filter = grading.get("target_header_part", "any")
        hf = after_info.get("headers_footers", {})
        candidates = [
            v["text_nfc"]
            for k, v in hf.items()
            if k.startswith("word/header") and (part_filter == "any" or k == part_filter)
        ]
        if not candidates:
            return False
        joined = "\n".join(candidates)
        for sub in grading["expected_header_text_contains"]:
            if nfc(sub) not in joined:
                return False

    if "expected_footnote_font_pt" in grading or "expected_footnote_font_name" in grading:
        checked_any = True
        fn_bytes = after_parts.get("word/footnotes.xml")
        if not fn_bytes:
            return False
        fn_root = parse_xml(fn_bytes)
        styles_ctx = after_info["styles_ctx"]
        saw_run = False
        for fn in fn_root.findall(f"{{{W}}}footnote"):
            fn_id = fn.get(f"{{{W}}}id")
            if fn_id in ("-1", "0"):
                continue
            for p_el in fn.findall(f"{{{W}}}p"):
                for r_el in p_el.findall(f"{{{W}}}r"):
                    if not text_of_element(r_el).strip():
                        continue
                    saw_run = True
                    rf = effective_run_format(r_el, p_el, styles_ctx)
                    if "expected_footnote_font_pt" in grading:
                        sz = rf.get("size_pt")
                        if sz is None or abs(sz - grading["expected_footnote_font_pt"]) > 0.1:
                            return False
                    if "expected_footnote_font_name" in grading:
                        want_f = grading["expected_footnote_font_name"]
                        if rf.get("font_ascii") != want_f and rf.get("font_hansi") != want_f:
                            return False
        if not saw_run:
            return False

    if "expected_nested_table_alignment" in grading or "expected_nested_table_shading_fill" in grading:
        checked_any = True
        nested = after_info.get("nested_tables", [])
        n_ord = grading.get("target_nested_table_ordinal", 1)
        if len(nested) < n_ord:
            return False
        nt = nested[n_ord - 1]
        if "expected_nested_table_alignment" in grading:
            if nt.get("alignment") != grading["expected_nested_table_alignment"]:
                return False
        if "expected_nested_table_shading_fill" in grading:
            want_fill = grading["expected_nested_table_shading_fill"].upper()
            fills = nt.get("header_shading_fills", [])
            if not fills or any((f or "").upper() != want_fill for f in fills):
                return False

    if "expected_hanging_indent_cm_range" in grading or target_para_ordinals:
        checked_any = True
        if not target_para_ordinals:
            return False
        if not all(_check_paragraph_assertions(after_info["paragraphs"][o - 1], grading) for o in target_para_ordinals):
            return False

    return checked_any


def grade_case(
    case: dict[str, Any],
    src_docx: str | Path,
    submission_dir: str | Path,
    *,
    run_render: bool = True,
    renderer_backend: str = "libreoffice",
    ecma376_xsd_path: str | Path | None = None,
    openxml_validator_cmd: str | None = None,
) -> dict[str, Any]:
    """Grade one submission directory independently of scripts/."""
    cid = case.get("case_id") or case["id"]
    src_path = Path(src_docx)
    sub_dir = Path(submission_dir)
    dst_docx = sub_dir / f"{cid}.docx"
    res_file = sub_dir / f"{cid}.result.json"

    res_json: dict[str, Any] | None = None
    if res_file.is_file():
        try:
            loaded = json.loads(res_file.read_text(encoding="utf-8-sig"))
            if isinstance(loaded, dict):
                res_json = loaded
        except (json.JSONDecodeError, OSError):
            res_json = None

    expected_outcome = case["expected_outcome"]
    grading = case["grading"]
    src_sha = _sha256(src_path)

    base_meta: dict[str, Any] = {
        "case_id": cid,
        "taxonomy_group": case.get("taxonomy_group"),
        "length_tier": case.get("length_tier"),
        "instruction_type": case.get("instruction_type"),
        "expected_outcome": expected_outcome,
        "near_boundary_valid": bool(case.get("near_boundary_valid", False)),
        "validator_backends": [],
        "renderer_backend": renderer_backend if (run_render and expected_outcome == "edit") else "none",
        "over_refusal": False,
        "unsafe_edit": False,
    }

    if expected_outcome == "refuse":
        status_refused = bool(
            res_json is not None
            and res_json.get("status") == "refused"
            and str(res_json.get("reason", "")).strip()
        )
        no_mutated = (not dst_docx.is_file()) or (_sha256(dst_docx) == src_sha)
        checkpoints = {
            "refusal_status": status_refused,
            "no_mutated_output": no_mutated,
        }
        unsafe_edit = not no_mutated or (res_json is not None and res_json.get("status") == "edited")
        hit = 1 if all(checkpoints.values()) else 0
        return {
            **base_meta,
            "hit": hit,
            "checkpoint_score": round(sum(1 for v in checkpoints.values() if v) / len(checkpoints), 4),
            "checkpoints": checkpoints,
            "unsafe_edit": bool(unsafe_edit),
            "failure_reasons": [k for k, ok in checkpoints.items() if not ok],
            "telemetry": {
                "xml_ntpr": 1.0 if no_mutated else 0.0,
                "zip_parts_byte_identical": 1.0 if no_mutated else 0.0,
                "semantic_non_target_preserved": no_mutated,
                "submitted_status": res_json.get("status") if res_json else None,
            },
        }

    # Edit branch.
    target_kind = grading.get("target_kind", "table")
    has_tbl_style = any(
        k in grading
        for k in (
            "expected_cell_margins_cm",
            "expected_data_font_pt",
            "expected_header_font_pt",
            "require_header_bold",
            "require_no_cell_margin_overrides",
            "expected_repeat_header",
            "expected_cant_split",
            "expected_border_style",
            "expected_border_color",
            "expected_header_shading_fill",
        )
    )

    checkpoints: dict[str, bool] = {"valid_output_pkg": False}
    if target_kind in ("table", "multi"):
        checkpoints["ooxml_target_width"] = False
        if has_tbl_style:
            checkpoints["ooxml_table_margins_and_fonts"] = False
        if run_render:
            checkpoints["pdf_geometry_match"] = False
    if target_kind in ("image", "multi"):
        checkpoints["ooxml_target_image"] = False
        if run_render:
            checkpoints["pdf_geometry_match" if target_kind == "image" else "pdf_image_geometry_match"] = False
    if target_kind in ("paragraph", "multi"):
        checkpoints["ooxml_target_paragraph"] = False
    if target_kind == "merge_runs":
        checkpoints["ooxml_runs_merged"] = False
    if target_kind == "style":
        checkpoints["ooxml_target_style"] = False
        if grading.get("require_document_xml_unchanged"):
            checkpoints["document_xml_unchanged"] = False
    if target_kind == "section":
        checkpoints["ooxml_target_section"] = False
    if target_kind == "out_of_scope":
        checkpoints["ooxml_out_of_scope_target"] = False
    if target_kind == "paginate":
        if run_render:
            checkpoints["page_count_target_met"] = False
            if grading.get("require_no_orphan_signature"):
                checkpoints["no_orphan_signature"] = False
        if grading.get("validate_profile"):
            checkpoints["profile_validation_pass"] = False
    checkpoints["ooxml_order_clean"] = False
    checkpoints["semantic_non_target_preserved"] = False
    if grading.get("require_track_changes_tags"):
        checkpoints["ooxml_track_changes_present"] = False
    if grading.get("require_namespace_audit_pass") or grading.get("require_mc_alternate_preserved"):
        checkpoints["ooxml_namespaces_preserved"] = False
    if grading.get("require_header_footer_byte_identical"):
        checkpoints["header_footer_byte_identical"] = False
    if grading.get("expected_pagination_ripple") is not None and run_render:
        checkpoints["pagination_ripple_reported"] = False
    if grading.get("validate_profile") and target_kind != "paginate":
        checkpoints["profile_validation_pass"] = False

    telemetry: dict[str, Any] = {
        "xml_ntpr": None,
        "zip_parts_byte_identical": None,
        "pdf_width_before_cm": None,
        "pdf_width_after_cm": None,
        "ooxml_width_after_cm": None,
        "ooxml_image_width_after_cm": None,
        "pdf_image_width_after_cm": None,
        "overflow_right_mm_after": None,
        "alignment_after": None,
        "pagination_ripple_detected": None,
        "submitted_status": res_json.get("status") if res_json else None,
    }

    if res_json is not None and res_json.get("status") == "refused":
        return {
            **base_meta,
            "hit": 0,
            "checkpoint_score": 0.0,
            "checkpoints": checkpoints,
            "over_refusal": True,
            "failure_reasons": ["over_refusal"],
            "telemetry": telemetry,
        }

    if not dst_docx.is_file():
        return {
            **base_meta,
            "hit": 0,
            "checkpoint_score": 0.0,
            "checkpoints": checkpoints,
            "failure_reasons": ["missing_output_docx"],
            "telemetry": telemetry,
        }

    val_res = validate_package(
        dst_docx,
        baseline_docx=src_path,
        ecma376_xsd_path=ecma376_xsd_path,
        openxml_validator_cmd=openxml_validator_cmd,
    )
    base_meta["validator_backends"] = val_res["validator_backends"]
    telemetry["validation_new_errors"] = val_res["new_errors"]
    telemetry["validation_all_error_count"] = val_res["all_error_count"]

    try:
        _, before_parts = read_zip_parts(src_path)
        _, after_parts = read_zip_parts(dst_docx)
        before_info = inspect_docx(src_path)
        after_info = inspect_docx(dst_docx)
        before_root = parse_xml(before_parts["word/document.xml"])
        after_root = parse_xml(after_parts["word/document.xml"])
    except Exception as exc:
        return {
            **base_meta,
            "hit": 0,
            "checkpoint_score": 0.0,
            "checkpoints": checkpoints,
            "failure_reasons": ["invalid_docx_package"],
            "telemetry": {**telemetry, "error": str(exc)},
        }

    opc_clean = not any(e["layer"] == "opc" for e in val_res["new_errors"])
    status_edited_ok = (res_json is None) or (res_json.get("status") == "edited")
    checkpoints["valid_output_pkg"] = bool(opc_clean and status_edited_ok)
    checkpoints["ooxml_order_clean"] = bool(val_res["valid"])

    target_tbl_ord = grading.get("target_ordinal")
    target_img_ord = grading.get("target_image_ordinal")
    if len(before_info["tables"]) != len(after_info["tables"]) or (
        target_tbl_ord is not None and len(before_info["tables"]) < target_tbl_ord
    ):
        return {
            **base_meta,
            "hit": 0,
            "checkpoint_score": 0.0,
            "checkpoints": checkpoints,
            "failure_reasons": ["table_count_mismatch"],
            "telemetry": telemetry,
        }
    if len(before_info["images"]) != len(after_info["images"]) or (
        target_img_ord is not None and len(before_info["images"]) < target_img_ord
    ):
        return {
            **base_meta,
            "hit": 0,
            "checkpoint_score": 0.0,
            "checkpoints": checkpoints,
            "failure_reasons": ["image_count_mismatch"],
            "telemetry": telemetry,
        }

    target_para_ordinals: list[int] = []
    need_paras = (
        target_kind in ("paragraph", "merge_runs", "multi")
        or (
            target_kind == "out_of_scope"
            and (
                "target_paragraph_ordinals" in grading
                or "target_paragraph_text_match" in grading
                or "expected_hanging_indent_cm_range" in grading
            )
        )
    )
    if need_paras:
        if len(before_info["paragraphs"]) != len(after_info["paragraphs"]):
            return {
                **base_meta,
                "hit": 0,
                "checkpoint_score": 0.0,
                "checkpoints": checkpoints,
                "failure_reasons": ["paragraph_count_mismatch"],
                "telemetry": telemetry,
            }
        for ord_num in grading.get("target_paragraph_ordinals", []):
            if 1 <= ord_num <= len(before_info["paragraphs"]) and ord_num not in target_para_ordinals:
                target_para_ordinals.append(ord_num)
        if grading.get("target_paragraph_text_match"):
            want_txt = nfc(grading["target_paragraph_text_match"])
            for p_item in before_info["paragraphs"]:
                if p_item["text_nfc"] == want_txt and p_item["ordinal"] not in target_para_ordinals:
                    target_para_ordinals.append(p_item["ordinal"])
        if not target_para_ordinals:
            return {
                **base_meta,
                "hit": 0,
                "checkpoint_score": 0.0,
                "checkpoints": checkpoints,
                "failure_reasons": ["target_paragraph_not_found"],
                "telemetry": telemetry,
            }

    # Compute allowed body indices for XML NTPR telemetry.
    allowed_bidx: set[int] = set()
    t_before = before_info["tables"][target_tbl_ord - 1] if target_tbl_ord else None
    t_after = after_info["tables"][target_tbl_ord - 1] if target_tbl_ord else None
    if t_before is not None:
        allowed_bidx.add(t_before["body_index"])

    img_before = before_info["images"][target_img_ord - 1] if target_img_ord else None
    img_after = after_info["images"][target_img_ord - 1] if target_img_ord else None
    if img_before is not None:
        allowed_bidx.add(img_before["body_index"])

    for p_ord in target_para_ordinals:
        allowed_bidx.add(before_info["paragraphs"][p_ord - 1]["body_index"])

    if target_kind == "section":
        sec_idx = grading.get("target_section_index", 1)
        sec_b = next((s for s in before_info["sections"] if s["section_index"] == sec_idx), None)
        if sec_b is not None:
            allowed_bidx.add(sec_b["body_index"])

    if target_kind == "paginate":
        for bt in before_info["tables"]:
            if not bt.get("is_layout_table"):
                allowed_bidx.add(bt["body_index"])
        for bi in before_info["images"]:
            allowed_bidx.add(bi["body_index"])
        for bp in before_info["paragraphs"]:
            if not (bp.get("style_id") or "").lower().startswith(("heading", "title")):
                allowed_bidx.add(bp["body_index"])
        for bs in before_info["sections"]:
            allowed_bidx.add(bs["body_index"])

    allow_hf: set[str] = set()
    if target_kind == "out_of_scope":
        if "expected_footer_text_contains" in grading:
            fp = grading.get("target_footer_part", "any")
            if fp == "any":
                allow_hf.update(k for k in after_info.get("headers_footers", {}) if k.startswith("word/footer"))
                allow_hf.update(k for k in before_info.get("headers_footers", {}) if k.startswith("word/footer"))
            else:
                allow_hf.add(fp)
        if "expected_header_text_contains" in grading:
            hp = grading.get("target_header_part", "any")
            if hp == "any":
                allow_hf.update(k for k in after_info.get("headers_footers", {}) if k.startswith("word/header"))
                allow_hf.update(k for k in before_info.get("headers_footers", {}) if k.startswith("word/header"))
            else:
                allow_hf.add(hp)

    allow_fn = target_kind == "out_of_scope" and (
        "expected_footnote_font_pt" in grading or "expected_footnote_font_name" in grading
    )

    ntpr_res = check_non_target_preservation(
        before_info,
        after_info,
        before_parts,
        after_parts,
        target_kind=target_kind,
        target_table_ordinals={target_tbl_ord} if target_tbl_ord else set(),
        target_image_ordinals={target_img_ord} if target_img_ord else set(),
        target_para_ordinals=set(target_para_ordinals),
        target_section_indices={grading.get("target_section_index", 1)} if target_kind == "section" else set(),
        allowed_body_indices=allowed_bidx,
        require_layout_tables_unchanged=grading.get("require_layout_tables_unchanged", True),
        allow_header_footer_parts=allow_hf or None,
        allow_footnote_changes=allow_fn,
    )
    checkpoints["semantic_non_target_preserved"] = ntpr_res["semantic_non_target_preserved"]
    telemetry["xml_ntpr"] = ntpr_res["xml_ntpr"]
    telemetry["zip_parts_byte_identical"] = ntpr_res["zip_parts_byte_identical"]
    telemetry["non_target_details"] = ntpr_res["details"]

    if "header_footer_byte_identical" in checkpoints:
        checkpoints["header_footer_byte_identical"] = ntpr_res["header_footer_byte_identical"]

    if target_kind in ("table", "multi") and t_before is not None and t_after is not None:
        telemetry["ooxml_width_after_cm"] = t_after["total_width_cm"]
        telemetry["alignment_after"] = t_after["alignment"]
        w_ok, st_ok = _check_table_assertions(t_before, t_after, grading)
        checkpoints["ooxml_target_width"] = w_ok
        if "ooxml_table_margins_and_fonts" in checkpoints:
            checkpoints["ooxml_table_margins_and_fonts"] = st_ok

    if target_kind in ("image", "multi") and img_before is not None and img_after is not None:
        telemetry["ooxml_image_width_after_cm"] = img_after["width_cm"]
        if target_kind == "image":
            telemetry["alignment_after"] = img_after["alignment"]
        checkpoints["ooxml_target_image"] = _check_image_assertions(img_before, img_after, grading)

    if target_kind in ("paragraph", "multi"):
        checkpoints["ooxml_target_paragraph"] = all(
            _check_paragraph_assertions(after_info["paragraphs"][ord_num - 1], grading)
            for ord_num in target_para_ordinals
        )

    if target_kind == "merge_runs":
        checkpoints["ooxml_runs_merged"] = all(
            _check_merge_runs(
                before_info["paragraphs"][ord_num - 1],
                after_info["paragraphs"][ord_num - 1],
                grading,
            )
            for ord_num in target_para_ordinals
        )

    if target_kind == "style":
        checkpoints["ooxml_target_style"] = _check_style_assertions(after_info, grading)
        if "document_xml_unchanged" in checkpoints:
            checkpoints["document_xml_unchanged"] = (
                before_parts.get("word/document.xml") == after_parts.get("word/document.xml")
            )

    if target_kind == "section":
        checkpoints["ooxml_target_section"] = _check_section_assertions(before_info, after_info, grading)

    if target_kind == "out_of_scope":
        checkpoints["ooxml_out_of_scope_target"] = _check_out_of_scope_assertions(
            after_info,
            after_parts,
            grading,
            target_para_ordinals,
        )

    if "ooxml_track_changes_present" in checkpoints and t_after is not None:
        checkpoints["ooxml_track_changes_present"] = bool(t_after.get("has_track_changes"))

    if "ooxml_namespaces_preserved" in checkpoints:
        ns_res = audit_namespaces(before_parts["word/document.xml"], after_parts["word/document.xml"])
        alt_ok = True
        if grading.get("require_mc_alternate_preserved"):
            b_alts = [etree.tostring(el) for el in before_root.iter(f"{{{MC_URI}}}AlternateContent")]
            a_alts = [etree.tostring(el) for el in after_root.iter(f"{{{MC_URI}}}AlternateContent")]
            alt_ok = len(b_alts) >= 1 and b_alts == a_alts
        checkpoints["ooxml_namespaces_preserved"] = bool(ns_res["preserved"] and alt_ok)

    if "profile_validation_pass" in checkpoints:
        comp = check_compliance_profile(after_info, grading["validate_profile"])
        checkpoints["profile_validation_pass"] = comp["status"] == "pass"
        telemetry["compliance_issues"] = comp["issues"]

    # Independent rendering & geometric checks.
    if run_render:
        with tempfile.TemporaryDirectory(prefix="eval-render-") as rtmp:
            tmp_dir = Path(rtmp)
            need_before_pdf = (
                target_kind in ("table", "multi", "image")
                or "pagination_ripple_reported" in checkpoints
                or "expected_width_max_drift_cm" in grading
            )
            b_ren = render_to_pdf(src_path, tmp_dir / "before", backend=renderer_backend) if need_before_pdf else None
            a_ren = render_to_pdf(dst_docx, tmp_dir / "after", backend=renderer_backend)
            base_meta["renderer_backend"] = a_ren.get("renderer_backend", renderer_backend)

            if a_ren.get("status") != "rendered":
                checkpoints["valid_output_pkg"] = False
                telemetry["render_error"] = a_ren.get("reason")
            else:
                after_pdf = a_ren["pdf_path"]
                after_layout = inspect_pdf_layout(after_pdf, after_info)
                before_layout = (
                    inspect_pdf_layout(b_ren["pdf_path"], before_info)
                    if (b_ren and b_ren.get("status") == "rendered")
                    else None
                )

                if target_kind in ("table", "multi") and target_tbl_ord is not None and t_after is not None:
                    a_tpdf = after_layout["tables"][target_tbl_ord - 1]
                    b_tpdf = before_layout["tables"][target_tbl_ord - 1] if before_layout else {}
                    telemetry["pdf_width_before_cm"] = b_tpdf.get("width_cm")
                    telemetry["pdf_width_after_cm"] = a_tpdf.get("width_cm")
                    telemetry["overflow_right_mm_after"] = a_tpdf.get("overflow_right_mm")

                    if not a_tpdf.get("matched") or a_tpdf.get("width_cm") is None:
                        checkpoints["pdf_geometry_match"] = False
                    else:
                        pw = a_tpdf["width_cm"]
                        if "expected_width_min_cm" in grading and "expected_width_max_cm" in grading:
                            pw_ok = (grading["expected_width_min_cm"] - 0.15) <= pw <= (grading["expected_width_max_cm"] + 0.15)
                        elif "expected_total_width_cm" in grading:
                            pw_ok = abs(pw - grading["expected_total_width_cm"]) <= (grading.get("width_tolerance_cm", 0.35) + 0.15)
                        elif "expected_width_max_drift_cm" in grading:
                            bw = b_tpdf.get("width_cm")
                            pw_ok = bw is not None and abs(pw - bw) <= (grading["expected_width_max_drift_cm"] + 0.15)
                        else:
                            pw_ok = True
                        err_ok = abs(pw - t_after["total_width_cm"]) <= 0.5
                        ovf_ok = (a_tpdf.get("overflow_right_mm") or 0.0) <= grading.get("max_overflow_right_mm", 0.5)
                        al_ok = True
                        if grading.get("expected_alignment") == "center":
                            sec_a = next(s for s in after_info["sections"] if s["section_index"] == t_after["section_index"])
                            exp_c = sec_a["margins_cm"]["left"] + sec_a["printable_width_cm"] / 2.0
                            al_ok = a_tpdf.get("center_cm") is not None and abs(a_tpdf["center_cm"] - exp_c) <= 0.65
                        checkpoints["pdf_geometry_match"] = bool(pw_ok and err_ok and ovf_ok and al_ok)

                if target_kind in ("image", "multi") and target_img_ord is not None and img_after is not None:
                    img_key = "pdf_geometry_match" if target_kind == "image" else "pdf_image_geometry_match"
                    a_ipdf = after_layout["images"][target_img_ord - 1]
                    b_ipdf = before_layout["images"][target_img_ord - 1] if before_layout else {}
                    telemetry["pdf_image_width_after_cm"] = a_ipdf.get("width_cm")
                    if target_kind == "image":
                        telemetry["pdf_width_before_cm"] = b_ipdf.get("width_cm")
                        telemetry["pdf_width_after_cm"] = a_ipdf.get("width_cm")
                        telemetry["overflow_right_mm_after"] = a_ipdf.get("overflow_right_mm")
                    if not a_ipdf.get("matched") or a_ipdf.get("width_cm") is None:
                        checkpoints[img_key] = False
                    else:
                        piw = a_ipdf["width_cm"]
                        tol_i = grading.get("image_width_tolerance_cm", 0.25) + 0.25
                        iw_ok = abs(piw - grading["expected_image_width_cm"]) <= tol_i if "expected_image_width_cm" in grading else True
                        ovf_ok = (a_ipdf.get("overflow_right_mm") or 0.0) <= grading.get("max_overflow_right_mm", 0.5)
                        al_ok = True
                        if grading.get("expected_image_alignment") == "center":
                            sec_a = next(s for s in after_info["sections"] if s["section_index"] == img_after["section_index"])
                            exp_c = sec_a["margins_cm"]["left"] + sec_a["printable_width_cm"] / 2.0
                            al_ok = a_ipdf.get("center_cm") is not None and abs(a_ipdf["center_cm"] - exp_c) <= 0.7
                        checkpoints[img_key] = bool(iw_ok and ovf_ok and al_ok)

                if target_kind == "paginate":
                    with pymupdf.open(after_pdf) as pdoc:
                        pages_after = len(pdoc)
                    telemetry["pages_after"] = pages_after
                    checkpoints["page_count_target_met"] = pages_after <= grading["expected_max_pages"]
                    if "no_orphan_signature" in checkpoints:
                        checkpoints["no_orphan_signature"] = not detect_orphan_signature_pdf(after_pdf)

                if b_ren and b_ren.get("status") == "rendered":
                    probes = [p["text_nfc"] for p in before_info["paragraphs"][-3:] if len(p["text_nfc"].strip()) >= 10]
                    ripple = detect_pagination_ripple(b_ren["pdf_path"], after_pdf, probe_texts=probes)
                    telemetry["pagination_ripple_detected"] = ripple
                    if "pagination_ripple_reported" in checkpoints:
                        checkpoints["pagination_ripple_reported"] = bool(
                            res_json is not None
                            and res_json.get("pagination_ripple") is grading["expected_pagination_ripple"]
                            and ripple is True
                        )

    hit = 1 if all(checkpoints.values()) else 0
    return {
        **base_meta,
        "hit": hit,
        "checkpoint_score": round(sum(1 for v in checkpoints.values() if v) / len(checkpoints), 4),
        "checkpoints": checkpoints,
        "failure_reasons": [k for k, ok in checkpoints.items() if not ok],
        "telemetry": telemetry,
    }


MC_URI = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def grade_directory(
    cases_path: str | Path,
    submission_dir: str | Path,
    corpus_dir: str | Path,
    *,
    run_render: bool = True,
    renderer_backend: str = "libreoffice",
) -> list[dict[str, Any]]:
    """Grade every case listed in a cases JSON manifest."""
    cfg = json.loads(Path(cases_path).read_text(encoding="utf-8"))
    cases = cfg["cases"] if isinstance(cfg, dict) and "cases" in cfg else cfg
    cdir = Path(corpus_dir)
    sdir = Path(submission_dir)
    results = []
    for case in cases:
        cid = case.get("case_id") or case["id"]
        src_rel = case.get("source_docx_relpath")
        src_docx = cdir / src_rel if src_rel else (cdir / f"{cid}.docx")
        results.append(
            grade_case(
                case,
                src_docx,
                sdir / cid,
                run_render=run_render,
                renderer_backend=renderer_backend,
            )
        )
    return results


def crosscheck_dev_ab(
    cases_path: str | Path,
    published_results_path: str | Path,
    ab_roots: list[str | Path],
    *,
    output_json: str | Path | None = None,
    output_md: str | Path | None = None,
) -> dict[str, Any]:
    """Re-grade all saved dev-v2.0 Live A/B trajectories using eval_harness and compare verdicts."""
    cases_cfg = json.loads(Path(cases_path).read_text(encoding="utf-8"))
    cases_by_id = {(c.get("case_id") or c["id"]): c for c in cases_cfg["cases"]}
    published = json.loads(Path(published_results_path).read_text(encoding="utf-8"))

    pub_lookup: dict[tuple[str, str, str], dict[str, Any]] = {}
    for section_key in ("live_ab_phase1_t01_t12", "live_ab_phase1b_t13_t18"):
        for r in published.get(section_key, {}).get("runs", []):
            pub_lookup[(r["model_tier"], r["arm"], r["case_id"])] = r

    rows: list[dict[str, Any]] = []
    disagreements: list[dict[str, Any]] = []

    for root_raw in ab_roots:
        root = Path(root_raw)
        rm_path = root / "runs_manifest.json"
        if not rm_path.is_file():
            continue
        rm = json.loads(rm_path.read_text(encoding="utf-8"))
        for entry in rm.get("runs", []):
            cid = entry["case_id"]
            case = cases_by_id[cid]
            src_docx = Path(entry["src_docx"])
            if not src_docx.is_file():
                # Resolve relative to repo cwd if needed.
                src_docx = Path.cwd() / entry["src_docx"]
            work_dir = Path(entry["work_dir"])
            if not work_dir.is_dir():
                work_dir = Path.cwd() / entry["work_dir"]

            ind = grade_case(case, src_docx, work_dir, run_render=True, renderer_backend="libreoffice")
            pub = pub_lookup.get((entry["model_tier"], entry["arm"], cid), {})
            pub_hit = pub.get("hit")
            ind_hit = ind["hit"]
            agrees = pub_hit == ind_hit
            rec = {
                "batch": root.name,
                "model_tier": entry["model_tier"],
                "arm": entry["arm"],
                "case_id": cid,
                "published_coupled_hit": pub_hit,
                "independent_hit": ind_hit,
                "agrees": agrees,
                "published_cp": pub.get("checkpoint_score"),
                "independent_cp": ind["checkpoint_score"],
                "independent_failures": ind["failure_reasons"],
                "validation_new_errors": ind["telemetry"].get("validation_new_errors", []),
                "validation_all_error_count": ind["telemetry"].get("validation_all_error_count", 0),
                "validator_backends": ind["validator_backends"],
                "renderer_backend": ind["renderer_backend"],
            }
            rows.append(rec)
            if not agrees:
                disagreements.append(rec)

    summary = {
        "total_runs_checked": len(rows),
        "agreements": sum(1 for r in rows if r["agrees"]),
        "disagreements": len(disagreements),
        "agreement_rate": round(sum(1 for r in rows if r["agrees"]) / len(rows), 4) if rows else 0.0,
        "disagreement_details": disagreements,
        "runs": rows,
    }

    if output_json:
        out_j = Path(output_json)
        out_j.parent.mkdir(parents=True, exist_ok=True)
        out_j.write_text(json.dumps(summary, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    if output_md:
        lines = [
            "# Independent Cross-Check of `dev-v2.0` Live A/B Trajectories",
            "",
            f"- Total trajectories re-graded independently: **{summary['total_runs_checked']}**",
            f"- Verdict agreement (`published_coupled_hit == independent_hit`): **{summary['agreements']}/{summary['total_runs_checked']} ({summary['agreement_rate']*100:.1f}%)**",
            f"- Disagreements found: **{summary['disagreements']}**",
            "",
            "| Batch | Model | Arm | Case | Coupled Hit | Independent Hit | Indep CP | New Schema Errors | Failures |",
            "|---|---|---|---|---:|---:|---:|---:|---|",
        ]
        for r in rows:
            fails = ", ".join(r["independent_failures"]) or "-"
            lines.append(
                f"| `{r['batch']}` | `{r['model_tier']}` | `{r['arm']}` | `{r['case_id']}` | "
                f"{r['published_coupled_hit']} | {r['independent_hit']} | {r['independent_cp']*100:.0f}% | "
                f"{len(r['validation_new_errors'])} | {fails} |"
            )
        out_m = Path(output_md)
        out_m.parent.mkdir(parents=True, exist_ok=True)
        out_m.write_text("\n".join(lines) + "\n", encoding="utf-8")

    return summary


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--grade-dir", metavar="SUBMISSION_DIR", help="Grade submission directory")
    mode.add_argument("--crosscheck-dev-ab", action="store_true", help="Cross-check saved dev-v2.0 Live A/B trajectories")
    mode.add_argument("--validate-file", metavar="DOCX_PATH", help="Validate one DOCX package independently")

    parser.add_argument("--cases", default="evals/layout_edit_cases.json", help="Path to cases JSON")
    parser.add_argument("--corpus-dir", help="Corpus directory (for --grade-dir)")
    parser.add_argument("--published-results", default="evals/layout-benchmark-results.json", help="Published results JSON")
    parser.add_argument("--ab-roots", nargs="+", default=["research/eval/live_ab", "research/eval/phase1b_ab"])
    parser.add_argument("--output-json", help="Optional JSON report path")
    parser.add_argument("--output-md", help="Optional Markdown report path")
    parser.add_argument("--no-render", action="store_true", help="Skip LibreOffice PDF rendering")
    args = parser.parse_args(argv)

    if args.validate_file:
        out: Any = validate_package(args.validate_file)
    elif args.grade_dir:
        if not args.corpus_dir:
            parser.error("--grade-dir requires --corpus-dir")
        out = grade_directory(
            args.cases,
            args.grade_dir,
            args.corpus_dir,
            run_render=not args.no_render,
        )
    else:
        out = crosscheck_dev_ab(
            args.cases,
            args.published_results,
            args.ab_roots,
            output_json=args.output_json,
            output_md=args.output_md,
        )

    print(json.dumps(out, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
