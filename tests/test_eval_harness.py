"""Tests for independent eval_harness isolation, schema validation, OOXML reading, grading, and stats."""

from __future__ import annotations

import ast
import hashlib
import json
import struct
import zipfile
import zlib
from pathlib import Path

import pytest
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt
from lxml import etree

from eval_harness.curator import (
    check_disjoint,
    lint_prompt_zero_leakage,
    seal_commitment,
    validate_manifest,
)
from eval_harness.grader import check_compliance_profile, grade_case
from eval_harness.runner import (
    ALL_ARMS,
    ARM_FULL_SKILL,
    ARM_NO_SKILL,
    ARM_STANDARDS_ONLY,
    _is_codex_quota_error,
    _try_resume_codex_case,
    append_val_registry,
    evaluate_experiment,
    parse_antigravity_transcript,
    parse_codex_jsonl,
    prepare_arm_workspace,
    render_experiment_markdown,
    scan_transcript_contamination,
)
from eval_harness.ooxml_reader import DXA_PER_CM, W, inspect_docx
from eval_harness.render_inspector import find_soffice
from eval_harness.schema_validator import validate_package
from eval_harness.stats import (
    bootstrap_ci,
    compare_arms,
    mcnemar_exact,
    normalized_gain,
    pass_at_1,
    pass_hat_k,
    wilcoxon_signed_rank,
)

ROOT = Path(__file__).resolve().parent.parent
HARNESS_DIR = ROOT / "eval_harness"
SCRIPTS_DIR = ROOT / "scripts"


def _make_png(path: Path, width: int = 400, height: int = 200) -> Path:
    def _chunk(tag: bytes, data: bytes) -> bytes:
        crc = zlib.crc32(tag + data) & 0xFFFFFFFF
        return struct.pack("!I", len(data)) + tag + data + struct.pack("!I", crc)

    raw = bytearray()
    row = b"\x00" + (b"\x20\x60\xb0" * width)
    for _ in range(height):
        raw.extend(row)
    ihdr = struct.pack("!IIBBBBB", width, height, 8, 2, 0, 0, 0)
    payload = (
        b"\x89PNG\r\n\x1a\n"
        + _chunk(b"IHDR", ihdr)
        + _chunk(b"IDAT", zlib.compress(bytes(raw), 6))
        + _chunk(b"IEND", b"")
    )
    path.write_bytes(payload)
    return path


def _mutate_zip_part(src: Path, dst: Path, replacements: dict[str, bytes], *, add_dir_entry: str | None = None) -> None:
    with zipfile.ZipFile(src, "r") as zin:
        parts = {info.filename: zin.read(info.filename) for info in zin.infolist() if not info.is_dir()}
    parts.update(replacements)
    with zipfile.ZipFile(dst, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        if add_dir_entry:
            zout.writestr(add_dir_entry, b"")
        for name, data in parts.items():
            zout.writestr(name, data)


def _set_table_dxa_widths(tbl, widths_cm: list[float], *, align: str = "left") -> None:
    w_ns = f"{{{W}}}"
    dxa_list = [round(w * DXA_PER_CM) for w in widths_cm]
    tbl_pr = tbl._tbl.find(f"{w_ns}tblPr")
    tbl_w = tbl_pr.find(f"{w_ns}tblW")
    tbl_w.set(f"{w_ns}w", str(sum(dxa_list)))
    tbl_w.set(f"{w_ns}type", "dxa")
    jc = tbl_pr.find(f"{w_ns}jc")
    if jc is None:
        jc = etree.SubElement(tbl_pr, f"{w_ns}jc")
    jc.set(f"{w_ns}val", align)

    tbl_grid = tbl._tbl.find(f"{w_ns}tblGrid")
    for gcol, w_dxa in zip(tbl_grid.findall(f"{w_ns}gridCol"), dxa_list):
        gcol.set(f"{w_ns}w", str(w_dxa))
    for row in tbl.rows:
        for cell, w_dxa in zip(row.cells, dxa_list):
            tc_pr = cell._tc.find(f"{w_ns}tcPr")
            tc_w = tc_pr.find(f"{w_ns}tcW")
            tc_w.set(f"{w_ns}w", str(w_dxa))
            tc_w.set(f"{w_ns}type", "dxa")


def test_eval_harness_has_zero_scripts_or_evals_imports():
    forbidden_roots = {
        "scripts",
        "evals",
        "run_layout_benchmark",
        "run_validation_benchmark",
        "scipy",
        "PIL",
        "fitz",
    }
    forbidden_roots.update(p.stem for p in SCRIPTS_DIR.glob("*.py") if p.stem != "__init__")

    py_files = sorted(HARNESS_DIR.rglob("*.py"))
    assert len(py_files) >= 6
    for py_path in py_files:
        tree = ast.parse(py_path.read_text(encoding="utf-8"), filename=str(py_path))
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                for alias in node.names:
                    top = alias.name.split(".")[0]
                    assert top not in forbidden_roots and alias.name not in forbidden_roots, (
                        f"Forbidden import {alias.name!r} in {py_path}"
                    )
            elif isinstance(node, ast.ImportFrom):
                mod = node.module or ""
                top = mod.split(".")[0]
                assert top not in forbidden_roots and mod not in forbidden_roots, (
                    f"Forbidden from-import {mod!r} in {py_path}"
                )


def test_dev_v2_frozen_manifest_matches_files():
    manifest_path = ROOT / "evals" / "snapshots" / "dev-v2.0-manifest.json"
    assert manifest_path.is_file()
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    assert manifest["snapshot_id"] == "dev-v2.0"
    assert manifest["split_role"] == "development_historical_baseline"
    assert manifest["summary_metrics_frozen"]["oracle_cases"] == 30
    assert manifest["summary_metrics_frozen"]["live_ab_trajectories"] == 50

    expected_paths = {
        "evals/layout_edit_cases.json",
        "evals/run_layout_benchmark.py",
        "evals/layout-benchmark-results.json",
        "evals/LAYOUT-BENCHMARK.md",
    }
    assert set(manifest["frozen_files"].keys()) == expected_paths
    for rel, meta in manifest["frozen_files"].items():
        data = (ROOT / rel).read_bytes()
        assert hashlib.sha256(data).hexdigest() == meta["sha256"], f"Frozen hash mismatch for {rel}"
        assert len(data) == meta["bytes"], f"Frozen byte size mismatch for {rel}"


def test_independent_case_schema_and_commitment_files():
    proto = ROOT / "evals" / "EVALUATION-PROTOCOL.md"
    schema_path = ROOT / "evals" / "schemas" / "independent-case.schema.json"
    commit_path = ROOT / "evals" / "SEALED-TEST-COMMITMENT.json"
    registry_path = ROOT / "evals" / "logs" / "val-run-registry.jsonl"

    assert proto.is_file() and "test-sealed-v1" in proto.read_text(encoding="utf-8")
    schema = json.loads(schema_path.read_text(encoding="utf-8"))
    assert schema["$schema"] == "https://json-schema.org/draft/2020-12/schema"
    assert set(schema["required"]) >= {"case_id", "split", "taxonomy_group", "provenance", "prompt_vi", "grading"}
    assert schema["properties"]["taxonomy_group"]["enum"] == ["A", "B", "C", "D", "E"]

    commitment = json.loads(commit_path.read_text(encoding="utf-8"))
    assert commitment["sealed_split_id"] == "test-sealed-v1"
    assert commitment["target_case_count"] == 30
    assert commitment["status"] in {
        "pending_phase3_curation",
        "committed_awaiting_eval",
        "evaluated_sealed_once",
    }
    if commitment["status"] in {"committed_awaiting_eval", "evaluated_sealed_once"}:
        for key in (
            "bundle_sha256",
            "per_case_manifest_sha256",
            "frozen_system_tag",
            "frozen_system_commit_sha",
            "committed_at",
        ):
            assert isinstance(commitment.get(key), str) and commitment[key]
    if commitment["status"] == "evaluated_sealed_once":
        assert isinstance(commitment.get("evaluated_at"), str) and commitment["evaluated_at"]

    assert registry_path.is_file()
    for line in registry_path.read_text(encoding="utf-8").splitlines():
        if line.strip():
            json.loads(line)


def test_schema_validator_positive_and_negative_mutations(tmp_path):
    doc = Document()
    doc.add_paragraph("Đoạn văn kiểm tra hợp lệ.")
    tbl = doc.add_table(rows=2, cols=2)
    tbl.cell(0, 0).text = "A"
    tbl.cell(0, 1).text = "B"
    valid_docx = tmp_path / "valid.docx"
    doc.save(valid_docx)

    res_ok = validate_package(valid_docx)
    assert res_ok["valid"] is True
    assert res_ok["validator_backends"] == ["standalone_opc", "standalone_ecma376_subset"]

    # 1. ZIP directory entry
    dir_docx = tmp_path / "dir_entry.docx"
    _mutate_zip_part(valid_docx, dir_docx, {}, add_dir_entry="word/")
    res_dir = validate_package(dir_docx)
    assert not res_dir["valid"]
    assert any(e["rule"] == "zip_dir_entries" for e in res_dir["new_errors"])

    # 2. Dangling .rels target
    with zipfile.ZipFile(valid_docx, "r") as zf:
        rels_xml = zf.read("word/_rels/document.xml.rels").decode("utf-8")
    broken_rels = rels_xml.replace(
        "</Relationships>",
        '<Relationship Id="rId999" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer99.xml"/></Relationships>',
    ).encode("utf-8")
    dangling_docx = tmp_path / "dangling_rels.docx"
    _mutate_zip_part(valid_docx, dangling_docx, {"word/_rels/document.xml.rels": broken_rels})
    res_dang = validate_package(dangling_docx)
    assert not res_dang["valid"]
    assert any(e["rule"] == "rels_dangling_target" for e in res_dang["new_errors"])

    # 3. Missing Content_Types coverage
    uncovered_docx = tmp_path / "uncovered.docx"
    _mutate_zip_part(valid_docx, uncovered_docx, {"word/customBlob.weirdext": b"abc"})
    res_ct = validate_package(uncovered_docx)
    assert not res_ct["valid"]
    assert any(e["rule"] == "content_type_coverage" for e in res_ct["new_errors"])

    # 4. Out-of-order tblPr (tblBorders before tblW), pPr after r, invalid tblW type, broken mc:Choice
    with zipfile.ZipFile(valid_docx, "r") as zf:
        doc_xml = zf.read("word/document.xml")
    root = etree.fromstring(doc_xml)
    tbl_pr = root.xpath(".//w:tblPr", namespaces={"w": W})[0]
    tbl_w = tbl_pr.find(f"{{{W}}}tblW")
    tbl_w.set(f"{{{W}}}type", "px")  # invalid type
    tbl_borders = etree.Element(f"{{{W}}}tblBorders")
    tbl_pr.insert(0, tbl_borders)  # out of order before tblW

    bad_p = etree.SubElement(root.find(f"{{{W}}}body"), f"{{{W}}}p")
    etree.SubElement(bad_p, f"{{{W}}}r").append(etree.Element(f"{{{W}}}t"))
    etree.SubElement(bad_p, f"{{{W}}}pPr")  # pPr after w:r

    mc_ns = "http://schemas.openxmlformats.org/markup-compatibility/2006"
    ac = etree.SubElement(root.find(f"{{{W}}}body"), f"{{{mc_ns}}}AlternateContent")
    etree.SubElement(ac, f"{{{mc_ns}}}Choice")  # missing @Requires

    bad_wml_docx = tmp_path / "bad_wml.docx"
    _mutate_zip_part(valid_docx, bad_wml_docx, {"word/document.xml": etree.tostring(root, encoding="UTF-8", xml_declaration=True)})
    res_wml = validate_package(bad_wml_docx)
    rules = {e["rule"] for e in res_wml["new_errors"]}
    assert {"invalid_w_type", "child_order", "pr_after_content", "mc_requires_missing"} <= rules


def test_ooxml_reader_pct_widths_and_style_inheritance(tmp_path):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.0)
    sec.bottom_margin = Cm(2.0)
    sec.left_margin = Cm(3.0)
    sec.right_margin = Cm(1.5)  # printable = 16.5 cm

    doc.add_paragraph("Đoạn văn kiểm tra kế thừa style nhiều tầng cho văn bản tiếng Việt.")
    tbl = doc.add_table(rows=2, cols=3)
    tbl.cell(0, 0).merge(tbl.cell(0, 1))
    tbl.cell(0, 0).text = "Nhóm A"
    tbl.cell(0, 2).text = "Giá trị"
    tbl.cell(1, 0).text = "Cột 1"
    tbl.cell(1, 1).text = "Cột 2"
    tbl.cell(1, 2).text = "100"

    base_docx = tmp_path / "base_real.docx"
    doc.save(base_docx)

    with zipfile.ZipFile(base_docx, "r") as zf:
        doc_root = etree.fromstring(zf.read("word/document.xml"))
        styles_root = etree.fromstring(zf.read("word/styles.xml"))

    # Add two-tier basedOn styles in styles.xml.
    s_base = etree.fromstring(
        f'<w:style xmlns:w="{W}" w:type="paragraph" w:styleId="BodyBase">'
        f'<w:name w:val="BodyBase"/>'
        f'<w:pPr><w:jc w:val="both"/><w:spacing w:after="120" w:line="400" w:lineRule="exact"/></w:pPr>'
        f'<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="28"/></w:rPr>'
        f'</w:style>'
    )
    s_child = etree.fromstring(
        f'<w:style xmlns:w="{W}" w:type="paragraph" w:styleId="PartyInherited">'
        f'<w:name w:val="PartyInherited"/><w:basedOn w:val="BodyBase"/>'
        f'<w:pPr><w:ind w:firstLine="567"/></w:pPr>'
        f'</w:style>'
    )
    styles_root.append(s_base)
    styles_root.append(s_child)

    # Point first paragraph to PartyInherited.
    p_el = doc_root.find(f"{{{W}}}body/{{{W}}}p")
    p_pr = p_el.find(f"{{{W}}}pPr")
    if p_pr is None:
        p_pr = etree.Element(f"{{{W}}}pPr")
        p_el.insert(0, p_pr)
    ps = etree.Element(f"{{{W}}}pStyle")
    ps.set(f"{{{W}}}val", "PartyInherited")
    p_pr.insert(0, ps)

    # Set tblW to 4000 pct (80% of 16.5 cm = 13.2 cm).
    tbl_el = doc_root.find(f"{{{W}}}body/{{{W}}}tbl")
    tbl_pr = tbl_el.find(f"{{{W}}}tblPr")
    tbl_w = tbl_pr.find(f"{{{W}}}tblW")
    tbl_w.set(f"{{{W}}}w", "4000")
    tbl_w.set(f"{{{W}}}type", "pct")

    out_docx = tmp_path / "pct_and_inherit.docx"
    _mutate_zip_part(
        base_docx,
        out_docx,
        {
            "word/document.xml": etree.tostring(doc_root, encoding="UTF-8", xml_declaration=True),
            "word/styles.xml": etree.tostring(styles_root, encoding="UTF-8", xml_declaration=True),
            "_xmlsignatures/sig1.xml": b"<Signature/>",
        },
    )

    info = inspect_docx(out_docx)
    assert info["zip_meta"]["is_signed"] is True
    assert abs(info["tables"][0]["total_width_cm"] - 13.2) < 0.05
    assert info["tables"][0]["width_source"] == "tblW_pct"
    assert info["tables"][0]["has_merged_cells"] is True

    p_fmt = info["paragraphs"][0]["effective_format"]
    assert p_fmt["align"] == "both"
    assert abs(p_fmt["first_line_cm"] - 1.0) < 0.03
    assert p_fmt["line_rule"] == "exact"
    assert abs(p_fmt["line_pt"] - 20.0) < 0.1
    assert abs(p_fmt["space_after_pt"] - 6.0) < 0.1
    assert info["paragraphs"][0]["run_sizes_pt"] == [14.0]
    assert info["paragraphs"][0]["run_fonts"] == ["Times New Roman"]


def _build_table_case_and_src(tmp_path: Path) -> tuple[dict, Path]:
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.0)
    sec.bottom_margin = Cm(2.0)
    sec.left_margin = Cm(3.0)
    sec.right_margin = Cm(1.5)

    p1 = doc.add_paragraph("Đoạn mở đầu không được phép thay đổi nội dung.")
    for r in p1.runs:
        r.font.name = "Times New Roman"
        r.font.size = Pt(14)

    tbl = doc.add_table(rows=2, cols=3)
    tbl.style = "Table Grid"
    for r_idx, vals in enumerate([["TT", "Nội dung", "Ghi chú"], ["1", "Hạng mục mẫu", "Đạt"]]):
        for c_idx, txt in enumerate(vals):
            p_c = tbl.cell(r_idx, c_idx).paragraphs[0]
            run = p_c.add_run(txt)
            run.font.name = "Times New Roman"
            run.font.size = Pt(14)
    _set_table_dxa_widths(tbl, [2.0, 10.5, 4.0], align="left")

    src = tmp_path / "src_tbl.docx"
    doc.save(src)
    case = {
        "case_id": "VAL-A01_table_scale",
        "taxonomy_group": "A",
        "length_tier": "short",
        "instruction_type": "descriptive",
        "expected_outcome": "edit",
        "near_boundary_valid": False,
        "prompt_vi": "Thu nhỏ bảng xuống 80% và căn giữa bảng.",
        "grading": {
            "target_kind": "table",
            "target_ordinal": 1,
            "expected_total_width_cm": 13.2,
            "width_tolerance_cm": 0.35,
            "expected_alignment": "center",
            "max_overflow_right_mm": 0.5,
        },
    }
    return case, src


def test_grader_adversarial_mutations_and_refusals(tmp_path):
    case, src = _build_table_case_and_src(tmp_path)

    # 1. Valid edit (no-render fast check)
    sub_ok = tmp_path / "sub_ok"
    sub_ok.mkdir()
    d_ok = Document(src)
    _set_table_dxa_widths(d_ok.tables[0], [1.6, 8.4, 3.2], align="center")
    d_ok.save(sub_ok / "VAL-A01_table_scale.docx")
    (sub_ok / "VAL-A01_table_scale.result.json").write_text(json.dumps({"status": "edited"}), encoding="utf-8")
    g_ok = grade_case(case, src, sub_ok, run_render=False)
    assert g_ok["hit"] == 1
    assert g_ok["over_refusal"] is False

    # 2. Over-refusal on valid edit
    sub_over = tmp_path / "sub_over"
    sub_over.mkdir()
    (sub_over / "VAL-A01_table_scale.result.json").write_text(
        json.dumps({"status": "refused", "reason": "false alarm"}), encoding="utf-8"
    )
    g_over = grade_case(case, src, sub_over, run_render=False)
    assert g_over["hit"] == 0
    assert g_over["over_refusal"] is True

    # 3. Collateral damage on non-target paragraph text
    sub_col = tmp_path / "sub_col"
    sub_col.mkdir()
    d_col = Document(src)
    d_col.paragraphs[0].text = "Đoạn mở đầu đã bị sửa nhầm!"
    _set_table_dxa_widths(d_col.tables[0], [1.6, 8.4, 3.2], align="center")
    d_col.save(sub_col / "VAL-A01_table_scale.docx")
    (sub_col / "VAL-A01_table_scale.result.json").write_text(json.dumps({"status": "edited"}), encoding="utf-8")
    g_col = grade_case(case, src, sub_col, run_render=False)
    assert g_col["hit"] == 0
    assert g_col["checkpoints"]["semantic_non_target_preserved"] is False

    # 4. Unsafe edit on refuse case
    ref_case = {
        "case_id": "VAL-D01_refuse",
        "taxonomy_group": "D",
        "length_tier": "short",
        "instruction_type": "inferential",
        "expected_outcome": "refuse",
        "near_boundary_valid": False,
        "grading": {"refusal_reason": "Signed document"},
    }
    sub_unsafe = tmp_path / "sub_unsafe"
    sub_unsafe.mkdir()
    d_ok.save(sub_unsafe / "VAL-D01_refuse.docx")
    (sub_unsafe / "VAL-D01_refuse.result.json").write_text(json.dumps({"status": "edited"}), encoding="utf-8")
    g_unsafe = grade_case(ref_case, src, sub_unsafe, run_render=False)
    assert g_unsafe["hit"] == 0
    assert g_unsafe["unsafe_edit"] is True

    # 5. Distorted image aspect ratio
    img_path = _make_png(tmp_path / "diag.png", 600, 300)
    d_img_src = Document()
    d_img_src.add_picture(str(img_path), width=Cm(14.0))
    src_img = tmp_path / "src_img.docx"
    d_img_src.save(src_img)

    d_img_bad = Document()
    p_bad = d_img_bad.add_paragraph()
    p_bad.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_bad.add_run().add_picture(str(img_path), width=Cm(10.0), height=Cm(9.0))  # distorted 2:1 -> 10:9
    sub_img_bad = tmp_path / "sub_img_bad"
    sub_img_bad.mkdir()
    d_img_bad.save(sub_img_bad / "VAL-C01_img.docx")
    img_case = {
        "case_id": "VAL-C01_img",
        "taxonomy_group": "C",
        "length_tier": "short",
        "instruction_type": "descriptive",
        "expected_outcome": "edit",
        "near_boundary_valid": False,
        "grading": {
            "target_kind": "image",
            "target_image_ordinal": 1,
            "expected_image_width_cm": 10.0,
            "image_width_tolerance_cm": 0.25,
            "max_aspect_ratio_delta_pct": 2.0,
            "expected_image_alignment": "center",
        },
    }
    g_img_bad = grade_case(img_case, src_img, sub_img_bad, run_render=False)
    assert g_img_bad["hit"] == 0
    assert g_img_bad["checkpoints"]["ooxml_target_image"] is False

    # 6. Independent compliance profile failure (Party HD05 wrong margins)
    info_nd = inspect_docx(src)
    assert check_compliance_profile(info_nd, "administrative")["status"] == "pass"
    d_bad_party = Document()
    d_bad_party.sections[0].left_margin = Cm(2.0)  # invalid for Party HD05 (requires 3.0 cm)
    d_bad_party.add_paragraph("Đoạn nội dung văn bản Đảng đủ dài để kiểm tra căn lề và giãn dòng theo quy định hiện hành năm 2026.")
    p_party = tmp_path / "bad_party.docx"
    d_bad_party.save(p_party)
    comp_party = check_compliance_profile(inspect_docx(p_party), "party")
    assert comp_party["status"] == "fail"


@pytest.mark.skipif(not find_soffice(), reason="LibreOffice soffice not installed")
def test_grader_render_positive_and_overflow_failure(tmp_path):
    case, src = _build_table_case_and_src(tmp_path)

    # Positive rendered pass
    sub_ok = tmp_path / "render_ok"
    sub_ok.mkdir()
    d_ok = Document(src)
    _set_table_dxa_widths(d_ok.tables[0], [1.6, 8.4, 3.2], align="center")
    d_ok.save(sub_ok / "VAL-A01_table_scale.docx")
    res_ok = grade_case(case, src, sub_ok, run_render=True)
    assert res_ok["hit"] == 1, res_ok
    assert res_ok["renderer_backend"] == "libreoffice"
    assert res_ok["checkpoints"]["pdf_geometry_match"] is True

    # Right-margin overflow failure (table wider than 16.5 cm printable width)
    sub_ovf = tmp_path / "render_ovf"
    sub_ovf.mkdir()
    d_ovf = Document(src)
    _set_table_dxa_widths(d_ovf.tables[0], [2.5, 13.0, 4.5], align="left")  # 20.0 cm > 16.5 cm
    d_ovf.save(sub_ovf / "VAL-A01_table_scale.docx")
    res_ovf = grade_case(case, src, sub_ovf, run_render=True)
    assert res_ovf["hit"] == 0
    assert res_ovf["checkpoints"]["pdf_geometry_match"] is False
    assert res_ovf["telemetry"]["overflow_right_mm_after"] > 5.0


def test_stats_exact_mcnemar_wilcoxon_bootstrap():
    # McNemar exact: b=0, c=8 -> 2 * (1 / 256) = 0.0078125
    mc = mcnemar_exact([0] * 8 + [1, 1], [1] * 8 + [1, 1])
    assert mc["b_a_only"] == 0
    assert mc["c_b_only"] == 8
    assert abs(mc["p_value"] - 0.0078125) < 1e-7

    # Wilcoxon exact (n <= 16)
    wx_exact = wilcoxon_signed_rank([100, 120, 140, 160, 180, 200], [50, 60, 70, 80, 90, 100])
    assert wx_exact["method"] == "exact"
    assert wx_exact["p_value"] == round(2 / 64, 8)
    assert wx_exact["median_delta"] == -75.0

    # Wilcoxon normal approx (n > 16)
    a_big = [100.0 + i for i in range(20)]
    b_big = [60.0 + i for i in range(20)]
    wx_norm = wilcoxon_signed_rank(a_big, b_big)
    assert wx_norm["method"] == "normal_approx"
    assert wx_norm["p_value"] < 0.001

    # Deterministic bootstrap
    ci1 = bootstrap_ci([0.6, 0.8, 1.0, 0.9, 0.7], seed=42)
    ci2 = bootstrap_ci([0.6, 0.8, 1.0, 0.9, 0.7], seed=42)
    assert ci1 == ci2
    assert ci1["ci_low"] <= ci1["mean"] <= ci1["ci_high"]

    # pass@1, pass^k, and ceiling saturation
    hits_by_case = {"C1": [1, 1, 1], "C2": [1, 0, 1], "C3": [0, 0, 0]}
    assert abs(pass_at_1(hits_by_case) - 0.5556) < 0.001
    assert abs(pass_hat_k(hits_by_case, 3) - 0.3333) < 0.001

    ng_sat = normalized_gain(1.0, 1.0)
    assert ng_sat["ceiling_saturated"] is True
    assert ng_sat["g"] is None

    # Arm comparison integration
    rows_a = [
        {"case_id": "C1", "trial_index": 0, "hit": 0, "checkpoint_score": 0.4, "expected_outcome": "edit", "transcript": {"missing": False, "total_tokens": 2000, "fresh_plus_output_tokens": 1000, "tool_calls_total": 10, "elapsed_seconds": 30.0}},
        {"case_id": "C2", "trial_index": 0, "hit": 1, "checkpoint_score": 1.0, "expected_outcome": "edit", "transcript": {"missing": False, "total_tokens": 1800, "fresh_plus_output_tokens": 900, "tool_calls_total": 8, "elapsed_seconds": 25.0}},
    ]
    rows_b = [
        {"case_id": "C1", "trial_index": 0, "hit": 1, "checkpoint_score": 1.0, "expected_outcome": "edit", "transcript": {"missing": False, "total_tokens": 800, "fresh_plus_output_tokens": 400, "tool_calls_total": 4, "elapsed_seconds": 12.0}},
        {"case_id": "C2", "trial_index": 0, "hit": 1, "checkpoint_score": 1.0, "expected_outcome": "edit", "transcript": {"missing": False, "total_tokens": 700, "fresh_plus_output_tokens": 350, "tool_calls_total": 3, "elapsed_seconds": 10.0}},
    ]
    cmp_res = compare_arms(rows_a, rows_b)
    assert cmp_res["hit_rate_pp_delta"] == 50.0
    assert cmp_res["normalized_gain"]["ceiling_saturated"] is False
    assert cmp_res["normalized_gain"]["g"] == 1.0


def test_curator_zero_leakage_disjoint_and_seal_commitment(tmp_path: Path) -> None:
    """Verify zero-leakage prompt linter, split disjoint checker, and commitment sealer."""
    clean_prompt = (
        "Hãy căn giữa bảng kế hoạch gồm 5 cột, giữ nguyên lề trang A4 và không sửa phần tiêu đề."
    )
    assert lint_prompt_zero_leakage(clean_prompt) == []

    leaky_prompt = "Dùng vndocs layout-edit sửa w:tblPr theo đơn vị dxa cho bảng T#1."
    hits = lint_prompt_zero_leakage(leaky_prompt)
    for expected_flag in ("vndocs", "layout-edit", "w:xml_tag", "dxa", "anchor_id"):
        assert any(expected_flag.lower() in h.lower() for h in hits), (expected_flag, hits)

    val_manifest = ROOT / "evals" / "val_v1" / "val_v1_cases.json"
    # Overlap detection
    overlap_manifest = tmp_path / "sealed_overlap.json"
    overlap_manifest.write_text(
        json.dumps(
            {
                "schema_version": "1.0",
                "split": "test-sealed-v1",
                "cases": [
                    {
                        "case_id": "SEAL-A01_dup",
                        "provenance": {"source_cluster_id": "ttn_edu_vn"},
                        "source_docx_sha256": "97936209270f7abe36ea9b7eeadb5e2fc010f7dd96bc5f969c261ce0b0edd17d",
                    }
                ],
            }
        ),
        encoding="utf-8",
    )
    disj_bad = check_disjoint(val_manifest, overlap_manifest)
    assert disj_bad["disjoint"] is False
    assert "ttn_edu_vn" in disj_bad["overlapping_clusters"]
    assert "97936209270f7abe36ea9b7eeadb5e2fc010f7dd96bc5f969c261ce0b0edd17d" in disj_bad["overlapping_hashes"]

    # Disjoint + 30-case synthetic sealed bundle commitment sealing
    sealed_dir = tmp_path / "sealed_pkg"
    (sealed_dir / "corpus").mkdir(parents=True)
    sealed_cases = []
    tiers = ["short"] * 8 + ["medium"] * 9 + ["long"] * 7 + ["very_long"] * 6
    idx = 0
    for grp in ("A", "B", "C", "D", "E"):
        for num in range(1, 7):
            cid = f"SEAL-{grp}{num:02d}_agency"
            f_path = sealed_dir / "corpus" / f"{cid}.docx"
            d_tmp = Document()
            d_tmp.add_paragraph(f"Tài liệu kiểm thử niêm phong {cid}.")
            d_tmp.save(str(f_path))
            f_sha = hashlib.sha256(f_path.read_bytes()).hexdigest()
            is_refuse = grp == "D" and num <= 5
            sealed_cases.append(
                {
                    "schema_version": "1.0",
                    "case_id": cid,
                    "split": "test-sealed-v1",
                    "taxonomy_group": grp,
                    "length_tier": tiers[idx],
                    "instruction_type": "descriptive" if idx % 2 == 0 else "inferential",
                    "expected_outcome": "refuse" if is_refuse else "edit",
                    "near_boundary_valid": (grp == "D" and num == 6) or (grp == "B" and num <= 2),
                    "out_of_scope_for_vndocs": grp == "E",
                    "provenance": {
                        "source_category": "gov_portal_nd30",
                        "authoring_software": "microsoft_word",
                        "source_cluster_id": f"unique_sealed_portal_{idx:02d}",
                        "license_or_status": "public_gov_portal",
                        "pii_redacted": False,
                        "created_without_vndocs_builders": True,
                    },
                    "source_docx_relpath": f"corpus/{cid}.docx",
                    "source_docx_sha256": f_sha,
                    "attachments": [],
                    "prompt_vi": f"Hãy căn chỉnh bảng tổng hợp số {idx + 1} trong tài liệu công vụ này cho cân đối trên trang A4.",
                    "grading": {"target_kind": "table", "target_ordinal": 1},
                }
            )
            idx += 1

    sealed_manifest = sealed_dir / "test_sealed_v1_cases.json"
    sealed_manifest.write_text(
        json.dumps({"schema_version": "1.0", "split": "test-sealed-v1", "cases": sealed_cases}, ensure_ascii=False),
        encoding="utf-8",
    )
    disj_ok = check_disjoint(val_manifest, sealed_manifest)
    assert disj_ok["disjoint"] is True
    assert disj_ok["overlapping_clusters"] == []
    assert disj_ok["overlapping_hashes"] == []

    commit_out = tmp_path / "commitment.json"
    commit = seal_commitment(
        sealed_dir,
        "test_sealed_v1_cases.json",
        val_manifest,
        commit_out,
        frozen_system_tag="eval-freeze-test",
        frozen_system_commit_sha="0" * 40,
    )
    assert commit["status"] == "committed_awaiting_eval"
    assert commit["target_case_count"] == 30
    assert commit["taxonomy_quota"] == {"A": 6, "B": 6, "C": 6, "D": 6, "E": 6}
    assert len(commit["bundle_sha256"]) == 64
    assert len(commit["per_case_manifest_sha256"]) == 64
    assert commit_out.is_file()


def test_val_v1_manifest_conforms_to_protocol_and_zero_leakage() -> None:
    """Verify public val-v1 manifest meets protocol quotas, zero-leakage, and solvability flags."""
    manifest_path = ROOT / "evals" / "val_v1" / "val_v1_cases.json"
    local_corpus = ROOT / "research" / "eval" / "val_v1" / "corpus"
    corpus_arg = local_corpus if local_corpus.is_dir() else None

    res = validate_manifest(manifest_path, corpus_dir=corpus_arg, enforce_quotas=True)
    assert res["valid"] is True, res["errors"]
    assert res["case_count"] == 20
    assert res["group_counts"] == {"A": 4, "B": 4, "C": 4, "D": 4, "E": 4}
    assert res["length_counts"] == {"short": 5, "medium": 6, "long": 5, "very_long": 4}
    assert res["instruction_type_counts"]["descriptive"] >= 8
    assert res["instruction_type_counts"]["inferential"] >= 8
    assert res["unique_source_clusters"] == 20
    assert res["near_boundary_valid_count"] >= 2
    assert 3 <= res["outcome_counts"]["refuse"] <= 5

    raw = json.loads(manifest_path.read_text(encoding="utf-8"))
    for case in raw["cases"]:
        sol = case.get("independent_solvability_verified", {})
        assert sol.get("positive_non_vndocs_verified") is True, case["case_id"]
        assert sol.get("negative_mutation_verified") is True, case["case_id"]
        assert lint_prompt_zero_leakage(case["prompt_vi"]) == [], case["case_id"]


def test_grader_out_of_scope_group_e_checks(tmp_path: Path) -> None:
    """Verify Group E out-of-scope checks (footnotes, footer label, nested table, hanging indent)."""
    # 1. Footnote font & size check
    fn_raw = tmp_path / "fn_raw.docx"
    fn_src = tmp_path / "fn_src.docx"
    doc_fn = Document()
    doc_fn.add_paragraph("Đoạn chính có chú thích.")
    doc_fn.save(str(fn_raw))

    fn_rels_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
    fn_ct = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"

    def _write_fn_pkg(src: Path, dst: Path, font_name: str, sz_half: str) -> None:
        footnotes_xml = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:footnotes xmlns:w="{W}">'
            f'<w:footnote w:id="1"><w:p><w:r>'
            f'<w:rPr><w:rFonts w:ascii="{font_name}" w:hAnsi="{font_name}"/>'
            f'<w:sz w:val="{sz_half}"/><w:szCs w:val="{sz_half}"/></w:rPr>'
            f'<w:t>Chú thích mẫu</w:t></w:r></w:p></w:footnote>'
            f'</w:footnotes>'
        ).encode("utf-8")
        with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(dst, "w") as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "[Content_Types].xml":
                    r = etree.fromstring(data)
                    ns = r.nsmap[None]
                    etree.SubElement(r, f"{{{ns}}}Override", PartName="/word/footnotes.xml", ContentType=fn_ct)
                    data = etree.tostring(r, xml_declaration=True, encoding="UTF-8", standalone=True)
                elif item.filename == "word/_rels/document.xml.rels":
                    r = etree.fromstring(data)
                    ns = r.nsmap[None]
                    etree.SubElement(r, f"{{{ns}}}Relationship", Id="rIdFn1", Type=fn_rels_type, Target="footnotes.xml")
                    data = etree.tostring(r, xml_declaration=True, encoding="UTF-8", standalone=True)
                zout.writestr(item, data)
            zout.writestr("word/footnotes.xml", footnotes_xml)

    _write_fn_pkg(fn_raw, fn_src, "Arial", "20")
    case_fn = {
        "case_id": "VAL-E01_test",
        "expected_outcome": "edit",
        "grading": {
            "target_kind": "out_of_scope",
            "expected_footnote_font_name": "Times New Roman",
            "expected_footnote_font_pt": 9.0,
        },
    }
    pos_fn = tmp_path / "pos_fn"
    pos_fn.mkdir()
    _write_fn_pkg(fn_raw, pos_fn / "VAL-E01_test.docx", "Times New Roman", "18")
    assert grade_case(case_fn, fn_src, pos_fn, run_render=False)["hit"] == 1
    neg_fn = tmp_path / "neg_fn"
    neg_fn.mkdir()
    _write_fn_pkg(fn_raw, neg_fn / "VAL-E01_test.docx", "Arial", "20")
    assert grade_case(case_fn, fn_src, neg_fn, run_render=False)["hit"] == 0

    # 2. Footer text check
    ftr_src = tmp_path / "ftr_src.docx"
    d_ftr = Document()
    d_ftr.add_paragraph("Nội dung.")
    d_ftr.sections[0].footer.paragraphs[0].text = "1"
    d_ftr.save(str(ftr_src))
    case_ftr = {
        "case_id": "VAL-E02_test",
        "expected_outcome": "edit",
        "grading": {
            "target_kind": "out_of_scope",
            "target_footer_part": "word/footer1.xml",
            "expected_footer_text_contains": ["Trang"],
        },
    }
    pos_ftr = tmp_path / "pos_ftr"
    pos_ftr.mkdir()
    d_ftr_pos = Document(str(ftr_src))
    d_ftr_pos.sections[0].footer.paragraphs[0].text = "Trang 1"
    d_ftr_pos.save(str(pos_ftr / "VAL-E02_test.docx"))
    assert grade_case(case_ftr, ftr_src, pos_ftr, run_render=False)["hit"] == 1

    neg_ftr = tmp_path / "neg_ftr"
    neg_ftr.mkdir()
    Document(str(ftr_src)).save(str(neg_ftr / "VAL-E02_test.docx"))
    assert grade_case(case_ftr, ftr_src, neg_ftr, run_render=False)["hit"] == 0

    # 3. Nested table alignment/shading + hanging indent check
    nt_src = tmp_path / "nt_src.docx"
    d_nt = Document()
    d_nt.add_paragraph("- Mục 1 thử nghiệm")
    outer = d_nt.add_table(rows=1, cols=1)
    outer.cell(0, 0).add_table(rows=2, cols=2)
    d_nt.save(str(nt_src))

    case_nt = {
        "case_id": "VAL-E03_test",
        "expected_outcome": "edit",
        "grading": {
            "target_kind": "out_of_scope",
            "target_nested_table_ordinal": 1,
            "expected_nested_table_alignment": "center",
            "expected_nested_table_shading_fill": "D9E2F3",
            "target_paragraph_ordinals": [1],
            "expected_hanging_indent_cm_range": [0.45, 0.75],
        },
    }

    def _mutate_nt_and_hang(src: Path, dst: Path, align: str, fill: str | None, hang_cm: float | None) -> None:
        with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(dst, "w") as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    root = etree.fromstring(data)
                    if hang_cm is not None:
                        p = root.find(f".//{{{W}}}body/{{{W}}}p")
                        pPr = p.find(f"{{{W}}}pPr")
                        if pPr is None:
                            pPr = etree.Element(f"{{{W}}}pPr")
                            p.insert(0, pPr)
                        ind = etree.SubElement(pPr, f"{{{W}}}ind")
                        ind.set(f"{{{W}}}left", str(int(round(hang_cm * DXA_PER_CM))))
                        ind.set(f"{{{W}}}hanging", str(int(round(hang_cm * DXA_PER_CM))))
                    nst = root.xpath(".//w:tbl//w:tbl", namespaces={"w": W})[0]
                    tblPr = nst.find(f"{{{W}}}tblPr")
                    if tblPr is None:
                        tblPr = etree.Element(f"{{{W}}}tblPr")
                        nst.insert(0, tblPr)
                    jc = etree.Element(f"{{{W}}}jc")
                    jc.set(f"{{{W}}}val", align)
                    look = tblPr.find(f"{{{W}}}tblLook")
                    if look is not None:
                        tblPr.insert(list(tblPr).index(look), jc)
                    else:
                        tblPr.append(jc)
                    if fill:
                        tr0 = nst.find(f"{{{W}}}tr")
                        for tc in tr0.findall(f"{{{W}}}tc"):
                            tcPr = tc.find(f"{{{W}}}tcPr")
                            if tcPr is None:
                                tcPr = etree.Element(f"{{{W}}}tcPr")
                                tc.insert(0, tcPr)
                            shd = etree.SubElement(tcPr, f"{{{W}}}shd")
                            shd.set(f"{{{W}}}val", "clear")
                            shd.set(f"{{{W}}}color", "auto")
                            shd.set(f"{{{W}}}fill", fill)
                    data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                zout.writestr(item, data)

    pos_nt = tmp_path / "pos_nt"
    pos_nt.mkdir()
    _mutate_nt_and_hang(nt_src, pos_nt / "VAL-E03_test.docx", "center", "D9E2F3", 0.55)
    assert grade_case(case_nt, nt_src, pos_nt, run_render=False)["hit"] == 1

    neg_nt = tmp_path / "neg_nt"
    neg_nt.mkdir()
    _mutate_nt_and_hang(nt_src, neg_nt / "VAL-E03_test.docx", "left", None, 0.0)
    assert grade_case(case_nt, nt_src, neg_nt, run_render=False)["hit"] == 0


def test_arm1_standards_guide_has_zero_leakage() -> None:
    """Verify Arm 1 domain guide stays strictly free of internal vndocs/CLI/XML tokens."""
    guide_path = ROOT / "evals" / "prompts" / "arm1_standards_only.md"
    assert guide_path.is_file()
    text = guide_path.read_text(encoding="utf-8")
    assert len(text) > 500
    assert lint_prompt_zero_leakage(text) == []
    assert "scripts/" not in text
    assert "SKILL.md" not in text


def test_runner_workspace_isolation_telemetry_and_contamination(tmp_path: Path) -> None:
    """Verify 3-arm workspace isolation, Codex JSONL parsing, and contamination scanning."""
    corpus_dir = tmp_path / "corpus"
    (corpus_dir / "attachments").mkdir(parents=True)
    src_doc = Document()
    src_doc.add_paragraph("Tài liệu thử nghiệm")
    src_doc.save(str(corpus_dir / "SAMPLE.docx"))
    _make_png(corpus_dir / "attachments" / "shot.png", width=200, height=100)

    case = {
        "case_id": "VAL-A01_sample",
        "taxonomy_group": "A",
        "length_tier": "short",
        "instruction_type": "descriptive",
        "expected_outcome": "edit",
        "source_docx_relpath": "SAMPLE.docx",
        "attachments": [{"kind": "screenshot_png", "relpath": "attachments/shot.png"}],
        "prompt_vi": "Hãy căn giữa bảng đầu tiên trong tài liệu.",
        "grading": {"target_kind": "table", "target_ordinal": 1},
    }

    w0 = prepare_arm_workspace(case, corpus_dir, tmp_path / "arm0", ARM_NO_SKILL, repo_root=ROOT)
    w1 = prepare_arm_workspace(case, corpus_dir, tmp_path / "arm1", ARM_STANDARDS_ONLY, repo_root=ROOT)
    w2 = prepare_arm_workspace(case, corpus_dir, tmp_path / "arm2", ARM_FULL_SKILL, repo_root=ROOT)

    assert sorted(p.name for p in Path(w0["work_dir"]).iterdir()) == ["shot.png", "src.docx"]
    assert sorted(p.name for p in Path(w1["work_dir"]).iterdir()) == ["STANDARDS_GUIDE.md", "shot.png", "src.docx"]
    assert not (Path(w1["work_dir"]) / "skill").exists()
    assert (Path(w2["work_dir"]) / "skill" / "SKILL.md").is_file()
    assert (Path(w2["work_dir"]) / "skill" / "scripts" / "cli.py").is_file()

    jsonl_file = tmp_path / "transcript.jsonl"
    jsonl_file.write_text(
        "\n".join(
            [
                json.dumps({"type": "thread.started", "model": "gpt-6.1-sol"}),
                json.dumps({"type": "item.completed", "item": {"type": "command_execution", "command": "python edit.py"}}),
                json.dumps({"type": "item.completed", "item": {"type": "file_change", "path": "VAL-A01_sample.docx"}}),
                json.dumps({"type": "turn.completed", "usage": {"input_tokens": 5000, "cached_input_tokens": 3500, "output_tokens": 400}}),
                json.dumps({"type": "turn.completed", "usage": {"input_tokens": 2000, "cached_input_tokens": 1000, "output_tokens": 250}}),
            ]
        )
        + "\n",
        encoding="utf-8",
    )
    tel = parse_codex_jsonl(jsonl_file, elapsed_seconds=12.345)
    assert tel["missing"] is False
    assert tel["model_id"] == "gpt-6.1-sol"
    assert tel["input_tokens"] == 7000
    assert tel["cached_input_tokens"] == 4500
    assert tel["output_tokens"] == 650
    assert tel["fresh_plus_output_tokens"] == (7000 - 4500) + 650
    assert tel["total_tokens"] == 7650
    assert tel["tool_calls_total"] == 2
    assert tel["elapsed_seconds"] == 12.35

    assert scan_transcript_contamination(jsonl_file, ARM_NO_SKILL)["contaminated"] is False
    assert scan_transcript_contamination(jsonl_file, ARM_STANDARDS_ONLY)["contaminated"] is False
    assert scan_transcript_contamination(jsonl_file, ARM_FULL_SKILL)["checked"] is False

    leaky_jsonl = tmp_path / "leaky.jsonl"
    leaky_jsonl.write_text('{"command": "python scripts/cli.py map src.docx && vndocs help"}\n', encoding="utf-8")
    c0 = scan_transcript_contamination(leaky_jsonl, ARM_NO_SKILL)
    c1 = scan_transcript_contamination(leaky_jsonl, ARM_STANDARDS_ONLY)
    assert c0["contaminated"] is True and c0["hit_count"] >= 2
    assert c1["contaminated"] is True
    assert scan_transcript_contamination(leaky_jsonl, ARM_FULL_SKILL)["contaminated"] is False


def test_runner_evaluate_and_val_registry_cap(tmp_path: Path) -> None:
    """Verify 3-arm offline experiment evaluation, Markdown summary, and 3-milestone registry cap."""
    corpus_dir = tmp_path / "corpus"
    corpus_dir.mkdir()
    src = corpus_dir / "REF.docx"
    d = Document()
    d.add_paragraph("Văn bản Đảng mẫu.")
    d.save(str(src))

    case = {
        "case_id": "VAL-D01_ref",
        "taxonomy_group": "D",
        "length_tier": "short",
        "instruction_type": "descriptive",
        "expected_outcome": "refuse",
        "source_docx_relpath": "REF.docx",
        "prompt_vi": "Đổi sang NĐ30",
        "grading": {"target_kind": "section"},
    }

    records = []
    for arm_idx, arm in enumerate(ALL_ARMS):
        wdir = tmp_path / arm / "VAL-D01_ref"
        prepare_arm_workspace(case, corpus_dir, wdir, arm, repo_root=ROOT)
        if arm == ARM_NO_SKILL:
            # Unsafe edit -> fail
            d_bad = Document(str(src))
            d_bad.save(str(wdir / "VAL-D01_ref.docx"))
            (wdir / "VAL-D01_ref.result.json").write_text(
                json.dumps({"case_id": "VAL-D01_ref", "status": "edited"}), encoding="utf-8"
            )
        else:
            (wdir / "VAL-D01_ref.result.json").write_text(
                json.dumps({"case_id": "VAL-D01_ref", "status": "refused", "reason": "Conflict"}), encoding="utf-8"
            )
        tr_file = wdir / "transcript.jsonl"
        tr_file.write_text(
            json.dumps({"type": "turn.completed", "usage": {"input_tokens": 1000 * (arm_idx + 1), "cached_input_tokens": 200, "output_tokens": 100}})
            + "\n",
            encoding="utf-8",
        )
        records.append(
            {
                "case_id": "VAL-D01_ref",
                "arm": arm,
                "trial_index": 0,
                "model_id": "gpt-6.1-sol",
                "work_dir": str(wdir),
                "transcript": parse_codex_jsonl(tr_file, elapsed_seconds=5.0),
            }
        )

    report = evaluate_experiment([case], records, corpus_dir, run_render=False, repo_root=ROOT)
    assert report["n_total_runs"] == 3
    assert report["contamination_count"] == 0
    assert report["summary"]["by_arm"][ARM_NO_SKILL]["pass_at_1"] == 0.0
    assert report["summary"]["by_arm"][ARM_STANDARDS_ONLY]["pass_at_1"] == 1.0
    assert report["summary"]["by_arm"][ARM_FULL_SKILL]["pass_at_1"] == 1.0
    md = render_experiment_markdown(report)
    assert "arm0_no_skill" in md and "arm2_full_skill" in md

    reg_path = tmp_path / "val-run-registry.jsonl"
    base_rec = {
        "timestamp": "2026-11-24T17:00:00+07:00",
        "git_commit_sha": "abc123",
        "skill_md_sha256": "0" * 64,
        "model_ids": ["gpt-6.1-sol"],
        "arms": list(ALL_ARMS),
        "trials_per_case": 1,
        "summary": {"pass_at_1": 1.0},
    }
    append_val_registry(reg_path, {**base_rec, "run_id": "m1", "milestone_index": 1})
    with pytest.raises(ValueError, match="Duplicate run_id"):
        append_val_registry(reg_path, {**base_rec, "run_id": "m1", "milestone_index": 2})
    append_val_registry(reg_path, {**base_rec, "run_id": "m2", "milestone_index": 2})
    append_val_registry(reg_path, {**base_rec, "run_id": "m3", "milestone_index": 3})
    with pytest.raises(ValueError, match="budget exhausted|outside"):
        append_val_registry(reg_path, {**base_rec, "run_id": "m4", "milestone_index": 3})

    # Real Antigravity PLANNER_RESPONSE transcript parsing
    ag_tr = tmp_path / "ag_transcript.jsonl"
    ag_tr.write_text(
        "\n".join(
            [
                json.dumps({"step_index": 0, "type": "SYSTEM_MESSAGE", "created_at": "2026-11-24T10:00:00Z"}),
                json.dumps({
                    "step_index": 1,
                    "type": "PLANNER_RESPONSE",
                    "created_at": "2026-11-24T10:00:02Z",
                    "model": "gemini-4-hydrogen-tiered-high",
                    "input_tokens": 5000,
                    "cache_read_tokens": 1000,
                    "output_tokens": 300,
                    "tool_calls": [{"name": "run_command"}],
                }),
                json.dumps({
                    "step_index": 2,
                    "type": "PLANNER_RESPONSE",
                    "created_at": "2026-11-24T10:00:12Z",
                    "model": "gemini-4-hydrogen-tiered-high",
                    "input_tokens": 2000,
                    "cache_read_tokens": 5000,
                    "output_tokens": 200,
                    "tool_calls": [{"name": "read_file"}, {"name": "run_command"}],
                }),
            ]
        )
        + "\n",
        encoding="utf-8",
    )
    ag_tel = parse_antigravity_transcript(ag_tr)
    assert ag_tel["missing"] is False
    assert ag_tel["model_id"] == "gemini-4-hydrogen-tiered-high"
    assert ag_tel["fresh_plus_output_tokens"] == 7500
    assert ag_tel["total_tokens"] == 13500
    assert ag_tel["tool_calls_total"] == 3
    assert ag_tel["elapsed_seconds"] == 12.0

    # Resume helper reloads clean prior Codex run metadata
    res_dir = tmp_path / "resume_case"
    res_dir.mkdir()
    (res_dir / "VAL-D01_ref.result.json").write_text('{"status":"refused","reason":"ok"}', encoding="utf-8")
    (res_dir / "transcript.jsonl").write_text(
        json.dumps({"type": "turn.completed", "usage": {"input_tokens": 900, "cached_input_tokens": 200, "output_tokens": 80}}) + "\n",
        encoding="utf-8",
    )
    (res_dir / "run_meta.json").write_text(
        json.dumps({"returncode": 0, "timed_out": False, "model_id": "gpt-6.1-sol", "elapsed_seconds": 6.4}),
        encoding="utf-8",
    )
    resumed = _try_resume_codex_case(case, ARM_NO_SKILL, res_dir)
    assert resumed is not None and resumed["resumed"] is True
    assert resumed["transcript"]["total_tokens"] == 980

    # Antigravity SYSTEM_MESSAGE skipped by contamination scanner, PLANNER_RESPONSE still checked
    ag_clean_tr = tmp_path / "ag_sys_only.jsonl"
    ag_clean_tr.write_text(
        "\n".join([
            json.dumps({"type": "SYSTEM_MESSAGE", "content": "AGENTS.md says SKILL.md vndocs"}),
            json.dumps({"type": "PLANNER_RESPONSE", "tool_calls": [{"name": "run_command", "args": {"CommandLine": "python fix.py"}}]}),
        ]) + "\n",
        encoding="utf-8",
    )
    assert scan_transcript_contamination(ag_clean_tr, ARM_NO_SKILL)["contaminated"] is False

    ag_dirty_tr = tmp_path / "ag_dirty.jsonl"
    ag_dirty_tr.write_text(
        "\n".join([
            json.dumps({"type": "SYSTEM_MESSAGE", "content": "ok"}),
            json.dumps({"type": "PLANNER_RESPONSE", "tool_calls": [{"name": "read_file", "args": {"Path": "vietnamese-docs-style/SKILL.md"}}]}),
        ]) + "\n",
        encoding="utf-8",
    )
    assert scan_transcript_contamination(ag_dirty_tr, ARM_NO_SKILL)["contaminated"] is True

    q_tr = tmp_path / "q_tr.jsonl"
    q_err = tmp_path / "q_err.log"
    q_tr.write_text('{"type":"error","message":"You’ve hit your usage limit. Upgrade to Pro."}\n', encoding="utf-8")
    q_err.write_text("", encoding="utf-8")
    assert _is_codex_quota_error(q_tr, q_err) is True


def test_diff_ntpr_allows_unlinked_footer_copy_preserving_original_text():
    from lxml import etree

    from eval_harness.diff_ntpr import check_non_target_preservation

    empty_doc = etree.tostring(
        etree.fromstring(b'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>')
    )
    base_info = {
        "paragraphs": [],
        "tables": [],
        "images": [],
        "sections": [],
        "notes_text_nfc": {},
        "headers_footers": {
            "word/footer1.xml": {"text_nfc": "Cover"},
            "word/footer2.xml": {"text_nfc": "i"},
        },
    }
    valid_after = {
        **base_info,
        "headers_footers": {
            "word/footer1.xml": {"text_nfc": "Cover"},
            "word/footer2.xml": {"text_nfc": "Trang i"},
            "word/footer3.xml": {"text_nfc": "i"},
        },
    }
    res_ok = check_non_target_preservation(
        base_info,
        valid_after,
        {"word/document.xml": empty_doc},
        {"word/document.xml": empty_doc},
        target_kind="out_of_scope",
        allow_header_footer_parts={"word/footer2.xml"},
    )
    assert res_ok["details"]["hf_text_ok"] is True

    bad_after = {
        **base_info,
        "headers_footers": {
            "word/footer1.xml": {"text_nfc": "Cover"},
            "word/footer2.xml": {"text_nfc": "Trang i"},
            "word/footer3.xml": {"text_nfc": "MODIFIED"},
        },
    }
    res_bad = check_non_target_preservation(
        base_info,
        bad_after,
        {"word/document.xml": empty_doc},
        {"word/document.xml": empty_doc},
        target_kind="out_of_scope",
        allow_header_footer_parts={"word/footer2.xml"},
    )
    assert res_bad["details"]["hf_text_ok"] is False


def test_sealed_verify_mark_and_blind_audit_from_report(tmp_path: Path):
    import json
    from unittest.mock import patch

    import pytest

    from eval_harness.curator import (
        mark_commitment_evaluated,
        prepare_blind_audit_from_report,
        seal_commitment,
        verify_commitment,
    )
    from eval_harness.stats import bootstrap_ci  # sanity ensure stats available

    assert bootstrap_ci([1.0, 1.0])["mean"] == 1.0

    # Minimal synthetic val manifest + valid disjoint 30-case sealed bundle using quota mock
    val_manifest = tmp_path / "val.json"
    val_manifest.write_text(json.dumps({"cases": [{"case_id": "VAL-A01_x", "provenance": {"source_cluster_id": "v_1"}, "source_docx_sha256": "abc"}]}), encoding="utf-8")

    sroot = tmp_path / "sealed"
    (sroot / "corpus").mkdir(parents=True)
    docx_f = sroot / "corpus" / "SEAL-A01_demo.docx"
    docx_f.write_bytes(b"PK-fake-docx")
    s_manifest = sroot / "test_sealed_v1_cases.json"
    s_manifest.write_text(
        json.dumps({
            "cases": [{
                "case_id": "SEAL-A01_demo",
                "split": "test-sealed-v1",
                "taxonomy_group": "A",
                "length_tier": "short",
                "instruction_type": "descriptive",
                "expected_outcome": "edit",
                "source_docx_relpath": "corpus/SEAL-A01_demo.docx",
                "source_docx_sha256": "def",
                "prompt_vi": "Chinh bang",
                "provenance": {"source_cluster_id": "seal_1"},
            }]
        }),
        encoding="utf-8",
    )
    comm_path = tmp_path / "SEALED-TEST-COMMITMENT.json"

    with patch("eval_harness.curator.validate_manifest", return_value={"valid": True, "group_counts": {"A": 6, "B": 6, "C": 6, "D": 6, "E": 6}, "errors": []}):
        comm = seal_commitment(sroot, "test_sealed_v1_cases.json", val_manifest, comm_path, frozen_system_tag="eval-freeze-v1")
    assert comm["status"] == "committed_awaiting_eval"

    ok_verify = verify_commitment(comm_path, sroot, "test_sealed_v1_cases.json")
    assert ok_verify["valid"] is True and ok_verify["errors"] == []

    # Tamper detection
    docx_f.write_bytes(b"PK-tampered")
    bad_verify = verify_commitment(comm_path, sroot, "test_sealed_v1_cases.json")
    assert bad_verify["valid"] is False
    assert any("bundle_sha256 mismatch" in e for e in bad_verify["errors"])
    docx_f.write_bytes(b"PK-fake-docx")

    # Single-use lock check
    marked = mark_commitment_evaluated(comm_path, notes="single shot")
    assert marked["status"] == "evaluated_sealed_once"
    assert verify_commitment(comm_path, sroot, "test_sealed_v1_cases.json")["valid"] is False
    with pytest.raises(ValueError, match="already evaluated"):
        mark_commitment_evaluated(comm_path)

    # prepare_blind_audit_from_report strips all arm/model identifiers from scoring_sheet.csv
    run_wdir = tmp_path / "runs" / "arm2_full_skill" / "SEAL-A01_demo"
    run_wdir.mkdir(parents=True)
    sub_docx = run_wdir / "SEAL-A01_demo.docx"
    sub_docx.write_bytes(b"PK-after")
    runs_mf = tmp_path / "runs_manifest.json"
    runs_mf.write_text(json.dumps({"runs": [{"case_id": "SEAL-A01_demo", "arm": "arm2_full_skill", "work_dir": str(run_wdir)}]}), encoding="utf-8")
    rep_path = tmp_path / "benchmark_report.json"
    rep_path.write_text(json.dumps({"rows": [{"case_id": "SEAL-A01_demo", "arm": "arm2_full_skill", "model_id": "test-model"}]}), encoding="utf-8")

    blind_out = tmp_path / "blind_out"
    with patch("eval_harness.curator.render_to_pdf", return_value={"status": "not_rendered"}):
        res_blind = prepare_blind_audit_from_report(
            rep_path,
            s_manifest,
            sroot,
            blind_out,
            runs_manifest_path=runs_mf,
            max_cases=5,
            seed=42,
        )
    assert res_blind["item_count"] == 1
    csv_txt = (blind_out / "scoring_sheet.csv").read_text(encoding="utf-8-sig")
    assert "BLIND-001" in csv_txt
    assert "arm2_full_skill" not in csv_txt
    assert "test-model" not in csv_txt
    viewer_f = blind_out / "viewer.html"
    assert viewer_f.is_file()
    viewer_txt = viewer_f.read_text(encoding="utf-8")
    assert "BLIND-001" in viewer_txt
    assert "arm2_full_skill" not in viewer_txt
    assert "test-model" not in viewer_txt
