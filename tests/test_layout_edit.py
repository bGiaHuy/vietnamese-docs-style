"""Unit and LibreOffice integration tests for layout mapping, editing, and diffing modules."""
import json
import unicodedata
import zipfile
from pathlib import Path
from unittest.mock import patch

import pytest
from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION_START
from docx.shared import Cm, Pt
from lxml import etree

import cli
import diff_layout
import layout_edit
import map_docx
import ooxml_order
from build_docx import add_body, add_table, setup_document
from docx_package import W, read_package
from render_docx import find_renderer

ROOT = Path(__file__).resolve().parent.parent
ASSETS_TEMPLATES = ROOT / 'assets' / 'templates'


def _save(tmp_path, doc, name='sample.docx'):
    out = tmp_path / name
    doc.save(out)
    return out


def _apply_test_col_widths(tbl, col_widths_cm):
    tbl.autofit = False
    for idx, w_cm in enumerate(col_widths_cm):
        width = Cm(w_cm)
        tbl.columns[idx].width = width
        for row in tbl.rows:
            row.cells[idx].width = width
    # Ensure OOXML gridCol matches requested test widths deterministically
    tbl_grid = tbl._tbl.find(f'{{{W}}}tblGrid')
    if tbl_grid is not None:
        grid_cols = tbl_grid.findall(f'{{{W}}}gridCol')
        for gcol, w_cm in zip(grid_cols, col_widths_cm):
            gcol.set(f'{{{W}}}w', str(round(w_cm * map_docx.DXA_PER_CM)))
    return tbl


def _make_simple_table_doc(
    tmp_path,
    name='table.docx',
    *,
    profile='academic',
    rows=None,
    col_widths_cm=None,
    table_text_before='Nội dung mở đầu báo cáo.',
    table_text_after='Phần kết luận văn bản.',
    landscape=False,
    multisection=False,
):
    doc = Document()
    setup_document(doc, profile)
    if landscape and not multisection:
        sec = doc.sections[0]
        sec.orientation = WD_ORIENT.LANDSCAPE
        sec.page_width, sec.page_height = sec.page_height, sec.page_width

    if multisection:
        sec2 = doc.add_section(WD_SECTION_START.NEW_PAGE)
        sec2.orientation = WD_ORIENT.LANDSCAPE
        sec2.page_width = Cm(29.7)
        sec2.page_height = Cm(21.0)
        sec2.left_margin = Cm(2.0)
        sec2.right_margin = Cm(2.0)
        sec2.top_margin = Cm(2.0)
        sec2.bottom_margin = Cm(2.0)

    add_body(doc, table_text_before)
    headers = ['STT', 'Hạng mục công việc', 'Kinh phí (VNĐ)']
    data_rows = rows or [
        ['1', 'Khảo sát thực địa và thu thập dữ liệu', '12.500.000'],
        ['2', 'Phân tích số liệu và biên tập tài liệu', '18.000.000'],
    ]
    tbl = add_table(doc, headers, data_rows, profile=profile)
    if col_widths_cm:
        _apply_test_col_widths(tbl, col_widths_cm)
    add_body(doc, table_text_after)
    return _save(tmp_path, doc, name), tbl


# ---------------------------------------------------------------------------
# 1. ooxml_order tests
# ---------------------------------------------------------------------------


def test_ooxml_order_insert_and_check():
    tbl = etree.Element(f'{{{W}}}tbl')
    tbl_grid = etree.Element(f'{{{W}}}tblGrid')
    tbl.append(tbl_grid)

    # Leading property must insert before tblGrid.
    tbl_pr = etree.Element(f'{{{W}}}tblPr')
    ooxml_order.insert_ordered_child(tbl, tbl_pr, replace_existing=False)
    assert list(tbl)[0] is tbl_pr
    assert list(tbl)[1] is tbl_grid

    # Out-of-order children inserted cleanly without duplicates.
    ooxml_order.insert_ordered_child(tbl_pr, etree.Element(f'{{{W}}}tblLayout', {f'{{{W}}}type': 'fixed'}))
    ooxml_order.insert_ordered_child(tbl_pr, etree.Element(f'{{{W}}}tblW', {f'{{{W}}}w': '5000', f'{{{W}}}type': 'dxa'}))
    ooxml_order.insert_ordered_child(tbl_pr, etree.Element(f'{{{W}}}jc', {f'{{{W}}}val': 'center'}))
    ooxml_order.insert_ordered_child(tbl_pr, etree.Element(f'{{{W}}}tblW', {f'{{{W}}}w': '6000', f'{{{W}}}type': 'dxa'}))

    names = [etree.QName(c).localname for c in tbl_pr]
    assert names == ['tblW', 'jc', 'tblLayout']
    assert tbl_pr.find(f'{{{W}}}tblW').get(f'{{{W}}}w') == '6000'
    assert ooxml_order.check_element_order(tbl) == []


def test_ooxml_order_detects_violations():
    doc = etree.Element(f'{{{W}}}document')
    body = etree.SubElement(doc, f'{{{W}}}body')
    tbl = etree.SubElement(body, f'{{{W}}}tbl')
    etree.SubElement(tbl, f'{{{W}}}tblGrid')
    tbl_pr = etree.SubElement(tbl, f'{{{W}}}tblPr')  # invalid: after tblGrid
    etree.SubElement(tbl_pr, f'{{{W}}}tblLayout')
    etree.SubElement(tbl_pr, f'{{{W}}}tblW')  # invalid: tblW after tblLayout

    violations = ooxml_order.check_element_order(doc)
    rules = {v['rule'] for v in violations}
    assert rules == {'leading_property', 'child_order'}

    # Normalizing tblPr puts tblBorders before tblLook cleanly
    tbl_pr2 = etree.Element(f'{{{W}}}tblPr')
    etree.SubElement(tbl_pr2, f'{{{W}}}tblLook')
    etree.SubElement(tbl_pr2, f'{{{W}}}tblBorders')
    ooxml_order.normalize_ordered_children(tbl_pr2)
    assert [etree.QName(c).localname for c in tbl_pr2] == ['tblBorders', 'tblLook']
    assert ooxml_order.check_element_order(tbl_pr2) == []


def test_ooxml_canonical_non_target_preservation():
    doc_before = etree.Element(f'{{{W}}}document')
    b_body = etree.SubElement(doc_before, f'{{{W}}}body')
    for text in ('Đoạn một', 'Bảng mục tiêu', 'Đoạn hai'):
        p = etree.SubElement(b_body, f'{{{W}}}p')
        r = etree.SubElement(p, f'{{{W}}}r')
        t = etree.SubElement(r, f'{{{W}}}t')
        t.text = text

    doc_after = etree.fromstring(etree.tostring(doc_before))
    a_body = doc_after.find(f'{{{W}}}body')
    # Edit index 1 (allowed target) -> xml_ntpr must remain 1.0
    a_body[1].find(f'.//{{{W}}}t').text = 'Đã sửa'
    ok = ooxml_order.check_non_target_preservation(doc_before, doc_after, {1})
    assert ok['xml_ntpr'] == 1.0
    assert ok['unexpected_indices'] == []

    # Collateral edit at index 2 -> caught cleanly
    a_body[2].find(f'.//{{{W}}}t').text = 'Bị sửa nhầm'
    bad = ooxml_order.check_non_target_preservation(doc_before, doc_after, {1})
    assert bad['xml_ntpr'] == 0.5
    assert [u['body_index'] for u in bad['unexpected_indices']] == [2]

    # Child count change -> structure_changed=True
    etree.SubElement(a_body, f'{{{W}}}p')
    struct = ooxml_order.check_non_target_preservation(doc_before, doc_after, {1})
    assert struct['xml_ntpr'] == 0.0
    assert struct['structure_changed'] is True


# ---------------------------------------------------------------------------
# 2. map_docx tests
# ---------------------------------------------------------------------------


def test_map_all_seven_starters():
    starters = sorted(ASSETS_TEMPLATES.glob('*.docx'))
    assert len(starters) == 7

    by_name = {}
    for starter in starters:
        mapped = map_docx.map_document(starter)
        by_name[starter.name] = mapped
        assert mapped['summary']['tables'] >= 1
        for tbl in mapped['tables']:
            assert tbl['grid_cols'] == len(tbl['grid_widths_dxa']) >= 1
            assert tbl['total_width_dxa'] > 0
            assert 'row_texts_nfc' not in tbl

    # Borderless layout templates vs data templates
    assert all(t['is_layout_table'] for t in by_name['administrative-dispatch.docx']['tables'])
    assert all(t['is_layout_table'] for t in by_name['administrative-notice.docx']['tables'])
    assert all(t['is_layout_table'] for t in by_name['party-resolution.docx']['tables'])
    assert all(t['is_layout_table'] for t in by_name['exam-paper.docx']['tables'])

    assert by_name['academic-short-report.docx']['tables'][0]['is_layout_table'] is False
    assert by_name['office-minutes.docx']['tables'][0]['is_layout_table'] is False
    assert all(not t['is_layout_table'] for t in by_name['thesis-research-outline.docx']['tables'])


def test_map_nfc_vs_nfd_identical_anchor_and_markdown(tmp_path):
    nfc_text = 'Cộng hòa xã hội chủ nghĩa Việt Nam — Đề tài nghiên cứu'
    nfd_text = unicodedata.normalize('NFD', nfc_text)
    assert nfc_text != nfd_text

    doc1 = Document()
    setup_document(doc1, 'academic')
    add_table(doc1, ['Tiêu đề'], [[nfc_text]])
    p1 = _save(tmp_path, doc1, 'nfc.docx')

    doc2 = Document()
    setup_document(doc2, 'academic')
    add_table(doc2, ['Tiêu đề'], [[nfd_text]])
    p2 = _save(tmp_path, doc2, 'nfd.docx')

    m1 = map_docx.map_document(p1)
    m2 = map_docx.map_document(p2)
    assert m1['tables'][0]['anchor'] == m2['tables'][0]['anchor']
    assert m1['paragraphs'] == m2['paragraphs']

    md = map_docx.format_markdown(m1)
    assert m1['tables'][0]['anchor'] in md
    assert 'printable_width=' in md


def test_map_multisection_landscape_printable_width(tmp_path):
    doc_path, _ = _make_simple_table_doc(tmp_path, 'multi.docx', multisection=True)
    mapped = map_docx.map_document(doc_path)
    assert len(mapped['sections']) == 2
    assert mapped['sections'][0]['orientation'] == 'portrait'
    assert mapped['sections'][1]['orientation'] == 'landscape'
    assert mapped['sections'][1]['section_index'] == 2
    assert mapped['sections'][1]['printable_width_cm'] == pytest.approx(25.7, abs=0.05)
    assert mapped['tables'][0]['section_index'] == 2
    assert mapped['tables'][0]['printable_width_cm'] == pytest.approx(25.7, abs=0.05)


# ---------------------------------------------------------------------------
# 3. layout_edit tests
# ---------------------------------------------------------------------------


def test_layout_edit_scale_col_widths_and_alignment(tmp_path):
    src, _ = _make_simple_table_doc(tmp_path, 'orig.docx', col_widths_cm=[2.0, 8.0, 5.0])
    m0 = map_docx.map_document(src)
    anchor = m0['tables'][0]['anchor']

    out_scale = tmp_path / 'scaled.docx'
    res_scale = layout_edit.apply_layout_plan(
        src,
        out_scale,
        [{'op': 'table.scale_width', 'target': anchor, 'factor': 0.8}],
        expected_sha256=m0['source_sha256'],
    )
    assert res_scale['status'] == 'ok'
    assert res_scale['non_target_preservation']['xml_ntpr'] == 1.0
    m1 = map_docx.map_document(out_scale)
    assert m1['tables'][0]['total_width_cm'] == pytest.approx(12.0, abs=0.1)
    assert m1['tables'][0]['layout'] == 'fixed'

    # Set explicit percentages & alignment
    out_pct = tmp_path / 'pct_center.docx'
    res_pct = layout_edit.apply_layout_plan(
        out_scale,
        out_pct,
        [
            {'op': 'table.set_col_widths', 'target': m1['tables'][0]['anchor'], 'widths': ['20%', '50%', '30%']},
            {'op': 'table.set_alignment', 'target': m1['tables'][0]['anchor'], 'align': 'center'},
        ],
    )
    assert res_pct['non_target_preservation']['xml_ntpr'] == 1.0
    m2 = map_docx.map_document(out_pct)
    assert m2['tables'][0]['grid_widths_pct'] == [pytest.approx(20.0, abs=0.2), pytest.approx(50.0, abs=0.2), pytest.approx(30.0, abs=0.2)]
    assert m2['tables'][0]['alignment'] == 'center'

    # Untouched ZIP entries byte-for-byte identical, and zero folder entries inside ZIP
    _, orig_parts = read_package(src)
    _, new_parts = read_package(out_scale)
    for part_name, raw in orig_parts.items():
        if part_name != 'word/document.xml':
            assert new_parts[part_name] == raw
    with zipfile.ZipFile(out_scale, 'r') as zf:
        assert not any(info.is_dir() for info in zf.infolist())


def test_layout_edit_starter_auto_tblw_fallback(tmp_path):
    starter = ASSETS_TEMPLATES / 'academic-short-report.docx'
    m0 = map_docx.map_document(starter)
    tbl = m0['tables'][0]
    assert tbl['tblW_raw']['type'] == 'auto'
    assert tbl['width_source'] == 'tblGrid'

    out = tmp_path / 'starter_scaled.docx'
    res = layout_edit.apply_layout_plan(
        starter,
        out,
        [{'op': 'table.scale_width', 'target': tbl['anchor'], 'target_width_pct': 75}],
        profile='academic',
    )
    assert res['status'] == 'ok'
    assert res['non_target_preservation']['xml_ntpr'] == 1.0
    m1 = map_docx.map_document(out)
    assert m1['tables'][0]['total_width_cm'] == pytest.approx(16.5 * 0.75, abs=0.1)


def test_layout_edit_merged_gridspan_and_vmerge(tmp_path):
    doc = Document()
    setup_document(doc, 'academic')
    tbl = doc.add_table(rows=3, cols=3)
    # Row 0: merge cols 0..1 (gridSpan=2), keep col 2
    tbl.cell(0, 0).merge(tbl.cell(0, 1))
    tbl.cell(0, 0).text = 'Tiêu đề gộp ngang'
    tbl.cell(0, 2).text = 'Tổng'
    # Col 0 rows 1..2: vMerge
    tbl.cell(1, 0).merge(tbl.cell(2, 0))
    tbl.cell(1, 0).text = 'Nhóm A'
    tbl.cell(1, 1).text = 'Mục 1'
    tbl.cell(1, 2).text = '100'
    tbl.cell(2, 1).text = 'Mục 2'
    tbl.cell(2, 2).text = '200'

    src = _save(tmp_path, doc, 'merged.docx')
    m0 = map_docx.map_document(src)
    assert m0['tables'][0]['has_merged_cells'] is True
    assert m0['tables'][0]['grid_cols'] == 3

    out = tmp_path / 'merged_edited.docx'
    res = layout_edit.apply_layout_plan(
        src,
        out,
        [{'op': 'table.set_col_widths', 'target': m0['tables'][0]['anchor'], 'widths': ['4cm', '6cm', '3cm']}],
    )
    assert res['status'] == 'ok'
    assert res['non_target_preservation']['xml_ntpr'] == 1.0

    # Verify spanned row 0 cell 0 received exact 4cm + 6cm = 10cm (~5669 dxa)
    _, parts = read_package(out)
    root = etree.fromstring(parts['word/document.xml'])
    tc_ws = root.xpath('.//w:tbl/w:tr[1]/w:tc/w:tcPr/w:tcW/@w:w', namespaces={'w': W})
    assert [int(x) for x in tc_ws] == [5670, 1701]


def test_layout_edit_preflight_errors_warnings_and_dry_run(tmp_path):
    src, _ = _make_simple_table_doc(
        tmp_path,
        'preflight.docx',
        rows=[['1', 'Thử nghiệmChuỗiRấtDàiKhôngNgắt', '999']],
        col_widths_cm=[2.0, 8.0, 5.0],
    )
    m0 = map_docx.map_document(src)
    anchor = m0['tables'][0]['anchor']

    # < 0.51 cm text cell refused
    with pytest.raises(layout_edit.LayoutEditError) as err_narrow:
        layout_edit.apply_layout_plan(
            src, tmp_path / 'x.docx',
            [{'op': 'table.set_col_widths', 'target': anchor, 'widths': ['0.3cm', '8cm', '5cm']}],
        )
    assert err_narrow.value.code == 'COLUMN_TOO_NARROW'

    # Stale anchor
    with pytest.raises(layout_edit.LayoutEditError) as err_stale:
        layout_edit.apply_layout_plan(
            src, tmp_path / 'x.docx',
            [{'op': 'table.scale_width', 'target': 'T99#0000', 'factor': 0.8}],
        )
    assert err_stale.value.code == 'ANCHOR_STALE'

    # Source SHA mismatch
    with pytest.raises(layout_edit.LayoutEditError) as err_sha:
        layout_edit.apply_layout_plan(
            src, tmp_path / 'x.docx',
            [{'op': 'table.scale_width', 'target': anchor, 'factor': 0.8}],
            expected_sha256='0' * 64,
        )
    assert err_sha.value.code == 'SOURCE_SHA_MISMATCH'

    # Warnings: wrap_risk and overflow_printable_width + dry_run does not write
    out_dry = tmp_path / 'should_not_exist.docx'
    res_warn = layout_edit.apply_layout_plan(
        src,
        out_dry,
        [{'op': 'table.set_col_widths', 'target': anchor, 'widths': ['10cm', '1.0cm', '8cm']}],
        dry_run=True,
    )
    assert res_warn['status'] == 'dry_run'
    assert not out_dry.exists()
    warn_codes = {w['code'] for w in res_warn['warnings']}
    assert 'wrap_risk' in warn_codes
    assert 'overflow_printable_width' in warn_codes


def test_layout_edit_layout_table_protection(tmp_path):
    starter = ASSETS_TEMPLATES / 'administrative-dispatch.docx'
    m0 = map_docx.map_document(starter)
    t1_anchor = m0['tables'][0]['anchor']

    with pytest.raises(layout_edit.LayoutEditError) as err_prot:
        layout_edit.apply_layout_plan(
            starter, tmp_path / 'x.docx',
            [{'op': 'table.scale_width', 'target': t1_anchor, 'factor': 0.8}],
        )
    assert err_prot.value.code == 'LAYOUT_TABLE_PROTECTED'

    out = tmp_path / 'allowed.docx'
    res_ok = layout_edit.apply_layout_plan(
        starter,
        out,
        [{'op': 'table.scale_width', 'target': t1_anchor, 'factor': 0.9}],
        allow_layout_tables=True,
    )
    assert res_ok['status'] == 'ok'


def test_cli_dispatch_map_and_layout_edit(tmp_path, capsys):
    src, _ = _make_simple_table_doc(tmp_path, 'cli.docx', col_widths_cm=[2.0, 8.0, 5.0])
    rc = cli.main(['map', str(src)])
    assert rc == 0
    mapped = json.loads(capsys.readouterr().out)
    anchor = mapped['tables'][0]['anchor']

    out = tmp_path / 'cli_out.docx'
    plan_file = tmp_path / 'plan.json'
    plan_file.write_text(
        json.dumps([{'op': 'table.scale_width', 'target': anchor, 'factor': 0.7}]),
        encoding='utf-8',
    )
    rc_edit = cli.main([
        'layout-edit', str(src), str(out),
        '--plan', str(plan_file),
        '--source-sha256', mapped['source_sha256'],
    ])
    assert rc_edit == 0
    edited = json.loads(capsys.readouterr().out)
    assert edited['non_target_preservation']['xml_ntpr'] == 1.0
    assert out.is_file()


# ---------------------------------------------------------------------------
# 4. diff_layout tests (offline + LibreOffice integration)
# ---------------------------------------------------------------------------


def test_diff_layout_not_checked_when_renderer_missing(tmp_path):
    before, _ = _make_simple_table_doc(tmp_path, 'b.docx', col_widths_cm=[2.0, 8.0, 5.0])
    after, _ = _make_simple_table_doc(tmp_path, 'a.docx', col_widths_cm=[2.0, 5.0, 5.0])
    with patch('render_docx.find_renderer', return_value=None):
        report = diff_layout.diff_layout(before, after)
    assert report['status'] == 'not_checked'


@pytest.mark.skipif(not find_renderer(), reason='LibreOffice renderer not installed')
def test_diff_layout_live_scale_alignment_crops_and_ripple(tmp_path):
    # Create a near-full-page document so expanding the table narrows columns,
    # wraps text across extra lines, and pushes a unique closing paragraph to page 2.
    doc = Document()
    setup_document(doc, 'academic')
    for i in range(13):
        p = add_body(doc, f'Đoạn trình bày số {i + 1}: tổng hợp điều kiện kích hoạt nhảy trang khi bảng thay đổi chiều cao.')
        p.paragraph_format.space_after = Pt(6)
    rows = [
        [str(i), f'Hạng mục kỹ thuật số {i} có mô tả đủ dài để xuống dòng nhiều hơn khi thu hẹp bề rộng cột giữa', '10.000.000']
        for i in range(1, 6)
    ]
    tbl = add_table(doc, ['TT', 'Nội dung công việc', 'Giá trị'], rows, profile='academic')
    _apply_test_col_widths(tbl, [1.5, 11.5, 3.5])
    add_body(doc, 'CHỮ KÝ XÁC NHẬN CUỐI BÁO CÁO ĐẶC BIỆT')
    src = _save(tmp_path, doc, 'before_ripple.docx')

    m0 = map_docx.map_document(src)
    anchor = m0['tables'][0]['anchor']

    # Scale down width to 55% + center -> increases wrapping height -> pushes tail paragraph onto page 2
    after_shrink = tmp_path / 'after_shrink.docx'
    layout_edit.apply_layout_plan(
        src,
        after_shrink,
        [
            {'op': 'table.scale_width', 'target': anchor, 'factor': 0.55},
            {'op': 'table.set_alignment', 'target': anchor, 'align': 'center'},
        ],
    )

    crop_dir = tmp_path / 'crops'
    diff_report = diff_layout.diff_layout(
        src,
        after_shrink,
        target=anchor,
        output_dir=tmp_path / 'pdfs',
        crop_dir=crop_dir,
    )
    assert diff_report['status'] == 'ok'
    tm = diff_report['target_metrics']
    assert tm['matched_before'] is True
    assert tm['matched_after'] is True
    assert tm['width_delta_cm'] < -5.0
    assert tm['render_vs_ooxml_error_cm'] <= 0.5
    assert tm['x0_delta_cm'] > 1.0  # centered on page compared to left-aligned
    assert Path(diff_report['crops']['before']).is_file()
    assert Path(diff_report['crops']['after']).is_file()

    # Unconditionally verify non_target_page_shifts catches the tail paragraph pushed onto page 2
    assert diff_report['page_count_before'] == 1
    assert diff_report['page_count_after'] == 2
    assert diff_report['page_count_delta'] == 1
    assert diff_report['pagination_ripple_detected'] is True
    shifted_previews = [s.get('preview', '') for s in diff_report['non_target_page_shifts']]
    assert any('CHỮ KÝ XÁC NHẬN' in pv for pv in shifted_previews)
