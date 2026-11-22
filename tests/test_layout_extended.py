"""Unit and integration tests for Word COM rendering, render cache, track changes, and extended layout ops."""
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import pytest
from docx import Document
from lxml import etree

import diff_layout
import layout_edit
import map_docx
import ooxml_order
import render_docx
from build_docx import add_body, add_table, setup_document
from docx_package import W, read_package, text_of


def _save(tmp_path, doc, name='sample.docx'):
    out = tmp_path / name
    doc.save(out)
    return out


def _read_xml(docx_path, member='word/document.xml'):
    _, parts = read_package(docx_path)
    return etree.fromstring(parts[member])


def test_table_pagination_borders_and_shading(tmp_path):
    doc = Document()
    setup_document(doc, 'academic')
    add_body(doc, 'Bảng mẫu thử phân trang và viền.')
    add_table(
        doc,
        ['Mã', 'Hạng mục', 'Ghi chú'],
        [['1', 'Mục A', 'Xong'], ['2', 'Mục B', 'Đang chạy'], ['3', 'Mục C', 'Chờ']],
        profile='academic',
    )
    src = _save(tmp_path, doc, 'tbl_p2.docx')
    t1 = map_docx.map_document(src)['tables'][0]['anchor']

    dst = tmp_path / 'tbl_p2_out.docx'
    report = layout_edit.apply_layout_plan(
        src,
        dst,
        [
            {'op': 'table.set_pagination', 'target': t1, 'repeat_header': True, 'header_rows': 1, 'cant_split': True},
            {'op': 'table.set_borders', 'target': t1, 'style': 'single', 'size_eighth_pt': 6, 'color': '1F4E79'},
            {'op': 'table.set_shading', 'target': t1, 'fill': 'D9E2F3', 'scope': 'header'},
            {'op': 'table.set_shading', 'target': t1, 'fill': 'F2F4F8', 'scope': 'odd_rows'},
        ],
    )
    assert report['status'] == 'ok'
    assert report['non_target_preservation']['xml_ntpr'] == 1.0

    root = _read_xml(dst)
    tbl = root.xpath('.//w:tbl', namespaces={'w': W})[0]
    assert not ooxml_order.check_element_order(tbl)

    rows = tbl.findall(f'{{{W}}}tr')
    assert rows[0].find(f'{{{W}}}trPr/{{{W}}}tblHeader') is not None
    assert rows[1].find(f'{{{W}}}trPr/{{{W}}}tblHeader') is None
    assert all(r.find(f'{{{W}}}trPr/{{{W}}}cantSplit') is not None for r in rows)

    top_bdr = tbl.find(f'{{{W}}}tblPr/{{{W}}}tblBorders/{{{W}}}top')
    assert top_bdr.get(f'{{{W}}}color') == '1F4E79'
    assert top_bdr.get(f'{{{W}}}sz') == '6'

    hdr_fill = rows[0].find(f'{{{W}}}tc/{{{W}}}tcPr/{{{W}}}shd').get(f'{{{W}}}fill')
    odd_fill = rows[1].find(f'{{{W}}}tc/{{{W}}}tcPr/{{{W}}}shd').get(f'{{{W}}}fill')
    even_shd = rows[2].find(f'{{{W}}}tc/{{{W}}}tcPr/{{{W}}}shd')
    assert hdr_fill == 'D9E2F3'
    assert odd_fill == 'F2F4F8'
    assert even_shd is None


def test_style_update_leaves_document_xml_untouched(tmp_path):
    doc = Document()
    setup_document(doc, 'academic')
    add_body(doc, 'Đoạn văn bản sử dụng style Normal.')
    src = _save(tmp_path, doc, 'styles_src.docx')
    _, orig_parts = read_package(src)

    dst = tmp_path / 'styles_dst.docx'
    report = layout_edit.apply_layout_plan(
        src,
        dst,
        [{
            'op': 'style.update',
            'style_name': 'Normal',
            'font': 'Times New Roman',
            'size_pt': 14.0,
            'color': '112233',
            'line_spacing_exact_pt': 20.0,
            'space_after_pt': 6.0,
        }],
    )
    assert report['status'] == 'ok'
    assert report['non_target_preservation']['xml_ntpr'] == 1.0
    _, after_parts = read_package(dst)
    assert after_parts['word/document.xml'] == orig_parts['word/document.xml']

    styles_root = _read_xml(dst, 'word/styles.xml')
    normal = layout_edit._find_style_node(styles_root, 'Normal')
    assert not ooxml_order.check_element_order(normal)
    assert normal.find(f'{{{W}}}rPr/{{{W}}}color').get(f'{{{W}}}val') == '112233'
    assert normal.find(f'{{{W}}}rPr/{{{W}}}sz').get(f'{{{W}}}val') == '28'
    assert normal.find(f'{{{W}}}pPr/{{{W}}}spacing').get(f'{{{W}}}lineRule') == 'exact'

    with pytest.raises(layout_edit.LayoutEditError) as exc_info:
        layout_edit.apply_layout_plan(
            src,
            tmp_path / 'missing_style.docx',
            [{'op': 'style.update', 'style_name': 'NonExistentStyleXYZ', 'size_pt': 12.0}],
        )
    assert exc_info.value.code == 'STYLE_NOT_FOUND'


def test_section_set_layout_and_profile_margin_warning(tmp_path):
    doc = Document()
    setup_document(doc, 'administrative')
    add_body(doc, 'Nội dung thử thay đổi khổ trang section.')
    src = _save(tmp_path, doc, 'sect_src.docx')

    dst = tmp_path / 'sect_dst.docx'
    report = layout_edit.apply_layout_plan(
        src,
        dst,
        [{
            'op': 'section.set_layout',
            'section_index': 1,
            'orientation': 'landscape',
            'page_size': 'A4',
            'margins_cm': {'top': 1.2, 'bottom': 1.2, 'left': 2.0, 'right': 1.5},
        }],
        profile='administrative',
    )
    assert report['status'] == 'ok'
    assert report['non_target_preservation']['xml_ntpr'] == 1.0
    assert any(w['code'] == 'profile_margin_warning' for w in report['warnings'])

    mapped = map_docx.map_document(dst)
    sec = mapped['sections'][0]
    assert sec['orientation'] == 'landscape'
    assert sec['page_width_cm'] > sec['page_height_cm']
    assert sec['margins_cm']['top'] == pytest.approx(1.2, abs=0.05)


def test_paragraph_merge_runs_preserves_mixed_formatting(tmp_path):
    doc = Document()
    setup_document(doc, 'academic')
    p = doc.add_paragraph()
    p.add_run('Xin ')
    p.add_run('chào ')
    p.add_run('Việt Nam ')
    bold_r = p.add_run('2026')
    bold_r.bold = True
    p.add_run(' hoàn ')
    p.add_run('chỉnh.')
    add_body(doc, 'Đoạn thứ hai giữ nguyên hoàn toàn.')
    src = _save(tmp_path, doc, 'merge_src.docx')

    mapping = map_docx.map_document(src)
    p1 = mapping['paragraphs'][0]['anchor']

    dst = tmp_path / 'merge_dst.docx'
    report = layout_edit.apply_layout_plan(
        src,
        dst,
        [{'op': 'paragraph.merge_runs', 'targets': [p1]}],
    )
    assert report['status'] == 'ok'
    assert report['operations'][0]['merged_runs'] == 3
    assert report['non_target_preservation']['xml_ntpr'] == 1.0

    root = _read_xml(dst)
    p_after = root.findall(f'{{{W}}}body/{{{W}}}p')[0]
    runs_after = p_after.findall(f'{{{W}}}r')
    assert len(runs_after) == 3
    assert text_of(p_after) == 'Xin chào Việt Nam 2026 hoàn chỉnh.'


def test_track_changes_records_revisions_in_valid_order(tmp_path):
    doc = Document()
    setup_document(doc, 'academic')
    add_body(doc, 'Đoạn văn sẽ đổi căn lề có track changes.')
    add_table(doc, ['Cột 1', 'Cột 2'], [['A', 'B']], profile='academic')
    src = _save(tmp_path, doc, 'tc_src.docx')

    m = map_docx.map_document(src)
    p1 = m['paragraphs'][0]['anchor']
    t1 = m['tables'][0]['anchor']

    dst = tmp_path / 'tc_dst.docx'
    report = layout_edit.apply_layout_plan(
        src,
        dst,
        [
            {'op': 'paragraph.set_format', 'target': p1, 'align': 'center', 'space_after_pt': 8.0},
            {'op': 'table.scale_width', 'target': t1, 'factor': 0.8},
            {'op': 'table.set_shading', 'target': t1, 'fill': 'EEEEEE', 'scope': 'header'},
        ],
        track_changes=True,
        author='qa-tester',
    )
    assert report['status'] == 'ok'
    assert report['track_changes'] is True

    root = _read_xml(dst)
    assert not ooxml_order.check_element_order(root)
    assert root.xpath('.//w:pPrChange[@w:author="qa-tester"]', namespaces={'w': W})
    assert root.xpath('.//w:tblPrChange[@w:author="qa-tester"]', namespaces={'w': W})
    assert root.xpath('.//w:tblGridChange[@w:author="qa-tester"]', namespaces={'w': W})
    assert root.xpath('.//w:tcPrChange[@w:author="qa-tester"]', namespaces={'w': W})


def test_render_word_com_mocked_and_cache_hit(tmp_path):
    doc = Document()
    setup_document(doc, 'academic')
    add_body(doc, 'Kiểm tra cache và backend word-win.')
    src = _save(tmp_path, doc, 'cache_src.docx')
    cache_dir = tmp_path / 'cache'

    # Mocked Word COM conversion + cache population & hit verification
    def fake_word_com(source, work_pdf, timeout=150):
        Path(work_pdf).write_bytes(b'%PDF-1.4\n1 0 obj<<>>endobj\ntrailer<<>>\n%%EOF\n')
        return 'ok', 'Rendered via Microsoft Word COM'

    with patch('render_docx._render_via_word_com', side_effect=fake_word_com) as mock_com:
        out1 = render_docx.render_doc(src, output_dir=tmp_path / 'r1', renderer='word-win', cache_dir=cache_dir)
        assert out1['status'] == 'rendered'
        assert out1['engine'] == 'word-win'
        assert out1['cache_hit'] is False
        assert out1['manifest'] and Path(out1['manifest']).is_file()
        assert mock_com.call_count == 1
        assert render_docx.is_render_fresh(src, out1['manifest']) is True

        out2 = render_docx.render_doc(src, output_dir=tmp_path / 'r2', renderer='word-win', cache_dir=cache_dir)
        assert out2['status'] == 'rendered'
        assert out2['cache_hit'] is True
        assert mock_com.call_count == 1  # Restored cleanly from cache without re-invoking COM

    # Modifying source DOCX must invalidate manifest freshness check immediately
    src2 = tmp_path / 'mutated.docx'
    doc.add_paragraph('Thêm dòng mới làm đổi SHA-256.')
    doc.save(src2)
    assert render_docx.is_render_fresh(src2, out1['manifest']) is False

    # Graceful not_checked when Word COM registry class is unavailable
    fake_proc = SimpleNamespace(returncode=1, stdout='', stderr='80040154 Class not registered')
    with patch('render_docx._is_windows', return_value=True), patch('render_docx.subprocess.run', return_value=fake_proc):
        out_missing = render_docx.render_doc(src, output_dir=tmp_path / 'r3', renderer='word-win')
        assert out_missing['status'] == 'not_checked'

    with patch('render_docx._is_windows', return_value=False):
        out_non_win = render_docx.render_doc(src, output_dir=tmp_path / 'r4', renderer='word-win')
        assert out_non_win['status'] == 'not_checked'

    with patch('render_docx.find_renderer', return_value=None):
        cmp_res = diff_layout.compare_renderers(src, output_dir=tmp_path / 'cmp')
        assert cmp_res['status'] == 'not_checked'
