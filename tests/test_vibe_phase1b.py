"""Unit and LibreOffice integration tests for Phase 1B Vibe Docs extensions."""
import struct
import zlib
from pathlib import Path

import pytest
from docx import Document
from docx.shared import Cm, Pt
from lxml import etree

import diff_layout
import layout_edit
import map_docx
import ooxml_order
import render_docx
from build_docx import add_body, add_table, setup_document
from docx_package import W, read_package


def _make_png(path, width=800, height=400, rgb=(24, 92, 168)):
    """Create a valid uncompressed-filter truecolor PNG without Pillow."""
    def _chunk(tag, data):
        crc = zlib.crc32(tag + data) & 0xFFFFFFFF
        return struct.pack('!I', len(data)) + tag + data + struct.pack('!I', crc)

    scanline = b'\x00' + bytes(rgb) * width
    raw = scanline * height
    ihdr = struct.pack('!IIBBBBB', width, height, 8, 2, 0, 0, 0)
    payload = (
        b'\x89PNG\r\n\x1a\n'
        + _chunk(b'IHDR', ihdr)
        + _chunk(b'IDAT', zlib.compress(raw, 6))
        + _chunk(b'IEND', b'')
    )
    path.write_bytes(payload)
    return path


def _save(tmp_path, doc, name='sample.docx'):
    out = tmp_path / name
    doc.save(out)
    return out


def _read_doc_xml(docx_path):
    _, parts = read_package(docx_path)
    return etree.fromstring(parts['word/document.xml'])


def test_ooxml_order_rpr_and_table_margins():
    run = etree.Element(f'{{{W}}}r')
    t_el = etree.SubElement(run, f'{{{W}}}t')
    t_el.text = 'Nội dung'
    r_pr = etree.Element(f'{{{W}}}rPr')
    ooxml_order.insert_ordered_child(run, r_pr)
    assert etree.QName(run[0]).localname == 'rPr'
    assert etree.QName(run[1]).localname == 't'

    tbl_mar = etree.Element(f'{{{W}}}tblCellMar')
    for side in ('right', 'bottom', 'top', 'left'):
        el = etree.Element(f'{{{W}}}{side}')
        el.set(f'{{{W}}}w', '100')
        el.set(f'{{{W}}}type', 'dxa')
        ooxml_order.insert_ordered_child(tbl_mar, el)
    assert [etree.QName(c).localname for c in tbl_mar] == ['top', 'left', 'bottom', 'right']

    tc_mar = etree.Element(f'{{{W}}}tcMar')
    for side in ('bottom', 'top', 'right', 'left'):
        el = etree.Element(f'{{{W}}}{side}')
        el.set(f'{{{W}}}w', '90')
        el.set(f'{{{W}}}type', 'dxa')
        ooxml_order.insert_ordered_child(tc_mar, el)
    assert [etree.QName(c).localname for c in tc_mar] == ['top', 'left', 'bottom', 'right']


def test_map_docx_indexes_images(tmp_path):
    img_path = _make_png(tmp_path / 'wide.png', 800, 400)
    doc = Document()
    setup_document(doc, 'academic')
    add_body(doc, 'Đoạn giới thiệu trước hình minh họa.')
    doc.add_picture(str(img_path), width=Cm(18.0))
    add_body(doc, 'Đoạn giải thích sau hình.')
    src = _save(tmp_path, doc, 'img_map.docx')

    mapping = map_docx.map_document(src)
    assert mapping['summary']['images'] == 1
    assert len(mapping['images']) == 1
    img = mapping['images'][0]
    assert img['anchor'].startswith('IMG1#')
    assert img['width_cm'] == pytest.approx(18.0, abs=0.05)
    assert img['height_cm'] == pytest.approx(9.0, abs=0.05)
    assert img['aspect_ratio'] == pytest.approx(2.0, rel=0.02)
    assert img['width_pct_of_printable'] > 100.0
    md = map_docx.format_markdown(mapping)
    assert '## Images' in md
    assert '## Paragraphs' not in md
    assert img['anchor'] in md

    md_para = map_docx.format_markdown(mapping, include_paragraphs=True)
    assert '## Paragraphs' in md_para
    assert mapping['paragraphs'][0]['anchor'] in md_para


def test_table_set_cell_margins_and_font_size(tmp_path):
    doc = Document()
    setup_document(doc, 'administrative')
    add_body(doc, 'Mở đầu bảng số liệu.')
    tbl = add_table(
        doc,
        ['STT', 'Nội dung', 'Ghi chú'],
        [['1', 'Kế hoạch năm', 'Hoàn thành'], ['2', 'Báo cáo quý', 'Đang làm']],
        profile='administrative',
    )
    # Give header 14pt bold and data 14pt, plus a stray cell tcMar override to test clearing.
    for row_idx, row in enumerate(tbl.rows):
        for cell in row.cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(14)
                    if row_idx == 0:
                        run.bold = True
    cell_tcpr = layout_edit._ensure_pr(tbl.rows[1].cells[0]._tc, 'tcPr')
    stale_mar = etree.Element(f'{{{W}}}tcMar')
    etree.SubElement(stale_mar, f'{{{W}}}left').set(f'{{{W}}}w', '400')
    ooxml_order.insert_ordered_child(cell_tcpr, stale_mar)

    src = _save(tmp_path, doc, 'tbl_fmt.docx')
    mapping = map_docx.map_document(src)
    t_anchor = mapping['tables'][0]['anchor']

    dst = tmp_path / 'tbl_fmt_out.docx'
    report = layout_edit.apply_layout_plan(
        src,
        dst,
        [
            {'op': 'table.set_cell_margins', 'target': t_anchor, 'top_cm': 0.12, 'bottom_cm': 0.12, 'left_cm': 0.15, 'right_cm': 0.15},
            {'op': 'table.set_font_size', 'target': t_anchor, 'size_pt': 12.0, 'preserve_header_size': True},
        ],
        profile='administrative',
    )
    assert report['status'] == 'ok'
    assert report['non_target_preservation']['xml_ntpr'] == 1.0
    assert not any(w['code'] == 'small_table_font_warning' for w in report['warnings'])

    root = _read_doc_xml(dst)
    tbl_el = root.xpath('.//w:tbl', namespaces={'w': W})[0]
    left_mar = tbl_el.find(f'{{{W}}}tblPr/{{{W}}}tblCellMar/{{{W}}}left')
    assert left_mar is not None
    assert int(left_mar.get(f'{{{W}}}w')) == pytest.approx(round(0.15 * map_docx.DXA_PER_CM), abs=2)
    assert not tbl_el.xpath('.//w:tcPr/w:tcMar', namespaces={'w': W})

    rows = tbl_el.findall(f'{{{W}}}tr')
    hdr_sz = rows[0].xpath('.//w:rPr/w:sz/@w:val', namespaces={'w': W})
    data_sz = rows[1].xpath('.//w:rPr/w:sz/@w:val', namespaces={'w': W})
    assert all(v == '28' for v in hdr_sz)
    assert all(v == '24' for v in data_sz)
    assert rows[0].xpath('.//w:rPr/w:b', namespaces={'w': W})

    # Verify warning when shrinking table below 12pt on administrative profile.
    dst_small = tmp_path / 'tbl_small.docx'
    warn_report = layout_edit.apply_layout_plan(
        src,
        dst_small,
        [{'op': 'table.set_font_size', 'target': t_anchor, 'size_pt': 10.5}],
        profile='administrative',
    )
    assert any(w['code'] == 'small_table_font_warning' for w in warn_report['warnings'])


def test_image_scale_aspect_ratio_align_and_overflow(tmp_path):
    wide_png = _make_png(tmp_path / 'chart.png', 800, 400)
    icon_png = _make_png(tmp_path / 'icon.png', 160, 160, rgb=(200, 60, 60))
    doc = Document()
    setup_document(doc, 'academic')
    add_body(doc, 'Hình 1: Biểu đồ rộng.')
    doc.add_picture(str(wide_png), width=Cm(18.0))
    add_body(doc, 'Hình 2: Biểu tượng giữ nguyên.')
    doc.add_picture(str(icon_png), width=Cm(3.0))
    src = _save(tmp_path, doc, 'imgs.docx')

    mapping = map_docx.map_document(src)
    assert len(mapping['images']) == 2
    img1 = mapping['images'][0]['anchor']

    dst = tmp_path / 'imgs_scaled.docx'
    report = layout_edit.apply_layout_plan(
        src,
        dst,
        [{'op': 'image.scale', 'target': img1, 'width_cm': 12.0, 'align': 'center'}],
    )
    assert report['status'] == 'ok'
    assert report['non_target_preservation']['xml_ntpr'] == 1.0

    after_map = map_docx.map_document(dst)
    a1, a2 = after_map['images']
    assert a1['width_cm'] == pytest.approx(12.0, abs=0.05)
    assert a1['height_cm'] == pytest.approx(6.0, abs=0.05)
    assert a1['alignment'] == 'center'
    assert a2['width_cm'] == pytest.approx(3.0, abs=0.05)

    # Check both wp:extent and a:ext stayed synchronized on IMG1.
    root = _read_doc_xml(dst)
    d1 = root.xpath('.//w:drawing', namespaces={'w': W})[0]
    ext1 = d1.find(f'{{{map_docx.WP}}}inline/{{{map_docx.WP}}}extent')
    aext1 = d1.xpath('.//a:xfrm/a:ext', namespaces={'a': map_docx.A})[0]
    assert ext1.get('cx') == aext1.get('cx')
    assert ext1.get('cy') == aext1.get('cy')

    # Scaling wider than printable width emits overflow_printable_width warning.
    dst_over = tmp_path / 'imgs_over.docx'
    over_report = layout_edit.apply_layout_plan(
        src,
        dst_over,
        [{'op': 'image.scale', 'target': img1, 'width_cm': 19.5}],
    )
    assert any(w['code'] == 'overflow_printable_width' for w in over_report['warnings'])


def test_paragraph_set_format_party_hd05_and_conflict(tmp_path):
    doc = Document()
    setup_document(doc, 'party')
    p1 = add_body(doc, 'Đoạn thứ nhất cần căn chuẩn Hướng dẫn 05-HD/VPTW.')
    # Inject a stray hanging indent to make sure first_line_indent_cm clears it.
    p1_pr = layout_edit._ensure_pr(p1._p, 'pPr')
    ind = etree.Element(f'{{{W}}}ind')
    ind.set(f'{{{W}}}hanging', '360')
    ooxml_order.insert_ordered_child(p1_pr, ind)
    add_body(doc, 'Đoạn thứ hai giữ nguyên hoàn toàn.')
    src = _save(tmp_path, doc, 'party_para.docx')

    mapping = map_docx.map_document(src)
    p_anchors = [p['anchor'] for p in mapping['paragraphs']]
    target_p1 = p_anchors[0]

    dst = tmp_path / 'party_para_out.docx'
    report = layout_edit.apply_layout_plan(
        src,
        dst,
        [{
            'op': 'paragraph.set_format',
            'targets': [target_p1],
            'align': 'justify',
            'first_line_indent_cm': 1.0,
            'line_spacing_exact_pt': 20.0,
            'space_after_pt': 6.0,
            'keep_with_next': True,
        }],
    )
    assert report['status'] == 'ok'
    assert report['non_target_preservation']['xml_ntpr'] == 1.0

    root = _read_doc_xml(dst)
    paras = root.findall(f'{{{W}}}body/{{{W}}}p')
    ppr = paras[0].find(f'{{{W}}}pPr')
    assert ppr.find(f'{{{W}}}jc').get(f'{{{W}}}val') == 'both'
    ind_out = ppr.find(f'{{{W}}}ind')
    assert ind_out.get(f'{{{W}}}hanging') is None
    assert int(ind_out.get(f'{{{W}}}firstLine')) == pytest.approx(round(map_docx.DXA_PER_CM), abs=2)
    sp = ppr.find(f'{{{W}}}spacing')
    assert sp.get(f'{{{W}}}lineRule') == 'exact'
    assert sp.get(f'{{{W}}}line') == '400'
    assert sp.get(f'{{{W}}}after') == '120'
    assert ppr.find(f'{{{W}}}keepNext') is not None

    with pytest.raises(layout_edit.LayoutEditError) as exc_info:
        layout_edit.apply_layout_plan(
            src,
            tmp_path / 'conflict.docx',
            [{
                'op': 'paragraph.set_format',
                'target': target_p1,
                'line_spacing': 1.15,
                'line_spacing_exact_pt': 20.0,
            }],
        )
    assert exc_info.value.code == 'INVALID_SPACING'


@pytest.mark.skipif(not render_docx.find_renderer(), reason='LibreOffice not installed')
def test_render_annotate_marks_and_diff_layout_image(tmp_path):
    wide_png = _make_png(tmp_path / 'diagram.png', 900, 450, rgb=(15, 110, 80))
    doc = Document()
    setup_document(doc, 'academic')
    add_body(doc, 'Bảng 1: Tổng hợp số liệu.')
    add_table(doc, ['Cột A', 'Cột B'], [['x', 'y']], profile='academic')
    add_body(doc, 'Hình 1: Sơ đồ quá khổ ban đầu.')
    doc.add_picture(str(wide_png), width=Cm(18.0))
    src = _save(tmp_path, doc, 'combo.docx')

    render_dir = tmp_path / 'render'
    ren = render_docx.render_doc(src, output_dir=render_dir, png=True, annotate_marks=True)
    assert ren['status'] == 'rendered'
    assert len(ren['pages']) >= 1
    assert len(ren['annotated_pages']) == len(ren['pages'])
    clean_png = Path(ren['pages'][0])
    marks_png = Path(ren['annotated_pages'][0])
    assert clean_png.is_file() and marks_png.is_file()
    assert clean_png.read_bytes() != marks_png.read_bytes()
    mark_anchors = [m['anchor'] for m in ren['marks']]
    assert any(a.startswith('T1#') for a in mark_anchors)
    assert any(a.startswith('IMG1#') for a in mark_anchors)

    mapping = map_docx.map_document(src)
    img1 = mapping['images'][0]['anchor']
    dst = tmp_path / 'combo_scaled.docx'
    edit_res = layout_edit.apply_layout_plan(
        src,
        dst,
        [{'op': 'image.scale', 'target': img1, 'width_cm': 11.0, 'align': 'center'}],
    )
    assert edit_res['status'] == 'ok'

    crop_dir = tmp_path / 'crops'
    diff_res = diff_layout.diff_layout(
        src,
        dst,
        target='IMG1',
        output_dir=tmp_path / 'diff_out',
        crop_dir=crop_dir,
    )
    assert diff_res['status'] == 'ok'
    assert diff_res['target_kind'] == 'image'
    assert len(diff_res['image_diffs']) == 1
    rec = diff_res['image_diffs'][0]
    assert rec['ooxml_width_delta_cm'] == pytest.approx(-7.0, abs=0.05)
    assert rec['aspect_ratio_delta_pct'] is not None and rec['aspect_ratio_delta_pct'] <= 2.0
    assert rec['rendered_width_cm_after'] == pytest.approx(11.0, abs=0.4)
    assert rec['overflow_right_mm_after'] == 0.0
    assert diff_res['crops'] is not None
    assert diff_res['crops']['before'] and Path(diff_res['crops']['before']).is_file()
    assert diff_res['crops']['after'] and Path(diff_res['crops']['after']).is_file()
