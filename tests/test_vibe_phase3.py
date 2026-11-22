"""Unit and LibreOffice integration tests for Vibe Docs Phase 3 (P3.1-P3.2)."""
import zipfile
from pathlib import Path

import pytest
from docx import Document
from docx.shared import Cm, Pt
from lxml import etree

import auto_paginate
import generate_docx
import layout_edit
import map_docx
import render_docx
import validate_docx
from build_docx import add_body, add_table, setup_document
from docx_package import audit_namespace_preservation, read_package


def _save(tmp_path, doc, name='sample.docx'):
    out = tmp_path / name
    doc.save(out)
    return out


def test_word365_namespaces_mc_alternate_and_zip_dir_cleanup(tmp_path):
    doc = Document()
    setup_document(doc, 'academic')
    add_body(doc, 'Đoạn mở đầu chứa mc:AlternateContent bên dưới.')
    add_table(doc, ['Cột A', 'Cột B'], [['x', 'y'], ['z', 'w']], profile='academic')
    base_docx = _save(tmp_path, doc, 'base.docx')

    # Inject Word 365 unused namespaces + mc:Ignorable + mc:AlternateContent + stray ZIP dir entry.
    with zipfile.ZipFile(base_docx, 'r') as zin:
        orig_parts = {info.filename: zin.read(info.filename) for info in zin.infolist() if not info.is_dir()}

    doc_root = etree.fromstring(orig_parts['word/document.xml'])
    merged_ns = dict(doc_root.nsmap)
    merged_ns.update({
        'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
        'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
        'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
        'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
    })
    new_root = etree.Element(doc_root.tag, attrib=dict(doc_root.attrib), nsmap=merged_ns)
    new_root.set('{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable', 'w14 w15')
    for ch in list(doc_root):
        new_root.append(ch)
    alt_para = etree.fromstring(
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
        ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
        ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
        ' w14:paraId="1A2B3C4D">'
        '<w:r><mc:AlternateContent><mc:Choice Requires="wps">'
        '<w:t>ChoiceContent</w:t></mc:Choice><mc:Fallback><w:t>FallbackContent</w:t></mc:Fallback>'
        '</mc:AlternateContent></w:r></w:p>'
    )
    new_root.find(f'{{{layout_edit.W}}}body').insert(0, alt_para)
    orig_parts['word/document.xml'] = etree.tostring(new_root, encoding='UTF-8', xml_declaration=True, standalone=True)

    complex_docx = tmp_path / 'word365_complex.docx'
    with zipfile.ZipFile(complex_docx, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        zout.writestr('word/', b'')  # Explicit folder entry to test directory cleanup
        for name, data in orig_parts.items():
            zout.writestr(name, data)

    mapping = map_docx.map_document(complex_docx)
    t1 = mapping['tables'][0]['anchor']

    out_docx = tmp_path / 'word365_out.docx'
    report = layout_edit.apply_layout_plan(
        complex_docx,
        out_docx,
        [{'op': 'table.scale_width', 'target': t1, 'factor': 0.85}],
    )
    assert report['status'] == 'ok'
    assert report['non_target_preservation']['xml_ntpr'] == 1.0

    with zipfile.ZipFile(out_docx, 'r') as zcheck:
        assert not any(info.is_dir() for info in zcheck.infolist())

    _, after_parts = read_package(out_docx)
    ns_audit = audit_namespace_preservation(orig_parts['word/document.xml'], after_parts['word/document.xml'])
    assert ns_audit['preserved'] is True
    assert ns_audit['ignorable_preserved'] is True

    out_root = etree.fromstring(after_parts['word/document.xml'])
    alt_nodes = out_root.xpath('.//mc:AlternateContent', namespaces={'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006'})
    assert len(alt_nodes) == 1
    assert 'ChoiceContent' in etree.tostring(alt_nodes[0], encoding='unicode')


@pytest.mark.skipif(not render_docx.find_renderer(), reason='LibreOffice unavailable')
def test_auto_paginate_orphan_signature_nd30_and_fit_checks(tmp_path):
    # 1. Build NĐ30 dispatch tuned so signature block sits orphaned on top of page 2.
    doc = Document()
    setup_document(doc, 'administrative')
    t_hdr = doc.add_table(rows=1, cols=2)
    t_hdr.cell(0, 0).text = 'BỘ KHOA HỌC VÀ CÔNG NGHỆ\nVĂN PHÒNG'
    t_hdr.cell(0, 1).text = 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc'
    generate_docx._borderless(t_hdr)

    # Give body section slightly larger bottom/top margin inside NĐ30 [2.0, 2.5]cm bounds so page 1 fills cleanly.
    sec = doc.sections[0]
    sec.top_margin = Cm(2.45)
    sec.bottom_margin = Cm(2.45)

    for idx in range(1, 10):
        p = add_body(
            doc,
            f'Căn cứ nội dung công văn số {idx:02d}: Văn phòng Bộ đề nghị các đơn vị trực thuộc khẩn trương '
            f'rà soát lại toàn bộ biểu mẫu báo cáo hành chính định kỳ quý I năm 2026 trước khi trình duyệt.',
        )
        p.paragraph_format.space_after = Pt(8)

    tbl = add_table(
        doc,
        ['STT', 'Đơn vị thực hiện', 'Hạn hoàn thành'],
        [['1', 'Phòng Tổng hợp', '25/11/2026'], ['2', 'Phòng Kế hoạch', '28/11/2026']],
        profile='administrative',
    )
    for row in tbl.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(14)

    t_sig = doc.add_table(rows=1, cols=2)
    t_sig.cell(0, 0).text = 'Nơi nhận:\n- Như trên;\n- Lưu: VT.'
    t_sig.cell(0, 1).text = 'CHÁNH VĂN PHÒNG\n\n\nNguyễn Minh Anh'
    generate_docx._borderless(t_sig)

    src_orphan = _save(tmp_path, doc, 'nd30_orphan.docx')
    ren_before = render_docx.render_doc(src_orphan, output_dir=tmp_path / 'rb')
    assert ren_before['status'] == 'rendered'
    import pymupdf
    with pymupdf.open(ren_before['pdf']) as pdf0:
        if len(pdf0) < 2:
            pytest.fail('Fixture did not spill signature block to page 2 cleanly')
    assert auto_paginate.detect_orphan_signature(src_orphan, ren_before['pdf']) is True

    dst_fixed = tmp_path / 'nd30_fixed.docx'
    res_fix = auto_paginate.auto_paginate(
        src_orphan,
        dst_fixed,
        fix_orphan_signature=True,
        profile='administrative',
        cache_dir=tmp_path / 'cache',
    )
    assert res_fix['status'] == 'ok', res_fix
    assert res_fix['pages_before'] == 2
    assert res_fix['pages_after'] == 1
    assert res_fix['orphan_signature_after'] is False
    assert validate_docx.inspect_doc(dst_fixed, profile='administrative')['status'] == 'pass'

    # Verify exact paragraph text count/strings preserved 100%
    orig_paras = [p.text for p in Document(src_orphan).paragraphs]
    fixed_paras = [p.text for p in Document(dst_fixed).paragraphs]
    assert orig_paras == fixed_paras

    # 2. Already-fitting 1-page doc -> status == 'already_fits'
    already_out = tmp_path / 'already_out.docx'
    res_already = auto_paginate.auto_paginate(
        dst_fixed, already_out, target_pages=1, profile='administrative',
    )
    assert res_already['status'] == 'already_fits'
    assert Path(already_out).is_file()

    # 3. 3-page academic report -> target_pages=2
    doc3 = Document()
    setup_document(doc3, 'academic')
    doc3.sections[0].top_margin = Cm(2.5)
    doc3.sections[0].bottom_margin = Cm(2.5)
    for i in range(1, 36):
        p = add_body(
            doc3,
            f'Đoạn phân tích số {i:02d}: Nội dung tổng hợp đánh giá hiệu năng hệ thống xử lý văn bản điện tử '
            f'và mức độ bảo toàn cấu trúc XML qua các lượt thử nghiệm thực nghiệm độc lập.',
        )
        p.paragraph_format.space_after = Pt(12)
    src3 = _save(tmp_path, doc3, 'report3p.docx')
    ren3 = render_docx.render_doc(src3, output_dir=tmp_path / 'r3')
    with pymupdf.open(ren3['pdf']) as p3:
        assert len(p3) >= 3

    dst2 = tmp_path / 'report2p.docx'
    res3 = auto_paginate.auto_paginate(src3, dst2, target_pages=2, profile='academic')
    assert res3['status'] == 'ok', res3
    assert res3['pages_after'] <= 2
    assert [p.text for p in Document(src3).paragraphs] == [p.text for p in Document(dst2).paragraphs]
