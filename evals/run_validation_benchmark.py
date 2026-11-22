"""Deterministic, reproducible DOCX tooling validation benchmark (Baseline vs. Skill-Guided).

This script measures structural and formatting checks enforced by `scripts/validate_docx.py`
across 8 naive unguided baseline DOCX documents and 8 skill-guided DOCX outputs across
administrative (ND30), Party (HD 05-HD/VPTW), legal-normative (VBQPPL / ND 187/2025),
academic, office, minutes, and custom profiles.
It does NOT measure standalone LLM prompt accuracy; see `evals/audit-comparison.md`
and `evals/manual-results.md` for scope and qualitative evaluations.
"""
import argparse
import json
import sys
import tempfile
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.shared import Cm, Pt, RGBColor

REPO_ROOT = Path(__file__).resolve().parents[1]


def _load_tools():
    scripts_dir = str(REPO_ROOT / 'scripts')
    if scripts_dir not in sys.path:
        sys.path.insert(0, scripts_dir)
    import build_docx
    import create_demo
    import create_templates
    import fill_template
    import validate_docx
    return build_docx, create_demo, create_templates, fill_template, validate_docx


def _para(container, text, style, *, align=None, fli_cm=None, before_pt=0, after_pt=0, line_spacing=1.0):
    p = container.paragraphs[0] if len(container.paragraphs) == 1 and not container.paragraphs[0].text else container.add_paragraph()
    p.style = style
    p.text = text
    if align is not None:
        p.alignment = align
    if fli_cm is not None:
        p.paragraph_format.first_line_indent = Cm(fli_cm)
    p.paragraph_format.space_before = Pt(before_pt)
    p.paragraph_format.space_after = Pt(after_pt)
    p.paragraph_format.line_spacing = line_spacing
    return p


def build_naive_notice(filepath):
    """Build an unguided default python-docx Thông báo simulating naive LLM output."""
    doc = Document()
    # Default US Letter (21.59 x 27.94 cm), default 2.54 cm margins, default Calibri theme font
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n').bold = True
    p.add_run('Độc lập - Tự do - Hạnh phúc\n').bold = True
    p.add_run('-------------------')

    h = doc.add_paragraph('ĐƠN VỊ THỬ NGHIỆM')
    h.runs[0].bold = True
    doc.add_paragraph('Số: 01/TB-TN    |    Hà Nội, ngày 13 tháng 10 năm 2026')

    t = doc.add_paragraph('THÔNG BÁO')
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t.runs[0].bold = True
    t.runs[0].font.size = Pt(14)

    s = doc.add_paragraph('Về lịch họp rà soát tài liệu thử nghiệm')
    s.alignment = WD_ALIGN_PARAGRAPH.CENTER
    s.runs[0].bold = True

    doc.add_paragraph('Đơn vị thử nghiệm thông báo lịch họp rà soát tài liệu như sau:')
    doc.add_paragraph('1. Thời gian: 09 giờ ngày 15 tháng 10 năm 2026 (giả định).')
    doc.add_paragraph('2. Địa điểm: phòng thử nghiệm (giả định).')
    doc.add_paragraph('3. Thành phần: nhóm biên soạn tài liệu thử nghiệm.')
    doc.add_paragraph('Đề nghị thành viên chuẩn bị bản nháp để trao đổi tại cuộc họp./.')
    doc.add_paragraph('Nơi nhận: Nhóm thử nghiệm; Lưu: VT.')
    sig = doc.add_paragraph('NGƯỜI PHỤ TRÁCH THỬ NGHIỆM\n\nNguyễn Văn A (giả định)')
    sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    sig.runs[0].bold = True
    doc.save(filepath)
    return Path(filepath)


def showcase_notice_config():
    """Return shared ND30-inspired configuration for the 2-page showcase notice."""
    return {
        'profile': 'administrative',
        'source': 'Skill showcase notice; ND30 Appendix I and III model 1.4 inspired; simulated data',
        'style_rules': {
            'AdminAgencyParent': {'size_pt': [12, 13], 'bold': False, 'italic': False},
            'AdminAgency': {'size_pt': [12, 13], 'bold': True, 'italic': False},
            'AdminNational': {'size_pt': [12, 13], 'bold': True, 'italic': False},
            'AdminMotto': {'size_pt': [13, 14], 'bold': True, 'italic': False},
            'AdminNumber': {'size_pt': [13, 13], 'bold': False, 'italic': False},
            'AdminDate': {'size_pt': [13, 14], 'bold': False, 'italic': True},
            'Title': {'size_pt': [14, 14], 'bold': True, 'italic': False},
            'AdminSubject': {'size_pt': [13, 14], 'bold': True, 'italic': False},
            'AdminBasis': {'size_pt': [13, 14], 'bold': False, 'italic': True},
            'AdminSectionHeading': {'size_pt': [13, 14], 'bold': True, 'italic': False},
            'AdminBody': {'size_pt': [13, 14], 'bold': False, 'italic': False},
            'AdminTableHeader': {'size_pt': [12, 13], 'bold': True, 'italic': False},
            'AdminTableCell': {'size_pt': [12, 13], 'bold': False, 'italic': False},
            'AdminSigner': {'size_pt': [13, 14], 'bold': True, 'italic': False},
            'AdminSignatureNote': {'size_pt': [12, 13], 'bold': False, 'italic': True},
            'AdminRecipientsLabel': {'size_pt': [12, 12], 'bold': True, 'italic': True},
            'AdminRecipients': {'size_pt': [11, 11], 'bold': False, 'italic': False},
        },
        'required_text': [
            'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM',
            'Độc lập - Tự do - Hạnh phúc',
            'THÔNG BÁO',
            'Nơi nhận:',
        ],
    }


def build_skill_showcase_notice(filepath):
    """Build a realistic 2-page ND30 Thông báo using skill helpers and audited layout."""
    build_docx, _, create_templates, _, _ = _load_tools()
    config = showcase_notice_config()
    doc = Document()
    build_docx.setup_document(doc, 'administrative', config)
    doc.sections[0].different_first_page_header_footer = True

    top = doc.add_table(rows=2, cols=2)
    top.autofit = False
    top.columns[0].width, top.columns[1].width = Cm(6.8), Cm(9.7)
    for row in top.rows:
        create_templates.prevent_row_split(row)
        row.cells[0].width, row.cells[1].width = Cm(6.8), Cm(9.7)
    create_templates.borderless(top)

    left_top, right_top = top.rows[0].cells
    left_meta, right_meta = top.rows[1].cells
    _para(left_top, 'CƠ QUAN CHỦ QUẢN GIẢ ĐỊNH', 'AdminAgencyParent', align=WD_ALIGN_PARAGRAPH.CENTER, fli_cm=0)
    _para(left_top, 'ĐƠN VỊ THỬ NGHIỆM', 'AdminAgency', align=WD_ALIGN_PARAGRAPH.CENTER, fli_cm=0)
    create_templates.rule(left_top, 2.0, 6.42, space_after_pt=2)

    _para(right_top, 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', 'AdminNational', align=WD_ALIGN_PARAGRAPH.CENTER, fli_cm=0)
    _para(right_top, 'Độc lập - Tự do - Hạnh phúc', 'AdminMotto', align=WD_ALIGN_PARAGRAPH.CENTER, fli_cm=0)
    create_templates.rule(right_top, 4.95, 9.32, space_after_pt=2)

    _para(left_meta, 'Số: 12/TB-TN', 'AdminNumber', align=WD_ALIGN_PARAGRAPH.CENTER, fli_cm=0, before_pt=4)
    _para(right_meta, 'Hà Nội, ngày 22 tháng 11 năm 2026', 'AdminDate', align=WD_ALIGN_PARAGRAPH.CENTER, fli_cm=0, before_pt=4)

    _para(doc, 'THÔNG BÁO', 'Title', align=WD_ALIGN_PARAGRAPH.CENTER, fli_cm=0, before_pt=14, after_pt=2)
    _para(doc, 'Về kế hoạch kiểm tra, rà soát hồ sơ lưu trữ cuối năm 2026', 'AdminSubject', align=WD_ALIGN_PARAGRAPH.CENTER, fli_cm=0)
    _para(doc, '(Tài liệu minh họa quy trình — dữ liệu giả định)', 'AdminSubject', align=WD_ALIGN_PARAGRAPH.CENTER, fli_cm=0)
    create_templates.rule(doc, 5.2, 16.5, space_after_pt=12)

    bases = [
        'Căn cứ quy định giả định về công tác văn thư, lưu trữ của Đơn vị thử nghiệm;',
        'Căn cứ Kế hoạch công tác quý IV năm 2026 (tình huống minh họa) đã được phê duyệt ngày 01 tháng 10 năm 2026;',
        'Xét đề nghị của Bộ phận Tổng hợp (giả định) về việc chuẩn bị công tác kiểm tra cuối năm,',
    ]
    for line in bases:
        _para(doc, line, 'AdminBasis', align=WD_ALIGN_PARAGRAPH.JUSTIFY, fli_cm=1.0, after_pt=6, line_spacing=1.15)

    _para(
        doc,
        'Đơn vị thử nghiệm thông báo kế hoạch kiểm tra, rà soát hồ sơ lưu trữ tại các bộ phận trực thuộc như sau:',
        'AdminBody',
        align=WD_ALIGN_PARAGRAPH.JUSTIFY,
        fli_cm=1.0,
        after_pt=6,
        line_spacing=1.15,
    )

    h1 = _para(doc, 'I. MỤC ĐÍCH, PHẠM VI VÀ LỊCH KIỂM TRA', 'AdminSectionHeading', align=WD_ALIGN_PARAGRAPH.LEFT, fli_cm=1.0, before_pt=4, after_pt=6, line_spacing=1.15)
    h1.paragraph_format.keep_with_next = True
    _para(doc, '1. Mục đích: Đánh giá việc lập hồ sơ, bảo quản và bàn giao tài liệu lưu trữ đúng thể thức trong tình huống giả định.', 'AdminBody', align=WD_ALIGN_PARAGRAPH.JUSTIFY, fli_cm=1.0, after_pt=6, line_spacing=1.15)
    _para(doc, '2. Lịch kiểm tra dự kiến tại các bộ phận (dữ liệu minh họa):', 'AdminBody', align=WD_ALIGN_PARAGRAPH.JUSTIFY, fli_cm=1.0, after_pt=6, line_spacing=1.15)

    headers = ['STT', 'Bộ phận (giả định)', 'Nội dung rà soát trọng tâm', 'Thời gian']
    rows = [
        ['1', 'Bộ phận Hành chính', 'Sổ đăng ký văn bản đi/đến và hồ sơ lưu trữ quý I–IV', '25/11/2026'],
        ['2', 'Bộ phận Kế hoạch', 'Hồ sơ kế hoạch định kỳ, biên bản họp và báo cáo tiến độ', '26/11/2026'],
        ['3', 'Bộ phận Tài chính', 'Danh mục chứng từ, hồ sơ quyết toán theo quy định nội bộ', '27/11/2026'],
        ['4', 'Bộ phận Kỹ thuật', 'Hồ sơ nghiệm thu, tài liệu quy trình và nhật ký bàn giao', '28/11/2026'],
    ]
    col_widths = [Cm(1.3), Cm(4.2), Cm(7.7), Cm(3.3)]
    table = doc.add_table(rows=1, cols=4)
    table.autofit = False
    table.style = 'Table Grid'
    repeat = OxmlElement('w:tblHeader')
    table.rows[0]._tr.get_or_add_trPr().append(repeat)
    create_templates.prevent_row_split(table.rows[0])
    for idx, (cell, text) in enumerate(zip(table.rows[0].cells, headers)):
        cell.width = col_widths[idx]
        _para(cell, text, 'AdminTableHeader', align=WD_ALIGN_PARAGRAPH.CENTER, fli_cm=0, before_pt=2, after_pt=2, line_spacing=1.05)
    for row_data in rows:
        row = table.add_row()
        create_templates.prevent_row_split(row)
        for idx, (cell, value) in enumerate(zip(row.cells, row_data)):
            cell.width = col_widths[idx]
            align = WD_ALIGN_PARAGRAPH.CENTER if idx in (0, 3) else WD_ALIGN_PARAGRAPH.LEFT
            _para(cell, value, 'AdminTableCell', align=align, fli_cm=0, before_pt=2, after_pt=2, line_spacing=1.05)

    h2 = _para(doc, 'II. TỔ CHỨC THỰC HIỆN', 'AdminSectionHeading', align=WD_ALIGN_PARAGRAPH.LEFT, fli_cm=1.0, before_pt=10, after_pt=6, line_spacing=1.15)
    h2.paragraph_format.keep_with_next = True
    clauses = [
        '1. Các bộ phận được kiểm tra (giả định): Tự rà soát, phân loại hồ sơ và lập danh mục tài liệu nộp lưu trước ngày 24 tháng 11 năm 2026; cử đầu mối làm việc trực tiếp với Tổ kiểm tra theo lịch nêu trên.',
        '2. Bộ phận Tổng hợp (giả định): Chủ trì chuẩn bị phiếu kiểm tra, hướng dẫn nghiệp vụ lập hồ sơ và tổng hợp kết quả rà soát từ các bộ phận để báo cáo Lãnh đạo đơn vị.',
        '3. Tổ kiểm tra hồ sơ (giả định): Thực hiện đối chiếu danh mục hồ sơ, kiểm tra hình thức trình bày văn bản theo quy định áp dụng, lập biên bản ghi nhận và đề xuất khắc phục đối với các thiếu sót nếu có.',
        '4. Yêu cầu báo cáo tổng hợp: Sau khi kết thúc đợt kiểm tra, Tổ kiểm tra hoàn thiện báo cáo tổng kết và gửi về Bộ phận Tổng hợp trước ngày 02 tháng 12 năm 2026 để trình Lãnh đạo xem xét.',
        '5. Lưu ý phạm vi minh họa: Toàn bộ tên cơ quan, số văn bản, mốc thời gian và bộ phận trong văn bản này đều là dữ liệu giả định phục vụ kiểm thử quy trình trình bày văn bản.',
    ]
    for clause in clauses:
        _para(doc, clause, 'AdminBody', align=WD_ALIGN_PARAGRAPH.JUSTIFY, fli_cm=1.0, after_pt=6, line_spacing=1.15)

    _para(
        doc,
        'Đơn vị thử nghiệm thông báo để các bộ phận và cá nhân liên quan trong tình huống giả định biết, nghiêm túc phối hợp triển khai thực hiện./.',
        'AdminBody',
        align=WD_ALIGN_PARAGRAPH.JUSTIFY,
        fli_cm=1.0,
        after_pt=12,
        line_spacing=1.15,
    )

    bottom = doc.add_table(rows=1, cols=2)
    bottom.autofit = False
    bottom.columns[0].width, bottom.columns[1].width = Cm(7.5), Cm(9.0)
    bottom.cell(0, 0).width, bottom.cell(0, 1).width = Cm(7.5), Cm(9.0)
    create_templates.prevent_row_split(bottom.rows[0])
    create_templates.borderless(bottom)
    left, right = bottom.rows[0].cells
    _para(left, 'Nơi nhận:', 'AdminRecipientsLabel', align=WD_ALIGN_PARAGRAPH.LEFT, fli_cm=0)
    _para(left, '- Lãnh đạo đơn vị (giả định, để b/c);', 'AdminRecipients', align=WD_ALIGN_PARAGRAPH.LEFT, fli_cm=0)
    _para(left, '- Các bộ phận trực thuộc (để t/h);', 'AdminRecipients', align=WD_ALIGN_PARAGRAPH.LEFT, fli_cm=0)
    _para(left, '- Tổ kiểm tra hồ sơ (giả định);', 'AdminRecipients', align=WD_ALIGN_PARAGRAPH.LEFT, fli_cm=0)
    _para(left, '- Lưu: VT, TH.', 'AdminRecipients', align=WD_ALIGN_PARAGRAPH.LEFT, fli_cm=0)

    _para(right, 'TM. ĐƠN VỊ THỬ NGHIỆM', 'AdminSigner', align=WD_ALIGN_PARAGRAPH.CENTER, fli_cm=0)
    _para(right, 'KT. TRƯỞNG ĐƠN VỊ', 'AdminSigner', align=WD_ALIGN_PARAGRAPH.CENTER, fli_cm=0)
    _para(right, 'PHÓ TRƯỞNG ĐƠN VỊ', 'AdminSigner', align=WD_ALIGN_PARAGRAPH.CENTER, fli_cm=0)
    _para(right, '(Không có chữ ký trong tài liệu minh họa)', 'AdminSignatureNote', align=WD_ALIGN_PARAGRAPH.CENTER, fli_cm=0, before_pt=4, after_pt=42)
    _para(right, 'Nguyễn Văn A (giả định)', 'AdminSigner', align=WD_ALIGN_PARAGRAPH.CENTER, fli_cm=0)

    doc.save(filepath)
    return Path(filepath), config


def run_benchmark():
    """Run all 16 deterministic DOCX validation cases and return structured results."""
    build_docx, create_demo, create_templates, fill_template, validate_docx = _load_tools()
    results = []

    with tempfile.TemporaryDirectory(prefix='vndocs-bench-') as tmp:
        root = Path(tmp)
        legal_cfg = json.loads((REPO_ROOT / 'assets/rules/legal-normative.json').read_text(encoding='utf-8'))

        # --- 8 Baseline (Unguided Default python-docx / Misapplied Profile Output) ---
        naive_notice = build_naive_notice(root / 'b1_naive_notice.docx')
        r1 = validate_docx.inspect_doc(naive_notice, 'administrative')
        results.append(('baseline', 'Unguided ND30 notice (Letter, Calibri, default margins)', 'administrative', r1))

        d2 = Document()
        d2.sections[0].page_width, d2.sections[0].page_height = Cm(21.0), Cm(29.7)
        d2.styles['Normal'].font.name = 'Arial'
        d2.add_paragraph('CÔNG VĂN THỬ NGHIỆM')
        p2 = root / 'b2_naive_dispatch.docx'
        d2.save(p2)
        results.append(('baseline', 'Unguided ND30 dispatch (Arial font, 2.54 cm margins)', 'administrative', validate_docx.inspect_doc(p2, 'administrative')))

        d3 = Document()
        d3.sections[0].page_width, d3.sections[0].page_height = Cm(21.0), Cm(29.7)
        d3.sections[0].left_margin = Cm(2.0)
        d3.sections[0].right_margin, d3.sections[0].top_margin, d3.sections[0].bottom_margin = Cm(2.0), Cm(2.0), Cm(2.0)
        d3.styles['Normal'].font.name = 'Times New Roman'
        d3.styles['Normal'].font.color.rgb = RGBColor(0, 0, 0)
        h3 = d3.add_paragraph('BIÊN BẢN HỌP')
        h3.runs[0].font.color.rgb = RGBColor(0x2F, 0x54, 0x96)
        p3 = root / 'b3_styled_minutes.docx'
        d3.save(p3)
        results.append(('baseline', 'Unguided minutes (blue Word heading, 2.0 cm left margin)', 'minutes-administrative', validate_docx.inspect_doc(p3, 'minutes-administrative')))

        d4 = Document()
        build_docx.setup_document(d4, 'administrative')
        d4.add_paragraph('Đơn vị: {{organization}}')
        p4 = root / 'b4_unfilled_token.docx'
        d4.save(p4)
        results.append(('baseline', 'Unresolved {{token}} left in final DOCX', 'administrative', validate_docx.inspect_doc(p4, 'administrative')))

        d5 = Document()
        build_docx.setup_document(d5, 'office')
        d5.add_paragraph('Thời hạn hoàn thành: [CẦN BỔ SUNG: ngày cụ thể]')
        p5 = root / 'b5_unresolved_draft.docx'
        d5.save(p5)
        results.append(('baseline', 'Unresolved [CẦN BỔ SUNG: ...] left in final mode', 'office', validate_docx.inspect_doc(p5, 'office')))

        d6 = Document()
        build_docx.setup_document(d6, 'academic')
        p_bad_num = d6.add_paragraph('Mục bị hỏng liên kết numbering')
        p_bad_num._p.get_or_add_pPr().get_or_add_numPr().get_or_add_numId().val = 99999
        p6 = root / 'b6_dangling_numbering.docx'
        d6.save(p6)
        results.append(('baseline', 'Broken Word numbering reference (numId=99999)', 'academic', validate_docx.inspect_doc(p6, 'academic')))

        d7 = Document()
        build_docx.setup_document(d7, 'administrative')
        d7.sections[0].left_margin = Cm(3.5)
        d7.sections[0].top_margin = Cm(2.5)
        build_docx.add_body(d7, 'NGHỊ QUYẾT CHI BỘ (dùng nhầm lề hành chính NĐ30)')
        p7 = root / 'b7_party_wrong_margins.docx'
        d7.save(p7)
        results.append(('baseline', 'Unguided Party doc (ND30 3.5/2.5 cm margins instead of HD 05 fixed)', 'party', validate_docx.inspect_doc(p7, 'party')))

        d8 = Document()
        build_docx.setup_document(d8, 'administrative')
        d8.sections[0].top_margin = Cm(2.5)
        d8.sections[0].bottom_margin = Cm(2.5)
        build_docx.add_body(d8, 'QUYẾT ĐỊNH QUY PHẠM (lề trên/dưới 2.5 cm thay vì 2.0 cm)')
        p8 = root / 'b8_vbqppl_wrong_margins.docx'
        d8.save(p8)
        results.append(('baseline', 'Unguided VBQPPL doc (2.5 cm top/bottom margins vs. ND 187/2025 2.0 cm)', 'administrative+legal-normative', validate_docx.inspect_doc(p8, 'administrative', legal_cfg)))

        # --- 8 Skill-Guided Outputs ---
        tpl_dir = REPO_ROOT / 'assets/templates'
        for case_id, label in [
            ('administrative-notice', 'ND30 notice starter filled + style role validation'),
            ('academic-short-report', 'Academic short report starter filled + validation'),
            ('office-minutes', 'Office minutes starter filled + validation'),
        ]:
            vals = json.loads((tpl_dir / f'{case_id}.values.json').read_text(encoding='utf-8'))
            cfg = json.loads((tpl_dir / f'{case_id}.rules.json').read_text(encoding='utf-8'))
            out = root / f's_{case_id}.docx'
            fill_template.fill_template(tpl_dir / f'{case_id}.docx', out, vals)
            results.append(('skill', label, cfg['profile'], validate_docx.inspect_doc(out, cfg['profile'], cfg)))

        showcase_path, showcase_cfg = build_skill_showcase_notice(root / 's_showcase_2page_notice.docx')
        results.append(('skill', 'Skill-generated 2-page ND30 notice (table + PAGE field)', 'administrative', validate_docx.inspect_doc(showcase_path, 'administrative', showcase_cfg)))

        demo_info = create_demo.create_demo(root / 's_demo')
        results.append(('skill', 'Multi-section office demo (portrait + landscape annex)', 'minutes-general', validate_docx.inspect_doc(demo_info['output'], 'minutes-general')))

        custom_cfg = json.loads((REPO_ROOT / 'assets/rules/custom-example.json').read_text(encoding='utf-8'))
        d_custom = Document()
        build_docx.setup_document(d_custom, 'custom', custom_cfg)
        build_docx.add_h1(d_custom, 'Nội dung')
        build_docx.add_body(d_custom, 'Tài liệu theo cấu hình riêng.')
        build_docx.add_h1(d_custom, 'Kết luận')
        build_docx.add_body(d_custom, 'Hoàn tất kiểm tra cấu trúc.')
        p_custom = root / 's_custom.docx'
        d_custom.save(p_custom)
        results.append(('skill', 'Custom organization config (Arial 12pt, no landscape)', 'custom', validate_docx.inspect_doc(p_custom, 'custom', custom_cfg)))

        d_party = Document()
        build_docx.setup_document(d_party, 'party')
        build_docx.add_body(d_party, 'Nội dung văn bản của Đảng theo Hướng dẫn 05-HD/VPTW (dữ liệu giả định).')
        p_party = root / 's_party.docx'
        d_party.save(p_party)
        results.append(('skill', 'Skill-guided Party doc (HD 05-HD/VPTW fixed margins, 14pt, 20pt exact)', 'party', validate_docx.inspect_doc(p_party, 'party')))

        d_vbqppl = Document()
        build_docx.setup_document(d_vbqppl, 'administrative', legal_cfg)
        build_docx.add_body(d_vbqppl, 'Dự thảo văn bản quy phạm pháp luật theo Phụ lục I NĐ 187/2025/NĐ-CP (giả định).')
        p_vbqppl = root / 's_vbqppl.docx'
        d_vbqppl.save(p_vbqppl)
        results.append(('skill', 'Skill-guided VBQPPL doc (ND 187/2025 Appendix I rules, 14pt, 2.0 cm top/bottom)', 'administrative+legal-normative', validate_docx.inspect_doc(p_vbqppl, 'administrative', legal_cfg)))

    rows = []
    for group, label, profile, report in results:
        codes = sorted({iss['code'] for iss in report['issues']})
        rows.append({
            'group': group,
            'case': label,
            'profile': profile,
            'status': report['status'],
            'issue_count': len(report['issues']),
            'issue_codes': codes,
        })
    baseline = [r for r in rows if r['group'] == 'baseline']
    skill = [r for r in rows if r['group'] == 'skill']
    summary = {
        'scope': 'Deterministic tooling validation via validate_docx.py; not an independent multi-LLM prompt benchmark',
        'baseline_pass': sum(1 for r in baseline if r['status'] == 'pass'),
        'baseline_total': len(baseline),
        'skill_pass': sum(1 for r in skill if r['status'] == 'pass'),
        'skill_total': len(skill),
        'cases': rows,
    }
    return summary


def format_markdown(summary):
    lines = [
        f"Baseline pass: {summary['baseline_pass']}/{summary['baseline_total']} | "
        f"Skill-guided pass: {summary['skill_pass']}/{summary['skill_total']}",
        '',
        '| Group | Case | Profile | Status | Issues | Codes |',
        '| --- | --- | --- | --- | ---: | --- |',
    ]
    for row in summary['cases']:
        codes = ', '.join(row['issue_codes']) if row['issue_codes'] else '—'
        lines.append(
            f"| {row['group']} | {row['case']} | `{row['profile']}` | `{row['status']}` | {row['issue_count']} | {codes} |"
        )
    return '\n'.join(lines)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--json', action='store_true', help='Emit raw JSON summary instead of Markdown table')
    parser.add_argument('--write-json', metavar='PATH', help='Write raw JSON summary to PATH')
    args = parser.parse_args(argv)
    summary = run_benchmark()
    if args.write_json:
        Path(args.write_json).write_text(json.dumps(summary, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    if args.json:
        print(json.dumps(summary, ensure_ascii=False, indent=2))
    else:
        print(format_markdown(summary))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
