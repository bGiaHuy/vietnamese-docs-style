"""Build project-authored starter templates, with explicit scopes and trial data."""
import argparse
import hashlib
import json
from pathlib import Path

from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

try:
    from .build_docx import setup_document, add_body, add_numbered, add_table
    from .docx_package import ensure_utf8_stdio
    from .fill_template import fill_template
except ImportError:
    from build_docx import setup_document, add_body, add_numbered, add_table
    from docx_package import ensure_utf8_stdio
    from fill_template import fill_template

LEGAL_SOURCE = 'https://datafiles.chinhphu.vn/cpp/files/vbpq/2020/03/30.signed.pdf'
PARTY_SOURCE = 'https://www.sav.gov.vn/Files/Others/761e611b-48d7-4b86-a5a2-4c17bac0f753/05.HD.VPTW.pdf'


def borderless(table):
    props = table._tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        node = OxmlElement('w:' + side)
        node.set(qn('w:val'), 'nil')
        borders.append(node)
    props.append(borders)


def prevent_row_split(row):
    row._tr.get_or_add_trPr().append(OxmlElement('w:cantSplit'))


def role_paragraph(container, text, style, centered=True, reuse_first=False):
    if reuse_first and len(container.paragraphs) == 1 and not container.paragraphs[0].text:
        p = container.paragraphs[0]
        p.style = style
        p.text = text
    else:
        p = container.add_paragraph(text, style=style)
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    if centered:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return p


def rule(container, width_cm, available_cm, space_after_pt=4):
    p = container.add_paragraph()
    p.paragraph_format.left_indent = Cm(max(0, (available_cm-width_cm)/2))
    p.paragraph_format.right_indent = p.paragraph_format.left_indent
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_after = Pt(space_after_pt)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.line_spacing = Pt(6)
    p.add_run().font.size = Pt(1)
    borders = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    for name, value in {'val': 'single', 'sz': '6', 'space': '0', 'color': '000000'}.items():
        bottom.set(qn('w:' + name), value)
    borders.append(bottom)
    p._p.get_or_add_pPr().append(borders)


def notice():
    config = {
        'profile': 'administrative',
        'source': 'Project-authored notice starter; NĐ30 Appendix I and Appendix III model 1.4; human review required',
        'style_rules': {
            'AdminAgency': {'size_pt': [12, 13], 'bold': True, 'italic': False},
            'AdminNational': {'size_pt': [12, 13], 'bold': True, 'italic': False},
            'AdminMotto': {'size_pt': [13, 14], 'bold': True, 'italic': False},
            'AdminNumber': {'size_pt': [13, 13], 'bold': False, 'italic': False},
            'AdminDate': {'size_pt': [13, 14], 'bold': False, 'italic': True},
            'Title': {'size_pt': [14, 14], 'bold': True, 'italic': False},
            'AdminSubject': {'size_pt': [13, 14], 'bold': True, 'italic': False},
            'AdminBody': {'size_pt': [13, 14], 'bold': False, 'italic': False},
            'AdminSigner': {'size_pt': [13, 14], 'bold': True, 'italic': False},
            'AdminSignatureNote': {'size_pt': [12, 13], 'bold': False, 'italic': True},
            'AdminRecipientsLabel': {'size_pt': [12, 12], 'bold': True, 'italic': True},
            'AdminRecipients': {'size_pt': [11, 11], 'bold': False, 'italic': False},
        },
        'required_text': ['CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', 'Độc lập - Tự do - Hạnh phúc', 'THÔNG BÁO', 'Nơi nhận:'],
    }
    doc = Document()
    setup_document(doc, 'administrative', config)
    doc.sections[0].different_first_page_header_footer = True
    top = doc.add_table(rows=2, cols=2)
    top.autofit = False
    top.columns[0].width, top.columns[1].width = Cm(6.8), Cm(9.7)
    for row in top.rows:
        prevent_row_split(row)
        row.cells[0].width, row.cells[1].width = Cm(6.8), Cm(9.7)
    borderless(top)
    left_top, right_top = top.rows[0].cells
    left_meta, right_meta = top.rows[1].cells
    role_paragraph(left_top, '{{organization}}', 'AdminAgency', reuse_first=True)
    rule(left_top, 2.0, 6.42, space_after_pt=2)
    role_paragraph(right_top, 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', 'AdminNational', reuse_first=True)
    role_paragraph(right_top, 'Độc lập - Tự do - Hạnh phúc', 'AdminMotto')
    rule(right_top, 4.95, 9.32, space_after_pt=2)
    p_num = role_paragraph(left_meta, 'Số: {{number}}', 'AdminNumber', reuse_first=True)
    p_num.paragraph_format.space_before = Pt(4)
    p_date = role_paragraph(right_meta, '{{date}}', 'AdminDate', reuse_first=True)
    p_date.paragraph_format.space_before = Pt(4)
    title = doc.add_paragraph('THÔNG BÁO', style='Title')
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.first_line_indent = Cm(0)
    title.paragraph_format.space_before, title.paragraph_format.space_after = Pt(14), Pt(2)
    role_paragraph(doc, '{{subject}}', 'AdminSubject')
    rule(doc, 4.0, 16.5, space_after_pt=12)
    for name in ('opening', 'time', 'location', 'participants', 'request'):
        doc.add_paragraph('{{' + name + '}}', style='AdminBody')
    note_p = add_body(doc, '{{status_note}}')
    note_p.paragraph_format.space_after = Pt(12)
    bottom = doc.add_table(rows=1, cols=2)
    bottom.autofit = False
    bottom.columns[0].width, bottom.columns[1].width = Cm(7.5), Cm(9.0)
    bottom.cell(0, 0).width, bottom.cell(0, 1).width = Cm(7.5), Cm(9.0)
    prevent_row_split(bottom.rows[0])
    borderless(bottom)
    left, right = bottom.rows[0].cells
    role_paragraph(left, 'Nơi nhận:', 'AdminRecipientsLabel', False, reuse_first=True)
    role_paragraph(left, '- {{recipients}};', 'AdminRecipients', False)
    role_paragraph(left, '- Lưu: VT.', 'AdminRecipients', False)
    role_paragraph(right, '{{position}}', 'AdminSigner', reuse_first=True)
    p = role_paragraph(right, '{{signature_note}}', 'AdminSignatureNote')
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(42)
    role_paragraph(right, '{{signer}}', 'AdminSigner')
    values = {
        'organization': 'ĐƠN VỊ THỬ NGHIỆM', 'number': '01/TB-TN',
        'date': 'Hà Nội, ngày 13 tháng 10 năm 2026',
        'subject': 'Về lịch họp rà soát tài liệu thử nghiệm',
        'opening': 'Đơn vị thử nghiệm thông báo lịch họp rà soát tài liệu như sau:',
        'time': '1. Thời gian: 09 giờ ngày 15 tháng 10 năm 2026 (giả định).',
        'location': '2. Địa điểm: phòng thử nghiệm (giả định).',
        'participants': '3. Thành phần: nhóm biên soạn tài liệu thử nghiệm.',
        'request': 'Đề nghị thành viên trong tình huống giả định chuẩn bị bản nháp để trao đổi tại cuộc họp./.',
        'recipients': 'Nhóm thử nghiệm', 'position': 'NGƯỜI PHỤ TRÁCH THỬ NGHIỆM',
        'signer': 'Nguyễn Văn A (giả định)',
        'status_note': 'Tài liệu thử nghiệm; dữ liệu giả định, không dùng để ban hành.',
        'signature_note': '(Không có chữ ký trong bản thử nghiệm)',
    }
    return doc, config, values


def academic():
    config = {'profile': 'academic', 'source': 'Project-authored short report; no university compliance claim',
              'margins_cm': {'left': [2.5, 2.5], 'right': [2, 2], 'top': [2.5, 2.5], 'bottom': [2.5, 2.5]},
              'style_rules': {'Normal': {'size_pt': [13, 13]}, 'Title': {'size_pt': [16, 16], 'bold': True}},
              'required_text': ['Mục đích', 'Phương pháp', 'Kết quả', 'Hạn chế', 'Tài liệu tham khảo']}
    doc = Document()
    setup_document(doc, 'academic', config)
    doc.add_paragraph('{{title}}', style='Title')
    add_body(doc, 'Tác giả: {{author}}')
    add_body(doc, '{{status_note}}')
    for heading, name in [('1. Mục đích', 'purpose'), ('2. Phương pháp', 'method'),
                          ('3. Kết quả', 'results'), ('4. Hạn chế', 'limitations'),
                          ('5. Tài liệu tham khảo', 'references')]:
        doc.add_heading(heading, level=1)
        add_body(doc, '{{' + name + '}}')
    add_table(doc, ['Hạng mục kiểm tra', 'Kỳ vọng'], [['Đánh số', 'Giữ cấp và thứ tự'], ['Dữ kiện', 'Không tự thêm thông tin']], 'academic')
    values = {'title': 'Báo cáo thử nghiệm tạo tài liệu tiếng Việt', 'author': 'Nhóm thử nghiệm (giả định)',
              'purpose': 'Kiểm tra cách tổ chức báo cáo ngắn bằng tiếng Việt và giữ nội dung đầu vào.',
              'method': 'Tạo mẫu có token, điền chuỗi tiếng Việt và đối chiếu nội dung sau khi lưu.',
              'results': 'Ví dụ này trình bày các phần của báo cáo. Không có khảo sát hoặc số liệu thực nghiệm để kết luận hiệu quả.',
              'limitations': 'Mẫu không thay thế hướng dẫn của trường hoặc yêu cầu riêng của giảng viên.',
              'references': 'Tài liệu minh họa không sử dụng nguồn nghiên cứu bên ngoài.',
              'status_note': 'Đây là tài liệu kiểm thử cấu trúc; dữ liệu minh họa, không phải kết quả nghiên cứu đã công bố.'}
    return doc, config, values


def office():
    config = {'profile': 'minutes-general', 'source': 'Project-authored internal meeting notes starter',
              'required_text': ['Nội dung', 'Quyết định', 'Phân công']}
    doc = Document()
    setup_document(doc, 'minutes-general', config)
    doc.add_paragraph('{{title}}', style='Title')
    add_body(doc, '{{metadata}}')
    add_body(doc, '{{status_note}}')
    doc.add_heading('Nội dung', level=1)
    add_body(doc, '{{discussion}}')
    doc.add_heading('Quyết định', level=1)
    add_numbered(doc, ['{{decision}}'])
    doc.add_heading('Phân công', level=1)
    add_table(doc, ['Công việc', 'Phụ trách', 'Thời hạn'], [['{{task}}', '{{owner}}', '{{deadline}}']], 'minutes-general')
    values = {'title': 'Biên bản họp nội bộ thử nghiệm', 'metadata': '09 giờ ngày 13/10/2026; phòng thử nghiệm; nhóm A (giả định).',
              'discussion': 'Nhóm trao đổi về cấu trúc tài liệu và cách ghi nhận công việc.',
              'decision': 'Trong tình huống giả định, nhóm thống nhất chuẩn bị một bản nháp để rà soát.',
              'task': 'Chuẩn bị bản nháp', 'owner': 'Thành viên A (giả định)', 'deadline': '15/10/2026 (giả định)',
              'status_note': 'Tình huống thử nghiệm; các tên và thời hạn dưới đây là giả định.'}
    return doc, config, values


def dispatch():
    config = {
        'profile': 'administrative',
        'source': 'Project-authored official letter starter; NĐ30 Appendix I and Appendix III model 1.5; human review required',
        'style_rules': {
            'AdminAgencyParent': {'size_pt': [12, 13], 'bold': False, 'italic': False},
            'AdminAgency': {'size_pt': [12, 13], 'bold': True, 'italic': False},
            'AdminNational': {'size_pt': [12, 13], 'bold': True, 'italic': False},
            'AdminMotto': {'size_pt': [13, 14], 'bold': True, 'italic': False},
            'AdminNumber': {'size_pt': [13, 13], 'bold': False, 'italic': False},
            'AdminDispatchBrief': {'size_pt': [12, 13], 'bold': False, 'italic': False},
            'AdminDate': {'size_pt': [13, 14], 'bold': False, 'italic': True},
            'AdminRecipientTo': {'size_pt': [13, 14], 'bold': False, 'italic': False},
            'AdminBody': {'size_pt': [13, 14], 'bold': False, 'italic': False},
            'AdminSigner': {'size_pt': [13, 14], 'bold': True, 'italic': False},
            'AdminSignatureNote': {'size_pt': [12, 13], 'bold': False, 'italic': True},
            'AdminRecipientsLabel': {'size_pt': [12, 12], 'bold': True, 'italic': True},
            'AdminRecipients': {'size_pt': [11, 11], 'bold': False, 'italic': False},
        },
        'required_text': ['CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', 'Độc lập - Tự do - Hạnh phúc', 'Kính gửi:', 'Nơi nhận:'],
    }
    doc = Document()
    setup_document(doc, 'administrative', config)
    doc.sections[0].different_first_page_header_footer = True
    top = doc.add_table(rows=2, cols=2)
    top.autofit = False
    top.columns[0].width, top.columns[1].width = Cm(6.8), Cm(9.7)
    for row in top.rows:
        prevent_row_split(row)
        row.cells[0].width, row.cells[1].width = Cm(6.8), Cm(9.7)
    borderless(top)
    left_top, right_top = top.rows[0].cells
    left_meta, right_meta = top.rows[1].cells
    role_paragraph(left_top, '{{parent_org}}', 'AdminAgencyParent', reuse_first=True)
    role_paragraph(left_top, '{{organization}}', 'AdminAgency')
    rule(left_top, 2.4, 6.42, space_after_pt=2)
    role_paragraph(right_top, 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', 'AdminNational', reuse_first=True)
    role_paragraph(right_top, 'Độc lập - Tự do - Hạnh phúc', 'AdminMotto')
    rule(right_top, 4.95, 9.32, space_after_pt=2)
    p_num = role_paragraph(left_meta, 'Số: {{number}}', 'AdminNumber', reuse_first=True)
    p_num.paragraph_format.space_before = Pt(4)
    p_brief = role_paragraph(left_meta, 'V/v {{subject}}', 'AdminDispatchBrief')
    p_brief.paragraph_format.space_before = Pt(6)
    p_date = role_paragraph(right_meta, '{{date}}', 'AdminDate', reuse_first=True)
    p_date.paragraph_format.space_before = Pt(4)

    kg = role_paragraph(doc, 'Kính gửi: {{recipient_to}}', 'AdminRecipientTo')
    kg.paragraph_format.space_before, kg.paragraph_format.space_after = Pt(12), Pt(10)
    for name in ('opening', 'item_1', 'item_2', 'closing'):
        doc.add_paragraph('{{' + name + '}}', style='AdminBody')
    note_p = doc.add_paragraph('{{status_note}}', style='AdminBody')
    note_p.paragraph_format.space_after = Pt(12)

    bottom = doc.add_table(rows=1, cols=2)
    bottom.autofit = False
    bottom.columns[0].width, bottom.columns[1].width = Cm(7.5), Cm(9.0)
    bottom.cell(0, 0).width, bottom.cell(0, 1).width = Cm(7.5), Cm(9.0)
    prevent_row_split(bottom.rows[0])
    borderless(bottom)
    left, right = bottom.rows[0].cells
    role_paragraph(left, 'Nơi nhận:', 'AdminRecipientsLabel', False, reuse_first=True)
    role_paragraph(left, '- Như trên;', 'AdminRecipients', False)
    role_paragraph(left, '- {{recipients}};', 'AdminRecipients', False)
    role_paragraph(left, '- Lưu: VT.', 'AdminRecipients', False)
    role_paragraph(right, '{{authority}}', 'AdminSigner', reuse_first=True)
    role_paragraph(right, '{{position}}', 'AdminSigner')
    p = role_paragraph(right, '{{signature_note}}', 'AdminSignatureNote')
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(42)
    role_paragraph(right, '{{signer}}', 'AdminSigner')
    values = {
        'parent_org': 'UBND TỈNH THỬ NGHIỆM',
        'organization': 'SỞ TỔNG HỢP (GIẢ ĐỊNH)',
        'number': '125/STH-TH',
        'subject': 'rà soát hồ sơ lưu trữ cuối năm (giả định)',
        'date': 'Hà Nội, ngày 22 tháng 11 năm 2026',
        'recipient_to': 'Các phòng chuyên môn trực thuộc (giả định)',
        'opening': 'Thực hiện kế hoạch công tác trong tình huống giả định, Sở Tổng hợp đề nghị các đơn vị triển khai các nội dung sau:',
        'item_1': '1. Rà soát toàn bộ hồ sơ công việc đã giải quyết xong trong năm 2026 và lập mục lục theo quy định nội bộ.',
        'item_2': '2. Gửi danh mục tổng hợp về Bộ phận Văn thư trước ngày 05 tháng 12 năm 2026 (dữ liệu giả định).',
        'closing': 'Đề nghị các đơn vị phối hợp thực hiện đúng tiến độ nêu trên./.',
        'status_note': 'Tài liệu thử nghiệm mẫu 1.5; dữ liệu giả định, không dùng để ban hành.',
        'recipients': 'Lãnh đạo Sở (để b/c)',
        'authority': 'KT. GIÁM ĐỐC',
        'position': 'PHÓ GIÁM ĐỐC',
        'signature_note': '(Không có chữ ký trong bản thử nghiệm)',
        'signer': 'Nguyễn Văn A (giả định)',
    }
    return doc, config, values


def party_resolution():
    config = {
        'profile': 'party',
        'source': 'Project-authored Party resolution starter; QĐ 399-QĐ/TW and HD 05-HD/VPTW (2026); human review required',
        'style_rules': {
            'PartyParent': {'size_pt': [14, 14], 'bold': False, 'italic': False},
            'PartyOrg': {'size_pt': [14, 14], 'bold': True, 'italic': False},
            'PartyTitleTop': {'size_pt': [15, 15], 'bold': True, 'italic': False},
            'PartyNumber': {'size_pt': [14, 14], 'bold': False, 'italic': False},
            'PartyDate': {'size_pt': [14, 14], 'bold': False, 'italic': True},
            'PartyDocType': {'size_pt': [15, 16], 'bold': True, 'italic': False},
            'PartySubject': {'size_pt': [14, 15], 'bold': True, 'italic': False},
            'PartyHeading': {'size_pt': [14, 15], 'bold': True, 'italic': False},
            'PartyBody': {'size_pt': [14, 15], 'bold': False, 'italic': False},
            'PartyAuthority': {'size_pt': [14, 14], 'bold': True, 'italic': False},
            'PartyPosition': {'size_pt': [14, 14], 'bold': False, 'italic': False},
            'PartySignatureNote': {'size_pt': [12, 13], 'bold': False, 'italic': True},
            'PartySigner': {'size_pt': [14, 14], 'bold': True, 'italic': False},
            'PartyRecipientsLabel': {'size_pt': [14, 14], 'bold': False, 'italic': False},
            'PartyRecipients': {'size_pt': [12, 12], 'bold': False, 'italic': False},
        },
        'required_text': ['ĐẢNG CỘNG SẢN VIỆT NAM', 'NGHỊ QUYẾT', 'Nơi nhận:'],
    }
    doc = Document()
    setup_document(doc, 'party', config)
    doc.sections[0].different_first_page_header_footer = True
    top = doc.add_table(rows=2, cols=2)
    top.autofit = False
    top.columns[0].width, top.columns[1].width = Cm(7.4), Cm(9.1)
    for row in top.rows:
        prevent_row_split(row)
        row.cells[0].width, row.cells[1].width = Cm(7.4), Cm(9.1)
    borderless(top)
    left_top, right_top = top.rows[0].cells
    left_meta, right_meta = top.rows[1].cells
    role_paragraph(left_top, '{{parent_party}}', 'PartyParent', reuse_first=True)
    role_paragraph(left_top, '{{party_org}}', 'PartyOrg')
    role_paragraph(left_top, '*', 'PartyBody').paragraph_format.space_after = Pt(2)
    role_paragraph(right_top, 'ĐẢNG CỘNG SẢN VIỆT NAM', 'PartyTitleTop', reuse_first=True)
    rule(right_top, 7.8, 8.72, space_after_pt=2)
    p_num = role_paragraph(left_meta, 'Số {{number}}', 'PartyNumber', reuse_first=True)
    p_num.paragraph_format.space_before = Pt(4)
    p_date = role_paragraph(right_meta, '{{date}}', 'PartyDate', reuse_first=True)
    p_date.paragraph_format.space_before = Pt(4)

    dt = role_paragraph(doc, 'NGHỊ QUYẾT', 'PartyDocType')
    dt.paragraph_format.space_before, dt.paragraph_format.space_after = Pt(12), Pt(2)
    role_paragraph(doc, '{{subject}}', 'PartySubject')
    dash = role_paragraph(doc, '-----', 'PartyBody')
    dash.paragraph_format.space_after = Pt(8)

    for text, is_heading in [
        ('- {{basis}}', False),
        ('I. {{section_1_title}}', True),
        ('{{section_1_body}}', False),
        ('II. {{section_2_title}}', True),
        ('{{section_2_body}}', False),
        ('{{status_note}}', False),
    ]:
        p = doc.add_paragraph(text, style='PartyHeading' if is_heading else 'PartyBody')
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT if is_heading else WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Cm(1.0)
        p.paragraph_format.line_spacing = Pt(20)
        p.paragraph_format.space_after = Pt(10 if text == '{{status_note}}' else 6)
        if is_heading:
            p.paragraph_format.keep_with_next = True

    bottom = doc.add_table(rows=1, cols=2)
    bottom.autofit = False
    bottom.columns[0].width, bottom.columns[1].width = Cm(7.5), Cm(9.0)
    bottom.cell(0, 0).width, bottom.cell(0, 1).width = Cm(7.5), Cm(9.0)
    prevent_row_split(bottom.rows[0])
    borderless(bottom)
    left, right = bottom.rows[0].cells
    lbl = role_paragraph(left, 'Nơi nhận:', 'PartyRecipientsLabel', False, reuse_first=True)
    lbl.runs[0].underline = True
    role_paragraph(left, '- {{recipients}};', 'PartyRecipients', False)
    role_paragraph(left, '- Lưu {{archive}}.', 'PartyRecipients', False)
    role_paragraph(right, '{{authority}}', 'PartyAuthority', reuse_first=True)
    role_paragraph(right, '{{position}}', 'PartyPosition')
    note = role_paragraph(right, '{{signature_note}}', 'PartySignatureNote')
    note.paragraph_format.space_before = Pt(4)
    note.paragraph_format.space_after = Pt(34)
    role_paragraph(right, '{{signer}}', 'PartySigner')
    values = {
        'parent_party': 'ĐẢNG BỘ XÃ THỬ NGHIỆM',
        'party_org': 'CHI BỘ TỔNG HỢP',
        'number': '01-NQ/CB',
        'date': 'Hà Nội, ngày 22 tháng 11 năm 2026',
        'subject': 'về nâng cao chất lượng sinh hoạt chi bộ năm 2027 (giả định)',
        'basis': 'Căn cứ Điều lệ Đảng Cộng sản Việt Nam (tình huống minh họa).',
        'section_1_title': 'TÌNH HÌNH THỰC HIỆN NHIỆM VỤ',
        'section_1_body': 'Tập thể Chi bộ duy trì sinh hoạt định kỳ và phân công nhiệm vụ cho từng đảng viên (nội dung giả định kiểm thử HD 05-HD/VPTW).',
        'section_2_title': 'NHIỆM VỤ TRỌNG TÂM',
        'section_2_body': 'Chuẩn bị tài liệu sinh hoạt chuyên đề trước kỳ họp và đánh giá kết quả thực hiện nghị quyết hàng quý.',
        'status_note': 'Mẫu thử nghiệm văn bản Đảng theo HD 05-HD/VPTW; dữ liệu giả định.',
        'recipients': 'Đảng uỷ xã thử nghiệm (b/c)',
        'archive': 'Chi bộ',
        'authority': 'T/M CHI BỘ',
        'position': 'BÍ THƯ',
        'signature_note': '(Không có chữ ký trong bản thử nghiệm)',
        'signer': 'Trần Văn B (giả định)',
    }
    return doc, config, values


def exam_paper():
    config = {
        'profile': 'academic',
        'source': 'Project-authored exam paper starter; no official institutional template claimed',
        'margins_cm': {'left': [2.0, 2.0], 'right': [2.0, 2.0], 'top': [2.0, 2.0], 'bottom': [2.0, 2.0]},
        'style_rules': {
            'ExamSchool': {'size_pt': [12, 12], 'bold': False, 'italic': False},
            'ExamDept': {'size_pt': [12, 12], 'bold': True, 'italic': False},
            'ExamHeaderTitle': {'size_pt': [12, 13], 'bold': True, 'italic': False},
            'ExamMeta': {'size_pt': [12, 12], 'bold': False, 'italic': True},
            'ExamQuestionTitle': {'size_pt': [13, 13], 'bold': True, 'italic': False},
            'ExamBody': {'size_pt': [13, 13], 'bold': False, 'italic': False},
            'ExamFooterLabel': {'size_pt': [12, 13], 'bold': True, 'italic': False},
        },
        'required_text': ['Thời gian làm bài:', 'Mã đề:', 'Câu 1', 'HẾT'],
    }
    doc = Document()
    setup_document(doc, 'academic', config)
    top = doc.add_table(rows=1, cols=2)
    top.autofit = False
    top.columns[0].width, top.columns[1].width = Cm(7.5), Cm(9.5)
    top.cell(0, 0).width, top.cell(0, 1).width = Cm(7.5), Cm(9.5)
    prevent_row_split(top.rows[0])
    borderless(top)
    left_top, right_top = top.rows[0].cells
    role_paragraph(left_top, '{{school}}', 'ExamSchool', reuse_first=True)
    role_paragraph(left_top, '{{department}}', 'ExamDept')
    rule(left_top, 3.0, 7.12, space_after_pt=2)
    role_paragraph(left_top, '{{exam_status}}', 'ExamMeta')
    role_paragraph(right_top, '{{exam_title}}', 'ExamHeaderTitle', reuse_first=True)
    role_paragraph(right_top, 'Học phần: {{course}}', 'ExamBody')
    role_paragraph(right_top, 'Thời gian làm bài: {{duration}}', 'ExamMeta')
    role_paragraph(right_top, 'Mã đề: {{exam_code}}', 'ExamHeaderTitle')

    rule(doc, 17.0, 17.0, space_after_pt=8)
    inst = role_paragraph(doc, '{{instructions}}', 'ExamMeta', centered=False)
    inst.paragraph_format.space_after = Pt(8)

    for idx in (1, 2, 3):
        qh = role_paragraph(doc, f'Câu {idx} ({{{{q{idx}_points}}}}):', 'ExamQuestionTitle', centered=False)
        qh.paragraph_format.keep_with_next = True
        qh.paragraph_format.space_after = Pt(2)
        qb = doc.add_paragraph('{{' + f'q{idx}_body' + '}}', style='ExamBody')
        qb.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        qb.paragraph_format.first_line_indent = Cm(0)
        qb.paragraph_format.space_after = Pt(6)

    end_p = role_paragraph(doc, '--- HẾT ---', 'ExamQuestionTitle')
    end_p.paragraph_format.space_before, end_p.paragraph_format.space_after = Pt(8), Pt(6)
    note_p = doc.add_paragraph('{{status_note}}', style='ExamBody')
    note_p.paragraph_format.first_line_indent = Cm(0)
    note_p.paragraph_format.space_after = Pt(12)

    bottom = doc.add_table(rows=1, cols=2)
    bottom.autofit = False
    bottom.columns[0].width, bottom.columns[1].width = Cm(8.5), Cm(8.5)
    bottom.cell(0, 0).width, bottom.cell(0, 1).width = Cm(8.5), Cm(8.5)
    prevent_row_split(bottom.rows[0])
    borderless(bottom)
    left_b, right_b = bottom.rows[0].cells
    role_paragraph(left_b, 'GIẢNG VIÊN RA ĐỀ', 'ExamFooterLabel', reuse_first=True)
    ln = role_paragraph(left_b, '{{author_note}}', 'ExamMeta')
    ln.paragraph_format.space_after = Pt(34)
    role_paragraph(left_b, '{{author}}', 'ExamFooterLabel')
    role_paragraph(right_b, 'TRƯỞNG BỘ MÔN', 'ExamFooterLabel', reuse_first=True)
    rn = role_paragraph(right_b, '{{reviewer_note}}', 'ExamMeta')
    rn.paragraph_format.space_after = Pt(34)
    role_paragraph(right_b, '{{reviewer}}', 'ExamFooterLabel')

    values = {
        'school': 'TRƯỜNG ĐẠI HỌC THỬ NGHIỆM',
        'department': 'KHOA CÔNG NGHỆ THÔNG TIN',
        'exam_status': '(Đề thi gồm 01 trang)',
        'exam_title': 'ĐỀ THI KẾT THÚC HỌC PHẦN (MẪU THỬ)',
        'course': 'Nhập môn Kỹ thuật phần mềm',
        'duration': '90 phút (không sử dụng tài liệu)',
        'exam_code': '01',
        'instructions': 'Lưu ý: Sinh viên trả lời trực tiếp vào giấy thi và ghi rõ mã đề.',
        'q1_points': '3,0 điểm',
        'q1_body': 'Phân tích sự khác biệt giữa kiểm thử đơn vị (unit test) và kiểm thử tích hợp trong tình huống giả định.',
        'q2_points': '4,0 điểm',
        'q2_body': 'Thiết kế cấu trúc module kiểm tra dữ liệu đầu vào theo định dạng JSON schema và trình bày các ca kiểm thử biên.',
        'q3_points': '3,0 điểm',
        'q3_body': 'Nêu hai rủi ro thường gặp khi xử lý tài liệu XML đóng gói và cách phòng tránh khi viết bộ công cụ tự động.',
        'status_note': 'Đây là mẫu đề thi giả định; điều chỉnh theo quy định của trường/khoa khi sử dụng thật.',
        'author_note': '(Chữ ký thử nghiệm)',
        'author': 'TS. Nguyễn Văn A (giả định)',
        'reviewer_note': '(Chữ ký duyệt thử nghiệm)',
        'reviewer': 'PGS.TS. Trần Thị B (giả định)',
    }
    return doc, config, values


def thesis_outline():
    config = {
        'profile': 'academic',
        'source': 'Project-authored thesis/research outline starter; adapt to university thesis rules',
        'margins_cm': {'left': [3.0, 3.0], 'right': [2.0, 2.0], 'top': [2.5, 2.5], 'bottom': [2.5, 2.5]},
        'style_rules': {
            'Normal': {'size_pt': [13, 13]},
            'Title': {'size_pt': [16, 16], 'bold': True},
            'ThesisInstitution': {'size_pt': [13, 13], 'bold': False, 'italic': False},
            'ThesisFaculty': {'size_pt': [13, 13], 'bold': True, 'italic': False},
            'ThesisSubtitle': {'size_pt': [13, 14], 'bold': True, 'italic': False},
        },
        'required_text': ['ĐỀ CƯƠNG NGHIÊN CỨU', 'Tính cấp thiết', 'Phương pháp', 'Tiến độ thực hiện', 'Tài liệu tham khảo'],
    }
    doc = Document()
    setup_document(doc, 'academic', config)
    role_paragraph(doc, '{{university}}', 'ThesisInstitution')
    role_paragraph(doc, '{{faculty}}', 'ThesisFaculty')
    rule(doc, 4.5, 16.0, space_after_pt=10)
    title = doc.add_paragraph('ĐỀ CƯƠNG NGHIÊN CỨU', style='Title')
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.first_line_indent = Cm(0)
    title.paragraph_format.space_after = Pt(4)
    sub = role_paragraph(doc, '{{title}}', 'ThesisSubtitle')
    sub.paragraph_format.space_after = Pt(10)

    add_table(doc, ['Thông tin', 'Nội dung'], [
        ['Người thực hiện', '{{author}}'],
        ['Người hướng dẫn', '{{advisor}}'],
        ['Chuyên ngành', '{{major}}'],
        ['Thời gian dự kiến', '{{timeline}}'],
    ], 'academic')
    add_body(doc, '{{status_note}}').paragraph_format.space_before = Pt(6)

    for heading, field in [
        ('1. Tính cấp thiết và câu hỏi nghiên cứu', 'rationale'),
        ('2. Tổng quan nghiên cứu', 'literature'),
        ('3. Phương pháp và dữ liệu', 'methodology'),
    ]:
        doc.add_heading(heading, level=1)
        add_body(doc, '{{' + field + '}}')

    doc.add_heading('4. Tiến độ thực hiện', level=1)
    add_table(doc, ['Giai đoạn', 'Nội dung công việc', 'Sản phẩm dự kiến'], [
        ['{{stage_1}}', '{{task_1}}', '{{deliverable_1}}'],
        ['{{stage_2}}', '{{task_2}}', '{{deliverable_2}}'],
    ], 'academic')

    doc.add_heading('5. Tài liệu tham khảo dự kiến', level=1)
    add_body(doc, '{{references}}')

    values = {
        'university': 'TRƯỜNG ĐẠI HỌC THỬ NGHIỆM',
        'faculty': 'KHOA KHOA HỌC MÁY TÍNH',
        'title': 'Tên đề tài nghiên cứu / khóa luận thử nghiệm (giả định)',
        'author': 'Nguyễn Văn A — Mã HV: 000000 (giả định)',
        'advisor': 'TS. Trần Thị B (giả định)',
        'major': 'Kỹ thuật phần mềm',
        'timeline': '01/2027 – 06/2027 (giả định)',
        'status_note': 'Đây là mẫu đề cương nghiên cứu giả định; bổ sung trang bìa và chuẩn trích dẫn theo quy định của trường.',
        'rationale': 'Trình bày bối cảnh bài toán, khoảng trống nghiên cứu và 1–2 câu hỏi nghiên cứu trọng tâm (nội dung giả định).',
        'literature': 'Tóm tắt các hướng tiếp cận chính trong nước và quốc tế liên quan trực tiếp đến câu hỏi nghiên cứu.',
        'methodology': 'Mô tả bộ dữ liệu dự kiến thu thập, quy trình thực nghiệm, chỉ số đánh giá định lượng và các kiểm tra đối chứng.',
        'stage_1': 'Tháng 1–2',
        'task_1': 'Khảo sát tài liệu, chốt thiết kế',
        'deliverable_1': 'Báo cáo tổng quan & bộ dữ liệu mẫu',
        'stage_2': 'Tháng 3–5',
        'task_2': 'Thử nghiệm, đánh giá và viết luận văn',
        'deliverable_2': 'Bản thảo hoàn chỉnh & mã nguồn',
        'references': 'Liệt kê các tài liệu tham khảo chính theo chuẩn trích dẫn (IEEE hoặc APA) do khoa quy định.',
    }
    return doc, config, values


BUILDERS = {
    'administrative-notice': notice,
    'administrative-dispatch': dispatch,
    'party-resolution': party_resolution,
    'exam-paper': exam_paper,
    'thesis-research-outline': thesis_outline,
    'academic-short-report': academic,
    'office-minutes': office,
}


def _source_for(name):
    if name in ('administrative-notice', 'administrative-dispatch'):
        return LEGAL_SOURCE
    if name == 'party-resolution':
        return PARTY_SOURCE
    return 'Original project structure'


def create_templates(destination):
    folder = Path(destination)
    folder.mkdir(parents=True, exist_ok=True)
    manifest = []
    for name, builder in BUILDERS.items():
        doc, config, values = builder()
        template = folder / (name + '.docx')
        doc.save(template)
        for suffix, data in [('rules', config), ('values', values)]:
            (folder / f'{name}.{suffix}.json').write_text(json.dumps(data, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
        fill_template(template, folder / (name + '.filled.docx'), values)
        manifest.append({'id': name, 'template': template.name, 'profile': config['profile'],
                         'origin': 'project-authored', 'version': 1,
                         'sha256': hashlib.sha256(template.read_bytes()).hexdigest(),
                         'source': _source_for(name),
                         'scope': config['source'], 'external_template_copied': False,
                         'status': 'generated; visual verification is recorded separately',
                         'limitations': 'Starter with single-line slots. Adapt long content and organization-specific requirements; not certified compliance.'})
    (folder / 'manifest.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    return manifest


def main(argv=None):
    ensure_utf8_stdio()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output-dir', default='artifacts/template-suite')
    args = parser.parse_args(argv)
    print(json.dumps(create_templates(args.output_dir), ensure_ascii=False, indent=2))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
