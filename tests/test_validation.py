"""Behavioral tests use temporary files, never overwrite repository fixtures."""
import json
from pathlib import Path
from unittest.mock import patch

import pytest
from docx import Document
from docx.enum.section import WD_SECTION_START, WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from build_docx import setup_document, add_body, add_bullet, add_table, add_numbered
from document_rules import load_rules
from validate_docx import inspect_doc
from fill_template import fill_template
from render_docx import render_doc
from render_docx import find_renderer
from edit_docx import edit_docx
from docx_package import read_package


def saved(tmp_path, doc, name='test.docx'):
    path = tmp_path / name
    doc.save(path)
    return path


def new_doc(profile='academic', config=None):
    doc = Document()
    setup_document(doc, profile, config)
    return doc


def codes(report):
    return {issue['code'] for issue in report['issues']}


def test_administrative_format_does_not_claim_complete_structure(tmp_path):
    doc = new_doc('administrative')
    add_body(doc, 'Văn bản thử nghiệm định dạng, không phải bản hành chính đầy đủ.')
    report = inspect_doc(saved(tmp_path, doc))
    assert report['status'] == 'pass'
    assert any('structure' in item for item in report['unchecked'])
    assert 'required' in codes(inspect_doc(tmp_path / 'test.docx', config={'required_text': ['Nơi nhận:']}))


def test_wrong_page_and_margin(tmp_path):
    doc = new_doc('administrative')
    doc.sections[0].page_width = Cm(22)
    doc.sections[0].right_margin = Cm(3)
    report = inspect_doc(saved(tmp_path, doc))
    assert {'page', 'margin'} <= codes(report)


def test_administrative_landscape_is_supported(tmp_path):
    doc = new_doc('administrative')
    doc.sections[0].page_width, doc.sections[0].page_height = Cm(29.7), Cm(21)
    add_body(doc, 'Bảng biểu trong trang ngang khi trường hợp cho phép.')
    assert 'page' not in codes(inspect_doc(saved(tmp_path, doc)))


def test_academic_margins_not_administrative(tmp_path):
    doc = new_doc()
    doc.sections[0].top_margin = Cm(3.5)
    add_body(doc, 'Mẫu trường có lề riêng.')
    assert 'margin' not in codes(inspect_doc(saved(tmp_path, doc), 'academic'))


def test_general_office_document_has_no_forced_administrative_furniture(tmp_path):
    doc = new_doc('office')
    doc.add_paragraph('Thông báo họp nhóm', style='Title')
    add_body(doc, 'Họp nhóm tại phòng A lúc 09 giờ (dữ liệu thử).')
    assert doc.sections[0].page_width.cm == pytest.approx(21, abs=.01)
    assert not doc.sections[0].different_first_page_header_footer
    assert not any('CỘNG HÒA' in p.text for p in doc.paragraphs)
    assert inspect_doc(saved(tmp_path, doc), 'office')['status'] == 'pass'


def test_custom_preserves_styles_and_page(tmp_path):
    doc = Document()
    doc.styles['Normal'].font.name = 'Arial'
    doc.sections[0].left_margin = Cm(4)
    before = doc.element.xml
    styles_before = doc.styles.element.xml
    setup_document(doc, 'custom')
    assert before == doc.element.xml
    assert styles_before == doc.styles.element.xml
    doc.add_paragraph('Theo mẫu riêng.')
    report = inspect_doc(saved(tmp_path, doc), 'custom')
    assert report['status'] == 'partial'
    assert not report['issues']


def test_administrative_page_field_hidden_only_on_document_first_page():
    doc = Document()
    doc.add_section(WD_SECTION_START.NEW_PAGE)
    setup_document(doc, 'administrative')
    assert doc.sections[0].different_first_page_header_footer
    assert not doc.sections[1].different_first_page_header_footer
    assert doc.sections[0].header._element.xpath('.//w:instrText')[0].text.strip() == 'PAGE'


def test_custom_explicit_rules_and_sections(tmp_path):
    config = {'font': 'Arial', 'color': '000000', 'page_cm': [21, 29.7],
              'margins_cm': {'left': [2, 2]},
              'sections': [{'page_cm': [21, 29.7]}, {'page_cm': [29.7, 21]}]}
    doc = Document()
    doc.add_section(WD_SECTION_START.NEW_PAGE)
    setup_document(doc, 'custom', config)
    assert doc.sections[1].orientation == WD_ORIENT.LANDSCAPE
    doc.add_paragraph('Dữ liệu thử.')
    path = saved(tmp_path, doc)
    assert inspect_doc(path, 'custom', config)['status'] == 'pass'
    config['sections'].append({})
    assert 'sections' in codes(inspect_doc(path, 'custom', config))


@pytest.mark.parametrize('container', ['body', 'table', 'nested_table', 'header', 'footer'])
def test_inherited_font_and_color_in_all_stories(tmp_path, container):
    doc = new_doc()
    doc.styles['Normal'].font.name = 'Arial'
    doc.styles['Normal'].font.color.rgb = RGBColor(255, 0, 0)
    if container == 'body':
        paragraph = doc.add_paragraph()
    elif container in ('header', 'footer'):
        paragraph = getattr(doc.sections[0], container).paragraphs[0]
        paragraph.style = doc.styles['Normal']
    else:
        cell = doc.add_table(rows=1, cols=1).cell(0, 0)
        if container == 'nested_table':
            cell = cell.add_table(rows=1, cols=1).cell(0, 0)
        paragraph = cell.paragraphs[0]
    paragraph.add_run('Định dạng kế thừa sai.')
    assert {'font', 'color'} <= codes(inspect_doc(saved(tmp_path, doc), 'academic'))


def test_split_placeholder_header_and_draft(tmp_path):
    doc = new_doc()
    p = doc.sections[0].header.paragraphs[0]
    p.add_run('[CẦN BỔ ')
    p.add_run('SUNG: người ký]')
    path = saved(tmp_path, doc)
    assert inspect_doc(path, 'academic')['status'] == 'fail'
    report = inspect_doc(path, 'academic', draft=True)
    assert all(x['severity'] == 'warning' for x in report['issues'])


def test_word_numbering_bullet_is_checked_when_config_forbids(tmp_path):
    doc = new_doc()
    add_bullet(doc, 'Bullet Word không có ký tự bullet trong text', profile='academic')
    path = saved(tmp_path, doc)
    assert 'bullet' not in codes(inspect_doc(path, 'academic'))
    assert 'bullet' in codes(inspect_doc(path, 'academic', {'forbid_bullets': True}))


def test_real_numbering_start(tmp_path):
    doc = new_doc()
    paragraphs = add_numbered(doc, ['Năm', 'Sáu'], start=5)
    assert paragraphs[0]._p.pPr.numPr.numId.val == paragraphs[1]._p.pPr.numPr.numId.val
    assert doc.part.numbering_part.element.xpath('./w:abstractNum/w:lvl/w:start[@w:val="5"]')
    assert paragraphs[0].text == 'Năm'
    assert inspect_doc(saved(tmp_path, doc), 'academic')['status'] == 'pass'


def test_dangling_numbering_is_reported_even_when_bullets_allowed(tmp_path):
    doc = new_doc()
    p = doc.add_paragraph('Danh sách hỏng định nghĩa')
    p._p.get_or_add_pPr().get_or_add_numPr().get_or_add_numId().val = 99999
    assert 'numbering' in codes(inspect_doc(saved(tmp_path, doc), 'academic'))


def test_vietnamese_font_override_is_not_hidden_by_ascii(tmp_path):
    doc = new_doc()
    run = doc.add_paragraph().add_run('Văn bản tiếng Việt')
    fonts = run._r.get_or_add_rPr().get_or_add_rFonts()
    fonts.set(qn('w:ascii'), 'Times New Roman')
    fonts.set(qn('w:hAnsi'), 'Arial')
    assert 'font' in codes(inspect_doc(saved(tmp_path, doc), 'academic'))


def test_named_role_checks_size_and_emphasis(tmp_path):
    config = {'style_rules': {'NoticeTitle': {'size_pt': [13, 14], 'bold': True}}}
    doc = new_doc('administrative', config)
    paragraph = doc.add_paragraph('THÔNG BÁO', style='NoticeTitle')
    path = saved(tmp_path, doc)
    assert inspect_doc(path, 'administrative', config)['status'] == 'pass'
    paragraph.runs[0].font.size = Pt(16)
    paragraph.runs[0].bold = False
    assert {'size', 'bold'} <= codes(inspect_doc(saved(tmp_path, doc), 'administrative', config))


def test_template_fill_across_runs_and_stories(tmp_path):
    doc = new_doc('custom', {'font': 'Arial', 'color': '000000'})
    p = doc.add_paragraph()
    p.add_run('Xin chào ')
    p.add_run('{{na').bold = True
    p.add_run('me}} kết thúc.').italic = True
    doc.sections[0].header.paragraphs[0].add_run('{{name}}')
    doc.add_table(rows=1, cols=1).cell(0, 0).text = '{{name}}'
    source = saved(tmp_path, doc, 'source.docx')
    original = source.read_bytes()
    out = tmp_path / 'out.docx'
    result = fill_template(source, out, {'name': 'Nhóm thử nghiệm'})
    changed = Document(out)
    assert changed.paragraphs[0].text == 'Xin chào Nhóm thử nghiệm kết thúc.'
    assert changed.paragraphs[0].runs[1].bold
    assert changed.paragraphs[0].runs[2].italic
    assert result['replacements']['name'] == 3
    assert changed.styles['Normal'].font.name == 'Arial'
    assert source.read_bytes() == original


def test_multiple_tokens_in_one_run_and_empty_value(tmp_path):
    doc = new_doc()
    doc.add_paragraph('{{x}} / {{x}}{{y}}')
    source = saved(tmp_path, doc)
    out = tmp_path / 'out.docx'
    fill_template(source, out, {'x': 'A', 'y': ''})
    assert Document(out).paragraphs[0].text == 'A / A'


def test_fill_refuses_in_place_or_unknown_tokens(tmp_path):
    doc = new_doc()
    doc.add_paragraph('{{missing}}')
    source = saved(tmp_path, doc)
    with pytest.raises(ValueError):
        fill_template(source, source, {})
    with pytest.raises(ValueError):
        fill_template(source, tmp_path / 'out.docx', {})
    assert not (tmp_path / 'out.docx').exists()


def test_fill_refuses_field_paragraph(tmp_path):
    doc = new_doc()
    p = doc.add_paragraph('{{name}}')
    field = OxmlElement('w:fldChar')
    field.set(qn('w:fldCharType'), 'begin')
    p.add_run()._r.append(field)
    source = saved(tmp_path, doc)
    with pytest.raises(ValueError):
        fill_template(source, tmp_path / 'out.docx', {'name': 'An'})


def test_unsupported_hyperlink_token_is_not_silently_left(tmp_path):
    doc = new_doc()
    p = doc.add_paragraph()
    hyperlink = OxmlElement('w:hyperlink')
    run = OxmlElement('w:r')
    text = OxmlElement('w:t')
    text.text = '{{missing}}'
    run.append(text)
    hyperlink.append(run)
    p._p.append(hyperlink)
    source = saved(tmp_path, doc)
    with pytest.raises(ValueError, match='Unfilled'):
        fill_template(source, tmp_path / 'out.docx', {})
    report = fill_template(source, tmp_path / 'draft.docx', {}, require_all=False)
    assert report['unfilled'] == ['missing']


def test_fill_preserves_opaque_members_and_style_bytes(tmp_path):
    from zipfile import ZipFile
    doc = new_doc()
    doc.add_paragraph('{{name}}')
    source = saved(tmp_path, doc)
    with ZipFile(source, 'a') as package:
        package.writestr('customXml/unrecognized-part.bin', b'opaque-data-123')
    out = tmp_path / 'out.docx'
    report = fill_template(source, out, {'name': 'An'})
    with ZipFile(source) as a, ZipFile(out) as b:
        assert set(a.namelist()) == set(b.namelist())
        for name in a.namelist():
            if name not in report['changed_parts']:
                assert a.read(name) == b.read(name), name
    assert report['changed_parts'] == ['word/document.xml']


def test_verified_slot_edit_and_checksum(tmp_path):
    import hashlib
    doc = new_doc()
    p = doc.add_paragraph('Thời gian: ')
    p.add_run('09 giờ').bold = True
    doc.add_paragraph('Không thay đổi.')
    source = saved(tmp_path, doc)
    out = tmp_path / 'out.docx'
    edits = [{'part': 'word/document.xml', 'paragraph': 0, 'before': '09 giờ', 'after': '10 giờ'}]
    report = edit_docx(source, out, edits, hashlib.sha256(source.read_bytes()).hexdigest())
    changed = Document(out)
    assert changed.paragraphs[0].text == 'Thời gian: 10 giờ'
    assert changed.paragraphs[0].runs[1].bold
    assert changed.paragraphs[1].text == 'Không thay đổi.'
    assert report['preserved_parts'] > 5
    with pytest.raises(ValueError):
        edit_docx(source, out, edits, '0'*64)
    edits[0]['before'] = '11 giờ'
    with pytest.raises(ValueError):
        edit_docx(source, out, edits)


def test_slot_edit_preserves_a_tab_outside_the_span(tmp_path):
    doc = new_doc()
    p = doc.add_paragraph()
    p.add_run().add_tab()
    p.add_run('Giờ: 09 giờ')
    source = saved(tmp_path, doc)
    output = tmp_path / 'out.docx'
    edit_docx(source, output, [{'part': 'word/document.xml', 'paragraph': 0,
                              'before': '09 giờ', 'after': '10 giờ'}])
    assert Document(output).paragraphs[0].text == '\tGiờ: 10 giờ'
    with pytest.raises(ValueError, match='cross'):
        edit_docx(source, output, [{'part': 'word/document.xml', 'paragraph': 0,
                                  'before': '\tGiờ: 09 giờ', 'after': '10 giờ'}])


def test_verified_empty_table_slot(tmp_path):
    doc = new_doc()
    doc.add_table(rows=1, cols=1).cell(0, 0).paragraphs[0].add_run().bold = True
    source = saved(tmp_path, doc)
    output = tmp_path / 'out.docx'
    edit_docx(source, output, [{'part': 'word/document.xml', 'paragraph': 0,
                              'before': '', 'after': 'Nội dung điền'}])
    cell = Document(output).tables[0].cell(0, 0)
    assert cell.text == 'Nội dung điền'
    assert cell.paragraphs[0].runs[0].bold
    with pytest.raises(ValueError):
        edit_docx(output, tmp_path / 'other.docx', [{'part': 'word/document.xml', 'paragraph': 0,
                                                   'before': '', 'after': 'Không được ghi đè'}])


def test_signed_package_is_not_silently_invalidated(tmp_path):
    from zipfile import ZipFile
    source = saved(tmp_path, new_doc())
    with ZipFile(source, 'a') as package:
        package.writestr('_xmlsignatures/sig1.xml', b'<signature/>')
    with pytest.raises(ValueError, match='signatures'):
        read_package(source)


def test_missing_renderer_is_not_success(tmp_path):
    source = saved(tmp_path, new_doc())
    with patch('render_docx.find_renderer', return_value=None):
        result = render_doc(source)
    assert result['status'] == 'not_checked'
    assert result['pdf'] is None


def test_stale_pdf_cannot_pass_conversion(tmp_path):
    source = saved(tmp_path, new_doc())
    output = tmp_path / 'render'
    output.mkdir()
    (output / 'test.pdf').write_bytes(b'%PDF-stale')
    with patch('render_docx.find_renderer', return_value='fake'), patch('render_docx.subprocess.run') as run:
        run.return_value.returncode = 0
        run.return_value.stdout = ''
        run.return_value.stderr = ''
        result = render_doc(source, output)
    assert result['status'] == 'fail'


def test_bad_rule_config_rejected():
    with pytest.raises(ValueError):
        load_rules('custom', {'margins_cm': {'left': [4, 2]}})
    with pytest.raises(ValueError):
        load_rules('custom', {'frobnicate': True})
    with pytest.raises(ValueError):
        load_rules('not-a-profile')


def test_table_row_shape():
    with pytest.raises(ValueError):
        add_table(new_doc(), ['A', 'B'], [['only one']])


@pytest.mark.parametrize('template_id', [
    'administrative-notice',
    'administrative-dispatch',
    'party-resolution',
    'exam-paper',
    'thesis-research-outline',
    'academic-short-report',
    'office-minutes',
])
def test_starter_template_fill_and_role_validation(tmp_path, template_id):
    from create_templates import BUILDERS
    doc, config, values = BUILDERS[template_id]()
    source = saved(tmp_path, doc, 'template.docx')
    output = tmp_path / 'filled.docx'
    fill_template(source, output, values)
    assert inspect_doc(output, config['profile'], config)['status'] == 'pass'
    assert not any(issue['code'] == 'placeholder' for issue in inspect_doc(output, config['profile'], config)['issues'])
    # Trial labels are data, not hardcoded content forced on a real user's output.
    assert '{{status_note}}' in ''.join(p.text for p in doc.paragraphs)


def test_administrative_notice_nd30_header_alignment_and_signature_layout(tmp_path):
    from create_templates import BUILDERS
    doc, config, values = BUILDERS['administrative-notice']()
    top, bottom = doc.tables
    assert len(top.rows) == 2 and len(top.columns) == 2
    assert '{{organization}}' in top.cell(0, 0).paragraphs[0].text
    assert 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM' in top.cell(0, 1).paragraphs[0].text
    assert 'Số: {{number}}' in top.cell(1, 0).paragraphs[0].text
    assert '{{date}}' in top.cell(1, 1).paragraphs[0].text
    assert bottom.rows[0]._tr.trPr.find(qn('w:cantSplit')) is not None
    assert bottom.cell(0, 0).paragraphs[0].text == 'Nơi nhận:'
    assert bottom.cell(0, 1).paragraphs[0].text == '{{position}}'
    source = saved(tmp_path, doc, 'notice.docx')
    filled = tmp_path / 'notice-filled.docx'
    fill_template(source, filled, values)
    report = inspect_doc(filled, config['profile'], config)
    assert report['status'] == 'pass', report
    assert Document(filled).tables[1].cell(0, 1).paragraphs[1].style.name == 'AdminSignatureNote'


def test_cli_doctor_reports_environment():
    import cli
    report = cli.doctor_report()
    assert report['python']['supported'] is True
    assert report['dependencies']['python-docx']['installed'] is True
    assert report['dependencies']['lxml']['installed'] is True
    assert report['status'] in ('ok', 'partial')
    with patch('render_docx.find_renderer', return_value=None):
        assert cli.run_doctor([]) == 2


def test_cli_validate_dispatch(tmp_path):
    import cli
    path = saved(tmp_path, new_doc('office'), 'office.docx')
    assert cli.main(['validate', str(path), '--profile', 'office']) == 0
    assert cli.main(['--help']) == 0
    assert cli.main(['unknown-command']) == 2


def test_package_excludes_research_and_detects_changed_template(tmp_path):
    import hashlib
    from zipfile import ZipFile
    from package_skill import package_skill
    root = tmp_path/'project'
    root.mkdir()
    (root/'SKILL.md').write_text('---\nname: vietnamese-docx\ndescription: Test fixture\n---\n', encoding='utf-8')
    (root/'requirements.txt').write_text('python-docx\n', encoding='utf-8')
    (root/'references/source-documents').mkdir(parents=True)
    (root/'references/source-documents/private-example.txt').write_text('excluded research', encoding='utf-8')
    assets = root/'assets/templates'
    assets.mkdir(parents=True)
    template = assets/'test.docx'
    new_doc().save(template)
    item = {'id': 'test', 'template': 'test.docx', 'origin': 'project-authored',
            'visual_trial': 'reviewed', 'sha256': hashlib.sha256(template.read_bytes()).hexdigest()}
    (assets/'manifest.json').write_text(json.dumps([item]), encoding='utf-8')
    for suffix in ('rules', 'values'):
        (assets/f'test.{suffix}.json').write_text('{}', encoding='utf-8')
    package = tmp_path/'skill.zip'
    package_skill(root, package)
    with ZipFile(package) as archive:
        assert 'vietnamese-docx/SKILL.md' in archive.namelist()
        assert not any('source-documents' in name for name in archive.namelist())
    template.write_bytes(b'changed since verification')
    with pytest.raises(ValueError, match='changed'):
        package_skill(root, package)


@pytest.mark.skipif(not find_renderer(), reason='LibreOffice unavailable')
def test_end_to_end_filled_template_render(tmp_path):
    from create_demo import create_demo
    import pymupdf
    result = create_demo(tmp_path)
    filled = Path(result['output'])
    assert inspect_doc(filled, 'minutes-general')['status'] == 'pass'
    rendering = render_doc(filled, tmp_path / 'render', png=True)
    assert rendering['status'] == 'rendered', rendering
    with pymupdf.open(rendering['pdf']) as pdf:
        assert len(pdf) == 2
        assert pdf[0].rect.height > pdf[0].rect.width
        assert pdf[1].rect.width > pdf[1].rect.height
        assert 'Thành viên A' in pdf[0].get_text()
        assert '{{' not in ''.join(page.get_text() for page in pdf)
    assert len(rendering['pages']) == 2


@pytest.mark.skipif(not find_renderer(), reason='LibreOffice unavailable')
def test_administrative_second_page_has_real_page_number(tmp_path):
    import pymupdf
    doc = new_doc('administrative')
    doc.add_paragraph('Trang bắt đầu')
    doc.add_page_break()
    doc.add_paragraph('Trang tiếp theo')
    report = render_doc(saved(tmp_path, doc), tmp_path/'render')
    assert report['status'] == 'rendered', report
    with pymupdf.open(report['pdf']) as pdf:
        assert len(pdf) == 2
        assert pdf[0].get_text().splitlines()[0].strip() == 'Trang bắt đầu'
        assert pdf[1].get_text().splitlines()[0].strip() == '2'


def test_party_profile_follows_guideline_05_layout(tmp_path):
    doc = new_doc('party')
    add_body(doc, 'Nội dung thử nghiệm văn bản của Đảng.')
    section = doc.sections[0]
    margins = [section.left_margin, section.right_margin, section.top_margin,
               section.bottom_margin, section.header_distance]
    assert [round(m.cm, 2) for m in margins] == [3.0, 1.5, 2.0, 2.0, 1.0]
    normal = doc.styles['Normal']
    assert normal.font.size == Pt(14)
    assert normal.paragraph_format.line_spacing == Pt(20)
    assert round(normal.paragraph_format.first_line_indent.cm, 2) == 1.0
    assert section.different_first_page_header_footer
    page_runs = [run for p in section.header.paragraphs for run in p.runs]
    assert page_runs and page_runs[0].font.size == Pt(14)
    assert inspect_doc(saved(tmp_path, doc), 'party')['status'] == 'pass'


def test_party_profile_rejects_administrative_margins(tmp_path):
    doc = new_doc('administrative')
    add_body(doc, 'Lề hành chính không khớp lề cố định của văn bản Đảng.')
    doc.sections[0].left_margin = Cm(3.5)
    assert 'margin' in codes(inspect_doc(saved(tmp_path, doc), 'party'))


def test_exact_line_spacing_config_is_validated():
    with pytest.raises(ValueError):
        load_rules('custom', {'line_spacing_exact_pt': 0})
    assert load_rules('office')['line_spacing_exact_pt'] is None


def test_legal_normative_rules_config_loads():
    root = Path(__file__).resolve().parents[1]
    rules = load_rules('administrative', root / 'assets/rules/legal-normative.json')
    assert rules['margins_cm']['top'] == [2.0, 2.0]
    assert rules['body_size_pt'] == 14


def test_skill_routes_only_to_existing_references():
    import re
    root = Path(__file__).resolve().parents[1]
    skill = (root / 'SKILL.md').read_text(encoding='utf-8')
    paths = set(re.findall(r'\]\((references/[\w./-]+\.md)\)', skill))
    assert {'references/party-documents.md', 'references/legal-normative-documents.md',
            'references/administrative-doc-types-nd30.md',
            'references/administrative-writing-style.md'} <= paths
    assert all((root / path).is_file() for path in paths)
    actual_refs = {f'references/{p.name}' for p in (root / 'references').glob('*.md')}
    assert paths == actual_refs
    assert not re.search(r'(?:references|scripts|assets)\\', skill)


def test_skill_conforms_to_anthropic_spec():
    import re
    import yaml

    root = Path(__file__).resolve().parents[1]
    skill = (root / 'SKILL.md').read_text(encoding='utf-8')
    fm_match = re.match(r'^---\r?\n(.*?)\r?\n---\r?\n(.*)$', skill, re.S)
    assert fm_match is not None
    fm = yaml.safe_load(fm_match.group(1))
    body = fm_match.group(2)

    assert isinstance(fm.get('name'), str)
    assert re.fullmatch(r'[a-z0-9-]+', fm['name'])
    assert 1 <= len(fm['name']) <= 64

    desc = fm.get('description', '').strip()
    assert desc and len(desc) <= 1024
    assert not re.search(r'<[^>]+>', desc)

    compat = fm.get('compatibility', '').strip()
    assert compat and len(compat) <= 500

    assert len(body.splitlines()) < 500
    assert 'source-documents/' not in skill

    for ref_path in sorted((root / 'references').glob('*.md')):
        ref_text = ref_path.read_text(encoding='utf-8')
        assert 'source-documents/' not in ref_text, ref_path.name
        lines = ref_text.splitlines()
        if len(lines) > 100:
            assert any(line.strip() == '## Mục lục' for line in lines[:30]), ref_path.name


def test_packaged_zip_cli_runs_standalone(tmp_path):
    import subprocess
    import sys
    from zipfile import ZipFile
    from package_skill import package_skill

    root = Path(__file__).resolve().parents[1]
    zip_path = tmp_path / 'vietnamese-docx.zip'
    result = package_skill(root, zip_path)
    assert result['skill_name'] == 'vietnamese-docx'

    with ZipFile(zip_path) as archive:
        names = archive.namelist()
        assert all(name.startswith('vietnamese-docx/') for name in names)
        assert 'vietnamese-docx/SKILL.md' in names
        assert 'vietnamese-docx/scripts/cli.py' in names
        assert 'vietnamese-docx/README.md' not in names
        assert 'vietnamese-docx/pyproject.toml' not in names
        assert not any('source-documents' in name for name in names)
        archive.extractall(tmp_path / 'unpacked')

    skill_dir = tmp_path / 'unpacked/vietnamese-docx'
    help_proc = subprocess.run(
        [sys.executable, 'scripts/cli.py', '--help'],
        cwd=skill_dir,
        capture_output=True,
        text=True,
        check=False,
    )
    assert help_proc.returncode == 0, help_proc.stderr
    assert 'Vietnamese DOCX' in help_proc.stdout

    doctor_proc = subprocess.run(
        [sys.executable, 'scripts/cli.py', 'doctor'],
        cwd=skill_dir,
        capture_output=True,
        text=True,
        check=False,
    )
    assert doctor_proc.returncode in (0, 2), doctor_proc.stderr
    doctor_data = json.loads(doctor_proc.stdout)
    assert doctor_data['status'] in ('ok', 'partial')


def test_cli_handles_vietnamese_stdout_under_cp1252(tmp_path):
    import os
    import subprocess
    import sys

    root = Path(__file__).resolve().parents[1]
    tpl_dir = root / 'assets/templates'
    values = json.loads((tpl_dir / 'party-resolution.values.json').read_text(encoding='utf-8-sig'))
    filled = tmp_path / 'party-filled.docx'
    fill_template(tpl_dir / 'party-resolution.docx', filled, values)

    env = {**os.environ, 'PYTHONIOENCODING': 'cp1252'}
    rules = str(tpl_dir / 'party-resolution.rules.json')
    for cmd in (
        [sys.executable, 'scripts/cli.py', 'validate', str(filled), '--profile', 'party', '--config', rules],
        [sys.executable, 'scripts/validate_docx.py', str(filled), '--profile', 'party', '--config', rules],
    ):
        proc = subprocess.run(cmd, cwd=root, env=env, capture_output=True, check=False)
        stderr_text = proc.stderr.decode('utf-8', errors='replace')
        assert proc.returncode == 0, stderr_text
        assert 'UnicodeEncodeError' not in stderr_text
        data = json.loads(proc.stdout.decode('utf-8'))
        assert data['status'] == 'pass'


def test_validation_benchmark_matches_published_json():
    import importlib.util
    root = Path(__file__).resolve().parents[1]
    spec = importlib.util.spec_from_file_location(
        'run_validation_benchmark', root / 'evals/run_validation_benchmark.py')
    bench = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(bench)
    summary = bench.run_benchmark()
    assert (summary['baseline_pass'], summary['baseline_total']) == (0, 8)
    assert (summary['skill_pass'], summary['skill_total']) == (8, 8)
    published = json.loads((root / 'evals/benchmark-results.json').read_text(encoding='utf-8'))
    assert summary == published


def test_inspect_docx_reports_structure_placeholders_and_eth_encoding(tmp_path):
    from inspect_docx import inspect_file
    doc = new_doc('administrative')
    doc.add_paragraph('\u00d0ƠN VỊ {{organization}} [CẦN BỔ SUNG: số] kính giửi .....')
    path = saved(tmp_path, doc, 'dirty.docx')
    report = inspect_file(path)
    assert len(report['sha256']) == 64
    assert report['sections'][0]['orientation'] == 'portrait'
    assert report['placeholders']['by_type']['token'] == ['{{organization}}']
    assert '[CẦN BỔ SUNG: số]' in report['placeholders']['by_type']['draft_marker']
    assert '.....' in report['placeholders']['by_type']['dot_fill']
    issue_codes = {item['code'] for item in report['encoding_issues']}
    assert {'eth_uppercase', 'typo'} <= issue_codes


def test_inspect_docx_normalize_encoding_is_lossless_and_rejects_signed(tmp_path):
    from zipfile import ZipFile
    from inspect_docx import inspect_file, normalize_encoding
    doc = new_doc('office')
    doc.add_paragraph('\u00d0ơn vị thực hiện \u00f0úng quy trình.')
    raw_docx = saved(tmp_path, doc, 'eth.docx')
    with_opaque = tmp_path / 'with-opaque.docx'
    opaque_bytes = b'<custom>keep-exact-bytes</custom>'
    with ZipFile(raw_docx) as zin, ZipFile(with_opaque, 'w') as zout:
        for info in zin.infolist():
            zout.writestr(info, zin.read(info))
        zout.writestr('word/custom-opaque.xml', opaque_bytes)
    normalized = tmp_path / 'normalized.docx'
    result = normalize_encoding(with_opaque, normalized)
    assert result['replacements'] == {'U+00D0->U+0110': 1, 'U+00F0->U+0111': 1}
    assert result['total_replacements'] == 2
    with ZipFile(normalized) as zcheck:
        assert zcheck.read('word/custom-opaque.xml') == opaque_bytes
    assert not inspect_file(normalized)['encoding_issues']

    signed = tmp_path / 'signed.docx'
    with ZipFile(raw_docx) as zin, ZipFile(signed, 'w') as zout:
        for info in zin.infolist():
            zout.writestr(info, zin.read(info))
        zout.writestr('_xmlsignatures/sig1.xml', b'<sig/>')
    with pytest.raises(ValueError, match='signature'):
        normalize_encoding(signed, tmp_path / 'blocked.docx')


def test_generate_docx_examples_pass_validation(tmp_path):
    from generate_docx import generate_from_spec
    root = Path(__file__).resolve().parents[1] / 'assets/schemas/examples'

    dispatch_spec = json.loads((root / 'nd30-dispatch.json').read_text(encoding='utf-8'))
    dispatch_out = generate_from_spec(dispatch_spec, tmp_path / 'dispatch.docx')
    assert inspect_doc(dispatch_out, 'administrative')['status'] == 'pass'
    dispatch_text = '\n'.join(p.text for p in Document(dispatch_out).paragraphs
                              + [p for t in Document(dispatch_out).tables for r in t.rows for c in r.cells for p in c.paragraphs])
    assert 'V/v ' in dispatch_text and 'Kính gửi:' in dispatch_text
    assert 'CÔNG VĂN' not in dispatch_text

    party_spec = json.loads((root / 'party-resolution.json').read_text(encoding='utf-8'))
    party_out = generate_from_spec(party_spec, tmp_path / 'party.docx')
    assert inspect_doc(party_out, 'party')['status'] == 'pass'
    party_text = '\n'.join(p.text for p in Document(party_out).paragraphs
                           + [p for t in Document(party_out).tables for r in t.rows for c in r.cells for p in c.paragraphs])
    assert 'ĐẢNG CỘNG SẢN VIỆT NAM' in party_text and 'Số 01-NQ/CB' in party_text
    assert 'Số:' not in party_text
    assert 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM' not in party_text

    acad_spec = json.loads((root / 'academic-report.json').read_text(encoding='utf-8'))
    acad_out = generate_from_spec(acad_spec, tmp_path / 'academic.docx')
    assert inspect_doc(acad_out, 'academic')['status'] == 'pass'
    acad_text = '\n'.join(p.text for p in Document(acad_out).paragraphs)
    assert 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM' not in acad_text
    assert 'ĐẢNG CỘNG SẢN VIỆT NAM' not in acad_text


def test_generate_docx_cli_json_string_validate_and_vbqppl(tmp_path):
    import generate_docx
    root = Path(__file__).resolve().parents[1]
    out = tmp_path / 'vbqppl.docx'
    spec_str = json.dumps({
        'profile': 'administrative',
        'header': {
            'organization': 'ỦY BAN NHÂN DÂN TỈNH THỬ NGHIỆM',
            'number': '01/2026/QĐ-UBND',
            'date': 'Hà Nội, ngày 22 tháng 11 năm 2026',
        },
        'title': 'QUYẾT ĐỊNH',
        'subtitle': 'Ban hành Quy định thử nghiệm (giả định)',
        'body': [{'type': 'paragraph', 'text': 'Điều 1. Phạm vi điều chỉnh thử nghiệm.'}],
    }, ensure_ascii=False)
    code = generate_docx.main([
        '--json-string', spec_str,
        '-o', str(out),
        '--config', str(root / 'assets/rules/legal-normative.json'),
        '--validate',
    ])
    assert code == 0
    assert inspect_doc(out, 'administrative', root / 'assets/rules/legal-normative.json')['status'] == 'pass'
    with pytest.raises(ValueError, match='\\.docx'):
        generate_docx.generate_from_spec({'body': [{'type': 'paragraph', 'text': 'x'}]}, tmp_path / 'bad.txt')


def test_cli_inspect_and_generate_dispatch(tmp_path):
    import cli
    root = Path(__file__).resolve().parents[1]
    out = tmp_path / 'via-cli.docx'
    assert cli.main([
        'generate',
        str(root / 'assets/schemas/examples/academic-report.json'),
        str(out),
        '--validate',
    ]) == 0
    assert cli.main(['inspect', str(out)]) == 0


def test_new_starters_keep_document_families_strictly_separated(tmp_path):
    from create_templates import BUILDERS

    def _all_text(doc):
        parts = [p.text for p in doc.paragraphs]
        parts += [p.text for t in doc.tables for r in t.rows for c in r.cells for p in c.paragraphs]
        return '\n'.join(parts)

    d_doc, _, d_vals = BUILDERS['administrative-dispatch']()
    d_filled = tmp_path / 'dispatch.filled.docx'
    fill_template(saved(tmp_path, d_doc, 'dispatch.docx'), d_filled, d_vals)
    d_text = _all_text(Document(d_filled))
    assert 'V/v ' in d_text and 'Kính gửi:' in d_text
    assert 'CÔNG VĂN' not in d_text

    p_doc, _, p_vals = BUILDERS['party-resolution']()
    p_filled = tmp_path / 'party.filled.docx'
    fill_template(saved(tmp_path, p_doc, 'party.docx'), p_filled, p_vals)
    p_loaded = Document(p_filled)
    p_text = _all_text(p_loaded)
    assert 'ĐẢNG CỘNG SẢN VIỆT NAM' in p_text
    assert 'Số 01-NQ/CB' in p_text and 'Số:' not in p_text
    assert 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM' not in p_text
    rec_label = p_loaded.tables[1].cell(0, 0).paragraphs[0]
    assert rec_label.text == 'Nơi nhận:' and rec_label.runs[0].underline is True
    pos_style = p_loaded.tables[1].cell(0, 1).paragraphs[1].style.name
    assert pos_style == 'PartyPosition' and not p_loaded.styles[pos_style].font.bold

    for tid in ('exam-paper', 'thesis-research-outline'):
        doc, _, vals = BUILDERS[tid]()
        filled = tmp_path / f'{tid}.filled.docx'
        fill_template(saved(tmp_path, doc, f'{tid}.docx'), filled, vals)
        text = _all_text(Document(filled))
        assert 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM' not in text
        assert 'ĐẢNG CỘNG SẢN VIỆT NAM' not in text


def test_manifest_matches_all_packaged_starters():
    import hashlib
    tpl_dir = Path(__file__).resolve().parents[1] / 'assets/templates'
    manifest = json.loads((tpl_dir / 'manifest.json').read_text(encoding='utf-8-sig'))
    ids = [item['id'] for item in manifest]
    assert set(ids) == {
        'administrative-notice',
        'administrative-dispatch',
        'party-resolution',
        'exam-paper',
        'thesis-research-outline',
        'academic-short-report',
        'office-minutes',
    }
    for item in manifest:
        assert item['visual_trial'] == 'reviewed'
        actual = hashlib.sha256((tpl_dir / item['template']).read_bytes()).hexdigest()
        assert actual == item['sha256'], item['id']
