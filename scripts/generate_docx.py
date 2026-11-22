"""Declarative multi-profile JSON-to-DOCX generator using skill helpers and validators."""
import argparse
import json
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

try:
    from . import build_docx, document_rules, validate_docx
except ImportError:
    import build_docx
    import document_rules
    import validate_docx

ALIGN_MAP = {
    'left': WD_ALIGN_PARAGRAPH.LEFT,
    'center': WD_ALIGN_PARAGRAPH.CENTER,
    'right': WD_ALIGN_PARAGRAPH.RIGHT,
    'justify': WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def _borderless(table):
    props = table._tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        node = OxmlElement('w:' + side)
        node.set(qn('w:val'), 'nil')
        borders.append(node)
    props.append(borders)


def _prevent_row_split(row):
    row._tr.get_or_add_trPr().append(OxmlElement('w:cantSplit'))


def _cell_para(container, text, *, align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=13,
               bold=False, italic=False, underline=False, before_pt=0, after_pt=0,
               indent_cm=0, line_spacing=1.0, reuse_first=False, style_name='Normal'):
    if reuse_first and len(container.paragraphs) == 1 and not container.paragraphs[0].text:
        p = container.paragraphs[0]
        p.text = ''
    else:
        p = container.add_paragraph()
    p.style = style_name
    p.alignment = align
    fmt = p.paragraph_format
    fmt.first_line_indent = Cm(indent_cm)
    fmt.space_before = Pt(before_pt)
    fmt.space_after = Pt(after_pt)
    fmt.line_spacing = line_spacing
    run = p.add_run(text)
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    if underline:
        run.underline = True
    return p


def _add_rule(container, width_cm, available_cm, space_after_pt=4):
    p = container.add_paragraph()
    p.paragraph_format.left_indent = Cm(max(0, (available_cm - width_cm) / 2))
    p.paragraph_format.right_indent = p.paragraph_format.left_indent
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(space_after_pt)
    p.paragraph_format.line_spacing = Pt(6)
    p.add_run().font.size = Pt(1)
    borders = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    for name, value in {'val': 'single', 'sz': '6', 'space': '0', 'color': '000000'}.items():
        bottom.set(qn('w:' + name), value)
    borders.append(bottom)
    p._p.get_or_add_pPr().append(borders)
    return p


def _format_number_line(number, is_party):
    raw = (number or '').strip()
    if not raw:
        return 'Số ...' if is_party else 'Số: ...'
    if is_party:
        stripped = raw[3:].strip() if raw.startswith('Số:') else (raw[2:].strip() if raw.startswith('Số ') else raw)
        return f'Số {stripped}'
    return raw if raw.startswith('Số:') else f'Số: {raw.removeprefix("Số").lstrip(": ")}'


def _render_official_header(doc, spec, profile, body_size_pt):
    is_party = profile == 'party'
    is_dispatch = spec.get('doc_kind', 'named') == 'dispatch'
    header = spec.get('header') or {}

    top = doc.add_table(rows=2, cols=2)
    top.autofit = False
    left_w, right_w = Cm(6.8), Cm(9.7)
    top.columns[0].width, top.columns[1].width = left_w, right_w
    for row in top.rows:
        _prevent_row_split(row)
        row.cells[0].width, row.cells[1].width = left_w, right_w
    _borderless(top)

    left_top, right_top = top.rows[0].cells
    left_meta, right_meta = top.rows[1].cells

    parent_org = (header.get('parent_org') or '').strip()
    org = (header.get('organization') or '[CẦN BỔ SUNG: cơ quan ban hành]').strip()
    org_size = 14 if is_party else 13
    reused = False
    if parent_org:
        _cell_para(left_top, parent_org.upper(), size_pt=org_size, bold=False, reuse_first=True)
        reused = True
    _cell_para(left_top, org.upper(), size_pt=org_size, bold=True, reuse_first=not reused)
    if is_party:
        _cell_para(left_top, '*', size_pt=14, bold=False, after_pt=2)
        _cell_para(right_top, 'ĐẢNG CỘNG SẢN VIỆT NAM', size_pt=15, bold=True, reuse_first=True)
        _add_rule(right_top, 6.8, 9.32, space_after_pt=2)
    else:
        _add_rule(left_top, 2.5, 6.42, space_after_pt=2)
        _cell_para(right_top, 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', size_pt=12, bold=True, reuse_first=True)
        _cell_para(right_top, 'Độc lập - Tự do - Hạnh phúc', size_pt=13, bold=True)
        _add_rule(right_top, 4.95, 9.32, space_after_pt=2)

    num_line = _format_number_line(header.get('number'), is_party)
    _cell_para(left_meta, num_line, size_pt=14 if is_party else 13, bold=False,
               before_pt=4, reuse_first=True)
    if is_dispatch and header.get('subject_brief'):
        _cell_para(left_meta, header['subject_brief'].strip(),
                   size_pt=12, italic=is_party, before_pt=6)

    date_text = (header.get('date') or '[CẦN BỔ SUNG: địa danh, ngày tháng năm]').strip()
    _cell_para(right_meta, date_text, size_pt=14 if is_party else body_size_pt,
               italic=True, before_pt=4, reuse_first=True)

    title = (spec.get('title') or '').strip()
    subtitle = (spec.get('subtitle') or '').strip()
    if not is_dispatch and title:
        _cell_para(doc, title.upper(), size_pt=15 if is_party else 14,
                   bold=True, before_pt=14, after_pt=2)
        if subtitle:
            _cell_para(doc, subtitle, size_pt=14 if is_party else body_size_pt,
                       bold=True, after_pt=2)
        if is_party:
            _cell_para(doc, '-----', size_pt=14, bold=False, after_pt=10)
        else:
            _add_rule(doc, 4.5, 16.5, space_after_pt=12)

    recipients_to = spec.get('recipients_to')
    if recipients_to:
        items = [recipients_to] if isinstance(recipients_to, str) else list(recipients_to)
        if len(items) == 1:
            _cell_para(doc, f'Kính gửi: {items[0]}', align=WD_ALIGN_PARAGRAPH.CENTER,
                       size_pt=14 if is_party else body_size_pt, italic=is_party,
                       before_pt=6, after_pt=10)
        else:
            _cell_para(doc, 'Kính gửi:', align=WD_ALIGN_PARAGRAPH.LEFT, indent_cm=2.0,
                       size_pt=14 if is_party else body_size_pt, italic=is_party,
                       before_pt=6, after_pt=2)
            for idx, target in enumerate(items):
                _cell_para(doc, f'- {target};' if idx < len(items) - 1 else f'- {target}.',
                           align=WD_ALIGN_PARAGRAPH.LEFT, indent_cm=3.5,
                           size_pt=14 if is_party else body_size_pt,
                           after_pt=10 if idx == len(items) - 1 else 2)

    for basis in spec.get('legal_bases') or []:
        text = basis.strip()
        if is_party and not text.startswith('- '):
            text = f'- {text}'
        _cell_para(doc, text, align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent_cm=1.0,
                   size_pt=body_size_pt, italic=not is_party, after_pt=6,
                   line_spacing=Pt(20) if is_party else 1.15)


def _render_general_header(doc, spec, body_size_pt):
    title = (spec.get('title') or '').strip()
    if title:
        p = doc.add_paragraph(title, style='Title')
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Cm(0)
    subtitle = (spec.get('subtitle') or '').strip()
    if subtitle:
        _cell_para(doc, subtitle, align=WD_ALIGN_PARAGRAPH.CENTER,
                   size_pt=body_size_pt + 1, italic=True, after_pt=6)
    for line in spec.get('metadata') or []:
        _cell_para(doc, line, align=WD_ALIGN_PARAGRAPH.LEFT,
                   size_pt=body_size_pt, after_pt=4)


def _render_body_blocks(doc, spec, profile, rules):
    is_official = profile in document_rules.ADMIN
    body_size_pt = rules['body_size_pt']
    spacing = Pt(rules['line_spacing_exact_pt']) if rules['line_spacing_exact_pt'] else rules['line_spacing']

    for block in spec['body']:
        if not isinstance(block, dict) or 'type' not in block:
            raise ValueError('Each body item must be an object with a "type" field')
        kind = block['type']
        if kind == 'heading':
            level = int(block.get('level', 1))
            if not 1 <= level <= 3:
                raise ValueError('Heading level must be 1, 2, or 3')
            text = block['text']
            if is_official:
                p = _cell_para(doc, text, align=WD_ALIGN_PARAGRAPH.LEFT,
                               size_pt=body_size_pt, bold=True, indent_cm=1.0,
                               before_pt=6, after_pt=4, line_spacing=spacing)
                p.paragraph_format.keep_with_next = True
            else:
                getattr(build_docx, f'add_h{level}')(doc, text)
        elif kind == 'paragraph':
            align_key = block.get('align', 'justify')
            if align_key not in ALIGN_MAP:
                raise ValueError(f'Unsupported alignment: {align_key}')
            _cell_para(
                doc, block['text'], align=ALIGN_MAP[align_key], size_pt=body_size_pt,
                bold=bool(block.get('bold', False)), italic=bool(block.get('italic', False)),
                indent_cm=1.0 if is_official and align_key in ('left', 'justify') else 0,
                after_pt=6, line_spacing=spacing,
            )
        elif kind == 'bullet':
            items = block.get('items') or []
            if not items:
                raise ValueError('Bullet block requires non-empty "items"')
            level = int(block.get('level', 1))
            for item in items:
                p = build_docx.add_bullet(doc, item, level=level, profile=profile)
                if rules['font']:
                    for r in p.runs:
                        r.font.size = Pt(body_size_pt)
        elif kind == 'numbered':
            items = block.get('items') or []
            if not items:
                raise ValueError('Numbered block requires non-empty "items"')
            start = int(block.get('start', 1))
            paragraphs = build_docx.add_numbered(doc, items, start=start)
            for p in paragraphs:
                for r in p.runs:
                    r.font.size = Pt(body_size_pt)
        elif kind == 'table':
            col_widths_cm = block.get('col_widths_cm')
            tbl_font_pt = block.get('font_size_pt')
            table = build_docx.add_table(
                doc, block['headers'], block['rows'], profile=profile,
                col_widths_cm=col_widths_cm, font_size_pt=tbl_font_pt,
            )
            for row in table.rows:
                _prevent_row_split(row)
        elif kind == 'page_break':
            doc.add_page_break()
        else:
            raise ValueError(f'Unsupported body block type: {kind}')


def _render_official_footer(doc, spec, profile):
    recipients = spec.get('recipients') or []
    signer = spec.get('signer')
    if not recipients and not signer:
        return
    if doc.paragraphs:
        doc.paragraphs[-1].paragraph_format.keep_with_next = True
    is_party = profile == 'party'
    bottom = doc.add_table(rows=1, cols=2)
    bottom.autofit = False
    bottom.columns[0].width, bottom.columns[1].width = Cm(7.5), Cm(9.0)
    bottom.cell(0, 0).width, bottom.cell(0, 1).width = Cm(7.5), Cm(9.0)
    _prevent_row_split(bottom.rows[0])
    _borderless(bottom)
    left, right = bottom.rows[0].cells

    if recipients:
        if is_party:
            _cell_para(left, 'Nơi nhận:', align=WD_ALIGN_PARAGRAPH.LEFT, size_pt=14,
                       bold=False, underline=True, reuse_first=True)
            for rec in recipients:
                _cell_para(left, rec, align=WD_ALIGN_PARAGRAPH.LEFT, size_pt=12)
        else:
            _cell_para(left, 'Nơi nhận:', align=WD_ALIGN_PARAGRAPH.LEFT, size_pt=12,
                       bold=True, italic=True, reuse_first=True)
            for rec in recipients:
                _cell_para(left, rec, align=WD_ALIGN_PARAGRAPH.LEFT, size_pt=11)

    if signer:
        authority = signer.get('authority_lines') or []
        position = (signer.get('position') or '').strip()
        note = (signer.get('note') or '').strip()
        name = (signer.get('name') or '').strip()
        sig_size = 14 if is_party else 13
        reused_right = False
        for auth in authority:
            _cell_para(right, auth.upper(), size_pt=sig_size, bold=True, reuse_first=not reused_right)
            reused_right = True
        if position:
            pos_bold = signer.get('position_bold', not is_party)
            _cell_para(right, position.upper(), size_pt=sig_size, bold=bool(pos_bold),
                       reuse_first=not reused_right)
            reused_right = True
        if note:
            _cell_para(right, note, size_pt=12, italic=True, before_pt=4, after_pt=42,
                       reuse_first=not reused_right)
            reused_right = True
        elif reused_right:
            _cell_para(right, '', size_pt=sig_size, after_pt=42)
        if name:
            _cell_para(right, name, size_pt=sig_size, bold=True, reuse_first=not reused_right)


def generate_from_spec(spec, output_path, config=None):
    """Build a styled DOCX document from a declarative dict spec and save to output_path."""
    if not isinstance(spec, dict):
        raise ValueError('Document spec must be a JSON object')
    out = Path(output_path)
    if out.suffix.lower() != '.docx':
        raise ValueError('Output path must end with .docx')
    profile = spec.get('profile', 'administrative')
    doc_kind = spec.get('doc_kind', 'named')
    if doc_kind not in ('named', 'dispatch'):
        raise ValueError('doc_kind must be "named" or "dispatch"')
    body = spec.get('body')
    if not isinstance(body, list) or not body:
        raise ValueError('Spec must contain a non-empty "body" list')

    rules = document_rules.load_rules(profile, config)
    doc = Document()
    build_docx.setup_document(doc, profile, config)

    if profile in document_rules.ADMIN:
        _render_official_header(doc, spec, profile, rules['body_size_pt'])
    else:
        _render_general_header(doc, spec, rules['body_size_pt'])

    _render_body_blocks(doc, spec, profile, rules)

    if profile in document_rules.ADMIN:
        _render_official_footer(doc, spec, profile)

    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out)
    return out


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input_source', nargs='?', help='JSON spec file path, or "-" to read from stdin')
    parser.add_argument('output_pos', nargs='?', help='Optional output .docx path (positional)')
    parser.add_argument('-o', '--output', dest='output_opt', help='Output .docx file path')
    parser.add_argument('--json-string', help='Inline JSON spec string instead of file input')
    parser.add_argument('--config', help='Optional JSON rules config (e.g. assets/rules/legal-normative.json)')
    parser.add_argument('--validate', action='store_true', help='Validate generated DOCX immediately')
    parser.add_argument('--draft', action='store_true', help='Allow [CẦN BỔ SUNG: ...] markers during validation')
    args = parser.parse_args(argv)

    if args.json_string and args.input_source:
        parser.error('Pass either input_source or --json-string, not both')
    if not args.json_string and not args.input_source:
        parser.error('Provide input_source (file or "-") or --json-string')

    output = args.output_opt or args.output_pos
    if not output:
        if args.input_source and args.input_source != '-':
            output = str(Path(args.input_source).with_suffix('.docx'))
        else:
            parser.error('--output (-o) is required when reading from stdin or --json-string')

    try:
        if args.json_string:
            raw = args.json_string
        elif args.input_source == '-':
            raw = sys.stdin.read()
        else:
            raw = Path(args.input_source).read_text(encoding='utf-8-sig')
        spec = json.loads(raw)
        out_path = generate_from_spec(spec, output, args.config)
    except (ValueError, OSError, json.JSONDecodeError) as exc:
        parser.error(str(exc))

    result = {'output': str(out_path), 'profile': spec.get('profile', 'administrative')}
    exit_code = 0
    if args.validate:
        report = validate_docx.inspect_doc(out_path, result['profile'], args.config, draft=args.draft)
        result['validation'] = report
        exit_code = {'pass': 0, 'fail': 1, 'partial': 2}[report['status']]
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return exit_code


if __name__ == '__main__':
    raise SystemExit(main())
