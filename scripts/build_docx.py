"""Reusable DOCX helpers. Template edits preserve existing formatting by default."""
import re
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
try:
    from .document_rules import load_rules, section_rules, ADMIN
    from . import ooxml_order
except ImportError:
    from document_rules import load_rules, section_rules, ADMIN
    import ooxml_order

BLACK = RGBColor(0, 0, 0)
DXA_PER_CM = 1440.0 / 2.54
_CENTER_COL_HDR_RE = re.compile(
    r'^(?:s?tt|mã(?:\s+\w+){0,2}|thời\s*(?:gian|hạn|điểm)|hạn(?:\s+\w+){0,2}|'
    r'ngày(?:\s+\w+){0,2}|năm|tỷ\s*lệ|điểm(?:\s+\w+){0,2}|số\s*lượng|quy\s*mô(?:\s+\w+)?)$',
    re.IGNORECASE,
)
_NARROW_COL_HDR_RE = re.compile(r'^(?:s?tt|mã(?:\s+\w+)?)$', re.IGNORECASE)
_CODE_VAL_RE = re.compile(
    r'^(?:\d+(?:[.,]\d+)?%?|(?:\d{1,2}/){1,2}\d{2,4}|[A-ZĐ]{2,6}-\d{1,4}(?:-\d{1,4})?|'
    r'câu\s*\d+|CLO\d+(?:\s*,\s*CLO\d+)*|\d+(?:[.,]\d+)?\s*(?:điểm|trang|hồ\s*sơ|ca|lượt))$',
    re.IGNORECASE,
)


def _w_el(tag, attrs=None):
    el = OxmlElement(f'w:{tag}')
    for key, val in (attrs or {}).items():
        el.set(qn(f'w:{key}'), str(val))
    return el


def _ensure_pr(parent, tag):
    node = parent.find(qn(f'w:{tag}'))
    if node is None:
        node = _w_el(tag)
        ooxml_order.insert_ordered_child(parent, node, replace_existing=False)
    return node


def _printable_width_cm(doc):
    if not doc.sections:
        return 16.5
    sec = doc.sections[-1]
    page_w = sec.page_width.cm if sec.page_width else 21.0
    left_m = sec.left_margin.cm if sec.left_margin else 2.5
    right_m = sec.right_margin.cm if sec.right_margin else 2.0
    return max(10.0, round(page_w - left_m - right_m, 2))


def _auto_col_widths_cm(headers, rows, printable_w_cm, col_widths_cm=None):
    n_cols = len(headers)
    if col_widths_cm is not None:
        if len(col_widths_cm) != n_cols:
            raise ValueError(f'col_widths_cm length ({len(col_widths_cm)}) must match headers ({n_cols})')
        vals = [float(x) for x in col_widths_cm]
        total = sum(vals)
        if total > printable_w_cm + 0.05:
            scale = printable_w_cm / total
            vals = [round(v * scale, 2) for v in vals]
        return vals

    weights = []
    min_widths = []
    for c_idx, hdr in enumerate(headers):
        h_text = str(hdr or '').strip()
        col_texts = [h_text] + [str(r[c_idx] or '').strip() for r in rows if c_idx < len(r)]
        max_len = max((len(t) for t in col_texts), default=1)
        avg_len = sum(len(t) for t in col_texts) / max(1, len(col_texts))
        longest_word = max((len(tok) for t in col_texts for tok in t.split()), default=1)

        if _NARROW_COL_HDR_RE.match(h_text) and max_len <= 4:
            min_w = 1.15
            weight = 1.15
        elif _NARROW_COL_HDR_RE.match(h_text) and max_len <= 12:
            min_w = 2.1
            weight = 2.2
        else:
            # Ensure longest unbreakable token fits without mid-word splitting, while weighting by text volume
            word_need_cm = min(4.2, max(1.6, longest_word * 0.235 + 0.55))
            cap_cm = printable_w_cm / max(2, n_cols) * (1.3 if longest_word >= 10 else 0.92)
            min_w = max(1.6, min(word_need_cm, cap_cm))
            effective_len = 0.65 * max_len + 0.35 * avg_len
            weight = max(min_w, (effective_len ** 0.58) * 0.75)
        min_widths.append(min_w)
        weights.append(weight)

    min_sum = sum(min_widths)
    if min_sum >= printable_w_cm:
        scale = printable_w_cm / min_sum
        return [round(w * scale, 3) for w in min_widths]

    rem = printable_w_cm - min_sum
    extra_weights = [max(0.05, w - mw * 0.6) for w, mw in zip(weights, min_widths)]
    ew_sum = sum(extra_weights)
    widths = [mw + rem * (ew / ew_sum) for mw, ew in zip(min_widths, extra_weights)]
    rounded = [round(w, 2) for w in widths]
    rounded[-1] = round(printable_w_cm - sum(rounded[:-1]), 2)
    return rounded


def _column_alignment(c_idx, hdr_text, rows):
    h = str(hdr_text or '').strip()
    if _CENTER_COL_HDR_RE.match(h):
        return WD_ALIGN_PARAGRAPH.CENTER
    vals = [str(r[c_idx] or '').strip() for r in rows if c_idx < len(r) and str(r[c_idx] or '').strip()]
    if vals and all(len(v) <= 18 and _CODE_VAL_RE.match(v) for v in vals):
        return WD_ALIGN_PARAGRAPH.CENTER
    return WD_ALIGN_PARAGRAPH.LEFT


def setup_document(doc, profile='administrative', config=None):
    """Apply explicit configuration to a new document, never to a retained template."""
    rules = load_rules(profile, config)
    if profile == 'custom' and config is None:
        return doc
    if rules['sections'] and len(rules['sections']) != len(doc.sections):
        raise ValueError('Create the configured sections before applying setup_document')
    for index, section in enumerate(doc.sections):
        spec = section_rules(rules, index)
        if spec['page_cm']:
            section.page_width, section.page_height = map(Cm, spec['page_cm'])
            section.orientation = (WD_ORIENT.LANDSCAPE if spec['page_cm'][0] > spec['page_cm'][1]
                                   else WD_ORIENT.PORTRAIT)
        margins = spec['margins_cm']
        if not margins and profile != 'custom':
            margins = {side: [value, value] for side, value in
                       {'left': 2.5, 'right': 2.0, 'top': 2.5, 'bottom': 2.5}.items()}
        for side, bounds in margins.items():
            setattr(section, side + '_margin', Cm(bounds[0]))
    setup_styles(doc, rules)
    if profile in ADMIN:
        add_page_numbers(doc, position='header', show_first=False,
                         size=rules['page_number_size_pt'])
        if profile == 'party':
            for section in doc.sections:
                section.header_distance = Cm(1.0)
    return doc


def setup_styles(doc, rules=None):
    rules = rules or load_rules()
    names = ['Normal', 'Title', 'Subtitle', 'Heading 1', 'Heading 2', 'Heading 3',
             'List Bullet', 'List Bullet 2', 'List Bullet 3', 'List Number']
    for name in names:
        if name not in doc.styles:
            continue
        style = doc.styles[name]
        if rules['font']:
            style.font.name = rules['font']
            fonts = style.element.get_or_add_rPr().get_or_add_rFonts()
            for key in ('asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'cstheme'):
                fonts.attrib.pop(qn('w:' + key), None)
            for key in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
                fonts.set(qn('w:' + key), rules['font'])
        if rules['color']:
            style.font.color.rgb = RGBColor.from_string(rules['color'])
            color = style.element.rPr.find(qn('w:color'))
            for key in ('themeColor', 'themeTint', 'themeShade'):
                color.attrib.pop(qn('w:' + key), None)
        style.font.size = Pt(rules['body_size_pt'])
        if name == 'Title':
            style.font.size = Pt(rules['body_size_pt'] + 3)
            style.font.bold = True
            if style.element.pPr is not None:
                border = style.element.pPr.find(qn('w:pBdr'))
                if border is not None:
                    style.element.pPr.remove(border)
        if name.startswith('Heading'):
            style.font.bold = True
            style.paragraph_format.keep_with_next = True
            style.paragraph_format.space_before = Pt(12)
            style.paragraph_format.space_after = Pt(6)
    normal = doc.styles['Normal'].paragraph_format
    normal.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    normal.space_after = Pt(6)
    normal.line_spacing = (Pt(rules['line_spacing_exact_pt'])
                           if rules['line_spacing_exact_pt'] else rules['line_spacing'])
    if rules['profile'] in ADMIN:
        normal.first_line_indent = Cm(1)
    for name, spec in rules['style_rules'].items():
        style = doc.styles[name] if name in doc.styles else doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        if style.base_style is None and name != 'Normal':
            style.base_style = doc.styles['Normal']
        if 'font' in spec:
            style.font.name = spec['font']
            fonts = style.element.get_or_add_rPr().get_or_add_rFonts()
            for key in ('asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'cstheme'):
                fonts.attrib.pop(qn('w:' + key), None)
            for key in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
                fonts.set(qn('w:' + key), spec['font'])
        if 'color' in spec:
            style.font.color.rgb = RGBColor.from_string(spec['color'])
            color = style.element.rPr.find(qn('w:color'))
            for key in ('themeColor', 'themeTint', 'themeShade'):
                color.attrib.pop(qn('w:' + key), None)
        if 'size_pt' in spec:
            style.font.size = Pt(spec['size_pt'][0])
        for prop in ('bold', 'italic'):
            if prop in spec:
                setattr(style.font, prop, spec[prop])


def add_page_numbers(doc, position='footer', show_first=True, size=13):
    """Add PAGE fields to a new document. Never use as a template reset operation."""
    if position not in ('header', 'footer'):
        raise ValueError('Page number position must be header or footer')
    seen = set()
    for index, section in enumerate(doc.sections):
        section.different_first_page_header_footer = index == 0 and not show_first
        story = getattr(section, position)
        part = story.part
        if str(part.partname) in seen:
            continue
        seen.add(str(part.partname))
        if any('PAGE' in (node.text or '').split() for node in story._element.iter(qn('w:instrText'))):
            continue
        p = story.paragraphs[0] if len(story.paragraphs) == 1 and not story.paragraphs[0].text else story.add_paragraph()
        p.style = doc.styles['Normal']
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        run = p.add_run()
        run.font.size, run.bold, run.italic = Pt(size), False, False
        for kind, text in [('begin', None), (None, ' PAGE '), ('separate', None), (None, '1'), ('end', None)]:
            node = OxmlElement('w:fldChar' if kind else ('w:instrText' if text.strip() == 'PAGE' else 'w:t'))
            if kind:
                node.set(qn('w:fldCharType'), kind)
            else:
                node.text = text
                node.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            run._r.append(node)
    settings = doc.settings.element
    update = settings.find(qn('w:updateFields'))
    if update is None:
        update = OxmlElement('w:updateFields')
        settings.append(update)
    update.set(qn('w:val'), 'true')


def add_body(doc, text):
    return doc.add_paragraph(text, style='Normal')


def add_blank(doc):
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    return paragraph


def add_centered(doc, text, size=None, bold=True, italic=False):
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.first_line_indent = Cm(0)
    run = paragraph.add_run(text)
    if size is not None:
        run.font.size = Pt(size)
    run.bold, run.italic = bold, italic
    return paragraph


def add_right(doc, text, bold=False, italic=False, size=None):
    paragraph = add_centered(doc, text, size, bold, italic)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    return paragraph


def add_italic(doc, text):
    paragraph = add_body(doc, text)
    for run in paragraph.runs:
        run.italic = True
    return paragraph


def add_bullet(doc, text, level=1, profile='administrative'):
    if level not in (1, 2, 3):
        raise ValueError('Bullet level must be 1, 2 or 3')
    if profile in ADMIN:
        paragraph = doc.add_paragraph({1: '- ', 2: '+ ', 3: '* '}[level] + text)
        paragraph.paragraph_format.left_indent = Cm((level-1) * 1.0)
        paragraph.paragraph_format.first_line_indent = Cm(0)
        return paragraph
    return doc.add_paragraph(text, style='List Bullet' if level == 1 else f'List Bullet {level}')


def add_bullets(doc, items, level=1, profile='administrative'):
    return [add_bullet(doc, text, level, profile) for text in items]


def add_numbered(doc, items, start=1):
    """Create a real Word decimal list with explicit start, independent of previous lists."""
    if not isinstance(start, int) or start < 1:
        raise ValueError('start must be a positive integer')
    numbering = doc.part.numbering_part.element
    ids = [int(node.get(qn('w:abstractNumId'))) for node in numbering.findall(qn('w:abstractNum'))]
    abstract_id = max(ids, default=-1) + 1
    abstract = OxmlElement('w:abstractNum')
    abstract.set(qn('w:abstractNumId'), str(abstract_id))
    level = OxmlElement('w:lvl')
    level.set(qn('w:ilvl'), '0')
    for tag, value in [('start', str(start)), ('numFmt', 'decimal'), ('lvlText', '%1.'), ('lvlJc', 'left')]:
        child = OxmlElement('w:' + tag)
        child.set(qn('w:val'), value)
        level.append(child)
    props = OxmlElement('w:pPr')
    indent = OxmlElement('w:ind')
    indent.set(qn('w:left'), '567')
    indent.set(qn('w:hanging'), '283')
    props.append(indent)
    level.append(props)
    abstract.append(level)
    numbering.append(abstract)
    num = numbering.add_num(abstract_id)
    result = []
    for text in items:
        paragraph = doc.add_paragraph(text)
        numbering_props = paragraph._p.get_or_add_pPr().get_or_add_numPr()
        numbering_props.get_or_add_ilvl().val = 0
        numbering_props.get_or_add_numId().val = num.numId
        result.append(paragraph)
    return result


def set_table_grid_style(table, profile='administrative'):
    table.style = 'Table Grid'


def _format_table_cell_paragraph(paragraph, align, size_pt=None, bold=None):
    paragraph.alignment = align
    pf = paragraph.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.space_before = Pt(3)
    pf.space_after = Pt(3)
    pf.line_spacing = 1.1
    if size_pt is not None or bold is not None:
        for run in paragraph.runs:
            if size_pt is not None:
                run.font.size = Pt(size_pt)
            if bold is not None:
                run.bold = bold


def _apply_table_layout_xml(table, col_widths_cm):
    cols_dxa = [max(288, round(w * DXA_PER_CM)) for w in col_widths_cm]
    total_dxa = sum(cols_dxa)
    tbl = table._tbl
    tbl_pr = _ensure_pr(tbl, 'tblPr')
    ooxml_order.insert_ordered_child(tbl_pr, _w_el('tblW', {'w': total_dxa, 'type': 'dxa'}))
    ooxml_order.insert_ordered_child(tbl_pr, _w_el('jc', {'val': 'center'}))
    ooxml_order.insert_ordered_child(tbl_pr, _w_el('tblLayout', {'type': 'fixed'}))

    cell_mar = tbl_pr.find(qn('w:tblCellMar'))
    if cell_mar is None:
        cell_mar = _w_el('tblCellMar')
        ooxml_order.insert_ordered_child(tbl_pr, cell_mar, replace_existing=False)
    for side, dxa in (('top', 80), ('left', 108), ('bottom', 80), ('right', 108)):
        ooxml_order.insert_ordered_child(cell_mar, _w_el(side, {'w': dxa, 'type': 'dxa'}))
    ooxml_order.normalize_ordered_children(cell_mar)
    ooxml_order.normalize_ordered_children(tbl_pr)

    tbl_grid = tbl.find(qn('w:tblGrid'))
    if tbl_grid is None:
        tbl_grid = _w_el('tblGrid')
        tbl_pr.addnext(tbl_grid)
    for child in list(tbl_grid):
        tbl_grid.remove(child)
    for w_dxa in cols_dxa:
        tbl_grid.append(_w_el('gridCol', {'w': w_dxa}))

    for r_idx, row in enumerate(table.rows):
        tr_pr = _ensure_pr(row._tr, 'trPr')
        if r_idx == 0:
            ooxml_order.insert_ordered_child(tr_pr, _w_el('tblHeader'))
        ooxml_order.insert_ordered_child(tr_pr, _w_el('cantSplit'))
        ooxml_order.normalize_ordered_children(tr_pr)
        for c_idx, cell in enumerate(row.cells):
            w_dxa = cols_dxa[min(c_idx, len(cols_dxa) - 1)]
            tc_pr = _ensure_pr(cell._tc, 'tcPr')
            ooxml_order.insert_ordered_child(tc_pr, _w_el('tcW', {'w': w_dxa, 'type': 'dxa'}))
            ooxml_order.insert_ordered_child(tc_pr, _w_el('vAlign', {'val': 'center' if r_idx == 0 else 'top'}))
            ooxml_order.normalize_ordered_children(tc_pr)


def add_table(doc, headers, rows, profile='administrative', col_widths_cm=None, font_size_pt=None):
    if not headers or any(len(row) != len(headers) for row in rows):
        raise ValueError('Table rows must match the nonempty header')
    table = doc.add_table(rows=1, cols=len(headers))
    table.autofit = False
    set_table_grid_style(table, profile)

    printable_w_cm = _printable_width_cm(doc)
    widths_cm = _auto_col_widths_cm(headers, rows, printable_w_cm, col_widths_cm=col_widths_cm)
    col_aligns = [_column_alignment(i, headers[i], rows) for i in range(len(headers))]

    for cell, text in zip(table.rows[0].cells, headers):
        paragraph = cell.paragraphs[0]
        paragraph.add_run(str(text)).bold = True
        _format_table_cell_paragraph(paragraph, WD_ALIGN_PARAGRAPH.CENTER, size_pt=font_size_pt, bold=True)

    for values in rows:
        row_cells = table.add_row().cells
        for c_idx, (cell, value) in enumerate(zip(row_cells, values)):
            paragraph = cell.paragraphs[0]
            paragraph.add_run(str(value))
            _format_table_cell_paragraph(paragraph, col_aligns[c_idx], size_pt=font_size_pt)

    _apply_table_layout_xml(table, widths_cm)
    return table


def add_cover_page(doc, org_name, title, subtitle, team_data, location=None, year=None):
    add_centered(doc, org_name)
    doc.add_paragraph(title, style='Title')
    if subtitle:
        add_centered(doc, subtitle, bold=False)
    if team_data:
        add_table(doc, ['Thông tin', 'Nội dung'], team_data, 'academic')
    if location or year:
        add_centered(doc, ', '.join(str(x) for x in (location, year) if x), bold=False)
    doc.add_page_break()


def add_cover_letter(doc, greeting, body_paragraphs, signer_name):
    doc.add_paragraph('THƯ NGỎ', style='Title')
    add_body(doc, greeting)
    for text in body_paragraphs:
        add_body(doc, text)
    add_right(doc, 'Trân trọng,')
    add_right(doc, signer_name, bold=True)
    doc.add_page_break()


def add_h1(doc, text):
    return doc.add_heading(text, level=1)


def add_h2(doc, text):
    return doc.add_heading(text, level=2)


def add_h3(doc, text):
    return doc.add_heading(text, level=3)
