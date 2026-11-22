"""Reusable DOCX helpers. Template edits preserve existing formatting by default."""
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
try:
    from .document_rules import load_rules, section_rules, ADMIN
except ImportError:
    from document_rules import load_rules, section_rules, ADMIN

BLACK = RGBColor(0, 0, 0)


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


def add_table(doc, headers, rows, profile='administrative'):
    if not headers or any(len(row) != len(headers) for row in rows):
        raise ValueError('Table rows must match the nonempty header')
    table = doc.add_table(rows=1, cols=len(headers))
    set_table_grid_style(table, profile)
    repeat = OxmlElement('w:tblHeader')
    table.rows[0]._tr.get_or_add_trPr().append(repeat)
    for cell, text in zip(table.rows[0].cells, headers):
        paragraph = cell.paragraphs[0]
        paragraph.paragraph_format.first_line_indent = Cm(0)
        paragraph.add_run(str(text)).bold = True
    for values in rows:
        for cell, value in zip(table.add_row().cells, values):
            paragraph = cell.paragraphs[0]
            paragraph.paragraph_format.first_line_indent = Cm(0)
            paragraph.add_run(str(value))
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
