"""Inspect DOCX properties, placeholders, and encoding issues, with optional lossless fix."""
import argparse
import hashlib
import json
import re
from pathlib import Path
from zipfile import BadZipFile, ZipFile

from docx import Document

try:
    from . import docx_package, validate_docx
except ImportError:
    import docx_package
    import validate_docx

RE_TOKEN = re.compile(r'\{\{[^{}]+\}\}')
RE_DRAFT = re.compile(r'\[CẦN BỔ SUNG[^\]]*\]', re.I)
RE_LEGACY = re.compile(r'\?\?\?|<[^<>]{2,40}>')
RE_DOTS = re.compile(r'\.{5,}|…{2,}')
RE_TYPO_KINH_GUI = re.compile(r'kính\s+giửi', re.I)
RE_HAS_LETTER = re.compile(r'[A-Za-zÀ-ỹĐđ]')


def _extract_placeholders(text):
    tokens = RE_TOKEN.findall(text)
    drafts = RE_DRAFT.findall(text)
    legacies = [m for m in RE_LEGACY.findall(text) if RE_HAS_LETTER.search(m) or m == '???']
    dots = RE_DOTS.findall(text)
    return {'token': tokens, 'draft_marker': drafts, 'legacy_marker': legacies, 'dot_fill': dots}


def inspect_file(filepath):
    """Inspect section geometry, styles, resolved fonts/colors, placeholders, and encoding."""
    path = Path(filepath)
    if not path.is_file():
        raise FileNotFoundError(f'DOCX file not found: {filepath}')
    raw = path.read_bytes()
    sha256 = hashlib.sha256(raw).hexdigest()
    try:
        with ZipFile(path) as archive:
            has_sig = any(entry.filename.startswith('_xmlsignatures/') for entry in archive.infolist())
    except BadZipFile as exc:
        raise ValueError(f'Unreadable or corrupt DOCX archive: {filepath}') from exc

    doc = Document(path)
    sections = []
    for idx, section in enumerate(doc.sections, start=1):
        w = round(section.page_width.cm, 2)
        h = round(section.page_height.cm, 2)
        hd = section.header_distance
        sections.append({
            'index': idx,
            'width_cm': w,
            'height_cm': h,
            'orientation': 'landscape' if w > h else 'portrait',
            'margins_cm': {
                'left': round(section.left_margin.cm, 2),
                'right': round(section.right_margin.cm, 2),
                'top': round(section.top_margin.cm, 2),
                'bottom': round(section.bottom_margin.cm, 2),
            },
            'header_distance_cm': round(hd.cm, 2) if hd is not None else None,
        })

    styles_used = set()
    fonts_detected = set()
    colors_detected = set()
    by_type = {'token': set(), 'draft_marker': set(), 'legacy_marker': set(), 'dot_fill': set()}
    placeholder_count = 0
    encoding_issues = []

    for story_name, part, element in validate_docx.story_parts(doc):
        for p_idx, node in enumerate(element.iter(validate_docx.qn('w:p')), start=1):
            paragraph = validate_docx.Paragraph(node, part)
            if paragraph.style is not None and paragraph.style.name:
                styles_used.add(paragraph.style.name)
            text = ''.join(t.text or '' for t in node.iter(validate_docx.qn('w:t')))
            location = f'{story_name}/paragraph {p_idx}'

            found_ph = _extract_placeholders(text)
            for kind, matches in found_ph.items():
                if matches:
                    placeholder_count += len(matches)
                    by_type[kind].update(matches)

            eth_upper = text.count('\u00d0')
            if eth_upper:
                encoding_issues.append({
                    'code': 'eth_uppercase', 'location': location,
                    'found': 'Ð (U+00D0)', 'suggestion': 'Đ (U+0110)', 'count': eth_upper,
                })
            eth_lower = text.count('\u00f0')
            if eth_lower:
                encoding_issues.append({
                    'code': 'eth_lowercase', 'location': location,
                    'found': 'ð (U+00F0)', 'suggestion': 'đ (U+0111)', 'count': eth_lower,
                })
            typos = RE_TYPO_KINH_GUI.findall(text)
            if typos:
                encoding_issues.append({
                    'code': 'typo', 'location': location,
                    'found': typos[0], 'suggestion': 'kính gửi', 'count': len(typos),
                })

            for run_node in node.iter(validate_docx.qn('w:r')):
                if not any(t.text for t in run_node.iter(validate_docx.qn('w:t'))):
                    continue
                run = validate_docx.Run(run_node, paragraph)
                font_val, font_ok = validate_docx.effective_property(run, paragraph, 'font', doc)
                if font_ok and font_val:
                    fonts_detected.add(font_val)
                color_val, color_ok = validate_docx.effective_property(run, paragraph, 'color', doc)
                if color_ok and color_val:
                    colors_detected.add(color_val)

    unique_ph = sorted({item for values in by_type.values() for item in values})
    return {
        'file': str(path),
        'sha256': sha256,
        'has_digital_signature': has_sig,
        'sections': sections,
        'styles_used': sorted(styles_used),
        'fonts_detected': sorted(fonts_detected),
        'colors_detected': sorted(colors_detected),
        'placeholders': {
            'count': placeholder_count,
            'unique': unique_ph,
            'by_type': {k: sorted(v) for k, v in by_type.items()},
        },
        'encoding_issues': encoding_issues,
    }


def normalize_encoding(source, destination):
    """Losslessly fix U+00D0 ('Ð') -> U+0110 ('Đ') and U+00F0 ('ð') -> U+0111 ('đ') in XML stories."""
    entries, original = docx_package.read_package(source)
    changed = {}
    upper_count = 0
    lower_count = 0
    t_tag = f'{{{docx_package.W}}}t'
    for name, root in docx_package.stories(original):
        modified = False
        for node in root.iter(t_tag):
            text = node.text
            if not text:
                continue
            cu = text.count('\u00d0')
            cl = text.count('\u00f0')
            if cu or cl:
                upper_count += cu
                lower_count += cl
                node.text = text.replace('\u00d0', '\u0110').replace('\u00f0', '\u0111')
                modified = True
        if modified:
            changed[name] = docx_package.xml_bytes(root)
    result = docx_package.write_package(source, destination, entries, original, changed)
    return {
        **result,
        'replacements': {'U+00D0->U+0110': upper_count, 'U+00F0->U+0111': lower_count},
        'total_replacements': upper_count + lower_count,
    }


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('filepath', help='DOCX file to inspect')
    parser.add_argument('--normalize-output', help='Write losslessly normalized DOCX (fixes Ð/ð encoding) to path')
    parser.add_argument('--report', help='Write JSON inspection report to file')
    args = parser.parse_args(argv)
    try:
        report = inspect_file(args.filepath)
        if args.normalize_output:
            report['normalization'] = normalize_encoding(args.filepath, args.normalize_output)
    except (ValueError, OSError) as exc:
        parser.error(str(exc))
    output = json.dumps(report, ensure_ascii=False, indent=2)
    print(output)
    if args.report:
        Path(args.report).write_text(output + '\n', encoding='utf-8')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
