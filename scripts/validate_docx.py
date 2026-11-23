"""Inspect supported DOCX properties without claiming legal or visual compliance."""
import argparse
import json
import re
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.text.run import Run

try:
    from .document_rules import load_rules, section_rules
    from .docx_package import ensure_utf8_stdio
except ImportError:
    from document_rules import load_rules, section_rules
    from docx_package import ensure_utf8_stdio

PLACEHOLDER = re.compile(r'\[CẦN BỔ SUNG[^\]]*\]|\{\{[^{}]+\}\}', re.I)


def story_parts(doc):
    """Return each real body/header/footer once, including linked stories."""
    yield 'body', doc.part, doc.element
    seen = set()
    for rel in doc.part.rels.values():
        if not rel.is_external and rel.reltype.rsplit('/', 1)[-1] in ('header', 'footer'):
            part = rel.target_part
            if part.partname not in seen:
                seen.add(part.partname)
                yield str(part.partname), part, part.element


def style_chain(style):
    seen = set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        yield style
        style = style.base_style


def effective_property(run, paragraph, name, doc):
    """Resolve direct/run-style/paragraph-style/default properties in order.

    Table conditional styles and theme font expansion remain explicitly unchecked.
    """
    candidates = [run._r.rPr]
    candidates += [s.element.rPr for s in style_chain(run.style)]
    candidates += [s.element.rPr for s in style_chain(paragraph.style)]
    defaults = doc.styles.element.find(qn('w:docDefaults'))
    if defaults is not None:
        default_run = defaults.find(qn('w:rPrDefault'))
        if default_run is not None:
            candidates.append(default_run.find(qn('w:rPr')))
    if name == 'font':
        resolved = []
        for key in ('ascii', 'hAnsi'):
            value = None
            for props in candidates:
                if props is None:
                    continue
                fonts = props.find(qn('w:rFonts'))
                if fonts is None:
                    continue
                if fonts.get(qn('w:' + key + 'Theme')):
                    break
                if fonts.get(qn('w:' + key)):
                    value = fonts.get(qn('w:' + key))
                    break
            if value is None:
                return None, False
            resolved.append(value)
        return '|'.join(sorted(set(resolved))), True
    if name in ('size', 'bold', 'italic'):
        tag = {'size': 'sz', 'bold': 'b', 'italic': 'i'}[name]
        for props in candidates:
            if props is None:
                continue
            element = props.find(qn('w:' + tag))
            if element is not None:
                value = element.get(qn('w:val'))
                if name == 'size':
                    return int(value)/2, True
                return value not in ('0', 'false', 'off'), True
        return (None, False) if name == 'size' else (False, True)
    for props in candidates:
        if props is None:
            continue
        element = props.find(qn('w:color'))
        if element is None:
            continue
        if element.get(qn('w:themeColor')):
            return None, False
        value = element.get(qn('w:val'))
        if value == 'auto':
            return None, False
        if value:
            return value.upper(), True
    return None, False


def _numbering(paragraph):
    props = [paragraph._p.pPr]
    props += [s.element.pPr for s in style_chain(paragraph.style)]
    for prop in props:
        if prop is not None:
            num = prop.find(qn('w:numPr'))
            if num is not None:
                number = num.find(qn('w:numId'))
                if number is not None:
                    return number.get(qn('w:val'))
    return None


def numbering_format(paragraph, doc):
    number = _numbering(paragraph)
    if number in (None, '0'):
        return None
    number = int(number)
    props = [paragraph._p.pPr] + [s.element.pPr for s in style_chain(paragraph.style)]
    level = 0
    for prop in props:
        if prop is not None:
            levels = prop.xpath('./w:numPr/w:ilvl')
            if levels:
                level = int(levels[0].get(qn('w:val')))
                break
    if not 0 <= level <= 8:
        raise ValueError('List level outside 0..8')
    numbering = doc.part.numbering_part.element
    nums = numbering.xpath(f'./w:num[@w:numId="{number}"]')
    if not nums:
        raise ValueError(f'Numbering id {number} has no definition')
    formats = nums[0].xpath(f'./w:lvlOverride[@w:ilvl="{level}"]/w:lvl/w:numFmt')
    if not formats:
        abstract = nums[0].find(qn('w:abstractNumId'))
        if abstract is None:
            raise ValueError('Numbering has no abstract definition')
        abstract_id = int(abstract.get(qn('w:val')))
        definitions = numbering.xpath(f'./w:abstractNum[@w:abstractNumId="{abstract_id}"]')
        if not definitions:
            raise ValueError('Abstract numbering definition is missing')
        if definitions[0].find(qn('w:numStyleLink')) is not None:
            return 'unresolved'
        levels = [node for node in definitions[0].findall(qn('w:lvl'))
                  if node.get(qn('w:ilvl')) == str(level)]
        if not levels:
            raise ValueError('Numbering level is missing')
        formats = levels[0].findall(qn('w:numFmt'))
        if not formats:
            return 'decimal'
    if not formats:
        raise ValueError('Numbering level has no format definition')
    return formats[0].get(qn('w:val'))


def inspect_doc(filepath, profile='administrative', config=None, draft=False):
    rules = load_rules(profile, config)
    report = {'profile': profile, 'source': rules['source'], 'status': 'pass',
              'issues': [], 'unchecked': [
                  'Visual layout, clipping, actual pagination and table overflow',
                  'Legal applicability, factual accuracy and signatures',
                  'Complete type-specific structure unless required_text is configured',
                  'Theme fonts/colors and conditional table styles',
                  'Text boxes, charts, footnotes/endnotes, tracked revisions and embedded objects',
                  'Number sequence correctness, cross-reference targets and field updates',
                  'Font size and emphasis outside configured named style roles; paragraph spacing',
              ]}

    def issue(code, location, message, severity='error'):
        report['issues'].append(dict(code=code, location=location,
                                     message=message, severity=severity))

    try:
        doc = Document(filepath)
    except Exception as exc:
        issue('open', str(filepath), str(exc))
        report['status'] = 'fail'
        return report

    if rules['sections'] and len(rules['sections']) != len(doc.sections):
        issue('sections', 'document', 'Section count differs from config')
    for index, section in enumerate(doc.sections):
        spec = section_rules(rules, index)
        actual = [section.page_width.cm, section.page_height.cm]
        expected = spec['page_cm']
        if expected:
            direct = all(abs(a-b) <= 0.03 for a, b in zip(actual, expected))
            rotated = (spec['allow_landscape'] and
                       all(abs(a-b) <= 0.03 for a, b in zip(actual, reversed(expected))))
            if not (direct or rotated):
                issue('page', f'section {index+1}', f'Page {actual} differs from {expected} cm')
        for side, bounds in spec['margins_cm'].items():
            actual_margin = getattr(section, side + '_margin').cm
            if not bounds[0] - .01 <= actual_margin <= bounds[1] + .01:
                issue('margin', f'section {index+1}', f'{side}: {actual_margin:.3f} cm outside {bounds}')

    all_text = []
    unresolved = set()
    for story, part, element in story_parts(doc):
        for index, node in enumerate(element.iter(qn('w:p'))):
            paragraph = Paragraph(node, part)
            role = {}
            for style in style_chain(paragraph.style):
                if style.name in rules['style_rules']:
                    role = rules['style_rules'][style.name]
                    break
            location = f'{story}/paragraph {index+1}'
            text = ''.join(t.text or '' for t in node.iter(qn('w:t')))
            all_text.append(text)
            if PLACEHOLDER.search(text):
                issue('placeholder', location, 'Unresolved placeholder', 'warning' if draft else 'error')
            try:
                num_format = numbering_format(paragraph, doc)
                if num_format == 'unresolved':
                    unresolved.add('numbering')
            except (ValueError, AttributeError, KeyError) as exc:
                issue('numbering', location, str(exc))
                num_format = None
            if rules['forbid_bullets']:
                bullet = '•' in text or num_format == 'bullet'
                if bullet:
                    issue('bullet', location, 'Bullet forbidden by selected config')
            for run_node in node.iter(qn('w:r')):
                run = Run(run_node, paragraph)
                if not any(t.text for t in run_node.iter(qn('w:t'))):
                    continue
                for name in ('font', 'color'):
                    expected_value = role.get(name, rules[name])
                    if expected_value is None:
                        continue
                    actual_value, known = effective_property(run, paragraph, name, doc)
                    if not known:
                        unresolved.add(name)
                    elif actual_value != expected_value:
                        issue(name, location, f'{actual_value} differs from {expected_value}')
                if 'size_pt' in role:
                    size, known = effective_property(run, paragraph, 'size', doc)
                    if not known:
                        unresolved.add('size')
                    elif not role['size_pt'][0] <= size <= role['size_pt'][1]:
                        issue('size', location, f'{size} pt outside {role["size_pt"]}')
                for name in ('bold', 'italic'):
                    if name in role:
                        value, known = effective_property(run, paragraph, name, doc)
                        if known and value != role[name]:
                            issue(name, location, f'{name} differs from named style requirement')
        if rules['forbid_shading']:
            for shading in element.iter(qn('w:shd')):
                if shading.get(qn('w:fill'), 'auto') not in ('auto', 'FFFFFF'):
                    issue('shading', story, 'Shading forbidden by selected config')

    joined = '\n'.join(all_text)
    for required in rules['required_text']:
        if required not in joined:
            issue('required', 'document', f'Missing literal: {required}')
    for name in sorted(unresolved):
        report['unchecked'].append(f'Could not resolve some effective {name} values')
    if profile == 'custom' and config is None:
        report['unchecked'].append('Custom formatting: no rule config supplied')
    report['status'] = ('fail' if any(x['severity'] == 'error' for x in report['issues'])
                        else 'partial' if unresolved or (profile == 'custom' and config is None)
                        else 'pass')
    return report


def validate_doc(filepath, profile='administrative', config=None, draft=False):
    """Compatibility boolean: no errors; consult inspect_doc for partial checks."""
    report = inspect_doc(filepath, profile, config, draft)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return report['status'] != 'fail'


def main(argv=None):
    ensure_utf8_stdio()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('filepath')
    parser.add_argument('--profile', default='administrative')
    parser.add_argument('--config', help='Explicit JSON rules for chosen template')
    parser.add_argument('--draft', action='store_true')
    parser.add_argument('--report', help='Write structured JSON report')
    args = parser.parse_args(argv)
    try:
        report = inspect_doc(args.filepath, args.profile, args.config, args.draft)
    except (ValueError, OSError) as exc:
        parser.error(str(exc))
    output = json.dumps(report, ensure_ascii=False, indent=2)
    print(output)
    if args.report:
        Path(args.report).write_text(output + '\n', encoding='utf-8')
    # 2 = incomplete formatting checks; 0 never means legal/visual certification.
    return {'pass': 0, 'fail': 1, 'partial': 2}[report['status']]


if __name__ == '__main__':
    raise SystemExit(main())
