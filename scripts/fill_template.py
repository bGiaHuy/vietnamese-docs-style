"""Fill explicit {{tokens}} across runs while preserving the surrounding DOCX."""
import argparse
import json
import re
from pathlib import Path
from lxml import etree

try:
    from .docx_package import ensure_utf8_stdio, read_package, stories, paragraphs, text_of, replace_span, xml_bytes, write_package
except ImportError:
    from docx_package import ensure_utf8_stdio, read_package, stories, paragraphs, text_of, replace_span, xml_bytes, write_package

TOKEN = re.compile(r'\{\{([^{}]+)\}\}')


def fill_template(source, destination, values, require_all=True):
    source, destination = Path(source).resolve(), Path(destination).resolve()
    if source == destination:
        raise ValueError('Write to a new file; keep the original template')
    if not isinstance(values, dict) or any(not isinstance(k, str) or not isinstance(v, str)
                                           for k, v in values.items()):
        raise ValueError('Values must be a JSON object of string keys and string values')
    if any('\n' in v or '\r' in v or '\t' in v for v in values.values()):
        raise ValueError('Inline replacement values cannot contain line breaks or tabs')
    entries, original = read_package(source)
    counts = {key: 0 for key in values}
    unresolved = set()
    changed = {}
    for story, root in stories(original):
        edited = False
        for paragraph in paragraphs(root):
            content = text_of(paragraph)
            for match in reversed(list(TOKEN.finditer(content))):
                key = match.group(1)
                if key not in values:
                    unresolved.add(key)
                    continue
                replace_span(paragraph, match.start(), match.end(), values[key])
                counts[key] += 1
                edited = True
        if edited:
            changed[story] = xml_bytes(root)
    unused = [key for key, count in counts.items() if count == 0]
    # Include unsupported visible stories in the unresolved report. Otherwise
    # tokens inside hyperlinks/footnotes could be silently left behind.
    for name, raw in original.items():
        if name.startswith('word/') and name.endswith('.xml'):
            root = etree.fromstring(changed.get(name, raw), etree.XMLParser(resolve_entities=False, no_network=True))
            text = ''.join(node.text or '' for node in root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'))
            unresolved.update(match.group(1) for match in TOKEN.finditer(text))
    if require_all and (unresolved or unused):
        raise ValueError(f'Unfilled tokens: {sorted(unresolved)}; unused values: {unused}')
    report = write_package(source, destination, entries, original, changed)
    report.update(replacements=counts, unfilled=sorted(unresolved), unused=unused,
                  unchecked=['Tokens outside supported plain body/table/header/footer paragraphs',
                             'Visual layout and cached fields require verification'])
    return report



def main(argv=None):
    ensure_utf8_stdio()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source')
    parser.add_argument('destination')
    parser.add_argument('--values', required=True, help='JSON file, plain token names to strings')
    parser.add_argument('--draft', action='store_true', help='Allow unfilled tokens/unused values')
    args = parser.parse_args(argv)
    try:
        values = json.loads(Path(args.values).read_text(encoding='utf-8-sig'))
        report = fill_template(args.source, args.destination, values, not args.draft)
    except (ValueError, OSError) as exc:
        parser.error(str(exc))
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
