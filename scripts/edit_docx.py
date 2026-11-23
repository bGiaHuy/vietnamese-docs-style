"""Patch verified text slots in a retained DOCX without reconstructing its package."""
import argparse
import hashlib
import json
from pathlib import Path

try:
    from .docx_package import ensure_utf8_stdio, read_package, stories, paragraphs, text_of, replace_span, xml_bytes, write_package
except ImportError:
    from docx_package import ensure_utf8_stdio, read_package, stories, paragraphs, text_of, replace_span, xml_bytes, write_package


def edit_docx(source, destination, edits, expected_sha256=None):
    if expected_sha256 and hashlib.sha256(Path(source).read_bytes()).hexdigest() != expected_sha256.lower():
        raise ValueError('Source checksum differs from the inspected template')
    if not isinstance(edits, list) or not edits:
        raise ValueError('Provide at least one verified slot edit')
    entries, original = read_package(source)
    roots = dict(stories(original))
    changed = {}
    seen = set()
    for edit in edits:
        if not isinstance(edit, dict) or set(edit) != {'part', 'paragraph', 'before', 'after'}:
            raise ValueError('Each edit needs exactly part, paragraph, before and after')
        part, index = edit['part'], edit['paragraph']
        if part not in roots or not isinstance(index, int) or isinstance(index, bool) or index < 0:
            raise ValueError('Unsupported part or paragraph index')
        if (part, index) in seen:
            raise ValueError('Combine edits targeting the same paragraph into one slot')
        seen.add((part, index))
        nodes = paragraphs(roots[part])
        if index >= len(nodes):
            raise ValueError('Paragraph index out of range')
        node = nodes[index]
        before, after = edit['before'], edit['after']
        content = text_of(node)
        if not isinstance(before, str) or (before == '' and content != '') or content.count(before) != 1:
            raise ValueError('Expected text must occur exactly once in the selected paragraph')
        start = content.index(before)
        replace_span(node, start, start+len(before), after)
        changed[part] = xml_bytes(roots[part])
    report = write_package(source, destination, entries, original, changed)
    report['edits'] = len(edits)
    report['unchecked'] = ['Visual layout and field refresh', 'Edits outside supported plain text slots']
    return report


def main(argv=None):
    ensure_utf8_stdio()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source')
    parser.add_argument('destination')
    parser.add_argument('--edits', required=True, help='JSON list of verified text slots')
    parser.add_argument('--source-sha256')
    args = parser.parse_args(argv)
    try:
        edits = json.loads(Path(args.edits).read_text(encoding='utf-8-sig'))
        report = edit_docx(args.source, args.destination, edits, args.source_sha256)
    except (ValueError, OSError) as exc:
        parser.error(str(exc))
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
