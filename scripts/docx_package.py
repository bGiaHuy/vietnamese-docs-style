"""Minimal OOXML text edits; untouched ZIP members stay byte-for-byte identical."""
import hashlib
import os
import sys
import tempfile
from pathlib import Path
from zipfile import ZipFile, BadZipFile

from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}


def ensure_utf8_stdio():
    """Reconfigure stdout/stderr to UTF-8 to avoid Windows cp1252 JSON print errors."""
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, 'reconfigure', None)
        if callable(reconfigure):
            try:
                reconfigure(encoding='utf-8', errors='replace')
            except (AttributeError, ValueError, OSError):
                pass


def read_package(source):
    try:
        with ZipFile(source) as archive:
            entries = archive.infolist()
            if len({x.filename for x in entries}) != len(entries):
                raise ValueError('Duplicate ZIP members are unsupported')
            if any(x.filename.startswith('_xmlsignatures/') for x in entries):
                raise ValueError('Editing would invalidate digital signatures; use an unsigned template')
            data = {entry.filename: archive.read(entry) for entry in entries}
    except BadZipFile as exc:
        raise ValueError('Unreadable or corrupt DOCX archive') from exc
    if 'word/document.xml' not in data:
        raise ValueError('Not a supported Word document package')
    return entries, data


def stories(data):
    for name, raw in data.items():
        if not name.startswith('word/') or not name.endswith('.xml'):
            continue
        parser = etree.XMLParser(resolve_entities=False, no_network=True)
        try:
            root = etree.fromstring(raw, parser)
        except etree.XMLSyntaxError as exc:
            raise ValueError(f'Invalid XML in {name}') from exc
        if root.tag in {f'{{{W}}}{tag}' for tag in ('document', 'hdr', 'ftr')}:
            yield name, root


def paragraphs(root):
    return [p for p in root.iter(f'{{{W}}}p')
            if not any(a.tag in {f'{{{W}}}{tag}' for tag in ('txbxContent', 'del', 'ins', 'sdt')}
                       for a in p.iterancestors())]


def text_of(paragraph):
    return ''.join(value for node, value in inline_items(paragraph))


def inline_items(paragraph):
    for run in paragraph.findall(f'{{{W}}}r'):
        for node in run:
            if node.tag == f'{{{W}}}t':
                yield node, node.text or ''
            elif node.tag == f'{{{W}}}tab':
                yield node, '\t'
            elif node.tag in (f'{{{W}}}br', f'{{{W}}}cr'):
                yield node, '\n'


def replace_span(paragraph, start, end, replacement):
    if not isinstance(replacement, str) or any(c in replacement for c in '\n\r\t'):
        raise ValueError('Inline edits require a single-line string without tabs')
    content = text_of(paragraph)
    empty_slot = start == end == 0 and not content
    if start < 0 or (end <= start and not empty_slot) or end > len(content):
        raise ValueError('Invalid replacement span')
    unsafe = ('fldChar', 'fldSimple', 'hyperlink', 'drawing', 'pict', 'sdt', 'ins', 'del')
    if any(list(paragraph.iter(f'{{{W}}}{tag}')) for tag in unsafe):
        raise ValueError('Editable paragraph contains fields, links, controls, drawings, revisions or breaks')
    if empty_slot:
        texts = paragraph.xpath('./w:r/w:t', namespaces=NS)
        if texts:
            node = texts[0]
        else:
            runs = paragraph.findall(f'{{{W}}}r')
            run = runs[0] if runs else etree.SubElement(paragraph, f'{{{W}}}r')
            node = etree.SubElement(run, f'{{{W}}}t')
        node.text = replacement
        node.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        return
    if any(c in text_of(paragraph)[start:end] for c in '\t\n'):
        raise ValueError('Text replacement cannot cross tabs or line breaks')
    offset = 0
    inserted = False
    for node, text in inline_items(paragraph):
        a, b = offset, offset + len(text)
        offset = b
        if b <= start or a >= end:
            continue
        left, right = max(0, start-a), min(len(text), end-a)
        node.text = text[:left] + (replacement if not inserted else '') + text[right:]
        node.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        inserted = True


MC_IGNORABLE = '{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable'


def preserve_root_namespaces(before_bytes, after_root):
    """Ensure every root xmlns declaration and mc:Ignorable from before_bytes survives serialization."""
    parser = etree.XMLParser(resolve_entities=False, no_network=True)
    before_root = etree.fromstring(before_bytes, parser)
    merged_ns = dict(before_root.nsmap or {})
    merged_ns.update(after_root.nsmap or {})

    attrs = dict(after_root.attrib)
    if MC_IGNORABLE in before_root.attrib and MC_IGNORABLE not in attrs:
        attrs[MC_IGNORABLE] = before_root.attrib[MC_IGNORABLE]

    if dict(after_root.nsmap or {}) == merged_ns and dict(after_root.attrib) == attrs:
        return after_root

    rebuilt = etree.Element(after_root.tag, attrib=attrs, nsmap=merged_ns)
    rebuilt.text = after_root.text
    for child in list(after_root):
        rebuilt.append(child)
    return rebuilt


def audit_namespace_preservation(before_bytes, after_bytes):
    """Verify after_bytes keeps all root namespace bindings and mc:Ignorable from before_bytes."""
    parser = etree.XMLParser(resolve_entities=False, no_network=True)
    b_root = etree.fromstring(before_bytes, parser)
    a_root = etree.fromstring(after_bytes, parser)
    b_ns = b_root.nsmap or {}
    a_ns = a_root.nsmap or {}
    missing_ns = [
        {'prefix': prefix or '', 'uri': uri}
        for prefix, uri in b_ns.items()
        if a_ns.get(prefix) != uri
    ]
    ign_ok = (
        True
        if MC_IGNORABLE not in b_root.attrib
        else (a_root.attrib.get(MC_IGNORABLE) == b_root.attrib[MC_IGNORABLE])
    )
    return {
        'preserved': (not missing_ns) and ign_ok,
        'missing_ns': missing_ns,
        'ignorable_preserved': ign_ok,
    }


def xml_bytes(root, original_bytes=None):
    target = preserve_root_namespaces(original_bytes, root) if original_bytes is not None else root
    return etree.tostring(target, encoding='UTF-8', xml_declaration=True, standalone=True)


def write_package(source, destination, entries, original, changed):
    source, destination = Path(source).resolve(), Path(destination).resolve()
    if source == destination:
        raise ValueError('Write to a new file; preserve the original template')
    destination.parent.mkdir(parents=True, exist_ok=True)
    clean_entries = [entry for entry in entries if not entry.is_dir()]
    # No extraction of archive paths. Atomic output avoids half-written DOCX files.
    handle, temporary = tempfile.mkstemp(suffix='.docx', dir=destination.parent)
    os.close(handle)
    try:
        with ZipFile(temporary, 'w') as archive:
            for entry in clean_entries:
                archive.writestr(entry, changed.get(entry.filename, original[entry.filename]))
        os.replace(temporary, destination)
    finally:
        if Path(temporary).exists():
            Path(temporary).unlink()
    return {
        'output': str(destination), 'source_sha256': hashlib.sha256(source.read_bytes()).hexdigest(),
        'changed_parts': sorted(changed), 'preserved_parts': len(original) - len(changed),
    }
