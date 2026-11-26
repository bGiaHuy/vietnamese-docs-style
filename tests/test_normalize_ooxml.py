"""Tests for scripts/ooxml_order.py normalize_docx_ooxml and CLI subcommand."""

from __future__ import annotations

import json
import subprocess
import sys
import zipfile
from pathlib import Path

from scripts.ooxml_order import check_element_order, normalize_docx_ooxml
from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _make_docx_with_order_violations(path: Path) -> dict[str, bytes]:
    content_types = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        b'<Default Extension="xml" ContentType="application/xml"/>'
        b"</Types>"
    )
    styles_clean = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        b'<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
        b"</w:styles>"
    )
    footer_clean = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        b"<w:p><w:r><w:t>Trang 1</w:t></w:r></w:p>"
        b"</w:ftr>"
    )
    # Deliberate child_order violation: <w:ind> placed AFTER <w:rPr> inside <w:pPr>.
    bad_doc = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        b'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14">'
        b"<w:body><w:p><w:pPr><w:rPr/><w:ind w:left=\"312\" w:hanging=\"312\"/></w:pPr>"
        b"<w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>"
    )
    # Deliberate child_order violation: <w:sz>/<w:szCs> placed AFTER <w:lang> inside <w:rPr>.
    bad_footnotes = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        b'<w:footnote w:id="1"><w:p><w:r><w:rPr><w:lang w:val="vi-VN"/>'
        b'<w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t>Note</w:t></w:r></w:p></w:footnote>'
        b"</w:footnotes>"
    )
    parts = {
        "[Content_Types].xml": content_types,
        "word/styles.xml": styles_clean,
        "word/footer1.xml": footer_clean,
        "word/document.xml": bad_doc,
        "word/footnotes.xml": bad_footnotes,
    }
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for k, v in parts.items():
            zf.writestr(k, v)
    return parts


def test_normalize_docx_ooxml_fixes_dirty_and_preserves_clean_parts(tmp_path: Path):
    src = tmp_path / "dirty.docx"
    dst = tmp_path / "clean.docx"
    orig_parts = _make_docx_with_order_violations(src)

    res = normalize_docx_ooxml(src, dst)
    assert res["violations_before"] >= 2
    assert res["violations_after"] == 0
    assert set(res["parts_normalized"]) == {"word/document.xml", "word/footnotes.xml"}

    with zipfile.ZipFile(dst, "r") as zf:
        assert zf.read("word/styles.xml") == orig_parts["word/styles.xml"]
        assert zf.read("word/footer1.xml") == orig_parts["word/footer1.xml"]
        assert zf.read("[Content_Types].xml") == orig_parts["[Content_Types].xml"]
        doc_root = etree.fromstring(zf.read("word/document.xml"))
        fn_root = etree.fromstring(zf.read("word/footnotes.xml"))
        assert check_element_order(doc_root) == []
        assert check_element_order(fn_root) == []
        assert doc_root.get("{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable") == "w14"


def test_cli_normalize_ooxml_roundtrip(tmp_path: Path):
    src = tmp_path / "cli_dirty.docx"
    dst = tmp_path / "cli_clean.docx"
    _make_docx_with_order_violations(src)

    proc = subprocess.run(
        [sys.executable, "scripts/cli.py", "normalize-ooxml", str(src), str(dst)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        check=False,
    )
    assert proc.returncode == 0, proc.stderr
    payload = json.loads(proc.stdout)
    assert payload["violations_after"] == 0
    assert dst.is_file()
