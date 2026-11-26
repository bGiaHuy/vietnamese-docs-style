"""Independent non-target preservation, namespace audit, and XML telemetry."""

from __future__ import annotations

import hashlib
from typing import Any

from lxml import etree

from eval_harness.ooxml_reader import MC, W, parse_xml


def audit_namespaces(before_doc_xml: bytes, after_doc_xml: bytes) -> dict[str, Any]:
    """Verify root namespace declarations and mc:Ignorable tokens were preserved."""
    b_root = parse_xml(before_doc_xml)
    a_root = parse_xml(after_doc_xml)

    b_ns = dict(b_root.nsmap)
    a_ns = dict(a_root.nsmap)
    dropped_ns = [
        {"prefix": k, "uri": v}
        for k, v in b_ns.items()
        if a_ns.get(k) != v
    ]

    b_ign = set((b_root.get(f"{{{MC}}}Ignorable") or "").split())
    a_ign = set((a_root.get(f"{{{MC}}}Ignorable") or "").split())
    dropped_ignorable = sorted(b_ign - a_ign)

    return {
        "preserved": len(dropped_ns) == 0 and len(dropped_ignorable) == 0,
        "dropped_namespaces": dropped_ns,
        "dropped_ignorable": dropped_ignorable,
    }


def _c14n_hash(node: etree._Element) -> str:
    payload = etree.tostring(node, method="c14n")
    return hashlib.sha256(payload).hexdigest()


def compute_xml_ntpr(
    before_root: etree._Element,
    after_root: etree._Element,
    allowed_body_indices: set[int],
) -> float:
    """Compute canonical XML Non-Target Preservation Ratio across non-target w:body children."""
    b_body = before_root.find(f"{{{W}}}body")
    a_body = after_root.find(f"{{{W}}}body")
    if b_body is None or a_body is None:
        return 0.0

    b_children = list(b_body)
    a_children = list(a_body)
    non_target = [idx for idx in range(len(b_children)) if idx not in allowed_body_indices]
    if not non_target:
        return 1.0 if len(b_children) == len(a_children) else 0.0
    if len(b_children) != len(a_children):
        matches = 0
        for idx in non_target:
            if idx < len(a_children) and _c14n_hash(b_children[idx]) == _c14n_hash(a_children[idx]):
                matches += 1
        return round(matches / len(non_target), 4)

    matches = sum(
        1 for idx in non_target
        if _c14n_hash(b_children[idx]) == _c14n_hash(a_children[idx])
    )
    return round(matches / len(non_target), 4)


def _norm_fmt_sig(fmt: dict[str, Any]) -> tuple[Any, ...]:
    return (
        fmt.get("align"),
        round(fmt["first_line_cm"], 2) if fmt.get("first_line_cm") is not None else None,
        round(fmt["hanging_cm"], 2) if fmt.get("hanging_cm") is not None else None,
        fmt.get("line_rule"),
        round(fmt["line_pt"], 2) if fmt.get("line_pt") is not None else None,
        round(fmt["line_multiple"], 2) if fmt.get("line_multiple") is not None else None,
        round(fmt["space_before_pt"], 2) if fmt.get("space_before_pt") is not None else None,
        round(fmt["space_after_pt"], 2) if fmt.get("space_after_pt") is not None else None,
        bool(fmt.get("keep_with_next")),
    )


def check_non_target_preservation(
    before_info: dict[str, Any],
    after_info: dict[str, Any],
    before_parts: dict[str, bytes],
    after_parts: dict[str, bytes],
    *,
    target_kind: str = "table",
    target_table_ordinals: set[int] | None = None,
    target_image_ordinals: set[int] | None = None,
    target_para_ordinals: set[int] | None = None,
    target_section_indices: set[int] | None = None,
    allowed_body_indices: set[int] | None = None,
    require_layout_tables_unchanged: bool = True,
    allow_header_footer_parts: set[str] | None = None,
    allow_footnote_changes: bool = False,
) -> dict[str, Any]:
    """Compare semantic and structural non-target preservation independently of scripts/."""
    t_tbls = target_table_ordinals or set()
    t_imgs = target_image_ordinals or set()
    t_paras = target_para_ordinals or set()
    t_secs = target_section_indices or set()
    allowed_bidx = allowed_body_indices or set()

    # 1. Body paragraph text invariance (NFC).
    b_paras = before_info.get("paragraphs", [])
    a_paras = after_info.get("paragraphs", [])
    para_texts_ok = [p["text_nfc"] for p in b_paras] == [p["text_nfc"] for p in a_paras]

    # 2. Non-target table invariance + target table text invariance.
    b_tbls = before_info.get("tables", [])
    a_tbls = after_info.get("tables", [])
    tables_ok = len(b_tbls) == len(a_tbls)
    if tables_ok:
        for bt, at in zip(b_tbls, a_tbls):
            if bt["row_texts_nfc"] != at["row_texts_nfc"]:
                tables_ok = False
                break
            if bt["ordinal"] in t_tbls:
                continue
            if target_kind == "paginate" and not bt.get("is_layout_table"):
                continue
            if bt.get("is_layout_table") and not require_layout_tables_unchanged and target_kind == "paginate":
                continue
            if (
                bt["rows"] != at["rows"]
                or bt["grid_cols"] != at["grid_cols"]
                or bt["alignment"] != at["alignment"]
                or abs(bt["total_width_cm"] - at["total_width_cm"]) > 0.1
                or len(bt["grid_widths_cm"]) != len(at["grid_widths_cm"])
                or any(abs(x - y) > 0.1 for x, y in zip(bt["grid_widths_cm"], at["grid_widths_cm"]))
            ):
                tables_ok = False
                break

    # 3. Non-target image invariance.
    b_imgs = before_info.get("images", [])
    a_imgs = after_info.get("images", [])
    images_ok = len(b_imgs) == len(a_imgs)
    if images_ok:
        for bi, ai in zip(b_imgs, a_imgs):
            if bi["ordinal"] in t_imgs or target_kind == "paginate":
                continue
            if (
                abs(bi["width_cm"] - ai["width_cm"]) > 0.05
                or abs(bi["height_cm"] - ai["height_cm"]) > 0.05
                or bi["alignment"] != ai["alignment"]
            ):
                images_ok = False
                break

    # 4. Non-target paragraph formatting invariance (when editing paragraphs or runs).
    non_target_ppr_ok = True
    if target_kind in ("paragraph", "merge_runs", "multi") or (target_kind == "out_of_scope" and t_paras):
        excluded_bidx = {p["body_index"] for p in b_paras if p["ordinal"] in t_paras}
        excluded_bidx |= {img["body_index"] for img in b_imgs if img["ordinal"] in t_imgs}
        b_sigs = [_norm_fmt_sig(p["effective_format"]) for p in b_paras if p["body_index"] not in excluded_bidx]
        a_sigs = [_norm_fmt_sig(p["effective_format"]) for p in a_paras if p["body_index"] not in excluded_bidx]
        non_target_ppr_ok = b_sigs == a_sigs

    # 5. Non-target sections invariance.
    b_secs = before_info.get("sections", [])
    a_secs = after_info.get("sections", [])
    sections_ok = len(b_secs) == len(a_secs)
    if sections_ok and target_kind != "paginate":
        for bs, as_ in zip(b_secs, a_secs):
            if target_kind == "section" and bs["section_index"] in t_secs:
                continue
            if (
                abs(bs["page_width_cm"] - as_["page_width_cm"]) > 0.1
                or abs(bs["page_height_cm"] - as_["page_height_cm"]) > 0.1
                or bs["orientation"] != as_["orientation"]
                or any(abs(bs["margins_cm"][k] - as_["margins_cm"][k]) > 0.05 for k in ("top", "bottom", "left", "right"))
            ):
                sections_ok = False
                break

    # 6. Footnotes/endnotes & header/footer preservation.
    if allow_footnote_changes:
        notes_ok = before_info.get("notes_text_nfc", {}).get("word/endnotes.xml") == after_info.get("notes_text_nfc", {}).get("word/endnotes.xml")
    else:
        notes_ok = before_info.get("notes_text_nfc", {}) == after_info.get("notes_text_nfc", {})

    b_hf = before_info.get("headers_footers", {})
    a_hf = after_info.get("headers_footers", {})
    allowed_hf = allow_header_footer_parts or set()
    if "*" in allowed_hf:
        hf_text_ok = True
    elif allowed_hf:
        hf_text_ok = (
            {k: v["text_nfc"] for k, v in b_hf.items() if k not in allowed_hf}
            == {k: v["text_nfc"] for k, v in a_hf.items() if k not in allowed_hf}
        )
    else:
        hf_text_ok = {k: v["text_nfc"] for k, v in b_hf.items()} == {k: v["text_nfc"] for k, v in a_hf.items()}

    hf_keys = [k for k in before_parts if k.startswith(("word/header", "word/footer")) and k.endswith(".xml")]
    hf_byte_identical = bool(hf_keys and all(before_parts.get(k) == after_parts.get(k) for k in hf_keys))

    other_parts = [k for k in before_parts if k != "word/document.xml"]
    same_other = sum(1 for k in other_parts if before_parts.get(k) == after_parts.get(k))
    zip_byte_ident = round(same_other / len(other_parts), 4) if other_parts else 1.0

    # 7. XML NTPR telemetry.
    before_root = parse_xml(before_parts["word/document.xml"])
    after_root = parse_xml(after_parts["word/document.xml"])
    xml_ntpr = compute_xml_ntpr(before_root, after_root, allowed_bidx)

    semantic_ok = bool(
        para_texts_ok
        and tables_ok
        and images_ok
        and non_target_ppr_ok
        and sections_ok
        and notes_ok
        and hf_text_ok
    )

    return {
        "semantic_non_target_preserved": semantic_ok,
        "details": {
            "para_texts_ok": para_texts_ok,
            "tables_ok": tables_ok,
            "images_ok": images_ok,
            "non_target_ppr_ok": non_target_ppr_ok,
            "sections_ok": sections_ok,
            "notes_ok": notes_ok,
            "hf_text_ok": hf_text_ok,
        },
        "header_footer_byte_identical": hf_byte_identical,
        "zip_parts_byte_identical": zip_byte_ident,
        "xml_ntpr": xml_ntpr,
    }
