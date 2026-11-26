"""Standalone OOXML package and document reader using only stdlib + lxml."""

from __future__ import annotations

import unicodedata
import zipfile
from pathlib import Path
from typing import Any

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

DXA_PER_CM = 1440.0 / 2.54
EMU_PER_CM = 360000.0
PT_PER_CM = 72.0 / 2.54

SAFE_XML_PARSER = etree.XMLParser(resolve_entities=False, no_network=True)

_LAYOUT_KEYWORDS = (
    "cộng hòa xã hội chủ nghĩa việt nam",
    "độc lập - tự do - hạnh phúc",
    "đảng cộng sản việt nam",
    "nơi nhận",
    "lưu:",
    "như trên",
    "kt.",
    "tm.",
    "t/m",
    "chủ tịch",
    "giám đốc",
    "cục trưởng",
    "hiệu trưởng",
    "chánh văn phòng",
    "ubnd",
    "ủy ban nhân dân",
    "ban chấp hành",
)


def nfc(text: str) -> str:
    """Normalize text to Unicode NFC."""
    return unicodedata.normalize("NFC", text or "")


def parse_xml(xml_bytes: bytes) -> etree._Element:
    """Parse XML bytes using hardened parser settings."""
    return etree.fromstring(xml_bytes, SAFE_XML_PARSER)


def _w_attr(node: etree._Element | None, attr: str) -> str | None:
    if node is None:
        return None
    return node.get(f"{{{W}}}{attr}")


def _int_w_attr(node: etree._Element | None, attr: str, default: int | None = None) -> int | None:
    raw = _w_attr(node, attr)
    if raw is None:
        return default
    try:
        return int(raw)
    except ValueError:
        try:
            return round(float(raw))
        except ValueError:
            return default


def _toggle_on(el: etree._Element | None) -> bool | None:
    if el is None:
        return None
    val = (_w_attr(el, "val") or "1").strip().lower()
    return val not in ("0", "false", "off")


def _norm_align(val: str | None) -> str | None:
    if val is None:
        return None
    low = val.strip().lower()
    if low == "start":
        return "left"
    if low == "end":
        return "right"
    if low == "justify":
        return "both"
    return low


def read_zip_parts(docx_path: str | Path) -> tuple[dict[str, Any], dict[str, bytes]]:
    """Read ZIP container metadata and raw non-directory member payloads."""
    path = Path(docx_path)
    names: list[str] = []
    has_dir_entries = False
    duplicate_names: list[str] = []
    parts: dict[str, bytes] = {}
    seen: set[str] = set()

    with zipfile.ZipFile(path, "r") as zf:
        for info in zf.infolist():
            if info.is_dir():
                has_dir_entries = True
                continue
            names.append(info.filename)
            if info.filename in seen:
                duplicate_names.append(info.filename)
            seen.add(info.filename)
            parts[info.filename] = zf.read(info.filename)

    is_signed = any(name.startswith("_xmlsignatures/") for name in names)
    meta = {
        "names": names,
        "has_dir_entries": has_dir_entries,
        "duplicate_names": duplicate_names,
        "is_signed": is_signed,
    }
    return meta, parts


def text_of_element(node: etree._Element | None) -> str:
    """Extract concatenated text from w:t, w:tab, and w:br/w:cr inside an element."""
    if node is None:
        return ""
    buf: list[str] = []
    for el in node.iter():
        if not isinstance(el.tag, str):
            continue
        q = etree.QName(el)
        if q.namespace != W:
            continue
        if q.localname == "t":
            if el.text:
                buf.append(el.text)
        elif q.localname == "tab":
            buf.append("\t")
        elif q.localname in ("br", "cr"):
            buf.append("\n")
    return "".join(buf)


def parse_styles(styles_xml_bytes: bytes | None) -> dict[str, Any]:
    """Parse docDefaults and style definitions including basedOn chains."""
    ctx: dict[str, Any] = {
        "doc_defaults_ppr": None,
        "doc_defaults_rpr": None,
        "styles": {},
        "name_to_id": {},
        "default_paragraph_style": None,
        "default_table_style": None,
    }
    if not styles_xml_bytes:
        return ctx

    root = parse_xml(styles_xml_bytes)
    doc_defs = root.find(f"{{{W}}}docDefaults")
    if doc_defs is not None:
        p_def = doc_defs.find(f"{{{W}}}pPrDefault/{{{W}}}pPr")
        r_def = doc_defs.find(f"{{{W}}}rPrDefault/{{{W}}}rPr")
        ctx["doc_defaults_ppr"] = p_def
        ctx["doc_defaults_rpr"] = r_def

    for st in root.findall(f"{{{W}}}style"):
        sid = _w_attr(st, "styleId")
        if not sid:
            continue
        stype = _w_attr(st, "type") or ""
        name_el = st.find(f"{{{W}}}name")
        sname = _w_attr(name_el, "val") or sid
        based_el = st.find(f"{{{W}}}basedOn")
        based_on = _w_attr(based_el, "val")
        is_default = (_w_attr(st, "default") or "").lower() in ("1", "true", "on")

        tbl_style_prs: dict[str, etree._Element] = {}
        for tsp in st.findall(f"{{{W}}}tblStylePr"):
            t_role = _w_attr(tsp, "type")
            if t_role:
                tbl_style_prs[t_role] = tsp

        rec = {
            "style_id": sid,
            "type": stype,
            "name": sname,
            "based_on": based_on,
            "is_default": is_default,
            "node": st,
            "pPr": st.find(f"{{{W}}}pPr"),
            "rPr": st.find(f"{{{W}}}rPr"),
            "tblPr": st.find(f"{{{W}}}tblPr"),
            "tcPr": st.find(f"{{{W}}}tcPr"),
            "tblStylePr": tbl_style_prs,
        }
        ctx["styles"][sid] = rec
        ctx["name_to_id"][sname.lower()] = sid
        if is_default and stype == "paragraph" and ctx["default_paragraph_style"] is None:
            ctx["default_paragraph_style"] = sid
        if is_default and stype == "table" and ctx["default_table_style"] is None:
            ctx["default_table_style"] = sid

    return ctx


def resolve_style_chain(style_id: str | None, styles_ctx: dict[str, Any]) -> list[dict[str, Any]]:
    """Resolve style inheritance chain from root ancestor to leaf style."""
    if not style_id:
        return []
    styles = styles_ctx.get("styles", {})
    chain: list[dict[str, Any]] = []
    visited: set[str] = set()
    cur: str | None = style_id
    while cur and cur not in visited and cur in styles:
        visited.add(cur)
        rec = styles[cur]
        chain.append(rec)
        cur = rec.get("based_on")
    chain.reverse()
    return chain


def find_style_record(name_or_id: str | None, styles_ctx: dict[str, Any]) -> dict[str, Any] | None:
    """Locate a style record by styleId or display name."""
    if not name_or_id:
        return None
    styles = styles_ctx.get("styles", {})
    if name_or_id in styles:
        return styles[name_or_id]
    mapped_id = styles_ctx.get("name_to_id", {}).get(name_or_id.lower())
    return styles.get(mapped_id) if mapped_id else None


def _apply_ppr_layer(fmt: dict[str, Any], p_pr: etree._Element | None) -> None:
    if p_pr is None:
        return
    jc = p_pr.find(f"{{{W}}}jc")
    if jc is not None:
        fmt["align"] = _norm_align(_w_attr(jc, "val"))

    ind = p_pr.find(f"{{{W}}}ind")
    if ind is not None:
        fl = _int_w_attr(ind, "firstLine")
        hg = _int_w_attr(ind, "hanging")
        if hg is not None:
            fmt["hanging_cm"] = round(hg / DXA_PER_CM, 4)
            fmt["first_line_cm"] = None
        elif fl is not None:
            fmt["first_line_cm"] = round(fl / DXA_PER_CM, 4)
            fmt["hanging_cm"] = None

    sp = p_pr.find(f"{{{W}}}spacing")
    if sp is not None:
        before = _int_w_attr(sp, "before")
        after = _int_w_attr(sp, "after")
        line = _int_w_attr(sp, "line")
        rule = _w_attr(sp, "lineRule")
        if before is not None:
            fmt["space_before_pt"] = round(before / 20.0, 3)
        if after is not None:
            fmt["space_after_pt"] = round(after / 20.0, 3)
        if line is not None or rule is not None:
            eff_rule = rule or fmt.get("line_rule") or "auto"
            fmt["line_rule"] = eff_rule
            if line is not None:
                if eff_rule in ("exact", "atLeast"):
                    fmt["line_pt"] = round(line / 20.0, 3)
                    fmt["line_multiple"] = None
                else:
                    fmt["line_multiple"] = round(line / 240.0, 4)
                    fmt["line_pt"] = None

    kn = p_pr.find(f"{{{W}}}keepNext")
    kn_tog = _toggle_on(kn)
    if kn_tog is not None:
        fmt["keep_with_next"] = kn_tog


def _apply_rpr_layer(fmt: dict[str, Any], r_pr: etree._Element | None) -> None:
    if r_pr is None:
        return
    rf = r_pr.find(f"{{{W}}}rFonts")
    if rf is not None:
        ascii_f = _w_attr(rf, "ascii")
        hansi_f = _w_attr(rf, "hAnsi")
        cs_f = _w_attr(rf, "cs")
        if ascii_f:
            fmt["font_ascii"] = ascii_f
        if hansi_f:
            fmt["font_hansi"] = hansi_f
        if cs_f:
            fmt["font_cs"] = cs_f

    sz = r_pr.find(f"{{{W}}}sz")
    sz_hp = _int_w_attr(sz, "val")
    if sz_hp is not None:
        fmt["size_pt"] = round(sz_hp / 2.0, 2)

    b_tog = _toggle_on(r_pr.find(f"{{{W}}}b"))
    if b_tog is not None:
        fmt["bold"] = b_tog

    col = r_pr.find(f"{{{W}}}color")
    if col is not None:
        cval = _w_attr(col, "val")
        if cval is not None:
            fmt["color"] = cval.upper()


def effective_paragraph_format(
    p_node: etree._Element,
    styles_ctx: dict[str, Any],
    *,
    tbl_style_id: str | None = None,
) -> dict[str, Any]:
    """Compute effective paragraph formatting through docDefaults, styles, and direct pPr."""
    fmt: dict[str, Any] = {
        "align": None,
        "first_line_cm": None,
        "hanging_cm": None,
        "line_rule": None,
        "line_pt": None,
        "line_multiple": None,
        "space_before_pt": None,
        "space_after_pt": None,
        "keep_with_next": False,
    }
    _apply_ppr_layer(fmt, styles_ctx.get("doc_defaults_ppr"))

    p_pr = p_node.find(f"{{{W}}}pPr")
    p_style_el = p_pr.find(f"{{{W}}}pStyle") if p_pr is not None else None
    p_style_id = _w_attr(p_style_el, "val") or styles_ctx.get("default_paragraph_style")

    if tbl_style_id:
        for st in resolve_style_chain(tbl_style_id, styles_ctx):
            _apply_ppr_layer(fmt, st.get("pPr"))
    for st in resolve_style_chain(p_style_id, styles_ctx):
        _apply_ppr_layer(fmt, st.get("pPr"))

    _apply_ppr_layer(fmt, p_pr)
    return fmt


def effective_run_format(
    run_node: etree._Element,
    p_node: etree._Element,
    styles_ctx: dict[str, Any],
    *,
    tbl_style_id: str | None = None,
    is_header_row: bool = False,
) -> dict[str, Any]:
    """Compute effective run formatting through docDefaults, style chains, pPr/rPr, and run rPr."""
    fmt: dict[str, Any] = {
        "font_ascii": None,
        "font_hansi": None,
        "font_cs": None,
        "size_pt": None,
        "bold": False,
        "color": None,
    }
    _apply_rpr_layer(fmt, styles_ctx.get("doc_defaults_rpr"))

    p_pr = p_node.find(f"{{{W}}}pPr")
    p_style_el = p_pr.find(f"{{{W}}}pStyle") if p_pr is not None else None
    p_style_id = _w_attr(p_style_el, "val") or styles_ctx.get("default_paragraph_style")
    for st in resolve_style_chain(p_style_id, styles_ctx):
        _apply_rpr_layer(fmt, st.get("rPr"))

    if tbl_style_id:
        for st in resolve_style_chain(tbl_style_id, styles_ctx):
            _apply_rpr_layer(fmt, st.get("rPr"))
            tsp_map = st.get("tblStylePr", {})
            if "wholeTable" in tsp_map:
                _apply_rpr_layer(fmt, tsp_map["wholeTable"].find(f"{{{W}}}rPr"))
            if is_header_row and "firstRow" in tsp_map:
                _apply_rpr_layer(fmt, tsp_map["firstRow"].find(f"{{{W}}}rPr"))

    if p_pr is not None:
        _apply_rpr_layer(fmt, p_pr.find(f"{{{W}}}rPr"))

    r_pr = run_node.find(f"{{{W}}}rPr")
    r_style_el = r_pr.find(f"{{{W}}}rStyle") if r_pr is not None else None
    r_style_id = _w_attr(r_style_el, "val")
    for st in resolve_style_chain(r_style_id, styles_ctx):
        _apply_rpr_layer(fmt, st.get("rPr"))

    _apply_rpr_layer(fmt, r_pr)
    return fmt


def _parse_sect_pr(sect_pr: etree._Element, section_index: int, body_index: int) -> dict[str, Any]:
    pg_sz = sect_pr.find(f"{{{W}}}pgSz")
    pg_mar = sect_pr.find(f"{{{W}}}pgMar")
    w_dxa = _int_w_attr(pg_sz, "w")
    h_dxa = _int_w_attr(pg_sz, "h")
    orient_attr = _w_attr(pg_sz, "orient")
    page_w_cm = round(w_dxa / DXA_PER_CM, 3) if w_dxa else 21.0
    page_h_cm = round(h_dxa / DXA_PER_CM, 3) if h_dxa else 29.7
    if orient_attr in ("portrait", "landscape"):
        orientation = orient_attr
    else:
        orientation = "landscape" if page_w_cm > page_h_cm else "portrait"

    def _mar(side: str, default_dxa: int) -> float:
        val = _int_w_attr(pg_mar, side, default_dxa)
        return round((val if val is not None else default_dxa) / DXA_PER_CM, 3)

    margins_cm = {
        "top": _mar("top", 1440),
        "bottom": _mar("bottom", 1440),
        "left": _mar("left", 1440),
        "right": _mar("right", 1440),
    }
    printable_w = round(max(0.0, page_w_cm - margins_cm["left"] - margins_cm["right"]), 3)
    return {
        "section_index": section_index,
        "body_index": body_index,
        "page_width_cm": page_w_cm,
        "page_height_cm": page_h_cm,
        "orientation": orientation,
        "margins_cm": margins_cm,
        "printable_width_cm": printable_w,
    }


def extract_sections(doc_root: etree._Element) -> list[dict[str, Any]]:
    """Extract ordered WordprocessingML sections from w:body."""
    body = doc_root.find(f"{{{W}}}body")
    if body is None:
        return []
    sections: list[dict[str, Any]] = []
    sec_num = 0
    for b_idx, child in enumerate(body):
        if not isinstance(child.tag, str):
            continue
        ln = etree.QName(child).localname
        if ln == "p":
            sect_pr = child.find(f"{{{W}}}pPr/{{{W}}}sectPr")
            if sect_pr is not None:
                sec_num += 1
                sections.append(_parse_sect_pr(sect_pr, sec_num, b_idx))
        elif ln == "sectPr":
            sec_num += 1
            sections.append(_parse_sect_pr(child, sec_num, b_idx))
    if not sections:
        sections.append(
            {
                "section_index": 1,
                "body_index": 0,
                "page_width_cm": 21.0,
                "page_height_cm": 29.7,
                "orientation": "portrait",
                "margins_cm": {"top": 2.54, "bottom": 2.54, "left": 2.54, "right": 2.54},
                "printable_width_cm": 15.92,
            }
        )
    return sections


def _section_for_body_index(body_index: int, sections: list[dict[str, Any]]) -> dict[str, Any]:
    for sec in sections:
        if body_index <= sec["body_index"]:
            return sec
    return sections[-1]


def _resolve_tbl_pr_chain(tbl_node: etree._Element, styles_ctx: dict[str, Any]) -> tuple[str | None, list[etree._Element]]:
    tbl_pr = tbl_node.find(f"{{{W}}}tblPr")
    st_el = tbl_pr.find(f"{{{W}}}tblStyle") if tbl_pr is not None else None
    tbl_style_id = _w_attr(st_el, "val") or styles_ctx.get("default_table_style")
    layers: list[etree._Element] = []
    if tbl_style_id:
        for st in resolve_style_chain(tbl_style_id, styles_ctx):
            if st.get("tblPr") is not None:
                layers.append(st["tblPr"])
    if tbl_pr is not None:
        layers.append(tbl_pr)
    return tbl_style_id, layers


def _extract_borders(tbl_pr_layers: list[etree._Element]) -> dict[str, dict[str, Any]]:
    borders: dict[str, dict[str, Any]] = {}
    aliases = {
        "top": ("top",),
        "bottom": ("bottom",),
        "left": ("left", "start"),
        "right": ("right", "end"),
        "insideH": ("insideH",),
        "insideV": ("insideV",),
    }
    for layer in tbl_pr_layers:
        tb = layer.find(f"{{{W}}}tblBorders")
        if tb is None:
            continue
        for canon, cands in aliases.items():
            for cand in cands:
                el = tb.find(f"{{{W}}}{cand}")
                if el is not None:
                    borders[canon] = {
                        "val": (_w_attr(el, "val") or "").lower(),
                        "sz": _int_w_attr(el, "sz"),
                        "color": (_w_attr(el, "color") or "").upper() or None,
                    }
                    break
    return borders


def _extract_tbl_cell_margins_cm(tbl_pr_layers: list[etree._Element]) -> dict[str, float] | None:
    margins: dict[str, float] = {}
    aliases = {
        "top": ("top",),
        "bottom": ("bottom",),
        "left": ("left", "start"),
        "right": ("right", "end"),
    }
    found = False
    for layer in tbl_pr_layers:
        mar = layer.find(f"{{{W}}}tblCellMar")
        if mar is None:
            continue
        found = True
        for canon, cands in aliases.items():
            for cand in cands:
                el = mar.find(f"{{{W}}}{cand}")
                if el is not None:
                    w_val = _int_w_attr(el, "w")
                    if w_val is not None:
                        margins[canon] = round(w_val / DXA_PER_CM, 4)
                    break
    return margins if found else None


def _is_borderless_layout_table(
    rows_text: list[list[str]],
    borders: dict[str, dict[str, Any]],
    tbl_node: etree._Element,
) -> bool:
    if len(rows_text) > 2 or not rows_text or len(rows_text[0]) != 2:
        return False
    # Any explicit non-none border on table or cells rules out borderless layout block.
    for info in borders.values():
        if info.get("val") not in ("", "none", "nil"):
            return False
    for b_el in tbl_node.xpath(".//w:tcPr/w:tcBorders/*", namespaces={"w": W}):
        val = (_w_attr(b_el, "val") or "").lower()
        if val not in ("", "none", "nil"):
            return False
    joined = " ".join(cell.lower() for row in rows_text for cell in row)
    return any(kw in joined for kw in _LAYOUT_KEYWORDS)


def extract_tables(
    doc_root: etree._Element,
    sections: list[dict[str, Any]],
    styles_ctx: dict[str, Any],
) -> tuple[list[dict[str, Any]], dict[str, int]]:
    """Extract top-level body tables with resolved widths, styles, and cell metadata."""
    body = doc_root.find(f"{{{W}}}body")
    if body is None:
        return [], {"nested_tables": 0, "container_tables": 0}

    tables: list[dict[str, Any]] = []
    ordinal = 0
    for b_idx, child in enumerate(body):
        if not isinstance(child.tag, str) or etree.QName(child).localname != "tbl":
            continue
        ordinal += 1
        sec = _section_for_body_index(b_idx, sections)
        tbl_style_id, tbl_pr_layers = _resolve_tbl_pr_chain(child, styles_ctx)
        tbl_pr = child.find(f"{{{W}}}tblPr")

        # Resolve alignment across table style + direct tblPr.
        align = "left"
        for layer in tbl_pr_layers:
            jc = layer.find(f"{{{W}}}jc")
            if jc is not None:
                norm_a = _norm_align(_w_attr(jc, "val"))
                if norm_a:
                    align = norm_a

        # Resolve grid columns and table width (dxa, pct, or gridCol fallback).
        grid_cols_dxa = [
            _int_w_attr(c, "w")
            for c in child.xpath("./w:tblGrid/w:gridCol", namespaces={"w": W})
        ]
        valid_grid = [g for g in grid_cols_dxa if g is not None and g > 0]
        grid_sum_cm = round(sum(valid_grid) / DXA_PER_CM, 3) if valid_grid else 0.0

        tbl_w_el = tbl_pr.find(f"{{{W}}}tblW") if tbl_pr is not None else None
        tbl_w_val = _int_w_attr(tbl_w_el, "w")
        tbl_w_type = (_w_attr(tbl_w_el, "type") or ("dxa" if tbl_w_val else "")).lower()

        if tbl_w_type == "dxa" and tbl_w_val and tbl_w_val > 0:
            total_w_cm = round(tbl_w_val / DXA_PER_CM, 3)
            width_source = "tblW_dxa"
        elif tbl_w_type == "pct" and tbl_w_val and tbl_w_val > 0:
            total_w_cm = round(sec["printable_width_cm"] * (tbl_w_val / 5000.0), 3)
            width_source = "tblW_pct"
        elif valid_grid:
            total_w_cm = grid_sum_cm
            width_source = "tblGrid"
        else:
            total_w_cm = 0.0
            width_source = "tcW_fallback"

        if valid_grid and len(valid_grid) == len(grid_cols_dxa):
            # When tblW_pct is scaled differently from gridCol, compute proportional widths if grid missing.
            grid_widths_cm = [round(g / DXA_PER_CM, 3) for g in valid_grid]
            denom = sum(valid_grid)
            grid_widths_pct = [round(g / denom * 100.0, 2) for g in valid_grid]
        else:
            grid_widths_cm = []
            grid_widths_pct = []

        # Row/cell extraction & gridSpan/tcW check.
        rows = child.findall(f"{{{W}}}tr")
        row_texts_nfc: list[list[str]] = []
        has_merged = False
        grid_span_ok = bool(valid_grid and len(valid_grid) == len(grid_cols_dxa))
        header_shd: list[str | None] = []
        data_shd: list[str | None] = []
        header_sizes: list[float] = []
        data_sizes: list[float] = []
        header_all_bold = True
        saw_header_text = False
        cant_split_all = bool(rows)

        for r_idx, tr in enumerate(rows):
            tr_pr = tr.find(f"{{{W}}}trPr")
            cs = tr_pr.find(f"{{{W}}}cantSplit") if tr_pr is not None else None
            if cs is None or _toggle_on(cs) is False:
                cant_split_all = False
            cursor = _int_w_attr(tr_pr.find(f"{{{W}}}gridBefore") if tr_pr is not None else None, "val", 0) or 0
            row_cells_txt: list[str] = []
            for tc in tr.findall(f"{{{W}}}tc"):
                tc_pr = tc.find(f"{{{W}}}tcPr")
                span = max(1, _int_w_attr(tc_pr.find(f"{{{W}}}gridSpan") if tc_pr is not None else None, "val", 1) or 1)
                if span > 1 or (tc_pr is not None and tc_pr.find(f"{{{W}}}vMerge") is not None):
                    has_merged = True

                tc_w = tc_pr.find(f"{{{W}}}tcW") if tc_pr is not None else None
                tc_w_val = _int_w_attr(tc_w, "w")
                tc_w_type = (_w_attr(tc_w, "type") or ("dxa" if tc_w_val else "")).lower()
                if grid_span_ok:
                    if cursor + span > len(valid_grid):
                        grid_span_ok = False
                    else:
                        exp_dxa = sum(valid_grid[cursor : cursor + span])
                        if tc_w_val is None:
                            grid_span_ok = False
                        elif tc_w_type == "dxa":
                            if abs(tc_w_val - exp_dxa) > 30:
                                grid_span_ok = False
                        elif tc_w_type == "pct":
                            exp_pct = exp_dxa / sum(valid_grid)
                            if abs((tc_w_val / 5000.0) - exp_pct) > 0.015:
                                grid_span_ok = False
                        else:
                            grid_span_ok = False
                cursor += span

                shd = tc_pr.find(f"{{{W}}}shd") if tc_pr is not None else None
                fill = (_w_attr(shd, "fill") or "").upper() or None if shd is not None else None
                if r_idx == 0:
                    header_shd.append(fill)
                else:
                    data_shd.append(fill)

                cell_paras = tc.findall(f"{{{W}}}p")
                cell_txt = nfc("\n".join(text_of_element(p) for p in cell_paras))
                row_cells_txt.append(cell_txt)

                for p_el in cell_paras:
                    for run_el in p_el.findall(f"{{{W}}}r"):
                        r_txt = text_of_element(run_el).strip()
                        if not r_txt:
                            continue
                        rf = effective_run_format(
                            run_el,
                            p_el,
                            styles_ctx,
                            tbl_style_id=tbl_style_id,
                            is_header_row=(r_idx == 0),
                        )
                        if r_idx == 0:
                            saw_header_text = True
                            if rf["size_pt"] is not None:
                                header_sizes.append(rf["size_pt"])
                            if not rf["bold"]:
                                header_all_bold = False
                        else:
                            if rf["size_pt"] is not None:
                                data_sizes.append(rf["size_pt"])

            row_texts_nfc.append(row_cells_txt)

        if width_source == "tcW_fallback" and rows:
            first_tcw = []
            for tc in rows[0].findall(f"{{{W}}}tc"):
                tc_w = tc.find(f"{{{W}}}tcPr/{{{W}}}tcW")
                w_v = _int_w_attr(tc_w, "w")
                if w_v and (_w_attr(tc_w, "type") or "dxa") == "dxa":
                    first_tcw.append(w_v)
            if first_tcw:
                total_w_cm = round(sum(first_tcw) / DXA_PER_CM, 3)

        borders = _extract_borders(tbl_pr_layers)
        margins_cm = _extract_tbl_cell_margins_cm(tbl_pr_layers)
        has_mar_overrides = bool(child.xpath(".//w:tcPr/w:tcMar", namespaces={"w": W}))
        repeat_header = bool(rows) and (
            _toggle_on(rows[0].find(f"{{{W}}}trPr/{{{W}}}tblHeader")) is True
        )
        has_tc = (
            child.find(f"{{{W}}}tblGrid/{{{W}}}tblGridChange") is not None
            and child.find(f"{{{W}}}tblPr/{{{W}}}tblPrChange") is not None
        )

        tables.append(
            {
                "ordinal": ordinal,
                "body_index": b_idx,
                "section_index": sec["section_index"],
                "tbl_style_id": tbl_style_id,
                "rows": len(rows),
                "grid_cols": len(grid_cols_dxa) if grid_cols_dxa else (len(row_texts_nfc[0]) if row_texts_nfc else 0),
                "grid_widths_cm": grid_widths_cm,
                "grid_widths_pct": grid_widths_pct,
                "grid_total_width_cm": grid_sum_cm,
                "total_width_cm": total_w_cm,
                "width_source": width_source,
                "tbl_w_type": tbl_w_type or None,
                "tbl_w_raw": tbl_w_val,
                "alignment": align,
                "row_texts_nfc": row_texts_nfc,
                "is_layout_table": _is_borderless_layout_table(row_texts_nfc, borders, child),
                "has_merged_cells": has_merged,
                "grid_span_tcw_consistent": grid_span_ok,
                "cell_margins_cm": margins_cm,
                "has_cell_margin_overrides": has_mar_overrides,
                "repeat_header": repeat_header,
                "cant_split_all_rows": cant_split_all,
                "borders": borders,
                "header_shading_fills": header_shd,
                "data_shading_fills": data_shd,
                "header_font_sizes_pt": header_sizes,
                "header_all_bold": bool(saw_header_text and header_all_bold),
                "data_font_sizes_pt": data_sizes,
                "has_track_changes": has_tc,
            }
        )

    nested_count = len(body.xpath(".//w:tbl//w:tbl", namespaces={"w": W}))
    container_count = len(body.xpath(".//w:sdt//w:tbl | .//w:txbxContent//w:tbl", namespaces={"w": W}))
    return tables, {"nested_tables": nested_count, "container_tables": container_count}


def _iter_active_drawings(p_node: etree._Element) -> list[etree._Element]:
    """Collect w:drawing nodes inside a paragraph without double-counting mc:Fallback."""
    drawings: list[etree._Element] = []
    for d in p_node.xpath(".//w:drawing", namespaces={"w": W}):
        # Skip drawings inside mc:Fallback if sibling mc:Choice exists.
        anc = d.getparent()
        skip = False
        while anc is not None and anc is not p_node:
            if isinstance(anc.tag, str) and anc.tag == f"{{{MC}}}Fallback":
                parent_ac = anc.getparent()
                if parent_ac is not None and parent_ac.find(f"{{{MC}}}Choice") is not None:
                    skip = True
                    break
            anc = anc.getparent()
        if not skip:
            drawings.append(d)
    return drawings


def extract_images(
    doc_root: etree._Element,
    sections: list[dict[str, Any]],
    styles_ctx: dict[str, Any],
) -> list[dict[str, Any]]:
    """Extract inline/anchored DrawingML images from top-level body paragraphs."""
    body = doc_root.find(f"{{{W}}}body")
    if body is None:
        return []
    images: list[dict[str, Any]] = []
    ordinal = 0
    for b_idx, child in enumerate(body):
        if not isinstance(child.tag, str) or etree.QName(child).localname != "p":
            continue
        sec = _section_for_body_index(b_idx, sections)
        p_fmt = effective_paragraph_format(child, styles_ctx)
        drawings = _iter_active_drawings(child)
        for d_idx, d_node in enumerate(drawings):
            box = d_node.find(f"{{{WP}}}inline")
            if box is None:
                box = d_node.find(f"{{{WP}}}anchor")
            if box is None:
                continue
            ext = box.find(f"{{{WP}}}extent")
            if ext is None:
                continue
            try:
                cx = int(ext.get("cx") or "0")
                cy = int(ext.get("cy") or "0")
            except ValueError:
                continue
            a_exts = box.xpath(".//a:xfrm/a:ext", namespaces={"a": A})
            xfrm_ok = bool(a_exts)
            for aext in a_exts:
                try:
                    acx = int(aext.get("cx") or "0")
                    acy = int(aext.get("cy") or "0")
                except ValueError:
                    xfrm_ok = False
                    break
                if abs(acx - cx) > 1000 or abs(acy - cy) > 1000:
                    xfrm_ok = False
                    break

            ordinal += 1
            w_cm = round(cx / EMU_PER_CM, 4)
            h_cm = round(cy / EMU_PER_CM, 4)
            ar = round(cx / cy, 6) if cy > 0 else None
            images.append(
                {
                    "ordinal": ordinal,
                    "body_index": b_idx,
                    "drawing_index_in_p": d_idx,
                    "section_index": sec["section_index"],
                    "wp_cx": cx,
                    "wp_cy": cy,
                    "xfrm_consistent": xfrm_ok,
                    "width_cm": w_cm,
                    "height_cm": h_cm,
                    "aspect_ratio": ar,
                    "alignment": p_fmt.get("align") or "left",
                }
            )
    return images


def extract_paragraphs(
    doc_root: etree._Element,
    sections: list[dict[str, Any]],
    styles_ctx: dict[str, Any],
) -> list[dict[str, Any]]:
    """Extract top-level body paragraphs with effective formatting and bold sequences."""
    body = doc_root.find(f"{{{W}}}body")
    if body is None:
        return []
    paras: list[dict[str, Any]] = []
    ordinal = 0
    for b_idx, child in enumerate(body):
        if not isinstance(child.tag, str) or etree.QName(child).localname != "p":
            continue
        ordinal += 1
        sec = _section_for_body_index(b_idx, sections)
        p_pr = child.find(f"{{{W}}}pPr")
        p_style_el = p_pr.find(f"{{{W}}}pStyle") if p_pr is not None else None
        style_id = _w_attr(p_style_el, "val")
        p_fmt = effective_paragraph_format(child, styles_ctx)
        runs = child.findall(f"{{{W}}}r")

        char_bold_seq: list[tuple[str, bool]] = []
        run_sizes_pt: list[float] = []
        run_fonts: list[str] = []
        for run in runs:
            rf = effective_run_format(run, child, styles_ctx)
            r_txt = text_of_element(run)
            if r_txt.strip():
                if rf["size_pt"] is not None:
                    run_sizes_pt.append(rf["size_pt"])
                if rf["font_ascii"]:
                    run_fonts.append(rf["font_ascii"])
                elif rf["font_hansi"]:
                    run_fonts.append(rf["font_hansi"])
            for ch in "".join(node.text or "" for node in run.findall(f"{{{W}}}t")):
                char_bold_seq.append((nfc(ch), bool(rf["bold"])))

        paras.append(
            {
                "ordinal": ordinal,
                "body_index": b_idx,
                "section_index": sec["section_index"],
                "style_id": style_id,
                "text_nfc": nfc(text_of_element(child)),
                "run_count": len(runs),
                "char_bold_seq": char_bold_seq,
                "effective_format": p_fmt,
                "run_sizes_pt": run_sizes_pt,
                "run_fonts": run_fonts,
            }
        )
    return paras


def extract_app_metadata(parts: dict[str, bytes]) -> dict[str, Any]:
    """Extract authoring application and document timestamps from docProps XML parts."""
    app_name = None
    app_version = None
    pages_hint = None
    creator = None
    last_modified_by = None

    if "docProps/app.xml" in parts:
        try:
            app_root = parse_xml(parts["docProps/app.xml"])
            for el in app_root:
                if not isinstance(el.tag, str):
                    continue
                ln = etree.QName(el).localname
                txt = (el.text or "").strip()
                if ln == "Application" and txt:
                    app_name = txt
                elif ln == "AppVersion" and txt:
                    app_version = txt
                elif ln == "Pages" and txt:
                    try:
                        pages_hint = int(txt)
                    except ValueError:
                        pass
        except etree.XMLSyntaxError:
            pass

    if "docProps/core.xml" in parts:
        try:
            core_root = parse_xml(parts["docProps/core.xml"])
            for el in core_root:
                if not isinstance(el.tag, str):
                    continue
                ln = etree.QName(el).localname
                txt = (el.text or "").strip()
                if ln == "creator" and txt:
                    creator = txt
                elif ln == "lastModifiedBy" and txt:
                    last_modified_by = txt
        except etree.XMLSyntaxError:
            pass

    low = (app_name or "").lower()
    if "microsoft" in low or "word" in low:
        guess = "microsoft_word"
    elif "google" in low:
        guess = "google_docs"
    elif "libreoffice" in low or "openoffice" in low:
        guess = "libreoffice_writer"
    elif "wps" in low or "kingsoft" in low:
        guess = "wps_office"
    else:
        guess = "mixed_or_unknown"

    return {
        "application": app_name,
        "app_version": app_version,
        "pages_hint": pages_hint,
        "creator": creator,
        "last_modified_by": last_modified_by,
        "authoring_software_guess": guess,
    }


def extract_nested_tables_summary(
    doc_root: etree._Element,
    styles_ctx: dict[str, Any],
) -> list[dict[str, Any]]:
    """Extract summary info for nested tables and tables inside sdt/textbox containers."""
    body = doc_root.find(f"{{{W}}}body")
    if body is None:
        return []
    seen_ids: set[int] = set()
    items: list[dict[str, Any]] = []
    xpath_expr = ".//w:tbl//w:tbl | .//w:sdt//w:tbl | .//w:txbxContent//w:tbl"
    for tbl_el in body.xpath(xpath_expr, namespaces={"w": W}):
        obj_id = id(tbl_el)
        if obj_id in seen_ids:
            continue
        seen_ids.add(obj_id)
        kind = "nested"
        anc = tbl_el.getparent()
        while anc is not None and anc is not body:
            if isinstance(anc.tag, str):
                ln = etree.QName(anc).localname
                if ln == "txbxContent":
                    kind = "txbx"
                    break
                if ln == "sdt":
                    kind = "sdt"
                    break
                if ln == "tbl":
                    kind = "nested"
                    break
            anc = anc.getparent()

        tbl_pr = tbl_el.find(f"{{{W}}}tblPr")
        jc = tbl_pr.find(f"{{{W}}}jc") if tbl_pr is not None else None
        align = _norm_align(_w_attr(jc, "val")) or "left"
        rows = tbl_el.findall(f"{{{W}}}tr")
        hdr_shd: list[str | None] = []
        if rows:
            for tc in rows[0].findall(f"{{{W}}}tc"):
                shd = tc.find(f"{{{W}}}tcPr/{{{W}}}shd")
                hdr_shd.append((_w_attr(shd, "fill") or "").upper() or None if shd is not None else None)
        items.append(
            {
                "ordinal": len(items) + 1,
                "kind": kind,
                "alignment": align,
                "rows": len(rows),
                "header_shading_fills": hdr_shd,
                "text_nfc": nfc(text_of_element(tbl_el)),
            }
        )
    return items


def inspect_docx(docx_path: str | Path) -> dict[str, Any]:
    """Inspect a DOCX package independently without importing any repo scripts."""
    meta, parts = read_zip_parts(docx_path)
    styles_ctx = parse_styles(parts.get("word/styles.xml"))
    doc_bytes = parts.get("word/document.xml")
    if doc_bytes is None:
        raise ValueError("Missing word/document.xml in package")
    doc_root = parse_xml(doc_bytes)

    sections = extract_sections(doc_root)
    tables, extra_tbl_counts = extract_tables(doc_root, sections, styles_ctx)
    nested_tables = extract_nested_tables_summary(doc_root, styles_ctx)
    images = extract_images(doc_root, sections, styles_ctx)
    paragraphs = extract_paragraphs(doc_root, sections, styles_ctx)
    app_metadata = extract_app_metadata(parts)

    header_footer_parts: dict[str, dict[str, Any]] = {}
    for name in sorted(parts):
        if name.startswith(("word/header", "word/footer")) and name.endswith(".xml"):
            hf_root = parse_xml(parts[name])
            hf_images = len(hf_root.xpath(".//w:drawing", namespaces={"w": W}))
            header_footer_parts[name] = {
                "text_nfc": nfc(text_of_element(hf_root)),
                "image_count": hf_images,
                "bytes": len(parts[name]),
            }

    notes_text: dict[str, str] = {}
    for note_part in ("word/footnotes.xml", "word/endnotes.xml"):
        if note_part in parts:
            notes_text[note_part] = nfc(text_of_element(parse_xml(parts[note_part])))

    mc_alts = [etree.tostring(el) for el in doc_root.iter(f"{{{MC}}}AlternateContent")]
    mc_ignorable = doc_root.get(f"{{{MC}}}Ignorable") or ""

    return {
        "zip_meta": meta,
        "app_metadata": app_metadata,
        "doc_xml_bytes": len(doc_bytes),
        "root_namespaces": dict(doc_root.nsmap),
        "mc_ignorable": mc_ignorable,
        "mc_alternate_count": len(mc_alts),
        "styles_ctx": styles_ctx,
        "sections": sections,
        "tables": tables,
        "nested_tables": nested_tables,
        "nested_table_count": extra_tbl_counts["nested_tables"],
        "container_table_count": extra_tbl_counts["container_tables"],
        "images": images,
        "paragraphs": paragraphs,
        "headers_footers": header_footer_parts,
        "notes_text_nfc": notes_text,
    }
