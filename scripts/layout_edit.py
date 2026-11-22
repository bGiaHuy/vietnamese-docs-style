"""Safe, anchor-locked OOXML layout edits with pre-flight, read-back, and non-target checks."""
import argparse
import hashlib
import json
import re
from copy import deepcopy
from datetime import datetime, timezone
from pathlib import Path

from lxml import etree

try:
    from . import document_rules, docx_package, map_docx, ooxml_order, validate_docx
except ImportError:
    import document_rules
    import docx_package
    import map_docx
    import ooxml_order
    import validate_docx

W = docx_package.W
WP = map_docx.WP
A = map_docx.A
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'
ALLOWED_OPS = frozenset({
    'table.scale_width',
    'table.set_col_widths',
    'table.set_alignment',
    'table.set_cell_margins',
    'table.set_font_size',
    'table.set_pagination',
    'table.set_borders',
    'table.set_shading',
    'image.scale',
    'paragraph.set_format',
    'paragraph.merge_runs',
    'style.update',
    'section.set_layout',
})
MIN_TEXT_COL_DXA = 288  # ~0.508 cm (0.2 in); narrower text-bearing cells wrap per character.
DXA_PER_CM = map_docx.DXA_PER_CM
EMU_PER_CM = map_docx.EMU_PER_CM
EMU_PER_DXA = EMU_PER_CM / DXA_PER_CM
LENGTH_RE = re.compile(r'^([0-9]+(?:\.[0-9]+)?)\s*(cm|mm|in|pt|dxa|tw)?$', re.IGNORECASE)
MARGIN_SIDES = ('top', 'left', 'bottom', 'right')
BORDER_SIDES = ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')
BORDER_STYLES = frozenset({'single', 'none', 'nil', 'double', 'dashed', 'dotted'})
HEX_COLOR_RE = re.compile(r'^[0-9a-fA-F]{6}$')
ADMIN_PROFILES = frozenset({'administrative', 'minutes-administrative', 'party'})


class LayoutEditError(ValueError):
    """Structured error carrying machine-readable code and remediation hints."""

    def __init__(self, code, message, hints=None):
        super().__init__(message)
        self.code = code
        self.message = message
        self.hints = list(hints or [])

    def to_dict(self):
        return {'status': 'error', 'code': self.code, 'error': self.message, 'hints': self.hints}


def _w_el(tag, attrs=None):
    el = etree.Element(f'{{{W}}}{tag}')
    for key, val in (attrs or {}).items():
        el.set(f'{{{W}}}{key}', str(val))
    return el


def _ensure_pr(parent, tag):
    node = parent.find(f'{{{W}}}{tag}')
    if node is None:
        node = _w_el(tag)
        ooxml_order.insert_ordered_child(parent, node, replace_existing=False)
    return node


def _parse_length_dxa(spec):
    if isinstance(spec, (int, float)) and not isinstance(spec, bool):
        if spec <= 0:
            raise LayoutEditError('INVALID_WIDTH', f'Column width must be positive: {spec}')
        return max(1, round(spec))
    if not isinstance(spec, str):
        raise LayoutEditError('INVALID_WIDTH', f'Unsupported width specification: {spec!r}')
    text = spec.strip()
    m = LENGTH_RE.match(text)
    if not m:
        raise LayoutEditError('INVALID_WIDTH', f'Invalid width syntax: {spec!r}')
    val = float(m.group(1))
    unit = (m.group(2) or 'cm').lower()
    if val <= 0:
        raise LayoutEditError('INVALID_WIDTH', f'Column width must be positive: {spec!r}')
    scale = {
        'cm': DXA_PER_CM,
        'mm': DXA_PER_CM / 10.0,
        'in': 1440.0,
        'pt': 20.0,
        'dxa': 1.0,
        'tw': 1.0,
    }[unit]
    return max(1, round(val * scale))


def _resolve_width_list(widths_spec, n_cols, printable_w_dxa, current_total_dxa, op):
    if not isinstance(widths_spec, list) or len(widths_spec) != n_cols:
        raise LayoutEditError(
            'COL_COUNT_MISMATCH',
            f'Expected {n_cols} grid column widths, got {len(widths_spec) if isinstance(widths_spec, list) else widths_spec!r}',
            ['Supply one value per <w:gridCol> (check `grid_cols` in `vndocs map`, not just row 0 visible cells).'],
        )
    all_pct = all(isinstance(x, str) and x.strip().endswith('%') for x in widths_spec)
    any_pct = any(isinstance(x, str) and x.strip().endswith('%') for x in widths_spec)
    if any_pct and not all_pct:
        raise LayoutEditError('INVALID_WIDTH', 'Cannot mix percentage and absolute width units in one list')

    if all_pct:
        pcts = []
        for raw in widths_spec:
            try:
                pct = float(raw.strip()[:-1])
            except ValueError as exc:
                raise LayoutEditError('INVALID_WIDTH', f'Invalid percentage width: {raw!r}') from exc
            if pct <= 0:
                raise LayoutEditError('INVALID_WIDTH', f'Column percentage must be positive: {raw!r}')
            pcts.append(pct)
        pct_sum = sum(pcts)
        if not 95.0 <= pct_sum <= 105.0:
            raise LayoutEditError('INVALID_WIDTH', f'Column percentages must sum near 100% (got {pct_sum:.1f}%)')
        if op.get('total_width_cm') is not None:
            total = _parse_length_dxa(f"{op['total_width_cm']}cm")
        elif op.get('total_width_pct') is not None:
            total = max(1, round(printable_w_dxa * float(op['total_width_pct']) / 100.0))
        else:
            total = current_total_dxa
        dxa_cols = [max(1, round(total * (p / pct_sum))) for p in pcts]
        dxa_cols[-1] = max(1, total - sum(dxa_cols[:-1]))
        return dxa_cols

    return [_parse_length_dxa(x) for x in widths_spec]


def _cell_text_and_longest_token(tc):
    txt = map_docx.normalize_nfc(' '.join(
        ''.join(node.text or '' for node in p.iter(f'{{{W}}}t'))
        for p in tc.findall(f'{{{W}}}p')
    ))
    tokens = [tok for tok in txt.split() if tok]
    longest = max(tokens, key=len) if tokens else ''
    return txt, longest


def _preflight_table_widths(tbl, anchor, new_cols_dxa, printable_w_dxa):
    """Refuse impossibly narrow text cells (< ~0.51cm) and flag wrap/overflow risks."""
    total_w = sum(new_cols_dxa)
    warnings = []
    if total_w > printable_w_dxa + 20:
        warnings.append({
            'code': 'overflow_printable_width',
            'table': anchor,
            'message': f'Total table width {round(total_w / DXA_PER_CM, 2)} cm exceeds printable width '
                       f'{round(printable_w_dxa / DXA_PER_CM, 2)} cm',
        })

    for r_idx, tr in enumerate(tbl.findall(f'{{{W}}}tr')):
        tr_pr = tr.find(f'{{{W}}}trPr')
        cursor = map_docx._int_attr(tr_pr.find(f'{{{W}}}gridBefore') if tr_pr is not None else None, 'val', 0)
        for tc in tr.findall(f'{{{W}}}tc'):
            tc_pr = tc.find(f'{{{W}}}tcPr')
            span = max(1, map_docx._int_attr(tc_pr.find(f'{{{W}}}gridSpan') if tc_pr is not None else None, 'val', 1))
            if cursor + span > len(new_cols_dxa):
                raise LayoutEditError(
                    'TABLE_STRUCTURE',
                    f'Row {r_idx} exceeds {len(new_cols_dxa)} grid columns on {anchor}',
                )
            cell_w_dxa = sum(new_cols_dxa[cursor:cursor + span])
            txt, longest = _cell_text_and_longest_token(tc)
            if txt and cell_w_dxa < MIN_TEXT_COL_DXA:
                raise LayoutEditError(
                    'COLUMN_TOO_NARROW',
                    f'Column span {cursor}..{cursor + span - 1} of {anchor} would shrink to '
                    f'{round(cell_w_dxa / DXA_PER_CM, 2)} cm (< 0.51 cm) while holding text {txt[:24]!r}',
                    ['Increase target table width or give this column a larger minimum width.'],
                )
            # Heuristic check: ~125 dxa/char (~6.25pt avg Latin/Vietnamese glyph advance) + 216 dxa cell margin.
            if longest and len(longest) >= 4:
                needed_dxa = len(longest) * 125 + 216
                if needed_dxa > cell_w_dxa:
                    warnings.append({
                        'code': 'wrap_risk',
                        'table': anchor,
                        'row': r_idx,
                        'grid_col_start': cursor,
                        'token': longest[:40],
                        'cell_width_cm': round(cell_w_dxa / DXA_PER_CM, 2),
                    })
            cursor += span
    return warnings


def _apply_fixed_grid_and_cells(tbl, new_cols_dxa):
    """Synchronize w:tblGrid, w:tblW, w:tblLayout, and every w:tcW per spanned grid columns."""
    total_w = sum(new_cols_dxa)
    tbl_pr = _ensure_pr(tbl, 'tblPr')
    ooxml_order.insert_ordered_child(tbl_pr, _w_el('tblW', {'w': total_w, 'type': 'dxa'}))
    ooxml_order.insert_ordered_child(tbl_pr, _w_el('tblLayout', {'type': 'fixed'}))

    tbl_grid = tbl.find(f'{{{W}}}tblGrid')
    if tbl_grid is None:
        tbl_grid = _w_el('tblGrid')
        tbl_pr.addnext(tbl_grid)
    for child in list(tbl_grid):
        tbl_grid.remove(child)
    for w_dxa in new_cols_dxa:
        tbl_grid.append(_w_el('gridCol', {'w': w_dxa}))

    for tr in tbl.findall(f'{{{W}}}tr'):
        tr_pr = tr.find(f'{{{W}}}trPr')
        cursor = map_docx._int_attr(tr_pr.find(f'{{{W}}}gridBefore') if tr_pr is not None else None, 'val', 0)
        for tc in tr.findall(f'{{{W}}}tc'):
            tc_pr = _ensure_pr(tc, 'tcPr')
            span = max(1, map_docx._int_attr(tc_pr.find(f'{{{W}}}gridSpan') if tc_pr is not None else None, 'val', 1))
            cell_w = sum(new_cols_dxa[cursor:cursor + span])
            ooxml_order.insert_ordered_child(tc_pr, _w_el('tcW', {'w': cell_w, 'type': 'dxa'}))
            cursor += span


def _set_alignment(tbl, align):
    if align not in ('left', 'center', 'right'):
        raise LayoutEditError('INVALID_ALIGN', f'Alignment must be left, center, or right (got {align!r})')
    tbl_pr = _ensure_pr(tbl, 'tblPr')
    ooxml_order.insert_ordered_child(tbl_pr, _w_el('jc', {'val': align}))


def _apply_cell_margins(tbl, sides_dxa, clear_overrides):
    tbl_pr = _ensure_pr(tbl, 'tblPr')
    cell_mar = tbl_pr.find(f'{{{W}}}tblCellMar')
    if cell_mar is None:
        cell_mar = _w_el('tblCellMar')
        ooxml_order.insert_ordered_child(tbl_pr, cell_mar, replace_existing=False)
    for side, dxa in sides_dxa.items():
        if side == 'left':
            for st in list(cell_mar.findall(f'{{{W}}}start')):
                cell_mar.remove(st)
        elif side == 'right':
            for en in list(cell_mar.findall(f'{{{W}}}end')):
                cell_mar.remove(en)
        ooxml_order.insert_ordered_child(cell_mar, _w_el(side, {'w': dxa, 'type': 'dxa'}))
    ooxml_order.normalize_ordered_children(cell_mar)
    if clear_overrides:
        for tc_pr in tbl.xpath('./w:tr/w:tc/w:tcPr', namespaces={'w': W}):
            for tc_mar in list(tc_pr.findall(f'{{{W}}}tcMar')):
                tc_pr.remove(tc_mar)


def _apply_table_font_size(tbl, size_hp, preserve_header_size):
    """Apply w:sz and w:szCs on all runs (and pPr/rPr) inside target table rows."""
    rows = tbl.findall(f'{{{W}}}tr')
    start_row = 1 if preserve_header_size else 0
    modified_runs = []
    for tr in rows[start_row:]:
        for p in tr.xpath('./w:tc//w:p', namespaces={'w': W}):
            p_pr = p.find(f'{{{W}}}pPr')
            if p_pr is not None:
                p_rpr = p_pr.find(f'{{{W}}}rPr')
                if p_rpr is not None:
                    ooxml_order.insert_ordered_child(p_rpr, _w_el('sz', {'val': size_hp}))
                    ooxml_order.insert_ordered_child(p_rpr, _w_el('szCs', {'val': size_hp}))
                    modified_runs.append(p_rpr)
            for run in p.findall(f'{{{W}}}r'):
                r_pr = _ensure_pr(run, 'rPr')
                ooxml_order.insert_ordered_child(r_pr, _w_el('sz', {'val': size_hp}))
                ooxml_order.insert_ordered_child(r_pr, _w_el('szCs', {'val': size_hp}))
                modified_runs.append(r_pr)
    return modified_runs


def _apply_image_scale(drawing_node, p_node, new_cx, new_cy, align):
    box = drawing_node.find(f'{{{WP}}}inline')
    if box is None:
        box = drawing_node.find(f'{{{WP}}}anchor')
    if box is None:
        raise LayoutEditError('IMAGE_STRUCTURE', 'Drawing does not contain wp:inline or wp:anchor')
    extent = box.find(f'{{{WP}}}extent')
    if extent is None:
        raise LayoutEditError('IMAGE_STRUCTURE', 'Drawing container missing wp:extent')
    extent.set('cx', str(new_cx))
    extent.set('cy', str(new_cy))

    for ext in box.xpath('.//a:xfrm/a:ext', namespaces={'a': A}):
        ext.set('cx', str(new_cx))
        ext.set('cy', str(new_cy))

    if align is not None:
        p_pr = _ensure_pr(p_node, 'pPr')
        ooxml_order.insert_ordered_child(p_pr, _w_el('jc', {'val': align}))


def _apply_paragraph_format(p_node, fmt):
    p_pr = _ensure_pr(p_node, 'pPr')
    if fmt.get('align') is not None:
        ooxml_order.insert_ordered_child(p_pr, _w_el('jc', {'val': fmt['align']}))

    if any(k in fmt for k in ('space_before_pt', 'space_after_pt', 'line_spacing', 'line_spacing_exact_pt')):
        existing_sp = p_pr.find(f'{{{W}}}spacing')
        sp_attrs = dict(existing_sp.attrib) if existing_sp is not None else {}
        if 'space_before_pt' in fmt:
            sp_attrs[f'{{{W}}}before'] = str(round(fmt['space_before_pt'] * 20))
        if 'space_after_pt' in fmt:
            sp_attrs[f'{{{W}}}after'] = str(round(fmt['space_after_pt'] * 20))
        if 'line_spacing' in fmt:
            sp_attrs[f'{{{W}}}line'] = str(round(fmt['line_spacing'] * 240))
            sp_attrs[f'{{{W}}}lineRule'] = 'auto'
        elif 'line_spacing_exact_pt' in fmt:
            sp_attrs[f'{{{W}}}line'] = str(round(fmt['line_spacing_exact_pt'] * 20))
            sp_attrs[f'{{{W}}}lineRule'] = 'exact'
        sp_el = etree.Element(f'{{{W}}}spacing', sp_attrs)
        ooxml_order.insert_ordered_child(p_pr, sp_el)

    if any(k in fmt for k in ('first_line_indent_cm', 'left_indent_cm', 'right_indent_cm')):
        existing_ind = p_pr.find(f'{{{W}}}ind')
        ind_attrs = dict(existing_ind.attrib) if existing_ind is not None else {}
        if 'first_line_indent_cm' in fmt:
            ind_attrs.pop(f'{{{W}}}hanging', None)
            ind_attrs[f'{{{W}}}firstLine'] = str(round(fmt['first_line_indent_cm'] * DXA_PER_CM))
        if 'left_indent_cm' in fmt:
            ind_attrs[f'{{{W}}}left'] = str(round(fmt['left_indent_cm'] * DXA_PER_CM))
        if 'right_indent_cm' in fmt:
            ind_attrs[f'{{{W}}}right'] = str(round(fmt['right_indent_cm'] * DXA_PER_CM))
        ind_el = etree.Element(f'{{{W}}}ind', ind_attrs)
        ooxml_order.insert_ordered_child(p_pr, ind_el)

    if 'keep_with_next' in fmt:
        if fmt['keep_with_next']:
            ooxml_order.insert_ordered_child(p_pr, _w_el('keepNext'))
        else:
            for kn in list(p_pr.findall(f'{{{W}}}keepNext')):
                p_pr.remove(kn)


def _next_rev_id_counter(doc_root):
    max_id = 0
    for elem in doc_root.iter():
        raw = elem.get(f'{{{W}}}id')
        if raw and raw.isdigit():
            max_id = max(max_id, int(raw))
    return [max_id]


def _clean_change_snapshot(node, strip_tags):
    snap = deepcopy(node)
    for tag in strip_tags:
        for ch in list(snap.findall(f'{{{W}}}{tag}')):
            snap.remove(ch)
    return snap


def _record_container_change(parent_pr, snapshot, change_tag, rev_counter, author, date_iso):
    if snapshot is None:
        return
    rev_counter[0] += 1
    change_el = _w_el(change_tag, {
        'id': rev_counter[0],
        'author': author,
        'date': date_iso,
    })
    change_el.append(snapshot)
    if ooxml_order._w_localname(parent_pr) == 'tblGrid':
        for existing in list(parent_pr.findall(f'{{{W}}}{change_tag}')):
            parent_pr.remove(existing)
        parent_pr.append(change_el)
    else:
        ooxml_order.insert_ordered_child(parent_pr, change_el)


def _apply_table_pagination(tbl, repeat_header, header_rows, cant_split):
    rows = tbl.findall(f'{{{W}}}tr')
    if repeat_header is not None:
        for r_idx, tr in enumerate(rows):
            tr_pr = _ensure_pr(tr, 'trPr')
            if repeat_header and r_idx < header_rows:
                ooxml_order.insert_ordered_child(tr_pr, _w_el('tblHeader'))
            else:
                for node in list(tr_pr.findall(f'{{{W}}}tblHeader')):
                    tr_pr.remove(node)
    if cant_split is not None:
        for tr in rows:
            tr_pr = _ensure_pr(tr, 'trPr')
            if cant_split:
                ooxml_order.insert_ordered_child(tr_pr, _w_el('cantSplit'))
            else:
                for node in list(tr_pr.findall(f'{{{W}}}cantSplit')):
                    tr_pr.remove(node)


def _apply_table_borders(tbl, spec):
    tbl_pr = _ensure_pr(tbl, 'tblPr')
    borders = tbl_pr.find(f'{{{W}}}tblBorders')
    if borders is None:
        borders = _w_el('tblBorders')
        ooxml_order.insert_ordered_child(tbl_pr, borders, replace_existing=False)
    style = spec['style']
    for side in spec['sides']:
        if style in ('none', 'nil'):
            attrs = {'val': style}
        else:
            attrs = {
                'val': style,
                'sz': spec['size_eighth_pt'],
                'space': 0,
                'color': spec['color'],
            }
        ooxml_order.insert_ordered_child(borders, _w_el(side, attrs))
    ooxml_order.normalize_ordered_children(borders)


def _apply_table_shading(tbl, spec, track_ctx=None):
    rows = tbl.findall(f'{{{W}}}tr')
    scope = spec['scope']
    fill = spec['fill']
    target_rows = []
    for r_idx, tr in enumerate(rows):
        if scope == 'all':
            target_rows.append(tr)
        elif scope == 'header' and r_idx == 0:
            target_rows.append(tr)
        elif scope == 'odd_rows' and r_idx % 2 == 1:
            target_rows.append(tr)
        elif scope == 'even_rows' and r_idx % 2 == 0:
            target_rows.append(tr)
        elif scope == 'row' and r_idx == spec['row_index']:
            target_rows.append(tr)

    touched = []
    for tr in target_rows:
        for tc in tr.findall(f'{{{W}}}tc'):
            tc_pr = _ensure_pr(tc, 'tcPr')
            snap = _clean_change_snapshot(tc_pr, ('tcPrChange',)) if track_ctx else None
            if fill == 'clear':
                for shd in list(tc_pr.findall(f'{{{W}}}shd')):
                    tc_pr.remove(shd)
            else:
                ooxml_order.insert_ordered_child(
                    tc_pr,
                    _w_el('shd', {'val': 'clear', 'color': 'auto', 'fill': fill}),
                )
            if track_ctx:
                _record_container_change(
                    tc_pr, snap, 'tcPrChange',
                    track_ctx['rev'], track_ctx['author'], track_ctx['date'],
                )
            touched.append(tc_pr)
    return touched


def _apply_style_update(style_node, spec):
    if any(k in spec for k in ('align', 'space_before_pt', 'space_after_pt', 'line_spacing', 'line_spacing_exact_pt', 'first_line_indent_cm')):
        p_pr = style_node.find(f'{{{W}}}pPr')
        if p_pr is None:
            p_pr = _w_el('pPr')
            r_pr_existing = style_node.find(f'{{{W}}}rPr')
            if r_pr_existing is not None:
                style_node.insert(list(style_node).index(r_pr_existing), p_pr)
            else:
                style_node.append(p_pr)
        fake_p = etree.Element(f'{{{W}}}p')
        fake_p.append(p_pr)
        _apply_paragraph_format(fake_p, spec)
        # Reattach directly to style_node keeping position
        style_idx = style_node.find(f'{{{W}}}rPr')
        if style_idx is not None:
            style_node.insert(list(style_node).index(style_idx), p_pr)
        else:
            style_node.append(p_pr)

    if any(k in spec for k in ('font', 'size_pt', 'color', 'bold', 'italic')):
        r_pr = style_node.find(f'{{{W}}}rPr')
        if r_pr is None:
            r_pr = _w_el('rPr')
            style_node.append(r_pr)
        if spec.get('font') is not None:
            fn = spec['font']
            ooxml_order.insert_ordered_child(r_pr, _w_el('rFonts', {'ascii': fn, 'hAnsi': fn, 'cs': fn}))
        if spec.get('bold') is not None:
            if spec['bold']:
                ooxml_order.insert_ordered_child(r_pr, _w_el('b'))
            else:
                ooxml_order.insert_ordered_child(r_pr, _w_el('b', {'val': '0'}))
        if spec.get('italic') is not None:
            if spec['italic']:
                ooxml_order.insert_ordered_child(r_pr, _w_el('i'))
            else:
                ooxml_order.insert_ordered_child(r_pr, _w_el('i', {'val': '0'}))
        if spec.get('color') is not None:
            ooxml_order.insert_ordered_child(r_pr, _w_el('color', {'val': spec['color']}))
        if spec.get('size_pt') is not None:
            hp = str(round(spec['size_pt'] * 2))
            ooxml_order.insert_ordered_child(r_pr, _w_el('sz', {'val': hp}))
            ooxml_order.insert_ordered_child(r_pr, _w_el('szCs', {'val': hp}))
        ooxml_order.normalize_ordered_children(r_pr)


def _apply_section_layout(sect_node, spec):
    if spec.get('orientation') is not None or spec.get('page_size') is not None:
        pg_sz = sect_node.find(f'{{{W}}}pgSz')
        cur_w = map_docx._int_attr(pg_sz, 'w', round(21.0 * DXA_PER_CM))
        cur_h = map_docx._int_attr(pg_sz, 'h', round(29.7 * DXA_PER_CM))
        if spec.get('page_size') == 'A4':
            base_w, base_h = 11906, 16838
        elif spec.get('page_size') == 'Letter':
            base_w, base_h = 12240, 15840
        else:
            base_w, base_h = min(cur_w, cur_h), max(cur_w, cur_h)
        orient = spec.get('orientation')
        if orient is None:
            orient = 'landscape' if cur_w > cur_h else 'portrait'
        attrs = (
            {'w': max(base_w, base_h), 'h': min(base_w, base_h), 'orient': 'landscape'}
            if orient == 'landscape'
            else {'w': min(base_w, base_h), 'h': max(base_w, base_h)}
        )
        ooxml_order.insert_ordered_child(sect_node, _w_el('pgSz', attrs))

    if spec.get('margins_cm'):
        existing = sect_node.find(f'{{{W}}}pgMar')
        mar_attrs = dict(existing.attrib) if existing is not None else {
            f'{{{W}}}top': '1134', f'{{{W}}}right': '851',
            f'{{{W}}}bottom': '1134', f'{{{W}}}left': '1701',
            f'{{{W}}}header': '567', f'{{{W}}}footer': '567',
        }
        for key, cm_val in spec['margins_cm'].items():
            mar_attrs[f'{{{W}}}{key}'] = str(round(cm_val * DXA_PER_CM))
        ooxml_order.insert_ordered_child(sect_node, etree.Element(f'{{{W}}}pgMar', mar_attrs))
    ooxml_order.normalize_ordered_children(sect_node)


def _is_pure_text_run(run_el):
    if not isinstance(run_el.tag, str) or etree.QName(run_el).localname != 'r':
        return False
    for ch in run_el:
        if not isinstance(ch.tag, str):
            continue
        if etree.QName(ch).localname not in ('rPr', 't'):
            return False
    return True


def _run_rpr_key(run_el):
    r_pr = run_el.find(f'{{{W}}}rPr')
    if r_pr is None or len(r_pr) == 0:
        return b''
    return etree.tostring(r_pr)


def _apply_merge_runs(p_node):
    merged_count = 0
    children = list(p_node)
    idx = 0
    while idx < len(children) - 1:
        cur = children[idx]
        nxt = children[idx + 1]
        if _is_pure_text_run(cur) and _is_pure_text_run(nxt) and _run_rpr_key(cur) == _run_rpr_key(nxt):
            cur_text = ''.join(t.text or '' for t in cur.findall(f'{{{W}}}t'))
            nxt_text = ''.join(t.text or '' for t in nxt.findall(f'{{{W}}}t'))
            combined = cur_text + nxt_text
            for extra_t in list(cur.findall(f'{{{W}}}t'))[1:]:
                cur.remove(extra_t)
            t_node = cur.find(f'{{{W}}}t')
            if t_node is None:
                t_node = _w_el('t')
                cur.append(t_node)
            t_node.text = combined
            if combined != combined.strip():
                t_node.set(XML_SPACE, 'preserve')
            p_node.remove(nxt)
            children.pop(idx + 1)
            merged_count += 1
        else:
            idx += 1
    return merged_count


def _verify_table_readback(item):
    tbl = item['tbl_node']
    expected_cols_dxa = item.get('new_cols')
    expected_align = item.get('new_align')
    if expected_cols_dxa is not None:
        actual_cols = [
            map_docx._int_attr(c, 'w')
            for c in tbl.xpath('./w:tblGrid/w:gridCol', namespaces={'w': W})
        ]
        if actual_cols != expected_cols_dxa:
            raise LayoutEditError('READBACK_MISMATCH', f'gridCol mismatch: expected {expected_cols_dxa}, got {actual_cols}')
        tbl_w = tbl.find(f'{{{W}}}tblPr/{{{W}}}tblW')
        if map_docx._int_attr(tbl_w, 'w') != sum(expected_cols_dxa) or tbl_w.get(f'{{{W}}}type') != 'dxa':
            raise LayoutEditError('READBACK_MISMATCH', 'tblW read-back mismatch after mutation')
        layout = tbl.find(f'{{{W}}}tblPr/{{{W}}}tblLayout')
        if layout is None or layout.get(f'{{{W}}}type') != 'fixed':
            raise LayoutEditError('READBACK_MISMATCH', 'tblLayout is not fixed after mutation')
    if expected_align is not None:
        jc = tbl.find(f'{{{W}}}tblPr/{{{W}}}jc')
        if jc is None or jc.get(f'{{{W}}}val') != expected_align:
            raise LayoutEditError('READBACK_MISMATCH', f'Alignment read-back mismatch (expected {expected_align})')
    if item.get('sides_dxa'):
        cell_mar = tbl.find(f'{{{W}}}tblPr/{{{W}}}tblCellMar')
        if cell_mar is None:
            raise LayoutEditError('READBACK_MISMATCH', 'tblCellMar missing after mutation')
        for side, dxa in item['sides_dxa'].items():
            node = cell_mar.find(f'{{{W}}}{side}')
            if node is None or map_docx._int_attr(node, 'w') != dxa:
                raise LayoutEditError('READBACK_MISMATCH', f'tblCellMar/{side} read-back mismatch')
        if item.get('clear_overrides') and tbl.xpath('./w:tr/w:tc/w:tcPr/w:tcMar', namespaces={'w': W}):
            raise LayoutEditError('READBACK_MISMATCH', 'Cell tcMar override remained after clear_cell_overrides=True')
    if item.get('size_hp') is not None:
        hp = str(item['size_hp'])
        for r_pr in item.get('modified_rprs', []):
            sz = r_pr.find(f'{{{W}}}sz')
            sz_cs = r_pr.find(f'{{{W}}}szCs')
            if sz is None or sz.get(f'{{{W}}}val') != hp or sz_cs is None or sz_cs.get(f'{{{W}}}val') != hp:
                raise LayoutEditError('READBACK_MISMATCH', f'Run font size mismatch (expected {hp})')
    if item.get('pagination') is not None:
        pag = item['pagination']
        rows = tbl.findall(f'{{{W}}}tr')
        if pag['repeat_header'] is not None:
            for r_idx, tr in enumerate(rows):
                has_hdr = tr.find(f'{{{W}}}trPr/{{{W}}}tblHeader') is not None
                want = pag['repeat_header'] and r_idx < pag['header_rows']
                if has_hdr != want:
                    raise LayoutEditError('READBACK_MISMATCH', f'tblHeader mismatch at row {r_idx}')
        if pag['cant_split'] is not None:
            for r_idx, tr in enumerate(rows):
                has_cs = tr.find(f'{{{W}}}trPr/{{{W}}}cantSplit') is not None
                if has_cs != pag['cant_split']:
                    raise LayoutEditError('READBACK_MISMATCH', f'cantSplit mismatch at row {r_idx}')
    if item.get('border_spec') is not None:
        bspec = item['border_spec']
        borders = tbl.find(f'{{{W}}}tblPr/{{{W}}}tblBorders')
        if borders is None:
            raise LayoutEditError('READBACK_MISMATCH', 'tblBorders missing after mutation')
        for side in bspec['sides']:
            el = borders.find(f'{{{W}}}{side}')
            if el is None or el.get(f'{{{W}}}val') != bspec['style']:
                raise LayoutEditError('READBACK_MISMATCH', f'tblBorders/{side} read-back mismatch')
    if item.get('shading_spec') is not None:
        fill = item['shading_spec']['fill']
        for tc_pr in item.get('modified_shd_tcprs', []):
            shd = tc_pr.find(f'{{{W}}}shd')
            if fill == 'clear':
                if shd is not None:
                    raise LayoutEditError('READBACK_MISMATCH', 'w:shd still present after clear')
            elif shd is None or shd.get(f'{{{W}}}fill') != fill:
                raise LayoutEditError('READBACK_MISMATCH', f'w:shd fill mismatch (expected {fill})')


def _verify_image_readback(item):
    drawing = item['drawing_node']
    box = drawing.find(f'{{{WP}}}inline')
    if box is None:
        box = drawing.find(f'{{{WP}}}anchor')
    extent = box.find(f'{{{WP}}}extent')
    if extent.get('cx') != str(item['new_cx']) or extent.get('cy') != str(item['new_cy']):
        raise LayoutEditError('READBACK_MISMATCH', 'wp:extent dimensions mismatch after image.scale')
    for ext in box.xpath('.//a:xfrm/a:ext', namespaces={'a': A}):
        if ext.get('cx') != str(item['new_cx']) or ext.get('cy') != str(item['new_cy']):
            raise LayoutEditError('READBACK_MISMATCH', 'a:ext dimensions mismatch after image.scale')
    if item.get('new_align') is not None:
        jc = item['p_node'].find(f'{{{W}}}pPr/{{{W}}}jc')
        if jc is None or jc.get(f'{{{W}}}val') != item['new_align']:
            raise LayoutEditError('READBACK_MISMATCH', 'Paragraph alignment mismatch after image.scale')


def _verify_paragraph_readback(item):
    p_pr = item['p_node'].find(f'{{{W}}}pPr')
    if p_pr is None:
        raise LayoutEditError('READBACK_MISMATCH', f'pPr missing on {item["target"]}')
    fmt = item['fmt']
    if fmt.get('align') is not None:
        jc = p_pr.find(f'{{{W}}}jc')
        if jc is None or jc.get(f'{{{W}}}val') != fmt['align']:
            raise LayoutEditError('READBACK_MISMATCH', f'Paragraph jc mismatch on {item["target"]}')
    if 'line_spacing' in fmt or 'line_spacing_exact_pt' in fmt or 'space_before_pt' in fmt or 'space_after_pt' in fmt:
        sp = p_pr.find(f'{{{W}}}spacing')
        if sp is None:
            raise LayoutEditError('READBACK_MISMATCH', f'w:spacing missing on {item["target"]}')
        if 'line_spacing' in fmt and (sp.get(f'{{{W}}}lineRule') != 'auto' or sp.get(f'{{{W}}}line') != str(round(fmt['line_spacing'] * 240))):
            raise LayoutEditError('READBACK_MISMATCH', f'line_spacing mismatch on {item["target"]}')
        if 'line_spacing_exact_pt' in fmt and (sp.get(f'{{{W}}}lineRule') != 'exact' or sp.get(f'{{{W}}}line') != str(round(fmt['line_spacing_exact_pt'] * 20))):
            raise LayoutEditError('READBACK_MISMATCH', f'line_spacing_exact_pt mismatch on {item["target"]}')


def _verify_merge_readback(item):
    if docx_package.text_of(item['p_node']) != item['before_text']:
        raise LayoutEditError('READBACK_MISMATCH', f'Paragraph text changed during merge_runs on {item["target"]}')


def _verify_style_readback(item):
    st = item['style_node']
    spec = item['style_spec']
    if spec.get('size_pt') is not None:
        hp = str(round(spec['size_pt'] * 2))
        sz = st.find(f'{{{W}}}rPr/{{{W}}}sz')
        if sz is None or sz.get(f'{{{W}}}val') != hp:
            raise LayoutEditError('READBACK_MISMATCH', f'Style {item["style_name"]} sz mismatch')
    if spec.get('font') is not None:
        rf = st.find(f'{{{W}}}rPr/{{{W}}}rFonts')
        if rf is None or rf.get(f'{{{W}}}ascii') != spec['font']:
            raise LayoutEditError('READBACK_MISMATCH', f'Style {item["style_name"]} rFonts mismatch')
    if spec.get('color') is not None:
        col = st.find(f'{{{W}}}rPr/{{{W}}}color')
        if col is None or col.get(f'{{{W}}}val') != spec['color']:
            raise LayoutEditError('READBACK_MISMATCH', f'Style {item["style_name"]} color mismatch')


def _verify_section_readback(item):
    sect = item['sect_node']
    spec = item['section_spec']
    if spec.get('orientation') is not None:
        pg_sz = sect.find(f'{{{W}}}pgSz')
        w = map_docx._int_attr(pg_sz, 'w')
        h = map_docx._int_attr(pg_sz, 'h')
        if w is None or h is None:
            raise LayoutEditError('READBACK_MISMATCH', 'pgSz missing after section.set_layout')
        if spec['orientation'] == 'landscape' and w <= h:
            raise LayoutEditError('READBACK_MISMATCH', 'Section width <= height after landscape change')
        if spec['orientation'] == 'portrait' and w >= h:
            raise LayoutEditError('READBACK_MISMATCH', 'Section width >= height after portrait change')
    if spec.get('margins_cm'):
        pg_mar = sect.find(f'{{{W}}}pgMar')
        if pg_mar is None:
            raise LayoutEditError('READBACK_MISMATCH', 'pgMar missing after section.set_layout')
        for key, cm_val in spec['margins_cm'].items():
            if map_docx._int_attr(pg_mar, key) != round(cm_val * DXA_PER_CM):
                raise LayoutEditError('READBACK_MISMATCH', f'Section margin {key} read-back mismatch')


def _check_mutated_order(prepared):
    violations = []
    for item in prepared:
        kind = item['kind']
        if kind == 'table':
            tbl_node = item['tbl_node']
            for idx, ch in enumerate(tbl_node):
                if isinstance(ch.tag, str) and etree.QName(ch).localname == 'tblPr' and idx > 0:
                    violations.append({'message': f'<w:tblPr> must be child 0 of <w:tbl> (found at {idx})'})
            tbl_pr = tbl_node.find(f'{{{W}}}tblPr')
            if tbl_pr is not None:
                violations.extend(ooxml_order.check_element_order(tbl_pr))
            if item.get('pagination') is not None:
                for tr in tbl_node.findall(f'{{{W}}}tr'):
                    for idx, ch in enumerate(tr):
                        if isinstance(ch.tag, str) and etree.QName(ch).localname == 'trPr' and idx > 0:
                            violations.append({'message': f'<w:trPr> must be child 0 of <w:tr> (found at {idx})'})
                    tr_pr = tr.find(f'{{{W}}}trPr')
                    if tr_pr is not None:
                        violations.extend(ooxml_order.check_element_order(tr_pr))
            if item.get('new_cols') is not None or item.get('clear_overrides') or item.get('shading_spec') is not None:
                for tc in tbl_node.xpath('./w:tr/w:tc', namespaces={'w': W}):
                    for idx, ch in enumerate(tc):
                        if isinstance(ch.tag, str) and etree.QName(ch).localname == 'tcPr' and idx > 0:
                            violations.append({'message': f'<w:tcPr> must be child 0 of <w:tc> (found at {idx})'})
                    tc_pr = tc.find(f'{{{W}}}tcPr')
                    if tc_pr is not None:
                        violations.extend(ooxml_order.check_element_order(tc_pr))
            if item.get('size_hp') is not None:
                for run in tbl_node.xpath('./w:tr/w:tc//w:r', namespaces={'w': W}):
                    for idx, ch in enumerate(run):
                        if isinstance(ch.tag, str) and etree.QName(ch).localname == 'rPr' and idx > 0:
                            violations.append({'message': f'<w:rPr> must be child 0 of <w:r> (found at {idx})'})
                    r_pr = run.find(f'{{{W}}}rPr')
                    if r_pr is not None:
                        violations.extend(ooxml_order.check_element_order(r_pr))
        elif kind == 'style':
            st = item['style_node']
            violations.extend(ooxml_order.check_element_order(st))
            child_names = [etree.QName(ch).localname for ch in st if isinstance(ch.tag, str)]
            if 'pPr' in child_names and 'rPr' in child_names and child_names.index('pPr') > child_names.index('rPr'):
                violations.append({'message': '<w:pPr> must precede <w:rPr> inside <w:style>'})
        elif kind == 'section':
            violations.extend(ooxml_order.check_element_order(item['sect_node']))
            if item.get('p_node') is not None:
                p_pr = item['p_node'].find(f'{{{W}}}pPr')
                if p_pr is not None:
                    violations.extend(ooxml_order.check_element_order(p_pr))
        else:
            p_node = item['p_node']
            for idx, ch in enumerate(p_node):
                if isinstance(ch.tag, str) and etree.QName(ch).localname == 'pPr' and idx > 0:
                    violations.append({'message': f'<w:pPr> must be child 0 of <w:p> (found at {idx})'})
            p_pr = p_node.find(f'{{{W}}}pPr')
            if p_pr is not None:
                violations.extend(ooxml_order.check_element_order(p_pr))
    return violations


def _expand_operations(ops):
    """Expand multi-target paragraph ops into uniform per-target items."""
    expanded = []
    for idx, op in enumerate(ops):
        if not isinstance(op, dict) or op.get('op') not in ALLOWED_OPS:
            raise LayoutEditError(
                'UNSUPPORTED_OP',
                f'Operation {idx} is not in allowlist {sorted(ALLOWED_OPS)}: {op.get("op") if isinstance(op, dict) else op!r}',
            )
        if op['op'] in ('paragraph.set_format', 'paragraph.merge_runs') and op.get('targets') is not None:
            targets = op['targets']
            if not isinstance(targets, list) or not targets:
                raise LayoutEditError('INVALID_PLAN', f"{op['op']} targets must be a nonempty list")
            for tgt in targets:
                clone = dict(op)
                clone['target'] = tgt
                clone.pop('targets', None)
                expanded.append(clone)
        else:
            expanded.append(op)
    return expanded


def _validate_float_range(val, low, high, code, name):
    try:
        num = float(val)
    except (TypeError, ValueError) as exc:
        raise LayoutEditError(code, f'{name} must be numeric, got {val!r}') from exc
    if not low <= num <= high:
        raise LayoutEditError(code, f'{name} must be in [{low}, {high}], got {num}')
    return num


def _validate_paragraph_fmt(op, *, require_nonempty=True):
    if op.get('line_spacing') is not None and op.get('line_spacing_exact_pt') is not None:
        raise LayoutEditError('INVALID_SPACING', 'Cannot set both line_spacing and line_spacing_exact_pt')
    fmt = {}
    if op.get('align') is not None:
        raw_align = str(op['align']).lower()
        align = 'both' if raw_align == 'justify' else raw_align
        if align not in ('left', 'center', 'right', 'both'):
            raise LayoutEditError('INVALID_ALIGN', f'Unsupported paragraph align: {op["align"]!r}')
        fmt['align'] = align
    if op.get('space_before_pt') is not None:
        fmt['space_before_pt'] = _validate_float_range(op['space_before_pt'], 0.0, 120.0, 'INVALID_SPACING', 'space_before_pt')
    if op.get('space_after_pt') is not None:
        fmt['space_after_pt'] = _validate_float_range(op['space_after_pt'], 0.0, 120.0, 'INVALID_SPACING', 'space_after_pt')
    if op.get('line_spacing') is not None:
        fmt['line_spacing'] = _validate_float_range(op['line_spacing'], 0.5, 4.0, 'INVALID_SPACING', 'line_spacing')
    if op.get('line_spacing_exact_pt') is not None:
        fmt['line_spacing_exact_pt'] = _validate_float_range(op['line_spacing_exact_pt'], 8.0, 72.0, 'INVALID_SPACING', 'line_spacing_exact_pt')
    if op.get('first_line_indent_cm') is not None:
        fmt['first_line_indent_cm'] = _validate_float_range(op['first_line_indent_cm'], 0.0, 5.0, 'INVALID_INDENT', 'first_line_indent_cm')
    if op.get('left_indent_cm') is not None:
        fmt['left_indent_cm'] = _validate_float_range(op['left_indent_cm'], 0.0, 10.0, 'INVALID_INDENT', 'left_indent_cm')
    if op.get('right_indent_cm') is not None:
        fmt['right_indent_cm'] = _validate_float_range(op['right_indent_cm'], 0.0, 10.0, 'INVALID_INDENT', 'right_indent_cm')
    if op.get('keep_with_next') is not None:
        if not isinstance(op['keep_with_next'], bool):
            raise LayoutEditError('INVALID_PLAN', 'keep_with_next must be boolean')
        fmt['keep_with_next'] = op['keep_with_next']
    if require_nonempty and not fmt:
        raise LayoutEditError('INVALID_PLAN', 'paragraph.set_format requires at least one formatting property')
    return fmt


def _find_style_node(styles_root, style_name):
    if styles_root is None or not style_name:
        return None
    want = str(style_name).strip()
    for st in styles_root.findall(f'{{{W}}}style'):
        if st.get(f'{{{W}}}styleId') == want:
            return st
    want_low = want.lower()
    for st in styles_root.findall(f'{{{W}}}style'):
        nm = st.find(f'{{{W}}}name')
        if nm is not None and (nm.get(f'{{{W}}}val') or '').strip().lower() == want_low:
            return st
    return None


def _locate_sections_xml(body):
    sections = []
    sec_idx = 0
    for b_idx, child in enumerate(body):
        qname = etree.QName(child).localname
        if qname == 'sectPr':
            sec_idx += 1
            sections.append((sec_idx, child, b_idx, None))
        elif qname == 'p':
            p_sect = child.find(f'{{{W}}}pPr/{{{W}}}sectPr')
            if p_sect is not None:
                sec_idx += 1
                sections.append((sec_idx, p_sect, b_idx, child))
    return sections


def _prepare_operations(ops, mapping, body, tbl_elements, allow_layout_tables, profile, styles_root=None):
    tables_by_anchor = {t['anchor']: t for t in mapping['tables']}
    images_by_anchor = {img['anchor']: img for img in mapping.get('images', [])}
    paragraphs_by_anchor = {p['anchor']: p for p in mapping['paragraphs']}
    sections_xml = _locate_sections_xml(body)
    body_children = list(body)

    prepared = []
    warnings = []
    allowed_body_indices = set()

    for op in _expand_operations(ops):
        op_name = op['op']
        target = op.get('target')

        if op_name.startswith('table.'):
            if target not in tables_by_anchor:
                raise LayoutEditError(
                    'ANCHOR_STALE',
                    f'Target table anchor {target!r} not found or stale in document',
                    [f"Available tables: {[t['anchor'] for t in mapping['tables']]}"],
                )
            t_info = tables_by_anchor[target]
            if t_info['is_layout_table'] and not allow_layout_tables:
                raise LayoutEditError(
                    'LAYOUT_TABLE_PROTECTED',
                    f'Table {target} is a protected layout table ({t_info["layout_role"]})',
                    ['Target a data table, or pass --allow-layout-tables explicitly.'],
                )
            tbl_node = tbl_elements[t_info['ordinal'] - 1]
            cur_cols = list(t_info['grid_widths_dxa'])
            cur_total = sum(cur_cols)
            printable_w = t_info['printable_width_dxa']

            item = {
                'kind': 'table',
                'op': op_name,
                'target': target,
                'tbl_node': tbl_node,
                'before_width_cm': t_info['total_width_cm'],
                'before_cols_cm': t_info['grid_widths_cm'],
                'before_align': t_info['alignment'],
                'new_cols': None,
                'new_align': None,
            }

            if op_name == 'table.scale_width':
                if op.get('factor') is not None:
                    factor = _validate_float_range(op['factor'], 0.2, 2.0, 'INVALID_FACTOR', 'factor')
                    target_total = max(1, round(cur_total * factor))
                elif op.get('target_width_pct') is not None:
                    pct = _validate_float_range(op['target_width_pct'], 10.0, 100.0, 'INVALID_WIDTH', 'target_width_pct')
                    target_total = max(1, round(printable_w * pct / 100.0))
                elif op.get('target_width_cm') is not None:
                    target_total = _parse_length_dxa(f"{op['target_width_cm']}cm")
                else:
                    raise LayoutEditError('INVALID_PLAN', 'table.scale_width requires factor, target_width_pct, or target_width_cm')
                new_cols = [max(1, round(target_total * (w / cur_total))) for w in cur_cols]
                new_cols[-1] = max(1, target_total - sum(new_cols[:-1]))
                warnings.extend(_preflight_table_widths(tbl_node, target, new_cols, printable_w))
                item['new_cols'] = new_cols

            elif op_name == 'table.set_col_widths':
                new_cols = _resolve_width_list(op.get('widths'), len(cur_cols), printable_w, cur_total, op)
                warnings.extend(_preflight_table_widths(tbl_node, target, new_cols, printable_w))
                item['new_cols'] = new_cols

            elif op_name == 'table.set_alignment':
                align = op.get('align')
                if align not in ('left', 'center', 'right'):
                    raise LayoutEditError('INVALID_ALIGN', f'table.set_alignment requires left, center, or right (got {align!r})')
                item['new_align'] = align

            elif op_name == 'table.set_cell_margins':
                sides_dxa = {}
                for side in MARGIN_SIDES:
                    key = f'{side}_cm'
                    if op.get(key) is not None:
                        val = _validate_float_range(op[key], 0.0, 1.5, 'INVALID_MARGIN', key)
                        sides_dxa[side] = round(val * DXA_PER_CM)
                if not sides_dxa:
                    raise LayoutEditError('INVALID_PLAN', 'table.set_cell_margins requires at least one of top_cm/bottom_cm/left_cm/right_cm')
                item['sides_dxa'] = sides_dxa
                item['clear_overrides'] = bool(op.get('clear_cell_overrides', True))

            elif op_name == 'table.set_font_size':
                size_pt = _validate_float_range(op.get('size_pt'), 6.0, 48.0, 'INVALID_FONT_SIZE', 'size_pt')
                if profile in ADMIN_PROFILES and size_pt < 12.0:
                    warnings.append({
                        'code': 'small_table_font_warning',
                        'table': target,
                        'size_pt': size_pt,
                        'message': f'Table font size {size_pt}pt is below 12pt minimum convention for profile {profile}',
                    })
                item['size_hp'] = round(size_pt * 2)
                item['size_pt'] = size_pt
                item['preserve_header_size'] = bool(op.get('preserve_header_size', False))

            elif op_name == 'table.set_pagination':
                repeat_header = op.get('repeat_header')
                cant_split = op.get('cant_split')
                if repeat_header is None and cant_split is None:
                    raise LayoutEditError('INVALID_PLAN', 'table.set_pagination requires repeat_header and/or cant_split')
                if repeat_header is not None and not isinstance(repeat_header, bool):
                    raise LayoutEditError('INVALID_PLAN', 'repeat_header must be boolean')
                if cant_split is not None and not isinstance(cant_split, bool):
                    raise LayoutEditError('INVALID_PLAN', 'cant_split must be boolean')
                header_rows = op.get('header_rows', 1)
                n_rows = len(tbl_node.findall(f'{{{W}}}tr'))
                if not isinstance(header_rows, int) or isinstance(header_rows, bool) or not 1 <= header_rows <= max(1, n_rows):
                    raise LayoutEditError('INVALID_PLAN', f'header_rows must be an int in [1, {n_rows}]')
                item['pagination'] = {
                    'repeat_header': repeat_header,
                    'header_rows': header_rows,
                    'cant_split': cant_split,
                }

            elif op_name == 'table.set_borders':
                style = str(op.get('style', 'single')).lower()
                if style not in BORDER_STYLES:
                    raise LayoutEditError('INVALID_PLAN', f'Unsupported border style: {style!r}')
                sz_raw = op.get('size_eighth_pt', 4)
                if not isinstance(sz_raw, int) or isinstance(sz_raw, bool) or not 1 <= sz_raw <= 96:
                    raise LayoutEditError('INVALID_PLAN', f'size_eighth_pt must be an int in [1, 96], got {sz_raw!r}')
                color = str(op.get('color', '000000'))
                if color.lower() != 'auto':
                    if not HEX_COLOR_RE.match(color):
                        raise LayoutEditError('INVALID_PLAN', f'Border color must be 6-hex or auto, got {color!r}')
                    color = color.upper()
                else:
                    color = 'auto'
                sides = op.get('sides') or list(BORDER_SIDES)
                if not isinstance(sides, list) or not sides or any(s not in BORDER_SIDES for s in sides):
                    raise LayoutEditError('INVALID_PLAN', f'Border sides must be subset of {BORDER_SIDES}')
                item['border_spec'] = {
                    'style': style,
                    'size_eighth_pt': sz_raw,
                    'color': color,
                    'sides': sides,
                }

            elif op_name == 'table.set_shading':
                raw_fill = str(op.get('fill') or '')
                if raw_fill.lower() in ('clear', 'auto'):
                    fill = 'clear'
                elif HEX_COLOR_RE.match(raw_fill):
                    fill = raw_fill.upper()
                else:
                    raise LayoutEditError('INVALID_PLAN', f'table.set_shading fill must be 6-hex or "clear", got {raw_fill!r}')
                scope = op.get('scope', 'all')
                if scope not in ('header', 'all', 'odd_rows', 'even_rows', 'row'):
                    raise LayoutEditError('INVALID_PLAN', f'Unsupported shading scope: {scope!r}')
                row_index = op.get('row_index')
                n_rows = len(tbl_node.findall(f'{{{W}}}tr'))
                if scope == 'row':
                    if not isinstance(row_index, int) or isinstance(row_index, bool) or not 0 <= row_index < n_rows:
                        raise LayoutEditError('INVALID_PLAN', f'row_index must be in [0, {n_rows - 1}]')
                item['shading_spec'] = {'fill': fill, 'scope': scope, 'row_index': row_index}

            allowed_body_indices.add(t_info['body_index'])
            prepared.append(item)

        elif op_name == 'image.scale':
            if target not in images_by_anchor:
                raise LayoutEditError(
                    'ANCHOR_STALE',
                    f'Target image anchor {target!r} not found or stale in document',
                    [f"Available images: {[i['anchor'] for i in mapping.get('images', [])]}"],
                )
            i_info = images_by_anchor[target]
            p_node = body_children[i_info['body_index']]
            drawings = p_node.xpath('.//w:drawing', namespaces={'w': W})
            if i_info['drawing_index_in_p'] >= len(drawings):
                raise LayoutEditError('IMAGE_STRUCTURE', f'Drawing missing at index for {target}')
            drawing_node = drawings[i_info['drawing_index_in_p']]
            orig_cx = i_info['width_emu']
            orig_cy = i_info['height_emu']
            printable_w_cm = i_info['printable_width_cm']

            size_keys = [k for k in ('factor', 'width_cm', 'width_pct') if op.get(k) is not None]
            if len(size_keys) != 1:
                raise LayoutEditError('INVALID_PLAN', 'image.scale requires exactly one of factor, width_cm, or width_pct')
            if op.get('factor') is not None:
                factor = _validate_float_range(op['factor'], 0.15, 3.0, 'INVALID_FACTOR', 'factor')
                new_cx = max(1, round(orig_cx * factor))
            elif op.get('width_cm') is not None:
                w_cm = _validate_float_range(op['width_cm'], 0.5, 40.0, 'INVALID_WIDTH', 'width_cm')
                new_cx = max(1, round(w_cm * EMU_PER_CM))
            else:
                pct = _validate_float_range(op['width_pct'], 5.0, 100.0, 'INVALID_WIDTH', 'width_pct')
                new_cx = max(1, round(printable_w_cm * (pct / 100.0) * EMU_PER_CM))

            new_cy = max(1, round(orig_cy * new_cx / orig_cx))
            new_w_cm = round(new_cx / EMU_PER_CM, 2)
            if new_w_cm > printable_w_cm + 0.04:
                warnings.append({
                    'code': 'overflow_printable_width',
                    'image': target,
                    'message': f'Scaled image width {new_w_cm} cm exceeds printable width {printable_w_cm} cm',
                })
            align = op.get('align')
            if align is not None and align not in ('left', 'center', 'right'):
                raise LayoutEditError('INVALID_ALIGN', f'Image align must be left, center, or right (got {align!r})')

            allowed_body_indices.add(i_info['body_index'])
            prepared.append({
                'kind': 'image',
                'op': op_name,
                'target': target,
                'p_node': p_node,
                'drawing_node': drawing_node,
                'before_width_cm': i_info['width_cm'],
                'before_height_cm': i_info['height_cm'],
                'before_align': i_info['alignment'],
                'new_cx': new_cx,
                'new_cy': new_cy,
                'new_width_cm': new_w_cm,
                'new_height_cm': round(new_cy / EMU_PER_CM, 2),
                'new_align': align,
            })

        elif op_name == 'paragraph.set_format':
            if target not in paragraphs_by_anchor:
                raise LayoutEditError(
                    'ANCHOR_STALE',
                    f'Target paragraph anchor {target!r} not found or stale in document',
                    [f"Paragraph count: {len(mapping['paragraphs'])}"],
                )
            p_info = paragraphs_by_anchor[target]
            p_node = body_children[p_info['body_index']]
            fmt = _validate_paragraph_fmt(op, require_nonempty=True)
            allowed_body_indices.add(p_info['body_index'])
            prepared.append({
                'kind': 'paragraph',
                'op': op_name,
                'target': target,
                'p_node': p_node,
                'fmt': fmt,
            })

        elif op_name == 'paragraph.merge_runs':
            if target not in paragraphs_by_anchor:
                raise LayoutEditError(
                    'ANCHOR_STALE',
                    f'Target paragraph anchor {target!r} not found or stale in document',
                    [f"Paragraph count: {len(mapping['paragraphs'])}"],
                )
            p_info = paragraphs_by_anchor[target]
            p_node = body_children[p_info['body_index']]
            allowed_body_indices.add(p_info['body_index'])
            prepared.append({
                'kind': 'merge_runs',
                'op': op_name,
                'target': target,
                'p_node': p_node,
                'before_text': docx_package.text_of(p_node),
            })

        elif op_name == 'style.update':
            if styles_root is None:
                raise LayoutEditError('STYLE_XML_MISSING', 'Package does not contain word/styles.xml')
            style_name = op.get('style_name')
            style_node = _find_style_node(styles_root, style_name)
            if style_node is None:
                raise LayoutEditError('STYLE_NOT_FOUND', f'Style {style_name!r} not found in word/styles.xml')
            spec = _validate_paragraph_fmt(op, require_nonempty=False)
            if op.get('font') is not None:
                fn = str(op['font']).strip()
                if not fn:
                    raise LayoutEditError('INVALID_PLAN', 'style.update font must be nonempty')
                spec['font'] = fn
            if op.get('size_pt') is not None:
                spec['size_pt'] = _validate_float_range(op['size_pt'], 6.0, 72.0, 'INVALID_FONT_SIZE', 'size_pt')
            if op.get('color') is not None:
                col = str(op['color']).strip()
                if not HEX_COLOR_RE.match(col):
                    raise LayoutEditError('INVALID_PLAN', f'style.update color must be 6-hex, got {col!r}')
                spec['color'] = col.upper()
            for bkey in ('bold', 'italic'):
                if op.get(bkey) is not None:
                    if not isinstance(op[bkey], bool):
                        raise LayoutEditError('INVALID_PLAN', f'style.update {bkey} must be boolean')
                    spec[bkey] = op[bkey]
            if not spec:
                raise LayoutEditError('INVALID_PLAN', 'style.update requires at least one style property')
            prepared.append({
                'kind': 'style',
                'op': op_name,
                'style_name': str(style_name),
                'style_node': style_node,
                'style_spec': spec,
            })

        elif op_name == 'section.set_layout':
            sec_idx = op.get('section_index', 1)
            if not isinstance(sec_idx, int) or isinstance(sec_idx, bool):
                raise LayoutEditError('INVALID_PLAN', 'section_index must be an integer')
            match = next((entry for entry in sections_xml if entry[0] == sec_idx), None)
            if match is None:
                raise LayoutEditError(
                    'SECTION_NOT_FOUND',
                    f'Section {sec_idx} not found (document has {len(sections_xml)} section(s))',
                )
            _, sect_node, b_idx, owner_p = match
            sec_spec = {}
            if op.get('orientation') is not None:
                orient = str(op['orientation']).lower()
                if orient not in ('portrait', 'landscape'):
                    raise LayoutEditError('INVALID_PLAN', f'Invalid section orientation: {orient!r}')
                sec_spec['orientation'] = orient
            if op.get('page_size') is not None:
                ps = str(op['page_size'])
                if ps not in ('A4', 'Letter'):
                    raise LayoutEditError('INVALID_PLAN', f'page_size must be A4 or Letter, got {ps!r}')
                sec_spec['page_size'] = ps
            if op.get('margins_cm') is not None:
                raw_m = op['margins_cm']
                if not isinstance(raw_m, dict) or not raw_m:
                    raise LayoutEditError('INVALID_PLAN', 'margins_cm must be a nonempty object')
                valid_keys = ('top', 'bottom', 'left', 'right', 'header', 'footer')
                bad = set(raw_m) - set(valid_keys)
                if bad:
                    raise LayoutEditError('INVALID_PLAN', f'Unknown margin keys: {sorted(bad)}')
                sec_spec['margins_cm'] = {
                    k: _validate_float_range(v, 0.2, 10.0, 'INVALID_MARGIN', f'margins_cm.{k}')
                    for k, v in raw_m.items()
                }
            if not sec_spec:
                raise LayoutEditError('INVALID_PLAN', 'section.set_layout requires orientation, page_size, or margins_cm')

            if profile in document_rules.PROFILES:
                rules = document_rules.load_rules(profile)
                sec_rules = document_rules.section_rules(rules, sec_idx - 1)
                if sec_spec.get('orientation') == 'landscape' and not sec_rules.get('allow_landscape', True):
                    warnings.append({
                        'code': 'profile_landscape_warning',
                        'section': sec_idx,
                        'message': f'Profile {profile} disallows landscape orientation in section {sec_idx}',
                    })
                bound_map = sec_rules.get('margins_cm', {})
                for side, val in sec_spec.get('margins_cm', {}).items():
                    if side in bound_map:
                        lo, hi = bound_map[side]
                        if val < lo - 0.01 or val > hi + 0.01:
                            warnings.append({
                                'code': 'profile_margin_warning',
                                'section': sec_idx,
                                'side': side,
                                'value_cm': val,
                                'allowed_cm': [lo, hi],
                                'message': f'Section {sec_idx} margin {side}={val} cm outside {profile} range [{lo}, {hi}] cm',
                            })

            allowed_body_indices.add(b_idx)
            prepared.append({
                'kind': 'section',
                'op': op_name,
                'section_index': sec_idx,
                'sect_node': sect_node,
                'p_node': owner_p,
                'section_spec': sec_spec,
            })

    return prepared, warnings, allowed_body_indices


def _summarize_prepared_op(item):
    if item['kind'] == 'table':
        return {
            'op': item['op'],
            'target': item['target'],
            'before_width_cm': item['before_width_cm'],
            'after_width_cm': round(sum(item['new_cols']) / DXA_PER_CM, 2)
            if item['new_cols'] is not None else item['before_width_cm'],
            'after_cols_cm': [round(w / DXA_PER_CM, 2) for w in item['new_cols']]
            if item['new_cols'] is not None else item['before_cols_cm'],
            'after_align': item['new_align'] or item['before_align'],
        }
    if item['kind'] == 'image':
        return {
            'op': item['op'],
            'target': item['target'],
            'before_width_cm': item['before_width_cm'],
            'before_height_cm': item['before_height_cm'],
            'after_width_cm': item['new_width_cm'],
            'after_height_cm': item['new_height_cm'],
            'after_align': item['new_align'] or item['before_align'],
        }
    if item['kind'] == 'merge_runs':
        return {
            'op': item['op'],
            'target': item['target'],
            'merged_runs': item.get('merged_runs', 0),
        }
    if item['kind'] == 'style':
        return {
            'op': item['op'],
            'style_name': item['style_name'],
            'applied': item['style_spec'],
        }
    if item['kind'] == 'section':
        return {
            'op': item['op'],
            'section_index': item['section_index'],
            'applied': item['section_spec'],
        }
    return {
        'op': item['op'],
        'target': item['target'],
        'applied': item['fmt'],
    }


def apply_layout_plan(
    source,
    destination,
    plan,
    *,
    expected_sha256=None,
    dry_run=False,
    allow_layout_tables=False,
    profile=None,
    track_changes=False,
    author='vndocs',
    date_iso=None,
):
    """Apply validated, anchor-scoped layout operations with non-target preservation checks."""
    src_path = Path(source).resolve()
    dst_path = Path(destination).resolve()
    if not src_path.is_file():
        raise LayoutEditError('SOURCE_NOT_FOUND', f'Source DOCX not found: {source}')
    if not dry_run and src_path == dst_path:
        raise LayoutEditError('SAME_PATH', 'Destination must differ from source to preserve original document')

    raw_bytes = src_path.read_bytes()
    actual_sha256 = hashlib.sha256(raw_bytes).hexdigest()
    if expected_sha256 and actual_sha256 != expected_sha256.lower():
        raise LayoutEditError(
            'SOURCE_SHA_MISMATCH',
            f'Source checksum {actual_sha256} differs from expected {expected_sha256}',
            ['Re-run `vndocs map` on the current document.'],
        )

    ops = plan.get('operations') if isinstance(plan, dict) else plan
    if not isinstance(ops, list) or not ops:
        raise LayoutEditError('INVALID_PLAN', 'Plan must contain a nonempty list of operations')

    mapping = map_docx.map_document(src_path)
    entries, original = docx_package.read_package(src_path)
    clean_entries = [entry for entry in entries if not entry.is_dir()]
    parser = etree.XMLParser(resolve_entities=False, no_network=True)
    doc_root = etree.fromstring(original['word/document.xml'], parser)
    styles_root = (
        etree.fromstring(original['word/styles.xml'], parser)
        if 'word/styles.xml' in original
        else None
    )
    before_root = deepcopy(doc_root)
    body = doc_root.find(f'{{{W}}}body')
    tbl_elements = [child for child in body if etree.QName(child).localname == 'tbl']

    prepared, warnings, allowed_body_indices = _prepare_operations(
        ops, mapping, body, tbl_elements, allow_layout_tables, profile, styles_root=styles_root,
    )

    if dry_run:
        return {
            'status': 'dry_run',
            'source': str(src_path),
            'source_sha256': actual_sha256,
            'operations': [_summarize_prepared_op(item) for item in prepared],
            'warnings': warnings,
        }

    track_ctx = None
    if track_changes:
        track_ctx = {
            'rev': _next_rev_id_counter(doc_root),
            'author': author or 'vndocs',
            'date': date_iso or datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ'),
        }

    # Mutate ONLY targeted blocks and perform deterministic read-back per operation.
    for item in prepared:
        kind = item['kind']
        if kind == 'table':
            tbl = item['tbl_node']
            tbl_pr_touches = (
                item['new_cols'] is not None
                or item['new_align'] is not None
                or bool(item.get('sides_dxa'))
                or item.get('border_spec') is not None
            )
            tbl_pr_snap = (
                _clean_change_snapshot(_ensure_pr(tbl, 'tblPr'), ('tblPrChange',))
                if (track_ctx and tbl_pr_touches)
                else None
            )
            tbl_grid_el = tbl.find(f'{{{W}}}tblGrid')
            tbl_grid_snap = (
                _clean_change_snapshot(tbl_grid_el, ('tblGridChange',))
                if (track_ctx and item['new_cols'] is not None and tbl_grid_el is not None)
                else None
            )

            if item['new_cols'] is not None:
                _apply_fixed_grid_and_cells(tbl, item['new_cols'])
                if track_ctx and tbl_grid_snap is not None:
                    cur_grid = tbl.find(f'{{{W}}}tblGrid')
                    if cur_grid is not None:
                        _record_container_change(
                            cur_grid, tbl_grid_snap, 'tblGridChange',
                            track_ctx['rev'], track_ctx['author'], track_ctx['date'],
                        )
            if item['new_align'] is not None:
                _set_alignment(tbl, item['new_align'])
            if item.get('sides_dxa'):
                _apply_cell_margins(tbl, item['sides_dxa'], item.get('clear_overrides', True))
            if item.get('border_spec') is not None:
                _apply_table_borders(tbl, item['border_spec'])
            if track_ctx and tbl_pr_snap is not None:
                _record_container_change(
                    tbl.find(f'{{{W}}}tblPr'), tbl_pr_snap, 'tblPrChange',
                    track_ctx['rev'], track_ctx['author'], track_ctx['date'],
                )

            if item.get('size_hp') is not None:
                item['modified_rprs'] = _apply_table_font_size(
                    tbl, item['size_hp'], item.get('preserve_header_size', False),
                )
            if item.get('pagination') is not None:
                pag = item['pagination']
                _apply_table_pagination(tbl, pag['repeat_header'], pag['header_rows'], pag['cant_split'])
            if item.get('shading_spec') is not None:
                item['modified_shd_tcprs'] = _apply_table_shading(tbl, item['shading_spec'], track_ctx=track_ctx)
            _verify_table_readback(item)
        elif kind == 'image':
            _apply_image_scale(
                item['drawing_node'], item['p_node'], item['new_cx'], item['new_cy'], item.get('new_align'),
            )
            _verify_image_readback(item)
        elif kind == 'paragraph':
            p_pr_snap = (
                _clean_change_snapshot(_ensure_pr(item['p_node'], 'pPr'), ('pPrChange',))
                if track_ctx
                else None
            )
            _apply_paragraph_format(item['p_node'], item['fmt'])
            if track_ctx and p_pr_snap is not None:
                _record_container_change(
                    item['p_node'].find(f'{{{W}}}pPr'), p_pr_snap, 'pPrChange',
                    track_ctx['rev'], track_ctx['author'], track_ctx['date'],
                )
            _verify_paragraph_readback(item)
        elif kind == 'merge_runs':
            item['merged_runs'] = _apply_merge_runs(item['p_node'])
            _verify_merge_readback(item)
        elif kind == 'style':
            _apply_style_update(item['style_node'], item['style_spec'])
            _verify_style_readback(item)
        elif kind == 'section':
            _apply_section_layout(item['sect_node'], item['section_spec'])
            _verify_section_readback(item)

    order_violations = _check_mutated_order(prepared)
    if order_violations:
        raise LayoutEditError('ELEMENT_ORDER_VIOLATION', order_violations[0]['message'])

    ntpr_report = ooxml_order.check_non_target_preservation(before_root, doc_root, allowed_body_indices)
    if ntpr_report['xml_ntpr'] != 1.0:
        raise LayoutEditError(
            'COLLATERAL_MUTATION',
            f'Non-target body elements mutated: {ntpr_report}',
        )

    before_profile = validate_docx.inspect_doc(src_path, profile=profile) if profile else None

    changed = {}
    if any(item['kind'] != 'style' for item in prepared):
        changed['word/document.xml'] = docx_package.xml_bytes(doc_root)
    if styles_root is not None and any(item['kind'] == 'style' for item in prepared):
        changed['word/styles.xml'] = docx_package.xml_bytes(styles_root)

    pkg_report = docx_package.write_package(src_path, dst_path, clean_entries, original, changed)
    out_sha256 = hashlib.sha256(dst_path.read_bytes()).hexdigest()

    profile_regression = None
    if profile and before_profile is not None:
        after_profile = validate_docx.inspect_doc(dst_path, profile=profile)
        before_set = {(i['code'], i['location'], i['message']) for i in before_profile['issues']}
        new_issues = [
            i for i in after_profile['issues']
            if (i['code'], i['location'], i['message']) not in before_set
        ]
        profile_regression = {
            'profile': profile,
            'before_status': before_profile['status'],
            'after_status': after_profile['status'],
            'new_issues': new_issues,
        }

    return {
        **pkg_report,
        'status': 'ok',
        'output_sha256': out_sha256,
        'track_changes': bool(track_changes),
        'operations': [_summarize_prepared_op(item) for item in prepared],
        'warnings': warnings,
        'non_target_preservation': ntpr_report,
        'profile_check': profile_regression,
    }


def _load_plan(spec):
    candidate = Path(spec)
    raw = candidate.read_text(encoding='utf-8-sig') if candidate.is_file() else spec
    return json.loads(raw)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source', help='Source DOCX path (preserved unmodified)')
    parser.add_argument('destination', help='Destination DOCX path')
    parser.add_argument('--plan', required=True, help='JSON path or inline JSON string of operations')
    parser.add_argument('--source-sha256', help='Expected SHA-256 of source DOCX')
    parser.add_argument('--dry-run', action='store_true', help='Validate & preview without writing destination')
    parser.add_argument('--allow-layout-tables', action='store_true', help='Permit editing protected layout tables')
    parser.add_argument('--profile', help='Optional Vietnamese profile to run before/after regression check')
    parser.add_argument('--track-changes', action='store_true', help='Record OOXML revision change elements (*PrChange / tblGridChange)')
    parser.add_argument('--author', default='vndocs', help='Revision author name when --track-changes is enabled')
    args = parser.parse_args(argv)
    try:
        plan = _load_plan(args.plan)
        report = apply_layout_plan(
            args.source,
            args.destination,
            plan,
            expected_sha256=args.source_sha256,
            dry_run=args.dry_run,
            allow_layout_tables=args.allow_layout_tables,
            profile=args.profile,
            track_changes=args.track_changes,
            author=args.author,
        )
    except LayoutEditError as exc:
        print(json.dumps(exc.to_dict(), ensure_ascii=False, indent=2))
        return 1
    except (ValueError, OSError, json.JSONDecodeError) as exc:
        print(json.dumps({'status': 'error', 'code': 'BAD_INPUT', 'error': str(exc)}, ensure_ascii=False, indent=2))
        return 2

    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
