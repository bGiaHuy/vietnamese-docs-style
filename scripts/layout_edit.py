"""Safe, anchor-locked OOXML layout edits with pre-flight, read-back, and non-target checks."""
import argparse
import hashlib
import json
import re
from copy import deepcopy
from pathlib import Path

from lxml import etree

try:
    from . import docx_package, map_docx, ooxml_order, validate_docx
except ImportError:
    import docx_package
    import map_docx
    import ooxml_order
    import validate_docx

W = docx_package.W
WP = map_docx.WP
A = map_docx.A
ALLOWED_OPS = frozenset({
    'table.scale_width',
    'table.set_col_widths',
    'table.set_alignment',
    'table.set_cell_margins',
    'table.set_font_size',
    'image.scale',
    'paragraph.set_format',
})
MIN_TEXT_COL_DXA = 288  # ~0.508 cm (0.2 in); narrower text-bearing cells wrap per character.
DXA_PER_CM = map_docx.DXA_PER_CM
EMU_PER_CM = map_docx.EMU_PER_CM
EMU_PER_DXA = EMU_PER_CM / DXA_PER_CM
LENGTH_RE = re.compile(r'^([0-9]+(?:\.[0-9]+)?)\s*(cm|mm|in|pt|dxa|tw)?$', re.IGNORECASE)
MARGIN_SIDES = ('top', 'left', 'bottom', 'right')
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


def _check_mutated_order(prepared):
    violations = []
    for item in prepared:
        if item['kind'] == 'table':
            tbl_node = item['tbl_node']
            for idx, ch in enumerate(tbl_node):
                if isinstance(ch.tag, str) and etree.QName(ch).localname == 'tblPr' and idx > 0:
                    violations.append({'message': f'<w:tblPr> must be child 0 of <w:tbl> (found at {idx})'})
            tbl_pr = tbl_node.find(f'{{{W}}}tblPr')
            if tbl_pr is not None:
                violations.extend(ooxml_order.check_element_order(tbl_pr))
            if item.get('new_cols') is not None or item.get('clear_overrides'):
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
    """Expand multi-target paragraph.set_format ops into uniform per-target items."""
    expanded = []
    for idx, op in enumerate(ops):
        if not isinstance(op, dict) or op.get('op') not in ALLOWED_OPS:
            raise LayoutEditError(
                'UNSUPPORTED_OP',
                f'Operation {idx} is not in allowlist {sorted(ALLOWED_OPS)}: {op.get("op") if isinstance(op, dict) else op!r}',
            )
        if op['op'] == 'paragraph.set_format' and op.get('targets') is not None:
            targets = op['targets']
            if not isinstance(targets, list) or not targets:
                raise LayoutEditError('INVALID_PLAN', 'paragraph.set_format targets must be a nonempty list')
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


def _prepare_operations(ops, mapping, body, tbl_elements, allow_layout_tables, profile):
    tables_by_anchor = {t['anchor']: t for t in mapping['tables']}
    images_by_anchor = {img['anchor']: img for img in mapping.get('images', [])}
    paragraphs_by_anchor = {p['anchor']: p for p in mapping['paragraphs']}
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
            if not fmt:
                raise LayoutEditError('INVALID_PLAN', 'paragraph.set_format requires at least one formatting property')

            allowed_body_indices.add(p_info['body_index'])
            prepared.append({
                'kind': 'paragraph',
                'op': op_name,
                'target': target,
                'p_node': p_node,
                'fmt': fmt,
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
    before_root = deepcopy(doc_root)
    body = doc_root.find(f'{{{W}}}body')
    tbl_elements = [child for child in body if etree.QName(child).localname == 'tbl']

    prepared, warnings, allowed_body_indices = _prepare_operations(
        ops, mapping, body, tbl_elements, allow_layout_tables, profile,
    )

    if dry_run:
        return {
            'status': 'dry_run',
            'source': str(src_path),
            'source_sha256': actual_sha256,
            'operations': [_summarize_prepared_op(item) for item in prepared],
            'warnings': warnings,
        }

    # Mutate ONLY targeted blocks and perform deterministic read-back per operation.
    for item in prepared:
        if item['kind'] == 'table':
            if item['new_cols'] is not None:
                _apply_fixed_grid_and_cells(item['tbl_node'], item['new_cols'])
            if item['new_align'] is not None:
                _set_alignment(item['tbl_node'], item['new_align'])
            if item.get('sides_dxa'):
                _apply_cell_margins(item['tbl_node'], item['sides_dxa'], item.get('clear_overrides', True))
            if item.get('size_hp') is not None:
                item['modified_rprs'] = _apply_table_font_size(
                    item['tbl_node'], item['size_hp'], item.get('preserve_header_size', False),
                )
            _verify_table_readback(item)
        elif item['kind'] == 'image':
            _apply_image_scale(
                item['drawing_node'], item['p_node'], item['new_cx'], item['new_cy'], item.get('new_align'),
            )
            _verify_image_readback(item)
        elif item['kind'] == 'paragraph':
            _apply_paragraph_format(item['p_node'], item['fmt'])
            _verify_paragraph_readback(item)

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

    changed = {'word/document.xml': docx_package.xml_bytes(doc_root)}
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
