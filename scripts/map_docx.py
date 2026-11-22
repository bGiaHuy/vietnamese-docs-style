"""Build compact structural, section-geometry, and optional PDF maps for DOCX files."""
import argparse
import hashlib
import json
import re
import unicodedata
from pathlib import Path

from lxml import etree

try:
    from . import docx_package, render_docx
except ImportError:
    import docx_package
    import render_docx

W = docx_package.W
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS = {'w': W, 'wp': WP, 'a': A, 'r': R}
DXA_PER_CM = 1440.0 / 2.54
PT_PER_CM = 72.0 / 2.54
EMU_PER_CM = 360000.0
BORDER_SIDES = ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')

LAYOUT_KEYWORDS = (
    'cộng hòa xã hội chủ nghĩa việt nam',
    'độc lập - tự do - hạnh phúc',
    'đảng cộng sản việt nam',
    'nơi nhận',
    'giảng viên ra đề',
    'trưởng bộ môn',
    '{{signature_note',
    '{{authority}}',
    '{{position}}',
)


def normalize_nfc(text):
    """Normalize Vietnamese Unicode to NFC and collapse internal whitespace."""
    return re.sub(r'\s+', ' ', unicodedata.normalize('NFC', text or '')).strip()


def hash4(text):
    return hashlib.sha256(text.encode('utf-8')).hexdigest()[:4]


def _int_attr(element, attr, default=None):
    if element is None:
        return default
    raw = element.get(f'{{{W}}}{attr}')
    if raw is None:
        return default
    try:
        return int(raw)
    except ValueError:
        return default


def _section_geometry(sect_pr, index):
    pg_sz = sect_pr.find(f'{{{W}}}pgSz') if sect_pr is not None else None
    pg_mar = sect_pr.find(f'{{{W}}}pgMar') if sect_pr is not None else None

    page_w = _int_attr(pg_sz, 'w', 11906)
    page_h = _int_attr(pg_sz, 'h', 16838)
    orient = pg_sz.get(f'{{{W}}}orient') if pg_sz is not None else None
    if orient is None:
        orient = 'landscape' if page_w > page_h else 'portrait'

    left = _int_attr(pg_mar, 'left', 1440)
    right = _int_attr(pg_mar, 'right', 1440)
    top = _int_attr(pg_mar, 'top', 1440)
    bottom = _int_attr(pg_mar, 'bottom', 1440)
    gutter = _int_attr(pg_mar, 'gutter', 0)
    printable_w = max(1, page_w - left - right - gutter)

    return {
        'section_index': index,
        'page_width_dxa': page_w,
        'page_height_dxa': page_h,
        'page_width_cm': round(page_w / DXA_PER_CM, 2),
        'page_height_cm': round(page_h / DXA_PER_CM, 2),
        'orientation': orient,
        'margins_dxa': {'left': left, 'right': right, 'top': top, 'bottom': bottom, 'gutter': gutter},
        'margins_cm': {
            'left': round(left / DXA_PER_CM, 2),
            'right': round(right / DXA_PER_CM, 2),
            'top': round(top / DXA_PER_CM, 2),
            'bottom': round(bottom / DXA_PER_CM, 2),
        },
        'printable_width_dxa': printable_w,
        'printable_width_cm': round(printable_w / DXA_PER_CM, 2),
    }


def _resolve_sections(body):
    children = list(body)
    sect_nodes = []
    block_to_section = {}
    pending = []
    sec_idx = 0

    for idx, child in enumerate(children):
        qname = etree.QName(child).localname
        if qname == 'sectPr':
            sec_idx += 1
            sect_nodes.append((sec_idx, child))
            for b_idx in pending:
                block_to_section[b_idx] = sec_idx
            pending = []
            continue
        pending.append(idx)
        if qname == 'p':
            p_pr = child.find(f'{{{W}}}pPr')
            p_sect = p_pr.find(f'{{{W}}}sectPr') if p_pr is not None else None
            if p_sect is not None:
                sec_idx += 1
                sect_nodes.append((sec_idx, p_sect))
                for b_idx in pending:
                    block_to_section[b_idx] = sec_idx
                pending = []

    if not sect_nodes:
        sect_nodes.append((1, None))
    last_sec = sect_nodes[-1][0]
    for b_idx in pending:
        block_to_section[b_idx] = last_sec

    sections = [_section_geometry(node, s_idx) for s_idx, node in sect_nodes]
    return sections, block_to_section


def _cell_text(tc):
    parts = [re.sub(r'\s+', ' ', docx_package.text_of(p)).strip() for p in tc.findall(f'{{{W}}}p')]
    return normalize_nfc(' '.join(part for part in parts if part))


def _row_logical_grid_count(tr):
    tr_pr = tr.find(f'{{{W}}}trPr')
    count = _int_attr(tr_pr.find(f'{{{W}}}gridBefore') if tr_pr is not None else None, 'val', 0)
    count += _int_attr(tr_pr.find(f'{{{W}}}gridAfter') if tr_pr is not None else None, 'val', 0)
    for tc in tr.findall(f'{{{W}}}tc'):
        tc_pr = tc.find(f'{{{W}}}tcPr')
        span = _int_attr(tc_pr.find(f'{{{W}}}gridSpan') if tc_pr is not None else None, 'val', 1)
        count += max(1, span)
    return count


def _resolve_table_widths(tbl, printable_w_dxa):
    rows = tbl.findall(f'{{{W}}}tr')
    tbl_grid = tbl.find(f'{{{W}}}tblGrid')
    grid_cols = tbl_grid.findall(f'{{{W}}}gridCol') if tbl_grid is not None else []
    grid_widths = [_int_attr(col, 'w') for col in grid_cols]

    tbl_pr = tbl.find(f'{{{W}}}tblPr')
    tbl_w = tbl_pr.find(f'{{{W}}}tblW') if tbl_pr is not None else None
    tbl_w_val = _int_attr(tbl_w, 'w')
    tbl_w_type = tbl_w.get(f'{{{W}}}type') if tbl_w is not None else None

    if grid_widths and all(w is not None and w > 0 for w in grid_widths):
        return grid_widths, sum(grid_widths), 'tblGrid'

    n_cols = len(grid_cols) or max((_row_logical_grid_count(tr) for tr in rows), default=1)
    if tbl_w_type == 'dxa' and tbl_w_val and tbl_w_val > 0:
        total = tbl_w_val
        source = 'tblW_dxa'
    elif tbl_w_type == 'pct' and tbl_w_val and tbl_w_val > 0:
        total = max(1, round(printable_w_dxa * (tbl_w_val / 5000.0)))
        source = 'tblW_pct'
    else:
        total = printable_w_dxa
        source = 'printable_fallback'

    base = total // n_cols
    resolved = [base] * n_cols
    resolved[-1] = total - base * (n_cols - 1)
    return resolved, total, source


def _detect_layout_table(tbl, rows, grid_count, preview_rows, style_val, borders_dict):
    all_text = normalize_nfc(' '.join(cell for row in preview_rows for cell in row)).lower()
    explicit_nil = bool(borders_dict) and all(
        borders_dict.get(side) in ('nil', 'none')
        for side in ('top', 'left', 'bottom', 'right')
    )
    keyword_hit = any(kw in all_text for kw in LAYOUT_KEYWORDS)

    if explicit_nil and grid_count <= 2 and len(rows) <= 3:
        if any(k in all_text for k in ('cộng hòa xã hội chủ nghĩa việt nam', 'độc lập - tự do - hạnh phúc')):
            return True, 'header_quoc_hieu'
        if 'đảng cộng sản việt nam' in all_text:
            return True, 'header_party'
        if any(k in all_text for k in ('học phần:', '{{exam_title}}', 'giảng viên ra đề')):
            if 'trưởng bộ môn' in all_text or 'giảng viên ra đề' in all_text:
                return True, 'signature_block'
            return True, 'header_exam'
        if any(k in all_text for k in ('nơi nhận', '{{signature_note', '{{authority}}', '{{position}}', 'kt.', 'tm.')):
            return True, 'signature_block'
        return True, 'borderless_layout'

    if explicit_nil and not style_val and keyword_hit:
        return True, 'borderless_layout'

    return False, None


def _table_entry(tbl, tbl_ordinal, body_index, sec_geom, heading_context):
    rows = tbl.findall(f'{{{W}}}tr')
    tbl_pr = tbl.find(f'{{{W}}}tblPr')
    style_el = tbl_pr.find(f'{{{W}}}tblStyle') if tbl_pr is not None else None
    style_val = style_el.get(f'{{{W}}}val') if style_el is not None else None
    layout_el = tbl_pr.find(f'{{{W}}}tblLayout') if tbl_pr is not None else None
    layout_val = layout_el.get(f'{{{W}}}type') if layout_el is not None else 'autofit'
    jc_el = tbl_pr.find(f'{{{W}}}jc') if tbl_pr is not None else None
    align_val = jc_el.get(f'{{{W}}}val') if jc_el is not None else 'left'
    tbl_w_el = tbl_pr.find(f'{{{W}}}tblW') if tbl_pr is not None else None
    tbl_w_raw = (
        {'w': tbl_w_el.get(f'{{{W}}}w'), 'type': tbl_w_el.get(f'{{{W}}}type')}
        if tbl_w_el is not None else None
    )

    borders_el = tbl_pr.find(f'{{{W}}}tblBorders') if tbl_pr is not None else None
    borders = {}
    if borders_el is not None:
        for side in BORDER_SIDES:
            child = borders_el.find(f'{{{W}}}{side}')
            if child is not None:
                borders[side] = child.get(f'{{{W}}}val')

    has_merged = any(
        tc_pr.find(f'{{{W}}}gridSpan') is not None or tc_pr.find(f'{{{W}}}vMerge') is not None
        for tc_pr in tbl.xpath('.//w:tr/w:tc/w:tcPr', namespaces=NS)
    )
    has_nested = bool(tbl.xpath('.//w:tr/w:tc//w:tbl', namespaces=NS))

    all_rows_text = []
    for r_idx, tr in enumerate(rows):
        cells_txt = [_cell_text(tc) for tc in tr.findall(f'{{{W}}}tc')]
        all_rows_text.append(cells_txt)

    printable_w = sec_geom['printable_width_dxa']
    grid_widths, total_w, width_source = _resolve_table_widths(tbl, printable_w)
    grid_count = len(grid_widths)

    sig_rows = ['|'.join(row) for row in all_rows_text[:2]]
    sig = f'{len(rows)}x{grid_count}|' + '||'.join(sig_rows)
    anchor = f'T{tbl_ordinal}#{hash4(sig)}'

    is_layout, layout_role = _detect_layout_table(
        tbl, rows, grid_count, all_rows_text[:3], style_val, borders
    )

    return {
        'anchor': anchor,
        'ordinal': tbl_ordinal,
        'body_index': body_index,
        'section_index': sec_geom['section_index'],
        'printable_width_dxa': printable_w,
        'printable_width_cm': sec_geom['printable_width_cm'],
        'rows': len(rows),
        'grid_cols': grid_count,
        'grid_widths_dxa': grid_widths,
        'grid_widths_cm': [round(w / DXA_PER_CM, 2) for w in grid_widths],
        'grid_widths_pct': [round(w * 100.0 / total_w, 1) if total_w else 0.0 for w in grid_widths],
        'total_width_dxa': total_w,
        'total_width_cm': round(total_w / DXA_PER_CM, 2),
        'width_pct_of_printable': round(total_w * 100.0 / printable_w, 1),
        'width_source': width_source,
        'tblW_raw': tbl_w_raw,
        'layout': layout_val,
        'alignment': align_val,
        'style': style_val,
        'borders': borders,
        'has_merged_cells': has_merged,
        'has_nested_tables': has_nested,
        'is_layout_table': is_layout,
        'layout_role': layout_role,
        'heading_context': heading_context,
        'preview_rows': all_rows_text[:3],
        'row_texts_nfc': all_rows_text,
    }


def _parse_doc_rels(parts):
    """Map rId -> Target media path from word/_rels/document.xml.rels when present."""
    raw = parts.get('word/_rels/document.xml.rels')
    if not raw:
        return {}
    parser = etree.XMLParser(resolve_entities=False, no_network=True)
    rels_root = etree.fromstring(raw, parser)
    mapping = {}
    for rel in rels_root.findall(f'{{{REL_NS}}}Relationship'):
        rid = rel.get('Id')
        target = rel.get('Target')
        if rid and target:
            mapping[rid] = target
    return mapping


def _extract_p_images(p_el, body_index, start_ordinal, sec_geom, heading_context, rels_map):
    """Extract inline/anchored DrawingML picture entries from a top-level paragraph."""
    drawings = p_el.xpath('.//w:drawing', namespaces=NS)
    if not drawings:
        return []

    p_pr = p_el.find(f'{{{W}}}pPr')
    jc_el = p_pr.find(f'{{{W}}}jc') if p_pr is not None else None
    align_val = jc_el.get(f'{{{W}}}val') if jc_el is not None else 'left'
    printable_dxa = sec_geom['printable_width_dxa']
    printable_cm = sec_geom['printable_width_cm']

    found = []
    for d_idx, drawing in enumerate(drawings):
        box = drawing.find(f'{{{WP}}}inline')
        if box is None:
            box = drawing.find(f'{{{WP}}}anchor')
        if box is None:
            continue
        extent = box.find(f'{{{WP}}}extent')
        if extent is None:
            continue
        try:
            cx = int(extent.get('cx') or 0)
            cy = int(extent.get('cy') or 0)
        except ValueError:
            continue
        if cx <= 0 or cy <= 0:
            continue

        doc_pr = box.find(f'{{{WP}}}docPr')
        name = normalize_nfc(doc_pr.get('name', '')) if doc_pr is not None else ''
        descr = normalize_nfc(doc_pr.get('descr', '')) if doc_pr is not None else ''
        blip_list = box.xpath('.//a:blip', namespaces=NS)
        rid = blip_list[0].get(f'{{{R}}}embed') if blip_list else None
        media_target = rels_map.get(rid) if rid else None

        ordinal = start_ordinal + len(found) + 1
        sig = f"{media_target or name}|{cx}x{cy}|{descr}"
        anchor = f'IMG{ordinal}#{hash4(sig)}'
        w_cm = round(cx / EMU_PER_CM, 2)
        h_cm = round(cy / EMU_PER_CM, 2)
        found.append({
            'anchor': anchor,
            'ordinal': ordinal,
            'body_index': body_index,
            'drawing_index_in_p': d_idx,
            'section_index': sec_geom['section_index'],
            'printable_width_dxa': printable_dxa,
            'printable_width_cm': printable_cm,
            'rId': rid,
            'media_target': media_target,
            'name': name,
            'descr': descr,
            'width_emu': cx,
            'height_emu': cy,
            'width_cm': w_cm,
            'height_cm': h_cm,
            'aspect_ratio': round(cx / cy, 4) if cy else None,
            'width_pct_of_printable': round((w_cm * 100.0 / printable_cm), 1) if printable_cm else 0.0,
            'alignment': align_val,
            'heading_context': heading_context,
        })
    return found


def _collect_page_bboxes(matches):
    """Build {str(page_num): [x0, y0, x1, y1]} combining boxes on the same page."""
    per_page = {}
    for _, page_idx, bbox in matches:
        key = str(page_idx)
        if key not in per_page:
            per_page[key] = list(bbox)
        else:
            cur = per_page[key]
            per_page[key] = [
                min(cur[0], bbox[0]),
                min(cur[1], bbox[1]),
                max(cur[2], bbox[2]),
                max(cur[3], bbox[3]),
            ]
    return {k: [round(v, 2) for v in box] for k, box in per_page.items()}


def _match_table_in_pdf(pdf_doc, table_entry, sections_by_idx):
    """Locate a table in a rendered PDF via border-table detection or cell text spans."""
    sec = sections_by_idx[table_entry['section_index']]
    printable_right_cm = sec['page_width_cm'] - sec['margins_cm']['right']

    tokens = []
    for row in table_entry.get('row_texts_nfc', []):
        for cell in row:
            words = [w for w in cell.split() if len(w) >= 2]
            tokens.extend(words)
    if not tokens:
        for row in table_entry.get('row_texts_nfc', []):
            for cell in row:
                tokens.extend(cell.split())
    target_token_set = {t.lower() for t in tokens}

    best_border_matches = []
    page_span_boxes = []

    for page_idx, page in enumerate(pdf_doc, start=1):
        # 1) Try PyMuPDF line-based border detection first (exact outer box on bordered tables).
        try:
            found_tables = page.find_tables()
            tables_list = list(found_tables.tables) if found_tables else []
        except Exception:
            tables_list = []

        for tbl_obj in tables_list:
            extracted = tbl_obj.extract() or []
            cell_strs = [normalize_nfc(str(c or '')) for row in extracted for c in row if c]
            pdf_words = {w.lower() for s in cell_strs for w in s.split() if len(w) >= 2}
            if not pdf_words or not target_token_set:
                continue
            overlap = len(target_token_set & pdf_words)
            recall = overlap / len(target_token_set)
            precision = overlap / len(pdf_words)
            score = (2 * precision * recall / (precision + recall)) if (precision + recall) else 0.0
            if overlap >= min(2, len(target_token_set)) and (recall >= 0.35 or score >= 0.35):
                best_border_matches.append((score, page_idx, tbl_obj.bbox))

        # 2) Text-span fallback (crucial for borderless NĐ30 tables or multi-page span edges).
        if not target_token_set:
            continue
        word_entries = page.get_text('words')
        matched_rects = []
        # Match contiguous phrases from non-empty cells first for high precision.
        non_empty_cells = [
            re.sub(r'\s+', ' ', c).strip().lower()
            for row in table_entry.get('row_texts_nfc', [])
            for c in row
            if len(re.sub(r'\s+', ' ', c).strip()) >= 3
        ]
        page_text_lower = re.sub(r'\s+', ' ', normalize_nfc(page.get_text('text'))).lower()
        hits = sum(1 for c in non_empty_cells if c[:30] in page_text_lower)
        if non_empty_cells and hits >= max(1, min(3, len(non_empty_cells) // 2 + 1)):
            for w_item in word_entries:
                w_norm = normalize_nfc(w_item[4]).lower()
                if w_norm and any(w_norm in c for c in non_empty_cells):
                    matched_rects.append(w_item[:4])
            if matched_rects:
                bbox = (
                    min(r[0] for r in matched_rects),
                    min(r[1] for r in matched_rects),
                    max(r[2] for r in matched_rects),
                    max(r[3] for r in matched_rects),
                )
                page_span_boxes.append((hits / len(non_empty_cells), page_idx, bbox))

    if best_border_matches:
        # Handle multi-page bordered tables by collecting segments with solid overlap on consecutive pages.
        best_score = max(m[0] for m in best_border_matches)
        selected = [m for m in best_border_matches if m[0] >= max(0.35, best_score * 0.65)]
        pages = sorted({m[1] for m in selected})
        first_bbox = selected[0][2]
        max_x1 = max(m[2][2] for m in selected)
        min_x0 = min(m[2][0] for m in selected)
        w_cm = round((max_x1 - min_x0) / PT_PER_CM, 2)
        h_cm = round((first_bbox[3] - first_bbox[1]) / PT_PER_CM, 2)
        overflow_mm = round(max(0.0, (max_x1 / PT_PER_CM - printable_right_cm) * 10.0), 2)
        return {
            'matched': True,
            'measurement_source': 'pdf_table_borders',
            'pdf_pages': pages,
            'spans_pages': len(pages) > 1,
            'bbox_pt': [round(v, 2) for v in first_bbox],
            'page_bboxes': _collect_page_bboxes(selected),
            'width_cm': w_cm,
            'height_cm': h_cm,
            'overflow_right_mm': overflow_mm,
        }

    if page_span_boxes:
        best_ratio = max(x[0] for x in page_span_boxes)
        selected = [x for x in page_span_boxes if x[0] >= best_ratio * 0.7]
        pages = sorted({x[1] for x in selected})
        first_bbox = selected[0][2]
        w_cm = round((first_bbox[2] - first_bbox[0]) / PT_PER_CM, 2)
        h_cm = round((first_bbox[3] - first_bbox[1]) / PT_PER_CM, 2)
        overflow_mm = round(max(0.0, (first_bbox[2] / PT_PER_CM - printable_right_cm) * 10.0), 2)
        return {
            'matched': True,
            'measurement_source': 'pdf_text_spans',
            'pdf_pages': pages,
            'spans_pages': len(pages) > 1,
            'bbox_pt': [round(v, 2) for v in first_bbox],
            'page_bboxes': _collect_page_bboxes(selected),
            'width_cm': w_cm,
            'height_cm': h_cm,
            'overflow_right_mm': overflow_mm,
        }

    return {'matched': False, 'measurement_source': None, 'pdf_pages': [], 'spans_pages': False, 'page_bboxes': {}}


def _match_images_in_pdf(pdf_doc, images, sections_by_idx):
    """Locate embedded images across rendered PDF pages using PyMuPDF image bounding boxes."""
    if not images:
        return

    pdf_boxes = []
    for page_idx, page in enumerate(pdf_doc, start=1):
        seen_on_page = []
        for img_info in page.get_images(full=True):
            xref = img_info[0]
            try:
                rects = page.get_image_rects(xref)
            except Exception:
                rects = []
            for rect in rects:
                if rect.is_empty or rect.width <= 2 or rect.height <= 2:
                    continue
                dup = any(
                    abs(rect.x0 - ex[0]) < 2.0
                    and abs(rect.y0 - ex[1]) < 2.0
                    and abs(rect.x1 - ex[2]) < 2.0
                    and abs(rect.y1 - ex[3]) < 2.0
                    for ex in seen_on_page
                )
                if not dup:
                    seen_on_page.append((rect.x0, rect.y0, rect.x1, rect.y1))
        seen_on_page.sort(key=lambda b: (round(b[1], 1), round(b[0], 1)))
        for box in seen_on_page:
            w_cm = (box[2] - box[0]) / PT_PER_CM
            h_cm = (box[3] - box[1]) / PT_PER_CM
            pdf_boxes.append({
                'page': page_idx,
                'bbox': box,
                'w_cm': w_cm,
                'h_cm': h_cm,
                'ar': (w_cm / h_cm) if h_cm else 1.0,
            })

    cursor = 0
    for img in images:
        sec = sections_by_idx[img['section_index']]
        printable_right_cm = sec['page_width_cm'] - sec['margins_cm']['right']
        target_ar = img['aspect_ratio'] or 1.0
        target_w = img['width_cm']
        best_idx = None
        if len(pdf_boxes) == len(images):
            best_idx = cursor
        else:
            best_score = float('inf')
            for idx in range(cursor, len(pdf_boxes)):
                cand = pdf_boxes[idx]
                score = (
                    abs(cand['ar'] - target_ar) / max(target_ar, 0.1)
                    + abs(cand['w_cm'] - target_w) / max(target_w, 1.0) * 0.5
                )
                if score < best_score:
                    best_score = score
                    best_idx = idx
        if best_idx is None or best_idx >= len(pdf_boxes):
            img['pdf'] = {'matched': False, 'pdf_pages': [], 'page_bboxes': {}}
            continue

        chosen = pdf_boxes[best_idx]
        cursor = best_idx + 1
        box = chosen['bbox']
        overflow_mm = round(max(0.0, (box[2] / PT_PER_CM - printable_right_cm) * 10.0), 2)
        img['pdf'] = {
            'matched': True,
            'measurement_source': 'pdf_image_rects',
            'pdf_pages': [chosen['page']],
            'spans_pages': False,
            'bbox_pt': [round(v, 2) for v in box],
            'page_bboxes': {str(chosen['page']): [round(v, 2) for v in box]},
            'width_cm': round(chosen['w_cm'], 2),
            'height_cm': round(chosen['h_cm'], 2),
            'overflow_right_mm': overflow_mm,
        }


def map_document(filepath, *, with_pdf=False, renderer=None, pdf_path=None, include_row_texts=False):
    """Build a compact structural + spatial index of a DOCX file."""
    path = Path(filepath)
    if not path.is_file():
        raise FileNotFoundError(f'DOCX file not found: {filepath}')
    source_sha256 = hashlib.sha256(path.read_bytes()).hexdigest()

    _, original = docx_package.read_package(path)
    roots = dict(docx_package.stories(original))
    doc_root = roots.get('word/document.xml')
    if doc_root is None:
        raise ValueError('Package does not contain word/document.xml')

    body = doc_root.find(f'{{{W}}}body')
    if body is None:
        raise ValueError('word/document.xml does not contain w:body')

    sections, block_to_section = _resolve_sections(body)
    sections_by_idx = {s['section_index']: s for s in sections}
    rels_map = _parse_doc_rels(original)

    paragraphs = []
    tables = []
    images = []
    p_ordinal = 0
    t_ordinal = 0
    last_heading = None

    for b_idx, child in enumerate(body):
        qname = etree.QName(child).localname
        sec_idx = block_to_section.get(b_idx, sections[-1]['section_index'])
        sec_geom = sections_by_idx[sec_idx]

        if qname == 'p':
            p_ordinal += 1
            raw_text = docx_package.text_of(child)
            norm = normalize_nfc(raw_text)
            anchor = f'P{p_ordinal}#{hash4(norm)}'
            p_pr = child.find(f'{{{W}}}pPr')
            p_style = p_pr.find(f'{{{W}}}pStyle') if p_pr is not None else None
            style_name = p_style.get(f'{{{W}}}val') if p_style is not None else None
            if norm and (
                (style_name and style_name.lower().startswith(('heading', 'title')))
                or (len(norm) <= 100 and re.match(r'^([IVX]+\.|[0-9]+\.|Điều\s+[0-9]+|Chương\s+[IVX0-9]+)', norm))
            ):
                last_heading = norm[:80]
            paragraphs.append({
                'anchor': anchor,
                'ordinal': p_ordinal,
                'body_index': b_idx,
                'section_index': sec_idx,
                'style': style_name,
                'text_preview': norm[:90],
                'char_count': len(norm),
            })
            p_imgs = _extract_p_images(child, b_idx, len(images), sec_geom, last_heading, rels_map)
            images.extend(p_imgs)
        elif qname == 'tbl':
            t_ordinal += 1
            tables.append(_table_entry(child, t_ordinal, b_idx, sec_geom, last_heading))

    render_status = 'not_requested'
    resolved_pdf = None
    if pdf_path:
        resolved_pdf = str(Path(pdf_path))
        render_status = 'provided'
    elif with_pdf:
        render_out = render_docx.render_doc(path, renderer=renderer)
        render_status = render_out['status']
        resolved_pdf = render_out.get('pdf')

    if resolved_pdf and Path(resolved_pdf).is_file():
        import pymupdf
        with pymupdf.open(resolved_pdf) as pdf_doc:
            page_count = len(pdf_doc)
            for t in tables:
                t['pdf'] = _match_table_in_pdf(pdf_doc, t, sections_by_idx)
            _match_images_in_pdf(pdf_doc, images, sections_by_idx)
    else:
        page_count = None

    if include_row_texts:
        public_tables = [dict(t) for t in tables]
    else:
        # Strip full row matrix from public output to preserve token budget (~keep compact).
        public_tables = []
        for t in tables:
            item = dict(t)
            item.pop('row_texts_nfc', None)
            public_tables.append(item)

    return {
        'file': str(path),
        'source_sha256': source_sha256,
        'render_status': render_status,
        'pdf_path': resolved_pdf,
        'page_count': page_count,
        'sections': sections,
        'summary': {
            'paragraphs': len(paragraphs),
            'tables': len(public_tables),
            'layout_tables': sum(1 for t in public_tables if t['is_layout_table']),
            'data_tables': sum(1 for t in public_tables if not t['is_layout_table']),
            'images': len(images),
        },
        'tables': public_tables,
        'images': images,
        'paragraphs': paragraphs,
    }


def format_markdown(mapping, *, include_paragraphs=False):
    """Render a compact token-efficient Markdown summary of tables and document sections."""
    lines = [
        f"# DOCX Map: {Path(mapping['file']).name}",
        f"- sha256: `{mapping['source_sha256'][:12]}...` | paragraphs: {mapping['summary']['paragraphs']} | "
        f"tables: {mapping['summary']['tables']} ({mapping['summary']['data_tables']} data, "
        f"{mapping['summary']['layout_tables']} layout) | images: {mapping['summary'].get('images', 0)}",
    ]
    for s in mapping['sections']:
        lines.append(
            f"- Section {s['section_index']}: {s['page_width_cm']}x{s['page_height_cm']} cm "
            f"({s['orientation']}), printable_width={s['printable_width_cm']} cm"
        )
    if mapping['tables']:
        lines.append('')
        lines.append('## Tables')
        for t in mapping['tables']:
            role = f" [LAYOUT:{t['layout_role']}]" if t['is_layout_table'] else ''
            merged = ' [MERGED]' if t['has_merged_cells'] else ''
            pdf_info = ''
            if t.get('pdf', {}).get('matched'):
                pdf = t['pdf']
                pdf_info = f" | pdf_p={pdf['pdf_pages']} w={pdf['width_cm']}cm ({pdf['measurement_source']})"
            row0 = ' | '.join(c[:22] for c in (t['preview_rows'][0] if t['preview_rows'] else []))
            lines.append(
                f"- `{t['anchor']}` sec={t['section_index']} {t['rows']}x{t['grid_cols']} "
                f"w={t['total_width_cm']}cm ({t['width_pct_of_printable']}% printable) "
                f"cols={t['grid_widths_cm']}cm align={t['alignment']} layout={t['layout']}"
                f"{role}{merged}{pdf_info} | ctx={t['heading_context']!r} | row0=[{row0}]"
            )
    if mapping.get('images'):
        lines.append('')
        lines.append('## Images')
        for img in mapping['images']:
            pdf_info = ''
            if img.get('pdf', {}).get('matched'):
                pdf = img['pdf']
                pdf_info = f" | pdf_p={pdf['pdf_pages']} w={pdf['width_cm']}cm"
            label = img['descr'] or img['name'] or img['media_target'] or ''
            lines.append(
                f"- `{img['anchor']}` sec={img['section_index']} body_idx={img['body_index']} "
                f"{img['width_cm']}x{img['height_cm']}cm ({img['width_pct_of_printable']}% printable, "
                f"ar={img['aspect_ratio']}) align={img['alignment']}{pdf_info} | "
                f"ctx={img['heading_context']!r} | media={img['media_target']!r} desc={label[:40]!r}"
            )
    if include_paragraphs and mapping.get('paragraphs'):
        non_empty = [p for p in mapping['paragraphs'] if p.get('char_count', 0) > 0]
        if non_empty:
            lines.append('')
            lines.append('## Paragraphs')
            for p in non_empty:
                lines.append(
                    f"- `{p['anchor']}` sec={p['section_index']} style={p['style']} | {p['text_preview'][:70]!r}"
                )
    return '\n'.join(lines)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('filepath', help='Input DOCX path')
    parser.add_argument('--format', choices=('json', 'markdown'), default='json')
    parser.add_argument('--with-pdf', action='store_true', help='Render via LibreOffice and extract PDF table geometry')
    parser.add_argument('--pdf-path', help='Existing rendered PDF to correlate instead of invoking LibreOffice')
    parser.add_argument('--renderer', help='Explicit LibreOffice executable path')
    parser.add_argument('--include-paragraphs', action='store_true', help='Include compact ## Paragraphs section in markdown output')
    parser.add_argument('--report', help='Write map output to file')
    args = parser.parse_args(argv)
    try:
        mapping = map_document(
            args.filepath,
            with_pdf=args.with_pdf,
            renderer=args.renderer,
            pdf_path=args.pdf_path,
        )
    except (ValueError, OSError) as exc:
        parser.error(str(exc))

    rendered = (
        format_markdown(mapping, include_paragraphs=args.include_paragraphs)
        if args.format == 'markdown'
        else json.dumps(mapping, ensure_ascii=False, indent=2)
    )
    print(rendered)
    if args.report:
        Path(args.report).write_text(rendered + '\n', encoding='utf-8')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
