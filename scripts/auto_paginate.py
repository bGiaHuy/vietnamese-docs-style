"""Bounded, profile-safe auto-pagination solver for fitting DOCX page targets and fixing orphan signatures."""
import argparse
import json
import re
import shutil
import tempfile
from pathlib import Path

from lxml import etree

try:
    from . import docx_package, layout_edit, map_docx, render_docx
except ImportError:
    import docx_package
    import layout_edit
    import map_docx
    import render_docx

W = docx_package.W


def detect_orphan_signature(docx_path, pdf_path):
    """Return True if the last PDF page holds only a signature block (and optional page number/short trailing line)."""
    import pymupdf

    with pymupdf.open(str(pdf_path)) as pdf_doc:
        page_count = len(pdf_doc)
        if page_count <= 1:
            return False
        last_text = map_docx.normalize_nfc(pdf_doc[-1].get_text('text'))

    mapping = map_docx.map_document(docx_path, pdf_path=pdf_path, include_row_texts=True)
    data_on_last = [
        t for t in mapping['tables']
        if not t.get('is_layout_table')
        and page_count in (t.get('pdf') or {}).get('pdf_pages', [])
    ]
    if data_on_last:
        return False

    all_sig_tables = [t for t in mapping['tables'] if t.get('layout_role') == 'signature_block']
    if not all_sig_tables:
        return False
    sig_on_last = [
        t for t in all_sig_tables
        if page_count in (t.get('pdf') or {}).get('pdf_pages', [])
    ]

    # Normalize whitespace so multi-line borderless cell text from PyMuPDF strips cleanly.
    remaining = re.sub(r'(?m)^\s*\d+\s*$', '', last_text)
    remaining_norm = re.sub(r'\s+', ' ', remaining).strip()
    removed_any_sig = False
    for st in (sig_on_last or all_sig_tables):
        for row in st.get('row_texts_nfc', []):
            for cell in row:
                tok = re.sub(r'\s+', ' ', map_docx.normalize_nfc(cell)).strip()
                if tok and tok in remaining_norm:
                    remaining_norm = remaining_norm.replace(tok, '')
                    removed_any_sig = True
                for word_chunk in re.split(r'[;\n]+', cell):
                    sub = re.sub(r'\s+', ' ', map_docx.normalize_nfc(word_chunk)).strip()
                    if len(sub) >= 6 and sub in remaining_norm:
                        remaining_norm = remaining_norm.replace(sub, '')
                        removed_any_sig = True
    if not sig_on_last and not removed_any_sig:
        return False
    compact_remaining = re.sub(r'\s+', '', remaining_norm)
    return len(compact_remaining) <= 80


def _is_heading_or_title_node(p_node, style_name):
    if style_name and style_name.lower().startswith(('heading', 'title', 'subtitle', 'caption')):
        return True
    p_pr = p_node.find(f'{{{W}}}pPr')
    jc = p_pr.find(f'{{{W}}}jc') if p_pr is not None else None
    if jc is not None and jc.get(f'{{{W}}}val') == 'center':
        return True
    runs = p_node.findall(f'{{{W}}}r')
    non_empty_runs = [
        r for r in runs
        if ''.join(t.text or '' for t in r.findall(f'{{{W}}}t')).strip()
    ]
    if non_empty_runs:
        all_bold = all(
            (r.find(f'{{{W}}}rPr/{{{W}}}b') is not None)
            and ((r.find(f'{{{W}}}rPr/{{{W}}}b').get(f'{{{W}}}val') or '1').lower() not in ('0', 'false', 'off'))
            for r in non_empty_runs
        )
        if all_bold:
            return True
    return False


def _build_step_ops(step_idx, mapping, doc_root, profile):
    """Build bounded layout ops for compaction step 1..5."""
    ops = []
    data_tables = [t for t in mapping['tables'] if not t.get('is_layout_table')]

    if step_idx == 1:
        for t in data_tables:
            ops.append({
                'op': 'table.set_cell_margins',
                'target': t['anchor'],
                'top_cm': 0.06,
                'bottom_cm': 0.06,
                'left_cm': 0.12,
                'right_cm': 0.12,
                'clear_cell_overrides': True,
            })
        return 'compact_table_margins', ops

    if step_idx == 2:
        for img in mapping.get('images', []):
            if img.get('width_pct_of_printable', 0) > 65.0 or img.get('height_cm', 0) > 6.5:
                ops.append({
                    'op': 'image.scale',
                    'target': img['anchor'],
                    'factor': 0.82,
                })
        return 'scale_large_images', ops

    if step_idx == 3:
        floor_pt = 12.0 if profile in layout_edit.ADMIN_PROFILES else 11.0
        for t in data_tables:
            ops.append({
                'op': 'table.set_font_size',
                'target': t['anchor'],
                'size_pt': floor_pt,
                'preserve_header_size': True,
            })
        return 'compact_table_font_size', ops

    if step_idx == 4:
        body = doc_root.find(f'{{{W}}}body')
        body_children = list(body)
        targets = []
        for p in mapping['paragraphs']:
            if p.get('char_count', 0) <= 0:
                continue
            p_node = body_children[p['body_index']]
            if _is_heading_or_title_node(p_node, p.get('style')):
                continue
            targets.append(p['anchor'])
        if targets:
            if profile == 'party':
                ops.append({
                    'op': 'paragraph.set_format',
                    'targets': targets,
                    'line_spacing_exact_pt': 18.0,
                    'space_after_pt': 6.0,
                })
            elif profile in ('administrative', 'minutes-administrative'):
                ops.append({
                    'op': 'paragraph.set_format',
                    'targets': targets,
                    'line_spacing': 1.15,
                    'space_after_pt': 4.0,
                })
            else:
                ops.append({
                    'op': 'paragraph.set_format',
                    'targets': targets,
                    'line_spacing': 1.12,
                    'space_after_pt': 3.0,
                })
        return 'tighten_paragraph_spacing', ops

    if step_idx == 5:
        if profile == 'party':
            return 'trim_section_margins', []
        target_mar = 2.0 if profile in ('administrative', 'minutes-administrative') else 1.8
        for sec in mapping['sections']:
            mar = sec['margins_cm']
            upd = {}
            if mar['top'] > target_mar + 0.05:
                upd['top'] = target_mar
            if mar['bottom'] > target_mar + 0.05:
                upd['bottom'] = target_mar
            if upd:
                ops.append({
                    'op': 'section.set_layout',
                    'section_index': sec['section_index'],
                    'margins_cm': upd,
                })
        return 'trim_section_margins', ops

    return 'unknown', []


def auto_paginate(
    source,
    destination,
    *,
    target_pages=None,
    fix_orphan_signature=False,
    profile=None,
    max_steps=5,
    renderer=None,
    cache_dir=None,
):
    """Apply bounded profile-safe compaction steps to meet page target or absorb orphan signature block."""
    src_path = Path(source).resolve()
    dst_path = Path(destination).resolve()
    if not src_path.is_file() or src_path.suffix.lower() != '.docx':
        raise FileNotFoundError(f'Source DOCX not found: {source}')
    if src_path == dst_path:
        raise ValueError('Destination must differ from source to preserve original document')
    if target_pages is not None and (not isinstance(target_pages, int) or isinstance(target_pages, bool) or target_pages < 1):
        raise ValueError(f'target_pages must be positive integer, got {target_pages!r}')

    with tempfile.TemporaryDirectory(prefix='vndocs-paginate-') as tmp_root:
        tmp_dir = Path(tmp_root)
        ren0 = render_docx.render_doc(
            src_path, output_dir=tmp_dir / 'step0_render', renderer=renderer, cache_dir=cache_dir,
        )
        if ren0['status'] == 'not_checked':
            return {'status': 'not_checked', 'message': ren0.get('message', 'Renderer unavailable')}
        if ren0['status'] not in ('rendered', 'partial') or not ren0.get('pdf'):
            return {'status': 'fail', 'message': ren0.get('message', 'Initial render failed')}

        import pymupdf
        with pymupdf.open(ren0['pdf']) as pdf0:
            pages_before = len(pdf0)
        orphan_before = detect_orphan_signature(src_path, ren0['pdf'])

        goal_pages = target_pages
        if fix_orphan_signature and goal_pages is None:
            goal_pages = pages_before - 1 if (orphan_before and pages_before > 1) else pages_before
        if goal_pages is None:
            raise ValueError('Specify target_pages or fix_orphan_signature=True')

        if pages_before <= goal_pages and (not fix_orphan_signature or not orphan_before):
            dst_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src_path, dst_path)
            return {
                'status': 'already_fits',
                'output': str(dst_path),
                'pages_before': pages_before,
                'pages_after': pages_before,
                'target_pages': goal_pages,
                'orphan_signature_before': orphan_before,
                'orphan_signature_after': orphan_before,
                'steps_applied': [],
            }

        cur_docx = src_path
        best_docx = src_path
        best_pages = pages_before
        best_orphan = orphan_before
        steps_applied = []

        limit = min(5, max(1, int(max_steps)))
        for step_idx in range(1, limit + 1):
            cur_map = map_docx.map_document(cur_docx)
            _, orig_parts = docx_package.read_package(cur_docx)
            parser = etree.XMLParser(resolve_entities=False, no_network=True)
            doc_root = etree.fromstring(orig_parts['word/document.xml'], parser)

            step_name, ops = _build_step_ops(step_idx, cur_map, doc_root, profile)
            if not ops:
                continue

            cand_docx = tmp_dir / f'step{step_idx}.docx'
            try:
                edit_res = layout_edit.apply_layout_plan(
                    cur_docx, cand_docx, ops, profile=profile,
                )
            except layout_edit.LayoutEditError:
                continue

            prof_chk = edit_res.get('profile_check') or {}
            if prof_chk.get('new_issues'):
                continue

            ren_s = render_docx.render_doc(
                cand_docx, output_dir=tmp_dir / f'step{step_idx}_render', renderer=renderer, cache_dir=cache_dir,
            )
            if ren_s['status'] not in ('rendered', 'partial') or not ren_s.get('pdf'):
                continue

            with pymupdf.open(ren_s['pdf']) as pdf_s:
                p_after = len(pdf_s)
            orphan_after = detect_orphan_signature(cand_docx, ren_s['pdf'])

            steps_applied.append({
                'step': step_idx,
                'name': step_name,
                'operations_count': len(ops),
                'pages': p_after,
                'orphan_signature': orphan_after,
            })

            # Accumulate layout tightening across steps cleanly.
            cur_docx = cand_docx
            if (p_after, orphan_after) <= (best_pages, best_orphan):
                best_docx = cand_docx
                best_pages = p_after
                best_orphan = orphan_after

            if p_after <= goal_pages and (not fix_orphan_signature or not orphan_after):
                break

        dst_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(best_docx, dst_path)
        goal_met = best_pages <= goal_pages and (not fix_orphan_signature or not best_orphan)
        return {
            'status': 'ok' if goal_met else 'best_effort_limit_reached',
            'output': str(dst_path),
            'pages_before': pages_before,
            'pages_after': best_pages,
            'target_pages': goal_pages,
            'orphan_signature_before': orphan_before,
            'orphan_signature_after': best_orphan,
            'steps_applied': steps_applied,
        }


def main(argv=None):
    docx_package.ensure_utf8_stdio()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source', help='Source DOCX path (preserved unmodified)')
    parser.add_argument('destination', help='Destination DOCX path')
    parser.add_argument('--target-pages', type=int, help='Desired maximum page count')
    parser.add_argument('--fix-orphan-signature', action='store_true', help='Pull trailing orphan signature block onto previous page')
    parser.add_argument('--profile', help='Vietnamese document profile guard (administrative, party, academic, ...)')
    parser.add_argument('--max-steps', type=int, default=5, help='Maximum compaction ladder steps (1..5)')
    parser.add_argument('--renderer', help='LibreOffice path or "word-win"')
    parser.add_argument('--cache-dir', help='Optional SHA-256 render cache directory')
    args = parser.parse_args(argv)
    try:
        report = auto_paginate(
            args.source,
            args.destination,
            target_pages=args.target_pages,
            fix_orphan_signature=args.fix_orphan_signature,
            profile=args.profile,
            max_steps=args.max_steps,
            renderer=args.renderer,
            cache_dir=args.cache_dir,
        )
    except (FileNotFoundError, ValueError, OSError) as exc:
        print(json.dumps({'status': 'fail', 'error': str(exc)}, ensure_ascii=False, indent=2))
        return 1

    print(json.dumps(report, ensure_ascii=False, indent=2))
    if report['status'] in ('ok', 'already_fits'):
        return 0
    if report['status'] in ('not_checked', 'best_effort_limit_reached'):
        return 2
    return 1


if __name__ == '__main__':
    raise SystemExit(main())
