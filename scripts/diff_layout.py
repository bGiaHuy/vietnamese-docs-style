"""Measure visual/spatial layout deltas and pagination ripple between before/after DOCX files."""
import argparse
import json
import re
import tempfile
from pathlib import Path

try:
    from . import docx_package, map_docx, render_docx
except ImportError:
    import docx_package
    import map_docx
    import render_docx

PT_PER_CM = map_docx.PT_PER_CM


def _norm_space(text):
    return re.sub(r'\s+', ' ', map_docx.normalize_nfc(text or '')).strip()


def _resolve_target_anchors(tables, target):
    """Resolve user/agent --target (exact `T1#hash` or ordinal `T1`) to matching anchor(s)."""
    if not target:
        return set()
    matched = set()
    for tbl in tables:
        anchor = tbl['anchor']
        prefix = anchor.split('#', 1)[0]
        if target == anchor or target == prefix:
            matched.add(anchor)
    return matched


def _pair_tables(before_tables, after_tables):
    """Pair before/after tables by anchor first, then fallback by positional ordinal."""
    after_by_anchor = {t['anchor']: t for t in after_tables}
    after_by_ord = {t['ordinal']: t for t in after_tables}
    paired = []
    used_after = set()
    for bt in before_tables:
        at = after_by_anchor.get(bt['anchor'])
        if at is None:
            cand = after_by_ord.get(bt['ordinal'])
            if cand is not None and cand['anchor'] not in used_after:
                at = cand
        if at is not None:
            used_after.add(at['anchor'])
        paired.append((bt, at))
    return paired


def _save_table_crop(pdf_path, pdf_info, out_path, pad_pt=18.0, zoom=2.0):
    """Render a cropped zoomed PNG around a table bbox on its first matched PDF page."""
    if not pdf_path or not pdf_info or not pdf_info.get('matched'):
        return None
    pages = pdf_info.get('pdf_pages') or []
    bbox = pdf_info.get('bbox_pt')
    if not pages or not bbox or len(bbox) != 4:
        return None
    import pymupdf

    with pymupdf.open(str(pdf_path)) as doc:
        page_index = pages[0] - 1
        if page_index < 0 or page_index >= len(doc):
            return None
        page = doc[page_index]
        clip = pymupdf.Rect(
            max(page.rect.x0, bbox[0] - pad_pt),
            max(page.rect.y0, bbox[1] - pad_pt),
            min(page.rect.x1, bbox[2] + pad_pt),
            min(page.rect.y1, bbox[3] + pad_pt),
        )
        if clip.is_empty or clip.width <= 1 or clip.height <= 1:
            return None
        pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), clip=clip)
        out_file = Path(out_path)
        out_file.parent.mkdir(parents=True, exist_ok=True)
        pix.save(str(out_file))
        return str(out_file)


def _table_diff_record(bt, at):
    """Compute OOXML vs rendered PDF geometric comparison for a paired table."""
    b_pdf = bt.get('pdf') or {}
    a_pdf = (at or {}).get('pdf') or {}

    w_before = b_pdf.get('width_cm')
    w_after = a_pdf.get('width_cm')
    h_before = b_pdf.get('height_cm')
    h_after = a_pdf.get('height_cm')

    b_bbox = b_pdf.get('bbox_pt')
    a_bbox = a_pdf.get('bbox_pt')
    x0_before = round(b_bbox[0] / PT_PER_CM, 2) if b_bbox else None
    x0_after = round(a_bbox[0] / PT_PER_CM, 2) if a_bbox else None

    expected_ooxml_w = at['total_width_cm'] if at else None
    ooxml_w_before = bt['total_width_cm']

    render_vs_ooxml_error = (
        round(abs(w_after - expected_ooxml_w), 2)
        if w_after is not None and expected_ooxml_w is not None
        else None
    )

    return {
        'anchor': bt['anchor'],
        'anchor_after': at['anchor'] if at else None,
        'ordinal': bt['ordinal'],
        'is_layout_table': bt['is_layout_table'],
        'layout_role': bt['layout_role'],
        'measurement_source_before': b_pdf.get('measurement_source'),
        'measurement_source_after': a_pdf.get('measurement_source'),
        'matched_before': bool(b_pdf.get('matched')),
        'matched_after': bool(a_pdf.get('matched')),
        'ooxml_width_before_cm': ooxml_w_before,
        'expected_ooxml_width_cm': expected_ooxml_w,
        'expected_width_delta_cm': (
            round(expected_ooxml_w - ooxml_w_before, 2)
            if expected_ooxml_w is not None
            else None
        ),
        'width_before_cm': w_before,
        'width_after_cm': w_after,
        'width_delta_cm': (
            round(w_after - w_before, 2)
            if w_before is not None and w_after is not None
            else None
        ),
        'height_before_cm': h_before,
        'height_after_cm': h_after,
        'height_delta_cm': (
            round(h_after - h_before, 2)
            if h_before is not None and h_after is not None
            else None
        ),
        'x0_before_cm': x0_before,
        'x0_after_cm': x0_after,
        'x0_delta_cm': (
            round(x0_after - x0_before, 2)
            if x0_before is not None and x0_after is not None
            else None
        ),
        'alignment_before': bt.get('alignment'),
        'alignment_after': at.get('alignment') if at else None,
        'render_vs_ooxml_error_cm': render_vs_ooxml_error,
        'overflow_right_mm_before': b_pdf.get('overflow_right_mm'),
        'overflow_right_mm_after': a_pdf.get('overflow_right_mm'),
        'pdf_pages_before': b_pdf.get('pdf_pages', []),
        'pdf_pages_after': a_pdf.get('pdf_pages', []),
        'spans_pages_before': b_pdf.get('spans_pages'),
        'spans_pages_after': a_pdf.get('spans_pages'),
    }


def _find_paragraph_pages(paragraphs, pdf_path):
    """Locate unique non-empty body paragraphs in normalized per-page PDF text."""
    import pymupdf

    norm_texts = [_norm_space(p.get('text_preview', '')) for p in paragraphs]
    freq = {}
    for txt in norm_texts:
        if txt:
            freq[txt] = freq.get(txt, 0) + 1

    with pymupdf.open(str(pdf_path)) as doc:
        page_texts = [
            _norm_space(page.get_text('text'))
            for page in doc
        ]

    hits = {}
    for p, txt in zip(paragraphs, norm_texts):
        if len(txt) < 6 or freq.get(txt, 0) != 1:
            continue
        snippet = txt[:55]
        matched_pages = [
            idx for idx, pt in enumerate(page_texts, start=1)
            if snippet in pt
        ]
        if not matched_pages and len(snippet) > 25:
            head = snippet[:28]
            matched_pages = [
                idx for idx, pt in enumerate(page_texts, start=1)
                if head in pt
            ]
        if matched_pages:
            hits[p['anchor']] = {
                'pages': matched_pages,
                'preview': p.get('text_preview', '')[:60],
            }
    return hits


def _pair_images(before_images, after_images):
    """Pair before/after images by anchor first, then by document ordinal."""
    by_anchor = {img['anchor']: img for img in after_images}
    by_ord = {img['ordinal']: img for img in after_images}
    used_after = set()
    pairs = []
    for b_img in before_images:
        a_img = by_anchor.get(b_img['anchor'])
        if a_img is None or a_img['anchor'] in used_after:
            a_img = by_ord.get(b_img['ordinal'])
        if a_img is not None and a_img['anchor'] in used_after:
            a_img = None
        if a_img is not None:
            used_after.add(a_img['anchor'])
        pairs.append((b_img, a_img))
    return pairs


def _image_diff_record(bi, ai):
    """Build OOXML and rendered measurement diff for one paired image."""
    b_pdf = bi.get('pdf') or {}
    a_pdf = (ai or {}).get('pdf') or {}
    b_w = bi.get('width_cm')
    a_w = ai.get('width_cm') if ai else None
    b_ar = bi.get('aspect_ratio')
    a_ar = ai.get('aspect_ratio') if ai else None
    b_rw = b_pdf.get('width_cm')
    a_rw = a_pdf.get('width_cm')
    return {
        'anchor': bi['anchor'],
        'name': bi.get('name') or bi.get('media_target') or '',
        'ooxml_width_cm_before': b_w,
        'ooxml_width_cm_after': a_w,
        'ooxml_width_delta_cm': (
            round(a_w - b_w, 3) if b_w is not None and a_w is not None else None
        ),
        'ooxml_height_cm_before': bi.get('height_cm'),
        'ooxml_height_cm_after': ai.get('height_cm') if ai else None,
        'aspect_ratio_before': b_ar,
        'aspect_ratio_after': a_ar,
        'aspect_ratio_delta_pct': (
            round(abs(a_ar - b_ar) / b_ar * 100.0, 2)
            if b_ar and a_ar is not None
            else None
        ),
        'alignment_before': bi.get('alignment'),
        'alignment_after': ai.get('alignment') if ai else None,
        'rendered_width_cm_before': b_rw,
        'rendered_width_cm_after': a_rw,
        'rendered_width_delta_cm': (
            round(a_rw - b_rw, 3) if b_rw is not None and a_rw is not None else None
        ),
        'rendered_height_cm_before': b_pdf.get('height_cm'),
        'rendered_height_cm_after': a_pdf.get('height_cm'),
        'overflow_right_mm_before': b_pdf.get('overflow_right_mm'),
        'overflow_right_mm_after': a_pdf.get('overflow_right_mm'),
        'pdf_pages_before': b_pdf.get('pdf_pages', []),
        'pdf_pages_after': a_pdf.get('pdf_pages', []),
    }


def _resolve_image_anchors(images, target):
    """Resolve exact IMG#hash or ordinal IMG# prefix tokens into matching anchors."""
    if not target:
        return set()
    tokens = [tok.strip() for tok in target.split(',') if tok.strip()]
    matched = set()
    for img in images:
        anchor = img['anchor']
        prefix = anchor.split('#', 1)[0]
        if anchor in tokens or prefix in tokens:
            matched.add(anchor)
    return matched


def _detect_non_target_shifts(
    before_map, after_map, paired_tables, target_anchors, pdf_before, pdf_after, paired_images=(),
):
    """Identify any non-target tables, images, or body paragraphs whose rendered PDF page(s) shifted."""
    shifts = []
    for bt, at in paired_tables:
        if bt['anchor'] in target_anchors or (at and at['anchor'] in target_anchors):
            continue
        b_pages = (bt.get('pdf') or {}).get('pdf_pages') or []
        a_pages = ((at or {}).get('pdf') or {}).get('pdf_pages') or []
        if b_pages and a_pages and b_pages != a_pages:
            shifts.append({
                'kind': 'table',
                'anchor': bt['anchor'],
                'role': bt.get('layout_role') or 'data',
                'pages_before': b_pages,
                'pages_after': a_pages,
            })

    for bi, ai in paired_images:
        if bi['anchor'] in target_anchors or (ai and ai['anchor'] in target_anchors):
            continue
        b_pages = (bi.get('pdf') or {}).get('pdf_pages') or []
        a_pages = ((ai or {}).get('pdf') or {}).get('pdf_pages') or []
        if b_pages and a_pages and b_pages != a_pages:
            shifts.append({
                'kind': 'image',
                'anchor': bi['anchor'],
                'name': bi.get('name') or bi.get('media_target') or '',
                'pages_before': b_pages,
                'pages_after': a_pages,
            })

    b_para_hits = _find_paragraph_pages(before_map.get('paragraphs', []), pdf_before)
    a_para_hits = _find_paragraph_pages(after_map.get('paragraphs', []), pdf_after)
    for anchor, b_info in b_para_hits.items():
        a_info = a_para_hits.get(anchor)
        if a_info and b_info['pages'] != a_info['pages']:
            shifts.append({
                'kind': 'paragraph',
                'anchor': anchor,
                'preview': b_info['preview'],
                'pages_before': b_info['pages'],
                'pages_after': a_info['pages'],
            })
    return shifts


def _render_missing_pdfs(before_path, after_path, pdf_before, pdf_after, render_root, renderer, cache_dir=None):
    """Resolve or render before/after PDFs via configured renderer with optional SHA-256 cache."""
    resolved_before = Path(pdf_before) if pdf_before else None
    resolved_after = Path(pdf_after) if pdf_after else None

    need_render = (resolved_before is None) or (resolved_after is None)
    if need_render and renderer != 'word-win' and cache_dir is None and not render_docx.find_renderer(renderer):
        return 'not_checked', None, None, 'LibreOffice renderer not found; skipping PDF layout diff'

    if resolved_before is None:
        b_dir = Path(render_root) / 'before'
        b_dir.mkdir(parents=True, exist_ok=True)
        res = render_docx.render_doc(before_path, output_dir=b_dir, renderer=renderer, cache_dir=cache_dir)
        if res['status'] == 'not_checked':
            return 'not_checked', None, None, res.get('message', 'Renderer unavailable')
        if res['status'] not in ('rendered', 'partial') or not res.get('pdf'):
            return 'fail', None, None, f"Before render failed: {res.get('message', res['status'])}"
        resolved_before = Path(res['pdf'])
    elif not resolved_before.is_file():
        return 'fail', None, None, f'Provided --pdf-before not found: {resolved_before}'

    if resolved_after is None:
        a_dir = Path(render_root) / 'after'
        a_dir.mkdir(parents=True, exist_ok=True)
        res = render_docx.render_doc(after_path, output_dir=a_dir, renderer=renderer, cache_dir=cache_dir)
        if res['status'] == 'not_checked':
            return 'not_checked', None, None, res.get('message', 'Renderer unavailable')
        if res['status'] not in ('rendered', 'partial') or not res.get('pdf'):
            return 'fail', str(resolved_before), None, f"After render failed: {res.get('message', res['status'])}"
        resolved_after = Path(res['pdf'])
    elif not resolved_after.is_file():
        return 'fail', str(resolved_before), None, f'Provided --pdf-after not found: {resolved_after}'

    return 'ok', str(resolved_before), str(resolved_after), 'Rendered and compared successfully'


def compare_renderers(docx_path, output_dir=None, cache_dir=None):
    """Compare page counts and matched table geometry between LibreOffice and Windows Word COM."""
    src = Path(docx_path).resolve()
    if not src.is_file():
        raise FileNotFoundError(f'DOCX not found: {docx_path}')

    temp_ctx = None
    if output_dir:
        out_root = Path(output_dir)
        out_root.mkdir(parents=True, exist_ok=True)
    else:
        temp_ctx = tempfile.TemporaryDirectory(prefix='vndocs-cmp-renderers-')
        out_root = Path(temp_ctx.name)

    try:
        lo_ren = render_docx.render_doc(src, output_dir=out_root / 'libreoffice', cache_dir=cache_dir)
        wd_ren = render_docx.render_doc(src, output_dir=out_root / 'word-win', renderer='word-win', cache_dir=cache_dir)
        if lo_ren['status'] not in ('rendered', 'partial') or wd_ren['status'] not in ('rendered', 'partial'):
            return {
                'status': 'not_checked',
                'docx': str(src),
                'libreoffice': lo_ren,
                'word_win': wd_ren,
                'message': 'Both LibreOffice and Microsoft Word COM are required to compare renderers.',
            }

        lo_map = map_docx.map_document(src, pdf_path=lo_ren['pdf'])
        wd_map = map_docx.map_document(src, pdf_path=wd_ren['pdf'])
        table_parity = []
        for lt, wt in zip(lo_map['tables'], wd_map['tables']):
            l_pdf, w_pdf = lt.get('pdf') or {}, wt.get('pdf') or {}
            table_parity.append({
                'anchor': lt['anchor'],
                'lo_pages': l_pdf.get('pdf_pages', []),
                'word_pages': w_pdf.get('pdf_pages', []),
                'lo_width_cm': l_pdf.get('width_cm'),
                'word_width_cm': w_pdf.get('width_cm'),
                'width_delta_cm': (
                    round(abs(l_pdf['width_cm'] - w_pdf['width_cm']), 2)
                    if l_pdf.get('width_cm') is not None and w_pdf.get('width_cm') is not None
                    else None
                ),
            })
        return {
            'status': 'ok',
            'docx': str(src),
            'lo_page_count': lo_map['page_count'],
            'word_page_count': wd_map['page_count'],
            'page_count_delta': (wd_map['page_count'] or 0) - (lo_map['page_count'] or 0),
            'table_parity': table_parity,
        }
    finally:
        if temp_ctx is not None:
            temp_ctx.cleanup()


def diff_layout(
    before_docx,
    after_docx,
    *,
    target=None,
    pdf_before=None,
    pdf_after=None,
    output_dir=None,
    crop_dir=None,
    renderer=None,
    cache_dir=None,
):
    """Compare rendered PDF geometry and non-target pagination between two DOCX files."""
    before_path = Path(before_docx)
    after_path = Path(after_docx)
    if not before_path.is_file():
        raise FileNotFoundError(f'Before DOCX not found: {before_docx}')
    if not after_path.is_file():
        raise FileNotFoundError(f'After DOCX not found: {after_docx}')

    temp_ctx = None
    if output_dir:
        render_root = Path(output_dir)
        render_root.mkdir(parents=True, exist_ok=True)
    else:
        temp_ctx = tempfile.TemporaryDirectory(prefix='vndocs-diff-layout-')
        render_root = Path(temp_ctx.name)

    try:
        render_state, resolved_b_pdf, resolved_a_pdf, message = _render_missing_pdfs(
            before_path, after_path, pdf_before, pdf_after, render_root, renderer, cache_dir=cache_dir,
        )
        if render_state != 'ok':
            before_only = map_docx.map_document(before_path)
            after_only = map_docx.map_document(after_path)
            return {
                'status': render_state,
                'message': message,
                'before_docx': str(before_path),
                'after_docx': str(after_path),
                'target': target,
                'before_ooxml_tables': before_only['tables'],
                'after_ooxml_tables': after_only['tables'],
                'before_ooxml_images': before_only.get('images', []),
                'after_ooxml_images': after_only.get('images', []),
            }

        before_map = map_docx.map_document(
            before_path, pdf_path=resolved_b_pdf, include_row_texts=True,
        )
        after_map = map_docx.map_document(
            after_path, pdf_path=resolved_a_pdf, include_row_texts=True,
        )

        paired = _pair_tables(before_map['tables'], after_map['tables'])
        table_diffs = [_table_diff_record(bt, at) for bt, at in paired]
        paired_images = _pair_images(
            before_map.get('images', []), after_map.get('images', []),
        )
        image_diffs = [_image_diff_record(bi, ai) for bi, ai in paired_images]

        target_anchors = _resolve_target_anchors(before_map['tables'], target)
        target_anchors |= _resolve_image_anchors(before_map.get('images', []), target)
        if not target_anchors and not target:
            # Auto-select modified tables or images when --target is omitted.
            for rec in table_diffs:
                if (
                    rec['expected_width_delta_cm'] not in (None, 0.0)
                    or rec['alignment_before'] != rec['alignment_after']
                ):
                    target_anchors.add(rec['anchor'])
            for rec in image_diffs:
                if (
                    rec['ooxml_width_delta_cm'] not in (None, 0.0)
                    or rec['alignment_before'] != rec['alignment_after']
                ):
                    target_anchors.add(rec['anchor'])

        target_record = None
        target_kind = None
        for rec in table_diffs:
            if rec['anchor'] in target_anchors:
                target_record = rec
                target_kind = 'table'
                break
        if target_record is None:
            for rec in image_diffs:
                if rec['anchor'] in target_anchors:
                    target_record = rec
                    target_kind = 'image'
                    break
        if target and target_record is None:
            return {
                'status': 'fail',
                'message': f'Target {target!r} not found in {before_path.name}',
                'before_docx': str(before_path),
                'after_docx': str(after_path),
                'target': target,
                'table_diffs': table_diffs,
                'image_diffs': image_diffs,
            }

        crops = None
        if crop_dir and target_record is not None:
            if target_kind == 'image':
                b_match = next(
                    (bi for bi, _ in paired_images if bi['anchor'] == target_record['anchor']),
                    None,
                )
                a_match = next(
                    (ai for bi, ai in paired_images if bi['anchor'] == target_record['anchor']),
                    None,
                )
            else:
                b_match = next(
                    (bt for bt, _ in paired if bt['anchor'] == target_record['anchor']),
                    None,
                )
                a_match = next(
                    (at for bt, at in paired if bt['anchor'] == target_record['anchor']),
                    None,
                )
            stem = target_record['anchor'].replace('#', '_')
            crop_root = Path(crop_dir)
            crops = {
                'before': _save_table_crop(
                    resolved_b_pdf,
                    b_match.get('pdf') if b_match else None,
                    crop_root / f'{stem}_before.png',
                ),
                'after': _save_table_crop(
                    resolved_a_pdf,
                    a_match.get('pdf') if a_match else None,
                    crop_root / f'{stem}_after.png',
                ),
            }

        non_target_shifts = _detect_non_target_shifts(
            before_map, after_map, paired, target_anchors, resolved_b_pdf, resolved_a_pdf,
            paired_images=paired_images,
        )
        page_delta = (after_map['page_count'] or 0) - (before_map['page_count'] or 0)

        return {
            'status': 'ok',
            'message': message,
            'before_docx': str(before_path),
            'after_docx': str(after_path),
            'pdf_before': resolved_b_pdf if output_dir or pdf_before else None,
            'pdf_after': resolved_a_pdf if output_dir or pdf_after else None,
            'page_count_before': before_map['page_count'],
            'page_count_after': after_map['page_count'],
            'page_count_delta': page_delta,
            'target': target,
            'target_kind': target_kind,
            'target_metrics': target_record,
            'crops': crops,
            'non_target_page_shifts': non_target_shifts,
            'pagination_ripple_detected': bool(non_target_shifts) or page_delta != 0,
            'table_diffs': table_diffs,
            'image_diffs': image_diffs,
        }
    finally:
        if temp_ctx is not None:
            temp_ctx.cleanup()


def main(argv=None):
    docx_package.ensure_utf8_stdio()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('before_docx', help='Original DOCX path before layout edit')
    parser.add_argument('after_docx', help='Edited DOCX path after layout edit')
    parser.add_argument('--target', help='Target table/image anchor (e.g. T1#a3f9, T1, IMG1#a3f9, IMG1)')
    parser.add_argument('--pdf-before', help='Existing rendered PDF for before_docx')
    parser.add_argument('--pdf-after', help='Existing rendered PDF for after_docx')
    parser.add_argument('--output-dir', help='Directory to persist rendered PDFs')
    parser.add_argument('--crop-dir', help='Directory to write zoomed PNG crops of the target table or image')
    parser.add_argument('--renderer', help='LibreOffice path or "word-win" for Windows Word COM')
    parser.add_argument('--cache-dir', help='Directory for SHA-256 PDF render cache')
    parser.add_argument('--report', help='Write JSON report to file')
    args = parser.parse_args(argv)

    try:
        result = diff_layout(
            args.before_docx,
            args.after_docx,
            target=args.target,
            pdf_before=args.pdf_before,
            pdf_after=args.pdf_after,
            output_dir=args.output_dir,
            crop_dir=args.crop_dir,
            renderer=args.renderer,
            cache_dir=args.cache_dir,
        )
    except (FileNotFoundError, ValueError, OSError) as exc:
        parser.error(str(exc))

    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    print(rendered)
    if args.report:
        report_path = Path(args.report)
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(rendered + '\n', encoding='utf-8')
    return 1 if result.get('status') == 'fail' else 0


if __name__ == '__main__':
    raise SystemExit(main())
