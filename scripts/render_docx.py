"""Render a DOCX using LibreOffice, optionally producing page PNGs via PyMuPDF."""
import argparse
import json
import os
import shutil
import subprocess
import tempfile
from pathlib import Path


def find_renderer(explicit=None):
    explicit = explicit or os.environ.get('VNDOCS_RENDERER')
    if explicit:
        return shutil.which(str(explicit)) or (str(Path(explicit).resolve()) if Path(explicit).is_file() else None)
    for name in ('soffice', 'libreoffice'):
        found = shutil.which(name)
        if found:
            return found
    for base in (os.environ.get('PROGRAMFILES'), os.environ.get('PROGRAMFILES(X86)')):
        if base:
            candidate = Path(base) / 'LibreOffice/program/soffice.com'
            if not candidate.is_file():
                candidate = Path(base) / 'LibreOffice/program/soffice.exe'
            if candidate.is_file():
                return str(candidate)
    return None


def _render_annotated_marks(source, destination_pdf, target_dir):
    """Overlay Set-of-Mark bounding boxes and tags ([T#], [IMG#]) onto PNG renders."""
    import pymupdf

    try:
        from . import map_docx
    except ImportError:
        import map_docx

    mapping = map_docx.map_document(source, pdf_path=destination_pdf)
    marks = []
    per_page_marks = {}

    for tbl in mapping.get('tables', []):
        pdf_info = tbl.get('pdf') or {}
        if not pdf_info.get('matched'):
            continue
        is_layout = bool(tbl.get('is_layout_table'))
        prefix = tbl['anchor'].split('#', 1)[0]
        label = f"[{prefix}:LAYOUT]" if is_layout else f"[{prefix}]"
        color = (0.85, 0.40, 0.0) if is_layout else (0.05, 0.35, 0.85)
        page_bboxes = pdf_info.get('page_bboxes') or {}
        if not page_bboxes and pdf_info.get('pdf_pages') and pdf_info.get('bbox_pt'):
            page_bboxes = {str(pdf_info['pdf_pages'][0]): pdf_info['bbox_pt']}
        for page_str, bbox in page_bboxes.items():
            p_num = int(page_str)
            entry = {
                'kind': 'layout_table' if is_layout else 'table',
                'anchor': tbl['anchor'],
                'label': label,
                'page': p_num,
                'bbox_pt': bbox,
                '_color': color,
            }
            marks.append({k: v for k, v in entry.items() if k != '_color'})
            per_page_marks.setdefault(p_num, []).append(entry)

    for img in mapping.get('images', []):
        pdf_info = img.get('pdf') or {}
        if not pdf_info.get('matched'):
            continue
        prefix = img['anchor'].split('#', 1)[0]
        label = f"[{prefix}]"
        color = (0.05, 0.55, 0.25)
        page_bboxes = pdf_info.get('page_bboxes') or {}
        for page_str, bbox in page_bboxes.items():
            p_num = int(page_str)
            entry = {
                'kind': 'image',
                'anchor': img['anchor'],
                'label': label,
                'page': p_num,
                'bbox_pt': bbox,
                '_color': color,
            }
            marks.append({k: v for k, v in entry.items() if k != '_color'})
            per_page_marks.setdefault(p_num, []).append(entry)

    annotated_pages = []
    # Draw strictly in memory and rasterize without modifying clean destination_pdf on disk.
    with pymupdf.open(str(destination_pdf)) as doc:
        for index, page in enumerate(doc):
            p_num = index + 1
            for item in per_page_marks.get(p_num, []):
                x0, y0, x1, y1 = item['bbox_pt']
                col = item['_color']
                rect = pymupdf.Rect(x0, y0, x1, y1)
                page.draw_rect(rect, color=col, width=1.8)
                pill_w = max(36.0, len(item['label']) * 6.5 + 10.0)
                pill_h = 15.0
                pill_y0 = max(page.rect.y0 + 1.0, y0 - pill_h)
                pill_rect = pymupdf.Rect(
                    max(page.rect.x0 + 1.0, x0),
                    pill_y0,
                    min(page.rect.x1 - 1.0, max(page.rect.x0 + 1.0, x0) + pill_w),
                    pill_y0 + pill_h,
                )
                page.draw_rect(pill_rect, color=col, fill=col)
                page.insert_text(
                    (pill_rect.x0 + 4.0, pill_rect.y1 - 3.5),
                    item['label'],
                    fontsize=9.5,
                    color=(1.0, 1.0, 1.0),
                )
            out_png = Path(target_dir) / f'page-{p_num}-marks.png'
            page.get_pixmap(matrix=pymupdf.Matrix(1.8, 1.8)).save(str(out_png))
            annotated_pages.append(str(out_png))
    return annotated_pages, marks


def render_doc(filepath, output_dir=None, renderer=None, timeout=120, png=False, annotate_marks=False):
    source = Path(filepath).resolve()
    report = {'status': 'not_checked', 'pdf': None, 'pages': [], 'message': ''}
    if not source.is_file() or source.suffix.lower() != '.docx':
        report.update(status='fail', message='Input must be an existing DOCX')
        return report
    executable = find_renderer(renderer)
    if not executable:
        report['message'] = 'LibreOffice unavailable; visual QA not performed'
        return report
    target = Path(output_dir).resolve() if output_dir else source.parent / (source.stem + '-render')
    target.mkdir(parents=True, exist_ok=True)
    # Fresh isolated output prevents stale PDFs from being mistaken for success.
    with tempfile.TemporaryDirectory(prefix='vndocs-render-') as temp:
        work = Path(temp)
        profile = work / 'profile'
        command = [executable, f'-env:UserInstallation={profile.as_uri()}', '--headless',
                   '--convert-to', 'pdf', '--outdir', str(work), str(source)]
        try:
            result = subprocess.run(command, capture_output=True, text=True, timeout=timeout, check=False)
        except (OSError, subprocess.TimeoutExpired) as exc:
            report.update(status='fail', message=str(exc))
            return report
        pdf = work / (source.stem + '.pdf')
        if result.returncode != 0 or not pdf.is_file() or not pdf.read_bytes().startswith(b'%PDF-'):
            report.update(status='fail', message=f'Conversion failed (exit {result.returncode}): {result.stdout} {result.stderr}')
            return report
        destination = target / pdf.name
        shutil.copy2(pdf, destination)
        report.update(status='rendered', pdf=str(destination),
                      message='PDF generated; visual inspection still required')
        if png:
            try:
                import pymupdf
                with pymupdf.open(destination) as document:
                    for index, page in enumerate(document):
                        path = target / f'page-{index+1}.png'
                        page.get_pixmap(matrix=pymupdf.Matrix(1.5, 1.5)).save(path)
                        report['pages'].append(str(path))
            except (ImportError, RuntimeError, OSError) as exc:
                report.update(status='partial', message=f'PDF generated, PNG step failed: {exc}')
        if annotate_marks:
            try:
                ann_pages, marks = _render_annotated_marks(source, destination, target)
                report['annotated_pages'] = ann_pages
                report['marks'] = marks
            except (ImportError, RuntimeError, OSError, ValueError) as exc:
                report.update(status='partial', message=f'PDF generated, annotate-marks step failed: {exc}')
    return report


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('filepath')
    parser.add_argument('--output-dir')
    parser.add_argument('--renderer', help='LibreOffice executable path')
    parser.add_argument('--timeout', type=float, default=120)
    parser.add_argument('--png', action='store_true')
    parser.add_argument('--annotate-marks', action='store_true', help='Render page PNGs with [T#] / [IMG#] Set-of-Mark overlays')
    args = parser.parse_args(argv)
    report = render_doc(
        args.filepath, args.output_dir, args.renderer, args.timeout,
        png=args.png, annotate_marks=args.annotate_marks,
    )
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return {'rendered': 0, 'fail': 1, 'not_checked': 2, 'partial': 2}[report['status']]


if __name__ == '__main__':
    raise SystemExit(main())
