"""Render a DOCX via LibreOffice or Windows Microsoft Word COM, with SHA-256 render cache."""
import argparse
import hashlib
import json
import os
import shutil
import subprocess
import tempfile
from pathlib import Path


def _sha256_file(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def find_word_win():
    """Return WINWORD executable path on Windows if detectable without launching COM."""
    if os.name != 'nt':
        return None
    for name in ('winword', 'WINWORD.EXE'):
        found = shutil.which(name)
        if found:
            return found
    for base in (os.environ.get('PROGRAMFILES'), os.environ.get('PROGRAMFILES(X86)')):
        if not base:
            continue
        for sub in (
            'Microsoft Office/root/Office16/WINWORD.EXE',
            'Microsoft Office/Office16/WINWORD.EXE',
            'Microsoft Office/root/Office15/WINWORD.EXE',
            'Microsoft Office/Office15/WINWORD.EXE',
        ):
            candidate = Path(base) / sub
            if candidate.is_file():
                return str(candidate)
    return None


def _render_via_word_com(source, work_pdf, timeout=150):
    """Render DOCX to PDF using Windows Microsoft Word COM automation."""
    if os.name != 'nt':
        return 'not_checked', 'Microsoft Word COM requires Windows'

    ps_script = (
        "$ErrorActionPreference = 'Stop'; "
        "$word = $null; $doc = $null; "
        "try { "
        "$word = New-Object -ComObject Word.Application; "
        "$word.Visible = $false; "
        "$word.DisplayAlerts = 0; "
        "$doc = $word.Documents.Open($env:VNDOCS_COM_SRC, $false, $true); "
        "$doc.ExportAsFixedFormat($env:VNDOCS_COM_DST, 17); "
        "Write-Output 'WORD_COM_OK' "
        "} finally { "
        "if ($null -ne $doc) { try { $doc.Close($false) } catch {} } "
        "if ($null -ne $word) { try { $word.Quit() } catch {} } "
        "}"
    )
    env = dict(os.environ)
    env['VNDOCS_COM_SRC'] = str(source)
    env['VNDOCS_COM_DST'] = str(work_pdf)
    try:
        result = subprocess.run(
            ['powershell', '-NoProfile', '-NonInteractive', '-Command', ps_script],
            capture_output=True,
            text=True,
            timeout=timeout,
            check=False,
            env=env,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        return 'fail', str(exc)

    combined = f'{result.stdout} {result.stderr}'
    if result.returncode != 0:
        low = combined.lower()
        if '80040154' in low or 'class not registered' in low or 'regdb_e_classnotreg' in low:
            return 'not_checked', 'Microsoft Word COM unavailable on this host'
        return 'fail', f'Word COM conversion failed (exit {result.returncode}): {combined.strip()}'

    pdf_file = Path(work_pdf)
    if not pdf_file.is_file() or not pdf_file.read_bytes().startswith(b'%PDF-'):
        return 'fail', f'Word COM did not produce a valid PDF: {combined.strip()}'
    return 'ok', 'Rendered via Microsoft Word COM'


def is_render_fresh(docx_path, manifest_path):
    """Verify render manifest still matches current DOCX and PDF SHA-256 checksums."""
    src = Path(docx_path)
    man = Path(manifest_path)
    if not src.is_file() or not man.is_file():
        return False
    try:
        meta = json.loads(man.read_text(encoding='utf-8'))
    except (json.JSONDecodeError, OSError):
        return False
    pdf_raw = meta.get('pdf')
    if not pdf_raw:
        return False
    pdf_file = Path(pdf_raw)
    if not pdf_file.is_file():
        return False
    return (
        meta.get('source_sha256') == _sha256_file(src)
        and meta.get('pdf_sha256') == _sha256_file(pdf_file)
    )


def find_renderer(explicit=None):
    explicit = explicit or os.environ.get('VNDOCS_RENDERER')
    if explicit == 'word-win':
        return find_word_win() or 'word-win'
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


def render_doc(
    filepath,
    output_dir=None,
    renderer=None,
    timeout=120,
    png=False,
    annotate_marks=False,
    cache_dir=None,
):
    source = Path(filepath).resolve()
    req_renderer = renderer or os.environ.get('VNDOCS_RENDERER')
    engine = 'word-win' if req_renderer == 'word-win' else 'libreoffice'
    report = {
        'status': 'not_checked',
        'pdf': None,
        'pages': [],
        'engine': engine,
        'source_sha256': None,
        'cache_hit': False,
        'manifest': None,
        'message': '',
    }
    if not source.is_file() or source.suffix.lower() != '.docx':
        report.update(status='fail', message='Input must be an existing DOCX')
        return report

    source_sha256 = _sha256_file(source)
    report['source_sha256'] = source_sha256
    target = Path(output_dir).resolve() if output_dir else source.parent / (source.stem + '-render')
    destination = target / (source.stem + '.pdf')

    cache_pdf = cache_manifest = None
    if cache_dir:
        cache_root = Path(cache_dir).resolve()
        cache_key = f'{source_sha256[:16]}-{engine}'
        cache_pdf = cache_root / f'{cache_key}.pdf'
        cache_manifest = cache_root / f'{cache_key}.manifest.json'
        if is_render_fresh(source, cache_manifest) and cache_pdf.is_file():
            target.mkdir(parents=True, exist_ok=True)
            shutil.copy2(cache_pdf, destination)
            report.update(
                status='rendered',
                pdf=str(destination),
                cache_hit=True,
                message='PDF restored from fresh SHA-256 render cache',
            )

    if not report['cache_hit']:
        if engine == 'word-win':
            target.mkdir(parents=True, exist_ok=True)
            with tempfile.TemporaryDirectory(prefix='vndocs-wordcom-') as temp:
                work_pdf = Path(temp) / (source.stem + '.pdf')
                com_state, com_msg = _render_via_word_com(
                    source, work_pdf, timeout=max(float(timeout), 150.0) if timeout == 120 else float(timeout),
                )
                if com_state == 'not_checked':
                    report.update(status='not_checked', message=com_msg)
                    return report
                if com_state != 'ok':
                    report.update(status='fail', message=com_msg)
                    return report
                shutil.copy2(work_pdf, destination)
        else:
            executable = find_renderer(renderer)
            if not executable:
                report['message'] = 'LibreOffice unavailable; visual QA not performed'
                return report
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
                shutil.copy2(pdf, destination)

        report.update(
            status='rendered',
            pdf=str(destination),
            message='PDF generated; visual inspection still required',
        )
        if cache_pdf is not None and cache_manifest is not None:
            cache_pdf.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(destination, cache_pdf)
            cache_manifest.write_text(
                json.dumps({
                    'source': str(source),
                    'source_sha256': source_sha256,
                    'engine': engine,
                    'pdf': str(cache_pdf),
                    'pdf_sha256': _sha256_file(cache_pdf),
                }, ensure_ascii=False, indent=2) + '\n',
                encoding='utf-8',
            )

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

    manifest_out = target / 'render-manifest.json'
    manifest_out.write_text(
        json.dumps({
            'source': str(source),
            'source_sha256': source_sha256,
            'engine': engine,
            'pdf': str(destination),
            'pdf_sha256': _sha256_file(destination),
            'cache_hit': report['cache_hit'],
            'pages': report['pages'],
            'annotated_pages': report.get('annotated_pages', []),
        }, ensure_ascii=False, indent=2) + '\n',
        encoding='utf-8',
    )
    report['manifest'] = str(manifest_out)
    return report


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('filepath')
    parser.add_argument('--output-dir')
    parser.add_argument('--renderer', help='LibreOffice path or "word-win" for Windows Word COM')
    parser.add_argument('--cache-dir', help='Directory for SHA-256 PDF render cache')
    parser.add_argument('--timeout', type=float, default=120)
    parser.add_argument('--png', action='store_true')
    parser.add_argument('--annotate-marks', action='store_true', help='Render page PNGs with [T#] / [IMG#] Set-of-Mark overlays')
    args = parser.parse_args(argv)
    report = render_doc(
        args.filepath, args.output_dir, args.renderer, args.timeout,
        png=args.png, annotate_marks=args.annotate_marks, cache_dir=args.cache_dir,
    )
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return {'rendered': 0, 'fail': 1, 'not_checked': 2, 'partial': 2}[report['status']]


if __name__ == '__main__':
    raise SystemExit(main())
