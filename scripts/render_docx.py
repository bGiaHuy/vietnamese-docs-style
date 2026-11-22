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


def render_doc(filepath, output_dir=None, renderer=None, timeout=120, png=False):
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
    return report


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('filepath')
    parser.add_argument('--output-dir')
    parser.add_argument('--renderer', help='LibreOffice executable path')
    parser.add_argument('--timeout', type=float, default=120)
    parser.add_argument('--png', action='store_true')
    args = parser.parse_args(argv)
    report = render_doc(args.filepath, args.output_dir, args.renderer, args.timeout, args.png)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return {'rendered': 0, 'fail': 1, 'not_checked': 2, 'partial': 2}[report['status']]


if __name__ == '__main__':
    raise SystemExit(main())
