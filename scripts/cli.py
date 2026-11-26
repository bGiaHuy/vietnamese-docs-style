"""Unified CLI entrypoint (`vndocs`) for Vietnamese DOCX tooling."""
import argparse
import importlib
import importlib.metadata
import json
import platform
import sys

try:
    from . import (
        auto_paginate,
        create_demo,
        create_templates,
        diff_layout,
        docx_package,
        edit_docx,
        fill_template,
        generate_docx,
        inspect_docx,
        layout_edit,
        map_docx,
        ooxml_order,
        package_skill,
        render_docx,
        validate_docx,
    )
except ImportError:
    import auto_paginate
    import create_demo
    import create_templates
    import diff_layout
    import docx_package
    import edit_docx
    import fill_template
    import generate_docx
    import inspect_docx
    import layout_edit
    import map_docx
    import ooxml_order
    import package_skill
    import render_docx
    import validate_docx

COMMANDS = {
    'generate': generate_docx.main,
    'inspect': inspect_docx.main,
    'validate': validate_docx.main,
    'render': render_docx.main,
    'fill': fill_template.main,
    'edit': edit_docx.main,
    'map': map_docx.main,
    'layout-edit': layout_edit.main,
    'diff-layout': diff_layout.main,
    'paginate': auto_paginate.main,
    'normalize-ooxml': ooxml_order.main,
    'demo': create_demo.main,
    'templates': create_templates.main,
    'package': package_skill.main,
}

HELP_TEXT = """usage: vndocs <command> [args...]

Vietnamese DOCX creation, inspection, validation, rendering, and packaging CLI.

commands:
  doctor           Check Python, DOCX runtime dependencies, and LibreOffice renderer
  generate         Build a styled DOCX from a multi-profile JSON spec (and optional --validate)
  inspect          Inspect DOCX geometry, styles, placeholders, and U+00D0 encoding issues
  validate         Inspect a DOCX file against profile or JSON rules
  render           Render a DOCX file to PDF (and optional PNGs) via LibreOffice
  fill             Fill {{token}} placeholders in a DOCX template
  edit             Patch verified plain-text slots in a retained DOCX
  map              Build a compact structural + spatial anchor map for DOCX editing
  layout-edit      Apply table-scoped OOXML layout operations with XML preservation checks
  diff-layout      Compare rendered PDF geometry and pagination ripple before/after
  paginate         Fit page targets or absorb orphan signature blocks cleanly
  normalize-ooxml  Normalize ECMA-376 child ordering in modified WordprocessingML XML parts
  demo             Generate a reproducible multi-section demo template and filled DOCX
  templates        Regenerate project-authored starter templates
  package          Build a clean distributable skill ZIP

Run `vndocs <command> --help` for command-specific options."""


def _check_module(module_name, dist_name):
    try:
        importlib.import_module(module_name)
    except ImportError:
        return {'installed': False, 'version': None}
    try:
        version = importlib.metadata.version(dist_name)
    except importlib.metadata.PackageNotFoundError:
        version = None
    return {'installed': True, 'version': version}


def doctor_report(renderer=None):
    """Inspect runtime dependencies and optional LibreOffice PDF/PNG renderer."""
    python_ok = sys.version_info >= (3, 10)
    deps = {
        'python-docx': _check_module('docx', 'python-docx'),
        'lxml': _check_module('lxml', 'lxml'),
        'PyMuPDF': _check_module('pymupdf', 'PyMuPDF'),
    }
    lo_path = render_docx.find_renderer(renderer)
    core_ok = python_ok and deps['python-docx']['installed'] and deps['lxml']['installed']
    render_ok = bool(lo_path) and deps['PyMuPDF']['installed']
    status = 'fail' if not core_ok else ('ok' if render_ok else 'partial')

    hints = []
    if not core_ok:
        hints.append('Install runtime dependencies with `pip install -e .` or `pip install -r requirements.txt` (Python 3.10+).')
    if not deps['PyMuPDF']['installed']:
        hints.append('Install PyMuPDF (`pip install PyMuPDF`) for `--png` page image rendering.')
    if not lo_path:
        hints.append(
            'LibreOffice not found (PDF/PNG rendering unavailable). Install: '
            'Windows `winget install TheDocumentFoundation.LibreOffice`, '
            'macOS `brew install --cask libreoffice`, '
            'Ubuntu/Debian `sudo apt-get install -y libreoffice-writer fonts-liberation`, '
            'or set VNDOCS_RENDERER.'
        )

    return {
        'status': status,
        'python': {'version': platform.python_version(), 'supported': python_ok},
        'dependencies': deps,
        'renderer': {
            'libreoffice': lo_path,
            'available': bool(lo_path),
            'word_win_available': bool(render_docx.find_word_win()),
        },
        'hints': hints,
    }


def run_doctor(argv=None):
    """Execute the `vndocs doctor` subcommand."""
    parser = argparse.ArgumentParser(prog='vndocs doctor', description=doctor_report.__doc__)
    parser.add_argument('--renderer', help='Explicit LibreOffice executable path to check')
    args = parser.parse_args(argv)
    report = doctor_report(args.renderer)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return {'ok': 0, 'partial': 2, 'fail': 1}[report['status']]


def main(argv=None):
    """Dispatch `vndocs` subcommands cleanly while preserving subcommand --help."""
    docx_package.ensure_utf8_stdio()
    args = list(sys.argv[1:] if argv is None else argv)
    if not args or args[0] in ('-h', '--help'):
        print(HELP_TEXT)
        return 0
    command, rest = args[0], args[1:]
    try:
        if command == 'doctor':
            return run_doctor(rest)
        handler = COMMANDS.get(command)
        if handler is None:
            print(f'Unknown command: {command}\n\n{HELP_TEXT}', file=sys.stderr)
            return 2
        result = handler(rest)
        return 0 if result is None else result
    except SystemExit as exc:
        return exc.code if isinstance(exc.code, int) else 1


if __name__ == '__main__':
    raise SystemExit(main())
