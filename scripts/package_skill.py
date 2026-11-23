"""Package the self-contained skill, excluding research datasets and local runtimes."""
import argparse
import hashlib
import json
import re
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED

try:
    from .docx_package import ensure_utf8_stdio
except ImportError:
    from docx_package import ensure_utf8_stdio


def _read_skill_name(skill_md_path):
    """Extract and validate the skill name from SKILL.md YAML frontmatter using stdlib only."""
    if not skill_md_path.is_file():
        raise ValueError(f'Missing SKILL.md: {skill_md_path}')
    content = skill_md_path.read_text(encoding='utf-8-sig')
    fm_match = re.match(r'^---\r?\n(.*?)\r?\n---', content, re.S)
    if not fm_match:
        raise ValueError('SKILL.md is missing YAML frontmatter')
    name_match = re.search(r'^\s*name:\s*[\'"]?([^\'"\r\n#]+?)[\'"]?\s*$', fm_match.group(1), re.M)
    if not name_match:
        raise ValueError('SKILL.md frontmatter is missing `name`')
    skill_name = name_match.group(1).strip()
    if not re.fullmatch(r'[a-z0-9-]+', skill_name) or len(skill_name) > 64:
        raise ValueError(f'Invalid skill name in SKILL.md frontmatter: {skill_name!r}')
    return skill_name


def package_skill(root, destination):
    root, destination = Path(root).resolve(), Path(destination).resolve()
    skill_name = _read_skill_name(root/'SKILL.md')
    candidates = [root/'SKILL.md', root/'requirements.txt']
    candidates += sorted((root/'references').glob('*.md'))
    candidates += sorted((root/'scripts').glob('*.py'))
    candidates += sorted((root/'assets/rules').glob('*.json'))
    if (root/'assets/schemas').is_dir():
        candidates += sorted((root/'assets/schemas').glob('*.json'))
    manifest_path = root/'assets/templates/manifest.json'
    if not manifest_path.is_file():
        raise ValueError('Missing verified starter template manifest')
    catalog = json.loads(manifest_path.read_text(encoding='utf-8-sig'))
    candidates.append(manifest_path)
    for item in catalog:
        if (not re.fullmatch('[a-z0-9-]+', item['id'])
                or item['template'] != item['id'] + '.docx'):
            raise ValueError('Invalid template identity or filename')
        if item['origin'] != 'project-authored' or item.get('visual_trial') != 'reviewed':
            raise ValueError('Only project-authored, reviewed starters can enter this package')
        path = root/'assets/templates'/item['template']
        if path.is_symlink() or not path.resolve().is_relative_to((root/'assets/templates').resolve()):
            raise ValueError('Template path escapes its asset directory')
        if hashlib.sha256(path.read_bytes()).hexdigest() != item['sha256']:
            raise ValueError('Template changed since verification')
        candidates += [path, root/'assets/templates'/(item['id']+'.rules.json'),
                       root/'assets/templates'/(item['id']+'.values.json')]
    entries = {}
    for path in candidates:
        resolved = path.resolve()
        if not resolved.is_relative_to(root) or path.is_symlink() or not path.is_file():
            raise ValueError(f'Invalid package file: {path}')
        entries[f'{skill_name}/' + path.relative_to(root).as_posix()] = path.read_bytes()
    destination.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(destination, 'w', compression=ZIP_DEFLATED) as archive:
        for name, content in sorted(entries.items()):
            info = ZipInfo(name, date_time=(2026, 10, 13, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, content)
    with ZipFile(destination) as archive:
        if archive.testzip() is not None:
            raise ValueError('Package CRC check failed')
        if set(archive.namelist()) != set(entries):
            raise ValueError('Package member mismatch')
    return {'skill_name': skill_name, 'output': str(destination), 'files': len(entries),
            'sha256': hashlib.sha256(destination.read_bytes()).hexdigest(),
            'excluded': ['references/source-documents', 'third-party research templates',
                         '.git', '.codex', '.test-deps', '.tools', 'artifacts', 'tests']}


def main(argv=None):
    ensure_utf8_stdio()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root', default=str(Path(__file__).resolve().parents[1]))
    parser.add_argument('--output', default='dist/vietnamese-docx.zip')
    args = parser.parse_args(argv)
    print(json.dumps(package_skill(args.root, args.output), ensure_ascii=False, indent=2))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
