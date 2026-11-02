"""Version checks and reproducible Windows release packaging from a Git commit."""
from pathlib import Path
import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import zipfile

ROOT = Path(__file__).resolve().parent.parent
APP = '青雁批量工具箱v3'
VERSION_FILES = {'app/main.py': 'APP_VERSION', 'installer_entry.py': 'VERSION',
                 'make_installer.py': 'VERSION'}


def run(arguments, **kwargs):
    environment = os.environ.copy()
    environment.update(kwargs.pop('env', {}))
    environment.update(PYTHONUTF8='1', PYTHONIOENCODING='utf-8')
    return subprocess.run(arguments, cwd=ROOT, check=True, env=environment, **kwargs)


def version(tag=None):
    versions = []
    for filename, variable in VERSION_FILES.items():
        matches = re.findall(rf"^{variable} = '([0-9]+\.[0-9]+\.[0-9]+)'$",
                             (ROOT / filename).read_text(encoding='utf-8'), re.M)
        if len(matches) != 1:
            raise ValueError(f'Cannot read version from {filename}')
        versions.append(matches[0])
    if len(set(versions)) != 1:
        raise ValueError(f'Product versions disagree: {versions}')
    value = versions[0]
    info = (ROOT / 'version_info.py').read_text(encoding='utf-8')
    numbers = tuple(int(part) for part in value.split('.')) + (0,)
    for key in ('filevers', 'prodvers'):
        match = re.search(rf'{key}=\(([^)]+)\)', info)
        if not match or tuple(int(x.strip()) for x in match[1].split(',')) != numbers:
            raise ValueError(f'version_info.py {key} disagrees with {value}')
    for key in ('FileVersion', 'ProductVersion'):
        if f"StringStruct('{key}', '{value}.0')" not in info:
            raise ValueError(f'version_info.py {key} disagrees with {value}')
    if tag and tag != f'v{value}':
        raise ValueError(f'Tag {tag} disagrees with product version v{value}')
    return value


def bump(new):
    old = version()
    if not re.fullmatch(r'(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)', new):
        raise ValueError('Use a numeric X.Y.Z version')
    if tuple(map(int, new.split('.'))) <= tuple(map(int, old.split('.'))):
        raise ValueError('A new release must increase the version')
    updates = {}
    for filename, variable in VERSION_FILES.items():
        path = ROOT / filename
        updates[path] = path.read_text(encoding='utf-8').replace(
            f"{variable} = '{old}'", f"{variable} = '{new}'", 1)
    path = ROOT / 'version_info.py'
    text = path.read_text(encoding='utf-8')
    new_tuple = ', '.join(new.split('.') + ['0'])
    for key in ('filevers', 'prodvers'):
        text = re.sub(rf'{key}=\([^)]+\)', f'{key}=({new_tuple})', text)
    text = text.replace(f"'{old}.0'", f"'{new}.0'")
    updates[path] = text
    readme = ROOT / 'README.md'
    updates[readme] = readme.read_text(encoding='utf-8').replace(f'**v{old}**', f'**v{new}**')
    for path, text in updates.items():
        path.write_text(text, encoding='utf-8', newline='\n')
    print(f'Updated v{old} -> v{version()}; update CHANGELOG and release notes next')


def build(tag=None):
    value = version(tag)
    notes = ROOT / 'docs' / 'releases' / f'v{value}.md'
    if not notes.is_file():
        raise ValueError(f'Missing release notes: {notes.relative_to(ROOT)}')
    status = run(['git', 'status', '--porcelain'], capture_output=True).stdout
    if status.strip():
        raise ValueError('Commit or remove pending changes before building a release')
    head = run(['git', 'rev-parse', 'HEAD'], capture_output=True, text=True).stdout.strip()
    run([sys.executable, '-m', 'PyInstaller', '--noconfirm', '--clean',
         '--distpath', 'dist', '--workpath', 'build/pyi', f'{APP}.spec'])
    dist = ROOT / 'dist' / APP
    for name in ('LICENSE', 'THIRD_PARTY_NOTICES.md', 'README.md'):
        shutil.copy2(ROOT / name, dist / name)
    run([sys.executable, 'scripts/collect_licenses.py', str(dist / 'licenses')])
    (dist / 'source-version.json').write_text(
        json.dumps({'version': value, 'commit': head}, indent=2), encoding='utf-8')
    smoke = ROOT / '.verification_data' / 'release-build'
    smoke.mkdir(parents=True, exist_ok=True)
    report = smoke / 'smoke.json'
    if report.exists():
        report.unlink()
    env = os.environ.copy()
    env.update(APPDATA=str(smoke / 'data'), LOCALAPPDATA=str(smoke / 'local'),
               QINGYAN_SMOKE_REPORT=str(report), QT_QPA_PLATFORM='offscreen', PYTHONUTF8='1')
    run([str(dist / f'{APP}.exe'), '--smoke'], env=env, timeout=180)
    if not report.is_file() or json.loads(report.read_text(encoding='utf-8'))['failed']:
        raise ValueError('Packaged executable smoke check failed')
    run([sys.executable, 'make_installer.py', '--dist', str(dist)])
    output = ROOT / 'release'
    output.mkdir(exist_ok=True)
    setup = output / f'QingYanToolbox-{value}-setup.exe'
    shutil.copy2(ROOT / '安装包' / f'{APP}_v{value}_安装包.exe', setup)
    portable = output / f'QingYanToolbox-{value}-windows-x64.zip'
    with zipfile.ZipFile(portable, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        for path in sorted(dist.rglob('*')):
            if path.is_file():
                archive.write(path, f'QingYanToolbox-{value}/' + path.relative_to(dist).as_posix())
    source = output / f'QingYanToolbox-{value}-source.zip'
    run(['git', 'archive', '--format=zip', f'--prefix=QingYanToolbox-{value}/',
         '-o', str(source), head])
    hashes = []
    for path in (setup, portable, source):
        with path.open('rb') as stream:
            digest = hashlib.file_digest(stream, 'sha256').hexdigest()
        hashes.append(f'{digest}  {path.name}')
    (output / 'SHA256SUMS.txt').write_text('\n'.join(hashes) + '\n', encoding='utf-8')
    print(f'Release v{value} from commit {head} ready in {output}')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='command', required=True)
    sub.add_parser('check').add_argument('--tag')
    sub.add_parser('build').add_argument('--tag')
    sub.add_parser('bump').add_argument('version')
    args = parser.parse_args()
    if args.command == 'bump':
        bump(args.version)
    elif args.command == 'build':
        build(args.tag)
    else:
        print(f'Version OK: v{version(args.tag)}')


if __name__ == '__main__':
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            stream.reconfigure(encoding='utf-8')
    main()
