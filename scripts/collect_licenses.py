"""Collect installed runtime dependency notices into a distribution directory."""
from importlib.metadata import distribution
from pathlib import Path
import json
import re
import shutil
import sys

from packaging.requirements import Requirement

ROOT = Path(__file__).resolve().parent.parent


def collect(target: Path) -> None:
    target.mkdir(parents=True, exist_ok=True)
    queue = []
    for line in (ROOT / 'requirements.txt').read_text(encoding='utf-8').splitlines():
        if line.strip() and not line.startswith('#'):
            requirement = Requirement(line)
            if not requirement.marker or requirement.marker.evaluate({'extra': ''}):
                queue.append(requirement.name)
    seen = set()
    manifest = []
    while queue:
        name = queue.pop(0)
        key = re.sub(r'[-_.]+', '-', name).lower()
        if key in seen:
            continue
        seen.add(key)
        dist = distribution(name)
        notices = []
        for item in dist.files or []:
            filename = Path(str(item)).name
            if not filename.lower().startswith(('license', 'licence', 'copying', 'notice')):
                continue
            source = Path(dist.locate_file(item))
            if not source.is_file() or source.suffix.lower() in ('.py', '.pyc', '.dll', '.pyd'):
                continue
            destination = target / key / f'{len(notices) + 1:02d}-{filename}'
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, destination)
            notices.append(destination.relative_to(target).as_posix())
        manifest.append({
            'name': dist.metadata['Name'], 'version': dist.version,
            'license': dist.metadata.get('License-Expression') or dist.metadata.get('License'),
            'project_urls': dist.metadata.get_all('Project-URL') or [],
            'notices': notices,
        })
        for raw in dist.requires or []:
            requirement = Requirement(raw)
            if not requirement.marker or requirement.marker.evaluate({'extra': ''}):
                queue.append(requirement.name)
    for filename in ('LICENSE.txt', 'LICENSE', 'LICENCE.txt'):
        candidate = Path(sys.base_prefix) / filename
        if candidate.is_file():
            shutil.copy2(candidate, target / 'Python-LICENSE.txt')
            break
    (target / 'dependencies.json').write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'Collected notices for {len(manifest)} installed runtime packages')


if __name__ == '__main__':
    collect(Path(sys.argv[1]))
