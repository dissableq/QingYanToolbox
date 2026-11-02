"""Tracked, versioned backups beside source files; never delete untracked files."""
from pathlib import Path
from datetime import datetime
import hashlib
import json
import os
import re
import shutil
import threading
import uuid
from engine.store import data_dir, _atomic_write

SESSION = uuid.uuid4().hex
_lock = threading.RLock()


def _index():
    return Path(data_dir())/'backups.json'


def _load():
    p = _index()
    if not p.exists():return []
    rows = json.loads(p.read_text(encoding='utf-8'))
    if not isinstance(rows, list):raise ValueError('备份清单格式异常，请保留备份并检查用户数据目录')
    return rows


def _hash(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for block in iter(lambda:f.read(1024*1024), b''):h.update(block)
    return h.hexdigest()


def is_backup_path(path):
    p = Path(path)
    return '备份' in p.parts or p.name.lower().endswith('.qingyanbak') or p.name.startswith('.qingyan_')


def _no_links(path):
    return not any(p.is_symlink() or (hasattr(p, 'is_junction') and p.is_junction())
                   for p in (path, *path.parents))


def register(path, original='', kind='file', legacy_from=''):
    path = Path(path).absolute()
    with _lock:
        rows = _load()
        entry = dict(id=uuid.uuid4().hex, path=str(path), original=str(original),
                     kind=kind, session=SESSION, time=datetime.now().isoformat(timespec='seconds'),
                     bytes=path.stat().st_size, sha256=_hash(path),legacy_from=legacy_from)
        rows.append(entry);_atomic_write(str(_index()), rows)
    return entry


def _folder(directory):
    directory = Path(directory).absolute()
    folder = directory/'备份'
    if not _no_links(folder):raise ValueError('备份目录不能是链接或目录联接：'+str(folder))
    folder.mkdir(exist_ok=True)
    return folder


def migrate_legacy(directory, recursive=False):
    """Move only recognized legacy suffixes, and repair the existing undo paths."""
    directory = Path(directory).absolute()
    folders = [directory]
    if recursive:
        folders = [Path(dp) for dp, _, _ in walk_sources(directory)]
    moved = []
    with _lock:
        for parent in folders:
            if is_backup_path(parent) or not _no_links(parent):continue
            for old in parent.glob('*.qingyanbak'):
                if not old.is_file() or not _no_links(old):continue
                dst = _folder(parent)/old.name
                if dst.exists():dst = dst.with_name(uuid.uuid4().hex[:12]+'_'+dst.name)
                original = re.sub(r'(?:\.[0-9a-f]{12})?\.qingyanbak$', '', str(old))
                shutil.copy2(old, dst)
                try:
                    entry = register(dst, original, 'legacy',str(old))
                    manifest = Path(data_dir())/'last_undo.json'
                    if manifest.exists():
                        data = json.loads(manifest.read_text(encoding='utf-8'))
                        for action in data.get('actions', []):
                            if os.path.normcase(os.path.abspath(action.get('src',''))) == os.path.normcase(str(old)):action['src'] = str(dst)
                        _atomic_write(str(manifest), data)
                    old.unlink()
                except Exception:
                    # Keep both copies if any manifest step fails; never discard data.
                    raise
                moved.append(entry)
    return moved


def create_backup(path):
    path = Path(path).absolute()
    if is_backup_path(path):raise ValueError('备份文件不参与处理，请选择原文件')
    if not path.is_file() or not _no_links(path):raise ValueError('原文件不存在或为链接：'+str(path))
    with _lock:
        migrate_legacy(path.parent)
        folder = _folder(path.parent)
        stamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
        name = f'{path.stem[:100]}_{stamp}_{uuid.uuid4().hex[:8]}{path.suffix}'
        dst = folder/name
        shutil.copy2(path, dst)
        try:register(dst, str(path))
        except Exception:
            dst.unlink(missing_ok=True)
            raise
        return str(dst)


def create_journal(directory, task_id, data):
    dst = _folder(directory)/('名称与任务记录_'+task_id+'.json')
    _atomic_write(str(dst), data)
    register(dst, str(directory), 'journal')
    return str(dst)


def existing(current_session=False):
    with _lock:
        return [r for r in _load() if Path(r['path']).is_file()
                and (not current_session or r.get('session') == SESSION)]


def forget(path):
    with _lock:
        rows = [r for r in _load() if os.path.normcase(r['path']) != os.path.normcase(str(path))]
        _atomic_write(str(_index()), rows)


def resolve_backup(path):
    if Path(path).is_file():return path
    with _lock:
        for row in _load():
            if os.path.normcase(row.get('legacy_from',''))==os.path.normcase(path) and Path(row['path']).is_file():return row['path']
    return path


def without_cleaned_actions(actions,deleted_paths):
    """Restore paths are historical; map through later renames before matching."""
    deleted={os.path.normcase(os.path.abspath(p)) for p in deleted_paths};keep=[]
    for i,action in enumerate(actions):
        src=action.get('src','')
        if action.get('kind','move')=='restore':
            for later in actions[i+1:]:
                if later.get('kind','move')!='move':continue
                old,new=later.get('src',''),later.get('dst','')
                if src==old or src.startswith(old+os.sep):src=new+src[len(old):]
            if os.path.normcase(os.path.abspath(src)) in deleted:continue
        keep.append(action)
    return keep


def remap(old, new):
    """Keep tracked backup locations correct when a containing directory moves."""
    old, new = os.path.abspath(old), os.path.abspath(new)
    def mapped(value):
        if not value:return value
        norm, base = os.path.normcase(value), os.path.normcase(old)
        if norm == base:return new
        if norm.startswith(base+os.sep):return new+value[len(old):]
        return value
    with _lock:
        rows = _load();changed = False
        for row in rows:
            for key in ('path','original'):
                value = mapped(row.get(key,''))
                if value != row.get(key):row[key] = value;changed = True
        if changed:_atomic_write(str(_index()), rows)


def cleanup(ids):
    """Delete exact tracked bytes only. Leave modified backups and unrelated files."""
    removed, errors, deleted_paths = [], [], []
    with _lock:
        rows = _load();keep = []
        for row in rows:
            if row['id'] not in ids:keep.append(row);continue
            path = Path(row['path'])
            try:
                if path.parent.name != '备份' or not _no_links(path):
                    raise ValueError('位置异常，已保留')
                if path.exists():
                    if _hash(path) != row['sha256']:raise ValueError('备份内容已被修改，已保留')
                    path.unlink()
                removed.append(row);deleted_paths.append(str(path))
                if row.get('legacy_from'):deleted_paths.append(row['legacy_from'])
                try:path.parent.rmdir()
                except OSError:pass
            except Exception as exc:
                errors.append(str(path)+'：'+str(exc));keep.append(row)
        _atomic_write(str(_index()), keep)
        manifest = Path(data_dir())/'last_undo.json'
        if manifest.exists():
            data = json.loads(manifest.read_text(encoding='utf-8'))
            data['actions'] = without_cleaned_actions(data.get('actions',[]),deleted_paths)
            _atomic_write(str(manifest), data)
    return removed, errors


def walk_sources(root, **kwargs):
    for dp, dirs, files in os.walk(root, **kwargs):
        if is_backup_path(dp):continue
        dirs[:] = [d for d in dirs if d != '备份' and not is_backup_path(d)
                   and _no_links(Path(dp)/d)]
        files[:] = [f for f in files if not is_backup_path(f) and _no_links(Path(dp)/f)]
        yield dp, dirs, files
