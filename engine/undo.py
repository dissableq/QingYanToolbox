"""Undo in reverse order, retaining unfinished dependent actions for retry."""
import os
import shutil
from engine.safety import atomic_save


def undo_actions(actions):
    remaining = list(actions)
    completed = []
    errors = []
    while remaining:
        action = remaining[-1]
        try:
            if not isinstance(action, dict):
                raise ValueError('撤销记录格式异常')
            src, dst = action.get('src'), action.get('dst')
            if not src or not dst:
                raise ValueError('撤销记录缺少文件路径')
            if action.get('kind', 'move') == 'restore':
                from engine.backups import resolve_backup
                src = resolve_backup(src)
                if not os.path.isfile(src):
                    raise FileNotFoundError('备份不存在：' + src)
                if not os.path.isfile(dst):
                    raise FileNotFoundError('待还原的文件不存在：' + dst)
                atomic_save(lambda temp: shutil.copy2(src, temp), dst)
                try:
                    os.remove(src)
                    from engine.backups import forget
                    forget(src)
                except OSError:
                    pass
            else:
                if not os.path.exists(dst):
                    raise FileNotFoundError('待移回的文件不存在：' + dst)
                if os.path.exists(src):
                    # Windows case-only renames refer to the same existing file.
                    if os.path.normcase(os.path.abspath(src)) != os.path.normcase(os.path.abspath(dst)):
                        raise FileExistsError('原位置已有同名文件：' + src)
                    os.rename(dst, src)
                else:
                    os.makedirs(os.path.dirname(os.path.abspath(src)), exist_ok=True)
                    shutil.move(dst, src)
                from engine.backups import remap
                remap(dst, src)
            completed.append(remaining.pop())
        except Exception as exc:
            errors.append(str(exc))
            # Earlier operations may depend on this one. Preserve their order.
            break
    return completed, remaining, errors
