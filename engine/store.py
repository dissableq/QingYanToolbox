# -*- coding: utf-8 -*-
"""青雁批量工具箱 v3 — 本地持久化存储 (收藏夹 / 最近使用 / 参数方案)。

全部 JSON 落盘于 %APPDATA%\\QingYanBatchToolbox\\, 重启不丢失。
与 v2 完全兼容: 沿用同一数据目录与文件名, 老用户升级无感。

v3 新增:
- ParamStore 增加 theme(外观)/pipelines(命名流水线)/trash_days(回收站自动清理天数)。
- 所有写盘均为原子写(临时文件 + 替换), 避免意外中断损坏数据。
"""
from __future__ import annotations

import json
import os
import tempfile
from datetime import datetime

APP_DIR_NAME = 'QingYanBatchToolbox'


def data_dir() -> str:
    base = os.environ.get('APPDATA') or os.path.expanduser('~')
    d = os.path.join(base, APP_DIR_NAME)
    os.makedirs(d, exist_ok=True)
    return d


def _atomic_write(path: str, data) -> None:
    parent = os.path.dirname(os.path.abspath(path))
    os.makedirs(parent, exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix='.qingyan_', suffix='.json', dir=parent)
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as fh:
            json.dump(data, fh, ensure_ascii=False, indent=1)
            fh.flush()
            os.fsync(fh.fileno())
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)



class Favorites:
    """收藏夹: 收藏首页功能入口 (功能名, 目标页, 操作名), 重启可恢复。"""

    def __init__(self, path: str = None):
        self.path = path or os.path.join(data_dir(), 'favorites.json')

    def load(self) -> list:
        try:
            with open(self.path, encoding='utf-8') as fh:
                data = json.load(fh)
            if not isinstance(data, list):
                return []
            valid=[]
            for item in data:
                if not isinstance(item, dict):
                    continue
                name, page = item.get('name'), item.get('page')
                if not isinstance(name,str) or not name or not isinstance(page,str) or not page:
                    continue
                if not any(x['name']==name and x['page']==page for x in valid):
                    valid.append({'name':name,'page':page,'op':item.get('op') or name})
            return valid
        except Exception:
            return []

    def save(self, items: list) -> None:
        _atomic_write(self.path, items)

    def add(self, name: str, page: str, op: str | None) -> None:
        items = self.load()
        # 同名功能可能属于不同页面(如「查找替换」在 Word/文本等),
        # 以 (name, page) 为唯一键, 避免误判已收藏(v3 修复)
        if not any(x.get('name') == name and x.get('page') == page
                   for x in items):
            items.insert(0, {'name': name, 'page': page, 'op': op})
            self.save(items)

    def remove(self, name: str, page: str = None) -> None:
        items = self.load()
        if page is None:
            items = [x for x in items if x.get('name') != name]
        else:
            items = [x for x in items
                     if not (x.get('name') == name and x.get('page') == page)]
        self.save(items)

    def is_favorite(self, name: str, page: str = None) -> bool:
        if page is None:
            return any(x.get('name') == name for x in self.load())
        return any(x.get('name') == name and x.get('page') == page
                   for x in self.load())


def load_recent(records_dir: str, limit: int = 12) -> list:
    """从任务台账聚合「最近使用」: [(页面标题, 操作名, 最近时间, 次数)]。

    台账 options 里存有 op_name; 无操作名的专用页(模板/合并等)按 title 返回。
    """
    if not os.path.isdir(records_dir):
        return []
    agg: dict = {}
    try:
        for fn in os.listdir(records_dir):
            if not fn.endswith('.json'):
                continue
            try:
                with open(os.path.join(records_dir, fn), encoding='utf-8') as fh:
                    rec = json.load(fh)
            except Exception:
                continue
            if not isinstance(rec, dict):
                continue
            title = rec.get('title') or ''
            op = ''
            opts = rec.get('options')
            if isinstance(opts, dict):
                op = opts.get('op_name') or ''
            key = title + '\x00' + op
            t = rec.get('time') or ''
            cur = agg.get(key)
            if cur is None:
                agg[key] = [title, op, t, 1]
            else:
                cur[3] += 1
                if t > cur[2]:
                    cur[2] = t
    except Exception:
        return []
    rows = sorted(agg.values(), key=lambda r: r[2], reverse=True)[:limit]
    return [(r[0], r[1], r[2], r[3]) for r in rows]


class TrashManifest:
    """回收备份区清单: 记录 trash 文件名 → 原路径, 支持恢复。"""

    def __init__(self, path: str = None):
        self.path = path or os.path.join(data_dir(), 'trash_manifest.json')

    def load(self) -> dict:
        try:
            with open(self.path, encoding='utf-8') as fh:
                data = json.load(fh)
            return data if isinstance(data, dict) else {}
        except Exception:
            return {}

    def save(self, data: dict) -> None:
        _atomic_write(self.path, data)

    def record(self, trash_name: str, original_path: str) -> None:
        data = self.load()
        data[trash_name] = {
            'src': original_path,
            'time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        }
        self.save(data)

    def remove(self, trash_name: str) -> None:
        data = self.load()
        data.pop(trash_name, None)
        self.save(data)


class ParamStore:
    """参数方案存储: 自动记住每页每操作的上次参数 + 命名方案 + 窗口状态。

    v3: 增加外观主题(theme/accent/font_scale)、命名流水线(pipelines)、
        回收站自动清理天数(trash_days)。
    """

    def __init__(self, path: str = None):
        self.path = path or os.path.join(data_dir(), 'params.json')

    # ------------------------------------------------------------- internal
    def _load_all(self) -> dict:
        try:
            with open(self.path, encoding='utf-8') as fh:
                data = json.load(fh)
            if isinstance(data, dict):
                data.setdefault('auto', {})
                data.setdefault('presets', {})
                data.setdefault('ui', {})
                data.setdefault('pipelines', {})
                return data
        except Exception:
            pass
        return {'auto': {}, 'presets': {}, 'ui': {}, 'pipelines': {}}

    def _save_all(self, data: dict) -> None:
        _atomic_write(self.path, data)

    # ------------------------------------------------- auto (last params)
    def save_auto(self, page_kind: str, op: str, state: dict) -> None:
        data = self._load_all()
        auto = data.setdefault('auto', {})
        if not isinstance(auto, dict):
            auto = data['auto'] = {}
        auto.setdefault(page_kind, {})[op or '_page'] = state
        self._save_all(data)

    def load_auto(self, page_kind: str, op: str) -> dict:
        data = self._load_all()
        page = data.get('auto', {}).get(page_kind, {})
        if isinstance(page, dict):
            v = page.get(op or '_page')
            if isinstance(v, dict):
                return v
        return {}

    # ------------------------------------------------------------- presets
    def save_preset(self, name: str, page_title: str, op: str, state: dict) -> None:
        data = self._load_all()
        data.setdefault('presets', {})[name] = {
            'page': page_title,
            'op': op,
            'state': state,
            'time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        }
        self._save_all(data)

    def load_presets(self) -> dict:
        presets = self._load_all().get('presets')
        return presets if isinstance(presets, dict) else {}

    def delete_preset(self, name: str) -> None:
        data = self._load_all()
        data.get('presets', {}).pop(name, None)
        self._save_all(data)

    PRESET_FILE_KEY = 'qingyan_presets'

    def export_presets(self, path: str) -> int:
        presets = self.load_presets()
        with open(path, 'w', encoding='utf-8') as fh:
            json.dump({self.PRESET_FILE_KEY: presets}, fh,
                      ensure_ascii=False, indent=1)
        return len(presets)

    def import_presets(self, path: str) -> int:
        with open(path, encoding='utf-8') as fh:
            data = json.load(fh)
        items = None
        if isinstance(data, dict):
            # 导出用的是 qingyan_presets; 兼容早期/手写的 presets 键
            items = data.get(self.PRESET_FILE_KEY) or data.get('presets')
        if not isinstance(items, dict):
            return 0
        all_data = self._load_all()
        cur = all_data.setdefault('presets', {})
        cur.update(items)
        self._save_all(all_data)
        return len(items)

    # ------------------------------------------------------------------ ui
    def save_ui(self, geo: dict) -> None:
        data = self._load_all()
        data.setdefault('ui', {}).update(geo)
        self._save_all(data)

    def load_ui(self) -> dict:
        ui = self._load_all().get('ui')
        return ui if isinstance(ui, dict) else {}

    # -------------------------------------------------- v3: theme / extras
    def get_theme(self) -> dict:
        ui = self.load_ui()
        return {
            'mode': ui.get('theme_mode', 'light'),        # light | dark
            'accent': ui.get('accent', '#635BDF'),
            'font_scale': ui.get('font_scale', 1.0),
        }

    def set_theme(self, mode: str = None, accent: str = None,
                  font_scale: float = None) -> None:
        data = self._load_all()
        ui = data.setdefault('ui', {})
        if mode is not None:
            ui['theme_mode'] = mode
        if accent is not None:
            ui['accent'] = accent
        if font_scale is not None:
            ui['font_scale'] = font_scale
        self._save_all(data)

    # ------------------------------------------------------- v3: pipelines
    def save_pipeline(self, name: str, steps: list) -> None:
        data = self._load_all()
        data.setdefault('pipelines', {})[name] = {
            'steps': steps,
            'time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        }
        self._save_all(data)

    def load_pipelines(self) -> dict:
        ps = self._load_all().get('pipelines')
        return ps if isinstance(ps, dict) else {}

    def delete_pipeline(self, name: str) -> None:
        data = self._load_all()
        data.get('pipelines', {}).pop(name, None)
        self._save_all(data)


def is_first_run() -> bool:
    """首次启动检测: flags/first_run.json 不存在即视为首次。"""
    flags = os.path.join(data_dir(), 'flags')
    marker = os.path.join(flags, 'first_run.json')
    if os.path.exists(marker):
        return False
    os.makedirs(flags, exist_ok=True)
    try:
        with open(marker, 'w', encoding='utf-8') as fh:
            json.dump({'first_seen': datetime.now().strftime('%Y-%m-%d %H:%M:%S')}, fh)
    except Exception:
        pass
    return True
