# -*- coding: utf-8 -*-
"""青雁批量工具箱 v3 — 批量任务核心引擎。

设计要点(承袭 v2, v3 增强):
- 纯离线: 不发起任何网络请求, 不依赖任何在线服务。
- 后台线程执行批量任务, 主线程通过 Qt 信号接收进度。
- 每个文件独立计成功/失败/跳过, 失败项带原因; 支持取消。
- v3: 每文件记录耗时; progress 附带已用秒(供 ETA/速度估算);
      FileResult 可携带 undo 动作, 任务结束写撤销清单(供台账页撤销)。
- 任务结束自动写台账 JSON, 可导出 CSV/Excel。
"""
from __future__ import annotations

import json
import os
import time
from dataclasses import dataclass, field
from datetime import datetime
from typing import Callable, List, Optional

from PySide6.QtCore import QThread, Signal
from engine.store import _atomic_write
from engine.backups import is_backup_path, walk_sources, remap
from uuid import uuid4


@dataclass
class FileResult:
    """单个文件的处理结果。"""
    path: str
    status: str = 'ok'          # ok | fail | skipped
    detail: str = ''
    output: str = ''
    duration_sec: float = 0.0                       # v3: 单文件耗时
    undo: List[dict] = field(default_factory=list)  # v3: [{src, dst}] 撤销动作
    outputs: List[str] = field(default_factory=list)


@dataclass
class TaskSpec:
    """一次批量任务的参数包。"""
    kind: str
    title: str
    files: List[str] = field(default_factory=list)
    options: dict = field(default_factory=dict)
    preview_only: bool = False


class BatchWorker(QThread):
    """批量执行线程。file_done 逐文件推送, finished_all 汇总推送。

    v3 增强:
    - progress(done, total, name, elapsed_sec): 供 ETA / 速度估算。
    - 取消后剩余文件不再处理; 已处理结果仍写台账。
    - 汇总撤销动作写入 undo_manifest_path, 供「撤销上次操作」。
    """

    progress = Signal(int, int, str, float)   # 已完成数, 总数, 当前文件名, 已用秒
    file_done = Signal(object)                # FileResult
    finished_all = Signal(list, str)          # 全部 FileResult, task_id
    cancelled = Signal()

    def __init__(self, spec: TaskSpec, processor: Callable[[str, dict], FileResult],
                 records_dir: str, undo_manifest_path: Optional[str] = None):
        super().__init__()
        self.spec = spec
        self.processor = processor
        self.records_dir = records_dir
        self.undo_manifest_path = undo_manifest_path
        self._cancel = False
        self.results: List[FileResult] = []
        self.undo_actions: List[dict] = []

    def cancel(self):
        self._cancel = True

    # ------------------------------------------------------------------ run
    def run(self):
        t0 = time.time()
        files = list(self.spec.files)
        total = len(files)
        done = ok = fail = skipped = 0
        now = datetime.now()
        task_id = now.strftime('%Y%m%d_%H%M%S_') + uuid4().hex[:10]

        for f in files:
            if self._cancel:
                break
            name = os.path.basename(f)
            t1 = time.time()
            # 先发一次进度: 大文件(几百页 PDF)处理时长的分钟级里,
            # 界面能显示正在处理谁, 而不是一直停在上一个文件的进度上
            self.progress.emit(done, total, name, time.time() - t0)
            try:
                r = self.processor(f, dict(self.spec.options))
                for action in getattr(r, 'undo', []):
                    if action.get('kind', 'move') == 'move':
                        remap(action['src'], action['dst'])
                if not isinstance(r, FileResult):
                    r = FileResult(path=str(f), status='ok', detail=str(r))
            except Exception as e:  # 单文件异常不影响后续
                r = FileResult(path=str(f), status='fail',
                               detail=type(e).__name__ + ': ' + str(e))
            if not r.path:
                r.path = str(f)
            r.duration_sec = round(time.time() - t1, 2)
            if r.undo:
                self.undo_actions.extend(r.undo)
            self.results.append(r)
            if r.status == 'ok':
                ok += 1
            elif r.status == 'fail':
                fail += 1
            else:
                skipped += 1
            done += 1
            self.file_done.emit(r)
            self.progress.emit(done, total, name, time.time() - t0)

        duration = time.time() - t0
        self._write_record(task_id, ok, fail, skipped, duration)
        if self.undo_manifest_path:
            self._write_undo_manifest(task_id)
        if self._cancel:
            self.cancelled.emit()
        else:
            self.finished_all.emit(self.results, task_id)

    # --------------------------------------------------------------- records
    def _write_record(self, task_id: str, ok: int, fail: int, skip: int,
                      duration: float):
        if not self.records_dir:
            return
        try:
            os.makedirs(self.records_dir, exist_ok=True)
            now = datetime.now()
            rec = {
                'task_id': task_id,
                'time': now.strftime('%Y-%m-%d %H:%M:%S'),
                'kind': self.spec.kind,
                'title': self.spec.title,
                'options': _jsonable(self.spec.options),
                'total': len(self.results),
                'ok': ok,
                'fail': fail,
                'skipped': skip,
                'duration_sec': round(duration, 2),
                'preview_only': self.spec.preview_only,
                'files': [
                    {'path': r.path, 'status': r.status,
                     'detail': (r.detail or ''),
                     'output': (r.output or ''), 'outputs': r.outputs}
                    for r in self.results
                ],
            }
            path = os.path.join(self.records_dir, task_id + '.json')
            _atomic_write(path, rec)
        except Exception:
            pass

    def _write_undo_manifest(self, task_id: str):
        try:
            manifest = {'task_id': task_id, 'actions': self.undo_actions}
            _atomic_write(self.undo_manifest_path, manifest)
        except Exception:
            pass


def _jsonable(o):
    """把 options 里的不可序列化对象转成字符串, 保证台账可写。"""
    try:
        json.dumps(o)
        return o
    except Exception:
        if isinstance(o, dict):
            return {str(k): _jsonable(v) for k, v in o.items()}
        if isinstance(o, (list, tuple, set)):
            return [_jsonable(x) for x in o]
        return str(o)


def collect_files(root: str, recursive: bool = False,
                  ext_filter: Optional[List[str]] = None) -> List[str]:
    """收集目录下文件; ext_filter 形如 ['.jpg','.png'](不区分大小写)。

    root 为单个文件时直接返回 [root]; 目录不存在时返回 []。
    """
    if is_backup_path(root):
        return []
    if os.path.isfile(root):
        return [root]
    exts = None
    if ext_filter:
        exts = [e.lower().lstrip('*.') for e in ext_filter
                if str(e).strip()]
    out: List[str] = []
    if recursive:
        for dirpath, dirs, files in walk_sources(root):
            # 跳过工具自己生成的产物目录。转换结果默认落在源文件旁的
            # _转换结果, 递归收集时若把它当输入, 第二次运行就会把上次的
            # 产物再转换一遍, 目录越滚越大(还可能结果转结果)。
            dirs[:] = [d for d in dirs
                       if d != '_转换结果' and not d.startswith('qingyan_pipe_')
                       and not os.path.islink(os.path.join(dirpath, d))]
            for fn in sorted(files):
                if fn.endswith('.qingyanbak') or fn.startswith('.qingyan_'):
                    continue
                ext = os.path.splitext(fn)[1].lower().lstrip('.')
                if not exts or ext in exts:
                    out.append(os.path.join(dirpath, fn))
    else:
        try:
            names = sorted(os.listdir(root))
        except OSError:
            return out
        for fn in names:
            if fn.endswith('.qingyanbak') or fn.startswith('.qingyan_'):
                continue
            p = os.path.join(root, fn)
            if os.path.isfile(p):
                ext = os.path.splitext(fn)[1].lower().lstrip('.')
                if not exts or ext in exts:
                    out.append(p)
    return out
