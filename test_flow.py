# -*- coding: utf-8 -*-
"""GUI 流程测试(v3 修复回归): FilesPage 清理/打包/删除 + 流水线全链路。"""
import os
import shutil
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), r'_flowtest')
if os.path.isdir(BASE):
    shutil.rmtree(BASE, ignore_errors=True)
DOCS = os.path.join(BASE, 'docs')
os.makedirs(DOCS)
# 冗余命名样本
for nm in ('报告_副本1.txt', '报告_最终版.txt', '报告final.txt', '报告(1).txt',
           '正常文件.txt'):
    open(os.path.join(DOCS, nm), 'w', encoding='utf-8').write('x' * 10)

ok = fail = 0


def wait_page(page, timeout_ms=20000):
    """等页面当前 worker 跑完。

    _run_dir_level 已改为后台执行(BatchWorker), 断言前必须等线程结束,
    否则读到的是尚未产生结果的状态。
    """
    w = getattr(page, 'worker', None)
    if w is None or not w.isRunning():
        return
    from PySide6.QtCore import QEventLoop, QTimer
    loop = QEventLoop()
    w.finished_all.connect(lambda *_: loop.quit())
    try:
        w.cancelled.connect(loop.quit)
    except Exception:
        pass
    QTimer.singleShot(timeout_ms, loop.quit)
    loop.exec()


def check(name, cond, detail=''):
    global ok, fail
    if cond:
        ok += 1
        print('PASS', name)
    else:
        fail += 1
        print('FAIL', name, '::', detail)


os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
from PySide6.QtWidgets import QApplication, QDialog

# 测试桩: 模态对话框立即返回(自动化无法点击)
QDialog.exec = lambda self: QDialog.Accepted

app = QApplication([])
import app.main as M

win = M.MainWindow()
win.close()

files_page = None
for p in win.pages:
    if getattr(p, 'kind', '') == 'files':
        files_page = p
files_page.CONFIRM_OPS = {}   # 测试中跳过确认框

# ---------------------------------------------------------- 冗余命名清理
files_page.file_edit.setText(DOCS)
files_page._widgets['op'].setCurrentText('冗余命名清理')
opts = files_page.collect_options()
plan = []
try:
    plan = __import__('engine.ops', fromlist=['build_rename_plan']) \
        .build_rename_plan(files_page.current_files(),
                           {'mode': 'cleanup',
                            'cleanup_patterns': __import__(
                                'engine.ops',
                                fromlist=['CLEANUP_PATTERNS'])
                            .CLEANUP_PATTERNS})
except Exception as e:
    print('plan error:', e)
files_page._plan_cache = plan
opts['_rename_plan'] = plan
files_page._run_dir_level(opts, files=[p[0] for p in plan])
wait_page(files_page)
names_after = sorted(os.listdir(DOCS))
print('after cleanup:', names_after)
check('cleanup plan passed & applied',
      '报告.txt' in names_after and '报告_副本1.txt' not in names_after
      and '正常文件.txt' in names_after, str(names_after))
check('cleanup result rows', files_page.result_table.rowCount() >= 4,
      str(files_page.result_table.rowCount()))

# ------------------------------------------------------------- 打包 zip
files_page.file_edit.setText(DOCS)
files_page._widgets['op'].setCurrentText('打包为 zip')
opts = files_page.collect_options()
opts['_src_paths'] = [DOCS]
files_page._run_dir_level(opts, files=[DOCS], batch=True)
wait_page(files_page)
zip_path = os.path.join(BASE, 'docs.zip')   # 打包输出在所选文件夹旁边
check('pack multi -> one zip', os.path.exists(zip_path),
      str(os.listdir(BASE)))
if os.path.exists(zip_path):
    with zipfile.ZipFile(zip_path) as zf:
        n_in = len(zf.namelist())
    check('zip contains files', n_in >= 5, str(n_in))

# -------------------------------------------------- 删除文件(per-file 语义)
del_dir = os.path.join(BASE, 'deltest')
os.makedirs(del_dir)
for i in range(3):
    open(os.path.join(del_dir, f'd{i}.txt'), 'w').write('x')
files_page.file_edit.setText(del_dir)
files_page._widgets['op'].setCurrentText('删除文件(进回收备份区)')
from engine.core import TaskSpec
opts = files_page.collect_options()
from engine.core import collect_files
targets = collect_files(del_dir, False)
spec = TaskSpec(kind='files', title='删除文件测试', files=targets, options=opts)
files_page._start_worker(spec, files_page.processor())

from PySide6.QtCore import QEventLoop, QTimer
loop = QEventLoop()
files_page.worker.finished_all.connect(lambda *_: loop.quit())
QTimer.singleShot(20000, loop.quit)
loop.exec()
check('delete per-file all ok',
      all(r.status == 'ok' for r in files_page.worker.results),
      str([(r.status, r.detail) for r in files_page.worker.results]))
check('deleted files gone', not os.path.exists(
    os.path.join(del_dir, 'd0.txt')), str(os.listdir(del_dir)))
check('files moved to trash', files_page.worker.undo_actions
      and len(files_page.worker.undo_actions) == 3,
      str(files_page.worker.undo_actions))
# 撤销删除(restore 回原位)
import shutil as _sh
n = 0
for a in files_page.worker.undo_actions:
    if os.path.exists(a['dst']) and not os.path.exists(a['src']):
        _sh.move(a['dst'], a['src'])
        n += 1
check('undo restores deleted', n == 3 and os.path.exists(
    os.path.join(del_dir, 'd0.txt')), str(os.listdir(del_dir)))

# ------------------------------------------------------------- 流水线
pipe_page = None
text_page = None
for p in win.pages:
    if getattr(p, 'title', '') == '任务流水线':
        pipe_page = p
    if getattr(p, 'kind', '') == 'text':
        text_page = p

src_dir = os.path.join(BASE, 'pipe_src')
os.makedirs(src_dir)
for i in range(4):
    # 必须含真正的空行: 第二步是"删除空行", 文件没有空行时该步会
    # skipped 且无输出, 流水线按设计提前结束(旧实现把末尾换行误判成
    # 空行才总是有输出, 那是 bug)。
    open(os.path.join(src_dir, f'p{i}.txt'), 'w', encoding='utf-8').write(
        f'内容{i} ALPHA beta\n\n第二行 ALPHA\n')

step1 = {'page': '文本工具', 'op': '查找替换', 'page_obj': text_page,
         'options': {'op': 'replace', 'op_name': '查找替换',
                     'find': 'ALPHA', 'replace': 'X',
                     'file_edit': src_dir,
                     'out_dir': os.path.join(BASE, 'pipe1')}}
step2 = {'page': '文本工具', 'op': '删除空行', 'page_obj': text_page,
         'options': {'op': 'del_blank', 'op_name': '删除空行',
                     'out_dir': os.path.join(BASE, 'pipe2')}}
for _s in (step1, step2):
    try:
        _s['page_obj'].restore_state(_s['options'])
        _s['ui_state'] = _s['page_obj'].capture_state()
    except Exception:
        _s['ui_state'] = {}
pipe_page._steps = [step1, step2]
pipe_page.steps_list.clear()
pipe_page.steps_list.addItem('1. [文本工具] 查找替换')
pipe_page.steps_list.addItem('2. [文本工具] 删除空行')
pipe_page.run_pipeline()
from PySide6.QtCore import QEventLoop, QTimer
_pipeline_loop = QEventLoop()
_pipeline_poll = QTimer()
_pipeline_poll.timeout.connect(lambda: _pipeline_loop.quit() if pipe_page.run_btn.isEnabled() else None)
_pipeline_poll.start(20)
QTimer.singleShot(30000, _pipeline_loop.quit)
_pipeline_loop.exec()
_pipeline_poll.stop()
status = pipe_page.status.text()
print('pipeline status:', status)
check('pipeline completed', '已完成' in status and '失败 0' in status, status)
check('step1 output replaced', any('X' in open(
    os.path.join(BASE, 'pipe1', f), encoding='utf-8').read()
    for f in os.listdir(os.path.join(BASE, 'pipe1'))), '')
check('step2 chained input', len(os.listdir(os.path.join(BASE, 'pipe2'))) == 4,
      str(os.listdir(os.path.join(BASE, 'pipe2'))))

# 保存方案 -> 载入 -> 步骤重建
from engine.store import ParamStore
ParamStore().save_pipeline('测试链', [
    {'page': '文本工具', 'op': '查找替换',
     'options': {'op': 'replace', 'find': 'A', 'out_dir': ''}},
    {'page': '文本工具', 'op': '删除空行', 'options': {'op': 'del_blank'}},
])
pipe_page._reload_pipelines()
pipe_page.pipe_combo.setCurrentText('测试链')
pipe_page._load_pipeline()
check('pipeline load rebuilds steps',
      len(pipe_page._steps) == 2
      and all(s.get('page_obj') is not None for s in pipe_page._steps),
      str(len(pipe_page._steps)))
check('pipeline load keeps options',
      pipe_page._steps[0]['options'].get('find') == 'A',
      str(pipe_page._steps[0]['options']))

# add_step 页面归还验证(模拟: 借出后归还, stack 索引不变)
idx_before = win.stack.indexOf(text_page)
dlg_result = pipe_page._show_page_in_dialog(text_page, '测试归还')
idx_after = win.stack.indexOf(text_page)
check('page returned to stack same index', idx_before == idx_after != -1,
      f'{idx_before} -> {idx_after}')

shutil.rmtree(BASE, ignore_errors=True)
print(f'\nRESULT: {ok} passed, {fail} failed')
sys.exit(1 if fail else 0)
