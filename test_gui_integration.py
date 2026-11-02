# -*- coding: utf-8 -*-
"""GUI 集成测试: 离屏构建 WordPage, 跑一次真实批量替换任务, 验证
current_files → worker → on_file_done → on_finished → 台账 落盘。"""
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), r'_itest')
if os.path.isdir(BASE):
    shutil.rmtree(BASE, ignore_errors=True)
os.makedirs(os.path.join(BASE, 'docs'))
for i in range(5):
    with open(os.path.join(BASE, 'docs', f'f{i}.txt'), 'w', encoding='utf-8') as f:
        f.write(f'文档{i} 青雁测试 内容ABC\n第二行\n')

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import QEventLoop, QTimer

app = QApplication([])
import app.main as M

win = M.MainWindow()
page = None
for p in win.pages:
    if getattr(p, 'kind', '') == 'text':
        page = p
        break
assert page is not None, 'TextPage not found'

page.file_edit.setText(os.path.join(BASE, 'docs'))
page.ext_edit.setText('')
page._widgets['out_dir'].setText('')  # 清空自动记忆的输出目录, 就地修改
page._widgets['op'].setCurrentText('查找替换')
page._widgets['find'].setText('青雁')
page._widgets['replace'].setText('大雁')

files = page.current_files()
assert len(files) == 5, f'expected 5 files, got {len(files)}'
print('current_files OK:', len(files))

opts = page.collect_options()
assert opts.get('op') == 'replace' and opts.get('find') == '青雁'
proc = page.processor()
assert proc is not None
print('collect_options + processor OK')

# 跑 worker(同步等待)
from engine.core import TaskSpec
spec = TaskSpec(kind=page.kind, title=page.title, files=files, options=opts)
page._start_worker(spec, proc)

loop = QEventLoop()
page.worker.finished_all.connect(lambda *_: loop.quit())
page.worker.cancelled.connect(loop.quit)
QTimer.singleShot(30000, loop.quit)
loop.exec()

rows = page.result_table.rowCount()
print('result rows:', rows)
assert rows == 5, 'result rows != 5'
summary = page.summary_label.text()
print('summary:', summary)
assert '成功 5' in summary

records = os.listdir(win.RECORDS_DIR)
print('records:', len(records))
assert records, 'no records written'
import json
with open(os.path.join(win.RECORDS_DIR, records[-1]), encoding='utf-8') as fh:
    rec = json.load(fh)
print('last record:', rec.get('title'), 'ok =', rec.get('ok'))
assert rec.get('ok') == 5

# 验证文件真的被替换
with open(os.path.join(BASE, 'docs', 'f0.txt'), encoding='utf-8') as f:
    content = f.read()
assert '大雁测试' in content and '青雁' not in content
print('file content replaced OK')

# 台账页统计
rp = None
for p in win.pages:
    if getattr(p, 'title', '') == '任务台账':
        rp = p
rp.refresh()
print('records table rows:', rp.table.rowCount())
assert rp.table.rowCount() >= 1

shutil.rmtree(BASE, ignore_errors=True)
print('\nGUI INTEGRATION TEST: ALL PASS')
