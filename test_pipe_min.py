# -*- coding: utf-8 -*-
"""最小复现: 定位 run_pipeline 卡点。"""
import os
import shutil
import sys
import faulthandler

faulthandler.dump_traceback_later(25, exit=True)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), r'_pipetest')
shutil.rmtree(BASE, ignore_errors=True)
src_dir = os.path.join(BASE, 'src')
os.makedirs(src_dir)
for i in range(4):
    open(os.path.join(src_dir, f'p{i}.txt'), 'w', encoding='utf-8').write(
        f'内容{i} ALPHA beta\n第二行 ALPHA\n')

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
from PySide6.QtWidgets import QApplication

app = QApplication([])
import app.main as M

print('building window...', flush=True)
win = M.MainWindow()
win.close()
print('window built', flush=True)

text_page = None
pipe_page = None
for p in win.pages:
    if getattr(p, 'kind', '') == 'text':
        text_page = p
    if getattr(p, 'title', '') == '任务流水线':
        pipe_page = p

print('step restore...', flush=True)
step1 = {'page': '文本工具', 'op': '查找替换', 'page_obj': text_page,
         'options': {'op': 'replace', 'op_name': '查找替换',
                     'find': 'ALPHA', 'replace': 'X',
                     'file_edit': src_dir,
                     'out_dir': os.path.join(BASE, 'pipe1')}}
try:
    text_page.restore_state(step1['options'])
    print('restored; op=', text_page._widgets['op'].currentText(),
          'files=', len(text_page.current_files()), flush=True)
except Exception as e:
    print('restore error:', e, flush=True)

pipe_page._steps = [step1]
pipe_page.steps_list.clear()
pipe_page.steps_list.addItem('1. [文本工具] 查找替换')
print('running pipeline...', flush=True)
pipe_page.run_pipeline()
from PySide6.QtCore import QEventLoop, QTimer
_pipeline_loop = QEventLoop()
_pipeline_poll = QTimer()
_pipeline_poll.timeout.connect(lambda: _pipeline_loop.quit() if pipe_page.run_btn.isEnabled() else None)
_pipeline_poll.start(20)
QTimer.singleShot(30000, _pipeline_loop.quit)
_pipeline_loop.exec()
_pipeline_poll.stop()
print('pipeline status:', pipe_page.status.text(), flush=True)
shutil.rmtree(BASE, ignore_errors=True)
print('DONE', flush=True)
