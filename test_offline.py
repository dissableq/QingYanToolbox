# -*- coding: utf-8 -*-
"""离线纯净性审计: 网络审计钩子 + 全页面 + 代表性操作, 断言零网络事件。

审计事件覆盖: socket.connect / socket.getaddrinfo / urllib.Request /
http.client.connect 等(Python 3.8+ audit hooks, C 层 socket 调用必触发)。
"""
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

NET_EVENTS = []


def _audit_hook(event, args):
    if event in ('socket.connect', 'socket.getaddrinfo', 'socket.sendto',
                 'urllib.Request', 'http.client.connect'):
        # 过滤本机环回/Unix 套接字(仍记录, 但区分)
        try:
            a = str(args)[:160]
        except Exception:
            a = '?'
        NET_EVENTS.append((event, a))


sys.addaudithook(_audit_hook)

B = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    r'_offlinetest')
if os.path.isdir(B):
    shutil.rmtree(B, ignore_errors=True)
DOCS = os.path.join(B, 'docs')
IMGS = os.path.join(B, 'imgs')
OFF = os.path.join(B, 'office')
for d in (DOCS, IMGS, OFF):
    os.makedirs(d)
open(os.path.join(DOCS, 'a.txt'), 'w', encoding='utf-8').write('离线测试 ABC\n你好世界\n')
open(os.path.join(DOCS, 'b.json'), 'w', encoding='utf-8').write('{"k": [1,2,3]}')
from PIL import Image
Image.new('RGB', (400, 300), (99, 66, 33)).save(os.path.join(IMGS, 'p.jpg'), 'JPEG')
import docx
d = docx.Document(); d.add_paragraph('离线 Word 段落'); d.save(os.path.join(OFF, 'a.docx'))
import openpyxl
wb = openpyxl.Workbook(); wb.active.append(['列1', 2]); wb.save(os.path.join(OFF, 'a.xlsx'))
from pptx import Presentation
prs = Presentation(); prs.slides.add_slide(prs.slide_layouts[5]); prs.save(os.path.join(OFF, 'a.pptx'))
from reportlab.pdfgen import canvas as _cv
p_pdf = os.path.join(OFF, 'a.pdf')
c = _cv.Canvas(p_pdf); c.showPage(); c.showPage(); c.save()

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
from PySide6.QtWidgets import QApplication

app = QApplication([])
print('== building MainWindow (all 19 pages) ==', flush=True)
import app.main as M
win = M.MainWindow()
win.show()
app.processEvents()

from PySide6.QtCore import QEventLoop, QTimer


def run_worker(page, files, options):
    from engine.core import TaskSpec
    proc = page.processor()
    spec = TaskSpec(kind=page.kind, title=page.title, files=files,
                    options=options)
    page._start_worker(spec, proc)
    loop = QEventLoop()
    page.worker.finished_all.connect(lambda *_: loop.quit())
    page.worker.cancelled.connect(loop.quit)
    QTimer.singleShot(30000, loop.quit)
    loop.exec()
    return page.worker.results


def page_by(kind):
    for p in win.pages:
        if getattr(p, 'kind', '') == kind:
            return p
    return None


ops_run = 0

# ---- 文本页: 替换 + 编码 + JSON
tp = page_by('text')
tp.file_edit.setText(DOCS)
tp._widgets['op'].setCurrentText('查找替换')
tp._widgets['find'].setText('ABC')
tp._widgets['replace'].setText('XYZ')
tp._widgets['out_dir'].setText(os.path.join(B, 'out1'))
tp._widgets['out_dir'].setText('')
tp.ext_edit.setText('')
run_worker(tp, tp.current_files(), tp.collect_options())
ops_run += 1

# ---- 图片页: 压缩 + 水印
ip = page_by('image')
ip.file_edit.setText(IMGS)
ip._widgets['op'].setCurrentText('压缩')
ip._widgets['out_dir'].setText(os.path.join(B, 'out2'))
run_worker(ip, ip.current_files(), ip.collect_options())
ops_run += 1
ip._widgets['op'].setCurrentText('添加水印')
ip._widgets['wm_text'].setText('离线水印')
run_worker(ip, ip.current_files(), ip.collect_options())
ops_run += 1

# ---- Word 页: 替换 + 水印
wp = page_by('word')
wp.file_edit.setText(OFF)
wp._widgets['op'].setCurrentText('查找替换')
wp._widgets['find'].setText('离线')
wp._widgets['replace'].setText('本地')
wp._widgets['out_dir'].setText('')
wp.ext_edit.setText('')
run_worker(wp, wp.current_files(), wp.collect_options())
ops_run += 1

# ---- Excel 页: 转CSV
xp = page_by('excel')
xp.file_edit.setText(OFF)
xp._widgets['op'].setCurrentText('转 CSV')
xp._widgets['out_dir'].setText(os.path.join(B, 'out3'))
xp.ext_edit.setText('')
run_worker(xp, xp.current_files(), xp.collect_options())
ops_run += 1

# ---- PPT 页: 水印
pp = page_by('ppt')
pp.file_edit.setText(OFF)
pp._widgets['op'].setCurrentText('添加水印')
pp._widgets['out_dir'].setText(os.path.join(B, 'out4'))
pp.ext_edit.setText('')
run_worker(pp, pp.current_files(), pp.collect_options())
ops_run += 1

# ---- PDF 页: 中文水印 + 页眉页脚
dp = page_by('pdf')
dp.file_edit.setText(OFF)
dp._widgets['op'].setCurrentText('添加水印')
dp._widgets['wm'].setText('青雁水印')
dp._widgets['out_dir'].setText(os.path.join(B, 'out5'))
dp.ext_edit.setText('')
run_worker(dp, dp.current_files(), dp.collect_options())
ops_run += 1

# ---- 重命名页: 序号
rp = page_by('rename')
rp.file_edit.setText(DOCS)
rp._widgets['mode'].setCurrentText('序号编排')
rp._override_files = None
rp._plan_cache = __import__('engine.ops', fromlist=['build_rename_plan']) \
    .build_rename_plan(rp.current_files(),
                       {'mode': 'sequence', 'seq_start': 1, 'seq_pad': 3})
run_worker(rp, [p0[0] for p0 in rp._plan_cache],
           {**rp.collect_options(), '_rename_plan': rp._plan_cache})
ops_run += 1

# ---- 引擎级: 模板生成 + 合并 + 重复文件 + 哈希(无 GUI worker)
from engine import advanced_ops, office_ops, ops
tpl_dir = os.path.join(B, 'tpl')
tpl, data, cols = office_ops.materialize_office_template('请假条', tpl_dir)
ops_run += 1
office_ops.pdf_merge_files([p_pdf, p_pdf], os.path.join(B, 'm.pdf'))
ops_run += 1
advanced_ops.pptx_merge_files([os.path.join(OFF, 'a.pptx')],
                              os.path.join(B, 'm.pptx'))
ops_run += 1
ops.find_duplicate_files(DOCS, {'out_dir': B})
ops_run += 1
ops.file_hasher(DOCS, {'alg': 'md5'})
ops_run += 1
ops.text_stats(DOCS, {'out_dir': B})
ops_run += 1

print(f'== {ops_run} operation groups executed ==', flush=True)

net = list(NET_EVENTS)
print('network audit events:', len(net))
for e in net[:10]:
    print('  NET:', e)

from unittest.mock import patch
from app.backup_dialog import BackupDialog
with patch.object(BackupDialog,'exec',return_value=BackupDialog.Accepted):win.close()
shutil.rmtree(B, ignore_errors=True)

# 允许 Qt 内部对 localhost 的无害探测? 实测应为 0 —— 断言零网络
if net:
    print('FAIL: detected network attempts')
    sys.exit(1)
print('OFFLINE AUDIT: CLEAN — zero network events')
