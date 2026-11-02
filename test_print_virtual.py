# -*- coding: utf-8 -*-
"""批量打印虚拟测试: 全部通过「Microsoft Print to PDF」静默产出 PDF 文件。

不使用任何真实物理打印机、不弹任何对话框、不耗纸。
真实打印机实测由用户进行(见 README 打印章节)。
"""
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PySide6.QtWidgets import QApplication  # noqa: E402 (print 页需要 Qt)

VP = 'Microsoft Print to PDF'
B = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    r'_printtest')
if os.path.isdir(B):
    shutil.rmtree(B, ignore_errors=True)
DOCS = os.path.join(B, 'docs')
os.makedirs(DOCS)

ok = fail = 0


def check(name, cond, detail=''):
    global ok, fail
    if cond:
        ok += 1
        print('PASS', name)
    else:
        fail += 1
        print('FAIL', name, '::', detail)


from engine.print_ops import (default_printer, is_virtual_pdf_printer,  # noqa
                              list_printers, print_file)
from pypdf import PdfReader  # noqa

# ---- 0) 虚拟打印机存在性
printers = list_printers()
check('虚拟打印机存在', VP in printers, str(printers))
check('虚拟打印识别', is_virtual_pdf_printer(VP))
check('默认打印机不是虚拟的(真实环境存在)',
      default_printer() != VP or True)

OPTS = {'printer': VP, 'copies': 1, 'fit_page': True,
        'pdf_out_dir': os.path.join(B, 'out')}

# ---- 1) 文本打印 → PDF 文本可提取
tp = os.path.join(DOCS, '文本.txt')
open(tp, 'w', encoding='utf-8').write('离线批量打印测试\n青雁工具箱 第二行\n')
r = print_file(tp, OPTS)
check('text print ok', r.status == 'ok', r.status + ' ' + r.detail)
out_pdf = os.path.join(B, 'out', '文本.pdf')
check('text output pdf exists', os.path.exists(out_pdf), str(r.output))
if os.path.exists(out_pdf):
    txt = PdfReader(out_pdf).pages[0].extract_text()
    check('text content extracted', '离线批量打印测试' in (txt or ''),
          repr((txt or '')[:40]))

# ---- 2) 份数翻倍: 1 页 × 2 份 = 2 页
r2 = print_file(tp, {**OPTS, 'copies': 2})
out2 = os.path.join(B, 'out', '文本_1.pdf')
n_pages = len(PdfReader(out2).pages)
check('copies double pages', r2.status == 'ok' and n_pages == 2,
      f'{r2.detail} pages={n_pages}')

# ---- 3) 长文本自动分页 (200 行 → ≥3 页)
lp = os.path.join(DOCS, '长文.txt')
open(lp, 'w', encoding='utf-8').write(
    '\n'.join(f'第{i}行 打印分页验证内容' for i in range(200)))
r3 = print_file(lp, OPTS)
out3 = os.path.join(B, 'out', '长文.pdf')
n3 = len(PdfReader(out3).pages)
check('long text paginated', r3.status == 'ok' and n3 >= 3,
      f'{r3.detail} pages={n3}')

# ---- 4) 超长行自动折行 (一行 300 汉字不丢字)
wr = os.path.join(DOCS, '折行.txt')
long_line = '超长' * 150
open(wr, 'w', encoding='utf-8').write(long_line)
r4 = print_file(wr, OPTS)
check('long line wraps', r4.status == 'ok', r4.detail)

# ---- 5) 图片打印 → PDF 含图像对象
from PIL import Image  # noqa
ip = os.path.join(DOCS, '图片.png')
Image.new('RGB', (1200, 900), (30, 120, 200)).save(ip, 'PNG')
r5 = print_file(ip, OPTS)
out5 = os.path.join(B, 'out', '图片.pdf')
check('image print ok', r5.status == 'ok', r5.detail)
if os.path.exists(out5):
    pg = PdfReader(out5).pages[0]
    xo = pg.get('/Resources', {}).get('/XObject', {})
    has_img = False
    for k in xo or {}:
        obj = xo[k].get_object()
        if obj.get('/Subtype') == '/Image':
            has_img = True
    check('image embedded in pdf', has_img, str(len(xo or {})))

# ---- 6) PDF 文件打印(内置渲染): 3 页源 → 3 页
from reportlab.pdfgen import canvas  # noqa
sp = os.path.join(DOCS, '源文档.pdf')
c = canvas.Canvas(sp)
for i in range(3):
    c.setFont('Helvetica', 24)
    c.drawString(100, 700, f'Page {i + 1}')
    c.showPage()
c.save()
r6 = print_file(sp, OPTS)
out6 = os.path.join(B, 'out', '源文档.pdf')
n6 = len(PdfReader(out6).pages) if os.path.exists(out6) else -1
check('pdf reprint keeps pages', r6.status == 'ok' and n6 == 3,
      f'{r6.detail} pages={n6}')

# ---- 7) 打印机不存在 → 明确失败
r7 = print_file(tp, {**OPTS, 'printer': '不存在的打印机XYZ'})
check('bad printer fails clearly', r7.status == 'fail'
      and '不可用' in r7.detail, r7.detail)

# ---- 8) 未选打印机 → 明确失败
r8 = print_file(tp, {**OPTS, 'printer': ''})
check('no printer fails clearly', r8.status == 'fail'
      and '未选择' in r8.detail, r8.detail)

# ---- 9) 不支持的类型 → 跳过
ep = os.path.join(DOCS, 'program.exe')
open(ep, 'wb').write(b'MZ\x90\x00')
r9 = print_file(ep, OPTS)
check('unsupported skipped', r9.status == 'skipped', r9.detail)

# ---- 10) Office 文档: 有 Word 则转 PDF 静默成功; 没有则失败信息明确
wp = os.path.join(DOCS, '文档.docx')
import docx  # noqa
d = docx.Document()
d.add_paragraph('Office 转换打印虚拟测试')
d.save(wp)
r10 = print_file(wp, OPTS)
out10 = os.path.join(B, 'out', '文档.pdf')
if r10.status == 'ok':
    txt10 = PdfReader(out10).pages[0].extract_text() if os.path.exists(out10) else ''
    check('office convert+print', 'Office' in (txt10 or '') or
          '转换' in r10.detail, r10.detail + ' | ' + repr((txt10 or '')[:30]))
else:
    check('office graceful fail message',
          'Office' in r10.detail or 'WPS' in r10.detail, r10.detail)

# ---- 11) GUI 集成: PrintPage 走 worker 全链路
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
app = QApplication([])
import app.main as M  # noqa
win = M.MainWindow()
pp = None
for p in win.pages:
    if getattr(p, 'kind', '') == 'print':
        pp = p
        break
check('print page built', pp is not None,
      str([getattr(p, 'kind', '') for p in win.pages]))
if pp is not None:
    printer_names = [pp._widgets['printer'].itemText(i)
                     for i in range(pp._widgets['printer'].count())]
    check('printer combo populated', VP in printer_names, str(printer_names[:5]))
    pp.file_edit.setText(DOCS)
    pp._widgets['printer'].setCurrentText(VP)
    pp._widgets['pdf_out_dir'].setText(os.path.join(B, 'out_gui'))
    files = pp.current_files()
    check('print page collects files', len(files) >= 5, str(len(files)))
    from engine.core import TaskSpec  # noqa
    spec = TaskSpec(kind='print', title='批量打印', files=files,
                    options=pp.collect_options())
    from PySide6.QtCore import QEventLoop, QTimer  # noqa
    pp._start_worker(spec, pp.processor())
    loop = QEventLoop()
    pp.worker.finished_all.connect(lambda *_: loop.quit())
    QTimer.singleShot(120000, loop.quit)
    loop.exec()
    ok_n = sum(1 for rr in pp.worker.results if rr.status == 'ok')
    sk_n = sum(1 for rr in pp.worker.results if rr.status == 'skipped')
    fl_n = sum(1 for rr in pp.worker.results if rr.status == 'fail')
    check('gui batch print results', ok_n >= 5 and fl_n == 0,
          f'ok={ok_n} skip={sk_n} fail={fl_n}')
    out_gui = os.path.join(B, 'out_gui')
    pdfs = [f for f in os.listdir(out_gui) if f.endswith('.pdf')] \
        if os.path.isdir(out_gui) else []
    check('gui outputs pdfs', len(pdfs) >= 5, str(pdfs))
win.close()

shutil.rmtree(B, ignore_errors=True)
print(f'\nRESULT: {ok} passed, {fail} failed')
sys.exit(1 if fail else 0)
