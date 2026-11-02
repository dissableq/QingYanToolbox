# -*- coding: utf-8 -*-
"""补充测试: 之前未覆盖的功能路径。"""
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), r'_testarea2')
if os.path.isdir(BASE):
    shutil.rmtree(BASE, ignore_errors=True)
os.makedirs(BASE)

ok = fail = 0


def check(name, cond, detail=''):
    global ok, fail
    if cond:
        ok += 1
        print('PASS', name)
    else:
        fail += 1
        print('FAIL', name, '::', detail)


from engine import advanced_ops, office_ops, ops

# --- 准备
wdir = os.path.join(BASE, 'office')
os.makedirs(wdir)
import docx
d1 = docx.Document()
d1.add_paragraph('文件一 内容')
d1.add_paragraph('带水印测试')
d1.save(os.path.join(wdir, 'a.docx'))
d2 = docx.Document()
d2.add_paragraph('文件二 内容')
d2.save(os.path.join(wdir, 'b.docx'))
import openpyxl
wb = openpyxl.Workbook()
wb.active.append(['x', 1])
wb.save(os.path.join(wdir, 'a.xlsx'))
wb2 = openpyxl.Workbook()
wb2.active.append(['y', 2])
wb2.save(os.path.join(wdir, 'b.xlsx'))
from pptx import Presentation
p1 = Presentation()
p1.slides.add_slide(p1.slide_layouts[5])
p1.save(os.path.join(wdir, 'a.pptx'))
p2 = Presentation()
p2.slides.add_slide(p2.slide_layouts[5])
p2.slides.add_slide(p2.slide_layouts[5])
p2.save(os.path.join(wdir, 'b.pptx'))
from reportlab.pdfgen import canvas
for nm in ('a.pdf', 'b.pdf'):
    c = canvas.Canvas(os.path.join(wdir, nm))
    c.showPage()
    c.showPage()
    c.save()

from PIL import Image
Image.new('RGB', (300, 200), (10, 100, 10)).save(
    os.path.join(wdir, 'img.png'), 'PNG')

# --- Word 水印(验证 Pt/RGBColor 导入)
r = office_ops.docx_watermark(os.path.join(wdir, 'a.docx'), {})
check('docx_watermark', r.status == 'ok', r.status + ' ' + r.detail)

# --- Word 接受/拒绝修订
r = advanced_ops.docx_accept_revisions(os.path.join(wdir, 'a.docx'),
                                       {'rev_mode': '接受'})
check('docx accept revisions', r.status in ('ok', 'skipped'), r.detail)

# --- Word 压缩/背景/页眉页脚/样式
r = office_ops.docx_header_footer(os.path.join(wdir, 'a.docx'),
                                  {'header_text': '页眉', 'footer_text': '页脚'})
check('docx_header_footer', r.status == 'ok', r.detail)
r = advanced_ops.docx_set_background(os.path.join(wdir, 'a.docx'),
                                     {'bg_color': '#EAF5F9'})
check('docx_set_background', r.status == 'ok', r.detail)
r = advanced_ops.docx_set_page_style(os.path.join(wdir, 'a.docx'),
                                     {'page_orientation': '横向'})
check('docx_set_page_style', r.status == 'ok', r.detail)
r = advanced_ops.docx_meta(os.path.join(wdir, 'a.docx'), {'meta_mode': '清空'})
check('docx_meta clear', r.status == 'ok', r.detail)

# --- PPT 水印/背景/备注/页码
r = office_ops.pptx_watermark(os.path.join(wdir, 'a.pptx'), {})
check('pptx_watermark', r.status == 'ok', r.detail)
r = advanced_ops.pptx_set_background(os.path.join(wdir, 'a.pptx'),
                                     {'bg_color': '#AACCDD'})
check('pptx_set_background', r.status == 'ok', r.detail)
r = advanced_ops.pptx_extract_slides(os.path.join(wdir, 'b.pptx'),
                                     {'page_from': 1, 'page_to': 1})
check('pptx_extract_slides', r.status == 'ok', r.detail)

# --- Excel 水印/Sheet 操作/合并
r = office_ops.xlsx_watermark(os.path.join(wdir, 'a.xlsx'), {})
check('xlsx_watermark', r.status == 'ok', r.detail)
r = advanced_ops.xlsx_sheet_insert(os.path.join(wdir, 'a.xlsx'),
                                   {'sheet_pos': 1, 'new_sheet_name': '新表'})
check('xlsx_sheet_insert', r.status == 'ok', r.detail)
r = advanced_ops.xlsx_sheet_delete_replace(os.path.join(wdir, 'a.xlsx'),
                                           {'sheet_names': '新表'})
check('xlsx_sheet_delete', r.status == 'ok', r.detail)
n = advanced_ops.xlsx_merge_files(
    [os.path.join(wdir, 'a.xlsx'), os.path.join(wdir, 'b.xlsx')],
    os.path.join(wdir, 'merged.xlsx'))
check('xlsx_merge_files', n == 2 and os.path.exists(
    os.path.join(wdir, 'merged.xlsx')), str(n))

# --- PDF 合并/加密解密/链接/裁剪
r = office_ops.pdf_ops(os.path.join(wdir, 'a.pdf'),
                       {'pdf_op': 'encrypt', 'password': '123'})
check('pdf encrypt', r.status == 'ok', r.detail)
r = office_ops.pdf_ops(os.path.join(wdir, 'a.pdf'),
                       {'pdf_op': 'decrypt', 'password': '123'})
check('pdf decrypt', r.status == 'ok', r.detail)
r = advanced_ops.pdf_add_link(os.path.join(wdir, 'a.pdf'),
                              {'link_url': 'https://example.com'})
check('pdf add link', r.status == 'ok', r.detail)
r = advanced_ops.pdf_remove_links(os.path.join(wdir, 'a.pdf'), {})
check('pdf remove links', r.status in ('ok', 'skipped'), r.detail)
r = advanced_ops.pdf_crop(os.path.join(wdir, 'a.pdf'),
                          {'crop_margin': '10,10,10,10'})
check('pdf crop', r.status == 'ok', r.detail)
r = advanced_ops.pdf_rotate_noop = None
total = office_ops.pdf_merge_files(
    [os.path.join(wdir, 'a.pdf'), os.path.join(wdir, 'b.pdf')],
    os.path.join(wdir, 'merged.pdf'))
check('pdf_merge_files', total == 4 and os.path.exists(
    os.path.join(wdir, 'merged.pdf')), str(total))

# --- PPT 跨文件合并(克隆验证)
n = advanced_ops.pptx_merge_files(
    [os.path.join(wdir, 'a.pptx'), os.path.join(wdir, 'b.pptx')],
    os.path.join(wdir, 'merged.pptx'))
check('pptx_merge_files', n == 3 and os.path.exists(
    os.path.join(wdir, 'merged.pptx')), str(n))
# 合并结果能被重新打开
try:
    p = Presentation(os.path.join(wdir, 'merged.pptx'))
    check('pptx merged reopenable', len(p.slides._sldIdLst) == 3,
          str(len(p.slides._sldIdLst)))
except Exception as e:
    check('pptx merged reopenable', False, str(e))

# --- 图片分割边界(行列大于像素)
r = ops.process_image(os.path.join(wdir, 'img.png'), {'op': 'split', 'grid': '99,99'})
check('image split too large fails gracefully', r.status == 'fail', r.detail)
r = ops.process_image(os.path.join(wdir, 'img.png'),
                      {'op': 'resize', 'width': 100, 'height': 0})
check('image resize by width', r.status == 'ok', r.detail)
r = ops.process_image(os.path.join(wdir, 'img.png'),
                      {'op': 'rotate', 'angle': 90})
check('image rotate', r.status == 'ok', r.detail)
r = ops.process_image(os.path.join(wdir, 'img.png'),
                      {'op': 'convert', 'format': 'JPG', 'quality': 85})
check('image convert jpg', r.status == 'ok', r.detail)

# --- 文本: trim 零变更应 skipped
tp = os.path.join(BASE, 'trim.txt')
open(tp, 'w', encoding='utf-8').write('无空格行\n第二行\n')
r = ops.process_text(tp, {'op': 'trim'})
check('trim no-change skipped', r.status == 'skipped', r.status + r.detail)
open(tp, 'w', encoding='utf-8').write('有空格   \n第二行  \n')
r = ops.process_text(tp, {'op': 'trim'})
check('trim with-change ok', r.status == 'ok' and '2 行' in r.detail, r.detail)
r = ops.process_text(tp, {'op': 'replace', 'find': '(', 'replace': 'x',
                          'regex': True})
check('invalid regex fails cleanly', r.status == 'fail' and
      '正则' in r.detail, r.detail)

# --- 重命名冲突保护
rdir = os.path.join(BASE, 'ren')
os.makedirs(rdir)
open(os.path.join(rdir, 'a.txt'), 'w').write('1')
open(os.path.join(rdir, 'a_1.txt'), 'w').write('2')
plan = ops.build_rename_plan([os.path.join(rdir, 'a.txt')],
                             {'mode': 'replace', 'find': '', 'replace': ''})
check('rename plan no-op safe', plan and os.path.exists(plan[0][1]), str(plan))

shutil.rmtree(BASE, ignore_errors=True)
print(f'\nRESULT: {ok} passed, {fail} failed')
sys.exit(1 if fail else 0)
