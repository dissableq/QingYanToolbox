# -*- coding: utf-8 -*-
"""engine 端到端功能测试(不启动 GUI)。"""
import os
import shutil
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), r'_testarea')
if os.path.isdir(BASE):
    shutil.rmtree(BASE, ignore_errors=True)
os.makedirs(BASE)

from engine.core import collect_files
from engine import ops, office_ops, advanced_ops
from engine.store import Favorites, ParamStore, load_recent

ok = fail = 0
def check(name, cond, detail=''):
    global ok, fail
    if cond:
        ok += 1
        print(f'PASS {name}')
    else:
        fail += 1
        print(f'FAIL {name} :: {detail}')

# --- 准备测试文件
d = os.path.join(BASE, 'docs')
os.makedirs(d)
with open(os.path.join(d, 'a.txt'), 'w', encoding='utf-8') as f:
    f.write('你好世界 hello\n第二行 test123\n\n第四行\n你好世界 again\n')
with open(os.path.join(d, 'b.log'), 'w', encoding='gbk') as f:
    f.write('中文GBK编码日志\n2026-01-01 开始\n')
with open(os.path.join(d, 'c.md'), 'w', encoding='utf-8') as f:
    f.write('# 标题\n- 项目1\n- 项目2\n**加粗**文本\n')
with open(os.path.join(d, 'd.json'), 'w', encoding='utf-8') as f:
    f.write('{"b":1,"a":[1,2,3],"c":"x"}')
with open(os.path.join(d, 'e_empty.txt'), 'w', encoding='utf-8') as f:
    f.write('')
empty_dir = os.path.join(d, '空文件夹')
os.makedirs(empty_dir)
with open(os.path.join(d, '副本a_副本1.txt'), 'w', encoding='utf-8') as f:
    f.write('x')
with open(os.path.join(d, 'same1.txt'), 'w', encoding='utf-8') as f:
    f.write('重复内容一模一样ABC')
with open(os.path.join(d, 'same2.txt'), 'w', encoding='utf-8') as f:
    f.write('重复内容一模一样ABC')

# zip
zdir = os.path.join(BASE, 'zips')
os.makedirs(zdir)
zpath = os.path.join(zdir, '测试.zip')
with zipfile.ZipFile(zpath, 'w') as zf:
    zf.writestr('中文文件.txt', 'zip 内容')

# 图片
from PIL import Image
img_dir = os.path.join(BASE, 'imgs')
os.makedirs(img_dir)
img1 = Image.new('RGB', (800, 600), (200, 100, 50))
img1.save(os.path.join(img_dir, 'p1.jpg'), 'JPEG', quality=95)
img2 = Image.new('RGB', (600, 400), (50, 200, 100))
img2.save(os.path.join(img_dir, 'p2.png'), 'PNG')

# word / excel / ppt / pdf
import docx
wdir = os.path.join(BASE, 'office')
os.makedirs(wdir)
doc = docx.Document()
doc.add_paragraph('会议时间：2026-01-01')
doc.add_paragraph('重复段落')
doc.add_paragraph('')
doc.add_paragraph('含链接 http://example.com')
doc.save(os.path.join(wdir, 't1.docx'))
import openpyxl
wb = openpyxl.Workbook()
ws = wb.active
ws.append(['姓名', '数量'])
ws.append(['张三', 10])
ws.append(['李四', 20])
ws.append([None, None])
wb.save(os.path.join(wdir, 't1.xlsx'))
from pptx import Presentation
prs = Presentation()
s = prs.slides.add_slide(prs.slide_layouts[5])
prs.save(os.path.join(wdir, 't1.pptx'))
from reportlab.pdfgen import canvas
pdf_path = os.path.join(wdir, 't1.pdf')
c = canvas.Canvas(pdf_path)
c.showPage(); c.showPage(); c.showPage()
c.save()

# --- 文本处理
r = ops.process_text(os.path.join(d, 'a.txt'), {'op': 'replace', 'find': '你好', 'replace': '再见'})
check('text replace', r.status == 'ok' and '替换 2 处' in r.detail, r.detail)
r = ops.process_text(os.path.join(d, 'a.txt'), {'op': 'del_blank', 'find': '', 'replace': ''})
check('text del_blank', r.status == 'ok', r.detail)
r = ops.process_text(os.path.join(d, 'a.txt'), {'op': 'json_fmt'})
check('json_fmt on non-json txt fails', r.status == 'fail', r.detail)
r = ops.process_text(os.path.join(d, 'd.json'), {'op': 'json_fmt'})
check('json_fmt', r.status == 'ok', r.detail)
r = ops.text_fullwidth_convert(os.path.join(d, 'a.txt'), {'fw_mode': '半角'})
check('fullwidth convert', r.status in ('ok', 'skipped'), r.detail)
r = ops.detect_encoding(os.path.join(d, 'b.log'))
check('detect gbk encoding', 'gb' in r, r)

# --- 重命名
files = collect_files(d, False)
plan = ops.build_rename_plan(files, {'mode': 'sequence', 'seq_start': 1, 'seq_pad': 3})
check('rename plan', len(plan) > 0 and plan[0][2].startswith('001'), str(plan[:2]))
r = ops.process_rename(plan[0][0], {'_rename_plan': plan})
check('process_rename ok', r.status == 'ok' and r.undo, r.detail)
# 撤销
a = r.undo[0]
os.rename(a['dst'], a['src'])
r = ops.rename_by_excel_map(os.path.join(d, 'a.txt'), {'a.txt': ' renamed.txt'})
check('excel rename map', r.status == 'ok', r.detail)

# --- 图片
r = ops.process_image(os.path.join(img_dir, 'p1.jpg'),
                      {'op': 'compress', 'quality': 60, 'max_width': 400})
check('image compress', r.status == 'ok' and os.path.exists(r.output), r.detail)
r = ops.process_image(os.path.join(img_dir, 'p1.jpg'),
                      {'op': 'watermark', 'wm_text': '青雁', 'wm_pos': 'br',
                       'wm_opacity': 60, 'wm_size': 24})
check('image watermark', r.status == 'ok', r.detail)
r = ops.process_image(os.path.join(img_dir, 'p1.jpg'),
                      {'op': 'split', 'grid': '2,2'})
check('image split 2x2', r.status == 'ok', r.detail)
r = advanced_ops.image_erase(os.path.join(img_dir, 'p1.jpg'),
                             {'erase_rect': '10,10,50,50', 'erase_mode': '马赛克'})
check('image erase mosaic', r.status == 'ok', r.detail)
r = advanced_ops.image_merge_files([os.path.join(img_dir, 'p1.jpg'),
                                    os.path.join(img_dir, 'p2.png')],
                                   os.path.join(img_dir, 'long.png'), '长图')
check('image merge long', os.path.exists(r), r)

# --- 文件整理
r = ops.process_zip(zpath, {'out_dir': os.path.join(zdir, 'out')})
check('unzip gbk', r.status == 'ok' and r.detail.startswith('解压 1'), r.detail)
r = ops.process_fileops(d, {'op': 'list', 'out_dir': os.path.join(BASE, 'reports')})
check('file list csv', r.status == 'ok', r.detail)
r = ops.files_export_excel(d, {'out_dir': os.path.join(BASE, 'reports')})
check('file list excel', r.status == 'ok', r.detail)
r = ops.find_duplicate_files(d, {'out_dir': os.path.join(BASE, 'reports')})
check('dup find', r.status == 'ok' and '1 组重复' in r.detail, r.detail)
r = ops.cleanup_empty(d, {'empty_mode': 'both'})
check('cleanup empty', r.status in ('ok', 'skipped'), r.detail)
r = ops.file_hasher(d, {'alg': 'sha256'})
check('file hasher', r.status == 'ok', r.detail)
r = ops.touch_timestamps(d, {'mode': 'fixed', 'ts_fixed': '2026-01-01 12:00:00'})
check('touch timestamps', r.status == 'ok', r.detail)

# --- Word
r = office_ops.docx_replace(os.path.join(wdir, 't1.docx'),
                            {'find': '会议', 'replace': '聚会'})
check('docx replace', r.status == 'ok' and '替换 1 处' in r.detail, r.detail)
r = office_ops.docx_clear_blank(os.path.join(wdir, 't1.docx'), {})
check('docx clear blank', r.status in ('ok', 'skipped'), r.detail)
r = advanced_ops.docx_remove_links(os.path.join(wdir, 't1.docx'), {})
check('docx remove links', r.status in ('ok', 'skipped'), r.detail)
r = advanced_ops.docx_set_content_style(os.path.join(wdir, 't1.docx'),
                                        {'style_font': '微软雅黑', 'style_size': 14})
check('docx style', r.status in ('ok', 'skipped'), r.detail)
r = office_ops.docx_compress(os.path.join(wdir, 't1.docx'), {})
check('docx compress', r.status == 'ok', r.detail)
r = advanced_ops.docx_split_by_paragraphs(os.path.join(wdir, 't1.docx'),
                                          {'docx_chunk': 1})
check('docx split', r.status == 'ok', r.detail)

# --- Excel
r = office_ops.xlsx_replace(os.path.join(wdir, 't1.xlsx'),
                            {'find': '张三', 'replace': '王五'})
check('xlsx replace', r.status == 'ok', r.detail)
r = office_ops.xlsx_delete_rows(os.path.join(wdir, 't1.xlsx'), {'row_mode': 'blank'})
check('xlsx del blank rows', r.status in ('ok', 'skipped'), r.detail)
r = advanced_ops.xlsx_convert_format(os.path.join(wdir, 't1.xlsx'),
                                     {'xlsx_target': 'csv'})
check('xlsx to csv', r.status == 'ok', r.detail)
r = advanced_ops.xlsx_split(os.path.join(wdir, 't1.xlsx'), {'xlsx_split_mode': 'sheet'})
check('xlsx split by sheet', r.status == 'ok', r.detail)

# --- PPT
r = office_ops.pptx_watermark(os.path.join(wdir, 't1.pptx'), {})
check('pptx watermark', r.status == 'ok', r.detail)
r = advanced_ops.pptx_split(os.path.join(wdir, 't1.pptx'), {'ppt_chunk': 1})
check('pptx split', r.status == 'ok', r.detail)
r = advanced_ops.pptx_extract_notes(os.path.join(wdir, 't1.pptx'), {})
check('pptx extract notes (skipped=无备注)', r.status == 'skipped', r.detail)

# --- PDF
r = office_ops.pdf_ops(pdf_path, {'pdf_op': 'extract', 'page_from': 1, 'page_to': 2})
check('pdf extract', r.status == 'ok', r.detail)
r = office_ops.pdf_ops(pdf_path, {'pdf_op': 'split', 'chunk_size': 2})
check('pdf split', r.status == 'ok', r.detail)
r = office_ops.pdf_watermark(pdf_path, {'wm_text': '青雁水印', 'wm_opacity': 15})
check('pdf watermark', r.status == 'ok', r.detail)
r = office_ops.pdf_header_footer(pdf_path, {'header_text': '页眉', 'footer_text': '页脚'})
check('pdf header/footer', r.status == 'ok', r.detail)
r = advanced_ops.pdf_resize(pdf_path, {'page_size': 'A4'})
check('pdf resize', r.status == 'ok', r.detail)
r = advanced_ops.pdf_delete_pages(pdf_path, {'pages_spec': '3'})
check('pdf delete pages', r.status == 'ok', r.detail)
r = advanced_ops.pdf_add_seal(pdf_path, {'seal_image': os.path.join(img_dir, 'p2.png')})
check('pdf seal', r.status == 'ok', r.detail)
r = ops.pdf_blank_and_count(pdf_path, {'mode': 'count'})
check('pdf count', r.status == 'ok', r.detail)
r = ops.pdf_blank_and_count(d, {'mode': 'blank', 'pdf_pages': 3})
check('pdf blank gen', r.status == 'ok', r.detail)

# --- 模板生成
from engine.office_ops import materialize_office_template, generate_from_template
tpl_dir = os.path.join(BASE, 'tpl')
tpl, data, cols = materialize_office_template('请假条', tpl_dir)
check('materialize template', os.path.exists(tpl) and os.path.exists(data), tpl)
import openpyxl as _op
wb2 = _op.load_workbook(data)
rows = list(wb2.active.iter_rows(values_only=True))
row = dict(zip(rows[0], rows[1]))
out_docx = os.path.join(tpl_dir, 'generated.docx')
generate_from_template(tpl, row, out_docx)
check('generate from template', os.path.exists(out_docx), out_docx)

# --- store
ps = ParamStore()
ps.save_auto('word', '查找替换', {'find': 'x'})
check('param auto', ps.load_auto('word', '查找替换').get('find') == 'x')
ps.save_preset('测试方案', 'Word 工具', '查找替换', {'find': 'y'})
check('preset save', '测试方案' in ps.load_presets())
ps.save_pipeline('p1', [{'page': 'Word 工具', 'op': '查找替换', 'options': {}}])
check('pipeline save', 'p1' in ps.load_pipelines())
fav = Favorites()
fav.add('查找替换', 'Word 工具', '查找替换')
check('favorite add', fav.is_favorite('查找替换'))
fav.remove('查找替换')
check('favorite remove', not fav.is_favorite('查找替换'))
recents = load_recent(os.path.join(os.environ.get('APPDATA', ''), 'QingYanBatchToolbox', 'records'), 5)
print('recent count (info):', len(recents))

# --- 正则生成器
p = advanced_ops.regexp_builder('中国大陆手机号', '', 0, '', '')
check('regexp builder phone', p == r'1[3-9]\d{9}', p)
p = advanced_ops.regexp_builder('单个数字', '', 2, '某个文本的左边位置', 'ID:')
check('regexp builder pos', 'ID:' in p and r'\d' in p, p)

# 清理
shutil.rmtree(BASE, ignore_errors=True)
print(f'\nRESULT: {ok} passed, {fail} failed')
sys.exit(1 if fail else 0)
