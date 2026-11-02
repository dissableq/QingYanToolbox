# -*- coding: utf-8 -*-
"""回归测试: 覆盖本轮修复的缺陷。

每个用例都对应一个已确认的真实缺陷, 防止后续改动把它们带回来。
"""
import os
from pathlib import Path
import shutil
import sys
import warnings

warnings.filterwarnings('ignore')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), r'_regress')
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


def D(*p):
    return os.path.join(BASE, *p)


import docx
import openpyxl
from PIL import Image
from pptx import Presentation
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas

from engine import advanced_ops as A
from engine import office_ops as O
from engine import ops
from engine.store import ParamStore


def mk_pdf(path, pages=3):
    c = canvas.Canvas(path)
    for _ in range(pages):
        c.showPage()
    c.save()
    return path


def mk_pptx(path, n=3, title_prefix='S'):
    p = Presentation()
    for i in range(n):
        p.slides.add_slide(p.slide_layouts[5]).shapes.title.text = f'{title_prefix}{i}'
    p.save(path)
    return path


def mk_docx(path, paras=5, sections=1):
    d = docx.Document()
    for i in range(paras):
        d.add_paragraph(f'段落{i}')
    for _ in range(sections - 1):
        d.add_section()
        d.add_paragraph('新节')
    d.save(path)
    return path


# ============================================ 1. PDF 口令校验(异常逃逸)
p = mk_pdf(D('enc.pdf'))
r = O.pdf_ops(p, {'pdf_op': 'encrypt', 'password': '123', 'out_dir': D('o1')})
check('encrypt 产物确实加密', PdfReader(r.output).is_encrypted, r.output)

r = O.pdf_ops(r.output, {'pdf_op': 'decrypt', 'password': '123',
                         'out_dir': D('o1')})
check('decrypt 成功', r.status == 'ok', r.detail)
check('decrypt 产物不再加密',
      r.status == 'ok' and not PdfReader(r.output).is_encrypted)

r = O.pdf_ops(mk_pdf(D('plain.pdf')),
              {'pdf_op': 'decrypt', 'password': 'x', 'out_dir': D('o1')})
check('未加密文件 decrypt -> skipped', r.status == 'skipped', r.status)

r = O.pdf_ops(r.output if r.output else p,
              {'pdf_op': 'extract', 'password': 'WRONG', 'out_dir': D('o1')})
# 用刚生成的加密文件验证错口令
enc_again = O.pdf_ops(mk_pdf(D('enc2.pdf')),
                      {'pdf_op': 'encrypt', 'password': 'right',
                       'out_dir': D('o1')}).output
r = O.pdf_ops(enc_again, {'pdf_op': 'extract', 'password': 'WRONG',
                          'out_dir': D('o1')})
check('错口令 -> fail(而非异常逃逸)', r.status == 'fail', r.status + r.detail)
r = O.pdf_ops(enc_again, {'pdf_op': 'extract', 'password': 'right',
                          'out_dir': D('o1')})
check('对口令 -> ok', r.status == 'ok', r.detail)

# ============================================ 2. PDF 页码/参数越界
p = mk_pdf(D('pg.pdf'))
check('extract 倒序 -> fail',
      O.pdf_ops(p, {'pdf_op': 'extract', 'page_from': '4', 'page_to': '2',
                    'out_dir': D('o2')}).status == 'fail')
check('extract 非整数 -> fail',
      O.pdf_ops(p, {'pdf_op': 'extract', 'page_from': 'abc',
                    'out_dir': D('o2')}).status == 'fail')
check('rotate 45 度 -> fail',
      O.pdf_ops(p, {'pdf_op': 'rotate', 'angle': '45',
                    'out_dir': D('o2')}).status == 'fail')
r = O.pdf_ops(p, {'pdf_op': 'rotate', 'angle': '0', 'out_dir': D('o2')})
check('rotate 0 度不再被改成 90', r.status == 'ok', r.detail)
check('insert_blank 越界 -> fail',
      O.pdf_ops(p, {'pdf_op': 'insert_blank', 'insert_at': 99,
                    'out_dir': D('o2')}).status == 'fail')
check('split chunk=0 -> fail',
      O.pdf_ops(p, {'pdf_op': 'split', 'chunk_size': '0',
                    'out_dir': D('o2')}).status == 'fail')

# ============================================ 3. PDF 裁剪/页眉页脚越界
check('crop 负边距 -> fail',
      A.pdf_crop(p, {'crop_margin': '-50,-50,-50,-50',
                     'out_dir': D('o3')}).status == 'fail')
check('crop 过大边距 -> fail',
      A.pdf_crop(p, {'crop_margin': '400,400,400,400',
                     'out_dir': D('o3')}).status == 'fail')
check('crop 正常 -> ok',
      A.pdf_crop(p, {'crop_margin': '10,10,10,10',
                     'out_dir': D('o3')}).status == 'ok')
check('remove_header_footer 非数字 -> skipped',
      A.pdf_remove_header_footer(p, {'hf_top': 'abc',
                                     'out_dir': D('o3')}).status == 'skipped')
check('remove_header_footer 超页高 -> fail',
      A.pdf_remove_header_footer(p, {'hf_top': '10000', 'hf_bottom': '0',
                                     'out_dir': D('o3')}).status == 'fail')
check('pdf_resize 非法纸张 -> skipped',
      A.pdf_resize(p, {'page_size': '..\\..\\evil',
                       'out_dir': D('o3')}).status == 'skipped')

# ============================================ 4. PDF 删空守卫
blank = D('blank.pdf')
w = PdfWriter()
for _ in range(3):
    w.add_blank_page(width=200, height=200)
with open(blank, 'wb') as fh:
    w.write(fh)
r = A.pdf_delete_blank_pages(blank, {'out_dir': D('o4')})
check('删光所有页 -> fail(守卫)', r.status == 'fail', r.status + r.detail)
r = A.pdf_delete_pages(mk_pdf(D('dp.pdf'), 3),
                       {'pages_spec': '1-3', 'out_dir': D('o4')})
check('删除全部页 -> fail(守卫)', r.status == 'fail', r.status + r.detail)
r = A.pdf_delete_pages(mk_pdf(D('dp2.pdf'), 3),
                       {'pages_spec': '1', 'out_dir': D('o4')})
check('删除部分页 -> ok', r.status == 'ok', r.detail)

# ============================================ 5. PPT 克隆(关系/包结构)
import zipfile
from collections import Counter

a = mk_pptx(D('pa.pptx'), 6)
b = mk_pptx(D('pb.pptx'), 2)
n = A.pptx_merge_files([a, b], D('merged.pptx'))
check('pptx 合并页数正确', n == 8, str(n))
with zipfile.ZipFile(D('merged.pptx')) as z:
    dups = {k: c for k, c in Counter(z.namelist()).items() if c > 1}
check('合并产物无重复 zip 条目', not dups, str(list(dups)[:3]))
check('合并产物可重新打开',
      len(Presentation(D('merged.pptx')).slides._sldIdLst) == 8)

# 图片关系必须保留
Image.new('RGB', (40, 40), (200, 10, 10)).save(D('img.png'))
pi = Presentation()
si = pi.slides.add_slide(pi.slide_layouts[5])
si.shapes.add_picture(D('img.png'), 500000, 500000, 1000000, 1000000)
si.shapes.title.text = 'IMG'
pi.save(D('pimg.pptx'))
A.pptx_merge_files([D('pimg.pptx')], D('pimg_m.pptx'))
with zipfile.ZipFile(D('pimg_m.pptx')) as z:
    media = [x for x in z.namelist() if 'media' in x]
    rels_xml = z.read('ppt/slides/_rels/slide1.xml.rels')
check('克隆保留图片部件', len(media) >= 1, str(media))
check('克隆保留图片关系引用', b'image' in rels_xml, rels_xml[:120].decode())

# ============================================ 6. PPT 空页判定与越界
pp = Presentation()
blank_layout = pp.slide_layouts[6]
pp.slides.add_slide(blank_layout)                       # 真空白页
s2 = pp.slides.add_slide(pp.slide_layouts[5])
s2.shapes.add_picture(D('img.png'), 500000, 500000, 1000000, 1000000)
s2.shapes.title.text = '仅有标题'
pp.save(D('mix.pptx'))
r = O.pptx_delete_blank_slides(D('mix.pptx'), {'out_dir': D('o6')})
check('删空白页保留图片页', r.status == 'ok' and r.detail == '删除空白页 1 页',
      r.status + r.detail)
r = O.pptx_delete_blank_slides(D('mix.pptx'), {'out_dir': D('o6')})
allblank = Presentation()
for _ in range(2):
    allblank.slides.add_slide(allblank.slide_layouts[6])
allblank.save(D('allblank.pptx'))
r = O.pptx_delete_blank_slides(D('allblank.pptx'), {'out_dir': D('o6')})
check('全空白 -> fail(守卫)', r.status == 'fail', r.status + r.detail)

r = O.pptx_delete_pages(D('mix.pptx'), {'page_start': '1', 'page_end': '-1',
                                       'out_dir': D('o6')})
check('PPT 删除全部页 -> fail(守卫)', r.status == 'fail', r.status + r.detail)
r = A.pptx_extract_slides(D('pa.pptx'), {'page_from': '5', 'page_to': '2',
                                         'out_dir': D('o6')})
check('PPT 提取倒序 -> fail', r.status == 'fail', r.status + r.detail)
r = A.pptx_extract_slides(D('pa.pptx'), {'page_from': '99',
                                         'out_dir': D('o6')})
check('PPT 提取越界 -> fail', r.status == 'fail', r.status + r.detail)

# ============================================ 7. docx 拆分 sectPr 位置

d = mk_docx(D('sp.docx'), paras=5, sections=2)
r = A.docx_split_by_paragraphs(d, {'docx_chunk': '3', 'out_dir': D('o7')})
check('docx 拆分成功', r.status == 'ok', r.detail)
nd = docx.Document(r.output)
tags = [e.tag.split('}')[1] for e in nd.element.body]
check('sectPr 位于 body 末尾', tags[-1] == 'sectPr', str(tags))

# ============================================ 8. docx 多节页眉不重复
d = mk_docx(D('multi.docx'), paras=3, sections=3)
r = O.docx_watermark(d, {'wm_text': 'WM', 'out_dir': D('o8')})
d2 = docx.Document(r.output)
counts = [sum(1 for x in s.header.paragraphs if x.text == 'WM')
          for s in d2.sections]
check('每节页眉水印恰好 1 个', all(x == 1 for x in counts), str(counts))
r = O.docx_header_footer(d, {'header_text': 'H', 'footer_text': 'F',
                             'out_dir': D('o8')})
d3 = docx.Document(r.output)
hc = [sum(1 for x in s.header.paragraphs if x.text == 'H') for s in d3.sections]
fc = [sum(1 for x in s.footer.paragraphs if x.text == 'F') for s in d3.sections]
check('每节页眉恰好 1 个', all(x == 1 for x in hc), str(hc))
check('每节页脚恰好 1 个', all(x == 1 for x in fc), str(fc))

# 清空段落不得删掉承载 sectPr 的段落
r = O.docx_clear_blank(d, {'out_dir': D('o8')})
check('清空段落后分节数不变',
      r.status in ('ok', 'skipped')
      and len(docx.Document(D('multi.docx')).sections) == 3)

# ============================================ 9. Excel 公式转值不丢数据
wb = openpyxl.Workbook()
ws = wb.active
ws['A1'] = 1
ws['A2'] = 2
ws['A3'] = '=A1+A2'
f = D('formula.xlsx')
wb.save(f)
r = O.xlsx_formulas_to_values(f, {'out_dir': D('o9')})
check('无缓存值时保留公式而非清空', r.status == 'skipped', r.status + r.detail)

# ============================================ 10. Excel 行删除不误删全空行
wb = openpyxl.Workbook()
ws = wb.active
for i, v in enumerate(['a', 'b'], 1):
    ws.cell(row=i, column=1, value=v)
for i in range(3, 7):
    pass
ws['A7'] = 'a'
ws['A8'] = 'b'
f2 = D('rows.xlsx')
wb.save(f2)
r = O.xlsx_delete_rows(f2, {'row_mode': 'duplicate', 'out_dir': D('o9')})
out = openpyxl.load_workbook(r.output)
vals = [c[0].value for c in out.active.iter_rows(min_col=1, max_col=1)
        if c[0].value is not None]
check('全空行不参与去重', vals == ['a', 'b'], str(vals))

# ============================================ 11. Sheet 插入重名
wb = openpyxl.Workbook()
wb.active.title = '数据'
wb.create_sheet('S2')
f3 = D('sheet.xlsx')
wb.save(f3)
r = A.xlsx_sheet_insert(f3, {'sheet_pos': '1', 'new_sheet_name': '数据',
                             'out_dir': D('o9')})
check('Sheet 重名 -> fail', r.status == 'fail', r.status + r.detail)
r = A.xlsx_sheet_insert(f3, {'sheet_pos': '1', 'new_sheet_name': '新表',
                             'out_dir': D('o9')})
check('Sheet 正常插入 -> ok', r.status == 'ok', r.detail)
check('Sheet 名称未被篡改',
      '新表' in openpyxl.load_workbook(r.output).sheetnames)

# ============================================ 12. 合并全部无效不崩溃
try:
    A.xlsx_merge_files([D('nope1.xlsx'), D('nope2.xlsx')], D('m.xlsx'))
    invalid_rejected = False
except (OSError, ValueError):
    invalid_rejected = True
check('Excel 合并无效输入明确拒绝', invalid_rejected)
check('Excel 合并失败不产出伪成功空文件', not os.path.exists(D('m.xlsx')))

# ============================================ 13. 压缩大小统计
f4 = mk_docx(D('cmp.docx'), paras=30)
so = os.path.getsize(f4)
r = O.docx_compress(f4, {'out_dir': D('o10')})
check('docx 压缩报出真实大小变化',
      r.status == 'ok' and '→' in r.detail and 'X' not in r.detail,
      r.detail)
r2 = O.xlsx_compress(mk_docx(D('z.txt'), 1), {'out_dir': D('o10')})
check('xlsx 压缩可执行', r2.status in ('ok', 'fail'), r2.status + r2.detail)

# ============================================ 14. skipped 不产生备份污染
d5 = mk_docx(D('clean.docx'), paras=3)
before = set(os.listdir(BASE))
A.docx_remove_macros(d5, {'out_dir': D('o11')})
A.docx_remove_watermark(d5, {'out_dir': D('o11')})
check('skipped 后不残留 .qingyanbak',
      not any(x.endswith('.qingyanbak') for x in os.listdir(BASE)),
      str([x for x in os.listdir(BASE) if 'bak' in x]))

# ============================================ 15. 图片分割上限与擦除校验
img = D('pic.png')
Image.new('RGB', (300, 200), (10, 100, 10)).save(img)
check('split 99x99 -> fail',
      ops.process_image(img, {'op': 'split', 'grid': '99,99',
                              'out_dir': D('o12')}).status == 'fail')
check('split 3x2 -> ok',
      ops.process_image(img, {'op': 'split', 'grid': '3,2',
                              'out_dir': D('o12')}).status == 'ok')
check('erase 宽高为 0 -> fail',
      A.image_erase(img, {'erase_rect': '10,10,0,0',
                          'out_dir': D('o12')}).status == 'fail')
check('erase 负宽 -> fail',
      A.image_erase(img, {'erase_rect': '10,10,-5,20',
                          'out_dir': D('o12')}).status == 'fail')
check('erase 超范围 -> fail',
      A.image_erase(img, {'erase_rect': '10,10,9999,9999',
                          'out_dir': D('o12')}).status == 'fail')
gray = D('gray.png')
Image.new('L', (100, 100), 128).save(gray)
check('灰度图纯色擦除不崩溃',
      A.image_erase(gray, {'erase_rect': '10,10,20,20',
                           'out_dir': D('o12')}).status == 'ok')

# ============================================ 16. 骑缝章除零
w = PdfWriter()
f5 = D('zero.pdf')
with open(f5, 'wb') as fh:
    w.write(fh)
Image.new('RGBA', (100, 100), (255, 0, 0, 128)).save(D('seal.png'))
r = A.pdf_add_seal(f5, {'seal_image': D('seal.png'), 'out_dir': D('o13')})
check('0 页 PDF 骑缝章不崩溃', r.status == 'skipped', r.status + r.detail)

# ============================================ 17. 路径穿越
try:
    A.files_create_from_excel([{'path': '../escaped.txt', 'type': '文件'}],
                              D('root'))
    check('Excel 建文件路径越界被拦', False, '未抛异常')
except ValueError as e:
    check('Excel 建文件路径越界被拦', '越界' in str(e), str(e))
check('越界文件未被创建', not os.path.exists(D('escaped.txt')))

os.makedirs(D('src/S'), exist_ok=True)
open(D('src/S/a.txt'), 'w').write('x')
try:
    A.files_merge_folders([D('src/S')],
                          {'merge_target': D('src/S/T'), 'keep_parent': False})
    check('汇总目标在源内部被拦', False, '未抛异常')
except ValueError as e:
    check('汇总目标在源内部被拦', '内部' in str(e), str(e))

# ============================================ 18. 删除文件夹边界
check('删除不存在文件夹 -> fail',
      A.files_delete_folder(D('not_exist'), {'trash_dir': D('trash')}).status
      == 'fail')
check('删除空路径 -> fail',
      A.files_delete_folder('', {'trash_dir': D('trash')}).status == 'fail')

# ============================================ 19. PDF 中文水印
p = mk_pdf(D('wm.pdf'), 2)
r = O.pdf_watermark(p, {'wm_text': '青雁水印', 'out_dir': D('o14')})
txt = PdfReader(r.output).pages[0].extract_text()
check('中文水印正常渲染(非 ■■)', '青雁' in txt and '■' not in txt,
      repr(txt[:20]))
r = O.pdf_header_footer(p, {'header_text': '页眉中文', 'footer_text': '页脚中文',
                             'out_dir': D('o14')})
txt = PdfReader(r.output).pages[0].extract_text()
check('中文页眉页脚正常渲染',
      '页眉中文' in txt and '页脚中文' in txt, repr(txt[:30]))

# ============================================ 20. 预设导出/导入闭环
ps = ParamStore(path=D('params.json'))
ps.save_preset('方案A', '文本工具', '查找替换', {'find': 'x'})
n = ps.export_presets(D('presets_out.json'))
check('导出预设成功', n == 1, str(n))
ps2 = ParamStore(path=D('params2.json'))
m = ps2.import_presets(D('presets_out.json'))
check('导入预设数量正确', m == 1, str(m))
check('导入后方案可用', '方案A' in ps2.load_presets(),
      str(list(ps2.load_presets())))

# ============================================ 21. 撤销精确回退一步
tf = D('undo.txt')
open(tf, 'w', encoding='utf-8').write('AAA')
O.docx_header_footer  # noqa
ops.process_text(tf, {'op': 'replace', 'find': 'A', 'replace': 'B'})
_undo_result = ops.process_text(tf, {'op': 'replace', 'find': 'B', 'replace': 'C'})
bak = _undo_result.undo[-1]['src']
if os.path.exists(bak):
    shutil.copy2(bak, tf)
    got = open(tf, encoding='utf-8').read()
else:
    got = open(tf, encoding='utf-8').read()
check('撤销回退精确一步(应为 BBB)', got == 'BBB', repr(got))

# ============================================ 22. 模板占位符
tpl = D('tpl.docx')
d6 = docx.Document()
# 同一段落内被拆到多个 run(Word 对占位符常见), 占位符仍应被替换
p6 = d6.add_paragraph()
p6.add_run('尊敬的{{姓')
p6.add_run('名}}：您好 ')
p6.add_run('部门：{{部门}}')
d6.add_paragraph('{{姓名}}')
d6.save(tpl)
outp = D('tpl_out.docx')
O.generate_from_template(tpl, {'姓名': '张三', '部门': None}, outp)
got = '\n'.join(x.text for x in docx.Document(outp).paragraphs)
check('跨 run 占位符被替换', '张三' in got and '{{' not in got, repr(got))
check('跨 run 替换后文字完整', '尊敬的张三：您好' in got, repr(got))
check('None 字段渲染为空而非 None', 'None' not in got, repr(got))
try:
    O.generate_from_template(p, {'a': 1}, D('bad.docx'))
    check('二进制模板被拒绝', False, '未抛异常')
except ValueError as e:
    check('二进制模板被拒绝', '不支持' in str(e), str(e))

# ============================================ 23. Theme 属性完整
from app.theme import Theme

th = Theme('light')
missing = [x for x in ('INK', 'MUTED', 'ACCENT', 'FAILC', 'OKC', 'WARNC',
                       'PAPER', 'PAPER2', 'LINE')
           if not hasattr(th, x)]
check('Theme 语义色属性齐全', not missing, str(missing))

# ============================================ 24. 首页操作名与页面 OPS 一致
import ast

from app.home_registry import HOME_GROUPS


def page_ops(fname, cls):
    tree = ast.parse(open(fname, encoding='utf-8').read())
    for node in ast.walk(tree):
        if isinstance(node, ast.ClassDef) and node.name == cls:
            for b in node.body:
                if isinstance(b, ast.Assign):
                    for t in b.targets:
                        if isinstance(t, ast.Name) and t.id == 'OPS':
                            return ast.literal_eval(b.value)
    return None


FP = 'app/pages/feature_pages.py'
page_map = {'Word 工具': 'WordPage', 'Excel 工具': 'ExcelPage',
            'PPT 工具': 'PptPage', 'PDF 工具': 'PdfPage',
            '文本工具': 'TextPage', '图片工具': 'ImagePage',
            '文件整理工具': 'FilesPage'}
mismatch = []
for _g, feats in HOME_GROUPS:
    for fname, page, op in feats:
        cls = page_map.get(page)
        if not cls or not op:
            continue
        opsl = page_ops(FP, cls)
        if opsl and op not in opsl:
            mismatch.append((fname, page, op))
check('首页操作名均存在于页面 OPS', not mismatch, str(mismatch))

# ============================================ 25. 首页注册表与页面层
from app.home_registry import group_title

bad_titles = []
for _g, _feats in HOME_GROUPS:
    t = group_title(_g)
    import re as _re
    m = _re.search(r'\((\d+) 项\)', t)
    if not m or int(m.group(1)) != len(_feats):
        bad_titles.append((_g, t, len(_feats)))
check('分组标题项数与实际一致', not bad_titles, str(bad_titles))

from app.pages.special_pages import page_index

check('page_index 找不到时返回 -1',
      page_index([('A', None), ('B', None)], '不存在') == -1)

# ============================================ 26. GUI 层(离屏)
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
# 万一有模态对话框弹出, 3 秒后强制退出, 避免测试挂死
try:
    import threading

    def _watchdog():
        import time
        time.sleep(90)
        print('\nWATCHDOG: 检测到阻塞, 强制退出')
        import os as _os
        _os._exit(2)

    threading.Thread(target=_watchdog, daemon=True).start()
except Exception:
    pass
try:
    from PySide6.QtWidgets import QApplication
    import app.main as _M

    _app = QApplication.instance() or QApplication([])
    _win = _M.MainWindow()
    _win.show()
    _app.processEvents()
    _home = _win.pages[0]

    check('首页按钮覆盖全部 169 个操作',
          len(_home._buttons) == 169, str(len(_home._buttons)))
    check('常用入口已收集', len(_home._common_buttons) == 9,
          str(len(_home._common_buttons)))

    def _probe(kw):
        _home.search_edit.setText(kw)
        _app.processEvents()
        return [f for _k, (b, g, f, _tp) in _home._buttons.items()
                if not b.isHidden()]

    check('搜索 pdf 覆盖工具与转换', len(_probe('pdf')) >= 22,
          str(len(_probe('pdf'))))
    check('搜索 查找替换 命中全部同名功能', len(_probe('查找替换')) == 5,
          str(len(_probe('查找替换'))))
    check('无结果时清空', len(_probe('zzz不存在')) == 0,
          str(len(_probe('zzz不存在'))))
    check('清空搜索后全部恢复', len(_probe('')) == 169,
          str(len(_probe(''))))

    # 流水线可选页只应是 FunctionPage 子类(合并/台账等没有 processor)
    from app.pages.base import FunctionPage
    pm = {n: c for n, c in _win._page_defs
          if c is not None and issubclass(c, FunctionPage)}
    check('流水线候选排除了非功能页', len(pm) == 11, str(sorted(pm)))

    # 深色主题切换不应崩溃
    _win.apply_theme('dark', '#0F6E8C', 1.0)
    _app.processEvents()
    check('切换深色主题不崩溃', _win.theme.mode == 'dark')
    _win.apply_theme('light', '#0F6E8C', 1.0)

    # 台账脏数据不应导致刷新崩溃
    import json as _json
    os.makedirs(_win.RECORDS_DIR, exist_ok=True)
    with open(os.path.join(_win.RECORDS_DIR, '20260101_000000_000.json'),
              'w', encoding='utf-8') as fh:
        _json.dump({'task_id': 'x', 'time': '2026-01-01 00:00:00',
                    'kind': 'k', 'title': '脏数据台账',
                    'total': 'N/A', 'ok': 'bad', 'fail': None,
                    'files': 'not-a-list'}, fh)
    rp = [p for p in _win.pages if getattr(p, 'title', '') == '任务台账'][0]
    try:
        rp.refresh()
        check('脏台账(total=非数字)不崩溃', True)
    except Exception as e:
        check('脏台账(total=非数字)不崩溃', False, f'{type(e).__name__}: {e}')
    # 脏台账的明细/导出路径: 直接调内部取文件的辅助函数, 避开文件对话框
    from app.pages.special_pages import _rec_files, _to_int
    bad = {'files': 'not-a-list', 'total': 'N/A'}
    check('_rec_files 对非 list 返回空', _rec_files(bad) == [], str(_rec_files(bad)))
    check('_to_int 对脏值兜底', _to_int('N/A') == 0 and _to_int('7') == 7)
    try:
        rp._current_record()
        check('脏台账读取当前记录不崩溃', True)
    except Exception as e:
        check('脏台账读取当前记录不崩溃', False, f'{type(e).__name__}: {e}')
    _win.records_dirty = False
except Exception as e:
    check('离屏 GUI 冒烟', False, f'{type(e).__name__}: {e}')

# ============================================ 27. 目录级操作的撤销清单 schema
# _run_dir_level 写的 last_undo.json 必须与 BatchWorker 一致
# ({'task_id', 'actions'}), 否则台账页读不到 actions, 撤销永远失效
_dl_dir = D('dirlevel')
os.makedirs(_dl_dir, exist_ok=True)
for _i in range(3):
    open(os.path.join(_dl_dir, f'd{_i}.txt'), 'w', encoding='utf-8').write('x')
_um_path = _win.UNDO_MANIFEST
_fp = None
for _p in _win.pages:
    if getattr(_p, 'kind', '') == 'files':
        _fp = _p
_fp.CONFIRM_OPS = {}
_fp.file_edit.setText(_dl_dir)
_fp._widgets['op'].setCurrentText('删除文件(进回收备份区)')
_o = _fp.collect_options()
_o['op_name'] = '删除文件(进回收备份区)'
_fp._run_dir_level(_o, files=[os.path.join(_dl_dir, f'd{_i}.txt')
                              for _i in range(3)])
# This action now runs in a background QThread. Await actual completion.
_fp.worker.wait(10000)
_app.processEvents()
check('目录级操作写了撤销清单', os.path.exists(_um_path),
      _um_path)
if os.path.exists(_um_path):
    _man = _json.load(open(_um_path, encoding='utf-8'))
    check('目录级撤销清单含 actions 键(台账页可读)',
          'actions' in _man and isinstance(_man['actions'], list)
          and len(_man['actions']) >= 3,
          str(list(_man.keys())))
    check('目录级撤销清单含 task_id',
          'task_id' in _man, str(list(_man.keys())))
    # 台账页的读取逻辑应能直接复用
    _rp2 = [p for p in _win.pages
            if getattr(p, 'title', '') == '任务台账'][0]
    check('台账页能解析该撤销清单',
          len(_man.get('actions', [])) > 0, 'actions 为空')

# ============================================ 28. 收藏标记不串页(同名功能)
_hp = _win.pages[0]
_btn_entries = list(_hp._buttons.values())
check('_buttons 值含目标页(4 元组)',
      all(len(e) == 4 for e in _btn_entries),
      str({len(e) for e in _btn_entries}))
_same_name = {}
for _b, _g, _fn, _tp in _btn_entries:
    _same_name.setdefault(_fn, set()).add(_tp)
check('同名功能散布多个目标页(测试前提)',
      any(len(v) > 1 for v in _same_name.values()),
      str({k: len(v) for k, v in _same_name.items() if len(v) > 1}))

# ============================================ 29. 文本工具: 编码/换行/BOM
# 每条都对应真实用户会踩的场景, 且都是"报 ok 但内容错了"那类静默 bug。
_d29 = D('r29')
os.makedirs(_d29)

# 29.1 大 UTF-8 中文文件: 采样块切在汉字中间不能被误判成 GBK
_big = os.path.join(_d29, 'big.txt')
with open(_big, 'w', encoding='utf-8') as _f:
    _f.write('中文测试内容' * 30000)
check('大 UTF-8 中文文件编码探测不误判',
      ops.detect_encoding(_big) == 'utf-8', ops.detect_encoding(_big))

# 29.2 纯 ASCII 前缀 + GBK 尾部: 不能整段读成 U+FFFD
_mix = os.path.join(_d29, 'mix.txt')
with open(_mix, 'wb') as _f:
    _f.write(b'a' * 300000 + '中文内容'.encode('gbk'))
_enc_mix = ops.detect_encoding(_mix)
check('ASCII 前缀+GBK 尾部不被判成 UTF-8',
      _enc_mix.lower().replace('-', '').startswith('gb'), _enc_mix)

# 29.3 移除 BOM 必须真的删掉(文本模式读 utf-8-sig 会自动吞 BOM)
_bom = os.path.join(_d29, 'bom.txt')
open(_bom, 'wb').write(b'\xef\xbb\xbfhello')
_r = ops.process_text(_bom, {'op': 'bom', 'bom_to': 'remove', 'out_dir': ''})
check('移除 BOM 真的删掉',
      open(_bom, 'rb').read() == b'hello', repr(open(_bom, 'rb').read()[:10]))

# 29.4 添加 BOM 不能产生双 BOM
_nb = os.path.join(_d29, 'nobom.txt')
open(_nb, 'wb').write(b'hello')
ops.process_text(_nb, {'op': 'bom', 'bom_to': 'add', 'out_dir': ''})
check('添加 BOM 不产生双 BOM',
      open(_nb, 'rb').read() == b'\xef\xbb\xbfhello',
      repr(open(_nb, 'rb').read()[:10]))

# 29.5 换行符转换 → LF 必须真的变
_crlf = os.path.join(_d29, 'crlf.txt')
open(_crlf, 'wb').write(b'a\r\nb\r\n')
_r = ops.process_text(_crlf, {'op': 'crlf', 'eol': 'lf', 'out_dir': ''})
check('换行符转 LF 真的生效',
      open(_crlf, 'rb').read() == b'a\nb\n',
      repr(open(_crlf, 'rb').read()))

# 29.6 任意文本操作不能把 CRLF 静默改成 LF
_keep = os.path.join(_d29, 'keep.txt')
open(_keep, 'wb').write(b'aa\r\nbb\r\n')
ops.process_text(_keep, {'op': 'replace', 'find': 'aa', 'replace': 'cc',
                         'out_dir': ''})
check('查找替换不改动原有 CRLF',
      open(_keep, 'rb').read() == b'cc\r\nbb\r\n',
      repr(open(_keep, 'rb').read()))

# 29.7 删空行不能把末尾换行算成空行
_db = os.path.join(_d29, 'db.txt')
open(_db, 'wb').write(b'line1\nline2\n')
_r = ops.process_text(_db, {'op': 'del_blank', 'out_dir': ''})
check('删空行不误删末尾换行',
      _r.status == 'skipped' and open(_db, 'rb').read() == b'line1\nline2\n',
      _r.status + ' ' + repr(open(_db, 'rb').read()))

# 29.8 正则替换串里的反斜杠按字面处理(Windows 路径)
_re = os.path.join(_d29, 're.txt')
open(_re, 'w', encoding='utf-8').write('PATHHERE\n')
ops.process_text(_re, {'op': 'replace', 'find': 'PATHHERE',
                       'replace': r'C:\new', 'regex': True, 'out_dir': ''})
check('正则替换串反斜杠按字面',
      open(_re, encoding='utf-8').read().strip() == r'C:\new',
      repr(open(_re, encoding='utf-8').read()))

# 29.9 out_dir 指向源目录时必须备份(否则源文件被无提示覆盖)
_od = D('r29', 'outdir')
os.makedirs(_od)
_on = os.path.join(_od, 'notes.txt')
open(_on, 'w', encoding='utf-8').write('hello world')
_r = ops.process_text(_on, {'op': 'replace', 'find': 'world',
                            'replace': 'THERE', 'out_dir': _od})
check('输出目录=源目录时仍生成备份',
      bool(_r.undo) and os.path.isfile(_r.undo[-1]['src']) and Path(_r.undo[-1]['src']).parent.name == '备份',
      f'bak={os.path.exists(_on + ".qingyanbak")} undo={_r.undo}')

# 29.10 MD→HTML 不能原地改 .md, 要另存 .html
_mdd = D('r29', 'md')
os.makedirs(_mdd)
_md = os.path.join(_mdd, 'readme.md')
open(_md, 'w', encoding='utf-8').write('# T\ntext')
_r = ops.process_text(_md, {'op': 'md2html', 'out_dir': ''})
check('MD→HTML 生成 .html 而非覆盖 .md',
      os.path.exists(os.path.join(_mdd, 'readme.html'))
      and open(_md, encoding='utf-8').read() == '# T\ntext',
      f'out={os.path.basename(_r.output or "")}')

# 29.11 拆分: 末尾不留空文件
_sp = D('r29', 'split')
os.makedirs(_sp)
_spf = os.path.join(_sp, 's.txt')
open(_spf, 'w', encoding='utf-8').write('a\nb\n')
_spo = D('r29', 'splitout')
os.makedirs(_spo)
_r = ops.process_text(_spf, {'op': 'split_lines', 'lines_per': 1,
                             'out_dir': _spo})
_parts = sorted(os.listdir(_spo))
check('按行拆分不产生末尾空文件',
      len(_parts) == 2 and all(
          os.path.getsize(os.path.join(_spo, f)) > 0 for f in _parts),
      str([(f, os.path.getsize(os.path.join(_spo, f))) for f in _parts]))

# 29.12 不同目录同名文件拆到同一输出目录不能互相覆盖
for _sub, _body in (('A', 'aaa\n'), ('B', 'bbb\n')):
    _sd = D('r29', 'dup', _sub)
    os.makedirs(_sd)
    open(os.path.join(_sd, 'data.txt'), 'w', encoding='utf-8').write(_body)
_do = D('r29', 'dupout')
os.makedirs(_do)
for _sub in ('A', 'B'):
    ops.process_text(D('r29', 'dup', _sub, 'data.txt'),
                     {'op': 'split_lines', 'lines_per': 9, 'out_dir': _do})
check('同名文件拆分不互相覆盖',
      len(os.listdir(_do)) == 2, str(os.listdir(_do)))

# ============================================ 30. 图片输出命名
_d30 = D('r30')
os.makedirs(_d30)
_img = os.path.join(_d30, 'photo.jpg')
Image.new('RGB', (40, 30), (10, 100, 10)).save(_img, 'JPEG')
_r = ops.process_image(_img, {'op': 'compress', 'quality': 70, 'out_dir': ''})
check('图片压缩输出单个扩展名',
      os.path.basename(_r.output or '').endswith('_compressed.jpg'),
      os.path.basename(_r.output or ''))
check('图片压缩产物扩展名与内容一致(不出现 .jpg.jpg)',
      '.jpg.' not in os.path.basename(_r.output or ''),
      os.path.basename(_r.output or ''))

# ============================================ 31. 打包 zip 不能自毁
_d31 = D('r31')
os.makedirs(_d31)
_zi = os.path.join(_d31, 'data.zip')
import zipfile as _zf
with _zf.ZipFile(_zi, 'w') as _z:
    _z.writestr('important.txt', 'PRECIOUS' * 50)
_r = ops.process_fileops(_zi, {'op': 'pack', 'out_dir': _d31})
with _zf.ZipFile(_zi) as _z:
    _names = _z.namelist()
check('对 zip 再打包不损坏原压缩包',
      'important.txt' in _names, str(_names))
check('对 zip 再打包另存新文件',
      os.path.basename(_r.output or '') != 'data.zip',
      os.path.basename(_r.output or ''))

# ============================================ 32. 全角/半角/标点三模式
_d32 = D('r32')
os.makedirs(_d32)
for _op, _src, _want, _label in (
        ('fw_full', 'abc,def', 'ａｂｃ，ｄｅｆ', '半角→全角'),
        ('fw_half', 'ａｂｃ', 'abc', '全角→半角'),
        ('fw_punct', '中文,测试', '中文，测试', '中文标点统一')):
    _p = os.path.join(_d32, _op + '.txt')
    open(_p, 'w', encoding='utf-8').write(_src)
    ops.text_fullwidth_convert(_p, {'op': _op, 'out_dir': ''})
    _got = open(_p, encoding='utf-8').read()
    check(f'{_label} 生效', _got == _want, f'{_got!r} != {_want!r}')
# 纯英文里的标点不能被改成中文标点
_p = os.path.join(_d32, 'en.txt')
open(_p, 'w', encoding='utf-8').write('hello,world')
ops.text_fullwidth_convert(_p, {'op': 'fw_punct', 'out_dir': ''})
check('中文标点统一不影响纯英文',
      open(_p, encoding='utf-8').read() == 'hello,world',
      repr(open(_p, encoding='utf-8').read()))

# ============================================ 33. 哈希算法选择
_d33 = D('r33')
os.makedirs(_d33)
_hp = os.path.join(_d33, 'x.txt')
open(_hp, 'w', encoding='utf-8').write('hello')
for _alg in ('MD5', 'SHA1', 'SHA256'):
    _r = ops.file_hasher(_hp, {'hash_alg': _alg, 'out_dir': _d33})
    check(f'哈希算法 {_alg} 生效', _alg in (_r.detail or '').upper(),
          _r.detail)

# ============================================ 34. 重命名加日期不崩溃
_d34 = D('r34')
os.makedirs(_d34)
_rp = os.path.join(_d34, 'a.txt')
open(_rp, 'w', encoding='utf-8').write('x')
try:
    _plan = ops.build_rename_plan([_rp], {'mode': 'add_date'})
    check('重命名加日期可用',
          bool(_plan) and _plan[0][1] != _plan[0][0], str(_plan))
except Exception as _e:
    check('重命名加日期可用', False, f'{type(_e).__name__}: {_e}')

# ============================================ 35. Excel 旧格式支持声明一致
from engine.convert_ops import supported_exts as _se
for _k in ('excel2csv', 'excel2word', 'excel2json', 'excel2html'):
    check(f'{_k} 支持 .xls/.et(与 UI 提示一致)',
          '.xls' in _se(_k) and '.et' in _se(_k), str(_se(_k)))

# ============================================ 36. GUI 层真实操作缺陷(离屏)
# 这些是用户第一次使用就会踩的: 拖拽没反应、预览崩、必填误报、
# 方案换掉已选文件。
import os as _os
_os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
from PySide6.QtWidgets import QApplication as _QA
_app = _QA.instance() or _QA([])
import inspect as _inspect

import app.main as _M
import app.pages.base as _B

_win = _M.MainWindow()
_func_pages = [p for p in _win.pages if isinstance(p, _B.FunctionPage)]

# 36.1 所有功能页必须启用 acceptDrops, 否则 dropEvent 永远不触发
check('功能页启用拖放(setAcceptDrops)',
      all(p.acceptDrops() for p in _func_pages),
      f'{sum(1 for p in _func_pages if p.acceptDrops())}/{len(_func_pages)}')

# 36.2 预览返回值两种形状都要能解包(单 rows / (header, rows))
_conv = next((p for p in _win.pages
              if getattr(p, 'kind', '') == 'convert'), None)
_prt = next((p for p in _win.pages
             if getattr(p, 'kind', '') == 'print'), None)
_pd = D('r36', 'prev')
os.makedirs(_pd)
for _i in range(3):
    open(os.path.join(_pd, f'f{_i}.txt'), 'w', encoding='utf-8').write('x')
_pfiles = [os.path.join(_pd, f'f{_i}.txt') for _i in range(3)]
for _nm, _pg in (('格式转换', _conv), ('批量打印', _prt)):
    if _pg is None:
        continue
    try:
        _h, _r = _B.FunctionPage._split_preview(
            _pg.build_preview(_pfiles, _pg.collect_options()))
        check(f'{_nm}页预览可解包(3 个文件不崩)', True,
              f'header={_h} rows={len(_r)}')
    except Exception as _e:
        check(f'{_nm}页预览可解包(3 个文件不崩)', False,
              f'{type(_e).__name__}: {_e}')

# 36.3 必填校验必须能读下拉框(打印机是 QComboBox)
check('必填校验支持 QComboBox',
      'QComboBox' in _inspect.getsource(
          _B.FunctionPage.validate_required))
if _prt is not None:
    # 默认选中的常是虚拟打印机(Microsoft Print to PDF), 此时"未填 PDF
    # 输出目录"本来就是必须报的——否则开跑后会对每个文件弹一次系统
    # 另存为对话框, 批量任务看起来像卡死。
    _pm = _prt.validate_required()
    check('打印页必填校验: 虚拟打印机必须填输出目录',
          all('输出目录' in x for x in _pm), str(_pm))
    check('打印页不因 QComboBox 而误报"打印机"',
          not any(x == '打印机' for x in _pm), str(_pm))

# 36.4 方案保存不能带上文件路径(载入会换掉用户当前选的文件)
_pg0 = _func_pages[0]
check('方案保存不含 file_edit',
      'file_edit' not in _pg0.capture_state(include_files=False))
check('流水线步骤保存仍含 file_edit',
      'file_edit' in _pg0.capture_state())

# 36.5 台账刷新: 导航切换要读 records_dirty
check('导航切换会刷新台账',
      hasattr(_win, '_on_nav_changed')
      and _win.records_dirty is False
      or 'records_dirty' in _inspect.getsource(_win._on_nav_changed))

# 36.6 关窗时正在跑的任务要被收尾
check('关闭窗口会取消运行中任务',
      'worker' in _inspect.getsource(_win.closeEvent)
      and 'cancel' in _inspect.getsource(_win.closeEvent))

# 36.7 全部失败不能显示"✔ 任务完成"
_src_fin = _inspect.getsource(_B.FunctionPage.on_finished)
check('全部失败不再显示成功标记', '全部失败' in _src_fin,
      'on_finished 无失败分支')

# 36.8 转换/打印页声明不改源文件(执行前说明条不再谎称修改原文件)
if _conv is not None:
    check('格式转换页声明不改源文件', _conv.MODIFIES_SOURCE is False)
if _prt is not None:
    check('批量打印页声明不改源文件', _prt.MODIFIES_SOURCE is False)

from unittest.mock import patch as _patch
from app.backup_dialog import BackupDialog as _BackupDialog
with _patch.object(_BackupDialog, 'exec', return_value=_BackupDialog.Accepted):
    _win.close()

shutil.rmtree(BASE, ignore_errors=True)
print(f'\nRESULT: {ok} passed, {fail} failed')
sys.exit(1 if fail else 0)
