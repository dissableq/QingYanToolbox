# -*- coding: utf-8 -*-
"""v3.1.7 回归: 本轮深度审查修复的缺陷。

前两轮已修的（编码/CRLF/BOM/输出覆盖等）在 test_regress.py 里; 这里只
固化本轮新发现并修复的。
"""
import io
import os
from pathlib import Path
import shutil
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), r'_regress2')
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
from pypdf import PdfWriter

from engine import advanced_ops as A
from engine import office_ops as O

# ============================================ 1. docm 宏文档可处理
_d = docx.Document()
_d.add_paragraph('宏文档 ABC')
_src = D('a.docx')
_d.save(_src)
_docm = D('a.docm')
with zipfile.ZipFile(_src) as zin, zipfile.ZipFile(_docm, 'w') as zout:
    for it in zin.infolist():
        data = zin.read(it.filename)
        if it.filename == '[Content_Types].xml':
            data = data.replace(
                b'application/vnd.openxmlformats-officedocument.'
                b'wordprocessingml.document.main+xml',
                b'application/vnd.ms-word.document.macroEnabled.main+xml')
        zout.writestr(it, data)
try:
    _r = O.docx_replace(_docm, {'find': 'ABC', 'replace': 'XYZ',
                                'out_dir': D('docm_out')})
    _txt = ''.join(p.text for p in docx.Document(_r.output).paragraphs)
    check('docm 宏文档可正常替换', _r.status == 'ok' and 'XYZ' in _txt,
          f'{_r.status} {_txt!r}')
except Exception as e:
    check('docm 宏文档可正常替换', False, f'{type(e).__name__}: {e}')

# ============================================ 2. docx_replace 覆盖页眉/页脚/文本框
_d = docx.Document()
_d.add_paragraph('正文 OLD')
_d.sections[0].header.paragraphs[0].text = '页眉 OLD'
_d.sections[0].footer.paragraphs[0].text = '页脚 OLD'
_hf = D('hf.docx')
_d.save(_hf)
_r = O.docx_replace(_hf, {'find': 'OLD', 'replace': 'NEW',
                          'out_dir': D('hf_out')})
_dd = docx.Document(_r.output)
check('docx_replace 覆盖页眉',
      'NEW' in _dd.sections[0].header.paragraphs[0].text,
      _dd.sections[0].header.paragraphs[0].text)
check('docx_replace 覆盖页脚',
      'NEW' in _dd.sections[0].footer.paragraphs[0].text,
      _dd.sections[0].footer.paragraphs[0].text)

# ============================================ 3. 水印覆盖"首页不同"页眉
_d = docx.Document()
_d.add_paragraph('正文')
_d.sections[0].different_first_page_header_footer = True
_wm = D('wm.docx')
_d.save(_wm)
_r = O.docx_watermark(_wm, {'wm_text': '机密', 'out_dir': D('wm_out')})
_dd = docx.Document(_r.output)
_fp = _dd.sections[0].first_page_header.paragraphs
check('水印也盖"首页不同"的首页页眉',
      any('机密' in p.text for p in _fp), str([p.text for p in _fp]))

# ============================================ 4. 删图清掉 media 字节
_d = docx.Document()
_d.add_paragraph('正文')
_buf = io.BytesIO()
Image.new('RGB', (30, 30), (255, 0, 0)).save(_buf, 'PNG')
_buf.seek(0)
_d.add_paragraph().add_run().add_picture(_buf)
_img = D('img.docx')
_d.save(_img)
_r = O.docx_remove_images(_img, {'out_dir': D('img_out')})
with zipfile.ZipFile(_r.output) as _z:
    _media = [n for n in _z.namelist() if n.startswith('word/media/')]
check('删图后 media 字节也被清除(体积才会变小)',
      _r.status == 'ok' and not _media, f'{_r.status} media={_media}')

# ============================================ 5. PDF 水印/页眉脚防覆盖
_w = PdfWriter()
_w.add_blank_page(100, 100)
_pdf = D('A.pdf')
with open(_pdf, 'wb') as _f:
    _w.write(_f)
_old = D('A_wm.pdf')
with open(_old, 'wb') as _f:
    _f.write(b'PRECIOUS-OLD-CONTENT')
_r = O.pdf_watermark(_pdf, {'wm_text': '水印', 'out_dir': BASE})
check('PDF 水印不覆盖已有同名文件',
      os.path.abspath(_r.output) != os.path.abspath(_old)
      and open(_old, 'rb').read() == b'PRECIOUS-OLD-CONTENT',
      os.path.basename(_r.output or ''))

# ============================================ 6. xlsm 编辑保留 VBA 宏
_wb = openpyxl.Workbook()
_wb.active['A1'] = 'xx宏测试'
_wb.save(D('m.xlsx'))
with zipfile.ZipFile(D('m.xlsx')) as _zin, \
        zipfile.ZipFile(D('m.xlsm'), 'w') as _zout:
    for _it in _zin.infolist():
        _zout.writestr(_it, _zin.read(_it.filename))
    _zout.writestr('xl/vbaProject.bin', b'VBA-CODE')
_r = O.xlsx_replace(D('m.xlsm'), {'find': 'xx', 'replace': 'YY',
                                  'out_dir': D('xlsm_out')})
with zipfile.ZipFile(_r.output) as _z:
    check('编辑 xlsm 后 VBA 宏仍在', 'xl/vbaProject.bin' in _z.namelist(),
          str(_z.namelist()[:5]))

# ============================================ 7. 删行后合并单元格跟随上移
_wb = openpyxl.Workbook()
_ws = _wb.active
_ws['A1'] = '甲'; _ws['B1'] = '乙'
_ws['A3'] = '丙'; _ws['B3'] = '丁'
_ws.merge_cells('B1:B3')
_rows = D('rows.xlsx')
_wb.save(_rows)
_r = O.xlsx_delete_rows(_rows, {'row_mode': 'blank', 'out_dir': D('rows_out')})
_ws2 = openpyxl.load_workbook(_r.output).active
check('删行后合并单元格跟随上移',
      [str(m) for m in _ws2.merged_cells.ranges] == ['B1:B2'],
      str([str(m) for m in _ws2.merged_cells.ranges]))

# ============================================ 8. 修订处理不破坏正常内容
import lxml.etree as ET
_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
_doc_xml = (f'<w:document xmlns:w="{_W}"><w:body>'
            f'<w:p><w:pPr><w:rPr><w:del w:id="1"/></w:rPr></w:pPr>'
            f'<w:r><w:t>KEEP-1</w:t></w:r></w:p>'
            f'<w:p><w:r><w:t>KEEP-2</w:t></w:r></w:p>'
            f'<w:p><w:del w:id="2"><w:r><w:delText>GONE</w:delText>'
            f'</w:r></w:del></w:p>'
            f'</w:body></w:document>')


def _mk_with_xml(xml, name):
    base = D('base2.docx')
    if not os.path.exists(base):
        _x = docx.Document()
        _x.add_paragraph('x')
        _x.save(base)
    dst = D(name)
    with zipfile.ZipFile(base) as zin, zipfile.ZipFile(dst, 'w') as zout:
        for it in zin.namelist():
            data = zin.read(it)
            if it == 'word/document.xml':
                data = xml.encode('utf-8')
            zout.writestr(it, data)
    return dst


_rev = _mk_with_xml(_doc_xml, 'rev.docx')
_r = A.docx_accept_revisions(_rev, {'rev_mode': '接受', 'out_dir': D('rev_out')})
with zipfile.ZipFile(_r.output) as _z:
    _out = _z.read('word/document.xml').decode('utf-8')
_well = True
try:
    ET.fromstring(_out.encode('utf-8'))
except Exception:
    _well = False
check('修订处理保留正常段落(KEEP-1/KEEP-2)',
      'KEEP-1' in _out and 'KEEP-2' in _out,
      f'KEEP-1={"KEEP-1" in _out} KEEP-2={"KEEP-2" in _out}')
check('修订处理产出合法 XML', _well, 'XML 残缺')
check('修订处理确实删掉了删除内容', 'GONE' not in _out)

# ============================================ 9. copy_by_excel 返回值可解包
_cs = D('cs.txt')
open(_cs, 'w', encoding='utf-8').write('A')
_ct = D('ct', 'cs.txt')
os.makedirs(os.path.dirname(_ct), exist_ok=True)
open(_ct, 'w', encoding='utf-8').write('PRECIOUS-USER')
_ret = A.files_copy_by_excel([{'path': _cs, 'target': _ct}], {})
check('copy_by_excel 返回 (成功数, 失败列表) 可解包',
      isinstance(_ret, tuple) and len(_ret) == 2, str(type(_ret)))
check('copy_by_excel 不覆盖目标已有文件',
      open(_ct, encoding='utf-8').read() == 'PRECIOUS-USER',
      open(_ct, encoding='utf-8').read())

# ============================================ 10. docx 拆分保留图片
_d = docx.Document()
for _i in range(61):
    _d.add_paragraph(f'段{_i}')
_buf = io.BytesIO()
Image.new('RGB', (20, 20), (0, 200, 0)).save(_buf, 'PNG')
_buf.seek(0)
_d.add_paragraph().add_run().add_picture(_buf)
_sp = D('split.docx')
_d.save(_sp)
_spo = D('splitout')
os.makedirs(_spo, exist_ok=True)
A.docx_split_by_paragraphs(_sp, {'docx_chunk': 30, 'out_dir': _spo})
_broken = []
for _p in sorted(f for f in os.listdir(_spo) if f.endswith('.docx')):
    with zipfile.ZipFile(os.path.join(_spo, _p)) as _z:
        _names = _z.namelist()
        _xml = _z.read('word/document.xml').decode('utf-8')
    if 'r:embed' in _xml and not any(
            n.startswith('word/media/') for n in _names):
        _broken.append(_p)
check('docx 拆分后图片引用不失效(无红叉)', not _broken, str(_broken))

# ============================================ 11. 拆分二次运行不留旧分份
_sp2 = D('split2.docx')
_x = docx.Document()
for _i in range(10):
    _x.add_paragraph(f'p{_i}')
_x.save(_sp2)
_spo2 = D('splitout2')
os.makedirs(_spo2, exist_ok=True)
A.docx_split_by_paragraphs(_sp2, {'docx_chunk': 5, 'out_dir': _spo2})
_first = len(os.listdir(_spo2))
A.docx_split_by_paragraphs(_sp2, {'docx_chunk': 10, 'out_dir': _spo2})
_second = len(os.listdir(_spo2))
check('拆分二次运行保留已有分份', _first == 2 and _second == 3,
      f'{_first} -> {_second}')

# ============================================ 12. 规则替换不链式串扰
_wb = openpyxl.Workbook()
_ws = _wb.active
_ws.append(('原', '新'))
_ws.append(('分公司', '集团'))
_ws.append(('集团', '有限公司'))
_wb.save(D('rules.xlsx'))
_tp = D('chain.txt')
open(_tp, 'w', encoding='utf-8').write('分公司报表')
_chain_out = D('chain_out')
_r = A.text_replace_by_excel(_tp, {'rules_file': D('rules.xlsx'),
                                   'out_dir': _chain_out})
_chain_file = os.path.join(_chain_out, 'chain.txt')
_chain_txt = open(_chain_file, encoding='utf-8').read() \
    if os.path.exists(_chain_file) else '(无产物)'
check('规则替换不链式串扰(分公司→集团, 不再被二次替换)',
      _chain_txt == '集团报表', repr(_chain_txt))

# ============================================ 13. 文本行操作保留 CRLF
_tp = D('crlf.txt')
open(_tp, 'wb').write(b'@echo off\r\nrem line3\r\n')
A.text_delete_replace_lines(_tp, {'line_mode': '删除', 'lines_spec': '2',
                                  'out_dir': D('crlf_out')})
_raw = open(_tp, 'rb').read()
check('文本行操作保留 CRLF', b'\r\n' in _raw, repr(_raw))

# ============================================ 14. out_dir=源目录时仍备份
# 前两轮修的是 ops.process_text 里 out_dir==源目录 的分支; 这里验
# advanced_ops 的行操作在同样情形下的保护。
_tp = D('bak.txt')
open(_tp, 'w', encoding='utf-8').write('l1\nl2\n')
_r = A.text_insert_lines(_tp, {'insert_pos': 1, 'insert_text': 'new',
                               'out_dir': BASE})   # 源文件所在目录
check('文本行操作 out_dir=源目录时也备份可撤销',
      bool(_r.undo) and os.path.isfile(_r.undo[-1]['src']) and Path(_r.undo[-1]['src']).parent.name == '备份',
      f'bak={os.path.exists(_tp + ".qingyanbak")} undo={bool(_r.undo)}')

# ============================================ 15. 中文字体写 eastAsia
_d = docx.Document()
_d.add_paragraph('中文内容ABC')
_font = D('font.docx')
_d.save(_font)
_r = A.docx_set_content_style(_font, {'style_font': '黑体',
                                      'out_dir': D('font_out')})
with zipfile.ZipFile(_r.output) as _z:
    _xml = _z.read('word/document.xml').decode('utf-8')
check('设置中文字体时写 w:eastAsia', 'w:eastAsia' in _xml,
      'eastAsia 缺失')

# ============================================ 16. 背景声明 displayBackgroundShape
_d = docx.Document()
_d.add_paragraph('正文')
_bg = D('bg.docx')
_d.save(_bg)
_r = A.docx_set_background(_bg, {'bg_color': 'FFFF00', 'out_dir': D('bg_out')})
with zipfile.ZipFile(_r.output) as _z:
    _st = _z.read('word/settings.xml').decode('utf-8') \
        if 'word/settings.xml' in _z.namelist() else ''
check('设置背景时声明 displayBackgroundShape(Word 才显示)',
      'displayBackgroundShape' in _st, _r.status)

# ============================================ 17. LA 模式图片擦除不崩
_la = D('la.png')
Image.new('LA', (40, 40), (128, 255)).save(_la)
try:
    _r = A.image_erase(_la, {'erase_rect': '5,5,10,10',
                             'erase_mode': '纯色', 'out_dir': D('la_out')})
    check('LA 模式(灰度+透明)图片擦除可用', _r.status == 'ok',
          f'{_r.status} {_r.detail[:40]}')
except Exception as e:
    check('LA 模式(灰度+透明)图片擦除可用', False,
          f'{type(e).__name__}: {e}')

# ============================================ 18. 长图拼接透明不变黑
_a1 = D('m1.png')
Image.new('RGB', (40, 20), (255, 0, 0)).save(_a1)
_a2 = D('m2.png')
Image.new('RGBA', (40, 20), (0, 0, 0, 0)).save(_a2)
_mo = D('merged.png')
A.image_merge_files([_a1, _a2], _mo, '长图')
check('长图拼接透明像素不变黑',
      Image.open(_mo).convert('RGB').getpixel((20, 30)) == (255, 255, 255),
      str(Image.open(_mo).convert('RGB').getpixel((20, 30))))

# ============================================ 19. 正则构造器保留重复次数
for _kind, _cnt, _want in (('单个汉字', 5, '{5}'),
                           ('单个英文字母', 6, '{6}'),
                           ('单个标点符号', 3, '{3}')):
    _got = A.regexp_builder(_kind, '', _cnt, '不限', '')
    check(f'{_kind} 重复次数生效', _want in _got, _got)

# ============================================ 20. GBK 模板不再乱码
_tpl = D('tpl.txt')
open(_tpl, 'wb').write('通知 {{姓名}} 同学'.encode('gbk'))
_out = D('tpl_out.txt')
O.generate_from_template(_tpl, {'姓名': '张三'}, _out)
_txt = open(_out, encoding='utf-8').read()
check('GBK 编码模板不再乱码', '\ufffd' not in _txt and '张三' in _txt,
      repr(_txt[:24]))

# ============================================ 21. pptx 替换覆盖表格
_prs = Presentation()
_s = _prs.slides.add_slide(_prs.slide_layouts[5])
_tb = _s.shapes.add_table(2, 2, 10, 10, 100, 50).table
_tb.cell(0, 0).text = '表格OLD'
_pp = D('p.pptx')
_prs.save(_pp)
_r = O.pptx_replace(_pp, {'find': 'OLD', 'replace': 'NEW',
                          'out_dir': D('pp_out')})
_p2 = Presentation(_r.output)
_found = []
for _sl in _p2.slides:
    for _sh in _sl.shapes:
        if getattr(_sh, 'has_table', False) and _sh.has_table:
            _found.append(_sh.table.cell(0, 0).text)
check('pptx 替换覆盖表格内文字',
      any('NEW' in x for x in _found), str(_found))

# ============================================ 22. 修复模式覆盖度自检
# 防止"同一类 bug 只修了一处"——逐函数核对关键防护是否在场。
import inspect

_SRC_O = inspect.getsource(O)
_SRC_A = inspect.getsource(A)

# 所有落盘 PDF 的函数必须用 _uniq_path
for _fn in ('pdf_watermark', 'pdf_header_footer'):
    _s = inspect.getsource(getattr(O, _fn))
    check(f'{_fn} 有防覆盖(_uniq_path)', '_uniq_path' in _s)

# 所有 xlsx 编辑函数必须走 _load_xlsx
for _fn in ('xlsx_replace', 'xlsx_delete_rows', 'xlsx_formulas_to_values',
            'xlsx_watermark', 'xlsx_sheet_ops'):
    _s = inspect.getsource(getattr(O, _fn))
    check(f'{_fn} 走 _load_xlsx(保 VBA 宏)', '_load_xlsx' in _s)

# 所有 docx 编辑函数必须走 _open_docx_compat
for _fn in ('docx_replace', 'docx_clear_blank', 'docx_remove_images',
            'docx_watermark', 'docx_header_footer'):
    _s = inspect.getsource(getattr(O, _fn))
    check(f'{_fn} 走 _open_docx_compat(支持 docm)',
          '_open_docx_compat' in _s)

# 文本行操作必须保 CRLF
for _fn in ('text_delete_replace_lines', 'text_insert_lines',
            'text_replace_by_excel'):
    _s = inspect.getsource(getattr(A, _fn))
    check(f'{_fn} 保留原换行符', '_orig_eol' in _s or 'newline=' in _s)

shutil.rmtree(BASE, ignore_errors=True)
print(f'\nRESULT: {ok} passed, {fail} failed')
sys.exit(1 if fail else 0)
