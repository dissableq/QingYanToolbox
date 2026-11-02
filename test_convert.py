# -*- coding: utf-8 -*-
"""格式转换虚拟测试: 不碰真实打印机, 全部本地文件级验证。

验证矩阵:
1. Office(docx/xlsx/pptx) → PDF — 需本机 Office COM
2. PDF → Word — 关键: 产物必须是可编辑文字, 不是图片
3. PDF → 图片 / TXT / HTML
4. Word → TXT / HTML
5. Excel → CSV / HTML
6. PPT → TXT(PPT→图片同样走 COM, 一并验证)
7. 图片 → PDF; 文本 → Word
8. 守卫: 错误扩展名 → skipped; 无文字 PDF → fail(明确报错)
"""
import os
import shutil
import sys
import warnings

warnings.filterwarnings('ignore')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), r'_convtest')
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
from pptx import Presentation
from pypdf import PdfReader

try:
    import pymupdf
except ImportError:
    pymupdf = None

from engine import convert_ops as C


# ------------------------------------------- 探测本机 COM 组件可用性
def _com_available(prog):
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        app = win32com.client.DispatchEx(prog)
        try:
            app.Quit()
        except Exception:
            pass
        return True
    except Exception:
        return False


HAS_WORD = _com_available('Word.Application')
HAS_EXCEL = _com_available('Excel.Application')
HAS_PPT = _com_available('PowerPoint.Application')
print(f'COM 探测: Word={HAS_WORD} Excel={HAS_EXCEL} PPT={HAS_PPT}')


# ---------------------------------------------------------- 样本文件
# 中文 PDF(带不同字号 → 标题层级)
def mk_pdf(path):
    import pymupdf
    d = pymupdf.open()
    p1 = d.new_page()
    p1.insert_text((72, 90), '青雁工具箱转换测试', fontsize=22,
                   fontname='china-ss')
    p1.insert_text((72, 130), '第二章 转换功能', fontsize=16,
                    fontname='china-ss')
    p1.insert_text((72, 170), '这是正文内容, 小字号。', fontsize=10.5,
                   fontname='china-ss')
    p1.insert_text((250, 210), '居中标题测试', fontsize=18, fontname='china-ss')
    p2 = d.new_page()
    p2.insert_text((72, 80), '第二页内容', fontsize=11, fontname='china-ss')
    d.save(path)
    d.close()
    return path


pdf = mk_pdf(D('sample.pdf'))

wd = docx.Document()
wd.add_heading('测试文档标题', 0)
wd.add_paragraph('第一段正文。')
wd.add_paragraph('第二段正文, 加粗测试。').runs[0].bold = True
wd.add_paragraph('表格如下:')
t = wd.add_table(rows=2, cols=2)
t.cell(0, 0).text = '姓名'
t.cell(0, 1).text = '分数'
t.cell(1, 0).text = '张三'
t.cell(1, 1).text = '95'
docx_path = D('sample.docx')
wd.save(docx_path)

wb = openpyxl.Workbook()
ws = wb.active
ws.title = '成绩表'
ws.append(['姓名', '语文', '数学'])
ws.append(['张三', 88, 95])
ws.append(['李四', 76, 89])
ws2 = wb.create_sheet('第二表')
ws2.append(['A', 'B'])
ws2.append([1, 2])
xlsx_path = D('sample.xlsx')
wb.save(xlsx_path)

pp = Presentation()
s1 = pp.slides.add_slide(pp.slide_layouts[5])
s1.shapes.title.text = '幻灯片标题'
_body = s1.placeholders[1] if len(s1.placeholders) > 1 else None
if _body is not None:
    _body.text = '幻灯片正文内容'
else:
    from pptx.util import Emu
    _tb = s1.shapes.add_textbox(Emu(500000), Emu(2000000),
                                Emu(6000000), Emu(1200000))
    _tb.text_frame.text = '幻灯片正文内容'
pptx_path = D('sample.pptx')
pp.save(pptx_path)

from PIL import Image
img_path = D('sample.png')
Image.new('RGB', (300, 200), (200, 40, 40)).save(img_path)

txt_path = D('sample.txt')
with open(txt_path, 'w', encoding='utf-8') as fh:
    fh.write('第一行文本\n第二行文本\n第三行文本\n')

# 无文字 PDF(纯图片) → PDF→Word 应明确失败
nopdf = D('notext.pdf')
_d = pymupdf.open()
_p = _d.new_page()
_p.insert_image(pymupdf.Rect(72, 72, 200, 150), filename=img_path)
_d.save(nopdf)
_d.close()


# ============================================== 1. Office → PDF
r = C.convert_office_to_pdf(docx_path, {'out_dir': D('o_office')})
check('Word→PDF 成功', r.status == 'ok', r.status + ' ' + r.detail)
check('Word→PDF 产物存在且非空',
      r.status == 'ok' and os.path.getsize(r.output) > 1000, str(r.output))
check('Word→PDF 产物可读页',
      r.status == 'ok' and len(PdfReader(r.output).pages) >= 1)

r = C.convert_office_to_pdf(xlsx_path, {'out_dir': D('o_office')})
check('Excel→PDF 成功', r.status == 'ok', r.status + ' ' + r.detail)

r = C.convert_office_to_pdf(pptx_path, {'out_dir': D('o_office')})
if HAS_PPT:
    check('PPT→PDF 成功', r.status == 'ok', r.status + ' ' + r.detail)
else:
    check('无 PowerPoint 时 PPT→PDF 明确报错(不伪造产物)',
          r.status == 'fail' and 'PowerPoint' in r.detail
          and not os.path.exists(r.output or 'x'),
          r.status + r.detail)

r = C.convert_office_to_pdf(txt_path, {'out_dir': D('o_office')})
check('非 Office 文件 → skipped(不报错)', r.status == 'skipped', r.status)


# ============================================== 2. PDF → Word(可编辑!)
r = C.convert_pdf_to_word(pdf, {'out_dir': D('o_p2w')})
check('PDF→Word 成功', r.status == 'ok', r.status + ' ' + r.detail)
check('PDF→Word 产物是 .docx', r.status == 'ok'
      and r.output.lower().endswith('.docx'), str(r.output))
_d2 = docx.Document(r.output)
paras = [p for p in _d2.paragraphs if p.text.strip()]
check('PDF→Word 段落可数', len(paras) >= 4, str(len(paras)))
all_text = '\n'.join(p.text for p in paras)
check('PDF→Word 含中文正文', '正文' in all_text, all_text[:60])
check('PDF→Word 文字为可编辑文本(非空段落)', bool(all_text.strip()))
styles = [p.style.name for p in paras]
check('PDF→Word 检测出标题样式',
      any('Heading' in s or 'Title' in s for s in styles), str(styles[:6]))
# 可编辑性验证: python-docx 能读到文本 run → 真文字
runs_ok = any(p.runs for p in paras)
check('PDF→Word 产物含文本 run(可编辑)', runs_ok)

# 无文字 PDF → 明确失败
r = C.convert_pdf_to_word(nopdf, {'out_dir': D('o_p2w')})
check('无文字扫描版 PDF → fail 且说明原因',
      r.status == 'fail' and '文字' in (r.detail + r.status),
      r.status + r.detail)

# 非 PDF → skipped
r = C.convert_pdf_to_word(docx_path, {'out_dir': D('o_p2w')})
check('PDF→Word 收到 docx → skipped', r.status == 'skipped', r.status)

# 标题字号与正文差距大时, 即使标题行数不少也要能识别层级
big_pdf = D('bigtitle.pdf')
_b = pymupdf.open()
_pg = _b.new_page()
for _i in range(2):        # 2 行大标题
    _pg.insert_text((72, 90 + _i * 34), f'第{_i}章 大标题文字', fontsize=22,
                    fontname='china-ss')
for _i in range(4):        # 4 行正文
    _pg.insert_text((72, 190 + _i * 24), f'正文第{_i}行小字内容', fontsize=10,
                    fontname='china-ss')
_b.save(big_pdf)
_b.close()
r = C.convert_pdf_to_word(big_pdf, {'out_dir': D('o_p2w')})
_d4 = docx.Document(r.output)
_styles = [p.style.name for p in _d4.paragraphs if p.text.strip()]
check('标题占比不低时仍识别标题',
      any('Heading' in s or 'Title' in s for s in _styles), str(_styles[:6]))
# 标题 run 不得写入 <w:b w:val="0"/>(那会关掉 Heading 样式自带加粗)
import zipfile as _zf
_xml = _zf.ZipFile(r.output).read('word/document.xml').decode('utf-8')
check('标题未被显式关闭加粗', 'w:val="0"' not in _xml, _xml[:400])

# 标题与正文数量相等(众数平票): 基准取较小字号, 标题仍应被识别
tie_pdf = D('tie.pdf')
_b = pymupdf.open()
_pg = _b.new_page()
for _i in range(3):
    _pg.insert_text((72, 90 + _i * 34), f'标题{_i}文字', fontsize=22,
                    fontname='china-ss')
for _i in range(3):
    _pg.insert_text((72, 220 + _i * 24), f'正文{_i}文字', fontsize=10,
                    fontname='china-ss')
_b.save(tie_pdf)
_b.close()
r = C.convert_pdf_to_word(tie_pdf, {'out_dir': D('o_p2w')})
_d5 = docx.Document(r.output)
_styles2 = [p.style.name for p in _d5.paragraphs if p.text.strip()]
check('众数平票时标题仍被识别',
      any('Heading' in s or 'Title' in s for s in _styles2), str(_styles2))


# ============================================== 3. PDF → 图片/TXT/HTML
r = C.convert_pdf_to_images(pdf, {'out_dir': D('o_p2i')})
check('PDF→图片 成功且张数正确',
      r.status == 'ok' and '2 张' in r.detail, r.status + r.detail)
check('PDF→图片 产物可读',
      r.status == 'ok' and os.path.getsize(r.output) > 1000)

r = C.convert_pdf_to_text(pdf, {'out_dir': D('o_p2t')})
check('PDF→TXT 成功', r.status == 'ok', r.status + r.detail)
txt = open(r.output, encoding='utf-8').read()
check('PDF→TXT 含中文与分页标记',
      '正文' in txt and '第 1 页' in txt, txt[:80])

r = C.convert_pdf_to_html(pdf, {'out_dir': D('o_p2h')})
check('PDF→HTML 成功', r.status == 'ok', r.status + r.detail)
html = open(r.output, encoding='utf-8').read()
check('PDF→HTML 含标题标签与中文',
      '<h' in html and '正文' in html, html[:120])
check('PDF→HTML 识别居中段', 'class="c"' in html)


# ============================================== 4. Word → TXT/HTML
r = C.convert_word_to_text(docx_path, {'out_dir': D('o_w2t')})
check('Word→TXT 成功', r.status == 'ok', r.status + r.detail)
txt = open(r.output, encoding='utf-8').read()
check('Word→TXT 含正文与表格内容',
      '第一段正文' in txt and '张三' in txt, txt[:100])

r = C.convert_word_to_html(docx_path, {'out_dir': D('o_w2h')})
check('Word→HTML 成功', r.status == 'ok', r.status + r.detail)
html = open(r.output, encoding='utf-8').read()
check('Word→HTML 含标题/表格/加粗',
      '<h1>' in html and '<table>' in html and '<strong>' in html,
      html[:200])


# ============================================== 5. Excel → CSV/HTML
r = C.convert_excel_to_csv(xlsx_path, {'out_dir': D('o_e2c')})
check('Excel→CSV 成功', r.status == 'ok', r.status + r.detail)
csv_txt = open(r.output, encoding='utf-8-sig').read()
check('Excel→CSV 含数据行', '张三' in csv_txt and '95' in csv_txt,
      csv_txt[:80])

r = C.convert_excel_to_html(xlsx_path, {'out_dir': D('o_e2h')})
check('Excel→HTML 成功', r.status == 'ok', r.status + r.detail)
html = open(r.output, encoding='utf-8').read()
check('Excel→HTML 含全部工作表',
      '成绩表' in html and '第二表' in html, html[:150])


# ============================================== 6. PPT → TXT/图片
r = C.convert_ppt_to_text(pptx_path, {'out_dir': D('o_p2t2')})
check('PPT→TXT 成功', r.status == 'ok', r.status + r.detail)
txt = open(r.output, encoding='utf-8').read()
check('PPT→TXT 含每页文字',
      '幻灯片标题' in txt and '幻灯片正文内容' in txt, txt[:80])

r = C.convert_ppt_to_images(pptx_path, {'out_dir': D('o_p2i2')})
if HAS_PPT:
    check('PPT→图片 成功(COM 渲染)', r.status == 'ok', r.status + r.detail)
else:
    check('无 PowerPoint 时 PPT→图片 明确报错',
          r.status == 'fail' and 'PowerPoint' in r.detail,
          r.status + r.detail)


# ============================================== 7. 图片→PDF / 文本→Word
r = C.convert_image_to_pdf(img_path, {'out_dir': D('o_i2p')})
check('图片→PDF 成功', r.status == 'ok', r.status + ' ' + r.detail)
check('图片→PDF 尺寸与原图一致',
      r.status == 'ok'
      and abs(PdfReader(r.output).pages[0].mediabox.width - 300) < 2
      and abs(PdfReader(r.output).pages[0].mediabox.height - 200) < 2)

# 多帧 GIF: 每帧一页, 不能只留首帧
gif_path = D('anim.gif')
_frames = [Image.new('RGB', (60, 40), (i * 50, 0, 0)) for i in range(3)]
_frames[0].save(gif_path, save_all=True, append_images=_frames[1:])
r = C.convert_image_to_pdf(gif_path, {'out_dir': D('o_i2p')})
check('多帧 GIF→PDF 每帧一页',
      r.status == 'ok'
      and len(PdfReader(r.output).pages) == 3,
      r.status + ' ' + r.detail)

r = C.convert_text_to_word(txt_path, {'out_dir': D('o_t2w')})
check('文本→Word 成功', r.status == 'ok', r.status + r.detail)
_d3 = docx.Document(r.output)
check('文本→Word 首行成标题',
      _d3.paragraphs[0].style.name.startswith('Heading'),
      _d3.paragraphs[0].style.name)
check('文本→Word 行数完整', len(_d3.paragraphs) >= 3)


# ============================================== 7.5 新增方向(v3.1.3)
# --- PDF → Excel: 画线网格 + 词坐标归格
pdf_tbl = D('tbl.pdf')
if pymupdf is not None:
    _d = pymupdf.open()
    _pg = _d.new_page()
    _cols = [60, 160, 260, 360, 460]      # 4 列 → 5 条竖线
    _rows = [60, 100, 140, 180, 220, 260]
    for _x in _cols:
        _pg.draw_line(pymupdf.Point(_x, _rows[0]),
                      pymupdf.Point(_x, _rows[-1]))
    for _y in _rows:
        _pg.draw_line(pymupdf.Point(_cols[0], _y), pymupdf.Point(_cols[-1], _y))
    _data = [['Name', 'Dept', 'Level', 'City'],
             ['Zhang', 'Tech', 'P7', 'Beijing'],
             ['Li', 'Mkt', 'P6', 'Shanghai'],
             ['Wang', 'Tech', 'P8', 'Shenzhen']]
    for _ri, _row in enumerate(_data):
        for _ci, _cell in enumerate(_row):
            _pg.insert_text((_cols[_ci] + 5, _rows[_ri] + 28), _cell,
                            fontsize=11)
    _d.save(pdf_tbl)
    _d.close()

    r = C.convert_pdf_to_excel(pdf_tbl, {'out_dir': D('o_p2x')})
    check('PDF→Excel 成功', r.status == 'ok', r.status + ' ' + r.detail)
    if r.status == 'ok':
        _wb = openpyxl.load_workbook(r.output)
        _ws = _wb.worksheets[0]
        _vals = [[c.value for c in row] for row in _ws.iter_rows()]
        _flat = [str(c) for row in _vals for c in row if c not in (None, '')]
        check('PDF→Excel 词全归格(16/16)', len(_flat) == 16, str(_flat))
        check('PDF→Excel 表头行正确',
              _vals and _vals[0][:4] == ['Name', 'Dept', 'Level', 'City'],
              str(_vals[:1]))
        check('PDF→Excel 表头加粗', _ws['A1'].font.bold is True)

# 无表格 PDF → fail(明确报错, 不产空文件)
r = C.convert_pdf_to_excel(pdf, {'out_dir': D('o_p2x')})
check('无表格 PDF→Excel 报 fail', r.status == 'fail', r.status + ' ' + r.detail)

# --- CSV → Excel
csv_path = D('people.csv')
with open(csv_path, 'w', encoding='utf-8-sig') as _fh:
    _fh.write('姓名,部门,城市\n张三,技术部,北京\n李四,市场部,上海\n')
r = C.convert_csv_to_excel(csv_path, {'out_dir': D('o_c2x')})
check('CSV→Excel 成功', r.status == 'ok', r.status + ' ' + r.detail)
if r.status == 'ok':
    _wb = openpyxl.load_workbook(r.output)
    _ws = _wb.active
    check('CSV→Excel 行数=3', _ws.max_row == 3, str(_ws.max_row))
    check('CSV→Excel 首行冻结', _ws.freeze_panes == 'A2', str(_ws.freeze_panes))
    check('CSV→Excel BOM 不残留',
          _ws['A1'].value == '姓名', repr(_ws['A1'].value))

# TSV + 空文件边界
tsv_path = D('people.tsv')
with open(tsv_path, 'w', encoding='utf-8') as _fh:
    _fh.write('a\tb\n1\t2\n')
r = C.convert_csv_to_excel(tsv_path, {'out_dir': D('o_c2x')})
check('TSV→Excel 成功', r.status == 'ok', r.status + ' ' + r.detail)
empty_csv = D('empty.csv')
open(empty_csv, 'w').close()
r = C.convert_csv_to_excel(empty_csv, {'out_dir': D('o_c2x')})
check('空 CSV → skipped', r.status == 'skipped', r.status)

# --- Excel → Word
r = C.convert_excel_to_word(xlsx_path, {'out_dir': D('o_x2w')})
check('Excel→Word 成功', r.status == 'ok', r.status + ' ' + r.detail)
if r.status == 'ok':
    _d = docx.Document(r.output)
    check('Excel→Word 有表格', len(_d.tables) >= 1, str(len(_d.tables)))
    check('Excel→Word 表名成标题',
          any(p.style.name.startswith('Heading') for p in _d.paragraphs))
    check('Excel→Word 表格内容可编辑',
          _d.tables and _d.tables[0].rows[0].cells[0].text != '')

# --- Excel → JSON
r = C.convert_excel_to_json(xlsx_path, {'out_dir': D('o_x2j')})
check('Excel→JSON 成功', r.status == 'ok', r.status + ' ' + r.detail)
if r.status == 'ok':
    import json as _json
    with open(r.output, encoding='utf-8') as _fh:
        _j = _json.load(_fh)
    check('Excel→JSON 按表名分组', isinstance(_j, dict) and len(_j) >= 1)
    _k = list(_j)[0]
    check('Excel→JSON 按表头建对象',
          isinstance(_j[_k], list) and isinstance(_j[_k][0], dict),
          str(_j)[:200])

# --- Word → Markdown
r = C.convert_word_to_md(docx_path, {'out_dir': D('o_w2m')})
check('Word→Markdown 成功', r.status == 'ok', r.status + ' ' + r.detail)
if r.status == 'ok':
    _md = open(r.output, encoding='utf-8').read()
    check('Word→Markdown 有标题井号', '#' in _md, _md[:120])
    check('Word→Markdown 表格转竖线',
          '|' in _md, _md[:200])

# --- HTML → Word
html_path = D('page.html')
with open(html_path, 'w', encoding='utf-8') as _fh:
    _fh.write('<html><head><title>t</title><style>p{color:red}</style></head>'
              '<body><h1>大标题</h1><p>第一段<strong>加粗</strong>。</p>'
              '<h2>小标题</h2><ul><li>条目一</li><li>条目二</li></ul>'
              '<script>alert(1)</script></body></html>')
r = C.convert_html_to_word(html_path, {'out_dir': D('o_h2w')})
check('HTML→Word 成功', r.status == 'ok', r.status + ' ' + r.detail)
if r.status == 'ok':
    _d = docx.Document(r.output)
    _texts = [p.text for p in _d.paragraphs if p.text.strip()]
    _all = '\n'.join(_texts)
    check('HTML→Word 标题层级', any(
        p.style.name == 'Heading 1' for p in _d.paragraphs), str(
        [p.style.name for p in _d.paragraphs]))
    check('HTML→Word 正文在', '第一段' in _all, _all[:100])
    check('HTML→Word 样式脚本不泄漏', 'color:red' not in _all
          and 'alert' not in _all, _all[:100])

# --- 图片 → Word
r = C.convert_image_to_word(img_path, {'out_dir': D('o_i2w')})
check('图片→Word 成功', r.status == 'ok', r.status + ' ' + r.detail)
if r.status == 'ok':
    _d = docx.Document(r.output)
    _n = len(_d.inline_shapes)
    check('图片→Word 插入 1 张', _n == 1, str(_n))
    check('图片→Word 宽度不超页宽',
          _n >= 1 and _d.inline_shapes[0].width.cm <= 16.5,
          str(_d.inline_shapes[0].width.cm if _n else 'none'))
r = C.convert_image_to_word(gif_path, {'out_dir': D('o_i2w')})
check('多帧 GIF→Word 每帧一图',
      r.status == 'ok'
      and len(docx.Document(r.output).inline_shapes) == 3,
      r.status + ' ' + r.detail)

# --- PDF → PPT
r = C.convert_pdf_to_ppt(pdf, {'out_dir': D('o_p2ppt')})
check('PDF→PPT 成功', r.status == 'ok', r.status + ' ' + r.detail)
if r.status == 'ok':
    _prs = Presentation(r.output)
    check('PDF→PPT 页数一致', len(_prs.slides) == len(PdfReader(pdf).pages),
          f'{len(_prs.slides)} vs {len(PdfReader(pdf).pages)}')
    check('PDF→PPT 16:9',
          _prs.slide_width == 12192000 and _prs.slide_height == 6858000,
          f'{_prs.slide_width}x{_prs.slide_height}')
    check('PDF→PPT 每页有图',
          len(_prs.slides[0].shapes) >= 1)


# ============================================== 7.6 保真度(v3.1.4)
# PDF 源: 混排行(红色加粗 + 黑色正文) + 下划线 + 彩色大字
fid_pdf = D('fid.pdf')
if pymupdf is not None:
    _d = pymupdf.open()
    _pg = _d.new_page()
    # 第一行: 大红加粗大字
    _pg.insert_text((72, 72), 'RedBoldTitle', fontsize=22,
                    color=(0.85, 0.1, 0.1), fontname='hebo')
    # 第二行: 普通黑字
    _pg.insert_text((72, 100), 'NormalBodyText', fontsize=11)
    # 第三行: 蓝色 + 下划线(下划线是画线)
    _pg.insert_text((72, 128), 'BlueUnderline', fontsize=11,
                    color=(0.1, 0.2, 0.85))
    _pg.draw_line(pymupdf.Point(72, 131), pymupdf.Point(150, 131), width=0.8)
    _d.save(fid_pdf)
    _d.close()

    r = C.convert_pdf_to_word(fid_pdf, {'out_dir': D('o_fid')})
    check('保真 PDF→Word 成功', r.status == 'ok', r.status + ' ' + r.detail)
    if r.status == 'ok':
        _d2 = docx.Document(r.output)
        _all_runs = [rn for p in _d2.paragraphs for rn in p.runs]
        _red = [rn for rn in _all_runs
                if rn.font.color.rgb is not None
                and str(rn.font.color.rgb).upper() == 'D91A1A']
        check('PDF→Word 保留红色', len(_red) >= 1,
              str([str(rn.font.color.rgb) for rn in _all_runs]))
        check('PDF→Word 红色段同时加粗', _red and _red[0].bold is True)
        check('PDF→Word 保留大字号',
              any(rn.font.size and rn.font.size.pt >= 20 for rn in _all_runs),
              str([rn.font.size and rn.font.size.pt for rn in _all_runs]))
        _ul = [rn for rn in _all_runs if rn.font.underline]
        check('PDF→Word 保留下划线', len(_ul) >= 1, str(len(_ul)))
        check('PDF→Word 下划线段是蓝色',
              _ul and _ul[0].font.color.rgb is not None
              and str(_ul[0].font.color.rgb).upper() in ('1A33D9', '1933D9',
                                                          '1A34D9'),
              str(_ul and _ul[0].font.color.rgb))

# Word 源: 颜色/字号/下划线/高亮 → HTML / MD
wf = D('fid.docx')
_w = docx.Document()
_p1 = _w.add_paragraph()
_r1 = _p1.add_run('红字加粗')
_r1.bold = True
_r1.font.color.rgb = __import__('docx').shared.RGBColor(0xD9, 0x1A, 0x1A)
_p2 = _w.add_paragraph()
_r2 = _p2.add_run('带下划线')
_r2.font.underline = True
_r2.font.size = __import__('docx').shared.Pt(18)
_p3 = _w.add_paragraph()
_r3 = _p3.add_run('高亮文本')
_r3.font.highlight_color = __import__(
    'docx').enum.text.WD_COLOR_INDEX.YELLOW
wf2 = wf
_w.save(wf)

r = C.convert_word_to_html(wf, {'out_dir': D('o_fid')})
check('保真 Word→HTML 成功', r.status == 'ok', r.status + ' ' + r.detail)
if r.status == 'ok':
    _h = open(r.output, encoding='utf-8').read()
    check('Word→HTML 保留颜色', '#D91A1A' in _h.upper(), _h[:400])
    check('Word→HTML 保留下划线', '<u>' in _h, _h[:400])
    check('Word→HTML 保留大字号', 'font-size:18' in _h, _h[:400])
    check('Word→HTML 保留高亮', 'background' in _h.lower(), _h[:400])

r = C.convert_word_to_md(wf, {'out_dir': D('o_fid')})
check('保真 Word→MD 成功', r.status == 'ok', r.status + ' ' + r.detail)
if r.status == 'ok':
    _m = open(r.output, encoding='utf-8').read()
    check('Word→MD 保留粗体', '**红字加粗**' in _m, _m[:300])
    check('Word→MD 保留颜色(行内 HTML)', '#D91A1A' in _m.upper(), _m[:300])
    check('Word→MD 保留下划线', '<u>' in _m, _m[:300])

# Excel 源: 数字格式/颜色/列宽 → Word / HTML
ef = D('fid.xlsx')
_wb = openpyxl.Workbook()
_ws = _wb.active
_ws.append(['名称', '金额', '比率'])
_c = _ws.cell(row=2, column=1, value='测试项')
_c.font = openpyxl.styles.Font(bold=True, color='FF0000')
_c.alignment = openpyxl.styles.Alignment(horizontal='center')
_ws.cell(row=2, column=2, value=1234567.89)
_ws.cell(row=2, column=2).number_format = '#,##0.00'
_ws.cell(row=2, column=3, value=0.156)
_ws.cell(row=2, column=3).number_format = '0.0%'
_ws.column_dimensions['A'].width = 22
_ws.column_dimensions['B'].width = 14
_wb.save(ef)

r = C.convert_excel_to_word(ef, {'out_dir': D('o_fid')})
check('保真 Excel→Word 成功', r.status == 'ok', r.status + ' ' + r.detail)
if r.status == 'ok':
    _d3 = docx.Document(r.output)
    _t = _d3.tables[0]
    _txt = '\n'.join(c.text for row in _t.rows for c in row.cells)
    check('Excel→Word 千分位格式保留', '1,234,567.89' in _txt, _txt[:200])
    check('Excel→Word 百分比格式保留', '15.6%' in _txt, _txt[:200])
    _bold_cells = [rn for row in _t.rows for c in row.cells
                   for p in c.paragraphs for rn in p.runs if rn.bold]
    check('Excel→Word 保留单元格加粗', len(_bold_cells) >= 1, str(len(_bold_cells)))
    _red_cells = [rn for row in _t.rows for c in row.cells
                  for p in c.paragraphs for rn in p.runs
                  if rn.font.color.rgb is not None
                  and str(rn.font.color.rgb).upper() == 'FF0000']
    check('Excel→Word 保留单元格颜色', len(_red_cells) >= 1,
          str(len(_red_cells)))
    _w1 = _t.rows[0].cells[0].width
    _w2 = _t.rows[0].cells[1].width
    check('Excel→Word 列宽按源比例映射',
          _w1 is not None and _w2 is not None and abs(_w1 - _w2) > 1000,
          f'{_w1} vs {_w2}')

r = C.convert_excel_to_html(ef, {'out_dir': D('o_fid')})
check('保真 Excel→HTML 成功', r.status == 'ok', r.status + ' ' + r.detail)
if r.status == 'ok':
    _h = open(r.output, encoding='utf-8').read()
    check('Excel→HTML 千分位格式', '1,234,567.89' in _h, _h[:400])
    check('Excel→HTML 百分比格式', '15.6%' in _h, _h[:400])
    check('Excel→HTML 保留加粗', 'font-weight:bold' in _h, _h[:400])
    check('Excel→HTML 保留颜色', '#FF0000' in _h, _h[:400])
    check('Excel→HTML 右对齐数字', 'class="r"' in _h, _h[:400])

# PDF 表格 → Excel: 数字转真数字 + 列宽
tbl_pdf2 = D('fidtbl.pdf')
if pymupdf is not None:
    _d = pymupdf.open()
    _pg = _d.new_page()
    _cols = [60, 180, 300, 420, 540]
    _rows = [60, 100, 140, 180]
    for _x in _cols:
        _pg.draw_line(pymupdf.Point(_x, _rows[0]),
                      pymupdf.Point(_x, _rows[-1]))
    for _y in _rows:
        _pg.draw_line(pymupdf.Point(_cols[0], _y), pymupdf.Point(_cols[-1], _y))
    _data2 = [['Item', 'Qty', 'Price', 'Total'],
              ['Apple', '100', '1,234.50', '123,450'],
              ['Pear', '50', '2,000', '100,000']]
    for _ri2, _row2 in enumerate(_data2):
        for _ci2, _cell2 in enumerate(_row2):
            _pg.insert_text((_cols[_ci2] + 5, _rows[_ri2] + 28), _cell2,
                            fontsize=11)
    _d.save(tbl_pdf2)
    _d.close()

    r = C.convert_pdf_to_excel(tbl_pdf2, {'out_dir': D('o_fid')})
    check('保真 PDF→Excel 成功', r.status == 'ok', r.status + ' ' + r.detail)
    if r.status == 'ok':
        _w2 = openpyxl.load_workbook(r.output)
        _s2 = _w2.active
        _vals2 = [[c.value for c in row] for row in _s2.iter_rows()]
        _flat2 = [v for row in _vals2 for v in row]
        _nums = [v for v in _flat2 if isinstance(v, (int, float))]
        check('PDF→Excel 数字是真数字类型(可求和)', len(_nums) >= 5,
              str(_flat2))
        check('PDF→Excel 千分位文本转数字', 1234.5 in _flat2, str(_flat2))
        _wa = _s2.column_dimensions['A'].width
        check('PDF→Excel 列宽按 PDF 映射(不等宽)',
              _wa is not None and _wa > 6, str(_wa))


# ============================================== 8. 路由与守卫
r = C.convert_file(pdf, {'convert': '不存在的方向'})
check('未知转换方向 → fail 不崩溃', r.status == 'fail', r.status)

check('supported_exts(pdf2word)=(pdf,)',
      C.supported_exts('pdf2word') == ('.pdf',))
check('supported_exts(office2pdf) 含 docx',
      '.docx' in C.supported_exts('office2pdf'))

# 默认输出目录: 不指定 out_dir → 源旁「_转换结果」
r = C.convert_pdf_to_text(D('内嵌测试.pdf') if False else pdf, {})
check('默认输出到「_转换结果」',
      r.status == 'ok' and '_转换结果' in r.output, str(r.output))

# 同名文件不覆盖: 追加 (1)
r2 = C.convert_pdf_to_text(pdf, {})
check('同名输出不覆盖(追加序号)',
      r2.status == 'ok' and r.output != r2.output, r2.output)


# ============================================== 9. GUI 集成(离屏)
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
try:
    from PySide6.QtWidgets import QApplication
    import app.main as M

    _app = QApplication.instance() or QApplication([])
    win = M.MainWindow()
    win.show()
    _app.processEvents()
    conv = [p for p in win.pages if getattr(p, 'title', '') == '格式转换']
    check('格式转换页已注册', len(conv) == 1)
    if conv:
        pg = conv[0]
        pg.file_edit.setText(pdf)
        _app.processEvents()
        pg.collect_options()
        # 切换转换方向: 下拉文本同步 + info 更新
        cb = pg._widgets['op']
        cb.setCurrentText('PDF → Word')
        _app.processEvents()
        o = pg.collect_options()
        check('下拉切换后机内键正确',
              o.get('convert') == 'pdf2word', str(o))
        check('预览可用', pg.supports_preview())
        data = pg.build_preview([pdf], o)
        check('预览报告 pdf 可转换',
              data is not None and data[1][1] == '转换', str(data[:2]))
        data2 = pg.build_preview([docx_path],
                                 {'convert': 'pdf2word'})
        check('预览对不符文件报跳过',
              data2 is not None and '跳过' in data2[1][1],
              str(data2[:2]))
        # 新方向: 下拉可选 + 切换后守卫生效
        _new_keys = {'PDF → Excel': 'pdf2excel', 'CSV → Excel': 'csv2excel',
                     'Excel → Word': 'excel2word',
                     'Excel → JSON': 'excel2json',
                     'Word → Markdown': 'word2md',
                     'HTML → Word': 'html2word',
                     '图片 → Word': 'img2word', 'PDF → PPT': 'pdf2ppt'}
        _items = [cb.itemText(i) for i in range(cb.count())]
        check('8 个新方向都在下拉里',
              all(k in _items for k in _new_keys),
              str([k for k in _new_keys if k not in _items]))
        for _label, _key in _new_keys.items():
            cb.setCurrentText(_label)
            _app.processEvents()
            _o = pg.collect_options()
            if _o.get('convert') != _key:
                check(f'切换 [{_label}] 机内键正确', False, str(_o))
                break
        else:
            check('8 个新方向切换后机内键都正确', True)
        # 守卫: PDF→Excel 方向预览时给 txt 应报跳过
        cb.setCurrentText('PDF → Excel')
        _app.processEvents()
        _d3 = pg.build_preview([txt_path], pg.collect_options())
        check('新方向守卫: txt 在 PDF→Excel 下报跳过',
              _d3 is not None and '跳过' in _d3[1][1], str(_d3[:2]))
    check('nav 含格式转换入口',
          '格式转换' in win.page_index_by_title)
    from app.icons import NAV_ICONS
    check('导航图标已配置', NAV_ICons := NAV_ICONS.get('格式转换') == 'convert')
except Exception as e:
    import traceback
    traceback.print_exc()
    check('离屏 GUI 集成', False, f'{type(e).__name__}: {e}')

shutil.rmtree(BASE, ignore_errors=True)
print(f'\nRESULT: {ok} passed, {fail} failed')
sys.exit(1 if fail else 0)
