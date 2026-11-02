# -*- coding: utf-8 -*-
"""青雁批量工具箱 v3 — 格式转换引擎 (全离线, 转换结果可编辑)。

设计原则: 产出物必须是「真格式」而非截图, 保证后续可编辑。

技术路线(按优先级):
1. Office(Word/Excel/PowerPoint) → PDF
   本机 COM ExportAsFixedFormat 静默转换, 保真度最高; 无 Office 时给出
   明确失败原因, 不伪造产物。
2. PDF → Word
   PyMuPDF 提取结构化文本(字号→标题层级 / 粗体 / 段落), 用 python-docx
   重建真 .docx — 每段都是可编辑文字, 不是图片。
3. PDF → 图片 / TXT / HTML
   位图渲染 / get_text 纯文本 / 带简单排版的 HTML。
4. Word → TXT / HTML
   python-docx 直接读段落与样式。
5. Excel → CSV / HTML
   openpyxl 逐表导出(CSV 取首个工作表, HTML 全表)。
6. PPT → TXT / 图片
   python-pptx 取文本; 图片走 COM→PDF→位图渲染。
7. 图片 → PDF
   PIL 读图 + reportlab 排版为 PDF(可再转回 Word)。

所有处理器签名与 BatchWorker 约定一致: (path, options) -> FileResult。
"""
from __future__ import annotations

from engine.safety import atomic_save

import os
import shutil
import tempfile

from engine.core import FileResult
from engine.ops import TXT_EXTS, read_text_auto
from engine.office_ops import _uniq_path

# 转换方向表: (源扩展名组, 目标格式, 处理器名)
IMG_EXTS = frozenset({'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.webp',
                      '.tif', '.tiff'})
OFFICE_EXTS = frozenset({'.doc', '.docx', '.wps',
                         '.xls', '.xlsx', '.et',
                         '.ppt', '.pptx', '.dps'})


class _NoOfficeError(RuntimeError):
    pass


def _xls_to_xlsx(path: str, workdir: str) -> str:
    """旧版 .xls/.et → 临时 .xlsx(openpyxl 只认 OOXML)。

    openpyxl 打不开 .xls, 但 UI 把 .xls 列为支持格式(用户手里老表很多),
    不转换就会每个文件都失败。这里借本机 Office/WPS 另存为 xlsx。
    没有 Office 时抛 _NoOfficeError, 由调用方给出明确提示。
    """
    import pythoncom

    abs_src = os.path.abspath(path)
    out = os.path.join(workdir, 'conv.xlsx')
    basename = 'KET.Application' if os.path.splitext(path)[1].lower() == '.et' \
        else 'KET.Application'
    progs = [('Excel.Application', 'excel'), (basename, 'excel')]
    pythoncom.CoInitialize()
    last_err = None
    for prog, _kind in progs:
        app = None
        try:
            import win32com.client
            app = win32com.client.DispatchEx(prog)
        except Exception as e:
            last_err = e
            continue
        try:
            try:
                app.Visible = False
            except Exception:
                pass
            try:
                app.DisplayAlerts = 0
            except Exception:
                pass
            wb = app.Workbooks.Open(abs_src, ReadOnly=True)
            try:
                # 51 = xlOpenXMLWorkbook (.xlsx)
                wb.SaveAs(os.path.abspath(out), FileFormat=51)
            finally:
                wb.Close(False)
            return out
        except Exception as e:
            last_err = e
        finally:
            try:
                app.Quit()
            except Exception:
                pass
    raise _NoOfficeError(
        f'读取 {os.path.splitext(path)[1]} 需要本机安装 Excel 或 WPS'
        f'({last_err})')


# ============================================================ 输出路径
def _conv_out_dir(path: str, options: dict) -> str:
    """转换输出目录: 指定 out_dir 用指定; 否则源文件旁建「_转换结果」。"""
    out_dir = (options.get('out_dir') or '').strip()
    if not out_dir:
        out_dir = os.path.join(os.path.dirname(os.path.abspath(path)),
                               '_转换结果')
    os.makedirs(out_dir, exist_ok=True)
    return out_dir


def _out_path(src: str, out_dir: str, ext: str) -> str:
    stem = os.path.splitext(os.path.basename(src))[0]
    return _uniq_path(os.path.join(out_dir, stem + ext))


# ==================================================== Office → PDF (COM)
def _office_to_pdf(path: str, out_path: str) -> None:
    """本机 Office/WPS COM 静默导出 PDF。失败抛 _NoOfficeError。"""
    import time

    import pythoncom

    ext = os.path.splitext(path)[1].lower()
    if ext in ('.doc', '.docx', '.wps'):
        progs = [('Word.Application', 'word'), ('KWPS.Application', 'word')]
    elif ext in ('.xls', '.xlsx', '.et'):
        progs = [('Excel.Application', 'excel'), ('KET.Application', 'excel')]
    else:
        progs = [('PowerPoint.Application', 'ppt'),
                 ('KWPP.Application', 'ppt')]

    abspath = os.path.abspath(path)
    pythoncom.CoInitialize()
    last_err = None
    tried = []
    for prog, kind in progs:
        tried.append(prog)
        app = None
        # PowerPoint 单实例: 上一个实例 Quit 后服务器可能仍在退出,
        # 此时 DispatchEx 会误报「无效的类字符串」— 短重试一次
        for attempt in range(3):
            try:
                import win32com.client
                app = win32com.client.DispatchEx(prog)
                break
            except Exception as e:
                last_err = e
                app = None
                if attempt < 2:
                    time.sleep(0.4)
        if app is None:
            continue
        try:
            try:
                # PowerPoint 无 Visible 属性(设置会抛 com_error), 跳过;
                # Word/Excel 可隐藏。DisplayAlerts 两者都支持。
                if kind != 'ppt':
                    app.Visible = False
            except Exception:
                pass
            try:
                app.DisplayAlerts = 0
            except Exception:
                pass
            if kind == 'word':
                doc = app.Documents.Open(abspath, ReadOnly=True)
                doc.ExportAsFixedFormat(out_path, 17)      # wdExportFormatPDF
                doc.Close(False)
            elif kind == 'excel':
                wb = app.Workbooks.Open(abspath, ReadOnly=True)
                wb.ExportAsFixedFormat(0, out_path)        # xlTypePDF
                wb.Close(False)
            else:
                pres = app.Presentations.Open(abspath, True, False, False)
                # PowerPoint 的 ExportAsFixedFormat 在 pywin32 早期绑定下
                # 参数转换失败(「无法转换为 COM 对象」), 改用等价的
                # SaveAs(Path, ppSaveAsPDF=32)
                pres.SaveAs(out_path, 32)
                pres.Close()
            if os.path.exists(out_path):
                return
            last_err = RuntimeError('转换未产出文件')
        except Exception as e:
            last_err = e
        finally:
            try:
                app.Quit()
            except Exception:
                pass
    need = {'word': 'Word', 'excel': 'Excel', 'ppt': 'PowerPoint'}[
        progs[0][1]]
    raise _NoOfficeError(
        f'未检测到可用的 {need}/WPS 组件'
        f'(已尝试: {", ".join(tried)})'
        + (f' — {last_err}' if last_err else ''))


def convert_office_to_pdf(path: str, options: dict) -> FileResult:
    ext = os.path.splitext(path)[1].lower()
    if ext not in OFFICE_EXTS:
        return FileResult(path, 'skipped', f'非 Office 文件 {ext}')
    out_dir = _conv_out_dir(path, options)
    out = _out_path(path, out_dir, '.pdf')
    try:
        _office_to_pdf(path, out)
    except _NoOfficeError as e:
        return FileResult(path, 'fail', str(e))
    except Exception as e:
        return FileResult(path, 'fail', f'转换失败: {e}')
    size = os.path.getsize(out) if os.path.exists(out) else 0
    return FileResult(path, 'ok', f'→ PDF ({size / 1024:.0f} KB)', out)


# ======================================================== PDF → Word
def _span_fmt(s: dict) -> dict:
    """PDF span → Word run 格式(字号/粗斜体/颜色/字体名)。

    PyMuPDF span 字段: font(名) flags(位: 2=斜体 16=粗体) color(int RGB)。
    下划线在 PDF 里是画线不是文字属性, 由调用方按页面画线单独判断。
    """
    font = (s.get('font') or '')
    flags = s.get('flags', 0) or 0
    return {
        'size': float(s.get('size', 0) or 0),
        'bold': bool('bold' in font.lower() or (flags & 16)),
        'italic': bool('italic' in font.lower() or (flags & 2)),
        'color': int(s.get('color', 0) or 0) & 0xFFFFFF,
        'font': font,
    }


def _map_cjk_font(font: str) -> str:
    """PDF 字体名 → Word 中文字体名(保真优先, 认不出回落雅黑)。

    例: 'SimSun'→宋体 'SimHei'→黑体 'Microsoft YaHei'→微软雅黑;
    带子集前缀(ABCDEF+SimSun)的先剥掉。
    """
    f = font or ''
    if '+' in f:
        f = f.split('+', 1)[1]
    low = f.lower()
    for key, name in (
            ('simsun', '宋体'), ('simhei', '黑体'), ('yahei', '微软雅黑'),
            ('simfang', '仿宋'), ('simkai', '楷体'), ('simli', '隶书'),
            ('fangsong', '仿宋'), ('kaiti', '楷体'), ('heiti', '黑体'),
            ('songti', '宋体'), ('msyh', '微软雅黑')):
        if key in low:
            return name
    if 'bold' in low:
        return '黑体'
    return ''


def _underline_spans(page, line_bbox):
    """判断一行文字是否带下划线: 页面画线里找落在该行文字下方的短线。

    PDF 下划线是独立画线(无文字属性), 只能靠几何关系推断。
    容差: 线在文字基线下 0~4pt, 水平与文字行有重叠。
    """
    y_hi = line_bbox[3]
    x0, x1 = line_bbox[0], line_bbox[2]
    for d in page.get_drawings():
        for item in d['items']:
            if item[0] != 'l':
                continue
            p1, p2 = item[1], item[2]
            if abs(p1.y - p2.y) > 1.0:
                continue                      # 只要水平线
            ly = p1.y
            if not (y_hi - 1.0 <= ly <= y_hi + 4.0):
                continue                      # 线必须在文字下方贴身
            lx0, lx1 = min(p1.x, p2.x), max(p1.x, p2.x)
            if min(x1, lx1) - max(x0, lx0) > 2.0:
                return True
    return False


# 字号阈值 → Word 标题级别(相对正文基准动态计算, 适配不同 PDF)
def _pdf_blocks_to_docx(pdf_path: str, out_path: str) -> int:
    """PyMuPDF 提取结构化文本 → python-docx 重建可编辑 Word。

    保留: 段落分块、标题层级(按字号)、行内按 span 切 run 分别保留
    粗体/斜体/颜色/字号/字体名, 下划线按几何关系推断。
    返回写入段落数。
    """
    import pymupdf
    import docx
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE

    doc = pymupdf.open(pdf_path)
    try:
        # 先统计字号分布, 用「出现次数最多的字号」作正文基准。
        # 不能用中位数/分位: 标题与图表文字占比高时基准会被抬高,
        # 导致真正的标题反被判成正文, 标题层级全部丢失。
        sizes_all = []
        raw_blocks = []
        for page in doc:
            d = page.get_text('dict')
            for b in d.get('blocks', []):
                if b.get('type') != 0:
                    continue
                for line in b.get('lines', []):
                    spans = [s for s in line.get('spans', [])
                             if s.get('text', '').strip()]
                    if not spans:
                        continue
                    raw_blocks.append((page, b, line, spans))
                    for s in spans:
                        sizes_all.append(round(s.get('size', 0), 1))
        if not raw_blocks:
            raise ValueError('PDF 中未提取到文字(可能是扫描件/纯图片)')
        # 众数字号 = 正文(行数远多于标题); 平票时取较小值更稳
        from collections import Counter
        cnt = Counter(sizes_all)
        best = max(cnt.values())
        base = min(sz for sz, c in cnt.items() if c == best)

        d_doc = docx.Document()
        # 正文样式基准字号, 让 Word 显示接近原 PDF
        try:
            st = d_doc.styles['Normal']
            st.font.name = '微软雅黑'
            st.font.size = Pt(max(9, min(14, round(base))))
        except Exception:
            pass

        n_par = 0
        for page, b, line, spans in raw_blocks:
            text = ''.join(s['text'] for s in spans).strip()
            if not text:
                continue
            size = max(s.get('size', base) for s in spans)
            # 居中判定: 行首 x 距左右边距差 < 20pt
            x0 = line['bbox'][0]
            x1 = line['bbox'][2]
            pw = page.rect.width
            centered = (abs(x0 - (pw - x1)) < 20 and size >= base) \
                if pw else False
            underlined = _underline_spans(page, line['bbox'])

            p = d_doc.add_paragraph()
            # 标题层级: 相对基准字号的比例
            if size >= base * 1.6:
                p.style = d_doc.styles['Title'] if size >= base * 2.2 \
                    else d_doc.styles['Heading 1']
            elif size >= base * 1.25:
                p.style = d_doc.styles['Heading 2']
            # 按 span 逐个建 run: 行内混排(红色加粗标题词 + 黑色正文)
            # 只有一整行同格式时合并, 否则每个 span 独立保留自己的格式
            fmts = [_span_fmt(s) for s in spans]
            same = all(f == fmts[0] for f in fmts)
            groups = [(text, fmts[0])] if same else [
                ((s['text'] or ''), _span_fmt(s))
                for s in spans if (s['text'] or '')]
            for gtext, gf in groups:
                gtext = gtext.strip() if same else gtext
                if not gtext:
                    continue
                run = p.add_run(gtext)
                # 只在确有该格式时显式开启; 不要写 bold=False —— 那会生成
                # <w:b w:val="0"/> 把 Heading/Title 样式自带的加粗关掉
                if gf['bold']:
                    run.bold = True
                if gf['italic']:
                    run.italic = True
                if underlined:
                    run.font.underline = WD_UNDERLINE.SINGLE
                if gf['color']:
                    try:
                        run.font.color.rgb = RGBColor.from_string(
                            f'{gf["color"]:06X}')
                    except Exception:
                        pass
                # 字号: 偏离基准超过 0.5pt 才显式写, 避免污染正文
                if abs(gf['size'] - base) > 0.5 and gf['size'] > 0:
                    run.font.size = Pt(round(gf['size'], 1))
                cn_font = _map_cjk_font(gf['font'])
                if cn_font:
                    run.font.name = cn_font
            # 居中对所有段落生效: 居中的标题在 Word 里同样要居中
            if centered:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            n_par += 1
        atomic_save(d_doc.save, out_path)
        return n_par
    finally:
        doc.close()


def convert_pdf_to_word(path: str, options: dict) -> FileResult:
    if os.path.splitext(path)[1].lower() != '.pdf':
        return FileResult(path, 'skipped', '非 PDF 文件')
    out_dir = _conv_out_dir(path, options)
    out = _out_path(path, out_dir, '.docx')
    try:
        n = _pdf_blocks_to_docx(path, out)
    except ValueError as e:
        return FileResult(path, 'fail', str(e))
    except Exception as e:
        return FileResult(path, 'fail', f'PDF→Word 失败: {e}')
    return FileResult(path, 'ok', f'→ Word 可编辑文档 ({n} 段)', out)


# ============================================ PDF → 图片 / TXT / HTML
def convert_pdf_to_images(path: str, options: dict) -> FileResult:
    if os.path.splitext(path)[1].lower() != '.pdf':
        return FileResult(path, 'skipped', '非 PDF 文件')
    dpi = int(options.get('dpi', 150) or 150)
    dpi = max(72, min(300, dpi))
    out_dir = os.path.join(_conv_out_dir(path, options), '图片')
    os.makedirs(out_dir, exist_ok=True)
    try:
        import pymupdf
        doc = pymupdf.open(path)
        stem = os.path.splitext(os.path.basename(path))[0]
        mat = pymupdf.Matrix(dpi / 72, dpi / 72)
        n = 0
        first = None
        outputs = []
        for page in doc:
            pix = page.get_pixmap(matrix=mat, alpha=False)
            out = _uniq_path(os.path.join(out_dir, f'{stem}_p{page.number + 1}.png'))
            pix.save(out)
            first = first or out
            outputs.append(out)
            n += 1
        doc.close()
    except Exception as e:
        return FileResult(path, 'fail', f'PDF→图片失败: {e}')
    return FileResult(path, 'ok', f'→ {n} 张 PNG 图片', first or '', outputs=outputs)


def convert_pdf_to_text(path: str, options: dict) -> FileResult:
    if os.path.splitext(path)[1].lower() != '.pdf':
        return FileResult(path, 'skipped', '非 PDF 文件')
    out_dir = _conv_out_dir(path, options)
    out = _out_path(path, out_dir, '.txt')
    try:
        import pymupdf
        doc = pymupdf.open(path)
        parts = []
        for page in doc:
            parts.append(f'\n===== 第 {page.number + 1} 页 =====\n')
            parts.append(page.get_text('text'))
        doc.close()
        with open(out, 'w', encoding='utf-8') as fh:
            fh.write(''.join(parts))
    except Exception as e:
        return FileResult(path, 'fail', f'PDF→TXT 失败: {e}')
    return FileResult(path, 'ok', '→ TXT 文本', out)


def convert_pdf_to_html(path: str, options: dict) -> FileResult:
    if os.path.splitext(path)[1].lower() != '.pdf':
        return FileResult(path, 'skipped', '非 PDF 文件')
    out_dir = _conv_out_dir(path, options)
    out = _out_path(path, out_dir, '.html')
    try:
        import pymupdf
        from html import escape
        doc = pymupdf.open(path)
        buf = ['<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8">'
               '<title>转换结果</title><style>'
               'body{font-family:"微软雅黑",sans-serif;max-width:900px;'
               'margin:2em auto;line-height:1.7;}'
               'h1,h2,h3{font-weight:600;}'
               '.pg{border-bottom:1px dashed #ccc;padding:1em 0;}'
               '.c{text-align:center;}'
               '</style></head><body>']
        for page in doc:
            buf.append('<div class="pg">')
            d = page.get_text('dict')
            sizes = []
            lines_raw = []
            for b in d.get('blocks', []):
                if b.get('type') != 0:
                    continue
                for line in b.get('lines', []):
                    spans = [s for s in line.get('spans', [])
                             if s.get('text', '').strip()]
                    if spans:
                        lines_raw.append((line, spans))
                        sizes.extend(round(s.get('size', 0), 1)
                                     for s in spans)
            if not lines_raw:
                continue
            # 同 PDF→Word: 众数字号为正文基准
            from collections import Counter
            cnt = Counter(sizes)
            best = max(cnt.values())
            base = min(sz for sz, c in cnt.items() if c == best) or 11
            for line, spans in lines_raw:
                text = ''.join(s['text'] for s in spans).strip()
                if not text:
                    continue
                size = max(s.get('size', base) for s in spans)
                x0, _y0, x1, _y1 = line['bbox']
                centered = abs(x0 - (page.rect.width - x1)) < 20
                # cls 对所有分支生效: 居中的标题也要居中, 不能只给 <p>
                cls = ' class="c"' if centered and size >= base else ''
                # 行内按 span 切: 每个 span 保留自己的粗斜体/颜色/字号,
                # 混排行(黑色正文里嵌红色加粗词)不再被抹平成一种格式
                parts = []
                for s in spans:
                    st = (s.get('text') or '')
                    if not st:
                        continue
                    f = _span_fmt(s)
                    frag = escape(st)
                    if f['italic']:
                        frag = f'<em>{frag}</em>'
                    if f['bold']:
                        frag = f'<strong>{frag}</strong>'
                    styles = []
                    if f['color']:
                        styles.append(f'color:#{f["color"]:06X}')
                    if f['size'] and abs(f['size'] - base) > 0.5:
                        styles.append(f'font-size:{round(f["size"], 1)}pt')
                    if styles:
                        frag = f'<span style="{";".join(styles)}">{frag}</span>'
                    parts.append(frag)
                inline = ''.join(parts)
                underlined = _underline_spans(page, line['bbox'])
                if underlined:
                    inline = f'<u>{inline}</u>'
                if size >= base * 1.6:
                    tag = 'h1' if size >= base * 2.2 else 'h2'
                    buf.append(f'<{tag}{cls}>{inline}</{tag}>')
                elif size >= base * 1.25:
                    buf.append(f'<h3{cls}>{inline}</h3>')
                else:
                    buf.append(f'<p{cls}>{inline}</p>')
            buf.append('</div>')
        doc.close()
        buf.append('</body></html>')
        with open(out, 'w', encoding='utf-8') as fh:
            fh.write('\n'.join(buf))
    except Exception as e:
        return FileResult(path, 'fail', f'PDF→HTML 失败: {e}')
    return FileResult(path, 'ok', '→ HTML 网页', out)


# ================================================ Word → TXT / HTML
def convert_word_to_text(path: str, options: dict) -> FileResult:
    ext = os.path.splitext(path)[1].lower()
    if ext not in ('.docx', '.docm', '.doc', '.wps'):
        return FileResult(path, 'skipped', f'非 Word 文件 {ext}')
    out_dir = _conv_out_dir(path, options)
    out = _out_path(path, out_dir, '.txt')
    try:
        if ext in ('.docx', '.docm'):
            import docx
            d = docx.Document(path)
            parts = [p.text for p in d.paragraphs]
            for tbl in d.tables:
                for row in tbl.rows:
                    parts.append('\t'.join(c.text for c in row.cells))
            content = '\n'.join(parts)
        else:
            # .doc 二进制: 走 COM 读全文
            content = _com_read_text(path, 'word')
        with open(out, 'w', encoding='utf-8') as fh:
            fh.write(content)
    except Exception as e:
        return FileResult(path, 'fail', f'Word→TXT 失败: {e}')
    return FileResult(path, 'ok', '→ TXT 文本', out)


def _run_html(r, base_pt=None) -> str:
    """Word run → HTML 片段(粗体/斜体/删除线/下划线/颜色/字号/高亮)。

    颜色与字号只在与正文基准有差异时才写 style, 避免全文刷满内联样式。
    """
    from html import escape
    t = escape(r.text or '')
    if not t:
        return ''
    if r.font.strike:
        t = f'<del>{t}</del>'
    if r.bold:
        t = f'<strong>{t}</strong>'
    if r.italic:
        t = f'<em>{t}</em>'
    if r.font.underline:
        t = f'<u>{t}</u>'
    styles = []
    col = getattr(r.font.color, 'rgb', None)
    try:
        if col is not None and str(col) not in ('None', ''):
            styles.append(f'color:#{col}')
    except Exception:
        pass
    try:
        sz = r.font.size
        if sz is not None and base_pt and abs(float(sz.pt) - base_pt) > 0.5:
            styles.append(f'font-size:{round(float(sz.pt), 1)}pt')
        elif sz is not None and not base_pt:
            styles.append(f'font-size:{round(float(sz.pt), 1)}pt')
    except Exception:
        pass
    hl = getattr(r.font, 'highlight_color', None)
    try:
        if hl is not None and str(hl) not in ('None', 'NONE', '0'):
            styles.append(f'background-color:{WD_HL_CSS.get(str(hl), "yellow")}')
    except Exception:
        pass
    if styles:
        t = f'<span style="{";".join(styles)}">{t}</span>'
    return t


# python-docx 高亮枚举 → CSS 颜色
WD_HL_CSS = {
    'YELLOW': '#ffff00', 'BRIGHT_GREEN': '#00ff00', 'TURQUOISE': '#00ffff',
    'PINK': '#ff00ff', 'BLUE': '#0000ff', 'RED': '#ff0000',
    'DARK_BLUE': '#000080', 'TEAL': '#008080', 'GREEN': '#008000',
    'VIOLET': '#800080', 'DARK_RED': '#800000', 'DARK_YELLOW': '#808000',
    'GRAY_50': '#808080', 'GRAY_25': '#c0c0c0', 'BLACK': '#000000',
}


def _fmt_run(text: str, bold, italic, strike) -> str:
    """Word run → HTML 片段(保留粗体/斜体/删除线, 先转义)。"""
    from html import escape
    t = escape(text or '')
    if not t:
        return ''
    if strike:
        t = f'<del>{t}</del>'
    if bold:
        t = f'<strong>{t}</strong>'
    if italic:
        t = f'<em>{t}</em>'
    return t


def _runs_md(runs, base_pt=None) -> str:
    """Word run 序列 → Markdown 行内文本。

    粗体/斜体/删除线用 Markdown 语法; 颜色/字号/下划线/高亮 Markdown
    没有对应记法, 退化为行内 HTML(主流 MD 渲染器都支持), 尽量不丢格式。
    先转义 Markdown 特殊字符, 避免正文里的 * _ 被误当格式符。
    """
    out = []
    for r in runs:
        t = (r.text or '')
        if not t:
            continue
        # 转义: 只转行内格式符, 不动中文标点
        t = t.replace('\\', '\\\\').replace('*', '\\*')
        t = t.replace('_', '\\_').replace('`', '\\`')
        if r.font.strike:
            t = f'~~{t}~~'
        if r.bold:
            t = f'**{t}**'
        if r.italic:
            t = f'*{t}*'
        styles = []
        try:
            if r.font.underline:
                t = f'<u>{t}</u>'
        except Exception:
            pass
        col = getattr(r.font.color, 'rgb', None)
        try:
            if col is not None and str(col) not in ('None', ''):
                styles.append(f'color:#{col}')
        except Exception:
            pass
        try:
            sz = r.font.size
            if sz is not None and base_pt and abs(float(sz.pt) - base_pt) > 0.5:
                styles.append(f'font-size:{round(float(sz.pt), 1)}pt')
        except Exception:
            pass
        hl = getattr(r.font, 'highlight_color', None)
        try:
            if hl is not None and str(hl) not in ('None', 'NONE', '0'):
                styles.append(
                    f'background-color:{WD_HL_CSS.get(str(hl), "yellow")}')
        except Exception:
            pass
        if styles:
            t = f'<span style="{";".join(styles)}">{t}</span>'
        out.append(t)
    return ''.join(out)


def convert_word_to_html(path: str, options: dict) -> FileResult:
    ext = os.path.splitext(path)[1].lower()
    if ext not in ('.docx', '.docm'):
        return FileResult(path, 'skipped',
                          f'Word→HTML 仅支持 .docx(收到 {ext}); '
                          '旧版 .doc 可先转 PDF 或另存为 docx')
    out_dir = _conv_out_dir(path, options)
    out = _out_path(path, out_dir, '.html')
    try:
        import docx
        from html import escape
        d = docx.Document(path)
        # 正文基准字号(用于只标差异, 不给全文刷内联样式)
        try:
            base_pt = float(d.styles['Normal'].font.size.pt)
        except Exception:
            base_pt = 11.0
        buf = ['<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8">'
               '<title>' + escape(os.path.basename(path)) +
               '</title><style>body{font-family:"微软雅黑",sans-serif;'
               'max-width:860px;margin:2em auto;line-height:1.8;}'
               'table{border-collapse:collapse;width:100%;}'
               'td,th{border:1px solid #bbb;padding:6px 10px;}'
               '</style></head><body>']
        for p in d.paragraphs:
            # run 级格式: 粗体/斜体/删除线/下划线/颜色/字号/高亮逐段拼装
            html_text = ''.join(_run_html(r, base_pt)
                                for r in p.runs) if p.runs else escape(p.text)
            text = html_text
            if not (text and text.strip()):
                buf.append('<p>&nbsp;</p>')
                continue
            style = (p.style.name or '').lower()
            if style.startswith('title'):
                buf.append(f'<h1>{text}</h1>')
            elif style.startswith('heading 1'):
                buf.append(f'<h1>{text}</h1>')
            elif style.startswith('heading 2'):
                buf.append(f'<h2>{text}</h2>')
            elif style.startswith('heading 3'):
                buf.append(f'<h3>{text}</h3>')
            elif 'list' in style:
                buf.append(f'<li>{text}</li>')
            else:
                buf.append(f'<p>{text}</p>')
        for tbl in d.tables:
            buf.append('<table>')
            for row in tbl.rows:
                buf.append('<tr>' + ''.join(
                    f'<td>{escape(c.text)}</td>' for c in row.cells) + '</tr>')
            buf.append('</table>')
        buf.append('</body></html>')
        with open(out, 'w', encoding='utf-8') as fh:
            fh.write('\n'.join(buf))
    except Exception as e:
        return FileResult(path, 'fail', f'Word→HTML 失败: {e}')
    return FileResult(path, 'ok', '→ HTML 网页', out)


# ============================================ Excel → CSV / HTML
def convert_excel_to_csv(path: str, options: dict) -> FileResult:
    ext = os.path.splitext(path)[1].lower()
    if ext not in ('.xlsx', '.xlsm', '.xls', '.et'):
        return FileResult(path, 'skipped', f'非 Excel 文件 {ext}')
    out_dir = _conv_out_dir(path, options)
    out = _out_path(path, out_dir, '.csv')
    tmpdir = None
    try:
        import csv
        import openpyxl
        src = path
        if ext in ('.xls', '.et'):
            # openpyxl 只认 OOXML, 旧格式先用 Office/WPS 另存为 xlsx
            tmpdir = tempfile.mkdtemp(prefix='qy_xls_')
            src = _xls_to_xlsx(path, tmpdir)
        wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
        ws = wb.worksheets[0]
        with open(out, 'w', encoding='utf-8-sig', newline='') as fh:
            w = csv.writer(fh)
            n = 0
            for row in ws.iter_rows(values_only=True):
                w.writerow(['' if c is None else c for c in row])
                n += 1
        wb.close()
        if n == 0:
            os.remove(out)
            return FileResult(path, 'skipped', '首表无数据')
    except _NoOfficeError as e:
        return FileResult(path, 'fail', str(e))
    except Exception as e:
        return FileResult(path, 'fail', f'Excel→CSV 失败: {e}')
    finally:
        if tmpdir:
            shutil.rmtree(tmpdir, ignore_errors=True)
    return FileResult(path, 'ok', f'→ CSV ({n} 行)', out)


def convert_excel_to_html(path: str, options: dict) -> FileResult:
    ext = os.path.splitext(path)[1].lower()
    if ext not in ('.xlsx', '.xlsm', '.xls', '.et'):
        return FileResult(path, 'skipped', f'非 Excel 文件 {ext}')
    out_dir = _conv_out_dir(path, options)
    out = _out_path(path, out_dir, '.html')
    tmpdir = None
    try:
        import datetime as _dt
        import openpyxl
        from html import escape
        src = path
        if ext in ('.xls', '.et'):
            tmpdir = tempfile.mkdtemp(prefix='qy_xls_')
            src = _xls_to_xlsx(path, tmpdir)
        wb = openpyxl.load_workbook(src, data_only=True)
        buf = ['<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8">'
               '<title>' + escape(os.path.basename(path)) +
               '</title><style>body{font-family:"微软雅黑",sans-serif;'
               'margin:2em;}table{border-collapse:collapse;margin-bottom:2em;}'
               'td,th{border:1px solid #bbb;padding:5px 9px;font-size:13px;}'
               'th{background:#f0f4f8;}h3{margin-bottom:.3em;}'
               'td.r{text-align:right;}td.c{text-align:center;}'
               '</style></head><body>']

        def _cell_html(c):
            """单元格 → <td>…</td>, 保留粗斜/颜色/底色/对齐/数字格式。"""
            v = c.value
            if v is None:
                return '<td></td>'
            if isinstance(v, (_dt.datetime, _dt.date)):
                txt = (v.strftime('%Y-%m-%d %H:%M:%S')
                       if isinstance(v, _dt.datetime)
                       else v.strftime('%Y-%m-%d'))
                cls, style = ' class="r"', ''
            elif isinstance(v, bool):
                txt, cls, style = ('TRUE' if v else 'FALSE'), '', ''
            elif isinstance(v, (int, float)):
                nf = c.number_format or ''
                cls = ' class="r"'
                if '%' in nf:
                    txt = f'{v * 100:g}%'
                elif '¥' in nf or '$' in nf:
                    txt = ('¥' if '¥' in nf else '$') + f'{v:,.2f}'
                elif nf.startswith('#,##0'):
                    txt = f'{v:,.{2 if "0.00" in nf else 0}f}'
                else:
                    txt = f'{v:g}'
                style = ''
            else:
                txt, cls, style = str(v), '', ''
            styles = []
            f = c.font
            try:
                if f.bold:
                    styles.append('font-weight:bold')
                if f.italic:
                    styles.append('font-style:italic')
                rgb = getattr(f.color, 'rgb', None)
                if isinstance(rgb, str) and len(rgb) >= 6:
                    styles.append(f'color:#{rgb[-6:]}')
            except Exception:
                pass
            try:
                fill = c.fill
                fg = getattr(fill, 'fgColor', None)
                rgb = getattr(fg, 'rgb', None)
                if fill is not None and str(
                        getattr(fill, 'patternType', '')) != 'none' \
                        and isinstance(rgb, str) and len(rgb) >= 6 \
                        and rgb[-6:].upper() not in ('000000', 'FFFFFF'):
                    styles.append(f'background:#{rgb[-6:]}')
            except Exception:
                pass
            try:
                al = c.alignment
                h = str(getattr(al, 'horizontal', '') or '')
                if h in ('center', 'right') and not cls:
                    cls = f' class="{h[0]}"'
            except Exception:
                pass
            if styles:
                style = f' style="{";".join(styles)}"'
            return f'<td{cls}{style}>{escape(txt)}</td>'

        for ws in wb.worksheets:
            buf.append(f'<h3>{escape(ws.title)}</h3><table>')
            for row in ws.iter_rows():
                buf.append('<tr>' + ''.join(_cell_html(c) for c in row)
                           + '</tr>')
            buf.append('</table>')
        wb.close()
        buf.append('</body></html>')
        with open(out, 'w', encoding='utf-8') as fh:
            fh.write('\n'.join(buf))
    except _NoOfficeError as e:
        return FileResult(path, 'fail', str(e))
    except Exception as e:
        return FileResult(path, 'fail', f'Excel→HTML 失败: {e}')
    finally:
        if tmpdir:
            shutil.rmtree(tmpdir, ignore_errors=True)
    return FileResult(path, 'ok', '→ HTML 网页(单元格格式保留)', out)


# ============================================ PPT → TXT / 图片
def convert_ppt_to_text(path: str, options: dict) -> FileResult:
    ext = os.path.splitext(path)[1].lower()
    if ext not in ('.pptx', '.pptm', '.ppt', '.dps'):
        return FileResult(path, 'skipped', f'非 PPT 文件 {ext}')
    out_dir = _conv_out_dir(path, options)
    out = _out_path(path, out_dir, '.txt')
    try:
        if ext in ('.pptx', '.pptm'):
            from pptx import Presentation
            prs = Presentation(path)
            parts = []
            for i, slide in enumerate(prs.slides, 1):
                parts.append(f'===== 第 {i} 页 =====')
                for sh in slide.shapes:
                    if sh.has_text_frame:
                        t = sh.text_frame.text.strip()
                        if t:
                            parts.append(t)
            content = '\n'.join(parts)
        else:
            content = _com_read_text(path, 'ppt')
        with open(out, 'w', encoding='utf-8') as fh:
            fh.write(content)
    except Exception as e:
        return FileResult(path, 'fail', f'PPT→TXT 失败: {e}')
    return FileResult(path, 'ok', '→ TXT 文本', out)


def convert_ppt_to_images(path: str, options: dict) -> FileResult:
    ext = os.path.splitext(path)[1].lower()
    if ext not in ('.pptx', '.pptm', '.ppt', '.dps'):
        return FileResult(path, 'skipped', f'非 PPT 文件 {ext}')
    dpi = max(72, min(200, int(options.get('dpi', 120) or 120)))
    out_dir = os.path.join(_conv_out_dir(path, options), '图片')
    os.makedirs(out_dir, exist_ok=True)
    tmp = tempfile.mkdtemp(prefix='qyconv_')
    try:
        try:
            pdf_tmp = os.path.join(tmp, 'slide.pdf')
            _office_to_pdf(path, pdf_tmp)
        except _NoOfficeError as e:
            return FileResult(path, 'fail', str(e) + '(PPT 转图片需要 Office)')
        import pymupdf
        doc = pymupdf.open(pdf_tmp)
        stem = os.path.splitext(os.path.basename(path))[0]
        mat = pymupdf.Matrix(dpi / 72, dpi / 72)
        n = 0
        first = None
        outputs = []
        for page in doc:
            pix = page.get_pixmap(matrix=mat, alpha=False)
            out = _uniq_path(
                os.path.join(out_dir, f'{stem}_p{page.number + 1}.png'))
            pix.save(out)
            first = first or out
            outputs.append(out)
            n += 1
        doc.close()
    except Exception as e:
        return FileResult(path, 'fail', f'PPT→图片失败: {e}')
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return FileResult(path, 'ok', f'→ {n} 张 PNG 图片', first or '', outputs=outputs)


# ============================================ 图片 → PDF
def convert_image_to_pdf(path: str, options: dict) -> FileResult:
    ext = os.path.splitext(path)[1].lower()
    if ext not in IMG_EXTS:
        return FileResult(path, 'skipped', f'非图片文件 {ext}')
    out_dir = _conv_out_dir(path, options)
    out = _out_path(path, out_dir, '.pdf')
    try:
        from reportlab.lib.utils import ImageReader
        from reportlab.pdfgen import canvas as rl_canvas
        from PIL import Image

        with Image.open(path) as im:
            # 动图(GIF)与多帧 TIFF: 每帧一页, 否则只保留首帧会丢内容
            frames = []
            try:
                n = getattr(im, 'n_frames', 1)
            except Exception:
                n = 1
            for i in range(max(1, n)):
                im.seek(i)
                frames.append(im.convert('RGB'))
        c = rl_canvas.Canvas(out)
        n_pages = 0
        for fr in frames:
            w, h = fr.size
            c.setPageSize((w, h))
            c.drawImage(ImageReader(fr), 0, 0, width=w, height=h)
            c.showPage()
            n_pages += 1
        c.save()
    except Exception as e:
        return FileResult(path, 'fail', f'图片→PDF 失败: {e}')
    return FileResult(path, 'ok',
                      f'→ PDF(原图尺寸, {n_pages} 页)', out)


# ============================================ 文本 → Word (HTML→PDF 类似)
def convert_text_to_word(path: str, options: dict) -> FileResult:
    ext = os.path.splitext(path)[1].lower()
    if ext not in TXT_EXTS and ext not in ('.md', '.ini', '.yaml', '.yml'):
        return FileResult(path, 'skipped', f'非文本文件 {ext}')
    out_dir = _conv_out_dir(path, options)
    out = _out_path(path, out_dir, '.docx')
    try:
        import docx
        text = read_text_auto(path)
        d = docx.Document()
        try:
            d.styles['Normal'].font.name = '微软雅黑'
        except Exception:
            pass
        first = True
        for line in text.splitlines():
            p = d.add_paragraph()
            p.add_run(line.rstrip('\r'))
            if first and line.strip():
                # 首行非空 → 设为标题, 便于生成目录
                try:
                    p.style = d.styles['Heading 1']
                except Exception:
                    pass
                first = False
        atomic_save(d.save, out)
    except Exception as e:
        return FileResult(path, 'fail', f'文本→Word 失败: {e}')
    return FileResult(path, 'ok', '→ Word 可编辑文档', out)


# ============================================ COM 读全文(旧版二进制)
def _com_read_text(path: str, kind: str) -> str:
    """旧版 .doc/.ppt 二进制: 用 COM 打开读全文。"""
    import pythoncom

    prog = {'word': 'Word.Application', 'ppt': 'PowerPoint.Application'}[kind]
    abspath = os.path.abspath(path)
    pythoncom.CoInitialize()
    import win32com.client
    app = win32com.client.DispatchEx(prog)
    try:
        app.Visible = False
        app.DisplayAlerts = 0
        if kind == 'word':
            doc = app.Documents.Open(abspath, ReadOnly=True)
            text = doc.Content.Text
            doc.Close(False)
        else:
            pres = app.Presentations.Open(abspath, True, False, False)
            parts = []
            for i, slide in enumerate(pres.Slides, 1):
                parts.append(f'===== 第 {i} 页 =====')
                for sh in slide.Shapes:
                    try:
                        if sh.HasTextFrame:
                            parts.append(sh.TextFrame.TextRange.Text)
                    except Exception:
                        continue
            text = '\n'.join(parts)
            pres.Close()
        return text
    finally:
        try:
            app.Quit()
        except Exception:
            pass


# ============================================================ 统一入口
# ==================================================== PDF → Excel(表格)
def _grid_from_lines(page):
    """从页面画线求表格的行列边界(竖线 x 集 + 横线 y 集)。

    不用 pymupdf 的 table.extract(): 实测它对单元格文本归位有缺陷
    (词被丢列/串格)。这里自己按词坐标归格, 12 词对 12 格一个不丢。
    """
    vxs, hys = [], []
    for d in page.get_drawings():
        for item in d['items']:
            if item[0] == 'l':
                p1, p2 = item[1], item[2]
                if abs(p1.x - p2.x) < 1.0:
                    vxs.append(round(p1.x, 1))
                elif abs(p1.y - p2.y) < 1.0:
                    hys.append(round(p1.y, 1))
            elif item[0] == 're':
                r = item[1]
                vxs.extend((round(r.x0, 1), round(r.x1, 1)))
                hys.extend((round(r.y0, 1), round(r.y1, 1)))
    return sorted(set(vxs)), sorted(set(hys))


def _grid_fill(page, vxs, hys):
    """把页面上的词按坐标填进网格; 返回二维列表(空串=空格)。"""
    if len(vxs) < 2 or len(hys) < 2:
        return []
    grid = [['' for _ in range(len(vxs) - 1)]
            for _ in range(len(hys) - 1)]
    for w in page.get_text('words'):
        # 词左边缘定列, 垂直中线定行 — 与肉眼视觉归属一致
        wx = w[0]
        wy = (w[1] + w[3]) / 2
        if not (vxs[0] <= wx <= vxs[-1] and hys[0] <= wy <= hys[-1]):
            continue
        ci = min(max(sum(1 for x in vxs if x <= wx) - 1, 0),
                 len(grid[0]) - 1)
        ri = min(max(sum(1 for y in hys if y <= wy) - 1, 0),
                 len(grid) - 1)
        grid[ri][ci] = (grid[ri][ci] + ' ' + w[4]).strip()
    return grid


def _dedup_lines(grid):
    """去掉全空的头尾行, 并压缩中间连续空行(只留一行)。"""
    rows = [r for r in grid if any(c.strip() for c in r)]
    return rows


def _num_cell(text: str):
    """单元格文本 → 数字(可解析时), 否则原样返回字符串。

    保真: PDF 表格里的 1234.5 写成 Excel 文本就没法求和了, 必须转真数字。
    带千分位/货币符号/百分号/括号负数的都认; 认不出就保持字符串。
    """
    t = (text or '').strip()
    if not t:
        return t
    neg = t.startswith('(') and t.endswith(')')
    body = t[1:-1] if neg else t
    body = body.replace(',', '').replace('，', '')
    pct = body.endswith('%')
    if pct:
        body = body[:-1]
    for sym in ('¥', '$', '￥'):
        if body.startswith(sym):
            body = body[len(sym):]
    try:
        v = float(body)
    except (TypeError, ValueError):
        return t
    if pct:
        v = v / 100.0
    if neg:
        v = -v
    # 整数形态回落 int: 100 不要写成 100.0
    if v == int(v) and abs(v) < 1e15 and '.' not in body:
        v = int(v)
    return v


def _pdf_tables_to_xlsx(pdf_path: str, out_path: str) -> tuple:
    """PDF 表格 → xlsx。返回 (表格数, 总行数)。无表格抛 ValueError。

    保真项: 列宽按 PDF 列边界等比映射(不是默认等宽),
    数字单元格转真数字(可求和), 首行表头加粗。
    """
    import openpyxl
    import pymupdf

    doc = pymupdf.open(pdf_path)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    n_tbl = 0
    n_rows = 0
    try:
        for page in doc:
            vxs, hys = _grid_from_lines(page)
            grid = _grid_fill(page, vxs, hys)
            if not grid:
                continue
            grid = _dedup_lines(grid)
            if not grid:
                continue
            ws = wb.create_sheet(title=f'第{page.number + 1}页')
            for row in grid:
                ws.append([_num_cell(c) for c in row])
                n_rows += 1
            # 列宽: PDF 列宽(pt) → Excel 字符宽(约 1pt ≈ 0.11 字符)
            for i in range(len(vxs) - 1):
                w_pt = vxs[i + 1] - vxs[i]
                ws.column_dimensions[
                    openpyxl.utils.get_column_letter(i + 1)].width = max(
                        6, round(w_pt * 0.11, 1))
            # 首行加粗(多半是表头)
            if ws.max_row >= 1:
                for c in ws[1]:
                    c.font = openpyxl.styles.Font(bold=True)
                ws.freeze_panes = 'A2'
            n_tbl += 1
    finally:
        doc.close()
    if n_tbl == 0:
        raise ValueError('PDF 中未检测到带边框线的表格'
                         '(纯文字或无线表格暂不支持)')
    atomic_save(wb.save, out_path)
    return n_tbl, n_rows


def convert_pdf_to_excel(path: str, options: dict) -> FileResult:
    if os.path.splitext(path)[1].lower() != '.pdf':
        return FileResult(path, 'skipped', '非 PDF 文件')
    out_dir = _conv_out_dir(path, options)
    out = _out_path(path, out_dir, '.xlsx')
    try:
        n_tbl, n_rows = _pdf_tables_to_xlsx(path, out)
    except ValueError as e:
        return FileResult(path, 'fail', str(e))
    except Exception as e:
        return FileResult(path, 'fail', f'PDF→Excel 失败: {e}')
    return FileResult(path, 'ok',
                      f'→ Excel({n_tbl} 个表格, {n_rows} 行, 可编辑)', out)


# ==================================================== CSV → Excel
def convert_csv_to_excel(path: str, options: dict) -> FileResult:
    ext = os.path.splitext(path)[1].lower()
    if ext not in ('.csv', '.tsv', '.txt'):
        return FileResult(path, 'skipped', f'非 CSV 文件 {ext}')
    out_dir = _conv_out_dir(path, options)
    out = _out_path(path, out_dir, '.xlsx')
    try:
        import csv
        import openpyxl
        delim = '\t' if ext == '.tsv' else ','
        # 编码自动探测: csv 常是 GBK/UTF-8-BOM
        text = read_text_auto(path)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = os.path.splitext(os.path.basename(path))[0][:31] or 'Sheet1'
        n = 0
        for row in csv.reader(text.splitlines(), delimiter=delim):
            if not row:
                continue
            ws.append(row)
            n += 1
        if n == 0:
            return FileResult(path, 'skipped', 'CSV 无数据行')
        for c in ws[1]:
            c.font = openpyxl.styles.Font(bold=True)
        ws.freeze_panes = 'A2'
        atomic_save(wb.save, out)
    except Exception as e:
        return FileResult(path, 'fail', f'CSV→Excel 失败: {e}')
    return FileResult(path, 'ok', f'→ Excel({n} 行, 表头已冻结)', out)


# ==================================================== Excel → Word
def convert_excel_to_word(path: str, options: dict) -> FileResult:
    ext = os.path.splitext(path)[1].lower()
    if ext not in ('.xlsx', '.xlsm', '.xls', '.et'):
        return FileResult(path, 'skipped', f'非 Excel 文件 {ext}')
    out_dir = _conv_out_dir(path, options)
    out = _out_path(path, out_dir, '.docx')
    tmpdir = None
    try:
        import datetime as _dt
        import docx
        import openpyxl
        from docx.enum.table import WD_TABLE_ALIGNMENT
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Cm, Pt, RGBColor

        src = path
        if ext in ('.xls', '.et'):
            tmpdir = tempfile.mkdtemp(prefix='qy_xls_')
            src = _xls_to_xlsx(path, tmpdir)

        # 保真必须读单元格对象(values_only 拿不到字体/颜色/对齐),
        # 且要非只读模式才有 style 信息
        wb = openpyxl.load_workbook(src, data_only=True)
        d = docx.Document()
        try:
            d.styles['Normal'].font.name = '微软雅黑'
            d.styles['Normal'].font.size = Pt(10)
        except Exception:
            pass

        def _fmt_cell_val(v, num_fmt):
            """单元格值 → (文本, 右对齐?) 。日期按 Excel 显示格式走。"""
            if v is None:
                return '', False
            if isinstance(v, (_dt.datetime, _dt.date)):
                if isinstance(v, _dt.datetime):
                    return v.strftime('%Y-%m-%d %H:%M:%S'), True
                return v.strftime('%Y-%m-%d'), True
            if isinstance(v, bool):
                return ('TRUE' if v else 'FALSE'), False
            if isinstance(v, (int, float)):
                # 货币/百分比格式按 Excel 显示还原, 否则 1234.5 会变 1234.500
                nf = (num_fmt or '')
                if '%' in nf:
                    return f'{v * 100:g}%', True
                if '¥' in nf or '$' in nf:
                    sym = '¥' if '¥' in nf else '$'
                    return f'{sym}{v:,.2f}', True
                if nf.startswith('#,##0'):
                    dec = 2 if '0.00' in nf else 0
                    return f'{v:,.{dec}f}', True
                return (f'{v:g}', True)
            return str(v), False

        def _xl_color(c):
            """openpyxl 颜色对象 → RGBColor 或 None。"""
            rgb = getattr(c, 'rgb', None)
            if isinstance(rgb, str) and len(rgb) >= 6:
                h = rgb[-6:]
                try:
                    return RGBColor.from_string(h)
                except Exception:
                    return None
            return None

        def _apply_to_par(par, cell):
            """把 Excel 单元格格式搬到 Word 段落 run 上(par.runs 已非空)。"""
            if cell is None or not par.runs:
                return
            run = par.runs[0]
            f = cell.font
            try:
                if f.bold:
                    run.bold = True
                if f.italic:
                    run.italic = True
                if f.underline:
                    run.font.underline = True
                sz = f.size
                if sz:
                    run.font.size = Pt(round(float(sz.pt), 1))
                col = _xl_color(f.color)
                if col is not None:
                    run.font.color.rgb = col
            except Exception:
                pass

        for ws in wb.worksheets:
            d.add_heading(ws.title, level=1)
            tbl = None
            for ri, row in enumerate(ws.iter_rows(), 1):
                vals, fmts = [], []
                for c in row:
                    txt, right = _fmt_cell_val(c.value,
                                              getattr(c, 'number_format', None))
                    vals.append(txt)
                    fmts.append((c, right))
                if not any(v.strip() for v in vals):
                    continue
                if tbl is None:
                    tbl = d.add_table(rows=1, cols=len(vals))
                    tbl.style = 'Table Grid'
                    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                    # 列宽: 按 Excel 列宽等比映射到 16cm 可用宽
                    widths = []
                    for i in range(len(vals)):
                        try:
                            cd = ws.column_dimensions[
                                openpyxl.utils.get_column_letter(i + 1)]
                            widths.append(float(cd.width or 9))
                        except Exception:
                            widths.append(9.0)
                    tot = sum(widths) or 1.0
                    scale = 16.0 / tot
                    for i, w in enumerate(widths):
                        for r in tbl.rows:
                            r.cells[i].width = Cm(max(0.8, w * scale))
                cells = tbl.rows[0].cells if ri == 1 else tbl.add_row().cells
                for i, (txt, right) in enumerate(fmts):
                    if i >= len(cells):
                        break
                    cell_obj = row[i] if i < len(row) else None
                    cells[i].text = ''
                    par = cells[i].paragraphs[0]
                    par.add_run(vals[i])
                    if right:
                        par.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    if ri == 1:
                        par.runs[0].bold = True        # 表头行加粗
                    _apply_to_par(par, cell_obj)
            if tbl is None:
                d.add_paragraph('(本工作表无数据)')
        wb.close()
        atomic_save(d.save, out)
    except _NoOfficeError as e:
        return FileResult(path, 'fail', str(e))
    except Exception as e:
        return FileResult(path, 'fail', f'Excel→Word 失败: {e}')
    finally:
        if tmpdir:
            shutil.rmtree(tmpdir, ignore_errors=True)
    return FileResult(path, 'ok', '→ Word(单元格格式保留, 可编辑)', out)


# ==================================================== Excel → JSON
def convert_excel_to_json(path: str, options: dict) -> FileResult:
    ext = os.path.splitext(path)[1].lower()
    if ext not in ('.xlsx', '.xlsm', '.xls', '.et'):
        return FileResult(path, 'skipped', f'非 Excel 文件 {ext}')
    out_dir = _conv_out_dir(path, options)
    out = _out_path(path, out_dir, '.json')
    tmpdir = None
    try:
        import json as _json
        import openpyxl
        src = path
        if ext in ('.xls', '.et'):
            tmpdir = tempfile.mkdtemp(prefix='qy_xls_')
            src = _xls_to_xlsx(path, tmpdir)
        wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
        data = {}
        for ws in wb.worksheets:
            rows = [['' if c is None else c for c in row]
                    for row in ws.iter_rows(values_only=True)]
            rows = [r for r in rows if any(str(c).strip() for c in r)]
            if not rows:
                data[ws.title] = []
                continue
            header = [str(h) if str(h).strip() else f'列{i + 1}'
                      for i, h in enumerate(rows[0])]
            objs = []
            for r in rows[1:]:
                # 列数多于表头时补键, 少于则留空, 保证对象结构一致
                padded = list(r) + [''] * (len(header) - len(r))
                objs.append({header[i]: (padded[i] if i < len(padded) else '')
                             for i in range(len(header))})
            data[ws.title] = objs
        wb.close()
        with open(out, 'w', encoding='utf-8') as fh:
            _json.dump(data, fh, ensure_ascii=False, indent=2)
    except _NoOfficeError as e:
        return FileResult(path, 'fail', str(e))
    except Exception as e:
        return FileResult(path, 'fail', f'Excel→JSON 失败: {e}')
    finally:
        if tmpdir:
            shutil.rmtree(tmpdir, ignore_errors=True)
    return FileResult(path, 'ok',
                      f'→ JSON({len(data)} 个工作表, 按表头建对象)', out)


# ==================================================== Word → Markdown
def convert_word_to_md(path: str, options: dict) -> FileResult:
    ext = os.path.splitext(path)[1].lower()
    if ext not in ('.docx', '.docm'):
        return FileResult(path, 'skipped',
                          f'Word→Markdown 仅支持 .docx(收到 {ext})')
    out_dir = _conv_out_dir(path, options)
    out = _out_path(path, out_dir, '.md')
    try:
        import docx
        from engine.convert_ops import _runs_md
        d = docx.Document(path)
        try:
            base_pt = float(d.styles['Normal'].font.size.pt)
        except Exception:
            base_pt = 11.0
        buf = []
        for p in d.paragraphs:
            style = (p.style.name or '').lower()
            text = _runs_md(p.runs, base_pt) if p.runs else (p.text or '')
            if not text.strip():
                buf.append('')
                continue
            if style.startswith('title'):
                buf.append(f'# {text}')
            elif style.startswith('heading 1'):
                buf.append(f'# {text}')
            elif style.startswith('heading 2'):
                buf.append(f'## {text}')
            elif style.startswith('heading 3'):
                buf.append(f'### {text}')
            elif style.startswith('heading 4'):
                buf.append(f'#### {text}')
            elif 'list number' in style:
                buf.append(f'1. {text}')
            elif 'list' in style:
                buf.append(f'- {text}')
            elif 'quote' in style:
                buf.append(f'> {text}')
            else:
                buf.append(text)
        for tbl in d.tables:
            buf.append('')
            for ri, row in enumerate(tbl.rows):
                cells = [c.text.replace('\n', ' ').strip() for c in row.cells]
                buf.append('| ' + ' | '.join(cells) + ' |')
                if ri == 0:
                    buf.append('|' + ' --- |' * len(cells))
            buf.append('')
        with open(out, 'w', encoding='utf-8') as fh:
            fh.write('\n'.join(buf).strip() + '\n')
    except Exception as e:
        return FileResult(path, 'fail', f'Word→Markdown 失败: {e}')
    return FileResult(path, 'ok', '→ Markdown 文档', out)


# ==================================================== HTML → Word
def convert_html_to_word(path: str, options: dict) -> FileResult:
    ext = os.path.splitext(path)[1].lower()
    if ext not in ('.html', '.htm'):
        return FileResult(path, 'skipped', f'非 HTML 文件 {ext}')
    out_dir = _conv_out_dir(path, options)
    out = _out_path(path, out_dir, '.docx')
    try:
        import docx
        from html.parser import HTMLParser

        class _P(HTMLParser):
            """HTML → (样式名, 文本) 序列; 只认块级标签, 不执行任何脚本。"""

            SKIP = {'script', 'style', 'head', 'meta', 'link', 'title'}

            def __init__(self):
                super().__init__(convert_charrefs=True)
                self.blocks = []
                self.cur = []
                self.skip_depth = 0
                self.style = 'Normal'

            def handle_starttag(self, tag, attrs):
                if tag in self.SKIP:
                    self.skip_depth += 1
                    return
                if self.skip_depth:
                    return
                if tag in ('h1', 'h2', 'h3', 'h4'):
                    self._flush()
                    self.style = 'Heading ' + tag[1]
                elif tag in ('h5', 'h6'):
                    self._flush()
                    self.style = 'Heading 3'
                elif tag == 'p':
                    self._flush()
                    self.style = 'Normal'
                elif tag == 'li':
                    self._flush()
                    self.style = 'List Bullet'
                elif tag in ('strong', 'b'):
                    self.cur.append(('bold', None))

            def handle_endtag(self, tag):
                if tag in self.SKIP:
                    self.skip_depth = max(0, self.skip_depth - 1)
                    return
                if tag in ('strong', 'b'):
                    self.cur.append(('/bold', None))
                elif tag in ('p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li'):
                    self._flush()

            def handle_data(self, data):
                # data 可能为 None(某些 CDATA/注释边界), 直接拼会 TypeError
                if not self.skip_depth and data:
                    self.cur.append(('text', data))

            def _flush(self):
                text = ''.join(v for k, v in self.cur
                               if v and k in ('text', 'bold', '/bold')
                               ).strip()
                if text:
                    self.blocks.append((self.style, text))
                self.cur = []
                self.style = 'Normal'

        p = _P()
        p.feed(read_text_auto(path))
        p.close()
        p._flush()
        d = docx.Document()
        try:
            d.styles['Normal'].font.name = '微软雅黑'
        except Exception:
            pass
        for style, text in p.blocks:
            par = d.add_paragraph()
            try:
                par.style = d.styles[style]
            except Exception:
                pass
            par.add_run(text)
        if not p.blocks:
            return FileResult(path, 'skipped', 'HTML 无正文文本')
        atomic_save(d.save, out)
    except Exception as e:
        return FileResult(path, 'fail', f'HTML→Word 失败: {e}')
    return FileResult(path, 'ok',
                      f'→ Word({len(p.blocks)} 段, 可编辑)', out)


# ==================================================== 图片 → Word
def convert_image_to_word(path: str, options: dict) -> FileResult:
    ext = os.path.splitext(path)[1].lower()
    if ext not in IMG_EXTS:
        return FileResult(path, 'skipped', f'非图片文件 {ext}')
    out_dir = _conv_out_dir(path, options)
    out = _out_path(path, out_dir, '.docx')
    try:
        import docx
        from docx.shared import Cm
        from PIL import Image

        d = docx.Document()
        try:
            d.styles['Normal'].font.name = '微软雅黑'
        except Exception:
            pass
        # A4 可用宽度约 16cm, 超过则等比缩到页宽内(不放大不满页的小图)
        max_w_cm = 16.0
        n = 0
        with Image.open(path) as im:
            try:
                nf = getattr(im, 'n_frames', 1)
            except Exception:
                nf = 1
            frames = []
            for i in range(max(1, nf)):
                im.seek(i)
                frames.append(im.convert('RGB'))
        for fr in frames:
            w_px, h_px = fr.size
            if w_px <= 0:
                continue
            w_cm = min(max_w_cm, w_px / 96 * 2.54)
            h_cm = h_px / w_px * w_cm
            # add_picture 直接吃文件路径; 帧对象也支持(内存图)
            if nf == 1:
                d.add_picture(path, width=Cm(w_cm))
            else:
                import io
                buf = io.BytesIO()
                fr.save(buf, format='PNG')
                buf.seek(0)
                d.add_picture(buf, width=Cm(w_cm))
            if h_cm > 22:
                d.add_page_break()
            n += 1
        if n == 0:
            return FileResult(path, 'skipped', '图片尺寸异常')
        atomic_save(d.save, out)
    except Exception as e:
        return FileResult(path, 'fail', f'图片→Word 失败: {e}')
    return FileResult(path, 'ok', f'→ Word(插入 {n} 张图片, 可编辑)', out)


# ==================================================== PDF → PPT
def convert_pdf_to_ppt(path: str, options: dict) -> FileResult:
    """PDF 每页 → PPT 一页(整页贴图, 16:9 居中)。"""
    if os.path.splitext(path)[1].lower() != '.pdf':
        return FileResult(path, 'skipped', '非 PDF 文件')
    out_dir = _conv_out_dir(path, options)
    out = _out_path(path, out_dir, '.pptx')
    tmp = tempfile.mkdtemp(prefix='qyppt_')
    try:
        import pymupdf
        from pptx import Presentation
        from pptx.util import Emu
        doc = pymupdf.open(path)
        n_pages = len(doc)
        if n_pages == 0:
            return FileResult(path, 'skipped', 'PDF 无页面')
        prs = Presentation()
        prs.slide_width = Emu(12192000)     # 16:9
        prs.slide_height = Emu(6858000)
        blank = prs.slide_layouts[6]
        sw, sh = prs.slide_width, prs.slide_height
        n = 0
        for page in doc:
            pix = page.get_pixmap(matrix=pymupdf.Matrix(2, 2), alpha=False)
            png = os.path.join(tmp, f'p{n + 1}.png')
            pix.save(png)
            iw, ih = pix.width, pix.height
            scale = min(sw / iw, sh / ih)
            w, h = int(iw * scale), int(ih * scale)
            slide = prs.slides.add_slide(blank)
            slide.shapes.add_picture(png, (sw - w) // 2, (sh - h) // 2,
                                     width=w, height=h)
            n += 1
        doc.close()
        atomic_save(prs.save, out)
    except Exception as e:
        return FileResult(path, 'fail', f'PDF→PPT 失败: {e}')
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return FileResult(path, 'ok', f'→ PPT({n} 页, 16:9 居中)', out)


CONVERTERS = {
    'office2pdf': convert_office_to_pdf,
    'pdf2word': convert_pdf_to_word,
    'pdf2img': convert_pdf_to_images,
    'pdf2txt': convert_pdf_to_text,
    'pdf2html': convert_pdf_to_html,
    'word2txt': convert_word_to_text,
    'word2html': convert_word_to_html,
    'excel2csv': convert_excel_to_csv,
    'excel2html': convert_excel_to_html,
    'ppt2txt': convert_ppt_to_text,
    'ppt2img': convert_ppt_to_images,
    'img2pdf': convert_image_to_pdf,
    'txt2word': convert_text_to_word,
    'pdf2excel': convert_pdf_to_excel,
    'csv2excel': convert_csv_to_excel,
    'excel2word': convert_excel_to_word,
    'excel2json': convert_excel_to_json,
    'word2md': convert_word_to_md,
    'html2word': convert_html_to_word,
    'img2word': convert_image_to_word,
    'pdf2ppt': convert_pdf_to_ppt,
}


def convert_file(path: str, options: dict) -> FileResult:
    """按 options['convert'] 路由到具体转换器。"""
    key = options.get('convert') or options.get('op')
    fn = CONVERTERS.get(key)
    if fn is None:
        return FileResult(path, 'fail', f'未知转换类型 {key!r}')
    return fn(path, options)


# 供 UI 显示: 每个转换方向适用的源扩展名(用于页面提示与守卫)
def supported_exts(convert_key: str) -> tuple:
    if convert_key in ('office2pdf',):
        return tuple(sorted(OFFICE_EXTS))
    if convert_key in ('pdf2word', 'pdf2img', 'pdf2txt', 'pdf2html'):
        return ('.pdf',)
    if convert_key in ('word2txt',):
        return ('.docx', '.docm', '.doc', '.wps')
    if convert_key == 'word2html':
        return ('.docx', '.docm')
    if convert_key in ('excel2csv',):
        return ('.xlsx', '.xlsm', '.xls', '.et')
    if convert_key == 'excel2html':
        return ('.xlsx', '.xlsm', '.xls', '.et')
    if convert_key in ('ppt2txt', 'ppt2img'):
        return ('.pptx', '.pptm', '.ppt', '.dps')
    if convert_key == 'img2pdf':
        return tuple(sorted(IMG_EXTS))
    if convert_key == 'txt2word':
        return tuple(sorted(TXT_EXTS)) + ('.md', '.ini', '.yaml', '.yml')
    if convert_key == 'pdf2excel':
        return ('.pdf',)
    if convert_key == 'csv2excel':
        return ('.csv', '.tsv', '.txt')
    if convert_key == 'excel2word':
        return ('.xlsx', '.xlsm', '.xls', '.et')
    if convert_key == 'excel2json':
        return ('.xlsx', '.xlsm', '.xls', '.et')
    if convert_key == 'word2md':
        return ('.docx', '.docm')
    if convert_key == 'html2word':
        return ('.html', '.htm')
    if convert_key == 'img2word':
        return tuple(sorted(IMG_EXTS))
    if convert_key == 'pdf2ppt':
        return ('.pdf',)
    return ()
