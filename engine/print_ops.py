# -*- coding: utf-8 -*-
"""青雁批量工具箱 v3 — 批量打印引擎 (Windows GDI, 全离线)。

路由:
- 图片 (jpg/png/bmp/tif/webp/gif) → GDI 直印 (PIL + win32ui)
- 文本 (txt/md/csv/json...)       → GDI 文本绘制 (自动换行/分页, 中文字体)
- PDF                             → PyMuPDF 渲染为位图 → GDI 直印
                                    (未安装 PyMuPDF 时回退系统关联程序)
- Office (doc/xls/ppt/wps/et/dps) → 本机 COM 转 PDF (Word/Excel/PowerPoint
                                    或 WPS) → 按 PDF 打印; 无 Office/WPS 时
                                    回退系统关联打印(用系统默认打印机)

虚拟测试: 选择「Microsoft Print to PDF」等虚拟打印机并在选项里给出输出
目录时, 打印任务静默产出 PDF 文件(不弹另存为、不耗纸), 供自动化验证。
"""
from __future__ import annotations

import os
import shutil
import tempfile

from engine.core import FileResult
from engine.ops import TXT_EXTS, read_text_auto

IMG_PRINT_EXTS = frozenset({'.jpg', '.jpeg', '.png', '.bmp', '.gif',
                            '.webp', '.tif', '.tiff'})
OFFICE_EXTS = frozenset({'.doc', '.docx', '.wps',
                         '.xls', '.xlsx', '.et',
                         '.ppt', '.pptx', '.dps'})


class _NoOfficeError(RuntimeError):
    pass


# ---------------------------------------------------------------- 打印机
def list_printers() -> list:
    try:
        import win32print
        flags = (win32print.PRINTER_ENUM_LOCAL |
                 win32print.PRINTER_ENUM_CONNECTIONS)
        return [p[2] for p in win32print.EnumPrinters(flags)]
    except Exception:
        return []


def default_printer() -> str:
    try:
        import win32print
        return win32print.GetDefaultPrinter()
    except Exception:
        return ''


def is_virtual_pdf_printer(name: str) -> bool:
    low = (name or '').lower()
    return 'print to pdf' in low or 'xps' in low or 'pdf editor' in low


# ------------------------------------------------------- 输出路径(虚拟)
def _pdf_output_path(options: dict, src: str):
    """虚拟打印机(Print to PDF 等)时的静默输出文件; 未填输出目录返回 None。"""
    out_dir = (options.get('pdf_out_dir') or '').strip()
    if not out_dir:
        return None
    os.makedirs(out_dir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(src))[0]
    cand = os.path.join(out_dir, stem + '.pdf')
    i = 1
    while os.path.exists(cand):
        cand = os.path.join(out_dir, f'{stem}_{i}.pdf')
        i += 1
    return cand


# ------------------------------------------------------------- GDI 底层
def _open_printer_dc(printer: str):
    import win32ui
    dc = win32ui.CreateDC()
    dc.CreatePrinterDC(printer)
    return dc


def _page_metrics(dc):
    HORZRES, VERTRES, LOGPIXELSX = 8, 10, 88   # GetDeviceCaps 索引
    page_w = dc.GetDeviceCaps(HORZRES)
    page_h = dc.GetDeviceCaps(VERTRES)
    dpi = max(72, dc.GetDeviceCaps(LOGPIXELSX))
    return page_w, page_h, dpi


def _select_cjk_font(dc, dpi: int, pt: int = 11):
    import win32ui
    font = win32ui.CreateFont({
        'name': 'Microsoft YaHei UI',
        'height': -int(pt * dpi / 72),
        'weight': 400,
        'charset': 134,           # GB2312 确保中文字形
    })
    dc.SelectObject(font)
    return font


def _wrap_lines(dc, text: str, max_w: int) -> list:
    """按像素宽度把逻辑行折成物理行(制表符展开为 4 空格)。"""
    phys: list = []
    for raw in text.replace('\t', '    ').splitlines():
        line = raw.rstrip('\r')
        if not line:
            phys.append('')
            continue
        while True:
            if dc.GetTextExtent(line)[0] <= max_w or not line:
                phys.append(line)
                break
            lo, hi, fit = 1, len(line), 1
            while lo <= hi:
                mid = (lo + hi) // 2
                if dc.GetTextExtent(line[:mid])[0] <= max_w:
                    fit = mid
                    lo = mid + 1
                else:
                    hi = mid - 1
            phys.append(line[:fit])
            line = line[fit:]
    return phys


def gdi_print_text(printer: str, text: str, copies: int, doc_name: str,
                   output: str = None) -> int:
    """GDI 打印文本, 返回总页数(含份数)。"""
    dc = _open_printer_dc(printer)
    page_w, page_h, dpi = _page_metrics(dc)
    margin = dpi // 5
    line_h = max(1, int(dpi * 11 / 72 * 1.4))
    max_w = page_w - 2 * margin

    dc.StartDoc(doc_name, output) if output else dc.StartDoc(doc_name)
    _select_cjk_font(dc, dpi)
    phys = _wrap_lines(dc, text, max_w)
    lpp = max(1, (page_h - 2 * margin) // line_h)
    pages = [phys[i:i + lpp] for i in range(0, len(phys), lpp)] or [[]]
    printed = 0
    for _copy in range(max(1, copies)):
        for pg in pages:
            dc.StartPage()
            _select_cjk_font(dc, dpi)
            y = margin
            for ln in pg:
                if ln:
                    dc.TextOut(margin, y, ln)
                y += line_h
            dc.EndPage()
            printed += 1
    dc.EndDoc()
    del dc
    return printed


def _draw_image(dc, im, page_w: int, page_h: int, dpi: int, fit: bool):
    from PIL import ImageWin
    margin = dpi // 5
    max_w = page_w - 2 * margin
    max_h = page_h - 2 * margin
    iw, ih = im.size
    if fit and iw > 0 and ih > 0:
        scale = min(max_w / iw, max_h / ih)
        dw, dh = max(1, int(iw * scale)), max(1, int(ih * scale))
    else:
        dw, dh = min(iw, max_w), min(ih, max_h)
    x = (page_w - dw) // 2
    y = (page_h - dh) // 2
    dib = ImageWin.Dib(im)
    dib.draw(dc.GetHandleOutput(), (x, y, x + dw, y + dh))


def gdi_print_images(printer: str, images: list, copies: int, doc_name: str,
                     output: str = None, fit: bool = True) -> int:
    """GDI 打印一组位图(每张一页), 返回总页数(含份数)。"""
    dc = _open_printer_dc(printer)
    page_w, page_h, dpi = _page_metrics(dc)
    dc.StartDoc(doc_name, output) if output else dc.StartDoc(doc_name)
    printed = 0
    for _copy in range(max(1, copies)):
        for im in images:
            if im.mode not in ('RGB', 'L'):
                im = im.convert('RGB')
            dc.StartPage()
            _draw_image(dc, im, page_w, page_h, dpi, fit)
            dc.EndPage()
            printed += 1
    dc.EndDoc()
    del dc
    return printed


def gdi_print_image_file(printer: str, img_path: str, copies: int,
                         doc_name: str, output: str = None,
                         fit: bool = True) -> int:
    from PIL import Image, ImageOps
    im = Image.open(img_path)
    try:
        im = ImageOps.exif_transpose(im)
    except Exception:
        pass
    return gdi_print_images(printer, [im], copies, doc_name, output, fit)


# ----------------------------------------------------------------- PDF
def _pdf_to_images(pdf_path: str, dpi: int = 150) -> list:
    """PDF → 位图列表 (PyMuPDF; 未安装时抛 ImportError)。"""
    try:
        import pymupdf
    except ImportError:
        raise ImportError('PyMuPDF 未安装')
    from PIL import Image
    imgs = []
    doc = pymupdf.open(pdf_path)
    try:
        for page in doc:
            pix = page.get_pixmap(dpi=dpi)
            img = Image.frombytes('RGB', (pix.width, pix.height),
                                  pix.samples)
            imgs.append(img)
    finally:
        doc.close()
    return imgs


# -------------------------------------------------------------- Office
def _office_to_pdf(path: str, workdir: str) -> str:
    """用本机 Office(或 WPS) COM 把文档静默转成 PDF, 返回 PDF 路径。

    只做格式转换(ExportAsFixedFormat), 不调用 PrintOut — 打印由统一的
    GDI 流程送到用户选择的打印机, 全程不弹窗、不出纸。
    """
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
    pdf_path = os.path.join(workdir, 'conv.pdf')
    import time
    pythoncom.CoInitialize()
    last_err = None
    for prog, kind in progs:
        app = None
        # PowerPoint 单实例: 上一实例 Quit 后服务器可能仍在退出,
        # 此时 DispatchEx 会误报「无效的类字符串」— 短重试
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
                # PowerPoint 无 Visible 属性(设置会抛 com_error), 跳过
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
                doc.ExportAsFixedFormat(pdf_path, 17)   # wdExportFormatPDF
                doc.Close(False)
            elif kind == 'excel':
                wb = app.Workbooks.Open(abspath, ReadOnly=True)
                wb.ExportAsFixedFormat(0, pdf_path)     # xlTypePDF
                wb.Close(False)
            else:
                pres = app.Presentations.Open(abspath, True, False, False)
                # PowerPoint 的 ExportAsFixedFormat 在 pywin32 早期绑定下
                # 参数转换失败, 改用等价的 SaveAs(Path, ppSaveAsPDF=32)
                pres.SaveAs(pdf_path, 32)
                pres.Close()
            if os.path.exists(pdf_path):
                return pdf_path
            last_err = RuntimeError('转换未产出文件')
        except SystemError as e:
            last_err = e
        except Exception as e:
            last_err = e
        finally:
            try:
                app.Quit()
            except Exception:
                pass
            app = None
    if last_err is not None and '0x80040154' in str(last_err):
        raise _NoOfficeError('本机未安装 Office/WPS, 无法直接打印 ' +
                             ext + ' 文档')
    raise _NoOfficeError('Office/WPS 转换失败: ' + str(last_err))


# ---------------------------------------------------------------- 路由
def print_file(path: str, options: dict) -> FileResult:
    """批量打印路由: (path, options) -> FileResult。"""
    printer = (options.get('printer') or '').strip()
    if not printer or printer.startswith('未检测到'):
        return FileResult(path, 'fail', '未选择可用打印机')
    if printer not in list_printers():
        return FileResult(path, 'fail', '打印机不可用或已断开: ' + printer)
    try:
        copies = max(1, int(options.get('copies', 1) or 1))
    except Exception:
        copies = 1
    fit = bool(options.get('fit_page', True))
    ext = os.path.splitext(path)[1].lower()
    doc_name = os.path.basename(path)
    virtual = is_virtual_pdf_printer(printer)
    output = _pdf_output_path(options, path) if virtual else None

    import win32ui
    _printer_dc_ready = win32ui.CreateDC   # 可用性探测(pywin32)
    try:
        _printer_dc_ready
    except AttributeError as e:
        return FileResult(path, 'fail', '打印组件不可用(pywin32): ' + str(e))

    try:
        if ext in IMG_PRINT_EXTS:
            pages = gdi_print_image_file(printer, path, copies, doc_name,
                                         output, fit)
            detail = f'已打印 {pages} 页 × {copies} 份'
        elif ext == '.pdf':
            try:
                imgs = _pdf_to_images(path)
            except ImportError:
                try:
                    os.startfile(path, 'print')
                    return FileResult(
                        path, 'ok',
                        '已发送到系统关联程序打印(未内置 PDF 渲染组件)')
                except Exception as e:
                    return FileResult(
                        path, 'fail',
                        'PDF 打印失败且无系统关联: ' + str(e))
            pages = gdi_print_images(printer, imgs, copies, doc_name,
                                     output, fit)
            detail = f'已打印 {pages} 页 × {copies} 份'
        elif ext in OFFICE_EXTS:
            workdir = tempfile.mkdtemp(prefix='qyprint_')
            try:
                pdf_path = _office_to_pdf(path, workdir)
                imgs = _pdf_to_images(pdf_path)
                pages = gdi_print_images(printer, imgs, copies, doc_name,
                                         output, fit)
                detail = (f'Office→PDF 转换成功, 已打印 '
                          f'{pages} 页 × {copies} 份')
            except _NoOfficeError as e:
                if virtual:
                    return FileResult(
                        path, 'fail', str(e) +
                        '; 虚拟打印机需要内置转换, 无法回退系统关联')
                try:
                    os.startfile(path, 'print')
                    return FileResult(
                        path, 'ok',
                        '未安装 Office/WPS, 已改用系统关联程序打印到'
                        '系统默认打印机')
                except Exception:
                    return FileResult(path, 'fail', str(e))
            finally:
                shutil.rmtree(workdir, ignore_errors=True)
        elif ext in TXT_EXTS:
            text = read_text_auto(path)
            if not text.strip():
                return FileResult(path, 'skipped', '空文件不打印')
            pages = gdi_print_text(printer, text, copies, doc_name, output)
            detail = f'已打印 {pages} 页 × {copies} 份'
        else:
            return FileResult(path, 'skipped',
                              f'不支持的打印类型 {ext} (支持图片/文本/PDF/Office)')
        return FileResult(path, 'ok', detail + f' → {printer}',
                          output or '')
    except Exception as e:
        return FileResult(path, 'fail',
                          type(e).__name__ + ': ' + str(e))
