# -*- coding: utf-8 -*-
"""青雁批量工具箱 v3 — 新增高级处理器 (Word/PPT/Excel/PDF 深度处理)。

依赖: python-docx / openpyxl / python-pptx / pypdf / Pillow (全离线)
每个函数: (path, options) -> FileResult

v3 增强:
- docx XML 级操作(批注/超链接/域/修订/水印)与 v2 行为一致。
- 所有就地修改经 _out() 统一走备份 + 撤销。
"""
from __future__ import annotations

from engine.safety import backup_file, atomic_save

import io
import os
import re
import shutil
import zipfile
from copy import deepcopy
from datetime import datetime

from engine.core import FileResult
from engine.ops import default_trash_dir, detect_encoding


# ------------------------------------------------------------------ 工具
def _replace(tmp: str, dst: str, attempts: int = 6) -> None:
    """带重试的原子替换。

    Windows 上目标文件刚被本进程/杀软/索引器碰过时, os.replace 会瞬时抛
    [WinError 5] 拒绝访问, 隔几十毫秒再试就过了。不重试的话用户会看到
    "处理失败", 而临时文件已经写好, 只是没换上去。
    """
    import time
    last = None
    for i in range(attempts):
        try:
            os.replace(tmp, dst)
            return
        except PermissionError as e:
            last = e
            time.sleep(0.05 * (i + 1))
    if last is not None:
        raise last


def _uniq_out(p: str) -> str:
    """输出文件名冲突时追加 (1)(2), 避免静默覆盖已有文件。"""
    if not os.path.exists(p):
        return p
    b, e = os.path.splitext(p)
    k = 1
    while os.path.exists(f'{b}({k}){e}'):
        k += 1
    return f'{b}({k}){e}'


def _bak(path: str) -> str:
    """就地修改前备份, 返回本次使用的备份路径。

    旧实现是"备份已存在就复用", 导致对同一文件连续处理两次时, 第二次的撤销会
    退回到第一次操作之前(而非本次操作之前), 第一次的撤销也同时失效。
    这里每次都重新备份, 保证撤销精确回退一步。
    """
    bak = backup_file(path)
    return bak


def _out(path: str, options: dict, tag: str):
    """输出路径: out_dir 指定则输出到输出目录; 否则就地(先备份)。"""
    out_dir = options.get('out_dir')
    base, ext = os.path.splitext(os.path.basename(path))
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
        out = os.path.abspath(os.path.join(out_dir, base + tag + ext))
        return _uniq_out(out), []
    bak = _bak(path)
    return os.path.abspath(path), [{'src': bak, 'dst': path, 'kind': 'restore'}]


def _word_xml_ops(path: str, transform, tag: str, options: dict,
                  settings_hook=None) -> FileResult:
    """解包 docx → 对 document.xml 应用 transform(字符串级) → 重打包。

    就地修改时先写临时文件再覆盖, 避免读写同一 zip 冲突。
    settings_hook: 可选, 对 word/settings.xml 做同样处理(背景可见性等
    需要动 settings 的操作用)。
    """
    with zipfile.ZipFile(path) as zf:
        src = zf.read('word/document.xml').decode('utf-8')
        settings = zf.read('word/settings.xml').decode('utf-8') \
            if 'word/settings.xml' in zf.namelist() else None
    doc, n = transform(src)
    if n == 0:
        return FileResult(path, 'skipped', '无可处理内容')
    if settings_hook is not None and settings is not None:
        settings = settings_hook(settings)
    out, undo = _out(path, options, tag)
    tmp_out = out + '.tmpzip'
    with zipfile.ZipFile(path) as zin, \
            zipfile.ZipFile(tmp_out, 'w', zipfile.ZIP_DEFLATED,
                            compresslevel=6) as zout:
        for item in zin.namelist():
            if item == 'word/document.xml':
                zout.writestr(item, doc.encode('utf-8'))
            elif item == 'word/settings.xml' and settings is not None:
                zout.writestr(item, settings.encode('utf-8'))
            else:
                zout.writestr(item, zin.read(item))
    if os.path.abspath(tmp_out) != os.path.abspath(out):
        shutil.move(tmp_out, out)
    return FileResult(path, 'ok', f'处理 {n} 处', out, undo=undo)


# ==================================================================== Word
def docx_remove_comments(path: str, options: dict) -> FileResult:
    from engine.comment_cleanup import remove_comments
    return remove_comments(path, options)


def docx_remove_links(path: str, options: dict) -> FileResult:
    def t(doc):
        n = len(re.findall(r'<w:hyperlink\b', doc))

        def unwrap(m):
            inner = m.group(1)
            runs = re.findall(r'<w:r\b.*?</w:r>', inner, flags=re.S)
            return ''.join(runs)

        doc2 = re.sub(r'<w:hyperlink\b[^>]*>(.*?)</w:hyperlink>', unwrap, doc,
                      flags=re.S)
        return doc2, n
    return _word_xml_ops(path, t, '_nolink', options)


def docx_remove_fields(path: str, options: dict) -> FileResult:
    mode = options.get('field_mode', '删除')

    def t(doc):
        n = len(re.findall(r'<w:fldChar\b', doc)) + \
            len(re.findall(r'<w:fldSimple\b', doc))
        if mode == '删除':
            doc2 = re.sub(r'<w:fldSimple\b[^>]*>(.*?)</w:fldSimple>', r'\1',
                          doc, flags=re.S)
            doc2 = re.sub(r'<w:r[^>]*>(?:(?!</w:r>).)*?'
                          r'<w:fldChar w:fldCharType="begin"[^>]*/>'
                          r'(?:(?!</w:r>).)*?</w:r>', '', doc2, flags=re.S)
            doc2 = re.sub(r'<w:r[^>]*>(?:(?!</w:r>).)*?<w:instrText[^>]*>'
                          r'(?:(?!</w:r>).)*?</w:r>', '', doc2, flags=re.S)
            doc2 = re.sub(r'<w:r[^>]*>(?:(?!</w:r>).)*?'
                          r'<w:fldChar w:fldCharType="(?:separate|end)"[^>]*/>'
                          r'(?:(?!</w:r>).)*?</w:r>', '', doc2, flags=re.S)
            return doc2, n
        # 转为纯文本: fldSimple 保留内部文本
        doc2 = re.sub(r'<w:fldSimple\b[^>]*>(.*?)</w:fldSimple>', r'\1',
                      doc, flags=re.S)
        return doc2, n
    return _word_xml_ops(path, t, '_nofld', options)


def docx_accept_revisions(path: str, options: dict) -> FileResult:
    """接受/拒绝修订 — 用 lxml 按元素处理。

    早期用正则改 XML: `[^>]*>` 会把自闭合标签的 `/` 吃进字符组, 从
    自闭合 <w:del/> 一直匹配到下一个成对 </w:del>, 把中间的正常段落
    整段删掉, 还会产生失配标签(Word 要求修复)。"删除回车"生成的
    段落属性修订正是自闭合形态, 真实修订文档极常见。
    """
    mode = options.get('rev_mode', '接受')

    def t(doc):
        import lxml.etree as ET
        W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        root = ET.fromstring(doc.encode('utf-8'))
        ins_list = list(root.iter(f'{W}ins', f'{W}moveTo'))
        del_list = list(root.iter(f'{W}del', f'{W}moveFrom'))
        n = len(ins_list) + len(del_list)

        def _unwrap(el):
            parent = el.getparent()
            if parent is None:
                return
            idx = list(parent).index(el)
            for child in list(el):
                parent.insert(idx, child)
                idx += 1
            parent.remove(el)

        def _drop(el):
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)

        if mode == '接受':
            # 删除内容随 w:del 子树移除; 插入内容解包保留
            for el in del_list:
                _drop(el)
            for el in ins_list:
                _unwrap(el)
        else:
            # 拒绝: 插入内容移除; 删除内容保留且 delText 必须转回 w:t
            for el in ins_list:
                _drop(el)
            for el in del_list:
                _unwrap(el)
            for dt in list(root.iter(f'{W}delText')):
                dt.tag = f'{W}t'
        # 修订前快照(pPrChange 等)无论接受/拒绝都应清除
        for tag in ('pPrChange', 'rPrChange', 'sectPrChange', 'tblPrChange',
                    'trPrChange', 'tcPrChange'):
            for el in list(root.iter(f'{W}{tag}')):
                _drop(el)
        out_xml = ET.tostring(root, encoding='unicode')
        return out_xml, n
    return _word_xml_ops(path, t, '_rev', options)


def docx_remove_macros(path: str, options: dict) -> FileResult:
    from engine.macro_cleanup import remove_macros
    return remove_macros(path, options, 'word')


def docx_remove_watermark(path: str, options: dict) -> FileResult:
    tmp_out = path + '.nowm.tmpzip'
    removed = 0
    try:
        with zipfile.ZipFile(path) as zin, \
                zipfile.ZipFile(tmp_out, 'w', zipfile.ZIP_DEFLATED,
                                compresslevel=6) as zout:
            for item in zin.namelist():
                data = zin.read(item)
                if re.match(r'word/header\d*\.xml', item):
                    xml = data.decode('utf-8', errors='replace')
                    before = xml
                    xml = re.sub(r'<v:shape\b[^>]*type="#_x0000_t136"[^>]*>.*?</v:shape>',
                                 '', xml, flags=re.S)
                    xml = re.sub(r'<v:shape\b[^>]*>[^<]*<v:textbox[^>]*>.*?'
                                 r'PowerPlusWaterMark.*?</v:shape>', '', xml, flags=re.S)
                    xml = re.sub(r'<w:pict\b[^>]*>(?:(?!</w:pict>).)*?'
                                 r'(?:PowerPlusWaterMark|_x0000_t136)'
                                 r'(?:(?!</w:pict>).)*?</w:pict>', '', xml, flags=re.S)
                    if xml != before:
                        removed += 1
                    data = xml.encode('utf-8')
                zout.writestr(item, data)
        if removed == 0:
            return FileResult(path, 'skipped', '未在页眉发现水印形状')
        out, undo = _out(path, options, '_nowm')
        shutil.move(tmp_out, out)
    finally:
        if os.path.exists(tmp_out):
            try:
                os.remove(tmp_out)
            except OSError:
                pass
    return FileResult(path, 'ok', f'清理 {removed} 个页眉水印', out, undo=undo)


def docx_remove_header_footer(path: str, options: dict) -> FileResult:
    # 先判后改: 没有页眉页脚内容的文档不该产出副本, 也不该报"已清空"
    has_content = False
    with zipfile.ZipFile(path) as zin:
        for item in zin.namelist():
            if not re.match(r'word/(header|footer)\d*\.xml', item):
                continue
            data = zin.read(item)
            if b'<w:t' in data or b'<w:drawing' in data \
                    or b'<w:pict' in data or b'<v:' in data:
                has_content = True
                break
    if not has_content:
        return FileResult(path, 'skipped', '未发现页眉页脚内容')
    out, undo = _out(path, options, '_hf')
    tmp_out = out + '.tmpzip'
    HDR = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           b'<w:hdr xmlns:w="http://schemas.openxmlformats.org/'
           b'wordprocessingml/2006/main"/>')
    FTR = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           b'<w:ftr xmlns:w="http://schemas.openxmlformats.org/'
           b'wordprocessingml/2006/main"/>')
    with zipfile.ZipFile(path) as zin, \
            zipfile.ZipFile(tmp_out, 'w', zipfile.ZIP_DEFLATED,
                            compresslevel=6) as zout:
        for item in zin.namelist():
            data = zin.read(item)
            if re.match(r'word/(header|footer)\d*\.xml', item):
                data = HDR if 'header' in item else FTR
            zout.writestr(item, data)
    if os.path.abspath(tmp_out) != os.path.abspath(out):
        shutil.move(tmp_out, out)
    return FileResult(path, 'ok', '页眉页脚已清空', out, undo=undo)


def docx_set_background(path: str, options: dict) -> FileResult:
    color = str(options.get('bg_color', '') or '').strip()
    remove = bool(options.get('bg_remove', False))

    def t(doc):
        doc2 = re.sub(r'<w:background\b[^>]*/>', '', doc, flags=re.S)
        doc2 = re.sub(r'<w:background\b[^>]*>.*?</w:background>', '', doc2,
                      flags=re.S)
        if remove:
            return doc2, 1
        m = re.match(r'^#?([0-9A-Fa-f]{6})$', color)
        if not m:
            return doc, 0
        hexv = m.group(1)
        bg = f'<w:background w:color="{hexv}"/>'
        doc2 = doc2.replace('<w:body>', bg + '<w:body>', 1)
        # OOXML 要求 settings.xml 里声明 displayBackgroundShape,
        # 否则 Word 不渲染背景(任务报 ok 用户却什么都看不到)
        return doc2, 1

    def settings_hook(settings_xml: str) -> str:
        if remove:
            return re.sub(r'<w:displayBackgroundShape\b[^>]*/>', '',
                          settings_xml)
        if '<w:displayBackgroundShape' not in settings_xml:
            m = re.search(r'(<w:settings\b[^>]*>)', settings_xml)
            if m:
                settings_xml = settings_xml.replace(
                    m.group(1), m.group(1) + '<w:displayBackgroundShape/>', 1)
        return settings_xml
    return _word_xml_ops(path, t, '_bg', options,
                         settings_hook=settings_hook)


def docx_extract_images(path: str, options: dict) -> FileResult:
    out_dir = options.get('out_dir') or (
        os.path.join(os.path.dirname(path),
                     os.path.splitext(os.path.basename(path))[0] + '_图片'))
    base = os.path.splitext(os.path.basename(path))[0]
    os.makedirs(out_dir, exist_ok=True)
    n = 0
    last = ''
    with zipfile.ZipFile(path) as zf:
        for item in zf.namelist():
            if re.match(r'word/media/', item):
                fn = base + '_' + os.path.basename(item)
                dst = _uniq_out(os.path.abspath(os.path.join(out_dir, fn)))
                with open(dst, 'wb') as fh:
                    fh.write(zf.read(item))
                n += 1
                last = dst
    if n == 0:
        return FileResult(path, 'skipped', '无图片')
    return FileResult(path, 'ok', f'提取 {n} 张图片', last)


def docx_set_content_style(path: str, options: dict) -> FileResult:
    import docx
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt, RGBColor

    d = docx.Document(path)
    font = options.get('style_font', '')
    size = options.get('style_size')
    color = str(options.get('style_color', '') or '').strip()
    align = options.get('style_align', '')
    line_spacing = options.get('style_line_spacing')
    n = 0
    amap = {'左': 'LEFT', '居中': 'CENTER', '右': 'RIGHT', '两端': 'JUSTIFY'}

    def _apply(p) -> None:
        nonlocal n
        if not p.text.strip():
            return
        for r in p.runs:
            if font:
                r.font.name = font
                # python-docx 的 font.name 只写 w:ascii/w:hAnsi, 汉字走
                # w:eastAsia 字体槽 —— 不设置它, 中文在 Word 里仍是原字体
                from docx.oxml.ns import qn
                rPr = r._element.get_or_add_rPr()
                rFonts = rPr.find(qn('w:rFonts'))
                if rFonts is None:
                    rFonts = rPr.makeelement(qn('w:rFonts'), {})
                    rPr.insert(0, rFonts)
                rFonts.set(qn('w:eastAsia'), font)
            if size:
                r.font.size = Pt(int(size))
            if color and re.match(r'^#?[0-9A-Fa-f]{6}$', color):
                r.font.color.rgb = RGBColor.from_string(
                    color.lstrip('#').upper())
        if align:
            attr = amap.get(align)
            if attr:
                p.alignment = getattr(WD_ALIGN_PARAGRAPH, attr)
        if line_spacing:
            p.paragraph_format.line_spacing = float(line_spacing)
        n += 1

    for p in d.paragraphs:
        _apply(p)
    for t in d.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _apply(p)
    if n == 0:
        return FileResult(path, 'skipped', '无可样式化段落')
    out, undo = _out(path, options, '_style')
    atomic_save(d.save, out)
    return FileResult(path, 'ok', f'样式化 {n} 段', out, undo=undo)


def docx_set_page_style(path: str, options: dict) -> FileResult:
    import docx
    from docx.enum.section import WD_ORIENT

    d = docx.Document(path)
    margin = options.get('page_margin_cm')
    orientation = options.get('page_orientation', '')
    n = 0
    for sec in d.sections:
        if margin:
            from docx.shared import Cm
            for attr in ('left_margin', 'right_margin', 'top_margin',
                         'bottom_margin'):
                setattr(sec, attr, Cm(float(margin)))
        if orientation in ('横向', '纵向'):
            w, h = sec.page_width, sec.page_height
            if orientation == '横向':
                sec.orientation = WD_ORIENT.LANDSCAPE
                sec.page_width, sec.page_height = max(w, h), min(w, h)
            else:
                sec.orientation = WD_ORIENT.PORTRAIT
                sec.page_width, sec.page_height = min(w, h), max(w, h)
        n += 1
    out, undo = _out(path, options, '_page')
    atomic_save(d.save, out)
    return FileResult(path, 'ok', f'页面设置 {n} 节', out, undo=undo)


def docx_meta(path: str, options: dict) -> FileResult:
    import docx

    d = docx.Document(path)
    mode = options.get('meta_mode', '清空')
    cp = d.core_properties
    if mode == '清空':
        cp.author = cp.last_modified_by = ''
        cp.title = cp.subject = cp.keywords = cp.comments = cp.category = ''
        detail = '元数据已清空'
    else:
        if options.get('meta_author'):
            cp.author = options['meta_author']
            cp.last_modified_by = options['meta_author']
        if options.get('meta_title'):
            cp.title = options['meta_title']
        if options.get('meta_subject'):
            cp.subject = options['meta_subject']
        if options.get('meta_keywords'):
            cp.keywords = options['meta_keywords']
        detail = '元数据已修改(作者=' + str(cp.author) + ')'
    out, undo = _out(path, options, '_meta')
    atomic_save(d.save, out)
    return FileResult(path, 'ok', detail, out, undo=undo)


def docx_split_by_paragraphs(path: str, options: dict) -> FileResult:
    import docx
    from docx.oxml.ns import qn

    d = docx.Document(path)
    try:
        chunk = max(1, int(options.get('docx_chunk', 50) or 50))
    except (TypeError, ValueError):
        return FileResult(path, 'fail', '每份段数必须是整数')
    out_dir = options.get('out_dir') or os.path.dirname(path)
    os.makedirs(out_dir, exist_ok=True)
    body = d.element.body
    # sectPr 必须是 w:body 的最后一个子元素。body.append() 会追加到它之后,
    # 违反 ECMA-376, Word 打开会报"需要修复"。每个分份保存前都把 sectPr
    # 摘出来再 append 回去(见下面循环内)。
    blocks = [el for el in body
              if el.tag in (qn('w:p'), qn('w:tbl'))]
    total = len(blocks)
    if total == 0:
        return FileResult(path, 'skipped', '空文档')
    base = os.path.splitext(os.path.basename(path))[0]
    n_parts = 0
    last = ''
    outputs = []
    src_abs = os.path.abspath(path)
    for i in range(0, total, chunk):
        n_parts += 1
        part = os.path.abspath(
            _uniq_out(os.path.join(out_dir, f'{base}_part{n_parts:03d}.docx')))
        # 整份复制原文件再删掉本份之外的段落: deepcopy 元素到新文档不带
        # 关系/媒体部件, 含图片的文档分份后会引用不存在的 rId(红叉丢图)。
        # 复制法保证每个分份都是完整的独立包。
        shutil.copy2(src_abs, part)
        nd = docx.Document(part)
        nb = nd.element.body
        nsect = nb.find(qn('w:sectPr'))
        nblocks = [el for el in nb if el.tag in (qn('w:p'), qn('w:tbl'))]
        keep = set(id(el) for el in nblocks[i:i + chunk])
        for el in nblocks:
            if id(el) not in keep:
                nb.remove(el)
        # sectPr 永远保持在 body 末尾
        if nsect is not None:
            nb.remove(nsect)
            nb.append(nsect)
        atomic_save(nd.save, part)
        last = part
        outputs.append(part)
    detail = f'拆为 {n_parts} 份(每份 {chunk} 段)'
    return FileResult(path, 'ok', detail, last, outputs=outputs)


# ==================================================================== Excel
def xlsx_remove_images(path: str, options: dict) -> FileResult:
    import openpyxl

    wb = openpyxl.load_workbook(path)
    n = 0
    for ws in wb.worksheets:
        imgs = getattr(ws, '_images', [])
        n += len(imgs)
        ws._images = []
    if n == 0:
        return FileResult(path, 'skipped', '无图片')
    out, undo = _out(path, options, '_noimg')
    atomic_save(wb.save, out)
    return FileResult(path, 'ok', f'删除图片 {n} 张', out, undo=undo)


def xlsx_remove_header_footer(path: str, options: dict) -> FileResult:
    import openpyxl

    wb = openpyxl.load_workbook(path)
    for ws in wb.worksheets:
        for hf in (ws.oddHeader, ws.evenHeader, ws.oddFooter, ws.evenFooter):
            hf.center.text = None
            hf.left.text = None
            hf.right.text = None
    out, undo = _out(path, options, '_hf')
    atomic_save(wb.save, out)
    return FileResult(path, 'ok', '页眉页脚已清空', out, undo=undo)


def xlsx_set_header_footer(path: str, options: dict) -> FileResult:
    import openpyxl

    wb = openpyxl.load_workbook(path)
    h = options.get('header_text', '')
    f = options.get('footer_text', '')
    for ws in wb.worksheets:
        if h:
            ws.oddHeader.center.text = h
            ws.oddHeader.center.size = 12
        if f:
            ws.oddFooter.center.text = f
            ws.oddFooter.center.size = 10
    out, undo = _out(path, options, '_hf')
    atomic_save(wb.save, out)
    return FileResult(path, 'ok', '页眉页脚已设置', out, undo=undo)


def xlsx_remove_macros(path: str, options: dict) -> FileResult:
    """xlsx 不含宏; xlsm 另存为 xlsx 即剔除。"""
    import openpyxl

    base, ext = os.path.splitext(path)
    if ext.lower() != '.xlsm':
        return FileResult(path, 'skipped', 'xlsx 本身不含宏')
    wb = openpyxl.load_workbook(path, keep_vba=False)
    # 注意: _out 的查重针对 .xlsm 文件名, 改成 .xlsx 后必须再查一次,
    # 否则第二次运行会覆盖用户已有的同名 xlsx。
    out_dir = options.get('out_dir')
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
        out = _uniq_out(os.path.abspath(
            os.path.join(out_dir, os.path.basename(base) + '_nomacro.xlsx')))
        undo = []
    else:
        bak = _bak(path)
        out = _uniq_out(os.path.abspath(base + '_nomacro.xlsx'))
        undo = [{'src': bak, 'dst': path, 'kind': 'restore'}]
    atomic_save(wb.save, out)
    return FileResult(path, 'ok', '已另存为无宏 xlsx', out, undo=undo)


def xlsx_convert_format(path: str, options: dict) -> FileResult:
    """Excel → CSV / TXT / HTML。

    两个坑:
    - data_only=True 只能拿到"上一次由 Excel 保存时写入的缓存值"。openpyxl
      生成或第三方导出的 xlsx 没有缓存, 公式格全变成空字符串, 用户会以为
      "数据丢了"。这里同时打开不含 data_only 的副本, 缓存值缺失时回落成
      公式文本(=A1+A2), 至少不丢信息, 用 Excel 打开还能直接算。
    - 早期只转 wb.active 一张表, 多 Sheet 工作簿其余表整张丢失。现在逐表
      输出成 base_Sheet名.ext; 单 Sheet 工作簿仍沿用 base.ext, 不改变旧行为。
    """
    import openpyxl

    target = options.get('xlsx_target', 'csv')
    if target not in ('csv', 'txt', 'html'):
        return FileResult(path, 'skipped', '暂不支持 ' + str(target))
    wb_val = openpyxl.load_workbook(path, data_only=True)
    wb_raw = openpyxl.load_workbook(path, data_only=False)
    out_dir = options.get('out_dir') or os.path.dirname(path)
    os.makedirs(out_dir, exist_ok=True)
    base, _ext = os.path.splitext(os.path.basename(path))
    sheets = list(wb_val.worksheets)
    multi = len(sheets) > 1

    def _val(cell, raw):
        v = cell.value
        if v is None and raw is not None and raw.data_type == 'f':
            return raw.value          # 无缓存值 → 用公式文本兜底
        return v

    outs = []
    for idx, ws in enumerate(sheets):
        raw_ws = wb_raw.worksheets[idx] if idx < len(wb_raw.worksheets) else None
        name = base
        if multi:
            safe = re.sub(r'[\\/:*?"<>|]', '_', ws.title).strip() or f'sheet{idx + 1}'
            name = base + '_' + safe
        out = os.path.abspath(os.path.join(out_dir, name + '.' + target))
        out = _uniq_out(out)
        if target == 'csv':
            import csv as _csv
            raw_iter = raw_ws.iter_rows() if raw_ws is not None else None
            with open(out, 'w', encoding='utf-8-sig', newline='') as fh:
                w = _csv.writer(fh)
                for row in ws.iter_rows():
                    raw_cells = None
                    if raw_iter is not None:
                        try:
                            raw_cells = list(next(raw_iter))
                        except StopIteration:
                            raw_cells = None
                    vals = []
                    for ci, cell in enumerate(row):
                        raw = raw_cells[ci] if raw_cells and ci < len(raw_cells) \
                            else None
                        v = _val(cell, raw)
                        vals.append('' if v is None else v)
                    w.writerow(vals)
        elif target == 'txt':
            with open(out, 'w', encoding='utf-8') as fh:
                for row in ws.iter_rows():
                    vals = []
                    for cell in row:
                        v = cell.value
                        vals.append('' if v is None else str(v))
                    fh.write('\t'.join(vals) + '\n')
        else:
            from html import escape as _escape
            with open(out, 'w', encoding='utf-8') as fh:
                fh.write("<!DOCTYPE html>\n<html><head><meta charset='utf-8'>"
                         '<style>table{border-collapse:collapse}td,th{border:1px '
                         'solid #999;padding:4px 8px}</style></head><body>\n'
                         + (f'<h3>{_escape(ws.title)}</h3>\n' if multi else '')
                         + '<table>\n')
                for row in ws.iter_rows():
                    fh.write('<tr>')
                    for cell in row:
                        v = cell.value
                        fh.write('<td>' + ('' if v is None
                                           else _escape(str(v))) + '</td>')
                    fh.write('</tr>\n')
                fh.write('</table>\n</body></html>\n')
        outs.append(out)

    if not outs:
        return FileResult(path, 'skipped', '工作簿里没有可转换的工作表')
    msg = '→ ' + target.upper()
    if multi:
        msg += f' ({len(outs)} 个 Sheet)'
    return FileResult(path, 'ok', msg, outs[0])


def xlsx_split(path, options):
    import openpyxl
    from engine.workbooks import copy_sheet
    from engine.office_ops import _load_xlsx
    mode = options.get('xlsx_split_mode', 'sheet')
    out_dir = options.get('out_dir') or os.path.dirname(path)
    os.makedirs(out_dir, exist_ok=True)
    base, ext = os.path.splitext(os.path.basename(path))
    outputs = []
    wb = _load_xlsx(path)
    try:
        if mode == 'sheet':
            for source in wb.worksheets:
                # Clone the source package so pictures, charts, widths and print layout survive.
                split = _load_xlsx(path)
                target = split[source.title]
                has_cross_refs = any(c.data_type == 'f' and '!' in str(c.value)
                                     for row in target for c in row)
                for other in list(split.worksheets):
                    if other.title != source.title:
                        if has_cross_refs: other.sheet_state = 'hidden'
                        else: split.remove(other)
                target.sheet_state = 'visible'; split.active = split.index(target)
                title = re.sub(r'[\\/:*?"<>|]', '_', source.title)
                out = _uniq_out(os.path.join(out_dir, base + '_' + title + ext))
                atomic_save(split.save, out); split.close(); outputs.append(out)
        else:
            per = max(1, int(options.get('xlsx_rows_per', 1000) or 1000))
            source = wb.active
            for i in range(2, max(3, source.max_row + 1), per):
                split = openpyxl.Workbook()
                target = split.active; target.title = source.title
                rows = [1] + list(range(i, min(i+per, source.max_row+1)))
                copy_sheet(source, target, rows)
                out = _uniq_out(os.path.join(out_dir, base + f'_part{len(outputs)+1:03d}.xlsx'))
                atomic_save(split.save, out); split.close(); outputs.append(out)
    finally:
        wb.close()
    return FileResult(path, 'ok', f'拆为 {len(outputs)} 个文件；跨表公式的依赖表保留为隐藏表',
                      outputs[-1] if outputs else '', outputs=outputs)


def xlsx_merge_files(paths, out_path):
    from engine.workbooks import merge_workbooks
    return merge_workbooks(paths, out_path)


def xlsx_sheet_delete_replace(path: str, options: dict) -> FileResult:
    import openpyxl

    wb = openpyxl.load_workbook(path)
    s = str(options.get('sheet_names', '') or '')
    names = [x.strip() for x in s.split(',') if x.strip()]
    if not names:
        return FileResult(path, 'skipped', '未填写 Sheet 名(逗号分隔)')
    n = 0
    for nm in names:
        if nm in wb.sheetnames:
            wb.remove(wb[nm])
            n += 1
    if n == 0:
        return FileResult(path, 'skipped', '未找到指定 Sheet')
    if not wb.sheetnames:
        wb.create_sheet('Sheet1')
    out, undo = _out(path, options, '_sheet')
    atomic_save(wb.save, out)
    return FileResult(path, 'ok', f'删除 {n} 个 Sheet', out, undo=undo)


def xlsx_sheet_insert(path: str, options: dict) -> FileResult:
    import openpyxl

    wb = openpyxl.load_workbook(path)
    try:
        pos = int(options.get('sheet_pos', 1) or 1)
    except (TypeError, ValueError):
        return FileResult(path, 'skipped', '插入位置需为整数')
    pos = max(1, pos)
    nm = str(options.get('new_sheet_name', '') or '').strip() or 'Sheet'
    nm = nm[:31]
    # openpyxl 遇到重名会静默改成 "数据1", 用户指定的名称被篡改
    if nm in wb.sheetnames:
        return FileResult(path, 'fail', f'已存在名为「{nm}」的 Sheet')
    if set(nm) & set(r'[]:*?/\\'):
        return FileResult(path, 'fail', f'Sheet 名含非法字符: {nm}')
    wb.create_sheet(title=nm, index=min(pos - 1, len(wb.sheetnames)))
    out, undo = _out(path, options, '_sheet')
    atomic_save(wb.save, out)
    return FileResult(path, 'ok', f'在第 {pos} 位插入 {nm}', out, undo=undo)


def xlsx_extract_images(path: str, options: dict) -> FileResult:
    out_dir = options.get('out_dir') or (
        os.path.join(os.path.dirname(path),
                     os.path.splitext(os.path.basename(path))[0] + '_图片'))
    base = os.path.splitext(os.path.basename(path))[0]
    os.makedirs(out_dir, exist_ok=True)
    n = 0
    last = ''
    with zipfile.ZipFile(path) as zf:
        for item in zf.namelist():
            if re.match(r'xl/media/', item):
                fn = base + '_' + os.path.basename(item)
                dst = _uniq_out(os.path.abspath(os.path.join(out_dir, fn)))
                with open(dst, 'wb') as fh:
                    fh.write(zf.read(item))
                n += 1
                last = dst
    if n == 0:
        return FileResult(path, 'skipped', '无图片')
    return FileResult(path, 'ok', f'提取 {n} 张图片', last)


def xlsx_meta(path: str, options: dict) -> FileResult:
    import openpyxl

    wb = openpyxl.load_workbook(path)
    p = wb.properties
    mode = options.get('meta_mode', '清空')
    if mode == '清空':
        p.creator = p.lastModifiedBy = ''
        p.title = p.subject = p.keywords = p.description = ''
        detail = '元数据已清空'
    else:
        if options.get('meta_author'):
            p.creator = p.lastModifiedBy = options['meta_author']
        if options.get('meta_title'):
            p.title = options['meta_title']
        if options.get('meta_keywords'):
            p.keywords = options['meta_keywords']
        detail = '元数据已修改(作者=' + str(p.creator) + ')'
    out, undo = _out(path, options, '_meta')
    atomic_save(wb.save, out)
    return FileResult(path, 'ok', detail, out, undo=undo)


# ====================================================================== PPT
def pptx_remove_images(path: str, options: dict) -> FileResult:
    from pptx import Presentation

    prs = Presentation(path)
    n = 0
    for slide in prs.slides:
        for shape in list(slide.shapes):
            try:
                if shape.shape_type == 13:  # PICTURE
                    shape._element.getparent().remove(shape._element)
                    n += 1
            except Exception:
                continue
    if n == 0:
        return FileResult(path, 'skipped', '无图片')
    out, undo = _out(path, options, '_noimg')
    atomic_save(prs.save, out)
    return FileResult(path, 'ok', f'删除图片 {n} 张', out, undo=undo)


def pptx_remove_notes(path: str, options: dict) -> FileResult:
    from pptx import Presentation

    prs = Presentation(path)
    n = 0
    for slide in prs.slides:
        if slide.has_notes_slide:
            slide.notes_slide.notes_text_frame.clear()
            n += 1
    if n == 0:
        return FileResult(path, 'skipped', '无备注')
    out, undo = _out(path, options, '_nonotes')
    atomic_save(prs.save, out)
    return FileResult(path, 'ok', f'清空 {n} 页备注', out, undo=undo)


def pptx_extract_notes(path: str, options: dict) -> FileResult:
    from pptx import Presentation

    prs = Presentation(path)
    lines = []
    pages = 0
    for i, slide in enumerate(prs.slides, 1):
        if slide.has_notes_slide:
            t = slide.notes_slide.notes_text_frame.text.strip()
            if t:
                lines.append(f'===== 第 {i} 页 =====\n{t}')
                pages += 1
    if pages == 0:
        return FileResult(path, 'skipped', '无备注')
    out_dir = options.get('out_dir') or os.path.dirname(path)
    os.makedirs(out_dir, exist_ok=True)
    base = os.path.splitext(os.path.basename(path))[0]
    out = _uniq_out(os.path.abspath(os.path.join(out_dir, base + '_备注.txt')))
    with open(out, 'w', encoding='utf-8') as fh:
        fh.write('\n\n'.join(lines))
    return FileResult(path, 'ok', f'提取 {pages} 页备注', out)


def pptx_remove_header_footer(path: str, options: dict) -> FileResult:
    from pptx import Presentation

    prs = Presentation(path)
    n = 0
    HF_IDX = (10, 11, 12)
    for slide in prs.slides:
        for shape in list(slide.shapes):
            try:
                if shape.is_placeholder and \
                        shape.placeholder_format.idx in HF_IDX:
                    shape._element.getparent().remove(shape._element)
                    n += 1
            except Exception:
                continue
    if n == 0:
        return FileResult(path, 'skipped', '未发现页眉页脚占位符')
    out, undo = _out(path, options, '_hf')
    atomic_save(prs.save, out)
    return FileResult(path, 'ok', f'清理 {n} 处', out, undo=undo)


def _drop_bak(undo: list):
    """撤销 _out 刚做的备份。

    _out 是"先备份再改", 但改之前的判断可能得出 skipped / fail —— 那种情况
    下 .qingyanbak 就是纯垃圾文件 (早期 pptx 无宏、坏包两份都留下)。
    """
    for u in undo or []:
        src = (u or {}).get('src') or ''
        if src and os.path.exists(src):
            try:
                os.remove(src)
            except OSError:
                pass
    return []


def _zip_ok(path: str) -> bool:
    """能当 zip 打开才算有效的 Office 文件。"""
    try:
        with zipfile.ZipFile(path) as z:
            z.namelist()
        return True
    except (zipfile.BadZipFile, OSError):
        return False


def pptx_remove_macros(path: str, options: dict) -> FileResult:
    from engine.macro_cleanup import remove_macros
    return remove_macros(path, options, 'ppt')


def pptx_set_background(path: str, options: dict) -> FileResult:
    from pptx import Presentation
    from pptx.dml.color import RGBColor

    prs = Presentation(path)
    color = str(options.get('bg_color', '') or '').strip()
    remove = bool(options.get('bg_remove', False))
    if not remove and not re.match(r'^#?([0-9A-Fa-f]{6})$', color):
        return FileResult(path, 'skipped', '未填写背景色(#RRGGBB)')
    n = 0
    for slide in prs.slides:
        try:
            fill = slide.background.fill
            if remove:
                fill.background()
            else:
                m = re.match(r'^#?([0-9A-Fa-f]{6})$', color)
                rgb = RGBColor.from_string(m.group(1).upper())
                fill.solid()
                fill.fore_color.rgb = rgb
            n += 1
        except Exception:
            continue
    if n == 0:
        return FileResult(path, 'skipped', '背景未修改')
    out, undo = _out(path, options, '_bg')
    atomic_save(prs.save, out)
    detail = ('背景已重置 ' + str(n) + ' 页' if remove
              else '背景色 ' + color + ' x ' + str(n) + ' 页')
    return FileResult(path, 'ok', detail, out, undo=undo)


def _clone_slide(src_prs, src_slide, dst_prs):
    """XML 级克隆一页幻灯片到目标 prs。

    两个必须处理的点:
    1. 必须用【目标包自己的】slide_layout。直接传源包的 layout 会把源 layout
       part 挂进目标包, 与目标包同名 part 撞名, 产出带重复 zip 条目的非法包
       (PowerPoint 报"需要修复")。
    2. 只 deepcopy 形状 XML 会丢掉 r:embed/r:id 等关系引用, 图片/超链接/图表
       变成悬空引用。这里复制关系并把 rId 重映射到新页的关系表。
    """
    from pptx.oxml.ns import qn
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT

    # 同名 layout 优先按源 layout 在其母版中的序号取目标包的 layout
    layout = None
    try:
        idx = list(src_slide.slide_layout.slide_master.slide_layouts).index(
            src_slide.slide_layout)
        if idx < len(dst_prs.slide_layouts):
            layout = dst_prs.slide_layouts[idx]
    except Exception:
        layout = None
    if layout is None:
        n = len(dst_prs.slide_layouts)
        layout = dst_prs.slide_layouts[min(6, n - 1)] if n else None
    new_slide = dst_prs.slides.add_slide(layout)
    for shp in list(new_slide.shapes):
        shp._element.getparent().remove(shp._element)

    # 复制源页关系(排除 layout), 建立 旧rId -> 新rId 映射
    rid_map: dict = {}
    try:
        for rid, rel in src_slide.part.rels.items():
            if rel.reltype == RT.SLIDE_LAYOUT:
                continue
            try:
                if rel.is_external:
                    new_rid = new_slide.part.rels.get_or_add_ext_rel(
                        rel.reltype, rel.target_ref)
                else:
                    new_rid = new_slide.part.relate_to(rel.target_part, rel.reltype)
                rid_map[rid] = new_rid
            except Exception:
                continue
    except Exception:
        pass

    r_attrs = ('r:embed', 'r:id', 'r:link', 'r:pict', 'r:dm', 'r:lo',
               'r:qs', 'r:cs', 'r:href')
    for el in src_slide.shapes._spTree:
        tag = el.tag
        if tag.endswith('}sp') or tag.endswith('}pic') or \
                tag.endswith('}graphicFrame') or tag.endswith('}grpSp') or \
                tag.endswith('}cxnSp'):
            new_el = deepcopy(el)
            if rid_map:
                for e in new_el.iter():
                    for a in r_attrs:
                        try:
                            v = e.get(qn(a))
                        except Exception:
                            v = None
                        if v and v in rid_map:
                            e.set(qn(a), rid_map[v])
            new_slide.shapes._spTree.append(new_el)
    return new_slide


def pptx_split(path: str, options: dict) -> FileResult:
    from pptx import Presentation

    per = max(1, int(options.get('ppt_chunk', 1) or 1))
    prs = Presentation(path)
    total = len(prs.slides._sldIdLst)
    out_dir = options.get('out_dir') or os.path.dirname(path)
    os.makedirs(out_dir, exist_ok=True)
    base = os.path.splitext(os.path.basename(path))[0]
    slides = list(prs.slides)
    n = 0
    last = ''
    outputs = []
    for i in range(0, total, per):
        nprs = Presentation()
        nprs.slide_width = prs.slide_width
        nprs.slide_height = prs.slide_height
        for sl in slides[i:i + per]:
            _clone_slide(prs, sl, nprs)
        out = os.path.abspath(
            os.path.join(out_dir, f'{base}_part{n + 1:03d}.pptx'))
        out = _uniq_out(out)
        atomic_save(nprs.save, out)
        outputs.append(out)
        n += 1
        last = out
    if n == 0:
        return FileResult(path, 'skipped', '无幻灯片')
    return FileResult(path, 'ok', f'拆为 {n} 个文件(每份 {per} 页)', last, outputs=outputs)


def pptx_merge_files(paths: list, out_path: str) -> int:
    """按顺序合并 PPT, 返回总页数。"""
    from pptx import Presentation

    first = paths[0] if paths else None
    if not first:
        return 0
    wb = Presentation(first)
    for p in paths[1:]:
        try:
            src = Presentation(p)
        except Exception:
            continue
        for sl in src.slides:
            _clone_slide(src, sl, wb)
    atomic_save(wb.save, out_path)
    return len(wb.slides._sldIdLst)


def pptx_extract_slides(path: str, options: dict) -> FileResult:
    from pptx import Presentation

    prs = Presentation(path)
    total = len(prs.slides._sldIdLst)
    if total == 0:
        return FileResult(path, 'skipped', '没有幻灯片')
    try:
        a = max(1, int(options.get('page_from', 1) or 1))
        b = int(options.get('page_to', 0) or 0)
    except (TypeError, ValueError):
        return FileResult(path, 'fail', '页码必须是整数')
    b = total if b <= 0 else min(total, b)     # 0/留空 = 到末尾
    if a > b:
        # 否则切片为空, 会产出 0 页 pptx(PowerPoint 无法打开)却报 ok
        return FileResult(path, 'fail',
                          f'起始页 {a} 大于结束页 {b}(共 {total} 页)')
    out_dir = options.get('out_dir') or os.path.dirname(path)
    os.makedirs(out_dir, exist_ok=True)
    base = os.path.splitext(os.path.basename(path))[0]
    nprs = Presentation()
    nprs.slide_width = prs.slide_width
    nprs.slide_height = prs.slide_height
    for sl in list(prs.slides)[a - 1:b]:
        _clone_slide(prs, sl, nprs)
    out = os.path.abspath(os.path.join(out_dir, f'{base}_p{a}-{b}.pptx'))
    nprs.save(out)
    return FileResult(path, 'ok', f'提取第 {a}-{b} 页', out)


def pptx_extract_images(path: str, options: dict) -> FileResult:
    out_dir = options.get('out_dir') or (
        os.path.join(os.path.dirname(path),
                     os.path.splitext(os.path.basename(path))[0] + '_图片'))
    base = os.path.splitext(os.path.basename(path))[0]
    os.makedirs(out_dir, exist_ok=True)
    n = 0
    last = ''
    with zipfile.ZipFile(path) as zf:
        for item in zf.namelist():
            if re.match(r'ppt/media/', item):
                fn = base + '_' + os.path.basename(item)
                dst = _uniq_out(os.path.abspath(os.path.join(out_dir, fn)))
                with open(dst, 'wb') as fh:
                    fh.write(zf.read(item))
                n += 1
                last = dst
    if n == 0:
        return FileResult(path, 'skipped', '无图片')
    return FileResult(path, 'ok', f'提取 {n} 张图片', last)


def pptx_compress(path: str, options: dict) -> FileResult:
    if not _zip_ok(path):
        return FileResult(path, 'fail', '文件已损坏或不是有效的 pptx')
    so = os.path.getsize(path)      # 必须在覆盖前取
    out, undo = _out(path, options, '_c')
    tmp_out = out + '.tmpzip'
    try:
        with zipfile.ZipFile(path) as zin, \
                zipfile.ZipFile(tmp_out, 'w', zipfile.ZIP_DEFLATED,
                                compresslevel=9) as zout:
            for item in zin.namelist():
                zout.writestr(item, zin.read(item))
        shutil.move(tmp_out, out)
    except (zipfile.BadZipFile, OSError) as e:
        return FileResult(path, 'fail', '压缩失败: ' + str(e)[:100],
                          undo=_drop_bak(undo))
    finally:
        if os.path.exists(tmp_out):
            try:
                os.remove(tmp_out)
            except OSError:
                pass
    sn = os.path.getsize(out)
    pct = (so - sn) / so * 100 if so else 0
    return FileResult(path, 'ok',
                      f'{so} → {sn} 字节({"减少" if pct >= 0 else "增加"} '
                      f'{abs(pct):.1f}%)', out, undo=undo)


def pptx_meta(path: str, options: dict) -> FileResult:
    from pptx import Presentation

    prs = Presentation(path)
    p = prs.core_properties
    mode = options.get('meta_mode', '清空')
    if mode == '清空':
        p.author = p.last_modified_by = ''
        p.title = p.subject = p.keywords = p.comments = ''
        detail = '元数据已清空'
    else:
        if options.get('meta_author'):
            p.author = p.last_modified_by = options['meta_author']
        if options.get('meta_title'):
            p.title = options['meta_title']
        detail = '元数据已修改(作者=' + str(p.author) + ')'
    out, undo = _out(path, options, '_meta')
    atomic_save(prs.save, out)
    return FileResult(path, 'ok', detail, out, undo=undo)


# ====================================================================== PDF
def _pdf_reader(path: str, options: dict):
    """打开 PDF; 已加密时用 options['password'] 解锁。

    pypdf 的 decrypt() 口令错误时不抛异常, 只返回 PasswordType.NOT_DECRYPTED,
    必须显式判定, 否则 FileNotDecryptedError 会在访问 .pages 时逃逸。
    """
    from pypdf import PdfReader
    reader = PdfReader(path)
    if reader.is_encrypted:
        try:
            res = reader.decrypt(options.get('password') or '')
        except Exception as e:
            raise RuntimeError('PDF 解析失败: ' + str(e)) from e
        # 只认 decrypt() 返回值: 口令正确时 is_encrypted 仍可能为 True
        if not res:
            raise RuntimeError('PDF 已加密且口令不对')
    return reader


def pdf_delete_blank_pages(path: str, options: dict) -> FileResult:
    from pypdf import PdfWriter

    try:
        reader = _pdf_reader(path, options)
    except RuntimeError as e:
        return FileResult(path, 'fail', str(e))
    w = PdfWriter()
    removed = 0
    for pg in reader.pages:
        text = ''
        try:
            text = (pg.extract_text() or '').strip()
        except Exception:
            pass
        has_img = False
        try:
            res = pg.get('/Resources')
            if res and '/XObject' in res:
                xo = res['/XObject'].get_object()
                has_img = len(xo) > 0
        except Exception:
            has_img = False
        has_annot = False
        try:
            has_annot = '/Annots' in pg
        except Exception:
            pass
        if not text and not has_img and not has_annot:
            removed += 1
            continue
        w.add_page(pg)
    if removed == 0:
        return FileResult(path, 'skipped', '无空白页')
    if len(w.pages) == 0:
        return FileResult(path, 'fail',
                          '所有页面均为空白, 已中止(会产生无效 PDF)')
    out, undo = _out(path, options, '_nob')
    with open(out, 'wb') as fh:
        w.write(fh)
    return FileResult(path, 'ok', f'删除空白页 {removed} 页', out, undo=undo)


def pdf_delete_pages(path: str, options: dict) -> FileResult:
    from pypdf import PdfWriter

    try:
        reader = _pdf_reader(path, options)
    except RuntimeError as e:
        return FileResult(path, 'fail', str(e))
    total = len(reader.pages)
    spec = str(options.get('pages_spec', '') or '').strip()
    if not spec:
        return FileResult(path, 'skipped', '未填写页码(如 1,3,5-7)')
    kill = set()
    for part in spec.split(','):
        part = part.strip()
        if not part:
            continue
        if '-' in part:
            a, _, b = part.partition('-')
            if a.strip().isdigit() and b.strip().isdigit():
                lo = max(1, min(total, int(a)))
                hi = max(1, min(total, int(b)))
                kill.update(range(lo, hi + 1))
        elif part.isdigit():
            p = int(part)
            if 1 <= p <= total:
                kill.add(p)
    if not kill:
        return FileResult(path, 'fail', '页码范围无效')
    if len(kill) >= total:
        return FileResult(path, 'fail',
                          '所选页码会删除全部页面, 已中止(会产生无效 PDF)')
    w = PdfWriter()
    for i, pg in enumerate(reader.pages, 1):
        if i not in kill:
            w.add_page(pg)
    out, undo = _out(path, options, '_pages')
    with open(out, 'wb') as fh:
        w.write(fh)
    return FileResult(path, 'ok', f'删除 {len(kill)} 页', out, undo=undo)


def pdf_extract_images(path: str, options: dict) -> FileResult:
    try:
        reader = _pdf_reader(path, options)
    except RuntimeError as e:
        return FileResult(path, 'fail', str(e))
    out_dir = options.get('out_dir') or (
        os.path.join(os.path.dirname(path),
                     os.path.splitext(os.path.basename(path))[0] + '_图片'))
    base = os.path.splitext(os.path.basename(path))[0]
    os.makedirs(out_dir, exist_ok=True)
    n = 0
    last = ''
    for pi, pg in enumerate(reader.pages, 1):
        try:
            for img in pg.images:
                fn = f'{base}_p{pi}_{img.name}'
                with open(os.path.join(out_dir, fn), 'wb') as fh:
                    fh.write(img.data)
                n += 1
                last = os.path.join(out_dir, fn)
        except Exception:
            continue
    if n == 0:
        return FileResult(path, 'skipped', '未提取到图片')
    return FileResult(path, 'ok', f'提取 {n} 张图片', last)


def pdf_meta(path: str, options: dict) -> FileResult:
    from pypdf import PdfWriter

    try:
        reader = _pdf_reader(path, options)
    except RuntimeError as e:
        return FileResult(path, 'fail', str(e))
    w = PdfWriter()
    w.append(reader)
    mode = options.get('meta_mode', '清空')
    if mode == '清空':
        w.add_metadata({'/Producer': '', '/Creator': '', '/Author': '',
                        '/Title': '', '/Subject': '', '/Keywords': ''})
        detail = '元数据已清空'
    else:
        meta = {}
        if options.get('meta_author'):
            meta['/Author'] = options['meta_author']
        if options.get('meta_title'):
            meta['/Title'] = options['meta_title']
        if options.get('meta_subject'):
            meta['/Subject'] = options['meta_subject']
        w.add_metadata(meta)
        detail = '元数据已修改'
    out, undo = _out(path, options, '_meta')
    with open(out, 'wb') as fh:
        w.write(fh)
    return FileResult(path, 'ok', detail, out, undo=undo)


def pdf_compress(path: str, options: dict) -> FileResult:
    from pypdf import PdfWriter

    try:
        reader = _pdf_reader(path, options)
    except RuntimeError as e:
        return FileResult(path, 'fail', str(e))
    w = PdfWriter()
    for pg in reader.pages:
        w.add_page(pg)
    try:
        w.compress_identical_objects(remove_identicals=True,
                                     remove_orphans=True)
    except Exception:
        pass
    so = os.path.getsize(path)      # 必须在覆盖前取
    out, undo = _out(path, options, '_c')
    with open(out, 'wb') as fh:
        w.write(fh)
    sn = os.path.getsize(out)
    pct = (so - sn) / so * 100 if so else 0
    return FileResult(path, 'ok',
                      f'{so} → {sn} 字节({"减少" if pct >= 0 else "增加"} '
                      f'{abs(pct):.1f}%)', out, undo=undo)


def pdf_add_link(path: str, options: dict) -> FileResult:
    from pypdf import PdfWriter
    from pypdf.annotations import Link

    try:
        reader = _pdf_reader(path, options)
    except RuntimeError as e:
        return FileResult(path, 'fail', str(e))
    url = str(options.get('link_url', '') or '').strip()
    if not url:
        return FileResult(path, 'skipped', '未填写网址')
    w = PdfWriter()
    w.append(reader)
    for i, pg in enumerate(w.pages):
        box = pg.mediabox
        ann = Link(rect=(0, 0, float(box.width), float(box.height)), url=url)
        w.add_annotation(page_number=i, annotation=ann)
    out, undo = _out(path, options, '_link')
    with open(out, 'wb') as fh:
        w.write(fh)
    return FileResult(path, 'ok', '全页链接 ' + url, out, undo=undo)


def pdf_remove_links(path: str, options: dict) -> FileResult:
    from pypdf import PdfWriter
    from pypdf.generic import ArrayObject, NameObject

    try:
        reader = _pdf_reader(path, options)
    except RuntimeError as e:
        return FileResult(path, 'fail', str(e))
    w = PdfWriter()
    n = 0
    for pg in reader.pages:
        annots = pg.get('/Annots')
        if annots:
            keep = ArrayObject()
            for a in annots:
                obj = a.get_object()
                if obj.get('/Subtype') == NameObject('/Link'):
                    n += 1
                    continue
                keep.append(a)
            if keep:
                pg[NameObject('/Annots')] = keep
            else:
                pg.pop(NameObject('/Annots'), None)
        w.add_page(pg)
    if n == 0:
        return FileResult(path, 'skipped', '无链接注释')
    out, undo = _out(path, options, '_nolink')
    with open(out, 'wb') as fh:
        w.write(fh)
    return FileResult(path, 'ok', f'移除 {n} 个链接', out, undo=undo)


def pdf_crop(path: str, options: dict) -> FileResult:
    from pypdf import PdfWriter
    from pypdf.generic import RectangleObject

    try:
        reader = _pdf_reader(path, options)
    except RuntimeError as e:
        return FileResult(path, 'fail', str(e))
    raw = str(options.get('crop_margin', '0,0,0,0') or '')
    try:
        l, t, r, b = (float(x.strip()) for x in raw.split(','))
    except Exception:
        return FileResult(path, 'skipped', '裁剪边距格式: 左,上,右,下(pt)')
    if min(l, t, r, b) < 0:
        return FileResult(path, 'fail', '裁剪边距不能为负数')
    w = PdfWriter()
    for i, pg in enumerate(reader.pages, 1):
        mb = pg.mediabox
        x0, y0 = float(mb.left) + l, float(mb.bottom) + b
        x1, y1 = float(mb.right) - r, float(mb.top) - t
        # left>right 或 bottom>top 会产出非法 cropbox, 后续阅读器行为未定义
        if x1 - x0 < 1 or y1 - y0 < 1:
            return FileResult(path, 'fail',
                              f'裁剪边距过大(第 {i} 页裁剪后不足 1pt)')
        pg.cropbox = RectangleObject((x0, y0, x1, y1))
        w.add_page(pg)
    out, undo = _out(path, options, '_crop')
    with open(out, 'wb') as fh:
        w.write(fh)
    return FileResult(path, 'ok', f'四边裁剪 {l}/{t}/{r}/{b} pt', out, undo=undo)


def pdf_resize(path: str, options: dict) -> FileResult:
    from pypdf import PdfWriter
    from pypdf.generic import RectangleObject

    try:
        reader = _pdf_reader(path, options)
    except RuntimeError as e:
        return FileResult(path, 'fail', str(e))
    sizes = {'A4': (595.28, 841.89), 'A3': (841.89, 1190.55),
             'A5': (419.53, 595.28), 'Letter': (612, 792)}
    target = str(options.get('page_size', 'A4') or 'A4')
    # page_size 会拼进输出文件名, 必须走白名单, 否则 '..\\..\\x' 可逃出输出目录
    if target not in sizes:
        return FileResult(path, 'skipped', '不支持的纸张规格: ' + target)
    tw, th = sizes[target]
    w = PdfWriter()
    for pg in reader.pages:
        pw = float(pg.mediabox.width)
        ph = float(pg.mediabox.height)
        scale = min(tw / pw, th / ph)
        pg.scale_by(scale)
        pg.mediabox = RectangleObject((0, 0, tw, th))
        w.add_page(pg)
    out, undo = _out(path, options, '_' + target)
    with open(out, 'wb') as fh:
        w.write(fh)
    return FileResult(path, 'ok', '统一为 ' + target, out, undo=undo)


def pdf_grayscale(path: str, options: dict) -> FileResult:
    """彩色转黑白: 逐页栅格化需要渲染引擎 — 如装了 PyMuPDF 则真转灰度。"""
    try:
        import fitz  # PyMuPDF
    except Exception:
        return FileResult(path, 'skipped',
                          '彩色转黑白需渲染引擎(PyMuPDF), 未安装时无法执行')
    out, undo = _out(path, options, '_gray')
    doc = fitz.open(path)
    ndoc = fitz.open()
    for pg in doc:
        pix = pg.get_pixmap(dpi=150, colorspace=fitz.csGRAY)
        npg = ndoc.new_page(width=pg.rect.width, height=pg.rect.height)
        npg.insert_image(npg.rect, pixmap=pix)
    ndoc.save(out, garbage=3, deflate=True)
    ndoc.close()
    doc.close()
    return FileResult(path, 'ok', '已转灰度(150dpi 栅格化)', out, undo=undo)


def pdf_remove_header_footer(path: str, options: dict) -> FileResult:
    """通过 cropbox 裁掉顶部/底部条带。"""
    from pypdf import PdfWriter
    from pypdf.generic import RectangleObject

    try:
        reader = _pdf_reader(path, options)
    except RuntimeError as e:
        return FileResult(path, 'fail', str(e))
    try:
        top = float(options.get('hf_top', 0) or 0)
        bottom = float(options.get('hf_bottom', 0) or 0)
    except (TypeError, ValueError):
        return FileResult(path, 'skipped', '页眉/页脚高度需为数字(pt)')
    if top <= 0 and bottom <= 0:
        return FileResult(path, 'skipped', '填写要裁掉的页眉/页脚高度(pt)')
    w = PdfWriter()
    for i, pg in enumerate(reader.pages, 1):
        mb = pg.mediabox
        if float(mb.height) - top - bottom < 1:
            return FileResult(path, 'fail',
                              f'页眉+页脚高度超过第 {i} 页页高, 已中止')
        pg.cropbox = RectangleObject((mb.left, mb.bottom + bottom,
                                      mb.right, mb.top - top))
        w.add_page(pg)
    out, undo = _out(path, options, '_hf')
    with open(out, 'wb') as fh:
        w.write(fh)
    return FileResult(path, 'ok', f'裁掉顶部{top:g}pt/底部{bottom:g}pt',
                      out, undo=undo)


def pdf_add_seal(path: str, options: dict) -> FileResult:
    """骑缝章: 将印章图片按页竖切, 依次贴到每页右边缘。"""
    from pypdf import PdfReader, PdfWriter
    from PIL import Image
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas as rl_canvas

    try:
        reader = _pdf_reader(path, options)
    except RuntimeError as e:
        return FileResult(path, 'fail', str(e))
    seal_path = str(options.get('seal_image', '') or '').strip()
    if not seal_path or not os.path.exists(seal_path):
        return FileResult(path, 'skipped', '未选择印章图片')
    pages = list(reader.pages)
    total = len(pages)
    if total == 0:
        return FileResult(path, 'skipped', 'PDF 没有页面')
    seal = Image.open(seal_path).convert('RGBA')
    sw, sh = seal.size
    strip_w = max(1, sw // total)
    w = PdfWriter()
    for i, pg in enumerate(pages):
        piece = seal.crop((i * strip_w, 0, (i + 1) * strip_w, sh))
        _io = io.BytesIO()
        piece.save(_io, 'PNG')
        _io.seek(0)
        pw = float(pg.mediabox.width)
        ph = float(pg.mediabox.height)
        draw_h = ph * 0.25
        draw_w = draw_h * (strip_w / sh)
        buf = io.BytesIO()
        c = rl_canvas.Canvas(buf, pagesize=(pw, ph))
        c.drawImage(ImageReader(_io), pw - draw_w, (ph - draw_h) / 2,
                    width=draw_w, height=draw_h, mask='auto')
        c.save()
        buf.seek(0)
        overlay = PdfReader(buf).pages[0]
        pg.merge_page(overlay)
        w.add_page(pg)
    out, undo = _out(path, options, '_seal')
    with open(out, 'wb') as fh:
        w.write(fh)
    return FileResult(path, 'ok', f'骑缝章已加({total} 页)', out, undo=undo)


# ==================================================================== 文本
def _sniff(path: str) -> str:
    try:
        return detect_encoding(path)
    except Exception:
        return 'utf-8'


def _text_out(path: str, options: dict, enc: str, text: str):
    """文本类输出: 保 CRLF、防覆盖、可撤销。返回 (out, undo)。

    - 读取用 splitlines 丢掉了行尾信息: 原文件是 CRLF 必须原样还原,
      否则 .bat/.ini/.sln 被"顺手"改成 LF 还报 ok。
    - out_dir 指向源文件所在目录时等同就地修改, 必须走备份+撤销,
      否则源文件被无提示覆盖且无法回退。
    - 其他输出目录里已有同名文件时改名 (1), 不静默覆盖。
    """
    out_dir = options.get('out_dir')
    if not out_dir:
        bak = _bak(path)
        return path, [{'src': bak, 'dst': path, 'kind': 'restore'}]
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, os.path.basename(path))
    if os.path.abspath(out) == os.path.abspath(path):
        bak = _bak(path)
        return path, [{'src': bak, 'dst': path, 'kind': 'restore'}]
    if os.path.exists(out):
        b, e = os.path.splitext(out)
        k = 1
        while os.path.exists(f'{b}({k}){e}'):
            k += 1
        out = f'{b}({k}){e}'
    return out, []


def _orig_eol(path: str) -> str:
    try:
        with open(path, 'rb') as fh:
            return '\r\n' if b'\r\n' in fh.read(65536) else '\n'
    except OSError:
        return '\n'


def text_delete_replace_lines(path: str, options: dict) -> FileResult:
    enc = _sniff(path)
    eol = _orig_eol(path)
    with open(path, 'r', encoding=enc, errors='replace',
              newline='') as fh:
        raw = fh.read()
    norm = raw.replace('\r\n', '\n').replace('\r', '\n')
    lines = norm.split('\n')
    tail = lines[-1] == '' if lines else False
    if tail:
        lines = lines[:-1]
    mode = options.get('line_mode', '删除')
    spec = str(options.get('lines_spec', '') or '').strip()
    match = str(options.get('line_match', '') or '').strip()
    new_text = options.get('line_new', '')

    def hit(i, ln):
        if spec:
            for part in spec.split(','):
                part = part.strip()
                if '-' in part:
                    a, _, b = part.partition('-')
                    if a.strip().isdigit() and b.strip().isdigit():
                        if int(a) <= i + 1 <= int(b):
                            return True
                elif part.isdigit() and int(part) == i + 1:
                    return True
            return False
        if match:
            return match in ln
        return False

    out_lines = []
    n = 0
    for i, ln in enumerate(lines):
        if hit(i, ln):
            n += 1
            if mode == '替换':
                out_lines.append(new_text)
            continue
        out_lines.append(ln)
    if n == 0:
        return FileResult(path, 'skipped', '没有命中的行')
    out, undo = _text_out(path, options, enc, '')
    text = eol.join(out_lines) + (eol if tail else '')
    tmp = out + '.qytmp'
    with open(tmp, 'w', encoding=enc, newline='') as fh:
        fh.write(text)
    _replace(tmp, out)
    return FileResult(path, 'ok', f'{mode} {n} 行', out, undo=undo)


def text_insert_lines(path: str, options: dict) -> FileResult:
    enc = _sniff(path)
    eol = _orig_eol(path)
    with open(path, 'r', encoding=enc, errors='replace',
              newline='') as fh:
        raw = fh.read()
    norm = raw.replace('\r\n', '\n').replace('\r', '\n')
    lines = norm.split('\n')
    tail = lines[-1] == '' if lines else False
    if tail:
        lines = lines[:-1]
    pos = max(1, int(options.get('insert_pos', 1) or 1))
    content = str(options.get('insert_text', '') or '')
    ins = content.split('\\n') if '\\n' in content else content.split('\n')
    out_lines = lines[:pos - 1] + ins + lines[pos - 1:]
    out, undo = _text_out(path, options, enc, '')
    text = eol.join(out_lines) + (eol if tail else '')
    tmp = out + '.qytmp'
    with open(tmp, 'w', encoding=enc, newline='') as fh:
        fh.write(text)
    _replace(tmp, out)
    return FileResult(path, 'ok', f'第 {pos} 行前插入 {len(ins)} 行', out,
                      undo=undo)


def text_md_to_docx(path: str, options: dict) -> FileResult:
    import docx
    from docx.shared import Pt

    enc = _sniff(path)
    with open(path, 'r', encoding=enc, errors='replace') as fh:
        md = fh.read()
    d = docx.Document()
    in_code = False
    for ln in md.splitlines():
        s = ln.strip()
        if s.startswith('```'):
            in_code = not in_code
            continue
        if in_code:
            p = d.add_paragraph()
            r = p.add_run(ln)
            r.font.name = 'Consolas'
            r.font.size = Pt(9)
            continue
        if s.startswith('### '):
            d.add_heading(s[4:], level=3)
        elif s.startswith('## '):
            d.add_heading(s[3:], level=2)
        elif s.startswith('# '):
            d.add_heading(s[2:], level=1)
        elif s.startswith('- ') or s.startswith('* '):
            d.add_paragraph(s[2:], style='List Bullet')
        elif s:
            segs = re.split(r'(\*\*.+?\*\*|\*.+?\*)', s)
            p = d.add_paragraph()
            for sgm in segs:
                if not sgm:
                    continue
                if sgm.startswith('**') and sgm.endswith('**') and len(sgm) > 4:
                    r = p.add_run(sgm[2:-2])
                    r.bold = True
                elif sgm.startswith('*') and sgm.endswith('*') and len(sgm) > 2:
                    r = p.add_run(sgm[1:-1])
                    r.italic = True
                else:
                    p.add_run(sgm)
    out_dir = options.get('out_dir') or os.path.dirname(path)
    os.makedirs(out_dir, exist_ok=True)
    base = os.path.splitext(os.path.basename(path))[0]
    out = _uniq_out(os.path.abspath(os.path.join(out_dir, base + '.docx')))
    atomic_save(d.save, out)
    return FileResult(path, 'ok', 'md → docx', out)


def text_html_to_docx(path: str, options: dict) -> FileResult:
    import docx
    from html.parser import HTMLParser

    enc = _sniff(path)
    with open(path, 'r', encoding=enc, errors='replace') as fh:
        raw = fh.read()

    class P(HTMLParser):
        def __init__(self):
            super().__init__()
            self.lines = []
            self.in_body = False

        def handle_starttag(self, tag, attrs):
            if tag == 'body':
                self.in_body = True

        def handle_data(self, data):
            if self.in_body:
                s = data.strip()
                if s:
                    self.lines.append(s)

    p = P()
    p.feed(raw)
    d = docx.Document()
    for ln in p.lines:
        d.add_paragraph(ln)
    out_dir = options.get('out_dir') or os.path.dirname(path)
    os.makedirs(out_dir, exist_ok=True)
    base = os.path.splitext(os.path.basename(path))[0]
    out = _uniq_out(os.path.abspath(os.path.join(out_dir, base + '.docx')))
    atomic_save(d.save, out)
    return FileResult(path, 'ok', 'html → docx', out)


def text_json_convert(path: str, options: dict) -> FileResult:
    import json as _json

    enc = _sniff(path)
    with open(path, 'r', encoding=enc, errors='replace') as fh:
        try:
            obj = _json.load(fh)
        except Exception as e:
            return FileResult(path, 'fail', 'JSON 无效: ' + str(e))
    target = options.get('json_target', 'txt')
    out_dir = options.get('out_dir') or os.path.dirname(path)
    os.makedirs(out_dir, exist_ok=True)
    base = os.path.splitext(os.path.basename(path))[0]
    if target == 'csv':
        import csv as _csv
        out = _uniq_out(os.path.abspath(os.path.join(out_dir, base + '.csv')))
        with open(out, 'w', encoding='utf-8-sig', newline='') as fh:
            w = _csv.writer(fh)
            if isinstance(obj, list) and obj and isinstance(obj[0], dict):
                headers = list(obj[0].keys())
                w.writerow(headers)
                for row in obj:
                    w.writerow([row.get(h, '') for h in headers])
            elif isinstance(obj, dict):
                w.writerow(('key', 'value'))
                for k, v in obj.items():
                    w.writerow((k, _json.dumps(v, ensure_ascii=False)))
            else:
                w.writerow(('value',))
                for v in (obj if isinstance(obj, list) else [obj]):
                    w.writerow((_json.dumps(v, ensure_ascii=False),))
    else:
        out = _uniq_out(os.path.abspath(os.path.join(out_dir, base + '.txt')))
        with open(out, 'w', encoding='utf-8') as fh:
            if isinstance(obj, dict):
                for k, v in obj.items():
                    fh.write(f'{k}\t{_json.dumps(v, ensure_ascii=False)}\n')
            else:
                fh.write(_json.dumps(obj, ensure_ascii=False, indent=2))
    return FileResult(path, 'ok', 'json → ' + target, out)


def text_replace_by_excel(path: str, options: dict) -> FileResult:
    import openpyxl

    rule_file = str(options.get('rules_file', '') or '')
    if not rule_file or not os.path.exists(rule_file):
        return FileResult(path, 'skipped', '未选择规则 Excel')
    wb = openpyxl.load_workbook(rule_file, read_only=True)
    ws = wb.active
    mapping = {}
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        if len(row) < 2:
            continue
        k, v = row[0], row[1]
        # 空行/误填的数字 0 会被 str() 变成 '' / '0', 变成一条命中全文的规则
        # (原文 "A0B0C" 里的 0 全被替换)。只有非空、非零的键才算规则。
        if k is None or v is None:
            continue
        ks = str(k).strip()
        if not ks or ks in ('0', '0.0'):
            continue
        if isinstance(k, float) and k == int(k):
            ks = str(int(k))       # 10.0 -> "10", 否则规则永远匹配不上
        mapping[ks] = str(v)
    wb.close()
    if not mapping:
        return FileResult(path, 'skipped', '无规则')
    enc = _sniff(path)
    with open(path, 'r', encoding=enc, errors='replace',
              newline='') as fh:
        content = fh.read()
    # 逐条 replace 会让前面的结果被后面的规则再替换(分公司→集团, 集团→
    # 有限公司, 最终"分公司"变"有限公司")。单遍正则同时替换才符合直觉。
    keys = sorted(mapping.keys(), key=len, reverse=True)
    pattern = re.compile('|'.join(re.escape(k) for k in keys))
    n = len(pattern.findall(content))
    content = pattern.sub(lambda m: mapping[m.group(0)], content)
    out, undo = _text_out(path, options, enc, content)
    tmp = out + '.qytmp'
    with open(tmp, 'w', encoding=enc, newline='') as fh:
        fh.write(content)
    _replace(tmp, out)
    return FileResult(path, 'ok', f'按 {len(mapping)} 条规则替换 {n} 处', out,
                      undo=undo)


# ==================================================================== 图片
def image_erase(path: str, options: dict) -> FileResult:
    from PIL import Image, ImageDraw

    ext = os.path.splitext(path)[1].lower()
    IMG = {'.jpg', '.jpeg', '.png', '.bmp', '.webp', '.tif', '.tiff'}
    if ext not in IMG:
        return FileResult(path, 'skipped', '非支持图片')
    raw = str(options.get('erase_rect', '') or '').strip()
    try:
        x, y, w, h = (int(v.strip()) for v in raw.split(','))
    except Exception:
        return FileResult(path, 'fail', '擦除区域格式: x,y,宽,高(像素)')
    if w <= 0 or h <= 0:
        return FileResult(path, 'fail', '宽和高必须为正整数')
    mode = options.get('erase_mode', '纯色')
    im = Image.open(path)
    if x < 0 or y < 0 or x + w > im.width or y + h > im.height:
        return FileResult(path, 'fail',
                          f'区域超出图片范围(图片 {im.width}x{im.height})')
    region = im.crop((x, y, x + w, y + h))
    if mode == '马赛克':
        small = region.resize((max(1, w // 12), max(1, h // 12)),
                              Image.NEAREST)
        region = small.resize((w, h), Image.NEAREST)
        im.paste(region, (x, y))
    else:
        color = str(options.get('erase_color', '#FFFFFF') or '#FFFFFF')
        m = re.match(r'^#?([0-9A-Fa-f]{6})$', color)
        fill = (255, 255, 255)
        if m:
            hexv = m.group(1)
            fill = tuple(int(hexv[i:i + 2], 16) for i in (0, 2, 4))
        d = ImageDraw.Draw(im)
        # 灰度/调色板图只接受标量或单元素元组填充, 直接给 RGB 会 TypeError
        if im.mode in ('L', '1'):
            fill = int(round(0.299 * fill[0] + 0.587 * fill[1]
                             + 0.114 * fill[2]))
        elif im.mode == 'LA':
            # 灰度+透明 PNG: 必须是 (灰度, 透明度) 二元组, 给 RGB 元组直接崩
            fill = (int(round(0.299 * fill[0] + 0.587 * fill[1]
                              + 0.114 * fill[2])), 255)
        elif im.mode == 'RGBA':
            fill = tuple(fill) + (255,)
        elif im.mode == 'CMYK':
            fill = tuple(fill) + (0,)
        d.rectangle((x, y, x + w - 1, y + h - 1), fill=fill)
    out_dir = options.get('out_dir') or os.path.dirname(path)
    os.makedirs(out_dir, exist_ok=True)
    base = os.path.splitext(os.path.basename(path))[0]
    # 保持原格式: 早期恒存 PNG, 照片批处理后被批量改成 .png 且 EXIF 全丢
    # (竖拍照片 Orientation=6 一丢, 打开就是横的)。
    out = _uniq_out(os.path.abspath(
        os.path.join(out_dir, base + '_erase' + ext)))
    save_kw = {}
    exif = None
    try:
        exif = im.info.get('exif')
    except Exception:
        exif = None
    if ext in ('.jpg', '.jpeg'):
        if im.mode not in ('RGB', 'L'):
            im = im.convert('RGB')
        save_kw['quality'] = 95
        if exif:
            save_kw['exif'] = exif
    elif ext in ('.tif', '.tiff', '.webp'):
        if exif:
            save_kw['exif'] = exif
    im.save(out, **save_kw)
    return FileResult(path, 'ok', f'擦除 ({x},{y},{w},{h})', out)


def image_merge_files(paths: list, out_path: str, fmt: str = '长图'):
    """多图合并: 长图(纵向拼接) 或 PDF(多页)。返回输出路径。"""
    from PIL import Image, ImageOps

    imgs = []
    for p in paths:
        try:
            im = Image.open(p)
            try:
                im = ImageOps.exif_transpose(im)
            except Exception:
                pass
            if fmt == 'PDF' and im.mode != 'RGB':
                bg = Image.new('RGB', im.size, (255, 255, 255))
                if im.mode == 'RGBA':
                    bg.paste(im, mask=im.split()[3])
                else:
                    bg.paste(im)
                im = bg
            imgs.append(im)
        except Exception:
            continue
    if not imgs:
        return ''
    if fmt == 'PDF':
        first, rest = imgs[0], imgs[1:]
        first.save(out_path, 'PDF', save_all=True, append_images=rest,
                   resolution=100)
        return out_path
    # 长图
    W = max(im.width for im in imgs)
    total_h = sum(round(im.height * W / im.width) for im in imgs)
    canvas = Image.new('RGB', (W, total_h), (255, 255, 255))
    y = 0
    for im in imgs:
        nh = round(im.height * W / im.width)
        r = im.resize((W, nh))
        # 有 alpha 的图必须带 mask 粘: 直接 paste 会丢掉透明度, 全透明像素
        # 按原始 RGB(通常是 0,0,0) 贴上去, 结果下半张图变成纯黑。
        if r.mode in ('RGBA', 'LA'):
            r = r.convert('RGBA')
            canvas.paste(r, (0, y), r.split()[-1])
        else:
            canvas.paste(r.convert('RGB'), (0, y))
        y += nh
    canvas.save(out_path, 'PNG')
    return out_path


# ==================================================================== 文件
def files_merge_folders(paths: list, options: dict):
    """把多个源文件夹的内容合并到目标文件夹。

    返回 (移动文件数, 目标目录, 失败列表[(源文件, 原因)])。单个文件被占用时
    早期会直接抛 PermissionError 中断: 已搬的停在目标、未搬的留在源, 用户看到
    fail 但两边各一半, 重跑还会多出时间戳副本。
    """
    import shutil

    target = str(options.get('merge_target', '') or '')
    if not target:
        return 0, ''
    target = os.path.abspath(target)
    # 目标若位于某个源目录内部, os.walk 会边走边把内容搬进自己的子目录
    for src in paths:
        if not os.path.isdir(src):
            continue
        src_abs = os.path.abspath(src)
        if src_abs == target:
            continue
        # 判断 target 是否位于 src 内部: 把 target 相对于 src 求相对路径,
        # 结果不含 '..' 说明确实在其内部。relpath 不会像 commonpath 那样
        # 对跨盘符抛 ValueError。
        rel = os.path.relpath(target, src_abs)
        if rel != os.curdir and not rel.startswith('..' + os.sep) \
                and rel != '..' and not os.path.isabs(rel):
            raise ValueError('汇总目标目录不能位于源文件夹内部: ' + target)
    os.makedirs(target, exist_ok=True)
    n = 0
    fails = []
    for src in paths:
        if not os.path.isdir(src):
            continue
        src = os.path.abspath(src)
        if src == target:
            continue
        for dp, _dns, fns in os.walk(src):
            for fn in fns:
                fp = os.path.join(dp, fn)
                rel = os.path.relpath(fp, src)
                dst_dir = os.path.join(target, os.path.dirname(rel)) \
                    if options.get('keep_parent') else target
                os.makedirs(dst_dir, exist_ok=True)
                f = os.path.join(dst_dir, fn if not options.get('keep_parent')
                                 else os.path.basename(os.path.dirname(fp))
                                 + '_' + fn)
                f = os.path.abspath(f)
                if os.path.abspath(fp) == f:
                    continue
                if os.path.exists(f):
                    b, e = os.path.splitext(fn)
                    f = os.path.join(dst_dir,
                                     b + '_' + datetime.now().strftime('%H%M%S%f')
                                     + e)
                try:
                    shutil.move(fp, f)
                    n += 1
                except OSError as e:
                    # 被 Word 占用的文件: shutil.move 可能已复制到目标却删不掉源,
                    # 不清理就没重跑一次多一份副本。单个失败不该中断整批。
                    if os.path.exists(f) and os.path.exists(fp):
                        try:
                            os.remove(f)
                        except OSError:
                            pass
                    fails.append((fp, str(e)[:120]))
    return n, target, fails


def files_delete_folder(path: str, options: dict) -> FileResult:
    """把整个文件夹移入回收备份区。

    早期直接 shutil.move(跨盘=先 copytree 再 rmtree): 源目录里只要有只读文件,
    rmtree 就会在删掉一部分之后抛 PermissionError —— 用户看到 "fail", 源目录
    已经少了几个文件, 而清单还没写, 界面无法恢复。现在改为:
      1) 同盘直接原子改名(零风险);
      2) 跨盘先完整复制到回收区并落清单, 再尽力逐个删除, 剩下删不掉的
         明确列出来而不是让用户以为整个操作失败了。
    """
    if not path or not os.path.isdir(path):
        return FileResult(path, 'fail', '文件夹不存在或不是目录')
    trash = options.get('trash_dir') or default_trash_dir()
    trash = os.path.abspath(trash)
    if os.path.abspath(path) == trash:
        return FileResult(path, 'fail', '不能把回收备份区本身移入回收备份区')
    os.makedirs(trash, exist_ok=True)
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    name = os.path.basename(path.rstrip('\\/')) or 'folder'
    dst = os.path.join(trash, stamp + '_' + name)
    if os.path.abspath(dst) == os.path.abspath(path):
        return FileResult(path, 'fail', '源与目标相同, 已中止')
    from engine.store import TrashManifest
    undo = [{'src': path, 'dst': dst}]

    renamed = False
    try:
        os.rename(path, dst)
        renamed = True
    except OSError:
        renamed = False
    if renamed:
        TrashManifest().record(os.path.basename(dst), path)
        return FileResult(path, 'ok', '文件夹已移入回收备份区(可恢复)', dst,
                          undo=undo)

    # 跨盘/被占用: 先保证回收区有完整副本, 副本成功即落清单, 之后才动源
    try:
        shutil.copytree(path, dst, symlinks=True)
    except Exception as e:
        return FileResult(path, 'fail', '复制到回收备份区失败, 源未改动: '
                          + str(e)[:120])
    TrashManifest().record(os.path.basename(dst), path)

    left = []
    for dp, dns, fns in os.walk(path, topdown=False):
        for fn in fns:
            fp = os.path.join(dp, fn)
            try:
                os.remove(fp)
            except OSError:
                try:
                    os.chmod(fp, 0o666)  # 只读属性会挡住删除
                    os.remove(fp)
                except OSError:
                    left.append(fp)
        for dn in dns:
            try:
                os.rmdir(os.path.join(dp, dn))
            except OSError:
                pass
    try:
        os.rmdir(path)
    except OSError:
        left.append(path)
    if left:
        return FileResult(
            path, 'ok',
            f'已备份到回收区, 但 {len(left)} 项删不掉(可能被占用), 可手动删除',
            dst, undo=undo)
    return FileResult(path, 'ok', '文件夹已移入回收备份区(可恢复)', dst,
                      undo=undo)


def files_create_from_excel(rows: list, base_dir: str) -> int:
    """按 Excel 路径清单批量创建文件/文件夹。

    rows: [{"path": ..., "type": "文件|文件夹"}]

    中途命中越界路径会抛错: 早期前面几行已经落盘且不回滚, 用户看到 fail
    却发现桌面上多了半批文件。现在记录已创建项, 出错时按创建的逆序清理。
    """
    made = 0
    created = []   # [(路径, 是否原本就存在)] —— 只回滚本次真正新建的
    root = os.path.abspath(base_dir)
    try:
        for r in rows:
            p = str(r.get('path', '') or '').strip()
            if not p:
                continue
            t = str(r.get('type', '') or '').strip()
            p = os.path.abspath(p if os.path.isabs(p)
                                else os.path.join(root, p))
            # Excel 清单是外部输入, '..' 会写到 base_dir 之外, 必须拦在根目录内
            try:
                inside = os.path.commonpath([root, p]) == root
            except ValueError:
                inside = False
            if not inside:
                raise ValueError('路径越界(超出清单根目录): ' + str(r.get('path')))
            if t == '文件夹':
                existed = os.path.isdir(p)
                os.makedirs(p, exist_ok=True)
                if not existed:
                    created.append((p, True))
                made += 1
            else:
                parent = os.path.dirname(p) or '.'
                p_existed = os.path.dirname(p) and os.path.isdir(parent)
                os.makedirs(parent, exist_ok=True)
                if not p_existed:
                    created.append((parent, True))
                if not os.path.exists(p):
                    open(p, 'w', encoding='utf-8').close()
                    created.append((p, False))
                made += 1
    except Exception:
        for p, is_dir in reversed(created):
            try:
                if is_dir and os.path.isdir(p):
                    os.rmdir(p)      # 只删空目录, 用户原有内容不受影响
                elif not is_dir and os.path.exists(p):
                    os.remove(p)
            except OSError:
                pass
        raise
    return made


def files_copy_by_excel(rows: list, options: dict):
    """按 Excel 路径清单批量复制: [{"path":源, "target":目标}]

    返回 (成功数, 失败列表[(源, 原因)]) — GUI 按二元组解包, 早期只
    return int 会让该功能 100% 抛 TypeError, 而复制实际已执行。
    目标已存在时不覆盖: 用户手头的同名文件可能正被编辑, 静默覆盖等于
    丢数据。改名加 (1) 保存。
    """
    import shutil

    n = 0
    fails = []
    for r in rows:
        src = str(r.get('path', '') or '').strip()
        tgt = str(r.get('target', '') or '').strip()
        if not src or not tgt:
            continue
        if not os.path.exists(src):
            fails.append((src, '源文件不存在'))
            continue
        try:
            if os.path.exists(tgt):
                b, e = os.path.splitext(tgt)
                k = 1
                while os.path.exists(f'{b}({k}){e}'):
                    k += 1
                tgt = f'{b}({k}){e}'
            if os.path.isdir(src):
                shutil.copytree(src, tgt, dirs_exist_ok=True)
            else:
                os.makedirs(os.path.dirname(tgt) or '.', exist_ok=True)
                shutil.copy2(src, tgt)
            n += 1
        except OSError as e:
            fails.append((src, str(e)))
    return n, fails


# ============================================================== 正则生成器
def regexp_builder(kind: str, text: str, count: int, position: str,
                   position_ref: str) -> str:
    """按结构化条件生成正则(简化版可视化生成器)。

    kind: 单个任意文本/单个数字/单个汉字/单个英文字母/单个标点/邮箱/手机号/
          身份证/网址/日期/整数
    count: 重复次数描述 (0=任意次, n=恰好n次, -1=n次以上)
    position_ref: 某个文本的左边/右边/之间
    """
    kind_map = {
        '单个任意文本': '.', '单个数字': r'\d', '单个汉字': '[一-鿿]',
        '单个英文字母': '[A-Za-z]', '单个空白': r'\s',
        '单个标点符号': r'[^\w\s]', '中国大陆手机号': r'1[3-9]\d{9}',
        '邮箱': r'[\w.+-]+@[\w-]+\.[\w.]+',
        '中国大陆18位身份证号': r'\d{17}[\dXx]',
        '网址': r'https?://[^\s]+', '日期': r'\d{4}-\d{1,2}-\d{1,2}',
        '整数': r'-?\d+',
    }
    body = kind_map.get(kind)
    # 只有"单个X"这一族才接受重复次数。早期用 len(body) > 2 判断复合模式,
    # 把 [一-鿿]、[A-Za-z]、[^\w\s] 这些字符类全误判成复合模式, 勾了
    # "重复 5 次"却输出 [一-鿿] 而不是 [一-鿿]{5}。
    SINGLE_KINDS = {'单个任意文本', '单个数字', '单个汉字', '单个英文字母',
                    '单个空白', '单个标点符号'}
    if body is None:
        body = re.escape(text or '')
        composite = False
    else:
        composite = kind not in SINGLE_KINDS
    if composite:
        core = body
    else:
        if count == 0:
            core = body + '*'
        elif count == -1:
            core = body + '+'
        elif count == 1:
            core = body
        else:
            core = body + '{' + str(count) + '}'
    if position == '某个文本的左边位置' and position_ref:
        return re.escape(position_ref) + core
    if position == '某个文本的右边位置' and position_ref:
        return core + re.escape(position_ref)
    return core
