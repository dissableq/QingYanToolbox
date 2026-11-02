# -*- coding: utf-8 -*-
"""青雁批量工具箱 v3 — Office/PDF 批量处理器 (Word/Excel/PPT/PDF), 全离线。

依赖: python-docx / openpyxl / python-pptx / pypdf / reportlab / Pillow
每个函数: (path, options) -> FileResult

v3 增强:
- _out_or_inplace 支持撤销动作(就地修改: 备份可回退)。
- docx 替换修复跨 run 断词问题(合并 run 文本后再替换)。
- 模板系统与 v2 完全兼容(6 套内置办公模板)。
"""
from __future__ import annotations

from engine.safety import backup_file, atomic_save, replace_text_nodes, update_row_references

import os
import re
import shutil
import zipfile

from engine.core import FileResult


# ------------------------------------------------------------------ 工具
def _uniq_path(p: str) -> str:
    """输出文件名冲突时追加 (1)(2), 避免静默覆盖已有文件。"""
    if not os.path.exists(p):
        return p
    b, e = os.path.splitext(p)
    k = 1
    while os.path.exists(f'{b}({k}){e}'):
        k += 1
    return f'{b}({k}){e}'


_DOCM_CT = b'application/vnd.ms-word.document.macroEnabled.main+xml'
_DOCX_CT = (b'application/vnd.openxmlformats-officedocument.'
            b'wordprocessingml.document.main+xml')


def _open_docx_compat(path: str):
    """打开 .docx/.docm。返回 (Document, 临时补丁文件或 None)。

    python-docx 对宏文档(.docm)的内容类型直接抛 ValueError。docm 本质
    就是 docx 加 vbaProject, 把 [Content_Types].xml 里的宏类型改成普通
    类型即可用 python-docx 处理(实测打开/编辑/保存都正常)。补丁落在
    临时副本上, 原文件不动, 用完由调用方删除。
    """
    import docx

    if os.path.splitext(path)[1].lower() != '.docm':
        return docx.Document(path), None
    import tempfile
    fd, tmp = tempfile.mkstemp(suffix='.docx')
    os.close(fd)
    with zipfile.ZipFile(path) as zin, \
            zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == '[Content_Types].xml':
                data = data.replace(_DOCM_CT, _DOCX_CT)
            zout.writestr(item, data)
    return docx.Document(tmp), tmp


def _out_or_inplace(path: str, options: dict, tag: str):
    """决定输出路径: out_dir 指定则输出到输出目录; 否则就地(先备份)。

    返回 (输出路径, 撤销动作列表)。就地修改时自动生成 .qingyanbak 备份,
    并产生 undo 动作供台账页回退。

    备份必须每次就地修改都重新生成: 若沿用旧备份, 对同一文件连续处理两次时,
    第二次的撤销会退回到第一次操作之前, 而不是本次操作之前。
    """
    out_dir = options.get('out_dir')
    base, ext = os.path.splitext(os.path.basename(path))
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
        out = os.path.abspath(os.path.join(out_dir, base + tag + ext))
        return _uniq_path(out), []
    bak = backup_file(path)
    out = os.path.abspath(path)
    undo = [{'src': bak, 'dst': path, 'kind': 'restore'}]
    return out, undo


# ===================================================================== Word
def docx_replace(path: str, options: dict) -> FileResult:
    d, _tmp = _open_docx_compat(path)
    find = options.get('find', '')
    if not find:
        return FileResult(path, 'skipped', '未填写查找内容')
    replace = options.get('replace', '')
    use_re = bool(options.get('regex', False))
    if use_re:
        try:
            re.compile(find)
        except re.error as e:
            return FileResult(path, 'fail', '正则表达式无效: ' + str(e))
    n = 0

    def _sub_in_paragraph(p) -> None:
        nonlocal n
        # Limit to this paragraph: nested textboxes are visited separately.
        nodes = [node for node in p._p.iter(f'{{{_W_NS}}}t')
                 if next((ancestor for ancestor in node.iterancestors()
                          if ancestor.tag == f'{{{_W_NS}}}p'), None) is p._p]
        n += replace_text_nodes(nodes, find, replace, use_re)

    # 遍历"全部故事部分": 正文/表格(含嵌套)/文本框走 document XML 的
    # w:p 全量迭代; 页眉/页脚(含首页不同/奇偶页)按节单独取。
    # 只扫 d.paragraphs + 顶层表格会漏掉页眉页脚与文本框里的文字,
    # 用户会看到"未匹配到内容"而内容明明存在。
    from docx.text.paragraph import Paragraph

    for part_story in [d.element] + [
            sec_part for sec in d.sections
            for sec_part in (sec.header._element, sec.footer._element,
                             sec.first_page_header._element,
                             sec.first_page_footer._element,
                             sec.even_page_header._element,
                             sec.even_page_footer._element)]:
        for el in part_story.iter(f'{{{_W_NS}}}p'):
            _sub_in_paragraph(Paragraph(el, None))

    if n == 0:
        return FileResult(path, 'skipped', '未匹配到内容')
    out, undo = _out_or_inplace(path, options, '_replaced')
    out = _docx_finish(d, _tmp, path, out)
    return FileResult(path, 'ok', f'替换 {n} 处', out, undo=undo)


_W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
_A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'


def _docx_finish(d, tmp, path: str, out: str) -> str:
    """docx 系列保存收尾: docm 补丁副本的扩展名处理 + 临时文件清理。"""
    if tmp:
        if os.path.splitext(out)[1].lower() == '.docm':
            # 经内容类型补丁的副本存出来是普通 docx, 写回 .docm 名会被
            # Word 报"内容与扩展名不符"; 落地为 .docx
            cand = os.path.splitext(out)[0] + '.docx'
            if os.path.abspath(cand) != os.path.abspath(path):
                cand = _uniq_path(cand)
            out = cand
    atomic_save(d.save, out)
    if tmp:
        try:
            os.remove(tmp)
        except OSError:
            pass
    return out


def docx_clear_blank(path: str, options: dict) -> FileResult:
    d, tmp = _open_docx_compat(path)
    n = 0
    for p in list(d.paragraphs):
        if p.text.strip() or p._element.findall(f'.//{{{_A_NS}}}blip'):
            continue
        # 承载 sectPr 的段落文字为空, 但删掉会丢失分节符与分页设置
        if p._element.find(f'./{{{_W_NS}}}pPr/{{{_W_NS}}}sectPr') is not None:
            continue
        p._element.getparent().remove(p._element)
        n += 1
    for t in d.tables:
        for row in t.rows:
            for cell in row.cells:
                paras = list(cell.paragraphs)
                for p in paras:
                    if p.text.strip() or p._element.findall(f'.//{{{_A_NS}}}blip'):
                        continue
                    # w:tc 按 ECMA-376 必须至少含一个块级元素,
                    # 删光段落会产出 Word 报"需要修复"的非法结构
                    if len(paras) <= 1:
                        continue
                    p._element.getparent().remove(p._element)
                    n += 1
    if n == 0:
        return FileResult(path, 'skipped', '无空白段落')
    out, undo = _out_or_inplace(path, options, '_cleaned')
    out = _docx_finish(d, tmp, path, out)
    return FileResult(path, 'ok', f'清理空段落 {n} 处', out, undo=undo)


def _prune_unused_media(path):
    """Prune each part's unused image relationships before removing its image bytes."""
    import posixpath
    from lxml import etree
    relationship_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    with zipfile.ZipFile(path) as z:
        content = {name: z.read(name) for name in z.namelist()}
    kept = set()
    for name, data in list(content.items()):
        if not name.endswith('.rels'): continue
        part_dir = posixpath.dirname(posixpath.dirname(name))
        owner = posixpath.join(part_dir, posixpath.basename(name)[:-5])
        used = set()
        if owner in content:
            owner_xml = etree.fromstring(content[owner])
            used = {value for el in owner_xml.iter() for key,value in el.attrib.items()
                    if key.startswith('{'+relationship_ns+'}')}
        rels = etree.fromstring(data)
        for rel in list(rels):
            if rel.get('TargetMode') == 'External': continue
            target = rel.get('Target','')
            resolved = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(part_dir,target))
            is_image = rel.get('Type','').endswith('/image')
            if is_image and owner in content and rel.get('Id') not in used:
                rels.remove(rel)
            else:
                kept.add(resolved)
        content[name] = etree.tostring(rels,xml_declaration=True,encoding='UTF-8',standalone=True)
    removed = [name for name in content if name.startswith('word/media/') and name not in kept]
    for name in removed: del content[name]
    def save(destination):
        with zipfile.ZipFile(destination,'w',zipfile.ZIP_DEFLATED) as z:
            for name,data in content.items(): z.writestr(name,data)
    atomic_save(save,path)
    return len(removed)


def docx_remove_images(path: str, options: dict) -> FileResult:
    d, tmp = _open_docx_compat(path)
    from docx.oxml.ns import qn

    n = 0
    # 只摘 wp:inline 会残留空的 <w:drawing/>, 违反 schema(Word 可能报修复)。
    # 正确做法是连同所在的 <w:r> 一起删除。
    # d.inline_shapes 只含内嵌图: 浮动(环绕)图 wp:anchor 和页眉/页脚里
    # 的图都删不到, 还会误报"无图片"。改按 XML 遍历全部故事部分里的
    # w:drawing(inline 与 anchor 都覆盖)。
    stories = [d.element] + [
        sp for sec in d.sections
        for sp in (sec.header._element, sec.footer._element,
                   sec.first_page_header._element,
                   sec.first_page_footer._element,
                   sec.even_page_header._element,
                   sec.even_page_footer._element)]
    for story in stories:
        for drawing in list(story.iter(qn('w:drawing'))):
            # 该 drawing 必须真的承载图片(pic:pic), 排除图表等非图片
            if drawing.find(f'.//{{{_A_NS}}}blip') is None:
                continue
            drawing.getparent().remove(drawing)
            n += 1
    if n == 0:
        return FileResult(path, 'skipped', '无图片')
    out, undo = _out_or_inplace(path, options, '_noimg')
    out = _docx_finish(d, tmp, path, out)
    # 清掉包里已无引用的图片字节, 否则文件体积一点不变
    _prune_unused_media(out)
    return FileResult(path, 'ok', f'删除图片 {n} 张', out, undo=undo)


def docx_watermark(path: str, options: dict) -> FileResult:
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt, RGBColor

    d, tmp = _open_docx_compat(path)
    text = options.get('wm_text') or '青雁水印'

    def _mark(hf) -> None:
        p = hf.add_paragraph(text)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.font.color.rgb = RGBColor(0xD9, 0xD9, 0xD9)
            r.font.size = Pt(24)

    for sec in d.sections:
        # "首页不同"的文档首页页眉是独立 part, 只写普通页眉会让第 1 页
        # 没有水印而任务仍报成功(公文/简历常态)。
        hfs = [sec.header]
        try:
            if sec.different_first_page_header_footer:
                hfs.append(sec.first_page_header)
        except Exception:
            pass
        for hf in hfs:
            # 未断链时多节共用同一个 header part, 逐节 add_paragraph
            # 会重复追加 N 遍
            if hf.is_linked_to_previous:
                hf.is_linked_to_previous = False
            _mark(hf)
    out, undo = _out_or_inplace(path, options, '_wm')
    out = _docx_finish(d, tmp, path, out)
    return FileResult(path, 'ok', f"页眉水印'{text}'", out, undo=undo)


def docx_header_footer(path: str, options: dict) -> FileResult:
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    d, tmp = _open_docx_compat(path)
    header = options.get('header_text', '')
    footer = options.get('footer_text', '')

    def _put(hf, text) -> None:
        p = hf.add_paragraph(text)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for sec in d.sections:
        # 页眉/页脚允许只填一个: 两个都必填会把只想加页眉的用户拦住
        parts = []
        if header:
            parts.append((sec.header, header))
            try:
                if sec.different_first_page_header_footer:
                    parts.append((sec.first_page_header, header))
            except Exception:
                pass
        if footer:
            parts.append((sec.footer, footer))
            try:
                if sec.different_first_page_header_footer:
                    parts.append((sec.first_page_footer, footer))
            except Exception:
                pass
        for hf, text in parts:
            # 多节默认共用同一个 header/footer part, 必须先断链,
            # 否则内容重复追加
            if hf.is_linked_to_previous:
                hf.is_linked_to_previous = False
            _put(hf, text)
    out, undo = _out_or_inplace(path, options, '_hf')
    out = _docx_finish(d, tmp, path, out)
    return FileResult(path, 'ok', '页眉页脚已设置', out, undo=undo)


def _repack_zip(src: str, dst: str) -> None:
    """把 OOXML 包重新打包为最高压缩率(zip 层面的真实压缩)。"""
    tmp = dst + '.tmpzip'
    try:
        with zipfile.ZipFile(src) as zin, \
                zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED,
                                compresslevel=9) as zout:
            for item in zin.namelist():
                zout.writestr(item, zin.read(item))
        shutil.move(tmp, dst)
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass


def _size_detail(so: int, sn: int) -> str:
    if so <= 0:
        return f'{sn} 字节'
    pct = (so - sn) / so * 100
    sign = '减少' if pct >= 0 else '增加'
    return f'{so} → {sn} 字节({sign} {abs(pct):.1f}%)'


def docx_compress(path: str, options: dict) -> FileResult:
    # 原大小必须在覆盖之前取, 否则就地模式下永远显示 X → X
    so = os.path.getsize(path)
    out, undo = _out_or_inplace(path, options, '_c')
    _repack_zip(path, out)
    return FileResult(path, 'ok',
                      _size_detail(so, os.path.getsize(out)), out, undo=undo)


# ==================================================================== Excel
def _load_xlsx(path: str, **kw):
    """openpyxl 载入: .xlsm 必须 keep_vba=True, 否则保存后 VBA 宏丢失
    (输出还是 .xlsm 名, 用户以为宏还在, 实际 vbaProject.bin 没了)。"""
    import openpyxl

    if os.path.splitext(path)[1].lower() == '.xlsm':
        kw.setdefault('keep_vba', True)
    return openpyxl.load_workbook(path, **kw)


def xlsx_replace(path: str, options: dict) -> FileResult:
    wb = _load_xlsx(path)
    find = options.get('find', '')
    if not find:
        return FileResult(path, 'skipped', '未填写查找内容')
    replace = options.get('replace', '')
    use_re = bool(options.get('regex', False))
    # 非法正则(a[b)不能让 PatternError 原样抛进结果表, 与 docx_replace 保持一致
    if use_re:
        try:
            re.compile(find)
        except re.error as e:
            return FileResult(path, 'fail', '正则表达式无效: ' + str(e))
    else:
        pat = None
    n = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if not isinstance(c.value, str):
                    continue
                if use_re:
                    if pat.search(c.value):
                        c.value = pat.sub(replace, c.value)
                        n += 1
                elif find in c.value:
                    c.value = c.value.replace(find, replace)
                    n += 1
    if n == 0:
        return FileResult(path, 'skipped', '未匹配到内容')
    out, undo = _out_or_inplace(path, options, '_replaced')
    atomic_save(wb.save, out)
    return FileResult(path, 'ok', f'替换单元格 {n} 个', out, undo=undo)


def _real_extent(ws):
    """Sheet 的真实数据范围 (max_row, max_column)。

    不能用 ws.max_row/max_column: Excel 的尾行常常虚高(残留格式/Ctrl+End),
    可达 1048576 行, 逐格 ws.cell() 遍历会卡死整批任务; 且 ws.cell() 是写操作,
    访问即凭空创建空单元格, 撑大工作簿并流入后续循环。
    """
    max_r = max_c = 0
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None:
                if c.row > max_r:
                    max_r = c.row
                if c.column > max_c:
                    max_c = c.column
    return max_r, max_c


def xlsx_delete_rows(path: str, options: dict) -> FileResult:
    mode = options.get('row_mode', 'blank')
    wb = _load_xlsx(path)
    n = 0
    deletions = {}
    for ws in wb.worksheets:
        max_r, max_c = _real_extent(ws)
        if max_r == 0:
            continue
        vals_of = {}
        for idx in range(1, max_r + 1):
            vals_of[idx] = [ws.cell(row=idx, column=c).value
                            for c in range(1, max_c + 1)]
        if mode == 'blank':
            doomed = [idx for idx, vals in vals_of.items()
                      if all(v is None or str(v).strip() == '' for v in vals)]
        else:  # duplicate
            seen = set()
            doomed = []
            for idx, vals in vals_of.items():
                # 全空行不参与去重, 否则所有空行互为"重复"被批量删光
                if all(v is None or str(v).strip() == '' for v in vals):
                    continue
                key = tuple((type(v).__name__, str(v)) for v in vals)
                if key in seen:
                    doomed.append(idx)
                else:
                    seen.add(key)
        # openpyxl 的 delete_rows 不搬移合并单元格、公式引用也不调整,
        # 删行后合并框钉在原地、公式全部串行。删之前记录合并区, 删完
        # 按被删行数整体上移重建。
        def _shift_merges(ws, doomed_rows: list):
            from openpyxl.utils import range_boundaries
            doomed = sorted(doomed_rows)
            old = list(ws.merged_cells.ranges)
            if not old:
                return
            # 不能用 unmerge_cells: delete_rows 已把单元格上移, 它按旧坐标
            # 清 MergedCell 会 KeyError。直接摘 ranges 再重建。
            for m in old:
                ws.merged_cells.remove(m)
            import bisect
            for m in old:
                c1, r1, c2, r2 = range_boundaries(str(m))
                nr1 = r1 - bisect.bisect_left(doomed, r1)
                nr2 = r2 - bisect.bisect_left(doomed, r2 + 1)
                if nr2 < nr1:
                    continue        # 整块都在被删行里
                ws.merge_cells(start_row=nr1, start_column=c1,
                               end_row=nr2, end_column=c2)

        deletions[ws.title] = sorted(doomed)
        for idx in reversed(doomed):
            ws.delete_rows(idx)
            n += 1
        _shift_merges(ws, doomed)
    if n == 0:
        return FileResult(path, 'skipped', '没有符合条件的行')
    update_row_references(wb, deletions)
    out, undo = _out_or_inplace(path, options, '_rows')
    atomic_save(wb.save, out)
    return FileResult(path, 'ok', f'删除 {n} 行', out, undo=undo)


def xlsx_formulas_to_values(path: str, options: dict) -> FileResult:
    """公式转值。

    data_only=True 读到的是 Excel 上次保存时缓存的计算结果。任何非 Excel 程序
    生成的文件(openpyxl / pandas / 程序导出)都没有缓存值, 此时 cv.value 为 None;
    若无条件写回会把公式清空, 造成静默数据丢失。故无缓存值时保留原公式。
    """
    wb = _load_xlsx(path)
    wbv = _load_xlsx(path, data_only=True)
    n = 0
    no_cache = 0
    try:
        for ws, wsv in zip(wb.worksheets, wbv.worksheets):
            for row, rowv in zip(ws.iter_rows(), wsv.iter_rows()):
                for c, cv in zip(row, rowv):
                    if isinstance(c.value, str) and c.value.startswith('='):
                        if cv.value is None:
                            no_cache += 1
                            continue      # 无缓存值: 保留公式
                        c.value = cv.value
                        n += 1
    finally:
        try:
            wbv.close()
        except Exception:
            pass
    if n == 0:
        if no_cache:
            return FileResult(path, 'skipped',
                              f'{no_cache} 个公式没有缓存计算结果'
                              '(需先用 Excel 打开并保存一次), 已保持原样')
        return FileResult(path, 'skipped', '未发现公式')
    out, undo = _out_or_inplace(path, options, '_vals')
    atomic_save(wb.save, out)
    detail = f'公式转值 {n} 个'
    if no_cache:
        detail += f'; {no_cache} 个无缓存值已保留公式'
    return FileResult(path, 'ok', detail, out, undo=undo)


def xlsx_watermark(path: str, options: dict) -> FileResult:
    wb = _load_xlsx(path)
    text = options.get('wm_text') or '青雁水印'
    for ws in wb.worksheets:
        ws.oddHeader.center.text = text
        ws.oddHeader.center.size = 36
        ws.oddHeader.center.color = 'D9D9D9'
    out, undo = _out_or_inplace(path, options, '_wm')
    atomic_save(wb.save, out)
    return FileResult(path, 'ok', f'{len(wb.worksheets)} 个 Sheet 加页眉水印',
                      out, undo=undo)


def xlsx_sheet_ops(path: str, options: dict) -> FileResult:
    op = options.get('sheet_op', 'rename')
    wb = _load_xlsx(path)
    n = 0
    if op == 'rename':
        find = options.get('find', '')
        replace = options.get('replace', '')
        for ws in wb.worksheets:
            if find and find in ws.title:
                ws.title = ws.title.replace(find, replace)
                n += 1
    elif op == 'delete_blank':
        for name in list(wb.sheetnames):
            ws = wb[name]
            v = ws['A1'].value
            if (v is None or str(v).strip() == '') and ws.max_row <= 1 and ws.max_column <= 1:
                wb.remove(ws)
                n += 1
    elif op == 'sort':
        names = sorted(wb.sheetnames)
        for i, nm in enumerate(names):
            idx = wb.sheetnames.index(nm)
            wb.move_sheet(nm, offset=i - idx)
            n += 1
    if n == 0:
        return FileResult(path, 'skipped', '没有符合条件的 Sheet')
    out, undo = _out_or_inplace(path, options, '_sheet')
    atomic_save(wb.save, out)
    return FileResult(path, 'ok', f'Sheet 操作 {n} 处', out, undo=undo)


def xlsx_compress(path: str, options: dict) -> FileResult:
    # load_workbook + save 不做任何压缩优化(实测字节数完全不变),
    # 这里直接在 zip 层面重打包, 压缩才真实生效。
    so = os.path.getsize(path)
    out, undo = _out_or_inplace(path, options, '_c')
    _repack_zip(path, out)
    return FileResult(path, 'ok',
                      _size_detail(so, os.path.getsize(out)), out, undo=undo)


# ====================================================================== PPT
def pptx_replace(path: str, options: dict) -> FileResult:
    from pptx import Presentation

    prs = Presentation(path)
    find = options.get('find', '')
    if not find:
        return FileResult(path, 'skipped', '未填写查找内容')
    replace = options.get('replace', '')
    n = 0

    def _sub_tf(tf) -> None:
        nonlocal n
        for p in tf.paragraphs:
            n += replace_text_nodes(p._p.iter(f'{{{_A_NS}}}t'), find, replace)

    def _walk(shapes) -> None:
        for shape in shapes:
            if shape.shape_type == 6:            # GROUP: 递归组内形状
                _walk(shape.shapes)
                continue
            if getattr(shape, 'has_table', False) and shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        _sub_tf(cell.text_frame)
                continue
            if shape.has_text_frame:
                _sub_tf(shape.text_frame)

    for slide in prs.slides:
        _walk(slide.shapes)
        # 演讲者备注也是常放内容的地方
        if slide.has_notes_slide:
            _sub_tf(slide.notes_slide.notes_text_frame)
    if n == 0:
        return FileResult(path, 'skipped', '未匹配到内容')
    out, undo = _out_or_inplace(path, options, '_replaced')
    atomic_save(prs.save, out)
    return FileResult(path, 'ok', f'替换 {n} 处', out, undo=undo)


def _slide_is_blank(slide) -> bool:
    """整页是否真的空白。

    只看 text_frame 会把"纯图片页 / 表格页 / 图表页"误判为空白并删掉,
    而 UI 明确承诺"纯图片页不受影响"。这里要求页面上除占位符外没有任何形状。
    """
    for shape in slide.shapes:
        # shape_type 为 None 或 PLACEHOLDER(14) 的多为版式占位符, 本身不算内容
        st = None
        try:
            st = int(shape.shape_type)
        except (TypeError, ValueError):
            st = None
        if st is not None and st != 14:
            return False
        if shape.has_text_frame and shape.text_frame.text.strip():
            return False
        if getattr(shape, 'has_table', False) and shape.has_table:
            return False
        if getattr(shape, 'has_chart', False) and shape.has_chart:
            return False
    return True


def pptx_delete_blank_slides(path: str, options: dict) -> FileResult:
    from pptx import Presentation

    prs = Presentation(path)
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    total = len(slides)
    n = 0
    for i, slide in enumerate(prs.slides):
        if _slide_is_blank(slide):
            xml_slides.remove(slides[i])
            n += 1
    if n == 0:
        return FileResult(path, 'skipped', '无空白页')
    if n >= total:
        return FileResult(path, 'fail',
                          '所有页面均为空白, 已中止(会产生打不开的空演示文稿)')
    out, undo = _out_or_inplace(path, options, '_cleaned')
    atomic_save(prs.save, out)
    return FileResult(path, 'ok', f'删除空白页 {n} 页', out, undo=undo)


def pptx_watermark(path: str, options: dict) -> FileResult:
    from pptx import Presentation
    from pptx.util import Emu, Pt

    prs = Presentation(path)
    text = options.get('wm_text') or '青雁水印'
    for slide in prs.slides:
        # 文本框宽度必须小于页宽, 否则右侧溢出幻灯片边界
        box = slide.shapes.add_textbox(
            Emu(457200), Emu(914400),
            Emu(max(1, prs.slide_width - 914400)),
            Emu(max(1, prs.slide_height - 1828800)))
        tf = box.text_frame
        p = tf.paragraphs[0]
        p.text = text
        for r in p.runs:
            r.font.size = Pt(12)
    out, undo = _out_or_inplace(path, options, '_wm')
    atomic_save(prs.save, out)
    return FileResult(path, 'ok', f'{len(prs.slides._sldIdLst)} 页加水印',
                      out, undo=undo)


def pptx_delete_pages(path: str, options: dict) -> FileResult:
    from pptx import Presentation

    prs = Presentation(path)
    try:
        start = int(options.get('page_start', 1) or 1)
        end = int(options.get('page_end', -1) or -1)
    except (TypeError, ValueError):
        return FileResult(path, 'fail', '页码需为整数')
    xml_slides = prs.slides._sldIdLst
    ids = list(xml_slides)
    total = len(ids)
    if end < 0:
        end = total
    start = max(1, start)
    end = min(total, end)
    n = 0
    for i in range(end - 1, start - 2, -1):
        if 0 <= i < total:
            xml_slides.remove(ids[i])
            n += 1
    if n == 0:
        return FileResult(path, 'skipped', '页码范围无效')
    if n >= total:
        return FileResult(path, 'fail',
                          '该范围会删除全部页面, 已中止(会产生打不开的演示文稿)')
    out, undo = _out_or_inplace(path, options, '_pages')
    atomic_save(prs.save, out)
    return FileResult(path, 'ok', f'删除第 {start}-{end} 页({n} 页)', out, undo=undo)


# ====================================================================== PDF
def _pdf_reader(path: str, options: dict):
    """打开 PDF; 已加密时用 options['password'] 解锁。

    pypdf 的 decrypt() 口令错误时 **不抛异常**, 只返回 PasswordType.NOT_DECRYPTED,
    且 is_encrypted 仍为 True。若不显式判定, 异常会推迟到访问 .pages 时以
    FileNotDecryptedError 逃逸(该异常不是 RuntimeError 子类, 调用方 catch 不到)。
    """
    from pypdf import PdfReader
    reader = PdfReader(path)
    if reader.is_encrypted:
        pwd = options.get('password') or ''
        try:
            res = reader.decrypt(pwd)
        except Exception as e:
            raise RuntimeError('PDF 解析失败: ' + str(e)) from e
        # 只认 decrypt() 的返回值: 口令正确时 is_encrypted 仍可能为 True
        # (owner password 场景), 拿它做判定会把正确口令误报成"口令不对"。
        if not res:
            raise RuntimeError('PDF 已加密且口令不对')
    return reader


def _pdf_was_encrypted(path: str) -> bool:
    """源文件本身是否加密(在 _pdf_reader 解密之前问)。"""
    from pypdf import PdfReader
    try:
        return bool(PdfReader(path).is_encrypted)
    except Exception:
        return False


def pdf_ops(path: str, options: dict) -> FileResult:
    from pypdf import PdfWriter

    op = options.get('pdf_op', 'merge_split')
    was_enc = _pdf_was_encrypted(path)
    try:
        reader = _pdf_reader(path, options)
    except RuntimeError as e:
        return FileResult(path, 'fail', str(e))
    except Exception as e:
        # 截断/垃圾 PDF 抛的是 pypdf 的 PdfStreamError 等(非 RuntimeError),
        # 早期会漏出去让用户看到英文堆栈
        return FileResult(path, 'fail',
                          'PDF 无法读取(文件损坏或不是 PDF): '
                          + type(e).__name__)
    out_dir = options.get('out_dir') or os.path.dirname(path)
    os.makedirs(out_dir, exist_ok=True)
    base, ext = os.path.splitext(os.path.basename(path))
    total = len(reader.pages)
    if total == 0:
        return FileResult(path, 'skipped', 'PDF 没有页面')

    if op == 'extract':
        try:
            a = max(1, int(options.get('page_from', 1) or 1))
            b = int(options.get('page_to', 0) or 0)
        except (TypeError, ValueError):
            return FileResult(path, 'fail', '页码必须是整数')
        b = total if b <= 0 else min(total, b)     # 0/留空 = 到末尾
        if a > b:
            return FileResult(path, 'fail',
                              f'起始页 {a} 大于结束页 {b}, 无内容可提取')
        w = PdfWriter()
        for i in range(a - 1, b):
            w.add_page(reader.pages[i])
        out = _uniq_path(os.path.abspath(
            os.path.join(out_dir, f'{base}_p{a}-{b}{ext}')))
        with open(out, 'wb') as fh:
            w.write(fh)
        return FileResult(path, 'ok', f'提取第 {a}-{b} 页', out)

    if op == 'split':
        try:
            chunk = int(options.get('chunk_size', 1))
        except (TypeError, ValueError):
            return FileResult(path, 'fail', '每份页数必须是整数')
        if chunk < 1:
            return FileResult(path, 'fail', '每份页数必须 ≥ 1')
        outs = []
        j = 0
        while j < total:
            w = PdfWriter()
            for i in range(j, min(j + chunk, total)):
                w.add_page(reader.pages[i])
            out = _uniq_path(os.path.abspath(os.path.join(
                out_dir, f'{base}_part{j // chunk + 1:03d}{ext}')))
            with open(out, 'wb') as fh:
                w.write(fh)
            outs.append(out)
            j += chunk
        return FileResult(path, 'ok', f'拆为 {len(outs)} 个文件',
                          outs[0] if outs else '', outputs=outs)

    if op == 'rotate':
        try:
            angle = int(options.get('angle', 90))
        except (TypeError, ValueError):
            return FileResult(path, 'fail', '旋转角度必须是整数')
        if angle % 90:
            return FileResult(path, 'fail', '旋转角度必须是 90 的倍数')
        angle %= 360
        w = PdfWriter()
        for pg in reader.pages:
            pg.rotate(angle)
            w.add_page(pg)
        out = _uniq_path(os.path.abspath(
            os.path.join(out_dir, base + '_rotated.pdf')))
        with open(out, 'wb') as fh:
            w.write(fh)
        return FileResult(path, 'ok', f'旋转 {angle} 度', out)

    if op == 'encrypt':
        pwd = str(options.get('password') or '').strip()
        if not pwd:
            return FileResult(path, 'fail', '未设置新口令')
        # 逐页复制而非 w.append(reader): append 会把源文件的 /Encrypt 字典
        # 带进 writer trailer, 再叠加 encrypt() 后口令状态不一致
        out, undo = _out_or_inplace(path, options, '_encrypted')
        w = PdfWriter()
        for pg in reader.pages:
            w.add_page(pg)
        w.encrypt(pwd)
        with open(out, 'wb') as fh:
            w.write(fh)
        return FileResult(path, 'ok', '已加密', out, undo=undo)

    if op == 'decrypt':
        # 判定源文件是否加密: _pdf_reader 已经解密过, 此时 is_encrypted 恒为 False
        if not was_enc:
            return FileResult(path, 'skipped', '未加密')
        out, undo = _out_or_inplace(path, options, '_decrypted')
        w = PdfWriter()
        for pg in reader.pages:
            w.add_page(pg)
        with open(out, 'wb') as fh:
            w.write(fh)
        return FileResult(path, 'ok', '已解密', out, undo=undo)

    if op == 'insert_blank':
        try:
            pos = max(1, int(options.get('insert_at', 1) or 1))
        except (TypeError, ValueError):
            return FileResult(path, 'fail', '插入位置必须是整数')
        if pos > total + 1:
            return FileResult(path, 'fail',
                              f'插入位置 {pos} 超出页数范围 1-{total + 1}')
        w = PdfWriter()
        from pypdf import PdfReader as _R
        from io import BytesIO
        from reportlab.pdfgen import canvas as rl_canvas
        buf = BytesIO()
        cv = rl_canvas.Canvas(buf, pagesize=(595.28, 841.89))
        cv.showPage()
        cv.save()
        buf.seek(0)
        blank = _R(buf).pages[0]
        for i, pg in enumerate(reader.pages, 1):
            if i == pos:
                w.add_page(blank)
            w.add_page(pg)
        if pos == total + 1:      # 末尾追加
            w.add_page(blank)
        out = _uniq_path(os.path.abspath(
            os.path.join(out_dir, base + '_blank.pdf')))
        with open(out, 'wb') as fh:
            w.write(fh)
        return FileResult(path, 'ok', f'第{pos}页前插入空白页', out)

    return FileResult(path, 'fail', '未知 PDF 操作 ' + str(op))


def pdf_merge_files(paths: list, out_path: str, password: str = '') -> int:
    """按顺序合并 PDF, 返回总页数。"""
    from pypdf import PdfReader, PdfWriter

    w = PdfWriter()
    total = 0
    try:
        for p in paths:
            r = PdfReader(p)
            try:
                if r.is_encrypted:
                    # decrypt 口令错误不抛异常, 只返回 NOT_DECRYPTED, 必须判定
                    if not r.decrypt(password or ''):
                        raise RuntimeError(f'PDF 已加密且口令不对: {p}')
                for pg in r.pages:
                    w.add_page(pg)
                    total += 1
            finally:
                try:
                    r.close()
                except Exception:
                    pass
        if total == 0:
            raise RuntimeError('没有可合并的页面')
        with open(out_path, 'wb') as fh:
            w.write(fh)
    finally:
        try:
            w.close()
        except Exception:
            pass
    return total


def pdf_header_footer(path: str, options: dict) -> FileResult:
    """reportlab 生成覆盖层, pypdf 合并。"""
    import io
    from pypdf import PdfReader, PdfWriter
    from reportlab.pdfgen import canvas as rl_canvas

    try:
        reader = _pdf_reader(path, options)
    except RuntimeError as e:
        return FileResult(path, 'fail', str(e))
    except Exception as e:
        # 截断/垃圾 PDF 抛的是 pypdf 的 PdfStreamError 等(非 RuntimeError),
        # 早期会漏出去让用户看到英文堆栈
        return FileResult(path, 'fail',
                          'PDF 无法读取(文件损坏或不是 PDF): '
                          + type(e).__name__)
    header = options.get('header_text', '')
    footer = options.get('footer_text', '')
    if not header and not footer:
        return FileResult(path, 'skipped', '未填写页眉/页脚内容')
    w = PdfWriter()
    for i, pg in enumerate(reader.pages):
        pw = float(pg.mediabox.width)
        ph = float(pg.mediabox.height)
        buf = io.BytesIO()
        c = rl_canvas.Canvas(buf, pagesize=(pw, ph))
        if header:
            c.setFont(_pdf_cjk_font(header), 9)
            c.drawCentredString(pw / 2, ph - 24, header)
        if footer:
            c.setFont(_pdf_cjk_font(footer), 9)
            c.drawCentredString(pw / 2, 16, footer)
        c.save()
        buf.seek(0)
        overlay = PdfReader(buf).pages[0]
        pg.merge_page(overlay)
        w.add_page(pg)
    out_dir = options.get('out_dir') or os.path.dirname(path)
    os.makedirs(out_dir, exist_ok=True)
    base, ext = os.path.splitext(os.path.basename(path))
    out = os.path.abspath(os.path.join(out_dir, base + '_hf.pdf'))
    # 必须防覆盖: 同批选了 合同.pdf 和 合同_hf.pdf 时, 直接写会把后者
    # (恰好是另一个待处理源文件)覆盖掉
    out = _uniq_path(out)
    with open(out, 'wb') as fh:
        w.write(fh)
    return FileResult(path, 'ok', '页眉页脚已添加', out)


_CJK_TTF_CANDIDATES = (
    r'C:\Windows\Fonts\msyh.ttc',      # 微软雅黑(嵌入子集, 任何阅读器可显示)
    r'C:\Windows\Fonts\simhei.ttf',    # 黑体
    r'C:\Windows\Fonts\simsun.ttc',    # 宋体
    r'C:\Windows\Fonts\Deng.ttf',      # 等线
)

_CJK_FONT_CACHE = {'name': None, 'resolved': False}


def TTFontSafe(name: str, path: str):
    """reportlab TTFont, 兼容 ttc(取第一个子字体); 失败由调用方兜底。"""
    from reportlab.pdfbase.ttfonts import TTFont
    try:
        return TTFont(name, path, subfontIndex=0)
    except TypeError:
        return TTFont(name, path)


def _pdf_cjk_font(text: str):
    """为 PDF 覆盖层挑一个能渲染 text 的字体。

    reportlab 内置的 Helvetica 等标准 14 字体没有 CJK 字形, 中文会被渲染成
    ■■。回退顺序:
    1. 系统中文字体 TTF/TTC — 字体子集直接嵌入 PDF, 任何阅读器可显示;
    2. reportlab 自带 STSong-Light CID 字体(不嵌入, 依赖阅读器亚洲字体包);
    3. 纯 ASCII 内容直接用 Helvetica。
    全部本地资源, 保持完全离线。
    """
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont

    def _has_cjk(s: str) -> bool:
        return any('\u4e00' <= ch <= '\u9fff' or
                   '\u3000' <= ch <= '\u303f' or
                   '\uff00' <= ch <= '\uffef' for ch in s)

    if not _has_cjk(text):
        return 'Helvetica'
    if _CJK_FONT_CACHE['resolved']:
        return _CJK_FONT_CACHE['name'] or 'STSong-Light'
    name = None
    # 1) 系统字体嵌入(最优)
    for fp in _CJK_TTF_CANDIDATES:
        try:
            if not os.path.exists(fp):
                continue
            cand = 'QYCJK'
            if cand not in pdfmetrics.getRegisteredFontNames():
                pdfmetrics.registerFont(TTFontSafe(cand, fp))
            name = cand
            break
        except Exception:
            continue
    # 2) CID 兜底
    if name is None:
        try:
            cand = 'STSong-Light'
            if cand not in pdfmetrics.getRegisteredFontNames():
                pdfmetrics.registerFont(UnicodeCIDFont(cand))
            name = cand
        except Exception:
            name = None
    _CJK_FONT_CACHE['name'] = name
    _CJK_FONT_CACHE['resolved'] = True
    return name or 'Helvetica'


def pdf_watermark(path: str, options: dict) -> FileResult:
    import io
    from pypdf import PdfReader, PdfWriter
    from reportlab.lib.colors import Color
    from reportlab.pdfgen import canvas as rl_canvas

    try:
        reader = _pdf_reader(path, options)
    except RuntimeError as e:
        return FileResult(path, 'fail', str(e))
    except Exception as e:
        # 截断/垃圾 PDF 抛的是 pypdf 的 PdfStreamError 等(非 RuntimeError),
        # 早期会漏出去让用户看到英文堆栈
        return FileResult(path, 'fail',
                          'PDF 无法读取(文件损坏或不是 PDF): '
                          + type(e).__name__)
    text = options.get('wm_text') or '青雁水印'
    # UI 标签是"水印透明度%": 100 表示完全透明, 因此要换算成 alpha
    transparency = max(1, min(100, int(options.get('wm_opacity', 15) or 15)))
    alpha = 1.0 - transparency / 100.0
    tile = bool(options.get('wm_tile', True))
    col = Color(0.5, 0.5, 0.5, alpha=max(0.02, alpha))
    font_name = _pdf_cjk_font(text)
    w = PdfWriter()
    for pg in reader.pages:
        pw = float(pg.mediabox.width)
        ph = float(pg.mediabox.height)
        buf = io.BytesIO()
        c = rl_canvas.Canvas(buf, pagesize=(pw, ph))
        c.setFillColor(col)
        c.setFont(font_name, 36)
        if tile:
            step_x, step_y = 220, 160
            yy = 30
            while yy < ph:
                xx = 2
                while xx < pw:
                    c.saveState()
                    c.translate(xx, yy)
                    c.rotate(30)
                    c.drawString(0, 0, text)
                    c.restoreState()
                    xx += step_x
                yy += step_y
        else:
            c.saveState()
            c.translate(pw / 2 - 60, ph / 2)
            c.rotate(30)
            c.drawString(0, 0, text)
            c.restoreState()
        c.save()
        buf.seek(0)
        overlay = PdfReader(buf).pages[0]
        pg.merge_page(overlay)
        w.add_page(pg)
    out_dir = options.get('out_dir') or os.path.dirname(path)
    os.makedirs(out_dir, exist_ok=True)
    base, ext = os.path.splitext(os.path.basename(path))
    out = os.path.abspath(os.path.join(out_dir, base + '_wm.pdf'))
    # 必须防覆盖: 目录里已有同名 _wm.pdf(可能是别的资料)会被静默覆盖;
    # 同批选 合同.pdf + 合同_wm.pdf 时更会把第二个源文件覆盖掉
    out = _uniq_path(out)
    with open(out, 'wb') as fh:
        w.write(fh)
    return FileResult(path, 'ok', f"水印'{text}'", out)


# ================================================================== 模板生成
def generate_from_template(template_path: str, row: dict, out_path: str) -> None:
    """{{字段}} 占位符模板 + 单行数据 -> 个性化文档。支持 docx/xlsx/pptx/txt。"""
    ext = os.path.splitext(template_path)[1].lower()

    def _subst(text: str) -> str:
        for k, v in row.items():
            # None 直接 str() 会把字面量 "None" 写进文档
            text = text.replace('{{' + str(k) + '}}',
                                '' if v is None else str(v))
        return text

    if ext == '.docx':
        import docx
        d = docx.Document(template_path)

        def _fill_para(p):
            runs = p.runs
            if not runs:
                return
            full = ''.join(r.text for r in runs)
            # 占位符常被 Word 拆到多个 run(如 '{{姓' + '名}}'),
            # 只按单 run 判断会导致占位符原样残留
            if '{{' not in full:
                return
            runs[0].text = _subst(full)
            for r in runs[1:]:
                r.text = ''

        for p in d.paragraphs:
            _fill_para(p)
        for t in d.tables:
            for row_ in t.rows:
                for cell in row_.cells:
                    for p in cell.paragraphs:
                        _fill_para(p)
        atomic_save(d.save, out_path)
        return
    if ext == '.xlsx':
        import openpyxl
        wb = openpyxl.load_workbook(template_path)
        for ws in wb.worksheets:
            for row_ in ws.iter_rows():
                for c in row_:
                    if isinstance(c.value, str) and '{{' in c.value:
                        c.value = _subst(c.value)
        atomic_save(wb.save, out_path)
        return
    if ext == '.pptx':
        from pptx import Presentation
        prs = Presentation(template_path)
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for p in shape.text_frame.paragraphs:
                        for r in p.runs:
                            if '{{' in r.text:
                                r.text = _subst(r.text)
        atomic_save(prs.save, out_path)
        return
    # txt: 必须显式限定扩展名, 否则 .pdf/.doc 等二进制文件会走进本分支,
    # 以 errors='replace' 读入再写出, 内容被彻底破坏
    if ext not in ('.txt', '.md', '.csv', '.log'):
        raise ValueError('不支持的模板格式: ' + ext +
                         '(仅支持 .docx/.xlsx/.pptx/.txt)')
    # 用自动编码探测读: 硬编码 utf-8 会把 GBK 模板(老记事本/业务系统
    # 导出, Windows 中文环境极常见)读成全篇 U+FFFD, 连 {{占位符}} 都
    # 匹配不上
    from engine.ops import read_text_auto
    text = read_text_auto(template_path)
    with open(out_path, 'w', encoding='utf-8') as fh:
        fh.write(_subst(text))


# ================================================================= 内置模板
OFFICE_TEMPLATES = {
    '请假条': {
        'ext': '.docx',
        'tpl_body': ('请假条', '', '尊敬的{{领导姓名}}：', '',
                     '    本人{{姓名}}（部门：{{部门}}），因{{请假事由}}，',
                     '需请假{{天数}}天（{{开始日期}} 至 {{结束日期}}）。',
                     '请假期间工作已交接给{{交接人}}，恳请批准。', '',
                     '                                申请人：{{姓名}}',
                     '                                {{申请日期}}'),
        'columns': ('领导姓名', '姓名', '部门', '请假事由', '天数',
                    '开始日期', '结束日期', '交接人', '申请日期'),
        'sample': (('王经理', '张三', '市场部', '家中有事', 2,
                    '2026-09-20', '2026-09-21', '李四', '2026-09-15'),
                   ('王经理', '赵六', '市场部', '身体不适', 1,
                    '2026-09-25', '2026-09-25', '李四', '2026-09-15')),
    },
    '会议纪要': {
        'ext': '.docx',
        'tpl_body': ('{{会议名称}} 会议纪要', '', '会议时间：{{会议时间}}',
                     '会议地点：{{会议地点}}',
                     '主持人：{{主持人}}    记录人：{{记录人}}',
                     '参会人员：{{参会人员}}', '', '一、会议议题', '{{会议议题}}',
                     '', '二、会议决议', '1. {{决议1}}', '2. {{决议2}}', '',
                     '三、后续行动',
                     '{{后续行动}}（责任人：{{责任人}}，截止：{{截止日期}}）'),
        'columns': ('会议名称', '会议时间', '会议地点', '主持人', '记录人',
                    '参会人员', '会议议题', '决议1', '决议2', '后续行动',
                    '责任人', '截止日期'),
        'sample': (('产品周会', '2026-09-14 10:00', '3号会议室', '陈总', '小刘',
                    '产品部全员', 'Q4 排期评审', '按期启动 v2.0', '增加压测环节',
                    '输出测试方案', '测试组', '2026-09-18'),),
    },
    '工作周报': {
        'ext': '.docx',
        'tpl_body': ('工作周报（{{周次}}）', '',
                     '姓名：{{姓名}}    部门：{{部门}}    日期：{{日期}}', '',
                     '一、本周完成', '{{本周完成}}', '', '二、下周计划',
                     '{{下周计划}}', '', '三、问题与风险', '{{问题风险}}'),
        'columns': ('周次', '姓名', '部门', '日期', '本周完成', '下周计划',
                    '问题风险'),
        'sample': (('第37周', '张三', '市场部', '2026-09-14',
                    '1.完成Q3复盘\n2.启动双11预热', '1.竞品分析\n2.物料设计',
                    '设计资源紧张'),),
    },
    '录用通知': {
        'ext': '.docx',
        'tpl_body': ('录用通知书', '', '{{姓名}} 先生/女士：', '',
                     '    感谢您应聘我司{{职位}}岗位。经评审，现正式通知您被录用。',
                     '报到时间：{{报到日期}} 上午{{报到时间}}',
                     '报到地点：{{报到地点}}',
                     '联系人：{{联系人}}（电话 {{联系电话}}）', '',
                     '    期待您的加入！', '',
                     '                                {{公司名称}}',
                     '                                {{发出日期}}'),
        'columns': ('姓名', '职位', '报到日期', '报到时间', '报到地点',
                    '联系人', '联系电话', '公司名称', '发出日期'),
        'sample': (('李四', '软件工程师', '2026-10-09', '09:30',
                    '科技园A座12F', '人事部王女士', '010-88886666',
                    '某某科技有限公司', '2026-09-15'),),
    },
    '工资条': {
        'ext': '.xlsx',
        'tpl_body': ('{{月份}} 工资条',),
        'columns': ('月份', '姓名', '基本工资', '绩效', '补贴', '社保扣款',
                    '个税', '实发'),
        'sample': (('2026-09', '张三', 8000, 2000, 500, 1200, 300, 9000),
                   ('2026-09', '李四', 9000, 2500, 600, 1500, 450, 10150)),
    },
    '快递通知': {
        'ext': '.docx',
        'tpl_body': ('取件通知', '', '{{姓名}} 您好：',
                     '    您的快递（运单号 {{运单号}}）已到达{{驿站名称}}，',
                     '取件码 {{取件码}}，请于 {{保留期限}} 前领取。',
                     '驿站地址：{{驿站地址}}'),
        'columns': ('姓名', '运单号', '驿站名称', '取件码', '保留期限',
                    '驿站地址'),
        'sample': (('王小明', 'SF1234567890', '丰巢驿站', '8-2-3009',
                    '2026-09-18', '小区南门东侧'),),
    },
}


def list_office_templates() -> list:
    """模板名列表(供 GUI 下拉)。"""
    return sorted(OFFICE_TEMPLATES)


def materialize_office_template(name: str, tpl_dir: str):
    """把内置模板物化为真实模板文件 + 示例数据 xlsx。

    返回 (模板路径, 数据路径, 字段列表)。
    """
    import openpyxl

    spec = OFFICE_TEMPLATES.get(name)
    if not spec:
        raise ValueError('未知模板: ' + name)
    os.makedirs(tpl_dir, exist_ok=True)
    safe = re.sub(r'[\\/:*?"<>|]', '_', name)
    if spec['ext'] == '.docx':
        import docx
        tpl_path = os.path.join(tpl_dir, safe + '.docx')
        d = docx.Document()
        for ln in spec['tpl_body']:
            d.add_paragraph(ln)
        d.save(tpl_path)
    else:
        tpl_path = os.path.join(tpl_dir, safe + '.xlsx')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = '模板'
        for i, col in enumerate(spec['columns'], 1):
            ws.cell(row=1, column=i, value='{{' + col + '}}')
        wb.save(tpl_path)
    data_path = os.path.join(tpl_dir, safe + '_数据.xlsx')
    wb2 = openpyxl.Workbook()
    ws2 = wb2.active
    ws2.title = '数据'
    for cidx, val in enumerate(spec['columns'], 1):
        ws2.cell(row=1, column=cidx, value=val)
    for r, row in enumerate(spec['sample'], 2):
        for cidx, val in enumerate(row, 1):
            ws2.cell(row=r, column=cidx, value=val)
    wb2.save(data_path)
    return tpl_path, data_path, list(spec['columns'])
