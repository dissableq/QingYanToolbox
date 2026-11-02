# -*- coding: utf-8 -*-
"""青雁批量工具箱 v3 — 格式转换功能页。

转换结果全部为「真格式」产物, 可继续编辑:
- Office(doc/xls/ppt) → PDF: 本机 COM 导出, 版式保真
- PDF → Word: 提取结构化文本重建 .docx(字号→标题/粗体/居中), 每段可编辑
- PDF → Excel: 识别表格画线求网格, 按词坐标归格(不用自带归位有缺陷的
  table.extract), 每页一个工作表
- PDF → 图片 / TXT / HTML / PPT
- Word → TXT / HTML(含表格) / Markdown
- Excel → CSV / HTML / Word / JSON
- PPT → TXT / 图片(COM 渲染)
- 图片 → PDF / Word; 文本 → Word; CSV → Excel; HTML → Word

引擎: engine/convert_ops.py。转换产物默认放到源文件旁的「_转换结果」
目录(也可指定输出目录), 原文件不动, 无需备份与撤销。
"""
from __future__ import annotations

import os

from PySide6.QtWidgets import (QComboBox, QFileDialog, QLabel, QLineEdit,
                               QPushButton)

from app.pages.base import FunctionPage


def _combo(items):
    c = QComboBox()
    c.addItems(list(items))
    return c


# 转换方向: 显示名 → (机内键, 白话说明, 源类型提示)
CONVERT_LIST = (
    ('Office → PDF', 'office2pdf',
     '把 Word / Excel / PPT 转成 PDF，版式不变，方便传阅',
     '源文件：.docx .doc .xlsx .xls .pptx .ppt 等'),
    ('PDF → Word', 'pdf2word',
     '把 PDF 转成可编辑的 Word 文档（文字、标题、加粗都会保留）',
     '源文件：.pdf'),
    ('PDF → 图片', 'pdf2img',
     '把 PDF 每一页转成 PNG 图片',
     '源文件：.pdf'),
    ('PDF → TXT', 'pdf2txt',
     '抽取 PDF 里的纯文字',
     '源文件：.pdf'),
    ('PDF → HTML', 'pdf2html',
     '把 PDF 转成网页，保留标题层级和居中排版',
     '源文件：.pdf'),
    ('Word → TXT', 'word2txt',
     '抽取 Word 里的文字（含表格内容）',
     '源文件：.docx .doc'),
    ('Word → HTML', 'word2html',
     '把 Word 转成网页，保留标题层级和表格',
     '源文件：.docx'),
    ('Excel → CSV', 'excel2csv',
     '把 Excel 首个工作表转成 CSV（Excel 可直接打开）',
     '源文件：.xlsx .xls'),
    ('Excel → HTML', 'excel2html',
     '把 Excel 全部工作表转成一个网页表格',
     '源文件：.xlsx'),
    ('PPT → TXT', 'ppt2txt',
     '抽取 PPT 每页的文字',
     '源文件：.pptx .ppt'),
    ('PPT → 图片', 'ppt2img',
     '把 PPT 每一页转成 PNG 图片（需要本机装有 Office/WPS）',
     '源文件：.pptx .ppt'),
    ('图片 → PDF', 'img2pdf',
     '把图片转成 PDF（原图尺寸，一页一图）',
     '源文件：.png .jpg .bmp .webp 等'),
    ('文本 → Word', 'txt2word',
     '把 TXT/Markdown 等文本转成可编辑的 Word 文档',
     '源文件：.txt .md .json 等文本'),
    ('PDF → Excel', 'pdf2excel',
     '识别 PDF 里的表格，转成可编辑的 Excel（每页一个工作表）',
     '源文件：.pdf（带边框线的表格）'),
    ('CSV → Excel', 'csv2excel',
     '把 CSV/TSV 转成 Excel，表头加粗并冻结首行',
     '源文件：.csv .tsv'),
    ('Excel → Word', 'excel2word',
     '把 Excel 每个工作表转成 Word 里的表格，可继续编辑',
     '源文件：.xlsx'),
    ('Excel → JSON', 'excel2json',
     '按表头把 Excel 转成 JSON 对象数组（全部工作表）',
     '源文件：.xlsx'),
    ('Word → Markdown', 'word2md',
     '把 Word 转成 Markdown，保留标题层级、粗体/斜体和表格',
     '源文件：.docx'),
    ('HTML → Word', 'html2word',
     '把网页 HTML 转成 Word 文档，保留标题层级',
     '源文件：.html .htm'),
    ('图片 → Word', 'img2word',
     '把图片插入 Word 文档生成图文文件（图片内的文字不会自动识别）',
     '源文件：.png .jpg .bmp .webp 等'),
    ('PDF → PPT', 'pdf2ppt',
     '把 PDF 每一页贴进 PPT（16:9 居中），保持页面外观，页内文字为图片',
     '源文件：.pdf'),
)


class ConvertPage(FunctionPage):
    title = '格式转换'
    kind = 'convert'
    MODIFIES_SOURCE = False     # 转换产物另存, 源文件从不改动
    desc = ('批量转换文档格式。文字型 PDF 可提取为可编辑文字；扫描件需先完成 OCR。'
            'Office 转 PDF 使用本机 Office/WPS 渲染。复杂版式请先预览转换结果。\n'
            '转换产物默认放在源文件旁的「_转换结果」文件夹，原文件保持不动。')
    OPS = tuple(c[0] for c in CONVERT_LIST)
    OP_FIELDS = {}
    ALWAYS_VISIBLE = ('out_dir', 'out_btn', 'conv_info')
    REQUIRED_FIELDS = {}
    EXPECTED_EXTS = ()   # 由 convert_ops.supported_exts 动态守卫

    def build_options(self, form):
        self._field(form, 'op', '转换方向', _combo(self.OPS))
        od = QLineEdit()
        od.setPlaceholderText('留空=源文件旁「_转换结果」文件夹')
        self._field(form, 'out_dir', '输出目录', od)
        ob = QPushButton('输出到…')
        ob.clicked.connect(self.pick_out)
        form.addRow('', ob)
        self._conv_info = QLabel('')
        self._conv_info.setWordWrap(True)
        self._conv_info.setStyleSheet(
            f'color:{self.main.theme.MUTED};font-size:12px;')
        form.addRow('', self._conv_info)
        self._widgets['op'].currentIndexChanged.connect(
            lambda _=None: self._update_info())
        self._update_info()

    def _update_info(self):
        op = self._widgets['op'].currentText()
        item = next((c for c in CONVERT_LIST if c[0] == op), None)
        if item is None:
            self._conv_info.setText('')
            return
        self._conv_info.setText(f'💡 {item[2]}\n支持 {item[3]}')

    def pick_out(self):
        d = QFileDialog.getExistingDirectory(self, '选择输出目录')
        if d:
            self._widgets['out_dir'].setText(d)

    def collect_options(self):
        op = self._widgets['op'].currentText()
        key = next((c[1] for c in CONVERT_LIST if c[0] == op), '')
        return {'op': key, 'convert': key, 'op_name': op,
                'out_dir': self._widgets['out_dir'].text().strip()}

    def supports_preview(self):
        return True

    def build_preview(self, files, options):
        from engine.convert_ops import supported_exts
        key = options.get('convert', '')
        exts = supported_exts(key)
        rows = [('文件', '将执行')]
        for f in files[:100]:
            ext = os.path.splitext(f)[1].lower()
            if not exts:
                ok = '转换'
            elif ext in exts:
                ok = '转换'
            else:
                ok = f'跳过(不支持 {ext})'
            rows.append((os.path.basename(f), ok))
        return rows

    def processor(self):
        from engine.convert_ops import convert_file
        return convert_file

    # ------------------------------------------------ 扩展名守卫(动态)
    def _guarded(self, proc):
        """按当前转换方向支持的扩展名过滤, 不符的优雅跳过。"""
        from engine.convert_ops import supported_exts
        key = self.collect_options().get('convert', '')
        exts = supported_exts(key)

        def wrapper(path, options):
            ext = os.path.splitext(str(path))[1].lower()
            if exts and ext not in exts:
                from engine.core import FileResult
                return FileResult(
                    path, 'skipped',
                    f'该转换方向不支持 {ext}(支持: {" ".join(exts)})')
            return proc(path, options)

        return wrapper
