# -*- coding: utf-8 -*-
"""青雁批量工具箱 v3 — 批量打印功能页。

路由见 engine/print_ops.py。核心亮点:
- 图片/文本直接 GDI 打印; PDF 内置渲染; Office 自动转 PDF 后打印
- 选择「Microsoft Print to PDF」时可静默输出 PDF(虚拟测试, 不耗纸)
"""
from __future__ import annotations

import os

from PySide6.QtWidgets import (QCheckBox, QComboBox, QFileDialog, QLabel,
                               QLineEdit, QPushButton, QSpinBox)

from app.pages.base import FunctionPage


def _combo(items):
    c = QComboBox()
    c.addItems(list(items))
    return c


class PrintPage(FunctionPage):
    title = '批量打印'
    kind = 'print'
    MODIFIES_SOURCE = False     # 打印不改动源文件
    desc = ('把图片/文本/PDF/Office 文档批量送到打印机。'
            '图片与文本直接打印；PDF 内置渲染保持版式；Office 自动转换为 '
            'PDF 后打印(需本机装有 Office 或 WPS)。\n'
            '虚拟测试: 选择「Microsoft Print to PDF」并填写下方输出目录，'
            '打印结果会变成 PDF 文件，不耗纸。')
    OPS = ('批量打印',)
    OP_FIELDS = {}
    ALWAYS_VISIBLE = ('printer', 'refresh_btn', 'copies', 'fit_page',
                      'pdf_out_dir', 'pdf_out_btn')
    REQUIRED_FIELDS = {'批量打印': (('printer', '打印机'),)}
    EXPECTED_EXTS = ()

    def build_options(self, form):
        self._field(form, 'op', '操作', _combo(self.OPS))
        self._field(form, 'printer', '打印机', QComboBox())
        rb = QPushButton('刷新打印机列表')
        self._field(form, 'refresh_btn', '', rb)
        rb.clicked.connect(self._reload_printers)
        cp = QSpinBox()
        cp.setRange(1, 99)
        cp.setValue(1)
        cp.setToolTip('每个文件打印几份')
        self._field(form, 'copies', '打印份数', cp)
        self._field(form, 'fit_page', '',
                    QCheckBox('图片/页面适应纸张大小(居中)'))
        self._widgets['fit_page'].setChecked(True)
        od = QLineEdit()
        od.setPlaceholderText('仅虚拟打印机(Print to PDF)生效: PDF 输出目录')
        self._field(form, 'pdf_out_dir', 'PDF 输出目录', od)
        ob = QPushButton('选输出目录')
        self._field(form, 'pdf_out_btn', '', ob)
        ob.clicked.connect(self._pick_out)
        info = QLabel('类型路由: 图片/文本 → 直接打印 · PDF → 内置渲染 · '
                      'Office(doc/xls/ppt) → 本机转 PDF 后打印')
        info.setWordWrap(True)
        info.setStyleSheet(f'color:{self.main.theme.MUTED};font-size:12px;')
        form.addRow('', info)
        self._reload_printers()

    def _reload_printers(self):
        from engine.print_ops import default_printer, list_printers
        cb = self._widgets['printer']
        # 必须在 clear() 之前读: clear() 会把 currentText 清成 '',
        # 之后再读就取不到用户原来的选择(刷新后选择被重置)
        cur = cb.currentText()
        cb.clear()
        ps = list_printers()
        if ps:
            cb.addItems(ps)
            d = default_printer()
            if cur in ps:
                cb.setCurrentText(cur)
            elif d in ps:
                cb.setCurrentText(d)
        else:
            cb.addItem('未检测到打印机')

    def _pick_out(self):
        d = QFileDialog.getExistingDirectory(self, '选择 PDF 输出目录')
        if d:
            self._widgets['pdf_out_dir'].setText(d)

    def collect_options(self):
        printer = self._widgets['printer'].currentText()
        if printer.startswith('未检测到'):
            printer = ''
        o = {'printer': printer,
             'copies': self._widgets['copies'].value(),
             'fit_page': self._widgets['fit_page'].isChecked(),
             'pdf_out_dir': self._widgets['pdf_out_dir'].text().strip(),
             'op_name': '批量打印'}
        return o

    def processor(self):
        from engine.print_ops import print_file
        return print_file

    def validate_required(self):
        """在基类必填校验之上, 追加"虚拟打印机必须填输出目录"。

        选了 Print to PDF 但不填目录时, 驱动会在工作线程里对**每个文件**
        弹一次系统"另存为"对话框, 批量任务看起来就像卡死(取消也要等当前
        文件处理完)。宁可在开跑前拦住。
        """
        missing = super().validate_required()
        try:
            from engine.print_ops import is_virtual_pdf_printer
            printer = self._widgets['printer'].currentText()
            out = self._widgets['pdf_out_dir'].text().strip()
            if printer and is_virtual_pdf_printer(printer) and not out:
                missing.append('PDF 输出目录(虚拟打印机必填)')
        except Exception:
            pass
        return missing

    def supports_preview(self):
        return True

    def build_preview(self, files, options):
        from engine.print_ops import is_virtual_pdf_printer, list_printers
        rows = [('文件', '打印方式')]
        printer = options.get('printer', '')
        known = printer in list_printers()
        for f in files[:100]:
            ext = os.path.splitext(f)[1].lower()
            if ext in ('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.webp',
                       '.tif', '.tiff'):
                how = 'GDI 直印'
            elif ext == '.pdf':
                how = '内置渲染直印(未装组件时用系统关联)'
            elif ext in ('.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx',
                         '.wps', '.et', '.dps'):
                how = '本机转 PDF 后直印(需 Office/WPS)'
            elif os.path.splitext(f)[1].lower() in (
                    '.txt', '.md', '.csv', '.log', '.json', '.xml', '.ini',
                    '.yaml', '.yml', '.srt'):
                how = 'GDI 文本打印'
            else:
                how = '跳过(不支持的类型)'
            if printer and not known:
                how += ' — 打印机不可用'
            elif printer and is_virtual_pdf_printer(printer):
                how += ' (虚拟打印 → PDF 文件)'
            rows.append((os.path.basename(f), how))
        return rows
