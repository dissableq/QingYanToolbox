# -*- coding: utf-8 -*-
"""青雁批量工具箱 v3 — 九大功能页 (Word/Excel/PPT/PDF/文本/图片/重命名/
文件整理/模板批量生成)。

每页: OPS(操作清单) + OP_FIELDS(操作→字段显隐) + build_options +
collect_options + processor(操作→处理器函数映射)。
"""
from __future__ import annotations

import os

from PySide6.QtWidgets import (QCheckBox, QComboBox, QDoubleSpinBox,
                               QFileDialog, QLineEdit, QMessageBox,
                               QPlainTextEdit, QPushButton, QSpinBox)

from app.pages.base import FunctionPage
from engine.core import FileResult
from engine import advanced_ops, office_ops, ops


def _combo(items):
    c = QComboBox()
    c.addItems(list(items))
    return c


class WordPage(FunctionPage):
    title = 'Word 工具'
    kind = 'word'
    EXPECTED_EXTS = ('.docx', '.docm')
    OP_HINTS = {
        '查找替换': '把文档里的指定文字替换成新文字，支持正则表达式',
        '删除空白段落': '删除文档中没有任何内容的空行，让排版更紧凑',
        '删除全部图片': '删除文档里插入的所有图片',
        '删除批注': '删除审阅时添加的所有批注气泡',
        '取消超链接': '去掉链接跳转，但保留链接文字本身',
        '删除域': '删除域代码（自动编号/目录等），可选转为普通文字',
        '接受修订': '接受全部修订标记，文档变成最终稿',
        '拒绝修订': '拒绝全部修订，恢复到修改前的样子',
        '删除宏': '移除文档里的宏组件，防止宏病毒',
        '删除水印': '清除页眉里的艺术字/图片水印',
        '删除页眉页脚': '清空每页顶部和底部的页眉页脚内容',
        '设置页眉页脚': '给每页添加统一的页眉/页脚文字',
        '添加水印': '在每页页眉位置加一行浅色水印文字',
        '删除或修改背景': '设置或清除整份文档的背景颜色',
        '提取图片': '把文档里所有图片导出成图片文件，方便复用',
        '设置内容样式': '一键统一正文的字体、字号、对齐和行距',
        '设置页面样式': '统一设置页边距和纸张方向（横/纵向）',
        '清空元数据': '抹掉作者、标题等文档属性，保护隐私',
        '修改元数据': '修改文档的作者/标题等属性',
        '压缩(重打包)': '重新压缩文件，体积变小，内容不变',
        '按段落拆分': '把长文档按每 N 段拆成多个文件',
    }
    desc = ('Word 批量处理：内容替换、图片/批注/链接/域/修订/宏清理、'
            '页眉页脚、水印、样式、元数据等')
    OPS = ('查找替换', '删除空白段落', '删除全部图片', '删除批注',
           '取消超链接', '删除域', '接受修订', '拒绝修订', '删除宏',
           '删除水印', '删除页眉页脚', '设置页眉页脚', '添加水印',
           '删除或修改背景', '提取图片', '设置内容样式', '设置页面样式',
           '清空元数据', '修改元数据', '压缩(重打包)', '按段落拆分')
    REQUIRED_FIELDS = {'查找替换': (('find', '查找内容'),),
                       '添加水印': (('wm', '水印文字'),),
                       }
    OP_FIELDS = {
        '查找替换': ('find', 'replace', 'regex_cb'),
        '添加水印': ('wm',),
        '设置页眉页脚': ('header', 'footer'),
        '删除或修改背景': ('bg_color', 'bg_remove'),
        '删除域': ('field_mode',),
        '设置内容样式': ('style_font', 'style_size', 'style_align', 'style_ls'),
        '设置页面样式': ('page_margin', 'page_orient'),
        '清空元数据': (),
        '修改元数据': ('meta_author', 'meta_title'),
        '压缩(重打包)': (),
        '按段落拆分': ('docx_chunk',),
    }

    def build_options(self, form):
        self._field(form, 'op', '操作', _combo(self.OPS))
        self._field(form, 'find', '查找', QLineEdit())
        self._field(form, 'replace', '替换为', QLineEdit())
        self._field(form, 'regex_cb', '', QCheckBox('使用正则表达式'))
        wm = QLineEdit()
        wm.setPlaceholderText('')
        self._field(form, 'wm', '水印文字', wm)
        self._field(form, 'header', '页眉', QLineEdit())
        self._field(form, 'footer', '页脚', QLineEdit())
        bgc = QLineEdit()
        bgc.setPlaceholderText('#RRGGBB, 如 #EAF5F9')
        self._field(form, 'bg_color', '背景', bgc)
        self._field(form, 'bg_remove', '', QCheckBox('删除背景'))
        # 引擎支持 删除 / 转为纯文本 两种处理, 但早期 UI 从不发 field_mode,
        # 于是「删除域」永远只会删掉, 提示里承诺的"可选转为普通文字"做不到。
        self._field(form, 'field_mode', '域处理方式',
                    _combo(('删除', '转为纯文本')))
        self._field(form, 'style_font', '正文字体',
                    _combo(('', '微软雅黑', '宋体', '黑体', '楷体', '仿宋',
                            'Times New Roman', 'Arial')))
        ss = QSpinBox()
        ss.setRange(1, 72)
        ss.setValue(12)
        self._field(form, 'style_size', '字号(pt)', ss)
        self._field(form, 'style_align', '对齐',
                    _combo(('', '左', '居中', '右', '两端')))
        ls = QDoubleSpinBox()
        ls.setRange(0.5, 5)
        ls.setSingleStep(0.25)
        ls.setValue(1.5)
        self._field(form, 'style_ls', '行距(倍)', ls)
        pm = QDoubleSpinBox()
        pm.setRange(0, 10)
        pm.setSingleStep(0.5)
        pm.setValue(2.54)
        self._field(form, 'page_margin', '页边距(cm)', pm)
        self._field(form, 'page_orient', '纸张方向',
                    _combo(('', '横向', '纵向')))
        self._field(form, 'meta_author', '作者', QLineEdit())
        self._field(form, 'meta_title', '标题', QLineEdit())
        ch = QSpinBox()
        ch.setRange(1, 100000)
        ch.setValue(50)
        self._field(form, 'docx_chunk', '每份段落数(拆分)', ch)
        od = QLineEdit()
        od.setPlaceholderText('留空=就地修改(自动保存到“备份”文件夹)')
        self._field(form, 'out_dir', '输出目录', od)
        ob = QPushButton('输出到…')
        ob.clicked.connect(self.pick_out)
        form.addRow('', ob)
        self.ext_edit.setText('.docx,.docm')

    def pick_out(self):
        d = QFileDialog.getExistingDirectory(self, '选择输出目录')
        if d:
            self._widgets['out_dir'].setText(d)

    def collect_options(self):
        op = self._widgets['op'].currentText()
        key_map = {'查找替换': 'replace', '删除空白段落': 'clear',
                   '删除全部图片': 'rmimg', '删除批注': 'nocmt',
                   '取消超链接': 'nolink', '删除域': 'nofld',
                   '接受修订': 'rev_acc', '拒绝修订': 'rev_rej',
                   '删除宏': 'nomacro', '删除水印': 'nowm',
                   '删除页眉页脚': 'hf_del', '设置页眉页脚': 'hf',
                   '添加水印': 'wm', '删除或修改背景': 'bg',
                   '提取图片': 'extimg', '设置内容样式': 'style',
                   '设置页面样式': 'page', '清空元数据': 'meta_c',
                   '修改元数据': 'meta_m', '压缩(重打包)': 'compress',
                   '按段落拆分': 'split'}
        o = {'op': key_map.get(op, ''), 'op_name': op}
        if op == '查找替换':
            o['find'] = self._widgets['find'].text()
            o['replace'] = self._widgets['replace'].text()
            o['regex'] = self._widgets['regex_cb'].isChecked()
        elif op == '添加水印':
            o['wm_text'] = self._widgets['wm'].text() or '青雁水印'
        elif op == '设置页眉页脚':
            o['header_text'] = self._widgets['header'].text()
            o['footer_text'] = self._widgets['footer'].text()
        elif op == '删除域':
            o['field_mode'] = self._widgets['field_mode'].currentText()
        elif op == '删除或修改背景':
            o['bg_color'] = self._widgets['bg_color'].text()
            o['bg_remove'] = self._widgets['bg_remove'].isChecked()
        elif op == '设置内容样式':
            o['style_font'] = self._widgets['style_font'].currentText()
            o['style_size'] = self._widgets['style_size'].value()
            o['style_align'] = self._widgets['style_align'].currentText()
            o['style_line_spacing'] = self._widgets['style_ls'].value()
        elif op == '设置页面样式':
            o['page_margin_cm'] = self._widgets['page_margin'].value()
            o['page_orientation'] = self._widgets['page_orient'].currentText()
        elif op == '修改元数据':
            o['meta_mode'] = '修改'
            o['meta_author'] = self._widgets['meta_author'].text()
            o['meta_title'] = self._widgets['meta_title'].text()
        elif op == '清空元数据':
            o['meta_mode'] = '清空'
        elif op == '按段落拆分':
            o['docx_chunk'] = self._widgets['docx_chunk'].value()
        if 'out_dir' in self._widgets:
            o['out_dir'] = self._widgets['out_dir'].text().strip()
        return o

    def processor(self):
        op = self.collect_options().get('op')
        return {
            'replace': office_ops.docx_replace, 'clear': office_ops.docx_clear_blank,
            'rmimg': office_ops.docx_remove_images,
            'nocmt': advanced_ops.docx_remove_comments,
            'nolink': advanced_ops.docx_remove_links,
            'nofld': advanced_ops.docx_remove_fields,
            'rev_acc': advanced_ops.docx_accept_revisions,
            'rev_rej': lambda p, o: advanced_ops.docx_accept_revisions(
                p, {**o, 'rev_mode': '拒绝'}),
            'nomacro': advanced_ops.docx_remove_macros,
            'nowm': advanced_ops.docx_remove_watermark,
            'hf_del': advanced_ops.docx_remove_header_footer,
            'hf': office_ops.docx_header_footer, 'wm': office_ops.docx_watermark,
            'bg': advanced_ops.docx_set_background, 'extimg': advanced_ops.docx_extract_images,
            'style': advanced_ops.docx_set_content_style,
            'page': advanced_ops.docx_set_page_style,
            'meta_c': advanced_ops.docx_meta, 'meta_m': advanced_ops.docx_meta,
            'compress': office_ops.docx_compress,
            'split': advanced_ops.docx_split_by_paragraphs,
        }.get(op)


class ExcelPage(FunctionPage):
    title = 'Excel 工具'
    kind = 'excel'
    EXPECTED_EXTS = ('.xlsx', '.xlsm')
    OP_HINTS = {
        '查找替换': '替换所有工作表单元格里的文字，支持正则表达式',
        '删除空行': '删除整行都没有内容的行',
        '删除重复行': '整行内容完全相同的行只保留第一行',
        '删除空白列': '删除整列都没有内容的列',
        '删除空白 Sheet': '删除没有任何数据的空工作表',
        '公式转值': '把公式固定成计算结果，防止引用失效变成 #REF',
        '删除图片': '删除表格里的浮动图片',
        '删除宏': '把含宏的 xlsm 另存为不含宏的 xlsx',
        '删除页眉页脚': '清空打印时每页顶部/底部的页眉页脚',
        '设置页眉页脚': '设置打印时每页的页眉/页脚文字',
        '添加水印': '打印预览时每页页眉显示浅色水印文字',
        'Sheet 改名': '批量替换工作表名称里的文字',
        'Sheet 排序': '按名称顺序给所有工作表排序',
        '删除指定 Sheet': '按名称删除指定的工作表（逗号分隔多个）',
        '插入 Sheet': '在指定位置插入一个新的空白工作表',
        '按 Sheet 拆分': '每个工作表另存为一个独立文件',
        '按行数拆分': '把数据按每 N 行拆成多个文件（首行表头自动保留）',
        '转 CSV': '另存为 CSV 纯文本格式（取第一个工作表）',
        '转 TXT': '另存为制表符分隔的 TXT 文本',
        '转 HTML': '另存为带表格边框的网页文件',
        '清空元数据': '抹掉作者、标题等文档属性，保护隐私',
        '修改元数据': '修改文档的作者等属性',
        '压缩重存': '重新压缩文件，体积变小，内容不变',
    }
    desc = ('Excel 批量处理：内容替换、行列清理、公式、Sheet 管理、页眉页脚、'
            '拆分合并、格式转换、元数据')
    OPS = ('查找替换', '删除空行', '删除重复行', '删除空白列', '删除空白 Sheet',
           '公式转值', '删除图片', '删除宏', '删除页眉页脚', '设置页眉页脚',
           '添加水印', 'Sheet 改名', 'Sheet 排序', '删除指定 Sheet',
           '插入 Sheet', '按 Sheet 拆分', '按行数拆分', '转 CSV', '转 TXT',
           '转 HTML', '清空元数据', '修改元数据', '压缩重存')
    REQUIRED_FIELDS = {'查找替换': (('find', '查找内容'),),
                       '添加水印': (('wm', '水印文字'),)}
    OP_FIELDS = {
        '查找替换': ('find', 'replace', 'regex_cb'),
        '添加水印': ('wm',),
        '设置页眉页脚': ('header', 'footer'),
        'Sheet 改名': ('find', 'replace'),
        '删除指定 Sheet': ('sheet_names',),
        '插入 Sheet': ('sheet_pos', 'sheet_new_name'),
        '按行数拆分': ('rows_per',),
        '修改元数据': ('meta_author',),
    }

    def build_options(self, form):
        self._field(form, 'op', '操作', _combo(self.OPS))
        self._field(form, 'find', '查找', QLineEdit())
        self._field(form, 'replace', '替换为', QLineEdit())
        self._field(form, 'regex_cb', '', QCheckBox('使用正则表达式'))
        self._field(form, 'wm', '水印文字', QLineEdit())
        self._field(form, 'header', '页眉', QLineEdit())
        self._field(form, 'footer', '页脚', QLineEdit())
        sn = QLineEdit()
        sn.setPlaceholderText('Sheet 名, 逗号分隔')
        self._field(form, 'sheet_names', '指定 Sheet', sn)
        sp = QSpinBox()
        sp.setRange(1, 999)
        sp.setValue(1)
        self._field(form, 'sheet_pos', '插入位置(第N个)', sp)
        nn = QLineEdit()
        nn.setPlaceholderText('新 Sheet 名')
        self._field(form, 'sheet_new_name', '新名称', nn)
        rp = QSpinBox()
        rp.setRange(1, 1000000)
        rp.setValue(1000)
        self._field(form, 'rows_per', '每份行数(拆分)', rp)
        self._field(form, 'meta_author', '作者', QLineEdit())
        od = QLineEdit()
        od.setPlaceholderText('留空=就地修改(自动保存到“备份”文件夹)')
        self._field(form, 'out_dir', '输出目录', od)
        ob = QPushButton('输出到…')
        ob.clicked.connect(self.pick_out)
        form.addRow('', ob)
        self.ext_edit.setText('.xlsx,.xlsm')

    def pick_out(self):
        d = QFileDialog.getExistingDirectory(self, '选择输出目录')
        if d:
            self._widgets['out_dir'].setText(d)

    def _collect_op(self):
        op = self._widgets['op'].currentText()
        key_map = {'查找替换': 'replace', '删除空行': 'del_blank',
                   '删除重复行': 'del_dup', '删除空白列': 'del_blank_col',
                   '删除空白 Sheet': 'del_blank_sheet', '公式转值': 'formula',
                   '删除图片': 'rmimg', '删除宏': 'nomacro',
                   '删除页眉页脚': 'hf_del', '设置页眉页脚': 'hf',
                   '添加水印': 'wm', 'Sheet 改名': 'rename',
                   'Sheet 排序': 'sort', '删除指定 Sheet': 'sheet_del',
                   '插入 Sheet': 'sheet_ins', '按 Sheet 拆分': 'split_sheet',
                   '按行数拆分': 'split_rows', '转 CSV': 'csv',
                   '转 TXT': 'txt', '转 HTML': 'html',
                   '清空元数据': 'meta_c', '修改元数据': 'meta_m',
                   '压缩重存': 'compress'}
        o = {'op': op, 'key': key_map.get(op, '')}
        if op == '查找替换':
            o['find'] = self._widgets['find'].text()
            o['replace'] = self._widgets['replace'].text()
            o['regex'] = self._widgets['regex_cb'].isChecked()
        elif op == '添加水印':
            o['wm_text'] = self._widgets['wm'].text() or '青雁水印'
        elif op == '设置页眉页脚':
            o['header_text'] = self._widgets['header'].text()
            o['footer_text'] = self._widgets['footer'].text()
        elif op == 'Sheet 改名':
            # 引擎按 find/replace 做子串替换改名; 早期这个分支被上一行的
            # in (...) 提前吃掉, 于是永远传不出 find/replace, 功能全废。
            o['find'] = self._widgets['find'].text()
            o['replace'] = self._widgets['replace'].text()
        elif op == '删除指定 Sheet':
            o['sheet_names'] = self._widgets['sheet_names'].text()
        elif op == '插入 Sheet':
            o['sheet_pos'] = self._widgets['sheet_pos'].value()
            o['new_sheet_name'] = self._widgets['sheet_new_name'].text()
        elif op == '按行数拆分':
            o['xlsx_rows_per'] = self._widgets['rows_per'].value()
        elif op == '修改元数据':
            o['meta_mode'] = '修改'
            o['meta_author'] = self._widgets['meta_author'].text()
        elif op == '清空元数据':
            o['meta_mode'] = '清空'
        if 'out_dir' in self._widgets:
            o['out_dir'] = self._widgets['out_dir'].text().strip()
        return o

    def collect_options(self):
        return self._collect_op()

    def _convert(self, path, options):
        tgt = options.get('key')
        return advanced_ops.xlsx_convert_format(
            path, {**options, 'xlsx_target': tgt})

    def _del_blank_col(self, path, options):
        import openpyxl
        wb = openpyxl.load_workbook(path)
        n = 0
        for ws in wb.worksheets:
            for c in range(ws.max_column, 0, -1):
                if all((ws.cell(row=r, column=c).value in (None, ''))
                       for r in range(1, ws.max_row + 1)):
                    ws.delete_cols(c)
                    n += 1
        if n == 0:
            return FileResult(path, 'skipped', '无空白列')
        out, undo = advanced_ops._out(path, options, '_nc')
        wb.save(out)
        return FileResult(path, 'ok', f'删除空白列 {n} 个', out, undo=undo)

    def processor(self):
        o = self._collect_op()
        key = o.get('key')
        return {
            'csv': self._convert, 'txt': self._convert, 'html': self._convert,
            'del_blank_col': self._del_blank_col,
            'replace': office_ops.xlsx_replace,
            'del_blank': office_ops.xlsx_delete_rows,
            'del_dup': lambda p, op: office_ops.xlsx_delete_rows(
                p, {**op, 'row_mode': 'duplicate'}),
            'formula': office_ops.xlsx_formulas_to_values,
            'rmimg': advanced_ops.xlsx_remove_images,
            'nomacro': advanced_ops.xlsx_remove_macros,
            'hf_del': advanced_ops.xlsx_remove_header_footer,
            'hf': advanced_ops.xlsx_set_header_footer,
            'wm': office_ops.xlsx_watermark,
            'rename': lambda p, op: office_ops.xlsx_sheet_ops(
                p, {**op, 'sheet_op': 'rename'}),
            'sort': lambda p, op: office_ops.xlsx_sheet_ops(
                p, {**op, 'sheet_op': 'sort'}),
            'del_blank_sheet': lambda p, op: office_ops.xlsx_sheet_ops(
                p, {**op, 'sheet_op': 'delete_blank'}),
            'sheet_del': advanced_ops.xlsx_sheet_delete_replace,
            'sheet_ins': advanced_ops.xlsx_sheet_insert,
            'split_sheet': advanced_ops.xlsx_split,
            'split_rows': lambda p, op: advanced_ops.xlsx_split(
                p, {**op, 'xlsx_split_mode': 'rows'}),
            'meta_c': advanced_ops.xlsx_meta, 'meta_m': advanced_ops.xlsx_meta,
            'compress': office_ops.xlsx_compress,
        }.get(key)


class PptPage(FunctionPage):
    title = 'PPT 工具'
    kind = 'ppt'
    EXPECTED_EXTS = ('.pptx', '.pptm')
    OP_HINTS = {
        '查找替换': '替换每一页文本框里的文字',
        '删除空白页': '删除没有任何文字的页面（纯图片页不受影响）',
        '删除图片': '删除所有页面里插入的图片',
        '删除备注': '清空演讲者备注里的内容',
        '提取备注': '把每页备注导出成一个 TXT 文本',
        '删除宏': '移除演示文稿里的宏组件',
        '删除页眉页脚': '删除页面上的页眉/页脚/日期/编号占位符',
        '添加水印': '在每页加一个文字水印框',
        '设置背景': '给每一页统一填充指定颜色背景',
        '重置背景': '恢复每页为模板默认背景',
        '按页拆分': '把幻灯片按每 N 页拆成多个文件',
        '提取页码范围': '只保留指定页码范围的页面另存',
        '提取图片': '把所有页面里的图片导出成图片文件',
        '清空元数据': '抹掉作者、标题等属性，保护隐私',
        '修改元数据': '修改作者/标题等属性',
        '压缩': '重新压缩文件，体积变小，内容不变',
    }
    desc = ('PPT 批量处理：内容替换、图片/备注/宏清理、页眉页脚、背景、'
            '拆分合并、元数据、压缩')
    OPS = ('查找替换', '删除空白页', '删除图片', '删除备注', '提取备注',
           '删除宏', '删除页眉页脚', '添加水印', '设置背景', '重置背景',
           '按页拆分', '提取页码范围', '提取图片', '清空元数据', '修改元数据',
           '压缩')
    REQUIRED_FIELDS = {'查找替换': (('find', '查找内容'),),
                       '添加水印': (('wm', '水印文字'),)}
    OP_FIELDS = {
        '查找替换': ('find', 'replace'),
        '添加水印': ('wm',),
        '设置背景': ('bg_color',),
        '按页拆分': ('p_chunk',),
        '提取页码范围': ('p_start', 'p_end'),
        '修改元数据': ('meta_author',),
    }

    def build_options(self, form):
        self._field(form, 'op', '操作', _combo(self.OPS))
        self._field(form, 'find', '查找', QLineEdit())
        self._field(form, 'replace', '替换为', QLineEdit())
        self._field(form, 'wm', '水印文字', QLineEdit())
        bgc = QLineEdit()
        bgc.setPlaceholderText('#RRGGBB')
        self._field(form, 'bg_color', '背景色', bgc)
        ps = QSpinBox()
        ps.setRange(1, 100000)
        ps.setValue(1)
        self._field(form, 'p_start', '从', ps)
        pe = QSpinBox()
        pe.setRange(0, 100000)
        pe.setValue(0)
        pe.setToolTip('到(0=末尾)')
        self._field(form, 'p_end', '页码范围', pe)
        ch = QSpinBox()
        ch.setRange(1, 1000)
        ch.setValue(1)
        self._field(form, 'p_chunk', '每份页数(拆分)', ch)
        self._field(form, 'meta_author', '作者', QLineEdit())
        od = QLineEdit()
        od.setPlaceholderText('留空=就地修改(自动保存到“备份”文件夹)')
        self._field(form, 'out_dir', '输出目录', od)
        ob = QPushButton('输出到…')
        ob.clicked.connect(self.pick_out)
        form.addRow('', ob)
        self.ext_edit.setText('.pptx,.pptm')

    def pick_out(self):
        d = QFileDialog.getExistingDirectory(self, '选择输出目录')
        if d:
            self._widgets['out_dir'].setText(d)

    def collect_options(self):
        op = self._widgets['op'].currentText()
        key_map = {'查找替换': 'replace', '删除空白页': 'del_blank',
                   '删除图片': 'rmimg', '删除备注': 'nonotes',
                   '提取备注': 'extnotes', '删除宏': 'nomacro',
                   '删除页眉页脚': 'hf_del', '添加水印': 'wm',
                   '设置背景': 'bg', '重置背景': 'bg_r',
                   '按页拆分': 'split', '提取页码范围': 'ext_pages',
                   '提取图片': 'extimg', '清空元数据': 'meta_c',
                   '修改元数据': 'meta_m', '压缩': 'compress'}
        o = {'op': key_map.get(op, ''), 'op_name': op}
        if op == '查找替换':
            o['find'] = self._widgets['find'].text()
            o['replace'] = self._widgets['replace'].text()
        elif op == '添加水印':
            o['wm_text'] = self._widgets['wm'].text() or '青雁水印'
        elif op == '设置背景':
            o['bg_color'] = self._widgets['bg_color'].text()
        elif op == '按页拆分':
            o['ppt_chunk'] = self._widgets['p_chunk'].value()
        elif op == '提取页码范围':
            # 键名必须与引擎一致(advanced_ops.pptx_extract_pages 读
            # page_from/page_to)。早期这里发的是 page_start/page_end,
            # 引擎取不到就默认整份, 用户选 2-3 页拿到的是全部页。
            o['page_from'] = self._widgets['p_start'].value()
            o['page_to'] = self._widgets['p_end'].value()
        elif op == '修改元数据':
            o['meta_mode'] = '修改'
            o['meta_author'] = self._widgets['meta_author'].text()
        elif op == '清空元数据':
            o['meta_mode'] = '清空'
        if 'out_dir' in self._widgets:
            o['out_dir'] = self._widgets['out_dir'].text().strip()
        return o

    def processor(self):
        key = self.collect_options().get('op')
        return {
            'replace': office_ops.pptx_replace,
            'del_blank': office_ops.pptx_delete_blank_slides,
            'rmimg': advanced_ops.pptx_remove_images,
            'nonotes': advanced_ops.pptx_remove_notes,
            'extnotes': advanced_ops.pptx_extract_notes,
            'nomacro': advanced_ops.pptx_remove_macros,
            'hf_del': advanced_ops.pptx_remove_header_footer,
            'wm': office_ops.pptx_watermark,
            'bg': advanced_ops.pptx_set_background,
            'bg_r': lambda p, o: advanced_ops.pptx_set_background(
                p, {**o, 'bg_remove': True}),
            'split': advanced_ops.pptx_split,
            'ext_pages': advanced_ops.pptx_extract_slides,
            'extimg': advanced_ops.pptx_extract_images,
            'meta_c': advanced_ops.pptx_meta, 'meta_m': advanced_ops.pptx_meta,
            'compress': advanced_ops.pptx_compress,
        }.get(key)


class PdfPage(FunctionPage):
    title = 'PDF 工具'
    kind = 'pdf'
    EXPECTED_EXTS = ('.pdf',)
    OP_HINTS = {
        '每N页拆分': '把 PDF 按每 N 页拆成多个小文件',
        '提取页码范围': '只保留指定页码范围的页面另存新文件',
        '删除指定页': '删除指定页码（如 1,3,5-7），其余保留',
        '删除空白页': '删除没有文字、图片、印章的空白页面',
        '批量旋转': '把所有页面旋转指定角度（顺时针）',
        '插入空白页': '在指定页码前插入一页空白页',
        '页面裁剪': '按四边留白裁掉页面边缘（单位 pt）',
        '统一页面尺寸(A4/A3)': '把所有页面缩放对齐到同一尺寸',
        '删除页眉页脚(裁边)': '按高度裁掉顶部/底部条带，去除页眉页脚',
        '添加页眉页脚': '给每页顶部/底部居中加文字',
        '添加水印': '给每页平铺半透明文字水印',
        '添加骑缝章': '把印章图片竖切分贴到每页右边缘（打印后可拼合）',
        '加密': '给 PDF 设置打开口令',
        '解密': '去掉 PDF 的打开口令',
        '全页链接': '给每一页整面加一个网址链接',
        '移除链接': '删除 PDF 里的全部链接注释',
        '提取图片': '把 PDF 里的内嵌图片导出成图片文件',
        '清空元数据': '抹掉作者、标题等属性，保护隐私',
        '修改元数据': '修改作者/标题等属性',
        '压缩': '清理冗余对象，体积变小，内容不变',
        '生成空白 PDF': '生成指定页数的空白 PDF（用于打印草稿）',
        'PDF 页数统计': '统计文件夹里所有 PDF 的页数并导出 Excel',
    }
    desc = ('PDF 批量处理：拆分合并、页面处理、水印印章、页眉页脚、加密、'
            '元数据、压缩、图片提取')
    OPS = ('每N页拆分', '提取页码范围', '删除指定页', '删除空白页', '批量旋转',
           '插入空白页', '页面裁剪', '统一页面尺寸(A4/A3)', '删除页眉页脚(裁边)',
           '添加页眉页脚', '添加水印', '添加骑缝章', '加密', '解密', '全页链接',
           '移除链接', '提取图片', '清空元数据', '修改元数据', '压缩',
           '生成空白 PDF', 'PDF 页数统计')
    REQUIRED_FIELDS = {'加密': (('pwd', '加密口令'),),
                       '解密': (('pwd', '解密口令'),),
                       '添加骑缝章': (('seal_edit', '印章图片'),),
                       '全页链接': (('link_url', '链接网址'),),
                       '删除指定页': (('pages_spec', '删除页码'),)}
    OP_FIELDS = {
        '每N页拆分': ('chunk',),
        '生成空白 PDF': ('chunk',),
        '提取页码范围': ('p_from', 'p_to'),
        '删除指定页': ('pages_spec',),
        '批量旋转': ('angle',),
        '插入空白页': ('insert_at',),
        '页面裁剪': ('crop_margin',),
        '统一页面尺寸(A4/A3)': ('page_size',),
        '删除页眉页脚(裁边)': ('hf_top', 'hf_bottom'),
        '添加页眉页脚': ('header', 'footer'),
        '添加水印': ('wm', 'wm_opacity'),
        '添加骑缝章': ('seal_edit', 'seal_btn'),
        '加密': ('pwd',),
        '解密': ('pwd',),
        '全页链接': ('link_url',),
        '修改元数据': ('meta_author',),
    }

    def build_options(self, form):
        self._field(form, 'op', '操作', _combo(self.OPS))
        ch = QSpinBox()
        ch.setRange(1, 10000)
        ch.setValue(1)
        ch.setToolTip('拆分时=每份页数；生成空白 PDF 时=总页数')
        self._field(form, 'chunk', '每份页数(拆分)', ch)
        pf = QSpinBox()
        pf.setRange(1, 100000)
        pf.setValue(1)
        self._field(form, 'p_from', '从', pf)
        pt = QSpinBox()
        pt.setRange(0, 100000)
        pt.setValue(0)
        pt.setToolTip('到(0=末尾)')
        self._field(form, 'p_to', '页码范围', pt)
        pspec = QLineEdit()
        pspec.setPlaceholderText('删除页码, 如 1,3,5-7')
        self._field(form, 'pages_spec', '删除页码', pspec)
        self._field(form, 'angle', '旋转角度', _combo(('90', '180', '270')))
        ia = QSpinBox()
        ia.setRange(1, 100000)
        ia.setValue(1)
        ia.setToolTip('空白页插入位置(第N页前)')
        self._field(form, 'insert_at', '插入位置', ia)
        cm = QLineEdit()
        cm.setPlaceholderText('左,上,右,下 (pt), 如 20,30,20,30')
        self._field(form, 'crop_margin', '裁剪边距', cm)
        self._field(form, 'page_size', '目标尺寸',
                    _combo(('A4', 'A3', 'A5', 'Letter')))
        ht = QDoubleSpinBox()
        ht.setRange(0, 200)
        ht.setValue(30)
        self._field(form, 'hf_top', '页眉高', ht)
        hb = QDoubleSpinBox()
        hb.setRange(0, 200)
        hb.setValue(30)
        self._field(form, 'hf_bottom', '页脚高', hb)
        hb.setToolTip('裁边(pt)')
        self._field(form, 'header', '页眉', QLineEdit())
        self._field(form, 'footer', '页脚', QLineEdit())
        self._field(form, 'wm', '水印文字', QLineEdit())
        wo = QSpinBox()
        wo.setRange(3, 100)
        wo.setValue(15)
        self._field(form, 'wm_opacity', '水印透明度%', wo)
        se = QLineEdit()
        se.setPlaceholderText('印章图片路径(骑缝章)')
        self._field(form, 'seal_edit', '印章', se)
        sb = QPushButton('选印章')
        sb.clicked.connect(self.pick_seal)
        self._field(form, 'seal_btn', '', sb)
        lu = QLineEdit()
        lu.setPlaceholderText('https://…')
        self._field(form, 'link_url', '链接网址', lu)
        pwd = QLineEdit()
        pwd.setEchoMode(QLineEdit.Password)
        pwd.setToolTip('口令(加密/解密)')
        self._field(form, 'pwd', '口令', pwd)
        self._field(form, 'meta_author', '作者', QLineEdit())
        od = QLineEdit()
        od.setPlaceholderText('输出目录(留空=源文件旁)')
        self._field(form, 'out_dir', '输出目录', od)
        ob = QPushButton('输出到…')
        ob.clicked.connect(self.pick_out)
        form.addRow('', ob)
        self.ext_edit.setText('.pdf')

    def pick_seal(self):
        f, _ = QFileDialog.getOpenFileName(self, '选择印章图片', '',
                                           '图片 (*.png *.jpg *.jpeg *.bmp)')
        if f:
            self._widgets['seal_edit'].setText(f)

    def pick_out(self):
        d = QFileDialog.getExistingDirectory(self, '选择输出目录')
        if d:
            self._widgets['out_dir'].setText(d)

    def collect_options(self):
        op = self._widgets['op'].currentText()
        key_map = {'每N页拆分': 'split', '提取页码范围': 'extract',
                   '删除指定页': 'del_pages', '删除空白页': 'del_blank',
                   '批量旋转': 'rotate', '插入空白页': 'insert_blank',
                   '页面裁剪': 'crop', '统一页面尺寸(A4/A3)': 'resize',
                   '删除页眉页脚(裁边)': 'hf_del', '添加页眉页脚': 'hf',
                   '添加水印': 'wm', '添加骑缝章': 'seal', '加密': 'encrypt',
                   '解密': 'decrypt', '全页链接': 'link', '移除链接': 'nolink',
                   '提取图片': 'extimg', '清空元数据': 'meta_c',
                   '修改元数据': 'meta_m', '压缩': 'compress',
                   '生成空白 PDF': 'blank', 'PDF 页数统计': 'pdfcount'}
        o = {'op': op, 'pdf_op': key_map.get(op, ''), 'op_name': op}
        if op == '每N页拆分':
            o['chunk_size'] = self._widgets['chunk'].value()
        elif op == '提取页码范围':
            o['page_from'] = self._widgets['p_from'].value()
            o['page_to'] = self._widgets['p_to'].value()
        elif op == '删除指定页':
            o['pages_spec'] = self._widgets['pages_spec'].text()
        elif op == '批量旋转':
            o['angle'] = int(self._widgets['angle'].currentText() or 90)
        elif op == '插入空白页':
            o['insert_at'] = self._widgets['insert_at'].value()
        elif op == '页面裁剪':
            o['crop_margin'] = self._widgets['crop_margin'].text()
        elif op == '统一页面尺寸(A4/A3)':
            o['page_size'] = self._widgets['page_size'].currentText()
        elif op == '删除页眉页脚(裁边)':
            o['hf_top'] = self._widgets['hf_top'].value()
            o['hf_bottom'] = self._widgets['hf_bottom'].value()
        elif op == '添加页眉页脚':
            o['header_text'] = self._widgets['header'].text()
            o['footer_text'] = self._widgets['footer'].text()
        elif op == '添加水印':
            o['wm_text'] = self._widgets['wm'].text() or '青雁水印'
            o['wm_opacity'] = self._widgets['wm_opacity'].value()
            o['wm_tile'] = True
        elif op == '添加骑缝章':
            o['seal_image'] = self._widgets['seal_edit'].text()
        elif op == '全页链接':
            o['link_url'] = self._widgets['link_url'].text()
        elif op in ('加密', '解密'):
            o['password'] = self._widgets['pwd'].text()
        elif op == '生成空白 PDF':
            o['pdf_pages'] = self._widgets['chunk'].value()
        elif op == '修改元数据':
            o['meta_mode'] = '修改'
            o['meta_author'] = self._widgets['meta_author'].text()
        elif op == '清空元数据':
            o['meta_mode'] = '清空'
        if 'out_dir' in self._widgets:
            o['out_dir'] = self._widgets['out_dir'].text().strip()
        return o

    def processor(self):
        o = self.collect_options()
        key = o.get('pdf_op')
        return {
            'split': office_ops.pdf_ops, 'extract': office_ops.pdf_ops,
            'rotate': office_ops.pdf_ops, 'insert_blank': office_ops.pdf_ops,
            'encrypt': office_ops.pdf_ops, 'decrypt': office_ops.pdf_ops,
            'hf': office_ops.pdf_header_footer, 'wm': office_ops.pdf_watermark,
            'del_pages': advanced_ops.pdf_delete_pages,
            'del_blank': advanced_ops.pdf_delete_blank_pages,
            'crop': advanced_ops.pdf_crop, 'resize': advanced_ops.pdf_resize,
            'hf_del': advanced_ops.pdf_remove_header_footer,
            'seal': advanced_ops.pdf_add_seal, 'link': advanced_ops.pdf_add_link,
            'nolink': advanced_ops.pdf_remove_links,
            'extimg': advanced_ops.pdf_extract_images,
            'meta_c': advanced_ops.pdf_meta, 'meta_m': advanced_ops.pdf_meta,
            'compress': advanced_ops.pdf_compress,
            'blank': lambda p, op: ops.pdf_blank_and_count(
                p, {**op, 'mode': 'blank'}),
            'pdfcount': lambda p, op: ops.pdf_blank_and_count(
                p, {**op, 'mode': 'count'}),
        }.get(key)


class TextPage(FunctionPage):
    title = '文本工具'
    kind = 'text'
    OP_HINTS = {
        '查找替换': '把文本里的指定内容替换成新内容，支持正则表达式',
        'Excel 规则批量替换': '按 Excel 里两列「查找|替换」的规则一次替换多条',
        '删除空行': '删除文本中的空白行',
        '删除重复行': '内容重复的行只保留第一次出现',
        '清理行尾空格': '去掉每行末尾多余的空格',
        '删除或替换指定行': '按行号（1,3,5-7）或内容匹配删除/替换指定行',
        '指定位置插入行': '在第 N 行前插入一行或多行内容',
        '换行符转换': '在 Windows(CRLF) 和 Linux/Mac(LF) 换行符之间转换',
        'BOM 添加/移除': '添加或移除 UTF-8 BOM 头（部分老软件需要）',
        '编码转换': '在 UTF-8 / GBK / GB18030 / UTF-16 之间转换编码',
        '按行数拆分': '把大文本按每 N 行拆成多个文件',
        'TXT→HTML': '把纯文本转成网页（每段一个 <p>）',
        'Markdown→HTML': '把 Markdown 转成带标题/列表/加粗的网页',
        'Markdown→Word': '把 Markdown 转成 Word 文档（标题/列表/加粗）',
        'HTML→TXT': '抽取网页里的纯文字',
        'HTML→Word': '把网页文字转成 Word 文档',
        'JSON 格式化': '把压缩的 JSON 排版成易读的缩进格式',
        'JSON→CSV': '把 JSON 数组转成 Excel 可打开的 CSV 表格',
        'JSON→TXT': '把 JSON 转成「键 Tab 值」的文本',
        '全角转半角': '把全角字母数字标点转成半角（排版清理常用）',
        '半角转全角': '把半角字母数字标点转成全角',
        '中文标点统一': '把中文语境里的英文标点替换成中文标点',
    }
    desc = ('文本批量处理：替换(正则/Excel规则)、行操作、编码、格式转换、拆分合并')
    OPS = ('查找替换', 'Excel 规则批量替换', '删除空行', '删除重复行',
           '清理行尾空格', '删除或替换指定行', '指定位置插入行', '换行符转换',
           'BOM 添加/移除', '编码转换', '按行数拆分', 'TXT→HTML',
           'Markdown→HTML', 'Markdown→Word', 'HTML→TXT', 'HTML→Word',
           'JSON 格式化', 'JSON→CSV', 'JSON→TXT', '全角转半角', '半角转全角',
           '中文标点统一')
    REQUIRED_FIELDS = {'查找替换': (('find', '查找内容'),),
                       'Excel 规则批量替换': (('rules_edit', '规则 Excel'),)}
    OP_FIELDS = {
        '查找替换': ('find', 'replace', 'regex_cb'),
        'Excel 规则批量替换': ('rules_edit',),
        '删除或替换指定行': ('lines_spec', 'line_match', 'line_new',
                            'line_mode'),
        '指定位置插入行': ('insert_pos', 'insert_text'),
        '换行符转换': ('eol',),
        'BOM 添加/移除': ('bom',),
        '编码转换': ('enc_to',),
        '按行数拆分': ('lines_per',),
    }

    def build_options(self, form):
        self._field(form, 'op', '操作', _combo(self.OPS))
        self._field(form, 'find', '查找', QLineEdit())
        self._field(form, 'replace', '替换为', QLineEdit())
        self._field(form, 'regex_cb', '', QCheckBox('使用正则表达式'))
        re_ = QLineEdit()
        re_.setPlaceholderText('规则 Excel(xlsx): 两列 查找|替换')
        self._field(form, 'rules_edit', 'Excel 规则', re_)
        rb = QPushButton('选规则表')
        rb.clicked.connect(self.pick_rules)
        form.addRow('', rb)
        ls = QLineEdit()
        ls.setPlaceholderText('行号, 如 1,3,5-7 或留空按内容匹配')
        self._field(form, 'lines_spec', '行号', ls)
        self._field(form, 'line_match', '行匹配内容', QLineEdit())
        self._field(form, 'line_new', '行替换为', QLineEdit())
        self._field(form, 'line_mode', '行操作', _combo(('删除', '替换')))
        ip = QSpinBox()
        ip.setRange(1, 1000000)
        ip.setValue(1)
        self._field(form, 'insert_pos', '插入位置(第N行前)', ip)
        it = QPlainTextEdit()
        it.setMaximumHeight(70)
        it.setPlaceholderText('要插入的内容(可多行)')
        self._field(form, 'insert_text', '插入内容', it)
        self._field(form, 'eol', '目标换行符', _combo(('LF (\\n)', 'CRLF (\\r\\n)')))
        self._field(form, 'bom', 'BOM', _combo(('添加 BOM', '移除 BOM')))
        self._field(form, 'enc_to', '目标编码',
                    _combo(('UTF-8', 'GBK', 'GB18030', 'UTF-16')))
        lp = QSpinBox()
        lp.setRange(1, 1000000)
        lp.setValue(1000)
        self._field(form, 'lines_per', '每份行数(拆分)', lp)
        od = QLineEdit()
        od.setPlaceholderText('留空=就地修改(自动保存到“备份”文件夹)')
        self._field(form, 'out_dir', '输出目录', od)
        ob = QPushButton('输出到…')
        ob.clicked.connect(self.pick_out)
        form.addRow('', ob)
        self.ext_edit.setText(
            '.txt,.md,.csv,.log,.json,.xml,.html,.htm,.srt,.ini,.yaml,.yml')

    def pick_rules(self):
        f, _ = QFileDialog.getOpenFileName(self, '选择规则 Excel', '',
                                           'Excel (*.xlsx)')
        if f:
            self._widgets['rules_edit'].setText(f)

    def pick_out(self):
        d = QFileDialog.getExistingDirectory(self, '选择输出目录')
        if d:
            self._widgets['out_dir'].setText(d)

    def collect_options(self):
        op = self._widgets['op'].currentText()
        key_map = {'查找替换': 'replace', 'Excel 规则批量替换': 'rules_replace',
                   '删除空行': 'del_blank', '删除重复行': 'del_dup',
                   '清理行尾空格': 'trim', '删除或替换指定行': 'line_del',
                   '指定位置插入行': 'line_ins', '换行符转换': 'crlf',
                   'BOM 添加/移除': 'bom', '编码转换': 'encoding',
                   '按行数拆分': 'split_lines', 'TXT→HTML': 'txt2html',
                   'Markdown→HTML': 'md2html', 'Markdown→Word': 'md2docx',
                   'HTML→TXT': 'html2txt', 'HTML→Word': 'html2docx',
                   'JSON 格式化': 'json_fmt', 'JSON→CSV': 'json_csv',
                   'JSON→TXT': 'json_txt', '全角转半角': 'fw_half',
                   '半角转全角': 'fw_full', '中文标点统一': 'fw_punct'}
        o = {'op': key_map.get(op, ''), 'op_name': op}
        if op == '查找替换':
            o['find'] = self._widgets['find'].text()
            o['replace'] = self._widgets['replace'].text()
            o['regex'] = self._widgets['regex_cb'].isChecked()
        elif op == 'Excel 规则批量替换':
            o['rules_file'] = self._widgets['rules_edit'].text()
        elif op == '删除或替换指定行':
            o['lines_spec'] = self._widgets['lines_spec'].text()
            o['line_match'] = self._widgets['line_match'].text()
            o['line_new'] = self._widgets['line_new'].text()
            o['line_mode'] = self._widgets['line_mode'].currentText()
        elif op == '指定位置插入行':
            o['insert_pos'] = self._widgets['insert_pos'].value()
            o['insert_text'] = self._widgets['insert_text'].toPlainText()
        elif op == '换行符转换':
            o['eol'] = 'lf' if self._widgets['eol'].currentText().startswith('LF') \
                else 'crlf'
        elif op == 'BOM 添加/移除':
            o['bom_to'] = 'add' if '添加' in self._widgets['bom'].currentText() \
                else 'remove'
        elif op == '编码转换':
            o['enc_to'] = self._widgets['enc_to'].currentText()
        elif op == '按行数拆分':
            o['lines_per'] = self._widgets['lines_per'].value()
        elif op in ('JSON→CSV', 'JSON→TXT'):
            # 引擎靠 json_target 决定输出格式, 默认是 txt。不传就等于
            # 用户点"JSON→CSV"却拿到 .txt。
            o['json_target'] = 'csv' if op == 'JSON→CSV' else 'txt'
        if 'out_dir' in self._widgets:
            o['out_dir'] = self._widgets['out_dir'].text().strip()
        return o

    def supports_preview(self):
        return self.collect_options().get('op') in ('replace', 'rules_replace',
                                                    'line_del')

    def build_preview(self, files, options):
        op = options.get('op')
        if op not in ('replace', 'rules_replace', 'line_del'):
            return None
        rows = [('文件', '命中数', '首个命中上下文')]
        from engine.ops import detect_encoding, load_excel_map
        mapping = None
        if op == 'rules_replace':
            rf = options.get('rules_file', '')
            if not rf or not os.path.exists(rf):
                return ('文件', '错误', ''), [('—', '规则表不存在', '')]
            mapping = load_excel_map(rf)
        for f in files[:100]:
            try:
                enc = detect_encoding(f)
                with open(f, 'r', encoding=enc, errors='replace') as fh:
                    content = fh.read()
            except Exception:
                rows.append((os.path.basename(f), '读取失败', ''))
                continue
            n, ctx = 0, ''
            if op == 'replace':
                find = options.get('find', '')
                if options.get('regex'):
                    # 勾了正则就得按正则数, 否则 \d+ 这种模式预览恒 0 命中,
                    # 用户以为没匹配就放弃了, 实际执行却能替换。
                    import re as _re
                    try:
                        pat = _re.compile(find) if find else None
                    except _re.error:
                        pat = None
                    if pat is not None:
                        hits = list(pat.finditer(content))
                        n = len(hits)
                        if hits:
                            m = hits[0]
                            start = max(0, m.start() - 15)
                            frag = content[start:m.end() + 15].replace('\n', ' ')
                            ctx = (content[start:m.start()] + '【' +
                                   m.group(0) + '】' + content[m.end():m.end() + 15]
                                   ).replace('\n', ' ')
                    else:
                        n = 0
                        ctx = '正则表达式无效'
                else:
                    n = content.count(find) if find else 0
                    if n:
                        idx = content.find(find)
                        start = max(0, idx - 15)
                        frag = content[start:idx + len(find) + 15].replace('\n', ' ')
                        ctx = frag.replace(find, '【' + find + '】')
            elif op == 'rules_replace':
                n = sum(content.count(k) for k in mapping)
                first_k = next((k for k in mapping if k in content), None)
                if first_k is not None:
                    idx = content.find(first_k)
                    start = max(0, idx - 15)
                    frag = content[start:idx + len(first_k) + 15].replace('\n', ' ')
                    ctx = frag.replace(first_k, '【' + first_k + '】')
            else:
                spec = options.get('lines_spec', '')
                match = options.get('line_match', '')
                lines = content.splitlines()
                for i, ln in enumerate(lines, 1):
                    hit = False
                    if spec:
                        for part in spec.split(','):
                            part = part.strip()
                            if '-' in part:
                                a, _, b = part.partition('-')
                                if a.strip().isdigit() and b.strip().isdigit() \
                                        and int(a) <= i <= int(b):
                                    hit = True
                            elif part.isdigit() and int(part) == i:
                                hit = True
                    elif match and match in ln:
                        hit = True
                    if hit:
                        n += 1
                        if n == 1:
                            ctx = f'第 {i} 行: ' + ln[:60]
            rows.append((os.path.basename(f), str(n), ctx))
        return rows

    def processor(self):
        key = self.collect_options().get('op')
        return {
            'replace': ops.process_text, 'del_blank': ops.process_text,
            'del_dup': ops.process_text, 'trim': ops.process_text,
            'crlf': ops.process_text, 'bom': ops.process_text,
            'encoding': ops.process_text, 'split_lines': ops.process_text,
            'txt2html': ops.process_text, 'md2html': ops.process_text,
            'html2txt': ops.process_text, 'json_fmt': ops.process_text,
            'rules_replace': advanced_ops.text_replace_by_excel,
            'line_del': advanced_ops.text_delete_replace_lines,
            'line_ins': advanced_ops.text_insert_lines,
            'md2docx': advanced_ops.text_md_to_docx,
            'html2docx': advanced_ops.text_html_to_docx,
            'json_csv': advanced_ops.text_json_convert,
            'json_txt': advanced_ops.text_json_convert,
            'fw_half': ops.text_fullwidth_convert,
            'fw_full': ops.text_fullwidth_convert,
            'fw_punct': ops.text_fullwidth_convert,
        }.get(key)


class ImagePage(FunctionPage):
    title = '图片工具'
    kind = 'image'
    # 擦除区域按像素定位, 动图只能处理首帧, 引擎明确不支持 gif,
    # 这里同步去掉声明, 免得放行后又被引擎判"非支持图片"。
    EXPECTED_EXTS = ('.jpg', '.jpeg', '.png', '.bmp', '.webp', '.tif', '.tiff')
    OP_HINTS = {
        '格式转换': '把图片转成 PNG / JPG / WEBP 等格式',
        '压缩': '减小图片体积（可同时限制最大宽度，按比例缩放）',
        '调整尺寸': '按宽或高等比缩放；两边都填且不锁比例则强制拉伸',
        '添加水印': '在图片指定位置印上文字（可调透明度和字号）',
        '旋转': '把图片旋转指定角度',
        '翻转': '把图片水平或垂直镜像翻转',
        '裁剪': '按「左,上,右,下」像素裁掉图片边缘',
        '颜色模式': '转成灰度 / 黑白 / RGBA 等颜色模式',
        '网格分割': '按「行,列」把一张图切成 N 块小图（九宫格等）',
        '擦除区域': '把指定矩形区域用纯色盖掉或打成马赛克（去水印/隐私）',
        '清空元数据': '抹掉 EXIF 拍摄信息（机型/GPS 位置等隐私）',
    }
    desc = '图片批量处理：压缩/尺寸/裁剪/旋转/水印/格式/颜色/分割/擦除/元数据'
    OPS = ('格式转换', '压缩', '调整尺寸', '添加水印', '旋转', '翻转', '裁剪',
           '颜色模式', '网格分割', '擦除区域', '清空元数据')
    REQUIRED_FIELDS = {'擦除区域': (('erase_rect', '擦除区域坐标'),)}
    OP_FIELDS = {
        '格式转换': ('fmt', 'quality'),
        '压缩': ('quality', 'max_width'),
        '调整尺寸': ('width', 'height', 'keep_ratio'),
        '添加水印': ('wm_text', 'wm_pos', 'wm_opacity', 'wm_size'),
        '旋转': ('angle',),
        '翻转': ('flip',),
        '裁剪': ('crop',),
        '颜色模式': ('color_mode',),
        '网格分割': ('grid',),
        '擦除区域': ('erase_rect', 'erase_mode', 'erase_color'),
    }

    def build_options(self, form):
        self._field(form, 'op', '操作', _combo(self.OPS))
        self._field(form, 'fmt', '目标格式',
                    _combo(('PNG', 'JPG', 'WEBP', 'BMP', 'TIFF')))
        q = QSpinBox()
        q.setRange(10, 100)
        q.setValue(85)
        self._field(form, 'quality', '质量(压缩)', q)
        mw = QSpinBox()
        mw.setRange(0, 20000)
        mw.setValue(0)
        mw.setToolTip('最大宽度, 0=不限')
        self._field(form, 'max_width', '最大宽度', mw)
        w = QSpinBox()
        w.setRange(0, 20000)
        self._field(form, 'width', '宽', w)
        h = QSpinBox()
        h.setRange(0, 20000)
        h.setToolTip('(0=按比例)')
        self._field(form, 'height', '高', h)
        self._field(form, 'keep_ratio', '', QCheckBox('锁定比例'))
        self._widgets['keep_ratio'].setChecked(True)
        self._field(form, 'angle', '角度(旋转)',
                    _combo(('-90', '-180', '90', '180', '270')))
        self._field(form, 'flip', '翻转方向', _combo(('水平', '垂直')))
        cr = QLineEdit()
        cr.setPlaceholderText('左,上,右,下 像素')
        self._field(form, 'crop', '裁剪边距', cr)
        self._field(form, 'color_mode', '颜色模式', _combo(('灰度', '黑白', 'RGBA')))
        g = QLineEdit()
        g.setPlaceholderText('行,列 如 3,3')
        self._field(form, 'grid', '分割网格', g)
        er = QLineEdit()
        er.setPlaceholderText('x,y,宽,高 像素')
        self._field(form, 'erase_rect', '擦除区域', er)
        self._field(form, 'erase_mode', '擦除方式', _combo(('纯色', '马赛克')))
        ec = QLineEdit()
        ec.setPlaceholderText('#FFFFFF')
        self._field(form, 'erase_color', '擦除颜色', ec)
        self._field(form, 'wm_text', '水印文字', QLineEdit())
        self._field(form, 'wm_pos', '水印位置',
                    _combo(('右下', '中下', '左下', '右中', '居中', '左中',
                            '右上', '中上', '左上')))
        wo = QSpinBox()
        wo.setRange(5, 100)
        wo.setValue(60)
        self._field(form, 'wm_opacity', '透明%', wo)
        ws_ = QSpinBox()
        ws_.setRange(8, 500)
        ws_.setValue(24)
        self._field(form, 'wm_size', '水印字号', ws_)
        od = QLineEdit()
        od.setPlaceholderText('留空=就地输出(原文件保留)')
        self._field(form, 'out_dir', '输出目录', od)
        ob = QPushButton('输出到…')
        ob.clicked.connect(self.pick_out)
        form.addRow('', ob)
        self.ext_edit.setText('.jpg,.jpeg,.png,.bmp,.webp,.tif,.tiff,.gif')

    def pick_out(self):
        d = QFileDialog.getExistingDirectory(self, '选择输出目录')
        if d:
            self._widgets['out_dir'].setText(d)

    def collect_options(self):
        op = self._widgets['op'].currentText()
        op_map = {'格式转换': 'convert', '压缩': 'compress',
                  '调整尺寸': 'resize', '添加水印': 'watermark',
                  '旋转': 'rotate', '翻转': 'flip', '裁剪': 'crop',
                  '颜色模式': 'color', '网格分割': 'split',
                  '擦除区域': 'erase', '清空元数据': 'strip'}
        pos_map = {'右下': 'br', '中下': 'bc', '左下': 'bl', '右中': 'cr',
                   '居中': 'cc', '左中': 'cl', '右上': 'tr', '中上': 'tc',
                   '左上': 'tl'}
        o = {'op': op_map.get(op, ''), 'op_name': op}
        if op == '格式转换':
            o['format'] = self._widgets['fmt'].currentText()
            o['quality'] = self._widgets['quality'].value()
        elif op == '压缩':
            o['quality'] = self._widgets['quality'].value()
            o['max_width'] = self._widgets['max_width'].value()
        elif op == '调整尺寸':
            o['width'] = self._widgets['width'].value()
            o['height'] = self._widgets['height'].value()
            o['keep_ratio'] = self._widgets['keep_ratio'].isChecked()
        elif op == '添加水印':
            o['wm_text'] = self._widgets['wm_text'].text() or '青雁'
            o['wm_pos'] = pos_map.get(self._widgets['wm_pos'].currentText(), 'br')
            o['wm_opacity'] = self._widgets['wm_opacity'].value()
            o['wm_size'] = self._widgets['wm_size'].value()
        elif op == '旋转':
            o['angle'] = self._widgets['angle'].currentText()
        elif op == '翻转':
            o['flip'] = self._widgets['flip'].currentText()
        elif op == '裁剪':
            o['crop'] = self._widgets['crop'].text()
        elif op == '颜色模式':
            o['color_mode'] = self._widgets['color_mode'].currentText()
        elif op == '网格分割':
            o['grid'] = self._widgets['grid'].text()
        elif op == '擦除区域':
            o['erase_rect'] = self._widgets['erase_rect'].text()
            o['erase_mode'] = self._widgets['erase_mode'].currentText()
            o['erase_color'] = self._widgets['erase_color'].text()
        if 'out_dir' in self._widgets:
            o['out_dir'] = self._widgets['out_dir'].text().strip()
        return o

    def processor(self):
        key = self.collect_options().get('op')
        if key == 'erase':
            return advanced_ops.image_erase
        return ops.process_image


class RenamePage(FunctionPage):
    title = '重命名工具'
    kind = 'rename'
    OP_HINTS = {
        '序号编排': '把文件按顺序改成 001、002…（可加前缀）',
        '查找替换': '替换文件名里的指定文字',
        '加前后缀': '在原文件名前后加固定文字',
        '加日期': '在文件名末尾追加今天的日期',
        '扩展名替换': '只改扩展名，如 .jpg 改 .png（文件内容不变）',
        'Excel 对照表改名(文件)': '按 Excel 两列「原名|新名」的对照关系改名',
        'Excel 对照表改名(文件夹)': '选择父目录，按 Excel 对照关系给其中的子文件夹改名',
        '文件夹序号编排': '选择父目录，把其中的子文件夹按顺序改成 001、002…',
    }
    desc = '批量重命名：序号/替换/前后缀/日期/扩展名/Excel 对照表，支持文件与文件夹'
    ALWAYS_VISIBLE = ()
    # 重命名页早期没设必填: 查找留空时 base.replace('', 'X') 会把单字符
    # 逐个插进文件名 (report.txt → XrXeXpXoXrXtX.txt), 其他页都有拦截。
    REQUIRED_FIELDS = {'查找替换': (('find', '查找内容'),)}
    OP_FIELDS = {
        '序号编排': ('seq_start', 'seq_pad', 'prefix'),
        '查找替换': ('find', 'replace'),
        '加前后缀': ('prefix', 'suffix'),
        '加日期': ('suffix',),
        '扩展名替换': ('new_ext',),
        'Excel 对照表改名(文件)': ('excel_edit',),
        'Excel 对照表改名(文件夹)': ('excel_edit',),
        '文件夹序号编排': ('seq_start', 'seq_pad', 'prefix'),
    }

    def build_options(self, form):
        self.modes = ('序号编排', '查找替换', '加前后缀', '加日期', '扩展名替换',
                      'Excel 对照表改名(文件)', 'Excel 对照表改名(文件夹)',
                      '文件夹序号编排')
        self._field(form, 'mode', '模式', _combo(self.modes))
        ss = QSpinBox()
        ss.setRange(0, 999999)
        ss.setValue(1)
        self._field(form, 'seq_start', '起始', ss)
        sp = QSpinBox()
        sp.setRange(1, 8)
        sp.setValue(3)
        self._field(form, 'seq_pad', '位数', sp)
        self._field(form, 'find', '查找', QLineEdit())
        self._field(form, 'replace', '替换为', QLineEdit())
        self._field(form, 'prefix', '前缀', QLineEdit())
        self._field(form, 'suffix', '后缀', QLineEdit())
        ne = QLineEdit()
        ne.setPlaceholderText('.jpg')
        self._field(form, 'new_ext', '新扩展名', ne)
        ee = QLineEdit()
        ee.setPlaceholderText('xlsx 两列: 原名 | 新名')
        self._field(form, 'excel_edit', 'Excel 对照', ee)
        eb = QPushButton('选 Excel')
        eb.clicked.connect(self.pick_excel)
        form.addRow('', eb)
        # 信号由基类 _setup_op_fields 统一连接, 此处不再重复连接
        # (重复连接会让一次切换触发多次 apply_op_fields)
        self.apply_op_fields()

    def pick_excel(self):
        f, _ = QFileDialog.getOpenFileName(self, '选择对照 Excel', '',
                                           'Excel (*.xlsx)')
        if f:
            self._widgets['excel_edit'].setText(f)

    def _collect_mode_options(self):
        mode = self._widgets['mode'].currentText()
        mode_map = {'序号编排': 'sequence', '查找替换': 'replace',
                    '加前后缀': 'prefix', '加日期': 'add_date',
                    '扩展名替换': 'ext', 'Excel 对照表改名(文件)': 'excel',
                    'Excel 对照表改名(文件夹)': 'excel_dir',
                    '文件夹序号编排': 'dir_seq'}
        o = {'mode': mode_map.get(mode, 'sequence'), 'op_name': mode}
        if mode in ('序号编排', '文件夹序号编排'):
            o['seq_start'] = self._widgets['seq_start'].value()
            o['seq_pad'] = self._widgets['seq_pad'].value()
            o['prefix'] = self._widgets['prefix'].text()
        elif mode == '查找替换':
            o['find'] = self._widgets['find'].text()
            o['replace'] = self._widgets['replace'].text()
        elif mode == '加前后缀':
            o['prefix'] = self._widgets['prefix'].text()
            o['suffix'] = self._widgets['suffix'].text()
        elif mode == '加日期':
            o['suffix'] = self._widgets['suffix'].text()
        elif mode == '扩展名替换':
            o['new_ext'] = self._widgets['new_ext'].text()
        elif mode.startswith('Excel 对照表'):
            o['excel_map'] = self._widgets['excel_edit'].text()
        return o

    def current_files(self):
        mode_widget = self._widgets.get('mode')
        mode = mode_widget.currentText() if mode_widget is not None else ''
        if mode not in ('Excel 对照表改名(文件夹)', '文件夹序号编排'):
            return super().current_files()
        raw = getattr(self, 'file_edit', None)
        paths = [os.path.abspath(p.strip()) for p in raw.text().split('||') if p.strip()] if raw else []
        if len(paths) == 1 and os.path.isdir(paths[0]):
            try:
                paths = [os.path.join(paths[0], name) for name in sorted(os.listdir(paths[0]))]
            except OSError:
                return []
        from engine.backups import is_backup_path
        return self._apply_filters([p for p in paths if os.path.isdir(p) and not is_backup_path(p)])

    def processor(self):
        opts = self.collect_options()
        if opts.get('mode') in ('excel', 'excel_dir'):
            return lambda p, o: ops.rename_by_excel_map(
                p, o.get('_excel_map') or {})
        if opts.get('mode') == 'dir_seq':
            return self._dir_seq_proc
        return ops.process_rename

    def _dir_seq_proc(self, path, options):
        if not os.path.isdir(path):
            return FileResult(path, 'skipped', '非文件夹')
        d, old = os.path.split(path)
        seq = options.get('seq_start', 1) + \
            getattr(self, '_dir_counter', 0)
        self._dir_counter = getattr(self, '_dir_counter', 0) + 1
        pad = int(options.get('seq_pad', 3) or 3)
        prefix = options.get('prefix', '')
        idx = seq
        from engine.safety import validate_filename
        new_name = validate_filename(prefix + str(idx).zfill(pad))
        new_path = os.path.join(os.path.abspath(d), new_name)
        if os.path.abspath(new_path) == os.path.abspath(path):
            return FileResult(path, 'skipped', '名称未变化')
        if os.path.exists(new_path):
            return FileResult(path, 'fail', '目标已存在: ' + new_name)
        os.rename(path, new_path)
        undo = [{'src': path, 'dst': new_path}]
        return FileResult(path, 'ok', '→ ' + new_name, new_path, undo=undo)

    def do_run(self):
        opts = self.collect_options()
        # 本页重写了 do_run, 绕过了基类的必填校验, 必须自己补上
        missing = self.validate_required()
        if missing:
            QMessageBox.warning(self, self.main.APP_NAME,
                                '请填写以下参数后再执行：' + '、'.join(missing))
            return
        self._dir_counter = 0
        if opts.get('mode') in ('excel', 'excel_dir'):
            xf = opts.get('excel_map', '')
            if not xf or not os.path.exists(xf):
                QMessageBox.information(self, self.main.APP_NAME,
                                        '请先选择 Excel 对照表')
                return
            try:
                self._excel_map_cache = ops.load_excel_map(xf)
            except Exception as e:
                QMessageBox.warning(self, self.main.APP_NAME,
                                    '读取对照表失败: ' + str(e))
                return
        if opts.get('mode') in ('dir_seq', 'excel_dir'):
            files = self.current_files()
            dirs = [f for f in files if os.path.isdir(f)] \
                if files else []
            if not dirs:
                QMessageBox.information(self, self.main.APP_NAME,
                                        '请先选择包含子文件夹的目录')
                return
            files = dirs
            plan = []
            for i, f in enumerate(files):
                d, old = os.path.split(f)
                idx = opts.get('seq_start', 1) + i
                pad = int(opts.get('seq_pad', 3) or 3)
                new_name = opts.get('prefix', '') + str(idx).zfill(pad)
                new_path = os.path.join(d, new_name)
                plan.append((f, new_path, new_name))
            self._plan_cache = plan
            files = [p[0] for p in plan]
            self._override_files = files
        else:
            self._override_files = None
        if opts.get('mode') not in ('dir_seq',):
            self._plan_cache = ops.build_rename_plan(self.current_files(), opts)
        from engine.core import TaskSpec
        files = self._override_files or [p[0] for p in self._plan_cache]
        if not files:
            QMessageBox.information(self, self.main.APP_NAME, '请先选择文件')
            return
        options = dict(opts)
        if opts.get('mode') in ('excel', 'excel_dir'):
            options['_excel_map'] = self._excel_map_cache
        else:
            options['_rename_plan'] = self._plan_cache
        self._autosave_params()
        spec = TaskSpec(kind=self.kind, title=self.title, files=files,
                        options=options)
        self._start_worker(spec, self.processor())

    def collect_options(self):
        o = self._collect_mode_options()
        return o

    def supports_preview(self):
        return True

    def build_preview(self, files, options):
        mode = options.get('mode')
        mapping = None
        if mode in ('excel', 'excel_dir'):
            xf = options.get('excel_map', '')
            try:
                mapping = ops.load_excel_map(xf) if xf else {}
            except Exception:
                return ('错误',), [('读取对照表失败', '', '')]
            rows = [('原名称', '新名称')]
            for f in files:
                old = os.path.basename(f)
                rows.append((old, mapping.get(old, '(对照表中无)')))
            return rows
        if mode == 'dir_seq':
            rows = [('文件夹', '新名称')]
            for i, f in enumerate(files):
                idx = options.get('seq_start', 1) + i
                pad = int(options.get('seq_pad', 3) or 3)
                rows.append((os.path.basename(f),
                             options.get('prefix', '') + str(idx).zfill(pad)))
            return rows
        plan = ops.build_rename_plan(files, options)
        rows = [('原文件名', '新文件名')]
        for _old, _new, nn in plan:
            rows.append((os.path.basename(_old), nn))
        return rows


class FilesPage(FunctionPage):
    title = '文件整理工具'
    kind = 'files'
    OP_HINTS = {
        '导出文件清单': '列出文件夹里所有文件（路径/大小）存成 CSV',
        '导出清单到 Excel': '同上，但生成排版好的 Excel 表格',
        '打包为 zip': '把选中的文件/文件夹压缩成一个 zip',
        '解压 zip': '解压 zip 文件（自动修复中文文件名乱码）',
        '按扩展名分类': '把文件按类型自动移进「jpg/」「pdf/」等子文件夹',
        '按日期分类': '把文件按修改日期自动移进「2026-01-01/」等子文件夹',
        '复制到多目录': '把文件同时复制到多个目标文件夹（分号分隔）',
        '多源复制到一处': '把多个文件夹的内容汇总复制到一个目标文件夹',
        '创建同名文件夹': '给每个文件建一个同名文件夹（整理素材常用）',
        '删除文件(进回收备份区)': '删除文件但可在「回收站」页随时恢复',
        '删除文件夹(进回收备份区)': '删除整个文件夹但可随时恢复',
        'Excel 清单批量复制': '按 Excel 清单（源路径|目标路径）批量复制',
        'Excel 批量创建文件/文件夹': '按 Excel 清单（路径|类型）批量创建',
        '冗余命名清理': '自动去掉文件名里的「副本/最终版/(1)」等冗余标记',
        '重复文件查找': '找出内容完全相同的重复文件并生成报告',
        '自动清理重复文件': '重复文件每组自动保留一份，其余移入回收备份区',
        '清理空文件夹': '把空文件夹移入回收备份区（可恢复）',
        '清理空文件': '把 0 字节的空文件移入回收备份区（可恢复）',
        '文件哈希校验': '计算 MD5/SHA 值用于校验文件一致性，导出 CSV',
        '文本行统计': '统计文本文件的行数/字数/字符数，导出 Excel',
        '修改时间批改(固定时间)': '把文件修改时间统一改成指定时间',
        '修改时间批改(按文件名日期)': '从文件名里识别日期并改成该日期',
    }
    CONFIRM_OPS = {
        '删除文件(进回收备份区)':
            '把选中的 {n} 个文件移入回收备份区？\n'
            '可在「回收站」页随时恢复到原位置。',
        '删除文件夹(进回收备份区)':
            '把选中的 {n} 个文件夹移入回收备份区？\n'
            '可在「回收站」页随时恢复到原位置。',
        '自动清理重复文件':
            '将自动清理重复文件：每组保留最早的一份，'
            '其余 {n} 个起始文件所在的范围内重复副本全部移入回收备份区。\n继续吗？',
        '清理空文件夹':
            '将把范围内的空文件夹移入回收备份区（可恢复）。\n继续吗？',
        '清理空文件':
            '将把范围内的空文件（0 字节）移入回收备份区（可恢复）。\n继续吗？',
    }
    desc = ('文件整理：清单/压缩/解压/分类/复制/同名文件夹/删除(可恢复)/'
            'Excel 批量操作/清理')
    OPS = ('导出文件清单', '导出清单到 Excel', '打包为 zip', '解压 zip',
           '按扩展名分类', '按日期分类', '复制到多目录', '多源复制到一处',
           '创建同名文件夹', '删除文件(进回收备份区)',
           '删除文件夹(进回收备份区)', 'Excel 清单批量复制',
           'Excel 批量创建文件/文件夹', '冗余命名清理', '重复文件查找',
           '自动清理重复文件', '清理空文件夹', '清理空文件', '文件哈希校验',
           '文本行统计', '修改时间批改(固定时间)',
           '修改时间批改(按文件名日期)')
    OP_FIELDS = {
        '复制到多目录': ('targets',),
        '多源复制到一处': ('merge_target',),
        'Excel 清单批量复制': ('excel_edit',),
        'Excel 批量创建文件/文件夹': ('excel_edit',),
        '文件哈希校验': ('hash_alg',),
        '修改时间批改(固定时间)': ('ts_fixed',),
    }
    DIR_LEVEL_OPS = frozenset({'清理空文件夹', '清理空文件', '重复文件查找',
                               '自动清理重复文件', '文件哈希校验', '文本行统计',
                               '修改时间批改(固定时间)',
                               '修改时间批改(按文件名日期)',
                               '按扩展名分类', '按日期分类',
                               '删除文件夹(进回收备份区)',
                               'Excel 清单批量复制',
                               'Excel 批量创建文件/文件夹'})

    def build_options(self, form):
        self._field(form, 'op', '操作', _combo(self.OPS))
        t = QLineEdit()
        t.setPlaceholderText('多目标目录, 分号分隔(复制到多目录)')
        self._field(form, 'targets', '目标目录', t)
        tb = QPushButton('选目录')
        tb.clicked.connect(self.pick_targets)
        form.addRow('', tb)
        m = QLineEdit()
        m.setPlaceholderText('汇总目标目录(多源复制)')
        self._field(form, 'merge_target', '汇总目标', m)
        mb = QPushButton('选汇总目录')
        mb.clicked.connect(self.pick_merge)
        form.addRow('', mb)
        e = QLineEdit()
        e.setPlaceholderText('清单 Excel 路径')
        self._field(form, 'excel_edit', 'Excel 清单', e)
        eb = QPushButton('选 Excel')
        eb.clicked.connect(self.pick_excel)
        form.addRow('', eb)
        od = QLineEdit()
        od.setPlaceholderText('输出/工作目录(留空=源目录)')
        self._field(form, 'out_dir', '输出目录', od)
        ob = QPushButton('输出到…')
        ob.clicked.connect(self.pick_out)
        form.addRow('', ob)
        self._field(form, 'hash_alg', '哈希算法', _combo(('md5', 'sha1', 'sha256')))
        tf = QLineEdit()
        tf.setPlaceholderText('YYYY-MM-DD HH:MM:SS(固定时间模式)')
        self._field(form, 'ts_fixed', '目标时间', tf)
        # 信号由基类 _setup_op_fields 统一连接, 避免重复触发
        self.apply_op_fields()

    def pick_targets(self):
        d = QFileDialog.getExistingDirectory(self, '选择目标目录')
        if d:
            cur = self._widgets['targets'].text()
            self._widgets['targets'].setText((cur + ';' + d).strip(';'))

    def pick_merge(self, checked=False):
        d = QFileDialog.getExistingDirectory(self, '选择汇总目标目录')
        if d:
            self._widgets['merge_target'].setText(d)

    def pick_excel(self):
        f, _ = QFileDialog.getOpenFileName(self, '选择清单 Excel', '',
                                           'Excel (*.xlsx)')
        if f:
            self._widgets['excel_edit'].setText(f)

    def pick_out(self):
        d = QFileDialog.getExistingDirectory(self, '选择输出目录')
        if d:
            self._widgets['out_dir'].setText(d)

    def collect_options(self):
        op = self._widgets['op'].currentText()
        key_map = {'导出文件清单': 'list', '导出清单到 Excel': 'list_excel',
                   '打包为 zip': 'pack', '解压 zip': 'unzip',
                   '按扩展名分类': 'by_ext', '按日期分类': 'by_date',
                   '复制到多目录': 'copy', '多源复制到一处': 'merge_src',
                   '创建同名文件夹': 'mkdir',
                   '删除文件(进回收备份区)': 'delete',
                   '删除文件夹(进回收备份区)': 'delete_dir',
                   'Excel 清单批量复制': 'copy_by_excel',
                   'Excel 批量创建文件/文件夹': 'create_by_excel',
                   '冗余命名清理': 'cleanup', '重复文件查找': 'dup_find',
                   '自动清理重复文件': 'dup_clean', '清理空文件夹': 'empty_dirs',
                   '清理空文件': 'empty_files', '文件哈希校验': 'hash',
                   '文本行统计': 'tstats',
                   '修改时间批改(固定时间)': 'ts_fixed',
                   '修改时间批改(按文件名日期)': 'ts_name'}
        # op=机内名(process_fileops 等处理器读取), op_name=显示名(台账/确认用)
        o = {'op': key_map.get(op, ''), 'key': key_map.get(op, ''),
             'op_name': op}
        if op == '复制到多目录':
            o['targets'] = self._widgets['targets'].text()
        elif op == '多源复制到一处':
            o['merge_target'] = self._widgets['merge_target'].text()
        elif op in ('Excel 清单批量复制', 'Excel 批量创建文件/文件夹'):
            o['excel_file'] = self._widgets['excel_edit'].text()
        if 'out_dir' in self._widgets:
            o['out_dir'] = self._widgets['out_dir'].text().strip()
        o['hash_alg'] = self._widgets['hash_alg'].currentText()
        o['ts_fixed'] = self._widgets['ts_fixed'].text()
        return o

    def processor(self):
        o = self.collect_options()
        key = o.get('key')
        return {
            'list': ops.process_fileops, 'pack': self._pack_multi,
            'unzip': ops.process_zip, 'by_ext': ops.process_fileops,
            'by_date': ops.process_fileops, 'copy': ops.process_fileops,
            'mkdir': ops.process_fileops, 'delete': ops.process_fileops,
            'delete_dir': advanced_ops.files_delete_folder,
            'list_excel': ops.files_export_excel, 'hash': ops.file_hasher,
            'tstats': ops.text_stats,
            'ts_fixed': lambda p, op: ops.touch_timestamps(p, {**op, 'mode': 'fixed'}),
            'ts_name': lambda p, op: ops.touch_timestamps(p, {**op, 'mode': 'from_name'}),
            'dup_find': ops.find_duplicate_files,
            'dup_clean': self._dup_clean_proc,
            'empty_dirs': lambda p, op: ops.cleanup_empty(p, {**op, 'empty_mode': 'empty_dirs'}),
            'empty_files': lambda p, op: ops.cleanup_empty(p, {**op, 'empty_mode': 'empty_files'}),
            'cleanup': self._cleanup_proc,
            'merge_src': self._merge_src,
            'copy_by_excel': self._copy_by_excel,
            'create_by_excel': self._create_by_excel,
        }.get(key)

    def _dup_clean_proc(self, path, options):
        return ops.find_duplicate_files(
            path, {**options, 'auto_keep': True,
                   'trash_dir': self.main.TRASH_DIR})

    def _cleanup_proc(self, path, options):
        plan = options.get('_rename_plan') or []
        for oldp, newp, nn in plan:
            if os.path.abspath(oldp) == os.path.abspath(path):
                if os.path.abspath(oldp) == os.path.abspath(newp):
                    return FileResult(path, 'skipped', '无冗余标记')
                os.rename(oldp, newp)
                undo = [{'src': oldp, 'dst': newp}]
                return FileResult(path, 'ok', '清理为 ' + nn, newp, undo=undo)
        return FileResult(path, 'fail', '计划中未找到该文件')

    def _merge_src(self, path, options):
        tgt = options.get('merge_target', '')
        if not tgt:
            return FileResult(path, 'fail', '未填写汇总目标目录')
        srcs = [s for s in (options.get('_src_paths') or [path])
                if os.path.isdir(s)]
        if not srcs:
            return FileResult(path, 'skipped', '所选项不是文件夹')
        n, dst, fails = advanced_ops.files_merge_folders(
            srcs, {'merge_target': tgt, 'keep_parent': False})
        detail = f'合并 {len(srcs)} 个源文件夹, {n} 个文件 → {dst}'
        if fails:
            detail += f'; {len(fails)} 个失败: ' + fails[0][1][:60]
        return FileResult(path, 'ok' if not fails else 'fail', detail, dst)

    def _pack_multi(self, path, options):
        """v3: 打包为 zip — 多选时合并为一个压缩包, 单选按原语义。"""
        import zipfile
        paths = [p for p in (options.get('_src_paths') or [path])
                 if os.path.exists(p)]
        if not paths:
            return FileResult(path, 'skipped', '没有可打包的内容')
        out_dir = options.get('out_dir') or os.path.dirname(paths[0])
        os.makedirs(out_dir, exist_ok=True)
        if len(paths) == 1:
            name = os.path.basename(paths[0].rstrip('\\/')) or '打包'
        else:
            from datetime import datetime as _dt
            name = '打包_' + _dt.now().strftime('%Y%m%d_%H%M%S')
        zip_path = os.path.abspath(os.path.join(out_dir, name + '.zip'))
        # 打包目标等于某个输入文件(典型: 对 data.zip 再打包, 且输出目录
        # 就是它所在目录)时, ZipFile('w') 会先截断该文件, 再把正在写的
        # 压缩包当内容打进去 —— 原压缩包内容全部丢失。
        if any(os.path.abspath(p) == zip_path for p in paths):
            zip_path = os.path.abspath(os.path.join(
                out_dir, name + '_打包.zip'))
        # 同名压缩包已存在时不能静默覆盖
        if os.path.exists(zip_path):
            b, e = os.path.splitext(zip_path)
            k = 1
            while os.path.exists(f'{b}({k}){e}'):
                k += 1
            zip_path = f'{b}({k}){e}'
        try:
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for p in paths:
                    if os.path.isfile(p):
                        zf.write(p, os.path.basename(p))
                    else:
                        base = os.path.basename(p.rstrip('\\/'))
                        for dp, _dns, fns in os.walk(p):
                            for fn in fns:
                                fp = os.path.join(dp, fn)
                                # 输出落在源目录内时, 排除正在写的 zip 自身,
                                # 否则包内会嵌一个半截的自己
                                if os.path.abspath(fp) == zip_path:
                                    continue
                                zf.write(fp, os.path.join(
                                    base, os.path.relpath(fp, p)))
        except Exception as e:
            try:
                if os.path.exists(zip_path):
                    os.remove(zip_path)
            except OSError:
                pass
            return FileResult(path, 'fail', f'打包失败: {e}')
        return FileResult(path, 'ok', f'已打包 {len(paths)} 项 → {zip_path}',
                          zip_path)

    def _read_excel_rows(self, options):
        import openpyxl
        f = options.get('excel_file', '')
        if not f or not os.path.exists(f):
            return None
        wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
        ws = wb.active
        rows = []
        headers = []
        for i, r in enumerate(ws.iter_rows(values_only=True)):
            vals = ['' if v is None else str(v).strip() for v in r]
            if i == 0:
                headers = vals
                continue
            if any(vals):
                rows.append({headers[j] if j < len(headers) else f'col{j}': v
                             for j, v in enumerate(vals)})
        wb.close()
        return rows

    def _copy_by_excel(self, path, options):
        rows = self._read_excel_rows(options)
        if rows is None:
            return FileResult(path, 'fail', '请选择清单 Excel(列: path | target)')
        n, fails = advanced_ops.files_copy_by_excel(rows, options)
        detail = f'复制 {n} 项' + (f', 失败 {len(fails)}' if fails else '')
        return FileResult(path, 'ok', detail, options.get('out_dir') or '')

    def _create_by_excel(self, path, options):
        rows = self._read_excel_rows(options)
        if rows is None:
            return FileResult(path, 'fail', '请选择清单 Excel(列: path | type)')
        base_dir = options.get('out_dir') or os.path.dirname(path)
        n = advanced_ops.files_create_from_excel(rows, base_dir)
        return FileResult(path, 'ok', f'创建 {n} 项 → {base_dir}', base_dir)

    def do_run(self):
        opts = self.collect_options()
        op = opts.get('op_name') or opts.get('op')  # 显示名
        if op == '冗余命名清理':
            files = self.current_files()
            if not files:
                QMessageBox.information(self, self.main.APP_NAME, '请先选择文件')
                return
            plan = ops.build_rename_plan(
                files, {'mode': 'cleanup',
                        'cleanup_patterns': ops.CLEANUP_PATTERNS})
            self._plan_cache = plan
            opts['_rename_plan'] = plan
            self._run_dir_level(opts, files=[p[0] for p in plan])
            return
        if op == '打包为 zip':
            paths = [p.strip() for p in self.file_edit.text().split('||')
                     if p.strip() and os.path.exists(p.strip())]
            if not paths:
                QMessageBox.information(self, self.main.APP_NAME,
                                        '请先选择文件或文件夹')
                return
            opts['_src_paths'] = paths
            self._run_dir_level(opts, files=paths, batch=True)
            return
        if op == '多源复制到一处':
            paths = [p.strip() for p in self.file_edit.text().split('||')
                     if p.strip() and os.path.exists(p.strip())]
            opts['_src_paths'] = paths
            self._run_dir_level(opts, files=paths, batch=True)
            return
        if op in self.DIR_LEVEL_OPS:
            self._run_dir_level(opts)
            return
        super().do_run()

    def _run_dir_level(self, opts, files=None, batch=False):
        """目录级操作: 不走逐文件 worker, 直接调用处理器并写台账。

        files 为 None 时处理全部已选路径(与预览一致 —— 早期只取
        split('||')[0], 第二个目录根本不动, 而预览却列了两行);
        files 为列表时逐项调用(如冗余命名清理的每一条计划);
        batch=True 时整批只调用一次(打包/多源合并这类"把 N 项合成一个
        结果"的操作 —— 逐项调用会生成 N 个内容相同的 zip)。
        """
        if files:
            targets = list(files)
        else:
            targets = [p.strip() for p in self.file_edit.text().split('||')
                       if p.strip()]
        targets = [t for t in targets if t and os.path.exists(t)]
        if not targets:
            QMessageBox.information(self, self.main.APP_NAME,
                                    '请先选择有效的文件或文件夹')
            return
        # 危险操作确认(与逐文件执行同一套 CONFIRM_OPS)
        op_name = opts.get('op_name', '')
        confirm_tpl = getattr(self, 'CONFIRM_OPS', {}).get(op_name)
        if confirm_tpl:
            ret = QMessageBox.question(
                self, self.main.APP_NAME, confirm_tpl.format(n=len(targets)))
            if ret != QMessageBox.Yes:
                return
        proc = self.processor()
        if proc is None:
            QMessageBox.information(self, self.main.APP_NAME,
                                    '该操作暂无可用处理器')
            return
        # 交给 BatchWorker 在后台跑。早期这里是主线程同步 for 循环:
        # 大目录去重(整树 MD5)/哈希/打包时界面直接冻结, 没有进度条也取消
        # 不了。worker 会顺带写台账和撤销清单, schema 与逐文件流程一致。
        from engine.core import TaskSpec
        self._autosave_params()
        calls = targets[:1] if batch else targets
        spec = TaskSpec(kind=self.kind, title=self.title,
                        files=calls, options=opts)
        self._start_worker(spec, proc)

    def supports_preview(self):
        return True

    def build_preview(self, files, options):
        rows = [('文件', '将执行')]
        op = options.get('op_name') or options.get('op', '')
        for f in files[:100]:
            nm = os.path.basename(f)
            if op == '按扩展名分类':
                ext = os.path.splitext(nm)[1].lower().lstrip('.') or '无扩展名'
                rows.append((nm, f'移动到 {ext}/ 子目录'))
            elif op == '按日期分类':
                from datetime import datetime
                day = datetime.fromtimestamp(os.path.getmtime(f)).date().isoformat()
                rows.append((nm, f'移动到 {day}/ 子目录'))
            elif op in ('删除文件(进回收备份区)', '删除文件夹(进回收备份区)'):
                rows.append((nm, '移入回收备份区(可恢复)'))
            elif op == '创建同名文件夹':
                rows.append((nm, '创建同名文件夹 ' +
                             os.path.splitext(nm)[0] + '/'))
            else:
                rows.append((nm, options.get('op_name', op)))
        return rows


class TemplatePage(FunctionPage):
    title = '模板批量生成'
    kind = 'template'
    desc = ('模板用 {{字段}} 占位; 数据源 Excel 每行一条, 列名=字段名。\n'
            '支持 docx/xlsx/pptx/txt 模板, 每行数据生成一个文件。')
    OPS = ('生成文档',)

    def build_options(self, form):
        self._field(form, 'tpl_edit', '模板', QLineEdit())
        self._field(form, 'tpl_btn', '', QPushButton('选模板'))
        self._widgets['tpl_btn'].clicked.connect(
            lambda: self._pick(self._widgets['tpl_edit'],
                               '模板文件 (*.docx *.xlsx *.pptx *.txt)'))
        self._field(form, 'data_edit', '数据源', QLineEdit())
        self._field(form, 'data_btn', '', QPushButton('选数据源'))
        self._widgets['data_btn'].clicked.connect(
            lambda: self._pick(self._widgets['data_edit'], 'Excel (*.xlsx)'))
        self._field(form, 'name_col', '文件名列', QLineEdit())
        self._widgets['name_col'].setPlaceholderText('输出文件名列(留空=序号)')
        od = QLineEdit()
        od.setPlaceholderText('输出目录(必填)')
        self._field(form, 'out_dir', '输出目录', od)
        ob = QPushButton('选目录')
        ob.clicked.connect(self.pick_out)
        form.addRow('', ob)
        self.apply_op_fields = lambda: None

    def _pick(self, edit, filt):
        f, _ = QFileDialog.getOpenFileName(self, '选择文件', '', filt)
        if f:
            edit.setText(f)

    def pick_out(self):
        d = QFileDialog.getExistingDirectory(self, '选择输出目录')
        if d:
            self._widgets['out_dir'].setText(d)

    def collect_options(self):
        o = {'template': self._widgets['tpl_edit'].text(),
             'data': self._widgets['data_edit'].text(),
             'name_col': self._widgets['name_col'].text(),
             'out_dir': self._widgets['out_dir'].text()}
        if getattr(self, '_rows_cache', None):
            o['_rows'] = self._rows_cache
        return o

    def processor(self):
        return self._gen_proc

    def _gen_proc(self, path, options):
        tpl = options.get('template', '')
        if not os.path.exists(tpl):
            return FileResult(path, 'fail', '模板不存在')
        rows = options.get('_rows') or []
        if not rows:
            return FileResult(path, 'fail', '无数据行')
        row = rows[options.get('_row_idx', 0)]
        out_dir = options.get('out_dir') or os.path.dirname(tpl)
        os.makedirs(out_dir, exist_ok=True)
        name_col = options.get('name_col', '')
        name = str(row.get(name_col, '')).strip() if name_col else ''
        if not name:
            name = 'doc_' + format(options.get('_row_idx', 0) + 1, '03d')
        for c in '\\/:*?"<>|':
            name = name.replace(c, '_')
        ext = os.path.splitext(tpl)[1]
        out = os.path.join(out_dir, name + ext)
        # 数据里文件名列重复时早期直接覆盖, 结果两个"张三"只剩最后一个,
        # 两行却都报成功。重名改成 张三(1).txt。
        if os.path.exists(out):
            stem, _e = os.path.splitext(out)
            k = 1
            while os.path.exists(f'{stem}({k}){_e}'):
                k += 1
            out = f'{stem}({k}){_e}'
        office_ops.generate_from_template(tpl, row, out)
        return FileResult(path, 'ok', '→ ' + out, out)

    def do_run(self):
        import openpyxl
        opts = self.collect_options()
        if not opts.get('template') or not os.path.exists(opts['template']):
            QMessageBox.information(self, self.main.APP_NAME, '请选择模板文件')
            return
        if not opts.get('data') or not os.path.exists(opts['data']):
            QMessageBox.information(self, self.main.APP_NAME, '请选择数据源 Excel')
            return
        if not opts.get('out_dir'):
            QMessageBox.information(self, self.main.APP_NAME, '请选择输出目录')
            return
        wb = openpyxl.load_workbook(opts['data'], read_only=True)
        ws = wb.active
        rows = []
        headers = []
        for i, r in enumerate(ws.iter_rows(values_only=True)):
            vals = ['' if v is None else str(v).strip() for v in r]
            if i == 0:
                headers = vals
                continue
            if any(vals):
                rows.append({headers[j] if j < len(headers) else f'col{j}': v
                             for j, v in enumerate(vals)})
        wb.close()
        if not rows:
            QMessageBox.information(self, self.main.APP_NAME, '数据源没有数据行')
            return
        self._rows_cache = rows

        def proc(_path, options):
            idx = int(str(_path).replace('row', '') or 0)
            return self._gen_proc(_path, {**options, '_row_idx': idx})

        files = [f'row{i}' for i in range(len(rows))]
        self._autosave_params()
        from engine.core import TaskSpec
        spec = TaskSpec(kind=self.kind, title=self.title, files=files,
                        options={**opts, '_rows': rows})
        self._start_worker(spec, proc)

    def _restore_page_extras(self):
        pass
