# -*- coding: utf-8 -*-
"""青雁批量工具箱 v3 — 特殊页面。

Home(搜索/收藏/最近) / Merge(合并拆分) / Tools(正则/JSON/差异) /
Records(台账+统计+撤销) / Trash(回收备份区) / Pipeline(流水线, 线程化+命名保存) /
Settings(主题!数据目录/方案管理) / Help / About / OfficeTemplates。
"""
from __future__ import annotations

import difflib
import json
import os

from PySide6.QtCore import Qt, QTimer, Signal
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import (QComboBox, QDialog, QFileDialog,
                               QFormLayout, QFrame, QGroupBox,
                               QHBoxLayout, QHeaderView, QLabel, QLineEdit,
                               QListWidget, QMessageBox,
                               QPlainTextEdit, QPushButton,
                               QSpinBox, QTableWidget,
                               QTableWidgetItem, QTabWidget, QVBoxLayout,
                               QWidget)

from app import icons
from app.layouts import WrapRow
from engine.store import ParamStore


def _to_int(v, default: int = 0) -> int:
    """宽松取整数: 台账 JSON 可能被手工编辑或由外部工具写入, 直接 int() 会崩。"""
    try:
        return int(v)
    except (TypeError, ValueError):
        return default


def _rec_files(rec: dict) -> list:
    """取台账的文件明细列表, 非 list 时返回空列表避免遍历崩溃。"""
    files = rec.get('files') if isinstance(rec, dict) else None
    if not isinstance(files, list):
        return []
    return [f for f in files if isinstance(f, dict)]


def _section_header(label: str, icon_name: str, color: str) -> QWidget:
    """统一区块标题: 图标 + 文字。"""
    w = QWidget()
    h = QHBoxLayout(w)
    h.setContentsMargins(0, 4, 0, 4)
    h.setSpacing(8)
    ic = QLabel()
    ic.setPixmap(icons.get_icon_pixmap(icon_name, 18, color))
    cap = QLabel(label)
    cap.setStyleSheet(
        f'font-size:14px;font-weight:600;letter-spacing:1px;color:{color};')
    h.addWidget(ic)
    h.addWidget(cap)
    h.addStretch(1)
    return w


# ==================================================================== 首页
class HomePage(QWidget):
    def __init__(self, parent_main):
        super().__init__()
        self.main = parent_main
        from app.home_workspace import build_home
        build_home(self)

    def _rebuild_recent(self):
        from app.home_workspace import refresh_recent
        refresh_recent(self)

    def refresh_recent(self):
        self._rebuild_recent()

    def _rebuild_favs(self):
        from app.home_workspace import refresh_favorites
        refresh_favorites(self)

    def _toggle_fav(self, fname, tp, op, remove):
        from app.home_workspace import toggle_favorite
        toggle_favorite(self, fname, tp, op, remove)

    def _fav_menu(self, pos, btn, fname, tp, op):
        from PySide6.QtWidgets import QMenu
        menu=QMenu(self)
        favorite=self._fav.is_favorite(fname,tp)
        action=menu.addAction('取消收藏' if favorite else '加入收藏')
        action.triggered.connect(lambda:self._toggle_fav(fname,tp,op,favorite))
        menu.exec(btn.mapToGlobal(pos))


# ================================================================ 合并拆分
class FileDropList(QListWidget):
    paths_dropped = Signal(list)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():event.acceptProposedAction()
        else:super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():event.acceptProposedAction()
        else:super().dragMoveEvent(event)

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            paths = [url.toLocalFile() for url in event.mimeData().urls() if os.path.isfile(url.toLocalFile())]
            for path in paths:self.addItem(path)
            self.paths_dropped.emit(paths)
            event.acceptProposedAction()
        else:super().dropEvent(event)


class MergePage(QWidget):
    title = '合并与拆分'

    def __init__(self, parent_main):
        super().__init__()
        self.main = parent_main
        v = QVBoxLayout(self)
        v.setContentsMargins(32, 26, 26, 16)
        v.setSpacing(10)
        cap = QLabel('合并与拆分')
        cap.setStyleSheet(
            f'font-size:21px;font-weight:600;color:{self.main.theme.INK};')
        v.addWidget(cap)
        info = QLabel('把多个 PDF/Word/文本/Excel/图片 按列表顺序合并为一个文件；'
                      'PDF/PPT 支持按页数拆分；图片支持拼接长图或合并 PDF。')
        info.setWordWrap(True)
        info.setStyleSheet(f'color:{self.main.theme.MUTED};')
        v.addWidget(info)

        self.listw = FileDropList()
        self.listw.paths_dropped.connect(lambda paths:self.main.show_status(f'已拖入 {len(paths)} 个文件'))
        self.listw.setAcceptDrops(True)
        self.listw.setDragDropMode(QListWidget.InternalMove)
        v.addWidget(self.listw, 1)

        brow = WrapRow()
        for text, cb in (('添加文件', self.add_files),
                         ('移除选中', self.remove_sel),
                         ('上移', lambda: self.move(-1)),
                         ('下移', lambda: self.move(1))):
            b = QPushButton(text)
            b.clicked.connect(cb)
            brow.addWidget(b)
        brow.addStretch(1)
        v.addLayout(brow)

        form = QFormLayout()
        self.type_box = QComboBox()
        self.type_box.addItems(('PDF 合并', 'PDF 每 N 页拆分', '文本合并',
                                'Word 合并(追加段落)', 'Excel 合并(多 Sheet)',
                                'PPT 合并', '图片拼接长图', '图片合并 PDF'))
        form.addRow('合并类型', self.type_box)
        self.chunk = QSpinBox()
        self.chunk.setValue(1)
        self.chunk.setRange(1, 10000)
        form.addRow('每份页数(拆分)', self.chunk)
        v.addLayout(form)

        self.run = QPushButton('执行合并/拆分')
        self.run.setProperty('accent', True)
        self.run.setStyleSheet(
            f'background:{self.main.theme.ACCENT};color:white;padding:9px 34px;'
            'font-weight:600;border:none;border-radius:9px;')
        self.run.clicked.connect(self.do_merge)
        v.addWidget(self.run)
        self.status = QLabel('')
        self.status.setStyleSheet('font-weight:700;')
        v.addWidget(self.status)

    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, '选择文件(按住 Ctrl 多选)')
        for f in files:
            self.listw.addItem(f)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        n = 0
        for url in event.mimeData().urls():
            p = url.toLocalFile()
            if os.path.exists(p):
                self.listw.addItem(p)
                n += 1
        if n:
            self.main.show_status(f'已从拖放加入 {n} 项')

    def remove_sel(self):
        for item in self.listw.selectedItems():
            self.listw.takeItem(self.listw.row(item))

    def move(self, delta):
        row = self.listw.currentRow()
        n = self.listw.count()
        if row < 0 or n == 0:
            return
        new = max(0, min(n - 1, row + delta))
        if new == row:
            return
        item = self.listw.takeItem(row)
        self.listw.insertItem(new, item)
        self.listw.setCurrentRow(new)

    def do_merge(self):
        from engine.core import BatchWorker, TaskSpec
        from engine.merge_ops import merge_files
        if getattr(self.main, '_task_running', False) or any(getattr(p, 'worker', None) is not None and p.worker.isRunning()
               for p in self.main.pages):
            QMessageBox.information(self, self.main.APP_NAME, '请等待当前任务完成。')
            return
        files = [self.listw.item(i).text() for i in range(self.listw.count())]
        if not files:
            QMessageBox.information(self, self.main.APP_NAME, '请先添加要合并的文件')
            return
        kind = self.type_box.currentText()
        if kind == 'PDF 每 N 页拆分':
            out = QFileDialog.getExistingDirectory(self, '选择拆分输出目录')
        else:
            extension = {'PDF 合并':'pdf','文本合并':'txt','Word 合并(追加段落)':'docx',
                         'Excel 合并(多 Sheet)':'xlsx','PPT 合并':'pptx',
                         '图片拼接长图':'png','图片合并 PDF':'pdf'}[kind]
            out, _ = QFileDialog.getSaveFileName(self, '保存结果', 'merged.'+extension,
                                                '*.'+extension)
        if not out: return
        spec = TaskSpec('merge', self.title, [files[0]],
                        {'files':files,'type':kind,'destination':out,'chunk':self.chunk.value()})
        self.worker = BatchWorker(spec, merge_files, self.main.RECORDS_DIR, self.main.UNDO_MANIFEST)
        self.worker.progress.connect(self.main.on_progress)
        self.worker.finished_all.connect(self._merge_finished)
        self.worker.cancelled.connect(lambda:self._merge_finished(self.worker.results,''))
        self.main.set_running(True, self.title)
        self.status.setText('正在处理，请稍候…')
        self.worker.start()

    def _merge_finished(self, results, task_id):
        self.main.set_running(False)
        self.main.records_dirty = True
        if results:
            r = results[0]
            self.status.setText(('完成 · ' if r.status == 'ok' else '未完成 · ') + r.detail +
                                (' → ' + r.output if r.output else ''))
        else:
            self.status.setText('任务已取消')


# ================================================================= 辅助工具
class ToolsPage(QWidget):
    title = '辅助工具'

    def __init__(self, parent_main):
        super().__init__()
        self.main = parent_main
        v = QVBoxLayout(self)
        v.setContentsMargins(32, 26, 26, 16)
        cap = QLabel('辅助工具')
        cap.setStyleSheet(
            f'font-size:21px;font-weight:600;color:{self.main.theme.INK};')
        v.addWidget(cap)
        tabs = QTabWidget()

        # 正则测试
        rx = QWidget()
        rv = QVBoxLayout(rx)
        self.rx_pattern = QLineEdit()
        self.rx_pattern.setPlaceholderText('输入正则, 如 \\d{4}-\\d{2}-\\d{2}')
        rv.addWidget(self.rx_pattern)
        self.rx_text = QPlainTextEdit()
        self.rx_text.setPlaceholderText('粘贴待测文本…')
        rv.addWidget(self.rx_text)
        self.rx_result = QLabel('')
        self.rx_result.setStyleSheet(
            f'color:{self.main.theme.ACCENT};font-weight:700;')
        self.rx_result.setWordWrap(True)
        rv.addWidget(self.rx_result)
        self.rx_common = QComboBox()
        self.rx_common.addItems(
            ('— 常用模板 —', '邮箱: [\\w.+-]+@[\\w-]+\\.[\\w.]+',
             '手机号: 1[3-9]\\d{9}', '日期: \\d{4}-\\d{2}-\\d{2}',
             'URL: https?://[^\\s]+', '整数: -?\\d+',
             '中文: [\u4e00-\u9fff]+', '身份证: \\d{17}[\\dXx]'))
        self.rx_common.currentIndexChanged.connect(self._rx_tpl)
        self.rx_pattern.textChanged.connect(self._rx_run)
        self.rx_text.textChanged.connect(self._rx_run)
        rv.addWidget(self.rx_common)
        tabs.addTab(rx, '正则测试')

        # 正则生成器
        rg = QWidget()
        gv = QVBoxLayout(rg)
        form = QFormLayout()
        self.rg_kind = QComboBox()
        self.rg_kind.addItems(('单个任意文本', '单个数字', '单个汉字',
                               '单个英文字母', '单个空白', '单个标点符号',
                               '中国大陆手机号', '邮箱',
                               '中国大陆18位身份证号', '网址', '日期', '整数'))
        self.rg_text = QLineEdit()
        self.rg_text.setPlaceholderText('自定义文本内容(选自定义文本时)')
        self.rg_count = QComboBox()
        self.rg_count.addItems(('任意次 (*)', '1 次以上 (+)', '恰好 1 次',
                                '恰好 2 次', '恰好 3 次', '恰好 4 次',
                                '恰好 6 次', '恰好 11 次', '恰好 17 次',
                                '恰好 18 次'))
        self.rg_pos = QComboBox()
        self.rg_pos.addItems(('', '某个文本的左边位置', '某个文本的右边位置'))
        self.rg_ref = QLineEdit()
        self.rg_ref.setPlaceholderText('参照文本(位置限定用)')
        form.addRow('内容类型', self.rg_kind)
        form.addRow('自定义文本', self.rg_text)
        form.addRow('重复次数', self.rg_count)
        form.addRow('位置限定', self.rg_pos)
        form.addRow('参照文本', self.rg_ref)
        gv.addLayout(form)
        rgr = QPushButton('生成正则')
        rgr.clicked.connect(self._rg_run)
        gv.addWidget(rgr)
        self.rg_out = QPlainTextEdit()
        self.rg_out.setReadOnly(True)
        self.rg_out.setMaximumHeight(90)
        gv.addWidget(self.rg_out)
        send = QPushButton('发送到正则测试')
        send.clicked.connect(self._rg_send)
        gv.addWidget(send)
        gv.addStretch(1)
        tabs.addTab(rg, '正则生成器')

        # JSON 工具
        js = QWidget()
        jv = QVBoxLayout(js)
        self.json_edit = QPlainTextEdit()
        self.json_edit.setPlaceholderText('粘贴 JSON…')
        jv.addWidget(self.json_edit)
        jrow = QHBoxLayout()
        for text, cb in (('格式化', self._json_fmt), ('压缩', self._json_min),
                         ('校验', self._json_check)):
            b = QPushButton(text)
            b.clicked.connect(cb)
            jrow.addWidget(b)
        jrow.addStretch(1)
        self.json_status = QLabel('')
        self.json_status.setStyleSheet('font-weight:700;')
        jrow.addWidget(self.json_status)
        jv.addLayout(jrow)
        tabs.addTab(js, 'JSON 工具')

        # 文本差异
        df = QWidget()
        dv = QVBoxLayout(df)
        self.diff_a = QPlainTextEdit()
        self.diff_a.setPlaceholderText('文本 A')
        self.diff_b = QPlainTextEdit()
        self.diff_b.setPlaceholderText('文本 B')
        dv.addWidget(self.diff_a)
        dv.addWidget(self.diff_b)
        dbtn = QPushButton('对比')
        dbtn.clicked.connect(self._diff)
        dv.addWidget(dbtn)
        self.diff_out = QPlainTextEdit()
        self.diff_out.setReadOnly(True)
        dv.addWidget(self.diff_out, 1)
        tabs.addTab(df, '文本差异')

        # 剪贴板
        cb = QWidget()
        cv = QVBoxLayout(cb)
        cbbtn = QPushButton('读取当前剪贴板')
        cbbtn.clicked.connect(self._clip)
        cv.addWidget(cbbtn)
        self.clip_list = QListWidget()
        cv.addWidget(self.clip_list, 1)
        tabs.addTab(cb, '剪贴板')

        v.addWidget(tabs, 1)

    def _rx_tpl(self, i):
        t = self.rx_common.currentText()
        if ': ' in t:
            self.rx_pattern.setText(t.split(': ', 1)[1])

    def _rx_run(self):
        import re
        pat = self.rx_pattern.text()
        if not pat:
            self.rx_result.setText('')
            return
        try:
            ms = re.findall(pat, self.rx_text.toPlainText())
            self.rx_result.setText(
                f'✔ 匹配 {len(ms)} 处: ' + str(ms[:10])[:200])
        except re.error as e:
            self.rx_result.setText('✘ 正则错误: ' + str(e))

    def _rg_run(self):
        from engine.advanced_ops import regexp_builder
        cnt_map = {'任意次 (*)': 0, '1 次以上 (+)': -1, '恰好 1 次': 1,
                   '恰好 2 次': 2, '恰好 3 次': 3, '恰好 4 次': 4,
                   '恰好 6 次': 6, '恰好 11 次': 11, '恰好 17 次': 17,
                   '恰好 18 次': 18}
        pat = regexp_builder(self.rg_kind.currentText(),
                             self.rg_text.text(),
                             cnt_map.get(self.rg_count.currentText(), 0),
                             self.rg_pos.currentText(), self.rg_ref.text())
        self.rg_out.setPlainText(pat)

    def _rg_send(self):
        pat = self.rg_out.toPlainText().strip()
        if pat:
            self.rx_pattern.setText(pat)
            tabs = self.findChild(QTabWidget)
            if tabs:
                tabs.setCurrentIndex(0)

    def _json_fmt(self):
        import json
        try:
            obj = json.loads(self.json_edit.toPlainText())
            self.json_edit.setPlainText(
                json.dumps(obj, ensure_ascii=False, indent=2))
            self.json_status.setText('✔ 已格式化')
        except Exception as e:
            self.json_status.setText('✘ ' + str(e))

    def _json_min(self):
        import json
        try:
            obj = json.loads(self.json_edit.toPlainText())
            self.json_edit.setPlainText(json.dumps(obj, ensure_ascii=False, separators=(',', ':')))
            self.json_status.setText('✔ 已压缩')
        except Exception as e:
            self.json_status.setText('✘ ' + str(e))

    def _json_check(self):
        import json
        try:
            json.loads(self.json_edit.toPlainText())
            self.json_status.setText('✔ JSON 有效')
        except Exception as e:
            self.json_status.setText('✘ ' + str(e))

    def _diff(self):
        a = self.diff_a.toPlainText().splitlines()
        b = self.diff_b.toPlainText().splitlines()
        d = difflib.unified_diff(a, b, lineterm='')
        self.diff_out.setPlainText('\n'.join(d) or '两段文本相同')

    def _clip(self):
        text = QGuiApplication.clipboard().text()
        if text:
            self.clip_list.insertItem(0, text[:500])


# ================================================================= 任务台账
class RecordsPage(QWidget):
    title = '任务台账'

    def __init__(self, parent_main):
        super().__init__()
        self.main = parent_main
        v = QVBoxLayout(self)
        v.setContentsMargins(32, 26, 26, 16)
        v.setSpacing(10)
        cap = QLabel('任务台账')
        cap.setStyleSheet(
            f'font-size:21px;font-weight:600;color:{self.main.theme.INK};')
        v.addWidget(cap)
        info = QLabel(f'每条批量任务自动记录于: {self.main.RECORDS_DIR}')
        info.setWordWrap(True)
        info.setStyleSheet(f'color:{self.main.theme.MUTED};')
        v.addWidget(info)

        self.stats_row = QHBoxLayout()
        self.stats_row.setSpacing(10)
        self._stat_cards = {}
        for key, label in (('tasks', '任务总数'), ('files', '处理文件数'),
                           ('ok', '成功'), ('fail', '失败'),
                           ('recent', '最近 7 天')):
            card = QFrame()
            card.setStyleSheet(
                f'QFrame{{background:{self.main.theme.PAPER2};'
                f'border:1px solid {self.main.theme.LINE};'
                f'border-top:4px solid {self.main.theme.ACCENT};}}')
            cv = QVBoxLayout(card)
            cv.setContentsMargins(14, 10, 14, 10)
            num = QLabel('0')
            num.setStyleSheet(
                f'font-size:22px;font-weight:700;color:{self.main.theme.ACCENT};'
                'border:none;')
            lbl = QLabel(label)
            lbl.setStyleSheet(
                f'font-size:11.5px;color:{self.main.theme.MUTED};border:none;')
            cv.addWidget(num)
            cv.addWidget(lbl)
            self._stat_cards[key] = num
            self.stats_row.addWidget(card)
        v.addLayout(self.stats_row)

        self.filter_edit = QLineEdit()
        self.filter_edit.setPlaceholderText(
            '🔍 过滤台账（按功能名/编号/文件关键字）')
        self.filter_edit.setClearButtonEnabled(True)
        # 防抖: refresh() 会把台账目录下所有 JSON 重新打开解析再重建表格
        # (2500 条台账实测 450ms/键), 每敲一个字符同步跑一次会明显卡顿。
        # 攒 250ms 再刷新, 打字过程只会在停顿后刷一次。
        self._filter_timer = QTimer(self)
        self._filter_timer.setSingleShot(True)
        self._filter_timer.setInterval(250)
        self._filter_timer.timeout.connect(self.refresh)
        self.filter_edit.textChanged.connect(
            lambda _=None: self._filter_timer.start())
        v.addWidget(self.filter_edit)

        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(
            ('台账编号', '时间', '功能', '总数', '成功/失败/跳过', '耗时(秒)'))
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.doubleClicked.connect(lambda _=None: self.show_detail())
        v.addWidget(self.table, 1)

        brow = WrapRow()
        for text, cb in (('刷新', self.refresh), ('查看明细', self.show_detail),
                         ('导出选中任务 CSV', self.export_csv),
                         ('撤销上次操作', self.undo_last),
                         ('导出全部台账 Excel', self.export_all_excel)):
            b = QPushButton(text)
            b.clicked.connect(cb)
            brow.addWidget(b)
        brow.addStretch(1)
        v.addLayout(brow)
        self.refresh()

    def refresh(self):
        self.table.setRowCount(0)
        kw = (self.filter_edit.text().strip().lower()
              if hasattr(self, 'filter_edit') else '')
        from datetime import datetime, timedelta
        week_ago = (datetime.now() - timedelta(days=7)).strftime(
            '%Y-%m-%d %H:%M:%S')
        recs = []
        if os.path.isdir(self.main.RECORDS_DIR):
            for fn in sorted(os.listdir(self.main.RECORDS_DIR), reverse=True):
                if not fn.endswith('.json'):
                    continue
                try:
                    with open(os.path.join(self.main.RECORDS_DIR, fn),
                              encoding='utf-8') as fh:
                        r = json.load(fh)
                except Exception:
                    continue
                if not isinstance(r, dict):
                    continue
                recs.append(r)
        st_tasks = st_files = st_ok = st_fail = st_recent = 0
        for r in recs:
            blob = (str(r.get('title', '')) + ' ' + str(r.get('task_id', ''))).lower()
            files = r.get('files')
            if not isinstance(files, list):
                files = []
            file_blob = ' '.join(str(fr.get('path', ''))
                                 for fr in files
                                 if isinstance(fr, dict)).lower()
            if kw and kw not in blob and kw not in file_blob:
                continue
            st_tasks += 1
            st_files += _to_int(r.get('total'))
            st_ok += _to_int(r.get('ok'))
            st_fail += _to_int(r.get('fail'))
            if str(r.get('time', '')) >= week_ago:
                st_recent += 1
            row = self.table.rowCount()
            self.table.insertRow(row)
            vals = (str(r.get('task_id', '')), str(r.get('time', '')),
                    str(r.get('title', '')), str(r.get('total', '')),
                    f"{r.get('ok', 0)} / {r.get('fail', 0)} / "
                    f"{r.get('skipped', 0)}",
                    str(r.get('duration_sec', '')))
            for j, v in enumerate(vals):
                self.table.setItem(row, j, QTableWidgetItem(v))
        self._stat_cards['tasks'].setText(str(st_tasks))
        self._stat_cards['files'].setText(str(st_files))
        self._stat_cards['ok'].setText(str(st_ok))
        self._stat_cards['fail'].setText(str(st_fail))
        self._stat_cards['recent'].setText(str(st_recent))

    def _current_record(self):
        row = self.table.currentRow()
        if row < 0:
            return None
        task_id = self.table.item(row, 0).text()
        path = os.path.join(self.main.RECORDS_DIR, task_id + '.json')
        if not os.path.exists(path):
            return None
        with open(path, encoding='utf-8') as fh:
            return json.load(fh)

    def show_detail(self):
        r = self._current_record()
        if not r:
            QMessageBox.information(self, self.main.APP_NAME, '请先选择一条台账')
            return
        dlg = QDialog(self)
        dlg.setWindowTitle('台账明细 ' + str(r.get('task_id', '')))
        dlg.resize(840, 540)
        v = QVBoxLayout(dlg)
        head = QLabel(f"{r.get('time', '')}  {r.get('title', '')}  "
                      f"总数 {r.get('total', '')}  成功 {r.get('ok', 0)} "
                      f"失败 {r.get('fail', 0)} 跳过 {r.get('skipped', 0)}")
        head.setWordWrap(True)
        v.addWidget(head)
        files = _rec_files(r)
        t = QTableWidget(len(files), 4)
        t.setHorizontalHeaderLabels(('文件', '状态', '说明', '输出'))
        t.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        t.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        t.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        t.setEditTriggers(QTableWidget.NoEditTriggers)
        for i, fr in enumerate(files):
            t.setItem(i, 0, QTableWidgetItem(str(fr.get('path', ''))))
            t.setItem(i, 1, QTableWidgetItem(str(fr.get('status', ''))))
            t.setItem(i, 2, QTableWidgetItem(str(fr.get('detail', ''))))
            t.setItem(i, 3, QTableWidgetItem(str(fr.get('output', ''))))
        v.addWidget(t)
        b = QPushButton('关闭')
        b.clicked.connect(dlg.accept)
        v.addWidget(b)
        dlg.exec()

    def export_csv(self):
        r = self._current_record()
        if not r:
            QMessageBox.information(self, self.main.APP_NAME, '请先选择一条台账')
            return
        import csv as _csv
        out, _ = QFileDialog.getSaveFileName(
            self, '导出 CSV', str(r.get('task_id', 'task')) + '.csv',
            'CSV (*.csv)')
        if not out:
            return
        with open(out, 'w', encoding='utf-8-sig', newline='') as fh:
            w = _csv.writer(fh)
            w.writerow(('台账', str(r.get('task_id', ''))))
            w.writerow(('时间', str(r.get('time', ''))))
            w.writerow(('功能', str(r.get('title', ''))))
            w.writerow(())
            w.writerow(('文件', '状态', '说明', '输出'))
            for fr in _rec_files(r):
                w.writerow((fr.get('path', ''), fr.get('status', ''),
                            fr.get('detail', ''), fr.get('output', '')))
        self.main.show_status('已导出: ' + out)

    def export_all_excel(self):
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill
        recs = []
        if os.path.isdir(self.main.RECORDS_DIR):
            for fn in sorted(os.listdir(self.main.RECORDS_DIR), reverse=True):
                if not fn.endswith('.json'):
                    continue
                try:
                    with open(os.path.join(self.main.RECORDS_DIR, fn),
                              encoding='utf-8') as fh:
                        _r = json.load(fh)
                    # 台账目录里可能有手工编辑/外部工具写入的 JSON(数组或字符串),
                    # 直接追加会让后面 r.get() 抛 AttributeError, 导出整个崩掉
                    if isinstance(_r, dict):
                        recs.append(_r)
                except Exception:
                    continue
        if not recs:
            QMessageBox.information(self, self.main.APP_NAME, '还没有台账记录')
            return
        out, _ = QFileDialog.getSaveFileName(self, '导出全部台账 Excel',
                                             '青雁台账汇总.xlsx',
                                             'Excel (*.xlsx)')
        if not out:
            return
        wb = Workbook()
        ws = wb.active
        ws.title = '任务汇总'
        headers = ('台账编号', '时间', '功能', '总数', '成功', '失败', '跳过',
                   '耗时(秒)')
        ws.append(headers)
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=c)
            cell.font = Font(bold=True, color='FFFFFF')
            cell.fill = PatternFill('solid', fgColor='0F6E8C')
        for r in recs:
            ws.append((str(r.get('task_id', '')), str(r.get('time', '')),
                       str(r.get('title', '')), r.get('total', 0),
                       r.get('ok', 0), r.get('fail', 0),
                       r.get('skipped', 0), r.get('duration_sec', 0)))
        for col, wd in (('A', 30), ('B', 20), ('C', 22), ('D', 8), ('E', 8),
                        ('F', 8), ('G', 8), ('H', 10)):
            ws.column_dimensions[col].width = wd
        ws2 = wb.create_sheet('文件明细')
        h2 = ('台账编号', '任务功能', '文件', '状态', '说明', '输出')
        ws2.append(h2)
        for c in range(1, len(h2) + 1):
            cell = ws2.cell(row=1, column=c)
            cell.font = Font(bold=True, color='FFFFFF')
            cell.fill = PatternFill('solid', fgColor='0F6E8C')
        for r in recs:
            for fr in _rec_files(r):
                ws2.append((str(r.get('task_id', '')),
                            str(r.get('title', '')),
                            str(fr.get('path', '')),
                            str(fr.get('status', '')),
                            str(fr.get('detail', '')),
                            str(fr.get('output', ''))))
        for col, wd in (('A', 30), ('B', 22), ('C', 46), ('D', 8), ('E', 28),
                        ('F', 40)):
            ws2.column_dimensions[col].width = wd
        wb.save(out)
        self.main.show_status('已导出全部台账 Excel: ' + out)

    def undo_last(self):
        """撤销上次操作: 交换 undo 清单中的 src/dst。"""
        if not os.path.exists(self.main.UNDO_MANIFEST):
            QMessageBox.information(self, self.main.APP_NAME, '没有可撤销的操作')
            return
        try:
            with open(self.main.UNDO_MANIFEST, encoding='utf-8') as fh:
                m = json.load(fh)
        except Exception:
            QMessageBox.information(self, self.main.APP_NAME, '没有可撤销的操作')
            return
        actions = m.get('actions', [])
        if not actions:
            QMessageBox.information(self, self.main.APP_NAME, '没有可撤销的操作')
            return
        if getattr(self.main, '_task_running', False):
            self.main.show_status('请等待当前任务完成后再撤销')
            return
        ret = QMessageBox.question(
            self, self.main.APP_NAME,
            f'将撤销上次任务的 {len(actions)} 项文件改动。\n'
            '若原文件在任务完成后被手动修改过，该修改会被覆盖。\n继续吗？',
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if ret != QMessageBox.Yes:
            return
        from engine.undo import undo_actions
        from engine.store import _atomic_write
        completed, remaining, fails = undo_actions(actions)
        try:
            _atomic_write(self.main.UNDO_MANIFEST, {**m, 'actions': remaining})
        except OSError as exc:
            fails.append('撤销记录保存失败：' + str(exc))
        msg = f'已撤销 {len(completed)} 项'
        if fails:
            msg += f'；剩余 {len(remaining)} 项已保留，可重试。' + fails[0]
        self.main.show_status(msg)
        QMessageBox.information(self, self.main.APP_NAME, msg)
        self.refresh()


# ================================================================ 回收备份区
class TrashPage(QWidget):
    title = '回收站'

    def __init__(self, parent_main):
        super().__init__()
        self.main = parent_main
        self.manifest = None
        from engine.store import TrashManifest
        self.manifest = TrashManifest()
        v = QVBoxLayout(self)
        v.setContentsMargins(32, 26, 26, 16)
        v.setSpacing(10)
        cap = QLabel('回收备份区')
        cap.setStyleSheet(
            f'font-size:21px;font-weight:600;color:{self.main.theme.INK};')
        v.addWidget(cap)
        info = QLabel(f'删除的文件/文件夹保存在: {self.main.TRASH_DIR}'
                      '\n可逐条恢复到原位置; 彻底清除后不可恢复。')
        info.setWordWrap(True)
        info.setStyleSheet(f'color:{self.main.theme.MUTED};')
        v.addWidget(info)
        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(('备份文件名', '原路径', '删除时间'))
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionMode(QTableWidget.ExtendedSelection)
        v.addWidget(self.table, 1)
        brow = WrapRow()
        for text, cb in (('恢复选中到原位置', self.restore_sel),
                         ('打开备份目录', self.open_trash),
                         ('彻底清除选中', self.purge_sel),
                         ('刷新', self.refresh)):
            b = QPushButton(text)
            b.clicked.connect(cb)
            brow.addWidget(b)
        brow.addStretch(1)
        v.addLayout(brow)
        self.refresh()

    def refresh(self):
        self.table.setRowCount(0)
        data = self.manifest.load()
        stale = []
        for trash_name in sorted(data, reverse=True):
            meta = data.get(trash_name, {})
            if not os.path.exists(os.path.join(self.main.TRASH_DIR, trash_name)):
                # 清单有但文件已不在(被手动清理/被杀软删除): 顺带清理清单,
                # 否则这条"幽灵"既看不到也清不掉, 清单还会无限增长
                stale.append(trash_name)
                continue
            row = self.table.rowCount()
            self.table.insertRow(row)
            self.table.setItem(row, 0, QTableWidgetItem(trash_name))
            self.table.setItem(row, 1, QTableWidgetItem(str(meta.get('src', ''))))
            self.table.setItem(row, 2, QTableWidgetItem(str(meta.get('time', ''))))
        for nm in stale:
            self.manifest.remove(nm)

    def _sel_names(self):
        rows = set()
        for r in range(self.table.rowCount()):
            if self.table.item(r, 0) and self.table.item(r, 0).isSelected():
                rows.add(r)
        return sorted(rows)

    def restore_sel(self):
        import shutil
        rows = self._sel_names()
        if not rows:
            QMessageBox.information(self, self.main.APP_NAME,
                                    '请先选择要恢复的条目')
            return
        okn = failn = 0
        notes = []
        for r in rows:
            nm = self.table.item(r, 0).text()
            meta = self.manifest.load().get(nm, {})
            src = meta.get('src', '')
            dstp = os.path.join(self.main.TRASH_DIR, nm)
            # src 为空时既不算成功也不算失败, 用户只看到"成功 0, 失败 0"
            if not src:
                failn += 1
                notes.append(f'{nm}: 清单缺少原路径，无法恢复')
                continue
            try:
                parent = os.path.dirname(src)
                if parent:
                    os.makedirs(parent, exist_ok=True)
                dst = src
                if os.path.exists(dst):
                    base, ext = os.path.splitext(src)
                    dst = f'{base}.恢复_{os.path.getmtime(dstp):.0f}{ext}'
                    notes.append(f'{nm}: 原路径已占用，恢复为 {os.path.basename(dst)}')
                shutil.move(dstp, dst)
                self.manifest.remove(nm)
                okn += 1
            except Exception as e:
                failn += 1
                notes.append(f'{nm}: {e}')

        self.refresh()
        msg = f'恢复完成: 成功 {okn}, 失败 {failn}'
        if notes:
            msg += '\n\n' + '\n'.join(notes[:8])
        QMessageBox.information(self, self.main.APP_NAME, msg)
        self.refresh()

    def purge_sel(self):
        import shutil
        rows = self._sel_names()
        if not rows:
            QMessageBox.information(self, self.main.APP_NAME,
                                    '请先选择要彻底清除的条目')
            return
        ret = QMessageBox.question(
            self, self.main.APP_NAME,
            f'彻底清除 {len(rows)} 项? 此操作不可恢复。')
        if ret != QMessageBox.Yes:
            return
        for r in rows:
            nm = self.table.item(r, 0).text()
            p = os.path.join(self.main.TRASH_DIR, nm)
            try:
                if os.path.isdir(p):
                    shutil.rmtree(p, ignore_errors=True)
                else:
                    os.remove(p)
                self.manifest.remove(nm)
            except Exception:
                pass
        self.refresh()

    def open_trash(self):
        os.makedirs(self.main.TRASH_DIR, exist_ok=True)
        try:
            os.startfile(self.main.TRASH_DIR)
        except Exception:
            pass


# ================================================================ 任务流水线
class PipelinePage(QWidget):
    title = '任务流水线'

    def __init__(self, parent_main):
        super().__init__()
        self.main = parent_main
        self._steps = []
        v = QVBoxLayout(self)
        v.setContentsMargins(32, 26, 26, 16)
        v.setSpacing(10)
        cap = QLabel('任务流水线')
        cap.setStyleSheet(
            f'font-size:21px;font-weight:600;color:{self.main.theme.INK};')
        v.addWidget(cap)
        info = QLabel('把多个操作串成一条链：第一步处理完，其输出自动作为第二步的输入，'
                      '依次执行。\n添加步骤 → 选功能与操作 → 填该步参数与输出目录 → '
                      '全部就绪后点「执行流水线」。\nv3: 流水线可保存为命名方案, '
                      '随时载入。')
        info.setWordWrap(True)
        info.setStyleSheet(f'color:{self.main.theme.MUTED};')
        v.addWidget(info)

        self.steps_list = QListWidget()
        v.addWidget(self.steps_list, 1)

        brow = WrapRow()
        for text, cb in (('添加步骤', self.add_step),
                         ('编辑选中步骤', self.edit_step),
                         ('移除选中', self.remove_step),
                         ('上移', lambda: self.move(-1)),
                         ('下移', lambda: self.move(1))):
            b = QPushButton(text)
            b.clicked.connect(cb)
            brow.addWidget(b)
        brow.addStretch(1)
        v.addLayout(brow)

        saverow = QHBoxLayout()
        saverow.addWidget(QLabel('命名方案:'))
        self.pipe_combo = QComboBox()
        self._reload_pipelines()
        saverow.addWidget(self.pipe_combo, 1)
        saveb = QPushButton('保存当前流水线')
        saveb.clicked.connect(self._save_pipeline)
        saverow.addWidget(saveb)
        loadb = QPushButton('载入')
        loadb.clicked.connect(self._load_pipeline)
        saverow.addWidget(loadb)
        delb = QPushButton('删除方案')
        delb.clicked.connect(self._delete_pipeline)
        saverow.addWidget(delb)
        v.addLayout(saverow)

        self.run_btn = QPushButton('执行流水线')
        self.run_btn.setProperty('accent', True)
        self.run_btn.setStyleSheet(
            f'background:{self.main.theme.ACCENT};color:white;padding:9px 34px;'
            'font-weight:600;border:none;border-radius:9px;')
        self.run_btn.clicked.connect(self.run_pipeline)
        v.addWidget(self.run_btn)
        from app.pipeline import cancel
        self.cancel_pipeline_btn = QPushButton('取消流水线')
        self.cancel_pipeline_btn.setEnabled(False)
        self.cancel_pipeline_btn.clicked.connect(lambda: cancel(self))
        v.addWidget(self.cancel_pipeline_btn)

        self.status = QLabel('')
        self.status.setStyleSheet('font-weight:700;')
        v.addWidget(self.status)

    # ------------------------------------------------------------ 步骤管理
    def _show_page_in_dialog(self, page, title, note_text=None):
        """把功能页临时放进模态对话框编辑; 关闭后归还给主窗口堆栈。

        必须在结束后把 page 放回 stack 原位置, 否则对话框销毁时会连同
        页面一起销毁, 导致主界面崩溃(v3 修复)。
        """
        stack = self.main.stack
        idx = stack.indexOf(page)
        cur = stack.currentIndex()
        dlg = QDialog(self)
        dlg.setWindowTitle(title)
        dlg.resize(860, 720)
        dv = QVBoxLayout(dlg)
        hidden = [page._splitter.widget(1), page.action_bar]
        states = [widget.isHidden() for widget in hidden]
        for widget in hidden:widget.hide()
        dv.addWidget(page)
        page.show()
        if note_text:
            note = QLabel(note_text)
            note.setWordWrap(True)
            note.setStyleSheet(f'color:{self.main.theme.MUTED};')
            dv.addWidget(note)
        okb = QPushButton('确定')
        okb.clicked.connect(dlg.accept)
        dv.addWidget(okb)
        cancelb = QPushButton('取消')
        cancelb.clicked.connect(dlg.reject)
        dv.addWidget(cancelb)
        try:
            dlg.exec()
        finally:
            stack.insertWidget(idx, page)
            stack.setCurrentIndex(cur)
            for widget, was_hidden in zip(hidden, states):widget.setVisible(not was_hidden)
        return dlg

    def add_step(self):
        from PySide6.QtWidgets import QInputDialog
        # _page_defs 是 (title, cls) 二元组, 不是三元组。
        # 只收 FunctionPage 子类: 合并/台账/设置等页没有 _op_widget/processor。
        from app.pages.base import FunctionPage
        page_map = {}
        for _n, _cls in self.main._page_defs:
            if _cls is not None and issubclass(_cls, FunctionPage):
                page_map[_n] = _cls
        names = sorted(page_map)
        if not names:
            QMessageBox.information(self, self.main.APP_NAME,
                                    '没有可用的功能页')
            return
        name, ok = QInputDialog.getItem(self, '添加步骤 - 选择功能',
                                        '功能页:', names, 0, False)
        if not ok:
            return
        tidx = page_index(self.main._page_defs, name)
        if tidx < 0:
            QMessageBox.warning(self, self.main.APP_NAME,
                                f'未找到页面: {name}')
            return
        page = self.main.pages[tidx]
        combo = page._op_widget()
        ops_list = []
        if combo is not None:
            ops_list = [combo.itemText(i) for i in range(combo.count())]
        op_name = None
        if ops_list:
            op_name, ok2 = QInputDialog.getItem(self, '添加步骤 - 选择操作',
                                                '操作:', ops_list, 0, False)
            if not ok2:
                return
            pos = combo.findText(op_name)
            if pos >= 0:
                combo.setCurrentIndex(pos)
        dlg = self._show_page_in_dialog(
            page, f'步骤参数 — {name} / {op_name or "(无操作下拉)"}',
            '提示: 「文件/文件夹」可不填 — 若此步不是第一步，运行时自动使用'
            '上一步的输出目录；若是第一步则执行时询问。')
        # 对话框右上角 X = rejected, 早期不检查返回值, 用户"取消"了
        # 参数设置, 步骤照样被加进列表。
        if dlg is not None and hasattr(dlg, 'result') \
                and dlg.result() != QDialog.Accepted:
            return
        options = page.collect_options()
        options.pop('_rows', None)
        # ui_state: 控件级快照(下拉为显示名), 保证运行时能精确还原
        ui_state = page.capture_state()
        self._steps.append({'page': name, 'op': op_name or '',
                            'page_obj': page, 'options': options,
                            'ui_state': ui_state})
        self.steps_list.addItem(
            f'{len(self._steps)}. [{name}] {op_name or "—"}')

    def edit_step(self):
        row = self.steps_list.currentRow()
        if row < 0 or row >= len(self._steps):
            QMessageBox.information(self, self.main.APP_NAME,
                                    '请先选择要编辑的步骤')
            return
        step = self._steps[row]
        page = step.get('page_obj')
        if page is None or not hasattr(page, 'restore_state'):
            QMessageBox.warning(self, self.main.APP_NAME,
                                f'步骤 {row + 1} 找不到对应功能页'
                                f'(「{step.get("page")}」), 请移除后重新添加')
            return
        # 先把页面参数恢复成步骤记录的值, 编辑所见即所得
        before = page.capture_state()
        page.restore_state(step.get('ui_state') or step.get('options', {}))
        dlg = self._show_page_in_dialog(
            page, f"编辑步骤 {row + 1} — {step.get('page')} / {step.get('op')}")
        if dlg is not None and hasattr(dlg, 'result') and dlg.result() != QDialog.Accepted:
            page.restore_state(before)
            return
        step['options'] = page.collect_options()
        step['ui_state'] = page.capture_state()
        # 用户在对话框里改了操作, 标签和台账里的 op 也必须跟着改 ——
        # 早期只回写 options/ui_state, 列表仍显示旧操作, 让人以为没改成。
        _o = step['options']
        new_op = ''
        if isinstance(_o, dict):
            new_op = str(_o.get('op_name') or _o.get('op') or '')
        if new_op:
            step['op'] = new_op
        self.steps_list.item(row).setText(
            f"{row + 1}. [{step.get('page')}] {step.get('op') or '—'}")

    def remove_step(self):
        row = self.steps_list.currentRow()
        if row < 0:
            return
        self.steps_list.takeItem(row)
        if row < len(self._steps):
            self._steps.pop(row)
        for i in range(self.steps_list.count()):
            txt = self.steps_list.item(i).text()
            parts = txt.split('. ', 1)
            if len(parts) == 2:
                self.steps_list.item(i).setText(f'{i + 1}. {parts[1]}')

    def move(self, delta):
        row = self.steps_list.currentRow()
        n = self.steps_list.count()
        if row < 0 or n == 0:
            return
        new = max(0, min(n - 1, row + delta))
        if new == row:
            return
        item = self.steps_list.takeItem(row)
        self.steps_list.insertItem(new, item)
        self.steps_list.setCurrentRow(new)
        if 0 <= row < len(self._steps):
            s = self._steps.pop(row)
            self._steps.insert(max(0, min(len(self._steps), row + delta)), s)
        # 编号必须重排: 早期只挪控件, 列表会显示 "2. [B] 1. [A]"
        for i in range(self.steps_list.count()):
            txt = self.steps_list.item(i).text()
            parts = txt.split('. ', 1)
            if len(parts) == 2:
                self.steps_list.item(i).setText(f'{i + 1}. {parts[1]}')

    # ----------------------------------------------------------- 命名方案
    def _reload_pipelines(self):
        self.pipe_combo.clear()
        for name in sorted(ParamStore().load_pipelines()):
            self.pipe_combo.addItem(name)

    def _save_pipeline(self):
        from PySide6.QtWidgets import QInputDialog
        if not self._steps:
            QMessageBox.information(self, self.main.APP_NAME,
                                    '当前没有步骤可保存')
            return
        name, ok = QInputDialog.getText(self, '保存流水线方案', '方案名称:')
        if ok and name.strip():
            steps = [{'page': s.get('page'), 'op': s.get('op'),
                      'options': s.get('options', {}),
                      'ui_state': s.get('ui_state', {})} for s in self._steps]
            ParamStore().save_pipeline(name.strip(), steps)
            self._reload_pipelines()
            self.main.show_status('流水线方案已保存: ' + name.strip())

    def _load_pipeline(self):
        name = self.pipe_combo.currentText()
        if not name:
            return
        pl = ParamStore().load_pipelines().get(name, {})
        steps = pl.get('steps', [])
        self._steps = []
        self.steps_list.clear()
        for i, s in enumerate(steps, 1):
            # 找回对应页面实例, 步骤参数保留在字典里(运行时恢复)
            page = None
            try:
                idx = page_index(self.main._page_defs, s.get('page', ''))
                if idx >= 0:
                    page = self.main.pages[idx]
            except Exception:
                page = None
            self._steps.append({'page': s.get('page', ''),
                                'op': s.get('op', ''),
                                'page_obj': page,
                                'options': dict(s.get('options', {})),
                                'ui_state': dict(s.get('ui_state', {}))})
            self.steps_list.addItem(
                f"{i}. [{s.get('page')}] {s.get('op') or '—'}")
        if not self._steps:
            QMessageBox.information(self, self.main.APP_NAME, '该方案没有步骤')
        else:
            self.main.show_status(f'已载入流水线方案: {name} '
                                  f'({len(self._steps)} 个步骤)')

    def _delete_pipeline(self):
        name = self.pipe_combo.currentText()
        if name:
            ParamStore().delete_pipeline(name)
            self._reload_pipelines()

    # --------------------------------------------------------------- 执行
    def run_pipeline(self):
        from app.pipeline import start
        start(self)



    def _write_pipe_undo(self, actions):
        """把整条流水线的撤销动作合并写入清单(schema 与 BatchWorker 一致)。"""
        from datetime import datetime as _dt
        now = _dt.now()
        task_id = now.strftime('%Y%m%d_%H%M%S_') + format(
            now.microsecond // 1000, '03d')
        try:
            os.makedirs(os.path.dirname(self.main.UNDO_MANIFEST) or '.',
                        exist_ok=True)
            with open(self.main.UNDO_MANIFEST, 'w', encoding='utf-8') as fh:
                json.dump({'task_id': task_id, 'actions': actions},
                          fh, ensure_ascii=False, indent=1)
        except Exception:
            pass


def page_index(page_defs, title):
    for i, (name, _cls) in enumerate(page_defs):
        if name == title:
            return i
    return -1        # 原先兜底返回 0, 会把步骤静默绑到 pages[0](首页)


# ================================================================= 全局设置
class SettingsPage(QWidget):
    title = '全局设置'

    def __init__(self, parent_main):
        super().__init__()
        self.main = parent_main
        v = QVBoxLayout(self)
        v.setContentsMargins(32, 26, 26, 16)
        v.setSpacing(12)
        cap = QLabel('全局设置')
        cap.setStyleSheet(
            f'font-size:21px;font-weight:600;color:{self.main.theme.INK};')
        v.addWidget(cap)

        # v3: 外观
        form = QFormLayout()
        self.startup_combo = QComboBox()
        self.startup_combo.addItems(('首页', '任务台账', '办公模板库',
                                     '任务流水线'))
        ui = ParamStore().load_ui()
        last = ui.get('startup_page', '首页')
        if last in ('任务台账', '办公模板库', '任务流水线'):
            self.startup_combo.setCurrentText(last)
        form.addRow('启动时显示', self.startup_combo)
        self.startup_combo.currentIndexChanged.connect(self._save_startup)

        self.mode_combo = QComboBox()
        self.mode_combo.addItems(('浅色', '深色'))
        th = ParamStore().get_theme()
        self.mode_combo.setCurrentIndex(1 if th.get('mode') == 'dark' else 0)
        form.addRow('界面主题 (v3)', self.mode_combo)
        self.mode_combo.currentIndexChanged.connect(self._change_theme)

        self.accent_combo = QComboBox()
        from app.theme import ACCENT_PRESETS
        for name, _hex in ACCENT_PRESETS:
            self.accent_combo.addItem(name)
        cur = th.get('accent', '#635BDF')
        for i, (_n, hx) in enumerate(ACCENT_PRESETS):
            if hx.lower() == cur.lower():
                self.accent_combo.setCurrentIndex(i)
        form.addRow('强调色', self.accent_combo)
        self.accent_combo.currentIndexChanged.connect(self._change_theme)

        self.font_combo = QComboBox()
        self.font_combo.addItems(('标准', '大', '特大'))
        fs = th.get('font_scale', 1.0)
        self.font_combo.setCurrentIndex(0 if fs <= 1.0 else (1 if fs <= 1.12
                                                             else 2))
        form.addRow('字号', self.font_combo)
        self.font_combo.currentIndexChanged.connect(self._change_theme)
        v.addLayout(form)

        gb = QGroupBox('数据与日志 (全部本地)')
        gv = QVBoxLayout(gb)
        info = QLabel(
            '任务记录、收藏与参数均保存在本机。\n'
            f'{self.main.APP_DATA_DIR}\n'
            '卸载应用时保留这些数据，重新安装后可继续使用。')
        info.setWordWrap(True)
        info.setTextInteractionFlags(Qt.TextSelectableByMouse)
        gv.addWidget(info)
        brow = WrapRow()
        b1 = QPushButton('打开数据目录')
        b1.clicked.connect(lambda: os.startfile(self.main.APP_DATA_DIR))
        b2 = QPushButton('打开日志目录')
        b2.clicked.connect(lambda: os.startfile(
            os.path.join(self.main.APP_DATA_DIR, 'logs')))
        b3 = QPushButton('打开回收备份区')
        b3.clicked.connect(lambda: os.startfile(self.main.TRASH_DIR))
        brow.addWidget(b1)
        brow.addWidget(b2)
        brow.addWidget(b3)
        brow.addStretch(1)
        gv.addLayout(brow)
        v.addWidget(gb)

        gb2 = QGroupBox('参数方案管理')
        g2v = QVBoxLayout(gb2)
        self.preset_table = QTableWidget(0, 4)
        self.preset_table.setHorizontalHeaderLabels(
            ('方案名', '页面', '操作', '保存时间'))
        self.preset_table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.Stretch)
        self.preset_table.setEditTriggers(QTableWidget.NoEditTriggers)
        g2v.addWidget(self.preset_table)
        pb = WrapRow()
        for text, cb in (('删除选中方案', self._del_preset),
                         ('导出全部方案…', self._export_presets),
                         ('导入方案…', self._import_presets),
                         ('刷新', self.refresh_presets)):
            b = QPushButton(text)
            b.clicked.connect(cb)
            pb.addWidget(b)
        pb.addStretch(1)
        g2v.addLayout(pb)
        v.addWidget(gb2, 1)
        self.refresh_presets()

    def _change_theme(self, _idx=0):
        from app.theme import ACCENT_PRESETS
        mode = 'dark' if self.mode_combo.currentText() == '深色' else 'light'
        accent = ACCENT_PRESETS[self.accent_combo.currentIndex()][1]
        scale = (1.0, 1.12, 1.25)[self.font_combo.currentIndex()]
        self.main.apply_theme(mode, accent, scale)

    def _save_startup(self, _=None):
        try:
            ParamStore().save_ui({'startup_page':
                                  self.startup_combo.currentText()})
        except Exception:
            pass

    def refresh_presets(self):
        presets = ParamStore().load_presets()
        self.preset_table.setRowCount(0)
        for name, p in sorted(presets.items()):
            r = self.preset_table.rowCount()
            self.preset_table.insertRow(r)
            self.preset_table.setItem(r, 0, QTableWidgetItem(name))
            self.preset_table.setItem(r, 1, QTableWidgetItem(p.get('page', '')))
            self.preset_table.setItem(r, 2, QTableWidgetItem(p.get('op') or '—'))
            self.preset_table.setItem(r, 3, QTableWidgetItem(p.get('time', '')))

    def _del_preset(self):
        r = self.preset_table.currentRow()
        if r < 0:
            QMessageBox.information(self, self.main.APP_NAME,
                                    '请先选择要删除的方案')
            return
        name = self.preset_table.item(r, 0).text()
        ParamStore().delete_preset(name)
        self.refresh_presets()
        self.main.show_status('已删除方案: ' + name)

    def _export_presets(self):
        out, _ = QFileDialog.getSaveFileName(self, '导出全部方案',
                                             'qingyan_presets.json',
                                             'JSON (*.json)')
        if out:
            n = ParamStore().export_presets(out)
            QMessageBox.information(self, self.main.APP_NAME,
                                    f'已导出 {n} 个方案:\n{out}')

    def _import_presets(self):
        f, _ = QFileDialog.getOpenFileName(self, '导入方案', '', 'JSON (*.json)')
        if f:
            n = ParamStore().import_presets(f)
            self.refresh_presets()
            QMessageBox.information(self, self.main.APP_NAME, f'已导入 {n} 个方案')


# ================================================================= 帮助中心
class HelpPage(QWidget):
    title = '帮助中心'

    def __init__(self, parent_main):
        super().__init__()
        self.main = parent_main
        v = QVBoxLayout(self)
        v.setContentsMargins(32, 26, 26, 16)
        v.setSpacing(10)
        cap = QLabel('帮助中心')
        cap.setStyleSheet(
            f'font-size:21px;font-weight:600;color:{self.main.theme.INK};')
        v.addWidget(cap)
        tabs = QTabWidget()

        quick = QWidget()
        qv = QVBoxLayout(quick)
        qv.setSpacing(10)
        steps = QLabel(
            '只需三步，完成第一次批量处理：\n'
            '① 选功能 — 首页点一个功能卡片，或在搜索框输入关键词（Ctrl+F）\n'
            '② 选文件 — 页面里点「选文件/选文件夹」，或把文件直接拖进窗口\n'
            '    · 「展开文件清单」可预览将处理的文件；扩展名/大小/日期可过滤\n'
            '③ 点执行 — 参数区只显示当前操作需要的参数，每个操作都有白话说明；\n'
            '    · 执行前说明条会告诉你：结果输出到哪 / 是否修改原文件\n'
            '    · 点「执行任务」，底部实时显示进度、速度、预计剩余时间\n\n'
            '放心用，处处有后悔药：\n'
            '  · 就地修改自动保存到原文件旁“备份”文件夹\n'
            '  · 任务完成后的绿色快捷条里可「打开输出目录」或「撤销本次操作」\n'
            '  · 删除类操作先确认，且都进「回收站」可恢复\n'
            '  · 重命名可在「任务台账」里撤销\n\n'
            '效率技巧:\n'
            '  · 首页右键任一功能可加入收藏\n'
            '  · 「方案 ▾」把当前参数存为命名方案，下次一键载入\n'
            '  · 执行过的参数自动记忆，下次打开不用重填\n'
            '  · 「任务流水线」把多步操作串成链自动执行（可保存命名方案）\n'
            '  · 「全局设置」可切换深色主题、强调色和字号')
        steps.setWordWrap(True)
        steps.setStyleSheet(f'font-size:{self.main.theme.px(13)}px;')
        qv.addWidget(steps)
        qv.addStretch(1)
        tabs.addTab(quick, '快速入门')

        keys = QWidget()
        kv = QVBoxLayout(keys)
        ktable = QTableWidget(5, 2)
        ktable.setHorizontalHeaderLabels(('快捷键', '功能'))
        ktable.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        ktable.setEditTriggers(QTableWidget.NoEditTriggers)
        for r, (k, d) in enumerate((
                ('Ctrl+F', '聚焦首页搜索框'),
                ('Ctrl+1 ~ Ctrl+9, Ctrl+0', '跳转到第 1~10 个页面'),
                ('Ctrl+L', '打开任务台账'),
                ('Ctrl+Enter', '在功能页直接执行任务'),
                ('F1', '打开本帮助中心'))):
            ktable.setItem(r, 0, QTableWidgetItem(k))
            ktable.setItem(r, 1, QTableWidgetItem(d))
        kv.addWidget(ktable)
        tabs.addTab(keys, '快捷键')

        faq = QWidget()
        fv = QVBoxLayout(faq)
        faql = QLabel(
            'Q: 需要联网吗？\n'
            'A: 完全不需要。安装、使用全程离线，不发送任何数据。\n\n'
            'Q: 为什么有些文件显示「跳过」？\n'
            'A: 跳过 = 无需处理（如未匹配到内容、已是目标编码），不是错误。\n\n'
            'Q: 处理会破坏原文件吗？\n'
            'A: 选「输出目录」则原文件不动；就地修改自动保存到原文件旁“备份”文件夹；'
            '重命名/删除可撤销或从回收备份区恢复。\n\n'
            'Q: 格式转换的结果能编辑吗？\n'
            'A: 能。所有转换产物都是真格式文件（不是截图）：'
            'PDF→Word 得到每段可编辑的 .docx；Office→PDF 用本机 Office 渲染，'
            '版式与原文件一致。产物默认放在源文件旁的「_转换结果」文件夹，'
            '原文件不动。\n\n'
            'Q: PDF 转 Word 后是图片吗？\n'
            'A: 不是。转换提取 PDF 里的文字层，重建为真正的 Word 段落'
            '（标题层级、加粗、居中都保留）。扫描件（纯图片 PDF）没有文字层，'
            '会明确提示不支持。\n\n'
            'Q: 支持 .doc/.xls/.ppt 老格式吗？\n'
            'A: 请先在 Office 中另存为 .docx/.xlsx/.pptx。\n\n'
            'Q: 参数每次都要重填吗？\n'
            'A: 不用。执行过的参数自动记忆；也可用「方案 ▾」保存为命名方案随时载入。')
        faql.setWordWrap(True)
        faql.setStyleSheet(f'font-size:{self.main.theme.px(13)}px;')
        fv.addWidget(faql)
        fv.addStretch(1)
        tabs.addTab(faq, '常见问题')

        data_tab = QWidget()
        dv = QVBoxLayout(data_tab)
        datal = QLabel(
            '全部数据保存在本机（零联网）:\n\n'
            f'  任务台账: {self.main.RECORDS_DIR}\n'
            f'  回收备份区: {self.main.TRASH_DIR}\n'
            f'  收藏夹/参数方案/窗口状态: {self.main.APP_DATA_DIR}\n'
            '  就地修改备份: 原文件旁“备份”文件夹\n\n'
            '备份建议: 定期复制上述数据目录即可完整备份使用习惯与历史台账。\n'
            '卸载软件不会删除数据目录；彻底清理请手动删除。')
        datal.setWordWrap(True)
        dv.addWidget(datal)
        openb = QPushButton('打开数据目录')
        openb.clicked.connect(lambda: os.startfile(self.main.APP_DATA_DIR))
        dv.addWidget(openb)
        dv.addStretch(1)
        tabs.addTab(data_tab, '数据与备份')

        v.addWidget(tabs, 1)


# ==================================================================== 关于
class AboutPage(QWidget):
    title = '关于'

    def __init__(self, parent_main=None):
        super().__init__()
        v = QVBoxLayout(self)
        v.setContentsMargins(32, 26, 26, 16)
        cap = QLabel(f'青雁批量工具箱 v{parent_main.APP_VERSION if parent_main else "3.4.0"}')
        cap.setStyleSheet('font-size:22px;font-weight:700;color:#0F6E8C;')
        v.addWidget(cap)
        from engine.store import data_dir
        body = QLabel(
            '一款完全免费、完全离线的桌面批量处理工具箱。\n\n承诺:\n'
            '  · 永久免费 — 无会员、无付费墙、无广告\n'
            '  · 完全离线 — 无账号、无授权服务器、不发送任何数据\n'
            '  · 本地台账 — 处理记录只保存在您自己的电脑上\n'
            '  · 安全可回退 — 预览、备份、撤销机制保护您的文件\n'
            '  · 全功能 — Word/Excel/PPT/PDF/文本/图片/重命名/文件整理 '
            '169 项操作，以及合并、模板与流水线\n\n'
            f'数据目录: {data_dir()}\n'
            f'任务台账: {os.path.join(data_dir(), "records")}\n\n'
            '常用工具可以收藏到首页；外观、字号和启动页面可在全局设置调整。\n'
            '处理完成后，可在任务台账查看每个文件的结果。')
        body.setWordWrap(True)
        v.addWidget(body)
        v.addStretch(1)


# ================================================================ 办公模板库
class OfficeTemplatesPage(QWidget):
    title = '办公模板库'

    def __init__(self, parent_main):
        super().__init__()
        self.main = parent_main
        from engine.office_ops import list_office_templates
        v = QVBoxLayout(self)
        v.setContentsMargins(32, 26, 26, 16)
        v.setSpacing(10)
        cap = QLabel('办公模板库')
        cap.setStyleSheet(
            f'font-size:21px;font-weight:600;color:{self.main.theme.INK};')
        v.addWidget(cap)
        info = QLabel('选择一套内置办公模板 → 一键生成模板文件与示例数据 → '
                      '自动跳到「模板批量生成」页，把 Excel 里的名单换成你的真实数据'
                      '即可批量产出正式文档。全离线、字段可改。')
        info.setWordWrap(True)
        info.setStyleSheet(f'color:{self.main.theme.MUTED};')
        v.addWidget(info)
        form = QFormLayout()
        self.tpl_combo = QComboBox()
        self.tpl_combo.addItems(list_office_templates())
        form.addRow('选择模板', self.tpl_combo)
        self.out_edit = QLineEdit()
        self.out_edit.setPlaceholderText(
            '模板输出目录(默认: 数据目录\\office_templates)')
        form.addRow('输出目录', self.out_edit)
        ob = QPushButton('选目录')
        ob.clicked.connect(self._pick_out)
        form.addRow('', ob)
        v.addLayout(form)
        self.desc = QFrame()
        self.desc.setStyleSheet(
            f'background:{self.main.theme.PAPER2};'
            f'border:1px solid {self.main.theme.LINE};padding:14px;'
            f'font-size:13px;')
        dvl = QVBoxLayout(self.desc)
        self.desc_label = QLabel('')
        self.desc_label.setWordWrap(True)
        dvl.addWidget(self.desc_label)
        v.addWidget(self.desc)
        gen = QPushButton('生成模板 + 示例数据 → 去批量生成')
        gen.setStyleSheet(
            f'background:{self.main.theme.ACCENT};color:white;padding:9px 28px;'
            'font-weight:600;border:none;border-radius:9px;')
        gen.clicked.connect(self._generate)
        v.addWidget(gen)
        self.status = QLabel('')
        self.status.setStyleSheet('font-weight:700;')
        v.addWidget(self.status)
        v.addStretch(1)
        self.tpl_combo.currentIndexChanged.connect(lambda _=None: self._show_desc())
        self._show_desc()

    def _pick_out(self):
        d = QFileDialog.getExistingDirectory(self, '选择模板输出目录')
        if d:
            self.out_edit.setText(d)

    def _show_desc(self):
        from engine.office_ops import OFFICE_TEMPLATES
        name = self.tpl_combo.currentText()
        spec = OFFICE_TEMPLATES.get(name, {})
        fields = '、'.join(spec.get('columns', ()))
        fmt = ('Word 文档 (.docx)' if spec.get('ext') == '.docx'
               else 'Excel 表格 (.xlsx)')
        rows = len(spec.get('sample', ()))
        self.desc_label.setText(
            f'输出格式：{fmt}\n包含字段：{fields}\n示例数据：{rows} 行'
            '（生成后可在 Excel 中替换为真实名单）')

    def _generate(self):
        from engine.office_ops import materialize_office_template
        out_dir = self.out_edit.text().strip() or os.path.join(
            os.environ.get('APPDATA', os.path.expanduser('~')),
            'QingYanBatchToolbox', 'office_templates')
        name = self.tpl_combo.currentText()
        try:
            tpl_path, data_path, columns = materialize_office_template(
                name, out_dir)
        except Exception as e:
            QMessageBox.warning(self, self.main.APP_NAME, '生成失败: ' + str(e))
            return
        self.status.setText(f'✔ 已生成: {tpl_path}\n✔ 示例数据: {data_path}')
        tidx = page_index(self.main._page_defs, '模板批量生成')
        if tidx < 0:
            self.main.show_status('模板已生成, 但未找到「模板批量生成」页')
            return
        self.main.nav.setCurrentRow(tidx)
        tpage = self.main.pages[tidx]
        tpage._widgets['tpl_edit'].setText(tpl_path)
        tpage._widgets['data_edit'].setText(data_path)
        tpage._widgets['name_col'].setText(columns[0] if columns else '')
        os.makedirs(out_dir, exist_ok=True)
        tpage._widgets['out_dir'].setText(out_dir)
        self.main.show_status(
            f'模板「{name}」已就绪 — 检查数据后点「开始处理」即可批量生成')
