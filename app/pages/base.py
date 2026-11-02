# -*- coding: utf-8 -*-
"""青雁批量工具箱 v3 — 功能页基类。

在 v2 基础上的改进:
- 文件清单面板: 展开可见已解析文件(名称/大小), 右键可打开/移除;
- 结果区汇总条: 成功/失败/跳过实时统计;
- 执行后自动定位失败项, 「只看失败」过滤;
- ETA 与速度显示(配合主窗口底部任务条);
- 参数自动记忆 + 命名方案 + 重置;
- Ctrl+Enter 快捷执行。
"""
from __future__ import annotations

import os

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QColor, QGuiApplication
from PySide6.QtWidgets import (QCheckBox, QComboBox, QDialog, QDoubleSpinBox,
                               QFileDialog, QFormLayout, QFrame, QGroupBox,
                               QHBoxLayout, QHeaderView, QInputDialog, QLabel,
                               QLineEdit, QListWidget, QListWidgetItem, QMenu,
                               QMessageBox, QPlainTextEdit, QPushButton,
                               QScrollArea, QSpinBox, QSplitter,
                               QTableWidget, QTableWidgetItem, QVBoxLayout,
                               QWidget)

from engine.core import BatchWorker, FileResult, TaskSpec
from engine.store import ParamStore
from app.layouts import WrapRow


class FunctionPage(QWidget):
    """功能页基类。子类定义: title/kind/desc/OPS/OP_FIELDS/REQUIRED_FIELDS,
    并实现 build_options(form) / collect_options() / processor()。"""

    title = '功能'
    kind = 'generic'
    group = ''
    desc = ''
    OPS: tuple = ()
    OP_FIELDS: dict = {}
    REQUIRED_FIELDS: dict = {}
    ALWAYS_VISIBLE: tuple = ()
    EXPECTED_EXTS: tuple = ()   # 非空时, 扩展名不符的文件优雅跳过(v3)
    # False = 该页操作不改动源文件(转换/打印/只读查询), 执行前说明条
    # 不再谎称"将直接修改原文件"
    MODIFIES_SOURCE: bool = True

    def __init__(self, parent_main):
        super().__init__()
        self.main = parent_main
        self.worker = None
        self._widgets: dict = {}
        self._field_rows: dict = {}
        self._noparam_hint = None
        self._extra: dict = {}      # 页面级缓存(规则表/计划等)
        self._summary_counts = {'ok': 0, 'fail': 0, 'skipped': 0}
        self.setObjectName("page_" + self.kind)
        self._build_ui()

    # ================================================================ UI 骨架
    def _build_ui(self):
        th = self.main.theme
        # 拖放: Qt 只在 acceptDrops 为真时才投递 drag/drop 事件。少了这一行,
        # 下面重写的 dragEnterEvent/dropEvent 永远不会被调用, 用户按提示
        # 拖文件进来毫无反应。
        self.setAcceptDrops(True)
        root = QVBoxLayout(self)
        root.setContentsMargins(26, 22, 26, 18)
        root.setSpacing(14)

        # ------------------------------------------------ 左栏: 输入与参数
        left_container = QWidget()
        left = QVBoxLayout(left_container)
        left.setContentsMargins(0, 0, 0, 0)
        left.setSpacing(10)

        cap = QLabel(self.title)
        cap.setStyleSheet(
            f'font-size:{th.px(21)}px;font-weight:600;color:{th.INK};')
        root.addWidget(cap)
        if self.desc:
            tip = QLabel(self.desc)
            tip.setWordWrap(True)
            tip.setStyleSheet(f'color:{th.MUTED};font-size:{th.px(12)}px;')
            root.addWidget(tip)

        # 文件选择卡片
        file_box = QGroupBox('01  添加文件')
        fv = QVBoxLayout(file_box)
        fv.setContentsMargins(12, 6, 12, 10)
        fv.setSpacing(8)
        row1 = QHBoxLayout()
        self.file_edit = QLineEdit()
        self.file_edit.setObjectName(self.kind + '_input')
        self.file_edit.setAccessibleName('输入文件或文件夹')
        self.file_edit.setPlaceholderText('选择文件或文件夹路径(支持拖入)')
        self.file_edit.textChanged.connect(self._on_files_changed)
        b1 = QPushButton('添加文件')
        b1.setToolTip('从电脑里挑几个文件(支持按住 Ctrl 多选)')
        b1.clicked.connect(self.pick_files)
        b2 = QPushButton('添加文件夹')
        b2.setToolTip('选中整个文件夹(可配合「包含子文件夹」)')
        b2.clicked.connect(self.pick_dir)
        row1.addWidget(self.file_edit, 1)
        fv.addLayout(row1)
        row1 = QHBoxLayout()
        row1.addWidget(b1)
        row1.addWidget(b2)
        row1.addStretch(1)
        fv.addLayout(row1)

        row2 = QHBoxLayout()
        self.recursive_cb = QCheckBox('包含子文件夹')
        self.recursive_cb.toggled.connect(self._on_files_changed)
        self.ext_edit = QLineEdit()
        self.ext_edit.setObjectName(self.kind + '_extensions')
        self.ext_edit.setPlaceholderText('扩展名过滤(逗号分隔, 留空=全部)')
        self.ext_edit.textChanged.connect(self._on_files_changed)
        row2.addWidget(self.recursive_cb)
        row2.addWidget(QLabel('扩展名:'))
        row2.addWidget(self.ext_edit, 1)
        fv.addLayout(row2)

        row3 = QHBoxLayout()
        self.filter_min_kb = QSpinBox()
        self.filter_min_kb.setRange(0, 1048576)
        self.filter_min_kb.setValue(0)
        self.filter_min_kb.setToolTip('只处理大于该值的文件, 0=不限')
        self.filter_date_from = QLineEdit()
        self.filter_date_from.setPlaceholderText('如 2026-01-01')
        self.filter_date_to = QLineEdit()
        self.filter_date_to.setPlaceholderText('如 2026-12-31')
        self.filter_min_kb.valueChanged.connect(self._on_files_changed)
        self.filter_date_from.textChanged.connect(self._on_files_changed)
        self.filter_date_to.textChanged.connect(self._on_files_changed)
        row3.addWidget(QLabel('最小KB:'))
        row3.addWidget(self.filter_min_kb)
        row3.addWidget(QLabel('日期从:'))
        row3.addWidget(self.filter_date_from)
        row3.addWidget(QLabel('到:'))
        row3.addWidget(self.filter_date_to)
        filters = QWidget()
        filters.setLayout(row3)
        filters.hide()
        filter_toggle = QPushButton('更多筛选')
        filter_toggle.setCheckable(True)
        filter_toggle.toggled.connect(filters.setVisible)
        row2.addWidget(filter_toggle)
        fv.addWidget(filters)

        row4 = QHBoxLayout()
        self.count_label = QLabel('已选 0 个文件')
        self.count_label.setStyleSheet(f'color:{th.MUTED};')
        self.expand_btn = QPushButton('展开文件清单')
        self.expand_btn.setCheckable(True)
        self.expand_btn.setFlat(True)
        self.expand_btn.setCursor(Qt.PointingHandCursor)
        self.expand_btn.toggled.connect(self._toggle_file_list)
        row4.addWidget(self.count_label, 1)
        row4.addWidget(self.expand_btn)
        fv.addLayout(row4)

        self.file_list = QListWidget()
        self.file_list.setVisible(False)
        self.file_list.setMaximumHeight(180)
        self.file_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.file_list.customContextMenuRequested.connect(self._file_menu)
        fv.addWidget(self.file_list)
        left.addWidget(file_box)

        # 参数卡片
        form_box = QGroupBox('02  设置处理方式')
        self.form = QFormLayout(form_box)
        self.form.setLabelAlignment(Qt.AlignLeft)
        self.form.setSpacing(12)
        self.form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        self.form.setRowWrapPolicy(QFormLayout.WrapLongRows)
        self.build_options(self.form)
        self._setup_op_fields()
        left.addWidget(form_box)

        # v3: 执行前说明条 — 用大白话告诉用户接下来会发生什么
        self.run_hint = QLabel('')
        self.run_hint.setWordWrap(True)
        self.run_hint.setStyleSheet(
            f'color:{th.MUTED};font-size:12px;padding:6px 10px;'
            f'background:{th.PAPER2};border:1px solid {th.LINE};'
            'border-radius:8px;')
        self.run_hint.setVisible(False)
        left.addWidget(self.run_hint)
        left.addStretch(1)

        # 动作条
        btn_row = WrapRow()
        self.preview_btn = QPushButton('预览处理结果')
        self.preview_btn.setObjectName(self.kind + '_preview')
        self.preview_btn.clicked.connect(self.do_preview)
        self.run_btn = QPushButton('开始处理')
        self.run_btn.setObjectName(self.kind + '_run')
        self.run_btn.setProperty('accent', True)
        self.run_btn.setStyleSheet(
            f'background:{th.ACCENT};color:white;padding:9px 34px;'
            f'font-weight:600;letter-spacing:1px;border:none;border-radius:9px;')
        self.run_btn.clicked.connect(self.do_run)
        self.cancel_btn = QPushButton('取消任务')
        self.cancel_btn.setObjectName(self.kind + '_cancel')
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.clicked.connect(self.do_cancel)
        preset_btn = QPushButton('方案')
        preset_btn.setToolTip('保存当前参数为命名方案 / 载入 / 管理 / 导出导入')
        preset_menu = QMenu(self)
        preset_menu.addAction('保存当前参数为方案…', self._save_preset_dialog)
        preset_menu.addAction('载入方案…', self._load_preset_dialog)
        preset_menu.addAction('管理方案…', self._manage_presets)
        preset_menu.addSeparator()
        preset_menu.addAction('导出全部方案…', self._export_presets)
        preset_menu.addAction('导入方案…', self._import_presets)
        preset_menu.addSeparator()
        preset_menu.addAction('重置当前参数', self._reset_params)
        preset_btn.setMenu(preset_menu)
        btn_row.addWidget(self.preview_btn)
        btn_row.addWidget(self.run_btn)
        btn_row.addWidget(self.cancel_btn)
        btn_row.addWidget(preset_btn)
        btn_row.addStretch(1)
        self.open_out_cb = QCheckBox('完成后打开输出目录')
        self.open_out_cb.setToolTip('任务结束后自动在资源管理器中打开输出目录')
        extra_row = QHBoxLayout()
        extra_row.addWidget(self.open_out_cb)
        self.fail_only_cb = QCheckBox('只看失败')
        self.fail_only_cb.setToolTip('勾选后结果表只显示失败项, 便于快速定位')
        self.fail_only_cb.toggled.connect(self._filter_fail_only)
        extra_row.addWidget(self.fail_only_cb)
        extra_row.addStretch(1)
        # Execution controls remain outside the scrolling options.

        scroll = QScrollArea()
        scroll.setObjectName(self.kind + '_optionsScroll')
        scroll.setWidget(left_container)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        # 左栏允许横向滚动: 拖窄分割条时内容可滚动而不是被裁掉
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll.setMinimumWidth(330)
        scroll.setMinimumHeight(160)
        scroll.viewport().setAutoFillBackground(False)

        # ------------------------------------------------ 右栏: 执行结果
        right = QWidget()
        right.setMinimumWidth(270)
        right.setMinimumHeight(150)
        rv = QVBoxLayout(right)
        rv.setContentsMargins(0, 0, 0, 0)
        rv.setSpacing(8)
        rcap = QHBoxLayout()
        rtitle = QLabel('执行结果')
        rtitle.setProperty('role', 'section')
        rtitle.setStyleSheet(
            f'font-size:{th.px(15)}px;font-weight:600;color:{th.INK};')
        self.summary_label = QLabel('')
        self.summary_label.setStyleSheet(f'color:{th.MUTED};')
        rcap.addWidget(rtitle)
        rcap.addStretch(1)
        rcap.addWidget(self.summary_label)
        rv.addLayout(rcap)

        # v3: 完成后快捷条 — 打开输出目录 / 撤销本次操作
        self.result_banner = QFrame()
        self.result_banner.setStyleSheet(
            f'QFrame{{background:{th.PAPER2};border:1px solid {th.LINE};'
            f'border-left:4px solid {th.ACCENT};border-radius:8px;}}')
        bl = QHBoxLayout(self.result_banner)
        bl.setContentsMargins(12, 8, 12, 8)
        self.banner_label = QLabel('')
        self.banner_label.setWordWrap(True)
        self.banner_label.setStyleSheet('border:none;font-weight:600;')
        bl.addWidget(self.banner_label, 1)
        self.banner_open_btn = QPushButton('打开输出目录')
        self.banner_open_btn.setCursor(Qt.PointingHandCursor)
        self.banner_open_btn.clicked.connect(self._banner_open)
        bl.addWidget(self.banner_open_btn)
        self.banner_undo_btn = QPushButton('撤销本次操作')
        self.banner_undo_btn.setCursor(Qt.PointingHandCursor)
        self.banner_undo_btn.setToolTip(
            '把本次任务改动过的文件恢复到执行前的状态')
        self.banner_undo_btn.clicked.connect(self._banner_undo)
        bl.addWidget(self.banner_undo_btn)
        banner_close = QPushButton('✕')
        banner_close.setAccessibleName('关闭结果提示')
        banner_close.setObjectName(self.kind + '_dismissResult')
        banner_close.setProperty('compact', True)
        banner_close.setStyleSheet('padding:0;')
        banner_close.setFlat(True)
        banner_close.setFixedWidth(28)
        banner_close.clicked.connect(lambda: self.result_banner.hide())
        bl.addWidget(banner_close)
        self.result_banner.hide()
        self._banner_outputs: list = []
        self._banner_undo_actions: list = []
        rv.addWidget(self.result_banner)

        self.result_table = QTableWidget(0, 3)
        self.result_table.setHorizontalHeaderLabels(('文件', '状态', '说明 / 输出'))
        self.result_table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.Stretch)
        self.result_table.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeToContents)
        self.result_table.horizontalHeader().setSectionResizeMode(
            2, QHeaderView.Stretch)
        self.result_table.verticalHeader().setVisible(False)
        self.result_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.result_table.setAlternatingRowColors(True)
        self.result_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.result_table.customContextMenuRequested.connect(self._result_menu)
        rv.addWidget(self.result_table, 1)

        # 空态提示: 必须挂到表格 viewport 上(否则会成为独立顶层窗口,
        # 在屏幕左上角出现黑色悬浮框 — v3.0.0 缺陷, 已修复)
        self.result_empty = QLabel(
            '还没有执行记录\n\n选好文件与参数后点「开始处理」，\n'
            '每个文件的处理结果会逐行显示在这里', self.result_table.viewport())
        self.result_empty.setAlignment(Qt.AlignCenter)
        self.result_empty.setWordWrap(True)
        self.result_empty.setContentsMargins(12, 12, 12, 12)
        self.result_empty.setAttribute(Qt.WA_TransparentForMouseEvents)
        self.result_empty.setStyleSheet(
            f'color:{th.MUTED};font-size:{th.px(13)}px;background:transparent;')
        self.result_empty.hide()
        self.result_table.viewport().installEventFilter(self)

        # ------------------- 左右分栏: 可拖动调节, 系统自适应并记住偏好
        # (v3 修复: 原来"左参数固定宽 + 右结果占满", 结果区空白偏大
        #  且无法调整; 现在用 QSplitter, 拖动手柄即可分配两侧宽度)
        self._splitter = QSplitter(Qt.Horizontal)
        self._splitter.setHandleWidth(8)
        self._splitter.setChildrenCollapsible(False)
        self._splitter.addWidget(scroll)     # 0: 左侧 输入与参数
        self._splitter.addWidget(right)      # 1: 右侧 执行结果
        self._splitter.setStretchFactor(0, 0)
        self._splitter.setStretchFactor(1, 1)
        self._splitter.setStyleSheet(
            'QSplitter::handle { background: transparent; }'
            f'QSplitter::handle:hover {{ background: {th.ACCENT_SOFT}; }}')
        root.addWidget(self._splitter, 1)
        actions = QFrame()
        actions.setProperty('actionBar', True)
        actions_layout = QVBoxLayout(actions)
        actions_layout.setContentsMargins(16, 12, 16, 12)
        actions_layout.addLayout(btn_row)
        actions_layout.addLayout(extra_row)
        root.addWidget(actions)
        self.action_bar = actions
        # 初始分配: 左侧给足参数区宽度, 其余给结果; 之后随窗口缩放自适应
        self._splitter.setSizes([600, 420])
        # 恢复用户上次拖动的宽度
        try:
            saved = (ParamStore().load_ui().get('splitters') or {}).get(
                self.kind)
            if (isinstance(saved, list) and len(saved) == 2
                    and all(isinstance(x, int) for x in saved)
                    and sum(saved) > 400):
                QTimer.singleShot(0, lambda: self._splitter.setSizes(saved))
        except Exception:
            pass

        self.result_table.setObjectName(self.kind + '_results')
        self.banner_label.setObjectName(self.kind + '_resultSummary')
        self._update_result_empty()
        # 恢复上次参数(自动记忆)
        self._restore_autosave()
        self._update_run_hint()

    # ----------------------------------------------------------- 文件区
    def pick_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, '选择文件')
        if files:
            self._excluded_selection = ('', set())
            self._files_immediate = True
            try:
                self.file_edit.setText('||'.join(files))
            finally:
                self._files_immediate = False
            self._refresh_file_stats()

    def pick_dir(self):
        d = QFileDialog.getExistingDirectory(self, '选择文件夹')
        if d:
            self._excluded_selection = ('', set())
            self._files_immediate = True
            try:
                self.file_edit.setText(d)
            finally:
                self._files_immediate = False
            self._refresh_file_stats()

    def _toggle_file_list(self, on: bool):
        self.expand_btn.setText('收起文件清单' if on else '展开文件清单')
        self.file_list.setVisible(on)
        if on:
            self._refresh_file_list()

    def _refresh_file_list(self):
        self.file_list.clear()
        try:
            files = self.current_files()
        except Exception:
            files = []
        for f in files:
            try:
                size = os.path.getsize(f)
                text = f'{os.path.basename(f)}   ({size / 1024:.0f} KB)'
            except OSError:
                text = os.path.basename(f)
            item = QListWidgetItem(text)
            item.setToolTip(f)
            self.file_list.addItem(item)

    def _file_menu(self, pos):
        item = self.file_list.itemAt(pos)
        if not item:
            return
        path = item.toolTip()
        menu = QMenu(self)
        menu.addAction('打开文件', lambda: self.main.open_path(path))
        menu.addAction('打开所在位置',
                       lambda: self.main.open_path(os.path.dirname(path)))
        menu.addSeparator()
        menu.addAction('复制路径', lambda: QGuiApplication.clipboard().setText(path))
        menu.addSeparator()
        menu.addAction('从本次选择中移除', lambda: self._remove_selected_file(path))
        menu.exec(self.file_list.mapToGlobal(pos))

    def _remove_selected_file(self, path):
        raw = self.file_edit.text()
        previous, excluded = getattr(self, '_excluded_selection', ('', set()))
        excluded = set(excluded) if previous == raw else set()
        excluded.add(os.path.normcase(os.path.abspath(path)))
        self._excluded_selection = (raw, excluded)
        self._refresh_file_stats()

    def _on_files_changed(self):
        previous, _ = getattr(self, '_excluded_selection', ('', set()))
        if previous != self.file_edit.text():
            self._excluded_selection = ('', set())
        # 防抖: 输入框每敲一个字符都会触发, 而 current_files() 要遍历整个
        # 目录(还可能 stat 每个文件)。大目录里逐字符全量扫描会把主线程
        # 卡住。延迟到用户停止输入后再算。
        if getattr(self, '_files_timer', None) is None:
            from PySide6.QtCore import QTimer
            self._files_timer = QTimer(self)
            self._files_timer.setSingleShot(True)
            self._files_timer.setInterval(280)
            self._files_timer.timeout.connect(self._refresh_file_stats)
        self._files_timer.start()
        # 选文件/选文件夹/拖放这类"一次性确定"的变更需要立刻反映到计数上,
        # 否则用户点完按钮要等 0.3 秒才看到数字
        if getattr(self, '_files_immediate', False):
            self._refresh_file_stats()

    def _refresh_file_stats(self):
        files = self.current_files()      # 只算一次, 不再重复遍历目录
        n = len(files)
        total = 0
        for f in files:
            try:
                total += os.path.getsize(f)
            except OSError:
                pass
        self.count_label.setText(
            f'已选 {n} 个文件' + (f' · 共 {total / 1048576:.1f} MB' if total else ''))
        # v3: 首次选到文件时自动展开文件清单, 让新手直观看到将处理什么
        if n > 0 and not self.expand_btn.isChecked() \
                and not getattr(self, '_auto_expanded', False):
            self._auto_expanded = True
            self.expand_btn.setChecked(True)
        elif n > 0 and self.expand_btn.isChecked():
            self._refresh_file_list()
        elif n == 0:
            self.file_list.clear()
        if hasattr(self, 'run_hint'):
            self._update_run_hint()

    def current_files(self):
        from engine.core import collect_files
        raw = self.file_edit.text().strip()
        if not raw:
            return []
        paths = [p.strip() for p in raw.split('||') if p.strip()]
        exts = None
        ext_text = self.ext_edit.text().replace('，', ',').strip()
        if ext_text:
            exts = [e.strip().lstrip('*.') for e in ext_text.split(',')
                    if e.strip()]
        files = []
        for p in paths:
            try:
                if os.path.isfile(p):
                    # 手动指定的文件不做扩展名过滤(用户明确选择了它)
                    files.append(p)
                else:
                    files.extend(collect_files(p, self.recursive_cb.isChecked(),
                                               exts))
            except Exception:
                continue
        return self._apply_filters(files)

    def _apply_filters(self, files):
        """应用大小/日期筛选器(对全部操作生效)。"""
        from datetime import datetime, timedelta
        min_kb = 0
        d_from = d_to = None
        if hasattr(self, 'filter_min_kb'):
            min_kb = self.filter_min_kb.value()
            s1 = self.filter_date_from.text().strip()
            s2 = self.filter_date_to.text().strip()
            try:
                if s1:
                    d_from = datetime.strptime(s1, '%Y-%m-%d').timestamp()
                if s2:
                    d_to = (datetime.strptime(s2, '%Y-%m-%d')
                            + timedelta(days=1)).timestamp()
            except ValueError:
                pass
        out = []
        seen = set()
        selection, excluded = getattr(self, '_excluded_selection', ('', set()))
        if selection != self.file_edit.text():excluded = set()
        for f in files:
            # 同一文件可能因"选了父文件夹又选了它本身"等路径重复出现。
            # 不去重会让重命名多占序号, 移动/删除第二次必然失败。
            key = os.path.normcase(os.path.abspath(f))
            if key in seen or key in excluded:
                continue
            seen.add(key)
            try:
                st = os.stat(f)
                if min_kb and st.st_size < min_kb * 1024:
                    continue
                if d_from and st.st_mtime < d_from:
                    continue
                if d_to and st.st_mtime > d_to:
                    continue
            except OSError:
                continue
            out.append(f)
        return out

    # ----------------------------------------------------------- 参数区
    def _field(self, form, attr, label, widget):
        """注册字段控件: 页面在 build_options 中调用。"""
        widget.setObjectName(self.kind + '_' + attr)
        widget.setAccessibleName(label or attr)
        self._widgets[attr] = widget
        form.addRow(label, widget)
        return widget

    def _op_widget(self):
        """操作/模式下拉(兼容 op 与 mode 两种命名)。"""
        w = self._widgets.get('op')
        if w is None:
            w = self._widgets.get('mode')
        return w

    def _op_key(self):
        if 'op' in self._widgets:
            return 'op'
        if 'mode' in self._widgets:
            return 'mode'
        return None

    def _set_row_visible(self, attr, vis):
        w = self._widgets.get(attr)
        if w is None:
            return
        label = self.form.labelForField(w)
        w.setVisible(vis)
        if label is not None:
            label.setVisible(vis)

    def _setup_op_fields(self):
        combo = self._op_widget()
        if combo is not None and isinstance(combo, QComboBox):
            combo.currentIndexChanged.connect(lambda _=None: self.apply_op_fields())
        self._make_op_hint()
        self.apply_op_fields()
        self._wire_run_hint()

    def _make_op_hint(self):
        """v3: 操作白话说明 — 在操作下拉下方显示一句人话解释。"""
        key = self._op_key()
        if key is None:
            return
        w = self._widgets[key]
        try:
            pos = self.form.getWidgetPosition(w)[0]
        except Exception:
            return
        if pos < 0:
            return
        self._op_hint = QLabel('')
        self._op_hint.setWordWrap(True)
        self._op_hint.setStyleSheet(
            f'color:{self.main.theme.MUTED};font-size:12px;')
        self._op_hint.setVisible(False)
        self.form.insertRow(pos + 1, '', self._op_hint)

    def apply_op_fields(self):
        combo = self._op_widget()
        sel = combo.currentText() if combo is not None else ''
        rows = self.OP_FIELDS.get(sel, [])
        has_params = bool(rows) or bool(self.ALWAYS_VISIBLE)
        # op_key 同时认 'op' 与 'mode'(如重命名页用 mode 命名),
        # 显隐循环必须豁免操作下拉本身, 否则它会被自己的规则藏掉
        op_key = self._op_key()
        for attr in self._widgets:
            if attr == op_key:
                continue
            vis = attr in rows or attr in self.ALWAYS_VISIBLE or attr == "out_dir"
            self._set_row_visible(attr, vis)
        if self._noparam_hint is not None:
            self._noparam_hint.setVisible(not has_params)
        # 操作白话说明
        hint_w = getattr(self, '_op_hint', None)
        if hint_w is not None:
            hints = getattr(self, 'OP_HINTS', {})
            text = hints.get(sel, '')
            hint_w.setText('提示：' + text if text else '')
            hint_w.setVisible(bool(text))
        if hasattr(self, 'run_hint'):
            self._update_run_hint()

    def _wire_run_hint(self):
        """v3: 执行说明条的联动信号。"""
        self.file_edit.textChanged.connect(lambda _=None: self._update_run_hint())
        self.recursive_cb.toggled.connect(lambda _=None: self._update_run_hint())
        for attr in ('out_dir',):
            w = self._widgets.get(attr)
            if isinstance(w, QLineEdit):
                w.textChanged.connect(lambda _=None: self._update_run_hint())

    def _update_run_hint(self):
        """v3: 用大白话预告执行行为(就地修改 or 输出到目录 + 必填检查)。"""
        hint_w = getattr(self, 'run_hint', None)
        if hint_w is None:
            return
        try:
            n = len(self.current_files())
        except Exception:
            n = 0
        th = self.main.theme
        parts = []
        combo = self._op_widget()
        operation = combo.currentText() if combo is not None else ''
        if n == 0:
            parts.append('① 先在上方选择要处理的文件或文件夹（支持直接拖入窗口）')
            color = th.MUTED
        elif self.kind == 'rename' or (self.kind == 'files' and any(
                word in operation for word in ('删除', '清理', '分类', '修改时间'))):
            if '删除' in operation or '清理空' in operation or '自动清理重复' in operation:
                parts.append('所选范围内的目标将移入回收备份区，可在「回收站」恢复。输出目录不会阻止此操作。')
            elif '修改时间' in operation:
                parts.append('将修改原文件的时间属性，完成后可撤销本次操作。')
            else:
                parts.append('将调整原文件或文件夹的名称、位置，完成后可撤销本次操作。')
            color = th.WARNC
        elif getattr(self, 'MODIFIES_SOURCE', True) is False:
            # 转换/打印/只读查询类操作从不改动原文件, 说成"直接修改原文件"
            # 会让用户以为有风险而不敢点, 或反过来以为有备份而放心乱点。
            parts.append(f'✔ 将处理 {n} 个文件，原文件保持不动')
            color = th.OKC
        else:
            out_dir = ''
            w = self._widgets.get('out_dir')
            if isinstance(w, QLineEdit):
                out_dir = w.text().strip()
            if out_dir:
                parts.append(f'✔ 将处理 {n} 个文件，结果输出到：{out_dir}（原文件保持不动）')
                color = th.OKC
            else:
                parts.append(f'⚠ 将直接修改 {n} 个原文件 — 原文件会自动备份为 '
                             '到“备份”文件夹，可在保留备份时恢复。若想保留原文件，请先在下方'
                             '「输出目录」选择新位置')
                color = th.WARNC
        missing = []
        try:
            missing = self.validate_required()
        except Exception:
            pass
        if missing:
            parts.append('还有必填参数未填：' + '、'.join(missing))
        hint_w.setText('\n'.join(parts))
        hint_w.setStyleSheet(
            f'color:{color};font-size:12px;padding:6px 10px;'
            f'background:{th.PAPER2};border:1px solid {th.LINE};'
            'border-radius:8px;')
        hint_w.setVisible(True)

    def collect_options(self) -> dict:
        return {}

    def processor(self):
        return None

    def validate_required(self):
        """返回当前操作语境下缺失的必填参数显示名列表。"""
        combo = self._op_widget()
        op_name = combo.currentText() if combo is not None else ''
        fields = self.REQUIRED_FIELDS.get(op_name, ())
        missing = []
        for attr, disp in fields:
            w = self._widgets.get(attr)
            v = ''
            if isinstance(w, QLineEdit):
                v = w.text().strip()
            elif isinstance(w, QPlainTextEdit):
                v = w.toPlainText().strip()
            elif isinstance(w, QComboBox):
                # 打印机这类必填项是下拉框。早期只认输入框, 导致"打印机"
                # 永远被判为空, 每次执行都弹"必填参数为空"把任务拦掉。
                v = w.currentText().strip()
            elif isinstance(w, QSpinBox):
                v = str(w.value())
            if not v:
                missing.append(disp)
        return missing

    # ------------------------------------------------------- 状态记忆
    def capture_state(self, include_files: bool = True) -> dict:
        """保存页面参数。

        include_files: 是否把当前文件路径一并存入。方案(preset)场景必须
        为 False —— 用户保存方案是为了复用**参数**, 之后选了新文件再载入
        方案时, 若连路径一起恢复, 会把他刚选的文件悄悄换成旧的, 接着处理
        错文件。流水线步骤相反, 每步各有自己的输入路径, 必须为 True。
        """
        state = {}
        for attr, w in self._widgets.items():
            try:
                if isinstance(w, QLineEdit):
                    state[attr] = w.text()
                elif isinstance(w, QPlainTextEdit):
                    state[attr] = w.toPlainText()
                elif isinstance(w, QComboBox):
                    state[attr] = w.currentText()
                elif isinstance(w, QCheckBox):
                    state[attr] = w.isChecked()
                elif isinstance(w, (QSpinBox, QDoubleSpinBox)):
                    state[attr] = w.value()
            except Exception:
                continue
        if include_files:
            fe = getattr(self, 'file_edit', None)
            if fe is not None:
                state['file_edit'] = fe.text()
        return state

    def restore_state(self, state: dict):
        for attr, w in self._widgets.items():
            if attr not in state:
                continue
            v = state[attr]
            try:
                if isinstance(w, QLineEdit) and isinstance(v, str):
                    w.setText(v)
                elif isinstance(w, QPlainTextEdit) and isinstance(v, str):
                    w.setPlainText(v)
                elif isinstance(w, QComboBox) and isinstance(v, str):
                    pos = w.findText(v)
                    if pos >= 0:
                        w.setCurrentIndex(pos)
                elif isinstance(w, QCheckBox):
                    w.setChecked(bool(v))
                elif isinstance(w, QSpinBox):
                    w.setValue(int(v))
                elif isinstance(w, QDoubleSpinBox):
                    w.setValue(float(v))
            except Exception:
                continue
        # 文件路径: 空值表示"保持现状"(自动记忆场景), 非空才覆盖
        fe = getattr(self, 'file_edit', None)
        v = state.get('file_edit')
        if fe is not None and isinstance(v, str) and v:
            fe.setText(v)
        # 操作/模式下拉兜底: state['op'] 可能是机内名(如 'del_blank'),
        # 下拉框是显示名 — findText 失败时用 op_name 反查(v3 修复)
        op_w = self._op_widget()
        if op_w is not None:
            vo = state.get('op')
            if isinstance(vo, str) and vo and op_w.findText(vo) < 0:
                disp = state.get('op_name')
                if isinstance(disp, str) and disp:
                    pos = op_w.findText(disp)
                    if pos >= 0:
                        op_w.setCurrentIndex(pos)
        self.apply_op_fields()

    def _state_attrs(self):
        return set(self._widgets.keys()) | {'file_edit'}

    def _autosave_params(self):
        try:
            combo = self._op_widget()
            op = combo.currentText() if combo is not None else ''
            state = self.capture_state()
            state['file_edit'] = ''
            ParamStore().save_auto(self.kind, op, state)
        except Exception:
            pass

    def _restore_autosave(self):
        try:
            combo = self._op_widget()
            op = combo.currentText() if combo is not None else ''
            state = ParamStore().load_auto(self.kind, op)
            if state:
                # file_edit 不走 capture/restore(它是文件路径, 不该被参数方案带走)
                f = state.pop('file_edit', '')
                self.restore_state(state)
                if f:
                    self.file_edit.setText(f)
        except Exception:
            pass
        # 页面级缓存(首次打开恢复一次)
        if not getattr(self, '_restored_once', False):
            self._restored_once = True
            self._restore_page_extras()

    def _restore_page_extras(self):
        """子类可覆盖: 恢复页面级缓存(如文件路径之外的输入)。"""
        pass

    # -------------------------------------------------------- 方案菜单
    def _reset_params(self):
        for attr, w in self._widgets.items():
            try:
                if isinstance(w, QLineEdit):
                    w.setText('')
                elif isinstance(w, QPlainTextEdit):
                    w.setPlainText('')
                elif isinstance(w, QCheckBox):
                    w.setChecked(False)
                elif isinstance(w, QComboBox):
                    w.setCurrentIndex(0)
                elif isinstance(w, QSpinBox):
                    w.setValue(0)
                elif isinstance(w, QDoubleSpinBox):
                    w.setValue(0.0)
            except Exception:
                continue
        self.apply_op_fields()
        self.main.show_status('参数已重置')

    def _save_preset_dialog(self):
        combo = self._op_widget()
        default = combo.currentText() if combo is not None else '方案'
        name, ok = QInputDialog.getText(self, '保存参数方案', '方案名称:',
                                        text=default)
        if ok and name.strip():
            op = combo.currentText() if combo is not None else ''
            ParamStore().save_preset(name.strip(), self.title, op,
                                     self.capture_state(include_files=False))
            self.main.show_status('参数方案已保存: ' + name.strip())

    def _load_preset_dialog(self):
        presets = {k: v for k, v in ParamStore().load_presets().items()
                   if v.get('page') == self.title}
        if not presets:
            QMessageBox.information(self, self.main.APP_NAME,
                                    '本页面还没有已保存的方案')
            return
        names = sorted(presets)
        name, ok = QInputDialog.getItem(self, '载入参数方案', '选择方案:',
                                        names, 0, False)
        if ok and name:
            self.restore_state(presets[name].get('state', {}))
            self.main.show_status('已载入方案: ' + name)

    def _manage_presets(self):
        presets = {k: v for k, v in ParamStore().load_presets().items()
                   if v.get('page') == self.title}
        dlg = QDialog(self)
        dlg.setWindowTitle('管理参数方案')
        dlg.resize(520, 420)
        v = QVBoxLayout(dlg)
        t = QTableWidget(len(presets), 4)
        t.setHorizontalHeaderLabels(('方案名', '页面', '操作', '保存时间'))
        t.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        t.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        t.setEditTriggers(QTableWidget.NoEditTriggers)
        for r, (name, p) in enumerate(sorted(presets.items())):
            t.setItem(r, 0, QTableWidgetItem(name))
            t.setItem(r, 1, QTableWidgetItem(p.get('page', '')))
            t.setItem(r, 2, QTableWidgetItem(p.get('op') or '—'))
            t.setItem(r, 3, QTableWidgetItem(p.get('time', '')))
        v.addWidget(t)

        def _del():
            r = t.currentRow()
            if r >= 0:
                name = t.item(r, 0).text()
                ParamStore().delete_preset(name)
                t.removeRow(r)

        h = QHBoxLayout()
        delb = QPushButton('删除选中方案')
        delb.clicked.connect(_del)
        h.addWidget(delb)
        h.addStretch(1)
        closeb = QPushButton('关闭')
        closeb.clicked.connect(dlg.accept)
        h.addWidget(closeb)
        v.addLayout(h)
        dlg.exec()

    def _export_presets(self):
        out, _ = QFileDialog.getSaveFileName(self, '导出方案',
                                             'qingyan_presets.json',
                                             'JSON (*.json)')
        if out:
            n = ParamStore().export_presets(out)
            QMessageBox.information(self, self.main.APP_NAME,
                                    f'已导出 {n} 个方案:\n{out}')

    def _import_presets(self):
        f, _ = QFileDialog.getOpenFileName(self, '导入方案', '',
                                           'JSON (*.json)')
        if f:
            n = ParamStore().import_presets(f)
            QMessageBox.information(self, self.main.APP_NAME, f'已导入 {n} 个方案')

    # -------------------------------------------------------------- 预览
    def build_preview(self, files, options):
        return None

    @staticmethod
    def _split_preview(data):
        """归一化 build_preview 的返回值 → (header, rows)。

        子类有两种写法: (header, rows) 二元组, 或直接返回 rows(首行当表头)。
        早期 do_preview 固定按二元组解包, 后一种页面选 2 个以上文件时
        会抛 "too many values to unpack" 把预览整个打崩。
        """
        if (isinstance(data, tuple) and len(data) == 2
                and isinstance(data[1], (list, tuple))
                and (not data[1] or isinstance(data[1][0], (list, tuple)))):
            return list(data[0]), list(data[1])
        rows = list(data)
        if not rows:
            return [], []
        header = list(rows[0])
        return header, rows[1:]

    def do_preview(self):
        files = self.current_files()
        if not files:
            QMessageBox.information(self, self.main.APP_NAME, '请先选择文件')
            return
        try:
            options = self.collect_options()
            data = self.build_preview(files, options)
        except Exception as error:
            QMessageBox.warning(self, self.main.APP_NAME,
                                '无法预览，请检查文件是否完整、参数是否正确。\n' + str(error))
            return
        plan_only = not data
        if plan_only:
            combo = self._op_widget()
            operation = combo.currentText() if combo is not None else self.title
            destination = options.get('out_dir') or options.get('pdf_out_dir') or '按当前操作保存到原位置或旁边'
            if self.kind == 'rename':destination = '在原位置调整名称'
            if self.kind == 'files' and ('删除' in operation or '清理' in operation):destination = '按操作移入回收备份区或调整名称'
            data = (('输入项目', '将执行', '输出设置'), [(path, operation, destination) for path in files])
        header, rows = self._split_preview(data)
        if not rows:
            QMessageBox.information(self, self.main.APP_NAME, '该功能暂无预览')
            return
        dlg = QDialog(self)
        dlg.setWindowTitle('预览(不修改文件)')
        dlg.resize(780, 500)
        v = QVBoxLayout(dlg)
        if plan_only:
            note = QLabel('请核对文件与处理方式。这里展示处理计划，文档的排版效果请在处理后打开输出文件查看。')
            note.setWordWrap(True)
            v.addWidget(note)
        if len(rows) > 500:
            note = QLabel(f'共 {len(rows)} 项，预览显示前 500 项；执行时会处理全部所选项目。')
            note.setWordWrap(True)
            v.addWidget(note)
            rows = rows[:500]
        t = QTableWidget(len(rows), len(header))
        t.setHorizontalHeaderLabels(header)
        t.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        t.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        if t.columnCount() > 2:
            t.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        t.setEditTriggers(QTableWidget.NoEditTriggers)
        for i, row in enumerate(rows):
            for j, cell in enumerate(row):
                t.setItem(i, j, QTableWidgetItem(str(cell)))
        v.addWidget(t)
        b = QPushButton('关闭')
        b.clicked.connect(dlg.accept)
        v.addWidget(b)
        dlg.exec()

    # -------------------------------------------------------------- 执行
    def do_run(self):
        files = self.current_files()
        if not files:
            QMessageBox.information(self, self.main.APP_NAME, '请先选择文件')
            return
        # v3: 删除/清理类操作先确认, 防止误操作
        combo = self._op_widget()
        op_name = combo.currentText() if combo is not None else ''
        confirm_tpl = getattr(self, 'CONFIRM_OPS', {}).get(op_name)
        if confirm_tpl:
            ret = QMessageBox.question(
                self, self.main.APP_NAME, confirm_tpl.format(n=len(files)))
            if ret != QMessageBox.Yes:
                return
        proc = self.processor()
        if proc is None:
            QMessageBox.information(self, self.main.APP_NAME,
                                    '该功能需要专用流程, 请查看手册')
            return
        options = self.collect_options()
        missing = self.validate_required()
        if missing:
            QMessageBox.warning(self, self.main.APP_NAME,
                                '请填写以下参数后再执行：\n' + '、'.join(missing))
            return
        self._autosave_params()
        spec = TaskSpec(kind=self.kind, title=self.title, files=files,
                        options=options)
        self._start_worker(spec, proc)

    def _guarded(self, proc):
        """v3: 扩展名守卫 — 格式不符的文件跳过并给出说明, 不报错。"""
        exts = getattr(self, 'EXPECTED_EXTS', ())
        if not exts:
            return proc

        def wrapper(path, options):
            ext = os.path.splitext(str(path))[1].lower()
            if ext and ext not in exts:
                return FileResult(
                    path, 'skipped',
                    f'格式不符 {ext} (本操作支持: {" ".join(exts)})')
            return proc(path, options)

        return wrapper

    def _start_worker(self, spec: TaskSpec, proc):
        combo = self._op_widget()
        if combo is not None:
            spec.options.setdefault('op_name', combo.currentText())
        elif len(self.OPS) == 1:
            spec.options.setdefault('op_name', self.OPS[0])
        if getattr(self.main, '_task_running', False) or any(getattr(p, 'worker', None) is not None and p.worker.isRunning()
               for p in self.main.pages):
            QMessageBox.information(self, self.main.APP_NAME, '请等待当前任务完成，或先取消当前任务。')
            return
        self.result_banner.hide()
        self.result_table.setRowCount(0)
        self._summary_counts = {'ok': 0, 'fail': 0, 'skipped': 0}
        self._update_summary()
        self._update_result_empty()
        self.worker = BatchWorker(spec, self._guarded(proc),
                                  self.main.RECORDS_DIR,
                                  self.main.UNDO_MANIFEST)
        self.worker.progress.connect(self.main.on_progress)
        self.worker.file_done.connect(self.on_file_done)
        self.worker.finished_all.connect(self.on_finished)
        self.worker.cancelled.connect(self.on_cancelled)
        self.run_btn.setEnabled(False)
        self.preview_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)
        self.main.set_running(True, self.title)
        self.worker.start()

    def do_cancel(self):
        if self.worker:
            self.worker.cancel()

    def on_file_done(self, r):
        self._update_result_empty()
        th = self.main.theme
        row = self.result_table.rowCount()
        self.result_table.insertRow(row)
        dot = '●' if r.status == 'ok' else ('●' if r.status == 'fail' else '○')
        label = {'ok': '成功', 'fail': '失败', 'skipped': '跳过'}.get(
            r.status, r.status)
        s_item = QTableWidgetItem(dot + ' ' + label)
        color = QColor(th.ACCENT)
        if r.status == 'fail':
            color = QColor(th.FAILC)
        elif r.status == 'skipped':
            # 跳过与成功用同一强调色时, 批量扫结果很容易漏看
            color = QColor(th.MUTED)
        s_item.setForeground(color)
        detail = r.detail or ''
        if r.output:
            detail = (detail + '  → ' + r.output).strip(' →')
        self.result_table.setItem(row, 0, QTableWidgetItem(os.path.basename(r.path)))
        self.result_table.setItem(row, 1, s_item)
        self.result_table.setItem(row, 2, QTableWidgetItem(detail))
        self.result_table.item(row, 0).setToolTip(r.path)
        self.result_table.item(row, 0).setData(Qt.UserRole, r.output)
        if r.duration_sec:
            self.result_table.item(row, 1).setText(
                s_item.text() + f'  {r.duration_sec:.1f}s')
        self._summary_counts[r.status if r.status in self._summary_counts
                             else 'skipped'] += 1
        self._update_summary()
        self._filter_fail_only(self.fail_only_cb.isChecked())

    def _update_summary(self):
        c = self._summary_counts
        total = c['ok'] + c['fail'] + c['skipped']
        if not total:
            self.summary_label.setText('')
            return
        th = self.main.theme
        self.summary_label.setText(
            f'共 {total} · <span style="color:{th.OKC}">成功 {c["ok"]}</span>'
            f' · <span style="color:{th.FAILC}">失败 {c["fail"]}</span>'
            f' · 跳过 {c["skipped"]}')

    def on_finished(self, results, task_id):
        self.run_btn.setEnabled(True)
        self.preview_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)
        self.main.set_running(False)
        ok = sum(1 for r in results if r.status == 'ok')
        fail = sum(1 for r in results if r.status == 'fail')
        skip = sum(1 for r in results if r.status == 'skipped')
        self.main.show_status(
            f'任务完成 {task_id} — 成功 {ok} / 失败 {fail} / 跳过 {skip}')
        self.main.records_dirty = True
        if self.open_out_cb.isChecked():
            dirs = []
            for r in results:
                if r.status == 'ok' and r.output:
                    d = os.path.dirname(r.output) if os.path.isabs(r.output) \
                        else os.path.dirname(os.path.abspath(r.path))
                    if os.path.isdir(d) and d not in dirs:
                        dirs.append(d)
            for d in dirs[:2]:
                try:
                    os.startfile(d)
                except Exception:
                    pass
        # v3: 完成后快捷条
        outs = []
        for r in results:
            if r.status == 'ok' and r.output:
                d = r.output if os.path.isabs(r.output) else \
                    os.path.abspath(os.path.join(os.path.dirname(r.path),
                                                 r.output))
                d2 = os.path.dirname(d)
                if os.path.isdir(d2) and d2 not in outs:
                    outs.append(d2)
        self._banner_outputs = outs[:3]
        self._banner_undo_actions = list(self.worker.undo_actions) \
            if self.worker is not None else []
        # 全部失败时不能还打勾说"任务完成", 否则用户扫一眼就走, 以为成了
        if ok == 0 and fail > 0:
            self.banner_label.setText(
                f'✘ 全部失败 — {fail} 个文件都没处理成功'
                '（看下方结果表的原因列）')
        else:
            self.banner_label.setText(
                f'✔ 任务完成 — 成功 {ok} 个'
                + (f'，失败 {fail} 个（可勾选「只看失败」定位）' if fail else '')
                + (f'，跳过 {skip} 个' if skip else ''))
        self.banner_open_btn.setVisible(bool(self._banner_outputs))
        self.banner_undo_btn.setVisible(bool(self._banner_undo_actions))
        self.result_banner.show()
        self._update_result_empty()

    def on_cancelled(self):
        self.on_finished(self.worker.results, '')
        self.banner_label.setText('任务已取消 · 已完成的结果和撤销记录已保留')
        self.main.show_status('任务已取消')

    # ------------------------------------------------- 完成后快捷条动作
    def _banner_open(self):
        for d in self._banner_outputs[:1]:
            try:
                os.startfile(d)
            except Exception:
                self.main.show_status('打开失败: ' + d)

    def _banner_undo(self):
        """撤销本次任务: move 类换回位置, restore 类用备份覆盖回原文件。"""
        from engine.undo import undo_actions
        from engine.store import _atomic_write
        import json
        if getattr(self.main, '_task_running', False):
            self.main.show_status('请等待当前任务完成后再撤销')
            return
        completed, remaining, errors = undo_actions(self._banner_undo_actions)
        try:
            with open(self.main.UNDO_MANIFEST, encoding='utf-8') as fh:
                manifest = json.load(fh)
            actions = list(manifest.get('actions', []))
            for action in completed:
                if action in actions:
                    actions.remove(action)
            _atomic_write(self.main.UNDO_MANIFEST, {**manifest, 'actions': actions})
        except (OSError, ValueError) as exc:
            errors.append('撤销记录保存失败：' + str(exc))
        self._banner_undo_actions = remaining
        self.banner_undo_btn.setVisible(bool(remaining))
        msg = f'已撤销 {len(completed)} 项'
        if errors:
            msg += f'；剩余 {len(remaining)} 项已保留，可重试。' + errors[0]
        self.banner_label.setText(msg)
        self.main.show_status(msg)

    # ---------------------------------------------------------- 右键菜单
    def _result_menu(self, pos):
        # 用右键点中的行, 而不是"上次左键选中的行": 先选第 1 行再右键第 5 行
        # 时, 打开文件/打开位置会作用到第 1 行, 用户会打开错误的文件。
        item = self.result_table.itemAt(pos)
        row = item.row() if item is not None else self.result_table.currentRow()
        if row < 0:
            return
        detail = self.result_table.item(row, 2).text() if \
            self.result_table.item(row, 2) else ''
        out_path = self.result_table.item(row, 0).data(Qt.UserRole) or ''
        name_item = self.result_table.item(row, 0)
        src_name = name_item.toolTip() if name_item else ''
        menu = QMenu(self)
        added = False
        if out_path and os.path.isabs(out_path) and os.path.isfile(out_path):
            a1 = menu.addAction('打开文件')
            a1.triggered.connect(lambda: self.main.open_path(out_path))
            added = True
        d = os.path.dirname(out_path) if out_path else ''
        if d and os.path.isdir(d):
            a2 = menu.addAction('打开所在位置')
            a2.triggered.connect(lambda: self.main.open_path(d))
            added = True
        if src_name and os.path.exists(src_name):
            a3 = menu.addAction('打开源文件位置')
            a3.triggered.connect(
                lambda: self.main.open_path(os.path.dirname(src_name)))
            added = True
        if added:
            menu.addSeparator()
        a4 = menu.addAction('复制说明内容')
        a4.triggered.connect(lambda: QGuiApplication.clipboard().setText(detail))
        menu.exec(self.result_table.viewport().mapToGlobal(pos))

    def _open_path(self, path):
        try:
            os.startfile(path)
        except Exception:
            QMessageBox.warning(self, self.main.APP_NAME, '打开失败: ' + path)

    def _filter_fail_only(self, only):
        for r in range(self.result_table.rowCount()):
            item = self.result_table.item(r, 1)
            status = item.text() if item else ''
            self.result_table.setRowHidden(r, only and ('失败' not in status))

    # --------------------------------------------------------- 空态与事件
    def _update_result_empty(self):
        vp = self.result_table.viewport()
        if vp.height() < 120:
            self.result_empty.setText('处理结果将显示在这里')
        else:
            self.result_empty.setText('还没有执行记录\n\n选好文件与参数后点「开始处理」，\n每个文件的处理结果会显示在这里')
        self.result_empty.setGeometry(0, 0, vp.width(), vp.height())
        self.result_empty.setVisible(self.result_table.rowCount() == 0)
        self.result_empty.raise_()

    def eventFilter(self, obj, event):
        if obj is self.result_table.viewport():
            self._update_result_empty()
        return super().eventFilter(obj, event)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        splitter = getattr(self, '_splitter', None)
        if splitter is not None:
            orientation = Qt.Vertical if self.width() < 880 else Qt.Horizontal
            if splitter.orientation() != orientation:
                splitter.setOrientation(orientation)
                splitter.setSizes([600, 350])
        self._update_result_empty()

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        paths = []
        seen = set()
        for url in event.mimeData().urls():
            p = url.toLocalFile()
            if not os.path.exists(p):
                continue
            # 同时拖入父文件夹和它里面的文件时会出现重复/包含关系, 同一
            # 文件被处理两次(重命名多占号、移动/删除第二次必然失败)
            key = os.path.normcase(os.path.abspath(p))
            if key in seen:
                continue
            seen.add(key)
            paths.append(p)
        if paths:
            if len(paths) == 1 and os.path.isdir(paths[0]):
                self.file_edit.setText(paths[0])
            else:
                # 已有内容时追加, 避免拖入一次就把之前选的清空
                old = [x for x in self.file_edit.text().split('||') if x.strip()]
                merged = list(old)
                for p in paths:
                    if os.path.normcase(os.path.abspath(p)) not in {
                            os.path.normcase(os.path.abspath(x)) for x in old}:
                        merged.append(p)
                self.file_edit.setText('||'.join(merged))
            self.main.show_status(f'已从拖放导入 {len(paths)} 项')
