# -*- coding: utf-8 -*-
"""青雁批量工具箱 v3 — 主窗口与程序入口。

架构: 首页(功能网格导航) + 左侧功能树 + 中央堆叠页 + 底部任务条。
v3 新增: 亮/暗主题 + 强调色 + 字号缩放(全局设置可切换)、ETA/速度显示、
任务取消按钮、Ctrl+Enter 执行。
"""
from __future__ import annotations

import os
import sys
import traceback

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QIcon, QKeySequence, QShortcut
from PySide6.QtWidgets import (QApplication, QFrame, QHBoxLayout, QLabel,
                               QListWidget, QListWidgetItem, QMainWindow,
                               QMessageBox, QProgressBar, QPushButton,
                               QStackedWidget, QStyleFactory, QVBoxLayout,
                               QWidget)

from app import icons
from app.theme import Theme
from engine.store import ParamStore, data_dir

APP_NAME = '青雁批量工具箱'
APP_VERSION = '3.4.1'


def _icon_path():
    """图标资源路径: 打包后随 _internal 分发, 源码运行用 icon/ 目录。"""
    base = getattr(sys, '_MEIPASS', None)
    cands = []
    if base:
        cands.append(os.path.join(base, 'icon.ico'))
    cands.append(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                              '..', 'icon.ico'))
    for c in cands:
        if c and os.path.exists(c):
            return os.path.abspath(c)
    return None


class MainWindow(QMainWindow):
    APP_NAME = APP_NAME
    APP_VERSION = APP_VERSION

    def __init__(self):
        super().__init__()
        store = ParamStore()
        self.store = store
        th = store.get_theme()
        self.theme = Theme(th.get('mode', 'light'), th.get('accent', '#635BDF'),
                           th.get('font_scale', 1.0))
        self.records_dirty = False

        self.APP_DATA_DIR = data_dir()
        self.RECORDS_DIR = os.path.join(self.APP_DATA_DIR, 'records')
        self.TRASH_DIR = os.path.join(self.APP_DATA_DIR, 'trash')
        self.UNDO_MANIFEST = os.path.join(self.APP_DATA_DIR, 'last_undo.json')

        self.setWindowTitle(f'{APP_NAME} v{APP_VERSION} — 免费 · 离线 · 本地')
        self.resize(1380, 900)
        self.setMinimumSize(980, 640)
        _ic = _icon_path()
        if _ic:
            self.setWindowIcon(QIcon(_ic))

        self._page_defs = self._default_page_titles()
        self.page_index_by_title = {}
        self.feature_page_map = {}
        self.pages = []

        central = QWidget()
        central.setObjectName("workspace")
        self.setCentralWidget(central)
        v = QVBoxLayout(central)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)

        # ------------------------------------------------------------- 头部
        header = QFrame()
        header.setObjectName("appHeader")
        header.setStyleSheet(
            f'background:{self.theme.PAPER};'
            f'border-bottom:1px solid {self.theme.LINE};')
        hv = QHBoxLayout(header)
        hv.setContentsMargins(24, 14, 24, 14)
        self.logo_btn = QPushButton(f'{APP_NAME} v{APP_VERSION}')
        self.logo_btn.setCursor(Qt.PointingHandCursor)
        self.logo_btn.setStyleSheet(
            f'color:{self.theme.INK};font-size:17px;font-weight:600;'
            'letter-spacing:5px;background:transparent;border:none;'
            'text-align:left;padding:2px 8px;border-radius:8px;')
        self.logo_btn.clicked.connect(lambda: self.goto_page(0))
        hv.addWidget(self.logo_btn)
        hv.addStretch(1)
        badges = QLabel('本地处理  /  隐私优先     Ctrl+F 搜索工具' )
        badges.setStyleSheet(
            f'color:{self.theme.MUTED};font-size:12px;letter-spacing:0.5px;')
        backup_button = QPushButton("备份管理")
        backup_button.setObjectName("manageBackups")
        backup_button.clicked.connect(self.manage_backups)
        hv.addWidget(backup_button)
        hv.addWidget(badges)
        v.addWidget(header)

        # ------------------------------------------------------------- 主体
        body = QHBoxLayout()
        body.setSpacing(2)
        body.setContentsMargins(12, 10, 0, 4)
        self.nav = QListWidget()
        self.nav.setFixedWidth(176)
        self.nav.setStyleSheet(
            'QListWidget{border:none;border-right:1px solid '
            f'{self.theme.LINE};background:{self.theme.PAPER2};font-size:13px;}}'
            f'QListWidget::item{{padding:12px 18px;margin:1px 8px;'
            f'border-radius:9px;color:{self.theme.INK};}}'
            'QListWidget::item:hover{background:'
            f'{self.theme.c["HOVER"]};}}'
            'QListWidget::item:selected{background:'
            f'{self.theme.c["SEL"]};color:{self.theme.ACCENT};'
            'font-weight:600;}')
        body.addWidget(self.nav)
        self.stack = QStackedWidget()
        body.addWidget(self.stack, 1)
        v.addLayout(body, 1)

        # 构建页面
        from app.pages.unified_page import UnifiedPage
        from app.pages.convert_page import ConvertPage
        from app.pages.feature_pages import (ExcelPage, FilesPage, ImagePage,
                                             PdfPage, PptPage, RenamePage,
                                             TemplatePage, TextPage, WordPage)
        from app.pages.print_page import PrintPage
        from app.pages.special_pages import (AboutPage, HelpPage, HomePage,
                                             MergePage, OfficeTemplatesPage,
                                             PipelinePage, RecordsPage,
                                             SettingsPage, ToolsPage,
                                             TrashPage)
        cls_map = {
            '首页': None,
            'Word 工具': WordPage, 'Excel 工具': ExcelPage, 'PPT 工具': PptPage,
            'PDF 工具': PdfPage, '文本工具': TextPage, '图片工具': ImagePage,
            '重命名工具': RenamePage, '文件整理工具': FilesPage,
            '批量打印': PrintPage, '格式转换': ConvertPage,
            '模板批量生成': TemplatePage, '办公模板库': OfficeTemplatesPage,
            '合并与拆分': MergePage, '辅助工具': ToolsPage,
            '任务台账': RecordsPage, '回收站': TrashPage,
            '任务流水线': PipelinePage, '统一替换': UnifiedPage, '帮助中心': HelpPage,
            '全局设置': SettingsPage, '关于': AboutPage,
        }
        self._page_defs = [(title, cls_map.get(title)) for title, _ in
                           self._default_page_titles()]
        for name, cls in self._page_defs:
            if cls is None:
                continue
            page = cls(self)
            self.pages.append(page)
            self.stack.addWidget(page)
        # 首页插入最前
        home = HomePage(self)
        self.pages.insert(0, home)
        self.stack.insertWidget(0, home)

        for i, (name, _cls) in enumerate(self._page_defs):
            self.page_index_by_title[name] = i
            it = QListWidgetItem(name)
            it.setIcon(icons.get_icon(icons.NAV_ICONS.get(name, 'star'),
                                      20, self.theme.INK))
            self.nav.addItem(it)
        self.nav.currentRowChanged.connect(self._on_nav_changed)
        self.nav.setCurrentRow(0)

        # ------------------------------------------------------------- 底部
        bottom = QFrame()
        bottom.setObjectName("taskFooter")
        bottom.setStyleSheet(
            f'border-top:1px solid {self.theme.LINE};'
            f'background:{self.theme.PAPER2};')
        bv = QHBoxLayout(bottom)
        bv.setContentsMargins(18, 8, 18, 8)
        self.status_label = QLabel('就绪 · 选择工具开始处理')
        self.status_label.setObjectName('globalStatus')
        self.status_label.setStyleSheet(
            f'color:{self.theme.MUTED};font-size:12px;')
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        self.progress.setFixedWidth(300)
        self.progress.setMaximumHeight(14)
        self.progress_label = QLabel('')
        self.progress_label.setStyleSheet(
            f'color:{self.theme.MUTED};font-size:12px;')
        self.cancel_btn = QPushButton('取消任务')
        self.cancel_btn.setVisible(False)
        self.cancel_btn.clicked.connect(self._cancel_current)
        bv.addWidget(self.status_label, 1)
        bv.addWidget(self.progress)
        bv.addWidget(self.progress_label)
        bv.addWidget(self.cancel_btn)
        v.addWidget(bottom)

        self._setup_shortcuts()
        self._apply_theme_hooks()

        # 恢复窗口状态
        try:
            ui = store.load_ui()
            size = ui.get('size')
            if isinstance(size, list) and len(size) == 2:
                self.resize(int(size[0]), int(size[1]))
            pos = ui.get('pos')
            if isinstance(pos, list) and len(pos) == 2:
                self.move(int(pos[0]), int(pos[1]))
            last = ui.get('last_page')
            startup = ui.get('startup_page', '首页')
            target = startup if startup in self.page_index_by_title else last
            if target and target in self.page_index_by_title:
                self.nav.setCurrentRow(self.page_index_by_title[target])
        except Exception:
            pass
        # The homepage provides onboarding without an interrupting first-run dialog.

    def manage_backups(self):
        from app.backup_dialog import BackupDialog
        if getattr(self, "_task_running", False):
            QMessageBox.information(self, "任务进行中", "请等待当前任务完成后管理备份。")
            return
        try:
            BackupDialog(self).exec()
        except Exception as exc:
            QMessageBox.warning(self, "备份管理", str(exc))

    # ------------------------------------------------------------- 页面定义
    @staticmethod
    def _default_page_titles():
        return (('首页', None), ('Word 工具', None), ('Excel 工具', None),
                ('PPT 工具', None), ('PDF 工具', None), ('文本工具', None),
                ('图片工具', None), ('重命名工具', None),
                ('文件整理工具', None), ('批量打印', None), ('格式转换', None),
                ('模板批量生成', None),
                ('办公模板库', None), ('合并与拆分', None),
                ('辅助工具', None), ('任务台账', None), ('回收站', None),
                ('任务流水线', None), ('统一替换', None), ('帮助中心', None),
                ('全局设置', None), ('关于', None))

    # ------------------------------------------------------------- 主题
    def apply_theme(self, mode: str, accent: str, scale: float):
        self.theme = Theme(mode, accent, scale)
        ParamStore().set_theme(mode=mode, accent=accent, font_scale=scale)
        self._apply_theme_hooks()

    def _apply_theme_hooks(self):
        from app.design import apply_design
        apply_design(self)

    # ------------------------------------------------------------- 快捷键
    def _setup_shortcuts(self):
        sc = QShortcut(QKeySequence('Ctrl+F'), self)
        sc.activated.connect(self._focus_search)
        for i in range(10):
            key = str((i + 1) % 10)
            sc = QShortcut(QKeySequence('Ctrl+' + key), self)
            sc.activated.connect(
                lambda _=False, idx=i: self.goto_page(idx))
        sc = QShortcut(QKeySequence('F1'), self)
        sc.activated.connect(
            lambda: self.goto_page(self.page_index_by_title.get('帮助中心', 0)))
        sc = QShortcut(QKeySequence('Ctrl+L'), self)
        sc.activated.connect(
            lambda: self.goto_page(self.page_index_by_title.get('任务台账', 0)))

    def _focus_search(self):
        self.goto_page(0)
        home = self.pages[0]
        if hasattr(home, 'search_edit'):
            home.search_edit.setFocus()
            home.search_edit.selectAll()

    # ------------------------------------------------------------- 导航
    def goto_page(self, idx):
        if 0 <= idx < self.nav.count():
            self.nav.setCurrentRow(idx)
            title = self.nav.item(idx).text()
            if title == '任务台账':
                page = self.pages[idx]
                if hasattr(page, 'refresh'):
                    page.refresh()
            elif title == '回收站':
                page = self.pages[idx]
                if hasattr(page, 'refresh'):
                    page.refresh()

    def goto_feature(self, target_page, op_name):
        idx = self.page_index_by_title.get(target_page)
        if idx is None:
            self.show_status('未找到页面: ' + str(target_page))
            return
        self.nav.setCurrentRow(idx)
        page = self.pages[idx]
        if op_name:
            combo = page._op_widget() if hasattr(page, '_op_widget') else None
            if combo is not None:
                pos = combo.findText(op_name)
                if pos >= 0:
                    combo.setCurrentIndex(pos)
                    self.show_status(
                        f'已定位: {target_page} → 当前操作「{op_name}」, '
                        '其余操作已隐藏(可用上方「操作」下拉切回)')
                    return
        self.show_status('已定位: ' + target_page)

    # ------------------------------------------------------------- 任务条
    def set_running(self, running: bool, title: str = ''):
        self._task_running = running
        self.progress.setVisible(running)
        self.progress.setRange(0, 0)
        self.cancel_btn.setVisible(running)
        self._cancel_target = None
        if running:
            self.status_label.setText(f'正在执行: {title} …')
            self.progress_label.setText('')
        else:
            self.progress_label.setText('')

    def on_progress(self, done, total, name, elapsed):
        if total > 0 and elapsed > 0:
            self.progress.setRange(0, 100)
            self.progress.setValue(round(done * 100 / total))
            eta = elapsed / done * (total - done) if done else 0
            speed = done / elapsed * 60
            self.progress_label.setText(
                f'{done}/{total}  {name[:24]}  '
                f'速度 {speed:.0f} 个/分  预计剩余 {eta:.0f}s')
        else:
            self.progress_label.setText(f'{done}/{total}')

    def _cancel_current(self):
        for p in self.pages:
            if getattr(p, '_pipeline_active', False):
                from app.pipeline import cancel
                cancel(p)
            w = getattr(p, 'worker', None)
            if w is not None and w.isRunning():
                w.cancel()
                self.show_status('正在取消…')

    def show_status(self, text):
        self.status_label.setText(text)

    def open_path(self, path):
        try:
            os.startfile(path)
        except Exception:
            QMessageBox.warning(self, APP_NAME, '打开失败: ' + str(path))

    # ------------------------------------------------------------- 欢迎
    def _maybe_welcome(self):
        if os.environ.get('QT_QPA_PLATFORM') == 'offscreen':
            return  # 自动化测试环境跳过
        try:
            ui = self.store.load_ui()
            if ui.get('welcomed_v3'):
                return
            self.store.save_ui({'welcomed_v3': True})
        except Exception:
            return
        from PySide6.QtWidgets import QDialog, QFrame
        dlg = QDialog(self)
        dlg.setWindowTitle(f'欢迎使用 {APP_NAME} v3')
        dlg.setFixedSize(560, 470)
        v = QVBoxLayout(dlg)
        v.setContentsMargins(32, 28, 32, 24)
        v.setSpacing(14)
        t = QLabel('👋 欢迎使用 青雁批量工具箱')
        t.setStyleSheet(
            f'font-size:21px;font-weight:700;color:{self.theme.ACCENT};')
        v.addWidget(t)
        sub = QLabel('完全免费 · 完全离线 · 数据不出本机')
        sub.setStyleSheet(f'color:{self.theme.MUTED};font-size:12px;')
        v.addWidget(sub)

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet(f'color:{self.theme.LINE};')
        v.addWidget(line)

        steps = (
            ('1', '首页点一个功能（或搜索）',
             '135+ 功能按类型分好组；右键可收藏常用功能'),
            ('2', '把文件拖进窗口，或点「选文件/选文件夹」',
             '可展开文件清单核对，支持扩展名/大小/日期过滤'),
            ('3', '点「执行任务」看结果',
             '参数只显示需要的部分，每个操作都有白话说明；'
             '改错了随时撤销，删除进回收站可恢复'),
        )
        for num, title, desc in steps:
            row = QHBoxLayout()
            badge = QLabel(num)
            # 居中要用 Qt 的 alignment 属性, 不能写进 QSS:
            # 'qproperty-alignment' 不是合法样式项(只会刷 "Unknown property" 警告)
            badge.setAlignment(Qt.AlignCenter)
            badge.setStyleSheet(
                f'background:{self.theme.ACCENT};color:white;border-radius:14px;'
                'min-width:28px;max-width:28px;min-height:28px;max-height:28px;'
                'font-weight:700;font-size:15px;')
            row.addWidget(badge)
            txt = QLabel(f'<b>{title}</b><br>'
                         f'<span style="color:{self.theme.MUTED};">{desc}</span>')
            txt.setWordWrap(True)
            row.addWidget(txt, 1)
            v.addLayout(row)

        tip = QLabel('💡 提示：所有处理都在本机完成，不联网、不上传任何文件。')
        tip.setStyleSheet(
            f'color:{self.theme.MUTED};font-size:12px;padding:8px;'
            f'background:{self.theme.PAPER2};border-radius:8px;')
        v.addWidget(tip)
        v.addStretch(1)

        brow = QHBoxLayout()
        brow.addStretch(1)
        helpb = QPushButton('打开帮助中心')
        helpb.clicked.connect(
            lambda: (dlg.accept(),
                     self.goto_page(self.page_index_by_title.get('帮助中心', 0))))
        brow.addWidget(helpb)
        startb = QPushButton('开始使用 →')
        startb.setProperty('accent', True)
        startb.setStyleSheet(
            f'background:{self.theme.ACCENT};color:white;padding:8px 26px;'
            'font-weight:600;border:none;border-radius:9px;')
        startb.clicked.connect(dlg.accept)
        brow.addWidget(startb)
        v.addLayout(brow)
        dlg.exec()

    def _on_nav_changed(self, idx: int):
        """切换左侧导航: 切页 + 按需刷新台账。

        原先只连 setCurrentIndex, 导致执行完任务直接点「任务台账」看到的
        还是启动时的旧数据(records_dirty 只被赋值、从没人读)。
        """
        self.stack.setCurrentIndex(idx)
        if idx < 0 or idx >= len(self.pages):
            return
        page = self.pages[idx]
        if getattr(page, 'title', '') == '任务台账' and self.records_dirty:
            try:
                page.refresh()
                self.records_dirty = False
            except Exception:
                pass
        elif hasattr(page, 'refresh_recent') and self.records_dirty:
            # 首页「最近使用」也得跟着新台账刷新: 早期它只在 __init__ 里
            # 读一次, 启动后执行的任务永远不会出现在首页。
            try:
                page.refresh_recent()
            except Exception:
                pass

    def closeEvent(self, event):
        # 先收尾正在跑的任务: 直接退出会让 QThread 在运行中被销毁
        # (Qt 报 "Destroyed while thread is still running" 并可能 abort),
        # 同时目标文件可能停在写了一半的状态。
        running = [p for p in self.pages
                   if getattr(p, '_pipeline_active', False) or
                   (getattr(p, 'worker', None) is not None and p.worker.isRunning())]
        if running:
            from PySide6.QtWidgets import QMessageBox
            r = QMessageBox.question(
                self, self.APP_NAME,
                f'还有 {len(running)} 个任务正在进行。\n'
                '现在退出会中断处理, 可能留下写了一半的文件。\n\n'
                '要先取消任务再退出吗?',
                QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel,
                QMessageBox.No)
            if r != QMessageBox.Yes:
                event.ignore()
                return
            if r == QMessageBox.Yes:
                for p in running:
                    try:
                        if getattr(p, '_pipeline_active', False):
                            from app.pipeline import cancel
                            cancel(p)
                        p.worker.cancel()
                    except Exception:
                        pass
                for p in running:
                    try:
                        if not p.worker.wait(300):
                            self.show_status("正在安全取消任务，处理结束后可关闭窗口。")
                            event.ignore()
                            return
                    except Exception:
                        pass
        from app.backup_dialog import allow_exit
        if not allow_exit(self):
            event.ignore()
            return
        try:
            ui = {'size': [self.width(), self.height()],
                  'pos': [self.x(), self.y()],
                  'last_page':
                      self.nav.item(self.nav.currentRow()).text()
                      if self.nav.currentRow() >= 0 else '首页'}
            # 记住每个功能页左右分栏的拖动宽度(v3)
            splitters = {}
            for p in self.pages:
                sp = getattr(p, '_splitter', None)
                if sp is not None:
                    try:
                        splitters[p.kind] = list(sp.sizes())
                    except Exception:
                        pass
            if splitters:
                ui['splitters'] = splitters
            self.store.save_ui(ui)
        except Exception:
            pass
        super().closeEvent(event)


# ---------------------------------------------------------------- 单实例
def _acquire_single_instance():
    """单实例锁: 防止重复启动叠出一堆窗口。

    没有锁时, 用户双击两次(或从快捷方式+安装包各启动一次)会出现多个
    满屏窗口, 表现为"任务栏一排条目 / 桌面被大片窗口盖住"。
    返回锁对象(需在 app 退出时释放); 已运行时弹出提示并返回 None。
    """
    from PySide6.QtCore import QLockFile, QCoreApplication
    lock_path = os.path.join(data_dir(), 'app.lock')
    lock = QLockFile(lock_path)
    # 不设 setStaleLockTime(0): 那样 Qt 会认为锁永不过期, 上次异常退出留下的
    # 残锁会导致之后再也启动不了。默认 30s 陈旧判定能自动回收死进程留下的锁,
    # 同时活着的实例仍会被正确拦住。
    if not lock.tryLock(100):
        # 已有实例在跑: 直接退出即可。
        # 注意不要用 QMessageBox 弹模态框 —— 此时还没有主窗口, 模态框会
        # 在事件循环里阻塞, 用户看不到就以为程序卡死; 无人值守场景
        # (双击多次) 也会把启动流程彻底卡住。改为写日志 + 立即退出。
        try:
            from datetime import datetime
            log_dir = os.path.join(data_dir(), 'logs')
            os.makedirs(log_dir, exist_ok=True)
            with open(os.path.join(log_dir, 'app.log'), 'a',
                      encoding='utf-8') as fh:
                fh.write(f'[{datetime.now():%Y-%m-%d %H:%M:%S}] '
                         '[INFO] 检测到已有实例在运行, 本次启动退出\n')
        except Exception:
            pass
        return None
    QCoreApplication.instance()  # 保持引用, 避免被 GC
    return lock


# ---------------------------------------------------------------- 日志/兜底
def _setup_logging():
    from PySide6.QtCore import qInstallMessageHandler, QtMsgType
    log_dir = os.path.join(data_dir(), 'logs')
    os.makedirs(log_dir, exist_ok=True)
    log_path = os.path.join(log_dir, 'app.log')

    def handler(msg_type, ctx, message):
        level = {QtMsgType.QtDebugMsg: 'DEBUG',
                 QtMsgType.QtInfoMsg: 'INFO',
                 QtMsgType.QtWarningMsg: 'WARN',
                 QtMsgType.QtCriticalMsg: 'ERROR',
                 QtMsgType.QtFatalMsg: 'FATAL'}.get(msg_type, 'INFO')
        try:
            with open(log_path, 'a', encoding='utf-8') as fh:
                from datetime import datetime
                ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                fh.write(f'[{ts}] [{level}] {message}\n')
        except Exception:
            pass

    qInstallMessageHandler(handler)
    return log_path


def _excepthook(etype, value, tb):
    detail = ''.join(traceback.format_exception(etype, value, tb))
    try:
        log_dir = os.path.join(data_dir(), 'logs')
        os.makedirs(log_dir, exist_ok=True)
        from datetime import datetime
        with open(os.path.join(log_dir, 'crash.log'), 'a',
                  encoding='utf-8') as fh:
            fh.write('\n===== [' +
                     datetime.now().strftime('%Y-%m-%d %H:%M:%S') +
                     '] UNCAUGHT =====\n' + detail + '\n')
    except Exception:
        pass
    try:
        QMessageBox.critical(None, APP_NAME,
                             '遇到未预期的错误, 程序将尝试继续运行。\n\n'
                             '错误详情已写入日志目录 logs/crash.log。\n'
                             + type(value).__name__ + ': ' + str(value))
    except Exception:
        pass


def main():
    log_path = _setup_logging()
    sys.excepthook = _excepthook

    # 高 DPI 策略必须在 QApplication 创建前设置, 否则无效。
    # Qt6 默认已启用 per-monitor DPI, 显式声明可避免在部分显卡驱动上
    # 出现窗口尺寸/DPI 缩放异常。
    from PySide6.QtCore import Qt
    from PySide6.QtGui import QGuiApplication
    QGuiApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)

    app = QApplication(sys.argv)
    # 单实例: 重复启动时提示并退出, 不再叠出一堆窗口
    lock = _acquire_single_instance()
    if lock is None:
        return

    # setStyle 在找不到样式时返回 None, 直接设置会让后续 setStyleSheet 行为异常
    style = QStyleFactory.create('Fusion')
    if style is not None:
        app.setStyle(style)
    # 关闭不透明窗口: 某些远程桌面/虚拟机环境下, 合成器缺 alpha 通道会把
    # 窗口画成全黑
    try:
        app.setAttribute(Qt.ApplicationAttribute.AA_UseCompositor, False)
        app.setAttribute(Qt.ApplicationAttribute.AA_OpaquePaintEvent, False)
    except Exception:
        pass
    from PySide6.QtGui import QFontDatabase
    for font in ('msyh.ttc', 'msyhbd.ttc', 'segoeui.ttf'):
        fp = os.path.join(os.environ.get('WINDIR', r'C:\Windows'), 'Fonts', font)
        if os.path.isfile(fp):
            QFontDatabase.addApplicationFont(fp)
    app.setApplicationName(APP_NAME)
    app.setOrganizationName('QingYan')
    app.setFont(QFont('Microsoft YaHei UI', 9))
    win = MainWindow()
    win._apply_theme_hooks()
    win.show()
    win.show_status('就绪 · 选择工具开始处理')
    win.status_label.setToolTip('运行日志：' + log_path)
    rc = app.exec()
    try:
        lock.unlock()
    except Exception:
        pass
    sys.exit(rc)


if __name__ == '__main__':
    main()
