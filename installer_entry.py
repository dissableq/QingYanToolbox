# -*- coding: utf-8 -*-
"""青雁批量工具箱 v3 — 安装器(由 PyInstaller 打成单文件 exe)。

为什么不直接用 Inno Setup: ISCC.exe 需从 GitHub 下载, 本机网络不可达
(InternetOpenUrl 0x80072efd)。本文件自实现等价安装流程。

安装向导(PySide6 实现, 提供常规安装体验):
  1. 欢迎页 —— 显示版本、占用空间、可选"安装后启动"
  2. 路径选择 —— 默认 %LOCALAPPDATA%\\Programs\\青雁批量工具箱v3, 可自选
  3. 安装前校验 —— 磁盘空间、可写性、覆盖已有安装确认
  4. 安装进度 —— 后台线程解压, 界面不卡死, 可取消
  5. 完成页 —— 建快捷方式、可立即启动

命令行(供自动化/无人值守):
    安装包.exe            # 交互式安装向导
    安装包.exe /S         # 静默安装到默认目录
    安装包.exe /S /D=路径  # 静默安装到指定目录
    安装包.exe /U         # 静默卸载
"""
from __future__ import annotations

import json
import hashlib
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import zipfile


APP_NAME = '青雁批量工具箱'
APP_DIR = '青雁批量工具箱v3'
VERSION = '3.4.0'
INSTALL_ROOT_DEFAULT = 'Programs'

MANIFEST = '.qingyan-install.json'

def _target_path(target):
    target = os.path.realpath(os.path.abspath(target))
    forbidden = [os.path.abspath(os.sep), os.path.expanduser('~'),
                 os.environ.get('WINDIR', r'C:\Windows'),
                 os.environ.get('ProgramFiles', r'C:\Program Files')]
    if os.path.dirname(target) == target or any(os.path.normcase(target) == os.path.normcase(x) for x in forbidden):
        raise ValueError('请选择独立的应用安装文件夹')
    return target

def _inside(target, name):
    from pathlib import PurePosixPath
    normalized = str(name).replace('\\', '/')
    parts = PurePosixPath(normalized).parts
    if not parts or normalized.startswith('/') or '..' in parts or ':' in normalized:
        raise ValueError('安装文件路径无效: ' + str(name))
    out = os.path.realpath(os.path.join(target, *parts))
    if os.path.commonpath([target, out]) != target or out == target:
        raise ValueError('安装文件超出目标目录: ' + str(name))
    return out

def desktop_dir():
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
            r'Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders') as key:
            return os.path.expandvars(winreg.QueryValueEx(key, 'Desktop')[0])
    except OSError:
        return os.path.join(os.environ.get('USERPROFILE', os.path.expanduser('~')), 'Desktop')



def _log(msg: str) -> None:
    """安装日志: 同时写文件与控制台, 便于排查 onefile 下的问题。"""
    try:
        print(msg, flush=True)
    except Exception:
        pass
    try:
        d = os.path.join(tempfile.gettempdir(), 'qingyan_install_log.txt')
        with open(d, 'a', encoding='utf-8') as fh:
            fh.write(msg + '\n')
    except Exception:
        pass


def appdata_local() -> str:
    return os.environ.get('LOCALAPPDATA') or os.path.expanduser('~')


def default_install_dir() -> str:
    return os.path.join(appdata_local(), INSTALL_ROOT_DEFAULT, APP_DIR)


def user_data_dir() -> str:
    return os.path.join(os.environ.get('APPDATA')
                        or appdata_local(), 'QingYanBatchToolbox')


def shortcut_dir() -> str:
    rd = os.environ.get('APPDATA') or appdata_local()
    return os.path.join(rd, 'Microsoft', 'Windows', 'Start Menu', 'Programs')


def free_space_mb(path: str) -> int:
    """目标盘剩余空间(MB); 取最近的已存在父目录。"""
    p = path
    while p and not os.path.isdir(p):
        parent = os.path.dirname(p)
        if parent == p:
            break
        p = parent
    try:
        return shutil.disk_usage(p).free // (1024 * 1024)
    except Exception:
        return 0


def _writable(path: str) -> bool:
    """目录是否可写(不实际创建文件)。"""
    p = path
    while p and not os.path.isdir(p):
        parent = os.path.dirname(p)
        if parent == p:
            break
        p = parent
    return os.path.isdir(p) and os.access(p, os.W_OK)


# ============================================================ 快捷方式
def _create_shortcut(exe_path: str, lnk_path: str, desc: str) -> None:
    """创建快捷方式。

    实现说明: 早期尝试用 ctypes 手工调 IShellLinkW 的 vtable, 但 vtable 索引
    极易写错(实测直接触发 access violation), 属于不可靠的底层代码。
    这里改用系统自带 PowerShell + WScript.Shell —— Windows 自带,
    不引入第三方依赖, 行为稳定。
    """
    def q(s: str) -> str:
        return str(s).replace("'", "''")

    ps = (
        "$s = New-Object -ComObject WScript.Shell; "
        f"$l = $s.CreateShortcut('{q(lnk_path)}'); "
        f"$l.TargetPath = '{q(os.path.abspath(exe_path))}'; "
        f"$l.WorkingDirectory = '{q(os.path.dirname(os.path.abspath(exe_path)))}'; "
        f"$l.Description = '{q(desc)}'; "
        "$l.Save()"
    )
    try:
        subprocess.run(
            ['powershell', '-NoProfile', '-NonInteractive',
             '-ExecutionPolicy', 'Bypass', '-Command', ps],
            capture_output=True, timeout=60,
            creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
    except Exception as e:
        _log(f'  快捷方式创建异常(可忽略): {e}')
    return os.path.exists(lnk_path)


def create_shortcuts(target: str, meta: dict,
                     desktop: bool = True) -> None:
    """建快捷方式。开始菜单始终创建; 桌面快捷方式按需。"""
    exe_path = os.path.join(target, meta.get('exe', f'{APP_DIR}.exe'))
    ver = meta.get('version', VERSION)
    try:
        sd = shortcut_dir()
        os.makedirs(sd, exist_ok=True)
        _create_shortcut(exe_path, os.path.join(sd, f'{APP_DIR}.lnk'),
                         f'{APP_NAME} v{ver}')
        _create_shortcut(os.path.join(target, '卸载.bat'),
                         os.path.join(sd, f'卸载 {APP_DIR}.lnk'),
                         f'卸载 {APP_NAME}')
        if desktop:
            dk = desktop_dir()
            if os.path.isdir(dk):
                _create_shortcut(exe_path, os.path.join(dk, f'{APP_DIR}.lnk'),
                                 f'{APP_NAME} v{ver}')
    except Exception as e:
        _log(f'快捷方式部分失败(不影响使用): {e}')
    try:
        import winreg
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, registry_key(target)) as key:
            values = {'DisplayName': APP_NAME, 'DisplayVersion': ver,
                      'InstallLocation': target, 'DisplayIcon': exe_path,
                      'UninstallString': 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "' +
                                         os.path.join(target, 'uninstall.ps1') + '"'}
            for name, value in values.items():
                winreg.SetValueEx(key, name, 0, winreg.REG_SZ, value)
    except OSError as exc:
        _log('系统应用列表登记失败：' + str(exc))


def registry_key(target):
    identity = hashlib.sha256(os.path.normcase(os.path.abspath(target)).encode('utf-8')).hexdigest()[:12]
    return r'Software\Microsoft\Windows\CurrentVersion\Uninstall\QingYan_' + identity


def shortcut_cleanup_script(target, desktop=True):
    def quote(value): return "'" + value.replace("'", "''") + "'"
    paths = [os.path.join(shortcut_dir(), f'{APP_DIR}.lnk'),
             os.path.join(shortcut_dir(), f'卸载 {APP_DIR}.lnk')]
    if desktop: paths.append(os.path.join(desktop_dir(), f'{APP_DIR}.lnk'))
    return ("$shell = New-Object -ComObject WScript.Shell\n" +
            "$shortcutRoot = " + quote(os.path.abspath(target).rstrip('\\') + '\\') + "\n" +
            "foreach ($link in @(" + ','.join(quote(p) for p in paths) + ")) {\n" +
            " if (Test-Path -LiteralPath $link) {\n" +
            "  $destination = $shell.CreateShortcut($link).TargetPath\n" +
            "  if ($destination.StartsWith($shortcutRoot, [StringComparison]::OrdinalIgnoreCase)) {\n" +
            "   Remove-Item -LiteralPath $link -Force\n  }\n }\n}\n" +
            "$key = " + quote('HKCU:\\' + registry_key(target)) + "\n" +
            "if (Test-Path -LiteralPath $key) { Remove-Item -LiteralPath $key -Force }\n")


def remove_shortcuts(target, desktop=True):
    subprocess.run(['powershell.exe', '-NoProfile', '-NonInteractive', '-Command',
                    shortcut_cleanup_script(target, desktop)], capture_output=True,
                   timeout=30, creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))


def write_uninstaller(target, meta, install_target=None):
    # PowerShell uses literal paths and validates every manifest entry before removal.
    script = r"""$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\')
$manifestPath = Join-Path $root '.qingyan-install.json'
try {
    $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($manifest.app -ne '青雁批量工具箱v3') { throw '安装清单无效' }
    $paths = @($manifest.files | ForEach-Object {
        $p = [IO.Path]::GetFullPath((Join-Path $root $_))
        if (-not $p.StartsWith($root + '\', [StringComparison]::OrdinalIgnoreCase)) { throw '路径无效' }
        $check = $p
        while ($check -ne $root) {
            if ((Test-Path -LiteralPath $check) -and
                ((Get-Item -LiteralPath $check -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)) {
                throw '安装目录内存在链接，请先移走链接再卸载'
            }
            $check = Split-Path -Parent $check
        }
        $p
    })
    foreach ($p in $paths) { if (Test-Path -LiteralPath $p -PathType Leaf) { Remove-Item -LiteralPath $p -Force } }
    # SHORTCUT_CLEANUP
    Remove-Item -LiteralPath $manifestPath -Force
    $folders = @($paths | ForEach-Object { Split-Path -Parent $_ } | Sort-Object -Unique) |
        Sort-Object Length -Descending
    foreach ($folder in $folders) {
        if ($folder.StartsWith($root + '\', [StringComparison]::OrdinalIgnoreCase) -and
            (Test-Path -LiteralPath $folder) -and @(Get-ChildItem -LiteralPath $folder -Force).Count -eq 0) {
            [IO.Directory]::Delete($folder)
        }
    }
    Write-Host '卸载完成。个人文件与用户数据已保留。'
} catch { Write-Host ('卸载未完成，请关闭应用后重试：' + $_.Exception.Message) }
Read-Host '按回车关闭'
"""
    script = script.replace('# SHORTCUT_CLEANUP', shortcut_cleanup_script(install_target or target))
    with open(os.path.join(target, '卸载.ps1'), 'w', encoding='utf-8-sig') as f:
        f.write(script)
    with open(os.path.join(target, '卸载.bat'), 'w', encoding='ascii') as f:
        f.write('@echo off\r\npowershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0'+
                'uninstall.ps1'+'"\r\n')
    # ASCII launcher avoids cmd code-page ambiguity for the script filename.
    os.replace(os.path.join(target, '卸载.ps1'), os.path.join(target, 'uninstall.ps1'))


def uninstall(target):
    target = _target_path(target)
    manifest_path = os.path.join(target, MANIFEST)
    try:
        with open(manifest_path, encoding='utf-8') as f:
            manifest = json.load(f)
    except (OSError, ValueError) as e:
        raise ValueError('无法确认安装目录，未删除任何文件') from e
    if manifest.get('app') != APP_DIR or not isinstance(manifest.get('files'), list):
        raise ValueError('安装清单无效')
    paths = [_inside(target, name) for name in manifest['files']]
    failures = []
    for path in paths:
        try:
            if os.path.isfile(path): os.remove(path)
        except OSError as e:
            failures.append(str(e))
    if failures:
        raise OSError('部分程序文件正在使用，请关闭程序后重试。安装清单已保留。')
    remove_shortcuts(target)
    os.remove(manifest_path)
    for dp, _, _ in os.walk(target, topdown=False):
        try: os.rmdir(dp)
        except OSError: pass
    _log('卸载完成，个人文件与用户数据已保留')


# ============================================================ 解压核心
def extract_payload(zip_path, target, progress_cb=None, should_cancel=None):
    target = _target_path(target)
    with zipfile.ZipFile(zip_path) as zf:
        items = zf.infolist()
        paths = [_inside(target, item.filename) for item in items]
        if len({os.path.normcase(p) for p in paths}) != len(paths):
            raise ValueError('安装包含重名文件')
        for item in items:
            if (item.external_attr >> 16) & 0o170000 == 0o120000:
                raise ValueError('安装包不支持符号链接')
        for index, (item, destination) in enumerate(zip(items, paths), 1):
            if should_cancel and should_cancel():
                return index - 1, True
            if item.is_dir():
                os.makedirs(destination, exist_ok=True)
            else:
                os.makedirs(os.path.dirname(destination), exist_ok=True)
                with zf.open(item) as source, open(destination, 'wb') as output:
                    shutil.copyfileobj(source, output)
            if progress_cb:
                progress_cb(index, len(items), item.filename)
        return len(items), False


def do_install(zip_path, target, meta, progress_cb=None, should_cancel=None,
               desktop_shortcut=True):
    target = _target_path(target)
    parent = os.path.dirname(target)
    os.makedirs(parent, exist_ok=True)
    # Both scratch directories are owned by this transaction, beside the target.
    stage = tempfile.mkdtemp(prefix='.qingyan-stage-', dir=parent)
    rollback = tempfile.mkdtemp(prefix='.qingyan-rollback-', dir=parent)
    installed = []
    retired = []
    existed = os.path.isdir(target)
    try:
        _, cancelled = extract_payload(zip_path, stage, progress_cb, should_cancel)
        if cancelled:
            return -1
        exe = meta.get('exe', f'{APP_DIR}.exe')
        if not os.path.isfile(_inside(stage, exe)):
            _log('安装包中缺少主程序')
            return 1
        if existed and os.listdir(target) and not os.path.isfile(os.path.join(target, exe)):
            raise ValueError('目标不是现有应用目录，请选择空文件夹或原安装目录')
        previous_names = []
        previous_manifest = os.path.join(target, MANIFEST)
        if os.path.isfile(previous_manifest):
            with open(previous_manifest, encoding='utf-8') as f:
                previous = json.load(f)
            if previous.get('app') != APP_DIR or not isinstance(previous.get('files'), list):
                raise ValueError('原安装清单无效，请选择新的空文件夹')
            previous_names = previous['files']
            for name in previous_names: _inside(target, name)
        write_uninstaller(stage, meta, target)
        names = [os.path.relpath(os.path.join(dp, fn), stage).replace('\\', '/')
                 for dp, _, fs in os.walk(stage) for fn in fs]
        with open(os.path.join(stage, MANIFEST), 'w', encoding='utf-8') as f:
            json.dump({'app': APP_DIR, 'version': meta.get('version', VERSION),
                       'files': names}, f, ensure_ascii=False)
        names.append(MANIFEST)
        # Cancellation is checked before committing. Once committing, finish or roll back.
        if should_cancel and should_cancel():
            return -1
        for name in previous_names:
            if name not in names:
                old_file = _inside(target, name)
                if os.path.isfile(old_file):
                    backup = _inside(rollback, name)
                    os.makedirs(os.path.dirname(backup), exist_ok=True)
                    os.replace(old_file, backup)
                    retired.append(name)
        for name in names:
            destination = _inside(target, name)
            backup = _inside(rollback, name)
            os.makedirs(os.path.dirname(destination), exist_ok=True)
            had_old = os.path.isfile(destination)
            if had_old:
                os.makedirs(os.path.dirname(backup), exist_ok=True)
                shutil.copy2(destination, backup)
            os.replace(_inside(stage, name), destination)
            installed.append((name, had_old))
        if meta.get('shortcuts', True):
            create_shortcuts(target, meta, desktop=desktop_shortcut)
        _log(f'安装完成: {target}')
        return 0
    except Exception as e:
        errors = []
        for name, had_old in reversed(installed):
            try:
                destination = _inside(target, name)
                if had_old:
                    os.replace(_inside(rollback, name), destination)
                elif os.path.isfile(destination):
                    os.remove(destination)
            except OSError as restore_error:
                errors.append(str(restore_error))
        for name in retired:
            try:
                os.replace(_inside(rollback, name), _inside(target, name))
            except OSError as restore_error:
                errors.append(str(restore_error))
        if errors:
            _log('回滚未完成，备份保留在: ' + rollback)
        _log(f'安装失败，旧文件已尝试还原: {e}')
        return 1
    finally:
        for scratch in (stage, rollback):
            # Never delete a directory other than the exact transaction directory.
            if os.path.dirname(os.path.realpath(scratch)) == parent:
                if scratch == rollback and 'errors' in locals() and errors:
                    continue
                shutil.rmtree(scratch)
        if not existed and os.path.isdir(target):
            for dp, _, _ in os.walk(target, topdown=False):
                try: os.rmdir(dp)
                except OSError: pass


# ============================================================ 安装向导 GUI
# Signal 必须在模块级导入, 且承载它的类必须继承 QObject —— 否则在实例上
# 访问到的仍是未绑定的 Signal 描述符, connect() 会抛
# "Signal object has no attribute 'connect'"。
try:
    from PySide6.QtCore import QObject, Signal
except Exception as exc:       # 未装 PySide6 时仍可做静默安装
    _log('界面组件加载失败：' + str(exc))
    QObject = object
    def Signal(*args):
        return None


class InstallWizard(QObject):
    """PySide6 安装向导: 路径可选、空间校验、后台解压、进度与取消。

    跨线程更新 UI 用 Qt 官方的信号槽机制(自动队列连接, 线程安全)。
    注意: **不要用 QApplication.call_later_threadsafe** —— 那是 PyQt 的 API,
    PySide6 没有这个属性, 调用会抛 AttributeError, 导致进度条永远停在 0%
    (异常还会被 worker 的 except 吞掉, 表现为"装完了但进度不动")。
    """

    # 后台线程 -> 主线程
    sig_progress = Signal(int, int, str)     # done, total, 当前文件
    sig_finished = Signal(str, int)          # target, rc

    def __init__(self, zip_path: str, meta: dict, splash=None):
        from PySide6.QtWidgets import QApplication
        # QObject 子类必须先初始化 Qt 侧对象, 否则访问 Signal 时会报
        # "Signal source has been deleted"
        super().__init__()
        self.zip_path = zip_path
        self.meta = meta
        self.target = default_install_dir()
        self.install_size_mb = 0
        self._cancel = False
        self._done = False
        self._installing = False
        self._rc = 0

        self.app = QApplication.instance() or QApplication(sys.argv)
        # 主界面就绪后关掉启动提示
        if splash is not None:
            try:
                splash.close()
            except Exception:
                pass
        self._build()

    # ---------------------------------------------------------- UI 构建
    def _build(self):
        from PySide6.QtWidgets import (QCheckBox, QFrame, QHBoxLayout,
                                       QLabel, QLineEdit, QProgressBar,
                                       QPushButton, QVBoxLayout, QWidget)

        self.sig_progress.connect(self._on_progress)
        self.sig_finished.connect(self._on_finish)

        wizard = self
        class WizardWindow(QWidget):
            def closeEvent(self, event):
                if wizard._installing:
                    wizard._cancel = True
                    wizard.lbl_status.setText('正在取消，请等待文件安全还原…')
                    event.ignore()
                else:
                    event.accept()
        self.w = WizardWindow()
        self.w.setWindowTitle(f'{APP_NAME} v{VERSION} 安装程序')
        self.w.setStyleSheet(_wizard_stylesheet())
        self.w.setMinimumSize(680, 620)
        self.w.resize(720, 680)
        v = QVBoxLayout(self.w)
        v.setContentsMargins(26, 22, 26, 20)
        v.setSpacing(14)

        # 标题
        t = QLabel(f'{APP_NAME}  v{VERSION}')
        t.setProperty('role', 'title')
        v.addWidget(t)
        sub = QLabel('让重复工作，在这里一次完成。' )
        sub.setProperty('role', 'muted')
        v.addWidget(sub)

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        v.addWidget(line)

        # 路径选择
        self.lbl_path = QLabel('安装到')
        self.lbl_path.setProperty('role', 'section')
        v.addWidget(self.lbl_path)
        row = QHBoxLayout()
        row.setSpacing(8)
        self.ed_path = QLineEdit(self.target)
        self.ed_path.setPlaceholderText('选择安装目录, 留空则使用默认位置')
        self.ed_path.textChanged.connect(self._on_path_changed)
        btn_browse = QPushButton('浏览…')
        self.btn_browse = btn_browse
        btn_browse.setFixedWidth(88)
        btn_browse.clicked.connect(self._browse)
        row.addWidget(self.ed_path, 1)
        row.addWidget(btn_browse)
        v.addLayout(row)

        # 空间信息
        self.lbl_space = QLabel('')
        self.lbl_space.setWordWrap(True)
        self.lbl_space.setProperty('role', 'muted')
        v.addWidget(self.lbl_space)

        # 功能清单
        info = QLabel('包含: Word / Excel / PPT / PDF / 文本 / 图片 / 重命名 / '
                      '文件整理 · 169 项批量操作' )
        info.setProperty('role', 'muted')
        info.setWordWrap(True)
        v.addWidget(info)
        tip = QLabel(f'用户数据(收藏 / 参数方案 / 任务台账)保存在:\n{user_data_dir()}\n'
                     '卸载时不会删除, 重新安装后自动恢复。')
        tip.setProperty('role', 'tip')
        tip.setWordWrap(True)
        v.addWidget(tip)

        # 进度
        self.pg = QProgressBar()
        self.pg.setVisible(False)
        self.pg.setFixedHeight(18)
        v.addWidget(self.pg)
        self.lbl_status = QLabel('')
        self.lbl_status.setWordWrap(True)
        self.lbl_status.setProperty('role', 'muted')
        v.addWidget(self.lbl_status)

        # 勾选项
        self.cb_launch = QCheckBox('安装完成后启动程序')
        self.cb_launch.setChecked(True)
        v.addWidget(self.cb_launch)
        self.cb_shortcut = QCheckBox('创建桌面快捷方式')
        self.cb_shortcut.setChecked(True)
        v.addWidget(self.cb_shortcut)

        v.addStretch(1)

        # 按钮
        brow = QHBoxLayout()
        brow.addStretch(1)
        self.btn_cancel = QPushButton('取消')
        self.btn_cancel.clicked.connect(self._on_cancel)
        self.btn_ok = QPushButton('开始安装')
        self.btn_ok.setProperty('accent', True)
        self.btn_ok.clicked.connect(self._on_install)
        brow.addWidget(self.btn_cancel)
        brow.addWidget(self.btn_ok)
        v.addLayout(brow)

        self._update_space()
        self.w.show()

    def _browse(self):
        from PySide6.QtWidgets import QFileDialog
        d = QFileDialog.getExistingDirectory(self.w, '选择安装位置',
                                             self.ed_path.text())
        if d:
            self.ed_path.setText(os.path.join(d, APP_DIR))

    def _on_path_changed(self, text):
        self.target = text.strip()
        self._update_space()

    def _update_space(self):
        need = self.install_size_mb or (self._payload_mb() * 3 + 40)
        free = free_space_mb(self.target or default_install_dir())
        ok_w = _writable(self.target or default_install_dir())
        parts = [f'需要空间约 {need} MB',
                 f'目标盘可用 {free} MB']
        if free and free < need:
            parts.append('⚠ 空间不足, 请更换安装位置')
        if not ok_w:
            parts.append('⚠ 该位置不可写, 请更换安装位置')
        if os.path.isdir(self.target or '') and os.listdir(self.target or ''):
            parts.append('该目录已有文件, 将执行覆盖安装')
        self.lbl_space.setText('　'.join(parts))
        self.lbl_space.setProperty('role', 'warn' if (free < need or not ok_w)
                                   else 'muted')
        self._refresh_btn()

    def _payload_mb(self) -> int:
        try:
            with zipfile.ZipFile(self.zip_path) as zf:
                return sum(item.file_size for item in zf.infolist()) // (1024 * 1024)
        except (OSError, zipfile.BadZipFile):
            return 0

    def _refresh_btn(self):
        free = free_space_mb(self.target or default_install_dir())
        need = self.install_size_mb or (self._payload_mb() * 3 + 40)
        self.btn_ok.setEnabled(bool(self.target) and _writable(self.target)
                               and free >= need)

    # ---------------------------------------------------------- 安装流程
    def _on_cancel(self):
        if self._done:
            self.app.quit()
            return
        if self._installing:
            self._cancel = True
            self.lbl_status.setText('正在取消…')
            return
        # 未开始 -> 确认后退出
        self.w.close()

    def _on_install(self):
        if self._done and self._rc == 0:
            if self.cb_launch.isChecked():
                try:
                    os.startfile(os.path.join(self.target, self.meta.get('exe', f'{APP_DIR}.exe')))
                except OSError as exc:
                    from PySide6.QtWidgets import QMessageBox
                    QMessageBox.warning(self.w, APP_NAME, '程序已安装，启动失败：' + str(exc))
                    return
            self.app.quit()
            return
        target = self.ed_path.text().strip()
        if not target:
            return
        if os.path.isdir(target) and os.listdir(target):
            from PySide6.QtWidgets import QMessageBox
            r = QMessageBox.question(
                self.w, f'{APP_NAME} 安装程序',
                f'检测到该目录已有文件:\n{target}\n\n'
                '继续将覆盖其中的同名文件, 是否继续?',
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if r != QMessageBox.Yes:
                return
        self._run_install(target)

    def _run_install(self, target: str):
        self.target = target
        self._done = False
        self._installing = True
        self._cancel = False
        desktop = self.cb_shortcut.isChecked()
        self.btn_ok.setEnabled(False)
        self.btn_cancel.setText('取消安装')
        self.ed_path.setEnabled(False)
        self.btn_browse.setEnabled(False)
        self.cb_shortcut.setEnabled(False)
        self.pg.setVisible(True)
        self.pg.setRange(0, 100)
        self.pg.setValue(0)
        self.lbl_status.setText('正在准备解压…')

        def worker():
            try:
                rc = do_install(
                    self.zip_path, target, self.meta,
                    progress_cb=lambda d, t, n: self.sig_progress.emit(d, t, n),
                    should_cancel=lambda: self._cancel,
                    desktop_shortcut=desktop)
            except Exception as e:
                _log(f'安装异常: {type(e).__name__}: {e}')
                rc = 1
            self.sig_finished.emit(target, rc)

        th = threading.Thread(target=worker, daemon=True)
        th.start()
        self._th = th

    def _on_progress(self, done, total, name):
        pct = int(done * 100 / total) if total else 0
        self.pg.setValue(pct)
        self.lbl_status.setText(
            f'正在解压 {pct}%  ({done}/{total})  {os.path.basename(name)[:36]}')

    def _on_finish(self, target, rc):
        self._installing = False
        self._rc = rc
        self._done = rc == 0
        self.btn_ok.setEnabled(True)
        self.btn_cancel.setText('关闭')
        if rc != 0:
            self.ed_path.setEnabled(True)
            self.btn_browse.setEnabled(True)
            self.cb_shortcut.setEnabled(True)
            self.btn_ok.setText('重新安装')
            self.lbl_status.setText('安装已取消，原有文件已保留。' if rc == -1 else
                '安装未完成。请关闭正在运行的青雁后重试，或选择其他位置。\n日志：' +
                os.path.join(tempfile.gettempdir(), 'qingyan_install_log.txt'))
            return
        self.pg.setValue(100)
        self.lbl_status.setText(f'安装完成 · v{VERSION}\n{target}\n点击「完成」即可开始使用。')
        self.btn_ok.setText('完成')

    def run(self) -> int:
        self.app.exec()
        return self._rc


def _wizard_stylesheet() -> str:
    return '''
QWidget { font-family: "Microsoft YaHei UI"; font-size: 13px; color: #1A2430; }
QWidget[role="title"] { font-size: 22px; font-weight: 700; color: #635BDF; }
QLabel[role="muted"] { color: #7A858C; font-size: 12px; }
QLabel[role="warn"] { color: #B36A00; font-size: 12px; }
QLabel[role="section"] { font-weight: 700; font-size: 14px; color: #1A2430;
                         padding-top: 4px; }
QLabel[role="tip"] { color: #5A6B75; font-size: 11px; background:#F5F5FA;
                     border:1px solid #E2E7F0; border-radius:14px;
                     padding:8px 10px; }
QLineEdit { border:1px solid #E2E7F0; border-radius:14px; padding:7px 10px;
            background:#FFFFFF; }
QPushButton { border:1px solid #C6D8E2; border-radius:14px; padding:7px 20px;
              background:#FFFFFF; }
QPushButton:hover { border-color:#635BDF; background:#F1F7FA; }
QPushButton[accent="true"] { background:#635BDF; color:white; border:none;
                             font-weight:600; padding:8px 26px; }
QPushButton[accent="true"]:hover { background:#5149C8; }
QPushButton:disabled { color:#7A858C; border-color:#E2E7F0; background:#F5F5FA; }
QCheckBox { spacing:8px; }
QProgressBar { border:none; border-radius:14px; background:#F5F5FA; height:16px;
               text-align:center; font-size:11px; color:#7A858C; }
QProgressBar::chunk { background:#635BDF; border-radius:14px; }
QFrame { color:#E2E7F0; }
'''


# ============================================================ 入口
def _load_embedded(base: str) -> tuple:
    """读取内嵌的 payload.zip 与 meta.json。"""
    pf = os.path.join(base, 'payload.zip')
    mf = os.path.join(base, 'meta.json')
    if not os.path.exists(pf):
        return '', '{}'
    meta = '{}'
    if os.path.exists(mf):
        with open(mf, encoding='utf-8') as fh:
            meta = fh.read().strip()
    return pf, meta


def _show_splash(app):
    """显示"正在启动"提示窗, 返回该窗口。

    onefile 启动时要先把 43MB 的 PySide6 解到临时目录, 这期间界面还没出来,
    用户会以为程序卡死。先弹一个短提示, 让用户知道程序活着。

    注意: 不能用子进程 —— onefile 下子进程会把运行库再解压一遍, 更慢。
    这里直接在主进程建窗, 由 InstallWizard 复用同一个 QApplication。
    """
    try:
        from PySide6.QtCore import Qt
        from PySide6.QtWidgets import QLabel, QVBoxLayout, QWidget

        w = QWidget()
        w.setWindowTitle(f'{APP_NAME} v{VERSION} 安装程序')
        w.setFixedWidth(420)
        w.setStyleSheet(_wizard_stylesheet())
        v = QVBoxLayout(w)
        v.setContentsMargins(24, 20, 24, 20)
        v.addSpacing(10)
        t = QLabel('正在启动安装程序…')
        t.setProperty('role', 'title')
        t.setAlignment(Qt.AlignCenter)
        v.addWidget(t)
        tip = QLabel('首次运行需先解压运行库，请稍候')
        tip.setProperty('role', 'muted')
        tip.setAlignment(Qt.AlignCenter)
        v.addWidget(tip)
        v.addSpacing(10)
        w.show()
        app.processEvents()
        return w
    except Exception as e:
        _log(f'splash 跳过: {e}')
        return None


def main(argv=None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    # 取开关名: '/S'、'-S'、'S:/'(Git Bash 转过的形态)、'--silent' 都归一为 'S'。
    # 只取开关本身, 不动路径参数, 避免把 C:\ 变成 C:// 。
    def flag(a: str) -> str:
        s = str(a).replace('\\', '/')
        return s.split(':', 1)[0].strip('/').lstrip('-').upper()

    silent = any(flag(a) in ('S', 'SILENT') for a in argv)
    do_uninstall = any(flag(a) in ('U', 'UNINSTALL') for a in argv)

    # /DIR=<path>(本安装器的正式开关), 兼容 /D= /-D= 形式。
    # 历史上裸 /D 分支会把 /DIR= 吞成 "IR=..." 当路径, 静默安装必失败;
    # 现在所有分支都要求等号, 且 /DIR= 单独列出。
    custom = ''
    for a in argv:
        s = str(a).strip('"')
        m = re.match(r'^(?:/DIR=|-DIR=|/D=|-D=)(.+)$', s, re.IGNORECASE)
        if m:
            custom = m.group(1).strip().strip('"')
            break
    if custom:
        # Git Bash 可能传 '/C:/Temp/QY' 这类带前导斜杠的路径, 统一规整。
        # 用标准库而非手写正则: 正则很容易把盘符/中文/反斜杠搞坏。
        if re.match(r'^/[A-Za-z]:', custom):
            custom = custom[1:]
        custom = os.path.normpath(custom)
        _log(f'custom target={custom}')

    base = getattr(sys, '_MEIPASS', None) or os.path.dirname(sys.executable)
    zip_path, meta_json = _load_embedded(base)
    if not zip_path:
        print('这是安装器内部组件, 请运行完整的「安装包.exe」。')
        return 1
    meta = json.loads(meta_json)
    if any(flag(a) == 'NOSHORTCUTS' for a in argv):
        meta['shortcuts'] = False

    # ---- 卸载
    if do_uninstall:
        tgt = custom or default_install_dir()
        uninstall(tgt)
        return 0

    # ---- 静默安装
    if silent:
        tgt = custom or default_install_dir()
        if not _writable(os.path.dirname(tgt) or tgt):
            print(f'目标目录不可写: {tgt}')
            return 1
        _log(f'silent install -> {tgt}')
        rc = do_install(zip_path, tgt, meta)
        print('installed' if rc == 0 else f'failed rc={rc}')
        return rc

    # ---- 交互式安装向导
    try:
        from PySide6.QtWidgets import QApplication
        _app = QApplication.instance() or QApplication(sys.argv)
        _sp = _show_splash(_app)
        wiz = InstallWizard(zip_path, meta, splash=_sp)
        if custom:
            wiz.ed_path.setText(custom)
        return wiz.run()
    except ImportError as e:
        _log(f'向导不可用({e}), 回退到静默安装')
        rc = do_install(zip_path, custom or default_install_dir(), meta)
        return rc


if __name__ == '__main__':
    sys.exit(main())
