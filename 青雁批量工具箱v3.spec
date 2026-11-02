# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files
import os
import sys
# Only resolve native libraries from Python/Windows; unrelated tools on PATH can
# shadow the Windows ICU and UCRT DLLs and produce an executable that cannot start.
windows = os.environ.get('WINDIR', r'C:\Windows')
os.environ['PATH'] = os.pathsep.join([os.path.dirname(sys.executable),
                                    os.path.join(windows, 'System32'), windows])


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('icon.ico', '.'), ('app/assets', 'app/assets')] + collect_data_files('docxcompose'),
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='青雁批量工具箱v3',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version='version_info.py',
    icon=['icon.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='青雁批量工具箱v3',
)
