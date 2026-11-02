# -*- coding: utf-8 -*-
"""青雁批量工具箱 v3 — 安装包制作驱动。

流程:
1) 用 PyInstaller 把 dist 产物打成 exe(见 QingYanToolboxV3.spec)
2) 把该目录打成 payload.zip, base64 编码后内嵌进安装器
3) 用 PyInstaller 把 installer_entry.py 打成单文件 安装包.exe

之所以自制安装器: Inno Setup 的 ISCC.exe 需从 GitHub 下载, 本机网络不可达。

用法:
    python make_installer.py --dist dist_new2/青雁批量工具箱v3
"""
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import zipfile

APP_NAME = '青雁批量工具箱'
APP_DIR = '青雁批量工具箱v3'
VERSION = '3.4.0'


def _run(cmd, **kw):
    print('$', ' '.join(cmd[:6]), '...')
    environment = os.environ.copy()
    windows = environment.get('WINDIR', r'C:\Windows')
    environment['PATH'] = os.pathsep.join([os.path.dirname(sys.executable),
                                          os.path.join(windows, 'System32'), windows])
    kw.setdefault('env', environment)
    return subprocess.run(cmd, check=True, **kw)


def zip_dist(dist_dir: str, out_zip: str) -> int:
    """把 dist 目录打成 zip(相对路径保留目录结构)。"""
    n = 0
    with zipfile.ZipFile(out_zip, 'w', zipfile.ZIP_DEFLATED,
                         compresslevel=9) as zf:
        for dp, _dns, fns in os.walk(dist_dir):
            for fn in fns:
                fp = os.path.join(dp, fn)
                rel = os.path.relpath(fp, dist_dir)
                zf.write(fp, rel.replace('\\', '/'))
                n += 1
    return n


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument('--dist', default='dist/' + APP_DIR)
    ap.add_argument('--out', default='安装包')
    ap.add_argument('--work', default='build_installer')
    args = ap.parse_args()

    root = os.path.dirname(os.path.abspath(__file__))
    dist_dir = os.path.join(root, args.dist.replace('/', os.sep))
    if not os.path.isdir(dist_dir):
        print(f'找不到打包产物: {dist_dir}')
        return 1
    exe_name = f'{APP_DIR}.exe'
    exe_path = os.path.join(dist_dir, exe_name)
    if not os.path.exists(exe_path):
        print(f'产物里没有 {exe_name}')
        return 1

    # 陈旧主 exe 检查: 本脚本只打 payload 与安装器, 不重打主 exe。若源码在
    # 主 exe 之后改过(典型: 改完 convert_ops/smoke_test 只跑 make_installer),
    # 装进去的还是旧逻辑, 而这种"改了半天没生效"极难发现。直接拦住。
    src_roots = [os.path.join(root, 'engine'), os.path.join(root, 'app'),
                 os.path.join(root, 'main.py'),
                 os.path.join(root, 'smoke_test.py')]
    exe_mtime = os.path.getmtime(exe_path)
    stale = []
    for sr in src_roots:
        if os.path.isfile(sr):
            fs = [sr]
        else:
            fs = [os.path.join(dp, f)
                  for dp, _d, fns in os.walk(sr) for f in fns
                  if f.endswith('.py') and '__pycache__' not in dp]
        for fp in fs:
            if os.path.getmtime(fp) > exe_mtime:
                stale.append(os.path.relpath(fp, root))
    if stale:
        print(f'主 exe 比以下源码旧(共 {len(stale)} 个), 先重打主 exe:')
        for f in stale[:8]:
            print('   ', f)
        if len(stale) > 8:
            print(f'    ...等 {len(stale) - 8} 个')
        print('    pyinstaller --noconfirm --clean --distpath dist '
              '--workpath build/pyi 青雁批量工具箱v3.spec')
        return 1

    os.makedirs(args.out, exist_ok=True)
    payload_zip = os.path.join(root, args.work.replace('/', os.sep),
                               'payload.zip')
    os.makedirs(args.work, exist_ok=True)

    print('==> 1/3 打包 payload.zip ...')
    n = zip_dist(dist_dir, payload_zip)
    size_zip = os.path.getsize(payload_zip)
    print(f'    {n} 个文件, zip {size_zip / 1048576:.1f} MB')

    print('==> 2/3 生成安装器 spec ...')
    meta = {'version': VERSION, 'exe': exe_name, 'app': APP_DIR}
    with open(os.path.join(root, args.work.replace('/', os.sep), 'meta.json'),
              'w', encoding='utf-8') as fh:
        json.dump(meta, fh, ensure_ascii=False)

    # 安装器本身不含业务逻辑, 用 onefile 打成单 exe
    # 安装器需要 PySide6 做图形化安装向导(选路径/进度/取消)
    spec_onefile = os.path.join(root, args.work.replace('/', os.sep),
                                'installer_onefile.spec')
    entry = os.path.join(root, 'installer_entry.py')
    meta_json = os.path.join(root, args.work.replace('/', os.sep),
                             'meta.json')
    with open(spec_onefile, 'w', encoding='utf-8') as fh:
        fh.write(
            "# -*- mode: python ; coding: utf-8 -*-\n"
            "# 安装器自带 PySide6 以提供图形化安装向导(选安装路径/进度/取消)\n"
            "a = Analysis([r'%s'],\n"
            "    pathex=[r'%s'],\n"
            "    binaries=[],\n"
            "    datas=[(r'%s', '.'), (r'%s', '.')],\n"
            "    hiddenimports=['PySide6.QtWidgets', 'PySide6.QtGui',\n"
            "                  'PySide6.QtCore'],\n"
            "    hookspath=[], hooksconfig={}, runtime_hooks=[],\n"
            "    excludes=['PySide6.QtWebEngineCore', 'PySide6.QtWebEngineWidgets',\n"
            "             'PySide6.QtQuick', 'PySide6.QtQml', 'PySide6.Qt3DCore',\n"
            "             'PySide6.QtMultimedia', 'PySide6.QtCharts',\n"
            "             'PySide6.QtDataVisualization', 'PySide6.QtNetworkAuth'],\n"
            "    noarchive=False, optimize=0)\n"
            "pyz = PYZ(a.pure)\n"
            "exe = EXE(pyz, a.scripts, a.binaries, a.datas, [],\n"
            "    name='%s_v%s_安装包',\n"
            "    debug=False, bootloader_ignore_signals=False, strip=False,\n"
            "    upx=False, console=False, disable_windowed_traceback=False,\n"
            "    icon=[r'%s\\icon.ico'])\n"
            % (entry, root, payload_zip, meta_json,
               APP_DIR.replace(' ', ''), VERSION, root))

    print('==> 3/3 用 PyInstaller 打包单文件安装器 ...')
    out_dir = os.path.join(root, args.out.replace('/', os.sep))
    os.makedirs(out_dir, exist_ok=True)
    _run([sys.executable, '-m', 'PyInstaller', '--noconfirm', '--clean',
          '--distpath', out_dir,
          '--workpath', os.path.join(root, args.work.replace('/', os.sep),
                                     'pyi'),
          spec_onefile])

    final = os.path.join(out_dir,
                         f'{APP_DIR.replace(" ", "")}_v{VERSION}_安装包.exe')
    if not os.path.exists(final):
        cand = [f for f in os.listdir(out_dir) if f.endswith('.exe')]
        print('产物目录内容:', cand)
        return 1

    print('\n==> 校验安装包 ...')
    sz = os.path.getsize(final)
    print(f'    文件: {final}')
    print(f'    大小: {sz / 1048576:.1f} MB')
    print(f'    (payload zip {size_zip / 1048576:.1f} MB, '
          f'{n} 个文件)')
    return 0


if __name__ == '__main__':
    sys.exit(main())
