# -*- coding: utf-8 -*-
"""青雁批量工具箱 v3 — 启动入口。

    python main.py                  # 直接运行
    pythonw main.py                 # 无控制台运行(打包用)
    青雁批量工具箱v3.exe --smoke     # 打包环境自检(安装后验证依赖完整性)
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app.main import main  # noqa: E402


def _smoke() -> int:
    """安装后自检: 验证打包产物里的第三方依赖真的可用。

    历史问题: 打包时漏掉 openpyxl/pypdf/reportlab, 导致装完后 Excel 与 PDF
    的全部功能崩溃。此处显式跑一遍, 有问题立刻暴露, 而不是等用户踩。
    """
    import smoke_test
    return smoke_test.main()


if __name__ == '__main__':
    if '--smoke' in sys.argv:
        sys.exit(_smoke())
    main()
