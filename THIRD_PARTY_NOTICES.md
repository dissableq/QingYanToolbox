# 第三方组件

青雁批量工具箱源码采用 GNU AGPL v3。下列组件保留各自版权与许可；安装依赖时请同时保留其许可证。

| 组件 | 版本 | 许可 |
| --- | --- | --- |
| PySide6 / Qt for Python / shiboken6 | 6.11.2 | LGPL-3.0 / GPL 许可选项，详见上游 |
| PyMuPDF / MuPDF | 1.28.2 | GNU AGPL v3；另有上游商业授权 |
| python-docx | 1.2.0 | MIT |
| docxcompose | 2.2.0 | MIT |
| openpyxl | 3.1.5 | MIT |
| python-pptx | 1.0.2 | MIT |
| pypdf | 6.19.0 | BSD-3-Clause |
| pycryptodome | 4.2.0 | BSD-2-Clause / Unlicense / MIT |
| Pillow | 12.3.0 | MIT-CMU |
| reportlab | 5.0.1 | BSD |
| pywin32 | 312 | PSF，部分组件另有许可 |
| charset-normalizer | 3.5.2 | MIT |
| chardet | 7.6.0 | 0BSD |
| lxml | 6.1.3 | BSD-3-Clause，含上游库许可 |

精确版本由 `requirements.txt` 固定。发布脚本将已安装运行依赖及其传递依赖的许可证、版本与上游链接收集到发行版的 `licenses/` 目录；应以各组件随附的许可原文为准。

Python 运行时使用 Python Software Foundation License。PyInstaller 使用 GPL 许可及其引导程序分发例外，详见其上游许可。

对应源码可通过同一 Release 的源码包获取。Qt/PySide 的动态库在 `_internal` 中随包分发；重新构建方法见 README。项目不授予第三方商标权。
