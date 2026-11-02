# 青雁批量工具箱 · QingYanToolbox

面向 Windows 的免费、离线文件批量处理工具。当前版本 **v3.4.1**，提供 22 个页面和 169 个操作入口。

集中处理 Word、Excel、PowerPoint、PDF、图片和文本，支持批量重命名、文件整理、格式转换、打印、任务流水线、备份和撤销。文档处理在本机进行。

## 下载与运行

在 [Releases 下载最新版](https://github.com/dissableq/QingYanToolbox/releases/latest)：

- `QingYanToolbox-版本号-setup.exe`：Windows 安装包。
- `QingYanToolbox-版本号-windows-x64.zip`：免安装版，完整解压后运行“青雁批量工具箱v3.exe”，保留 `_internal` 目录。
- `QingYanToolbox-版本号-source.zip`：对应版本的源码。
- `SHA256SUMS.txt`：上述文件的 SHA-256 校验值。

支持 Windows 10/11 64 位。软件未进行商业代码签名，Windows 可能显示发布者未知。

## 主要功能

| 分类 | 功能示例 |
| --- | --- |
| Word / Excel / PPT | 替换内容、清理空白和批注、图片与页眉页脚、样式和元数据、合并与拆分 |
| PDF | 拆分合并、提取页面、旋转裁剪、水印、加密解密、提取文本和图片 |
| 图片与文本 | 格式与编码转换、尺寸调整、压缩、水印、文本行处理 |
| 文件整理 | 批量重命名、复制、分类、清理重复文件、哈希与清单导出 |
| 统一替换 | 一次设置多组规则，处理文档内容、文件名和文件夹名 |
| 任务与恢复 | 流水线、参数方案、收藏、任务台账、备份管理、取消与撤销 |

## 统一替换与备份

选择目标文件夹，勾选要处理的内容、文件名和文件夹名，填写一组或多组规则，先预览再执行。可递归处理子文件夹，支持规则导入导出、取消和撤销。

原地修改前，备份保存到原文件旁的“备份”目录，每次保存独立版本。退出时可选择保留备份、清理本次备份或返回软件。清理后无法再通过已删除的备份恢复原文件内容。

内容替换支持现代 Word、Excel、PPT 和常见文本格式；PDF、图片及旧版 Office 文件在统一替换中仅处理名称。规则区分大小写，按顺序执行，Excel 公式保持不变。详见[使用说明](docs/usage.md)。

Office 保真转换需要本机安装 Office/WPS。PDF 转 PPT 以页面图片为主，图片转 Word 不包含 OCR。复杂文档和真实打印设备请使用对应样例验证。

## 从源码运行

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python main.py
```

推荐 Python 3.13。收藏、方案和任务台账保存在 `%APPDATA%/QingYanBatchToolbox`。

## 验证和构建

```powershell
python verify.py
python -m pip install -r requirements-build.txt
python scripts/release.py check
python scripts/release.py build
```

完整测试需要 Windows、本机 Office 和“Microsoft Print to PDF”。GitHub Actions 运行不依赖 Office 与打印机的 12 组测试；转换、真实设备和完整界面流程仍需本机验证。打包脚本会校验版本、构建程序与安装包、运行打包自检，并生成源码包和校验值。

构建使用正式的中文 `青雁批量工具箱v3.spec`。它限制构建时的 DLL 搜索路径，避免混入其他软件的运行库。

## 参与与许可

欢迎通过 [Issues](https://github.com/dissableq/QingYanToolbox/issues) 反馈问题。请附版本、复现步骤和经过脱敏的最小样例。

本项目采用 [GNU AGPL v3](LICENSE)。第三方组件保留各自许可证，见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。

[更新记录](CHANGELOG.md) · [贡献指南](CONTRIBUTING.md) · [版本发布流程](docs/releasing.md)
