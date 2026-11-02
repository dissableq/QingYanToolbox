# 迭代与发布

每次完成修改后，先验证，再提交并推送 GitHub。文档修改可只推送 `main`；功能或问题修复使用新的版本号及 `vX.Y.Z` 标签。

## 维护者流程

1. 完成修改和对应验证，更新 `CHANGELOG.md`、使用说明及 `docs/releases/vX.Y.Z.md`。
2. 执行 `python scripts/release.py bump X.Y.Z` 同步四处产品版本，并核对差异。
3. 在配有 Office 和虚拟打印机的 Windows 环境运行 `python verify.py`，对涉及的界面和输出文件进行实际复查。
4. 审阅 `git diff` 与新增文件，确认公开内容仅包含项目源码和文档。
5. 执行 `scripts/publish.ps1 -Version X.Y.Z -Message '描述此次变化' -Python 'python解释器路径'`。脚本再次验证，创建提交与标签，原子推送 `main` 和标签。
6. 查看 GitHub Actions。标签验证和构建通过后，自动发布安装包、免安装包、源码包与 SHA-256 文件。

`scripts/publish.ps1 -DocumentationOnly -Message '更新说明'` 用于不改变产品版本的文档更新。

发布失败时检查 Actions 日志并修复；已经发布的版本不重写标签。不要强制推送覆盖远端历史。登录过期时使用 GitHub CLI 的 `gh auth login`，不把令牌写进仓库。

## 本地构建

```powershell
python -m pip install -r requirements-build.txt
python scripts/release.py check --tag v3.4.0
python scripts/release.py build --tag v3.4.0
```

必须在已提交且工作区干净的源码中构建，源码包由当前 Git 提交生成，避免安装包与源码不一致。产物输出到 `release/`。构建会从 PyInstaller 产物运行自检，并将许可证放进免安装包与安装负载。

## 自动检查范围

GitHub Windows Runner 运行 12 组不依赖 Office 和打印机的测试。`test_convert.py`、`test_print_virtual.py`、`test_surface.py` 留给具备对应软件和设备的本机完整验证。CI 通过不能替代这些环境相关检查。
