# 文件搜索工具（标题 + 正文）v8

本地 Windows 文件管理工具，纯 Python 编写，双击 exe 即可使用，无需安装 Python 环境。

> **仓库**：https://github.com/BryanKo-CHN/file-search-pro （私有）
> **状态**：v8 源码与文档已发布；exe 请从 Releases 下载（若尚未发布，见 `RELEASE.md`）。
> **v8 更新**：老式 .doc/.wps 正文改用 OLE2 分片表精确提取，中文不再乱码。

## 功能

- **功能 1｜目录扫描与分类导出**：扫描指定硬盘/文件夹，按后缀分类，导出可折叠目录树（HTML/CSV/TXT）。
- **功能 2｜标题 + 正文搜索**：文件名模糊/精确/正则搜索，PDF/DOCX/PPTX/XLSX 等正文全文搜索，AND/OR/NOT 高级筛选、大小/日期/后缀过滤、分页、跨页全选、结果导出。
- **浏览文件夹**：可视化选目录，支持关键词跨层级模糊定位。
- **文件操作**：结果区行内系统复制/剪切、粘贴，目录树超链接直接打开文件。
- **预览面板**：左侧保留原功能，右侧预览区可拖拽宽度、可关闭。支持图片、PDF 逐页、文本、docx/pptx/xlsx 原排版、老式 Office 文字降级。
- **Office/WPS 兼容**：自动探测本机 Microsoft Office 或 WPS，调用原版引擎导出 PDF 保留排版；两者都没有时自动回退到内置渲染，功能不缺失。
- **安全**：只读扫描，不修改原文件。

## 源码与构建

| 文件 | 说明 |
|------|------|
| `file_search_pro.py` | 最终主程序 v7（标题+正文搜索、预览、导出） |
| `dir_scanner.py` | 程序 1：目录扫描导出 |
| `file_classifier.py` | 程序 2：后缀分类、复制/剪切/粘贴 |
| `file_searcher.py` | 程序 3：文件名搜索 |
| `FileSearchPro.spec` | PyInstaller 打包配置（build5） |
| `交接文档.md` | 完整交接文档（含终局总览、17 节功能说明、27 条避坑经验） |
| `RELEASE.md` | exe 手动发布步骤与校验值 |

构建命令（需 PyInstaller 与依赖，见 `交接文档.md`）：

```
pyinstaller FileSearchPro.spec --clean --noconfirm
```

## exe 下载

见本仓库 Releases 页面。若尚未发布，按 `RELEASE.md` 的步骤上传。

- v8 主程序 exe：121,714,910 字节
- MD5：`c4e050bdbf1cb04c5b0c546f57807755`

## 使用

双击 `文件搜索工具(标题+正文).exe`，浏览器会自动打开本地界面（默认端口 8812）。关闭窗口即退出，不写注册表、不改系统设置。

---

生成日期：2026-11-22
