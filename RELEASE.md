# v8 exe 发布步骤

GitHub 仓库单文件上限为 100 MB，v8 的 exe 是 116 MiB，因此 exe 不能作为普通文件提交，必须通过 **Release 附件**上传（Release 附件上限 2 GB）。

## 文件信息

- 主程序：`dist/文件搜索工具(标题+正文).exe`
- 大小：121,714,910 字节（约 116 MiB）
- MD5：`c4e050bdbf1cb04c5b0c546f57807755`
- 英文名副本 `FileSearchPro.exe` 内容完全相同（MD5 一致）

## 网页上传步骤

1. 打开仓库 `https://github.com/BryanKo-CHN/file-search-pro`
2. 右侧点击 **Releases** → **Create a new release**
3. Tag 填 `v8.0.0`，标题填 `v8 — 老式 .doc 精确提取版`
4. 说明可写：
   ```
   v8 更新：老式 .doc/.wps 正文改用 OLE2 分片表精确提取，中文不再乱码；
   v7 的 Office/WPS 双引擎兼容与原排版预览全部保留。
   MD5: c4e050bdbf1cb04c5b0c546f57807755
   ```
5. 拖入 `文件搜索工具(标题+正文).exe`（可同时拖入 `FileSearchPro.exe`）
6. 勾选 **Set as a pre-release**（可选），点击 **Publish release**

## 命令行上传（如果你有 GitHub PAT）

```bash
# 先创建 release
gh release create v8.0.0 "文件搜索工具(标题+正文).exe" \
  --repo BryanKo-CHN/file-search-pro \
  --title "v8 — 老式 .doc 精确提取版" \
  --notes "MD5: c4e050bdbf1cb04c5b0c546f57807755"
```

需要本机安装 GitHub CLI 并已 `gh auth login`。
