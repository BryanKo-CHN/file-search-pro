# v7 exe 发布步骤

GitHub 仓库单文件上限为 100 MB，v7 的 exe 是 116 MiB，因此 exe 不能作为普通文件提交，必须通过 **Release 附件**上传（Release 附件上限 2 GB）。

## 文件信息

- 主程序：`dist/文件搜索工具(标题+正文).exe`
- 大小：121,558,469 字节（约 116 MiB）
- MD5：`51879b780268d828108182d79c17f6ea`
- 英文名副本 `FileSearchPro.exe` 内容完全相同（MD5 一致）

## 网页上传步骤

1. 打开仓库 `https://github.com/BryanKo-CHN/file-search-pro`
2. 右侧点击 **Releases** → **Create a new release**
3. Tag 填 `v7.0.0`，标题填 `v7 — Office/WPS 双引擎兼容版`
4. 说明可写：
   ```
   v7 更新：新增 Microsoft Office / WPS 自动探测，原版引擎可用时保留原排版；
   未安装或调用失败时自动回退内置渲染。新增 --port 测试参数、FS_OFFICE_COM=0 紧急开关。
   MD5: 51879b780268d828108182d79c17f6ea
   ```
5. 拖入 `文件搜索工具(标题+正文).exe`（可同时拖入 `FileSearchPro.exe`）
6. 勾选 **Set as a pre-release**（可选），点击 **Publish release**

## 命令行上传（如果你有 GitHub PAT）

```bash
# 先创建 release
gh release create v7.0.0 "文件搜索工具(标题+正文).exe" \
  --repo BryanKo-CHN/file-search-pro \
  --title "v7 — Office/WPS 双引擎兼容版" \
  --notes "MD5: 51879b780268d828108182d79c17f6ea"
```

需要本机安装 GitHub CLI 并已 `gh auth login`。
