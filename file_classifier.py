#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文件分类与目录管理工具

功能：
  1. 全盘/文件夹扫描，按扩展名分类统计（办公文档、图片、视频、种子等）
  2. 扫描与导出均为只读，绝不改动原始文件和目录
  3. 选择某个后缀，导出该格式全部文件的目录清单到指定文件夹
  4. 对指定后缀文件执行复制 / 剪切 / 粘贴（需用户显式操作并确认）
  5. 图形化文件夹浏览器，逐级预览、点选，无需手打路径

运行：
    python file_classifier.py
    python file_classifier.py --no-browser
"""

import os
import sys
import csv
import json
import time
import uuid
import shutil
import string
import threading
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs

PORT = 8791
TASKS = {}        # 扫描任务
PASTE_JOBS = {}   # 粘贴任务
CLIPBOARD = {"mode": None, "paths": [], "task_id": None, "ext": None, "source": None}
SERVER = None

# ----------------------------- 分类定义 -----------------------------

CATEGORIES = [
    ("办公文档", {"doc", "docx", "xls", "xlsx", "ppt", "pptx", "wps", "et", "dps",
                "odt", "ods", "odp", "rtf", "pdf"}),
    ("文本与数据", {"txt", "csv", "tsv", "log", "md", "json", "xml", "yaml", "yml",
                 "ini", "cfg", "conf", "sql"}),
    ("图片", {"jpg", "jpeg", "png", "gif", "bmp", "svg", "webp", "ico", "tif",
             "tiff", "heic", "raw", "psd", "ai"}),
    ("视频", {"mp4", "avi", "mkv", "mov", "wmv", "flv", "webm", "m4v", "rmvb",
             "rm", "mpeg", "mpg", "3gp"}),
    ("音频", {"mp3", "wav", "flac", "aac", "ogg", "m4a", "wma", "ape", "mid"}),
    ("压缩包与镜像", {"zip", "rar", "7z", "tar", "gz", "bz2", "xz", "iso", "cab", "arj"}),
    ("程序与脚本", {"exe", "msi", "bat", "cmd", "sh", "ps1", "app", "deb", "rpm",
                "apk", "jar", "com", "scr"}),
    ("种子与下载", {"torrent", "magnet", "nfo", "sfv"}),
    ("编程开发", {"py", "js", "ts", "jsx", "tsx", "html", "htm", "css", "scss",
              "java", "c", "cpp", "cc", "h", "hpp", "cs", "go", "rs", "php",
              "rb", "swift", "kt", "vue", "ipynb", "less", "sass"}),
    ("字体", {"ttf", "otf", "woff", "woff2", "eot", "fon"}),
    ("其他", set()),
]

EXT_CATEGORY = {}
for _cat, _exts in CATEGORIES:
    for _e in _exts:
        EXT_CATEGORY[_e] = _cat


def category_of(ext):
    return EXT_CATEGORY.get((ext or "").lower().lstrip("."), "其他")


# ----------------------------- 工具函数 -----------------------------

def human_size(num):
    try:
        n = float(num)
    except Exception:
        return "0 B"
    for unit in ["B", "KB", "MB", "GB", "TB", "PB"]:
        if n < 1024.0 or unit == "PB":
            return f"{int(n)} B" if unit == "B" else f"{n:.2f} {unit}"
        n /= 1024.0
    return f"{n:.2f} PB"


def is_hidden(path):
    name = os.path.basename(path)
    if os.name == "nt":
        try:
            import ctypes
            attrs = ctypes.windll.kernel32.GetFileAttributesW(str(path))
            if attrs != -1 and (attrs & 0x2):
                return True
        except Exception:
            pass
    return name.startswith(".")


def to_long_path(path):
    r"""Windows 下给绝对路径加 \\?\ 前缀，突破 MAX_PATH(260) 限制。

    大容量数据盘里的中文深层目录很容易超过 260 字符，不加前缀会被系统拒绝
    访问，导致 os.walk 静默跳过整个子树——这正是"全盘扫描数量不符"的主因。
    """
    if os.name != "nt":
        return path
    p = os.path.abspath(path)
    if p.startswith("\\\\?\\"):
        return p
    if p.startswith("\\\\"):                    # UNC 网络路径
        return "\\\\?\\UNC\\" + p[2:]
    return "\\\\?\\" + p


def from_long_path(path):
    r"""去掉 \\?\ 前缀，还原成普通路径用于展示、导出与文件操作。"""
    if path.startswith("\\\\?\\UNC\\"):
        return "\\\\" + path[8:]
    if path.startswith("\\\\?\\"):
        return path[4:]
    return path


def list_drives():
    out = []
    if os.name == "nt":
        for letter in string.ascii_uppercase:
            p = letter + ":\\"
            if os.path.exists(p):
                try:
                    u = shutil.disk_usage(p)
                    out.append({"path": p, "total": u.total, "free": u.free})
                except Exception:
                    out.append({"path": p, "total": 0, "free": 0})
    else:
        out.append({"path": "/", "total": 0, "free": 0})
    return out


def safe_name(s):
    return "".join(c if (c.isalnum() or c in "-_") else "_" for c in str(s))


def make_record(full, root_path):
    # full / root_path 可能带 \\?\ 长路径前缀；stat 用长路径，展示与操作用普通路径
    full_stat = full
    full = from_long_path(full)
    root_path = from_long_path(root_path)
    name = os.path.basename(full) or full
    try:
        rel = os.path.relpath(full, root_path)
    except Exception:
        rel = full
    if rel == ".":
        rel = name
    size = 0
    mtime = ""
    try:
        st = os.stat(full_stat)
        size = st.st_size
        mtime = datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        try:
            st = os.stat(full)
            size = st.st_size
            mtime = datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            pass
    ext = os.path.splitext(name)[1].lstrip(".").lower()
    return {
        "name": name,
        "path": full,
        "rel": rel,
        "size": size,
        "mtime": mtime,
        "ext": ext,
        "category": category_of(ext),
    }


# ----------------------------- 文件夹浏览 -----------------------------

def browse_directory(path):
    """返回某路径下的子文件夹与文件，供前端图形化点选。"""
    if not path:
        return {
            "path": "", "parent": None, "is_root": True,
            "dirs": [{"name": d["path"], "path": d["path"], "hidden": False,
                      "drive": True} for d in list_drives()],
            "files": [], "files_total": 0,
        }
    path = os.path.abspath(path)
    if not os.path.isdir(path):
        return {"error": "文件夹不存在或无法访问：" + path}
    dirs, files = [], []
    # 浏览也用长路径前缀，保证能进入超深中文目录
    scan_path = to_long_path(path)
    try:
        for name in os.listdir(scan_path):
            full_long = os.path.join(scan_path, name)
            full = from_long_path(full_long)
            try:
                if os.path.isdir(full_long):
                    dirs.append({"name": name, "path": full, "hidden": is_hidden(full_long),
                                 "drive": False})
                elif os.path.isfile(full_long):
                    st = os.stat(full_long)
                    files.append({
                        "name": name, "path": full, "size": st.st_size,
                        "ext": os.path.splitext(name)[1].lstrip(".").lower(),
                        "mtime": datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
                    })
            except Exception:
                continue
    except Exception as e:
        return {"error": "读取文件夹失败：" + str(e)}
    dirs.sort(key=lambda x: x["name"].lower())
    files.sort(key=lambda x: x["name"].lower())
    parent = os.path.dirname(path)
    if parent == path:   # 盘符根目录
        parent = ""
    return {
        "path": path, "parent": parent, "is_root": False,
        "dirs": dirs, "files": files[:500], "files_total": len(files),
    }


# ----------------------------- 扫描任务 -----------------------------

class ScanTask:
    def __init__(self, path, opts):
        self.id = uuid.uuid4().hex[:12]
        self.path = os.path.abspath(path)
        self.opts = opts
        self.status = "running"
        self.status_text = "准备扫描..."
        self.count_dir = 0
        self.count_file = 0
        self.total_size = 0
        self.files = []
        self.by_ext = {}   # ext -> {count, size}
        self.errors = []
        self.skipped = []        # 无法访问的目录（路径 + 原因）
        self.followed_links = 0  # 已跟随的符号链接/联接数量
        self.current_level = 0   # 当前正在扫描的层级（逐层 BFS）
        self.max_level = 0       # 已到达的最深层数
        self.level_dirs = {}     # 每层目录数
        self.level_files = {}    # 每层新增文件数
        self.start = time.time()
        self.end = None
        self.cancel = False

    def snapshot(self):
        exts = []
        for ext, info in self.by_ext.items():
            exts.append({
                "ext": ext or "(无扩展名)",
                "key": ext,
                "category": category_of(ext),
                "count": info["count"],
                "size": info["size"],
            })
        exts.sort(key=lambda x: (-x["count"], x["ext"]))
        return {
            "id": self.id,
            "path": self.path,
            "status": self.status,
            "status_text": self.status_text,
            "dirs": self.count_dir,
            "files": self.count_file,
            "total_size": self.total_size,
            "errors": len(self.errors),
            "skipped": len(self.skipped),
            "skipped_samples": self.skipped[:20],
            "followed_links": self.followed_links,
            "current_level": self.current_level,
            "max_level": self.max_level,
            "level_dirs": dict(self.level_dirs),
            "level_files": dict(self.level_files),
            "elapsed": round((self.end or time.time()) - self.start, 2),
            "extensions": exts,
        }


def scan_worker(task):
    r"""严格按"一层一层"的广度优先（BFS）方式扫描：

        根目录(第0层) → 第1层文件夹 → 第2层 → 第3层 → …… → 最内层

    与 os.walk 的深度优先不同，BFS 保证每一层的所有文件夹都先被完整登记，
    再进入下一层，因此"二十年精选收藏"这类一级目录下的文件绝不会被深层
    遍历的顺序或中断所掩盖。

    同时根治历史漏扫点：
      1) 超长中文路径 —— \\?\ 前缀突破 MAX_PATH(260)
      2) 隐藏文件 —— 数据盘里 90% 的种子都带隐藏属性，默认强制全包含
      3) 目录联接/符号链接 —— 跟随并做 realpath 去重，避免死循环
      4) 去重误伤 —— 只对"重解析点"去重，普通目录一律不跳过
      5) 无法访问的目录 —— 记录到 skipped 并在界面明示
    """
    from collections import deque

    max_depth = int(task.opts.get("max_depth", 0) or 0)
    # 只有用户明确勾选"排除隐藏项"时才过滤；默认强制全包含
    exclude_hidden = bool(task.opts.get("exclude_hidden", False))
    follow_links = bool(task.opts.get("follow_links", True))

    scan_root = to_long_path(task.path)

    def _on_error(path, reason):
        try:
            p = from_long_path(str(path))
        except Exception:
            p = str(path)
        rec = {"path": p, "reason": str(reason)}
        if rec not in task.skipped:
            task.skipped.append(rec)
        task.errors.append(f"{p} :: {reason}")

    def _is_reparse(path):
        """判断是否为重解析点（junction / 符号链接）。"""
        try:
            import ctypes
            attrs = ctypes.windll.kernel32.GetFileAttributesW(str(path))
            if attrs != -1 and (attrs & 0x400):   # FILE_ATTRIBUTE_REPARSE_POINT
                return True
        except Exception:
            pass
        try:
            return os.path.islink(path)
        except Exception:
            return False

    # 队列元素：(目录长路径, 层数)。逐层推进，天然保证"一层一层"扫描。
    queue = deque([(scan_root, 0)])
    visited_real = set()
    try:
        visited_real.add(os.path.realpath(scan_root).lower())
    except Exception:
        pass

    current_level = 0
    try:
        while queue:
            if task.cancel:
                task.status = "cancelled"
                task.end = time.time()
                task.status_text = "已取消"
                return

            root, level = queue.popleft()
            current_level = level
            task.current_level = level
            if level > task.max_level:
                task.max_level = level

            # 读取当前目录。整目录快照：先取到全部条目，再分层处理。
            try:
                with os.scandir(root) as it:
                    entries = list(it)
            except Exception as e:
                _on_error(root, e)
                continue

            subdirs = []
            for entry in entries:
                try:
                    is_dir = entry.is_dir(follow_symlinks=follow_links)
                except Exception:
                    is_dir = False
                if is_dir:
                    # 隐藏目录过滤（默认关闭）
                    if exclude_hidden and is_hidden(entry.path):
                        continue
                    subdirs.append(entry)
                    continue
                try:
                    if not entry.is_file(follow_symlinks=follow_links):
                        continue
                except Exception:
                    continue
                if exclude_hidden and is_hidden(entry.path):
                    continue
                # 记录文件
                try:
                    rec = make_record(entry.path, scan_root)
                except Exception as e:
                    _on_error(entry.path, e)
                    continue
                task.files.append(rec)
                task.count_file += 1
                task.total_size += rec["size"]
                key = rec["ext"]
                info = task.by_ext.setdefault(key, {"count": 0, "size": 0})
                info["count"] += 1
                info["size"] += rec["size"]
                task.level_files[level] = task.level_files.get(level, 0) + 1

            task.count_dir += len(subdirs)
            task.level_dirs[level] = task.level_dirs.get(level, 0) + len(subdirs)

            # 深度限制：到达上限后不再把子目录加入队列
            if max_depth and level >= max_depth:
                pass
            else:
                for entry in subdirs:
                    child = entry.path
                    # 只对重解析点做去重；普通目录绝不跳过，避免误伤真实文件夹
                    if _is_reparse(child):
                        try:
                            child_real = os.path.realpath(child).lower()
                        except Exception:
                            child_real = child.lower()
                        if child_real in visited_real:
                            continue
                        visited_real.add(child_real)
                        task.followed_links += 1
                    queue.append((child, level + 1))

            task.status_text = (
                f"逐层扫描中 · 第 {level} 层 · 已发现 {task.count_file} 个文件 / "
                f"{task.count_dir} 个文件夹"
                + (f"（跳过 {len(task.skipped)} 处）" if task.skipped else "")
            )
    except Exception as e:
        task.errors.append(str(e))

    task.status = "done"
    task.end = time.time()
    extra = [f"共扫描 {task.max_level + 1} 层"]
    if task.followed_links:
        extra.append(f"跟随 {task.followed_links} 个链接")
    if task.skipped:
        extra.append(f"跳过 {len(task.skipped)} 处不可访问")
    suffix = ("（" + "，".join(extra) + "）") if extra else ""
    task.status_text = (
        f"扫描完成：{task.count_file} 个文件，涉及 {len(task.by_ext)} 种后缀{suffix}"
    )


def get_task_files(task, ext):
    """ext 为 '__all__' 时返回全部文件。"""
    if ext == "__all__":
        return list(task.files)
    return [f for f in task.files if f["ext"] == ext]


# ----------------------------- 导出功能（只读） -----------------------------

def export_rows(rows, meta, fmt, out_path):
    """把一组文件记录导出为目录清单。仅读取元数据，不触碰源文件。"""
    fmt = fmt.lower().lstrip(".")
    rows = sorted(rows, key=lambda r: r["path"].lower())
    label = meta["ext"]
    if fmt == "csv":
        with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["序号", "后缀名", "分类", "文件名", "相对路径", "完整路径",
                        "大小(字节)", "大小", "修改时间"])
            for i, r in enumerate(rows, 1):
                w.writerow([i, r["ext"], r["category"], r["name"], r["rel"], r["path"],
                            r["size"], human_size(r["size"]), r["mtime"]])
        return len(rows)
    if fmt == "txt":
        with open(out_path, "w", encoding="utf-8") as f:
            f.write("=" * 76 + "\n")
            f.write("文件目录清单\n")
            f.write("=" * 76 + "\n")
            f.write(f"扫描范围  : {meta['scope']}\n")
            f.write(f"筛选后缀  : {label}\n")
            f.write(f"生成时间  : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"文件数量  : {len(rows)}\n")
            f.write(f"合计大小  : {human_size(sum(r['size'] for r in rows))}\n")
            f.write("-" * 76 + "\n")
            f.write(f"{'序号':<8}{'后缀':<10}{'大小':<13}{'修改时间':<21}{'相对路径'}\n")
            f.write("-" * 76 + "\n")
            for i, r in enumerate(rows, 1):
                f.write(f"{i:<8}{r['ext']:<10}{human_size(r['size']):<13}"
                        f"{r['mtime']:<21}{r['rel']}\n")
            f.write("-" * 76 + "\n")
        return len(rows)
    if fmt in ("html", "htm"):
        import html as _h
        def esc(s):
            return _h.escape(str(s))
        trs = []
        for i, r in enumerate(rows, 1):
            trs.append(
                "<tr><td class='num'>{i}</td><td><span class='ext'>{ext}</span></td>"
                "<td>{cat}</td><td class='rel'>{rel}</td><td class='size'>{size}</td>"
                "<td class='time'>{mtime}</td></tr>".format(
                    i=i, ext=esc(r["ext"] or "-"), cat=esc(r["category"]),
                    rel=esc(r["rel"]), size=human_size(r["size"]), mtime=esc(r["mtime"])
                )
            )
        doc = """<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8"><title>文件目录清单</title>
<style>
 body{font-family:"Microsoft YaHei",Arial,sans-serif;margin:0;background:#f5f7fa;color:#1f2937;}
 .wrap{max-width:1180px;margin:24px auto;background:#fff;border-radius:12px;
  box-shadow:0 4px 20px rgba(0,0,0,.06);padding:28px;}
 h1{margin:0 0 6px;font-size:23px;color:#0f172a;}
 .meta{color:#64748b;font-size:13px;line-height:1.9;margin-bottom:20px;}
 .cards{display:flex;gap:14px;flex-wrap:wrap;margin-bottom:20px;}
 .card{flex:1;min-width:150px;background:#f8fafc;border:1px solid #e2e8f0;
  border-radius:10px;padding:15px;}
 .card .k{font-size:12px;color:#64748b;}
 .card .v{font-size:21px;font-weight:700;color:#0f172a;margin-top:5px;}
 table{width:100%;border-collapse:collapse;font-size:12px;}
 th,td{padding:8px 10px;text-align:left;border-bottom:1px solid #eef2f7;}
 th{background:#f8fafc;color:#475569;position:sticky;top:0;}
 tr:hover td{background:#fafbfc;}
 .num{color:#94a3b8;width:56px;}
 .ext{background:#e0f2fe;color:#0369a1;padding:2px 8px;border-radius:12px;font-size:11px;}
 .rel{word-break:break-all;max-width:520px;}
 .size{text-align:right;white-space:nowrap;color:#475569;}
 .time{white-space:nowrap;color:#64748b;}
</style></head><body><div class="wrap">
<h1>文件目录清单</h1>
<div class="meta">扫描范围：__SCOPE__<br>筛选后缀：__EXT__<br>生成时间：__TIME__</div>
<div class="cards">
 <div class="card"><div class="k">文件数量</div><div class="v">__COUNT__</div></div>
 <div class="card"><div class="k">合计大小</div><div class="v">__TOTAL__</div></div>
 <div class="card"><div class="k">后缀筛选</div><div class="v">__EXT__</div></div>
</div>
<table><thead><tr><th>序号</th><th>后缀</th><th>分类</th><th>相对路径</th><th>大小</th><th>修改时间</th></tr></thead>
<tbody>__ROWS__</tbody></table>
</div></body></html>"""
        doc = (doc.replace("__SCOPE__", esc(meta["scope"]))
                  .replace("__EXT__", esc(label))
                  .replace("__TIME__", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                  .replace("__COUNT__", str(len(rows)))
                  .replace("__TOTAL__", human_size(sum(r["size"] for r in rows)))
                  .replace("__ROWS__", "\n".join(trs)))
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(doc)
        return len(rows)
    raise ValueError("不支持的格式：" + fmt)


# ----------------------------- 复制 / 剪切 / 粘贴 -----------------------------

class PasteJob:
    def __init__(self, dest, collision, files, mode):
        self.id = uuid.uuid4().hex[:12]
        self.dest = os.path.abspath(dest)
        self.collision = collision
        self.files = files          # [(src, basename), ...]
        self.mode = mode            # copy / cut
        self.status = "running"
        self.total = len(files)
        self.done = 0
        self.copied = 0
        self.moved = 0
        self.skipped = 0
        self.errors = []
        self.start = time.time()
        self.cancel = False

    def snapshot(self):
        return {
            "id": self.id, "status": self.status, "mode": self.mode,
            "total": self.total, "done": self.done,
            "copied": self.copied, "moved": self.moved, "skipped": self.skipped,
            "errors": self.errors[:20], "error_count": len(self.errors),
            "elapsed": round(time.time() - self.start, 2),
        }


def unique_path(path):
    base, ext = os.path.splitext(path)
    i = 1
    while True:
        cand = f"{base} ({i}){ext}"
        # 用长路径检查，避免超长路径下误判为不存在而覆盖
        if not os.path.exists(to_long_path(cand)):
            return cand
        i += 1


def paste_worker(job):
    try:
        os.makedirs(to_long_path(job.dest), exist_ok=True)
    except Exception as e:
        job.errors.append("无法创建目标文件夹：" + str(e))
        job.status = "error"
        return
    for src, name in job.files:
        if job.cancel:
            break
        dst = os.path.join(job.dest, name)
        # 文件系统操作统一走长路径，展示/记录仍用普通路径
        src_op = to_long_path(src)
        dst_op = to_long_path(dst)
        try:
            if src_op.lower() == dst_op.lower():
                job.skipped += 1     # 源和目标相同，跳过
                continue
            if not os.path.isfile(src_op):
                job.errors.append("源文件已不存在：" + src)
                continue
            if os.path.exists(dst_op):
                if job.collision == "rename":
                    dst_op = to_long_path(unique_path(dst))
                    dst = from_long_path(dst_op)
                else:
                    job.skipped += 1
                    continue
            if job.mode == "copy":
                shutil.copy2(src_op, dst_op)
                job.copied += 1
            else:
                shutil.move(src_op, dst_op)
                job.moved += 1
        except Exception as e:
            job.errors.append(f"{name}: {e}")
        job.done += 1
    job.status = "cancelled" if job.cancel else "done"
    # 剪切完成后清空剪贴板，避免误操作导致重复移动
    if job.status == "done" and job.mode == "cut" and not job.cancel:
        CLIPBOARD.update({"mode": None, "paths": [], "task_id": None,
                          "ext": None, "source": None})


def resolve_clip_files():
    """把剪贴板内容解析为 [(源文件, 文件名), ...]。"""
    result = []
    if CLIPBOARD.get("task_id") and CLIPBOARD.get("ext"):
        task = TASKS.get(CLIPBOARD["task_id"])
        if task:
            ext = CLIPBOARD["ext"]
            for f in (get_task_files(task, ext) if ext != "__all__" else task.files):
                result.append((f["path"], f["name"]))
    elif CLIPBOARD.get("paths"):
        for p in CLIPBOARD["paths"]:
            if os.path.isfile(to_long_path(p)):
                result.append((p, os.path.basename(p)))
    return result


# ----------------------------- Web 界面 -----------------------------

HTML_PAGE = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>文件分类与目录管理工具</title>
<style>
:root{--bg:#f5f7fa;--panel:#fff;--line:#e2e8f0;--text:#1f2937;--muted:#64748b;
 --primary:#2563eb;--primary-dark:#1d4ed8;--ok:#16a34a;--warn:#d97706;--soft:#f8fafc;}
*{box-sizing:border-box;}
body{margin:0;background:var(--bg);color:var(--text);font-size:14px;
 font-family:"Microsoft YaHei","Segoe UI",Arial,sans-serif;}
.header{background:linear-gradient(120deg,#1e40af,#2563eb);color:#fff;padding:20px 0;}
.header .inner{max-width:1180px;margin:0 auto;padding:0 24px;}
.header h1{margin:0;font-size:22px;}
.header p{margin:6px 0 0;opacity:.85;font-size:13px;}
.main{max-width:1180px;margin:20px auto;padding:0 24px 60px;}
.panel{background:var(--panel);border:1px solid var(--line);border-radius:12px;
 padding:20px;margin-bottom:16px;box-shadow:0 2px 10px rgba(15,23,42,.04);}
.panel h2{font-size:15px;margin:0 0 16px;color:#0f172a;
 border-left:3px solid var(--primary);padding-left:10px;}
label{font-size:13px;color:var(--muted);display:block;margin-bottom:6px;}
input[type=text]{padding:10px 12px;border:1px solid var(--line);border-radius:8px;
 font-size:14px;outline:none;background:#fff;color:var(--text);}
input[type=text]:focus{border-color:var(--primary);box-shadow:0 0 0 3px rgba(37,99,235,.12);}
.pick-input{flex:1;min-width:240px;background:#f8fafc;}
button{border:none;border-radius:8px;padding:10px 16px;font-size:14px;cursor:pointer;
 font-family:inherit;transition:.15s;white-space:nowrap;}
.btn-primary{background:var(--primary);color:#fff;}
.btn-primary:hover{background:var(--primary-dark);}
.btn-light{background:#eef2ff;color:#3730a3;}
.btn-light:hover{background:#e0e7ff;}
.btn-ghost{background:#f1f5f9;color:#475569;}
.btn-ghost:hover{background:#e2e8f0;}
.btn-danger{background:#fee2e2;color:#b91c1c;}
.btn-danger:hover{background:#fecaca;}
button:disabled{opacity:.45;cursor:not-allowed;}
.input-row{display:flex;gap:9px;align-items:center;flex-wrap:wrap;}
.quick{margin-top:12px;display:flex;gap:8px;flex-wrap:wrap;align-items:center;}
.quick span{font-size:12px;color:var(--muted);}
.chip{background:#f1f5f9;border:1px solid var(--line);color:#475569;
 padding:5px 11px;border-radius:20px;font-size:12px;}
.chip:hover{background:#e0ecff;color:#1d4ed8;border-color:#bfdbfe;}
.opts{display:flex;gap:26px;flex-wrap:wrap;align-items:center;margin-top:14px;}
.check{display:flex;align-items:center;gap:8px;cursor:pointer;color:var(--text);font-size:13px;}
.check input{width:16px;height:16px;accent-color:var(--primary);}
.opt input[type=number]{width:100px;padding:8px 10px;border:1px solid var(--line);
 border-radius:8px;outline:none;background:#fff;color:var(--text);}
.progress-wrap{margin-top:16px;display:none;}
.bar{height:8px;background:#e2e8f0;border-radius:20px;overflow:hidden;}
.bar>i{display:block;height:100%;width:8%;border-radius:20px;
 background:linear-gradient(90deg,#60a5fa,#2563eb);animation:move 1.2s ease-in-out infinite alternate;}
@keyframes move{from{width:6%;}to{width:92%;}}
.done .bar>i{animation:none;width:100%;background:var(--ok);}
.status{margin-top:8px;font-size:13px;color:var(--muted);}
.stats{display:none;grid-template-columns:repeat(4,1fr);gap:12px;margin-top:16px;}
.stat{background:var(--soft);border:1px solid var(--line);border-radius:10px;padding:13px;}
.stat .k{font-size:12px;color:var(--muted);}
.stat .v{font-size:20px;font-weight:700;margin-top:4px;color:#0f172a;}
.cat-group{margin-top:16px;display:none;}
.cat-title{font-size:13px;font-weight:700;color:#334155;margin-bottom:9px;
 display:flex;align-items:center;gap:8px;}
.cat-title:before{content:"";width:6px;height:6px;border-radius:50%;background:var(--primary);}
.ext-list{display:flex;gap:8px;flex-wrap:wrap;}
.ext-item{border:1px solid var(--line);background:#fff;border-radius:9px;
 padding:8px 12px;cursor:pointer;min-width:118px;transition:.15s;}
.ext-item:hover{border-color:#93c5fd;background:#f8fbff;}
.ext-item.active{border-color:var(--primary);background:#eff6ff;
 box-shadow:0 0 0 2px rgba(37,99,235,.12);}
.ext-item .e{font-weight:700;color:#0f172a;font-size:13px;}
.ext-item .c{font-size:11px;color:var(--muted);margin-top:3px;}
.curbar{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-bottom:12px;}
.curbar b{color:var(--primary);font-size:16px;}
.toolbar{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px;}
.filelist{border:1px solid var(--line);border-radius:8px;max-height:360px;overflow:auto;display:none;}
.filelist table{font-size:12px;}
.filelist td.path{color:#475569;max-width:520px;overflow:hidden;text-overflow:ellipsis;
 white-space:nowrap;}
.filelist .num{width:44px;}
.pager{display:flex;gap:8px;align-items:center;margin-top:10px;font-size:12px;color:var(--muted);}
.msg{margin-top:12px;padding:10px 14px;border-radius:8px;display:none;font-size:13px;}
.msg.ok{background:#dcfce7;color:#166534;display:block;}
.msg.err{background:#fee2e2;color:#991b1b;display:block;}
.msg.info{background:#e0f2fe;color:#075985;display:block;}
select{padding:10px 12px;border:1px solid var(--line);border-radius:8px;
 background:#fff;color:var(--text);font-size:14px;outline:none;font-family:inherit;}
.clipbox{background:var(--soft);border:1px dashed #cbd5e1;border-radius:9px;
 padding:13px;margin-bottom:12px;font-size:13px;color:#475569;}
.clipbox b{color:#0f172a;}
/* 文件夹浏览弹窗 */
.mask{position:fixed;inset:0;background:rgba(15,23,42,.45);display:none;
 align-items:center;justify-content:center;z-index:50;}
.mask.show{display:flex;}
.modal{background:#fff;border-radius:14px;width:min(760px,94vw);max-height:86vh;
 display:flex;flex-direction:column;box-shadow:0 20px 60px rgba(0,0,0,.25);}
.modal-head{padding:16px 20px;border-bottom:1px solid var(--line);
 display:flex;justify-content:space-between;align-items:center;}
.modal-head h3{margin:0;font-size:16px;}
.modal-head button{padding:4px 10px;font-size:18px;background:#f1f5f9;color:#475569;}
.modal-bar{padding:12px 20px;border-bottom:1px solid var(--line);display:flex;gap:8px;}
.modal-bar input{flex:1;min-width:0;font-size:12px;}
.modal-list{padding:8px 12px;overflow:auto;flex:1;min-height:280px;}
.brow-row{display:flex;align-items:center;gap:10px;padding:8px 10px;border-radius:7px;
 cursor:pointer;font-size:13px;}
.brow-row:hover{background:#f1f5f9;}
.brow-row .ic{width:22px;color:#2563eb;font-weight:700;}
.brow-row.folder .ic{color:#d97706;}
.brow-row .nm{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.brow-row .sz{color:#94a3b8;font-size:11px;}
.brow-tip{padding:10px;color:#94a3b8;font-size:12px;text-align:center;}
.modal-foot{padding:14px 20px;border-top:1px solid var(--line);
 display:flex;justify-content:space-between;align-items:center;gap:10px;}
.modal-foot .hint{font-size:12px;color:var(--muted);}
@media(max-width:820px){.stats{grid-template-columns:repeat(2,1fr);}}
</style>
</head>
<body>
<div class="header"><div class="inner">
 <h1>文件分类与目录管理工具</h1>
 <p>按后缀名分类统计 · 只读导出指定格式目录 · 显式复制/剪切/粘贴 · 可视化选择文件夹</p>
</div></div>

<div class="main">
  <div class="panel">
    <h2>1. 选择扫描范围</h2>
    <div class="input-row">
      <input type="text" id="scanPath" class="pick-input" readonly placeholder="点击“浏览选择”选取硬盘或文件夹">
      <button class="btn-light" id="btnBrowseScan">浏览选择…</button>
      <button class="btn-primary" id="btnScan">开始扫描</button>
      <button class="btn-ghost" id="btnCancel" disabled>取消</button>
    </div>
    <div class="quick" id="quick"><span>快捷盘符：</span></div>
    <div class="opts">
      <label class="check"><input type="checkbox" id="followLinks" checked> 跟随目录联接/符号链接（推荐，避免漏扫）</label>
      <label class="check"><input type="checkbox" id="excHidden"> 排除隐藏文件
        <span style="color:#b45309;">（⚠ 数据盘约九成种子带隐藏属性，勾选会大量漏扫，请勿勾选）</span></label>
      <div class="opt"><label>最大层数（0 = 一直扫到最内层）</label>
        <input type="number" id="maxDepth" value="0" min="0" max="99"></div>
    </div>
    <div class="progress-wrap" id="progressWrap">
      <div class="bar"><i></i></div>
      <div class="status" id="status">准备中...</div>
      <div class="status" id="levelInfo" style="color:#2563eb;font-weight:600;"></div>
    </div>
    <div class="stats" id="stats">
      <div class="stat"><div class="k">文件总数</div><div class="v" id="sFiles">0</div></div>
      <div class="stat"><div class="k">文件夹数</div><div class="v" id="sDirs">0</div></div>
      <div class="stat"><div class="k">文件总大小</div><div class="v" id="sSize">0 B</div></div>
      <div class="stat"><div class="k">已扫层数 / 后缀种类</div><div class="v" id="sKinds">0 / 0</div></div>
    </div>
    <div id="levelBars" style="display:none;margin-top:12px;padding:10px 12px;background:#eff6ff;
      border:1px solid #bfdbfe;border-radius:8px;font-size:12px;color:#1e40af;"></div>
    <div id="skipNote" style="display:none;margin-top:12px;padding:10px 12px;background:#fff7ed;
      border:1px solid #fed7aa;border-radius:8px;font-size:12px;color:#9a3412;"></div>
  </div>

  <div class="panel">
    <h2>2. 按后缀名分类的扫描结果</h2>
    <div id="categories"></div>
    <div class="note" id="catNote" style="font-size:12px;color:#94a3b8;margin-top:12px;"></div>
  </div>

  <div class="panel">
    <h2>3. 查看某个格式的文件</h2>
    <div class="curbar">
      <span>当前格式：</span><b id="curExt">未选择</b>
      <span id="curExtInfo" style="font-size:12px;color:#64748b;"></span>
    </div>
    <div class="toolbar">
      <button class="btn-ghost" id="btnSelAll">全选本页</button>
      <button class="btn-ghost" id="btnSelNone">取消全选本页</button>
      <button class="btn-light" id="btnCopySel">复制选中</button>
      <button class="btn-light" id="btnCutSel">剪切选中</button>
      <button class="btn-light" id="btnCopyAll">复制本格式全部</button>
      <button class="btn-danger" id="btnCutAll">剪切本格式全部</button>
    </div>
    <div class="filelist" id="fileList">
      <table><thead><tr><th></th><th>文件名</th><th>相对路径</th><th>大小</th><th>修改时间</th></tr></thead>
      <tbody id="fileBody"></tbody></table>
    </div>
    <div class="pager" id="pager"></div>

    <h2 style="margin-top:24px;">导出该格式目录清单（只读，不改动原文件）</h2>
    <div class="input-row">
      <select id="fmt">
        <option value="html">HTML 网页（推荐）</option>
        <option value="csv">CSV 表格（Excel）</option>
        <option value="txt">TXT 文本</option>
      </select>
      <input type="text" id="outDir" class="pick-input" readonly placeholder="选择输出文件夹">
      <button class="btn-light" id="btnBrowseOut">浏览…</button>
      <button class="btn-primary" id="btnExport" disabled>导出</button>
    </div>
    <div class="msg" id="msg"></div>
  </div>

  <div class="panel">
    <h2>4. 剪贴板与粘贴</h2>
    <div class="clipbox" id="clipInfo">剪贴板为空。请在上一步点击“复制”或“剪切”。</div>
    <div class="input-row">
      <input type="text" id="pasteDir" class="pick-input" readonly placeholder="选择粘贴的目标文件夹">
      <button class="btn-light" id="btnBrowsePaste">浏览…</button>
      <button class="btn-primary" id="btnPaste" disabled>粘贴到该文件夹</button>
      <button class="btn-ghost" id="btnClearClip">清空剪贴板</button>
    </div>
    <div class="opts">
      <label class="check"><input type="checkbox" id="renameOnCollision">
        文件名冲突时自动重命名（默认跳过，不覆盖已有文件）</label>
    </div>
    <div class="progress-wrap" id="pasteWrap">
      <div class="bar"><i></i></div>
      <div class="status" id="pasteStatus"></div>
    </div>
    <div class="msg" id="pasteMsg"></div>
  </div>

  <div style="text-align:center;margin-top:6px;">
    <button class="btn-ghost" id="btnQuit" style="padding:7px 16px;font-size:12px;">退出程序</button>
  </div>
</div>

<!-- 文件夹浏览弹窗 -->
<div class="mask" id="browserMask">
 <div class="modal">
  <div class="modal-head"><h3 id="browserTitle">选择文件夹</h3>
   <button id="btnCloseBrowser">×</button></div>
  <div class="modal-bar">
   <button class="btn-ghost" id="btnUp">↑ 上级</button>
   <button class="btn-ghost" id="btnRoots">盘符</button>
   <input type="text" id="browserPath" readonly>
  </div>
  <div class="modal-list" id="browserList"></div>
  <div class="modal-foot">
   <span class="hint" id="browserHint">点击文件夹进入，查看内容；选好后点右侧按钮。</span>
   <button class="btn-primary" id="btnSelectFolder">选择当前文件夹</button>
  </div>
 </div>
</div>

<script>
const $ = id => document.getElementById(id);
let taskId=null, currentExt=null, currentFiles=[], selected=new Set();
let page=0, pageSize=80, pollTimer=null, pastePoll=null;
let browserTarget=null, browserCurrent='';

function esc(s){return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;')
 .replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');}
function fmtSize(n){if(!n)return '0 B';const u=['B','KB','MB','GB','TB','PB'];let i=0,x=Number(n);
 while(x>=1024&&i<u.length-1){x/=1024;i++;}return i===0?Math.round(x)+' B':x.toFixed(2)+' '+u[i];}
function msg(id,text,kind){const m=$(id);m.textContent=text;m.className='msg '+(kind||'info');}
async function api(path,data){
 const opt=data?{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)}:{};
 const r=await fetch(path,opt);return await r.json();
}

/* ---------- 文件夹浏览器 ---------- */
function openBrowser(target,initial){
 browserTarget=target;
 browserCurrent=initial||'';
 $('browserTitle').textContent =
  ({scan:'选择扫描范围',out:'选择导出目录',paste:'选择粘贴目标'})[target]||'选择文件夹';
 $('browserMask').classList.add('show');
 browseTo(browserCurrent||'');
}
function closeBrowser(){$('browserMask').classList.remove('show');}
$('btnCloseBrowser').onclick=closeBrowser;
$('browserMask').onclick=function(e){if(e.target===this)closeBrowser();};
$('btnRoots').onclick=()=>browseTo('');
$('btnUp').onclick=()=>browseTo(__upPath(browserCurrent));
function __upPath(p){
 if(!p)return '';
 const clean=p.replace(/[\\/]+$/,'');
 const parent=clean.replace(/[\\/][^\\/]*$/,'');
 return (parent===clean)?'':parent;
}
async function browseTo(p){
 const r=await api('/api/browse?path='+encodeURIComponent(p));
 if(r.error){alert(r.error);return;}
 browserCurrent=r.path||'';
 $('browserPath').value=r.is_root?'（本机盘符）':browserCurrent;
 const box=$('browserList');box.innerHTML='';
 if(r.is_root){
  const tip=document.createElement('div');tip.className='brow-tip';
  tip.textContent='双击盘符进入，逐级浏览选择文件夹';box.appendChild(tip);
 }
 (r.dirs||[]).forEach(d=>{
  const row=document.createElement('div');row.className='brow-row folder';
  row.innerHTML='<span class="ic">['+'</span><span class="nm">'+esc(d.name)+
   '</span><span class="sz">'+(d.drive?'磁盘分区':'文件夹')+'</span>';
  row.querySelector('.ic').textContent=d.drive?'[盘]':'[+]';
  row.onclick=()=>browseTo(d.path);
  box.appendChild(row);
 });
 (r.files||[]).forEach(f=>{
  const row=document.createElement('div');row.className='brow-row';
  row.innerHTML='<span class="ic">[-]</span><span class="nm">'+esc(f.name)+
   '</span><span class="sz">'+(f.ext?'.'+esc(f.ext)+' · ':'')+fmtSize(f.size)+'</span>';
  box.appendChild(row);
 });
 if(r.files_total>(r.files||[]).length){
  const tip=document.createElement('div');tip.className='brow-tip';
  tip.textContent='仅预览前 500 个文件（共 '+r.files_total+' 个），文件夹已全部列出';
  box.appendChild(tip);
 }
 $('browserHint').textContent=r.is_root?'选择一个磁盘分区开始浏览':'当前：'+browserCurrent;
}
$('btnSelectFolder').onclick=()=>{
 if(!browserCurrent){alert('请先进入一个磁盘分区或文件夹');return;}
 if(browserTarget==='scan')$('scanPath').value=browserCurrent;
 if(browserTarget==='out')$('outDir').value=browserCurrent;
 if(browserTarget==='paste')$('pasteDir').value=browserCurrent;
 closeBrowser();refreshActionButtons();
};
$('btnBrowseScan').onclick=()=>openBrowser('scan',$('scanPath').value);
$('btnBrowseOut').onclick=()=>openBrowser('out',$('outDir').value);
$('btnBrowsePaste').onclick=()=>openBrowser('paste',$('pasteDir').value);

/* ---------- 快捷盘符 ---------- */
async function loadDrives(){
 const d=await api('/api/drives');const box=$('quick');
 (d.drives||[]).forEach(dr=>{
  const b=document.createElement('button');b.className='chip';
  b.textContent=dr.path+' ('+fmtSize(dr.free)+'可用)';
  b.onclick=()=>{$('scanPath').value=dr.path;refreshActionButtons();};
  box.appendChild(b);
 });
 [['桌面','Desktop'],['文档','Documents'],['下载','Downloads']].forEach(([n,k])=>{
  const b=document.createElement('button');b.className='chip';b.textContent=n;
  b.onclick=async()=>{const r=await api('/api/quickpath?name='+k);
   if(r.path){$('scanPath').value=r.path;refreshActionButtons();}};
  box.appendChild(b);
 });
}
function refreshActionButtons(){
 $('btnScan').disabled=!$('scanPath').value;
 $('btnExport').disabled=!taskId||!currentExt||!$('outDir').value;
 $('btnPaste').disabled=!(CLIP.hasMode&&$('pasteDir').value);
}

/* ---------- 扫描 ---------- */
$('btnScan').onclick=async()=>{
 const p=$('scanPath').value.trim();if(!p){alert('请先选择扫描范围');return;}
 $('btnScan').disabled=true;$('btnCancel').disabled=false;
 $('btnExport').disabled=true;
 $('progressWrap').style.display='block';$('progressWrap').classList.remove('done');
 $('stats').style.display='none';$('categories').innerHTML='';
 $('catNote').textContent='';$('status').textContent='正在启动...';
 $('skipNote').style.display='none';$('levelBars').style.display='none';
 $('levelInfo').textContent='准备按层扫描：根目录 → 第1层 → 第2层 → …';
 const r=await api('/api/scan',{path:p,max_depth:$('maxDepth').value,
  exclude_hidden:$('excHidden').checked, follow_links:$('followLinks').checked});
 if(r.error){alert(r.error);$('btnScan').disabled=false;
  $('btnCancel').disabled=true;return;}
 taskId=r.id;pollTimer=setInterval(poll,350);
};
$('btnCancel').onclick=async()=>{if(taskId)await api('/api/cancel',{id:taskId});};
async function poll(){
 const s=await api('/api/progress?id='+encodeURIComponent(taskId));
 if(s.error){clearInterval(pollTimer);alert(s.error);return;}
 $('status').textContent=s.status_text+'（用时 '+s.elapsed+'s）';
 $('levelInfo').textContent='正在扫描第 '+s.current_level+' 层（已到达最深第 '+s.max_level+' 层）';
 if(s.status==='done'||s.status==='cancelled'){
  clearInterval(pollTimer);$('progressWrap').classList.add('done');
  $('btnScan').disabled=false;$('btnCancel').disabled=true;
  $('stats').style.display='grid';
  $('sFiles').textContent=s.files;$('sDirs').textContent=s.dirs;
  $('sSize').textContent=fmtSize(s.total_size);
  $('sKinds').textContent=(s.max_level+1)+' / '+(s.extensions||[]).length;
  // 逐层结果明细：让用户清楚看到第 1、2、3… 层各扫到多少文件夹和文件
  const lf=s.level_files||{}, ld=s.level_dirs||{};
  const keys=Object.keys(lf).concat(Object.keys(ld))
    .filter((v,i,a)=>a.indexOf(v)===i).map(Number).sort((a,b)=>a-b);
  if(keys.length){
   $('levelBars').innerHTML='<b>逐层扫描明细</b>（根目录为第 0 层）：<br>'+
     keys.map(k=>'· 第 '+k+' 层：'+ (ld[k]||0) +' 个文件夹，'+(lf[k]||0)+' 个文件').join('<br>');
   $('levelBars').style.display='block';
  }
  renderCategories(s);
  // 显示被跳过的不可访问目录，让用户明确知道哪里没扫到
  const sn=$('skipNote');
  if(s.skipped>0){
   let lines=(s.skipped_samples||[]).map(x=>'· '+esc(x.path)+(x.reason?'（'+esc(x.reason)+'）':''));
   sn.innerHTML='<b>有 '+s.skipped+' 个目录无法访问，已跳过</b>（通常是权限不足或系统保护目录）：<br>'
     +lines.join('<br>')+(s.skipped>lines.length?'<br>… 另有 '+(s.skipped-lines.length)+' 个':'');
   sn.style.display='block';
  }else{sn.style.display='none';}
  if(s.status==='cancelled')msg('msg','扫描已取消','err');
 }
}
function renderCategories(s){
 const box=$('categories');box.innerHTML='';
 if(!s.extensions||!s.extensions.length){box.textContent='未发现文件';return;}
 const groups={};
 s.extensions.forEach(e=>{(groups[e.category]=groups[e.category]||[]).push(e);});
 Object.keys(groups).forEach(cat=>{
  const g=document.createElement('div');g.className='cat-group';g.style.display='block';
  const t=document.createElement('div');t.className='cat-title';
  t.textContent=cat+'（'+groups[cat].reduce((a,b)=>a+b.count,0)+' 个文件）';
  g.appendChild(t);
  const list=document.createElement('div');list.className='ext-list';
  groups[cat].forEach(e=>{
   const item=document.createElement('div');item.className='ext-item';
   if(e.key===currentExt)item.classList.add('active');
   item.innerHTML='<div class="e">.'+esc(e.ext)+'</div><div class="c">'+
    e.count+' 个 · '+fmtSize(e.size)+'</div>';
   item.onclick=()=>selectExt(e.key,e.ext,e.count,e.size);
   list.appendChild(item);
  });
  g.appendChild(list);box.appendChild(g);
 });
 $('catNote').textContent='共 '+s.extensions.length+' 种后缀。点击任意后缀查看具体文件并导出清单。';
}
async function selectExt(key,label,count,size){
 currentExt=key;page=0;selected.clear();
 $('curExt').textContent=(key==='__all__'?'全部格式':'.'+label);
 $('curExtInfo').textContent=count+' 个文件 · '+fmtSize(size);
 document.querySelectorAll('.ext-item').forEach(el=>el.classList.remove('active'));
 if(label!==undefined){
  document.querySelectorAll('.ext-item').forEach(el=>{
   if(el.querySelector('.e').textContent==='.'+label)el.classList.add('active');});
 }
 $('fileList').style.display='block';
 await loadFiles();
 refreshActionButtons();
}
async function loadFiles(){
 const url='/api/files?id='+encodeURIComponent(taskId)+'&ext='+encodeURIComponent(currentExt)+
  '&offset='+(page*pageSize)+'&limit='+pageSize;
 const r=await api(url);if(r.error){alert(r.error);return;}
 currentFiles=r.files||[];
 const body=$('fileBody');body.innerHTML='';
 currentFiles.forEach((f,i)=>{
  const tr=document.createElement('tr');
  tr.innerHTML='<td class="num"><input type="checkbox" class="ck" data-i="'+i+'" '+
   (selected.has(f.path)?'checked':'')+'></td><td>'+esc(f.name)+'</td>'+
   '<td class="path" title="'+esc(f.path)+'">'+esc(f.rel)+'</td><td class="size">'+
   fmtSize(f.size)+'</td><td class="time">'+esc(f.mtime)+'</td>';
  body.appendChild(tr);
 });
 body.querySelectorAll('.ck').forEach(ck=>{
  ck.onchange=()=>{const p=currentFiles[Number(ck.dataset.i)].path;
   if(ck.checked)selected.add(p);else selected.delete(p);};
 });
 const pages=Math.ceil((r.total||0)/pageSize);
 $('pager').textContent='共 '+(r.total||0)+' 个文件 · 第 '+(page+1)+'/'+
  Math.max(pages,1)+' 页';
 if(pages>1){
  const prev=document.createElement('button');prev.className='btn-ghost';
  prev.textContent='上一页';prev.disabled=page===0;
  prev.onclick=()=>{page--;loadFiles();};
  const next=document.createElement('button');next.className='btn-ghost';
  next.textContent='下一页';next.disabled=page>=pages-1;
  next.onclick=()=>{page++;loadFiles();};
  $('pager').appendChild(prev);$('pager').appendChild(next);
 }
}
$('btnSelAll').onclick=()=>{currentFiles.forEach(f=>selected.add(f.path));loadFiles();};
$('btnSelNone').onclick=()=>{currentFiles.forEach(f=>selected.delete(f.path));loadFiles();};

/* ---------- 剪贴板 ---------- */
const CLIP={hasMode:false,mode:null,count:0,ext:null};
async function setClip(mode,useSelection){
 if(!taskId){alert('请先扫描');return;}
 let body={mode:mode};
 if(useSelection){
  if(!selected.size){alert('请先勾选文件');return;}
  body.paths=Array.from(selected);
 }else{
  body.task_id=taskId;body.ext=currentExt;
  if(!currentExt){alert('请先选择一种后缀');return;}
 }
 const r=await api('/api/clip/set',body);
 if(r.error){alert(r.error);return;}
 CLIP.hasMode=true;CLIP.mode=mode;CLIP.count=r.count;CLIP.ext=r.ext;
 renderClip();refreshActionButtons();
 msg('msg','已将 '+r.count+' 个文件放入剪贴板（'+(mode==='copy'?'复制':'剪切')+
  '）。请到第 4 步选择目标文件夹并粘贴。','ok');
}
function renderClip(){
 if(!CLIP.hasMode){$('clipInfo').innerHTML='剪贴板为空。请在上一步点击“复制”或“剪切”。';return;}
 $('clipInfo').innerHTML='剪贴板：<b>'+(CLIP.mode==='copy'?'复制':'剪切')+'</b> · <b>'+
  CLIP.count+'</b> 个文件 · 格式 '+(CLIP.ext==='__all__'?'全部':'.'+esc(CLIP.ext));
}
$('btnCopySel').onclick=()=>setClip('copy',true);
$('btnCutSel').onclick=()=>setClip('cut',true);
$('btnCopyAll').onclick=()=>setClip('copy',false);
$('btnCutAll').onclick=()=>{
 if(!confirm('确定要“剪切”该格式的全部文件吗？\n粘贴成功后，源文件会从原位置移动走。'))return;
 setClip('cut',false);
};
$('btnClearClip').onclick=async()=>{
 await api('/api/clip/clear',{});
 CLIP.hasMode=false;CLIP.mode=null;CLIP.count=0;
 renderClip();refreshActionButtons();msg('pasteMsg','剪贴板已清空','info');
};

/* ---------- 粘贴 ---------- */
$('btnPaste').onclick=async()=>{
 const dest=$('pasteDir').value.trim();if(!dest){alert('请先选择粘贴目标文件夹');return;}
 if(CLIP.mode==='cut'){
  if(!confirm('即将移动 '+CLIP.count+' 个文件到：\n'+dest+
   '\n\n源文件会从原位置移走，确定继续吗？'))return;
 }
 const r=await api('/api/paste',{dest:dest,collision:$('renameOnCollision').checked?'rename':'skip'});
 if(r.error){alert(r.error);return;}
 $('pasteWrap').style.display='block';$('pasteWrap').classList.remove('done');
 $('pasteStatus').textContent='正在准备粘贴...';
 pastePoll=setInterval(async()=>{
  const s=await api('/api/paste/status?id='+encodeURIComponent(r.id));
  if(s.error){clearInterval(pastePoll);return;}
  $('pasteStatus').textContent='已处理 '+s.done+'/'+s.total+' 个文件'+
   '（复制 '+s.copied+' · 移动 '+s.moved+' · 跳过 '+s.skipped+'）';
  if(s.status==='done'||s.status==='cancelled'||s.status==='error'){
   clearInterval(pastePoll);$('pasteWrap').classList.add('done');
   msg('pasteMsg','粘贴完成：复制 '+s.copied+' 个，移动 '+s.moved+
    ' 个，跳过 '+s.skipped+' 个'+(s.error_count?'，失败 '+s.error_count+' 个':'')+
    (s.mode==='cut'?'。剪切后剪贴板已清空，原扫描结果已过期，请重新扫描。':''),'ok');
   if(s.mode==='cut'){CLIP.hasMode=false;renderClip();refreshActionButtons();}
  }
 },300);
};

/* ---------- 导出 ---------- */
$('btnExport').onclick=async()=>{
 if(!taskId||!currentExt)return;
 const out=$('outDir').value.trim();if(!out){alert('请选择输出文件夹');return;}
 $('btnExport').disabled=true;msg('msg','正在生成目录清单...','info');
 const r=await api('/api/export',{id:taskId,ext:currentExt,fmt:$('fmt').value,
  output_dir:out,include_dirs:false});
 $('btnExport').disabled=false;
 if(r.error){msg('msg',r.error,'err');return;}
 msg('msg','导出成功！'+r.rows+' 条记录 → '+r.path,'ok');
};

/* ---------- 退出 ---------- */
$('btnQuit').onclick=async()=>{
 if(!confirm('确定要退出文件分类工具吗？'))return;
 await api('/api/quit',{});
 document.body.innerHTML='<div style="text-align:center;padding:80px 20px;'+
  'font-family:Microsoft YaHei,sans-serif;color:#64748b;">'+
  '<h2 style="color:#0f172a;">程序已退出</h2><p>可以关闭此浏览器标签页。</p></div>';
};

loadDrives();renderClip();refreshActionButtons();
</script>
</body>
</html>"""


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        pass

    def _send(self, code, body, ctype="application/json; charset=utf-8"):
        if isinstance(body, (dict, list)):
            body = json.dumps(body, ensure_ascii=False).encode("utf-8")
        elif isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        try:
            self.wfile.write(body)
        except Exception:
            pass

    def _json_body(self):
        try:
            length = int(self.headers.get("Content-Length", 0))
            if not length:
                return {}
            return json.loads(self.rfile.read(length).decode("utf-8"))
        except Exception:
            return {}

    def do_GET(self):
        u = urlparse(self.path)
        if u.path == "/":
            self._send(200, HTML_PAGE, "text/html; charset=utf-8")
        elif u.path == "/api/drives":
            self._send(200, {"drives": list_drives()})
        elif u.path == "/api/quickpath":
            name = parse_qs(u.query).get("name", [""])[0]
            home = os.path.expanduser("~")
            mapping = {
                "Desktop": ["Desktop", "OneDrive/Desktop", "桌面"],
                "Documents": ["Documents", "My Documents", "文档"],
                "Downloads": ["Downloads", "下载"],
            }
            path = None
            for cand in mapping.get(name, []):
                p = os.path.join(home, cand)
                if os.path.isdir(p):
                    path = p
                    break
            self._send(200, {"path": path})
        elif u.path == "/api/browse":
            path = parse_qs(u.query).get("path", [""])[0]
            self._send(200, browse_directory(path))
        elif u.path == "/api/progress":
            task = TASKS.get(parse_qs(u.query).get("id", [""])[0])
            self._send(200, task.snapshot() if task else {"error": "任务不存在，请重新扫描"})
        elif u.path == "/api/files":
            q = parse_qs(u.query)
            task = TASKS.get(q.get("id", [""])[0])
            if not task:
                self._send(200, {"error": "任务不存在"})
                return
            ext = q.get("ext", [""])[0]
            try:
                offset = max(0, int(q.get("offset", ["0"])[0]))
                limit = min(500, max(1, int(q.get("limit", ["100"])[0])))
            except Exception:
                offset, limit = 0, 100
            files = get_task_files(task, ext)
            self._send(200, {"total": len(files), "files": files[offset:offset + limit]})
        elif u.path == "/api/paste/status":
            job = PASTE_JOBS.get(parse_qs(u.query).get("id", [""])[0])
            self._send(200, job.snapshot() if job else {"error": "粘贴任务不存在"})
        else:
            self._send(404, {"error": "not found"})

    def do_POST(self):
        u = urlparse(self.path)
        data = self._json_body()

        if u.path == "/api/scan":
            path = (data.get("path") or "").strip().strip('"')
            if not os.path.isdir(path):
                self._send(200, {"error": "路径不存在或不是文件夹：" + path})
                return
            task = ScanTask(path, data)
            TASKS[task.id] = task
            threading.Thread(target=scan_worker, args=(task,), daemon=True).start()
            self._send(200, {"id": task.id, "path": task.path})

        elif u.path == "/api/cancel":
            task = TASKS.get(data.get("id", ""))
            if task:
                task.cancel = True
            self._send(200, {"ok": True})

        elif u.path == "/api/export":
            task = TASKS.get(data.get("id", ""))
            if not task:
                self._send(200, {"error": "未找到扫描结果"})
                return
            ext = data.get("ext", "")
            rows = get_task_files(task, ext)
            if not rows:
                self._send(200, {"error": "该格式没有文件，无法导出"})
                return
            fmt = (data.get("fmt") or "html").lower()
            dest = (data.get("output_dir") or "").strip().strip('"')
            if not dest or not os.path.isdir(dest):
                self._send(200, {"error": "输出文件夹不存在：" + dest})
                return
            name_part = "全部格式" if ext == "__all__" else safe_name(ext)
            out = os.path.join(
                dest, f"目录清单_{name_part}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{fmt}"
            )
            try:
                n = export_rows(rows, {"scope": task.path, "ext": name_part}, fmt, out)
                self._send(200, {"ok": True, "path": out, "rows": n,
                                 "size": os.path.getsize(out)})
            except Exception as e:
                self._send(200, {"error": "导出失败：" + str(e)})

        elif u.path == "/api/clip/set":
            mode = data.get("mode")
            if mode not in ("copy", "cut"):
                self._send(200, {"error": "无效的操作模式"})
                return
            paths = data.get("paths")
            if isinstance(paths, list) and paths:
                CLIPBOARD.update({"mode": mode, "paths": paths, "task_id": None,
                                  "ext": "已选文件", "source": "selection"})
                count = len(paths)
            else:
                task = TASKS.get(data.get("task_id", ""))
                if not task:
                    self._send(200, {"error": "扫描结果不存在"})
                    return
                ext = data.get("ext", "")
                files = get_task_files(task, ext)
                CLIPBOARD.update({"mode": mode, "paths": [], "task_id": task.id,
                                  "ext": ext, "source": "task"})
                count = len(files)
            self._send(200, {"ok": True, "count": count, "ext": CLIPBOARD["ext"]})

        elif u.path == "/api/clip/clear":
            CLIPBOARD.update({"mode": None, "paths": [], "task_id": None,
                              "ext": None, "source": None})
            self._send(200, {"ok": True})

        elif u.path == "/api/paste":
            dest = (data.get("dest") or "").strip().strip('"')
            if not dest:
                self._send(200, {"error": "请选择粘贴目标文件夹"})
                return
            if not CLIPBOARD.get("mode"):
                self._send(200, {"error": "剪贴板为空"})
                return
            files = resolve_clip_files()
            if not files:
                self._send(200, {"error": "剪贴板中没有可操作的文件"})
                return
            job = PasteJob(dest, data.get("collision", "skip"), files, CLIPBOARD["mode"])
            PASTE_JOBS[job.id] = job
            threading.Thread(target=paste_worker, args=(job,), daemon=True).start()
            self._send(200, {"id": job.id})

        elif u.path == "/api/quit":
            self._send(200, {"ok": True})

            def _shutdown():
                time.sleep(0.4)
                if SERVER is not None:
                    SERVER.shutdown()

            threading.Thread(target=_shutdown, daemon=True).start()
        else:
            self._send(404, {"error": "not found"})


def main():
    global SERVER
    port = PORT
    server = None
    for _ in range(20):
        try:
            server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
            break
        except OSError:
            port += 1
    if server is None:
        print("无法启动服务：端口均被占用")
        return
    SERVER = server
    url = f"http://127.0.0.1:{port}"
    print("文件分类与目录管理工具已启动：", url)
    if "--no-browser" not in sys.argv:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n已退出")


if __name__ == "__main__":
    main()
