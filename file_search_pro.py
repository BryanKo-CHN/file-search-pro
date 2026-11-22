#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
全盘文件名搜索工具（第三个程序）

功能：
  1. 全盘 / 文件夹逐层（BFS）扫描，长路径 + 隐藏项 + 目录联接全覆盖
  2. 关键词搜索：智能模糊（子序列）、精确子串、通配符 * ?、正则表达式
  3. 多关键词：空格=AND，竖线|=OR，减号-/叹号!=排除
  4. 只按“文件名主干”匹配（默认忽略扩展名）：记不清格式也能搜到
  5. 丰富筛选：扩展名 / 分类 / 大小区间 / 修改日期 / 路径包含
  6. 相关度评分与排序：相关度 / 名称 / 大小 / 时间 / 路径
  7. 搜索历史、保存的搜索（智能文件夹）
  8. 重复文件检测（按大小 + 摘要分组），只读、不改文件
  9. 结果导出 TXT / CSV / HTML；定位到文件所在文件夹
  10. 可视化选择扫描范围；无控制台 exe 的“退出程序”按钮

运行：
    python file_searcher.py
    python file_searcher.py --no-browser
"""

import os
import sys
import re
import csv
import json
import time
import uuid
import html
import base64
import string
import fnmatch
import hashlib
import shutil
import threading
import tempfile
import webbrowser
import multiprocessing
from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor, as_completed
from collections import deque, defaultdict
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs, quote
try:
    import ctypes
    from ctypes import wintypes
except Exception:
    ctypes = None
    wintypes = None

PORT = 8812
TASKS = {}          # 扫描任务
DUP_JOBS = {}       # 重复检测任务
CONTENT_JOBS = {}   # 正文搜索任务
PASTE_JOBS = {}     # 粘贴任务
CLIPBOARD = {"mode": None, "paths": [], "task_id": None, "ext": None, "source": None}
SERVER = None
DATA_FILE = os.path.join(os.path.expanduser("~"), ".file_searcher_data.json")
DATA = {"history": [], "saved": []}


# ----------------------------- 工具函数 -----------------------------

def load_data():
    global DATA
    try:
        if os.path.exists(DATA_FILE):
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                d = json.load(f)
            if isinstance(d, dict):
                DATA = {"history": d.get("history", [])[:30],
                        "saved": d.get("saved", [])[:50]}
    except Exception:
        pass


def save_data():
    try:
        with open(DATA_FILE, "w", encoding="utf-8") as f:
            json.dump(DATA, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


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


def parse_size(s):
    """把 '10MB' / '1.5G' / '500' 之类解析成字节数。纯数字按 MB。"""
    if s is None or str(s).strip() == "":
        return None
    t = str(s).strip().upper().replace(" ", "")
    m = re.match(r"^([\d.]+)\s*(B|KB|K|MB|M|GB|G|TB|T)?$", t)
    if not m:
        return None
    try:
        val = float(m.group(1))
    except Exception:
        return None
    unit = (m.group(2) or "MB")
    factor = {"B": 1, "KB": 1024, "K": 1024, "MB": 1024**2, "M": 1024**2,
              "GB": 1024**3, "G": 1024**3, "TB": 1024**4, "T": 1024**4}[unit]
    return int(val * factor)


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
    r"""Windows 下加 \\?\ 前缀，突破 MAX_PATH(260)。"""
    if os.name != "nt":
        return path
    p = os.path.abspath(path)
    if p.startswith("\\\\?\\"):
        return p
    if p.startswith("\\\\"):
        return "\\\\?\\UNC\\" + p[2:]
    return "\\\\?\\" + p


def from_long_path(path):
    r"""去掉 \\?\ 前缀。"""
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
    return "".join(c if (c.isalnum() or c in "-_") else "_" for c in str(s))[:60]


CATEGORIES = [
    ("办公文档", {"doc", "docx", "xls", "xlsx", "ppt", "pptx", "wps", "et", "dps", "odt", "ods", "odp", "rtf", "pdf"}),
    ("文本与数据", {"txt", "csv", "tsv", "log", "md", "json", "xml", "yaml", "yml", "ini", "cfg", "conf", "sql"}),
    ("图片", {"jpg", "jpeg", "png", "gif", "bmp", "svg", "webp", "ico", "tif", "tiff", "heic", "raw", "psd", "ai"}),
    ("视频", {"mp4", "avi", "mkv", "mov", "wmv", "flv", "webm", "m4v", "rmvb", "rm", "mpeg", "mpg", "3gp"}),
    ("音频", {"mp3", "wav", "flac", "aac", "ogg", "m4a", "wma", "ape", "mid"}),
    ("压缩包与镜像", {"zip", "rar", "7z", "tar", "gz", "bz2", "xz", "iso", "cab", "arj"}),
    ("程序与脚本", {"exe", "msi", "bat", "cmd", "sh", "ps1", "app", "deb", "rpm", "apk", "jar", "com", "scr"}),
    ("种子与下载", {"torrent", "magnet", "nfo", "sfv"}),
    ("编程开发", {"py", "js", "ts", "jsx", "tsx", "html", "htm", "css", "scss", "java", "c", "cpp", "h", "hpp", "cs", "go", "rs", "php", "rb", "swift", "kt", "vue", "ipynb"}),
    ("字体", {"ttf", "otf", "woff", "woff2", "eot", "fon"}),
    ("其他", set()),
]
EXT_CATEGORY = {}
for _cat, _exts in CATEGORIES:
    for _e in _exts:
        EXT_CATEGORY[_e] = _cat


def category_of(ext):
    return EXT_CATEGORY.get((ext or "").lower().lstrip("."), "其他")


def make_record(full, root_path):
    """生成文件记录，附带小写索引字段以加速搜索。"""
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
    mtime_ts = 0.0
    try:
        st = os.stat(full_stat)
        size = st.st_size
        mtime_ts = st.st_mtime
        mtime = datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        try:
            st = os.stat(full)
            size = st.st_size
            mtime_ts = st.st_mtime
            mtime = datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            pass
    ext = os.path.splitext(name)[1].lstrip(".").lower()
    stem = os.path.splitext(name)[0]
    return {
        "name": name, "path": full, "rel": rel, "ext": ext,
        "stem": stem, "category": category_of(ext),
        "size": size, "mtime": mtime, "mtime_ts": mtime_ts,
        # 小写索引（搜索时避免重复 lower）
        "_name_l": name.lower(), "_stem_l": stem.lower(), "_rel_l": rel.lower(),
    }


def browse_directory(path):
    if not path:
        return {"path": "", "parent": None, "is_root": True,
                "dirs": [{"name": d["path"], "path": d["path"], "hidden": False,
                          "drive": True} for d in list_drives()],
                "files": [], "files_total": 0}
    path = os.path.abspath(path)
    if not os.path.isdir(path):
        return {"error": "文件夹不存在或无法访问：" + path}
    dirs, files = [], []
    scan_path = to_long_path(path)
    try:
        for name in os.listdir(scan_path):
            full_long = os.path.join(scan_path, name)
            full = from_long_path(full_long)
            try:
                if os.path.isdir(full_long):
                    dirs.append({"name": name, "path": full, "hidden": is_hidden(full_long), "drive": False})
                elif os.path.isfile(full_long):
                    st = os.stat(full_long)
                    files.append({"name": name, "path": full, "size": st.st_size,
                                  "ext": os.path.splitext(name)[1].lstrip(".").lower(),
                                  "mtime": datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M:%S")})
            except Exception:
                continue
    except Exception as e:
        return {"error": "读取文件夹失败：" + str(e)}
    dirs.sort(key=lambda x: x["name"].lower())
    files.sort(key=lambda x: x["name"].lower())
    parent = os.path.dirname(path)
    if parent == path:
        parent = ""
    return {"path": path, "parent": parent, "is_root": False,
            "dirs": dirs, "files": files[:500], "files_total": len(files)}


def search_browse_dirs(root, keyword, max_depth=9, limit=300,
                       max_dirs=80000, budget=6.0):
    """文件夹浏览弹窗里的「关键词快速定位」：从当前目录逐层 BFS，按文件夹名
    模糊匹配（子序列 + 子串优先），返回可直接跳转的二/三级子文件夹。
    从盘符列表发起时各盘并行搜索，避免小盘把时间预算耗尽却搜不到目标盘。"""
    kw = (keyword or "").strip().lower()
    if not kw:
        return {"dirs": [], "total": 0, "truncated": False, "searched": 0}
    roots = []
    if root and os.path.isdir(root):
        roots.append(root)
    else:
        roots = [d["path"] for d in list_drives()]

    def scan_one(starts, deadline, dir_budget):
        found, seen, visited, truncated = [], set(), 0, False
        queue = deque((s, 0) for s in starts)
        while queue and visited < dir_budget:
            if time.time() > deadline:
                truncated = True
                break
            cur, depth = queue.popleft()
            key = os.path.normcase(os.path.abspath(cur))
            if key in seen:
                continue
            seen.add(key)
            visited += 1
            try:
                with os.scandir(to_long_path(cur)) as it:
                    for entry in it:
                        try:
                            if not entry.is_dir(follow_symlinks=False):
                                continue
                            full = from_long_path(entry.path)
                            score = _subseq_score(entry.name.lower(), kw)
                            if score:
                                found.append({
                                    "name": entry.name, "path": full, "parent": cur,
                                    "hidden": is_hidden(entry.path), "score": round(score, 1),
                                })
                            if depth < max_depth:
                                queue.append((full, depth + 1))
                        except Exception:
                            continue
            except Exception:
                continue
        return found, visited, truncated

    results, visited_all, truncated = [], 0, False
    deadline = time.time() + budget
    if len(roots) > 1:
        per_root = max(12000, max_dirs // len(roots))
        with ThreadPoolExecutor(max_workers=min(8, len(roots))) as pool:
            futures = [pool.submit(scan_one, [r], deadline, per_root) for r in roots]
            for fut in futures:
                try:
                    found, visited, tr = fut.result()
                except Exception:
                    continue
                results.extend(found)
                visited_all += visited
                truncated = truncated or tr
    else:
        found, visited, tr = scan_one(roots, deadline, max_dirs)
        results, visited_all, truncated = found, visited, tr

    results.sort(key=lambda x: (-x["score"], x["path"].lower()))
    return {"dirs": results[:limit], "total": len(results),
            "truncated": truncated or visited_all >= max_dirs, "searched": visited_all}


def set_system_clipboard(paths, cut=False):
    """把文件写入 Windows 系统剪贴板（CF_HDROP）。
    写完后可在任意文件夹 Ctrl+V 粘贴；cut=True 额外写入 Preferred DropEffect=MOVE，
    因此在资源管理器里 Ctrl+V 等同于「移动」。不依赖第三方库。"""
    if os.name != "nt" or ctypes is None:
        raise RuntimeError("当前环境不支持 Windows 系统剪贴板")
    files = []
    for p in paths or []:
        if not p:
            continue
        ap = os.path.abspath(p)
        if os.path.exists(ap):
            files.append(ap)
    if not files:
        raise RuntimeError("没有可放入剪贴板的文件（可能已被移动或删除）")

    CF_HDROP, CF_UNICODETEXT = 15, 13
    GMEM_MOVEABLE, GMEM_ZEROINIT = 0x0002, 0x0040
    DROPEFFECT_COPY, DROPEFFECT_MOVE = 1, 2

    class DROPFILES(ctypes.Structure):
        _fields_ = [("pFiles", wintypes.DWORD), ("pt", wintypes.POINT),
                    ("fNC", wintypes.BOOL), ("fWide", wintypes.BOOL)]

    u32, k32 = ctypes.windll.user32, ctypes.windll.kernel32
    u32.OpenClipboard.argtypes = [wintypes.HWND]
    u32.OpenClipboard.restype = wintypes.BOOL
    u32.CloseClipboard.argtypes = []
    u32.CloseClipboard.restype = wintypes.BOOL
    u32.EmptyClipboard.argtypes = []
    u32.EmptyClipboard.restype = wintypes.BOOL
    u32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
    u32.SetClipboardData.restype = wintypes.HANDLE
    u32.RegisterClipboardFormatW.argtypes = [wintypes.LPCWSTR]
    u32.RegisterClipboardFormatW.restype = wintypes.UINT
    k32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
    k32.GlobalAlloc.restype = wintypes.HGLOBAL
    k32.GlobalLock.argtypes = [wintypes.HGLOBAL]
    k32.GlobalLock.restype = wintypes.LPVOID
    k32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
    k32.GlobalUnlock.restype = wintypes.BOOL
    k32.GlobalFree.argtypes = [wintypes.HGLOBAL]
    k32.GlobalFree.restype = wintypes.HGLOBAL

    def alloc(data):
        h = k32.GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, len(data))
        if not h:
            raise RuntimeError("分配剪贴板内存失败")
        ptr = k32.GlobalLock(h)
        if not ptr:
            k32.GlobalFree(h)
            raise RuntimeError("锁定剪贴板内存失败")
        ctypes.memmove(ptr, data, len(data))
        k32.GlobalUnlock(h)
        return h

    opened = False
    for _ in range(12):
        if u32.OpenClipboard(None):
            opened = True
            break
        time.sleep(0.05)
    if not opened:
        raise RuntimeError("剪贴板被其他程序占用，请稍后重试")
    try:
        u32.EmptyClipboard()
        df = DROPFILES()
        df.pFiles = ctypes.sizeof(DROPFILES)
        df.fWide = 1
        listing = b"".join((p + "\x00").encode("utf-16le") for p in files) + b"\x00\x00"
        h_drop = alloc(bytes(df) + listing)
        if not u32.SetClipboardData(CF_HDROP, h_drop):
            k32.GlobalFree(h_drop)
            raise RuntimeError("写入剪贴板失败")

        fmt = u32.RegisterClipboardFormatW("Preferred DropEffect")
        if fmt:
            effect = wintypes.DWORD(DROPEFFECT_MOVE if cut else DROPEFFECT_COPY)
            h_effect = alloc(ctypes.string_at(ctypes.byref(effect), ctypes.sizeof(effect)))
            if not u32.SetClipboardData(fmt, h_effect):
                k32.GlobalFree(h_effect)

        text = ("\r\n".join(files) + "\r\n").encode("utf-16le") + b"\x00\x00"
        h_text = alloc(text)
        if not u32.SetClipboardData(CF_UNICODETEXT, h_text):
            k32.GlobalFree(h_text)
    finally:
        u32.CloseClipboard()
    return len(files)


def open_file_external(path):
    """用系统默认程序打开文件或文件夹。"""
    p = os.path.abspath(path)
    if not os.path.exists(p):
        raise RuntimeError("文件不存在：" + p)
    if os.name == "nt":
        os.startfile(p)  # noqa: S606 - 本地文件，由用户点击触发
    elif sys.platform == "darwin":
        import subprocess
        subprocess.Popen(["open", p])
    else:
        import subprocess
        subprocess.Popen(["xdg-open", p])
    return True


# ----------------------------- 扫描任务（逐层 BFS） -----------------------------

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
        self.by_ext = defaultdict(lambda: {"count": 0, "size": 0})
        self.errors = []
        self.skipped = []
        self.followed_links = 0
        self.current_level = 0
        self.max_level = 0
        self.level_dirs = {}
        self.level_files = {}
        self.content_cache = {}   # path -> (text, err)，正文缓存，重复搜索秒出
        self.start = time.time()
        self.end = None
        self.cancel = False

    def snapshot(self):
        exts = []
        for ext, info in self.by_ext.items():
            exts.append({"ext": ext or "(无扩展名)", "key": ext,
                         "category": category_of(ext), "count": info["count"], "size": info["size"]})
        exts.sort(key=lambda x: (-x["count"], x["ext"]))
        return {"id": self.id, "path": self.path, "status": self.status,
                "status_text": self.status_text, "dirs": self.count_dir,
                "files": self.count_file, "total_size": self.total_size,
                "errors": len(self.errors), "skipped": len(self.skipped),
                "skipped_samples": self.skipped[:20], "followed_links": self.followed_links,
                "current_level": self.current_level, "max_level": self.max_level,
                "level_dirs": dict(self.level_dirs), "level_files": dict(self.level_files),
                "elapsed": round((self.end or time.time()) - self.start, 2),
                "extensions": exts}


def scan_worker(task):
    """逐层广度优先扫描：根目录 → 第1层 → 第2层 → … → 最内层。"""
    max_depth = int(task.opts.get("max_depth", 0) or 0)
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
        try:
            import ctypes
            attrs = ctypes.windll.kernel32.GetFileAttributesW(str(path))
            if attrs != -1 and (attrs & 0x400):
                return True
        except Exception:
            pass
        try:
            return os.path.islink(path)
        except Exception:
            return False

    queue = deque([(scan_root, 0)])
    visited = set()
    try:
        visited.add(os.path.realpath(scan_root).lower())
    except Exception:
        pass

    try:
        while queue:
            if task.cancel:
                task.status = "cancelled"
                task.end = time.time()
                task.status_text = "已取消"
                return
            root, level = queue.popleft()
            task.current_level = level
            task.max_level = max(task.max_level, level)
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
                try:
                    rec = make_record(entry.path, scan_root)
                except Exception as e:
                    _on_error(entry.path, e)
                    continue
                # 不向前端下发内部小写索引字段
                for k in ("_name_l", "_stem_l", "_rel_l"):
                    rec.pop(k, None)
                task.files.append(rec)
                task.count_file += 1
                task.total_size += rec["size"]
                info = task.by_ext[rec["ext"]]
                info["count"] += 1
                info["size"] += rec["size"]
                task.level_files[level] = task.level_files.get(level, 0) + 1

            task.count_dir += len(subdirs)
            task.level_dirs[level] = task.level_dirs.get(level, 0) + len(subdirs)

            if not (max_depth and level >= max_depth):
                for entry in subdirs:
                    child = entry.path
                    if _is_reparse(child):
                        try:
                            cr = os.path.realpath(child).lower()
                        except Exception:
                            cr = child.lower()
                        if cr in visited:
                            continue
                        visited.add(cr)
                        task.followed_links += 1
                    queue.append((child, level + 1))

            task.status_text = (f"逐层扫描中 · 第 {level} 层 · 已发现 {task.count_file} 个文件 / "
                                f"{task.count_dir} 个文件夹"
                                + (f"（跳过 {len(task.skipped)} 处）" if task.skipped else ""))
    except Exception as e:
        task.errors.append(str(e))

    task.status = "done"
    task.end = time.time()
    extra = [f"共 {task.max_level + 1} 层"]
    if task.followed_links:
        extra.append(f"跟随 {task.followed_links} 个链接")
    if task.skipped:
        extra.append(f"跳过 {len(task.skipped)} 处不可访问")
    task.status_text = (f"扫描完成：{task.count_file} 个文件，涉及 {len(task.by_ext)} 种后缀"
                        + ("（" + "，".join(extra) + "）"))


# ----------------------------- 搜索核心 -----------------------------

def _rebuild_indexes(task):
    """为扫描记录补充小写索引字段（扫描时为节省内存未下发，搜索时重建一次）。"""
    for r in task.files:
        if "_stem_l" not in r:
            r["_name_l"] = r["name"].lower()
            r["_stem_l"] = r.get("stem", os.path.splitext(r["name"])[0]).lower()
            r["_rel_l"] = r["rel"].lower()


def _subseq_score(text, q):
    """子序列模糊匹配：按顺序找到 q 的每个字符，返回评分；不匹配返回 0。
    评分偏好：连续命中、起始命中、覆盖比例高、间隔小。"""
    n, m = len(text), len(q)
    if m == 0:
        return 1
    if m > n:
        return 0
    # 贪心从左到右匹配
    idx = -1
    first = -1
    gaps = 0
    matched = 0
    for ch in q:
        pos = text.find(ch, idx + 1)
        if pos < 0:
            return 0
        if first < 0:
            first = pos
        if idx >= 0:
            gaps += pos - idx - 1
        idx = pos
        matched += 1
    coverage = matched / max(1, n)
    contiguity = matched - min(gaps, matched * 3)
    score = 40 + (30 if first <= 1 else 0) + contiguity * 3 + coverage * 20 - min(gaps, 20)
    return max(1, score)


def _term_score(rec, term, scope, fuzzy=True):
    """对单条记录、单个词计算匹配分（0 表示不匹配）。
    fuzzy=False 时只做子串匹配（精确包含模式）。"""
    if scope == "path":
        fields = [rec.get("_rel_l", "")]
    elif scope == "name_ext":
        fields = [rec.get("_name_l", "")]
    else:  # 默认只按文件名主干（忽略扩展名），解决“记不清格式”的痛点
        fields = [rec.get("_stem_l", ""), rec.get("_name_l", "")]
    best = 0
    for text in fields:
        if not text:
            continue
        pos = text.find(term)
        if pos >= 0:
            # 子串命中：越高分越靠前
            s = 100 + min(30, len(term) * 4) + (25 if pos == 0 else 0)
            best = max(best, s)
            continue
        if fuzzy:
            s = _subseq_score(text, term)
            if s > 0:
                best = max(best, s)
    return best


def parse_query(query, mode, scope):
    """把查询串解析成 OR 组：每组 [and_terms, exclude_terms]。"""
    query = (query or "").strip()
    if not query:
        return []
    if mode in ("regex", "wildcard"):
        # 特殊：单组单 term，交给正则/通配符处理
        return [[[query], []]]

    groups = []
    for group_raw in query.split("|"):
        and_terms, excl = [], []
        for tok in group_raw.split():
            if len(tok) > 1 and tok[0] in ("-", "!"):
                excl.append(tok[1:].lower())
            elif tok not in ("-", "!"):
                and_terms.append(tok.lower())
        if and_terms or excl:
            groups.append([and_terms, excl])
    return groups


def _match_regex(rec, pattern, scope, compiled):
    if scope == "path":
        text = rec.get("rel", "")
    elif scope == "name_ext":
        text = rec.get("name", "")
    else:
        text = rec.get("stem", "")
    try:
        m = compiled.search(text)
        return (70 if m else 0)
    except Exception:
        return 0


def _match_wildcard(rec, pattern, scope):
    if scope == "path":
        text = rec.get("rel", "")
    elif scope == "name_ext":
        text = rec.get("name", "")
    else:
        text = rec.get("stem", "")
    try:
        return 70 if fnmatch.fnmatch(text.lower(), pattern.lower()) else 0
    except Exception:
        return 0


def apply_filters(rec, f):
    ext = (f.get("ext") or "").strip().lower().lstrip(".")
    if ext:
        wanted = {x.strip().lstrip(".") for x in ext.replace(",", " ").split() if x.strip()}
        if rec["ext"] not in wanted:
            return False
    cat = (f.get("category") or "").strip()
    if cat and rec["category"] != cat:
        return False
    smin = parse_size(f.get("size_min"))
    smax = parse_size(f.get("size_max"))
    if smin is not None and rec["size"] < smin:
        return False
    if smax is not None and rec["size"] > smax:
        return False
    df = (f.get("date_from") or "").strip()
    dt = (f.get("date_to") or "").strip()
    if df or dt:
        day = rec["mtime"][:10] if rec["mtime"] else ""
        if df and (not day or day < df):
            return False
        if dt and (not day or day > dt):
            return False
    pc = (f.get("path_contains") or "").strip().lower()
    if pc and pc not in rec["_rel_l"]:
        return False
    return True


def search_files(task, params):
    """返回 (total, rows, took_ms, positive_terms)。"""
    query = params.get("query", "")
    mode = params.get("mode", "smart")
    scope = params.get("scope", "stem")
    sort = params.get("sort", "relevance")
    f = params.get("filters") or {}
    try:
        offset = max(0, int(params.get("offset", 0)))
        limit = min(1000, max(1, int(params.get("limit", 200))))
        if params.get("_unlimited"):
            limit = 100000000
    except Exception:
        offset, limit = 0, 200

    _rebuild_indexes(task)
    t0 = time.time()
    groups = parse_query(query, mode, scope)
    compiled = None
    if mode == "regex":
        try:
            compiled = re.compile(query, re.I)
        except Exception:
            compiled = None

    matched = []
    files = task.files
    if not groups and not any([f.get("ext"), f.get("category"), f.get("size_min"),
                               f.get("size_max"), f.get("date_from"), f.get("date_to"),
                               f.get("path_contains")]):
        matched = list(files)
    else:
        for rec in files:
            if groups:
                score = 0
                any_group = False
                for and_terms, excl in groups:
                    ok = True
                    gs = 0
                    for term in and_terms:
                        if mode == "regex":
                            s = _match_regex(rec, query, scope, compiled) if compiled else 0
                        elif mode == "wildcard":
                            s = _match_wildcard(rec, query, scope)
                        else:
                            s = _term_score(rec, term, scope, fuzzy=(mode == "smart"))
                        if s <= 0:
                            ok = False
                            break
                        gs += s
                    if not ok:
                        continue
                    bad = False
                    for term in excl:
                        if mode in ("regex", "wildcard"):
                            s = _match_wildcard(rec, term, scope) if mode == "wildcard" else _match_regex(rec, term, scope, compiled)
                        else:
                            s = _term_score(rec, term, scope, fuzzy=(mode == "smart"))
                        if s > 0:
                            bad = True
                            break
                    if bad:
                        continue
                    any_group = True
                    score = max(score, gs)
                if not any_group:
                    continue
            else:
                score = 0
            if not apply_filters(rec, f):
                continue
            rec2 = dict(rec)
            rec2["_score"] = score
            matched.append(rec2)

    # 排序
    if sort == "name":
        matched.sort(key=lambda r: r["name"].lower())
    elif sort == "path":
        matched.sort(key=lambda r: r["rel"].lower())
    elif sort == "size":
        matched.sort(key=lambda r: -r["size"])
    elif sort == "mtime":
        matched.sort(key=lambda r: -r["mtime_ts"])
    else:  # relevance
        matched.sort(key=lambda r: (-r.get("_score", 0), r["name"].lower()))

    total = len(matched)
    page = matched[offset:offset + limit]
    for r in page:
        r.pop("_name_l", None); r.pop("_stem_l", None); r.pop("_rel_l", None)
    took = int((time.time() - t0) * 1000)
    pos_terms = []
    for grp in groups:
        pos_terms.extend(grp[0])
    return total, page, took, pos_terms


def export_rows(rows, meta, fmt, out_path):
    fmt = fmt.lower().lstrip(".")
    rows = sorted(rows, key=lambda r: r["path"].lower())
    if fmt == "csv":
        with open(out_path, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(["序号", "文件名", "后缀", "分类", "大小(字节)", "大小",
                        "修改时间", "相对路径", "完整路径"])
            for i, r in enumerate(rows, 1):
                w.writerow([i, r["name"], r["ext"], r["category"], r["size"],
                            human_size(r["size"]), r["mtime"], r["rel"], r["path"]])
        return len(rows)
    if fmt == "txt":
        with open(out_path, "w", encoding="utf-8") as fh:
            fh.write("=" * 78 + "\n文件搜索结果清单\n" + "=" * 78 + "\n")
            fh.write(f"扫描范围 : {meta.get('scope', '')}\n")
            fh.write(f"搜索词   : {meta.get('query', '') or '(全部)'}\n")
            fh.write(f"生成时间 : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            fh.write(f"结果数量 : {len(rows)}\n")
            fh.write(f"合计大小 : {human_size(sum(r['size'] for r in rows))}\n")
            fh.write("-" * 78 + "\n")
            for i, r in enumerate(rows, 1):
                fh.write(f"{i}. {r['name']}  [{human_size(r['size'])}]  {r['rel']}\n")
        return len(rows)
    if fmt in ("html", "htm"):
        esc = html.escape
        trs = []
        for i, r in enumerate(rows, 1):
            trs.append(
                "<tr><td class='n'>{i}</td><td class='nm'>{nm}</td>"
                "<td><span class='tag'>{ext}</span></td><td>{cat}</td>"
                "<td class='sz'>{sz}</td><td class='tm'>{tm}</td>"
                "<td class='rp'>{rel}</td></tr>".format(
                    i=i, nm=esc(r["name"]), ext=esc(r["ext"] or "-"),
                    cat=esc(r["category"]), sz=human_size(r["size"]),
                    tm=esc(r["mtime"]), rel=esc(r["rel"])))
        doc = """<!doctype html><html><head><meta charset="utf-8">
<title>文件搜索结果</title><style>
body{font-family:"Microsoft YaHei",system-ui,sans-serif;margin:24px;color:#1f2937;background:#fff}
h1{font-size:20px;margin:0 0 6px}.meta{color:#6b7280;font-size:13px;margin-bottom:16px;line-height:1.7}
table{border-collapse:collapse;width:100%;font-size:13px}th,td{border:1px solid #e5e7eb;padding:7px 9px;text-align:left}
th{background:#f3f4f6;position:sticky;top:0}.n{color:#9ca3af;width:46px}.nm{font-weight:600}
.tag{background:#eff6ff;color:#2563eb;padding:1px 7px;border-radius:10px;font-size:11px}
.sz,.tm{white-space:nowrap;color:#4b5563}.rp{color:#6b7280;word-break:break-all}
tr:nth-child(even){background:#fafafa}</style></head><body>
<h1>文件搜索结果清单</h1>
<div class="meta">扫描范围：__SCOPE__<br>搜索词：__QUERY__<br>结果数量：__CNT__ · 合计大小：__TOTAL__<br>生成时间：__NOW__</div>
<table><thead><tr><th>#</th><th>文件名</th><th>后缀</th><th>分类</th><th>大小</th><th>修改时间</th><th>相对路径</th></tr></thead>
<tbody>__ROWS__</tbody></table></body></html>"""
        doc = (doc.replace("__SCOPE__", html.escape(meta.get("scope", "")))
                  .replace("__QUERY__", html.escape(meta.get("query", "") or "(全部)"))
                  .replace("__CNT__", str(len(rows)))
                  .replace("__TOTAL__", human_size(sum(r["size"] for r in rows)))
                  .replace("__NOW__", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                  .replace("__ROWS__", "".join(trs)))
        with open(out_path, "w", encoding="utf-8") as fh:
            fh.write(doc)
        return len(rows)
    raise ValueError("不支持的格式：" + str(fmt))


def export_tree(rows, meta, out_path):
    """把文件清单导出为「可折叠目录树」HTML：按相对路径建树，每个文件夹用
    <details><summary> 实现展开/折叠，双击即可在浏览器里像资源管理器一样逐层翻阅。"""
    def esc(s):
        return html.escape(str(s))

    # 导出页点击文件名时回调本地接口（用系统默认程序打开）；
    # 若程序已关闭，JS 会自动回退到 file:// 链接。
    server_url = ""
    if SERVER is not None:
        try:
            server_url = "http://127.0.0.1:%d" % SERVER.server_address[1]
        except Exception:
            server_url = ""

    def file_uri(p):
        q = str(p).replace("\\", "/")
        if q.startswith("//"):
            return "file:" + quote(q, safe="/")
        return "file:///" + quote(q.lstrip("/"), safe="/:")

    # 构建树：目录节点 = {name: node}，文件放 __files__
    root = {"__files__": [], "__dirs__": {}}
    for r in rows:
        parts = [p for p in re.split(r"[\\/]+", r["rel"] or "") if p and p not in (".", "..")]
        if len(parts) <= 1:
            root["__files__"].append(r)
            continue
        node = root
        for seg in parts[:-1]:
            dirs = node["__dirs__"]
            if seg not in dirs:
                dirs[seg] = {"__files__": [], "__dirs__": {}}
            node = dirs[seg]
        node["__files__"].append(r)

    def render_dir(node, path_prefix, depth):
        lines = []
        files = sorted(node.get("__files__", []), key=lambda x: x["name"].lower())
        dirs = sorted(node.get("__dirs__", {}).items(), key=lambda x: x[0].lower())
        # 文件夹在前，文件在后
        for name, child in dirs:
            child_path = path_prefix + name
            cnt = _count_node(child)
            lines.append(
                '<div class="folder"><details><summary>'
                '<span class="ico">📁</span><span class="fname">' + esc(name) + '</span>'
                '<span class="cnt">' + str(cnt) + ' 项</span>'
                '</summary><div class="children">'
            )
            lines.append(render_dir(child, child_path + "/", depth + 1))
            lines.append('</div></details></div>')
        for r in files:
            api_url = (server_url + "/api/open?p=" + quote(r["path"], safe="")) if server_url else ""
            uri = file_uri(r["path"])
            lines.append(
                '<div class="file"><span class="ico">📄</span>'
                '<a class="fname" href="' + esc(api_url or uri) + '" data-api="' + esc(api_url) +
                '" data-file="' + esc(uri) + '" title="点击打开：' + esc(r["path"]) + '">' +
                esc(r["name"]) + '</a>'
                '<span class="sz">' + human_size(r["size"]) + '</span>'
                '<span class="tm">' + esc(r["mtime"][:10] if r.get("mtime") else "") + '</span>'
                '</div>'
            )
        return "\n".join(lines)

    def _count_node(node):
        return len(node.get("__files__", [])) + sum(_count_node(c) for c in node.get("__dirs__", {}).values())

    body = render_dir(root, "", 0)
    total = len(rows)
    doc = r"""<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">
<title>可折叠目录树</title><style>
body{font-family:"Microsoft YaHei",system-ui,sans-serif;margin:0;background:#f5f7fa;color:#1f2937}
.wrap{max-width:1080px;margin:24px auto;background:#fff;border-radius:12px;padding:26px 30px;
 box-shadow:0 4px 20px rgba(0,0,0,.06)}
h1{font-size:21px;margin:0 0 6px}.meta{color:#6b7280;font-size:13px;line-height:1.8;margin-bottom:16px}
.bar{background:#eff6ff;border:1px solid #dbeafe;border-radius:8px;padding:9px 12px;margin-bottom:14px;font-size:13px}
.bar button{padding:5px 12px;margin-right:8px;font-size:12px;border:1px solid #bfdbfe;background:#fff;
 color:#1d4ed8;border-radius:6px;cursor:pointer}
.bar button:hover{background:#dbeafe}
.tree{font-size:13px;line-height:1.5;border:1px solid #e5e7eb;border-radius:9px;padding:14px 16px}
details{margin:1px 0}summary{cursor:pointer;padding:3px 4px;border-radius:5px;list-style:none;user-select:none}
summary::-webkit-details-marker{display:none}summary:hover{background:#f1f5f9}
summary .ico{margin-right:6px}summary .fname{font-weight:600;color:#1e40af}
summary .cnt{color:#9ca3af;font-size:11px;margin-left:8px}
.children{margin-left:18px;border-left:1px dashed #cbd5e1;padding-left:10px}
.file{padding:3px 4px;border-radius:5px;display:flex;gap:8px;align-items:center}
.file:hover{background:#f8fafc}
.file a.fname{color:#1d4ed8;text-decoration:none;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;
 border-bottom:1px dashed transparent}
.file a.fname:hover{color:#2563eb;border-bottom-color:#93c5fd}
.file a.fname:after{content:" ↗";font-size:11px;color:#93c5fd}
.file .ico{opacity:.7}.sz{color:#64748b;font-size:11px;margin-left:auto;white-space:nowrap}
.tm{color:#9ca3af;font-size:11px;white-space:nowrap}
</style></head><body><div class="wrap">
<h1>📂 可折叠目录树</h1>
<div class="meta">扫描范围：__SCOPE__<br>文件总数：__CNT__ · 合计大小：__TOTAL__<br>生成时间：__NOW__</div>
<div class="bar"><button onclick="expandAll(true)">全部展开</button>
<button onclick="expandAll(false)">全部收起</button>
<span style="color:#6b7280">点击文件夹名展开/折叠；<b>点击文件名可直接打开原文件</b>（程序运行时用系统默认程序打开，程序关闭后回退浏览器打开）。</span></div>
<div class="tree">__BODY__</div>
</div>
<script>
function expandAll(open){document.querySelectorAll('details').forEach(function(d){d.open=open;});}
function openTreeFile(ev,a){
 ev.preventDefault();
 var api=a.getAttribute('data-api'), file=a.getAttribute('data-file');
 function fallback(){if(file){location.href=file;}}
 if(!api){fallback();return;}
 fetch(api,{mode:'cors'}).then(function(r){return r.json();}).then(function(j){
  if(!j.ok){fallback();}
 }).catch(fallback);
}
document.querySelectorAll('a[data-api]').forEach(function(a){
 a.addEventListener('click',function(e){openTreeFile(e,a);});
});
</script>
</body></html>"""
    doc = (doc.replace("__SCOPE__", esc(meta.get("scope", "")))
              .replace("__CNT__", str(total))
              .replace("__TOTAL__", human_size(sum(r["size"] for r in rows)))
              .replace("__NOW__", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
              .replace("__BODY__", body))
    with open(out_path, "w", encoding="utf-8") as fh:
        fh.write(doc)
    return len(rows)


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
        src_op = to_long_path(src)
        dst_op = to_long_path(dst)
        try:
            if src_op.lower() == dst_op.lower():
                job.skipped += 1
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
            files = list(task.files) if ext == "__all__" else [f for f in task.files if f["ext"] == ext]
            for f in files:
                result.append((f["path"], f["name"]))
    elif CLIPBOARD.get("paths"):
        for p in CLIPBOARD["paths"]:
            if os.path.isfile(to_long_path(p)):
                result.append((p, os.path.basename(p)))
    return result


# ----------------------------- 重复文件检测 -----------------------------

class DupJob:
    def __init__(self, task_id, min_size):
        self.id = uuid.uuid4().hex[:12]
        self.task_id = task_id
        self.min_size = min_size
        self.status = "running"
        self.status_text = "准备检测..."
        self.groups = []
        self.start = time.time()
        self.end = None
        self.cancel = False
        self.scanned = 0
        self.total = 0

    def snapshot(self):
        wasted = sum(g["size"] * (len(g["files"]) - 1) for g in self.groups)
        dup_files = sum(len(g["files"]) for g in self.groups)
        return {"id": self.id, "status": self.status, "status_text": self.status_text,
                "scanned": self.scanned, "total": self.total,
                "groups": self.groups[:100], "group_count": len(self.groups),
                "dup_files": dup_files, "wasted": wasted,
                "elapsed": round((self.end or time.time()) - self.start, 2)}


def _quick_hash(path, size):
    """快速摘要：文件大小 + 头部 64KB + 尾部 64KB（对同大小文件再全量校验时可选）。"""
    h = hashlib.md5()
    h.update(str(size).encode())
    try:
        with open(to_long_path(path), "rb") as fh:
            head = fh.read(65536)
            h.update(head)
            if size > 131072:
                fh.seek(-65536, os.SEEK_END)
                h.update(fh.read(65536))
    except Exception:
        return None
    return h.hexdigest()


def dup_worker(job):
    task = TASKS.get(job.task_id)
    if not task:
        job.status = "error"
        job.status_text = "扫描结果已不存在"
        job.end = time.time()
        return
    candidates = [r for r in task.files if r["size"] >= job.min_size and r["size"] > 0]
    job.total = len(candidates)
    by_size = defaultdict(list)
    for r in candidates:
        by_size[r["size"]].append(r)
    # 仅处理有同体积候选的组
    same = [(sz, rs) for sz, rs in by_size.items() if len(rs) > 1]
    job.status_text = f"正在校验 {len(same)} 组同大小文件..."
    groups = []
    for gi, (sz, rs) in enumerate(same):
        if job.cancel:
            job.status = "cancelled"
            job.end = time.time()
            return
        buckets = defaultdict(list)
        for r in rs:
            hh = _quick_hash(r["path"], sz)
            if hh:
                buckets[hh].append(r)
        for hh, items in buckets.items():
            if len(items) > 1:
                groups.append({"size": sz, "hash": hh[:12],
                               "files": [{"name": x["name"], "path": x["path"],
                                          "rel": x["rel"], "mtime": x["mtime"]} for x in items]})
        job.scanned = gi + 1
        if (gi + 1) % 50 == 0:
            job.status_text = f"已处理 {gi + 1}/{len(same)} 组，发现 {len(groups)} 组重复"
    groups.sort(key=lambda g: -(g["size"] * (len(g["files"]) - 1)))
    job.groups = groups
    job.status = "done"
    job.end = time.time()
    dup = sum(len(g["files"]) for g in groups)
    job.status_text = f"检测完成：{len(groups)} 组重复，涉及 {dup} 个文件"


# ----------------------------- 正文搜索 -----------------------------

# 需要读取正文的扩展名（决定哪些文件会被打开解析）
TEXT_EXTS = {"txt", "md", "csv", "tsv", "log", "json", "xml", "yaml", "yml", "ini",
             "cfg", "conf", "sql", "py", "js", "ts", "jsx", "tsx", "java", "c", "cpp",
             "h", "hpp", "cs", "go", "rs", "php", "rb", "swift", "kt", "html", "htm",
             "css", "scss", "bat", "cmd", "sh", "ps1", "rtf", "tex", "srt", "lrc"}
DOC_EXTS = {"docx"}
XLS_EXTS = {"xlsx"}
PPT_EXTS = {"pptx"}
PDF_EXTS = {"pdf"}
OLD_DOC_EXTS = {"doc", "xls", "ppt", "wps", "et", "dps"}
# OLE2 文字类格式（Word 97-2003 二进制）：可用分片表精确提取正文
OLE2_WORD_EXTS = {"doc", "wps"}
CONTENT_EXTS = TEXT_EXTS | DOC_EXTS | XLS_EXTS | PPT_EXTS | PDF_EXTS | OLD_DOC_EXTS
MAX_CONTENT_BYTES = 40 * 1024 * 1024   # 单个文件正文上限 40MB，跳过超大文件


def _decode_bytes(raw):
    """多编码尝试解码，优先中文常见编码。"""
    for enc in ("utf-8", "gb18030", "utf-16", "big5", "latin-1"):
        try:
            return raw.decode(enc)
        except Exception:
            continue
    return raw.decode("utf-8", errors="ignore")


def _strip_xml(xml_text):
    """去 XML 标签，保留文本，把常见段落边界换成空格。"""
    x = re.sub(r"<w:tab[^>]*/>|</w:p>|</a:p>|<w:br[^>]*/>", " ", xml_text)
    x = re.sub(r"<[^>]+>", "", x)
    x = (x.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
          .replace("&quot;", '"').replace("&apos;", "'").replace("&#10;", " "))
    return re.sub(r"[ \t]+", " ", x)


_ole2_word_cache = {}       # path+mtime+size -> (text, err)，进程内缓存
_ole2_word_cache_lock = threading.Lock()
_OLE2_CACHE_MAX = 24


def _extract_ole2_word(lp):
    """用 OLE2 分片表（piece table）精确提取 Word 97-2003 .doc 正文。

    基于 legacy-doc-text-extract 技能：读 WordDocument 流的 FIB，定位
    1Table/0Table 流里的 Clx→Pcdt→PlcPcd，逐片按压缩标志用 cp1252 或
    utf-16-le 解码，再清理 Word 控制字符。失败时抛异常由调用方回退。
    """
    import struct
    import olefile

    ole = olefile.OleFileIO(lp)
    try:
        wd = ole.openstream("WordDocument").read()
        flags = struct.unpack_from("<H", wd, 0x0A)[0]
        tbl_name = "1Table" if (flags >> 9) & 1 else "0Table"
        csw = struct.unpack_from("<H", wd, 32)[0]
        off = 34 + csw * 2
        cslw = struct.unpack_from("<H", wd, off)[0]
        off += 2 + cslw * 4
        cbRgFcLcb = struct.unpack_from("<H", wd, off)[0]
        rgFcLcb_off = off + 2
        if cbRgFcLcb < 34:
            raise ValueError("FIB 太短，非标准 .doc")
        fcClx, lcbClx = struct.unpack_from("<II", wd, rgFcLcb_off + 33 * 8)
        tbl = ole.openstream(tbl_name).read()
        clx = tbl[fcClx:fcClx + lcbClx]
        i, pcdt = 0, None
        while i < len(clx):
            t = clx[i]
            if t == 1:
                cb = struct.unpack_from("<h", clx, i + 1)[0]
                i += 3 + cb
            elif t == 2:
                lcb = struct.unpack_from("<I", clx, i + 1)[0]
                pcdt = clx[i + 5:i + 5 + lcb]
                break
            else:
                break
        if pcdt is None or len(pcdt) < 4:
            raise ValueError("未找到 Pcdt")
        n = (len(pcdt) - 4) // 12
        cps = struct.unpack_from("<%dI" % (n + 1), pcdt, 0)
        out, pcd_base = [], (n + 1) * 4
        for k in range(n):
            pcd = pcdt[pcd_base + k * 8: pcd_base + k * 8 + 8]
            fc = struct.unpack_from("<I", pcd, 2)[0]
            compressed = bool(fc & 0x40000000)
            fc &= 0x3FFFFFFF
            nchars = cps[k + 1] - cps[k]
            if compressed:
                raw = wd[fc // 2: fc // 2 + nchars]
                s = raw.decode("cp1252", errors="replace")
            else:
                raw = wd[fc: fc + nchars * 2]
                s = raw.decode("utf-16-le", errors="replace")
            out.append(s)
        text = "".join(out)
    finally:
        try:
            ole.close()
        except Exception:
            pass

    # 清理 Word 控制字符，保留段落/制表边界
    reps = {"\r": "\n", "\x07": "\t", "\x0b": "\n", "\x0c": "\n",
            "\x1e": "-", "\x1f": "", "\x13": "", "\x14": "", "\x15": "",
            "\x01": "", "\x02": "", "\x05": "", "\x08": "", "\u3000": " "}
    for a, b in reps.items():
        text = text.replace(a, b)
    text = "".join(ch for ch in text if ch >= " " or ch in "\n\t")
    lines = [ln.rstrip() for ln in text.split("\n")]
    res, blank = [], 0
    for ln in lines:
        if ln.strip() == "":
            blank += 1
            if blank > 2:
                continue
        else:
            blank = 0
        res.append(ln)
    return "\n".join(res)


def _extract_ole2_word_cached(lp):
    """带缓存的 .doc 正文提取（文件未变则复用结果）。"""
    try:
        st = os.stat(lp)
        key = (lp, st.st_mtime, st.st_size)
    except Exception:
        key = (lp,)
    with _ole2_word_cache_lock:
        hit = _ole2_word_cache.get(key)
    if hit is not None:
        return hit
    try:
        text = _extract_ole2_word(lp)
        result = (text, None if text.strip() else "正文为空")
    except Exception as e:
        result = ("", "OLE2 解析失败：" + str(e))
    with _ole2_word_cache_lock:
        if len(_ole2_word_cache) >= _OLE2_CACHE_MAX:
            try:
                _ole2_word_cache.pop(next(iter(_ole2_word_cache)))
            except Exception:
                pass
        _ole2_word_cache[key] = result
    return result


def extract_text(path, ext, needle=None, stop_text=None):
    """尽力提取一个文件的正文文本，失败返回 (文本, 错误信息)。只读打开。
    needle 为关键词字节集合时，先做原始字节预筛，明显不含关键词的文件直接跳过。
    stop_text 为小写关键词列表时，逐页提取命中后提前停止，避免解析整本大文档。"""
    ext = (ext or "").lower().lstrip(".")
    lp = to_long_path(path)
    try:
        st = os.stat(lp)
        if st.st_size > MAX_CONTENT_BYTES:
            return "", "文件超过40MB已跳过"
    except Exception as e:
        return "", "无法读取：" + str(e)
    # 原始字节预筛：对未压缩文本(txt/doc)很有效，docx/xlsx/pdf 因压缩可能误判，跳过预筛
    if needle and ext in ((TEXT_EXTS | OLD_DOC_EXTS) - OLE2_WORD_EXTS):
        try:
            with open(lp, "rb") as fh:
                raw = fh.read(MAX_CONTENT_BYTES)
            if not any(nb in raw for nb in needle):
                return "", None
        except Exception:
            pass
    try:
        if ext in TEXT_EXTS:
            with open(lp, "rb") as fh:
                return _decode_bytes(fh.read(MAX_CONTENT_BYTES)), None
        if ext in DOC_EXTS:
            import zipfile
            with zipfile.ZipFile(lp) as z:
                parts = [n for n in z.namelist()
                         if n == "word/document.xml" or n.startswith("word/header")
                         or n.startswith("word/footer") or n == "word/footnotes.xml"
                         or n == "word/endnotes.xml"]
                chunks = []
                for n in parts:
                    try:
                        chunks.append(_strip_xml(z.read(n).decode("utf-8", errors="ignore")))
                    except Exception:
                        pass
                return "\n".join(chunks), None
        if ext in XLS_EXTS:
            import zipfile
            with zipfile.ZipFile(lp) as z:
                chunks = []
                if "xl/sharedStrings.xml" in z.namelist():
                    raw = z.read("xl/sharedStrings.xml").decode("utf-8", errors="ignore")
                    chunks.append(_strip_xml(raw))
                for n in z.namelist():
                    if n.startswith("xl/worksheets/") and n.endswith(".xml"):
                        try:
                            chunks.append(_strip_xml(z.read(n).decode("utf-8", errors="ignore")))
                        except Exception:
                            pass
                return "\n".join(chunks), None
        if ext in PPT_EXTS:
            import zipfile
            with zipfile.ZipFile(lp) as z:
                chunks = []
                for n in sorted(z.namelist()):
                    if n.startswith("ppt/slides/slide") and n.endswith(".xml"):
                        try:
                            chunks.append(_strip_xml(z.read(n).decode("utf-8", errors="ignore")))
                        except Exception:
                            pass
                return "\n".join(chunks), None
        if ext in PDF_EXTS:
            # 优先用 PyMuPDF（C 扩展，实测比 pypdf 快约 11 倍），失败回退 pypdf
            try:
                try:
                    import pymupdf as pdf_lib  # 新 API
                except Exception:
                    import fitz as pdf_lib     # 旧 API
                doc = pdf_lib.open(lp)
                chunks = []
                for i in range(min(len(doc), 300)):   # 最多 300 页
                    try:
                        t = doc[i].get_text()
                    except Exception:
                        t = ""
                    chunks.append(t)
                    if stop_text and any(k in t.lower() for k in stop_text):
                        break   # 命中即停，无需解析剩余页
                doc.close()
                return "\n".join(chunks), None
            except Exception:
                pass
            try:
                import logging
                logging.getLogger("pypdf").setLevel(logging.ERROR)
                from pypdf import PdfReader
                reader = PdfReader(lp)
                chunks = []
                for pg in reader.pages[:200]:   # 最多 200 页，防止超大 PDF 卡死
                    try:
                        t = pg.extract_text() or ""
                    except Exception:
                        t = ""
                    chunks.append(t)
                    if stop_text and any(k in t.lower() for k in stop_text):
                        break   # 命中即停，无需解析剩余页
                return "\n".join(chunks), None
            except Exception:
                return "", "未安装PDF解析库"
        if ext in OLE2_WORD_EXTS:
            # Word 97-2003 二进制 .doc / .wps：用 OLE2 分片表精确提取，
            # 失败再回退到旧的字节捞取法（兼容非标准/OLE 损坏文件）
            text, err = _extract_ole2_word_cached(lp)
            if text.strip():
                return text, err
            with open(lp, "rb") as fh:
                raw = fh.read(MAX_CONTENT_BYTES)
            fallback = _decode_bytes(raw)
            chunks = re.findall(r"[\u4e00-\u9fffA-Za-z0-9，。、；：？！\s]{4,}", fallback)
            return "\n".join(chunks), (err or "分片表提取失败，已回退为字节捞取，可能不完整")
        if ext in OLD_DOC_EXTS:
            # 其他老式二进制 Office（xls/ppt/et/dps）：尽力从字节流里捞可读文本
            with open(lp, "rb") as fh:
                raw = fh.read(MAX_CONTENT_BYTES)
            text = _decode_bytes(raw)
            # 只保留连续可读片段，避免二进制噪声
            chunks = re.findall(r"[\u4e00-\u9fffA-Za-z0-9，。、；：？！\s]{4,}", text)
            return "\n".join(chunks), "老式格式仅尽力提取，可能不完整"
        return "", "不支持的正文格式"
    except Exception as e:
        return "", "解析失败：" + str(e)


# ----------------------------- 文件预览 -----------------------------

# 图片：可直接按原始字节交给浏览器显示
IMAGE_EXTS = {"png", "jpg", "jpeg", "gif", "bmp", "webp", "svg", "ico", "tif", "tiff"}
IMAGE_MIME = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
              "gif": "image/gif", "bmp": "image/bmp", "webp": "image/webp",
              "svg": "image/svg+xml", "ico": "image/x-icon", "tif": "image/tiff",
              "tiff": "image/tiff"}
# 预览时能提取出文字的类型（纯文本 + 新版 Office + 老式 Office）
PREVIEW_TEXT_EXTS = TEXT_EXTS | DOC_EXTS | XLS_EXTS | PPT_EXTS | OLD_DOC_EXTS
MAX_PREVIEW_TEXT = 300000        # 文本预览最多 30 万字符，超出标记截断
MAX_PREVIEW_IMAGE = 120 * 1024 * 1024   # 图片超过 120MB 不直接内联
PREVIEW_PDF_ZOOM = 1.4           # PDF 渲染缩放，兼顾清晰度与速度
# 高保真 Office 渲染：
# 1) 优先调用本机 Microsoft Office / WPS 的原版引擎导出 PDF，能覆盖 doc/docx/xls/xlsx/ppt/pptx/wps/et/dps 等；
# 2) 未安装、COM 调用失败或被禁用时，新版格式用纯 Python 回退，老式格式降级为文字预览。
RICH_OFFICE_EXTS = {"docx", "pptx", "xlsx"}   # 纯 Python 能高保真渲染排版的新版格式
OFFICE_FAMILY = {
    "doc": "word", "docx": "word", "docm": "word", "dot": "word", "dotx": "word",
    "wps": "word", "rtf": "word", "odt": "word",
    "xls": "excel", "xlsx": "excel", "xlsm": "excel", "xlsb": "excel",
    "et": "excel", "ods": "excel",
    "ppt": "powerpoint", "pptx": "powerpoint", "pps": "powerpoint",
    "ppsx": "powerpoint", "dps": "powerpoint", "odp": "powerpoint",
}
COM_OFFICE_EXTS = set(OFFICE_FAMILY)
COM_PROGIDS = {
    # 顺序很重要：优先 Microsoft Office；只装 WPS 时自动落到 WPS 的 ProgID
    "word": [("Microsoft Word", "Word.Application"),
             ("WPS 文字", "KWPS.Application"),
             ("WPS 文字", "wps.Application")],
    "excel": [("Microsoft Excel", "Excel.Application"),
              ("WPS 表格", "KET.Application"),
              ("WPS 表格", "et.Application")],
    "powerpoint": [("Microsoft PowerPoint", "PowerPoint.Application"),
                   ("WPS 演示", "KWPP.Application"),
                   ("WPS 演示", "wpp.Application")],
}
# 环境变量 FS_OFFICE_COM=0 可在极端情况下强制关闭 COM 调用，不影响纯 Python 回退
OFFICE_COM_DISABLED = os.environ.get("FS_OFFICE_COM", "").strip().lower() in {"0", "false", "no", "off"}
MAX_OFFICE_FILE = 60 * 1024 * 1024   # 单文件超过 60MB 不做排版渲染（防止内存爆炸）
_OFFICE_CACHE = {}                   # 渲染缓存：key=(path,mtime,size) -> (html,pages,err)
_OFFICE_CACHE_LOCK = threading.Lock()
_OFFICE_CACHE_MAX = 12
_OFFICE_PDF_CACHE = {}               # COM 转换出的临时 PDF：key -> (pdf_path, engine)
_OFFICE_PDF_CACHE_MAX = 8
_OFFICE_PDF_LOCK = threading.Lock()


def _office_libs():
    """惰性加载 Office 渲染库，返回 (mammoth, pptx, openpyxl) 或 None。"""
    try:
        import mammoth
        from pptx import Presentation
        from pptx.util import Emu
        import openpyxl
        return mammoth, Presentation, openpyxl
    except Exception:
        return None


def _com_client():
    """惰性加载 Windows COM 客户端：优先 pywin32，失败回退 comtypes。"""
    if os.name != "nt":
        return None, "当前不是 Windows 系统"
    errors = []
    try:
        import pythoncom
        import win32com.client as win32com_client
        return {"kind": "pywin32", "pythoncom": pythoncom, "client": win32com_client}, None
    except Exception as e:
        errors.append("pywin32: " + str(e))
    try:
        import comtypes.client as comtypes_client
        return {"kind": "comtypes", "client": comtypes_client}, None
    except Exception as e:
        errors.append("comtypes: " + str(e))
    return None, "缺少 COM 组件（" + "; ".join(errors) + "）"


def _progid_registered(progid):
    """先在注册表里确认 ProgID 存在，避免无谓地启动 COM 报错。"""
    try:
        import winreg
        for view in (getattr(winreg, "KEY_WOW64_64KEY", 0), getattr(winreg, "KEY_WOW64_32KEY", 0)):
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid, 0, winreg.KEY_READ | view):
                    return True
            except Exception:
                continue
        return False
    except Exception:
        return False


def _progid_engine(progid, family, default_label):
    """读取 CLSID 的 LocalServer32，识别 ProgID 实际指向 Microsoft 还是 WPS。"""
    wps_names = {"word": "WPS 文字", "excel": "WPS 表格", "powerpoint": "WPS 演示"}
    ms_names = {"word": "Microsoft Word", "excel": "Microsoft Excel", "powerpoint": "Microsoft PowerPoint"}
    try:
        import winreg
        views = (getattr(winreg, "KEY_WOW64_64KEY", 0), getattr(winreg, "KEY_WOW64_32KEY", 0))
        for view in views:
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid + "\\CLSID", 0,
                                    winreg.KEY_READ | view) as key:
                    clsid = winreg.QueryValueEx(key, "")[0]
            except Exception:
                continue
            for cmd_view in views:
                try:
                    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "CLSID\\" + clsid + "\\LocalServer32",
                                        0, winreg.KEY_READ | cmd_view) as key:
                        cmd = str(winreg.QueryValueEx(key, "")[0])
                    low = cmd.lower()
                    if ("kingsoft" in low or "wps" in low or "et.exe" in low
                            or "wpp.exe" in low or "dps.exe" in low):
                        return wps_names.get(family, default_label)
                    if ("microsoft" in low or "winword.exe" in low or "excel.exe" in low
                            or "powerpnt.exe" in low or "\\office16\\" in low or "\\office15\\" in low):
                        return ms_names.get(family, default_label)
                    return default_label
                except Exception:
                    continue
    except Exception:
        pass
    return default_label


def _create_office_app(family, clients):
    """按 Microsoft Office → WPS 的顺序创建一个不可见的办公自动化实例。"""
    kind = clients.get("kind")
    for label, progid in COM_PROGIDS.get(family, []):
        if not _progid_registered(progid):
            continue
        engine = _progid_engine(progid, family, label)
        try:
            if kind == "pywin32":
                app = clients["client"].DispatchEx(progid)
            else:
                app = clients["client"].CreateObject(progid, dynamic=True)
        except Exception:
            continue
        # 尽量关掉所有会弹窗/更新链接的行为，防止无人值守时卡住
        for attr, value in (("Visible", False), ("DisplayAlerts", False),
                            ("AskToUpdateLinks", False), ("AutomationSecurity", 3)):
            try:
                setattr(app, attr, value)
            except Exception:
                pass
        return app, engine
    return None, None


def _office_to_pdf(path, ext, family):
    """调用本机 Microsoft Office/WPS 把文档导出为临时 PDF。返回 (pdf_path, engine, err)。"""
    if OFFICE_COM_DISABLED:
        return None, None, "已禁用本机 Office/WPS 引擎"
    clients, err = _com_client()
    if not clients:
        return None, None, err
    pythoncom = clients.get("pythoncom") if clients.get("kind") == "pywin32" else None
    fd, pdf_path = tempfile.mkstemp(prefix="fspro_office_", suffix=".pdf")
    os.close(fd)
    try:
        os.remove(pdf_path)
    except Exception:
        pass

    app = None
    engine = None
    ok = False
    try:
        if pythoncom is not None:
            pythoncom.CoInitialize()
        app, engine = _create_office_app(family, clients)
        if app is None:
            return None, None, "未检测到可用的 Microsoft Office 或 WPS"
        abs_path = os.path.abspath(path)

        if family == "word":
            doc = None
            try:
                try:
                    doc = app.Documents.Open(abs_path, ConfirmConversions=False, ReadOnly=True,
                                             AddToRecentFiles=False, Revert=False)
                except Exception:
                    doc = app.Documents.Open(abs_path)
                try:
                    doc.ExportAsFixedFormat(pdf_path, 17)   # wdFormatPDF
                except Exception:
                    doc.SaveAs(pdf_path, 17)
            finally:
                if doc is not None:
                    try:
                        doc.Close(SaveChanges=0)
                    except Exception:
                        pass
        elif family == "excel":
            wb = None
            try:
                try:
                    wb = app.Workbooks.Open(abs_path, UpdateLinks=0, ReadOnly=True,
                                            AddToRecentFiles=False, IgnoreReadOnlyRecommended=True)
                except Exception:
                    wb = app.Workbooks.Open(abs_path)
                try:
                    wb.ExportAsFixedFormat(0, pdf_path)     # xlTypePDF
                except Exception:
                    wb.SaveAs(pdf_path, 57)                 # xlOpenXML? 失败则交给后续解析判断
            finally:
                if wb is not None:
                    try:
                        wb.Close(SaveChanges=False)
                    except Exception:
                        pass
        else:
            prs = None
            try:
                try:
                    prs = app.Presentations.Open(abs_path, WithWindow=False, ReadOnly=True)
                except Exception:
                    prs = app.Presentations.Open(abs_path)
                try:
                    prs.ExportAsFixedFormat(pdf_path, 2)    # ppFixedFormatTypePDF
                except Exception:
                    prs.SaveAs(pdf_path, 32)                # ppSaveAsPDF
            finally:
                if prs is not None:
                    try:
                        prs.Close()
                    except Exception:
                        pass

        if not os.path.exists(pdf_path) or os.path.getsize(pdf_path) < 100:
            return None, engine, "Office/WPS 未生成可用的预览文件"
        if pdf_page_count(pdf_path) <= 0:
            return None, engine, "Office/WPS 导出的预览无法解析"
        ok = True
        return pdf_path, engine, None
    except Exception as e:
        return None, engine, "Office/WPS 预览失败：" + str(e)
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        if pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        if not ok:
            try:
                os.remove(pdf_path)
            except Exception:
                pass


def _get_office_pdf(path, ext):
    """源文件未变化时复用已转换的 PDF，避免每次翻页都启动 Office/WPS。"""
    try:
        st = os.stat(to_long_path(path))
        key = (path, st.st_mtime, st.st_size, ext)
    except Exception:
        key = (path, ext)
    with _OFFICE_PDF_LOCK:
        cached = _OFFICE_PDF_CACHE.get(key)
    if cached and os.path.exists(cached[0]):
        return cached[0], cached[1], None

    family = OFFICE_FAMILY.get(ext)
    pdf_path, engine, err = _office_to_pdf(path, ext, family)
    if pdf_path:
        with _OFFICE_PDF_LOCK:
            if key in _OFFICE_PDF_CACHE:
                old_pdf = _OFFICE_PDF_CACHE[key][0]
                if old_pdf != pdf_path:
                    try:
                        os.remove(old_pdf)
                    except Exception:
                        pass
            _OFFICE_PDF_CACHE[key] = (pdf_path, engine)
            while len(_OFFICE_PDF_CACHE) > _OFFICE_PDF_CACHE_MAX:
                old_key = next(iter(_OFFICE_PDF_CACHE))
                old_pdf, _eng = _OFFICE_PDF_CACHE.pop(old_key)
                try:
                    os.remove(old_pdf)
                except Exception:
                    pass
    return pdf_path, engine, err


def _office_query_terms(query):
    """从搜索框表达式里取用于高亮的正向关键词，忽略 -排除 与 | 分隔符。"""
    terms = []
    for part in re.split(r"[\s|]+", query or ""):
        t = part.strip().strip('"').strip("'")
        if not t or t.startswith("-") or t.startswith("!"):
            continue
        terms.append(t)
        if len(terms) >= 12:
            break
    return terms


def _render_office_pdf(path, ext, page, query):
    """用本机 Office/WPS 导出的 PDF 渲染成 HTML 图片页，并按搜索词叠加高亮。"""
    family = OFFICE_FAMILY.get(ext)
    if not family:
        return None, 0, None, "不支持的格式"
    pdf_path, engine, err = _get_office_pdf(path, ext)
    if err or not pdf_path:
        return None, 0, engine, err or "本机 Office/WPS 转换失败"
    lib = _pymupdf()
    if lib is None:
        return None, 0, engine, "缺少 PDF 渲染组件"
    try:
        doc = lib.open(pdf_path)
        try:
            total = doc.page_count
            if total <= 0:
                return None, 0, engine, "预览文档无页面"
            page = max(1, min(page, total))
            pdf_page = doc.load_page(page - 1)
            rect = pdf_page.rect
            png, png_err = render_pdf_page(pdf_path, page, PREVIEW_PDF_ZOOM)
            if png is None:
                return None, total, engine, png_err or "PDF 页面渲染失败"

            overlay = []
            for term in _office_query_terms(query):
                try:
                    found = pdf_page.search_for(term)
                except Exception:
                    continue
                for r in found or []:
                    x0 = max(0.0, min(float(r.x0), rect.width))
                    y0 = max(0.0, min(float(r.y0), rect.height))
                    x1 = max(x0, min(float(r.x1), rect.width))
                    y1 = max(y0, min(float(r.y1), rect.height))
                    if x1 <= x0 or y1 <= y0:
                        continue
                    overlay.append(
                        '<div class="pv-hit" style="left:%.3f%%;top:%.3f%%;width:%.3f%%;height:%.3f%%"></div>'
                        % (x0 / rect.width * 100, y0 / rect.height * 100,
                           (x1 - x0) / rect.width * 100, (y1 - y0) / rect.height * 100)
                    )
            body = ('<div class="pv-pdf-page"><img class="pv-page-img" src="'
                    + _base64_data_uri(png, "image/png") + '">'
                    + "".join(overlay) + "</div>")
            return body, total, engine, None
        finally:
            doc.close()
    except Exception as e:
        return None, 0, engine, "原排版预览失败：" + str(e)


def _base64_data_uri(blob, mime):
    return "data:" + (mime or "application/octet-stream") + ";base64," + \
        base64.b64encode(blob).decode("ascii")


def _css_color(color, default=""):
    """从 python-pptx 的颜色对象取 CSS 颜色字符串。"""
    if color is None:
        return default
    try:
        ctype = color.type
        if ctype is None:
            return default
        if str(ctype).find("RGB") >= 0 or hasattr(color, "rgb"):
            rgb = color.rgb
            if rgb is not None:
                return "#" + str(rgb)
        if hasattr(color, "theme_color") and color.theme_color is not None:
            return ""
    except Exception:
        pass
    return default


def _xml_escape(s):
    return html.escape(str(s if s is not None else ""), quote=False)


def _render_docx(path):
    """用 mammoth 把 docx 转成保留标题/表格/列表/图片的 HTML。"""
    libs = _office_libs()
    if not libs:
        return None, 1, "缺少 Office 渲染组件"
    mammoth = libs[0]

    def img_handler(image):
        try:
            return {"src": _base64_data_uri(image.image.blob, image.content_type)}
        except Exception:
            return {"src": ""}

    try:
        with open(to_long_path(path), "rb") as fh:
            result = mammoth.convert_to_html(fh, convert_image=mammoth.images.img_element(img_handler))
        body = result.value or ""
        if not body.strip():
            return None, 1, "文档中未提取到可见内容"
        return body, 1, None
    except Exception as e:
        return None, 1, "Word 排版渲染失败：" + str(e)


def _pptx_shape_html(shape, slide_w, slide_h):
    """把 PPTX 单个形状渲染成绝对定位的 HTML（尽量还原位置、大小、文字、图片、表格）。"""
    try:
        L = float(shape.left or 0) / slide_w * 100
        T = float(shape.top or 0) / slide_h * 100
        W = float(shape.width or 0) / slide_w * 100
        H = float(shape.height or 0) / slide_h * 100
        if W <= 0 or H <= 0:
            return ""
        pos = "position:absolute;left:%.3f%%;top:%.3f%%;width:%.3f%%;height:%.3f%%;" % (L, T, W, H)
        # 图片
        if getattr(shape, "shape_type", None) is not None and str(shape.shape_type).find("PICTURE") >= 0:
            try:
                blob = shape.image.blob
                mime = shape.image.content_type
                return '<img src="' + _base64_data_uri(blob, mime) + '" style="' + pos + 'object-fit:fill">'
            except Exception:
                return ""
        # 表格
        if getattr(shape, "has_table", False):
            tbl = shape.table
            rows = []
            for r in tbl.rows:
                cells = []
                for c in r.cells:
                    cells.append("<td>" + _xml_escape(c.text) + "</td>")
                rows.append("<tr>" + "".join(cells) + "</tr>")
            return '<div style="' + pos + 'overflow:auto"><table class="pv-ppt-table">' + "".join(rows) + "</table></div>"
        # 文本框 / 自选图形
        if getattr(shape, "has_text_frame", False):
            fill = ""
            try:
                fc = _css_color(shape.fill.fore_color)
                if fc:
                    fill = "background:" + fc + ";"
            except Exception:
                pass
            lines = []
            tf = shape.text_frame
            try:
                anchor = str(getattr(tf, "vertical_anchor", "") or "")
                if "MIDDLE" in anchor:
                    fill += "display:flex;flex-direction:column;justify-content:center;"
                elif "BOTTOM" in anchor:
                    fill += "display:flex;flex-direction:column;justify-content:flex-end;"
            except Exception:
                pass
            for p in tf.paragraphs:
                align = ""
                try:
                    a = str(p.alignment or "")
                    if "CENTER" in a:
                        align = "text-align:center;"
                    elif "RIGHT" in a:
                        align = "text-align:right;"
                    elif "JUSTIFY" in a:
                        align = "text-align:justify;"
                except Exception:
                    pass
                runs = []
                for run in p.runs:
                    st = ""
                    try:
                        if run.font.size is not None:
                            st += "font-size:%.1fpt;" % (run.font.size.pt)
                    except Exception:
                        pass
                    try:
                        if run.font.bold:
                            st += "font-weight:bold;"
                    except Exception:
                        pass
                    try:
                        if run.font.italic:
                            st += "font-style:italic;"
                    except Exception:
                        pass
                    try:
                        if run.font.underline:
                            st += "text-decoration:underline;"
                    except Exception:
                        pass
                    fc = _css_color(getattr(run.font, "color", None))
                    if fc:
                        st += "color:" + fc + ";"
                    try:
                        if run.font.name:
                            st += "font-family:'" + str(run.font.name).replace("'", "") + "',sans-serif;"
                    except Exception:
                        pass
                    txt = _xml_escape(run.text)
                    runs.append("<span" + (" style='" + st + "'" if st else "") + ">" + txt + "</span>")
                para_text = "".join(runs) if runs else _xml_escape(p.text)
                lines.append("<div style='min-height:1em;" + align + "'>" + para_text + "</div>")
            return '<div style="' + pos + fill + 'overflow:hidden;padding:2px">' + "".join(lines) + "</div>"
        # 其他形状：画个空框，避免空白
        return '<div style="' + pos + 'border:1px dashed #cbd5e1;opacity:.35"></div>'
    except Exception:
        return ""


def _render_pptx(path, page):
    """用 python-pptx 逐页渲染，返回 (html, total_pages, err)。"""
    libs = _office_libs()
    if not libs:
        return None, 0, "缺少 Office 渲染组件"
    Presentation = libs[1]
    try:
        prs = Presentation(to_long_path(path))
        total = len(prs.slides)
        if total <= 0:
            return None, 0, "演示文稿无幻灯片"
        page = max(1, min(page, total))
        slide = prs.slides[page - 1]
        sw = float(prs.slide_width or 12192000)
        sh = float(prs.slide_height or 6858000)
        parts = []
        for shape in slide.shapes:
            parts.append(_pptx_shape_html(shape, sw, sh))
        return '<div class="pv-ppt-slide">' + "".join(parts) + "</div>", total, None
    except Exception as e:
        return None, 0, "PPT 排版渲染失败：" + str(e)


def _render_xlsx(path, page):
    """用 openpyxl 渲染 xlsx 工作表为带样式的表格，返回 (html, total_sheets, err)。"""
    libs = _office_libs()
    if not libs:
        return None, 0, "缺少 Office 渲染组件"
    openpyxl = libs[2]
    try:
        wb = openpyxl.load_workbook(to_long_path(path), read_only=True, data_only=True)
        try:
            sheets = list(wb.sheetnames)
            if not sheets:
                return None, 0, "工作簿无工作表"
            page = max(1, min(page, len(sheets)))
            ws = wb[sheets[page - 1]]
            rows = []
            max_row = min(ws.max_row or 0, 2000)   # 防止超大数据卡死
            max_col = min(ws.max_column or 0, 120)
            for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
                cells = []
                for c in row:
                    v = c.value
                    if v is None:
                        cells.append("<td class='pv-x-empty'></td>")
                        continue
                    txt = _xml_escape(v)
                    st = ""
                    try:
                        if c.font and c.font.bold:
                            st += "font-weight:bold;"
                    except Exception:
                        pass
                    try:
                        if c.font and c.font.sz:
                            st += "font-size:%.0fpt;" % float(c.font.sz)
                    except Exception:
                        pass
                    try:
                        if c.font and c.font.color and c.font.color.rgb:
                            st += "color:#" + str(c.font.color.rgb)[-6:] + ";"
                    except Exception:
                        pass
                    try:
                        if c.fill and c.fill.fgColor and c.fill.fgColor.rgb:
                            rgb = str(c.fill.fgColor.rgb)
                            if rgb and rgb.upper() != "00000000":
                                st += "background:#" + rgb[-6:] + ";"
                    except Exception:
                        pass
                    try:
                        al = c.alignment
                        if al and al.horizontal == "center":
                            st += "text-align:center;"
                        elif al and al.horizontal == "right":
                            st += "text-align:right;"
                    except Exception:
                        pass
                    cells.append("<td" + (" style='" + st + "'" if st else "") + ">" + txt + "</td>")
                rows.append("<tr>" + "".join(cells) + "</tr>")
            head = "<div class='pv-sheet-name'>工作表：" + _xml_escape(sheets[page - 1]) + "</div>"
            note = "<div class='pv-note'>仅展示前 %d 行 × %d 列（只读预览）</div>" % (max_row, max_col)
            return head + '<table class="pv-x-table">' + "".join(rows) + "</table>" + note, len(sheets), None
        finally:
            wb.close()
    except Exception as e:
        return None, 0, "Excel 排版渲染失败：" + str(e)


def render_office_rich(path, ext, page=1, query=""):
    """Office 高保真渲染入口：优先本机 Office/WPS，失败再走纯 Python 回退。"""
    try:
        st = os.stat(to_long_path(path))
        key = (path, st.st_mtime, st.st_size, ext)
        size = st.st_size
    except Exception:
        key = (path, ext)
        size = 0

    com_err = None
    if ext in COM_OFFICE_EXTS and size <= MAX_OFFICE_FILE:
        body, pages, engine, com_err = _render_office_pdf(path, ext, page, query)
        if body is not None:
            doc = _wrap_office_html(body, pages, page, "pdf", query, path, engine=engine)
            doc["engine"] = engine
            return doc

    if size > MAX_OFFICE_FILE:
        return {"ok": False, "previewable": False,
                "reason": "文件超过 %dMB，已跳过排版预览" % (MAX_OFFICE_FILE // 1024 // 1024)}
    if ext not in RICH_OFFICE_EXTS:
        return {"ok": False, "previewable": False,
                "reason": com_err or "未检测到可用的 Microsoft Office/WPS，无法还原该格式的原排版"}

    with _OFFICE_CACHE_LOCK:
        cached = _OFFICE_CACHE.get(key)
    if cached and page <= cached[1]:
        body, pages, err = cached[0].get(page), cached[1], cached[2]
        if body is not None:
            doc = _wrap_office_html(body, pages, page, ext, query, path)
            doc["engine"] = "内置渲染"
            return doc

    if ext == "docx":
        body, pages, err = _render_docx(path)
    elif ext == "pptx":
        body, pages, err = _render_pptx(path, page)
    elif ext == "xlsx":
        body, pages, err = _render_xlsx(path, page)
    else:
        body, pages, err = None, 0, "不支持的格式"

    if body is None:
        return {"ok": False, "previewable": False, "reason": err or "排版渲染失败"}

    with _OFFICE_CACHE_LOCK:
        entry = _OFFICE_CACHE.get(key)
        if not entry:
            entry = ({}, pages, err)
            if len(_OFFICE_CACHE) >= _OFFICE_CACHE_MAX:
                try:
                    _OFFICE_CACHE.pop(next(iter(_OFFICE_CACHE)))
                except Exception:
                    pass
            _OFFICE_CACHE[key] = entry
        entry[0][page] = body
    doc = _wrap_office_html(body, pages, page, ext, query, path)
    doc["engine"] = "内置渲染"
    return doc


def _wrap_office_html(body, pages, page, ext, query, path, engine=None):
    """把渲染出的正文包成一个自带样式的完整 HTML 文档（含搜索词高亮脚本）。"""
    doc = """<!DOCTYPE html><html><head><meta charset="utf-8"><style>
*{box-sizing:border-box}html,body{margin:0;padding:0;background:#eef2f7;
font-family:"Microsoft YaHei","Segoe UI",Arial,sans-serif;color:#1f2937;font-size:13px}
.pv-paper{background:#fff;margin:14px auto;padding:22px 26px;max-width:100%;
box-shadow:0 1px 4px rgba(0,0,0,.12);min-height:96%;position:relative}
.pv-engine{color:#64748b;font-size:11px;margin:-10px 0 12px;text-align:right}
.pv-ppt-wrap{width:100%;padding-top:56.25%;position:relative;background:#000}
.pv-ppt-slide{position:absolute;inset:0;background:#fff;overflow:hidden}
.pv-ppt-table{border-collapse:collapse;font-size:12px}
.pv-ppt-table td{border:1px solid #94a3b8;padding:2px 5px}
.pv-x-table{border-collapse:collapse;width:100%;font-size:12px;background:#fff}
.pv-x-table td{border:1px solid #d7dee8;padding:3px 6px;min-width:54px}
.pv-pdf-page{position:relative;width:100%;line-height:0;background:#fff}
.pv-page-img{width:100%;height:auto;display:block}
.pv-hit{position:absolute;background:rgba(253,230,138,.55);border:1px solid rgba(180,83,9,.35);
border-radius:2px;pointer-events:none}
.pv-x-empty{background:#fafbfc}.pv-sheet-name{font-weight:600;color:#2563eb;margin-bottom:8px}
.pv-note{color:#94a3b8;font-size:11px;margin-top:8px}
body img{max-width:100%}
h1,h2,h3,h4{color:#111827;line-height:1.3}
table{border-collapse:collapse}td,th{border:1px solid #d7dee8;padding:3px 6px}
mark{background:#fde68a;color:#b91c1c;padding:0 1px;border-radius:2px}
</style></head><body><div class="pv-paper">__BODY__</div>
<script>
(function(){var q=__QUERY__||'';if(!q)return;
var terms=q.split(/[\\s|]+/).filter(function(x){return x&&x.charAt(0)!=='-'&&x.charAt(0)!=='!';});
if(!terms.length)return;
var walker=document.createTreeWalker(document.body,NodeFilter.SHOW_TEXT,null,false);
var nodes=[];while(walker.nextNode())nodes.push(walker.currentNode);
nodes.forEach(function(node){var t=node.nodeValue;if(!t)return;var low=t.toLowerCase(),r=[];
terms.forEach(function(k){var kk=k.toLowerCase(),from=0,idx;
while((idx=low.indexOf(kk,from))>=0){r.push([idx,idx+k.length]);from=idx+k.length;if(r.length>1000)break;}});
if(!r.length)return;r.sort(function(a,b){return a[0]-b[0];});
var merged=[];r.forEach(function(x){if(merged.length&&x[0]<=merged[merged.length-1][1])
merged[merged.length-1][1]=Math.max(merged[merged.length-1][1],x[1]);else merged.push([x[0],x[1]]);});
var frag=document.createDocumentFragment(),pos=0;
merged.forEach(function(x){frag.appendChild(document.createTextNode(t.slice(pos,x[0])));
var m=document.createElement('mark');m.textContent=t.slice(x[0],x[1]);frag.appendChild(m);pos=x[1];});
frag.appendChild(document.createTextNode(t.slice(pos)));
node.parentNode.replaceChild(frag,node);});
})();
</script></body></html>"""
    # PPTX 外层加 16:9 画布容器；COM 导出的 PDF 页保持原始比例
    if ext == "pptx":
        body = '<div class="pv-ppt-wrap">' + body + "</div>"
    if engine:
        body = '<div class="pv-engine">预览引擎：' + _xml_escape(engine) + "</div>" + body
    query_js = json.dumps(query or "", ensure_ascii=False)
    full = doc.replace("__BODY__", body).replace("__QUERY__", query_js)
    return {"ok": True, "html": full, "pages": pages, "page": page}


def _pymupdf():
    """返回可用的 PyMuPDF 模块（新名 pymupdf / 旧名 fitz），失败返回 None。"""
    try:
        import pymupdf as pdf_lib
        return pdf_lib
    except Exception:
        try:
            import fitz as pdf_lib
            return pdf_lib
        except Exception:
            return None


def pdf_page_count(path):
    lib = _pymupdf()
    if lib is None:
        return -1
    try:
        doc = lib.open(to_long_path(path))
        n = doc.page_count
        doc.close()
        return n
    except Exception:
        return -1


def render_pdf_page(path, page_no, zoom=PREVIEW_PDF_ZOOM):
    """把 PDF 指定页（从 1 开始）渲染成 PNG 字节，失败返回 (None, 错误信息)。"""
    lib = _pymupdf()
    if lib is None:
        return None, "未安装 PDF 渲染组件"
    try:
        doc = lib.open(to_long_path(path))
        try:
            if page_no < 1 or page_no > doc.page_count:
                return None, f"页码超出范围（共 {doc.page_count} 页）"
            pix = doc[page_no - 1].get_pixmap(matrix=lib.Matrix(zoom, zoom), alpha=False)
            return pix.tobytes("png"), None
        finally:
            doc.close()
    except Exception as e:
        return None, "PDF 渲染失败：" + str(e)


def preview_info(path, page=1, query=""):
    """判定文件能否预览并返回预览元信息。完全只读。"""
    ext = os.path.splitext(path)[1].lower().lstrip(".")
    name = os.path.basename(path)
    info = {"name": name, "ext": ext, "path": path, "ok": False,
            "reason": "此文件类型不支持预览", "previewable": False}
    try:
        lp = to_long_path(path)
        st = os.stat(lp)
        if not os.path.isfile(lp):
            info["reason"] = "路径不是文件"
            return info
        info["size"] = st.st_size
        info["mtime"] = time.strftime("%Y-%m-%d %H:%M", time.localtime(st.st_mtime))
    except Exception as e:
        info["reason"] = "无法读取文件：" + str(e)
        return info

    if ext in IMAGE_EXTS:
        if st.st_size > MAX_PREVIEW_IMAGE:
            info["reason"] = f"图片超过 {MAX_PREVIEW_IMAGE // 1024 // 1024}MB，已跳过预览"
            return info
        info.update({"ok": True, "previewable": True, "kind": "image",
                     "mime": IMAGE_MIME.get(ext, "application/octet-stream"),
                     "url": "/api/preview/blob?p=" + quote(path), "pages": 1, "page": 1})
        return info

    if ext in PDF_EXTS:
        pages = pdf_page_count(path)
        if pages <= 0:
            info["reason"] = "无法打开 PDF（或缺少 PDF 渲染组件）"
            return info
        page = max(1, min(page, pages))
        info.update({"ok": True, "previewable": True, "kind": "pdf",
                     "mime": "image/png", "pages": pages, "page": page,
                     "url": f"/api/preview/blob?p={quote(path)}&page={page}"})
        return info

    # Office/WPS：优先原版引擎导出 PDF（保留原排版），失败后新版格式走内置渲染、老式格式走文字
    office_err = None
    if ext in COM_OFFICE_EXTS:
        doc = render_office_rich(path, ext, page, query)
        if doc.get("ok") and doc.get("html"):
            html_b64 = base64.b64encode(doc["html"].encode("utf-8")).decode("ascii")
            engine = doc.get("engine")
            note = ("本机 " + engine + " 原版引擎导出，保留原排版") if engine else "保留原排版"
            if not engine or engine == "内置渲染":
                note = ("内置渲染保留原排版的图文混排"
                        if ext != "xlsx" else "内置渲染按工作表展示单元格样式与内容")
            info.update({"ok": True, "previewable": True, "kind": "office",
                         "mime": "text/html", "html_b64": html_b64,
                         "pages": doc.get("pages", 1), "page": doc.get("page", page),
                         "layout": True, "engine": engine, "note": note})
            return info
        office_err = doc.get("reason") or "Office/WPS 原排版预览失败"
        if ext in RICH_OFFICE_EXTS:
            info["reason"] = office_err
            return info
        # 老式 doc/xls/ppt/wps/et/dps 继续向下走纯文字回退

    if ext in PREVIEW_TEXT_EXTS:
        text, err = extract_text(path, ext)
        truncated = len(text) > MAX_PREVIEW_TEXT
        if truncated:
            text = text[:MAX_PREVIEW_TEXT]
        if ext in OLE2_WORD_EXTS and not err:
            note = "Word 97-2003 正文已精确提取（分片表解析）；排版/图片不保留"
            if office_err:
                note = office_err + "；" + note
        elif office_err:
            note = office_err + "；已降级为文字预览"
        else:
            note = err or ("未检测到可用 Office/WPS，老式格式仅支持文字预览"
                           if ext in OLD_DOC_EXTS else "")
        info.update({"ok": True, "previewable": True, "kind": "text",
                     "text": text, "truncated": truncated, "note": note})
        return info

    return info


def preview_blob(path, page=1):
    """返回预览用的原始字节 (bytes, mime, error)。"""
    ext = os.path.splitext(path)[1].lower().lstrip(".")
    if ext in PDF_EXTS:
        data, err = render_pdf_page(path, max(1, page))
        return data, "image/png", err
    if ext in IMAGE_EXTS:
        try:
            with open(to_long_path(path), "rb") as fh:
                return fh.read(), IMAGE_MIME.get(ext, "application/octet-stream"), None
        except Exception as e:
            return None, "application/octet-stream", "读取图片失败：" + str(e)
    return None, "application/octet-stream", "此文件类型不支持预览"


def _content_hits(text, groups, mode, compiled, query, scope):
    """在正文中匹配，返回 (总评分, [(关键词, 位置)...])。"""
    if not text:
        return 0, []
    low = text.lower()
    all_hits = []
    total = 0
    if mode in ("regex", "wildcard"):
        if mode == "regex" and compiled:
            try:
                for m in compiled.finditer(text):
                    all_hits.append((m.group(0) or query, m.start()))
                    total += 70
                    if len(all_hits) >= 30:
                        break
            except Exception:
                pass
        elif mode == "wildcard":
            pat = query.lower().replace("*", ".*").replace("?", ".")
            try:
                for m in re.finditer(pat, text, re.I):
                    all_hits.append((m.group(0), m.start()))
                    total += 70
                    if len(all_hits) >= 30:
                        break
            except Exception:
                pass
        return total, all_hits
    for and_terms, excl in groups:
        ok, gs, local = True, 0, []
        for term in and_terms:
            s = _term_score({"_rel_l": low, "_name_l": low, "_stem_l": low},
                            term, "stem", fuzzy=(mode == "smart"))
            if s <= 0:
                # 正文里直接用子串定位
                pos = low.find(term)
                if pos < 0:
                    ok = False
                    break
                s = 60
            gs += s
            start = 0
            while True:
                pos = low.find(term, start)
                if pos < 0:
                    break
                local.append((term, pos))
                start = pos + len(term)
                if len(local) >= 30:
                    break
        if not ok:
            continue
        bad = False
        for term in excl:
            if low.find(term) >= 0:
                bad = True
                break
        if bad:
            continue
        total = max(total, gs)
        all_hits.extend(local)
    all_hits.sort(key=lambda x: x[1])
    return total, all_hits[:40]


def _record_content_hit(job, rec, text, err, groups, compiled):
    """把一个文件的提取结果做关键词匹配并记录命中（供串行/并行共用）。"""
    job.scanned += 1
    if text:
        score, hits = _content_hits(text, groups, job.mode, compiled, job.query, "stem")
        if hits:
            first_pos = hits[0][1]
            ctx = 55
            a = max(0, first_pos - ctx)
            b = min(len(text), first_pos + ctx)
            snippet = ("…" if a > 0 else "") + text[a:b].replace("\n", " ") + ("…" if b < len(text) else "")
            row = {"name": rec["name"], "path": rec["path"], "rel": rec["rel"],
                   "ext": rec["ext"], "category": rec["category"], "size": rec["size"],
                   "mtime": rec["mtime"], "snippet": snippet,
                   "hit_count": len(hits), "score": score,
                   "terms": sorted({h[0] for h in hits if h[0]})}
            job.hits.append(row)
    if err and len(job.errors) < 200:
        job.errors.append(rec["name"] + "：" + err)
    if job.scanned % 20 == 0 or job.scanned == job.total:
        job.status_text = f"已读取 {job.scanned}/{job.total} 个文件，命中 {len(job.hits)} 个"


def _silence_mupdf():
    """静默 PyMuPDF 对损坏 PDF 的告警（会写到 fd 2，Python 日志拦不住）。"""
    try:
        import pymupdf as _m
    except Exception:
        try:
            import fitz as _m
        except Exception:
            return
    try:
        _m.TOOLS.mupdf_display_errors(False)
    except Exception:
        pass


def _content_extract_worker(args):
    """进程池 worker：提取单个文件正文，返回 (path, text, err)。
    必须是模块顶层函数，Windows spawn 子进程才能 pickle。"""
    path, ext, needle, stop_text = args
    _silence_mupdf()
    try:
        text, err = extract_text(path, ext, needle, stop_text)
    except Exception as e:
        text, err = "", "解析失败：" + str(e)
    return path, text, err


class ContentJob:
    """后台正文搜索任务。"""

    def __init__(self, task_id, params):
        self.id = uuid.uuid4().hex[:12]
        self.task_id = task_id
        self.params = params
        self.query = (params.get("query") or "").strip()
        self.mode = params.get("mode", "smart")
        self.status = "running"
        self.status_text = "准备正文搜索..."
        self.scanned = 0
        self.total = 0
        self.hits = []
        self.errors = []
        self.start = time.time()
        self.end = None
        self.cancel = False

    def snapshot(self, offset=0, limit=200):
        page = self.hits[offset:offset + limit]
        return {"id": self.id, "status": self.status, "status_text": self.status_text,
                "scanned": self.scanned, "total": self.total,
                "total_hits": len(self.hits), "rows": page,
                "errors": len(self.errors),
                "elapsed": round((self.end or time.time()) - self.start, 2),
                "query": self.query}


def content_worker(job, task):
    """遍历文件、提取正文、匹配关键词，产出带片段的命中结果。"""
    _silence_mupdf()
    params = job.params
    f = params.get("filters") or {}
    # 只扫描可能含正文的文件；尊重后缀/分类筛选
    candidates = [r for r in task.files if r["ext"] in CONTENT_EXTS]
    if f.get("skip_slow"):
        slow = PDF_EXTS | OLD_DOC_EXTS
        candidates = [r for r in candidates if r["ext"] not in slow]
    if (f.get("ext") or "").strip():
        want = {x.strip().lstrip(".") for x in f["ext"].replace(",", " ").split() if x.strip()}
        candidates = [r for r in candidates if r["ext"] in want]
    if (f.get("category") or "").strip():
        candidates = [r for r in candidates if r["category"] == f["category"].strip()]
    smin = parse_size(f.get("size_min"))
    smax = parse_size(f.get("size_max"))
    if smin is not None:
        candidates = [r for r in candidates if r["size"] >= smin]
    if smax is not None:
        candidates = [r for r in candidates if r["size"] <= smax]
    pc = (f.get("path_contains") or "").strip().lower()
    if pc:
        candidates = [r for r in candidates if pc in r["rel"].lower()]

    job.total = len(candidates)
    job.status_text = f"正在读取正文（共 {job.total} 个候选文件）..."

    groups = parse_query(job.query, job.mode, "stem")
    compiled = None
    if job.mode == "regex":
        try:
            compiled = re.compile(job.query, re.I)
        except Exception:
            compiled = None
    if not groups and job.mode not in ("regex", "wildcard"):
        job.status = "done"
        job.status_text = "请输入关键词"
        job.end = time.time()
        return

    # 构造关键词的多编码字节（预筛）和小写词（逐页提前停止）
    needle = None
    stop_text = None
    if job.mode in ("smart", "exact"):
        terms = [t for grp in groups for t in grp[0] if t]
        stop_text = [t.lower() for t in terms if t]
        long_terms = [t for t in terms if len(t) >= 2]
        if long_terms:
            needle = set()
            for t in long_terms:
                for enc in ("utf-8", "gb18030", "utf-16-le"):
                    try:
                        b = t.encode(enc)
                        if b:
                            needle.add(b)
                    except Exception:
                        pass

    # 按路径建索引，供并行结果回填
    by_path = {r["path"]: r for r in candidates}
    # 缓存命中立即处理，未命中的送进程池并行提取
    todo = []
    for rec in candidates:
        key = rec["path"]
        if key in task.content_cache:
            text, err = task.content_cache[key]
            _record_content_hit(job, rec, text, err, groups, compiled)
            if job.cancel:
                job.status = "cancelled"; job.end = time.time(); return
        else:
            todo.append(rec)

    if todo:
        # pypdf 为纯 Python，GIL 下线程无效，必须用进程池利用多核。
        # 优先 PyMuPDF（若已装）：C 扩展释放 GIL，进程池同样受益。
        workers = min(8, (os.cpu_count() or 4))
        ex = None
        use_pool = len(todo) >= 4
        if use_pool:
            try:
                ctx = multiprocessing.get_context("spawn")
                ex = ProcessPoolExecutor(max_workers=workers, mp_context=ctx)
                futures = {}
                for rec in todo:
                    fut = ex.submit(_content_extract_worker,
                                    (rec["path"], rec["ext"], needle, stop_text))
                    futures[fut] = rec["path"]
                for fut in as_completed(futures):
                    if job.cancel:
                        job.status = "cancelled"; job.end = time.time()
                        for f in futures:
                            f.cancel()
                        ex.shutdown(wait=False, cancel_futures=True)
                        return
                    path = futures[fut]
                    try:
                        path, text, err = fut.result()
                    except Exception as e:
                        text, err = "", "解析失败：" + str(e)
                    task.content_cache[path] = (text, err)
                    _record_content_hit(job, by_path[path], text, err, groups, compiled)
            except Exception:
                # 进程池不可用时回退串行
                if ex:
                    try:
                        ex.shutdown(wait=False, cancel_futures=True)
                    except Exception:
                        pass
                for rec in todo:
                    if job.cancel:
                        job.status = "cancelled"; job.end = time.time(); return
                    key = rec["path"]
                    text, err = extract_text(rec["path"], rec["ext"], needle, stop_text)
                    task.content_cache[key] = (text, err)
                    _record_content_hit(job, rec, text, err, groups, compiled)
            finally:
                if ex:
                    try:
                        ex.shutdown(wait=False, cancel_futures=True)
                    except Exception:
                        pass
        else:
            for rec in todo:
                if job.cancel:
                    job.status = "cancelled"; job.end = time.time(); return
                key = rec["path"]
                text, err = extract_text(rec["path"], rec["ext"], needle, stop_text)
                task.content_cache[key] = (text, err)
                _record_content_hit(job, rec, text, err, groups, compiled)
    # 按评分排序
    job.hits.sort(key=lambda x: (-x.get("score", 0), x["name"].lower()))
    job.status = "done"
    job.end = time.time()
    job.status_text = f"正文搜索完成：{job.total} 个文件中命中 {len(job.hits)} 个"


# ----------------------------- HTTP 服务 -----------------------------

class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        pass

    def _send(self, code, body, ctype="application/json; charset=utf-8", headers=None):
        if isinstance(body, (dict, list)):
            body = json.dumps(body, ensure_ascii=False).encode("utf-8")
        elif isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (headers or {}).items():
            self.send_header(k, v)
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
        q = parse_qs(u.query)
        if u.path == "/":
            self._send(200, HTML_PAGE, "text/html; charset=utf-8")
        elif u.path == "/api/drives":
            self._send(200, {"drives": list_drives()})
        elif u.path == "/api/quickpath":
            name = q.get("name", [""])[0]
            home = os.path.expanduser("~")
            mapping = {"Desktop": ["Desktop", "OneDrive/Desktop", "桌面"],
                       "Documents": ["Documents", "My Documents", "文档"],
                       "Downloads": ["Downloads", "下载"]}
            path = None
            for cand in mapping.get(name, []):
                p = os.path.join(home, cand)
                if os.path.isdir(p):
                    path = p
                    break
            self._send(200, {"path": path})
        elif u.path == "/api/browse":
            self._send(200, browse_directory(q.get("path", [""])[0]))
        elif u.path == "/api/browse/search":
            self._send(200, search_browse_dirs(
                q.get("path", [""])[0], q.get("kw", [""])[0]))
        elif u.path == "/api/open":
            # 供导出的目录树 HTML 调用：点击文件名用系统默认程序打开。
            # 允许跨域，这样从 file:// 打开的导出页也能调用本地接口。
            try:
                p = q.get("p", [""])[0]
                open_file_external(p)
                self._send(200, {"ok": True}, headers={"Access-Control-Allow-Origin": "*"})
            except Exception as e:
                self._send(200, {"ok": False, "error": str(e)},
                           headers={"Access-Control-Allow-Origin": "*"})
        elif u.path == "/api/preview":
            # 预览元信息：判定能否预览，文本类直接返回文字，图片/PDF 返回 blob URL
            try:
                page = max(1, int(q.get("page", ["1"])[0]))
            except Exception:
                page = 1
            query = q.get("query", [""])[0]
            self._send(200, preview_info(q.get("p", [""])[0], page, query))
        elif u.path == "/api/preview/blob":
            # 预览字节：图片原样返回，PDF 渲染指定页为 PNG
            try:
                page = max(1, int(q.get("page", ["1"])[0]))
            except Exception:
                page = 1
            data, mime, err = preview_blob(q.get("p", [""])[0], page)
            if err:
                self._send(200, {"error": err})
            else:
                self._send(200, data, mime or "application/octet-stream")
        elif u.path == "/api/progress":
            task = TASKS.get(q.get("id", [""])[0])
            self._send(200, task.snapshot() if task else {"error": "任务不存在，请重新扫描"})
        elif u.path == "/api/dup/status":
            job = DUP_JOBS.get(q.get("id", [""])[0])
            self._send(200, job.snapshot() if job else {"error": "任务不存在"})
        elif u.path == "/api/content/status":
            job = CONTENT_JOBS.get(q.get("id", [""])[0])
            if not job:
                self._send(200, {"error": "任务不存在"})
                return
            try:
                offset = max(0, int(q.get("offset", ["0"])[0]))
            except Exception:
                offset = 0
            try:
                limit = min(1000, max(1, int(q.get("limit", ["100"])[0])))
            except Exception:
                limit = 100
            self._send(200, job.snapshot(offset=offset, limit=limit))
        elif u.path == "/api/content/all":
            # 正文搜索：跨页全选，只回传路径，避免大结果集传输开销
            job = CONTENT_JOBS.get(q.get("id", [""])[0])
            if not job:
                self._send(200, {"error": "任务不存在"})
                return
            self._send(200, {"total": len(job.hits),
                             "paths": [r["path"] for r in job.hits]})
        elif u.path == "/api/reveal":
            p = q.get("p", [""])[0]
            try:
                if os.name == "nt":
                    subprocess_open = 'explorer'
                    import subprocess
                    subprocess.Popen([subprocess_open, "/select,", os.path.normpath(p)])
                else:
                    import subprocess
                    subprocess.Popen(["open", os.path.dirname(p)])
                self._send(200, {"ok": True})
            except Exception as e:
                self._send(200, {"error": str(e)})
        elif u.path == "/api/data":
            self._send(200, DATA)
        elif u.path == "/api/paste/status":
            job = PASTE_JOBS.get(q.get("id", [""])[0])
            self._send(200, job.snapshot() if job else {"error": "任务不存在"})
        elif u.path == "/api/clip":
            info = dict(CLIPBOARD)
            info["count"] = len(resolve_clip_files())
            self._send(200, info)
        elif u.path == "/api/type-files":
            # 功能1：列出某个后缀类型的全部文件（供勾选/复制/导出）
            task = TASKS.get(q.get("id", [""])[0])
            if not task:
                self._send(200, {"error": "请先扫描"})
                return
            ext = (q.get("ext", [""])[0] or "").lower().lstrip(".")
            files = list(task.files) if ext in ("", "__all__") else [f for f in task.files if f["ext"] == ext]
            files = sorted(files, key=lambda r: r["path"].lower())
            self._send(200, {"ext": ext, "total": len(files), "rows": files})
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
            # 保留最近 5 个任务
            if len(TASKS) > 5:
                for k in list(TASKS.keys())[:-5]:
                    TASKS.pop(k, None)
            threading.Thread(target=scan_worker, args=(task,), daemon=True).start()
            self._send(200, {"id": task.id, "path": task.path})

        elif u.path == "/api/cancel":
            task = TASKS.get(data.get("id", ""))
            if task:
                task.cancel = True
            self._send(200, {"ok": True})

        elif u.path == "/api/search":
            task = TASKS.get(data.get("id", ""))
            if not task:
                self._send(200, {"error": "请先扫描"})
                return
            total, rows, took, terms = search_files(task, data)
            # 记录历史
            q = (data.get("query") or "").strip()
            if q:
                DATA["history"] = [q] + [x for x in DATA["history"] if x != q]
                DATA["history"] = DATA["history"][:30]
                save_data()
            self._send(200, {"total": total, "rows": rows, "took": took, "terms": terms})

        elif u.path == "/api/search/paths":
            # 标题搜索：跨页全选，只回传路径
            task = TASKS.get(data.get("id", ""))
            if not task:
                self._send(200, {"error": "请先扫描"})
                return
            params = dict(data)
            params["offset"] = 0
            params["limit"] = 1000000
            params["_unlimited"] = True
            total, rows, _took, _terms = search_files(task, params)
            self._send(200, {"total": total, "paths": [r["path"] for r in rows]})

        elif u.path == "/api/export":
            task = TASKS.get(data.get("id", ""))
            if not task:
                self._send(200, {"error": "未找到扫描结果"})
                return
            # 取全部结果（不分页）：功能1类型列表 source=type / 正文搜索 source=content
            source = data.get("source", "")
            query_label = data.get("query", "")
            if source == "content":
                job = CONTENT_JOBS.get(data.get("content_id", ""))
                if not job:
                    self._send(200, {"error": "正文搜索任务已失效，请重新搜索"})
                    return
                rows = list(job.hits)
                if data.get("paths_only"):
                    wanted = set(data["paths_only"])
                    rows = [f for f in rows if f["path"] in wanted]
                scope_label = task.path + "  ·  正文关键词：" + (job.query or "(空)")
                query_label = job.query or data.get("query", "")
            elif source == "type":
                ext = (data.get("ext") or "").lower().lstrip(".")
                rows = list(task.files) if ext in ("", "__all__") else [f for f in task.files if f["ext"] == ext]
                rows = sorted(rows, key=lambda r: r["path"].lower())
                if data.get("paths_only"):
                    wanted = set(data["paths_only"])
                    rows = [f for f in rows if f["path"] in wanted]
                scope_label = task.path + (f"  ·  类型 .{ext}" if ext else "")
            else:
                params = dict(data)
                params["offset"] = 0
                params["limit"] = 1000000
                params["_unlimited"] = True
                total, rows, _took, _terms = search_files(task, params)
                scope_label = task.path
            if not rows:
                self._send(200, {"error": "没有可导出的结果"})
                return
            fmt = (data.get("fmt") or "html").lower()
            dest = (data.get("output_dir") or "").strip().strip('"')
            if not dest or not os.path.isdir(dest):
                self._send(200, {"error": "输出文件夹不存在：" + dest})
                return
            stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            if fmt == "tree":
                out = os.path.join(dest, f"目录树_{stamp}.html")
            else:
                out = os.path.join(dest, f"搜索结果_{stamp}.{fmt}")
            try:
                meta = {"scope": scope_label, "query": query_label}
                if fmt == "tree":
                    n = export_tree(rows, meta, out)
                else:
                    n = export_rows(rows, meta, fmt, out)
                self._send(200, {"ok": True, "path": out, "rows": n, "size": os.path.getsize(out)})
            except Exception as e:
                self._send(200, {"error": "导出失败：" + str(e)})

        elif u.path == "/api/clip/set":
            # 设置剪贴板：paths=[...] 或 task_id + ext（整类）
            mode = data.get("mode", "copy")
            if mode not in ("copy", "cut"):
                mode = "copy"
            CLIPBOARD.update({
                "mode": mode,
                "paths": data.get("paths") or [],
                "task_id": data.get("task_id"),
                "ext": data.get("ext"),
                "source": data.get("source"),
            })
            n = len(resolve_clip_files())
            info = dict(CLIPBOARD); info["count"] = n
            self._send(200, {"ok": True, "count": n, "clip": info})

        elif u.path == "/api/clip/clear":
            CLIPBOARD.update({"mode": None, "paths": [], "task_id": None,
                              "ext": None, "source": None})
            self._send(200, {"ok": True})

        elif u.path == "/api/sysclip/set":
            # 写入 Windows 系统剪贴板：可在程序外用 Ctrl+V 粘贴 / 移动
            try:
                n = set_system_clipboard(data.get("paths") or [],
                                         cut=(data.get("mode") == "cut"))
                self._send(200, {"ok": True, "count": n})
            except Exception as e:
                self._send(200, {"error": str(e)})

        elif u.path == "/api/paste":
            files = resolve_clip_files()
            if not files:
                self._send(200, {"error": "剪贴板为空"})
                return
            dest = (data.get("dest") or "").strip().strip('"')
            if not dest:
                self._send(200, {"error": "请选择目标文件夹"})
                return
            job = PasteJob(dest, data.get("collision", "skip"), files,
                           CLIPBOARD.get("mode") or "copy")
            PASTE_JOBS[job.id] = job
            if len(PASTE_JOBS) > 8:
                for k in list(PASTE_JOBS.keys())[:-8]:
                    PASTE_JOBS.pop(k, None)
            threading.Thread(target=paste_worker, args=(job,), daemon=True).start()
            self._send(200, {"id": job.id, "total": len(files)})

        elif u.path == "/api/paste/cancel":
            job = PASTE_JOBS.get(data.get("id", ""))
            if job:
                job.cancel = True
            self._send(200, {"ok": True})

        elif u.path == "/api/content/start":
            task = TASKS.get(data.get("id", ""))
            if not task:
                self._send(200, {"error": "请先扫描"})
                return
            q = (data.get("query") or "").strip()
            if not q:
                self._send(200, {"error": "请输入正文关键词"})
                return
            job = ContentJob(task.id, data)
            CONTENT_JOBS[job.id] = job
            if len(CONTENT_JOBS) > 8:
                for k in list(CONTENT_JOBS.keys())[:-8]:
                    CONTENT_JOBS.pop(k, None)
            threading.Thread(target=content_worker, args=(job, task), daemon=True).start()
            self._send(200, {"id": job.id})
        elif u.path == "/api/content/cancel":
            job = CONTENT_JOBS.get(data.get("id", ""))
            if job:
                job.cancel = True
            self._send(200, {"ok": True})
        elif u.path == "/api/dup/start":
            task = TASKS.get(data.get("id", ""))
            if not task:
                self._send(200, {"error": "请先扫描"})
                return
            try:
                min_size = parse_size(data.get("min_size")) or 1024 * 1024
            except Exception:
                min_size = 1024 * 1024
            job = DupJob(task.id, min_size)
            DUP_JOBS[job.id] = job
            if len(DUP_JOBS) > 3:
                for k in list(DUP_JOBS.keys())[:-3]:
                    DUP_JOBS.pop(k, None)
            threading.Thread(target=dup_worker, args=(job,), daemon=True).start()
            self._send(200, {"id": job.id})

        elif u.path == "/api/dup/cancel":
            job = DUP_JOBS.get(data.get("id", ""))
            if job:
                job.cancel = True
            self._send(200, {"ok": True})

        elif u.path == "/api/saved/add":
            item = {"name": (data.get("name") or "").strip(), "params": data.get("params") or {}}
            if not item["name"]:
                self._send(200, {"error": "请填写名称"})
                return
            DATA["saved"] = [x for x in DATA["saved"] if x.get("name") != item["name"]]
            DATA["saved"].insert(0, item)
            DATA["saved"] = DATA["saved"][:50]
            save_data()
            self._send(200, {"ok": True})

        elif u.path == "/api/saved/delete":
            name = data.get("name", "")
            DATA["saved"] = [x for x in DATA["saved"] if x.get("name") != name]
            save_data()
            self._send(200, {"ok": True})

        elif u.path == "/api/data/clear_history":
            DATA["history"] = []
            save_data()
            self._send(200, {"ok": True})

        elif u.path == "/api/quit":
            self._send(200, {"ok": True})

            def _shutdown():
                time.sleep(0.4)
                if SERVER is not None:
                    SERVER.shutdown()
            threading.Thread(target=_shutdown, daemon=True).start()
        else:
            self._send(404, {"error": "not found"})


# ----------------------------- 前端页面 -----------------------------

HTML_PAGE = r"""<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>文件搜索工具（标题 + 正文）</title>
<style>
:root{--bg:#f5f7fa;--card:#fff;--line:#e5e7eb;--text:#1f2937;--muted:#6b7280;
--blue:#2563eb;--blue-l:#eff6ff;--green:#059669;--amber:#b45309;--red:#dc2626;--red-l:#fef2f2;}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--text);
font-family:"Microsoft YaHei",-apple-system,system-ui,sans-serif;font-size:14px;line-height:1.5}
/* 左右分栏骨架：左侧全部功能，右侧文件预览 */
.app{display:flex;height:100vh;overflow:hidden}
.left{flex:1;min-width:360px;height:100vh;overflow:auto;display:flex;flex-direction:column}
.wrap{width:100%;max-width:1280px;margin:0 auto;padding:22px 18px 60px;flex:1}
.splitter{width:9px;flex:none;cursor:col-resize;background:#eef1f5;border-left:1px solid var(--line);
  position:relative;transition:background .15s}
.splitter:hover,.splitter.drag{background:#bfdbfe}
.splitter::after{content:'⋮';position:absolute;left:0;top:50%;transform:translateY(-50%);
  width:100%;text-align:center;color:#94a3b8;font-size:14px;line-height:1;pointer-events:none}
.preview{width:var(--pv-w,500px);flex:none;height:100vh;background:#fff;border-left:1px solid var(--line);
  display:flex;flex-direction:column;box-shadow:-2px 0 10px rgba(15,23,42,.04)}
.app.no-preview .preview,.app.no-preview .splitter{display:none}
.pv-head{padding:12px 14px;border-bottom:1px solid var(--line);display:flex;align-items:center;gap:8px;
  flex-wrap:wrap;background:#f8fafc;flex:none}
.pv-head b{font-size:14px}
.pv-head .pv-name{font-size:12px;color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:100%}
.pv-tools{margin-left:auto;display:flex;gap:6px;align-items:center;flex:none}
.pv-tools button{padding:5px 9px;font-size:12px}
.pv-close{font-size:17px;line-height:1;padding:3px 9px!important}
.pv-bar{padding:7px 12px;border-bottom:1px solid var(--line);background:#fff;display:flex;
  align-items:center;gap:8px;font-size:12px;color:var(--muted);flex:none}
.pv-bar button{padding:4px 9px;font-size:12px}
.pv-body{flex:1;overflow:auto;padding:14px;background:#fafbfc;display:flex;flex-direction:column}
.pv-empty{margin:auto;text-align:center;color:var(--muted);font-size:13px;line-height:2;max-width:300px}
.pv-empty .ico{font-size:40px;display:block;margin-bottom:8px;opacity:.6}
.pv-imgwrap{margin:auto;text-align:center;max-width:100%}
.pv-imgwrap img{max-width:100%;max-height:calc(100vh - 190px);object-fit:contain;
  border:1px solid var(--line);border-radius:8px;background:#fff;box-shadow:0 2px 12px rgba(0,0,0,.06)}
.pv-iframe{flex:1;width:100%;border:1px solid var(--line);border-radius:8px;background:#fff;
  box-shadow:0 2px 12px rgba(0,0,0,.06);min-height:300px}
.pv-text{margin:0;white-space:pre-wrap;word-break:break-word;font-size:13px;line-height:1.75;
  font-family:Consolas,"Microsoft YaHei",monospace;background:#fff;border:1px solid var(--line);
  border-radius:8px;padding:14px;color:#1f2937;flex:1}
.pv-text mark{background:#fee2e2;color:var(--red);font-weight:700}
.pv-note{color:var(--amber);font-size:12px;margin-bottom:8px;background:#fffbeb;
  border:1px solid #fde68a;border-radius:7px;padding:7px 10px}
.pv-meta{font-size:12px;color:var(--muted);margin-top:10px;line-height:1.9;word-break:break-all}
.pv-unsupported{margin:auto;text-align:center;color:var(--muted);line-height:2}
.pv-unsupported .ico{font-size:42px;display:block;margin-bottom:10px}
.pv-open-btn{margin-top:12px}
.pv-float{position:fixed;right:16px;top:16px;z-index:40;display:none;
  background:#fff;border:1px solid var(--line);border-radius:9px;padding:8px 13px;
  font-size:13px;box-shadow:0 2px 10px rgba(0,0,0,.1);cursor:pointer}
.pv-float:hover{background:var(--blue-l);color:var(--blue);border-color:#bfdbfe}
.op2.pv{color:#7c3aed}
.c-name.pv-click{cursor:pointer}
.c-name.pv-click:hover{color:var(--blue);text-decoration:underline}
.type-tbl .fn.pv-click{cursor:pointer}
.type-tbl .fn.pv-click:hover{color:var(--blue);text-decoration:underline}
h1{font-size:22px;margin:0 0 2px}h2{font-size:16px;margin:0 0 14px}
.sub{color:var(--muted);font-size:13px;margin-bottom:18px}
.panel{background:var(--card);border:1px solid var(--line);border-radius:12px;
padding:18px 20px;margin-bottom:16px;box-shadow:0 1px 3px rgba(0,0,0,.03)}
.input-row{display:flex;gap:8px;flex-wrap:wrap;align-items:center}
input[type=text],input[type=number],input[type=date],select{
border:1px solid #d1d5db;border-radius:8px;padding:8px 10px;font-size:14px;background:#fff;color:var(--text)}
.pick-input{flex:1;min-width:240px;background:#fafafa}
button{border:0;border-radius:8px;padding:9px 15px;font-size:14px;cursor:pointer;font-family:inherit}
button:disabled{opacity:.45;cursor:not-allowed}
.btn-primary{background:var(--blue);color:#fff;font-weight:600}
.btn-primary:hover:not(:disabled){background:#1d4ed8}
.btn-red{background:var(--red);color:#fff;font-weight:600}
.btn-red:hover:not(:disabled){background:#b91c1c}
.btn-light{background:var(--blue-l);color:var(--blue)}
.btn-ghost{background:#f3f4f6;color:#374151}
.btn-amber{background:#fffbeb;color:var(--amber)}
.quick{display:flex;gap:6px;align-items:center;flex-wrap:wrap;margin-top:12px;color:var(--muted);font-size:12px}
.quick button{padding:5px 10px;font-size:12px;background:#f3f4f6;color:#374151}
.opts{display:flex;gap:18px;align-items:center;flex-wrap:wrap;margin-top:14px;
padding-top:14px;border-top:1px dashed var(--line)}
.check{display:inline-flex;gap:6px;align-items:center;cursor:pointer;color:#374151}
.opt{display:inline-flex;gap:8px;align-items:center;color:#374151}
.progress-wrap{display:none;margin-top:14px}
.bar{height:8px;background:#eef2ff;border-radius:6px;overflow:hidden}
.bar i{display:block;height:100%;width:30%;background:linear-gradient(90deg,#60a5fa,#2563eb);
animation:mv 1.2s ease-in-out infinite}
@keyframes mv{0%{width:25%}50%{width:70%}100%{width:25%}}
.status{margin-top:7px;font-size:13px;color:var(--muted)}
/* 功能2统一进度面板 */
.search-prog{margin-top:12px;padding:12px 14px;background:#f8fafc;border:1px solid var(--line);border-radius:9px}
.search-prog.sp-run{background:#eff6ff;border-color:#bfdbfe}
.search-prog.sp-done{background:#f0fdf4;border-color:#bbf7d0}
.search-prog.sp-err{background:#fef2f2;border-color:#fecaca}
.sp-head{display:flex;justify-content:space-between;align-items:center;gap:12px;font-size:13px;font-weight:700;color:#1e40af}
.search-prog.sp-done .sp-head{color:#15803d}
.search-prog.sp-err .sp-head{color:#b91c1c}
.sp-time{font-weight:400;font-size:12px;color:var(--muted);white-space:nowrap}
.sp-bar{height:10px;background:#e0e7ff;border-radius:6px;overflow:hidden;margin-top:8px}
.sp-bar i{display:block;height:100%;width:35%;border-radius:6px;background:linear-gradient(90deg,#60a5fa,#2563eb);transition:width .3s}
.sp-bar i.indet{animation:mv 1.2s ease-in-out infinite}
.search-prog.sp-done .sp-bar i{background:linear-gradient(90deg,#4ade80,#16a34a)}
.sp-detail{margin-top:7px;font-size:12px;color:var(--muted);line-height:1.6}
.sp-cancel{margin-top:8px}
/* 功能1扫描结果折叠面板 */
.fold-head{display:flex;justify-content:space-between;align-items:center;margin-top:14px;padding:9px 12px;
  background:#f8fafc;border:1px solid var(--line);border-radius:9px;cursor:pointer;user-select:none;
  font-size:13px;font-weight:600;color:#374151;transition:.15s}
.fold-head:hover{background:#f1f5f9;border-color:#bfdbfe}
.fold-btn{color:var(--blue);font-size:12px;font-weight:600;white-space:nowrap}
.fold-hint{font-weight:400;color:var(--muted);font-size:12px;margin-left:6px}
.fold-body{margin-top:10px}
.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-top:14px}
.stat{background:#f8fafc;border:1px solid var(--line);border-radius:9px;padding:10px 12px}
.stat .k{color:var(--muted);font-size:12px}.stat .v{font-size:18px;font-weight:700;margin-top:2px}
.note{color:var(--muted);font-size:12px}
/* 搜索类型切换 */
.tabs{display:flex;gap:8px;margin-bottom:14px}
.tab{flex:1;padding:12px;border:2px solid var(--line);border-radius:10px;cursor:pointer;
background:#fff;text-align:center;font-weight:600;color:#374151;transition:.15s}
.tab small{display:block;font-weight:400;color:var(--muted);font-size:12px;margin-top:2px}
.tab.on{border-color:var(--blue);background:var(--blue-l);color:var(--blue)}
.tab.on small{color:#1d4ed8}
.tab.ct.on{border-color:var(--red);background:var(--red-l);color:var(--red)}
.tab.ct.on small{color:#b91c1c}
/* 搜索区 */
.search-bar{display:flex;gap:8px;align-items:center}
.qbox{flex:1;position:relative}
.qinput{width:100%;font-size:17px;padding:12px 14px 12px 50px;border:2px solid #dbe3ee;border-radius:10px}
.qinput:focus{outline:0;border-color:var(--blue);box-shadow:0 0 0 3px rgba(37,99,235,.1)}
.qicon{position:absolute;left:17px;top:50%;transform:translateY(-50%);color:#9ca3af;font-size:16px}
.modes{display:flex;gap:5px;margin-top:10px;flex-wrap:wrap}
.chip{padding:5px 12px;border-radius:20px;background:#f3f4f6;color:#4b5563;font-size:12px;cursor:pointer;border:1px solid transparent}
.chip.on{background:var(--blue-l);color:var(--blue);border-color:#bfdbfe;font-weight:600}
.chip.ct.on{background:var(--red-l);color:var(--red);border-color:#fecaca}
.adv{margin-top:12px;padding-top:12px;border-top:1px dashed var(--line);display:none}
.adv.show{display:block}
.fgrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px}
.fitem label{display:block;font-size:12px;color:var(--muted);margin-bottom:4px}
.fitem input,.fitem select{width:100%}
.fitem .pair{display:flex;gap:6px}
.fitem .pair input{flex:1;min-width:0;width:auto}
.adv .fsel{margin-bottom:5px}
/* 高级筛选：选中"自定义"时高亮对应输入框 */
.fitem input.hl-input{border-color:var(--blue);box-shadow:0 0 0 2px rgba(37,99,235,.12)}
/* 结果区快捷筛选 */
.res-filters{display:flex;gap:10px;align-items:flex-end;flex-wrap:wrap;margin:8px 0 12px;
  padding:10px 12px;background:#f8fafc;border:1px solid var(--line);border-radius:9px}
.res-filters .rf{display:inline-flex;flex-direction:column;gap:4px}
.res-filters .rf label{font-size:11px;color:var(--muted)}
.res-filters select{padding:6px 8px;font-size:12px;min-width:104px}
.res-filters .rreset{padding:6px 10px;font-size:12px}
.res-filters .rtip{align-self:center;font-size:11px;color:var(--muted)}
/* 结果区"自定义"输入行 */
.res-filters .rcustom{display:flex;gap:5px;margin-top:2px}
.res-filters .rcustom input{padding:6px 8px;font-size:12px;min-width:76px;width:100%;border:1px solid var(--line);border-radius:7px}
.res-filters .rcustom input[type=date]{min-width:120px}
.res-filters .rcustom input:focus{outline:0;border-color:var(--blue);box-shadow:0 0 0 2px rgba(37,99,235,.1)}
.hint{margin-top:10px;font-size:12px;color:var(--muted);background:#f8fafc;
padding:8px 11px;border-radius:8px;line-height:1.8}
.hint code{background:#eef2ff;color:#3730a3;padding:1px 5px;border-radius:4px}
.hint.red{background:var(--red-l);color:#9a3412}
/* 分类展示 */
.cat-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:12px;margin-top:14px}
.cat-card{border:1px solid var(--line);border-radius:10px;padding:12px;background:#fcfdff}
.cat-card h4{margin:0 0 8px;font-size:13px;color:#374151;display:flex;justify-content:space-between}
.cat-card h4 span{color:var(--muted);font-weight:400}
.ext-chip{display:inline-flex;gap:6px;align-items:center;padding:4px 9px;margin:0 5px 5px 0;
border-radius:14px;background:var(--blue-l);color:var(--blue);font-size:12px;cursor:pointer;border:1px solid #dbeafe}
.ext-chip:hover{background:#dbeafe}
.ext-chip b{color:#1e40af}
/* 结果区 */
.res-head{display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:10px}
.res-info{color:var(--muted);font-size:13px}
.res-info b{color:var(--blue);font-size:16px}
.tbl-wrap{overflow:auto;max-height:560px;border:1px solid var(--line);border-radius:9px}
table{border-collapse:collapse;width:100%;font-size:13px}
th,td{padding:8px 10px;text-align:left;border-bottom:1px solid #f0f1f3;white-space:nowrap}
th{background:#f8fafc;position:sticky;top:0;z-index:2;color:#374151;cursor:pointer;user-select:none}
th:hover{background:#f1f5f9}
tr:hover td{background:#fafcff}
.c-name{font-weight:600;max-width:320px;overflow:hidden;text-overflow:ellipsis}
.c-rel{color:var(--muted);max-width:260px;overflow:hidden;text-overflow:ellipsis}
.c-snip{max-width:420px;white-space:normal;color:#4b5563;font-size:12px}
mark{background:#fee2e2;color:var(--red);font-weight:700;border-radius:2px;padding:0 2px}
.tag{display:inline-block;background:var(--blue-l);color:var(--blue);font-size:11px;
padding:1px 7px;border-radius:10px}
.op{color:var(--blue);cursor:pointer;font-size:12px}
.pager{display:flex;gap:8px;align-items:center;justify-content:center;margin-top:12px;flex-wrap:wrap}
.pager button{padding:6px 11px;font-size:13px}
.empty{text-align:center;color:var(--muted);padding:40px 0}
.content-prog{margin-bottom:12px;padding:10px 12px;background:var(--red-l);
border:1px solid #fecaca;border-radius:8px;color:#9a3412;font-size:13px;display:none}
/* 弹窗 */
.mask{display:none;position:fixed;inset:0;background:rgba(15,23,42,.45);z-index:50;
align-items:center;justify-content:center;padding:20px}
.mask.show{display:flex}
.modal{background:#fff;border-radius:14px;width:min(860px,96vw);max-height:88vh;overflow:hidden;
display:flex;flex-direction:column;box-shadow:0 20px 60px rgba(0,0,0,.2)}
.modal h3{margin:0;padding:16px 20px;border-bottom:1px solid var(--line);font-size:16px}
.modal .body{padding:16px 20px;overflow:auto;flex:1}
.modal .foot{padding:12px 20px;border-top:1px solid var(--line);display:flex;justify-content:flex-end;gap:8px}
.crumb{font-size:12px;color:var(--muted);margin-bottom:10px;word-break:break-all;background:#f8fafc;padding:8px 10px;border-radius:7px}
.dirs{display:grid;grid-template-columns:repeat(auto-fill,minmax(190px,1fr));gap:7px}
.dirs .d{padding:8px 10px;border:1px solid var(--line);border-radius:7px;cursor:pointer;
overflow:hidden;text-overflow:ellipsis;white-space:nowrap;background:#fff}
.dirs .d:hover{background:var(--blue-l);border-color:#bfdbfe;color:var(--blue)}
.bsearch{display:flex;gap:8px;align-items:center;margin-bottom:10px}
.bsearch input{flex:1;padding:8px 10px;border:1px solid var(--line);border-radius:7px;font-size:13px;min-width:0}
.bsearch span{font-size:12px;color:var(--muted);white-space:nowrap}
.bresults{max-height:300px;overflow:auto;border:1px solid #bfdbfe;border-radius:8px;margin-bottom:10px;background:#f8fbff}
.bresults .r{padding:7px 10px;border-bottom:1px solid #e8f0fe;cursor:pointer;font-size:13px}
.bresults .r:hover{background:#eaf2ff}
.bresults .r .p{color:var(--muted);font-size:11px;margin-left:8px;word-break:break-all}
.bresults .pick{float:right;color:var(--blue);font-size:12px}
.bresults .tip{padding:8px 10px;color:var(--muted);font-size:12px}
.toast{position:fixed;left:50%;bottom:28px;transform:translateX(-50%);z-index:99;
 background:#1e293b;color:#fff;padding:10px 18px;border-radius:8px;font-size:13px;
 box-shadow:0 8px 24px rgba(0,0,0,.2);opacity:0;pointer-events:none;transition:opacity .2s}
.toast.show{opacity:1}
.op2.sys{color:#0f766e}.op2.sys.cut{color:#b45309}
.saved-item{display:flex;justify-content:space-between;align-items:center;padding:8px 10px;
border:1px solid var(--line);border-radius:8px;margin-bottom:7px}
.tagchip{display:inline-block;background:#f3f4f6;color:#4b5563;font-size:11px;padding:1px 7px;border-radius:10px;margin:2px}
.footer{text-align:center;color:var(--muted);font-size:12px;margin-top:22px}
/* 类型文件列表面板（功能1）与复制粘贴 */
.type-panel{margin-top:14px;border:1px solid #bfdbfe;border-radius:10px;background:#f8fbff;display:none}
.type-panel.show{display:block}
.type-head{display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;
 padding:10px 13px;border-bottom:1px solid #dbeafe}
.type-head b{color:#1e40af;font-size:14px}
.clip-bar{display:flex;gap:7px;align-items:center;flex-wrap:wrap;padding:9px 13px;background:#fff;
 border-bottom:1px solid #dbeafe}
.clip-bar button{padding:6px 11px;font-size:12px}
.clip-state{font-size:12px;color:var(--muted);margin-left:auto}
.clip-state.has{color:#1d4ed8;font-weight:600}
.type-tbl-wrap{max-height:420px;overflow:auto}
.type-tbl{width:100%;border-collapse:collapse;font-size:12px}
.type-tbl th,.type-tbl td{padding:6px 9px;text-align:left;border-bottom:1px solid #eef2f7;white-space:nowrap}
.type-tbl th{background:#f1f5f9;position:sticky;top:0;color:#475569;z-index:2}
.type-tbl tr:hover td{background:#fafcff}
.type-tbl .ck{width:30px;text-align:center}
.type-tbl .fn{max-width:300px;overflow:hidden;text-overflow:ellipsis}
.type-tbl .rp{max-width:280px;overflow:hidden;text-overflow:ellipsis;color:var(--muted)}
.res-ck{width:28px;text-align:center}
.op2{color:var(--blue);cursor:pointer;font-size:12px;margin-right:8px}
.op2.cut{color:var(--amber)}
.op2:hover{text-decoration:underline}
.paste-prog{margin-top:10px}
.small-btn{padding:5px 10px;font-size:12px}
</style></head><body>
<div class="app" id="app">
<div class="left" id="leftPane">
<div class="wrap">
  <h1>🔎 文件搜索工具（标题 + 正文）</h1>
  <div class="sub">扫描硬盘/文件夹后，可按「文件名」或「文档正文」搜索，支持智能模糊、多关键词、筛选与高亮标红。全文只读，绝不改动原文件。</div>

  <!-- 功能1 -->
  <div class="panel">
    <h2>1. 扫描范围与文件类型概览</h2>
    <div class="input-row">
      <input type="text" id="scanPath" class="pick-input" readonly placeholder="点击“浏览选择”选取硬盘或文件夹">
      <button class="btn-light" id="btnBrowseScan">浏览选择…</button>
      <button class="btn-primary" id="btnScan">开始扫描</button>
      <button class="btn-ghost" id="btnCancel" disabled>取消</button>
    </div>
    <div class="quick" id="quick"><span>快捷位置：</span></div>
    <div class="opts">
      <label class="check"><input type="checkbox" id="followLinks" checked> 跟随目录联接/符号链接</label>
      <label class="check"><input type="checkbox" id="excHidden"> 排除隐藏文件
        <span style="color:var(--amber)">（数据盘很多文件带隐藏属性，勾选会漏扫）</span></label>
      <div class="opt"><label>最大层数（0=扫到最内层）</label>
        <input type="number" id="maxDepth" value="0" min="0" max="99" style="width:76px"></div>
    </div>
    <div class="progress-wrap" id="progressWrap">
      <div class="bar"><i></i></div>
      <div class="status" id="status">准备中...</div>
      <div class="status" id="levelInfo" style="color:var(--blue);font-weight:600"></div>
    </div>
    <div class="fold-head" id="scanFoldHead" title="点击展开或收起扫描结果">
      <span>📊 扫描结果概览 <span class="fold-hint" id="scanFoldHint">（完成扫描后自动展开，可点击收起）</span></span>
      <span class="fold-btn" id="scanFoldBtn">展开 ▼</span>
    </div>
    <div class="fold-body" id="scanResult" style="display:none">
      <div class="stats" id="stats" style="display:none">
        <div class="stat"><div class="k">文件总数</div><div class="v" id="sFiles">0</div></div>
        <div class="stat"><div class="k">文件夹数</div><div class="v" id="sDirs">0</div></div>
        <div class="stat"><div class="k">已扫层数 / 后缀种类</div><div class="v" id="sKinds">0 / 0</div></div>
        <div class="stat"><div class="k">扫描用时</div><div class="v" id="sTime">0s</div></div>
      </div>
      <div id="levelBars" style="display:none;margin-top:12px;padding:10px 12px;background:#eff6ff;
        border:1px solid #bfdbfe;border-radius:8px;font-size:12px;color:#1e40af"></div>
      <div id="skipNote" style="display:none;margin-top:12px;padding:10px 12px;background:#fff7ed;
        border:1px solid #fed7aa;border-radius:8px;font-size:12px;color:#9a3412"></div>
      <div class="cat-grid" id="catGrid" style="display:none"></div>
      <div class="note" id="catTip" style="display:none;margin-top:8px">💡 点击后缀标签可列出该类型全部文件（勾选后复制/剪切/导出）；也可以用该后缀跳转到功能2搜索。</div>
    </div>

    <!-- 功能1：类型文件列表（点后缀标签后展开，可勾选/复制/剪切/导出） -->
    <div class="type-panel" id="typePanel">
      <div class="type-head">
        <b id="typeTitle">类型文件列表</b>
        <div style="display:flex;gap:7px;flex-wrap:wrap">
          <button class="btn-ghost small-btn" id="btnTypeClose">收起列表</button>
        </div>
      </div>
      <div class="clip-bar">
        <button class="btn-light small-btn" id="btnTypeAll">全选</button>
        <button class="btn-ghost small-btn" id="btnTypeNone">取消全选</button>
        <button class="btn-primary small-btn" id="btnTypeCopy">复制选中</button>
        <button class="btn-amber small-btn" id="btnTypeCut">剪切选中</button>
        <button class="btn-light small-btn" id="btnTypeMove">移动到文件夹…</button>
        <button class="btn-light small-btn" id="btnTypePaste">粘贴到文件夹…</button>
        <button class="btn-ghost small-btn" id="btnTypeSysCopy" title="写入 Windows 系统剪贴板，可在资源管理器直接 Ctrl+V 粘贴">系统复制</button>
        <button class="btn-ghost small-btn" id="btnTypeSysCut" title="写入 Windows 系统剪贴板，在资源管理器 Ctrl+V 即移动">系统剪切</button>
        <button class="btn-light small-btn" id="btnTypeCopyAll">复制本类型全部</button>
        <button class="btn-light small-btn" id="btnTypeExport">导出目录…</button>
        <span class="clip-state" id="typeSelInfo">已选 0 个</span>
      </div>
      <div class="type-tbl-wrap">
        <table class="type-tbl"><thead><tr>
          <th class="ck">选</th><th>文件名</th><th>大小</th><th>修改时间</th><th>相对路径</th><th>操作</th>
        </tr></thead><tbody id="typeBody"></tbody></table>
        <div class="empty" id="typeEmpty" style="display:none">该类型没有文件</div>
      </div>
    </div>
  </div>

  <!-- 功能2 -->
  <div class="panel">
    <h2>2. 关键词搜索</h2>
    <div class="tabs" id="searchTabs">
      <div class="tab on" data-type="title">📝 标题搜索<small>按文件名查找（记不清格式也能找到）</small></div>
      <div class="tab ct" data-type="content">📄 正文搜索<small>读取文档内容查找（标题里没有也能找到）</small></div>
    </div>
    <div class="input-row" style="margin-bottom:12px">
      <input type="text" id="searchPath" class="pick-input" readonly placeholder="在此“浏览选择”要搜索的文件夹（可不预先扫描，搜索时自动扫描）">
      <button class="btn-light" id="btnBrowseSearch">浏览选择…</button>
      <button class="btn-ghost" id="btnUseScanPath">用功能1的文件夹</button>
    </div>
    <div class="search-bar">
      <div class="qbox"><span class="qicon">🔎</span>
        <input type="text" id="query" class="qinput" placeholder="输入关键词，例如：方沸石" autocomplete="off"></div>
      <button class="btn-red" id="btnSearch" style="padding:12px 22px">搜索</button>
      <button class="btn-ghost" id="btnClear" style="padding:12px 14px">清空</button>
    </div>
    <!-- 功能2统一进度面板：紧挨搜索按钮，实时反馈扫描/搜索/正文读取 -->
    <div class="search-prog sp-run" id="searchProg" style="display:none">
      <div class="sp-head">
        <span id="spPhase">准备中…</span>
        <span class="sp-time" id="spTime"></span>
      </div>
      <div class="sp-bar"><i id="spBar"></i></div>
      <div class="sp-detail" id="spDetail"></div>
      <button class="btn-ghost sp-cancel" id="spCancel" style="display:none">停止当前操作</button>
    </div>
    <div id="titleOpts">
      <div class="modes" id="modes">
        <span style="color:var(--muted);font-size:12px;align-self:center;margin-right:4px">匹配模式：</span>
        <span class="chip on" data-mode="smart">智能模糊（推荐）</span>
        <span class="chip" data-mode="exact">精确包含</span>
        <span class="chip" data-mode="wildcard">通配符 * ?</span>
        <span class="chip" data-mode="regex">正则表达式</span>
      </div>
      <div class="modes" id="scopes">
        <span style="color:var(--muted);font-size:12px;align-self:center;margin-right:4px">搜索范围：</span>
        <span class="chip on" data-scope="stem">仅文件名（忽略格式）</span>
        <span class="chip" data-scope="name_ext">文件名含后缀</span>
        <span class="chip" data-scope="path">完整路径</span>
      </div>
    </div>
    <div class="hint red" id="contentHint" style="display:none">
      📄 <b>正文搜索</b>会打开文档读取内容：支持 TXT / DOCX / XLSX / PPTX / PDF / DOC / XLS / PPT / CSV 等，
      标题中没有、正文里出现的关键词也能找到。命中结果会显示<b style="color:var(--red)">正文片段并标红</b>。<br>
      <label class="check" style="margin-top:6px"><input type="checkbox" id="skipSlow"> 只看纯文本与新版 Office（跳过 PDF 与老式 DOC/XLS/PPT）——已启用多核并行加速，一般无需勾选</label>
    </div>
    <div class="opts">
      <label class="check"><input type="checkbox" id="advToggle"> 高级筛选（大小 / 日期 / 后缀 / 分类 / 路径）</label>
      <div class="opt"><label>排序</label>
        <select id="sort">
          <option value="relevance">相关度</option><option value="name">文件名</option>
          <option value="path">路径</option><option value="size">大小（大到小）</option>
          <option value="mtime">修改时间（新到旧）</option>
        </select></div>
      <button class="btn-light" id="btnResearch">按当前条件重新搜索</button>
      <button class="btn-light" id="btnSave">保存此搜索</button>
      <button class="btn-ghost" id="btnSaved">已保存的搜索</button>
    </div>
    <div class="adv" id="adv">
      <div class="fgrid">
        <div class="fitem">
          <label>后缀格式（下拉选或手输）</label>
          <select class="fsel" id="fExtSel"></select>
          <input type="text" id="fExt" placeholder="留空=全部，可手输 pdf,docx">
        </div>
        <div class="fitem"><label>分类</label><select id="fCat"><option value="">全部分类</option></select></div>
        <div class="fitem">
          <label>文件大小（下拉选或手输）</label>
          <select class="fsel" id="fSizeSel"></select>
          <div class="pair">
            <input type="text" id="fMin" placeholder="最小 如 10MB">
            <input type="text" id="fMax" placeholder="最大 如 2GB">
          </div>
        </div>
        <div class="fitem">
          <label>修改日期（下拉选或手选）</label>
          <select class="fsel" id="fDateSel"></select>
          <div class="pair">
            <input type="date" id="fDateFrom" title="开始日期">
            <input type="date" id="fDateTo" title="结束日期">
          </div>
        </div>
        <div class="fitem"><label>路径包含</label><input type="text" id="fPath" placeholder="如 科研"></div>
      </div>
      <div class="note" style="margin-top:8px">💡 下拉选预设会自动填入输入框；选中“自定义”后对应输入框会<b style="color:var(--blue)">高亮并聚焦</b>，直接输入即可（多个后缀用英文逗号分隔；大小填最小/最大，日期选起止日）。改动自动重搜，并与结果区筛选双向同步。</div>
    </div>
    <div class="hint" id="titleHint">
      💡 <b>多关键词</b>：空格分隔表示“同时包含”；<code>|</code> 表示“或”；<code>-词</code> 或 <code>!词</code> 表示“排除”。
      例：<code>中考 -答案</code>、<code>报告 | 总结</code>。<br>
      <b>智能模糊</b>：按顺序命中即可，如 <code>中学生考点</code> 能匹配「中学生<b>考试</b>考点汇总.doc」——
      即使记错格式（以为 PDF 其实是 DOC）也能搜到。<br>
      <b>通配符</b>（匹配整个文件名）：<code>*</code> 任意多个字符，<code>?</code> 单个字符，如 <code>*报告*.doc</code>。
    </div>
    <div class="quick" id="history"></div>
  </div>

  <!-- 结果 -->
  <div class="panel">
    <div class="res-head">
      <div class="res-info" id="resInfo">尚未搜索</div>
      <div class="input-row">
        <button class="btn-ghost" id="btnContentCancel" style="display:none">停止正文搜索</button>
        <button class="btn-amber" id="btnDup">查找重复文件</button>
        <button class="btn-light" id="btnExport">导出结果…</button>
      </div>
    </div>
    <!-- 功能2结果区：复制/剪切/粘贴操作条 -->
    <div class="clip-bar" id="resClipBar">
      <button class="btn-ghost small-btn" id="btnResAll">全选本页</button>
      <button class="btn-ghost small-btn" id="btnResAllPages" title="选中当前搜索条件下的所有文件，不限当前页">全选全部结果</button>
      <button class="btn-ghost small-btn" id="btnResNone">取消全选</button>
      <button class="btn-primary small-btn" id="btnResCopy">复制选中</button>
      <button class="btn-amber small-btn" id="btnResCut">剪切选中</button>
      <button class="btn-light small-btn" id="btnResMove">移动到文件夹…</button>
      <button class="btn-light small-btn" id="btnResPaste">粘贴到文件夹…</button>
      <button class="btn-ghost small-btn" id="btnResSysCopy" title="写入 Windows 系统剪贴板，可在资源管理器直接 Ctrl+V 粘贴">系统复制</button>
      <button class="btn-ghost small-btn" id="btnResSysCut" title="写入 Windows 系统剪贴板，在资源管理器 Ctrl+V 即移动">系统剪切</button>
      <button class="btn-ghost small-btn" id="btnClipClear">清空剪贴板</button>
      <span class="clip-state" id="resSelInfo">已选 0 个</span>
    </div>
    <div class="res-filters" id="resFilters">
      <div class="rf"><label>格式</label><select id="rExt"></select>
        <div class="rcustom" id="rExtBox" style="display:none"><input type="text" id="rExtCustom" placeholder="手输如 pdf,docx"></div></div>
      <div class="rf"><label>大小</label><select id="rSize"></select>
        <div class="rcustom" id="rSizeBox" style="display:none">
          <input type="text" id="rMinCustom" placeholder="最小 如 10MB"><input type="text" id="rMaxCustom" placeholder="最大 如 2GB"></div></div>
      <div class="rf"><label>日期</label><select id="rDate"></select>
        <div class="rcustom" id="rDateBox" style="display:none">
          <input type="date" id="rDateFromCustom" title="开始日期"><input type="date" id="rDateToCustom" title="结束日期"></div></div>
      <div class="rf"><label>排序</label><select id="rSort"></select></div>
      <button class="btn-ghost rreset" id="rReset">清除筛选</button>
      <span class="rtip">改动后自动重新搜索；选“自定义”可直接在此输入，并与高级筛选双向同步</span>
    </div>
    <div class="content-prog" id="contentProg"></div>
    <div id="dupArea" style="display:none;margin-bottom:12px;padding:12px;background:#fffbeb;
      border:1px solid #fde68a;border-radius:9px;font-size:13px;color:#92400e"></div>
    <div class="tbl-wrap">
      <table><thead><tr id="resultHead">
        <th style="width:30px"><input type="checkbox" id="resCheckAll" title="全选本页"></th>
        <th data-sort="name" style="width:44px">#</th>
        <th data-sort="name">文件名</th><th>后缀</th><th>分类</th>
        <th data-sort="size">大小</th><th data-sort="mtime">修改时间</th>
        <th data-sort="path">相对路径</th><th class="snip-col" style="display:none">正文片段</th>
        <th>操作</th>
      </tr></thead><tbody id="resultBody"></tbody></table>
      <div class="empty" id="empty" style="display:none">没有找到匹配的文件</div>
    </div>
    <div class="pager" id="pager"></div>
  </div>

  <div class="footer">纯本地运行 · 扫描与搜索只读 · 数据不离开本机　
    <button class="btn-ghost" id="btnQuit" style="padding:5px 12px;font-size:12px">退出程序</button></div>
</div>
</div><!-- /left -->

<div class="splitter" id="splitter" title="拖动调节左右区域宽度"></div>

<aside class="preview" id="previewPane">
  <div class="pv-head">
    <b>👁 文件预览</b>
    <span class="pv-name" id="pvName">未选择文件</span>
    <span class="pv-tools">
      <button class="btn-ghost pv-close" id="btnPvClose" title="关闭预览面板">×</button>
    </span>
  </div>
  <div class="pv-bar" id="pvBar" style="display:none">
    <span id="pvPageInfo"></span>
    <button class="btn-ghost" id="btnPvPrev">◀ 上一页</button>
    <button class="btn-ghost" id="btnPvNext">下一页 ▶</button>
    <button class="btn-light" id="btnPvOpen" title="用系统默认程序打开原文件">打开原文件 ↗</button>
  </div>
  <div class="pv-body" id="pvBody">
    <div class="pv-empty"><span class="ico">🖼️</span>
      点击搜索结果或类型列表中的<b>文件名 / 「预览」按钮</b>，<br>
      可预览的文件会在此显示内容。<br>
      图片、PDF、TXT/Office 文档支持预览；<br>
      Office 优先用本机 Microsoft/WPS 原版引擎还原排版，<br>
      未安装或调用失败时自动回退内置渲染/文字预览。</div>
  </div>
</aside>
<button class="pv-float" id="pvFloat" title="重新打开预览面板">👁 显示预览面板</button>

<!-- 文件夹浏览弹窗 -->
<div class="mask" id="browseMask"><div class="modal">
  <h3 id="browseTitle">选择文件夹</h3>
  <div class="body">
    <div class="bsearch"><input type="text" id="browseKw" autocomplete="off"
      placeholder="输入关键词快速找文件夹：如「二十年」「精选」，支持跨层级模糊匹配">
      <span id="browseSearchInfo"></span></div>
    <div class="bresults" id="browseResults" style="display:none"></div>
    <div class="crumb" id="crumb"></div>
    <div class="dirs" id="dirList"></div>
    <div id="filePreview" style="margin-top:14px;color:var(--muted);font-size:12px"></div>
  </div>
  <div class="foot">
    <button class="btn-ghost" id="btnUp">↑ 上一级</button>
    <button class="btn-light" id="btnPickHere">选择当前文件夹</button>
    <button class="btn-ghost" id="btnBrowseClose">关闭</button>
  </div>
</div></div>

<!-- 导出弹窗 -->
<div class="mask" id="exportMask"><div class="modal" style="width:min(520px,95vw)">
  <h3>导出搜索结果</h3>
  <div class="body">
    <div class="opt" style="margin-bottom:12px"><label>格式：</label>
      <select id="expFmt"><option value="tree">可折叠目录树 HTML（逐层展开，推荐）</option>
      <option value="html">HTML 表格</option>
      <option value="csv">CSV（Excel 打开，纯表格无折叠）</option><option value="txt">TXT 文本</option></select></div>
    <div class="note" style="margin-bottom:10px">💡 想在导出文件里像资源管理器一样逐层展开/折叠文件夹，请选「可折叠目录树 HTML」；CSV 是纯表格，无法折叠。</div>
    <div class="input-row"><input type="text" id="expDir" class="pick-input" readonly placeholder="选择输出文件夹">
      <button class="btn-light" id="btnBrowseExp">浏览…</button></div>
    <div class="note" style="margin-top:8px" id="expInfo"></div>
  </div>
  <div class="foot"><button class="btn-primary" id="btnDoExport">导出</button>
    <button class="btn-ghost" id="btnExpClose">取消</button></div>
</div></div>

<!-- 粘贴弹窗 -->
<div class="mask" id="pasteMask"><div class="modal" style="width:min(560px,95vw)">
  <h3>粘贴文件</h3>
  <div class="body">
    <div class="note" id="pasteClipInfo" style="margin-bottom:12px"></div>
    <div class="input-row"><input type="text" id="pasteDir" class="pick-input" readonly placeholder="选择目标文件夹">
      <button class="btn-light" id="btnBrowsePaste">浏览…</button></div>
    <div class="opt" style="margin-top:12px"><label>同名冲突：</label>
      <select id="pasteCollision"><option value="skip">跳过（不覆盖）</option>
      <option value="rename">自动重命名（保留两份）</option></select></div>
    <div class="paste-prog search-prog sp-run" id="pasteProg" style="display:none">
      <div class="sp-head"><span id="pastePhase">准备粘贴…</span><span class="sp-time" id="pasteTime"></span></div>
      <div class="sp-bar"><i id="pasteBar" class="indet"></i></div>
      <div class="sp-detail" id="pasteDetail"></div>
    </div>
  </div>
  <div class="foot"><button class="btn-primary" id="btnDoPaste">开始粘贴</button>
    <button class="btn-ghost" id="btnPasteCancel">关闭</button></div>
</div></div>

<!-- 保存的搜索弹窗 -->
<div class="mask" id="savedMask"><div class="modal" style="width:min(560px,95vw)">
  <h3>已保存的搜索（智能文件夹）</h3>
  <div class="body" id="savedList"></div>
  <div class="foot"><button class="btn-ghost" id="btnSavedClose">关闭</button></div>
</div></div>

<div class="toast" id="toast"></div>

<script>
const $=id=>document.getElementById(id);
let taskId=null, taskPath='', mode='smart', scope='stem', searchType='title',
    currentPage=0, pageSize=100, lastTotal=0, browseCtx={path:'',purpose:'scan'},
    expDir='', dupId=null, contentId=null, pollTimer=null, contentTimer=null,
    composing=false, lastParams=null, progKind='', progScanId='', progHideTimer=null,
    pasteDir='', pasteId=null, pasteTimer=null, clipState={mode:null,count:0},
    typeExt='', typeRows=[], resRows=[], expSource='',
    browseSearchTimer=null, toastTimer=null, resSelected=new Set();
function toast(text){const el=$('toast');if(!el)return;el.textContent=text;
 el.classList.add('show');clearTimeout(toastTimer);toastTimer=setTimeout(()=>el.classList.remove('show'),2600);}
function esc(s){return String(s==null?'':s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
function fmtSize(n){n=Number(n)||0;const u=['B','KB','MB','GB','TB'];let i=0;
 while(n>=1024&&i<u.length-1){n/=1024;i++;}return (i===0?n:n.toFixed(2))+' '+u[i];}
async function api(p,d){
 if(d!==undefined){const r=await fetch(p,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(d)});return r.json();}
 const r=await fetch(p);return r.json();
}
function msg(id,text,color){const e=$(id);if(e){e.textContent=text;e.style.color=color||'';}}

/* 筛选预设：格式 / 大小 / 日期，高级筛选与结果区共用 */
const EXT_PRESETS=['pdf','docx','doc','xlsx','xls','pptx','ppt','txt','csv','md','epub','mobi',
 'jpg','png','gif','bmp','mp4','mkv','avi','mp3','wav','zip','rar','7z','iso','torrent',
 'exe','msi','py','js','html','json','xml','dwg','psd'];
const SIZE_PRESETS=[['lt1m','小于 1MB','','1MB'],['1m10m','1–10MB','1MB','10MB'],
 ['10m100m','10–100MB','10MB','100MB'],['100m1g','100MB–1GB','100MB','1GB'],['gt1g','大于 1GB','1GB','']];
function ymd(d){return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');}
function dateRange(key){
 const now=new Date(),from=new Date(now),to=ymd(now);
 if(key==='7d')from.setDate(now.getDate()-7);
 else if(key==='30d')from.setDate(now.getDate()-30);
 else if(key==='90d')from.setDate(now.getDate()-90);
 else if(key==='year'){from.setMonth(0,1);from.setDate(1);}
 else return ['',''];
 return [ymd(from),to];
}
function initFilterUI(){
 const extOpts='<option value="">全部格式</option>'+EXT_PRESETS.map(e=>'<option value="'+e+'">.'+e+'</option>').join('')+'<option value="__custom">自定义（手输）</option>';
 const sizeOpts='<option value="">不限大小</option>'+SIZE_PRESETS.map(x=>'<option value="'+x[0]+'">'+x[1]+'</option>').join('')+'<option value="custom">自定义（手输）</option>';
 const dateOpts='<option value="">不限日期</option><option value="7d">最近 7 天</option><option value="30d">最近 30 天</option><option value="90d">最近 90 天</option><option value="year">今年</option><option value="custom">自定义（手选）</option>';
 ['fExtSel','rExt'].forEach(id=>$(id).innerHTML=extOpts);
 ['fSizeSel','rSize'].forEach(id=>$(id).innerHTML=sizeOpts);
 ['fDateSel','rDate'].forEach(id=>$(id).innerHTML=dateOpts);
 $('rSort').innerHTML=$('sort').innerHTML;
 // 下拉预设 → 填输入框 → 同步两处下拉与自定义输入 → 自动重搜
 [['fExtSel','rExt',applyExtPreset],['fSizeSel','rSize',applySizePreset],['fDateSel','rDate',applyDatePreset]].forEach(([a,b,fn])=>{
   const onSel=(self,other)=>{fn(self);$(other).value=$(self).value;syncResultCustom();doSearch();};
   $(a).onchange=()=>onSel(a,b);
   $(b).onchange=()=>onSel(b,a);
 });
 $('sort').onchange=()=>{$('rSort').value=$('sort').value;doSearch();};
 $('rSort').onchange=()=>{$('sort').value=$('rSort').value;doSearch();};
 // 高级筛选手输/手选 → 下拉回显示"自定义" → 同步结果区
 ['fExt','fMin','fMax','fDateFrom','fDateTo'].forEach(id=>{
   const h=()=>{syncPresetFromInputs();syncResultCustom();};
   $(id).addEventListener('input',h);
   $(id).addEventListener('change',h);
 });
 // 结果区自定义输入 → 写回高级筛选 → 防抖自动重搜
 ['rExtCustom','rMinCustom','rMaxCustom','rDateFromCustom','rDateToCustom'].forEach(id=>{
   $(id).addEventListener('input',onResultCustomInput);
   $(id).addEventListener('change',onResultCustomInput);
 });
 $('rReset').onclick=()=>{
   ['fExt','fMin','fMax','fDateFrom','fDateTo','fCat','fPath'].forEach(id=>$(id).value='');
   syncPresetFromInputs();syncResultCustom();doSearch();
 };
 syncPresetFromInputs();syncResultCustom();
}
/* 高级筛选：选中"自定义"时高亮并聚焦对应输入框 */
function clearHl(){['fExt','fMin','fMax','fDateFrom','fDateTo'].forEach(x=>$(x).classList.remove('hl-input'));}
function hlAdvanced(id){clearHl();if($(id)){$(id).classList.add('hl-input');try{$(id).focus();}catch(e){}}}
function applyExtPreset(sel){
 const v=$(sel).value;
 if(v!=='__custom'){clearHl();$('fExt').value=v;}else hlAdvanced('fExt');
}
function applySizePreset(sel){
 const v=$(sel).value;
 if(v==='custom'){hlAdvanced('fMin');return;}
 clearHl();
 const p=SIZE_PRESETS.find(x=>x[0]===v);
 $('fMin').value=p?p[2]:'';$('fMax').value=p?p[3]:'';
}
function applyDatePreset(sel){
 const v=$(sel).value;
 if(v==='custom'){hlAdvanced('fDateFrom');return;}
 clearHl();
 const r=dateRange(v);$('fDateFrom').value=r[0];$('fDateTo').value=r[1];
}
/* 结果区"自定义"输入：按当前下拉显示对应输入框，并与高级筛选保持同步 */
let customSearchTimer=null;
function syncResultCustom(){
 const show=(box,on)=>{$(box).style.display=on?'flex':'none';};
 show('rExtBox',$('rExt').value==='__custom');
 show('rSizeBox',$('rSize').value==='custom');
 show('rDateBox',$('rDate').value==='custom');
 $('rExtCustom').value=$('fExt').value;
 $('rMinCustom').value=$('fMin').value;
 $('rMaxCustom').value=$('fMax').value;
 $('rDateFromCustom').value=$('fDateFrom').value;
 $('rDateToCustom').value=$('fDateTo').value;
}
function onResultCustomInput(){
 $('fExt').value=$('rExtCustom').value;
 $('fMin').value=$('rMinCustom').value;
 $('fMax').value=$('rMaxCustom').value;
 $('fDateFrom').value=$('rDateFromCustom').value;
 $('fDateTo').value=$('rDateToCustom').value;
 syncPresetFromInputs();  // 两处下拉统一回显示为自定义/对应预设
 // 只更新输入框显示状态，不回写正在输入的值（避免打断光标）
 const show=(box,on)=>{$(box).style.display=on?'flex':'none';};
 show('rExtBox',$('rExt').value==='__custom');
 show('rSizeBox',$('rSize').value==='custom');
 show('rDateBox',$('rDate').value==='custom');
 clearTimeout(customSearchTimer);
 customSearchTimer=setTimeout(()=>doSearch(),350);
}
function syncPresetFromInputs(){
 const ev=$('fExt').value.trim();
 const ek=(ev===''?'':(EXT_PRESETS.includes(ev)?ev:'__custom'));
 $('fExtSel').value=ek;$('rExt').value=ek;
 const mn=$('fMin').value.trim(),mx=$('fMax').value.trim();
 let sk='custom';
 if(!mn&&!mx)sk='';else{const p=SIZE_PRESETS.find(x=>x[2]===mn&&x[3]===mx);if(p)sk=p[0];}
 $('fSizeSel').value=sk;$('rSize').value=sk;
 const df=$('fDateFrom').value,dt=$('fDateTo').value;
 let dk='custom';
 if(!df&&!dt)dk='';else{for(const k of ['7d','30d','90d','year']){const r=dateRange(k);if(r[0]===df&&r[1]===dt){dk=k;break;}}}
 $('fDateSel').value=dk;$('rDate').value=dk;
}

/* 初始化 */
(async()=>{
 initFilterUI();
 const dr=await api('/api/drives');
 $('quick').innerHTML='<span>快捷位置：</span>'+['Desktop','Documents','Downloads'].map(n=>
   '<button data-q="'+n+'">'+({Desktop:'桌面',Documents:'文档',Downloads:'下载'}[n])+'</button>').join('')+
   dr.drives.map(d=>'<button data-drive="'+d.path+'">'+d.path.slice(0,2)+'</button>').join('');
 $('quick').querySelectorAll('button').forEach(b=>b.onclick=async()=>{
   if(b.dataset.q){const r=await api('/api/quickpath?name='+b.dataset.q);if(r.path)chooseScanPath(r.path);}
   else if(b.dataset.drive)chooseScanPath(b.dataset.drive);
 });
 const cats=['办公文档','文本与数据','图片','视频','音频','压缩包与镜像','程序与脚本','种子与下载','编程开发','字体','其他'];
 $('fCat').innerHTML='<option value="">全部分类</option>'+cats.map(c=>'<option>'+c+'</option>').join('');
 const data=await api('/api/data');renderHistory(data.history||[]);
 // 恢复剪贴板状态（服务重启会清空，这里只读一次）
 try{const c=await api('/api/clip');clipState={mode:c.mode,count:c.count||0};refreshClipUI();}catch(e){}
})();
function chooseScanPath(p){$('scanPath').value=p;}
function setRes(msgTxt){$('resInfo').innerHTML=msgTxt||'尚未搜索';}
/* 功能2统一进度面板：phase=阶段文字, pct=百分比(null=不确定动画), detail=明细, elapsed=用时, kind=run/done/err, canCancel=显示停止按钮 */
function showProg(phase,pct,detail,elapsed,kind,canCancel){
 if(progHideTimer){clearTimeout(progHideTimer);progHideTimer=null;}
 const w=$('searchProg');w.style.display='block';
 w.className='search-prog'+(kind==='done'?' sp-done':'')+(kind==='err'?' sp-err':'')+(kind==='run'?' sp-run':'');
 $('spPhase').textContent=phase||'';
 $('spTime').textContent=(elapsed!=null&&elapsed!=='')?('用时 '+elapsed+'s'):'';
 $('spDetail').textContent=detail||'';
 const bar=$('spBar');
 if(pct==null){bar.className='indet';bar.style.width='';}
 else{bar.className='';bar.style.width=Math.max(0,Math.min(100,pct))+'%';}
 $('spCancel').style.display=canCancel?'':'none';
}
function hideProg(){$('searchProg').style.display='none';$('spCancel').style.display='none';}
function resetSearchBtn(){$('btnSearch').disabled=false;$('btnSearch').textContent='搜索';}
$('spCancel').onclick=async()=>{
 showProg('正在停止…',null,'正在通知后台停止，请稍候',null,'run',false);
 try{
  if(progKind==='content'&&contentId)await api('/api/content/cancel',{id:contentId});
  else if(progKind==='scan'&&progScanId)await api('/api/cancel',{id:progScanId});
 }catch(e){}
};

/* 模式/范围/类型切换 */
$('modes').querySelectorAll('.chip').forEach(c=>c.onclick=()=>{
 $('modes').querySelectorAll('.chip').forEach(x=>x.classList.remove('on'));c.classList.add('on');mode=c.dataset.mode;
});
$('scopes').querySelectorAll('.chip').forEach(c=>c.onclick=()=>{
 $('scopes').querySelectorAll('.chip').forEach(x=>x.classList.remove('on'));c.classList.add('on');scope=c.dataset.scope;
});
$('searchTabs').querySelectorAll('.tab').forEach(t=>t.onclick=()=>{
 $('searchTabs').querySelectorAll('.tab').forEach(x=>x.classList.remove('on'));t.classList.add('on');
 searchType=t.dataset.type;
 const isContent=searchType==='content';
 $('titleOpts').style.display=isContent?'none':'block';
 $('titleHint').style.display=isContent?'none':'block';
 $('contentHint').style.display=isContent?'block':'none';
 document.querySelectorAll('.snip-col').forEach(x=>x.style.display=isContent?'':'none');
 $('query').placeholder=isContent?'输入正文中可能出现的关键词，例如：方沸石':'输入文件名关键词，例如：方沸石';
});
$('advToggle').onchange=e=>$('adv').classList.toggle('show',e.target.checked);
/* 中文输入法安全：组合期间不触发搜索；keyCode 229 为输入法处理中 */
$('query').addEventListener('compositionstart',()=>composing=true);
$('query').addEventListener('compositionend',()=>{composing=false;});
$('query').addEventListener('keydown',e=>{
 if(e.key==='Enter'&&!composing&&!e.isComposing&&e.keyCode!==229){e.preventDefault();doSearch();}
});
$('btnClear').onclick=()=>{$('query').value='';['fExt','fMin','fMax','fPath','fDateFrom','fDateTo'].forEach(id=>$(id).value='');
 $('fCat').value='';syncPresetFromInputs();setRes();$('resultBody').innerHTML='';};
$('btnResearch').onclick=()=>doSearch();
$('btnSearch').onclick=()=>doSearch();   // 修复：此前搜索按钮从未绑定点击事件
document.querySelectorAll('th[data-sort]').forEach(th=>th.onclick=()=>{
 const s=th.dataset.sort;if(['name','size','mtime','path'].includes(s)){$('sort').value=s;doSearch();}
});

/* 功能1扫描结果折叠面板 */
let scanFoldOpen=false;
function setScanFold(open){
  scanFoldOpen=!!open;
  $('scanResult').style.display=scanFoldOpen?'block':'none';
  $('scanFoldBtn').textContent=scanFoldOpen?'收起 ▲':'展开 ▼';
  $('scanFoldHint').textContent=scanFoldOpen?'（点击收起以节省空间）':'（完成扫描后自动展开，可点击收起）';
}
$('scanFoldHead').onclick=()=>setScanFold(!scanFoldOpen);

/* 扫描 */
$('btnBrowseScan').onclick=()=>openBrowse('scan');
$('btnBrowseSearch').onclick=()=>openBrowse('search');
$('btnUseScanPath').onclick=()=>{if(!$('scanPath').value){alert('请先在功能1选择文件夹');return;}
 $('searchPath').value=$('scanPath').value;};
$('btnScan').onclick=async()=>{
 const p=$('scanPath').value.trim();if(!p){alert('请先选择扫描范围');return;}
 $('btnScan').disabled=true;$('btnCancel').disabled=false;$('btnSearch').disabled=true;
 setScanFold(false);
 $('progressWrap').style.display='block';$('stats').style.display='none';$('catGrid').style.display='none';$('catTip').style.display='none';
 $('levelBars').style.display='none';$('skipNote').style.display='none';
 $('status').textContent='正在启动...';$('levelInfo').textContent='准备按层扫描：根目录 → 第1层 → 第2层 → …';
 const r=await api('/api/scan',{path:p,max_depth:$('maxDepth').value,
   exclude_hidden:$('excHidden').checked,follow_links:$('followLinks').checked});
 if(r.error){alert(r.error);$('btnScan').disabled=false;$('btnCancel').disabled=true;return;}
 taskId=r.id;taskPath=p;pollTimer=setInterval(poll,300);
};
$('btnCancel').onclick=async()=>{if(taskId)await api('/api/cancel',{id:taskId});};
async function poll(){
 const s=await api('/api/progress?id='+encodeURIComponent(taskId));
 if(s.error){clearInterval(pollTimer);alert(s.error);return;}
 $('status').textContent=s.status_text+'（用时 '+s.elapsed+'s）';
 $('levelInfo').textContent='正在扫描第 '+s.current_level+' 层（已到达最深第 '+s.max_level+' 层）';
 if(s.status==='done'||s.status==='cancelled'){
  clearInterval(pollTimer);
  $('btnScan').disabled=false;$('btnCancel').disabled=true;$('btnSearch').disabled=false;
  $('stats').style.display='grid';
  $('sFiles').textContent=s.files;$('sDirs').textContent=s.dirs;
  $('sKinds').textContent=(s.max_level+1)+' / '+(s.extensions||[]).length;
  $('sTime').textContent=s.elapsed+'s';
  const lf=s.level_files||{},ld=s.level_dirs||{};
  const keys=Object.keys(lf).concat(Object.keys(ld)).filter((v,i,a)=>a.indexOf(v)===i).map(Number).sort((a,b)=>a-b);
  if(keys.length){$('levelBars').innerHTML='<b>逐层扫描明细</b>（第0层=根目录）：<br>'+
    keys.map(k=>'· 第 '+k+' 层：'+(ld[k]||0)+' 个文件夹，'+(lf[k]||0)+' 个文件').join('<br>');
    $('levelBars').style.display='block';}
  if(s.skipped>0){$('skipNote').innerHTML='<b>'+s.skipped+' 处目录无法访问，已跳过</b><br>'+
    (s.skipped_samples||[]).map(x=>'· '+esc(x.path)+(x.reason?'（'+esc(x.reason)+'）':'')).join('<br>');
    $('skipNote').style.display='block';}
  renderCategories(s);
  setScanFold(true);  // 扫描完成自动展开，之后用户可手动收起
 }
}

/* 分类展示 */
function renderCategories(s){
 const exts=s.extensions||[];if(!exts.length)return;
 const groups={};
 exts.forEach(e=>{(groups[e.category]||(groups[e.category]={count:0,size:0,items:[]}));
   groups[e.category].count+=e.count;groups[e.category].size+=e.size;groups[e.category].items.push(e);});
 const order=['办公文档','文本与数据','图片','视频','音频','压缩包与镜像','程序与脚本','种子与下载','编程开发','字体','其他'];
 const keys=Object.keys(groups).sort((a,b)=>order.indexOf(a)-order.indexOf(b));
 $('catGrid').innerHTML=keys.map(k=>{
   const g=groups[k];
   const chips=g.items.map(e=>'<span class="ext-chip" data-ext="'+esc(e.key)+'" data-label="'+esc(e.ext)+'"><b>.'+esc(e.ext)+'</b> '+e.count+
     ' <span style="color:#9ca3af">'+fmtSize(e.size)+'</span></span>').join('');
   return '<div class="cat-card"><h4>'+esc(k)+'<span>'+g.count+' 个 · '+fmtSize(g.size)+'</span></h4>'+chips+'</div>';
 }).join('');
 $('catGrid').style.display='grid';$('catTip').style.display='block';
 $('catGrid').querySelectorAll('[data-ext]').forEach(el=>el.onclick=()=>{
   // 功能1：点击后缀标签 → 列出该类型全部文件（可勾选/复制/剪切/导出）
   showTypeFiles(el.dataset.ext, el.dataset.label||el.dataset.ext);
 });
}

/* 功能1：列出某后缀类型的全部文件 */
async function showTypeFiles(ext,label){
 if(!taskId){alert('请先扫描');return;}
 typeExt=ext;
 const r=await api('/api/type-files?id='+encodeURIComponent(taskId)+'&ext='+encodeURIComponent(ext));
 if(r.error){alert(r.error);return;}
 typeRows=r.rows||[];
 $('typeTitle').textContent='类型文件列表：'+(label?'.'+label:'全部')+'（共 '+r.total+' 个）';
 const body=$('typeBody');body.innerHTML='';
 $('typeEmpty').style.display=typeRows.length?'none':'block';
 typeRows.forEach(f=>{
  const tr=document.createElement('tr');
  tr.innerHTML='<td class="ck"><input type="checkbox" class="type-ckbox" data-path="'+esc(f.path)+'"></td>'+
   '<td class="fn pv-click" data-tpreview="'+esc(f.path)+'" title="'+esc(f.path)+'（点击预览）">'+esc(f.name)+'</td>'+
   '<td>'+fmtSize(f.size)+'</td><td>'+esc(f.mtime)+'</td>'+
   '<td class="rp" title="'+esc(f.rel)+'">'+esc(f.rel)+'</td>'+
   '<td><span class="op2" data-tcopy="'+esc(f.path)+'" title="复制到本程序剪贴板">内复制</span>'+
   '<span class="op2 cut" data-tcut="'+esc(f.path)+'" title="剪切到本程序剪贴板">内剪切</span>'+
   '<span class="op2 sys" data-tsyscopy="'+esc(f.path)+'" title="写入系统剪贴板，可在资源管理器 Ctrl+V 粘贴">系统复制</span>'+
   '<span class="op2 sys cut" data-tsyscut="'+esc(f.path)+'" title="写入系统剪贴板，在资源管理器 Ctrl+V 即移动">系统剪切</span>'+
   '<span class="op2 pv" data-tpreview="'+esc(f.path)+'">预览</span>'+
   '<span class="op2" data-topen="'+esc(f.path)+'">定位</span></td>';
  body.appendChild(tr);
 });
 body.querySelectorAll('.type-ckbox').forEach(cb=>cb.onchange=updateTypeSelInfo);
 body.querySelectorAll('[data-tpreview]').forEach(el=>el.onclick=()=>showPreview(el.dataset.tpreview));
 body.querySelectorAll('[data-topen]').forEach(el=>el.onclick=()=>api('/api/reveal?p='+encodeURIComponent(el.dataset.topen)));
 body.querySelectorAll('[data-tcopy]').forEach(el=>el.onclick=()=>setClip([el.dataset.tcopy],'copy'));
 body.querySelectorAll('[data-tcut]').forEach(el=>el.onclick=()=>setClip([el.dataset.tcut],'cut'));
 body.querySelectorAll('[data-tsyscopy]').forEach(el=>el.onclick=()=>setSystemClip([el.dataset.tsyscopy],'copy'));
 body.querySelectorAll('[data-tsyscut]').forEach(el=>el.onclick=()=>setSystemClip([el.dataset.tsyscut],'cut'));
 $('typePanel').classList.add('show');
 updateTypeSelInfo();
 $('typePanel').scrollIntoView({behavior:'smooth',block:'start'});
}
function selectedTypePaths(){return Array.from(document.querySelectorAll('.type-ckbox:checked')).map(c=>c.dataset.path);}
function updateTypeSelInfo(){if($('typeSelInfo'))$('typeSelInfo').textContent='已选 '+selectedTypePaths().length+' 个';}
$('btnTypeClose').onclick=()=>$('typePanel').classList.remove('show');
$('btnTypeAll').onclick=()=>{document.querySelectorAll('.type-ckbox').forEach(c=>c.checked=true);updateTypeSelInfo();};
$('btnTypeNone').onclick=()=>{document.querySelectorAll('.type-ckbox').forEach(c=>c.checked=false);updateTypeSelInfo();};
$('btnTypeCopy').onclick=()=>setClip(selectedTypePaths(),'copy');
$('btnTypeCut').onclick=()=>setClip(selectedTypePaths(),'cut');
$('btnTypeMove').onclick=async()=>{const p=selectedTypePaths();
 if(!p.length){alert('请先勾选文件');return;}await setClip(p,'cut');openPaste();};
$('btnTypePaste').onclick=openPaste;
$('btnTypeSysCopy').onclick=()=>setSystemClip(selectedTypePaths(),'copy');
$('btnTypeSysCut').onclick=()=>setSystemClip(selectedTypePaths(),'cut');
$('btnTypeCopyAll').onclick=()=>setClip(typeRows.map(f=>f.path),'copy');
$('btnTypeExport').onclick=()=>{expSource='type';$('exportMask').classList.add('show');};

/* 搜索：确保目标文件夹已扫描，再执行 */
function collectFilters(){return {ext:$('fExt').value,category:$('fCat').value,size_min:$('fMin').value,
   size_max:$('fMax').value,date_from:$('fDateFrom').value,date_to:$('fDateTo').value,
   path_contains:$('fPath').value};}
function collectParams(offset){return {id:taskId,query:$('query').value.trim(),mode:mode,scope:scope,
   sort:$('sort').value,offset:offset,limit:pageSize,filters:collectFilters()};}
function targetPath(){return $('searchPath').value.trim()||$('scanPath').value.trim();}

function ensureScanned(){
 return new Promise((resolve,reject)=>{
  const p=targetPath();
  if(!p){reject('请先选择要搜索的文件夹');return;}
  if(taskId&&taskPath&&p.replace(/\\/g,'/').toLowerCase()===taskPath.replace(/\\/g,'/').toLowerCase()){
    showProg('① 文件夹已扫描',100,'复用已有扫描结果，正在准备搜索',null,'run',false);
    resolve();return;}
  // 自动扫描目标文件夹：这是"点搜索没反应"的主要环节，必须给出可见进度
  progKind='scan';progScanId='';
  setRes('<b>正在扫描目标文件夹…</b>');
  showProg('① 正在扫描文件夹…',null,'正在启动扫描任务（大文件夹可能需要几十秒到几分钟）',null,'run',true);
  $('btnSearch').disabled=true;
  api('/api/scan',{path:p,max_depth:'0',exclude_hidden:false,follow_links:true}).then(r=>{
   if(r.error){progKind='';resetSearchBtn();reject(r.error);return;}
   const id=r.id;progScanId=id;
   const tm=setInterval(async()=>{
    let s;
    try{s=await api('/api/progress?id='+encodeURIComponent(id));}
    catch(e){return;}   // 网络抖动：跳过本次轮询，下次再取
    if(s.error){clearInterval(tm);progKind='';resetSearchBtn();reject(s.error);return;}
    showProg('① 正在扫描文件夹…',null,
      '第 '+s.current_level+' 层 · 已发现 '+s.files+' 个文件 / '+s.dirs+' 个文件夹 · '+s.status_text,
      s.elapsed,'run',true);
    setRes('<b>正在扫描：'+esc(s.status_text)+'</b>');
    if(s.status==='done'||s.status==='cancelled'){
     clearInterval(tm);progKind='';progScanId='';resetSearchBtn();
     if(s.status==='cancelled'){reject('扫描已取消');return;}
     taskId=id;taskPath=p;
     // 同步功能1概览
     $('scanPath').value=p;
     renderCategories(s);
     $('stats').style.display='grid';$('sFiles').textContent=s.files;$('sDirs').textContent=s.dirs;
     $('sKinds').textContent=(s.max_level+1)+' / '+(s.extensions||[]).length;$('sTime').textContent=s.elapsed+'s';
     resolve();
    }
   },300);
  });
 });
}
function doSearch(offset,keepSel){
 offset=offset||0;
 if(!keepSel){resSelected.clear();updateResSelInfo();}
 const q=$('query').value.trim();
 if(!q){if(!Object.values(collectFilters()).some(v=>v&&v.toString().trim())){alert('请输入关键词或设置筛选条件');return;}}
 // 立即反馈：让用户在按钮附近就能看到程序正在工作
 $('btnSearch').disabled=true;$('btnSearch').textContent='搜索中…';
 showProg('① 准备中…',null,'正在检查搜索范围，如需扫描将自动开始',null,'run',false);
 ensureScanned().then(()=>{
   currentPage=Math.floor(offset/pageSize);
   if(searchType==='content')startContentSearch(offset);
   else runTitleSearch(offset);
 }).catch(e=>{
   resetSearchBtn();
   showProg('操作未完成',null,String(e),null,'err',false);
   setRes(esc(e));
 });
}
/* 标题搜索 */
function runTitleSearch(offset){
 const params=collectParams(offset);lastParams=params;
 progKind='search';
 showProg('② 正在匹配文件名…',null,'正在从已扫描的文件中查找关键词（大文件夹约需 1–3 秒）',null,'run',false);
 setRes('<b>搜索中…</b>');
 api('/api/search',params).then(r=>{
  progKind='';resetSearchBtn();
  if(r.error){showProg('搜索失败',null,r.error,null,'err',false);setRes(esc(r.error));return;}
  lastTotal=r.total;renderResults(r.rows,r.terms,false);renderPager();
  setRes('找到 <b>'+r.total+'</b> 个文件　<span style="color:var(--muted)">用时 '+r.took+' ms</span>');
  showProg('搜索完成',100,'找到 '+r.total+' 个文件',null,'done',false);
  progHideTimer=setTimeout(hideProg,4000);
 });
}

/* 正文搜索 */
function startContentSearch(offset){
 const params=collectParams(offset);lastParams=params;
 progKind='content';
 showProg('② 正在准备正文搜索…',null,'正在筛选候选文件',null,'run',true);
 setRes('<b>正文搜索启动中…</b>');
 $('contentProg').style.display='block';$('contentProg').textContent='正在准备正文搜索…';
 $('btnContentCancel').style.display='';
 const cf=Object.assign({},params.filters,{skip_slow:$('skipSlow').checked});
 api('/api/content/start',{id:taskId,query:params.query,mode:params.mode,filters:cf}).then(r=>{
  if(r.error){progKind='';resetSearchBtn();$('contentProg').style.display='none';$('btnContentCancel').style.display='none';
    showProg('正文搜索失败',null,r.error,null,'err',false);setRes(esc(r.error));return;}
  contentId=r.id;pollContent();
 });
}
async function pollContent(){
 let s;
 try{s=await api('/api/content/status?id='+encodeURIComponent(contentId)+'&offset='+(currentPage*pageSize)+'&limit='+pageSize);}
 catch(e){contentTimer=setTimeout(pollContent,500);return;}
 if(s.error){$('contentProg').textContent=esc(s.error);showProg('正文搜索失败',null,s.error,null,'err',false);return;}
 const pct=s.total>0?Math.round(s.scanned*100/s.total):null;
 showProg('② 正在读取正文…',pct,
   '已读 '+s.scanned+'/'+s.total+' 个 · 命中 '+s.total_hits+' 个 · '+s.status_text,
   s.elapsed,'run',true);
 $('contentProg').innerHTML='📄 '+esc(s.status_text)+'（已读 '+s.scanned+'/'+s.total+'，命中 '+s.total_hits+'）';
 if(s.status==='done'||s.status==='cancelled'||s.status==='error'){
  progKind='';resetSearchBtn();
  $('btnContentCancel').style.display='none';
  if(s.status==='done'){
   lastTotal=s.total_hits;renderResults(s.rows,s.query?[s.query]:[],true);renderPager();
   setRes('正文中找到 <b>'+s.total_hits+'</b> 个文件　<span style="color:var(--muted)">用时 '+s.elapsed+'s</span>');
   showProg('正文搜索完成',100,'命中 '+s.total_hits+' 个文件',s.elapsed,'done',false);
   progHideTimer=setTimeout(hideProg,4000);
  }else{
   setRes(esc(s.status_text));
   showProg(s.status==='cancelled'?'已停止':'正文搜索失败',null,s.status_text,s.elapsed,
     s.status==='cancelled'?'done':'err',false);
   progHideTimer=setTimeout(hideProg,4000);
  }
  return;
 }
 contentTimer=setTimeout(pollContent,500);
}
$('btnContentCancel').onclick=async()=>{if(contentId){await api('/api/content/cancel',{id:contentId});
 $('contentProg').textContent='正在停止…';}};

/* 高亮标红 */
function highlight(text,terms){
 let h=esc(text);if(!terms||!terms.length)return h;
 const sorted=terms.slice().filter(Boolean).sort((a,b)=>b.length-a.length);
 for(const t of sorted){
  try{const re=new RegExp('('+t.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')+')','gi');
   h=h.replace(re,'<mark>$1</mark>');}catch(e){}
 }
 return h.replace(/<mark>(<mark>(.*?)<\/mark>)<\/mark>/g,'<mark>$2</mark>');
}
function renderResults(rows,terms,isContent){
 resRows=rows;
 const body=$('resultBody');body.innerHTML='';
 $('empty').style.display=rows.length?'none':'block';
 rows.forEach((r,i)=>{
  const tr=document.createElement('tr');
  let html='<td class="res-ck"><input type="checkbox" class="res-ckbox" data-path="'+esc(r.path)+'"'+
   (resSelected.has(r.path)?' checked':'')+'></td>'+
   '<td style="color:#9ca3af">'+(currentPage*pageSize+i+1)+'</td>'+
   '<td class="c-name pv-click" data-preview="'+esc(r.path)+'" title="'+esc(r.path)+'（点击预览）">'+highlight(r.name,terms)+'</td>'+
   '<td><span class="tag">'+esc(r.ext||'-')+'</span></td>'+
   '<td>'+esc(r.category)+'</td>'+
   '<td>'+fmtSize(r.size)+'</td>'+
   '<td>'+esc(r.mtime)+'</td>'+
   '<td class="c-rel" title="'+esc(r.rel)+'">'+esc(r.rel)+'</td>';
  if(isContent){html+='<td class="c-snip">'+highlight(r.snippet||'',r.terms||terms)+
    (r.hit_count>1?' <span style="color:#9ca3af">('+r.hit_count+'处)</span>':'')+'</td>';}
  html+='<td><span class="op2" data-copy="'+esc(r.path)+'" title="复制到本程序剪贴板">内复制</span>'+
   '<span class="op2 cut" data-cut="'+esc(r.path)+'" title="剪切到本程序剪贴板">内剪切</span>'+
   '<span class="op2 sys" data-syscopy="'+esc(r.path)+'" title="写入 Windows 系统剪贴板，可在资源管理器 Ctrl+V 粘贴">系统复制</span>'+
   '<span class="op2 sys cut" data-syscut="'+esc(r.path)+'" title="写入系统剪贴板，在资源管理器 Ctrl+V 即移动">系统剪切</span>'+
   '<span class="op2 pv" data-preview="'+esc(r.path)+'">预览</span>'+
   '<span class="op" data-open="'+esc(r.path)+'">定位</span></td>';
  tr.innerHTML=html;body.appendChild(tr);
 });
 body.querySelectorAll('[data-open]').forEach(el=>el.onclick=()=>{
  api('/api/reveal?p='+encodeURIComponent(el.dataset.open));
 });
 body.querySelectorAll('[data-preview]').forEach(el=>el.onclick=()=>showPreview(el.dataset.preview));
 body.querySelectorAll('[data-copy]').forEach(el=>el.onclick=()=>setClip([el.dataset.copy],'copy'));
 body.querySelectorAll('[data-cut]').forEach(el=>el.onclick=()=>setClip([el.dataset.cut],'cut'));
 body.querySelectorAll('[data-syscopy]').forEach(el=>el.onclick=()=>setSystemClip([el.dataset.syscopy],'copy'));
 body.querySelectorAll('[data-syscut]').forEach(el=>el.onclick=()=>setSystemClip([el.dataset.syscut],'cut'));
 body.querySelectorAll('.res-ckbox').forEach(cb=>cb.onchange=()=>{
  if(cb.checked)resSelected.add(cb.dataset.path);else resSelected.delete(cb.dataset.path);
  updatePageCheckAll();updateResSelInfo();
 });
 syncPageChecks();updateResSelInfo();
}
function selectedResPaths(){return Array.from(resSelected);}
function updatePageCheckAll(){
 const boxes=document.querySelectorAll('.res-ckbox');
 if($('resCheckAll'))$('resCheckAll').checked=boxes.length>0&&Array.from(boxes).every(c=>c.checked);
}
function syncPageChecks(){
 document.querySelectorAll('.res-ckbox').forEach(c=>{c.checked=resSelected.has(c.dataset.path);});
 updatePageCheckAll();
}
function updateResSelInfo(){
 if($('resSelInfo'))$('resSelInfo').textContent='已选 '+resSelected.size+' 个'+
  (resSelected.size>resRows.length?'（含其他页）':'');
}

/* 剪贴板 / 复制 / 剪切 / 粘贴（功能1类型列表与功能2结果区共用） */
async function setClip(paths,mode){
 if(!paths||!paths.length){alert('请先勾选文件');return;}
 const r=await api('/api/clip/set',{mode:mode,paths:paths,source:'search'});
 if(r.error){alert(r.error);return;}
 clipState={mode:mode,count:r.count};refreshClipUI();
}
async function setSystemClip(paths,mode){
 if(!paths||!paths.length){alert('请先勾选文件');return;}
 const r=await api('/api/sysclip/set',{mode:mode,paths:paths});
 if(r.error){alert(r.error);return;}
 toast('已放入系统剪贴板：'+(mode==='cut'?'剪切':'复制')+' '+r.count+' 个文件，可在文件夹里 Ctrl+V '+(mode==='cut'?'移动':'粘贴'));
}
async function clearClip(){await api('/api/clip/clear',{});clipState={mode:null,count:0};refreshClipUI();}
function refreshClipUI(){
 const s=$('resSelInfo');
 if(clipState.count>0){
  if(s){s.innerHTML='剪贴板：<b style="color:#1d4ed8">'+(clipState.mode==='cut'?'剪切':'复制')+' '+clipState.count+' 个</b>';}
 }else{updateResSelInfo();}
 updatePasteClipInfo();
}
function updatePasteClipInfo(){
 const el=$('pasteClipInfo');if(!el)return;
 if(clipState.count>0){el.innerHTML='当前剪贴板：<b>'+(clipState.mode==='cut'?'剪切':'复制')+' '+clipState.count+' 个文件</b>（来自搜索结果或类型列表）。';}
 else{el.textContent='剪贴板为空，请先在结果区或类型列表勾选文件并点「复制选中」或「剪切选中」。';}
}
function openPaste(){updatePasteClipInfo();
 if(clipState.count<=0){alert('请先勾选文件并复制或剪切');return;}
 $('pasteMask').classList.add('show');$('pasteProg').style.display='none';}
$('btnResCopy').onclick=()=>setClip(selectedResPaths(),'copy');
$('btnResCut').onclick=()=>setClip(selectedResPaths(),'cut');
$('btnResMove').onclick=async()=>{const p=selectedResPaths();
 if(!p.length){alert('请先勾选文件');return;}await setClip(p,'cut');openPaste();};
$('btnResSysCopy').onclick=()=>setSystemClip(selectedResPaths(),'copy');
$('btnResSysCut').onclick=()=>setSystemClip(selectedResPaths(),'cut');
$('btnResAll').onclick=()=>{
 document.querySelectorAll('.res-ckbox').forEach(c=>{c.checked=true;resSelected.add(c.dataset.path);});
 syncPageChecks();updateResSelInfo();
 toast('已选本页 '+document.querySelectorAll('.res-ckbox').length+' 个，累计 '+resSelected.size+' 个');
};
$('btnResNone').onclick=()=>{resSelected.clear();syncPageChecks();updateResSelInfo();};
$('btnResAllPages').onclick=async()=>{
 if(!taskId){alert('请先搜索');return;}
 const btn=$('btnResAllPages'),old=btn.textContent;
 btn.disabled=true;btn.textContent='正在选择…';
 try{
  let r;
  if(searchType==='content'){
   if(!contentId){alert('请先做正文搜索');return;}
   r=await api('/api/content/all?id='+encodeURIComponent(contentId));
  }else{
   const p=lastParams?Object.assign({},lastParams):collectParams(0);
   r=await api('/api/search/paths',p);
  }
  if(r.error){alert(r.error);return;}
  r.paths.forEach(p=>resSelected.add(p));
  syncPageChecks();updateResSelInfo();
  toast('已全选符合条件的 '+r.paths.length+' 个文件（跨页）');
 }finally{btn.disabled=false;btn.textContent=old;}
};
$('resCheckAll').onchange=e=>{
 document.querySelectorAll('.res-ckbox').forEach(c=>{
  c.checked=e.target.checked;
  if(e.target.checked)resSelected.add(c.dataset.path);else resSelected.delete(c.dataset.path);
 });
 updateResSelInfo();
};
$('btnClipClear').onclick=clearClip;
$('btnResPaste').onclick=openPaste;
$('btnBrowsePaste').onclick=()=>openBrowse('paste');
$('btnPasteCancel').onclick=()=>{$('pasteMask').classList.remove('show');};
$('btnDoPaste').onclick=async()=>{
 if(!pasteDir){alert('请选择目标文件夹');return;}
 if(clipState.count<=0){alert('剪贴板为空');return;}
 const r=await api('/api/paste',{dest:pasteDir,collision:$('pasteCollision').value});
 if(r.error){alert(r.error);return;}
 pasteId=r.id;$('pasteProg').style.display='block';$('pasteProg').className='search-prog sp-run';
 $('btnDoPaste').disabled=true;pollPaste();
};
async function pollPaste(){
 const s=await api('/api/paste/status?id='+encodeURIComponent(pasteId));
 if(s.error){$('pastePhase').textContent=s.error;return;}
 const pct=s.total?Math.round(s.done/s.total*100):100;
 $('pasteBar').style.width=pct+'%';$('pasteBar').classList.remove('indet');
 $('pastePhase').textContent=s.status==='done'?'✓ 粘贴完成':(s.status==='error'?'粘贴出错':(s.mode==='cut'?'正在移动…':'正在复制…'));
 $('pasteDetail').textContent='已处理 '+s.done+'/'+s.total+'（复制 '+s.copied+' · 移动 '+s.moved+' · 跳过 '+s.skipped+'）'+
   (s.error_count?'　出错 '+s.error_count+' 个':'');
 $('pasteTime').textContent=s.elapsed+'s';
 if(s.status==='done'||s.status==='error'||s.status==='cancelled'){
  if(s.status==='done')$('pasteProg').className='search-prog sp-done';
  $('btnDoPaste').disabled=false;clearClip();
  return;
 }
 pasteTimer=setTimeout(pollPaste,400);
}
function renderPager(){
 const pages=Math.ceil(lastTotal/pageSize),p=$('pager');p.innerHTML='';
 if(pages<=1)return;
 const mk=(label,off,dis)=>{const b=document.createElement('button');b.textContent=label;
  b.className='btn-ghost';b.disabled=dis;b.onclick=()=>doSearch(off,true);p.appendChild(b);};
 mk('首页',0,currentPage===0);mk('上一页',Math.max(0,currentPage-1)*pageSize,currentPage===0);
 let s=Math.max(0,currentPage-2),e=Math.min(pages-1,currentPage+2);
 for(let i=s;i<=e;i++){const b=document.createElement('button');
  b.textContent=(i+1);b.className=i===currentPage?'btn-primary':'btn-ghost';
  b.onclick=()=>doSearch(i*pageSize,true);p.appendChild(b);}
 mk('下一页',Math.min(pages-1,currentPage+1)*pageSize,currentPage>=pages-1);
 mk('末页',(pages-1)*pageSize,currentPage>=pages-1);
}

/* 历史 / 保存 */
function renderHistory(list){
 $('history').innerHTML=list.length?('<span>最近搜索：</span>'+list.slice(0,8).map(q=>
   '<button data-h="'+esc(q)+'">'+esc(q)+'</button>').join('<button id="clearHist">清空</button>')):'';
 $('history').querySelectorAll('[data-h]').forEach(b=>b.onclick=()=>{$('query').value=b.dataset.h;doSearch();});
 if($('clearHist'))$('clearHist').onclick=async()=>{await api('/api/data/clear_history',{});renderHistory([]);};
}
$('btnSave').onclick=async()=>{
 const name=prompt('给这个搜索起个名字：',$('query').value.trim()||'未命名搜索');
 if(!name)return;
 const r=await api('/api/saved/add',{name:name,params:Object.assign({},collectParams(0),{searchType:searchType})});
 alert(r.error||('已保存：'+name));
};
$('btnSaved').onclick=async()=>{
 const data=await api('/api/data');const box=$('savedList');box.innerHTML='';
 if(!(data.saved||[]).length){box.innerHTML='<div class="empty">还没有保存的搜索</div>';}
 else data.saved.forEach(it=>{
   const p=it.params||{};
   const d=document.createElement('div');d.className='saved-item';
   d.innerHTML='<div><b>'+esc(it.name)+'</b><div>'+
     ((p.query?'<span class="tagchip">词: '+esc(p.query)+'</span>':'')+
      (p.searchType==='content'?'<span class="tagchip">正文</span>':'')+
      (p.filters&&p.filters.ext?'<span class="tagchip">后缀: '+esc(p.filters.ext)+'</span>':'')+
      (p.mode?'<span class="tagchip">'+esc(p.mode)+'</span>':''))+'</div></div>'+
     '<div><button class="btn-light" data-load="'+esc(it.name)+'">应用</button> '+
     '<button class="btn-ghost" data-del="'+esc(it.name)+'">删除</button></div>';
   box.appendChild(d);
 });
 $('savedMask').classList.add('show');
 box.querySelectorAll('[data-load]').forEach(b=>b.onclick=()=>{
   const it=data.saved.find(x=>x.name===b.dataset.load),p=it.params||{};
   $('query').value=p.query||'';$('sort').value=p.sort||'relevance';
   const f=p.filters||{};$('fExt').value=f.ext||'';$('fCat').value=f.category||'';
   $('fMin').value=f.size_min||'';$('fMax').value=f.size_max||'';
   $('fDateFrom').value=f.date_from||'';$('fDateTo').value=f.date_to||'';$('fPath').value=f.path_contains||'';
   if(p.mode){$('modes').querySelectorAll('.chip').forEach(x=>x.classList.toggle('on',x.dataset.mode===p.mode));mode=p.mode;}
   if(p.scope){$('scopes').querySelectorAll('.chip').forEach(x=>x.classList.toggle('on',x.dataset.scope===p.scope));scope=p.scope;}
   const st=p.searchType||'title';
   $('searchTabs').querySelectorAll('.tab').forEach(x=>x.classList.toggle('on',x.dataset.type===st));searchType=st;
   $('titleOpts').style.display=st==='content'?'none':'block';$('titleHint').style.display=st==='content'?'none':'block';
   $('contentHint').style.display=st==='content'?'block':'none';
   document.querySelectorAll('.snip-col').forEach(x=>x.style.display=st==='content'?'':'none');
   $('savedMask').classList.remove('show');doSearch();
 });
 box.querySelectorAll('[data-del]').forEach(b=>b.onclick=async()=>{
   await api('/api/saved/delete',{name:b.dataset.del});b.closest('.saved-item').remove();
 });
};
$('btnSavedClose').onclick=()=>$('savedMask').classList.remove('show');

/* 重复文件检测 */
$('btnDup').onclick=async()=>{
 if(!taskId){alert('请先扫描或搜索');return;}
 const area=$('dupArea');area.style.display='block';area.innerHTML='正在启动重复文件检测…';
 const r=await api('/api/dup/start',{id:taskId,min_size:'1MB'});
 if(r.error){area.textContent=r.error;return;}
 dupId=r.id;pollDup();
};
async function pollDup(){
 const s=await api('/api/dup/status?id='+encodeURIComponent(dupId));
 if(s.error){$('dupArea').textContent=s.error;return;}
 const area=$('dupArea');
 if(s.status==='running'||s.status==='pending'){
  area.innerHTML='🔄 '+esc(s.status_text)+'（'+s.scanned+'/'+s.total+'）';
  setTimeout(pollDup,600);return;
 }
 if(s.status==='done'){
  let h='<b>重复文件检测完成</b>：'+s.group_count+' 组重复，涉及 '+s.dup_files+
   ' 个文件，可释放约 <b>'+fmtSize(s.wasted)+'</b>　<button class="btn-ghost" id="dupClose">收起</button>';
  if(s.groups.length){
   h+='<div style="max-height:220px;overflow:auto;margin-top:8px;font-size:12px">'+
    s.groups.slice(0,40).map(g=>'· '+fmtSize(g.size)+' × '+g.files.length+' 个：'+
      esc(g.files[0].name)+' 等').join('<br>')+'</div>';
  }
  area.innerHTML=h;$('dupClose').onclick=()=>area.style.display='none';
 }else area.textContent=s.status_text;
}

/* 导出 */
$('btnExport').onclick=()=>{if(!taskId){alert('请先扫描或搜索');return;}
 expSource='';
 const sel=selectedResPaths().length;
 $('expInfo').textContent=(searchType==='content')
  ? (sel?'正文搜索：将导出结果区已选的 '+sel+' 个文件；清空选择则导出全部命中。'
        :'正文搜索：将导出全部正文命中结果（可勾选文件后只导出勾选的）。')
  : '';
 $('exportMask').classList.add('show');};
$('btnExpClose').onclick=()=>$('exportMask').classList.remove('show');
$('btnBrowseExp').onclick=()=>openBrowse('export');
$('btnDoExport').onclick=async()=>{
 if(!expDir){alert('请选择输出文件夹');return;}
 $('expInfo').textContent='正在导出…';
 let payload;
 if(expSource==='type'){
  // 功能1类型列表导出：可导出当前勾选，或整个类型
  const sel=selectedTypePaths();
  payload={source:'type',id:taskId,ext:typeExt,paths_only:sel,fmt:$('expFmt').value,output_dir:expDir};
 }else if(searchType==='content'){
  // 正文搜索结果导出：有勾选则只导勾选，否则导全部命中
  if(!contentId){$('expInfo').textContent='请先执行一次正文搜索';return;}
  payload={source:'content',id:taskId,content_id:contentId,
           paths_only:selectedResPaths(),fmt:$('expFmt').value,output_dir:expDir};
 }else{
  const base=lastParams?Object.assign({},lastParams):collectParams(0);
  base.offset=0;base.limit=1000000;
  payload=Object.assign(base,{fmt:$('expFmt').value,output_dir:expDir});
 }
 const r=await api('/api/export',payload);
 if(r.error){$('expInfo').textContent=r.error;return;}
 expSource='';
 $('expInfo').innerHTML='✓ 已导出 '+r.rows+' 条到<br>'+esc(r.path);
};

/* 文件夹浏览弹窗（含关键词跨层级模糊定位） */
function resetBrowseView(){
 $('browseResults').style.display='none';$('browseResults').innerHTML='';
 $('browseSearchInfo').textContent='';
 $('crumb').style.display='';$('dirList').style.display='';$('filePreview').style.display='';
}
function openBrowse(purpose){browseCtx={path:'',purpose:purpose};
 $('browseKw').value='';resetBrowseView();
 $('exportMask').classList.remove('show');$('browseMask').classList.add('show');loadBrowse('');}
async function loadBrowse(path){
 const r=await api('/api/browse?path='+encodeURIComponent(path));
 if(r.error){msg('filePreview',r.error,'var(--red)');return;}
 browseCtx.path=r.path;$('crumb').textContent=r.is_root?'选择硬盘：':('当前：'+(r.path||''));
 const list=$('dirList');list.innerHTML='';
 r.dirs.forEach(d=>{const el=document.createElement('div');el.className='d';
  el.textContent=(d.hidden?'🔒 ':'📁 ')+d.name;el.title=d.path;
  el.onclick=()=>loadBrowse(d.path);list.appendChild(el);});
 $('filePreview').textContent=r.is_root?'点击盘符进入逐级选择，或直接在上方输入关键词跨层级查找。':
  ('本层文件 '+r.files_total+' 个'+(r.files.length?'，预览：'+r.files.slice(0,8).map(f=>f.name).join('、'):''));
 $('btnPickHere').style.display=r.is_root?'none':'';
}
async function doBrowseSearch(kw){
 const box=$('browseResults');
 box.style.display='block';box.innerHTML='<div class="tip">正在查找，稍候…</div>';
 $('crumb').style.display='none';$('dirList').style.display='none';$('filePreview').style.display='none';
 const r=await api('/api/browse/search?path='+encodeURIComponent(browseCtx.path)+'&kw='+encodeURIComponent(kw));
 if(r.error){box.innerHTML='<div class="tip" style="color:var(--red)">'+esc(r.error)+'</div>';return;}
 const list=r.dirs||[];
 $('browseSearchInfo').textContent='找到 '+list.length+' 个'+(r.truncated?'（范围较大，已截断）':'');
 if(!list.length){box.innerHTML='<div class="tip">没有匹配的文件夹。支持乱序模糊，例如「二选」也能匹配「二十年精选收藏」。</div>';return;}
 box.innerHTML='';
 list.forEach(d=>{
  const el=document.createElement('div');el.className='r';
  el.innerHTML='<span class="pick">选这里</span>📁 '+esc(d.name)+'<span class="p">'+esc(d.parent)+'</span>';
  el.onclick=e=>{
   if(e.target && e.target.classList.contains('pick')){pickBrowsePath(d.path);return;}
   $('browseKw').value='';resetBrowseView();loadBrowse(d.path);
  };
  box.appendChild(el);
 });
}
$('browseKw').oninput=()=>{
 clearTimeout(browseSearchTimer);
 const kw=$('browseKw').value.trim();
 if(!kw){resetBrowseView();return;}
 browseSearchTimer=setTimeout(()=>doBrowseSearch(kw),300);
};
$('browseKw').onkeydown=e=>{
 if(e.key==='Enter'){e.preventDefault();clearTimeout(browseSearchTimer);
  const kw=$('browseKw').value.trim();if(kw)doBrowseSearch(kw);}
};
$('btnUp').onclick=async()=>{
 const r=await api('/api/browse?path='+encodeURIComponent(browseCtx.path));
 $('browseKw').value='';resetBrowseView();loadBrowse(r.parent||'');
};
function pickBrowsePath(p){
 if(!p)return;
 if(browseCtx.purpose==='export'){expDir=p;$('expDir').value=p;$('browseMask').classList.remove('show');
  $('exportMask').classList.add('show');}
 else if(browseCtx.purpose==='search'){$('searchPath').value=p;$('browseMask').classList.remove('show');}
 else if(browseCtx.purpose==='paste'){pasteDir=p;$('pasteDir').value=p;$('browseMask').classList.remove('show');
  $('pasteMask').classList.add('show');}
 else{chooseScanPath(p);$('browseMask').classList.remove('show');}
}
$('btnPickHere').onclick=()=>pickBrowsePath(browseCtx.path);
$('btnBrowseClose').onclick=()=>$('browseMask').classList.remove('show');

/* ============ 右侧文件预览 ============ */
let pvPath='', pvKind='', pvPage=1, pvPages=1, pvToken=0, pvTextAll='';
function pvSetOpen(open){
 $('app').classList.toggle('no-preview',!open);
 $('pvFloat').style.display=open?'none':'block';
 if(open){try{localStorage.setItem('fsp_preview_open','1');}catch(e){}}
 else{try{localStorage.setItem('fsp_preview_open','0');}catch(e){}}
}
function pvResetBody(tip){
 $('pvName').textContent=tip||'未选择文件';
 $('pvBar').style.display='none';
 $('pvBody').innerHTML='<div class="pv-empty"><span class="ico">🖼️</span>'+
  '点击搜索结果或类型列表中的<b>文件名 / 「预览」按钮</b>，<br>'+
  '可预览的文件会在此显示内容。<br>图片、PDF 按原样显示；Office 优先用本机 Microsoft/WPS 还原原排版；<br>未安装时自动回退内置渲染或文字预览。</div>';
}
async function showPreview(path, page){
 if(!path)return;
 page=page||1;
 pvSetOpen(true);
 const token=++pvToken;
 $('pvName').textContent='读取中…';
 $('pvBar').style.display='none';
 $('pvBody').innerHTML='<div class="pv-empty"><span class="ico">⏳</span>正在读取预览…</div>';
 let r;
 const pvQuery=($('query').value||'').trim();
 try{r=await api('/api/preview?p='+encodeURIComponent(path)+'&page='+page+'&query='+encodeURIComponent(pvQuery));}
 catch(e){r={ok:false,reason:'预览请求失败：'+e};}
 if(token!==pvToken)return;   // 防止快速连点时旧请求覆盖新结果
 if(!r||!r.ok){
  pvPath=path;pvKind='none';
  $('pvName').textContent=r&&r.name?r.name:'无法预览';
  $('pvBar').style.display='none';
  let sz=r&&r.size!=null?fmtSize(r.size):'-';
  let mt=r&&r.mtime||'-';
  $('pvBody').innerHTML='<div class="pv-unsupported"><span class="ico">📄</span>'+
   '<div style="color:#374151;font-weight:600">'+esc(r&&r.name||'文件')+'</div>'+
   '<div style="font-size:12px;margin:4px 0 10px">'+sz+'　'+esc(mt)+'</div>'+
   '<div>'+esc((r&&r.reason)||'此文件类型不支持预览')+'</div>'+
   '<button class="btn-light pv-open-btn" id="pvUnsupportedOpen">用系统默认程序打开 ↗</button></div>';
  const ob=$('pvUnsupportedOpen');
  if(ob)ob.onclick=()=>api('/api/open?p='+encodeURIComponent(path));
  return;
 }
 pvPath=path;pvKind=r.kind;pvPage=r.page||page||1;pvPages=r.pages||1;pvTextAll=r.text||'';
 $('pvName').textContent=r.name;
 const meta='<div class="pv-meta">'+esc(r.path)+'<br>'+fmtSize(r.size||0)+'　'+
  esc(r.mtime||'')+'　·　'+esc(r.ext||'无后缀')+'</div>';
 if(r.kind==='image'){
  $('pvBar').style.display='flex';
  $('pvPageInfo').textContent='图片预览';
  $('btnPvPrev').style.display='none';$('btnPvNext').style.display='none';
  $('pvBody').innerHTML='<div class="pv-imgwrap"><img id="pvImg" alt="图片预览"></div>'+meta;
  $('pvImg').src=r.url+'&_='+Date.now();
  $('pvImg').onerror=()=>{$('pvBody').innerHTML='<div class="pv-unsupported"><span class="ico">⚠️</span>图片加载失败</div>'+meta;};
 }else if(r.kind==='pdf'){
  $('pvBar').style.display='flex';
  $('btnPvPrev').style.display='';$('btnPvNext').style.display='';
  $('pvPageInfo').textContent='PDF 第 '+pvPage+' / '+pvPages+' 页';
  $('pvBody').innerHTML='<div class="pv-imgwrap"><img id="pvImg" alt="PDF 页面预览"></div>'+meta;
  $('pvImg').src=r.url+'&_='+Date.now();
  $('pvImg').onerror=()=>{$('pvBody').innerHTML='<div class="pv-unsupported"><span class="ico">⚠️</span>PDF 页面渲染失败</div>'+meta;};
 }else if(r.kind==='office'){
  // 新版 Office：后端已渲染成带样式的完整 HTML（含搜索词高亮），用 iframe 展示原排版
  $('pvBar').style.display='flex';
  const multi=(pvPages>1);
  $('btnPvPrev').style.display=multi?'':'none';$('btnPvNext').style.display=multi?'':'none';
  const kindName=r.ext==='pptx'?'幻灯片':(r.ext==='xlsx'?'工作表':'页');
  $('pvPageInfo').textContent=(r.ext==='docx')?'原排版预览':(kindName+' '+pvPage+' / '+pvPages);
  const note=(r.note?'<div class="pv-note">'+esc(r.note)+'　·　搜索词已高亮</div>':'');
  $('pvBody').innerHTML=note+'<iframe class="pv-iframe" id="pvFrame" title="原排版预览"></iframe>'+meta;
  const html='data:text/html;charset=utf-8;base64,'+r.html_b64;
  const fr=$('pvFrame');
  if(fr)fr.src=html;
 }else if(r.kind==='text'){
  $('pvBar').style.display='flex';
  $('btnPvPrev').style.display='none';$('btnPvNext').style.display='none';
  $('pvPageInfo').textContent='文字预览';
  const q=($('query').value||'').trim();
  const note=(r.note?'<div class="pv-note">'+esc(r.note)+'</div>':'')+
   (r.truncated?'<div class="pv-note">内容过长，仅显示前 '+Math.round(r.text.length/1000)+'K 字符</div>':'');
  $('pvBody').innerHTML=note+'<pre class="pv-text" id="pvText"></pre>'+meta;
  const txt=pvTextAll||'（文件为空）';
  const terms=q?q.split(/[\s|]+/).filter(x=>x&&!x.startsWith('-')&&!x.startsWith('!')):[];
  if(terms.length){
   // 先收集所有命中区间（原文坐标），再按区间对原文分段 esc，避免标签注入
   const ranges=[];
   terms.forEach(t=>{
    let from=0,idx;const low=txt.toLowerCase(),lt=t.toLowerCase();
    while((idx=low.indexOf(lt,from))>=0){ranges.push([idx,idx+t.length]);from=idx+t.length;
     if(ranges.length>2000)break;}
   });
   ranges.sort((a,b)=>a[0]-b[0]);
   const merged=[];ranges.forEach(rg=>{
    if(merged.length&&rg[0]<=merged[merged.length-1][1])merged[merged.length-1][1]=Math.max(merged[merged.length-1][1],rg[1]);
    else merged.push(rg.slice());});
   let out='',pos=0;
   merged.forEach(rg=>{out+=esc(txt.slice(pos,rg[0]))+'<mark>'+esc(txt.slice(rg[0],rg[1]))+'</mark>';pos=rg[1];});
   out+=esc(txt.slice(pos));
   $('pvText').innerHTML=out;
  }else{
   $('pvText').textContent=txt;
  }
 }
}
function pvFlip(delta){
 if(pvKind!=='pdf'&&pvKind!=='office')return;
 const np=pvPage+delta;
 if(np<1||np>pvPages)return;
 showPreview(pvPath, np);
}
$('btnPvPrev').onclick=()=>pvFlip(-1);
$('btnPvNext').onclick=()=>pvFlip(1);
$('btnPvOpen').onclick=()=>{if(pvPath)api('/api/open?p='+encodeURIComponent(pvPath));};
$('btnPvClose').onclick=()=>pvSetOpen(false);
$('pvFloat').onclick=()=>pvSetOpen(true);
try{if(localStorage.getItem('fsp_preview_open')==='0')pvSetOpen(false);}catch(e){}
/* 拖动分隔条调节左右宽度 */
(function(){
 const sp=$('splitter');let dragging=false;
 sp.addEventListener('mousedown',e=>{dragging=true;sp.classList.add('drag');
  document.body.style.cursor='col-resize';document.body.style.userSelect='none';e.preventDefault();});
 document.addEventListener('mousemove',e=>{
  if(!dragging)return;
  const w=Math.max(320,Math.min(window.innerWidth-360,window.innerWidth-e.clientX));
  $('previewPane').style.setProperty('--pv-w',w+'px');
 });
 document.addEventListener('mouseup',()=>{
  if(!dragging)return;dragging=false;sp.classList.remove('drag');
  document.body.style.cursor='';document.body.style.userSelect='';
  try{localStorage.setItem('fsp_pv_width',$('previewPane').style.getPropertyValue('--pv-w'));}catch(e){}
 });
 try{const w=localStorage.getItem('fsp_pv_width');
  if(w)$('previewPane').style.setProperty('--pv-w',w);}catch(e){}
})();

$('btnQuit').onclick=async()=>{await api('/api/quit',{});document.body.innerHTML=
 '<div style="padding:60px;text-align:center;color:#6b7280">程序已退出，可以直接关闭此窗口。</div>';};
</script>
</body></html>
"""


def main():
    global SERVER
    load_data()
    port = PORT
    # 可选的测试/多开参数：--port 8901（普通用户双击时不需要）
    try:
        if "--port" in sys.argv:
            port = int(sys.argv[sys.argv.index("--port") + 1])
    except Exception:
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
    print("全盘文件名搜索工具已启动：", url)
    if "--no-browser" not in sys.argv:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n已退出")


if __name__ == "__main__":
    # PyInstaller onefile 下用多进程池必须调用，否则子进程会重新进入 main 导致递归
    multiprocessing.freeze_support()
    main()
