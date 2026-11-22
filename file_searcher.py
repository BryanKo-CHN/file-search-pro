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
import string
import fnmatch
import hashlib
import shutil
import threading
import webbrowser
from collections import deque, defaultdict
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs

PORT = 8805
TASKS = {}          # 扫描任务
DUP_JOBS = {}       # 重复检测任务
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


# ----------------------------- HTTP 服务 -----------------------------

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
        elif u.path == "/api/progress":
            task = TASKS.get(q.get("id", [""])[0])
            self._send(200, task.snapshot() if task else {"error": "任务不存在，请重新扫描"})
        elif u.path == "/api/dup/status":
            job = DUP_JOBS.get(q.get("id", [""])[0])
            self._send(200, job.snapshot() if job else {"error": "任务不存在"})
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

        elif u.path == "/api/export":
            task = TASKS.get(data.get("id", ""))
            if not task:
                self._send(200, {"error": "未找到扫描结果"})
                return
            # 取全部匹配结果（不分页）
            params = dict(data)
            params["offset"] = 0
            params["limit"] = 1000000
            total, rows, _took, _terms = search_files(task, params)
            if not rows:
                self._send(200, {"error": "没有可导出的结果"})
                return
            fmt = (data.get("fmt") or "html").lower()
            dest = (data.get("output_dir") or "").strip().strip('"')
            if not dest or not os.path.isdir(dest):
                self._send(200, {"error": "输出文件夹不存在：" + dest})
                return
            out = os.path.join(dest, f"搜索结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{fmt}")
            try:
                n = export_rows(rows, {"scope": task.path, "query": data.get("query", "")}, fmt, out)
                self._send(200, {"ok": True, "path": out, "rows": n, "size": os.path.getsize(out)})
            except Exception as e:
                self._send(200, {"error": "导出失败：" + str(e)})

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
<title>全盘文件名搜索工具</title>
<style>
:root{--bg:#f5f7fa;--card:#fff;--line:#e5e7eb;--text:#1f2937;--muted:#6b7280;
--blue:#2563eb;--blue-l:#eff6ff;--green:#059669;--amber:#b45309;--red:#dc2626;}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--text);
font-family:"Microsoft YaHei",-apple-system,system-ui,sans-serif;font-size:14px;line-height:1.5}
.wrap{max-width:1240px;margin:0 auto;padding:22px 18px 60px}
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
.btn-light{background:var(--blue-l);color:var(--blue)}
.btn-ghost{background:#f3f4f6;color:#374151}
.btn-green{background:#ecfdf5;color:var(--green)}
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
.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-top:14px}
.stat{background:#f8fafc;border:1px solid var(--line);border-radius:9px;padding:10px 12px}
.stat .k{color:var(--muted);font-size:12px}.stat .v{font-size:18px;font-weight:700;margin-top:2px}
.note{color:var(--muted);font-size:12px}
/* 搜索区 */
.search-bar{display:flex;gap:8px;align-items:center}
.qbox{flex:1;position:relative}
.qinput{width:100%;font-size:17px;padding:12px 14px 12px 40px;border:2px solid #dbe3ee;border-radius:10px}
.qinput:focus{outline:0;border-color:var(--blue);box-shadow:0 0 0 3px rgba(37,99,235,.1)}
.qicon{position:absolute;left:13px;top:50%;transform:translateY(-50%);color:#9ca3af;font-size:17px}
.modes{display:flex;gap:5px;margin-top:10px;flex-wrap:wrap}
.chip{padding:5px 12px;border-radius:20px;background:#f3f4f6;color:#4b5563;font-size:12px;cursor:pointer;border:1px solid transparent}
.chip.on{background:var(--blue-l);color:var(--blue);border-color:#bfdbfe;font-weight:600}
.adv{margin-top:12px;padding-top:12px;border-top:1px dashed var(--line);display:none}
.adv.show{display:block}
.fgrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px}
.fitem label{display:block;font-size:12px;color:var(--muted);margin-bottom:4px}
.fitem input,.fitem select{width:100%}
.hint{margin-top:10px;font-size:12px;color:var(--muted);background:#f8fafc;
padding:8px 11px;border-radius:8px;line-height:1.8}
.hint code{background:#eef2ff;color:#3730a3;padding:1px 5px;border-radius:4px}
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
.c-name{font-weight:600;max-width:360px;overflow:hidden;text-overflow:ellipsis}
.c-rel{color:var(--muted);max-width:320px;overflow:hidden;text-overflow:ellipsis}
mark{background:#fde68a;color:#7c2d12;border-radius:2px;padding:0 1px}
.tag{display:inline-block;background:var(--blue-l);color:var(--blue);font-size:11px;
padding:1px 7px;border-radius:10px}
.op{color:var(--blue);cursor:pointer;font-size:12px}
.pager{display:flex;gap:8px;align-items:center;justify-content:center;margin-top:12px;flex-wrap:wrap}
.pager button{padding:6px 11px;font-size:13px}
.empty{text-align:center;color:var(--muted);padding:40px 0}
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
.saved-item{display:flex;justify-content:space-between;align-items:center;padding:8px 10px;
border:1px solid var(--line);border-radius:8px;margin-bottom:7px}
.tagchip{display:inline-block;background:#f3f4f6;color:#4b5563;font-size:11px;padding:1px 7px;border-radius:10px;margin:2px}
.footer{text-align:center;color:var(--muted);font-size:12px;margin-top:22px}
</style></head><body>
<div class="wrap">
  <h1>🔍 全盘文件名搜索工具</h1>
  <div class="sub">逐层扫描硬盘/文件夹，按文件名智能模糊搜索 —— 记不清全名、记错格式都能找到。只读操作，绝不改动原文件。</div>

  <!-- 扫描 -->
  <div class="panel">
    <h2>1. 选择扫描范围</h2>
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
        <span style="color:var(--amber)">（数据盘多数种子/文件带隐藏属性，勾选会漏扫）</span></label>
      <div class="opt"><label>最大层数（0=扫到最内层）</label>
        <input type="number" id="maxDepth" value="0" min="0" max="99" style="width:76px"></div>
    </div>
    <div class="progress-wrap" id="progressWrap">
      <div class="bar"><i></i></div>
      <div class="status" id="status">准备中...</div>
      <div class="status" id="levelInfo" style="color:var(--blue);font-weight:600"></div>
    </div>
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
  </div>

  <!-- 搜索 -->
  <div class="panel">
    <h2>2. 关键词搜索</h2>
    <div class="search-bar">
      <div class="qbox"><span class="qicon">🔎</span>
        <input type="text" id="query" class="qinput" placeholder="输入文件名关键词，例如：中学生考点"
               autocomplete="off"></div>
      <button class="btn-primary" id="btnSearch" style="padding:12px 22px">搜索</button>
      <button class="btn-ghost" id="btnClear" style="padding:12px 14px">清空</button>
    </div>
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
    <div class="opts">
      <label class="check"><input type="checkbox" id="advToggle"> 高级筛选（大小 / 日期 / 后缀 / 路径）</label>
      <div class="opt"><label>排序</label>
        <select id="sort">
          <option value="relevance">相关度</option><option value="name">文件名</option>
          <option value="path">路径</option><option value="size">大小（大到小）</option>
          <option value="mtime">修改时间（新到旧）</option>
        </select></div>
      <button class="btn-light" id="btnSave">保存此搜索</button>
      <button class="btn-ghost" id="btnSaved">已保存的搜索</button>
    </div>
    <div class="adv" id="adv">
      <div class="fgrid">
        <div class="fitem"><label>后缀（如 pdf,docx）</label><input type="text" id="fExt" placeholder="留空=全部"></div>
        <div class="fitem"><label>分类</label><select id="fCat"><option value="">全部分类</option></select></div>
        <div class="fitem"><label>最小大小</label><input type="text" id="fMin" placeholder="如 10MB"></div>
        <div class="fitem"><label>最大大小</label><input type="text" id="fMax" placeholder="如 2GB"></div>
        <div class="fitem"><label>修改日期从</label><input type="date" id="fDateFrom"></div>
        <div class="fitem"><label>修改日期到</label><input type="date" id="fDateTo"></div>
        <div class="fitem"><label>路径包含</label><input type="text" id="fPath" placeholder="如 二十年精选收藏"></div>
      </div>
    </div>
    <div class="hint">
      💡 <b>多关键词</b>：空格分隔表示“同时包含”；<code>|</code> 表示“或”；<code>-词</code> 或 <code>!词</code> 表示“排除”。
      例：<code>中考 -答案</code>、<code>报告 | 总结</code>。<br>
      <b>智能模糊</b>：按顺序命中即可，如 <code>中学生考点</code> 能匹配「中学生<b>考试</b>考点汇总.doc」——
      所以即使记错格式（以为 PDF 其实是 DOC）也能搜到。<br>
      <b>通配符</b>（匹配整个文件名）：<code>*</code> 任意多个字符，<code>?</code> 单个字符，
      如 <code>*报告*.doc</code>（含“报告”的 doc 文件）。
    </div>
    <div class="quick" id="history"></div>
  </div>

  <!-- 结果 -->
  <div class="panel">
    <div class="res-head">
      <div class="res-info" id="resInfo">尚未搜索</div>
      <div class="input-row">
        <button class="btn-amber" id="btnDup">查找重复文件</button>
        <button class="btn-light" id="btnExport">导出结果…</button>
      </div>
    </div>
    <div id="dupArea" style="display:none;margin-bottom:12px;padding:12px;background:#fffbeb;
      border:1px solid #fde68a;border-radius:9px;font-size:13px;color:#92400e"></div>
    <div class="tbl-wrap">
      <table><thead><tr>
        <th data-sort="name" style="width:44px">#</th>
        <th data-sort="name">文件名</th><th>后缀</th><th>分类</th>
        <th data-sort="size">大小</th><th data-sort="mtime">修改时间</th>
        <th data-sort="path">相对路径</th><th>操作</th>
      </tr></thead><tbody id="resultBody"></tbody></table>
      <div class="empty" id="empty" style="display:none">没有找到匹配的文件</div>
    </div>
    <div class="pager" id="pager"></div>
  </div>

  <div class="footer">纯本地运行 · 扫描与搜索只读 · 数据不离开本机　
    <button class="btn-ghost" id="btnQuit" style="padding:5px 12px;font-size:12px">退出程序</button></div>
</div>

<!-- 文件夹浏览弹窗 -->
<div class="mask" id="browseMask"><div class="modal">
  <h3 id="browseTitle">选择扫描范围</h3>
  <div class="body">
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
      <select id="expFmt"><option value="html">HTML（可浏览，推荐）</option>
      <option value="csv">CSV（Excel 打开）</option><option value="txt">TXT 文本</option></select></div>
    <div class="input-row"><input type="text" id="expDir" class="pick-input" readonly placeholder="选择输出文件夹">
      <button class="btn-light" id="btnBrowseExp">浏览…</button></div>
    <div class="note" style="margin-top:8px" id="expInfo"></div>
  </div>
  <div class="foot"><button class="btn-primary" id="btnDoExport">导出</button>
    <button class="btn-ghost" id="btnExpClose">取消</button></div>
</div></div>

<!-- 保存的搜索弹窗 -->
<div class="mask" id="savedMask"><div class="modal" style="width:min(560px,95vw)">
  <h3>已保存的搜索（智能文件夹）</h3>
  <div class="body" id="savedList"></div>
  <div class="foot"><button class="btn-ghost" id="btnSavedClose">关闭</button></div>
</div></div>

<script>
const $=id=>document.getElementById(id);
let taskId=null, mode='smart', scope='stem', currentPage=0, pageSize=100,
    lastTotal=0, browseCtx={path:'',purpose:'scan'}, expDir='', dupId=null, pollTimer=null;
function esc(s){return String(s==null?'':s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
function fmtSize(n){n=Number(n)||0;const u=['B','KB','MB','GB','TB'];let i=0;
 while(n>=1024&&i<u.length-1){n/=1024;i++;}return (i===0?n:n.toFixed(2))+' '+u[i];}
async function api(p,d){
 if(d!==undefined){const r=await fetch(p,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(d)});return r.json();}
 const r=await fetch(p);return r.json();
}
function msg(id,text,color){const e=$(id);if(e){e.textContent=text;e.style.color=color||'';}}

/* 初始化 */
(async()=>{
 const dr=await api('/api/drives');
 $('quick').innerHTML='<span>快捷位置：</span>'+['Desktop','Documents','Downloads'].map(n=>
   '<button data-q="'+n+'">'+({Desktop:'桌面',Documents:'文档',Downloads:'下载'}[n])+'</button>').join('')+
   dr.drives.map(d=>'<button data-drive="'+d.path+'">'+d.path.slice(0,2)+'</button>').join('');
 $('quick').querySelectorAll('button').forEach(b=>b.onclick=async()=>{
   if(b.dataset.q){const r=await api('/api/quickpath?name='+b.dataset.q);if(r.path)choosePath(r.path);}
   else if(b.dataset.drive)choosePath(b.dataset.drive);
 });
 const cats=['办公文档','文本与数据','图片','视频','音频','压缩包与镜像','程序与脚本','种子与下载','编程开发','字体','其他'];
 $('fCat').innerHTML='<option value="">全部分类</option>'+cats.map(c=>'<option>'+c+'</option>').join('');
 const data=await api('/api/data');renderHistory(data.history||[]);
})();
function choosePath(p){$('scanPath').value=p;$('resInfo').textContent='尚未搜索';$('resultBody').innerHTML='';}

/* 模式/范围切换 */
$('modes').querySelectorAll('.chip').forEach(c=>c.onclick=()=>{
 $('modes').querySelectorAll('.chip').forEach(x=>x.classList.remove('on'));c.classList.add('on');mode=c.dataset.mode;
});
$('scopes').querySelectorAll('.chip').forEach(c=>c.onclick=()=>{
 $('scopes').querySelectorAll('.chip').forEach(x=>x.classList.remove('on'));c.classList.add('on');scope=c.dataset.scope;
});
$('advToggle').onchange=e=>$('adv').classList.toggle('show',e.target.checked);
$('query').addEventListener('keydown',e=>{if(e.key==='Enter')doSearch();});
$('btnClear').onclick=()=>{$('query').value='';['fExt','fMin','fMax','fPath','fDateFrom','fDateTo'].forEach(id=>$(id).value='');
 $('fCat').value='';doSearch();};
document.querySelectorAll('th[data-sort]').forEach(th=>th.onclick=()=>{
 const s=th.dataset.sort;if(['name','size','mtime','path'].includes(s)){$('sort').value=s;doSearch();}
});

/* 扫描 */
$('btnBrowseScan').onclick=()=>openBrowse('scan');
$('btnScan').onclick=async()=>{
 const p=$('scanPath').value.trim();if(!p){alert('请先选择扫描范围');return;}
 $('btnScan').disabled=true;$('btnCancel').disabled=false;$('btnSearch').disabled=true;
 $('progressWrap').style.display='block';$('stats').style.display='none';
 $('levelBars').style.display='none';$('skipNote').style.display='none';
 $('status').textContent='正在启动...';$('levelInfo').textContent='准备按层扫描：根目录 → 第1层 → 第2层 → …';
 const r=await api('/api/scan',{path:p,max_depth:$('maxDepth').value,
   exclude_hidden:$('excHidden').checked,follow_links:$('followLinks').checked});
 if(r.error){alert(r.error);$('btnScan').disabled=false;$('btnCancel').disabled=true;return;}
 taskId=r.id;pollTimer=setInterval(poll,300);
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
  if(s.status==='done')doSearch();
 }
}
function collectParams(offset){
 return {id:taskId,query:$('query').value.trim(),mode:mode,scope:scope,
   sort:$('sort').value,offset:offset,limit:pageSize,
   filters:{ext:$('fExt').value,category:$('fCat').value,size_min:$('fMin').value,
     size_max:$('fMax').value,date_from:$('fDateFrom').value,date_to:$('fDateTo').value,
     path_contains:$('fPath').value}};
}
function doSearch(offset){offset=offset||0;if(!taskId){alert('请先扫描');return;}
 currentPage=Math.floor(offset/pageSize);
 $('resInfo').innerHTML='<b>搜索中…</b>';
 api('/api/search',collectParams(offset)).then(r=>{
  if(r.error){msg('resInfo',r.error,'var(--red)');return;}
  lastTotal=r.total;renderResults(r.rows,r.terms);renderPager();
  $('resInfo').innerHTML='找到 <b>'+r.total+'</b> 个文件'+
   (r.query?'':'')+'　<span style="color:var(--muted)">用时 '+r.took+' ms</span>';
 });
}
function highlight(name,terms){
 let h=esc(name);if(!terms||!terms.length)return h;
 const sorted=terms.slice().sort((a,b)=>b.length-a.length);
 for(const t of sorted){
  if(!t)continue;
  try{const re=new RegExp('('+t.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')+')','gi');
   h=h.replace(re,'<mark>$1</mark>');}catch(e){}
 }
 // 避免嵌套 mark 造成的重复
 return h.replace(/<mark>(<mark>(.*?)<\/mark>)<\/mark>/g,'<mark>$2</mark>');
}
function renderResults(rows,terms){
 const body=$('resultBody');body.innerHTML='';
 $('empty').style.display=rows.length?'none':'block';
 rows.forEach((r,i)=>{
  const tr=document.createElement('tr');
  tr.innerHTML='<td style="color:#9ca3af">'+(currentPage*pageSize+i+1)+'</td>'+
   '<td class="c-name" title="'+esc(r.path)+'">'+highlight(r.name,terms)+'</td>'+
   '<td><span class="tag">'+esc(r.ext||'-')+'</span></td>'+
   '<td>'+esc(r.category)+'</td>'+
   '<td>'+fmtSize(r.size)+'</td>'+
   '<td>'+esc(r.mtime)+'</td>'+
   '<td class="c-rel" title="'+esc(r.rel)+'">'+esc(r.rel)+'</td>'+
   '<td><span class="op" data-open="'+esc(r.path)+'">定位</span></td>';
  body.appendChild(tr);
 });
 body.querySelectorAll('[data-open]').forEach(el=>el.onclick=()=>{
  api('/api/reveal?p='+encodeURIComponent(el.dataset.open));
 });
}
function renderPager(){
 const pages=Math.ceil(lastTotal/pageSize),p=$('pager');p.innerHTML='';
 if(pages<=1)return;
 const mk=(label,off,dis)=>{const b=document.createElement('button');b.textContent=label;
  b.className='btn-ghost';b.disabled=dis;b.onclick=()=>doSearch(off);p.appendChild(b);};
 mk('首页',0,currentPage===0);mk('上一页',Math.max(0,currentPage-1)*pageSize,currentPage===0);
 let s=Math.max(0,currentPage-2),e=Math.min(pages-1,currentPage+2);
 for(let i=s;i<=e;i++){const b=document.createElement('button');
  b.textContent=(i+1);b.className=i===currentPage?'btn-primary':'btn-ghost';
  b.onclick=()=>doSearch(i*pageSize);p.appendChild(b);}
 mk('下一页',Math.min(pages-1,currentPage+1)*pageSize,currentPage>=pages-1);
 mk('末页',(pages-1)*pageSize,currentPage>=pages-1);
}

/* 历史 / 保存 */
function renderHistory(list){
 $('history').innerHTML=list.length?('<span>最近搜索：</span>'+list.slice(0,8).map(q=>
   '<button data-h="'+esc(q)+'">'+esc(q)+'</button>').join('<button id="clearHist">清空</button>')):'';
 $('history').querySelectorAll('[data-h]').forEach(b=>b.onclick=()=>{$('query').value=b.dataset.h;doSearch();});
 if($('clearHist'))$('clearHist').onclick=async()=>{await api('/api/data/clear_history',{});
   renderHistory([]);};
}
$('btnSave').onclick=async()=>{
 const name=prompt('给这个搜索起个名字：',$('query').value.trim()||'未命名搜索');
 if(!name)return;
 const r=await api('/api/saved/add',{name:name,params:collectParams(0)});
 alert(r.error||('已保存：'+name));
};
$('btnSaved').onclick=async()=>{
 const data=await api('/api/data');const box=$('savedList');box.innerHTML='';
 if(!(data.saved||[]).length){box.innerHTML='<div class="empty">还没有保存的搜索</div>';}
 else data.saved.forEach(it=>{
   const d=document.createElement('div');d.className='saved-item';
   const p=it.params||{};
   d.innerHTML='<div><b>'+esc(it.name)+'</b><div>'+
     ((p.query?'<span class="tagchip">词: '+esc(p.query)+'</span>':'')+
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
   $('savedMask').classList.remove('show');doSearch();
 });
 box.querySelectorAll('[data-del]').forEach(b=>b.onclick=async()=>{
   await api('/api/saved/delete',{name:b.dataset.del});b.closest('.saved-item').remove();
 });
};
$('btnSavedClose').onclick=()=>$('savedMask').classList.remove('show');

/* 重复文件检测 */
$('btnDup').onclick=async()=>{
 if(!taskId){alert('请先扫描');return;}
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
$('btnExport').onclick=()=>{if(!taskId){alert('请先扫描');return;}
 $('exportMask').classList.add('show');};
$('btnExpClose').onclick=()=>$('exportMask').classList.remove('show');
$('btnBrowseExp').onclick=()=>openBrowse('export');
$('btnDoExport').onclick=async()=>{
 if(!expDir){alert('请选择输出文件夹');return;}
 $('expInfo').textContent='正在导出…';
 const r=await api('/api/export',Object.assign(collectParams(0),{fmt:$('expFmt').value,output_dir:expDir}));
 if(r.error){$('expInfo').textContent=r.error;return;}
 $('expInfo').innerHTML='✓ 已导出 '+r.rows+' 条到<br>'+esc(r.path);
};

/* 文件夹浏览弹窗 */
function openBrowse(purpose){browseCtx={path:'',purpose:purpose};
 $('exportMask').classList.remove('show');$('browseMask').classList.add('show');loadBrowse('');}
async function loadBrowse(path){
 const r=await api('/api/browse?path='+encodeURIComponent(path));
 if(r.error){msg('filePreview',r.error,'var(--red)');return;}
 browseCtx.path=r.path;$('crumb').textContent=r.is_root?'选择硬盘：':('当前：'+(r.path||''));
 const list=$('dirList');list.innerHTML='';
 r.dirs.forEach(d=>{const el=document.createElement('div');el.className='d';
  el.textContent=(d.hidden?'🔒 ':'📁 ')+d.name;el.title=d.path;
  el.onclick=()=>loadBrowse(d.path);list.appendChild(el);});
 $('filePreview').textContent=r.is_root?'点击盘符进入，逐级选择文件夹。':
  ('本层文件 '+r.files_total+' 个'+(r.files.length?'，预览：'+r.files.slice(0,8).map(f=>f.name).join('、'):''));
 $('btnPickHere').style.display=r.is_root?'none':'';
}
$('btnUp').onclick=async()=>{
 const r=await api('/api/browse?path='+encodeURIComponent(browseCtx.path));
 loadBrowse(r.parent||'');
};
$('btnPickHere').onclick=()=>{
 const p=browseCtx.path;if(!p)return;
 if(browseCtx.purpose==='export'){expDir=p;$('expDir').value=p;$('browseMask').classList.remove('show');
  $('exportMask').classList.add('show');}
 else{choosePath(p);$('browseMask').classList.remove('show');}
};
$('btnBrowseClose').onclick=()=>$('browseMask').classList.remove('show');

$('btnQuit').onclick=async()=>{await api('/api/quit',{});document.body.innerHTML=
 '<div style="padding:60px;text-align:center;color:#6b7280">程序已退出，可以直接关闭此窗口。</div>';};
</script>
</body></html>"""


def main():
    global SERVER
    load_data()
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
    main()
