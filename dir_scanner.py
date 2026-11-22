#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
目录扫描导出工具
扫描指定硬盘/文件夹下的所有文件，导出为 TXT / CSV / HTML 目录清单。

运行方式：
    python dir_scanner.py            # 启动本地网页界面（自动打开浏览器）
    python dir_scanner.py --no-browser
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

PORT = 8765
TASKS = {}
SERVER = None  # 当前 HTTP 服务实例，供退出接口调用


# ----------------------------- 工具函数 -----------------------------

def is_hidden(path):
    """判断文件/文件夹是否为隐藏项（兼容 Windows 隐藏属性与点开头命名）。"""
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


def list_drives():
    """列出本机可用磁盘分区及其容量信息。"""
    out = []
    if os.name == "nt":
        for letter in string.ascii_uppercase:
            p = letter + ":\\"
            if os.path.exists(p):
                try:
                    u = shutil.disk_usage(p)
                    out.append({"path": p, "total": u.total, "used": u.used, "free": u.free})
                except Exception:
                    out.append({"path": p, "total": 0, "used": 0, "free": 0})
    else:
        out.append({"path": "/", "total": 0, "used": 0, "free": 0})
    return out


def human_size(num):
    """把字节数转成可读格式。"""
    try:
        n = float(num)
    except Exception:
        return "0 B"
    for unit in ["B", "KB", "MB", "GB", "TB", "PB"]:
        if n < 1024.0 or unit == "PB":
            return f"{int(n)} B" if unit == "B" else f"{n:.2f} {unit}"
        n /= 1024.0
    return f"{n:.2f} PB"


def make_record(full, typ, root_path):
    """生成一条文件/文件夹记录。"""
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
        st = os.stat(full)
        if typ == "file":
            size = st.st_size
        mtime = datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        pass
    ext = os.path.splitext(name)[1].lstrip(".").lower()
    return {
        "type": typ,
        "name": name,
        "path": full,
        "rel": rel,
        "size": size,
        "mtime": mtime,
        "ext": ext,
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
        self.dirs = []
        self.errors = []
        self.start = time.time()
        self.end = None
        self.cancel = False

    def snapshot(self):
        return {
            "id": self.id,
            "path": self.path,
            "status": self.status,
            "status_text": self.status_text,
            "dirs": self.count_dir,
            "files": self.count_file,
            "total_size": self.total_size,
            "errors": len(self.errors),
            "elapsed": round((self.end or time.time()) - self.start, 2),
            "preview": (self.dirs[:40] + self.files[:160]),
        }


def parse_ext_set(text):
    """把 'txt,log,.csv' 解析成 {'txt','log','csv'}。"""
    result = set()
    for part in (text or "").replace("，", ",").split(","):
        p = part.strip().lstrip(".").lower()
        if p:
            result.add(p)
    return result


def scan_worker(task):
    opts = task.opts
    max_depth = int(opts.get("max_depth", 0) or 0)
    include_hidden = bool(opts.get("include_hidden", False))
    inc_exts = parse_ext_set(opts.get("include_ext", ""))
    exc_exts = parse_ext_set(opts.get("exclude_ext", ""))
    base_depth = task.path.rstrip("\\/").count(os.sep)

    def match_ext(name):
        ext = os.path.splitext(name)[1].lstrip(".").lower()
        key = ext or name.lower()
        if inc_exts and ext not in inc_exts and key not in inc_exts:
            return False
        if exc_exts and (ext in exc_exts or key in exc_exts):
            return False
        return True

    try:
        for root, dirs, files in os.walk(
            task.path, topdown=True, onerror=lambda e: task.errors.append(str(e))
        ):
            if task.cancel:
                task.status = "cancelled"
                task.end = time.time()
                task.status_text = "已取消"
                return

            # 最大深度限制
            if max_depth and (root.count(os.sep) - base_depth) >= max_depth:
                dirs[:] = []

            # 隐藏项过滤
            if not include_hidden:
                dirs[:] = [d for d in dirs if not is_hidden(os.path.join(root, d))]
                files = [f for f in files if not is_hidden(os.path.join(root, f))]

            for d in dirs:
                task.dirs.append(make_record(os.path.join(root, d), "dir", task.path))
                task.count_dir += 1

            for f in files:
                if not match_ext(f):
                    continue
                rec = make_record(os.path.join(root, f), "file", task.path)
                task.files.append(rec)
                task.count_file += 1
                task.total_size += rec["size"]

            task.status_text = (
                f"扫描中... 已发现 {task.count_file} 个文件 / {task.count_dir} 个文件夹"
            )
    except Exception as e:
        task.errors.append(str(e))

    task.status = "done"
    task.end = time.time()
    task.status_text = f"扫描完成：{task.count_file} 个文件，{task.count_dir} 个文件夹"


# ----------------------------- 导出功能 -----------------------------

def build_rows(task, include_dirs):
    rows = []
    if include_dirs:
        rows.extend(task.dirs)
    rows.extend(task.files)
    rows.sort(key=lambda r: r["path"].lower())
    return rows


def export_txt(task, out_path, include_dirs):
    rows = build_rows(task, include_dirs)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("=" * 72 + "\n")
        f.write("目录扫描清单\n")
        f.write("=" * 72 + "\n")
        f.write(f"扫描路径  : {task.path}\n")
        f.write(f"生成时间  : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"文件数量  : {task.count_file}\n")
        f.write(f"文件夹数量: {task.count_dir}\n")
        f.write(f"文件总大小: {human_size(task.total_size)} ({task.total_size} 字节)\n")
        f.write(f"读取错误  : {len(task.errors)}\n")
        f.write("-" * 72 + "\n")
        f.write(f"{'序号':<8}{'类型':<8}{'大小':<14}{'修改时间':<21}{'相对路径'}\n")
        f.write("-" * 72 + "\n")
        for i, r in enumerate(rows, 1):
            typ = "文件夹" if r["type"] == "dir" else "文件"
            size = "" if r["type"] == "dir" else human_size(r["size"])
            f.write(
                f"{i:<8}{typ:<8}{size:<14}{r['mtime']:<21}{r['rel']}\n"
            )
        f.write("-" * 72 + "\n")
        f.write(f"共 {len(rows)} 条记录\n")
    return len(rows)


def export_csv(task, out_path, include_dirs):
    rows = build_rows(task, include_dirs)
    with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["序号", "类型", "文件名", "相对路径", "完整路径", "大小(字节)",
                    "大小", "修改时间", "扩展名"])
        for i, r in enumerate(rows, 1):
            w.writerow([
                i,
                "文件夹" if r["type"] == "dir" else "文件",
                r["name"],
                r["rel"],
                r["path"],
                "" if r["type"] == "dir" else r["size"],
                "" if r["type"] == "dir" else human_size(r["size"]),
                r["mtime"],
                r["ext"],
            ])
    return len(rows)


def export_html(task, out_path, include_dirs):
    rows = build_rows(task, include_dirs)
    def esc(s):
        import html as _h
        return _h.escape(str(s))
    trs = []
    for i, r in enumerate(rows, 1):
        typ = "文件夹" if r["type"] == "dir" else "文件"
        size = "" if r["type"] == "dir" else human_size(r["size"])
        trs.append(
            "<tr><td class='num'>{i}</td><td><span class='tag {cls}'>{typ}</span></td>"
            "<td class='rel'>{rel}</td><td class='size'>{size}</td>"
            "<td class='time'>{mtime}</td></tr>".format(
                i=i, cls="dir" if r["type"] == "dir" else "file", typ=typ,
                rel=esc(r["rel"]), size=esc(size), mtime=esc(r["mtime"])
            )
        )
    doc = """<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<title>目录扫描清单</title>
<style>
 body{font-family:"Microsoft YaHei",Arial,sans-serif;margin:0;background:#f5f7fa;color:#1f2937;}
 .wrap{max-width:1100px;margin:24px auto;background:#fff;border-radius:12px;
       box-shadow:0 4px 20px rgba(0,0,0,.06);padding:28px;}
 h1{margin:0 0 4px;font-size:24px;color:#0f172a;}
 .meta{color:#64748b;font-size:13px;margin-bottom:20px;line-height:1.9;}
 .cards{display:flex;gap:14px;flex-wrap:wrap;margin-bottom:22px;}
 .card{flex:1;min-width:150px;background:#f8fafc;border:1px solid #e2e8f0;
       border-radius:10px;padding:16px;}
 .card .k{font-size:12px;color:#64748b;}
 .card .v{font-size:22px;font-weight:700;color:#0f172a;margin-top:6px;}
 table{width:100%;border-collapse:collapse;font-size:13px;}
 th,td{padding:9px 12px;text-align:left;border-bottom:1px solid #eef2f7;}
 th{background:#f8fafc;color:#475569;position:sticky;top:0;}
 tr:hover td{background:#fafbfc;}
 .num{color:#94a3b8;width:60px;}
 .tag{padding:2px 9px;border-radius:20px;font-size:12px;}
 .tag.dir{background:#e0f2fe;color:#0369a1;}
 .tag.file{background:#dcfce7;color:#15803d;}
 .rel{word-break:break-all;}
 .size{text-align:right;white-space:nowrap;color:#475569;}
 .time{white-space:nowrap;color:#64748b;}
 .foot{margin-top:18px;color:#94a3b8;font-size:12px;text-align:right;}
</style></head><body><div class="wrap">
<h1>目录扫描清单</h1>
<div class="meta">扫描路径：__PATH__<br>生成时间：__TIME__</div>
<div class="cards">
 <div class="card"><div class="k">文件数量</div><div class="v">__FILES__</div></div>
 <div class="card"><div class="k">文件夹数量</div><div class="v">__DIRS__</div></div>
 <div class="card"><div class="k">文件总大小</div><div class="v">__SIZE__</div></div>
 <div class="card"><div class="k">总记录数</div><div class="v">__ROWS__</div></div>
</div>
<table><thead><tr><th>序号</th><th>类型</th><th>相对路径</th><th>大小</th><th>修改时间</th></tr></thead>
<tbody>__ROWS_HTML__</tbody></table>
<div class="foot">由 目录扫描导出工具 生成</div>
</div></body></html>"""
    doc = (doc.replace("__PATH__", esc(task.path))
              .replace("__TIME__", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
              .replace("__FILES__", str(task.count_file))
              .replace("__DIRS__", str(task.count_dir))
              .replace("__SIZE__", human_size(task.total_size))
              .replace("__ROWS__", str(len(rows)))
              .replace("__ROWS_HTML__", "\n".join(trs)))
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(doc)
    return len(rows)


def do_export(task, fmt, out_path, include_dirs):
    fmt = fmt.lower().lstrip(".")
    if fmt == "txt":
        return export_txt(task, out_path, include_dirs)
    if fmt == "csv":
        return export_csv(task, out_path, include_dirs)
    if fmt in ("html", "htm"):
        return export_html(task, out_path, include_dirs)
    raise ValueError("不支持的导出格式: " + fmt)


# ----------------------------- Web 界面 -----------------------------

HTML_PAGE = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>目录扫描导出工具</title>
<style>
:root{
  --bg:#f5f7fa; --panel:#ffffff; --line:#e2e8f0; --text:#1f2937;
  --muted:#64748b; --primary:#2563eb; --primary-dark:#1d4ed8;
  --ok:#16a34a; --warn:#d97706; --soft:#f8fafc;
}
*{box-sizing:border-box;}
body{margin:0;background:var(--bg);color:var(--text);
     font-family:"Microsoft YaHei","Segoe UI",Arial,sans-serif;font-size:14px;}
.header{background:linear-gradient(120deg,#1e40af,#2563eb);color:#fff;padding:22px 0;}
.header .inner{max-width:1080px;margin:0 auto;padding:0 24px;}
.header h1{margin:0;font-size:22px;font-weight:700;}
.header p{margin:6px 0 0;opacity:.85;font-size:13px;}
.main{max-width:1080px;margin:22px auto;padding:0 24px 60px;}
.panel{background:var(--panel);border:1px solid var(--line);border-radius:12px;
       padding:20px;margin-bottom:18px;box-shadow:0 2px 10px rgba(15,23,42,.04);}
.panel h2{font-size:15px;margin:0 0 16px;color:#0f172a;
          border-left:3px solid var(--primary);padding-left:10px;}
label{font-size:13px;color:var(--muted);display:block;margin-bottom:6px;}
.input-row{display:flex;gap:10px;align-items:center;flex-wrap:wrap;}
input[type=text]{flex:1;min-width:260px;padding:10px 12px;border:1px solid var(--line);
   border-radius:8px;font-size:14px;outline:none;background:#fff;color:var(--text);}
input[type=text]:focus{border-color:var(--primary);box-shadow:0 0 0 3px rgba(37,99,235,.12);}
button{border:none;border-radius:8px;padding:10px 18px;font-size:14px;cursor:pointer;
       font-family:inherit;transition:.15s;}
.btn-primary{background:var(--primary);color:#fff;}
.btn-primary:hover{background:var(--primary-dark);}
.btn-light{background:#eef2ff;color:#3730a3;}
.btn-light:hover{background:#e0e7ff;}
.btn-ghost{background:#f1f5f9;color:#475569;}
.btn-ghost:hover{background:#e2e8f0;}
button:disabled{opacity:.5;cursor:not-allowed;}
.quick{margin-top:12px;display:flex;gap:8px;flex-wrap:wrap;align-items:center;}
.quick span{font-size:12px;color:var(--muted);margin-right:2px;}
.chip{background:#f1f5f9;border:1px solid var(--line);color:#475569;
      padding:5px 11px;border-radius:20px;font-size:12px;cursor:pointer;}
.chip:hover{background:#e0ecff;color:#1d4ed8;border-color:#bfdbfe;}
.opts{display:flex;gap:24px;flex-wrap:wrap;align-items:flex-end;}
.opt{display:flex;flex-direction:column;gap:6px;}
.opt small{color:#94a3b8;font-size:11px;}
.opt input[type=text]{width:150px;min-width:0;padding:8px 10px;}
.opt input[type=number]{width:110px;padding:8px 10px;border:1px solid var(--line);
   border-radius:8px;outline:none;background:#fff;color:var(--text);}
.check{display:flex;align-items:center;gap:8px;padding-bottom:10px;cursor:pointer;
       color:var(--text);font-size:13px;}
.check input{width:16px;height:16px;accent-color:var(--primary);}
.actions{margin-top:18px;display:flex;gap:10px;align-items:center;flex-wrap:wrap;}
.progress-wrap{margin-top:16px;display:none;}
.bar{height:8px;background:#e2e8f0;border-radius:20px;overflow:hidden;}
.bar > i{display:block;height:100%;width:8%;border-radius:20px;
   background:linear-gradient(90deg,#60a5fa,#2563eb);
   animation:move 1.2s ease-in-out infinite alternate;}
@keyframes move{from{width:6%;}to{width:92%;}}
.done .bar > i{animation:none;width:100%;background:var(--ok);}
.status{margin-top:8px;font-size:13px;color:var(--muted);}
.stats{display:none;grid-template-columns:repeat(4,1fr);gap:12px;margin-top:18px;}
.stat{background:var(--soft);border:1px solid var(--line);border-radius:10px;padding:14px;}
.stat .k{font-size:12px;color:var(--muted);}
.stat .v{font-size:21px;font-weight:700;margin-top:4px;color:#0f172a;}
.preview{display:none;margin-top:16px;max-height:320px;overflow:auto;
   border:1px solid var(--line);border-radius:8px;}
table{width:100%;border-collapse:collapse;font-size:12px;}
th,td{padding:8px 10px;text-align:left;border-bottom:1px solid #eef2f7;white-space:nowrap;}
th{background:var(--soft);color:#475569;position:sticky;top:0;}
td.path{color:#475569;max-width:560px;overflow:hidden;text-overflow:ellipsis;}
.tag{padding:2px 8px;border-radius:12px;font-size:11px;}
.tag.dir{background:#e0f2fe;color:#0369a1;}
.tag.file{background:#dcfce7;color:#15803d;}
.note{font-size:12px;color:var(--muted);margin-top:10px;}
.export-grid{display:grid;grid-template-columns:180px 1fr auto;gap:12px;align-items:end;}
select{padding:10px 12px;border:1px solid var(--line);border-radius:8px;
   background:#fff;color:var(--text);font-size:14px;outline:none;font-family:inherit;}
.msg{margin-top:12px;padding:10px 14px;border-radius:8px;display:none;font-size:13px;}
.msg.ok{background:#dcfce7;color:#166534;display:block;}
.msg.err{background:#fee2e2;color:#991b1b;display:block;}
@media(max-width:760px){.stats{grid-template-columns:repeat(2,1fr);}
  .export-grid{grid-template-columns:1fr;}}
</style>
</head>
<body>
<div class="header"><div class="inner">
  <h1>目录扫描导出工具</h1>
  <p>扫描指定硬盘或文件夹的全部文件，一键导出 TXT / CSV / HTML 目录清单</p>
</div></div>

<div class="main">
  <div class="panel">
    <h2>1. 选择扫描位置</h2>
    <label>硬盘盘符或文件夹路径</label>
    <div class="input-row">
      <input type="text" id="path" placeholder="例如：C:\ 或 D:\项目资料 或 C:\Users\...\Desktop">
      <button class="btn-primary" id="btnScan">开始扫描</button>
      <button class="btn-ghost" id="btnCancel" disabled>取消</button>
    </div>
    <div class="quick" id="quick"><span>快速选择：</span></div>
  </div>

  <div class="panel">
    <h2>2. 扫描选项</h2>
    <div class="opts">
      <div class="opt">
        <label>最大深度（0 = 不限）</label>
        <input type="number" id="maxDepth" value="0" min="0" max="99">
      </div>
      <div class="opt">
        <label>只含扩展名（逗号分隔，留空=全部）</label>
        <input type="text" id="incExt" placeholder="例如：txt,docx,pdf">
      </div>
      <div class="opt">
        <label>排除扩展名（逗号分隔）</label>
        <input type="text" id="excExt" placeholder="例如：tmp,bak">
      </div>
      <label class="check"><input type="checkbox" id="incHidden"> 包含隐藏文件/文件夹</label>
      <label class="check"><input type="checkbox" id="incDirs" checked> 导出中包含文件夹</label>
    </div>

    <div class="progress-wrap" id="progressWrap">
      <div class="bar"><i></i></div>
      <div class="status" id="status">准备中...</div>
    </div>

    <div class="stats" id="stats">
      <div class="stat"><div class="k">文件数量</div><div class="v" id="sFiles">0</div></div>
      <div class="stat"><div class="k">文件夹数量</div><div class="v" id="sDirs">0</div></div>
      <div class="stat"><div class="k">文件总大小</div><div class="v" id="sSize">0 B</div></div>
      <div class="stat"><div class="k">耗时 / 错误</div><div class="v" id="sTime">0s / 0</div></div>
    </div>

    <div class="preview" id="preview">
      <table><thead><tr><th>类型</th><th>相对路径</th><th>大小</th><th>修改时间</th></tr></thead>
      <tbody id="previewBody"></tbody></table>
    </div>
    <div class="note" id="previewNote"></div>
  </div>

  <div class="panel">
    <h2>3. 导出目录文件</h2>
    <div class="export-grid">
      <div class="opt">
        <label>导出格式</label>
        <select id="fmt">
          <option value="html">HTML 网页（可浏览，推荐）</option>
          <option value="csv">CSV 表格（Excel 打开）</option>
          <option value="txt">TXT 文本</option>
        </select>
      </div>
      <div class="opt" style="flex:1;">
        <label>保存路径（留空则保存到桌面）</label>
        <input type="text" id="output" placeholder="留空 = 桌面\目录清单_日期.格式">
      </div>
      <button class="btn-primary" id="btnExport" disabled>导出文件</button>
    </div>
    <div class="msg" id="msg"></div>
    <div class="actions">
      <button class="btn-light" id="btnOpen" style="display:none;">打开导出文件所在文件夹</button>
      <span class="note" id="exportInfo"></span>
    </div>
  </div>

  <div style="text-align:center;margin-top:8px;">
    <button class="btn-ghost" id="btnQuit" style="padding:7px 16px;font-size:12px;">退出程序</button>
  </div>
</div>

<script>
const $ = id => document.getElementById(id);
let taskId = null, pollTimer = null, lastResult = null, lastOutput = null;

function fmtSize(n){
  if(!n) return '0 B';
  const u=['B','KB','MB','GB','TB','PB']; let i=0, x=Number(n);
  while(x>=1024 && i<u.length-1){x/=1024;i++;}
  return i===0 ? Math.round(x)+' B' : x.toFixed(2)+' '+u[i];
}

function showMsg(text, ok){
  const m=$('msg'); m.textContent=text; m.className='msg '+(ok?'ok':'err');
}

async function api(path, data){
  const opt = data ? {method:'POST',headers:{'Content-Type':'application/json'},
                      body:JSON.stringify(data)} : {};
  const r = await fetch(path, opt);
  return await r.json();
}

async function loadDrives(){
  const d = await api('/api/drives');
  const box = $('quick');
  (d.drives||[]).forEach(dr=>{
    const b=document.createElement('button');
    b.className='chip';
    b.textContent=dr.path+'  ('+fmtSize(dr.free)+' 可用)';
    b.onclick=()=>{ $('path').value=dr.path; };
    box.appendChild(b);
  });
  [['桌面','Desktop'],['文档','Documents'],['下载','Downloads']].forEach(([name,key])=>{
    const b=document.createElement('button');
    b.className='chip'; b.textContent=name;
    b.onclick=async()=>{
      const r=await api('/api/quickpath?name='+key);
      if(r.path) $('path').value=r.path;
    };
    box.appendChild(b);
  });
}

$('btnScan').onclick = async ()=>{
  const p = $('path').value.trim();
  if(!p){ alert('请先填写要扫描的硬盘或文件夹路径'); return; }
  $('btnScan').disabled=true; $('btnCancel').disabled=false;
  $('btnExport').disabled=true; $('progressWrap').style.display='block';
  $('progressWrap').classList.remove('done');
  $('stats').style.display='none'; $('preview').style.display='none';
  $('previewNote').textContent=''; $('msg').style.display='none';
  $('status').textContent='正在启动扫描...';

  const r = await api('/api/scan', {
    path:p,
    max_depth: $('maxDepth').value,
    include_ext: $('incExt').value,
    exclude_ext: $('excExt').value,
    include_hidden: $('incHidden').checked,
    include_dirs: $('incDirs').checked
  });
  if(r.error){ showMsg(r.error,false); $('btnScan').disabled=false;
               $('btnCancel').disabled=true; return; }
  taskId = r.id;
  pollTimer = setInterval(poll, 350);
};

$('btnCancel').onclick = async ()=>{
  if(!taskId) return;
  await api('/api/cancel', {id:taskId});
  $('status').textContent='正在取消...';
};

async function poll(){
  if(!taskId) return;
  const s = await api('/api/progress?id='+encodeURIComponent(taskId));
  if(s.error){ clearInterval(pollTimer); showMsg(s.error,false); return; }
  $('status').textContent = s.status_text + '  (用时 '+s.elapsed+'s)';
  if(s.status==='done' || s.status==='cancelled'){
    clearInterval(pollTimer);
    $('progressWrap').classList.add('done');
    $('btnScan').disabled=false; $('btnCancel').disabled=true;
    if(s.status==='done') $('btnExport').disabled=false;
    lastResult=s;
    $('stats').style.display='grid';
    $('sFiles').textContent=s.files;
    $('sDirs').textContent=s.dirs;
    $('sSize').textContent=fmtSize(s.total_size);
    $('sTime').textContent=s.elapsed+'s / '+s.errors;
    const tb=$('previewBody'); tb.innerHTML='';
    (s.preview||[]).forEach(r=>{
      const tr=document.createElement('tr');
      tr.innerHTML='<td><span class="tag '+(r.type==='dir'?'dir':'file')+'">'+
        (r.type==='dir'?'文件夹':'文件')+'</span></td><td class="path" title="'+
        r.path.replace(/"/g,'&quot;')+'">'+r.rel+'</td><td>'+
        (r.type==='dir'?'':fmtSize(r.size))+'</td><td>'+r.mtime+'</td>';
      tb.appendChild(tr);
    });
    $('preview').style.display = (s.preview&&s.preview.length)?'block':'none';
    const total = s.files + s.dirs;
    $('previewNote').textContent = (total > (s.preview||[]).length)
      ? '预览仅显示前 '+((s.preview||[]).length)+' 条，导出文件包含全部 '+total+' 条记录。' : '';
    if(s.status==='cancelled') showMsg('扫描已取消',false);
  }
}

$('btnExport').onclick = async ()=>{
  if(!taskId) return;
  $('btnExport').disabled=true;
  showMsg('正在生成导出文件...', true);
  const r = await api('/api/export', {
    id:taskId,
    format:$('fmt').value,
    output:$('output').value.trim(),
    include_dirs:$('incDirs').checked
  });
  $('btnExport').disabled=false;
  if(r.error){ showMsg(r.error,false); return; }
  lastOutput = r.path;
  showMsg('导出成功！共 '+r.rows+' 条记录，文件大小 '+fmtSize(r.size), true);
  $('exportInfo').textContent = r.path;
  $('btnOpen').style.display='inline-block';
};

$('btnOpen').onclick = ()=>{ if(lastOutput) api('/api/reveal?p='+encodeURIComponent(lastOutput)); };

$('btnQuit').onclick = async ()=>{
  if(!confirm('确定要退出目录扫描工具吗？')) return;
  await api('/api/quit', {});
  document.body.innerHTML = '<div style="text-align:center;padding:80px 20px;'
    + 'font-family:Microsoft YaHei,sans-serif;color:#64748b;">'
    + '<h2 style="color:#0f172a;">程序已退出</h2><p>可以关闭此浏览器标签页。</p></div>';
};

loadDrives();
</script>
</body>
</html>"""


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        pass  # 保持控制台干净

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
        elif u.path == "/api/progress":
            tid = parse_qs(u.query).get("id", [""])[0]
            task = TASKS.get(tid)
            self._send(200, task.snapshot() if task else {"error": "任务不存在，请重新扫描"})
        elif u.path == "/api/reveal":
            p = parse_qs(u.query).get("p", [""])[0]
            try:
                if os.name == "nt":
                    import subprocess
                    folder = os.path.dirname(p) if os.path.isfile(p) else p
                    subprocess.Popen(["explorer", folder])
                elif sys.platform == "darwin":
                    import subprocess
                    subprocess.Popen(["open", "-R", p])
                else:
                    import subprocess
                    subprocess.Popen(["xdg-open", os.path.dirname(p) or p])
                self._send(200, {"ok": True})
            except Exception as e:
                self._send(200, {"error": str(e)})
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
                self._send(200, {"error": "未找到扫描结果，请先执行扫描"})
                return
            fmt = (data.get("format") or "csv").lower()
            out = (data.get("output") or "").strip().strip('"')
            if not out:
                desktop = os.path.join(os.path.expanduser("~"), "Desktop")
                if not os.path.isdir(desktop):
                    desktop = os.path.expanduser("~")
                out = os.path.join(
                    desktop,
                    "目录清单_" + datetime.now().strftime("%Y%m%d_%H%M%S") + "." + fmt,
                )
            out = os.path.abspath(os.path.expanduser(out))
            try:
                parent = os.path.dirname(out)
                if parent:
                    os.makedirs(parent, exist_ok=True)
                rows = do_export(task, fmt, out, bool(data.get("include_dirs", True)))
                self._send(200, {"ok": True, "path": out, "size": os.path.getsize(out),
                                 "rows": rows})
            except Exception as e:
                self._send(200, {"error": "导出失败：" + str(e)})

        elif u.path == "/api/quit":
            self._send(200, {"ok": True})
            # 延迟关闭，确保上面的响应先发送完成
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
    print("目录扫描导出工具已启动：", url)
    if "--no-browser" not in sys.argv:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n已退出")


if __name__ == "__main__":
    main()
