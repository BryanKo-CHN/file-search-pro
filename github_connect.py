#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
github_connect.py — GitHub 连接控制流程（Watt Toolkit 加速场景）

背景（2026-11-23 本机实测）：
  * Watt Toolkit（Steam++）加速方式 = hosts 劫持 + 本地 443 接管：
    它把 github.com / raw.githubusercontent.com 等域名写进 hosts 指向 127.0.0.1，
    由 Steam++.Accelerator.exe 监听 0.0.0.0:443 转发流量。
  * Windows schannel 的证书吊销检查会报 CRYPT_E_NO_REVOCATION_CHECK (0x80092012)，
    这是"假失败"：curl 加 --ssl-no-revoke 立刻 200；git 配 sslVerify=false 即通。
  * WorkBuddy 内部代理（环境变量 http_proxy=127.0.0.1:29992）不稳定，
    偶发 502 / TLS 握手失败；绕过它直连（走 Watt 本地 443）最稳。

用法：
  python github_connect.py check            # 环境诊断 + 多路径连通性测试
  python github_connect.py push "提交说明"  # 自动选路 → add/commit/push → 验证
  python github_connect.py fetch            # 自动选路 → fetch
  python github_connect.py verify           # 只对照本地/远程 HEAD

设计原则：
  1. 先诊断再动作；先读后写；能验证就验证。
  2. 绝不打印 token / 环境变量敏感值。
  3. 任何一步失败都给出明确下一步指引，不静默吞错。
"""

import os
import re
import sys
import subprocess

REPO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "github-publish")
HOSTS = r"C:\Windows\System32\drivers\etc\hosts"
TIMEOUT = 12
OK_CODES = {"200", "201", "301", "302", "304"}

# 需要绕过的代理变量（大小写都列，Windows 环境常混用）
PROXY_VARS = ("http_proxy", "https_proxy", "HTTP_PROXY", "HTTPS_PROXY",
              "ALL_PROXY", "all_proxy")
# 只读探测 URL（仓库公开时 raw 可读；私有时用 github.com 首页即可）
PROBE_URLS = ("https://github.com", "https://raw.githubusercontent.com")


def _run(cmd, env=None, timeout=TIMEOUT, capture=True, cwd=None):
    """统一执行命令，返回 (returncode, stdout, stderr)。"""
    try:
        p = subprocess.run(cmd, env=env, timeout=timeout, cwd=cwd,
                           stdout=subprocess.PIPE if capture else None,
                           stderr=subprocess.PIPE if capture else None)
        out = p.stdout.decode("utf-8", "replace") if p.stdout else ""
        err = p.stderr.decode("utf-8", "replace") if p.stderr else ""
        return p.returncode, out, err
    except subprocess.TimeoutExpired:
        return -1, "", "TIMEOUT"
    except FileNotFoundError:
        return -2, "", "NOT_FOUND"


def _git(args, env=None, timeout=120):
    # 所有 git 命令都在仓库目录执行，保证 status/config/remote 读到正确仓库
    return _run(["git"] + list(args), env=env, timeout=timeout, cwd=REPO)


def base_env():
    return os.environ.copy()


def strip_proxy(env):
    """生成"绕过环境代理"的环境变量副本（走 Watt 本地 443 / 真实直连）。"""
    e = env.copy()
    for k in PROXY_VARS:
        e.pop(k, None)
    return e


def proxy_desc(env):
    """只显示代理主机，不显示任何 token。"""
    v = env.get("https_proxy") or env.get("HTTPS_PROXY") or ""
    m = re.match(r"https?://([^:/]+):?(\d*)", v)
    return m.group(0).replace(v, f"{m.group(1)}:{m.group(2)}") if m else "(无)"


# ---------------------------------------------------------------------------
# 1. 环境诊断
# ---------------------------------------------------------------------------
def diagnose():
    print("=" * 68)
    print("【1/4】环境诊断")
    print("=" * 68)

    # Watt Toolkit 进程
    rc, out, _ = _run(["tasklist"], timeout=20)
    watt = [ln for ln in out.splitlines()
            if re.search(r"Steam\+\+", ln, re.I)]
    print(f"  Watt Toolkit 进程 : {'运行中 (' + str(len(watt)) + ' 个)' if watt else '未检测到'}")
    for ln in watt[:3]:
        print("      " + ln.strip()[:90])

    # 本地 443 监听
    rc, out, _ = _run(["netstat", "-ano"], timeout=20)
    listening = bool(re.search(r":443\s+\S+\s+LISTENING", out, re.I))
    print(f"  本地 443 监听     : {'有（Watt 加速接管）' if listening else '无'}")

    # hosts GitHub 劫持
    hijack = 0
    try:
        with open(HOSTS, "r", encoding="utf-8", errors="ignore") as fh:
            for line in fh:
                if re.search(r"\bgithub(\.com|\.io|usercontent\.com)", line, re.I):
                    hijack += 1
    except Exception as e:
        print(f"  hosts 读取失败    : {e}")
    print(f"  hosts GitHub 条目 : {hijack} 条"
          + ("（Watt 正在劫持域名到本地）" if hijack else ""))

    # git SSL 配置
    rc, out, _ = _git(["config", "--get", "http.sslVerify"], env=base_env(), timeout=20)
    sslv = out.strip() or "(未设置)"
    print(f"  git sslVerify     : {sslv}"
          + ("  ← 正确" if sslv.lower() == "false" else "  ← 建议设为 false"))
    if sslv.lower() != "false":
        print("      修复: cd github-publish && git config http.sslVerify false")

    # 仓库状态
    if os.path.isdir(os.path.join(REPO, ".git")):
        rc, out, _ = _git(["status", "-sb"], env=base_env(), timeout=20)
        print(f"  git 仓库          : {out.strip().splitlines()[0] if out.strip() else '?'} (分支状态)")
    else:
        print("  git 仓库          : 未找到 github-publish/.git")
    print(f"  环境代理          : {proxy_desc(base_env())}")
    return {"watt": bool(watt), "listen443": listening, "hosts": hijack,
            "ssl_ok": sslv.lower() == "false"}


# ---------------------------------------------------------------------------
# 2. 多路径连通性探测
# ---------------------------------------------------------------------------
def probe(route, env):
    """探测单条路径是否可用，返回 (ok, detail)。"""
    e = strip_proxy(env) if route == "direct" else env.copy()
    curl = ["curl", "--ssl-no-revoke", "-sS", "-o", os.devnull,
            "-w", "%{http_code}", "--max-time", str(TIMEOUT)]
    if route == "direct":
        curl += ["--noproxy", "*"]
    curl.append(PROBE_URLS[0])
    rc, out, err = _run(curl, env=e, timeout=TIMEOUT + 4)
    code = out.strip().splitlines()[-1] if out.strip() else ""
    if code in OK_CODES:
        return True, f"HTTP {code}"
    return False, (err.strip().splitlines()[0][:70] if err.strip()
                   else (f"HTTP {code}" if code else "无响应"))


def pick_route():
    print()
    print("=" * 68)
    print("【2/4】连通性探测（curl --ssl-no-revoke 规避 schannel 假失败）")
    print("=" * 68)
    env = base_env()
    routes = [("direct", "绕过环境代理 → 直连/Watt 本地 443  [首选]"),
              ("envproxy", f"使用环境代理 {proxy_desc(env)}  [备用]")]
    for name, label in routes:
        ok, detail = probe(name, env)
        print(f"  [{'通' if ok else '不通'}] {label:<52} {detail}")
        if ok:
            print(f"  → 选定路线: {name}")
            return name
    print("  ✗ 两条路径都不通。请检查：")
    print("     1) Watt Toolkit 是否已开启「GitHub 加速」开关")
    print("     2) 任务管理器是否有 Steam++.Accelerator.exe 进程")
    print("     3) hosts 里 github.com 是否指向 127.0.0.1")
    return None


# ---------------------------------------------------------------------------
# 3. git 操作（用选定路线）
# ---------------------------------------------------------------------------
def git_env(route):
    e = base_env()
    return strip_proxy(e) if route == "direct" else e


def do_fetch(route, remote="origin"):
    print()
    print("=" * 68)
    print(f"【3/4】git fetch（路线={route}）")
    print("=" * 68)
    rc, out, err = _git(["fetch", remote], env=git_env(route), timeout=180)
    print("  " + (out.strip() or err.strip() or "(无输出)").replace("\n", "\n  ")[:600])
    if rc != 0:
        print(f"  ✗ fetch 失败 (rc={rc})")
        return False
    print("  ✓ fetch 成功")
    return True


def do_push(route, message):
    env = git_env(route)
    # 有改动才 commit
    rc, out, _ = _git(["status", "--porcelain"], env=env, timeout=30)
    if out.strip():
        print()
        print("=" * 68)
        print("【3/4】提交本地改动")
        print("=" * 68)
        rc, out, err = _git(["add", "-A"], env=env, timeout=60)
        if rc != 0:
            print("  ✗ git add 失败:", err.strip()[:200]); return False
        rc, out, err = _git(["commit", "-m", message], env=env, timeout=60)
        print("  " + (out.strip() or err.strip()).replace("\n", "\n  ")[:400])
        if rc != 0 and "nothing to commit" not in out + err:
            print("  ✗ commit 失败"); return False
    else:
        print("  (无本地改动，跳过 commit)")

    # push：最多 3 次尝试
    print()
    print("=" * 68)
    print(f"【4/4】git push（路线={route}，最多重试 3 次）")
    print("=" * 68)
    for attempt in range(1, 4):
        print(f"  --- 第 {attempt} 次 ---")
        rc, out, err = _git(["push", "origin", "main"], env=env, timeout=180)
        tail = (out.strip() or err.strip()).splitlines()
        print("  " + "\n  ".join(tail[-6:])[:500])
        if rc == 0:
            print("  ✓ push 成功")
            return True
        print("  重试中…")
    print("  ✗ push 三次失败。建议：确认 Watt 加速开关 → 重试；")
    print("     或手动执行: cd github-publish && git push origin main")
    return False


def verify(route):
    print()
    print("=" * 68)
    print("【验证】本地 HEAD vs 远程 HEAD")
    print("=" * 68)
    env = git_env(route)
    _, lout, _ = _git(["rev-parse", "HEAD"], env=env, timeout=30)
    _, rout, rerr = _git(["ls-remote", "origin", "HEAD"], env=env, timeout=60)
    local = lout.strip()
    remote = rout.strip().split()[0] if rout.strip() else ""
    print(f"  本地 HEAD: {local[:12]}")
    print(f"  远程 HEAD: {remote[:12] if remote else '(读取失败) ' + rerr.strip()[:60]}")
    if local and remote:
        ok = local == remote
        print("  " + ("✓ 完全一致，发布成功" if ok else "✗ 不一致，可能需要再 push"))
        return ok
    return False


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return
    cmd = sys.argv[1].lower()

    if cmd == "check":
        diagnose()
        pick_route()
        return

    if cmd not in ("push", "fetch", "verify"):
        print("未知命令:", cmd); print(__doc__); return

    if not os.path.isdir(os.path.join(REPO, ".git")):
        print("✗ 未找到 git 仓库:", REPO); return

    info = diagnose()
    if not info["ssl_ok"]:
        print("\n  自动修复 git sslVerify=false（仓库级，可随时改回）…")
        _git(["config", "http.sslVerify", "false"], env=base_env(), timeout=20)

    route = pick_route()
    if not route:
        return

    if cmd == "fetch":
        do_fetch(route)
    elif cmd == "push":
        message = sys.argv[2] if len(sys.argv) > 2 else "Update via github_connect.py"
        if do_push(route, message):
            verify(route)
    elif cmd == "verify":
        verify(route)


if __name__ == "__main__":
    main()
