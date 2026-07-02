"""
02324 · 从 WPS 云文档下载最新 .docx

用法:
  python download_wps.py

工作模式:
  1. 优先尝试已保存的 cookies（wps_cookies.json），用 Playwright 独立浏览器静默下载
  2. 如果 cookies 无效或不存在，提示用户关闭 Edge，用 Edge 配置登录下载

首次使用:
  - 需要关闭 Edge 一次 → 脚本自动拉取 Edge 登录态 → 下载并保存 cookies
  - 后续无需再关 Edge
"""

import asyncio
import json
import os
import sys
import time

from playwright.async_api import async_playwright

ROOT = os.path.dirname(os.path.abspath(__file__))
COOKIES_FILE = os.path.join(ROOT, "wps_cookies.json")
TEMP_PROFILE = os.path.join(ROOT, ".wps-temp-profile")
WPS_URL = "https://www.kdocs.cn/l/cp1CPS0QvBWV"
EDGE_USER_DATA = os.path.expandvars(
    r"%LOCALAPPDATA%\Microsoft\Edge\User Data"
)


def log(*args, **kwargs):
    print("[WPS]", *args, **kwargs)


# ── 模式 A：用已保存 cookies ──


async def download_with_cookies():
    """用已保存的 cookies + 独立浏览器下载。返回 (文件路径, 大小) 或 None。"""
    if not os.path.exists(COOKIES_FILE):
        return None

    cookies = json.load(open(COOKIES_FILE, "r", encoding="utf-8"))
    log(f"📂 已找到保存的 cookies（{len(cookies)} 条）")

    async with async_playwright() as p:
        context = await p.chromium.launch_persistent_context(
            TEMP_PROFILE,
            headless=False,  # 有窗口但不需要用户操作，方便看到进度
            args=["--window-size=1024,768"],
        )
        page = context.pages[0] if context.pages else await context.new_page()

        # 注入 cookies（必须在 navigate 前）
        await context.add_cookies(cookies)

        await page.goto(WPS_URL, wait_until="networkidle", timeout=30000)
        await asyncio.sleep(4)

        # 检查是否登录成功
        has_login = await page.evaluate(
            '() => document.body.innerText.includes("立即登录")'
        )
        if has_login:
            log("⚠️  Cookies 已失效，需要重新登录")
            await context.close()
            return None

        log("✅ Cookies 有效，已登录")
        result = await _trigger_download(page)
        await context.close()
        return result


# ── 模式 B：用 Edge 配置（首次/fallback）──


async def download_with_edge():
    """用 Edge 用户配置下载。需要关闭 Edge。返回 (文件路径, 大小) 或 None。"""
    log("🔄 需要借助 Edge 浏览器...")

    # 检查 Edge 是否在运行
    import subprocess

    check = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq msedge.exe"],
        capture_output=True, text=True, timeout=5,
    )
    if "msedge.exe" in check.stdout:
        log("⛔ Edge 正在运行，请关闭后按 Enter 继续...")
        input("   按 Enter 继续（请先关闭所有 Edge 窗口）...")
        # 再次确认
        time.sleep(2)
        check2 = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq msedge.exe"],
            capture_output=True, text=True, timeout=5,
        )
        if "msedge.exe" in check2.stdout:
            log("❌ Edge 仍在运行，请在任务管理器中结束 msedge.exe 后重试")
            return None

    async with async_playwright() as p:
        context = await p.chromium.launch_persistent_context(
            EDGE_USER_DATA,
            channel="msedge",
            headless=False,
            args=["--profile-directory=Default", "--window-size=1024,768"],
        )
        page = context.pages[0] if context.pages else await context.new_page()

        await page.goto(WPS_URL, wait_until="networkidle", timeout=30000)
        await asyncio.sleep(5)

        # 检查是否登录
        has_login = await page.evaluate(
            '() => document.body.innerText.includes("立即登录")'
        )
        if has_login:
            log("❌ Edge 中 WPS 未登录，请在浏览器中登录后重试")
            await context.close()
            return None

        log("✅ Edge 已登录 WPS")
        result = await _trigger_download(page)

        # 保存 cookies 供后续使用
        if result:
            cookies = await context.cookies()
            with open(COOKIES_FILE, "w", encoding="utf-8") as f:
                json.dump(cookies, f, ensure_ascii=False, indent=2)
            log(f"💾 Cookies 已保存到 {COOKIES_FILE}（后续无需再关 Edge）")

        await context.close()
        return result


# ── 公共：触以下载 ──


async def _trigger_download(page):
    """在页面上点击下载按钮，返回 (文件路径, 大小) 或 None。"""
    try:
        # 点击右上角更多按钮（三个点）
        more_btn = await page.query_selector(".app-header-more-btn")
        if more_btn:
            await more_btn.click()
            await asyncio.sleep(2)
        else:
            log("⚠️  未找到更多按钮，尝试直接寻找下载按钮")

        # 点"下载"
        download_btn = await page.query_selector("text=下载")
        if not download_btn:
            log("❌ 找不到下载按钮")
            return None

        async with page.expect_download(timeout=20000) as download_info:
            await download_btn.click()

        download = await download_info.value
        filename = download.suggested_filename
        dest = os.path.join(ROOT, filename)

        # 如果同名文件已存在，先删除（playwright save_as 不会覆盖）
        if os.path.exists(dest):
            os.remove(dest)

        await download.save_as(dest)
        size = os.path.getsize(dest)
        log(f"📥 下载成功: {filename} ({size/1024:.1f} KB)")
        return (dest, size)

    except Exception as e:
        log(f"❌ 下载失败: {e}")
        return None


# ── 入口 ──


async def main():
    log("🚀 开始从 WPS 下载最新笔记...\n")

    # 先尝试 cookies 模式
    result = await download_with_cookies()
    if result:
        return 0

    # 降级到 Edge 模式
    print()  # 空行
    result = await download_with_edge()
    if result:
        return 0

    log("❌ 下载失败，请检查后重试")
    return 1


if __name__ == "__main__":
    exit_code = asyncio.run(main())
    sys.exit(exit_code)
