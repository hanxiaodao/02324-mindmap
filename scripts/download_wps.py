"""
02324 · 从 WPS 云文档下载最新 .docx

用法:
  python download_wps.py

工作模式:
  1. 优先尝试已保存的 cookies（wps_cookies.json），静默下载
  2. 如果 cookies 无效或不存在，启动独立 Chrome 让用户手动登录一次

首次使用:
  - 脚本会启动一个独立 Chrome 窗口（不影响你已打开的 Chrome）
  - 在窗口中登录 WPS → 自动检测到登录后开始下载
  - 保存 cookies，后续运行无需再登录
  - Chrome 正常用，不用关
"""

import asyncio
import json
import os
import sys
import time

from playwright.async_api import async_playwright

ROOT = os.path.dirname(os.path.abspath(__file__))       # scripts/
PROJECT_ROOT = os.path.dirname(ROOT)                     # 项目根目录
COOKIES_FILE = os.path.join(ROOT, "wps_cookies.json")
TEMP_PROFILE = os.path.join(ROOT, ".wps-temp-profile")
TEMP_PROFILE_CHROME = os.path.join(ROOT, ".wps-temp-chrome")
WPS_URL = "https://www.kdocs.cn/l/cp1CPS0QvBWV"


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


# ── 模式 B：首次/fallback（需要时才让用户登录）──

BROWSER_TITLE = "Chrome"


async def _needs_login(page) -> bool:
    """检测 WPS 页面是否需要登录。"""
    text = await page.evaluate('() => document.body.innerText')
    if "立即登录" in text:
        return True
    # 检查登录弹窗遮罩
    popup = await page.query_selector("#util-popup")
    if popup:
        popup_text = await popup.inner_text()
        if "登录" in popup_text or "扫码" in popup_text:
            return True
    return False


async def download_with_chrome():
    """用 Playwright 独立 Chrome 实例下载。
    
    临时配置目录会保留登录状态，只需登录一次。
    Chrome 正常用，不用关。
    """
    log(f"🔄 正在启动 {BROWSER_TITLE} 浏览器（独立配置，不影响已打开的窗口）...")

    import subprocess
    r = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq chrome.exe"],
        capture_output=True, text=True, timeout=5,
    )
    if "chrome.exe" in r.stdout:
        log(f"💡 {BROWSER_TITLE} 正在运行中，脚本会另开一个独立实例，无需关闭")

    async with async_playwright() as p:
        context = await p.chromium.launch_persistent_context(
            TEMP_PROFILE_CHROME,
            channel="chrome",
            headless=False,
            args=["--window-size=1024,768"],
        )
        page = context.pages[0] if context.pages else await context.new_page()

        await page.goto(WPS_URL, wait_until="networkidle", timeout=30000)
        await asyncio.sleep(3)

        # 只有页面确实需要登录时才提示
        if await _needs_login(page):
            print()
            log("🔑 请在浏览器窗口中登录 WPS（扫码或账号均可）")
            log("   登录后按 Enter 继续...")
            print()
            input("   按 Enter 继续...")
            # 等页面登录完成重定向稳定，不重新 goto（避免导航冲突）
            await page.wait_for_load_state("networkidle", timeout=30000)
            await asyncio.sleep(3)

        log(f"✅ {BROWSER_TITLE} 已登录 WPS")
        result = await _trigger_download(page)

        if result:
            cookies = await context.cookies()
            with open(COOKIES_FILE, "w", encoding="utf-8") as f:
                json.dump(cookies, f, ensure_ascii=False, indent=2)
            log("💾 Cookies 已保存（后续无须再登录）")

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
        dest = os.path.join(PROJECT_ROOT, filename)

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

    # 降级到 Chrome 模式
    print()  # 空行
    result = await download_with_chrome()
    if result:
        return 0

    log("❌ 下载失败，请检查后重试")
    return 1


if __name__ == "__main__":
    exit_code = asyncio.run(main())
    sys.exit(exit_code)
