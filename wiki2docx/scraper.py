"""Playwright-based scraper for RoboMaster Wiki pages."""

import json
import os
import sys
import time

from playwright.sync_api import sync_playwright

# Fix Windows GBK console encoding
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


class WikiScraper:
    """Scrape a RoboMaster Wiki page by clicking through every sidebar node."""

    WIKI_URL_TEMPLATE = "https://bbs.robomaster.com/wiki/{wiki_id}"

    DJI_LOGIN_URL = (
        "https://account.dji.com/login"
        "?appId=rm-bbs-backend"
        "&autologin=y"
        "&backUrl=https://bbs.robomaster.com/"
        "&locale=zh_CN&mode=redirect"
    )

    def __init__(self, wiki_id: str, profile_dir: str):
        self.wiki_id = wiki_id
        self.wiki_url = self.WIKI_URL_TEMPLATE.format(wiki_id=wiki_id)
        self.profile_dir = os.path.abspath(profile_dir)

    # ------------------------------------------------------------------
    # Public
    # ------------------------------------------------------------------

    def scrape(self) -> dict:
        """Entry: open browser, QR login, expand tree, click-scrape all sections."""
        os.makedirs(self.profile_dir, exist_ok=True)

        with sync_playwright() as p:
            browser = p.chromium.launch(
                channel="chrome",
                headless=False,
                args=["--disable-blink-features=AutomationControlled"],
            )
            context = browser.new_context(viewport={"width": 1400, "height": 900})
            page = context.new_page()

            self._qr_login_and_wait(page)
            self._navigate_to_wiki(page)
            all_sections = self._extract_menu_tree(page)
            self._scrape_sections(page, all_sections)

            context.close()
            browser.close()

        return {
            "wiki_url": self.wiki_url,
            "title": all_sections[0]["title"] if all_sections else "",
            "scraped_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
            "total_nodes": len(all_sections),
            "leaf_nodes": sum(1 for s in all_sections if s["isLeaf"]),
            "with_content": sum(1 for s in all_sections if s["content"]),
            "empty": sum(1 for s in all_sections if not s["content"]),
            "sections": all_sections,
        }

    # ------------------------------------------------------------------
    # Login
    # ------------------------------------------------------------------

    def _qr_login_and_wait(self, page):
        """Open DJI login page, screenshot WeChat QR, poll until user scans."""

        # ---- Step 1: Open DJI login ----
        page.goto(self.DJI_LOGIN_URL, timeout=30000)
        time.sleep(5)

        # ---- Step 2: Switch to WeChat tab ----
        try:
            page.locator("text=微信登录").first.click(timeout=5000)
            time.sleep(5)
        except Exception:
            print("⚠ 未找到「微信登录」选项，继续截图...")

        # ---- Step 3: Screenshot QR code ----
        qr_path = os.path.join(self.profile_dir, "wechat_login_qr.png")
        page.screenshot(path=qr_path)

        print()
        print("=" * 55)
        print(f" 微信登录二维码已保存至：")
        print(f"   {qr_path}")
        print()
        print(" 请打开该图片，用微信扫描二维码完成登录。")
        print(" 脚本将自动检测登录状态，无需操作。")
        print("=" * 55)
        print()

        # ---- Step 4: Poll for login success ----
        # After WeChat scan, DJI redirects to https://bbs.robomaster.com/
        logged_in = False
        for i in range(600):  # 10 minutes
            time.sleep(1)
            cur_url = page.url
            # User explicitly stated: after login, redirect to https://bbs.robomaster.com/
            if cur_url.rstrip("/") == "https://bbs.robomaster.com":
                body = page.content()
                if "哎呀" not in body:
                    logged_in = True
                    print("检测到已重定向至 bbs.robomaster.com，登录成功！")
                    break
            if i > 0 and i % 30 == 0:
                print(f"  已等待 {i} 秒，请扫码...")

        if not logged_in:
            raise RuntimeError("登录超时（10 分钟），请重新运行。")

    # ------------------------------------------------------------------
    # Wiki navigation
    # ------------------------------------------------------------------

    def _navigate_to_wiki(self, page):
        """After login, navigate to the wiki page and verify access."""
        print(f"导航到 wiki 页面: {self.wiki_url}")
        page.goto(self.wiki_url, timeout=30000)
        time.sleep(5)

        if "哎呀" in page.content() or "无权限" in page.content():
            raise RuntimeError(
                "无法访问 wiki 页面。请确认账号有该页面的访问权限。"
            )
        print("wiki 页面访问成功。")

    # ------------------------------------------------------------------
    # Tree extraction
    # ------------------------------------------------------------------

    def _extract_menu_tree(self, page) -> list:
        """Expand all sidebar nodes and extract the full hierarchy."""
        print("展开侧边栏所有节点...")

        for round_num in range(12):
            closed = page.evaluate(
                "() => document.querySelectorAll('.ant-tree-switcher_close').length"
            )
            if closed == 0:
                print(f"  第 {round_num + 1} 轮: 全部已展开")
                break

            print(f"  第 {round_num + 1} 轮: 展开 {closed} 个折叠节点...")
            page.evaluate(
                """() => {
                    document.querySelectorAll('.ant-tree-switcher_close')
                        .forEach(el => el.click());
                }"""
            )
            time.sleep(1.2)

        # Final sweep
        page.evaluate(
            """() => {
                document.querySelectorAll('.ant-tree-switcher_close')
                    .forEach(el => el.click());
            }"""
        )
        time.sleep(1.5)

        menu_items = page.evaluate(
            """() => {
                function extract(ul) {
                    const items = [];
                    if (!ul) return items;
                    ul.querySelectorAll(':scope > li').forEach(li => {
                        const box = li.querySelector('.menuAction__box');
                        const title = box
                            ? (box.getAttribute('title') || box.innerText.trim())
                            : '';
                        const child = li.querySelector(':scope > ul');
                        items.push({
                            title,
                            children: child ? extract(child) : [],
                        });
                    });
                    return items;
                }
                const root = document.querySelector('.wikiMenu__tree .ant-tree');
                return root ? extract(root) : [];
            }"""
        )

        all_sections = []

        def flatten(items, level):
            for item in items:
                is_leaf = len(item["children"]) == 0
                all_sections.append(
                    {
                        "title": item["title"],
                        "level": level,
                        "isLeaf": is_leaf,
                        "content_html": "",
                        "content": "",
                    }
                )
                if item["children"]:
                    flatten(item["children"], level + 1)

        flatten(menu_items, 0)
        leaf_count = sum(1 for s in all_sections if s["isLeaf"])
        print(f"  共 {len(all_sections)} 个节点 ({leaf_count} 个叶子节点)")
        return all_sections

    # ------------------------------------------------------------------
    # Content scraping
    # ------------------------------------------------------------------

    def _scrape_sections(self, page, all_sections):
        """Click every section in the sidebar and capture rendered HTML content."""
        print(f"抓取 {len(all_sections)} 个章节内容...")

        success = 0
        for i, section in enumerate(all_sections):
            title = section["title"]
            label = f"[{i + 1}/{len(all_sections)}]"

            clicked = page.evaluate(
                """(t) => {
                    for (const box of document.querySelectorAll('.menuAction__box')) {
                        if (
                            (box.getAttribute('title') || box.innerText.trim()) === t
                        ) {
                            box.click();
                            return true;
                        }
                    }
                    return false;
                }""",
                title,
            )

            if not clicked:
                print(f"  {label} {title} — 未找到")
                continue

            time.sleep(0.6)

            html_content = page.evaluate(
                """() => {
                    const c = document.getElementById('rick-viewer-container');
                    if (!c || c.querySelector('.client-only-placeholder')) return '';
                    const e = c.querySelector('[data-slate-editor], .editor-content-container');
                    return e ? e.innerHTML : c.innerHTML;
                }"""
            )
            text_content = page.evaluate(
                """() => {
                    const c = document.getElementById('rick-viewer-container');
                    if (!c || c.querySelector('.client-only-placeholder')) return '';
                    const e = c.querySelector('[data-slate-editor], .editor-content-container');
                    return e ? (e.innerText || '') : (c.innerText || '');
                }"""
            )

            section["content_html"] = html_content
            section["content"] = text_content

            if text_content.strip():
                success += 1
                preview = text_content.strip()[:60].replace("\n", " ")
                print(f"  {label} {title} [{len(text_content)} chars] {preview}...")
            else:
                print(f"  {label} {title} — (空)")

        print(f"抓取完成: {success}/{len(all_sections)} 个章节有内容")
