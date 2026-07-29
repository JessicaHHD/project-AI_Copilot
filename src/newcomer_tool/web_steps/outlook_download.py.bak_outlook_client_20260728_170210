from __future__ import annotations

import argparse
import re
import time
from pathlib import Path


ATTACHMENT_PATTERN = re.compile(r"\.(xlsx|xls|csv|zip)$", re.IGNORECASE)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Outlook网页邮箱：打开查价结果邮件并下载附件")
    parser.add_argument("--mail-url", default="https://outlook.office.com/mail/", help="Outlook网页邮箱地址")
    parser.add_argument(
        "--download-dir",
        default=r"D:\hansiying.1\Desktop\查价前sku\查价结果导出",
        type=Path,
        help="附件下载保存文件夹",
    )
    parser.add_argument(
        "--profile-dir",
        default=r"D:\hansiying.1\Desktop\查价前sku\邮箱浏览器登录缓存",
        type=Path,
        help="邮箱浏览器登录缓存目录；首次运行需手动登录，之后复用登录态",
    )
    parser.add_argument("--search", default="查价 新人价", help="Outlook搜索关键词，可按实际邮件标题修改")
    parser.add_argument("--max-mails", default=20, type=int, help="最多尝试打开多少封邮件")
    parser.add_argument("--manual", action="store_true", help="手动模式：你先在邮箱页面筛出邮件，再按回车让脚本下载当前列表附件")
    parser.add_argument("--headless", action="store_true", help="无界面运行；首次登录不建议开启")
    return parser.parse_args()


def ensure_playwright() -> None:
    try:
        import playwright  # noqa: F401
    except Exception as exc:
        raise RuntimeError("缺少 playwright。请先运行：python -m pip install playwright && python -m playwright install chromium") from exc


def click_if_visible(page, labels: list[str], timeout: int = 3000) -> bool:
    for label in labels:
        for locator in [page.get_by_role("button", name=label).first, page.get_by_text(label, exact=False).first]:
            try:
                locator.wait_for(state="visible", timeout=timeout)
                locator.click(timeout=timeout)
                return True
            except Exception:
                continue
    return False


def wait_for_login(page) -> None:
    print("如果 Outlook 要求登录，请在打开的浏览器里完成登录。登录后脚本会继续。")
    candidates = [
        "input[aria-label*='搜索']",
        "input[placeholder*='搜索']",
        "input[aria-label*='Search']",
        "input[placeholder*='Search']",
        "[role='searchbox']",
        "div[role='main']",
    ]
    for selector in candidates:
        try:
            page.locator(selector).first.wait_for(state="visible", timeout=180000)
            return
        except Exception:
            continue


def search_mail(page, query: str) -> bool:
    selectors = [
        "input[aria-label*='搜索']",
        "input[placeholder*='搜索']",
        "input[aria-label*='Search']",
        "input[placeholder*='Search']",
        "[role='searchbox']",
    ]
    for selector in selectors:
        try:
            box = page.locator(selector).first
            box.wait_for(state="visible", timeout=8000)
            box.click()
            box.fill(query)
            box.press("Enter")
            time.sleep(8)
            return True
        except Exception:
            continue
    return False


def visible_message_rows(page):
    selectors = [
        "div[role='option']",
        "div[role='listitem']",
        "[data-convid]",
        "[aria-label*='附件']",
        "[aria-label*='attachment']",
    ]
    best = None
    best_count = 0
    for selector in selectors:
        try:
            locator = page.locator(selector)
            count = min(locator.count(), 100)
            visible_count = 0
            for index in range(count):
                try:
                    if locator.nth(index).is_visible(timeout=300):
                        visible_count += 1
                except Exception:
                    pass
            if visible_count > best_count:
                best = locator
                best_count = visible_count
        except Exception:
            continue
    return best, best_count


def download_current_mail_attachments(page, download_dir: Path, mail_index: int) -> int:
    downloaded = 0
    attachment_locators = [
        page.locator("a, button, [role='button']").filter(has_text=ATTACHMENT_PATTERN),
        page.locator("[title$='.xlsx'], [title$='.xls'], [title$='.csv'], [title$='.zip']"),
        page.locator("[aria-label$='.xlsx'], [aria-label$='.xls'], [aria-label$='.csv'], [aria-label$='.zip']"),
    ]

    seen_names: set[str] = set()
    for attachments in attachment_locators:
        try:
            count = min(attachments.count(), 20)
        except Exception:
            continue
        for index in range(count):
            item = attachments.nth(index)
            try:
                if not item.is_visible(timeout=500):
                    continue
                name = (item.inner_text(timeout=1000) or item.get_attribute("title") or item.get_attribute("aria-label") or f"attachment_{mail_index}_{index}").strip()
                if name in seen_names:
                    continue
                seen_names.add(name)

                try:
                    with page.expect_download(timeout=20000) as download_info:
                        item.click(timeout=5000)
                    download = download_info.value
                except Exception:
                    item.click(timeout=5000)
                    with page.expect_download(timeout=30000) as download_info:
                        if not click_if_visible(page, ["下载", "Download", "全部下载", "Download all"], timeout=8000):
                            raise RuntimeError("打开附件后未找到下载按钮")
                    download = download_info.value

                suffix = Path(download.suggested_filename).suffix or Path(name).suffix or ".xlsx"
                safe_name = re.sub(r"[\\/:*?\"<>|]+", "_", Path(download.suggested_filename).stem or Path(name).stem or f"mail_{mail_index}_attachment_{index}")
                save_path = download_dir / f"{mail_index:03d}_{safe_name}{suffix}"
                download.save_as(str(save_path))
                print(f"已下载：{save_path}")
                downloaded += 1
            except Exception as exc:
                print(f"附件尝试下载失败：{exc}")
                continue
    return downloaded


def main() -> int:
    args = parse_args()
    ensure_playwright()
    from playwright.sync_api import sync_playwright

    args.download_dir.mkdir(parents=True, exist_ok=True)
    args.profile_dir.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as playwright:
        context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(args.profile_dir),
            headless=args.headless,
            accept_downloads=True,
            viewport={"width": 1500, "height": 900},
            args=["--start-maximized"],
        )
        page = context.pages[0] if context.pages else context.new_page()
        page.set_default_timeout(30000)
        page.goto(args.mail_url, wait_until="domcontentloaded")
        wait_for_login(page)

        if args.manual:
            input("请在邮箱页面手动搜索/筛选出查价结果邮件列表，然后回到此窗口按回车继续...")
        else:
            if not search_mail(page, args.search):
                input("脚本没有找到搜索框。请手动搜索出查价结果邮件列表，然后回到此窗口按回车继续...")

        rows, count = visible_message_rows(page)
        if not rows or count == 0:
            input("没有自动识别到邮件列表。请点开第一封查价结果邮件，然后按回车，脚本会尝试下载当前邮件附件...")
            total = download_current_mail_attachments(page, args.download_dir, 1)
            print(f"完成，下载附件数：{total}")
            context.close()
            return 0

        total_downloaded = 0
        attempts = min(count, args.max_mails)
        print(f"识别到约 {count} 条邮件记录，准备尝试前 {attempts} 条。")
        for index in range(attempts):
            try:
                rows.nth(index).click(timeout=8000)
                time.sleep(4)
                total_downloaded += download_current_mail_attachments(page, args.download_dir, index + 1)
            except Exception as exc:
                print(f"第 {index + 1} 封邮件处理失败：{exc}")
                continue

        context.close()
    print(f"全部完成，下载附件数：{total_downloaded}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
