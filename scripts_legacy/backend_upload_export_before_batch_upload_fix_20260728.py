from __future__ import annotations

import argparse
import time
from pathlib import Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="新人价后台查价：自动上传SKU拆分文件、开始查询，并点击批量导出到邮箱")
    parser.add_argument(
        "--url",
        default="https://yx.jd.com/user-operate#/customize/findLowestActivity?pathkey=goodspool",
        help="新人价查价工具页面地址",
    )
    parser.add_argument(
        "--sku-dir",
        default=r"D:\hansiying.1\Desktop\查价前sku",
        type=Path,
        help="拆分后的SKU文件夹",
    )
    parser.add_argument("--pattern", default="8月第一批查价前sku_PART*.xlsx", help="待上传文件匹配规则")
    parser.add_argument("--only", help="只处理某一个文件名，例如：8月第一批查价前sku_PART01.xlsx")
    parser.add_argument(
        "--profile-dir",
        default=r"D:\hansiying.1\Desktop\查价前sku\后台浏览器登录缓存",
        type=Path,
        help="后台浏览器登录缓存目录；首次运行需手动登录，之后复用登录态",
    )
    parser.add_argument(
        "--log-file",
        default=r"D:\hansiying.1\Desktop\查价前sku\后台查价自动上传日志.txt",
        type=Path,
        help="运行日志保存位置",
    )
    parser.add_argument("--wait-after-query", default=45, type=int, help="点击开始查询后等待秒数，用于等待结果刷新")
    parser.add_argument("--wait-after-export", default=10, type=int, help="点击批量导出后等待秒数")
    parser.add_argument("--pause-after-each", action="store_true", help="每个文件批量导出后暂停，方便人工检查页面")
    parser.add_argument("--headless", action="store_true", help="无界面运行；首次登录和调试不建议开启")
    return parser.parse_args()


def ensure_playwright() -> None:
    try:
        import playwright  # noqa: F401
    except Exception as exc:
        raise RuntimeError("缺少 playwright。请先运行：python -m pip install playwright && python -m playwright install chromium") from exc


def get_files(sku_dir: Path, pattern: str, only: str | None) -> list[Path]:
    files = [sku_dir / only] if only else sorted(sku_dir.glob(pattern))
    files = [path for path in files if path.is_file()]
    if not files:
        raise FileNotFoundError(f"未找到待上传文件：{sku_dir / pattern}")
    return files


def log(message: str, log_file: Path) -> None:
    print(message)
    log_file.parent.mkdir(parents=True, exist_ok=True)
    with log_file.open("a", encoding="utf-8") as file:
        file.write(message + "\n")


def click_first_available(page, selectors: list[tuple[str, str]], timeout: int = 8000) -> bool:
    for selector_type, selector_value in selectors:
        try:
            if selector_type == "role_button":
                locator = page.get_by_role("button", name=selector_value).first
            elif selector_type == "text":
                locator = page.get_by_text(selector_value, exact=False).first
            elif selector_type == "css":
                locator = page.locator(selector_value).first
            else:
                continue
            locator.wait_for(state="visible", timeout=timeout)
            locator.click(timeout=timeout)
            return True
        except Exception:
            continue
    return False


def page_and_frames(page):
    targets = [page]
    for frame in page.frames:
        if frame not in targets:
            targets.append(frame)
    return targets


def save_debug_snapshot(page, debug_dir: Path, reason: str) -> None:
    debug_dir.mkdir(parents=True, exist_ok=True)
    safe_reason = "".join(ch if ch.isalnum() else "_" for ch in reason)[:40]
    try:
        page.screenshot(path=str(debug_dir / f"{safe_reason}.png"), full_page=True)
    except Exception:
        pass
    try:
        (debug_dir / f"{safe_reason}.html").write_text(page.content(), encoding="utf-8")
    except Exception:
        pass
    try:
        lines = []
        for target in page_and_frames(page):
            try:
                buttons = target.locator("button, [role=button], input[type=file], .ant-upload, .el-upload")
                count = min(buttons.count(), 80)
                lines.append(f"TARGET={getattr(target, 'url', 'page')}")
                for index in range(count):
                    item = buttons.nth(index)
                    text = ""
                    try:
                        text = item.inner_text(timeout=500).strip()
                    except Exception:
                        pass
                    title = item.get_attribute("title") or ""
                    aria = item.get_attribute("aria-label") or ""
                    tag = item.evaluate("el => el.tagName").strip()
                    lines.append(f"{index}: tag={tag} text={text} title={title} aria={aria}")
            except Exception as exc:
                lines.append(f"TARGET_ERROR={exc}")
        (debug_dir / f"{safe_reason}_elements.txt").write_text("\n".join(lines), encoding="utf-8")
    except Exception:
        pass


def click_first_available_anywhere(page, selectors: list[tuple[str, str]], timeout: int = 8000) -> bool:
    for target in page_and_frames(page):
        if click_first_available(target, selectors, timeout=timeout):
            return True
    return False


def upload_file(page, file_path: Path, debug_dir: Path) -> None:
    # 直接设置网页/iframe中的上传控件，不依赖 Windows 文件选择窗口。
    for target in page_and_frames(page):
        try:
            inputs = target.locator("input[type=file]")
            count = inputs.count()
            for index in range(count):
                try:
                    inputs.nth(index).set_input_files(str(file_path))
                    return
                except Exception:
                    continue
        except Exception:
            continue

    # 如果上传控件只有点击后才出现，则捕获文件选择器。
    selectors = [
        ("role_button", "上传文件"),
        ("text", "上传文件"),
        ("text", "点击上传"),
        ("text", "点击上传 Excel 文件"),
        ("text", "上传Excel"),
        ("text", "上传 Excel"),
        ("css", "button:has-text('上传')"),
        ("css", ".ant-upload"),
        ("css", ".el-upload"),
        ("css", "[class*='upload']"),
    ]
    try:
        with page.expect_file_chooser(timeout=15000) as chooser_info:
            clicked = click_first_available_anywhere(page, selectors, timeout=10000)
            if not clicked:
                save_debug_snapshot(page, debug_dir, "upload_button_not_found")
                raise RuntimeError(f"找不到上传文件按钮。已保存调试文件到：{debug_dir}")
        chooser_info.value.set_files(str(file_path))
        return
    except Exception as exc:
        save_debug_snapshot(page, debug_dir, "upload_failed")
        raise RuntimeError(f"上传文件失败：{exc}。已保存调试文件到：{debug_dir}") from exc


def wait_query_finished(page, wait_after_query: int) -> None:
    try:
        page.wait_for_load_state("networkidle", timeout=120000)
    except Exception:
        pass
    time.sleep(wait_after_query)
    try:
        page.get_by_text("目前查找到sku价格信息", exact=False).first.wait_for(state="visible", timeout=30000)
    except Exception:
        pass


def confirm_export_if_needed(page) -> None:
    # 某些页面点击批量导出后会弹确认框；没有弹窗时会自动跳过。
    for label in ["确定", "确认", "知道了", "我知道了"]:
        click_first_available(page, [("role_button", label), ("text", label)], timeout=3000)


def main() -> int:
    args = parse_args()
    ensure_playwright()
    from playwright.sync_api import sync_playwright

    files = get_files(args.sku_dir, args.pattern, args.only)
    args.profile_dir.mkdir(parents=True, exist_ok=True)
    args.log_file.write_text("", encoding="utf-8")

    with sync_playwright() as playwright:
        context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(args.profile_dir),
            headless=args.headless,
            accept_downloads=False,
            viewport={"width": 1500, "height": 900},
            args=["--start-maximized"],
        )
        page = context.pages[0] if context.pages else context.new_page()
        page.set_default_timeout(30000)
        page.goto(args.url, wait_until="domcontentloaded")

        log("如果页面要求登录，请在打开的浏览器里完成登录；脚本会等待并继续。", args.log_file)
        click_first_available(page, [("text", "查价工具")], timeout=120000)
        click_first_available_anywhere(page, [("text", "批量上传")], timeout=15000)

        for index, file_path in enumerate(files, start=1):
            log(f"[{index}/{len(files)}] 准备上传：{file_path.name}", args.log_file)
            click_first_available_anywhere(page, [("role_button", "清空"), ("text", "清空")], timeout=3000)
            upload_file(page, file_path, args.log_file.parent / "debug_backend")

            started = click_first_available_anywhere(
                page,
                [("role_button", "开始查询"), ("text", "开始查询"), ("css", "button:has-text('开始查询')")],
                timeout=15000,
            )
            if not started:
                raise RuntimeError("找不到“开始查询”按钮。")
            log(f"[{index}/{len(files)}] 已点击开始查询：{file_path.name}", args.log_file)
            wait_query_finished(page, args.wait_after_query)
            log(f"[{index}/{len(files)}] 查询等待结束，准备点击批量导出。", args.log_file)

            exported = click_first_available_anywhere(
                page,
                [("role_button", "批量导出"), ("text", "批量导出"), ("css", "button:has-text('批量导出')")],
                timeout=30000,
            )
            if not exported:
                raise RuntimeError("找不到“批量导出”按钮。")
            confirm_export_if_needed(page)
            time.sleep(args.wait_after_export)
            log(f"[{index}/{len(files)}] 已点击批量导出，等待邮件发送：{file_path.name}", args.log_file)

            if args.pause_after_each:
                input("已完成当前文件上传、查询和批量导出。检查页面后，按回车继续下一个文件...")

        log("全部文件已完成上传、开始查询并点击批量导出。后台网页操作结束。", args.log_file)
        input("后台网页操作结束。确认后按回车关闭浏览器...")
        context.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

