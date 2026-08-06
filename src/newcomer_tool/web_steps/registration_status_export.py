from __future__ import annotations

import argparse
import time
from datetime import datetime, timedelta
from pathlib import Path
from types import SimpleNamespace
from typing import Any

from backend_upload_export import (
    ensure_playwright,
    fast_click_text_anywhere,
    handle_inner_erp_login_if_needed,
    log,
    page_and_frames,
    save_debug_snapshot,
    timed_step,
    visible_text_exists_anywhere,
)
from outlook_download import (
    TIME_FORMAT,
    download_url,
    flatten_links,
    parse_received_time,
    powershell_extract_mails,
    write_link_report,
)
from registration_submit import (
    cachebust_url,
    enter_first_order_registration,
    enter_registration_home,
    reset_scroll_anywhere,
    target_month_labels,
    wait_for_any_text,
)


DEFAULT_URL = "https://yx.jd.com/user-operate#/customize/newCanSignUpActivity?pathkey=goodspool"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="新人价已报名提报情况：后台触发下载，并从 Outlook 下载导出 Excel")
    parser.add_argument("--url", default=DEFAULT_URL, help="新人价补贴报名页面地址")
    parser.add_argument(
        "--mode",
        choices=["dry_run", "export_only", "export_and_download", "download_latest_only"],
        default="dry_run",
        help="dry_run 到下载页不点击下载；export_only 点击下载不抓邮件；export_and_download 点击下载并从 Outlook 下载；download_latest_only 只下载最新匹配邮件",
    )
    parser.add_argument("--confirm-export", default="", help="点击后台下载必须传入固定确认语：确认导出")
    parser.add_argument("--profile-dir", type=Path, required=True, help="后台浏览器登录缓存目录")
    parser.add_argument("--log-file", type=Path, required=True, help="运行日志保存位置")
    parser.add_argument("--download-dir", type=Path, required=True, help="Outlook 链接下载保存目录")
    parser.add_argument("--month-threshold-day", type=int, default=25, help="大于等于该日期时优先选择下月计划")
    parser.add_argument("--mail-subject-keyword", default="商品下载完成通知", help="Outlook 邮件主题关键词")
    parser.add_argument("--mail-sender-keyword", default="增长权益", help="Outlook 发件人关键词；为空则不限制")
    parser.add_argument("--url-keyword", default=".xlsx", help="邮件下载链接关键词")
    parser.add_argument("--max-mails", type=int, default=120, help="最多扫描多少封 Outlook 邮件")
    parser.add_argument("--mail-timeout", type=int, default=1200, help="等待邮件到达最长秒数")
    parser.add_argument("--mail-poll-seconds", type=int, default=30, help="轮询 Outlook 间隔秒数")
    parser.add_argument("--mail-window-minutes", type=int, default=30, help="只接受距离点击下载时间多少分钟内的邮件")
    parser.add_argument("--download-timeout", type=int, default=180, help="下载链接超时秒数")
    parser.add_argument("--headless", action="store_true", help="无界面运行；首次登录和调试不建议开启")
    return parser.parse_args()


def click_registered_tab(page, debug_dir: Path, log_file: Path) -> None:
    script = r"""
        () => {
            const visible = (el) => {
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return style && style.visibility !== 'hidden' && style.display !== 'none' && rect.width > 0 && rect.height > 0;
            };
            const norm = (text) => String(text || '').replace(/\s+/g, '');
            const textOf = (el) => el.innerText || el.textContent || '';
            const nodes = Array.from(document.querySelectorAll('button, a, [role=tab], [role=button], div, span'))
                .filter(el => visible(el) && norm(textOf(el)) === '已报名')
                .sort((a, b) => a.getBoundingClientRect().top - b.getBoundingClientRect().top || a.getBoundingClientRect().left - b.getBoundingClientRect().left);
            const target = nodes[0];
            if (!target) return null;
            const clickable = target.closest('button, a, [role=tab], [role=button]') || target;
            clickable.scrollIntoView({ block: 'center', inline: 'center' });
            const rect = clickable.getBoundingClientRect();
            clickable.click();
            return { text: textOf(clickable).replace(/\s+/g, ' ').trim(), tag: clickable.tagName, x: Math.round(rect.left + rect.width / 2), y: Math.round(rect.top + rect.height / 2) };
        }
    """
    clicked = False
    for target in page_and_frames(page):
        try:
            result = target.evaluate(script)
            if result:
                log(f"已点击“已报名”tab：文本={result.get('text')}，标签={result.get('tag')}，坐标=({result.get('x')},{result.get('y')})", log_file)
                clicked = True
                break
        except Exception:
            continue
    if not clicked:
        clicked = fast_click_text_anywhere(page, ["已报名"], timeout_ms=1200, clickable_only=False)
    if not clicked:
        save_debug_snapshot(page, debug_dir, "registered_tab_not_found")
        raise RuntimeError(f"找不到“已报名”tab。已保存调试文件到：{debug_dir}")
    if not wait_for_any_text(page, ["补贴计划名称", "查看"], timeout_seconds=10):
        save_debug_snapshot(page, debug_dir, "registered_plan_list_not_ready")
        raise RuntimeError(f"已报名计划列表未加载完成。已保存调试文件到：{debug_dir}")


def registered_plan_ready(page, month_labels: list[str]) -> bool:
    script = r"""
        ({ monthLabels }) => {
            const visible = (el) => {
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return style && style.visibility !== 'hidden' && style.display !== 'none' && rect.width > 0 && rect.height > 0;
            };
            const norm = (text) => String(text || '').replace(/\s+/g, '');
            const bodyText = norm(document.body.innerText || document.body.textContent || '');
            const hasMonth = monthLabels.some(label => bodyText.includes(norm(label)));
            const hasView = Array.from(document.querySelectorAll('button, a, [role=button], span, div'))
                .some(el => visible(el) && norm(el.innerText || el.textContent || '').includes('查看'));
            return hasMonth && hasView;
        }
    """
    for target in page_and_frames(page):
        try:
            if target.evaluate(script, {"monthLabels": month_labels}):
                return True
        except Exception:
            continue
    return False


def wait_registered_plan_ready(page, month_labels: list[str], debug_dir: Path, log_file: Path, timeout_seconds: int = 20) -> None:
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        if registered_plan_ready(page, month_labels):
            log(f"已检测到目标月份已报名计划和“查看”：{'、'.join(month_labels)}", log_file)
            return
        time.sleep(0.5)
    save_debug_snapshot(page, debug_dir, "registered_plan_not_ready")
    raise RuntimeError(f"等待目标月份已报名计划超时：{'、'.join(month_labels)}。已保存调试文件到：{debug_dir}")


def wait_after_view_click(page, before_url: str, timeout_seconds: int = 15) -> bool:
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        if visible_text_exists_anywhere(page, ["下载", "提报状态", "补贴状态", "skuid", "查看详情"], timeout_ms=500):
            return True
        if page.url != before_url and visible_text_exists_anywhere(page, ["下载", "提报状态", "补贴状态"], timeout_ms=500):
            return True
        time.sleep(0.3)
    return False


def click_registered_plan_view(page, month_labels: list[str], debug_dir: Path, log_file: Path) -> str:
    script = r"""
        ({ monthLabels }) => {
            const visible = (el) => {
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return style && style.visibility !== 'hidden' && style.display !== 'none' && rect.width > 0 && rect.height > 0 && rect.bottom > 0 && rect.right > 0 && rect.top < window.innerHeight && rect.left < window.innerWidth;
            };
            const textOf = (el) => (el.innerText || el.textContent || '').trim();
            const norm = (text) => String(text || '').replace(/\s+/g, '');
            const clickableOf = (el) => el.closest('button, a, [role=button], .ant-btn, .el-button, [class*=btn], [class*=button]') || el;
            const dispatchClick = (el) => {
                const clickable = clickableOf(el);
                clickable.scrollIntoView({ block: 'center', inline: 'center' });
                const rect = clickable.getBoundingClientRect();
                const x = rect.left + rect.width / 2;
                const y = rect.top + rect.height / 2;
                clickable.focus && clickable.focus();
                for (const type of ['mouseover', 'mousemove', 'mousedown', 'mouseup', 'click']) {
                    clickable.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window, clientX: x, clientY: y, button: 0 }));
                }
                clickable.click && clickable.click();
                return { tag: clickable.tagName, text: textOf(clickable), x: Math.round(x), y: Math.round(y) };
            };
            const rowSelectors = ['tr', '[role=row]', '.ant-table-row', '.el-table__row', '[class*=table] [class*=row]', '[class*=Table] [class*=Row]'];
            const actionSelector = 'button, a, [role=button], .ant-btn, .el-button, [class*=btn], [class*=button], span, div';
            const candidates = [];
            for (const monthLabel of monthLabels) {
                const targetMonth = norm(monthLabel);
                for (const rowSelector of rowSelectors) {
                    for (const row of Array.from(document.querySelectorAll(rowSelector))) {
                        if (!visible(row)) continue;
                        const rowText = norm(textOf(row));
                        if (!rowText.includes(targetMonth) || !rowText.includes('查看')) continue;
                        for (const action of Array.from(row.querySelectorAll(actionSelector))) {
                            if (visible(action) && norm(textOf(action)).includes('查看')) {
                                const rect = action.getBoundingClientRect();
                                candidates.push({ action, row, monthLabel, method: 'visible_row', right: rect.right, top: rect.top });
                            }
                        }
                    }
                }
                if (candidates.length) break;
            }
            if (!candidates.length) {
                for (const monthLabel of monthLabels) {
                    const targetMonth = norm(monthLabel);
                    const monthItems = Array.from(document.querySelectorAll('body *'))
                        .filter(el => visible(el) && norm(textOf(el)).includes(targetMonth))
                        .map(el => ({ el, rect: el.getBoundingClientRect(), text: textOf(el) }))
                        .filter(item => item.text.length <= 120)
                        .sort((a, b) => (a.rect.width * a.rect.height) - (b.rect.width * b.rect.height));
                    const actions = Array.from(document.querySelectorAll(actionSelector))
                        .filter(el => visible(el) && norm(textOf(el)).includes('查看'))
                        .map(el => ({ el, rect: el.getBoundingClientRect(), text: textOf(el) }));
                    for (const monthItem of monthItems) {
                        const sameLine = actions
                            .filter(action => Math.abs((action.rect.top + action.rect.height / 2) - (monthItem.rect.top + monthItem.rect.height / 2)) <= 32 && action.rect.left > monthItem.rect.left)
                            .sort((a, b) => b.rect.left - a.rect.left)[0];
                        if (sameLine) candidates.push({ action: sameLine.el, row: monthItem.el, monthLabel, method: 'visible_same_line', right: sameLine.rect.right, top: sameLine.rect.top });
                    }
                    if (candidates.length) break;
                }
            }
            candidates.sort((a, b) => b.right - a.right || a.top - b.top);
            const candidate = candidates[0];
            if (!candidate) return null;
            const clicked = dispatchClick(candidate.action);
            return { monthLabel: candidate.monthLabel, method: candidate.method, rowText: textOf(candidate.row).replace(/\s+/g, ' ').slice(0, 180), ...clicked };
        }
    """
    for target in page_and_frames(page):
        before_url = page.url
        try:
            result = target.evaluate(script, {"monthLabels": month_labels})
        except Exception:
            continue
        if not result:
            continue
        selected = str(result.get("monthLabel", ""))
        log(
            f"已点击已报名计划“查看”：月份={selected}，方式={result.get('method')}，"
            f"文本={result.get('text')}，坐标=({result.get('x')},{result.get('y')})",
            log_file,
        )
        if wait_after_view_click(page, before_url, timeout_seconds=15):
            log(f"已进入提报情况详情页：{selected}", log_file)
            return selected
    save_debug_snapshot(page, debug_dir, "registered_plan_view_not_found")
    raise RuntimeError(f"未找到目标月份已报名计划的“查看”：{'、'.join(month_labels)}。已保存调试文件到：{debug_dir}")


def click_download_button(page, debug_dir: Path, log_file: Path) -> datetime:
    script = r"""
        () => {
            const visible = (el) => {
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return style && style.visibility !== 'hidden' && style.display !== 'none' && rect.width > 0 && rect.height > 0;
            };
            const norm = (text) => String(text || '').replace(/\s+/g, '');
            const textOf = (el) => el.innerText || el.textContent || el.value || '';
            const nodes = Array.from(document.querySelectorAll('button, [role=button], a, .ant-btn, .el-button, input[type=button]'))
                .filter(el => visible(el) && !el.disabled && !el.getAttribute('disabled') && norm(textOf(el)) === '下载')
                .map(el => ({ el, rect: el.getBoundingClientRect(), text: textOf(el) }))
                .sort((a, b) => a.rect.top - b.rect.top || b.rect.right - a.rect.right);
            const target = nodes[0];
            if (!target) return null;
            target.el.scrollIntoView({ block: 'center', inline: 'center' });
            const rect = target.el.getBoundingClientRect();
            target.el.click();
            return { text: String(target.text || '').replace(/\s+/g, ' ').trim(), tag: target.el.tagName, x: Math.round(rect.left + rect.width / 2), y: Math.round(rect.top + rect.height / 2) };
        }
    """
    for target in page_and_frames(page):
        try:
            trigger_time = datetime.now()
            result = target.evaluate(script)
            if result:
                log(f"已点击“下载”按钮：文本={result.get('text')}，标签={result.get('tag')}，坐标=({result.get('x')},{result.get('y')})", log_file)
                time.sleep(0.8)
                for label in ["确定", "确认", "我知道了"]:
                    fast_click_text_anywhere(page, [label], timeout_ms=300, clickable_only=True)
                return trigger_time
        except Exception:
            continue
    save_debug_snapshot(page, debug_dir, "status_download_button_not_found")
    raise RuntimeError(f"找不到“下载”按钮。已保存调试文件到：{debug_dir}")


def find_latest_outlook_link(args: argparse.Namespace) -> dict[str, str] | None:
    outlook_args = SimpleNamespace(
        max_mails=args.max_mails,
        subject_keyword=args.mail_subject_keyword,
        sender_keyword=args.mail_sender_keyword,
        url_keyword=args.url_keyword,
        include_read=True,
    )
    mails = powershell_extract_mails(outlook_args)
    rows = flatten_links(mails)
    if not rows:
        return None
    return rows[0]


def download_latest_outlook_file(args: argparse.Namespace, log_file: Path) -> Path:
    row = find_latest_outlook_link(args)
    if not row:
        raise RuntimeError(
            f"未找到 Outlook 邮件：主题包含“{args.mail_subject_keyword}”、"
            f"发件人包含“{args.mail_sender_keyword or '不限制'}”、链接包含“{args.url_keyword or '不限制'}”。"
        )
    log(f"已找到最新匹配 Outlook 邮件：{row.get('received_time')} | {row.get('subject')} | {row.get('sender')}", log_file)
    report = write_link_report([row], log_file)
    log(f"链接报告：{report}", log_file)
    target = download_url(row["url"], args.download_dir, args.download_timeout)
    log(f"提报情况 Excel 下载完成：{target}", log_file)
    return target


def find_closest_outlook_link(args: argparse.Namespace, trigger_time: datetime) -> dict[str, str] | None:
    outlook_args = SimpleNamespace(
        max_mails=args.max_mails,
        subject_keyword=args.mail_subject_keyword,
        sender_keyword=args.mail_sender_keyword,
        url_keyword=args.url_keyword,
        include_read=True,
    )
    mails = powershell_extract_mails(outlook_args)
    rows = flatten_links(mails)
    if not rows:
        return None

    window_end = trigger_time + timedelta(minutes=max(1, args.mail_window_minutes))
    earliest_allowed = trigger_time - timedelta(minutes=1)
    candidates: list[tuple[int, float, dict[str, str]]] = []
    for row in rows:
        received_time = parse_received_time(row.get("received_time", ""))
        if not received_time:
            continue
        if received_time < earliest_allowed or received_time > window_end:
            continue
        after_trigger = received_time >= trigger_time
        delta = abs((received_time - trigger_time).total_seconds())
        # 优先选择点击下载之后收到的邮件；1分钟内的旧邮件仅作为邮箱时间误差兜底。
        candidates.append((0 if after_trigger else 1, delta, row))
    if not candidates:
        return None
    candidates.sort(key=lambda item: (item[0], item[1]))
    return candidates[0][2]


def wait_and_download_outlook_file(args: argparse.Namespace, trigger_time: datetime, log_file: Path) -> Path:
    deadline = time.time() + args.mail_timeout
    log(
        f"开始后台轮询 Outlook：最长等待 {args.mail_timeout} 秒，轮询间隔 {args.mail_poll_seconds} 秒；"
        f"只匹配 {trigger_time.strftime(TIME_FORMAT)} 之后、主题包含“{args.mail_subject_keyword}”、发件人包含“{args.mail_sender_keyword or '不限制'}”的 xlsx 链接。",
        log_file,
    )
    while time.time() < deadline:
        row = find_closest_outlook_link(args, trigger_time)
        if row:
            log(
                f"已找到点击下载后的 Outlook 邮件：{row.get('received_time')} | {row.get('subject')}",
                log_file,
            )
            report = write_link_report([row], log_file)
            log(f"链接报告：{report}", log_file)
            target = download_url(row["url"], args.download_dir, args.download_timeout)
            log(f"提报情况 Excel 下载完成：{target}", log_file)
            return target
        remaining = max(0, int(deadline - time.time()))
        log(f"暂未找到新邮件，剩余等待约 {remaining} 秒；{args.mail_poll_seconds} 秒后重试。", log_file)
        time.sleep(max(3, args.mail_poll_seconds))
    raise RuntimeError(f"等待 Outlook 邮件超时：{args.mail_timeout} 秒。请检查邮件主题、下载是否触发成功。")


def process_export(page, args: argparse.Namespace, month_labels: list[str], debug_dir: Path) -> None:
    with timed_step("打开新人价补贴报名页面", args.log_file):
        page.goto(cachebust_url(args.url), wait_until="domcontentloaded")
    with timed_step("处理内网ERP登录", args.log_file):
        handle_inner_erp_login_if_needed(page, args.log_file, max_wait_seconds=60)
    with timed_step("进入新人补贴报名首页", args.log_file):
        enter_registration_home(page, debug_dir, args.log_file, args.url)
    with timed_step("进入首单新人价计划列表", args.log_file):
        enter_first_order_registration(page, debug_dir)
    with timed_step("切换到已报名计划列表", args.log_file):
        click_registered_tab(page, debug_dir, args.log_file)
        wait_registered_plan_ready(page, month_labels, debug_dir, args.log_file, timeout_seconds=20)
    with timed_step("选择补贴计划并查看", args.log_file):
        reset_scroll_anywhere(page)
        selected_month = click_registered_plan_view(page, month_labels, debug_dir, args.log_file)
        save_debug_snapshot(page, debug_dir, "status_detail_ready")

    if args.mode == "dry_run":
        log(f"dry_run 演练完成：已进入提报情况详情页，未点击下载。目标月份：{selected_month}", args.log_file)
        return

    with timed_step("点击下载并触发邮件", args.log_file):
        trigger_time = click_download_button(page, debug_dir, args.log_file)
        save_debug_snapshot(page, debug_dir, "status_download_clicked")
        log(f"下载触发时间：{trigger_time.strftime(TIME_FORMAT)}", args.log_file)

    if args.mode == "export_only":
        log("export_only 完成：已点击下载，未从 Outlook 抓取邮件。", args.log_file)
        return

    with timed_step("从 Outlook 下载最近增长权益邮件附件", args.log_file):
        wait_and_download_outlook_file(args, trigger_time, args.log_file)


def main() -> int:
    args = parse_args()
    if args.mode not in {"dry_run", "download_latest_only"} and args.confirm_export != "确认导出":
        raise RuntimeError("点击后台下载必须传入 --confirm-export 确认导出")
    ensure_playwright()
    args.profile_dir.mkdir(parents=True, exist_ok=True)
    args.download_dir.mkdir(parents=True, exist_ok=True)
    args.log_file.parent.mkdir(parents=True, exist_ok=True)
    args.log_file.write_text("", encoding="utf-8")
    debug_dir = args.log_file.parent / "debug_registration_status_export"
    debug_dir.mkdir(parents=True, exist_ok=True)
    month_labels = target_month_labels(args.month_threshold_day)

    log(f"运行模式：{args.mode}", args.log_file)
    log(f"月份选择优先级：{'、'.join(month_labels)}", args.log_file)
    log(f"导出下载目录：{args.download_dir}", args.log_file)
    if args.mode == "dry_run":
        log("安全提示：当前模式不会点击下载。", args.log_file)
    if args.mode == "download_latest_only":
        download_latest_outlook_file(args, args.log_file)
        return 0

    from playwright.sync_api import sync_playwright

    with sync_playwright() as playwright:
        context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(args.profile_dir),
            headless=args.headless,
            accept_downloads=False,
            viewport={"width": 1500, "height": 900},
            args=["--start-maximized"],
        )
        page = context.pages[0] if context.pages else context.new_page()
        page.set_default_timeout(10000)
        try:
            process_export(page, args, month_labels, debug_dir)
            input("新人价提报情况导出流程结束。确认后按回车关闭浏览器...")
        finally:
            context.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
