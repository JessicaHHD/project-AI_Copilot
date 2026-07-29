from __future__ import annotations

import argparse
import time
from datetime import datetime
from pathlib import Path

from backend_upload_export import (
    click_first_available_anywhere,
    ensure_playwright,
    fast_click_text_anywhere,
    get_files,
    handle_inner_erp_login_if_needed,
    log,
    page_and_frames,
    save_debug_snapshot,
    timed_step,
    upload_file,
    visible_text_exists_anywhere,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="新人价补贴报名：安全演练、上传不提交、或真实提交")
    parser.add_argument(
        "--url",
        default="https://yx.jd.com/user-operate#/customize/newCanSignUpActivity?pathkey=goodspool",
        help="新人价补贴报名页面地址",
    )
    parser.add_argument("--submit-dir", type=Path, required=True, help="提报 PART 文件目录")
    parser.add_argument("--pattern", default="新人价提报_PART*.xlsx", help="提报 PART 文件匹配规则")
    parser.add_argument("--only", help="只处理某一个文件名")
    parser.add_argument("--start-index", default=1, type=int, help="从匹配文件列表的第 N 个文件开始处理；默认 1")
    parser.add_argument("--limit", default=1, type=int, help="从起始文件开始只处理 N 个文件；默认 1；0 表示不限制")
    parser.add_argument(
        "--mode",
        choices=["dry_run", "upload_only", "submit"],
        default="dry_run",
        help="dry_run 不上传不提交；upload_only 上传但不提交；submit 真实提交",
    )
    parser.add_argument("--confirm-submit", default="", help="真实提交必须传入固定确认语：确认提报")
    parser.add_argument("--profile-dir", type=Path, required=True, help="后台浏览器登录缓存目录")
    parser.add_argument("--log-file", type=Path, required=True, help="运行日志保存位置")
    parser.add_argument("--month-threshold-day", type=int, default=25, help="大于等于该日期时优先选择下月计划")
    parser.add_argument("--headless", action="store_true", help="无界面运行；首次登录和调试不建议开启")
    return parser.parse_args()


def cachebust_url(url: str) -> str:
    separator = "&" if "?" in url else "?"
    return f"{url}{separator}time={int(time.time() * 1000)}"


def target_month_labels(threshold_day: int) -> list[str]:
    now = datetime.now()
    target_year = now.year
    target_month = now.month
    fallback_year = now.year
    fallback_month = now.month
    if now.day >= threshold_day:
        target_month += 1
        if target_month > 12:
            target_month = 1
            target_year += 1
    else:
        fallback_month += 1
        if fallback_month > 12:
            fallback_month = 1
            fallback_year += 1
    labels = month_label_variants(target_year, target_month)
    for label in month_label_variants(fallback_year, fallback_month):
        if label not in labels:
            labels.append(label)
    return labels


def month_label_variants(year: int, month: int) -> list[str]:
    return [
        f"{year}年{month}月",
        f"{year}年{month:02d}月",
        f"{month}月",
        f"{month:02d}月",
    ]


def click_action_in_block(page, block_text: str, action_texts: list[str], debug_dir: Path, reason: str) -> None:
    script = """
        ({ blockText, actionTexts }) => {
            const visible = (el) => {
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return style && style.visibility !== 'hidden' && style.display !== 'none' && rect.width > 0 && rect.height > 0;
            };
            const candidates = Array.from(document.querySelectorAll('button, a, span, div'))
                .filter(el => visible(el) && actionTexts.some(text => (el.innerText || el.textContent || '').trim().includes(text)));
            for (const candidate of candidates) {
                let node = candidate;
                for (let depth = 0; node && depth < 9; depth += 1, node = node.parentElement) {
                    const text = node.innerText || node.textContent || '';
                    if (text.includes(blockText)) {
                        candidate.click();
                        return true;
                    }
                }
            }
            return false;
        }
    """
    for target in page_and_frames(page):
        try:
            if target.evaluate(script, {"blockText": block_text, "actionTexts": action_texts}):
                return
        except Exception:
            continue
    save_debug_snapshot(page, debug_dir, reason)
    raise RuntimeError(f"找不到包含“{block_text}”区域内的操作按钮：{'、'.join(action_texts)}。已保存调试文件到：{debug_dir}")


def click_first_order_apply(page, debug_dir: Path) -> None:
    script = """
        () => {
            const visible = (el) => {
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return style && style.visibility !== 'hidden' && style.display !== 'none' && rect.width > 0 && rect.height > 0;
            };
            const textOf = (el) => (el.innerText || el.textContent || '').trim();
            const titleNodes = Array.from(document.querySelectorAll('div, span, p, h1, h2, h3, h4'))
                .filter(el => visible(el) && textOf(el) === '首单新人价');
            for (const title of titleNodes) {
                let card = title;
                for (let depth = 0; card && depth < 8; depth += 1, card = card.parentElement) {
                    const cardText = textOf(card);
                    if (!cardText.includes('首单新人价') || !cardText.includes('前往报名')) continue;
                    if (cardText.includes('重复价') || cardText.includes('电商新人价') || cardText.includes('品类新人价')) continue;
                    const actions = Array.from(card.querySelectorAll('button, a, span, div'))
                        .filter(el => visible(el) && textOf(el).includes('前往报名'));
                    const preferred = actions.find(el => {
                        const tag = el.tagName.toLowerCase();
                        const cls = el.className || '';
                        return tag === 'button' || tag === 'a' || String(cls).includes('button') || String(cls).includes('btn');
                    }) || actions[0];
                    if (preferred) {
                        preferred.click();
                        return { ok: true, cardText: cardText.slice(0, 200) };
                    }
                }
            }
            return { ok: false, titles: titleNodes.map(el => textOf(el)).slice(0, 10) };
        }
    """
    last_result = None
    for target in page_and_frames(page):
        try:
            result = target.evaluate(script)
            last_result = result
            if result and result.get("ok"):
                return
        except Exception:
            continue
    save_debug_snapshot(page, debug_dir, "first_order_apply_not_found")
    raise RuntimeError(f"找不到“首单新人价”卡片内的“前往报名”按钮：{last_result}。已保存调试文件到：{debug_dir}")


def wait_for_any_text(page, texts: list[str], timeout_seconds: int = 10) -> bool:
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        if visible_text_exists_anywhere(page, texts, timeout_ms=300):
            return True
        time.sleep(0.3)
    return False


def plan_row_ready(page, month_labels: list[str]) -> bool:
    script = """
        ({ monthLabels }) => {
            const visible = (el) => {
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return style && style.visibility !== 'hidden' && style.display !== 'none' && rect.width > 0 && rect.height > 0;
            };
            const norm = (text) => String(text || '').replace(/\\s+/g, '');
            const bodyText = norm(document.body.innerText || document.body.textContent || '');
            const hasMonth = monthLabels.some(label => bodyText.includes(norm(label)));
            const hasApply = Array.from(document.querySelectorAll('button, a, span, div'))
                .some(el => visible(el) && (el.innerText || el.textContent || '').includes('去报名'));
            return hasMonth && hasApply;
        }
    """
    for target in page_and_frames(page):
        try:
            if target.evaluate(script, {"monthLabels": month_labels}):
                return True
        except Exception:
            continue
    return False


def wait_plan_row_ready(page, month_labels: list[str], debug_dir: Path, log_file: Path, timeout_seconds: int = 20) -> None:
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        if plan_row_ready(page, month_labels):
            log(f"已检测到目标月份计划行和“去报名”：{'、'.join(month_labels)}", log_file)
            return
        time.sleep(0.5)
    for text in collect_plan_debug_texts(page):
        log(f"等待计划行超时，可见计划文本：{text}", log_file)
    save_debug_snapshot(page, debug_dir, "plan_row_not_ready")
    raise RuntimeError(f"等待目标月份计划行超时：{'、'.join(month_labels)}。已保存调试文件到：{debug_dir}")



def registration_home_ready(page) -> bool:
    return visible_text_exists_anywhere(page, ["首单新人价"], timeout_ms=500) and visible_text_exists_anywhere(page, ["前往报名"], timeout_ms=500)


def wait_registration_home_ready(page, timeout_seconds: int = 10) -> bool:
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        if registration_home_ready(page):
            return True
        time.sleep(0.3)
    return False


def click_registration_home_tab(page, log_file: Path) -> bool:
    script = r"""
        () => {
            const visible = (el) => {
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return style && style.visibility !== 'hidden' && style.display !== 'none' && rect.width > 0 && rect.height > 0;
            };
            const norm = (text) => String(text || '').replace(/\s+/g, '');
            const textOf = (el) => el.innerText || el.textContent || '';
            const dispatchClick = (el) => {
                el.scrollIntoView({ block: 'center', inline: 'center' });
                const rect = el.getBoundingClientRect();
                const x = rect.left + rect.width / 2;
                const y = rect.top + rect.height / 2;
                for (const type of ['mouseover', 'mousedown', 'mouseup', 'click']) {
                    el.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window, clientX: x, clientY: y }));
                }
                return { tag: el.tagName, text: textOf(el).replace(/\s+/g, ' ').trim(), x: Math.round(x), y: Math.round(y), path: el.getAttribute('path') || '' };
            };
            const exactMenu = Array.from(document.querySelectorAll('[path="/newCanSignUpActivity"], [name="新人补贴报名"], li[role="menuitem"]'))
                .filter(el => visible(el) && (el.getAttribute('path') === '/newCanSignUpActivity' || el.getAttribute('name') === '新人补贴报名' || norm(textOf(el)) === '新人补贴报名'))
                .sort((a, b) => a.getBoundingClientRect().left - b.getBoundingClientRect().left || norm(textOf(a)).length - norm(textOf(b)).length)[0];
            if (exactMenu) return dispatchClick(exactMenu);

            const exactText = Array.from(document.querySelectorAll('button, a, [role=button], li, div, span'))
                .filter(el => visible(el) && norm(textOf(el)) === '新人补贴报名')
                .sort((a, b) => a.getBoundingClientRect().left - b.getBoundingClientRect().left || norm(textOf(a)).length - norm(textOf(b)).length)[0];
            if (!exactText) return null;
            const clickable = exactText.closest('[path="/newCanSignUpActivity"], [name="新人补贴报名"], button, a, [role=button], li') || exactText;
            return dispatchClick(clickable);
        }
    """
    for target in page_and_frames(page):
        try:
            result = target.evaluate(script)
            if result:
                log(f"已点击左侧“新人补贴报名”tab：文本={result.get('text')}，标签={result.get('tag')}，坐标=({result.get('x')},{result.get('y')})", log_file)
                return True
        except Exception:
            continue
    return False


def force_registration_home_iframe(page, log_file: Path) -> bool:
    script = r"""
        () => {
            const iframe = document.querySelector('#currntIframe, iframe[src*="Activity"]');
            if (!iframe) return false;
            iframe.src = 'https://goodspool-pro.local-pf.jd.com/newCanSignUpActivity?nomenu=no&priceType=3&time=' + Date.now();
            return true;
        }
    """
    try:
        if page.evaluate(script):
            log("点击 tab 后未出现卡片，已强制刷新新人补贴报名 iframe。", log_file)
            return True
    except Exception:
        pass
    return False


def enter_registration_home(page, debug_dir: Path, log_file: Path, home_url: str) -> None:
    if wait_registration_home_ready(page, timeout_seconds=2):
        return
    if click_registration_home_tab(page, log_file) and wait_registration_home_ready(page, timeout_seconds=8):
        return
    log("点击左侧“新人补贴报名”后未检测到首单新人价卡片，改用带时间戳重新打开报名首页。", log_file)
    page.goto(cachebust_url(home_url), wait_until="domcontentloaded")
    if wait_registration_home_ready(page, timeout_seconds=8):
        return
    if force_registration_home_iframe(page, log_file) and wait_registration_home_ready(page, timeout_seconds=10):
        return
    save_debug_snapshot(page, debug_dir, "registration_home_not_ready")
    raise RuntimeError(f"未检测到新人补贴报名首页的“首单新人价/前往报名”卡片。已保存调试文件到：{debug_dir}")


def enter_first_order_registration(page, debug_dir: Path) -> None:
    before_url = page.url
    click_first_order_apply(page, debug_dir)
    try:
        page.wait_for_load_state("domcontentloaded", timeout=5000)
    except Exception:
        pass
    deadline = time.time() + 10
    while time.time() < deadline:
        if page.url != before_url and visible_text_exists_anywhere(page, ["待报名", "已报名", "补贴计划名称"], timeout_ms=500):
            break
        if visible_text_exists_anywhere(page, ["待报名", "已报名", "补贴计划名称"], timeout_ms=500):
            break
        time.sleep(0.3)
    if not visible_text_exists_anywhere(page, ["待报名", "已报名", "补贴计划名称"], timeout_ms=800):
        save_debug_snapshot(page, debug_dir, "plan_list_not_ready")
        raise RuntimeError(f"进入首单新人价报名后未检测到计划列表。已保存调试文件到：{debug_dir}")
    fast_click_text_anywhere(page, ["待报名"], timeout_ms=1000, clickable_only=False)


def reset_scroll_anywhere(page) -> None:
    script = r"""
        () => {
            try { window.scrollTo(0, 0); } catch (e) {}
            for (const el of Array.from(document.querySelectorAll('*'))) {
                try {
                    if ((el.scrollTop || el.scrollLeft) && el.scrollHeight > el.clientHeight) {
                        el.scrollTop = 0;
                        el.scrollLeft = 0;
                    }
                } catch (e) {}
            }
        }
    """
    for target in page_and_frames(page):
        try:
            target.evaluate(script)
        except Exception:
            continue


def wait_after_plan_click(page, before_url: str, timeout_seconds: int = 12) -> bool:
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        if visible_text_exists_anywhere(page, ["批量上传10个以上sku", "批量上传 10个以上sku", "单个设置"], timeout_ms=500):
            return True
        if page.url != before_url and visible_text_exists_anywhere(page, ["批量上传", "单个设置", "选择文件"], timeout_ms=500):
            return True
        time.sleep(0.3)
    return False




def find_visible_plan_apply_button(page, month_labels: list[str]) -> dict | None:
    script = r"""
        ({ monthLabels }) => {
            const visible = (el) => {
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return style && style.visibility !== 'hidden' && style.display !== 'none' && rect.width > 0 && rect.height > 0 && rect.bottom > 0 && rect.right > 0 && rect.top < window.innerHeight && rect.left < window.innerWidth;
            };
            const textOf = (el) => (el.innerText || el.textContent || '').trim();
            const norm = (text) => String(text || '').replace(/\s+/g, '');
            const rectInfo = (el) => {
                const rect = el.getBoundingClientRect();
                return { left: rect.left, top: rect.top, right: rect.right, bottom: rect.bottom, width: rect.width, height: rect.height, x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 };
            };
            const clickableOf = (el) => el.closest('button, a, [role=button], .ant-btn, .el-button, [class*=btn], [class*=button]') || el;
            const rowSelectors = ['tr', '[role=row]', '.ant-table-row', '.el-table__row', '[class*=table] [class*=row]', '[class*=Table] [class*=Row]'];
            const actionSelector = 'button, a, [role=button], .ant-btn, .el-button, [class*=btn], [class*=button], span, div';
            const candidates = [];
            const seen = new Set();
            const addCandidate = (action, row, monthLabel, method) => {
                const clickable = clickableOf(action);
                if (!clickable || !visible(clickable)) return;
                const rect = rectInfo(clickable);
                if (rect.x < 0 || rect.y < 0 || rect.x > window.innerWidth || rect.y > window.innerHeight) return;
                if (seen.has(clickable)) return;
                seen.add(clickable);
                candidates.push({
                    monthLabel,
                    method,
                    tag: clickable.tagName,
                    className: String(clickable.className || ''),
                    text: textOf(clickable) || textOf(action),
                    rowText: textOf(row).replace(/\s+/g, ' ').slice(0, 180),
                    x: Math.round(rect.x),
                    y: Math.round(rect.y),
                    left: Math.round(rect.left),
                    top: Math.round(rect.top),
                    right: Math.round(rect.right),
                    bottom: Math.round(rect.bottom),
                });
            };
            for (const monthLabel of monthLabels) {
                const monthTarget = norm(monthLabel);
                for (const rowSelector of rowSelectors) {
                    for (const row of Array.from(document.querySelectorAll(rowSelector))) {
                        if (!visible(row)) continue;
                        const rowText = norm(textOf(row));
                        if (!rowText.includes(monthTarget) || !rowText.includes('去报名')) continue;
                        for (const action of Array.from(row.querySelectorAll(actionSelector))) {
                            if (visible(action) && norm(textOf(action)).includes('去报名')) {
                                addCandidate(action, row, monthLabel, 'visible_row');
                            }
                        }
                    }
                }
                if (candidates.length) break;
            }
            if (!candidates.length) {
                for (const monthLabel of monthLabels) {
                    const monthTarget = norm(monthLabel);
                    const monthItems = Array.from(document.querySelectorAll('body *'))
                        .filter(el => visible(el) && norm(textOf(el)).includes(monthTarget))
                        .map(el => ({ el, rect: rectInfo(el), text: textOf(el) }))
                        .filter(item => item.text.length <= 120)
                        .sort((a, b) => (a.rect.width * a.rect.height) - (b.rect.width * b.rect.height));
                    const actions = Array.from(document.querySelectorAll(actionSelector))
                        .filter(el => visible(el) && norm(textOf(el)).includes('去报名'))
                        .map(el => ({ el, rect: rectInfo(el), text: textOf(el) }));
                    for (const monthItem of monthItems) {
                        const sameLine = actions
                            .filter(action => Math.abs(action.rect.y - monthItem.rect.y) <= 32 && action.rect.left > monthItem.rect.left)
                            .sort((a, b) => b.rect.left - a.rect.left)[0];
                        if (sameLine) addCandidate(sameLine.el, monthItem.el, monthLabel, 'visible_same_line');
                    }
                    if (candidates.length) break;
                }
            }
            candidates.sort((a, b) => b.right - a.right || a.top - b.top);
            return candidates[0] || null;
        }
    """
    for target in page_and_frames(page):
        try:
            result = target.evaluate(script, {"monthLabels": month_labels})
            if result:
                return result
        except Exception:
            continue
    return None



def click_visible_plan_apply_button(page, month_labels: list[str], log_file: Path) -> str:
    script = r"""
        ({ monthLabels }) => {
            const visible = (el) => {
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return style && style.visibility !== 'hidden' && style.display !== 'none' && rect.width > 0 && rect.height > 0 && rect.bottom > 0 && rect.right > 0 && rect.top < window.innerHeight && rect.left < window.innerWidth;
            };
            const textOf = (el) => (el.innerText || el.textContent || '').trim();
            const norm = (text) => String(text || '').replace(/\s+/g, '');
            const rectInfo = (el) => {
                const rect = el.getBoundingClientRect();
                return { left: rect.left, top: rect.top, right: rect.right, bottom: rect.bottom, width: rect.width, height: rect.height, x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 };
            };
            const clickableOf = (el) => el.closest('button, a, [role=button], .ant-btn, .el-button, [class*=btn], [class*=button]') || el;
            const dispatchClick = (el) => {
                el.scrollIntoView({ block: 'center', inline: 'center' });
                const rect = el.getBoundingClientRect();
                const x = rect.left + rect.width / 2;
                const y = rect.top + rect.height / 2;
                const realTarget = document.elementFromPoint(x, y) || el;
                const clickable = clickableOf(realTarget) || el;
                clickable.focus && clickable.focus();
                for (const type of ['mouseover', 'mousemove', 'mousedown', 'mouseup', 'click']) {
                    clickable.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window, clientX: x, clientY: y, button: 0 }));
                }
                clickable.click && clickable.click();
                return { target: clickable, rect, x, y };
            };
            const rowSelectors = ['tr', '[role=row]', '.ant-table-row', '.el-table__row', '[class*=table] [class*=row]', '[class*=Table] [class*=Row]'];
            const actionSelector = 'button, a, [role=button], .ant-btn, .el-button, [class*=btn], [class*=button], span, div';
            const candidates = [];
            const seen = new Set();
            const addCandidate = (action, row, monthLabel, method) => {
                const clickable = clickableOf(action);
                if (!clickable || !visible(clickable)) return;
                const rect = rectInfo(clickable);
                if (rect.x < 0 || rect.y < 0 || rect.x > window.innerWidth || rect.y > window.innerHeight) return;
                if (seen.has(clickable)) return;
                seen.add(clickable);
                candidates.push({ action, row, monthLabel, method, rect });
            };
            for (const monthLabel of monthLabels) {
                const monthTarget = norm(monthLabel);
                for (const rowSelector of rowSelectors) {
                    for (const row of Array.from(document.querySelectorAll(rowSelector))) {
                        if (!visible(row)) continue;
                        const rowText = norm(textOf(row));
                        if (!rowText.includes(monthTarget) || !rowText.includes('去报名')) continue;
                        for (const action of Array.from(row.querySelectorAll(actionSelector))) {
                            if (visible(action) && norm(textOf(action)).includes('去报名')) addCandidate(action, row, monthLabel, 'visible_row');
                        }
                    }
                }
                if (candidates.length) break;
            }
            if (!candidates.length) {
                for (const monthLabel of monthLabels) {
                    const monthTarget = norm(monthLabel);
                    const monthItems = Array.from(document.querySelectorAll('body *'))
                        .filter(el => visible(el) && norm(textOf(el)).includes(monthTarget))
                        .map(el => ({ el, rect: rectInfo(el), text: textOf(el) }))
                        .filter(item => item.text.length <= 120)
                        .sort((a, b) => (a.rect.width * a.rect.height) - (b.rect.width * b.rect.height));
                    const actions = Array.from(document.querySelectorAll(actionSelector))
                        .filter(el => visible(el) && norm(textOf(el)).includes('去报名'))
                        .map(el => ({ el, rect: rectInfo(el), text: textOf(el) }));
                    for (const monthItem of monthItems) {
                        const sameLine = actions
                            .filter(action => Math.abs(action.rect.y - monthItem.rect.y) <= 32 && action.rect.left > monthItem.rect.left)
                            .sort((a, b) => b.rect.left - a.rect.left)[0];
                        if (sameLine) addCandidate(sameLine.el, monthItem.el, monthLabel, 'visible_same_line');
                    }
                    if (candidates.length) break;
                }
            }
            candidates.sort((a, b) => b.rect.right - a.rect.right || a.rect.top - b.rect.top);
            const candidate = candidates[0];
            if (!candidate) return null;
            const clickResult = dispatchClick(clickableOf(candidate.action));
            return {
                monthLabel: candidate.monthLabel,
                method: candidate.method,
                text: textOf(clickResult.target) || textOf(candidate.action),
                tag: clickResult.target.tagName,
                rowText: textOf(candidate.row).replace(/\s+/g, ' ').slice(0, 180),
                x: Math.round(clickResult.x),
                y: Math.round(clickResult.y),
            };
        }
    """
    for attempt in range(1, 4):
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
                f"按可见元素点击补贴计划行“去报名”：月份={selected}，方式={result.get('method')}，"
                f"文本={result.get('text')}，坐标=({result.get('x')},{result.get('y')})",
                log_file,
            )
            if wait_after_plan_click(page, before_url, timeout_seconds=4):
                log(f"已进入补贴计划报名页：{selected}", log_file)
                return selected
        log(f"可见元素点击后未进入报名页，继续尝试第 {attempt + 1} 次。", log_file)
        time.sleep(0.5)
    return ""


def click_plan_with_locator(page, month_labels: list[str], log_file: Path, attempted_clicks: list[str]) -> str:
    row_selectors = [
        "tr",
        "[role=row]",
        ".ant-table-row",
        ".el-table__row",
        "[class*=table] [class*=row]",
        "[class*=Table] [class*=Row]",
    ]
    action_selector = (
        "button:has-text('去报名'), a:has-text('去报名'), [role=button]:has-text('去报名'), "
        ".ant-btn:has-text('去报名'), .el-button:has-text('去报名')"
    )
    selected_by_visible_coordinate = click_visible_plan_apply_button(page, month_labels, log_file)
    if selected_by_visible_coordinate:
        return selected_by_visible_coordinate

    for target in page_and_frames(page):
        for month_label in month_labels:
            for row_selector in row_selectors:
                try:
                    rows = target.locator(row_selector).filter(has_text=month_label).filter(has_text="去报名")
                    row_count = min(rows.count(), 5)
                except Exception:
                    continue
                for index in range(row_count):
                    row = rows.nth(index)
                    try:
                        row_text = row.inner_text(timeout=1000).strip().replace("\n", " ")[:180]
                    except Exception:
                        row_text = ""
                    try:
                        row.scroll_into_view_if_needed(timeout=2000)
                    except Exception:
                        pass
                    try:
                        actions = row.locator(action_selector)
                        action = actions.last if actions.count() else row.get_by_text("去报名", exact=False).last
                    except Exception:
                        action = row.get_by_text("去报名", exact=False).last
                    log(f"尝试点击补贴计划行“去报名”：月份={month_label}，行={row_text}", log_file)
                    before_url = page.url
                    try:
                        action.scroll_into_view_if_needed(timeout=2000)
                    except Exception:
                        pass
                    try:
                        action.click(timeout=5000)
                    except Exception as exc:
                        log(f"常规点击“去报名”失败，改用强制点击：{exc}", log_file)
                        try:
                            action.click(timeout=3000, force=True)
                        except Exception as force_exc:
                            log(f"强制点击“去报名”失败，继续尝试其他候选：{force_exc}", log_file)
                            continue
                    if wait_after_plan_click(page, before_url, timeout_seconds=12):
                        log(f"已进入补贴计划报名页：{month_label}", log_file)
                        return month_label
                    attempted_clicks.append(month_label)
                    log("点击“去报名”后未检测到报名表单页，继续尝试其他候选。", log_file)
    return ""


def select_plan(page, month_labels: list[str], debug_dir: Path, log_file: Path) -> str:
    reset_scroll_anywhere(page)
    try:
        page.wait_for_load_state("networkidle", timeout=3000)
    except Exception:
        pass
    time.sleep(0.3)

    locator_clicked_without_navigation: list[str] = []
    selected_by_locator = click_plan_with_locator(page, month_labels, log_file, locator_clicked_without_navigation)
    if selected_by_locator:
        return selected_by_locator

    clicked_without_navigation: list[str] = []
    script = """
        ({ monthLabels }) => {
            const visible = (el) => {
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return style && style.visibility !== 'hidden' && style.display !== 'none' && rect.width > 0 && rect.height > 0;
            };
            const textOf = (el) => (el.innerText || el.textContent || '').trim();
            const norm = (text) => String(text || '').replace(/\\s+/g, '');
            const clickTarget = (row) => {
                const actions = Array.from(row.querySelectorAll('button, a, span, div'))
                    .filter(el => visible(el) && textOf(el).includes('去报名'));
                const preferred = actions.find(el => {
                    const tag = el.tagName.toLowerCase();
                    const cls = String(el.className || '');
                    return tag === 'button' || tag === 'a' || cls.includes('button') || cls.includes('btn');
                }) || actions[0];
                if (preferred) {
                    preferred.scrollIntoView({ block: 'center', inline: 'center' });
                    preferred.click();
                    return true;
                }
                return false;
            };
            const clickActionOnSameLine = (monthLabel) => {
                const target = norm(monthLabel);
                const monthElements = Array.from(document.querySelectorAll('body *'))
                    .filter(el => visible(el) && norm(textOf(el)).includes(target))
                    .map(el => ({ el, rect: el.getBoundingClientRect(), text: textOf(el) }))
                    .filter(item => item.text.length <= 120)
                    .sort((a, b) => (a.rect.width * a.rect.height) - (b.rect.width * b.rect.height));
                const actions = Array.from(document.querySelectorAll('button, a, span, div'))
                    .filter(el => visible(el) && textOf(el).includes('去报名'))
                    .map(el => ({ el, rect: el.getBoundingClientRect(), text: textOf(el) }));
                for (const monthItem of monthElements) {
                    const y = monthItem.rect.top + monthItem.rect.height / 2;
                    const sameLine = actions
                        .filter(action => Math.abs((action.rect.top + action.rect.height / 2) - y) <= 28)
                        .sort((a, b) => b.rect.left - a.rect.left)[0];
                    if (sameLine) {
                        sameLine.el.scrollIntoView({ block: 'center', inline: 'center' });
                        sameLine.el.click();
                        return true;
                    }
                }
                return false;
            };
            const rows = Array.from(document.querySelectorAll('tr, .ant-table-row, .el-table__row, [class*=table] [class*=row]'))
                .filter(row => visible(row) && textOf(row).includes('去报名'));
            for (const monthLabel of monthLabels) {
                for (const row of rows) {
                    const rowText = textOf(row);
                    if (rowText.includes(monthLabel) && clickTarget(row)) {
                        return monthLabel;
                    }
                }
            }
            for (const monthLabel of monthLabels) {
                if (clickActionOnSameLine(monthLabel)) {
                    return monthLabel;
                }
            }
            const allActions = Array.from(document.querySelectorAll('button, a, span, div'))
                .filter(el => visible(el) && textOf(el).includes('去报名'));
            for (const monthLabel of monthLabels) {
                for (const action of allActions) {
                    let node = action;
                    for (let depth = 0; node && depth < 8; depth += 1, node = node.parentElement) {
                        if (textOf(node).includes(monthLabel)) {
                            action.scrollIntoView({ block: 'center', inline: 'center' });
                            action.click();
                            return monthLabel;
                        }
                    }
                }
            }
            return '';
        }
    """
    for target in page_and_frames(page):
        try:
            before_url = page.url
            selected = target.evaluate(script, {"monthLabels": month_labels})
            if selected:
                if wait_after_plan_click(page, before_url, timeout_seconds=12):
                    log(f"已进入补贴计划报名页：{selected}", log_file)
                    return str(selected)
                clicked_without_navigation.append(str(selected))
                log(f"JS 兜底点击“去报名”后未进入报名页：{selected}", log_file)
        except Exception:
            continue
    for text in collect_plan_debug_texts(page):
        log(f"可见计划文本：{text}", log_file)
    all_click_attempts = locator_clicked_without_navigation + clicked_without_navigation
    if all_click_attempts:
        save_debug_snapshot(page, debug_dir, "plan_apply_click_no_navigation")
        raise RuntimeError(
            "已找到目标月份计划行的“去报名”，但点击后未进入报名页。"
            f"已尝试月份：{'、'.join(all_click_attempts)}。已保存调试文件到：{debug_dir}"
        )
    save_debug_snapshot(page, debug_dir, "plan_not_found")
    raise RuntimeError(f"未找到目标月份补贴计划的“去报名”：{'、'.join(month_labels)}。已保存调试文件到：{debug_dir}")


def collect_plan_debug_texts(page) -> list[str]:
    script = """
        () => {
            const visible = (el) => {
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return style && style.visibility !== 'hidden' && style.display !== 'none' && rect.width > 0 && rect.height > 0;
            };
            const seen = new Set();
            const result = [];
            for (const el of Array.from(document.querySelectorAll('tr, .ant-table-row, .el-table__row, [class*=table] [class*=row], div'))) {
                if (!visible(el)) continue;
                const text = (el.innerText || el.textContent || '').trim().replace(/\\s+/g, ' ');
                if (!text || text.length > 300) continue;
                if ((text.includes('去报名') || text.includes('补贴计划') || text.includes('2026年')) && !seen.has(text)) {
                    seen.add(text);
                    result.push(text);
                }
                if (result.length >= 20) break;
            }
            return result;
        }
    """
    texts: list[str] = []
    for target in page_and_frames(page):
        try:
            for text in target.evaluate(script):
                if text not in texts:
                    texts.append(str(text))
        except Exception:
            continue
    return texts[:20]


def enter_batch_upload_area(page, debug_dir: Path) -> None:
    if not wait_for_any_text(page, ["批量上传10个以上sku", "单个设置"], timeout_seconds=10):
        save_debug_snapshot(page, debug_dir, "signup_form_not_ready")
        raise RuntimeError(f"报名表单页面未加载完成。已保存调试文件到：{debug_dir}")
    clicked = fast_click_text_anywhere(page, ["批量上传10个以上sku", "批量上传 10个以上sku"], timeout_ms=1000, clickable_only=False)
    if not clicked:
        clicked = click_first_available_anywhere(
            page,
            [("text", "批量上传10个以上sku"), ("text", "批量上传 10个以上sku")],
            timeout=1000,
        )
    if not clicked:
        save_debug_snapshot(page, debug_dir, "batch_signup_tab_not_found")
        raise RuntimeError(f"找不到“批量上传10个以上sku”入口。已保存调试文件到：{debug_dir}")
    if not wait_for_any_text(page, ["选择文件", "下载模板", "批量上传"], timeout_seconds=6):
        save_debug_snapshot(page, debug_dir, "batch_signup_area_not_ready")
        raise RuntimeError(f"批量上传区域未出现。已保存调试文件到：{debug_dir}")


def click_normalized_button_anywhere(page, normalized_text: str, log_file: Path, reason: str) -> bool:
    script = r"""
        ({ normalizedText }) => {
            const visible = (el) => {
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return style && style.visibility !== 'hidden' && style.display !== 'none' && rect.width > 0 && rect.height > 0;
            };
            const norm = (text) => String(text || '').replace(/\s+/g, '');
            const textOf = (el) => el.innerText || el.textContent || el.value || '';
            const nodes = Array.from(document.querySelectorAll('button, [role=button], a, .ant-btn, .el-button, input[type=button], input[type=submit]'))
                .filter(el => visible(el) && !el.disabled && !el.getAttribute('disabled'))
                .map(el => ({ el, rect: el.getBoundingClientRect(), text: textOf(el), normalized: norm(textOf(el)) }))
                .filter(item => item.normalized === normalizedText || item.normalized.includes(normalizedText))
                .sort((a, b) => (b.rect.bottom - a.rect.bottom) || (b.rect.right - a.rect.right));
            const target = nodes[0];
            if (!target) return null;
            target.el.scrollIntoView({ block: 'center', inline: 'center' });
            target.el.click();
            return {
                tag: target.el.tagName,
                text: String(target.text || '').replace(/\s+/g, ' ').trim(),
                x: Math.round(target.rect.left + target.rect.width / 2),
                y: Math.round(target.rect.top + target.rect.height / 2),
            };
        }
    """
    for target in page_and_frames(page):
        try:
            result = target.evaluate(script, {"normalizedText": normalized_text})
            if result:
                log(f"已点击{reason}按钮：文本={result.get('text')}，标签={result.get('tag')}，坐标=({result.get('x')},{result.get('y')})", log_file)
                return True
        except Exception:
            continue
    return False


def wait_submit_feedback(page, debug_dir: Path, log_file: Path) -> None:
    deadline = time.time() + 12
    success_texts = ["提交成功", "报名成功", "上传成功", "导入成功", "已提报", "操作成功"]
    error_texts = ["提交失败", "报名失败", "上传失败", "导入失败", "错误", "异常"]
    while time.time() < deadline:
        if visible_text_exists_anywhere(page, success_texts, timeout_ms=500):
            log("已检测到提交后的成功/已提报提示。", log_file)
            return
        if visible_text_exists_anywhere(page, error_texts, timeout_ms=500):
            save_debug_snapshot(page, debug_dir, "submit_after_error")
            raise RuntimeError(f"点击提交后页面出现失败/错误提示。已保存调试文件到：{debug_dir}")
        time.sleep(0.5)
    log("已点击提交，但未在 12 秒内检测到明确成功提示；已保存提交后截图供人工确认。", log_file)
    save_debug_snapshot(page, debug_dir, "submit_clicked_waiting_feedback")


def click_submit(page, debug_dir: Path, log_file: Path) -> None:
    clicked = click_normalized_button_anywhere(page, "提交", log_file, "提交")
    if not clicked:
        clicked = fast_click_text_anywhere(page, ["提交", "提 交"], timeout_ms=1000, clickable_only=True)
    if not clicked:
        clicked = click_first_available_anywhere(
            page,
            [("role_button", "提交"), ("role_button", "提 交"), ("css", "button:has-text('提 交')"), ("css", "button:has-text('提交')")],
            timeout=1000,
        )
    if not clicked:
        save_debug_snapshot(page, debug_dir, "submit_button_not_found")
        raise RuntimeError(f"找不到“提交”按钮。已保存调试文件到：{debug_dir}")
    time.sleep(0.8)
    for label in ["确定", "确认"]:
        if click_normalized_button_anywhere(page, label, log_file, label):
            break
    wait_submit_feedback(page, debug_dir, log_file)


def upload_registration_file(page, file_path: Path, debug_dir: Path) -> None:
    try:
        upload_file(page, file_path, debug_dir)
        return
    except Exception:
        pass
    selectors = [
        ("role_button", "选择文件"),
        ("text", "选择文件"),
        ("text", "上传文件"),
        ("css", "button:has-text('选择文件')"),
        ("css", ".ant-upload"),
        ("css", ".el-upload"),
        ("css", "[class*='upload']"),
    ]
    try:
        with page.expect_file_chooser(timeout=5000) as chooser_info:
            clicked = click_first_available_anywhere(page, selectors, timeout=2000)
            if not clicked:
                save_debug_snapshot(page, debug_dir, "registration_upload_button_not_found")
                raise RuntimeError(f"找不到选择文件按钮。已保存调试文件到：{debug_dir}")
        chooser_info.value.set_files(str(file_path))
    except Exception as exc:
        save_debug_snapshot(page, debug_dir, "registration_upload_failed")
        raise RuntimeError(f"提报文件上传失败：{exc}。已保存调试文件到：{debug_dir}") from exc


def process_one_file(page, file_path: Path, args: argparse.Namespace, month_labels: list[str], debug_dir: Path) -> None:
    with timed_step("打开新人价补贴报名页面", args.log_file):
        page.goto(cachebust_url(args.url), wait_until="domcontentloaded")
    with timed_step("处理内网ERP登录", args.log_file):
        handle_inner_erp_login_if_needed(page, args.log_file, max_wait_seconds=60)
    with timed_step("进入新人补贴报名首页", args.log_file):
        enter_registration_home(page, debug_dir, args.log_file, args.url)
    with timed_step("进入首单新人价计划列表", args.log_file):
        enter_first_order_registration(page, debug_dir)
    with timed_step("选择补贴计划并去报名", args.log_file):
        wait_plan_row_ready(page, month_labels, debug_dir, args.log_file, timeout_seconds=20)
        selected_month = select_plan(page, month_labels, debug_dir, args.log_file)
    with timed_step("进入批量上传区域", args.log_file):
        enter_batch_upload_area(page, debug_dir)
        save_debug_snapshot(page, debug_dir, f"ready_to_upload_{file_path.stem}")

    if args.mode == "dry_run":
        log(f"dry_run 演练完成：已到达上传页，未选择文件、未点击提交。目标月份：{selected_month}，文件：{file_path.name}", args.log_file)
        return

    with timed_step(f"选择文件 {file_path.name}", args.log_file):
        upload_registration_file(page, file_path, debug_dir)
        save_debug_snapshot(page, debug_dir, f"file_selected_{file_path.stem}")

    if args.mode == "upload_only":
        log(f"upload_only 完成：已选择文件但未点击提交。目标月份：{selected_month}，文件：{file_path.name}", args.log_file)
        return

    with timed_step("点击提交", args.log_file):
        click_submit(page, debug_dir, args.log_file)
        save_debug_snapshot(page, debug_dir, f"submitted_{file_path.stem}")
    log(f"submit 完成：已点击提交。目标月份：{selected_month}，文件：{file_path.name}", args.log_file)


def main() -> int:
    args = parse_args()
    if args.mode == "submit" and args.confirm_submit != "确认提报":
        raise RuntimeError("真实提交模式必须传入 --confirm-submit 确认提报")
    ensure_playwright()
    all_files = get_files(args.submit_dir, args.pattern, args.only, 0)
    if args.only:
        files = all_files
    else:
        start_index = max(args.start_index, 1)
        files = all_files[start_index - 1 :]
        if args.limit and args.limit > 0:
            files = files[: args.limit]
        if not files:
            raise FileNotFoundError(f"从第 {start_index} 个文件开始没有可处理文件；匹配文件总数：{len(all_files)}")
    args.profile_dir.mkdir(parents=True, exist_ok=True)
    args.log_file.parent.mkdir(parents=True, exist_ok=True)
    args.log_file.write_text("", encoding="utf-8")
    debug_dir = args.log_file.parent / "debug_registration_submit"
    debug_dir.mkdir(parents=True, exist_ok=True)
    month_labels = target_month_labels(args.month_threshold_day)

    log(f"运行模式：{args.mode}", args.log_file)
    log(f"起始文件序号：{args.start_index if not args.only else '指定单文件'}", args.log_file)
    log(f"文件数量：{len(files)}；文件：{', '.join(path.name for path in files)}", args.log_file)
    log(f"月份选择优先级：{'、'.join(month_labels)}", args.log_file)
    if args.mode != "submit":
        log("安全提示：当前模式不会点击提交。", args.log_file)

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
            for index, file_path in enumerate(files, start=1):
                log(f"[{index}/{len(files)}] 准备处理：{file_path.name}", args.log_file)
                process_one_file(page, file_path, args, month_labels, debug_dir)
        finally:
            input("新人价提报安全测试流程结束。确认后按回车关闭浏览器...")
            context.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
