from __future__ import annotations

import argparse
import time
from contextlib import contextmanager
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
    parser.add_argument("--file-list", type=Path, help="只处理清单中的文件；文件内容为一行一个 Excel 完整路径")
    parser.add_argument("--limit", default=0, type=int, help="只处理前 N 个匹配文件；0 表示不限制")
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
    parser.add_argument("--wait-after-export", default=0, type=int, help="点击批量导出后等待秒数；默认不等待邮件")
    parser.add_argument("--pause-after-each", action="store_true", help="每个文件批量导出后暂停，方便人工检查页面")
    parser.add_argument("--headless", action="store_true", help="无界面运行；首次登录和调试不建议开启")
    return parser.parse_args()


def ensure_playwright() -> None:
    try:
        import playwright  # noqa: F401
    except Exception as exc:
        raise RuntimeError("缺少 playwright。请先运行：python -m pip install playwright && python -m playwright install chromium") from exc


def get_files(sku_dir: Path, pattern: str, only: str | None, limit: int = 0, file_list: Path | None = None) -> list[Path]:
    if file_list:
        files = [Path(line.strip().strip('"')) for line in file_list.read_text(encoding="utf-8-sig").splitlines() if line.strip()]
    else:
        files = [sku_dir / only] if only else sorted(sku_dir.glob(pattern))
    files = [path for path in files if path.is_file()]
    if limit and limit > 0 and not only:
        files = files[:limit]
    if not files:
        target = file_list if file_list else sku_dir / pattern
        raise FileNotFoundError(f"未找到待上传文件：{target}")
    return files


def log(message: str, log_file: Path) -> None:
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {message}"
    print(line)
    log_file.parent.mkdir(parents=True, exist_ok=True)
    with log_file.open("a", encoding="utf-8") as file:
        file.write(line + "\n")


@contextmanager
def timed_step(name: str, log_file: Path):
    start = time.perf_counter()
    log(f"开始：{name}", log_file)
    try:
        yield
    finally:
        elapsed = time.perf_counter() - start
        log(f"结束：{name}，耗时 {elapsed:.2f} 秒", log_file)


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


def visible_text_exists_anywhere(page, texts: list[str], timeout_ms: int = 1000) -> bool:
    deadline = time.time() + timeout_ms / 1000
    script = """
        (texts) => {
            const isVisible = (el) => {
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return style && style.visibility !== 'hidden' && style.display !== 'none' && rect.width > 0 && rect.height > 0;
            };
            const nodes = Array.from(document.querySelectorAll('button,[role=button],label,a,.ant-btn,.ant-radio-wrapper,span,div'));
            return texts.some(text => nodes.some(el => isVisible(el) && (el.innerText || el.value || '').trim().includes(text)));
        }
    """
    while time.time() < deadline:
        for target in page_and_frames(page):
            try:
                if target.evaluate(script, texts):
                    return True
            except Exception:
                continue
        time.sleep(0.1)
    return False


def fast_click_text_anywhere(page, texts: list[str], timeout_ms: int = 1200, clickable_only: bool = True) -> bool:
    deadline = time.time() + timeout_ms / 1000
    script = r"""
        ({texts, clickableOnly}) => {
            const isVisible = (el) => {
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return style && style.visibility !== 'hidden' && style.display !== 'none' && rect.width > 0 && rect.height > 0;
            };
            const norm = (el) => (el.innerText || el.value || '').trim().replace(/\s+/g, ' ');
            const selector = clickableOnly
                ? 'button,[role=button],a,.ant-btn,label,.ant-radio-wrapper'
                : 'button,[role=button],a,.ant-btn,label,.ant-radio-wrapper,span,div';
            const nodes = Array.from(document.querySelectorAll(selector)).filter(isVisible);
            for (const text of texts) {
                let node = nodes.find(el => norm(el) === text);
                if (!node) node = nodes.find(el => norm(el).includes(text));
                if (!node) continue;
                const clickable = node.closest('button,[role=button],a,label,.ant-btn,.ant-radio-wrapper') || node;
                clickable.click();
                return true;
            }
            return false;
        }
    """
    while time.time() < deadline:
        for target in page_and_frames(page):
            try:
                if target.evaluate(script, {"texts": texts, "clickableOnly": clickable_only}):
                    return True
            except Exception:
                continue
        time.sleep(0.08)
    return False


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



def wait_upload_area_ready(page, timeout_ms: int = 3000) -> bool:
    deadline = time.time() + timeout_ms / 1000
    while time.time() < deadline:
        for target in page_and_frames(page):
            try:
                if target.locator("input[type=file]").count() > 0:
                    return True
            except Exception:
                pass
        if visible_text_exists_anywhere(page, ["上传文件", "点击上传 Excel 文件", "下载模板"], timeout_ms=150):
            return True
        time.sleep(0.1)
    return False


def select_lowest_price_query(page, debug_dir: Path, log_file: Path) -> None:
    # 新版页面进入查价工具后，必须先切到“近N天最低价查询”tab，再选择批量上传。
    # 注意：默认“100天最低价”页也有“批量上传”，不能仅凭批量上传可见就跳过切 tab。
    exact_tab_texts = ["近N天最低价查询", "近 N 天最低价查询"]
    fuzzy_tab_texts = [*exact_tab_texts, "近N天最低价", "近 N 天最低价"]
    deadline = time.time() + 90
    last_log_at = 0.0
    seen_target_tab = False
    clicked_target_tab = False

    def target_tab_visible() -> bool:
        return visible_text_exists_anywhere(page, fuzzy_tab_texts, timeout_ms=700)

    def click_target_tab() -> bool:
        # 优先用 JS 按可见文本点击 tab/button，避免点击到说明文字。
        script = r"""
            (texts) => {
                const isVisible = (el) => {
                    const style = window.getComputedStyle(el);
                    const rect = el.getBoundingClientRect();
                    return style && style.visibility !== 'hidden' && style.display !== 'none' && rect.width > 0 && rect.height > 0;
                };
                const norm = (el) => (el.innerText || el.textContent || '').trim().replace(/\s+/g, ' ');
                const clickableSelector = [
                    'button', '[role=tab]', '[role=button]', 'a', 'label',
                    '.ant-tabs-tab', '.ant-radio-button-wrapper', '.ant-radio-wrapper',
                    '.ant-btn', '.el-tabs__item', '.el-radio-button', '.el-radio'
                ].join(',');
                const likelyTabClass = (el) => /tab|radio|btn|button|item|nav|menu|price|lowest|query/i.test(String(el.className || ''));
                const clickableCandidates = (node, text) => {
                    const candidates = [];
                    const closestClickable = node.closest(clickableSelector);
                    if (closestClickable) candidates.push(closestClickable);
                    candidates.push(node);
                    let current = node.parentElement;
                    while (current && current !== document.body && candidates.length < 10) {
                        const currentText = norm(current);
                        const style = window.getComputedStyle(current);
                        if (
                            currentText.includes(text) &&
                            currentText.length <= 120 &&
                            (style.cursor === 'pointer' || likelyTabClass(current) || current.getAttribute('role'))
                        ) {
                            candidates.push(current);
                        }
                        current = current.parentElement;
                    }
                    return [...new Set(candidates)].filter(Boolean);
                };
                const selectors = [
                    clickableSelector, 'span', 'div'
                ].join(',');
                const nodes = Array.from(document.querySelectorAll(selectors)).filter(isVisible);
                for (const text of texts) {
                    const matches = [
                        ...nodes.filter(el => norm(el) === text),
                        ...nodes.filter(el => norm(el) !== text && norm(el).includes(text))
                    ];
                    for (const node of matches) {
                        for (const clickable of clickableCandidates(node, text)) {
                            try {
                                clickable.scrollIntoView({block: 'center', inline: 'center'});
                                clickable.click();
                                return true;
                            } catch (error) {
                                continue;
                            }
                        }
                    }
                }
                return false;
            }
        """
        for target in page_and_frames(page):
            try:
                if target.evaluate(script, exact_tab_texts):
                    return True
            except Exception:
                continue

        for target in page_and_frames(page):
            try:
                if target.evaluate(script, fuzzy_tab_texts):
                    return True
            except Exception:
                continue
        return False

    def target_tab_confirmed() -> bool:
        script = r"""
            (texts) => {
                const isVisible = (el) => {
                    const style = window.getComputedStyle(el);
                    const rect = el.getBoundingClientRect();
                    return style && style.visibility !== 'hidden' && style.display !== 'none' && rect.width > 0 && rect.height > 0;
                };
                const norm = (el) => (el.innerText || el.textContent || '').trim().replace(/\s+/g, ' ');
                const selectedClasses = [
                    'active', 'selected', 'checked', 'current',
                    'ant-tabs-tab-active', 'ant-radio-button-wrapper-checked', 'ant-radio-wrapper-checked',
                    'el-tabs__item is-active', 'is-active', 'is-checked'
                ];
                const selectedAttrs = (el) =>
                    el.getAttribute('aria-selected') === 'true' ||
                    el.getAttribute('aria-checked') === 'true' ||
                    el.getAttribute('aria-current') === 'true';
                const classSelected = (el) => {
                    const className = String(el.className || '');
                    return selectedClasses.some(item => className.includes(item));
                };
                const rgbValues = (value) => {
                    const match = String(value || '').match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
                    return match ? match.slice(1, 4).map(Number) : null;
                };
                const looksActive = (el) => {
                    const style = window.getComputedStyle(el);
                    const colors = [style.color, style.backgroundColor, style.borderColor];
                    return colors.some(value => {
                        const rgb = rgbValues(value);
                        if (!rgb) return false;
                        const [red, green, blue] = rgb;
                        return blue >= 180 && blue > red + 40 && blue > green + 20;
                    });
                };
                const selector = [
                    'button', '[role=tab]', '[role=button]', 'a', 'label',
                    '.ant-tabs-tab', '.ant-radio-button-wrapper', '.ant-radio-wrapper',
                    '.ant-btn', '.el-tabs__item', '.el-radio-button', '.el-radio', 'span', 'div'
                ].join(',');
                const clickableSelector = [
                    'button', '[role=tab]', '[role=button]', 'a', 'label',
                    '.ant-tabs-tab', '.ant-radio-button-wrapper', '.ant-radio-wrapper',
                    '.ant-btn', '.el-tabs__item', '.el-radio-button', '.el-radio'
                ].join(',');
                const nodes = Array.from(document.querySelectorAll(selector)).filter(isVisible);
                for (const text of texts) {
                    const matches = [
                        ...nodes.filter(el => norm(el) === text),
                        ...nodes.filter(el => norm(el) !== text && norm(el).includes(text))
                    ];
                    for (const node of matches) {
                        const candidates = [];
                        let current = node;
                        while (current && current !== document.body && candidates.length < 8) {
                            candidates.push(current);
                            current = current.parentElement;
                        }
                        const clickable = node.closest(clickableSelector);
                        if (clickable) candidates.push(clickable);
                        for (const candidate of candidates) {
                            if (!candidate) continue;
                            if (selectedAttrs(candidate) || classSelected(candidate) || looksActive(candidate)) return true;
                            const checkedInput = candidate.querySelector && candidate.querySelector('input:checked');
                            if (checkedInput) return true;
                        }
                    }
                }
                return false;
            }
        """
        deadline_confirm = time.time() + 3
        while time.time() < deadline_confirm:
            for target in page_and_frames(page):
                try:
                    if target.evaluate(script, exact_tab_texts):
                        return True
                except Exception:
                    continue
            time.sleep(0.3)
        return False

    while time.time() < deadline:
        try:
            page.wait_for_load_state("domcontentloaded", timeout=1000)
        except Exception:
            pass

        if target_tab_visible():
            seen_target_tab = True
            if not click_target_tab():
                time.sleep(0.5)
                continue

            if not clicked_target_tab:
                log("已点击“近N天最低价查询”入口，等待页面确认。", log_file)
            clicked_target_tab = True
            try:
                page.wait_for_load_state("networkidle", timeout=3000)
            except Exception:
                pass
            if target_tab_confirmed():
                log("已确认进入“近N天最低价查询”页面。", log_file)
                return
            time.sleep(0.5)

        now = time.time()
        if now - last_log_at > 10:
            last_log_at = now
            try:
                log(f"等待查价页面渲染/目标tab出现中，当前URL：{page.url}", log_file)
            except Exception:
                log("等待查价页面渲染/目标tab出现中。", log_file)
        time.sleep(0.7)

    if clicked_target_tab:
        save_debug_snapshot(page, debug_dir, "lowest_price_query_tab_not_confirmed")
        raise RuntimeError(f"已点击“近N天最低价查询”，但无法确认已进入对应页面。已保存调试文件到：{debug_dir}")
    if seen_target_tab:
        save_debug_snapshot(page, debug_dir, "lowest_price_query_tab_click_failed")
        raise RuntimeError(f"已找到“近N天最低价查询”，但点击失败。已保存调试文件到：{debug_dir}")
    save_debug_snapshot(page, debug_dir, "lowest_price_query_tab_not_found")
    raise RuntimeError(f"未找到“近N天最低价查询”。页面可能仍在加载或后台未渲染入口；已保存调试文件到：{debug_dir}")

def select_batch_upload(page, debug_dir: Path) -> None:
    # 测速结果：主定位 js visible text 批量上传；兜底 .ant-radio-wrapper / label。
    if fast_click_text_anywhere(page, ["批量上传"], timeout_ms=1000, clickable_only=False):
        if wait_upload_area_ready(page, timeout_ms=2500):
            return

    for target in page_and_frames(page):
        try:
            target.locator(".ant-radio-wrapper:has-text('批量上传'), label:has-text('批量上传')").first.click(timeout=800)
            if wait_upload_area_ready(page, timeout_ms=2500):
                return
        except Exception:
            continue

    if wait_upload_area_ready(page, timeout_ms=500):
        return
    save_debug_snapshot(page, debug_dir, "batch_upload_not_selected")
    raise RuntimeError(f"未能选中“批量上传”或上传区域未出现。已保存调试文件到：{debug_dir}")


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
        with page.expect_file_chooser(timeout=5000) as chooser_info:
            clicked = click_first_available_anywhere(page, selectors, timeout=2000)
            if not clicked:
                save_debug_snapshot(page, debug_dir, "upload_button_not_found")
                raise RuntimeError(f"找不到上传文件按钮。已保存调试文件到：{debug_dir}")
        chooser_info.value.set_files(str(file_path))
        return
    except Exception as exc:
        save_debug_snapshot(page, debug_dir, "upload_failed")
        raise RuntimeError(f"上传文件失败：{exc}。已保存调试文件到：{debug_dir}") from exc



def get_result_state(page) -> str:
    states: list[str] = []
    for target in page_and_frames(page):
        for selector in [".ant-table-row", ".el-table__row", "tbody tr"]:
            try:
                locator = target.locator(selector)
                count = locator.count()
                if count > 0:
                    first_text = ""
                    try:
                        first_text = locator.first.inner_text(timeout=500).strip()[:120]
                    except Exception:
                        pass
                    states.append(f"{selector}:{count}:{first_text}")
            except Exception:
                pass
        for text in ["暂无数据", "商品名称", "skuid", "100天最低价", "最低价"]:
            try:
                if target.get_by_text(text, exact=False).first.is_visible(timeout=300):
                    states.append(f"text:{text}")
            except Exception:
                pass
    return "|".join(states) if states else "unknown"


def has_result_rows(page) -> bool:
    for target in page_and_frames(page):
        for selector in [".ant-table-row", ".el-table__row", "tbody tr"]:
            try:
                locator = target.locator(selector)
                count = locator.count()
                if count > 0:
                    for index in range(min(count, 5)):
                        try:
                            text = locator.nth(index).inner_text(timeout=500).strip()
                            if text and "暂无数据" not in text:
                                return True
                        except Exception:
                            continue
            except Exception:
                continue
    return False


def wait_query_finished(page, max_wait_seconds: int, before_state: str | None = None) -> None:
    # 不再固定等待 45 秒；查询后至少等 1 秒，然后检测表格出现或刷新即继续导出。
    max_wait_seconds = max(2, int(max_wait_seconds or 2))
    time.sleep(1)
    try:
        page.wait_for_load_state("networkidle", timeout=2000)
    except Exception:
        pass

    deadline = time.time() + max_wait_seconds
    while time.time() < deadline:
        current_state = get_result_state(page)
        if has_result_rows(page):
            if not before_state or current_state != before_state:
                return
            # 如果页面没有清空旧结果，但表格已存在，也不要为了固定等待而卡住太久。
            if time.time() > deadline - max_wait_seconds + 2:
                return
        if before_state and current_state != before_state and current_state != "unknown":
            return
        time.sleep(0.5)

def confirm_export_if_needed(page) -> None:
    # 短轮询处理可能出现的确认弹窗；无弹窗时最多等待约 2 秒。
    deadline = time.time() + 2
    labels = ["确定", "确认", "知道了", "我知道了"]
    while time.time() < deadline:
        clicked = False
        for label in labels:
            if fast_click_text_anywhere(page, [label], timeout_ms=120, clickable_only=True):
                clicked = True
                time.sleep(0.2)
                break
        if not clicked:
            time.sleep(0.15)



def handle_inner_erp_login_if_needed(page, log_file: Path, max_wait_seconds: int = 60) -> None:
    """等待登录页/业务页稳定：看到“内网用户ERP登录”就点击，看到业务页元素就继续。"""
    deadline = time.time() + max_wait_seconds
    login_texts = ["内网用户ERP登录", "内网用户erp登录", "内网用户ERP登陆", "内网用户登录"]
    business_texts = ["查价工具", "批量上传", "上传文件", "开始查询", "运营工作台", "商品价格力", "新人价"]
    clicked_login = False
    last_log_at = 0.0

    while time.time() < deadline:
        # 1) 已经进入业务页/工具页，则结束登录处理。
        try:
            if visible_text_exists_anywhere(page, business_texts, timeout_ms=300):
                if clicked_login:
                    log("已检测到登录后页面元素，继续执行。", log_file)
                return
        except Exception:
            pass

        # 2) 登录按钮出现时直接点击，不再依赖 URL 是否包含 login。
        clicked = False
        try:
            clicked = fast_click_text_anywhere(page, login_texts, timeout_ms=700, clickable_only=False)
        except Exception:
            clicked = False

        if not clicked:
            for target in page_and_frames(page):
                for text in login_texts:
                    try:
                        target.get_by_text(text, exact=False).first.click(timeout=300)
                        clicked = True
                        break
                    except Exception:
                        continue
                if clicked:
                    break

        if clicked:
            clicked_login = True
            log("已点击内网用户ERP登录，等待自动登录跳转。", log_file)
            try:
                page.wait_for_load_state("domcontentloaded", timeout=5000)
            except Exception:
                pass
            time.sleep(1)
            continue

        # 3) 未看到登录按钮也未看到业务页时，短暂等待页面跳转/渲染。
        now = time.time()
        if now - last_log_at > 10:
            last_log_at = now
            try:
                log(f"等待登录页或业务页渲染中，当前URL：{page.url}", log_file)
            except Exception:
                log("等待登录页或业务页渲染中。", log_file)
        time.sleep(0.5)

    # 超时不直接报错：可能用户仍需手动登录，后续步骤如果找不到元素会保存调试文件。
    log("未自动检测到内网ERP登录按钮或业务页元素；如浏览器停在登录页，请手动点击内网用户ERP登录。", log_file)


def main() -> int:
    args = parse_args()
    ensure_playwright()
    from playwright.sync_api import sync_playwright

    files = get_files(args.sku_dir, args.pattern, args.only, args.limit, args.file_list)
    args.profile_dir.mkdir(parents=True, exist_ok=True)
    args.log_file.write_text("", encoding="utf-8")
    log(f"文件来源：{'file-list' if args.file_list else 'directory-pattern'}", args.log_file)
    log(f"文件数量：{len(files)}；文件：{', '.join(path.name for path in files)}", args.log_file)

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
        with timed_step("打开后台页面", args.log_file):
            page.goto(args.url, wait_until="domcontentloaded")

        with timed_step("处理内网ERP登录", args.log_file):
            handle_inner_erp_login_if_needed(page, args.log_file, max_wait_seconds=60)

        log("如果页面仍要求登录，请在打开的浏览器里完成登录；脚本会等待并继续。", args.log_file)
        with timed_step("进入查价工具", args.log_file):
            if visible_text_exists_anywhere(page, ["批量上传", "上传文件", "开始查询"], timeout_ms=2500):
                log("已检测到查价工具页面元素，跳过点击查价工具。", args.log_file)
            else:
                clicked_tool = fast_click_text_anywhere(page, ["查价工具"], timeout_ms=1500, clickable_only=False)
                if not clicked_tool:
                    click_first_available(page, [("text", "查价工具")], timeout=5000)
        with timed_step("选择近N天最低价查询", args.log_file):
            select_lowest_price_query(page, args.log_file.parent / "debug_backend", args.log_file)

        with timed_step("选择批量上传", args.log_file):
            select_batch_upload(page, args.log_file.parent / "debug_backend")

        for index, file_path in enumerate(files, start=1):
            log(f"[{index}/{len(files)}] 准备上传：{file_path.name}", args.log_file)
            with timed_step(f"上传文件 {file_path.name}", args.log_file):
                upload_file(page, file_path, args.log_file.parent / "debug_backend")

            before_query_state = get_result_state(page)
            with timed_step("点击开始查询", args.log_file):
                started = fast_click_text_anywhere(page, ["开始查询"], timeout_ms=1000, clickable_only=True)
                if not started:
                    started = fast_click_text_anywhere(page, ["查询"], timeout_ms=1000, clickable_only=True)
                if not started:
                    started = click_first_available_anywhere(
                        page,
                        [("role_button", "开始查询"), ("css", "button:has-text('开始查询')"), ("role_button", "查询"), ("css", "button:has-text('查询')")],
                        timeout=800,
                    )
            if not started:
                raise RuntimeError("找不到“开始查询”按钮。")
            log(f"[{index}/{len(files)}] 已点击开始查询：{file_path.name}", args.log_file)
            with timed_step("等待查价结果出现/刷新", args.log_file):
                wait_query_finished(page, args.wait_after_query, before_query_state)
            log(f"[{index}/{len(files)}] 查询等待结束，准备点击批量导出。", args.log_file)

            with timed_step("点击批量导出", args.log_file):
                exported = fast_click_text_anywhere(page, ["批量导出"], timeout_ms=1000, clickable_only=True)
                if not exported:
                    exported = click_first_available_anywhere(
                        page,
                        [("role_button", "批量导出"), ("css", "button:has-text('批量导出')")],
                        timeout=800,
                    )
            if not exported:
                raise RuntimeError("找不到“批量导出”按钮。")
            with timed_step("处理导出确认（如有）", args.log_file):
                confirm_export_if_needed(page)
            log(f"[{index}/{len(files)}] 已点击批量导出：{file_path.name}", args.log_file)

            if args.pause_after_each:
                input("已完成当前文件上传、查询和批量导出。检查页面后，按回车继续下一个文件...")

        log("全部文件已完成上传、开始查询并点击批量导出。后台网页操作结束。", args.log_file)
        input("后台网页操作结束。确认后按回车关闭浏览器...")
        context.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())


