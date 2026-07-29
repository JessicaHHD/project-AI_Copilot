from __future__ import annotations

from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

from .config import AppConfig


NO_VALUES = {"否", "n", "no", "不", "不参加", "不报名", "不提报", "退出", "取消", "放弃", "false", "0"}
EXIT_VALUES = {"是", "y", "yes", "true", "1", "退出", "不参加", "不报名", "不提报", "取消", "放弃"}


def normalize(value: object) -> str:
    return str(value or "").strip().replace(" ", "").replace("\n", "").replace("\r", "").replace("\t", "")


def to_price(value: object) -> Decimal | None:
    text = str(value or "").strip().replace(",", "")
    if not text or text in {"-", "--"}:
        return None
    try:
        return Decimal(text).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except InvalidOperation:
        return None


def format_price(value: Decimal) -> float:
    return float(value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def find_index(headers: list[object], candidates: list[str]) -> int | None:
    normalized = [normalize(header).lower() for header in headers]
    for candidate in candidates:
        target = normalize(candidate).lower()
        if target in normalized:
            return normalized.index(target)
    for index, header in enumerate(normalized):
        if any(normalize(candidate).lower() in header for candidate in candidates):
            return index
    return None


def find_header_row(rows: list[tuple[Any, ...]]) -> tuple[int, list[object], int, int, int, str]:
    for row_index, row in enumerate(rows[:30]):
        headers = list(row)
        sku_index = find_index(headers, ["sku", "skuid", "SKU", "SKU(含影)"])
        price_index = find_index(headers, ["期望新人价", "新人价", "首单新人价", "促销价"])
        join_index = find_index(headers, ["是否退出", "是否登记", "是否报名", "是否参加", "报名", "是否提报"])
        if sku_index is not None and price_index is not None and join_index is not None:
            return row_index, headers, sku_index, price_index, join_index, str(headers[join_index] or "")
    raise ValueError("未找到表头行：需要同时包含 SKU、期望新人价/新人价、是否退出/是否报名 字段")


def is_not_joining(value: object, field_name: str) -> bool:
    text = normalize(value).lower()
    if not text:
        return False
    if "退出" in normalize(field_name):
        if text in EXIT_VALUES:
            return True
        return any(keyword in text for keyword in ["不参加", "不报名", "不提报", "退出", "取消", "放弃"])
    if text in NO_VALUES:
        return True
    return any(keyword in text for keyword in ["不参加", "不报名", "不提报", "退出", "取消", "放弃"])


def read_confirmation_rows(path: Path) -> tuple[list[dict[str, Any]], list[str]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        worksheet = workbook.active
        if hasattr(worksheet, "reset_dimensions"):
            worksheet.reset_dimensions()
        all_rows = list(worksheet.iter_rows(values_only=True))
    finally:
        workbook.close()

    header_row, headers, sku_index, price_index, join_index, join_field = find_header_row(all_rows)
    selected: list[dict[str, Any]] = []
    skipped_no = 0
    skipped_missing = 0
    seen_skus: set[str] = set()
    duplicate_skus = 0

    for row in all_rows[header_row + 1:]:
        join_value = row[join_index] if join_index < len(row) else ""
        if is_not_joining(join_value, join_field):
            skipped_no += 1
            continue
        sku = str(row[sku_index] if sku_index < len(row) else "").strip()
        price = to_price(row[price_index] if price_index < len(row) else "")
        if not sku or price is None:
            skipped_missing += 1
            continue
        if sku in seen_skus:
            duplicate_skus += 1
            continue
        seen_skus.add(sku)
        selected.append({"skuid": sku, "促销价": price})

    logs = [
        f"确认表文件：{path}",
        f"识别表头行：第 {header_row + 1} 行",
        f"识别报名判断字段：{join_field}",
        f"报名确认行数：{len(selected)}",
        f"明确退出/不报名跳过：{skipped_no}",
        f"SKU或新人价缺失跳过：{skipped_missing}",
        f"重复SKU跳过：{duplicate_skus}",
    ]
    return selected, logs


def write_submit_file(path: Path, rows: list[dict[str, Any]], sheet_name: str) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = sheet_name
    worksheet.append(["skuid", "促销价"])
    for row in rows:
        worksheet.append([row["skuid"], format_price(row["促销价"])])
    for cell in worksheet[1]:
        cell.font = Font(bold=True)
    worksheet.freeze_panes = "A2"
    worksheet.column_dimensions["A"].width = 22
    worksheet.column_dimensions["B"].width = 12
    workbook.save(path)


def write_log_file(path: Path, logs: list[str]) -> None:
    path.write_text("\n".join(logs) + "\n", encoding="utf-8")


def run_registration_prepare(config: AppConfig, source_file: Path) -> dict[str, Any]:
    if not source_file.exists():
        raise FileNotFoundError(f"采销确认表不存在：{source_file}")
    rows, logs = read_confirmation_rows(source_file)
    if not rows:
        raise ValueError("未提取到任何可提报SKU，请检查 SKU、期望新人价/新人价、是否退出 字段。")

    output_dir = Path(str(config.get("registration_submit_dir")))
    output_dir.mkdir(parents=True, exist_ok=True)
    base_name = str(config.get("registration_base_name", "新人价提报") or "新人价提报")
    chunk_size = int(config.get("registration_chunk_size", 5000) or 5000)
    if chunk_size <= 0:
        raise ValueError("registration_chunk_size 必须大于 0")

    summary_path = output_dir / f"{base_name}.xlsx"
    write_submit_file(summary_path, rows, "汇总")
    part_paths: list[Path] = []
    for start in range(0, len(rows), chunk_size):
        part_no = start // chunk_size + 1
        part_path = output_dir / f"{base_name}_PART{part_no:02d}.xlsx"
        write_submit_file(part_path, rows[start:start + chunk_size], f"PART{part_no:02d}")
        part_paths.append(part_path)
    logs.extend([
        f"输出汇总文件：{summary_path}",
        f"拆分大小：{chunk_size}",
        f"输出PART文件数：{len(part_paths)}",
    ])
    log_path = output_dir / f"{base_name}_生成日志.txt"
    write_log_file(log_path, logs)
    return {
        "source": source_file,
        "summary": summary_path,
        "parts": part_paths,
        "log": log_path,
        "rows": len(rows),
        "logs": logs,
    }
