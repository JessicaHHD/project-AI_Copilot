from __future__ import annotations

from collections import Counter
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any
import warnings

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font

from .config import AppConfig


warnings.filterwarnings("ignore", message="Workbook contains no default style, apply openpyxl's default")


def normalize(value: str) -> str:
    return value.replace(" ", "").replace("\n", "").replace("\r", "").replace("\t", "")


def find_column(headers: list[str], candidates: list[str], required: bool = True) -> str | None:
    normalized = {normalize(header): header for header in headers}
    for candidate in candidates:
        if normalize(candidate) in normalized:
            return normalized[normalize(candidate)]
    for header in headers:
        compact = normalize(header)
        if any(normalize(candidate) in compact for candidate in candidates if candidate):
            return header
    if required:
        raise KeyError("缺少必要字段：" + "、".join([item for item in candidates if item]))
    return None


def to_decimal(value: object) -> Decimal | None:
    text = str(value or "").strip().replace(",", "")
    if not text or text in {"-", "--"}:
        return None
    try:
        return Decimal(text)
    except InvalidOperation:
        return None


def money(value: Decimal) -> Decimal:
    return value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def format_value(value: object) -> object:
    if isinstance(value, Decimal):
        return float(money(value))
    return value


def read_sheet(path: Path, sheet_name: str | None = None) -> tuple[list[str], list[dict[str, Any]]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        worksheets = [workbook[sheet_name]] if sheet_name and sheet_name in workbook.sheetnames else list(workbook.worksheets)
        for worksheet in worksheets:
            headers, rows = read_worksheet(worksheet)
            if any(headers):
                return headers, rows
        return [], []
    finally:
        workbook.close()


def read_filter_logs(path: Path) -> list[str]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        if "筛品日志" not in workbook.sheetnames:
            return ["未读取到筛品日志：源筛品结果文件无“筛品日志”sheet。"]
        worksheet = workbook["筛品日志"]
        if hasattr(worksheet, "reset_dimensions"):
            worksheet.reset_dimensions()
        lines: list[str] = []
        for row in worksheet.iter_rows(values_only=True):
            values = [str(value).strip() for value in row if str(value or "").strip()]
            if values:
                lines.append(" | ".join(values))
        return lines or ["筛品日志为空。"]
    finally:
        workbook.close()


def read_worksheet(worksheet) -> tuple[list[str], list[dict[str, Any]]]:
    if hasattr(worksheet, "reset_dimensions"):
        worksheet.reset_dimensions()
    rows_iter = worksheet.iter_rows(values_only=True)
    try:
        headers = [str(value or "").strip() for value in next(rows_iter)]
    except StopIteration:
        return [], []
    rows: list[dict[str, Any]] = []
    for row in rows_iter:
        rows.append({headers[index]: value for index, value in enumerate(row[: len(headers)])})
    return headers, rows


def latest_filter_file(config: AppConfig) -> Path:
    configured = str(config.get("final_filter_file", "") or "").strip()
    if configured:
        path = Path(configured)
        if path.exists():
            return path
        raise FileNotFoundError(f"筛品结果文件不存在：{path}")
    candidates = sorted(
        config.filter_output_dir.glob("新人价筛品结果_*.xlsx"),
        key=lambda item: item.stat().st_mtime,
        reverse=True,
    )
    if not candidates:
        raise FileNotFoundError(f"未找到筛品结果文件：{config.filter_output_dir / '新人价筛品结果_*.xlsx'}")
    return candidates[0]


def price_files(config: AppConfig) -> list[Path]:
    directory = Path(str(config.get("final_price_dir", "") or config.download_output_dir))
    pattern = str(config.get("final_price_pattern", "*.xlsx") or "*.xlsx")
    if not directory.exists():
        raise FileNotFoundError(f"查价结果目录不存在：{directory}")
    files = sorted(path for path in directory.glob(pattern) if not path.name.startswith("~$"))
    if not files:
        raise FileNotFoundError(f"未找到查价结果文件：{directory / pattern}")
    return files


def date_label(config: AppConfig) -> str:
    configured = str(config.get("final_price_date_label", "") or "").strip()
    if configured:
        return configured
    now = datetime.now()
    return f"{now.month}.{now.day}"


def writable_output_file(path: Path) -> Path:
    if not path.exists():
        return path
    try:
        with path.open("a+b"):
            return path
    except PermissionError:
        stamp = datetime.now().strftime("%H%M%S")
        return path.with_name(f"{path.stem}_{stamp}{path.suffix}")


def merge_price_rows(files: list[Path], config: AppConfig) -> tuple[list[str], list[dict[str, Any]], dict[str, Decimal], list[str]]:
    all_headers: list[str] = []
    merged_rows: list[dict[str, Any]] = []
    price_by_sku: dict[str, Decimal] = {}
    duplicate_skus: Counter[str] = Counter()
    sku_candidates = [str(config.get("final_price_sku_column", "skuId")), "skuId", "sku", "SKU", "SKU(含影)"]
    price_candidates = [str(config.get("final_price_value_column", "100天最低价")), "100天最低价"]

    for file in files:
        headers, rows = read_sheet(file, None)
        if not headers:
            continue
        for header in headers:
            if header and header not in all_headers:
                all_headers.append(header)
        try:
            sku_col = find_column(headers, sku_candidates)
            price_col = find_column(headers, price_candidates)
        except KeyError as exc:
            raise KeyError(f"{file.name}：{exc}") from exc
        for row in rows:
            row = dict(row)
            row["来源文件"] = file.name
            merged_rows.append(row)
            sku = str(row.get(sku_col, "") or "").strip()
            price = to_decimal(row.get(price_col))
            if not sku or price is None:
                continue
            if sku in price_by_sku:
                duplicate_skus[sku] += 1
                price_by_sku[sku] = min(price_by_sku[sku], price)
            else:
                price_by_sku[sku] = price

    if "来源文件" not in all_headers:
        all_headers.append("来源文件")
    logs = [
        f"查价文件数：{len(files)}",
        f"查价明细行数：{len(merged_rows)}",
        f"可匹配价格SKU数：{len(price_by_sku)}",
        f"重复SKU数：{len(duplicate_skus)}，重复时取100天最低价较小值。",
    ]
    return all_headers, merged_rows, price_by_sku, logs


def calculate_newcomer_price(base_price: Decimal | None, mapping: str) -> Decimal | None:
    if base_price is None:
        return None
    subsidy = Decimal("8") if mapping.strip() == "自营" else Decimal("6.5")
    result = base_price - subsidy
    if result <= 0:
        return Decimal("0.01")
    return money(result)


def write_rows(workbook: Workbook, title: str, headers: list[str], rows: list[dict[str, Any]]) -> None:
    sheet = workbook.create_sheet(title)
    sheet.append(headers)
    for row in rows:
        sheet.append([format_value(row.get(header, "")) for header in headers])
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    sheet.freeze_panes = "A2"
    for column in sheet.columns:
        sheet.column_dimensions[column[0].column_letter].width = min(max(len(str(column[0].value or "")) + 2, 12), 32)


def write_log_sheet(workbook: Workbook, logs: list[str]) -> None:
    sheet = workbook.create_sheet("整合日志")
    for line in logs:
        sheet.append([line])
    sheet.column_dimensions["A"].width = 120
    for row in sheet.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")


def unique_erp_values(rows: list[dict[str, Any]], erp_column: str) -> list[str]:
    values = {
        str(row.get(erp_column, "") or "").strip()
        for row in rows
        if str(row.get(erp_column, "") or "").strip()
    }
    return sorted(values)


def write_erp_sheet(workbook: Workbook, erp_values: list[str]) -> str:
    joined_text = ";".join(erp_values)
    sheet = workbook.create_sheet("ERP汇总")
    sheet.append(["销售员ERP去重合并", joined_text])
    sheet.append([])
    sheet.append(["销售员ERP帐号"])
    for value in erp_values:
        sheet.append([value])
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    sheet[3][0].font = Font(bold=True)
    sheet.freeze_panes = "A4"
    sheet.column_dimensions["A"].width = 24
    sheet.column_dimensions["B"].width = 120
    sheet[1][1].alignment = Alignment(wrap_text=True, vertical="top")
    return joined_text


def run_final_pricing(config: AppConfig) -> dict[str, Any]:
    filter_file = latest_filter_file(config)
    files = price_files(config)
    label = date_label(config)
    lookup_column = f"{label}_查价"

    filter_headers, filter_rows = read_sheet(filter_file, "筛选结果")
    if not filter_headers:
        raise ValueError(f"筛品结果文件为空：{filter_file}")
    filter_logs = read_filter_logs(filter_file)
    sku_col = find_column(filter_headers, [str(config.get("sku_column")), "SKU(含影)", "sku", "SKU"])
    mapping_col = find_column(filter_headers, [str(config.get("mapping_column")), "1自营3POP_映射"])
    erp_col = find_column(filter_headers, [str(config.get("erp_column")), "销售员ERP帐号", "销售员ERP账号"])

    price_headers, price_rows, price_by_sku, price_logs = merge_price_rows(files, config)
    final_headers = list(filter_headers)
    for header in [lookup_column, "新人价"]:
        if header not in final_headers:
            final_headers.append(header)

    matched_count = 0
    missing_count = 0
    removed_count = 0
    final_rows: list[dict[str, Any]] = []
    for row in filter_rows:
        row = dict(row)
        sku = str(row.get(sku_col, "") or "").strip()
        base_price = price_by_sku.get(sku)
        newcomer_price = calculate_newcomer_price(base_price, str(row.get(mapping_col, "") or ""))
        if base_price is None:
            missing_count += 1
            row[lookup_column] = ""
            row["新人价"] = ""
            final_rows.append(row)
            continue
        matched_count += 1
        row[lookup_column] = money(base_price)
        row["新人价"] = newcomer_price
        if newcomer_price is not None and newcomer_price >= Decimal("200"):
            removed_count += 1
            continue
        final_rows.append(row)

    output_dir = config.output_root / "最终结果"
    output_dir.mkdir(parents=True, exist_ok=True)
    date_tag = str(config.get("date_tag"))
    output_file = output_dir / f"新人价最终结果_{date_tag}_{label.replace('.', '')}.xlsx"
    output_file = writable_output_file(output_file)

    logs = [
        "【筛品日志】",
        *filter_logs,
        "",
        "【最终整合日志】",
        f"筛品结果文件：{filter_file}",
        f"查价日期字段：{lookup_column}",
        *price_logs,
        f"筛品结果行数：{len(filter_rows)}",
        f"查价匹配行数：{matched_count}",
        f"未匹配查价行数：{missing_count}",
        f"剔除新人价大于等于200元：{removed_count}",
        f"最终保留行数：{len(final_rows)}",
    ]
    erp_values = unique_erp_values(final_rows, erp_col)
    logs.append(f"最终保留商品去重销售员ERP数：{len(erp_values)}")

    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)
    write_rows(workbook, "筛选结果", final_headers, final_rows)
    write_rows(workbook, "查价结果合并", price_headers, price_rows)
    erp_joined_text = write_erp_sheet(workbook, erp_values)
    write_log_sheet(workbook, logs)
    workbook.save(output_file)

    return {
        "filter_file": filter_file,
        "price_files": files,
        "output_file": output_file,
        "lookup_column": lookup_column,
        "source_rows": len(filter_rows),
        "matched_rows": matched_count,
        "missing_rows": missing_count,
        "removed_rows": removed_count,
        "final_rows": len(final_rows),
        "erp_count": len(erp_values),
        "erp_joined_text": erp_joined_text,
        "logs": logs,
    }
