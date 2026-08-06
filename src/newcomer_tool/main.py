from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from .batch import archive_matching_files, batch_date, batch_name, load_manifest, manifest_path, registration_submit_pattern, update_manifest
from .config import PROJECT_ROOT, load_config, print_config
from .final_pricing import run_final_pricing
from .filtering import run_filtering
from .registration_prepare import run_registration_prepare
from .registration_status_merge import run_registration_status_merge
from .split_sku import run_split


def pause() -> None:
    input("\n按回车继续...")


def confirm(message: str) -> bool:
    return input(f"{message}\n输入 y 继续，其他键取消：").strip().lower() == "y"


def manifest_part_files(config, step_names: list[str], key: str) -> tuple[list[Path], str]:
    manifest = load_manifest(config)
    steps = manifest.get("steps", {}) if isinstance(manifest, dict) else {}
    for step_name in step_names:
        step = steps.get(step_name, {}) if isinstance(steps, dict) else {}
        raw_files = step.get(key, []) if isinstance(step, dict) else []
        if not isinstance(raw_files, list) or not raw_files:
            continue
        files = [Path(str(item)) for item in raw_files]
        missing = [path for path in files if not path.exists()]
        if missing:
            print(f"批次清单中的文件有缺失，已忽略该清单：{step_name}")
            for path in missing[:5]:
                print(f"  缺失：{path}")
            continue
        return files, f"批次清单：{step_name}"
    return [], ""


def directory_part_files(directory: Path, patterns: list[str]) -> tuple[list[Path], str]:
    for pattern in patterns:
        files = sorted(path for path in directory.glob(pattern) if path.is_file() and not path.name.startswith("~$"))
        if files:
            return files, f"目录回退：{pattern}"
    return [], f"目录回退：{patterns[0] if patterns else ''}"


def duplicate_suffix_files(files: list[Path]) -> list[Path]:
    result = []
    for path in files:
        parts = path.stem.rsplit("_", 1)
        if len(parts) == 2 and parts[1].isdigit():
            result.append(path)
    return result


def print_upload_file_scope(files: list[Path], source_label: str) -> None:
    print(f"- 文件来源：{source_label}")
    print(f"- 文件数量：{len(files)}")
    for path in files[:5]:
        print(f"  {path.name}")
    if len(files) > 5:
        print(f"  ... 等 {len(files)} 个文件")
    suspicious = duplicate_suffix_files(files)
    if suspicious:
        print("- 风险提示：发现疑似重跑生成的重名后缀文件，可能是历史文件：")
        for path in suspicious[:5]:
            print(f"  {path.name}")


def write_upload_file_list(config, label: str, files: list[Path]) -> Path:
    directory = config.log_root / "upload_file_lists"
    directory.mkdir(parents=True, exist_ok=True)
    path = directory / f"{batch_name(config)}_{batch_date(config)}_{label}.txt"
    path.write_text("\n".join(str(item.resolve()) for item in files) + "\n", encoding="utf-8")
    return path


def archived_count(result: dict) -> int:
    return len(result.get("archived_files") or [])


def is_legacy_sku_list_warning(row: dict) -> bool:
    item = str(row.get("检查项", ""))
    detail = str(row.get("说明", ""))
    legacy_items = {"目前提报成功SKU名单", "已经提报SKU名单"}
    legacy_keys = ("current_success_skus_file", "submitted_skus_file")
    return row.get("状态") == "预警" and (item in legacy_items or any(key in detail for key in legacy_keys))


def print_filter_result(result: dict) -> None:
    visible_warnings = [row for row in result.get("warnings", []) if not is_legacy_sku_list_warning(row)]
    visible_failures = list(result.get("failures", []))
    display_status = result["status"]
    if display_status == "WARN" and not visible_warnings and not visible_failures:
        display_status = "PASS"

    print("\n筛品完成")
    print(f"状态：{display_status}")
    print(f"源文件：{result['source_file']}")
    print(f"源数据：{result['source_rows']}条")
    print(f"最终结果：{result['rows']}条")
    print(f"结果文件：{result['xlsx']}")
    print(f"SKU文件：{result['sku_file']}")
    print(f"日志文件：{result['log']}")
    if result.get("audit_report"):
        print(f"过程记录：{result['audit_report']}")
    if result.get("csv"):
        print(f"CSV文件：{result['csv']}")

    print("\n筛品日志：")
    for line in result.get("logs", []):
        print(line)

    if visible_warnings or visible_failures:
        print("\n审核提示：")
        for row in visible_failures:
            print(f"- [失败] {row['检查项']}：{row['说明']}")
        for row in visible_warnings:
            print(f"- [预警] {row['检查项']}：{row['说明']}")
    if display_status == "WARN":
        print("建议确认以上预警，无异常可继续进入下一模块。")
    if display_status == "FAIL":
        print("筛品失败，请先处理失败项，不建议进入下一模块。")

def run_filter_step(config) -> dict | None:
    print_config(config)
    if not confirm("准备执行【筛品】。"):
        return None
    result = run_filtering(config)
    print_filter_result(result)
    return result


def run_split_step(config) -> None:
    print_config(config)
    if not confirm("准备执行【提取并拆分查价前SKU】。"):
        return
    result = run_split(config)
    print("\n拆分完成：")
    print(f"- 来源：{result['source']}")
    print(f"- SKU数量：{result['sku_count']}")
    print(f"- 汇总文件：{result['summary']}")
    print(f"- PART文件数：{len(result['parts'])}")
    if archived_count(result):
        print(f"- 已归档旧查价文件：{archived_count(result)}")
        print(f"- 归档目录：{result.get('archive_dir')}")
    for path in result["parts"]:
        print(f"  {path}")
    update_manifest(config, "阶段1-筛品拆分", {
        "查价SKU来源": str(result["source"]),
        "查价SKU汇总": str(result["summary"]),
        "查价SKU_PART文件": [str(path) for path in result["parts"]],
        "查价SKU_PART数": len(result["parts"]),
        "查价旧文件归档数": archived_count(result),
        "查价旧文件归档目录": result.get("archive_dir", ""),
    })

def run_local_pipeline(config) -> None:
    print_config(config)
    if not confirm("准备执行【筛品 + 提取拆分】。"):
        return
    filter_result = run_filtering(config)
    print_filter_result(filter_result)
    if filter_result["status"] == "FAIL":
        print("筛品失败，已停止拆分。")
        return
    split_result = run_split(config, filter_result["sku_file"])
    print("\n本地流程完成：")
    print(f"- 筛品结果：{filter_result['xlsx']}")
    print(f"- 过程记录：{filter_result['audit_report']}")
    print(f"- 筛品SKU文件：{filter_result['sku_file']}")
    print(f"- 拆分SKU数量：{split_result['sku_count']}")
    print(f"- 拆分目录：{config.split_output_dir}")
    if archived_count(split_result):
        print(f"- 已归档旧查价文件：{archived_count(split_result)}")
        print(f"- 归档目录：{split_result.get('archive_dir')}")
    update_manifest(config, "阶段1-筛品拆分", {
        "筛品结果": str(filter_result["xlsx"]),
        "筛品SKU文件": str(filter_result["sku_file"]),
        "查价SKU汇总": str(split_result["summary"]),
        "查价SKU_PART文件": [str(path) for path in split_result["parts"]],
        "查价SKU_PART数": len(split_result["parts"]),
        "查价旧文件归档数": archived_count(split_result),
        "查价旧文件归档目录": split_result.get("archive_dir", ""),
    })


def run_final_pricing_step(config) -> None:
    print_config(config)
    if not confirm("准备执行【查价结果整合并生成新人价】。"):
        return
    result = run_final_pricing(config)
    print("\n最终整合完成：")
    print(f"- 筛品结果：{result['filter_file']}")
    print(f"- 查价文件数：{len(result['price_files'])}")
    print(f"- 查价字段：{result['lookup_column']}")
    print(f"- 筛品行数：{result['source_rows']}")
    print(f"- 匹配查价：{result['matched_rows']}")
    print(f"- 未匹配查价：{result['missing_rows']}")
    print(f"- 剔除新人价>=200：{result['removed_rows']}")
    print(f"- 最终保留：{result['final_rows']}")
    print(f"- 去重销售员ERP：{result['erp_count']}")
    print(f"- 查价文件来源：{result['price_file_source']}")
    print(f"- 疑似误下载隔离：{result['misdownload_count']}")
    print(f"- 写入方式：{'更新已有筛品查价表' if result['updated_existing'] else '新建筛品查价表'}")
    if result.get("backup_file"):
        print(f"- 更新前备份：{result['backup_file']}")
    print(f"- 输出文件：{result['output_file']}")
    update_manifest(config, "阶段1-筛品查价整合", {
        "筛品查价表": str(result["output_file"]),
        "查价文件数": len(result["price_files"]),
        "查价文件来源": result["price_file_source"],
        "疑似误下载隔离数": result["misdownload_count"],
        "筛品查价表写入方式": "更新已有表" if result["updated_existing"] else "新建表",
        "更新前备份文件": str(result.get("backup_file") or ""),
        "最终保留行数": result["final_rows"],
        "ERP数量": result["erp_count"],
    })


def run_script(script_path: Path, args: list[str]) -> None:
    if not script_path.exists():
        raise FileNotFoundError(f"脚本不存在：{script_path}")
    command = [sys.executable, str(script_path), *args]
    print("即将运行：")
    print(" ".join(command))
    subprocess.run(command, check=True)


def run_backend_experiment(config) -> None:
    pattern = f"{config.get('split_base_name')}_PART*.xlsx"
    files, source_label = manifest_part_files(config, ["阶段1-筛品拆分"], "查价SKU_PART文件")
    if not files:
        print("未读取到本批次查价 PART 清单，已回退到目录匹配。目录中历史文件可能被一起上传，请确认。")
        files, source_label = directory_part_files(config.split_output_dir, [pattern])
    if not files:
        print(f"未找到待上传文件：{config.split_output_dir / pattern}")
        return

    print("\n后台上传范围：")
    print(f"当前目录：{config.split_output_dir}")
    print(f"匹配规则：{pattern}")
    print_upload_file_scope(files, source_label)
    print("1 只跑 1 个文件（从 PART01 开始）")
    print("2 只跑 2 个文件（从 PART01 开始）")
    print("3 跑全量文件")
    print("0 取消")
    upload_choice = input("请选择后台上传范围：").strip()

    if upload_choice == "0":
        return
    if upload_choice == "1":
        limit = 1
        selected_files = files[:1]
        scope_label = f"前 1 个文件：{files[0].name}"
    elif upload_choice == "2":
        limit = 2
        selected_files = files[:2]
        preview = "、".join(path.name for path in selected_files)
        scope_label = f"前 {min(2, len(files))} 个文件：{preview}"
    elif upload_choice == "3":
        limit = 0
        selected_files = files
        scope_label = f"全量 {len(files)} 个文件"
    else:
        print("无效选择，已取消。")
        return

    if not confirm(f"准备执行【后台上传并批量导出到邮箱（实验）】。本次范围：{scope_label}"):
        return
    file_list = write_upload_file_list(config, "backend_query_files", selected_files)

    script = PROJECT_ROOT / "src" / "newcomer_tool" / "web_steps" / "backend_upload_export.py"
    args = [
        "--url", str(config.get("backend_url")),
        "--sku-dir", str(config.split_output_dir),
        "--pattern", pattern,
        "--profile-dir", str(config.get("backend_profile_dir")),
        "--log-file", str(config.log_root / "后台查价自动上传日志.txt"),
        "--wait-after-query", str(config.get("backend_wait_after_query")),
        "--wait-after-export", "0",
        "--file-list", str(file_list),
    ]
    run_script(script, args)


def run_outlook_experiment(config) -> None:
    print("\nOutlook客户端邮件下载：")
    print("1 只提取未读邮件下载链接（推荐先测）")
    print("2 直接下载链接到项目输出目录")
    print("0 取消")
    outlook_choice = input("请选择：").strip()

    if outlook_choice == "0":
        return
    if outlook_choice == "1":
        mode = "list"
        confirm_text = "准备执行【只提取 Outlook 未读邮件下载链接】。"
    elif outlook_choice == "2":
        mode = "download"
        confirm_text = "准备执行【直接下载 Outlook 邮件链接】。"
    else:
        print("无效选择，已取消。")
        return

    pattern = f"{config.get('split_base_name')}_PART*.xlsx"
    part_files, part_source = manifest_part_files(config, ["阶段1-筛品拆分"], "查价SKU_PART文件")
    if not part_files:
        print("未读取到本批次查价 PART 清单，Outlook 目标数量已回退到目录匹配。")
        part_files, part_source = directory_part_files(config.split_output_dir, [pattern])
    part_count = len(part_files)
    print(f"本批查价 PART 数量来源：{part_source}")

    print("\n本批目标下载数量：")
    print("1 只处理 1 个链接")
    print("2 只处理 2 个链接")
    print(f"3 按查价前 SKU PART 文件数处理（当前 {part_count} 个）")
    print("4 手动输入数量")
    print("0 取消")
    count_choice = input("请选择：").strip()
    if count_choice == "0":
        return
    if count_choice == "1":
        expected_count = 1
    elif count_choice == "2":
        expected_count = 2
    elif count_choice == "3":
        expected_count = part_count
        if expected_count <= 0:
            print(f"未找到 PART 文件：{config.split_output_dir / pattern}")
            return
    elif count_choice == "4":
        raw_count = input("请输入本批目标链接数量：").strip()
        if not raw_count.isdigit() or int(raw_count) <= 0:
            print("数量无效，已取消。")
            return
        expected_count = int(raw_count)
    else:
        print("无效选择，已取消。")
        return

    if not confirm(f"{confirm_text}\n本次目标数量：{expected_count} 个；只处理最新一批 10 分钟内邮件。"):
        return
    archive_result = {"archived_files": [], "archive_dir": ""}
    download_list_file = config.log_root / "upload_file_lists" / f"{batch_name(config)}_{batch_date(config)}_price_download_files.txt"
    if mode == "download":
        archive_result = archive_matching_files(config.download_output_dir, ["*.xlsx"], "price_download")
        if archived_count(archive_result):
            print(f"已归档旧查价结果文件：{archived_count(archive_result)}")
            print(f"归档目录：{archive_result.get('archive_dir')}")

    script = PROJECT_ROOT / "src" / "newcomer_tool" / "web_steps" / "outlook_download.py"
    args = [
        "--mode", mode,
        "--download-dir", str(config.download_output_dir),
        "--log-file", str(config.log_root / "Outlook客户端下载日志.txt"),
        "--subject-keyword", "商品下载完成通知",
        "--sender-keyword", "selectionserver@jd.com",
        "--url-keyword", ".xlsx",
        "--max-mails", str(config.get("outlook_max_mails") or 80),
        "--expected-count", str(expected_count),
        "--batch-window-minutes", "10",
    ]
    if mode == "download":
        args.extend(["--download-list-file", str(download_list_file)])
    run_script(script, args)
    if mode == "download":
        downloaded_files = []
        if download_list_file.exists():
            downloaded_files = [line.strip() for line in download_list_file.read_text(encoding="utf-8-sig").splitlines() if line.strip()]
        update_manifest(config, "阶段1-Outlook查价下载", {
            "查价结果文件": downloaded_files,
            "查价结果文件数": len(downloaded_files),
            "查价结果旧文件归档数": archived_count(archive_result),
            "查价结果旧文件归档目录": archive_result.get("archive_dir", ""),
            "下载文件清单": str(download_list_file),
        })
        print(f"本次查价结果文件数：{len(downloaded_files)}")
        for path in downloaded_files[:5]:
            print(f"  {path}")


def run_registration_prepare_step(config) -> None:
    default_dir = Path(str(config.get("registration_input_dir", PROJECT_ROOT / "data" / "input" / "提报确认表")))
    print("\n从采销确认表生成提报文件：")
    print(f"建议把确认表放在：{default_dir}")
    print("说明：本步骤会直接更新当前最新筛品查价最终表，写入“业务是否参与报名”和“业务确认明细”。")
    source_text = input("请输入采销确认表Excel完整路径：").strip().strip('"')
    if not source_text:
        print("未输入文件路径，已取消。")
        return
    result = run_registration_prepare(config, Path(source_text))
    print("\n提报文件生成完成：")
    print(f"- 来源：{result['source']}")
    print(f"- 确认报名SKU数：{result['rows']}")
    print(f"- 已更新筛品查价最终表：{result['final_file']}")
    print(f"- 业务确认参与报名SKU：{result.get('business_join_skus', 0)}")
    print(f"- 业务确认不报名SKU：{result.get('business_no_join_skus', 0)}")
    print(f"- 业务确认未确认SKU：{result.get('business_unknown_skus', 0)}")
    print(f"- 汇总文件：{result['summary']}")
    print(f"- PART文件数：{len(result['parts'])}")
    if archived_count(result):
        print(f"- 已归档旧提报文件：{archived_count(result)}")
        print(f"- 归档目录：{result.get('archive_dir')}")
    for path in result["parts"]:
        print(f"  {path}")
    update_manifest(config, "阶段2-业务确认生成提报", {
        "业务确认表": str(result["source"]),
        "已更新筛品查价最终表": str(result["final_file"]),
        "提报SKU汇总": str(result["summary"]),
        "提报PART文件": [str(path) for path in result["parts"]],
        "提报PART数": len(result["parts"]),
        "提报SKU数": result["rows"],
        "提报旧文件归档数": archived_count(result),
        "提报旧文件归档目录": result.get("archive_dir", ""),
        "业务确认参与报名SKU": result.get("business_join_skus", 0),
        "业务确认不报名SKU": result.get("business_no_join_skus", 0),
        "业务确认未确认SKU": result.get("business_unknown_skus", 0),
    })
    print(f"- 生成日志：{result['log']}")

def run_registration_submit_step(config) -> None:
    submit_dir = Path(str(config.get("registration_submit_dir")))
    candidate_patterns = [registration_submit_pattern(config), str(config.get("registration_submit_pattern") or "新人价提报_PART*.xlsx"), "新人价提报_PART*.xlsx"]
    pattern = candidate_patterns[0]
    files, source_label = manifest_part_files(config, ["阶段2-业务确认生成提报"], "提报PART文件")
    if not files:
        print("未读取到本批次提报 PART 清单，已回退到目录匹配。目录中历史文件可能被一起上传，请确认。")
        files, source_label = directory_part_files(submit_dir, candidate_patterns)
        if source_label.startswith("目录回退："):
            pattern = source_label.replace("目录回退：", "", 1)
    if not files:
        print(f"未找到提报文件：{submit_dir / pattern}")
        return

    print("\n新人价后台批量提报文件（安全模式）：")
    print(f"当前目录：{submit_dir}")
    print(f"匹配规则：{pattern}")
    print_upload_file_scope(files, source_label)
    print("1 演练到上传页，不上传不提交（推荐）")
    print("2 上传文件但不提交")
    print("3 真实提交")
    print("0 取消")
    mode_choice = input("请选择运行模式：").strip()
    if mode_choice == "0":
        return
    if mode_choice == "1":
        mode = "dry_run"
        mode_label = "演练到上传页，不上传不提交"
    elif mode_choice == "2":
        mode = "upload_only"
        mode_label = "上传文件但不提交"
    elif mode_choice == "3":
        mode = "submit"
        mode_label = "真实提交"
    else:
        print("无效选择，已取消。")
        return

    print("\n提报文件起始位置：")
    print(f"找到文件数：{len(files)}")
    print(f"第 1 个文件：{files[0].name}")
    if len(files) >= 2:
        print(f"第 2 个文件：{files[1].name}")
    raw_start_index = input("请输入从第几个文件开始（默认 1；已提交PART01则输入 2）：").strip()
    if raw_start_index:
        try:
            start_index = int(raw_start_index)
        except ValueError:
            print("起始序号不是数字，已取消。")
            return
    else:
        start_index = 1
    if start_index < 1 or start_index > len(files):
        print(f"起始序号超出范围：1-{len(files)}，已取消。")
        return
    remaining_files = files[start_index - 1 :]

    print("\n提报文件范围：")
    print("1 从起始位置只跑 1 个文件（默认安全测试）")
    print("2 从起始位置只跑 2 个文件")
    print("3 从起始位置跑到最后")
    print("0 取消")
    scope_choice = input("请选择文件范围：").strip()
    if scope_choice == "0":
        return
    if scope_choice == "1":
        limit = 1
        selected_files = remaining_files[:1]
    elif scope_choice == "2":
        limit = 2
        selected_files = remaining_files[:2]
    elif scope_choice == "3":
        limit = 0
        selected_files = remaining_files
        if input("全量执行风险更高。如确认从起始位置跑到最后，请输入 全量提报 ：").strip() != "全量提报":
            print("未确认全量，已取消。")
            return
    else:
        print("无效选择，已取消。")
        return
    preview = "、".join(path.name for path in selected_files[:5])
    if len(selected_files) > 5:
        preview += f" 等 {len(selected_files)} 个文件"
    scope_label = f"从第 {start_index} 个开始，处理 {len(selected_files)} 个文件：{preview}"

    print("\n本次提报安全确认：")
    print(f"- 运行模式：{mode_label}")
    print(f"- 文件范围：{scope_label}")
    print(f"- 后台URL：{config.get('registration_url')}")
    print(f"- 月份选择阈值：每月 {config.get('registration_month_threshold_day')} 号及以后优先下月")
    if mode == "submit":
        if input("真实提交会点击后台“提交”。如确认真实提交，请输入 确认提报 ：").strip() != "确认提报":
            print("未输入确认提报，已取消。")
            return
    elif not confirm("准备执行安全测试流程。当前模式不会点击提交。"):
        return
    file_list = write_upload_file_list(config, "registration_submit_files", selected_files)

    script = PROJECT_ROOT / "src" / "newcomer_tool" / "web_steps" / "registration_submit.py"
    args = [
        "--url", str(config.get("registration_url")),
        "--submit-dir", str(submit_dir),
        "--pattern", pattern,
        "--profile-dir", str(config.get("registration_profile_dir")),
        "--log-file", str(config.log_root / "新人价提报自动上传日志.txt"),
        "--mode", mode,
        "--start-index", "1",
        "--month-threshold-day", str(config.get("registration_month_threshold_day") or 25),
        "--file-list", str(file_list),
    ]
    args.extend(["--limit", "0"])
    if mode == "submit":
        args.extend(["--confirm-submit", "确认提报"])
    run_script(script, args)


def run_registration_status_export_step(config) -> None:
    download_dir = Path(str(config.get("registration_status_export_dir") or PROJECT_ROOT / "data" / "output" / "提报情况导出"))
    print("\n后台导出提报/审核情况：")
    print(f"导出文件保存目录：{download_dir}")
    print(f"后台URL：{config.get('registration_url')}")
    print("1 演练到提报情况详情页，不点击下载（推荐先测）")
    print("2 点击下载，只触发邮箱推送，不下载邮件")
    print("3 点击下载，并后台检测 Outlook 新邮件自动下载（最多约20分钟）")
    print("4 只从 Outlook 下载最新提报情况邮件（不点后台下载）")
    print("0 取消")
    choice = input("请选择运行模式：").strip()
    if choice == "0":
        return
    if choice == "1":
        mode = "dry_run"
        mode_label = "演练到详情页，不点击下载"
    elif choice == "2":
        mode = "export_only"
        mode_label = "点击下载，只触发邮箱推送"
    elif choice == "3":
        mode = "export_and_download"
        mode_label = "点击下载，并后台检测 Outlook 新邮件自动下载"
    elif choice == "4":
        mode = "download_latest_only"
        mode_label = "只从 Outlook 下载最新提报情况邮件"
    else:
        print("无效选择，已取消。")
        return

    print("\n本次导出确认：")
    print(f"- 运行模式：{mode_label}")
    print(f"- 目标月份规则：每月 {config.get('registration_month_threshold_day')} 号及以后优先下月")
    print(f"- 邮件主题关键词：{config.get('registration_status_mail_subject') or '商品下载完成通知'}")
    print(f"- 邮件发件人关键词：{config.get('registration_status_mail_sender') or '增长权益'}")
    print(f"- 最长等待邮件：{config.get('registration_status_mail_timeout') or 1200} 秒")
    print(f"- 保存目录：{download_dir}")
    if mode == "dry_run":
        if not confirm("准备执行演练流程。当前模式不会点击下载。"):
            return
    elif mode == "download_latest_only":
        if not confirm("准备只读取 Outlook 最新匹配邮件并下载，不会点击后台下载。"):
            return
    else:
        if input("点击下载会触发后台邮件推送。如确认导出，请输入 确认导出 ：").strip() != "确认导出":
            print("未输入确认导出，已取消。")
            return

    script = PROJECT_ROOT / "src" / "newcomer_tool" / "web_steps" / "registration_status_export.py"
    args = [
        "--url", str(config.get("registration_url")),
        "--mode", mode,
        "--profile-dir", str(config.get("registration_profile_dir")),
        "--log-file", str(config.log_root / "新人价提报情况导出日志.txt"),
        "--download-dir", str(download_dir),
        "--month-threshold-day", str(config.get("registration_month_threshold_day") or 25),
        "--mail-subject-keyword", str(config.get("registration_status_mail_subject") or "商品下载完成通知"),
        "--mail-sender-keyword", str(config.get("registration_status_mail_sender") or "增长权益"),
        "--max-mails", str(max(int(config.get("outlook_max_mails") or 80), 80)),
        "--mail-timeout", str(config.get("registration_status_mail_timeout") or 1200),
    ]
    if mode not in {"dry_run", "download_latest_only"}:
        args.extend(["--confirm-export", "确认导出"])
    run_script(script, args)


def run_registration_status_merge_step(config) -> None:
    print("\n整理提报情况并回填最终结果：")
    print("将自动读取：")
    print(f"- 提报情况导出目录：{config.get('registration_status_export_dir')}")
    print(f"- 最终结果目录：{config.output_root / '最终结果'}")
    print(f"- 采销确认表目录：{config.get('registration_input_dir')}")
    print("输出：新的最终结果 Excel，不覆盖原文件。")
    if not confirm("准备执行【整理提报情况并回填最终结果】。"):
        return
    result = run_registration_status_merge(config)
    print("\n提报情况整理完成：")
    print(f"- 提报情况文件：{result['status_file']}")
    print(f"- 最终结果源文件：{result['final_file']}")
    print(f"- 采销确认表：{result['confirm_file']}")
    print(f"- 提报记录数：{result['status_records']}")
    print(f"- 提报情况去重SKU：{result['merged_skus']}")
    print(f"- 匹配最终结果SKU：{result['matched_skus']}")
    print(f"- 未提报/未匹配SKU：{result['missing_skus']}")
    print(f"- 后台额外提报SKU：{result.get('extra_skus', 0)}")
    print(f"- 已追加额外SKU：{result.get('appended_extra_skus', 0)}")
    print(f"- 业务确认不报名SKU：{result.get('business_no_join_skus', 0)}")
    print(f"- 输出文件：{result['output_file']}")


def run_registration_menu(config) -> None:
    print("\n提报/审核整理：")
    print("1 从采销确认表生成提报文件")
    print("2 后台批量提报文件（安全测试）")
    print("3 后台导出提报/审核情况")
    print("4 整理提报情况并回填最终结果")
    print("0 返回")
    choice = input("请选择：").strip()
    if choice == "1":
        run_registration_prepare_step(config)
    elif choice == "2":
        run_registration_submit_step(config)
    elif choice == "3":
        run_registration_status_export_step(config)
    elif choice == "4":
        run_registration_status_merge_step(config)
    elif choice == "0":
        return
    else:
        print("无效选择。")



def latest_business_confirm_file(config) -> Path | None:
    directory = Path(str(config.get("registration_input_dir", PROJECT_ROOT / "data" / "input" / "提报确认表")))
    if not directory.exists():
        return None
    patterns = [f"{batch_name(config)}_业务已确认表_*.xlsx", "*业务已确认表_*.xlsx", "*新人价采销确认表*.xlsx", "*.xlsx"]
    for pattern in patterns:
        files = sorted(
            (path for path in directory.glob(pattern) if path.is_file() and not path.name.startswith("~$")),
            key=lambda item: item.stat().st_mtime,
            reverse=True,
        )
        if files:
            return files[0]
    return None


def print_manifest(config) -> None:
    path = manifest_path(config)
    manifest = load_manifest(config)
    print("\n本批次文件清单：")
    print(f"- 批次：{batch_name(config)}")
    print(f"- 日期：{batch_date(config)}")
    print(f"- 清单文件：{path}")
    steps = manifest.get("steps", {})
    if not steps:
        print("- 暂无已记录步骤。")
        return
    for step, data in steps.items():
        print(f"\n【{step}】")
        for key, value in data.items():
            print(f"- {key}: {value}")


def workflow_stage_1(config) -> None:
    print_config(config)
    print("\n阶段1：筛品 → 拆分SKU → 后台查价 → Outlook下载 → 最终整合")
    print("说明：后台查价和 Outlook 下载仍会在各自步骤内二次确认。")
    if not confirm("准备开始/继续【筛品查价整合】。"):
        return
    filter_result = run_filtering(config)
    print_filter_result(filter_result)
    if filter_result["status"] == "FAIL":
        print("筛品失败，已停止后续流程。")
        return
    split_result = run_split(config, filter_result["sku_file"])
    print("\n拆分完成：")
    print(f"- SKU数量：{split_result['sku_count']}")
    print(f"- 汇总文件：{split_result['summary']}")
    print(f"- PART文件数：{len(split_result['parts'])}")
    if archived_count(split_result):
        print(f"- 已归档旧查价文件：{archived_count(split_result)}")
        print(f"- 归档目录：{split_result.get('archive_dir')}")
    update_manifest(config, "阶段1-筛品拆分", {
        "筛品结果": str(filter_result["xlsx"]),
        "筛品SKU文件": str(filter_result["sku_file"]),
        "查价SKU汇总": str(split_result["summary"]),
        "查价SKU_PART文件": [str(path) for path in split_result["parts"]],
        "查价SKU_PART数": len(split_result["parts"]),
        "查价旧文件归档数": archived_count(split_result),
        "查价旧文件归档目录": split_result.get("archive_dir", ""),
    })

    if confirm("是否继续执行【后台查价上传并导出到邮箱】？"):
        run_backend_experiment(config)
    else:
        print("已暂停。下一步：确认后从菜单 4 或工作流向导继续后台查价。")
        return

    if confirm("是否继续执行【Outlook 下载查价结果】？"):
        run_outlook_experiment(config)
    else:
        print("已暂停。下一步：确认邮件到达后从菜单 5 或工作流向导继续下载。")
        return

    if confirm("是否继续执行【查价结果整合并生成筛品查价表】？"):
        final_result = run_final_pricing(config)
        print("\n最终整合完成：")
        print(f"- 输出文件：{final_result['output_file']}")
        print(f"- 最终保留：{final_result['final_rows']}")
        print(f"- 去重销售员ERP：{final_result['erp_count']}")
        print(f"- 查价文件来源：{final_result['price_file_source']}")
        print(f"- 疑似误下载隔离：{final_result['misdownload_count']}")
        print(f"- 写入方式：{'更新已有筛品查价表' if final_result['updated_existing'] else '新建筛品查价表'}")
        if final_result.get("backup_file"):
            print(f"- 更新前备份：{final_result['backup_file']}")
        update_manifest(config, "阶段1-筛品查价整合", {
            "筛品查价表": str(final_result["output_file"]),
            "查价文件数": len(final_result["price_files"]),
            "查价文件来源": final_result["price_file_source"],
            "疑似误下载隔离数": final_result["misdownload_count"],
            "筛品查价表写入方式": "更新已有表" if final_result["updated_existing"] else "新建表",
            "更新前备份文件": str(final_result.get("backup_file") or ""),
            "最终保留行数": final_result["final_rows"],
            "ERP数量": final_result["erp_count"],
        })
        print("\n阶段1完成。下一步人工动作：迁移筛品查价表到线上表，使用 ERP 汇总文本发业务确认邮件。")


def workflow_stage_2(config) -> None:
    print("\n阶段2：业务确认后生成提报文件")
    default_file = latest_business_confirm_file(config)
    print(f"业务确认表目录：{config.get('registration_input_dir')}")
    if default_file:
        print(f"自动识别到最新业务确认表：{default_file}")
    source_text = input("请输入业务确认表完整路径；直接回车使用自动识别文件：").strip().strip('"')
    source_file = Path(source_text) if source_text else default_file
    if source_file is None:
        print("未找到业务确认表，已取消。")
        return
    if not confirm(f"准备读取业务确认表，生成提报文件，并更新当前最新筛品查价最终表：\n{source_file}"):
        return
    result = run_registration_prepare(config, source_file)
    print("\n提报文件生成完成：")
    print(f"- 确认报名SKU数：{result['rows']}")
    print(f"- 已更新筛品查价最终表：{result['final_file']}")
    print(f"- 业务确认参与报名SKU：{result.get('business_join_skus', 0)}")
    print(f"- 业务确认不报名SKU：{result.get('business_no_join_skus', 0)}")
    print(f"- 业务确认未确认SKU：{result.get('business_unknown_skus', 0)}")
    print(f"- 汇总文件：{result['summary']}")
    print(f"- PART文件数：{len(result['parts'])}")
    if archived_count(result):
        print(f"- 已归档旧提报文件：{archived_count(result)}")
        print(f"- 归档目录：{result.get('archive_dir')}")
    for path in result["parts"]:
        print(f"  {path}")
    update_manifest(config, "阶段2-业务确认生成提报", {
        "业务确认表": str(result["source"]),
        "已更新筛品查价最终表": str(result["final_file"]),
        "提报SKU汇总": str(result["summary"]),
        "提报PART文件": [str(path) for path in result["parts"]],
        "提报PART数": len(result["parts"]),
        "提报SKU数": result["rows"],
        "提报旧文件归档数": archived_count(result),
        "提报旧文件归档目录": result.get("archive_dir", ""),
        "业务确认参与报名SKU": result.get("business_join_skus", 0),
        "业务确认不报名SKU": result.get("business_no_join_skus", 0),
        "业务确认未确认SKU": result.get("business_unknown_skus", 0),
    })
    print("\n阶段2完成。下一步：进入阶段3后台提报，建议先 dry_run。")


def workflow_stage_4(config) -> None:
    print("\n阶段4：导出并整理审核结果")
    print("说明：点击后台下载会触发邮件推送，仍需按菜单二次确认。")
    if not confirm("是否进入审核结果导出流程？"):
        return
    run_registration_status_export_step(config)
    if confirm("是否继续整理提报情况并回填最终结果？"):
        result = run_registration_status_merge(config)
        print("\n提报情况整理完成：")
        print(f"- 输出文件：{result['output_file']}")
        print(f"- 后台额外提报SKU：{result.get('extra_skus', 0)}")
        print(f"- 已追加额外SKU：{result.get('appended_extra_skus', 0)}")
        print(f"- 业务确认不报名SKU：{result.get('business_no_join_skus', 0)}")
        update_manifest(config, "阶段4-审核整理", {
            "提报情况文件": str(result["status_file"]),
            "提报情况整理表": str(result["output_file"]),
            "提报记录数": result["status_records"],
            "提报情况去重SKU": result["merged_skus"],
            "后台额外提报SKU": result.get("extra_skus", 0),
            "已追加额外SKU": result.get("appended_extra_skus", 0),
            "业务确认不报名SKU": result.get("business_no_join_skus", 0),
        })
        print("\n阶段4完成。该提报情况整理表将作为下一轮复提判断依据。")


def run_workflow_menu(config) -> None:
    print("\n新人价工作流向导：")
    print(f"当前批次：{batch_name(config)}")
    print("1 开始/继续筛品查价整合")
    print("2 业务确认后生成提报文件")
    print("3 后台提报")
    print("4 导出并整理审核结果")
    print("5 查看本批次文件清单")
    print("0 返回")
    choice = input("请选择：").strip()
    if choice == "1":
        workflow_stage_1(config)
    elif choice == "2":
        workflow_stage_2(config)
    elif choice == "3":
        run_registration_submit_step(config)
    elif choice == "4":
        workflow_stage_4(config)
    elif choice == "5":
        print_manifest(config)
    elif choice == "0":
        return
    else:
        print("无效选择。")


def main() -> int:
    while True:
        config = load_config()
        print("\n========== 新人价自动化工具 MVP ==========")
        print("1 筛品")
        print("2 提取并拆分查价前SKU")
        print("3 一键执行本地流程")
        print("4 后台上传并批量导出到邮箱（实验）")
        print("5 Outlook邮件下载（实验）")
        print("6 查看当前配置")
        print("7 查价结果整合并生成新人价")
        print("8 提报/审核整理")
        print("9 新人价工作流向导")
        print("0 退出")
        choice = input("请选择：").strip()
        try:
            if choice == "1":
                run_filter_step(config)
                pause()
            elif choice == "2":
                run_split_step(config)
                pause()
            elif choice == "3":
                run_local_pipeline(config)
                pause()
            elif choice == "4":
                run_backend_experiment(config)
                pause()
            elif choice == "5":
                run_outlook_experiment(config)
                pause()
            elif choice == "6":
                print_config(config)
                pause()
            elif choice == "7":
                run_final_pricing_step(config)
                pause()
            elif choice == "8":
                run_registration_menu(config)
                pause()
            elif choice == "9":
                run_workflow_menu(config)
                pause()
            elif choice == "0":
                return 0
            else:
                print("无效选择。")
        except Exception as exc:
            print(f"\n执行失败：{exc}")
            pause()


if __name__ == "__main__":
    raise SystemExit(main())


