from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from .config import PROJECT_ROOT, load_config, print_config
from .filtering import run_filtering
from .split_sku import run_split


def pause() -> None:
    input("\n按回车继续...")


def confirm(message: str) -> bool:
    return input(f"{message}\n输入 y 继续，其他键取消：").strip().lower() == "y"


def print_filter_result(result: dict) -> None:
    print("\n筛品完成")
    print(f"状态：{result['status']}")
    print(f"源文件：{result['source_file']}")
    print(f"源数据：{result['source_rows']}条")
    print(f"最终结果：{result['rows']}条")
    print(f"预警：{len(result['warnings'])}个")
    for warning in result["warnings"][:5]:
        print(f"- {warning['说明']}")
    if result["failures"]:
        print(f"失败：{len(result['failures'])}个")
        for failure in result["failures"][:5]:
            print(f"- {failure['说明']}")
    print(f"结果文件：{result['xlsx']}")
    print(f"SKU文件：{result['sku_file']}")
    print(f"审核报告：{result['audit_report']}")
    print(f"日志文件：{result['log']}")
    if result.get("csv"):
        print(f"CSV文件：{result['csv']}")
    if result["status"] == "WARN":
        print("建议查看审核报告后再进入下一模块。")
    if result["status"] == "FAIL":
        print("筛品失败，不建议进入下一模块。")


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
    for path in result["parts"]:
        print(f"  {path}")


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
    print(f"- 审核报告：{filter_result['audit_report']}")
    print(f"- 筛品SKU文件：{filter_result['sku_file']}")
    print(f"- 拆分SKU数量：{split_result['sku_count']}")
    print(f"- 拆分目录：{config.split_output_dir}")


def run_script(script_path: Path, args: list[str]) -> None:
    if not script_path.exists():
        raise FileNotFoundError(f"脚本不存在：{script_path}")
    command = [sys.executable, str(script_path), *args]
    print("即将运行：")
    print(" ".join(command))
    subprocess.run(command, check=True)


def run_backend_experiment(config) -> None:
    if not confirm("准备执行【后台上传并批量导出到邮箱（实验）】。建议先只跑 PART01。"):
        return
    script = PROJECT_ROOT / "src" / "newcomer_tool" / "web_steps" / "backend_upload_export.py"
    args = ["--url", str(config.get("backend_url")), "--sku-dir", str(config.split_output_dir), "--pattern", f"{config.get('split_base_name')}_PART*.xlsx", "--profile-dir", str(config.get("backend_profile_dir")), "--log-file", str(PROJECT_ROOT / "logs" / "后台查价自动上传日志.txt"), "--wait-after-query", str(config.get("backend_wait_after_query")), "--wait-after-export", str(config.get("backend_wait_after_export")), "--pause-after-each"]
    only_file = str(config.get("backend_only_file", "") or "").strip()
    if only_file:
        args.extend(["--only", only_file])
    run_script(script, args)


def run_outlook_experiment(config) -> None:
    if not confirm("准备执行【Outlook邮件下载（实验）】。默认手动模式。"):
        return
    script = PROJECT_ROOT / "src" / "newcomer_tool" / "web_steps" / "outlook_download.py"
    args = ["--mail-url", str(config.get("outlook_url")), "--download-dir", str(config.download_output_dir), "--profile-dir", str(config.get("outlook_profile_dir")), "--search", str(config.get("outlook_search")), "--max-mails", str(config.get("outlook_max_mails")), "--manual"]
    run_script(script, args)


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
            elif choice == "0":
                return 0
            else:
                print("无效选择。")
        except Exception as exc:
            print(f"\n执行失败：{exc}")
            pause()


if __name__ == "__main__":
    raise SystemExit(main())
