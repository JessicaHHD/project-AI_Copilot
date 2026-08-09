from __future__ import annotations

from pathlib import Path
from typing import Any

from .batch import update_manifest
from .config import AppConfig
from .final_pricing import run_final_pricing
from .registration_prepare import run_registration_prepare


def action_success(message: str, outputs: dict[str, Any], metrics: dict[str, Any], warnings: list[str] | None = None) -> dict[str, Any]:
    return {
        "ok": True,
        "message": message,
        "outputs": outputs,
        "metrics": metrics,
        "warnings": warnings or [],
        "error": "",
    }


def action_failure(message: str, error: Exception) -> dict[str, Any]:
    return {
        "ok": False,
        "message": message,
        "outputs": {},
        "metrics": {},
        "warnings": [],
        "error": str(error),
    }


def archived_count(result: dict[str, Any]) -> int:
    return len(result.get("archived_files") or [])


def run_final_pricing_action(config: AppConfig) -> dict[str, Any]:
    try:
        result = run_final_pricing(config)
        update_manifest(config, "阶段1-筛品查价整合", {
            "筛品查价表": str(result["output_file"]),
            "查价文件数": len(result["price_files"]),
            "查价文件来源": result.get("price_file_source", ""),
            "疑似误下载隔离数": result.get("misdownload_count", 0),
            "筛品查价表写入方式": "更新已有表" if result.get("updated_existing") else "新建表",
            "更新前备份文件": str(result.get("backup_file") or ""),
            "最终保留行数": result.get("final_rows", 0),
            "ERP数量": result.get("erp_count", 0),
        })
        warnings: list[str] = []
        if result.get("misdownload_count", 0):
            warnings.append(f"已隔离疑似误下载文件 {result['misdownload_count']} 个，本次未参与新人价计算。")
        if result.get("missing_rows", 0):
            warnings.append(f"有 {result['missing_rows']} 行未匹配到查价结果，请按需检查。")
        return action_success(
            "新人价表已生成/更新。",
            {
                "输出文件": str(result["output_file"]),
                "更新前备份": str(result.get("backup_file") or "无"),
                "查价字段": result.get("lookup_column", ""),
                "ERP汇总文本": result.get("erp_joined_text", ""),
            },
            {
                "筛品行数": result.get("source_rows", 0),
                "匹配查价": result.get("matched_rows", 0),
                "未匹配查价": result.get("missing_rows", 0),
                "新人价>=200剔除": result.get("removed_rows", 0),
                "价格复核剔除": result.get("price_review_excluded_rows", 0),
                "最终保留SKU": result.get("final_rows", 0),
                "ERP数量": result.get("erp_count", 0),
            },
            warnings,
        )
    except Exception as exc:
        return action_failure("新人价表生成失败。", exc)


def run_business_confirmation_action(config: AppConfig, confirm_file: str) -> dict[str, Any]:
    try:
        text = confirm_file.strip().strip('"')
        if not text:
            raise ValueError("请先填写业务确认表完整路径。")
        source_file = Path(text)
        result = run_registration_prepare(config, source_file)
        update_manifest(config, "阶段2-业务确认生成提报", {
            "业务确认表": str(result["source"]),
            "更新后筛品查价表": str(result["final_file"]),
            "提报汇总文件": str(result["summary"]),
            "提报PART文件": [str(path) for path in result["parts"]],
            "提报PART数": len(result["parts"]),
            "提报SKU数": result["rows"],
            "业务确认参与报名SKU": result.get("business_join_skus", 0),
            "业务确认不报名SKU": result.get("business_no_join_skus", 0),
            "业务确认未确认SKU": result.get("business_unknown_skus", 0),
            "提报旧文件归档数": archived_count(result),
            "提报旧文件归档目录": result.get("archive_dir", ""),
        })
        warnings: list[str] = []
        if result.get("business_unknown_skus", 0):
            warnings.append(f"有 {result['business_unknown_skus']} 个 SKU 未确认是否报名，请按需检查业务确认明细。")
        return action_success(
            "业务确认表已读取，提报文件已生成。",
            {
                "更新后的筛品查价表": str(result["final_file"]),
                "提报汇总文件": str(result["summary"]),
                "提报生成日志": str(result["log"]),
                "提报PART文件": [str(path) for path in result["parts"]],
            },
            {
                "可提报SKU": result.get("rows", 0),
                "参与报名SKU": result.get("business_join_skus", 0),
                "不报名SKU": result.get("business_no_join_skus", 0),
                "未确认SKU": result.get("business_unknown_skus", 0),
                "提报PART数": len(result.get("parts") or []),
                "旧提报文件归档数": archived_count(result),
            },
            warnings,
        )
    except Exception as exc:
        return action_failure("业务确认处理失败。", exc)

