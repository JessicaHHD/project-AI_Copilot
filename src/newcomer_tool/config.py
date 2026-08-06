from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any


PROJECT_ROOT = Path(__file__).resolve().parents[2]
CONFIG_PATH = PROJECT_ROOT / "config.yaml"


DEFAULTS: dict[str, Any] = {
    "batch_name": "8月第1批提报",
    "batch_date": "",
    "source_file": "",
    "output_root": str(PROJECT_ROOT / "data" / "output"),
    "log_root": str(PROJECT_ROOT / "logs"),
    "date_tag": "20260727",
    "sku_column": "SKU(含影)",
    "product_name_column": "sku名称",
    "mapping_column": "1自营3POP_映射",
    "shop_column": "店铺名称",
    "department_column": "1级部门名称（类目运营）",
    "category2_column": "商品二级分类名称",
    "erp_column": "销售员ERP帐号",
    "keep_mapping_values": ["自营"],
    "keep_pop_shop_keyword": "京喜",
    "exclude_category2": ["健康服务套餐", "其他服务", "特殊商品"],
    "exclude_name_keywords": ["非卖品", "赠品"],
    "exclude_erp_accounts": ["zhaodong.103"],
    "target_departments": ["医药", "营养保健", "医疗器械", "消费器械"],
    "metric_sum_columns": ["去重计数_用户PIN加密", "求和_成交金额", "去重计数_销售订单编号", "求和_销售数量"],
    "current_success_skus_file": "",
    "submitted_skus_file": "",
    "filter_round_mode": "auto",
    "resubmit_final_file": "",
    "resubmit_success_keywords": ["转化不满足", "删促"],
    "resubmit_exit_keywords": ["退出", "不参加", "不报名", "不提报", "取消", "放弃"],
    "resubmit_price_review_enabled": True,
    "resubmit_price_diff_abs_threshold": 10,
    "resubmit_price_diff_rate_threshold": 0.2,
    "resubmit_price_review_detail_sheet": True,
    "strict_field_validation": True,
    "export_csv": False,
    "export_removed_details": False,
    "debug_mode": False,
    "split_base_name": "8月第一批查价前sku",
    "split_chunk_size": 5000,
    "split_source_sheet": "汇总",
    "split_source_column": "sku",
    "backend_url": "https://yx.jd.com/user-operate#/customize/findLowestActivity?pathkey=goodspool",
    "backend_profile_dir": str(PROJECT_ROOT / "browser_profiles" / "backend"),
    "backend_wait_after_query": 45,
    "backend_wait_after_export": 10,
    "backend_only_file": "8月第一批查价前sku_PART01.xlsx",
    "outlook_url": "https://outlook.office.com/mail/",
    "outlook_profile_dir": str(PROJECT_ROOT / "browser_profiles" / "outlook"),
    "outlook_search": "查价 新人价",
    "outlook_max_mails": 20,
    "final_filter_file": "",
    "final_price_dir": "",
    "final_price_pattern": "*.xlsx",
    "final_price_sku_column": "skuId",
    "final_price_value_column": "价格",
    "final_price_date_label": "",
    "registration_submit_dir": str(PROJECT_ROOT / "data" / "output" / "提报文件"),
    "registration_input_dir": str(PROJECT_ROOT / "data" / "input" / "提报确认表"),
    "registration_base_name": "新人价提报",
    "registration_chunk_size": 5000,
    "registration_submit_pattern": "新人价提报_PART*.xlsx",
    "registration_url": "https://yx.jd.com/user-operate#/customize/newCanSignUpActivity?pathkey=goodspool",
    "registration_profile_dir": str(PROJECT_ROOT / "browser_profiles" / "backend"),
    "registration_month_threshold_day": 25,
    "registration_status_export_dir": str(PROJECT_ROOT / "data" / "output" / "提报情况导出"),
    "registration_status_mail_subject": "商品下载完成通知",
    "registration_status_mail_sender": "增长权益",
    "registration_status_mail_timeout": 1200,
    "registration_status_file": "",
    "registration_status_final_file": "",
    "registration_status_confirm_file": "",
}


@dataclass(frozen=True)
class AppConfig:
    values: dict[str, Any]

    def get(self, key: str, default: Any = None) -> Any:
        return self.values.get(key, default)

    def path(self, key: str) -> Path:
        return Path(str(self.values.get(key, ""))).expanduser()

    @property
    def output_root(self) -> Path:
        return Path(str(self.values["output_root"]))

    @property
    def log_root(self) -> Path:
        return Path(str(self.values.get("log_root") or PROJECT_ROOT / "logs"))

    @property
    def filter_output_dir(self) -> Path:
        return self.output_root / "筛品结果"

    @property
    def split_output_dir(self) -> Path:
        return self.output_root / "查价前sku"

    @property
    def download_output_dir(self) -> Path:
        return self.output_root / "查价结果导出"


def parse_scalar(value: str) -> Any:
    text = value.strip()
    if text == "":
        return ""
    if (text.startswith('"') and text.endswith('"')) or (text.startswith("'") and text.endswith("'")):
        return text[1:-1]
    lowered = text.lower()
    if lowered in {"true", "yes", "y"}:
        return True
    if lowered in {"false", "no", "n"}:
        return False
    try:
        return int(text)
    except ValueError:
        pass
    try:
        return float(text)
    except ValueError:
        return text


def load_simple_yaml(path: Path) -> dict[str, Any]:
    data: dict[str, Any] = {}
    current_key: str | None = None
    for raw_line in path.read_text(encoding="utf-8-sig").splitlines():
        line = raw_line.rstrip()
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            continue
        if stripped.startswith("-") and current_key:
            current_value = data.get(current_key)
            if not isinstance(current_value, list):
                data[current_key] = [] if current_value in (None, "") else [current_value]
            data[current_key].append(parse_scalar(stripped[1:].strip()))
            continue
        if ":" not in line:
            continue
        key, value = line.split(":", 1)
        current_key = key.strip()
        if value.strip() == "":
            data[current_key] = [] if current_key in DEFAULTS and isinstance(DEFAULTS[current_key], list) else ""
        else:
            data[current_key] = parse_scalar(value)
    return data


def load_config(path: Path | None = None) -> AppConfig:
    configured_path = os.environ.get("NEWCOMER_TOOL_CONFIG")
    path = Path(configured_path) if configured_path else (path or CONFIG_PATH)
    values = dict(DEFAULTS)
    if path.exists():
        values.update(load_simple_yaml(path))
    for key, default in DEFAULTS.items():
        if isinstance(default, list) and isinstance(values.get(key), str):
            values[key] = [values[key]] if values[key] else []
    return AppConfig(values)


def print_config(config: AppConfig) -> None:
    print("\n当前配置：")
    important_keys = [
        "batch_name", "batch_date",
        "source_auto_detect", "source_auto_dir", "source_auto_patterns",
        "output_root", "log_root", "date_tag", "sku_column", "keep_mapping_values",
        "keep_pop_shop_keyword", "exclude_category2", "exclude_name_keywords",
        "exclude_erp_accounts",
        "filter_round_mode", "resubmit_final_file", "resubmit_success_keywords", "resubmit_exit_keywords",
        "resubmit_price_review_enabled", "resubmit_price_diff_abs_threshold", "resubmit_price_diff_rate_threshold",
        "split_base_name", "split_chunk_size",
        "final_filter_file", "final_price_dir", "final_price_pattern", "final_price_date_label",
        "registration_input_dir", "registration_submit_dir", "registration_submit_pattern", "registration_month_threshold_day",
    ]
    labels = {
        "source_file": "source_file（旧配置兜底；自动识别到新源表时会被覆盖）",
        "source_auto_detect": "source_auto_detect（自动判断是否有新增大盘源表）",
        "source_auto_dir": "source_auto_dir（自动检测目录）",
        "source_auto_patterns": "source_auto_patterns（自动检测文件名规则）",
        "filter_round_mode": "filter_round_mode（auto=有源表先筛品，无源表走复提）",
        "resubmit_final_file": "resubmit_final_file（空=自动取最新提报整理最终表）",
        "resubmit_success_keywords": "resubmit_success_keywords（成功但需复提关键词）",
        "resubmit_exit_keywords": "resubmit_exit_keywords（采销退出关键词）",
        "resubmit_price_review_enabled": "resubmit_price_review_enabled（复提二次查价价格复核开关）",
        "resubmit_price_diff_abs_threshold": "resubmit_price_diff_abs_threshold（复提查价差值阈值）",
        "resubmit_price_diff_rate_threshold": "resubmit_price_diff_rate_threshold（复提查价差异比例阈值）",
    }
    for key in important_keys:
        print(f"- {labels.get(key, key)}: {config.get(key)}")
    print("\n说明：源表自动检测、提报状态字段、批次配置为常规流程核心；旧版名单配置仅保留兼容，不再要求使用者提供。")
    print("")
