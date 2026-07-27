from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any


PROJECT_ROOT = Path(__file__).resolve().parents[2]
CONFIG_PATH = PROJECT_ROOT / "config.yaml"


DEFAULTS: dict[str, Any] = {
    "source_file": "",
    "output_root": str(PROJECT_ROOT / "data" / "output"),
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
    "price_column": "",
    "price_threshold": 200,
    "export_csv": False,
    "export_removed_details": False,
    "debug_mode": False,
    "split_base_name": "8月第一批查价前sku",
    "split_chunk_size": 5000,
    "split_source_sheet": "筛选结果",
    "split_source_column": "SKU(含影)",
    "backend_url": "https://yx.jd.com/user-operate#/customize/findLowestActivity?pathkey=goodspool",
    "backend_profile_dir": str(PROJECT_ROOT / "browser_profiles" / "backend"),
    "backend_wait_after_query": 45,
    "backend_wait_after_export": 10,
    "backend_only_file": "8月第一批查价前sku_PART01.xlsx",
    "outlook_url": "https://outlook.office.com/mail/",
    "outlook_profile_dir": str(PROJECT_ROOT / "browser_profiles" / "outlook"),
    "outlook_search": "查价 新人价",
    "outlook_max_mails": 20,
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
            data.setdefault(current_key, []).append(parse_scalar(stripped[1:].strip()))
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


def load_config(path: Path = CONFIG_PATH) -> AppConfig:
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
        "source_file", "output_root", "date_tag", "sku_column", "keep_mapping_values",
        "keep_pop_shop_keyword", "exclude_category2", "exclude_name_keywords",
        "exclude_erp_accounts", "current_success_skus_file", "submitted_skus_file",
        "price_column", "price_threshold", "split_base_name", "split_chunk_size",
        "export_csv", "debug_mode",
    ]
    for key in important_keys:
        print(f"- {key}: {config.get(key)}")
    print("")
