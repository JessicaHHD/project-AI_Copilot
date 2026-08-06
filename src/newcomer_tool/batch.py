from __future__ import annotations

import json
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any

from .config import AppConfig


def text(value: Any) -> str:
    return str(value or "").strip()


def safe_filename_part(value: str, max_length: int = 90) -> str:
    invalid_chars = '<>:"/\\|?*'
    cleaned = "".join("_" if char in invalid_chars else char for char in text(value)).strip(" ._")
    return (cleaned or "未命名批次")[:max_length]


def batch_name(config: AppConfig) -> str:
    return safe_filename_part(str(config.get("batch_name") or "未命名批次"), 60)


def batch_date(config: AppConfig) -> str:
    configured = text(config.get("batch_date"))
    return configured or datetime.now().strftime("%Y%m%d")


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    for index in range(2, 1000):
        candidate = path.with_name(f"{path.stem}_{index}{path.suffix}")
        if not candidate.exists():
            return candidate
    raise FileExistsError(f"输出文件重名过多，请清理目录后重试：{path.parent}")


def archive_matching_files(directory: Path, patterns: list[str], reason_label: str) -> dict[str, Any]:
    if not directory.exists():
        return {"archived_files": [], "archive_dir": ""}
    matched: list[Path] = []
    seen: set[Path] = set()
    for pattern in patterns:
        for path in directory.glob(pattern):
            if not path.is_file() or path.name.startswith("~$"):
                continue
            resolved = path.resolve()
            if resolved in seen:
                continue
            seen.add(resolved)
            matched.append(path)
    if not matched:
        return {"archived_files": [], "archive_dir": ""}

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    archive_dir = directory / "_archive" / f"{timestamp}_{safe_filename_part(reason_label, 40)}"
    archive_dir.mkdir(parents=True, exist_ok=True)
    archived_files: list[str] = []
    for path in matched:
        target = archive_dir / path.name
        try:
            shutil.move(str(path), str(target))
        except OSError as exc:
            raise RuntimeError(f"旧文件归档失败：{path}。请关闭相关 Excel 文件后重试。") from exc
        archived_files.append(str(target))
    return {"archived_files": archived_files, "archive_dir": str(archive_dir)}


def final_pricing_output_path(config: AppConfig, output_dir: Path, price_label: str) -> Path:
    label = safe_filename_part(price_label, 20)
    return unique_path(output_dir / f"{batch_name(config)}_筛品查价表{label}.xlsx")


def registration_submit_base(config: AppConfig) -> str:
    return f"{batch_name(config)}_提报sku_{batch_date(config)}"


def registration_submit_pattern(config: AppConfig) -> str:
    return f"{registration_submit_base(config)}*_PART*.xlsx"


def registration_status_output_path(config: AppConfig, output_dir: Path, status_file: Path) -> Path:
    export_name = safe_filename_part(status_file.stem, 80)
    return unique_path(output_dir / f"{batch_name(config)}_提报情况整理表{export_name}_{batch_date(config)}.xlsx")


def manifest_dir(config: AppConfig) -> Path:
    return config.output_root / "批次清单"


def manifest_path(config: AppConfig) -> Path:
    return manifest_dir(config) / f"{batch_name(config)}_{batch_date(config)}.json"


def load_manifest(config: AppConfig) -> dict[str, Any]:
    path = manifest_path(config)
    if not path.exists():
        return {"batch_name": batch_name(config), "batch_date": batch_date(config), "steps": {}}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {"batch_name": batch_name(config), "batch_date": batch_date(config), "steps": {}}


def update_manifest(config: AppConfig, step: str, data: dict[str, Any]) -> Path:
    path = manifest_path(config)
    path.parent.mkdir(parents=True, exist_ok=True)
    manifest = load_manifest(config)
    manifest["batch_name"] = batch_name(config)
    manifest["batch_date"] = batch_date(config)
    manifest.setdefault("steps", {})[step] = {
        **data,
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return path
