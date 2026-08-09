from __future__ import annotations

import json
import os
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import streamlit as st

from newcomer_tool.batch import batch_date, batch_name, load_manifest, manifest_path
from newcomer_tool.config import PROJECT_ROOT, AppConfig, load_config
from newcomer_tool.gui_actions import run_business_confirmation_action, run_final_pricing_action


DEFAULT_CONFIG_PATH = PROJECT_ROOT / "config.test.yaml"


@dataclass(frozen=True)
class GuiContext:
    config: AppConfig
    config_path: Path
    environment: str
    batch_name: str
    batch_date: str
    output_root: Path
    log_root: Path
    manifest_path: Path


def readable_path(path: Path | str | None) -> str:
    return str(path) if path else "未找到"


def latest_file(directory: Path, patterns: list[str]) -> Path | None:
    if not directory.exists():
        return None
    candidates: list[Path] = []
    for pattern in patterns:
        candidates.extend(path for path in directory.glob(pattern) if path.is_file() and not path.name.startswith("~$"))
    return sorted(candidates, key=lambda item: item.stat().st_mtime, reverse=True)[0] if candidates else None


def list_files(directory: Path, patterns: list[str], limit: int = 6) -> list[Path]:
    if not directory.exists():
        return []
    files: list[Path] = []
    seen: set[Path] = set()
    for pattern in patterns:
        for path in directory.glob(pattern):
            if not path.is_file() or path.name.startswith("~$"):
                continue
            resolved = path.resolve()
            if resolved in seen:
                continue
            seen.add(resolved)
            files.append(path)
    return sorted(files, key=lambda item: item.stat().st_mtime, reverse=True)[:limit]


@st.cache_data(ttl=5)
def load_gui_context() -> GuiContext:
    config_path = DEFAULT_CONFIG_PATH if DEFAULT_CONFIG_PATH.exists() else PROJECT_ROOT / "config.yaml"
    config = load_config(config_path)
    environment = "测试环境" if config_path.name == "config.test.yaml" else "正式环境"
    return GuiContext(
        config=config,
        config_path=config_path,
        environment=environment,
        batch_name=batch_name(config),
        batch_date=batch_date(config),
        output_root=config.output_root,
        log_root=config.log_root,
        manifest_path=manifest_path(config),
    )


@st.cache_data(ttl=5)
def load_batch_status(config_values: dict[str, Any]) -> dict[str, Any]:
    config = AppConfig(config_values)
    path = manifest_path(config)
    manifest = load_manifest(config)
    steps = manifest.get("steps", {}) if path.exists() else {}
    completed_steps = list(steps.keys())
    return {
        "manifest_exists": path.exists(),
        "manifest_path": path,
        "completed_steps": completed_steps,
        "steps": steps,
        "next_action": infer_next_action(completed_steps),
    }


@st.cache_data(ttl=5)
def scan_latest_outputs(config_values: dict[str, Any]) -> dict[str, Any]:
    config = AppConfig(config_values)
    output_root = config.output_root
    final_dir = output_root / "最终结果"
    submit_dir = Path(str(config.get("registration_submit_dir") or output_root / "提报文件"))
    download_dir = config.download_output_dir
    split_dir = config.split_output_dir
    status_export_dir = Path(str(config.get("registration_status_export_dir") or output_root / "提报情况导出"))
    suspicious_dir = download_dir / "疑似误下载"
    return {
        "final_pricing_file": latest_file(final_dir, ["*筛品查价表*.xlsx"]),
        "registration_status_file": latest_file(final_dir, ["*提报情况整理表*.xlsx", "新人价提报整理_*.xlsx"]),
        "price_download_dir": download_dir,
        "price_download_files": list_files(download_dir, ["*.xlsx"], limit=5),
        "suspicious_dir": suspicious_dir,
        "suspicious_files": list_files(suspicious_dir, ["*.xlsx"], limit=5),
        "split_dir": split_dir,
        "split_files": list_files(split_dir, ["*PART*.xlsx"], limit=5),
        "submit_dir": submit_dir,
        "submit_files": list_files(submit_dir, ["*PART*.xlsx"], limit=5),
        "status_export_dir": status_export_dir,
        "status_export_file": latest_file(status_export_dir, ["*.xlsx"]),
        "log_dir": config.log_root,
        "log_files": list_files(config.log_root, ["*.txt", "*.log"], limit=6),
    }


def infer_next_action(completed_steps: list[str]) -> str:
    step_set = set(completed_steps)
    if "阶段4-审核整理" in step_set:
        return "本批次已完成审核整理，可查看提报情况整理表，或准备下一轮复提。"
    if "阶段2-业务确认生成提报" in step_set:
        return "已生成提报文件，下一步建议通过原命令行工具执行后台提报。"
    if "阶段1-筛品查价整合" in step_set:
        return "已生成筛品查价表，下一步可以上传业务确认表并生成提报文件。"
    if "阶段1-Outlook查价下载" in step_set:
        return "已下载查价结果，推荐点击“整合查价结果，生成新人价表”。"
    if "阶段1-筛品拆分" in step_set:
        return "已生成查价 SKU 文件，下一步通过原命令行工具完成后台查价和 Outlook 下载。"
    return "建议先确认配置和输入文件，再通过原命令行工具开始筛品查价。"


def stage_cards(completed_steps: list[str], outputs: dict[str, Any]) -> list[dict[str, str]]:
    step_set = set(completed_steps)
    final_pricing_found = bool(outputs.get("final_pricing_file"))
    submit_found = bool(outputs.get("submit_files"))
    status_found = bool(outputs.get("registration_status_file"))
    return [
        {"name": "准备任务", "status": "已完成" if completed_steps else "可继续", "description": "确认批次、配置和工作目录。", "output": "批次清单" if completed_steps else "等待开始"},
        {"name": "筛品查价", "status": "已完成" if "阶段1-Outlook查价下载" in step_set or final_pricing_found else "需原工具", "description": "筛品、拆分 SKU、后台查价和下载结果。", "output": "查价文件" if outputs.get("price_download_files") else "待查价"},
        {"name": "生成新人价表", "status": "已完成" if "阶段1-筛品查价整合" in step_set or final_pricing_found else "可执行", "description": "整合查价结果，生成/更新筛品查价表。", "output": "筛品查价表" if final_pricing_found else "待生成"},
        {"name": "业务确认与提报文件", "status": "已完成" if "阶段2-业务确认生成提报" in step_set or submit_found else "需人工", "description": "读取确认表，剔除不报名 SKU，生成提报文件。", "output": "提报 PART" if submit_found else "等待确认表"},
        {"name": "审核回填", "status": "已完成" if "阶段4-审核整理" in step_set or status_found else "需原工具", "description": "整理后台结果，沉淀下一轮复提依据。", "output": "提报情况整理表" if status_found else "待整理"},
    ]


def status_badge(status: str) -> str:
    colors = {"已完成": "#16a34a", "可执行": "#2563eb", "可继续": "#2563eb", "需人工": "#d97706", "需原工具": "#7c3aed", "未开始": "#6b7280", "异常": "#dc2626"}
    color = colors.get(status, "#6b7280")
    return f"<span class='status-badge' style='background:{color}'>{status}</span>"


def inject_style() -> None:
    st.markdown(
        """
        <style>
        .block-container { padding-top: 1.5rem; }
        .hero { background: linear-gradient(135deg, #eff6ff 0%, #f8fafc 60%, #fff7ed 100%); border: 1px solid #dbeafe; border-radius: 22px; padding: 24px; margin-bottom: 18px; }
        .hero-title { font-size: 32px; font-weight: 800; color: #0f172a; margin-bottom: 6px; }
        .hero-sub { color: #475569; font-size: 15px; }
        .env-pill { display:inline-block; padding:5px 12px; border-radius:999px; background:#dbeafe; color:#1d4ed8; font-weight:700; font-size:13px; margin-left:8px; }
        .section-card { border:1px solid #e5e7eb; border-radius:18px; padding:18px; background:white; box-shadow: 0 3px 14px rgba(15,23,42,0.05); min-height: 190px; }
        .status-badge { color:white; padding:4px 10px; border-radius:999px; font-size:13px; font-weight:700; }
        .muted { color:#64748b; font-size:14px; }
        .action-note { background:#f8fafc; border-left:4px solid #2563eb; padding:12px 14px; border-radius:12px; color:#334155; }
        .danger-note { background:#fff7ed; border-left:4px solid #f97316; padding:12px 14px; border-radius:12px; color:#9a3412; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_file_item(label: str, path: Path | None) -> None:
    exists = bool(path and path.exists())
    icon = "✅" if exists else "⚪"
    st.markdown(f"**{icon} {label}**")
    st.code(readable_path(path), language=None)


def render_file_list(label: str, files: list[Path]) -> None:
    st.markdown(f"**{label}**")
    if not files:
        st.caption("暂未找到本类文件。")
        return
    for file in files:
        st.caption(str(file))


def open_path(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)
    os.startfile(str(path))  # type: ignore[attr-defined]


def start_bat(path: Path) -> None:
    subprocess.Popen(["cmd", "/c", "start", "", str(path)], cwd=str(PROJECT_ROOT))


def render_action_result(result: dict[str, Any]) -> None:
    if result.get("ok"):
        st.success(result.get("message", "执行完成。"))
    else:
        st.error(result.get("message", "执行失败。"))
        st.code(result.get("error", "未知错误"), language=None)
        return
    for warning in result.get("warnings") or []:
        st.warning(warning)
    metrics = result.get("metrics") or {}
    if metrics:
        metric_cols = st.columns(min(len(metrics), 4))
        for index, (label, value) in enumerate(metrics.items()):
            metric_cols[index % len(metric_cols)].metric(label, value)
    outputs = result.get("outputs") or {}
    if outputs:
        st.markdown("**本次输出**")
        for label, value in outputs.items():
            if isinstance(value, list):
                st.markdown(f"- **{label}**")
                for item in value[:8]:
                    st.caption(str(item))
            else:
                st.markdown(f"- **{label}**：`{value}`")


def render_primary_actions(context: GuiContext) -> None:
    st.subheader("推荐操作")
    st.markdown(
        """
        <div class='action-note'>
        <b>按钮说明：</b>蓝色按钮会真实处理文件；灰色按钮只用于查看；后台查价、真实提报、审核导出暂时通过原命令行工具完成。
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.write("")
    action_col1, action_col2 = st.columns([1, 1])
    with action_col1:
        st.markdown("### 生成新人价表")
        st.caption("适用于：已经完成后台查价，并且 Outlook 已下载查价结果。")
        if st.button("整合查价结果，生成新人价表", type="primary", use_container_width=True):
            with st.spinner("正在整合查价结果并生成新人价表..."):
                result = run_final_pricing_action(context.config)
            st.session_state["last_action_result"] = result
            st.cache_data.clear()
            st.rerun()
    with action_col2:
        st.markdown("### 处理业务确认")
        st.caption("适用于：业务已反馈确认表，需要剔除不报名 SKU 并生成提报文件。")
        confirm_file = st.text_input("业务确认表完整路径", placeholder="例如：D:\\...\\业务已确认表.xlsx")
        if st.button("读取业务确认表，生成提报文件", type="primary", use_container_width=True):
            if not confirm_file.strip():
                st.session_state["last_action_result"] = {"ok": False, "message": "业务确认处理失败。", "error": "请先粘贴业务确认表完整路径。", "outputs": {}, "metrics": {}, "warnings": []}
            else:
                with st.spinner("正在读取业务确认表并生成提报文件..."):
                    result = run_business_confirmation_action(context.config, confirm_file)
                st.session_state["last_action_result"] = result
                st.cache_data.clear()
            st.rerun()
    if "last_action_result" in st.session_state:
        with st.container(border=True):
            st.markdown("### 最近一次执行结果")
            render_action_result(st.session_state["last_action_result"])


def render_helper_actions(context: GuiContext) -> None:
    st.subheader("辅助操作")
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("刷新状态", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
    with col2:
        if st.button("打开输出目录", use_container_width=True):
            open_path(context.output_root)
            st.toast("已尝试打开输出目录。")
    with col3:
        if st.button("打开日志目录", use_container_width=True):
            open_path(context.log_root)
            st.toast("已尝试打开日志目录。")


def render_advanced_actions() -> None:
    with st.expander("高级操作 / 原命令行入口", expanded=False):
        st.markdown(
            """
            <div class='danger-note'>
            后台查价、真实提报、审核导出仍建议通过原命令行工具执行。这里是兜底入口，不是主流程按钮。
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.write("")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("打开测试命令行工具", use_container_width=True):
                start_bat(PROJECT_ROOT / "run_test.bat")
                st.info("已打开测试入口命令行窗口。")
        with col2:
            if st.button("打开正式命令行工具", use_container_width=True):
                start_bat(PROJECT_ROOT / "run.bat")
                st.warning("已打开正式入口。真实提报相关动作仍需在命令行中人工确认。")


def main() -> None:
    st.set_page_config(page_title="新人价自动化任务工作台", page_icon="📊", layout="wide")
    inject_style()
    context = load_gui_context()
    status = load_batch_status(context.config.values)
    outputs = scan_latest_outputs(context.config.values)

    st.markdown(
        f"""
        <div class='hero'>
            <div class='hero-title'>新人价自动化任务工作台 <span class='env-pill'>{context.environment}</span></div>
            <div class='hero-sub'>当前批次：<b>{context.batch_name}</b>　｜　批次日期：<b>{context.batch_date}</b></div>
            <div style='height:12px'></div>
            <div class='action-note'><b>下一步建议：</b>{status['next_action']}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    top_cols = st.columns(4)
    top_cols[0].metric("批次清单", "已找到" if status["manifest_exists"] else "未找到")
    top_cols[1].metric("已记录阶段", len(status["completed_steps"]))
    top_cols[2].metric("最近查价文件", len(outputs["price_download_files"]))
    top_cols[3].metric("提报 PART", len(outputs["submit_files"]))

    with st.expander("当前工作目录", expanded=False):
        st.code(f"配置文件：{context.config_path}\n输出目录：{context.output_root}\n日志目录：{context.log_root}\n批次清单：{context.manifest_path}", language=None)

    st.subheader("业务阶段")
    cards = stage_cards(status["completed_steps"], outputs)
    stage_cols = st.columns(5)
    for column, card in zip(stage_cols, cards):
        with column:
            st.markdown(
                f"""
                <div class='section-card'>
                    <h4 style='margin:0 0 10px 0'>{card['name']}</h4>
                    {status_badge(card['status'])}
                    <p class='muted' style='margin-top:12px'>{card['description']}</p>
                    <p style='font-size:13px'><b>关键产物：</b>{card['output']}</p>
                </div>
                """,
                unsafe_allow_html=True,
            )

    render_primary_actions(context)

    st.subheader("关键结果")
    file_col1, file_col2, file_col3, file_col4 = st.columns(4)
    with file_col1:
        render_file_item("筛品查价表", outputs["final_pricing_file"])
    with file_col2:
        render_file_item("提报文件目录", outputs["submit_dir"])
    with file_col3:
        render_file_item("提报情况整理表", outputs["registration_status_file"])
    with file_col4:
        render_file_item("日志目录", outputs["log_dir"] if outputs["log_dir"].exists() else None)

    with st.expander("更多文件与排障信息", expanded=False):
        more_cols = st.columns(3)
        with more_cols[0]:
            render_file_item("查价结果导出目录", outputs["price_download_dir"])
            render_file_item("疑似误下载目录", outputs["suspicious_dir"] if outputs["suspicious_dir"].exists() else None)
            render_file_item("批次清单", context.manifest_path if context.manifest_path.exists() else None)
        with more_cols[1]:
            render_file_list("最近查价下载文件", outputs["price_download_files"])
            render_file_list("疑似误下载文件", outputs["suspicious_files"])
        with more_cols[2]:
            render_file_list("最近查价 SKU PART", outputs["split_files"])
            render_file_list("最近提报 PART", outputs["submit_files"])
            render_file_list("最近日志文件", outputs["log_files"])

    with st.expander("批次清单原始记录", expanded=False):
        if status["manifest_exists"]:
            st.json(json.loads(context.manifest_path.read_text(encoding="utf-8")))
        else:
            st.caption("暂未找到本批次文件清单。完成命令行流程后，这里会显示本批次产物记录。")

    render_helper_actions(context)
    render_advanced_actions()
    st.divider()
    st.caption("当前是 GUI MVP：低风险步骤可在页面直接执行；完整后台自动化和真实提报仍通过原命令行工具完成。V1.2 可继续接入正式服务接口、实时进度和异常中心。")


if __name__ == "__main__":
    main()
