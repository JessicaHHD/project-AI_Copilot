# 新人价自动化工作台 MVP — Implementation Plan

## Context

The user provided a detailed Chinese-language prototype spec for a desktop operations dashboard ("新人价自动化工作台 MVP"). The goal is a 1440px-wide workbench that wraps a Python automation pipeline into a batch-status UI: current stage, KPI overview, key-file access, and log tail. The current `App.tsx` only renders a dot-grid animation — it will be fully replaced.

## Aesthetic

**Stance:** Functional o下面给你一套可以直接拿去给 **Figma Make** 用的原型规划 + Prompt。目标不是做高保真商业 SaaS，而是做一个**桌面端运营工作台原型**，用于你后续作品集展示、老板汇报、以及指导 Codex/开发 AI 做 GUI。

---

# 一、原型定位

## 原型名称

**新人价自动化工作台 MVP**

## 原型目标

这是一个面向新人价运营的本地工具 GUI 原型，用来把现有 Python 自动化脚本包装成一个**批次状态看板**。

用户打开后能快速知道：

1. 当前处理的是哪一批新人价。
2. 当前处于哪个阶段。
3. 上轮提报结果如何。
4. 本批次关键文件在哪里。
5. 下一步应该做什么。
6. 如果失败，应该去哪里看日志。

## 重要边界

原型不要设计成手机 App，也不要设计成复杂 SaaS 后台。  
它应该是一个 **Windows 桌面端运营工作台**。

当前 MVP 不做：

- 不做完整服务接口
- 不做真实提报按钮
- 不做实时后台查价进度
- 不做复杂配置编辑
- 不做 SKU 明细大表格
- 不做完整异常中心
- 不做自动发邮件

---

# 二、桌面端信息架构

推荐整体布局：

```text
┌──────────────────────────────────────────────────────────────┐
│ 顶部：批次信息 + 环境 + 当前阶段 + 下一步建议                  │
├──────────────────────────────────────────────────────────────┤
│ 品池状态概览：上轮提报总数 / 成功 / 失败 / 可复提 / 不报名       │
│ 失败原因 Top 3                                                 │
├──────────────┬──────────────────────────────┬────────────────┤
│ 左侧流程导航  │ 中间主工作区                   │ 右侧关键文件     │
│ 准备任务      │ 当前阶段详情 / 结果摘要 / 操作   │ 整合表 / 日志    │
│ 筛品查价      │ 下一步建议 / 低风险按钮          │ 提报表 / 整理表  │
│ 业务确认      │                                │ 批次清单         │
│ 后台提报      │                                │                │
│ 审核回填      │                                │                │
├──────────────┴──────────────────────────────┴────────────────┤
│ 底部：最近日志 / 异常提示 / 调试文件入口                       │
└──────────────────────────────────────────────────────────────┘
```

---

# 三、页面模块规划

## 1. 顶部批次状态栏

展示：

| 字段 | 示例 |
|---|---|
| 工具名称 | 新人价自动化工作台 |
| 当前批次 | 8月第1批提报 |
| 当前环境 | 测试环境 |
| 当前阶段 | 筛品查价已完成，待业务确认 |
| 下一步建议 | 请复制 ERP 文本并发送业务确认邮件 |
| 批次清单 | 已找到 |
| 操作 | 刷新状态、打开输出目录、打开日志目录 |

---

## 2. 品池状态概览看板

展示轻量数据，不展示 SKU 明细。

| 指标 | 示例 |
|---|---|
| 上轮提报 SKU | 12,340 |
| 提报成功 | 9,860 |
| 提报失败 | 1,240 |
| 可复提 | 860 |
| 不报名 / 退出 | 240 |

失败原因 Top 3：

| 原因 | SKU 数 |
|---|---:|
| 转化不满足 | 420 |
| 删促 | 260 |
| 活动规则不满足 | 180 |

明细入口：

- 打开提报情况整理表
- 查看失败 SKU 明细
- 查看复提候选

注意：  
首页只显示摘要，不要展开完整 SKU 列表。

---

## 3. 左侧流程导航

5 个阶段：

| 阶段 | 状态示例 |
|---|---|
| 准备任务 | 已完成 |
| 筛品查价 | 已完成 |
| 业务确认 | 需人工 |
| 后台提报 | 未开始 |
| 审核回填 | 未开始 |

状态样式：

| 状态 | 视觉建议 |
|---|---|
| 已完成 | 绿色勾 |
| 当前阶段 | 蓝色高亮 |
| 需人工 | 橙色提示 |
| 异常 | 红色警告 |
| 未开始 | 灰色 |

---

## 4. 中间主工作区

默认展示当前阶段详情。

### 当前阶段：业务确认

展示：

- 阶段说明：业务需要确认哪些 SKU 参与报名。
- 当前状态：待人工发送业务确认邮件。
- 输入材料：
  - 新人价整合表
  - ERP 合并文本
- 用户动作：
  - 复制 ERP 文本
  - 打开新人价整合表
  - 上传业务确认表，MVP 可展示为占位按钮
- 下一步：
  - 收到确认表后，生成提报文件。

---

## 5. 右侧关键文件区

始终展示关键文件入口。

| 文件类型 | 状态 |
|---|---|
| 新人价整合表 | 已找到 |
| 查价失败 SKU | 已找到 |
| ERP 合并文本 | 可复制 |
| 提报 PART 文件目录 | 暂未生成 |
| 提报情况整理表 | 暂未生成 |
| 批次清单 | 已找到 |
| 日志目录 | 可打开 |
| 疑似误下载目录 | 暂无 |

每个文件项建议包含：

- 文件名称
- 状态标签
- 路径摘要
- 打开按钮

---

## 6. 底部日志与异常区

展示：

- 最近日志：`2026-08-07 14:32 查价整合完成，生成新人价整合表`
- 当前异常：暂无异常 / 查价结果数量不一致 / Outlook 未登录
- 调试入口：
  - 打开日志目录
  - 打开调试截图目录
  - 打开疑似误下载目录

---

# 四、原型交互设计

MVP 原型只需要这些交互：

| 交互 | 说明 |
|---|---|
| 点击左侧阶段 | 中间区域切换阶段详情 |
| 点击刷新状态 | 模拟刷新，看板数据不必真实变化 |
| 点击打开文件 | 原型里展示 toast：已打开文件 |
| 点击复制 ERP 文本 | 展示 toast：已复制 ERP 文本 |
| 点击打开日志目录 | 展示 toast：已打开日志目录 |
| 点击查看失败 SKU | 弹出轻量弹窗，展示 Top 3 失败原因和“打开明细表”按钮 |

不要设计真实提报确认弹窗，最多在“后台提报”阶段展示：

> 真实提报为高风险动作，当前 MVP 仅展示状态，不在 GUI 中执行。

---

# 五、给 Figma Make 的完整 Prompt

你可以直接复制下面这段给 Figma Make：

```text
请为我设计一个桌面端 Web 原型，不是手机 App。产品名称是「新人价自动化工作台 MVP」。

这是一个面向新人价运营人员的本地自动化工具 GUI 原型。工具背后已有 Python 自动化脚本，可以完成筛品、查价、Outlook 下载、查价整合、生成提报文件、后台提报、审核回填和复提候选识别。当前 MVP 不替代命令行完整操作，而是做一个「批次状态看板」，帮助运营快速知道当前批次做到哪一步、关键结果文件在哪里、下一步该做什么、出问题去哪看日志。

请设计一个 1440px 宽的桌面端工作台界面，风格专业、清爽、适合运营后台工具，不要做成手机页面，不要做成营销落地页。

整体布局：
1. 顶部是批次状态栏。
2. 顶部下方是品池状态概览看板。
3. 主体区域分为三栏：左侧流程导航，中间当前阶段工作区，右侧关键文件入口。
4. 底部是日志与异常提示区。

页面标题：
新人价自动化工作台

顶部批次状态栏需要展示：
- 当前批次：8月第1批提报
- 当前环境：测试环境
- 当前阶段：筛品查价已完成，待业务确认
- 下一步建议：请复制 ERP 文本并发送业务确认邮件
- 批次清单：已找到
- 按钮：刷新状态、打开输出目录、打开日志目录

品池状态概览看板需要展示 5 个核心指标：
- 上轮提报 SKU：12,340
- 提报成功：9,860
- 提报失败：1,240
- 可复提：860
- 不报名 / 退出：240

同时展示失败原因 Top 3：
- 转化不满足：420 SKU
- 删促：260 SKU
- 活动规则不满足：180 SKU

注意：这里只展示摘要，不展示完整 SKU 明细。提供三个入口按钮：
- 打开提报情况整理表
- 查看失败 SKU 明细
- 查看复提候选

左侧流程导航包含 5 个阶段：
1. 准备任务，状态：已完成
2. 筛品查价，状态：已完成
3. 业务确认，状态：需人工，当前高亮
4. 后台提报，状态：未开始
5. 审核回填，状态：未开始

状态视觉：
- 已完成：绿色勾
- 当前阶段：蓝色高亮
- 需人工：橙色提示
- 未开始：灰色
- 异常：红色

中间主工作区默认展示「业务确认」阶段详情。
需要包含：
- 阶段标题：业务确认
- 状态标签：需人工处理
- 阶段说明：业务需要确认哪些 SKU 参与报名，普通复提无异常时可跳过业务确认；首次提报或价格异常时需要确认。
- 本阶段输入材料：
  1. 新人价整合表：已生成
  2. ERP 合并文本：可复制
- 本阶段用户动作：
  1. 复制 ERP 合并文本
  2. 打开新人价整合表
  3. 上传业务确认表，占位按钮即可
- 下一步提示：收到业务确认表后，工具将识别参与报名、不报名和未确认 SKU，并生成提报文件。

右侧关键文件入口需要展示：
- 新人价整合表：已找到
- 查价失败 SKU：已找到
- ERP 合并文本：可复制
- 提报 PART 文件目录：暂未生成
- 提报情况整理表：暂未生成
- 批次清单：已找到
- 日志目录：可打开
- 疑似误下载目录：暂无

每个文件项需要有：
- 文件名称
- 状态标签
- 简短路径，例如 data/output/最终结果
- 打开按钮或复制按钮

底部日志与异常区展示：
- 最近日志：2026-08-07 14:32 查价整合完成，生成新人价整合表
- 当前异常：暂无异常
- 调试入口：打开日志目录、打开调试截图目录、打开疑似误下载目录

交互要求：
- 点击左侧阶段时，中间区域可以切换内容，但如果复杂，可以先做静态原型。
- 点击复制 ERP 文本时，显示一个 toast：已复制 ERP 文本。
- 点击打开文件或目录时，显示 toast：已打开文件或目录。
- 点击查看失败 SKU 明细时，弹出一个轻量弹窗，展示失败原因 Top 3 和「打开明细表」按钮。
- 不要设计真实提报按钮。
- 不要设计一键完整运行流程按钮。
- 不要设计复杂配置编辑器。
- 不要展示完整 SKU 表格。

视觉风格：
- 桌面端运营后台风格
- 浅色背景
- 卡片式布局
- 蓝色作为主色
- 绿色表示完成
- 橙色表示需人工
- 红色表示异常
- 灰色表示未开始
- 信息密度适中，不要太拥挤
- 适合放入 AI 产品经理秋招作品集

请输出一个完整的桌面端 Dashboard 原型。
```

---

# 六、给 Figma Make 的增强版 Prompt：带接口映射

如果你想让原型和脚本更对齐，可以用这个增强版追加：

```text
请在原型中隐含以下数据来源逻辑，但不要把代码展示给用户：

1. 批次信息来自 load_gui_context()
   - batch_name
   - environment
   - output_root
   - log_root
   - manifest_exists

2. 品池状态来自 latest_registration_summary
   - last_submitted_count
   - success_count
   - failed_count
   - resubmit_candidate_count
   - business_exit_count
   - top_failure_reasons

3. 流程阶段状态来自 build_stage_status()
   - prepare_task
   - pricing
   - business_confirmation
   - registration_submit
   - status_merge

4. 关键文件来自 scan_key_outputs()
   - final_pricing_file
   - failed_sku_file
   - erp_text
   - registration_part_dir
   - status_merge_file
   - manifest_file
   - log_dir
   - suspicious_download_dir

5. 日志和异常来自 scan_recent_logs()
   - latest_log_message
   - current_error
   - debug_dir

原型中每个模块都要对应这些数据来源，确保后续开发可以按这个结构实现。
```

---

# 七、给开发 AI 的轻量接口说明

如果后续你要让 Codex/开发 AI 根据原型做 GUI，可以给它这个接口规划。

## 1. `load_gui_context()`

返回：

```json
{
  "batch_name": "8月第1批提报",
  "environment": "测试环境",
  "batch_date": "2026-08-07",
  "output_root": "data/output",
  "log_root": "logs",
  "manifest_exists": true,
  "current_stage": "业务确认",
  "next_action": "请复制 ERP 文本并发送业务确认邮件"
}
```

## 2. `load_pool_overview()`

返回：

```json
{
  "last_submitted_count": 12340,
  "success_count": 9860,
  "failed_count": 1240,
  "resubmit_candidate_count": 860,
  "business_exit_count": 240,
  "top_failure_reasons": [
    {
      "reason": "转化不满足",
      "sku_count": 420
    },
    {
      "reason": "删促",
      "sku_count": 260
    },
    {
      "reason": "活动规则不满足",
      "sku_count": 180
    }
  ]
}
```

## 3. `build_stage_status()`

返回：

```json
[
  {
    "stage": "准备任务",
    "status": "completed",
    "description": "配置和批次信息已识别"
  },
  {
    "stage": "筛品查价",
    "status": "completed",
    "description": "新人价整合表已生成"
  },
  {
    "stage": "业务确认",
    "status": "need_user_action",
    "description": "请发送业务确认邮件并上传确认表"
  },
  {
    "stage": "后台提报",
    "status": "not_started",
    "description": "等待业务确认后生成提报文件"
  },
  {
    "stage": "审核回填",
    "status": "not_started",
    "description": "提报完成后可导出审核结果"
  }
]
```

## 4. `scan_key_outputs()`

返回：

```json
{
  "final_pricing_file": {
    "label": "新人价整合表",
    "status": "found",
    "path": "data/output/最终结果/8月第1批提报_筛品查价表.xlsx"
  },
  "failed_sku_file": {
    "label": "查价失败 SKU",
    "status": "found",
    "path": "data/output/最终结果/查价失败SKU.xlsx"
  },
  "erp_text": {
    "label": "ERP 合并文本",
    "status": "available",
    "value": "erp1;erp2;erp3"
  },
  "registration_part_dir": {
    "label": "提报 PART 文件目录",
    "status": "missing",
    "path": "data/output/提报文件"
  },
  "status_merge_file": {
    "label": "提报情况整理表",
    "status": "missing",
    "path": "data/output/最终结果"
  },
  "manifest_file": {
    "label": "批次清单",
    "status": "found",
    "path": "data/output/批次清单/manifest.json"
  },
  "log_dir": {
    "label": "日志目录",
    "status": "found",
    "path": "logs"
  },
  "suspicious_download_dir": {
    "label": "疑似误下载目录",
    "status": "empty",
    "path": "data/output/查价结果导出/疑似误下载"
  }
}
```

## 5. `load_recent_log_status()`

返回：

```json
{
  "latest_log_message": "2026-08-07 14:32 查价整合完成，生成新人价整合表",
  "current_error": null,
  "debug_dirs": [
    "logs/debug_backend",
    "logs/debug_registration_submit"
  ]
}
```

---

# 八、原型页面里的信息量控制原则

你可以给 Figma Make 加这段，防止它做得太复杂：

```text
请控制信息层级：
1. 首页只展示摘要数字，不展示 SKU 明细。
2. 品池状态只展示 5 个核心指标和 Top 3 失败原因。
3. SKU 明细通过“打开整理表”或“查看失败 SKU 明细”进入，不在首页展开。
4. 业务确认阶段只展示当前要做什么，不展示复杂规则配置。
5. 高风险动作只展示状态，不提供真实提交按钮。
6. 页面最多使用 3 栏布局，不要堆太多卡片。
```

---

# 九、如果 Figma Make 生成得太花，怎么纠偏

你可以继续追问它：

```text
请降低视觉复杂度，改成更像 Windows 桌面端运营工具，而不是 SaaS 营销后台。减少渐变、减少大面积插画，突出批次状态、流程阶段、关键文件和日志入口。
```

或者：

```text
请把页面信息密度降低，保留顶部批次栏、品池状态概览、左侧流程导航、中间阶段详情、右侧关键文件、底部日志提示，不要增加其他模块。
```

---

# 十、最终建议

你这次原型的核心不是“好看”，而是证明一个产品判断：

> 新人价自动化工具不应该先做成一堆按钮，而应该先做成一个批次工作台，让用户知道当前品池状态、流程阶段、结果文件和下一步动作。

Figma Make 的任务就是把这个判断视觉化。  
后续 Codex/开发 AI 再根据这套原型去做 Streamlit 或桌面 GUI，才不会和脚本能力脱节。ps workbench — clean Windows-style admin tool, not SaaS marketing.  
**Fonts:** `Noto Sans SC` (Chinese + Latin body) + `JetBrains Mono` (file paths, log lines, numeric labels).  
**Palette:** Near-white ground `#F5F7FA`, blue primary `#1A6CF6`, semantic green `#22C55E` / orange `#F59E0B` / red `#EF4444`, mid-gray borders.

## Layout (CSS Grid)

```
┌──── header (batch status bar) ────────────────────────────────────────┐
├──── pool overview (KPI row + failure top-3 + action buttons) ─────────┤
├── left nav (220px) ──┬── main workspace (flex-1) ──┬── right panel (260px) ─┤
├──── footer (log + exception + debug links) ────────────────────────────┤
```

## Files to Modify / Create

| File | Action |
|---|---|
| `src/index.css` | Add Google Fonts `@import` for Noto Sans SC & JetBrains Mono |
| `src/App.tsx` | Full replacement — main layout shell + all inline sub-components |

No new files are needed; all components live in `App.tsx` as local functions to keep scope tight for a prototype.

## Data Model (mock, in-module constants)

Match the interface spec from the attachment exactly:
- `guiContext` — batch name, environment, current stage, next action, manifest exists
- `poolOverview` — 5 KPIs + top-3 failure reasons
- `stageList` — 5 stages with status `completed | current | need_user | not_started`
- `keyFiles` — 8 file entries with status `found | available | missing | empty`
- `logStatus` — latest log message, current error (null), debug dirs

## Components (all in App.tsx)

1. **`TopBar`** — batch badge, environment tag, stage label, next-action text, three action buttons (刷新状态 / 打开输出目录 / 打开日志目录)
2. **`PoolOverview`** — 5 stat tiles in a row, failure Top-3 table, 3 entry buttons (打开整理表 / 查看失败SKU明细 / 查看复提候选)
3. **`StageNav`** — vertical list of 5 stages; click switches active stage; status icons (✓ green / dot blue / ⚠ orange / – gray)
4. **`WorkspacePanel`** — renders detail view for the selected stage; each stage has title, status badge, description, input materials list, action buttons; "后台提报" shows a disabled warning state instead of a real submit button
5. **`FilePanel`** — 8 file rows each with name, status badge, path snippet, and open/copy button
6. **`Footer`** — log line, error indicator, 3 debug buttons
7. **`Toast`** — small bottom-right notification; auto-dismisses after 2.5 s; managed by a single `useState`
8. **`FailureModal`** — lightweight overlay with Top-3 failure table and "打开明细表" button; triggered by pool overview

## Interactions

| Trigger | Effect |
|---|---|
| Click stage in nav | `setActiveStage(id)` → WorkspacePanel rerenders |
| Click any file / dir open button | `showToast('已打开文件/目录')` |
| Click "复制 ERP 文本" | `showToast('已复制 ERP 文本')` |
| Click "查看失败 SKU 明细" | `setModalOpen(true)` |
| Click "刷新状态" | brief spinner (300 ms) then `showToast('状态已刷新')` |

## Verification

1. The Vite dev server hot-reloads on save — confirm the layout renders at full width without horizontal scroll.
2. Click each of the 5 stage nav items and verify WorkspacePanel updates.
3. Trigger each toast and verify it auto-dismisses.
4. Open the failure modal and close it.
5. Verify no build errors in the terminal.
