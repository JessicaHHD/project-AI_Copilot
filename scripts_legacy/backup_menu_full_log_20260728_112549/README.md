# 新人价自动化工具 MVP

## 如何启动
双击 运行工具.bat，按菜单数字执行。

## 推荐验证顺序
1. 选择 6 查看当前配置，确认源文件和输出目录正确。
2. 选择 1 筛品，验证筛品结果、SKU文件和审核报告。
3. 选择 2 提取并拆分查价前SKU，验证汇总和 PART 文件。
4. 选择 3 一键执行本地流程，验证本地全流程。
5. 选择 4 后台上传并批量导出到邮箱，建议先只跑 PART01。
6. 收到邮件后选择 5 Outlook邮件下载。

## 筛品模块 v1 接口

输入：固定格式和字段的源数据 CSV/Excel 文件；可选补充名单 current_success_skus_file 和 submitted_skus_file。

输出：
- data/output/筛品结果/新人价筛品结果_{date_tag}.xlsx，包含 筛选结果、筛品日志 两个 sheet。
- data/output/筛品结果/筛品结果SKU_{date_tag}.xlsx，只含 sku 列，供下一模块使用。
- data/output/筛品结果/新人价筛品日志_{date_tag}.txt。
- data/output/筛品结果/筛品审核报告_{date_tag}.txt，单独保存自动审核结果。
- 可选 CSV，仅当 export_csv: true 时生成。

字段缺失处理：
- 关键字段缺失：终止运行，不生成正式筛品结果。
- 可选名单未提供：对应步骤跳过，审核报告标记为预警。

审核状态：
- PASS：无异常。
- WARN：有预警，可继续但建议查看审核报告。
- FAIL：关键失败，不建议进入下一模块。

关键字段：
- SKU(含影)
- sku名称 或 商品名称
- 1自营3POP_映射
- 店铺名称
- 商品二级分类名称
- 销售员ERP帐号
- 去重计数_用户PIN加密 或 去重计数*用户PIN加密
- 求和_成交金额 或 求和*成交金额
- 去重计数_销售订单编号 或 去重计数*销售订单编号
- 求和_销售数量 或 求和*销售数量

## 修改规则
打开 config.yaml 修改路径、筛选规则、拆分数量和实验功能参数。

常改项：
- source_file：源数据文件。
- current_success_skus_file：目前提报成功 SKU 名单，可为空。
- submitted_skus_file：已经提报过 SKU 名单，可为空。
- exclude_category2：剔除二级分类。
- exclude_name_keywords：商品名剔除关键词。
- exclude_erp_accounts：剔除 ERP 账号。
- split_chunk_size：每个 PART 文件记录数。
- export_csv：是否额外输出 CSV。

## 注意事项
默认不强制人工逐步确认；请通过单独生成的 筛品审核报告_{date_tag}.txt 复核。状态为 WARN 时建议查看审核报告后再进入下一模块。状态为 FAIL 时请先修复源文件或配置。

网页和 Outlook 自动化需要 playwright。如果未安装，请先运行：python -m pip install playwright，然后运行：python -m playwright install chromium。

## 回退
更新脚本会把被覆盖的文件备份到 scripts_legacy/backup_filter_v1_时间戳。
