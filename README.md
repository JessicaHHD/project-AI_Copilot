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
7. 下载完成后选择 7 查价结果整合并生成新人价。
8. 提报前选择 8 提报/审核整理，先从采销确认表生成提报文件，再用演练模式验证页面路径。

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

菜单会直接显示完整筛品日志，避免只看摘要时不清楚每一步含义。

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

## 最终整合模块

菜单 7 会自动读取最新的筛品结果和查价结果导出目录，生成新的整合 Excel，不覆盖原筛品结果。

输出：
- data/output/最终结果/新人价最终结果_{date_tag}_{查价日期}.xlsx。
- Sheet 筛选结果：原筛品结果追加“xx.xx_查价”和“新人价”，并剔除新人价大于等于 200 元的商品。
- Sheet 查价结果合并：合并所有查价导出 xlsx 明细。
- Sheet ERP汇总：最终保留商品的销售员ERP去重列表，以及英文分号连接后的整段文本。
- Sheet 整合日志：先展示筛品模块日志，再记录匹配、未匹配和剔除数量。

新人价规则：自营 = 100天最低价 - 8；非自营 = 100天最低价 - 6.5；若结果小于等于 0，则填 0.01。

## 提报安全测试模块

菜单 8 包含两个步骤：先从采销确认表生成提报 PART 文件，再进入新人价补贴报名后台处理提报 PART 文件。默认推荐演练模式，不上传、不提交。

提报确认表：
- 建议放在 data/input/提报确认表。
- 工具会自动识别 SKU、期望新人价/新人价、是否退出/是否报名 字段。
- 如果字段是“是否退出”：空白默认参加；填写 是、退出、不参加、不报名、不提报、取消、放弃 等字样才剔除。
- 如果字段是“是否报名/是否登记”：空白默认参加；填写 否、不参加、不报名、不提报、退出、取消、放弃 等字样才剔除。
- 输出到 data/output/提报文件，列名为 skuid、促销价，每个 PART 最多 5000 条。

运行模式：
- 演练到上传页：只打开后台、选择首单新人价计划、进入批量上传页，不选择文件、不点击提交。
- 上传文件但不提交：选择文件上传到页面，方便人工确认文件识别情况，不点击提交。
- 真实提交：会点击后台“提交”，必须输入“确认提报”才会继续。

安全限制：
- 菜单默认建议只跑 1 个文件。
- 全量执行需要额外输入“全量提报”。
- 脚本找不到明确按钮时会停止并保存调试截图，不会模糊点击提交。
- 日志保存到 logs/新人价提报自动上传日志.txt。

常改项：
- source_file：源数据文件。
- current_success_skus_file：目前提报成功 SKU 名单，可为空。
- submitted_skus_file：已经提报过 SKU 名单，可为空。
- exclude_category2：剔除二级分类。
- exclude_name_keywords：商品名剔除关键词。
- exclude_erp_accounts：剔除 ERP 账号。
- split_chunk_size：每个 PART 文件记录数。
- final_filter_file：指定筛品结果文件，可为空，默认自动取最新。
- final_price_dir：指定查价结果目录，可为空，默认使用 data/output/查价结果导出。
- final_price_date_label：指定查价日期标签，可为空，默认使用当天日期，如 7.28。
- registration_input_dir：采销确认表建议存放目录，默认 data/input/提报确认表。
- registration_submit_dir：工具生成的提报 PART 文件目录，默认 data/output/提报文件。
- registration_chunk_size：每个提报 PART 文件记录数，默认 5000。
- registration_submit_pattern：提报 PART 文件匹配规则，默认 新人价提报_PART*.xlsx。
- registration_month_threshold_day：月份选择阈值，默认 25 号及以后优先下月计划。
- export_csv：是否额外输出 CSV。

## 注意事项
默认不强制人工逐步确认；请通过单独生成的 筛品审核报告_{date_tag}.txt 复核。状态为 WARN 时建议查看审核报告后再进入下一模块。状态为 FAIL 时请先修复源文件或配置。

网页和 Outlook 自动化需要 playwright。如果未安装，请先运行：python -m pip install playwright，然后运行：python -m playwright install chromium。

## 回退
更新脚本会把被覆盖的文件备份到 scripts_legacy/backup_filter_v1_时间戳。
