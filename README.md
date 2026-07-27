# 新人价自动化工具 MVP

## 如何启动
双击 `运行工具.bat`，按菜单数字执行。

## 推荐验证顺序
1. 选择 `6 查看当前配置`，确认源文件和输出目录正确。
2. 选择 `1 筛品`，先验证筛品结果和日志。
3. 选择 `2 提取并拆分查价前SKU`，验证汇总和 PART 文件。
4. 选择 `3 一键执行本地流程`，验证本地全流程。
5. 选择 `4 后台上传并批量导出到邮箱（实验）`，建议先只跑 `PART01`。
6. 收到邮件后选择 `5 Outlook邮件下载（实验）`。

## 默认输出
- 筛品结果：`data\output\筛品结果\新人价筛品结果_{date_tag}.xlsx`
- 筛品日志：`data\output\筛品结果\新人价筛品日志_{date_tag}.txt`
- 查价前 SKU：`data\output\查价前sku\8月第一批查价前sku.xlsx`
- PART 文件：`data\output\查价前sku\8月第一批查价前sku_PART01.xlsx` 等
- 邮件附件：`data\output\查价结果导出`

## 修改规则
打开 `config.yaml` 修改路径、筛选规则、拆分数量和实验功能参数。

常改项：
- `source_file`：源数据文件。
- `exclude_category2`：剔除二级分类。
- `exclude_name_keywords`：商品名剔除关键词。
- `exclude_erp_accounts`：剔除 ERP 账号。
- `split_chunk_size`：每个 PART 文件记录数。
- `export_csv`：是否额外输出 CSV。

## 注意事项
- 默认只输出筛品 Excel 和筛品日志，不输出 CSV。
- 如果 `current_success_skus_file`、`submitted_skus_file`、`price_column` 为空，相关步骤会在日志中标记为未执行。
- 网页和 Outlook 自动化需要 `playwright`；如果未安装，请先运行：

```bat
python -m pip install playwright
python -m playwright install chromium
```

## 回退
原始脚本已备份在 `scripts_legacy`。
