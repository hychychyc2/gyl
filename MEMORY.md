# MEMORY

## 重要教训

### 2026-06-25: .xls 解析必须完全复用 openpyxl 路径的 item 结构

`.xls` 文件的 xlrd fallback 必须放在 `parse_domestic_from_attachment()` 函数内部，
构建的 item 字段要与 openpyxl 路径**完全一致**（不能多字段也不能少字段），
不能让 `generate_xlsm()` 收到不同结构的 item。

**当前能正常工作的代码版本：5462d95**
- 验证通过的场景：5条国内(.xls) + 5条海外，10行xlsm正常打开
- HTML/XML全部通过校验
- 业务逻辑：PO号/编码/型号/数量/价格/供应商/目的地全部正确

### 2026-06-22: 禁止手写 xlsm rebuild

之前多次尝试手写简化版 `generate_xlsm()` 来 re生成采购订单模板，结果每次都出现 XML 格式错误、空行拼接 bug（列名和行号混导致格式损坏）、sharedStrings 索引越界等问题。

**规则：永远不要手写 xlsm 重新生成逻辑。** 如果需要 re生成 xlsm，必须调用 `scripts/po_automation_v2.py` 中原版的 `generate_xlsm()` 函数。该函数经过多次验证是正确且稳定的。

如果需要对特定日期 re生成，使用 `scripts/run_date.py`（通过 monkey-patch date.today() 来指定日期），该脚本会 exec 完整的原版 po_automation_v2.py，保证所有逻辑一致。

### 手写版本曾导致的问题
- 空行生成时 format 参数错位：`<c r="B15"15 s="B"/>` 而不是 `<c r="B15" s="7"/>`
- sharedStrings 索引超界：max_ref >= ss_count
- 国内订单合并逻辑使数量翻倍
- XML not well-formed 错误反复出现

## 项目背景
- 采购订单自动化：每天 17:30 定时运行，从杨娜邮件获取订单，生成 WebADI xlsm
- GitHub: https://github.com/hychychyc2/gyl.git
