# 共享能力请求 #7：Umya MoveSheet 丢失 pageSetUpPr

- 提出方：Audit（audit-dev）
- 状态：空 `pageSetUpPr` 默认值归一已由 main 修复并同步；真实候选现推进到 Relationships 检查后仍失败，见共享请求 #8
- 关联：请求 #6 的逐表顺序无关比较已由 `fc6b121` 同步；该修复使完整保留校验继续检查逐表内容。

## 真实复现

Windows 原生将 `AUDIT_S1_HISTORY_REAL_WORKBOOK` 设为用户提供的只读 `config-real` 非空历史工作簿，并运行 `cargo test --offline --locked -p audit-core --lib s1_history::tests::real_s1_history_workbook_engine_candidate_only -- --ignored --exact --nocapture`。样本 5 张表、历史数据 9 行。`UmyaXlsxBackend::apply` 成功，完整计划随后被 `OpenpyxlXlsxValidator` 拒绝：`openpyxl changed worksheet page settings`。

对临时输入副本与候选输出做只读 ZIP/XML 检查：输入的 5 个 worksheet 均含空的 `<sheetPr><pageSetUpPr/></sheetPr>`；Umya 输出的 5 个 worksheet 均缺少该子项。源样本 SHA-256 前后相同。该差异违反当前 PageSettings OOXML 保留快照；未据此断言 Excel 渲染结果不同。

## 请求

请 main 评估 Umya MoveSheet 的序列化是否应保留明确存在的 `pageSetUpPr`（包括无属性的空元素），并以真实模板验证修复。Audit 侧没有设置生产 capability，也未改共享代码。

## 验收边界

完整 EditPlan 必须通过共享保留校验，并确认逐表页面设置、工作表顺序、历史写入、其他工作表和输入不可变。当前失败发生在共享保留校验；S1-F01 历史写回仍为 Unverified。

## 2026-10-25 main 修复后复验

- main 提交 `0bb17c7` 已以 `cc64e15` 同步；校验器仅把无属性且无子元素的空 `pageSetUpPr` 视为默认值归一，显式页面设置仍严格比较。
- 同一真实候选已通过 PageSettings 检查，但继续失败于 `openpyxl changed OOXML relationships`。观察到工作簿新增 `sharedStrings` 关系；被更新的管理历史表原有 9 个外链关系、输出为 0；批注/VML 关系 ID 也被重新编号。Python `write_history_xlsx` 会删除并重建该管理表，因此目标表关系移除可能是预期业务语义；需按编辑计划有限处理关系差异，不能全局放宽。
- 源样本保持不变。完整计划仍未通过；S1 历史写回仍 Unverified。详见共享请求 #8。
