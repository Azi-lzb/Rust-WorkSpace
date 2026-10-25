# 共享能力请求 #7：Umya MoveSheet 丢失 pageSetUpPr

- 提出方：Audit（audit-dev）
- 状态：待 main 评估；Audit 未修改共享 Excel-Core
- 关联：请求 #6 的逐表顺序无关比较已由 `fc6b121` 同步；该修复使完整保留校验继续检查逐表内容。

## 真实复现

Windows 原生将 `AUDIT_S1_HISTORY_REAL_WORKBOOK` 设为用户提供的只读 `config-real` 非空历史工作簿，并运行 `cargo test --offline --locked -p audit-core --lib s1_history::tests::real_s1_history_workbook_engine_candidate_only -- --ignored --exact --nocapture`。样本 5 张表、历史数据 9 行。`UmyaXlsxBackend::apply` 成功，完整计划随后被 `OpenpyxlXlsxValidator` 拒绝：`openpyxl changed worksheet page settings`。

对临时输入副本与候选输出做只读 ZIP/XML 检查：输入的 5 个 worksheet 均含空的 `<sheetPr><pageSetUpPr/></sheetPr>`；Umya 输出的 5 个 worksheet 均缺少该子项。源样本 SHA-256 前后相同。该差异违反当前 PageSettings OOXML 保留快照；未据此断言 Excel 渲染结果不同。

## 请求

请 main 评估 Umya MoveSheet 的序列化是否应保留明确存在的 `pageSetUpPr`（包括无属性的空元素），并以真实模板验证修复。Audit 侧没有设置生产 capability，也未改共享代码。

## 验收边界

完整 EditPlan 必须通过共享保留校验，并确认逐表页面设置、工作表顺序、历史写入、其他工作表和输入不可变。当前失败发生在共享保留校验；S1-F01 历史写回仍为 Unverified。
