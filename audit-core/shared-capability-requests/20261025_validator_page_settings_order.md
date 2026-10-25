# 共享能力请求（第六项）：OOXML 保留快照逐表比较的 MoveSheet 感知

- 提出方：Audit（audit-dev 分支）
- 日期：2026-10-25
- 状态：待 main 评估；Audit 不自行修改共享 crates
- 背景：请求 #4（e73bc1a）落地后，S1 历史 38 操作完整计划在真实 config-real 非空历史簿上的删除区样式误报已消失；候选在 umya apply 与读回验收全部通过后，仅剩 `OpenpyxlXlsxValidator` 的 page-settings 检查因顺序敏感而拒绝。

## 缺陷

`validate_preserved_ooxml_metadata` 的 PageSettings 检查用
`source.page_settings == output.page_settings`（`Vec<(String, Vec<CanonicalXmlElement>)>`
的顺序敏感相等）。`page_settings` 按 workbook.xml 中的工作表顺序填充；
`MoveSheet` 合法地重排工作表后，输出 Vec 顺序必然不同——即使逐表 pageSetup
内容完全一致也会被拒。

真实证据（config-real 非空历史簿，5 表，MoveSheet 把管理页移到首位）：
- 源顺序：使用说明 → 核查表校验结果 → 本地校验结果 → 业务说明 → 任务说明
- 输出顺序：核查表校验结果 → 使用说明 → 本地校验结果 → 业务说明 → 任务说明
- 逐表 pageSetup 属性字典完全一致（Rust `per_sheet_page_setup` 与 Python
  zipfile+ElementTree 双侧验证均 True）。

同一快照结构中 `header_footer_text`（及任何按工作簿顺序填充的逐表 Vec）存在
同类顺序敏感风险；`sheet_order` 与 dimensions/defined names 已是感知或键控
形态。

## 期望行为

`OpenpyxlXlsxValidator` 在计划含 `MoveSheet` 时（或一律），把
`page_settings`（与同结构的逐表 Vec：`header_footer_text` 等）的相等性比较
改为按工作表名键控、顺序不敏感；`sheet_order` 的既有 move-aware 检查继续
承担顺序正确性。与 07196d5 对 sheet-order 的处理同型。

## 复现方式

`AUDIT_S1_HISTORY_REAL_WORKBOOK` 指向真实 config-real 非空历史簿，跑
`cargo test -p audit-core --lib real_s1_history_workbook_engine_candidate_only
-- --ignored`（audit-dev 6d67fe2+）：umya apply 与读回验收全过，校验器以
"openpyxl changed worksheet page settings" 拒绝；测试内 `per_sheet_page_setup`
证明逐表属性字典一致、仅顺序不同。

## 若不采纳

S1 历史完整计划维持"apply/读回通过、校验器顺序误拒"状态；MoveSheet 路线
的引擎无关验收被此最后一个检查挡住。
