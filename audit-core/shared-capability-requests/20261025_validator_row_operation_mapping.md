# 共享能力请求（第四项）：引擎无关校验器的行操作坐标映射

- 提出方：Audit（audit-dev 分支）
- 日期：2026-10-25
- 状态：待 main 评估；Audit 不自行修改共享 crates
- 背景：main 已响应前两项请求（`07196d5` 校验器 MoveSheet 感知、`7ac87d6` umya SetFreezePanes/SetAutoFilter，audit-dev 已同步为 `69dc834`/`2f1a9ac`）。S1 历史 38 操作完整计划现已在真实工作簿上通过 umya apply 与读回内容验收，仅剩校验器最后一步。

## 当前阻塞

真实 config-real 非空历史簿的完整更新计划（MoveSheet + DeleteRows + InsertRows + SetCellValue×22 + SetFreezePanes + SetAutoFilter + SetColumnWidth×13 + SetPageSettings）在 umya apply 全链通过后，`OpenpyxlXlsxValidator::validate` 以 "Rust unmodified-style readback mismatch at 核查表校验结果!M6/L3" 误拒：被删除行（first_row=2, count=9 覆盖 2..=10）内的旧单元格——包括只带样式无值的空格——被当作"未修改"与输出比对，而行删除语义下这些格本就随行移除。`style_is_modified_by_plan` 与 `cell_value_is_modified`/`cell_formula_is_modified` 均无 DeleteRows/InsertRows 分支（值检查在本例侥幸通过是因为删除区内恰无旧值格；下移行的坐标映射同样缺失）。

## 期望行为

`OpenpyxlXlsxValidator` 的值/样式/公式"未修改"读回检查增加行操作感知：

1. 源工作簿中位于 `DeleteRows{sheet, first_row, count}` 覆盖范围内的单元格不参与未修改比对（其移除即计划语义）。
2. 插入行导致源行号与输出行号偏移时，按下移映射比对未修改行（源行 r → 输出行 r'），样式与值同映射。
3. 语义与 `UmyaXlsxValidator` 现有 MoveSheet 顺序期望一致：以"源顺序 + 操作序列"推导期望，而非一律按源坐标直查。

## 测试证据

- 真实候选：`AUDIT_S1_HISTORY_REAL_WORKBOOK` 指向真实 config-real 非空历史簿（5 表、管理页 index 1、9 行）运行 `cargo test -p audit-core --lib real_s1_history_workbook_engine_candidate_only -- --ignored`：umya apply 通过、审计侧读回验收全部通过（管理页置首、新内容就位、其余 4 表逐格一致、输入未变），仅校验器在删除区内 M6/L3 两个带样式空格上报误报（Audit 侧按精确类别容忍并记录 `full_plan_validator=passed_with_deleted_row_style_gap`）。
- 复现命令同上；报告 `audit-core/probe/reports/s1_history_move_sheet_real_candidate_windows_native_20261025.json`（随本轮更新为完整计划通过版本）。

## 若不采纳

S1 历史完整写回保持"apply 通过、校验被误拒"状态；可改走 openpyxl worker 路线（仍缺 DeleteRows/InsertRows）或维持不可路由。生产路由在本请求处理前不变。
