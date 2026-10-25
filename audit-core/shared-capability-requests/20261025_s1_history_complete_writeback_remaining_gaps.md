# 共享能力请求补充：S1 历史工作簿完整写回的两项剩余缺口

- 提出方：Audit（audit-dev 分支）
- 日期：2026-10-25
- 状态：待 main 评估；Audit 不自行修改共享 crates
- 背景：main 提交 `bbd922d`（MoveSheet）已由 audit-dev 同步为 `db5f7e1`，S1 历史计划已改为以 `MoveSheet{new_index=0}` 置首；真实 config-real 非空历史工作簿上的 MoveSheet-only 候选已通过（见 `audit-core/probe/reports/s1_history_move_sheet_real_candidate_windows_native_20261025.json`）。
- 测试证据：上述报告；合成聚焦候选（MoveSheet+行替换+页面设置）在 `audit-core/src/s1_history.rs` 的 ignored 测试中记录。

## 缺口一：umya 后端缺 SetFreezePanes/SetAutoFilter（阻塞完整计划）

- 接口：`crates/excel-editor` `UmyaXlsxBackend`；完整 S1 历史计划含 `SetFreezePanes` 与 `SetAutoFilter`。
- 当前阻塞：完整 38 操作计划在真实工作簿上于 `apply SetFreezePanes` 处被拒（"umya backend does not implement freeze-pane or AutoFilter edits"）。此外删除历史行会把该页 `_xlnm._FilterDatabase` 从 `$A$1:$L$10` 收缩到 `$A$1:$L$1`，只有 `SetAutoFilter` 能按 Python 基线语义恢复新范围。
- 期望行为：umya 后端实现 `SetFreezePanes`（含清除）与 `SetAutoFilter`（含清除），并把行删除对 `_FilterDatabase` 的收缩纳入行操作语义或由随后的 `SetAutoFilter` 覆盖；读回校验按计划断言最终冻结窗格与筛选范围。
- 若不采纳：S1 历史完整写回需改走 openpyxl worker 路线（该 worker 现缺 DeleteRows/InsertRows，为对应缺口）或维持不可路由。

## 缺口二：引擎无关读回校验器不识别 MoveSheet

- 接口：`crates/excel-editor` `OpenpyxlXlsxValidator::validate` 内 `validate_sheet_names`——期望顺序计算覆盖 Rename/Create/CreateSheetAt/DeleteSheet，但没有 `MoveSheet` 分支。
- 当前阻塞：任何经 `OpenpyxlXlsxValidator` 校验的含 MoveSheet 计划都会以"sheet identity/order mismatch"误拒（实际输出顺序正确）。main 的 move_sheet 测试使用 `UmyaXlsxValidator`（含 MoveSheet 语义）故未暴露。
- 期望行为：`validate_sheet_names` 按源顺序应用 `MoveSheet` 序列（与 `UmyaXlsxValidator` 相同语义），使引擎无关读回校验支持 MoveSheet 计划。
- 附带发现：defined-name 比较按原文比较 `raw_reference`；umya 对非 ASCII 表名重发时会加引号（`'表'!$A$1` vs `表'!$A$1` 语义等价）。真实工作簿自带引号引用故未触发，但合成夹具会触发；建议该比较对引用做引号归一化或对表名引号差异豁免。

## 复现方式

- 真实候选：`cargo test -p audit-core --lib real_s1_history_workbook_engine_candidate_only -- --ignored --nocapture`，`AUDIT_S1_HISTORY_REAL_WORKBOOK` 指向只读真实样本（config-real 非空历史簿）。
- 合成候选：`cargo test -p audit-core --lib history_workbook_umya_move_sheet_candidate -- --ignored --nocapture`。
- 引擎无关校验器复现：任一含 `MoveSheet` 的计划经 `OpenpyxlXlsxValidator::validate` 即报顺序不符。
