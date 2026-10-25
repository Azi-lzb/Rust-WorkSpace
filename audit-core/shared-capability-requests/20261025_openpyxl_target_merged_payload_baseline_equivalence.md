# 共享能力请求：S1 公式复制目标簿合并非锚载荷的基线等价豁免

- 提出方：Audit（audit-dev 分支）
- 日期：2026-10-25
- 状态：待 main 评估；Audit 不自行修改共享 crates
- 测试证据：`audit-core/probe/reports/s1_formula_copy_target_payload_baseline_20261025.json`（本轮新增）；`audit-core/probe/reports/s1_formula_copy_target_merged_payload_guard_20261025.json`；`audit-core/probe/reports/s1_formula_copy_candidate_openpyxl_windows_native_20261025.json`

## 涉及的共享接口

1. `crates/excel-editor` `OpenpyxlXlsxBackend::apply` 写前预检中的 `find_non_anchor_merged_payload`：main 提交 `6ecf941` 已把 ImportRange 源侧检查收窄到 source_sheet/source_range，但目标侧仍全局检查所有工作表；目标簿任一合并不锚载荷都会以 `target_merged_non_anchor_payload` 拒绝。
2. `OpenpyxlXlsxValidator` 的 `validate_preserved_values` / `validate_preserved_formulas`：在 `plan.preserve` 含 `CellValuesAndTypes`/`FormulaText` 时对未被操作覆盖的单元格逐一比对输入与输出，输出为空即报 `unmodified cell value changed`。
3. `ExcelEditPlan::drop_allowed`：现仅支持工作表级、且要求显式 `DeleteSheet` 操作，无法表达"允许在保留工作表上丢弃特定合并位置内容"。

## 当前阻塞

S1-F01 完整 60 操作公式复制计划（4×ImportRange + SetComment + SetWorkbookCalculationProperties）在 7/7 本真实目标簿上写入前被 `target_merged_non_anchor_payload` 拒绝：每本目标簿在 4 个复制区之外有 855 个合并不锚载荷（31 个载荷合并区域）。即使预检放行，读回校验器也会在这 855 个单元格上报 `unmodified cell value changed`，路线仍不可用。

## 新证据（本轮取证，7/7 真实样本）

对每本真实目标簿在临时目录执行两条写穿路径并按共享守卫同一语义清点：

- 原始目标簿：855 个合并不锚载荷，全部位于 4 个复制区之外，0 个在区内；37 个合并区域中 31 个含载荷。
- Python 基线（`base_audit.native.openpyxl_workbook.copy_formula_ranges_from_template`，本质是 openpyxl load + 修改 + save）：写后载荷保留 0/855，新增 0；合并区域 37→37 不变。
- 纯 openpyxl load+save（不执行任何业务操作）：载荷保留 0/855，新增 0；合并区域 37→37 不变。
- 结论：**既有 Python 产品基线本来就不保留这些载荷**；openpyxl worker 的丢弃与基线逐格等价。守卫防御的内容损失与已验收 Python 行为完全相同。
- 模板口径交叉核对：4 个复制区、1,072 个复制公式、55 个区内批注，与既有候选报告一致；7/7 输入未变。

## 期望行为（建议契约，实现方式由 main 决定）

引入计划级显式 opt-in（例如 `ExcelEditPlan` 新字段，或新的 `PreservationRequirement` 作用域语义），仅对声明该选项的计划生效：

1. 目标侧预检：仅拒绝与 ImportRange 目标矩形（按 target_sheet）相交的合并不锚载荷；位于全部操作矩形之外的合并不锚载荷不再作为拒绝原因。
2. 读回校验：对输入中位于合并区域非锚位置、且输出为空的单元格，在合并区域本身不变（`MergedRanges` 校验保持严格）的前提下，不计入 unmodified value/formula 变化。
3. 未声明 opt-in 的计划与其他产品默认行为不变（保持全局严格保护）；`drop_allowed` 既有语义不变。
4. 该 opt-in 仅限既有工作簿整批修改计划按已验证能力登记后使用，不构成生产能力验收。

## 复现方式

- 脚本：`audit-core/tests/compare/probe_s1_target_payload_baseline.py`（本轮新增，只读输入、临时输出、报告仅含聚合计数）。
- 参数：`--source-root`（只读源 Python 项目）、`--template-workbook`（真实命名区域模板，校验区域_001..004 展开 4 区）、`--mapping-config`（任一 config-real 配置的字节级临时 .xlsx 副本；无"模块化功能"表时按两侧内置默认解析，与既有报告口径一致）、`--input-dir`（真实批次目标簿目录）、`--report`。
- 判定：`summary.baseline_retained_zero_payloads_all_samples == true` 且 `summary.loadsave_retained_zero_payloads_all_samples == true` 且 `inputs_unchanged == true`。

## 若不采纳

若 main 评估后认为合并不锚载荷必须严格保留，则 openpyxl worker 对该模板族客观不可路由（载荷在 openpyxl load 阶段即被丢弃，worker 无法挽救），S1-F01 需另择引擎或修订对拍基准；请 main 明确裁决，Audit 据此调整迁移计划。
