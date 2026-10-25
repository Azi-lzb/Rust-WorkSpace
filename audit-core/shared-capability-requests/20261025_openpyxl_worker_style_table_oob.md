# 共享能力请求（第五项）：openpyxl worker 输出样式表越界引用

- 提出方：Audit（audit-dev 分支）
- 日期：2026-10-25
- 状态：待 main 评估；Audit 不自行修改共享 crates
- 背景：main 已落地请求 #3（becd3cd 载荷基线等价豁免）与 #4（e73bc1a 校验器行操作映射），audit-dev 已同步为 d5350f4/1d7c22c。S1 公式复制业务计划按请求 #3 的裁决对目标簿全部表声明 `MergedNonAnchorPayloads` 豁免后，openpyxl worker 首次写穿完整 60 操作计划。

## 缺陷

同一输入、同一 openpyxl 3.1.5：

- **Python 基线**（`copy_formula_ranges_from_template`，openpyxl load+save）：sheet7 `CJ2` 写出 `s=24`，cellXfs 共 38 个，**引用在界内**，xf24 含 `<alignment vertical="center"/>`——语义保留。
- **worker 路线**（openpyxl_worker.py ImportRange 组件样式赋值 + SetComment×55 + SetWorkbookCalculationProperties）：sheet7 `CJ2` 写出 `s=24` 但 cellXfs 仅 **24 个（0..23）**——**越界引用**。Excel/umya/openpyxl 读回回退默认样式，源 xf 0 携带的 `<alignment vertical="center"/>` 语义丢失（源 cellXfs=23、xf0 非默认：带 vertical=center）。

读回校验器在 sheet7 CJ2 报 `unmodified_style_semantic_mismatch`（alignment vertical center → 默认），worker 输出因此不可验收。值/公式在该格两侧完全一致——缺陷只在样式表完整性。

## 涉及的共享接口

`crates/excel-editor/python/openpyxl_worker.py`：ImportRange 的 `_copy_style` 做格间组件赋值（dst.font/fill/border/alignment/number_format/protection = ...）。该赋值模式触发 openpyxl 样式注册表重建；重建后的 cellXfs 数量少于部分"未触碰"格（如源 xf 0 非默认的表）所引用的最高样式索引。

## 期望行为（方向由 main 裁决）

worker 输出中所有格的样式索引必须落在 cellXfs 范围内，且未触碰格的语义与源一致（至少：不产生越界引用；理想：越界格回写源 xf 语义或保留源样式表条目）。可选方向包括：保存前校验最高引用索引与 cellXfs 数量并在不一致时重建/修复映射；或改用不触发注册表重建的样式赋值方式。

## 复现方式

- 命令：`AUDIT_S1_COPY_CANDIDATE_*` 环境变量指向真实模板/目标簿，`AUDIT_S1_COPY_CANDIDATE_ENGINE=openpyxl-worker` 跑 `cargo test -p audit-core --test s1_formula_copy_candidate -- --ignored`（audit-dev d5350f4 + 业务计划豁免声明之后）。
- 证据：`audit-core/probe/reports/s1_formula_copy_openpyxl_worker_e2e_windows_native_20261025.json`（worker 写穿通过、读回拒绝、样式表越界定位）。

## 若不采纳

openpyxl worker 对该真实模板族不可路由（样式表完整性缺陷），S1-F01 公式复制维持无可用引擎；umya 路线已有 1,036 项对拍差异同样不可路由。请 main 明确裁决。
