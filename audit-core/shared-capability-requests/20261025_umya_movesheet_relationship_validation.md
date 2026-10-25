# 共享能力请求 #8：MoveSheet/历史表更新的关系快照按操作验证

- 提出方：Audit（audit-dev）
- 状态：待 main 评估；Audit 未修改共享 Excel-Core
- 前置修复：main `0bb17c7`（audit-dev `cc64e15`）归一空 `pageSetUpPr` 后，真实计划已越过 PageSettings 检查。

## 真实复现

复用请求 #7 的只读 `config-real` 样本，5 张表、9 行历史。Windows 原生 `UmyaXlsxBackend::apply` 成功，`OpenpyxlXlsxValidator` 随后以 `openpyxl changed OOXML relationships` 拒绝完整计划；源样本 SHA-256 前后相同。

只读 ZIP/XML 检查观察到：

- workbook relationships 从 7 项变为 8 项，新增 `sharedStrings`；
- 被更新的管理历史表原有 9 个外部 hyperlink relationships，输出中为 0；
- 批注/VML 关系 ID 从 `comments`/`anysvml` 改为 `rId1`/`rId2`；
- 原 Python `core/src/base_audit/native/history_io.py::write_history_xlsx` 会删除并重建管理历史工作表，因此该表旧超链接关系被移除符合其业务语义；不能仅因 XML 关系清单不同就把它判为非目标表损坏，也不能全局忽略关系变化。

## 请求

请 main 评估 Relationships 保留校验如何与编辑计划的工作表/区域操作对应：目标历史表允许符合 Python 契约的关系变化，未修改工作表仍须严格验证；输出新增的共享字符串关系及被重编号的批注/VML 关系需验证其内容与引用有效。请以真实模板和 Python 基线验证完整计划。

当前 S1 历史写回仍 Unverified；未登记生产 capability。
