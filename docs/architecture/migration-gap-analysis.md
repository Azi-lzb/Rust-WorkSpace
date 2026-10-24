# Python → Rust 迁移缺口分析（migration-gap-analysis）

> 审计日期：2026-10-23。编号与 [feature-inventory.md](feature-inventory.md)、[feature-dataflow-matrix.md](feature-dataflow-matrix.md) 一致。
> 状态六分类定义见 feature-inventory §0.2：尚未迁移 / 部分核心 / 代码未接入口 / 接入依赖 Python / 接入未验 / 完整迁移。
> 判定纪律：不因存在同名 Rust 函数认定迁移完成；"接入"指迁移产品入口（Rust CLI 子命令 / `open-data-audit serve`）可达；全部合成对拍不等于真实样本验收。

## 1. 总量与状态分布

| 状态 | pytools（48） | ODA（23） | 合计（71） |
|---|---|---|---|
| 完整迁移并通过验证 | 0 | 0 | **0** |
| 接入未验 | 44 | 1（AUD-S2-F01） | 45 |
| 部分核心 | 1（PT-1-2） | 2（AUD-S1-F01、AUD-S3-F01） | 3 |
| 尚未迁移 | 0 | 19 | 19 |
| 接入依赖 Python | 0 | 0 | 0 |
| 范围外/源已移除/待确认 | 3（PT-3-5、PT-4-11、PT-SV1） | 1（AUD-S4） | 4 |

**核心结论：两产品没有任何一项功能达到"完整迁移并通过验证"。** 主要缺口不是代码量（pytools 44 项已有 CLI），而是三类验收缺失：真实业务样本（0 个）、ExcelModifier 生产能力登记（0 条）、目标平台（glibc 2.28 deb）。

## 2. pytools 映射明细

证据列缩写：`PR`=pytools-core/probe/reports 下 JSON；CLI 列为 `pytools-core/src/bin/pytools.rs` 分发表（:70-117）中的命令。

| 编号 | Python 源 | Rust 模块 | CLI | 状态 | 对拍证据 | 主要缺口 |
|---|---|---|---|---|---|---|
| PT-1-1 | `compare_393.py` | `timeline_ops.rs` | timeline-compare | 接入未验 | date_parse 22/22、path_map/关键字/列规格 28/28 向量（`PR/date_parse_baseline.json` 等）；无端到端 | 端到端 Python 对拍；.xls COM 路由语义（PyCode 自动路由） |
| PT-1-2 | `timeline_396.py` + `rule_engine` | `rule_engine.rs`、`workbook_extract.rs`、`sheet_extract.rs` | extract-sheets(-area)（范围外接入） | 部分核心 | workbook_extract 6/6（`workbook_extract_parity.json`）；sheet_extract 规则 11/11；area 候选端到端（`sheet_extract_area_parity.json`）；整表 ImportRange 候选 0 差异（`sheet_extract_full_import_parity_20261023.json`） | 用户范围明确排除 1-2 完整功能（`MIGRATION_PLAN.md:20`）；追加既有簿未实现；xlsb 合并区读取 API 缺口 |
| PT-1-3/1-4 | `wide_397_398.py` | `timeline_ops.rs` | wide-summary(-row-suffix) | 接入未验 | 无端到端 | 同 PT-1-1；追加扩列契约 |
| PT-1-5 | `precheck_399.py` | `timeline_ops.rs` | timeline-precheck | 接入未验 | 配置加载 11/11（`config_task_loaders_parity.json`，间接） | 预校验文案/行为对拍 |
| PT-2-1..2-4 | `dedup_311.py` | `dedup_ops.rs` | dedup-comment-plan / dedup-config-plan / dedup-precheck | 接入未验 | **无 Python 对拍**（`dedup_execution_wiring_20261023.json` 标 implementation_pending_validation） | SetCellFillColor/DeleteRows 在真实模板的保留能力无验收记录；批注读取（OOXML 补读）未对该功能对拍；删行不修正合并区的语义口径未决 |
| PT-2-5 | `dedup_311.py:368` | `dedup_ops.rs` | dedup-append-plan/execute | 接入未验 | 无对拍；新目标走 writer 直写 | ImportRange 追加能力未验收；公式原文搬运语义（源 data_only=False）未决 |
| PT-2-6 | `dedup_311.py:530+` | `dedup_ops.rs` | dedup-precheck | 接入未验 | 无 | 只读，风险低 |
| PT-3-1/3-2 | `print_310.py` | `print_ops.rs` | print-comment-plan | 接入未验 | **无对拍**（`print_config_execution_wiring_20261023.json`） | 样式+批注+行高列宽+合并整批复制能力未验收；页面设置十属性未逐项对拍；calamine 无样式模型——需要 umya 或重建式写入的契约决策 |
| PT-3-3 | `print_310.py:593-815` | `print_ops.rs` | print-config-plan | 接入未验 | 无 | "清空目标 sheet 重写"是原位修改，需已验收后端；模式 1/2/3 差异未对拍 |
| PT-3-4 | `print_310.py:510-590` | `print_ops.rs` | print-config-precheck | 接入未验 | 无 | 只读 |
| PT-3-5 | `print_310_com.py`（COM PDF） | — | — | 范围排除 | — | workspace 规则排除 PDF；无需迁移 |
| PT-4-1 | `convert_33_34_35_37.py` | `batch_rename.rs` | rename-batch | 接入未验（最接近完整） | 合成端到端 + 真实配置窄 smoke（`batch_rename_parity.json`、`batch_rename_real_config_smoke.json`） | 真实业务文件覆盖率；目标文件系统竞争行为 |
| PT-4-2 | `xls_to_xlsx.py`（NoCom）/`run_excel_convert_34`（PyCode COM-only） | `conversion.rs` + `sidecar/lo-convert` | convert-excel / convert-file | 接入未验 | sidecar 校验器内置（sheet/值/公式文本对拍，`sidecar/lo-convert/src/main.rs:221-289`）；真实 .et/.xls 样本未验证 | 两变体行为差异（NoCom LO 分支 vs PyCode COM-only）需定契约；WPS ProgID 链 |
| PT-4-3 | `word_convert.py` | `conversion.rs` | convert-word | 接入未验 | sidecar docx 文本比对（`main.rs:291-354`） | 真实 .doc 样本 |
| PT-4-4 | `run_batch_rename_sheet_37` | `sheet_rename.rs` | rename-sheets-plan | 接入未验 | 无 | sheet 改名涉及公式/名称/打印区域引用（规则 §5 W03/W04），计划未验证执行 |
| PT-4-5 | `run_submission_check_48` | `submission_check.rs` | submission-check | 接入未验 | 无 | 报告格式对拍 |
| PT-4-6 | `run_batch_rename_by_content_46` | `local_filename_map.rs` 等 | rename-by-content | 接入未验 | 合成 12 路径（`content_rename_46_parity.json`） | 真实 .xlsm/宏；.xls/.xlt COM 读取不迁（源基线）；日期/公式缓存格 |
| PT-4-7 | `run_batch_unrename_by_content_47` | 同上 | unrename-by-content | 接入未验 | 合成 13+2 路径（`content_unrename_47_parity.json`） | 同上 |
| PT-4-8 | `run_batch_rename_local_map_48` | `local_filename_map.rs` | rename-local-map | 接入未验 | 合成 10 路径（`local_filename_map_48_parity.json`） | 非 UTF-8 文件名；真实配置 |
| PT-4-9 | `run_file_sort_prefix_49` | `file_sort_prefix.rs` | file-sort-prefix | 接入未验 | 合成 8 路径 + 日志断言（`file_sort_prefix_operation_parity.json`） | 日志时间/并发与 Python logger 逐字对拍 |
| PT-4-10 | `run_remove_file_sort_prefix_410` | `file_sort_prefix.rs` | remove-file-sort-prefix | 接入未验 | 合成 10 路径（`file_sort_prefix_remove_parity.json`） | .xls/.xlt 内容读取；多单元格规格 |
| PT-4-11 | 历史功能 | — | — | 源已移除 | grep `PYTOOLS_MODE` 0 命中 | 无迁移对象；确认产品范围不含即可 |
| PT-5-1 | `config_init.py` | `config_admin.rs` | config-init / config-repair-plan | 接入未验 | init+重读 smoke（`pytools_batch_scope_20261023.json`）；修复计划无对拍 | 11-sheet 源模板结构、默认值、报错语义未对拍；修复是原位修改需后端 |
| PT-5-2 | `_view_config` | `config_admin.rs` | config-view | 接入未验 | 无 | 只读 |
| PT-6-1 | `summary_321_322.py` | `summary_ops.rs` | summary-used-range | 接入未验 | 无 | used-range 语义（calamine extent vs COM UsedRange 差异）；.xls/.xlsb 输入 |
| PT-6-2 | `summary_321_322.py`（批注指令） | `summary_ops.rs` | summary-comments | 接入未验 | 无 | 批注指令解析（OOXML 补读）未对该功能对拍；.xls 模板转换路径 |
| PT-7-1 | `feature_1_1_split_village_bank.py` | `batch_ops.rs` | village-split-plan | 接入未验 | 无对拍 | 删插行+合并区语义（COM 版随行移动 vs openpyxl 不修正）必须先定口径 |
| PT-7-2 | `feature_1_2_normalize_institution.py` | `institution_normalize.rs` | institution-normalize-plan | 接入未验 | 计划级 3 变更格（`institution_normalize_plan_parity.json`） | 旧 `run_normalize_institution` 直接调 umya，仅编译进测试（`institution_normalize.rs:99-101` 注释）——**代码未接入口**的反面教材；inventory 不完整故计划不可路由 |
| PT-7-3 | `feature_1_3_remove_foreign_bank.py` | `batch_ops.rs` | remove-foreign-bank-plan | 接入未验 | 无 | "(金融局)"副本保护语义须保留；删行语义同 PT-7-1 |
| PT-7-4 | `feature_1_4_fx_header.py` | `fx_header.rs` | fx-header-plan | 接入未验 | 合成 1 样本（`fx_header_parity.json`，页眉补丁 passed） | 批量文件统计聚合缺失；.xlsm/VBA；源路径原位替换未做 |
| PT-7-5 | `feature_1_5_region_sum_check.py` | `region_sum.rs` | region-sum-check | 接入未验 | 合成 5 场景 + CLI 集成测试（`region_sum_parity.json`） | 真实模板；非 XLSX；性能 |
| PT-7-6 | `feature_1_6_extract_sheets.py` | `sheet_extract.rs` | extract-sheets(-area) | 部分核心 | area 值提取端到端（`sheet_extract_area_parity.json`）；整表 ImportRange 候选 0 差异（对照源工作表属性；旧 Python helper 输出本身损坏——openpyxl 读样式 IndexError，见 `sheet_extract_full_import_parity_20261023.json`） | 整表/追加路径被能力门控；COM 批注版不迁移（已确认）；旧 Python helper 样式索引异常需源项目层面决定（本 workspace 不改源） |
| PT-7-7/7-8 | `feature_1_8/1_9` | `batch_ops.rs` | rural-loan-plan / quarterly-rural-loan-plan | 接入未验 | 无 | 金额类型/舍入/公式覆盖为值的语义对拍未做 |
| PT-8-1..8-5 | `convert_33_34_35_37.py` 归档段 | `archive_ops.rs` | archive-*（5 命令） | 接入未验 | 合成 6 场景端到端（`archive_parity_20261023.json`） | 真实目录/权限/竞争；GUI 菜单集成 |
| PT-9-1..9-3 | `main.py` 快捷段 | `quick.rs` | open-config / open-latest-output / open-output-dir | 接入未验 | smoke | 目标平台文件管理器集成（xdg-open） |
| PT-SV1 | `survey_stats.py`（源未接线） | — | — | 待确认 | — | 是否纳入产品范围需用户决定；不静默遗漏 |

## 3. ODA 映射明细

| 编号 | Python 源 | Rust 现状 | 状态 | 证据/缺口 |
|---|---|---|---|---|
| AUD-S1-F01 | `web_app.py::start("audit")`、`service.py`、`workflow/*`、`conditional_engine.py`（COM）、`native/*`、Rust writer | `s1_template.rs`（模板推荐，CLI template-suggest）、`discovery.rs`（periods 49/49、目录 9/9）、`models.rs`、`template.rs`（解析 helper 24 case）、`name_config.rs`（31 case）、`period_pairs.rs`（配对预览，经 web serve 可达）、`unit_conversion.rs`（62 case） | 部分核心 | `probe/reports/template_recommendation_parity.json`（4 合成）+ `template_recommendation_real_smoke.json`（真实 config-real 8 模板 × 68 源文件一致）；缺口：条件格式求值引擎（COM 逐格触发/优先级/相对引用）无 Rust 对应物——S1 最大单项缺口；规则执行、结果回写、历史、重复运行一致性 |
| AUD-S1-F02 | `service.summarize_regions`、`native/summary.py`、python-calamine 读取、Rust writer | 无 | 尚未迁移 | 读取引擎三态语义（AUTO 回退/手动严格错误，`test_s1_summary_backends.py` 合同）需在 Rust 重新定义 |
| AUD-S1-F03 | `merge_org.py`、`simple_handlers.handle_combine_sheets` | 无 | 尚未迁移 | 跨簿 sheet 复制+命名区域本地化=整批修改，依赖未验收的 umya/Excelize 能力 |
| AUD-S1-F04 | `workflow/*`（DAG 619 行 runner） | 无 | 尚未迁移 | 固定 DAG 配置契约、失败策略 |
| AUD-S1-F05 | `native/template_merge.py`、COM template_merge | 无 | 尚未迁移 | 名称冲突、外链重写（`rebase_external_references`） |
| AUD-S2-F01 | `PeriodAuditV2Service` + Direct OOXML exporter | `s2_service.rs::run_s2_audit` + CLI `open-data-audit s2` + `serve` web bridge（`web_bridge.rs`，8 个方法 @292-306，未迁移接口 HTTP 501） | 接入未验 | `s2_service_parity.json`：3 场景 0 差异（15/18 条发现，268/304 格比较）；缺口：真实脱敏模板、完整失败矩阵、性能、Windows 原生、deb；1-ULP 浮点差异已按 1e-15 容差处理（规则 §9.1） |
| AUD-S3-F01 | `s3 service.run_comparison` + `csv_importer` + Rust s3 writer | `central_csv.rs`（3 场景对拍） | 部分核心 | 比较引擎/规则动作/表达式/Rust writer 对等物未迁 |
| AUD-S3-F02..F05 | `explanation_exporter.py`、`cross_period_engine.py`、`fast_ooxml.py` | 无 | 尚未迁移 | S3-F05 的表单 OOXML 渲染与校验是独立大项 |
| AUD-S5-F01..F04 | `systems/s5_configured_statistics/*` + `s1audit_configured/*` | 无 | 尚未迁移 | F02-F04 行为待确认；S5 与 s1audit_configured 共享引擎使迁移应与 S1 合并规划 |
| AUD-S4 | 源码无独立 S4 系统 | — | 待验证 | systems/ 仅 s1/s2/s3/s5（已实地核对）；workspace 清单"待确认"应向用户澄清后关闭或立项 |
| AUD-C01..C07 | 配置/历史/LO 安装/诊断/DAG 编辑器/外部参考/S5 表达式 CLI | 无（serve 前端仅 S2 可用，其余 501） | 尚未迁移 | 前端入口不表示功能已迁移（`audit-core/AGENTS.md` 前端边界） |

## 4. "代码未接入口"级清单（feature 级之外）

| 项 | 位置 | 说明 |
|---|---|---|
| `run_normalize_institution`（直接 umya 执行） | `pytools-core/src/institution_normalize.rs:99-101` | 注释明示仅编译进模块测试；生产必须走 ExcelModifier。不得复制该模式（`pytools-core/AGENTS.md:11`） |
| `EngineId::ExcelizeWorker` | `crates/excel-editor/src/plan.rs:560`；`pytools-core/src/edit_plan.rs:81` bail | 枚举占位，无运行时适配器；`core-go/` 仅 README；`packaging/build.sh:43` 判 go.mod 缺失 |
| sidecar `--mode native` | `sidecar/lo-convert/src/main.rs:53` | 显式拒绝（设计如此，非缺陷） |
| umya-probe / rust-fastpath-probe | `audit-core/src/bin/*` | 探针 bin，非业务入口 |
| `s1audit_configured`（S5 引擎） | 源项目 | Rust 无对应；AUD-S5 与 S1 共享引擎的迁移路径未规划 |
| `rust/s5_result_writer`（源项目内） | `3.open-data-audit/core/rust/` | 已构建但 `core/src` 无 .py 调用方——**待验证**（历史遗留或未索引调用） |

## 5. 反向依赖（源产品已嵌入 Rust）

见 feature-inventory §3.4。影响：ODA 的 S1-F01/F02、S3-F01 在源产品中已是"Python 业务 + Rust 写出"混合链路；workspace 规则 §6.2/§6.3 的"Python S5/S2 专用 writer"性能基线实际对应这些实现。迁移基线登记时应把这三个 writer 视为**既有 Rust 资产**（可参考其输出契约与临时文件/回退模式），而非纯 Python 行为。

## 6. 关键问题清单（按优先级，Top 10）

| # | 问题 | 证据 | 影响 |
|---|---|---|---|
| P1 | ExcelModifier 无任何已验收生产能力登记/validator，全部修改类 CLI 只能产出计划 | `pytools-core/src/edit_plan.rs:69-85`（仅注册两后端，能力文件必须外部提供）；workspace 无 capability JSON；`dedup/print_config_execution_wiring_*.json` 标 pending | 阻塞约 19 个修改类功能从"计划"到"可执行" |
| P2 | 真实业务样本验收数为 0（两产品全部对拍为合成） | 各 probe/reports"不能据此推断的范围"列 | 迁移完成度整体不可声明；合成→真实的保真风险未度量 |
| P3 | "追加到已有簿"契约未定：Python 追加=整 sheet 重建替换（`io_excel.py:278-284`），Rust 候选显式拒绝追加 | `sheet_extract_area_parity.json` 拒绝项；PT-1-2/1-3/1-4/2-5/7-6 均涉追加 | 五项功能无对等实现路径 |
| P4 | 样式/批注读取能力缺口：calamine 无样式模型；OOXML 补读仅批注/CF/DV/页眉/extents | `crates/excel-reader/src/lib.rs` API 清单 | 打印（全样式复制）、去重（批注定位）、按批注汇总（批注指令）在 Rust 无完整读取链 |
| P5 | 行插入/删除与 sheet 结构操作的引擎未验收，且源两档语义不同（openpyxl 不修正合并区 vs COM 随行移动） | `feature_1_1_split_village_bank.py:124-125`、`dedup_311.py:160-161` vs COM 版 | PT-7-1/7-3/2-2/2-4/4-4 的对拍基准本身二义，须先定语义口径 |
| P6 | ODA S1 条件格式审核引擎（COM 条件格式求值）无 Rust 对应 | `conditional_engine.py`、`conditional_evaluators.py`；workspace 仅有 CF 元数据读取 | AUD-S1-F01 核心不可迁移；是 audit 侧最大工作量 |
| P7 | 公式重算 sidecar 有 `recalculate` XLSX 候选命令：临时副本先清除旧公式缓存，再请求 full-calc；Ubuntu/WSL2 LO 合成公式 `=SUM(A1:B1)` 的输出缓存实测为 5，且公式文本保留。缓存节点结构本身不证明通用公式结果正确 | `sidecar/lo-convert/src/main.rs`、`crates/ooxml-patch/src/lib.rs`；`ODA engines/libreoffice_adapter.py` | AUD-S1 仍须显式触发并与真实旧版重算结果对拍；COM、麒麟/UOS 与真实 S1 样本未验收 |
| P8 | 文档滞后与编号不一致：架构文档称 sidecar"未接入"（实际 `conversion.rs:155` 已接）；清单称 fx-header"未接 CLI"（实际 `pytools.rs:102` 已接）；`run_batch_rename_33` 记在"3-3"标题而菜单为 4-1；PyCode/NoCom 转换组菜单序不同 | `EXCEL_ENGINE_ARCHITECTURE.md:50`、`MIGRATION_INVENTORY.md:59` 与代码对照 | 误导后续迁移决策与验收口径 |
| P9 | 源产品死代码/未接线：PT-SV1 菜单路由丢失；去重 write_mode 死配置；PyCode 冗余 dispatch 键 | `PyC/main.py:198-203,1327-1345`；`config_xlsx.py:477` | 迁移范围确认遗漏风险 |
| P10 | 重复读取/重复序列化：P-01/P-02 每文件双开（样式+值）；PT-6-1 整批两遍打开；Rust region_sum 写出后补丁再重开 zip | `print_310.py:61-62,452-457`；`summary_321_322.py:217,240`；`region_sum.rs:564` | 性能与一致性成本；迁移时应单次读取+定点补丁（详见 excel-engine-routing-review §4） |

## 7. 下一阶段建议迁移顺序

依据：状态越接近验收越先收敛；修改类功能受 P1 制约；audit 大项（S1/S3/S5）依赖 P6/P7 立项。

1. **pytools 真实样本收敛批**（状态最近）：PT-4-1、4-6..4-10、8-1..8-5、7-4、7-5 在真实脱敏配置/目录上对拍 → 产生第一批"完整迁移"功能，验证全链验收流程本身。
2. **ExcelModifier 首批生产能力登记**（解 P1）：以 PT-7-4 页眉补丁（已有 passed 窄能力）和 PT-7-2 机构规范化为试点，建立 capability 证据模板 + validator 验收流程，再扩展到去重/打印计划。
3. **追加契约决策**（解 P3）：定义"追加=重建替换"是否为可接受契约；若否，为 ImportRange 追加能力立项对拍（PT-1-2/1-3/1-4/2-5/7-6 解锁）。
4. **时序/汇总/配置组验收**：PT-1-x、6-x、5-x 端到端对拍（读多写少，风险低）。
5. **AUD-S2-F01 真实模板验收**：已是接入未验，真实脱敏配置+两期目录通过后成为 audit 第一个完整功能。
6. **AUD-S3-F01 链路**：csv 切片已迁，比较引擎+Rust writer（源产品已有 Rust writer 资产可参考）迁移。
7. **AUD-S1-F01 立项**：条件格式求值引擎设计（P6）+ 公式重算 sidecar（P7）先行技术验证；同时澄清 AUD-S4 与 AUD-S5-F02..F04 范围。
8. **配套功能**（AUD-C01..C07）与目标平台 deb 验收收尾。
