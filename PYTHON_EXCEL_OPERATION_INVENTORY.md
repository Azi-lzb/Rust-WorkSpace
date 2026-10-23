# Python Excel 操作入口清单

盘点日期：2026-10-23  
基线：Open-Data-Audit core、py_statistics_tools PyCode 与 NoComPyTools  
状态：静态源码/API 入口清单，尚非逐菜单端到端调用链审计。

配套架构与路由原则见 [EXCEL_ENGINE_ARCHITECTURE.md](EXCEL_ENGINE_ARCHITECTURE.md)。业务功能范围/状态继续以 [audit-core/MIGRATION_INVENTORY.md](audit-core/MIGRATION_INVENTORY.md) 与 [pytools-core/MIGRATION_INVENTORY.md](pytools-core/MIGRATION_INVENTORY.md) 为准。

## 1. 盘点口径

- 搜索范围为两个只读源项目的业务 core 目录，不读取 tests、.venv 或生成缓存。
- 使用 Python AST 静态识别函数定义、Workbook/openpyxl 对象操作、保存、单元格/样式/批注/尺寸/页面设置赋值，以及 ZIP/XML 访问。没有导入或执行源项目模块，没有向源目录写文件。
- 本表把可疑的 Excel 写入函数列为“待核函数”；静态 AST 不能解析别名、反射、跨函数 helper、DataFrame/COM 封装或配置运行时分支。调用链未核完前，不将任何操作标为迁移完成。
- 保存/写入函数分为新建输出、原位修改、配置修复或探针报告。一个函数包含 save 不等于它只做单一操作；后续必须逐条补上具体输入、输出和前后置条件。
- PyCode 是 py_statistics_tools 主行为基线。NoComPyTools 是同产品的变体参考；差异尚未全部核实。工作表提取只迁移普通无批注版，COM 批注版不纳入。
- 本轮对 PyCode、NoComPyTools 业务 core 和 Open-Data-Audit `core/src` 补做了 openpyxl/工作簿对象/ZIP/XML/旧 Office 引擎静态扫描，并在 `EXCEL_ENGINE_ARCHITECTURE.md` 增加操作到候选路径映射。静态扫描只定位入口，不能证明菜单可达、执行顺序、保存时机或错误回滚。
- 2026-10-22 收尾时两个源项目 Git 工作树都存在未提交变化，且没有本轮开始前的状态快照；其来源未知。当前模型对拍导入的 `base_audit/models.py` 未显示为修改文件。后续 Python 基线记录应固定源文件版本/工作树状态，禁止清理源目录。

## 2. py_statistics_tools

源根目录：

- PyCode：C:/Users/AZI/Desktop/VibeCode/4.pytools-main/PyCode/core
- NoComPyTools：C:/Users/AZI/Desktop/VibeCode/4.pytools-main/NoComPyTools/nocompytools/core

### 2.1 工作簿写入与修改函数

| 源模块/函数 | AST 可见操作 | 分类与迁移验证重点 |
|---|---|---|
| config_init.py：initialize_or_repair_config；_ensure_sheet_headers；_ensure_rename_sheet_extra_cells；_ensure_global_defaults；_ensure_archive_type_examples；_ensure_submission_check_examples；_ensure_file_sort_examples；_ensure_header_comments | 新建/补齐工作表、单元格、说明批注、样式、保存 | 配置初始化/修复；区分新增、覆盖和保留其他配置；校验批注和样式不因修复丢失 |
| dedup_311.py：run_check_by_comment；run_delete_by_comment；run_check_by_config；run_delete_by_config；run_append_by_config；_append_unique_source_to_target；_delete_duplicates；_mark_duplicates | 批注定位、单元格写入/标色、删除重复行、创建/清理目标 sheet、追加目标簿、保存 | 分开检查与执行；确认哪些分支只生成报告，哪些更改源副本或目标簿；核对重复键/顺序/追加幂等 |
| io_excel.py：write_workbook；append_to_target；_apply_data_date_format | 新建多 sheet 工作簿、追加/替换目标 sheet、写日期数字格式 | 迁移到 rust_xlsxwriter 或 ExcelModifier 前确认 append 与 replace 对未触碰 sheet 的保留要求 |
| feature_1_1_split_village_bank.py：_process_one_workbook | 行插入/删除及工作簿保存 | 按银行类别拆分与行重排；验证原始 sheet、公式和格式变化范围 |
| feature_1_2_normalize_institution.py：_process_one_workbook | 保存；具体 cell 变更可能由 helper/其他路径完成 | 逐层追踪赋值和 COM 对照，不从 save 语句推断修改内容 |
| feature_1_3_remove_foreign_bank.py：_delete_foreign_rows；_process_one_workbook | 删除行、保存 | 行删除对公式引用、合并、筛选和其他 sheet 的影响 |
| feature_1_4_fx_header.py：_process_one_workbook | 遍历名称包含“外汇”的工作表，将 odd/even/first 页眉 left/center/right 中所有“万元”替换为“万美元”；仅在至少一个页眉变化时保存。对 `.xlsm` 使用 keep_vba | Rust 候选 `pytools_core::fx_header::build_fx_header_plan` 仅接受 `.xlsx`，通过 ExcelEditPlan + ExcelModifier 执行窄 OOXML 页眉文本补丁；PyCode 与 NoComPyTools 对应源文件 SHA-256 相同。候选与 Python 基线的合成对拍见 `pytools-core/probe/reports/fx_header_parity.json`；真实工作簿、`.xlsm`/VBA、菜单/CLI 和源路径替换仍未验证 |
| feature_1_5_region_sum_check.py：run_region_sum_check；_write_result_header；_write_result_row；_autofit | 新建结果簿、写值、写字体/填充/对齐、列宽、保存 | `pytools-core::region_sum::run_region_sum_check` 已接入 Rust `pytools region-sum-check` CLI；合成 XLSX 对拍统计、结果值/公式、样式和列宽见 `pytools-core/probe/reports/region_sum_parity.json`。原 GUI 菜单及真实模板仍未验收 |
| feature_1_6_extract_sheets.py：run_extract_sheets；_copy_full_sheet_with_style；_remove_default_blank；_write_values | 新建/删除 sheet、复制值与样式、行列尺寸/隐藏/大纲、合并区域、保存 | 仅普通无批注版迁移；完整文件流程仍需测批次发现、目标追加、整表复制和未修改对象 |
| feature_1_8_adjust_rural_loan.py：run_adjust_rural_loan | 保存；具体改单元格和范围经 helper/分支展开后登记 | 迁移前逐项确定金额类型、舍入、公式处理和写回范围 |
| feature_1_9_adjust_rural_loan_quarterly.py：run_adjust_rural_loan_quarterly | 保存；具体改单元格和范围经 helper/分支展开后登记 | 同上，并覆盖季度/分地区输出契约 |
| print_310.py：run_print_fast_by_comment；run_print_keep_by_comment；run_print_config_all_modes；_copy_block；_apply_print_setup；_apply_sheet_order_footer；_save_with_suffix_fallback | 新建/清除 sheet、块复制、值与样式/批注复制、行列尺寸、合并、打印区域/页面设置/页脚、保存 | 打印只写页面设置、不渲染；测试方向、fit、scale 清除、纸张、打印区域、居中和页脚；不迁移 PDF 渲染 |
| summary_321_322.py：run_summary_by_usedrange；run_summary_by_comment | 新建汇总工作簿、写值与样式、批注驱动读取、创建工作表、保存 | 核对行/列/来源定位和输出顺序；批注只按旧基线业务语义解析 |
| convert_33_34_35_37.py：run_batch_rename_sheet_37；_python_excel_convert_one | Sheet 改名、Excel 文件转换和保存 | .xls 改名先走显式 .xls→.xlsx 转换候选；原 Python 转换器仅作对拍基线，不能作为生产依赖 |
| convert_33_34_35_37.py：_load_content_rename_config；_read_cells_openpyxl；_read_cells_com；_already_renamed_by_content；run_batch_rename_by_content_46；_load_file_local_rename_map；_apply_local_name_map；run_batch_rename_local_map_48；_remove_content_prefix_keep_sort_prefix；run_batch_unrename_by_content_47；_load_file_sort_rules；_match_file_sort_rule；_strip_sort_prefix；run_file_sort_prefix_49；run_remove_file_sort_prefix_410 | 4-6 读取配置单元格值并添加内容前缀（`.xls/.xlt` 旧路径走 COM）；4-7 读取内容规格和指定单元格后取消文件名前缀；4-8 按 M:N 配置顺序替换任意文件名片段；4-9 按文件名匹配并改名、处理冲突、生成 xlsx 报告；4-10 删除排序前缀，必要时读取内容单元格 | 4-8 已接入 `pytools rename-local-map`、4-9 已接入 `pytools file-sort-prefix` Rust CLI；4-6/4-7 已接入 Rust CLI；4-10 已接入 `pytools remove-file-sort-prefix` Rust CLI。4-6/4-7 候选只接受 `.xlsx/.xlsm` 并用共享 Calamine 读取，属于“只读工作簿 + 文件系统改名”，不进入 ExcelModifier；4-8 以共享 Calamine 读取 M:N 映射配置，接受任意扩展名并只改名。4-8 覆盖映射顺序、链式替换、中文片段、空目标、同名冲突、路径分隔符拒绝及失败保留源文件；合成场景 5 改名/5 跳过。4-9 使用共享 Calamine 读取排序规则、rust_xlsxwriter 输出报告；8 路径合成对拍的统计、报告打开性、最终文件名集合和文件字节与 Python 基线一致；Rust CLI 输出时间戳文件及 stderr 日志，成功项消息已断言，具体时间/并发行为未逐项对拍。4-10 的 10 路径主场景与空输入场景均经 Rust CLI 对拍，报告行、最终文件名集合和空报告行为与 Python 基线一致，并检查了成功日志和改名文件字节。4-10 的 `.xls/.xlt` 内容读取和完整 logger 行为仍未验证。原 GUI 菜单、真实配置、非 UTF-8 文件名、Windows 原生 Rust 和目标 Linux 文件系统仍未验证。4-6/4-7 的 `.xls/.xlt` COM 路径只作源行为基线，不迁入主路径 |
| NoComPyTools/core/xls_to_xlsx.py：run_xls_to_xlsx；_convert_one；_convert_one_with_office | 显式通过 LibreOffice 或 Windows Office/COM 转换 xls 为 xlsx | 只迁移到允许的 sidecar 转换入口；需校验 sheet、数据、公式和输出可打开性；缺少引擎时报告“该功能需要外部引擎” |

以上普通路径在 PyCode 与 NoComPyTools 中有大部分同名实现。NoComPyTools 静态扫描对应的主要函数为：config 初始化、dedup 五个入口、io_excel 两个入口、feature_1_1/1_2/1_3/1_4/1_5/1_6/1_8/1_9、print 三个入口、summary 两个入口和 Sheet 改名。PyCode 另外存在多个 COM 实现；两套实现的全部行为差异、数据口径和异常行为尚未对齐。

### 2.2 PyCode/NoCom 分支边界

- pytools-core 当前已有配置加载、纯规则 helper、工作簿抽取、工作表区域值提取 CLI、地区总分校验 CLI、4-8 局部文件名映射 CLI、4-9 文件排序前缀 CLI、4-10 取消排序前缀 CLI 和机构名称规范化候选；内容前缀添加/取消 4-6/4-7 已有显式文件列表 CLI 的窄范围端到端候选，具体范围与对拍限制见 pytools-core/MIGRATION_INVENTORY.md 和对应 probe/reports 报告。
- 机构名称规范化新增 `build_institution_normalize_plan`：通过共享 `excel-reader` 读取值并构造统一 `ExcelEditPlan`；3 个单元格操作及统计与 Python 基线匹配，报告为 `pytools-core/probe/reports/institution_normalize_plan_parity.json`。该测试的 workbook inventory 标为不完整，未路由、未执行修改。原 `run_normalize_institution` 仍直接调用 excel-editor/umya，是未获准的旧候选执行路径；该业务功能尚未收敛到 ExcelModifier。
- 普通版工作表提取现有 Rust area-only CLI 候选 `pytools extract-sheets-area`，配置加载、显式输入文件批处理及新建 XLSX 值输出与合成 Python baseline 对拍通过，报告见 `pytools-core/probe/reports/sheet_extract_area_parity.json`。整表路径 `_copy_full_sheet_with_style` 涉及值、样式、行列尺寸、隐藏、大纲和合并，仍未实现；CLI 不含产品菜单/dispatcher、目录扫描、既有输出簿追加或真实模板验收。
- feature_1_6_extract_sheets_com.py 的批注分支不进入迁移范围；普通版是否仍有未由 AST 识别的间接复制行为需结合函数调用链确认。
- NoComPyTools/core/convert_33_34_35_37.py 的内容读取路径包含 python-calamine 与 openpyxl 回退；生产 Rust 路径必须改为 Rust calamine/已验证的 OOXML 补充，Python 仅保留为对拍基线。该模块还包含重命名、归档等文件系统行为，需在业务功能清单核对，但不属于 ExcelModifier 修改操作本身。
- PyCode 还存在 `feature_1_5_region_sum_check_com.py`、`feature_1_6_extract_sheets_com.py`、`feature_1_8_adjust_rural_loan_com.py`、`feature_1_9_adjust_rural_loan_quarterly_com.py`、`print_310_com.py`、`summary_321_322_com.py`、`timeline_396_com.py`、`wide_397_398_com.py` 等 COM 对照实现；它们不是可随产品保留的 COM 业务旁路。除公式重算与同类格式转换外，业务行为需由 Rust/允许的 Go 修改器实现并对拍。

### 2.3 菜单与 dispatcher 差异

- `PyCode/main.py` 同时导入普通版与 COM 版实现，并在打印、汇总、地区总分校验、工作表提取及涉农贷款调整等路径分派；`NoComPyTools/nocompytools/main.py` 使用独立 dispatcher。具体选择条件、默认值、失败回退和两个产品变体的行为差异仍待逐项核实，不能仅凭函数同名认定等价。
- 工作表提取以用户已确认的普通无批注版为基线，`feature_1_6_extract_sheets_com.py` 的批注版不迁移。`print_310_com.py` 的 PDF 导出入口也因当前目标排除 PDF 而不迁移；非 PDF 打印任务只写页面设置/输出工作簿。
- 普通版与 COM 版都包含的规则/结果处理，除公式重算本身外必须迁入 Rust 主路径；不得以 COM 菜单分支作为“允许旁路”的理由。

## 3. Open-Data-Audit

源根目录：C:/Users/AZI/Desktop/VibeCode/3.open-data-audit/core/src

### 3.1 既有工作簿编辑、模板复制与 OOXML 定点写入

| 源模块/函数 | 已观察到的操作类型 | 候选修改计划类别 |
|---|---|---|
| base_audit/native/openpyxl_workbook.py：copy_formula_ranges_from_template；add_external_sheets；_copy_sheet | 公式/值和模板区域复制、样式/批注/合并/尺寸、跨簿增加 sheet、保存 | 区域/工作表复制和公式写入；必须核对引用、defined names 和跨簿 relationship |
| base_audit/native/template_merge.py：copy_sheet；run_template_merge_native；rebase_external_references；copy_sheet_names；localize_base_names；_drop_original | 复制工作表值/公式/样式/批注/合并/尺寸，名称增删改，引用重写，工作表结构变更 | 多操作整批修改；不能按单元格拆成不同引擎 |
| base_audit/merge_org.py：_copy_sheet_openpyxl；_merge_one_org_openpyxl；_copy_sheet_named_ranges_as_local；_localize_workbook_named_ranges；_write_template_merge_report | 合并多个组织工作簿、工作表复制、样式/批注/合并/尺寸、名称本地化、生成报告 | 模板/工作簿合并；需要真实模板与全 package parts 对拍 |
| base_audit/formula_region_writer.py：DirectOoxmlFormulaRegionWriter.write；_copy_structured_rule_sheet；_TemplateReader.sheet_part | OOXML ZIP/XML 直接读写，复制结构规则 sheet、写入公式区域和部分相关结构 | 若是确定位置/已知 part，可作为 namespace-aware patch 候选；现有 Python 逻辑需按 Rust 原语拆分、对拍关系和 sharedStrings |
| base_audit/external_sheet_writer.py：DirectOoxmlExternalSheetWriter.prepare；write；_parse_data_sheet；_validate_output | 直接组装外部 sheet、读写 workbook/worksheet/rels/sharedStrings/Content_Types 等 | 高关联包操作；仅在明确每个部件变化和关系维护规则后考虑 Rust Patch，否则整批交给通过测试的编辑引擎 |
| base_audit/systems/s2_report_collection/exporter.py：_write_direct_ooxml | 直接 OOXML 写结果文件 | 判断是从零生成还是修改模板；从零生成优先 rust_xlsxwriter，不能默认保留旧 writer |
| base_audit/systems/s3_central_statistics/fast_ooxml.py：TemplateCompiler._compile；FastOoxmlRenderer.render_org；_validate_output；ooxml_structure.py 的关系/名称/公式读取 | 自定义模板编译和 OOXML 输出/校验，直接处理共享字符串、定义名称、关系 | 逐函数判定新建还是既有文件 patch；同时保留 validation 作为 probe 要求 |
| s1audit_configured/fast_ooxml.py：write_tables；_write_sheet；_styles | 从零构建 ZIP/XML 工作簿、样式、工作表、关系和名称 | 新建表优先评估 rust_xlsxwriter；不得把该定制实现扩展成第三套完整 OOXML writer |

### 3.2 新建结果、配置、历史与报告输出

以下函数含 Workbook/worksheet 创建或保存路径，需按“新建输出”与“原位更新配置/历史”拆开；新建输出优先 rust_xlsxwriter，原位改模板进入 ExcelEditPlan。

- 配置/历史/系统设置：base_audit/config_guide.py::write_guide_to_file；base_audit/feature_log.py::write_feature_log_xlsx；base_audit/name_config.py 的配置初始化、重置、历史迁移与编辑草稿函数；base_audit/node_flow_config.py 的 DAG/工作表重置、替换和保存；base_audit/period_compare.py::ensure_default_config 与 append/write 输出；S2/S3 config 的默认配置写入函数。
- 检查/汇总/审核结果：base_audit/native/history_io.py::write_history_xlsx；native/result_writer.py 的 write_openpyxl；native/summary.py 的 write_openpyxl/汇总输出；preflight_xlsx.py 的两个 xlsx 报告函数；region_summary.py 的历史补充；S2 exporter/regression；S3 comparison/explanation/service；S5 expression/formula candidate 输出；s1audit_configured/outputs.py::_write_result_workbook_openpyxl。
- 本节列出候选入口，不代表每项新输出都应经过原位编辑器。正式迁移前应通过源调用方和输出路径分配确认生命周期。

### 3.3 元数据读取而非写入的 OOXML 能力

- base_audit/cf_reader.py::OoxmlConditionalRuleReader：ZIP/XML 读取条件格式、共享字符串和 dxf。
- base_audit/conditional_context.py::OoxmlCommentReader：ZIP/XML 解析 comments 与工作表 relationship。
- base_audit/systems/s3_central_statistics/ooxml_structure.py：工作表、relationship、公式文本和 defined names 结构读取。
- s1audit_configured/source_reader.py：OOXML/样式/工作表元数据 reader。
- base_audit/discovery.py::read_xlsx_sheet_names、preflight_xlsx.py::_merged_anchor_lookup：读取 sheet / merge 元数据。

这些读取路径不应与修改器合并成完整第二读取引擎；Rust 默认 reader 仍是 calamine，OOXML reader 只补经证实缺失的属性。

### 3.4 旧 Office 引擎入口（只作行为基线）

- `base_audit/excel_com.py` 的 `ExcelSession` 覆盖模板打开/检查、规则执行、公式计算、条件格式问题提取、跨工作簿 sheet 复制，以及 `write_navigation`、`write_history`、`write_preflight_report`、`write_summary` 等结果写入。不能将整个 `ExcelSession` 当作 sidecar；只有显式公式重算职责可以调用 `sidecar/lo-convert`，其余读、算、改、写都要迁入纯 Rust 主路径或已验证的 ExcelModifier。
- 已观察到 `ExcelSession` 被 `engines/excel_wps.py`、`merge_org.py`、`service.py`、`region_summary.py`、`workflow/runner.py`、`workflow/adapters/audit_handlers.py` 和 S3 `office_eval.py` 调用；`formula_region_writer.py` 还调用 `session.apply_rules`。这些入口仍需映射各自的业务数据流，不能把调用方整个工作流当作重算请求。
- `base_audit/engines/libreoffice.py` 提供旧 LO/KARE 发现和重算流程。迁移时只允许显式重算能力进入 sidecar，必须在输入/输出两端验证公式文本及缓存值；旧 Python 的 ZIP/XML 修复代码需在 Rust 侧按 namespace-aware 规则重写。
- `NoComPyTools/core/xls_to_xlsx.py` 和 `convert_33_34_35_37.py` 的旧转换入口仅作为行为基线。符合已确认同类格式转换范围的 `.xls/.et→.xlsx`、`.doc→.docx` 可由独立 sidecar 提供；PDF 和跨类别转换不在范围内。
- Office/COM 方法中的业务读取、筛选、规则、公式运算或报告输出，不得因为旧实现使用 Excel COM 就进入 sidecar。公式重算需与业务计算区分，sidecar 不执行审核/统计规则。

## 4. 当前 Rust 对应关系与缺口

| Python 行为类别 | 现有 Rust 对应 | 当前结论 |
|---|---|---|
| 工作簿值/公式/合并等读取 | crates/excel-reader；pytools-core::config、workbook_extract、institution_normalize probe、file_sort_prefix 配置加载使用 | 有合成 xlsx 对拍；xls/xlsm/xlsb/csv 和真实模板矩阵未完成 |
| 新建 xlsx 输出 | crates/excel-writer 对 rust_xlsxwriter 的类型导出；audit-core 探针与 pytools-core 4-9/4-10 候选 API 使用 | 探针 9/9；4-9、4-10 的简单结果报告均有端到端合成对拍；尚无通用批量写入 API、完整产品输出或真实模板验收 |
| 既有 xlsx 修改 | crates/excel-editor 有 ExcelEditPlan/ExcelModifier 骨架；机构名称规范化可生成计划但旧候选仍直调 umya | 计划与 3 个 Python 基线变更匹配；未完成 package inventory、统一 backend/validator 执行、真实模板与默认路由验收 |
| 定点 OOXML 修改 | crates/ooxml-patch 的未验证部件拒绝与列定义/列宽补丁 | 仅少量 patch 行为有合成样本报告；没有公式区/批注/页面设置完整能力 |
| 比较和报告 | crates/probe 导出与比较 JSON | reader probe 为主；没有自动消费 ExcelEditPlan 的修改 probe |
| Go worker | 当前不存在 | 尚无 go.mod、excelize 调用或 IPC 协议 |
| sidecar | sidecar/lo-convert 独立 binary 骨架 | 公式重算与格式转换均未接通 |

## 5. 下一轮逐项核实模板

对上表每个业务入口补齐一行结构化记录：

1. 旧项目菜单/dispatch → 源文件与函数 → 被调 helper → 输入格式/模板 → 配置字段 → 输出文件/文件结构。
2. 该路径是新建、原位修改、拷贝、转换、只读检查中的哪一种；若原位修改，生成完整 ExcelEditPlan 操作序列。
3. 操作前文件清点的对象集合与修改后必须保留的对象；指出实际未涉及项目为“不适用”，没有对拍的为“未验证”。
4. Python 基线命令、Rust probe 命令、JSON 比较字段、真实/合成样本区分、性能/内存测量、失败时源文件安全和回退路径。
5. PyCode 与 NoComPyTools 的差异证据及迁移决策；任何 COM 专有功能不能悄悄混入普通 Rust 路径。

静态 AST 能确认“存在该类操作语法”，不能确认菜单是否可达、操作是否执行、错误情况下是否保存或属性是否真的变化。只有完成上述记录及对拍后，才能把业务入口状态由未验证改成保留/不适用等结论。
