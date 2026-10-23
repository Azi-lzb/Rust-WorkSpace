# Legacy test group map

> Migrated reference copied from the Open-Data-Audit Python project. Its `core/tests` and shell paths are not present in this Rust workspace, so these commands describe the legacy regression suite and are not runnable here. Keep this map for future business-migration planning; Rust probe evidence is under `probe/reports/`.
# S1 / S2 / S3 系统回归测试分组

先区分三种范围：单文件或单用例是**定向测试**；`s1`、`s2`、`s3` 是**系统回归**；不带标记运行 `core/tests` 是**core 全量回归**。`test_rule_action_engine.py` 的 27 项只覆盖 S3 规则动作相关部分，不代表全部 11 个功能。

在仓库根目录执行：

```powershell
python -m pytest -q -rs -m s1 core/tests
python -m pytest -q -rs -m s2 core/tests
python -m pytest -q -rs -m s3 core/tests
python -m pytest -q -rs core/tests
```

Windows 用户可以双击仓库根目录的 `运行全量测试.bat` 运行 Windows 全量回归，或使用 `运行全量测试.bat --no-pause`。若已安装并配置 WSL，可执行 `运行全量测试.bat --wsl-linux` 仅在 WSL 中运行 Linux 全量测试，或执行 `运行全量测试.bat --with-wsl-linux` 先测 Windows、再测 WSL Linux；多个发行版时可追加 `--wsl-distro Ubuntu` 指定发行版。WSL 发行版必须预先安装 Bash、Python 3.8+、项目运行依赖和 pytest；脚本不会自动安装依赖。Linux runner 会优先尝试 `$HOME/.venvs/open-data-audit-test/bin/python`，可在 WSL 中一次性创建环境：`python3 -m venv "$HOME/.venvs/open-data-audit-test" && "$HOME/.venvs/open-data-audit-test/bin/python" -m pip install -r shell-flask/requirements.txt pytest xlwt`。Windows AMD64 上的 WSL 通常只验证 Linux AMD64，不能替代 ARM64 实机测试。Windows 脚本会自动切换到仓库根目录，优先选择能导入 pytest 的项目虚拟环境，其次使用 PATH 中的 Python 或 `py -3`，并在报告中记录 Windows 版本和架构。

ARM64 可通过 Docker/QEMU 纳入 Windows 全量测试：`运行全量测试.bat --arm64-qemu` 仅运行 ARM64 模拟，`--with-arm64-qemu` 在 Windows 全量测试后追加 ARM64 模拟，`--all-platforms` 依次运行 Windows、WSL 原生架构 Linux 和 ARM64 QEMU。所选 WSL 发行版需能访问 Docker Engine；首次运行会联网准备 ARM64 Ubuntu 22.04 测试镜像及依赖，并在 ARM64 环境执行 core 全量 pytest 和可用发行包启动测试；有 ARM64 glibc2.23 包时，还会在 Ubuntu 16.04 ARM64 用户空间执行包启动冒烟和 DEB 安装检查。报告标出 QEMU 模拟方式。QEMU 可验证 ARM64 用户态与发行包 ABI，不能替代麒麟 V10 真机的内核、桌面、LibreOffice 和真实业务工作簿验收。

Linux 用户也可在目标机运行 `bash run_full_test_linux.sh`；应分别在原生 amd64 和 arm64 环境运行，报告会自动记录发行版、架构、glibc、Python、LibreOffice 和 Node.js 信息，并将 pytest 临时文件放在系统临时目录。Linux 源码全量 pytest 需要 Python 3.8+、项目运行依赖和 pytest；脚本优先使用项目 `.venv/bin/python` 或 `bin/python`，随后使用用户测试虚拟环境或系统 Python，不会为了包启动测试强制指定目标机 Python 版本。脚本可加 `--artifact 路径`（可重复）调用现有发行包启动冒烟测试，例如 `bash run_full_test_linux.sh --artifact shell-flask/dist/审核工具_V26.1.6.0_amd64.deb`；该冒烟会使用临时 HOME/XDG_DATA_HOME 并在完成后清理。麒麟 V10 上的 Python 3.5 不满足源码测试要求，但冻结发行包自带 Python 运行时，可在目标机使用 `bash run_full_test_linux.sh --package-only --artifact <目标架构的包>` 单独生成带系统信息的发行包启动报告；这只验证包启动和页面资源，不等同于 S1/S2/S3 全量功能验收。Windows 报告命名为 `FullTest_windows-<架构>_<时间戳>.txt`，Linux 报告命名为 `FullTest_linux-<架构>_<时间戳>.txt`，均保存在 `test-output/full-test-reports/`；全量报告包含按 S1/S2/S3 功能整理的验证目标、代表用例和本次逐用例状态（通过/跳过/失败）。

Linux 桌面真机可双击仓库根目录 `运行Linux全量测试.sh`：脚本会尝试在桌面终端显示结果、自动选择当前架构最新的本地发行包；只有检测到 `dpkg-deb` 时才自动选择 DEB，否则只查找 TAR.GZ。有 Python 3.8+ 与 pytest/项目依赖时运行 core 全量测试，否则直接用发行包自带 Python 运行启动冒烟并在报告中标明 core 未执行。脚本不联网、不安装依赖。将脚本拷到目标机时，需同时带上 `run_full_test_linux.sh`、`shell-flask/Linux-1-启动审核工具.sh`、`shell-flask/Linux-3-测试发行包启动.sh` 和同架构发行包；完整源码 pytest 还需 `core/` 与兼容的 Python 测试环境。该包测试面向 UOS、麒麟、Ubuntu、Debian 等兼容的 glibc Linux；DEB 需 Debian 系包管理工具，其他系统使用 TAR.GZ。架构、glibc 和系统共享库必须兼容，不能据此承诺所有 Linux 发行版（如 Alpine/musl）。

分组由 `core/tests/conftest.py` 给现有测试模块自动标记，标记定义在 `core/pytest.ini`。设置、Web API、路径、配置指引、应用身份和版本清单等公共测试同时进入三组；因此三组测试数之和大于全量测试数。新增测试模块必须在 `conftest.py` 归类，否则收集时明确报错。不要把某一组通过写成“全部功能通过”。

`test_distribution_profile.py` 同时进入三组：执行共用前端的完整版/S1-S2 版显示逻辑，
覆盖常规配置、操作指引、问答、按系统标签过滤更新日志（五连击解锁不扩大包的日志范围）、
Windows 双语 PE 产品属性、关于五连击密码显示入口及脚本语法；
后端用例检查错误密码、本次运行状态与设置文件不变。前端需要 Node.js，缺失时明确跳过。

## S5：逐笔统计系统（配置版）

运行 `python -m pytest -q -rs -m s5 core/tests`。`core/tests/s5/` 包含迁入的合成读取、
表达式、OOXML、条件格式测试及宿主适配测试，覆盖显式模板、只读输入、V_PRE 缺失与配对、
配置坐标错误、取消、互斥和独立状态；`test_s5_integration.py` 另验证 S5-F02/F03/F04 隐藏维护工具
默认不可见、五次点击 S5 导航后显示、再次五击隐藏并返回 F1。真实 Office 对拍在无引擎环境明确跳过；
Flask 浏览器路由另做隔离冒烟。来源独立 CLI/HTTP 层不迁入，不将这些壳层用例计为原样通过。
共用测试同时进入 S5。

## 现成配置的选择顺序

读取项目现成工作簿的测试统一使用 `config_paths.resolve_test_config()`：逐文件优先取仓库根目录 `config-real/`，对应文件不存在时才回退 `config/`。`默认配置/` 快照按同样的相对路径独立选择，不能用正在使用的正式配置替代默认快照。公开配置只保留空表头，回退后不具备正式规则的测试可能失败；失败应按实际配置来源解释，不能当作完整业务验收。

测试只读真实配置；需要删表、改规则或写入的场景先复制到临时目录。`test_config_paths.py` 进入三组，覆盖优先级、逐文件回退、默认快照及文件不变；这些测试使用临时样例文件，不依赖本机配置内容。

## S1：逐笔统计系统，5 个功能

| 功能 | 主要自动测试 | 不能由单测替代的验收 |
|---|---|---|
| S1-F01 汇总核查表校验 | `golden/test_golden_audit_flow.py`、`golden/test_golden_native_flow.py`、`test_conditional_*.py`、`test_preflight_xlsx.py`、`test_result_writer.py`、`test_history.py` | Windows Excel/WPS 条件格式渲染与公式重算；UOS LibreOffice 真机；真实核查表数量及结果对拍 |
| S1-F02 汇总校验结果说明 | `golden/test_golden_summary_flow.py`、`test_region_summary.py` | 真实说明文件与历史审核结果对拍；需要 Office 的路径须真机验证 |
| S1-F03 组合工作表（核查表） | `test_merge_org.py`、`test_external_sheet_writer.py`、`test_formula_region_writer.py`、`test_native_uos.py` | 多机构真实文件组合、公式/样式与源文件保护 |
| S1-F04 组合工作表（按配置） | `workflow/test_workflow_core.py`、`workflow/test_node_flow_config.py`、`test_combine_settings.py`、`test_native_uos.py` | 真实配置 DAG 和组合输出；目前没有独立的完整黄金流程对拍 |
| S1-F05 制作联合模板 | `test_native_template_merge.py`、`test_merge_org.py`、`test_native_uos.py` | Excel/WPS 与 LibreOffice 原生差异、命名区域及真实模板验收 |

## S2：报表采集系统，1 个功能

| 功能 | 主要自动测试 | 不能由单测替代的验收 |
|---|---|---|
| S2-F01 执行报表规则 | `test_period_v2.py`（正式 S2 包：导入、环比、外部核对、校验、导出）、`test_s2_units.py`、`test_s2_config_check.py`、`test_fast_reader.py`（快速读取引擎对拍与回退） | 真实报表与大集中数据重跑，核对三 Sheet 的行数、结论、数值和配置；目前没有专属黄金端到端对拍 |

`test_period_compare.py`、`test_percentage_indicators.py`、`test_percentage_point_compare.py` 也归在 S2 组，但测的是历史 `base_audit.period_compare` 兼容路径，**不能替代**当前 `systems/s2_report_collection` 的测试。

## S3：大集中统计系统，5 个功能

| 功能 | 主要自动测试 | 不能由单测替代的验收 |
|---|---|---|
| S3-F01 两期数据到比较文件 | `test_central_comparison.py`、`test_rule_action_engine.py`、`test_rule_applicability.py`、`test_central_indicator_rules.py`、`test_central_expression.py`、`test_central_five_segment.py`、`test_s3_rollover_dates.py`、`test_s3_comparison_config_check.py` | 真实两期数据与 VBA/黄金结果对拍，核对规则命中、备注、说明及计算过程 |
| S3-F02 比较文件到说明文件 | `test_central_explanation_exporter.py` 中说明文件导出用例 | 真实比较文件到说明文件的完整输出核对 |
| S3-F03 机构说明到说明文件 | `test_central_explanation_exporter.py` 中反馈导入、身份键与冲突报告用例 | 多机构真实反馈文件回写及冲突对拍 |
| S3-F04 本期数值核对 | `test_central_cross_period.py`、`test_s3_config_check_ui.py`、表达式相关测试 | 真实日报/月报批次与规则结果对拍；目前以组件测试为主 |
| S3-F05 比较文件到金融表单 | `test_central_forms.py`、`test_central_renderers.py`、`test_s3_forms_config_check.py`、`test_ooxml_structure.py`、`test_fast_delete_defined_names.py` | 真实金融表单的 OOXML 结构、格式和办公套件打开结果 |

## 逐功能测试用例：供人工检查

下表写的是**实际存在的代表性 pytest 用例**与其断言目标，不是新写的测试需求。路径以 `core/tests/` 为起点；同一格中后续只写 `::...` 时沿用前一个文件，只写 `::test_...` 时还沿用前一个测试类。完整清单用 `python -m pytest --collect-only -q -m s1 core/tests`（或将 `s1` 换成 `s2`、`s3`）查看。一个测试文件可能覆盖多个功能；系统分组是回归入口，不能把每项文件数当作端到端验收数。

### S1-F01 汇总核查表校验

| 核对点 | pytest 用例 |
|---|---|
| 审核流程黄金输出及重复运行 | `golden/test_golden_audit_flow.py::test_audit_flow_golden_contract`、`::test_audit_flow_rerun_repeats_contract` |
| 源文件工作表全部可在模板中找到时视为匹配；模板允许保留未出现在源文件中的表 | `test_discovery.py::test_source_sheet_subset_matches_template_with_other_sheets` |
| 待处理列表按 F1 模板匹配状态和报表模板筛选；筛选不改变勾选状态，全选/取消只作用于可见文件 | `test_distribution_profile.py::test_s1_pending_file_filters_and_visible_select_all_preserve_hidden_selection` |
| 原生 DAG 输出、日志开关不改变业务结果 | `golden/test_golden_native_flow.py::test_dag_native_flow_golden_contract`、`::test_s1_business_outputs_are_identical_when_run_log_export_changes` |
| 条件格式逐格触发、优先级、相对引用 | `test_conditional_engine.py::test_triggering_cell_reported_with_result_fields`、`::test_higher_priority_wins_over_lower`、`::test_expression_rule_triggers_with_relative_shift` |
| 没有命名区域时仍扫描“应用于”区域 | `test_conditional_format_fallback.py::test_fallback_scans_applies_to_without_named_range` |
| 模板结构差异及结果格式 | `test_preflight_xlsx.py::PreflightXlsxTests::test_structure_details_keep_sheet_name_and_cell_address`、`test_result_writer.py::test_openpyxl_result_writer_preserves_summary_contract` |

### S1-F02 汇总校验结果说明

- `test_discovery.py::test_f2_partial_hint_does_not_relax_f1_classification_or_auto_selection`、`::test_f2_hint_tracks_active_template_and_does_not_claim_fallback_matches`、`::test_f2_extra_sheets_with_fallback_are_shown_as_fallback_candidates`：模板多出的表不算未匹配；源文件新增表用中性提示，只有模板含“兜底模板”时才提示可尝试兜底。
- `test_web_app.py::WebAppCompatibilityTests::test_summary_preview_preserves_f1_unrecognized_and_manual_template`、`::test_s1_f2_file_badges_distinguish_new_sheets_from_unmatched_templates`：列表说明区分源新增表、兜底处理与无同名工作表，保持手选模板、文件选择及输入哈希。

| 核对点 | pytest 用例 |
|---|---|
| 说明文件黄金输出、无校验区场景 | `golden/test_golden_summary_flow.py::test_summary_flow_golden_contract`、`::test_summary_flow_without_validation_region` |
| 源文件缺少一个/全部模板工作表时继续生成汇总；存在表头错位时仍剔除该文件 | `test_native_uos.py::test_s1_summary_keeps_other_sheet_when_source_is_missing_one_template_sheet`、`test_preflight_xlsx.py::PreflightXlsxTests::test_summary_keeps_present_sheets_when_another_template_sheet_is_missing`、`::test_summary_keeps_header_only_result_when_all_target_sheets_are_missing`、`::test_summary_still_rejects_misaligned_existing_sheet_if_another_is_missing` |
| 只有模板含“兜底模板”时才汇总未匹配工作表；空工作表不输出；仅兜底模板可通过结构预检 | `test_native_uos.py::test_s1_summary_uses_optional_fallback_template_for_unmatched_sheets`、`test_preflight_xlsx.py::PreflightXlsxTests::test_summary_accepts_a_fallback_only_template`；不含兜底模板时的忽略行为由 `::test_s1_summary_keeps_other_sheet_when_source_is_missing_one_template_sheet` 覆盖 |
| 输出名称、历史说明、历史键、重复运行 | `test_region_summary.py::test_explanation_summary_uses_business_facing_output_name`、`::test_history_context_appends_by_builtin_rule`、`::test_history_key_matches_numeric_and_text_values`、`::test_enrich_history_columns_is_idempotent` |

### S1-F03 组合工作表（核查表）

| 核对点 | pytest 用例 |
|---|---|
| 多机构分组、工作表命名与公式保留 | `test_merge_org.py::test_organisation_key_is_first_filename_segment`、`::test_sheet_prefix_is_sanitized_and_deduplicated`、`::test_openpyxl_merge_keeps_all_sheets_and_formulas` |
| UOS/openpyxl 合并路径 | `test_native_uos.py::test_dag_combine_sheets_runs_with_openpyxl_on_linux` |

### S1-F04 组合工作表（按配置）

| 核对点 | pytest 用例 |
|---|---|
| 固定 DAG 配置不能被旧配置任意改写 | `workflow/test_node_flow_config.py::test_legacy_node_workbook_cannot_override_fixed_dag`、`::test_fixed_workflows_validate_for_com_and_native` |
| 组合预设与业务编译器绑定 | `workflow/test_workflow_core.py::test_union_combine_preset_reuses_combine_sheets_module`、`::test_business_compiler_binds_combine_plan_without_changing_module` |
| DAG 顺序、失败策略、输入校验 | `workflow/test_workflow_core.py::test_linear_graph_runs_in_order`、`::test_failure_policy_fail_workflow_stops_run`、`::test_all_default_business_workflows_have_explicit_valid_inputs` |

### S1-F05 制作联合模板

| 核对点 | pytest 用例 |
|---|---|
| 基准模板、命名区域不冲突、外部公式引用本地化 | `test_native_template_merge.py::test_base_template_is_authoritative`、`::test_named_ranges_get_sheet_suffix_and_no_collision`、`::test_cross_workbook_reference_is_localized` |
| 不覆盖原始模板、拒绝不受支持的源文件 | `test_native_template_merge.py::test_output_never_modifies_originals`、`::test_rejects_xls_and_empty_sources` |

### S2-F01 执行报表规则

| 核对点 | pytest 用例 |
|---|---|
| 三 Sheet 路由、真实 0 数值、未知类型与级别冲突 | `test_period_v2.py::test_v2_export_keeps_zero_numeric_and_routes_three_empty_sheets`、`::test_v2_export_rejects_unknown_audit_type`、`::test_v2_export_rejects_severity_conflict` |
| 进度按实际文件数及审核/导出/保存阶段推进，保存成功才到100%；有无回调业务结果及输入哈希相同；损坏文件不报告完成、导入保留来源并只取A:C | `test_period_v2.py::test_v2_progress_covers_file_counts_export_and_changes_no_results`、`::test_v2_import_progress_retains_sources_and_only_reads_contract_columns`、`::test_v2_failed_read_does_not_report_completion` |
| S2后台进度成功/失败均释放busy、失败保留完成量，已有配对免重扫且空清单重试；运行记录按阶段、文件/结果计数写入并按比例限流；前端仅保留运行记录上方进度条，同一行显示阶段/数量/耗时，无0/0和独立框；完成/失败及其他任务旧值隔离 | `test_web_app.py::S2ProgressLifecycleTests::test_worker_reports_completed_or_failed_without_refreshing_existing_pairs`、`::test_worker_logs_progress_at_bounded_intervals_with_phase_totals`、`test_distribution_profile.py::test_s2_phase_progress_reflects_actual_state_and_hides_stale_values` |
| 对比对象表头在列表内固定，相同数据/筛选复用DOM，重建保留滚动位置且筛选仍实时生效 | `test_distribution_profile.py::test_period_pair_filters_combine_clear_and_preserve_full_audit_scope`；固定表头另经浏览器滚动验证 |
| 普通数值环比与百分数百分点差 | `test_period_v2.py::test_v2_normal_amount_still_uses_relative_change`、`::test_v2_percentage_change_rate_is_percentage_point_difference` |
| 对比对象显示独立信用代码，三种导出/空表/单边保留，banks 未取得名称仍可定位，不用机构名称冒充代码，前端正确转义并保留原配对数 | `test_period_compare.py::test_pair_list_matches_empty_sheets_using_only_names`、`::test_pair_list_keeps_credit_code_for_one_sided_banks_without_org_name`、`::test_pair_list_does_not_use_institution_name_as_credit_code`、`test_distribution_profile.py::test_period_pair_targets_show_credit_codes_without_org_name_mapping` |
| 对比对象按机构/信用代码关键词及表单组合筛选，忽略大小写和首尾空格；开关/清空/无结果提示、刷新保留、选项更新和HTML转义；正式审核仍使用全部对比对象 | `test_distribution_profile.py::test_period_pair_filters_combine_clear_and_preserve_full_audit_scope` |
| 配对状态筛选全部/单边/本期缺少/上期缺少/两期齐全，pre表示本期缺少、cur表示上期缺少，缺失位置明确期别；支持四条件组合、刷新保留和统一清空，不改变审核数据 | `test_distribution_profile.py::test_period_pair_status_filter_identifies_the_missing_period`、`::test_period_pair_filters_combine_clear_and_preserve_full_audit_scope` |
| 三种导出方式只按文件名/工作表名配对；xlsx/xls 空表保留、多机构不合并、预览不扫描正文，空表/错误正文身份不影响正式环比结果的机构字段且输入不变 | `test_period_compare.py::test_pair_list_matches_empty_sheets_using_only_names`、`::test_pair_list_does_not_open_worksheet_xml_or_shared_strings`、`::test_pair_list_keeps_multiple_empty_institutions_and_true_one_sided_sheet`、`test_period_v2.py::test_v2_import_identity_comes_from_names_when_body_is_blank_or_wrong`、`::test_v2_import_keeps_filename_credit_code_when_name_profile_has_other_code`、`::test_v2_export_preserves_identity_for_completely_empty_counterpart` |
| banks 导出提示依赖机构参照，建议重新导出默认单sheet/按报表单sheet；本期/上期任一选择触发、换目录清除，执行确认同步提醒 | `test_web_app.py::WebAppCompatibilityTests::test_banks_selection_warns_and_recommended_source_clears_warning`、`test_distribution_profile.py::test_period_banks_warning_renders_and_reaches_execution_confirmation` |
| 外部核对容差、缺失值、六种比较方式 | `test_period_v2.py::test_v2_external_rule_uses_its_own_tolerance_and_rule_identity`、`test_s2_units.py::test_external_missing_value_is_not_treated_as_real_zero`、`::test_external_comparison_methods_use_report_value_as_left_operand_and_test_boundaries` |
| 外部核对区分未匹配报表项目、缺少指标列和已匹配但值为空；匹配行的空白/非数值按0参与比较并生成软性提示；行名称空且指标有值时提示并忽略，时序元数据存在但指标为空时忽略不提示，仅内部整行全空行提示；提示写入结果“校验规则”且完成状态显示数量 | `test_period_v2.py::test_v2_external_export_shows_selected_report_item_next_to_credit_code`、`::test_v2_external_time_series_blank_matched_values_are_zero_with_soft_warning`、`::test_v2_external_plain_layout_treats_blank_value_as_zero_with_soft_warning`、`::test_v2_external_non_numeric_value_is_zero_and_warned`、`::test_v2_external_blank_row_label_is_ignored_with_soft_warning`、`::test_v2_service_exports_external_import_warnings_as_soft_findings`、`test_s2_units.py::test_external_missing_value_is_not_treated_as_real_zero`、`::test_external_uses_explicit_zero_fallback_when_institution_label_is_missing` |
| 外部核对结果在信用代码右侧显示报表项目；外部映射优先、CF2回退、已映射缺数仍保留、无映射留空、真实0/标签缺失兜底、时序匹配及文本标签完整导出 | `test_period_v2.py::test_v2_external_export_shows_selected_report_item_next_to_credit_code`、`::test_v2_service_exports_current_time_series_value_without_modifying_inputs`、`test_s2_units.py::test_external_mapping_precedes_cf2_and_cf2_is_used_as_fallback` |
| 外部数据按表头识别普通/时序布局；时序要求“数据日期”与“行名称”同时出现，纯报表兼容原空白标签头和显式“行名称”；时序仍只匹配本期完整日期，不借用其他日期，缺失/无效日期行提示后忽略，重复记录阻断；空白和非数值值按0并提示，原文件不变 | `test_period_v2.py::test_v2_external_import_supports_xlsx_only`、`::test_v2_external_import_time_series_uses_named_headers_and_full_dates`、`::test_v2_external_import_time_series_rejects_ambiguous_or_invalid_data`、`::test_v2_external_time_series_invalid_row_date_is_warned_and_ignored`、`::test_v2_external_time_series_matches_exact_current_date_and_preserves_missing_values`、`::test_v2_external_time_series_blank_matched_values_are_zero_with_soft_warning`、`::test_v2_external_plain_layout_treats_blank_value_as_zero_with_soft_warning`、`::test_v2_external_non_numeric_value_is_zero_and_warned`、`::test_v2_external_blank_row_label_is_ignored_with_soft_warning`、`::test_v2_service_exports_current_time_series_value_without_modifying_inputs` |
| 金额单位统一后比较、保持小数精度 | `test_s2_units.py::test_amount_comparison_converts_both_sides_to_output_unit_and_keeps_tolerance_in_that_unit`、`::test_rule_expression_preserves_decimal_digits_beyond_excel_float_precision` |
| 正式配置检查、环比策略区间与外部规则字段 | `test_s2_config_check.py::test_formal_and_default_s2_configs_pass_read_only_check`、`::test_checker_blocks_gap_and_duplicate_strategy`、`::test_checker_requires_external_rule_header_order` |
| 日志设置不改变业务输出 | `test_period_v2.py::test_s2_business_output_is_independent_of_unrelated_log_preferences` |
| 快速读取引擎与传统引擎逐值对拍（xls 对齐 xlrd、xlsx 对齐 openpyxl），缺包静默回退、单文件失败提示后重读，两引擎导入结果全量一致 | `test_fast_reader.py::test_fast_reader_values_match_legacy_engine`、`::test_fast_reader_absent_falls_back_silently`、`::test_fast_reader_failure_falls_back_with_warning`、`::test_import_results_are_identical_between_engines` |

### S3-F01 两期数据到比较文件

| 核对点 | pytest 用例 |
|---|---|
| 比较文件表头、行数、值、备注、说明与计算过程黄金对拍 | `test_central_comparison.py::ComparisonGoldenTests::test_headers_match_golden`、`::test_row_count_matches_golden`、`::test_values_match_golden`、`::test_explain_and_process_match_golden` |
| 12MU2 一类的实际变动值判断、金额阈值 | `test_central_indicator_rules.py::test_change_action_requires_actual_change_even_without_threshold`、`::test_change_action_uses_change_amount_not_current_balance` |
| 结转规则与表达式只在结转日生效 | `test_rule_action_engine.py::test_real_config_all_rollover_scene_actions_are_date_gated`、`test_s3_rollover_dates.py::test_real_config_expression_groups_follow_date_gate` |
| 命中说明含规则编号、阈值、备注；空备注不产生尾部分隔符 | `test_rule_action_engine.py::test_action_output_includes_id_action_effective_threshold_and_config_note`、`::test_action_output_omits_empty_config_note_without_trailing_separator` |
| 五段/八段语法、配置动作检查 | `test_central_five_segment.py::TestGoldenParity::test_sbe_parity_row_by_row`、`test_s3_comparison_config_check.py::test_unsupported_enabled_action_reports_sheet_row_rule_and_original_action` |

### S3-F02 比较文件到说明文件

| 核对点 | pytest 用例 |
|---|---|
| 只导出标记为需要说明的行 | `test_central_explanation_exporter.py::ExplanationExporterTests::test_exports_only_rows_marked_for_explanation` |
| 机构精确/通配参照和源地区保留 | `test_central_explanation_exporter.py::ExplanationExporterTests::test_uses_precise_then_wildcard_org_reference_and_keeps_source_region` |

### S3-F03 机构说明到说明文件

| 核对点 | pytest 用例 |
|---|---|
| 按四字段业务身份匹配反馈 | `test_central_explanation_exporter.py::ExplanationExporterTests::test_imports_feedback_by_four_business_identity_columns` |
| 同名指标按代码区分、冲突报告保留来源与原因 | `test_central_explanation_exporter.py::ExplanationExporterTests::test_keeps_same_named_indicators_separate_by_indicator_code`、`::test_writes_conflict_report_with_feedback_source_and_reason` |

### S3-F04 本期数值核对

| 核对点 | pytest 用例 |
|---|---|
| 五段指标引用、频度与前提条件 | `test_central_cross_period.py::TokenTests::test_five_segment_token_joins_uni_code`、`::DataStoreTests::test_frequency_check`、`::RunCrossPeriodTests::test_premise_false_skips` |
| 阈值触发、取反、缺一侧与公式错误 | `test_central_cross_period.py::RunCrossPeriodTests::test_ratio_rule_hits_and_suppresses`、`::test_inverted_rule`、`::test_one_side_missing_message`、`::test_formula_error_still_outputs_and_logs` |
| 3.2 配置必需 Sheet 与语法检查 | `test_s3_config_check_ui.py::test_formal_32_check_reports_selected_workbooks_sheet_and_syntax_scope`、`::test_32_check_reports_a_malformed_enabled_rule_from_a_disposable_copy` |

### S3-F05 比较文件到金融表单

| 核对点 | pytest 用例 |
|---|---|
| 按机构回填及内嵌模板使用 | `test_central_forms.py::RenderFormsTests::test_renders_one_file_per_org_with_values_and_colors`、`::test_embedded_config_is_used_as_template_and_not_emitted` |
| FAST_OOXML 与 openpyxl 的表单、数值、格式、颜色、隐藏行一致 | `test_central_renderers.py::DualModeParityTests::test_sheet_names_and_order_match`、`::test_cell_values_match`、`::test_number_formats_match`、`::test_warning_fills_match`、`::test_hidden_rows_match` |
| 3.3 配置检查及内嵌表单引用 | `test_s3_forms_config_check.py::test_formal_config_passes_and_is_not_modified`、`::test_missing_referenced_embedded_form_sheet_is_reported` |

### S5-F01 汇总核查表校验

- `s5/test_s5_recognition.py`：模板、业务与报告期识别；日期冲突/缺失、模板/业务歧义明确提示；源工作表全部包含于模板时视为匹配，模板多出的表不降级，源工作簿多出的表单独标注；源目录输出跟随与固定。
- `s5/test_s5_integration.py::test_complete_f1_explicit_template_source_without_names_final_cf_and_read_only`：源表无命名区域时使用显式模板；批次与单机构结果保留条件格式；输入哈希不变。
- `::test_vpre_without_previous_is_visible_and_matching_previous_is_retained`：缺上期明确报告取数失败，提供上期后按机构配对。
- `::test_external_code_mapping_runs_complete_f1_with_both_layouts_and_readers`：普通/时序外部布局按本期报告期取数，并验证 Calamine 与强制 openpyxl 回退。
- `s5/test_s5_configured_engine.py::test_external_numeric_blank_or_text_is_zero_with_a_quality_warning`、`s5/test_s5_integration.py::test_external_layout_anomalies_warn_and_continue_for_both_layouts`：即使外部配置含固定数据源契约，也按普通/时序布局读取；匹配值空白/非数值按0，有值但行名称为空时提示并忽略；只有整行全空的内部行提示，时序元数据本身不触发行名称异常；批次与单机构结果写入“数据质量提示”，任务结束展示提示数量，时序日期仍按本期精确匹配。
- `s5/test_s5_integration.py::test_s5_maintenance_tools_are_hidden_by_default_and_toggle_with_five_nav_clicks`：S5 主面板 F1-F4 按钮位于同一卡片；F2-F4 默认隐藏并由左侧 S5 导航五击切换，空闲说明卡片不显示。
- `::test_time_series_missing_current_date_only_fails_affected_rule_and_continues_batch`：某机构缺少本期有效日期时，对应规则明确失败，批次继续审核其他机构。
- `::test_s5_maintenance_uses_the_current_panel_paths`、`::test_s5_vext_reference_export_supports_both_external_layouts`：三个维护工具沿用 S5 面板模板/外部文件路径，V_EXT 参考同时支持普通与时序外部布局。
- `::test_configuration_error_reports_physical_sheet_cell_and_rule`、`::test_no_enabled_rules_is_not_reported_as_pass`：配置物理坐标与无启用规则诊断。
- `::test_cancellation_stops_before_next_institution_and_never_returns_success`：实际停止后续机构，不把取消报告为成功。

- `::test_s5_start_freezes_parameters_before_thread`、`::test_s5_state_does_not_write_s1_fields_and_mutual_exclusion`：参数快照、状态隔离与执行互斥。

## 后续执行口径

1. 改动某一功能：先跑对应测试文件，再跑所属系统标记。跨系统公共代码：跑所有受影响系统标记。
2. 准备提交或合入 `main`：跑 `compileall`、相关系统标记和 core 全量测试；前端改动另跑 JavaScript 语法检查。
3. 涉及真实 Office 计算、条件格式或跨平台行为：额外跑相应 Windows/UOS 真机验收。涉及规则迁移：额外做配置可执行性检查与真实数据/VBA 对拍。
4. 报告每组的 passed、skipped、failed、errors，以及未执行的真机/真实数据项目。`0 skipped` 仅表示本次 pytest 未跳过，**不表示**所有平台或业务数据都已验收。
