# Open-Data-Audit 功能迁移清单

## 基线与状态口径

- 只读源项目：C:\Users\AZI\Desktop\VibeCode\3.open-data-audit
- 当前功能目录依据源项目 core/tests/TEST_GROUPS.md 与源项目 AGENTS.md 建立；本清单不是对全部实现细节的最终代码审计。
- 来源文档列出的 S1/S2/S3/S5 活跃业务项作为本轮清单；S4 在源项目中标为预览流程，是否纳入正式迁移仍待确认。
- S5-F02/F03/F04 在来源测试文档中有隐藏维护工具测试，但不在主要功能汇总表内；需继续沿源码确认其业务行为，不能仅凭隐藏状态视为废弃。
- workspace fixtures 只包含合成或脱敏测试数据，没有复制真实业务工作簿；真实配置和 2026-07-31 S1 文件仅从只读源目录做过一次模板推荐 smoke 对拍，报告只保留文件集合哈希和汇总结果。该 smoke 不能替代真实模板全集验收。
- 代码盘点（2026-10-23）：`audit-core/src/`、`pytools-core/src/`、`crates/` 和 `sidecar/` 没有 Python 源文件或 Python 子进程入口；workspace 当前 50 个 `.py` 文件仅位于 `probe/`、`tests/compare/`，作为旧实现基线和测试辅助工具保留，不进入产品运行路径。workspace 尚无 `.go`/`go.mod` 生产实现；`core-go/` 目前只有说明文档。
- 每项状态初始为“未验证”；只有逐项实现、对拍并保存报告后才能改为“保留”“丢失”“损坏”“不适用”或“通过”。

## 功能清单

| 来源标识 | 功能 | 需要覆盖的业务结果 | 当前状态 |
|---|---|---|---|
| S1-F01 | 汇总核查表校验 | 输入发现与筛选、模板匹配、规则执行、结果回写、历史记录及重复运行一致性 | 未验证；模板推荐局部已接 Rust CLI，4 个合成场景对拍通过，真实 config-real + 2026-07-31 样本树 smoke 对拍通过；条件格式/业务规则执行、公式结果、审核结果写入、历史和重复运行仍未迁移 |
| S1-F02 | 汇总校验结果说明 | 汇总说明、未匹配/兜底行为、历史说明与输出结构 | 未验证；待端到端样本 |
| S1-F03 | 组合工作表（核查表） | 多来源工作表合并、命名、顺序、公式和命名区域保留 | 未验证；待真实组合文件 |
| S1-F04 | 组合工作表（按配置） | 按配置/DAG 执行组合、输入验证和失败策略 | 未验证；待配置契约和黄金输出 |
| S1-F05 | 制作联合模板 | 多模板合并、基准模板选择、名称冲突处理和引用修正 | 未验证；待模板组合样本 |
| S2-F01 | 执行报表规则 | 报表导入、环比、外部核对、规则校验、三类结果导出 | 未验证；Rust service/CLI 对拍通过，loopback Web bridge 可由原样 HTML 调用并完成 synthetic 端到端运行，目录选择已接入两期文件配对预览；真实模板、完整输入/导出错误矩阵、目标平台和 deb 仍未验收 |
| S3-F01 | 两期数据到比较文件 | 两期数据匹配、规则命中、比较计算、说明及结果输出 | 未验证；待两期真实脱敏数据 |
| S3-F02 | 比较文件到说明文件 | 将比较结果转换为说明文件并保持输出契约 | 未验证；待比较文件和黄金说明 |
| S3-F03 | 机构说明到说明文件 | 机构反馈读取、身份键匹配、冲突处理和说明回写 | 未验证；待多机构反馈样本 |
| S3-F04 | 本期数值核对 | 本期数据核对规则、跨期关联和异常报告 | 未验证；待真实日报/月报批次 |
| S3-F05 | 比较文件到金融表单 | 比较结果写入金融表单并保留表单结构、格式和名称引用 | 未验证；待金融表单样本 |
| S5-F01 | 汇总核查表校验（逐笔统计配置版） | 配置驱动读取、逐笔规则、校验结果与输出 | 未验证；待核对与 S1 的行为差异 |
| S5-F02 | 隐藏维护工具（来源测试标识） | 从源代码确认实际业务功能、输入输出和触发方式 | 待确认；不能因默认隐藏而跳过 |
| S5-F03 | 隐藏维护工具（来源测试标识） | 从源代码确认实际业务功能、输入输出和触发方式 | 待确认；不能因默认隐藏而跳过 |
| S5-F04 | 隐藏维护工具（来源测试标识） | 从源代码确认实际业务功能、输入输出和触发方式 | 待确认；不能因默认隐藏而跳过 |

## 格式与共用能力矩阵

| 输入/对象 | 当前处理计划 | 验收状态 |
|---|---|---|
| xlsx | calamine 读取；OOXML 补充缺失元数据；新建写入用 rust_xlsxwriter；原位编辑候选 umya | 合成探针已有报告；业务格式与真实模板未验证 |
| xlsm | 逐项验证公式和宏相关部件；生产不执行 VBA；不得仅凭 OOXML 容器判断保真 | 未验证 |
| xls | calamine 读取能力逐项验证；需改 sheet 名等不支持操作时允许显式转换为 xlsx | 未验证 |
| xlsb | 逐项验证 calamine 值读取；其余元数据缺口按适配器能力处理 | 未验证 |
| csv | Rust `csv` + `encoding_rs` 已迁移 S3 固定 14 列导数的 CSV 解析/标准化；保留 GBK 优先、UTF-8 BOM 回退、引号字段、空记录行号、日期/数值及异常值语义 | 3 个合成 Python 对拍场景通过；仅覆盖 S3 CSV 路径，通用 CSV 契约及 XLS/XLSX 导入仍未验证 |
| et | 经用户显式选择后由 sidecar 转成 xlsx；验证 sheet 清单和数据 | 未验证 |
| 批注/条件格式/数据验证 | calamine 未提供时由 OOXML 或 umya 补；元数据缺失必须报告未验证 | 合成样例仅覆盖已有探针列出的对象 |
| 公式重算 | 默认不运行；用户显式要求本轮结果时调 sidecar 并检查缓存刷新 | sidecar 专项尚未完成 |
| 打印 | 仅迁移页面设置写入；PDF、渲染不在范围 | 页面属性逐项待验收 |

## 进入实现前的门槛

1. 对每个功能找到源代码入口、直接调用链、配置字段、输入/输出和错误行为。
2. 将需要保留的 workbook parts 和公式口径登记到独立 probe config。
3. 取得真实脱敏模板；无样本时可先做合成探针，但保持真实模板状态为未验证。
4. 完成共享 reader/writer/editor 对拍并记录决策后，才开始该功能的业务迁移。
5. 每次状态变化在 audit-core/GPT修改记录/ 增加日期报告；不以单测或合成文件推断真实模板通过。

## 已对拍的局部迁移（不代表业务流程迁移）

| 来源 | Rust 实现 | 状态 | 证据与范围 |
|---|---|---|---|
| `base_audit.discovery.recommend_template` 的按工作表集合模板推荐 | `audit_core::s1_template::{recommend_template, recommend_from_sheet_sets}`，经 `open-data-audit template-suggest` CLI 调用并用共享 namespace-aware OOXML sheet-directory adapter 读取工作表名 | 保留（模板推荐局部） | `probe/reports/template_recommendation_parity.json`：4 个合成场景逐字段匹配只读 Python 基线，覆盖源工作表为模板子集、混合报表类型停止自动选择、候选接近时要求人工确认、无可信模板。`probe/reports/template_recommendation_real_smoke.json`：只读源 `config-real` 的 8 个模板和 2026-07-31 S1 目录中的 68 个 `.xlsx` 源文件，Rust 与 Python 的匹配状态、置信度、说明和候选一致，置信度 1；读前/读后模板和源文件集合 SHA-256 未变。该单次 WSL2 目录树性能观察为 Rust Release OOXML metadata adapter 中位数 1.37 秒、Python 基线 1.76 秒；同一任务使用 Calamine 的 Rust Release 中位数 3.70 秒。此数字只描述该样本和完整进程耗时，不外推到其他模板、目录、引擎任务或目标平台。仅是单个真实目录树的模板推荐；不覆盖真实配置全部模板、其他期间、S1 规则计算/条件格式、公式重算、输出回写、历史、重复运行、Windows 原生运行、麒麟/UOS 或 deb，S1-F01 仍未验证。 |
| `base_audit.discovery._periods_in_text` | `audit_core::discovery::periods_in_text` | 保留（局部） | `probe/reports/period_parser_parity.json`：49/49 个 Python 基线向量一致；仅覆盖纯文本期间解析，不覆盖目录扫描、`detect_period` 或任何端到端业务流程。 |
| `base_audit.discovery.source_workbooks` + `detect_period` | `audit_core::discovery::{source_workbooks, detect_period}` | 保留（局部） | `probe/reports/period_detection_parity.json`：9/9 个合成目录场景与 Python 基线一致，覆盖递归深度、输入过滤、文件名识别、目录回退及冲突输出；不覆盖真实模板、文件系统错误策略或端到端业务流程。 |
| `base_audit.period_compare.list_period_pairs` + `base_audit.report_source.report_source_warning` | `audit_core::period_pairs::{list_period_pairs, report_source_warning}` + `audit_core::report_source::report_source_identity`；工作表名经共享 `excel_reader::read_sheet_names` 使用 Calamine | 保留（前端预览局部） | `probe/reports/period_pairs_parity.json`：1 个合成目录场景、6 个配对行及 banks 提醒与只读 Python 基线逐字段一致；覆盖标准逐机构文件、reports、banks、多工作表、临时文件过滤、重复身份最后项规则和单边项目，路径按文件名归一。工作表名只取 Calamine workbook sheet directory，不读取单元格正文。没有有效 `.xls` fixture，故该格式的 sheet-name 读取仍未验证；不覆盖真实报送目录或端到端业务审核。 |
| `base_audit.models` 的 10 个 dataclass 与两个结果摘要方法 | `audit_core::models` | 保留（模型层） | `probe/reports/models_parity.json`：10 个模型类型、11 个实例、7 组默认值和结果计数/摘要文本与 Python 基线一致；不代表任何审核、预检、模板或文件写入流程已经迁移。 |
| `s3_central_statistics.csv_importer` 的 CSV 解码与数据标准化分支 | `audit_core::central_csv::parse_central_csv` | 保留（S3 导入局部） | `probe/reports/central_csv_parity.json`：3 个合成 CSV 场景逐字段与只读 Python `read_central_csv` 基线一致，覆盖 GBK、UTF-8 BOM、带逗号/双引号字段、空记录行号、日期解析、单位换算/豁免、异常值、短行补齐和重复键后值覆盖；不覆盖 XLS/XLSX reader dispatch、真实批次和端到端 S3 流程。 |
| `base_audit.template` 的模板名、批注、公式结果和规则行纯解析 helper | `audit_core::template::{normalize_template_name, parse_comment_rule_id, clean_rule_comment, parse_formula_result, parse_rule_rows}` | 保留（纯解析局部） | `probe/reports/template_parser_parity.json`：24 个 fixture case 与只读 Python 基线逐项一致（模板名 6、批注 7、公式结果 7、规则行 4）；覆盖 Unicode 15.0 十进制数字的日期/区域标记。只覆盖值级 helper，不覆盖 workbook/批注读取、ExcelSession、规则执行或审核流程；JSON 数组/对象的 Python `str()` 等价性未验证。 |
| `base_audit.name_config` 的功能映射数据类及纯值 helper | `audit_core::name_config::{FeatureMapping, split_names, is_limited, features_of_type, matches_named_range}` | 保留（纯配置 helper 局部） | `probe/reports/name_config_helpers_parity.json`：31 个 fixture case 与只读 Python 基线逐项一致（映射 4、名称拆分 5、限定标记 9、按类型筛选 3、命名区域匹配 10）。不覆盖 config.xlsx 读取/校验/保存、默认模块/流程生成或工作簿行为；Python `casefold()` 的特殊 Unicode 展开字符未验证。 |
| `systems.s2_report_collection.unit_conversion` 的金额类型标记、Decimal 解析、单位缩放与 `to_yuan` 兼容入口；`importer._normalize_fast_value` / `_decimal_number` 的单元格数值适配 | `audit_core::unit_conversion::{is_amount_data_type, as_decimal, convert_amount, to_yuan, decimal_number_from_cell, convert_amount_from_cell}`；共享类型 `excel_reader::CellValue` | 保留（金额单位 helper 与 typed-cell adapter） | `probe/reports/unit_conversion_parity.json`：62 个 fixture case 与只读 Python 基线逐项一致，包括 43 个既有文本/helper case 和 19 个 reader cell case；reader cell case 直接对拍源 importer 的 XLS/XLSX 归一化、Decimal tuple 和单位转换结果。Calamine 值的共享 DTO 保持原 `WorkbookDump` JSON 结构。S2 目录导入调用该适配器的范围另见下一行；此 helper 报告不覆盖真实工作簿或规则/周期/外部比较流程。 |
| `systems.s2_report_collection.config.{load_config, load_unit_settings}` 的 S2 运行配置工作簿解析 | `audit_core::s2_config::{load_config, load_unit_settings}`，通过共享 `excel_reader::read_value_columns_with_iso_dates` 使用 Calamine；配置可投影到现有导入/周期/规则/外部比较模型 | 保留（合成 XLSX 配置加载局部） | `probe/reports/s2_config_parity.json`：15 个合成 XLSX 配置场景逐字段/错误文本与只读 Python 基线一致，包含 3 个成功配置、12 个预期配置错误，以及独立单位设置读取；覆盖设置默认值、表头/别名、指标、机构、外部规则、环比策略/旧环比表、表达式规则、告警与错误。源 `config.py`、`models.py`、`unit_conversion.py` SHA-256 记录在报告中。配置检查器局部切片见下一行；默认工作簿生成、真实配置文件、服务编排和报告输出未迁移；S2-F01 仍未验证且未接入生产默认路由。 |
| `systems.s2_report_collection.config_check.{check_s2_config, format_config_report}` 的只读配置结构、策略区间和运行时解析检查 | `audit_core::s2_config_check::{check_s2_config, format_config_report}`；结构读取经共享 `excel_reader::read_value_columns_with_iso_dates`（包括 Calamine worksheet extent 供统计），运行时检查复用 `s2_config::load_config` | 保留（合成 XLSX 配置检查器局部） | `probe/reports/s2_config_check_parity.json`：19 个合成 XLSX 案例的结构化报告与 UI 文本均与只读 Python 基线逐项一致，含有效配置、旧审核级别兼容警告、策略空档/重复、日期值错误提示、完整综合错误和必需工作表缺失；对拍同时验证输入夹具字节不变。Python 基线记录 `config_check.py`、`config.py`、`models.py`、`unit_conversion.py` 的 SHA-256。未覆盖真实业务配置、损坏/加密工作簿的错误文案、服务阻断流程接线、默认工作簿生成、S2 报告输出、非 XLSX 格式或目标 Linux 平台；S2-F01 仍未验证且未接入生产默认路由。 |
| `systems.s2_report_collection.importer.import_period_directory` 的文件发现、来源身份、A-C 列指标行提取、机构/指标补充、金额单位归一、去重和警告 | `audit_core::s2_import::import_period_directory`，通过共享 `excel_reader::read_value_columns` 使用 Calamine | 保留（合成 XLSX 目录导入局部） | `probe/reports/s2_import_parity.json`：12 条记录、2 对重复项、3 条警告与只读 Python importer 结果逐字段相同；夹具覆盖标准名/按机构名映射、两张表单、余额/累发/百分数/个数/文字/零值、无效金额、未识别文件跳过及大于 `i64::MAX` 的整数指标代码归一。只覆盖合成 `.xlsx` 夹具；有效 BIFF `.xls`、`.xlsb`、真实工作簿、公式缓存边界、文件读取失败策略、进度回调、配置加载、周期/规则/外部比较、报告写入及服务路由均未验证，尚未接入生产默认路径。 |
| `systems.s2_report_collection.importer.import_external_workbook` 的外部工作簿解析 | `audit_core::s2_external_import::import_external_workbook`，通过共享 `excel_reader::read_value_columns_with_iso_dates` 使用纯 Rust Calamine | 保留（合成 XLSX 外部导入局部） | `probe/reports/s2_external_import_parity.json`：12 个合成输入场景与只读 OpenPyXL 基线逐字段/错误文本一致；覆盖普通/时序布局、1900/1904 日期系统、日期与数值归一、空值/非数值警告、机构参照表优先级/旧表兼容、重复日期行/重复指标/不完整表头和不支持扩展名拒绝。基线 Python 3.12.10，源 `importer.py` SHA-256 为 `EEF44740DE8EB7AB0E09DE7C074FF06F22F515CC9A1518982FA662D123A49BEC`。未覆盖真实外部工作簿、成功解析 `.xls/.xlsb/.et/.csv`、S2 服务编排/外部比较集成/报告写入、性能及目标平台；尚未接入生产默认路由，S2-F01 仍未验证。 |
| `systems.s2_report_collection.validation.{validate_import, validate_config_warnings, validate_external_warnings}` | `audit_core::s2_validation::{validate_import, validate_config_warnings, validate_external_warnings}` | 保留（质量发现转换局部） | `probe/reports/s2_validation_parity.json`：11 条 Python `AuditFinding` 逐字段对拍通过，覆盖重复项/跨表单冲突、缺少机构或表单身份、导入警告、配置警告及外部数据警告；路径仅为跨平台比较归一为文件名。发现生成已迁移，S2 规则引擎、配置工作簿加载、外部导入及报告写入未验证，不代表 S2-F01 流程完成。 |
| `systems.s2_report_collection.exporter.build_summary` 的报告统计数据生成 | `audit_core::s2_summary::build_summary` | 保留（报告统计数据局部） | `probe/reports/s2_summary_parity.json`：2 个合成案例与只读 Python `build_summary` 结果一致，覆盖唯一机构/非空表单计数、指标及发现总数、异常发现分类、空类型/空级别和空输入；源 `exporter.py` SHA-256 为 `EF99D056BA9038BC9AED8BD23D5D8D0D2813941BA54366C1136380AC70A2E02B`。只覆盖摘要聚合，不覆盖 S2 服务编排、XLSX 报告输出、真实模板、性能或目标平台；S2-F01 仍未验证且未接入生产默认路由。 |
| `systems.s2_report_collection.period_engine.PeriodComparisonEngine` 的两期比较计算 | `audit_core::s2_period::compare_periods` | 保留（规则计算局部） | `probe/reports/s2_period_parity.json`：13 个合成规则场景与只读 Python 基线逐字段一致，产生 15 条发现并覆盖 2 个预期规则配置错误；覆盖数值/文字/缺失项、百分数百分点、Decimal 金额、策略区间、严重度、计算轨迹和错误分支。基线源文件 SHA-256 记录在报告中。该切片使用已导入的 `IndicatorRecord` 和合成 JSON 配置，不覆盖配置工作簿加载、报表导入、外部比较、规则表达式、服务编排、报告写入或真实模板，S2-F01 仍未验证且未接入生产默认路由。 |
| `systems.s2_report_collection.rule_engine.RuleEngine` 的累计不降与表达式规则计算 | `audit_core::s2_rule_engine::evaluate_rules` + `audit_core::s2_expression` | 保留（规则计算局部） | `probe/reports/s2_rule_engine_parity.json`：13 个合成案例与只读 Python 基线逐字段一致，产生 10 条规则发现和 3 条规则质量发现；覆盖累计不降、当前/上期占位符、旧八段指标码、Decimal 表达式、逻辑函数、Thd 金额单位换算、缺失/无效值诊断及来源/计算轨迹。`rule_engine.py`、`models.py`、`unit_conversion.py` 的 SHA-256 记录在报告中。仅覆盖已导入记录和内存配置，不覆盖配置工作簿加载、导入、外部比较、服务编排、报告写入或真实模板；S2-F01 仍未验证且未接入生产默认路由。 |
| `systems.s2_report_collection.service.PeriodAuditV2Service.run` 的已导入数据分析与发现汇总编排（不含文件/配置 I/O 与导出） | `audit_core::s2_analysis::analyze_imports` | 保留（内存分析编排局部） | `probe/reports/s2_analysis_parity.json`：2 个合成案例与只读 Python 服务组件顺序逐字段一致，覆盖本期首条逻辑键保留、环比/外部核对/表达式规则发现、配置/导入/外部警告、规则质量发现、发现拼接次序和摘要，以及无外部数据的空输入分支。Python 基线记录 `service.py`、`period_engine.py`、`external_engine.py`、`rule_engine.py`、`validation.py`、`exporter.py`、`models.py` 和 `importer.py` SHA-256。未覆盖配置检查/加载、工作簿导入、进度回调、XLSX 导出、真实模板、性能、目标平台或生产路由；S2-F01 仍未验证。 |
| `systems.s2_report_collection.exporter.ExcelExporter.export` 的三表 XLSX 报告写出候选 | `audit_core::s2_exporter::export_s2_report`，通过共享 `excel_writer` 新建工作簿、共享 `excel_reader` 重开校验 | 保留（合成 finding 到 XLSX 的局部候选；由 S2 service/CLI 候选调用，非生产默认） | `probe/reports/s2_exporter_parity.json`：5 个合成 finding、3 张工作表、116 个比较单元格（11 个数值格、1 个公式格），与只读 Python `ExcelExporter` 的文件级对拍 0 差异；数值相对容差 `1e-15`，公式缓存为空。对拍覆盖值/类型、公式与公式/文字区分、样式、列宽区间、冻结窗格、筛选、隐藏行和页面边距。`probe/reports/s2_exporter_python_baseline.json` 记录 Python 3.12.10、fixture SHA-256，以及源 `exporter.py` / `models.py` SHA-256。仅覆盖合成 finding；日期/时长值和不能精确转为 Excel `f64` 的整数明确拒绝。真实脱敏模板、性能、目标 Linux 或 deb 未验证；S2-F01 仍未验证。 |
| `systems.s2_report_collection.service.PeriodAuditV2Service.run` 的文件级 S2 编排（配置检查/加载、两期目录导入、期间/规则/可选外部比较、发现汇总、三表报告写出） | `audit_core::s2_service::run_s2_audit`，由 `audit-core/src/bin/open-data-audit.rs` 的 `open-data-audit s2` 接入 Rust CLI，连接已对拍的 `s2_config_check`、`s2_config`、`s2_import`、`s2_period`、`s2_rule_engine`、`s2_external_import`、`s2_external`、`s2_validation`、`s2_analysis`、`s2_exporter` | 保留（synthetic service + CLI 端到端候选；非生产默认） | `probe/reports/s2_service_parity.json`：无外部数据 15 条发现、含外部 `plain_layout.xlsx` 18 条发现；服务 API 与 CLI 的摘要/进度均和只读 Python 对拍 0 差异，CLI 生成的三表 XLSX 文件级比较 0 差异（分别比较 268/304 个单元格、36 个数值格）；阻断配置错误和进度一致，且不创建输出目录。相对容差 `1e-15`；Windows/WSL 绝对路径记法归一后比较；Python/Rust 输入哈希不变，Python 基线源哈希前后不变。夹具为合成文件，当前/上期复用相同报表内容并只调整文件名日期；外部读取进度中的引擎/耗时动态字段不逐字比较。未验证真实脱敏模板、全部输入/导出失败回退、性能、Windows 原生 Rust、麒麟/UOS/glibc 2.28 或 deb；S2-F01 整体仍未完成。 |

其他功能清单状态不因这些局部迁移对拍而改变，仍按上文逐项保持未验证/待确认。

## Rust 前端接入状态

- 源项目 `core/frontend/web/index.html` 和 `core/frontend/assets/app-icon.{png,ico}` 已原样复制到 `frontend/`；HTML 与源文件 SHA-256 一致。源项目只读，没有修改。
- `open-data-audit serve` 仅监听 `127.0.0.1`，并通过响应时注入 bridge 保持 HTML 文件本身不变。页面启动后导航至已接入的 S2。
- 当前连接 `get_state`、浏览器路径浏览、S2 配置检查、S2 路径选择/清除、输出目录固定/跟随、审核启动、进度和结果回传，以及有限的 UI 设置更新。选择或清除本期/上期目录会刷新两期文件配对预览，并支持 `.xls`/`.xlsx` 文件名列举。
- 未迁移接口返回 HTTP 501。取消和设置持久化尚未接入；S1/S3/S4/S5 的前端入口不表示其功能已经迁移。
- 合成 S2 输入经 HTTP bridge 完整运行并生成三表 XLSX；该结果不替代真实业务模板、Windows 原生、麒麟/UOS、glibc 2.28 与 deb 验收。
