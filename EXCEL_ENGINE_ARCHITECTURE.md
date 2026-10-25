# Excel 引擎架构与迁移清单

> Excel 读写、修改与引擎路由的规范入口：[rules/13_Rust_Excel读写修改与引擎路由规范.md](rules/13_Rust_Excel读写修改与引擎路由规范.md)。本文件记录实际代码、候选能力矩阵和待完成迁移项。

盘点日期：2026-10-23  
状态：架构设计、静态源码盘点，以及通用 ExcelEditPlan/ExcelModifier 路由与安全输出骨架；新增隔离 openpyxl worker、Rust 后端及 pytools capability dispatcher 注册，并通过合成 Python worker 测试与 Rust 跨进程集成测试。真实模板、生产 capability JSON、隔离 runtime deb staging、glibc 2.28 和麒麟/UOS 均未验收。Audit S2 文件服务已接入 `open-data-audit s2` 独立 Rust CLI，并完成合成 Python 对拍。pytools 工作表区域值提取、地区总分校验、局部文件名映射 4-8、文件排序前缀 4-9 和取消排序前缀 4-10 已接入 Rust CLI 子命令，但尚未接入原产品菜单/dispatcher。pytools PT-7-4 外汇页眉已接入 Rust CLI `fx-header-plan`，默认只生成计划，执行仍须匹配能力证据；4-6/4-7 内容前缀操作已接入 Rust CLI，但未接入原 GUI 菜单/dispatcher。4-9/4-10 CLI 已输出时间戳日志文件和 stderr 成功/失败消息，精确日志时序及并发行为仍未验收。所有候选均非生产默认，也不代表真实业务模板和完整业务调用链已验收。

## 1. 目标与边界

本文件把两个只读 Python 项目观察到的工作簿读写/修改操作，映射到 workspace 已有共享 crate 与待评估候选引擎。函数入口级静态清单另见 [PYTHON_EXCEL_OPERATION_INVENTORY.md](PYTHON_EXCEL_OPERATION_INVENTORY.md)。Python 旧实现继续作为行为对拍基线。业务模块只能产生与具体 Excel 库无关的修改计划；引擎选择、执行、校验和回退统一由共享编辑边界负责。

本次静态盘点覆盖的参考路径：

- Open-Data-Audit：C:\Users\AZI\Desktop\VibeCode\3.open-data-audit\core
- py_statistics_tools PyCode：C:\Users\AZI\Desktop\VibeCode\4.pytools-main\PyCode
- py_statistics_tools NoComPyTools：C:\Users\AZI\Desktop\VibeCode\4.pytools-main\NoComPyTools

以上目录只读。本轮没有在源项目写入代码、配置、测试、缓存或报告。

收尾检查时，两个源项目的 Git 工作树都显示未提交变化（Open-Data-Audit 约 86 条状态行，pytools 约 16 条）；本轮没有在开始时保存其状态快照，不能判断这些改动的来源或时间。`base_audit/models.py` 本身未出现在状态清单中。后续基线必须明确使用并记录当前源文件/工作树版本，不得清理或重置源目录。

## 2. 当前 workspace 与已有验证

### 2.1 共享 Excel crate 现状

| crate | 当前公开边界/用途 | 现状与缺口 |
|---|---|---|
| crates/excel-reader | calamine 读取值、公式文本、sheet 顺序、calamine 可读的合并区域；`read_value_columns_with_iso_dates` 将日期型值按工作簿日期系统读为 ISO 日期时间；roxmltree + zip 补充 OOXML 批注、条件格式、数据验证和 odd/even/first 页眉文本（`read_header_inventory`）；`read_xlsx_sheet_names_ooxml` 仅读取 XLSX `xl/workbook.xml` 中的 sheet directory，不读取单元格值，也不在失败时回退到 Calamine | 元数据专用 reader 使用 namespace-aware XML；页眉 reader 按 package/workbook/sheet relationships 定位 part；拒绝嵌套富文本以免返回不完整结果。主要 API 仍包括 JSON 形状的 WorkbookDump；多格式/错误行为和业务需要的元数据仍需继续盘点 |
| crates/excel-writer | 重新导出 rust_xlsxwriter 的 Workbook、Worksheet、Format 等类型 | 尚未形成 workspace 自有的批量/流式写入 API；pytools 4-9 候选 API 与 audit-core 的 S2 报表写出候选通过该 crate 新建结果簿并由 excel-reader 读回；这些功能级候选不等于通用写入验收 |
| crates/excel-editor | 保留 open_xlsx / save_xlsx 的 umya 薄封装；提供 engine-independent edit plan、精确能力筛选、整批单后端路由、原始输入临时副本、输出验证与 no-clobber 发布；新增隔离 openpyxl XLSX worker adapter | openpyxl 仅已在 synthetic worker/进程集成测试中验证；无生产 capability JSON 或真实模板验收；Excelize 仍为占位；计划操作集合仍是首批子集 |
| crates/ooxml-patch | 检查未验证部件并拒绝；提供 namespace-aware 列定义、普通公式空缓存和 odd/even/first 页眉文本窄补丁 | 页眉补丁现按 `_rels/.rels` 解析 workbook，再按 workbook relationships 解析 worksheet，不假设 `xl/workbook.xml`；仅外汇页眉合成 XLSX 操作有专门对拍。当前不能视为通用编辑器，通用单元格修改仍未实现 |
| crates/probe | 通过 excel-reader 导出/比较工作簿 JSON | 是探针 CLI，不是业务工作流，也不负责修改引擎路由 |

### 2.2 业务侧现状

- audit-core 已有期间/源文件发现、模型与纯值 helper、S3 固定 14 列 CSV 解析，以及独立对拍的 S2 配置读取、期间目录导入、外部 `.xlsx` 导入、质量发现转换、期间比较、规则/表达式计算、外部数值比较和报告摘要统计。CSV 模块不替代共享 Excel reader，也不含 XLS/XLSX dispatch。其 umya 与 writer 探针位于 audit-core/src/bin，属于候选引擎验证代码，不是迁移后的业务工作流。
- `audit_core::s1_template` 和 `open-data-audit template-suggest` 已迁移 `base_audit.discovery.recommend_template` 的工作表集合推荐切片；工作表名只经 `read_xlsx_sheet_names_ooxml` 读取 OOXML workbook metadata。4 个合成案例逐字段匹配，另在只读真实样本上对 8 个模板和 68 个源 XLSX 与 Python 完整输出相同，前后集合哈希一致。该样本完整进程 Release 中位数为 Rust OOXML adapter 1.37 秒、Calamine 3.70 秒、Python 1.76 秒。此为单个 Ubuntu/WSL2 目录树观测，不改变通用 reader 默认路由，也不代表 S1-F01 完整审核迁移；证据见 `audit-core/probe/reports/template_recommendation_parity.json` 和 `template_recommendation_real_smoke.json`。
- 新增的 `audit_core::s2_analysis::analyze_imports` 将已加载配置、已导入的两期记录和可选外部数据接入内存分析：沿用首条逻辑键记录、调用已对拍的周期/规则/外部比较与校验模块、按 Python 服务顺序拼接发现并生成摘要。`s2_analysis_parity.json` 的两例合成对拍通过；它不执行配置检查/加载、工作簿导入、进度回调或 XLSX 导出，因此尚未完成 S2 服务端到端迁移，S2-F01 仍未验证。
- 新增的 `audit_core::s2_exporter::export_s2_report` 是三张 S2 结果表的 Rust 新建工作簿候选：经共享 `excel_writer` 写出、共享 Calamine reader 重开校验，再与只读 Python `ExcelExporter` 作文件级对拍。5 个合成 finding、116 个单元格（含 11 个数值格和 1 个公式格）对拍 0 差异，数值相对容差为 `1e-15`，公式缓存保持为空；样式、列宽区间、冻结/筛选、隐藏行和页面边距均一致。候选显式拒绝未完成对拍的日期/时长值及不能精确保留为 Excel `f64` 的整数；现已由 `s2_service` 合成文件流程调用，但仍未验证真实模板、性能和目标平台，S2-F01 仍未完成。证据见 `audit-core/probe/reports/s2_exporter_parity.json`。
- 新增的 `audit_core::s2_service::run_s2_audit` 将配置检查/加载、本期和上期 XLSX 目录导入、进度、期间与规则计算、可选外部工作簿比较、发现汇总及三表报告写出串成文件级候选；由 `audit-core/src/bin/open-data-audit.rs` 的 `open-data-audit s2` 接入独立 Rust CLI。Python/Rust 对拍覆盖无外部数据（15 条发现）、含外部 `plain_layout.xlsx`（18 条发现）和阻断配置（匹配错误且不创建输出）；服务 API 和 CLI 的摘要/进度均无差异，CLI 输出工作簿与 Python 三表结果无差异。输入夹具和源 Python 基线文件哈希保持不变。证据见 `audit-core/probe/reports/s2_service_parity.json`。这轮只使用合成 fixtures；当前/上期复用了同一报表内容并调整文件名日期，外部 reader 的动态耗时/引擎提示不逐字比较。尚未成为生产默认；真实模板、完整失败矩阵、性能、目标平台和 deb 仍未验证。
- pytools-core 已有日期/路径/规则纯逻辑、配置加载、工作簿单元格抽取、工作表提取区域值 CLI、地区总分校验 CLI、局部文件名映射 CLI 和机构名称规范化模块；这些命令尚未接入原产品菜单/dispatcher。4-6/4-7 已分别接入 Rust CLI pytools rename-by-content 和 pytools unrename-by-content，显式处理文件列表；它们通过共享 Calamine 读取配置和 .xlsx/.xlsm 内容单元格，按不覆盖目标的规则改名，不保存工作簿，也不属于 ExcelModifier。真实宏、.xls/.xlt COM 路径、数字/日期/公式缓存格、真实目录和目标平台仍未验收。4-8 经共享 Calamine 读取 M:N 映射配置，Rust 命令 `pytools rename-local-map` 对任意扩展名的普通文件按映射顺序改名，不打开或保存目标文件。4-8 CLI 与 Python 基线在合成 10 路径场景中统计和最终文件集合相同、文件字节保持。4-9 使用共享 Calamine 读配置、rust_xlsxwriter 新建报告；4-10 对配置中可选内容单元格规格和文件内容使用共享 Calamine reader，对改名前缀生成报告使用 rust_xlsxwriter。上述结果仅限各自合成样本，不是生产默认或完整菜单调用链；4-6/4-7 的真实 `.xlsm`、源 `.xls/.xlt` COM 路径、日期/公式缓存格及目标平台文件系统未验收。
- pytools `fx_header::build_fx_header_plan` 已把外汇页眉单位替换编译成 `ExcelEditPlan`，业务层使用共享 `excel-reader::read_header_inventory` 读页眉，执行经合成样本验证的 OOXML Patch。PyCode 与 NoComPyTools 对应源文件 SHA-256 相同；Python baseline 与 Rust candidate 对 1 个合成 XLSX 的统计、odd/even/first 页眉内容、未修改工作簿内容通过。候选仅支持 XLSX，已接入 Rust CLI `fx-header-plan`（默认 plan_only，执行须显式提供匹配的 capabilities/platform），未接原 Python 菜单、不替换源文件；真实工作簿、真实 XLSM/VBA、目标平台仍未验证。详情见 `pytools-core/probe/reports/fx_header_parity.json`。
- 机构名称规范化计划 API 新增后，2026-10-23 再次执行 `cargo fmt --all -- --check`、`cargo test -p pytools-core --test institution_normalize_plan_parity --locked` 和 `cargo test --workspace --locked -q`；格式检查通过，专项 1/1 通过，全 workspace 命令退出码为 0。该轮只验证计划与合成 Python baseline 一致，不代表 ExcelModifier 有可用的 umya backend，也不验证计划执行后的文件保真。
- pytools `region_sum::run_region_sum_check` 是地区总分校验普通非 COM 路径，已接入独立 Rust 命令 `pytools region-sum-check`：用共享 Calamine 读取 XLSX，并通过 rust_xlsxwriter 从零生成“核对结果”工作簿；为匹配 Python 基线的无缓存公式标记，在保存后调用一项 namespace-aware OOXML 空缓存补丁。它是新建输出，不是既有工作簿编辑或 ExcelModifier 路由；CLI 可显式传入工作簿列表并沿用全局输出/日志配置。CLI 对拍覆盖 5 个合成 XLSX 场景；原 GUI 菜单、非 XLSX、真实模板、性能及目标平台仍未验证，Python 原功能保留。
- pytools `institution_normalize::build_institution_normalize_plan` 已改用共享 `excel-reader` 读取工作簿值，并将合成夹具上的 3 个变更编译成 `SetCellValue` 操作；与 Python 基线统计和变更格对拍见 `pytools-core/probe/reports/institution_normalize_plan_parity.json`。该计划明确保留 `inventory_complete=false`，所以不能被 ExcelModifier 路由执行；测试只验证计划，不验证实际引擎输出。原 `run_normalize_institution` 仍是直接调用 umya 的候选执行路径，未接入统一路由，也不是生产默认。后续须先补齐可审计的 workbook inventory，再实现共享 backend/validator 并用真实模板复验；在完成前不得删除旧候选或宣称该功能已迁移。
- sidecar/lo-convert 已通过 `pytools-core/src/conversion.rs::invoke_sidecar` 显式子进程接入同类格式转换；CLI 为 convert，含输出校验。`recalculate` XLSX 候选命令在临时副本中先清除既有公式缓存、设置 full-calc 标志，再调用单一显式引擎；Ubuntu/WSL2 LibreOffice CLI 合成测试验证 `=SUM(A1:B1)` 缓存由占位 0 更新为 5，且公式文本保留。缓存节点齐全仍不证明真实业务公式正确；Audit 样本对拍、COM 和麒麟/UOS 尚未验收，因此不能作为已验收路由。真实 .xls/.et/.doc 转换及目标平台仍未验收。
- workspace 中没有 Go module 或 Go worker。audit-core/core-go 只有预留说明，不是已实现的 Excelize 引擎。
- 当前 workspace Git 状态为“尚无首次提交，已有文件均未跟踪”。后续变更不得清理、覆盖或假定这些文件已提交。

### 2.3 已保存对拍证据的范围

下列 JSON 报告记录各自对拍轮次。2026-10-23 S2 外部 `.xlsx` 导入轮次曾在 WSL Ubuntu 执行 `cargo fmt --all -- --check` 和 `cargo test --workspace --locked -q`，81 个 Rust 单元/集成测试通过；该数为历史轮次。外汇页眉候选完成后，本轮再次在 WSL Ubuntu 执行 `cargo fmt --all -- --check` 和 `cargo test --workspace --locked`：97 个 Rust 单元/集成测试通过、0 失败，doc-tests 为 0 项；外汇页眉专项集成测试 2/2 通过。S2 报表写出候选完成后 workspace 测试命令退出码为 0；候选集成测试 1/1 通过，Python 文件级对拍 0 差异。S2 文件级 service 候选后续在 WSL Ubuntu 执行 `cargo fmt --all -- --check` 和 `cargo test --workspace --locked`：105 个测试通过、0 失败，doc-tests 为 0 项；Python/Rust 三场景服务对拍通过，详情见 `s2_service_parity.json`。外汇页眉 Python baseline 由只读旧实现生成，详见对应 JSON；原项目源文件保持只读。既有 pytools Python `test_rename_local_map_48.py`、`test_convert_rename_edges.py` 共 19 项以及更早的 4-6/4-7 源测试均为历史轮次证据，不是本轮重跑。4-6、4-7、4-8、4-9 和 4-10 各自有独立合成文件集对拍。2026-10-23 接入 4-6/4-7 CLI 后再次在 WSL Ubuntu 执行格式检查、workspace 测试和 cargo check：121 个单元/集成测试通过、0 失败；4-6 集成测试 2/2、4-7 集成测试 4/4 通过。上述证据不等于真实模板、Windows 原生 Rust 或目标 Linux 打包验收：

| 报告 | 已验证结果 | 不能据此推断的范围 |
|---|---|---|
| audit-core/probe/reports/models_parity.json | 10 个 audit 模型类型的字段/序列化、7 组 dataclass 默认值、预检/审核结果计数和摘要文本对拍通过 | 预检、规则执行、审核结果生成和报告写入流程 |
| audit-core/probe/reports/s2_external_import_parity.json | 12 个合成输入场景中，Calamine 外部 `.xlsx` 解析与 OpenPyXL 源基线逐字段/错误文本一致；日期、空值、警告、机构映射和预期拒绝分支有独立夹具 | 真实外部工作簿、多格式成功读取、S2 服务编排、外部比较集成、报告写入、性能/内存及目标 Linux/deb |
| audit-core/probe/reports/s2_analysis_parity.json | 2 个合成输入案例：周期/外部/表达式规则结果、质量发现、拼接顺序与摘要逐字段匹配只读 Python 服务组件顺序 | 该分析切片自身不覆盖配置检查/加载、Excel 导入、进度或输出；更大服务范围另见 s2_service_parity.json；真实模板、性能/内存、目标 Linux/deb 及生产路由仍未验证 |
| audit-core/probe/reports/s2_exporter_parity.json | 5 个合成 finding 写出三表；116 个比较单元格、11 个数值格、1 个公式格；与只读 Python `ExcelExporter` 文件级对拍 0 差异（相对容差 `1e-15`）；公式缓存为空；值/类型、样式、列宽区间、冻结/筛选、隐藏行和页面边距均比较 | 真实模板、非合成数据分布、性能/峰值内存、目标 Linux/deb；候选仅在合成 S2 service 流程中被调用，未接产品 CLI/生产路由 |
| audit-core/probe/reports/s2_service_parity.json | 无外部/有外部/阻断配置 3 个文件流程；15/18 条发现、22/26 个进度事件；service API 和 `open-data-audit s2` CLI 的摘要、进度及 CLI 三表工作簿与 Python 对拍均 0 差异；输入与源 Python 哈希未变；相对容差 `1e-15` | 合成输入（本期/上期复用内容）；外部 reader 完成消息动态字段不逐字比较；真实模板、完整失败矩阵、性能/峰值内存、Windows 原生运行、目标平台/deb；CLI 尚非生产默认 |
| audit-core/probe/reports/fastpath_read_compare.json | 合成 xlsx：sheet 1/1、值 65/65、公式 1/1、合并 1/1、批注 1/1 | 真实模板与 xls/xlsm/xlsb/csv |
| audit-core/probe/reports/fastpath_writer_compare.json | rust_xlsxwriter 新建合成 xlsx：9/9 | 编辑既有工作簿、复杂样式保留、真实模板 |
| audit-core/probe/reports/umya_raw_compare.json | 合成修改样本：17/18；列定义集合有额外显式宽度 | 不能作为无损默认编辑器 |
| audit-core/probe/reports/umya_patched_compare.json 与 ooxml_patch.json | 列定义补丁后：18/18；补丁范围：3/3，报告记录仅 sheet XML 的 cols 区域改变 | 仅覆盖该合成文件及该补丁，不代表完整 ExcelModifier |
| pytools-core/probe/reports/config_task_loaders_parity.json | 配置读取向量 11/11 | 配置初始化/修复、真实配置及消费配置的业务 |
| pytools-core/probe/reports/workbook_extract_parity.json | 一个合成工作簿中 helper 提取 6/6 | 批处理、写回/追加、真实模板及多格式 |
| pytools-core/probe/reports/sheet_extract_parity.json | 工作表提取纯规则向量 11/11 | 完整抽取文件工作流、整表样式复制、写入/追加 |
| pytools-core/probe/reports/sheet_extract_area_parity.json 与 sheet_extract_area_python_baseline.json | area-only Rust CLI 候选：两个有效合成 XLSX、一个损坏输入；4 张输出表、1 个新工作簿；文件/表名/区域值/尺寸/统计与 PyCode 基线相同；CLI 配置加载、输入字节保持、日期拒绝、append/full-sheet 拒绝及目标冲突不覆盖测试通过 | 显式 XLSX 文件列表；公式夹具、真实模板、xls/xlsm/xlsb/csv、目录扫描、日志等价、完整菜单、整表保真与追加已有簿 |
| pytools-core/probe/reports/institution_normalize_parity.json 与 institution_normalize_preservation.json | 合成 xlsx 规则与值/公式/合并/批注、所列样式尺寸/打印属性通过 | 真实模板、xlsm、菜单集成及完整对象空间 |
| pytools-core/probe/reports/file_sort_prefix_operation_parity.json | 4-9：合成 xlsx 配置、8 个输入路径、5 次改名/1 未命中/2 跳过，API 与 Rust CLI 的统计、报告工作簿和文件名集合与 Python 基线一致；文件内容保持、成功日志行、报告写入失败回滚和 no-clobber 碰撞测试通过 | 原 GUI 菜单、日志时序/并发等完整 logger 等价、真实目录、Windows 原生 Rust、硬链接不支持的文件系统、4-10、目标 Linux/glibc 2.28 |
| pytools-core/probe/reports/file_sort_prefix_remove_parity.json | 4-10：10 个输入/路径案例，5 次改名/5 次跳过/0 失败；内容前缀读取覆盖合成 xlsx，API/CLI 报告行和最终文件集合与 Python 一致；空输入 API/CLI 均输出 Python 兼容空报告；CLI 日志成功项通过合成断言；Rust 还检查改名文件字节未变化、报告写入失败回滚和无规则拒绝 | `.xls/.xlt` 内容读取、日期/公式缓存单元格、多单元格规格、原 GUI 菜单、日志时序/并发等完整 logger 等价、真实目录、Windows 原生 Rust、目标 Linux/glibc 2.28 |
| pytools-core/probe/reports/local_filename_map_48_parity.json 与 local_filename_map_48_python_baseline.json | 4-8：合成 10 个输入路径，5 次改名/5 次跳过；6 条规则覆盖顺序替换、中文片段、空目标、冲突后缀和失败保留；Rust CLI 检查最终文件集合和文件字节；源 Python 两个聚焦测试文件 19 项通过；本轮 CLI 集成测试 3/3 通过 | 仅合成文件/配置；原 GUI 菜单、真实配置和目录、非 UTF-8 名称、Windows 原生 Rust、不同文件系统、目标 Linux/glibc 2.28 |
| pytools-core/probe/reports/content_rename_46_parity.json 与 content_rename_46_python_baseline.json | 4-6：合成 12 个输入路径，5 次改名/7 次跳过；配置统计与最终文件名集合和 Python 一致，Rust 测试检查原工作簿字节未变及冲突目标保留；Python 聚焦源测试 13/13 为前轮历史证据（本轮未重跑）；Rust API+CLI 集成测试 2/2、pytools-core 49 项、workspace 121 项通过 | 仅合成文本格与候选 `.xlsx/.xlsm`；真实 `.xlsm`、`.xls/.xlt` COM、数字/日期/公式缓存格、真实目录、不同文件系统、原 GUI 菜单/dispatcher、Windows 原生 Rust、目标 Linux/glibc 2.28 |
| pytools-core/probe/reports/content_unrename_47_parity.json 与 content_unrename_47_python_baseline.json | 4-7：合成主场景 13 个输入（改名 8、跳过 5），旧配置缺排序工作表场景 2 个输入（改名 1、跳过 1）；Rust 断言统计和最终文件集合与 Python 一致，并检查改名文件字节不变；Python `test_rename_by_content_46.py` 13/13（前轮历史，本轮未重跑）、Rust 4-7 API+CLI 集成测试 4/4 通过，pytools-core 49 项、workspace 121 项通过 | 仅合成 `.xlsx` 与无 VBA XLSX 包改扩展名的 `.xlsm`；真实 `.xlsm`、`.xls/.xlt`、日期/公式缓存格、真实目录、不同文件系统、原 GUI 菜单/dispatcher、目标 Linux/deb |
| pytools-core/probe/reports/region_sum_parity.json 与 region_sum_python_baseline.json | 地区总分校验普通版：5 个合成 XLSX 输入，Python/Rust 统计均为文件 4、工作表 3、错误 1、失败文件 1；结果值、公式文本、列宽和所列样式通过；输入字节保持、输出文件/日志冲突不覆盖及输出目录失败保护测试通过。2026-10-23 接入 `pytools region-sum-check` 后，Rust CLI 统计/输出值/公式及配置目录路由通过基线测试；workspace 全量格式检查、测试与 cargo check 通过 | 原 GUI 菜单、真实业务模板、非 XLSX 格式、完整日志等价、性能/峰值内存、Windows 原生 Rust、麒麟/UOS glibc 2.28 与 deb |
| pytools-core/probe/reports/fx_header_parity.json 与 fx_header_python_baseline.json | 1 个合成 XLSX：Rust 候选计划发现 2 个名称含“外汇”的工作表、1 个待修改工作表，与 Python 对应计数一致；单文件计划成功，页眉修改、非目标页眉、值/公式/合并/批注/条件格式/数据验证及 Calamine 重开检查通过；OOXML Patch 限定验证仅请求的页眉文本节点变化 | Rust 批量文件统计、真实工作簿、XLSM/VBA、菜单/CLI、原位替换、Windows 原生 Rust、麒麟/UOS glibc 2.28 与 deb |

详细历史环境、命令与限制见上述 JSON 和对应 GPT修改记录。存档结果证明相应旧轮测试曾执行，不是本轮当前代码状态的重新验证。

## 3. 业务修改抽象

### 3.1 ExcelEditPlan

业务层将规则结果编译为一个文件级、可序列化或可检查的 ExcelEditPlan。计划只描述“修改什么”和“必须保留什么”，不得包含 umya、rust_xlsxwriter、calamine 或 Excelize 的类型。

建议计划至少包含：

- schema_version、源文件身份/格式、目标 output_format；
- write_pattern（append_only / random_access / edit_existing），按实际算法区分追加、随机写入与原位修改；
- value_view（stored 默认；displayed 必须有已验收的格式解释能力）；
- 按工作表和单元格/区域定位的有序修改操作；
- 修改前置条件（例如 sheet/单元格必须存在、原值匹配预期）；
- preserve 清单、drop_allowed（默认空）及需要修改的对象清单；
- allow_conversion（默认 false）；输入输出格式不同必须明确为转换任务，不能冒充保真编辑；
- external_apps（主路径默认 false）及 reference_app（默认 none），避免隐式依赖办公应用或其显示语义；
- 操作来源/业务调用标识，供对拍报告定位；不得把业务名称放入共享 crate 的通用能力名。

先从已观察到的 Python 操作定义最小操作代数：写入/清除单元格值、写入公式文本、样式设定/复制、批注新增/替换/删除、合并/取消合并、行列尺寸与隐藏属性、页面设置、工作表创建/删除/重命名/排序、命名区域更新、区域复制。只有源代码调用链证明需要时才扩展操作类型；不把完整 Excel 对象模型塞进首版计划。

同一逻辑结果中的公式文本与缓存值必须是不同语义。ExcelEditPlan 不表达“已重算”，除非明确选择并验证了公式重算 sidecar。

当前 `crates/excel-editor/src/plan.rs` 已实现首批类型：普通值、公式文本、批注、行列尺寸、删行、sheet 改名及本规范范围内的页面设置。尚缺 schema_version、写入模式/值视图、修改前置条件、允许丢弃项和转换标记；样式、合并、名称、跨表复制、隐藏/大纲等操作仍待从实际调用链中确认并扩展，不能据此声称完整覆盖源项目操作。

### 3.2 ExcelModifier

共享 excel-editor 演进为统一 facade，业务仅调用该 facade。引擎能力以“文件格式 + 工作簿对象特征 + 整批操作组合 + 保留契约 + 目标平台 + 引擎版本”为单位登记，状态使用：

- passed：该适配器能力组合有对应 fixture、对拍结果和报告，可以进入路由候选；
- unsupported：引擎明确不支持该格式、对象或操作；
- unverified：尚无足够证据，不能进入自动路由。

能力通过状态与保留项报告状态分开：保留项逐项标注“保留/丢失/损坏/不适用/未验证”，不得用“适配器 passed”代替对象级验证。

当前代码的能力声明由运行时调用方提供，路由检查声明字段与本次计划是否匹配，并要求报告/版本/平台字段非空且平台一致；它不读取报告文件或验证报告内容。尚无从 probe/reports 自动加载并核验能力登记的生产注册表。适配器接入前仍须人工/CI 确认对拍证据，禁止把任意声明直接注册为 passed。当前没有已接入的生产 backend，因此这些类型仅构成架构边界。

引擎选择按计划整批进行：先检查输入包与保留契约，再根据操作类型和对象特征筛出 status=passed、覆盖整批操作的候选；在合格候选中结合真实保真结果、实测性能、维护成本和失败模式选择。不得按单元格把同一文件在多个引擎间来回打开/保存，也不得只按语言或理论速度选引擎。没有单一引擎覆盖该计划时，拒绝执行或把计划拆成互不依赖的独立输出任务；不得悄悄降级为不完整结果。

### 3.3 执行、回退与提交边界

1. 保留不可变源文件；对其做只读工作簿/部件清点并确认计划的前置条件。
2. 将完整计划一次性提交给所选引擎，写入同文件系统的临时输出；禁止覆盖源文件。
3. 重新打开输出，验证文件可读、计划修改确已生效、指定保留项及未修改区域未损坏。
4. 只有后端在写出前明确拒绝当前文件特性且另一后端已覆盖整批计划时，才从原始输入重新执行；写入失败、协议违约、输出验证失败或未知异常停止，不自动切换后端。
5. 只有验证通过后才提交目标文件；失败保留原始文件，并返回结构化错误及路由/尝试报告。当前骨架以 no-clobber 方式发布到新路径，不替换已有目标。

批处理按文件边界并行；同一工作簿的操作顺序由一个计划保持确定性。

当前执行骨架只接受输入/输出扩展名与登记格式一致的新目标路径；为每个候选引擎从原始输入创建独立临时副本。只有 `UnsupportedBeforeWrite` 且候选输出未写入时才继续下一个已合格引擎；写入失败、协议违约或输出校验失败均停止。通过 validator 后以同目录 hard link 做原子 no-clobber 发布；目标已存在时拒绝覆盖。覆盖/备份替换策略、实际 Excel validator 和 backend 仍未实现。

## 4. 三类候选引擎及能力边界

三个候选引擎不要求全部实现或启用。下表是测试优先级/候选能力，不是已支持声明。

| 候选 | 适合先验证的操作 | 主要风险与拒绝条件 |
|---|---|---|
| Rust OOXML Patch | 现有 xlsx 中位置明确的单元格数值/文本/普通公式文本；少量明确 XML 节点，如页面属性或定义名称的局部更新 | 必须 namespace-aware；关系、Content_Types、sheet 身份/顺序、sharedStrings、样式索引和公式引用都需按操作维护。首批只接受已验证的部件组合；新增字符串若需扩充 sharedStrings、跨 part 联动或依赖语义不清时拒绝 |
| Rust umya-spreadsheet | 以工作簿对象模型处理的复杂 XLSX 修改：样式、批注、合并、行列属性、页面设置、工作表结构 | 打开/保存会重写工作簿，须用真实模板核验未修改部件。当前只有一份合成样本和列定义补丁证据；不能默认覆盖 xlsm、xlsb 或未清点对象 |
| Go Excelize worker | 复杂 XLSX 修改的候选/保底；一条 IPC 消息接收一个完整文件计划并返回结果和校验摘要 | 当前尚不存在 worker。必须先比较同一真实模板/操作批次，证明保真、性能或维护收益，再引入；CGO_ENABLED=0 与 worker 二进制一起交付，不要求用户安装 Go |

OOXML Patch 是狭窄的补丁工具，不是完整读写器。新建工作簿仍由 rust_xlsxwriter 负责；不为统一栈重写完整 OOXML 写入器。复杂结构不得通过忽略未知 part 或复制 ZIP 条目后不校验来冒险保存。

### 4.1 按操作类型划分的首轮候选

下表是进入真实模板对拍的候选分组，不代表已支持。表中凡未有对应 fixture、报告和目标文件可打开性验证的能力都保持 `unverified`。

| 操作 | 首轮候选 | 暂不允许的推断 |
|---|---|---|
| 对现有单元格写数值/文本，且不插行、不改样式 | OOXML Patch | 新文本是否使用 sharedStrings、目标是否为共享/数组公式单元格、单元格类型及工作表尺寸更新均需验证；无法证明时拒绝 |
| 对现有单元格写普通公式文本 | OOXML Patch 或对象模型引擎整批对拍 | 写公式不等于重算；共享公式、数组公式、外部引用、calcChain/calcPr 语义未验证时拒绝 |
| 只改 pageSetup / printOptions / 明确的 definedName 节点 | 小范围 OOXML Patch；umya/Excelize 作为同模板比较组 | printArea 还涉及 workbook 名称定义；sheet 改名后的名称引用不是单节点任务 |
| 替换既有 worksheet 的 odd/even/first 页眉文本 | OOXML Patch + ExcelEditPlan；metadata 通过共享 excel-reader 获取 | 目前仅合成 XLSX 的 `万元`→`万美元` 操作通过；不推断页脚、页眉富文本、XLSM/VBA 或其他文本变更能力 |
| 样式、批注、合并、行列尺寸/隐藏/大纲 | umya、Excelize 与 openpyxl 对同一完整计划、同一模板比较；openpyxl 只覆盖已实现的 XLSX 操作 | 不预设其中任一必需或胜出；OOXML Patch 仅对已经验证的窄操作组合开放；openpyxl 不保证保留所有 Excel 对象 |
| 插入/删除行、创建/删除/重命名/排序 sheet、跨 sheet/工作簿复制 | umya 与 Excelize 候选；openpyxl 目前拒绝插删行，已实现的 sheet 与区域操作仍须逐项对拍 | 不按单元格拆分；必须覆盖公式引用、relationship、definedNames、条件格式/验证和未改对象；不支持的操作必须写前拒绝 |
| 新建报表、配置、历史簿 | rust_xlsxwriter | 不为新簿使用 ExcelModifier；若旧逻辑要求在用户配置簿上修复，则按既有簿计划评估 |
| `.xls/.et→.xlsx`、`.doc→.docx` 同类转换 | 显式触发 sidecar/lo-convert | 不属于 ExcelModifier；不自动启动，不宣称保真成功，输出需校验 sheet/数据或文档可打开性 |
| 公式重算 | 显式触发 sidecar/lo-convert | 不属于业务规则计算；主路径不读旧缓存冒充本轮重算 |

所有候选都应先完成操作入口调用链和真实模板/合成夹具契约，再建立路由。样式、批注、合并、尺寸、行列结构和 sheet 结构属于需要比较 umya、Excelize、openpyxl 的操作族，不等于现在已确认“必须用”其中某个库。Go worker 是否加入取决于实测收益。

### 4.2 操作级能力矩阵（2026-10-22 初始建立；2026-10-23 增补；2026-10-25 增补 MoveSheet、冻结窗格/AutoFilter、校验器行操作映射、合并不锚载荷豁免与逐表快照顺序感知）

本矩阵只登记仓库已有代码和报告，作为后续补齐依据；不是完整操作覆盖声明。2026-10-23 增补外汇页眉与 openpyxl synthetic worker/进程测试证据。passed 仅限同一行明确写出的样本/操作；openpyxl 合成测试不产生生产 capability，也不代表真实模板或目标平台已验收。

| 操作子类与保留契约 | Rust OOXML Patch | Rust umya-spreadsheet 3.1.1 | Go Excelize worker | openpyxl worker | 当前证据与缺口 |
|---|---|---|---|---|---|
| 修改已有普通单元格值/文本，不插行、不改样式 | unsupported（尚无对应写入适配器） | unverified（合成混合探针修改成功，未隔离验证独立契约） | unsupported（无 worker/IPC） | implemented; synthetic Rust 跨进程集成通过 | 无真实模板/目标平台 capability；umya 混合探针也不能推导所有类型、共享字符串和模板保真。 |
| 写入普通公式文本，不要求重算 | unsupported（尚无对应公式写入适配器） | unverified（合成混合探针保留公式文本） | unsupported（无 worker/IPC） | implemented; synthetic Rust 跨进程集成通过；非空公式缓存写前拒绝 | 公式写入不等于重算；公式类型、缓存、计算链和真实模板仍未验证。 |
| 新增/编辑/删除批注及其位置、作者/外观 | unsupported（尚无 comments/VML 写入适配器） | unverified（探针只验证既有批注文本保留） | unsupported（无 worker/IPC） | implemented; synthetic worker 测试覆盖新增/替换与未修改批注保留 | 只断言文本/作者，不验证批注外观；真实 comments/VML/relationship 组合未验收。 |
| 样式、合并、行高列宽、隐藏、页面设置的完整混合批次 | unverified（仅窄列定义补丁） | unverified（原始探针 17/18；补列定义后合成样本 18/18；SetFreezePanes/SetAutoFilter 已实现并通过合成测试：冻结窗格按 openpyxl 语义设置/替换，AutoFilter 与 _FilterDatabase 联动更新/清除；清除已有冻结窗格写前拒绝，需其他引擎路由） | unsupported（无 worker/IPC） | implemented; synthetic worker/进程测试覆盖填充、合并导入、尺寸、页面设置；Rust 独立读回检查已覆盖样式语义、显式行列尺寸、填充色、页面设置及 OOXML Print_Area | openpyxl 会重写 package；ImportRange 的行列尺寸/隐藏/outline 复制标志仍写前拒绝；保存后以 namespace-aware XML 解析器检查 cellXfs 计数与元素数，并检查单元格/行 s 和列 style 引用均在界内；真实模板保留项、打印应用行为和目标平台未验证。冻结窗格/AutoFilter 合成证据见 crates/excel-editor/tests/freeze_panes_auto_filter.rs。 |
| 设置目标列宽并保留样本中原有列定义 | passed（synthetic_probe.xlsx 定点补丁报告 3/3） | unsupported（同样本原始输出未保留列定义） | unsupported（无 worker/IPC） | unverified（非本 worker 的特定列定义契约） | OOXML 补丁只改 worksheet 的 cols 区域；不得扩展为通用编辑器能力。 |
| 替换现有 worksheet 的 odd/even/first 页眉普通文本，并保留其余部件 | passed（仅 `fx_header` 合成 XLSX） | unverified（未评估） | unsupported（无 worker/IPC） | unsupported（Rust adapter 不开放 worker 的此操作） | `ReplaceHeaderText` 继续路由至窄 OOXML Patch；不能因 worker 有内部实现就绕过既有已验收窄路径。 |
| 插删行列、改名/重排/删除 sheet、跨表/工作簿复制 | unsupported（通用结构适配器未实现） | unverified（MoveSheet 已实现并通过合成往返测试：顺序重排、sheet-scoped definedName localSheetId 重映射、activeTab 跟随、页面设置保留；CreateSheetAt/DeleteSheet 已同步修复 localSheetId/activeTab 滞留；真实模板与依赖对拍未做） | unsupported（无 worker/IPC） | partial: synthetic create/rename/delete sheet 与 ImportRange 进程测试通过；行插删 unsupported；引擎无关校验器 validate_sheet_names 已支持 MoveSheet 顺序期望，未修改值/公式/批注/样式读回比较已按 DeleteRows/InsertRows 行坐标映射（删除区内源格跳过、下移行按映射坐标比对；下移公式的引用重写不做逐字比较） | 复制/导入和结构操作需核对公式、definedNames、relationships、条件格式/验证及未知部件；无真实 capability。MoveSheet 合成证据见 crates/excel-editor/tests/move_sheet.rs 与 tests/sheet_index_metadata.rs；校验器行映射证据见 tests/validator_row_mapping.rs。合并不锚载荷基线等价豁免已实现：drop_allowed 新增 MergedNonAnchorPayloads 按表显式声明（无需 DeleteSheet），openpyxl 目标侧预检收窄到与操作目标矩形相交的载荷，读回对合并区不变的非锚空值豁免值/公式比较；未声明的计划保持全局严格守卫；证据见 tests/merged_payload_exemption.rs 与 excel-editor lib 单测，真实 7/7 样本基线等价证据由 Audit 提供。引擎无关校验器的逐表 OOXML 快照比较（pageSettings/页眉页脚/条件格式/数据验证）已改为按表名键控、顺序不敏感，MoveSheet 重排不再误拒；openpyxl worker 保存后新增样式表完整性断言（cellXfs count 与 xf 元素一致、全部 s= 引用在界内），不一致显式失败不发布损坏候选。证据见 crates/excel-editor/tests/python/test_openpyxl_worker.py 与 lib 单测；真实模板上的 worker 样式表缺陷（audit-dev 请求 #5）根因待真实样本材料确认。 |

样式完整性断言使用 namespace-aware XML 解析：核对 cellXfs 计数及元素数，并检查单元格/行 s 与列 style 索引均在界内。AutoFilter 的 _xlnm._FilterDatabase 范围更新/清除应同时保持 hidden 属性；合成回归见 crates/excel-editor/tests/freeze_panes_auto_filter.rs。

能力矩阵须继续拆到具体源操作与 ExcelEditPlan 操作序列；不能用上表代替 PYTHON_EXCEL_OPERATION_INVENTORY.md 的调用链核验。操作、样本或契约未确认时保持 unverified。引擎注册和合成测试不等于生产验收；不得据此自动改路由或创建 Passed capability。

## 5. Python Excel 操作入口与修改操作族盘点

以下清单基于 PyCode、NoComPyTools 和 Open-Data-Audit 业务源码的 AST/文本静态检索，涵盖当前找到的主要 openpyxl 写入点、工作簿对象修改、ZIP/XML 输出以及旧 COM/LibreOffice 引擎入口。它比初始函数清单增加了按操作类型的首轮候选映射，但仍不替代逐菜单/逐函数的完整调用链审计；具体哪些分支进入正式迁移范围，仍以两个项目的功能清单及后续源码追踪为准。

| 操作类别 | 源码中观察到的行为与代表位置 | 候选路径 | 必须增加的保真验证 |
|---|---|---|---|
| 单元格值写入/清除、批量行输出 | audit：base_audit/native/result_writer.py、summary.py、history_io.py、s1audit_configured/outputs.py、S2/S3 exporter；pytools：io_excel.py、dedup_311.py、机构标准化、外资行删除、拆分、地区校验、贷款调整、汇总 | 新建目标：rust_xlsxwriter；原位计划：OOXML Patch、umya 或 Excelize 按整批能力选择 | 值/类型/空值、行顺序、公式文本、单元格地址、未改区域、目标文件结构 |
| 普通公式写入/公式区域回填/跨簿引用重写 | audit：base_audit/formula_region_writer.py、external_sheet_writer.py、native/template_merge.py、S3 fast_ooxml.py；pytools 多数读公式或复制公式相关数据 | 定点公式区域可先评估 OOXML Patch；复杂模板合并/多引用关系交给整体编辑器候选；重算另属 sidecar | 公式文本逐格一致、名称与跨表引用、calcChain/fullCalcOnLoad 语义、缓存不冒充重算 |
| 样式写入/复制 | 字体、填充、边框、对齐、数字格式及综合样式复制；audit：native/template_merge.py、s1audit_configured/outputs.py、S2/S3 输出；pytools：feature_1_5_region_sum_check.py、feature_1_6_extract_sheets.py、print_310.py、summary_321_322.py | rust_xlsxwriter 用于新建；既有模板对 OOXML Patch、umya 与 Excelize 按同一完整操作批次对拍，不预设优先级 | styles.xml 索引、主题/自定义数字格式、目标与非目标单元格样式、Excel/WPS 可打开性 |
| 批注读取/新增/替换/删除/复制 | audit：conditional_context.py、native/template_merge.py、native/openpyxl_workbook.py 等；pytools：去重标识、打印/汇总标记 | 读取用 excel-reader 的 OOXML 补充；写入先评估 umya/Excelize，OOXML Patch 仅在 comments、关系和 VML/Drawing 部件组合验证后使用 | 文本、作者（如属于行为契约）、位置、comments relationship、VML/drawing、未改批注 |
| 合并区域读取/创建/复制 | audit：native/template_merge.py 与结果模板；pytools：feature_1_6_extract_sheets.py、print_310.py | 读取 calamine/OOXML；写入/复制由整批编辑引擎候选处理 | 区域边界、锚点值/样式、重叠冲突、复制前后其他合并区 |
| 行高/列宽/隐藏/大纲属性 | pytools：整表抽取、打印复制、地区校验；audit：模板合并及输出格式化 | umya/Excelize；已存在的窄列定义修补可用现有 OOXML patch 作为单操作候选 | 显式与默认尺寸区别、隐藏/大纲、范围覆盖、未修改尺寸。已有 umya 样例说明列定义可能被重写 |
| 页面设置、打印区域、居中、页脚 | pytools：print_310.py；audit：s1audit_configured/outputs.py 与格式化输出 | umya/Excelize 或严格限定的 pageSetup/printOptions/definedNames OOXML Patch | orientation、fitToPage、fitToWidth/Height、scale、paperSize、printArea、居中、页脚；只验证页面设置，不渲染/导出 PDF |
| 工作表新建/删除/改名/排序/清空 | audit：config_guide.py、feature_log.py、template_merge.py、S2/S3 输出；pytools：配置初始化、拆分、删除分支、打印、汇总、提取 | 新建簿 rust_xlsxwriter；既有簿的结构操作用完整编辑器候选；xls 改名按已确认流程先显式转换 xlsx | sheet identity/order/state、关系 ID、公式/名称引用、打印名称、内容类型 |
| 工作表/区域复制及模板合并 | audit：native/template_merge.py；pytools：feature_1_6_extract_sheets.py、print_310.py、summary_321_322.py | 评估 umya/Excelize 的工作簿对象能力；若操作可被降为全新工作簿输出则用 writer，但先确认业务保留契约 | 值/公式、样式、批注、合并、尺寸、页眉页脚、冻结/筛选/验证/条件格式及名称；单测不够 |
| 命名区域及引用更新 | audit：native/template_merge.py、formula_region_writer.py、S3 formula/reference 检查 | 少量确定的定义名称节点可评估 OOXML Patch；成组工作簿修改交给整体编辑器 | workbook/local scope、公式目标、sheet rename 后引用、内部打印名称 |
| 新建工作簿/多表报表 | audit：s1audit_configured/outputs.py、feature_log.py、S2/S3 exporter；pytools：io_excel.py、打印/抽取/汇总/地区报告 | rust_xlsxwriter；优先评估其逐行写入/低内存模式 | 与旧输出逐表、逐值、逐格式比较；错误退出时不留下半成品 |
| 自定义 OOXML/ZIP 路径 | audit：s1audit_configured/fast_ooxml.py（生成工作簿）、base_audit/formula_region_writer.py、external_sheet_writer.py、S3 fast_ooxml.py | 逐个源函数决定：全新生成转 rust_xlsxwriter；局部已存在 part 的修改才考虑 namespace-aware OOXML Patch；复杂复制先评估对象编辑器 | package parts 清单、relationship/Content_Types、sharedStrings、defined names、未改 bytes/语义、目标 Excel/WPS 能打开 |

### 5.1 静态扫描发现的补充边界

- pytools 另有配置簿初始化/修复（`config_init.py`）、去重检查与执行（`dedup_311.py`）、逐表数据生成与追加（`io_excel.py`）、行插入/删除（`feature_1_1_split_village_bank.py`、`feature_1_3_remove_foreign_bank.py`）、整表复制（普通版 `feature_1_6_extract_sheets.py`）、打印文件块复制及页面设置（`print_310.py`）、地区检查/汇总新簿（`feature_1_5_region_sum_check.py`、`summary_321_322.py`）、贷款调整（`feature_1_8/1_9`）和 Sheet 改名/文件转换（`convert_33_34_35_37.py`）。`compare_393.py` 等只读或纯计算入口不因含有列表 `append` 被归为工作簿修改；跨 helper 的真实写入点仍待调用链核验。
- audit 除模板/结果操作外，还有用户配置簿与系统配置簿更新：`base_audit/name_config.py`、`node_flow_config.py`、S2/S3 `config.py`、`web_app.py` 等。它们常将单元格值、工作表结构、样式和批注组合保存，按完整计划评估，不将“配置文件”误判成纯文本配置。
- `base_audit/excel_com.py` 的 `ExcelSession` 包含模板读取/检查、规则执行、条件格式与 issue 读取、跨簿 sheet 复制，以及导航链接、历史、预检和汇总输出等业务工作簿操作。它是旧业务引擎参考，不是允许整体留在 sidecar 的 COM 服务。仅 `calculate_pending_workbooks` / `calculate_workbook` 一类明确的公式重算职责可拆到显式 sidecar；其他数据读取、规则和报告输出必须迁至 Rust 路径。
- 已发现的 `ExcelSession` 调用方包括 `base_audit/engines/excel_wps.py`、`merge_org.py`、`service.py`、`region_summary.py`、`workflow/runner.py`、`workflow/adapters/audit_handlers.py`、`systems/s3_central_statistics/office_eval.py`；`formula_region_writer.py` 也调用 `session.apply_rules`。这表明该旧 COM 引擎跨多个业务流程，不可整体移入 sidecar。每个调用方仍需逐入口映射输入、输出和错误路径。
- pytools 中 `*_com.py` 模块是旧实现分支/对拍参考；工作表提取 COM 批注版仍明确不迁移。除同类格式转换、显式公式重算外，不得把这些业务 COM 工作流包装为旁路以规避 Rust 迁移。
- `PyCode/main.py` 会导入并在打印、汇总、地区总分校验、工作表提取及贷款调整等路径显式路由普通版和 COM 版；`NoComPyTools/main.py` 是另一条 dispatcher。菜单分支和版本差异尚未逐项对拍；打印 PDF 的 COM 入口按当前范围排除。
- audit 的 `base_audit/engines/libreoffice.py` 主要提供旧计算引擎发现和批量公式重算。Rust sidecar 只可继承对拍已确认的重算需求；其 Python 辅助的 calcPr 修复逻辑不能未经重写直接搬用，OOXML 修改仍遵循 namespace-aware 约束。

额外只读元数据不是“修改操作”，但决定编辑器能否安全保存：公式、批注、合并、条件格式、数据验证、超链接、defined names、外部关系、冻结/筛选、页眉页脚、VBA/扩展部件都应进入文件清点和保留契约。模板历史检查未发现图表、透视表、切片器、图片、保护或 VBA 宏，但这不能替代对新输入文件的检查。

## 6. 分阶段迁移与验收清单

### 当前阶段：先完成盘点与统一边界设计

- [x] 检查 workspace、业务 crate、共享 crate、历史对拍报告。
- [x] 盘点两份 pytools 变体与 audit 项目里可见的主要工作簿操作类别。
- [x] 建立 Python Excel 操作函数入口静态清单：PYTHON_EXCEL_OPERATION_INVENTORY.md。
- [x] 记录 ExcelEditPlan / ExcelModifier、整批路由、原始输入回退和临时文件校验原则。
- [x] 对照 openpyxl、ZIP/XML 与旧 Office 引擎入口，补充操作族候选路由；引擎能力本身仍未验证。
- [x] 建立 §4.2 操作级能力矩阵初始版本；只登记现有证据，完整操作与真实样本验收仍待补齐。
- [ ] 逐菜单/逐源函数建立完整读、写、修改调用链映射，确认所有输出文件/配置/失败行为。
- [ ] 逐项比对 PyCode 与 NoComPyTools 的共有/独有工作簿行为，沿用 pytools-core/MIGRATION_INVENTORY.md 中的待确认口径。
- [ ] 追踪 `ExcelSession`、pytools `*_com.py` 与菜单/dispatcher 的所有调用点；旧 COM 业务工作流继续迁移到 Rust。未来经独立验收的 Windows COM 编辑仅通过 ExcelModifier，公式计算仅通过独立 ExcelCalculator；同类格式转换仍走显式 sidecar。
- [ ] 取得/确认真实脱敏模板；无样本时只推进合成测试并持续标记真实模板未验证。

### 后续阶段

1. **读能力矩阵**：以格式签名和真实文件为准，分别验证 xls、xlsx、xlsm、xlsb、csv；列明值/公式/合并及 OOXML 元数据覆盖。不可从扩展名或 ZIP 容器推断。
2. **新建写入**：用 rust_xlsxwriter 验证所有实际新建输出；测量批量与低内存写入，比较文件结构和业务输出。
3. **标准修改计划**：先定义语义操作模型、前置条件、保留契约、报告关联及计划校验；把现存直接调用 umya 的候选业务适配器收敛到该边界。
4. **候选编辑器矩阵**：按已盘点的真实操作组合逐批测试 OOXML Patch、umya 与现有 openpyxl worker；只有结果显示 Excelize 有实际收益时，才开发批量 Go worker。不得把实现四种引擎作为迁移前置要求。
5. **业务增量迁移**：每个功能独立 fixture、Python 基线、Rust 结果、差异分类、失败回退和旧实现保留/切换条件；不合并 audit-core 与 pytools-core。
6. **交付验收**：两个 deb 分开，x86_64/glibc 2.28 构建；核心 Linux 路径不要求用户预装 Python、Go、Office/WPS/LO/JVM。产品若启用 openpyxl worker，将隔离 runtime 随包内置并做目标机动态依赖验收。未来 Windows COM 编辑不进入 Linux 包；若 Linux 启用 LibreOffice Calc 计算，先明确独立运行时来源并验证目标平台。外部能力缺失不阻塞其他功能。

每个修改操作矩阵至少记录：输入扩展名与包特征、计划操作组合、候选引擎/版本、样本来源、保留/丢失/损坏/不适用/未验证状态、性能和峰值内存、失败类型、回退引擎、报告路径。API 不存在或样本缺失都不能写成“保留”。

## 7. 本轮未完成与风险

- 已建立静态源文件/函数入口清单；尚未完成每个菜单到 helper、配置、输入输出的调用链映射，因此 PYTHON_EXCEL_OPERATION_INVENTORY.md 仍是初始入口清单，不宣称穷尽所有动态/间接 openpyxl 调用。
- 所有现存编辑器证据来自合成 xlsx；真实模板数为 0。umya 的 18/18 是带列定义补丁的单一合成场景。
- S2 文件级 service 对拍只使用合成配置/工作簿；当前与上期复用同一报表内容并更改文件名日期。外部 reader 完成事件中的引擎/耗时消息未逐字对拍；该 service 仍未连接 CLI 或生产默认路由。
- xls/xlsm/xlsb/csv/et 的业务样本与保存行为尚未形成统一矩阵；xlsb 的合并区域读取 API 存在已记录缺口。
- ExcelModifier 已有 UmyaXlsxBackend/OoxmlHeaderTextBackend、OpenpyxlXlsxBackend 与对应 validator；pytools capability dispatcher 注册 openpyxl。没有已验收的生产 capability JSON；真实模板、runtime staging/deb、目标 glibc 和平台均未验收。Excelize worker 与 Windows Excel/WPS COM 修改 worker 未实现。sidecar 同类转换已接入；`recalculate` 候选已实现，含 LibreOffice 合成 `SUM` 测试，但真实业务公式、Excel/WPS 等价性、超时和目标 Linux 离线运行时仍未验收。LibreOffice 基础公式测试另有用户实测反馈，范围待明确。最新状态以代码及日期化修改记录为准。
- 当前没有首次 Git 提交。本轮不提交、不重置、不清理任何既有未跟踪文件。
