# py_statistics_tools 功能迁移清单

## 基线与状态口径

- 只读源项目：C:\Users\AZI\Desktop\VibeCode\4.pytools-main
- 主入口基线：PyCode/main.py 的用户菜单和 dispatch 映射；继续追踪 PyCode/core 中的规则、配置、文件处理与工作簿调用链。
- NoComPyTools 是同一产品目标的变体参考；共有和独有行为尚未逐项比较。独有行为先登记为待确认，不静默丢弃，也不生成第三个产品。
- 状态统一使用保留、丢失、损坏、不适用、未验证；源代码入口和真实业务样本尚未对拍的功能一律不得标为完成。
- 工作表提取以不带批注普通版为准；COM 批注版不迁移。Python 运行模式切换和 PDF 导出不迁移。

## 菜单功能清单

| 菜单组 | 来源入口中的功能 | 迁移边界 | 当前状态 |
|---|---|---|---|
| 时序 | 表头路径比对；时序提取；宽表汇总（行头加后缀）；宽表汇总（行头不加后缀）；预校验 | 迁移时序/路径映射规则、输出列和异常行为 | PT-1-3/1-4 两种行头模式已通过合成 Python/Rust 对拍和显式安全目标提交；真实样本、生产 capability、Windows 原生执行和目标平台仍未验收，整体状态未验证 |
| 去重 | 按批注检查；按批注删除；按配置检查；按配置删除；按配置追加到目标；按配置预校验 | 批注读取使用通用 Excel 原语；保留检查与执行的区别 | 六项均有合成语义对拍；PT-2-1 有合成 CLI→ExcelModifier→openpyxl 执行验证，PT-2-2 有合成 Excel COM 对拍及 CLI→ExcelModifier→umya 执行验证，PT-2-4 有合成 CLI→ExcelModifier→umya 执行验证；PT-2-5 新目标不覆盖发布及既有目标 `--commit-target` 提交通过合成测试专用 capability 验证；生产能力、真实样本和其余编辑能力仍未验收 |
| 打印 | 按批注保留源格式；按批注快速复制；按配置准备打印文件；按配置预校验 | 只写页面设置并准备 xlsx；PDF 菜单不迁移 | 3-1/3-2 计划语义及 3-4 预校验通过合成对拍；3-3 模式 1/2/3 已通过合成端到端候选对拍，并按用户合同拒绝已有目标；生产能力、真实模板与目标平台仍未验证 |
| 转换与文件命名 | 批量重命名；Excel 格式转换；Word 格式转换；批量修改 Sheet 名；按内容加前缀重命名；取消内容前缀；局部映射重命名；提交缺失检查；文件排序前缀；取消排序前缀 | 保留文件操作与同类格式转换；首批 sidecar 为 .et→.xlsx、.xls→.xlsx、.doc→.docx；不迁移 Python 模式切换 | 未验证 |
| 配置 | 初始化/修复配置表；查看配置并从磁盘重新读取 | 盘点配置模板、默认值、修复行为、字段映射与错误提示 | 5-1 新建、5-2 查看通过合成及真实只读对拍；5-1 修复已在合成及真实配置临时副本上通过 openpyxl worker 候选执行与语义对拍；生产 capability、完整文件保真与目标平台仍未验收 |
| 汇总 | 按使用区域汇总；按批注汇总 | 逐项确认坐标、模板/源文件关系、行顺序和输出格式 | PT-6-1 已接入真实 summary-used-range CLI；合成 XLSX 对拍（2 个工作簿、13 行）值/类型、样式、维度、冻结窗格、公式/合并/名称/页面属性通过，Excel 错误单元格类型保留。真实业务样本、XLSM、NoComPyTools 独有格式和目标平台仍未验收；PT-6-2 仍仅有合成对拍，整体未验证 |
| 一二批处理 | 拆分村镇银行数据；机构名称标准化；删除分机构表外资行；外汇页眉修改；地区总分校验；工作表数据提取；涉农贷款比上月修正；季度涉农贷款比上月修正及分地区核对 | 仅工作表提取普通无批注版；对 COM 路由的业务行为与普通实现差异做盘点 | 7-1/7-3/7-7/7-8 计划语义合成对拍通过；7-7/7-8 精度与失败地区回归分支已补测；7-2/7-4/7-5/7-6 有各自候选证据；既有簿执行和整组验收仍未完成 |
| 归档 | 按日期预览；按类型预览；执行归档计划；按日期直接复制；按类型直接复制 | 保留预览、冲突检查、复制结果和归档规则 | Rust CLI 候选；5 类工作流通过合成 Python 对拍，真实业务目录/配置与目标平台未验证 |
| 快捷工具 | 打开配置文件；打开 output 最新文件；打开 output 文件夹 | 迁移到目标平台可用的文件打开/定位体验，系统集成细节待确认 | Rust 三入口通过 WSL 假 `xdg-open` 合成集成验证；旧版 Windows `os.startfile`、Windows 桌面和目标平台尚未实测 |

## 输入与转换矩阵

| 格式 | 计划 | 验收状态 |
|---|---|---|
| xlsx | 默认读写格式；calamine/OOXML reader 与 writer/editor 按对拍范围使用 | 业务模板未验证 |
| xls | Rust reader 能力逐项验证；需改 Sheet 名时显式转为 xlsx，并检查工作表及数据 | 未验证 |
| xlsm | 按真实样本确认值读取、VBA 部件保留和输出格式限制；不执行宏 | 未验证 |
| xlsb | 按真实样本确认 calamine 读取范围；不能从 zip 容器推断可修改 | 未验证 |
| csv | 盘点每项功能是否接收 CSV，确定编码、分隔符、空值和列类型 | 未验证 |
| et | 用户显式触发 sidecar 转换为 xlsx；检查 sheet 数、名称及数据 | 未验证 |
| doc | 只允许同类文字文档转换到 docx；由 sidecar 显式调用，转换后校验可打开性和内容 | 未验证 |

## 当前 Rust 代码与缺口

已有日期解析、路径映射、规则引擎、配置加载、工作簿抽取和机构名称规范化候选模块。日期解析 helper 已使用 22 个共享向量验证原 Python 与 Rust：22/22 匹配；关键字/列规格/路径标准化 helper 又以 28 个共享向量验证：28/28 匹配。配置加载已覆盖 `load_global`、`load_global_value`、`load_timeline_rules`、`load_path_maps`、`load_dedup_tasks`、`load_print_tasks`、`load_extract_tasks`、`load_foreign_banks` 和 `load_institution_mapping`；前四个 API 的基线见 `probe/reports/config_loader_parity.json`，新增五类 API、旧版去重表回退及缺表行为见 `probe/reports/config_task_loaders_parity.json`（11/11 比较组通过）。两组均为合成 xlsx 配置，不代表真实用户配置或消费这些配置的业务流程已验收。时序工作簿抽取适配器以 calamine/共享 `excel-reader` 读取 xlsx，并把合并区域交给规则引擎；一份合成工作簿的六个抽取单元格与原 Python `extract_cells` 结果一致。

普通版工作表提取除 `sheet_extract` 纯逻辑（Python helper 11/11，`probe/reports/sheet_extract_parity.json`）外，现有 area-only 文件流程候选 `run_extract_sheets_area_candidate`，并通过 `pytools extract-sheets-area --config <配置簿> [--output-dir <目录>] <源工作簿>...` 接入 Rust CLI。它加载现有“工作表提取”任务配置，使用共享 `excel-reader`/Calamine 读取值，以 rust_xlsxwriter 新建输出簿；支持文件/工作表关键字匹配与排除、A1 多区域横向/纵向拼接、空行列压缩、输出路径和工作表重名处理。写出后重开校验工作表数量、名称、尺寸和非空格数量；输出先写临时文件，经校验后用不覆盖目标的硬链接提交，提交失败时撤销本轮输出。合成 XLSX 对拍覆盖两个有效源工作簿和一个损坏输入，输出 4 张表至 1 个新工作簿，统计 files=2、extracted_sheets=4、output_files=1、failed_files=1；值、表名、尺寸与只读 PyCode 基线一致，输入文件字节未变。Python 基线和 Rust 对拍报告分别为 `probe/reports/sheet_extract_area_python_baseline.json`、`probe/reports/sheet_extract_area_parity.json`，人工记录见 `GPT修改记录/20261023_081520_pytools工作表区域提取Rust候选与CLI对拍.txt`。

此候选仅覆盖新建 XLSX 的区域值提取：源 Python 同样以 `data_only=True` 读取；本候选不复制公式文本。整表模式（样式、合并、行列尺寸等保真）和追加既有输出簿会显式拒绝；日期、时长、错误值及非有限数目前显式拒绝。夹具不含公式，日期拒绝仅作候选安全边界测试。源文件由调用方显式列出，不含目录扫描、现有产品菜单接入、日志等价、真实模板或非 XLSX 格式验证。因此仅有一个 CLI 候选切片通过合成对拍，不能将工作表提取菜单或完整功能标为已迁移。

机构名称规范化已有 `.xlsx` Rust CLI（`pytools institution-normalize --config ... <workbook>...`），默认只生成计划；执行须显式 `--apply` 并提供匹配的 ExcelModifier capability。候选按原规则在“分机构”工作表统一选择 A 列或 B 列，精确映射优先，其次按映射顺序不区分大小写做子串映射；多文件用 Rayon 并行，同一路径重复出现时退回顺序处理。原 `run_normalize_institution` 候选与 Python `_process_sheet` 在一份四工作表合成簿上对拍，Rust probe 的 sheet 名、20 个值、1 条公式、1 个合并区域和 1 条批注均保留；文件统计与 Python 一致（3 个匹配工作表、2 个改动工作表、3 行改动）。openpyxl 元数据检查还验证了该夹具的样式、行列尺寸、批注、合并区域、页面设置和页眉页脚。`build_institution_normalize_plan` 通过共享 `excel-reader` 构造 3 个 `SetCellValue` 操作，并与 Python 基线的统计及变更格一致，计划对拍报告为 `probe/reports/institution_normalize_plan_parity.json`，该单独报告未执行修改引擎。另有显式忽略的集成测试在合成四工作表夹具上临时登记 fixture-scoped capability，实际执行 CLI → ExcelModifier → `openpyxl_worker`，并对拍值、公式、批注、合并、样式、行列尺寸、页面设置、页眉页脚和源文件安全性；执行报告为 `probe/reports/institution_normalize_openpyxl_execution_20261024.json`，状态明确为 `synthetic_passed_not_production_acceptance`。worker 修复了 openpyxl 写出时遗漏的 `xml:space="preserve"`，并由 Rust 读取候选临时文件完成保留项验证。该 capability 仅为测试临时文件，不产生生产 Passed 登记，也不改变默认路由。以上仅为合成 XLSX 证据；无经授权真实样本，`.xlsm`、真实模板、Windows 原生执行、目标 Linux/glibc 2.28 和 deb 尚未验证；PT-7-2 生产状态仍为未验证。旧直接 umya 候选未改为 ExcelModifier 默认路由。

文件排序前缀（4-9/4-10）已有纯逻辑 helper 对拍，报告为 `probe/reports/file_sort_prefix_parity.json`。4-9 `run_file_sort_prefix_49` 使用共享 `excel-reader`/Calamine 读取“文件排序配置”，按输入顺序匹配并改名；目标已存在时自动添加 `_1` 等后缀，不覆盖占用文件；使用共享 `excel-writer`/rust_xlsxwriter 新建“文件排序前缀结果”工作簿。报表写入失败时会撤销本轮已完成的改名。该实现现已接入 `pytools file-sort-prefix --config <配置簿> [--output-dir <目录>] [--log-dir <目录>] <文件>...` Rust CLI。与只读 PyCode 的 `run_file_sort_prefix_49` 在一份合成配置簿和 8 个文件/路径案例上对拍通过，API 与 CLI 均验证统计、报告路径、文件名集合及文件内容；CLI 为成功/失败项写带时间戳的日志并输出到 stderr。日志消息/成功项行数已做合成断言，但未逐字与 Python logger 的时间、并发和重复运行行为对拍。完整机器报告为 `probe/reports/file_sort_prefix_operation_parity.json`。原 GUI 菜单/dispatcher、真实目录和目标平台仍未验收。

4-10 `run_remove_file_sort_prefix_410` 已接入 `pytools remove-file-sort-prefix --config <配置簿> [--output-dir <目录>] [--log-dir <目录>] [<文件>...]` Rust CLI。它读取排序规则和可选的 `重命名配置!O:P:Q` 内容单元格规格；直接以排序前缀开头的文件名不打开工作簿，排序前缀位于内容前缀之后时才用共享 Calamine reader 读取指定内容格，并只在内容前缀吻合时移除排序片段。改名使用不覆盖目标的硬链接预留，报告写入失败会回滚本轮改名；结果通过共享 rust_xlsxwriter 写入“取消文件排序前缀结果”。CLI 写时间戳日志文件和 stderr，允许空文件列表并输出兼容 Python 的空报告。合成 XLSX 对拍覆盖 10 个文件/路径案例：规则 4、改名 5、跳过 5、失败 0；报告行、文件名集合与 Python 基线一致，并检查改名不改变输入文件字节。空输入 API 与 CLI 均验证：保留报告工作表但不写表头或单元格。完整机器报告为 `probe/reports/file_sort_prefix_remove_parity.json`，Python 明细基线为 `probe/reports/file_sort_prefix_remove_python_baseline.json`。原 GUI 菜单、日志时序/并发行为、真实目录和目标平台尚未验收，不代表整个菜单功能已迁移。

4-8 局部文件名映射 `run_local_filename_map_48_candidate` 已接入 `pytools rename-local-map --config <配置簿> <文件>...` Rust CLI，并与只读 PyCode 的 `run_batch_rename_local_map_48` 对拍。它从“重命名配置”工作表按 `原文件名片段`/`新文件名片段` 表头读取映射，按配置顺序在文件 stem 中逐项替换，并保留最后一个扩展名；允许任意扩展名，空目标表示删除片段。改名目标冲突时使用 `_1` 等编号且不覆盖已存在文件；目标名不合法、过长或文件系统操作失败时逐项跳过并保留源文件。合成场景 10 个输入路径，含 6 条映射：Python 与 Rust 均改名 5、跳过 5，映射顺序、中文片段、空目标、冲突后缀、缺失路径、目录、未命中和改名错误的最终目录集合一致；Rust 测试检查输入文件字节保持及冲突目标不变。只读 Python `test_rename_local_map_48.py` 和 `test_convert_rename_edges.py` 共 19 项通过；Rust 4-8 集成测试现为 3 项（含直接 CLI 调用）及模块纯逻辑测试 2 项通过。机器报告为 `probe/reports/local_filename_map_48_parity.json`，Python 基线为 `probe/reports/local_filename_map_48_python_baseline.json`。该 Rust 命令不等于原 GUI 菜单集成或生产默认路由；真实目录和目标平台尚未验证。

4-6 内容前缀添加 `run_content_prefix_rename_46_xlsx_candidate` 已接入 `pytools rename-by-content --config <配置簿> <文件>...` Rust CLI，并与只读 PyCode 的 `run_batch_rename_by_content_46` 对拍。它从唯一启用的 `重命名配置!O:P:Q` 读取规格，用共享 Calamine reader 读取 `.xlsx/.xlsm` 指定单元格；沿用缺失工作表的 `sheet!A1` 回退、内容项已出现在文件名中则跳过、前缀最多 180 个 Unicode 字符、目标冲突添加 `_1` 等后缀的行为。候选只改文件名、不保存工作簿；改名通过硬链接预留避免覆盖目标。合成场景 12 个输入路径：改名 5、跳过 7；API 与 CLI 输出统计、最终文件名集合均与 Python 基线一致，Rust CLI 测试逐文件检查输入工作簿字节未变并确认占用目标未覆盖。既有 Python `test_rename_by_content_46.py` 13/13 是前一轮历史证据；本轮未重跑 Python。Rust 4-6 集成测试 2/2、pytools-core 49 项、workspace 121 项通过。机器报告为 `probe/reports/content_rename_46_parity.json`，Python 基线为 `probe/reports/content_rename_46_python_baseline.json`。该证据仅覆盖合成文本格与 `.xlsx/.xlsm`；源 `.xls/.xlt` COM 分支、真实宏、数字/日期/公式缓存值、真实目录、GUI dispatcher、Windows 原生 Rust、跨平台文件系统和目标 Linux/deb 仍未验收。CLI 是显式文件列表操作，不是 GUI 菜单接入或生产默认路由。

4-7 内容前缀取消 `run_content_prefix_unrename_47_xlsx_candidate` 已接入 `pytools unrename-by-content --config <配置簿> <文件>...` Rust CLI，并与只读 PyCode 的 `run_batch_unrename_by_content_47` 对拍。它读取 `重命名配置!O:P:Q` 中唯一启用的内容规格，使用共享 Calamine reader 读取 `.xlsx/.xlsm` 指定单元格；缺失工作表/空值沿用 `sheet!A1` 作为回退项，移除文件名开头匹配的内容项并保留已配置排序前缀。候选不保存工作簿，只执行逐文件改名；目标已存在时以不覆盖的硬链接预留并尝试 `_1` 后缀，未匹配、读取失败或未验收扩展名逐项跳过。合成主场景 13 个输入为改名 8、跳过 5；无“文件排序配置”工作表的旧配置场景为改名 1、跳过 1；Rust API 与 CLI 的统计和最终文件名集合均与 Python 基线一致，并检查已改名文件字节未变化。旧 Python 聚焦测试 13/13 是前一轮历史证据，本轮未重跑；Rust 4-7 集成测试 4/4、pytools-core 49 项、workspace 121 项通过。机器报告为 `probe/reports/content_unrename_47_parity.json`，基线为 `probe/reports/content_unrename_47_python_baseline.json`。当前仅为合成样本候选；真实 `.xlsm`、源 `.xls/.xlt` COM、日期/公式缓存格、真实目录、GUI dispatcher、不同文件系统、目标 Linux/deb 未验证。CLI 不是 GUI 菜单接入或生产默认路由。

地区总分校验 `feature_1_5_region_sum_check.py::run_region_sum_check` 的非 COM 主路径已由 Rust `region_sum::run_region_sum_check` 实现，并通过 `pytools region-sum-check` 接入 Rust CLI。命令接受显式工作簿列表；可从 `--config` 读取全局输出/日志目录，也可用 `--output-dir`、`--log-dir` 覆盖；不传配置时默认写入 `./output` 和 `./logs`。读取通过共享 `excel-reader`/Calamine，写入通过共享 `excel-writer`/rust_xlsxwriter；COM 分支不迁移到此命令。5 个合成 XLSX 场景的文件/工作表/错误/失败统计与 Python 一致（4/3/1/1），结果单元格、公式文本（按 Python 公式前缀归一）、数据区域尺寸、样式和列宽（绝对差不超过 0.01）通过对拍，日志关键记录存在且输入文件字节不变。Python 基线为 Python 3.12.10、OpenPyXL 3.1.5；源函数与 logger 的 SHA-256 见 `probe/reports/region_sum_python_baseline.json`。为保持汇总标记公式的空缓存，结果生成后使用 `ooxml-patch::clear_empty_formula_cache` 仅清除对应普通公式单元格的空缓存；单元格坐标按结果行数计算，不是通用公式修改能力。新增的目标名冲突与输出目录失败测试确认既有结果/日志及输入不会被覆盖或更改。完整机器报告为 `probe/reports/region_sum_parity.json`；本轮新增 CLI 集成对拍在 `tests/region_sum_parity.rs` 直接启动产品二进制并与同一基线比较统计、输出单元格/公式及输入字节。真实业务模板、非 XLSX 格式、性能/峰值内存、Windows 原生 Rust、麒麟/UOS glibc 2.28 与 deb 均未验收；原 GUI 菜单仍未迁移，Python 原功能保留作回退，因此不能将地区校验菜单标为已迁移。

外汇页眉单位修正 `feature_1_4_fx_header.py` 已有窄范围 Rust 候选 `fx_header::build_fx_header_plan`：业务层读取共享 `excel-reader::read_header_inventory` 的 odd/even/first 页眉文本，只生成 `ReplaceHeaderText` 计划，经 ExcelModifier 路由到 OOXML Patch。PyCode 与 NoComPyTools 对应 Python 源文件当前 SHA-256 相同；普通非 COM 行为以 PyCode 基线为准。Python 基线统计为文件 1、外汇工作表 2、修改工作表 1、失败 0；Rust 候选验证外汇工作表计数 2 和计划修改工作表计数 1 与 Python 相符，并成功执行/校验单文件输出。验证各页眉文本、非目标工作表、值/公式/合并/批注/条件格式/数据验证以及补丁范围；当前尚无 Rust 批量文件统计聚合。机器报告为 `probe/reports/fx_header_parity.json`，Python 基线为 `probe/reports/fx_header_python_baseline.json`。PT-7-4 候选仅接受 `.xlsx`，已接入 Rust CLI `fx-header-plan`（默认 plan_only，执行须显式提供匹配的 capabilities/platform），未接原 Python 菜单、不执行源路径原位替换；旧 Python 原功能保留。真实业务模板、`.xlsm`/VBA、目标平台和 deb 未验证，不能将页眉菜单标为已迁移。

### 3-3 配置驱动的批量文件重命名

`PyCode.core.convert_33_34_35_37.run_batch_rename_33` 已由 `pytools_core::batch_rename::run_batch_rename_33` 改写，并通过 `pytools rename-batch --config <工作簿> [--log-dir <目录>] [<文件或目录>...]` 接入 Rust CLI。实现读取 `重命名配置` 的简称/全称、全称/代码映射及数据日期/报表名称；按配置首次插入顺序执行不区分大小写的子串匹配，重复配置键覆盖值但保留原顺序；缺少代码、无命中、原名不变、路径不存在及 OS 改名错误按源函数计入跳过；目标冲突使用 `_1` 等后缀。支持文件和目录，不读取或改写工作簿内容。

Python 基线由源 `run_batch_rename_33` 对同一合成配置及路径场景实跑，Python 3.12.10、pandas 2.2.3、openpyxl 3.1.5：改名 3、跳过 4，最终文件/目录名与 `probe/reports/batch_rename_python_baseline.json` 一致；覆盖匹配顺序、代码缺失、无命中、缺失路径、同名跳过、连续冲突后缀、目录内容和 ASCII 大小写匹配。Rust `tests/batch_rename_parity.rs` 验证 API 与 CLI 结果、日志、输入/目标文件内容及配置字节不变，汇总见 `probe/reports/batch_rename_parity.json`。对原只读 `PyCode/config/config.xlsx` 还做了两项验证：CLI 空输入读取成功（0/0）；随后从真实配置选择一个有效别名，在系统临时目录创建合成文件并由 Rust CLI 改名（1/0），确认原配置字节与临时文件内容都不变。结果摘要见 `probe/reports/batch_rename_real_config_smoke.json`。未对原文件或真实业务文件执行改名。

当前证据仅覆盖脱敏合成配置与临时路径场景；真实业务输入文件、真实配置对应文件名覆盖率、重名冲突在目标文件系统上的竞争行为、麒麟/UOS 和 deb 未验证。此局部 Rust CLI 候选不代表 Python GUI 菜单整体已迁移，也不改变 3-3 的总体“未验证”状态。

上述结果不代表完整菜单或真实业务模板已验收。尚需：
- 将每个菜单命令映射到源函数、配置项、文件格式、输出契约和回归用例。
- 对比 PyCode 与 NoComPyTools，列出共享与独有行为。
- 为尚未覆盖的配置表、业务工作簿及各迁移功能建立 pytools 专属 fixtures、probe config、对拍输出和日期报告。
- 时序抽取目前没有批量目录调度、CSV 输入、结果工作簿写入或向目标簿追加；日期单元格序列值、xlsb 合并区域和真实模板仍未验证。
- 普通版工作表区域值提取已有配置加载、显式文件列表批处理、Rust CLI、Python 基线和新建 XLSX 输出候选；尚缺现有产品菜单/dispatcher 接入、目录发现、日志等价、真实模板/多格式验证。整表提取现已有 XLSX `ImportRange` 编辑计划与 umya 候选实现，可复制值/公式缓存、样式、合并、批注开关、行列尺寸/隐藏/outline 及 A1 矩形内空白格。一个合成样本上，Rust 输出与源工作表属性（值、样式、合并、行列尺寸/隐藏/outline）0 项差异；输入哈希不变。旧 Python `_copy_full_sheet_with_style` 生成的跨簿输出在同一带样式样本上被 openpyxl 读取样式时触发 `IndexError`，因此此次以源工作表属性作为语义基准，不宣称旧输出文件级保真。没有已验收的生产 capability，完整整表 CLI 仍要求匹配的 `--capabilities` 和 `--platform`。追加既有输出簿和其他格式仍未实现或验证。
- 机构名称规范化已有显式文件列表和配置 CLI；仍未接入原 GUI 菜单/dispatcher 或目录发现，当前候选只接受 `.xlsx`。合成夹具已有通过临时测试 capability 的 ExcelModifier/openpyxl worker/validator 执行证据，但未形成生产 capability；计划报告与执行测试覆盖范围不同。源实现支持的 `.xlsm` 保留行为、真实模板中的全部 OOXML 部件、真实样本、Windows 原生执行及目标平台仍需验收。
- 地区总分校验 `region_sum::run_region_sum_check` 现已接入 `pytools region-sum-check` CLI，并通过合成 XLSX 对拍；原 GUI 菜单、真实模板和非 XLSX 格式、执行性能/峰值内存、Windows 原生与目标 Linux/deb 均未验证，现有 Python 功能仍保留。
- 4-6/4-7、4-8、4-9 和 4-10 均已接入显式文件列表 Rust CLI；只有 4-9/4-10 写操作日志。4-6/4-7 CLI 对拍仅覆盖合成 `.xlsx/.xlsm`，源 `.xls/.xlt` COM 读取未迁移，真实宏、数字/日期/公式缓存单元格和不同文件系统仍未验证；CLI 不等于原 GUI 菜单/dispatcher 集成。4-9/4-10 日志时间、并发及重复运行行为尚未逐项与 Python logger 对拍；4-8 尚未验证真实配置、目标平台和非 UTF-8 文件名。上述操作的真实目录、Windows 原生 Rust 构建和目标 Linux 包均未验收。4-10 的 `.xls/.xlt` 内容读取、多单元格规格及真实文件仍未验证，源 COM 内容读取仅作基线，不进入 Rust 主路径。旧 COM 业务读取未迁入 sidecar。
- 拿到经授权的脱敏业务文件；当前合成测试不能证明真实模板兼容。

### 归档工具：五类工作流的合成对拍

`PyCode.core.convert_33_34_35_37` 中的按日期预览（含可选扩展名配置）、按类型预览、按日期直接复制、按类型直接复制、执行归档计划，现均有 Rust 实现并接入 `pytools` CLI。`tests/archive_parity.rs` 将每个工作流的统计和报告行与只读 Python 基线逐行比较；对复制成功项另核对源/目标字节。六个基线场景的主要统计为：日期预览 12 个文件/1 个无效路径/13 行；带扩展名筛选的日期预览 12/1/13 且筛除 2 项；类型预览 12/1/13、3 条规则且筛除 2 项；日期直接复制 12/12 成功；类型直接复制 12 项中 10 成功、2 跳过；执行计划复制 13 项中 12 成功、1 跳过。类型配置和输入均为合成夹具，不是生产数据。

对拍期间修正了两处兼容差异：日期预览 CLI 增加与 Python 对应的可选 `--config <config.xlsx>` 扩展名筛选；执行计划时清除预览报告中的旧目标冲突备注，仅在复制时再次发生冲突后生成执行备注。Python 函数源文件 SHA-256 为 `9cfd0de33f6cc0cc7c665ace7c445feb125b24f0c201165bcf281a53f24617a3`。测试命令 `cargo test -p pytools-core --test archive_parity` 为 2/2 通过；本次变更后 `cargo test --workspace` 通过。机器报告为 `probe/reports/archive_parity_20261023.json`，Python 基线为 `probe/reports/archive_python_baseline.json`。

该结果仅证明当前合成夹具和 Windows 开发机上的 API/CLI 候选行为一致。真实配置及真实归档目录、GUI 菜单集成、权限/竞争条件、非 UTF-8 路径、麒麟/UOS 与 glibc 2.28 deb 尚未验收；因此不宣称归档功能整体已完成迁移。

### 去重工具：计划执行入口

现有 `dedup_ops` 覆盖批注检查/标记与删除、配置检查/标记与删除、按配置去重追加以及预校验的 Rust 计划构造。CLI 的批注与配置计划命令现可将计划写入候选文件并可选接收 `--capabilities <json> --platform <id>`，由统一 `ExcelModifier` 按能力登记和保留项验证执行；不提供能力证据时只输出/保存计划，不修改输入工作簿。2-5 的新目标 XLSX 由共享 writer 写到候选路径；既有目标使用 edit plan 且必须提供已登记能力文件。`dedup-append-execute` 默认保留候选模式；显式 `--commit-target` 才会在验证后提交目标。2-6 仅预校验。

计划执行入口本身仍受能力登记约束：`SetCellFillColor`、`DeleteRows`、追加涉及的 `ImportRange` 在真实 pytools 模板上的后端保留能力没有可用的生产验收记录；不得自行创建生产能力记录或切换默认。六项菜单现已完成一组合成 Python/Rust 语义对拍；PT-2-4 和 PT-2-5 的执行候选证据及限制见下文和 `probe/reports/dedup_append_execution_20261023.json`。新增的显式目标提交协议及安全校验见本清单“PT-1-3/PT-1-4 显式提交”条目；临时测试 capability 仍不能登记为生产能力。真实配置/工作簿、完整运行日志与交互进度、冲突/失败回退和目标 Linux/deb 仍待完成。机器状态记录为 `probe/reports/dedup_execution_wiring_20261023.json`。

### 2026-10-23 去重六项合成语义对拍

新增 `probe/compare_dedup_tools.py`，对旧版 `dedup_311.py` 与 Rust 计划/CLI 比较 2-1 批注查重、2-2 批注删除、2-3 配置查重、2-4 配置删除、2-5 配置追加新目标和 2-6 预校验。合成 XLSX 结果：批注查重 3 行、配置查重 5 行；批注删除 3 行、配置删除 4 行；新目标追加 7 行并过滤 7 行；预校验 ok/warn/fail 为 4/1/2。删除计划由 probe 应用到临时副本后与 Python 工作簿单元格投影比较；标记计划的填色坐标与 Python 输出比较。追加候选与 Python 输出比较工作表名和值。

对拍发现跨任务去重差异：同一目标表先由一个键列组合追加、后由另一组合追加时，旧 Rust 只从目标原表构造新键缓存，没有纳入前一任务已追加行，导致少过滤重复项。`dedup_ops` 现将已接受行保存在当前目标工作表状态中，新建选择器时从这些行建缓存，并在每次任务后同步已有选择器。复测 Python/Rust 追加和过滤统计均为 7，候选工作簿值投影一致。

机器报告为 `probe/reports/dedup_parity_20261023.json`。所有场景均为合成 XLSX；标记与删除尚未通过 production `ExcelModifier` 后端执行，2-5 仅验证新目标候选创建，既有目标编辑仍受能力证据门控。真实模板、运行日志、冲突/失败恢复、其他格式、Windows 发布构建、麒麟/UOS/glibc 2.28 和 deb 未验证。此结果不代表去重菜单整体完成迁移。

### 2026-10-23 PT-2-5 新建及既有目标追加候选执行

新增 `tests/dedup_append_execution.rs`，通过 Rust CLI 分别验证 PT-2-5 的两种目标状态。目标不存在时，两个追加任务通过 `excel-reader`/Calamine 读取源簿并由 `excel-writer`/rust_xlsxwriter 新建候选簿；合成输入共追加 4 行、过滤 4 行，公式文本保留，源簿和配置不变。目标已存在时，CLI 生成 `ExcelEditPlan`，仅在测试临时目录内生成与该合成计划精确匹配的临时 Passed capability，再经 `ExcelModifier → UmyaSpreadsheet 3.1.1` 生成候选文件；两个任务共追加 4 行、过滤 4 行。测试检查原目标、源簿及配置未被改写，并检查第二张未触及工作表的顺序、批注、合并区域、行高，以及目标工作表原有样式。

扩展 `probe/compare_dedup_tools.py`，将旧 `PyCode` 的 `run_append_by_config` 与 Rust CLI 既有目标候选作同输入对拍。此单任务场景在原目标已有 2 行时追加 3 行、过滤 1 行；工作表名、单元格值和公式文本投影一致，Rust 原目标字节不变。新目标场景仍为追加 7 行、过滤 7 行，工作表/值投影一致。Python 基线运行于 Windows Python 3.12.10，Rust CLI 及集成测试运行于 Ubuntu/WSL2 x86_64；测试数据是临时生成的合成 XLSX，不是实际用户模板或配置。机器报告为 `probe/reports/dedup_append_execution_20261023.json`。

测试：`py -3.12 pytools-core/probe/compare_dedup_tools.py` 通过；`cargo test -p pytools-core --offline --test dedup_append_execution -- --nocapture` 为 2/2 通过；`cargo test -p pytools-core --offline --test dedup_comment_delete_umya --test dedup_append_execution -- --nocapture` 为 3/3 通过；`cargo fmt --all -- --check` 通过；`cargo test --workspace --offline --locked -q` 通过。既有目标测试 capability 仅存在于临时测试目录，不产生生产能力登记，也没有改变默认路由。

仍未验证真实配置/工作簿、XLS/XLSM/XLSB/ET、未列出的 OOXML 部件、Windows 原生 Rust 构建、麒麟 V10/UOS V20、glibc 2.28、离线 deb、生产 capability 保留项和完整去重菜单。故 PT-2-5 生产验收状态仍为**未验证**；本候选证据只覆盖列明的合成操作批次与属性。

后续安全提交试点见下文 PT-1-3/PT-1-4 条目：`AppendTargetPlan` 记录既有目标的 SHA-256，`dedup-append-execute --commit-target` 在 capability 门控和候选验证后才提交；新目标采用不覆盖发布。该共享 CLI 执行能力的 PT-2-5 生产后端仍受真实 capability 验收约束。

### 打印与配置工具：计划执行入口

`print-comment-plan`、`print-config-plan` 和 `config-repair-plan` 现支持可选 `--capabilities <json> --platform <id>`，在计划输出后调用统一 `ExcelModifier`；未提供匹配的能力记录时仍只产出计划。批注打印计划另可通过 `--plan-output` 保存。`print-config-precheck` 和 `config-view` 是只读操作；配置初始化沿用 `excel-writer` 新建配置簿。此接线没有增加渲染、PDF、打印机或外部 Office 路径。

上述计划执行仍受能力证据门控：真实打印区复制、样式/合并/批注/行列尺寸与页面设置整批保留能力未完成 pytools 对拍，不能创建或启用对应默认 capability。配置修复模板结构、默认值和源项目报错语义也尚未对拍。机器状态记录为 `probe/reports/print_config_execution_wiring_20261023.json`。

### 2026-10-23 批注打印 3-1/3-2 计划语义对拍

新增 `probe/compare_print_comment_plans.py`，生成含两个批注标记区域、合并格、样式、批注和零值的合成源簿，分别调用旧版保留格式/快速打印函数与 Rust `print-comment-plan`。两条路径均识别 `B2:D4` 和 `F2:G3`，输出表顺序与名称一致；对拍了 ImportRange 起点、全/基础样式模式、批注和合并策略、零值隐藏及方向、缩放、纸张、打印区域、居中和页脚设置。机器报告为 `probe/reports/print_comment_plan_parity_20261023.json`。

旧 Python 输出由 openpyxl 在临时目录生成；Rust 只生成 ExcelEditPlan，没有执行修改。该结果是计划语义一致，不证明 umya/OOXML 后端实际写出保真。真实模板、3-3 配置打印、XLSM、Windows 发布构建、麒麟/UOS 与 deb 未验证；PDF 3-5 仍不迁移。

### 2026-10-23 配置初始化与查看（5-1/5-2）合成对拍

新增 `probe/compare_config_admin.py`，将 Python `initialize_or_repair_config` 的新建行为与 Rust `config-init` 输出比较，并验证 Rust `config-view` 的行和值、批注文本。11 张新模板表的顺序、所有非空单元格值、表头批注文本及作者字段一致；在已存在文件上再次调用 `config-init` 会拒绝操作，文件 SHA-256 不变。另只读检查旧项目真实 `PyCode/config/config.xlsx` 的 12 张表，表序、值/类型语义和批注投影匹配，文件 SHA-256 前后均为 `e000c784e85b63b4fc19fd45684aa8729dbd032238d78321fdc9fb4c33cee716`。机器报告为 `probe/reports/config_admin_parity_20261023.json`，Python 配置初始化源文件 SHA-256 为 `c6cd86e5275e63845b48db6029c14e09ee3edc2ec30057047999da244a14a0ab`。

对拍修正 Rust 注释写入行为：rust_xlsxwriter 默认把批注作者前缀拼进正文，现显式关闭此前缀、仍写入独立 `pytools` 作者元数据；同时补齐“机构映射表!是否为外资行”说明缺少的“1.3 使用”。真实配置中按日期格式存储的日期也曾显示为 Excel 序列值；`config-view` 现通过共享 reader 的 ISO 日期读取接口显示日期语义。真实配置仅只读，无文件内容变化。配置修复计划与候选执行证据见下方新小节；生产能力和目标平台尚未验收。

### 2026-10-23 PT-5-1 配置修复 openpyxl worker 候选执行

修正 `build_config_repair_plan`：配置示例中的空字符串按旧 openpyxl 保存后的 XLSX 空单元格语义生成 `Blank` 操作；计划只声明当前 openpyxl validator 能检查的值/类型、公式文本、批注与位置、样式、合并区域和工作表身份/顺序，不再宣称保留未验证的页面设置、行列尺寸、关系或 OOXML package 字节。计划仍通过 `ExcelModifier` 能力门控。

新增 `tests/config_repair_execution.rs`，临时装配隔离测试 runtime，并通过真实 `pytools config-repair-plan` CLI 调用 `ExcelModifier → openpyxl_worker`。合成配置覆盖错误表头/批注、缺表、全局默认及重复键、仅空表填充示例、删除旧表、重命名配置补充值和保留自定义表；计划含 258 个操作，候选的 sheet 顺序/值/批注与 Python 基线一致。集成测试另检查源文件不变、未触及自定义表的样式/批注/合并/行高/列宽，以及目标表数据。该操作组合的 Umya 候选未保留修复后的表头批注，Python readback 对拍失败，故本项不选 Umya。

扩展 `probe/compare_config_repair_plan.py`，对真实只读源 `4.pytools-main/PyCode/config/config.xlsx` 的临时副本分别执行旧 Python 修复和 Rust 候选修复：Python 报告新建 0 张、修复 1 张；Rust 计划 2 项；实际路由为 `openpyxl_worker`，表序、值及批注正文/作者投影一致。真实源配置 SHA-256 为 `e000c784e85b63b4fc19fd45684aa8729dbd032238d78321fdc9fb4c33cee716`，运行前后未变化；Rust 输入副本也未改写。报告还覆盖合成配置场景。机器报告为 `probe/reports/config_repair_execution_20261023.json`。

本次使用 PyTools 原项目 `.venv`（Windows Python 3.12.10/openpyxl 3.1.5）运行基线；Rust CLI 和隔离 worker 在 Ubuntu/WSL2 x86_64 执行（rustc/cargo 1.99.1，worker openpyxl 3.1.5）。临时 Passed capability 只用于上述测试副本，没有生成生产 capability，也未改变默认路由。合成与真实配置投影对拍通过；Rust 集成测试需设置 `RUST_WORKSPACE_OPENPYXL_PYTHON` 并用 `--ignored` 显式执行。

尚未验证真实业务数据簿中的全部 OOXML 对象、页面设置/行列维度的完整总体保真、其他输入格式、Windows 原生 Rust 构建、麒麟 V10/UOS V20、glibc 2.28 和离线 deb。因此 PT-5-1 修复生产验收状态仍为**未验证**；该候选只证明当前配置对象和所列断言。

### 2026-10-23 打印配置预校验（3-4）合成对拍

新增 `probe/compare_print_precheck.py`，对旧版 `print_310.py::run_print_config_precheck` 与 Rust `print-config-precheck` 比较计数。合成配置包含有效项、FitToPagesTall 为 0 的旧版默认处理、负缩放警告、方向同时指定横纵、无效模式、缺失源工作表、无效 A1 区域、空目标路径和被禁用的无效行；Python/Rust 均为 ok=2、warn=2、fail=4。

对拍修正 Rust 的旧版语义：零值 `FitToPagesTall` 在 Python 预校验的 `or 1` 默认表达式下变成 1，不产生缩放警告；负值产生警告。非空方向只有明确且不冲突的“横向”或“纵向”可用，同时含两者时按无效方向警告。机器报告为 `probe/reports/print_precheck_parity_20261023.json`。

本节仅验证合成配置和源 XLSX 的预校验计数；现有真实配置中的打印任务表没有任务行。PT-3-1/3-2 与 PT-3-3 的后续合成执行候选证据见下方各自小节。整组打印菜单仍未验收；真实模板、XLSM、Windows 发布构建、麒麟/UOS 与 deb 未验证。

### 2026-10-23 PT-3-1/PT-3-2 按批注打印执行候选对拍

新增 `tests/print_comment_execution.rs` 并扩展 `probe/compare_print_comment_plans.py`：用合成 XLSX 调用只读 Python 基线 `print_310.py::run_print_keep_by_comment` / `run_print_fast_by_comment`，再调用 Rust `print-comment-plan` 的实际执行路径。Rust 计划从共享 `excel-reader` 获取标记区域并生成批量操作，执行候选为 `ExcelModifier → UmyaSpreadsheet`；通过仅存于测试临时目录的 capability 执行，未登记生产能力。

Keep 与 fast 两种模式均逐项比较工作表顺序/名称、单元格值与类型、字体/填充/边框/对齐/数值格式/保护、批注、合并区域、行高、有效列宽、方向/缩放/纸张/打印区域/居中/页脚；浮点使用相对及绝对容差 `1e-15`，同宽相邻 OOXML 列区间按每列有效宽度展开比较。Python 与 Rust 结果一致，源文件字节保持不变。报告：`probe/reports/print_comment_execution_parity_20261023.json`；计划语义报告：`probe/reports/print_comment_plan_parity_20261023.json`。

对拍发现并修正 Umya 在删除输入工作表后仍使用旧 `localSheetId` 写打印区域的问题；打印区域现在在全部工作表操作结束后按最终顺序写入。Keep 模式还显式写入 openpyxl 对未设置列宽的默认值 `13.0`，随后由 `ImportRange` 覆盖源表中显式列宽，以复现旧打印函数的有效列宽。

这只是合成 XLSX 的执行候选验证：没有真实打印模板或 `.xlsm` 样本，未验证 Windows 原生 Rust 构建、麒麟/UOS、glibc 2.28 或离线 deb；未创建生产 capability，也未改变默认路由。此候选通过 ExcelModifier 生成新工作簿，而 workspace 目标对新建 XLSX 优先使用 `rust_xlsxwriter`；在评估生产路径前仍需对照该新建 writer 合同完成候选评估。PT-3-1/PT-3-2 生产验收状态仍为**未验证**。

### 2026-10-23 PT-3-3 配置打印模式 1/2/3 执行对拍

修正 `build_config_print_plans`：新输出使用配置中的 `目标工作簿` 路径，不再另写随机候选文件；目标路径已存在（包括符号链接）时，预校验和计划构建均拒绝。计划阶段不创建目标目录；只有显式提供 capability 并进入实际执行时，CLI 才创建目标父目录。通过 `print-config-plan` → `ExcelEditPlan` → `ExcelModifier` → `UmyaSpreadsheet` 执行，计划/执行仍受 capability 门控，本轮的 Passed 记录只在临时测试目录使用。

扩展 `tests/print_config_new_target.rs`：已有目标测试确认 precheck/plan 均失败、目标字节未变且不创建 output 目录；新目标测试确认 plan-only 不创建目标父目录，实际执行按配置路径写出并读回工作表、值、样式、批注、合并、行列尺寸及页面设置。扩展 `probe/compare_print_config_execution.py`，分别以旧版 `print_310.py::run_print_config_all_modes` 和 Rust CLI 执行模式 1/2/3，对比工作表、值/类型、字体/填充/边框/对齐/保护/数值格式、批注、合并、行列尺寸及打印设置；三个模式的单任务合成样本均通过。模式 2 的列宽浮点结果相差约 1 ULP，按 `rules/13 §9.1` 使用相对/绝对容差 `1e-15` 比较。报告为 `probe/reports/print_config_execution_parity_20261023.json`。

这是合成 XLSX 候选验证，不证明真实配置、真实打印源、同一目标下多任务组合、未列 OOXML 部件或目标平台保真。用户确认的“输出新文件、已有目标拒绝”合同与旧 Python 可打开并修改既有目标的分支不同；Rust 候选按用户合同实现。未创建生产 capability，未改变生产默认路由，PT-3-3 生产验收状态仍为**未验证**。Rust 集成测试与实际对拍在 Ubuntu/WSL2 x86_64 执行，因为当前 Windows 会话没有 `cargo` 命令；Windows 原生 Rust 构建、麒麟/UOS、glibc 2.28 和离线 deb 未验证。

### 2026-10-23 快捷工具（9-1/9-2/9-3）隔离系统集成验证

新增 `probe/compare_quick_tools.py`，以临时配置簿和临时目录调用 Rust CLI，并将 `xdg-open` 替换为记录参数的假启动器。9-1 打开配置文件、9-2 递归选择 output 中修改时间最新且文件名不以 `~$` 开头的文件、9-3 不存在时创建 output 目录后打开，三项均通过。空 output 的 9-2 返回空结果且不调用启动器。报告为 `probe/reports/quick_tools_parity_20261023.json`。

此验证没有启动真实桌面应用，也没有在 Windows 上执行旧版 `os.startfile`；它只证明 Rust 在 WSL/Linux 下选择并传递的路径和目录行为。Windows Explorer、macOS、真实用户目录、麒麟/UOS 桌面集成仍未验证，不构成 Python/Rust 对拍或快捷工具整体验收。

### 2026-10-23 批处理 7-1/7-3/7-7/7-8 计划语义合成对拍

新增 `probe/compare_batch_remaining.py`，在临时 XLSX 上调用只读旧 PyCode 与 Rust 计划 CLI。7-1 村镇银行拆分、7-3 外资行删除、7-7 市级涉农贷款差额、7-8 按地区涉农贷款差额的修改后工作簿值投影一致；汇总数字按 `1e-15` 容差比较，排序按 Python 工作表顺序核对。机器报告为 `probe/reports/batch_remaining_plan_parity_20261023.json`。

对拍修正了两处 Rust 计划语义：7-7/7-8 未匹配机构的旧版 `openpyxl cell(..., value=None)` 不会清除已有 D 列值，Rust 计划现保留该值；7-8 地区结果按当前工作簿工作表顺序输出，避免无序 `HashMap` 改变报告顺序。7-1 旧版 `branch_sheets` 统计所有名称含“分机构”的工作表（本例为 6），而 Rust `matched_sheets` 统计可生成修改计划的工作表（本例为 3）；计划的 3 条警告列出其余跳过原因，两个输出工作簿的值投影一致。

计划由 probe 使用 openpyxl 只应用本组涉及的删行、插行和写值操作到独立临时副本；没有执行 `ExcelModifier`，没有创建生产 capability，也没有修改生产默认路由。真实模板、样式/批注/页面设置等保留项、实际文件提交与回退、GUI/CLI 完整日志契约、XLSM、Windows 原生构建、麒麟/UOS 和 deb 未验证；因此四项仍是计划语义候选，不是整体业务验收完成。

### 2026-10-23 PT-7-7/PT-7-8 精度与地区错误回归

静态复核发现旧探针只覆盖整数差额与所有地区均可处理的季度样本，遗漏了两个分支：Rust 原先将 7-7/7-8 写入差额固定格式化为两位小数；7-8 遇到缺少上期工作表或重复机构名时只记录 warning 并跳过该地区，可能使其他地区通过总体一致判断。另有缺少地区汇总行时没有与 Python 一样的逐地区 `error` 字段。

`batch_ops` 现对 7-7/7-8 差额使用可往返完整浮点表示写入，不再截成两位；季度摘要保留缺上期表、重复机构及缺少汇总行的地区结果和对应错误，并将其计入总体不一致判断。探针跨实现数值比较按 `1e-15` 相对/绝对容差，继续比较完整单元格投影及统计结构。新机器报告 `probe/reports/batch_remaining_plan_parity_20261023_r2.json` 对精度差额、缺少上期地区、重复机构和缺少地区统计行逐项对拍通过。

Rust 回归测试 `tests/rural_loan_regressions.rs` 2/2 通过；`cargo fmt --all -- --check`、`cargo test --workspace --offline --locked -q`、`cargo check -p pytools-core --target x86_64-pc-windows-gnu --offline --locked -q` 通过（Windows GNU 项为 WSL 交叉编译检查，不是 Windows 原生运行）。全量测试仍有既有 ignored openpyxl runtime 用例、umya 弃用 API 和 unused import 警告。所有新样本均为合成 XLSX，Python 仅用于只读基线；未执行 `ExcelModifier` 修改器，也未创建 production capability 或更改路由。缺少真实业务样本、完整文件保真、Windows 原生运行、麒麟/UOS glibc 2.28 与 deb 验收，因此 PT-7-7/PT-7-8 仍为**未验证**。

## 2026-10-23 用户指定批量范围：代码入口与集中验证

按用户本轮选定范围，Rust CLI 已为以下功能提供入口：时序 1-1/1-3/1-4/1-5；全部六项去重；打印 3-1 至 3-4；转换与文件操作 4-1 至 4-10；配置 5-1/5-2；汇总 6-1/6-2；批处理 7-1 至 7-8；归档全部功能；快捷工具全部功能。明确不含时序 1-2、打印 3-5 PDF 和 4-11 Python 模式切换。命令入口见 `src/bin/pytools.rs`，业务模块集中在 `src/` 下的 `timeline_ops.rs`、`dedup_ops.rs`、`print_ops.rs`、`conversion.rs`、`config_admin.rs`、`summary_ops.rs`、`batch_ops.rs`、`archive_ops.rs`、`quick.rs` 及文件操作模块。

本轮集中执行的验证：`cargo fmt --all` 通过，`cargo check -p pytools-core --message-format short` 通过，`cargo test --workspace` 全部通过；只读源项目 `PyCode` 测试集 78/78 通过。Rust CLI 另通过读取源配置、初始化并重读临时配置、以及临时文件归档预览/复制冒烟。机器摘要见 `probe/reports/pytools_batch_scope_20261023.json`，人工记录见 `GPT修改记录/20261023_120310_pytools用户指定范围Rust入口与集中验证.txt`。

**以上是代码入口存在和当前自动化测试通过的记录，不将所选菜单组标记为完整迁移或业务验收通过。** 集中测试没有为下列整组功能提供真实业务文件端到端对拍：时序完整流程、打印 3-1 至 3-3、配置初始化/修复、快捷工具、Excel/Word 格式转换，以及 7-1/7-3/7-7/7-8。去重六项和打印 3-4 已补合成 Python/Rust 语义/计数对拍，但真实模板与目标平台未验收；去重标记/删除和打印输出仍受修改引擎能力门控。归档五类工作流已有独立合成 Python/Rust 对拍（见本清单“归档工具”小节），但真实配置、真实目录和目标平台仍未验收。其他已通过的合成对拍也只覆盖各专项报告明确写出的操作，不能外推到同组其他功能。

### 2026-10-23 时序 1-1 表头路径比对合成对拍

新增 `probe/compare_timeline_compare.py`，在合成 XLSX 上调用旧 `compare_393.py` 与 Rust `timeline-compare`。覆盖无差异、位置与路径差异、工作簿关键字未命中、模板作为源文件时跳过。首轮发现 Rust 把汇总表两个计数列写成文本，已改为数值单元格；复跑后两张表的名称、全部值与 Python 类型、差异顺序及汇总计数均一致。机器报告为 `probe/reports/timeline_compare_parity_20261023.json`。

该结果仅覆盖合成 XLSX 的 1-1 数据输出契约；未比较格式保真，也未验证真实配置/业务簿、其他格式或目标平台。故 1-1 仍未完成真实业务验收。

### 2026-10-23 时序 1-5 预校验合成对拍

补齐 Rust `run_timeline_precheck` 对旧版配置表列契约及启用规则/路径映射的检查。新增 `probe/compare_timeline_precheck.py` 对有效配置、详细错误/警告配置和缺必需列配置进行 Python/Rust 计数对拍，三组错误与警告数分别完全一致（0/0、11/3、1/0）。机器报告为 `probe/reports/timeline_precheck_parity_20261023.json`。

对拍只验证合成配置和错误/警告计数，不声称逐字复现旧版终端日志，也未用真实业务配置或目标平台验收。

随后将真实配置 `PyCode/config/config.xlsx` 复制到临时目录，只改副本中的输入/输出/日志目录值并分别运行旧预校验和 Rust CLI；两边均返回 0 错误/0 警告。源文件 SHA-256 和边界见 `probe/reports/timeline_precheck_real_config_20261023.json`。此项只验配置，不代表时序工作簿业务计算已对拍；未发现可用真实业务工作簿。

### 2026-10-23 宽表汇总（1-3/1-4）候选对拍（历史记录）

`probe/compare_wide_summary.py` 以合成 XLSX 分别调用只读旧实现 `wide_397_398.py` 与 Rust CLI `wide-summary` / `wide-summary-row-suffix`。两种行头模式下，新建汇总簿和新目标追加候选的工作表名、维度、单元格值及数据日期格式均通过；数字按 1e-15 相对/绝对容差比较。新建目标报告为 `probe/reports/wide_summary_parity_20261023.json`。

本轮补齐既有目标路径：夹具在 `Append` 中放入一条与新批次重复的数据、一条历史数据，并在前后各放一个应保留的工作表。两种行头模式都按旧版 `append_to_target(..., allow_extend=True)` 语义重建目标 sheet：保留旧数据、追加非重复新行、过滤重复行，并保留其他工作表。Rust 计划经临时测试专用 Passed capability 调用 `ExcelModifier → UmyaSpreadsheet` 产生候选；未登记生产 capability。Python/Rust 的工作表顺序、值/类型投影和日期格式比较通过；夹具不含公式，公式文本保留未验证。替换目标 sheet 的样式、批注、合并、行列尺寸标记按旧版 `if_sheet_exists="replace"` 重建行为丢弃；未触及的 `Before`/`After` sheet 检查值、A1 填充/批注、合并、行高和列宽。源工作簿、配置及原目标文件字节均未改变。机器报告 `probe/reports/wide_summary_existing_target_parity_20261023.json` 保存两种模式的逐格计数与输出文件 SHA-256（仅作为产物标识，不用于跨引擎逐格判等）；临时测试产物和 capability 已清理。

该轮 Rust 数据流：`wide-summary` / `wide-summary-row-suffix` → `run_wide_summary`（共享 `excel-reader`/Calamine 读取源簿，Rust 规则和宽表转换）→ `excel-writer`/rust_xlsxwriter 新建汇总簿 → `build_wide_target_append_plans` 构造目标写入载荷 → `dedup-append-execute` 经能力门控、`ExcelModifier → UmyaSpreadsheet` 写出候选。当时只验证候选、未修改原目标；该限制已由下方显式提交实现补齐。生产 capability 未验收前，禁止把临时测试 capability 推入生产目录或更改路由。

Windows Python 3.12.10 的临时对拍环境使用 NumPy 2.1.3、pandas 2.2.3、openpyxl 3.1.5；Windows 未安装 Rust 工具链，因此 Rust CLI 与 `cargo fmt --all -- --check`、`cargo test --workspace --offline --locked -q` 在 WSL2 Ubuntu x86_64 执行。格式检查和 workspace 测试通过；测试输出保留既有弃用/unused-import 警告，两个要求隔离 openpyxl runtime 的用例仍为 ignored。

以上为当时的合成 XLSX 候选执行证据；既有目标提交补齐情况见下一条。生产 capability、真实业务样本、全量目标 sheet 格式及 OOXML 保留项、Windows 原生 Rust、麒麟/UOS glibc 2.28 和 deb 均未验收，因此 PT-1-3/1-4 仍为**未验证**。

### 2026-10-23 PT-1-3/PT-1-4 显式目标提交与执行摘要

为补齐旧 Python 直接更新配置目标、Rust 仅返回候选的差距，`AppendTargetPlan` 新增既有目标 SHA-256：计划构建前读取哈希，完成目标读取与操作计划后再次核验；目标在构建期间变化或原本不存在的目标出现时拒绝生成计划。`dedup-append-execute` 默认仍只生成候选；只有显式传 `--commit-target` 才执行目标提交。既有目标提交前要求计划含 SHA-256、当前目标仍与计划一致、ExcelModifier capability 通过、候选/替换 sheet 合同通过；随后把候选复制至目标同目录、核对 SHA-256、备份原目标、原子替换并用共享 Calamine reader 重开验证。新目标发布使用不覆盖操作。提交报告给出实际 operation、前后哈希、耗时、备份/清理诊断；执行摘要记录实际 Excel 引擎、阶段状态/耗时、任务成功/跳过数量及追加行计数。提交失败、目标已变化、无效候选或缺少能力证据时不得覆盖原目标。该提交是命令显式选择，尚未接入生产菜单或设为默认。

集成测试覆盖既有目标通过临时测试 capability→`ExcelModifier → UmyaSpreadsheet` 后安全替换、新目标发布、计划后目标变化拒绝、无效候选拒绝，以及默认候选模式不触碰原目标。Windows Python 3.12.10 + openpyxl 3.1.5 对拍脚本对行头加/不加后缀的四种场景（新/既有目标）重跑通过；两个既有目标的目标文件字节与验证候选一致，工作表顺序、值/类型投影和日期格式通过，未触及工作表抽查属性保留，目标 sheet 按 replace 语义丢弃的属性符合 Python 基线。源工作簿和配置字节未变，公式文本因夹具不含公式仍未验证。机器报告 `probe/reports/wide_summary_existing_target_parity_20261023.json` 现含 target commit 和 execution summary 信息；SHA-256 仅用于标识和提交完整性，不作为跨引擎逐格值比较。

验证命令：`cargo fmt --all -- --check`；`cargo test -p pytools-core --test file_commit --test dedup_append_execution --offline --locked`（最终 11/11）；`cargo test --workspace --offline --locked -q`（通过）；`cargo check -p pytools-core --target x86_64-pc-windows-gnu --offline --locked`（通过，WSL 中的 Windows GNU 交叉检查，非 Windows 原生运行测试）；`pytools-core/probe/compare_wide_summary.py`（四种合成模式通过）。没有真实脱敏样本，生产 capability 未验收。实时进度推送、历史运行记录、普通用户/开发者切换视图尚未实现；当前执行摘要是 CLI 完成后的结构化 JSON，不能视为日志系统完整验收。全量样式/OOXML 保真、Excel/WPS 打开、Windows 原生执行、麒麟/UOS、glibc 2.28、离线 deb 仍未验证；PT-1-3/1-4 状态保持**未验证**。

### 2026-10-23 汇总工具 6-1/6-2 合成对拍

新增 `probe/compare_summary_tools.py`，分别对比旧 `summary_321_322.py` 与 Rust `summary-used-range`、`summary-comments`。6-1 覆盖多个工作表、格式定义的 used-range 尾部、日期/布尔/数值/文本；12 行一致。6-2 覆盖行/列区域批注、`set(name)` 批注、重复列头和“模板”回退；6 行一致。两个输出的 sheet 名、单元格值/类型、维度、冻结窗格和表头加粗一致。机器报告见 `probe/reports/summary_tools_parity_20261023.json`。

仅合成 XLSX；没有真实批注模板或真实汇总来源。XLSM/XLS/XLSB、目标平台及 deb 未验证。

涉及既有工作簿修改的去重、打印、sheet 改名和若干批处理命令仍通过 `ExcelEditPlan` 生成计划；只有对应格式、对象特征、整批操作和平台均有能力登记时，`apply-edit-plan` 才能执行。此能力门槛未被本轮更改。格式转换只通过显式 sidecar 命令调用；真实 `.et/.xls/.doc` 样本的转换保留性未验证。真实 PyCode 配置目前只有 3-3 的窄范围 smoke 记录（见 `batch_rename_real_config_smoke.json`），本轮没有将源配置或真实业务文件复制进 fixtures，也没有在源目录执行业务修改。

Git 状态：workspace `main` 尚无首次提交，`git ls-files` 为 0；现有 workspace 文件目前整体处于未跟踪状态。本轮没有 stage、commit 或 push，以免把整个未建基线的 workspace 和既有用户内容一起纳入提交。该状态须先由用户确认首个版本纳入范围后再建立可审查的 Git 基线。

### 2026-10-23 工作表提取整表导入候选扩展

为普通无批注版 `feature_1_6_extract_sheets.py::_copy_full_sheet_with_style` 增加 XLSX 全表导入候选：读取端通过共享 `excel-reader` 的 OOXML worksheet extent 识别实际存储单元格和合并范围，不依赖可能过宽或过时的 worksheet dimension；`ImportRange` 计划可要求 A1 起始的完整矩形，并由 umya 候选复制值/公式缓存、样式、合并区域、行列高宽、隐藏状态、outline 以及矩形内空白坐标。超出单元格范围但显式保存的行列尺寸也会复制。outline 修改继续走 namespace-aware OOXML 定点适配器。`print_ops` 的计划构造同步填入新增导入策略字段。

新增合成测试验证缓存公式按普通版 `data_only=True` 语义导入为值、单元格样式和合并区域、单元格矩形之外显式行高/列宽及隐藏状态、合并范围纳入 worksheet extent，以及合并列定义被切分时的 outline 复制。umya 会在创建单元格时物化默认 A 列宽，本轮增加清理，只保留目标原有列维度和源中显式导入的维度。Rust 输出与源工作表属性对拍 0 差异，机器报告为 `probe/reports/sheet_extract_full_import_parity_20261023.json`。旧 Python 辅助函数的生成副本在 openpyxl 读取样式时触发 `IndexError: list index out of range`；因此对拍对照源工作表的 `data_only=True` 属性和由 helper 复制的样式/结构，而非将该损坏副本当作正确样本。`cargo fmt --all`、`cargo check --workspace --message-format short`、`cargo test --workspace` 均通过。只读 Python 源项目测试以 `PYTHONDONTWRITEBYTECODE=1`、`-p no:cacheprovider` 执行，78/78 通过，3 条 pandas `FutureWarning`；它们只验证旧实现，不替代 Rust 语义属性对拍。

该候选只接受 XLSX，且拒绝含未验证 OOXML 部件的输入。整表任务仍被 ExcelModifier 的精确能力证据门控；本轮没有创建 production capability，也没有改变默认路由。未完成项：真实业务模板、批注/外链等保留项验证、追加既有输出簿、XLS/XLSM/XLSB/ET、Windows 原生和麒麟/UOS deb；旧 Python helper 输出的样式索引异常也需要源项目层面的单独修复决定，本 workspace 不修改源项目。机器摘要：`probe/reports/sheet_extract_full_import_20261023.json` 与 `probe/reports/sheet_extract_full_import_parity_20261023.json`；修改记录：`GPT修改记录/20261023_132321_pytools整表提取ImportRange候选扩展.txt` 和后续对拍记录。

## 2026-10-23 双代理复核后的真实样本闸门

首批范围统一使用审计编号：PT-4-1、PT-4-6..PT-4-10、PT-8-1..PT-8-5、PT-7-4、PT-7-5。用户明确当前“还没有样本”，13 项真实样本及目标平台状态保持**未验证**。已新增 `probe/REAL_SAMPLE_PLAN.md`、`tests/fixtures/real-sample-plan.json` 和开发机辅助 `probe/real_sample_harness.ps1`：清单列出逐项材料，harness 只比较独立捕获的基线/Rust JSON 产物，缺样本退出 2，不启动业务或生成能力记录。当前 `probe/reports/PT_real_sample_intake_20261023.json` 只记录缺样本，不能替代真实端到端对拍；隔离运行捕获适配器须取得样本目录与配置后接入。本闸门通过前不进入生产 capability 登记或 Audit 后续主链。

用户已确认三项目标合同，证据记录在 `probe/reports/PT_semantic_decisions_20261023.json`：已有簿写入分别保留时序/宽表重建目标 sheet、去重追加新增行、提取新增 sheet（PT-1-2 仍排除）；删/插行按 Excel/COM 体验使合并、批注及公式引用随行调整；PT-3-3 只允许输出新文件，已有目标拒绝。PT-3-3 的 CLI 现按该合同输出至配置目标路径并拒绝已有路径，模式 1/2/3 有合成执行对拍；对应记录和限制见本清单 PT-3-3 小节。生产能力/默认路由未更改。PT-2-2/PT-2-4/PT-7-1/PT-7-3 的随行合同和保留项仍须按既定证据要求验收。

缺样本清单最终复核报告为 `probe/reports/PT_real_sample_intake_20261023_r2.json`，包含清单及 harness 哈希；初轮报告原样保留。双代理复核曾指出 PT-7-7/PT-7-8 差额提前舍入、PT-7-8 季度异常地区未进入总体一致性判断的静态风险；这些分支现已由本清单前述 `batch_remaining_plan_parity_20261023_r2.json` 合成回归覆盖，但仍缺真实样本与生产执行/目标平台验收，因此不据此提升功能状态。

本轮另按授权修正文档中的 sidecar/页眉 CLI 滞后描述，并将 PT-4-2 native 输出和帮助明确标注“按值重建（C01，丢弃样式/公式缓存）”；Excel 公式文本仍保留且不重算，未更改转换逻辑、依赖、既有默认值或修改能力。静态对拍完整报告见 `../docs/architecture/audit-crosscheck-2026-10-23.md`。

### 2026-10-23 PT-2-1 批注重复标记执行链验证

在生成的合成 XLSX 上，通过 `pytools dedup-comment-plan --action check` 先验证无能力证据时仅输出计划且源文件字节不变，再为本次测试计划临时生成精确匹配的 Passed capability，调用 Rust CLI → `ExcelModifier` → `OpenpyxlWorker` 批量执行。执行结果确认路由引擎为 `openpyxl_worker`，6 个计划填充单元格均读回为红色，源批注文本/位置投影保持一致，源文件前后字节一致。测试专用 capability 与暂存 runtime 随测试清理，不进入生产能力目录，也未改变默认路由。

本次 CLI 执行夹具单独覆盖 2 个重复行、6 个标记单元格；不要与 `probe/reports/dedup_parity_20261023.json` 的 PT-2-1 语义基线（3 个重复行、8 个标记单元格）混为同一输入。后者确认 Rust 计划填充坐标与 Python 基线红色单元格坐标一致。两项证据均为合成数据；无真实 pytools 模板、真实配置、麒麟/UOS、glibc 2.28 或 deb 验收，因此 PT-2-1 仍为**未验证**（生产验收状态），不得据此创建生产 capability 或宣称迁移完成。机器执行报告为 `probe/reports/dedup_comment_openpyxl_execution_20261023_r1.json`。

### 2026-10-23 PT-2-2 批注重复删除 Excel COM 对拍

新增 Windows/开发机探针 `probe/compare_dedup_row_delete_excel_com.ps1`，用 Microsoft Excel COM 16.0 创建临时 XLSX，并以 Excel 原生删除第 3 行的结果作为本轮用户确认的“相关对象与引用随行调整”基准。探针再调用 Rust `pytools dedup-comment-plan --action delete`，以临时测试 capability 经 `ExcelModifier → UmyaSpreadsheet 3.1.1` 生成候选，对比 sheet 顺序、A1:L6 值/公式/数值格式/填充/粗体、批注坐标与文本、合并区、工作簿定义名称、行高/列宽和筛选范围；全部一致，源文件 SHA-256 前后不变。精简复现的 Rust 集成测试 `tests/dedup_comment_delete_umya.rs` 也通过 CLI 执行并检查计划、行数据、公式引用、批注移动/删除、合并区域、样式及行高。

该执行采用测试目录中的临时 Passed capability，不是生产能力登记。COM 仅作为开发期语义对照，没有进入生产路径。此合同故意遵循用户选定的 Excel/COM 随行行为，与旧普通 Python `openpyxl.delete_rows` 对合并区域/公式引用不修正的行为不同；因此本项是用户确认的目标契约对拍，不应描述为与旧 Python 原样一致。输入是探针生成的合成簿，不是真实业务模板；PT-2-2 生产验收仍为**未验证**。机器报告为 `probe/reports/dedup_comment_delete_excel_com_20261023.json`。

### 2026-10-23 PT-7-3 按 Excel/COM 语义删除行并更新相关引用

`crates/excel-editor/src/umya_backend.rs` 的 Umya 候选现按批次预检并更新静态 A1 公式引用与直接定义名称；公式中绝对行 `$` 锁定保留，同时按结构删除更新行号。预检按公式工作表和删除操作批处理，遇到动态 `INDIRECT`、复杂定义名称、3-D/外部/结构化/整行引用、数组/共享/数据表公式等当前不能安全调整的结构时，在写临时候选前拒绝。跨表公式保留在原工作表坐标，不把目标表删行误施加到其他表本地引用。输出仍经 ExcelModifier 的临时文件与验证流程；测试只使用临时 capability。

`tests/dedup_comment_delete_umya.rs` 覆盖 PT-2-2/PT-2-4 删除后的行值、合并区、批注、样式、行高、目标表本地公式、绝对公式、绝对定义名称、跨表公式及非目标表本地公式；源文件字节保持不变。`tests/foreign_bank_batch_execution.rs` 覆盖 PT-7-3 多 sheet 删除后的静态跨表公式/名称范围更新，并验证动态 `INDIRECT` 和复杂 `OFFSET` 定义名称写入前拒绝。

Python/openpyxl 合成端到端对拍报告 `probe/reports/PT-7-3_execution_parity_20261023_r2.json`：Windows 11 x64 / Python 3.12.10 / openpyxl 3.1.5；2 个输入、4 张分支表、3 行删除、失败数 0，业务计数、sheet 顺序和值投影一致，原输入未改且无变更输出为字节复制。该报告不验证 Excel 对象保真。Excel COM 16.0 对拍报告 `probe/reports/PT-7-3_excel_com_parity_20261023_r8.json` 另行验证公式文本、绝对定义名称、sheet 顺序、目标/非目标值、批注、合并区、样式和行高，语义投影一致且源文件 SHA-256 不变；COM 仅作开发期 oracle，未进入生产路径。

本轮命令：`cargo fmt --all -- --check`、`cargo test -p pytools-core --test dedup_comment_delete_umya --offline --locked -q`、`cargo test -p pytools-core --test foreign_bank_batch_execution --offline --locked -q`、`cargo test --workspace --offline --locked -q`（均在 Ubuntu/WSL 2 执行；Windows 原生 Rust 工具链不可用）、`python pytools-core/probe/compare_pt73_execution.py --report-path pytools-core/probe/reports/PT-7-3_execution_parity_20261023_r2.json` 和 `probe/compare_pt73_row_delete_excel_com.ps1` 指向 r8 报告。以上通过；workspace 全测仍有预先标为 ignored 的 openpyxl runtime 测试，本轮未运行隔离 runtime 专项。

以上均为合成样本证据，PT-7-3 生产验收状态仍为**未验证**：无真实 pytools 工作簿/配置，Windows 原生 Rust、麒麟/UOS、glibc 2.28 和离线 deb 未验收；运行日志/实时进度契约及专门性能基准也未在本功能核验。未创建生产 capability、未改变默认路由。Excel COM 结果遵循用户确认的目标语义，不应表述为普通 openpyxl 删行的原样行为。

### 2026-10-23 PT-4-4 批量工作表改名 CLI 候选

新增 `pytools rename-sheets` 批量入口。Rust 读取配置与 XLSX 工作表名，生成 `ExcelEditPlan` 的 `RenameSheet` 批次；执行仍需调用者显式提供精确匹配的 ExcelModifier capability。候选由 Umya 处理，先写入同目录临时文件并验证，再按输入 SHA-256 安全替换；报告记录实际引擎以及规划、编辑、提交和总耗时。支持文件/目录及递归扫描，忽略 Office 锁文件；当前只处理 `.xlsx`，未验收格式会明确跳过。

旧只读 PyCode 的 `run_batch_rename_sheet_37` 与 Rust CLI 在一份合成工作簿上对拍：1 个文件、5 次改名、5 次跳过；工作表顺序、单元格值、公式文本、自定义名称、样式、批注、合并区域、行列尺寸和页面方向一致。打印区域按引用语义比较，忽略简单工作表名是否带引号这一 OOXML 词法差异。Umya 改名会自动重写其附着的工作表引用名称；为匹配旧 openpyxl 行为，候选保留自定义名称原引用，仅更新打印区域。Rust 机器报告为 `probe/reports/PT-4-4_rust_candidate_20261023.json`，Python 基线为 `probe/reports/PT-4-4_python_baseline_20261023.json`。

验证：`cargo fmt --all -- --check`；`cargo test -p excel-editor -p pytools-core --offline --locked`（Ubuntu/WSL2 x86_64，通过；openpyxl 隔离 runtime 专项按现有配置 ignored）。合成测试还验证无 Passed capability 时源文件不变，以及递归批处理对未验收 `.xls` 的跳过行为。测试使用临时 Passed capability，不构成生产能力登记或默认路由变更。现有 JSON 结果在任务结束时提供阶段耗时和实际引擎，但实时进度、普通用户/开发者双视图及统一历史运行日志尚未接入；后续统一日志实现时补齐，不阻断其他功能迁移。没有真实 pytools 样本，`.xlsm/.xls`、Windows 原生 Rust、麒麟/UOS、glibc 2.28、离线 deb 和专门性能基准仍未验收；PT-4-4 生产状态为**未验证**。

## 顺序与验收门槛

1. 按菜单与 dispatch 完成源代码调用链盘点，不把菜单数量当作功能完成数量。
2. 先完成共享读取、新建写入、原位修改对拍，再逐功能迁移。
3. Rust probe 是首选比较器；只有它无法读取或断言某项属性时，测试辅助才使用 Python/openpyxl。
4. 报告逐项标注保留、丢失、损坏、不适用或未验证；合成和真实模板结果分别记录。
5. 每次迁移状态更新在 pytools-core/probe/reports/ 生成机器报告，并在本目录记录人工可读报告。

### 2026-10-24 PT-6-1 错误单元格保留与 CLI 属性对拍

扩展 probe/compare_summary_tools.py，通过真实 pytools summary-used-range CLI 对照只读 PyCode 基线；合成夹具覆盖两个工作簿、多个工作表、格式定义 used-range 尾部、合并范围、日期/布尔/数值/文本/空白、有缓存及缺缓存公式和 #DIV/0!。13 行的值/类型、已填充单元格样式、工作表维度、冻结窗格、表头、公式/合并/定义名称/打印区和页面设置一致。机器报告 probe/reports/summary_used_range_parity_20261024.json 记录实际 calamine 读取、rust_xlsxwriter 新建、阶段耗时及 1 个 OOXML 错误类型补丁；write_elapsed 覆盖工作簿构建、逐格写入、序列化和补丁。

Rust 在新建报表时保留受支持的 Excel 错误类型；使用 namespace-aware XML 解析器定位报表工作表关系与单元格，写入后由 Calamine 重新读取校验。未知错误值会失败关闭。同步对齐 PyCode 的日期格式及页面边距。NoComPyTools Python 对照与 Rust 在 汇总.rows/styles/types/values 有差异；NoComPyTools 支持的 .xls/.xlsb 与 PyCode 的 .xlsx/.xlsm 范围差异仍待确认，不将候选差异或格式能力登记为生产范围。

验证：cargo fmt -p pytools-core -- --check 通过；patch_error_cells 输出所有权回归单测和 summary_used_range_error CLI 集成测试各 1/1 通过；python pytools-core/probe/compare_summary_tools.py 合成 CLI 对拍通过；git diff --check 通过。cargo test -p pytools-core --lib --offline --locked -q 的 26 个库测试中 25 个通过，workbook_extraction_matches_original_python_helper 因当前 checkout 缺少且 Git 未跟踪 tests/fixtures/季度报表_县区报表_2024-12-31.xlsx 失败；本轮未创建替代夹具。Cargo.lock 未变。

真实业务样本、XLSM/XLS/XLSB、Windows 原生运行、麒麟/UOS、glibc 2.28、deb 尚未验证；日志仅以 CLI JSON 记录该流程引擎和耗时，普通用户/开发者双视图运行日志尚未实现。因此 PT-6-1 仍未完成验收。
