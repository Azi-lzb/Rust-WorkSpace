# py_statistics_tools 功能迁移清单

## 基线与状态口径

- 只读源项目：C:\Users\AZI\Desktop\VibeCode\4.pytools-main
- 主入口基线：PyCode/main.py 的用户菜单和 dispatch 映射；继续追踪 PyCode/core 中的规则、配置、文件处理与工作簿调用链。
- NoComPyTools 是同一产品目标的变体参考；独有行为先登记为待确认，不静默丢弃，也不生成第三个产品。PT-6-1 已对原项目 xlsx/ 下全部 70 份 XLSX 完成 PyCode/NoCom 对拍；差异仍按待确认处理，不并入 PyCode 生产基线，详见本清单末尾记录。
- 状态统一使用保留、丢失、损坏、不适用、未验证；源代码入口和真实业务样本尚未对拍的功能一律不得标为完成。
- 工作表提取以不带批注普通版为准；COM 批注版不迁移。Python 运行模式切换和 PDF 导出不迁移。

## 菜单功能清单

| 菜单组 | 来源入口中的功能 | 迁移边界 | 当前状态 |
|---|---|---|---|
| 时序 | 表头路径比对；时序提取；宽表汇总（行头加后缀）；宽表汇总（行头不加后缀）；预校验 | 迁移时序/路径映射规则、输出列和异常行为 | PT-1-1 现通过原交互菜单 1→1 路由接入 user/developer/off 三视图及开发者运行历史；合成 Python/Rust 输出和样式对拍无差异，真实样本未验收，整体仍未验证。PT-1-3/1-4 两种行头模式已通过 Windows 原生 Rust CLI 合成 Python/Rust 对拍及显式安全目标提交；汇总输出、新目标追加和临时测试 capability 下的既有目标替换有效样式逐格一致，源与配置未变。新增无可解析日期且含缓存公式的菜单对拍：缓存值按 PyCode 行为投影，空日期列格式一致。真实样本、生产 capability 和目标平台仍未验收，整体状态未验证 |
| 去重 | 按批注检查；按批注删除；按配置检查；按配置删除；按配置追加到目标；按配置预校验 | 批注读取使用通用 Excel 原语；保留检查与执行的区别 | 六项均有合成语义对拍；PT-2-1 有合成 CLI→ExcelModifier→openpyxl 执行验证，PT-2-2 有合成 Excel COM 对拍及 CLI→ExcelModifier→umya 执行验证，PT-2-4 有合成 CLI→ExcelModifier→umya 执行验证；PT-2-5 新目标不覆盖发布及既有目标 `--commit-target` 提交通过合成测试专用 capability 验证；生产能力、真实样本和其余编辑能力仍未验收 |
| 打印 | 按批注保留源格式；按批注快速复制；按配置准备打印文件；按配置预校验 | 只写页面设置并准备 xlsx；PDF 菜单不迁移 | 3-1/3-2 计划语义及候选执行已有合成对拍，现已接入菜单 capability/platform 显式确认执行；3-4 预校验通过合成对拍；3-3 模式 1/2/3 已通过合成端到端候选对拍，并按用户合同拒绝已有目标；生产能力、真实模板与目标平台仍未验证 |
| 转换与文件命名 | 批量重命名；Excel 格式转换；Word 格式转换；批量修改 Sheet 名；按内容加前缀重命名；取消内容前缀；局部映射重命名；提交缺失检查；文件排序前缀；取消排序前缀 | 保留文件操作与同类格式转换；首批 sidecar 为 .et→.xlsx、.xls→.xlsx、.doc→.docx；不迁移 Python 模式切换 | PT-4-4 已接通菜单计划与 capability/platform 显式传入后的确认执行；仍使用合成 Passed capability 测试，不登记生产能力。其他转换项、真实工作簿及整组目标平台验收仍未验证 |
| 配置 | 初始化/修复配置表；查看配置并从磁盘重新读取 | 盘点配置模板、默认值、修复行为、字段映射与错误提示 | 5-1 新建、5-2 查看通过合成及真实只读对拍；5-1 修复已在合成及真实配置临时副本上通过 openpyxl worker 候选执行与语义对拍；生产 capability、完整文件保真与目标平台仍未验收 |
| 汇总 | 按使用区域汇总；按批注汇总 | 逐项确认坐标、模板/源文件关系、行顺序和输出格式 | PT-6-1 已接入真实 summary-used-range CLI，并接入增量交互菜单（无参数或 pytools menu → 主菜单 6 → 汇总 1；从全局配置读取输出/日志目录；Win32 多选对话框及手工分号路径回退均已通过真实菜单实测；2026-10-25 原生对话框双文件业务样本 1,167 行的完整输出投影与 PyCode 精确一致，详见 `probe/reports/summary_used_range_real_menu_parity_20261025.json`）；新菜单集成测试 3/3（含 PT-6-1 多文件、PT-6-2 模板/源文件与损坏输入失败记录）；原 PyCode 只读真实 XLSX 经 menu→6→1 路由对拍 7,462 行完整工作簿投影精确一致，14 条开发者事件与运行历史一致，源未改变；合成 XLSX/XLSM 对拍（2 个工作簿、19 行）值/类型、样式、维度、冻结窗格、公式/合并/名称/页面属性通过；空、隐藏和 veryHidden 工作表及 D5:F6 偏移区域按 PyCode 行为汇总；7 种常见 Excel 错误值保留为 error 类型。新增 15 个字面 t=e token 边界夹具；7 个常见错误类型保持为 Excel error，8 个扩展 token 与 PyCode 一致输出为文本。原项目只读 `xlsx/` 目录中的 13 份 XLSX 分 4 批经 PyCode/Rust CLI 对拍，涵盖 96 个工作表和 203,580 行；值/类型、样式、工作表与输出属性完全一致，输入未改变；最大批次 188,365 行。真实宽报表中的合并子格 B1 曾被 Calamine 投影为值，Rust 现依据共享 reader 合并范围忽略非锚点值，与 openpyxl 行为一致。另有 15 个工作表缺少 OOXML `<dimension>` 元素的真实簿，实际范围达第 70 行/第 257 列，PyCode 的 427 行结果与 Rust 批次对拍一致。宽报表 1,212 个受支持错误值经既有 OOXML Patch 保留。未知 `#DATA!` 单元格由 PyCode 写成字符串；PT-6-1 显式使用共享 reader 的 opt-in 错误 token 文本化 API，再由 Calamine 读取；默认 reader 仍失败关闭。该合成 XLSX 边界夹具已在 Windows GNU debug 与 WSL glibc 2.28 release 的 user/developer/off 三视图中与 PyCode 完整输出对拍一致。最大批次 Rust 阶段耗时记录于真实报告；Windows/WSL 结果不作性能对比，峰值内存未测量。PT-6-2 常规模板已有 6 行合成对拍；重复同编号起止标记按最后匹配项生效的 PyCode/Rust CLI 对拍通过。NoComPyTools 独有格式、其他输入来源、原项目业务含 VBA 的 XLSM 和目标平台仍未验收；官方 VBA 测试样本已完成 PyCode/Rust 对拍，Rust CLI 现提供普通用户、开发者及 off 三种日志视图，阶段状态、实际引擎与耗时写入 JSON 运行记录；合成 CLI 测试与 13 份真实 XLSX 的四批对拍验证开发者事件流和历史 JSON 一致；PT-6-1 的旧 PyCode 输入合同只支持 .xlsx/.xlsm；混合有效 .xlsx/.xlsm 与无效 .xls/.xlsb 的 Python/Rust CLI 对拍输出一致，.xls/.xlsb 过滤事件未调用 reader；仍待原项目业务含 VBA 的 XLSM 和目标平台验收；最新 pytools-dev HEAD `98a702f` 已在 WSL 用 Zig 0.17.0 重建 `x86_64-unknown-linux-gnu.2.28` release，`objdump -T` 最高 GLIBC 符号为 `GLIBC_2.28`；该 release 对 PyCode 13 份真实 XLSX、96 张工作表、203,580 行四批完整投影对拍一致（详见 2026-10-24 报告） |
| 一二批处理 | 拆分村镇银行数据；机构名称标准化；删除分机构表外资行；外汇页眉修改；地区总分校验；工作表数据提取；涉农贷款比上月修正；季度涉农贷款比上月修正及分地区核对 | 仅工作表提取普通无批注版；对 COM 路由的业务行为与普通实现差异做盘点 | 7-1/7-3/7-7/7-8 计划语义合成对拍通过；7-7/7-8 已接入真实交互菜单、三种日志视图及显式 capability 安全提交，计划/摘要和菜单候选执行通过合成对拍；7-2/7-5/7-6 有各自候选证据；7-4 直接 CLI、交互菜单与合成安全提交对拍通过；既有簿真实样本和整组验收仍未完成 |
| 归档 | 按日期预览；按类型预览；执行归档计划；按日期直接复制；按类型直接复制 | 保留预览、冲突检查、复制结果和归档规则 | 五项 Rust CLI 已接入真实 `pytools menu` → 主菜单 8 → 1–5；选择支持多选文件或递归文件夹，目标目录、输出目录和日志目录走交互/全局配置。Windows 原生菜单集成 5 条路由、三种日志视图、预览只生成计划、计划执行与直接复制字节均通过；`archive_parity` 继续覆盖五项工作流对只读 Python 基线的合成对拍。运行历史记录处理计数、耗时及 Calamine、rust_xlsxwriter、Rust 文件复制引擎，不写输入路径/文件名。文件弹窗 UI 未直接操作；真实业务目录/配置与目标平台仍未验收，整体状态未验证 |
| 快捷工具 | 打开配置文件；打开 output 最新文件；打开 output 文件夹 | 迁移到目标平台可用的文件打开/定位体验，系统集成细节待确认 | Rust 三个业务 CLI 已有 WSL 假 `xdg-open` 合成集成验证；现已接入 `pytools menu` → 主菜单 9 → 1–3，Windows 原生 CLI 测试覆盖三条菜单路由、配置打开失败/无最新文件/output 目录失败及返回主菜单。Windows Explorer 实际打开、旧版 Windows `os.startfile`、目标 Linux 桌面尚未实测，整体仍未验证 |

PT-6-1 另于 2026-10-25 在 Windows MSVC HEAD `13c83fb` 经真实 `pytools menu → 6 → 1` 手工分号路径回退处理 2 份原项目只读 XLSX（1,102 行），PyCode 与 Rust 完整工作簿投影精确一致；在后续 HEAD `0500b20` 通过同一真实菜单打开 Win32 原生多选对话框处理 2 份只读 XLSX（1,167 行），完整投影同样精确一致，运行历史成功且源文件未改变，实际引擎为 Calamine 与 rust_xlsxwriter。两种路径的机器证据见 `probe/reports/summary_used_range_real_menu_parity_20261025.json`。

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

此候选仅覆盖 XLSX：区域值由 Calamine 读取并以 rust_xlsxwriter 新建输出，区域模式沿用 Python data_only=True 且不复制公式文本；整表模式需通过 ExcelModifier capability。extract-sheets --append-existing 现支持对既有 XLSX 区域或整表结果追加新工作表，但仅在提供匹配 --capabilities、--platform 且目标簿 inventory 完整时执行；默认 CLI/菜单仍拒绝覆盖既有输出。日期、时长、错误值及非有限数目前显式拒绝。其他格式、真实模板及文件弹窗尚未验证。

机构名称规范化已有 `.xlsx` Rust CLI（`pytools institution-normalize --config ... <workbook>...`），默认只生成计划；执行须显式 `--apply` 并提供匹配的 ExcelModifier capability。候选按原规则在“分机构”工作表统一选择 A 列或 B 列，精确映射优先，其次按映射顺序不区分大小写做子串映射；多文件用 Rayon 并行，同一路径重复出现时退回顺序处理。原 `run_normalize_institution` 候选与 Python `_process_sheet` 在一份四工作表合成簿上对拍，Rust probe 的 sheet 名、20 个值、1 条公式、1 个合并区域和 1 条批注均保留；文件统计与 Python 一致（3 个匹配工作表、2 个改动工作表、3 行改动）。openpyxl 元数据检查还验证了该夹具的样式、行列尺寸、批注、合并区域、页面设置和页眉页脚。`build_institution_normalize_plan` 通过共享 `excel-reader` 构造 3 个 `SetCellValue` 操作，并与 Python 基线的统计及变更格一致，计划对拍报告为 `probe/reports/institution_normalize_plan_parity.json`，该单独报告未执行修改引擎。另有显式忽略的集成测试在合成四工作表夹具上临时登记 fixture-scoped capability，实际执行 CLI → ExcelModifier → `openpyxl_worker`，并对拍值、公式、批注、合并、样式、行列尺寸、页面设置、页眉页脚和源文件安全性；执行报告为 `probe/reports/institution_normalize_openpyxl_execution_20261024.json`，状态明确为 `synthetic_passed_not_production_acceptance`。worker 修复了 openpyxl 写出时遗漏的 `xml:space="preserve"`，并由 Rust 读取候选临时文件完成保留项验证。该 capability 仅为测试临时文件，不产生生产 Passed 登记，也不改变默认路由。以上仅为合成 XLSX 证据；无经授权真实样本，`.xlsm`、真实模板、Windows 原生执行、目标 Linux/glibc 2.28 和 deb 尚未验证；PT-7-2 生产状态仍为未验证。旧直接 umya 候选未改为 ExcelModifier 默认路由。

文件排序前缀（4-9/4-10）已有纯逻辑 helper 对拍，报告为 `probe/reports/file_sort_prefix_parity.json`。4-9 `run_file_sort_prefix_49` 使用共享 `excel-reader`/Calamine 读取“文件排序配置”，按输入顺序匹配并改名；目标已存在时自动添加 `_1` 等后缀，不覆盖占用文件；使用共享 `excel-writer`/rust_xlsxwriter 新建“文件排序前缀结果”工作簿。报表写入失败时会撤销本轮已完成的改名。该实现现已接入 `pytools file-sort-prefix --config <配置簿> [--output-dir <目录>] [--log-dir <目录>] <文件>...` Rust CLI。与只读 PyCode 的 `run_file_sort_prefix_49` 在一份合成配置簿和 8 个文件/路径案例上对拍通过，API 与 CLI 均验证统计、报告路径、文件名集合及文件内容；CLI 为成功/失败项写带时间戳的日志并输出到 stderr。日志消息/成功项行数已做合成断言，但未逐字与 Python logger 的时间、并发和重复运行行为对拍。完整机器报告为 `probe/reports/file_sort_prefix_operation_parity.json`。原 GUI 菜单/dispatcher、真实目录和目标平台仍未验收。

4-10 `run_remove_file_sort_prefix_410` 已接入 `pytools remove-file-sort-prefix --config <配置簿> [--output-dir <目录>] [--log-dir <目录>] [<文件>...]` Rust CLI。它读取排序规则和可选的 `重命名配置!O:P:Q` 内容单元格规格；直接以排序前缀开头的文件名不打开工作簿，排序前缀位于内容前缀之后时才用共享 Calamine reader 读取指定内容格，并只在内容前缀吻合时移除排序片段。改名使用不覆盖目标的硬链接预留，报告写入失败会回滚本轮改名；结果通过共享 rust_xlsxwriter 写入“取消文件排序前缀结果”。CLI 写时间戳日志文件和 stderr，允许空文件列表并输出兼容 Python 的空报告。合成 XLSX 对拍覆盖 10 个文件/路径案例：规则 4、改名 5、跳过 5、失败 0；报告行、文件名集合与 Python 基线一致，并检查改名不改变输入文件字节。空输入 API 与 CLI 均验证：保留报告工作表但不写表头或单元格。完整机器报告为 `probe/reports/file_sort_prefix_remove_parity.json`，Python 明细基线为 `probe/reports/file_sort_prefix_remove_python_baseline.json`。原 GUI 菜单、日志时序/并发行为、真实目录和目标平台尚未验收，不代表整个菜单功能已迁移。

4-8 局部文件名映射 `run_local_filename_map_candidate` 已接入 `pytools rename-local-map --config <配置簿> [--log-dir <目录>] [--log-view user|developer|off] <文件>...` Rust CLI 和类别 4 菜单 4→7，并与只读 PyCode 的 `run_batch_rename_local_map_48` 对拍。它从“重命名配置”工作表按 `原文件名片段`/`新文件名片段` 表头读取映射，按配置顺序在文件 stem 中逐项替换，并保留最后一个扩展名；允许任意扩展名，空目标表示删除片段。改名目标冲突时使用 `_1` 等编号且不覆盖已存在文件；目标名不合法、过长或文件系统操作失败时逐项跳过并保留源文件。合成场景 10 个输入路径，含 6 条映射：Python 与 Rust 均改名 5、跳过 5，映射顺序、中文片段、空目标、冲突后缀、缺失路径、目录、未命中和改名错误的最终目录集合一致；Rust 测试检查输入文件字节保持及冲突目标不变。只读 Python `test_rename_local_map_48.py` 和 `test_convert_rename_edges.py` 共 19 项为历史证据，本轮未重跑；Rust 4-8 对拍测试 3/3、真实菜单三视图/类别回归 7/7 通过。user 显示阶段进度和耗时；developer 写结构化运行历史，记录 Calamine、RustFilesystem 和阶段耗时，不写输入路径或文件名；off 不显示进度或写历史。机器报告为 `probe/reports/local_filename_map_48_parity.json`，Python 基线为 `probe/reports/local_filename_map_48_python_baseline.json`。当前仍限合成配置/目录和已对拍行为，未验证真实业务目录、非 UTF-8 文件名跨平台语义、Windows 原生文件选择器、目标 Linux/glibc 2.28 与 deb；日志文本本轮未与 Python logger 逐字对拍。

4-6 内容前缀添加 `run_content_prefix_rename_46_xlsx_candidate` 已接入 `pytools rename-by-content --config <配置簿> <文件>...` Rust CLI，并与只读 PyCode 的 `run_batch_rename_by_content_46` 对拍。它从唯一启用的 `重命名配置!O:P:Q` 读取规格，用共享 Calamine reader 读取 `.xlsx/.xlsm` 指定单元格；沿用缺失工作表的 `sheet!A1` 回退、内容项已出现在文件名中则跳过、前缀最多 180 个 Unicode 字符、目标冲突添加 `_1` 等后缀的行为。候选只改文件名、不保存工作簿；改名通过硬链接预留避免覆盖目标。合成场景 12 个输入路径：改名 5、跳过 7；API 与 CLI 输出统计、最终文件名集合均与 Python 基线一致，Rust CLI 测试逐文件检查输入工作簿字节未变并确认占用目标未覆盖。既有 Python `test_rename_by_content_46.py` 13/13 是前一轮历史证据；本轮未重跑 Python。Rust 4-6 集成测试 2/2、pytools-core 49 项、workspace 121 项通过。机器报告为 `probe/reports/content_rename_46_parity.json`，Python 基线为 `probe/reports/content_rename_46_python_baseline.json`。4-6 现已接入类别 4 真实菜单 4→5；CLI 支持 `--log-dir` 与 `--log-view user|developer|off`，菜单沿用全局日志视图。用户视图报告逐文件进度和阶段耗时；开发者视图记录 Rust、Calamine、RustFilesystem、阶段耗时和任务总耗时，不记录输入路径/文件名；off 不输出进度或写结构化历史。Windows 原生真实菜单三视图测试与 4-6/4-7 回归共 12/12 通过；本轮日志文本未与 Python logger 逐字对拍。该证据仅覆盖合成文本格与 `.xlsx/.xlsm`；源 `.xls/.xlt` COM 分支、真实宏、数字/日期/公式缓存值、真实目录、跨平台文件系统和目标 Linux/deb 仍未验收。

4-7 内容前缀取消 `run_content_prefix_unrename_47_xlsx_candidate` 已接入 `pytools unrename-by-content --config <配置簿> <文件>...` Rust CLI，并与只读 PyCode 的 `run_batch_unrename_by_content_47` 对拍。它读取 `重命名配置!O:P:Q` 中唯一启用的内容规格，使用共享 Calamine reader 读取 `.xlsx/.xlsm` 指定单元格；缺失工作表/空值沿用 `sheet!A1` 作为回退项，移除文件名开头匹配的内容项并保留已配置排序前缀。候选不保存工作簿，只执行逐文件改名；目标已存在时以不覆盖的硬链接预留并尝试 `_1` 后缀，未匹配、读取失败或未验收扩展名逐项跳过。合成主场景 13 个输入为改名 8、跳过 5；无“文件排序配置”工作表的旧配置场景为改名 1、跳过 1；Rust API 与 CLI 的统计和最终文件名集合均与 Python 基线一致，并检查已改名文件字节未变化。旧 Python 聚焦测试 13/13 是前一轮历史证据，本轮未重跑；Rust 4-7 集成测试 4/4、pytools-core 49 项、workspace 121 项通过。机器报告为 `probe/reports/content_unrename_47_parity.json`，基线为 `probe/reports/content_unrename_47_python_baseline.json`。4-7 现已接入类别 4 真实菜单 4→6；CLI 支持 `--log-dir` 与 `--log-view user|developer|off`，菜单沿用全局日志视图。用户视图报告逐文件进度和阶段耗时；开发者视图记录 Rust、Calamine、RustFilesystem、阶段耗时和任务总耗时，不记录输入路径/文件名；off 不输出进度或写结构化历史。Windows 原生真实菜单三视图测试与 4-6/4-7 回归共 12/12 通过；本轮日志文本未与 Python logger 逐字对拍。该证据仅为合成配置和 `.xlsx/.xlsm`；真实 `.xlsm`、源 `.xls/.xlt` COM、日期/公式缓存格、真实目录、不同文件系统、目标 Linux/deb 未验证。

地区总分校验 `feature_1_5_region_sum_check.py::run_region_sum_check` 的非 COM 主路径已由 Rust `region_sum::run_region_sum_check` 实现，并通过 `pytools region-sum-check` 接入 Rust CLI。命令接受显式工作簿列表；可从 `--config` 读取全局输出/日志目录，也可用 `--output-dir`、`--log-dir` 覆盖；不传配置时默认写入 `./output` 和 `./logs`。读取通过共享 `excel-reader`/Calamine，写入通过共享 `excel-writer`/rust_xlsxwriter；COM 分支不迁移到此命令。5 个合成 XLSX 场景的文件/工作表/错误/失败统计与 Python 一致（4/3/1/1），结果单元格、公式文本（按 Python 公式前缀归一）、数据区域尺寸、样式和列宽（绝对差不超过 0.01）通过对拍，日志关键记录存在且输入文件字节不变。Python 基线为 Python 3.12.10、OpenPyXL 3.1.5；源函数与 logger 的 SHA-256 见 `probe/reports/region_sum_python_baseline.json`。为保持汇总标记公式的空缓存，结果生成后使用 `ooxml-patch::clear_empty_formula_cache` 仅清除对应普通公式单元格的空缓存；单元格坐标按结果行数计算，不是通用公式修改能力。新增的目标名冲突与输出目录失败测试确认既有结果/日志及输入不会被覆盖或更改。完整机器报告为 `probe/reports/region_sum_parity.json`；本轮新增 CLI 集成对拍在 `tests/region_sum_parity.rs` 直接启动产品二进制并与同一基线比较统计、输出单元格/公式及输入字节。真实业务模板、非 XLSX 格式、性能/峰值内存、Windows 原生 Rust、麒麟/UOS glibc 2.28 与 deb 均未验收；原 GUI 菜单仍未迁移，Python 原功能保留作回退，因此不能将地区校验菜单标为已迁移。

外汇页眉单位修正 `feature_1_4_fx_header.py` 已有窄范围 Rust 候选 `fx_header::build_fx_header_plan`：业务层读取共享 `excel-reader::read_header_inventory` 的 odd/even/first 页眉文本，只生成 `ReplaceHeaderText` 计划，经 ExcelModifier 路由到 OOXML Patch。PyCode 与 NoComPyTools 对应 Python 源文件当前 SHA-256 相同；普通非 COM 行为以 PyCode 基线为准。Python 基线统计为文件 1、外汇工作表 2、修改工作表 1、失败 0；Rust 候选验证外汇工作表计数 2 和计划修改工作表计数 1 与 Python 相符，并成功执行/校验单文件输出。验证各页眉文本、非目标工作表、值/公式/合并/批注/条件格式/数据验证以及补丁范围；批量文件统计现由下述 PT-7-4 入口聚合。机器报告为 `probe/reports/fx_header_parity.json`，Python 基线为 `probe/reports/fx_header_python_baseline.json`。PT-7-4 候选仅接受 `.xlsx`，现已接入 `pytools fx-header` 直接 CLI 与交互菜单 7→4，默认只生成计划；提交目标必须显式提供匹配的 ExcelModifier capability、platform 和 `--commit-target`。菜单选择、计划默认行为、扩展名跳过、能力门控提交及 user/developer/off 日志视图由 Windows 原生 Rust 集成测试覆盖。合成执行使用临时 fixture-scoped capability，经 ExcelModifier→OOXML Patch 修改测试副本并安全替换目标；回滚备份在提交验证期间使用，成功后清理。Python 基线普通路径统计、页眉文本、值、公式文本、合并区、批注、条件格式和数据验证与 Rust 对拍一致，机器报告为 `probe/reports/fx_header_menu_parity_20261024.json`；测试明细见 `tests/fx_header_menu_cli.rs`。此项仍未通过真实菜单文件选择器、真实业务工作簿、`.xlsm`/VBA、生产 capability、Kylin/UOS、glibc 2.28 和 deb 验收，PT-7-4 整体状态保持未验证。

### 3-3 配置驱动的批量文件重命名

`PyCode.core.convert_33_34_35_37.run_batch_rename_33` 已迁移到语义化的 `pytools_core::batch_rename::run_batch_rename`；Rust 仍提供标记为 deprecated 的 `run_batch_rename_33` 兼容别名。该操作通过 `pytools rename-batch --config <工作簿> [--log-dir <目录>] [--log-view user|developer|off] [<文件或目录>...]` 接入 Rust CLI，也接入真实交互菜单“转换工具→1”（主菜单 4→1）。实现读取 `重命名配置` 的简称/全称、全称/代码映射及数据日期/报表名称；按配置首次插入顺序执行不区分大小写的子串匹配，重复配置键覆盖值但保留原顺序；缺少代码、无命中、原名不变、路径不存在及 OS 改名错误按源函数计入跳过；目标冲突使用 `_1` 等后缀。支持文件和目录，不读取或改写工作簿内容。

Python 基线由源 `run_batch_rename_33` 对同一合成配置及路径场景实跑，Python 3.12.10、pandas 2.2.3、openpyxl 3.1.5：改名 3、跳过 4，最终文件/目录名与 `probe/reports/batch_rename_python_baseline.json` 一致；覆盖匹配顺序、代码缺失、无命中、缺失路径、同名跳过、连续冲突后缀、目录内容和 ASCII 大小写匹配。Rust `tests/batch_rename_parity.rs` 验证 API 与 CLI 结果、日志、输入/目标文件内容及配置字节不变，汇总见 `probe/reports/batch_rename_parity.json`。对原只读 `PyCode/config/config.xlsx` 还做了两项验证：CLI 空输入读取成功（0/0）；随后从真实配置选择一个有效别名，在系统临时目录创建合成文件并由 Rust CLI 改名（1/0），确认原配置字节与临时文件内容都不变。结果摘要见 `probe/reports/batch_rename_real_config_smoke.json`。本轮新增 API 配置读取失败事件断言，以及 Windows 原生交互菜单三视图测试：菜单传入全局 `log-view`；user 显示步骤、计数和耗时，developer 输出事件并写 JSON 历史，包含 Calamine/RustFilesystem、阶段耗时和累计耗时且不记录输入路径/文件名，off 不输出进度或新增结构化历史；旧逐项 `.log` 文件仍保留。`cargo test --offline --locked -p pytools-core --test batch_rename_parity --test conversion_menu_cli -j 1`：3/3 和 5/5 通过。未重新执行 Python 基线脚本；对拍测试读取既有 Python 实跑基线，业务改名结果和配置行为未更改。

当前证据仅覆盖脱敏合成配置与临时路径场景；真实业务输入文件、真实配置对应文件名覆盖率、重名冲突在目标文件系统上的竞争行为、桌面原生文件选择器、麒麟/UOS 和 deb 未验证。交互菜单选择及路径回退接线已有测试，但不代表完整原 Python GUI 行为或真实样本已验收；3-3/PT-4-1 整体状态仍为“未验证”。

上述结果不代表完整菜单或真实业务模板已验收。尚需：
- 将每个菜单命令映射到源函数、配置项、文件格式、输出契约和回归用例。
- 对比 PyCode 与 NoComPyTools，列出共享与独有行为。
- 为尚未覆盖的配置表、业务工作簿及各迁移功能建立 pytools 专属 fixtures、probe config、对拍输出和日期报告。
- 时序抽取目前没有批量目录调度、CSV 输入、结果工作簿写入或向目标簿追加；日期单元格序列值、xlsb 合并区域和真实模板仍未验证。
- 普通版工作表提取已有 Rust CLI 与 7→6 主菜单入口；extract-sheets --append-existing 仅作为显式 CLI 选项追加既有 XLSX 输出，需精确 capability/platform。区域及整表追加已有合成执行与既有簿属性回归；生产 capability 未登记。真实模板、其他格式、Windows 文件弹窗及麒麟/UOS glibc 2.28/deb 未验证。
- 机构名称规范化已有显式文件列表和配置 CLI；仍未接入原 GUI 菜单/dispatcher 或目录发现，当前候选只接受 `.xlsx`。合成夹具已有通过临时测试 capability 的 ExcelModifier/openpyxl worker/validator 执行证据，但未形成生产 capability；计划报告与执行测试覆盖范围不同。源实现支持的 `.xlsm` 保留行为、真实模板中的全部 OOXML 部件、真实样本、Windows 原生执行及目标平台仍需验收。
- 地区总分校验 `region_sum::run_region_sum_check` 现已接入 `pytools region-sum-check` CLI，并通过合成 XLSX 对拍；原 GUI 菜单、真实模板和非 XLSX 格式、执行性能/峰值内存、Windows 原生与目标 Linux/deb 均未验证，现有 Python 功能仍保留。
- 4-6/4-7、4-8、4-9 和 4-10 均已接入显式文件列表 Rust CLI；只有 4-9/4-10 写操作日志。4-6/4-7 已接入 `pytools menu`→4→5/6，并有交互菜单集成测试；本轮原项目随附 13 份 `.xlsx` 样本经 Python/Rust CLI 对拍，详见 `probe/reports/content_prefix_project_asset_parity_20261025.json`。该样本不是授权客户模板，真实宏、数字/日期/公式缓存边界、`.xls/.xlt` COM 读取和不同文件系统仍未验证；Windows 桌面文件弹窗未人工操作。4-9/4-10 日志时间、并发及重复运行行为尚未逐项与 Python logger 对拍；4-8 尚未验证真实配置、目标平台和非 UTF-8 文件名。上述操作的真实目录、Windows 原生 Rust 构建和目标 Linux 包均未验收。4-10 的 `.xls/.xlt` 内容读取、多单元格规格及真实文件仍未验证，源 COM 内容读取仅作基线，不进入 Rust 主路径。旧 COM 业务读取未迁入 sidecar。
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

### 2026-10-24 PT-1-3/PT-1-4 Windows 原生样式对拍

按 PyCode `write_workbook` 的 XLSX 输出行为补齐宽表汇总表头粗体、细边框、居中和顶端对齐；追加载荷增加显式表头样式与“数据日期”列格式，确保空白日期单元格也有 `yyyy/m/d`。普通去重追加载荷保持不设置这些宽表专用样式。`compare_wide_summary.py` 支持通过 `PYTOOLS_RUST_CLI` 直接调用 Windows CLI，并逐格比较有效样式；本轮移除已修复样式差异的过期限制说明。

真实入口：Windows PowerShell 启动 `target/debug/pytools.exe`，覆盖 `wide-summary`、`wide-summary-row-suffix` 与 `dedup-append-execute --commit-target`。两种行头模式的汇总输出、新目标追加、使用临时测试 capability 的既有目标替换均通过 Python/Rust 合成对拍；汇总工作簿和被替换目标 sheet 的有效单元格样式完全一致。目标替换按原 Python replace 行为重建目标 sheet；未触及工作表顺序和值、抽查样式/批注/合并/行高/列宽通过检查，源输入与配置未变化。候选字节与提交后的目标一致。报告：`probe/reports/wide_summary_parity_20261023.json`、`probe/reports/wide_summary_existing_target_parity_20261023.json`。

验证：Windows 原生 `cargo build --offline --locked -p pytools-core --bin pytools` 通过；`cargo fmt --manifest-path pytools-core/Cargo.toml -- --check` 通过；`cargo test --offline --locked -p pytools-core --test dedup_append_execution` 通过（6/6）；使用原项目 Python 3.12.10 venv 执行 `probe/compare_wide_summary.py`，4 种汇总/目标模式通过且样式差异数为 0。完整 `cargo test --offline --locked -p pytools-core` 因仓库缺少既有夹具 `pytools-core/tests/fixtures/季度报表_县区报表_2024-12-31.xlsx` 失败；跳过对应单测后，该套件的另一个既有 `config_task_loader_parity` 集成测试在 Windows `\\?\` canonical path 前缀断言失败，与本次业务代码无关。格式与 diff 检查均通过。

范围仍为合成 XLSX；没有可用脱敏真实业务样本。临时 capability 仅用于测试，没有登记生产路由；PT-1-3/1-4 整体仍为**未验证**。真实工作簿、完整 OOXML 保留、麒麟/UOS、glibc 2.28 和离线 deb 未验收。此轮没有共享 crate、根 Cargo 配置、Cargo.lock、Audit 或长期 Goal 变更。

### 2026-10-25 PT-1-3/PT-1-4 交互日志视图与回归对拍

保留 `run_wide_summary` 原调用接口，增加 `run_wide_summary_with_progress`、`WideSummaryProgress` 和 `WideSummaryRunEvent`。真实交互入口为 `pytools menu --log-view <user|developer|off>` → 时序工具 1 → 宽表汇总 3/4；阶段事件覆盖源工作簿读取、规则汇总、汇总簿写入及目标追加计划，标记实际 Rust 组件并记录数量、阶段/总耗时和实际 Calamine、rust_xlsxwriter 引擎。user 视图打印业务步骤，developer 视图将结构化事件写到 stderr 并写入配置日志目录下的历史 JSON，off 保持原调用且不输出进度或创建历史。运行历史仅记录输入序号和处理统计，不记录源路径或单元格内容；CLI stdout 保持单行 JSON。

Windows 原生验证：`cargo check --manifest-path Cargo.toml -p pytools-core --bin pytools` 通过；`cargo test --manifest-path Cargo.toml -p pytools-core --test timeline_menu_cli` 3/3、`--test dedup_append_execution` 6/6 通过；`cargo fmt --manifest-path pytools-core/Cargo.toml -- --check` 和 `git diff --check` 通过。使用原 PyTools Python 3.12.10 venv 与原只读 PyCode 基线运行 `probe/compare_wide_summary.py`，两种行头模式的新建汇总和既有目标替换共 4 个合成场景通过；两个汇总结果各比较 36 个单元格且样式精确，既有目标替换样式差异为 0，源输入和配置未变化。报告：`probe/reports/wide_summary_parity_20261023.json`、`probe/reports/wide_summary_existing_target_parity_20261023.json`。

该结果仍仅覆盖合成 XLSX 和菜单管道输入，不包含原生文件选择弹窗实测或脱敏真实业务样本；公式保留、生产 capability、完整真实模板保真、麒麟/UOS glibc 2.28 与离线 deb 未验收，PT-1-3/1-4 整体保持**未验证**。

### 2026-10-25 PT-1-3/PT-1-4 交互式目标写入接入与菜单对拍

主菜单现在将 capability/platform 传入时序子菜单。宽表汇总 3/4 生成计划后显示目标及追加行数，并要求用户确认；新目标确认后通过 rust_xlsxwriter 写入、重开验证并不覆盖发布，既有目标没有 capability/platform 时仅保留计划，有匹配能力时沿用 `dedup-append-execute --commit-target` 的快照核验、候选验证和安全替换。目标执行阶段写入 user/developer/off 兼容的结果摘要及开发者历史，历史记录实际引擎和阶段耗时，不记录路径或单元格内容。

Windows 原生 `cargo fmt --manifest-path pytools-core/Cargo.toml -- --check` 通过；`cargo test --offline --locked -p pytools-core --test timeline_menu_cli --test dedup_append_execution` 为 12/12。`probe/compare_wide_summary.py` 已改为通过只读 Python PyCode 基线对比真实 `pytools menu → 1 → 3/4 → 确认` 新目标路径，两种行头模式共 2 个合成菜单场景通过；两个汇总各对比 36 个单元格、有效样式精确，目标输出、候选与已提交簿一致。该脚本另对两种模式的既有目标替换执行 2 个 Rust CLI 对拍，分别比较 467/476 个单元格，目标 sheet 样式精确、未触及工作表属性保留；既有目标的交互菜单计划模式/临时 capability 确认替换由 `timeline_menu_cli` 集成测试覆盖。开发者历史的 `target_append_execute` 状态、实际引擎及阶段耗时通过检查。机器报告更新为 `probe/reports/wide_summary_parity_20261023.json` 与 `probe/reports/wide_summary_existing_target_parity_20261023.json`。

证据仍限于合成 XLSX 和管道输入。未实测 Windows 原生文件弹窗或脱敏真实业务样本；既有目标执行仅使用临时测试 capability，生产 capability、完整真实模板保真、麒麟/UOS glibc 2.28 和离线 deb 未验收，PT-1-3/1-4 整体仍为**未验证**。

### 2026-10-25 PT-1-3/PT-1-4 空日期格式与缓存公式对拍

PyCode `io_excel._apply_data_date_format` 会对宽表输出中“数据日期”列的所有数据行设置 `yyyy/m/d`，包括日期未解析时的空值和文本。Rust 原先只在日期可解析时套用格式；现已修正为日期列的数值、空白和未解析文本均使用该格式。扩展 `probe/compare_wide_summary.py`，通过真实 Rust 菜单入口对无可解析日期、含 `B5 = 8+2.25`/缓存值 `10.25` 的合成输入执行 Python/Rust 对拍，检查汇总簿、目标候选和已提交新目标的值/类型/样式一致、缓存值按旧行为作为值投影、空日期仍为 `yyyy/m/d`，源与配置不变。

Windows 原生 `cargo fmt --manifest-path pytools-core/Cargo.toml -- --check`、`cargo build --offline --locked -p pytools-core --bin pytools` 通过；`cargo test --offline --locked -p pytools-core --test timeline_menu_cli -j 1` 8/8；Python 3.12.10/openpyxl 3.1.5 `probe/compare_wide_summary.py` 通过 3 个新目标菜单场景（含新增公式缓存/空日期边界）和 2 个既有目标替换场景。既有两个行头模式的汇总各比对 36 个单元格，既有目标分别比对 467/476 个单元格，样式均精确；既有目标仅使用临时测试 capability。机器报告更新于 `probe/reports/wide_summary_parity_20261023.json` 和 `probe/reports/wide_summary_existing_target_parity_20261023.json`。

证据仍为合成 XLSX；只验证读取已保存公式缓存，不触发公式重算。没有脱敏真实业务样本或生产 capability；Windows 原生文件弹窗操作、完整真实模板保真、麒麟/UOS glibc 2.28 和离线 deb 未验收，PT-1-3/1-4 整体仍为**未验证**。

### 2026-10-25 PT-1-3/PT-1-4 输入路径筛选对齐

对照 PyCode `run_wide_summary` 后，Rust 汇总入口现只读取已存在、扩展名在全局 `源文件扩展名` 中或为 `.xls` 的路径；其他扩展名与不存在路径在 reader 调用前忽略。user/developer 运行历史中的 `total_inputs` 同步按实际纳入处理的路径计数。Windows 原生 `timeline_menu_cli` 9/9 通过；新增 Python/Rust CLI 对拍以 3 个输入路径验证 `.xlsb` 和不存在的 `.xlsx` 被过滤、输出工作簿投影一致、源和配置未改变。机器报告：`probe/reports/wide_summary_input_filter_parity_20261025.json`。

本项仅为合成输入路径边界对拍；没有可用的脱敏 PT-1-3/1-4 真实业务样本，`.xls` 实际文件、原生文件弹窗、生产 capability、麒麟/UOS、glibc 2.28 和 deb 仍未验证，PT-1-3/1-4 整体状态不变。

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

扩展 probe/compare_summary_tools.py，通过真实 pytools summary-used-range CLI 对照只读 PyCode 基线；合成夹具覆盖两个工作簿（xlsx 与无 VBA 的 macro-enabled XLSM）、多个工作表、空/隐藏/veryHidden 工作表、从 D5:F6 开始的偏移 used range、格式定义 used-range 尾部、合并范围、日期/布尔/数值/文本/空白、有缓存及缺缓存公式和 7 种常见 Excel 错误值。19 行的值/类型、已填充单元格样式、工作表维度、冻结窗格、表头、公式/合并/定义名称/打印区和页面设置一致；空工作表产生默认范围行，隐藏及 veryHidden 工作表按 PyCode 工作簿顺序参与汇总；偏移区域保留源行号 5/6 和列坐标。机器报告 probe/reports/summary_used_range_parity_20261024.json 记录实际 calamine 读取、rust_xlsxwriter 新建、阶段耗时、偏移投影坐标及 7 个 OOXML 错误类型补丁；write_elapsed 覆盖工作簿构建、逐格写入、序列化和补丁。Rust 预扫描通过 extent-only 元数据建立范围映射，写出时每轮只保留当前输入 WorkbookDump；这是结构性内存占用变化，未测量 RSS。

另以原 PyTools 项目只读 XLSX 中的 13 份工作簿执行真实批次对拍，共 96 个工作表、203,580 条汇总数据，分成 4 批。四批的值/类型、已填充样式、工作表与输出维度、冻结窗格、表头、公式/合并/定义名称/页面属性均与 PyCode 一致；输入 SHA-256 前后相同，输出只存在于临时目录。最大批次包含 4 份校验工作簿、188,365 行；该批 Rust CLI 阶段 read 219,151 ms、write 59,415 ms，内部 total 279,292 ms；四批 Rust CLI 墙钟合计 347,861 ms（包含 WSL/Cargo 启动），不用于性能比较或默认路由调整，峰值 RSS 未测量。真实宽报表将源 B1（合并区 A1:AG1 的非锚点）由 Calamine 读出的值误写到摘要 E3753；PyCode 将合并子格视为空，Rust 现按 reader 的合并范围跳过该值。`2.报表校验关系.xlsx` 的 15 个 worksheet 均缺少 `<dimension>` 元素，但真实单元格/合并范围仍到第 70 行、第 257 列；PyCode 汇总 427 行，Rust 的 extent 扫描和批次结果一致。宽报表另外保留 1,212 个受支持的 Excel 错误值。机器报告 `probe/reports/summary_used_range_real_parity_20261024.json` 不含单元格值或输出簿。

Rust 在新建报表时保留受支持的 Excel 错误类型；使用 namespace-aware XML 解析器定位报表工作表关系与单元格，写入后由 Calamine 重新读取校验。合成 XLSX `t=e` 单元格 `#DATA!` 的旧 PyCode 输出为字符串；Calamine 0.37 对该 token 报错，Rust CLI 报出具体值、失败关闭且不产生部分报表。按 main 对 workspace reader 边界的确认，此类值登记为不支持/未验证；未添加普通值 OOXML 回退，也未改变共享 reader 或七种已支持错误映射。同步对齐 PyCode 的日期格式及页面边距。PyCode/Rust 的当前合成 XLSX 汇总表尺寸均为 20 行×14 列；NoComPyTools 为 19 行×14 列并跳过空工作表。两种实现都按工作簿顺序纳入隐藏和 veryHidden 工作表；D5:F6 偏移区域的 A19:C20/G19/I20 投影坐标与 PyCode 一致。NoCom 还将 汇总!H5:N5 的 7 种错误值输出为空值（PyCode 保留 error 类型），并将 汇总!E4 日期样式写为 yyyy-mm-dd（PyCode 为 yyyy-mm-dd h:mm:ss）。详细差异见机器报告的 nocom_pytools_vs_pycode。NoComPyTools 支持的 .xls/.xlsb 与 PyCode 的 .xlsx/.xlsm 范围差异仍待确认，不将 NoCom 差异或其独有格式登记为生产范围。

验证：cargo fmt -p pytools-core -- --check 通过；patch_error_cells 输出所有权回归单测和 summary_used_range_error CLI 集成测试各 1/1 通过；python pytools-core/probe/compare_summary_tools.py 合成 CLI 对拍通过；git diff --check 通过。cargo test -p pytools-core --lib --offline --locked -q 的 26 个库测试中 25 个通过，workbook_extraction_matches_original_python_helper 因当前 checkout 缺少且 Git 未跟踪 tests/fixtures/季度报表_县区报表_2024-12-31.xlsx 失败；本轮未创建替代夹具。Cargo.lock 未变。

原项目 13 份 .xlsx 已在 Windows Python 与 WSL Rust CLI 间完成四批真实对拍（96 张工作表、203,580 行），结果无差异；开发者视图在四批中均成功验证进度事件序号连续、历史 JSON 与 CLI 事件数一致；此外，WSL 交叉构建的 Windows GNU CLI 在 Windows PowerShell 上对同一 13 份 XLSX 四批重跑（203,580 行）均与 PyCode 投影一致，源文件大小/修改时间未变，事件历史全部匹配。普通用户、开发者及 off 模式的输出一致性由真实 CLI 集成测试验证。Windows 原生 Cargo 构建仍未验证（WSL 交叉构建的 Windows GNU 目标已在 PowerShell 实际运行并通过 VBA 样本对拍）；一般 XLSM（宏自由合成样本及官方真实 VBA 样本已测；原项目业务样本未提供）、麒麟/UOS、glibc 2.28、deb 和正式 peak RSS 尚未验证；PT-6-1 按原 Python 合同过滤 .xls/.xlsb，工作区这些格式的通用 reader 能力仍未验收。PT-6-1 整体仍未完成验收。

### 2026-10-24 PT-6-1 VBA 启用 XLSM 对拍

使用 XlsxWriter 固定提交 `a880ecc15c1573449dba051e5c21025f00c7c9fc` 的测试样本 `macro01.xlsm`（BSD-2-Clause），保存为 `tests/fixtures/summary_used_range/vba_enabled_xlsxwriter_macro01.xlsm` 并附上许可文本。样本含 13,312 字节 `xl/vbaProject.bin`；测试未执行宏。PyCode `run_summary_by_usedrange` 与实际 `summary-used-range` Rust CLI 对拍输出完全一致（1 行），两边均生成不带 VBA 部件的 `.xlsx`；输入未改变。Rust 实际使用 Calamine 和 rust_xlsxwriter，开发者事件历史一致，输出工作簿阶段 205 ms。机器结果保存在 `probe/reports/summary_used_range_parity_20261024.json`。原 PyTools 业务目录仍无含 VBA 的 XLSM 样本，目标 Linux 平台和 deb 仍待验收。

### 2026-10-24 PT-6-1 musl Linux 阶段验证

WSL Ubuntu（host glibc 2.39）上离线构建 `pytools` CLI 到 `x86_64-unknown-linux-musl` debug 目标成功。debug 产物为 ELF static PIE，`readelf -l` 未发现动态解释器段，`objdump -T` 未发现 GLIBC 动态符号；相同 musl 目标的 `summary_used_range_error` 集成测试 4/4 通过，包含 VBA 启用 XLSM 输入。随后以 `--release --offline --locked` 成功构建同目标；release 产物同为 static PIE，无动态解释器段或 GLIBC 动态符号。直接运行 release `pytools summary-used-range` 处理官方 VBA 启用 XLSM 测试夹具，PyCode/Rust 工作簿投影完全一致（1 行）；源文件未改变，新建 XLSX 不含 VBA，11 条开发者事件与保存历史一致，实际读取/写入引擎为 Calamine/rust_xlsxwriter，本次 `output_workbook` 阶段记录 104 ms。以上仅覆盖 WSL 上的 musl debug/release 阶段，不等于麒麟 V10/UOS V20 实机、deb 安装或 glibc 2.28 验收；PT-6-1 仍未完成最终验收。

### 2026-10-24 PT-6-1 Windows 目标运行阶段验证

Windows 主机未安装 Rust/Cargo。复用 WSL Ubuntu 中已安装的 MinGW 与 `x86_64-pc-windows-gnu` target，离线构建 `pytools.exe` 成功，并从 Windows PowerShell 进程直接运行 `summary-used-range`。固定 VBA 启用 XLSM 样本与只读 PyCode 的工作簿投影完全一致（1 行），使用 Calamine/rust_xlsxwriter；源文件未改变，输出 `.xlsx` 不含 VBA，11 条开发者事件与保存历史一致，`output_workbook` 阶段为 14 ms。随后 Windows 目标 CLI 对原 PyTools 只读 `xlsx/` 下 13 份 XLSX 分四批运行，203,580 行的完整工作簿投影逐批与 PyCode 一致；各批 26/14/27/63 条事件均与保存历史一致，源文件大小和修改时间未变。各批 `output_workbook` 阶段观测值为 472/6,665/1,438/55,095 ms，仅记录本次运行，不作性能比较。该证据覆盖交叉构建后的 Windows 目标二进制运行，不等于 Windows 原生 Rust 构建；麒麟/UOS 实机、deb 与 glibc 2.28 仍未验证。

### 2026-10-24 PT-6-2 页边距契约补齐

真实 `summary-comments` CLI 对拍发现新工作簿页边距和 PyCode 的 openpyxl 默认值不同；旧 probe 曾关闭页面设置比较。Rust 现对输出工作表显式写入 PyCode 默认左/右 0.75、上/下 1.0、页眉/页脚 0.5，并恢复完整页面设置对拍。Windows PowerShell 实际运行 Windows GNU 目标 CLI，对标准标记和重复行标记两类合成模板、各含 2 个源工作簿，输出均为 21 行，完整值/类型、样式、工作表维度、冻结窗格和页面设置投影与 PyCode 完全一致。扫描原项目中的 72 份 XLSX/XLSM 未发现含 PT-6-2 行/列/set 标记的真实模板；含批注的两个配置工作簿没有该汇总标记。

### 2026-10-24 PT-6-2 双视图运行日志接入

`summary-comments` 真实 CLI 已接入 PT-6-2 结构化运行记录，默认普通用户视图，并支持 `--log-view developer|off` 与 `--log-dir`。Rust 业务层按运行、输出目录准备、模板格式/读取/区域解析、输入筛选、源工作簿读取、区域提取、工作簿生成和保存发出阶段事件；事件记录进度、文件/行计数、阶段与总耗时、错误阶段及实际引擎。模板/源读取实际经 `excel-reader`（Calamine 读值并执行 namespace-aware OOXML 元数据补充），输出使用 `rust_xlsxwriter`。日志不含单元格原始值。

新增 `tests/summary_comments_cli.rs`，验证普通用户/开发者/off 三种视图、事件流与历史 JSON 一致、模板格式失败阶段、实际引擎与耗时字段、跳过输入无引擎、日志不含原始单元格值，以及三种视图生成的工作簿一致；与 PT-6-1 的 `summary_used_range_error` 回归共 6 项通过。Windows PowerShell 直接运行 Windows GNU 目标 CLI，对原 Python `run_summary_by_comment` 的标准标记及重复行标记两场景（每场景 2 个合成源簿、21 行）做完整工作簿投影对拍，值、类型、样式、维度、冻结窗格、页面设置/页边距均精确一致；developer stderr 事件和历史 JSON 24 条逐项一致。10,000 行合成 XLSX 上 developer/off 各运行 3 次，三次结果投影均一致，中位墙钟耗时约 2,640/2,618ms，developer 中位数约慢 0.8%；差异在该小样本的运行间波动范围内，不能外推为其他平台或真实模板的性能保证。机器结果见 `probe/reports/summary_used_range_parity_20261024.json`。

PT-6-2 原项目真实批注区域模板仍未找到，当前业务及日志对拍使用合成模板；本轮 Windows 实际入口通过交叉生成的 Windows GNU 二进制执行，当前 Windows 会话没有原生 Cargo。麒麟 V10/UOS V20、glibc 2.28 与 deb 仍待验收，不据此宣称 PT-6-2 整体完成。

### 2026-10-24 PT-6-2 合并区域子格取值对拍修复

合成源簿在 `A2:C3` 合并区域的非锚点 `B2` 中放入原始 OOXML 子格值。此前 Rust `summary-comments` 会优先返回 Calamine 的 `B2` 残值，而 PyCode `_gcv` 对合并区域内的任意坐标都返回左上角 `A2`。Rust 现先解析合并区域，再读取普通单元格值；以 WSL MinGW 离线构建 Windows GNU CLI，并从 Windows Python 3.12 运行实际入口与 PyCode 对拍，`D2/E2/D3/E3` 均为锚点值，3 行输出的完整工作簿投影精确一致，读取/写入引擎为 Calamine/rust_xlsxwriter，源夹具未改变。回归夹具为 `tests/fixtures/summary_comments/merged_children_source.xlsx`，机器结果补入 `probe/reports/summary_used_range_parity_20261024.json`。

`cargo fmt --manifest-path pytools-core/Cargo.toml -- --check`（WSL 离线）通过；`cargo test --offline --locked -p pytools-core --test summary_comments_cli` 3/3 通过。PT-6-2 真实批注模板、麒麟 V10/UOS V20、glibc 2.28 与 deb 验收仍未完成。

### 2026-10-24 PT-6-1 历史与备份真实工作簿扩展对拍

在原 PyTools 只读目录的 xlsx/历史测试数据/ 选取 4 份工作簿，并从 xlsx/备份/ 选取最小的 1 份工作簿；通过 Windows PowerShell 实际运行 Windows GNU 目标 pytools summary-used-range，与只读 PyCode run_summary_by_usedrange 输出对拍。5 份工作簿共 791 行：历史测试簿 448 行、备份簿 343 行。普通用户、开发者、off 三种日志视图生成的完整工作簿投影均与 PyCode 一致，值/类型、样式及工作簿输出属性无差异。开发者视图分别产生 26/11 条事件，均与保存的历史事件流一致；实际读取/写入引擎字段为 Calamine/rust_xlsxwriter。5 份源文件的大小及修改时间未变，输出仅写入临时目录，没有将源工作簿复制进仓库。

机器结果见 probe/reports/summary_used_range_real_parity_20261024.json 的 additional_project_assets_20261024；报告只记录来源目录类别、计数和对拍结果，不包含源工作簿名称或单元格内容。此次仅补充 PT-6-1 证据，不改变其未完成最终验收的状态。

### 2026-10-24 PT-6-1 最大备份簿真实负载对拍

从原 PyTools 只读目录 xlsx/备份/ 单独选取最大 XLSX（3,800,854 字节）；Windows PowerShell 运行 Windows GNU 目标 pytools summary-used-range，按 user/developer/off 三种日志视图逐次生成输出，并与只读 PyCode run_summary_by_usedrange 的完整工作簿投影比较。输出 39,171 行，三视图均与 PyCode 精确一致；实际读取/写入引擎为 Calamine/rust_xlsxwriter。developer 视图 18 条事件与保存历史逐条一致；user 历史状态成功；off 未创建运行记录。源簿大小和修改时间未变，临时输出已清理。本结果只作为该真实样本的保真证据，不作性能结论或其他格式推断。

机器结果见 probe/reports/summary_used_range_real_parity_20261024.json 的 additional_project_assets_20261024.largest_backup_workbook；只记录文件大小、来源目录和汇总结果，不记录工作簿名称或单元格内容。PT-6-1 整体仍未完成最终验收。

### 2026-10-24 PT-6-1 musl Linux 最大备份簿对拍

使用 WSL Ubuntu 上已构建的 x86_64-unknown-linux-musl release static PIE pytools CLI，单独处理原 PyTools 最大备份 XLSX（3,800,854 字节），并与 Windows 原 PyCode 的只读基线完整投影对拍。39,171 行输出完全一致；实际读取/写入引擎为 Calamine/rust_xlsxwriter；developer 18 条事件与保存历史逐条一致。源文件大小和修改时间未变，输出及日志位于临时目录并已清理。此项仅证明 WSL musl 阶段行为，不等于麒麟 V10/UOS V20 实机、deb 或 glibc 2.28 验收。

合入 main 的共享 reader comment author 与 editor merged-style 更新后，在当前 HEAD `15244db` 重新构建 musl release，并再次对同一最大备份 XLSX 与只读 PyCode 完整投影对拍。39,171 行数据（输出含标题共 39,172 行）在 user/developer/off 三视图均精确一致；user/developer 历史各有 18 条事件，developer stderr 事件流与保存历史一致，off 未写历史。实际引擎为 Calamine、`excel-reader::read_xlsx_worksheet_extents`、rust_xlsxwriter；源文件大小和修改时间未变。此结果仍仅代表 WSL musl 阶段，不代表麒麟/UOS 实机或 deb 验收。

原始机器证据见 `probe/reports/summary_used_range_real_parity_20261024.json` 的 `additional_project_assets_20261024.platform_validation.linux_musl_release_largest_backup`；main 同步后当前 HEAD 复验见 `linux_musl_release_after_main_sync_15244db`。


### 2026-10-24 PT-1-1 关键字筛选与 CLI 语义对拍

timeline_ops::keyword_match 改为复用共享 Python 语义匹配器，修复工作表关键字大小写敏感及 ASCII/中文逗号分隔差异。真实 Rust CLI 与只读 PyCode run_compare 基线在大小写变化、逗号关键字、未命中源和模板同路径过滤场景下，输出值、单元格类型、工作表维度/顺序、合并、冻结窗格、页面设置和定义名称一致。

真实配置的输入目录在当前开发机不存在；原项目 xlsx/ 中 70 份 XLSX 没有文件名和工作表同时命中时序规则，因此没有可运行的真实业务样本。本轮合成 CLI 对拍不构成 PT-1-1 完整验收。

输出样式差异已在 Rust `timeline-compare` 真实 CLI 中补齐：两张结果表的 17 个标题单元格现为粗体、水平居中、顶部对齐并带细边框。对拍通过 Windows 原生 CLI 与只读 PyCode 基线比较值/类型、工作表结构及有效 Excel 样式，17/17 个标题样式一致；样式比较将隐式默认 Calibri 11/主题文字色和省略/自动边框色视为等效，机器报告为 `probe/reports/timeline_compare_keyword_parity_20261024.json`。配置输入目录在当前开发机不存在，原项目 `xlsx/` 的 70 份 XLSX 无文件名和工作表同时命中时序规则，因此仍没有可运行的真实业务样本；PT-1-1 尚未完成真实业务验收。

### 2026-10-24 PT-6-1 剩余备份工作簿完整批次对拍

对原 PyTools 只读 xlsx/备份/ 中此前已测最小和最大端点以外的 51 份 XLSX，调用只读 PyCode run_summary_by_usedrange 和真实 Rust summary-used-range CLI，按 17 批、每批 3 份进行对拍。输出共 1,016,511 行；17/17 批的完整工作簿投影一致，覆盖值/行、类型、样式、工作表维度、冻结窗格、表头、合并、公式、定义名称、打印区域、页面设置、页边距及居中属性。开发者视图累计 549 条事件，与每次运行保存的历史事件流逐项一致；报告记录实际组件为 Calamine、excel-reader::read_xlsx_worksheet_extents 和 rust_xlsxwriter。

51 份源文件合计 84,340,966 字节，测试前后文件大小与修改时间均未变化；生成的工作簿和运行日志仅位于临时目录。机器报告 probe/reports/summary_used_range_backup_assets_20261024.json 只记录计数、投影结果及实际组件，不含源文件名或单元格内容。此项补充真实样本证据，不改变 PT-6-1 的最终验收状态：未知 #DATA! 错误值仍按 Calamine 未支持能力失败关闭；麒麟/UOS 实机、deb 与 glibc 2.28 仍未验证。

### 2026-10-24 PT-6-1 glibc 2.28 release 阶段对拍

使用 Zig 0.17.0 与 cargo-zigbuild，在 WSL Ubuntu 离线构建 `x86_64-unknown-linux-gnu.2.28` release `pytools`。ELF 为动态链接 PIE，`ldd` 成功解析系统运行库；`objdump -T` 观测到的最高 GLIBC 符号为 `GLIBC_2.28`。直接运行 release 二进制的 `summary-used-range` 处理原 PyTools 最大备份 XLSX（3,800,854 字节），39,171 行完整工作簿投影与只读 PyCode 一致；实际组件为 Calamine、`excel-reader::read_xlsx_worksheet_extents` 和 `rust_xlsxwriter`。developer 视图 18 条事件与保存历史一致，输出工作簿阶段记录 1,406 ms，输入大小和修改时间未变。

机器结果见 `probe/reports/summary_used_range_real_parity_20261024.json` 的 `additional_project_assets_20261024.platform_validation.linux_glibc_2_28_release_largest_backup`。此证据只覆盖 WSL 上的 glibc 2.28 ABI 构建检查和真实 CLI 对拍，不是麒麟 V10/UOS V20 实机或 deb 验收；PT-6-1 未知 `#DATA!` token、目标系统、deb 及最终验收仍未完成。

同一 glibc 2.28 release CLI 在合成 `getting_data_error.xlsx` 边界夹具上以退出码 2 失败关闭；运行记录为 `failed`，stage 为 `value_read`、error type 为 `value_read_error`，未生成部分汇总且运行日志不包含单元格 token。随后在 Windows 上直接运行只读 PyCode `run_summary_by_usedrange` 与 Windows GNU debug Rust CLI 对同一夹具进行实际入口对拍：Python 成功生成 1 行，`#DATA!` 输出为普通字符串类型 `s`；Rust 退出码 2，stage `value_read`，未生成部分汇总，仅生成失败运行记录。该结果确认真实迁移差异。Rust 继续失败关闭，不加入业务侧 OOXML 值回退；支持该 token 需由 main 协调共享 `excel-reader` 能力。

glibc 2.28 release CLI 对同一最大真实工作簿的 user、developer、off 三种日志视图均与 PyCode 完整工作簿投影精确一致（各 39,171 行）。user/developer 运行记录均成功且含 18 条事件；developer stderr 事件流与保存历史逐项一致；off 未创建运行历史且没有 CLI 进度输出（WSL 主机启动提示单独识别）。三种模式均验证了读取/写入引擎字段；输入大小和修改时间未变。该证据仍不代表麒麟/UOS 实机或 deb 验收。

### 2026-10-24 PT-6-1 glibc 2.28 release 备份簿完整批次对拍

将此前完成 debug CLI 对拍的 51 份 XLSX 再由 `x86_64-unknown-linux-gnu.2.28` release 二进制直接处理，并调用只读 PyCode 基线按原 17 批、每批 3 份逐批比较。51/51 份、1,016,511 行的值/行、类型、样式、工作表维度、冻结窗格、表头、合并、公式、定义名称、打印区、页面设置、页边距和居中属性完整投影一致；17/17 批通过，无失败组。开发者事件历史匹配，51 份源文件大小和修改时间均未变化。

机器报告 `probe/reports/summary_used_range_backup_assets_20261024.json` 的 `platform_validation.linux_glibc_2_28_release` 记录目标二进制及汇总结果。此项扩大 glibc release 的 PyTools 真实样本覆盖，不等于麒麟 V10/UOS V20 实机或 deb 验收；PT-6-1 最终验收仍未完成。

### 2026-10-24 PT-6-1 glibc 2.28 release 原始主样本对拍

使用已验证的 `x86_64-unknown-linux-gnu.2.28` release `pytools` 直接执行原 PyTools `xlsx/` 中 13 份主样本，沿用 PyCode/Rust 真实样本比较器按 4 批运行。96 张工作表、203,580 行在 4/4 批完整投影一致，覆盖值/类型、样式、工作表名称与维度、冻结窗格、表头、公式、合并、定义名称及页面属性；developer 事件数量为 26/14/27/63，均与保存历史匹配。实际引擎为 Calamine、`excel-reader::read_xlsx_worksheet_extents`、OOXML Patch 和 rust_xlsxwriter；13 份源文件的前后指纹一致。

机器结果见 `probe/reports/summary_used_range_real_parity_20261024.json` 的 `additional_project_assets_20261024.platform_validation.linux_glibc_2_28_release_primary_samples`。此项验证的是 WSL glibc 2.28 ABI release 阶段，不代表麒麟/UOS 实机或 deb 验收；PT-6-1 最终验收仍未完成。

### 2026-10-24 PT-6-1 glibc 2.28 release 历史测试簿对拍

使用 glibc 2.28 release `pytools summary-used-range` 处理原 PyTools `xlsx/历史测试数据/` 的 4 份工作簿。448 行完整工作簿投影与只读 PyCode 一致；developer 26 条事件与保存历史逐项匹配；实际引擎为 Calamine、`excel-reader::read_xlsx_worksheet_extents` 和 rust_xlsxwriter；源文件大小与修改时间未变。此结果补充主样本及备份簿以外的历史数据目录覆盖，不代表目标发行版实机或 deb 验收。

机器结果见 `probe/reports/summary_used_range_real_parity_20261024.json` 的 `additional_project_assets_20261024.platform_validation.linux_glibc_2_28_release_historical_samples`。

### 2026-10-24 PT-6-1 #DATA! 未知错误值兼容迁移与三视图对拍

通过 main 协调提交的共享 reader opt-in API `read_workbook_with_unknown_error_tokens_as_text`，仅在 PT-6-1 值读取处显式调用。对合成 XLSX 字面 `t=e` `#DATA!`，共享 reader 在临时包中将该未知错误 token 规范化为文本后交由 Calamine 读取；PT-6-1 仍由 Rust 汇总并用 rust_xlsxwriter 新建输出。普通 `read_workbook` 默认行为未改变；未增加业务侧 OOXML 值 fallback，也未推断其他 token/格式的支持。

Windows GNU debug CLI 在 PowerShell 与只读 PyCode `PyCode.core.summary_321_322.run_summary_by_usedrange` 的三视图对拍均为 1 行完整工作簿投影一致。WSL 上 `x86_64-unknown-linux-gnu.2.28` release CLI 的 user/developer/off 三视图也均为 1 行精确一致；`objdump -T` 最高 GLIBC 符号为 `GLIBC_2.28`。两边开发者视图各有 12 条事件，其中兼容事件准确报告实际读取引擎和阶段耗时，事件流与历史 JSON 匹配且日志不含单元格值；输入夹具未变化。机器详情见 `probe/reports/summary_used_range_parity_20261024.json` 的 `unknown_cell_error_probe`。

验证包括 `cargo fmt --check -p pytools-core`、`cargo test --offline --locked -p pytools-core --test summary_used_range_error`、Windows GNU CLI 交叉构建及执行、WSL glibc 2.28 release CLI 执行与完整 Python/Rust 工作簿投影对拍。麒麟 V10/UOS V20 实机、deb 安装和离线交付仍未验证；PT-6-1 整体仍未完成最终验收。

共享 opt-in reader API 接入后，Windows GNU debug CLI 对原 PyTools xlsx/ 13 份主样本按 4 批完整回归；96 张工作表、203,580 行在 4/4 批与只读 PyCode 完整投影一致，各批 reader 为 Calamine、未知错误 token 归一化计数为 0，writer 为 rust_xlsxwriter，开发者事件数 26/14/27/63 均与历史 JSON 匹配，输入前后指纹一致。公共 reader 更新到 `aa2797d` 后，重新构建 Windows GNU CLI 并运行完整 `compare_summary_tools.py`：合成汇总为 19 行，7 个已支持 Excel 错误单元格及 `#DATA!` 字符串语义与 PyCode 一致；上述 13 份样本再次 4/4 完整投影一致。之后同步的 `9c869de` 只改共享 `excel-editor` 工作表布局能力；在该 HEAD 上 `cargo fmt --check -p pytools-core`、`cargo test --offline --locked -p pytools-core --test summary_used_range_error`（4/4）和 Windows GNU CLI 构建均通过，未因该编辑器变更重复读取对拍。当前环境缺少 Zig，未能在本轮重建最新 glibc 2.28 release；既有 release 证据见 `additional_project_assets_20261024.platform_validation`，不外推为最新共享 crate 构建或目标发行版验收。机器结果见 `probe/reports/summary_used_range_parity_20261024.json` 与 `probe/reports/summary_used_range_real_parity_20261024.json`。

### 2026-10-24 PT-6-1 PyCode/NoComPyTools 真实样本变体对拍

在 Windows Python 3.12.10 中分别使用 PyCode 项目虚拟环境（openpyxl 3.1.5）和 NoComPyTools 项目虚拟环境（openpyxl 3.1.5、python-calamine 0.9.0），与当前 Windows GNU `summary-used-range` CLI 对同一份原项目真实 XLSX 对拍。PyCode/Rust 完整投影精确一致，均输出 7,462 行；NoComPyTools 输出维度相同，但值、类型、样式和行投影存在差异：1,213 个单元格，其中 1,212 个 PyCode 错误类型单元格在 NoCom 输出中为空，另有 1 个空值变为字符串。Rust 实际使用 Calamine 与 rust_xlsxwriter，并由既有 OOXML Patch 保留 1,212 个错误单元格；源文件大小和修改时间未变，报告不含源文件名或单元格值。

该结果只确认此真实样本上的变体差异；不据此判定 NoComPyTools 独有行为属于最终产品范围，也不改变已确认的 PyCode 基线。原项目 xlsx/ 下其余 XLSX 的同一汇总入口变体对拍见下节；其他格式和 NoCom 独有行为仍待逐项确认。机器结果见 `probe/reports/summary_used_range_real_parity_20261024.json` 的 `nocom_pytools_pt61_variant_comparison`。


### 2026-10-24 PT-6-1 扩展 Excel 错误值 token 对拍

新增合成 `error_token_matrix.xlsx`，覆盖 7 种常见 Excel 错误和 8 种扩展字面 `t=e` token。只读 PyCode 与真实 `pytools summary-used-range` Windows GNU CLI 输出的完整工作簿投影一致：7 个常见错误保留为 error 类型，8 个扩展 token 输出为文本。Rust 使用 Calamine + OOXML error-token normalization 和 rust_xlsxwriter；阶段计数分别为 OOXML Patch 保留 7 个错误值、文本兼容读取 8 个未知 token。开发者事件历史与 CLI 事件流一致，历史不含 token 内容，输入夹具字节未变。同一夹具再由 WSL Ubuntu `cargo run` Linux debug CLI 处理，与 PyCode 投影仍完全一致；14 条开发者事件与历史匹配，结果与 Windows GNU 记录并列。该次不是 glibc 2.28 release 验证。

Rust 集成回归 `summary_used_range_cli_preserves_common_and_textual_error_token_matrix` 检查真实 CLI 报告、输入 XML 错误 token、输出类型与值以及日志隐私。Windows Python/Rust 对拍机器报告为 `probe/reports/summary_used_range_error_token_matrix_20261024.json`。验证：Windows Python 3.12.10 对拍通过；WSL `cargo fmt -p pytools-core -- --check` 通过；WSL `cargo test --offline --locked -p pytools-core --test summary_used_range_error -q` 为 5/5 通过。

本证据仅覆盖列出的 XLSX 字面 token，不推断其他 token、XLS/XLSB、麒麟/UOS、deb 或生产默认路由；PT-6-1 最终验收状态不变。无真实业务数据或凭据进入夹具和报告。

### 2026-10-24 PT-6-1 PyCode/NoComPyTools 全量 XLSX 变体对拍

对原项目 xlsx/ 递归发现的 70 份 XLSX，分别运行原 PyCode 与 NoComPyTools 的按使用区域汇总入口，并按最多 3 份工作簿分为 24 批比较完整输出投影。24/24 批有结果，其中 12 批精确一致、12 批存在差异；一致批次包含 34 份输入，差异批次包含其余 36 份输入（比较粒度为批次，不能归因到单个工作簿）。两边失败输入和缺失输出均为 0。汇总行数（含标题行）分别为 1,260,077 和 1,258,014，PyCode 多 2,063 行。差异涉及工作表维度，以及汇总表的行、样式、类型、值和表头加粗属性。

比较运行于 Windows Python 3.12.10；两边 openpyxl 为 3.1.5，NoCom 环境的 python-calamine 为 0.9.0。源工作簿大小和修改时间均未改变。按输出坐标计数的单元格差异可能受删除行后的坐标移动影响，不代表不同源单元格数。该对拍只比较两种 Python 变体，不新增 Rust 行为结论、不改变 PyCode 生产基线；NoCom 独有行为仍标为待确认。机器结果见 `probe/reports/summary_used_range_nocom_assets_20261024.json`，不记录源文件名、单元格值、路径或哈希。

### 2026-10-24 PT-6-2 空 `set()` 批注标记兼容

真实 PyCode `run_summary_by_comment` 的 `set` 正则要求括号内至少有一个字符；空 `set()` 不构成集合标记。Rust 原解析曾接受空内容，导致多出一个空表头和数据列。`summary_ops::parse_set_name` 现跳过零长度捕获，同时保留 `set( )` 经裁剪后为空名称的原 Python 行为。新增真实 `summary-comments` CLI 集成回归与专用 Python/Rust 对拍：2 行输出的值/类型、样式、维度、冻结窗格和页面设置完整投影精确一致，输出 5 列且无空集合列；源模板和源工作簿未变化。机器报告为 `probe/reports/summary_comments_empty_set_marker_parity_20261024.json`。

验证：WSL Ubuntu `cargo test --offline --locked -p pytools-core --lib comment_region_tests` 2/2、`cargo test --offline --locked -p pytools-core --test summary_comments_cli` 4/4、Windows `cargo fmt --manifest-path pytools-core/Cargo.toml -- --check` 通过。Windows SDK 后续由 main 配置完成；Windows 原生 MSVC `cargo test --offline --locked -p pytools-core --lib comment_region_tests` 2/2、`--test summary_comments_cli` 4/4，且原生 Rust CLI 与 PyCode 此场景完整投影一致。PyTools crate 全量测试有一项无关的既有失败：`workbook_extract` 单测引用的 `tests/fixtures/季度报表_县区报表_2024-12-31.xlsx` 在当前 checkout 缺失（WSL lib 测试为 29/30）。该修复只覆盖合成模板边界；PT-6-2 真实业务批注模板、麒麟/UOS、glibc 2.28 与 deb 验收仍未完成。

### 2026-10-24 PT-6-1 Windows 原生 MSVC CLI 对拍

Windows SDK 安装完成后，使用 native `cargo test` 构建的 `target/debug/pytools.exe` 运行既有 `compare_summary_tools.py`。PyCode 与 Rust 真实入口对原 PyTools `xlsx/` 目录 13 份工作簿分 4 批完整投影对拍一致：96 张工作表、203,580 行；值/类型、样式、工作表及维度、公式、合并、定义名称、冻结窗格与页面属性通过。1,212 个 Excel 错误单元格由既有 OOXML Patch 保留；批次事件流与保存运行记录一致，源文件前后完整性检查通过。合成错误值、未知 `#DATA!`、扩展错误 token、VBA 启用测试簿、输入扩展名过滤和双视图测试均通过。NoComPyTools 与 PyCode 的既有差异保持为变体差异，不作为 PyCode/Rust 失败或生产合同变更。

Windows 原生测试：`cargo test --offline --locked -p pytools-core --lib comment_region_tests` 2/2；`--test summary_comments_cli` 4/4；`--test summary_used_range_error` 5/5；`cargo fmt --manifest-path pytools-core/Cargo.toml -- --check` 通过。Windows 原生结果记录于 `probe/reports/summary_used_range_windows_native_20261024.json`；此前 WSL/glibc 阶段报告保留不变。运行时间仅记录该次 Windows debug 环境，不作跨实现性能结论；本轮没有重建 glibc 2.28 release。PT-6-1 整体仍待原 PyTools 业务含 VBA 的 XLSM 样本、麒麟 V10/UOS V20、deb 和离线交付验收。

### 2026-10-24 PT-6-1 当前分支 musl release 真实样本复验

在当前 `pytools-dev` HEAD `9837336` 使用 WSL musl release 二进制运行真实 CLI `summary-used-range`，对原 PyTools 最大备份 XLSX（3,800,854 字节）的 39,171 行汇总，与只读 PyCode 完整工作簿投影精确一致。实际引擎为 Calamine 与 rust_xlsxwriter；developer 视图 18 条事件与保存历史逐项匹配，源文件大小和修改时间未变。ELF 检查为静态 PIE，`ldd` 报告静态链接。

当前 HEAD 的 Windows 原生 PT-6-1/PT-6-2 回归 9/9 通过；VBA 启用 XLSM 合成夹具由原 PyCode 与真实 Rust CLI 输出精确一致，宏部件未复制到新汇总簿。此夹具不是原 PyTools 业务 XLSM 样本。

机器结果见 `probe/reports/summary_used_range_real_parity_20261024.json` 的 `additional_project_assets_20261024.platform_validation.linux_musl_release_current_9837336`。本轮未能重建 glibc 2.28 release：WSL 缺 Zig，Windows/WSL 对 Zig 官方下载站的 TLS 连接均超时。此项只证明当前 HEAD 的 WSL musl 阶段行为，不代表麒麟 V10/UOS V20 实机、glibc 2.28、deb 或离线验收；PT-6-1 最终验收仍未完成。

### 2026-10-24 PT-6-1 当前 HEAD 真实样本复验

在 pytools-dev HEAD a376a76 运行现有 PT-6-1 真实样本 probe。Windows Python 基线与 WSL Linux debug cargo run 对原 PyTools xlsx/ 的 13 份 XLSX 分 4 批对拍，96 张工作表、203,580 行完整投影精确一致；值/类型、样式、工作表名称与维度、冻结窗格/表头加粗、公式/合并/定义名称/页面属性通过。批次事件数依次为 26、14、27、63，均与保存运行历史匹配；输入前后 SHA-256 相同。Rust CLI 使用 Calamine、excel-reader::read_xlsx_worksheet_extents、rust_xlsxwriter；其中含 Excel 错误值的批次还使用 OOXML Patch。此次 WSL 验证包含 cargo run debug 调用，不是 glibc 2.28 release 或目标发行版验证。

在此前 HEAD a376a76 上，Windows 原生 target/debug/pytools.exe 已单独复验最大批次：4 份工作簿、188,365 行完整投影精确一致；63 条 developer 事件与保存历史匹配，输入前后 SHA-256 一致。随后在当前分支 HEAD 19c96fc，使用 Windows 原生真实 CLI 对全部 13 份工作簿分 4 批全量重跑，96 张工作表、203,580 行完整投影精确一致；值/类型、已填单元格样式、工作表及维度、冻结窗格/表头加粗、公式/合并/名称/页面属性通过。批次 developer 事件数为 26、14、27、63，均与保存历史匹配，源文件前后 SHA-256 相同。机器摘要位于 additional_current_head_windows_native_full_20261024；本结果记录正确性，不作 Windows/WSL 性能比较。

机器结果见 probe/reports/summary_used_range_real_parity_20261024.json。此前独立 Windows 全量与 glibc 2.28 release 报告保留各自测量范围。原项目业务含 VBA 的 XLSM 样本、麒麟 V10/UOS V20、deb 与离线交付仍未验收；PT-6-1 整体尚未最终验收。

### 2026-10-24 PT-1 时序交互菜单入口

增量主菜单类别 1 已接入 1-1 表头路径比对、1-3/1-4 两种宽表汇总和 1-5 预校验，入口复用既有 Rust CLI 业务处理。Windows 原生 Rust CLI 的合成 PyCode/Rust 菜单对拍通过：1-1 五个所选文件中三个匹配，输出结构、值/类型及有效样式一致；1-3/1-4 工作簿投影与有效样式一致并生成目标追加计划，配置目标未改写；1-5 的有效、详细错误和缺必需列三种场景计数一致。tests/timeline_menu_cli.rs 1/1 通过。机器摘要见 probe/reports/timeline_menu_parity_20261024.json。

证据仅覆盖合成 XLSX 和手工路径输入。真实业务样本、原生文件弹窗、目标追加生产 capability、结构化用户/开发者/off 日志、麒麟/UOS、glibc 2.28 和 deb 仍未验证。

### 2026-10-24 PT-2 去重交互菜单接入

增量主菜单类别 2 已接入 2-1 至 2-6 六项去重入口。批注检查和按配置检查按 PyCode 行为生成红色标记编辑计划；删除、配置追加也只生成计划，不在菜单中直接改写源或目标工作簿；每种菜单操作写入独立计划目录，避免覆盖既有计划。Windows 原生 Rust 菜单集成测试 1/1 覆盖六项入口、重复计数、计划文件、候选/目标未写入及源文件字节未变。底层业务语义的既有 PyCode/Rust 合成对拍见 probe/reports/dedup_parity_20261023.json。

真实业务样本、生产 ExcelModifier capability、计划执行菜单、结构化用户/开发者/off 日志、麒麟/UOS、glibc 2.28 和 deb 仍未验证。

### 2026-10-24 PT-3 打印交互菜单接入

增量主菜单类别 3 已接入 3-1 按批注保留格式、3-2 按批注快速复制、3-3 按配置准备打印文件和 3-4 配置预校验。评论标记的两种计划分别映射保留批注/样式和快速模式；配置输出保持 plan-only，未生成候选文件，也不渲染或导出 PDF。Windows 原生 Rust 菜单集成测试 1/1 验证四个入口、两个打印区域、模式属性、目标未创建和配置预校验；底层 PyCode/Rust 合成对拍见 probe/reports/print_comment_plan_parity_20261023.json、print_comment_execution_parity_20261023.json、print_precheck_parity_20261023.json 和 print_config_execution_parity_20261023.json。

真实业务模板、生产 capability、结构化用户/开发者/off 日志、麒麟/UOS、glibc 2.28 和 deb 仍未验证；页面设置工作流不包含打印渲染。

### 2026-10-24 PT-4 转换与文件操作交互菜单及提交检查对拍补充

PyTools 主菜单类别 4 已通过真实 `menu` 交互入口接入 4-1 至 4-10 十个文件操作：批量重命名、Excel/Word 转换、Sheet 改名计划、内容前缀添加/取消、局部文件名映射、提交检查、排序前缀添加/取消。Windows 实现了按业务格式筛选的文件对话框、当前层/递归目录选择及路径手工输入；4-4 仍只生成 `ExcelEditPlan`，不执行未获 capability 验收的工作簿修改。

新增 `tests/conversion_menu_cli.rs` 直接启动 Rust `pytools menu`，合成验证十个分支均接通；同格式复制样本验证输出位置和源文件内容，`.xls/.doc` 样本在显式测试缺失 sidecar 时报告“该功能需要外部引擎”、不生成输出且不改源文件。该项不是实际格式转换或 sidecar 保真验收。

`submission_check.rs` 对只读 PyCode `run_submission_check_48` 建立了合成基线及真实 Rust CLI 对拍，报告四张表的顺序、单元格值/类型、维度、表头加粗和统计均一致（3 条启用规则、3 个文件、提交 2、缺失规则 1、未识别文件 1、多重命中 1）。对拍发现并修正了多重命中描述缺少关键词，以及命中次数/配置行被写为文本的问题。报告与基线位于 `probe/reports/submission_check_python_baseline.json`，基线生成器位于 `probe/submission_check_baseline.py`；输入、配置和路径均为合成内容。

本轮将 3-3/4-6 至 4-10 的 Rust API 改为语义名称，CLI 和 Rust 测试使用新名称；原数字名称保留为弃用兼容别名，未更改 CLI 命令、交互菜单序号或 Excel 配置字段。相关 Python/Rust 函数级对拍继续保留其历史操作号，便于定位基线。

Windows 原生验证：`cargo fmt --manifest-path pytools-core/Cargo.toml -- --check`、`cargo check --offline --locked -p pytools-core --bin pytools` 通过；12 个关联集成测试目标 28/28 通过，`cargo test --offline --locked -p pytools-core --lib conversion::tests` 4/4 通过。未运行 PyTools 全量测试或 Linux/目标系统验证。

最后补入 Windows 目录弹窗的 OLE STA 初始化后，原生 `cargo check` 和格式检查再次通过；真实菜单与提交检查 CLI 测试在 F 盘独立验证目标中 2/2 通过。默认 MSVC 调试链接曾因 PDB `LNK1318 LIMIT` 失败，独立无调试符号验证成功；之前 28/28 的完整关联回归发生在此次弹窗初始化之前。

真实业务文件、用户实际原生文件/目录对话框操作、有效 sidecar 对 `.xls/.et/.doc` 的转换结果、隔离运行打包、麒麟/UOS、glibc 2.28 与 deb 均未验收。此更新证明菜单路由及合成对拍，不将转换/文件操作整组标为已验收。

### 2026-10-25 PT-4-2 native XLS 空文本安全预检

对一个临时副本的真实 `.xls` 样本复验 `pytools convert-excel --engine native`。上一轮 native 候选读取 18,080 个存储值坐标，写出并由共享 reader 重读后为 15,957 个；少掉的 2,123 个源坐标均为类型化空文本，没有额外输出坐标，因此候选未通过数据坐标校验且没有发布。`conversion.rs` 现于创建临时候选前拒绝无法保真的空文本值，并提示需要外部保真引擎；对应公式单元格的空缓存仍跳过该值预检，按原公式文本路径处理。

真实 `convert-excel` CLI 复跑退出码 2，结果为 0 转换、1 失败，输出目录无文件。Windows 原生 `conversion::tests` 6/6、`conversion_menu_cli` 5/5，格式检查通过。未调用 Excel/COM/LibreOffice；该样本的 Python/保真转换对拍仍未验证，C01 native 路径不接受此空文本输入，PT-4-2 整体状态保持**未验证**。机器摘要为 `probe/reports/PT-4-2_native_empty_text_guard_20261025.json`。共享 reader/writer 对显式零长度文本的能力问题已提交 main 协调；未修改共享 crate，也未登记生产能力。

### 2026-10-24 PT-7-5 地区总分校验交互入口

主菜单新增类别 7，并接入原 PyCode 菜单 7-5 地区总分校验：文件多选、当前层/递归文件夹扫描和手工路径均进入既有 `region-sum-check` Rust CLI；输出目录和日志目录读取全局配置。当前入口只接受 `.xlsx/.xlsm`；手工指定 `.xls` 会明确提示先显式转换，避免在业务路径使用未批准的 COM 读取。核对仍为只读输入并新建结果工作簿。

新增 `tests/batch_tools_menu_cli.rs` 启动真实 `pytools menu`，通过 7→5 对五个地区核对合成夹具（正常、不一致、无目标、非数值/缺项、损坏文件）验证统计、结果表、日志和输入不变；额外的合成 `.xls` 路径被跳过并提示显式转换。与既有 `region_sum_parity` 的 PyCode 基线逐值/格式对拍一并回归通过。该证据仅覆盖 PT-7-5 合成样本，不代表 7-1 至 7-4、7-6 至 7-8 或整个类别已接入。

Windows 原生 Rust：格式检查通过；`batch_tools_menu_cli`、`region_sum_parity`、`timeline_menu_cli`、`conversion_menu_cli`、`summary_menu_cli` 共 11/11 通过。为消除 Windows 并行测试时纳秒临时目录名可能碰撞，`region_sum_parity` 测试目录追加进程内原子序号。真实业务模板、Windows 原生弹窗手动操作、目标平台和 deb 仍未验收；未运行全量 PyTools 测试。

### 2026-10-24 PT-7-1 村镇银行拆分菜单与运行事件

在既有 `batch_ops::build_village_split_plan` 计划逻辑上接入主菜单 7→1 和 `village-split` 批处理 CLI。菜单选择本期、上期及 `.xlsx` 目标文件；默认只生成计划、不修改原文件。实际执行仍要求显式 ExcelModifier capability 并经过候选校验、安全替换。补入 `user`/`developer`/`off` 事件视图和配置日志目录中的 PT-7-1 运行历史；历史仅保留状态、计数、阶段时间及实际引擎，不写文件名或单元格值。

Windows 原生 `cargo fmt --manifest-path pytools-core/Cargo.toml -- --check` 通过；`cargo test --offline --locked -p pytools-core --test batch_tools_menu_cli -j 1` 4/4 通过，覆盖 7→1 计划只读、用户日志历史、开发者实际引擎事件、off 模式不写历史，以及真实菜单经合成 capability 执行后的 Excel 值、同表与跨表公式、全局定义名称、合并区域、批注、行高和非目标工作表保真。真实 Rust 菜单与只读 PyCode `feature_1_1_split_village_bank.run_split_village_bank` 执行对拍使用 Python 3.12.10/openpyxl 3.1.5、3 张填充表及未匹配/未修改表；`probe/reports/village_split_menu_parity_20261024.json` 状态为 `passed_synthetic_menu_execution_parity`，全部工作表值投影一致，输入本期/上期不变，目标已安全提交；报告不含单元格值。

共享修复由 main 完成并提交，本分支经 Git cherry-pick 包含 `4a08852`（目标 worksheet 范围内插行及共享回归）和 `afaed4f`（共享修改记录终结）。共享回归由 main 报告 `excel-editor/tests/insert_rows` 1/1 通过；本分支没有重复运行该共享测试，以上 PT-7-1 菜单测试和多工作表执行对拍验证了业务调用链。

仍未取得授权的真实业务模板或真实业务样本；synthetic capability 仅用于隔离测试，不登记为生产 Passed 能力或默认路由。真实模板/平台、Windows 弹窗手工验收、目标 Linux/glibc 2.28/deb 验收未执行；PT-7-1 的生产验收状态仍为未验证。

### 2026-10-24 PT-7-2 机构名称规范化菜单入口

在既有 `institution-normalize` CLI/计划实现上接入主菜单 7→2，映射从全局配置读取；默认仅生成计划，显式 capability 才能执行候选验证和安全替换。菜单拒绝未验收格式；PT-7-2 当前只允许 `.xlsx`，原 Python 的 `.xlsm`/VBA 路径仍未迁移。普通用户与开发者视图沿用现有结构化事件；新增 `off` 视图，不输出阶段事件、不写历史文件。

Windows 原生验证：`batch_tools_menu_cli` 6/6、`institution_normalize_batch_cli` 5/5、`institution_normalize_plan_parity` 1/1 通过，格式检查通过。真实菜单覆盖计划只读、配置映射、`.xlsm` 跳过、user/developer/off 视图与临时 synthetic capability 的真实 CLI→ExcelModifier 执行；计划语义及执行后所有工作表的值、公式、合并区域、批注投影均与既有 PyCode 基线一致，编辑器候选校验通过计划声明的样式与未修改包部件保留项。以上均为合成 XLSX 证据；真实模板/样本、`.xlsm` 保真、生产 capability、目标平台和 deb 未验收，PT-7-2 仍为未验证。

### 2026-10-24 PT-7-3 外资行删除交互菜单入口

在既有 `remove-foreign-bank` Rust 批处理入口和 `batch_ops::build_remove_foreign_plan` 上接入主菜单 7→3。菜单从全局配置读取机构映射，只选择 `.xlsx`；无 capability 时保存 ExcelEditPlan 至 `output/plans/PT-7-3` 且保持源文件不变，显式提供 capability/platform 时通过 ExcelModifier 生成同目录 `(金融局)` 副本。`user`/`developer` 记录阶段、进度、耗时与实际读取/修改引擎到配置日志目录；历史报告不记录文件名、路径或单元格内容。`off` 保留最终摘要，不输出事件或写历史。

新增菜单集成覆盖计划模式、无操作工作簿、手工选入 `.xls` 后跳过、源文件只读、用户/开发者/off 视图，以及以临时 synthetic capability 经真实菜单→CLI→ExcelModifier 执行。执行后检查删除行后的值和公式、原始输入不变；同一 Rust 执行端到端回归仍覆盖跨表公式、绝对定义名称、批注、合并区域、样式、行高及无操作时的字节复制。Python 基线与 Windows 原生 Rust CLI 的合成对拍结果见 `probe/reports/PT-7-3_execution_parity_windows_native_20261024.json`：2 个输入、4 张分机构表、删除 3 行、失败 0；计数、工作表顺序和值投影一致，源输入未修改，无操作输出为字节复制。合成 Excel COM 对象级对拍见 `probe/reports/PT-7-3_excel_com_parity_20261023_r8.json`。

Windows 原生验证：`cargo fmt --manifest-path pytools-core/Cargo.toml -- --check` 通过；`batch_tools_menu_cli` 10/10、`foreign_bank_batch_execution` 2/2 通过。完整关联测试使用 `CARGO_TARGET_DIR=F:\pytools-dev-verify-20261024` 与 `RUSTFLAGS=-C debuginfo=0`，避开默认目标目录链接 `pytools.exe` 时既有的 MSVC PDB `LNK1318 LIMIT`。Python 3.12.10 / openpyxl 3.1.5 基线与 native Rust CLI 合成对拍通过。既有共享 `excel-editor` 弃用 API 与未使用导入警告未在本功能中修改。

PT-7-3 生产验收仍为**未验证**：没有真实工作簿/配置或授权样本，没有在 Windows 桌面人工操作文件弹窗，未迁移旧 `.xls` COM 路径及 `.xlsm` VBA 保留，生产 capability 未登记；麒麟/UOS、glibc 2.28 和离线 deb 未验证。本轮对拍和执行 capability 均为合成 XLSX，不更改 ExcelModifier 默认路由。

### 2026-10-24 PT-7-6 区域提取菜单与运行记录

在既有区域提取候选与 `extract-sheets` CLI 上接入主菜单 7→6。入口读取配置中的启用任务；无任务时中止；整表任务在没有 ExcelModifier capability 与 platform 参数时先中止，再选择输入。输入支持多选，但当前只把 `.xlsx` 交给读取器，其他扩展名在菜单处提示并跳过。普通无批注版保持 PyCode `feature_1_6_extract_sheets.run_extract_sheets` 的迁移范围，COM 批注版本不接入。

核心通过共享 `excel-reader`/Calamine 读取区域值，以 `rust_xlsxwriter` 新建 XLSX，重开校验工作表名称、尺寸和非空格数量，并在校验后用不覆盖目标的提交方式发布结果。读取、输出写入/修改、校验和提交阶段提供进度、阶段耗时及实际引擎事件；菜单支持 user/developer/off 视图，user/developer 在配置日志目录写 PT-7-6 运行历史，记录统计、耗时和引擎，不写单元格值。整表导入继续走 ExcelModifier capability；本轮没有登记生产能力或改变默认路由。

Windows 原生 `cargo fmt --manifest-path pytools-core/Cargo.toml -- --check` 通过；`sheet_extract_area_operation_parity` 7/7、关联的 `batch_tools_menu_cli` 10/10 通过。新增真实 `pytools menu` → 7→6 集成对拍：两个有效合成工作簿和一个损坏 XLSX 得到 4 张表、1 个输出、1 个失败；另有占位 `.xls` 在菜单处过滤。输出 44 个比较单元格的值/类型、工作表名和尺寸与只读 PyCode 基线一致，输入字节不变；user/developer/off 视图回归通过；开发者事件与运行历史记录 Calamine 和 rust_xlsxwriter，历史不含抽样单元格值。机器报告为 `probe/reports/sheet_extract_menu_parity.json`。

2026-10-24 的证据仅覆盖合成 XLSX、管道输入的手动分号路径和区域模式；Windows 原生文件弹窗未操作。整表菜单候选对拍见下节。真实业务模板、其他格式及麒麟/UOS glibc 2.28/deb 未验证，PT-7-6 完整功能及生产验收仍为未验证。

### 2026-10-25 PT-7-6 整表菜单端到端合成对拍

在普通无批注版的真实 `pytools menu` → 7→6 入口中执行一个合成 XLSX 整表任务，任务配置、源文件选择、输出提交和运行记录均走产品路径。测试仅创建 fixture-scoped `ImportRange` 临时 capability，明确匹配该合成工作簿与 `umya_spreadsheet`；没有登记生产 capability 或改变默认路由。Python 基线由 `probe/create_sheet_extract_full_menu_baseline.py` 调用只读 `PyCode/core/feature_1_6_extract_sheets.py::_copy_full_sheet_with_style` 生成，工作簿只包含合成值和样式。

Rust 输出与 Python helper 可读部分对拍：8×5 矩形的 40 个值/类型位置、工作表名、存储范围、A1:C1 合并、行 12 和列 G 的尺寸/隐藏/outline，以及普通版丢弃 D5 批注，均匹配；被选取的粗体、填充色和数字格式按源工作簿属性读回一致。Python helper 将跨簿私有 style 数组复制后生成了无法读取的样式表，openpyxl 重开样式时抛出 `IndexError: list index out of range`，因此报告状态为 `partial`，不宣称 Python 输出样式文件级对拍通过。报告分别为 `probe/reports/sheet_extract_full_menu_python_baseline.json` 和 `probe/reports/sheet_extract_full_menu_parity.json`。

端到端测试发现通用 staged-workbook 校验用 Calamine 的非空值尺寸判断整表边界，会漏掉带样式的空白 E8，并错误拒绝有效 8×5 输出。`validate_staged_workbook` 现以共享 OOXML worksheet extent reader 校验存储单元格/合并范围边界，仍由 Calamine 校验值、类型与非空格数。验证没有写入单元格数据的运行历史、源文件字节未变，并记录实际引擎 Calamine、rust_xlsxwriter 与 umya_spreadsheet。

Windows 原生 Rust（rustc/cargo 1.99.1）验证：`cargo fmt --manifest-path pytools-core/Cargo.toml -- --check` 通过；`sheet_extract_area_operation_parity` 8/8、`batch_tools_menu_cli` 10/10、`excel-editor` 的 `detects_stored_content_in_merged_non_anchor_cells` 1/1、`import_range` 1/1 通过；`git diff --check` 无空白错误。仍未操作 Windows 原生文件弹窗，未使用真实业务模板或真实生产 capability，未验证其他格式、麒麟/UOS glibc 2.28 和 deb；整表生产能力与 PT-7-6 完整验收继续标记未验证。

### 2026-10-25 PT-7-7/PT-7-8 涉农贷款交互菜单接入

在既有 `rural-loan-plan`、`quarterly-rural-loan-plan` 及计划构建逻辑上接入真实 `pytools menu` → 7→7/8 入口。各步骤分别选择本期和上期 `.xlsx`；默认只生成 ExcelEditPlan，不改本期工作簿。只有明确提供 capability 与 platform 时才经 ExcelModifier 生成并校验候选，随后按规划快照安全替换本期工作簿。user/developer/off 视图及配置日志目录中的运行历史记录处理状态、机构/地区计数、阶段时间和实际引擎；历史不含路径、文件名或单元格内容。

Windows 原生 Rust 验证：`cargo fmt --manifest-path pytools-core/Cargo.toml -- --check` 通过；`batch_tools_menu_cli` 15/15、`rural_loan_regressions` 2/2 通过。新菜单测试覆盖 7-7 与 7-8 计划只读、user/developer/off 视图、地区错误摘要，以及临时 synthetic `umya_spreadsheet` capability 下真实菜单→ExcelModifier→安全提交；提交后检查工作簿值、公式、样式、合并区域、批注及未改地区，且上期输入保持不变。只读 PyCode 基线与 Windows native Rust CLI 的临时合成对拍中，7-7、7-8 正常摘要/输出投影以及 7-8 缺失上期地区、重复机构、缺统计地区等错误合同均一致；未生成机器报告文件。

这些证据仅覆盖合成 `.xlsx`，不登记生产 Passed capability，也不改变默认修改路由。真实业务模板/数据、Windows 原生文件弹窗手工操作、目标 Linux/glibc 2.28 与离线 deb 尚未验收；PT-7-7/PT-7-8 的产品验收状态仍为未验证。

### 2026-10-25 PT-6-2 Windows 原生 Rust 菜单与 PyCode 对拍

使用 Windows 原生 Rust 1.99.1 验证既有真实入口：`cargo test --manifest-path Cargo.toml -p pytools-core --test summary_comments_cli --test summary_menu_cli --offline --locked -j 1` 中 `summary_comments_cli` 4/4、`summary_menu_cli` 3/3 通过。CLI 测试覆盖 user/developer/off 三种视图和合并区域锚点语义；菜单测试验证 `pytools menu` → 6→2、配置输出/日志目录、运行历史及损坏输入错误。

另将只读 PyCode `run_summary_by_comment` 与本机编译的 `target/debug/pytools.exe summary-comments` 对拍三个合成场景：标准行/列区域（6 行）、重复行区域标记（6 行）和合并区域非锚点（3 行）。完整工作簿投影一致，比较值/类型、工作表名/维度、样式、冻结窗格、合并区域、公式、定义名称、打印区域、页面设置、页边距和居中属性。机器报告为 `probe/reports/summary_comments_windows_native_parity_20261025.json`。

本次只补充 PT-6-2 的 Windows 原生环境证据，没有修改业务代码。证据仍只覆盖合成输入；真实批注区域模板、Windows 原生文件弹窗手工操作、麒麟 V10/UOS V20、glibc 2.28 和离线 deb 未验证，PT-6-2 生产验收仍为未验证。

补充真实拒绝路径：以原项目只读 `PyCode/config/config.xlsx`（含批注但不含 `行区域`/`列区域` 标记）分别运行 PyCode `run_summary_by_comment` 与 Rust `summary-comments` CLI；两端错误文本逐字一致，均以失败码 2 终止且未生成结果工作簿，输入簿大小/修改时间保持不变。Rust 提示已对齐 Python，并由 `summary_menu_cli` 集成测试覆盖。该证据只验证无有效区域标记时的拒绝行为；合法真实区域模板、PT-6-2 正向业务验收和目标平台仍未验证。

### 2026-10-25 PT-6-2 空白与公式模板表头对拍

对拍发现并修复 `summary-comments` 的模板表头语义差异：模板空白表头现按 PyCode 回退为列字母；模板公式文本由 Calamine 公式投影读取，并在输出簿中作为公式表头保留，公式缓存留空，不执行公式计算。布尔表头按 Python `str(bool)` 大小写输出，数值表头保留 Python 的整数/小数字符串表现；日期表头通过共享 Calamine ISO 日期投影解析，再按 Python datetime 字符串语义输出。

新增 Rust CLI 回归，覆盖公式、空白、布尔、数值和日期表头、输出公式类型及输入不变；`probe/compare_summary_comments_template_header.py` 使用只读 PyCode 基线与 Windows 原生 `pytools summary-comments` CLI 对同一合成模板/源簿执行完整工作簿投影对拍，结果精确一致（3 行、1 个输出；4×8 输出维度；probe 模板使用 1904 日期系统；Calamine/rust_xlsxwriter，源未改变）。报告：`probe/reports/summary_comments_template_header_parity_20261025.json`。

Windows 原生验证：`cargo fmt --manifest-path pytools-core/Cargo.toml -- --check` 通过；`cargo build --offline --locked -q -p pytools-core --bin pytools` 通过；`cargo test --offline --locked -p pytools-core --lib comment_region_tests` 3/3；`--test summary_comments_cli --test summary_menu_cli -j 1` 5/5 与 4/4；`py -3.12 pytools-core/probe/compare_summary_comments_template_header.py` 完整投影 exact。证据限合成模板，不代替真实有效批注区域模板、麒麟 V10/UOS V20、glibc 2.28 或离线 deb 验收；PT-6-2 仍未通过生产验收。

另补数值表头 Python float 格式边界对拍：`1e-5`、`1e20`、`1e-4` 经 `summary-comments` CLI 回归与只读 PyCode 完整工作簿投影比较一致，输出 3 行、4×11，`different_fields` 为空。该证据限合成 XLSX，不扩大到所有 Excel 数值词法形式。

后续持续时间表头最小复现发现共享 reader 类型信息缺口：合成 XLSX 的 `[h]:mm:ss.000` 模板单元格由只读 PyCode/openpyxl 读取为 timedelta，表头文本例如 `1 day, 2:03:04.123000`；当前 Rust 日期投影把相同序列值格式化成 1900 日历日期。问题涉及共享 `excel-reader` 对日历日期和持续时间的区分，已向 main 提供复现和最小修改需求；PyTools 暂不自行解析 OOXML 样式绕过 reader，持续时间表头仍未验证。

### 2026-10-25 PT-5-1/PT-5-2 配置工具交互菜单接入

主菜单新增类别 5，并按只读 PyCode `main.py` 的 5-1/5-2 入口接入现有 Rust 配置能力。配置缺失或无法加载时仍可打开菜单；其他依赖配置的业务类别会提示先处理配置。5-1 在缺少配置时通过 Rust `initialize_new_config` 新建 11 张表；配置已存在时生成修复计划并保留原文件。只有显式提供匹配 capability/platform 才经 ExcelModifier 生成并校验候选，用户再输入 `y` 后才按原文件规划快照安全替换，保留备份。5-2 每次从磁盘读取，用户视图展示工作表维度、全局路径、启用规则/映射/提取任务摘要，开发者视图附原始配置结构及实际读取组件；打印阶段用时。没有已验收的生产修复 capability，因此菜单不默认修改现有配置。

Windows 原生 `cargo fmt --manifest-path pytools-core/Cargo.toml -- --check` 通过；主菜单回归 31/31：`config_menu_cli` 2/2、`timeline_menu_cli` 3/3、`dedup_menu_cli` 1/1、`print_menu_cli` 1/1、`conversion_menu_cli` 1/1、`summary_menu_cli` 3/3、`batch_tools_menu_cli` 15/15、`fx_header_menu_cli` 4/4、`archive_menu_cli` 1/1。真实菜单覆盖缺失配置提示、5-1 初始化、5-2 磁盘重读以及既有配置的计划模式；新配置为 11 张工作表，计划模式的修复计划非空、候选未生成且输入配置字节不变。配置模板和值、表头批注与只读真实配置的 Python/Rust 对拍沿用 `probe/reports/config_admin_parity_20261023.json`；修复候选的 openpyxl worker 语义及保留项沿用 `tests/config_repair_execution.rs` 的既有候选验证。本轮未重新运行 Python 基线，也未用 capability 执行菜单中的候选提交。

仍未完成真实用户配置文件的修改验收、生产 capability 登记、Windows 桌面交互、目标平台和离线 deb 验收；PT-5-1 修复生产验收继续为未验证。配置查看和菜单集成测试使用合成/本地配置，不写真实业务数据或运行历史。

### 2026-10-25 PT-4-9/PT-4-10 文件排序双视图运行日志

`pytools file-sort-prefix` 与 `pytools remove-file-sort-prefix` 现在支持 `--log-view user|developer|off`，交互菜单类别 4 的第 9/10 项沿用全局视图设置。Rust 业务回调记录配置读取、输出目录准备、逐文件处理、报表写入和任务摘要；开发者事件包含实际组件/引擎、阶段耗时与累计耗时。PT-4-10 按是否读取内容前缀工作簿标记 Calamine，纯文件改名标记 RustFilesystem；报表写入标记 rust_xlsxwriter。用户视图显示处理数量和阶段耗时，off 不输出进度或新建日志，开发者视图保存结构化历史并保留原详细操作日志。两项操作的文件名、统计及报表业务语义未改变。

Windows 原生验证：`cargo fmt --manifest-path pytools-core/Cargo.toml` 通过；`cargo test --offline --locked -p pytools-core --test conversion_menu_cli --test file_sort_prefix_operation_parity --test file_sort_prefix_remove_parity -q` 中 2/2、3/3、6/6 通过；`pytools --help` 显示两条命令的三种视图选项。对拍使用既有只读 PyCode 机器基线，4-9 八个文件案例及 4-10 内容前缀/冲突/跳过案例的统计、目录结果和 XLSX 报表检查无业务差异；新增菜单三视图测试、Calamine 实际调用标记及输出目录失败阶段测试。结构化历史不含输入文件名或路径。没有重新运行 Python 解释器，也没有可用真实业务目录样本；手工桌面文件选择器、目标 Linux/glibc 2.28、麒麟/UOS 与 deb 未验证，PT-4-9/PT-4-10 整体迁移状态仍未验证。日志性能开销未测量。

### 2026-10-25 PT-4-2/PT-4-3 转换批处理双视图进度日志

Excel/Word 批量转换 Rust API 新增可选进度回调；直接 `convert-excel`/`convert-word` CLI 和 `pytools menu` → 4→2/3 现支持 `--log-view user|developer|off` 并沿用全局日志目录。user 视图提供发现数量、逐文件进度、阶段耗时和脱敏错误摘要；developer 视图输出事件流并保存 JSON 历史，记录实际组件/引擎、计数、阶段和总耗时；off 不输出事件或写入历史。历史不包含输入路径、文件名或错误明细。same-format 文件复制标记 `Rust std::fs::copy`；native CSV 标记 `csv + rust_xlsxwriter`；native Excel 标记 Calamine 与 rust_xlsxwriter；sidecar 成功时从其结果读取真实 LibreOffice/COM 引擎。转换业务路由和输出行为未改变。

Windows 原生验证：`cargo fmt --manifest-path pytools-core/Cargo.toml -- --check` 通过；`cargo test --offline --locked -p pytools-core --test conversion_menu_cli -j 1` 8/8 通过；`cargo test --offline --locked -p pytools-core --lib conversion::tests -j 1` 6/6 通过。新增真实菜单测试覆盖 user/developer/off、直接 native CSV CLI、实际引擎与耗时字段、历史脱敏、缺失 sidecar 错误及源文件不变。native CSV 输出通过既有 Rust 候选自校验。Windows 未发现已安装 LibreOffice，故无成功 `.xls`/`.et`/`.doc` sidecar 样本，未作该类 Python/Rust 结果对拍；PT-4-2/PT-4-3、手工桌面文件选择器、目标平台和 deb 仍未验收。

### 2026-10-25 PT-7-5 地区总分校验双视图运行事件

既有 pytools region-sum-check CLI 增加可选 --log-view user|developer|off，主菜单 7→5 将全局视图传入真实 CLI。Rust 地区核对核心提供阶段进度事件，记录处理计数、失败类型、阶段/总耗时及实际调用的 Calamine、rust_xlsxwriter；产生核对结果行时同时标记 ooxml-patch。user 输出简明进度，developer 输出结构化事件并保存 JSON 运行历史，off 不输出结构化事件或新增历史。未指定 --log-view 时保留原 CLI JSON 输出；既有详细 .log 保持兼容。新运行历史不保存输入路径、文件名或单元格值，核对业务与 Excel 引擎路由未改。

Windows 原生 Rust 验证：cargo test --offline --locked -p pytools-core --test region_sum_parity --test batch_tools_menu_cli -j 1 中 5/5 和 16/16 通过；cargo fmt --manifest-path pytools-core/Cargo.toml -- --check、git diff --check 通过。测试启动真实 pytools menu → 7→5 → region-sum-check，覆盖 user/developer/off 视图、成功/损坏输入、历史隐私字段、实际引擎与阶段耗时；PT-7-5 既有合成工作簿输出继续与只读 PyCode 保存基线对拍，值和展示断言无回归。未重新运行 Python 解释器或全量 pytools-core 测试。

合成夹具之外的真实业务工作簿、Windows 桌面文件弹窗手工操作、麒麟 V10/UOS V20、glibc 2.28 和离线 deb 未验证；PT-7-5 产品验收仍为未验证。日志性能开销未测量。

### 2026-10-25 PT-9 快捷工具主菜单接入

Rust 交互主菜单现提供与 PyCode 对应的“9 快捷工具”入口，包含打开配置、打开 output 最新文件、打开 output 文件夹三项。各项复用既有 `quick` 业务实现；完成、未找到文件和启动/建目录失败时显示对应结果，并按旧菜单行为返回主菜单。直接 CLI 行为、文件选择规则及系统启动器未更改。

Windows 原生验证：`cargo fmt --manifest-path pytools-core/Cargo.toml -- --check` 通过；`cargo test --offline --locked -p pytools-core --test quick_menu_cli --test archive_menu_cli --test config_menu_cli -j 1` 共 4/4 通过。新集成测试通过真实 `pytools menu` 进程输入覆盖主菜单 9→子菜单 1/2/3，使用无效配置路径、空输出目录及被文件阻塞的输出父目录，断言各分支提示并返回主菜单；不会启动桌面程序。菜单标签与 `PyCode/main.py` 的快捷菜单及 `q1/q2/q3` dispatcher 对齐。既有直接 CLI 假 `xdg-open` 对拍报告为 `probe/reports/quick_tools_parity_20261023.json`，本轮未重复运行 Python 对拍。

本轮未直接启动 Windows Explorer/其他桌面应用；真实配置和业务 output 目录、旧 Windows `os.startfile`、麒麟 V10/UOS V20 桌面集成、glibc 2.28 与离线 deb 尚未验收，PT-9 整体仍未验证。

### 2026-10-25 PT-4-4 菜单 capability 确认执行

类别 4→4 现在接收主菜单启动时传入的 `--capabilities` 与 `--platform`。缺少这两个参数时保持 plan-only；两者都提供时，菜单先明确提示将原位修改所选 `.xlsx`，只有用户输入 `y/yes/是` 才调用现有 `rename-sheets` 候选执行路径，否则保留原工作簿并只生成计划。执行仍通过 ExcelModifier capability 校验、临时候选验证、输入快照核验和安全替换，不登记或改变生产 capability。

Windows 原生验证：`cargo fmt --manifest-path pytools-core/Cargo.toml -- --check`；`cargo test --offline --locked -p pytools-core --test conversion_menu_cli --test sheet_rename_execution -j 1` 共 7/7 通过。菜单测试覆盖不确认时源字节不变并只生成计划，以及 synthetic Passed capability + 明确确认后从真实 `pytools menu` 入口改名并读回工作表和值；同一测试组还确认非 Passed capability 在源写入前失败。`sheet_rename_execution` 的既有 Python baseline 对拍也随本轮重跑通过，核对工作表顺序、值/公式、定义名称、样式、批注、合并、行列尺寸、页面设置及有效提交。

上述修改执行仅使用合成工作簿和测试专用 capability。没有授权的真实业务簿或生产 capability；Windows 文件弹窗手工操作、XLSM 保留、麒麟 V10/UOS V20、glibc 2.28 和离线 deb 仍未验证，PT-4-4 整体仍为未验证。

### 2026-10-25 PT-3-1/PT-3-2 打印菜单 capability 确认执行

主菜单 3→打印 1/2 现在接收已配对的 `--capabilities` 与 `--platform`。不提供时保持 plan-only；提供两项时先提示生成新的打印工作簿且源簿不变，只有输入 `y/yes/是` 才把两项传给既有 `print-comment-plan` CLI，经 ExcelModifier 按 capability 校验后执行。拒绝确认仍只保留计划。未登记生产 Passed capability，也未改变默认路由。

Windows 原生验证：`cargo fmt --manifest-path pytools-core/Cargo.toml -- --check` 通过；`cargo test --offline --locked -p pytools-core --test print_menu_cli --test print_comment_execution -j 1` 中打印菜单测试 3/3、候选执行测试 1/1 通过。菜单测试由真实 CLI 进程验证 PT-3-1/3-2 确认执行、拒绝后的 plan-only、输出路径、候选工作表、批注/零值语义及源簿不变。使用 PyTools Python 3.12.10/openpyxl 3.1.5 与 Windows 原生 `target/debug/pytools.exe` 新跑 `probe/compare_print_comment_plans.py`，keep/fast 两种计划与 ExcelModifier 候选的完整投影均通过；机器报告为 `probe/reports/print_comment_plan_parity_20261025.json`、`probe/reports/print_comment_execution_parity_20261025.json`。对拍样本为合成 XLSX，源未改变，执行 capability 仅临时测试使用。

本轮只用合成 XLSX 与测试专用 capability。生产 capability、真实打印模板、Windows 桌面弹窗手工操作、麒麟 V10/UOS V20、glibc 2.28 和离线 deb 未验收；PT-3-1/PT-3-2 整体仍为未验证。
### 2026-10-25 PT-6-1 Windows Rust CLI 真实 XLSX 样本复验

在 Windows Python 3.12.10/openpyxl 3.1.5 环境下，直接运行现成的 `compare_real_used_range_sample()`，以原 PyTools 只读 `xlsx/` 目录的 13 份真实 XLSX 与 Windows 上运行的 `target/debug/pytools.exe` 按四批对拍。96 张工作表、203,580 行的完整输出投影均一致；行数/文件数和开发者事件流/保存历史断言通过，源文件前后指纹一致，输出只在临时目录。机器报告：`probe/reports/summary_used_range_real_parity_20261025.json`。

完整组合 `compare_summary_tools.py` 未能完成，因为当前 PyTools Python venv 缺少 NoComPyTools 对拍所需的 `python_calamine`；本次已隔离运行 PyCode/Rust PT-6-1 主样本对拍，不安装依赖、不修改源项目。此结果补充 Windows 上运行现有 Rust CLI 的阶段证据，不代表 Windows 原生 Cargo 构建、原 PyTools 业务 VBA XLSM、麒麟 V10/UOS V20、glibc 2.28、deb 或离线交付验收；PT-6-1 整体仍未完成最终验收。

### 2026-10-25 PT-6-1 Windows 原生 MSVC 真实样本复验

使用 Windows 原生 `x86_64-pc-windows-msvc`（Cargo/rustc 1.99.1）构建当前分支 `pytools` CLI；`summary_menu_cli` 与 `summary_comments_cli` 回归 7/7 通过。以该 MSVC 二进制直接运行现有 `compare_real_used_range_sample()`：原 PyTools 只读 `xlsx/` 目录 13 份真实 XLSX，按四批共 96 个工作表、203,580 行与 PyCode 对拍；值/类型、非空单元格样式、工作表名称与维度、冻结窗格/表头粗体、公式/合并/定义名称/页面属性均精确一致，输入前后指纹相同，输出仅写入临时目录。机器报告仍为 `probe/reports/summary_used_range_real_parity_20261025.json`，现标出 Windows MSVC 执行环境。

另经 Windows 原生 `pytools menu → 6 → 1` 与原生文件对话框选择 1 个合成错误值矩阵 XLSX；PyCode/Rust 完整输出投影 exact（2×18），8 个扩展错误 token 按文本兼容，7 个标准错误值由 OOXML Patch 保留。14 条运行历史事件记录了 Calamine、窄化错误 token 兼容、OOXML Patch 和 `rust_xlsxwriter`；输入夹具未变，输出/历史仅在 TEMP。机器证据并入 `probe/reports/summary_used_range_real_menu_parity_20261025.json`。该项仅补菜单入口的合成错误值矩阵，不扩大 PT-6-1 生产验收。

本轮没有修改 PT-6-1 业务逻辑或默认引擎；这补齐的是 Windows 原生 Cargo 阶段验证，不覆盖 VBA-bearing 原业务 XLSM、NoComPyTools 独有格式、麒麟 V10/UOS V20、glibc 2.28 或离线 deb，PT-6-1 整体仍为未完成最终验收。完整 NoCom 对拍仍因 Python venv 缺少 `python_calamine` 未执行。

### 2026-10-25 PT-7-6 显式追加既有输出簿

新增 pytools extract-sheets --append-existing，映射原 PyCode run_extract_sheets(..., append_existing=True) 的显式追加参数；现已接入真实交互菜单 pytools menu → 7 → 6，并在选源前询问是否追加到配置指定的已有输出簿，默认保持新建输出行为。用户确认追加但未提供 --capabilities 与 --platform 时，菜单在选源前停止。对已存在的目标，仅支持 XLSX；读取并验证目标 workbook inventory，保存目标 SHA-256 快照，要求匹配的 --capabilities 与 --platform。Rust 先以 rust_xlsxwriter 生成待追加工作表，再经 ExcelModifier 按 CreateSheet + ImportRange 对既有簿执行整批修改；Calamine 校验新表，最终仅在目标快照未变化时安全替换。缺少/不匹配 capability、目标 inventory 不完整或输出冲突时失败关闭，目标不被覆盖。目标路径尚不存在时，沿用新建输出行为。

Python 只读基线测试 test_extract_sheets_append_existing_output_keeps_previous_sheets（区域提取）通过 1/1。Rust 真实 CLI 合成回归验证首次新建后再次追加，表名/顺序和值正确；无 capability 与平台不匹配时原目标字节不变；运行历史包含实际引擎/提交阶段且不含单元格值。现有交互菜单以明确确认接入该追加模式；新增端到端测试验证无 capability 时选源前拒绝，以及临时匹配 capability 下通过 pytools menu → 7 → 6 → pytools extract-sheets --append-existing 执行，既有值/公式/批注/合并保留，新表值正确，运行历史不含单元格值。整表追加另验证既有值、公式、批注、合并、粗体样式和行列尺寸保留；新表缓存值、合并、样式、行高/列宽复制，普通无批注版不复制公式文本与批注。Windows 原生 `sheet_extract_area_operation_parity` 11/11、`batch_tools_menu_cli` 16/16、cargo fmt --check 通过。完整机器报告：probe/reports/sheet_extract_append_existing_parity_20261025.json。

本次使用的 Umya CreateSheet + ImportRange capability 仅由测试临时生成，未登记生产 Passed capability。Python 既有测试只覆盖区域追加；整表 Python helper 生成文件的样式重开问题仍按前节记录，本次不宣称整表输出文件级 Python 对拍通过。真实模板、未列举 OOXML 对象、非 XLSX、Windows 原生文件弹窗、目标 Linux/glibc 2.28 与 deb 未验证，PT-7-6 整体验收仍未完成。

### 2026-10-25 PT-1-2 时序提取 Rust 主输出、目标追加与入口对拍

新增 `timeline_extract` Rust 业务模块，按原 `PyCode.core.timeline_396.run_timeline_slim` 复用现有规则引擎，读取配置规则与路径映射，执行行头后缀、跨规则去重、日期和值写出，并以 `rust_xlsxwriter` 生成七列 `时序提取结果`。新增 `timeline-extract` CLI 及真实菜单 `pytools menu → 1 → 2`，接入 user/developer/off 日志视图与 PT-1-2 运行历史。读取、写入和目标提交事件记录实际引擎、进度与阶段耗时。

目标写入按原 Python `append_to_target` 语义迁移：目标不存在时以 Rust writer 创建；已有目标默认只生成计划，不修改原文件；需显式 capability/platform 与提交确认后，经 ExcelEditPlan/ExcelModifier 执行 XLSX 修改。按既有目标列头检查、去重键 `[0,1,3,4,5]`、禁用扩列；表头不匹配时跳过并保留旧目标。现有目标测试使用临时合成 capability `umya_spreadsheet`，没有登记生产 Passed capability，也没有改变默认路由。

Windows 原生构建与回归通过：`cargo fmt --manifest-path pytools-core/Cargo.toml -- --check`、`cargo build --offline --locked -p pytools-core --bin pytools -j 1`、`cargo test --offline --locked -p pytools-core --test timeline_menu_cli -j 1`（9/9）、`cargo test --offline --locked -p pytools-core --lib -j 1`（33/33）。只读 PyCode/Rust 合成对拍通过：2 个有效 XLSX、8 个规则命中工作表、20 条汇总输出；新目标与 PyCode 完整工作簿投影一致；已有目标保留 1 条重复并追加 19 条；表头不匹配时 PyCode 与 Rust 均不修改目标，汇总投影一致；损坏文件的 continue/fail_fast 行为一致。普通用户视图输出进度且在 stdout/stderr 隐藏引擎名，输出投影一致；持久化历史仍记录实际引擎。未修改源工作簿。

Windows 原生菜单 `1→2` 另以系统 Open 对话框选取 1 个临时合成工作簿，命中 4 个工作表、提取 20 条记录、写出 10 行；与 PyCode 输出的完整工作簿投影一致。机器对拍报告：`probe/reports/timeline_extract_parity_20261025.json`；可重跑自动探针：`probe/compare_timeline_extract.py`。

新增无位置参数入口 `pytools timeline-extract --config <配置簿>`：按 PyCode `source_paths=None` 语义递归扫描配置输入目录，按 `source_exts` 匹配扩展名、跳过 `~$` Excel 锁文件并排序；显式源文件列表与菜单选取不变。Windows 原生集成测试覆盖递归子目录、扩展名大小写、锁文件、未配置格式和空目录。合成 Python/Rust 对拍中默认扫描 4 个已配置扩展名文件（3 个有效、1 个损坏），锁文件排除，输出完整工作簿投影一致。报告已更新至 `probe/reports/timeline_extract_parity_20261025.json`。本项只证明合成样本行为，真实业务样本及目标 Linux/deb 验收仍未完成，PT-1-2 整体保持未验证。

真实业务样本尚不可用：原 PyCode 默认配置的输入目录不存在，源项目树中的工作簿文件没有命中该配置的时序规则。现有证据仅为合成 XLSX；PT-1-2 仍未通过真实数据验收。D01 将 PT-1-2 排除在已确认目标写入合同之外，因此既有目标写入只作为 capability-gated 候选，不能作为生产能力或迁移完成依据。其他格式、麒麟 V10/UOS V20、glibc 2.28 与 deb 离线交付未验证；PT-1-2 整体状态仍为未验证。

### 2026-10-25 PT-1-3/PT-1-4 Windows 原生文件对话框实测

Windows 原生 `pytools.exe menu --config ... --log-view developer` 分别经 `pytools menu → 1 → 3`（行头加后缀）和 `→ 1 → 4`（行头不加后缀）打开系统文件选择对话框。每种模式选择一份隔离的合成 XLSX，确认在临时目录新建目标；只读原 PyCode 基线与 Rust 汇总、追加候选及提交后目标均通过现有完整投影/有效样式对拍。

- PT-1-3：汇总 3 行（4×9），新目标追加 3 行、过滤重复 0 行；PyCode/Rust 投影精确一致，开发者历史记录 Calamine、`rust_xlsxwriter` 和目标执行 40 ms。
- PT-1-4：汇总 3 行（4×9），新目标追加 2 行、过滤重复 1 行；PyCode/Rust 投影精确一致，开发者历史记录 Calamine、`rust_xlsxwriter` 和目标执行 42 ms。
- PT-1-3 原生对话框多选：一次选择 2 份源簿，汇总 6 行（7×9），新目标追加 6 行、过滤重复 0 行；汇总、候选及提交目标与 PyCode 精确一致，记录 2 次 Calamine 读取、`rust_xlsxwriter` 写入和 43 ms 目标执行；源簿与配置逐字节未变。
- 两种模式的源簿语义投影及配置单元格值均未改变；目标仅为隔离临时目录中的新文件。Windows 原生 `cargo test --offline --locked -p pytools-core --test timeline_menu_cli` 9/9 通过。机器证据追加至 `probe/reports/wide_summary_parity_20261023.json`。

该验证只补齐合成数据上的原生文件对话框和菜单入口证据，不代表脱敏真实业务样本、真实模板全量保真、既有目标生产 capability、麒麟 V10/UOS V20、glibc 2.28 或离线 deb 验收；PT-1-3/1-4 整体仍为**未验证**。

### 2026-10-25 PT-6-2 输出工作表名称与 PyCode 对齐

用合成模板复现并修正两项页签差异：`  Branch_A  ` 输出时去除首尾空白；`  Region `/`region` 因 Excel 页签名大小写不敏感而冲突时，Rust 与 PyCode/openpyxl 一样生成 `Region`、`region1`，不再保存失败。`summary-comments` 现对输出页签执行与源实现一致的首尾空白清理、非法字符替换、31 字符截断、空名回退、`_2` 重名后缀及 openpyxl 大小写冲突数字后缀；源工作表匹配和值仍使用原页签名。

Windows 原生验证：`cargo fmt --manifest-path pytools-core/Cargo.toml -- --check` 通过；`cargo test --offline --locked -p pytools-core --lib comment_region_tests -j 1` 4/4、`cargo test --offline --locked -p pytools-core --test summary_comments_cli -j 1` 5/5 通过。只读 PyCode 与真实 `pytools summary-comments` CLI 的完整工作簿投影对拍通过，6 行数据、两个输出页签 `Branch_A`（4×11）和 `branch_a1`（4×5），值/类型、样式、页面属性均无差异；合成输入未改变。机器报告：`probe/reports/summary_comments_template_header_parity_20261025.json`。

该修复不改变 PT-6-2 的数据读取、规则或引擎路由；仍只有合成标记模板，没有找到有效真实业务模板。麒麟 V10/UOS V20、glibc 2.28、deb 与离线交付仍未验收，PT-6-2 整体状态保持未验证。

### 2026-10-25 PT-1-3/PT-1-4 宽表 CLI 默认输入目录发现

`wide-summary` 与 `wide-summary-row-suffix` 现可省略位置源文件参数，按配置输入目录递归发现匹配 `source_exts` 的文件；扩展名大小写不敏感，忽略 Excel `~$` 锁文件并按路径排序。显式源文件参数行为保持不变。扫描逻辑抽到 `source_discovery::discover_configured_source_files`，并由现有 `timeline-extract` 默认扫描入口复用；`discover_timeline_extract_sources` 保留为旧 Rust 调用方的兼容委托。启用 user/developer 日志时记录发现阶段和耗时，developer 历史包含发现数量，后续事件序号保持连续。

Windows 原生 `cargo fmt --manifest-path pytools-core/Cargo.toml -- --check`、`cargo build --offline --locked -p pytools-core --bin pytools -j 1` 通过；`cargo test --offline --locked -p pytools-core --test timeline_menu_cli -j 1` 12/12 通过。只读 PyCode/Rust 自动对拍脚本新增两种行头模式的无位置参数案例：各扫描根目录与嵌套目录内的 2 份 `.xlsx`，跳过 `~$` 文件和未配置 `.xlsb`，汇总工作簿完整投影一致，开发者历史计数和发现事件正确，源/锁文件/配置未变。报告：`probe/reports/wide_summary_parity_20261023.json`；重跑过程中保留原有 3 条原生文件对话框证据。

原 PyCode 配置与其源树工作簿的只读预检没有发现可产生宽表输出的真实输入；本次仅能确认合成样本行为。生产 capability 未登记，真实业务匹配样本、麒麟 V10/UOS V20、glibc 2.28 与离线 deb 未验证，PT-1-3/PT-1-4 整体仍未验收。
