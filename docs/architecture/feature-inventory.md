# 功能清单（feature-inventory）

> 审计日期：2026-10-23。审计人：Rust 架构审计（自动化 + 人工交叉验证）。
> 审计对象：`1.rust-workspace`（迁移目标）、`3.open-data-audit`（只读源）、`4.pytools-main`（只读源）。
> 本轮为只读审计：未修改任何业务代码、依赖与 workspace 结构。所有结论以审计时刻的代码状态为准。
> 关联文档：[feature-dataflow-matrix.md](feature-dataflow-matrix.md)、[migration-gap-analysis.md](migration-gap-analysis.md)、[excel-engine-routing-review.md](excel-engine-routing-review.md)。四份文档使用相同功能编号。

## 0. 审计口径

### 0.1 功能判定口径

- 功能 = 用户可独立触发的业务操作，从源项目 GUI 菜单/按钮/dispatch 分支、CLI 参数或 Web bridge 方法出发确认，**不以文件名或函数名推断**。每项功能追踪到实际打开/写出文件的调用行。
- pytools 以 PyCode 菜单编号为规范编号（PT-4-6 等与 workspace 既有文档一致）；NoComPyTools 菜单序差异逐项标注。PyCode 为行为基线（`pytools-core/MIGRATION_PLAN.md:5`）。
- ODA 以源项目 `core/tests/TEST_GROUPS.md` 与 `audit-core/MIGRATION_INVENTORY.md` 的 S1/S2/S3/S5 功能号为规范编号（AUD-S1-F01 等）；另登记源项目中实际存在的配套功能（AUD-C01..C07）。
- pytools 变体处理：工作表提取 COM 批注版**不迁移**（用户已确认，见 `AGENTS.md` §8）；打印 PDF（PT-3-5）按 workspace 范围排除；Python 模式切换（PT-4-11）在当前 PyCode 代码中已不存在（grep `PYTOOLS_MODE` 0 命中）；`survey_stats`（PT-SV1）在源项目中实现存在但菜单未接线，登记待确认，不静默遗漏。

### 0.2 迁移状态分类（六类，全文统一）

| 标签 | 定义 |
|---|---|
| **尚未迁移** | Rust workspace 无对应业务实现。 |
| **已迁移部分核心逻辑** | Rust 已实现并对拍该功能的可辨识切片（纯逻辑 helper、局部流程），但完整功能链未迁移。 |
| **代码未接入口** | 已有 Rust 代码，但任何产品入口（CLI 子命令 / web serve）都不可达，仅测试或 probe 使用。 |
| **接入依赖 Python** | 已接入入口，但业务处理仍由 Python 执行（本轮审计中 Rust 产品路径为 0 项；反向情形见 §3.4）。 |
| **接入未验** | 已从迁移产品入口（Rust CLI / `open-data-audit serve`）可达，但未完成结果一致性验证（无对拍、仅纯逻辑向量对拍、或仅合成样本对拍——真实模板/真实目录/目标平台均未验收）。 |
| **完整迁移** | 已接入且通过真实样本一致性验证与目标平台验收。**本轮审计：0 项。** |

> 注意：pytools 各 CLI 均为"迁移产品的新入口"，不等于已替换原 Python 菜单；原产品仍完整保留。某项标"接入未验"仅表示 Rust 产品入口可达，不代表生产默认已切换（workspace 规则 §6.2/§9.3 亦如此要求）。

### 0.3 证据标记

- `PyCode/...:行号`、`NoComPyTools/...:行号` 指源项目 `C:\Users\AZI\Desktop\VibeCode\4.pytools-main` 下文件。
- `core/src/...:行号` 指源项目 `C:\Users\AZI\Desktop\VibeCode\3.open-data-audit` 下文件。
- `pytools-core/src/...`、`audit-core/src/...`、`crates/...` 指 workspace 文件。
- 无法确认的调用链明确标 **待验证**。

## 1. 总量统计

| 维度 | 数量 |
|---|---|
| 登记功能总数 | **71** |
| pytools（PT） | 48（47 项菜单功能含 2 项范围外 + 1 项源项目未接线） |
| ODA（AUD） | 23（15 项核心系统功能 + 1 项 S4 待确认 + 7 项配套功能） |
| 已完整迁移并通过验证（S6） | **0** |
| 接入未验（S5） | 45（pytools 44 + ODA 1） |
| 已迁移部分核心逻辑（S2） | 3（PT-1-2、AUD-S1-F01、AUD-S3-F01） |
| 尚未迁移（S1） | 19（全部为 ODA 功能） |
| 范围外/源已移除/待确认 | 4（PT-3-5、PT-4-11、PT-SV1、AUD-S4） |
| 接入依赖 Python（S4 类） | 0（Rust 产品路径无 Python 依赖；反向依赖见 §3.4） |

pytools 44 项"接入未验"中，有合成端到端对拍证据的约 15 项（4-1、4-6..4-10、7-2 计划级、7-4、7-5、7-6 area 切片、8-1..8-5），其余为仅 CLI 接线或纯逻辑向量对拍（详见 [migration-gap-analysis.md](migration-gap-analysis.md) §2）。

## 2. pytools 功能清单（48 项）

状态列：接入未验=●；部分核心=◐；尚未迁移=○；范围外/移除/待确认=△。

### 2.1 时序工具（菜单 1）

| 编号 | 功能 | 用户入口 | 输入 → 输出 | 状态 | Rust 入口/证据 |
|---|---|---|---|---|---|
| PT-1-1 | 表头路径比对 | NoCom `main.py:875-879`；PyCode `main.py:1657-1661`（dispatch `"1com"` @725-751，.xls 自动路由 COM `_split_sources_for_auto_route` @490-499） | 模板 xlsx/xlsm(+.xls) + 多源 → 新建 xlsx（2 sheet 报告） | ● | `pytools timeline-compare`（`pytools-core/src/bin/pytools.rs:72`） |
| PT-1-2 | 时序提取（快速精简） | NoCom `main.py:477-491`；PyCode `"2com"` @752-779 | 多源 → 新建 xlsx 或**追加已有簿**（`NoComPyTools/nocompytools/core/io_excel.py:278-284`，`mode="a", if_sheet_exists="replace"`） | ◐ | 提取规则/工作簿抽取 helper + `pytools extract-sheets-area`（`pytools.rs:108`）；完整功能（整表样式、追加）在用户范围排除（`MIGRATION_PLAN.md:20`） |
| PT-1-3 | 宽表汇总（行头加后缀） | NoCom `main.py:492-509`；PyCode `"3com"` @780-816 | 多源 → 新建 xlsx 每规则一 sheet（`wide_397_398.py:322`），可扩列追加 | ● | `pytools wide-summary-row-suffix`（`pytools.rs:74`） |
| PT-1-4 | 宽表汇总（行头不加后缀） | 同 PT-1-3 入口组 | 同上 | ● | `pytools wide-summary`（`pytools.rs:73`） |
| PT-1-5 | 预校验 | NoCom `main.py:457-459`；PyCode @701-703 | config.xlsx 只读校验（`precheck_399.py:24-50`，pandas）→ 目录 + 控制台报告，不写工作簿 | ● | `pytools timeline-precheck`（`pytools.rs:71`） |

### 2.2 去重工具（菜单 2，两变体同代码 `dedup_311.py`，无 COM 版）

| 编号 | 功能 | 用户入口 | 输入 → 输出 | 状态 | Rust 入口/证据 |
|---|---|---|---|---|---|
| PT-2-1 | 按批注检查重复 | NoCom `main.py:510-516`；PyCode `main.py:850-856` | 单工作簿 → **原位标红**（`dedup_311.py:17,143` 写 PatternFill 后 `wb.save` @185）+ 控制台统计 | ● | `pytools dedup-comment-plan`（`pytools.rs:88`，计划输出；执行需能力登记，见 gap 文档 §4） |
| PT-2-2 | 按批注删除重复 | NoCom `main.py:517-523`；PyCode @857-863 | 倒序 `delete_rows`（`dedup_311.py:160-161`）后保存 @210；**不修正合并区/批注位置**（现状） | ● | 同上（dedup-comment-plan） |
| PT-2-3 | 按配置检查重复 | NoCom `main.py:524-539`；PyCode @864-889（dispatch `"8"`） | 配置分组，同簿只开一次（`run_check_by_config` `dedup_311.py:215-250`，变更时保存 @249） | ● | `pytools dedup-config-plan`（`pytools.rs:89`） |
| PT-2-4 | 按配置删除重复 | 同上（dispatch `"9"`） | `run_delete_by_config` `dedup_311.py:259+` | ● | 同上 |
| PT-2-5 | 按配置去重追加到目标 | 同上（dispatch `"a"`） | `_append_unique_source_to_target` `dedup_311.py:368`（值拷贝；`data_only=False` 会搬运公式原文） | ● | `pytools dedup-append-plan` / `dedup-append-execute`（`pytools.rs:91-92`；新目标走 writer，既有目标需能力证据） |
| PT-2-6 | 按配置预校验 | 同上（dispatch `"b"`） | 只读（`_open_wb(read_only=True)` `dedup_311.py:553`） | ● | `pytools dedup-precheck`（`pytools.rs:90`） |

去重配置中的"写入方式/覆盖方式"列被读入（`config_xlsx.py` NoCom:477/494）但全项目无消费点——死配置，迁移时不实现。

### 2.3 打印工具（菜单 3；全部只产出带页面设置的 xlsx，全项目无 PrintOut 调用）

| 编号 | 功能 | 用户入口 | 输入 → 输出 | 状态 | Rust 入口/证据 |
|---|---|---|---|---|---|
| PT-3-1 | 按批注打印（保留源格式） | NoCom `main.py:540-546`；PyCode `"p1com"` @934-959 | 源簿（openpyxl **双开**：样式视图 @`print_310.py:61-62` + 缓存值视图 @452-457）→ 新建 `{ts}_按批注打印（保留源格式）.xlsx`（@491），含值+全部样式+批注+行高列宽+合并（`_copy_block` mode=1 @244-288）+ 页面设置（`_apply_print_setup` @121-147） | ● | `pytools print-comment-plan`（`pytools.rs:94`，计划+可选能力执行） |
| PT-3-2 | 按批注打印（快速复制） | NoCom `"p3"` `main.py:547-553`；PyCode `"p3com"` @960-985 | 值=calamine（`_read_values_calamine` @201-203），样式仍 openpyxl；mode=3 复制值+基础样式+合并（@256-307），**不复制批注/保护/行高列宽** → 新建 @417 | ● | 同上 |
| PT-3-3 | 按配置打印（模式 1/2/3） | NoCom `"p7"` `main.py:554-566`；PyCode `"p7com"` @986-1004 | **修改已有目标簿**：清空目标 sheet（`delete_rows(1,max_row)` @740-742）后写入并 `wb.save` @797；目标不存在则新建 @727；四级缓存 @595-598；mode=2 自适应列宽行高 @309-364 | ● | `pytools print-config-plan`（`pytools.rs:96`） |
| PT-3-4 | 按配置打印预校验 | NoCom `"p8"` `main.py:567-570`；PyCode @930-933 | 只读（`_open_wb(read_only=True)` @574） | ● | `pytools print-config-precheck`（`pytools.rs:95`） |
| PT-3-5 | 按配置打印→导出 PDF（仅 PyCode/COM） | PyCode `"ppdf"` `main.py:1005-1026`；`print_310_com.py:813-819` `ExportAsFixedFormat` | → PDF 渲染 | △ 范围排除 | 无（workspace 规则 §1.1 排除 PDF） |

页面设置写入清单（迁移对拍基准）：orientation、pageSetUpPr.fitToPage、fitToWidth、fitToHeight（0=不限页高）、scale=None、paperSize=9(A4)、print_area、horizontalCentered、verticalCentered、oddFooter.center `"第 &P 页 / 共 &N 页"`（`NoComPyTools/.../print_310.py:121-147`，已实地核对）；另 `_apply_sheet_order_footer` @150-156。全项目未写 margins/printTitleRows/gridLines/headings/blackAndWhite/draft/firstPageNumber（与 workspace 规则 §1.2 一致）。

### 2.4 转换与文件命名（菜单 4）

| 编号 | 功能 | 用户入口 | 输入 → 输出 | 状态 | Rust 入口/证据 |
|---|---|---|---|---|---|
| PT-4-1 | 批量重命名（配置驱动） | NoCom `"t3"` `main.py:571-577`；PyCode @1027-1035 | 配置簿 + 文件/目录 → `os.rename`（`convert_33_34_35_37.py` NoCom:599-643），冲突自动加序号；不读工作簿内容 | ● | `pytools rename-batch`（`pytools.rs:115`）；合成对拍 + 真实配置窄 smoke（`probe/reports/batch_rename_parity.json`、`batch_rename_real_config_smoke.json`） |
| PT-4-2 | Excel 格式转换（xls 等→xlsx） | NoCom `"t4"` `main.py:578-590`→`xls_to_xlsx.py:54-128`（Win=Excel/WPS COM @27-51，Linux=LibreOffice @`libreoffice.py:32-58`）；PyCode `"t4"` @1036-1054→`run_excel_convert_34` @1025-1071（**COM-only**，8 种 FileFormat @986-988，.et 优先 WPS @1055） | .xls/.et/... → xlsx | ● | `pytools convert-excel/convert-file` → **sidecar 子进程**（`pytools-core/src/conversion.rs:155`）或 native 按值转换（`conversion.rs:120-137`）；真实 .et/.xls 样本未验证 |
| PT-4-3 | Word 格式转换（doc↔docx） | NoCom `"t5"` `main.py:591-600`→`word_convert.py:66-143`；PyCode @1055-1072（COM-only @1077-1089） | doc/docx 互转（同类） | ● | `pytools convert-word` → sidecar（同上）；真实 .doc 未验证 |
| PT-4-4 | 批量修改 Sheet 名 | NoCom `"t7"` `main.py:601-607`→convert @841-877（openpyxl `ws.title=` 后保存 @869/874，keep_vba @852）；PyCode `"t7com"` @1083-1107（.xls 走 COM @1596-1668） | xlsx/xlsm 原位改 sheet 名 | ● | `pytools rename-sheets-plan`（`pytools.rs:103`，计划输出） |
| PT-4-5 | 提交缺失检查 | NoCom `"t11"` `main.py:631-637`→`run_submission_check_48` NoCom @198-228 | 文件名关键词匹配 → pandas 报告 xlsx | ● | `pytools submission-check`（`pytools.rs:104`） |
| PT-4-6 | 按内容加前缀重命名 | NoCom `"t8"` `main.py:608-615`（calamine 优先 `convert @776-801`，回退 openpyxl @819-835）；PyCode @1108-1122（openpyxl @1295；.xls/.xlt COM Value2 @1314-1332） | 读指定单元格拼前缀 → `os.rename`，不保存工作簿 | ● | `pytools rename-by-content`（`pytools.rs:112`）；合成 12 路径对拍（`content_rename_46_parity.json`） |
| PT-4-7 | 取消内容前缀 | NoCom `"t9"` `main.py:616-623`；PyCode @1123-1137 | 同上反向 | ● | `pytools unrename-by-content`（`pytools.rs:113`）；合成 13+2 路径对拍 |
| PT-4-8 | 局部映射重命名（M:N） | NoCom `"t10"` `main.py:624-630`；PyCode @1138-1147 | 纯文件名片段替换（convert NoCom @948-974），任意扩展名 | ● | `pytools rename-local-map`（`pytools.rs:114`）；合成 10 路径对拍 |
| PT-4-9 | 文件排序前缀 | NoCom `"t12"` `main.py:638-644`；PyCode @1162-1175 | 规则匹配改名 + xlsx 报告（convert NoCom @286-334） | ● | `pytools file-sort-prefix`（`pytools.rs:110`）；合成 8 路径对拍 |
| PT-4-10 | 取消文件排序前缀 | NoCom `"t13"` `main.py:645-651`；PyCode @1176-1188 | convert NoCom @336-405（必要时读内容单元格） | ● | `pytools remove-file-sort-prefix`（`pytools.rs:111`）；合成 10 路径对拍 |
| PT-4-11 | Python 运行模式切换 | 历史功能；当前 PyCode grep `PYTOOLS_MODE` **0 命中**（laptop 基线恢复时移除） | — | △ 源已移除 | 无迁移对象 |

> 编号说明：PT-4-6..4-10 与函数号 46/47/48/49/410 及 workspace 既有文档一致；PT-4-5 提交缺失检查在 NoCom 菜单中排在 4-8（两变体菜单序不同，见 gap 文档 §6 P9）。`run_batch_rename_33` 的"33"为历史编号（pytools-core/MIGRATION_INVENTORY.md 将其记在"3-3"标题下），与菜单 4-1 并存，属编号不一致问题。

### 2.5 配置工具（菜单 5）

| 编号 | 功能 | 用户入口 | 输入 → 输出 | 状态 | Rust 入口/证据 |
|---|---|---|---|---|---|
| PT-5-1 | 初始化/修复配置表 | NoCom `"c"` `main.py:420-424`；PyCode @661-669 → `config_init.py::initialize_or_repair_config`（openpyxl 原位 @306，写表头批注 @188-196，可清除旧批注 @213-214） | config.xlsx（11 个业务 sheet）原位修复/新建 | ● | `pytools config-init`（`pytools.rs:85`，writer 新建）+ `config-repair-plan`（`pytools.rs:86`，计划）；源模板对拍未完成 |
| PT-5-2 | 查看当前配置 | NoCom `"cv"` `main.py:426-427`→`_view_config` @358-408 | pandas 只读打印 | ● | `pytools config-view`（`pytools.rs:87`） |

### 2.6 汇总工具（菜单 6）

| 编号 | 功能 | 用户入口 | 输入 → 输出 | 状态 | Rust 入口/证据 |
|---|---|---|---|---|---|
| PT-6-1 | 按使用区域汇总（全量） | NoCom `"s21"` `main.py:703-712`；PyCode `"s21com"` @1268-1289 | 多源（含 .xls/.xlsb）→ 新建单 sheet"汇总"（`summary_321_322.py` NoCom `Workbook()` @203，freeze A2 @271）；NoCom 用 CalamineWorkbook 直连且**整批打开两遍**（扫列宽 @217-221 + 汇总 @240,251）；PyCode=openpyxl data_only @198/224 | ● | `pytools summary-used-range`（`pytools.rs:75`） |
| PT-6-2 | 按批注汇总（模板+源文件） | NoCom `"s22"` `main.py:847-860`；PyCode `"s22com"` @1290-1326 | 模板批注是核心输入（openpyxl 解析 `_extract_comments` @94-100、`_extract_regions` @124-148、`_extract_set_info` @151-159）；NoCom 源值=calamine @351-352；COM 版 .xls 模板先转临时 xlsx（`_ensure_openpyxl_compatible` @47-71，`tempfile.mkdtemp` @245）→ 新建多 sheet @409-413 | ● | `pytools summary-comments`（`pytools.rs:76`） |

### 2.7 一二批处理（菜单 7；非 COM 版两变体字节一致，PyCode 每项有 COM 版）

| 编号 | 功能 | 用户入口 | 输入 → 输出 | 状态 | Rust 入口/证据 |
|---|---|---|---|---|---|
| PT-7-1 | 拆分村镇银行数据 | NoCom `main.py:713-735`；PyCode @1346-1390（自动路由） | 读两期（data_only+read_only @`feature_1_1_split_village_bank.py:47-48`）算 6×8 矩阵；**原位改目标** keep_vba @108、`delete_rows+insert_rows` @124-125、save @131（openpyxl 删插行不修正合并区） | ● | `pytools village-split-plan`（`pytools.rs:97`，计划） |
| PT-7-2 | 机构名称标准化 | NoCom `main.py:736-748`；PyCode @1391-1418 | 原位改分机构表 A 列（keep_vba @`feature_1_2_normalize_institution.py:80`） | ● | `pytools institution-normalize-plan`（`pytools.rs:101`）；计划级合成对拍 3 变更格（`institution_normalize_plan_parity.json`）；旧直接 umya 执行路径仅编译进测试 |
| PT-7-3 | 删除分机构表外资行 | NoCom `main.py:749-762`；PyCode @1419-1447 | **先复制"{stem}(金融局)"副本再改副本**（`shutil.copy2` @`feature_1_3_remove_foreign_bank.py:55-62`）；连续块自底向上 `delete_rows` @39-52、save @73 | ● | `pytools remove-foreign-bank-plan`（`pytools.rs:98`，计划） |
| PT-7-4 | 外汇页眉修改（万元→万美元） | NoCom `main.py:763-769`；PyCode `"x14com"` @1497-1520 | 原位改 odd/even/first 三套页眉 left/center/right（@`feature_1_4_fx_header.py:23-31`），有变化才保存 @48 | ● | `pytools fx-header-plan`（`pytools.rs:102`）→ ExcelEditPlan + OOXML 页眉补丁后端；合成对拍（`fx_header_parity.json`） |
| PT-7-5 | 地区总分校验 | NoCom `main.py:770-778`；PyCode @1459-1480 | 只读（data_only+read_only @`feature_1_5_region_sum_check.py:172`）→ **新建报告** `Workbook()` @163，写字体/填充/对齐/列宽 @122-156，输出 @206 | ● | `pytools region-sum-check`（`pytools.rs:109`）；合成 5 场景对拍 + CLI 集成测试；写出后经 `ooxml_patch::clear_empty_formula_cache` 定点补丁（`pytools-core/src/region_sum.rs:564`，已实地核对） |
| PT-7-6 | 工作表数据提取 | NoCom `main.py:779-791`（纯值版：`_write_values` @358-362，`read_sheet_2d` @450）；PyCode `"x16com"` @1521-1558（普通版整表走 `_copy_full_sheet_with_style` @381-413：值+样式+行高列宽/隐藏+合并，**无批注**）；COM 批注版 `feature_1_6_extract_sheets_com.py:140` `src_sheet.Copy`（**不迁移**） | 按配置提取 → 新建/追加输出簿（@441-443，save @497） | ◐ | `pytools extract-sheets(-area)`（`pytools.rs:108`）：area 值提取合成对拍通过；整表 ImportRange 候选已实现（`sheet_extract_full_import_parity_20261023.json`）但需能力证据、追加未实现 |
| PT-7-7 | 涉农贷款比上月修正 | NoCom `main.py:792-816`；PyCode @1559-1591 | 本期可写 keep_vba @`feature_1_8_adjust_rural_loan.py:126`、上期只读 @127；分机构表写差额列 D @148-156（**覆盖公式为值**），save @168；容差核对 @174 | ● | `pytools rural-loan-plan`（`pytools.rs:99`，计划） |
| PT-7-8 | 季度涉农修正及分地区核对 | NoCom `main.py:817-846`；PyCode @1592-1633 | 同构（@`feature_1_9_adjust_rural_loan_quarterly.py:110-111`），回填全部地区分机构表后核对 | ● | `pytools quarterly-rural-loan-plan`（`pytools.rs:100`，计划） |

### 2.8 归档工具（菜单 8；两变体同代码，纯文件系统，唯一 Excel 写点是计划/结果报告）

| 编号 | 功能 | 用户入口 | 输入 → 输出 | 状态 | Rust 入口/证据 |
|---|---|---|---|---|---|
| PT-8-1 | 按日期归档预览 | NoCom `main.py:652-662`；PyCode @1189-1202 | 文件名正则提年月（`convert` NoCom `_archive_date_from_name` @145-160）→ `归档计划.xlsx` 报告 | ● | `pytools archive-date-preview`（`pytools.rs:77`）；合成对拍（`archive_parity_20261023.json`） |
| PT-8-2 | 按类型归档预览 | NoCom `main.py:663-673`；PyCode @1203-1217 | 按"归档类型配置"关键词规则 | ● | `pytools archive-type-preview`（`pytools.rs:78`）；合成对拍 |
| PT-8-3 | 执行归档计划 | NoCom `main.py:674-680`；PyCode @1218-1227 | pandas 读计划 → `shutil.copy2` 复制 + 结果报告（convert NoCom @549-597） | ● | `pytools archive-execute-plan`（`pytools.rs:81`）；合成对拍 13 项 12 成功 1 跳过 |
| PT-8-4 | 按日期直接复制 | NoCom `main.py:681-702`；PyCode @1228-1257 | 预览+copy2，扩展名黑白名单（配置 E2/F2 @120-136），兜底目录 | ● | `pytools archive-date-copy`（`pytools.rs:79`）；合成对拍 12/12 |
| PT-8-5 | 按类型直接复制 | 同上 | 同上 | ● | `pytools archive-type-copy`（`pytools.rs:80`）；合成对拍 10 成功 2 跳过 |

### 2.9 快捷工具（菜单 9）

| 编号 | 功能 | 用户入口 | 输入 → 输出 | 状态 | Rust 入口/证据 |
|---|---|---|---|---|---|
| PT-9-1 | 打开配置文件 | NoCom `main.py:429-456`（`os.startfile`/`xdg-open` @290-297）；PyCode 仅 `os.startfile` @502-505 | 打开 config.xlsx | ● | `pytools open-config`（`pytools.rs:105`→`quick.rs`） |
| PT-9-2 | 打开 output 最新文件 | 同上（mtime 最新 @299-305） | 打开文件 | ● | `pytools open-latest-output`（`pytools.rs:106`） |
| PT-9-3 | 打开 output 文件夹 | 同上 | 打开目录 | ● | `pytools open-output-dir`（`pytools.rs:107`） |

### 2.10 未接线入口

| 编号 | 功能 | 证据 | 状态 |
|---|---|---|---|
| PT-SV1 | 调研选项统计（sv1） | `SURVEY_MENU` 定义 `PyCode/main.py:198-203`；dispatch `"sv1"` @1327-1345；`survey_stats.py`（636 行，openpyxl 读问卷+config，输出明细/展示两 xlsx）；但 `main()` @1650-1717 菜单映射无任何项指向 sv1——2026-10-23 laptop 基线恢复时菜单路由未带回（源项目 rules 规范 233 行注明恢复） | △ 待确认（是否纳入迁移范围需用户决定） |

## 3. Open-Data-Audit 功能清单（23 项）

### 3.1 产品形态与入口

- 壳层：`shell-flask/run.py`（Flask 服务器，Linux deb / Windows exe 打包）与 `shell-pywebview/run.py`（桌面 WebView），二者均包装 `core/src/base_audit/web_app.py` 的 `WebApp`（3633 行）。前端为 `core/frontend/web/index.html`，经 JS bridge 调用 `WebApp` 公开方法（方法清单见 web_app.py:568-3396 的类方法）。
- 无独立 CLI 业务入口；`systems/s5_configured_statistics/expression_export.py:355-370` 是唯一的 argparse 子工具（S5 表达式参考导出）。
- 计算引擎抽象：`engines/__init__.py:99-105` `pipeline_kind`——Windows=`"com"`（ExcelSession，Excel/WPS COM），非 Windows=`"native"`（openpyxl + soffice 重算）；引擎协议 `engines/protocol.py:29-36`。

### 3.2 核心系统功能（15 项 + S4 待确认）

| 编号 | 功能 | 用户入口 | 输入 → 输出 | 状态 | Rust 对应 |
|---|---|---|---|---|---|
| AUD-S1-F01 | 汇总核查表校验 | `web_app.py::start("audit")` @3195-3208；审核前检查 `start("check")` → `service.preflight` @451；辅助入口 `recognize` @3160、`inspect_template_formulas` @3164、`set_selected_files` @3187 | 模板 + 源核查表目录 → 审核结果工作簿（Rust writer 子进程，`native/s1_summary_writer.py:267-303`，AUTO/OPENPYXL/RUST 三态 @`web_app.py:434-440`）、历史记录（`native/history_io.py:67-90`）、导航链接 | ◐ 部分核心 | `open-data-audit template-suggest`（`audit-core/src/bin/open-data-audit.rs:127-136`）+ `s1_template.rs`/`discovery.rs`/`models.rs`/`template.rs` 等切片；条件格式求值（`conditional_engine.py`，COM）无 Rust 对应 |
| AUD-S1-F02 | 汇总校验结果说明 | `start("summary")` → `service.summarize_regions` @183-323 | 模板+源 → 说明汇总工作簿（native：`native/summary.py:257,512,638`；读取引擎三态 `engines/__init__.py:31-34`：纯 Python(python-calamine AUTO，`native/summary_source_reader.py:17`)/Excel-WPS COM/LibreOffice）；Rust writer 写结果（`s1_summary_writer.py:39-40,105-133`）；历史说明 | ○ 尚未迁移 | 无 |
| AUD-S1-F03 | 组合工作表（核查表） | `start_dag_flow`（combine 预设）→ `service.merge_org_files` @357；`simple_handlers.py::handle_combine_sheets` @165 | 多机构工作簿 → 合并簿（`merge_org.py` COM 与 openpyxl 双路径）；源文件保护 | ○ | 无 |
| AUD-S1-F04 | 组合工作表（按配置） | DAG 编辑器 API（`web_app.py:1118-1236`）+ `start_dag_flow` @3209；`workflow/runner.py`（619 行）+ `combine_settings.py` | 固定 DAG 配置 → 按配置组合输出（COM/native 两套适配器） | ○ | 无 |
| AUD-S1-F05 | 制作联合模板 | `web_app.py::create_template_merge` @1702 → `service.merge_template_files` @407-438 | 多模板 → 联合模板（native `native/template_merge.py::run_template_merge_native` 或 COM `template_merge`）；名称冲突处理、引用重写 | ○ | 无 |
| AUD-S2-F01 | 执行报表规则 | `web_app.py::start_period_compare_v2` @2741 → `systems.s2_report_collection.PeriodAuditV2Service` @2861；配置检查 `check_period_audit_config` @568 | 两期报表目录 + 配置簿 + 可选外部簿 → 三表结果 xlsx。读取：python-calamine 0.3.1 快速路径 + openpyxl 回退 + xlrd .xls（`systems/s2_report_collection/importer.py:25-30,146,151`）；写出：Direct OOXML `_write_direct_ooxml`（`exporter.py:348,409,441`，ZIP_DEFLATED level 6，临时路径） | ● 接入未验 | `open-data-audit s2`（`open-data-audit.rs:104-126`）+ `open-data-audit serve`（@100-103，web bridge 接 S2）；合成 3 场景对拍 0 差异（`s2_service_parity.json`） |
| AUD-S3-F01 | 两期数据到比较文件 | `web_app.py::start_central_comparison` @2367 → `s3 service.run_comparison` @2389（office_eval @2397） | CSV 导数（`csv_importer.py`，python-calamine）+ 配置 → 比较文件；**Rust writer 子进程** `rust/s3_comparison_writer`（`native/s3_comparison_writer.py:44-49,76-146`，调用点 `comparison_exporter.py:130`） | ◐ 部分核心 | `audit_core::central_csv::parse_central_csv` 切片对拍（`central_csv_parity.json`） |
| AUD-S3-F02 | 比较文件到说明文件 | `start_central_explanation_export` @2523 → `explanation_exporter.export_selected_explanations` @2546 | 比较文件 → 说明文件 | ○ | 无 |
| AUD-S3-F03 | 机构说明到说明文件 | `import_central_explanation_feedback` @2565 → @2592 | 多机构反馈 → 说明回写+冲突报告 | ○ | 无 |
| AUD-S3-F04 | 本期数值核对 | `start_central_cross_period` @2433 → `run_cross_period_check` @2458 | 日报/月报批次 + 配置 → 核对结果 | ○ | 无 |
| AUD-S3-F05 | 比较文件到金融表单 | `start_central_forms` @2478 → `render_financial_forms` @2501 | 比较结果 → 金融表单（`fast_ooxml.py` TemplateCompiler/FastOoxmlRenderer，OOXML 直写） | ○ | 无 |
| AUD-S5-F01 | 逐笔统计（配置版） | `systems/s5_configured_statistics/web_api.py::start_s5` @684-685（仅接受 `"S5-F01"`） | 模板 + 只读输入 → 识别/校验结果（引擎复用 `s1audit_configured`）；结果写出 `s1audit_configured/outputs.py`：fast_ooxml `write_tables`（direct）或 openpyxl（`_write_result_workbook_openpyxl` @309、save @405；env `AUDIT3_WRITER_ENGINE=direct|openpyxl`） | ○ | 无 |
| AUD-S5-F02 | 隐藏维护工具一 | `web_api.py::start_s5_maintenance` @577（五连击解锁，见 `TEST_GROUPS.md` S5 节） | 具体业务行为待源码确认 | ○ 待确认 | 无 |
| AUD-S5-F03 | 隐藏维护工具二 | 同上 | 待确认 | ○ 待确认 | 无 |
| AUD-S5-F04 | 隐藏维护工具三 | 同上 | 待确认 | ○ 待确认 | 无 |
| AUD-S4 | （预览流程，workspace 清单标记待确认） | **当前源码未发现独立 S4 系统**：`core/src/base_audit/systems/` 仅 s1/s2/s3/s5 | — | △ 待验证 | 无 |

### 3.3 配套功能（7 项，均为用户可触发的独立操作）

| 编号 | 功能 | 用户入口（web_app.py） | 实现位置 | 状态 |
|---|---|---|---|---|
| AUD-C01 | 配置指引/初始化/重置（分系统） | `initialize_config` @859、`reset_config` @865、`set_summary_config_default` @952、`set_period_config_default` @958、`set/reset_external_reference_config` @962-977、`set_central_config_default` @1008、`reset_summary/period_config` @1012/1052、S3 检查与重置 @2020-2363 | `config_guide.py`、`period_compare.py`、`name_config.py`、s2/s3 config（openpyxl 写配置簿） | ○ |
| AUD-C02 | 历史记录查询/历史说明 | `get_history_page` @1265、`open_history_explanation` @1248 | `history.py`、`native/history_io.py::write_history_xlsx` @67-90（openpyxl 原位追加） | ○ |
| AUD-C03 | LibreOffice 安装/卸载/麒麟维护 | `get_libreoffice_install` @1358、`choose_libreoffice_source` @1385、`start_libreoffice_install/uninstall` @1404/1410、`start_kylin_maintenance` @1416 | `native/libreoffice_install.py` | ○（sidecar 骨架未接通重算） |
| AUD-C04 | 诊断工作簿导出 | `export_diagnostic_workbook` @664 | web_app 内联（openpyxl） | ○ |
| AUD-C05 | DAG 工作流编辑器 | `get_dag_editor_data` @1118、`validate/save_config_editor_draft` @1137-1164、`save_business_workflow` @1166、`validate/save/delete_dag_workflow` @1194-1236 | `node_flow_config.py`、`workflow/*` | ○ |
| AUD-C06 | 外部参考配置工作簿 | `open_external_reference_config` @962 及 set/reset；`choose_history_config` @1682 | `external.py`、`external_sheet_writer.py` | ○ |
| AUD-C07 | S5 表达式参考导出（CLI） | `expression_export.py:355-370` argparse main（V_CUR/V_EXT 副本批量生成） | 独立 CLI | ○ |

### 3.4 反向跨语言现状（重要发现）

源 ODA 产品**已在生产路径嵌入 Rust 子进程 writer**，与"Python→Rust 单向迁移"的直觉相反：

| Rust 二进制 | 消费方 | 证据 |
|---|---|---|
| `core/rust/s1_summary_writer`（rust_xlsxwriter） | S1-F01 审核结果与 S1-F02 汇总结果写出，临时文件 + openpyxl 回退，AUTO/OPENPYXL/RUST 用户可选 | `native/s1_summary_writer.py:39-40`（二进制定位）、`:105-133`（临时文件+子进程+回退）、`web_app.py:434-440`（设置项） |
| `core/rust/s3_comparison_writer` | S3-F01 比较文件写出 | `native/s3_comparison_writer.py:44-49,76-146`、`comparison_exporter.py:14,130` |
| `core/rust/s5_result_writer` | **未发现任何 .py 调用方**（grep `core/src` 0 命中）——待验证：可能是历史遗留或经未索引路径调用；S5 实际写出走 `s1audit_configured/outputs.py`（fast_ooxml/openpyxl） | 本轮 grep 证据 |

这证实 workspace 规则 §6.2/§6.3 中"Python S5/S2 专用 writer"的部分性能基线来自这些已存在的 Rust/直写实现，迁移基线口径应把"源产品已部分 Rust 化"登记在案。

## 4. 状态口径补充

- "接入未验"的统一限制（对全部 45 项成立）：真实业务模板/真实目录数为 **0**；目标平台（麒麟 V10 / 统信 V20 / glibc 2.28 deb）未验收；Windows 原生 Rust 运行未系统验证（pytools 部分对拍在 WSL Ubuntu 执行）。证据边界逐项见各 probe/reports JSON 的"不能据此推断的范围"列。
- pytools 全部 44 个 CLI 命令入口见 `pytools-core/src/bin/pytools.rs:70-117`（手写分发，44 个命令字符串，已实地核对）；ODA 3 个子命令见 `audit-core/src/bin/open-data-audit.rs:24-32`。
