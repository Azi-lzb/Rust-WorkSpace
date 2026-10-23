# 双代理功能审计对拍复核（2026-10-23）

> 重大勘误：① pytools 为 **45 个分发分支、46 个可接受命令名称**（含一个别名），不是 44；② ODA 的 `start("audit"/"summary")` 没有活动业务分支，S1 实际通过内置 DAG 进入；③ 已有输出 validator 和 XLSX inventory 实现，“无 validator”“机构规范化始终 inventory_complete=false”不成立；④“追加=重建替换”只适用于时序/宽表的 `append_to_target`，不能推广到 PT-2-5、PT-7-6；⑤ S1 已有真实模板推荐 smoke 和真实公式实验，不能写成所有证据均为合成。这些修正均不构成任何功能的真实样本及 glibc 2.28 deb 验收。

复核对象为 ZCode 的四份架构审计文档和当前代码；审计快照按任务指定约为 2026-10-23 15:00。复核结论只使用“确认 / 不符 / 过时”，功能保留状态另用“保留 / 丢失 / 损坏 / 不适用 / 未验证 / 待确认”。本阶段未修改业务代码、四份原审计、规则、默认路由、依赖或源项目。此前启动的 `cargo fmt --all && cargo test --workspace` 已完成，退出码 0；这是前轮验证，不是本轮只读复核新跑的业务验收。

路径缩写：`WS` 为本 workspace；`ODA` 为 `../3.open-data-audit/core/src/base_audit/`；`PT` 为 `../4.pytools-main/`。所有 file:line 均针对本次读取的文件；后续变更可能使行号移动。源项目只读。根仓库无首次提交，无法用 commit 锁定审计；本次以文件及行号、已有报告、修改记录和文件 SHA-256 核对，未 stage/commit/push。

## ① 确认一致

承重结论的主判定如下；复合结论有一部分不成立即记“不符”，细分边界见本节及②。新增证据的快照过时单列于②。

| 任务结论 | 主判定 | 决定性当前证据 |
|---|---|---|
| 1：44 个 CLI 子命令 | 不符 | `pytools-core/src/bin/pytools.rs:71`、`:108`、`:115`，45 分支/46 名称 |
| 2：修改门控且无生产能力 JSON | 确认 | `pytools-core/src/edit_plan.rs:69`；`pytools.rs:668`；非 target JSON 检索 |
| 3：唯一 sidecar 调用且 native 没有损失标注 | 不符 | `conversion.rs:155` 调用确认；`:658` 已有英文损失声明，但缺 C01/缓存声明 |
| 4：region_sum 直接调用补丁 | 确认 | `pytools-core/src/region_sum.rs:564`，仅新建报告产物 |
| 5：旧 umya 执行只编译到测试 | 确认 | `institution_normalize.rs:276`、`:323` |
| 6：Excelize 是占位 | 确认 | `crates/excel-editor/src/plan.rs:559`；`edit_plan.rs:75` |
| 7：ODA start audit/summary/check 与系统入口 | 不符 | `ODA/web_app.py:3352`、`:3355`、`:3367`；S2/S3 独立入口存在 |
| 8：S2 Direct OOXML 写和 Python reader 回退 | 确认 | `ODA/systems/s2_report_collection/exporter.py:348`；`importer.py:140` |
| 9：S1/S3 Rust writer 消费者与 S5 待查 | 确认 | `ODA/native/s1_summary_writer.py:115`；`native/s3_comparison_writer.py:123`；S5 静态消费者搜索关闭 |
| 10：pytools 无业务 Direct OOXML，追加一概重建 | 不符 | 无 ZIP/XML 直写命中；`dedup_311.py:398` 和 `feature_1_6_extract_sheets.py:471` 不重建旧 sheet |

### 承重结论逐项复核：确认项

- **结论 2：确认，须限定为已有工作簿修改的默认执行路径。** `pytools-core/src/edit_plan.rs:69` 遍历外部能力文件，`:71` 注册 `UmyaXlsxBackend`，`:72` 注册 `OoxmlHeaderTextBackend`，`:75` 拒绝其他引擎。`pytools-core/src/bin/pytools.rs:651` 的 `execute_optional_plan` 仅在同时提供 capabilities/platform 时执行；`:668` 默认返回 `plan_only`。`crates/excel-editor/src/plan.rs:759` 检查 Passed，`:765` 检查 inventory，`:768` 检查整批操作，`:780` 检查保留项，`:784` 检查平台，`:787` 检查证据字段。本次检索非 target 的 139 个 JSON，未发现符合 `BackendCapability` 字段结构并已验收的生产记录。**不存在现成能力登记**，但不等于没有 validator；该错误另见②。新建目标簿、只读预校验和文件系统操作不属于“所有命令仅产计划”的范围。
- **结论 3 的 sidecar 调用点：确认。** `pytools-core/src/conversion.rs:155` 是业务源码中唯一启动 `lo-convert` 的调用，固定 `convert`、`--mode fidelity`（`:156`、`:164`）。`quick.rs:67` 等是系统文件打开器，不是转换 sidecar。`sidecar/lo-convert/src/main.rs:124`、`:197` 是 sidecar 内部的 soffice/PowerShell 子进程；没有从业务库静态依赖 sidecar。native 输出标注的复合结论见②。
- **结论 4：确认，但“无门控写路径”只可理解为生产业务源码直接调用 OOXML 写补丁。** `pytools-core/src/region_sum.rs:544` 创建新报告临时路径，`:546` 由 rust_xlsxwriter 新建，`:564` 只对该新产物的结果标记公式清空缓存，`:580` 不覆盖地提交。它不是对用户既有输入簿绕过 ExcelModifier。检索 `audit-core/src`、`pytools-core/src` 的直接 `ooxml_patch::` 写调用，除探针 bin 外只有此处；不能把普通新建报告的 `.save()` 也称为门控例外。
- **结论 5：确认。** `pytools-core/src/institution_normalize.rs:276` 的 `#[cfg(test)]` 保护旧 `run_normalize_institution`，`:323` 同样保护实际直连 umya 的 `process_workbook`；具体打开/保存位于 `:335`、`:396`。生产 CLI 走 `pytools-core/src/bin/pytools.rs:588` 的计划构建与 `:596` 的统一执行边界。注释 `institution_normalize.rs:100` 与编译条件一致。
- **结论 6：确认。** `crates/excel-editor/src/plan.rs:559` 仅有 `ExcelizeWorker` 枚举；`pytools-core/src/edit_plan.rs:75` 对它拒绝执行。`audit-core/core-go/README.md:1` 是该目录唯一文件；没有 Go module、worker 或 IPC 实现。本轮用户要求生产禁 Go，故占位不能据此启用。
- **结论 7 的 S2 和五个 S3 公开入口：确认。** `ODA/web_app.py:2741` → `:2861` 的 `PeriodAuditV2Service().run`；S3-F01 `:2367` → `:2389`/`:2403`；S3-F04 `:2433` → `:2458`/`:2462`；S3-F05 `:2478` → `:2501`/`:2507`；S3-F02 `:2523` → `:2546`/`:2550`；S3-F03 `:2565` → `:2592`/`:2596`。`start("audit"/"summary")` 部分不符，见②。
- **结论 8：确认，版本限源代码声明。** `ODA/systems/s2_report_collection/exporter.py:348` 定义 `_write_direct_ooxml`，`:409` 用 Python ZipFile/ZIP_DEFLATED level 6，`:441` 向临时路径写出。`importer.py:25` 注释声明 python-calamine 0.3.1，`:28` 可选导入；`:140` 先 `_open_fast`，`:145` 对 XLS 走 xlrd，`:151` 对其余输入回退 openpyxl。**代码注释不是本次实际安装版本测量**；workspace Rust 读取方向仍为纯 Rust Calamine，不移植该 Python fallback。
- **结论 9：确认，并关闭 s5 消费方的仓库内静态待验证项。** S1 两类调用见 `ODA/native/s1_summary_writer.py:115`、`:277`，调用者为 `native/summary.py:30`、`native/result_writer.py:11`；S3 见 `native/s3_comparison_writer.py:123`，消费点 `systems/s3_central_statistics/comparison_exporter.py:130`。全源仓库检索 `s5_result_writer`、`s5-result-writer`、`write-tables`、`AUDIT_RUST_ZIP_LEVEL`，排除 .git/.venv/target/node_modules；S5 名称仅出现在自身 crate 和历史技术盘点，无产品、打包或测试消费方；其他 `write-tables` 命中明确属于 S1。`../3.open-data-audit/core/rust/s5_result_writer/Cargo.toml:2` 定义实验二进制，`src/main.rs:186` 和 `src/bin/umya_writer.rs:68` 是各自 CLI 协议。结论为“**仓库内未接生产调用的原型**”，不推断仓库外人工调用是否存在。S5 现用 `../3.open-data-audit/core/src/s1audit_configured/outputs.py:211` 的 direct/openpyxl 路径；不能把 S5 原型算作第三个已接生产 Rust writer。
- **结论 10 的 Direct OOXML 搜索：确认，限两个源业务目录。** 对 `PT/PyCode/core` 和 `PT/NoComPyTools/nocompytools/core` 检索 `zipfile|ElementTree|ZipFile|lxml|writestr` 为零命中；这说明未定位到源业务的 ZIP/XML writer，不证明第三方库内部不用 OOXML。`PT/NoComPyTools/nocompytools/core/io_excel.py:278` 的确用 `if_sheet_exists="replace"` 重建单个目标 sheet；它保留其他 sheet（`:277`），并非重建整个工作簿。适用功能边界见②。

### 关闭 AUD-S3-F02/F03/F04 写出引擎待验证项

| 功能编号 | 当前源代码完整写入路径 | 保护机制与边界 | 复核结论 / 功能状态 |
|---|---|---|---|
| AUD-S3-F02 | `ODA/systems/s3_central_statistics/explanation_exporter.py:86` → `:92` 只读 data_only 打开比较文件 → `:150` openpyxl Workbook → `:191` save | 新建“说明文件”“生成信息”，标题合并/样式/冻结/筛选及数值格式见 `:152`、`:169`、`:175`；函数直接保存输出，不声称已有原子提交 | 确认 / 未验证（仅静态调用链） |
| AUD-S3-F03 | 同文件 `:283` → `:293` 可写打开目标 → `:307` 只读读反馈 → `:352` 写变动原因 → `:358` copy2 备份 → `:359` save；有冲突则 `:365` 新建 Workbook、`:381` save | 四项业务标识归一、重复/多原因冲突不回写；目标为既有工作簿编辑，冲突报告为独立新建任务，备份是产品契约 | 确认 / 未验证 |
| AUD-S3-F04 | `ODA/systems/s3_central_statistics/service.py:315` → `:344`/`:351` `_load_dataset` → `:360` `run_cross_period` → `:376` `_write_check_workbook` → `:394` openpyxl Workbook、`:425` save | 本期/可选上期 + 配置，纯新建结果；可选日志走 `:383` → `comparison_exporter.py:222` → `:226` Workbook、`:232` save，亦是 openpyxl | 确认 / 未验证 |

### workflow 处理器的活动调用链

共同入口：`ODA/web_app.py:3209` → `:3234` `_worker` → `:3367` `run_dag_flow`。`workflow/runner.py:599` 将审核流程交给 `run_dag_audit`，`:608` 将说明汇总/两类组合交给 `run_dag_simple`；native 审核链见 `:46`、`:97`。`build_registry` 在 `:137` 按平台选择 COM/native 审核 handlers，`:147` 注册审核定义，`:149` 注册汇总定义；`workflow/scheduler.py:210` 取 handler、`:214` 真正调用。**已注册 13 个审核 module_id + 4 个汇总 module_id；“14 个”只是 audit_handlers.py 中 handle_* 函数个数，不是活动注册数。**

下表 COM 与 native 两列共用 `ODA/workflow/adapters/` 前缀；表内的注册行和执行函数均已静态核对，未在本轮运行源 DAG。

| 注册 module_id / 所属功能 | COM 处理器及实际能力 | native 处理器及实际能力 |
|---|---|---|
| source.excel_files / AUD-S1 | `audit_handlers.py:111` → service._source_files，目录枚举 | `native_audit_handlers.py:87`，同类枚举 |
| source.excel_template / AUD-S1 | `audit_handlers.py:135` → `template_snapshot.py:90`/`:97`，openpyxl 只读语义快照与体检 | `native_audit_handlers.py:109`，`:118` openpyxl，汇总模板分支同样读取 |
| excel.named_range_check / AUD-S1 | `audit_handlers.py:331`，模板快照/命名区域检查 | `native_audit_handlers.py:171`、`:189`，openpyxl 名称检查 |
| excel.structure_compare / AUD-S1 | `audit_handlers.py:381` → `:402` validate_source_xlsx，结构筛选 | `native_audit_handlers.py:211`，同类结构筛选 |
| audit.copy_files / AUD-S1 | `audit_handlers.py:468` → `:477` shutil.copy2 + 哈希保护 | `native_audit_handlers.py:269`，副本保护 |
| excel.external_sheet_copy / AUD-S1 | `audit_handlers.py:497`；`:539` DirectOoxmlExternalSheetWriter.write，或 `:544` native.openpyxl_workbook.add_external_sheets | `native_audit_handlers.py:282`；`:302` Direct OOXML，或 `:305` openpyxl；**不能仅按 COM 管线名称推断此节点用 COM 复制** |
| excel.formula_copy_recalculate / AUD-S1-F01 | `audit_handlers.py:613`；`:660` OOXML 写公式或 `:674` COM apply_rules，之后 `:668`/`:681` calculate_workbook、Save | `native_audit_handlers.py:314`；`:324` openpyxl 复制公式，`:329` LibreOfficeCalculator.recalculate_many；复制和重算为不同能力 |
| excel.formula_issue_extract / AUD-S1-F01 | `audit_handlers.py:736` → `:750` native.openpyxl_workbook.extract_issues，读已保存缓存 | `native_audit_handlers.py:348`，同缓存提取契约 |
| excel.conditional_format_issue_extract / AUD-S1-F01 | `audit_handlers.py:763` → `:790`/`:796` ConditionalFormatService：PYTHON、COM_EVALUATE、COM_DISPLAY | `native_audit_handlers.py:357` → `:372` 强制 PYTHON 规则求值，非 LO 条件格式渲染 |
| audit.result_merge / AUD-S1-F01 | `audit_handlers.py:864`，按源文件顺序交错汇聚两类问题，内存操作 | native 复用此函数（`native_audit_handlers.py:459`） |
| audit.history_enrich / AUD-S1-F01 | `audit_handlers.py:885`，merge_history 内存富化 | `native_audit_handlers.py:432`，merge_history + classify_current_issues；两分支细节不能当作完全相同 |
| audit.result_workbook_write / AUD-S1-F01 | `audit_handlers.py:950` → `:962` native.result_writer.write_audit_summary → S1 Rust writer/明确模式下 openpyxl | `native_audit_handlers.py:446` → `:451` 同 writer 边界 |
| log.workbook_write / AUD-S1 | `audit_handlers.py:973` 收集日志；真正保存由 runner 的 FeatureLog.write 执行（`runner.py:365`） | 复用相同收集器；`feature_log.py:82` → `:85` → `:28` Workbook、`:38` save；不能把节点本身称为已完成磁盘写入 |
| summary.source_preflight / AUD-S1-F02 | `simple_handlers.py:49` → `:54` AuditService.preflight（write_report=False） | 相同 wrapper，service 再按平台选择 |
| summary.region_summary / AUD-S1-F02 | `simple_handlers.py:74` → `:116` AuditService.summarize_regions，见 `ODA/service.py:183` | `service.py:218` 引入 native merge/run_region_summaries；输出仍经 S1 summary writer 契约 |
| summary.history_enrich / AUD-S1-F02 | `simple_handlers.py:143` → `:152` service.enrich_summary_history，既有汇总输出簿富化 | 相同 wrapper，源方法 `service.py:325` |
| excel.combine_sheets / AUD-S1-F03/F04 | `simple_handlers.py:165` → `:198` run_combine_sheets；Windows 可 COM 兼容回退 | `:193` native 标识、`:206` allow_com_fallback=not native；UOS 为 openpyxl 组合，不依赖计算 |

`audit_handlers.py:897` 和 `native_audit_handlers.py:390` 的导航处理器均存在，但不在 `AUDIT_MODULE_DEFINITIONS`、`AUDIT_HANDLERS`（`:1156`）或 `AUDIT_NATIVE_HANDLERS`（`:461`）中，也未发现其他生产调用。故只能登记为“未接线的源处理器”，不能据此把活动 S1 输出描述为一定写审核副本导航。

## ② 与代码不符（误报、限定不准及快照过时）

### 承重结论中的不符项与建议改法

1. **结论 1：不符。** 对 `pytools-core/src/bin/pytools.rs:71` 至 `:115` 的分发实际计数为 45 个业务分支；`:108` 同时接受 `extract-sheets`/`extract-sheets-area`，所以可接受名称为 46。help 和未知命令不计。文件修改时间早于审计约 15:00，未有快照之后新增分支的记录，故不以“过时”豁免。建议四文档统一区分“菜单功能数 / 分发分支数 / 名称含别名数”，不要由 CLI 个数直接反推功能完成数。
2. **结论 3：不符（“完全没标丢弃”的表述过强）。** native 路径实际在 `conversion.rs:134` 至 `:142`；`:658` 已声明 formatting/comments/merges 等不保留及“不重算”，CSV 的 `:482` 也声明 text/no Office type inference/no source formatting。尚缺显式 C01 中文标识与旧公式缓存不保留说明。`:620` 至 `:631` 会保留公式文本，不能把 C01 文案写成“丢弃公式文本”。建议只补输出/帮助声明，不改数据处理和默认引擎。
3. **结论 7：不符（S1 audit/summary 部分）。** `ODA/web_app.py:3203` 只是日志标签；`:3352` 处理 check，`:3355` 对非 dag:* 抛“旧版流程已移除”，`:3367` 才调用 DAG。实际前端 `../3.open-data-audit/core/frontend/web/index.html:4577` 调 `start_dag_flow`。建议 AUD-S1-F01/F02 主入口改为内置 DAG，check 单列预检；不要以同名参数/日志字典证明业务可达。源文件修改时间 2026-10-22，早于审计，不属于本次并行修改导致的过时。
4. **结论 10：不符（追加范围推广）。** `io_excel.append_to_target` 的 replace 契约只由 PT-1-2/1-3/1-4 使用，见 `PT/PyCode/core/timeline_396.py:86`、`wide_397_398.py:349`。PT-2-5 是 `dedup_311.py:384` 从 max_row+1 追加、`:398` 逐格写新增行；PT-7-6 是 `feature_1_6_extract_sheets.py:463` 打开既有输出、`:470` 生成唯一名称、`:471` 新增 sheet。建议将三种追加合同拆开：**时序目标 sheet 重建替换 / 去重增量追加行 / 提取新增 sheet**，不能以审计假设覆盖这些源差异。后续用户已确认分别保留三种行为，见③第8项。
5. **“无 validator”：不符。** `pytools-core/src/edit_plan.rs:82` 已有 PlanOutputValidator；页眉验证器在 `crates/excel-editor/src/ooxml_header_backend.rs:55`/`:60` 校验补丁，umya 验证器在 `umya_backend.rs:145`/`:148` 重开并检查 inventory。真正缺口是生产证据登记及功能保留项验收；**现有 umya validator 只做可读性/包清点，不检查目标变更和全部 preserve**，不能把它算作机构规范化整体验证器。建议复用现有边界补功能验收，避免重写另一套路由。
6. **“PT-7-2 runtime inventory 恒为 false”：不符。** CLI 在 `pytools.rs:586` 调用 `inspect_xlsx_profile` 后将 profile 传给 builder；`crates/excel-editor/src/workbook_profile.rs:27` 初始为 true，未知部件/重复 entry 才置 false（`:120`、`:126`）；`institution_normalize.rs:206` 原样放入计划。计划对拍 fixture 故意设 false，不等于实际 CLI 恒 false。建议分别登记“检查器已存在但白名单/关系覆盖仍需验收”“测试 fixture 使用不完整 profile”。没有生产 capability 的门控结论不变。
7. **“S1 只有 COM 条件格式求值”：不符。** `ODA/conditional_evaluators.py:56` 定义三模式，`:191` 实现 PythonConditionalRuleEvaluator，`:222` 真正求规则；native handler `:372` 强制该路径。Rust 尚无对等业务引擎仍成立，但迁移基线已有纯规则求值资产可参考。建议同时登记“Python规则真假”和“COM显示色/COM公式求值”，按输出契约区分，不把 Office 依赖扩展到 Rust 业务规则。
8. **基准与源 Rust 资产混同：不符。** `feature-inventory.md` §3.4、`migration-gap-analysis.md` §5 把“Python S5/S2 专用 writer”与源已存在的 Rust writer 相联系，不能由名称得出它们是同一被测对象。S2 源 exporter 明确是 Python ZipFile；S5 原型 Rust 是独立未接线程序。§6.2/§6.3 证据边界应维持原报告实现身份，不改写为 Rust Direct OOXML 或已接生产 writer。源 S1/S3 Cargo.toml 的 rust_xlsxwriter 均为 0.100.0（各 `:9`/`:8`），workspace 锁定 0.101.0，不能抹平版本差异。

9. **“openpyxl 删插行不移动批注”：不符，不能把合并区和批注绑成同一行为。** 只读源环境的 `PT/.venv/Lib/site-packages/openpyxl/worksheet/worksheet.py:721`、`:739` 调 `_move_cells`，`:810` 移动原 Cell 对象，`:812` 更新 row；`openpyxl/worksheet/_writer.py:139` 从该 Cell 建批注，`openpyxl/comments/comment_sheet.py:126` 取其新 coordinate。因此存活格附着的普通批注可随格移动；被删格的批注随 Cell 删除。合并区域表及公式依赖的修正是另一问题，不能由此推断完整 COM 等价。建议删插行合同逐对象确定；本轮为读取库源码的静态核查，未把它当作真实模板批注保真验收。

### 按修改记录确认的快照过时

以下只改变“有什么证据”，不把对应功能状态从未验证升级为整体保留，不因此建立 capability：

| 功能编号 | 审计表述 | 当前证据 / 修改记录 | 复核 |
|---|---|---|---|
| PT-1-5 | 无直接预校验对拍 | `pytools-core/GPT修改记录/20261023_150000_pytools时序15预校验补齐与合成对拍.txt:1`；`20261023_151100_pytools真实配置预校验只读对拍.txt:1`；对应 JSON 比较计数和真实配置副本 | 过时（临界/后续记录） |
| PT-2-1..2-6 | 无 Python 对拍 | `pytools-core/probe/reports/dedup_parity_20261023.json:5`；`GPT修改记录/20261023_155200_pytools去重六项合成对拍.txt:1`（相对 pytools-core） | 过时；标记/删除仍是计划语义，新目标是候选写出 |
| PT-3-1/3-2/3-4 | 无对拍 | `print_comment_plan_parity_20261023.json`、`print_precheck_parity_20261023.json`；16:30/16:05 修改记录 | 过时；PT-3-3 与实际复制保真仍未验证 |
| PT-5-1/5-2 | 只有初始化 smoke / 无查看和修复计划对拍 | `config_admin_parity_20261023.json:2`、`:26`；`config_repair_plan_parity_20261023.json:2`；16:15 修改记录 | 过时；真实配置仅只读投影，修复仍未执行后端 |
| PT-6-1/6-2 | 无对拍 | `summary_tools_parity_20261023.json`；15:40 修改记录 | 过时；合成 XLSX 输出契约 |
| PT-7-1/7-3/7-7/7-8 | 无对拍 | `batch_remaining_plan_parity_20261023.json:3`；15:29 修改记录；`batch_ops.rs` 已修未匹配机构旧值保留与地区顺序 | 过时；仅计划语义，没有真实编辑后端执行 |
| PT-9-1..9-3 | 仅 smoke | `quick_tools_parity_20261023.json:2`、`:12`；15:21 修改记录 | 过时；假 xdg-open 集成，不能称旧 Windows 桌面端到端对拍 |

PT-1-1 的 `timeline_compare_parity_20261023.json`（14:42 记录）与 PT-1-3/1-4 的 `wide_summary_parity_20261023.json`（14:33 记录）已在约 15:00 之前登记；审计仍称“无端到端”应补列为遗漏，不能仅因新旧文档不同就统称后续快照过时。记录标注的时间用于本次排序，不把文件名时间当精确执行轨迹。Git 尚无 HEAD，未推断这些记录分别对应哪个提交。

## ③ 审计遗漏（补充功能、调用点与迁移闸门）

1. **AUD-S1-F01 真实候选实验已存在。** `audit-core/GPT修改记录/20261023_100439_S1F01公式引擎真实样本对拍.txt:1` 记录 Excel/WPS、Rust Formualizer 和两种 Go 候选对一组真实模板/输入的实验；Rust/Go 结果未达到等价，不能启用为生产引擎。`audit-core/MIGRATION_INVENTORY.md:59` 的真实模板推荐 smoke 另有 JSON 证据。因此“真实业务整体验收为0”仍成立，但“所有对拍均合成、没有真实输入观察”不成立。公式实验只有人工记录，不在本轮补造机器报告；完整目标平台验收仍为未验证。
2. **已接入的元数据专用读路径。** AUD-S1-F01 template-suggest 使用 sheet-directory OOXML 适配器而非 Calamine 正文读取；迁移清单 `audit-core/MIGRATION_INVENTORY.md:59` 已明确，符合对象专用元数据例外。“读取100% Calamine”必须限普通值读取，不能推广为所有元数据实现身份。打印批注/去重批注已有 reader 与计划对拍，缺的是真实保留能力，不是所有读取链从零缺失。
3. **S1 当前活动 DAG 没有导航写节点。** ①中的注册图已关闭该待验证项；“导航链接”要区分结果报表的超链接和审核副本中独立导航页。源未接线处理器是否仍属用户希望保留的功能需待确认，不能直接移植为当前行为。
4. **AUD-S5-F02/F03/F04 的操作内容可初步定位，但编号仍待确认。** `ODA/systems/s5_configured_statistics/web_api.py:639` 的维护 worker 分派 `vcur`（`:646`）、`vext`（`:655`）、`candidates`（`:663`）；分别调用表达式参考导出和公式转配置候选。源 `start_s5` 在 `:685` 只接受 S5-F01，维护为独立入口。三项确有业务，不应简单登记废弃；没有证据把某个 tool 值强行对应 F02/F03/F04，编号映射与最终范围保留待确认。
5. **能力证据的真实性没有由现路由器自动验证。** `plan.rs:787` 只检查字符串非空及平台一致，不读取 report_path、不验证报告哈希或功能测试结果。PT-7-4 测试 `pytools-core/tests/fx_header_parity.rs:39` 内构造的 capability 是合成测试证据，不是可直接分发的生产 JSON。第二阶段应建立“样本哈希+报告哈希+操作/保留项+平台/版本+validator结果+人工批准”证据模板；禁止把此测试记录直接转成 Passed 生产能力。
6. **原审计与 workspace 快照表均有遗留文本。** 除任务指定的 sidecar/页眉两处外，`rules/13` §6.1 的静态快照也仍写 sidecar 骨架、无统一路由；`pytools-core/AGENTS.md:11` 保留旧候选 gap 描述，必须和实际 cfg(test) 区分。本轮不直接改四审计/AGENTS/rules，建议下一轮按本报告统一落修订；不要借此重构已对拍实现。
7. **第二阶段的入口及停止条件。** 先为 PT-4-1、PT-4-6..PT-4-10、PT-8-1..PT-8-5、PT-7-4、PT-7-5 建脱敏真实样本清单、隔离对拍 harness 和验收报告；样本未取得及目标 glibc 2.28 deb 未验收时，状态一律未验证。此闸门通过前，不扩展到生产 capability 登记或 Audit 主链。允许独立完成已授权的 sidecar/页眉滞后文档修正及 C01 输出/帮助标注。
8. **用户已确认目标合同，保持源操作差异。** 2026-10-23 后续答复确认：PT-1-3/1-4 重建替换单个目标 sheet、PT-2-5 仅追加新增行、PT-7-6 仅新增 sheet，三种旧行为分别保留，PT-1-2 仍排除；删/插行按 Excel/COM 体验调整相关对象与引用；PT-3-3 只输出新文件，已有目标直接拒绝。决定记录在 `pytools-core/probe/reports/PT_semantic_decisions_20261023.json`；这不代表现有计划、执行器和保留项已经符合新目标。本轮不启用能力或改默认，后续实现仍须逐项验收。
9. **PT-7-7/PT-7-8 金额精度存在未覆盖差异。** `pytools-core/src/batch_ops.rs:210`、`:319` 将差额传给 `decimal_text`，`:637`/`:638` 固定格式为两位小数；源 `PT/PyCode/core/feature_1_8_adjust_rural_loan.py:150`/`:151` 和 `feature_1_9_adjust_rural_loan_quarterly.py:151`/`:152` 直接写原始差额。由代码可知，多于两位小数的差额可能被 Rust 计划提前舍入，且汇总使用未舍入值，与写出计划值不再相同。已有合成对拍不能证明此分支等价；功能状态保持未验证，后续补小数金额、正负差额及汇总一致性用例，不把固定两位小数假设成用户确认的金额合同。本轮仅登记，不改业务实现。
10. **PT-7-8 部分地区错误可能遗漏于总体核对。** Rust `batch_ops.rs:291`、`:299` 对上期缺地区表或重复机构警告后跳过，只有成功地区在 `:338` 加入结果，`:353` 只按这些地区计算总体 consistent。源季度函数 `feature_1_9_adjust_rural_loan_quarterly.py:134`、`:142` 保留错误地区记录，`:171`/`:172` 要求全部地区一致且无错误。因此有一个成功地区及另一个被跳过的错误地区时，Rust 存在错误地区缺席、总体仍可能一致的静态风险。后续须补缺表/重复机构混合成功地区的对拍，不能沿用仅正常样本的结果宣称此分支保留；本轮未运行真实样本，也未改业务代码。

本步静态对拍已完成；四份原审计保持原文，修正建议集中在本文件。真实样本、性能、保真、故障回退和目标平台均没有由本报告新增验收结论。
