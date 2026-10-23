# Excel 引擎使用情况与路由评估（excel-engine-routing-review）

> 审计日期：2026-10-23。编号与 [feature-inventory.md](feature-inventory.md) 一致。规范依据：`rules/13_Rust_Excel读写修改与引擎路由规范.md`（下称"规则"）。
> 本文档区分三个层面：**源产品当前实际引擎**（Python 现状）、**workspace 当前实际引擎**（Rust 已实现）、**建议目标引擎**（按规则 §4/§5 的候选方向，均未验收，不代表已可通过路由）。

## 1. 引擎清单与版本

### 1.1 workspace（Rust）

| 引擎/库 | 版本（Cargo.lock） | 用途 | 生产接入点 |
|---|---|---|---|
| calamine | =0.37.0 | 全部值/sheet/公式文本/合并读取 | `crates/excel-reader`（全 workspace 统一读取入口；xlsb 合并读取显式 unsupported，`lib.rs:283-287`） |
| rust_xlsxwriter | =0.101.0 | 新建工作簿 | `crates/excel-writer`（纯 re-export）；消费方：audit s2_exporter + pytools 10 个模块 |
| roxmltree + zip | 0.21 / 6 | namespace-aware OOXML 补丁与元数据补读 | `crates/ooxml-patch`（页眉文本、列定义、空公式缓存、outline）、`excel-reader`（批注/CF/DV/页眉/sheet 目录/extents） |
| umya-spreadsheet | =3.1.1 | 既有簿对象模型修改候选 | `crates/excel-editor/src/umya_backend.rs`（受能力门控） |
| quick-xml | （传递依赖） | **未使用**（workspace 代码零引用） | — |
| Go Excelize | 不存在 | 枚举占位 | `plan.rs:560`；`core-go/` 仅 README |
| LO/COM sidecar | — | 同类格式转换 | `sidecar/lo-convert`（convert + 校验器），唯一生产调用 `pytools-core/src/conversion.rs:155` |

### 1.2 源产品（Python 现状）

| 引擎 | pytools 使用 | ODA 使用 |
|---|---|---|
| openpyxl | 全部原位修改/新建/批注解析（去重、打印、配置、一二批、Sheet 改名） | native 管线主力、history、配置簿、S3 读取、S5 回退写出 |
| pandas ExcelWriter(openpyxl) | 时序/宽表/汇总/报告输出 | 归档报告 |
| python-calamine | NoCom 值读取主力（pandas 引擎 + CalamineWorkbook 直连） | S2 导入快速路径 0.3.1（`importer.py:25-30`）、S1-F02 汇总读取 AUTO（`summary_source_reader.py:17`）、S3 csv_importer |
| xlrd | PyCode 时序 .xls | S2 .xls（`importer.py:146`） |
| pyxlsb | PyCode .xlsb（pandas 隐式） | — |
| Excel/WPS COM | PyCode 全部 COM 版 + NoCom 转换；`com_engine.py` 启动即手动计算 | ExcelSession（条件格式求值、规则执行、跨簿复制、office_eval）；S1 计算引擎 |
| LibreOffice | NoCom 转换（xls/word） | S1 计算引擎（native 管线重算）、转换、安装器 |
| Direct OOXML（zipfile+XML） | **无**（grep zipfile/ElementTree 业务代码 0 命中——pytools 无 Direct OOXML 实现） | S2 exporter `_write_direct_ooxml`（`exporter.py:348`）、S3-F05 fast_ooxml、s1audit fast_ooxml、`formula_region_writer.py`、`external_sheet_writer.py` |
| **Rust 子进程 writer（rust_xlsxwriter）** | 无 | S1-F01/F02（`rust/s1_summary_writer`）、S3-F01（`rust/s3_comparison_writer`）；`rust/s5_result_writer` 无调用方（待验证） |

## 2. 引擎使用分布汇总

- **workspace**：读取 100% calamine（+roxmltree 元数据补读）；新建 100% rust_xlsxwriter；原位修改 100% 经 ExcelEditPlan/ExcelModifier（仅两个测试内后端：UmyaXlsxBackend、OoxmlHeaderTextBackend）+ 1 处无门控定点补丁（`region_sum.rs:564` `clear_empty_formula_cache`）；转换经 sidecar 子进程。
- **pytools 源**：读=calamine(NoCom)/openpyxl(PyCode) 双轨；写/改=openpyxl 一统；转换=COM/LO。
- **ODA 源**：读=python-calamine 快速+openpyxl 回退；写=Direct OOXML（S2/S3 表单/s1audit）+ Rust writer（S1/S3 比较）+ openpyxl；算=COM/LO。
- 结论：workspace 与源产品的引擎格局**不重叠**——迁移不是"换掉同一个引擎"，而是把 5 类源引擎行为收敛到 4 类 Rust 原语 + sidecar。

## 3. 逐功能组路由评估

> 目标引擎列为规则 §5 操作编号 + 候选；所有"候选"在通过对应格式/对象/整批验收前均不得成为生产默认（规则 §6.1/§9.3）。

| 功能组 | 当前实际引擎（源/workspace） | 建议目标 | 理由 | 是否需调整 |
|---|---|---|---|---|
| PT-1-x 时序读算（1-1/1-3/1-4/1-5） | 源：pandas+calamine/openpyxl；workspace：calamine | R01 Calamine（已对拍范围内）+ W01 rust_xlsxwriter 报告 | Rust-owned 值读取、纯数据输出，符合 §6.2/§6.3 适用键 | 补端到端对拍；**追加目标簿**另走 W03 决策（见 §4-2） |
| PT-1-2 时序提取 | 源：同上；workspace：calamine+writer（area）/umya ImportRange 候选（整表） | area=W01；整表=ImportRange 计划→umya（已 0 差异样本）| 值提取是新建；整表复制是 W05-W07 跨簿复制 | 整表/追加能力登记与真实模板验收；范围排除状态维持至用户确认 |
| PT-2-x 去重 | 源：openpyxl 原位标色/删行；workspace：计划（SetCellFillColor/DeleteRows/ImportRange） | 标色：umya（styles.xml 索引共享，OOXML patch 风险高）；删行：umya/Excelize 整批（须先定合并/批注修正口径，见 §4-3）；追加：ImportRange | 原位修改 W02/W03；批注读取经 OOXML 补读已具备 | **需调整**：当前无对拍、无能力登记，不能执行；语义口径未决前不得路由 |
| PT-3-1/3-2 打印复制 | 源：openpyxl 双开+新建；workspace：计划 | 块复制=W05/W07（umya 首选候选，OOXML sheet 级复制为备选）；页面设置=P01 十属性定点写入（OOXML patch 或 umya） | 值+样式+批注+尺寸整批复制的对象面大，适合对象模型引擎整批 | **需调整**：样式读取链缺失（§4-4）；十属性页面设置对拍未做 |
| PT-3-3 按配置打印 | 源：openpyxl 清空目标 sheet 重写 | 二选一并登记：①目标簿视为产品输出→改为 W01 重建（丢弃目标簿其他对象需明示）；②真原位修改→W02 整批（umya） | "清空+重写"介于重建与编辑之间，现契约模糊 | **需调整**：契约决策优先于引擎 |
| PT-4-1/4-5..4-10 文件操作 | 源：文件系统+pandas 报告；workspace：同构 | 维持（文件系统 API + W01 报告）；规则 F01/F02 | 不触碰工作簿内容 | 已达标（真实样本验收后） |
| PT-4-2/4-3 转换 | 源：COM/LO；workspace：sidecar + **native 按值分支**（`conversion.rs:120-137`） | 保真=C02 sidecar；按值=C01 且**必须明示丢弃范围** | 规则 §5：C01 按值重建必须声明丢弃；不得与保真转换混称 | **需调整**：native 分支输出须显式标注"按值重建，样式/公式缓存丢弃"，否则违反输出契约登记要求 |
| PT-4-4 Sheet 改名 | 源：openpyxl 原位（.xls COM）；workspace：计划（RenameSheet） | W03：xls 先显式 C02 转 xlsx；xlsx/xlsm 走 umya 整批并核对公式/名称/打印区域引用 | 规则 §5 W03/W04 要求引用检查 | 计划未验证执行 |
| PT-5-x 配置 | 源：openpyxl 原位+批注；workspace：writer 新建 + 修复计划 | 初始化=W01 新建；修复=W02（批注+表头+样式整批，umya 候选） | 修复对象含批注写入——OOXML patch 未经 comments/VML 验收不可用 | 修复对拍缺失 |
| PT-6-1/6-2 汇总 | 源：CalamineWorkbook 直连（**两遍打开**）/openpyxl；workspace：calamine+writer | R01+W01；批注指令=OOXML 补读（已具备 read_header_inventory 类能力，需按批注解析扩展验收） | 纯新建报表 | PT-6-1 的 used-range 语义需按 calamine extents 定义并对拍 |
| PT-7-1/7-3/7-7/7-8 批处理 | 源：openpyxl 原位；workspace：计划 | W02/W03 整批（umya/Excelize 候选）；7-3 的"(金融局)"副本保护保留在业务层 | 删插行/区域写值整批 | 语义口径（合并/批注修正）+ 能力登记双缺口 |
| PT-7-2 机构标准化 | 源：openpyxl 原位；workspace：计划（对拍通过）+ 直连 umya 旧路径（仅测试） | SetCellValue 窄操作：补齐 workbook inventory 后经 OOXML patch 或 umya 整批 | 值替换不改样式——若 inventory 完整，定点 patch 最小侵入 | **需调整**：inventory_complete=false 阻止路由（按设计）；先补清点 |
| PT-7-4 页眉 | 源：openpyxl 页眉；workspace：ExcelEditPlan→OoxmlHeaderTextBackend（合成 passed） | 维持 OOXML 页眉补丁（唯一已 passed 的修改能力） | 窄、可验证、已对拍 | 补批量统计+真实模板+xlsm 决策 |
| PT-7-5 地区校验 | 源：openpyxl 新建；workspace：calamine+writer+`clear_empty_formula_cache` | 维持；补丁范围必须锁定"结果簿空公式缓存" | 新建簿+定点补丁组合已被合成对拍覆盖 | 补真实模板/性能 |
| PT-7-6 提取 | 源：三档；workspace：area CLI+ImportRange 候选 | area=W01；整表=ImportRange→umya（已 0 差异样本，对照源属性）；追加未实现 | 同 PT-1-2 | 同 PT-1-2 |
| PT-8-x/PT-9-x | 文件系统 | 维持 | 无工作簿修改 | 真实目录验收 |
| AUD-S1-F01/F02 | 源：COM 条件格式/LO 重算 + **Rust writer** | R01 Calamine 值读取 + OOXML 元数据（CF/批注/名称）+ 条件格式求值引擎（Rust 业务）+ W01（沿用源 Rust writer 契约）+ P? 无 | 源产品 s1_summary_writer 已是 rust_xlsxwriter——其输出契约可直接作为 workspace 对拍基准，避免两套 Rust writer 分叉 | **需调整**：条件格式求值为最大新增组件；重算 sidecar 立项；两套 rust_xlsxwriter 代码的版本/契约统一（源版本**待验证**） |
| AUD-S2-F01 | 源：python-calamine+Direct OOXML exporter；workspace：Calamine 全链+rust_xlsxwriter | 维持 workspace 路由（已 0 差异对拍）；Direct OOXML exporter 保留为源基线直至真实模板验收 | Calamine 值读取与纯数据三表输出符合 §6.3 适用键；1-ULP 按 1e-15 容差处理（规则 §9.1） | 真实模板+失败矩阵+性能验收 |
| AUD-S3-F01 | 源：csv+openpyxl+Rust writer；workspace：central_csv 切片 | CSV=Rust csv+encoding_rs（已对拍）；比较引擎=Rust 业务；写出沿用源 Rust writer 契约 | CSV 解析已验收；比较计算是纯业务 | 比较引擎迁移 |
| AUD-S3-F05 | 源：fast_ooxml Direct 写 | 模板渲染=W01 rust_xlsxwriter（若可重建）或对象编辑器；**不**为它新写通用 OOXML writer | 规则 §4.2：Direct OOXML 仅在专用输出契约下进入候选 | 先判定"金融表单"能否降为新建任务 |
| AUD-S5-F01..F04 | 源：s1audit_configured（OOXML 元数据读+fast_ooxml/openpyxl 写） | 与 AUD-S1 合并规划（共享引擎）；读=OOXML 元数据适配器（样式/布局），写=W01 或沿用源 direct 契约 | S5 与 s1audit_configured 同源 | 范围先确认（F02-F04 待确认） |
| AUD-C01/C02 | 源：openpyxl 配置/历史原位 | C02 历史=W02 追加（ImportRange 候选）或维持 Python？——**不得**，生产不得依赖 Python；历史追加需登记能力 | 原位追加 | 随 S1 主链迁移 |

## 4. 六项专项检查结果

### 4-1 不必要的整本工作簿重新读取
- **发现（源）**：PT-6-1 NoCom 对同批文件打开两遍（先扫列宽 @`summary_321_322.py:217-221`，后汇总 @240）；PT-3-1/3-2 每文件双开（样式+data_only 值，@`print_310.py:61-62,452-457`）。缓解现状：lru_cache（`io_excel.py:108-120`）、P-03 四级缓存（@595-598）。
- **发现（workspace）**：未见整本重读；`inspect_xlsx_profile`、region_sum 补丁重开为定点必要行为（见 4-6）。
- **结论**：迁移实现必须单次读取（calamine 值 + OOXML 元数据一次解析），不得复刻双开模式。

### 4-2 只需新建文件却调用修改引擎
- **发现**：PT-3-3 源实现"目标存在→清空重写"混淆了新建与编辑（目标簿如果是产品上轮输出，重建即可；如果含用户对象，则必须真编辑）。workspace 目前无此混淆——所有新建走 writer、修改走计划。
- **风险项**：`pytools convert-file --engine native`（`conversion.rs:120-137`）是 Rust 按值重建，却挂在"转换"命令语义下；按规则 §5 C01 必须明示丢弃范围。**建议**：输出与帮助文本显式标注"按值重建（丢弃样式/公式缓存）"。
- 结论：workspace 无违规实例；PT-3-3 契约与 native 转换标注需修正。

### 4-3 需保留原样式却使用不当写入方式
- **发现（源语义债）**：①追加=整 sheet 重建替换（`io_excel.py:278-284`）——目标 sheet 格式丢失，但产品菜单称之为"追加"；②删/插行不修正合并区（`feature_1_1:124-125`、`dedup_311.py:160-161`），COM 版却会随行移动——同一功能两变体输出不同；③PT-7-7 写差额列覆盖公式为值——是否为期望语义未登记。
- **发现（workspace 已知引擎行为）**：umya 打开/保存会重写全簿（列定义重写致 17/18，`umya_raw_compare.json`）；umya 物化默认 A 列宽（整表导入中已加清理）。
- **结论**：这三处是"对拍基准本身二义"，必须先由业务定口径（重建是否可接受、删行是否需修正合并），再谈引擎选择；否则任何引擎都无法"对拍通过"。

### 4-4 calamine 被当作完整样式读取器
- **结论：workspace 无此违规**——`excel-reader` 明确只提供值/公式/合并/批注/CF/DV/页眉/extents，xlsb 合并显式 unsupported；样式模型（字体/填充/边框/数字格式）无任何读取 API。
- **风险提示**：打印组（全样式复制）、去重（标色但需保留其余样式）、按批注汇总（模板批注+合并左上）的 Rust 实现将需要：①OOXML 样式元数据补读（新增适配器，须逐对象验收），或②umya 对象模型整批。禁止用"calamine 读值+手工重建样式"冒充复制（那属于 C01 重建，需明示丢弃）。

### 4-5 逐单元格跨进程/worker 调用
- **结论：workspace 无违规**。修改类命令整文件单计划提交（`ExcelModifier::execute`）；规则 §4.3 禁止逐单元格 IPC，`edit_plan.rs` 结构上也只接受整批计划。
- **源观察**：ODA COM 条件格式求值在 Excel 会话内逐格触发——这是源行为，Rust 迁移时必须实现为批量求值引擎，不得映射为逐格子进程调用。

### 4-6 同一功能重复读取/解析/序列化
- **发现**：①PT-6-1 两遍打开（同 4-1）；②`region-sum-check` 写出后由 `clear_empty_formula_cache` 重开 zip 再写一次（`region_sum.rs:564`）——双重序列化，但属定点必要（保持公式空缓存与 Python 基线一致），代价可接受，建议未来由 writer 原生支持"公式不带缓存"后移除；③源 PT-6-2 COM 版为解析批注把 .xls 模板转临时 xlsx（`summary_321_322_com.py:47-71,245`）——格式转换+解析串联，迁移后应改为直接读 .xls 批注（OOXML 不适用 .xls，需另评估）或维持显式两步并登记。
- **workspace 良性项**：sidecar 转换后用 excel-reader 重开校验——这是验证步骤，非重复处理。

## 5. 规范符合性结论

1. 符合：calamine 值读取未越界、无 python-calamine/PyO3/Go 生产依赖、修改全部经 ExcelEditPlan、sidecar 显式调用、无正则改 XML、xlsb 缺口如实标注、伪 .xls 拒绝（sidecar/基准附录）。
2. 需修正：native 按值转换的契约标注（4-2）；`EXCEL_ENGINE_ARCHITECTURE.md:50` 与 `pytools-core/MIGRATION_INVENTORY.md:59` 两处文档滞后（P8）。
3. 需决策：追加契约（4-3）、PT-3-3 新建/编辑定性（4-2）、删行合并口径（4-3）、.xls 批注读取路线（4-6）。
4. 待登记：s1_summary_writer / s3_comparison_writer 作为既有 Rust 资产纳入 §6.2 类证据边界（避免与 workspace writer 双轨漂移）。
