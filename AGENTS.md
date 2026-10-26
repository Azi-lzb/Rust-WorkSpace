# Workspace development rules

## 规则入口

- 涉及 Excel/CSV 读取、新建、修改、格式转换、打印页面设置、引擎路由或性能优化时，必须先阅读 workspace 级规范 [rules/13_Rust_Excel读写修改与引擎路由规范.md](rules/13_Rust_Excel读写修改与引擎路由规范.md)，再按需查看根目录的引擎架构与 Python 操作清单。评估性能证据前，先按 data_owner、输入/输出格式和输出契约匹配该规则 §6.2；改变既有默认路由还须满足 §9.2 的完整调用链测量门槛。报告只用于其明确样本/契约/平台的候选评估，不等于已验收路由，也不授权改变现有生产默认。
- `audit-core/AGENTS.md` 和 `pytools-core/AGENTS.md` 是各自子树的自动适用规则。Audit 原 rules/ 下文档来自旧 Python 项目：业务语义可作为迁移基线，技术指令必须遵循本 workspace 的 Rust/Go 与外部引擎边界；详见 `audit-core/AGENTS.md`。涉及 Audit OOXML 结构、测试要求或 Excel 操作映射时，按需读取 rules/03、07、13。
- pytools-core 没有单独复制旧项目 rules；其迁移范围与完成状态以 `pytools-core/AGENTS.md`、`MIGRATION_PLAN.md`、`MIGRATION_INVENTORY.md` 为准。probe/ 下的 Python 仅作旧实现对拍基线，不得进入产品运行路径。
- 修改记录按 audit-core/rules/08 中兼容的留痕要求执行。audit-core/rules/11 是旧产品组件版本制度，不直接适用于本 workspace；workspace crate 版本以各 Cargo.toml 和 Cargo.lock 为准，除非另行确认，不创建旧版 VERSION_MANIFEST 或 Sx-Fxx 版本体系。
- 本 workspace 根目录和 audit-core/rules/ 下目前都没有 rules/12；不得引用不存在的 workspace 规则。Open-Data-Audit 原项目的 rules/12 仅作为本规范采纳优化原则的只读来源。
- 规则优先级：用户本轮明确要求 > 本文件硬约束 > workspace rules/13 专项规则 > 子项目 AGENTS.md > 子项目迁移清单和规则中的兼容业务约束。旧规则中的 Python/openpyxl/COM/Office 实现指引只记录源项目现状，不覆盖 workspace 技术边界。出现业务语义冲突时先报告，不自行猜测。

## 1. 项目目标与背景

本 workspace 的目标是将 Open-Data-Audit 与 py_statistics_tools 的全部确认功能迁移为以 Rust 为主、Go 为条件性备选的实现。产品在麒麟 V10 x86_64（glibc 2.31）和统信 V20 x86_64（glibc 2.28）内网环境以独立 deb 运行。开发机为 Windows 11 x64；目标机现有 Python 3.8.10 不构成生产依赖。

核心主路径不依赖 Office、WPS、LibreOffice、COM 或 UNO。公式计算、同类格式转换，以及未来 Windows 上经能力验收的 ExcelModifier COM 修改，可作为显式选择的可选能力；其运行时不可用或未验收时，核心主路径继续运行，并对该项报告“该功能需要外部引擎”。读取仍以 calamine 为主，新建仍以 rust_xlsxwriter 为主；本阶段不增加 COM Reader/Writer。

## 2. 原始 Python 项目与只读约束

- Open-Data-Audit：C:\Users\AZI\Desktop\VibeCode\3.open-data-audit
- py_statistics_tools：C:\Users\AZI\Desktop\VibeCode\4.pytools-main

两个源目录只读。不得在其中改代码、配置、测试、数据或版本记录。生产读取、业务计算和新建工作簿使用 Rust；禁止 python-calamine、PyO3、pandas 或复制 Python 业务路径。生产 Python 例外限于隔离的 openpyxl XLSX 修改 worker，以及未来 Windows 专用 pywin32 / `win32com.client` Office COM 修改/计算 worker；均不得承接业务规则或统计。openpyxl 只能通过 ExcelEditPlan/ExcelModifier 按工作簿批量执行已登记能力；COM worker 只能实现已验收的 Windows 修改或计算能力。不得依赖系统 Python/PATH；Linux 不依赖 pywin32/COM。Calamine 的优先候选范围是 Rust 持有数据、普通存储值 XLSX 且输出契约与专项基准一致的功能。其他数据归属、格式和专属元数据按既有验收路由或对象级对拍处理，不从本次基准外推。

## 3. Workspace 结构与依赖边界

rust-workspace/
├── Cargo.toml
├── AGENTS.md
├── crates/
│   ├── excel-reader/
│   ├── excel-writer/
│   ├── excel-editor/
│   ├── ooxml-patch/
│   └── probe/
├── audit-core/
│   └── tests/fixtures/
├── pytools-core/
│   ├── probe/
│   └── tests/fixtures/
├── sidecar/
│   └── lo-convert/

- Excel 原语及跨产品复用能力必须经 crates/ 下的共享库提供；audit-core 与 pytools-core 不得相互依赖。通用第三方依赖统一在 workspace 根 Cargo.toml 管理版本。
- 共享 crate 只提供 Excel 原语，不包含审核、时序、打印业务等概念。
- sidecar/lo-convert 是独立二进制。主路径显式启动子进程，不静态链接、不导入该 crate。
- 两个业务 crate 各自维护 fixtures、probe 配置和报告；报告不得混用。
- 产品分别打包为 open-data-audit.deb 和 pytools.deb；共享 Rust 库静态链接进对应二进制，不单独发 .so。
- 当前 Workspace 仅作为两个产品的临时联合迁移环境。迁移及真实业务验收完成后，按 DEVELOPMENT_GOAL.md 第十七章拆分为 Excel-Core、Audit、PyTools 三个独立 Rust 项目和 Git 仓库；现在不提前拆分。公共 Excel 接口应保持通用、可独立 Cargo 依赖，各产品最终须能独立离线构建。

## 4. 技术分工与外部引擎边界

### Rust

- 读取 xlsx、xlsm、xls、xlsb：使用纯 Rust calamine 作为读取实现的主要候选，并按格式、值契约和对象逐项对拍。性能报告只支持 Rust 持有数据、普通存储值 XLSX 且坐标、类型、空值和值语义均与报告契约一致的 reader 候选排序；不得外推到其他格式、Python 持有数据或不同输出契约。Direct OOXML 不是固定 reader 或通用 fallback；Calamine 失败必须如实报告。CSV 使用纯 Rust CSV 实现。ET 先由用户选择转换或按已确认策略转成 xlsx。
- 批注、条件格式、数据验证等 OOXML 专属读取能力，由针对该对象且已验收的 namespace-aware OOXML 适配器补齐；能力不足时报告缺项。Python-owned 迁移期流程保留原项目已验收的同契约路由，候选调整须测完整调用链，不能外推 Python reader 排序。
- Rust 持有结果、从零生成纯数据 XLSX 时优先评估 rust_xlsxwriter，须先通过该功能的输出契约对拍且符合预先定义的峰值内存预算。报告中 30 万行 × 20 列测试的 rust_xlsxwriter 峰值 RSS 观测值为 1,376.7 MiB（约 1.34 GiB）；它只属于该测试规模和默认内存模式，不是通用内存阈值或任意大文件承诺。报告测量的 Direct OOXML 写入是 Python S5/S2 专用实现，不是 Rust writer。
- 原位修改通过 ExcelEditPlan 与 ExcelModifier 统一边界。OOXML Patch、umya-spreadsheet、Go Excelize 和隔离 openpyxl worker 都是按格式、对象特征和整批操作组合验证的候选，不预设 workspace 级默认编辑器；细节见 EXCEL_ENGINE_ARCHITECTURE.md。
- 多文件并行使用 rayon，按文件级并行。

### Go

Go Excelize 是复杂 XLSX 修改的候选/保底引擎。是否保留由同一真实模板、整批操作对拍、性能和维护成本决定；若没有可验证收益，则不引入 Go。Go worker 只做既有工作簿修改，不承担批量读取或新建报表；采用批量 IPC、CGO_ENABLED=0 静态编译，并随离线安装包交付，不要求用户安装 Go。

### 隔离 openpyxl 修改 worker

openpyxl 是可选的第四个 Excel 修改引擎，不是读取器、新建 writer、公式计算器或业务逻辑运行时。Rust 在一个 workbook 级 ExcelEditPlan 中提交完整操作批次，worker 通过内部绝对路径和 `-I` 启动；禁止逐单元格启动进程、PATH 查找或依赖系统 Python。路由必须匹配已验证的格式、操作批次、OOXML 对象、保留契约、平台与 worker 版本；真实模板验收前不得登记生产 Passed 能力或设为默认。worker/运行时缺失时只可在写入前转到另一个已验证引擎，否则报告不可用。超时、崩溃、写入失败或结果校验失败都停止，不静默回退。

### 未来 Windows Office COM 修改后端

ExcelModifier 可在未来增加两个独立、可选的 Windows 后端：Microsoft Excel COM 与 WPS 表格 COM。优先复用 ExcelEditPlan、EditBackend、能力证据和工作簿级 JSON IPC；Python 侧候选为 pywin32 / `win32com.client`，Excel ProgID 为 `Excel.Application`，WPS 优先验证 `KET.Application`。Windows worker 必须随产品以隔离 Python runtime 和锁定的 pywin32 交付，不依赖系统 Python/PATH；Office/WPS 应用及实际 COM 接口仍需在运行环境验证。两者必须有独立引擎标识、版本范围和 capability 验收。应验证实际 COM 激活与所需接口，不能只按软件安装状态判断可用。用户显式选择不可用后端时明确报错，不静默切换。它们只修改既有工作簿，不提供读取器或新建 writer；不使用 Rust COM、PyO3 嵌入或 FFI。只有完整调用链实测证明 JSON IPC 是瓶颈时，才可另行评估通信机制。当前尚未实现，不能作为生产路由或界面选项。

### OOXML

OOXML 是补丁和缺失元数据读取层，不是主要读写库。使用 namespace-aware XML parser，禁止正则替换 XML 全文。维护关系、Content_Types、sheet 身份和顺序、sharedStrings、defined names。新增适配器必须对拍所涉及的保留项；遇到未验证部件直接拒绝处理。

### 公式计算、转换与 COM 修改的可选边界

以下能力都必须显式触发、按具体能力登记并验证结果；它们不得成为核心主路径的启动依赖：

1. **公式计算（独立 ExcelCalculator）**：默认不计算。本阶段目标为 LibreOffice Calc（Windows/Linux）和 Microsoft Excel COM（Windows）；此前列出的 WPS COM 是历史候选，不属于当前计算范围，除非用户另行确认。允许显式请求普通重算或完整重算；必须对拍公式语义、结果和依赖关系，并检查保存引起的公式、样式、结构变化。不得宣称 LO 与 Excel 完全等价。LibreOffice 基础公式计算的可行性已由用户实际测试确认，不重复泛泛的可行性调研；具体业务公式覆盖、文件保真和目标系统部署仍须验收。ExcelCalculator 与 ExcelModifier 分离；不强制其全部实现 JSON IPC。运行时不可用时报告“该功能需要外部引擎”，不得用旧缓存或值重建冒充重算。
2. **同类格式转换**：首批为 Excel .et→.xlsx、Excel .xls→.xlsx（供 sheet 改名等 Rust 不支持的修改前转换）、Word .doc→.docx。用户可选择保真转换或经说明的 Rust/Go 原生转换；转换后检查工作表和数据。PDF 导出、跨类别转换不在范围内。
3. **Windows COM 修改**：未来由 ExcelModifier 显式调用 Microsoft Excel COM 或 WPS 表格 COM，仅在真实 COM 接口可用且具体操作 capability 已验收时开放。优先用独立 pywin32 Worker 与 JSON IPC。COM 不提供读取或新建能力，也不作为 Linux 依赖。

主程序启动不检查外部引擎；用户显式触发能力时才检测其可用性。失败不能阻塞其他主路径功能，也不能将未校验输出报告为成功。主 deb 不含 Office/WPS/LO/COM 运行时。可选 LibreOffice 计算能力在 Linux 上发布前须明确其独立运行时来源；不得假设目标机已安装。

### 当前 Excel-Core 优先级（2026-10-26）

Read 继续使用 calamine，Create 继续使用 rust_xlsxwriter；Edit 新增能力优先落在 OOXML Patch 与 umya-spreadsheet。按少量工作簿、工作表和单元格修改及固定区域复制的真实业务规模实现，并保护非目标区域。保留现有 openpyxl worker、接口、测试和仍依赖它的兼容路径，但暂停新增扩展；只有对应 Rust 替代方案通过真实模板保真验收后才能移除。Go Excelize worker 与 FFI 暂缓开发，不是业务迁移前置条件。公式复制保持原样复制为既有语义，引用平移必须显式选择；公式重算使用 LibreOffice 或 Windows Microsoft Excel COM，并验证缓存和关键业务结果。打印整理须单独通过真实模板端到端保真验收。以上优先级不改变 Audit/PyTools 的业务范围或任何已验收生产默认路由，合成测试不等于生产验收。

### 明确禁止

禁止将 python-calamine、PyO3、pandas 或通用 Python 运行时用作读取、业务计算、新建写入或其他生产后端。openpyxl 仅可作为隔离 XLSX 修改 worker；pywin32 仅可用于未来经验证的 Windows Office COM 修改/计算 worker，不得扩展为业务兼容层。禁止在一个任务中无测量地串联 Rust 读取和 Go 改写；禁止将打印页面设置实现成渲染或 PDF 导出。LibreOffice 仅可作为独立公式计算候选及已批准的格式转换引擎；COM 仅可用于 Windows 上能力验收通过的修改/计算及获准的格式转换，不作为跨平台依赖。

## 5. 读取、写入、修改与打印默认策略

- 读取：Calamine 作为 Rust-owned 普通存储值 XLSX 的优先候选，前提是输出契约与限定基准相符；其他数据归属、格式和契约保持其已验收路由。OOXML/umya 仅作为所需专属元数据的对象级适配器，不能遮盖 Calamine 值读取错误。每种格式的能力按实际对拍登记，不能从 ZIP 容器格式推断对象支持。
- 新建：Rust-owned 纯数据 XLSX 优先评估 rust_xlsxwriter；只有输出契约通过且峰值内存符合该功能预算时才采用。Direct OOXML writer 必须有专用输出契约/OOXML 对象要求或同契约资源收益实测；压缩配置相同不代表实现等价。具体 benchmark 边界见规则 §6.2。
- 原位修改：业务只生成 ExcelEditPlan，由 ExcelModifier 按整批已验证能力从 OOXML Patch、umya-spreadsheet、Excelize worker、openpyxl worker，以及未来 Windows-only 的 Excel COM/WPS COM 中选择单一引擎。COM 两后端须独立登记能力；只有用户显式选择且 capability、平台和实际接口均匹配时才参与路由。引擎失败后必须从原始文件重新执行已验证备用方案；显式选择不可用引擎时报错，不静默切换；禁止将同一批修改依次交给多个引擎打开/保存。源操作入口盘点见 PYTHON_EXCEL_OPERATION_INVENTORY.md，具体能力与迁移证据见 EXCEL_ENGINE_ARCHITECTURE.md。
- 打印：只写页面设置，包括 orientation、fitToPage、fitToWidth、fitToHeight、清除 scale、paperSize、printArea、horizontalCentered、verticalCentered 和页脚页码。不渲染、不导出 PDF。
- 新打印只对等旧项目现状属性。现状代码从未写入 margins、printTitleRows/Columns、gridLines、headings、blackAndWhite、draft、firstPageNumber；新增需求另行立项。
- xls、xlsx、csv 是主要输入；可能遇到 xlsm、xlsb、et。对不支持或未验收的输入，提示转换为 xlsx 或显式触发已批准的转换流程，验证 sheet 与数据后再继续。

## 6. 对拍要求与决策规则

每个后端成为默认候选前必须完成对拍，逐项标注保留、丢失、损坏、不适用或未验证；API 不存在写未验证，模板没有对象写不适用，没测量写未测量。

1. 读取：probe 导出值、sheet 名、合并区域、公式文本、批注、条件格式和数据验证元数据，与旧项目基线逐项比较。
2. 新建写入：比较值、样式、合并、公式文本、命名区域和页面设置。
3. 原位修改：覆盖改值、合并、样式、公式、命名区域、批注位置和值、行列尺寸与隐藏、页眉页脚、页面设置、公式文本、未修改区域及文件可打开性。
4. 打印页面设置：逐属性读回断言；不做 PDF 渲染验收。
5. 公式重算旁路：验证缓存刷新，并与旧项目重算结果比较；无引擎时验证明确报错且不影响其他功能。

读取或新建写入达到 95% 是具体功能、格式、对象与输出契约进入默认评估的必要条件，不是充分条件。读取时，Calamine 的默认评估仅限已验收的对应功能；需要 OOXML 元数据时，应把专用适配器作为独立元数据能力验收，不得拼成通用 fallback 链或用来掩盖普通值读取错误。新建时，rust_xlsxwriter 只在对应 Rust-owned、从零生成 XLSX 的输出契约通过，且峰值内存满足该功能预先定义的预算后，才可作为该功能默认。其他数据归属、格式、对象、平台和不同契约沿用各自已验收路由；通用基准不得自动改写生产默认。原位修改不设全局默认：每个“格式 + 对象特征 + 整批操作”组合必须无损通过对应保留项验证后，才可路由至该引擎。单项失败优先评估定点补丁；是否引入 Go 取决于同一操作组合的实际保真度、性能及维护成本。目标文件不能正常打开时停止该模板迁移并记录不支持。机器可读 JSON 报告保存在各自 probe/reports/；日期化人工报告保存在各自 GPT修改记录/。两个项目的配置、夹具和报告完全分开。

## 7. glibc 与 deb 约束

- Rust 使用 cargo-zigbuild 的 x86_64-unknown-linux-gnu.2.28 目标，或经确认的 musl 静态目标。不得在 glibc 2.31 上直接编译后交付给 glibc 2.28。
- Go 使用 GOOS=linux GOARCH=amd64 CGO_ENABLED=0。
- 构建后以 objdump -T 检查 GLIBC 符号不高于 2.28。
- 核心 Linux deb 不依赖系统 Python、用户 PATH、pip、网络或 Office/WPS/LibreOffice/COM 运行时。若产品显式启用 openpyxl 修改能力，产品 deb 可在 `/usr/libexec/<product>/openpyxl/` 内带隔离 Python runtime、锁定版本的 openpyxl/et-xmlfile 和 worker；该目录不得导入业务代码。核验运行时版本、`-I` 导入、worker 协议、扩展模块动态依赖和目标 glibc，未在麒麟 V10/UOS V20 实测前不得宣称可发布。Windows COM Worker 如交付，须使用隔离 Python runtime 和锁定 pywin32；Office/WPS COM 应用需在兼容 Windows 环境可实际激活。COM Worker 不进入 Linux deb。可选 LibreOffice Calc 计算功能须明确运行时来源（独立离线组件/包或受支持的外部安装）并验证 glibc 2.28 与离线运行；未确定来源前不得作为 Linux 交付能力。
- 分别在麒麟 V10、统信 V20 安装 deb，以 ldd 检查依赖，并对各自 fixtures 跑验收。无外部引擎环境下纯 Rust 主路径必须完整运行；可选能力缺失时明确提示所需外部引擎。

## 8. 已确认事实与待确认问题

已确认：
- 现有两个项目真实模板不含图表、透视表、切片器、图片、工作表保护或 VBA 宏。此事实只覆盖已检查模板，不等于未来输入文件都没有这些对象。
- py_statistics_tools 的工作表提取以不带批注普通版为准；COM 批注版不迁移。
- 公式重算默认关闭，但允许用户显式触发独立 sidecar。
- LibreOffice Calc 的基础公式计算可行性已由用户实际测试确认，不重复泛泛调研；其具体公式覆盖、Excel/WPS 结果一致性、保存保真及目标 Linux 离线运行时均待验收。未来 Windows Excel COM/WPS COM 修改与计算是可选候选，不是当前能力。
- 打印只写页面设置；PDF 不做。
- 首批转换为 .et→.xlsx、.xls→.xlsx、.doc→.docx；不做 PDF 或跨类别转换。
- 开发环境是 Windows 11 x64；目标系统是麒麟 V10 x86_64/glibc 2.31 和统信 V20 x86_64/glibc 2.28。目标机 Python 3.8.10 不是系统依赖；如 deb 启用 openpyxl，必须另行内置并验证隔离 runtime。

待确认：
- 未来输入中 .xlsm、.xlsb、.et 的实际样本、对象保留要求及转换默认交互方式。
- NoComPyTools 与 PyCode 的行为差异中，哪些独有功能属于最终 pytools 产品范围；盘点时先列出差异，不静默遗漏。
- 用户是否需要未来新增打印属性或非首批格式转换。

## 9. 修改记录与版本治理

遵循 audit-core/rules/08_AI协作与修改记录.md 和 audit-core/rules/11_版本管理与组件编号.md 的留痕要求：修改前检查 WIP；每次实际变更新增 GPT修改记录，写明文件、目的、测试、未执行项、已知问题和回滚信息。纯文档治理修改不升级产品/组件版本，不写入用户版更新日志。源 Python 项目始终只读。

## 10. 执行顺序

按 `DEVELOPMENT_GOAL.md` 的当前阶段执行，不改变 Audit、PyTools 的约定业务范围：

1. **迁移**：Audit、PyTools 在各自 Worktree 并行完成全部约定功能，包括 Rust 业务核心、真实 GUI/CLI/前端入口、配置读写、进度反馈、错误提示及文件操作。前端必须调用真实 Rust 能力，不保留成功形状的占位实现。每项至少通过相关 crate 编译、范围内单测、最小冒烟、入口接线及必要的文件安全检查，并更新状态。公共能力阻塞须附证据并转 main，同时继续其他独立功能。
2. **Windows 集中验收**：全部约定功能具备可操作实现后，由两个 Agent 并行进行用户操作复查、真实样本对拍、集成/回归测试和缺陷修复；逐项记录通过、失败、受阻或未验证。实现完成不代表业务验收通过。
3. **Linux 集中验收**：关键 Windows 业务通过后，以稳定成果为基线，统一验证 x86_64、glibc 2.28、离线 deb、运行时依赖、LibreOffice 计算及麒麟 V10/UOS V20。迁移期间不重复执行完整 WSL 构建和发行测试；明确的 Linux 专属风险可提前针对性验证。保留现有跨平台实现和测试约束。
4. **交付优化**：修复集中验收遗留问题，整理发布包和文档；功能正确及兼容性确认后再评估非必要性能或架构优化，最终仓库拆分沿用长期目标。

迁移阶段每项功能的最低质量门槛和统一状态见 `DEVELOPMENT_GOAL.md` 第八、九章。阶段验收不得清除或重复运行已有有效证据；数据损坏、错误覆盖、公式错误及保真问题发现后立即处理，不延期到最终阶段。
