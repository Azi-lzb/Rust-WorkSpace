# Excel、OOXML 与历史 Python/Office 边界

> 本文件保留 Open-Data-Audit 的 OOXML 结构保护和旧 Python 实现背景。当前读取/写入/修改路由以根 `AGENTS.md` 和 [workspace 规则 13](../../rules/13_Rust_Excel读写修改与引擎路由规范.md) 为准；openpyxl 除旧实现/测试基线外，仅在根规则 13 明确限定的隔离 ExcelModifier worker 中作为可选修改后端。

## 1. 决策顺序

当前 Rust workspace 处理 Excel 文件时按以下职责选择：

```text
读取既有工作簿      → 纯 Rust calamine；缺失元数据由 namespace-aware OOXML reader 补足
新建 XLSX           → rust_xlsxwriter
修改既有工作簿      → ExcelEditPlan / ExcelModifier 路由至整批操作已验收的单一引擎
本轮公式重算        → 仅用户显式触发时，经独立 sidecar/lo-convert 调用 LO/COM
同类格式转换        → 仅用户显式触发时，经独立 sidecar/lo-convert 调用 LO/COM
```

OOXML Patch、umya-spreadsheet、Go Excelize 和隔离 openpyxl worker 是修改候选，不以本文件的历史实现或理论能力代替真实模板对拍。打印只写页面设置；PDF、渲染和跨类别转换不在主路径范围内。

## 2. Direct OOXML

`xlsx/xlsm = ZIP + OOXML`。

适合：
- 固定模板；
- 大量固定模板操作且性能瓶颈已测量；
- 只修改明确 XML；
- 需要保留其他部件；
- 性能瓶颈明确。

要求：
- XML 使用 namespace-aware parser；
- 禁止依赖 `ns0:` 等具体前缀的正则解析结构；
- 正确维护 Relationship、Content Types、Sheet graph；
- 删除 Sheet 后校验不存在悬空引用；
- 输出必须通过包结构校验和独立 reader 重开；openpyxl 可作为 Python 对拍基线，生产用途仅限根规则 13 限定的隔离修改 worker。Office/WPS/LibreOffice 仅在获准的 sidecar 功能中使用，不作为一般修改验收的必备运行时。

不适合：
- 工作簿真实计算；
- 未知任意 Excel 的万能编辑器；
- 不能证明结构安全的复杂对象修改。

## 3. openpyxl（对拍基线与可选修改 worker）

openpyxl 不得用于 workspace 的生产读取、新建工作簿、业务计算或公式重算。唯一生产使用方式是根规则 13 定义的隔离 ExcelModifier worker：一个 workbook 的完整 ExcelEditPlan 由内部绝对路径启动，不依赖系统 Python/PATH，不导入 Audit 业务代码；只有格式、对象、保留项、平台和整批操作都具有匹配的验收证据时才可路由。worker 缺失时只允许写入前拒绝或转到另一已验收引擎；执行失败、超时或输出验证不通过时停止，不掩盖错误。

worker 不保证保留所有 Excel 对象；当前不支持行插入/删除，并会在写入前拒绝非空公式缓存和未验证 package parts。产品 deb 若启用 worker，只能内置隔离 Python/openpyxl runtime，不依赖系统 Python、用户 PATH、pip 或网络。生产读取仍使用 calamine/已验收的 OOXML 专用 reader，新建工作簿仍使用 rust_xlsxwriter。

旧项目对拍覆盖范围：
- 单元格；
- 工作表；
- 公式文本；
- 样式；
- 命名区域；
- 静态规则；
- 不涉及真实计算的模板处理。

旧项目 openpyxl 行为基线：
- 可以读写公式文本；
- **不会真实计算公式**。

## 4. Office / LO / COM sidecar 边界

### Windows
- Excel
- WPS

### UOS/Linux
- LibreOffice Calc

只有两类操作允许显式调用独立 sidecar：公式重算、同类格式转换。其他 Office/WPS/LibreOffice/COM/UNO 语义不得进入主路径；原生渲染不是可用旁路范围。旁路缺失时该项报告“该功能需要外部引擎”，其他功能继续运行。

## 5. 工作簿计算硬边界

禁止：
- Python 自制不完整公式引擎冒充 Office；
- openpyxl 冒充公式计算；
- 仅依赖旧 cached value 当作本轮计算结果。

公式重算 sidecar 流程必须：
1. 打开；
2. 重算；
3. 保存；
4. 可靠关闭；
5. 由 Rust reader 检查输出和缓存刷新状态；Python 仅可用于对拍基线。

## 6. Fast OOXML

旧实现 `systems/s3_central_statistics/fast_ooxml.py` 可作为算法和保留行为参考；workspace 中如需该能力，应在 `crates/ooxml-patch` 用 Rust 重写并对拍。

必须：
- workbook.xml / rels 用 XML parser；
- Relationship target 用 POSIX package path 解析；
- 输出前检查结构完整性；
- 用户错误说人话；
- 技术日志保留 rId、Target、part 等诊断信息。

## 7. Sheet 删除

Direct OOXML 物理删除 Sheet 时至少同步考虑：
- workbook.xml
- workbook.xml.rels
- `[Content_Types].xml`
- sheet rels
- definedNames
- bookViews
- 相关 drawing/table/comment 等依赖

若保留 Sheet 引用了待删除 Sheet：
- 不得静默制造 `#REF!`
- 自动清理空 Sheet 且存在未解除依赖时，保留该 Sheet（及仍被引用的名称）并隐藏；明确要求硬删除时阻止操作并报告依赖
- 不自动改业务公式

### 7.1 工作表身份与顺序字段（铁律）

- SheetName / rId / worksheet part 用于识别「删的是谁」（工作表身份）；
- old_index / new_index 用于维护依赖工作表顺序的字段
  （`localSheetId`、bookViews 的 `firstSheet`/`activeTab`）；
- **`localSheetId` 不是工作表身份**，它只是 workbook.xml 当前顺序中的位置，
  禁止把它当身份用于判定，也禁止用逐次 `-=1` 的写法回移（一次删除多张表
  会累计偏差）；必须先建立删除集合的 old→new 映射，再一次性重算。

### 7.2 Sheet-local definedNames（铁律）

Direct OOXML 删除 Sheet 时必须维护 Sheet-local definedNames
（`_xlnm._FilterDatabase` / `_xlnm.Print_Area` / `_xlnm.Print_Titles`）：

- 被物理删除 Sheet 的 local definedName（owner=被删表）应同步删除；如果保留公式仍引用该名称，不物理删除 Sheet，按自动清理规则隐藏并保留名称；
- 属于保留 Sheet 的 localSheetId 必须根据新的 Sheet 顺序重映射；
- 应结合 `localSheetId` 对应的 owner 与名称公式文本中的 SheetName 做
  交叉验证；两者矛盾视为作用域/引用异常，不得自动猜测改写；
- 无法证明安全的复杂 definedName（自定义名称、跨表复杂引用、外部引用、
  动态公式）必须保守处理：与被删表无关则保留；疑似引用待删空表时，
  自动删空表场景优先把该空表降级为隐藏，不得生成可能损坏的工作簿；
- 引用「明确业务删除」表（如内嵌配置页）且无法挽留的名称一并移除并记日志；
- 用户提示不得出现 localSheetId / definedName / rId / part 等底层术语；
  技术日志保留 defined_name / owner / old/new localSheetId / action / reason。

### 7.3 definedName 公式依赖保护（铁律）

Direct OOXML 删除工作表及其 definedName 前，必须确认最终保留工作表没有继续引用该名称；无法证明安全时不得物理删除，也不得静默制造 `#NAME?`。处置方式固定为：自动清理场景保留工作表和依赖名称并隐藏；用户明确要求硬删除时阻止并报告，由用户决定如何处理依赖。

- 名称引用检测按 Excel 名称语义做 token 边界 + 大小写不敏感匹配
  （``Rate`` 不得误判 ``InterestRate``；``myconfigvalue`` 命中
  ``MyConfigValue``），禁止裸 substring；
- 无人使用的名称可随表安全删除；自动清理时若有公式引用名称，保留工作表和名称并隐藏；明确业务硬删时阻止并报告；
- 绝不自动重写业务公式（改引用目标、内联名称都属于业务公式重写）；
- 保留表公式识别不到时宁可保守，不得产出 #NAME? 工作簿。

删除后的完整性验证不得降低标准：sheet↔relationship↔part 无悬空、
被删表 part 不再被引用、被删表的 local definedName 不存在、保留表
localSheetId 与新顺序一致且不越界、剩余名称不指向已删除的 Sheet、
输出必须能被 workspace 独立 reader 重开，并通过 OOXML 关系完整性检查；openpyxl 重开只作为对拍证据。若复用 FastOoxmlCompiler，其实现必须是 workspace 内经过对拍的 Rust 适配器，而非旧 Python 模块。

## 8. 条件格式

规则是否触发与显示颜色不是一回事。

静态条件格式求值只有在 Rust 实现且规则范围经过对拍后才可用于业务结果；不支持的规则标记为 UNSUPPORTED，不得伪装为 FALSE。当前 workspace 不提供真实显示渲染或 PDF 验收路径。

统一状态：
- TRUE
- FALSE
- UNSUPPORTED
- ERROR

不支持的规则不得伪装成 FALSE。
