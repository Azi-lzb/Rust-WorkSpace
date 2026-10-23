# 功能级数据流矩阵（feature-dataflow-matrix）

> 审计日期：2026-10-23。编号与 [feature-inventory.md](feature-inventory.md) 一致；引擎评估见 [excel-engine-routing-review.md](excel-engine-routing-review.md)，迁移状态细节见 [migration-gap-analysis.md](migration-gap-analysis.md)。
> 每项功能记录 14 个字段：①名称/项目 ②用户入口 ③输入与来源 ④实际读取方式及位置 ⑤解析/标准化/业务处理 ⑥是否读取原始样式 ⑦是否修改已有工作簿 ⑧修改引擎及位置 ⑨是否新建工作簿 ⑩写入引擎及位置 ⑪输出与保真要求 ⑫语言角色 ⑬跨语言/跨进程调用 ⑭异常处理/临时文件/原文件保护。
> 源项目路径缩写：`PyC`=`4.pytools-main/PyCode`，`NoCom`=`4.pytools-main/NoComPyTools/nocompytools`，`ODA`=`3.open-data-audit/core/src`。行号经实地核对或探索代理核对，无法确认处标**待验证**。

## 1. 通用事实（适用于全部功能，避免逐条重复）

- **语言角色**：源产品 pytools/ODA 业务逻辑全部为 Python；ODA 的 S1-F01/F02、S3-F01 结果写出已嵌入 Rust writer 子进程（§3.4 反向依赖）。Rust workspace 产品路径零 Python 依赖（无 pyo3；`Command::new` 生产调用仅 3 处：`pytools-core/src/conversion.rs:155` sidecar、`quick.rs:67-71` 系统打开文件、sidecar 内部 soffice/powershell）。
- **公式口径**：源产品全部功能从不重算公式——读取依赖文件已存缓存值（openpyxl `data_only=True` / COM Value2 / calamine 存储值）；ODA COM 会话启动即 `Calculation=xlCalculationManual`（`ODA/base_audit/com_engine` 经 `engines/excel_wps.py`），唯一重算入口是 S1 计算引擎（Excel/WPS COM 或 LibreOffice，`engines/__init__.py:99-105`）。
- **并行**：pytools 全部批量为 tkinter 多选后**单线程串行**；ODA 用 threading 跑后台任务（`web_app.py::start` @3202）但工作簿处理仍串行。Rust workspace 业务批量用 rayon 文件级并行（如 `institution_normalize.rs`）。
- **打印页面设置写入基准**：见 feature-inventory §2.3 表后说明。

## 2. pytools 数据流（PT-1-1 … PT-SV1）

### PT-1-1 表头路径比对
① 時序-表头路径比对（pytools，两变体） ② 菜单 1-1：`NoCom/main.py:875-879`、`PyC/main.py:1657-1661` ③ 模板 xlsx/xlsm + 多源 xlsx/xlsm（PyCode 另允许 .xls），tkinter 多选 ④ NoCom：pandas→calamine（`NoCom/core/io_excel.py:63-67`，异常回退 openpyxl @74-77），实际打开在 `pd.read_excel` @110/103；PyCode：pandas→openpyxl/xlrd（`PyC/core/io_excel.py:63-69`）；COM 版 `compare_393_com.py` 经 `com_engine.read_sheet_values`（UsedRange+Value2） ⑤ `compare_393.py:134-187` list_sheet_names→read_sheet_2d 逐 sheet 表头路径比对；合并区经 `rule_engine._load_merge_ranges` ⑥ 否（仅合并区坐标） ⑦ 否 ⑧ — ⑨ 是 ⑩ pandas ExcelWriter(engine="openpyxl")（`io_excel.py:88-97`） ⑪ 新建 2-sheet 比对报告 xlsx，纯值 ⑫ 读 Python(NoCom=calamine/PyCode=openpyxl)，算 Python，写 Python ⑬ PyCode 下 .xls 经 COM（进程内 Dispatch） ⑭ 读缓存 lru_cache（路径+mtime+size+sheet，`io_excel.py:108-120`），单文件处理完 `clear_sheet_cache`；无临时文件（直接写目标路径）

### PT-1-2 时序提取（快速精简）
① 菜单 1-2 ② `NoCom/main.py:477-491`、`PyC/main.py`（"2com" @752-779） ③ 多源工作簿 + 配置"时序提取规则" ④ 同 PT-1-1 读引擎；合并区来源差异：NoCom=python-calamine `merged_cell_ranges`（`rule_engine.py:151-163`，异常回退 openpyxl @164-177）；PyCode=xlsx/xlsm 用 openpyxl `merged_cells`（@147-160），**.xls 回退 COM** `_load_merge_ranges_file_by_com` @164-223（MergeAreas，失败逐格扫描，异常静默返回 {}） ⑤ `rule_engine.extract_cells_from_arrays`（`rule_engine.py:428-633`，表头合并区取左上 @192-199） ⑥ 否 ⑦ 否（新建/追加目标簿，不改源） ⑧ — ⑨ 是（新建或追加已有簿） ⑩ 新建：pandas ExcelWriter(openpyxl)；追加：`io_excel.append_to_target` @161-291——读旧 sheet（`pd.read_excel` @208-212）后 `mode="a", if_sheet_exists="replace"` 重写目标 sheet（@278-284，**重建替换非保真追加**） ⑪ 值 + "数据日期"列 number_format="yyyy/m/d"（`_apply_data_date_format` @24-34）；追加时目标 sheet 其余格式丢失（现状契约） ⑫ 全 Python ⑬ PyCode .xls 经 COM ⑭ 同 PT-1-1 缓存策略；追加直接写目标路径，无原子提交

### PT-1-3 / PT-1-4 宽表汇总（加/不加后缀）
① 菜单 1-3/1-4 ② `NoCom/main.py:492-509`、`PyC/main.py`（"3com"/"4com" @780-816） ③ 多源工作簿 + 规则 ④ 同 PT-1-1 ⑤ `wide_397_398.py:164-194` 读→`rule_engine.compute_wide_col_paths_from_arrays`（`rule_engine.py:314-425`，列头路径+同名消歧） ⑥ 否 ⑦ 否 ⑧ — ⑨ 是（每规则一 sheet @`wide_397_398.py:322`，可扩列追加 `append_to_target(allow_extend=True)`） ⑩ pandas ExcelWriter(openpyxl)；追加同 PT-1-2 重建替换 ⑪ 值+表头，无样式要求 ⑫ 全 Python ⑬ PyCode .xls COM ⑭ 同 PT-1-1

### PT-1-5 预校验
① 菜单 1-5 ② `NoCom/main.py:457-459`、`PyC/main.py:701-703` ③ config.xlsx ④ pandas `pd.ExcelFile/pd.read_excel`（`precheck_399.py:24-50`） ⑤ 配置完整性校验 ⑥-⑨ 否 ⑩ — ⑪ 控制台报告+建目录，无文件输出 ⑫ Python ⑬ 无 ⑭ 只读

### PT-2-1 按批注检查重复
① 去重-按批注检查 ② `NoCom/main.py:510-516`、`PyC/main.py:850-856` ③ 单工作簿（xlsx/xlsm） ④ **openpyxl** `load_workbook(data_only=False, keep_vba)`（`dedup_311.py:94-96`）；A1 批注含"源数据"判定命中 sheet（@176-177），表头批注含"标识列"定键列（`_get_key_cols_by_comment` @112-120） ⑤ 按键列分组找重复行 ⑥ **是（批注为功能核心输入）** ⑦ **是（原位标红）** ⑧ openpyxl 写红色 PatternFill（@17,143）后 `wb.save(workbook_path)` @185 ⑨ 否 ⑩ — ⑪ 原文件其余样式不动；重复行着色为输出契约 ⑫ Python ⑬ 无 ⑭ 直接覆盖保存原文件，**无备份/无临时文件**

### PT-2-2 按批注删除重复
② `NoCom/main.py:517-523`、`PyC/main.py:857-863` ④ 同 PT-2-1 ⑤ 倒序 `ws.delete_rows`（`dedup_311.py:160-161`） ⑥ 是（批注） ⑦ 是 ⑧ openpyxl 原位删行后保存 @210；**删行不修正合并区/批注位置**（现状语义） ⑨ 否 ⑪ 同上 ⑭ 直接覆盖保存

### PT-2-3 / PT-2-4 按配置检查/删除重复
② `NoCom/main.py:524-539`、`PyC/main.py:864-889` ③ 配置簿（去重追加数据配置）+ 目标簿 ④ openpyxl；同配置分组同簿只开一次（`run_check_by_config` @215-250，变更才 `wb.save` @249） ⑤ 按配置键列重复检测/删除（`run_delete_by_config` @259+） ⑥ 是 ⑦ 是 ⑧ openpyxl（同上） ⑨ 否 ⑩ — ⑪ 未触碰 sheet 保留 ⑫ Python ⑬ 无 ⑭ 条件保存（无变更不写盘）

### PT-2-5 按配置去重追加到目标
③ 配置 + 源簿 + 目标簿 ④ openpyxl（源 data_only=False——**公式原文会被当作值搬运**） ⑤ `_append_unique_source_to_target`（`dedup_311.py:368`）唯一键过滤后值拷贝 ⑥ 否 ⑦ 是（目标簿） ⑧ openpyxl 原位追加行 ⑨ 否（除非目标不存在则新建） ⑩ openpyxl ⑪ 追加幂等性、顺序为目标契约 ⑫ Python ⑬ 无 ⑭ 直接保存

### PT-2-6 按配置预校验
④ `_open_wb(read_only=True)`（@553）只读 ⑤-⑪ 仅报告 ⑭ 无写

### PT-3-1 按批注打印（保留源格式）
① 打印-保留源格式 ② `NoCom/main.py:540-546`、`PyC/main.py`（"p1com" @934-959） ③ 源工作簿（A1 批注含"打印/输出"命中 @391/465；块边界由单元格批注"打印区域N/#N"标注，`_find_comment_ranges` @171-198） ④ **同一文件 openpyxl 双开**：样式视图 `_open_wb`（`print_310.py:61-62`）+ 缓存值视图 `load_workbook(data_only=True)`（@452-457） ⑤ `_copy_block` mode=1 @244-254：值+font/fill/border/alignment/protection/number_format 逐属性复制+批注（可配 drop_comments 清除 @253-254）+行高列宽+合并重建 @264-288 ⑥ **是（样式+批注全量）** ⑦ 否 ⑧ — ⑨ **是** ⑩ openpyxl `Workbook()` @458 ⑪ 值+样式+批注+行高列宽+合并+页面设置（`_apply_print_setup` @121-147）+顺序页脚 @150-156；输出名 `{ts}_按批注打印（保留源格式）.xlsx` @491 ⑫ Python ⑬ PyCode .xls 经 COM（PasteSpecial 复制+COM 页面设置） ⑭ `_save_with_suffix_fallback` @41-54：PermissionError 自动换 `_1/_2…` 后缀重试（≤1000 次），不覆盖被占用文件；**无临时文件+验证机制**

### PT-3-2 按批注打印（快速复制）
② `NoCom/main.py:547-553`、`PyC/main.py`（"p3com" @960-985） ④ 值=calamine（`_read_values_calamine` @201-203→read_sheet_2d），样式仍需 openpyxl 打开一次（**双引擎双开**） ⑤ mode=3 @256-307：font/fill/border/alignment/number_format+合并；**不复制批注/保护/行高列宽** ⑥ 部分（基础样式，无批注） ⑦ 否 ⑨ 是 ⑩ openpyxl @417 输出 ⑪ 值+框线+合并+页面设置 ⑫ Python（读混 calamine+openpyxl） ⑬ PyCode COM 版同口径 ⑭ 同 PT-3-1

Rust 候选（2026-10-23，合成 XLSX）：CLI `print-comment-plan` 经共享 `excel-reader` 识别标记区域并构造批量导入/页面设置操作；执行候选为 `ExcelModifier → UmyaSpreadsheet`，通过测试临时 capability 写新工作簿。Keep 与 fast 输出均与只读 Python 基线一致（值/类型、样式属性、批注策略、合并、行列尺寸、页面设置、打印区域、页脚）；列宽按 OOXML 列范围展开为每列有效宽度比较，浮点容差为 `1e-15`。执行对拍报告：`pytools-core/probe/reports/print_comment_execution_parity_20261023.json`；测试：`pytools-core/tests/print_comment_execution.rs`。源文件未改变。此候选仅覆盖合成 XLSX，生产能力未登记、默认路由未改变，真实模板和目标平台仍未验证；由于输出是新工作簿，生产采用前仍须评估 `rust_xlsxwriter` 新建路径。

### PT-3-3 按配置打印（全部模式）
② `NoCom/main.py:554-566`、`PyC/main.py`（"p7com" @986-1004） ③ 打印配置 + 源簿 + 目标簿 ④ 四级缓存（source/source_value/source_calamine/target，@595-598）避免重复打开；mode3 值走 calamine ⑤ `run_print_config_all_modes` @593-815：块复制（mode1 全样式同 PT-3-1；mode2 值+基础样式+自适应列宽行高 @309-364；mode3 同 PT-3-2） ⑥ 是 ⑦ **是（修改已有目标簿）** ⑧ openpyxl：目标存在则打开、清空目标 sheet（`delete_rows(1,max_row)` @740-742）后写入，`wb.save` @797 ⑨ 条件（目标不存在则 `Workbook()` @727） ⑩ openpyxl ⑪ 块内容+页面设置；目标簿其他 sheet 保留（现状） ⑫ Python ⑬ COM 版 `print_310_com.py::run_print_config_all_modes_com`（`_apply_print_setup_com` @245-276，Excel 内存簿+Save） ⑭ 同 PT-3-1 后缀回退

Rust 候选（2026-10-23）：CLI `print-config-plan` 加载配置后按配置目标簿分组构造 `ExcelEditPlan`；源数据经共享 `excel-reader` 读取，整批修改经 capability 门控的 `ExcelModifier → UmyaSpreadsheet` 执行，计划阶段不输出工作簿。按用户确认的新合同，配置目标必须是新路径，已有目标拒绝；实际执行写入配置目标路径。合成单任务模式 1/2/3 结果与旧 Python 投影一致，含值/类型、列出的样式属性、批注、合并、行列尺寸及页面设置，浮点列宽按相对/绝对容差 `1e-15` 比较。测试专用 Passed capability 未进入产品能力登记。真实配置和业务模板、同目标多任务组合、Windows 原生构建及目标 Linux/deb 未验证；状态仍为未验证。

### PT-3-4 按配置打印预校验
④ `_open_wb(read_only=True)` @574，同簿缓存 @570-575 ⑤ 配置与批注协议校验 ⑦-⑨ 否 ⑭ 只读

### PT-3-5 打印→PDF（范围排除）
② `PyC/main.py:1005-1026`（"ppdf"；非 .xls 任务提示不支持 @1010-1011） ⑩ COM 内存簿 → `ExportAsFixedFormat`（`print_310_com.py:813-819`） ⑬ COM 渲染 ⑪ PDF 文件。NoComPyTools 无此项。

### PT-4-1 批量重命名
② `NoCom/main.py:571-577`、`PyC/main.py:1027-1035` ③ 配置簿 + 文件/目录 ④ pandas 只读配置（简称/全称、全称/代码、数据日期/报表名） ⑤ 不区分大小写子串匹配，按配置首次插入顺序；缺代码/无命中/原名不变/路径不存在计入跳过 ⑥-⑦ 否 ⑧ — ⑨ 否 ⑩ — ⑪ 仅文件名变化，内容字节不变 ⑫ Python ⑬ 无 ⑭ 冲突目标 `_next_path` 加序号不覆盖

### PT-4-2 Excel 格式转换
② `NoCom/main.py:578-590`、`PyC/main.py:1036-1054` ③ .xls（NoCom 只收 .xls @59；PyCode 支持 xls/xlsx/xlsm/xlsb/xlt/xltx/xltm/et/ett/csv 互转） ④ 引擎按平台：NoCom Windows=Excel/WPS COM `DispatchEx("Excel.Application"/"ket")`（`xls_to_xlsx.py:27-29`，`Workbooks.Open+SaveAs FileFormat=51` @44-51）；NoCom Linux=LibreOffice headless（`libreoffice.py:32-58`，soffice --convert-to，timeout 180s）；PyCode=COM-only（`_excel_convert_one` @1017-1022，8 种 FileFormat，.et 优先 WPS @1055-1056） ⑤ 引擎转换，无业务处理 ⑥ 由 Office 引擎保真 ⑦ 否 ⑧ — ⑨ 是（转换输出） ⑩ Office SaveAs ⑪ `源目录/xlsx/同名.xlsx`（@63-65），要求 sheet/数据保真 ⑫ Python 编排 + Office/LO 执行 ⑬ **COM/子进程 soffice** ⑭ LibreOffice 路径经临时目录+`shutil.move`（@35-58）

### PT-4-3 Word 格式转换
② `NoCom/main.py:591-600`、`PyC/main.py:1055-1072` ③ doc/docx ④ NoCom Windows=Word/WPS Writer COM（`SaveAs2` FileFormat 0/12 @49-63）；Linux=LO Writer（`_WORD_FILTERS` @10-13）；PyCode COM-only（ProgID 链 @1077-1089） ⑤ 同格式跳过 @79-82 ⑩ Office/LO SaveAs ⑪ `源目录/{目标格式}/` 子目录 @83-84 ⑬ COM/soffice ⑭ 同 PT-4-2

### PT-4-4 批量修改 Sheet 名
② `NoCom/main.py:601-607`、`PyC/main.py:1083-1107` ④ openpyxl keep_vba @852 ⑤ `ws.title=` 映射 @869/874（仅 xlsx/xlsm） ⑦ **是** ⑧ openpyxl 原位改 sheet 名并保存；PyCode .xls 走 COM `Worksheets(idx).Name=`+`wb.Save()`（@1644/1652，Excel 重存保真） ⑪ sheet 名变化，其余对象依赖引擎保真（openpyxl 重写全簿） ⑬ PyCode .xls COM ⑭ 直接覆盖保存

### PT-4-5 提交缺失检查
② `NoCom/main.py:631-637` ④ pandas 只读配置 ⑤ 文件名关键词比对 ⑨ 是（pandas 报告 xlsx，`run_submission_check_48` NoCom @198-228） ⑪ 报告 ⑭ 无

### PT-4-6 按内容加前缀重命名
② `NoCom/main.py:608-615`、`PyC/main.py:1108-1122` ③ 配置簿（重命名配置!O:P:Q）+ 文件 ④ NoCom：calamine 优先 `_read_cells_calamine`（`convert` @776-801，`to_python(skip_empty_area=False)`），OOXML 边角回退 openpyxl @819-835；PyCode：openpyxl @1295-1311，.xls/.xlt COM Value2 @1314-1332 ⑤ 读指定单元格拼前缀（缺失 sheet 回退 `sheet!A1`；前缀上限 180 字符）；内容项已在文件名中则跳过 ⑥ 否（仅单元格值） ⑦ 否（**不保存工作簿**） ⑨ 否 ⑩ — ⑪ 仅文件名 ⑫ Python ⑬ PyCode .xls/.xlt COM ⑭ 冲突 `_next_path` 序号

### PT-4-7 取消内容前缀
② `NoCom/main.py:616-623`、`PyC/main.py:1123-1137` ④ 同 PT-4-6 ⑤ 移除文件名开头内容项，兼容保留排序前缀 @920-922 ⑦ 否 ⑭ 同上

### PT-4-8 局部映射重命名
② `NoCom/main.py:624-630`、`PyC/main.py:1138-1147` ④ pandas 读 M:N 配置 ⑤ 按配置顺序 stem 片段替换；空目标=删除片段 ⑦ 否 ⑭ 冲突序号

### PT-4-9 / PT-4-10 文件排序前缀 / 取消
② `NoCom/main.py:638-644`/`645-651`；`PyC/main.py:1162-1175`/`1176-1188` ④ 排序规则配置（4-10 另可读内容格） ⑤ 规则匹配改名 + 报告簿（4-9 `run_file_sort_prefix_49` NoCom @286-334；4-10 @336-405，直接以排序前缀开头的不打开工作簿） ⑥ 4-10 部分（可选内容单元格值） ⑦ 否 ⑨ 是（报告 xlsx，pandas） ⑩ pandas/openpyxl ⑪ 报告+文件名集合 ⑬ PyCode .xls/.xlt 内容读取 COM（仅源基线） ⑭ 冲突不覆盖

### PT-5-1 初始化/修复配置表
② `NoCom/main.py:420-424`、`PyC/main.py:661-669` ③ config.xlsx（11 个业务 sheet） ④ **openpyxl 原位打开** `load_workbook(cfg_path)`（`config_init.py:306`） ⑤ 修表头、补缺失 sheet/单元格/说明批注（`Comment(text,"pytools")` @188-196，可清除旧批注 @213-214；`_ensure_*` 系列各 sheet 检查） ⑥ 是（批注+样式是修复对象） ⑦ 是 ⑧ openpyxl 原位补齐后保存 ⑨ 条件（缺失则新建） ⑩ openpyxl ⑪ 表头/批注/默认值与使用指南 @352-366 字段一致 ⑭ 直接保存

### PT-5-2 查看当前配置
④ pandas 只读 ⑪ 控制台打印 ⑭ 无写

### PT-6-1 按使用区域汇总
② `NoCom/main.py:703-712`（选择器含 .xls/.xlsb）、`PyC/main.py`（"s21com" @1268-1289） ④ NoCom：**CalamineWorkbook 直连**（`summary_321_322.py:11` import；先全表扫列宽 @217-221，再逐文件重新打开 @240，`sheet.to_python(skip_empty_area=False)` @251——**同批文件打开两遍**）；PyCode：openpyxl data_only @198/224；COM 版 Value2 ⑤ used-range 全量聚合一表 ⑥ 否 ⑦ 否 ⑨ 是 ⑩ openpyxl `Workbook()` @203，表头粗体 @232，freeze A2 @271 ⑪ `{ts}_按使用区域汇总（全量）.xlsx` @272，值+粗表头 ⑫ Python（NoCom 读 calamine） ⑬ PyCode COM ⑭ 无专门保护

### PT-6-2 按批注汇总
② `NoCom/main.py:847-860`、`PyC/main.py`（"s22com" @1290-1326） ③ 模板簿 + 源文件 ④ 模板批注解析=openpyxl（`_extract_comments` @94-100、`_extract_regions` @124-148、`_extract_set_info` @151-159——行/列区域N/#N 与 set(…) 指令是**功能核心输入**）；源值 NoCom=calamine `to_python` @351 + `merged_cell_ranges or ()` @352（None 与"无合并"混淆——已登记能力缺口）；PyCode=openpyxl @330 ⑤ 按批注区域指令聚合；合并取左上 `_gcv` @84-91 ⑥ **是（模板批注+合并）** ⑦ 否 ⑨ 是 ⑩ openpyxl @409-413（表头粗体+freeze A2） ⑪ `{ts}_按批注汇总（模板+源文件）.xlsx` @413 ⑫ Python ⑬ COM 版源 Value2；**.xls 模板先 COM SaveAs=51 转临时 xlsx** 供 openpyxl 解析批注（`_ensure_openpyxl_compatible` @47-71，`tempfile.mkdtemp(prefix="3_2_2_xls_")` @245） ⑭ 临时目录转换；输出直接保存

### PT-7-1 拆分村镇银行数据
② `NoCom/main.py:713-735`（三选：本期/去年同期/待改目标）、`PyC/main.py:1346-1390` ④ 读两期 `load_workbook(data_only=True, read_only=True)`（`feature_1_1_split_village_bank.py:47-48`） ⑤ 计算 6×8 矩阵 @43-77 ⑥ 否 ⑦ **是** ⑧ 目标 keep_vba 打开 @108→`delete_rows+insert_rows` @124-125→写值→`wb.save(path)` @131；**openpyxl 删插行不修正合并区**（COM 版 Excel 原生删插会随行移动） ⑪ 行结构+值 ⑬ PyCode .xls COM 版 ⑭ 直接覆盖保存目标

### PT-7-2 机构名称标准化
② `NoCom/main.py:736-748`、`PyC/main.py:1391-1418` ④ 映射配置（机构映射表 A→B，openpyxl/pandas）+ 目标簿 keep_vba @80 ⑤ A 列机构名精确/子串映射替换 ⑥ 否 ⑦ 是 ⑧ openpyxl 原位替换后保存 ⑪ 值替换，其余保留 ⑬ 无 ⑭ 直接保存

### PT-7-3 删除分机构表外资行
② `NoCom/main.py:749-762`、`PyC/main.py:1419-1447` ④ 映射配置（C=1 标记外资）+ 目标簿 ⑤ **先 `shutil.copy2` 生成"{stem}(金融局)"副本**（`feature_1_3_remove_foreign_bank.py:55-62`），对副本操作 ⑥ 否 ⑦ 是（副本） ⑧ 连续块自底向上批量 `delete_rows` @39-52，openpyxl 保存 @73 ⑪ 副本行结构 ⑬ 无 ⑭ **原文件不动**（唯一有副本保护的原位改功能）

### PT-7-4 外汇页眉修改
② `NoCom/main.py:763-769`、`PyC/main.py`（"x14com" @1497-1520） ④ openpyxl（对 .xlsm keep_vba）遍历名称含"外汇"的 sheet ⑤ odd/even/first 三套页眉 left/center/right 文本"万元"→"万美元"（@23-31） ⑥ 是（页眉文本） ⑦ 是 ⑧ openpyxl 改页眉后 `wb.save(path)` @48（仅至少一处变化时保存）；COM 版额外感知 `DifferentFirstPageHeaderFooter`（`feature_1_4_fx_header_com.py:62-64`） ⑪ 页眉文本，其余保真 ⑬ PyCode COM ⑭ 直接保存

### PT-7-5 地区总分校验
② `NoCom/main.py:770-778`、`PyC/main.py:1459-1480` ④ 只读 `load_workbook(data_only=True, read_only=True)`（@172） ⑤ 地区合计 vs 总分容差核对 ⑥ 否 ⑦ 否 ⑨ **是** ⑩ openpyxl `Workbook()` @163，写字体/填充/对齐/自适应列宽（@122-156），输出 `地区数值核对结果_{ts}.xlsx` @206 ⑪ 报告样式是输出契约一部分 ⑬ 无 ⑭ 输出名冲突无保护（现状）

### PT-7-6 工作表数据提取（三档实现）
② `NoCom/main.py:779-791`；`PyC/main.py`（"x16com" @1521-1558） ④ NoCom 纯值版：`read_sheet_2d`（pandas+calamine）@450；PyCode 普通版：openpyxl；COM 批注版：Value2 批量读 @86-99 / `src_sheet.Copy` @140 ⑤ 按配置"工作表提取"关键字/区域选择；整表模式走 `_copy_full_sheet_with_style`（PyCode @381-413：值+`_style`/number_format/protection/alignment+行高列宽/隐藏/outlineLevel+合并，**不复制批注**）；NoCom `_write_values` 只写值 @358-362 ⑥ PyCode 普通版是（样式+尺寸）；NoCom 否；COM 版是（全对象含批注，**不迁移**） ⑦ 否（源只读） ⑨ 是（新建或追加输出簿 @441-443） ⑩ openpyxl，`wb_out.save(path)` @497 ⑪ 输出 sheet 名/区域值/顺序；样式档位按变体不同 ⑫ Python ⑬ PyCode .xls COM 版（SaveAs 51 @323） ⑭ 每文件处理后 `clear_sheet_cache` @486；输出直接保存

### PT-7-7 / PT-7-8 涉农贷款修正（月度/季度）
② `NoCom/main.py:792-816`/`817-846`；`PyC/main.py:1559-1591`/`1592-1633` ④ 本期 `load_workbook(keep_vba)`（`feature_1_8_adjust_rural_loan.py:126`/`feature_1_9:110`）可写；上期 `load_workbook(data_only=True, read_only=True)` @127/111 ⑤ 分机构表逐机构写差额列 D @148-156（**覆盖公式为值**——现状）；季度版回填全部地区分机构表后核对 ⑥ 否 ⑦ 是（本期簿） ⑧ openpyxl 原位写值后 `wb_cur.save(current_path)` @168 ⑨ 否 ⑪ 差额列数值+容差核对 @174（超出容差报告） ⑬ PyCode 上期/本期为 .xls 时走 COM ⑭ 直接覆盖保存

### PT-8-1 … PT-8-5 归档（五项同管道）
② `NoCom/main.py:652-702`；`PyC/main.py:1189-1257` ③ 业务目录文件 + 可选"归档类型配置" ④ 文件名正则（`_archive_date_from_name` convert NoCom @145-160）/配置关键词分类；递归 rglob（`_pick_archive_paths` main.py:229-263） ⑤ 生成归档计划（8-1/8-2 仅预览）或执行 `shutil.copy2`（8-3 `run_archive_plan_copy` @549-597；8-4/8-5 预览+直接复制），后缀黑白名单 @120-136，未识别进兜底目录 ⑥ 否 ⑦ 否 ⑨ 仅计划/结果报告 xlsx（pandas） ⑩ pandas/openpyxl ⑪ 报告+复制结果 ⑫ Python ⑬ 无 ⑭ copy2 不覆盖语义按 Python 默认（覆盖已存在目标——现状）；**待验证**：冲突时具体行为两变体是否一致

### PT-9-1 … PT-9-3 快捷工具
② `NoCom/main.py:429-456`；`PyC/main.py:673-700` ④ 无 Excel 读取 ⑤ output 目录 mtime 最新文件定位（@299-305） ⑬ NoCom：win32 `os.startfile` / 其他平台 `subprocess.Popen(["xdg-open"])` @290-297；PyCode 仅 `os.startfile` @502-505 ⑭ 无

### PT-SV1 调研选项统计（未接线）
`survey_stats.py`（636 行）：openpyxl 读问卷+config → 明细/展示两 xlsx。入口缺失（feature-inventory §2.10）。调用链**待验证**（无法从菜单到达，行为以静态阅读为准）。

## 3. ODA 数据流（AUD-S1 … AUD-C07）

### AUD-S1-F01 汇总核查表校验
① ODA-S1 ② 前端"开始审核"→ `web_app.py::start("audit")` @3195-3208（worker 线程 @3202）；审核前检查 `start("check")`→`service.preflight` @451-473 ③ 模板目录 + 源核查表目录（前端选择）+ config-real 配置 ④ 双管线：Windows COM 管线=ExcelSession（`excel_com.py`，模板打开/检查/规则执行/条件格式与 issue 读取）；native 管线=openpyxl（`native/*`）+ soffice 重算；管线选择 `engines/__init__.py:99-105`（win32→"com"）；模板匹配/源发现 `discovery.py`（recommend_template 等） ⑤ 条件格式规则求值（`conditional_engine.py`，COM 逐格触发/优先级/相对引用）、公式规则、结果模型 `models.py` ⑥ **是（条件格式、批注规则、命名区域、公式结果是核心输入）** ⑦ 源文件可选回写/导航链接写入（`workflow/adapters/audit_handlers.py::handle_navigation_write` @897） ⑧ COM ExcelSession / native openpyxl+OOXML（`conditional_context.py`、`cf_reader.py` 读元数据） ⑨ **是** ⑩ **Rust writer 子进程**（rust_xlsxwriter，`native/s1_summary_writer.py:267-303`，临时文件 `.rust-audit-*.tmp.xlsx` @267 + openpyxl 回退；用户可选 AUTO/OPENPYXL/RUST @`web_app.py:434-440`） ⑪ 审核结果工作簿（值+超链接+冻结+筛选+列宽，`test_result_writer.py` 合同）+ 历史记录（`native/history_io.py:67-90` openpyxl 原位追加） ⑫ Python 业务 + Rust 写出 + Office/LO 计算 ⑬ **COM（Excel/WPS）或 soffice 重算 + Rust 子进程 writer** ⑭ 临时文件写结果后落位；single_instance（`single_instance.py`）；取消经 `_cancel_event`（不强杀 Office @3237-3240）

### AUD-S1-F02 汇总校验结果说明
② `start("summary")` → `service.summarize_regions` @183-323 ③ 模板+源目录 ④ 读取引擎三态（`engines/__init__.py:31-34`）：纯 Python=python-calamine AUTO（`native/summary_source_reader.py:17`，openpyxl 回退 @164-177）/Excel-WPS COM/LibreOffice；native 管线走 `native/summary.py::merge_workbook_tables/run_region_summaries`（@257,512,638，result_writer 参数 AUTO） ⑤ 按模板工作表聚合并保持绝对坐标 ⑥ 否（明确不复制公式/条件格式，`engines/__init__.py:29-33` 注释） ⑦ 否 ⑨ 是 ⑩ Rust writer（`s1_summary_writer.py:39-40,105-133`，临时文件 @105）或 openpyxl 回退 ⑪ 汇总说明工作簿+历史说明 ⑫ Python+Rust ⑬ COM/LO 读取（可选）+Rust 写 ⑭ Rust 失败回退 openpyxl 并记录 writer 版本/耗时（@139）

### AUD-S1-F03 组合工作表（核查表）
② DAG combine 预设（`start_dag_flow` @3209）→ `service.merge_org_files` @357 ③ 多机构工作簿目录 ④ COM 管线=ExcelSession 跨簿复制；openpyxl 管线=`merge_org.py`（`_copy_sheet_openpyxl`/`_merge_one_org_openpyxl`/`_copy_sheet_named_ranges_as_local`/`_localize_workbook_named_ranges`/`_write_template_merge_report`） ⑤ 机构分组（文件名首段）、sheet 命名去重、命名区域本地化 ⑥ **是（样式/批注/合并/尺寸/命名区域）** ⑦ 是（新建合并簿过程中的跨簿复制） ⑧ COM / openpyxl ⑨ 是 ⑩ 同引擎 ⑪ 全部 sheet+公式+名称 ⑬ COM 可选 ⑭ 源文件只读

### AUD-S1-F04 组合工作表（按配置）
② DAG 编辑器（@1118-1236）+ `start_dag_flow` ③ 固定 DAG 配置（`workflow/defaults.py`，旧配置不可覆盖固定 DAG——`workflow/test_node_flow_config.py` 合同）+ 输入 ④ `workflow/runner.py`（619 行）按 DAG 调度 `adapters/*` 节点（source_files→template_load→structure_compare→files_copy→external_copy→formula_copy→…→result_workbook） ⑤ 节点式编排，失败策略（fail-workflow） ⑥ 节点相关 ⑦/⑨ 按节点 ⑩ 同 S1 管线 ⑬ COM/native 适配器二选一 ⑭ artifact store（`workflow/artifact_store.py`）暂存中间产物

### AUD-S1-F05 制作联合模板
② `web_app.py::create_template_merge` @1702 → `service.merge_template_files` @407-438 ③ 基准模板+所选模板 ④ native=`native/template_merge.py::run_template_merge_native`（@425-427）或 COM `template_merge` ⑤ 复制 sheet 值/公式/样式/批注/合并/尺寸、名称增删改（`copy_sheet_names`/`localize_base_names`）、外链重写（`rebase_external_references`）、删除原表（`_drop_original`） ⑥ **是（全对象）** ⑦ 是 ⑧ 同 ④ ⑨ 是 ⑩ 同 ⑪ 联合模板可打开、名称引用正确 ⑬ COM 可选 ⑭ 源只读

### AUD-S2-F01 执行报表规则
② `web_app.py::start_period_compare_v2` @2741 → `PeriodAuditV2Service` @2861（进度回调 @2799-2823） ③ 本期/上期报表目录 + S2 配置簿 + 可选外部核对簿 ④ **python-calamine 0.3.1 快速读取**（`systems/s2_report_collection/importer.py:25-30`，`_calamine_load` @110-113），openpyxl `load_workbook(read_only=True, data_only=True)` 回退 @151/476，.xls 用 xlrd @146；外部簿 calamine @482-504 ⑤ A-C 列指标行提取、机构/指标补充、金额单位归一（`unit_conversion.py` Decimal）、去重、期间比较（`period_engine.py`）、规则引擎（`rule_engine.py` 累计不降+表达式）、外部核对（`external_engine.py`）、校验（`validation.py`）、摘要（`exporter.build_summary`） ⑥ 否（普通存储值） ⑦ 否 ⑧ — ⑨ 是 ⑩ **Direct OOXML**：`exporter.py::_write_direct_ooxml` @348（`zipfile.ZipFile(path,"w",ZIP_DEFLATED,compresslevel=6)` @409），经临时路径再落位 @441 ⑪ 三表结果（数值/文字/公式格），行数与配置一致 ⑫ Python 业务 + Python 直写 OOXML ⑬ 无（纯 Python 进程） ⑭ 配置检查阻断（`config_check.py`）时不建输出；导出走临时文件

### AUD-S3-F01 两期数据到比较文件
② `web_app.py::start_central_comparison` @2367 → `s3_central_statistics/service.py::run_comparison` @2389（office_eval @2397） ③ 两期 CSV 导数 + 比较配置（COMMON/COMPARISON） ④ CSV：`csv_importer.py`（python-calamine/编码处理）；比较阶段 openpyxl `load_workbook(BytesIO(raw), read_only=True, data_only=True)`（`service.py:447,475`） ⑤ 规则动作引擎（`rule_action_engine.py`）、适用性（`rule_applicability.py`）、指标规则（`indicator_rule_engine.py`）、表达式（`expression_*`）、五段式（SCHEMA_FIVE_SEGMENT_V1） ⑥ 部分（渲染器阶段） ⑦ 否 ⑨ 是 ⑩ **Rust writer 子进程** `rust/s3_comparison_writer`（`native/s3_comparison_writer.py:44-49` 定位、@76-146 执行+临时文件 @115+错误回退 @140-146；调用点 `comparison_exporter.py:14,130`） ⑪ 比较文件（值+说明+计算过程） ⑫ Python 业务 + Rust 写出 + Office/LO 表达式求值（office_eval，可选） ⑬ Rust 子进程 + COM/LO 可选 ⑭ 临时文件+失败报错保留原状

### AUD-S3-F02 比较文件到说明文件
② `start_central_explanation_export` @2523 → `explanation_exporter.export_selected_explanations` @2546 ④ openpyxl 读比较文件 ⑨ 是（说明文件导出） ⑩ `explanation_exporter.py`（**待验证**：写出引擎细节本轮未逐行核对） ⑪ 说明文件契约

### AUD-S3-F03 机构说明到说明文件
② `import_central_explanation_feedback` @2565 → @2592 ③ 多机构反馈文件 ⑤ 身份键匹配、冲突报告 ⑩ 同上模块（**待验证**） ⑪ 说明回写+冲突清单

### AUD-S3-F04 本期数值核对
② `start_central_cross_period` @2433 → `run_cross_period_check` @2458 ④ openpyxl（`service.py` 读路径同 S3-F01） ⑤ 跨期规则（`cross_period_engine.py`、CROSS_SHEET 配置、`check_cross_formulas` @2198） ⑩ 引擎（**待验证**） ⑪ 核对结果

### AUD-S3-F05 比较文件到金融表单
② `start_central_forms` @2478 → `render_financial_forms` @2501 ④ openpyxl + `ooxml_structure.py` 结构读取 ⑤ `fast_ooxml.py::TemplateCompiler._compile`/`FastOoxmlRenderer.render_org`（共享字符串/定义名称/关系直操作）、`form_template.py`、`renderers.py` ⑥ **是（表单结构+格式）** ⑨ 是 ⑩ **Direct OOXML**（`fast_ooxml.py` + `_validate_output`） ⑪ 金融表单 OOXML 结构保真 + `fast_delete_defined_names` ⑬ 无 ⑭ `_validate_output` 校验

### AUD-S5-F01 逐笔统计（配置版）
② `web_api.py::start_s5("S5-F01")` @684-685 ③ 显式模板 + 只读输入 + V_PRE 配对 ④ `s1audit_configured/sources.py`（`_Workbook`/SourceBinding，含布局探测 @861-863）、`source_reader.py`（OOXML/样式/工作表元数据） ⑤ `s1audit_configured/engine.py`（dsl.py 规则、conditional_audit、excel_precision、numeric_scope）、`recognition.py` ⑥ **是（样式/布局/精度）** ⑦ 否（源只读） ⑨ 是 ⑩ `s1audit_configured/outputs.py`：env `AUDIT3_WRITER_ENGINE`——"direct"=fast_ooxml `write_tables`（@17 import；`supports_tables` 判断 @12-13）或回退 openpyxl `_write_result_workbook_openpyxl` @309（save @405） ⑪ 审核结果工作簿+源单元格链接（`_source_cell_link` @301） ⑫ Python ⑬ 无 ⑭ 只读输入保证；结果含完整性刷新（`_refresh_integrity` @102）

### AUD-S5-F02/F03/F04 隐藏维护工具
② `web_api.py::update_s5_maintenance` @527、`start_s5_maintenance` @577、`_s5_maintenance_worker` @639（五连击解锁显示，再次五击隐藏——`TEST_GROUPS.md` S5 节） ③-⑪ **待验证**（维护工具具体业务行为需逐个源码确认，workspace 清单亦标待确认）

### AUD-C01 配置指引/初始化/重置
② `web_app.py:859-1087`、S3 检查重置 @2020-2363 ④ openpyxl（`config_guide.py`、`period_compare.py::ensure_default_config`、`name_config.py`、s2/s3 config 默认写入函数） ⑦/⑨ 是 ⑪ 配置簿模板结构 ⑭ 重置前确认（前端）

### AUD-C02 历史记录
② `get_history_page` @1265、`open_history_explanation` @1248 ④ openpyxl（`history.py`、`native/history_io.py::write_history_xlsx` @67-90：`load_workbook(path)` 原位 @70 → append → `book.save(path)` @90） ⑦ 是 ⑪ 历史簿追加行、HISTORY_HEADERS 列契约

### AUD-C03 LibreOffice 安装/维护
② `web_app.py:1358-1438` ④ `native/libreoffice_install.py`（安装包发现/安装/卸载；麒麟维护动作 @1416） ⑬ 子进程（dpkg/安装器） ⑪ 引擎可用性 ⑭ 安装失败不影响其他功能

### AUD-C04 诊断工作簿导出
② `export_diagnostic_workbook` @664（clean() @680，probe 线程 @827,849） ⑨ 是（openpyxl 诊断簿） ⑪ 环境探测结果

### AUD-C05 DAG 工作流编辑器
② `web_app.py:1118-1236` ④ `node_flow_config.py`（DAG/工作表重置、替换、保存；openpyxl 配置簿） ⑦ 是（配置簿） ⑪ 固定 DAG 不可被旧配置改写（测试合同）

### AUD-C06 外部参考配置
② `web_app.py:962-977`、`choose_history_config` @1682 ④ `external.py`/`external_sheet_writer.py`（Direct OOXML：`prepare`/`write`/`_parse_data_sheet`/`_validate_output`——外部 sheet 组装，关系/Content_Types/sharedStrings 直操作） ⑨ 是 ⑩ Direct OOXML ⑪ 关系完整性+`_validate_output`

### AUD-C07 S5 表达式参考导出
② `expression_export.py:355-370` argparse CLI（V_CUR/V_EXT 副本批量生成） ⑩ 独立 CLI 输出（**待验证**：写出引擎本轮未逐行核对）

## 4. 汇总矩阵（功能组 × 引擎）

| 功能组（编号） | 读取引擎（源） | 写出引擎（源） | 修改既有簿 | 样式/元数据依赖 | Rust 现状引擎 |
|---|---|---|---|---|---|
| PT-1-x 时序（1-1..1-4） | NoCom pandas+calamine；PyCode pandas+openpyxl/xlrd；COM 版 Value2 | pandas ExcelWriter(openpyxl)；追加=重建替换 | 否（追加目标簿除外） | 合并区坐标 | calamine 读 + rust_xlsxwriter 新建（CLI 已接） |
| PT-2-x 去重 | openpyxl（含批注） | openpyxl 原位（标色/删行/追加） | **是** | 批注核心 | 计划构造（SetCellFillColor/DeleteRows/ImportRange），执行被能力门控 |
| PT-3-x 打印（3-1..3-4） | openpyxl 双开（3-2 值=calamine） | openpyxl 新建/改目标簿 | 3-3 是（清空目标 sheet 重写） | 样式+批注+页面设置 | 计划构造 + umya 候选 ImportRange，执行被能力门控 |
| PT-4-1/4-5..4-10 文件操作 | 配置簿 pandas/calamine（4-6/4-7 内容格 calamine/openpyxl） | 报告 pandas（4-9/4-10） | 否 | 值 | calamine 读 + rust_xlsxwriter 报告 + 文件系统，多数有合成对拍 |
| PT-4-2/4-3 转换 | Office/LO | Office SaveAs / LO | 否 | 引擎保真 | sidecar（LO/COM）+ 校验，或 native 按值转换 |
| PT-4-4 Sheet 改名 | openpyxl | openpyxl 原位（PyCode .xls=COM） | **是** | keep_vba | 计划构造（RenameSheet） |
| PT-5-x 配置 | pandas 查看 / openpyxl 修复 | openpyxl 原位+批注 | **是** | 批注+表头 | writer 新建 + 修复计划 |
| PT-6-x 汇总 | NoCom CalamineWorkbook 直连（**双开**）；PyCode openpyxl；模板批注=openpyxl | openpyxl 新建 | 否 | 批注指令+合并 | calamine 读 + rust_xlsxwriter 新建（CLI 已接，无对拍） |
| PT-7-1/7-3 批处理结构操作 | openpyxl data_only | openpyxl 原位删插行 | **是** | 合并区（现状不修正） | 计划构造（DeleteRows 等） |
| PT-7-2 机构标准化 | openpyxl/共享 reader | openpyxl 原位（源）；Rust 计划+umya 候选（测试内） | 是 | 值 | 计划级对拍；执行未路由 |
| PT-7-4 页眉 | openpyxl 页眉 | openpyxl 原位 | 是 | 页眉文本 | ExcelEditPlan→OOXML 页眉补丁（合成对拍通过） |
| PT-7-5 地区校验 | openpyxl 只读 | openpyxl 新建报告 | 否 | 报告样式 | calamine 读 + rust_xlsxwriter + `clear_empty_formula_cache` 定点补丁（CLI+合成对拍） |
| PT-7-6 提取 | calamine（NoCom 纯值）/openpyxl 样式版/COM 批注版 | openpyxl 新建/追加 | 否（追加输出簿除外） | 档位不同 | area 值提取 CLI（合成对拍）；整表 ImportRange 候选 |
| PT-7-7/7-8 贷款修正 | openpyxl（本期可写/上期只读） | openpyxl 原位 | 是 | 值+公式覆盖语义 | 计划构造 |
| PT-8-x 归档 | 无 | pandas 报告 | 否 | 否 | 文件系统 + rust_xlsxwriter 报告（合成对拍） |
| AUD-S1-F01/F02 | ExcelSession COM / openpyxl native / python-calamine 汇总读取 | **Rust 子进程 writer**（rust_xlsxwriter）+ openpyxl 回退 + history openpyxl 原位 | 导航/历史写入 | 条件格式/批注/名称/公式 | template-suggest 等切片已迁；主链未迁 |
| AUD-S2-F01 | python-calamine 快速 + openpyxl 回退 + xlrd | **Direct OOXML**（exporter） | 否 | 否 | Calamine 全链 + rust_xlsxwriter 三表（CLI+合成对拍 0 差异） |
| AUD-S3-F01 | csv_importer + openpyxl | **Rust 子进程 writer** | 否 | 部分 | central_csv 切片已迁 |
| AUD-S3-F05 | openpyxl + ooxml_structure | **Direct OOXML**（fast_ooxml） | 否 | 表单结构 | 未迁 |
| AUD-S5-F01 | s1audit_configured source_reader（OOXML 元数据） | fast_ooxml direct / openpyxl 回退 | 否 | 样式/布局/精度 | 未迁 |
