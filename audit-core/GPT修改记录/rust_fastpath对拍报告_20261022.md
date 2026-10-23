# Calamine 读取与 rust_xlsxwriter 新建写入对拍报告

日期：2026-10-22  
项目分支：`rust-migration`  
依赖：[Calamine 0.37.0 Reader API](https://docs.rs/calamine/0.37.0/calamine/trait.Reader.html)、[rust_xlsxwriter 0.101.0 Workbook API](https://docs.rs/rust_xlsxwriter/0.101.0/rust_xlsxwriter/workbook/struct.Workbook.html)

## 范围与环境

未找到可用于本项目的脱敏真实工作簿，因此仅以 `tests/fixtures/synthetic_probe.xlsx` 作为合成读取样本，并另行生成一个 rust_xlsxwriter smoke workbook。两者均不包含真实业务数据。本报告不是生产模板验收。OpenPyXL 3.1.5 仅作为独立读取基线；没有引入 Python 业务运行时。

- Rust/Cargo 1.99.1，WSL Ubuntu 24.04 x86_64，glibc 2.39。
- 对拍基线：Windows 11 x86_64、Python 3.12.10、OpenPyXL 3.1.5。
- 读取预热 1 次、计时 5 次；写入预热 1 次、计时 5 次。Rust 在 WSL 访问 Windows 工作区，基线在 Windows 运行，计时环境不对称。
- 未使用 Office/WPS/LibreOffice；没有公式重算。

## Calamine 读取对拍

读取 JSON 导出值、sheet 名、合并区域、公式文本、批注文本，并与 OpenPyXL 结果比较：

| 检查项 | 匹配 | 通过率 |
|---|---:|---:|
| Sheet 名称及顺序 | 1/1 | 100% |
| 单元格值 | 65/65 | 100% |
| 公式文本 | 1/1 | 100% |
| 合并区域 | 1/1 | 100% |
| 批注文本 | 1/1 | 100% |

批注通过 namespace-aware OOXML 补充读取，因为 Calamine 没有批注读取 API；因此不能把批注结果归因于 Calamine。当前 Calamine 0.37.0 探针明确报告 xlsb 合并区域 API 不可用；该字段对 xlsb 为“不支持/未验证”，不会伪造为空合并集。样本仅为一个 xlsx，未测 xls、xlsm、xlsb、错误单元格或大文件。

## rust_xlsxwriter 新建写入对拍

探针从零生成 xlsx，覆盖 sheet 名及顺序、值、公式文本、合并、字体/填充/数字格式、defined name、行高列宽、打印区域；OpenPyXL 独立重开检查文件和结构。结果 **9/9，100%**。列宽因 Excel 宽度编码与读取换算有小数差异，使用小于 1 个宽度单位的容差；实际读取值 25.7109375，对应请求值 25。

该结果只证明所列新建对象在合成工作簿上的读回结果，不证明编辑既有模板、批注/条件格式保留、打印渲染或目标办公套件兼容。

## 性能记录

| 阶段 | 中位数 | 范围 |
|---|---:|---:|
| Calamine 值/公式/合并/批注导出（含 OOXML 批注补充） | 149 ms | 97–170 ms |
| OpenPyXL 读取基线 | 31.68 ms | 21.14–41.93 ms |
| rust_xlsxwriter 新建工作簿 | 223 ms | 138–251 ms |

Rust 与 OpenPyXL 的运行环境、文件路径和实现阶段并不对称；该小样本不能用于声称 Rust 更快或更慢。写入任务只测了 Rust 探针内部时间，没有做 OpenPyXL 写入性能对比。

## 结论与决策

- **读取**：本合成 xlsx 上，Calamine 主数据能力达到 100%；批注由 OOXML 补读。按规则可以将 Calamine 作为待真实模板验收的 reader 候选，不能据此宣称 xls/xlsm/xlsb 已通过。
- **新建写入**：rust_xlsxwriter 在本样本达到 100%，可作为新建 xlsx writer 候选；范围不包括编辑已有工作簿。
- **Go**：读取和新建写入路径不需要 Go。结合 umya 报告中的单项列定义补丁结果，目前不引入 Go excelize。
- **业务迁移**：探针阶段不迁移 Python 业务代码。需先用实际脱敏模板补齐格式与对象覆盖，再开始业务迁移。

## 共享 workspace 接入复核

探针移入 `audit-core` 并改用 `excel-reader` 与 `excel-writer` 后，同一 fixture 的 sheet 名/顺序 **1/1**、值 **65/65**、公式 **1/1**、合并 **1/1**、批注 **1/1** 均匹配；新建工作簿仍为 **9/9**。本节复核不新增性能测量。

## 未完成的验收

- 真实模板：0 个可用，未测。
- xls/xlsm/xlsb 多格式真实样本：未测；xlsb 合并区域不支持/未验证。
- 麒麟 V10 glibc 2.31、统信 V20 glibc 2.28 实机运行与 deb 安装：未测。
- 公式计算结果、原生打印与渲染：不在本轮支持范围。

## Linux ABI 构建记录

`packaging/build.sh` 使用 `cargo-zigbuild 0.24.0` 以 `x86_64-unknown-linux-gnu.2.28` 构建 release workspace，构建成功。`objdump -T` 检查两个探针二进制的最高 GLIBC 符号版本均为 **2.28**；动态依赖为 `libm`、`libc`、`ld-linux-x86-64`、`libpthread`、`libdl`。两个二进制在 Ubuntu 24.04 构建主机通过 `--help` 启动检查。该项不替代麒麟/统信实机及 deb 安装验收；当前没有生产 CLI，故未生成 deb。

