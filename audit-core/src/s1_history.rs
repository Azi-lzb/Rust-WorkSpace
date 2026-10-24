//! S1 issue history classification and manual feedback carry-over.

use crate::models::Issue;
use anyhow::{Context, Result, ensure};
use excel_editor::{
    CellValue as EditCellValue, EditOperation, ExcelEditPlan, PreservationRequirement,
    WorkbookFormat, inspect_xlsx_profile,
};
use excel_reader::{CellValue, SheetValues, read_value_columns_with_iso_dates};
use excel_writer::Workbook as XlsxWorkbook;
use serde::{Deserialize, Serialize};
use std::{
    collections::{BTreeMap, HashMap, HashSet},
    fs,
    path::Path,
    time::{Instant, SystemTime, UNIX_EPOCH},
};

const HISTORY_AUDIT_SHEET: &str = "核查表校验结果";
const LEGACY_HISTORY_SHEET: &str = "历史审核结果";
const INITIAL_HISTORY_READ_COLUMNS: usize = 14;
pub const S1_HISTORY_HEADERS: [&str; 13] = [
    "工作簿名",
    "工作表名",
    "定位单元格",
    "错误类型",
    "校验类型",
    "校验指标",
    "描述",
    "当前值",
    "对比值",
    "差值",
    "规则编号",
    "历史校验说明",
    "审核意见",
];

pub type S1HistoryWorkbookRow = [EditCellValue; 13];

/// Convert Rust issues into the ordered cells used by Python's history writer.
///
/// This keeps numeric, boolean, and blank issue values typed for ExcelEditPlan;
/// it does not open or modify a workbook.
pub fn s1_history_workbook_row(issue: &Issue) -> Result<S1HistoryWorkbookRow> {
    let source_name = Path::new(&issue.source_file)
        .file_name()
        .map(|name| name.to_string_lossy().into_owned())
        .unwrap_or_default();
    let description = if issue.detail.is_empty() {
        issue.message.clone()
    } else {
        issue.detail.clone()
    };
    Ok([
        EditCellValue::Text(source_name),
        EditCellValue::Text(issue.sheet_name.clone()),
        EditCellValue::Text(issue.target_cell.clone()),
        EditCellValue::Text(issue.severity.clone()),
        EditCellValue::Text(issue.check_type.clone()),
        EditCellValue::Text(issue.check_field.clone()),
        EditCellValue::Text(description),
        history_json_cell_value(&issue.target_value, "当前值")?,
        history_json_cell_value(&issue.comparison_value, "对比值")?,
        history_json_cell_value(&issue.difference_value, "差值")?,
        EditCellValue::Text(issue.issue_id.clone()),
        EditCellValue::Text(issue.institution_feedback.clone()),
        EditCellValue::Text(issue.auditor_opinion.clone()),
    ])
}

pub fn s1_history_header_row() -> S1HistoryWorkbookRow {
    S1_HISTORY_HEADERS.map(|header| EditCellValue::Text(header.to_owned()))
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct S1HistoryWorkbookWriteSummary {
    pub path: std::path::PathBuf,
    pub sheet_name: &'static str,
    pub issue_count: usize,
    pub engine: &'static str,
    pub elapsed_ms: u128,
}

#[derive(Clone, Debug)]
pub struct S1HistoryWorkbookEditPlan {
    pub edit_plan: ExcelEditPlan,
    pub issue_count: usize,
    pub replaces_existing_history_sheet: bool,
}

/// Build the ExcelModifier batch corresponding to Python's
/// `write_history_xlsx`: replace only the managed history worksheet, keep other
/// workbook sheets under the preservation contract, and place the new history
/// sheet first. This prepares Audit business data and does not execute an
/// editor or qualify a production capability.
pub fn build_s1_history_workbook_edit_plan(
    input_path: impl AsRef<Path>,
    output_path: impl AsRef<Path>,
    issues: &[Issue],
) -> Result<S1HistoryWorkbookEditPlan> {
    let input_path = input_path.as_ref();
    let output_path = output_path.as_ref();
    ensure!(input_path.is_file(), "S1 历史工作簿不存在");
    ensure!(
        xlsx_extension(input_path),
        "S1 历史工作簿更新目前仅支持 .xlsx"
    );
    ensure!(
        xlsx_extension(output_path),
        "S1 历史工作簿输出目前仅支持 .xlsx"
    );
    ensure!(
        !output_path.exists(),
        "拒绝覆盖已有 S1 历史工作簿输出：{}",
        output_path.display()
    );
    let output_parent = output_path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    ensure!(output_parent.is_dir(), "S1 历史工作簿输出目录不存在");
    ensure!(
        issues.len() <= 1_048_575,
        "S1 历史工作簿超过 Excel 的最大数据行数"
    );

    let input_path = fs::canonicalize(input_path).context("规范化 S1 历史工作簿路径")?;
    let output_parent = fs::canonicalize(output_parent).context("规范化 S1 历史工作簿输出目录")?;
    let output_name = output_path
        .file_name()
        .context("S1 历史工作簿输出路径缺少文件名")?;
    let output_path = output_parent.join(output_name);
    ensure!(
        !output_path.exists(),
        "拒绝覆盖已有 S1 历史工作簿输出：{}",
        output_path.display()
    );

    let workbook_profile = inspect_xlsx_profile(&input_path).context("检查 S1 历史工作簿结构")?;
    ensure!(
        workbook_profile.format == WorkbookFormat::Xlsx,
        "S1 历史工作簿更新目前仅支持 XLSX"
    );
    ensure!(
        workbook_profile.inventory_complete,
        "S1 历史工作簿 OOXML 对象清单不完整，拒绝准备更新计划"
    );
    let sheet_names =
        excel_reader::read_sheet_names(&input_path).context("读取 S1 历史工作簿工作表名称")?;
    let replaces_existing_history_sheet =
        sheet_names.iter().any(|name| name == HISTORY_AUDIT_SHEET);

    let mut operations = Vec::new();
    if replaces_existing_history_sheet {
        operations.push(EditOperation::DeleteSheet {
            name: HISTORY_AUDIT_SHEET.to_owned(),
        });
    }
    operations.push(EditOperation::CreateSheetAt {
        name: HISTORY_AUDIT_SHEET.to_owned(),
        index: 0,
    });
    for (column, value) in s1_history_header_row().into_iter().enumerate() {
        operations.push(EditOperation::SetCellValue {
            sheet: HISTORY_AUDIT_SHEET.to_owned(),
            cell: s1_history_cell_address(0, column),
            value,
        });
    }
    for (row_index, issue) in issues.iter().enumerate() {
        for (column, value) in s1_history_workbook_row(issue)?.into_iter().enumerate() {
            if matches!(value, EditCellValue::Blank)
                || matches!(&value, EditCellValue::Text(text) if text.is_empty())
            {
                continue;
            }
            operations.push(EditOperation::SetCellValue {
                sheet: HISTORY_AUDIT_SHEET.to_owned(),
                cell: s1_history_cell_address(row_index + 1, column),
                value,
            });
        }
    }
    let last_row = issues.len() + 1;
    operations.push(EditOperation::SetFreezePanes {
        sheet: HISTORY_AUDIT_SHEET.to_owned(),
        cell: Some("A2".to_owned()),
    });
    operations.push(EditOperation::SetAutoFilter {
        sheet: HISTORY_AUDIT_SHEET.to_owned(),
        range: Some(format!("A1:M{last_row}")),
    });
    for column_index in 1..=S1_HISTORY_HEADERS.len() {
        operations.push(EditOperation::SetColumnWidth {
            sheet: HISTORY_AUDIT_SHEET.to_owned(),
            column_index: column_index as u32,
            width: 22.0,
        });
    }

    let preserve: std::collections::BTreeSet<_> = [
        PreservationRequirement::CellValuesAndTypes,
        PreservationRequirement::FormulaText,
        PreservationRequirement::CommentsAndPositions,
        PreservationRequirement::Styles,
        PreservationRequirement::MergedRanges,
        PreservationRequirement::ConditionalFormatting,
        PreservationRequirement::DataValidations,
        PreservationRequirement::DefinedNames,
        PreservationRequirement::RowAndColumnDimensions,
        PreservationRequirement::PageSettings,
        PreservationRequirement::HeaderFooterText,
        PreservationRequirement::Relationships,
        PreservationRequirement::SheetIdentityAndOrder,
        PreservationRequirement::FreezePanes,
        PreservationRequirement::AutoFilter,
    ]
    .into_iter()
    .collect();
    let mut drop_allowed = BTreeMap::new();
    if replaces_existing_history_sheet {
        drop_allowed.insert(HISTORY_AUDIT_SHEET.to_owned(), preserve.clone());
    }
    let edit_plan = ExcelEditPlan {
        request_id: "audit-s1-history-workbook-update".to_owned(),
        input: input_path,
        output: output_path,
        format: WorkbookFormat::Xlsx,
        workbook: workbook_profile,
        operations,
        preserve,
        drop_allowed,
    };
    edit_plan
        .validate()
        .context("校验 S1 历史工作簿 ExcelEditPlan")?;

    Ok(S1HistoryWorkbookEditPlan {
        edit_plan,
        issue_count: issues.len(),
        replaces_existing_history_sheet,
    })
}

fn xlsx_extension(path: &Path) -> bool {
    path.extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
}

/// Create a standalone XLSX containing the merged S1 history rows.
///
/// This writer deliberately creates a new file and never edits an existing
/// workbook. Preservation-style updates to existing history workbooks must use
/// the ExcelEditPlan/ExcelModifier route after that operation capability is
/// validated for its real template and platform.
pub fn write_s1_history_workbook(
    output_path: impl AsRef<Path>,
    issues: &[Issue],
) -> Result<S1HistoryWorkbookWriteSummary> {
    const ENGINE: &str = "rust_xlsxwriter 0.101.0";

    let started = Instant::now();
    let output_path = output_path.as_ref();
    ensure!(
        output_path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx")),
        "S1 历史结果工作簿输出路径必须使用 .xlsx"
    );
    ensure!(
        !output_path.exists(),
        "拒绝覆盖已有 S1 历史结果工作簿：{}",
        output_path.display()
    );
    ensure!(
        issues.len() <= 1_048_575,
        "S1 历史结果工作簿超过 Excel 的最大数据行数"
    );

    let parent = output_path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or(Path::new("."));
    fs::create_dir_all(parent)
        .with_context(|| format!("创建 S1 历史结果工作簿目录 {}", parent.display()))?;
    let temp_path = s1_history_temporary_path(output_path)?;

    let write_result = (|| -> Result<()> {
        {
            let mut workbook = XlsxWorkbook::new();
            let worksheet = workbook.add_worksheet();
            worksheet
                .set_name(HISTORY_AUDIT_SHEET)
                .context("设置 S1 历史结果工作表名称")?;
            for (column, header) in S1_HISTORY_HEADERS.iter().enumerate() {
                worksheet
                    .write_string(0, column as u16, *header)
                    .with_context(|| format!("写入 S1 历史表头第 {} 列", column + 1))?;
            }

            for (row_index, issue) in issues.iter().enumerate() {
                let row = s1_history_workbook_row(issue)?;
                for (column, value) in row.iter().enumerate() {
                    let row_number = row_index as u32 + 1;
                    let column_number = column as u16;
                    match value {
                        EditCellValue::Text(value) if !value.is_empty() => {
                            worksheet.write_string(row_number, column_number, value)?;
                        }
                        EditCellValue::Text(_) | EditCellValue::Blank => {}
                        EditCellValue::Decimal(value) => {
                            let number = value.parse::<f64>().with_context(|| {
                                format!("解析 S1 历史数值第 {} 行", row_index + 2)
                            })?;
                            ensure!(number.is_finite(), "S1 历史数值必须是有限数值");
                            worksheet.write_number(row_number, column_number, number)?;
                        }
                        EditCellValue::Boolean(value) => {
                            worksheet.write_boolean(row_number, column_number, *value)?;
                        }
                        EditCellValue::Error(value) => {
                            worksheet.write_string(row_number, column_number, value)?;
                        }
                        EditCellValue::DateTime { .. } => {
                            anyhow::bail!("S1 历史结果暂不接受未经日期契约映射的日期单元格")
                        }
                    }
                }
            }

            worksheet
                .set_freeze_panes(1, 0)
                .context("冻结 S1 历史表头")?;
            worksheet
                .autofilter(
                    0,
                    0,
                    issues.len() as u32,
                    S1_HISTORY_HEADERS.len() as u16 - 1,
                )
                .context("设置 S1 历史结果表筛选")?;
            for column in 0..S1_HISTORY_HEADERS.len() {
                // Keep Python's stored character width (22) exact in OOXML.
                worksheet
                    .set_column_width_pixels(column as u16, 22 * 7)
                    .with_context(|| format!("设置 S1 历史表第 {} 列宽度", column + 1))?;
            }
            workbook
                .save(&temp_path)
                .context("保存 S1 历史结果临时工作簿")?;
        }

        validate_s1_history_workbook(&temp_path, issues)?;
        // A same-directory hard link commits without replacing a file that may
        // have appeared after the initial existence check.
        fs::hard_link(&temp_path, output_path)
            .with_context(|| format!("提交 S1 历史结果工作簿 {}", output_path.display()))?;
        let _ = fs::remove_file(&temp_path);
        Ok(())
    })();
    if write_result.is_err() {
        let _ = fs::remove_file(&temp_path);
    }
    write_result?;

    Ok(S1HistoryWorkbookWriteSummary {
        path: fs::canonicalize(output_path)
            .with_context(|| format!("规范化 S1 历史结果路径 {}", output_path.display()))?,
        sheet_name: HISTORY_AUDIT_SHEET,
        issue_count: issues.len(),
        engine: ENGINE,
        elapsed_ms: started.elapsed().as_millis(),
    })
}

fn s1_history_temporary_path(output_path: &Path) -> Result<std::path::PathBuf> {
    let parent = output_path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or(Path::new("."));
    let stem = output_path
        .file_stem()
        .context("S1 历史结果工作簿输出路径缺少文件名")?
        .to_string_lossy();
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .context("系统时间早于 Unix epoch")?
        .as_nanos();
    Ok(parent.join(format!(
        ".{stem}.rust-audit-{}-{nonce}.tmp.xlsx",
        std::process::id()
    )))
}

fn validate_s1_history_workbook(path: &Path, issues: &[Issue]) -> Result<()> {
    let workbook = excel_reader::read_workbook(path).context("重新打开 S1 历史结果工作簿")?;
    ensure!(
        workbook.sheets.len() == 1 && workbook.sheets[0].name == HISTORY_AUDIT_SHEET,
        "S1 历史结果工作簿的工作表名称或数量不匹配"
    );
    let sheet = &workbook.sheets[0];
    ensure!(sheet.formulas.is_empty(), "S1 历史结果工作簿不应包含公式");
    ensure!(
        sheet.merges.is_empty(),
        "S1 历史结果工作簿不应包含合并单元格"
    );

    let mut expected = Vec::with_capacity((issues.len() + 1) * S1_HISTORY_HEADERS.len());
    expected.extend(
        s1_history_header_row()
            .into_iter()
            .enumerate()
            .map(|(column, value)| (s1_history_cell_address(0, column), value)),
    );
    for (row_index, issue) in issues.iter().enumerate() {
        expected.extend(
            s1_history_workbook_row(issue)?
                .into_iter()
                .enumerate()
                .map(|(column, value)| (s1_history_cell_address(row_index + 1, column), value)),
        );
    }
    let expected_nonblank = expected
        .iter()
        .filter(|(_, value)| match value {
            EditCellValue::Blank => false,
            EditCellValue::Text(value) => !value.is_empty(),
            _ => true,
        })
        .count();
    ensure!(
        sheet.values.len() == expected_nonblank,
        "S1 历史结果工作簿单元格数量不匹配：预期 {}，实际 {}",
        expected_nonblank,
        sheet.values.len()
    );
    for (address, value) in expected {
        let actual = sheet.values.get(&address);
        match value {
            EditCellValue::Blank => {
                ensure!(
                    actual.is_none_or(|actual| actual
                        .get("type")
                        .and_then(serde_json::Value::as_str)
                        == Some("empty")),
                    "S1 历史结果工作簿空白单元格 {address} 意外有值"
                );
            }
            EditCellValue::Text(value) if value.is_empty() => {
                ensure!(
                    actual.is_none_or(|actual| actual
                        .get("type")
                        .and_then(serde_json::Value::as_str)
                        == Some("empty")),
                    "S1 历史结果工作簿空白单元格 {address} 意外有值"
                );
            }
            EditCellValue::Text(value) | EditCellValue::Error(value) => {
                let actual =
                    actual.with_context(|| format!("S1 历史结果工作簿缺少单元格 {address}"))?;
                ensure!(
                    actual.get("type").and_then(serde_json::Value::as_str) == Some("string")
                        && actual.get("value").and_then(serde_json::Value::as_str)
                            == Some(value.as_str()),
                    "S1 历史结果工作簿单元格 {address} 内容或类型不匹配"
                );
            }
            EditCellValue::Decimal(value) => {
                let expected = value.parse::<f64>().context("解析 S1 历史预期数值")?;
                let actual =
                    actual.with_context(|| format!("S1 历史结果工作簿缺少单元格 {address}"))?;
                let actual_type = actual.get("type").and_then(serde_json::Value::as_str);
                let actual_value = actual.get("value").and_then(serde_json::Value::as_f64);
                ensure!(
                    matches!(actual_type, Some("integer" | "number"))
                        && actual_value.is_some_and(|value| (value - expected).abs() <= 1e-12),
                    "S1 历史结果工作簿单元格 {address} 数值或类型不匹配"
                );
            }
            EditCellValue::Boolean(expected) => {
                let actual =
                    actual.with_context(|| format!("S1 历史结果工作簿缺少单元格 {address}"))?;
                ensure!(
                    actual.get("type").and_then(serde_json::Value::as_str) == Some("boolean")
                        && actual.get("value").and_then(serde_json::Value::as_bool)
                            == Some(expected),
                    "S1 历史结果工作簿单元格 {address} 布尔值或类型不匹配"
                );
            }
            EditCellValue::DateTime { .. } => {
                anyhow::bail!("S1 历史结果工作簿暂不支持日期值读回校验")
            }
        }
    }
    Ok(())
}

fn s1_history_cell_address(row: usize, column: usize) -> String {
    let mut number = column + 1;
    let mut letters = Vec::new();
    while number > 0 {
        let remainder = (number - 1) % 26;
        letters.push((b'A' + remainder as u8) as char);
        number = (number - 1) / 26;
    }
    letters.reverse();
    format!("{}{}", letters.into_iter().collect::<String>(), row + 1)
}

fn history_json_cell_value(value: &serde_json::Value, header: &str) -> Result<EditCellValue> {
    match value {
        serde_json::Value::Null => Ok(EditCellValue::Blank),
        serde_json::Value::Bool(value) => Ok(EditCellValue::Boolean(*value)),
        serde_json::Value::Number(value) => Ok(EditCellValue::Decimal(value.to_string())),
        serde_json::Value::String(value) => Ok(EditCellValue::Text(value.clone())),
        serde_json::Value::Array(_) | serde_json::Value::Object(_) => {
            anyhow::bail!("S1 历史工作簿字段 {header} 不是 Excel 标量值")
        }
    }
}

#[derive(Clone, Debug, Deserialize)]
pub struct S1HistoryRequest {
    pub period: String,
    #[serde(default)]
    pub batch_id: String,
    #[serde(default)]
    pub audit_time: String,
    #[serde(default)]
    pub audited_org_codes: Option<HashSet<String>>,
    #[serde(default)]
    pub current_issues: Vec<Issue>,
    #[serde(default)]
    pub history: Vec<Issue>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize)]
pub struct S1HistorySummary {
    pub history_rows: usize,
    pub merged_history_rows: usize,
    pub current_issues: usize,
    pub new_issues: usize,
    pub continuous_issues: usize,
    pub recurring_issues: usize,
    pub resolved_issues: usize,
    pub previous_period: String,
    pub elapsed_ms: u128,
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct S1HistoryResult {
    pub current_issues: Vec<Issue>,
    pub resolved_issues: Vec<Issue>,
    pub merged_history: Vec<Issue>,
    pub summary: S1HistorySummary,
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct S1HistoryWorkbookReadSummary {
    pub status: &'static str,
    pub reader: &'static str,
    pub sheet_name: Option<String>,
    pub rows_total: usize,
    pub rows_read: usize,
    pub rows_skipped: usize,
    pub elapsed_ms: u128,
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct S1HistoryWorkbookReadResult {
    pub issues: Vec<Issue>,
    pub summary: S1HistoryWorkbookReadSummary,
}

/// Read stored values from the Audit history sheet using the shared Calamine reader.
///
/// A missing file or supported history sheet matches the Python native adapter's
/// empty-history behavior. Formula values come from stored caches; this stage does
/// not recalculate or modify the workbook.
pub fn read_s1_history_workbook(
    path: impl AsRef<Path>,
    requested_sheet: Option<&str>,
) -> Result<S1HistoryWorkbookReadResult> {
    let started = Instant::now();
    let path = path.as_ref();
    if !path.is_file() {
        return Ok(S1HistoryWorkbookReadResult {
            issues: Vec::new(),
            summary: S1HistoryWorkbookReadSummary {
                status: "missing_file",
                reader: "none",
                sheet_name: None,
                rows_total: 0,
                rows_read: 0,
                rows_skipped: 0,
                elapsed_ms: started.elapsed().as_millis(),
            },
        });
    }
    ensure!(
        path.extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| {
                extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm")
            }),
        "S1 历史工作簿读取目前支持 .xlsx 和 .xlsm"
    );

    let mut sheets = read_value_columns_with_iso_dates(path, INITIAL_HISTORY_READ_COLUMNS)
        .with_context(|| format!("读取 S1 历史工作簿值：{}", path.display()))?;
    let selected_name = select_history_sheet_name(&sheets, requested_sheet);
    let Some(selected_name) = selected_name else {
        return Ok(S1HistoryWorkbookReadResult {
            issues: Vec::new(),
            summary: S1HistoryWorkbookReadSummary {
                status: "sheet_missing",
                reader: "calamine",
                sheet_name: None,
                rows_total: 0,
                rows_read: 0,
                rows_skipped: 0,
                elapsed_ms: started.elapsed().as_millis(),
            },
        });
    };
    let selected_extent = sheets
        .iter()
        .find(|sheet| sheet.name == selected_name)
        .map(|sheet| sheet.max_column)
        .unwrap_or(INITIAL_HISTORY_READ_COLUMNS);
    if selected_extent > INITIAL_HISTORY_READ_COLUMNS {
        sheets = read_value_columns_with_iso_dates(path, selected_extent)
            .with_context(|| format!("读取 S1 历史工作簿完整表头：{}", path.display()))?;
    }
    let sheet = sheets
        .iter()
        .find(|sheet| sheet.name == selected_name)
        .with_context(|| format!("Calamine 重读后缺少历史工作表：{selected_name}"))?;

    let issues = parse_history_sheet(sheet);
    let rows_total = sheet.max_row.saturating_sub(1);
    let rows_read = issues.len();
    let rows_skipped = rows_total.saturating_sub(rows_read);
    Ok(S1HistoryWorkbookReadResult {
        issues,
        summary: S1HistoryWorkbookReadSummary {
            status: if rows_total == 0 {
                "empty_sheet"
            } else {
                "loaded"
            },
            reader: "calamine",
            sheet_name: Some(selected_name),
            rows_total,
            rows_read,
            rows_skipped,
            elapsed_ms: started.elapsed().as_millis(),
        },
    })
}

fn select_history_sheet_name(sheets: &[SheetValues], requested: Option<&str>) -> Option<String> {
    let exists = |name: &str| sheets.iter().any(|sheet| sheet.name == name);
    requested
        .filter(|name| exists(name))
        .or_else(|| exists(HISTORY_AUDIT_SHEET).then_some(HISTORY_AUDIT_SHEET))
        .or_else(|| exists(LEGACY_HISTORY_SHEET).then_some(LEGACY_HISTORY_SHEET))
        .map(str::to_owned)
}

fn parse_history_sheet(sheet: &SheetValues) -> Vec<Issue> {
    let mut rows = BTreeMap::<usize, BTreeMap<usize, CellValue>>::new();
    for (address, value) in &sheet.values {
        if let Some((row, column)) = parse_cell_address(address) {
            rows.entry(row).or_default().insert(column, value.clone());
        }
    }
    let Some(header_row) = rows.get(&1) else {
        return Vec::new();
    };
    let mut headers = HashMap::<String, usize>::new();
    for (column, value) in header_row {
        let name = python_truthy_text(Some(value)).trim().to_owned();
        if !name.is_empty() {
            headers.insert(name, *column);
        }
    }

    let mut issues = Vec::new();
    for (_, row) in rows.range(2..) {
        let source_file = history_text(row, &headers, "工作簿名");
        let sheet_name = history_text(row, &headers, "工作表名");
        let target_cell = history_text(row, &headers, "定位单元格");
        let rule_id = history_text(row, &headers, "规则编号");
        if ![&source_file, &sheet_name, &target_cell, &rule_id]
            .iter()
            .any(|value| !value.is_empty())
        {
            continue;
        }
        let indicator = history_text(row, &headers, "校验指标");
        let identity = if rule_id.is_empty() {
            let indicator = if indicator.is_empty() {
                "校验指标未识别"
            } else {
                indicator.as_str()
            };
            format!("{source_file}｜{sheet_name}｜{target_cell}｜{indicator}")
        } else {
            rule_id
        };
        let description = history_text(row, &headers, "描述");
        let severity = history_text(row, &headers, "错误类型");
        issues.push(Issue {
            issue_id: identity.clone(),
            period: history_text(row, &headers, "数据期"),
            batch_id: String::new(),
            audit_time: String::new(),
            triggered: true,
            status: String::new(),
            first_seen_period: String::new(),
            previous_seen_period: String::new(),
            consecutive_count: 1,
            org_code: String::new(),
            org_name: String::new(),
            report_code: String::new(),
            sheet_name,
            rule_id: identity,
            severity: if severity.is_empty() {
                "错误".to_owned()
            } else {
                severity
            },
            formula_cell: target_cell.clone(),
            target_cell,
            target_value: history_value(row, &headers, "当前值"),
            formula_result: serde_json::Value::String(String::new()),
            message: description.clone(),
            source_file,
            audit_file: String::new(),
            check_field: indicator,
            check_type: history_text(row, &headers, "校验类型"),
            comparison_value: history_value(row, &headers, "对比值"),
            reference_value: serde_json::Value::String(String::new()),
            difference_value: history_value(row, &headers, "差值"),
            detail: description,
            institution_feedback: history_text(row, &headers, "历史校验说明"),
            auditor_opinion: history_text(row, &headers, "审核意见"),
        });
    }
    issues
}

fn history_text(
    row: &BTreeMap<usize, CellValue>,
    headers: &HashMap<String, usize>,
    header: &str,
) -> String {
    headers
        .get(header)
        .and_then(|column| row.get(column))
        .map(|value| python_truthy_text(Some(value)))
        .unwrap_or_default()
}

fn history_value(
    row: &BTreeMap<usize, CellValue>,
    headers: &HashMap<String, usize>,
    header: &str,
) -> serde_json::Value {
    let Some(column) = headers.get(header) else {
        return serde_json::Value::String(String::new());
    };
    row.get(column)
        .map(cell_value_to_json)
        .unwrap_or(serde_json::Value::Null)
}

fn python_truthy_text(value: Option<&CellValue>) -> String {
    match value {
        None | Some(CellValue::Empty) | Some(CellValue::Boolean(false)) => String::new(),
        Some(CellValue::Integer(0)) => String::new(),
        Some(CellValue::Number(value)) if *value == 0.0 => String::new(),
        Some(CellValue::Text(value)) => value.clone(),
        Some(CellValue::Number(value)) | Some(CellValue::DateTimeSerial(value)) => {
            python_float_text(*value)
        }
        Some(CellValue::Integer(value)) => value.to_string(),
        Some(CellValue::Boolean(value)) => if *value { "True" } else { "False" }.to_owned(),
        Some(CellValue::DateTimeIso(value)) => python_datetime_text(value),
        Some(CellValue::DurationIso(value)) | Some(CellValue::Error(value)) => value.clone(),
    }
}

fn python_float_text(value: f64) -> String {
    if value == 0.0 {
        return if value.is_sign_negative() {
            "-0.0".to_owned()
        } else {
            "0.0".to_owned()
        };
    }
    let text = value.to_string();
    if text.contains('.') || text.contains('e') || text.contains('E') {
        text
    } else {
        format!("{text}.0")
    }
}

fn python_datetime_text(value: &str) -> String {
    let (date_time, fraction) = value.split_once('.').unwrap_or((value, ""));
    let mut text = date_time.replace('T', " ");
    let fraction = fraction.trim_end_matches('0');
    if !fraction.is_empty() {
        text.push('.');
        text.push_str(&format!("{fraction:0<6}")[..6]);
    }
    text
}

fn cell_value_to_json(value: &CellValue) -> serde_json::Value {
    match value {
        CellValue::Empty => serde_json::Value::Null,
        CellValue::Text(value) | CellValue::DurationIso(value) | CellValue::Error(value) => {
            serde_json::Value::String(value.clone())
        }
        CellValue::DateTimeIso(value) => {
            let (date_time, fraction) = value.split_once('.').unwrap_or((value, ""));
            let fraction = fraction.trim_end_matches('0');
            let mut iso = date_time.to_owned();
            if !fraction.is_empty() {
                iso.push('.');
                iso.push_str(&format!("{fraction:0<6}")[..6]);
            }
            serde_json::Value::String(iso)
        }
        CellValue::Number(value) | CellValue::DateTimeSerial(value) => {
            serde_json::Number::from_f64(*value)
                .map(serde_json::Value::Number)
                .unwrap_or(serde_json::Value::Null)
        }
        CellValue::Integer(value) => serde_json::Value::from(*value),
        CellValue::Boolean(value) => serde_json::Value::Bool(*value),
    }
}

fn parse_cell_address(address: &str) -> Option<(usize, usize)> {
    let split = address.bytes().position(|byte| byte.is_ascii_digit())?;
    let (letters, digits) = address.split_at(split);
    if letters.is_empty() || !letters.bytes().all(|byte| byte.is_ascii_alphabetic()) {
        return None;
    }
    let mut column = 0usize;
    for letter in letters.bytes() {
        let value = usize::from(
            letter
                .to_ascii_uppercase()
                .checked_sub(b'A')?
                .checked_add(1)?,
        );
        column = column.checked_mul(26)?.checked_add(value)?;
    }
    let row = digits.parse::<usize>().ok()?;
    (row > 0 && column > 0).then_some((row, column))
}

/// Carry forward user-entered notes and merge the new rows into the history set.
///
/// Existing rows are keyed by `issue_id`, matching the Python Audit workflow.
/// Additions replace rows with the same key after inheriting distinct notes.
pub fn merge_history(existing: &[Issue], additions: &mut [Issue]) -> Vec<Issue> {
    let mut grouped: HashMap<&str, Vec<&Issue>> = HashMap::new();
    for item in existing {
        grouped.entry(&item.issue_id).or_default().push(item);
    }

    for item in additions.iter_mut() {
        if let Some(previous) = grouped.get(item.issue_id.as_str()) {
            item.institution_feedback = combine_notes(previous, |row| &row.institution_feedback);
            item.auditor_opinion = combine_notes(previous, |row| &row.auditor_opinion);
        }
    }

    // Python dict replacement keeps the first insertion position for a key.
    let mut merged = Vec::<Issue>::new();
    let mut indexes = HashMap::<String, usize>::new();
    for item in existing.iter().chain(additions.iter()) {
        if let Some(index) = indexes.get(&item.issue_id).copied() {
            merged[index] = item.clone();
        } else {
            indexes.insert(item.issue_id.clone(), merged.len());
            merged.push(item.clone());
        }
    }
    merged.sort_by_key(|item| {
        (
            source_basename(&item.source_file),
            item.sheet_name.clone(),
            item.formula_cell.clone(),
            item.issue_id.clone(),
        )
    });
    merged
}

/// Classify current issues against triggered history and synthesize resolved rows.
pub fn classify_current_issues(
    current_period: &str,
    current: &mut [Issue],
    history: &[Issue],
    audited_org_codes: Option<&HashSet<String>>,
    batch_id: &str,
    audit_time: &str,
) -> (Vec<Issue>, String) {
    let old_rows = history
        .iter()
        .filter(|row| row.period != current_period)
        .collect::<Vec<_>>();
    let old_triggered = old_rows
        .iter()
        .copied()
        .filter(|row| row.triggered)
        .collect::<Vec<_>>();

    let mut periods = old_rows
        .iter()
        .map(|row| row.period.as_str())
        .collect::<Vec<_>>();
    periods.sort_by(|left, right| period_key(left).cmp(&period_key(right)));
    periods.dedup();
    let previous_period = periods.last().copied().unwrap_or_default().to_owned();

    let mut by_id: HashMap<&str, Vec<&Issue>> = HashMap::new();
    for row in &old_triggered {
        by_id.entry(&row.issue_id).or_default().push(row);
    }
    for rows in by_id.values_mut() {
        rows.sort_by(|left, right| period_key(&left.period).cmp(&period_key(&right.period)));
    }

    let current_ids = current
        .iter()
        .map(|item| item.issue_id.clone())
        .collect::<HashSet<_>>();
    for item in current.iter_mut() {
        match by_id.get(item.issue_id.as_str()) {
            None => {
                item.status = "新增".to_owned();
                item.first_seen_period = current_period.to_owned();
                item.previous_seen_period.clear();
                item.consecutive_count = 1;
            }
            Some(prior) => {
                let first = prior.first().expect("history group is nonempty");
                let last = prior.last().expect("history group is nonempty");
                item.first_seen_period = if first.first_seen_period.is_empty() {
                    first.period.clone()
                } else {
                    first.first_seen_period.clone()
                };
                item.previous_seen_period = last.period.clone();
                if !previous_period.is_empty() && last.period == previous_period {
                    item.status = "连续出现".to_owned();
                    item.consecutive_count = last.consecutive_count.max(1) + 1;
                } else {
                    item.status = "再次出现".to_owned();
                    item.consecutive_count = 1;
                }
                item.institution_feedback = last.institution_feedback.clone();
                item.auditor_opinion = last.auditor_opinion.clone();
            }
        }
    }

    let mut latest_by_id = Vec::<&Issue>::new();
    let mut latest_indexes = HashMap::<&str, usize>::new();
    if !previous_period.is_empty() {
        for row in &old_triggered {
            if row.period == previous_period {
                if let Some(index) = latest_indexes.get(row.issue_id.as_str()).copied() {
                    latest_by_id[index] = row;
                } else {
                    latest_indexes.insert(&row.issue_id, latest_by_id.len());
                    latest_by_id.push(row);
                }
            }
        }
    }

    let fallback_batch = current
        .first()
        .map(|item| item.batch_id.as_str())
        .unwrap_or("");
    let fallback_audit_time = current
        .first()
        .map(|item| item.audit_time.as_str())
        .unwrap_or("");
    let mut resolved = Vec::new();
    for old in latest_by_id {
        let in_scope = audited_org_codes.is_none_or(|codes| codes.contains(&old.org_code));
        if in_scope && !current_ids.contains(&old.issue_id) {
            let mut item = old.clone();
            item.period = current_period.to_owned();
            item.batch_id = if batch_id.is_empty() {
                fallback_batch.to_owned()
            } else {
                batch_id.to_owned()
            };
            item.audit_time = if audit_time.is_empty() {
                fallback_audit_time.to_owned()
            } else {
                audit_time.to_owned()
            };
            item.triggered = false;
            item.status = "已整改".to_owned();
            item.previous_seen_period = old.period.clone();
            item.consecutive_count = 0;
            item.target_value = serde_json::Value::String(String::new());
            item.comparison_value = serde_json::Value::String(String::new());
            item.reference_value = serde_json::Value::String(String::new());
            item.difference_value = serde_json::Value::String(String::new());
            item.formula_result = serde_json::Value::String(String::new());
            resolved.push(item);
        }
    }

    (resolved, previous_period)
}

/// Apply the same ordering as the source workflow: inherit notes, then classify.
pub fn process_s1_history(request: S1HistoryRequest) -> S1HistoryResult {
    let mut current = request.current_issues;
    let history_rows = request.history.len();
    let mut merged_history = merge_history(&request.history, &mut current);
    let (resolved_issues, previous_period) = classify_current_issues(
        &request.period,
        &mut current,
        &request.history,
        request.audited_org_codes.as_ref(),
        &request.batch_id,
        &request.audit_time,
    );
    let merged_indexes = merged_history
        .iter()
        .enumerate()
        .map(|(index, item)| (item.issue_id.clone(), index))
        .collect::<HashMap<_, _>>();
    for item in &current {
        if let Some(index) = merged_indexes.get(&item.issue_id).copied() {
            // The Python workflow's merged list shares these Issue objects;
            // classification updates are visible through that list as well.
            merged_history[index] = item.clone();
        }
    }
    merged_history.sort_by_key(|item| {
        (
            source_basename(&item.source_file),
            item.sheet_name.clone(),
            item.formula_cell.clone(),
            item.issue_id.clone(),
        )
    });
    let summary = S1HistorySummary {
        history_rows,
        merged_history_rows: merged_history.len(),
        current_issues: current.len(),
        new_issues: current.iter().filter(|item| item.status == "新增").count(),
        continuous_issues: current
            .iter()
            .filter(|item| item.status == "连续出现")
            .count(),
        recurring_issues: current
            .iter()
            .filter(|item| item.status == "再次出现")
            .count(),
        resolved_issues: resolved_issues.len(),
        previous_period,
        elapsed_ms: 0,
    };
    S1HistoryResult {
        current_issues: current,
        resolved_issues,
        merged_history,
        summary,
    }
}

fn combine_notes<'a>(rows: &[&'a Issue], field: impl Fn(&'a Issue) -> &'a str) -> String {
    let mut values = Vec::new();
    for row in rows {
        let value = field(row).trim();
        if !value.is_empty() && !values.iter().any(|existing| *existing == value) {
            values.push(value);
        }
    }
    values.join("\n")
}

fn source_basename(path: &str) -> String {
    path.rsplit(['/', '\\']).next().unwrap_or(path).to_owned()
}

fn period_key(value: &str) -> (i32, i32, &str) {
    let text = value.trim();
    if let Some((year, month)) = parse_year_month(text) {
        return (year, month, text);
    }
    let uppercase = text.to_uppercase();
    let chars = uppercase.chars().collect::<Vec<_>>();
    if chars.len() >= 6 {
        if let Some(index) = chars.iter().position(|value| *value == 'Q') {
            let year_text = chars[..4].iter().collect::<String>();
            let quarter_text = chars.get(index + 1).copied().unwrap_or_default();
            if year_text.chars().all(|value| value.is_ascii_digit())
                && quarter_text.is_ascii_digit()
            {
                let year = year_text.parse::<i32>().unwrap_or_default();
                let quarter = quarter_text.to_digit(10).unwrap_or_default() as i32;
                return (year, quarter * 3, text);
            }
        }
    }
    (0, 0, text)
}

fn parse_year_month(value: &str) -> Option<(i32, i32)> {
    let (year_text, month_text, day_text) =
        if value.len() == 6 && value.chars().all(|value| value.is_ascii_digit()) {
            (&value[..4], &value[4..], None)
        } else {
            let separator = if value.contains('-') {
                '-'
            } else if value.contains('/') {
                '/'
            } else {
                return None;
            };
            let parts = value.split(separator).collect::<Vec<_>>();
            match parts.as_slice() {
                [year, month] => (*year, *month, None),
                [year, month, day] => (*year, *month, Some(*day)),
                _ => return None,
            }
        };
    if year_text.len() != 4
        || !year_text.chars().all(|value| value.is_ascii_digit())
        || month_text.is_empty()
        || month_text.len() > 2
        || !month_text.chars().all(|value| value.is_ascii_digit())
    {
        return None;
    }
    let year = year_text.parse::<i32>().ok()?;
    let month = month_text.parse::<i32>().ok()?;
    if !(1..=12).contains(&month) {
        return None;
    }
    if let Some(day_text) = day_text {
        if day_text.is_empty()
            || day_text.len() > 2
            || !day_text.chars().all(|value| value.is_ascii_digit())
        {
            return None;
        }
        let day = day_text.parse::<u32>().ok()?;
        let max_day = days_in_month(year, month as u32)?;
        if !(1..=max_day).contains(&day) {
            return None;
        }
    }
    Some((year, month))
}

fn days_in_month(year: i32, month: u32) -> Option<u32> {
    let leap = year % 4 == 0 && (year % 100 != 0 || year % 400 == 0);
    match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => Some(31),
        4 | 6 | 9 | 11 => Some(30),
        2 => Some(if leap { 29 } else { 28 }),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use excel_editor::{
        BackendCapability, CapabilityStatus, EditBackend, EngineId, OpenpyxlXlsxValidator,
        OutputValidator, UmyaXlsxBackend,
    };
    use serde_json::json;

    fn issue(issue_id: &str, period: &str, triggered: bool) -> Issue {
        Issue {
            issue_id: issue_id.to_owned(),
            period: period.to_owned(),
            batch_id: "batch-old".to_owned(),
            audit_time: "old-time".to_owned(),
            triggered,
            status: "新增".to_owned(),
            first_seen_period: period.to_owned(),
            previous_seen_period: String::new(),
            consecutive_count: 1,
            org_code: "001".to_owned(),
            org_name: "测试机构".to_owned(),
            report_code: "R01".to_owned(),
            sheet_name: "Sheet1".to_owned(),
            rule_id: issue_id.to_owned(),
            severity: "错误".to_owned(),
            formula_cell: "C2".to_owned(),
            target_cell: "B2".to_owned(),
            target_value: json!(1),
            formula_result: json!("错误"),
            message: "测试".to_owned(),
            source_file: "source.xlsx".to_owned(),
            audit_file: "audit.xlsx".to_owned(),
            check_field: String::new(),
            check_type: String::new(),
            comparison_value: json!(""),
            reference_value: json!(""),
            difference_value: json!(""),
            detail: String::new(),
            institution_feedback: String::new(),
            auditor_opinion: String::new(),
        }
    }

    fn history_fixture_path(name: &str) -> (std::path::PathBuf, std::path::PathBuf) {
        let directory = std::env::temp_dir().join(format!(
            "audit_s1_history_plan_{}_{}",
            std::process::id(),
            name
        ));
        fs::create_dir_all(&directory).expect("create S1 history plan fixture directory");
        (directory.clone(), directory.join("input.xlsx"))
    }

    fn save_history_plan_fixture(path: &Path, sheet_names: &[&str]) {
        let mut workbook = XlsxWorkbook::new();
        for sheet_name in sheet_names {
            let worksheet = workbook.add_worksheet();
            worksheet
                .set_name(*sheet_name)
                .expect("set history plan fixture sheet name");
            worksheet
                .write_string(0, 0, format!("{sheet_name} existing"))
                .expect("write history plan fixture cell");
        }
        workbook.save(path).expect("save history plan fixture");
    }

    #[test]
    fn history_workbook_edit_plan_replaces_only_the_managed_sheet() {
        let (directory, input) = history_fixture_path("replace");
        let output = directory.join("updated.xlsx");
        save_history_plan_fixture(&input, &[HISTORY_AUDIT_SHEET, "业务说明"]);
        let original = fs::read(&input).expect("read original history plan fixture");
        let mut current = issue("RULE-1", "2026-10", true);
        current.target_value = json!(0);
        current.comparison_value = json!(false);
        current.difference_value = serde_json::Value::Null;

        let result = build_s1_history_workbook_edit_plan(&input, &output, &[current])
            .expect("build history workbook edit plan");
        let plan = result.edit_plan;
        assert!(result.replaces_existing_history_sheet);
        assert_eq!(result.issue_count, 1);
        assert_eq!(plan.input, fs::canonicalize(&input).unwrap());
        assert_eq!(
            plan.output,
            fs::canonicalize(&directory).unwrap().join("updated.xlsx")
        );
        assert!(plan.workbook.inventory_complete);
        assert!(matches!(
            plan.operations.first(),
            Some(EditOperation::DeleteSheet { name }) if name == HISTORY_AUDIT_SHEET
        ));
        assert!(matches!(
            plan.operations.get(1),
            Some(EditOperation::CreateSheetAt { name, index: 0 }) if name == HISTORY_AUDIT_SHEET
        ));
        assert!(plan.operations.iter().any(|operation| matches!(
            operation,
            EditOperation::SetCellValue {
                sheet,
                cell,
                value: EditCellValue::Decimal(value),
            } if sheet == HISTORY_AUDIT_SHEET && cell == "H2" && value == "0"
        )));
        assert!(plan.operations.iter().any(|operation| matches!(
            operation,
            EditOperation::SetCellValue {
                sheet,
                cell,
                value: EditCellValue::Boolean(false),
            } if sheet == HISTORY_AUDIT_SHEET && cell == "I2"
        )));
        assert!(!plan.operations.iter().any(|operation| matches!(
            operation,
            EditOperation::SetCellValue { cell, .. } if cell == "J2"
        )));
        assert!(plan.operations.iter().any(|operation| matches!(
            operation,
            EditOperation::SetAutoFilter { sheet, range: Some(range) }
                if sheet == HISTORY_AUDIT_SHEET && range == "A1:M2"
        )));
        assert!(plan.operations.iter().any(|operation| matches!(
            operation,
            EditOperation::SetFreezePanes { sheet, cell: Some(cell) }
                if sheet == HISTORY_AUDIT_SHEET && cell == "A2"
        )));
        assert_eq!(
            plan.operations
                .iter()
                .filter(|operation| matches!(operation, EditOperation::SetColumnWidth { .. }))
                .count(),
            S1_HISTORY_HEADERS.len()
        );
        assert!(
            plan.preserve
                .contains(&PreservationRequirement::CellValuesAndTypes)
        );
        assert!(
            plan.preserve
                .contains(&PreservationRequirement::FormulaText)
        );
        assert_eq!(
            plan.drop_allowed.get(HISTORY_AUDIT_SHEET),
            Some(&plan.preserve)
        );
        assert_eq!(fs::read(&input).unwrap(), original);
        assert!(!output.exists());
        fs::remove_dir_all(directory).expect("remove S1 history plan fixture directory");
    }

    #[test]
    fn history_workbook_edit_plan_creates_the_managed_sheet_when_missing() {
        let (directory, input) = history_fixture_path("create");
        let output = directory.join("updated.xlsx");
        save_history_plan_fixture(&input, &["业务说明"]);

        let result = build_s1_history_workbook_edit_plan(&input, &output, &[])
            .expect("build history workbook create plan");
        assert!(!result.replaces_existing_history_sheet);
        assert!(matches!(
            result.edit_plan.operations.first(),
            Some(EditOperation::CreateSheetAt { name, index: 0 }) if name == HISTORY_AUDIT_SHEET
        ));
        assert!(matches!(
            result.edit_plan.operations.iter().find(|operation| matches!(
                operation,
                EditOperation::SetAutoFilter { .. }
            )),
            Some(EditOperation::SetAutoFilter { range: Some(range), .. }) if range == "A1:M1"
        ));
        assert!(result.edit_plan.drop_allowed.is_empty());
        assert!(!output.exists());
        fs::remove_dir_all(directory).expect("remove S1 history plan fixture directory");
    }

    #[test]
    #[ignore = "set AUDIT_S1_HISTORY_REAL_WORKBOOK to an authorized XLSX sample; candidate only"]
    fn real_s1_history_workbook_engine_candidate_only() {
        let source = std::env::var_os("AUDIT_S1_HISTORY_REAL_WORKBOOK")
            .map(std::path::PathBuf::from)
            .expect("AUDIT_S1_HISTORY_REAL_WORKBOOK must identify a read-only sample");
        assert!(
            source.is_file(),
            "real S1 history workbook sample is missing"
        );
        let source_bytes = fs::read(&source).expect("read source sample for immutability check");
        let (directory, input) = history_fixture_path("real_candidate");
        fs::copy(&source, &input).expect("copy real sample into the temporary candidate directory");
        let input_bytes = fs::read(&input).expect("read temporary candidate input");
        let output = directory.join("candidate-output.xlsx");
        let mut candidate_issue = issue("CANDIDATE-ONLY", "2026-10", true);
        candidate_issue.source_file = "candidate-source.xlsx".to_owned();
        candidate_issue.target_value = json!(0);
        let result = build_s1_history_workbook_edit_plan(&input, &output, &[candidate_issue])
            .expect("build real-sample history edit plan");
        let plan = &result.edit_plan;
        let sheet_count = excel_reader::read_sheet_names(&input)
            .expect("read real-sample worksheet names")
            .len();
        println!(
            "candidate_profile format={:?} inventory_complete={} features={:?} history_sheet_present={} other_sheet_count={} operations={:?} preserve={:?}",
            plan.workbook.format,
            plan.workbook.inventory_complete,
            plan.workbook.features_present,
            result.replaces_existing_history_sheet,
            sheet_count.saturating_sub(usize::from(result.replaces_existing_history_sheet)),
            plan.operations
                .iter()
                .map(EditOperation::kind)
                .collect::<Vec<_>>(),
            plan.preserve
        );

        let capability = BackendCapability {
            engine: EngineId::UmyaSpreadsheet,
            status: CapabilityStatus::Unverified,
            format: plan.format,
            operations: plan.operations.iter().map(EditOperation::kind).collect(),
            supported_features: plan.workbook.features_present.clone(),
            preserved: plan.preserve.clone(),
            platform: std::env::consts::OS.to_owned(),
            evidence: None,
        };
        let backend = UmyaXlsxBackend::new(capability).expect("construct unverified candidate");
        match backend.apply(plan) {
            Ok(()) => match OpenpyxlXlsxValidator.validate(plan, &output) {
                Ok(()) => println!(
                    "candidate_validation=passed engine=umya_spreadsheet production_capability=unverified"
                ),
                Err(error) => println!(
                    "candidate_validation=rejected engine=umya_spreadsheet production_capability=unverified reason={error:#}"
                ),
            },
            Err(error) => println!(
                "candidate_apply=rejected engine=umya_spreadsheet kind={:?} production_capability=unverified reason={}",
                error.kind, error.message
            ),
        }
        assert_eq!(fs::read(&input).unwrap(), input_bytes);
        assert_eq!(fs::read(&source).unwrap(), source_bytes);
        fs::remove_dir_all(directory).expect("remove temporary real-sample candidate directory");
    }

    #[test]
    fn classifies_new_continuous_recurring_and_resolved_issues() {
        let continuous = issue("CONT", "2026-03", true);
        let recurring = issue("AGAIN", "2026-03", true);
        let new = issue("NEW", "2026-03", true);
        let mut history_cont = issue("CONT", "2026-02", true);
        history_cont.first_seen_period = "2025-12".to_owned();
        history_cont.consecutive_count = 2;
        let old_again = issue("AGAIN", "2026-01", true);
        let resolved = issue("RESOLVED", "2026-02", true);
        let history = vec![history_cont, old_again, resolved];
        let mut current = vec![continuous, recurring, new];
        let scope = HashSet::from(["001".to_owned()]);

        let (resolved_issues, previous_period) = classify_current_issues(
            "2026-03",
            &mut current,
            &history,
            Some(&scope),
            "batch-new",
            "new-time",
        );

        assert_eq!(previous_period, "2026-02");
        assert_eq!(current[0].status, "连续出现");
        assert_eq!(current[0].first_seen_period, "2025-12");
        assert_eq!(current[0].consecutive_count, 3);
        assert_eq!(current[1].status, "再次出现");
        assert_eq!(current[1].consecutive_count, 1);
        assert_eq!(current[2].status, "新增");
        assert_eq!(resolved_issues.len(), 1);
        assert_eq!(resolved_issues[0].issue_id, "RESOLVED");
        assert_eq!(resolved_issues[0].status, "已整改");
        assert_eq!(resolved_issues[0].batch_id, "batch-new");
        assert!(!resolved_issues[0].triggered);
        assert_eq!(resolved_issues[0].target_value, json!(""));
    }

    #[test]
    fn merges_distinct_manual_notes_and_replaces_history_by_issue_id() {
        let mut first = issue("SAME", "2026-01", true);
        first.institution_feedback = "  机构说明甲  ".to_owned();
        first.auditor_opinion = "审核意见".to_owned();
        let mut second = issue("SAME", "2026-02", true);
        second.institution_feedback = "机构说明乙".to_owned();
        second.auditor_opinion = "审核意见".to_owned();
        let current = issue("SAME", "2026-03", true);
        let mut additions = vec![current];

        let merged = merge_history(&[first, second], &mut additions);

        assert_eq!(additions[0].institution_feedback, "机构说明甲\n机构说明乙");
        assert_eq!(additions[0].auditor_opinion, "审核意见");
        assert_eq!(merged.len(), 1);
        assert_eq!(merged[0].period, "2026-03");
    }

    #[test]
    fn history_writer_rows_match_python_column_order_and_keep_scalar_types() {
        let mut current = issue("RULE-7", "2026-03", true);
        current.source_file = Path::new("inputs")
            .join("机构甲.xlsx")
            .display()
            .to_string();
        current.check_type = "公式".to_owned();
        current.check_field = "资产总额".to_owned();
        current.detail.clear();
        current.message = "指标不匹配".to_owned();
        current.target_value = json!(12.5);
        current.comparison_value = json!(false);
        current.difference_value = serde_json::Value::Null;
        current.institution_feedback = "机构反馈".to_owned();
        current.auditor_opinion = "审核意见".to_owned();

        let headers = s1_history_header_row();
        let row = s1_history_workbook_row(&current).unwrap();

        assert_eq!(
            headers,
            S1_HISTORY_HEADERS.map(|header| EditCellValue::Text(header.to_owned()))
        );
        assert_eq!(row[0], EditCellValue::Text("机构甲.xlsx".to_owned()));
        assert_eq!(row[1], EditCellValue::Text("Sheet1".to_owned()));
        assert_eq!(row[2], EditCellValue::Text("B2".to_owned()));
        assert_eq!(row[4], EditCellValue::Text("公式".to_owned()));
        assert_eq!(row[5], EditCellValue::Text("资产总额".to_owned()));
        assert_eq!(row[6], EditCellValue::Text("指标不匹配".to_owned()));
        assert_eq!(row[7], EditCellValue::Decimal("12.5".to_owned()));
        assert_eq!(row[8], EditCellValue::Boolean(false));
        assert_eq!(row[9], EditCellValue::Blank);
        assert_eq!(row[10], EditCellValue::Text("RULE-7".to_owned()));
        assert_eq!(row[11], EditCellValue::Text("机构反馈".to_owned()));
        assert_eq!(row[12], EditCellValue::Text("审核意见".to_owned()));
    }

    #[test]
    fn history_writer_rejects_non_scalar_excel_values() {
        let mut current = issue("RULE-ARRAY", "2026-03", true);
        current.target_value = json!([1, 2]);

        let error = s1_history_workbook_row(&current).unwrap_err();

        assert!(error.to_string().contains("当前值"));
    }

    #[test]
    fn history_result_writer_creates_a_typed_standalone_workbook() {
        let directory = std::env::temp_dir().join(format!(
            "audit_s1_history_writer_{}_{}",
            std::process::id(),
            SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
        fs::create_dir_all(&directory).unwrap();
        let output = directory.join("history.xlsx");
        let mut current = issue("RULE-7", "2026-03", true);
        current.source_file = Path::new("inputs")
            .join("机构甲.xlsx")
            .display()
            .to_string();
        current.target_value = json!(12.5);
        current.comparison_value = json!(false);
        current.difference_value = serde_json::Value::Null;

        let summary = write_s1_history_workbook(&output, &[current]).unwrap();

        assert_eq!(summary.issue_count, 1);
        assert_eq!(summary.sheet_name, HISTORY_AUDIT_SHEET);
        assert_eq!(summary.engine, "rust_xlsxwriter 0.101.0");
        let workbook = excel_reader::read_workbook(&output).unwrap();
        assert_eq!(workbook.sheets.len(), 1);
        let sheet = &workbook.sheets[0];
        assert_eq!(sheet.name, HISTORY_AUDIT_SHEET);
        assert_eq!(sheet.values["A1"]["value"], "工作簿名");
        assert_eq!(sheet.values["A2"]["value"], "机构甲.xlsx");
        assert_eq!(sheet.values["H2"]["type"], "number");
        assert_eq!(sheet.values["H2"]["value"], 12.5);
        assert_eq!(sheet.values["I2"]["type"], "boolean");
        assert_eq!(sheet.values["I2"]["value"], false);
        assert!(!sheet.values.contains_key("J2"));
        assert!(write_s1_history_workbook(&output, &[]).is_err());
        fs::remove_dir_all(directory).unwrap();
    }

    #[test]
    fn period_sorting_matches_supported_month_date_and_quarter_forms() {
        let mut periods = vec!["2026-Q1", "2025-12", "2026-01-31", "2026/02", "bad"];
        periods.sort_by(|left, right| period_key(left).cmp(&period_key(right)));
        assert_eq!(
            periods,
            vec!["bad", "2025-12", "2026-01-31", "2026/02", "2026-Q1"]
        );
        assert_eq!(period_key("2026-02-29").0, 0);
        assert_eq!(period_key("2024-02-29").0, 2024);
    }

    #[test]
    fn empty_current_uses_requested_batch_and_timestamp_for_resolved_rows() {
        let history = vec![issue("OLD", "2026-01", true)];
        let request = S1HistoryRequest {
            period: "2026-02".to_owned(),
            batch_id: "batch-new".to_owned(),
            audit_time: "new-time".to_owned(),
            audited_org_codes: None,
            current_issues: Vec::new(),
            history,
        };
        let result = process_s1_history(request);
        assert_eq!(result.resolved_issues[0].batch_id, "batch-new");
        assert_eq!(result.resolved_issues[0].audit_time, "new-time");
    }
}
