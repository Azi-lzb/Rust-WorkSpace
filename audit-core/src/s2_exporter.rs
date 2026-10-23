//! Candidate Rust writer for the three S2-F01 result worksheets.
//!
//! This module mirrors the source `ExcelExporter` table contract. It remains a
//! candidate until its real-template and file-level parity gates are accepted.

use crate::{
    s2_import::ImportedValue, s2_validation::S2AuditFinding, unit_conversion::DecimalTuple,
};
use anyhow::{bail, ensure, Context, Result};
use chrono::Local;
use excel_reader::read_workbook;
use excel_writer::{Color, Format, FormatAlign, Workbook, Worksheet};
use serde_json::{Number, Value};
use std::{
    fs::{self, OpenOptions},
    path::{Path, PathBuf},
    sync::atomic::{AtomicU64, Ordering},
};

const STATUS_NORMAL: &str = "正常";
const STATUS_ABNORMAL: &str = "异常";
const SEVERITIES: &[&str] = &["硬性", "软性"];
const MAX_TEXT_CHARS: usize = 32_767;
const MAX_COLUMN_WIDTH: usize = 42;

const SHEET_ORDER: [&str; 3] = ["环比规则", "外部核对规则", "校验规则"];
const PERIOD_HEADERS: [&str; 16] = [
    "地区",
    "承接行",
    "机构类别",
    "机构名称",
    "社会信用代码",
    "表单",
    "指标代码",
    "数据属性",
    "指标名称",
    "审核级别",
    "审核说明",
    "本期值",
    "上期值",
    "差异值",
    "变动幅度",
    "计算过程",
];
const EXTERNAL_HEADERS: [&str; 16] = [
    "地区",
    "承接行",
    "机构类别",
    "机构名称",
    "社会信用代码",
    "报表项目",
    "表单",
    "指标代码",
    "指标名称",
    "外部指标名称",
    "审核级别",
    "审核说明",
    "报表值",
    "外部值",
    "差异值",
    "计算过程",
];
const VALIDATION_HEADERS: [&str; 12] = [
    "地区",
    "承接行",
    "机构类别",
    "机构名称",
    "社会信用代码",
    "涉及表单",
    "规则编号",
    "规则类型",
    "审核级别",
    "审核说明",
    "取值明细",
    "计算过程",
];
const LONG_TEXT_HEADERS: &[&str] = &["审核说明", "取值明细", "计算过程", "报表项目"];

static TEMP_COUNTER: AtomicU64 = AtomicU64::new(0);

/// Write a candidate three-sheet S2 report and return its final path.
///
/// The result is first written to a uniquely named sibling file, reopened via
/// the shared Calamine reader, and only then published with a filesystem
/// rename. Decimal calculations stay exact upstream; XLSX numeric cells are
/// emitted as Excel's IEEE-754 numeric values, matching the source workbook's
/// numeric-cell contract. Cross-writer parity uses the workspace's declared
/// floating-point comparison policy.
pub fn export_s2_report(
    output_dir: impl AsRef<Path>,
    findings: &[S2AuditFinding],
    hide_unchanged_period_rows: bool,
    hide_matching_external_rows: bool,
    on_progress: Option<&mut dyn FnMut(&str, usize, usize)>,
) -> Result<PathBuf> {
    let output_dir = output_dir.as_ref();
    fs::create_dir_all(output_dir)
        .with_context(|| format!("创建 S2 报告目录失败：{}", output_dir.display()))?;

    let tables = build_tables(
        findings,
        hide_unchanged_period_rows,
        hide_matching_external_rows,
    )?;
    let timestamp = Local::now().format("%Y%m%d_%H%M%S").to_string();
    let result_path = output_dir.join(format!("报表审核清单_{timestamp}.xlsx"));
    let mut temporary = TemporaryWorkbook::create(output_dir, &result_path)?;
    let expected_formulas = tables
        .iter()
        .flat_map(|table| &table.rows)
        .flatten()
        .filter(|cell| matches!(cell, ExcelCell::Formula(_)))
        .count();

    match on_progress {
        Some(progress) => write_workbook(&temporary.path, &tables, progress)?,
        None => {
            let mut no_progress = |_: &str, _: usize, _: usize| {};
            write_workbook(&temporary.path, &tables, &mut no_progress)?;
        }
    }
    validate_workbook(&temporary.path, expected_formulas)?;
    temporary.publish(&result_path)?;
    Ok(result_path)
}

#[derive(Clone, Debug, PartialEq)]
enum ExcelCell {
    Blank,
    Text(String),
    Number { value: f64, display: String },
    Boolean(bool),
    Formula(String),
}

impl ExcelCell {
    fn width_text(&self) -> &str {
        match self {
            Self::Blank | Self::Boolean(false) => "",
            Self::Text(value) | Self::Formula(value) => value,
            Self::Number { value, .. } if *value == 0.0 => "",
            Self::Number { display, .. } => display,
            Self::Boolean(true) => "True",
        }
    }

    fn is_hidden_description(&self, expected: &str) -> bool {
        matches!(self, Self::Text(value) if value == expected)
    }
}

struct Table {
    name: &'static str,
    headers: &'static [&'static str],
    rows: Vec<Vec<ExcelCell>>,
    percentage_point_rows: Vec<bool>,
    hidden_description: Option<&'static str>,
}

fn build_tables(
    findings: &[S2AuditFinding],
    hide_unchanged_period_rows: bool,
    hide_matching_external_rows: bool,
) -> Result<Vec<Table>> {
    let mut period = Vec::new();
    let mut period_points = Vec::new();
    let mut external = Vec::new();
    let mut validation = Vec::new();

    for item in findings {
        validate_finding(item)?;
        match item.audit_type.as_str() {
            "环比" | "环比规则" => {
                period.push(period_values(item)?);
                period_points.push(item.value_type == "百分数");
            }
            "外部核对" | "外部核对规则" => external.push(external_values(item)?),
            "规则" | "校验规则" => validation.push(validation_values(item)?),
            _ => bail!(
                "S2-F01 遇到未知审核类型，未生成结果：{} finding_id={}",
                python_repr(&item.audit_type),
                item.finding_id
            ),
        }
    }

    Ok(vec![
        Table {
            name: SHEET_ORDER[0],
            headers: &PERIOD_HEADERS,
            rows: period,
            percentage_point_rows: period_points,
            hidden_description: hide_unchanged_period_rows.then_some("无变动"),
        },
        Table {
            name: SHEET_ORDER[1],
            headers: &EXTERNAL_HEADERS,
            rows: external,
            percentage_point_rows: Vec::new(),
            hidden_description: hide_matching_external_rows.then_some("与大集中系统数据一致"),
        },
        Table {
            name: SHEET_ORDER[2],
            headers: &VALIDATION_HEADERS,
            rows: validation,
            percentage_point_rows: Vec::new(),
            hidden_description: None,
        },
    ])
}

fn validate_finding(item: &S2AuditFinding) -> Result<()> {
    if item.status == STATUS_NORMAL {
        ensure!(
            item.severity.is_empty(),
            "S2-F01 审核级别冲突：正常 finding 不应有审核级别；finding_id={} 机构={} 表单={} 规则={} 指标={} 状态={} 级别={}",
            item.finding_id,
            item.institution_name,
            item.form_code,
            item.rule_id,
            item.indicator_code,
            item.status,
            item.severity
        );
        return Ok(());
    }
    ensure!(
        item.status == STATUS_ABNORMAL && SEVERITIES.contains(&item.severity.as_str()),
        "S2-F01 审核级别冲突：异常 finding 必须是硬性/软性；finding_id={} 机构={} 表单={} 规则={} 指标={} 状态={} 级别={}",
        item.finding_id,
        item.institution_name,
        item.form_code,
        item.rule_id,
        item.indicator_code,
        item.status,
        item.severity
    );
    Ok(())
}

fn period_values(item: &S2AuditFinding) -> Result<Vec<ExcelCell>> {
    Ok(vec![
        text_cell(&item.region, false)?,
        text_cell(&item.handling_branch, false)?,
        text_cell(&item.institution_type, false)?,
        text_cell(&item.institution_name, false)?,
        text_cell(&item.social_credit_code, false)?,
        text_cell(&form_label(item), false)?,
        text_cell(&item.indicator_code, false)?,
        text_cell(&item.data_type, false)?,
        text_cell(&item.indicator_name, false)?,
        text_cell(&item.severity, false)?,
        text_cell(
            first_nonempty(&item.rule_description, &item.description),
            false,
        )?,
        value_cell(&item.current_value, false)?,
        value_cell(&item.previous_value, false)?,
        value_cell(&item.difference_value, false)?,
        value_cell(&item.change_rate, false)?,
        text_cell(&item.calculation_trace, false)?,
    ])
}

fn external_values(item: &S2AuditFinding) -> Result<Vec<ExcelCell>> {
    Ok(vec![
        text_cell(&item.region, false)?,
        text_cell(&item.handling_branch, false)?,
        text_cell(&item.institution_type, false)?,
        text_cell(&item.institution_name, false)?,
        text_cell(&item.social_credit_code, false)?,
        text_cell(&item.report_item, true)?,
        text_cell(&form_label(item), false)?,
        text_cell(&item.indicator_code, false)?,
        text_cell(&item.indicator_name, false)?,
        text_cell(&item.external_indicator_name, false)?,
        text_cell(&item.severity, false)?,
        text_cell(
            first_nonempty(&item.rule_description, &item.description),
            false,
        )?,
        value_cell(&item.current_value, false)?,
        value_cell(&item.external_value, false)?,
        value_cell(&item.difference_value, false)?,
        text_cell(&item.calculation_trace, false)?,
    ])
}

fn validation_values(item: &S2AuditFinding) -> Result<Vec<ExcelCell>> {
    let rule_type = if item.rule_type.is_empty() {
        "未识别"
    } else {
        &item.rule_type
    };
    let value_details = if item.value_details.is_empty() {
        if matches!(rule_type, "导入校验" | "配置检查") {
            "不适用"
        } else {
            "未识别"
        }
    } else {
        &item.value_details
    };
    let institution_missing = item.institution_name.is_empty();

    Ok(vec![
        text_cell(
            if item.region.is_empty() && institution_missing {
                "不适用"
            } else {
                &item.region
            },
            false,
        )?,
        text_cell(
            if item.handling_branch.is_empty() && institution_missing {
                "不适用"
            } else {
                &item.handling_branch
            },
            false,
        )?,
        text_cell(
            if item.institution_type.is_empty() && institution_missing {
                "不适用"
            } else {
                &item.institution_type
            },
            false,
        )?,
        text_cell(
            if institution_missing {
                "全局"
            } else {
                &item.institution_name
            },
            false,
        )?,
        text_cell(
            if item.social_credit_code.is_empty() && institution_missing {
                "不适用"
            } else {
                &item.social_credit_code
            },
            false,
        )?,
        text_cell(&validation_form_scope(item), false)?,
        text_cell(
            if item.rule_id.is_empty() {
                "未配置"
            } else {
                &item.rule_id
            },
            false,
        )?,
        text_cell(rule_type, false)?,
        text_cell(&item.severity, false)?,
        text_cell(
            first_nonempty(&item.rule_description, &item.description),
            false,
        )?,
        text_cell(value_details, false)?,
        text_cell(&item.calculation_trace, false)?,
    ])
}

fn form_label(item: &S2AuditFinding) -> String {
    if item.form_codes.is_empty() {
        item.form_code.clone()
    } else {
        item.form_codes.join("、")
    }
}

fn validation_form_scope(item: &S2AuditFinding) -> String {
    let mut forms = item
        .form_codes
        .iter()
        .chain(std::iter::once(&item.form_code))
        .filter(|value| !value.is_empty())
        .cloned()
        .collect::<std::collections::BTreeSet<_>>();
    for record in &item.related_records {
        if let Some(form_code) = record.get("form_code").and_then(Value::as_str) {
            if !form_code.is_empty() {
                forms.insert(form_code.to_owned());
            }
        }
    }
    if !forms.is_empty() {
        return forms.into_iter().collect::<Vec<_>>().join("、");
    }
    if matches!(item.rule_type.as_str(), "导入校验" | "配置检查") {
        "不适用".to_owned()
    } else {
        "未识别".to_owned()
    }
}

fn first_nonempty<'a>(primary: &'a str, fallback: &'a str) -> &'a str {
    if primary.is_empty() {
        fallback
    } else {
        primary
    }
}

fn text_cell(value: &str, literal: bool) -> Result<ExcelCell> {
    let value = checked_text(value)?;
    if !literal && value.starts_with('=') {
        Ok(ExcelCell::Formula(value))
    } else {
        Ok(ExcelCell::Text(value))
    }
}

fn value_cell(value: &Value, literal: bool) -> Result<ExcelCell> {
    match value {
        Value::Null => Ok(ExcelCell::Blank),
        Value::Bool(value) => Ok(ExcelCell::Boolean(*value)),
        Value::Number(number) => number_cell(number),
        Value::String(value) => text_cell(value, literal),
        Value::Array(_) => bail!("S2-F01 导出值类型不支持：数组"),
        Value::Object(_) => {
            let imported: ImportedValue =
                serde_json::from_value(value.clone()).context("解析 S2 finding 中的带类型值")?;
            imported_value_cell(imported, literal)
        }
    }
}

fn imported_value_cell(value: ImportedValue, literal: bool) -> Result<ExcelCell> {
    match value {
        ImportedValue::Empty => Ok(ExcelCell::Blank),
        ImportedValue::Text(value) | ImportedValue::Error(value) => text_cell(&value, literal),
        ImportedValue::Number(value) => {
            finite_number(value, crate::s2_external::python_float_text(value))
        }
        ImportedValue::DateTimeSerial(_)
        | ImportedValue::DateTimeIso(_)
        | ImportedValue::DurationIso(_) => {
            bail!("S2-F01 导出日期/时长单元格尚未完成对拍，已拒绝生成")
        }
        ImportedValue::Integer(value) => exact_integer_cell(value),
        ImportedValue::Boolean(value) => Ok(ExcelCell::Boolean(value)),
        ImportedValue::Decimal(value) => decimal_cell(&value),
    }
}

fn number_cell(number: &Number) -> Result<ExcelCell> {
    if let Some(value) = number.as_i64() {
        return exact_integer_cell(value);
    }
    if let Some(value) = number.as_u64() {
        if value > i64::MAX as u64 {
            let as_f64 = value as f64;
            ensure!(
                as_f64 as u128 == value as u128,
                "S2-F01 导出整数超出 Excel 可精确保留范围，已拒绝生成"
            );
            return finite_number(as_f64, value.to_string());
        }
        return exact_integer_cell(value as i64);
    }
    let value = number
        .as_f64()
        .context("S2-F01 导出数值无法转换为 Excel 数值")?;
    finite_number(value, crate::s2_external::python_float_text(value))
}

fn exact_integer_cell(value: i64) -> Result<ExcelCell> {
    let as_f64 = value as f64;
    ensure!(
        as_f64 as i64 == value,
        "S2-F01 导出整数超出 Excel 可精确保留范围，已拒绝生成"
    );
    finite_number(as_f64, value.to_string())
}

fn finite_number(value: f64, display: String) -> Result<ExcelCell> {
    ensure!(value.is_finite(), "S2-F01 导出数值不是有限数，已拒绝生成");
    Ok(ExcelCell::Number { value, display })
}

fn decimal_cell(value: &DecimalTuple) -> Result<ExcelCell> {
    let text = decimal_python_text(value)?;
    let number = text
        .parse::<f64>()
        .with_context(|| format!("S2-F01 Decimal 无法写为 Excel 数值：{text}"))?;
    finite_number(number, text)
}

fn decimal_python_text(value: &DecimalTuple) -> Result<String> {
    ensure!(value.sign <= 1, "S2-F01 Decimal 符号无效，已拒绝生成");
    ensure!(
        !value.digits.is_empty() && value.digits.iter().all(|digit| *digit <= 9),
        "S2-F01 Decimal 数字无效，已拒绝生成"
    );
    let digits = value
        .digits
        .iter()
        .map(|digit| char::from(b'0' + *digit))
        .collect::<String>();
    let adjusted = value
        .exponent
        .checked_add(digits.len() as i64 - 1)
        .context("S2-F01 Decimal 指数超出可处理范围")?;
    let sign = if value.sign == 1 { "-" } else { "" };

    if value.exponent > 0 || adjusted < -6 {
        let coefficient = if digits.len() == 1 {
            digits
        } else {
            format!("{}.{}", &digits[..1], &digits[1..])
        };
        let exponent = if adjusted >= 0 {
            format!("+{adjusted}")
        } else {
            adjusted.to_string()
        };
        return Ok(format!("{sign}{coefficient}E{exponent}"));
    }

    let decimal_point = digits.len() as i64 + value.exponent;
    let fixed = if decimal_point <= 0 {
        format!("0.{}{}", "0".repeat((-decimal_point) as usize), digits)
    } else if decimal_point as usize >= digits.len() {
        format!(
            "{}{}",
            digits,
            "0".repeat(decimal_point as usize - digits.len())
        )
    } else {
        format!(
            "{}.{}",
            &digits[..decimal_point as usize],
            &digits[decimal_point as usize..]
        )
    };
    Ok(format!("{sign}{fixed}"))
}

fn checked_text(value: &str) -> Result<String> {
    ensure!(
        !value.chars().any(is_excel_illegal_control),
        "S2-F01 导出值包含 Excel 不支持的控制字符，无法生成审核清单。"
    );
    Ok(value.chars().take(MAX_TEXT_CHARS).collect())
}

fn is_excel_illegal_control(ch: char) -> bool {
    matches!(ch as u32, 0x00..=0x08 | 0x0B..=0x0C | 0x0E..=0x1F)
}

fn python_repr(value: &str) -> String {
    format!("'{}'", value.replace('\\', "\\\\").replace('\'', "\\'"))
}

fn write_workbook(
    path: &Path,
    tables: &[Table],
    on_progress: &mut dyn FnMut(&str, usize, usize),
) -> Result<()> {
    let total_rows = tables.iter().map(|table| table.rows.len()).sum::<usize>();
    let header = Format::new()
        .set_bold()
        .set_font_color(Color::RGB(0xFF_FF_FF))
        .set_background_color(Color::RGB(0x1F_4E_78))
        .set_align(FormatAlign::VerticalCenter)
        .set_align(FormatAlign::Center)
        .set_text_wrap();
    let long_header = Format::new()
        .set_bold()
        .set_font_color(Color::RGB(0xFF_FF_FF))
        .set_background_color(Color::RGB(0x1F_4E_78))
        .set_align(FormatAlign::Top)
        .set_text_wrap();
    let long_text = Format::new().set_align(FormatAlign::Top).set_text_wrap();
    let percentage = Format::new().set_num_format("0.00%");
    let percentage_points = Format::new().set_num_format("0.00");

    let mut workbook = Workbook::new();
    let mut completed_rows = 0;
    for table in tables {
        on_progress(table.name, completed_rows, total_rows);
        let worksheet = workbook.add_worksheet();
        worksheet
            .set_name(table.name)
            .with_context(|| format!("设置工作表名称失败：{}", table.name))?;
        worksheet.set_formula_result_default("");
        write_table(
            worksheet,
            table,
            &header,
            &long_header,
            &long_text,
            &percentage,
            &percentage_points,
            completed_rows,
            total_rows,
            &mut *on_progress,
        )?;
        completed_rows += table.rows.len();
    }

    on_progress("save", completed_rows, total_rows);
    workbook.save(path).context("保存 S2 XLSX 报告失败")?;
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn write_table(
    worksheet: &mut Worksheet,
    table: &Table,
    header: &Format,
    long_header: &Format,
    long_text: &Format,
    percentage: &Format,
    percentage_points: &Format,
    completed_rows: usize,
    total_rows: usize,
    on_progress: &mut dyn FnMut(&str, usize, usize),
) -> Result<()> {
    let long_columns = table
        .headers
        .iter()
        .enumerate()
        .filter_map(|(index, name)| LONG_TEXT_HEADERS.contains(name).then_some(index))
        .collect::<std::collections::HashSet<_>>();
    let description_column = if table.hidden_description.is_some() {
        table
            .headers
            .iter()
            .position(|header| *header == "审核说明")
    } else {
        None
    };
    let percentage_column = table
        .headers
        .iter()
        .position(|header| *header == "变动幅度");

    for (column, title) in table.headers.iter().enumerate() {
        let cell = text_cell(title, false)?;
        write_cell(
            worksheet,
            0,
            column as u16,
            &cell,
            Some(if long_columns.contains(&column) {
                long_header
            } else {
                header
            }),
        )?;
    }

    for (row_index, row) in table.rows.iter().enumerate() {
        for (column, cell) in row.iter().enumerate() {
            let format = if long_columns.contains(&column) {
                Some(long_text)
            } else if Some(column) == percentage_column {
                if table
                    .percentage_point_rows
                    .get(row_index)
                    .copied()
                    .unwrap_or(false)
                {
                    Some(percentage_points)
                } else {
                    Some(percentage)
                }
            } else {
                None
            };
            write_cell(
                worksheet,
                (row_index + 1) as u32,
                column as u16,
                cell,
                format,
            )?;
        }

        if let (Some(description_column), Some(expected)) =
            (description_column, table.hidden_description)
        {
            if row
                .get(description_column)
                .is_some_and(|cell| cell.is_hidden_description(expected))
            {
                worksheet
                    .set_row_hidden((row_index + 1) as u32)
                    .with_context(|| format!("隐藏 {} 的无变动行失败", table.name))?;
            }
        }

        let completed = row_index + 1;
        if completed % 250 == 0 {
            on_progress(table.name, completed_rows + completed, total_rows);
        }
    }

    on_progress(table.name, completed_rows + table.rows.len(), total_rows);

    for column in 0..table.headers.len() {
        let mut max_chars = table.headers[column].chars().count();
        for row in &table.rows {
            if let Some(cell) = row.get(column) {
                max_chars = max_chars.max(cell.width_text().chars().take(MAX_COLUMN_WIDTH).count());
            }
        }
        let width_chars = (max_chars + 2).clamp(12, MAX_COLUMN_WIDTH) as u32;
        // The legacy writer stores the requested character-width value directly
        // in OOXML. `set_column_width()` converts character width to pixels and
        // adds the default 5 px cell padding, which changes the serialized
        // width read back by openpyxl. Convert the legacy width to the pixel
        // value whose OOXML character width is exact (Calibri 11: 7 px/digit).
        let width_pixels = width_chars * 7;
        worksheet
            .set_column_width_pixels(column as u16, width_pixels)
            .with_context(|| format!("设置 {} 第 {} 列宽度失败", table.name, column + 1))?;
    }

    worksheet
        .set_freeze_panes(1, 0)
        .with_context(|| format!("冻结 {} 表头失败", table.name))?;
    worksheet
        .autofilter(
            0,
            0,
            table.rows.len() as u32,
            table.headers.len() as u16 - 1,
        )
        .with_context(|| format!("设置 {} 筛选区域失败", table.name))?;
    Ok(())
}

fn write_cell(
    worksheet: &mut Worksheet,
    row: u32,
    column: u16,
    cell: &ExcelCell,
    format: Option<&Format>,
) -> Result<()> {
    match (cell, format) {
        (ExcelCell::Blank, Some(format)) => {
            worksheet.write_string_with_format(row, column, "", format)?
        }
        (ExcelCell::Blank, None) => worksheet.write_string(row, column, "")?,
        (ExcelCell::Text(value), Some(format)) => {
            worksheet.write_string_with_format(row, column, value, format)?
        }
        (ExcelCell::Text(value), None) => worksheet.write_string(row, column, value)?,
        (ExcelCell::Number { value, .. }, Some(format)) => {
            worksheet.write_number_with_format(row, column, *value, format)?
        }
        (ExcelCell::Number { value, .. }, None) => worksheet.write_number(row, column, *value)?,
        (ExcelCell::Boolean(value), Some(format)) => {
            worksheet.write_boolean_with_format(row, column, *value, format)?
        }
        (ExcelCell::Boolean(value), None) => worksheet.write_boolean(row, column, *value)?,
        (ExcelCell::Formula(value), Some(format)) => {
            worksheet.write_formula_with_format(row, column, value.as_str(), format)?
        }
        (ExcelCell::Formula(value), None) => {
            worksheet.write_formula(row, column, value.as_str())?
        }
    };
    Ok(())
}

fn validate_workbook(path: &Path, expected_formulas: usize) -> Result<()> {
    let workbook = read_workbook(path).context("重新打开并校验 S2 XLSX 报告失败")?;
    let names = workbook
        .sheets
        .iter()
        .map(|sheet| sheet.name.as_str())
        .collect::<Vec<_>>();
    ensure!(names == SHEET_ORDER, "S2 XLSX 报告工作表名称或顺序不匹配");
    let formula_count = workbook
        .sheets
        .iter()
        .map(|sheet| sheet.formulas.len())
        .sum::<usize>();
    ensure!(
        formula_count == expected_formulas,
        "S2 XLSX 报告公式数不匹配：预期 {expected_formulas}，实际 {formula_count}"
    );
    Ok(())
}

struct TemporaryWorkbook {
    path: PathBuf,
    published: bool,
}

impl TemporaryWorkbook {
    fn create(directory: &Path, result_path: &Path) -> Result<Self> {
        let filename = result_path
            .file_name()
            .context("S2 输出文件名缺失")?
            .to_string_lossy();
        for _ in 0..100 {
            let counter = TEMP_COUNTER.fetch_add(1, Ordering::Relaxed);
            let path = directory.join(format!(
                ".{filename}.{}.{}.tmp.xlsx",
                std::process::id(),
                counter
            ));
            match OpenOptions::new().write(true).create_new(true).open(&path) {
                Ok(file) => {
                    drop(file);
                    return Ok(Self {
                        path,
                        published: false,
                    });
                }
                Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
                Err(error) => {
                    return Err(error)
                        .with_context(|| format!("创建临时 S2 报告失败：{}", path.display()));
                }
            }
        }
        bail!("无法为 S2 报告分配唯一临时文件名")
    }

    fn publish(&mut self, result_path: &Path) -> Result<()> {
        fs::rename(&self.path, result_path).with_context(|| {
            format!(
                "发布 S2 报告失败：{} -> {}",
                self.path.display(),
                result_path.display()
            )
        })?;
        self.published = true;
        Ok(())
    }
}

impl Drop for TemporaryWorkbook {
    fn drop(&mut self) {
        if !self.published {
            let _ = fs::remove_file(&self.path);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::{
        build_tables, checked_text, decimal_python_text, exact_integer_cell, value_cell, ExcelCell,
    };
    use crate::{s2_validation::S2AuditFinding, unit_conversion::DecimalTuple};
    use serde_json::json;

    #[test]
    fn decimal_text_preserves_python_decimal_scale_and_sign() {
        assert_eq!(
            decimal_python_text(&DecimalTuple {
                sign: 0,
                digits: vec![1, 2, 0, 0],
                exponent: -2,
            })
            .unwrap(),
            "12.00"
        );
        assert_eq!(
            decimal_python_text(&DecimalTuple {
                sign: 1,
                digits: vec![1],
                exponent: -7,
            })
            .unwrap(),
            "-1E-7"
        );
        assert_eq!(
            decimal_python_text(&DecimalTuple {
                sign: 0,
                digits: vec![1, 2, 3],
                exponent: 2,
            })
            .unwrap(),
            "1.23E+4"
        );
    }

    #[test]
    fn writer_values_keep_zero_boolean_formula_and_literal_label_distinct() {
        assert_eq!(
            value_cell(&json!(0), false).unwrap(),
            ExcelCell::Number {
                value: 0.0,
                display: "0".to_owned(),
            }
        );
        assert_eq!(
            value_cell(&json!(false), false).unwrap(),
            ExcelCell::Boolean(false)
        );
        assert_eq!(
            value_cell(&json!("=1+1"), false).unwrap(),
            ExcelCell::Formula("=1+1".to_owned())
        );
        assert_eq!(
            value_cell(&json!("=SUM(1,2)"), true).unwrap(),
            ExcelCell::Text("=SUM(1,2)".to_owned())
        );
        assert!(matches!(
            value_cell(
                &json!({"type":"decimal","value":{"sign":0,"digits":[1,2,0],"exponent":-1}}),
                false
            )
            .unwrap(),
            ExcelCell::Number { value: 12.0, .. }
        ));
    }

    #[test]
    fn illegal_controls_fail_and_long_text_is_truncated_by_scalar_count() {
        assert!(checked_text("before\u{000B}after").is_err());
        let value = "字".repeat(32_770);
        assert_eq!(checked_text(&value).unwrap().chars().count(), 32_767);
    }

    #[test]
    fn export_groups_findings_and_rejects_unrecognized_categories() {
        let mut item = S2AuditFinding::default();
        item.finding_id = "p".to_owned();
        item.audit_type = "环比".to_owned();
        item.status = "正常".to_owned();
        let tables = build_tables(&[item.clone()], false, false).unwrap();
        assert_eq!(tables[0].rows.len(), 1);
        assert!(tables[1].rows.is_empty() && tables[2].rows.is_empty());

        item.audit_type = "未知".to_owned();
        assert!(build_tables(&[item], false, false).is_err());
    }

    #[test]
    fn unsafe_integer_is_rejected_before_xlsx_precision_is_lost() {
        assert!(exact_integer_cell(9_007_199_254_740_993).is_err());
    }
}
