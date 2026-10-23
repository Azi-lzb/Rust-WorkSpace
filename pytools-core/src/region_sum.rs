//! Candidate Rust implementation of PyCode's read-only regional total check.
//!
//! The source feature reads stored worksheet values and creates a new summary
//! workbook. It does not modify the input workbooks. The Rust implementation
//! is available through the `pytools region-sum-check` CLI.

use anyhow::{bail, ensure, Context, Result};
use chrono::Local;
use excel_reader::{read_value_columns, CellValue, SheetValues};
use excel_writer::{Color, Format, FormatAlign, Workbook};
use std::collections::BTreeMap;
use std::fs::{self, File, OpenOptions};
use std::io::{BufWriter, Write};
use std::path::{Path, PathBuf};

const REGION_SHEET_KEY: &str = "分地区";
const TOTAL_REGION: &str = "惠州市";
const SUB_REGIONS: [&str; 6] = [
    "惠阳区",
    "博罗县",
    "惠东县",
    "龙门县",
    "惠州市本部",
    "大亚湾",
];
const ABS_TOLERANCE: f64 = 0.01;
const REL_TOLERANCE: f64 = 0.001;
const RESULT_SHEET: &str = "核对结果";
const RESULT_HEADERS: [&str; 6] = [
    "序号",
    "文件名",
    "工作表名",
    "核对状态",
    "错误数量",
    "错误信息",
];
const MAX_XLSX_COLUMNS: usize = 16_384;

#[derive(Clone, Debug, PartialEq, Eq)]
struct SheetCheck {
    error_count: usize,
    detail: String,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct ResultRow {
    sequence: usize,
    file_name: String,
    sheet_name: String,
    status: String,
    error_count: usize,
    detail: String,
}

/// Aggregate returned by `run_region_sum_check`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct RegionSumRun {
    pub files: usize,
    pub sheets: usize,
    pub errors: usize,
    pub failed_files: usize,
    pub path: Option<PathBuf>,
    pub log_path: PathBuf,
}

/// Run the standard feature using the current local time for output names.
pub fn run_region_sum_check<I, P>(
    target_workbooks: I,
    output_dir: impl AsRef<Path>,
    log_dir: impl AsRef<Path>,
) -> Result<RegionSumRun>
where
    I: IntoIterator<Item = P>,
    P: AsRef<Path>,
{
    let now = Local::now();
    let file_stamp = now.format("%Y%m%d_%H%M%S").to_string();
    let log_stamp = now.format("%Y-%m-%d %H:%M:%S").to_string();
    run_region_sum_check_with_timestamp(
        target_workbooks,
        output_dir,
        log_dir,
        &file_stamp,
        &log_stamp,
    )
}

/// Deterministic entry used by probes and tests; the timestamp must be
/// `YYYYMMDD_HHMMSS`. Product callers should use `run_region_sum_check`.
pub fn run_region_sum_check_with_timestamp<I, P>(
    target_workbooks: I,
    output_dir: impl AsRef<Path>,
    log_dir: impl AsRef<Path>,
    file_stamp: &str,
    log_stamp: &str,
) -> Result<RegionSumRun>
where
    I: IntoIterator<Item = P>,
    P: AsRef<Path>,
{
    validate_timestamp(file_stamp)?;
    let output_dir = output_dir.as_ref();
    let log_dir = log_dir.as_ref();
    fs::create_dir_all(log_dir)
        .with_context(|| format!("create log directory {}", log_dir.display()))?;
    let (log_path, initial_output_path) = unique_paths(output_dir, log_dir, file_stamp)?;
    let log_file =
        File::create(&log_path).with_context(|| format!("create log {}", log_path.display()))?;
    let mut log = BufWriter::new(log_file);

    let mut result_rows = Vec::new();
    let mut files = 0usize;
    let mut sheets_checked = 0usize;
    let mut errors = 0usize;
    let mut failed_files = 0usize;

    for input in target_workbooks {
        let source = input.as_ref();
        if !openpyxl_compatible_extension(source) {
            failed_files += 1;
            write_log(
                &mut log,
                log_stamp,
                "ERROR",
                &format!(
                    "打开失败: {} - unsupported workbook extension",
                    source.display()
                ),
            );
            continue;
        }

        let sheets = match read_value_columns(source, MAX_XLSX_COLUMNS) {
            Ok(sheets) => sheets,
            Err(error) => {
                failed_files += 1;
                write_log(
                    &mut log,
                    log_stamp,
                    "ERROR",
                    &format!("打开失败: {} - {error:#}", source.display()),
                );
                continue;
            }
        };

        files += 1;
        let mut file_sheets = 0usize;
        let mut file_errors = 0usize;
        let file_name = source
            .file_name()
            .map(|name| name.to_string_lossy().into_owned())
            .unwrap_or_default();

        for sheet in &sheets {
            if !sheet.name.contains(REGION_SHEET_KEY) {
                continue;
            }
            let check = check_one_sheet(sheet);
            sheets_checked += 1;
            errors += check.error_count;
            file_sheets += 1;
            file_errors += check.error_count;
            result_rows.push(make_result_row(
                result_rows.len() + 1,
                &file_name,
                &sheet.name,
                check,
            ));
        }
        write_log(
            &mut log,
            log_stamp,
            "INFO",
            &format!("{file_name}: 校验 {file_sheets} 个分地区 sheet，错误 {file_errors}"),
        );
    }

    fs::create_dir_all(output_dir)
        .with_context(|| format!("create output directory {}", output_dir.display()))?;
    let output_path = save_report(
        output_dir,
        initial_output_path,
        &result_rows,
        files,
        sheets_checked,
        errors,
    )?;
    write_log(
        &mut log,
        log_stamp,
        "INFO",
        &format!("结果已保存: {}", output_path.display()),
    );
    log.flush().context("flush feature log")?;

    Ok(RegionSumRun {
        files,
        sheets: sheets_checked,
        errors,
        failed_files,
        path: Some(output_path),
        log_path,
    })
}

fn validate_timestamp(value: &str) -> Result<()> {
    ensure!(
        value.len() == 15
            && value.as_bytes().get(8) == Some(&b'_')
            && value
                .bytes()
                .enumerate()
                .all(|(index, byte)| index == 8 || byte.is_ascii_digit()),
        "timestamp must use YYYYMMDD_HHMMSS"
    );
    Ok(())
}

fn openpyxl_compatible_extension(path: &Path) -> bool {
    path.extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| {
            matches!(
                extension.to_ascii_lowercase().as_str(),
                "xlsx" | "xlsm" | "xltx" | "xltm"
            )
        })
}

fn unique_paths(output_dir: &Path, log_dir: &Path, stamp: &str) -> Result<(PathBuf, PathBuf)> {
    for suffix in 0u32..10_000 {
        let suffix = if suffix == 0 {
            String::new()
        } else {
            format!("_{suffix}")
        };
        let output = output_dir.join(format!("地区数值核对结果_{stamp}{suffix}.xlsx"));
        let log = log_dir.join(format!("feature_1_5_region_sum_check_{stamp}{suffix}.log"));
        if !output.exists() && !log.exists() {
            return Ok((log, output));
        }
    }
    bail!("could not allocate a unique regional-check output name")
}

fn write_log(log: &mut BufWriter<File>, timestamp: &str, level: &str, message: &str) {
    let _ = writeln!(log, "{timestamp} [{level}] {message}");
}

fn make_result_row(
    sequence: usize,
    file_name: &str,
    sheet_name: &str,
    check: SheetCheck,
) -> ResultRow {
    let status = if check.error_count > 0 {
        "错误"
    } else {
        "正确"
    }
    .to_owned();
    let detail = if check.error_count > 0 {
        check.detail
    } else {
        "核对通过".to_owned()
    };
    ResultRow {
        sequence,
        file_name: file_name.to_owned(),
        sheet_name: sheet_name.to_owned(),
        status,
        error_count: check.error_count,
        detail,
    }
}

fn check_one_sheet(sheet: &SheetValues) -> SheetCheck {
    if sheet.max_row == 0 || sheet.max_column == 0 {
        return SheetCheck {
            error_count: 1,
            detail: "工作表为空或没有数据".to_owned(),
        };
    }

    let mut rows = BTreeMap::<String, usize>::new();
    for row in 1..=sheet.max_row {
        let Some(value) = cell_at(sheet, row, 2) else {
            continue;
        };
        let Some(name) = python_text(value) else {
            continue;
        };
        let name = name.trim();
        if (name == TOTAL_REGION || SUB_REGIONS.contains(&name)) && !rows.contains_key(name) {
            rows.insert(name.to_owned(), row);
        }
    }

    let Some(&total_row) = rows.get(TOTAL_REGION) else {
        return SheetCheck {
            error_count: 1,
            detail: format!("未找到'{TOTAL_REGION}'行"),
        };
    };
    if sheet.max_column < 3 {
        return SheetCheck {
            error_count: 1,
            detail: "未找到有效数据区域".to_owned(),
        };
    }

    let mut error_details = Vec::new();
    let mut numeric_error_count = 0usize;
    for column in 3..=sheet.max_column {
        let Some(total_value) = number_or_none(cell_at(sheet, total_row, column)) else {
            continue;
        };

        let mut sub_sum = 0.0;
        let mut any_sub_present = false;
        for name in SUB_REGIONS {
            let Some(&row) = rows.get(name) else {
                continue;
            };
            if let Some(value) = number_or_none(cell_at(sheet, row, column)) {
                sub_sum += value;
                any_sub_present = true;
            }
        }
        if !any_sub_present {
            continue;
        }

        let difference = total_value - sub_sum;
        let significant = if difference.abs() > ABS_TOLERANCE {
            true
        } else if total_value != 0.0 {
            (difference / total_value).abs() > REL_TOLERANCE
        } else {
            sub_sum != 0.0
        };
        if significant {
            numeric_error_count += 1;
            let header = cell_at(sheet, 2, column)
                .and_then(python_text)
                .map(|value| value.trim().to_owned())
                .filter(|value| !value.is_empty())
                .unwrap_or_else(|| format!("列{}", column_letters(column)));
            error_details.push(format!(
                "{header}: {} ≠ {} (差{})",
                fixed_2(total_value),
                fixed_2(sub_sum),
                fixed_2(difference)
            ));
        }
    }

    let missing = SUB_REGIONS
        .iter()
        .filter(|name| !rows.contains_key(**name))
        .copied()
        .collect::<Vec<_>>();
    if !missing.is_empty() {
        error_details.push(format!("未找到: {}", missing.join(", ")));
    }
    SheetCheck {
        error_count: numeric_error_count,
        detail: error_details.join("; "),
    }
}

fn cell_at(sheet: &SheetValues, row: usize, column: usize) -> Option<&CellValue> {
    sheet.values.get(&cell_address(row, column))
}

fn cell_address(row: usize, column: usize) -> String {
    format!("{}{}", column_letters(column), row)
}

fn column_letters(mut column: usize) -> String {
    let mut letters = String::new();
    while column > 0 {
        column -= 1;
        letters.insert(0, char::from(b'A' + (column % 26) as u8));
        column /= 26;
    }
    letters
}

fn python_text(value: &CellValue) -> Option<String> {
    match value {
        CellValue::Empty => None,
        CellValue::Text(value)
        | CellValue::DateTimeIso(value)
        | CellValue::DurationIso(value)
        | CellValue::Error(value) => Some(value.clone()),
        CellValue::Integer(value) => Some(value.to_string()),
        CellValue::Number(value) => {
            let mut text = value.to_string();
            if value.is_finite()
                && value.fract() == 0.0
                && !text.contains('e')
                && !text.contains('E')
            {
                text.push_str(".0");
            }
            Some(text)
        }
        CellValue::Boolean(value) => Some(if *value { "True" } else { "False" }.to_owned()),
        CellValue::DateTimeSerial(value) => Some(value.to_string()),
    }
}

fn number_or_none(value: Option<&CellValue>) -> Option<f64> {
    match value {
        None | Some(CellValue::Empty) => Some(0.0),
        Some(CellValue::Integer(value)) => Some(*value as f64),
        Some(CellValue::Number(value)) => Some(*value),
        Some(CellValue::Boolean(value)) => Some(if *value { 1.0 } else { 0.0 }),
        Some(CellValue::Text(value)) => parse_python_float(value),
        Some(
            CellValue::DateTimeSerial(_)
            | CellValue::DateTimeIso(_)
            | CellValue::DurationIso(_)
            | CellValue::Error(_),
        ) => None,
    }
}

fn parse_python_float(value: &str) -> Option<f64> {
    let value = value.trim();
    if value.is_empty() || value == "-" {
        return Some(0.0);
    }
    value.parse::<f64>().ok()
}

fn fixed_2(value: f64) -> String {
    if value.is_nan() {
        "nan".to_owned()
    } else if value == f64::INFINITY {
        "inf".to_owned()
    } else if value == f64::NEG_INFINITY {
        "-inf".to_owned()
    } else {
        format!("{value:.2}")
    }
}

fn save_report(
    output_dir: &Path,
    output_path: PathBuf,
    rows: &[ResultRow],
    files: usize,
    sheets: usize,
    errors: usize,
) -> Result<PathBuf> {
    let mut report_rows = Vec::<Vec<String>>::new();
    report_rows.push(
        RESULT_HEADERS
            .iter()
            .map(|value| (*value).to_owned())
            .collect(),
    );
    for row in rows {
        report_rows.push(vec![
            row.sequence.to_string(),
            row.file_name.clone(),
            row.sheet_name.clone(),
            row.status.clone(),
            row.error_count.to_string(),
            row.detail.clone(),
        ]);
    }
    if !rows.is_empty() {
        report_rows.push(vec![
            "=== 统计汇总 ===".to_owned(),
            format!("已检查文件数: {files}"),
            format!("已检查工作表数: {sheets}"),
            format!("发现错误总数: {errors}"),
            String::new(),
            String::new(),
        ]);
    }

    let header_format = Format::new()
        .set_bold()
        .set_align(FormatAlign::Center)
        .set_background_color(Color::RGB(0xC8DCF0));
    let error_format = Format::new()
        .set_font_color(Color::RGB(0xC00000))
        .set_background_color(Color::RGB(0xFFC8C8));
    let ok_format = Format::new()
        .set_font_color(Color::RGB(0x006400))
        .set_background_color(Color::RGB(0xC8FFC8));

    let mut workbook = Workbook::new();
    {
        let worksheet = workbook.add_worksheet();
        worksheet.set_name(RESULT_SHEET)?;
        // The Python baseline stores this marker as a formula with no cached
        // value, so preserve that cache state instead of rust_xlsxwriter's 0.
        worksheet.set_formula_result_default("");
        for (column, header) in RESULT_HEADERS.iter().enumerate() {
            worksheet.write_string_with_format(0, column as u16, *header, &header_format)?;
        }
        for (index, row) in rows.iter().enumerate() {
            let excel_row = index as u32 + 1;
            let format = if row.error_count > 0 {
                &error_format
            } else {
                &ok_format
            };
            worksheet.write_number_with_format(excel_row, 0, row.sequence as f64, format)?;
            worksheet.write_string_with_format(excel_row, 1, &row.file_name, format)?;
            worksheet.write_string_with_format(excel_row, 2, &row.sheet_name, format)?;
            worksheet.write_string_with_format(excel_row, 3, &row.status, format)?;
            worksheet.write_number_with_format(excel_row, 4, row.error_count as f64, format)?;
            worksheet.write_string_with_format(excel_row, 5, &row.detail, format)?;
        }
        if !rows.is_empty() {
            let row = rows.len() as u32 + 1;
            worksheet.write_formula(row, 0, "=== 统计汇总 ===")?;
            worksheet.write_string(row, 1, format!("已检查文件数: {files}"))?;
            worksheet.write_string(row, 2, format!("已检查工作表数: {sheets}"))?;
            worksheet.write_string(row, 3, format!("发现错误总数: {errors}"))?;
        }
        for column in 0..RESULT_HEADERS.len() {
            let max_length = report_rows
                .iter()
                .filter_map(|row| row.get(column))
                .map(|value| display_width(value))
                .max()
                .unwrap_or_default();
            let width_units = (max_length + 2).clamp(8, 80) as u32;
            // The source writes openpyxl's raw width value. rust_xlsxwriter's
            // character-unit API adds 5 px of padding, so use its pixel API to
            // reproduce the same OOXML width attribute for the default 7 px digit.
            worksheet.set_column_width_pixels(column as u16, width_units * 7)?;
        }
    }

    let temporary = create_temporary_path(output_dir)?;
    let save_result = workbook
        .save(&temporary)
        .with_context(|| format!("write temporary report {}", temporary.display()));
    if let Err(error) = save_result {
        let _ = fs::remove_file(&temporary);
        return Err(error);
    }

    let committed_source = if rows.is_empty() {
        temporary.clone()
    } else {
        let patched = match create_temporary_path(output_dir) {
            Ok(path) => path,
            Err(error) => {
                let _ = fs::remove_file(&temporary);
                return Err(error);
            }
        };
        let formula_marker_cell = format!("A{}", rows.len() + 2);
        if let Err(error) = ooxml_patch::clear_empty_formula_cache(
            &temporary,
            &patched,
            RESULT_SHEET,
            &formula_marker_cell,
        ) {
            let _ = fs::remove_file(&temporary);
            let _ = fs::remove_file(&patched);
            return Err(error).context("preserve empty formula cache in result marker");
        }
        fs::remove_file(&temporary).with_context(|| {
            format!("remove unpatched temporary report {}", temporary.display())
        })?;
        patched
    };

    let committed_path = commit_without_overwrite(&committed_source, output_path)?;
    Ok(committed_path)
}

fn display_width(value: &str) -> usize {
    value
        .chars()
        .map(|character| if character as u32 > 127 { 2 } else { 1 })
        .sum()
}

fn create_temporary_path(output_dir: &Path) -> Result<PathBuf> {
    for attempt in 0u32..10_000 {
        let path = output_dir.join(format!(
            ".region_sum_{}_{}.tmp.xlsx",
            std::process::id(),
            attempt
        ));
        match OpenOptions::new().write(true).create_new(true).open(&path) {
            Ok(file) => {
                drop(file);
                return Ok(path);
            }
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(error) => return Err(error).with_context(|| format!("reserve {}", path.display())),
        }
    }
    bail!("could not allocate temporary report path")
}

fn commit_without_overwrite(temporary: &Path, desired: PathBuf) -> Result<PathBuf> {
    for suffix in 0u32..10_000 {
        let candidate = if suffix == 0 {
            desired.clone()
        } else {
            let stem = desired
                .file_stem()
                .and_then(|name| name.to_str())
                .unwrap_or("report");
            let extension = desired
                .extension()
                .and_then(|name| name.to_str())
                .unwrap_or("xlsx");
            desired.with_file_name(format!("{stem}_{suffix}.{extension}"))
        };
        match fs::hard_link(temporary, &candidate) {
            Ok(()) => {
                let _ = fs::remove_file(temporary);
                return Ok(candidate);
            }
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(error) => {
                let _ = fs::remove_file(temporary);
                return Err(error).with_context(|| {
                    format!("commit report without overwriting {}", candidate.display())
                });
            }
        }
    }
    let _ = fs::remove_file(temporary);
    bail!("could not allocate a unique report output path")
}

#[cfg(test)]
mod tests {
    use super::{check_one_sheet, number_or_none, parse_python_float};
    use excel_reader::{CellValue, SheetValues};
    use std::collections::BTreeMap;

    fn sheet(values: &[(&str, CellValue)], max_row: usize, max_column: usize) -> SheetValues {
        SheetValues {
            name: "分地区测试".to_owned(),
            values: values
                .iter()
                .map(|(address, value)| ((*address).to_owned(), value.clone()))
                .collect::<BTreeMap<_, _>>(),
            max_row,
            max_column,
        }
    }

    #[test]
    fn numeric_conversion_preserves_source_blank_boolean_and_date_rules() {
        assert_eq!(number_or_none(None), Some(0.0));
        assert_eq!(number_or_none(Some(&CellValue::Boolean(true))), Some(1.0));
        assert_eq!(
            number_or_none(Some(&CellValue::Text(" - ".to_owned()))),
            Some(0.0)
        );
        assert_eq!(
            number_or_none(Some(&CellValue::Text("2.5".to_owned()))),
            Some(2.5)
        );
        assert_eq!(
            number_or_none(Some(&CellValue::Text("bad".to_owned()))),
            None
        );
        assert_eq!(
            number_or_none(Some(&CellValue::DateTimeSerial(45_000.0))),
            None
        );
        assert_eq!(parse_python_float(" Infinity "), Some(f64::INFINITY));
    }

    #[test]
    fn sheet_check_counts_numeric_mismatches_but_preserves_missing_region_text() {
        let input = sheet(
            &[
                ("B2", CellValue::Text("地区".to_owned())),
                ("C2", CellValue::Text("金额".to_owned())),
                ("B3", CellValue::Text("惠州市".to_owned())),
                ("C3", CellValue::Integer(10)),
                ("B4", CellValue::Text("惠阳区".to_owned())),
                ("C4", CellValue::Integer(3)),
            ],
            4,
            3,
        );
        let result = check_one_sheet(&input);
        assert_eq!(result.error_count, 1);
        assert!(result.detail.starts_with("金额: 10.00 ≠ 3.00 (差7.00)"));
        assert!(result
            .detail
            .contains("未找到: 博罗县, 惠东县, 龙门县, 惠州市本部, 大亚湾"));
    }
}
