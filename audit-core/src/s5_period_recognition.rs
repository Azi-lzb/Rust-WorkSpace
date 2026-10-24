//! S5 configured-statistics period recognition.
//!
//! This keeps the S5 date contract separate from S1's `YYYY-MM` discovery
//! helper: S5 recognizes exact days and maps month-only tokens to month-end.

use anyhow::{bail, Context, Result};
use chrono::{Duration, NaiveDate};
use excel_reader::{
    read_value_columns_with_iso_dates, read_workbook, CellValue, SheetDump, WorkbookDump,
};
use serde::Serialize;
use std::{
    collections::{BTreeMap, BTreeSet},
    path::{Path, PathBuf},
};

const PERIOD_LABELS: &[&str] = &[
    "数据日期",
    "报告日期",
    "报告期",
    "统计日期",
    "数据期",
    "数据截止日期",
    "报表日期",
];

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct S5PeriodDetection {
    pub period: String,
    pub source: String,
    pub selected_file_count: usize,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub reader: Option<String>,
}

/// Detect the unique S5 report period from selected filenames, workbook fields,
/// and then the input directory's ancestor names, in that order.
pub fn detect_s5_period(files: &[PathBuf], input_dir: &Path) -> Result<S5PeriodDetection> {
    let mut periods = BTreeSet::new();
    for path in files {
        let stem = path
            .file_stem()
            .and_then(|value| value.to_str())
            .unwrap_or_default();
        periods.extend(
            analyze_period_text(stem)
                .with_context(|| format!("解析源文件名报告期：{}", path.display()))?
                .periods,
        );
    }

    let mut source = "源文件名".to_owned();
    let mut reader = None;
    if periods.is_empty() {
        source = "源表报告期字段".to_owned();
        reader = Some("calamine".to_owned());
        for path in files {
            periods.extend(
                periods_in_workbook_fields(path)
                    .with_context(|| format!("读取源表报告期字段：{}", path.display()))?,
            );
        }
    }

    if periods.is_empty() {
        let mut folder = Some(input_dir);
        while let Some(path) = folder {
            let name = path
                .file_name()
                .and_then(|value| value.to_str())
                .unwrap_or_default();
            let analysis = analyze_period_text(name)
                .with_context(|| format!("解析所属文件夹报告期：{name}"))?;
            if !analysis.periods.is_empty() {
                periods = analysis.periods;
                source = format!("文件夹“{name}”");
                if !analysis.has_day_pattern {
                    source.push_str("（仅年月时按月末）");
                }
                break;
            }
            folder = path.parent();
        }
    }

    if periods.len() > 1 {
        bail!(
            "源文件存在多个报告期：{}；请按报告期分目录或调整勾选文件",
            periods.iter().cloned().collect::<Vec<_>>().join("、")
        );
    }
    let period = periods
        .into_iter()
        .next()
        .context("未识别到报告期；请在源文件或所属文件夹名称中标明 YYYY-MM-DD 或 YYYYMM")?;

    Ok(S5PeriodDetection {
        period,
        source,
        selected_file_count: files.len(),
        reader,
    })
}

/// Return the S5 file-list period hint using only the filename and parent
/// directory names. This intentionally does not open the workbook.
pub fn s5_file_period_hint(path: &Path) -> String {
    let stem = path
        .file_stem()
        .and_then(|value| value.to_str())
        .unwrap_or_default();
    match analyze_period_text(stem) {
        Ok(analysis) => {
            if let Some(hint) = period_hint(&analysis.periods) {
                return hint;
            }
        }
        Err(_) => return "日期格式异常".to_owned(),
    }

    let parent = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let resolved_parent = std::fs::canonicalize(parent).unwrap_or_else(|_| parent.to_path_buf());
    for folder in resolved_parent.ancestors() {
        let name = folder
            .file_name()
            .and_then(|value| value.to_str())
            .unwrap_or_default();
        match analyze_period_text(name) {
            Ok(analysis) => {
                if let Some(hint) = period_hint(&analysis.periods) {
                    return hint;
                }
            }
            Err(_) => return "日期格式异常".to_owned(),
        }
    }
    String::new()
}

fn period_hint(periods: &BTreeSet<String>) -> Option<String> {
    match periods.len() {
        0 => None,
        1 => periods.first().cloned(),
        _ => Some("多个日期".to_owned()),
    }
}

/// Parse S5's accepted day/month date tokens. Month-only values become the
/// last calendar day of that month; invalid day tokens are rejected.
pub fn periods_in_s5_text(value: &str) -> Result<BTreeSet<String>> {
    Ok(analyze_period_text(value)?.periods)
}

#[derive(Default)]
struct PeriodTextAnalysis {
    periods: BTreeSet<String>,
    has_day_pattern: bool,
}

fn analyze_period_text(value: &str) -> Result<PeriodTextAnalysis> {
    if value.starts_with('=') {
        return Ok(PeriodTextAnalysis::default());
    }

    let chars = value.chars().collect::<Vec<_>>();
    let mut remaining = chars.clone();
    let mut analysis = PeriodTextAnalysis::default();
    let mut position = 0;
    while position < chars.len() {
        let Some(token) = day_token_at(&chars, position) else {
            position += 1;
            continue;
        };
        let date = NaiveDate::from_ymd_opt(token.year, token.month, token.day)
            .ok_or_else(|| anyhow::anyhow!("报告期日期不正确：{}", token.text))?;
        analysis.periods.insert(date.format("%Y-%m-%d").to_string());
        analysis.has_day_pattern = true;
        remaining[token.start..token.end].fill(' ');
        position = token.end;
    }

    for position in 0..remaining.len() {
        let Some(token) = month_token_at(&remaining, position) else {
            continue;
        };
        let last_day = month_end(token.year, token.month)
            .ok_or_else(|| anyhow::anyhow!("报告期日期不正确：{}", token.text))?;
        analysis
            .periods
            .insert(last_day.format("%Y-%m-%d").to_string());
    }

    Ok(analysis)
}

struct DayToken {
    start: usize,
    end: usize,
    year: i32,
    month: u32,
    day: u32,
    text: String,
}

fn day_token_at(chars: &[char], start: usize) -> Option<DayToken> {
    let (year, after_year) = year_at(chars, start)?;

    if chars
        .get(after_year)
        .is_some_and(|value| is_day_separator(*value))
    {
        let month_start = after_year + 1;
        for month_len in [2, 1] {
            let Some((month, after_month)) = decimal_digits(chars, month_start, month_len) else {
                continue;
            };
            if !chars
                .get(after_month)
                .is_some_and(|value| is_day_separator(*value))
            {
                continue;
            }
            let day_start = after_month + 1;
            for day_len in [2, 1] {
                let Some((day, mut end)) = decimal_digits(chars, day_start, day_len) else {
                    continue;
                };
                if chars.get(end) == Some(&'日') {
                    end += 1;
                }
                if !is_digit_at(chars, end) {
                    return Some(DayToken {
                        start,
                        end,
                        year,
                        month,
                        day,
                        text: chars[start..end].iter().collect(),
                    });
                }
            }
        }
        return None;
    }

    let (month, after_month) = decimal_digits(chars, after_year, 2)?;
    let (day, mut end) = decimal_digits(chars, after_month, 2)?;
    if chars.get(end) == Some(&'日') {
        end += 1;
    }
    if is_digit_at(chars, end) {
        return None;
    }
    Some(DayToken {
        start,
        end,
        year,
        month,
        day,
        text: chars[start..end].iter().collect(),
    })
}

struct MonthToken {
    year: i32,
    month: u32,
    text: String,
}

fn month_token_at(chars: &[char], start: usize) -> Option<MonthToken> {
    let (year, after_year) = year_at(chars, start)?;
    if chars
        .get(after_year)
        .is_some_and(|value| is_month_separator(*value))
    {
        let month_start = after_year + 1;
        for month_len in [2, 1] {
            let Some((month, mut end)) = decimal_digits(chars, month_start, month_len) else {
                continue;
            };
            if chars.get(end) == Some(&'月') {
                end += 1;
            }
            if !is_digit_at(chars, end) {
                return Some(MonthToken {
                    year,
                    month,
                    text: chars[start..end].iter().collect(),
                });
            }
        }
        return None;
    }

    let (month, end) = decimal_digits(chars, after_year, 2)?;
    if is_digit_at(chars, end) {
        return None;
    }
    Some(MonthToken {
        year,
        month,
        text: chars[start..end].iter().collect(),
    })
}

fn year_at(chars: &[char], start: usize) -> Option<(i32, usize)> {
    if start > 0 && crate::discovery::decimal_digit(chars[start - 1]).is_some() {
        return None;
    }
    if chars.get(start..start + 2)? != ['2', '0'] {
        return None;
    }
    let (last_two, end) = decimal_digits(chars, start + 2, 2)?;
    Some((2000 + last_two as i32, end))
}

fn decimal_digits(chars: &[char], start: usize, count: usize) -> Option<(u32, usize)> {
    let end = start.checked_add(count)?;
    let slice = chars.get(start..end)?;
    let mut value = 0_u32;
    for character in slice {
        let digit = crate::discovery::decimal_digit(*character)?;
        value = value.checked_mul(10)?.checked_add(digit)?;
    }
    Some((value, end))
}

fn is_digit_at(chars: &[char], position: usize) -> bool {
    chars
        .get(position)
        .is_some_and(|value| crate::discovery::decimal_digit(*value).is_some())
}

fn is_day_separator(value: char) -> bool {
    matches!(value, '-' | '.' | '_' | '年' | '月' | '/')
}

fn is_month_separator(value: char) -> bool {
    matches!(value, '-' | '.' | '_' | '年' | '/')
}

fn month_end(year: i32, month: u32) -> Option<NaiveDate> {
    let (next_year, next_month) = if month == 12 {
        (year.checked_add(1)?, 1)
    } else {
        (year, month.checked_add(1)?)
    };
    NaiveDate::from_ymd_opt(next_year, next_month, 1)?.checked_sub_signed(Duration::days(1))
}

fn periods_in_workbook_fields(path: &Path) -> Result<BTreeSet<String>> {
    let workbook = read_workbook(path).context("使用 Calamine 读取工作簿值与公式")?;
    let (mut periods, needs_iso_dates) = scan_workbook_fields(&workbook, None)?;
    if !needs_iso_dates {
        return Ok(periods);
    }

    // The shared reader exposes date-formatted values as ISO strings while
    // preserving the workbook's 1900/1904 date system. Use that projection only
    // when a recognized period field contains a date serial.
    let mut iso_dates = BTreeMap::new();
    for sheet in read_value_columns_with_iso_dates(path, 17)
        .context("使用 Calamine 读取报告期字段的 ISO 日期值")?
    {
        for (coordinate, value) in sheet.values {
            if let CellValue::DateTimeIso(value) = value {
                iso_dates.insert((sheet.name.clone(), coordinate), value);
            }
        }
    }
    periods = scan_workbook_fields(&workbook, Some(&iso_dates))?.0;
    Ok(periods)
}

fn scan_workbook_fields(
    workbook: &WorkbookDump,
    iso_dates: Option<&BTreeMap<(String, String), String>>,
) -> Result<(BTreeSet<String>, bool)> {
    let mut periods = BTreeSet::new();
    let mut needs_iso_dates = false;

    for sheet in &workbook.sheets {
        for row in 1..=60 {
            for column in 1..=15 {
                let label_coordinate = cell_coordinate(row, column);
                let (label_text, _) = cell_text(sheet, &label_coordinate, iso_dates, false)?;
                let Some(label_text) = label_text else {
                    continue;
                };
                let label = label_text
                    .split(|character| matches!(character, ':' | '：'))
                    .next()
                    .unwrap_or_default()
                    .trim();
                if !PERIOD_LABELS.contains(&label) {
                    continue;
                }

                add_text_periods(&mut periods, &label_text)?;
                for (neighbor_row, neighbor_column) in
                    [(row, column + 1), (row, column + 2), (row + 1, column)]
                {
                    let coordinate = cell_coordinate(neighbor_row, neighbor_column);
                    let (text, needs_decode) = cell_text(sheet, &coordinate, iso_dates, true)?;
                    needs_iso_dates |= needs_decode;
                    if let Some(text) = text {
                        add_text_periods(&mut periods, &text)?;
                    }
                }
            }
        }
    }

    Ok((periods, needs_iso_dates))
}

fn add_text_periods(periods: &mut BTreeSet<String>, value: &str) -> Result<()> {
    periods.extend(analyze_period_text(value)?.periods);
    Ok(())
}

fn cell_text(
    sheet: &SheetDump,
    coordinate: &str,
    iso_dates: Option<&BTreeMap<(String, String), String>>,
    request_iso_for_serial: bool,
) -> Result<(Option<String>, bool)> {
    if sheet.formulas.contains_key(coordinate) {
        return Ok((None, false));
    }
    let Some(value) = sheet.values.get(coordinate) else {
        return Ok((None, false));
    };
    let value = CellValue::from_probe_json(value)?;
    let (text, needs_iso) = match value {
        CellValue::Empty => (None, false),
        CellValue::Text(value) => (Some(value), false),
        CellValue::Number(value) if value.is_finite() => (Some(value.to_string()), false),
        CellValue::Number(value) => (Some(value.to_string()), false),
        CellValue::Integer(value) => (Some(value.to_string()), false),
        CellValue::Boolean(true) => (Some("True".to_owned()), false),
        CellValue::Boolean(false) => (None, false),
        CellValue::DateTimeSerial(_) => {
            let key = (sheet.name.clone(), coordinate.to_owned());
            (
                iso_dates.and_then(|dates| dates.get(&key).cloned()),
                request_iso_for_serial,
            )
        }
        CellValue::DateTimeIso(value) | CellValue::DurationIso(value) => (Some(value), false),
        CellValue::Error(value) => (Some(value), false),
    };
    Ok((text, needs_iso))
}

fn cell_coordinate(row: usize, mut column: usize) -> String {
    let mut letters = Vec::new();
    while column > 0 {
        let remainder = (column - 1) % 26;
        letters.push((b'A' + remainder as u8) as char);
        column = (column - 1) / 26;
    }
    letters.reverse();
    format!("{}{row}", letters.into_iter().collect::<String>())
}

#[cfg(test)]
mod tests {
    use super::{
        analyze_period_text, detect_s5_period, periods_in_s5_text, s5_file_period_hint,
        scan_workbook_fields,
    };
    use excel_reader::{SheetDump, WorkbookDump};
    use serde::Deserialize;
    use serde_json::json;
    use std::{collections::BTreeMap, fs, path::PathBuf};

    #[derive(Deserialize)]
    struct PythonPeriodBaseline {
        cases: Vec<PythonPeriodCase>,
    }

    #[derive(Deserialize)]
    struct PythonPeriodCase {
        input: String,
        periods: Option<Vec<String>>,
        error: Option<String>,
    }

    #[derive(Deserialize)]
    struct PythonFileHintBaseline {
        cases: Vec<PythonFileHintCase>,
    }

    #[derive(Deserialize)]
    struct PythonFileHintCase {
        input: PathBuf,
        period_hint: String,
    }

    #[test]
    fn period_parser_matches_python_s5_baseline_vectors() {
        let baseline: PythonPeriodBaseline = serde_json::from_str(include_str!(
            "../probe/reports/s5_period_parser_python_baseline.json"
        ))
        .expect("valid Python S5 period parser baseline");
        assert!(baseline.cases.len() >= 12, "baseline vector count changed");
        for (index, case) in baseline.cases.iter().enumerate() {
            let actual = periods_in_s5_text(&case.input);
            if case.error.is_some() {
                assert!(
                    actual.is_err(),
                    "baseline error case {}: {:?}",
                    index + 1,
                    case.input
                );
            } else {
                assert_eq!(
                    actual.unwrap().into_iter().collect::<Vec<_>>(),
                    case.periods.clone().unwrap_or_default(),
                    "Python parity case {}: {:?}",
                    index + 1,
                    case.input
                );
            }
        }
    }

    #[test]
    fn file_period_hint_matches_python_filename_and_directory_baseline() {
        let baseline: PythonFileHintBaseline = serde_json::from_str(include_str!(
            "../probe/reports/s5_file_period_hint_python_baseline.json"
        ))
        .expect("valid Python S5 file period hint baseline");
        assert_eq!(baseline.cases.len(), 9);
        for (index, case) in baseline.cases.iter().enumerate() {
            assert_eq!(
                s5_file_period_hint(&case.input),
                case.period_hint,
                "Python file-period-hint case {}: {:?}",
                index + 1,
                case.input
            );
        }
    }

    #[test]
    fn parses_python_s5_period_formats_and_month_end() {
        let cases = [
            ("甲机构_20260831_核查表", "2026-08-31"),
            ("reports#2026-08-31#institution", "2026-08-31"),
            ("报告2026年7月31日", "2026-07-31"),
            ("2026/7/3", "2026-07-03"),
            ("202608", "2026-08-31"),
            ("2024年2月", "2024-02-29"),
            ("2026.8.3", "2026-08-03"),
            ("2026_08_03", "2026-08-03"),
        ];
        for (input, expected) in cases {
            assert_eq!(
                periods_in_s5_text(input)
                    .unwrap()
                    .into_iter()
                    .collect::<Vec<_>>(),
                [expected],
                "input {input:?}"
            );
        }
        assert!(periods_in_s5_text("=DATE(2026,8,31)").unwrap().is_empty());
        assert!(periods_in_s5_text("2026-02-30").is_err());
    }

    #[test]
    fn rejects_multiple_s5_periods_in_selected_filenames() {
        let files = [
            PathBuf::from("first_20260831.xlsx"),
            PathBuf::from("second_20260930.xlsx"),
        ];
        let error = detect_s5_period(&files, PathBuf::from("input").as_path()).unwrap_err();
        assert!(format!("{error:#}").contains("多个报告期"));
    }

    #[test]
    fn prefers_selected_filenames_without_opening_workbooks() {
        let files = [PathBuf::from("selected_20260831.xlsx")];
        let result = detect_s5_period(&files, PathBuf::from("input").as_path()).unwrap();
        assert_eq!(result.period, "2026-08-31");
        assert_eq!(result.source, "源文件名");
        assert_eq!(result.reader, None);
    }

    #[test]
    fn reads_s5_period_field_adjacent_to_label_and_ignores_formula_cells() {
        let mut values = BTreeMap::new();
        values.insert(
            "A1".to_owned(),
            json!({"type":"string", "value":"报告日期："}),
        );
        values.insert(
            "B1".to_owned(),
            json!({"type":"string", "value":"2026年8月31日"}),
        );
        values.insert(
            "A2".to_owned(),
            json!({"type":"string", "value":"数据日期："}),
        );
        values.insert(
            "B2".to_owned(),
            json!({"type":"string", "value":"2026-09-30"}),
        );
        let mut formulas = BTreeMap::new();
        formulas.insert("B2".to_owned(), "=DATE(2026,9,30)".to_owned());
        let workbook = WorkbookDump {
            schema_version: 1,
            reader: "calamine",
            elapsed_ms: 0,
            sheets: vec![SheetDump {
                name: "Summary".to_owned(),
                values,
                formulas,
                formula_status: "available".to_owned(),
                merges: Vec::new(),
                merge_status: "available".to_owned(),
                comments: BTreeMap::new(),
                comments_status: "available".to_owned(),
                conditional_formatting: Vec::new(),
                conditional_formatting_status: "available".to_owned(),
                data_validations: Vec::new(),
                data_validations_status: "available".to_owned(),
            }],
        };
        let (periods, needs_iso) = scan_workbook_fields(&workbook, None).unwrap();
        assert_eq!(periods.into_iter().collect::<Vec<_>>(), ["2026-08-31"]);
        assert!(!needs_iso);
    }

    #[test]
    fn detects_a_real_date_typed_cell_through_the_calamine_reader() {
        use excel_writer::{ExcelDateTime, Format, Workbook};

        let directory = std::env::temp_dir().join("audit-s5-period-recognition-test");
        fs::create_dir_all(&directory).unwrap();
        let path = directory.join("period-source.xlsx");
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();
        worksheet.write_string(0, 0, "报告日期：").unwrap();
        let date = ExcelDateTime::parse_from_str("2026-08-31").unwrap();
        let date_format = Format::new().set_num_format("yyyy-mm-dd");
        worksheet
            .write_datetime_with_format(0, 1, date, &date_format)
            .unwrap();
        workbook.save(&path).unwrap();

        let result = detect_s5_period(&[path], &directory).unwrap();
        assert_eq!(result.period, "2026-08-31");
        assert_eq!(result.source, "源表报告期字段");
        assert_eq!(result.reader.as_deref(), Some("calamine"));

        fs::remove_dir_all(directory).unwrap();
    }

    #[test]
    fn month_name_evidence_marks_month_end_and_exact_day_evidence_does_not() {
        let month = analyze_period_text("input_202608").unwrap();
        assert_eq!(
            month.periods.into_iter().collect::<Vec<_>>(),
            ["2026-08-31"]
        );
        assert!(!month.has_day_pattern);
        let day = analyze_period_text("input_2026-08-31").unwrap();
        assert_eq!(day.periods.into_iter().collect::<Vec<_>>(), ["2026-08-31"]);
        assert!(day.has_day_pattern);
    }
}
