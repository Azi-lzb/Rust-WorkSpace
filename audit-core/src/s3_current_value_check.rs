//! Business rules for S3-F04, the current/previous period numeric check.
//!
//! This module mirrors the Python `cross_period_engine.run_cross_period`
//! contract. Input loading, rule configuration, and workbook presentation stay
//! at the service/CLI boundary so the calculation can be compared independently.

use crate::central_csv::{
    header_matches, parse_central_csv, parse_central_grid, CentralDataset, CentralValue,
};
use anyhow::{bail, ensure, Context, Result};
use chrono::{Local, NaiveDate, NaiveDateTime};
use excel_reader::{read_value_columns, read_value_columns_with_iso_dates, CellValue, SheetValues};
use excel_writer::{Color, Format, Workbook};
use regex::Regex;
use serde::Serialize;
use std::{
    collections::{BTreeMap, BTreeSet, HashMap, HashSet},
    fs,
    path::{Path, PathBuf},
    time::Instant,
};

const EQUALITY_EPSILON: f64 = 1e-12;

#[derive(Clone, Debug, Default, PartialEq)]
pub struct S3CrossPeriodRule {
    pub disabled: String,
    pub rule_code: String,
    pub rule_name: String,
    pub institution_pattern: String,
    pub region_pattern: String,
    pub premise: String,
    pub left_expression: String,
    pub right_expression: String,
    pub formula: String,
    pub hint: String,
    pub invert: String,
    pub suppress_success: String,
    pub logical_check: String,
    pub remark: String,
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct S3CurrentValueCheckRow {
    pub data_date: String,
    pub institution_code: String,
    pub institution_name: String,
    pub region_code: String,
    pub region_name: String,
    pub rule_code: String,
    pub rule_name: String,
    pub check_type: String,
    pub remark: String,
    pub left_value: f64,
    pub right_value: f64,
    pub difference: f64,
    pub absolute_difference: f64,
    pub difference_percent: Option<f64>,
    pub explanation_required: String,
    pub explanation: String,
    pub calculation: String,
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct S3CurrentValueCheckLog {
    pub institution_code: String,
    pub region_code: String,
    pub error_type: String,
    pub rule: String,
    pub detail: String,
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct S3CurrentValueCheckResult {
    pub rows: Vec<S3CurrentValueCheckRow>,
    pub logs: Vec<S3CurrentValueCheckLog>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct S3CurrentValueCheckConfig {
    pub source_unit: String,
    pub target_unit: String,
    pub exempt_indicators: BTreeSet<String>,
    pub rules: Vec<S3CrossPeriodRule>,
}

#[derive(Clone, Debug, Serialize)]
pub struct S3CurrentValueCheckReport {
    pub status: &'static str,
    pub output_workbook: PathBuf,
    pub log_workbook: Option<PathBuf>,
    pub rows: usize,
    pub total_rules: usize,
    pub enabled_rules: usize,
    pub disabled_rules: usize,
    pub current_records: usize,
    pub previous_records: usize,
    pub config_reader: &'static str,
    pub input_reader: String,
    pub writer: &'static str,
    pub config_read_time_ms: u128,
    pub input_read_time_ms: u128,
    pub calculation_time_ms: u128,
    pub workbook_write_time_ms: u128,
    pub total_time_ms: u128,
}

/// Match institution and region scopes with Python `re.search` behavior.
/// Invalid patterns are cached as non-matches, matching `_safe_regex_match`.
pub fn python_regex_scope_matcher() -> impl FnMut(&str, &str) -> bool {
    let mut compiled: HashMap<String, Option<Regex>> = HashMap::new();
    move |pattern, text| {
        if pattern.is_empty() {
            return true;
        }
        compiled
            .entry(pattern.to_owned())
            .or_insert_with(|| Regex::new(pattern).ok())
            .as_ref()
            .is_some_and(|regex| regex.is_match(text))
    }
}

#[derive(Clone, Copy, Debug, Default)]
struct PairedValue {
    current: Option<f64>,
    previous: Option<f64>,
}

#[derive(Clone, Debug, Default)]
struct InstitutionData {
    values: HashMap<String, PairedValue>,
}

/// Evaluate the configured cross-period rules in input order.
///
/// `scope_matches` implements Python `re.search` semantics. The caller owns
/// regex compilation and may cache patterns; an invalid pattern should return
/// `false`, matching the legacy `_safe_regex_match` behavior.
pub fn run_current_value_check(
    current: &CentralDataset,
    previous: Option<&CentralDataset>,
    rules: &[S3CrossPeriodRule],
    target_unit: &str,
    mut scope_matches: impl FnMut(&str, &str) -> bool,
) -> S3CurrentValueCheckResult {
    let (mut store, frequencies, institution_names, region_names) =
        build_data_store(current, previous, target_unit);
    let mut result = S3CurrentValueCheckResult::default();

    for rule in rules {
        if is_disabled(&rule.disabled) {
            continue;
        }
        let rule_id = if !rule.rule_code.is_empty() {
            rule.rule_code.as_str()
        } else {
            rule.rule_name.as_str()
        };
        let left_expression = rule.left_expression.trim();
        let right_expression = rule.right_expression.trim();
        let formula = rule.formula.trim();

        if previous.is_none()
            && (left_expression.contains('{')
                || right_expression.contains('{')
                || rule.premise.contains('{'))
        {
            continue;
        }
        if left_expression.is_empty() || right_expression.is_empty() {
            continue;
        }
        if !rule_frequency_ok(
            &format!("{left_expression}{right_expression}{}", rule.premise),
            &frequencies,
        ) {
            continue;
        }

        for (institution_region, data) in &mut store.ordered {
            let institution_code = &institution_region[..4.min(institution_region.len())];
            let region_start = institution_region.len().saturating_sub(7);
            let region_code = &institution_region[region_start..];
            if !scope_matches(&rule.institution_pattern, institution_code)
                || !scope_matches(&rule.region_pattern, region_code)
            {
                continue;
            }

            let premise = rule.premise.trim();
            if !premise.is_empty() {
                let Some(premise_value) = calculate_expression(
                    premise,
                    data,
                    &[("Thd", 1.0)],
                    &mut result.logs,
                    institution_code,
                    region_code,
                    "前提条件",
                    rule_id,
                ) else {
                    continue;
                };
                if premise_value == 0.0 {
                    continue;
                }
            }

            let Some(left_value) = calculate_expression(
                left_expression,
                data,
                &[("Thd", 1.0)],
                &mut result.logs,
                institution_code,
                region_code,
                "left_data",
                rule_id,
            ) else {
                continue;
            };
            let Some(right_value) = calculate_expression(
                right_expression,
                data,
                &[("Thd", 1.0)],
                &mut result.logs,
                institution_code,
                region_code,
                "right_data",
                rule_id,
            ) else {
                continue;
            };

            let mut formula_error = String::new();
            let formula_result = if formula.is_empty() {
                None
            } else {
                match evaluate_expression(
                    formula,
                    &[("left", left_value), ("right", right_value), ("Thd", 1.0)],
                ) {
                    Ok(value) => Some(value),
                    Err(error) => {
                        formula_error = format!("校验公式计算有误：{error}");
                        result.logs.push(S3CurrentValueCheckLog {
                            institution_code: institution_code.to_owned(),
                            region_code: region_code.to_owned(),
                            error_type: "校验公式 计算有误".to_owned(),
                            rule: rule_id.to_owned(),
                            detail: formula.to_owned(),
                        });
                        None
                    }
                }
            };

            let mut difference = left_value - right_value;
            if difference.abs() <= EQUALITY_EPSILON {
                difference = 0.0;
            }
            let absolute_difference = difference.abs();
            let difference_percent = if left_value == 0.0 && right_value == 0.0 {
                None
            } else {
                Some(absolute_difference / left_value.abs().max(right_value.abs()) * 100.0)
            };

            let inverted = !rule.invert.trim().is_empty();
            let suppress_success = !rule.suppress_success.trim().is_empty();
            if !formula.is_empty() && formula_error.is_empty() && suppress_success {
                let truth = formula_result.is_some_and(ExpressionValue::truthy);
                if (inverted && truth) || (!inverted && !truth) {
                    continue;
                }
            }

            if left_value == 0.0 && right_value == 0.0 {
                continue;
            }

            let check_type = if !matches!(rule.logical_check.trim(), "" | "0") {
                "逻辑校验"
            } else {
                "数值核对"
            };
            let mut messages = Vec::new();
            if check_type == "数值核对" && (left_value == 0.0) != (right_value == 0.0) {
                messages.push("一边有，一边无;".to_owned());
            }
            if !formula.is_empty() && formula_error.is_empty() {
                let truth = formula_result.is_some_and(ExpressionValue::truthy);
                let hit = if inverted { !truth } else { truth };
                if hit && !rule.hint.trim().is_empty() {
                    messages.push(rule.hint.trim().to_owned());
                }
            }
            if !formula_error.is_empty() {
                messages.push(formula_error);
            }

            let left_substituted = substitute_tokens(left_expression, data);
            let right_substituted = substitute_tokens(right_expression, data);
            result.rows.push(S3CurrentValueCheckRow {
                data_date: current.record_date.clone(),
                institution_code: institution_code.to_owned(),
                institution_name: institution_names
                    .get(institution_code)
                    .cloned()
                    .unwrap_or_default(),
                region_code: region_code.to_owned(),
                region_name: region_names.get(region_code).cloned().unwrap_or_default(),
                rule_code: rule.rule_code.clone(),
                rule_name: rule.rule_name.clone(),
                check_type: check_type.to_owned(),
                remark: rule.remark.trim().to_owned(),
                left_value,
                right_value,
                difference,
                absolute_difference,
                difference_percent,
                explanation_required: messages.join(""),
                explanation: String::new(),
                calculation: format!(
                    "left:{left_expression};{left_substituted};right:{right_expression};{right_substituted};校验公式:{formula};"
                ),
            });
        }
    }

    result
}

/// Read the 3.0 run parameters/exemptions and 3.2 cross-period rules from one
/// or more configuration workbooks. Later workbooks override duplicate run
/// parameters, while row sheets are appended in file order like Python.
pub fn load_current_value_check_config(paths: &[PathBuf]) -> Result<S3CurrentValueCheckConfig> {
    ensure!(!paths.is_empty(), "未选择本期数值核对配置工作簿");
    let mut source_unit = "元".to_owned();
    let mut target_unit = "亿元".to_owned();
    let mut exempt_indicators = BTreeSet::new();
    let mut rules = Vec::new();
    let mut found_cross_sheet = false;

    for path in paths {
        ensure!(
            path.extension()
                .and_then(|value| value.to_str())
                .is_some_and(|value| value.eq_ignore_ascii_case("xlsx")),
            "本期数值核对配置目前只支持 XLSX：{}",
            path.display()
        );
        let sheets = read_value_columns(path, 64)
            .with_context(|| format!("读取本期数值核对配置：{}", path.display()))?;
        for sheet in &sheets {
            match sheet.name.as_str() {
                "运行参数" => {
                    let headers = header_columns(sheet);
                    let Some(name_column) = headers.get("参数") else {
                        bail!("配置“运行参数”工作表缺少“参数”列：{}", path.display());
                    };
                    let Some(value_column) = headers.get("值") else {
                        bail!("配置“运行参数”工作表缺少“值”列：{}", path.display());
                    };
                    for row in 2..=sheet.max_row {
                        let name = cell_text(sheet, row, *name_column);
                        let value = cell_text(sheet, row, *value_column);
                        match name.trim() {
                            "默认源数据单位" if !value.trim().is_empty() => {
                                source_unit = value.trim().to_owned();
                            }
                            "默认目标单位" if !value.trim().is_empty() => {
                                target_unit = value.trim().to_owned();
                            }
                            _ => {}
                        }
                    }
                }
                "单位换算例外" => {
                    let headers = header_columns(sheet);
                    if let Some(column) = headers.get("指标代码") {
                        for row in 2..=sheet.max_row {
                            let code = cell_text(sheet, row, *column)
                                .replace('\'', "")
                                .trim()
                                .to_owned();
                            if !code.is_empty() {
                                exempt_indicators.insert(code);
                            }
                        }
                    }
                }
                "跨期数值核对" => {
                    found_cross_sheet = true;
                    rules.extend(
                        read_cross_period_rules(sheet).with_context(|| {
                            format!("解析“跨期数值核对”规则：{}", path.display())
                        })?,
                    );
                }
                _ => {}
            }
        }
    }

    ensure!(found_cross_sheet, "配置工作簿中未找到“跨期数值核对”工作表");
    Ok(S3CurrentValueCheckConfig {
        source_unit,
        target_unit,
        exempt_indicators,
        rules,
    })
}

/// Run the S3-F04 service and create the Python-compatible result
/// workbook plus the optional separate flow-log workbook.
pub fn run_current_value_check_files(
    current_paths: &[PathBuf],
    previous_paths: &[PathBuf],
    config_paths: &[PathBuf],
    output_path: impl AsRef<Path>,
    log_path: Option<&Path>,
    scope_matches: impl FnMut(&str, &str) -> bool,
) -> Result<S3CurrentValueCheckReport> {
    let total_started = Instant::now();
    ensure!(!current_paths.is_empty(), "本期数值核对：请先选择数据文件");
    let output_path = resolve_new_xlsx_path(output_path.as_ref(), "本期数值核对结果")?;
    let log_path = log_path
        .map(|path| resolve_new_xlsx_path(path, "本期数值核对运行日志"))
        .transpose()?;
    if let Some(log_path) = &log_path {
        ensure!(
            !same_path(&output_path, log_path),
            "结果工作簿和运行日志不能使用同一输出路径"
        );
    }

    let config_started = Instant::now();
    let config = load_current_value_check_config(config_paths)?;
    let config_read_time_ms = config_started.elapsed().as_millis();
    let conversion_factor = unit_factor(&config.source_unit) / unit_factor(&config.target_unit);

    let input_started = Instant::now();
    let current = load_central_dataset(current_paths, conversion_factor, &config.exempt_indicators)
        .context("读取本期数据失败")?;
    let previous = if previous_paths.is_empty() {
        None
    } else {
        Some(
            load_central_dataset(previous_paths, conversion_factor, &config.exempt_indicators)
                .context("读取上期数据失败")?,
        )
    };
    let current_records = current.records.len();
    let previous_records = previous.as_ref().map_or(0, |dataset| dataset.records.len());
    let input_read_time_ms = input_started.elapsed().as_millis();

    let calculation_started = Instant::now();
    let mut result = run_current_value_check(
        &current,
        previous.as_ref(),
        &config.rules,
        &config.target_unit,
        scope_matches,
    );
    let calculation_time_ms = calculation_started.elapsed().as_millis();
    let disabled_rules = config
        .rules
        .iter()
        .filter(|rule| is_disabled(&rule.disabled))
        .count();
    let enabled_rules = config.rules.len().saturating_sub(disabled_rules);
    let previous_date = previous
        .as_ref()
        .map(|dataset| dataset.record_date.as_str());
    let date_message = match previous_date {
        Some(date) => format!(
            "数据日期：本期 {}；上期 {date}",
            date_or_unknown(&current.record_date)
        ),
        None => format!(
            "数据日期：本期 {}（未选上期，{{...}} 规则跳过）",
            date_or_unknown(&current.record_date)
        ),
    };
    result.logs.splice(
        0..0,
        [
            S3CurrentValueCheckLog {
                institution_code: "配置载入".to_owned(),
                region_code: format!(
                    "本期数值核对规则：读取 {} 条，启用 {enabled_rules} 条，禁用 {disabled_rules} 条",
                    config.rules.len()
                ),
                error_type: String::new(),
                rule: String::new(),
                detail: String::new(),
            },
            S3CurrentValueCheckLog {
                institution_code: "配置载入".to_owned(),
                region_code: date_message,
                error_type: String::new(),
                rule: String::new(),
                detail: String::new(),
            },
            S3CurrentValueCheckLog {
                institution_code: "规则命中".to_owned(),
                region_code: format!("输出 {} 行", result.rows.len()),
                error_type: String::new(),
                rule: String::new(),
                detail: String::new(),
            },
        ],
    );

    let workbook_started = Instant::now();
    write_check_workbook(&output_path, &result.rows)?;
    validate_check_workbook(&output_path, result.rows.len())?;
    if let Some(path) = log_path.as_deref() {
        write_log_workbook(path, &result.logs)?;
    }
    let workbook_write_time_ms = workbook_started.elapsed().as_millis();

    let mut input_paths = current_paths.to_vec();
    input_paths.extend_from_slice(previous_paths);
    Ok(S3CurrentValueCheckReport {
        status: "completed",
        output_workbook: output_path,
        log_workbook: log_path,
        rows: result.rows.len(),
        total_rules: config.rules.len(),
        enabled_rules,
        disabled_rules,
        current_records,
        previous_records,
        config_reader: "calamine",
        input_reader: input_reader_for_paths(&input_paths),
        writer: "rust_xlsxwriter",
        config_read_time_ms,
        input_read_time_ms,
        calculation_time_ms,
        workbook_write_time_ms,
        total_time_ms: total_started.elapsed().as_millis(),
    })
}

pub(crate) fn resolve_new_xlsx_path(path: &Path, description: &str) -> Result<PathBuf> {
    ensure!(
        path.extension()
            .and_then(|value| value.to_str())
            .is_some_and(|value| value.eq_ignore_ascii_case("xlsx")),
        "{description} 输出必须使用 .xlsx"
    );
    let parent = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    ensure!(
        parent.is_dir(),
        "{description} 输出目录不存在：{}",
        parent.display()
    );
    let parent = fs::canonicalize(parent)
        .with_context(|| format!("规范化 {description} 输出目录 {}", parent.display()))?;
    let file_name = path
        .file_name()
        .with_context(|| format!("{description} 输出路径缺少文件名"))?;
    let resolved = parent.join(file_name);
    ensure!(
        !resolved.exists(),
        "拒绝覆盖已有{description}：{}",
        resolved.display()
    );
    Ok(resolved)
}

#[cfg(windows)]
fn same_path(left: &Path, right: &Path) -> bool {
    left.to_string_lossy()
        .eq_ignore_ascii_case(&right.to_string_lossy())
}

#[cfg(not(windows))]
fn same_path(left: &Path, right: &Path) -> bool {
    left == right
}

fn date_or_unknown(date: &str) -> &str {
    if date.is_empty() {
        "未知"
    } else {
        date
    }
}

/// Return the reader selected by the input file formats for safe failure logs.
/// Successful reports use the same value after all files have been read.
pub fn input_reader_for_paths(paths: &[PathBuf]) -> String {
    let mut has_csv = false;
    let mut has_calamine = false;
    let mut unsupported = false;
    for path in paths {
        match path
            .extension()
            .and_then(|value| value.to_str())
            .map(str::to_ascii_lowercase)
            .as_deref()
        {
            Some("csv") => has_csv = true,
            Some("xlsx" | "xls") => has_calamine = true,
            _ => unsupported = true,
        }
    }
    if unsupported || (!has_csv && !has_calamine) {
        return "unavailable".to_owned();
    }
    match (has_csv, has_calamine) {
        (true, true) => "calamine+csv".to_owned(),
        (true, false) => "csv".to_owned(),
        (false, true) => "calamine".to_owned(),
        (false, false) => unreachable!(),
    }
}

pub(crate) fn load_central_dataset(
    paths: &[PathBuf],
    conversion_factor: f64,
    exempt_indicators: &BTreeSet<String>,
) -> Result<CentralDataset> {
    ensure!(!paths.is_empty(), "数据文件列表为空");
    let mut combined = CentralDataset::default();
    for path in paths {
        let file_name = path.file_name().unwrap_or_default().to_string_lossy();
        let extension = path
            .extension()
            .and_then(|value| value.to_str())
            .unwrap_or_default();
        let dataset = if extension.eq_ignore_ascii_case("csv") {
            let bytes =
                fs::read(path).with_context(|| format!("读取导数文件：{}", path.display()))?;
            parse_central_csv(&file_name, &bytes, conversion_factor, exempt_indicators)?
        } else if extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xls") {
            load_central_workbook_dataset(path, conversion_factor, exempt_indicators)?
        } else {
            bail!("S3-F04 只支持 CSV、XLSX、XLS 导数：{}", path.display());
        };
        combined.records.extend(dataset.records);
        combined.issues.extend(dataset.issues);
    }
    let dates = combined
        .records
        .iter()
        .filter(|record| !record.record_date.is_empty())
        .map(|record| record.record_date.clone())
        .collect::<BTreeSet<_>>();
    ensure!(
        dates.len() <= 1,
        "同一期文件数据日期不一致：{}",
        dates.into_iter().collect::<Vec<_>>().join("、")
    );
    combined.record_date = dates.into_iter().next().unwrap_or_default();
    Ok(combined)
}

fn load_central_workbook_dataset(
    path: &Path,
    conversion_factor: f64,
    exempt_indicators: &BTreeSet<String>,
) -> Result<CentralDataset> {
    let extension = path
        .extension()
        .and_then(|value| value.to_str())
        .unwrap_or_default();
    let is_xls = extension.eq_ignore_ascii_case("xls");
    let format_name = if is_xls { "XLS" } else { "XLSX" };
    let sheets = read_value_columns_with_iso_dates(path, usize::MAX)
        .with_context(|| format!("读取{format_name}导数工作簿失败：{}", path.display()))?;

    for sheet in sheets {
        let rows = workbook_rows(&sheet, is_xls)
            .with_context(|| format!("规范化工作表“{}”：{}", sheet.name, path.display()))?;
        if rows
            .first()
            .is_some_and(|(row_number, fields)| *row_number == 1 && header_matches(fields))
        {
            return parse_central_grid(
                &path.file_name().unwrap_or_default().to_string_lossy(),
                &rows,
                conversion_factor,
                exempt_indicators,
            );
        }
    }

    bail!(
        "{format_name} 首行表头与大集中导数 14 列不符：{}",
        path.display()
    )
}

fn workbook_rows(sheet: &SheetValues, is_xls: bool) -> Result<Vec<(usize, Vec<String>)>> {
    let mut sparse_rows: BTreeMap<usize, BTreeMap<usize, String>> = BTreeMap::new();
    for (address, value) in &sheet.values {
        let (row, column) = parse_a1_coordinate(address)
            .with_context(|| format!("解析工作表“{}”单元格坐标 {address}", sheet.name))?;
        let text = workbook_cell_text(value, is_xls)
            .with_context(|| format!("读取工作表“{}”单元格 {address}", sheet.name))?;
        sparse_rows.entry(row).or_default().insert(column, text);
    }

    Ok(sparse_rows
        .into_iter()
        .map(|(row_number, cells)| {
            let width = cells.last_key_value().map_or(0, |(column, _)| *column);
            let mut fields = vec![String::new(); width];
            for (column, value) in cells {
                fields[column - 1] = value;
            }
            (row_number, fields)
        })
        .collect())
}

fn parse_a1_coordinate(address: &str) -> Result<(usize, usize)> {
    let split = address
        .find(|character: char| character.is_ascii_digit())
        .context("坐标缺少行号")?;
    let (column_letters, row_digits) = address.split_at(split);
    ensure!(!column_letters.is_empty(), "坐标缺少列号");
    ensure!(!row_digits.is_empty(), "坐标缺少行号");

    let mut column = 0usize;
    for character in column_letters.bytes() {
        let upper = character.to_ascii_uppercase();
        ensure!(upper.is_ascii_uppercase(), "坐标包含非法列号");
        column = column
            .checked_mul(26)
            .and_then(|value| value.checked_add((upper - b'A' + 1) as usize))
            .context("列号超出范围")?;
    }
    let row = row_digits.parse::<usize>().context("行号不是正整数")?;
    ensure!(row > 0 && column > 0, "坐标必须从 A1 开始");
    Ok((row, column))
}

fn workbook_cell_text(value: &CellValue, is_xls: bool) -> Result<String> {
    Ok(match value {
        CellValue::Empty => String::new(),
        CellValue::Text(value) => value.clone(),
        CellValue::Number(value) | CellValue::DateTimeSerial(value) => legacy_number_text(*value),
        CellValue::Integer(value) => value.to_string(),
        CellValue::Boolean(value) if is_xls => if *value { "TRUE" } else { "FALSE" }.to_owned(),
        // Python's XLSX reader uses str(bool), whose spelling is title case.
        CellValue::Boolean(value) => if *value { "True" } else { "False" }.to_owned(),
        CellValue::DateTimeIso(value) => legacy_datetime_text(value, is_xls)?,
        CellValue::DurationIso(value) => {
            bail!("XLS/XLSX 时长单元格暂不支持按 Python timedelta 文本转换：{value}");
        }
        CellValue::Error(value) => legacy_excel_error_text(value, is_xls)?,
    })
}

fn legacy_number_text(value: f64) -> String {
    if value == 0.0 {
        return "0".to_owned();
    }
    if value.is_finite() && value.fract() == 0.0 {
        return format!("{value:.0}");
    }
    python_float_repr(value)
}

fn legacy_datetime_text(value: &str, is_xls: bool) -> Result<String> {
    let datetime = NaiveDateTime::parse_from_str(value, "%Y-%m-%dT%H:%M:%S%.f")
        .or_else(|_| NaiveDateTime::parse_from_str(value, "%Y-%m-%d %H:%M:%S%.f"))
        .or_else(|_| {
            NaiveDate::parse_from_str(value, "%Y-%m-%d")
                .map(|date| date.and_hms_opt(0, 0, 0).expect("valid midnight"))
        })
        .with_context(|| format!("无法解析 Calamine 日期值 {value:?}"))?;
    if is_xls {
        return Ok(datetime.date().format("%Y-%m-%d").to_string());
    }
    let mut text = datetime.format("%Y-%m-%d %H:%M:%S").to_string();
    let microseconds = datetime.and_utc().timestamp_subsec_micros();
    if microseconds != 0 {
        text.push_str(&format!(".{microseconds:06}"));
    }
    Ok(text)
}

fn legacy_excel_error_text(value: &str, is_xls: bool) -> Result<String> {
    let (text, xls_code) = match value {
        "Div0" => ("#DIV/0!", 7),
        "NA" => ("#N/A", 42),
        "Name" => ("#NAME?", 29),
        "Null" => ("#NULL!", 0),
        "Num" => ("#NUM!", 36),
        "Ref" => ("#REF!", 23),
        "Value" => ("#VALUE!", 15),
        "GettingData" => ("#GETTING_DATA", 43),
        value if value.starts_with('#') => return Ok(value.to_owned()),
        _ => bail!("不支持的 Excel 错误单元格类型：{value}"),
    };
    if is_xls {
        Ok(format!("#ERR{xls_code}"))
    } else {
        Ok(text.to_owned())
    }
}

fn header_columns(sheet: &SheetValues) -> HashMap<String, usize> {
    let mut headers = HashMap::new();
    for column in 1..=sheet.max_column {
        let value = cell_text(sheet, 1, column);
        if !value.trim().is_empty() {
            headers.insert(value.trim().to_owned(), column);
        }
    }
    headers
}

fn read_cross_period_rules(sheet: &SheetValues) -> Result<Vec<S3CrossPeriodRule>> {
    const REQUIRED: [&str; 14] = [
        "校验编码",
        "校验名称",
        "机构类代码",
        "地区代码",
        "前提条件",
        "left",
        "right",
        "校验公式",
        "校验提示",
        "取反标识",
        "无误是否提示",
        "是否逻辑校验",
        "禁用",
        "备注",
    ];
    let headers = header_columns(sheet);
    let missing = REQUIRED
        .iter()
        .filter(|header| !headers.contains_key(**header))
        .copied()
        .collect::<Vec<_>>();
    ensure!(
        missing.is_empty(),
        "“跨期数值核对”工作表缺少列：{}",
        missing.join("、")
    );
    let mut rules = Vec::new();
    for row in 2..=sheet.max_row {
        let mut values = HashMap::new();
        for header in REQUIRED {
            values.insert(header, cell_text(sheet, row, headers[header]));
        }
        if values.values().all(|value| value.trim().is_empty()) {
            continue;
        }
        rules.push(S3CrossPeriodRule {
            disabled: values["禁用"].clone(),
            rule_code: values["校验编码"].trim().to_owned(),
            rule_name: values["校验名称"].trim().to_owned(),
            institution_pattern: values["机构类代码"].trim().to_owned(),
            region_pattern: values["地区代码"].trim().to_owned(),
            premise: values["前提条件"].clone(),
            left_expression: values["left"].clone(),
            right_expression: values["right"].clone(),
            formula: values["校验公式"].clone(),
            hint: values["校验提示"].clone(),
            invert: values["取反标识"].clone(),
            suppress_success: values["无误是否提示"].clone(),
            logical_check: values["是否逻辑校验"].clone(),
            remark: values["备注"].clone(),
        });
    }
    Ok(rules)
}

fn cell_text(sheet: &SheetValues, row: usize, column: usize) -> String {
    let address = format!("{}{}", column_name(column), row);
    match sheet.values.get(&address).unwrap_or(&CellValue::Empty) {
        CellValue::Empty => String::new(),
        CellValue::Text(value) => value.clone(),
        CellValue::Number(value) => format!("{value:?}"),
        CellValue::Integer(value) => value.to_string(),
        CellValue::Boolean(value) => value.to_string(),
        CellValue::DateTimeSerial(value) => format!("{value:?}"),
        CellValue::DateTimeIso(value) | CellValue::DurationIso(value) => value.clone(),
        CellValue::Error(value) => value.clone(),
    }
}

fn column_name(mut column: usize) -> String {
    let mut name = String::new();
    while column > 0 {
        let remainder = (column - 1) % 26;
        name.insert(0, (b'A' + remainder as u8) as char);
        column = (column - 1) / 26;
    }
    name
}

fn write_check_workbook(path: &Path, rows: &[S3CurrentValueCheckRow]) -> Result<()> {
    const HEADERS: [&str; 17] = [
        "数据日期",
        "机构类代码",
        "机构类名称",
        "地区代码",
        "地区名称",
        "校验编码",
        "校验名称",
        "校验类型",
        "备注",
        "左值",
        "右值",
        "差异",
        "差异绝对值",
        "差异幅度(%)",
        "是否说明",
        "说明内容",
        "计算公式",
    ];
    let widths = [
        12.0, 10.0, 22.0, 10.0, 10.0, 10.0, 22.0, 10.0, 14.0, 12.0, 12.0, 12.0, 12.0, 12.0, 24.0,
        24.0, 44.0,
    ];
    let header_format = Format::new()
        .set_bold()
        .set_font_color(Color::RGB(0xFF_FF_FF))
        .set_background_color(Color::RGB(0x1F_4E_78))
        .set_text_wrap();
    let number_format = Format::new().set_num_format("0.00;-0.00;");
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .set_name("对比结果")
        .context("设置 S3-F04 结果工作表名称失败")?;
    set_python_default_page_margins(worksheet);
    for (column, header) in HEADERS.iter().enumerate() {
        worksheet.write_string_with_format(0, column as u16, *header, &header_format)?;
    }
    for (index, row) in rows.iter().enumerate() {
        let output_row = index as u32 + 1;
        let strings = [
            row.data_date.as_str(),
            row.institution_code.as_str(),
            row.institution_name.as_str(),
            row.region_code.as_str(),
            row.region_name.as_str(),
            row.rule_code.as_str(),
            row.rule_name.as_str(),
            row.check_type.as_str(),
            row.remark.as_str(),
        ];
        for (column, value) in strings.iter().enumerate() {
            worksheet.write_string(output_row, column as u16, *value)?;
        }
        for (column, value) in [
            row.left_value,
            row.right_value,
            row.difference,
            row.absolute_difference,
        ]
        .into_iter()
        .enumerate()
        {
            worksheet.write_number_with_format(
                output_row,
                (column + 9) as u16,
                value,
                &number_format,
            )?;
        }
        if let Some(value) = row.difference_percent {
            worksheet.write_number_with_format(output_row, 13, value, &number_format)?;
        } else {
            worksheet.write_blank(output_row, 13, &Format::new())?;
        }
        for (column, value) in [
            row.explanation_required.as_str(),
            row.explanation.as_str(),
            row.calculation.as_str(),
        ]
        .into_iter()
        .enumerate()
        {
            worksheet.write_string(output_row, (column + 14) as u16, value)?;
        }
    }
    worksheet
        .set_freeze_panes(1, 0)
        .context("冻结 S3-F04 结果表标题行失败")?;
    for (column, width) in widths.iter().enumerate() {
        // Python/openpyxl stores these widths directly in the XLSX column
        // dimension. rust_xlsxwriter::set_column_width() treats its argument
        // as the visible character width and adds the default 5 px cell
        // padding in the serialized width. Use the equivalent pixel width so
        // the output keeps the Python workbook's exact width metadata.
        let width_pixels = (*width * 7.0_f64).round() as u32;
        worksheet
            .set_column_width_pixels(column as u16, width_pixels)
            .with_context(|| format!("设置 S3-F04 结果表第 {} 列宽度失败", column + 1))?;
    }
    workbook
        .save(path)
        .with_context(|| format!("保存 S3-F04 结果工作簿失败：{}", path.display()))?;
    Ok(())
}

fn write_log_workbook(path: &Path, logs: &[S3CurrentValueCheckLog]) -> Result<()> {
    let header_format = Format::new()
        .set_bold()
        .set_font_color(Color::RGB(0xFF_FF_FF))
        .set_background_color(Color::RGB(0x1F_4E_78));
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .set_name("运行日志")
        .context("设置 S3-F04 运行日志工作表名称失败")?;
    set_python_default_page_margins(worksheet);
    for (column, header) in ["时间", "机构类代码", "地区代码", "错误类型", "规则", "详情"]
        .iter()
        .enumerate()
    {
        worksheet.write_string_with_format(0, column as u16, *header, &header_format)?;
    }
    for (index, log) in logs.iter().enumerate() {
        let row = index as u32 + 1;
        let timestamp = Local::now().format("%Y-%m-%d %H:%M:%S").to_string();
        for (column, value) in [
            timestamp.as_str(),
            log.institution_code.as_str(),
            log.region_code.as_str(),
            log.error_type.as_str(),
            log.rule.as_str(),
            log.detail.as_str(),
        ]
        .iter()
        .enumerate()
        {
            worksheet.write_string(row, column as u16, *value)?;
        }
    }
    workbook
        .save(path)
        .with_context(|| format!("保存 S3-F04 运行日志失败：{}", path.display()))?;
    Ok(())
}

fn validate_check_workbook(path: &Path, expected_rows: usize) -> Result<()> {
    let sheets = read_value_columns(path, 17).context("独立读回 S3-F04 结果工作簿失败")?;
    ensure!(sheets.len() == 1, "S3-F04 结果工作簿应只有一个工作表");
    let sheet = &sheets[0];
    ensure!(sheet.name == "对比结果", "S3-F04 结果工作表名称错误");
    ensure!(
        sheet.max_row == expected_rows + 1,
        "S3-F04 结果工作簿行数校验失败"
    );
    ensure!(sheet.max_column == 17, "S3-F04 结果工作簿应有 17 列");
    ensure!(
        cell_text(sheet, 1, 1) == "数据日期",
        "S3-F04 结果工作簿表头读回不一致"
    );
    Ok(())
}

fn set_python_default_page_margins(worksheet: &mut excel_writer::Worksheet) {
    worksheet.set_margins(0.75, 0.75, 1.0, 1.0, 0.5, 0.5);
}

#[derive(Default)]
struct DataStore {
    ordered: Vec<(String, InstitutionData)>,
    positions: HashMap<String, usize>,
}

fn build_data_store(
    current: &CentralDataset,
    previous: Option<&CentralDataset>,
    target_unit: &str,
) -> (
    DataStore,
    HashSet<String>,
    HashMap<String, String>,
    HashMap<String, String>,
) {
    let mut store = DataStore::default();
    let mut frequencies = HashSet::new();
    let mut institution_names = HashMap::new();
    let mut region_names = HashMap::new();
    let sentinel = 0.01 / unit_factor(target_unit);

    feed_dataset(
        current,
        "cur",
        sentinel,
        &mut store,
        &mut frequencies,
        &mut institution_names,
        &mut region_names,
    );
    if let Some(previous) = previous {
        feed_dataset(
            previous,
            "pre",
            sentinel,
            &mut store,
            &mut frequencies,
            &mut institution_names,
            &mut region_names,
        );
    }
    (store, frequencies, institution_names, region_names)
}

fn feed_dataset(
    dataset: &CentralDataset,
    side: &str,
    sentinel: f64,
    store: &mut DataStore,
    frequencies: &mut HashSet<String>,
    institution_names: &mut HashMap<String, String>,
    region_names: &mut HashMap<String, String>,
) {
    for record in &dataset.records {
        let institution_region = format!("{}{}", record.org_code, record.region_code);
        frequencies.insert(format!("{}{}", record.frequency, record.batch));
        institution_names
            .entry(record.org_code.clone())
            .or_insert_with(|| record.org_name.clone());
        region_names
            .entry(record.region_code.clone())
            .or_insert_with(|| record.region_name.clone());

        let position = match store.positions.get(&institution_region) {
            Some(position) => *position,
            None => {
                let position = store.ordered.len();
                store.positions.insert(institution_region.clone(), position);
                store
                    .ordered
                    .push((institution_region.clone(), InstitutionData::default()));
                position
            }
        };
        let value = numeric_value(&record.value);
        let value = if value == 0.0 { sentinel } else { value };
        let entry = store.ordered[position]
            .1
            .values
            .entry(record.attr_key())
            .or_default();
        match side {
            "cur" => entry.current = Some(value),
            "pre" => entry.previous = Some(value),
            _ => unreachable!("only current and previous datasets are fed"),
        }
    }
}

fn numeric_value(value: &CentralValue) -> f64 {
    match value {
        CentralValue::Number(value) => *value,
        CentralValue::Text(_) | CentralValue::Empty => 0.0,
    }
}

fn unit_factor(unit: &str) -> f64 {
    match unit {
        "元" => 1.0,
        "十元" => 10.0,
        "百元" => 100.0,
        "千元" => 1_000.0,
        "万元" => 10_000.0,
        "十万元" => 100_000.0,
        "百万元" => 1_000_000.0,
        "千万" | "千万元" => 10_000_000.0,
        "亿元" => 100_000_000.0,
        _ => 1.0,
    }
}

pub(crate) fn s3_target_unit_factor(unit: &str) -> f64 {
    unit_factor(unit)
}

fn python_float_repr(value: f64) -> String {
    if value.is_nan() {
        return "nan".to_owned();
    }
    if value == f64::INFINITY {
        return "inf".to_owned();
    }
    if value == f64::NEG_INFINITY {
        return "-inf".to_owned();
    }

    let representation = format!("{value:?}");
    if let Some((mantissa, exponent)) = representation.split_once('e') {
        let exponent = exponent.parse::<i32>().unwrap_or_default();
        if (-4..16).contains(&exponent) {
            return expand_scientific_notation(mantissa, exponent);
        }
        return format!("{}e{exponent:+03}", mantissa.trim_end_matches(".0"));
    }
    if value != 0.0 && (value.abs() < 1e-4 || value.abs() >= 1e16) {
        return fixed_to_scientific_notation(&representation);
    }
    representation
}

pub(crate) fn s3_python_float_repr(value: f64) -> String {
    python_float_repr(value)
}

fn expand_scientific_notation(mantissa: &str, exponent: i32) -> String {
    let (sign, unsigned) = mantissa
        .strip_prefix('-')
        .map_or(("", mantissa), |value| ("-", value));
    let (integer, fraction) = unsigned.split_once('.').unwrap_or((unsigned, ""));
    let digits = format!("{integer}{fraction}");
    let decimal_position = integer.len() as i32 + exponent;
    if decimal_position <= 0 {
        format!(
            "{sign}0.{}{}",
            "0".repeat(decimal_position.unsigned_abs() as usize),
            digits
        )
    } else if decimal_position as usize >= digits.len() {
        let mut expanded = digits;
        expanded.push_str(&"0".repeat(decimal_position as usize - expanded.len()));
        format!("{sign}{expanded}.0")
    } else {
        let position = decimal_position as usize;
        format!("{sign}{}.{}", &digits[..position], &digits[position..])
    }
}

fn fixed_to_scientific_notation(value: &str) -> String {
    let (sign, unsigned) = value
        .strip_prefix('-')
        .map_or(("", value), |value| ("-", value));
    let decimal_position = unsigned.find('.').unwrap_or(unsigned.len()) as i32;
    let mut digits = unsigned.replace('.', "");
    let leading_zeroes = digits.chars().take_while(|digit| *digit == '0').count();
    digits.drain(..leading_zeroes);
    while digits.len() > 1 && digits.ends_with('0') {
        digits.pop();
    }
    if digits.is_empty() {
        return format!("{value}e+00");
    }
    let exponent = decimal_position - leading_zeroes as i32 - 1;
    let mantissa = if digits.len() == 1 {
        digits
    } else {
        format!("{}.{}", &digits[..1], &digits[1..])
    };
    format!("{sign}{mantissa}e{exponent:+03}")
}

fn is_disabled(value: &str) -> bool {
    matches!(value.trim(), "是" | "1" | "true" | "True" | "Y" | "y")
}

fn rule_frequency_ok(expression: &str, frequencies: &HashSet<String>) -> bool {
    for token in expression_tokens(expression) {
        let segments = token.split(',').map(str::trim).collect::<Vec<_>>();
        if segments.len() != 1 && segments.len() != 5 {
            return false;
        }
        if segments.len() == 5 && !frequencies.contains(&format!("{}{}", segments[3], segments[4]))
        {
            return false;
        }
    }
    true
}

fn expression_tokens(expression: &str) -> Vec<&str> {
    let mut tokens = Vec::new();
    let mut opening: Option<(char, usize)> = None;
    for (index, character) in expression.char_indices() {
        if matches!(character, '[' | '{') {
            opening = Some((character, index + character.len_utf8()));
        } else if character == ']' || character == '}' {
            if let Some((kind, start)) = opening.take() {
                if (kind == '[' && character == ']') || (kind == '{' && character == '}') {
                    tokens.push(&expression[start..index]);
                }
            }
        }
    }
    tokens
}

fn substitute_tokens(expression: &str, data: &InstitutionData) -> String {
    let mut output = String::with_capacity(expression.len());
    let mut cursor = 0;
    let characters = expression.char_indices().collect::<Vec<_>>();
    let mut index = 0;
    while index < characters.len() {
        let (start, character) = characters[index];
        let (closing, value_side) = match character {
            '{' => ('}', "pre"),
            '[' => (']', "cur"),
            _ => {
                index += 1;
                continue;
            }
        };
        let Some((closing_index, (end, _))) = characters[index + 1..]
            .iter()
            .enumerate()
            .find(|(_, (_, candidate))| *candidate == closing)
            .map(|(offset, item)| (index + 1 + offset, *item))
        else {
            index += 1;
            continue;
        };
        output.push_str(&expression[cursor..start]);
        let body = &expression[start + 1..end];
        let segments = body.split(',').map(str::trim).collect::<Vec<_>>();
        let unique_code = match segments.as_slice() {
            [single] => (*single).to_owned(),
            [indicator, data_attr, currency, frequency, batch] => {
                format!("{indicator}{data_attr}{currency}{frequency}{batch}")
            }
            _ => String::new(),
        };
        let value = data
            .values
            .get(&unique_code)
            .and_then(|value| match value_side {
                "cur" => value.current,
                "pre" => value.previous,
                _ => None,
            });
        match value {
            Some(value) => output.push_str(&python_float_repr(value)),
            // Python's token substitution uses the literal "0" when no
            // matching current/previous record exists. Keep that distinct
            // from a present numeric value rendered with float repr ("0.0").
            None => output.push('0'),
        }
        cursor = end + closing.len_utf8();
        index = closing_index + 1;
    }
    output.push_str(&expression[cursor..]);
    output
}

fn calculate_expression(
    expression: &str,
    data: &InstitutionData,
    variables: &[(&str, f64)],
    logs: &mut Vec<S3CurrentValueCheckLog>,
    institution_code: &str,
    region_code: &str,
    mode: &str,
    rule_id: &str,
) -> Option<f64> {
    let substituted = substitute_tokens(expression, data);
    match evaluate_expression(substituted.as_str(), variables) {
        Ok(value) => Some(value.as_number()),
        Err(error) => {
            logs.push(S3CurrentValueCheckLog {
                institution_code: institution_code.to_owned(),
                region_code: region_code.to_owned(),
                error_type: format!("{mode}计算有误"),
                rule: rule_id.to_owned(),
                detail: format!("{expression} -> {error}"),
            });
            None
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub(crate) enum ExpressionValue {
    Number(f64),
    Boolean(bool),
}

impl ExpressionValue {
    pub(crate) fn as_number(self) -> f64 {
        match self {
            Self::Number(value) => value,
            Self::Boolean(value) => f64::from(value),
        }
    }

    pub(crate) fn truthy(self) -> bool {
        self.as_number() != 0.0
    }

    pub(crate) fn as_boolean(self) -> Option<bool> {
        match self {
            Self::Boolean(value) => Some(value),
            Self::Number(_) => None,
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
enum ExpressionToken {
    Number(f64),
    Identifier(String),
    Plus,
    Minus,
    Multiply,
    Divide,
    Power,
    Equal,
    NotEqual,
    Less,
    LessEqual,
    Greater,
    GreaterEqual,
    LeftParen,
    RightParen,
    Comma,
    End,
}

fn evaluate_expression(
    expression: &str,
    variables: &[(&str, f64)],
) -> Result<ExpressionValue, String> {
    let tokens = tokenize_expression(expression)?;
    let mut parser = ExpressionParser {
        tokens,
        position: 0,
        variables,
    };
    let value = parser.parse_or()?;
    if parser.peek() != &ExpressionToken::End {
        return Err(format!(
            "表达式包含无法识别的内容「{}」",
            parser.describe_next()
        ));
    }
    Ok(value)
}

pub(crate) fn evaluate_s3_business_expression(
    expression: &str,
    tolerance: f64,
) -> Result<ExpressionValue, String> {
    evaluate_expression(expression, &[("Thd", tolerance)])
}

fn tokenize_expression(expression: &str) -> Result<Vec<ExpressionToken>, String> {
    let characters = expression.char_indices().collect::<Vec<_>>();
    let mut tokens = Vec::new();
    let mut index = 0;
    while index < characters.len() {
        let (byte_index, character) = characters[index];
        if character.is_whitespace() {
            index += 1;
            continue;
        }
        if character.is_ascii_digit()
            || (character == '.'
                && characters
                    .get(index + 1)
                    .is_some_and(|(_, next)| next.is_ascii_digit()))
        {
            let start = index;
            index += 1;
            while index < characters.len()
                && (characters[index].1.is_ascii_digit() || characters[index].1 == '.')
            {
                index += 1;
            }
            if index < characters.len() && matches!(characters[index].1, 'e' | 'E') {
                index += 1;
                if index < characters.len() && matches!(characters[index].1, '+' | '-') {
                    index += 1;
                }
                let exponent_start = index;
                while index < characters.len() && characters[index].1.is_ascii_digit() {
                    index += 1;
                }
                if index == exponent_start {
                    return Err("数值指数缺少数字".to_owned());
                }
            }
            let start_byte = characters[start].0;
            let end_byte = characters
                .get(index)
                .map(|(byte, _)| *byte)
                .unwrap_or(expression.len());
            let number = expression[start_byte..end_byte]
                .parse::<f64>()
                .map_err(|_| "数值格式错误".to_owned())?;
            tokens.push(ExpressionToken::Number(number));
            continue;
        }
        if character.is_alphabetic() || character == '_' {
            let start = index;
            index += 1;
            while index < characters.len()
                && (characters[index].1.is_alphanumeric() || characters[index].1 == '_')
            {
                index += 1;
            }
            let start_byte = characters[start].0;
            let end_byte = characters
                .get(index)
                .map(|(byte, _)| *byte)
                .unwrap_or(expression.len());
            tokens.push(ExpressionToken::Identifier(
                expression[start_byte..end_byte].to_owned(),
            ));
            continue;
        }
        let next = characters.get(index + 1).map(|(_, next)| *next);
        match (character, next) {
            ('<', Some('=')) => {
                tokens.push(ExpressionToken::LessEqual);
                index += 2;
            }
            ('>', Some('=')) => {
                tokens.push(ExpressionToken::GreaterEqual);
                index += 2;
            }
            ('<', Some('>')) | ('!', Some('=')) => {
                tokens.push(ExpressionToken::NotEqual);
                index += 2;
            }
            ('=', Some('=')) => {
                tokens.push(ExpressionToken::Equal);
                index += 2;
            }
            ('=', _) => {
                tokens.push(ExpressionToken::Equal);
                index += 1;
            }
            ('<', _) => {
                tokens.push(ExpressionToken::Less);
                index += 1;
            }
            ('>', _) => {
                tokens.push(ExpressionToken::Greater);
                index += 1;
            }
            ('+', _) => {
                tokens.push(ExpressionToken::Plus);
                index += 1;
            }
            ('-', _) => {
                tokens.push(ExpressionToken::Minus);
                index += 1;
            }
            ('*', _) => {
                tokens.push(ExpressionToken::Multiply);
                index += 1;
            }
            ('/', _) => {
                tokens.push(ExpressionToken::Divide);
                index += 1;
            }
            ('^', _) => {
                tokens.push(ExpressionToken::Power);
                index += 1;
            }
            ('(', _) => {
                tokens.push(ExpressionToken::LeftParen);
                index += 1;
            }
            (')', _) => {
                tokens.push(ExpressionToken::RightParen);
                index += 1;
            }
            (',', _) => {
                tokens.push(ExpressionToken::Comma);
                index += 1;
            }
            _ => {
                return Err(format!(
                    "表达式包含无法识别的字符「{character}」（位置 {}）",
                    byte_index + 1
                ));
            }
        }
    }
    tokens.push(ExpressionToken::End);
    Ok(tokens)
}

struct ExpressionParser<'a> {
    tokens: Vec<ExpressionToken>,
    position: usize,
    variables: &'a [(&'a str, f64)],
}

impl ExpressionParser<'_> {
    fn peek(&self) -> &ExpressionToken {
        &self.tokens[self.position]
    }

    fn next(&mut self) -> ExpressionToken {
        let token = self.tokens[self.position].clone();
        self.position += 1;
        token
    }

    fn parse_or(&mut self) -> Result<ExpressionValue, String> {
        let mut value = self.parse_and()?;
        while self.consume_keyword("OR") {
            let right = self.parse_and()?;
            value = ExpressionValue::Boolean(value.truthy() || right.truthy());
        }
        Ok(value)
    }

    fn parse_and(&mut self) -> Result<ExpressionValue, String> {
        let mut value = self.parse_comparison()?;
        while self.consume_keyword("AND") {
            let right = self.parse_comparison()?;
            value = ExpressionValue::Boolean(value.truthy() && right.truthy());
        }
        Ok(value)
    }

    fn parse_comparison(&mut self) -> Result<ExpressionValue, String> {
        let left = self.parse_additive()?;
        let operator = self.peek().clone();
        let result = match operator {
            ExpressionToken::Equal
            | ExpressionToken::NotEqual
            | ExpressionToken::Less
            | ExpressionToken::LessEqual
            | ExpressionToken::Greater
            | ExpressionToken::GreaterEqual => {
                self.next();
                let right = self.parse_additive()?;
                let (left, right) = (left.as_number(), right.as_number());
                let value = match operator {
                    ExpressionToken::Equal => left == right,
                    ExpressionToken::NotEqual => left != right,
                    ExpressionToken::Less => left < right,
                    ExpressionToken::LessEqual => left <= right,
                    ExpressionToken::Greater => left > right,
                    ExpressionToken::GreaterEqual => left >= right,
                    _ => unreachable!(),
                };
                ExpressionValue::Boolean(value)
            }
            _ => left,
        };
        Ok(result)
    }

    fn parse_additive(&mut self) -> Result<ExpressionValue, String> {
        let mut value = self.parse_multiplicative()?;
        loop {
            let operator = self.peek().clone();
            let operation = match operator {
                ExpressionToken::Plus => 1,
                ExpressionToken::Minus => -1,
                _ => break,
            };
            self.next();
            let right = self.parse_multiplicative()?.as_number();
            let left = value.as_number();
            value = checked_number(if operation == 1 {
                left + right
            } else {
                left - right
            })?;
        }
        Ok(value)
    }

    fn parse_multiplicative(&mut self) -> Result<ExpressionValue, String> {
        let mut value = self.parse_unary()?;
        loop {
            let operator = self.peek().clone();
            let operation = match operator {
                ExpressionToken::Multiply => 1,
                ExpressionToken::Divide => -1,
                _ => break,
            };
            self.next();
            let right = self.parse_unary()?.as_number();
            let left = value.as_number();
            if operation == -1 && right == 0.0 {
                return Err("除零错误".to_owned());
            }
            value = checked_number(if operation == 1 {
                left * right
            } else {
                left / right
            })?;
        }
        Ok(value)
    }

    fn parse_unary(&mut self) -> Result<ExpressionValue, String> {
        match self.peek() {
            ExpressionToken::Minus => {
                self.next();
                checked_number(-self.parse_unary()?.as_number())
            }
            ExpressionToken::Plus => {
                self.next();
                checked_number(self.parse_unary()?.as_number())
            }
            _ => self.parse_power(),
        }
    }

    fn parse_power(&mut self) -> Result<ExpressionValue, String> {
        let mut value = self.parse_primary()?;
        while self.peek() == &ExpressionToken::Power {
            self.next();
            let right = self.parse_unary()?.as_number();
            let left = value.as_number();
            value = checked_number(left.powf(right))?;
        }
        Ok(value)
    }

    fn parse_primary(&mut self) -> Result<ExpressionValue, String> {
        match self.next() {
            ExpressionToken::Number(value) => checked_number(value),
            ExpressionToken::Identifier(name) => {
                if self.peek() == &ExpressionToken::LeftParen {
                    self.next();
                    return self.parse_function(&name);
                }
                if name.eq_ignore_ascii_case("TRUE") {
                    return Ok(ExpressionValue::Boolean(true));
                }
                if name.eq_ignore_ascii_case("FALSE") {
                    return Ok(ExpressionValue::Boolean(false));
                }
                self.variables
                    .iter()
                    .find(|(candidate, _)| candidate.eq_ignore_ascii_case(&name))
                    .map(|(_, value)| ExpressionValue::Number(*value))
                    .ok_or_else(|| format!("表达式引用了未知名称「{name}」"))
            }
            ExpressionToken::LeftParen => {
                let value = self.parse_or()?;
                if self.next() != ExpressionToken::RightParen {
                    return Err("括号不匹配".to_owned());
                }
                Ok(value)
            }
            token => Err(format!("意外的符号「{}」", token_label(&token))),
        }
    }

    fn parse_function(&mut self, name: &str) -> Result<ExpressionValue, String> {
        let normalized = name.to_ascii_uppercase();
        let mut arguments = Vec::new();
        if self.peek() == &ExpressionToken::RightParen {
            self.next();
        } else {
            loop {
                arguments.push(self.parse_or()?);
                match self.next() {
                    ExpressionToken::Comma => {}
                    ExpressionToken::RightParen => break,
                    token => {
                        return Err(format!(
                            "函数 {normalized} 参数列表格式错误：「{}」",
                            token_label(&token)
                        ));
                    }
                }
            }
        }
        match normalized.as_str() {
            "AND" if !arguments.is_empty() => Ok(ExpressionValue::Boolean(
                arguments.into_iter().all(ExpressionValue::truthy),
            )),
            "OR" if !arguments.is_empty() => Ok(ExpressionValue::Boolean(
                arguments.into_iter().any(ExpressionValue::truthy),
            )),
            "NOT" if arguments.len() == 1 => Ok(ExpressionValue::Boolean(!arguments[0].truthy())),
            "AND" | "OR" | "NOT" => Err(format!("{normalized} 参数数量错误")),
            "ABS" if arguments.len() == 1 => checked_number(arguments[0].as_number().abs()),
            "ABS" => Err("ABS 参数数量错误".to_owned()),
            "SUM" => checked_number(arguments.iter().map(|value| value.as_number()).sum()),
            _ => Err(format!("未知函数「{name}」")),
        }
    }

    fn consume_keyword(&mut self, keyword: &str) -> bool {
        let matches = matches!(
            self.peek(),
            ExpressionToken::Identifier(value) if value.eq_ignore_ascii_case(keyword)
        );
        if matches {
            self.next();
        }
        matches
    }

    fn describe_next(&self) -> String {
        token_label(self.peek())
    }
}

fn checked_number(value: f64) -> Result<ExpressionValue, String> {
    if value.is_finite() {
        Ok(ExpressionValue::Number(value))
    } else {
        Err("数值溢出".to_owned())
    }
}

fn token_label(token: &ExpressionToken) -> String {
    match token {
        ExpressionToken::Number(value) => value.to_string(),
        ExpressionToken::Identifier(value) => value.clone(),
        ExpressionToken::Plus => "+".to_owned(),
        ExpressionToken::Minus => "-".to_owned(),
        ExpressionToken::Multiply => "*".to_owned(),
        ExpressionToken::Divide => "/".to_owned(),
        ExpressionToken::Power => "^".to_owned(),
        ExpressionToken::Equal => "=".to_owned(),
        ExpressionToken::NotEqual => "<>".to_owned(),
        ExpressionToken::Less => "<".to_owned(),
        ExpressionToken::LessEqual => "<=".to_owned(),
        ExpressionToken::Greater => ">".to_owned(),
        ExpressionToken::GreaterEqual => ">=".to_owned(),
        ExpressionToken::LeftParen => "(".to_owned(),
        ExpressionToken::RightParen => ")".to_owned(),
        ExpressionToken::Comma => ",".to_owned(),
        ExpressionToken::End => "<结束>".to_owned(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::central_csv::{CentralRecord, CentralValue, EXPECTED_HEADERS};
    use std::sync::atomic::{AtomicU64, Ordering};

    static TEMP_COUNTER: AtomicU64 = AtomicU64::new(0);

    fn temporary_path(label: &str, extension: &str) -> PathBuf {
        let sequence = TEMP_COUNTER.fetch_add(1, Ordering::Relaxed);
        std::env::temp_dir().join(format!(
            "audit-s3-f04-{label}-{}-{sequence}.{extension}",
            std::process::id()
        ))
    }

    fn create_test_config(path: &Path) {
        let mut workbook = Workbook::new();
        let parameters = workbook.add_worksheet();
        parameters.set_name("运行参数").unwrap();
        parameters.write_string(0, 0, "参数").unwrap();
        parameters.write_string(0, 1, "值").unwrap();
        parameters.write_string(1, 0, "默认源数据单位").unwrap();
        parameters.write_string(1, 1, "元").unwrap();
        parameters.write_string(2, 0, "默认目标单位").unwrap();
        parameters.write_string(2, 1, "亿元").unwrap();

        let exceptions = workbook.add_worksheet();
        exceptions.set_name("单位换算例外").unwrap();
        exceptions.write_string(0, 0, "指标代码").unwrap();
        exceptions.write_string(1, 0, "I0").unwrap();

        let rules = workbook.add_worksheet();
        rules.set_name("跨期数值核对").unwrap();
        let headers = [
            "校验编码",
            "校验名称",
            "机构类代码",
            "地区代码",
            "前提条件",
            "left",
            "right",
            "校验公式",
            "校验提示",
            "取反标识",
            "无误是否提示",
            "是否逻辑校验",
            "禁用",
            "备注",
        ];
        for (column, header) in headers.iter().enumerate() {
            rules.write_string(0, column as u16, *header).unwrap();
        }
        let values = [
            "R1",
            "单位测试规则",
            "^0001$",
            "1234567",
            "",
            "[I1,余额,人民币,月,2]",
            "{I1,余额,人民币,月,2}",
            "left > right",
            "请核实",
            "",
            "是",
            "",
            "",
            "真实语义",
        ];
        for (column, value) in values.iter().enumerate() {
            rules.write_string(1, column as u16, *value).unwrap();
        }
        workbook.save(path).unwrap();
    }

    fn create_test_central_xlsx(path: &Path) {
        let mut workbook = Workbook::new();
        let cover = workbook.add_worksheet();
        cover.set_name("说明").unwrap();
        cover.write_string(0, 0, "非导数工作表").unwrap();

        let data = workbook.add_worksheet();
        data.set_name("本期导数").unwrap();
        for (column, header) in EXPECTED_HEADERS.iter().enumerate() {
            data.write_string(0, column as u16, *header).unwrap();
        }
        // Leave two physical rows blank to verify row numbers survive sparse
        // Calamine output, as they do in the old openpyxl grid.
        data.write_string(3, 0, "人民币").unwrap();
        data.write_string(3, 1, "bad-date").unwrap();
        data.write_string(3, 2, "0001").unwrap();
        data.write_boolean(3, 3, true).unwrap();
        data.write_string(3, 4, "1234567").unwrap();
        data.write_boolean(3, 5, false).unwrap();
        data.write_string(3, 7, "I1").unwrap();
        data.write_string(3, 13, "12.5").unwrap();
        workbook.save(path).unwrap();
    }

    fn record(
        org: &str,
        region: &str,
        indicator: &str,
        frequency: &str,
        batch: &str,
        value: f64,
    ) -> CentralRecord {
        CentralRecord {
            biz_class: "人民币".to_owned(),
            record_date: "2026-08-31".to_owned(),
            org_code: org.to_owned(),
            org_name: format!("机构{org}"),
            region_code: region.to_owned(),
            region_name: format!("地区{region}"),
            order_code: "1".to_owned(),
            indicator: indicator.to_owned(),
            indicator_name: indicator.to_owned(),
            data_attr: "余额".to_owned(),
            currency: "人民币".to_owned(),
            frequency: frequency.to_owned(),
            batch: batch.to_owned(),
            value: CentralValue::Number(value),
        }
    }

    fn dataset(records: Vec<CentralRecord>, date: &str) -> CentralDataset {
        CentralDataset {
            records,
            record_date: date.to_owned(),
            ..CentralDataset::default()
        }
    }

    fn rule(left: &str, right: &str, formula: &str) -> S3CrossPeriodRule {
        S3CrossPeriodRule {
            rule_code: "R1".to_owned(),
            rule_name: "规则一".to_owned(),
            institution_pattern: "".to_owned(),
            region_pattern: "".to_owned(),
            left_expression: left.to_owned(),
            right_expression: right.to_owned(),
            formula: formula.to_owned(),
            hint: "核实".to_owned(),
            ..S3CrossPeriodRule::default()
        }
    }

    #[test]
    fn python_regex_scope_matcher_uses_search_and_rejects_invalid_patterns() {
        let mut matches = python_regex_scope_matcher();
        assert!(matches("", "anything"));
        assert!(matches(r"机构-(?:12|34)\d+", "前缀机构-123后缀"));
        assert!(!matches(r"^机构-(?:12|34)\d+$", "前缀机构-123后缀"));
        assert!(!matches("[", "anything"));
    }

    #[test]
    fn evaluates_five_segment_tokens_across_current_and_previous_data() {
        let current = dataset(
            vec![record("0001", "1234567", "I1", "月", "2", 8.0)],
            "2026-08-31",
        );
        let previous = dataset(
            vec![record("0001", "1234567", "I1", "月", "2", 10.0)],
            "2026-07-31",
        );
        let rules = [rule(
            "[I1,余额,人民币,月,2]",
            "{I1,余额,人民币,月,2}",
            "left < right",
        )];

        let result =
            run_current_value_check(&current, Some(&previous), &rules, "亿元", |_, _| true);

        assert_eq!(result.rows.len(), 1);
        assert_eq!(result.rows[0].left_value, 8.0);
        assert_eq!(result.rows[0].right_value, 10.0);
        assert_eq!(result.rows[0].difference, -2.0);
        assert_eq!(result.rows[0].difference_percent, Some(20.0));
        assert_eq!(result.rows[0].explanation_required, "核实");
        assert!(result.logs.is_empty());
    }

    #[test]
    fn skips_previous_period_rules_when_previous_dataset_is_absent() {
        let current = dataset(
            vec![record("0001", "1234567", "I1", "月", "2", 8.0)],
            "2026-08-31",
        );
        let rules = [rule("[I1,余额,人民币,月,2]", "{I1,余额,人民币,月,2}", "")];

        let result = run_current_value_check(&current, None, &rules, "亿元", |_, _| true);

        assert!(result.rows.is_empty());
    }

    #[test]
    fn records_formula_errors_and_keeps_the_result_row() {
        let current = dataset(
            vec![record("0001", "1234567", "I1", "月", "2", 8.0)],
            "2026-08-31",
        );
        let rules = [rule(
            "[I1,余额,人民币,月,2]",
            "[I1,余额,人民币,月,2] + 1",
            "left / 0 > right",
        )];

        let result = run_current_value_check(&current, None, &rules, "亿元", |_, _| true);

        assert_eq!(result.rows.len(), 1);
        assert!(result.rows[0]
            .explanation_required
            .contains("校验公式计算有误"));
        assert_eq!(result.logs.len(), 1);
        assert_eq!(result.logs[0].error_type, "校验公式 计算有误");
    }

    #[test]
    fn preserves_python_zero_sentinel_and_omits_both_zero_rows() {
        let current = dataset(
            vec![
                record("0001", "1234567", "I1", "月", "2", 0.0),
                record("0002", "1234567", "I1", "月", "2", 5.0),
            ],
            "2026-08-31",
        );
        let rules = [rule(
            "[I1,余额,人民币,月,2]",
            "[MISSING,余额,人民币,月,2]",
            "",
        )];

        let result = run_current_value_check(&current, None, &rules, "亿元", |_, _| true);

        assert_eq!(result.rows.len(), 2);
        assert_eq!(result.rows[0].left_value, 1e-10);
        assert_eq!(result.rows[0].explanation_required, "一边有，一边无;");
        assert_eq!(result.rows[1].right_value, 0.0);
    }

    #[test]
    fn evaluates_arithmetic_comparisons_and_boolean_functions() {
        assert_eq!(
            evaluate_expression(
                "AND(0.8 * left > right, left - right >= 1)",
                &[("left", 10.0), ("right", 7.0)]
            )
            .unwrap(),
            ExpressionValue::Boolean(true)
        );
        assert!(evaluate_expression("1 / 0", &[]).is_err());
        assert!(evaluate_expression("left > unknown", &[("left", 1.0)]).is_err());
    }

    #[test]
    fn business_expression_supports_configured_abs_and_sum_functions() {
        assert_eq!(
            evaluate_s3_business_expression("AND(ABS(-3)=3,SUM(1,2,3)=6)", 0.01).unwrap(),
            ExpressionValue::Boolean(true)
        );
        let value = evaluate_s3_business_expression("SUM(Thd,ABS(-2))", 0.01)
            .unwrap()
            .as_number();
        assert!((value - 2.01).abs() < 1e-12);
    }

    #[test]
    fn substituted_float_text_matches_python_repr_boundaries() {
        assert_eq!(python_float_repr(1.0), "1.0");
        assert_eq!(python_float_repr(0.0001), "0.0001");
        assert_eq!(python_float_repr(0.00001), "1e-05");
        assert_eq!(python_float_repr(0.0000001), "1e-07");
        assert_eq!(python_float_repr(1e16), "1e+16");
    }

    #[test]
    fn missing_tokens_use_python_zero_literal() {
        let data = InstitutionData {
            values: HashMap::from([(
                "present".to_owned(),
                PairedValue {
                    current: Some(1.0),
                    previous: Some(2.0),
                },
            )]),
        };

        assert_eq!(substitute_tokens("[missing]+{missing}", &data), "0+0");
        assert_eq!(substitute_tokens("[present]+{present}", &data), "1.0+2.0");
    }

    #[test]
    fn reads_s3_cross_period_configuration_and_unit_exceptions() {
        let path = temporary_path("config", "xlsx");
        create_test_config(&path);

        let config = load_current_value_check_config(std::slice::from_ref(&path)).unwrap();

        assert_eq!(config.source_unit, "元");
        assert_eq!(config.target_unit, "亿元");
        assert!(config.exempt_indicators.contains("I0"));
        assert_eq!(config.rules.len(), 1);
        assert_eq!(config.rules[0].institution_pattern, "^0001$");
        assert_eq!(config.rules[0].left_expression, "[I1,余额,人民币,月,2]");
        let _ = fs::remove_file(path);
    }

    #[test]
    fn workbook_cell_text_matches_python_xlsx_and_xls_scalar_strings() {
        assert_eq!(
            workbook_cell_text(&CellValue::Boolean(true), false).unwrap(),
            "True"
        );
        assert_eq!(
            workbook_cell_text(&CellValue::Boolean(false), false).unwrap(),
            "False"
        );
        assert_eq!(
            workbook_cell_text(&CellValue::Boolean(true), true).unwrap(),
            "TRUE"
        );
        assert_eq!(
            workbook_cell_text(&CellValue::Number(12.0), false).unwrap(),
            "12"
        );
        assert_eq!(
            workbook_cell_text(
                &CellValue::DateTimeIso("2026-08-31T00:00:00".to_owned()),
                false,
            )
            .unwrap(),
            "2026-08-31 00:00:00"
        );
        assert_eq!(
            workbook_cell_text(
                &CellValue::DateTimeIso("2026-08-31T00:00:00".to_owned()),
                true,
            )
            .unwrap(),
            "2026-08-31"
        );
        assert_eq!(
            workbook_cell_text(&CellValue::Error("Div0".to_owned()), false).unwrap(),
            "#DIV/0!"
        );
        assert_eq!(
            workbook_cell_text(&CellValue::Error("Div0".to_owned()), true).unwrap(),
            "#ERR7"
        );
    }

    #[test]
    fn xlsx_import_selects_matching_sheet_and_preserves_python_values_and_rows() {
        let path = temporary_path("central-import", "xlsx");
        create_test_central_xlsx(&path);

        let dataset = load_central_workbook_dataset(&path, 1.0, &BTreeSet::new()).unwrap();

        assert_eq!(dataset.records.len(), 1);
        assert_eq!(dataset.records[0].org_name, "True");
        assert_eq!(dataset.records[0].region_name, "False");
        assert_eq!(dataset.issues.len(), 1);
        assert_eq!(dataset.issues[0].row, 4);
        let _ = fs::remove_file(path);
    }

    #[test]
    fn csv_service_writes_python_contract_workbook_and_flow_log() {
        let config_path = temporary_path("service-config", "xlsx");
        let current_path = temporary_path("current", "csv");
        let previous_path = temporary_path("previous", "csv");
        let output_path = temporary_path("output", "xlsx");
        let log_path = temporary_path("logs", "xlsx");
        create_test_config(&config_path);
        let header = EXPECTED_HEADERS.join(",");
        fs::write(
            &current_path,
            format!(
                "{header}\n人民币,2026-08-31,0001,机构,1234567,地区,1,I1,指标,余额,人民币,月,2,200000000\n"
            ),
        )
        .unwrap();
        fs::write(
            &previous_path,
            format!(
                "{header}\n人民币,2026-07-31,0001,机构,1234567,地区,1,I1,指标,余额,人民币,月,2,100000000\n"
            ),
        )
        .unwrap();

        let report = run_current_value_check_files(
            std::slice::from_ref(&current_path),
            std::slice::from_ref(&previous_path),
            std::slice::from_ref(&config_path),
            &output_path,
            Some(&log_path),
            |pattern, value| {
                pattern.is_empty() || value.contains(pattern.trim_matches('^').trim_matches('$'))
            },
        )
        .unwrap();

        assert_eq!(report.status, "completed");
        assert_eq!(report.rows, 1);
        assert_eq!(report.total_rules, 1);
        assert_eq!(report.current_records, 1);
        assert_eq!(report.previous_records, 1);
        let result_sheet = read_value_columns(&output_path, 17).unwrap().remove(0);
        assert_eq!(cell_text(&result_sheet, 1, 1), "数据日期");
        assert_eq!(result_sheet.values.get("J2"), Some(&CellValue::Number(2.0)));
        assert_eq!(result_sheet.values.get("K2"), Some(&CellValue::Number(1.0)));
        assert_eq!(
            result_sheet.values.get("O2"),
            Some(&CellValue::Text("请核实".to_owned()))
        );
        let log_sheet = read_value_columns(&log_path, 6).unwrap().remove(0);
        assert_eq!(log_sheet.name, "运行日志");
        assert_eq!(log_sheet.max_row, 4);

        for path in [
            config_path,
            current_path,
            previous_path,
            output_path,
            log_path,
        ] {
            let _ = fs::remove_file(path);
        }
    }

    #[test]
    fn output_paths_are_preflighted_before_any_workbook_is_written() {
        let output_path = temporary_path("preflight-output", "xlsx");
        let same_log_path = output_path.clone();
        let missing_inputs = [PathBuf::from("not-read.csv")];

        let same_path_error = run_current_value_check_files(
            &missing_inputs,
            &[],
            &[],
            &output_path,
            Some(&same_log_path),
            |_, _| true,
        )
        .expect_err("result and log must not share a destination");
        assert!(format!("{same_path_error:#}").contains("不能使用同一输出路径"));
        assert!(!output_path.exists());

        let existing_log_path = temporary_path("preflight-existing-log", "xlsx");
        fs::write(&existing_log_path, b"keep existing log").unwrap();
        let protected_output_path = temporary_path("preflight-protected-output", "xlsx");
        let existing_log_error = run_current_value_check_files(
            &missing_inputs,
            &[],
            &[],
            &protected_output_path,
            Some(&existing_log_path),
            |_, _| true,
        )
        .expect_err("existing log must reject the run before writing its result");
        assert!(format!("{existing_log_error:#}").contains("拒绝覆盖已有本期数值核对运行日志"));
        assert!(!protected_output_path.exists());
        assert_eq!(fs::read(&existing_log_path).unwrap(), b"keep existing log");
        let _ = fs::remove_file(existing_log_path);
    }
}
