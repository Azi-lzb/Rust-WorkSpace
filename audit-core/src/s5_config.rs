//! S5 configuration workbook loading.
//!
//! Faithful Rust port of the source project's
//! `s1audit_configured.workbook_config.load_configuration`: it reads the audit
//! configuration workbook (rule sheet plus the optional metric, scope,
//! external-metric, contract, and version sheets), validates every rule row
//! with the same order of checks as the Python implementation (including the
//! DSL segment from [`crate::s5_config_dsl`]), derives metric specifications
//! from `V_CUR`/`V_PRE` calls when no metric catalog exists, and assembles the
//! [`Configuration`] value. Raw error messages mirror the Python source so
//! real-configuration parity can compare them exactly.
//!
//! Known reading delta: openpyxl converts date-formatted numeric cells into
//! Python dates; this port reads the stored value text. Real configurations
//! store their fields as text, so the parity run covers the shipped format.

use crate::s5_config_dsl::{
    calls_in, parse_dictionary, parse_expression, static_kind, validate_dictionary_dependencies,
    validate_static_types, variables_in, CallNode, DslError, DslErrorKind, Node,
};
use excel_reader::{read_workbook, SheetDump};
use serde::Serialize;
use sha2::{Digest, Sha256};
use std::{collections::BTreeMap, fs, path::Path};

const METRIC_HEADERS: [&str; 15] = [
    "业务类型",
    "工作表名称",
    "行路径",
    "列路径",
    "行表头标题",
    "数据类型",
    "指标类型",
    "来源单位",
    "标准单位",
    "存储尺度",
    "允许空值",
    "精度",
    "口径依据",
    "确认状态",
    "布局版本",
];
const EXTERNAL_METRIC_HEADERS: [&str; 11] = [
    "数据源ID",
    "指标表头",
    "数据类型",
    "指标类型",
    "来源单位",
    "标准单位",
    "存储尺度",
    "允许空值",
    "精度",
    "口径依据",
    "确认状态",
];
const SOURCE_HEADERS: [&str; 22] = [
    "数据源ID",
    "来源类型",
    "默认Sheet",
    "机构Sheet",
    "机构代码标签",
    "机构名称标签",
    "期间标签",
    "映射Sheet",
    "映射表头行",
    "映射数据起始行",
    "映射机构字段",
    "映射目标字段",
    "数据表头行",
    "数据起始行",
    "数据键列",
    "指标表头行",
    "指标起始列",
    "报表项目标签",
    "期望期间",
    "契约版本",
    "期间必须",
    "机构必须",
];
const VERSION_HEADERS: [&str; 6] = [
    "配置版本号",
    "状态",
    "发布日期",
    "生效起始日",
    "生效截止日",
    "发布说明",
];
const SCOPE_HEADERS: [&str; 6] = [
    "适用范围ID",
    "业务类型",
    "工作表名称",
    "行路径前缀",
    "范围说明",
    "确认状态",
];

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct LoadError {
    pub kind: DslErrorKind,
    pub message: String,
}

impl LoadError {
    pub fn configuration(message: impl Into<String>) -> Self {
        Self {
            kind: DslErrorKind::Configuration,
            message: message.into(),
        }
    }
}

impl std::fmt::Display for LoadError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(formatter, "{}", self.message)
    }
}

impl std::error::Error for LoadError {}

impl From<DslError> for LoadError {
    fn from(error: DslError) -> Self {
        Self {
            kind: error.kind,
            message: error.message,
        }
    }
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct RuleDefinition {
    pub rule_id: String,
    pub revision: i64,
    pub enabled: bool,
    pub business: String,
    pub default_source: String,
    pub default_sheet: String,
    pub name: String,
    pub severity: String,
    pub check_type: String,
    pub metric: String,
    pub audit_expression: String,
    pub description: String,
    pub audit_expression_inverted: bool,
    pub current_expression: String,
    pub comparison_expression: String,
    pub difference_expression: String,
    pub value_dictionary: String,
    pub effective_from: String,
    pub effective_to: String,
    pub applicable_expression: String,
    pub applicability_scope: String,
    pub numeric_mode: String,
    pub formula_cell: String,
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct MetricSpec {
    pub business: String,
    pub sheet: String,
    pub row_path: Vec<String>,
    pub column_path: Vec<String>,
    pub row_header_titles: Vec<String>,
    pub data_type: String,
    pub metric_type: String,
    pub source_unit: String,
    pub standard_unit: String,
    pub storage_scale: String,
    pub allow_blank: bool,
    pub precision: i64,
    pub evidence: String,
    pub confirmed: bool,
    pub layout_version: String,
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct ExternalMetricSpec {
    pub source_id: String,
    pub metric_header: String,
    pub data_type: String,
    pub metric_type: String,
    pub source_unit: String,
    pub standard_unit: String,
    pub storage_scale: String,
    pub allow_blank: bool,
    pub precision: i64,
    pub evidence: String,
    pub confirmed: bool,
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct SourceContract {
    pub source_id: String,
    pub source_type: String,
    pub default_sheet: String,
    pub identity_sheet: String,
    pub institution_code_label: String,
    pub institution_name_label: String,
    pub period_label: String,
    pub mapping_sheet: String,
    pub mapping_header_row: i64,
    pub mapping_data_start_row: i64,
    pub mapping_identity_field: String,
    pub mapping_target_field: String,
    pub data_header_row: i64,
    pub data_start_row: i64,
    pub data_key_column: i64,
    pub metric_header_row: i64,
    pub metric_first_column: i64,
    pub report_item_label: String,
    pub expected_period: String,
    pub version: String,
    pub period_required: bool,
    pub institution_required: bool,
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct ApplicabilityScope {
    pub scope_id: String,
    pub business: String,
    pub sheet: String,
    pub row_path_prefix: Vec<String>,
    pub description: String,
    pub confirmed: bool,
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct Configuration {
    pub rules: Vec<RuleDefinition>,
    pub metrics: Vec<MetricSpec>,
    pub external_metrics: Vec<ExternalMetricSpec>,
    pub source_contracts: Vec<SourceContract>,
    pub applicability_scopes: Vec<ApplicabilityScope>,
    pub version: String,
    pub version_status: String,
    pub effective_from: String,
    pub effective_to: String,
    pub release_date: String,
    pub release_note: String,
    pub sha256: String,
}

/// One configuration worksheet: header names in column order plus the
/// non-blank data rows with all values converted through the Python `_text`
/// semantics.
struct ConfigSheet {
    columns: Vec<String>,
    rows: Vec<ConfigRow>,
}

struct ConfigRow {
    row_num: usize,
    values: BTreeMap<String, String>,
}

impl ConfigRow {
    fn get(&self, header: &str) -> String {
        self.values.get(header).cloned().unwrap_or_default()
    }
}

fn column_letter(mut column: usize) -> String {
    let mut letters = Vec::new();
    while column > 0 {
        let digit = (column - 1) % 26;
        letters.push((b'A' + digit as u8) as char);
        column = (column - 1) / 26;
    }
    letters.into_iter().rev().collect()
}

fn cell_text(value: &serde_json::Value) -> String {
    let kind = value.get("type").and_then(serde_json::Value::as_str);
    let raw = value.get("value");
    match kind {
        Some("string") => raw
            .and_then(serde_json::Value::as_str)
            .unwrap_or_default()
            .trim()
            .to_owned(),
        Some("number") => {
            let number = raw.and_then(serde_json::Value::as_f64).unwrap_or(0.0);
            if number.fract() == 0.0 && number.abs() < 1e15 {
                format!("{}", number as i64)
            } else {
                format!("{number}")
            }
        }
        Some("boolean") => {
            if raw.and_then(serde_json::Value::as_bool) == Some(true) {
                "True".to_owned()
            } else {
                "False".to_owned()
            }
        }
        _ => String::new(),
    }
}

fn parse_address(address: &str) -> Option<(usize, usize)> {
    let mut row_part = String::new();
    let mut column = 0usize;
    for character in address.chars() {
        if character.is_ascii_alphabetic() {
            column = column * 26 + (character.to_ascii_uppercase() as usize - 'A' as usize + 1);
        } else if character.is_ascii_digit() {
            row_part.push(character);
        } else {
            return None;
        }
    }
    let row = row_part.parse::<usize>().ok()?;
    if row == 0 || column == 0 {
        return None;
    }
    Some((row, column))
}

fn header_row_columns(sheet: &SheetDump) -> Vec<String> {
    let mut max_column = 0usize;
    for address in sheet.values.keys().chain(sheet.formulas.keys()) {
        if let Some((row, column)) = parse_address(address) {
            if row == 1 && column > max_column {
                max_column = column;
            }
        }
    }
    let mut headers: Vec<String> = (1..=max_column)
        .map(|column| {
            sheet
                .values
                .get(&format!("{}1", column_letter(column)))
                .map(cell_text)
                .unwrap_or_default()
        })
        .collect();
    while headers.last().is_some_and(|header| header.is_empty()) {
        headers.pop();
    }
    headers
}

/// Mirror of the Python `_sheet_rows`: header validation, blank-row skip, and
/// the per-cell formula rejection.
fn read_sheet(
    workbook_sheets: &[SheetDump],
    sheet_name: &str,
    required_headers: &[&str],
    required: bool,
) -> Result<Option<ConfigSheet>, LoadError> {
    let Some(sheet) = workbook_sheets
        .iter()
        .find(|sheet| sheet.name == sheet_name)
    else {
        if required {
            return Err(LoadError::configuration(format!(
                "配置工作簿缺少工作表：{sheet_name}"
            )));
        }
        return Ok(None);
    };
    let headers = header_row_columns(sheet);
    if headers.is_empty() {
        return Err(LoadError::configuration("缺少表头，请在第1行填写字段名称"));
    }
    let mut seen = std::collections::BTreeSet::new();
    for header in &headers {
        if !seen.insert(header.as_str()) {
            let repeated = if header.is_empty() {
                "空白表头".to_owned()
            } else {
                header.clone()
            };
            return Err(LoadError::configuration(format!(
                "存在重复表头：{repeated}，请保留唯一字段名称"
            )));
        }
    }
    let missing: Vec<&str> = required_headers
        .iter()
        .copied()
        .filter(|header| !headers.iter().any(|existing| existing == header))
        .collect();
    if !missing.is_empty() {
        return Err(LoadError::configuration(format!(
            "第1行缺少字段：{}；缺失字段尚无列号，请补充表头",
            missing.join(", ")
        )));
    }
    let mut max_row = 0usize;
    for address in sheet.values.keys().chain(sheet.formulas.keys()) {
        if let Some((row, _)) = parse_address(address) {
            if row > max_row {
                max_row = row;
            }
        }
    }
    let mut rows = Vec::new();
    for row_num in 2..=max_row {
        let mut values = BTreeMap::new();
        let mut has_content = false;
        for (index, header) in headers.iter().enumerate() {
            let address = format!("{}{row_num}", column_letter(index + 1));
            if sheet.formulas.contains_key(&address) {
                return Err(LoadError::configuration(
                    "不允许使用 Excel 公式，请将审核表达式按文本保存，开头不要加 =",
                ));
            }
            let text = sheet
                .values
                .get(&address)
                .map(cell_text)
                .unwrap_or_default();
            if !text.is_empty() {
                has_content = true;
            }
            values.insert(header.clone(), text);
        }
        if !has_content {
            continue;
        }
        rows.push(ConfigRow { row_num, values });
    }
    Ok(Some(ConfigSheet {
        columns: headers,
        rows,
    }))
}

fn flag(value: &str, field: &str, default: Option<bool>) -> Result<bool, LoadError> {
    let text = value.trim().to_ascii_lowercase();
    if text.is_empty() {
        if let Some(default) = default {
            return Ok(default);
        }
    }
    match text.as_str() {
        "是" | "是的" | "yes" | "y" | "true" | "1" | "启用" | "已确认" => Ok(true),
        "否" | "不是" | "no" | "n" | "false" | "0" | "禁用" | "未确认" => Ok(false),
        _ => Err(LoadError::configuration(format!("{field} 必须填写 是/否"))),
    }
}

fn integer(value: &str, field: &str, default: Option<i64>) -> Result<i64, LoadError> {
    let text = value.trim();
    if text.is_empty() {
        if let Some(default) = default {
            return Ok(default);
        }
        return Err(LoadError::configuration(format!("{field} 不能为空")));
    }
    if let Ok(number) = text.parse::<i64>() {
        return Ok(number);
    }
    if let Ok(number) = text.parse::<f64>() {
        let truncated = number as i64;
        if number == truncated as f64 {
            return Ok(truncated);
        }
    }
    Err(LoadError::configuration(format!("{field} 必须是整数")))
}

fn path_parts(value: &str, field: &str) -> Result<Vec<String>, LoadError> {
    let text = value.trim();
    if text.is_empty() {
        return Err(LoadError::configuration(format!("{field} 不能为空")));
    }
    let parts: Vec<String> = text
        .split("\r\n")
        .flat_map(|part| part.split(['\n', '\r']))
        .map(|part| part.trim().to_owned())
        .filter(|part| !part.is_empty())
        .collect();
    if parts.is_empty() {
        return Err(LoadError::configuration(format!("{field} 不能为空")));
    }
    Ok(parts)
}

fn parse_field_date(value: &str, field: &str) -> Result<(i32, i32, i32), LoadError> {
    let digits = value.chars().all(|c| c.is_ascii_digit());
    let (year, month, day) = if digits && value.len() == 8 {
        (
            value[0..4].parse::<i32>().unwrap_or(0),
            value[4..6].parse::<i32>().unwrap_or(0),
            value[6..8].parse::<i32>().unwrap_or(0),
        )
    } else if value.len() == 10
        && value.as_bytes().get(4) == Some(&b'-')
        && value.as_bytes().get(7) == Some(&b'-')
    {
        (
            value[0..4].parse::<i32>().unwrap_or(0),
            value[5..7].parse::<i32>().unwrap_or(0),
            value[8..10].parse::<i32>().unwrap_or(0),
        )
    } else if digits && value.len() == 6 {
        (
            value[0..4].parse::<i32>().unwrap_or(0),
            value[4..6].parse::<i32>().unwrap_or(0),
            1,
        )
    } else if value.len() == 7 && value.as_bytes().get(4) == Some(&b'-') {
        (
            value[0..4].parse::<i32>().unwrap_or(0),
            value[5..7].parse::<i32>().unwrap_or(0),
            1,
        )
    } else {
        return Err(LoadError::configuration(format!(
            "{field}日期格式必须为 YYYYMMDD、YYYY-MM-DD、YYYYMM 或 YYYY-MM"
        )));
    };
    let leap = year % 4 == 0 && (year % 100 != 0 || year % 400 == 0);
    let month_days = match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if leap => 29,
        2 => 28,
        _ => {
            return Err(LoadError::configuration(format!(
                "{field}日期无效：{value}"
            )))
        }
    };
    if day < 1 || day > month_days {
        return Err(LoadError::configuration(format!(
            "{field}日期无效：{value}"
        )));
    }
    Ok((year, month, day))
}

fn validate_date_range(start: &str, end: &str, field: &str) -> Result<(), LoadError> {
    let start_date = if start.trim().is_empty() {
        None
    } else {
        Some(parse_field_date(start, field)?)
    };
    let end_date = if end.trim().is_empty() {
        None
    } else {
        Some(parse_field_date(end, field)?)
    };
    if let (Some(start_date), Some(end_date)) = (start_date, end_date) {
        if start_date > end_date {
            return Err(LoadError::configuration(format!(
                "{field}的开始日期晚于结束日期"
            )));
        }
    }
    Ok(())
}

/// Mirror of `layout.split_path` for string inputs: split on `@` and trim each
/// part, keeping empty parts.
fn split_at_path(value: &str) -> Vec<String> {
    value
        .split('@')
        .map(|part| part.trim().to_owned())
        .collect()
}

fn literal_string_value(node: &Node) -> Option<&str> {
    match node {
        Node::Literal(crate::s5_config_dsl::LiteralNode {
            value: crate::s5_config_dsl::LiteralValue::String(value),
            kind: "string",
        }) => Some(value),
        _ => None,
    }
}

fn literal_path(node: &Node) -> Vec<String> {
    match node {
        Node::Literal(_) => literal_string_value(node)
            .map(split_at_path)
            .unwrap_or_default(),
        Node::Tuple(tuple) => tuple
            .values
            .iter()
            .map(|value| value.trim().to_owned())
            .filter(|value| !value.is_empty())
            .collect(),
        _ => Vec::new(),
    }
}

/// Sheet and path arguments of a `V_CUR`/`V_PRE` call: three-argument calls
/// carry an explicit sheet string first, two-argument calls use `sheet=`.
fn call_sheet_and_paths(call: &CallNode) -> (String, Vec<String>, Vec<String>) {
    if call.arguments.len() == 3 {
        (
            literal_string_value(&call.arguments[0])
                .map(str::trim)
                .unwrap_or_default()
                .to_owned(),
            literal_path(&call.arguments[1]),
            literal_path(&call.arguments[2]),
        )
    } else if call.arguments.len() == 2 {
        let sheet = call
            .keywords
            .iter()
            .find(|(name, _)| name == "sheet")
            .map(|(_, value)| value.trim().to_owned())
            .unwrap_or_default();
        (
            sheet,
            literal_path(&call.arguments[0]),
            literal_path(&call.arguments[1]),
        )
    } else {
        (String::new(), Vec::new(), Vec::new())
    }
}

fn canonical_metric_type(value: &str) -> String {
    let text = value.trim().to_lowercase();
    match text.as_str() {
        "" | "金额" | "余额" | "发生额" => "amount".to_owned(),
        "利率" | "比例" => "rate".to_owned(),
        "户数" | "笔数" | "计数" => "count".to_owned(),
        "普通数值" => "number".to_owned(),
        "文本" => "text".to_owned(),
        _ => text,
    }
}

fn canonical_data_type(value: &str) -> String {
    let text = value.trim().to_lowercase();
    match text.as_str() {
        "" | "数值" | "数值型" | "decimal" => "number".to_owned(),
        "整数" => "integer".to_owned(),
        "文本" | "字符串" | "string" => "text".to_owned(),
        _ => text,
    }
}

fn canonical_storage_scale(value: &str) -> String {
    let text = value.trim().to_lowercase();
    match text.as_str() {
        "" | "原值" => "native".to_owned(),
        "比例存储" | "小数比例" => "fraction".to_owned(),
        "百分比存储" | "百分数存储" => "percent".to_owned(),
        _ => text,
    }
}

fn canonical_unit(value: &str, metric_type: &str) -> String {
    if metric_type == "rate" && (value == "%" || value == "百分点") {
        return "百分比".to_owned();
    }
    value.to_owned()
}

fn metric_profile(
    labels: &[&str],
    source_unit: &str,
    standard_unit: &str,
    metric_type: &str,
) -> (String, String, String, String) {
    let label = labels.join(" ").to_lowercase();
    let mut source_unit = source_unit.to_owned();
    let kind = if metric_type.is_empty() {
        String::new()
    } else {
        canonical_metric_type(metric_type)
    };
    let kind = if kind.is_empty() {
        if source_unit == "户" || source_unit == "笔" {
            "count".to_owned()
        } else if ["百分比", "比例", "百分点", "%"].contains(&source_unit.as_str())
            || ["利率", "利息率", "百分比", "占比", "%"]
                .iter()
                .any(|token| label.contains(token))
        {
            "rate".to_owned()
        } else if ["户数", "贷款户", "客户户数", "（户）", "(户)"]
            .iter()
            .any(|token| label.contains(token))
        {
            if source_unit.is_empty() {
                source_unit = "户".to_owned();
            }
            "count".to_owned()
        } else if ["笔数", "业务笔数", "发生笔数", "（笔）", "(笔)"]
            .iter()
            .any(|token| label.contains(token))
        {
            if source_unit.is_empty() {
                source_unit = "笔".to_owned();
            }
            "count".to_owned()
        } else {
            "amount".to_owned()
        }
    } else {
        kind
    };
    if kind == "amount" {
        let source = if source_unit.is_empty() {
            "万元".to_owned()
        } else {
            source_unit
        };
        let standard = if standard_unit.is_empty() {
            "万元".to_owned()
        } else {
            standard_unit.to_owned()
        };
        (kind, source, standard, "native".to_owned())
    } else if kind == "rate" {
        let source = canonical_unit(
            &if source_unit.is_empty() {
                "百分比".to_owned()
            } else {
                source_unit
            },
            "rate",
        );
        let standard = canonical_unit(
            &if standard_unit.is_empty() {
                source.clone()
            } else {
                standard_unit.to_owned()
            },
            "rate",
        );
        (kind, source, standard, "native".to_owned())
    } else if kind == "count" {
        if source_unit.is_empty() {
            source_unit = if label.contains("户数") || label.contains("贷款户") {
                "户".to_owned()
            } else {
                "笔".to_owned()
            };
        }
        let standard = if standard_unit.is_empty() {
            source_unit.clone()
        } else {
            standard_unit.to_owned()
        };
        (kind, source_unit, standard, "native".to_owned())
    } else {
        let source = if source_unit.is_empty() {
            "万元".to_owned()
        } else {
            source_unit
        };
        let standard = if standard_unit.is_empty() {
            source.clone()
        } else {
            standard_unit.to_owned()
        };
        (kind, source, standard, "native".to_owned())
    }
}

fn validate_metric_units(
    metric_type: &str,
    source_unit: &str,
    standard_unit: &str,
    storage_scale: &str,
) -> Result<(), LoadError> {
    if !matches!(metric_type, "amount" | "rate" | "count" | "number" | "text") {
        return Err(LoadError::configuration(format!(
            "指标类型不支持：{metric_type}"
        )));
    }
    if metric_type == "amount" {
        let allowed = ["元", "万元", "亿元"];
        if !allowed.contains(&source_unit) || !allowed.contains(&standard_unit) {
            return Err(LoadError::configuration("金额指标单位必须是元/万元/亿元"));
        }
    } else if metric_type == "rate" {
        let allowed = ["百分比", "%", "百分点", "比例"];
        if !allowed.contains(&source_unit) || !allowed.contains(&standard_unit) {
            return Err(LoadError::configuration("利率指标必须明确比例单位/尺度"));
        }
    } else if metric_type == "count" {
        if (source_unit != "户" && source_unit != "笔") || standard_unit != source_unit {
            return Err(LoadError::configuration(
                "计数指标单位必须是户或笔且保持一致",
            ));
        }
    } else if metric_type == "number" && source_unit != standard_unit {
        return Err(LoadError::configuration(
            "普通数值指标来源单位与标准单位必须相同",
        ));
    }
    if metric_type == "rate" && !matches!(storage_scale, "native" | "fraction" | "percent") {
        return Err(LoadError::configuration(format!(
            "比例存储尺度不支持：{storage_scale}"
        )));
    }
    Ok(())
}

fn validate_data_type(data_type: &str, metric_type: &str, row_num: usize) -> Result<(), LoadError> {
    if !matches!(
        data_type,
        "number" | "numeric" | "decimal" | "integer" | "text" | "string"
    ) {
        return Err(LoadError::configuration(format!(
            "指标目录第 {row_num} 行数据类型不支持：{data_type}"
        )));
    }
    if metric_type == "text" && !matches!(data_type, "text" | "string") {
        return Err(LoadError::configuration(format!(
            "文本指标目录第 {row_num} 行的数据类型必须是文本"
        )));
    }
    if metric_type != "text" && matches!(data_type, "text" | "string") {
        return Err(LoadError::configuration(format!(
            "数值指标目录第 {row_num} 行的数据类型不能是文本"
        )));
    }
    Ok(())
}

fn validate_precision(precision: i64, row_num: usize) -> Result<(), LoadError> {
    if !(0..=12).contains(&precision) {
        return Err(LoadError::configuration(format!(
            "指标目录第 {row_num} 行精度必须介于 0 到 12"
        )));
    }
    Ok(())
}

/// Python `repr` of a tuple of strings, used verbatim inside two error
/// messages.
fn py_tuple_repr(parts: &[String]) -> String {
    let quoted: Vec<String> = parts
        .iter()
        .map(|part| format!("'{}'", part.replace('\\', "\\\\").replace('\'', "\\'")))
        .collect();
    let joined = quoted.join(", ");
    if parts.len() == 1 {
        format!("({joined},)")
    } else {
        format!("({joined})")
    }
}

fn unit_suffix_from_header(header: &str) -> String {
    let characters: Vec<char> = header.chars().collect();
    let open = characters.iter().position(|c| *c == '（' || *c == '(');
    let close = characters.iter().position(|c| *c == '）' || *c == ')');
    if let (Some(open), Some(close)) = (open, close) {
        if open < close {
            let inner: String = characters[open + 1..close].iter().collect();
            let inner = inner.trim();
            if matches!(
                inner,
                "元" | "万元" | "亿元" | "百分比" | "百分点" | "户" | "笔"
            ) {
                return inner.to_owned();
            }
        }
    }
    String::new()
}

fn template_metric_from_call(
    call: &CallNode,
    rule_id: &str,
    rule_business: &str,
    rule_default_sheet: &str,
) -> Result<MetricSpec, LoadError> {
    if call.arguments.len() != 2 && call.arguments.len() != 3 {
        return Err(LoadError::configuration(format!(
            "{} 必须传入行路径、列路径，可选工作表名",
            call.name
        )));
    }
    let (sheet_from_call, row_path, column_path) = call_sheet_and_paths(call);
    let mut sheet = sheet_from_call;
    if sheet.is_empty() {
        sheet = rule_default_sheet.to_owned();
    }
    if sheet.is_empty() {
        return Err(LoadError::configuration(format!(
            "规则 {rule_id} 的 {} 缺少工作表名，请使用 V_CUR(工作表, 行路径, 列路径)",
            call.name
        )));
    }
    if row_path.is_empty()
        || column_path.is_empty()
        || row_path.iter().any(|part| part.is_empty())
        || column_path.iter().any(|part| part.is_empty())
    {
        return Err(LoadError::configuration(format!(
            "规则 {rule_id} 的 {} 行路径和列路径必须是字符串或字符串元组",
            call.name
        )));
    }
    let keyword = |name: &str| {
        call.keywords
            .iter()
            .find(|(key, _)| key == name)
            .map(|(_, value)| value.clone())
            .unwrap_or_default()
    };
    let unit = if keyword("unit").is_empty() {
        keyword("source_unit")
    } else {
        keyword("unit")
    };
    let label_parts: Vec<&str> = row_path
        .iter()
        .chain(column_path.iter())
        .map(String::as_str)
        .collect();
    let (kind, source_unit, standard_unit, storage_scale) = metric_profile(
        &label_parts,
        &unit,
        &keyword("standard_unit"),
        &keyword("metric_type"),
    );
    Ok(MetricSpec {
        business: rule_business.to_owned(),
        sheet,
        row_path,
        column_path,
        row_header_titles: Vec::new(),
        data_type: "number".to_owned(),
        metric_type: kind,
        source_unit,
        standard_unit,
        storage_scale,
        allow_blank: true,
        precision: 2,
        evidence: String::new(),
        confirmed: true,
        layout_version: String::new(),
    })
}

fn external_metric_from_call(
    call: &CallNode,
    rule_id: &str,
) -> Result<ExternalMetricSpec, LoadError> {
    let Some(header) = call.arguments.first() else {
        return Err(LoadError::configuration(format!(
            "规则 {rule_id} 的 V_EXT 指标名必须是字符串"
        )));
    };
    let Some(metric_header) = literal_string_value(header) else {
        return Err(LoadError::configuration(format!(
            "规则 {rule_id} 的 V_EXT 指标名必须是字符串"
        )));
    };
    let metric_header = metric_header.trim();
    if metric_header.is_empty() {
        return Err(LoadError::configuration(format!(
            "规则 {rule_id} 的 V_EXT 指标名不能为空"
        )));
    }
    let keyword = |name: &str| {
        call.keywords
            .iter()
            .find(|(key, _)| key == name)
            .map(|(_, value)| value.clone())
            .unwrap_or_default()
    };
    let mut unit = keyword("unit");
    if unit.is_empty() {
        unit = keyword("source_unit");
    }
    if unit.is_empty() {
        unit = unit_suffix_from_header(metric_header);
    }
    let (kind, source_unit, standard_unit, storage_scale) = metric_profile(
        &[metric_header],
        &unit,
        &keyword("standard_unit"),
        &keyword("metric_type"),
    );
    let source_id = if keyword("source").is_empty() {
        "EXT".to_owned()
    } else {
        keyword("source")
    };
    Ok(ExternalMetricSpec {
        source_id,
        metric_header: metric_header.to_owned(),
        data_type: "number".to_owned(),
        metric_type: kind,
        source_unit,
        standard_unit,
        storage_scale,
        allow_blank: true,
        precision: 2,
        evidence: String::new(),
        confirmed: true,
    })
}

fn description_variables(text: &str) -> Result<std::collections::BTreeSet<String>, LoadError> {
    let characters: Vec<char> = text.chars().collect();
    let mut found = std::collections::BTreeSet::new();
    let mut index = 0usize;
    while index < characters.len() {
        if characters[index] == '{' && characters.get(index + 1) == Some(&'{') {
            index += 2;
            continue;
        }
        if characters[index] == '}' && characters.get(index + 1) == Some(&'}') {
            index += 2;
            continue;
        }
        if characters[index] == '{' {
            let end = characters[index + 1..]
                .iter()
                .position(|c| *c == '}')
                .map(|offset| offset + index + 1);
            let Some(end) = end else {
                return Err(LoadError::configuration(
                    "规则描述中的变量占位符缺少结束大括号",
                ));
            };
            let name: String = characters[index + 1..end].iter().collect();
            let valid = !name.is_empty()
                && name.chars().all(|c| {
                    c == '_' || c.is_alphanumeric() || ('\u{3400}'..='\u{9fff}').contains(&c)
                });
            if !valid {
                return Err(LoadError::configuration("规则描述中的变量占位符格式错误"));
            }
            found.insert(name);
            index = end + 1;
            continue;
        }
        if characters[index] == '}' {
            return Err(LoadError::configuration("规则描述中的右大括号需要写成 }}"));
        }
        index += 1;
    }
    Ok(found)
}

fn infer_rule_business(
    rule_num: usize,
    calls: &[&CallNode],
    metric_catalog: &[(String, String, Vec<String>, Vec<String>)],
    scopes: &[ApplicabilityScope],
    applicability_scope: &str,
) -> Result<String, LoadError> {
    if !applicability_scope.is_empty() {
        if let Some(scope) = scopes
            .iter()
            .find(|scope| scope.scope_id == applicability_scope)
        {
            return Ok(scope.business.clone());
        }
    }
    let mut matching_business_sets: Vec<std::collections::BTreeSet<String>> = Vec::new();
    for call in calls {
        if call.name != "V_CUR" && call.name != "V_PRE" {
            continue;
        }
        if call.arguments.len() != 2 && call.arguments.len() != 3 {
            continue;
        }
        let (sheet, row_path, column_path) = call_sheet_and_paths(call);
        if row_path == vec!["@当前行".to_owned()] {
            continue;
        }
        if row_path.is_empty() || column_path.is_empty() {
            continue;
        }
        let matches: std::collections::BTreeSet<String> = metric_catalog
            .iter()
            .filter(|(_, metric_sheet, metric_row_path, metric_column_path)| {
                metric_row_path == &row_path
                    && metric_column_path == &column_path
                    && (sheet.is_empty() || metric_sheet == &sheet)
            })
            .map(|(business, _, _, _)| business.clone())
            .collect();
        if !matches.is_empty() {
            matching_business_sets.push(matches);
        }
    }
    if !matching_business_sets.is_empty() {
        let common = matching_business_sets
            .iter()
            .map(|set| set.clone())
            .reduce(|left, right| left.intersection(&right).cloned().collect())
            .unwrap_or_default();
        if common.len() == 1 {
            return Ok(common.into_iter().next().unwrap());
        }
        if common.len() > 1 {
            return Err(LoadError::configuration(format!(
                "审核规则第 {rule_num} 行的业务类型无法由指标路径唯一推断，请补充业务类型"
            )));
        }
        return Err(LoadError::configuration(format!(
            "审核规则第 {rule_num} 行引用的指标跨越多个业务，请补充业务类型"
        )));
    }
    let businesses: std::collections::BTreeSet<String> = metric_catalog
        .iter()
        .map(|(business, _, _, _)| business.clone())
        .filter(|business| !business.is_empty())
        .collect();
    if businesses.len() == 1 {
        return Ok(businesses.into_iter().next().unwrap());
    }
    Err(LoadError::configuration(format!(
        "审核规则第 {rule_num} 行缺少业务类型，且无法从指标路径唯一推断"
    )))
}

fn call_argument_text(call: &CallNode, index: usize) -> String {
    call.arguments
        .get(index)
        .and_then(literal_string_value)
        .unwrap_or_default()
        .to_owned()
}

fn is_current_row_call(call: &CallNode) -> bool {
    if call.name != "V_CUR" && call.name != "V_PRE" {
        return false;
    }
    if call.arguments.len() == 2 {
        return call_argument_text(call, 0) == "@当前行";
    }
    if call.arguments.len() == 3 {
        return call_argument_text(call, 1) == "@当前行";
    }
    false
}

fn required_text(row: &ConfigRow, header: &str) -> Result<String, LoadError> {
    let value = row.get(header);
    if value.is_empty() {
        return Err(LoadError::configuration(format!("{header} 不能为空")));
    }
    Ok(value)
}

fn validate_rule_dates(row: &ConfigRow) -> Result<(), LoadError> {
    let start = row.get("生效起始日");
    let end = row.get("生效截止日");
    if !start.is_empty() {
        validate_date_range(&start, &start, "生效起始日")?;
    }
    if !end.is_empty() {
        validate_date_range(&end, &end, "生效截止日")?;
    }
    validate_date_range(&start, &end, "生效期")
}

/// Load and validate one S5 audit configuration workbook.
pub fn load_configuration(
    path: &Path,
    include_disabled: bool,
    business: &str,
) -> Result<Configuration, LoadError> {
    if !path.is_file() {
        return Err(LoadError::configuration(format!(
            "配置工作簿不存在：{}",
            path.display()
        )));
    }
    let config_bytes =
        fs::read(path).map_err(|_| LoadError::configuration("配置工作簿无法读取：文件无法打开"))?;
    let mut hasher = Sha256::new();
    hasher.update(&config_bytes);
    let config_hash = format!("{:x}", hasher.finalize());

    let dump = read_workbook(path)
        .map_err(|error| LoadError::configuration(format!("配置工作簿无法读取：{error}")))?;
    let sheets = &dump.sheets;

    let Some(rules_sheet) = read_sheet(sheets, "审核规则", &[], true)? else {
        return Err(LoadError::configuration("配置工作簿缺少工作表：审核规则"));
    };
    if rules_sheet.rows.is_empty() {
        return Err(LoadError::configuration("审核规则工作表没有规则数据"));
    }
    let rule_headers = &rules_sheet.columns;
    let Some(id_header) = ["规则ID", "规则编号", "规则序号"]
        .into_iter()
        .find(|header| rule_headers.iter().any(|existing| existing == header))
    else {
        return Err(LoadError::configuration(
            "第1行缺少字段：规则ID（也可使用规则编号或规则序号）；请补充表头",
        ));
    };
    let required_rule_headers = [
        "是否启用",
        "规则名称",
        "审核级别",
        "校验类型",
        "校验指标",
        "审核表达式",
        "规则描述",
    ];
    let mut missing_rule_headers = required_rule_headers
        .iter()
        .copied()
        .filter(|header| !rule_headers.iter().any(|existing| existing == header))
        .collect::<Vec<_>>();
    missing_rule_headers.sort();
    if !missing_rule_headers.is_empty() {
        return Err(LoadError::configuration(format!(
            "第1行缺少字段：{}；缺失字段尚无列号，请补充表头",
            missing_rule_headers.join(", ")
        )));
    }

    let requested_business = business.trim().to_owned();
    let mut rule_rows: Vec<ConfigRow> = Vec::new();
    for mut row in rules_sheet.rows {
        if row.get("规则ID").is_empty() {
            let fallback = row.get(id_header);
            row.values.insert("规则ID".to_owned(), fallback);
        }
        for (alias, preferred) in [("业务类型", "业务名称"), ("默认Sheet", "工作表名称")]
        {
            if rule_headers.iter().any(|header| header == preferred) {
                let value = row.get(preferred);
                row.values.entry(alias.to_owned()).or_insert(value);
            }
        }
        if !requested_business.is_empty() {
            let row_business = if row.get("业务名称").is_empty() {
                row.get("业务类型")
            } else {
                row.get("业务名称")
            };
            if row_business != requested_business {
                continue;
            }
        }
        rule_rows.push(row);
    }

    let metric_sheet = read_sheet(sheets, "模板指标目录", &METRIC_HEADERS, false)?;
    let mut raw_metric_rows: Vec<ConfigRow> = Vec::new();
    if let Some(sheet) = &metric_sheet {
        for row in &sheet.rows {
            if !requested_business.is_empty() && row.get("业务类型") != requested_business {
                continue;
            }
            raw_metric_rows.push(ConfigRow {
                row_num: row.row_num,
                values: row.values.clone(),
            });
        }
    }
    let mut metric_catalog: Vec<(String, String, Vec<String>, Vec<String>)> = Vec::new();
    for row in &raw_metric_rows {
        metric_catalog.push((
            required_text(row, "业务类型")?,
            required_text(row, "工作表名称")?,
            path_parts(&row.get("行路径"), "行路径")?,
            path_parts(&row.get("列路径"), "列路径")?,
        ));
    }

    let mut raw_scope_rows: Vec<ConfigRow> = Vec::new();
    if let Some(sheet) = read_sheet(sheets, "适用范围目录", &SCOPE_HEADERS, false)? {
        for row in &sheet.rows {
            if !requested_business.is_empty() && row.get("业务类型") != requested_business {
                continue;
            }
            raw_scope_rows.push(ConfigRow {
                row_num: row.row_num,
                values: row.values.clone(),
            });
        }
    }
    {
        let mut seen: std::collections::BTreeSet<String> = std::collections::BTreeSet::new();
        for row in &raw_scope_rows {
            let key = row.get("适用范围ID");
            if !seen.insert(key) {
                return Err(LoadError::configuration("适用范围目录存在重复适用范围ID"));
            }
        }
    }
    let mut scopes: Vec<ApplicabilityScope> = Vec::new();
    for row in &raw_scope_rows {
        let field = |name: &str| format!("适用范围目录第 {} 行{name}", row.row_num);
        scopes.push(ApplicabilityScope {
            scope_id: required_text(row, "适用范围ID")?,
            business: required_text(row, "业务类型")?,
            sheet: required_text(row, "工作表名称")?,
            row_path_prefix: if row.get("行路径前缀").trim().is_empty() {
                Vec::new()
            } else {
                path_parts(&row.get("行路径前缀"), "行路径前缀")?
            },
            description: row.get("范围说明"),
            confirmed: flag(&row.get("确认状态"), &field("确认状态"), None)?,
        });
    }

    let has_metric_sheet = metric_sheet.is_some();
    let has_external_sheet = sheets.iter().any(|sheet| sheet.name == "外部指标目录");
    let has_source_sheet = sheets.iter().any(|sheet| sheet.name == "数据源契约");

    let mut rules: Vec<RuleDefinition> = Vec::new();
    let mut derived_metrics: Vec<((String, String, Vec<String>, Vec<String>), MetricSpec)> =
        Vec::new();
    let mut derived_external: Vec<((String, String), ExternalMetricSpec)> = Vec::new();
    let mut referenced_template = false;
    let mut referenced_external: Vec<(String, String)> = Vec::new();
    let mut referenced_external_fields: Vec<(String, usize)> = Vec::new();

    for row in &rule_rows {
        let row_num = row.row_num;
        let rule_field = |name: &str| format!("审核规则第 {row_num} 行{name}");
        let rule_id = required_text(row, "规则ID")?;
        let revision = integer(&row.get("规则修订"), &rule_field("规则修订"), Some(1))?;
        if revision < 1 {
            return Err(LoadError::configuration(format!(
                "审核规则第 {row_num} 行规则修订必须大于零"
            )));
        }
        let enabled = flag(&row.get("是否启用"), &rule_field("是否启用"), None)?;
        let mut business = if row.get("业务名称").is_empty() {
            row.get("业务类型")
        } else {
            row.get("业务名称")
        };
        let default_sheet = if row.get("工作表名称").is_empty() {
            row.get("默认Sheet")
        } else {
            row.get("工作表名称")
        };
        let name = required_text(row, "规则名称")?;
        let severity = required_text(row, "审核级别")?;
        let check_type = required_text(row, "校验类型")?;
        let metric = required_text(row, "校验指标")?;
        let audit_expression = required_text(row, "审核表达式")?;
        let numeric_mode_label = if row.get("数值运算模式").is_empty() {
            "十进制".to_owned()
        } else {
            row.get("数值运算模式")
        };
        let numeric_mode = match numeric_mode_label.as_str() {
            "十进制" | "Decimal" => "decimal".to_owned(),
            "Excel双精度" | "Excel binary64" => "excel_binary64".to_owned(),
            _ => {
                return Err(LoadError::configuration(
                    "数值运算模式仅支持 十进制 或 Excel双精度",
                ))
            }
        };
        let formula_cell = row.get("校验公式单元格");
        let invert_audit_expression = flag(
            &row.get("表达式布尔值取反"),
            &rule_field("表达式布尔值取反"),
            Some(false),
        )?;
        let description = required_text(row, "规则描述")?;
        let rule_effective_from = row.get("生效起始日");
        let rule_effective_to = row.get("生效截止日");
        let applicability_scope = row.get("适用范围");
        validate_rule_dates(row)?;
        if !applicability_scope.is_empty()
            && !scopes
                .iter()
                .any(|scope| scope.scope_id == applicability_scope)
        {
            return Err(LoadError::configuration(format!(
                "审核规则第 {row_num} 行引用了不存在的适用范围ID：{applicability_scope}"
            )));
        }
        if enabled || include_disabled {
            let dictionary_text = row.get("取值字典");
            let dictionary = parse_dictionary(if dictionary_text.is_empty() {
                None
            } else {
                Some(dictionary_text.as_str())
            })?;
            validate_dictionary_dependencies(&dictionary)?;
            for (_, node) in &dictionary {
                validate_static_types(node, false)?;
            }
            let mut expression_nodes: Vec<(String, Node)> = Vec::new();
            expression_nodes.push((
                "审核表达式".to_owned(),
                parse_expression(&audit_expression)?,
            ));
            for field_name in ["当前值表达式", "对比值表达式", "差值表达式", "适用表达式"]
            {
                let expression = row.get(field_name);
                if !expression.is_empty() {
                    expression_nodes.push((field_name.to_owned(), parse_expression(&expression)?));
                }
            }
            for (field_name, node) in &expression_nodes {
                validate_static_types(node, false)?;
                let kind = static_kind(node);
                if matches!(field_name.as_str(), "审核表达式" | "适用表达式")
                    && !matches!(kind, None | Some("boolean"))
                {
                    return Err(LoadError::configuration(format!(
                        "{field_name} 必须返回布尔值"
                    )));
                }
                if field_name == "差值表达式" && !matches!(kind, None | Some("number")) {
                    return Err(LoadError::configuration("差值表达式必须返回数值"));
                }
            }
            for (field_name, node) in &expression_nodes {
                let undefined: Vec<String> = variables_in(node)
                    .into_iter()
                    .filter(|variable| !dictionary.iter().any(|(key, _)| key == variable))
                    .collect();
                if !undefined.is_empty() {
                    return Err(LoadError::configuration(format!(
                        "{field_name} 引用未定义变量：{}",
                        undefined.join(", ")
                    )));
                }
            }
            let rule_calls: Vec<&CallNode> = dictionary
                .iter()
                .flat_map(|(_, node)| calls_in(node))
                .chain(expression_nodes.iter().flat_map(|(_, node)| calls_in(node)))
                .collect();
            let current_row_calls: Vec<&CallNode> = rule_calls
                .iter()
                .copied()
                .filter(|call| is_current_row_call(call))
                .collect();
            let uses_current_row = !current_row_calls.is_empty();
            if uses_current_row && applicability_scope.is_empty() {
                return Err(LoadError::configuration(
                    "使用 @当前行，但没有配置适用范围ID，请补充适用范围",
                ));
            }
            if !applicability_scope.is_empty() && !uses_current_row {
                return Err(LoadError::configuration(format!(
                    "审核规则第 {row_num} 行配置了适用范围，但表达式没有使用 @当前行"
                )));
            }
            if business.is_empty() {
                business = infer_rule_business(
                    row_num,
                    &rule_calls,
                    &metric_catalog,
                    &scopes,
                    &applicability_scope,
                )?;
            }
            if !applicability_scope.is_empty() {
                let scope = scopes
                    .iter()
                    .find(|scope| scope.scope_id == applicability_scope)
                    .expect("scope existence validated above");
                if !scope.confirmed {
                    return Err(LoadError::configuration(format!(
                        "审核规则第 {row_num} 行引用的适用范围尚未确认：{applicability_scope}"
                    )));
                }
                if scope.business != business {
                    return Err(LoadError::configuration(format!(
                        "审核规则第 {row_num} 行适用范围业务与规则业务不一致"
                    )));
                }
                if !default_sheet.is_empty() && scope.sheet != default_sheet {
                    return Err(LoadError::configuration(format!(
                        "审核规则第 {row_num} 行适用范围工作表与默认Sheet不一致"
                    )));
                }
                for call in &current_row_calls {
                    if call.arguments.len() == 3 && call_argument_text(call, 0) != scope.sheet {
                        return Err(LoadError::configuration(format!(
                            "审核规则第 {row_num} 行 @当前行 的显式工作表与适用范围工作表不一致"
                        )));
                    }
                }
            }
            let mut field_nodes: Vec<(String, &Node)> = dictionary
                .iter()
                .map(|(_, node)| ("取值字典".to_owned(), node))
                .collect();
            for (field, node) in &expression_nodes {
                field_nodes.push((field.clone(), node));
            }
            for (field_name, node) in &field_nodes {
                for call in calls_in(node) {
                    if call.name == "V_CUR" || call.name == "V_PRE" {
                        referenced_template = true;
                    } else if call.name == "V_EXT" {
                        let metric_header = call
                            .arguments
                            .first()
                            .and_then(literal_string_value)
                            .unwrap_or_default();
                        let source_id = call
                            .keywords
                            .iter()
                            .find(|(key, _)| key == "source")
                            .map(|(_, value)| value.clone())
                            .unwrap_or_default();
                        referenced_external.push((source_id, metric_header.to_owned()));
                        referenced_external_fields.push((field_name.clone(), row_num));
                    }
                }
            }
            let description_vars = description_variables(&description)?;
            let dictionary_keys: std::collections::BTreeSet<&str> =
                dictionary.iter().map(|(key, _)| key.as_str()).collect();
            let undefined_description: Vec<String> = description_vars
                .iter()
                .filter(|variable| !dictionary_keys.contains(variable.as_str()))
                .cloned()
                .collect();
            if !undefined_description.is_empty() {
                return Err(LoadError::configuration(format!(
                    "规则描述引用未定义变量：{}",
                    undefined_description.join(", ")
                )));
            }
            let rule_definition = RuleDefinition {
                rule_id,
                revision,
                enabled,
                business: business.clone(),
                default_source: row.get("默认数据源"),
                default_sheet: default_sheet.clone(),
                name,
                severity,
                check_type,
                metric,
                audit_expression: audit_expression.clone(),
                description,
                audit_expression_inverted: invert_audit_expression,
                current_expression: row.get("当前值表达式"),
                comparison_expression: row.get("对比值表达式"),
                difference_expression: row.get("差值表达式"),
                value_dictionary: row.get("取值字典"),
                effective_from: rule_effective_from,
                effective_to: rule_effective_to,
                applicable_expression: row.get("适用表达式"),
                applicability_scope,
                numeric_mode,
                formula_cell,
            };
            if !has_metric_sheet || !has_external_sheet {
                for (field_name, node) in &field_nodes {
                    for call in calls_in(node) {
                        if (call.name == "V_CUR" || call.name == "V_PRE") && !has_metric_sheet {
                            let spec = template_metric_from_call(
                                call,
                                &rule_definition.rule_id,
                                &rule_definition.business,
                                &rule_definition.default_sheet,
                            )?;
                            let key = (
                                spec.business.clone(),
                                spec.sheet.clone(),
                                spec.row_path.clone(),
                                spec.column_path.clone(),
                            );
                            if let Some((_, previous)) = derived_metrics
                                .iter()
                                .find(|(key_existing, _)| key_existing == &key)
                            {
                                if previous != &spec {
                                    return Err(LoadError::configuration(format!(
                                        "审核规则引用同一模板路径时单位定义不一致：{} / {} / {}",
                                        spec.sheet,
                                        py_tuple_repr(&spec.row_path),
                                        py_tuple_repr(&spec.column_path)
                                    )));
                                }
                            } else {
                                derived_metrics.push((key, spec));
                            }
                        } else if call.name == "V_EXT" && !has_external_sheet {
                            let spec = external_metric_from_call(call, &rule_definition.rule_id)?;
                            let key = (spec.source_id.clone(), spec.metric_header.clone());
                            if let Some((_, previous)) = derived_external
                                .iter()
                                .find(|(key_existing, _)| key_existing == &key)
                            {
                                if previous != &spec {
                                    return Err(LoadError::configuration(format!(
                                        "审核规则引用同一外部指标时单位定义不一致：{} / {}",
                                        spec.source_id, spec.metric_header
                                    )));
                                }
                            } else {
                                derived_external.push((key, spec));
                            }
                        }
                    }
                    let _ = field_name;
                }
            }
            rules.push(rule_definition);
        } else if business.is_empty() {
            business =
                infer_rule_business(row_num, &[], &metric_catalog, &scopes, &applicability_scope)?;
            rules.push(RuleDefinition {
                rule_id,
                revision,
                enabled,
                business,
                default_source: row.get("默认数据源"),
                default_sheet,
                name,
                severity,
                check_type,
                metric,
                audit_expression,
                description,
                audit_expression_inverted: invert_audit_expression,
                current_expression: row.get("当前值表达式"),
                comparison_expression: row.get("对比值表达式"),
                difference_expression: row.get("差值表达式"),
                value_dictionary: row.get("取值字典"),
                effective_from: rule_effective_from,
                effective_to: rule_effective_to,
                applicable_expression: row.get("适用表达式"),
                applicability_scope,
                numeric_mode,
                formula_cell,
            });
        } else {
            rules.push(RuleDefinition {
                rule_id,
                revision,
                enabled,
                business,
                default_source: row.get("默认数据源"),
                default_sheet,
                name,
                severity,
                check_type,
                metric,
                audit_expression,
                description,
                audit_expression_inverted: invert_audit_expression,
                current_expression: row.get("当前值表达式"),
                comparison_expression: row.get("对比值表达式"),
                difference_expression: row.get("差值表达式"),
                value_dictionary: row.get("取值字典"),
                effective_from: rule_effective_from,
                effective_to: rule_effective_to,
                applicable_expression: row.get("适用表达式"),
                applicability_scope,
                numeric_mode,
                formula_cell,
            });
        }
    }

    let mut metric_specs: Vec<MetricSpec> = Vec::new();
    if let Some(sheet) = &metric_sheet {
        for row in &sheet.rows {
            let row_num = row.row_num;
            let field = |name: &str| format!("模板指标目录第 {row_num} 行{name}");
            let metric_type = canonical_metric_type(&row.get("指标类型"));
            let row_header_raw = row.get("行表头标题");
            let spec = MetricSpec {
                business: required_text(row, "业务类型")?,
                sheet: required_text(row, "工作表名称")?,
                row_path: path_parts(&row.get("行路径"), "行路径")?,
                column_path: path_parts(&row.get("列路径"), "列路径")?,
                row_header_titles: if row_header_raw.trim().is_empty() {
                    Vec::new()
                } else {
                    path_parts(&row_header_raw, "行表头标题")?
                },
                data_type: canonical_data_type(&row.get("数据类型")),
                metric_type: metric_type.clone(),
                source_unit: canonical_unit(&row.get("来源单位"), &metric_type),
                standard_unit: canonical_unit(&row.get("标准单位"), &metric_type),
                storage_scale: canonical_storage_scale(&row.get("存储尺度")),
                allow_blank: flag(&row.get("允许空值"), &field("允许空值"), Some(false))?,
                precision: integer(&row.get("精度"), &field("精度"), Some(2))?,
                evidence: row.get("口径依据"),
                confirmed: flag(&row.get("确认状态"), &field("确认状态"), None)?,
                layout_version: row.get("布局版本"),
            };
            if spec.row_path.len() > spec.row_header_titles.len() {
                return Err(LoadError::configuration(format!(
                    "模板指标目录第 {row_num} 行行路径层级数不能多于行表头标题数"
                )));
            }
            validate_data_type(&spec.data_type, &spec.metric_type, row_num)?;
            validate_precision(spec.precision, row_num)?;
            validate_metric_units(
                &spec.metric_type,
                &spec.source_unit,
                &spec.standard_unit,
                &spec.storage_scale,
            )?;
            metric_specs.push(spec);
        }
    }
    if !has_metric_sheet {
        metric_specs = derived_metrics
            .iter()
            .map(|(_, spec)| spec.clone())
            .collect();
    }

    let mut external_specs: Vec<ExternalMetricSpec> = Vec::new();
    if let Some(sheet) = read_sheet(sheets, "外部指标目录", &EXTERNAL_METRIC_HEADERS, false)?
    {
        for row in &sheet.rows {
            let row_num = row.row_num;
            let field = |name: &str| format!("外部指标目录第 {row_num} 行{name}");
            let metric_type = canonical_metric_type(&row.get("指标类型"));
            let spec = ExternalMetricSpec {
                source_id: required_text(row, "数据源ID")?,
                metric_header: required_text(row, "指标表头")?,
                data_type: canonical_data_type(&row.get("数据类型")),
                metric_type: metric_type.clone(),
                source_unit: canonical_unit(&row.get("来源单位"), &metric_type),
                standard_unit: canonical_unit(&row.get("标准单位"), &metric_type),
                storage_scale: canonical_storage_scale(&row.get("存储尺度")),
                allow_blank: flag(&row.get("允许空值"), &field("允许空值"), Some(false))?,
                precision: integer(&row.get("精度"), &field("精度"), Some(2))?,
                evidence: row.get("口径依据"),
                confirmed: flag(&row.get("确认状态"), &field("确认状态"), None)?,
            };
            validate_data_type(&spec.data_type, &spec.metric_type, row_num)?;
            validate_precision(spec.precision, row_num)?;
            validate_metric_units(
                &spec.metric_type,
                &spec.source_unit,
                &spec.standard_unit,
                &spec.storage_scale,
            )?;
            external_specs.push(spec);
        }
    }
    if !has_external_sheet {
        external_specs = derived_external
            .iter()
            .map(|(_, spec)| spec.clone())
            .collect();
    }

    let mut contracts: Vec<SourceContract> = Vec::new();
    if let Some(sheet) = read_sheet(sheets, "数据源契约", &SOURCE_HEADERS, false)? {
        for row in &sheet.rows {
            let row_num = row.row_num;
            let field = |name: &str| format!("数据源契约第 {row_num} 行{name}");
            let source_type = required_text(row, "来源类型")?.to_lowercase();
            if !matches!(
                source_type.as_str(),
                "current" | "previous" | "external" | "本期" | "上期" | "外部"
            ) {
                return Err(LoadError::configuration(format!(
                    "数据源契约第 {row_num} 行来源类型不支持：{source_type}"
                )));
            }
            let contract = SourceContract {
                source_id: required_text(row, "数据源ID")?,
                source_type: source_type.clone(),
                default_sheet: row.get("默认Sheet"),
                identity_sheet: row.get("机构Sheet"),
                institution_code_label: row.get("机构代码标签"),
                institution_name_label: row.get("机构名称标签"),
                period_label: row.get("期间标签"),
                mapping_sheet: row.get("映射Sheet"),
                mapping_header_row: integer(&row.get("映射表头行"), &field("映射表头行"), Some(1))?,
                mapping_data_start_row: integer(
                    &row.get("映射数据起始行"),
                    &field("映射数据起始行"),
                    Some(2),
                )?,
                mapping_identity_field: row.get("映射机构字段"),
                mapping_target_field: row.get("映射目标字段"),
                data_header_row: integer(&row.get("数据表头行"), &field("数据表头行"), Some(1))?,
                data_start_row: integer(&row.get("数据起始行"), &field("数据起始行"), Some(2))?,
                data_key_column: integer(&row.get("数据键列"), &field("数据键列"), Some(1))?,
                metric_header_row: integer(&row.get("指标表头行"), &field("指标表头行"), Some(1))?,
                metric_first_column: integer(
                    &row.get("指标起始列"),
                    &field("指标起始列"),
                    Some(2),
                )?,
                report_item_label: row.get("报表项目标签"),
                expected_period: row.get("期望期间"),
                version: row.get("契约版本"),
                period_required: flag(&row.get("期间必须"), &field("期间必须"), Some(false))?,
                institution_required: flag(&row.get("机构必须"), &field("机构必须"), Some(false))?,
            };
            if [
                contract.mapping_header_row,
                contract.mapping_data_start_row,
                contract.data_header_row,
                contract.data_start_row,
                contract.data_key_column,
                contract.metric_header_row,
                contract.metric_first_column,
            ]
            .iter()
            .any(|value| *value < 1)
            {
                return Err(LoadError::configuration("表头/起始行/列号必须大于零"));
            }
            if matches!(source_type.as_str(), "external" | "外部") {
                let missing: Vec<&str> = [
                    ("默认Sheet", contract.default_sheet.is_empty()),
                    ("映射Sheet", contract.mapping_sheet.is_empty()),
                    ("映射机构字段", contract.mapping_identity_field.is_empty()),
                    ("映射目标字段", contract.mapping_target_field.is_empty()),
                ]
                .into_iter()
                .filter(|(_, missing)| *missing)
                .map(|(name, _)| name)
                .collect();
                if !missing.is_empty() {
                    return Err(LoadError::configuration(
                        "外部数据源契约必须配置默认Sheet和映射字段",
                    ));
                }
            }
            contracts.push(contract);
        }
    }

    let mut version = String::new();
    let mut version_status = String::new();
    let mut release_date = String::new();
    let mut effective_from = String::new();
    let mut effective_to = String::new();
    let mut release_note = String::new();
    if let Some(sheet) = read_sheet(sheets, "配置版本", &VERSION_HEADERS, false)? {
        if sheet.rows.len() > 1 {
            return Err(LoadError::configuration(
                "配置版本工作表只能有一行当前版本信息",
            ));
        }
        if let Some(row) = sheet.rows.first() {
            version = required_text(row, "配置版本号")?;
            version_status = required_text(row, "状态")?;
            release_date = row.get("发布日期");
            if !release_date.is_empty() {
                validate_date_range(&release_date, &release_date, "发布日期")?;
            }
            effective_from = row.get("生效起始日");
            effective_to = row.get("生效截止日");
            validate_date_range(&effective_from, &effective_to, "生效期")?;
            release_note = row.get("发布说明");
        }
    }

    {
        let mut seen: std::collections::BTreeSet<(String, String, Vec<String>, Vec<String>)> =
            std::collections::BTreeSet::new();
        for spec in &metric_specs {
            let key = (
                spec.business.clone(),
                spec.sheet.clone(),
                spec.row_path.clone(),
                spec.column_path.clone(),
            );
            if !seen.insert(key) {
                return Err(LoadError::configuration(
                    "模板指标目录存在重复业务/Sheet/行路径/列路径",
                ));
            }
        }
    }
    {
        let mut seen: std::collections::BTreeSet<(String, String)> =
            std::collections::BTreeSet::new();
        for spec in &external_specs {
            let key = (spec.source_id.clone(), spec.metric_header.clone());
            if !seen.insert(key) {
                return Err(LoadError::configuration(
                    "外部指标目录存在重复数据源/指标表头",
                ));
            }
        }
    }
    {
        let mut seen: std::collections::BTreeSet<String> = std::collections::BTreeSet::new();
        for contract in &contracts {
            if !seen.insert(contract.source_id.clone()) {
                return Err(LoadError::configuration("数据源契约存在重复数据源ID"));
            }
        }
    }
    for scope in &scopes {
        let matched = metric_specs.iter().any(|metric| {
            metric.business == scope.business
                && metric.sheet == scope.sheet
                && metric.row_path.len() >= scope.row_path_prefix.len()
                && metric.row_path[..scope.row_path_prefix.len()] == scope.row_path_prefix[..]
        });
        if !matched {
            return Err(LoadError::configuration(format!(
                "适用范围 {} 没有匹配到模板指标目录行路径",
                scope.scope_id
            )));
        }
    }
    if referenced_template && metric_specs.is_empty() {
        return Err(LoadError::configuration(
            "规则使用 V_CUR/V_PRE，但无法从审核规则表达式生成模板取值路径，请检查模板指标目录是否为空",
        ));
    }
    if !referenced_external.is_empty() && external_specs.is_empty() {
        return Err(LoadError::configuration(
            "规则使用 V_EXT，但无法从审核规则表达式生成外部指标信息，请检查外部指标目录是否为空",
        ));
    }
    for spec in &external_specs {
        let contract = contracts
            .iter()
            .find(|contract| contract.source_id == spec.source_id);
        if contract.is_none() && has_source_sheet {
            return Err(LoadError::configuration(format!(
                "外部指标 {} 缺少对应数据源契约：{}",
                spec.metric_header, spec.source_id
            )));
        }
        if let Some(contract) = contract {
            if !matches!(contract.source_type.as_str(), "external" | "外部") {
                return Err(LoadError::configuration(format!(
                    "外部指标 {} 引用了非外部来源契约：{}",
                    spec.metric_header, spec.source_id
                )));
            }
        }
    }
    if has_external_sheet {
        let external_keys: std::collections::BTreeSet<(String, String)> = external_specs
            .iter()
            .map(|spec| (spec.source_id.clone(), spec.metric_header.clone()))
            .collect();
        let external_headers: std::collections::BTreeSet<String> = external_specs
            .iter()
            .map(|spec| spec.metric_header.clone())
            .collect();
        for (source_id, metric_header) in &referenced_external {
            if !source_id.is_empty()
                && !external_keys.contains(&(source_id.clone(), metric_header.clone()))
            {
                return Err(LoadError::configuration(format!(
                    "V_EXT 引用的外部指标未登记：{source_id} / {metric_header}"
                )));
            }
            if source_id.is_empty() && !external_headers.contains(metric_header) {
                return Err(LoadError::configuration(format!(
                    "V_EXT 引用的外部指标未登记：{metric_header}"
                )));
            }
        }
    }
    let _ = referenced_external_fields;

    Ok(Configuration {
        rules,
        metrics: metric_specs,
        external_metrics: external_specs,
        source_contracts: contracts,
        applicability_scopes: scopes,
        version,
        version_status,
        effective_from,
        effective_to,
        release_date,
        release_note,
        sha256: config_hash,
    })
}
