//! S1 formula-result extraction from values already stored in an audit copy.
//!
//! This stage consumes Calamine's stored cell values, which are normally Excel
//! formula caches. It does not recalculate formulas or modify either workbook.

use crate::{
    models::{AuditRule, Issue, TemplateDefinition},
    template::{normalize_template_name, parse_formula_result},
};
use anyhow::{ensure, Context, Result};
use excel_reader::{read_value_columns, CellValue, SheetValues};
use serde::Serialize;
use serde_json::Value;
use std::{
    collections::{BTreeMap, BTreeSet},
    fs,
    path::{Path, PathBuf},
    time::Instant,
};

const READER: &str = "calamine";
const RESULT_KEYWORDS: [&str; 4] = ["错误", "硬性", "错", "软性"];
const FORMULA_ERRORS: [&str; 11] = [
    "#REF!",
    "#VALUE!",
    "#N/A",
    "#DIV/0!",
    "#NAME?",
    "#NUM!",
    "#NULL!",
    "#SPILL!",
    "#CALC!",
    "#FIELD!",
    "#GETTING_DATA",
];

#[derive(Clone, Debug, Default)]
pub struct S1FormulaMetadata {
    pub period: String,
    pub batch_id: String,
    pub audit_time: String,
    pub org_code: String,
    pub org_name: String,
}

#[derive(Clone, Debug, Serialize)]
pub struct S1FormulaAuditSummary {
    pub reader: &'static str,
    pub rules_total: usize,
    pub rules_enabled: usize,
    pub rules_selected: usize,
    pub formula_result_values_nonempty: usize,
    pub issue_count: usize,
    pub read_time_ms: u128,
    pub issue_build_time_ms: u128,
    pub total_time_ms: u128,
}

#[derive(Clone, Debug, Serialize)]
pub struct S1FormulaAuditResult {
    pub status: &'static str,
    pub issues: Vec<Issue>,
    pub summary: S1FormulaAuditSummary,
}

/// Extract standard S1 Issues using values already stored in the audit copy.
///
/// `source_file` defaults to the audit workbook, matching the Python stage.
/// `extraction_ranges` restrict rule formula cells when configured; an empty
/// list means all enabled rules. Errors returned by formula cells are still
/// emitted for `error_only` rules.
pub fn extract_s1_formula_issues(
    definition: &TemplateDefinition,
    audit_workbook_path: impl AsRef<Path>,
    source_file: Option<&Path>,
    metadata: &S1FormulaMetadata,
) -> Result<S1FormulaAuditResult> {
    let started = Instant::now();
    let audit_path = audit_workbook_path.as_ref();
    ensure!(
        audit_path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| {
                extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm")
            }),
        "S1 公式结果读取目前支持 .xlsx 和 .xlsm"
    );
    let source_path = source_file.unwrap_or(audit_path);
    let selected_rules = definition
        .rules
        .iter()
        .filter(|rule| {
            rule.enabled
                && (definition.extraction_ranges.is_empty()
                    || definition.extraction_ranges.iter().any(|range| {
                        range.sheet_name == rule.sheet_name
                            && address_is_in_range(&rule.formula_cell, &range.address)
                    }))
        })
        .collect::<Vec<_>>();

    let mut required_sheets = BTreeSet::new();
    let mut max_column = 1;
    for rule in &selected_rules {
        required_sheets.insert(rule.sheet_name.as_str());
        max_column = max_column.max(
            parse_cell_address(&rule.formula_cell)
                .with_context(|| {
                    format!(
                        "无效的公式单元格：{}!{}",
                        rule.sheet_name, rule.formula_cell
                    )
                })?
                .1,
        );
        if !rule.value_from_result {
            max_column = max_column.max(
                parse_cell_address(&rule.target_cell)
                    .with_context(|| {
                        format!("无效的定位单元格：{}!{}", rule.sheet_name, rule.target_cell)
                    })?
                    .1,
            );
        }
    }

    let read_started = Instant::now();
    let workbook = read_value_columns(audit_path, max_column)
        .with_context(|| format!("读取审核副本中的公式结果值：{}", audit_path.display()))?;
    let read_time_ms = read_started.elapsed().as_millis();
    let sheets = workbook
        .iter()
        .map(|sheet| (sheet.name.as_str(), sheet))
        .collect::<BTreeMap<_, _>>();
    for sheet_name in required_sheets {
        ensure!(
            sheets.contains_key(sheet_name),
            "审核副本缺少工作表：{sheet_name}"
        );
    }

    let issue_build_started = Instant::now();
    let mut issues = Vec::new();
    let mut formula_result_values_nonempty = 0;
    let source_display = resolved_path(source_path);
    let audit_display = resolved_path(audit_path);
    for rule in &selected_rules {
        let sheet = sheets
            .get(rule.sheet_name.as_str())
            .with_context(|| format!("审核副本缺少工作表：{}", rule.sheet_name))?;
        let formula_result = cell_value(sheet, &rule.formula_cell);
        if !is_empty_formula_result(&formula_result) {
            formula_result_values_nonempty += 1;
        }

        if let Some(error) = formula_error(&formula_result) {
            issues.push(formula_error_issue(
                rule,
                source_path,
                &source_display,
                &audit_display,
                metadata,
                error,
            ));
            continue;
        }
        if rule.result_mode == "error_only" || is_empty_formula_result(&formula_result) {
            continue;
        }
        if rule.result_mode == "keyword" && !has_result_marker(&formula_result) {
            continue;
        }

        let formula_result_json = cell_value_to_json(&formula_result);
        let parsed = parse_formula_result(
            &formula_result_json,
            Some(&rule.severity),
            Some(&rule.message),
        );
        let target_value = if rule.value_from_result {
            Value::String(parsed.value.clone())
        } else {
            cell_value_to_json(&cell_value(sheet, &rule.target_cell))
        };
        let target_value = if is_python_falsey_json(&target_value) {
            Value::String(String::new())
        } else {
            target_value
        };
        let identity = issue_identity(source_path, rule, &parsed.indicator, &parsed.check_type);
        issues.push(Issue {
            issue_id: identity.clone(),
            period: metadata.period.clone(),
            batch_id: metadata.batch_id.clone(),
            audit_time: metadata.audit_time.clone(),
            triggered: true,
            status: String::new(),
            first_seen_period: String::new(),
            previous_seen_period: String::new(),
            consecutive_count: 1,
            org_code: metadata.org_code.clone(),
            org_name: metadata.org_name.clone(),
            report_code: rule.report_code.clone(),
            sheet_name: rule.sheet_name.clone(),
            rule_id: identity,
            severity: parsed.severity,
            formula_cell: rule.formula_cell.clone(),
            target_cell: rule.target_cell.clone(),
            target_value,
            formula_result: formula_result_json,
            message: parsed.message,
            source_file: source_display.clone(),
            audit_file: audit_display.clone(),
            check_field: parsed.indicator,
            check_type: parsed.check_type,
            comparison_value: Value::String(parsed.comparison_value),
            reference_value: Value::String(parsed.reference_value),
            difference_value: Value::String(parsed.difference_value),
            detail: parsed.detail,
            institution_feedback: String::new(),
            auditor_opinion: String::new(),
        });
    }

    let issue_build_time_ms = issue_build_started.elapsed().as_millis();
    let summary = S1FormulaAuditSummary {
        reader: READER,
        rules_total: definition.rules.len(),
        rules_enabled: definition.rules.iter().filter(|rule| rule.enabled).count(),
        rules_selected: selected_rules.len(),
        formula_result_values_nonempty,
        issue_count: issues.len(),
        read_time_ms,
        issue_build_time_ms,
        total_time_ms: started.elapsed().as_millis(),
    };
    Ok(S1FormulaAuditResult {
        status: "completed",
        issues,
        summary,
    })
}

fn formula_error_issue(
    rule: &AuditRule,
    source_path: &Path,
    source_display: &str,
    audit_display: &str,
    metadata: &S1FormulaMetadata,
    error: &str,
) -> Issue {
    let detail = if error == "#REF!" {
        "校验公式引用无效（#REF!）。请确认外部文件添加步骤已执行，且工作表名称与模板公式一致。"
            .to_owned()
    } else {
        format!("校验公式计算异常：{error}，请检查模板公式及引用数据")
    };
    let identity = issue_identity(source_path, rule, "", "");
    Issue {
        issue_id: identity.clone(),
        period: metadata.period.clone(),
        batch_id: metadata.batch_id.clone(),
        audit_time: metadata.audit_time.clone(),
        triggered: true,
        status: String::new(),
        first_seen_period: String::new(),
        previous_seen_period: String::new(),
        consecutive_count: 1,
        org_code: metadata.org_code.clone(),
        org_name: metadata.org_name.clone(),
        report_code: rule.report_code.clone(),
        sheet_name: rule.sheet_name.clone(),
        rule_id: identity,
        severity: "错误".to_owned(),
        formula_cell: rule.formula_cell.clone(),
        target_cell: rule.formula_cell.clone(),
        target_value: Value::String(String::new()),
        formula_result: Value::String(error.to_owned()),
        message: detail.clone(),
        source_file: source_display.to_owned(),
        audit_file: audit_display.to_owned(),
        detail,
        check_field: String::new(),
        check_type: String::new(),
        comparison_value: Value::String(String::new()),
        reference_value: Value::String(String::new()),
        difference_value: Value::String(String::new()),
        institution_feedback: String::new(),
        auditor_opinion: String::new(),
    }
}

fn issue_identity(
    source_path: &Path,
    rule: &AuditRule,
    indicator: &str,
    check_type: &str,
) -> String {
    let source_stem = source_path
        .file_stem()
        .map(|stem| stem.to_string_lossy().into_owned())
        .unwrap_or_default();
    let source_name = normalize_template_name(&source_stem);
    let mut parts = vec![
        if source_name.is_empty() {
            "工作簿未识别".to_owned()
        } else {
            source_name
        },
        if rule.sheet_name.is_empty() {
            "工作表未识别".to_owned()
        } else {
            rule.sheet_name.clone()
        },
        if rule.formula_cell.is_empty() {
            "公式单元格未识别".to_owned()
        } else {
            rule.formula_cell.clone()
        },
    ];
    if !check_type.is_empty() {
        parts.push(check_type.to_owned());
    }
    parts.push(if indicator.is_empty() {
        "校验指标未识别".to_owned()
    } else {
        indicator.to_owned()
    });
    parts.join("｜")
}

fn cell_value(sheet: &SheetValues, address: &str) -> CellValue {
    sheet
        .values
        .get(&normalize_address(address))
        .cloned()
        .unwrap_or(CellValue::Empty)
}

fn cell_value_to_json(value: &CellValue) -> Value {
    match value {
        CellValue::Empty => Value::Null,
        CellValue::Text(value) => Value::String(value.clone()),
        CellValue::Number(value) | CellValue::DateTimeSerial(value) => {
            serde_json::Number::from_f64(*value)
                .map(Value::Number)
                .unwrap_or(Value::Null)
        }
        CellValue::Integer(value) => Value::from(*value),
        CellValue::Boolean(value) => Value::Bool(*value),
        CellValue::DateTimeIso(value) | CellValue::DurationIso(value) => {
            Value::String(value.clone())
        }
        CellValue::Error(value) => {
            Value::String(excel_error_text(value).unwrap_or(value.as_str()).to_owned())
        }
    }
}

fn formula_error(value: &CellValue) -> Option<&'static str> {
    let text = match value {
        CellValue::Error(value) => excel_error_text(value).unwrap_or(value).trim(),
        CellValue::Text(value) => value.trim(),
        _ => return None,
    };
    let upper = text.to_uppercase();
    FORMULA_ERRORS
        .iter()
        .copied()
        .find(|error| upper.starts_with(error))
}

fn excel_error_text(value: &str) -> Option<&'static str> {
    let normalized = value
        .trim()
        .trim_start_matches('#')
        .trim_end_matches('!')
        .trim_end_matches('?')
        .replace('_', "")
        .replace('-', "")
        .replace(' ', "")
        .to_ascii_lowercase();
    match normalized.as_str() {
        "ref" => Some("#REF!"),
        "value" => Some("#VALUE!"),
        "na" | "notavailable" => Some("#N/A"),
        "div0" | "divzero" => Some("#DIV/0!"),
        "name" => Some("#NAME?"),
        "num" => Some("#NUM!"),
        "null" => Some("#NULL!"),
        "spill" => Some("#SPILL!"),
        "calc" => Some("#CALC!"),
        "field" => Some("#FIELD!"),
        "gettingdata" => Some("#GETTING_DATA"),
        _ => None,
    }
}

fn is_empty_formula_result(value: &CellValue) -> bool {
    match value {
        CellValue::Empty => true,
        CellValue::Boolean(value) => !value,
        CellValue::Integer(value) => *value == 0,
        CellValue::Number(value) | CellValue::DateTimeSerial(value) => *value == 0.0,
        CellValue::Text(value) => value.trim().is_empty(),
        CellValue::DateTimeIso(value) | CellValue::DurationIso(value) => value.trim().is_empty(),
        CellValue::Error(value) => value.trim().is_empty(),
    }
}

fn has_result_marker(value: &CellValue) -> bool {
    matches!(value, CellValue::Text(text) if RESULT_KEYWORDS.iter().any(|marker| text.trim().starts_with(marker)))
}

fn is_python_falsey_json(value: &Value) -> bool {
    match value {
        Value::Null => true,
        Value::Bool(value) => !value,
        Value::Number(value) => value.as_f64().is_some_and(|number| number == 0.0),
        Value::String(value) => value.is_empty(),
        Value::Array(_) | Value::Object(_) => false,
    }
}

fn normalize_address(value: &str) -> String {
    value.replace('$', "").trim().to_ascii_uppercase()
}

fn address_is_in_range(cell: &str, address: &str) -> bool {
    let Some((row, column)) = parse_cell_address(cell) else {
        return false;
    };
    let address = address.replace('$', "");
    let mut parts = address.split(':');
    let Some((first_row, first_column)) = parts.next().and_then(parse_cell_address) else {
        return false;
    };
    let (last_row, last_column) = match parts.next() {
        Some(value) => match parse_cell_address(value) {
            Some(bounds) => bounds,
            None => return false,
        },
        None => (first_row, first_column),
    };
    parts.next().is_none()
        && first_row <= last_row
        && first_column <= last_column
        && (first_row..=last_row).contains(&row)
        && (first_column..=last_column).contains(&column)
}

fn parse_cell_address(value: &str) -> Option<(usize, usize)> {
    let value = value.replace('$', "");
    let split = value.find(|character: char| character.is_ascii_digit())?;
    let (column, row) = value.split_at(split);
    if column.is_empty()
        || !column.bytes().all(|byte| byte.is_ascii_alphabetic())
        || row.is_empty()
        || !row.bytes().all(|byte| byte.is_ascii_digit())
    {
        return None;
    }
    let column = column.bytes().try_fold(0usize, |value, byte| {
        value
            .checked_mul(26)?
            .checked_add(usize::from(byte.to_ascii_uppercase() - b'A' + 1))
    })?;
    let row = row.parse::<usize>().ok()?;
    (column > 0 && row > 0).then_some((row, column))
}

fn resolved_path(path: &Path) -> String {
    if let Ok(path) = fs::canonicalize(path) {
        return path.display().to_string();
    }
    if path.is_absolute() {
        return path.display().to_string();
    }
    std::env::current_dir()
        .map(|directory| directory.join(path).display().to_string())
        .unwrap_or_else(|_| PathBuf::from(path).display().to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::models::{AuditRule, CopyRange};
    use excel_writer::Workbook;
    use std::{fs, path::PathBuf, time::SystemTime};

    fn temporary_workbook_path() -> PathBuf {
        let nonce = SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .expect("system clock after Unix epoch")
            .as_nanos();
        std::env::temp_dir().join(format!(
            "audit_s1_formula_audit_{}_{}.xlsx",
            std::process::id(),
            nonce
        ))
    }

    fn rule(
        formula_cell: &str,
        target_cell: &str,
        result_mode: &str,
        value_from_result: bool,
        enabled: bool,
    ) -> AuditRule {
        AuditRule {
            rule_id: format!("rule-{formula_cell}"),
            enabled,
            report_code: "报表甲".to_owned(),
            sheet_name: "检查表".to_owned(),
            formula_cell: formula_cell.to_owned(),
            target_cell: target_cell.to_owned(),
            severity: "错误".to_owned(),
            message: "默认说明".to_owned(),
            copy_range: formula_cell.to_owned(),
            result_mode: result_mode.to_owned(),
            value_from_result,
        }
    }

    #[test]
    fn extracts_formula_issues_with_python_filter_and_error_semantics() {
        let path = temporary_workbook_path();
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.set_name("检查表").expect("name formula sheet");
        sheet
            .write_string(0, 0, "错误|复核|存款|低于要求|100|200|100")
            .expect("write keyword result");
        sheet.write_string(1, 0, "通过").expect("write non-marker");
        sheet
            .write_string(2, 0, "#REF!")
            .expect("write formula error");
        sheet
            .write_string(3, 0, "通过")
            .expect("write skipped error-only result");
        sheet
            .write_string(4, 0, "软性|比率|资本充足率|偏低|12|15|3")
            .expect("write structured result");
        sheet.write_number(4, 1, 42).expect("write target value");
        sheet.write_number(5, 0, 0).expect("write falsey result");
        workbook.save(&path).expect("save cached values workbook");

        let definition = TemplateDefinition {
            rules: vec![
                rule("A1", "A1", "keyword", true, true),
                rule("A2", "A2", "keyword", true, true),
                rule("A3", "A3", "error_only", true, true),
                rule("A4", "A4", "error_only", true, true),
                rule("A5", "B5", "nonempty", false, true),
                rule("A6", "A6", "nonempty", true, true),
                rule("A7", "A7", "keyword", true, false),
            ],
            copy_ranges: Vec::new(),
            structured: false,
            structure_ranges: Vec::new(),
            extraction_ranges: vec![CopyRange {
                sheet_name: "检查表".to_owned(),
                address: "$A$1:$A$5".to_owned(),
            }],
        };
        let metadata = S1FormulaMetadata {
            period: "2026-07".to_owned(),
            batch_id: "batch-01".to_owned(),
            audit_time: "2026-10-24 12:00:00".to_owned(),
            org_code: "1234".to_owned(),
            org_name: "机构甲".to_owned(),
        };
        let source = PathBuf::from("!报表甲_20260731.xlsx");
        let result = extract_s1_formula_issues(&definition, &path, Some(&source), &metadata)
            .expect("extract formula issues");
        fs::remove_file(&path).expect("remove synthetic workbook");

        assert_eq!(result.status, "completed");
        assert_eq!(result.summary.reader, "calamine");
        assert_eq!(result.summary.rules_total, 7);
        assert_eq!(result.summary.rules_enabled, 6);
        assert_eq!(result.summary.rules_selected, 5);
        assert_eq!(result.summary.formula_result_values_nonempty, 5);
        assert_eq!(result.issues.len(), 3);
        let keyword = &result.issues[0];
        assert_eq!(keyword.period, "2026-07");
        assert_eq!(keyword.org_code, "1234");
        assert_eq!(keyword.severity, "错误");
        assert_eq!(keyword.check_type, "复核");
        assert_eq!(keyword.check_field, "存款");
        assert_eq!(keyword.target_value, Value::String("100".to_owned()));
        assert_eq!(
            keyword.formula_result,
            Value::String("错误|复核|存款|低于要求|100|200|100".to_owned())
        );

        let formula_error = &result.issues[1];
        assert_eq!(formula_error.severity, "错误");
        assert_eq!(
            formula_error.formula_result,
            Value::String("#REF!".to_owned())
        );
        assert_eq!(formula_error.target_cell, formula_error.formula_cell);
        assert!(formula_error.message.contains("外部文件添加步骤"));
        assert!(formula_error.check_field.is_empty());

        let target_value_issue = &result.issues[2];
        assert_eq!(target_value_issue.severity, "核实");
        assert_eq!(target_value_issue.check_field, "资本充足率");
        assert_eq!(target_value_issue.target_value.as_f64(), Some(42.0));
        assert_eq!(
            target_value_issue.comparison_value,
            Value::String("15".to_owned())
        );
        assert_eq!(
            target_value_issue.difference_value,
            Value::String("3".to_owned())
        );
    }

    #[test]
    fn maps_calamine_formula_error_variants_to_python_error_text() {
        for (source, expected) in [
            ("Ref", "#REF!"),
            ("Div0", "#DIV/0!"),
            ("Value", "#VALUE!"),
            ("NA", "#N/A"),
            ("GettingData", "#GETTING_DATA"),
        ] {
            let error = CellValue::Error(source.to_owned());
            assert_eq!(formula_error(&error), Some(expected));
            assert_eq!(
                cell_value_to_json(&error),
                Value::String(expected.to_owned())
            );
        }
    }

    #[test]
    fn empty_extraction_ranges_run_every_enabled_rule() {
        let path = temporary_workbook_path();
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.set_name("检查表").expect("name formula sheet");
        sheet
            .write_string(0, 0, "错误|类型|指标|说明")
            .expect("write formula result");
        workbook.save(&path).expect("save formula results");
        let definition = TemplateDefinition {
            rules: vec![rule("A1", "A1", "nonempty", true, true)],
            copy_ranges: Vec::new(),
            structured: false,
            structure_ranges: Vec::new(),
            extraction_ranges: Vec::new(),
        };
        let result =
            extract_s1_formula_issues(&definition, &path, None, &S1FormulaMetadata::default())
                .expect("extract all enabled rules without a configured gate");
        fs::remove_file(&path).expect("remove synthetic workbook");
        assert_eq!(result.summary.rules_selected, 1);
        assert_eq!(result.issues.len(), 1);
        assert!(result.issues[0]
            .issue_id
            .starts_with("audit_s1_formula_audit_"));
    }
}
