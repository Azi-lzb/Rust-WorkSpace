//! S1 conditional-format extraction from evaluator results into Audit issues.
//!
//! This is the business-output stage for the existing static Rust evaluator.
//! Template configuration loading, the remaining S1 rule engine, history, and
//! workbook writeback remain separate migration work.

use crate::{
    conditional_format::{evaluate_xlsx_workbook_with_dump, ConditionalEvaluationReport},
    models::{CopyRange, Issue},
    template::normalize_template_name,
};
use anyhow::{Context, Result};
use excel_reader::{CellValue, SheetDump};
use serde::Serialize;
use serde_json::{json, Value};
use std::{collections::BTreeSet, path::Path, time::Instant};

#[derive(Clone, Debug, Default)]
pub struct S1ConditionalMetadata {
    pub period: String,
    pub batch_id: String,
    pub audit_time: String,
    pub org_code: String,
    pub org_name: String,
}

#[derive(Clone, Debug, Default, Serialize)]
pub struct S1ConditionalContextSummary {
    pub unresolved_indicator_count: usize,
    pub comments_unverified_sheet_count: usize,
}

#[derive(Clone, Debug, Serialize)]
pub struct S1ConditionalAuditResult {
    pub status: &'static str,
    pub issues: Vec<Issue>,
    pub evaluation: ConditionalEvaluationReport,
    pub context: S1ConditionalContextSummary,
    pub issue_build_time_ms: u128,
}

/// Run the static conditional-format evaluator and construct the S1-F01 issue
/// rows that its Python counterpart feeds into AuditResult/history processing.
/// The workbook is read once; no recalculation or workbook mutation occurs.
pub fn extract_conditional_issues(
    workbook_path: impl AsRef<Path>,
    source_file: impl AsRef<Path>,
    scan_areas: &[String],
    structure_ranges: &[CopyRange],
    metadata: &S1ConditionalMetadata,
) -> Result<S1ConditionalAuditResult> {
    let workbook_path = workbook_path.as_ref();
    let source_file = source_file.as_ref();
    let (evaluation, workbook) = evaluate_xlsx_workbook_with_dump(workbook_path, scan_areas)?;
    let build_started = Instant::now();
    let source_stem = source_file
        .file_stem()
        .map(|stem| stem.to_string_lossy().into_owned())
        .unwrap_or_default();
    let normalized_source = normalize_template_name(&source_stem);
    let workbook_key = if normalized_source.is_empty() {
        source_stem
    } else {
        normalized_source
    };

    let mut issues = Vec::with_capacity(evaluation.triggered.len());
    let mut unresolved_indicator_count = 0usize;
    let mut comments_unverified_sheets = BTreeSet::new();
    for trigger in &evaluation.triggered {
        let sheet = workbook
            .sheets
            .iter()
            .find(|sheet| sheet.name == trigger.sheet)
            .with_context(|| format!("条件格式结果引用了未知工作表：{}", trigger.sheet))?;
        let (column, row) = parse_cell_address(&trigger.cell)
            .with_context(|| format!("无法解析条件格式触发单元格：{}", trigger.cell))?;
        let check_field = resolve_indicator(sheet, structure_ranges, row, column);
        if check_field.is_empty() {
            unresolved_indicator_count += 1;
        }
        let comment_supported = sheet
            .comments_status
            .starts_with("read_by_namespace_aware_ooxml_supplement");
        if !comment_supported {
            comments_unverified_sheets.insert(sheet.name.clone());
        }
        let check_type = if comment_supported {
            sheet
                .comments
                .get(&trigger.cell)
                .map(|comment| comment.trim().to_owned())
                .unwrap_or_default()
        } else {
            String::new()
        };

        let identity = [
            workbook_key.as_str(),
            trigger.sheet.as_str(),
            trigger.cell.as_str(),
            if check_field.is_empty() {
                "校验指标未识别"
            } else {
                check_field.as_str()
            },
        ]
        .join("｜");
        issues.push(Issue {
            issue_id: identity,
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
            report_code: String::new(),
            sheet_name: trigger.sheet.clone(),
            rule_id: "条件格式填充".to_owned(),
            severity: "条件格式触发".to_owned(),
            formula_cell: trigger.cell.clone(),
            target_cell: trigger.cell.clone(),
            target_value: issue_value(sheet.values.get(&trigger.cell))?,
            formula_result: json!(""),
            message: trigger.message.clone(),
            source_file: source_file.display().to_string(),
            audit_file: workbook_path.display().to_string(),
            check_field,
            check_type: check_type.clone(),
            comparison_value: json!(""),
            reference_value: json!(""),
            difference_value: json!(""),
            detail: trigger.message.clone(),
            institution_feedback: String::new(),
            auditor_opinion: String::new(),
        });
    }

    let issue_build_time_ms = build_started.elapsed().as_millis();
    let context = S1ConditionalContextSummary {
        unresolved_indicator_count,
        comments_unverified_sheet_count: comments_unverified_sheets.len(),
    };
    let has_partial_context =
        context.unresolved_indicator_count > 0 || context.comments_unverified_sheet_count > 0;
    let status = if evaluation.diagnostics.is_empty() && !has_partial_context {
        "completed"
    } else {
        "completed_with_diagnostics"
    };
    Ok(S1ConditionalAuditResult {
        status,
        issues,
        evaluation,
        context,
        issue_build_time_ms,
    })
}

fn issue_value(raw: Option<&Value>) -> Result<Value> {
    let empty = json!({"type":"empty"});
    let value = CellValue::from_probe_json(raw.unwrap_or(&empty))?;
    Ok(match value {
        CellValue::Empty => Value::Null,
        CellValue::Text(value) | CellValue::DateTimeIso(value) | CellValue::DurationIso(value) => {
            json!(value)
        }
        CellValue::Number(value) | CellValue::DateTimeSerial(value) => json!(value),
        CellValue::Integer(value) => json!(value),
        CellValue::Boolean(value) => json!(value),
        CellValue::Error(value) => json!(value),
    })
}

fn resolve_indicator(
    sheet: &SheetDump,
    structure_ranges: &[CopyRange],
    row: i32,
    column: i32,
) -> String {
    let bounds = structure_ranges
        .iter()
        .filter(|area| area.sheet_name == sheet.name)
        .filter_map(|area| parse_cell_range(&area.address))
        .collect::<Vec<_>>();
    let Some(min_row) = bounds.iter().map(|item| item.1).min() else {
        return String::new();
    };
    let Some(min_column) = bounds.iter().map(|item| item.0).min() else {
        return String::new();
    };

    let mut left = Vec::new();
    for current_column in min_column..column {
        if let Some(text) = display_value(sheet, row, current_column) {
            if !left.contains(&text) {
                left.push(text);
            }
        }
    }
    let mut top = Vec::new();
    for current_row in min_row..row {
        if let Some(text) = display_value(sheet, current_row, column) {
            if !top.contains(&text) {
                top.push(text);
            }
        }
    }
    let mut parts = Vec::new();
    if !left.is_empty() {
        parts.push(left.join("_"));
    }
    if !top.is_empty() {
        parts.push(top.join("_"));
    }
    parts.join("｜")
}

fn display_value(sheet: &SheetDump, row: i32, column: i32) -> Option<String> {
    let mut anchor = (row, column);
    for merge in &sheet.merges {
        let Some(bounds) = parse_cell_range(merge) else {
            continue;
        };
        if bounds.0 <= column && column <= bounds.2 && bounds.1 <= row && row <= bounds.3 {
            anchor = (bounds.1, bounds.0);
            break;
        }
    }
    let address = cell_address(anchor.1, anchor.0);
    let raw = sheet.values.get(&address)?;
    let value = CellValue::from_probe_json(raw).ok()?;
    let text = match value {
        CellValue::Text(value) => value.trim().to_owned(),
        CellValue::Boolean(true) => "True".to_owned(),
        CellValue::Boolean(false) => "False".to_owned(),
        CellValue::Empty
        | CellValue::Number(_)
        | CellValue::Integer(_)
        | CellValue::DateTimeSerial(_)
        | CellValue::DateTimeIso(_)
        | CellValue::DurationIso(_)
        | CellValue::Error(_) => return None,
    };
    if text.is_empty() || is_numeric_text(&text) {
        return None;
    }
    Some(text)
}

fn is_numeric_text(value: &str) -> bool {
    let compact = value.replace(',', "");
    let digits = if compact.starts_with('+') || compact.starts_with('-') {
        &compact[1..]
    } else {
        &compact
    };
    if digits.is_empty() {
        return false;
    }
    let mut parts = digits.split('.');
    let whole = parts.next().unwrap_or_default();
    let fraction = parts.next();
    if parts.next().is_some() || whole.is_empty() || !whole.chars().all(|ch| ch.is_ascii_digit()) {
        return false;
    }
    fraction.map_or(true, |fraction| {
        !fraction.is_empty() && fraction.chars().all(|ch| ch.is_ascii_digit())
    })
}

fn parse_cell_address(value: &str) -> Option<(i32, i32)> {
    let address = value.replace('$', "");
    let split = address.find(|ch: char| ch.is_ascii_digit())?;
    if split == 0 {
        return None;
    }
    let (letters, digits) = address.split_at(split);
    if !letters.chars().all(|ch| ch.is_ascii_alphabetic())
        || !digits.chars().all(|ch| ch.is_ascii_digit())
    {
        return None;
    }
    let mut column = 0i32;
    for ch in letters.chars() {
        column = column * 26 + (ch.to_ascii_uppercase() as i32 - 'A' as i32 + 1);
    }
    let row = digits.parse::<i32>().ok()?;
    (column > 0 && row > 0).then_some((column, row))
}

fn cell_address(column: i32, row: i32) -> String {
    let mut number = column;
    let mut letters = String::new();
    while number > 0 {
        number -= 1;
        letters.insert(0, char::from(b'A' + (number % 26) as u8));
        number /= 26;
    }
    format!("{letters}{row}")
}

/// Parse one A1 cell or range and return (min column, min row, max column,
/// max row), matching the 1-based coordinates used by the Python resolver.
fn parse_cell_range(value: &str) -> Option<(i32, i32, i32, i32)> {
    let mut endpoints = value.split(':');
    let first = parse_cell_address(endpoints.next()?.trim())?;
    let last = match endpoints.next() {
        Some(value) => parse_cell_address(value.trim())?,
        None => first,
    };
    if endpoints.next().is_some() {
        return None;
    }
    Some((
        first.0.min(last.0),
        first.1.min(last.1),
        first.0.max(last.0),
        first.1.max(last.1),
    ))
}

#[cfg(test)]
mod tests {
    use super::{resolve_indicator, CopyRange};
    use excel_reader::SheetDump;
    use serde_json::json;

    #[test]
    fn resolves_indicators_using_structure_ranges_and_merged_anchors() {
        let sheet = SheetDump {
            name: "Data".to_owned(),
            values: [
                ("A1".to_owned(), json!({"type":"string","value":"贷款余额"})),
                ("B1".to_owned(), json!({"type":"string","value":"本期"})),
                ("B2".to_owned(), json!({"type":"number","value":1000.0})),
            ]
            .into_iter()
            .collect(),
            formulas: Default::default(),
            formula_status: "read_by_calamine".to_owned(),
            merges: vec!["A1:A3".to_owned()],
            merge_status: "read_by_calamine".to_owned(),
            comments: Default::default(),
            comments_status: "read_by_namespace_aware_ooxml_supplement".to_owned(),
            conditional_formatting: Vec::new(),
            conditional_formatting_status: "read_by_namespace_aware_ooxml_supplement".to_owned(),
            data_validations: Vec::new(),
            data_validations_status: "read_by_namespace_aware_ooxml_supplement".to_owned(),
        };
        let structure_ranges = [CopyRange {
            sheet_name: "Data".to_owned(),
            address: "A1:B4".to_owned(),
        }];

        assert_eq!(
            resolve_indicator(&sheet, &structure_ranges, 3, 2),
            "贷款余额｜本期"
        );
    }

    #[test]
    fn numeric_display_text_is_not_used_as_an_indicator() {
        let sheet = SheetDump {
            name: "Data".to_owned(),
            values: [
                ("A1".to_owned(), json!({"type":"string","value":"1,234.5"})),
                ("A2".to_owned(), json!({"type":"string","value":"余额"})),
            ]
            .into_iter()
            .collect(),
            formulas: Default::default(),
            formula_status: "read_by_calamine".to_owned(),
            merges: Vec::new(),
            merge_status: "read_by_calamine".to_owned(),
            comments: Default::default(),
            comments_status: "read_by_namespace_aware_ooxml_supplement".to_owned(),
            conditional_formatting: Vec::new(),
            conditional_formatting_status: "read_by_namespace_aware_ooxml_supplement".to_owned(),
            data_validations: Vec::new(),
            data_validations_status: "read_by_namespace_aware_ooxml_supplement".to_owned(),
        };
        let structure_ranges = [CopyRange {
            sheet_name: "Data".to_owned(),
            address: "A1:B3".to_owned(),
        }];

        assert_eq!(resolve_indicator(&sheet, &structure_ranges, 2, 2), "余额");
    }
}
