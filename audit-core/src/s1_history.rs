//! S1 issue history classification and manual feedback carry-over.

use crate::models::Issue;
use anyhow::{ensure, Context, Result};
use excel_reader::{read_value_columns_with_iso_dates, CellValue, SheetValues};
use serde::{Deserialize, Serialize};
use std::{
    collections::{BTreeMap, HashMap, HashSet},
    path::Path,
    time::Instant,
};

const HISTORY_AUDIT_SHEET: &str = "核查表校验结果";
const LEGACY_HISTORY_SHEET: &str = "历史审核结果";
const INITIAL_HISTORY_READ_COLUMNS: usize = 14;

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
