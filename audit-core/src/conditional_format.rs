//! Static conditional-format evaluation used by the Audit S1 migration.
//!
//! This is a Rust port of the source project's bounded OOXML rule evaluator.
//! It evaluates cellIs comparisons and the supported expression subset over
//! stored/cached cell values. Unsupported rule types and formula syntax remain
//! explicit diagnostics; they are never treated as FALSE.

use anyhow::{bail, Context, Result};
use excel_reader::{read_workbook, SheetDump, WorkbookDump};
use serde::Serialize;
use serde_json::Value as JsonValue;
use std::{
    collections::{BTreeMap, BTreeSet},
    path::Path,
    time::Instant,
};

#[derive(Clone, Debug, PartialEq)]
enum Value {
    Empty,
    Number(f64),
    Text(String),
    Boolean(bool),
    Range(RangeRef),
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
struct CellRef {
    column: i32,
    row: i32,
    column_absolute: bool,
    row_absolute: bool,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
struct RangeRef {
    first: CellRef,
    last: CellRef,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
struct Bounds {
    min_column: i32,
    min_row: i32,
    max_column: i32,
    max_row: i32,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize)]
#[serde(rename_all = "SCREAMING_SNAKE_CASE")]
pub enum EvaluationStatus {
    True,
    False,
    Unsupported,
    Error,
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct RuleEvaluation {
    pub status: EvaluationStatus,
    #[serde(skip_serializing_if = "String::is_empty")]
    pub reason: String,
}

impl RuleEvaluation {
    fn new(status: EvaluationStatus) -> Self {
        Self {
            status,
            reason: String::new(),
        }
    }

    fn with_reason(status: EvaluationStatus, reason: impl Into<String>) -> Self {
        Self {
            status,
            reason: reason.into(),
        }
    }
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct ConditionalTrigger {
    pub sheet: String,
    pub cell: String,
    pub applies_to: String,
    pub rule_type: String,
    pub operator: String,
    pub formula: String,
    pub message: String,
    pub priority: u32,
    pub confidence: String,
    #[serde(skip_serializing_if = "String::is_empty")]
    pub unsupported_reason: String,
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct ConditionalDiagnostic {
    pub sheet: String,
    pub cell: String,
    pub applies_to: String,
    pub rule_type: String,
    pub formula: String,
    pub priority: u32,
    pub status: EvaluationStatus,
    pub reason: String,
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct ConditionalEvaluationReport {
    pub status: &'static str,
    pub reader: &'static str,
    pub read_time_ms: u128,
    pub evaluation_time_ms: u128,
    pub total_time_ms: u128,
    pub rules_total: usize,
    pub rules_supported: usize,
    pub scanned_cells: usize,
    pub triggered: Vec<ConditionalTrigger>,
    pub diagnostics: Vec<ConditionalDiagnostic>,
}

#[derive(Clone, Debug)]
struct Rule {
    sheet: String,
    applies_to: String,
    rule_type: String,
    operator: String,
    formulas: Vec<String>,
    priority: u32,
    order: usize,
    anchor_row: i32,
    anchor_column: i32,
    bounds: Vec<Bounds>,
}

#[derive(Clone, Debug)]
struct ScanArea {
    sheet: String,
    bounds: Bounds,
}

#[derive(Clone, Debug)]
enum EvalError {
    Unsupported(String),
    Error(String),
}

impl EvalError {
    fn status(&self) -> EvaluationStatus {
        match self {
            Self::Unsupported(_) => EvaluationStatus::Unsupported,
            Self::Error(_) => EvaluationStatus::Error,
        }
    }

    fn reason(&self) -> &str {
        match self {
            Self::Unsupported(reason) | Self::Error(reason) => reason,
        }
    }
}

/// Read a workbook through the shared Rust reader and evaluate its static
/// conditional-format rules. scan_areas entries use Sheet!A1:B20; when empty,
/// the rules' AppliesTo ranges define the scan area.
///
/// This does not recalculate formula cells. Calamine's stored/cached values
/// are used as read. Callers requiring fresh formula results must explicitly
/// run the formula-recalculation sidecar first and validate its cache refresh.
pub fn evaluate_xlsx_workbook(
    path: impl AsRef<Path>,
    scan_areas: &[String],
) -> Result<ConditionalEvaluationReport> {
    let (report, _) = evaluate_xlsx_workbook_with_dump(path, scan_areas)?;
    Ok(report)
}

/// Evaluate conditional formatting and return the exact workbook projection
/// used by the evaluator so Audit can construct business findings without
/// opening and parsing the input a second time.
pub fn evaluate_xlsx_workbook_with_dump(
    path: impl AsRef<Path>,
    scan_areas: &[String],
) -> Result<(ConditionalEvaluationReport, WorkbookDump)> {
    let total_started = Instant::now();
    let path = path.as_ref();
    if !path
        .extension()
        .and_then(|ext| ext.to_str())
        .is_some_and(|ext| matches!(ext.to_ascii_lowercase().as_str(), "xlsx" | "xlsm"))
    {
        bail!("条件格式规则求值目前只支持 OOXML 工作簿（.xlsx/.xlsm）");
    }
    let read_started = Instant::now();
    let workbook = read_workbook(path).with_context(|| format!("读取工作簿 {}", path.display()))?;
    let read_time_ms = read_started.elapsed().as_millis();
    let report = evaluate_dump(
        &workbook.sheets,
        scan_areas,
        read_time_ms,
        total_started.elapsed().as_millis(),
    )?;
    Ok((report, workbook))
}

fn evaluate_dump(
    sheets: &[SheetDump],
    scan_area_specs: &[String],
    read_time_ms: u128,
    elapsed_before_eval_ms: u128,
) -> Result<ConditionalEvaluationReport> {
    let total_started = Instant::now();
    let mut rules = Vec::new();
    let mut sheet_values = BTreeMap::new();
    for sheet in sheets {
        if !sheet
            .conditional_formatting_status
            .starts_with("read_by_namespace_aware_ooxml_supplement")
            && !sheet.conditional_formatting.is_empty()
        {
            bail!(
                "工作表 {} 的条件格式读取状态未验证：{}",
                sheet.name,
                sheet.conditional_formatting_status
            );
        }
        let mut values = BTreeMap::new();
        for (address, raw) in &sheet.values {
            values.insert(address.to_ascii_uppercase(), value_from_reader(raw)?);
        }
        sheet_values.insert(sheet.name.clone(), values);
        for node in &sheet.conditional_formatting {
            parse_conditional_node(&sheet.name, node, rules.len(), &mut rules)
                .with_context(|| format!("解析工作表 {} 的条件格式元数据", sheet.name))?;
        }
    }
    let rules_total = rules.len();
    let rules_supported = rules
        .iter()
        .filter(|rule| matches!(rule.rule_type.as_str(), "cellis" | "expression"))
        .count();
    rules.sort_by(|left, right| {
        left.sheet
            .cmp(&right.sheet)
            .then(left.priority.cmp(&right.priority))
            .then(left.order.cmp(&right.order))
    });

    let scan_areas = if scan_area_specs.is_empty() {
        rules
            .iter()
            .flat_map(|rule| {
                rule.bounds.iter().copied().map(|bounds| ScanArea {
                    sheet: rule.sheet.clone(),
                    bounds,
                })
            })
            .collect::<Vec<_>>()
    } else {
        scan_area_specs
            .iter()
            .map(|spec| parse_scan_area(spec))
            .collect::<Result<Vec<_>>>()?
    };

    let known_sheets = sheets
        .iter()
        .map(|sheet| sheet.name.as_str())
        .collect::<BTreeSet<_>>();
    for area in &scan_areas {
        if !known_sheets.contains(area.sheet.as_str()) {
            bail!("扫描范围引用了不存在的工作表：{}", area.sheet);
        }
    }

    let mut visited = BTreeSet::new();
    let mut triggered = Vec::new();
    let mut diagnostics = Vec::new();
    for area in &scan_areas {
        let sheet_rules = rules
            .iter()
            .filter(|rule| rule.sheet == area.sheet)
            .collect::<Vec<_>>();
        let Some(values) = sheet_values.get(&area.sheet) else {
            continue;
        };
        for row in area.bounds.min_row..=area.bounds.max_row {
            for column in area.bounds.min_column..=area.bounds.max_column {
                if !visited.insert((area.sheet.clone(), row, column)) {
                    continue;
                }
                let address = cell_address(column, row)?;
                let cell_value = values.get(&address).cloned().unwrap_or(Value::Empty);
                let mut partial_reason = String::new();
                for rule in &sheet_rules {
                    if !rule
                        .bounds
                        .iter()
                        .any(|bounds| contains(*bounds, column, row))
                    {
                        continue;
                    }
                    let formula = rule.formulas.first().cloned().unwrap_or_default();
                    if !matches!(rule.rule_type.as_str(), "cellis" | "expression") {
                        let reason = format!("暂不支持的条件格式类型：{}", rule.rule_type);
                        if partial_reason.is_empty() {
                            partial_reason = reason.clone();
                        }
                        diagnostics.push(ConditionalDiagnostic {
                            sheet: area.sheet.clone(),
                            cell: address.clone(),
                            applies_to: rule.applies_to.clone(),
                            rule_type: rule.rule_type.clone(),
                            formula,
                            priority: rule.priority,
                            status: EvaluationStatus::Unsupported,
                            reason,
                        });
                        continue;
                    }

                    let grid = ValueGrid { values };
                    let evaluation = evaluate_rule(rule, &cell_value, &grid, row, column);
                    match evaluation.status {
                        EvaluationStatus::True => {
                            triggered.push(ConditionalTrigger {
                                sheet: area.sheet.clone(),
                                cell: address.clone(),
                                applies_to: rule.applies_to.clone(),
                                rule_type: rule.rule_type.clone(),
                                operator: rule.operator.clone(),
                                formula,
                                message: trigger_message(&rule, &address, row, column),
                                priority: rule.priority,
                                confidence: if partial_reason.is_empty() {
                                    "FULL".to_owned()
                                } else {
                                    "PARTIAL".to_owned()
                                },
                                unsupported_reason: partial_reason,
                            });
                            break;
                        }
                        EvaluationStatus::False => {}
                        EvaluationStatus::Unsupported | EvaluationStatus::Error => {
                            if partial_reason.is_empty()
                                && evaluation.status == EvaluationStatus::Unsupported
                            {
                                partial_reason.clone_from(&evaluation.reason);
                            }
                            diagnostics.push(ConditionalDiagnostic {
                                sheet: area.sheet.clone(),
                                cell: address.clone(),
                                applies_to: rule.applies_to.clone(),
                                rule_type: rule.rule_type.clone(),
                                formula,
                                priority: rule.priority,
                                status: evaluation.status,
                                reason: evaluation.reason,
                            });
                        }
                    }
                }
            }
        }
    }
    let evaluation_time_ms = total_started.elapsed().as_millis();
    Ok(ConditionalEvaluationReport {
        status: "evaluated_static_rules",
        reader: "calamine+namespace_aware_ooxml",
        read_time_ms,
        evaluation_time_ms,
        total_time_ms: elapsed_before_eval_ms.saturating_add(evaluation_time_ms),
        rules_total,
        rules_supported,
        scanned_cells: visited.len(),
        triggered,
        diagnostics,
    })
}

fn parse_conditional_node(
    sheet: &str,
    node: &JsonValue,
    first_order: usize,
    rules: &mut Vec<Rule>,
) -> Result<()> {
    let sqref = attribute(node, "sqref").unwrap_or_default();
    let bounds = parse_bounds_list(sqref)?;
    if bounds.is_empty() {
        bail!("条件格式 sqref 为空或无法解析");
    }
    let anchor_column = bounds.iter().map(|item| item.min_column).min().unwrap_or(1);
    let anchor_row = bounds.iter().map(|item| item.min_row).min().unwrap_or(1);
    let children = node
        .get("content")
        .and_then(JsonValue::as_array)
        .context("conditionalFormatting 缺少 content 数组")?;
    let mut local_order = 0usize;
    for child in children {
        if child.get("name").and_then(JsonValue::as_str) != Some("cfRule") {
            continue;
        }
        let rule_type = attribute(child, "type")
            .unwrap_or_default()
            .to_ascii_lowercase();
        let operator = attribute(child, "operator").unwrap_or_default().to_owned();
        let priority = attribute(child, "priority")
            .and_then(|text| text.parse::<u32>().ok())
            .unwrap_or((first_order + local_order + 1) as u32);
        let formulas = child
            .get("content")
            .and_then(JsonValue::as_array)
            .into_iter()
            .flatten()
            .filter(|item| item.get("name").and_then(JsonValue::as_str) == Some("formula"))
            .map(node_text)
            .collect::<Vec<_>>();
        rules.push(Rule {
            sheet: sheet.to_owned(),
            applies_to: sqref.to_owned(),
            rule_type,
            operator,
            formulas,
            priority,
            order: first_order + local_order,
            anchor_row,
            anchor_column,
            bounds: bounds.clone(),
        });
        local_order += 1;
    }
    Ok(())
}

fn attribute<'a>(node: &'a JsonValue, key: &str) -> Option<&'a str> {
    node.get("attributes")?.get(key)?.as_str()
}

fn node_text(node: &JsonValue) -> String {
    node.get("content")
        .and_then(JsonValue::as_array)
        .into_iter()
        .flatten()
        .filter_map(|child| child.get("text").and_then(JsonValue::as_str))
        .collect::<String>()
}

fn parse_scan_area(spec: &str) -> Result<ScanArea> {
    let (sheet, address) = spec
        .rsplit_once('!')
        .with_context(|| format!("扫描范围格式应为 Sheet!A1:B10：{spec}"))?;
    if sheet.is_empty() {
        bail!("扫描范围缺少工作表名称：{spec}");
    }
    let bounds = parse_bounds(address).with_context(|| format!("无效的扫描范围地址：{spec}"))?;
    Ok(ScanArea {
        sheet: sheet.trim_matches('\'').replace("''", "'"),
        bounds,
    })
}

fn parse_bounds_list(sqref: &str) -> Result<Vec<Bounds>> {
    sqref.split_whitespace().map(parse_bounds).collect()
}

fn parse_bounds(address: &str) -> Result<Bounds> {
    let mut parts = address.split(':');
    let first = parse_cell_ref(parts.next().unwrap_or_default())?;
    let last = match parts.next() {
        Some(value) => parse_cell_ref(value)?,
        None => first,
    };
    if parts.next().is_some() {
        bail!("区域包含多个冒号");
    }
    Ok(Bounds {
        min_column: first.column.min(last.column),
        min_row: first.row.min(last.row),
        max_column: first.column.max(last.column),
        max_row: first.row.max(last.row),
    })
}

fn parse_cell_ref(raw: &str) -> Result<CellRef> {
    let chars = raw.chars().collect::<Vec<_>>();
    let mut pos = 0usize;
    let column_absolute = chars.get(pos) == Some(&'$');
    if column_absolute {
        pos += 1;
    }
    let column_start = pos;
    while chars.get(pos).is_some_and(|ch| ch.is_ascii_alphabetic()) && pos - column_start < 3 {
        pos += 1;
    }
    if pos == column_start || chars.get(pos).is_some_and(|ch| ch.is_ascii_alphabetic()) {
        bail!("单元格引用列格式无效：{raw}");
    }
    let mut column = 0i32;
    for ch in chars[column_start..pos].iter().copied() {
        column = column * 26 + (ch.to_ascii_uppercase() as i32 - 'A' as i32 + 1);
    }
    let row_absolute = chars.get(pos) == Some(&'$');
    if row_absolute {
        pos += 1;
    }
    let row_start = pos;
    while chars.get(pos).is_some_and(|ch| ch.is_ascii_digit()) {
        pos += 1;
    }
    if row_start == pos || pos != chars.len() {
        bail!("单元格引用行格式无效：{raw}");
    }
    let row = chars[row_start..pos]
        .iter()
        .collect::<String>()
        .parse::<i32>()
        .with_context(|| format!("行号无效：{raw}"))?;
    if row < 1 || row > 1_048_576 || column < 1 || column > 16_384 {
        bail!("单元格引用超出 Excel 工作表范围：{raw}");
    }
    Ok(CellRef {
        column,
        row,
        column_absolute,
        row_absolute,
    })
}

fn cell_address(column: i32, row: i32) -> Result<String> {
    if column < 1 || row < 1 {
        bail!("单元格坐标越界");
    }
    let mut number = column;
    let mut letters = String::new();
    while number > 0 {
        number -= 1;
        letters.insert(0, char::from(b'A' + (number % 26) as u8));
        number /= 26;
    }
    Ok(format!("{letters}{row}"))
}

fn trigger_message(rule: &Rule, address: &str, row: i32, column: i32) -> String {
    let raw_formula = rule
        .formulas
        .first()
        .map(|formula| formula.trim_start_matches('='))
        .unwrap_or_default();
    let formula = translate_formula_refs(
        raw_formula,
        row - rule.anchor_row,
        column - rule.anchor_column,
    )
    .replace('$', "");
    match rule.rule_type.as_str() {
        "cellis" => {
            let operator = match rule.operator.to_ascii_lowercase().as_str() {
                "equal" => "=",
                "notequal" => "<>",
                "greaterthan" => ">",
                "greaterthanorequal" => ">=",
                "lessthan" => "<",
                "lessthanorequal" => "<=",
                _ if !rule.operator.is_empty() => rule.operator.as_str(),
                _ => "?",
            };
            format!(
                "条件格式规则：{}{operator}{formula}",
                address.replace('$', "")
            )
        }
        "expression" => format!("条件格式规则：{formula}"),
        _ => format!("条件格式规则：{}", rule.rule_type),
    }
}

/// Shift cell references in a displayed conditional-format formula using the
/// same relative/absolute reference rules used when evaluating the rule.
fn translate_formula_refs(formula: &str, row_delta: i32, column_delta: i32) -> String {
    if row_delta == 0 && column_delta == 0 {
        return formula.to_owned();
    }

    let chars = formula.chars().collect::<Vec<_>>();
    let mut output = String::with_capacity(formula.len());
    let mut position = 0usize;
    while position < chars.len() {
        if chars[position] == '"' {
            output.push('"');
            position += 1;
            while position < chars.len() {
                let current = chars[position];
                output.push(current);
                position += 1;
                if current == '"' {
                    if chars.get(position) == Some(&'"') {
                        output.push('"');
                        position += 1;
                    } else {
                        break;
                    }
                }
            }
            continue;
        }

        let previous_is_word = position > 0 && is_formula_word(chars[position - 1]);
        if !previous_is_word {
            if let Some((first, first_end)) = parse_formula_ref(&chars, position) {
                let (last, end) = if chars.get(first_end) == Some(&':') {
                    parse_formula_ref(&chars, first_end + 1)
                        .filter(|(_, second_end)| *second_end > first_end + 1)
                        .map_or((None, first_end), |(cell, second_end)| {
                            (Some(cell), second_end)
                        })
                } else {
                    (None, first_end)
                };
                let next_non_space = chars[end..]
                    .iter()
                    .position(|ch| !ch.is_whitespace())
                    .map(|offset| chars[end + offset]);
                let has_valid_end = chars.get(end).map_or(true, |ch| !is_formula_word(*ch))
                    && next_non_space != Some('(');
                if has_valid_end {
                    output.push_str(&shift_formula_ref(first, row_delta, column_delta));
                    if let Some(last) = last {
                        output.push(':');
                        output.push_str(&shift_formula_ref(last, row_delta, column_delta));
                    }
                    position = end;
                    continue;
                }
            }
        }
        output.push(chars[position]);
        position += 1;
    }
    output
}

fn parse_formula_ref(chars: &[char], start: usize) -> Option<(CellRef, usize)> {
    let mut position = start;
    let column_absolute = chars.get(position) == Some(&'$');
    if column_absolute {
        position += 1;
    }
    let column_start = position;
    while chars
        .get(position)
        .is_some_and(|ch| ch.is_ascii_alphabetic())
        && position - column_start < 3
    {
        position += 1;
    }
    if position == column_start
        || chars
            .get(position)
            .is_some_and(|ch| ch.is_ascii_alphabetic())
    {
        return None;
    }
    let mut column = 0i32;
    for ch in &chars[column_start..position] {
        column = column * 26 + (ch.to_ascii_uppercase() as i32 - 'A' as i32 + 1);
    }

    let row_absolute = chars.get(position) == Some(&'$');
    if row_absolute {
        position += 1;
    }
    let row_start = position;
    while chars.get(position).is_some_and(|ch| ch.is_ascii_digit()) {
        position += 1;
    }
    if position == row_start {
        return None;
    }
    let row = chars[row_start..position]
        .iter()
        .collect::<String>()
        .parse::<i32>()
        .ok()?;
    if column < 1 || row < 1 {
        return None;
    }
    Some((
        CellRef {
            column,
            row,
            column_absolute,
            row_absolute,
        },
        position,
    ))
}

fn shift_formula_ref(cell: CellRef, row_delta: i32, column_delta: i32) -> String {
    let row = if cell.row_absolute {
        cell.row
    } else {
        cell.row + row_delta
    };
    let column = if cell.column_absolute {
        cell.column
    } else {
        cell.column + column_delta
    };
    if row < 1 || column < 1 {
        return "#REF!".to_owned();
    }
    let mut address = cell_address(column, row).unwrap_or_else(|_| "#REF!".to_owned());
    if cell.column_absolute {
        address.insert(0, '$');
    }
    let column_end = address
        .chars()
        .take_while(|ch| ch.is_ascii_alphabetic() || *ch == '$')
        .count();
    if cell.row_absolute {
        address.insert(column_end, '$');
    }
    address
}

fn is_formula_word(ch: char) -> bool {
    ch.is_ascii_alphanumeric() || ch == '_'
}

fn contains(bounds: Bounds, column: i32, row: i32) -> bool {
    bounds.min_column <= column
        && column <= bounds.max_column
        && bounds.min_row <= row
        && row <= bounds.max_row
}

fn value_from_reader(raw: &JsonValue) -> Result<Value> {
    let kind = raw
        .get("type")
        .and_then(JsonValue::as_str)
        .context("Calamine 单元格值缺少 type")?;
    let value = raw.get("value");
    Ok(match kind {
        "empty" => Value::Empty,
        "string" => Value::Text(
            value
                .and_then(JsonValue::as_str)
                .unwrap_or_default()
                .to_owned(),
        ),
        "number" | "datetime_serial" => {
            Value::Number(value.and_then(JsonValue::as_f64).unwrap_or(f64::NAN))
        }
        "integer" => Value::Number(value.and_then(JsonValue::as_i64).unwrap_or_default() as f64),
        "boolean" => Value::Boolean(value.and_then(JsonValue::as_bool).unwrap_or(false)),
        "datetime_iso" | "duration_iso" | "error" => Value::Text(
            value
                .and_then(JsonValue::as_str)
                .unwrap_or_default()
                .to_owned(),
        ),
        _ => bail!("不支持的 Calamine 值类型：{kind}"),
    })
}

fn evaluate_rule(
    rule: &Rule,
    value: &Value,
    grid: &ValueGrid<'_>,
    row: i32,
    column: i32,
) -> RuleEvaluation {
    match rule.rule_type.as_str() {
        "cellis" => {
            let resolve = |index| resolve_cellis_operand(rule, index, grid, row, column);
            evaluate_cell_is(&rule.operator, &rule.formulas, value, resolve)
        }
        "expression" => {
            let formula = rule
                .formulas
                .first()
                .map(String::as_str)
                .unwrap_or_default();
            if formula.trim().is_empty() {
                return RuleEvaluation::with_reason(
                    EvaluationStatus::Error,
                    "expression 规则公式为空",
                );
            }
            match evaluate_expression(
                formula,
                grid,
                rule.anchor_row,
                rule.anchor_column,
                row,
                column,
            ) {
                Ok(value) => RuleEvaluation::new(if truthy(&value) {
                    EvaluationStatus::True
                } else {
                    EvaluationStatus::False
                }),
                Err(error) => RuleEvaluation::with_reason(error.status(), error.reason()),
            }
        }
        _ => RuleEvaluation::with_reason(
            EvaluationStatus::Unsupported,
            format!("暂不支持的条件格式类型：{}", rule.rule_type),
        ),
    }
}

fn evaluate_cell_is<F>(
    operator: &str,
    _formulas: &[String],
    value: &Value,
    mut resolve_operand: F,
) -> RuleEvaluation
where
    F: FnMut(usize) -> std::result::Result<Value, EvalError>,
{
    let op = operator.trim().to_ascii_lowercase();
    if !matches!(
        op.as_str(),
        "equal"
            | "notequal"
            | "greaterthan"
            | "greaterthanorequal"
            | "lessthan"
            | "lessthanorequal"
            | "between"
            | "notbetween"
    ) {
        return RuleEvaluation::with_reason(
            EvaluationStatus::Unsupported,
            format!(
                "未知 cellIs 运算符：{}",
                if operator.is_empty() { "空" } else { operator }
            ),
        );
    }
    if matches!(value, Value::Empty) || matches!(value, Value::Text(text) if text.is_empty()) {
        return RuleEvaluation::new(EvaluationStatus::False);
    }
    let first = match resolve_operand(0) {
        Ok(value) => value,
        Err(error) => return RuleEvaluation::with_reason(error.status(), error.reason()),
    };
    let second = if matches!(op.as_str(), "between" | "notbetween") {
        match resolve_operand(1) {
            Ok(value) => Some(value),
            Err(error) => return RuleEvaluation::with_reason(error.status(), error.reason()),
        }
    } else {
        None
    };
    if matches!(first, Value::Empty) {
        return RuleEvaluation::with_reason(
            EvaluationStatus::Unsupported,
            "cellIs 操作数无法解析为字面量或单格引用",
        );
    }

    let value_number = expression_number(value);
    let first_number = expression_number(&first);
    if let (Some(value_number), Some(first_number)) = (value_number, first_number) {
        let value_number = excel_round15(value_number);
        let first_number = excel_round15(first_number);
        let second_number = second
            .as_ref()
            .and_then(expression_number)
            .map(excel_round15);
        if matches!(op.as_str(), "between" | "notbetween") {
            let Some(second_number) = second_number else {
                return RuleEvaluation::with_reason(
                    EvaluationStatus::Unsupported,
                    format!("cellIs {op} 第二操作数不是数值"),
                );
            };
            let inside = first_number <= value_number && value_number <= second_number;
            return RuleEvaluation::new(if inside == (op == "between") {
                EvaluationStatus::True
            } else {
                EvaluationStatus::False
            });
        }
        return RuleEvaluation::new(if compare_ordered(value_number, first_number, &op) {
            EvaluationStatus::True
        } else {
            EvaluationStatus::False
        });
    }
    if let (Value::Text(value), Value::Text(first)) = (value, &first) {
        if matches!(op.as_str(), "between" | "notbetween") {
            return RuleEvaluation::with_reason(
                EvaluationStatus::Unsupported,
                "cellIs between 暂不支持文本比较",
            );
        }
        return RuleEvaluation::new(if compare_text(value, first, &op) {
            EvaluationStatus::True
        } else {
            EvaluationStatus::False
        });
    }
    if let (Value::Boolean(value), Value::Boolean(first)) = (value, &first) {
        return RuleEvaluation::new(
            if compare_ordered(i32::from(*value) as f64, i32::from(*first) as f64, &op) {
                EvaluationStatus::True
            } else {
                EvaluationStatus::False
            },
        );
    }
    RuleEvaluation::new(EvaluationStatus::False)
}

fn compare_ordered(left: f64, right: f64, op: &str) -> bool {
    match op {
        "equal" => left == right,
        "notequal" => left != right,
        "greaterthan" => left > right,
        "greaterthanorequal" => left >= right,
        "lessthan" => left < right,
        "lessthanorequal" => left <= right,
        _ => false,
    }
}

fn compare_text(left: &str, right: &str, op: &str) -> bool {
    match op {
        "equal" => left == right,
        "notequal" => left != right,
        "greaterthan" => left > right,
        "greaterthanorequal" => left >= right,
        "lessthan" => left < right,
        "lessthanorequal" => left <= right,
        _ => false,
    }
}

fn resolve_cellis_operand(
    rule: &Rule,
    index: usize,
    grid: &ValueGrid<'_>,
    row: i32,
    column: i32,
) -> std::result::Result<Value, EvalError> {
    let raw = rule
        .formulas
        .get(index)
        .map(|text| text.trim())
        .unwrap_or_default();
    if raw.is_empty() {
        return Err(EvalError::Unsupported("cellIs 操作数为空".to_owned()));
    }
    if let Some(number) = parse_number(raw) {
        return Ok(Value::Number(number));
    }
    if raw.eq_ignore_ascii_case("true") || raw.eq_ignore_ascii_case("false") {
        return Ok(Value::Boolean(raw.eq_ignore_ascii_case("true")));
    }
    if let Ok(cell) = parse_cell_ref(raw) {
        let shifted = shift_cell(cell, row - rule.anchor_row, column - rule.anchor_column)?;
        return grid.cell(shifted.column, shifted.row);
    }
    Err(EvalError::Unsupported(format!(
        "cellIs 操作数无法解析：{raw}"
    )))
}

struct ValueGrid<'a> {
    values: &'a BTreeMap<String, Value>,
}

impl ValueGrid<'_> {
    fn cell(&self, column: i32, row: i32) -> std::result::Result<Value, EvalError> {
        let address =
            cell_address(column, row).map_err(|error| EvalError::Error(error.to_string()))?;
        Ok(self.values.get(&address).cloned().unwrap_or(Value::Empty))
    }

    fn range(&self, range: RangeRef, dr: i32, dc: i32) -> std::result::Result<Value, EvalError> {
        let first = shift_cell(range.first, dr, dc)?;
        let last = shift_cell(range.last, dr, dc)?;
        let _ = cell_address(first.column, first.row)
            .and_then(|_| cell_address(last.column, last.row))
            .map_err(|error| EvalError::Error(error.to_string()))?;
        if first.column == last.column && first.row == last.row {
            return self.cell(first.column, first.row);
        }
        Ok(Value::Range(RangeRef { first, last }))
    }

    fn range_values(&self, range: RangeRef) -> std::result::Result<Vec<Value>, EvalError> {
        let min_column = range.first.column.min(range.last.column);
        let max_column = range.first.column.max(range.last.column);
        let min_row = range.first.row.min(range.last.row);
        let max_row = range.first.row.max(range.last.row);
        let mut values = Vec::new();
        for row in min_row..=max_row {
            for column in min_column..=max_column {
                values.push(self.cell(column, row)?);
            }
        }
        Ok(values)
    }
}

fn shift_cell(cell: CellRef, dr: i32, dc: i32) -> std::result::Result<CellRef, EvalError> {
    let row = if cell.row_absolute {
        cell.row
    } else {
        cell.row + dr
    };
    let column = if cell.column_absolute {
        cell.column
    } else {
        cell.column + dc
    };
    if row < 1 || row > 1_048_576 || column < 1 || column > 16_384 {
        return Err(EvalError::Error("条件格式引用越界".to_owned()));
    }
    Ok(CellRef {
        row,
        column,
        ..cell
    })
}

fn evaluate_expression(
    formula: &str,
    grid: &ValueGrid<'_>,
    anchor_row: i32,
    anchor_column: i32,
    current_row: i32,
    current_column: i32,
) -> std::result::Result<Value, EvalError> {
    let tokens = tokenize(formula)?;
    let mut parser = Parser {
        tokens,
        position: 0,
        formula,
        grid,
        dr: current_row - anchor_row,
        dc: current_column - anchor_column,
    };
    let value = parser.expression()?;
    if parser.peek() != &Token::End {
        return Err(EvalError::Error(format!(
            "无法解析的条件格式公式：{formula}"
        )));
    }
    Ok(value)
}

#[derive(Clone, Debug, PartialEq)]
enum Token {
    Number(String),
    String(String),
    Reference(RangeRef),
    Name(String),
    Operator(String),
    End,
}

fn tokenize(formula: &str) -> std::result::Result<Vec<Token>, EvalError> {
    let chars = formula.chars().collect::<Vec<_>>();
    let mut tokens = Vec::new();
    let mut pos = 0usize;
    while pos < chars.len() {
        if chars[pos].is_whitespace() {
            pos += 1;
            continue;
        }
        if chars[pos].is_ascii_digit() {
            let start = pos;
            pos += 1;
            while pos < chars.len() && chars[pos].is_ascii_digit() {
                pos += 1;
            }
            if chars.get(pos) == Some(&'.') {
                pos += 1;
                while pos < chars.len() && chars[pos].is_ascii_digit() {
                    pos += 1;
                }
            }
            tokens.push(Token::Number(chars[start..pos].iter().collect()));
            continue;
        }
        if chars[pos] == '"' {
            pos += 1;
            let start = pos;
            while pos < chars.len() && chars[pos] != '"' {
                pos += 1;
            }
            if pos == chars.len() {
                return Err(EvalError::Error(format!(
                    "无法解析的条件格式公式：{formula}"
                )));
            }
            tokens.push(Token::String(chars[start..pos].iter().collect()));
            pos += 1;
            continue;
        }
        if chars[pos] == '$' || chars[pos].is_ascii_alphabetic() {
            if let Some((reference, end)) = parse_reference_at(&chars, pos) {
                tokens.push(Token::Reference(reference));
                pos = end;
                continue;
            }
            if chars[pos].is_ascii_alphabetic() || chars[pos] == '_' {
                let start = pos;
                pos += 1;
                while pos < chars.len()
                    && (chars[pos].is_ascii_alphanumeric()
                        || chars[pos] == '_'
                        || chars[pos] == '.')
                {
                    pos += 1;
                }
                tokens.push(Token::Name(chars[start..pos].iter().collect()));
                continue;
            }
        }
        let (op, width) = match (chars[pos], chars.get(pos + 1).copied()) {
            ('<', Some('>')) => ("<>", 2),
            ('<', Some('=')) => ("<=", 2),
            ('>', Some('=')) => (">=", 2),
            ('<', _) => ("<", 1),
            ('>', _) => (">", 1),
            ('=', _) => ("=", 1),
            ('+', _) => ("+", 1),
            ('-', _) => ("-", 1),
            ('*', _) => ("*", 1),
            ('/', _) => ("/", 1),
            ('(', _) => ("(", 1),
            (')', _) => (")", 1),
            (',', _) => (",", 1),
            _ => {
                return Err(EvalError::Error(format!(
                    "无法解析的条件格式公式：{formula}"
                )));
            }
        };
        tokens.push(Token::Operator(op.to_owned()));
        pos += width;
    }
    tokens.push(Token::End);
    Ok(tokens)
}

fn parse_reference_at(chars: &[char], start: usize) -> Option<(RangeRef, usize)> {
    let (first, mut pos) = parse_reference_part(chars, start)?;
    let mut last = first;
    if chars.get(pos) == Some(&':') {
        let (second, end) = parse_reference_part(chars, pos + 1)?;
        last = second;
        pos = end;
    }
    let previous = start
        .checked_sub(1)
        .and_then(|index| chars.get(index))
        .copied();
    let next = chars.get(pos).copied();
    if previous.is_some_and(|ch| ch.is_ascii_alphanumeric() || ch == '_' || ch == '$')
        || next.is_some_and(|ch| ch.is_ascii_alphanumeric() || ch == '_' || ch == '.')
        || next == Some('(')
    {
        return None;
    }
    Some((RangeRef { first, last }, pos))
}

fn parse_reference_part(chars: &[char], start: usize) -> Option<(CellRef, usize)> {
    let mut pos = start;
    let column_absolute = chars.get(pos) == Some(&'$');
    if column_absolute {
        pos += 1;
    }
    let column_start = pos;
    while chars.get(pos).is_some_and(|ch| ch.is_ascii_alphabetic()) && pos - column_start < 3 {
        pos += 1;
    }
    if pos == column_start || chars.get(pos).is_some_and(|ch| ch.is_ascii_alphabetic()) {
        return None;
    }
    let column_end = pos;
    let row_absolute = chars.get(pos) == Some(&'$');
    if row_absolute {
        pos += 1;
    }
    let row_start = pos;
    while chars.get(pos).is_some_and(|ch| ch.is_ascii_digit()) {
        pos += 1;
    }
    if pos == row_start {
        return None;
    }
    let column = chars[column_start..column_end]
        .iter()
        .fold(0i32, |sum, ch| {
            sum * 26 + (ch.to_ascii_uppercase() as i32 - 'A' as i32 + 1)
        });
    let row = chars[row_start..pos]
        .iter()
        .collect::<String>()
        .parse::<i32>()
        .ok()?;
    if row < 1 || row > 1_048_576 || column < 1 || column > 16_384 {
        return None;
    }
    Some((
        CellRef {
            column,
            row,
            column_absolute,
            row_absolute,
        },
        pos,
    ))
}

struct Parser<'a> {
    tokens: Vec<Token>,
    position: usize,
    formula: &'a str,
    grid: &'a ValueGrid<'a>,
    dr: i32,
    dc: i32,
}

impl Parser<'_> {
    fn peek(&self) -> &Token {
        self.tokens.get(self.position).unwrap_or(&Token::End)
    }

    fn eat(&mut self) -> Token {
        let token = self.peek().clone();
        self.position += 1;
        token
    }

    fn eat_operator(&mut self, operator: &str) -> bool {
        if self.peek() == &Token::Operator(operator.to_owned()) {
            self.eat();
            true
        } else {
            false
        }
    }

    fn expression(&mut self) -> std::result::Result<Value, EvalError> {
        let mut left = self.additive()?;
        loop {
            let op = match self.peek() {
                Token::Operator(op)
                    if matches!(op.as_str(), "<>" | "<=" | ">=" | "<" | ">" | "=") =>
                {
                    op.clone()
                }
                _ => break,
            };
            self.eat();
            let right = self.additive()?;
            left = compare_values(&op, left, right);
        }
        Ok(left)
    }

    fn additive(&mut self) -> std::result::Result<Value, EvalError> {
        let mut left = self.multiplicative()?;
        loop {
            let op = match self.peek() {
                Token::Operator(op) if matches!(op.as_str(), "+" | "-") => op.clone(),
                _ => break,
            };
            self.eat();
            let right = self.multiplicative()?;
            left = if op == "+"
                && (matches!(left, Value::Text(_)) || matches!(right, Value::Text(_)))
            {
                Value::Text(format!(
                    "{}{}",
                    py_text_or_empty(&left),
                    py_text_or_empty(&right)
                ))
            } else if matches!(left, Value::Text(_)) || matches!(right, Value::Text(_)) {
                Value::Empty
            } else {
                let left = expression_number(&left).unwrap_or(0.0);
                let right = expression_number(&right).unwrap_or(0.0);
                Value::Number(excel_round15(if op == "+" {
                    left + right
                } else {
                    left - right
                }))
            };
        }
        Ok(left)
    }

    fn multiplicative(&mut self) -> std::result::Result<Value, EvalError> {
        let mut left = self.unary()?;
        loop {
            let op = match self.peek() {
                Token::Operator(op) if matches!(op.as_str(), "*" | "/") => op.clone(),
                _ => break,
            };
            self.eat();
            let right = self.unary()?;
            let (Some(left_number), Some(right_number)) =
                (expression_number(&left), expression_number(&right))
            else {
                left = Value::Empty;
                continue;
            };
            if op == "/" && right_number == 0.0 {
                left = Value::Empty;
            } else {
                left = Value::Number(excel_round15(if op == "*" {
                    left_number * right_number
                } else {
                    left_number / right_number
                }));
            }
        }
        Ok(left)
    }

    fn unary(&mut self) -> std::result::Result<Value, EvalError> {
        if self.eat_operator("-") {
            return Ok(Value::Number(
                -expression_number(&self.unary()?).unwrap_or(0.0),
            ));
        }
        self.primary()
    }

    fn primary(&mut self) -> std::result::Result<Value, EvalError> {
        match self.eat() {
            Token::Number(text) => text
                .parse::<f64>()
                .map(Value::Number)
                .map_err(|_| EvalError::Error(format!("无法解析的条件格式公式：{}", self.formula))),
            Token::String(text) => Ok(Value::Text(text)),
            Token::Reference(reference) => self.grid.range(reference, self.dr, self.dc),
            Token::Name(name) => {
                if name.eq_ignore_ascii_case("true") {
                    return Ok(Value::Boolean(true));
                }
                if name.eq_ignore_ascii_case("false") {
                    return Ok(Value::Boolean(false));
                }
                if !self.eat_operator("(") {
                    return Err(EvalError::Error(format!(
                        "无法解析的条件格式公式：{}",
                        self.formula
                    )));
                }
                let mut arguments = Vec::new();
                if !self.eat_operator(")") {
                    loop {
                        arguments.push(self.expression()?);
                        if self.eat_operator(")") {
                            break;
                        }
                        if !self.eat_operator(",") {
                            return Err(EvalError::Error(format!(
                                "无法解析的条件格式公式：{}",
                                self.formula
                            )));
                        }
                    }
                }
                call_function(&name, &arguments, self.grid)
            }
            Token::Operator(op) if op == "(" => {
                let value = self.expression()?;
                if !self.eat_operator(")") {
                    return Err(EvalError::Error(format!(
                        "无法解析的条件格式公式：{}",
                        self.formula
                    )));
                }
                Ok(value)
            }
            _ => Err(EvalError::Error(format!(
                "无法解析的条件格式公式：{}",
                self.formula
            ))),
        }
    }
}

fn call_function(
    name: &str,
    arguments: &[Value],
    grid: &ValueGrid<'_>,
) -> std::result::Result<Value, EvalError> {
    match name.to_ascii_uppercase().as_str() {
        "AND" => Ok(Value::Boolean(arguments.iter().all(truthy))),
        "OR" => Ok(Value::Boolean(arguments.iter().any(truthy))),
        "NOT" => Ok(Value::Boolean(
            arguments
                .first()
                .map(|value| !truthy(value))
                .unwrap_or(false),
        )),
        "ABS" => {
            let argument = arguments
                .first()
                .ok_or_else(|| EvalError::Error("ABS 函数缺少参数".to_owned()))?;
            Ok(expression_number(argument)
                .map(|value| Value::Number(value.abs()))
                .unwrap_or(Value::Empty))
        }
        "SUM" => {
            let mut total = 0.0;
            for argument in arguments {
                match argument {
                    Value::Range(range) => {
                        for item in grid.range_values(*range)? {
                            if let Some(number) = expression_number(&item) {
                                total += number;
                            }
                        }
                    }
                    value => {
                        if let Some(number) = expression_number(value) {
                            total += number;
                        }
                    }
                }
            }
            Ok(Value::Number(excel_round15(total)))
        }
        _ => Err(EvalError::Unsupported(format!(
            "暂不支持的条件格式函数：{}",
            name.to_ascii_uppercase()
        ))),
    }
}

fn compare_values(operator: &str, mut left: Value, mut right: Value) -> Value {
    let left_empty = is_empty(&left);
    let right_empty = is_empty(&right);
    if !left_empty
        && !right_empty
        && !matches!(left, Value::Boolean(_))
        && !matches!(right, Value::Boolean(_))
    {
        if let (Some(left_number), Some(right_number)) =
            (expression_number(&left), expression_number(&right))
        {
            left = Value::Number(excel_round15(left_number));
            right = Value::Number(excel_round15(right_number));
        }
    }
    let result = match operator {
        "=" => {
            if left_empty && right_empty {
                true
            } else if left_empty || right_empty {
                false
            } else {
                left == right
            }
        }
        "<>" => {
            if left_empty && right_empty {
                false
            } else if left_empty || right_empty {
                true
            } else {
                left != right
            }
        }
        _ => {
            let left_number = expression_number(if left_empty {
                &Value::Number(0.0)
            } else {
                &left
            });
            let right_number = expression_number(if right_empty {
                &Value::Number(0.0)
            } else {
                &right
            });
            match (left_number, right_number) {
                (Some(left), Some(right)) => match operator {
                    "<" => left < right,
                    ">" => left > right,
                    "<=" => left <= right,
                    ">=" => left >= right,
                    _ => false,
                },
                _ => false,
            }
        }
    };
    Value::Boolean(result)
}

fn is_empty(value: &Value) -> bool {
    matches!(value, Value::Empty) || matches!(value, Value::Text(text) if text.is_empty())
}

fn truthy(value: &Value) -> bool {
    match value {
        Value::Empty => false,
        Value::Number(number) => *number != 0.0,
        Value::Text(text) => !text.is_empty(),
        Value::Boolean(value) => *value,
        Value::Range(_) => true,
    }
}

fn expression_number(value: &Value) -> Option<f64> {
    match value {
        Value::Number(number) => Some(*number),
        Value::Text(text) => parse_number(text.trim()),
        Value::Empty | Value::Boolean(_) | Value::Range(_) => None,
    }
}

fn parse_number(text: &str) -> Option<f64> {
    if let Some(percent) = text.strip_suffix('%') {
        return percent
            .trim()
            .parse::<f64>()
            .ok()
            .map(|value| value / 100.0);
    }
    text.parse::<f64>().ok()
}

fn excel_round15(value: f64) -> f64 {
    if value == 0.0 || !value.is_finite() {
        return value;
    }
    format!("{value:.14e}").parse::<f64>().unwrap_or(value)
}

fn py_text_or_empty(value: &Value) -> String {
    match value {
        Value::Empty => String::new(),
        Value::Number(number) if *number == 0.0 => String::new(),
        Value::Boolean(false) => String::new(),
        Value::Text(text) => text.clone(),
        Value::Number(number) => number.to_string(),
        Value::Boolean(true) => "True".to_owned(),
        Value::Range(_) => "[]".to_owned(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn grid(entries: &[(&str, Value)]) -> BTreeMap<String, Value> {
        entries
            .iter()
            .map(|(address, value)| ((*address).to_owned(), value.clone()))
            .collect()
    }

    fn evaluate(
        formula: &str,
        values: &[(&str, Value)],
        anchor: (i32, i32),
        target: (i32, i32),
    ) -> RuleEvaluation {
        let values = grid(values);
        let grid = ValueGrid { values: &values };
        match evaluate_expression(formula, &grid, anchor.0, anchor.1, target.0, target.1) {
            Ok(value) => RuleEvaluation::new(if truthy(&value) {
                EvaluationStatus::True
            } else {
                EvaluationStatus::False
            }),
            Err(error) => RuleEvaluation::with_reason(error.status(), error.reason()),
        }
    }

    #[test]
    fn cellis_numeric_between_empty_and_unsupported_match_source_contract() {
        let resolve = |index: usize| -> std::result::Result<Value, EvalError> {
            Ok([Value::Number(0.0), Value::Number(10.0)][index].clone())
        };
        assert_eq!(
            evaluate_cell_is("greaterThan", &[], &Value::Number(5.0), resolve).status,
            EvaluationStatus::True
        );
        assert_eq!(
            evaluate_cell_is("between", &[], &Value::Number(11.0), resolve).status,
            EvaluationStatus::False
        );
        assert_eq!(
            evaluate_cell_is("greaterThan", &[], &Value::Empty, resolve).status,
            EvaluationStatus::False
        );
        assert_eq!(
            evaluate_cell_is("containsText", &[], &Value::Text("x".into()), resolve).status,
            EvaluationStatus::Unsupported
        );
    }

    #[test]
    fn expression_resolves_relative_and_locked_references() {
        let values = [
            ("C5", Value::Number(2.0)),
            ("D5", Value::Number(0.0)),
            ("E5", Value::Number(0.5)),
            ("C26", Value::Number(3.0)),
        ];
        assert_eq!(
            evaluate("AND(C5>0,C5>1)", &values, (5, 3), (5, 3)).status,
            EvaluationStatus::True
        );
        assert_eq!(
            evaluate("AND(C5>0,C5>1)", &values, (5, 3), (5, 4)).status,
            EvaluationStatus::False
        );
        assert_eq!(
            evaluate("AND(C5>0,C5>1)", &values, (5, 3), (26, 3)).status,
            EvaluationStatus::True
        );
        assert_eq!(
            evaluate("AND($C$5>0,C5>1)", &values, (5, 3), (5, 4)).status,
            EvaluationStatus::False
        );
    }

    #[test]
    fn expression_supports_sum_ranges_and_excel_15_digit_rounding() {
        let values = [
            ("D27", Value::Number(31_894.55)),
            ("D28", Value::Number(31.28)),
            ("D34", Value::Empty),
            ("D35", Value::Number(31_925.83)),
        ];
        assert_eq!(
            evaluate("$D35<>$D27+$D28+$D34", &values, (35, 4), (35, 4)).status,
            EvaluationStatus::False
        );
        let values = [
            ("A1", Value::Number(1.0)),
            ("A2", Value::Number(2.0)),
            ("A3", Value::Number(3.0)),
            ("C2", Value::Number(6.0)),
        ];
        assert_eq!(
            evaluate("SUM(A1:A3)=C2", &values, (2, 3), (2, 3)).status,
            EvaluationStatus::True
        );
    }

    #[test]
    fn expression_preserves_unsupported_and_error_states() {
        let no_values = [];
        assert_eq!(
            evaluate("INDIRECT(\"A1\")>0", &no_values, (1, 1), (1, 1)).status,
            EvaluationStatus::Unsupported
        );
        assert_eq!(
            evaluate("(((", &no_values, (1, 1), (1, 1)).status,
            EvaluationStatus::Error
        );
    }

    #[test]
    fn expression_matches_empty_boolean_text_and_percent_rules() {
        let values = [("A1", Value::Empty), ("B1", Value::Text(String::new()))];
        assert_eq!(
            evaluate("A1=B1", &values, (1, 1), (1, 1)).status,
            EvaluationStatus::True
        );
        assert_eq!(
            evaluate("A1<>B1", &values, (1, 1), (1, 1)).status,
            EvaluationStatus::False
        );
        assert_eq!(
            evaluate("TRUE=TRUE", &[], (1, 1), (1, 1)).status,
            EvaluationStatus::True
        );
        assert_eq!(
            evaluate("0.5=50%", &[], (1, 1), (1, 1)).status,
            EvaluationStatus::Error
        );
    }
}
