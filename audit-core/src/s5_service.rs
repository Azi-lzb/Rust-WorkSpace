//! S5 audit service rule evaluation (`engine._evaluate_rule` port).
//!
//! Wires the configuration, the retrieval core ([`crate::s5_retrieval`]), and
//! the expression runtime ([`crate::s5_eval`]) into the per-rule semantics of
//! the source project's `AuditEngine._evaluate_rule`: dictionary variable
//! caching, retrieval caching and traces, the applicable-expression gate, the
//! audit/current/comparison/difference expressions with kind checks,
//! description rendering, evidence aggregation, and the exact error-status
//! classification. Applicability-scope (`@当前行`) expansion belongs to the
//! batch orchestration slice and raises the same errors as the Python
//! single-rule path.

use crate::s5_config::RuleDefinition;
use crate::s5_config_dsl::{parse_dictionary, parse_expression, Node};
use crate::s5_eval::{
    evaluate, EvalArgument, EvalError, EvalErrorKind, EvalScalar, EvalValue, EvaluationHooks,
};
use crate::s5_retrieval::{RetrievalFailure, S5Retrieval};
use serde::Serialize;
use serde_json::{json, Value as JsonValue};
use std::cell::RefCell;
use std::collections::{BTreeMap, BTreeSet};

#[derive(Clone, Debug, Default, PartialEq, Serialize)]
pub struct S5RuleOutcome {
    pub rule_id: String,
    pub revision: i64,
    pub rule_name: String,
    pub business: String,
    pub period: String,
    pub institution_code: String,
    pub institution_name: String,
    pub severity: String,
    pub check_type: String,
    pub metric: String,
    pub sheet_name: String,
    pub formula_cell: String,
    pub status: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub triggered: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error_code: Option<String>,
    #[serde(skip_serializing_if = "String::is_empty", default)]
    pub error_message: String,
    #[serde(default)]
    pub description: String,
    #[serde(default)]
    pub scope_path: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub current_value: Option<JsonValue>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub comparison_value: Option<JsonValue>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub difference_value: Option<JsonValue>,
    #[serde(skip_serializing_if = "String::is_empty", default)]
    pub current_unit: String,
    #[serde(skip_serializing_if = "String::is_empty", default)]
    pub comparison_unit: String,
    #[serde(skip_serializing_if = "String::is_empty", default)]
    pub difference_unit: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub current_precision: Option<i64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub comparison_precision: Option<i64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub difference_precision: Option<i64>,
    #[serde(default)]
    pub variables: BTreeMap<String, JsonValue>,
    #[serde(default)]
    pub evidence: Vec<JsonValue>,
    #[serde(default)]
    pub retrieval_trace: Vec<JsonValue>,
    #[serde(default)]
    pub unused_variables: Vec<String>,
}

const STRUCTURE_ERROR_CODES: [&str; 14] = [
    "SHEET_NOT_FOUND",
    "STRUCTURE_ANCHOR_NOT_FOUND",
    "ROW_PATH_NOT_FOUND",
    "COLUMN_PATH_NOT_FOUND",
    "STRUCTURE_ERROR",
    "METRIC_AMBIGUOUS",
    "MAPPING_SHEET_MISSING",
    "MAPPING_FIELD_NOT_FOUND",
    "MAPPING_FIELD_AMBIGUOUS",
    "METRIC_NOT_FOUND",
    "REPORT_ITEM_NOT_FOUND",
    "REPORT_ITEM_DUPLICATE",
    "NUMERIC_REGION_NOT_FOUND",
    "NUMERIC_REGION_AMBIGUOUS",
];

const CONFIG_ERROR_CODES: [&str; 7] = [
    "SOURCE_CONTRACT_MISSING",
    "CATALOG_METRIC_NOT_FOUND",
    "METRIC_NOT_CONFIRMED",
    "UNIT_CONTRACT_MISSING",
    "PERIOD_CONTRACT_MISSING",
    "INSTITUTION_CONTRACT_MISSING",
    "SOURCE_AMBIGUOUS",
];

/// Mirror of `_display_value`.
pub fn display_value(value: &EvalValue) -> String {
    match &value.value {
        EvalScalar::Blank => "（空值）".to_owned(),
        EvalScalar::Number(number) => match number {
            crate::s5_eval::NumberValue::Binary64(raw) => python_g15(*raw),
            crate::s5_eval::NumberValue::Decimal(decimal) => decimal.to_canonical_string(),
        },
        EvalScalar::Boolean(flag) => {
            if *flag {
                "是".to_owned()
            } else {
                "否".to_owned()
            }
        }
        EvalScalar::Text(text) => text.clone(),
        EvalScalar::Tuple(values) => values.join(","),
    }
}

/// Python `format(value, ".15g")` with the lowercase `e` upper-cased.
fn python_g15(raw: f64) -> String {
    let text = format!("{raw:.14e}");
    let (mantissa, exponent_text) = text.split_once('e').unwrap_or((text.as_str(), "0"));
    let exponent: i32 = exponent_text.parse().unwrap_or(0);
    if (-5..15).contains(&exponent) {
        let plain = format!("{raw:.15}");
        let trimmed = plain.trim_end_matches('0').trim_end_matches('.');
        if trimmed.is_empty() {
            "0".to_owned()
        } else {
            trimmed.to_owned()
        }
    } else {
        let sign = if exponent < 0 { "-" } else { "+" };
        format!("{mantissa}E{}{abs}", sign, abs = exponent.abs())
    }
}

/// Mirror of `_render_description`.
fn render_description(
    template: &str,
    resolve: &mut dyn FnMut(&str) -> Result<EvalValue, EvalError>,
) -> Result<String, EvalError> {
    let characters: Vec<char> = template.chars().collect();
    let mut result = String::new();
    let mut index = 0usize;
    while index < characters.len() {
        let marker_at = |character: char, at: usize| characters.get(at) == Some(&character);
        if marker_at('{', index) && marker_at('{', index + 1) {
            result.push('{');
            index += 2;
            continue;
        }
        if marker_at('}', index) && marker_at('}', index + 1) {
            result.push('}');
            index += 2;
            continue;
        }
        if characters[index] == '{' {
            let end = characters[index + 1..]
                .iter()
                .position(|c| *c == '}')
                .map(|offset| offset + index + 1);
            let Some(end) = end else {
                return Err(EvalError::configuration(
                    "规则描述中的变量占位符缺少结束大括号",
                ));
            };
            let name: String = characters[index + 1..end].iter().collect();
            let valid = !name.is_empty()
                && name.chars().all(|c| {
                    c == '_' || c.is_alphanumeric() || ('\u{3400}'..='\u{9fff}').contains(&c)
                });
            if !valid {
                return Err(EvalError::configuration(format!(
                    "规则描述中的变量名无效：{name}"
                )));
            }
            let value = resolve(&name)?;
            result.push_str(&display_value(&value));
            index = end + 1;
            continue;
        }
        if characters[index] == '}' {
            return Err(EvalError::configuration("规则描述中的右大括号需要写成 }}"));
        }
        result.push(characters[index]);
        index += 1;
    }
    Ok(result)
}

/// Mirror of `_error_code`.
fn error_code(message: &str, fallback: &str) -> String {
    if let Some((prefix, _)) = message.split_once('：') {
        let upper = !prefix.is_empty()
            && prefix
                .chars()
                .next()
                .is_some_and(|c| c.is_ascii_uppercase())
            && prefix
                .chars()
                .all(|c| c.is_ascii_uppercase() || c.is_ascii_digit() || c == '_');
        if upper {
            return prefix.to_owned();
        }
    }
    fallback.to_owned()
}

struct RuleHooks<'a> {
    retrieval: &'a RefCell<&'a mut S5Retrieval>,
    scope_path: Option<Vec<String>>,
    rule: &'a RuleDefinition,
    dictionary: &'a BTreeMap<String, Node>,
    numeric_mode: String,
    variable_cache: RefCell<BTreeMap<String, EvalValue>>,
    variable_used: RefCell<BTreeSet<String>>,
    variable_stack: RefCell<Vec<String>>,
    retrieval_cache: RefCell<BTreeMap<String, Result<EvalValue, RetrievalFailure>>>,
    trace: RefCell<Vec<JsonValue>>,
}

fn trace_source_id(
    name: &str,
    keywords: &BTreeMap<String, String>,
    default_source: &str,
) -> String {
    if let Some(requested) = keywords.get("source") {
        if !requested.is_empty() {
            return requested.clone();
        }
    }
    default_source.to_owned()
}

fn argument_text(value: &EvalArgument) -> String {
    match value {
        EvalArgument::Text(text) => text.clone(),
        EvalArgument::Tuple(values) => values.join("@"),
    }
}

fn path_argument(value: &EvalArgument) -> Vec<String> {
    match value {
        EvalArgument::Text(text) => crate::s5_config::split_at_path(text),
        EvalArgument::Tuple(values) => values.clone(),
    }
}

fn eval_scalar_json(value: &EvalScalar) -> JsonValue {
    match value {
        EvalScalar::Blank => JsonValue::Null,
        EvalScalar::Boolean(flag) => json!(flag),
        EvalScalar::Text(text) => json!(text),
        EvalScalar::Number(number) => match number {
            crate::s5_eval::NumberValue::Binary64(raw) => serde_json::Number::from_f64(*raw)
                .map(JsonValue::Number)
                .unwrap_or(JsonValue::Null),
            crate::s5_eval::NumberValue::Decimal(decimal) => json!(decimal.to_canonical_string()),
        },
        EvalScalar::Tuple(values) => json!(values),
    }
}

fn resolved_value_json(value: &EvalValue) -> JsonValue {
    eval_scalar_json(&value.value)
}

impl<'a> RuleHooks<'a> {
    fn push_trace(
        &self,
        name: &str,
        args: &[EvalArgument],
        keywords: &BTreeMap<String, String>,
        outcome: Result<&EvalValue, &str>,
    ) {
        let mut source_sheet = keywords.get("sheet").cloned().unwrap_or_default();
        let mut row_path = String::new();
        let mut column_path = String::new();
        let mut metric_header = String::new();
        if matches!(name, "V_CUR" | "V_PRE") {
            match args {
                [EvalArgument::Text(sheet), row, column] => {
                    source_sheet = sheet.clone();
                    row_path = path_argument(row).join("@");
                    column_path = path_argument(column).join("@");
                }
                [row, column] => {
                    row_path = path_argument(row).join("@");
                    column_path = path_argument(column).join("@");
                }
                _ => {}
            }
        } else if name == "V_EXT" {
            if let Some(EvalArgument::Text(header)) = args.first() {
                metric_header = header.clone();
                column_path = header.clone();
            }
        }
        let source_id = trace_source_id(name, keywords, &self.rule.default_source);
        let evidence_items: Vec<JsonValue> = match outcome {
            Ok(value) if !value.evidence.is_empty() => value.evidence.clone(),
            _ => vec![json!(null)],
        };
        let mut trace = self.trace.borrow_mut();
        for item in &evidence_items {
            let field = |key: &str| -> String {
                item.get(key)
                    .map(|value| {
                        value
                            .as_str()
                            .map(str::to_owned)
                            .unwrap_or_else(|| value.to_string())
                    })
                    .unwrap_or_default()
            };
            trace.push(json!({
                "variable": self.variable_stack.borrow().last().cloned().unwrap_or_else(|| "审核表达式".to_owned()),
                "function": name,
                "source_id": if item.is_null() { source_id.clone() } else { field("source_id") },
                "sheet": if item.is_null() { source_sheet.clone() } else { field("sheet") },
                "requested_sheet": source_sheet,
                "row_path": row_path,
                "column_path": column_path,
                "metric": metric_header,
                "cell": if item.is_null() { String::new() } else { field("cell") },
                "raw_value": if item.is_null() { String::new() } else { field("raw_value") },
                "resolved_value": match outcome {
                    Ok(value) => resolved_value_json(value),
                    Err(_) => json!(""),
                },
                "source_unit": if item.is_null() { String::new() } else { field("source_unit") },
                "standard_unit": if item.is_null() { String::new() } else { field("standard_unit") },
                "conversion": if item.is_null() { String::new() } else { field("conversion") },
                "report_item": if item.is_null() { String::new() } else { field("report_item") },
                "mapping_cell": if item.is_null() { String::new() } else { field("mapping_cell") },
                "institution_code": if item.is_null() { String::new() } else { field("institution_code") },
                "status": if outcome.is_err() { "取数失败" } else { "已取到" },
                "error": match outcome {
                    Err(message) => message.to_owned(),
                    Ok(_) => String::new(),
                },
            }));
        }
    }
}

impl<'a> crate::s5_eval::EvaluationHooks for RuleHooks<'a> {
    fn resolve_variable(&self, name: &str) -> Result<EvalValue, EvalError> {
        if let Some(value) = self.variable_cache.borrow().get(name) {
            self.variable_used.borrow_mut().insert(name.to_owned());
            return Ok(value.clone());
        }
        let node = self.dictionary.get(name).ok_or_else(|| {
            EvalError::evaluation(format!("UNDEFINED_VARIABLE：未定义变量 {name}"))
        })?;
        self.variable_used.borrow_mut().insert(name.to_owned());
        self.variable_stack.borrow_mut().push(name.to_owned());
        let result = evaluate(node, self, &self.numeric_mode);
        self.variable_stack.borrow_mut().pop();
        let value = result?;
        self.variable_cache
            .borrow_mut()
            .insert(name.to_owned(), value.clone());
        Ok(value)
    }

    fn call_value(
        &self,
        name: &str,
        arguments: &[EvalArgument],
        keywords: &BTreeMap<String, String>,
    ) -> Result<EvalValue, EvalError> {
        let mut arguments: Vec<EvalArgument> = arguments.to_vec();
        let current_row_dynamic = match (name, arguments.len()) {
            ("V_CUR" | "V_PRE", 2) => arguments
                .first()
                .map(|argument| argument_text(argument) == "@当前行")
                .unwrap_or(false),
            ("V_CUR" | "V_PRE", 3) => arguments
                .get(1)
                .map(|argument| argument_text(argument) == "@当前行")
                .unwrap_or(false),
            _ => false,
        };
        if current_row_dynamic {
            let Some(scope_path) = &self.scope_path else {
                return Err(EvalError::configuration(
                    "取值函数使用 @当前行，但当前规则没有展开适用范围",
                ));
            };
            let scope = EvalArgument::Tuple(scope_path.clone());
            match arguments.len() {
                2 => {
                    arguments[0] = scope;
                }
                _ => {
                    arguments[1] = scope;
                }
            }
        }
        let cache_key = format!(
            "{}|{}|{}",
            name,
            arguments
                .iter()
                .map(argument_text)
                .collect::<Vec<_>>()
                .join("|"),
            {
                let mut sorted: Vec<String> = keywords
                    .iter()
                    .map(|(key, value)| format!("{key}={value}"))
                    .collect();
                sorted.sort();
                sorted.join("|")
            }
        );
        if let Some(existing) = self.retrieval_cache.borrow().get(&cache_key) {
            return match existing {
                Ok(value) => Ok(value.clone()),
                Err(failure) => Err(failure.clone().to_eval_error()),
            };
        }
        let outcome = self.retrieval.borrow_mut().retrieve(
            name,
            &arguments,
            keywords,
            &self.rule.business,
            &self.rule.default_sheet,
            &self.rule.default_source,
        );
        match outcome {
            Ok(value) => {
                self.push_trace(name, &arguments, keywords, Ok(&value));
                self.retrieval_cache
                    .borrow_mut()
                    .insert(cache_key, Ok(value.clone()));
                Ok(value)
            }
            Err(failure) => {
                self.push_trace(name, &arguments, keywords, Err(&failure.message));
                self.retrieval_cache
                    .borrow_mut()
                    .insert(cache_key, Err(failure.clone()));
                Err(failure.to_eval_error())
            }
        }
    }
}

fn unique_evidence(values: &[EvalValue]) -> Vec<JsonValue> {
    let mut seen: BTreeSet<(String, String, String, String)> = BTreeSet::new();
    let mut unique = Vec::new();
    for value in values {
        for item in &value.evidence {
            let key = (
                item.get("source_id")
                    .and_then(JsonValue::as_str)
                    .unwrap_or_default()
                    .to_owned(),
                item.get("sheet")
                    .and_then(JsonValue::as_str)
                    .unwrap_or_default()
                    .to_owned(),
                item.get("cell")
                    .and_then(JsonValue::as_str)
                    .unwrap_or_default()
                    .to_owned(),
                item.get("source_sha256")
                    .and_then(JsonValue::as_str)
                    .unwrap_or_default()
                    .to_owned(),
            );
            if !seen.contains(&key) {
                seen.insert(key);
                unique.push(item.clone());
            }
        }
    }
    unique
}

fn cache_evidence(cache: &BTreeMap<String, EvalValue>) -> Vec<JsonValue> {
    let mut evidence = Vec::new();
    for value in cache.values() {
        evidence.extend(value.evidence.iter().cloned());
    }
    evidence
}

/// Mirror of `AuditEngine._evaluate_rule` for one rule without
/// applicability-scope expansion.
#[allow(clippy::too_many_arguments)]
fn base_outcome(
    rule: &RuleDefinition,
    period: &str,
    institution_code: &str,
    institution_name: &str,
    scope_path: Option<&[String]>,
) -> S5RuleOutcome {
    S5RuleOutcome {
        rule_id: rule.rule_id.clone(),
        revision: rule.revision,
        rule_name: rule.name.clone(),
        business: rule.business.clone(),
        period: period.to_owned(),
        institution_code: institution_code.to_owned(),
        institution_name: institution_name.to_owned(),
        severity: rule.severity.clone(),
        check_type: rule.check_type.clone(),
        metric: rule.metric.clone(),
        sheet_name: rule.default_sheet.clone(),
        formula_cell: rule.formula_cell.clone(),
        status: "NOT_EXECUTED".to_owned(),
        triggered: None,
        description: rule.description.clone(),
        scope_path: scope_path.map(|path| path.join(" / ")).unwrap_or_default(),
        ..Default::default()
    }
}

#[allow(clippy::too_many_arguments)]
pub fn evaluate_rule(
    rule: &RuleDefinition,
    retrieval: &mut S5Retrieval,
    run_id: &str,
    context_period: &str,
    institution_code: &str,
    institution_name: &str,
    configuration_version: &str,
    scope_path: Option<&[String]>,
) -> S5RuleOutcome {
    let _ = (run_id, configuration_version);
    let mut outcome = base_outcome(
        rule,
        context_period,
        institution_code,
        institution_name,
        scope_path,
    );

    let dictionary_nodes = match parse_dictionary(if rule.value_dictionary.is_empty() {
        None
    } else {
        Some(rule.value_dictionary.as_str())
    }) {
        Ok(nodes) => nodes,
        Err(error) => {
            outcome.status = "CONFIG_ERROR".to_owned();
            outcome.error_code = Some("CONFIG_ERROR".to_owned());
            outcome.error_message = error.message;
            return outcome;
        }
    };
    if let Err(error) = crate::s5_config_dsl::validate_dictionary_dependencies(&dictionary_nodes) {
        outcome.status = "CONFIG_ERROR".to_owned();
        outcome.error_code = Some("CONFIG_ERROR".to_owned());
        outcome.error_message = error.message;
        return outcome;
    }
    let dictionary: BTreeMap<String, Node> = dictionary_nodes
        .iter()
        .map(|(name, node)| (name.clone(), node.clone()))
        .collect();

    let retrieval_cell = RefCell::new(retrieval);
    let hooks = RuleHooks {
        retrieval: &retrieval_cell,
        scope_path: scope_path.map(|path| path.to_vec()),
        rule,
        dictionary: &dictionary,
        numeric_mode: rule.numeric_mode.clone(),
        variable_cache: RefCell::new(BTreeMap::new()),
        variable_used: RefCell::new(BTreeSet::new()),
        variable_stack: RefCell::new(Vec::new()),
        retrieval_cache: RefCell::new(BTreeMap::new()),
        trace: RefCell::new(Vec::new()),
    };

    let run_expression = |hooks: &RuleHooks<'_>, text: &str| -> Result<EvalValue, EvalError> {
        if text.trim().is_empty() {
            return Err(EvalError::configuration("需要执行的表达式为空"));
        }
        let node =
            parse_expression(text).map_err(|error| EvalError::configuration(error.message))?;
        evaluate(&node, hooks, &rule.numeric_mode)
    };

    let evaluated = (|| -> Result<S5RuleOutcome, EvalError> {
        if !rule.applicable_expression.is_empty() {
            let applicable = run_expression(&hooks, &rule.applicable_expression)?;
            match applicable.value {
                EvalScalar::Boolean(flag) => {
                    if !flag {
                        let mut not_applicable = base_outcome(
                            rule,
                            context_period,
                            institution_code,
                            institution_name,
                            scope_path,
                        );
                        not_applicable.status = "NOT_APPLICABLE".to_owned();
                        not_applicable.error_message = "规则适用表达式判定为不适用".to_owned();
                        not_applicable.evidence = unique_evidence(&[applicable]);
                        not_applicable.unused_variables =
                            unused_dictionary_variables(&dictionary, &hooks.variable_used.borrow());
                        return Ok(not_applicable);
                    }
                }
                _ => {
                    return Err(EvalError::configuration("适用表达式必须返回布尔值"));
                }
            }
        }
        let audit_value = run_expression(&hooks, &rule.audit_expression)?;
        let audit_is_boolean = matches!(audit_value.value, EvalScalar::Boolean(_));
        if !audit_is_boolean {
            return Err(EvalError::evaluation(
                "TYPE_ERROR：审核表达式必须返回布尔值",
            ));
        }
        let audit_passes = match audit_value.value {
            EvalScalar::Boolean(passes) => {
                if rule.audit_expression_inverted {
                    !passes
                } else {
                    passes
                }
            }
            _ => unreachable!("kind checked above"),
        };
        let current_value = if !rule.current_expression.is_empty() {
            Some(run_expression(&hooks, &rule.current_expression)?)
        } else {
            None
        };
        let comparison_value = if !rule.comparison_expression.is_empty() {
            Some(run_expression(&hooks, &rule.comparison_expression)?)
        } else {
            None
        };
        let difference_value = if !rule.difference_expression.is_empty() {
            let difference = run_expression(&hooks, &rule.difference_expression)?;
            match difference.kind {
                "number" | "blank" => Some(difference),
                _ => {
                    return Err(EvalError::evaluation(
                        "TYPE_ERROR：差值表达式必须返回数值或空值",
                    ))
                }
            }
        } else {
            None
        };
        let description =
            render_description(&rule.description, &mut |name| hooks.resolve_variable(name))?;

        let mut outcome = base_outcome(
            rule,
            context_period,
            institution_code,
            institution_name,
            scope_path,
        );
        outcome.status = if audit_passes { "PASSED" } else { "TRIGGERED" }.to_owned();
        outcome.triggered = Some(!audit_passes);
        outcome.description = description;
        if let Some(value) = &current_value {
            outcome.current_value = Some(resolved_value_json(value));
            outcome.current_unit = value.unit.clone().unwrap_or_default();
            outcome.current_precision = value.precision;
        }
        if let Some(value) = &comparison_value {
            outcome.comparison_value = Some(resolved_value_json(value));
            outcome.comparison_unit = value.unit.clone().unwrap_or_default();
            outcome.comparison_precision = value.precision;
        }
        if let Some(value) = &difference_value {
            outcome.difference_value = Some(resolved_value_json(value));
            outcome.difference_unit = value.unit.clone().unwrap_or_default();
            outcome.difference_precision = value.precision;
        }
        for (name, value) in hooks.variable_cache.borrow().iter() {
            outcome
                .variables
                .insert(name.clone(), resolved_value_json(value));
        }
        let mut evidence_inputs: Vec<EvalValue> = vec![audit_value.clone()];
        if let Some(value) = &current_value {
            evidence_inputs.push(value.clone());
        }
        if let Some(value) = &comparison_value {
            evidence_inputs.push(value.clone());
        }
        if let Some(value) = &difference_value {
            evidence_inputs.push(value.clone());
        }
        for value in hooks.variable_cache.borrow().values() {
            evidence_inputs.push(value.clone());
        }
        outcome.evidence = unique_evidence(&evidence_inputs);
        outcome.retrieval_trace = hooks.trace.borrow().clone();
        outcome.unused_variables =
            unused_dictionary_variables(&dictionary, &hooks.variable_used.borrow());
        Ok(outcome)
    })();

    match evaluated {
        Ok(outcome) => outcome,
        Err(error) => {
            let (status, code) = match error.kind {
                EvalErrorKind::Configuration => {
                    ("CONFIG_ERROR".to_owned(), "CONFIG_ERROR".to_owned())
                }
                EvalErrorKind::Evaluation => {
                    if let Some(code) = &error.code {
                        // Retrieval failures carry their own code verbatim.
                        let status = if STRUCTURE_ERROR_CODES.contains(&code.as_str()) {
                            "STRUCTURE_ERROR"
                        } else if CONFIG_ERROR_CODES.contains(&code.as_str()) {
                            "CONFIG_ERROR"
                        } else {
                            "RETRIEVAL_FAILED"
                        };
                        (status.to_owned(), code.clone())
                    } else {
                        (
                            "CALCULATION_FAILED".to_owned(),
                            error_code(&error.message, "CALCULATION_FAILED"),
                        )
                    }
                }
            };
            outcome.status = status;
            outcome.error_code = Some(code);
            outcome.error_message = error.message;
            let cache_snapshot: Vec<EvalValue> =
                hooks.variable_cache.borrow().values().cloned().collect();
            outcome.evidence = unique_evidence(&cache_snapshot);
            for (name, value) in hooks.variable_cache.borrow().iter() {
                outcome
                    .variables
                    .insert(name.clone(), resolved_value_json(value));
            }
            outcome.retrieval_trace = hooks.trace.borrow().clone();
            outcome
        }
    }
}

fn unused_dictionary_variables(
    dictionary: &BTreeMap<String, Node>,
    used: &BTreeSet<String>,
) -> Vec<String> {
    let mut unused: Vec<String> = dictionary
        .keys()
        .filter(|name| !used.contains(*name))
        .cloned()
        .collect();
    unused.sort();
    unused
}

fn parse_period_date(value: &str, field: &str) -> Result<(i32, u32, u32), EvalError> {
    let text = value.trim();
    let digits = text.chars().all(|c| c.is_ascii_digit());
    let (year, month, day) = if digits && text.len() == 8 {
        (
            text[0..4].parse().unwrap_or(0),
            text[4..6].parse().unwrap_or(0),
            text[6..8].parse().unwrap_or(0),
        )
    } else if text.len() == 10
        && text.as_bytes().get(4) == Some(&b'-')
        && text.as_bytes().get(7) == Some(&b'-')
    {
        (
            text[0..4].parse().unwrap_or(0),
            text[5..7].parse().unwrap_or(0),
            text[8..10].parse().unwrap_or(0),
        )
    } else if digits && text.len() == 6 {
        (
            text[0..4].parse().unwrap_or(0),
            text[4..6].parse().unwrap_or(0),
            1,
        )
    } else if text.len() == 7 && text.as_bytes().get(4) == Some(&b'-') {
        (
            text[0..4].parse().unwrap_or(0),
            text[5..7].parse().unwrap_or(0),
            1,
        )
    } else {
        return Err(EvalError::configuration(format!(
            "{field} 日期格式必须为 YYYYMMDD、YYYY-MM-DD、YYYYMM 或 YYYY-MM"
        )));
    };
    let leap = year % 4 == 0 && (year % 100 != 0 || year % 400 == 0);
    let days = match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if leap => 29,
        2 => 28,
        _ => {
            return Err(EvalError::configuration(format!(
                "{field} 日期格式或日期值无效：{text}"
            )))
        }
    };
    if !(1..=days).contains(&day) {
        return Err(EvalError::configuration(format!(
            "{field} 日期格式或日期值无效：{text}"
        )));
    }
    Ok((year, month, day))
}

fn within_span(period: &str, start: &str, end: &str) -> Result<bool, EvalError> {
    let value = parse_period_date(period, "报告期")?;
    if !start.is_empty() && value < parse_period_date(start, "配置版本生效起始日")? {
        return Ok(false);
    }
    if !end.is_empty() && value > parse_period_date(end, "配置版本生效截止日")? {
        return Ok(false);
    }
    Ok(true)
}

fn within_rule_period(rule: &RuleDefinition, period: (i32, u32, u32)) -> bool {
    let parse = |text: &str, field: &str| parse_period_date(text, field).ok();
    if !rule.effective_from.is_empty() {
        match parse(&rule.effective_from, "规则生效起始日") {
            Some(from) if period < from => return false,
            None => return false,
            _ => {}
        }
    }
    if !rule.effective_to.is_empty() {
        match parse(&rule.effective_to, "规则生效截止日") {
            Some(to) if period > to => return false,
            None => return false,
            _ => {}
        }
    }
    true
}

fn scope_paths(
    rule: &RuleDefinition,
    configuration: &crate::s5_config::Configuration,
) -> Vec<Vec<String>> {
    if rule.applicability_scope.is_empty() {
        return vec![Vec::new()];
    }
    let Some(scope) = configuration
        .applicability_scopes
        .iter()
        .find(|scope| scope.scope_id == rule.applicability_scope)
    else {
        return Vec::new();
    };
    let mut paths: Vec<Vec<String>> = configuration
        .metrics
        .iter()
        .filter(|metric| {
            metric.business == scope.business
                && metric.sheet == scope.sheet
                && metric.row_path.len() >= scope.row_path_prefix.len()
                && metric.row_path[..scope.row_path_prefix.len()] == scope.row_path_prefix[..]
        })
        .map(|metric| metric.row_path.clone())
        .collect();
    paths.sort();
    paths.dedup();
    paths
}

fn status_outcome(
    rule: &RuleDefinition,
    period: &str,
    institution_code: &str,
    institution_name: &str,
    status: &str,
    message: &str,
) -> S5RuleOutcome {
    let mut outcome = base_outcome(rule, period, institution_code, institution_name, None);
    outcome.status = status.to_owned();
    outcome.error_message = message.to_owned();
    outcome
}

/// Mirror of `AuditEngine.run` for the single-binding current-source path:
/// rule selection with revision/effective-date gating, applicability-scope
/// expansion (including the `@当前行` substitution), and the status ledger.
/// Hash-integrity downgrade and output writing belong to later slices.
pub fn run_rules(
    configuration: &crate::s5_config::Configuration,
    retrieval: &mut S5Retrieval,
    period: &str,
    businesses: &[String],
    institution_code: &str,
    institution_name: &str,
) -> Result<(Vec<S5RuleOutcome>, BTreeMap<String, usize>), EvalError> {
    if period.is_empty() {
        return Err(EvalError::configuration("运行上下文必须提供报告期"));
    }
    let period_date = parse_period_date(period, "报告期")?;
    let config_in_span = within_span(
        period,
        &configuration.effective_from,
        &configuration.effective_to,
    )?;

    // Group rules by rule_id, choose the max eligible revision.
    let mut groups: BTreeMap<String, Vec<&RuleDefinition>> = BTreeMap::new();
    for rule in &configuration.rules {
        if !businesses.is_empty() && !businesses.iter().any(|business| business == &rule.business) {
            continue;
        }
        groups.entry(rule.rule_id.clone()).or_default().push(rule);
    }
    let mut outcomes: Vec<S5RuleOutcome> = Vec::new();
    for (_rule_id, rules) in &groups {
        let eligible: Vec<&&RuleDefinition> = rules
            .iter()
            .filter(|rule| within_rule_period(rule, period_date))
            .collect();
        let eligible_is_empty = eligible.is_empty();
        let selected: &RuleDefinition = if eligible_is_empty {
            rules
                .iter()
                .max_by_key(|rule| rule.revision)
                .expect("groups are nonempty")
        } else {
            eligible
                .iter()
                .copied()
                .max_by_key(|rule| rule.revision)
                .expect("eligible nonempty")
        };
        let status_override = if eligible_is_empty {
            Some(("NOT_APPLICABLE", "本次报告期不在规则生效范围内"))
        } else if !config_in_span {
            Some(("NOT_APPLICABLE", "本次报告期不在配置版本生效范围内"))
        } else {
            None
        };
        if let Some((status, message)) = status_override {
            outcomes.push(status_outcome(
                selected,
                period,
                institution_code,
                institution_name,
                status,
                message,
            ));
            continue;
        }
        if !config_in_span {
            continue;
        }
        let paths = scope_paths(selected, configuration);
        if paths.is_empty() {
            outcomes.push(status_outcome(
                selected,
                period,
                institution_code,
                institution_name,
                "CONFIG_ERROR",
                "适用范围没有匹配到任何模板行路径",
            ));
            continue;
        }
        for path in paths {
            let outcome = evaluate_rule(
                selected,
                retrieval,
                "RUN",
                period,
                institution_code,
                institution_name,
                &configuration.version,
                Some(&path),
            );
            outcomes.push(outcome);
        }
    }

    let mut counts: BTreeMap<String, usize> = BTreeMap::new();
    for outcome in &outcomes {
        *counts.entry(outcome.status.clone()).or_default() += 1;
    }
    Ok((outcomes, counts))
}
