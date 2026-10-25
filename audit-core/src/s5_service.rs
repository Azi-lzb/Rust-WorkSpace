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
            return Err(EvalError::configuration(
                "取值函数使用 @当前行，但当前规则没有展开适用范围",
            ));
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
            arguments,
            keywords,
            &self.rule.business,
            &self.rule.default_sheet,
            &self.rule.default_source,
        );
        match outcome {
            Ok(value) => {
                self.push_trace(name, arguments, keywords, Ok(&value));
                self.retrieval_cache
                    .borrow_mut()
                    .insert(cache_key, Ok(value.clone()));
                Ok(value)
            }
            Err(failure) => {
                self.push_trace(name, arguments, keywords, Err(&failure.message));
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
pub fn evaluate_rule(
    rule: &RuleDefinition,
    retrieval: &mut S5Retrieval,
    run_id: &str,
    context_period: &str,
    institution_code: &str,
    institution_name: &str,
    configuration_version: &str,
) -> S5RuleOutcome {
    let _ = (
        run_id,
        context_period,
        institution_code,
        institution_name,
        configuration_version,
    );
    let mut outcome = S5RuleOutcome {
        status: "NOT_EXECUTED".to_owned(),
        triggered: None,
        description: rule.description.clone(),
        ..Default::default()
    };

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
                        let mut not_applicable = S5RuleOutcome {
                            status: "NOT_APPLICABLE".to_owned(),
                            error_message: "规则适用表达式判定为不适用".to_owned(),
                            evidence: unique_evidence(&[applicable]),
                            ..Default::default()
                        };
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

        let mut outcome = S5RuleOutcome {
            status: if audit_passes { "PASSED" } else { "TRIGGERED" }.to_owned(),
            triggered: Some(!audit_passes),
            description,
            ..Default::default()
        };
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
                        if STRUCTURE_ERROR_CODES.contains(&code.as_str()) {
                            ("STRUCTURE_ERROR".to_owned(), code.clone())
                        } else if CONFIG_ERROR_CODES.contains(&code.as_str()) {
                            ("CONFIG_ERROR".to_owned(), code.clone())
                        } else {
                            (
                                "RETRIEVAL_FAILED".to_owned(),
                                error_code(&error.message, "RETRIEVAL_FAILED"),
                            )
                        }
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
