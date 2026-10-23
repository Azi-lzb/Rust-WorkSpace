//! S2 expression and cumulative-rule evaluation over already imported records.
//!
//! This is a business-calculation slice only. It does not load workbooks,
//! parse the configuration workbook, orchestrate a report run, or write XLSX.

use crate::{
    s2_expression::{decimal_plain, evaluate, python_float_repr},
    s2_import::{ImportedValue, IndicatorRecord, S2IndicatorConfig, S2SourceRef, S2UnitSettings},
    s2_period::{compare_decimal, decimal_fixed, decimal_subtract},
    s2_validation::S2AuditFinding,
    unit_conversion::{as_decimal, convert_amount, is_amount_data_type, DecimalTuple},
};
use anyhow::{anyhow, Result};
use serde::{Deserialize, Serialize};
use serde_json::{Number, Value};
use sha1::{Digest, Sha1};
use std::{
    cmp::Ordering,
    collections::{BTreeMap, BTreeSet, HashMap},
};

const STATUS_ABNORMAL: &str = "异常";

#[derive(Clone, Debug, Deserialize, Default)]
pub struct S2RuleEngineConfig {
    #[serde(default)]
    pub indicators: HashMap<String, S2IndicatorConfig>,
    #[serde(default)]
    pub rules: Vec<S2Rule>,
    #[serde(default)]
    pub unit_settings: S2UnitSettings,
}

#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
pub struct S2Rule {
    pub rule_id: String,
    pub rule_type: String,
    pub description: String,
    pub expression: String,
    #[serde(default = "default_severity")]
    pub severity: String,
    #[serde(default)]
    pub threshold: Option<f64>,
    #[serde(default)]
    pub invert: bool,
    #[serde(default)]
    pub disabled: bool,
    #[serde(default)]
    pub note: String,
    #[serde(default)]
    pub source: Option<S2SourceRef>,
}

fn default_severity() -> String {
    "软性".to_owned()
}

/// Evaluate the rules for one already-imported reporting period pair.
///
/// Input slices must follow the source importer's record order. The source
/// Python engine receives insertion-ordered dictionaries; preserving this
/// order keeps its per-institution anchor and duplicate-indicator behavior.
pub fn evaluate_rules(
    current: &[IndicatorRecord],
    previous: &[IndicatorRecord],
    config: &S2RuleEngineConfig,
) -> Result<(Vec<S2AuditFinding>, Vec<S2AuditFinding>)> {
    let current_by = group_by_institution(current);
    let previous_by = group_by_institution(previous);
    let institutions = current_by
        .keys()
        .chain(previous_by.keys())
        .cloned()
        .collect::<BTreeSet<_>>();
    let amount_codes = config
        .indicators
        .iter()
        .filter(|(_, indicator)| is_amount_data_type(&indicator.data_type))
        .map(|(code, _)| code.clone())
        .collect::<BTreeSet<_>>();
    let mut findings = Vec::new();
    let mut quality = Vec::new();

    for rule in &config.rules {
        if rule.disabled {
            continue;
        }
        for institution_id in &institutions {
            let current_records = current_by.get(institution_id);
            let previous_records = previous_by.get(institution_id);
            let anchor = current_records
                .and_then(OrderedRecords::first)
                .or_else(|| previous_records.and_then(OrderedRecords::first));
            let Some(anchor) = anchor else { continue };

            if rule.rule_type.starts_with("累计不降") {
                let codes = split_cumulative_codes(&rule.expression);
                for code in codes {
                    let (Some(current_record), Some(previous_record)) = (
                        current_records.and_then(|records| records.get(&code)),
                        previous_records.and_then(|records| records.get(&code)),
                    ) else {
                        continue;
                    };
                    if !number_less(
                        &current_record.normalized_value,
                        &previous_record.normalized_value,
                    )
                    .unwrap_or(false)
                    {
                        continue;
                    }
                    let difference = number_subtract(
                        &current_record.normalized_value,
                        &previous_record.normalized_value,
                    )?;
                    let current_display = display_number(&current_record.normalized_value);
                    let previous_display = display_number(&previous_record.normalized_value);
                    let mut finding = S2AuditFinding::empty(finding_id(&[
                        "规则",
                        &rule.rule_id,
                        institution_id,
                        &code,
                        &current_record.period,
                    ]));
                    finding.audit_type = "规则".to_owned();
                    finding.status = STATUS_ABNORMAL.to_owned();
                    finding.severity = rule.severity.clone();
                    finding.region = current_record.region.clone();
                    finding.handling_branch = current_record.handling_branch.clone();
                    finding.institution_type = current_record.institution_type.clone();
                    finding.institution_name = current_record.institution_name.clone();
                    finding.social_credit_code = current_record.social_credit_code.clone();
                    finding.form_code = current_record.form_code.clone();
                    finding.form_codes = if current_record.form_code.is_empty() {
                        Vec::new()
                    } else {
                        vec![current_record.form_code.clone()]
                    };
                    finding.period = current_record.period.clone();
                    finding.indicator_code = code.clone();
                    finding.indicator_name = current_record.indicator_name.clone();
                    finding.rule_id = rule.rule_id.clone();
                    finding.rule_description = rule.description.clone();
                    finding.description = rule.description.clone();
                    finding.current_value = raw_value_json(&current_record.normalized_value);
                    finding.previous_value = raw_value_json(&previous_record.normalized_value);
                    finding.difference_value = difference;
                    finding.calculation_trace =
                        format!("{}: {} < {}", code, current_display, previous_display);
                    finding.rule_type = rule_type(rule);
                    finding.value_details = format!(
                        "{} {}: 本期={}；上期={}",
                        code, current_record.indicator_name, current_display, previous_display
                    );
                    finding.source_refs = vec![current_record.source.clone()];
                    finding.related_records = related_records(
                        current_records,
                        previous_records,
                        std::iter::once(code.as_str()),
                    );
                    findings.push(finding);
                }
                continue;
            }

            let evaluated = evaluate_rule(
                rule,
                current_records,
                previous_records,
                config,
                &amount_codes,
            )?;
            if !evaluated.error.is_empty() {
                let error = evaluated.error;
                let mut finding = S2AuditFinding::empty(finding_id(&[
                    "规则配置",
                    &rule.rule_id,
                    institution_id,
                    &error,
                ]));
                finding.audit_type = "校验规则".to_owned();
                finding.status = STATUS_ABNORMAL.to_owned();
                finding.severity = "软性".to_owned();
                finding.institution_name = anchor.institution_name.clone();
                finding.social_credit_code = anchor.social_credit_code.clone();
                finding.period = anchor.period.clone();
                finding.rule_id = rule.rule_id.clone();
                finding.description =
                    format!("规则无法执行：{}（规则：{}）", error, rule.description);
                finding.calculation_trace =
                    expression_trace(&rule.expression, &evaluated.display, false);
                finding.rule_type = rule_type(rule);
                finding.value_details =
                    expression_value_details(&rule.expression, current_records, previous_records);
                quality.push(finding);
                continue;
            }
            if !evaluated.hit {
                continue;
            }

            let related = related_record_refs(
                current_records,
                previous_records,
                evaluated.involved.iter().map(String::as_str),
            );
            let forms = related
                .iter()
                .filter_map(|record| {
                    (!record.form_code.is_empty()).then_some(record.form_code.clone())
                })
                .collect::<BTreeSet<_>>()
                .into_iter()
                .collect::<Vec<_>>();
            let primary = related.first().copied().unwrap_or(anchor);
            let mut finding = S2AuditFinding::empty(finding_id(&[
                "规则",
                &rule.rule_id,
                institution_id,
                &anchor.period,
            ]));
            finding.audit_type = "规则".to_owned();
            finding.status = STATUS_ABNORMAL.to_owned();
            finding.severity = rule.severity.clone();
            finding.region = anchor.region.clone();
            finding.handling_branch = anchor.handling_branch.clone();
            finding.institution_type = anchor.institution_type.clone();
            finding.institution_name = anchor.institution_name.clone();
            finding.social_credit_code = anchor.social_credit_code.clone();
            finding.form_code = primary.form_code.clone();
            finding.form_codes = forms;
            finding.period = anchor.period.clone();
            finding.indicator_code = primary.indicator_code.clone();
            finding.indicator_name = primary.indicator_name.clone();
            finding.rule_id = rule.rule_id.clone();
            finding.rule_description = rule.description.clone();
            finding.description = rule.description.clone();
            finding.calculation_trace =
                expression_trace(&rule.expression, &evaluated.display, rule.invert);
            finding.rule_type = rule_type(rule);
            finding.value_details =
                expression_value_details(&rule.expression, current_records, previous_records);
            finding.source_refs = match &rule.source {
                Some(source) => vec![primary.source.clone(), source.clone()],
                None => vec![primary.source.clone()],
            };
            finding.related_records = related_records(
                current_records,
                previous_records,
                evaluated.involved.iter().map(String::as_str),
            );
            findings.push(finding);
        }
    }
    Ok((findings, quality))
}

#[derive(Default)]
struct OrderedRecords<'a> {
    codes: Vec<String>,
    by_code: HashMap<String, &'a IndicatorRecord>,
}

impl<'a> OrderedRecords<'a> {
    fn insert(&mut self, record: &'a IndicatorRecord) {
        if !self.by_code.contains_key(&record.indicator_code) {
            self.codes.push(record.indicator_code.clone());
        }
        self.by_code.insert(record.indicator_code.clone(), record);
    }

    fn first(&self) -> Option<&'a IndicatorRecord> {
        self.codes
            .first()
            .and_then(|code| self.by_code.get(code))
            .copied()
    }

    fn get(&self, code: &str) -> Option<&'a IndicatorRecord> {
        self.by_code.get(code).copied()
    }
}

fn group_by_institution<'a>(
    records: &'a [IndicatorRecord],
) -> BTreeMap<String, OrderedRecords<'a>> {
    let mut groups = BTreeMap::<String, OrderedRecords<'a>>::new();
    for record in records {
        groups
            .entry(record.institution_id.clone())
            .or_default()
            .insert(record);
    }
    groups
}

fn split_cumulative_codes(expression: &str) -> Vec<String> {
    expression
        .split(|character: char| {
            matches!(character, ',' | '，' | '、' | ';' | '；') || character.is_whitespace()
        })
        .filter(|part| !part.is_empty())
        .map(str::to_owned)
        .collect()
}

#[derive(Clone, Debug)]
enum PythonNumber {
    Float(f64),
    Decimal(DecimalTuple),
}

fn python_number(value: &ImportedValue) -> PythonNumber {
    match value {
        ImportedValue::Decimal(value) => PythonNumber::Decimal(value.clone()),
        ImportedValue::Empty => PythonNumber::Float(0.0),
        ImportedValue::Text(value) => value
            .parse::<f64>()
            .map(PythonNumber::Float)
            .unwrap_or(PythonNumber::Float(0.0)),
        ImportedValue::Number(value) | ImportedValue::DateTimeSerial(value) => {
            PythonNumber::Float(*value)
        }
        ImportedValue::Integer(value) => PythonNumber::Float(*value as f64),
        ImportedValue::Boolean(value) => PythonNumber::Float(if *value { 1.0 } else { 0.0 }),
        ImportedValue::DateTimeIso(_) | ImportedValue::DurationIso(_) | ImportedValue::Error(_) => {
            PythonNumber::Float(0.0)
        }
    }
}

fn number_less(left: &ImportedValue, right: &ImportedValue) -> Option<bool> {
    let left = python_number(left);
    let right = python_number(right);
    Some(match (&left, &right) {
        (PythonNumber::Float(left), PythonNumber::Float(right)) => {
            return left
                .partial_cmp(right)
                .map(|ordering| ordering == Ordering::Less)
        }
        (PythonNumber::Decimal(left), PythonNumber::Decimal(right)) => {
            compare_decimal(left, right) == Ordering::Less
        }
        (PythonNumber::Decimal(left), PythonNumber::Float(right)) => {
            let right = float_to_decimal(*right)?;
            compare_decimal(left, &right) == Ordering::Less
        }
        (PythonNumber::Float(left), PythonNumber::Decimal(right)) => {
            let left = float_to_decimal(*left)?;
            compare_decimal(&left, right) == Ordering::Less
        }
    })
}

fn number_subtract(left: &ImportedValue, right: &ImportedValue) -> Result<Value> {
    match (python_number(left), python_number(right)) {
        (PythonNumber::Float(left), PythonNumber::Float(right)) => Number::from_f64(left - right)
            .map(Value::Number)
            .ok_or_else(|| anyhow!("difference is not a finite JSON number")),
        (PythonNumber::Decimal(left), PythonNumber::Decimal(right)) => {
            let value = decimal_subtract(&left, &right)?;
            Ok(serde_json::to_value(ImportedValue::Decimal(value))?)
        }
        _ => anyhow::bail!("unsupported operand type(s) for -: Decimal and float"),
    }
}

fn float_to_decimal(value: f64) -> Option<DecimalTuple> {
    if !value.is_finite() {
        return None;
    }
    let bits = value.to_bits();
    let sign = (bits >> 63) as u8;
    let exponent_bits = ((bits >> 52) & 0x7ff) as i32;
    let fraction = bits & ((1_u64 << 52) - 1);
    if exponent_bits == 0 && fraction == 0 {
        return Some(DecimalTuple {
            sign,
            digits: vec![0],
            exponent: 0,
        });
    }
    let (mantissa, exponent_two) = if exponent_bits == 0 {
        (fraction, -1022 - 52)
    } else {
        ((1_u64 << 52) | fraction, exponent_bits - 1023 - 52)
    };
    let mut digits = mantissa
        .to_string()
        .bytes()
        .map(|byte| byte - b'0')
        .collect::<Vec<_>>();
    if exponent_two >= 0 {
        for _ in 0..exponent_two {
            digits = multiply_small(&digits, 2);
        }
        Some(DecimalTuple {
            sign,
            digits,
            exponent: 0,
        })
    } else {
        for _ in 0..-exponent_two {
            digits = multiply_small(&digits, 5);
        }
        Some(DecimalTuple {
            sign,
            digits,
            exponent: i64::from(exponent_two),
        })
    }
}

fn multiply_small(digits: &[u8], multiplier: u8) -> Vec<u8> {
    let mut output = Vec::with_capacity(digits.len() + 2);
    let mut carry = 0_u16;
    for digit in digits.iter().rev() {
        let product = u16::from(*digit) * u16::from(multiplier) + carry;
        output.push((product % 10) as u8);
        carry = product / 10;
    }
    while carry > 0 {
        output.push((carry % 10) as u8);
        carry /= 10;
    }
    output.reverse();
    output
}

struct EvaluatedRule {
    hit: bool,
    display: String,
    error: String,
    involved: BTreeSet<String>,
}

fn evaluate_rule(
    rule: &S2Rule,
    current: Option<&OrderedRecords<'_>>,
    previous: Option<&OrderedRecords<'_>>,
    config: &S2RuleEngineConfig,
    amount_codes: &BTreeSet<String>,
) -> Result<EvaluatedRule> {
    let substitution = substitute_placeholders(&rule.expression, current, previous);
    if !substitution.errors.is_empty() {
        return Ok(EvaluatedRule {
            hit: false,
            display: substitution.display,
            error: substitution
                .errors
                .into_iter()
                .collect::<BTreeSet<_>>()
                .into_iter()
                .collect::<Vec<_>>()
                .join("；"),
            involved: substitution.involved,
        });
    }

    let mut expression = substitution.evaluable;
    if has_thd_token(&expression) {
        let Some(threshold) = rule.threshold else {
            return Ok(EvaluatedRule {
                hit: false,
                display: substitution.display,
                error: "表达式使用 Thd 但规则未填写 Thd值(万元)".to_owned(),
                involved: substitution.involved,
            });
        };
        let mut threshold = as_decimal(&python_float_repr(threshold))?;
        if substitution
            .involved
            .iter()
            .any(|code| amount_codes.contains(code))
        {
            threshold = convert_amount(
                &decimal_plain(&threshold),
                "万元",
                &config.unit_settings.output_file_unit,
            )?;
        }
        expression = replace_thd_token(&expression, &decimal_plain(&threshold));
    }

    let result = evaluate(&expression);
    Ok(match result {
        Ok(hit) => EvaluatedRule {
            hit: if rule.invert { !hit } else { hit },
            display: substitution.display,
            error: String::new(),
            involved: substitution.involved,
        },
        Err(error) => EvaluatedRule {
            hit: false,
            display: substitution.display,
            error: format!("表达式无法执行：{error}"),
            involved: substitution.involved,
        },
    })
}

struct Substitution {
    evaluable: String,
    display: String,
    errors: Vec<String>,
    involved: BTreeSet<String>,
}

#[derive(Clone, Debug)]
struct Placeholder {
    start: usize,
    end: usize,
    period_side: char,
    code: String,
}

fn placeholders(expression: &str) -> Vec<Placeholder> {
    let chars = expression.chars().collect::<Vec<_>>();
    let mut result = Vec::new();
    let mut index = 0;
    while index < chars.len() {
        if !matches!(chars[index], '[' | '{') {
            index += 1;
            continue;
        }
        let mut close = index + 1;
        while close < chars.len() && !matches!(chars[close], ']' | '}') {
            close += 1;
        }
        if close >= chars.len() || close == index + 1 {
            index += 1;
            continue;
        }
        let inner = chars[index + 1..close].iter().collect::<String>();
        result.push(Placeholder {
            start: index,
            end: close + 1,
            period_side: chars[index],
            code: indicator_code(&inner),
        });
        index = close + 1;
    }
    result
}

fn substitute_placeholders(
    expression: &str,
    current: Option<&OrderedRecords<'_>>,
    previous: Option<&OrderedRecords<'_>>,
) -> Substitution {
    let characters = expression.chars().collect::<Vec<_>>();
    let placeholders = placeholders(expression);
    let mut evaluable = String::new();
    let mut display = String::new();
    let mut errors = Vec::new();
    let mut involved = BTreeSet::new();
    let mut cursor = 0;
    for placeholder in placeholders {
        evaluable.extend(characters[cursor..placeholder.start].iter());
        display.extend(characters[cursor..placeholder.start].iter());
        cursor = placeholder.end;
        involved.insert(placeholder.code.clone());
        let records = if placeholder.period_side == '[' {
            current
        } else {
            previous
        };
        let Some(record) = records.and_then(|records| records.get(&placeholder.code)) else {
            errors.push(format!("缺少指标 {} 的当期数据", placeholder.code));
            evaluable.push_str("0.0");
            display.push_str("0.00");
            continue;
        };
        match expression_number(record) {
            Ok((number_text, shown)) => {
                evaluable.push_str(&number_text);
                display.push_str(&shown);
            }
            Err(error) => {
                errors.push(error);
                evaluable.push_str("0.0");
                display.push_str(&expression_error_display(record));
            }
        }
    }
    evaluable.extend(characters[cursor..].iter());
    display.extend(characters[cursor..].iter());
    Substitution {
        evaluable,
        display,
        errors,
        involved,
    }
}

fn expression_error_display(record: &IndicatorRecord) -> String {
    match &record.normalized_value {
        ImportedValue::Text(value) if !value.trim().is_empty() => value.clone(),
        _ => "0.00".to_owned(),
    }
}

fn expression_number(record: &IndicatorRecord) -> std::result::Result<(String, String), String> {
    match &record.normalized_value {
        ImportedValue::Empty => Err(format!("缺少指标 {} 的有效数值", record.indicator_code)),
        ImportedValue::Text(value) if value.trim().is_empty() => {
            Err(format!("缺少指标 {} 的有效数值", record.indicator_code))
        }
        ImportedValue::Text(value) => {
            let cleaned = value.replace(',', "").replace('_', "");
            let number = cleaned.trim().parse::<f64>().map_err(|_| {
                format!(
                    "指标 {} 的值不是数值：{}",
                    record.indicator_code,
                    python_repr_string(value)
                )
            })?;
            let text = python_float_repr(number);
            Ok((text, format!("{number:.2}")))
        }
        ImportedValue::Decimal(value) => Ok((decimal_plain(value), decimal_fixed(value, 2))),
        ImportedValue::Number(value) | ImportedValue::DateTimeSerial(value) => {
            let text = python_float_repr(*value);
            Ok((text, format!("{value:.2}")))
        }
        ImportedValue::Integer(value) => {
            let value = *value as f64;
            Ok((python_float_repr(value), format!("{value:.2}")))
        }
        ImportedValue::Boolean(value) => {
            let text = if *value { "1.0" } else { "0.0" };
            let decimal = as_decimal(text).expect("finite boolean decimal");
            Ok((text.to_owned(), decimal_fixed(&decimal, 2)))
        }
        ImportedValue::DateTimeIso(_) | ImportedValue::DurationIso(_) | ImportedValue::Error(_) => {
            let decimal = as_decimal("0.0").expect("zero decimal");
            Ok(("0.0".to_owned(), decimal_fixed(&decimal, 2)))
        }
    }
}

fn python_repr_string(value: &str) -> String {
    let quote = if !value.contains('\'') {
        '\''
    } else if !value.contains('"') {
        '"'
    } else {
        '\''
    };
    let mut output = String::with_capacity(value.len() + 2);
    output.push(quote);
    for character in value.chars() {
        match character {
            '\\' => output.push_str("\\\\"),
            '\n' => output.push_str("\\n"),
            '\r' => output.push_str("\\r"),
            '\t' => output.push_str("\\t"),
            character if character == quote => {
                output.push('\\');
                output.push(character);
            }
            character => output.push(character),
        }
    }
    output.push(quote);
    output
}

fn indicator_code(inner: &str) -> String {
    let text = inner.trim();
    if !text.contains(',') {
        return text.to_owned();
    }
    let parts = text.split(',').map(str::trim).collect::<Vec<_>>();
    if parts.len() >= 4 && !parts[3].is_empty() {
        return parts[3].to_owned();
    }
    parts
        .iter()
        .rev()
        .find(|part| !part.is_empty())
        .copied()
        .unwrap_or(text)
        .to_owned()
}

fn has_thd_token(text: &str) -> bool {
    identifier_tokens(text)
        .iter()
        .any(|token| token.eq_ignore_ascii_case("Thd"))
}

fn replace_thd_token(text: &str, replacement: &str) -> String {
    let characters = text.chars().collect::<Vec<_>>();
    let mut output = String::new();
    let mut index = 0;
    while index < characters.len() {
        if characters[index] == '_' || characters[index].is_alphabetic() {
            let start = index;
            index += 1;
            while index < characters.len()
                && (characters[index] == '_' || characters[index].is_alphanumeric())
            {
                index += 1;
            }
            let token = characters[start..index].iter().collect::<String>();
            if token.eq_ignore_ascii_case("Thd") {
                output.push_str(replacement);
            } else {
                output.push_str(&token);
            }
        } else {
            output.push(characters[index]);
            index += 1;
        }
    }
    output
}

fn identifier_tokens(text: &str) -> Vec<String> {
    let characters = text.chars().collect::<Vec<_>>();
    let mut tokens = Vec::new();
    let mut index = 0;
    while index < characters.len() {
        if characters[index] == '_' || characters[index].is_alphabetic() {
            let start = index;
            index += 1;
            while index < characters.len()
                && (characters[index] == '_' || characters[index].is_alphanumeric())
            {
                index += 1;
            }
            tokens.push(characters[start..index].iter().collect());
        } else {
            index += 1;
        }
    }
    tokens
}

fn rule_type(rule: &S2Rule) -> String {
    if rule.rule_type.starts_with("累计不降") {
        "累计不降".to_owned()
    } else if rule.rule_type == "表达式" {
        "表达式".to_owned()
    } else {
        "未识别".to_owned()
    }
}

fn expression_trace(expression: &str, rendered: &str, invert: bool) -> String {
    let suffix = if invert { "（取反）" } else { "" };
    format!("配置表达式：{expression}；代入表达式：{rendered}{suffix}")
}

fn expression_value_details(
    expression: &str,
    current: Option<&OrderedRecords<'_>>,
    previous: Option<&OrderedRecords<'_>>,
) -> String {
    let mut entries = Vec::new();
    let mut seen = BTreeSet::new();
    for placeholder in placeholders(expression) {
        let period_label = if placeholder.period_side == '[' {
            "本期"
        } else {
            "上期"
        };
        if !seen.insert((period_label, placeholder.code.clone())) {
            continue;
        }
        let records = if placeholder.period_side == '[' {
            current
        } else {
            previous
        };
        let record = records.and_then(|records| records.get(&placeholder.code));
        let mut label = format!("{} {}", period_label, placeholder.code);
        if let Some(record) = record.filter(|record| !record.indicator_name.is_empty()) {
            label.push(' ');
            label.push_str(&record.indicator_name);
        }
        entries.push(format!("{}={}", label, display_value(record)));
    }
    entries.join("；")
}

fn display_value(record: Option<&IndicatorRecord>) -> String {
    let Some(record) = record else {
        return "缺少指标".to_owned();
    };
    match &record.normalized_value {
        ImportedValue::Empty => "无数据".to_owned(),
        ImportedValue::Text(value) if value.is_empty() => "无数据".to_owned(),
        value => display_number(value),
    }
}

fn display_number(value: &ImportedValue) -> String {
    match value {
        ImportedValue::Boolean(value) => if *value { "True" } else { "False" }.to_owned(),
        ImportedValue::Integer(value) => format!("{value}.00"),
        ImportedValue::Number(value) | ImportedValue::DateTimeSerial(value) => {
            format!("{value:.2}")
        }
        ImportedValue::Decimal(value) => decimal_fixed(value, 2),
        ImportedValue::Text(value) => {
            let cleaned = value.replace(',', "");
            match as_decimal(cleaned.trim()) {
                Ok(number) => decimal_fixed(&number, 2),
                Err(_) => value.clone(),
            }
        }
        ImportedValue::Empty => "".to_owned(),
        ImportedValue::DateTimeIso(value)
        | ImportedValue::DurationIso(value)
        | ImportedValue::Error(value) => value.clone(),
    }
}

fn related_record_refs<'a>(
    current: Option<&'a OrderedRecords<'a>>,
    previous: Option<&'a OrderedRecords<'a>>,
    codes: impl Iterator<Item = &'a str>,
) -> Vec<&'a IndicatorRecord> {
    let mut by_code = BTreeMap::new();
    if let Some(previous) = previous {
        for (code, record) in &previous.by_code {
            by_code.insert(code.as_str(), *record);
        }
    }
    if let Some(current) = current {
        for (code, record) in &current.by_code {
            by_code.insert(code.as_str(), *record);
        }
    }
    codes
        .filter_map(|code| by_code.get(code).copied())
        .collect()
}

fn related_records<'a>(
    current: Option<&'a OrderedRecords<'a>>,
    previous: Option<&'a OrderedRecords<'a>>,
    codes: impl Iterator<Item = &'a str>,
) -> Vec<Value> {
    related_record_refs(current, previous, codes)
        .into_iter()
        .map(|record| {
            serde_json::json!({
                "indicator_code": record.indicator_code,
                "indicator_name": record.indicator_name,
                "form_code": record.form_code,
                "period": record.period,
                "value": raw_value_json(&record.normalized_value),
                "source": record.source,
            })
        })
        .collect()
}

fn raw_value_json(value: &ImportedValue) -> Value {
    match value {
        ImportedValue::Empty => Value::Null,
        ImportedValue::Text(value) => Value::String(value.clone()),
        ImportedValue::Number(value) | ImportedValue::DateTimeSerial(value) => {
            Number::from_f64(*value)
                .map(Value::Number)
                .unwrap_or(Value::Null)
        }
        ImportedValue::Integer(value) => Value::from(*value),
        ImportedValue::Boolean(value) => Value::Bool(*value),
        ImportedValue::Decimal(value) => {
            serde_json::to_value(ImportedValue::Decimal(value.clone()))
                .expect("decimal tuple serializes")
        }
        ImportedValue::DateTimeIso(value)
        | ImportedValue::DurationIso(value)
        | ImportedValue::Error(value) => Value::String(value.clone()),
    }
}

fn finding_id(parts: &[&str]) -> String {
    let digest = Sha1::digest(parts.join("|").as_bytes());
    digest[..8]
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect()
}

#[cfg(test)]
mod tests {
    use super::{indicator_code, split_cumulative_codes};

    #[test]
    fn legacy_eight_segment_indicator_code_selects_fourth_segment() {
        assert_eq!(indicator_code(",,,20203045,"), "20203045");
        assert_eq!(indicator_code(",,,20203045,"), "20203045");
    }

    #[test]
    fn cumulative_code_lists_accept_legacy_separators() {
        assert_eq!(
            split_cumulative_codes("A，B、C;D；E F"),
            ["A", "B", "C", "D", "E", "F"]
        );
    }
}
