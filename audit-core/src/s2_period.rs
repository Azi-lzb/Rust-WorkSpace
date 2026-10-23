//! Period comparison rules migrated from the read-only S2 Python baseline.
//!
//! Monetary values stay as base-10 decimal tuples through subtraction,
//! thresholds, and band matching. Percentage outputs and Python-float inputs
//! retain the source engine's binary-float semantics. Decimal division rounds
//! to 28 significant digits with half-even rounding, matching Python's default
//! `Decimal` context used by the source implementation.

use crate::{
    s2_import::{ImportedValue, IndicatorRecord},
    s2_validation::S2AuditFinding,
    unit_conversion::{as_decimal, DecimalTuple, UnitConversionError, AMOUNT_DATA_TYPES},
};
use anyhow::{bail, Result};
use serde::{Deserialize, Serialize};
use serde_json::{Number, Value};
use sha1::{Digest, Sha1};
use std::cmp::Ordering;
use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};

const STATUS_NORMAL: &str = "正常";
const STATUS_ABNORMAL: &str = "异常";
const AUDIT_TYPE: &str = "环比";
const DECIMAL_CONTEXT_PRECISION: usize = 28;

#[derive(Clone, Debug, Deserialize, Default)]
pub struct S2PeriodConfig {
    #[serde(default)]
    pub indicators: HashMap<String, S2PeriodIndicatorConfig>,
    #[serde(default)]
    pub bands: Vec<S2PeriodBandConfig>,
    #[serde(default = "hard_severity")]
    pub text_change_severity: String,
    #[serde(default = "hard_severity")]
    pub current_missing_severity: String,
    #[serde(default = "hard_severity")]
    pub previous_missing_severity: String,
}

fn hard_severity() -> String {
    "硬性".to_owned()
}

#[derive(Clone, Debug, Deserialize)]
#[serde(default)]
pub struct S2PeriodIndicatorConfig {
    pub data_type: String,
    pub disabled: bool,
    pub period_enabled: bool,
    pub period_strategy_group: String,
    pub min_change_value: Option<f64>,
}

impl Default for S2PeriodIndicatorConfig {
    fn default() -> Self {
        Self {
            data_type: String::new(),
            disabled: false,
            period_enabled: true,
            period_strategy_group: String::new(),
            min_change_value: None,
        }
    }
}

#[derive(Clone, Debug, Deserialize, Default, PartialEq, Serialize)]
#[serde(default)]
pub struct S2PeriodBandConfig {
    pub rule_id: String,
    pub lower: Option<f64>,
    pub upper: Option<f64>,
    pub description: String,
    pub severity: String,
    pub disabled: bool,
    pub strategy_group: String,
    pub data_type: String,
}

#[derive(Clone, Debug, PartialEq)]
enum Numeric {
    Float(f64),
    Decimal(DecimalTuple),
}

impl Numeric {
    fn from_imported(value: &ImportedValue) -> Option<Self> {
        match value {
            ImportedValue::Empty | ImportedValue::Boolean(_) => None,
            ImportedValue::Decimal(value) => Some(Self::Decimal(value.clone())),
            ImportedValue::Integer(value) => Some(Self::Float(*value as f64)),
            ImportedValue::Number(value) => Some(Self::Float(*value)),
            ImportedValue::Text(value) => {
                let cleaned = value.replace(',', "").replace('_', "");
                cleaned.trim().parse::<f64>().ok().map(Self::Float)
            }
            ImportedValue::DateTimeSerial(_)
            | ImportedValue::DateTimeIso(_)
            | ImportedValue::DurationIso(_)
            | ImportedValue::Error(_) => None,
        }
    }

    fn compare(&self, other: &Self) -> Option<Ordering> {
        match (self, other) {
            (Self::Float(left), Self::Float(right)) => left.partial_cmp(right),
            (Self::Decimal(left), Self::Decimal(right)) => Some(compare_decimal(left, right)),
            (Self::Decimal(left), Self::Float(right)) => {
                float_to_decimal(*right).map(|right| compare_decimal(left, &right))
            }
            (Self::Float(left), Self::Decimal(right)) => {
                float_to_decimal(*left).map(|left| compare_decimal(&left, right))
            }
        }
    }

    fn subtract(&self, other: &Self) -> Result<Self> {
        match (self, other) {
            (Self::Float(left), Self::Float(right)) => Ok(Self::Float(left - right)),
            (Self::Decimal(left), Self::Decimal(right)) => {
                Ok(Self::Decimal(decimal_subtract(left, right)?))
            }
            _ => bail!("环比数值类型不一致：Decimal 与 float 不能直接运算"),
        }
    }

    fn absolute(&self) -> Self {
        match self {
            Self::Float(value) => Self::Float(value.abs()),
            Self::Decimal(value) => {
                let mut value = value.clone();
                value.sign = 0;
                Self::Decimal(value)
            }
        }
    }

    fn is_zero(&self) -> bool {
        match self {
            Self::Float(value) => *value == 0.0,
            Self::Decimal(value) => value.digits.iter().all(|digit| *digit == 0),
        }
    }

    fn divide(&self, other: &Self) -> Result<Self> {
        match (self, other) {
            (Self::Float(left), Self::Float(right)) => Ok(Self::Float(left / right)),
            (Self::Decimal(left), Self::Decimal(right)) => {
                Ok(Self::Decimal(decimal_divide(left, right)?))
            }
            _ => bail!("环比数值类型不一致：Decimal 与 float 不能直接运算"),
        }
    }

    fn multiply_hundred(&self) -> Result<Self> {
        match self {
            Self::Float(value) => Ok(Self::Float(value * 100.0)),
            Self::Decimal(value) => {
                let mut value = value.clone();
                value.exponent = value
                    .exponent
                    .checked_add(2)
                    .ok_or_else(|| anyhow::anyhow!("Decimal 指数超出可处理范围"))?;
                Ok(Self::Decimal(value))
            }
        }
    }

    fn divide_hundred(&self) -> Result<Self> {
        match self {
            Self::Float(value) => Ok(Self::Float(value / 100.0)),
            Self::Decimal(value) => {
                let mut value = value.clone();
                value.exponent = value
                    .exponent
                    .checked_sub(2)
                    .ok_or_else(|| anyhow::anyhow!("Decimal 指数超出可处理范围"))?;
                Ok(Self::Decimal(value))
            }
        }
    }

    fn to_json(&self) -> Value {
        match self {
            Self::Float(value) => Number::from_f64(*value)
                .map(Value::Number)
                .unwrap_or(Value::Null),
            Self::Decimal(value) => serde_json::to_value(ImportedValue::Decimal(value.clone()))
                .expect("decimal tuple serializes"),
        }
    }

    fn python_repr(&self) -> String {
        match self {
            Self::Float(value) => {
                let mut text = value.to_string();
                if value.is_finite() && !text.contains(['.', 'e', 'E']) {
                    text.push_str(".0");
                }
                text
            }
            Self::Decimal(value) => decimal_to_string(value),
        }
    }

    fn trace_number(&self) -> String {
        match self {
            Self::Float(value) => format!("{value:.2}"),
            Self::Decimal(value) => decimal_fixed(value, 2),
        }
    }

    fn trace_percent(&self) -> String {
        match self {
            Self::Float(value) => format!("{:.2}%", value * 100.0),
            Self::Decimal(_) => {
                let scaled = self
                    .multiply_hundred()
                    .expect("adding two to Decimal exponent");
                format!("{}%", scaled.trace_number())
            }
        }
    }
}

/// Compare the same institution/indicator pair across the two supplied periods.
/// Each input slice follows source insertion order; if it contains duplicate
/// comparison keys, the last record wins just as in the Python dict-comprehension.
pub fn compare_periods(
    current: &[IndicatorRecord],
    previous: &[IndicatorRecord],
    config: &S2PeriodConfig,
) -> Result<Vec<S2AuditFinding>> {
    let current_by = comparison_index(current);
    let previous_by = comparison_index(previous);
    let keys = current_by
        .keys()
        .chain(previous_by.keys())
        .cloned()
        .collect::<BTreeSet<_>>();
    let text_severity = if config.text_change_severity.is_empty() {
        "硬性"
    } else {
        &config.text_change_severity
    };
    let current_missing_severity = if config.current_missing_severity.is_empty() {
        "硬性"
    } else {
        &config.current_missing_severity
    };
    let previous_missing_severity = if config.previous_missing_severity.is_empty() {
        "硬性"
    } else {
        &config.previous_missing_severity
    };
    let mut findings = Vec::new();

    for key in keys {
        let current_record = current_by.get(&key).copied();
        let previous_record = previous_by.get(&key).copied();
        let anchor = current_record
            .or(previous_record)
            .expect("key has a record");
        let indicator = config.indicators.get(&anchor.indicator_code);
        if indicator.is_some_and(|item| item.disabled || !item.period_enabled) {
            continue;
        }
        let attribute = indicator
            .map(|item| item.data_type.as_str())
            .filter(|value| !value.is_empty())
            .unwrap_or(&anchor.value_type);
        let (current_side_severity, previous_side_severity) = if attribute == "文字" {
            (text_severity, text_severity)
        } else {
            (current_missing_severity, previous_missing_severity)
        };

        let Some(current_record) = current_record else {
            let previous_record = previous_record.expect("previous record exists");
            let mut finding = base_finding(
                previous_record,
                STATUS_ABNORMAL,
                current_side_severity,
                "本期无，上期有",
                attribute,
            );
            finding.previous_value = raw_value_json(&previous_record.normalized_value);
            finding.band.clear();
            finding.retrieval_note = join_notes(&[
                retrieval_note(None, "本期"),
                retrieval_note(Some(previous_record), "上期"),
            ]);
            findings.push(finding);
            continue;
        };
        let Some(previous_record) = previous_record else {
            let mut finding = base_finding(
                current_record,
                STATUS_ABNORMAL,
                previous_side_severity,
                "本期有，上期无",
                attribute,
            );
            finding.current_value = raw_value_json(&current_record.normalized_value);
            finding.band.clear();
            finding.retrieval_note = join_notes(&[
                retrieval_note(Some(current_record), "本期"),
                retrieval_note(None, "上期"),
            ]);
            findings.push(finding);
            continue;
        };

        if current_record.form_code != previous_record.form_code {
            let mut finding = base_finding(
                current_record,
                STATUS_ABNORMAL,
                "硬性",
                &format!(
                    "两期表单不一致：本期 {}，上期 {}",
                    current_record.form_code, previous_record.form_code
                ),
                attribute,
            );
            finding.current_value = raw_value_json(&current_record.normalized_value);
            finding.previous_value = raw_value_json(&previous_record.normalized_value);
            finding.band.clear();
            finding.retrieval_note = join_notes(&[
                retrieval_note(Some(current_record), "本期"),
                retrieval_note(Some(previous_record), "上期"),
            ]);
            finding.source_refs = vec![
                current_record.source.clone(),
                previous_record.source.clone(),
            ];
            findings.push(finding);
            continue;
        }

        if current_record.value_type == "文字" || previous_record.value_type == "文字" {
            let changed = text_value(&current_record.normalized_value)
                != text_value(&previous_record.normalized_value);
            let mut finding = base_finding(
                current_record,
                if changed {
                    STATUS_ABNORMAL
                } else {
                    STATUS_NORMAL
                },
                if changed { text_severity } else { "" },
                if changed {
                    "文字指标变化"
                } else {
                    "无变动"
                },
                attribute,
            );
            finding.current_value = raw_value_json(&current_record.normalized_value);
            finding.previous_value = raw_value_json(&previous_record.normalized_value);
            finding.band.clear();
            finding.retrieval_note = join_notes(&[
                retrieval_note(Some(current_record), "本期"),
                retrieval_note(Some(previous_record), "上期"),
                "变动幅度未计算".to_owned(),
            ]);
            finding.source_refs = vec![
                current_record.source.clone(),
                previous_record.source.clone(),
            ];
            findings.push(finding);
            continue;
        }

        let current_number = Numeric::from_imported(&current_record.normalized_value);
        let previous_number = Numeric::from_imported(&previous_record.normalized_value);
        if current_number.is_none() && previous_number.is_none() {
            let mut finding = base_finding(current_record, STATUS_NORMAL, "", "无变动", attribute);
            finding.current_value = raw_value_json(&current_record.normalized_value);
            finding.previous_value = raw_value_json(&previous_record.normalized_value);
            finding.band.clear();
            finding.retrieval_note = join_notes(&[
                retrieval_note(Some(current_record), "本期"),
                retrieval_note(Some(previous_record), "上期"),
            ]);
            finding.source_refs = vec![
                current_record.source.clone(),
                previous_record.source.clone(),
            ];
            findings.push(finding);
            continue;
        }
        if let (Some(current_number), Some(previous_number)) =
            (current_number.as_ref(), previous_number.as_ref())
        {
            if current_number.compare(previous_number) == Some(Ordering::Equal) {
                let mut finding =
                    base_finding(current_record, STATUS_NORMAL, "", "无变动", attribute);
                finding.current_value = current_number.to_json();
                finding.previous_value = previous_number.to_json();
                finding.difference_value = Value::from(0.0);
                finding.change_rate = Value::from(0.0);
                finding.band.clear();
                finding.source_refs = vec![
                    current_record.source.clone(),
                    previous_record.source.clone(),
                ];
                findings.push(finding);
                continue;
            }
        }
        if current_number.is_none() {
            let mut finding = base_finding(
                current_record,
                STATUS_ABNORMAL,
                current_side_severity,
                "本期无，上期有",
                attribute,
            );
            finding.previous_value = previous_number
                .as_ref()
                .map(Numeric::to_json)
                .unwrap_or_else(|| raw_value_json(&previous_record.normalized_value));
            finding.band.clear();
            finding.retrieval_note = join_notes(&[
                retrieval_note(Some(current_record), "本期"),
                retrieval_note(Some(previous_record), "上期"),
            ]);
            finding.source_refs = vec![
                current_record.source.clone(),
                previous_record.source.clone(),
            ];
            findings.push(finding);
            continue;
        }
        if previous_number.is_none() {
            let mut finding = base_finding(
                current_record,
                STATUS_ABNORMAL,
                previous_side_severity,
                "本期有，上期无",
                attribute,
            );
            finding.current_value = current_number
                .as_ref()
                .map(Numeric::to_json)
                .unwrap_or_else(|| raw_value_json(&current_record.normalized_value));
            finding.band.clear();
            finding.retrieval_note = join_notes(&[
                retrieval_note(Some(current_record), "本期"),
                retrieval_note(Some(previous_record), "上期"),
            ]);
            finding.source_refs = vec![
                current_record.source.clone(),
                previous_record.source.clone(),
            ];
            findings.push(finding);
            continue;
        }
        let current_number = current_number.expect("current numeric value");
        let previous_number = previous_number.expect("previous numeric value");
        let configured_type = indicator.map_or("", |item| item.data_type.as_str());
        let percentage = configured_type == "百分数"
            || current_record.value_type == "百分数"
            || previous_record.value_type == "百分数";

        if percentage {
            let difference = current_number.subtract(&previous_number)?;
            let legacy_decimal = difference.divide_hundred()?;
            let (group, band) =
                match_strategy(config, current_record, &difference, Some(&legacy_decimal))?;
            let minimum = minimum_change(config, current_record)?;
            let gate_passed = gate_passed(&difference.absolute(), minimum.as_ref());
            let abnormal = band.as_ref().is_some_and(|band| {
                (!band.description.is_empty() || !band.severity.is_empty()) && gate_passed
            });
            let severity = if abnormal {
                band.as_ref()
                    .map(|band| {
                        if band.severity.is_empty() {
                            "软性"
                        } else {
                            &band.severity
                        }
                    })
                    .unwrap_or("")
            } else {
                ""
            };
            let trace = format!(
                "{}% - {}% = {} 个百分点",
                current_number.trace_number(),
                previous_number.trace_number(),
                difference.trace_number()
            );
            let description = strategy_description(band, abnormal);
            let mut finding = base_finding(
                current_record,
                if abnormal {
                    STATUS_ABNORMAL
                } else {
                    STATUS_NORMAL
                },
                severity,
                &description,
                attribute,
            );
            finding.current_value = current_number.to_json();
            finding.previous_value = previous_number.to_json();
            finding.difference_value = difference.to_json();
            finding.change_rate = difference.to_json();
            finding.value_type = "百分数".to_owned();
            finding.band = strategy_identity(&group, band);
            finding.calculation_trace = trace;
            finding.source_refs = vec![
                current_record.source.clone(),
                previous_record.source.clone(),
            ];
            findings.push(finding);
            continue;
        }

        if previous_number.is_zero() {
            let difference = current_number.clone();
            let mut finding = base_finding(
                current_record,
                STATUS_ABNORMAL,
                previous_side_severity,
                "本期有，上期无",
                attribute,
            );
            finding.current_value = current_number.to_json();
            finding.previous_value = previous_number.to_json();
            finding.difference_value = difference.to_json();
            finding.band = "未命中".to_owned();
            finding.retrieval_note = "变动幅度未计算".to_owned();
            finding.source_refs = vec![
                current_record.source.clone(),
                previous_record.source.clone(),
            ];
            findings.push(finding);
            continue;
        }
        if current_number.is_zero() {
            let difference = current_number.subtract(&previous_number)?;
            let mut finding = base_finding(
                current_record,
                STATUS_ABNORMAL,
                current_side_severity,
                "本期无，上期有",
                attribute,
            );
            finding.current_value = current_number.to_json();
            finding.previous_value = previous_number.to_json();
            finding.difference_value = difference.to_json();
            finding.band = "未命中".to_owned();
            finding.retrieval_note = "变动幅度未计算".to_owned();
            finding.source_refs = vec![
                current_record.source.clone(),
                previous_record.source.clone(),
            ];
            findings.push(finding);
            continue;
        }

        let difference = current_number.subtract(&previous_number)?;
        let change_rate = difference.divide(&previous_number.absolute())?;
        let percent_value = change_rate.multiply_hundred()?;
        let (group, band) =
            match_strategy(config, current_record, &percent_value, Some(&change_rate))?;
        let minimum = minimum_change(config, current_record)?;
        let threshold_passed = gate_passed(&difference.absolute(), minimum.as_ref());
        let abnormal = band.as_ref().is_some_and(|band| {
            (!band.description.is_empty() || !band.severity.is_empty()) && threshold_passed
        });
        let severity = if abnormal {
            band.as_ref()
                .map(|band| {
                    if band.severity.is_empty() {
                        "软性"
                    } else {
                        &band.severity
                    }
                })
                .unwrap_or("")
        } else {
            ""
        };
        let trace = format!(
            "({} - {}) / abs({}) = {}",
            current_number.trace_number(),
            previous_number.trace_number(),
            previous_number.trace_number(),
            change_rate.trace_percent()
        );
        let description = strategy_description(band, abnormal);
        let mut finding = base_finding(
            current_record,
            if abnormal {
                STATUS_ABNORMAL
            } else {
                STATUS_NORMAL
            },
            severity,
            &description,
            attribute,
        );
        finding.current_value = current_number.to_json();
        finding.previous_value = previous_number.to_json();
        finding.difference_value = difference.to_json();
        finding.change_rate = change_rate.to_json();
        finding.value_type = if current_record.value_type.is_empty() {
            "数值".to_owned()
        } else {
            current_record.value_type.clone()
        };
        finding.band = strategy_identity(&group, band);
        finding.calculation_trace = trace;
        finding.source_refs = vec![
            current_record.source.clone(),
            previous_record.source.clone(),
        ];
        findings.push(finding);
    }
    Ok(findings)
}

fn comparison_index(records: &[IndicatorRecord]) -> BTreeMap<(String, String), &IndicatorRecord> {
    let mut result = BTreeMap::new();
    for record in records {
        result.insert(
            (record.institution_id.clone(), record.indicator_code.clone()),
            record,
        );
    }
    result
}

fn base_finding(
    record: &IndicatorRecord,
    status: &str,
    severity: &str,
    description: &str,
    data_type: &str,
) -> S2AuditFinding {
    let mut finding = S2AuditFinding {
        finding_id: finding_id(&[
            AUDIT_TYPE,
            &record.institution_id,
            &record.indicator_code,
            &record.period,
        ]),
        audit_type: AUDIT_TYPE.to_owned(),
        status: status.to_owned(),
        severity: severity.to_owned(),
        region: record.region.clone(),
        handling_branch: record.handling_branch.clone(),
        institution_type: record.institution_type.clone(),
        institution_name: record.institution_name.clone(),
        social_credit_code: record.social_credit_code.clone(),
        form_code: record.form_code.clone(),
        form_codes: if record.form_code.is_empty() {
            Vec::new()
        } else {
            vec![record.form_code.clone()]
        },
        period: record.period.clone(),
        indicator_code: record.indicator_code.clone(),
        indicator_name: record.indicator_name.clone(),
        rule_id: String::new(),
        rule_description: String::new(),
        external_indicator_name: String::new(),
        current_value: Value::String(String::new()),
        previous_value: Value::String(String::new()),
        external_value: Value::String(String::new()),
        difference_value: Value::String(String::new()),
        change_rate: Value::Null,
        value_type: String::new(),
        data_type: data_type.to_owned(),
        band: String::new(),
        description: description.to_owned(),
        calculation_trace: String::new(),
        source_refs: vec![record.source.clone()],
        related_records: Vec::new(),
        retrieval_note: String::new(),
        rule_type: String::new(),
        value_details: String::new(),
        report_item: String::new(),
    };
    if status == STATUS_NORMAL {
        finding.severity.clear();
    }
    finding
}

fn finding_id(parts: &[&str]) -> String {
    let digest = Sha1::digest(parts.join("|").as_bytes());
    digest[..8]
        .iter()
        .map(|byte| format!("{byte:02x}"))
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

fn python_truthy(value: &ImportedValue) -> bool {
    match value {
        ImportedValue::Empty => false,
        ImportedValue::Text(value) => !value.is_empty(),
        ImportedValue::Number(value) | ImportedValue::DateTimeSerial(value) => *value != 0.0,
        ImportedValue::Integer(value) => *value != 0,
        ImportedValue::Boolean(value) => *value,
        ImportedValue::Decimal(value) => value.digits.iter().any(|digit| *digit != 0),
        ImportedValue::DateTimeIso(value)
        | ImportedValue::DurationIso(value)
        | ImportedValue::Error(value) => !value.is_empty(),
    }
}

fn text_value(value: &ImportedValue) -> String {
    if !python_truthy(value) {
        return String::new();
    }
    match value {
        ImportedValue::Empty => String::new(),
        ImportedValue::Text(value) => value.clone(),
        ImportedValue::Number(value) | ImportedValue::DateTimeSerial(value) => value.to_string(),
        ImportedValue::Integer(value) => value.to_string(),
        ImportedValue::Boolean(value) => if *value { "True" } else { "False" }.to_owned(),
        ImportedValue::Decimal(value) => decimal_to_string(value),
        ImportedValue::DateTimeIso(value)
        | ImportedValue::DurationIso(value)
        | ImportedValue::Error(value) => value.clone(),
    }
}

fn retrieval_note(record: Option<&IndicatorRecord>, label: &str) -> String {
    let Some(record) = record else {
        return format!("{label}缺少指标");
    };
    if matches!(&record.normalized_value, ImportedValue::Empty)
        || matches!(&record.normalized_value, ImportedValue::Text(value) if value.is_empty())
    {
        return format!("{label}无数据");
    }
    if record.value_type != "文字" && Numeric::from_imported(&record.normalized_value).is_none() {
        return format!("{label}取值非数值");
    }
    String::new()
}

fn join_notes(notes: &[String]) -> String {
    let mut seen = HashSet::new();
    notes
        .iter()
        .filter(|note| !note.is_empty() && seen.insert(note.as_str()))
        .cloned()
        .collect::<Vec<_>>()
        .join("；")
}

fn strategy_bands<'a>(
    config: &'a S2PeriodConfig,
    record: &IndicatorRecord,
) -> (String, Vec<&'a S2PeriodBandConfig>) {
    let indicator = config.indicators.get(&record.indicator_code);
    let group = indicator
        .map(|item| item.period_strategy_group.trim())
        .filter(|group| !group.is_empty())
        .unwrap_or("");
    if group.is_empty() || group == "不适用" {
        let candidates = config
            .bands
            .iter()
            .filter(|band| !band.disabled && band.strategy_group.is_empty())
            .collect();
        return (group.to_owned(), candidates);
    }
    let data_type = indicator
        .map(|item| item.data_type.as_str())
        .filter(|value| !value.is_empty())
        .unwrap_or(&record.value_type);
    let candidates = config
        .bands
        .iter()
        .filter(|band| {
            !band.disabled
                && band.strategy_group == group
                && (band.data_type.is_empty() || band.data_type == data_type)
        })
        .collect();
    (group.to_owned(), candidates)
}

fn match_strategy<'a>(
    config: &'a S2PeriodConfig,
    record: &IndicatorRecord,
    percent_value: &Numeric,
    legacy_decimal_value: Option<&Numeric>,
) -> Result<(String, Option<&'a S2PeriodBandConfig>)> {
    let (group, candidates) = strategy_bands(config, record);
    let value = if group.is_empty() {
        legacy_decimal_value.unwrap_or(percent_value)
    } else {
        percent_value
    };
    let matches = candidates
        .into_iter()
        .filter(|band| band_includes(band, value))
        .collect::<Vec<_>>();
    if matches.len() > 1 {
        let ids = matches
            .iter()
            .map(|band| band.rule_id.as_str())
            .collect::<Vec<_>>()
            .join("、");
        bail!(
            "指标 {} 环比策略组 {} 在变动值 {} 上匹配多个规则：{}",
            record.indicator_code,
            if group.is_empty() {
                "旧配置"
            } else {
                &group
            },
            value.python_repr(),
            ids
        );
    }
    if !group.is_empty() && matches.is_empty() {
        bail!(
            "指标 {} 环比策略组 {} 在变动值 {} 上没有唯一匹配规则",
            record.indicator_code,
            group,
            value.python_repr()
        );
    }
    Ok((group, matches.first().copied()))
}

fn band_includes(band: &S2PeriodBandConfig, value: &Numeric) -> bool {
    let lower = band
        .lower
        .map(|bound| value.compare(&Numeric::Float(bound)) != Some(Ordering::Less))
        .unwrap_or(true);
    let upper = band
        .upper
        .map(|bound| value.compare(&Numeric::Float(bound)) == Some(Ordering::Less))
        .unwrap_or(true);
    lower && upper
}

fn minimum_change(config: &S2PeriodConfig, record: &IndicatorRecord) -> Result<Option<Numeric>> {
    let Some(indicator) = config.indicators.get(&record.indicator_code) else {
        return Ok(None);
    };
    if indicator.data_type == "文字" {
        return Ok(None);
    }
    let Some(value) = indicator.min_change_value else {
        return Ok(None);
    };
    if AMOUNT_DATA_TYPES.contains(&indicator.data_type.as_str()) {
        let decimal = as_decimal(&value.to_string()).map_err(decimal_error)?;
        Ok(Some(Numeric::Decimal(decimal)))
    } else {
        Ok(Some(Numeric::Float(value)))
    }
}

fn decimal_error(error: UnitConversionError) -> anyhow::Error {
    anyhow::anyhow!(error.to_string())
}

fn gate_passed(abs_change: &Numeric, minimum: Option<&Numeric>) -> bool {
    let Some(minimum) = minimum else {
        return true;
    };
    if minimum.is_zero() {
        return true;
    }
    matches!(
        abs_change.compare(minimum),
        Some(Ordering::Equal | Ordering::Greater)
    )
}

fn strategy_identity(group: &str, band: Option<&S2PeriodBandConfig>) -> String {
    let Some(band) = band else {
        return String::new();
    };
    if group.is_empty() {
        band.rule_id.clone()
    } else {
        format!("{group}-{}", band.rule_id)
    }
}

fn strategy_description(band: Option<&S2PeriodBandConfig>, abnormal: bool) -> String {
    if abnormal {
        band.map(|band| band.description.clone())
            .unwrap_or_default()
    } else {
        String::new()
    }
}

pub(crate) fn decimal_to_string(value: &DecimalTuple) -> String {
    let mut digits = digits_string(&value.digits);
    let exponent = value.exponent;
    if exponent >= 0 {
        digits.push_str(&"0".repeat(exponent.min(100_000) as usize));
    } else {
        let point = digits.len() as i64 + exponent;
        if point > 0 {
            digits.insert(point as usize, '.');
        } else {
            digits = format!("0.{}{}", "0".repeat((-point).min(100_000) as usize), digits);
        }
    }
    if value.sign == 1 {
        format!("-{digits}")
    } else {
        digits
    }
}

pub(crate) fn decimal_fixed(value: &DecimalTuple, places: usize) -> String {
    let digits = normalized_digits(&value.digits);
    let target_exponent = -(places as i64);
    let shift = value.exponent.saturating_sub(target_exponent);
    let mut cents = if shift >= 0 {
        let mut result = digits;
        result.extend(std::iter::repeat_n(0, shift.min(100_000) as usize));
        result
    } else {
        let drop = (-shift).min(1_000_000) as usize;
        let keep = digits.len().saturating_sub(drop);
        let mut result = if keep == 0 {
            vec![0]
        } else {
            digits[..keep].to_vec()
        };
        let first_discard = if drop > digits.len() { 0 } else { digits[keep] };
        let rest_nonzero = drop <= digits.len()
            && digits[keep.saturating_add(1)..]
                .iter()
                .any(|digit| *digit != 0);
        let last_kept = *result.last().unwrap_or(&0);
        let round_up =
            first_discard > 5 || (first_discard == 5 && (rest_nonzero || last_kept % 2 == 1));
        if round_up {
            increment_digits(&mut result);
        }
        result
    };
    trim_leading_zeroes(&mut cents);
    let mut rendered = digits_string(&cents);
    if places > 0 {
        if rendered.len() <= places {
            rendered = format!("{}{}", "0".repeat(places + 1 - rendered.len()), rendered);
        }
        rendered.insert(rendered.len() - places, '.');
    }
    if value.sign == 1 {
        format!("-{rendered}")
    } else {
        rendered
    }
}

pub(crate) fn compare_decimal(left: &DecimalTuple, right: &DecimalTuple) -> Ordering {
    let left_digits = normalized_digits(&left.digits);
    let right_digits = normalized_digits(&right.digits);
    let left_zero = left_digits.iter().all(|digit| *digit == 0);
    let right_zero = right_digits.iter().all(|digit| *digit == 0);
    match (left_zero, right_zero) {
        (true, true) => return Ordering::Equal,
        (true, false) => {
            return if right.sign == 1 {
                Ordering::Greater
            } else {
                Ordering::Less
            }
        }
        (false, true) => {
            return if left.sign == 1 {
                Ordering::Less
            } else {
                Ordering::Greater
            }
        }
        (false, false) => {}
    }
    if left.sign != right.sign {
        return if left.sign == 1 {
            Ordering::Less
        } else {
            Ordering::Greater
        };
    }
    let absolute = compare_decimal_abs(left, right, &left_digits, &right_digits);
    if left.sign == 1 {
        absolute.reverse()
    } else {
        absolute
    }
}

fn compare_decimal_abs(
    left: &DecimalTuple,
    right: &DecimalTuple,
    left_digits: &[u8],
    right_digits: &[u8],
) -> Ordering {
    let left_magnitude = left_digits.len() as i128 + i128::from(left.exponent);
    let right_magnitude = right_digits.len() as i128 + i128::from(right.exponent);
    match left_magnitude.cmp(&right_magnitude) {
        Ordering::Equal => {
            let count = left_digits.len().max(right_digits.len());
            for index in 0..count {
                let left_digit = left_digits.get(index).copied().unwrap_or(0);
                let right_digit = right_digits.get(index).copied().unwrap_or(0);
                match left_digit.cmp(&right_digit) {
                    Ordering::Equal => {}
                    ordering => return ordering,
                }
            }
            Ordering::Equal
        }
        ordering => ordering,
    }
}

pub(crate) fn decimal_subtract(left: &DecimalTuple, right: &DecimalTuple) -> Result<DecimalTuple> {
    let exponent = left.exponent.min(right.exponent);
    let left_shift = usize::try_from(left.exponent - exponent)
        .map_err(|_| anyhow::anyhow!("Decimal 指数差超出可处理范围"))?;
    let right_shift = usize::try_from(right.exponent - exponent)
        .map_err(|_| anyhow::anyhow!("Decimal 指数差超出可处理范围"))?;
    if left_shift > 100_000 || right_shift > 100_000 {
        bail!("Decimal 指数差超出可处理范围");
    }
    let mut left_digits = normalized_digits(&left.digits);
    let mut right_digits = normalized_digits(&right.digits);
    left_digits.extend(std::iter::repeat_n(0, left_shift));
    right_digits.extend(std::iter::repeat_n(0, right_shift));
    let right_sign = if right_digits.iter().all(|digit| *digit == 0) {
        right.sign
    } else {
        1 - right.sign
    };
    let (sign, digits) = if left.sign == right_sign {
        (left.sign, add_abs(&left_digits, &right_digits))
    } else {
        match compare_abs(&left_digits, &right_digits) {
            Ordering::Equal => (0, vec![0]),
            Ordering::Greater => (left.sign, subtract_abs(&left_digits, &right_digits)),
            Ordering::Less => (right_sign, subtract_abs(&right_digits, &left_digits)),
        }
    };
    Ok(round_significant(DecimalTuple {
        sign,
        digits,
        exponent,
    }))
}

fn decimal_divide(left: &DecimalTuple, right: &DecimalTuple) -> Result<DecimalTuple> {
    let right_digits = normalized_digits(&right.digits);
    if right_digits.iter().all(|digit| *digit == 0) {
        bail!("Decimal division by zero");
    }
    let left_digits = normalized_digits(&left.digits);
    if left_digits.iter().all(|digit| *digit == 0) {
        return Ok(DecimalTuple {
            sign: left.sign ^ right.sign,
            digits: vec![0],
            exponent: 0,
        });
    }
    let left_significand = padded_digits(&left_digits, right_digits.len());
    let right_significand = padded_digits(&right_digits, left_digits.len());
    let coefficient_exponent = left_digits.len() as i64
        - right_digits.len() as i64
        - i64::from(left_significand < right_significand);
    let exponent_difference = left
        .exponent
        .checked_sub(right.exponent)
        .ok_or_else(|| anyhow::anyhow!("Decimal 指数超出可处理范围"))?;
    let result_order = coefficient_exponent
        .checked_add(exponent_difference)
        .ok_or_else(|| anyhow::anyhow!("Decimal 指数超出可处理范围"))?;
    let numerator_shift = DECIMAL_CONTEXT_PRECISION as i64 - 1 - coefficient_exponent;
    let mut numerator = left_digits;
    let mut denominator = right_digits;
    if numerator_shift >= 0 {
        if numerator_shift > 100_000 {
            bail!("Decimal division precision shift exceeds supported range");
        }
        numerator.extend(std::iter::repeat_n(0, numerator_shift as usize));
    } else {
        if numerator_shift < -100_000 {
            bail!("Decimal division precision shift exceeds supported range");
        }
        denominator.extend(std::iter::repeat_n(0, (-numerator_shift) as usize));
    }
    let (mut quotient, remainder) = divide_abs(&numerator, &denominator);
    let twice_remainder = multiply_small(&remainder, 2);
    let rounding = compare_abs(&twice_remainder, &denominator);
    if rounding == Ordering::Greater
        || (rounding == Ordering::Equal && quotient.last().copied().unwrap_or(0) % 2 == 1)
    {
        increment_digits(&mut quotient);
    }
    if quotient.len() > DECIMAL_CONTEXT_PRECISION {
        quotient.pop();
        let exponent = result_order
            .checked_sub(DECIMAL_CONTEXT_PRECISION as i64 - 2)
            .ok_or_else(|| anyhow::anyhow!("Decimal 指数超出可处理范围"))?;
        let (digits, exponent) = strip_division_trailing_zeros(quotient, exponent)?;
        return Ok(DecimalTuple {
            sign: left.sign ^ right.sign,
            digits,
            exponent,
        });
    }
    let exponent = result_order - (DECIMAL_CONTEXT_PRECISION as i64 - 1);
    let (digits, exponent) = strip_division_trailing_zeros(quotient, exponent)?;
    Ok(DecimalTuple {
        sign: left.sign ^ right.sign,
        digits,
        exponent,
    })
}

fn strip_division_trailing_zeros(mut digits: Vec<u8>, mut exponent: i64) -> Result<(Vec<u8>, i64)> {
    while digits.len() > 1 && digits.last() == Some(&0) {
        digits.pop();
        exponent = exponent
            .checked_add(1)
            .ok_or_else(|| anyhow::anyhow!("Decimal 指数超出可处理范围"))?;
    }
    Ok((digits, exponent))
}

fn round_significant(mut value: DecimalTuple) -> DecimalTuple {
    let mut digits = normalized_digits(&value.digits);
    if digits.len() <= DECIMAL_CONTEXT_PRECISION {
        value.digits = digits;
        return value;
    }
    let drop = digits.len() - DECIMAL_CONTEXT_PRECISION;
    let first = digits[DECIMAL_CONTEXT_PRECISION];
    let rest_nonzero = digits[DECIMAL_CONTEXT_PRECISION + 1..]
        .iter()
        .any(|digit| *digit != 0);
    let last = digits[DECIMAL_CONTEXT_PRECISION - 1];
    let round_up = first > 5 || (first == 5 && (rest_nonzero || last % 2 == 1));
    digits.truncate(DECIMAL_CONTEXT_PRECISION);
    if round_up {
        increment_digits(&mut digits);
    }
    if digits.len() > DECIMAL_CONTEXT_PRECISION {
        digits.pop();
        digits.fill(0);
        digits[0] = 1;
        value.exponent += drop as i64 + 1;
    } else {
        value.exponent += drop as i64;
    }
    value.digits = digits;
    value
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
        .map(|b| b - b'0')
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
        let scale = -exponent_two;
        for _ in 0..scale {
            digits = multiply_small(&digits, 5);
        }
        Some(DecimalTuple {
            sign,
            digits,
            exponent: i64::from(exponent_two),
        })
    }
}

fn normalized_digits(digits: &[u8]) -> Vec<u8> {
    let first = digits
        .iter()
        .position(|digit| *digit != 0)
        .unwrap_or(digits.len().saturating_sub(1));
    digits[first..].to_vec()
}

fn trim_leading_zeroes(digits: &mut Vec<u8>) {
    let first = digits
        .iter()
        .position(|digit| *digit != 0)
        .unwrap_or(digits.len().saturating_sub(1));
    if first > 0 {
        digits.drain(..first);
    }
    if digits.is_empty() {
        digits.push(0);
    }
}

fn digits_string(digits: &[u8]) -> String {
    digits
        .iter()
        .map(|digit| char::from(b'0' + *digit))
        .collect()
}

fn padded_digits(digits: &[u8], length: usize) -> Vec<u8> {
    let mut result = digits.to_vec();
    result.resize(length, 0);
    result
}

fn compare_abs(left: &[u8], right: &[u8]) -> Ordering {
    let left = normalized_digits(left);
    let right = normalized_digits(right);
    left.len().cmp(&right.len()).then_with(|| left.cmp(&right))
}

fn add_abs(left: &[u8], right: &[u8]) -> Vec<u8> {
    let mut output = Vec::with_capacity(left.len().max(right.len()) + 1);
    let (mut i, mut j, mut carry) = (left.len(), right.len(), 0_u8);
    while i > 0 || j > 0 || carry > 0 {
        let a = if i > 0 {
            i -= 1;
            left[i]
        } else {
            0
        };
        let b = if j > 0 {
            j -= 1;
            right[j]
        } else {
            0
        };
        let sum = a + b + carry;
        output.push(sum % 10);
        carry = sum / 10;
    }
    output.reverse();
    output
}

fn subtract_abs(left: &[u8], right: &[u8]) -> Vec<u8> {
    let mut output = Vec::with_capacity(left.len());
    let (mut i, mut j, mut borrow) = (left.len(), right.len(), 0_i16);
    while i > 0 {
        i -= 1;
        let a = i16::from(left[i]) - borrow;
        let b = if j > 0 {
            j -= 1;
            i16::from(right[j])
        } else {
            0
        };
        if a < b {
            output.push((a + 10 - b) as u8);
            borrow = 1;
        } else {
            output.push((a - b) as u8);
            borrow = 0;
        }
    }
    output.reverse();
    trim_leading_zeroes(&mut output);
    output
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
    trim_leading_zeroes(&mut output);
    output
}

fn increment_digits(digits: &mut Vec<u8>) {
    let mut index = digits.len();
    while index > 0 {
        index -= 1;
        if digits[index] < 9 {
            digits[index] += 1;
            return;
        }
        digits[index] = 0;
    }
    digits.insert(0, 1);
}

fn divide_abs(numerator: &[u8], denominator: &[u8]) -> (Vec<u8>, Vec<u8>) {
    let mut quotient = Vec::with_capacity(numerator.len());
    let mut remainder = vec![0];
    for digit in numerator {
        if remainder.len() == 1 && remainder[0] == 0 {
            remainder[0] = *digit;
        } else {
            remainder.push(*digit);
        }
        trim_leading_zeroes(&mut remainder);
        let mut quotient_digit = 0_u8;
        while compare_abs(&remainder, denominator) != Ordering::Less {
            remainder = subtract_abs(&remainder, denominator);
            quotient_digit += 1;
        }
        quotient.push(quotient_digit);
    }
    trim_leading_zeroes(&mut quotient);
    (quotient, remainder)
}

#[cfg(test)]
mod tests {
    use super::{
        compare_decimal, decimal_divide, decimal_fixed, decimal_subtract, float_to_decimal,
    };
    use crate::unit_conversion::as_decimal;
    use std::cmp::Ordering;

    #[test]
    fn decimal_subtraction_preserves_scale_and_rounds_like_python_context() {
        let current = as_decimal("8685065.63000").unwrap();
        let previous = as_decimal("8663670.02000").unwrap();
        let difference = decimal_subtract(&current, &previous).unwrap();
        assert_eq!(difference, as_decimal("21395.61000").unwrap());
        assert_eq!(decimal_fixed(&difference, 2), "21395.61");

        let wide = as_decimal("99999999999999999999999999999").unwrap();
        let one = as_decimal("1").unwrap();
        assert_eq!(
            decimal_subtract(&wide, &one).unwrap(),
            as_decimal("1.000000000000000000000000000E+29").unwrap()
        );
    }

    #[test]
    fn decimal_division_matches_default_decimal_precision_and_half_even() {
        let one = as_decimal("1").unwrap();
        let three = as_decimal("3").unwrap();
        let third = decimal_divide(&one, &three).unwrap();
        assert_eq!(third, as_decimal("0.3333333333333333333333333333").unwrap());
        assert_eq!(
            decimal_divide(
                &as_decimal("50000").unwrap(),
                &as_decimal("100000").unwrap()
            )
            .unwrap(),
            as_decimal("0.5").unwrap()
        );
        assert_eq!(decimal_fixed(&as_decimal("0.005").unwrap(), 2), "0.00");
        assert_eq!(decimal_fixed(&as_decimal("0.015").unwrap(), 2), "0.02");
    }

    #[test]
    fn decimal_float_comparison_uses_the_exact_binary_float_value() {
        let decimal = as_decimal("0.3").unwrap();
        let bound = float_to_decimal(0.3).unwrap();
        assert_eq!(compare_decimal(&decimal, &bound), Ordering::Greater);
    }
}
