//! S2 external-data comparison rules.
//!
//! Workbook parsing is deliberately outside this module. The caller supplies
//! already-imported indicator records, external values, row-label mappings,
//! and metadata used to explain missing matches.

use crate::{
    s2_import::{ImportedValue, IndicatorRecord, S2SourceRef},
    s2_period::{compare_decimal, decimal_fixed, decimal_subtract, decimal_to_string},
    s2_validation::S2AuditFinding,
    unit_conversion::{
        as_decimal, convert_amount, DecimalTuple, UnitConversionError, AMOUNT_DATA_TYPES,
    },
};
use anyhow::{bail, Result};
use chrono::{DateTime, FixedOffset, NaiveDate, NaiveDateTime};
use serde::{Deserialize, Serialize};
use serde_json::Value;
use sha1::{Digest, Sha1};
use std::{
    cmp::Ordering,
    collections::{HashMap, HashSet},
};

const STATUS_NORMAL: &str = "正常";
const STATUS_ABNORMAL: &str = "异常";
const MISSING_LABEL_FALLBACK_REASON: &str = "机构行标签列未识别，外部数按0参与比较";

#[derive(Clone, Debug, Deserialize, Default)]
pub struct S2ExternalConfig {
    #[serde(default)]
    pub indicators: HashMap<String, S2ExternalIndicatorConfig>,
    #[serde(default)]
    pub institutions_by_id: HashMap<String, S2ExternalInstitutionProfile>,
    #[serde(default)]
    pub institutions_by_name: HashMap<String, S2ExternalInstitutionProfile>,
    #[serde(default)]
    pub external_rules: Vec<S2ExternalRule>,
    #[serde(default)]
    pub unit_settings: S2ExternalUnitSettings,
}

#[derive(Clone, Debug, Deserialize, Default)]
pub struct S2ExternalIndicatorConfig {
    #[serde(default)]
    pub data_type: String,
}

#[derive(Clone, Debug, Deserialize, Default)]
pub struct S2ExternalInstitutionProfile {
    #[serde(default)]
    pub report_item: String,
}

#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(default)]
pub struct S2ExternalRule {
    pub external_system: String,
    pub rule_id: String,
    pub indicator_code: String,
    pub external_indicator_name: String,
    pub comparison_method: String,
    pub tolerance: f64,
    pub severity: String,
    pub enabled: bool,
    pub source: Option<S2SourceRef>,
}

impl Default for S2ExternalRule {
    fn default() -> Self {
        Self {
            external_system: "大集中".to_owned(),
            rule_id: String::new(),
            indicator_code: String::new(),
            external_indicator_name: String::new(),
            comparison_method: "等于".to_owned(),
            tolerance: 0.0,
            severity: "硬性".to_owned(),
            enabled: true,
            source: None,
        }
    }
}

#[derive(Clone, Debug, Deserialize)]
#[serde(default)]
pub struct S2ExternalUnitSettings {
    pub external_file_unit: String,
    pub output_file_unit: String,
}

impl Default for S2ExternalUnitSettings {
    fn default() -> Self {
        Self {
            external_file_unit: "元".to_owned(),
            output_file_unit: "元".to_owned(),
        }
    }
}

#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
pub struct S2ExternalValue {
    #[serde(default)]
    pub data_date: String,
    #[serde(default)]
    pub row_label: String,
    #[serde(default)]
    pub indicator_name: String,
    pub value: ImportedValue,
    pub source: S2SourceRef,
    #[serde(default)]
    pub fallback_reason: String,
}

#[derive(Clone, Debug, Deserialize, Eq, Hash, Ord, PartialEq, PartialOrd, Serialize)]
#[serde(tag = "kind", content = "key")]
pub enum S2ExternalValueKey {
    Plain(String, String),
    TimeSeries(String, String, String),
}

#[derive(Clone, Debug, Deserialize, Default, PartialEq, Serialize)]
pub struct S2ExternalData {
    #[serde(default)]
    pub time_series: bool,
    #[serde(default)]
    pub values: Vec<S2ExternalValue>,
    #[serde(default)]
    pub available_dates: HashSet<String>,
    #[serde(default)]
    pub date_labels: HashSet<(String, String)>,
    #[serde(default)]
    pub row_labels: HashSet<String>,
    #[serde(default)]
    pub indicator_names: HashSet<String>,
    #[serde(default)]
    pub blank_values: HashSet<S2ExternalValueKey>,
    #[serde(default)]
    pub missing_label_fallback: Option<S2ExternalValue>,
}

#[derive(Clone, Debug, Eq, Hash, PartialEq)]
enum ExternalKey {
    Plain(String, String),
    TimeSeries(String, String, String),
}

fn external_key(value: &S2ExternalValue) -> ExternalKey {
    if value.data_date.is_empty() {
        ExternalKey::Plain(value.row_label.clone(), value.indicator_name.clone())
    } else {
        ExternalKey::TimeSeries(
            value.data_date.clone(),
            value.row_label.clone(),
            value.indicator_name.clone(),
        )
    }
}

/// Compare imported report values with the already-parsed external data.
///
/// Records are processed in the same institution/code order as the Python
/// implementation; enabled rules retain their configuration order.
pub fn compare_external(
    current: &[IndicatorRecord],
    external: &S2ExternalData,
    external_org_map: &HashMap<String, String>,
    config: &S2ExternalConfig,
) -> Result<Vec<S2AuditFinding>> {
    let mut rules_by_indicator: HashMap<&str, Vec<&S2ExternalRule>> = HashMap::new();
    for rule in &config.external_rules {
        if rule.enabled {
            rules_by_indicator
                .entry(&rule.indicator_code)
                .or_default()
                .push(rule);
        }
    }

    let external_index = external
        .values
        .iter()
        .map(|value| (external_key(value), value))
        .collect::<HashMap<_, _>>();
    let fallback = if external.time_series {
        None
    } else {
        external
            .missing_label_fallback
            .as_ref()
            .filter(|item| item.fallback_reason == MISSING_LABEL_FALLBACK_REASON)
    };

    let mut records = current.iter().collect::<Vec<_>>();
    records.sort_by(|left, right| {
        (&left.institution_id, &left.indicator_code)
            .cmp(&(&right.institution_id, &right.indicator_code))
    });
    let mut findings = Vec::new();

    for record in records {
        if record.value_type == "文字" {
            continue;
        }
        let Some(rules) = rules_by_indicator.get(record.indicator_code.as_str()) else {
            continue;
        };
        let Some(raw_current) = decimal_candidate(&record.normalized_value)? else {
            continue;
        };

        let indicator = config.indicators.get(&record.indicator_code);
        let is_amount =
            indicator.is_some_and(|item| AMOUNT_DATA_TYPES.contains(&item.data_type.as_str()));
        let profile = config
            .institutions_by_id
            .get(&record.institution_id)
            .or_else(|| config.institutions_by_name.get(&record.institution_name));
        let mapping_key = if record.social_credit_code.is_empty() {
            &record.institution_id
        } else {
            &record.social_credit_code
        };
        let row_label = external_org_map
            .get(mapping_key)
            .filter(|label| !label.is_empty())
            .cloned()
            .or_else(|| profile.map(|item| item.report_item.clone()))
            .unwrap_or_default();
        let data_date = if external.time_series {
            normalize_external_date(&record.source_date)
        } else {
            String::new()
        };

        for rule in rules {
            let (mut missing_description, mut missing_reason) = if row_label.is_empty() {
                (
                    "请完善报表采集系统配置或外部文件中的机构参照映射".to_owned(),
                    "缺少机构映射或报表项目".to_owned(),
                )
            } else {
                (
                    "大集中系统数据未找到该报表项目，请先完善大集中数据的报表定义".to_owned(),
                    format!(
                        "报表项目“{row_label}”、指标“{}”无可读取数值",
                        rule.external_indicator_name
                    ),
                )
            };

            let mut external_item = if external.time_series {
                (!data_date.is_empty() && !row_label.is_empty())
                    .then(|| {
                        external_index.get(&ExternalKey::TimeSeries(
                            data_date.clone(),
                            row_label.clone(),
                            rule.external_indicator_name.clone(),
                        ))
                    })
                    .flatten()
                    .copied()
            } else if !row_label.is_empty() {
                external_index
                    .get(&ExternalKey::Plain(
                        row_label.clone(),
                        rule.external_indicator_name.clone(),
                    ))
                    .copied()
            } else {
                None
            };

            if external.time_series {
                if data_date.is_empty() {
                    missing_description = "本期报送日期无效，无法匹配大集中时序数据".to_owned();
                    missing_reason = format!("本期报送日期“{}”不是完整日期", record.source_date);
                } else if !external.available_dates.contains(&data_date) {
                    missing_description = format!("未找到本期日期（{data_date}）的大集中系统数据");
                    missing_reason =
                        format!("外部时序数据没有数据日期 {data_date}，不借用其他日期");
                } else if row_label.is_empty() {
                    missing_reason = "缺少机构映射或报表项目".to_owned();
                } else if !external
                    .date_labels
                    .contains(&(data_date.clone(), row_label.clone()))
                {
                    missing_reason = format!("数据日期 {data_date} 未找到报表项目“{row_label}”");
                } else if !external
                    .indicator_names
                    .contains(&rule.external_indicator_name)
                {
                    missing_description = format!(
                        "已匹配报表项目“{row_label}”，但大集中数据中没有指标列“{}”，请检查导出字段或外部文件。",
                        rule.external_indicator_name
                    );
                    missing_reason = format!(
                        "数据日期 {data_date}、报表项目“{row_label}”已找到，但缺少指标列“{}”",
                        rule.external_indicator_name
                    );
                }
            } else if !row_label.is_empty() && external_item.is_none() {
                if external.row_labels.contains(&row_label) {
                    if !external
                        .indicator_names
                        .contains(&rule.external_indicator_name)
                    {
                        missing_description = format!(
                            "已匹配报表项目“{row_label}”，但大集中数据中没有指标列“{}”，请检查导出字段或外部文件。",
                            rule.external_indicator_name
                        );
                        missing_reason = format!(
                            "报表项目“{row_label}”已找到，但缺少指标列“{}”",
                            rule.external_indicator_name
                        );
                    } else {
                        missing_description = format!(
                            "已匹配报表项目“{row_label}”，但指标“{}”没有可读取数值，请检查大集中数据内容。",
                            rule.external_indicator_name
                        );
                        missing_reason = format!(
                            "报表项目“{row_label}”、指标“{}”无可读取数值",
                            rule.external_indicator_name
                        );
                    }
                }
            }

            let used_missing_label_fallback = external_item.is_none() && fallback.is_some();
            if used_missing_label_fallback {
                external_item = fallback;
            }
            let current_value = decimal_from_candidate(&raw_current)?;
            let external_value = if let Some(item) = external_item {
                let value = decimal_from_external(&item.value)?;
                if is_amount {
                    Some(
                        convert_amount(
                            &decimal_to_string(&value),
                            &config.unit_settings.external_file_unit,
                            &config.unit_settings.output_file_unit,
                        )
                        .map_err(|error| {
                            anyhow::anyhow!(
                                "外部核对规则 {} 的金额单位换算失败：{}",
                                rule.rule_id,
                                error
                            )
                        })?,
                    )
                } else {
                    Some(value)
                }
            } else {
                None
            };

            let (status, severity, description, difference_value, trace, retrieval_note) =
                if let Some(external_value) = external_value.as_ref() {
                    let difference = absolute(decimal_subtract(&current_value, external_value)?);
                    let matched = match rule.comparison_method.as_str() {
                        "等于" => {
                            let tolerance = as_decimal(&python_float_text(rule.tolerance))
                                .map_err(decimal_error)?;
                            compare_decimal(&difference, &tolerance).is_le()
                        }
                        "不等于" => {
                            compare_decimal(&current_value, external_value) != Ordering::Equal
                        }
                        "大于" => compare_decimal(&current_value, external_value).is_gt(),
                        "大于等于" => compare_decimal(&current_value, external_value).is_ge(),
                        "小于" => compare_decimal(&current_value, external_value).is_lt(),
                        "小于等于" => compare_decimal(&current_value, external_value).is_le(),
                        other => bail!("外部核对规则 {} 的比较方式无效：{other}", rule.rule_id),
                    };
                    let status = if matched {
                        STATUS_NORMAL
                    } else {
                        STATUS_ABNORMAL
                    };
                    let mut description = if rule.comparison_method == "等于" {
                        if matched {
                            "与大集中系统数据一致".to_owned()
                        } else {
                            "与大集中系统数据不一致".to_owned()
                        }
                    } else if matched {
                        "满足外部核对规则".to_owned()
                    } else {
                        "不满足外部核对规则".to_owned()
                    };
                    if used_missing_label_fallback {
                        description = format!("{MISSING_LABEL_FALLBACK_REASON}；{description}");
                    }
                    let unit_text = if is_amount {
                        format!(" {}", config.unit_settings.output_file_unit)
                    } else {
                        String::new()
                    };
                    let trace = format!(
                        "差异绝对值= abs({} - {}) = {}{}",
                        decimal_fixed(&current_value, 2),
                        decimal_fixed(external_value, 2),
                        decimal_fixed(&difference, 2),
                        unit_text
                    );
                    let retrieval_note = if used_missing_label_fallback {
                        MISSING_LABEL_FALLBACK_REASON.to_owned()
                    } else {
                        String::new()
                    };
                    (
                        status,
                        if matched {
                            String::new()
                        } else {
                            rule.severity.clone()
                        },
                        description,
                        decimal_json(&difference),
                        trace,
                        retrieval_note,
                    )
                } else {
                    (
                        STATUS_ABNORMAL,
                        rule.severity.clone(),
                        missing_description,
                        Value::String(String::new()),
                        format!("未比较：{missing_reason}"),
                        "外部未匹配".to_owned(),
                    )
                };

            let mut source_refs = vec![record.source.clone()];
            if let Some(item) = external_item {
                source_refs.push(item.source.clone());
            }
            let finding_id = finding_id(&[
                "外部核对",
                &rule.rule_id,
                &record.institution_id,
                &record.indicator_code,
                &record.period,
            ]);
            findings.push(S2AuditFinding {
                finding_id,
                audit_type: "外部核对".to_owned(),
                status: status.to_owned(),
                severity,
                region: record.region.clone(),
                handling_branch: record.handling_branch.clone(),
                institution_type: record.institution_type.clone(),
                institution_name: record.institution_name.clone(),
                social_credit_code: record.social_credit_code.clone(),
                form_code: record.form_code.clone(),
                form_codes: vec![record.form_code.clone()],
                period: record.period.clone(),
                indicator_code: record.indicator_code.clone(),
                indicator_name: record.indicator_name.clone(),
                rule_id: rule.rule_id.clone(),
                rule_description: description.clone(),
                external_indicator_name: rule.external_indicator_name.clone(),
                current_value: decimal_json(&current_value),
                previous_value: Value::String(String::new()),
                external_value: external_value
                    .as_ref()
                    .map(decimal_json)
                    .unwrap_or(Value::Null),
                difference_value,
                change_rate: Value::Null,
                value_type: String::new(),
                data_type: String::new(),
                band: String::new(),
                description,
                calculation_trace: trace,
                source_refs,
                related_records: Vec::new(),
                retrieval_note,
                rule_type: String::new(),
                value_details: String::new(),
                report_item: row_label.clone(),
            });
        }
    }
    Ok(findings)
}

fn decimal_candidate(value: &ImportedValue) -> Result<Option<ImportedValue>> {
    match value {
        ImportedValue::Decimal(_) | ImportedValue::Number(_) | ImportedValue::Integer(_) => {
            Ok(Some(value.clone()))
        }
        ImportedValue::Boolean(value) => {
            let text = if *value { "True" } else { "False" };
            bail!("金额不是有效数值：{text}")
        }
        ImportedValue::Empty
        | ImportedValue::Text(_)
        | ImportedValue::DateTimeSerial(_)
        | ImportedValue::DateTimeIso(_)
        | ImportedValue::DurationIso(_)
        | ImportedValue::Error(_) => Ok(None),
    }
}

fn decimal_from_candidate(value: &ImportedValue) -> Result<DecimalTuple> {
    match value {
        ImportedValue::Decimal(value) => Ok(value.clone()),
        ImportedValue::Integer(value) => as_decimal(&value.to_string()).map_err(decimal_error),
        ImportedValue::Number(value) if value.is_finite() => {
            as_decimal(&python_float_text(*value)).map_err(decimal_error)
        }
        ImportedValue::Number(value) => Err(anyhow::anyhow!(
            "金额不是有限数值：{}",
            python_float_text(*value)
        )),
        _ => bail!("当前指标值不是数值"),
    }
}

fn decimal_from_external(value: &ImportedValue) -> Result<DecimalTuple> {
    match value {
        ImportedValue::Decimal(value) => Ok(value.clone()),
        ImportedValue::Integer(value) => as_decimal(&value.to_string()).map_err(decimal_error),
        ImportedValue::Number(value) if value.is_finite() => {
            as_decimal(&python_float_text(*value)).map_err(decimal_error)
        }
        ImportedValue::Text(value) => as_decimal(value).map_err(decimal_error),
        ImportedValue::Number(value) => Err(anyhow::anyhow!(
            "金额不是有限数值：{}",
            python_float_text(*value)
        )),
        _ => bail!("外部指标值不是数值"),
    }
}

pub(crate) fn python_float_text(value: f64) -> String {
    if value.is_nan() {
        return "nan".to_owned();
    }
    if value.is_infinite() {
        return if value.is_sign_negative() {
            "-inf".to_owned()
        } else {
            "inf".to_owned()
        };
    }
    if value == 0.0 {
        return if value.is_sign_negative() {
            "-0.0".to_owned()
        } else {
            "0.0".to_owned()
        };
    }
    let absolute = value.abs();
    if !(1e-4..1e16).contains(&absolute) {
        let scientific = format!("{value:?}");
        if let Some((coefficient, exponent)) = scientific.split_once('e') {
            let exponent = exponent.parse::<i32>().unwrap_or(0);
            let sign = if exponent >= 0 { "+" } else { "-" };
            return format!("{coefficient}e{sign}{:02}", exponent.unsigned_abs());
        }
    }
    let text = value.to_string();
    if text.contains(['.', 'e', 'E']) {
        text
    } else {
        format!("{text}.0")
    }
}

fn decimal_json(value: &DecimalTuple) -> Value {
    serde_json::to_value(ImportedValue::Decimal(value.clone())).expect("Decimal tuple serializes")
}

fn absolute(mut value: DecimalTuple) -> DecimalTuple {
    value.sign = 0;
    value
}

fn decimal_error(error: UnitConversionError) -> anyhow::Error {
    anyhow::anyhow!(error.to_string())
}

fn finding_id(parts: &[&str]) -> String {
    let digest = Sha1::digest(parts.join("|").as_bytes());
    digest[..8]
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect()
}

pub fn normalize_external_date(value: &str) -> String {
    let text = value.trim();
    if text.is_empty() {
        return String::new();
    }
    if text.chars().all(|character| character.is_ascii_digit()) && text.len() != 8 {
        return String::new();
    }

    if text.len() == 8 && text.bytes().all(|byte| byte.is_ascii_digit()) {
        if let (Ok(year), Ok(month), Ok(day)) = (
            text[..4].parse::<i32>(),
            text[4..6].parse::<u32>(),
            text[6..8].parse::<u32>(),
        ) {
            if let Some(date) = NaiveDate::from_ymd_opt(year, month, day) {
                return date.format("%Y-%m-%d").to_string();
            }
        }
    }
    for format in ["%Y%m%d", "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y年%m月%d日"] {
        if let Ok(date) = NaiveDate::parse_from_str(text, format) {
            return date.format("%Y-%m-%d").to_string();
        }
    }

    if let Ok(datetime) = DateTime::<FixedOffset>::parse_from_rfc3339(text) {
        return datetime.date_naive().format("%Y-%m-%d").to_string();
    }
    for format in [
        "%Y-%m-%dT%H:%M",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%dT%H:%M:%S%.f",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M:%S%.f",
        "%Y-%m-%dt%H:%M",
        "%Y-%m-%dt%H:%M:%S",
        "%Y-%m-%dt%H:%M:%S%.f",
    ] {
        if let Ok(datetime) = NaiveDateTime::parse_from_str(text, format) {
            return datetime.date().format("%Y-%m-%d").to_string();
        }
    }
    for format in [
        "%Y-%m-%dT%H:%M:%S%:z",
        "%Y-%m-%dT%H:%M:%S%.f%:z",
        "%Y-%m-%d %H:%M:%S%:z",
        "%Y-%m-%d %H:%M:%S%.f%:z",
    ] {
        if let Ok(datetime) = DateTime::<FixedOffset>::parse_from_str(text, format) {
            return datetime.date_naive().format("%Y-%m-%d").to_string();
        }
    }
    String::new()
}

#[cfg(test)]
mod tests {
    use super::{normalize_external_date, python_float_text};

    #[test]
    fn date_normalization_accepts_only_complete_valid_dates() {
        assert_eq!(normalize_external_date("20260630"), "2026-06-30");
        assert_eq!(normalize_external_date("2026/07/31"), "2026-07-31");
        assert_eq!(normalize_external_date("2026年8月9日"), "2026-08-09");
        assert_eq!(normalize_external_date("2026-06-30T12:30:00"), "2026-06-30");
        assert_eq!(normalize_external_date("2026-02-30"), "");
        assert_eq!(normalize_external_date("202611"), "");
        assert_eq!(normalize_external_date("2026/07/31 12:05:00"), "");
        assert_eq!(
            normalize_external_date("2026-07-31 12:05:00+08:00"),
            "2026-07-31"
        );
    }

    #[test]
    fn float_text_uses_python_decimal_string_scale_for_common_values() {
        assert_eq!(python_float_text(1.0), "1.0");
        assert_eq!(python_float_text(0.0001), "0.0001");
        assert_eq!(python_float_text(0.00001), "1e-05");
        assert_eq!(python_float_text(1e16), "1e+16");
        assert_eq!(python_float_text(-0.0), "-0.0");
        assert_eq!(python_float_text(f64::NAN), "nan");
        assert_eq!(python_float_text(f64::INFINITY), "inf");
        assert_eq!(python_float_text(f64::NEG_INFINITY), "-inf");
    }
}
