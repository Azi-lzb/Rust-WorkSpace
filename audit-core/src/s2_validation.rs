//! S2 import-quality findings migrated from the source validation helpers.

use crate::s2_import::{IndicatorRecord, S2ImportResult, S2SourceRef};
use serde::{Deserialize, Serialize};
use serde_json::Value;
use sha1::{Digest, Sha1};

const STATUS_ABNORMAL: &str = "异常";

/// S2 finding wire model. The full field set preserves the source dataclass
/// shape even though import-quality findings populate only a subset.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(default)]
pub struct S2AuditFinding {
    pub finding_id: String,
    pub audit_type: String,
    pub status: String,
    pub severity: String,
    pub region: String,
    pub handling_branch: String,
    pub institution_type: String,
    pub institution_name: String,
    pub social_credit_code: String,
    pub form_code: String,
    pub form_codes: Vec<String>,
    pub period: String,
    pub indicator_code: String,
    pub indicator_name: String,
    pub rule_id: String,
    pub rule_description: String,
    pub external_indicator_name: String,
    pub current_value: Value,
    pub previous_value: Value,
    pub external_value: Value,
    pub difference_value: Value,
    pub change_rate: Value,
    pub value_type: String,
    pub data_type: String,
    pub band: String,
    pub description: String,
    pub calculation_trace: String,
    pub source_refs: Vec<S2SourceRef>,
    pub related_records: Vec<Value>,
    pub retrieval_note: String,
    pub rule_type: String,
    pub value_details: String,
    pub report_item: String,
}

impl S2AuditFinding {
    pub(crate) fn empty(finding_id: String) -> Self {
        Self {
            finding_id,
            audit_type: "校验规则".to_owned(),
            status: STATUS_ABNORMAL.to_owned(),
            severity: String::new(),
            region: String::new(),
            handling_branch: String::new(),
            institution_type: String::new(),
            institution_name: String::new(),
            social_credit_code: String::new(),
            form_code: String::new(),
            form_codes: Vec::new(),
            period: String::new(),
            indicator_code: String::new(),
            indicator_name: String::new(),
            rule_id: String::new(),
            rule_description: String::new(),
            external_indicator_name: String::new(),
            current_value: Value::String(String::new()),
            previous_value: Value::String(String::new()),
            external_value: Value::String(String::new()),
            difference_value: Value::String(String::new()),
            change_rate: Value::Null,
            value_type: String::new(),
            data_type: String::new(),
            band: String::new(),
            description: String::new(),
            calculation_trace: String::new(),
            source_refs: Vec::new(),
            related_records: Vec::new(),
            retrieval_note: String::new(),
            rule_type: String::new(),
            value_details: String::new(),
            report_item: String::new(),
        }
    }
}

impl Default for S2AuditFinding {
    fn default() -> Self {
        Self::empty(String::new())
    }
}

/// Convert the source importer result into stable data-quality findings.
pub fn validate_import(result: &S2ImportResult) -> Vec<S2AuditFinding> {
    let mut findings = Vec::new();
    for (left, right) in &result.duplicates {
        let suffix = if left.form_code != right.form_code {
            "；同一指标编码跨表单冲突"
        } else {
            ""
        };
        let description = format!("重复指标：同一机构、指标编码、数据期出现多条记录{suffix}");
        let trace = format!(
            "首条：{}/{}!第{}行；重复：{}/{}!第{}行",
            left.source.source_file,
            left.source.source_sheet,
            left.source.source_row,
            right.source.source_file,
            right.source.source_sheet,
            right.source.source_row
        );
        findings.push(quality_finding(right, &description, "硬性", &trace));
    }

    for record in &result.records {
        if record.institution_id.is_empty() {
            findings.push(quality_finding(record, "缺少稳定机构身份", "硬性", ""));
        }
        if record.form_code.is_empty() {
            findings.push(quality_finding(record, "未识别表单代码", "软性", ""));
        }
    }

    for warning in &result.warnings {
        findings.push(system_finding("质量", warning, "", "Q001", "导入校验"));
    }
    findings
}

pub fn validate_config_warnings(warnings: &[String]) -> Vec<S2AuditFinding> {
    warnings
        .iter()
        .map(|warning| system_finding("配置", warning, "配置检查：", "Q001", "配置检查"))
        .collect()
}

pub fn validate_external_warnings(warnings: &[String]) -> Vec<S2AuditFinding> {
    warnings
        .iter()
        .map(|warning| system_finding("外部数据导入", warning, "", "Q002", "外部数据导入"))
        .collect()
}

fn quality_finding(
    record: &IndicatorRecord,
    description: &str,
    severity: &str,
    trace: &str,
) -> S2AuditFinding {
    let id = finding_id(&[
        "质量",
        &record.institution_id,
        &record.indicator_code,
        &record.period,
        description,
    ]);
    let mut finding = S2AuditFinding::empty(id);
    finding.severity = severity.to_owned();
    finding.region = record.region.clone();
    finding.handling_branch = record.handling_branch.clone();
    finding.institution_type = record.institution_type.clone();
    finding.institution_name = record.institution_name.clone();
    finding.social_credit_code = record.social_credit_code.clone();
    finding.form_code = record.form_code.clone();
    if !record.form_code.is_empty() {
        finding.form_codes.push(record.form_code.clone());
    }
    finding.period = record.period.clone();
    finding.indicator_code = record.indicator_code.clone();
    finding.indicator_name = record.indicator_name.clone();
    finding.rule_id = "Q001".to_owned();
    finding.rule_description = description.to_owned();
    finding.description = description.to_owned();
    finding.calculation_trace = trace.to_owned();
    finding.source_refs.push(record.source.clone());
    finding.rule_type = "导入校验".to_owned();
    finding
}

fn system_finding(
    category: &str,
    warning: &str,
    description_prefix: &str,
    rule_id: &str,
    rule_type: &str,
) -> S2AuditFinding {
    let id = finding_id(&[category, warning]);
    let description = format!("{description_prefix}{warning}");
    let mut finding = S2AuditFinding::empty(id);
    finding.severity = "软性".to_owned();
    finding.rule_id = rule_id.to_owned();
    finding.rule_description = description.clone();
    finding.description = description;
    finding.rule_type = rule_type.to_owned();
    finding
}

/// The source uses SHA-1 only as a deterministic finding identifier, not as a
/// security primitive. Keep the first 16 lowercase hexadecimal characters.
fn finding_id(parts: &[&str]) -> String {
    let input = parts.join("|");
    let digest = Sha1::digest(input.as_bytes());
    let mut output = String::with_capacity(16);
    for byte in digest.iter().take(8) {
        use std::fmt::Write as _;
        write!(output, "{byte:02x}").expect("writing to String cannot fail");
    }
    output
}
