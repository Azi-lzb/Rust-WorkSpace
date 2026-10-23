//! In-memory S2 analysis orchestration over already imported inputs.
//!
//! This connects the migrated period, external, rule, quality, and summary
//! components. The caller remains responsible for configuration checking and
//! loading, workbook imports, progress/UI behavior, and report export.

use crate::{
    s2_config::S2AuditConfig,
    s2_external::compare_external,
    s2_external_import::S2ExternalImportResult,
    s2_import::{IndicatorRecord, S2ImportResult},
    s2_period::compare_periods,
    s2_rule_engine::evaluate_rules,
    s2_summary::{build_summary, S2AuditSummary},
    s2_validation::{
        validate_config_warnings, validate_external_warnings, validate_import, S2AuditFinding,
    },
};
use anyhow::Result;
use serde::Serialize;
use std::collections::HashSet;

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct S2AnalysisResult {
    pub current_records: Vec<IndicatorRecord>,
    pub previous_records: Vec<IndicatorRecord>,
    pub period_findings: Vec<S2AuditFinding>,
    pub external_findings: Vec<S2AuditFinding>,
    pub rule_findings: Vec<S2AuditFinding>,
    pub quality_findings: Vec<S2AuditFinding>,
    pub rule_quality_findings: Vec<S2AuditFinding>,
    pub findings: Vec<S2AuditFinding>,
    pub summary: S2AuditSummary,
}

/// Analyze already imported S2 inputs using the loaded business configuration.
///
/// First records for repeated logical keys are retained in input order, matching
/// Python's insertion-ordered `ImportResult.by_logical_key()`. Findings are
/// combined in the source service order: period, external, rules, import/config
/// quality, then rule-engine quality. No filesystem or workbook I/O occurs.
pub fn analyze_imports(
    current_import: &S2ImportResult,
    previous_import: &S2ImportResult,
    config: &S2AuditConfig,
    external_import: Option<&S2ExternalImportResult>,
) -> Result<S2AnalysisResult> {
    let current_records = first_records_by_logical_key(current_import);
    let previous_records = first_records_by_logical_key(previous_import);

    let period_findings =
        compare_periods(&current_records, &previous_records, &config.period_config())?;
    let (rule_findings, rule_quality_findings) = evaluate_rules(
        &current_records,
        &previous_records,
        &config.rule_engine_config(),
    )?;

    let mut quality_findings = validate_config_warnings(&config.warnings);
    quality_findings.extend(validate_import(current_import));
    quality_findings.extend(validate_import(previous_import));

    let external_findings = if let Some(external_import) = external_import {
        quality_findings.extend(validate_external_warnings(&external_import.warnings));
        compare_external(
            &current_records,
            &external_import.data,
            &external_import.external_org_map,
            &config.external_config(),
        )?
    } else {
        Vec::new()
    };

    Ok(build_analysis_result(
        current_records,
        previous_records,
        period_findings,
        external_findings,
        rule_findings,
        quality_findings,
        rule_quality_findings,
    ))
}

pub(crate) fn build_analysis_result(
    current_records: Vec<IndicatorRecord>,
    previous_records: Vec<IndicatorRecord>,
    period_findings: Vec<S2AuditFinding>,
    external_findings: Vec<S2AuditFinding>,
    rule_findings: Vec<S2AuditFinding>,
    quality_findings: Vec<S2AuditFinding>,
    rule_quality_findings: Vec<S2AuditFinding>,
) -> S2AnalysisResult {
    let mut findings = Vec::with_capacity(
        period_findings.len()
            + external_findings.len()
            + rule_findings.len()
            + quality_findings.len()
            + rule_quality_findings.len(),
    );
    findings.extend(period_findings.iter().cloned());
    findings.extend(external_findings.iter().cloned());
    findings.extend(rule_findings.iter().cloned());
    findings.extend(quality_findings.iter().cloned());
    findings.extend(rule_quality_findings.iter().cloned());

    let summary = build_summary(&current_records, &findings);
    S2AnalysisResult {
        current_records,
        previous_records,
        period_findings,
        external_findings,
        rule_findings,
        quality_findings,
        rule_quality_findings,
        findings,
        summary,
    }
}

pub(crate) fn first_records_by_logical_key(import: &S2ImportResult) -> Vec<IndicatorRecord> {
    let mut seen = HashSet::new();
    import
        .records
        .iter()
        .filter(|record| seen.insert(record.logical_key()))
        .cloned()
        .collect()
}

#[cfg(test)]
mod tests {
    use super::first_records_by_logical_key;
    use crate::s2_import::{ImportedValue, IndicatorRecord, S2ImportResult, S2SourceRef};

    fn record(code: &str, row: usize) -> IndicatorRecord {
        IndicatorRecord {
            institution_id: "I1".to_owned(),
            institution_name: "机构一".to_owned(),
            social_credit_code: "C1".to_owned(),
            institution_type: String::new(),
            handling_branch: String::new(),
            region: String::new(),
            form_code: "20202".to_owned(),
            indicator_code: code.to_owned(),
            indicator_name: code.to_owned(),
            period: "2026-06".to_owned(),
            source_date: "2026-06-30".to_owned(),
            value: ImportedValue::Integer(row as i64),
            value_type: "数值".to_owned(),
            source_unit: "元".to_owned(),
            normalized_value: ImportedValue::Integer(row as i64),
            normalized_unit: "元".to_owned(),
            source: S2SourceRef {
                source_file: "fixture.xlsx".to_owned(),
                source_sheet: "20202".to_owned(),
                source_row: row,
            },
        }
    }

    #[test]
    fn logical_key_deduplication_keeps_first_record_and_input_order() {
        let import = S2ImportResult {
            records: vec![record("B", 1), record("A", 2), record("B", 3)],
            duplicates: Vec::new(),
            warnings: Vec::new(),
        };

        let actual = first_records_by_logical_key(&import);
        assert_eq!(
            actual
                .iter()
                .map(|record| (record.indicator_code.as_str(), record.source.source_row))
                .collect::<Vec<_>>(),
            vec![("B", 1), ("A", 2)]
        );
    }
}
