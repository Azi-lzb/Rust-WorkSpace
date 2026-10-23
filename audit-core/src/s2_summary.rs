//! S2-F01 report summary data derived from imported current-period records and findings.
//!
//! This mirrors the source exporter's `build_summary`; it does not write a
//! workbook or imply that the surrounding S2 workflow has been migrated.

use crate::{s2_import::IndicatorRecord, s2_validation::S2AuditFinding};
use serde::Serialize;
use std::collections::{BTreeMap, HashSet};

const STATUS_ABNORMAL: &str = "异常";

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct S2AuditSummary {
    pub institutions: usize,
    pub forms: usize,
    pub indicators: usize,
    pub total_findings: usize,
    pub abnormal_findings: usize,
    pub by_type: BTreeMap<String, usize>,
    pub by_severity: BTreeMap<String, usize>,
}

/// Build summary counts from the same record set that will be written to the
/// report. Findings with any status other than `异常` contribute only to the
/// total; empty severities are omitted from the severity breakdown.
pub fn build_summary(records: &[IndicatorRecord], findings: &[S2AuditFinding]) -> S2AuditSummary {
    let institutions = records
        .iter()
        .map(|record| record.institution_id.as_str())
        .collect::<HashSet<_>>()
        .len();
    let forms = records
        .iter()
        .map(|record| record.form_code.as_str())
        .filter(|form_code| !form_code.is_empty())
        .collect::<HashSet<_>>()
        .len();

    let mut abnormal_findings = 0;
    let mut by_type = BTreeMap::new();
    let mut by_severity = BTreeMap::new();
    for finding in findings {
        if finding.status != STATUS_ABNORMAL {
            continue;
        }
        abnormal_findings += 1;
        *by_type.entry(finding.audit_type.clone()).or_insert(0) += 1;
        if !finding.severity.is_empty() {
            *by_severity.entry(finding.severity.clone()).or_insert(0) += 1;
        }
    }

    S2AuditSummary {
        institutions,
        forms,
        indicators: records.len(),
        total_findings: findings.len(),
        abnormal_findings,
        by_type,
        by_severity,
    }
}

#[cfg(test)]
mod tests {
    use super::build_summary;
    use crate::{
        s2_import::{ImportedValue, IndicatorRecord, S2SourceRef},
        s2_validation::S2AuditFinding,
    };
    use serde::Deserialize;
    use serde_json::{Map, Value};
    use std::{fs, path::Path};

    #[derive(Deserialize)]
    struct FixtureRoot {
        cases: Vec<FixtureCase>,
    }

    #[derive(Deserialize)]
    struct FixtureCase {
        id: String,
        records: Vec<RecordInput>,
        findings: Vec<FindingInput>,
    }

    #[derive(Deserialize)]
    struct RecordInput {
        institution_id: String,
        form_code: String,
    }

    #[derive(Deserialize)]
    struct FindingInput {
        audit_type: String,
        status: String,
        severity: String,
    }

    fn workspace_path(relative: &str) -> std::path::PathBuf {
        Path::new(env!("CARGO_MANIFEST_DIR")).join(relative)
    }

    fn into_record(input: RecordInput, row: usize) -> IndicatorRecord {
        IndicatorRecord {
            institution_id: input.institution_id,
            institution_name: String::new(),
            social_credit_code: String::new(),
            institution_type: String::new(),
            handling_branch: String::new(),
            region: String::new(),
            form_code: input.form_code.clone(),
            indicator_code: String::new(),
            indicator_name: String::new(),
            period: String::new(),
            source_date: String::new(),
            value: ImportedValue::Empty,
            value_type: String::new(),
            source_unit: String::new(),
            normalized_value: ImportedValue::Empty,
            normalized_unit: String::new(),
            source: S2SourceRef {
                source_file: "s2_summary_fixture.xlsx".to_owned(),
                source_sheet: input.form_code,
                source_row: row,
            },
        }
    }

    fn into_finding(input: FindingInput, index: usize) -> S2AuditFinding {
        let mut finding = S2AuditFinding::empty(format!("summary-fixture-{index}"));
        finding.audit_type = input.audit_type;
        finding.status = input.status;
        finding.severity = input.severity;
        finding
    }

    #[test]
    fn summary_counts_match_read_only_python_exporter_baseline() {
        let fixture: FixtureRoot = serde_json::from_slice(
            &fs::read(workspace_path("tests/fixtures/s2_summary/cases.json"))
                .expect("read S2 summary fixture"),
        )
        .expect("decode S2 summary fixture");
        let baseline: Value = serde_json::from_slice(
            &fs::read(workspace_path(
                "probe/reports/s2_summary_python_baseline.json",
            ))
            .expect("read S2 summary Python baseline"),
        )
        .expect("decode S2 summary Python baseline");

        let mut actual_results = Map::new();
        let mut case_results = Vec::new();
        for case in fixture.cases {
            let records = case
                .records
                .into_iter()
                .enumerate()
                .map(|(index, record)| into_record(record, index + 1))
                .collect::<Vec<_>>();
            let findings = case
                .findings
                .into_iter()
                .enumerate()
                .map(|(index, finding)| into_finding(finding, index + 1))
                .collect::<Vec<_>>();
            let actual = serde_json::to_value(build_summary(&records, &findings))
                .expect("serialize Rust S2 summary");
            assert_eq!(
                actual, baseline["result"][&case.id],
                "summary differs from Python baseline for {}",
                case.id
            );
            actual_results.insert(case.id.clone(), actual);
            case_results.push(serde_json::json!({"case": case.id, "status": "passed"}));
        }

        if let Some(path) = std::env::var_os("AUDIT_S2_SUMMARY_PARITY_REPORT") {
            let report = serde_json::json!({
                "status": "passed",
                "case_count": case_results.len(),
                "python_baseline": baseline["source"],
                "result": Value::Object(actual_results),
                "cases": case_results,
                "scope": "S2 summary aggregation only; does not cover S2 orchestration, workbook export, real templates, or target Linux packaging."
            });
            fs::write(
                path,
                format!("{}\n", serde_json::to_string_pretty(&report).unwrap()),
            )
            .expect("write S2 summary parity report");
        }
    }
}
