use audit_core::conditional_format::{evaluate_xlsx_workbook, EvaluationStatus};
use serde_json::Value;
use std::path::PathBuf;

fn fixture() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("s1_conditional_eval.xlsx")
}

#[test]
fn reads_and_evaluates_synthetic_s1_conditional_rules() {
    let report = evaluate_xlsx_workbook(fixture(), &[]).expect("evaluate fixture");
    let baseline: Value = serde_json::from_str(include_str!(
        "../probe/reports/s1_conditional_python_baseline.json"
    ))
    .expect("valid Python conditional-format baseline");
    let cells = report
        .triggered
        .iter()
        .map(|trigger| trigger.cell.as_str())
        .collect::<Vec<_>>();
    assert_eq!(cells, ["B2", "C5", "C7", "E2"]);
    let expected = baseline["triggered"].as_array().unwrap();
    assert_eq!(report.triggered.len(), expected.len());
    for (trigger, expected) in report.triggered.iter().zip(expected) {
        assert_eq!(trigger.sheet, expected["sheet"]);
        assert_eq!(trigger.cell, expected["cell"]);
        assert_eq!(trigger.applies_to, expected["applies_to"]);
        assert_eq!(trigger.rule_type, expected["rule_type"]);
        assert_eq!(trigger.formula, expected["formula"]);
    }
    let messages = report
        .triggered
        .iter()
        .map(|trigger| trigger.message.as_str())
        .collect::<Vec<_>>();
    assert_eq!(
        messages,
        [
            "条件格式规则：B2>0",
            "条件格式规则：AND(C5>0,C5>1)",
            "条件格式规则：AND(C7>0,C7>1)",
            "条件格式规则：A1+A2+A3=C2",
        ]
    );
    assert_eq!(report.rules_total, 7);
    assert_eq!(report.rules_supported, 6);
    assert_eq!(report.scanned_cells, 12);
    assert_eq!(
        report
            .diagnostics
            .iter()
            .filter(|diagnostic| diagnostic.status == EvaluationStatus::Unsupported)
            .count(),
        3
    );
    assert_eq!(
        report
            .diagnostics
            .iter()
            .filter(|diagnostic| diagnostic.status == EvaluationStatus::Error)
            .count(),
        1
    );
}

#[test]
fn rejects_scan_ranges_for_missing_sheets() {
    let error = evaluate_xlsx_workbook(fixture(), &["Missing!A1:A2".to_owned()])
        .expect_err("missing worksheet must be rejected");
    assert!(error.to_string().contains("不存在的工作表"));
}
