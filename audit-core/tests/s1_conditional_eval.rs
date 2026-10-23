use audit_core::conditional_format::{EvaluationStatus, evaluate_xlsx_workbook};
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
    let cells = report
        .triggered
        .iter()
        .map(|trigger| trigger.cell.as_str())
        .collect::<Vec<_>>();
    assert_eq!(cells, ["B2", "C5", "C7", "E2"]);
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
