use audit_core::s1_conditional_audit::{extract_conditional_issues, S1ConditionalMetadata};
use std::path::PathBuf;

fn fixture() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("s1_conditional_eval.xlsx")
}

#[test]
fn constructs_s1_issue_records_from_python_parity_fixture() {
    let workbook = fixture();
    let metadata = S1ConditionalMetadata {
        period: "2026-07-31".to_owned(),
        batch_id: "synthetic-batch".to_owned(),
        audit_time: "2026-10-24 12:00:00".to_owned(),
        org_code: "9144".to_owned(),
        org_name: "样例机构".to_owned(),
    };
    let result = extract_conditional_issues(&workbook, &workbook, &[], &[], &metadata)
        .expect("extract S1 conditional issues");

    assert_eq!(result.issues.len(), 4);
    assert_eq!(result.status, "completed_with_diagnostics");
    assert_eq!(result.context.unresolved_indicator_count, 4);
    assert_eq!(result.context.comments_unverified_sheet_count, 0);
    let by_cell = result
        .issues
        .iter()
        .map(|issue| (issue.target_cell.as_str(), issue))
        .collect::<std::collections::BTreeMap<_, _>>();
    assert_eq!(by_cell["B2"].target_value.as_f64(), Some(5.0));
    assert_eq!(by_cell["B2"].message, "条件格式规则：B2>0");
    assert_eq!(by_cell["C7"].message, "条件格式规则：AND(C7>0,C7>1)");
    assert_eq!(by_cell["E2"].target_value, serde_json::Value::Null);
    assert_eq!(by_cell["B2"].period, "2026-07-31");
    assert_eq!(by_cell["B2"].batch_id, "synthetic-batch");
    assert_eq!(by_cell["B2"].org_code, "9144");
    assert_eq!(by_cell["B2"].org_name, "样例机构");
    assert_eq!(by_cell["B2"].rule_id, "条件格式填充");
    assert_eq!(by_cell["B2"].severity, "条件格式触发");
    assert_eq!(
        by_cell["B2"].issue_id,
        "s1_conditional_eval｜S1Conditional｜B2｜校验指标未识别"
    );
}
