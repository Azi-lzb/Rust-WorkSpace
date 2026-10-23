use serde_json::{json, Value};
use std::{
    fs,
    process::Command,
    time::{SystemTime, UNIX_EPOCH},
};

fn issue(issue_id: &str, period: &str, triggered: bool) -> Value {
    json!({
        "issue_id": issue_id,
        "period": period,
        "batch_id": "old-batch",
        "audit_time": "old-time",
        "triggered": triggered,
        "status": "新增",
        "first_seen_period": period,
        "previous_seen_period": "",
        "consecutive_count": 1,
        "org_code": "001",
        "org_name": "测试机构",
        "report_code": "R01",
        "sheet_name": "Sheet1",
        "rule_id": issue_id,
        "severity": "错误",
        "formula_cell": "C2",
        "target_cell": "B2",
        "target_value": 1,
        "formula_result": "错误",
        "message": "测试",
        "source_file": "source.xlsx",
        "audit_file": "audit.xlsx",
        "check_field": "",
        "check_type": "",
        "comparison_value": "",
        "reference_value": "",
        "difference_value": "",
        "detail": "",
        "institution_feedback": "",
        "auditor_opinion": ""
    })
}

fn run(input: &std::path::Path, developer_view: bool) -> std::process::Output {
    let mut command = Command::new(env!("CARGO_BIN_EXE_open-data-audit"));
    command.arg("s1-history-classify").arg("--input").arg(input);
    if developer_view {
        command.arg("--developer-view");
    }
    command.output().expect("run S1 history CLI")
}

#[test]
fn history_cli_classifies_issues_and_reports_user_and_developer_logs() {
    let unique = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("clock is after epoch")
        .as_nanos();
    let path = std::env::temp_dir().join(format!(
        "audit_s1_history_cli_{}_{}.json",
        std::process::id(),
        unique
    ));
    let mut continuous_old = issue("CONT", "2026-01", true);
    continuous_old["first_seen_period"] = json!("2025-12");
    let mut continuous_latest = issue("CONT", "2026-02", true);
    continuous_latest["first_seen_period"] = json!("2025-12");
    continuous_latest["consecutive_count"] = json!(2);
    let resolved_old = issue("RESOLVED", "2026-02", true);
    let current_continuous = issue("CONT", "2026-03", true);
    let current_new = issue("NEW", "2026-03", true);
    let input = json!({
        "period": "2026-03",
        "batch_id": "batch-new",
        "audit_time": "new-time",
        "audited_org_codes": ["001"],
        "current_issues": [current_continuous, current_new],
        "history": [continuous_old, continuous_latest, resolved_old]
    });
    fs::write(&path, serde_json::to_vec(&input).expect("serialize input"))
        .expect("write temporary history input");

    let ordinary = run(&path, false);
    assert!(
        ordinary.status.success(),
        "{}",
        String::from_utf8_lossy(&ordinary.stderr)
    );
    let output: Value = serde_json::from_slice(&ordinary.stdout).expect("parse result JSON");
    assert_eq!(output["current_issues"][0]["status"], "连续出现");
    assert_eq!(output["current_issues"][0]["consecutive_count"], 3);
    assert_eq!(output["current_issues"][1]["status"], "新增");
    assert_eq!(output["resolved_issues"][0]["issue_id"], "RESOLVED");
    assert_eq!(output["resolved_issues"][0]["batch_id"], "batch-new");
    let ordinary_log = String::from_utf8_lossy(&ordinary.stderr);
    assert!(ordinary_log.contains("历史归类完成"));
    assert!(ordinary_log.contains("引擎 Rust"));
    assert!(ordinary_log.contains("耗时"));

    let developer = run(&path, true);
    assert!(
        developer.status.success(),
        "{}",
        String::from_utf8_lossy(&developer.stderr)
    );
    let developer_output: Value =
        serde_json::from_slice(&developer.stdout).expect("parse developer result JSON");
    assert_eq!(developer_output["current_issues"], output["current_issues"]);
    assert_eq!(
        developer_output["resolved_issues"],
        output["resolved_issues"]
    );
    let developer_log = String::from_utf8_lossy(&developer.stderr);
    assert!(developer_log.contains("history_issue_classification"));
    assert!(developer_log.contains("rust_s1_history"));
    assert!(developer_log.contains("elapsed_ms"));

    fs::remove_file(path).expect("remove temporary history input");
}
