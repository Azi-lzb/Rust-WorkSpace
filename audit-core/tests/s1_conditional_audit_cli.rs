use serde_json::Value;
use std::{path::PathBuf, process::Command};

fn fixture() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("s1_conditional_eval.xlsx")
}

fn run_cli(developer_view: bool) -> std::process::Output {
    let workbook = fixture();
    let mut command = Command::new(env!("CARGO_BIN_EXE_open-data-audit"));
    command
        .arg("conditional-audit")
        .arg("--workbook")
        .arg(&workbook)
        .arg("--period")
        .arg("2026-07-31")
        .arg("--batch-id")
        .arg("synthetic-batch")
        .arg("--audit-time")
        .arg("2026-10-24 12:00:00");
    if developer_view {
        command.arg("--developer-view");
    }
    command.output().expect("run open-data-audit CLI")
}

#[test]
fn conditional_audit_cli_returns_issues_and_both_log_views() {
    let user = run_cli(false);
    assert!(
        user.status.success(),
        "{}",
        String::from_utf8_lossy(&user.stderr)
    );
    let user_report: Value = serde_json::from_slice(&user.stdout).expect("valid CLI JSON");
    assert_eq!(user_report["issues"].as_array().unwrap().len(), 4);
    assert!(String::from_utf8_lossy(&user.stderr).contains("条件格式检查完成"));

    let developer = run_cli(true);
    assert!(
        developer.status.success(),
        "{}",
        String::from_utf8_lossy(&developer.stderr)
    );
    let developer_report: Value =
        serde_json::from_slice(&developer.stdout).expect("valid developer-view CLI JSON");
    assert_eq!(developer_report["issues"].as_array().unwrap().len(), 4);
    let developer_log = String::from_utf8_lossy(&developer.stderr);
    assert!(developer_log.contains("conditional_format_issue_extraction"));
    assert!(developer_log.contains("calamine+namespace_aware_ooxml"));
    assert!(developer_log.contains("evaluation_time_ms"));
}
