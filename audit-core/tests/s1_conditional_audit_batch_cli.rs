use serde_json::{json, Value};
use std::{
    fs,
    path::PathBuf,
    process::{Command, Output},
    time::{SystemTime, UNIX_EPOCH},
};

fn fixture() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("s1_conditional_eval.xlsx")
}

fn manifest() -> PathBuf {
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system clock after Unix epoch")
        .as_nanos();
    let path = std::env::temp_dir().join(format!(
        "open_data_audit_s1_batch_{}_{}.json",
        std::process::id(),
        nonce
    ));
    let document = json!({
        "files": [
            {
                "workbook": fixture(),
                "source_file": "甲银行_20260731.xlsx",
                "org_code": "9144",
                "org_name": "甲银行"
            },
            {
                "workbook": fixture(),
                "source_file": "乙银行_20260731.xlsx",
                "org_code": "9145",
                "org_name": "乙银行"
            }
        ]
    });
    fs::write(
        &path,
        serde_json::to_vec_pretty(&document).expect("serialize manifest"),
    )
    .expect("write temporary test manifest");
    path
}

fn failing_manifest() -> PathBuf {
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system clock after Unix epoch")
        .as_nanos();
    let path = std::env::temp_dir().join(format!(
        "open_data_audit_s1_batch_failure_{}_{}.json",
        std::process::id(),
        nonce
    ));
    let document = json!({
        "files": [
            {
                "workbook": fixture(),
                "source_file": "甲银行_20260731.xlsx",
                "org_code": "9144",
                "org_name": "甲银行"
            },
            {
                "workbook": "missing-workbook.xlsx",
                "source_file": "乙银行_20260731.xlsx",
                "org_code": "9145",
                "org_name": "乙银行"
            }
        ]
    });
    fs::write(
        &path,
        serde_json::to_vec_pretty(&document).expect("serialize failing manifest"),
    )
    .expect("write temporary failing batch manifest");
    path
}

fn run_cli(manifest: &std::path::Path, developer_view: bool) -> Output {
    let mut command = Command::new(env!("CARGO_BIN_EXE_open-data-audit"));
    command
        .arg("conditional-audit-batch")
        .arg("--manifest")
        .arg(manifest)
        .arg("--period")
        .arg("2026-07-31")
        .arg("--batch-id")
        .arg("batch-1")
        .arg("--audit-time")
        .arg("2026-10-24 12:00:00");
    if developer_view {
        command.arg("--developer-view");
    }
    command.output().expect("run batch conditional-audit CLI")
}

#[test]
fn conditional_audit_batch_cli_merges_issues_and_logs_per_file_engine_timings() {
    let manifest = manifest();
    let user = run_cli(&manifest, false);
    assert!(
        user.status.success(),
        "{}",
        String::from_utf8_lossy(&user.stderr)
    );
    let user_report: Value = serde_json::from_slice(&user.stdout).expect("valid batch JSON");
    assert_eq!(user_report["files"].as_array().unwrap().len(), 2);
    assert_eq!(user_report["issues"].as_array().unwrap().len(), 8);
    assert_eq!(user_report["files"][0]["issue_count"], 4);
    assert_eq!(user_report["files"][1]["issue_count"], 4);
    assert_eq!(user_report["issues"][0]["org_code"], "9144");
    assert_eq!(user_report["issues"][4]["org_code"], "9145");
    let user_log = String::from_utf8_lossy(&user.stderr);
    assert!(user_log.contains("甲银行_20260731.xlsx"));
    assert!(user_log.contains("引擎 calamine+namespace_aware_ooxml"));
    assert!(user_log.contains("读取"));

    let developer = run_cli(&manifest, true);
    assert!(
        developer.status.success(),
        "{}",
        String::from_utf8_lossy(&developer.stderr)
    );
    let developer_report: Value =
        serde_json::from_slice(&developer.stdout).expect("valid developer-view batch JSON");
    assert_eq!(developer_report["issues"].as_array().unwrap().len(), 8);
    let developer_events = String::from_utf8_lossy(&developer.stderr)
        .lines()
        .map(|line| serde_json::from_str::<Value>(line).expect("JSON progress line"))
        .collect::<Vec<_>>();
    assert!(developer_events.len() >= 6);
    let file_events = developer_events
        .iter()
        .filter_map(|event| event.get("event"))
        .collect::<Vec<_>>();
    assert_eq!(
        file_events
            .iter()
            .filter(|event| event["status"] == "started")
            .count(),
        2
    );
    assert_eq!(
        file_events
            .iter()
            .filter(|event| event["status"] == "completed")
            .count(),
        2
    );
    assert_eq!(
        file_events
            .iter()
            .find(|event| event["file_index"] == 1 && event["status"] == "completed")
            .unwrap()["reader"],
        "calamine+namespace_aware_ooxml"
    );
    assert_eq!(developer_events.last().unwrap()["issues"], 8);

    fs::remove_file(manifest).expect("remove temporary batch manifest");
}

#[test]
fn conditional_audit_batch_cli_stops_without_partial_result_on_file_error() {
    let manifest = failing_manifest();
    let user = run_cli(&manifest, false);
    assert!(!user.status.success());
    assert!(user.stdout.is_empty(), "partial JSON must not be returned");
    let user_log = String::from_utf8_lossy(&user.stderr);
    assert!(user_log.contains("甲银行_20260731.xlsx"));
    assert!(user_log.contains("乙银行_20260731.xlsx：条件格式检查失败"));

    let developer = run_cli(&manifest, true);
    assert!(!developer.status.success());
    assert!(
        developer.stdout.is_empty(),
        "partial JSON must not be returned"
    );
    let developer_events = String::from_utf8_lossy(&developer.stderr)
        .lines()
        .filter_map(|line| serde_json::from_str::<Value>(line).ok())
        .collect::<Vec<_>>();
    assert!(developer_events.len() >= 5);
    let file_events = developer_events
        .iter()
        .filter_map(|event| event.get("event"))
        .collect::<Vec<_>>();
    assert!(file_events
        .iter()
        .any(|event| event["file_index"] == 2 && event["status"] == "failed"));
    assert_eq!(developer_events.last().unwrap()["status"], "failed");

    fs::remove_file(manifest).expect("remove temporary failing batch manifest");
}
