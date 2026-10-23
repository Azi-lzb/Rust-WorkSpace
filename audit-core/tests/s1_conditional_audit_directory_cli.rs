use serde_json::Value;
use std::{
    fs,
    path::{Path, PathBuf},
    process::{Command, Output},
    time::{SystemTime, UNIX_EPOCH},
};

fn fixture() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("s1_conditional_eval.xlsx")
}

fn temporary_directory(label: &str) -> PathBuf {
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system clock after Unix epoch")
        .as_nanos();
    let path = std::env::temp_dir().join(format!(
        "open_data_audit_{label}_{}_{}",
        std::process::id(),
        nonce
    ));
    fs::create_dir_all(&path).expect("create temporary input directory");
    path
}

fn copy_fixture(directory: &Path, name: &str) {
    fs::copy(fixture(), directory.join(name)).expect("copy synthetic XLSX fixture");
}

fn run_cli(
    directory: &Path,
    period: Option<&str>,
    recursive: bool,
    developer_view: bool,
) -> Output {
    let mut command = Command::new(env!("CARGO_BIN_EXE_open-data-audit"));
    command
        .arg("conditional-audit-directory")
        .arg("--input-dir")
        .arg(directory)
        .arg("--batch-id")
        .arg("directory-batch")
        .arg("--audit-time")
        .arg("2026-10-24 12:00:00");
    if let Some(period) = period {
        command.arg("--period").arg(period);
    }
    if recursive {
        command.arg("--recursive");
    }
    if developer_view {
        command.arg("--developer-view");
    }
    command.output().expect("run directory S1 conditional CLI")
}

#[test]
fn directory_cli_discovers_period_derives_organization_and_logs_both_views() {
    let directory = temporary_directory("s1_directory");
    copy_fixture(&directory, "9144_甲银行_2026-07-31.xlsx");

    let user = run_cli(&directory, None, false, false);
    assert!(
        user.status.success(),
        "{}",
        String::from_utf8_lossy(&user.stderr)
    );
    let user_report: Value = serde_json::from_slice(&user.stdout).expect("valid batch JSON");
    assert_eq!(user_report["files"].as_array().unwrap().len(), 1);
    assert_eq!(user_report["issues"].as_array().unwrap().len(), 4);
    assert_eq!(user_report["issues"][0]["period"], "2026-07");
    assert_eq!(user_report["issues"][0]["org_code"], "9144");
    assert_eq!(user_report["issues"][0]["org_name"], "甲银行");
    let user_log = String::from_utf8_lossy(&user.stderr);
    assert!(user_log.contains("源目录发现 1 个 XLSX"));
    assert!(user_log.contains("数据期 2026-07（源文件名）"));
    assert!(user_log.contains("引擎 calamine+namespace_aware_ooxml"));
    assert!(user_log.contains("读取"));

    let developer = run_cli(&directory, None, false, true);
    assert!(
        developer.status.success(),
        "{}",
        String::from_utf8_lossy(&developer.stderr)
    );
    let events = String::from_utf8_lossy(&developer.stderr)
        .lines()
        .map(|line| serde_json::from_str::<Value>(line).expect("JSON progress line"))
        .collect::<Vec<_>>();
    let discovery = events
        .iter()
        .find(|event| {
            event["step"] == "source_workbook_discovery" && event["status"] == "completed"
        })
        .expect("discovery completion event");
    assert_eq!(discovery["files_total"], 1);
    assert_eq!(discovery["period"], "2026-07");
    assert_eq!(discovery["period_source"], "源文件名");
    assert!(discovery["elapsed_ms"].as_u64().is_some());
    assert!(events.iter().any(|event| {
        event["step"] == "conditional_format_issue_extraction" && event["status"] == "completed"
    }));

    fs::remove_dir_all(directory).expect("remove temporary input directory");
}

#[test]
fn directory_cli_requires_explicit_period_when_filenames_conflict() {
    let directory = temporary_directory("s1_period_conflict");
    copy_fixture(&directory, "9144_甲银行_2026-06-30.xlsx");
    copy_fixture(&directory, "9145_乙银行_2026-07-31.xlsx");

    let output = run_cli(&directory, None, false, false);
    assert!(!output.status.success());
    assert!(output.stdout.is_empty());
    let error = String::from_utf8_lossy(&output.stderr);
    assert!(error.contains("多个数据期"), "{error}");
    assert!(error.contains("请显式指定 --period"), "{error}");

    fs::remove_dir_all(directory).expect("remove temporary input directory");
}

#[test]
fn directory_cli_recurses_only_when_requested() {
    let directory = temporary_directory("s1_directory_recursive");
    let nested = directory.join("nested");
    fs::create_dir_all(&nested).expect("create nested input directory");
    copy_fixture(&nested, "9144_甲银行_2026-07-31.xlsx");

    let root_only = run_cli(&directory, Some("2026-07"), false, false);
    assert!(!root_only.status.success());
    assert!(String::from_utf8_lossy(&root_only.stderr).contains("没有可审核的 .xlsx 文件"));

    let recursive = run_cli(&directory, Some("2026-07"), true, false);
    assert!(
        recursive.status.success(),
        "{}",
        String::from_utf8_lossy(&recursive.stderr)
    );
    let report: Value =
        serde_json::from_slice(&recursive.stdout).expect("valid recursive batch JSON");
    assert_eq!(report["files"].as_array().unwrap().len(), 1);
    assert_eq!(report["issues"].as_array().unwrap().len(), 4);

    fs::remove_dir_all(directory).expect("remove temporary input directory");
}
