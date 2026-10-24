use serde_json::Value;
use std::{
    fs,
    path::PathBuf,
    process::{Command, Output},
    time::{SystemTime, UNIX_EPOCH},
};

fn selected_workbooks() -> (PathBuf, Vec<PathBuf>) {
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system clock after Unix epoch")
        .as_nanos();
    let directory = std::env::temp_dir().join(format!(
        "audit_s5_file_hints_cli_{}_{}",
        std::process::id(),
        nonce
    ));
    fs::create_dir_all(&directory).expect("create temporary S5 input directory");
    let files = ["source_20260831.xlsx", "source_20260731_20260831.xlsx"]
        .into_iter()
        .map(|name| {
            let path = directory.join(name);
            fs::write(&path, []).expect("create filename-only selected input");
            path
        })
        .collect();
    (directory, files)
}

fn run_cli(files: &[PathBuf], developer_view: bool) -> Output {
    let mut command = Command::new(env!("CARGO_BIN_EXE_open-data-audit"));
    command.arg("s5-file-period-hints");
    for path in files {
        command.arg("--input-file").arg(path);
    }
    if developer_view {
        command.arg("--developer-view");
    }
    command.output().expect("run S5 file period hints CLI")
}

#[test]
fn s5_file_period_hints_cli_reports_both_log_views_without_workbook_reader() {
    let (directory, files) = selected_workbooks();
    let user = run_cli(&files, false);
    let developer = run_cli(&files, true);
    fs::remove_dir_all(&directory).expect("remove temporary S5 inputs");

    assert!(
        user.status.success(),
        "{}",
        String::from_utf8_lossy(&user.stderr)
    );
    let user_output: Value = serde_json::from_slice(&user.stdout).expect("valid user-view JSON");
    assert_eq!(user_output["status"], "completed");
    assert_eq!(user_output["selected_file_count"], 2);
    assert_eq!(user_output["reader"], "none");
    assert_eq!(user_output["items"][0]["period_hint"], "2026-08-31");
    assert_eq!(user_output["items"][1]["period_hint"], "多个日期");
    let user_log = String::from_utf8_lossy(&user.stderr);
    assert!(user_log.contains("不读取工作簿"));
    assert!(user_log.contains("耗时"));

    assert!(
        developer.status.success(),
        "{}",
        String::from_utf8_lossy(&developer.stderr)
    );
    let developer_output: Value =
        serde_json::from_slice(&developer.stdout).expect("valid developer-view JSON");
    assert_eq!(developer_output["selected_file_count"], 2);
    let events = String::from_utf8_lossy(&developer.stderr)
        .lines()
        .map(|line| serde_json::from_str::<Value>(line).expect("valid progress JSON"))
        .collect::<Vec<_>>();
    assert_eq!(events.len(), 2);
    assert_eq!(events[0]["step"], "file_period_hints");
    assert_eq!(events[0]["status"], "started");
    assert_eq!(events[1]["status"], "completed");
    assert_eq!(events[1]["engine"], "rust_path_name_parser");
    assert_eq!(events[1]["reader"], "none");
    assert!(events[1]["elapsed_ms"].as_u64().is_some());
}
