use serde_json::Value;
use std::{
    fs,
    path::{Path, PathBuf},
    process::{Command, Output},
    time::{SystemTime, UNIX_EPOCH},
};

fn temporary_directory() -> PathBuf {
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system clock after Unix epoch")
        .as_nanos();
    let path = std::env::temp_dir().join(format!(
        "audit_s5_template_match_cli_{}_{}",
        std::process::id(),
        nonce
    ));
    fs::create_dir_all(&path).expect("create temporary S5 template-match directory");
    path
}

fn fixture() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("s1_region_summary_template.xlsx")
}

fn run_cli(template_dir: &Path, input: &Path, developer_view: bool) -> Output {
    let mut command = Command::new(env!("CARGO_BIN_EXE_open-data-audit"));
    command
        .arg("s5-match-template")
        .arg("--template-dir")
        .arg(template_dir)
        .arg("--input-file")
        .arg(input);
    if developer_view {
        command.arg("--developer-view");
    }
    command.output().expect("run S5 template-matching CLI")
}

#[test]
fn s5_template_match_cli_reports_result_and_both_log_views() {
    let directory = temporary_directory();
    let template_dir = directory.join("templates");
    fs::create_dir_all(&template_dir).expect("create template directory");
    let template = template_dir.join("selected-template.xlsx");
    let input = directory.join("source.xlsx");
    fs::copy(fixture(), &template).expect("copy synthetic template fixture");
    fs::copy(fixture(), &input).expect("copy synthetic source fixture");

    let user = run_cli(&template_dir, &input, false);
    assert!(
        user.status.success(),
        "{}",
        String::from_utf8_lossy(&user.stderr)
    );
    let user_output: Value = serde_json::from_slice(&user.stdout).expect("valid user JSON");
    assert_eq!(user_output["status"], "completed");
    assert_eq!(user_output["selected_file_count"], 1);
    assert_eq!(user_output["reader"], "excel_reader_ooxml_sheet_names");
    assert_eq!(
        Path::new(user_output["template_path"].as_str().unwrap())
            .file_name()
            .unwrap(),
        "selected-template.xlsx"
    );
    let user_log = String::from_utf8_lossy(&user.stderr);
    assert!(user_log.contains("OOXML 工作表名称读取"));
    assert!(user_log.contains("耗时"));

    let developer = run_cli(&template_dir, &input, true);
    assert!(
        developer.status.success(),
        "{}",
        String::from_utf8_lossy(&developer.stderr)
    );
    let events = String::from_utf8_lossy(&developer.stderr)
        .lines()
        .map(|line| serde_json::from_str::<Value>(line).expect("valid progress JSON"))
        .collect::<Vec<_>>();
    assert_eq!(events.len(), 2);
    assert_eq!(events[0]["step"], "template_matching");
    assert_eq!(events[0]["status"], "started");
    assert_eq!(events[1]["status"], "completed");
    assert_eq!(events[1]["engine"], "rust_s5_sheet_set_matcher");
    assert_eq!(events[1]["reader"], "excel_reader_ooxml_sheet_names");
    assert!(events[1]["elapsed_ms"].as_u64().is_some());

    fs::remove_dir_all(directory).expect("remove temporary S5 test files");
}

#[test]
fn s5_template_match_cli_rejects_ambiguous_profiles_and_reports_failure() {
    let directory = temporary_directory();
    let template_dir = directory.join("templates");
    fs::create_dir_all(&template_dir).expect("create template directory");
    let first = template_dir.join("template-a.xlsx");
    let second = template_dir.join("template-b.xlsx");
    let input = directory.join("source.xlsx");
    fs::copy(fixture(), &first).expect("copy first template fixture");
    fs::copy(fixture(), &second).expect("copy second template fixture");
    fs::copy(fixture(), &input).expect("copy synthetic source fixture");

    let output = run_cli(&template_dir, &input, true);
    assert_eq!(output.status.code(), Some(2));
    let events = String::from_utf8_lossy(&output.stderr)
        .lines()
        .filter_map(|line| serde_json::from_str::<Value>(line).ok())
        .collect::<Vec<_>>();
    assert_eq!(events.len(), 2);
    assert_eq!(events[1]["status"], "failed");
    assert!(events[1]["error"]
        .as_str()
        .unwrap()
        .contains("有多个相近模板"));
    assert!(String::from_utf8_lossy(&output.stderr).contains("open-data-audit:"));

    fs::remove_dir_all(directory).expect("remove temporary S5 test files");
}

#[test]
fn s5_template_match_cli_explains_pending_xlsm_reader_support() {
    let directory = temporary_directory();
    let template_dir = directory.join("templates");
    fs::create_dir_all(&template_dir).expect("create template directory");
    let input = directory.join("source.xlsm");
    fs::copy(fixture(), &input).expect("copy synthetic source fixture");

    let output = run_cli(&template_dir, &input, false);
    assert_eq!(output.status.code(), Some(2));
    assert!(String::from_utf8_lossy(&output.stderr).contains("仍待共享 excel-reader 协调"));

    fs::remove_dir_all(directory).expect("remove temporary S5 test files");
}
