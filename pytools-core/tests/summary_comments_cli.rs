use anyhow::{Context, Result};
use excel_reader::read_workbook;
use serde_json::Value;
use std::{
    fs,
    path::{Path, PathBuf},
    process::{Command, Output},
    time::{SystemTime, UNIX_EPOCH},
};

#[test]
fn summary_comments_cli_logs_actual_stages_and_preserves_output_across_views() -> Result<()> {
    let template = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/summary_comments/comment_template.xlsx");
    let template_before = fs::read(&template)?;
    let root = unique_test_dir();
    fs::create_dir_all(&root)?;
    let source = root.join("comment_source.xlsx");
    create_source_workbook(&source)?;
    let source_before = fs::read(&source)?;

    let user = run_cli(&template, &[&source], &root.join("user"), None)?;
    assert!(
        user.status.success(),
        "user-view CLI failed: {}",
        String::from_utf8_lossy(&user.stderr)
    );
    let user_report = parse_stdout(&user)?;
    assert_eq!(user_report["saved"], true);
    assert_eq!(user_report["rows"], 3);
    assert_eq!(user_report["files"], 1);
    assert_eq!(user_report["execution"]["reader_engine"], "calamine");
    assert_eq!(user_report["execution"]["writer_engine"], "rust_xlsxwriter");
    assert!(user_report["execution"]["read_elapsed_ms"].is_number());
    assert!(user_report["execution"]["write_elapsed_ms"].is_number());
    assert!(user_report["execution"]["total_elapsed_ms"].is_number());
    let user_stderr = String::from_utf8_lossy(&user.stderr);
    assert!(user_stderr.contains("正在读取批注汇总模板"));
    assert!(user_stderr.contains("批注汇总工作簿已保存"));
    assert!(!user_stderr.contains("calamine"));
    assert!(!user_stderr.contains("rust_xlsxwriter"));
    let user_history = read_history(&user_report)?;
    assert_eq!(user_history["product"], "pytools");
    assert_eq!(user_history["feature_id"], "PT-6-2");
    assert_eq!(user_history["function"], "summary-comments");
    assert_eq!(user_history["status"], "succeeded");
    assert_eq!(user_history["completed_files"], 1);
    assert_eq!(user_history["failed_files"], 0);
    let user_history_text = fs::read_to_string(
        user_report["run_log_path"]
            .as_str()
            .context("user run log path missing")?,
    )?;
    for value in ["华东", "甲", "乙"] {
        assert!(
            !user_history_text.contains(value),
            "run history must not contain raw cell value {value}"
        );
    }

    let unsupported = root.join("ignored.xls");
    fs::write(&unsupported, b"unsupported input skipped by PyCode")?;
    let developer = run_cli(
        &template,
        &[&source, &unsupported],
        &root.join("developer"),
        Some("developer"),
    )?;
    assert!(
        developer.status.success(),
        "developer-view CLI failed: {}",
        String::from_utf8_lossy(&developer.stderr)
    );
    let developer_report = parse_stdout(&developer)?;
    let developer_history = read_history(&developer_report)?;
    let developer_events = String::from_utf8_lossy(&developer.stderr)
        .lines()
        .map(serde_json::from_str::<Value>)
        .collect::<std::result::Result<Vec<_>, _>>()
        .context("parse developer progress events")?;
    assert!(developer_events.len() >= 11);
    assert_eq!(developer_events.first().unwrap()["phase"], "run");
    assert_eq!(developer_events.first().unwrap()["status"], "started");
    assert_eq!(developer_events.last().unwrap()["status"], "succeeded");
    assert_eq!(
        developer_history["events"],
        serde_json::json!(developer_events)
    );
    for (index, event) in developer_events.iter().enumerate() {
        assert_eq!(event["sequence"], index + 1);
        assert!(event["total_elapsed_ms"].is_number());
    }
    for phase in ["template_read", "source_read", "row_extract", "output_save"] {
        assert!(
            developer_events.iter().any(|event| {
                event["phase"] == phase
                    && event["status"] == "succeeded"
                    && event["stage_elapsed_ms"].is_number()
            }),
            "completed {phase} event with elapsed time is missing"
        );
    }
    assert!(developer_events.iter().any(|event| {
        event["phase"] == "template_read"
            && event["actual_engine"] == "calamine + namespace-aware OOXML supplement"
            && event["components"].as_array().is_some_and(|components| {
                components
                    .iter()
                    .any(|component| component == "namespace-aware OOXML supplement")
            })
    }));
    assert!(developer_events.iter().any(|event| {
        event["phase"] == "source_read"
            && event["actual_engine"] == "calamine + namespace-aware OOXML supplement"
    }));
    assert!(developer_events.iter().any(|event| {
        event["phase"] == "input_filter"
            && event["status"] == "skipped"
            && event["file_name"] == "ignored.xls"
            && event["actual_engine"].is_null()
    }));
    let save_event = developer_events
        .iter()
        .find(|event| {
            event["phase"] == "output_save"
                && event["status"] == "succeeded"
                && event["actual_engine"] == "rust_xlsxwriter"
        })
        .context("rust_xlsxwriter save event is missing")?;
    assert!(save_event["stage_elapsed_ms"].is_number());
    assert_eq!(
        developer_events.last().unwrap()["progress"]["processed_rows"],
        3
    );
    assert_eq!(
        developer_events.last().unwrap()["progress"]["total_rows"],
        3
    );
    assert_eq!(
        developer_events.last().unwrap()["progress"]["completed_files"],
        1
    );

    let off = run_cli(&template, &[&source], &root.join("off"), Some("off"))?;
    assert!(
        off.status.success(),
        "off-view CLI failed: {}",
        String::from_utf8_lossy(&off.stderr)
    );
    assert!(off.stderr.is_empty(), "off view must not emit logs");
    let off_report = parse_stdout(&off)?;
    assert!(off_report.get("run_id").is_none());
    assert!(off_report.get("run_log_path").is_none());
    assert!(!root.join("off/logs").exists());

    let user_output = output_workbook(&user_report)?;
    let developer_output = output_workbook(&developer_report)?;
    let off_output = output_workbook(&off_report)?;
    let expected_sheets = serde_json::to_value(&read_workbook(&user_output)?.sheets)?;
    assert_eq!(
        expected_sheets,
        serde_json::to_value(&read_workbook(&developer_output)?.sheets)?
    );
    assert_eq!(
        expected_sheets,
        serde_json::to_value(&read_workbook(&off_output)?.sheets)?
    );
    let summary_sheet = read_workbook(&user_output)?;
    let summary_sheet = summary_sheet
        .sheets
        .iter()
        .find(|sheet| sheet.name == "模板")
        .context("comment summary output sheet missing")?;
    assert_eq!(
        summary_sheet.values.get("A2"),
        Some(&serde_json::json!({"type":"string", "value":"comment_source"}))
    );
    assert_eq!(
        summary_sheet.values.get("B2"),
        Some(&serde_json::json!({"type":"string", "value":"Branch_A"}))
    );
    assert_eq!(
        summary_sheet.values.get("D2"),
        Some(&serde_json::json!({"type":"number", "value":10.5}))
    );
    assert_eq!(fs::read(&template)?, template_before, "template changed");
    assert_eq!(fs::read(&source)?, source_before, "source workbook changed");

    fs::remove_dir_all(root)?;
    Ok(())
}

#[test]
fn summary_comments_cli_records_template_validation_failure() -> Result<()> {
    let root = unique_test_dir();
    fs::create_dir_all(&root)?;
    let invalid_template = root.join("template.txt");
    fs::write(&invalid_template, b"not a workbook")?;
    let source = root.join("source.xlsx");
    create_source_workbook(&source)?;

    let output = run_cli(
        &invalid_template,
        &[&source],
        &root.join("failure"),
        Some("developer"),
    )?;
    assert!(!output.status.success());
    let events = String::from_utf8_lossy(&output.stderr)
        .lines()
        .filter_map(|line| serde_json::from_str::<Value>(line).ok())
        .collect::<Vec<_>>();
    anyhow::ensure!(
        !events.is_empty(),
        "failed-run developer events are missing"
    );
    assert_eq!(events.last().unwrap()["phase"], "template_validation");
    assert_eq!(events.last().unwrap()["status"], "failed");
    assert_eq!(
        events.last().unwrap()["error_type"],
        "template_format_error"
    );
    let logs = fs::read_dir(root.join("failure/logs"))?
        .map(|entry| entry.map(|entry| entry.path()))
        .collect::<std::io::Result<Vec<_>>>()?;
    assert_eq!(logs.len(), 1);
    let history: Value = serde_json::from_slice(&fs::read(&logs[0])?)?;
    assert_eq!(history["feature_id"], "PT-6-2");
    assert_eq!(history["status"], "failed");
    assert_eq!(history["failure"]["stage"], "template_validation");
    assert_eq!(history["failure"]["error_type"], "template_format_error");

    fs::remove_dir_all(root)?;
    Ok(())
}

fn run_cli(
    template: &Path,
    sources: &[&Path],
    root: &Path,
    log_view: Option<&str>,
) -> Result<Output> {
    let mut command = Command::new(env!("CARGO_BIN_EXE_pytools"));
    command
        .arg("summary-comments")
        .arg("--template")
        .arg(template)
        .arg("--output-dir")
        .arg(root.join("output"))
        .arg("--log-dir")
        .arg(root.join("logs"));
    if let Some(log_view) = log_view {
        command.arg("--log-view").arg(log_view);
    }
    Ok(command
        .args(sources)
        .output()
        .context("run summary-comments CLI")?)
}

fn parse_stdout(output: &Output) -> Result<Value> {
    serde_json::from_slice(&output.stdout).context("parse summary-comments CLI JSON")
}

fn read_history(report: &Value) -> Result<Value> {
    let path = report["run_log_path"]
        .as_str()
        .context("run log path missing from CLI output")?;
    serde_json::from_slice(&fs::read(path)?).context("parse PT-6-2 run history")
}

fn output_workbook(report: &Value) -> Result<PathBuf> {
    Ok(PathBuf::from(
        report["path"]
            .as_str()
            .context("summary output path missing from CLI output")?,
    ))
}

fn create_source_workbook(path: &Path) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    {
        let worksheet = workbook.add_worksheet();
        worksheet.set_name("Branch_A")?;
        worksheet.write_string(0, 1, "余额")?;
        worksheet.write_string(0, 2, "余额")?;
        worksheet.write_string(0, 3, "区域")?;
        worksheet.write_number(1, 1, 10.5)?;
        worksheet.write_string(1, 2, "甲")?;
        worksheet.write_number(2, 1, 20.0)?;
        worksheet.write_string(2, 2, "乙")?;
        worksheet.write_number(3, 1, 0.0)?;
        worksheet.write_string(3, 2, "丙")?;
        worksheet.write_string(0, 3, "华东")?;
    }
    workbook
        .save(path)
        .context("save synthetic source workbook")
}

fn unique_test_dir() -> PathBuf {
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system clock after Unix epoch")
        .as_nanos();
    std::env::temp_dir().join(format!(
        "pytools-summary-comments-{}-{nonce}",
        std::process::id()
    ))
}
