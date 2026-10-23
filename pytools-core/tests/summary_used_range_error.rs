use anyhow::{Context, Result};
use serde_json::Value;
use std::{
    fs,
    path::{Path, PathBuf},
    process::Command,
    time::{SystemTime, UNIX_EPOCH},
};

#[test]
fn summary_used_range_cli_preserves_excel_error_cells() -> Result<()> {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/summary_used_range/error_value.xlsx");
    let original = fs::read(&fixture)?;
    let output_dir = unique_test_dir();
    let unsupported_dir = output_dir.join("unsupported_inputs");
    fs::create_dir_all(&unsupported_dir)?;
    let unsupported_xls = unsupported_dir.join("ignored.xls");
    let unsupported_xlsb = unsupported_dir.join("ignored.xlsb");
    fs::write(
        &unsupported_xls,
        b"not an Excel workbook; PyCode skips .xls",
    )?;
    fs::write(
        &unsupported_xlsb,
        b"not an Excel workbook; PyCode skips .xlsb",
    )?;
    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("summary-used-range")
        .arg("--output-dir")
        .arg(output_dir.join("output"))
        .arg("--log-dir")
        .arg(output_dir.join("logs"))
        .arg(&fixture)
        .arg(&unsupported_xls)
        .arg(&unsupported_xlsb)
        .output()
        .context("run pytools summary-used-range")?;
    assert!(
        output.status.success(),
        "CLI failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report: Value = serde_json::from_slice(&output.stdout).context("parse CLI JSON")?;
    assert_eq!(report["saved"], true);
    assert_eq!(report["rows"], 2);
    assert_eq!(report["files"], 1);
    assert!(report["run_log_path"].is_string());
    assert_eq!(report["execution"]["reader_engine"], "calamine");
    assert_eq!(report["execution"]["writer_engine"], "rust_xlsxwriter");
    assert_eq!(report["execution"]["ooxml_error_cells_patched"], 7);
    assert!(report["execution"]["total_elapsed_ms"].is_number());
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.contains("正在扫描工作簿范围"));
    assert!(stderr.contains("工作簿读取完成"));
    assert!(stderr.contains("汇总工作簿生成完成"));
    assert!(stderr.contains("忽略不支持的输入"));
    assert!(!stderr.contains("calamine"));
    assert!(!stderr.contains("rust_xlsxwriter"));

    let run_log_path = PathBuf::from(
        report["run_log_path"]
            .as_str()
            .context("CLI run log path must be a string")?,
    );
    let run_log_bytes = fs::read(&run_log_path)?;
    let run_log_text = String::from_utf8_lossy(&run_log_bytes);
    let run_log: Value = serde_json::from_slice(&run_log_bytes).context("parse PT-6-1 run log")?;
    assert_eq!(run_log["product"], "pytools");
    assert_eq!(run_log["feature_id"], "PT-6-1");
    assert_eq!(run_log["status"], "succeeded");
    assert_eq!(run_log["total_inputs"], 3);
    assert_eq!(run_log["events"][0]["progress"]["total_files"], 1);
    assert!(run_log["events"]
        .as_array()
        .is_some_and(|events| !events.is_empty()));
    for extension in ["xls", "xlsb"] {
        assert!(run_log["events"].as_array().unwrap().iter().any(|event| {
            event["phase"] == "input_filter"
                && event["status"] == "skipped"
                && event["file_name"] == format!("ignored.{extension}")
                && event["actual_engine"].is_null()
        }));
    }
    assert!(run_log_text.contains("calamine"));
    assert!(run_log_text.contains("rust_xlsxwriter"));
    assert!(!run_log_text.contains("#DATA!"));

    let developer_output_dir = output_dir.join("developer");
    let developer = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("summary-used-range")
        .arg("--output-dir")
        .arg(developer_output_dir.join("output"))
        .arg("--log-dir")
        .arg(developer_output_dir.join("logs"))
        .arg("--log-view")
        .arg("developer")
        .arg(&fixture)
        .arg(&unsupported_xls)
        .arg(&unsupported_xlsb)
        .output()
        .context("run pytools summary-used-range in developer log view")?;
    assert!(
        developer.status.success(),
        "developer CLI failed: {}",
        String::from_utf8_lossy(&developer.stderr)
    );
    let developer_report: Value =
        serde_json::from_slice(&developer.stdout).context("parse developer CLI JSON")?;
    let developer_stderr = String::from_utf8_lossy(&developer.stderr);
    let developer_events = developer_stderr
        .lines()
        .map(serde_json::from_str::<Value>)
        .collect::<std::result::Result<Vec<_>, _>>()
        .context("parse developer progress events")?;
    assert!(developer_events.len() >= 8);
    assert_eq!(developer_events.first().unwrap()["phase"], "run");
    assert_eq!(developer_events.first().unwrap()["status"], "started");
    assert_eq!(developer_events.last().unwrap()["status"], "succeeded");
    assert!(developer_events.iter().any(|event| {
        event["phase"] == "value_read"
            && event["actual_engine"] == "calamine"
            && event["stage_elapsed_ms"].is_number()
    }));
    assert!(developer_events.iter().any(|event| {
        event["phase"] == "output_save"
            && event["actual_engine"] == "rust_xlsxwriter"
            && event["stage_elapsed_ms"].is_number()
    }));
    let workbook_write_event = developer_events
        .iter()
        .find(|event| {
            event["phase"] == "output_workbook"
                && event["status"] == "succeeded"
                && event["actual_engine"] == "rust_xlsxwriter"
                && event["stage_elapsed_ms"].is_number()
        })
        .context("completed output workbook event missing")?;
    assert_eq!(
        developer_report["execution"]["write_elapsed_ms"],
        workbook_write_event["stage_elapsed_ms"]
    );
    assert!(developer_events
        .iter()
        .filter(|event| { event["phase"] == "input_filter" && event["status"] == "skipped" })
        .all(|event| event["actual_engine"].is_null()));
    for (index, event) in developer_events.iter().enumerate() {
        assert_eq!(event["sequence"], index + 1);
    }
    let developer_log_path = PathBuf::from(
        developer_report["run_log_path"]
            .as_str()
            .context("developer CLI run log path must be a string")?,
    );
    let developer_log: Value = serde_json::from_slice(&fs::read(developer_log_path)?)
        .context("parse developer PT-6-1 run log")?;
    assert_eq!(
        developer_log["events"].as_array().unwrap().len(),
        developer_events.len()
    );

    let quiet_output_dir = output_dir.join("quiet");
    let quiet = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("summary-used-range")
        .arg("--output-dir")
        .arg(quiet_output_dir.join("output"))
        .arg("--log-dir")
        .arg(quiet_output_dir.join("logs"))
        .arg("--log-view")
        .arg("off")
        .arg(&fixture)
        .arg(&unsupported_xls)
        .arg(&unsupported_xlsb)
        .output()
        .context("run pytools summary-used-range with logging disabled")?;
    assert!(
        quiet.status.success(),
        "quiet CLI failed: {}",
        String::from_utf8_lossy(&quiet.stderr)
    );
    assert!(
        quiet.stderr.is_empty(),
        "off view must not emit progress logs"
    );
    let quiet_report: Value =
        serde_json::from_slice(&quiet.stdout).context("parse quiet CLI JSON")?;
    assert!(quiet_report.get("run_id").is_none());
    assert!(quiet_report.get("run_log_path").is_none());
    assert!(!quiet_output_dir.join("logs").exists());

    let path = PathBuf::from(
        report["path"]
            .as_str()
            .context("CLI report path must be a string")?,
    );
    let workbook = excel_reader::read_workbook(&path)?;
    let sheet = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "汇总")
        .context("summary sheet missing")?;
    let expected_errors = [
        ("D2", "Div0"),
        ("E2", "NA"),
        ("F2", "Name"),
        ("G2", "Null"),
        ("H2", "Num"),
        ("I2", "Ref"),
        ("J2", "Value"),
    ];
    for (address, error) in expected_errors {
        assert_eq!(
            sheet.values.get(address),
            Some(&serde_json::json!({"type":"error", "value":error}))
        );
    }
    assert_eq!(
        sheet.values.get("A3"),
        Some(&serde_json::json!({"type":"string", "value":"error_value"}))
    );
    assert_eq!(
        sheet.values.get("B3"),
        Some(&serde_json::json!({"type":"string", "value":"Empty"}))
    );
    assert!(
        sheet.formulas.is_empty(),
        "summary must export values, not formulas"
    );
    let user_projection = serde_json::to_value(&workbook.sheets)?;
    let developer_path = PathBuf::from(
        developer_report["path"]
            .as_str()
            .context("developer summary path must be a string")?,
    );
    let developer_workbook = excel_reader::read_workbook(&developer_path)?;
    assert_eq!(
        user_projection,
        serde_json::to_value(&developer_workbook.sheets)?
    );
    let quiet_path = PathBuf::from(
        quiet_report["path"]
            .as_str()
            .context("quiet summary path must be a string")?,
    );
    let quiet_workbook = excel_reader::read_workbook(&quiet_path)?;
    assert_eq!(
        user_projection,
        serde_json::to_value(&quiet_workbook.sheets)?
    );
    assert_eq!(fs::read(&fixture)?, original, "source workbook changed");

    fs::remove_dir_all(output_dir)?;
    Ok(())
}

#[test]
fn summary_used_range_cli_fails_closed_on_unknown_xlsx_error_value() -> Result<()> {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/summary_used_range/getting_data_error.xlsx");
    let original = fs::read(&fixture)?;
    let output_dir = unique_test_dir();
    let summary_dir = output_dir.join("output");
    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("summary-used-range")
        .arg("--output-dir")
        .arg(&summary_dir)
        .arg("--log-dir")
        .arg(output_dir.join("logs"))
        .arg(&fixture)
        .output()
        .context("run pytools summary-used-range with an unknown XLSX error")?;
    assert!(!output.status.success(), "unknown error should fail closed");
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(
        stderr.contains("Unsupported cell error value '#DATA!'"),
        "unexpected CLI error: {stderr}"
    );
    let run_logs = fs::read_dir(output_dir.join("logs"))?
        .filter_map(std::result::Result::ok)
        .map(|entry| entry.path())
        .filter(|path| {
            path.extension()
                .is_some_and(|extension| extension == "json")
        })
        .collect::<Vec<_>>();
    assert_eq!(
        run_logs.len(),
        1,
        "failed run should keep one structured log"
    );
    let run_log_bytes = fs::read(&run_logs[0])?;
    let run_log_text = String::from_utf8_lossy(&run_log_bytes);
    let run_log: Value =
        serde_json::from_slice(&run_log_bytes).context("parse failed PT-6-1 run log")?;
    assert_eq!(run_log["status"], "failed");
    assert_eq!(run_log["failure"]["stage"], "value_read");
    assert_eq!(run_log["failure"]["error_type"], "value_read_error");
    assert_eq!(
        run_log["events"].as_array().unwrap().last().unwrap()["status"],
        "failed"
    );
    assert!(
        !run_log_text.contains("#DATA!"),
        "run log must omit cell contents"
    );
    assert!(
        fs::read_dir(&summary_dir)?.next().is_none(),
        "failed read must not create a partial summary"
    );
    assert_eq!(fs::read(&fixture)?, original, "source workbook changed");

    fs::remove_dir_all(output_dir)?;
    Ok(())
}

#[test]
fn summary_used_range_cli_reports_no_reader_when_no_workbook_is_read() -> Result<()> {
    let output_dir = unique_test_dir();
    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("summary-used-range")
        .arg("--output-dir")
        .arg(output_dir.join("output"))
        .arg("--log-dir")
        .arg(output_dir.join("logs"))
        .arg(output_dir.join("missing.xlsx"))
        .output()
        .context("run pytools summary-used-range with a missing source")?;
    assert!(
        output.status.success(),
        "CLI failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report: Value = serde_json::from_slice(&output.stdout).context("parse CLI JSON")?;
    assert_eq!(report["saved"], false);
    assert_eq!(report["files"], 0);
    assert!(report["execution"]["reader_engine"].is_null());
    assert!(report["execution"]["writer_engine"].is_null());
    let run_log_path = PathBuf::from(
        report["run_log_path"]
            .as_str()
            .context("skipped CLI run log path must be a string")?,
    );
    let run_log: Value =
        serde_json::from_slice(&fs::read(run_log_path)?).context("parse skipped PT-6-1 run log")?;
    assert_eq!(run_log["status"], "skipped");
    assert_eq!(
        run_log["events"].as_array().unwrap().last().unwrap()["status"],
        "skipped"
    );
    assert!(run_log["events"]
        .as_array()
        .unwrap()
        .iter()
        .all(|event| { event["actual_engine"].is_null() }));

    fs::remove_dir_all(output_dir)?;
    Ok(())
}

fn unique_test_dir() -> PathBuf {
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system clock after Unix epoch")
        .as_nanos();
    std::env::temp_dir().join(format!(
        "pytools-summary-error-{}-{nonce}",
        std::process::id()
    ))
}
