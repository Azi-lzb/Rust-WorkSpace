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
    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("summary-used-range")
        .arg("--output-dir")
        .arg(output_dir.join("output"))
        .arg(&fixture)
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
    assert_eq!(report["execution"]["reader_engine"], "calamine");
    assert_eq!(report["execution"]["writer_engine"], "rust_xlsxwriter");
    assert_eq!(report["execution"]["ooxml_error_cells_patched"], 7);
    assert!(report["execution"]["total_elapsed_ms"].is_number());

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
        .arg(&fixture)
        .output()
        .context("run pytools summary-used-range with an unknown XLSX error")?;
    assert!(!output.status.success(), "unknown error should fail closed");
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(
        stderr.contains("Unsupported cell error value '#DATA!'"),
        "unexpected CLI error: {stderr}"
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
