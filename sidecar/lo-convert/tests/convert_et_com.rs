#[path = "support/et_fixture.rs"]
mod support;

use anyhow::{ensure, Context, Result};
use excel_reader::{read_workbook, WorkbookDump};
use std::{env, fs, path::Path, process::Command};
use support::{copy_and_seed_excel8_et_named_fixture, copy_native_et_fixture, temp_test_dir};

#[test]
fn com_conversion_validates_cfb_excel8_with_et_extension_data_and_formula() -> Result<()> {
    if !wps_et_tests_enabled() {
        return Ok(());
    }
    let Some(fixture) = env::var_os("LO_CONVERT_CFB_ET_NAMED_XLS_FIXTURE") else {
        eprintln!("skipping CFB/.et extension regression: fixture path is not set");
        return Ok(());
    };
    let root = temp_test_dir("lo_convert_cfb_et_extension")?;
    let input = copy_and_seed_excel8_et_named_fixture(&root.0, Path::new(&fixture))?;
    let workbook = convert_and_read(&input, &root.0)?;
    let data = workbook
        .sheets
        .first()
        .context("converted workbook has no worksheets")?;
    ensure!(
        data.name == "Data",
        "converted sheet name changed: {}",
        data.name
    );
    ensure!(
        number_at(data, "A1") == Some(12.0),
        "converted CFB/.et input lost A1: {:?}",
        data.values.get("A1")
    );
    ensure!(
        number_at(data, "A2") == Some(3.0),
        "converted CFB/.et input lost A2: {:?}",
        data.values.get("A2")
    );
    ensure!(
        data.formulas.get("B1").map(String::as_str) == Some("SUM(A1:A2)"),
        "converted CFB/.et input lost B1 formula: {:?}",
        data.formulas.get("B1")
    );
    Ok(())
}

#[test]
fn com_conversion_validates_native_et_fixture_when_provided() -> Result<()> {
    if !wps_et_tests_enabled() {
        return Ok(());
    }
    let Some(fixture) = env::var_os("LO_CONVERT_NATIVE_ET_FIXTURE") else {
        eprintln!("skipping native ET acceptance: no verified native ET fixture was supplied");
        return Ok(());
    };
    let root = temp_test_dir("lo_convert_native_et")?;
    let input = copy_native_et_fixture(&root.0, Path::new(&fixture))?;
    let workbook = convert_and_read(&input, &root.0)?;
    ensure!(
        !workbook.sheets.is_empty(),
        "converted ET has no worksheets"
    );
    ensure!(
        workbook
            .sheets
            .iter()
            .any(|sheet| !sheet.values.is_empty() || !sheet.formulas.is_empty()),
        "native ET fixture must contain at least one stored value or formula"
    );
    Ok(())
}

fn convert_and_read(input: &Path, root: &Path) -> Result<WorkbookDump> {
    let output = root.join("output.xlsx");
    let source_bytes = fs::read(input)?;
    let result = Command::new(env!("CARGO_BIN_EXE_lo-convert"))
        .args([
            "convert",
            "--input",
            input.to_str().context("ET input path is not UTF-8")?,
            "--output",
            output.to_str().context("XLSX output path is not UTF-8")?,
            "--to",
            "xlsx",
            "--engine",
            "com",
        ])
        .output()?;
    ensure!(
        result.status.success(),
        "ET conversion failed: {}",
        String::from_utf8_lossy(&result.stderr)
    );
    let report: serde_json::Value = serde_json::from_slice(&result.stdout)?;
    ensure!(report["status"] == "converted");
    ensure!(report["engine"] == "com");
    ensure!(
        report["validation"]["level"] == "worksheet_order_names_values_and_formula_text",
        "conversion did not compare worksheet/data/formula semantics: {}",
        report["validation"]
    );
    ensure!(
        report["validation"]["source_probe"] == "separate_conversion_using_selected_engine",
        "validation did not separately convert the source: {}",
        report["validation"]
    );
    ensure!(output.is_file(), "validated XLSX output was not published");
    let workbook = read_workbook(&output)?;
    ensure!(
        !workbook.sheets.is_empty(),
        "converted XLSX has no worksheets"
    );
    ensure!(
        fs::read(input)? == source_bytes,
        "source input was modified during COM conversion"
    );
    Ok(workbook)
}

fn number_at(sheet: &excel_reader::SheetDump, address: &str) -> Option<f64> {
    sheet
        .values
        .get(address)
        .and_then(|value| value.get("value"))
        .and_then(serde_json::Value::as_f64)
}

fn wps_et_tests_enabled() -> bool {
    if !cfg!(windows) {
        eprintln!("skipping WPS COM tests: COM is only available on Windows");
        return false;
    }
    if env::var("LO_CONVERT_RUN_WPS_ET_TESTS").as_deref() != Ok("1") {
        eprintln!("skipping WPS COM tests: set LO_CONVERT_RUN_WPS_ET_TESTS=1 to enable");
        return false;
    }
    true
}
