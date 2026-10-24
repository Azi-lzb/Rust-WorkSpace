//! Explicit same-category conversion requests delegated to the independent
//! sidecar executable. This module does not link the sidecar crate.

use anyhow::{anyhow, bail, Context, Result};
use encoding_rs::GBK;
use excel_reader::{read_workbook, CellValue};
use excel_writer::Workbook;
use serde::Serialize;
use std::{
    collections::HashSet,
    env, fs,
    path::{Path, PathBuf},
    process::{Command, Stdio},
};

#[derive(Clone, Debug, Serialize)]
pub struct ConversionItem {
    pub input: PathBuf,
    pub output: Option<PathBuf>,
    pub status: String,
    pub detail: String,
}
#[derive(Clone, Debug, Serialize)]
pub struct ConversionRun {
    pub converted: usize,
    pub skipped: usize,
    pub failed: usize,
    pub items: Vec<ConversionItem>,
}

pub fn convert_excel_files(
    inputs: &[PathBuf],
    output_dir: &Path,
    engine: &str,
) -> Result<ConversionRun> {
    convert_excel_files_to(inputs, output_dir, "xlsx", engine)
}
pub fn convert_excel_files_to(
    inputs: &[PathBuf],
    output_dir: &Path,
    target: &str,
    engine: &str,
) -> Result<ConversionRun> {
    anyhow::ensure!(
        target
            .trim()
            .trim_start_matches('.')
            .eq_ignore_ascii_case("xlsx"),
        "Excel conversion currently targets .xlsx only"
    );
    convert_files(
        inputs,
        Some(output_dir),
        target,
        engine,
        ConversionCategory::Excel,
    )
}
pub fn convert_excel_files_in_source_folders(
    inputs: &[PathBuf],
    target: &str,
    engine: &str,
) -> Result<ConversionRun> {
    convert_files(inputs, None, target, engine, ConversionCategory::Excel)
}
pub fn convert_word_files(
    inputs: &[PathBuf],
    output_dir: &Path,
    engine: &str,
) -> Result<ConversionRun> {
    convert_word_files_to(inputs, output_dir, "docx", engine)
}
pub fn convert_word_files_to(
    inputs: &[PathBuf],
    output_dir: &Path,
    target: &str,
    engine: &str,
) -> Result<ConversionRun> {
    anyhow::ensure!(
        target
            .trim()
            .trim_start_matches('.')
            .eq_ignore_ascii_case("docx"),
        "Word conversion currently targets .docx only"
    );
    convert_files(
        inputs,
        Some(output_dir),
        target,
        engine,
        ConversionCategory::Word,
    )
}
pub fn convert_word_files_in_source_folders(
    inputs: &[PathBuf],
    target: &str,
    engine: &str,
) -> Result<ConversionRun> {
    convert_files(inputs, None, target, engine, ConversionCategory::Word)
}
pub fn convert_one(input: &Path, output: &Path, target: &str, engine: &str) -> Result<String> {
    anyhow::ensure!(
        input.is_file(),
        "input file does not exist: {}",
        input.display()
    );
    anyhow::ensure!(
        !output.exists(),
        "refusing to overwrite existing file: {}",
        output.display()
    );
    let target = target.trim().trim_start_matches('.').to_ascii_lowercase();
    let source_ext = input
        .extension()
        .and_then(|ext| ext.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();
    anyhow::ensure!(
        is_same_category(&source_ext, &target),
        "cross-category conversion is not supported: .{source_ext} -> .{target}"
    );
    anyhow::ensure!(
        matches!(target.as_str(), "xlsx" | "docx"),
        "only .xlsx and .docx conversion targets are supported"
    );
    anyhow::ensure!(
        source_ext != target,
        "input is already in the requested format: .{target}"
    );
    anyhow::ensure!(
        matches!(engine, "auto" | "lo" | "com" | "native"),
        "--engine must be auto, lo, com or native"
    );
    if engine == "native" {
        anyhow::ensure!(
            target == "xlsx",
            "native value-only conversion is currently available only for Excel to .xlsx"
        );
        if source_ext == "csv" {
            return convert_native_csv_values(input, output);
        }
        return convert_native_excel_values(input, output, &source_ext);
    }
    let sidecar = find_sidecar()?;
    invoke_sidecar(&sidecar, input, output, &target, engine)
}

fn invoke_sidecar(
    sidecar: &Path,
    input: &Path,
    output: &Path,
    target: &str,
    engine: &str,
) -> Result<String> {
    let result = Command::new(&sidecar)
        .arg("convert")
        .arg("--input")
        .arg(input)
        .arg("--output")
        .arg(output)
        .arg("--to")
        .arg(&target)
        .arg("--mode")
        .arg("fidelity")
        .arg("--engine")
        .arg(engine)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .output();
    let result = match result {
        Ok(result) => result,
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => {
            bail!(
                "该功能需要外部引擎: conversion sidecar is unavailable ({})",
                sidecar.display()
            );
        }
        Err(error) => {
            return Err(error)
                .with_context(|| format!("start conversion sidecar {}", sidecar.display()));
        }
    };
    if !result.status.success() {
        let message = String::from_utf8_lossy(&result.stderr).trim().to_owned();
        if message.contains("unavailable")
            || message.contains("install soffice")
            || message.contains("COM conversion is only available")
            || message.contains("No supported Office/WPS COM engine is available")
        {
            bail!("该功能需要外部引擎: {message}");
        }
        bail!("conversion failed: {message}");
    }
    anyhow::ensure!(
        output.is_file() && fs::metadata(output)?.len() > 0,
        "conversion sidecar reported success but output is missing or empty"
    );
    Ok(String::from_utf8_lossy(&result.stdout).trim().to_owned())
}

#[derive(Clone, Copy)]
enum ConversionCategory {
    Excel,
    Word,
}

fn convert_files(
    inputs: &[PathBuf],
    output_dir: Option<&Path>,
    target: &str,
    engine: &str,
    category: ConversionCategory,
) -> Result<ConversionRun> {
    let target = target.trim().trim_start_matches('.').to_ascii_lowercase();
    anyhow::ensure!(
        matches!(target.as_str(), "xlsx" | "docx"),
        "conversion target must be xlsx or docx"
    );
    if let Some(output_dir) = output_dir {
        fs::create_dir_all(output_dir)?;
    }
    let mut run = ConversionRun {
        converted: 0,
        skipped: 0,
        failed: 0,
        items: Vec::new(),
    };
    let inputs = expand_conversion_inputs(inputs, category)?;
    for input in &inputs {
        if !input.is_file() {
            run.failed += 1;
            run.items.push(ConversionItem {
                input: input.clone(),
                output: None,
                status: "failed".into(),
                detail: "input file does not exist".into(),
            });
            continue;
        }
        let ext = input
            .extension()
            .and_then(|x| x.to_str())
            .unwrap_or_default()
            .to_ascii_lowercase();
        let category_ok = match category {
            ConversionCategory::Excel => {
                matches!(
                    ext.as_str(),
                    "csv" | "et" | "xls" | "xlsx" | "xlsm" | "xlsb"
                )
            }
            ConversionCategory::Word => matches!(ext.as_str(), "doc" | "docx"),
        };
        if !category_ok {
            run.failed += 1;
            run.items.push(ConversionItem {
                input: input.clone(),
                output: None,
                status: "failed".into(),
                detail: format!("unsupported same-category conversion source: .{ext}"),
            });
            continue;
        }
        let stem = input
            .file_stem()
            .map(|s| s.to_string_lossy().into_owned())
            .unwrap_or_else(|| "converted".into());
        let file_output_dir = if let Some(output_dir) = output_dir {
            output_dir.to_path_buf()
        } else {
            let parent = input
                .parent()
                .filter(|path| !path.as_os_str().is_empty())
                .unwrap_or_else(|| Path::new("."));
            parent.join(&target)
        };
        fs::create_dir_all(&file_output_dir)?;
        let mut output = file_output_dir.join(format!("{stem}.{target}"));
        let mut i = 1;
        while output.exists() {
            output = file_output_dir.join(format!("{stem}_{i}.{target}"));
            i += 1;
        }
        let result = if ext == target {
            fs::copy(input, &output)
                .map(|_| "same-format copy; source bytes retained".to_owned())
                .map_err(anyhow::Error::from)
        } else {
            convert_one(input, &output, &target, engine)
        };
        match result {
            Ok(detail) => {
                run.converted += 1;
                run.items.push(ConversionItem {
                    input: input.clone(),
                    output: Some(output),
                    status: if ext == target {
                        "copied".into()
                    } else {
                        "converted".into()
                    },
                    detail,
                });
            }
            Err(error) => {
                run.failed += 1;
                run.items.push(ConversionItem {
                    input: input.clone(),
                    output: None,
                    status: "failed".into(),
                    detail: error.to_string(),
                });
            }
        }
    }
    Ok(run)
}

fn expand_conversion_inputs(
    inputs: &[PathBuf],
    category: ConversionCategory,
) -> Result<Vec<PathBuf>> {
    let mut files = Vec::new();
    let mut seen = HashSet::new();
    for input in inputs {
        if input.is_dir() {
            collect_conversion_directory(input, category, &mut files, &mut seen)?;
        } else if !is_lock_file(input) {
            remember_conversion_file(input, &mut files, &mut seen);
        }
    }
    Ok(files)
}

fn collect_conversion_directory(
    directory: &Path,
    category: ConversionCategory,
    files: &mut Vec<PathBuf>,
    seen: &mut HashSet<String>,
) -> Result<()> {
    let mut entries = fs::read_dir(directory)
        .with_context(|| format!("list conversion input directory {}", directory.display()))?
        .map(|entry| entry.map(|entry| entry.path()))
        .collect::<std::io::Result<Vec<_>>>()?;
    entries.sort_by_key(|path| {
        path.strip_prefix(directory)
            .unwrap_or(path)
            .to_string_lossy()
            .to_lowercase()
    });
    for path in entries {
        let file_type = fs::symlink_metadata(&path)
            .with_context(|| format!("inspect conversion input {}", path.display()))?
            .file_type();
        if file_type.is_dir() {
            collect_conversion_directory(&path, category, files, seen)?;
        } else if file_type.is_file()
            && !is_lock_file(&path)
            && conversion_source_matches_category(&path, category)
        {
            remember_conversion_file(&path, files, seen);
        }
    }
    Ok(())
}

fn remember_conversion_file(path: &Path, files: &mut Vec<PathBuf>, seen: &mut HashSet<String>) {
    let identity = path
        .canonicalize()
        .unwrap_or_else(|_| path.to_path_buf())
        .to_string_lossy()
        .to_lowercase();
    if seen.insert(identity) {
        files.push(path.to_path_buf());
    }
}

fn conversion_source_matches_category(path: &Path, category: ConversionCategory) -> bool {
    let extension = path
        .extension()
        .and_then(|extension| extension.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();
    match category {
        ConversionCategory::Excel => {
            matches!(
                extension.as_str(),
                "csv" | "et" | "xls" | "xlsx" | "xlsm" | "xlsb"
            )
        }
        ConversionCategory::Word => matches!(extension.as_str(), "doc" | "docx"),
    }
}

fn is_lock_file(path: &Path) -> bool {
    path.file_name()
        .and_then(|name| name.to_str())
        .is_some_and(|name| name.starts_with("~$"))
}

fn is_same_category(source: &str, target: &str) -> bool {
    match target {
        "xlsx" => matches!(source, "csv" | "et" | "xls" | "xlsx" | "xlsm" | "xlsb"),
        "docx" => matches!(source, "doc" | "docx"),
        _ => false,
    }
}

fn convert_native_csv_values(input: &Path, output: &Path) -> Result<String> {
    const MAX_ROWS: usize = 1_048_576;
    const MAX_COLUMNS: usize = 16_384;
    anyhow::ensure!(
        input.is_file(),
        "input file does not exist: {}",
        input.display()
    );
    anyhow::ensure!(
        !output.exists(),
        "refusing to overwrite existing file: {}",
        output.display()
    );
    let bytes = fs::read(input).with_context(|| format!("read CSV {}", input.display()))?;
    let decoded = decode_csv_text(&bytes)?;
    let mut reader = csv::ReaderBuilder::new()
        .has_headers(false)
        .flexible(true)
        .from_reader(decoded.as_bytes());
    let mut rows = Vec::<Vec<String>>::new();
    for record in reader.records() {
        let record = record.with_context(|| format!("parse CSV {}", input.display()))?;
        anyhow::ensure!(
            rows.len() < MAX_ROWS,
            "CSV exceeds Excel's 1,048,576-row limit"
        );
        anyhow::ensure!(
            record.len() <= MAX_COLUMNS,
            "CSV row {} exceeds Excel's 16,384-column limit",
            rows.len() + 1
        );
        rows.push(record.iter().map(str::to_owned).collect());
    }

    let parent = output
        .parent()
        .filter(|path| !path.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    fs::create_dir_all(parent)?;
    let temp_dir = unique_native_temp_dir(parent)?;
    let _guard = TempDirGuard(temp_dir.clone());
    let candidate = temp_dir.join(
        output
            .file_name()
            .context("native CSV output must have a file name")?,
    );
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_name("Sheet1")?;
    for (row_index, row) in rows.iter().enumerate() {
        let row_index = u32::try_from(row_index).context("CSV exceeds Excel row limit")?;
        for (column_index, value) in row.iter().enumerate() {
            let column_index =
                u16::try_from(column_index).context("CSV exceeds Excel column limit")?;
            if !value.is_empty() {
                worksheet.write_string(row_index, column_index, value)?;
            }
        }
    }
    workbook.save(&candidate).with_context(|| {
        format!(
            "write native CSV conversion candidate {}",
            candidate.display()
        )
    })?;
    validate_native_csv_output(&candidate, &rows)?;
    fs::hard_link(&candidate, output).with_context(|| {
        format!(
            "commit native CSV conversion without replacing an existing file {}",
            output.display()
        )
    })?;
    Ok(format!(
        "按值重建（C01，丢弃样式/公式缓存）；native CSV-to-XLSX conversion; {} rows written as text; no Office type inference or source formatting",
        rows.len()
    ))
}

fn decode_csv_text(bytes: &[u8]) -> Result<String> {
    let bytes = bytes.strip_prefix(&[0xEF, 0xBB, 0xBF]).unwrap_or(bytes);
    if let Ok(text) = std::str::from_utf8(bytes) {
        return Ok(text.to_owned());
    }
    let (text, _, had_errors) = GBK.decode(bytes);
    anyhow::ensure!(
        !had_errors,
        "CSV is neither valid UTF-8 nor valid GBK; choose an external conversion engine or transcode it first"
    );
    Ok(text.into_owned())
}

fn validate_native_csv_output(candidate: &Path, expected: &[Vec<String>]) -> Result<()> {
    let workbook = read_workbook(candidate)
        .with_context(|| format!("reopen native CSV candidate {}", candidate.display()))?;
    anyhow::ensure!(
        workbook.sheets.len() == 1 && workbook.sheets[0].name == "Sheet1",
        "native CSV candidate must contain exactly one Sheet1 worksheet"
    );
    let sheet = &workbook.sheets[0];
    let mut expected_addresses = std::collections::BTreeSet::new();
    for (row_index, row) in expected.iter().enumerate() {
        for (column_index, value) in row.iter().enumerate() {
            if value.is_empty() {
                continue;
            }
            let address = format!("{}{}", csv_column_name(column_index + 1), row_index + 1);
            expected_addresses.insert(address.clone());
            let actual = sheet
                .values
                .get(&address)
                .with_context(|| format!("native CSV candidate is missing {address}"))?;
            anyhow::ensure!(
                actual.get("type").and_then(serde_json::Value::as_str) == Some("string")
                    && actual.get("value").and_then(serde_json::Value::as_str)
                        == Some(value.as_str()),
                "native CSV value differs at {address}"
            );
        }
    }
    let actual_addresses = sheet
        .values
        .keys()
        .cloned()
        .collect::<std::collections::BTreeSet<_>>();
    anyhow::ensure!(
        actual_addresses == expected_addresses,
        "native CSV candidate contains missing or extra non-empty cells"
    );
    Ok(())
}

fn csv_column_name(mut column: usize) -> String {
    let mut name = String::new();
    while column > 0 {
        let remainder = (column - 1) % 26;
        name.insert(0, (b'A' + remainder as u8) as char);
        column = (column - 1) / 26;
    }
    name
}

fn convert_native_excel_values(input: &Path, output: &Path, source_ext: &str) -> Result<String> {
    anyhow::ensure!(
        matches!(source_ext, "xls" | "xlsm" | "xlsb"),
        "native value-only conversion cannot read .{source_ext}; use an external fidelity engine"
    );
    anyhow::ensure!(
        !output.exists(),
        "refusing to overwrite existing file: {}",
        output.display()
    );

    let source = read_workbook(input).with_context(|| {
        format!(
            "read source workbook for native conversion {}",
            input.display()
        )
    })?;
    validate_native_source_values(&source)?;

    let parent = output
        .parent()
        .filter(|path| !path.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    fs::create_dir_all(parent)?;
    let temp_dir = unique_native_temp_dir(parent)?;
    let guard = TempDirGuard(temp_dir.clone());
    let candidate = temp_dir.join(
        output
            .file_name()
            .context("native conversion output must have a file name")?,
    );

    let mut workbook = Workbook::new();
    for sheet in &source.sheets {
        let worksheet = workbook.add_worksheet();
        worksheet
            .set_name(&sheet.name)
            .with_context(|| format!("create worksheet {:?}", sheet.name))?;
        for (address, formula) in &sheet.formulas {
            let (row, column) = parse_a1(address).with_context(|| {
                format!("invalid formula coordinate {}!{}", sheet.name, address)
            })?;
            let formula = if formula.starts_with('=') {
                formula.clone()
            } else {
                format!("={formula}")
            };
            worksheet
                .write_formula(row, column, formula.as_str())
                .with_context(|| format!("write formula {}!{}", sheet.name, address))?;
        }
        for (address, value) in &sheet.values {
            if sheet.formulas.contains_key(address) {
                continue;
            }
            let (row, column) = parse_a1(address)
                .with_context(|| format!("invalid value coordinate {}!{}", sheet.name, address))?;
            write_native_cell(worksheet, row, column, value)
                .with_context(|| format!("write value {}!{}", sheet.name, address))?;
        }
    }
    workbook
        .save(&candidate)
        .with_context(|| format!("write native conversion candidate {}", candidate.display()))?;

    let converted = read_workbook(&candidate)
        .with_context(|| format!("reopen native conversion candidate {}", candidate.display()))?;
    validate_native_conversion(&source, &converted)?;
    fs::hard_link(&candidate, output).with_context(|| {
        format!(
            "publish validated native conversion without replacing {}",
            output.display()
        )
    })?;
    drop(guard);

    Ok("按值重建（C01，丢弃样式/公式缓存）；native value-only conversion validated: worksheet names/order, stored values and formula text preserved; original formula caches are not preserved; formatting, comments, merges, dimensions, names, page settings, validations, conditional formatting, drawings and macros are not preserved; formulas are not recalculated".to_owned())
}

fn validate_native_source_values(source: &excel_reader::WorkbookDump) -> Result<()> {
    for sheet in &source.sheets {
        for (address, value) in &sheet.values {
            if sheet.formulas.contains_key(address) {
                continue;
            }
            let parsed = CellValue::from_probe_json(value)
                .with_context(|| format!("decode {}!{}", sheet.name, address))?;
            validate_native_cell_value(&sheet.name, address, parsed)?;
        }
    }
    Ok(())
}

fn validate_native_cell_value(sheet_name: &str, address: &str, value: CellValue) -> Result<()> {
    match value {
        CellValue::Text(text) if text.is_empty() => bail!(
            "native value-only conversion cannot preserve an empty-text cell {}!{}; use an external fidelity engine",
            sheet_name,
            address
        ),
        CellValue::Error(_) => bail!(
            "native value-only conversion refuses error cell {}!{} because the writer cannot preserve its typed error value",
            sheet_name,
            address
        ),
        CellValue::Number(number) | CellValue::DateTimeSerial(number) if !number.is_finite() => {
            bail!(
                "native value-only conversion refuses non-finite numeric cell {}!{}",
                sheet_name,
                address
            )
        }
        CellValue::Integer(number) if number.unsigned_abs() > (1_u64 << 53) => {
            bail!(
                "native value-only conversion refuses integer cell {}!{} outside the exact XLSX numeric range",
                sheet_name,
                address
            )
        }
        _ => {}
    }
    Ok(())
}

fn write_native_cell(
    worksheet: &mut excel_writer::Worksheet,
    row: u32,
    column: u16,
    value: &serde_json::Value,
) -> Result<()> {
    match CellValue::from_probe_json(value)? {
        CellValue::Empty => Ok(()),
        CellValue::Text(text) | CellValue::DateTimeIso(text) | CellValue::DurationIso(text) => {
            worksheet.write_string(row, column, text)?;
            Ok(())
        }
        CellValue::Number(number) | CellValue::DateTimeSerial(number) => {
            worksheet.write_number(row, column, number)?;
            Ok(())
        }
        CellValue::Integer(number) => {
            worksheet.write_number(row, column, number as f64)?;
            Ok(())
        }
        CellValue::Boolean(value) => {
            worksheet.write_boolean(row, column, value)?;
            Ok(())
        }
        CellValue::Error(error) => {
            bail!("typed Excel error cells cannot be written natively: {error}")
        }
    }
}

fn validate_native_conversion(
    source: &excel_reader::WorkbookDump,
    converted: &excel_reader::WorkbookDump,
) -> Result<()> {
    let source_names = source
        .sheets
        .iter()
        .map(|sheet| &sheet.name)
        .collect::<Vec<_>>();
    let converted_names = converted
        .sheets
        .iter()
        .map(|sheet| &sheet.name)
        .collect::<Vec<_>>();
    anyhow::ensure!(
        source_names == converted_names,
        "native conversion changed worksheet count, names or order"
    );
    for (source_sheet, output_sheet) in source.sheets.iter().zip(&converted.sheets) {
        anyhow::ensure!(
            source_sheet.formulas.len() == output_sheet.formulas.len(),
            "native conversion changed formula count in worksheet {:?}",
            source_sheet.name
        );
        for (address, formula) in &source_sheet.formulas {
            let output_formula = output_sheet.formulas.get(address).with_context(|| {
                format!(
                    "native conversion removed formula {}!{}",
                    source_sheet.name, address
                )
            })?;
            anyhow::ensure!(
                formula.trim_start_matches('=').trim()
                    == output_formula.trim_start_matches('=').trim(),
                "native conversion changed formula text at {}!{}",
                source_sheet.name,
                address
            );
        }
        let source_cells = source_sheet
            .values
            .keys()
            .filter(|address| !source_sheet.formulas.contains_key(*address))
            .collect::<std::collections::BTreeSet<_>>();
        let output_cells = output_sheet
            .values
            .keys()
            .filter(|address| !output_sheet.formulas.contains_key(*address))
            .collect::<std::collections::BTreeSet<_>>();
        if source_cells != output_cells {
            let missing_addresses = source_cells
                .difference(&output_cells)
                .copied()
                .collect::<Vec<_>>();
            let unexpected_count = output_cells.difference(&source_cells).count();
            let missing = missing_addresses
                .iter()
                .take(12)
                .map(|address| address.as_str())
                .collect::<Vec<_>>();
            let unexpected = output_cells
                .difference(&source_cells)
                .take(12)
                .map(|address| address.as_str())
                .collect::<Vec<_>>();
            let mut missing_types = std::collections::BTreeMap::<&str, usize>::new();
            let mut missing_empty_strings = 0usize;
            for address in &missing_addresses {
                if let Some(value) = source_sheet.values.get(*address) {
                    let value_type = value
                        .get("type")
                        .and_then(serde_json::Value::as_str)
                        .unwrap_or("unknown");
                    *missing_types.entry(value_type).or_default() += 1;
                    if value_type == "string"
                        && value.get("value").and_then(serde_json::Value::as_str) == Some("")
                    {
                        missing_empty_strings += 1;
                    }
                }
            }
            bail!(
                "native conversion changed data cell coordinates in worksheet {:?} (source={}, output={}, missing_in_output={} {:?}, missing_value_types={:?}, missing_empty_strings={}, unexpected_in_output={} {:?})",
                source_sheet.name,
                source_cells.len(),
                output_cells.len(),
                missing_addresses.len(),
                missing,
                missing_types,
                missing_empty_strings,
                unexpected_count,
                unexpected
            );
        }
        for address in source_cells {
            anyhow::ensure!(
                native_values_equal(&source_sheet.values[address], &output_sheet.values[address])?,
                "native conversion changed stored value at {}!{}",
                source_sheet.name,
                address
            );
        }
    }
    Ok(())
}

fn native_values_equal(left: &serde_json::Value, right: &serde_json::Value) -> Result<bool> {
    let left = CellValue::from_probe_json(left)?;
    let right = CellValue::from_probe_json(right)?;
    Ok(match (left, right) {
        (CellValue::Text(left), CellValue::Text(right))
        | (CellValue::DateTimeIso(left), CellValue::DateTimeIso(right))
        | (CellValue::DurationIso(left), CellValue::DurationIso(right)) => left == right,
        (CellValue::Text(left), CellValue::DateTimeIso(right))
        | (CellValue::DateTimeIso(left), CellValue::Text(right))
        | (CellValue::Text(left), CellValue::DurationIso(right))
        | (CellValue::DurationIso(left), CellValue::Text(right)) => left == right,
        (CellValue::Boolean(left), CellValue::Boolean(right)) => left == right,
        (CellValue::Integer(left), CellValue::Integer(right)) => left == right,
        (CellValue::Integer(left), CellValue::Number(right))
        | (CellValue::Number(right), CellValue::Integer(left))
        | (CellValue::Integer(left), CellValue::DateTimeSerial(right))
        | (CellValue::DateTimeSerial(right), CellValue::Integer(left)) => {
            numeric_values_equal(left as f64, right)
        }
        (CellValue::Number(left), CellValue::Number(right))
        | (CellValue::Number(left), CellValue::DateTimeSerial(right))
        | (CellValue::DateTimeSerial(left), CellValue::Number(right))
        | (CellValue::DateTimeSerial(left), CellValue::DateTimeSerial(right)) => {
            numeric_values_equal(left, right)
        }
        (CellValue::Empty, CellValue::Empty) => true,
        _ => false,
    })
}

fn numeric_values_equal(left: f64, right: f64) -> bool {
    if left == right {
        return true;
    }
    left.is_finite()
        && right.is_finite()
        && (left - right).abs() <= 1e-15 * left.abs().max(right.abs())
}

fn parse_a1(address: &str) -> Option<(u32, u16)> {
    let split = address.find(|character: char| character.is_ascii_digit())?;
    let mut column = 0_u32;
    for byte in address[..split].bytes() {
        if !byte.is_ascii_alphabetic() {
            return None;
        }
        column = column.checked_mul(26)?;
        column = column.checked_add((byte.to_ascii_uppercase() - b'A' + 1) as u32)?;
    }
    let row = address[split..].parse::<u32>().ok()?;
    if row == 0 || column == 0 || column > 16_384 {
        return None;
    }
    Some((row - 1, (column - 1) as u16))
}

fn unique_native_temp_dir(parent: &Path) -> Result<PathBuf> {
    use std::time::{SystemTime, UNIX_EPOCH};
    let nanos = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
    for index in 0..100 {
        let path = parent.join(format!(
            ".pytools-native-convert-{}-{nanos}-{index}",
            std::process::id()
        ));
        match fs::create_dir(&path) {
            Ok(()) => return Ok(path),
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(error) => return Err(error.into()),
        }
    }
    bail!("cannot allocate native conversion temporary directory")
}

struct TempDirGuard(PathBuf);

impl Drop for TempDirGuard {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

fn find_sidecar() -> Result<PathBuf> {
    if let Some(configured) = env::var_os("PYTOOLS_LO_CONVERT") {
        let path = PathBuf::from(configured);
        if path.is_file() {
            return Ok(path);
        }
        return Err(anyhow!(
            "该功能需要外部引擎: PYTOOLS_LO_CONVERT does not point to a sidecar executable: {}",
            path.display()
        ));
    }
    let executable = env::current_exe().context("resolve pytools executable path")?;
    let suffix = if cfg!(windows) {
        "lo-convert.exe"
    } else {
        "lo-convert"
    };
    if let Some(parent) = executable.parent() {
        let sibling = parent.join(suffix);
        if sibling.is_file() {
            return Ok(sibling);
        }
    }
    Ok(PathBuf::from(suffix))
}

#[cfg(test)]
mod tests {
    use super::{
        convert_excel_files_in_source_folders, convert_one, invoke_sidecar,
        validate_native_source_values,
    };
    use excel_reader::{read_workbook, SheetDump, WorkbookDump};
    use excel_writer::Workbook;
    use serde_json::json;
    use std::{collections::BTreeMap, fs, path::PathBuf, time::SystemTime};

    struct TempDir(PathBuf);

    impl TempDir {
        fn new(label: &str) -> Self {
            let stamp = SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos();
            let path = std::env::temp_dir().join(format!(
                "pytools-conversion-{label}-{}-{stamp}",
                std::process::id()
            ));
            fs::create_dir_all(&path).unwrap();
            Self(path)
        }
    }

    impl Drop for TempDir {
        fn drop(&mut self) {
            let _ = fs::remove_dir_all(&self.0);
        }
    }

    #[test]
    fn missing_sidecar_reports_external_engine_requirement() {
        let root = TempDir::new("missing-sidecar");
        let error = invoke_sidecar(
            &root.0.join("lo-convert-missing"),
            &root.0.join("input.xls"),
            &root.0.join("output.xlsx"),
            "xlsx",
            "lo",
        )
        .expect_err("missing converter executable must be reported");
        assert!(error.to_string().contains("该功能需要外部引擎"));
    }

    #[test]
    fn explicit_native_conversion_preserves_sheet_data_and_formula_text() {
        let root = TempDir::new("native-values");
        let source_xlsx = root.0.join("source.xlsx");
        let source_xlsm = root.0.join("source.xlsm");
        let output = root.0.join("converted.xlsx");

        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.set_name("数据").unwrap();
        sheet.write_string(0, 0, "代码").unwrap();
        sheet.write_number(0, 1, 123.25).unwrap();
        sheet.write_boolean(0, 2, true).unwrap();
        sheet.write_formula(0, 3, "=B1*2").unwrap();
        workbook.add_worksheet().set_name("空工作表").unwrap();
        workbook.save(&source_xlsx).unwrap();
        fs::copy(&source_xlsx, &source_xlsm).unwrap();
        let original_bytes = fs::read(&source_xlsm).unwrap();

        let detail = convert_one(&source_xlsm, &output, "xlsx", "native").unwrap();
        assert!(detail.contains("value-only conversion validated"));
        assert!(detail.contains("macros are not preserved"));
        assert_eq!(fs::read(&source_xlsm).unwrap(), original_bytes);

        let converted = read_workbook(&output).unwrap();
        assert_eq!(
            converted
                .sheets
                .iter()
                .map(|sheet| sheet.name.as_str())
                .collect::<Vec<_>>(),
            vec!["数据", "空工作表"]
        );
        assert_eq!(converted.sheets[0].values["A1"]["value"], "代码");
        assert_eq!(converted.sheets[0].values["B1"]["value"], 123.25);
        assert_eq!(converted.sheets[0].values["C1"]["value"], true);
        assert_eq!(
            converted.sheets[0].formulas["D1"].trim_start_matches('='),
            "B1*2"
        );
    }

    #[test]
    fn native_value_conversion_rejects_empty_text_before_writing() {
        let source = workbook_with_value_and_formula("", None);
        let error = validate_native_source_values(&source)
            .expect_err("empty-text cells cannot be represented by the native writer");
        let message = format!("{error:#}");
        assert!(message.contains("empty-text cell 数据!A1"));
        assert!(message.contains("external fidelity engine"));
        assert!(!message.contains("value="));
    }

    #[test]
    fn native_value_conversion_skips_empty_formula_cache_during_preflight() {
        let source = workbook_with_value_and_formula("", Some("=\"\""));
        validate_native_source_values(&source)
            .expect("empty cached formula values are represented by preserved formula text");
    }

    #[test]
    fn native_value_conversion_refuses_et_without_external_engine() {
        let root = TempDir::new("native-et");
        let input = root.0.join("input.et");
        let output = root.0.join("output.xlsx");
        fs::write(&input, b"not an OOXML workbook").unwrap();
        let error = convert_one(&input, &output, "xlsx", "native")
            .expect_err("native value-only reader must reject ET");
        assert!(error.to_string().contains("cannot read .et"));
        assert!(!output.exists());
    }

    fn workbook_with_value_and_formula(value: &str, formula: Option<&str>) -> WorkbookDump {
        let values = BTreeMap::from([("A1".to_owned(), json!({"type":"string", "value":value}))]);
        let formulas = formula
            .map(|formula| BTreeMap::from([("A1".to_owned(), formula.to_owned())]))
            .unwrap_or_default();
        WorkbookDump {
            schema_version: 1,
            reader: "unit-test",
            elapsed_ms: 0,
            sheets: vec![SheetDump {
                name: "数据".to_owned(),
                values,
                formulas,
                formula_status: "available".to_owned(),
                merges: Vec::new(),
                merge_status: "available".to_owned(),
                comments: BTreeMap::new(),
                comments_status: "available".to_owned(),
                conditional_formatting: Vec::new(),
                conditional_formatting_status: "available".to_owned(),
                data_validations: Vec::new(),
                data_validations_status: "available".to_owned(),
            }],
        }
    }

    #[test]
    fn batch_conversion_matches_source_format_folder_and_copies_same_format() {
        let root = TempDir::new("source-folder-layout");
        let input = root.0.join("report.xlsx");
        let output = root.0.join("xlsx/report.xlsx");
        let mut workbook = Workbook::new();
        workbook
            .add_worksheet()
            .write_string(0, 0, "preserve bytes")
            .unwrap();
        workbook.save(&input).unwrap();
        let source_bytes = fs::read(&input).unwrap();

        let result = convert_excel_files_in_source_folders(&[input.clone()], "xlsx", "lo").unwrap();
        assert_eq!(result.converted, 1);
        assert_eq!(result.skipped, 0);
        assert_eq!(result.failed, 0);
        assert_eq!(result.items[0].status, "copied");
        assert_eq!(result.items[0].output.as_deref(), Some(output.as_path()));
        assert_eq!(fs::read(&output).unwrap(), source_bytes);
    }
}
