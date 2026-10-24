use excel_reader::{read_value_columns_with_iso_dates, CellValue};
use pytools_core::config::ExtractTask;
use pytools_core::sheet_extract::run_extract_sheets_area_candidate;
use serde_json::{json, Value};
use std::fs;
use std::io::Write;
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use std::time::{SystemTime, UNIX_EPOCH};

fn cases() -> Value {
    serde_json::from_str(include_str!("fixtures/sheet_extract_area/cases.json"))
        .expect("valid worksheet extraction operation cases")
}

fn baseline() -> Value {
    serde_json::from_str(include_str!(
        "../probe/reports/sheet_extract_area_python_baseline.json"
    ))
    .expect("valid Python worksheet extraction baseline")
}

fn load_tasks(cases: &Value) -> Vec<ExtractTask> {
    cases["tasks"]
        .as_array()
        .unwrap()
        .iter()
        .map(|task| ExtractTask {
            row_no: task["row_no"].as_u64().unwrap() as usize,
            workbook_keywords: task["workbook_keywords"].as_str().unwrap().to_owned(),
            sheet_keywords: task["sheet_keywords"].as_str().unwrap().to_owned(),
            exclude_keywords: task["exclude_keywords"].as_str().unwrap().to_owned(),
            area_spec: task["area_spec"].as_str().unwrap().to_owned(),
            stitch_mode: task["stitch_mode"].as_str().unwrap().to_owned(),
            output_folder: task["output_folder"].as_str().unwrap().to_owned(),
            output_name: task["output_name"].as_str().unwrap().to_owned(),
            remark: task["remark"].as_str().unwrap().to_owned(),
        })
        .collect()
}

fn fixture_path(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/sheet_extract_area")
        .join(name)
}

fn temporary_directory(label: &str) -> PathBuf {
    let unique = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap()
        .as_nanos();
    let path =
        std::env::temp_dir().join(format!("pytools-{label}-{}-{unique}", std::process::id()));
    fs::create_dir_all(&path).unwrap();
    path
}

fn create_menu_config(root: &Path, tasks: &[ExtractTask]) -> PathBuf {
    let app = root.join("app");
    let config = app.join("config/config.xlsx");
    fs::create_dir_all(config.parent().unwrap()).unwrap();
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let global = workbook.add_worksheet();
    global.set_name("全局配置").unwrap();
    global.write_string(0, 0, "配置项").unwrap();
    global.write_string(0, 1, "值").unwrap();
    for (row, (key, value)) in [
        ("输入目录", app.join("input")),
        ("输出目录", app.join("output")),
        ("日志目录", app.join("logs")),
    ]
    .iter()
    .enumerate()
    {
        global.write_string((row + 1) as u32, 0, *key).unwrap();
        global
            .write_string((row + 1) as u32, 1, value.to_string_lossy())
            .unwrap();
    }

    let extraction = workbook.add_worksheet();
    extraction.set_name("工作表提取").unwrap();
    let headers = [
        "是否启用",
        "工作簿关键字",
        "工作表关键字",
        "排除关键字",
        "提取区域",
        "多区域拼接方式",
        "输出文件夹",
        "文件名",
        "备注",
    ];
    for (column, header) in headers.iter().enumerate() {
        extraction.write_string(0, column as u16, *header).unwrap();
    }
    for (index, task) in tasks.iter().enumerate() {
        let row = (index + 1) as u32;
        for (column, value) in [
            "是",
            task.workbook_keywords.as_str(),
            task.sheet_keywords.as_str(),
            task.exclude_keywords.as_str(),
            task.area_spec.as_str(),
            task.stitch_mode.as_str(),
            task.output_folder.as_str(),
            task.output_name.as_str(),
            task.remark.as_str(),
        ]
        .iter()
        .enumerate()
        {
            extraction.write_string(row, column as u16, *value).unwrap();
        }
    }
    workbook.save(&config).unwrap();
    config
}

fn run_menu(config: &Path, input: &str) -> std::process::Output {
    run_menu_with_view(config, input, "developer")
}

fn run_menu_with_view(config: &Path, input: &str, view: &str) -> std::process::Output {
    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("menu")
        .arg("--config")
        .arg(config)
        .arg("--log-view")
        .arg(view)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .unwrap();
    child
        .stdin
        .take()
        .unwrap()
        .write_all(input.as_bytes())
        .unwrap();
    child.wait_with_output().unwrap()
}

fn excel_cell(row: usize, column: usize) -> String {
    let mut column = column;
    let mut letters = String::new();
    while column > 0 {
        column -= 1;
        letters.insert(0, (b'A' + (column % 26) as u8) as char);
        column /= 26;
    }
    format!("{letters}{row}")
}

fn close_enough(left: f64, right: f64) -> bool {
    if left == right {
        return true;
    }
    let scale = left.abs().max(right.abs());
    scale > 0.0 && (left - right).abs() <= 1e-15 * scale
}

#[test]
fn floating_point_comparison_uses_relative_tolerance() {
    let value = 0.125_f64;
    assert!(close_enough(value, f64::from_bits(value.to_bits() + 1)));
    assert!(!close_enough(value, f64::from_bits(value.to_bits() + 8)));
}

fn assert_cell(expected: &Value, actual: Option<&CellValue>, address: &str, sheet: &str) {
    let expected_type = expected["type"].as_str();
    match (expected_type, actual) {
        (None, None) => {}
        (None, Some(CellValue::Empty)) => {}
        (Some("string"), None) if expected["value"].as_str() == Some("") => {}
        (Some("string"), Some(CellValue::Text(actual))) => {
            assert_eq!(
                actual,
                expected["value"].as_str().unwrap(),
                "{sheet}!{address}"
            );
        }
        (Some("integer" | "number"), Some(CellValue::Integer(actual))) => {
            let expected = expected["value"].as_f64().unwrap();
            assert!(
                close_enough(expected, *actual as f64),
                "{sheet}!{address}: expected {expected}, got {actual}"
            );
        }
        (Some("integer" | "number"), Some(CellValue::Number(actual))) => {
            let expected = expected["value"].as_f64().unwrap();
            assert!(
                close_enough(expected, *actual),
                "{sheet}!{address}: expected {expected}, got {actual}"
            );
        }
        (Some("boolean"), Some(CellValue::Boolean(actual))) => {
            assert_eq!(
                *actual,
                expected["value"].as_bool().unwrap(),
                "{sheet}!{address}"
            );
        }
        (Some("string"), None) if expected["value"].as_str().unwrap().trim().is_empty() => {}
        (kind, value) => panic!("{sheet}!{address}: unexpected {kind:?} / {value:?}"),
    }
}

#[test]
fn area_extraction_matches_python_baseline_and_preserves_inputs() {
    let cases = cases();
    let baseline = baseline();
    assert_eq!(baseline["status"], "passed");
    let parity_report_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("probe/reports/sheet_extract_area_parity.json");
    if parity_report_path.exists() {
        fs::remove_file(&parity_report_path).unwrap();
    }
    let source_paths = cases["input_files"]
        .as_array()
        .unwrap()
        .iter()
        .map(|name| fixture_path(name.as_str().unwrap()))
        .collect::<Vec<_>>();
    let input_bytes = source_paths
        .iter()
        .map(|path| fs::read(path).unwrap())
        .collect::<Vec<_>>();
    let root = temporary_directory("sheet-extract-area");
    let output_dir = root.join("output");
    let result = run_extract_sheets_area_candidate(
        &source_paths,
        &load_tasks(&cases),
        &output_dir,
        false,
        Some("unused-default.xlsx"),
    )
    .unwrap();

    for key in ["files", "extracted_sheets", "output_files", "failed_files"] {
        assert_eq!(
            result_value(&result, key),
            cases["expected"][key].as_u64().unwrap() as usize
        );
        assert_eq!(
            result_value(&result, key),
            baseline["stats"][key].as_u64().unwrap() as usize
        );
    }
    assert_eq!(result.paths.len(), 1);
    assert_eq!(
        Path::new(&result.paths[0])
            .file_name()
            .unwrap()
            .to_string_lossy(),
        cases["expected"]["output_file"].as_str().unwrap()
    );

    let output_sheets = read_value_columns_with_iso_dates(&result.paths[0], 16_384).unwrap();
    let expected_sheets = baseline["worksheets"].as_array().unwrap();
    assert_eq!(output_sheets.len(), expected_sheets.len());
    for (actual, expected) in output_sheets.iter().zip(expected_sheets) {
        let expected_name = expected["name"].as_str().unwrap();
        assert_eq!(actual.name, expected_name);
        assert_eq!(
            actual.max_row,
            expected["max_row"].as_u64().unwrap() as usize
        );
        assert_eq!(
            actual.max_column,
            expected["max_column"].as_u64().unwrap() as usize
        );
        for (row_index, row) in expected["values"].as_array().unwrap().iter().enumerate() {
            for (column_index, expected_cell) in row.as_array().unwrap().iter().enumerate() {
                let address = excel_cell(row_index + 1, column_index + 1);
                assert_cell(
                    expected_cell,
                    actual.values.get(&address),
                    &address,
                    expected_name,
                );
            }
        }
    }

    for (path, bytes) in source_paths.iter().zip(input_bytes) {
        assert_eq!(
            fs::read(path).unwrap(),
            bytes,
            "input changed: {}",
            path.display()
        );
    }
    let compared_cells = expected_sheets
        .iter()
        .map(|sheet| {
            sheet["values"]
                .as_array()
                .unwrap()
                .iter()
                .map(|row| row.as_array().unwrap().len())
                .sum::<usize>()
        })
        .sum::<usize>();
    let report = json!({
        "schema_version": 1,
        "status": "passed",
        "suite": "pytools.sheet_extract_area_parity",
        "baseline": "probe/reports/sheet_extract_area_python_baseline.json",
        "rust_operation": "run_extract_sheets_area_candidate",
        "stats": {
            "files": result.files,
            "extracted_sheets": result.extracted_sheets,
            "output_files": result.output_files,
            "failed_files": result.failed_files,
        },
        "worksheet_count": output_sheets.len(),
        "compared_cells": compared_cells,
        "numeric_comparison": "relative tolerance <= 1e-15",
        "checks": {
            "statistics": "preserved",
            "worksheet_names": "preserved",
            "worksheet_dimensions": "preserved",
            "cell_values_and_types": "preserved",
            "input_bytes": "preserved",
        }
    });
    let report_bytes = serde_json::to_vec_pretty(&report).unwrap();
    let mut report_with_newline = report_bytes;
    report_with_newline.push(b'\n');
    fs::write(&parity_report_path, report_with_newline).unwrap();
    fs::remove_dir_all(root).unwrap();
}

fn result_value(result: &pytools_core::sheet_extract::ExtractSheetsAreaResult, key: &str) -> usize {
    match key {
        "files" => result.files,
        "extracted_sheets" => result.extracted_sheets,
        "output_files" => result.output_files,
        "failed_files" => result.failed_files,
        _ => panic!("unknown result field {key}"),
    }
}

#[test]
fn cli_loads_config_and_runs_the_area_extraction_candidate() {
    let cases = cases();
    let root = temporary_directory("sheet-extract-cli");
    let output_dir = root.join("output");
    let mut command = Command::new(env!("CARGO_BIN_EXE_pytools"));
    command
        .arg("extract-sheets-area")
        .arg("--config")
        .arg(fixture_path("extract_config.xlsx"))
        .arg("--output-dir")
        .arg(&output_dir);
    for name in cases["input_files"].as_array().unwrap() {
        command.arg(fixture_path(name.as_str().unwrap()));
    }
    let output = command.output().unwrap();
    assert!(
        output.status.success(),
        "CLI failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report: Value = serde_json::from_slice(&output.stdout).unwrap();
    assert_eq!(report["files"], cases["expected"]["files"]);
    assert_eq!(
        report["extracted_sheets"],
        cases["expected"]["extracted_sheets"]
    );
    assert_eq!(report["output_files"], cases["expected"]["output_files"]);
    assert_eq!(report["failed_files"], cases["expected"]["failed_files"]);
    assert!(output_dir.join("区域提取/area-result.xlsx").exists());
    fs::remove_dir_all(root).unwrap();
}

#[test]
fn interactive_menu_routes_area_extraction_matches_python_and_records_engines() {
    let cases = cases();
    let root = temporary_directory("sheet-extract-menu");
    let tasks = load_tasks(&cases);
    let config = create_menu_config(&root, &tasks);
    let mut sources = cases["input_files"]
        .as_array()
        .unwrap()
        .iter()
        .map(|name| fixture_path(name.as_str().unwrap()))
        .collect::<Vec<_>>();
    let unsupported = root.join("not-validated.xls");
    fs::write(&unsupported, b"synthetic legacy workbook placeholder").unwrap();
    sources.push(unsupported.clone());
    let original_inputs = sources
        .iter()
        .map(|path| (path.clone(), fs::read(path).unwrap()))
        .collect::<Vec<_>>();
    let input = format!(
        "7\n6\n{}\n",
        sources
            .iter()
            .map(|path| path.to_string_lossy())
            .collect::<Vec<_>>()
            .join(";")
    );

    let output = run_menu(&config, &input);
    assert!(
        output.status.success(),
        "menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stdout.contains("6  工作表数据提取（普通无批注版）"));
    assert!(stdout.contains("PT-7-6 当前仅验收 .xlsx"));
    assert!(stdout.contains("\"files\": 2"));
    assert!(stdout.contains("\"extracted_sheets\": 4"));
    assert!(stdout.contains("\"output_files\": 1"));
    assert!(stdout.contains("\"failed_files\": 1"));

    let output_path = root.join("app/output/区域提取/area-result.xlsx");
    assert!(output_path.is_file());
    let actual_sheets = read_value_columns_with_iso_dates(&output_path, 16_384).unwrap();
    let expected_sheets = baseline()["worksheets"].as_array().unwrap().clone();
    assert_eq!(actual_sheets.len(), expected_sheets.len());
    for (actual, expected) in actual_sheets.iter().zip(&expected_sheets) {
        let expected_name = expected["name"].as_str().unwrap();
        assert_eq!(actual.name, expected_name);
        assert_eq!(
            actual.max_row,
            expected["max_row"].as_u64().unwrap() as usize
        );
        assert_eq!(
            actual.max_column,
            expected["max_column"].as_u64().unwrap() as usize
        );
        for (row_index, row) in expected["values"].as_array().unwrap().iter().enumerate() {
            for (column_index, expected_cell) in row.as_array().unwrap().iter().enumerate() {
                let address = excel_cell(row_index + 1, column_index + 1);
                assert_cell(
                    expected_cell,
                    actual.values.get(&address),
                    &address,
                    expected_name,
                );
            }
        }
    }
    for (path, original) in original_inputs {
        assert_eq!(
            fs::read(&path).unwrap(),
            original,
            "input changed: {}",
            path.display()
        );
    }

    let events = stderr
        .lines()
        .filter(|line| line.starts_with('{'))
        .filter_map(|line| serde_json::from_str::<Value>(line).ok())
        .collect::<Vec<_>>();
    for phase in [
        "source_read",
        "output_write",
        "output_validation",
        "output_commit",
    ] {
        assert!(
            events.iter().any(|event| event["phase"] == phase),
            "missing {phase} event"
        );
    }
    assert!(events
        .iter()
        .any(|event| event["actual_engine"] == "calamine"));
    assert!(events
        .iter()
        .any(|event| event["actual_engine"] == "rust_xlsxwriter"));
    assert!(events
        .iter()
        .any(|event| { event["phase"] == "source_read" && event["status"] == "failed" }));
    assert!(events
        .iter()
        .any(|event| { event["phase"] == "source_read" && event["stage_elapsed_ms"].is_number() }));

    let log_dir = root.join("app/logs");
    let history_path = fs::read_dir(&log_dir)
        .unwrap()
        .filter_map(|entry| entry.ok().map(|entry| entry.path()))
        .find(|path| {
            path.file_name()
                .unwrap()
                .to_string_lossy()
                .starts_with("PT-7-6_")
        })
        .unwrap();
    let history_bytes = fs::read(&history_path).unwrap();
    let history_text = String::from_utf8_lossy(&history_bytes);
    let history: Value = serde_json::from_slice(&history_bytes).unwrap();
    assert_eq!(history["feature_id"], "PT-7-6");
    assert_eq!(history["status"], "partial");
    assert_eq!(history["total_inputs"], 3);
    assert_eq!(history["result"]["extracted_sheets"], 4);
    assert_eq!(history["actual_engines"].as_array().unwrap().len(), 2);
    let sample_cell_value = expected_sheets
        .iter()
        .flat_map(|sheet| sheet["values"].as_array().unwrap())
        .flat_map(|row| row.as_array().unwrap())
        .find_map(|cell| {
            (cell["type"] == "string")
                .then(|| cell["value"].as_str())
                .flatten()
                .filter(|value| !value.is_empty())
        })
        .unwrap();
    assert!(
        !history_text.contains(sample_cell_value),
        "run history must not contain raw cell values"
    );

    let compared_cells = expected_sheets
        .iter()
        .map(|sheet| {
            sheet["values"]
                .as_array()
                .unwrap()
                .iter()
                .map(|row| row.as_array().unwrap().len())
                .sum::<usize>()
        })
        .sum::<usize>();
    let report = json!({
        "schema_version": 1,
        "status": "passed",
        "suite": "pytools.sheet_extract_area_menu_parity",
        "entry_point": "pytools menu -> 7 -> 6",
        "baseline": "probe/reports/sheet_extract_area_python_baseline.json",
        "inputs": {
            "xlsx": 3,
            "filtered_unverified_formats": 1,
            "damaged_xlsx_expected_failures": 1,
        },
        "stats": {
            "files": 2,
            "extracted_sheets": 4,
            "output_files": 1,
            "failed_files": 1,
        },
        "worksheet_count": actual_sheets.len(),
        "compared_cells": compared_cells,
        "actual_engines": ["calamine", "rust_xlsxwriter"],
        "checks": {
            "menu_task_loading_and_file_selection": "passed",
            "unsupported_format_filtered_before_reader": "passed",
            "damaged_xlsx_reported_as_partial_failure": "passed",
            "worksheet_names_dimensions_values_and_types": "preserved",
            "source_input_bytes": "preserved",
            "run_history_excludes_cell_values": "passed",
            "production_capability_created": false,
        }
    });
    let report_path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("probe/reports/sheet_extract_menu_parity.json");
    fs::write(report_path, serde_json::to_vec_pretty(&report).unwrap()).unwrap();
    fs::remove_dir_all(root).unwrap();
}

#[test]
fn interactive_menu_stops_when_extract_tasks_are_missing_or_full_sheet_capability_is_unavailable() {
    let cases = cases();
    let root = temporary_directory("sheet-extract-menu-gates");
    let no_tasks_config = create_menu_config(&root.join("no-tasks"), &[]);
    let no_tasks = run_menu(&no_tasks_config, "7\n6\n");
    assert!(no_tasks.status.success());
    let no_tasks_stdout = String::from_utf8_lossy(&no_tasks.stdout);
    assert!(no_tasks_stdout.contains("配置无启用任务"));
    assert!(!no_tasks_stdout.contains("文件路径>"));
    assert!(!root.join("no-tasks/app/output").exists());

    let mut full_sheet_tasks = load_tasks(&cases);
    full_sheet_tasks[0].area_spec.clear();
    let full_sheet_config = create_menu_config(&root.join("full-sheet"), &full_sheet_tasks);
    let full_sheet = run_menu(&full_sheet_config, "7\n6\n");
    assert!(full_sheet.status.success());
    let full_sheet_stdout = String::from_utf8_lossy(&full_sheet.stdout);
    assert!(full_sheet_stdout.contains("整表提取任务"));
    assert!(full_sheet_stdout.contains("没有匹配的 ExcelModifier 能力文件与平台"));
    assert!(!full_sheet_stdout.contains("文件路径>"));
    assert!(!root.join("full-sheet/app/output").exists());
    fs::remove_dir_all(root).unwrap();
}

#[test]
fn interactive_menu_supports_user_and_off_log_views() {
    let cases = cases();
    let tasks = load_tasks(&cases);
    let input_path = fixture_path("省级月度报表_A.xlsx");

    let user_root = temporary_directory("sheet-extract-menu-user-view");
    let user_config = create_menu_config(&user_root, &tasks);
    let user_output = run_menu_with_view(
        &user_config,
        &format!("7\n6\n{}\n", input_path.display()),
        "user",
    );
    assert!(user_output.status.success());
    let user_stdout = String::from_utf8_lossy(&user_output.stdout);
    let user_stderr = String::from_utf8_lossy(&user_output.stderr);
    assert!(user_stdout.contains("[完成] 源文件=1 提取工作表=2 输出文件=1 失败=0"));
    assert!(user_stderr.contains("正在读取源工作簿"));
    assert!(user_stderr.contains("工作表提取运行记录已保存"));
    assert!(user_root.join("app/logs").is_dir());
    fs::remove_dir_all(user_root).unwrap();

    let off_root = temporary_directory("sheet-extract-menu-off-view");
    let off_config = create_menu_config(&off_root, &tasks);
    let off_output = run_menu_with_view(
        &off_config,
        &format!("7\n6\n{}\n", input_path.display()),
        "off",
    );
    assert!(off_output.status.success());
    let off_stdout = String::from_utf8_lossy(&off_output.stdout);
    assert!(off_stdout.contains("\"files\": 1"));
    assert!(String::from_utf8_lossy(&off_output.stderr).is_empty());
    assert!(!off_root.join("app/logs").exists());
    fs::remove_dir_all(off_root).unwrap();
}

#[test]
fn unsupported_operations_and_existing_outputs_are_rejected_safely() {
    let cases = cases();
    let source = fixture_path("省级月度报表_A.xlsx");
    let root = temporary_directory("sheet-extract-reject");
    let output_dir = root.join("output");
    let mut loaded_tasks = load_tasks(&cases);

    let append_error = run_extract_sheets_area_candidate(
        std::slice::from_ref(&source),
        &loaded_tasks,
        &output_dir,
        true,
        Some("unused.xlsx"),
    )
    .unwrap_err();
    assert!(append_error.to_string().contains("append mode"));
    assert!(!output_dir.exists());

    loaded_tasks[0].area_spec.clear();
    let full_sheet_error = run_extract_sheets_area_candidate(
        std::slice::from_ref(&source),
        &loaded_tasks,
        &output_dir,
        false,
        Some("unused.xlsx"),
    )
    .unwrap_err();
    assert!(full_sheet_error
        .to_string()
        .contains("full-sheet extraction"));
    loaded_tasks = load_tasks(&cases);

    let date_output = root.join("date-output");
    let mut date_task = loaded_tasks[0].clone();
    date_task.workbook_keywords.clear();
    date_task.sheet_keywords.clear();
    date_task.exclude_keywords.clear();
    date_task.area_spec = "E1".to_owned();
    date_task.output_name = "date-output.xlsx".to_owned();
    let date_error = run_extract_sheets_area_candidate(
        std::slice::from_ref(&source),
        &[date_task],
        &date_output,
        false,
        Some("unused.xlsx"),
    )
    .unwrap_err();
    assert!(format!("{date_error:#}").contains("unsupported cell type"));
    assert!(!date_output.join("区域提取/date-output.xlsx").exists());

    let collision_dir = root.join("collision");
    let collision_path = collision_dir.join("区域提取/area-result.xlsx");
    fs::create_dir_all(collision_path.parent().unwrap()).unwrap();
    fs::write(&collision_path, b"existing output must survive").unwrap();
    let collision_error = run_extract_sheets_area_candidate(
        std::slice::from_ref(&source),
        &loaded_tasks,
        &collision_dir,
        false,
        Some("unused.xlsx"),
    )
    .unwrap_err();
    assert!(collision_error.to_string().contains("refuse to overwrite"));
    assert_eq!(
        fs::read(collision_path).unwrap(),
        b"existing output must survive"
    );

    fs::remove_dir_all(root).unwrap();
}
