use excel_reader::{read_workbook, CellValue, SheetDump};
use excel_writer::{Format, Workbook};
use serde_json::Value;
use std::{
    fs,
    path::PathBuf,
    process::{Command, Output},
    time::{SystemTime, UNIX_EPOCH},
};

#[derive(Clone)]
enum SampleCell<'a> {
    Text(&'a str),
    Number(f64),
    Blank,
    Formula(&'a str),
}

fn temporary_root(label: &str) -> PathBuf {
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system time after Unix epoch")
        .as_nanos();
    let root = std::env::temp_dir().join(format!(
        "audit_s1_summary_merge_{label}_{}_{}",
        std::process::id(),
        nonce
    ));
    fs::create_dir_all(&root).expect("create temporary test directory");
    root
}

fn write_sample_workbook(path: &PathBuf, sheets: &[(&str, Vec<Vec<SampleCell<'_>>>)]) {
    let mut workbook = Workbook::new();
    let blank_format = Format::new();
    for (name, rows) in sheets {
        let worksheet = workbook.add_worksheet();
        worksheet.set_name(*name).expect("set test worksheet name");
        for (row_index, row) in rows.iter().enumerate() {
            for (column_index, cell) in row.iter().enumerate() {
                let row_index = row_index as u32;
                let column_index = column_index as u16;
                match cell {
                    SampleCell::Text(value) => {
                        worksheet
                            .write_string(row_index, column_index, *value)
                            .expect("write test text");
                    }
                    SampleCell::Number(value) => {
                        worksheet
                            .write_number(row_index, column_index, *value)
                            .expect("write test number");
                    }
                    SampleCell::Blank => {
                        worksheet
                            .write_blank(row_index, column_index, &blank_format)
                            .expect("write test blank");
                    }
                    SampleCell::Formula(value) => {
                        worksheet
                            .write_formula(row_index, column_index, *value)
                            .expect("write test formula");
                    }
                }
            }
        }
    }
    workbook.save(path).expect("save synthetic workbook");
}

fn run_merge(input_dir: &PathBuf, output_path: &PathBuf, developer_view: bool) -> Output {
    let mut command = Command::new(env!("CARGO_BIN_EXE_open-data-audit"));
    command
        .arg("s1-merge-workbook-tables")
        .arg("--input-dir")
        .arg(input_dir)
        .arg("--output-workbook")
        .arg(output_path);
    if developer_view {
        command.arg("--developer-view");
    }
    command.output().expect("run S1-F02 table-merge CLI")
}

fn read_text(sheet: &SheetDump, address: &str) -> String {
    match sheet
        .values
        .get(address)
        .map(|value| CellValue::from_probe_json(value).expect("decode output cell"))
    {
        Some(CellValue::Text(value)) => value,
        other => panic!("expected text in {address}, got {other:?}"),
    }
}

fn read_numeric(sheet: &SheetDump, address: &str) -> f64 {
    match sheet
        .values
        .get(address)
        .map(|value| CellValue::from_probe_json(value).expect("decode output cell"))
    {
        Some(CellValue::Number(value) | CellValue::DateTimeSerial(value)) => value,
        Some(CellValue::Integer(value)) => value as f64,
        other => panic!("expected number in {address}, got {other:?}"),
    }
}

fn create_inputs(root: &PathBuf) {
    let nested = root.join("机构报送");
    fs::create_dir_all(&nested).expect("create nested source directory");
    write_sample_workbook(
        &root.join("上一版汇总.xlsx"),
        &[(
            "Data",
            vec![
                vec![SampleCell::Text("旧表")],
                vec![SampleCell::Text("不应进入")],
            ],
        )],
    );
    write_sample_workbook(
        &nested.join("!Alpha_20260731.xlsx"),
        &[
            (
                "Data",
                vec![
                    vec![SampleCell::Text("单位"), SampleCell::Text("数值")],
                    vec![SampleCell::Text("甲"), SampleCell::Number(10.0)],
                    vec![SampleCell::Text("乙"), SampleCell::Number(20.0)],
                ],
            ),
            (
                "Stats",
                vec![
                    vec![SampleCell::Text("指标"), SampleCell::Text("数值")],
                    vec![SampleCell::Text("总数"), SampleCell::Number(2.0)],
                ],
            ),
            ("Empty", vec![vec![SampleCell::Blank]]),
        ],
    );
    write_sample_workbook(
        &nested.join("Beta_20260731.xlsx"),
        &[(
            "Data",
            vec![
                vec![SampleCell::Text("单位"), SampleCell::Text("数值")],
                vec![SampleCell::Text("丙"), SampleCell::Number(30.0)],
                vec![SampleCell::Blank, SampleCell::Blank],
            ],
        )],
    );
    write_sample_workbook(
        &nested.join("Gamma_20260731.xlsx"),
        &[(
            "Different",
            vec![
                vec![SampleCell::Text("地区"), SampleCell::Text("数量")],
                vec![SampleCell::Text("东"), SampleCell::Number(5.0)],
            ],
        )],
    );
}

#[test]
fn cli_groups_tables_by_header_and_exposes_user_and_developer_views() {
    let root = temporary_root("groups");
    create_inputs(&root);

    let user_output_path = root.join("user_summary.xlsx");
    let user = run_merge(&root, &user_output_path, false);
    assert!(
        user.status.success(),
        "{}",
        String::from_utf8_lossy(&user.stderr)
    );
    let user_json: Value = serde_json::from_slice(&user.stdout).expect("parse user CLI result");
    assert_eq!(user_json["status"], "completed");
    assert_eq!(user_json["summary"]["source_workbooks"], 3);
    assert_eq!(user_json["summary"]["source_sheets"], 5);
    assert_eq!(user_json["summary"]["merged_sheets"], 3);
    assert_eq!(user_json["summary"]["data_rows"], 5);
    assert!(user_json["summary"].get("reader").is_none());
    assert!(user_json["summary"].get("writer").is_none());
    assert!(!String::from_utf8_lossy(&user.stderr).contains("rust_xlsxwriter"));

    let workbook = read_workbook(&user_output_path).expect("read Rust summary workbook");
    let names = workbook
        .sheets
        .iter()
        .map(|sheet| sheet.name.as_str())
        .collect::<Vec<_>>();
    assert_eq!(names, ["Different", "Stats", "Data"]);
    assert_eq!(read_text(&workbook.sheets[2], "A1"), "来源文件");
    assert_eq!(read_text(&workbook.sheets[2], "B1"), "来源工作表");
    assert_eq!(read_text(&workbook.sheets[2], "A2"), "!Alpha_20260731.xlsx");
    assert_eq!(read_text(&workbook.sheets[2], "B2"), "Data");
    assert_eq!(read_text(&workbook.sheets[2], "C2"), "甲");
    assert_eq!(read_numeric(&workbook.sheets[2], "D2"), 10.0);
    assert_eq!(read_text(&workbook.sheets[2], "A4"), "Beta_20260731.xlsx");
    assert_eq!(read_text(&workbook.sheets[0], "C2"), "东");

    let developer_output_path = root.join("developer_summary.xlsx");
    let developer = run_merge(&root, &developer_output_path, true);
    assert!(
        developer.status.success(),
        "{}",
        String::from_utf8_lossy(&developer.stderr)
    );
    let events = String::from_utf8_lossy(&developer.stderr)
        .lines()
        .map(|line| serde_json::from_str::<Value>(line).expect("parse developer event"))
        .collect::<Vec<_>>();
    assert_eq!(events.len(), 2);
    assert_eq!(events[0]["status"], "started");
    assert_eq!(events[1]["status"], "completed");
    assert_eq!(events[1]["reader"], "calamine+namespace_aware_ooxml");
    assert_eq!(events[1]["writer"], "rust_xlsxwriter 0.101.0");
    assert!(events[1]["read_time_ms"].is_number());
    assert!(events[1]["write_time_ms"].is_number());
    let developer_json: Value =
        serde_json::from_slice(&developer.stdout).expect("parse developer CLI result");
    assert_eq!(
        developer_json["summary"]["writer"],
        "rust_xlsxwriter 0.101.0"
    );

    fs::remove_dir_all(&root).expect("remove synthetic test directory");
}

#[test]
fn cli_refuses_formula_sources_without_creating_output() {
    let root = temporary_root("formula_guard");
    write_sample_workbook(
        &root.join("formula.xlsx"),
        &[(
            "Data",
            vec![
                vec![SampleCell::Text("指标")],
                vec![SampleCell::Formula("=1+1")],
            ],
        )],
    );
    let output_path = root.join("summary.xlsx");
    let result = run_merge(&root, &output_path, false);
    assert!(!result.status.success());
    let error = String::from_utf8_lossy(&result.stderr);
    assert!(error.contains("含公式的工作表"), "{error}");
    assert!(!output_path.exists());
    fs::remove_dir_all(&root).expect("remove synthetic test directory");
}
