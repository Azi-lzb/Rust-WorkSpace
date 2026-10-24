use excel_reader::{read_workbook, CellValue, SheetDump};
use excel_writer::{Format, Workbook};
use serde_json::Value;
use std::{
    fs,
    path::{Path, PathBuf},
    process::{Command, Output},
    time::{SystemTime, UNIX_EPOCH},
};

const TEMPLATE: &str = "tests/fixtures/s1_region_summary_template.xlsx";

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
        "audit_s1_region_summary_{label}_{}_{}",
        std::process::id(),
        nonce
    ));
    fs::create_dir_all(&root).expect("create temporary test directory");
    root
}

fn write_sample_workbook(path: &Path, rows: &[Vec<SampleCell<'_>>]) {
    let mut workbook = Workbook::new();
    let blank_format = Format::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_name("Data").expect("set test worksheet name");
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
    workbook.save(path).expect("save synthetic source workbook");
}

fn run_summary(root: &Path, output: &Path, developer_view: bool) -> Output {
    let mut command = Command::new(env!("CARGO_BIN_EXE_open-data-audit"));
    command
        .arg("s1-region-summary")
        .arg("--template-workbook")
        .arg(Path::new(env!("CARGO_MANIFEST_DIR")).join(TEMPLATE))
        .arg("--input-dir")
        .arg(root)
        .arg("--output-workbook")
        .arg(output)
        .arg("--feature-name")
        .arg("任意行汇总")
        .arg("--feature-name")
        .arg("固定行汇总");
    if developer_view {
        command.arg("--developer-view");
    }
    command
        .output()
        .expect("run S1-F02 named-region summary CLI")
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

fn create_sources(root: &Path) -> Vec<(PathBuf, Vec<u8>)> {
    fs::create_dir_all(root).expect("create synthetic source directory");
    let alpha = root.join("!Alpha_20260731.xlsx");
    write_sample_workbook(
        &alpha,
        &[
            vec![SampleCell::Blank; 4],
            vec![SampleCell::Blank; 4],
            vec![SampleCell::Blank; 4],
            vec![
                SampleCell::Blank,
                SampleCell::Text("甲"),
                SampleCell::Number(10.0),
            ],
            vec![
                SampleCell::Blank,
                SampleCell::Blank,
                SampleCell::Number(0.0),
            ],
            vec![SampleCell::Blank; 3],
            vec![
                SampleCell::Blank,
                SampleCell::Text("乙"),
                SampleCell::Number(20.0),
            ],
        ],
    );
    let beta = root.join("Beta_20260731.xlsx");
    write_sample_workbook(
        &beta,
        &[
            vec![SampleCell::Blank; 3],
            vec![SampleCell::Blank; 3],
            vec![SampleCell::Blank; 3],
            vec![
                SampleCell::Blank,
                SampleCell::Text("丙"),
                SampleCell::Number(30.0),
            ],
            vec![
                SampleCell::Blank,
                SampleCell::Text("丁"),
                SampleCell::Number(40.0),
            ],
        ],
    );
    vec![
        (
            alpha.clone(),
            fs::read(&alpha).expect("read source snapshot"),
        ),
        (beta.clone(), fs::read(&beta).expect("read source snapshot")),
    ]
}

#[test]
fn cli_summarizes_used_and_fixed_named_ranges_and_logs_both_views() {
    let root = temporary_root("output");
    let input_dir = root.join("input");
    let sources = create_sources(&input_dir);

    let user_output_path = root.join("user_summary.xlsx");
    let user = run_summary(&input_dir, &user_output_path, false);
    assert!(
        user.status.success(),
        "{}",
        String::from_utf8_lossy(&user.stderr)
    );
    let user_json: Value = serde_json::from_slice(&user.stdout).expect("parse user CLI result");
    assert_eq!(user_json["status"], "completed");
    assert_eq!(user_json["summary"]["source_workbooks"], 2);
    assert_eq!(user_json["summary"]["source_sheets"], 2);
    assert_eq!(user_json["summary"]["output_sheets"], 2);
    assert_eq!(user_json["summary"]["data_rows"], 9);
    assert!(user_json["summary"].get("reader").is_none());
    assert!(user_json["summary"].get("writer").is_none());
    assert!(!String::from_utf8_lossy(&user.stderr).contains("rust_xlsxwriter"));

    let output = read_workbook(&user_output_path).expect("read Rust region summary workbook");
    assert_eq!(
        output
            .sheets
            .iter()
            .map(|sheet| sheet.name.as_str())
            .collect::<Vec<_>>(),
        ["Data_2", "Data"]
    );
    assert_eq!(read_text(&output.sheets[0], "A1"), "来源文件");
    assert_eq!(read_text(&output.sheets[0], "B1"), "来源工作表");
    assert_eq!(read_text(&output.sheets[0], "C1"), "财务_单位");
    assert_eq!(read_text(&output.sheets[0], "D1"), "财务_金额");
    assert_eq!(read_text(&output.sheets[0], "A2"), "Alpha");
    assert_eq!(read_text(&output.sheets[0], "B2"), "Data");
    assert_eq!(read_text(&output.sheets[0], "C2"), "甲");
    assert_eq!(read_numeric(&output.sheets[0], "D3"), 0.0);
    assert_eq!(read_text(&output.sheets[0], "C4"), "丙");
    assert_eq!(read_text(&output.sheets[0], "C5"), "丁");
    assert_eq!(read_text(&output.sheets[1], "C4"), "乙");
    assert_eq!(read_text(&output.sheets[1], "C5"), "丙");
    assert_eq!(read_text(&output.sheets[1], "C6"), "丁");
    for (path, expected) in sources {
        assert_eq!(fs::read(path).expect("read source after run"), expected);
    }

    let developer_output_path = root.join("developer_summary.xlsx");
    let developer = run_summary(&input_dir, &developer_output_path, true);
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
    assert!(events[1]["source_read_time_ms"].is_number());
    assert!(events[1]["write_time_ms"].is_number());

    fs::remove_dir_all(&root).expect("remove synthetic test directory");
}

#[test]
fn cli_refuses_formula_sources_without_creating_output() {
    let root = temporary_root("formula_guard");
    let input_dir = root.join("input");
    fs::create_dir_all(&input_dir).expect("create synthetic source directory");
    write_sample_workbook(
        &input_dir.join("formula.xlsx"),
        &[
            vec![SampleCell::Blank; 3],
            vec![SampleCell::Blank; 3],
            vec![SampleCell::Blank; 3],
            vec![
                SampleCell::Blank,
                SampleCell::Text("甲"),
                SampleCell::Formula("=1+1"),
            ],
        ],
    );
    let output_path = root.join("summary.xlsx");
    let result = run_summary(&input_dir, &output_path, false);
    assert!(!result.status.success());
    let error = String::from_utf8_lossy(&result.stderr);
    assert!(error.contains("含公式的工作表"), "{error}");
    assert!(!output_path.exists());
    fs::remove_dir_all(&root).expect("remove synthetic test directory");
}
