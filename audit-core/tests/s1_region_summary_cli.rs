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
    write_sample_workbook_with_metadata(path, rows, None, None);
}

fn write_sample_workbook_with_metadata(
    path: &Path,
    rows: &[Vec<SampleCell<'_>>],
    local_value: Option<&str>,
    global_value: Option<&str>,
) {
    let mut workbook = Workbook::new();
    let blank_format = Format::new();
    {
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
        if let Some(value) = local_value {
            worksheet
                .write_string(0, 5, value)
                .expect("write local metadata field");
        }
    }
    if let Some(value) = global_value {
        let worksheet = workbook.add_worksheet();
        worksheet
            .set_name("Meta")
            .expect("set metadata worksheet name");
        worksheet
            .write_string(0, 0, value)
            .expect("write global metadata field");
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

fn run_summary_with_history(
    template: &Path,
    input_dir: &Path,
    history: &Path,
    output: &Path,
    developer_view: bool,
) -> Output {
    let mut command = Command::new(env!("CARGO_BIN_EXE_open-data-audit"));
    command
        .arg("s1-region-summary")
        .arg("--template-workbook")
        .arg(template)
        .arg("--input-dir")
        .arg(input_dir)
        .arg("--output-workbook")
        .arg(output)
        .arg("--history-workbook")
        .arg(history)
        .arg("--feature-name")
        .arg("任意行汇总");
    if developer_view {
        command.arg("--developer-view");
    }
    command
        .output()
        .expect("run S1-F02 history-enriched summary CLI")
}

fn run_fallback_summary(
    template: &Path,
    input_dir: &Path,
    output: &Path,
    feature_name: &str,
    developer_view: bool,
) -> Output {
    let mut command = Command::new(env!("CARGO_BIN_EXE_open-data-audit"));
    command
        .arg("s1-region-summary")
        .arg("--template-workbook")
        .arg(template)
        .arg("--input-dir")
        .arg(input_dir)
        .arg("--output-workbook")
        .arg(output)
        .arg("--feature-name")
        .arg(feature_name);
    if developer_view {
        command.arg("--developer-view");
    }
    command.output().expect("run S1-F02 fallback summary CLI")
}

fn write_fallback_summary_template(path: &Path) {
    let mut workbook = Workbook::new();
    let fallback = workbook.add_worksheet();
    fallback
        .set_name("兜底模板")
        .expect("set fallback template sheet name");
    fallback
        .write_string(0, 1, "字段甲")
        .expect("write fallback first header");
    fallback
        .write_string(0, 2, "字段乙")
        .expect("write fallback second header");
    workbook
        .add_worksheet()
        .set_name("Meta")
        .expect("set fallback template metadata sheet name");
    workbook
        .define_name("任意行汇总区域", "='兜底模板'!$B$2:$C$100")
        .expect("define fallback data range");
    workbook
        .define_name("固定行汇总区域", "='兜底模板'!$B$2:$C$6")
        .expect("define fixed-row fallback data range");
    workbook
        .define_name("表头区域", "='兜底模板'!$B$1:$C$1")
        .expect("define fallback header range");
    workbook
        .define_name("全局单元格区域.批次", "=Meta!$A$1")
        .expect("define fallback global metadata field");
    workbook
        .define_name("单元格区域.区域", "='兜底模板'!$F$1")
        .expect("define fallback local metadata field");
    workbook.save(path).expect("save fallback summary template");
}

fn write_fallback_summary_source(path: &Path, second_path: &Path) {
    let mut workbook = Workbook::new();
    let source = workbook.add_worksheet();
    source
        .set_name("未匹配表")
        .expect("set fallback source sheet name");
    source
        .write_string(0, 5, "区域甲")
        .expect("write fallback local metadata value");
    source
        .write_string(3, 1, "记录甲")
        .expect("write first fallback row label");
    source
        .write_number(3, 2, 11.0)
        .expect("write first fallback row value");
    source
        .write_string(5, 1, "记录乙")
        .expect("write second fallback row label");
    source
        .write_number(5, 2, 0.0)
        .expect("write second fallback row value");

    let empty = workbook.add_worksheet();
    empty
        .set_name("无兜底数据")
        .expect("set empty source sheet name");
    empty
        .write_string(0, 0, "不在兜底列范围内")
        .expect("write value outside fallback data columns");

    let other = workbook.add_worksheet();
    other
        .set_name("另一张表")
        .expect("set second unmatched source sheet name");
    other
        .write_string(0, 5, "区域乙")
        .expect("write second unmatched local metadata value");
    other
        .write_string(1, 1, "另一条记录")
        .expect("write second unmatched row label");
    other
        .write_number(1, 2, 12.0)
        .expect("write second unmatched row value");

    let metadata = workbook.add_worksheet();
    metadata
        .set_name("Meta")
        .expect("set fallback source metadata sheet name");
    metadata
        .write_string(0, 0, "批次-01")
        .expect("write fallback global metadata value");
    workbook.save(path).expect("save fallback summary source");

    let mut second = Workbook::new();
    let source = second.add_worksheet();
    source
        .set_name("未匹配表")
        .expect("set repeated fallback source sheet name");
    source
        .write_string(0, 5, "区域丙")
        .expect("write repeated fallback local metadata value");
    source
        .write_string(1, 1, "记录丙")
        .expect("write repeated fallback row label");
    source
        .write_number(1, 2, 5.0)
        .expect("write repeated fallback row value");
    let metadata = second.add_worksheet();
    metadata
        .set_name("Meta")
        .expect("set repeated fallback metadata sheet name");
    metadata
        .write_string(0, 0, "批次-02")
        .expect("write repeated fallback global metadata value");
    second
        .save(second_path)
        .expect("save repeated fallback summary source");
}

fn write_history_summary_template(path: &Path) {
    let mut workbook = Workbook::new();
    {
        let worksheet = workbook.add_worksheet();
        worksheet
            .set_name("本地校验结果")
            .expect("set history template sheet name");
        for (column, header) in [
            "批次",
            "表单名称",
            "规则编号",
            "规则类型",
            "规则描述",
            "校验字段",
        ]
        .iter()
        .enumerate()
        {
            worksheet
                .write_string(1, column as u16 + 1, *header)
                .expect("write history template header");
        }
    }
    workbook
        .define_name("任意行汇总区域", "='本地校验结果'!$B$3:$G$10")
        .expect("define history summary range");
    workbook
        .define_name("表头区域", "='本地校验结果'!$B$2:$G$2")
        .expect("define history header range");
    workbook.save(path).expect("save history summary template");
}

fn write_history_summary_source(path: &Path) {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .set_name("本地校验结果")
        .expect("set history source sheet name");
    let rows = [
        [
            "2026-07",
            "存量单位贷款",
            "1.0",
            "阈值",
            "不应低于",
            "贷款余额",
        ],
        [
            "2026-07",
            "存量单位贷款",
            "2.0",
            "阈值",
            "不应低于",
            "贷款余额",
        ],
    ];
    for (row_index, row) in rows.iter().enumerate() {
        for (column_index, value) in row.iter().enumerate() {
            let row = row_index as u32 + 2;
            let column = column_index as u16 + 1;
            if column_index == 2 {
                worksheet
                    .write_number(
                        row,
                        column,
                        value.parse::<f64>().expect("parse rule identifier"),
                    )
                    .expect("write numeric rule identifier");
            } else {
                worksheet
                    .write_string(row, column, *value)
                    .expect("write source history key");
            }
        }
    }
    workbook.save(path).expect("save history summary source");
}

fn write_history_workbook(path: &Path, include_formula: bool) {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .set_name("本地校验结果")
        .expect("set history worksheet name");
    let headers = [
        "来源文件",
        "来源工作表",
        "批次",
        "表单名称",
        "规则编号",
        "规则类型",
        "规则描述",
        "校验字段",
        "历史触发条数",
        "历史分类说明",
        "审核结果",
    ];
    for (column, header) in headers.iter().enumerate() {
        worksheet
            .write_string(0, column as u16, *header)
            .expect("write history header");
    }
    for (row, count, class, result) in [
        ("1", 10.0, "类别甲", "通过"),
        ("1", 12.0, "类别乙", "未通过"),
    ] {
        let row_index = if class == "类别甲" { 1 } else { 2 };
        let values = [
            "Sample",
            "本地校验结果",
            "2026-07",
            "存量单位贷款",
            row,
            "阈值",
            "不应低于",
            "贷款余额",
        ];
        for (column, value) in values.iter().enumerate() {
            worksheet
                .write_string(row_index, column as u16, *value)
                .expect("write history identity value");
        }
        worksheet
            .write_number(row_index, 8, count)
            .expect("write history trigger count");
        worksheet
            .write_string(row_index, 9, class)
            .expect("write history classification");
        worksheet
            .write_string(row_index, 10, result)
            .expect("write history audit result");
    }
    if include_formula {
        worksheet
            .write_formula(1, 10, "=1+1")
            .expect("write formula to history sheet");
    }
    workbook.save(path).expect("save history workbook");
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

fn read_value(sheet: &SheetDump, address: &str) -> CellValue {
    sheet
        .values
        .get(address)
        .map(|value| CellValue::from_probe_json(value).expect("decode output cell"))
        .unwrap_or(CellValue::Empty)
}

fn create_sources(root: &Path) -> Vec<(PathBuf, Vec<u8>)> {
    fs::create_dir_all(root).expect("create synthetic source directory");
    let alpha = root.join("!Alpha_20260731.xlsx");
    write_sample_workbook_with_metadata(
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
        Some("甲组"),
        Some("批次甲"),
    );
    let beta = root.join("Beta_20260731.xlsx");
    write_sample_workbook_with_metadata(
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
        Some("乙组"),
        None,
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
    assert_eq!(user_json["summary"]["source_sheets"], 3);
    assert_eq!(user_json["summary"]["output_sheets"], 2);
    assert_eq!(user_json["summary"]["data_rows"], 9);
    assert_eq!(user_json["summary"]["metadata_fields_left_blank"], 2);
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
    assert_eq!(read_text(&output.sheets[0], "C1"), "批次");
    assert_eq!(read_text(&output.sheets[0], "D1"), "分组");
    assert_eq!(read_text(&output.sheets[0], "E1"), "财务_单位");
    assert_eq!(read_text(&output.sheets[0], "F1"), "财务_金额");
    assert_eq!(read_text(&output.sheets[0], "A2"), "Alpha");
    assert_eq!(read_text(&output.sheets[0], "B2"), "Data");
    assert_eq!(read_text(&output.sheets[0], "C2"), "批次甲");
    assert_eq!(read_text(&output.sheets[0], "D2"), "甲组");
    assert_eq!(read_text(&output.sheets[0], "E2"), "甲");
    assert_eq!(read_numeric(&output.sheets[0], "F3"), 0.0);
    assert_eq!(read_value(&output.sheets[0], "C4"), CellValue::Empty);
    assert_eq!(read_text(&output.sheets[0], "D4"), "乙组");
    assert_eq!(read_text(&output.sheets[0], "E4"), "丙");
    assert_eq!(read_text(&output.sheets[0], "E5"), "丁");
    assert_eq!(read_text(&output.sheets[1], "E4"), "乙");
    assert_eq!(read_text(&output.sheets[1], "E5"), "丙");
    assert_eq!(read_text(&output.sheets[1], "E6"), "丁");
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
    assert_eq!(events[1]["metadata_fields_left_blank"], 2);

    fs::remove_dir_all(&root).expect("remove synthetic test directory");
}

#[test]
fn cli_summarizes_unmatched_source_sheets_with_the_fallback_template() {
    let root = temporary_root("fallback_summary");
    let input_dir = root.join("input");
    fs::create_dir_all(&input_dir).expect("create fallback source directory");
    let template = root.join("fallback_template.xlsx");
    let source = input_dir.join("FallbackSource.xlsx");
    let second_source = input_dir.join("FallbackSource2.xlsx");
    let output_path = root.join("fallback_summary.xlsx");
    let fixed_output_path = root.join("fixed_fallback_summary.xlsx");
    write_fallback_summary_template(&template);
    write_fallback_summary_source(&source, &second_source);
    let template_before = fs::read(&template).expect("read fallback template snapshot");
    let source_before = fs::read(&source).expect("read fallback source snapshot");
    let second_source_before = fs::read(&second_source).expect("read second source snapshot");

    let result = run_fallback_summary(&template, &input_dir, &output_path, "任意行汇总", true);
    assert!(
        result.status.success(),
        "{}",
        String::from_utf8_lossy(&result.stderr)
    );
    let payload: Value = serde_json::from_slice(&result.stdout).expect("parse fallback CLI result");
    assert_eq!(payload["status"], "completed");
    assert_eq!(payload["summary"]["features"], 1);
    assert_eq!(payload["summary"]["output_sheets"], 2);
    assert_eq!(payload["summary"]["data_rows"], 4);

    let output = read_workbook(&output_path).expect("read fallback summary workbook");
    assert_eq!(
        output
            .sheets
            .iter()
            .map(|sheet| sheet.name.as_str())
            .collect::<Vec<_>>(),
        ["另一张表", "未匹配表"]
    );
    let other_sheet = &output.sheets[0];
    assert_eq!(read_text(other_sheet, "B2"), "另一张表");
    assert_eq!(read_text(other_sheet, "C2"), "批次-01");
    assert_eq!(read_text(other_sheet, "D2"), "区域乙");
    assert_eq!(read_text(other_sheet, "E2"), "另一条记录");
    assert_eq!(read_numeric(other_sheet, "F2"), 12.0);

    let sheet = &output.sheets[1];
    assert_eq!(read_text(sheet, "A1"), "来源文件");
    assert_eq!(read_text(sheet, "B1"), "来源工作表");
    assert_eq!(read_text(sheet, "C1"), "批次");
    assert_eq!(read_text(sheet, "D1"), "区域");
    assert_eq!(read_text(sheet, "E1"), "字段甲");
    assert_eq!(read_text(sheet, "F1"), "字段乙");
    assert_eq!(read_text(sheet, "B2"), "未匹配表");
    assert_eq!(read_text(sheet, "C2"), "批次-01");
    assert_eq!(read_text(sheet, "D2"), "区域甲");
    assert_eq!(read_text(sheet, "E2"), "记录甲");
    assert_eq!(read_numeric(sheet, "F2"), 11.0);
    assert_eq!(read_text(sheet, "E3"), "记录乙");
    assert_eq!(read_numeric(sheet, "F3"), 0.0);
    assert_eq!(read_text(sheet, "B4"), "未匹配表");
    assert_eq!(read_text(sheet, "C4"), "批次-02");
    assert_eq!(read_text(sheet, "D4"), "区域丙");
    assert_eq!(read_text(sheet, "E4"), "记录丙");
    assert_eq!(read_numeric(sheet, "F4"), 5.0);
    assert_eq!(
        fs::read(&template).expect("read fallback template after run"),
        template_before
    );
    assert_eq!(
        fs::read(&source).expect("read fallback source after run"),
        source_before
    );
    assert_eq!(
        fs::read(&second_source).expect("read second source after run"),
        second_source_before
    );

    let fixed_result = run_fallback_summary(
        &template,
        &input_dir,
        &fixed_output_path,
        "固定行汇总",
        false,
    );
    assert!(
        fixed_result.status.success(),
        "{}",
        String::from_utf8_lossy(&fixed_result.stderr)
    );
    let fixed_payload: Value =
        serde_json::from_slice(&fixed_result.stdout).expect("parse fixed fallback CLI result");
    assert_eq!(fixed_payload["summary"]["output_sheets"], 2);
    assert_eq!(fixed_payload["summary"]["data_rows"], 4);
    let fixed_output = read_workbook(&fixed_output_path).expect("read fixed fallback summary");
    assert_eq!(fixed_output.sheets.len(), 2);
    assert_eq!(read_text(&fixed_output.sheets[1], "E2"), "记录甲");
    assert_eq!(read_numeric(&fixed_output.sheets[1], "F4"), 5.0);
    assert_eq!(
        fs::read(&template).expect("read fixed fallback template after run"),
        template_before
    );

    fs::remove_dir_all(&root).expect("remove fallback summary test directory");
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

#[test]
fn cli_appends_builtin_history_fields_with_numeric_text_key_normalization() {
    let root = temporary_root("history_append");
    let input_dir = root.join("input");
    fs::create_dir_all(&input_dir).expect("create history summary source directory");
    let template = root.join("template.xlsx");
    let source = input_dir.join("Sample.xlsx");
    let history = root.join("history.xlsx");
    let output_path = root.join("summary.xlsx");
    write_history_summary_template(&template);
    write_history_summary_source(&source);
    write_history_workbook(&history, false);
    let source_before = fs::read(&source).expect("read original source snapshot");

    let result = run_summary_with_history(&template, &input_dir, &history, &output_path, true);
    assert!(
        result.status.success(),
        "{}",
        String::from_utf8_lossy(&result.stderr)
    );
    let payload: Value = serde_json::from_slice(&result.stdout).expect("parse CLI summary");
    assert_eq!(
        payload["summary"]["history_reader"],
        "calamine+namespace_aware_ooxml"
    );
    assert_eq!(payload["summary"]["history_rows_matched"], 1);
    assert_eq!(payload["summary"]["sheets"][0]["column_count"], 11);
    assert!(payload["summary"]["history_elapsed_ms"].is_number());

    let events = String::from_utf8_lossy(&result.stderr)
        .lines()
        .map(|line| serde_json::from_str::<Value>(line).expect("parse developer event"))
        .collect::<Vec<_>>();
    assert_eq!(events.len(), 2);
    assert_eq!(
        events[1]["history_reader"],
        "calamine+namespace_aware_ooxml"
    );
    assert_eq!(events[1]["history_rows_matched"], 1);
    assert!(events[1]["history_elapsed_ms"].is_number());

    let output = read_workbook(&output_path).expect("read history-enriched summary");
    let sheet = output
        .sheets
        .iter()
        .find(|sheet| sheet.name == "本地校验结果")
        .expect("find history-enriched summary sheet");
    assert_eq!(read_text(sheet, "I1"), "历史触发条数");
    assert_eq!(read_text(sheet, "J1"), "历史分类说明");
    assert_eq!(read_text(sheet, "K1"), "审核结果");
    assert_eq!(read_text(sheet, "I2"), "10|12");
    assert_eq!(read_text(sheet, "J2"), "类别甲|类别乙");
    assert_eq!(read_text(sheet, "K2"), "通过|未通过");
    assert_eq!(read_value(sheet, "I3"), CellValue::Empty);
    assert_eq!(read_value(sheet, "J3"), CellValue::Empty);
    assert_eq!(read_value(sheet, "K3"), CellValue::Empty);
    assert_eq!(
        fs::read(&source).expect("read source after run"),
        source_before
    );

    fs::remove_dir_all(&root).expect("remove synthetic test directory");
}

#[test]
fn cli_refuses_formula_history_sheet_before_creating_output() {
    let root = temporary_root("history_formula_guard");
    let input_dir = root.join("input");
    fs::create_dir_all(&input_dir).expect("create history summary source directory");
    let template = root.join("template.xlsx");
    let source = input_dir.join("Sample.xlsx");
    let history = root.join("history.xlsx");
    let output_path = root.join("summary.xlsx");
    write_history_summary_template(&template);
    write_history_summary_source(&source);
    write_history_workbook(&history, true);
    let source_before = fs::read(&source).expect("read original source snapshot");

    let result = run_summary_with_history(&template, &input_dir, &history, &output_path, false);
    assert!(!result.status.success());
    let error = String::from_utf8_lossy(&result.stderr);
    assert!(error.contains("包含公式"), "{error}");
    assert!(!output_path.exists());
    assert_eq!(
        fs::read(&source).expect("read source after refusal"),
        source_before
    );

    fs::remove_dir_all(&root).expect("remove synthetic test directory");
}
