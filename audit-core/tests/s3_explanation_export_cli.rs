use excel_reader::{read_workbook, CellValue};
use excel_writer::Workbook;
use serde_json::Value;
use std::{
    fs,
    path::PathBuf,
    process::{Command, Output},
    time::{SystemTime, UNIX_EPOCH},
};

fn temporary_root() -> PathBuf {
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system time after Unix epoch")
        .as_nanos();
    let root = std::env::temp_dir().join(format!(
        "audit_s3_explanation_export_{}_{}",
        std::process::id(),
        nonce
    ));
    fs::create_dir_all(&root).expect("create temporary test directory");
    root
}

fn write_comparison(path: &PathBuf) {
    let mut workbook = Workbook::new();
    let notes = workbook.add_worksheet();
    notes.set_name("封面").unwrap();
    notes.write_string(0, 0, "不是比较结果").unwrap();

    let comparison = workbook.add_worksheet();
    comparison.set_name("比较结果").unwrap();
    let headers = [
        "机构类代码",
        "机构类名称",
        "地区代码",
        "地区名称",
        "指标代码",
        "指标名称",
        "数据值",
        "上期值",
        "输出说明",
        "说明内容",
        "增减额（先出现）",
        "增减额（后出现）",
    ];
    for (column, header) in headers.iter().enumerate() {
        comparison.write_string(0, column as u16, *header).unwrap();
    }
    let rows = [
        [
            "7020001",
            "示例村镇银行",
            "4401",
            "惠城区",
            "I-01",
            "指标一",
            "0.418462",
            "0.250000",
            "是",
            "补充说明一",
            "-0.168462",
            "999",
        ],
        [
            "12345",
            "普通银行",
            "510101",
            "省级",
            "I-02",
            "指标二",
            "10.5",
            "10",
            "否",
            "不应导出",
            "0.5",
            "998",
        ],
        [
            "12345",
            "普通银行",
            "510102",
            "城北区",
            "I-03",
            "指标三",
            "25",
            "20",
            " YES ",
            "补充说明三",
            "5",
            "997",
        ],
        [
            "99999",
            "地方机构",
            "6101",
            "新区",
            "I-04",
            "指标四",
            "40",
            "35",
            "✓",
            "补充说明四",
            "5",
            "996",
        ],
        [
            "4567",
            "未映射机构",
            "3201",
            "南区",
            "I-05",
            "指标五",
            "12",
            "11",
            "1",
            "补充说明五",
            "1",
            "995",
        ],
    ];
    for (row, values) in rows.iter().enumerate() {
        for (column, value) in values.iter().enumerate() {
            if matches!(column, 6 | 7 | 10) {
                comparison
                    .write_number(row as u32 + 1, column as u16, value.parse::<f64>().unwrap())
                    .unwrap();
            } else {
                comparison
                    .write_string(row as u32 + 1, column as u16, *value)
                    .unwrap();
            }
        }
    }
    workbook.save(path).unwrap();
}

fn write_config(path: &PathBuf) {
    let mut workbook = Workbook::new();
    let mapping = workbook.add_worksheet();
    mapping.set_name("机构地区参照").unwrap();
    for (column, header) in ["机构类代码", "地区代码", "机构名称"].iter().enumerate() {
        mapping.write_string(0, column as u16, *header).unwrap();
    }
    for (row, values) in [
        ["7020001", "4401", "精确映射机构"],
        ["12345", "510102", "精确市级机构"],
        ["12345", "*", "显式通配机构"],
        ["99999", "*", "全国通配机构"],
    ]
    .iter()
    .enumerate()
    {
        for (column, value) in values.iter().enumerate() {
            mapping
                .write_string(row as u32 + 1, column as u16, *value)
                .unwrap();
        }
    }
    workbook.save(path).unwrap();
}

fn run_cli(comparison: &PathBuf, config: &PathBuf, output: &PathBuf, developer: bool) -> Output {
    let mut command = Command::new(env!("CARGO_BIN_EXE_open-data-audit"));
    command
        .arg("s3-export-explanations")
        .arg("--comparison-file")
        .arg(comparison)
        .arg("--output-workbook")
        .arg(output)
        .arg("--config-workbook")
        .arg(config);
    if developer {
        command.arg("--developer-view");
    }
    command.output().expect("run S3-F02 CLI")
}

fn text_value(workbook: &excel_reader::WorkbookDump, sheet: usize, address: &str) -> String {
    match workbook.sheets[sheet]
        .values
        .get(address)
        .map(|value| CellValue::from_probe_json(value).expect("decode output cell"))
    {
        Some(CellValue::Text(value)) => value,
        other => panic!("expected text in {address}, got {other:?}"),
    }
}

fn numeric_value(workbook: &excel_reader::WorkbookDump, sheet: usize, address: &str) -> f64 {
    match workbook.sheets[sheet]
        .values
        .get(address)
        .map(|value| CellValue::from_probe_json(value).expect("decode output cell"))
    {
        Some(CellValue::Number(value) | CellValue::DateTimeSerial(value)) => value,
        Some(CellValue::Integer(value)) => value as f64,
        other => panic!("expected numeric value in {address}, got {other:?}"),
    }
}

#[test]
fn s3_f02_cli_exports_selected_rows_and_reports_both_views() {
    let root = temporary_root();
    let comparison = root.join("comparison.xlsx");
    let config = root.join("mapping.xlsx");
    let user_output_path = root.join("user_explanations.xlsx");
    let developer_output_path = root.join("developer_explanations.xlsx");
    write_comparison(&comparison);
    write_config(&config);
    let comparison_before = fs::read(&comparison).unwrap();
    let config_before = fs::read(&config).unwrap();

    let user = run_cli(&comparison, &config, &user_output_path, false);
    assert!(
        user.status.success(),
        "{}",
        String::from_utf8_lossy(&user.stderr)
    );
    let user_json: Value = serde_json::from_slice(&user.stdout).unwrap();
    assert_eq!(user_json["status"], "completed");
    assert_eq!(user_json["selected_rows"], 4);
    assert!(user_json.get("reader").is_none());
    assert!(user_json.get("writer").is_none());
    assert!(!String::from_utf8_lossy(&user.stderr).contains("calamine"));

    let output = read_workbook(&user_output_path).unwrap();
    assert_eq!(
        output
            .sheets
            .iter()
            .map(|sheet| sheet.name.as_str())
            .collect::<Vec<_>>(),
        ["说明文件", "生成信息"]
    );
    assert_eq!(text_value(&output, 0, "A1"), "xx说明");
    assert_eq!(text_value(&output, 0, "A2"), "机构代码");
    assert_eq!(text_value(&output, 0, "E2"), "数据值");
    assert_eq!(text_value(&output, 0, "G2"), "增减额（先出现）");
    assert_eq!(text_value(&output, 0, "A3"), "70200014401");
    assert_eq!(text_value(&output, 0, "B3"), "精确映射机构（惠城区）");
    assert_eq!(text_value(&output, 0, "H3"), "补充说明一");
    assert_eq!(numeric_value(&output, 0, "E3"), 0.418462);
    assert_eq!(numeric_value(&output, 0, "G3"), -0.168462);
    assert_eq!(text_value(&output, 0, "A4"), "12345510102");
    assert_eq!(text_value(&output, 0, "B4"), "精确市级机构（城北区）");
    assert_eq!(text_value(&output, 0, "A5"), "99999");
    assert_eq!(text_value(&output, 0, "B5"), "全国通配机构（新区）");
    assert_eq!(text_value(&output, 0, "A6"), "4567");
    assert_eq!(text_value(&output, 0, "B6"), "未映射机构（南区）");
    assert_eq!(output.sheets[0].merges, ["A1:H1"]);
    assert_eq!(text_value(&output, 1, "A1"), "项目");
    assert_eq!(
        text_value(&output, 1, "B2"),
        comparison.to_string_lossy().into_owned()
    );
    assert_eq!(numeric_value(&output, 1, "B4"), 4.0);

    let existing_output_before = fs::read(&user_output_path).unwrap();
    let existing_output = run_cli(&comparison, &config, &user_output_path, false);
    assert!(!existing_output.status.success());
    assert_eq!(fs::read(&user_output_path).unwrap(), existing_output_before);

    let developer = run_cli(&comparison, &config, &developer_output_path, true);
    assert!(
        developer.status.success(),
        "{}",
        String::from_utf8_lossy(&developer.stderr)
    );
    let developer_json: Value = serde_json::from_slice(&developer.stdout).unwrap();
    assert_eq!(developer_json["reader"], "calamine");
    assert_eq!(developer_json["writer"], "rust_xlsxwriter");
    let events = String::from_utf8_lossy(&developer.stderr)
        .lines()
        .map(|line| serde_json::from_str::<Value>(line).unwrap())
        .collect::<Vec<_>>();
    assert_eq!(events.len(), 2);
    assert_eq!(events[0]["status"], "started");
    assert_eq!(events[1]["status"], "completed");
    assert_eq!(events[1]["feature"], "S3-F02");
    assert_eq!(events[1]["reader"], "calamine");
    assert_eq!(events[1]["writer"], "rust_xlsxwriter");
    assert!(events[1]["comparison_read_time_ms"].is_number());
    assert!(events[1]["workbook_write_time_ms"].is_number());
    assert_eq!(fs::read(&comparison).unwrap(), comparison_before);
    assert_eq!(fs::read(&config).unwrap(), config_before);

    fs::remove_dir_all(root).unwrap();
}
