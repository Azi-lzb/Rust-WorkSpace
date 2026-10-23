use excel_writer::Workbook;
use serde_json::{json, Value};
use std::{
    fs,
    path::{Path, PathBuf},
    process::{Command, Output},
    time::{SystemTime, UNIX_EPOCH},
};

const HISTORY_HEADERS: [&str; 15] = [
    "工作簿名",
    "工作表名",
    "定位单元格",
    "错误类型",
    "校验类型",
    "校验指标",
    "描述",
    "当前值",
    "对比值",
    "差值",
    "规则编号",
    "历史校验说明",
    "审核意见",
    "扩展列",
    "数据期",
];

fn temporary_path(label: &str, extension: &str) -> PathBuf {
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system clock after Unix epoch")
        .as_nanos();
    std::env::temp_dir().join(format!(
        "audit_s1_history_workbook_{label}_{}_{}.{}",
        std::process::id(),
        nonce,
        extension
    ))
}

fn write_history_sheet(
    workbook: &mut Workbook,
    sheet_name: &str,
    first_rule_id: &str,
    first_period: &str,
) {
    let sheet = workbook.add_worksheet();
    sheet.set_name(sheet_name).expect("set history sheet name");
    for (column, header) in HISTORY_HEADERS.iter().enumerate() {
        sheet
            .write_string(0, column as u16, *header)
            .expect("write history header");
    }

    write_history_row(
        sheet,
        1,
        "机构一.xlsx",
        "汇总表",
        "D8",
        "核实",
        "数值",
        "资产总额",
        "金额不一致",
        123.5,
        100.0,
        23.5,
        first_rule_id,
        "机构历史说明",
        "审核历史意见",
        first_period,
    );

    write_history_row(
        sheet,
        2,
        "机构一.xlsx",
        "汇总表",
        "E9",
        "错误",
        "数值",
        "存款余额",
        "历史问题已整改",
        50.0,
        40.0,
        10.0,
        "RULE-RESOLVED",
        "",
        "",
        first_period,
    );

    // This row has no identity columns and must be skipped by the Python contract.
    sheet
        .write_string(3, 5, "仅有指标的空身份行")
        .expect("write skippable row");
}

#[allow(clippy::too_many_arguments)]
fn write_history_row(
    sheet: &mut excel_writer::Worksheet,
    row: u32,
    source_file: &str,
    sheet_name: &str,
    target_cell: &str,
    severity: &str,
    check_type: &str,
    indicator: &str,
    description: &str,
    current_value: f64,
    comparison_value: f64,
    difference_value: f64,
    rule_id: &str,
    institution_feedback: &str,
    auditor_opinion: &str,
    period: &str,
) {
    for (column, value) in [
        (0, source_file),
        (1, sheet_name),
        (2, target_cell),
        (3, severity),
        (4, check_type),
        (5, indicator),
        (6, description),
        (10, rule_id),
        (11, institution_feedback),
        (12, auditor_opinion),
        (14, period),
    ] {
        sheet
            .write_string(row, column, value)
            .expect("write history cell");
    }
    sheet
        .write_number(row, 7, current_value)
        .expect("write current value");
    sheet
        .write_number(row, 8, comparison_value)
        .expect("write comparison value");
    sheet
        .write_number(row, 9, difference_value)
        .expect("write difference value");
}

fn history_workbook(include_standard: bool, include_custom: bool, include_legacy: bool) -> PathBuf {
    let path = temporary_path("fixture", "xlsx");
    let mut workbook = Workbook::new();
    if include_standard {
        write_history_sheet(&mut workbook, "核查表校验结果", " RULE-CONT ", "2026-02");
    }
    if include_custom {
        write_history_sheet(&mut workbook, "自定义历史", "RULE-CUSTOM", "2026-01");
    }
    if include_legacy {
        write_history_sheet(&mut workbook, "历史审核结果", "RULE-LEGACY", "2026-01");
    }
    workbook.save(&path).expect("save history workbook");
    path
}

fn current_issue(issue_id: &str) -> Value {
    json!({
        "issue_id": issue_id,
        "period": "2026-03",
        "batch_id": "batch-new",
        "audit_time": "2026-10-24 12:00:00",
        "triggered": true,
        "status": "",
        "first_seen_period": "",
        "previous_seen_period": "",
        "consecutive_count": 1,
        "org_code": "001",
        "org_name": "测试机构",
        "report_code": "R01",
        "sheet_name": "汇总表",
        "rule_id": issue_id,
        "severity": "错误",
        "formula_cell": "C8",
        "target_cell": "D8",
        "target_value": 125.0,
        "formula_result": "错误",
        "message": "当前错误",
        "source_file": "机构一.xlsx",
        "audit_file": "机构一_审核版.xlsx"
    })
}

fn input_file(current_issues: Vec<Value>) -> PathBuf {
    let path = temporary_path("input", "json");
    fs::write(
        &path,
        serde_json::to_vec(&json!({
            "period": "2026-03",
            "batch_id": "batch-new",
            "audit_time": "2026-10-24 12:00:00",
            "current_issues": current_issues,
            // Workbook input must replace JSON history when the workbook option is present.
            "history": [current_issue("JSON-DECOY")]
        }))
        .expect("serialize history request"),
    )
    .expect("write history request");
    path
}

fn run_cli(input: &Path, workbook: &Path, sheet: Option<&str>, developer_view: bool) -> Output {
    let mut command = Command::new(env!("CARGO_BIN_EXE_open-data-audit"));
    command
        .arg("s1-history-classify")
        .arg("--input")
        .arg(input)
        .arg("--history-workbook")
        .arg(workbook);
    if let Some(sheet) = sheet {
        command.arg("--history-sheet").arg(sheet);
    }
    if developer_view {
        command.arg("--developer-view");
    }
    command.output().expect("run S1 history CLI")
}

fn successful_json(output: &Output) -> Value {
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    serde_json::from_slice(&output.stdout).expect("parse S1 history CLI JSON")
}

#[test]
fn cli_reads_standard_history_classifies_and_reports_both_log_views() {
    let workbook = history_workbook(true, true, true);
    let input = input_file(vec![current_issue(" RULE-CONT ")]);
    let ordinary = run_cli(&input, &workbook, None, false);
    let developer = run_cli(&input, &workbook, None, true);
    let fallback = run_cli(&input, &workbook, Some("不存在的工作表"), false);
    fs::remove_file(&workbook).expect("remove synthetic history workbook");
    fs::remove_file(&input).expect("remove synthetic history request");
    let ordinary_result = successful_json(&ordinary);
    let developer_result = successful_json(&developer);
    let fallback_result = successful_json(&fallback);

    assert_eq!(ordinary_result["history_workbook_read"]["status"], "loaded");
    assert_eq!(
        ordinary_result["history_workbook_read"]["reader"],
        "calamine"
    );
    assert_eq!(
        ordinary_result["history_workbook_read"]["sheet_name"],
        "核查表校验结果"
    );
    assert_eq!(ordinary_result["history_workbook_read"]["rows_total"], 3);
    assert_eq!(ordinary_result["history_workbook_read"]["rows_read"], 2);
    assert_eq!(ordinary_result["history_workbook_read"]["rows_skipped"], 1);
    assert_eq!(ordinary_result["summary"]["history_rows"], 2);
    assert_eq!(ordinary_result["current_issues"][0]["status"], "连续出现");
    assert_eq!(
        ordinary_result["current_issues"][0]["institution_feedback"],
        "机构历史说明"
    );
    assert_eq!(
        ordinary_result["current_issues"][0]["auditor_opinion"],
        "审核历史意见"
    );
    assert_eq!(
        ordinary_result["current_issues"][0]["previous_seen_period"],
        "2026-02"
    );
    assert_eq!(
        ordinary_result["resolved_issues"].as_array().unwrap().len(),
        1
    );
    assert_eq!(
        ordinary_result["resolved_issues"][0]["rule_id"],
        "RULE-RESOLVED"
    );
    assert_eq!(ordinary_result["resolved_issues"][0]["status"], "已整改");
    assert_eq!(
        developer_result["history_workbook_read"]["sheet_name"],
        "核查表校验结果"
    );
    assert_eq!(
        fallback_result["history_workbook_read"]["sheet_name"],
        "核查表校验结果"
    );
    assert!(String::from_utf8_lossy(&ordinary.stderr).contains("引擎 calamine"));
    assert!(String::from_utf8_lossy(&ordinary.stderr).contains("耗时"));
    let developer_events = String::from_utf8_lossy(&developer.stderr);
    assert!(developer_events.contains("history_workbook_read"));
    assert!(developer_events.contains("\"engine\":\"calamine\""));
    assert!(developer_events.contains("\"elapsed_ms\""));
}

#[test]
fn cli_honors_requested_sheet_and_falls_back_to_legacy_sheet() {
    let workbook = history_workbook(true, true, true);
    let input = input_file(Vec::new());
    let requested_output = run_cli(&input, &workbook, Some("自定义历史"), false);
    fs::remove_file(&workbook).expect("remove synthetic workbook");
    let requested = successful_json(&requested_output);

    assert_eq!(
        requested["history_workbook_read"]["sheet_name"],
        "自定义历史"
    );
    assert_eq!(requested["merged_history"][0]["rule_id"], "RULE-CUSTOM");

    let legacy_workbook = history_workbook(false, false, true);
    let legacy_output = run_cli(&input, &legacy_workbook, None, false);
    fs::remove_file(&legacy_workbook).expect("remove legacy workbook");
    fs::remove_file(&input).expect("remove synthetic history request");
    let legacy = successful_json(&legacy_output);

    assert_eq!(
        legacy["history_workbook_read"]["sheet_name"],
        "历史审核结果"
    );
    assert_eq!(legacy["merged_history"][0]["rule_id"], "RULE-LEGACY");
}

#[test]
fn cli_reports_missing_workbook_and_missing_history_sheet_as_empty_history() {
    let missing_workbook = temporary_path("missing", "xlsx");
    let input = input_file(Vec::new());
    let missing_output = run_cli(&input, &missing_workbook, None, false);
    let missing_developer_output = run_cli(&input, &missing_workbook, None, true);
    fs::remove_file(&input).expect("remove synthetic history request");
    let missing = successful_json(&missing_output);
    let missing_developer = successful_json(&missing_developer_output);

    assert_eq!(missing["history_workbook_read"]["status"], "missing_file");
    assert_eq!(missing["history_workbook_read"]["reader"], "none");
    assert_eq!(missing["summary"]["history_rows"], 0);
    assert!(String::from_utf8_lossy(&missing_output.stderr).contains("未调用 Excel 读取引擎"));
    assert!(
        String::from_utf8_lossy(&missing_developer_output.stderr).contains("\"engine\":\"none\"")
    );
    assert!(missing_developer["history_workbook_read"]["reader"] == "none");

    let unrelated_workbook = temporary_path("unrelated", "xlsx");
    let mut workbook = Workbook::new();
    workbook
        .add_worksheet()
        .set_name("其他表")
        .expect("set unrelated worksheet name");
    workbook
        .save(&unrelated_workbook)
        .expect("save unrelated workbook");
    let input = input_file(Vec::new());
    let output = run_cli(&input, &unrelated_workbook, None, false);
    fs::remove_file(&input).expect("remove synthetic history request");
    fs::remove_file(&unrelated_workbook).expect("remove unrelated workbook");
    let missing_sheet = successful_json(&output);

    assert_eq!(
        missing_sheet["history_workbook_read"]["status"],
        "sheet_missing"
    );
    assert_eq!(missing_sheet["history_workbook_read"]["reader"], "calamine");
    assert_eq!(missing_sheet["summary"]["history_rows"], 0);
}
