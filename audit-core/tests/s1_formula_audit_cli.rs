use excel_writer::Workbook;
use serde_json::Value;
use std::{
    fs,
    path::PathBuf,
    process::{Command, Output},
    time::{SystemTime, UNIX_EPOCH},
};

fn temporary_workbook_path(label: &str) -> PathBuf {
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system clock after Unix epoch")
        .as_nanos();
    std::env::temp_dir().join(format!(
        "audit_s1_formula_audit_cli_{label}_{}_{}.xlsx",
        std::process::id(),
        nonce
    ))
}

fn formula_template() -> PathBuf {
    let path = temporary_workbook_path("template");
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_name("检查表").expect("name rule sheet");
    for (row, formula) in [
        "=IF(TRUE,\"错误|复核|存款|低于要求|100|200|100\",\"\")",
        "=IF(TRUE,\"通过\",\"错误\")",
        "=IF(TRUE,\"\",\"\")",
        "=IF(TRUE,\"\",\"\")",
        "=IF(TRUE,\"错误|字段|指标|描述\",\"通过\")",
    ]
    .iter()
    .enumerate()
    {
        worksheet
            .write_formula(row as u32, 0, *formula)
            .expect("write formula rule");
    }
    workbook
        .define_name("校验区域", "='检查表'!$A$1:$A$5")
        .expect("define formula rules");
    workbook.save(&path).expect("save formula template");
    path
}

fn audit_copy_with_stored_results() -> PathBuf {
    let path = temporary_workbook_path("results");
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_name("检查表").expect("name result sheet");
    for (row, value) in [
        "错误|复核|存款|低于要求|100|200|100",
        "通过",
        "#REF!",
        "通过",
        "通过",
    ]
    .iter()
    .enumerate()
    {
        worksheet
            .write_string(row as u32, 0, *value)
            .expect("write stored formula result");
    }
    workbook.save(&path).expect("save audit copy values");
    path
}

fn run_cli(template: &PathBuf, audit: &PathBuf, developer_view: bool) -> Output {
    let mut command = Command::new(env!("CARGO_BIN_EXE_open-data-audit"));
    command
        .arg("s1-formula-audit")
        .arg("--template-workbook")
        .arg(template)
        .arg("--audit-workbook")
        .arg(audit)
        .arg("--period")
        .arg("2026-07")
        .arg("--batch-id")
        .arg("batch-01")
        .arg("--audit-time")
        .arg("2026-10-24 12:00:00")
        .arg("--org-code")
        .arg("1234")
        .arg("--org-name")
        .arg("机构甲");
    if developer_view {
        command.arg("--developer-view");
    }
    command.output().expect("run S1 formula audit CLI")
}

#[test]
fn cli_extracts_cached_formula_issues_with_user_and_developer_logs() {
    let template = formula_template();
    let audit = audit_copy_with_stored_results();
    let ordinary = run_cli(&template, &audit, false);
    let developer = run_cli(&template, &audit, true);
    fs::remove_file(&template).expect("remove formula template");
    fs::remove_file(&audit).expect("remove audit copy");

    assert!(
        ordinary.status.success(),
        "{}",
        String::from_utf8_lossy(&ordinary.stderr)
    );
    assert!(
        developer.status.success(),
        "{}",
        String::from_utf8_lossy(&developer.stderr)
    );

    let ordinary_output: Value =
        serde_json::from_slice(&ordinary.stdout).expect("parse formula audit result");
    assert_eq!(ordinary_output["status"], "completed");
    assert_eq!(ordinary_output["template_summary"]["rules_total"], 5);
    assert_eq!(ordinary_output["summary"]["reader"], "calamine");
    assert_eq!(ordinary_output["summary"]["rules_selected"], 5);
    assert_eq!(
        ordinary_output["summary"]["formula_result_values_nonempty"],
        5
    );
    assert_eq!(ordinary_output["summary"]["issue_count"], 2);
    assert_eq!(ordinary_output["issues"][0]["formula_cell"], "A1");
    assert_eq!(ordinary_output["issues"][0]["check_field"], "存款");
    assert_eq!(ordinary_output["issues"][0]["target_value"], "100");
    assert_eq!(ordinary_output["issues"][1]["formula_cell"], "A3");
    assert_eq!(ordinary_output["issues"][1]["formula_result"], "#REF!");
    let ordinary_log = String::from_utf8_lossy(&ordinary.stderr);
    assert!(ordinary_log.contains("S1 公式结果提取完成"));
    assert!(ordinary_log.contains("读取器 calamine"));
    assert!(ordinary_log.contains("未执行公式重算"));
    assert!(!ordinary_log.contains("低于要求"));

    let developer_output: Value =
        serde_json::from_slice(&developer.stdout).expect("parse developer formula audit result");
    assert_eq!(developer_output["issues"], ordinary_output["issues"]);
    let developer_events = String::from_utf8_lossy(&developer.stderr)
        .lines()
        .map(|line| serde_json::from_str::<Value>(line).expect("parse developer event"))
        .collect::<Vec<_>>();
    assert_eq!(developer_events.len(), 2);
    assert_eq!(developer_events[0]["status"], "started");
    assert_eq!(developer_events[0]["result_reader"], "calamine");
    assert_eq!(developer_events[1]["status"], "completed");
    assert_eq!(developer_events[1]["rules_total"], 5);
    assert_eq!(developer_events[1]["issues"], 2);
    assert!(developer_events[1]["total_time_ms"].is_number());
    assert!(!String::from_utf8_lossy(&developer.stderr).contains("低于要求"));
}
