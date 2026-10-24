use excel_writer::Workbook;
use serde_json::Value;
use std::{
    fs,
    path::PathBuf,
    process::{Command, Output},
    time::{SystemTime, UNIX_EPOCH},
};

fn temporary_workbook_path() -> PathBuf {
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system clock after Unix epoch")
        .as_nanos();
    std::env::temp_dir().join(format!(
        "audit_s1_template_rules_cli_{}_{}.xlsx",
        std::process::id(),
        nonce
    ))
}

fn template_with_formula_rules() -> PathBuf {
    let path = temporary_workbook_path();
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .set_name("检查表")
        .expect("set formula worksheet name");
    worksheet
        .write_formula(0, 0, "=IF(TRUE,\"错误|类型|指标|描述\",\"通过\")")
        .expect("write formula");
    workbook
        .define_name("校验区域", "='检查表'!$A$1")
        .expect("define formula range");
    workbook
        .define_name("表结构区域", "='检查表'!$B$1")
        .expect("define structure range");
    workbook.save(&path).expect("save formula template");
    path
}

fn run_cli(path: &PathBuf, developer_view: bool, include_definition: bool) -> Output {
    let mut command = Command::new(env!("CARGO_BIN_EXE_open-data-audit"));
    command
        .arg("s1-template-rules")
        .arg("--template-workbook")
        .arg(path);
    if developer_view {
        command.arg("--developer-view");
    }
    if include_definition {
        command.arg("--include-rule-definitions");
    }
    command.output().expect("run S1 template-rule CLI")
}

#[test]
fn cli_returns_sanitized_summary_and_both_log_views() {
    let template = template_with_formula_rules();
    let ordinary = run_cli(&template, false, false);
    let developer = run_cli(&template, true, true);
    fs::remove_file(&template).expect("remove synthetic template");

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
        serde_json::from_slice(&ordinary.stdout).expect("parse ordinary CLI output");
    assert_eq!(ordinary_output["status"], "completed");
    assert_eq!(ordinary_output["summary"]["rules_total"], 1);
    assert_eq!(ordinary_output["template_health"]["passed"], true);
    assert_eq!(ordinary_output["template_health"]["result_rules"], 1);
    assert_eq!(ordinary_output["template_health"]["structure_ranges"], 1);
    assert!(ordinary_output.get("definition").is_none());
    let ordinary_log = String::from_utf8_lossy(&ordinary.stderr);
    assert!(ordinary_log.contains("S1 模板规则检查完成"));
    assert!(ordinary_log.contains("calamine+namespace_aware_ooxml"));

    let developer_output: Value =
        serde_json::from_slice(&developer.stdout).expect("parse developer CLI output");
    assert_eq!(
        developer_output["definition"]["rules"][0]["formula_cell"],
        "A1"
    );
    let developer_events = String::from_utf8_lossy(&developer.stderr)
        .lines()
        .map(|line| serde_json::from_str::<Value>(line).expect("parse progress event"))
        .collect::<Vec<_>>();
    assert_eq!(developer_events.len(), 2);
    assert_eq!(developer_events[0]["status"], "started");
    assert_eq!(developer_events[1]["status"], "completed");
    assert_eq!(
        developer_events[1]["summary"]["engine"],
        "calamine+namespace_aware_ooxml"
    );
    assert_eq!(developer_events[1]["template_health"]["passed"], true);
    assert_eq!(developer_output["template_health"]["passed"], true);
}
