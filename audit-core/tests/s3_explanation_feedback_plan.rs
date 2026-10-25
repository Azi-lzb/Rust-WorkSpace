use excel_reader::{read_workbook, CellValue};
use excel_writer::Workbook;
use serde_json::Value;
use std::{
    fs,
    path::{Path, PathBuf},
    process::{Command, Output},
    time::{SystemTime, UNIX_EPOCH},
};

fn temporary_root() -> PathBuf {
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system time after Unix epoch")
        .as_nanos();
    let root = std::env::temp_dir().join(format!(
        "audit_s3_explanation_feedback_{}_{}",
        std::process::id(),
        nonce
    ));
    fs::create_dir_all(&root).expect("create test directory");
    root
}

fn write_target(path: &PathBuf) {
    let mut workbook = Workbook::new();
    let cover = workbook.add_worksheet();
    cover.set_name("封面").unwrap();
    cover.write_string(0, 0, "机构说明导入样本").unwrap();
    let sheet = workbook.add_worksheet();
    sheet.set_name("说明文件").unwrap();
    sheet
        .merge_range(0, 0, 0, 7, "xx说明", &Default::default())
        .unwrap();
    let headers = [
        "机构代码",
        "机构名称",
        "指标代码",
        "指标名称",
        "本期金额（亿元）",
        "上期值",
        "增减额（亿元）",
        "变动原因",
    ];
    for (column, header) in headers.iter().enumerate() {
        sheet.write_string(1, column as u16, *header).unwrap();
    }
    write_target_row(
        sheet,
        2,
        "6k0u",
        "目标机构",
        "A001",
        "指标甲",
        0.418462,
        -0.0002,
    );
    write_target_row(sheet, 3, "DUP", "重复机构", "D1", "重复指标", 1.0, 2.0);
    write_target_row(sheet, 4, "DUP", "重复机构", "D1", "重复指标", 1.0, 2.0);
    write_target_row(sheet, 5, "MULTI", "冲突机构", "M1", "冲突指标", 3.0, 4.0);
    for row in 2..6 {
        sheet.write_string(row, 7, "待机构填写").unwrap();
    }
    workbook.save(path).unwrap();
}

fn write_target_row(
    sheet: &mut excel_writer::Worksheet,
    row: u32,
    organization_code: &str,
    organization_name: &str,
    indicator_code: &str,
    indicator_name: &str,
    data_value: f64,
    change_amount: f64,
) {
    for (column, value) in [
        organization_code,
        organization_name,
        indicator_code,
        indicator_name,
    ]
    .iter()
    .enumerate()
    {
        sheet.write_string(row, column as u16, *value).unwrap();
    }
    sheet.write_number(row, 4, data_value).unwrap();
    sheet.write_number(row, 5, 1.0).unwrap();
    sheet.write_number(row, 6, change_amount).unwrap();
}

fn write_feedback(path: &PathBuf, rows: &[[&str; 8]]) {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("机构反馈").unwrap();
    let headers = [
        "机构代码",
        "机构名称",
        "指标代码",
        "指标名称",
        "数据值",
        "上期值",
        "增减额（亿元）",
        "变动原因",
    ];
    for (column, header) in headers.iter().enumerate() {
        sheet.write_string(0, column as u16, *header).unwrap();
    }
    for (row, values) in rows.iter().enumerate() {
        for (column, value) in values.iter().enumerate() {
            sheet
                .write_string(row as u32 + 1, column as u16, *value)
                .unwrap();
        }
    }
    workbook.save(path).unwrap();
}

fn run_plan(
    target: &PathBuf,
    feedback_paths: &[PathBuf],
    candidate: &PathBuf,
    plan: &PathBuf,
    developer_view: bool,
) -> Output {
    let mut command = Command::new(env!("CARGO_BIN_EXE_open-data-audit"));
    command
        .arg("s3-plan-explanation-feedback-import")
        .arg("--target-workbook")
        .arg(target);
    for feedback in feedback_paths {
        command.arg("--feedback-workbook").arg(feedback);
    }
    command
        .arg("--candidate-output-workbook")
        .arg(candidate)
        .arg("--plan-output")
        .arg(plan);
    if developer_view {
        command.arg("--developer-view");
    }
    command.output().expect("run S3-F03 planning CLI")
}

fn text_cell(workbook: &excel_reader::WorkbookDump, sheet: usize, address: &str) -> String {
    match workbook.sheets[sheet]
        .values
        .get(address)
        .map(|value| CellValue::from_probe_json(value).expect("decode cell"))
    {
        Some(CellValue::Text(value)) => value,
        other => panic!("expected text in {address}, got {other:?}"),
    }
}

#[test]
fn s3_f03_plan_matches_python_identity_rules_and_keeps_inputs_unchanged() {
    let root = temporary_root();
    let target = root.join("指标说明.xlsx");
    let feedback_a = root.join("机构反馈A.xlsx");
    let feedback_b = root.join("机构反馈B.xlsx");
    let candidate = root.join("候选回写.xlsx");
    let plan_path = root.join("导入计划.json");
    write_target(&target);
    write_feedback(
        &feedback_a,
        &[
            [
                "'6K0U",
                "名称有差异",
                "a001",
                "指标名称也有差异",
                "0.4185",
                "1.0",
                "-0.00020",
                "本期业务正常增加",
            ],
            [
                "DUP",
                "重复机构",
                "D1",
                "重复指标",
                "1.0",
                "1.0",
                "2.0",
                "重复目标原因",
            ],
            [
                "MULTI",
                "冲突机构",
                "M1",
                "冲突指标",
                "3.0",
                "1.0",
                "4.0",
                "原因甲",
            ],
            [
                "MISSING",
                "未命中机构",
                "U1",
                "未命中指标",
                "9",
                "1",
                "8",
                "无目标原因",
            ],
            [
                "IGNORED",
                "忽略机构",
                "I1",
                "忽略指标",
                "1",
                "1",
                "0",
                "   ",
            ],
        ],
    );
    write_feedback(
        &feedback_b,
        &[[
            "MULTI",
            "冲突机构",
            "M1",
            "冲突指标",
            "3.0000",
            "1.0",
            "4.0000",
            "原因乙",
        ]],
    );
    let target_before = fs::read(&target).unwrap();
    let feedback_a_before = fs::read(&feedback_a).unwrap();
    let feedback_b_before = fs::read(&feedback_b).unwrap();

    let output = run_plan(
        &target,
        &[feedback_a.clone(), feedback_b.clone()],
        &candidate,
        &plan_path,
        true,
    );
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    let summary: Value = serde_json::from_slice(&output.stdout).unwrap();
    assert_eq!(summary["status"], "planned_not_applied");
    assert_eq!(summary["applied"], false);
    assert_eq!(summary["imported"], 1);
    assert_eq!(summary["unmatched"], 1);
    assert_eq!(summary["conflicts"], 2);
    assert_eq!(summary["planned_cell_updates"], 1);
    assert_eq!(summary["reader"], "calamine");
    assert_eq!(summary["business_engine"], "audit-core (Rust)");
    assert_eq!(summary["conflict_writer"], "rust_xlsxwriter");
    assert!(
        !candidate.exists(),
        "planning must not create a candidate workbook"
    );
    assert_eq!(fs::read(&target).unwrap(), target_before);
    assert_eq!(fs::read(&feedback_a).unwrap(), feedback_a_before);
    assert_eq!(fs::read(&feedback_b).unwrap(), feedback_b_before);

    let plan: Value = serde_json::from_slice(&fs::read(&plan_path).unwrap()).unwrap();
    assert_eq!(plan["feature"], "S3-F03");
    assert_eq!(plan["counts"]["imported"], 1);
    assert_eq!(plan["counts"]["unmatched"], 1);
    assert_eq!(plan["counts"]["conflicts"], 2);
    let operations = plan["edit_plan"]["operations"].as_array().unwrap();
    assert_eq!(operations.len(), 1);
    assert_eq!(operations[0]["operation"], "set_cell_value");
    assert_eq!(operations[0]["sheet"], "说明文件");
    assert_eq!(operations[0]["cell"], "H3");
    assert_eq!(operations[0]["value"]["value"], "本期业务正常增加");

    let conflict_path = summary["conflict_report_path"].as_str().unwrap();
    assert!(Path::new(conflict_path)
        .file_name()
        .unwrap()
        .to_string_lossy()
        .starts_with("指标说明_导入冲突清单_"));
    let conflicts = read_workbook(conflict_path).unwrap();
    assert_eq!(conflicts.sheets[0].name, "冲突明细");
    assert_eq!(text_cell(&conflicts, 0, "A1"), "状态");
    assert_eq!(text_cell(&conflicts, 0, "A2"), "冲突");
    assert_eq!(text_cell(&conflicts, 0, "H2"), "4、5");
    assert!(text_cell(&conflicts, 0, "J2").contains("目标说明文件存在 2 行相同标识"));
    assert_eq!(text_cell(&conflicts, 0, "A3"), "冲突");
    assert!(text_cell(&conflicts, 0, "I3").contains("原因甲"));
    assert!(text_cell(&conflicts, 0, "I3").contains("原因乙"));
    assert!(text_cell(&conflicts, 0, "J3").contains("多个反馈文件给出不同变动原因"));

    let logs = String::from_utf8_lossy(&output.stderr);
    assert!(logs.contains("\"feature\":\"S3-F03\""));
    assert!(logs.contains("\"status\":\"completed\""));
    let _ = fs::remove_dir_all(root);
}

#[test]
fn s3_f03_user_view_logs_engines_and_stage_timings() {
    let root = temporary_root();
    let target = root.join("target.xlsx");
    let feedback = root.join("feedback.xlsx");
    let candidate = root.join("candidate.xlsx");
    let plan = root.join("plan.json");
    write_target(&target);
    write_feedback(
        &feedback,
        &[[
            "6K0U",
            "名称不同",
            "A001",
            "指标名称不同",
            "0.4185",
            "1",
            "-0.0002",
            "原因",
        ]],
    );
    let target_before = fs::read(&target).unwrap();

    let output = run_plan(&target, &[feedback], &candidate, &plan, false);
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    let logs = String::from_utf8_lossy(&output.stderr);
    assert!(logs.contains("读取器 calamine"));
    assert!(logs.contains("业务引擎 audit-core (Rust)"));
    assert!(logs.contains("无需生成冲突清单"));
    assert!(logs.contains("目标读取"));
    assert!(logs.contains("总耗时"));
    assert_eq!(fs::read(&target).unwrap(), target_before);
    assert!(!candidate.exists());
    assert!(plan.exists());
    let _ = fs::remove_dir_all(root);
}

#[test]
fn s3_f03_plan_rejects_existing_outputs_and_bad_feedback_without_mutating_target() {
    let root = temporary_root();
    let target = root.join("target.xlsx");
    let feedback = root.join("feedback.xlsx");
    let candidate = root.join("candidate.xlsx");
    let plan = root.join("plan.json");
    write_target(&target);
    write_feedback(
        &feedback,
        &[[
            "6k0u", "机构", "A001", "指标", "0.418462", "1", "-0.0002", "原因",
        ]],
    );
    let existing_output = root.join("exists.json");
    fs::write(&existing_output, b"keep").unwrap();
    let target_before = fs::read(&target).unwrap();

    let output = run_plan(&target, &[feedback], &candidate, &existing_output, false);
    assert!(!output.status.success());
    assert_eq!(fs::read(&existing_output).unwrap(), b"keep");
    assert_eq!(fs::read(&target).unwrap(), target_before);
    assert!(!candidate.exists());
    assert!(!plan.exists());
    let _ = fs::remove_dir_all(root);
}
