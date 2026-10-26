use excel_reader::read_workbook;
use excel_writer::Workbook;
use serde_json::{Value, json};
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

fn temporary_directory(label: &str) -> PathBuf {
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system clock after Unix epoch")
        .as_nanos();
    let path = std::env::temp_dir().join(format!(
        "audit_s1_issue_chain_{label}_{}_{}",
        std::process::id(),
        nonce
    ));
    fs::create_dir_all(&path).expect("create temporary directory");
    path
}

fn formula_template(path: &Path) {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_name("检查表").expect("name formula sheet");
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
    let structure_sheet = workbook.add_worksheet();
    structure_sheet
        .set_name("S1Conditional")
        .expect("name structure sheet");
    structure_sheet
        .write_string(1, 0, "资产总额")
        .expect("write structure label");
    workbook
        .define_name("表结构区域_集成样例", "='S1Conditional'!$A$2:$A$2")
        .expect("define structure area");
    workbook.save(path).expect("save formula template");
}

fn formula_template_without_result_rule_or_structure_area(path: &Path) {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_name("检查表").expect("name formula sheet");
    worksheet
        .write_formula(0, 0, "=1+1")
        .expect("write error-only formula");
    workbook
        .define_name("校验区域", "='检查表'!$A$1")
        .expect("define formula range");
    workbook
        .save(path)
        .expect("save unhealthy formula template");
}

fn source_workbook_with_validation_error(path: &Path) {
    let mut workbook = Workbook::new();
    let validation_sheet = workbook.add_worksheet();
    validation_sheet
        .set_name("检查表")
        .expect("name source validation sheet");
    validation_sheet
        .write_string(3, 0, "#DIV/0!")
        .expect("write in-range source error");
    validation_sheet
        .write_string(0, 1, "#REF!")
        .expect("write out-of-range source error");
    let structure_sheet = workbook.add_worksheet();
    structure_sheet
        .set_name("S1Conditional")
        .expect("name source structure sheet");
    structure_sheet
        .write_string(1, 0, "资产总额")
        .expect("write source structure label");
    workbook.save(path).expect("save source workbook");
}

fn audit_copy(path: &Path) {
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
    workbook.save(path).expect("save calculated audit copy");
}

fn history_workbook(path: &Path, continuing_issue_id: &str) {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet
        .set_name("核查表校验结果")
        .expect("name S1 history sheet");
    for (column, header) in HISTORY_HEADERS.iter().enumerate() {
        sheet
            .write_string(0, column as u16, *header)
            .expect("write S1 history header");
    }
    for (row, issue_id, period, feedback, opinion) in [
        (
            1,
            continuing_issue_id,
            "2026-06",
            "历史机构反馈",
            "历史审核意见",
        ),
        (
            2,
            "history-only-issue",
            "2026-06",
            "已整改项反馈",
            "已整改项意见",
        ),
    ] {
        for (column, value) in [
            (0, "1001_机构甲_2026-07.xlsx"),
            (1, "检查表"),
            (2, "A1"),
            (3, "错误"),
            (4, "数值"),
            (5, "存款余额"),
            (6, "历史问题"),
            (10, issue_id),
            (11, feedback),
            (12, opinion),
            (14, period),
        ] {
            sheet
                .write_string(row, column as u16, value)
                .expect("write S1 history row");
        }
        sheet
            .write_number(row, 7, 123.0)
            .expect("write history current value");
        sheet
            .write_number(row, 8, 100.0)
            .expect("write history comparison value");
        sheet
            .write_number(row, 9, 23.0)
            .expect("write history difference value");
    }
    workbook.save(path).expect("save S1 history workbook");
}

fn run_cli(input: &Path, template: &Path, manifest: &Path, developer_view: bool) -> Output {
    run_cli_with_history(input, template, manifest, None, developer_view)
}

fn run_cli_with_history(
    input: &Path,
    template: &Path,
    manifest: &Path,
    history: Option<&Path>,
    developer_view: bool,
) -> Output {
    run_cli_with_options(
        input,
        Some(template),
        None,
        manifest,
        history,
        None,
        None,
        developer_view,
    )
}

fn run_cli_with_options(
    input: &Path,
    template: Option<&Path>,
    template_dir: Option<&Path>,
    manifest: &Path,
    history: Option<&Path>,
    result_workbook: Option<&Path>,
    history_result_workbook: Option<&Path>,
    developer_view: bool,
) -> Output {
    let mut command = Command::new(env!("CARGO_BIN_EXE_open-data-audit"));
    command
        .arg("s1-audit-directory")
        .arg("--input-dir")
        .arg(input)
        .arg("--formula-manifest")
        .arg(manifest)
        .arg("--period")
        .arg("2026-07")
        .arg("--batch-id")
        .arg("issue-chain-test")
        .arg("--allow-saved-formula-cache-diagnostics")
        .arg("--audit-time")
        .arg("2026-10-24 12:00:00");
    if let Some(template) = template {
        command.arg("--template-workbook").arg(template);
    }
    if let Some(template_dir) = template_dir {
        command.arg("--template-dir").arg(template_dir);
    }
    if let Some(history) = history {
        command.arg("--history-workbook").arg(history);
    }
    if let Some(result_workbook) = result_workbook {
        command.arg("--result-workbook").arg(result_workbook);
    }
    if let Some(history_result_workbook) = history_result_workbook {
        command
            .arg("--history-result-workbook")
            .arg(history_result_workbook);
    }
    if developer_view {
        command.arg("--developer-view");
    }
    command.output().expect("run S1 audit issue-chain CLI")
}

#[test]
fn cli_runs_both_issue_sources_merges_then_classifies_and_logs_actual_engines() {
    let directory = temporary_directory("complete");
    let input = directory.join("input");
    fs::create_dir_all(&input).expect("create source input directory");
    let source = input.join("1001_机构甲_2026-07.xlsx");
    fs::copy(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("tests")
            .join("fixtures")
            .join("s1_issue_chain_source.xlsx"),
        &source,
    )
    .expect("copy conditional-format source fixture");
    let template_dir = directory.join("templates");
    fs::create_dir_all(&template_dir).expect("create template directory");
    let template = template_dir.join("template.xlsx");
    formula_template(&template);
    let audit = directory.join("audit-copy.xlsx");
    audit_copy(&audit);
    let manifest = directory.join("formula-copies.json");
    fs::write(
        &manifest,
        serde_json::to_vec(&json!({
            "files": [{
                "source_file": source.file_name().unwrap().to_string_lossy(),
                "audit_workbook": audit.file_name().unwrap().to_string_lossy()
            }]
        }))
        .unwrap(),
    )
    .expect("write formula copy manifest");

    let result_workbook = directory.join("s1-result.xlsx");
    let history_result_workbook = directory.join("s1-history.xlsx");
    let ordinary = run_cli_with_options(
        &input,
        None,
        Some(&template_dir),
        &manifest,
        None,
        Some(&result_workbook),
        Some(&history_result_workbook),
        false,
    );
    assert!(
        ordinary.status.success(),
        "{}",
        String::from_utf8_lossy(&ordinary.stderr)
    );
    let developer_result_workbook = directory.join("s1-result-developer.xlsx");
    let developer_history_result_workbook = directory.join("s1-history-developer.xlsx");
    let developer = run_cli_with_options(
        &input,
        None,
        Some(&template_dir),
        &manifest,
        None,
        Some(&developer_result_workbook),
        Some(&developer_history_result_workbook),
        true,
    );
    assert!(
        developer.status.success(),
        "{}",
        String::from_utf8_lossy(&developer.stderr)
    );

    let report: Value = serde_json::from_slice(&ordinary.stdout).expect("parse issue-chain output");
    assert_eq!(report["status"], "completed_with_diagnostics");
    assert_eq!(
        report["formula_execution"]["status"],
        "saved_cache_diagnostic_only"
    );
    assert_eq!(
        report["formula_execution"]["rust_formula_copy_performed_by_this_command"],
        false
    );
    assert_eq!(report["template_selection"]["mode"], "automatic");
    assert_eq!(
        report["template_selection"]["template_file"],
        "template.xlsx"
    );
    assert_eq!(report["formula_summary"]["files"], 1);
    assert_eq!(report["formula_summary"]["issues"], 2);
    assert_eq!(report["issues"][0]["org_code"], "1001");
    assert_eq!(report["issues"][0]["org_name"], "机构甲");
    assert_eq!(
        report["conditional_summary"]["files"]
            .as_array()
            .unwrap()
            .len(),
        1
    );
    assert_eq!(report["conditional_summary"]["issues"], 1);
    assert_eq!(report["issues"].as_array().unwrap().len(), 4);
    assert_eq!(report["issues"][3]["org_code"], "");
    assert_eq!(report["issues"][3]["org_name"], "");
    assert_eq!(
        report["issues"]
            .as_array()
            .unwrap()
            .iter()
            .map(|issue| issue["formula_cell"].as_str().unwrap())
            .collect::<Vec<_>>(),
        ["A1", "A3", "", "B2"]
    );
    assert_eq!(report["validation_formula_scan"]["status"], "completed");
    assert_eq!(report["validation_formula_scan"]["missing_sheets"], 1);
    assert_eq!(
        report["issues"][2]["rule_id"],
        "TEMPLATE_FORMULA_SCAN:检查表"
    );
    assert_eq!(
        report["issues"][3]["audit_file"],
        audit.canonicalize().unwrap().to_string_lossy().as_ref()
    );
    assert_eq!(report["summary"]["new_issues"], 4);
    assert_eq!(report["resolved_issues"].as_array().unwrap().len(), 0);
    assert_eq!(report["result_workbook"]["issue_count"], 4);
    assert_eq!(report["result_workbook"]["hyperlink_count"], 3);
    assert_eq!(report["history_result_workbook"]["issue_count"], 4);
    assert_eq!(
        report["history_result_workbook"]["path"],
        history_result_workbook
            .canonicalize()
            .unwrap()
            .to_string_lossy()
            .as_ref()
    );
    assert_eq!(
        report["result_workbook"]["path"],
        result_workbook
            .canonicalize()
            .unwrap()
            .to_string_lossy()
            .as_ref()
    );
    let result = read_workbook(&result_workbook).expect("read generated S1 result workbook");
    assert_eq!(result.sheets.len(), 1);
    assert_eq!(result.sheets[0].name, "本期审核结果");
    assert!(result.sheets[0].values.len() > 13);
    assert_eq!(result.sheets[0].values["A1"]["value"], "工作簿名");
    assert_eq!(
        result.sheets[0].values["A2"]["value"],
        source.file_name().unwrap().to_string_lossy().as_ref()
    );
    assert_eq!(result.sheets[0].values["B2"]["value"], "检查表");
    assert_eq!(result.sheets[0].values["C2"]["value"], "A1");
    assert_eq!(
        result.sheets[0].values["K2"]["value"],
        report["issues"][0]["issue_id"]
    );
    let history_result =
        read_workbook(&history_result_workbook).expect("read generated S1 history workbook");
    assert_eq!(history_result.sheets.len(), 1);
    assert_eq!(history_result.sheets[0].name, "核查表校验结果");
    assert_eq!(history_result.sheets[0].values["A1"]["value"], "工作簿名");
    let first_history_issue = &report["merged_history"][0];
    let first_history_source = Path::new(
        first_history_issue["source_file"]
            .as_str()
            .expect("first merged history issue source file"),
    )
    .file_name()
    .expect("first history source filename")
    .to_string_lossy();
    assert_eq!(
        history_result.sheets[0].values["A2"]["value"],
        first_history_source.as_ref()
    );
    assert_eq!(
        history_result.sheets[0].values["K2"]["value"],
        first_history_issue["issue_id"]
    );

    let user_log = String::from_utf8_lossy(&ordinary.stderr);
    assert!(user_log.contains("条件格式审核完成"));
    assert!(user_log.contains("公式结果审核完成"));
    assert!(user_log.contains("S1 校验公式扫描完成"));
    assert!(!user_log.contains("calamine"), "{user_log}");
    assert!(!user_log.contains("namespace_aware_ooxml"), "{user_log}");
    assert!(!user_log.contains("读取器"), "{user_log}");
    assert!(user_log.contains("未执行公式重算"));
    assert!(user_log.contains("本期结果工作簿已生成"));
    assert!(user_log.contains("历史结果工作簿已生成"));
    assert!(user_log.contains("S1-F01 模板推荐完成"));
    assert!(user_log.contains("引擎 namespace-aware OOXML 工作表目录读取"));
    assert!(user_log.contains("耗时"));
    assert!(!user_log.contains("rust_xlsxwriter"), "{user_log}");

    let developer_report: Value =
        serde_json::from_slice(&developer.stdout).expect("parse developer issue-chain output");
    assert_eq!(developer_report["issues"], report["issues"]);
    assert_eq!(developer_report["template_selection"]["mode"], "automatic");
    let events = String::from_utf8_lossy(&developer.stderr)
        .lines()
        .map(|line| serde_json::from_str::<Value>(line).expect("parse developer event"))
        .collect::<Vec<_>>();
    assert!(events.iter().any(|event| {
        event["step"] == "template_recommendation"
            && event["status"] == "completed"
            && event["engine"] == "namespace_aware_ooxml_sheet_directory"
            && event["selected_template"] == "template.xlsx"
            && event["confidence"].is_number()
            && event["elapsed_ms"].is_number()
    }));
    assert!(events.iter().any(|event| {
        event["step"] == "formula_result_issue_extraction"
            && event["engine"] == "calamine"
            && event["issues"] == 2
    }));
    assert!(events.iter().any(|event| {
        event["step"] == "validation_formula_error_scan"
            && event["engine"] == "calamine+namespace_aware_ooxml_defined_names"
            && event["formula_cells_scanned"] == 5
            && event["missing_sheets"] == 1
    }));
    assert!(events.iter().any(|event| {
        event["step"] == "conditional_format_issue_extraction"
            && event["status"] == "completed"
            && event["engine"] == "calamine+namespace_aware_ooxml"
    }));
    assert!(events.iter().any(|event| {
        event["step"] == "issue_merge_and_history_classification"
            && event["current_issues"] == 4
            && event["elapsed_ms"].is_number()
    }));
    assert!(events.iter().any(|event| {
        event["step"] == "result_workbook_write"
            && event["status"] == "completed"
            && event["engine"] == "rust_xlsxwriter 0.101.0"
            && event["issues"] == 4
            && event["hyperlinks"] == 3
            && event["elapsed_ms"].is_number()
    }));
    assert!(events.iter().any(|event| {
        event["step"] == "history_result_workbook_write"
            && event["status"] == "completed"
            && event["engine"] == "rust_xlsxwriter 0.101.0"
            && event["issues"] == 4
            && event["elapsed_ms"].is_number()
    }));

    let replay_history_result = directory.join("s1-history-replay.xlsx");
    let replay = run_cli_with_options(
        &input,
        None,
        Some(&template_dir),
        &manifest,
        Some(&history_result_workbook),
        None,
        Some(&replay_history_result),
        false,
    );
    assert!(
        replay.status.success(),
        "{}",
        String::from_utf8_lossy(&replay.stderr)
    );
    let replay_report: Value =
        serde_json::from_slice(&replay.stdout).expect("parse replayed issue-chain output");
    assert_eq!(
        replay_report["history_workbook_summary"]["status"],
        "loaded"
    );
    assert_eq!(
        replay_report["history_workbook_summary"]["reader"],
        "calamine"
    );
    assert_eq!(replay_report["history_workbook_summary"]["rows_read"], 4);
    assert_eq!(replay_report["summary"]["history_rows"], 4);
    assert_eq!(replay_report["summary"]["merged_history_rows"], 4);
    assert_eq!(replay_report["summary"]["new_issues"], 0);
    assert_eq!(replay_report["summary"]["recurring_issues"], 4);
    assert_eq!(replay_report["summary"]["continuous_issues"], 0);
    assert_eq!(replay_report["issues"].as_array().unwrap().len(), 4);
    assert!(
        replay_report["issues"]
            .as_array()
            .unwrap()
            .iter()
            .all(|issue| issue["status"] == "再次出现")
    );
    assert_eq!(replay_report["history_result_workbook"]["issue_count"], 4);
    let replay_log = String::from_utf8_lossy(&replay.stderr);
    assert!(replay_log.contains("S1 历史工作簿读取完成"));
    assert!(replay_log.contains("历史结果工作簿已生成"));

    if std::env::var_os("AUDIT_S1_KEEP_RESULT_ARTIFACTS").is_some() {
        fs::write(directory.join("rust-result.json"), &ordinary.stdout)
            .expect("save synthetic Rust issues for Python writer parity");
        println!("S1_RESULT_PARITY_DIR={}", directory.display());
    } else {
        fs::remove_dir_all(directory).expect("remove temporary issue-chain workbooks");
    }
}

#[test]
fn cli_scans_template_validation_formula_cells_and_merges_source_errors() {
    let directory = temporary_directory("validation_formula_scan");
    let input = directory.join("input");
    fs::create_dir_all(&input).expect("create source input directory");
    let source = input.join("1001_机构甲_2026-07.xlsx");
    source_workbook_with_validation_error(&source);
    let template = directory.join("template.xlsx");
    formula_template(&template);
    let audit = directory.join("audit-copy.xlsx");
    audit_copy(&audit);
    let result_workbook = directory.join("result.xlsx");
    let history_result_workbook = directory.join("history-result.xlsx");
    let manifest = directory.join("formula-copies.json");
    fs::write(
        &manifest,
        serde_json::to_vec(&json!({
            "files": [{
                "source_file": source.file_name().unwrap().to_string_lossy(),
                "audit_workbook": audit.file_name().unwrap().to_string_lossy()
            }]
        }))
        .unwrap(),
    )
    .expect("write formula copy manifest");

    let output = run_cli_with_options(
        &input,
        Some(&template),
        None,
        &manifest,
        None,
        Some(&result_workbook),
        Some(&history_result_workbook),
        true,
    );
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report: Value = serde_json::from_slice(&output.stdout).expect("parse S1 report");
    assert_eq!(report["validation_formula_scan"]["status"], "completed");
    assert_eq!(
        report["validation_formula_scan"]["formula_cells_scanned"],
        5
    );
    assert_eq!(report["validation_formula_scan"]["formula_errors"], 1);
    assert_eq!(report["validation_formula_scan"]["missing_sheets"], 0);
    assert_eq!(report["validation_formula_scan"]["issues_added"], 1);
    assert_eq!(report["validation_formula_scan"]["failures"], 0);
    assert_eq!(
        report["validation_formula_scan"]["errors"][0]["sheet"],
        "检查表"
    );
    assert_eq!(report["validation_formula_scan"]["errors"][0]["cell"], "A4");
    assert_eq!(
        report["validation_formula_scan"]["errors"][0]["formula"],
        "=IF(TRUE,\"\",\"\")"
    );
    assert_eq!(
        report["validation_formula_scan"]["errors"][0]["origin"],
        "source_cached_value"
    );
    let scanner_issue = report["issues"]
        .as_array()
        .unwrap()
        .iter()
        .find(|issue| issue["check_type"] == "校验公式错误")
        .expect("source validation formula error issue");
    assert_eq!(scanner_issue["formula_cell"], "A4");
    assert_eq!(scanner_issue["formula_result"], "#DIV/0!");
    assert_eq!(
        scanner_issue["source_file"],
        source.canonicalize().unwrap().to_string_lossy().as_ref()
    );
    assert_eq!(
        report["result_workbook"]["issue_count"],
        report["issues"].as_array().unwrap().len()
    );
    let result = read_workbook(&result_workbook).expect("read S1 result workbook with scan issue");
    let scan_issue_id = scanner_issue["issue_id"]
        .as_str()
        .expect("scanner issue identity");
    let scan_issue_row = (2..=report["result_workbook"]["issue_count"].as_u64().unwrap() + 1)
        .find(|row| result.sheets[0].values[&format!("K{row}")]["value"] == scan_issue_id)
        .expect("validation scan issue in generated result workbook");
    assert_eq!(
        result.sheets[0].values[&format!("E{scan_issue_row}")]["value"],
        "校验公式错误"
    );
    assert_eq!(
        result.sheets[0].values[&format!("H{scan_issue_row}")]["value"],
        "#DIV/0!"
    );
    assert_eq!(
        report["history_result_workbook"]["issue_count"],
        report["merged_history"].as_array().unwrap().len()
    );
    let history_result = read_workbook(&history_result_workbook)
        .expect("read S1 history workbook with validation scan issue");
    let history_issue_row = (2..=report["history_result_workbook"]["issue_count"]
        .as_u64()
        .unwrap()
        + 1)
        .find(|row| history_result.sheets[0].values[&format!("K{row}")]["value"] == scan_issue_id)
        .expect("validation scan issue in generated S1 history workbook");
    assert_eq!(
        history_result.sheets[0].values[&format!("E{history_issue_row}")]["value"],
        "校验公式错误"
    );
    assert!(
        !report["validation_formula_scan"]["per_file"][0]
            .to_string()
            .contains("#REF!")
    );
    let events = String::from_utf8_lossy(&output.stderr)
        .lines()
        .map(|line| serde_json::from_str::<Value>(line).expect("parse developer event"))
        .collect::<Vec<_>>();
    assert!(events.iter().any(|event| {
        event["step"] == "validation_formula_error_scan"
            && event["engine"] == "calamine+namespace_aware_ooxml_defined_names"
            && event["formula_errors"] == 1
            && event["issues_added"] == 1
    }));

    if std::env::var_os("AUDIT_S1_KEEP_VALIDATION_SCAN_ARTIFACTS").is_some() {
        fs::write(directory.join("rust-result.json"), &output.stdout)
            .expect("save synthetic Rust scan result for Python parity");
        println!("S1_VALIDATION_SCAN_PARITY_DIR={}", directory.display());
    } else {
        fs::remove_dir_all(directory).expect("remove temporary validation scan fixtures");
    }
}

#[test]
fn cli_stops_before_formula_copy_manifest_when_template_health_fails() {
    let directory = temporary_directory("template_health_failure");
    let input = directory.join("input");
    fs::create_dir_all(&input).expect("create source input directory");
    let source = input.join("1001_机构甲_2026-07.xlsx");
    audit_copy(&source);
    let template = directory.join("template.xlsx");
    formula_template_without_result_rule_or_structure_area(&template);
    let manifest = directory.join("invalid-manifest.json");
    fs::write(&manifest, "{ not valid JSON").expect("write invalid manifest sentinel");

    let output = run_cli(&input, &template, &manifest, true);
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(!output.status.success(), "{stderr}");
    assert!(stderr.contains("S1 模板体检不通过"), "{stderr}");
    let events = stderr
        .lines()
        .filter_map(|line| serde_json::from_str::<Value>(line).ok())
        .collect::<Vec<_>>();
    let health_event = events
        .iter()
        .find(|event| event["step"] == "template_health_preflight")
        .expect("template-health progress event");
    assert_eq!(health_event["status"], "failed");
    assert_eq!(health_event["enabled_rules"], 1);
    assert_eq!(health_event["result_rules"], 0);
    assert_eq!(health_event["structure_ranges"], 0);
    assert_eq!(health_event["errors"], 3);
    assert!(
        !stderr.contains("解析公式审核副本清单"),
        "manifest JSON must not be parsed on template-health failure"
    );

    fs::remove_dir_all(&directory).expect("remove temporary directory");
}

#[test]
fn cli_rejects_incomplete_formula_copy_coverage_before_history_classification() {
    let directory = temporary_directory("missing_formula");
    let input = directory.join("input");
    fs::create_dir_all(&input).expect("create source input directory");
    fs::copy(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("tests")
            .join("fixtures")
            .join("s1_issue_chain_source.xlsx"),
        input.join("1001_机构甲_2026-07.xlsx"),
    )
    .expect("copy conditional-format source fixture");
    let template = directory.join("template.xlsx");
    formula_template(&template);
    let manifest = directory.join("formula-copies.json");
    fs::write(&manifest, r#"{"files":[]}"#).expect("write incomplete manifest");

    let output = run_cli(&input, &template, &manifest, false);
    assert!(!output.status.success());
    let error = String::from_utf8_lossy(&output.stderr);
    assert!(error.contains("公式审核副本清单中没有工作簿映射"));

    fs::remove_dir_all(directory).expect("remove temporary incomplete manifest fixture");
}

#[test]
fn cli_reads_nonempty_history_before_classification_and_reports_reader_and_timing() {
    let directory = temporary_directory("history");
    let input = directory.join("input");
    fs::create_dir_all(&input).expect("create source input directory");
    let source = input.join("1001_机构甲_2026-07.xlsx");
    fs::copy(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("tests")
            .join("fixtures")
            .join("s1_issue_chain_source.xlsx"),
        &source,
    )
    .expect("copy conditional-format source fixture");
    let template = directory.join("template.xlsx");
    formula_template(&template);
    let audit = directory.join("audit-copy.xlsx");
    audit_copy(&audit);
    let manifest = directory.join("formula-copies.json");
    fs::write(
        &manifest,
        serde_json::to_vec(&json!({
            "files": [{
                "source_file": source.file_name().unwrap().to_string_lossy(),
                "audit_workbook": audit.file_name().unwrap().to_string_lossy()
            }]
        }))
        .unwrap(),
    )
    .expect("write formula copy manifest");

    let without_history = run_cli(&input, &template, &manifest, false);
    assert!(
        without_history.status.success(),
        "{}",
        String::from_utf8_lossy(&without_history.stderr)
    );
    let baseline: Value =
        serde_json::from_slice(&without_history.stdout).expect("parse baseline issue-chain JSON");
    let continuing_id = baseline["issues"][0]["issue_id"]
        .as_str()
        .expect("formula issue identity");
    let history = directory.join("history.xlsx");
    history_workbook(&history, continuing_id);

    let ordinary = run_cli_with_history(&input, &template, &manifest, Some(&history), false);
    assert!(
        ordinary.status.success(),
        "{}",
        String::from_utf8_lossy(&ordinary.stderr)
    );
    let report: Value = serde_json::from_slice(&ordinary.stdout).expect("parse history report");
    assert_eq!(report["template_selection"]["mode"], "explicit");
    assert_eq!(
        report["template_selection"]["template_file"],
        "template.xlsx"
    );
    assert_eq!(report["history_workbook_summary"]["status"], "loaded");
    assert_eq!(report["history_workbook_summary"]["reader"], "calamine");
    assert_eq!(report["history_workbook_summary"]["rows_read"], 2);
    assert_eq!(report["summary"]["history_rows"], 2);
    assert_eq!(report["summary"]["continuous_issues"], 1);
    assert_eq!(report["summary"]["resolved_issues"], 1);
    let continuing = report["issues"]
        .as_array()
        .unwrap()
        .iter()
        .find(|issue| issue["issue_id"] == continuing_id)
        .expect("current issue matching history");
    assert_eq!(continuing["status"], "连续出现");
    assert_eq!(continuing["institution_feedback"], "历史机构反馈");
    assert_eq!(continuing["auditor_opinion"], "历史审核意见");
    assert_eq!(
        report["resolved_issues"][0]["issue_id"],
        "history-only-issue"
    );
    assert_eq!(report["resolved_issues"][0]["status"], "已整改");

    let user_log = String::from_utf8_lossy(&ordinary.stderr);
    assert!(user_log.contains("S1 历史工作簿读取完成"));
    assert!(!user_log.contains("calamine"), "{user_log}");
    assert!(!user_log.contains("读取器"), "{user_log}");
    assert!(user_log.contains("读取 2/2 行"));

    let developer = run_cli_with_history(&input, &template, &manifest, Some(&history), true);
    assert!(
        developer.status.success(),
        "{}",
        String::from_utf8_lossy(&developer.stderr)
    );
    let developer_report: Value =
        serde_json::from_slice(&developer.stdout).expect("parse developer history report");
    assert_eq!(developer_report["issues"], report["issues"]);
    let events = String::from_utf8_lossy(&developer.stderr)
        .lines()
        .map(|line| serde_json::from_str::<Value>(line).expect("parse developer progress event"))
        .collect::<Vec<_>>();
    assert!(events.iter().any(|event| {
        event["step"] == "history_workbook_read"
            && event["status"] == "loaded"
            && event["engine"] == "calamine"
            && event["rows_read"] == 2
            && event["elapsed_ms"].is_number()
    }));

    fs::remove_dir_all(directory).expect("remove synthetic S1 history workbooks");
}

#[test]
fn production_s1_directory_command_requires_fresh_formula_calculation() {
    let root = temporary_directory("recalculation_required");
    let output = Command::new(env!("CARGO_BIN_EXE_open-data-audit"))
        .arg("s1-audit-directory")
        .arg("--input-dir")
        .arg(&root)
        .arg("--formula-manifest")
        .arg(root.join("missing-manifest.json"))
        .arg("--template-workbook")
        .arg(root.join("missing-template.xlsx"))
        .arg("--batch-id")
        .arg("cache-policy-test")
        .output()
        .expect("run S1 production entry without calculation sidecar");
    assert!(!output.status.success());
    assert!(String::from_utf8_lossy(&output.stderr).contains("强制 Office/LibreOffice 重算"));
}

#[test]
#[ignore = "requires AUDIT_S1_RECALC_TEST_SIDECAR and Windows Microsoft Excel COM"]
fn s1_directory_with_excel_sidecar_recalculates_a_copy_before_auditing() -> anyhow::Result<()> {
    let sidecar = std::env::var_os("AUDIT_S1_RECALC_TEST_SIDECAR")
        .map(PathBuf::from)
        .expect("AUDIT_S1_RECALC_TEST_SIDECAR is required for the Excel COM integration test");
    let directory = temporary_directory("excel_sidecar_end_to_end");
    let input = directory.join("input");
    fs::create_dir_all(&input).expect("create S1 input directory");
    let source = input.join("1001_机构甲_2026-07.xlsx");
    fs::copy(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("tests")
            .join("fixtures")
            .join("s1_issue_chain_source.xlsx"),
        &source,
    )
    .expect("copy synthetic source with a triggered conditional format");
    let template = directory.join("template.xlsx");
    formula_template(&template);
    let audit_copy = directory.join("audit-copy.xlsx");
    fs::copy(&template, &audit_copy).expect("create separate formula audit copy");
    let manifest = directory.join("formula-copies.json");
    fs::write(
        &manifest,
        serde_json::to_vec(&json!({
            "files": [{
                "source_file": source.file_name().unwrap().to_string_lossy(),
                "audit_workbook": audit_copy.file_name().unwrap().to_string_lossy()
            }]
        }))?,
    )
    .expect("write S1 formula-copy manifest");
    let source_before = fs::read(&source)?;
    let audit_copy_before = fs::read(&audit_copy)?;
    let recalculated_dir = directory.join("recalculated");
    let result_workbook = directory.join("s1-result.xlsx");
    let history_workbook = directory.join("s1-history.xlsx");

    let output = Command::new(env!("CARGO_BIN_EXE_open-data-audit"))
        .arg("s1-audit-directory")
        .arg("--input-dir")
        .arg(&input)
        .arg("--formula-manifest")
        .arg(&manifest)
        .arg("--period")
        .arg("2026-07")
        .arg("--batch-id")
        .arg("excel-sidecar-e2e")
        .arg("--audit-time")
        .arg("2026-10-27 12:00:00")
        .arg("--template-workbook")
        .arg(&template)
        .arg("--recalculate-sidecar")
        .arg(&sidecar)
        .arg("--recalculated-copy-dir")
        .arg(&recalculated_dir)
        .arg("--result-workbook")
        .arg(&result_workbook)
        .arg("--history-result-workbook")
        .arg(&history_workbook)
        .arg("--developer-view")
        .output()
        .expect("run S1 full flow with Excel calculation sidecar");
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report: Value = serde_json::from_slice(&output.stdout)?;
    assert_eq!(report["status"], "completed");
    assert_eq!(
        report["formula_execution"]["status"],
        "recalculated_by_sidecar"
    );
    assert_eq!(report["formula_execution"]["files_recalculated"], 1);
    assert!(
        report["formula_summary"]["issues"]
            .as_u64()
            .unwrap_or_default()
            > 0
    );
    assert_eq!(report["conditional_summary"]["issues"], 1);
    assert!(report["result_workbook"].is_object());
    assert!(report["history_result_workbook"].is_object());
    assert!(result_workbook.is_file());
    assert!(history_workbook.is_file());
    assert_eq!(fs::read(&source)?, source_before);
    assert_eq!(fs::read(&audit_copy)?, audit_copy_before);

    let progress = String::from_utf8_lossy(&output.stderr)
        .lines()
        .filter_map(|line| serde_json::from_str::<Value>(line).ok())
        .collect::<Vec<_>>();
    let recalculated = progress
        .iter()
        .find(|event| {
            event["step"] == "formula_copy_recalculation" && event["status"] == "completed"
        })
        .expect("completed formula-copy recalculation progress event");
    assert_eq!(recalculated["engine"], "Excel.Application");
    let recalculated_path = PathBuf::from(
        recalculated["output_audit_copy"]
            .as_str()
            .expect("recalculated audit-copy path"),
    );
    let recalculated_workbook = excel_reader::read_workbook(&recalculated_path)?;
    assert_eq!(
        recalculated_workbook.sheets[0].values["A1"]["value"],
        "错误|复核|存款|低于要求|100|200|100"
    );
    let result_workbook_read = read_workbook(&result_workbook)?;
    let result_sheet = &result_workbook_read.sheets[0];
    let issues = report["issues"].as_array().expect("combined S1 issues");
    let conditional_issue_index = issues
        .iter()
        .position(|issue| issue["formula_cell"] == "B2")
        .expect("triggered conditional-format issue");
    assert_eq!(
        result_sheet.values[&format!("K{}", conditional_issue_index + 2)]["value"],
        issues[conditional_issue_index]["issue_id"]
    );
    assert_eq!(
        report["result_workbook"]["issue_count"],
        report["summary"]["new_issues"]
    );

    fs::remove_dir_all(directory).expect("remove S1 Excel sidecar E2E fixtures");
    Ok(())
}
