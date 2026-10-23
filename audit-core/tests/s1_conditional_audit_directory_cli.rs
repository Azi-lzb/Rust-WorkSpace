use excel_writer::Workbook;
use serde_json::Value;
use std::{
    fs,
    path::{Path, PathBuf},
    process::{Command, Output},
    time::{SystemTime, UNIX_EPOCH},
};

fn fixture() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("s1_conditional_eval.xlsx")
}

fn temporary_directory(label: &str) -> PathBuf {
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system clock after Unix epoch")
        .as_nanos();
    let path = std::env::temp_dir().join(format!(
        "open_data_audit_{label}_{}_{}",
        std::process::id(),
        nonce
    ));
    fs::create_dir_all(&path).expect("create temporary input directory");
    path
}

fn copy_fixture(directory: &Path, name: &str) {
    fs::copy(fixture(), directory.join(name)).expect("copy synthetic XLSX fixture");
}

fn run_cli(
    directory: &Path,
    period: Option<&str>,
    recursive: bool,
    developer_view: bool,
) -> Output {
    run_cli_with_mapping(
        directory,
        period,
        recursive,
        developer_view,
        None,
        None,
        None,
        None,
        None,
    )
}

fn run_cli_with_mapping(
    directory: &Path,
    period: Option<&str>,
    recursive: bool,
    developer_view: bool,
    template_workbook: Option<&Path>,
    template_dir: Option<&Path>,
    mapping_config: Option<&Path>,
    scan_area: Option<&str>,
    selected_file: Option<&Path>,
) -> Output {
    let mut command = Command::new(env!("CARGO_BIN_EXE_open-data-audit"));
    command
        .arg("conditional-audit-directory")
        .arg("--input-dir")
        .arg(directory)
        .arg("--batch-id")
        .arg("directory-batch")
        .arg("--audit-time")
        .arg("2026-10-24 12:00:00");
    if let Some(period) = period {
        command.arg("--period").arg(period);
    }
    if recursive {
        command.arg("--recursive");
    }
    if let Some(template_workbook) = template_workbook {
        command.arg("--template-workbook").arg(template_workbook);
    }
    if let Some(template_dir) = template_dir {
        command.arg("--template-dir").arg(template_dir);
    }
    if let Some(mapping_config) = mapping_config {
        command.arg("--mapping-config").arg(mapping_config);
    }
    if let Some(scan_area) = scan_area {
        command.arg("--scan-area").arg(scan_area);
    }
    if let Some(selected_file) = selected_file {
        command.arg("--selected-file").arg(selected_file);
    }
    if developer_view {
        command.arg("--developer-view");
    }
    command.output().expect("run directory S1 conditional CLI")
}

#[test]
fn directory_cli_discovers_period_derives_organization_and_logs_both_views() {
    let directory = temporary_directory("s1_directory");
    copy_fixture(&directory, "9144_甲银行_2026-07-31.xlsx");

    let user = run_cli(&directory, None, false, false);
    assert!(
        user.status.success(),
        "{}",
        String::from_utf8_lossy(&user.stderr)
    );
    let user_report: Value = serde_json::from_slice(&user.stdout).expect("valid batch JSON");
    assert_eq!(user_report["files"].as_array().unwrap().len(), 1);
    assert_eq!(user_report["issues"].as_array().unwrap().len(), 4);
    assert_eq!(user_report["issues"][0]["period"], "2026-07");
    assert_eq!(user_report["issues"][0]["org_code"], "9144");
    assert_eq!(user_report["issues"][0]["org_name"], "甲银行");
    let user_log = String::from_utf8_lossy(&user.stderr);
    assert!(user_log.contains("源目录发现 1 个 XLSX"));
    assert!(user_log.contains("数据期 2026-07（源文件名）"));
    assert!(user_log.contains("引擎 calamine+namespace_aware_ooxml"));
    assert!(user_log.contains("读取"));

    let developer = run_cli(&directory, None, false, true);
    assert!(
        developer.status.success(),
        "{}",
        String::from_utf8_lossy(&developer.stderr)
    );
    let events = String::from_utf8_lossy(&developer.stderr)
        .lines()
        .map(|line| serde_json::from_str::<Value>(line).expect("JSON progress line"))
        .collect::<Vec<_>>();
    let discovery = events
        .iter()
        .find(|event| {
            event["step"] == "source_workbook_discovery" && event["status"] == "completed"
        })
        .expect("discovery completion event");
    assert_eq!(discovery["files_total"], 1);
    assert_eq!(discovery["period"], "2026-07");
    assert_eq!(discovery["period_source"], "源文件名");
    assert!(discovery["elapsed_ms"].as_u64().is_some());
    assert!(events.iter().any(|event| {
        event["step"] == "conditional_format_issue_extraction" && event["status"] == "completed"
    }));

    fs::remove_dir_all(directory).expect("remove temporary input directory");
}

fn named_template_fixture() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("s1_conditional_named_template.xlsx")
}

fn mapping_config(path: &Path) {
    let mut workbook = Workbook::new();
    {
        let worksheet = workbook.add_worksheet();
        worksheet
            .set_name("模块化功能")
            .expect("set mapping configuration sheet name");
        let rows = [
            vec!["功能名", "执行模块", "命名区域名", "输入", "输出", "备注"],
            vec![
                "条件格式结果提取",
                "汇总_条件格式结果提取",
                "条件格式区域_集成样例",
                "",
                "",
                "",
            ],
            vec![
                "表结构比对",
                "核对_表结构比对",
                "表结构区域_集成样例",
                "",
                "",
                "",
            ],
        ];
        for (row_index, row) in rows.iter().enumerate() {
            for (column_index, value) in row.iter().enumerate() {
                if !value.is_empty() {
                    worksheet
                        .write_string(row_index as u32, column_index as u16, *value)
                        .expect("write mapping configuration value");
                }
            }
        }
    }
    workbook
        .save(path)
        .expect("save synthetic mapping configuration");
}

fn workbook_with_sheet_name(path: &Path, name: &str) {
    let mut workbook = Workbook::new();
    workbook
        .add_worksheet()
        .set_name(name)
        .expect("set synthetic workbook sheet name");
    workbook.save(path).expect("save synthetic workbook");
}

#[test]
fn directory_cli_reads_template_defined_names_and_applies_configured_areas() {
    let directory = temporary_directory("s1_named_area_mapping");
    copy_fixture(&directory, "9144_甲银行_2026-07-31.xlsx");
    let template = directory.join("template_2026-09.xlsx");
    fs::copy(named_template_fixture(), &template).expect("copy synthetic named template");
    let config = directory.join("mapping-config_2026-08.xlsx");
    mapping_config(&config);

    let output = run_cli_with_mapping(
        &directory,
        None,
        false,
        true,
        Some(&template),
        None,
        Some(&config),
        None,
        None,
    );
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report: Value =
        serde_json::from_slice(&output.stdout).expect("valid named-range batch JSON");
    assert_eq!(report["files"].as_array().unwrap().len(), 1);
    assert_eq!(report["files"][0]["scanned_cells"], 1);
    assert_eq!(report["files"][0]["issue_count"], 1);
    assert_eq!(report["issues"].as_array().unwrap().len(), 1);
    assert_eq!(report["issues"][0]["target_cell"], "B2");

    let events = String::from_utf8_lossy(&output.stderr)
        .lines()
        .map(|line| serde_json::from_str::<Value>(line).expect("JSON progress line"))
        .collect::<Vec<_>>();
    let mapping = events
        .iter()
        .find(|event| event["step"] == "conditional_area_mapping" && event["status"] == "completed")
        .expect("template mapping completion event");
    let discovery = events
        .iter()
        .find(|event| {
            event["step"] == "source_workbook_discovery" && event["status"] == "completed"
        })
        .expect("source discovery completion event");
    assert_eq!(discovery["period"], "2026-07");
    assert_eq!(discovery["period_source"], "源文件名");
    assert_eq!(mapping["engine"], "calamine+namespace_aware_ooxml");
    assert_eq!(mapping["area_source"], "template_defined_names");
    assert_eq!(mapping["mapped_named_areas"], 2);
    assert_eq!(mapping["scan_areas"], 1);
    assert_eq!(mapping["structure_areas"], 1);
    assert!(mapping["elapsed_ms"].as_u64().is_some());

    let override_output = run_cli_with_mapping(
        &directory,
        Some("2026-07"),
        false,
        false,
        Some(&template),
        None,
        Some(&config),
        Some("S1Conditional!C5:C5"),
        None,
    );
    assert!(
        override_output.status.success(),
        "{}",
        String::from_utf8_lossy(&override_output.stderr)
    );
    let override_report: Value =
        serde_json::from_slice(&override_output.stdout).expect("valid override batch JSON");
    assert_eq!(override_report["issues"].as_array().unwrap().len(), 1);
    assert_eq!(override_report["issues"][0]["target_cell"], "C5");

    fs::remove_dir_all(directory).expect("remove temporary input directory");
}

#[test]
fn directory_cli_recommends_template_for_selected_sources_and_excludes_template_files() {
    let directory = temporary_directory("s1_directory_template_recommendation");
    let input_dir = directory.join("input");
    let template_dir = input_dir.join("templates");
    fs::create_dir_all(&input_dir).expect("create source input directory");
    fs::create_dir_all(&template_dir).expect("create template directory inside source tree");
    copy_fixture(&input_dir, "9144_甲银行_2026-07-31.xlsx");
    let template = template_dir.join("单位模板.xlsx");
    fs::copy(named_template_fixture(), &template).expect("copy synthetic named template");
    let config = directory.join("mapping-config.xlsx");
    mapping_config(&config);

    let user = run_cli_with_mapping(
        &input_dir,
        None,
        true,
        false,
        None,
        Some(&template_dir),
        Some(&config),
        None,
        None,
    );
    assert!(
        user.status.success(),
        "{}",
        String::from_utf8_lossy(&user.stderr)
    );
    let user_report: Value = serde_json::from_slice(&user.stdout).expect("valid audit JSON");
    assert_eq!(user_report["files"].as_array().unwrap().len(), 1);
    assert_eq!(user_report["issues"].as_array().unwrap().len(), 1);
    assert_eq!(user_report["issues"][0]["target_cell"], "B2");
    assert!(String::from_utf8_lossy(&user.stderr).contains("模板推荐完成"));
    assert!(String::from_utf8_lossy(&user.stderr).contains("namespace-aware OOXML"));

    let developer = run_cli_with_mapping(
        &input_dir,
        None,
        true,
        true,
        None,
        Some(&template_dir),
        Some(&config),
        None,
        None,
    );
    assert!(
        developer.status.success(),
        "{}",
        String::from_utf8_lossy(&developer.stderr)
    );
    let events = String::from_utf8_lossy(&developer.stderr)
        .lines()
        .map(|line| serde_json::from_str::<Value>(line).expect("JSON progress line"))
        .collect::<Vec<_>>();
    let discovery = events
        .iter()
        .find(|event| {
            event["step"] == "source_workbook_discovery" && event["status"] == "completed"
        })
        .expect("source discovery completion event");
    assert_eq!(discovery["available_source_files"], 1);
    let recommendation = events
        .iter()
        .find(|event| event["step"] == "template_recommendation" && event["status"] == "completed")
        .expect("template recommendation completion event");
    assert_eq!(
        recommendation["engine"],
        "namespace_aware_ooxml_sheet_directory"
    );
    assert_eq!(recommendation["source_workbooks"], 1);
    assert_eq!(recommendation["selected_template"], "单位模板.xlsx");
    assert_eq!(recommendation["confidence"], 1.0);
    assert!(recommendation["elapsed_ms"].as_u64().is_some());

    fs::remove_dir_all(directory).expect("remove temporary input directory");
}

#[test]
fn directory_cli_stops_when_template_recommendation_is_ambiguous() {
    let directory = temporary_directory("s1_directory_template_ambiguity");
    let input_dir = directory.join("input");
    let template_dir = directory.join("templates");
    fs::create_dir_all(&input_dir).expect("create source input directory");
    fs::create_dir_all(&template_dir).expect("create template directory");
    copy_fixture(&input_dir, "9144_甲银行_2026-07-31.xlsx");
    fs::copy(named_template_fixture(), template_dir.join("候选甲.xlsx"))
        .expect("copy first synthetic template");
    fs::copy(named_template_fixture(), template_dir.join("候选乙.xlsx"))
        .expect("copy second synthetic template");

    let output = run_cli_with_mapping(
        &input_dir,
        None,
        false,
        true,
        None,
        Some(&template_dir),
        None,
        None,
        None,
    );
    assert!(!output.status.success());
    assert!(output.stdout.is_empty());
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.contains("模板自动推荐未能唯一确定模板"), "{stderr}");
    assert!(
        stderr.contains("请显式指定 --template-workbook"),
        "{stderr}"
    );
    assert!(stderr.contains("候选甲.xlsx"), "{stderr}");
    assert!(stderr.contains("候选乙.xlsx"), "{stderr}");
    let failed_event = stderr
        .lines()
        .filter(|line| line.trim_start().starts_with('{'))
        .map(|line| serde_json::from_str::<Value>(line).expect("JSON progress line"))
        .find(|event| event["step"] == "template_recommendation" && event["status"] == "failed")
        .expect("template ambiguity failure event");
    assert_eq!(
        failed_event["engine"],
        "namespace_aware_ooxml_sheet_directory"
    );
    assert_eq!(failed_event["matched"], false);
    assert_eq!(failed_event["alternatives"].as_array().unwrap().len(), 2);

    fs::remove_dir_all(directory).expect("remove temporary input directory");
}

#[test]
fn directory_cli_template_recommendation_respects_selected_files_for_mixed_types() {
    let directory = temporary_directory("s1_directory_template_mixed_types");
    let input_dir = directory.join("input");
    let template_dir = directory.join("templates");
    fs::create_dir_all(&input_dir).expect("create source input directory");
    fs::create_dir_all(&template_dir).expect("create template directory");
    copy_fixture(&input_dir, "9144_甲银行_2026-07-31.xlsx");
    workbook_with_sheet_name(&input_dir.join("9145_乙银行_2026-07-31.xlsx"), "其他报表");
    fs::copy(named_template_fixture(), template_dir.join("单位模板.xlsx"))
        .expect("copy unit template");
    workbook_with_sheet_name(&template_dir.join("其他模板.xlsx"), "其他报表");
    let config = directory.join("mapping-config.xlsx");
    mapping_config(&config);

    let mixed = run_cli_with_mapping(
        &input_dir,
        None,
        false,
        true,
        None,
        Some(&template_dir),
        Some(&config),
        None,
        None,
    );
    assert!(!mixed.status.success());
    let mixed_stderr = String::from_utf8_lossy(&mixed.stderr);
    assert!(
        mixed_stderr.contains("源数据目录可能包含多种报表"),
        "{mixed_stderr}"
    );

    let selected = run_cli_with_mapping(
        &input_dir,
        None,
        false,
        true,
        None,
        Some(&template_dir),
        Some(&config),
        None,
        Some(Path::new("9144_甲银行_2026-07-31.xlsx")),
    );
    assert!(
        selected.status.success(),
        "{}",
        String::from_utf8_lossy(&selected.stderr)
    );
    let report: Value =
        serde_json::from_slice(&selected.stdout).expect("valid selected audit JSON");
    assert_eq!(report["files"].as_array().unwrap().len(), 1);
    assert_eq!(report["issues"].as_array().unwrap().len(), 1);
    let events = String::from_utf8_lossy(&selected.stderr)
        .lines()
        .map(|line| serde_json::from_str::<Value>(line).expect("JSON progress line"))
        .collect::<Vec<_>>();
    let recommendation = events
        .iter()
        .find(|event| event["step"] == "template_recommendation" && event["status"] == "completed")
        .expect("selected-source recommendation event");
    assert_eq!(recommendation["source_workbooks"], 1);
    assert_eq!(recommendation["selected_template"], "单位模板.xlsx");

    fs::remove_dir_all(directory).expect("remove temporary input directory");
}

#[test]
fn directory_cli_limits_period_detection_to_selected_files() {
    let directory = temporary_directory("s1_selected_source");
    copy_fixture(&directory, "9144_甲银行_2026-06-30.xlsx");
    copy_fixture(&directory, "9145_乙银行_2026-07-31.xlsx");
    let template = directory.join("template_2026-09.xlsx");
    fs::copy(named_template_fixture(), &template).expect("copy synthetic named template");
    let config = directory.join("mapping-config_2026-08.xlsx");
    mapping_config(&config);

    let selected = run_cli_with_mapping(
        &directory,
        None,
        false,
        true,
        Some(&template),
        None,
        Some(&config),
        None,
        Some(Path::new("9145_乙银行_2026-07-31.xlsx")),
    );
    assert!(
        selected.status.success(),
        "{}",
        String::from_utf8_lossy(&selected.stderr)
    );
    let report: Value =
        serde_json::from_slice(&selected.stdout).expect("valid selected batch JSON");
    assert_eq!(report["files"].as_array().unwrap().len(), 1);
    assert_eq!(report["issues"].as_array().unwrap().len(), 1);
    assert_eq!(report["issues"][0]["period"], "2026-07");

    let events = String::from_utf8_lossy(&selected.stderr)
        .lines()
        .map(|line| serde_json::from_str::<Value>(line).expect("JSON progress line"))
        .collect::<Vec<_>>();
    let discovery = events
        .iter()
        .find(|event| {
            event["step"] == "source_workbook_discovery" && event["status"] == "completed"
        })
        .expect("source discovery completion event");
    assert_eq!(discovery["available_source_files"], 2);
    assert_eq!(discovery["files_total"], 1);
    assert_eq!(discovery["selected_files_requested"], 1);
    assert_eq!(discovery["selection_applied"], true);
    assert_eq!(discovery["period"], "2026-07");

    let rejected = run_cli_with_mapping(
        &directory,
        None,
        false,
        false,
        Some(&template),
        None,
        Some(&config),
        None,
        Some(Path::new("../outside.xlsx")),
    );
    assert!(!rejected.status.success());
    assert!(
        String::from_utf8_lossy(&rejected.stderr)
            .contains("选择的待审核文件不在源数据目录或当前扫描范围内"),
        "{}",
        String::from_utf8_lossy(&rejected.stderr)
    );

    fs::remove_dir_all(directory).expect("remove temporary input directory");
}

#[test]
fn directory_cli_requires_explicit_period_when_filenames_conflict() {
    let directory = temporary_directory("s1_period_conflict");
    copy_fixture(&directory, "9144_甲银行_2026-06-30.xlsx");
    copy_fixture(&directory, "9145_乙银行_2026-07-31.xlsx");

    let output = run_cli(&directory, None, false, false);
    assert!(!output.status.success());
    assert!(output.stdout.is_empty());
    let error = String::from_utf8_lossy(&output.stderr);
    assert!(error.contains("多个数据期"), "{error}");
    assert!(error.contains("请显式指定 --period"), "{error}");

    fs::remove_dir_all(directory).expect("remove temporary input directory");
}

#[test]
fn directory_cli_recurses_only_when_requested() {
    let directory = temporary_directory("s1_directory_recursive");
    let nested = directory.join("nested");
    fs::create_dir_all(&nested).expect("create nested input directory");
    copy_fixture(&nested, "9144_甲银行_2026-07-31.xlsx");

    let root_only = run_cli(&directory, Some("2026-07"), false, false);
    assert!(!root_only.status.success());
    assert!(String::from_utf8_lossy(&root_only.stderr).contains("没有可审核的 .xlsx 文件"));

    let recursive = run_cli(&directory, Some("2026-07"), true, false);
    assert!(
        recursive.status.success(),
        "{}",
        String::from_utf8_lossy(&recursive.stderr)
    );
    let report: Value =
        serde_json::from_slice(&recursive.stdout).expect("valid recursive batch JSON");
    assert_eq!(report["files"].as_array().unwrap().len(), 1);
    assert_eq!(report["issues"].as_array().unwrap().len(), 4);

    fs::remove_dir_all(directory).expect("remove temporary input directory");
}
