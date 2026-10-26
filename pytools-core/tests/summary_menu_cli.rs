use anyhow::{Context, Result};
use excel_reader::read_workbook;
use serde_json::Value;
use std::{
    fs,
    io::Write,
    path::{Path, PathBuf},
    process::{Command, Output, Stdio},
    sync::atomic::{AtomicU64, Ordering},
    time::{SystemTime, UNIX_EPOCH},
};

static NEXT_TEST_DIRECTORY_ID: AtomicU64 = AtomicU64::new(0);

struct TestDirectory(PathBuf);

impl TestDirectory {
    fn new() -> Result<Self> {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .context("system clock before Unix epoch")?
            .as_nanos();
        let serial = NEXT_TEST_DIRECTORY_ID.fetch_add(1, Ordering::Relaxed);
        let path = std::env::temp_dir().join(format!(
            "pytools-summary-menu-{}-{nonce}-{serial}",
            std::process::id()
        ));
        fs::create_dir_all(&path)?;
        Ok(Self(path))
    }

    fn path(&self) -> &Path {
        &self.0
    }
}

impl Drop for TestDirectory {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

#[test]
fn interactive_menu_routes_used_range_summary_through_global_config() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_app_config(root.path())?;
    let source_one = root.path().join("source branch one.xlsx");
    let source_two = root.path().join("source branch two.xlsx");
    create_used_range_source(&source_one, "north", 11.5)?;
    create_used_range_source(&source_two, "south", 22.5)?;
    let source_one_before = fs::read(&source_one)?;
    let source_two_before = fs::read(&source_two)?;

    let input = format!(
        r#"1
2
{};{}
0
"#,
        source_one.display(),
        source_two.display()
    );
    let output = run_menu(&config, &input)?;
    assert!(
        output.status.success(),
        "interactive PT-6-1 failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("1  数据提取与汇总"));
    assert!(stdout.contains("2  按使用区域快速汇总"));
    assert!(stdout.contains("可手动输入路径"));
    assert!(stdout.contains(r#""saved": true"#));
    assert!(stdout.contains("已选择文件 2 个"));

    let logs = root.path().join("app").join("logs");
    let (history_path, history, history_text) = read_single_history(&logs, "PT-6-1_")?;
    assert_eq!(history["product"], "pytools");
    assert_eq!(history["feature_id"], "PT-6-1");
    assert_eq!(history["function"], "summary-used-range");
    assert_eq!(history["status"], "succeeded");
    assert_eq!(history["total_inputs"], 2);
    assert_eq!(history["summary"]["rows"], 4);
    assert_eq!(history["summary"]["files"], 2);
    assert_eq!(history["summary"]["execution"]["reader_engine"], "calamine");
    assert_eq!(
        history["summary"]["execution"]["writer_engine"],
        "rust_xlsxwriter"
    );
    assert!(!history_text.contains("north"));
    assert!(!history_text.contains("south"));

    let developer_events = String::from_utf8_lossy(&output.stderr)
        .lines()
        .map(serde_json::from_str::<Value>)
        .collect::<std::result::Result<Vec<_>, _>>()
        .context("parse interactive developer events")?;
    assert!(!developer_events.is_empty());
    assert_eq!(history["events"], serde_json::json!(developer_events));
    for phase in ["extent_scan", "value_read", "row_write", "output_save"] {
        assert!(
            developer_events.iter().any(|event| {
                event["phase"] == phase
                    && event["status"] == "succeeded"
                    && event["stage_elapsed_ms"].is_number()
            }),
            "successful {phase} timing is absent from the interactive run"
        );
    }
    assert!(developer_events
        .iter()
        .any(|event| { event["phase"] == "value_read" && event["actual_engine"] == "calamine" }));
    assert!(developer_events.iter().any(|event| {
        event["phase"] == "output_save" && event["actual_engine"] == "rust_xlsxwriter"
    }));

    let summary_path = PathBuf::from(
        history["summary"]["path"]
            .as_str()
            .context("summary path is missing from history")?,
    );
    assert!(summary_path.starts_with(root.path().join("app").join("output")));
    let summary = read_workbook(&summary_path)?;
    let sheet = summary
        .sheets
        .iter()
        .find(|sheet| sheet.name == "汇总")
        .context("interactive summary sheet missing")?;
    assert_eq!(
        sheet.values.get("A3"),
        Some(&serde_json::json!({"type":"string", "value":"source branch one"}))
    );
    assert_eq!(
        sheet.values.get("D3"),
        Some(&serde_json::json!({"type":"string", "value":"north"}))
    );
    assert_eq!(
        sheet.values.get("A5"),
        Some(&serde_json::json!({"type":"string", "value":"source branch two"}))
    );
    assert_eq!(
        sheet.values.get("D5"),
        Some(&serde_json::json!({"type":"string", "value":"south"}))
    );
    assert_eq!(fs::read(&source_one)?, source_one_before);
    assert_eq!(fs::read(&source_two)?, source_two_before);
    assert!(history_path.starts_with(&logs));
    Ok(())
}

#[test]
fn interactive_v2_named_range_summary_uses_named_regions_and_skips_bad_structure() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_app_config(root.path())?;
    let template = root.path().join("named template.xlsx");
    let source = root.path().join("source branch.xlsx");
    let shorter_source = root.path().join("source branch short.xlsx");
    let mismatched = root.path().join("mismatched branch.xlsx");
    create_named_range_template(&template)?;
    create_named_range_source(&source, "结构A", true)?;
    create_named_range_source(&shorter_source, "结构A", false)?;
    create_named_range_source(&mismatched, "结构B", true)?;
    let template_before = fs::read(&template)?;
    let source_before = fs::read(&source)?;
    let shorter_source_before = fs::read(&shorter_source)?;
    let mismatched_before = fs::read(&mismatched)?;
    let input = format!(
        "1\n3\n{}\n{};{};{}\n0\n0\n",
        template.display(),
        source.display(),
        shorter_source.display(),
        mismatched.display()
    );
    let output = run_menu(&config, &input)?;
    assert!(
        output.status.success(),
        "named-range menu route failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("3  按模板命名区域汇总"));
    assert!(stdout.contains("已选择文件 3 个"));
    assert!(stdout.contains(r#""saved": true"#));
    assert!(!stdout.contains("尚未实现"));

    let logs = root.path().join("app").join("logs");
    let (history_path, history, history_text) = read_single_history(&logs, "PT-6-2_")?;
    assert_eq!(history["feature_id"], "PT-6-2");
    assert_eq!(history["function"], "summary-named-ranges");
    assert_eq!(history["status"], "succeeded");
    assert_eq!(history["summary"]["rows"], 8);
    assert_eq!(history["summary"]["files"], 3);
    assert_eq!(
        history["summary"]["skipped_sources"]
            .as_array()
            .unwrap()
            .len(),
        1
    );
    assert!(history["summary"]["skipped_sources"][0]["reason"]
        .as_str()
        .unwrap()
        .contains("A1"));
    assert!(!history_text.contains("记录甲"));
    assert!(history_path.starts_with(&logs));

    let summary_path = PathBuf::from(history["summary"]["path"].as_str().unwrap());
    assert!(summary_path.starts_with(root.path().join("app").join("output")));
    let summary = read_workbook(&summary_path)?;
    let sheet = summary
        .sheets
        .iter()
        .find(|sheet| sheet.name == "数据")
        .context("named-range output must use the template sheet name")?;
    assert_eq!(sheet.values["C1"]["value"], "资产_贷款余额");
    assert_eq!(sheet.values["F1"]["value"], "指标");
    assert_eq!(sheet.values["G1"]["value"], "金额");
    assert_eq!(sheet.values["C2"]["value"], "甲");
    assert_eq!(sheet.values["D2"]["value"], 0.0);
    assert_eq!(sheet.values["H2"]["value"], 4.0);
    assert_eq!(sheet.values["C3"]["value"], "乙");
    assert_eq!(sheet.values["H3"]["value"], 6.0);
    assert_eq!(sheet.values["D4"]["value"], 0.0);
    assert_eq!(sheet.values["H4"]["value"], 7.0);
    assert_eq!(sheet.values["F5"]["value"], "总计");
    assert_eq!(sheet.values["G5"]["value"], 3.0);
    assert_eq!(sheet.values["H6"]["value"], 11.0);
    assert_eq!(sheet.values["A7"]["value"], "source branch short");
    assert_eq!(sheet.values["C7"]["value"], "甲");
    assert_eq!(sheet.values["H7"]["value"], 4.0);
    assert_eq!(sheet.values["F8"]["value"], "总计");
    assert!(!sheet
        .values
        .values()
        .any(|value| value["value"] == "不应输出"));
    assert_eq!(fs::read(&template)?, template_before);
    assert_eq!(fs::read(&source)?, source_before);
    assert_eq!(fs::read(&shorter_source)?, shorter_source_before);
    assert_eq!(fs::read(&mismatched)?, mismatched_before);
    Ok(())
}

#[test]
fn summary_named_ranges_cli_is_a_real_standalone_entry() -> Result<()> {
    let root = TestDirectory::new()?;
    let template = root.path().join("named template.xlsx");
    let source = root.path().join("source branch.xlsx");
    let output_dir = root.path().join("output");
    create_named_range_template(&template)?;
    create_named_range_source(&source, "结构A", true)?;
    let template_before = fs::read(&template)?;
    let source_before = fs::read(&source)?;
    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("summary-named-ranges")
        .arg("--template")
        .arg(&template)
        .arg("--output-dir")
        .arg(&output_dir)
        .arg("--log-view")
        .arg("off")
        .arg(&source)
        .output()
        .context("run standalone named-range summary CLI")?;
    assert!(
        output.status.success(),
        "named-range CLI failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let result: Value = serde_json::from_slice(&output.stdout)?;
    assert_eq!(result["saved"], true);
    assert_eq!(result["rows"], 5);
    assert_eq!(result["execution"]["reader_engine"], "calamine");
    assert_eq!(result["execution"]["writer_engine"], "rust_xlsxwriter");
    assert_eq!(fs::read(&template)?, template_before);
    assert_eq!(fs::read(&source)?, source_before);
    Ok(())
}

#[test]
fn legacy_comment_summary_is_not_exposed_by_the_v2_data_menu() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_app_config(root.path())?;
    let output = run_menu(&config, "1\n0\n0\n")?;
    assert!(output.status.success());
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(!stdout.contains("按批注汇总"));
    assert!(!stdout.contains("comment summary"));
    Ok(())
}

#[test]
fn summary_comments_cli_matches_python_missing_region_error() -> Result<()> {
    let root = TestDirectory::new()?;
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/summary_comments/merged_children_source.xlsx");
    let output_dir = root.path().join("output");
    let log_dir = root.path().join("logs");
    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("summary-comments")
        .arg("--template")
        .arg(&fixture)
        .arg("--output-dir")
        .arg(&output_dir)
        .arg("--log-dir")
        .arg(&log_dir)
        .arg("--log-view")
        .arg("off")
        .arg(&fixture)
        .output()
        .context("run PT-6-2 with a workbook lacking comment regions")?;

    assert_eq!(output.status.code(), Some(2));
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(
        stderr.contains("模板未识别到可用批注区域（需包含 行区域N/#N 与 列区域N/#N）。"),
        "unexpected PT-6-2 error: {stderr}"
    );
    assert!(
        !output_dir.exists() || fs::read_dir(&output_dir)?.next().is_none(),
        "invalid comment template must not produce an output workbook"
    );
    Ok(())
}

#[test]
fn interactive_menu_returns_failure_when_summary_input_is_corrupt() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_app_config(root.path())?;
    let corrupt = root.path().join("corrupt input.xlsx");
    fs::write(&corrupt, b"not an xlsx workbook")?;

    let input = format!(
        r#"1
2
{}
"#,
        corrupt.display()
    );
    let output = run_menu(&config, &input)?;
    assert!(!output.status.success());
    assert!(
        String::from_utf8_lossy(&output.stderr).contains("pytools:"),
        "CLI failure should be surfaced to stderr"
    );

    let logs = root.path().join("app").join("logs");
    let (_history_path, history, _) = read_single_history(&logs, "PT-6-1_")?;
    assert_eq!(history["status"], "failed");
    assert_eq!(history["failure"]["stage"], "extent_scan");
    Ok(())
}

fn create_app_config(root: &Path) -> Result<PathBuf> {
    let config = root.join("app").join("config").join("config.xlsx");
    fs::create_dir_all(config.parent().context("config parent missing")?)?;
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_name("全局配置")?;
    worksheet.write_string(0, 0, "序号")?;
    worksheet.write_string(0, 1, "配置项")?;
    worksheet.write_string(0, 2, "配置值")?;
    worksheet.write_string(1, 1, "输出目录")?;
    worksheet.write_string(1, 2, "output")?;
    worksheet.write_string(2, 1, "日志目录")?;
    worksheet.write_string(2, 2, "logs")?;
    workbook
        .save(&config)
        .context("save interactive global configuration")?;
    Ok(config)
}

fn create_used_range_source(path: &Path, area: &str, amount: f64) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_name("Data")?;
    worksheet.write_string(0, 0, "名称")?;
    worksheet.write_string(0, 1, "金额")?;
    worksheet.write_string(1, 0, area)?;
    worksheet.write_number(1, 1, amount)?;
    workbook
        .save(path)
        .context("save interactive used-range workbook")
}

fn create_named_range_template(path: &Path) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_name("数据")?;
    worksheet.write_string(0, 0, "结构A")?;
    worksheet.write_string(1, 1, "资产")?;
    worksheet.write_string(2, 1, "贷款余额")?;
    worksheet.write_string(1, 2, "资产")?;
    worksheet.write_string(2, 2, "存款余额")?;
    worksheet.write_string(1, 3, "机构")?;
    worksheet.write_string(2, 3, "名称")?;
    worksheet.write_string(8, 5, "指标")?;
    worksheet.write_string(8, 6, "金额")?;
    workbook.define_name("表结构区域_1", "'数据'!$A$1")?;
    workbook.define_name("表头区域_1", "'数据'!$B$2:$D$3")?;
    workbook.define_name("任意行汇总区域_1", "'数据'!$B$4:$D$5")?;
    workbook.define_name("表头区域_2", "'数据'!$F$9:$G$9")?;
    workbook.define_name("固定行汇总区域_2", "'数据'!$F$10:$G$11")?;
    workbook.save(path).context("save named-range template")
}

fn create_named_range_source(
    path: &Path,
    structure: &str,
    include_second_detail: bool,
) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_name("数据")?;
    worksheet.write_string(0, 0, structure)?;
    worksheet.write_string(3, 1, "甲")?;
    worksheet.write_number(3, 2, 0.0)?;
    worksheet.write_string(3, 3, "记录甲")?;
    if include_second_detail {
        worksheet.write_string(5, 1, "乙")?;
        worksheet.write_number(5, 2, 2.0)?;
        worksheet.write_string(5, 3, "记录乙")?;
        worksheet.write_number(6, 2, 0.0)?;
    }
    worksheet.write_string(9, 5, "总计")?;
    worksheet.write_number(9, 6, 3.0)?;
    worksheet.write_string(499, 25, "不应输出")?;
    let stale_format = rust_xlsxwriter::Format::new().set_bold();
    worksheet.write_blank(399, 1, &stale_format)?;
    workbook.save(path).context("save named-range source")
}

fn run_menu(config: &Path, input: &str) -> Result<Output> {
    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("menu")
        .arg("--config")
        .arg(config)
        .arg("--log-view")
        .arg("developer")
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .context("start interactive pytools menu")?;
    child
        .stdin
        .take()
        .context("interactive stdin missing")?
        .write_all(input.as_bytes())
        .context("send interactive menu selections")?;
    child
        .wait_with_output()
        .context("wait for interactive menu")
}

fn read_single_history(log_dir: &Path, prefix: &str) -> Result<(PathBuf, Value, String)> {
    let path = fs::read_dir(log_dir)?
        .map(|entry| entry.map(|entry| entry.path()))
        .collect::<std::io::Result<Vec<_>>>()?
        .into_iter()
        .find(|path| {
            path.file_name()
                .is_some_and(|name| name.to_string_lossy().starts_with(prefix))
        })
        .with_context(|| format!("no {prefix} history in {}", log_dir.display()))?;
    let bytes = fs::read(&path)?;
    let text = String::from_utf8_lossy(&bytes).into_owned();
    let history = serde_json::from_slice(&bytes).context("parse interactive run history")?;
    Ok((path, history, text))
}
