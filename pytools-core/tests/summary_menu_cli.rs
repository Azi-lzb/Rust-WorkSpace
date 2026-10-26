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
        r#"6
1
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
    assert!(stdout.contains("6  汇总工具"));
    assert!(stdout.contains("1  按使用区域汇总"));
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
fn v2_menu_routes_used_range_summary_and_displays_seven_categories() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_app_config(root.path())?;
    let source = root.path().join("v2 source.xlsx");
    create_used_range_source(&source, "central", 13.0)?;
    let source_before = fs::read(&source)?;
    let input = format!("1\n2\n{}\n0\n", source.display());

    let output = run_menu_v2(&config, &input)?;
    assert!(
        output.status.success(),
        "V2 PT-6-1 entry failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    for label in [
        "PyTools Rust V2",
        "数据提取与汇总",
        "数据去重与追加",
        "报表打印",
        "文件处理",
        "专用报表处理",
        "归档整理",
        "配置与快捷操作",
        "1-2 按使用区域快速汇总（PT-6-1）",
        "可手动输入路径",
    ] {
        assert!(stdout.contains(label), "V2 menu output omitted {label:?}");
    }
    assert!(
        !stdout.contains("按批注汇总（模板+源文件）"),
        "V2 must not expose the retired annotation-based summary"
    );
    assert!(stdout.contains("stable_feature_id=data.used_range_summary"));

    let logs = root.path().join("app").join("logs");
    let (_history_path, history, _) = read_single_history(&logs, "PT-6-1_")?;
    assert_eq!(history["feature_id"], "PT-6-1");
    assert_eq!(history["status"], "succeeded");
    assert_eq!(history["summary"]["files"], 1);
    assert_eq!(history["summary"]["rows"], 2);
    assert_eq!(fs::read(&source)?, source_before);
    assert!(root.path().join("app").join("output").is_dir());
    Ok(())
}

#[test]
fn v2_named_range_summary_is_explicitly_unimplemented() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_app_config(root.path())?;
    let output = run_menu_v2(&config, "1\n3\n0\n")?;
    assert!(output.status.success());
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("1-3 按模板命名区域汇总（尚未实现）"));
    assert!(stdout.contains("[尚未实现] 模板命名区域汇总尚未实现"));
    assert!(!stdout.contains("PT-6-2"));
    assert!(!root.path().join("app").join("output").exists());
    Ok(())
}

#[test]
fn direct_launch_opens_the_seven_category_v2_menu() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_app_config(root.path())?;
    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .env("PYTOOLS_CONFIG_PATH", &config)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .context("start PyTools without a command")?;
    child
        .stdin
        .take()
        .context("direct-launch stdin missing")?
        .write_all(b"0\n")?;
    let output = child
        .wait_with_output()
        .context("wait for direct PyTools launch")?;
    assert!(output.status.success());
    let stdout = String::from_utf8_lossy(&output.stdout);
    for category in [
        "PyTools Rust V2",
        "数据提取与汇总",
        "数据去重与追加",
        "报表打印",
        "文件处理",
        "专用报表处理",
        "归档整理",
        "配置与快捷操作",
    ] {
        assert!(
            stdout.contains(category),
            "direct launch omitted {category:?}"
        );
    }
    Ok(())
}

#[test]
fn interactive_summary_menu_routes_comment_summary_with_template_and_sources() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_app_config(root.path())?;
    let fixtures = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/summary_comments");
    let template = root.path().join("comment template.xlsx");
    let source = root.path().join("comment source.xlsx");
    fs::copy(fixtures.join("comment_template.xlsx"), &template)?;
    fs::copy(fixtures.join("merged_children_source.xlsx"), &source)?;
    let template_before = fs::read(&template)?;
    let source_before = fs::read(&source)?;

    let input = format!(
        r#"6
2
{}
{}
0
"#,
        template.display(),
        source.display()
    );
    let output = run_menu(&config, &input)?;
    assert!(
        output.status.success(),
        "interactive PT-6-2 failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("2  按批注汇总"));
    assert!(stdout.contains(r#""saved": true"#));

    let logs = root.path().join("app").join("logs");
    let (_history_path, history, _history_text) = read_single_history(&logs, "PT-6-2_")?;
    assert_eq!(history["feature_id"], "PT-6-2");
    assert_eq!(history["function"], "summary-comments");
    assert_eq!(history["status"], "succeeded");
    assert_eq!(history["completed_files"], 1);
    assert_eq!(history["summary"]["saved"], true);
    assert_eq!(history["summary"]["files"], 1);
    assert_eq!(history["summary"]["rows"], 3);
    assert!(history["events"].as_array().is_some_and(|events| {
        events.iter().any(|event| {
            event["phase"] == "source_read"
                && event["status"] == "succeeded"
                && event["actual_engine"] == "calamine + namespace-aware OOXML supplement"
        }) && events.iter().any(|event| {
            event["phase"] == "output_save"
                && event["status"] == "succeeded"
                && event["actual_engine"] == "rust_xlsxwriter"
        })
    }));
    assert_eq!(fs::read(&template)?, template_before);
    assert_eq!(fs::read(&source)?, source_before);
    assert!(root.path().join("app").join("output").is_dir());
    Ok(())
}

#[test]
fn interactive_comment_summary_formats_elapsed_duration_header_as_python_timedelta() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_app_config(root.path())?;
    let template = root.path().join("duration template.xlsx");
    let source = root.path().join("duration source.xlsx");
    create_elapsed_duration_template(&template)?;
    create_elapsed_duration_source(&source)?;
    let template_before = fs::read(&template)?;
    let source_before = fs::read(&source)?;

    let input = format!("6\n2\n{}\n{}\n0\n", template.display(), source.display());
    let output = run_menu(&config, &input)?;
    assert!(
        output.status.success(),
        "interactive PT-6-2 duration summary failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let logs = root.path().join("app").join("logs");
    let (_history_path, history, _) = read_single_history(&logs, "PT-6-2_")?;
    assert_eq!(history["status"], "succeeded");
    assert_eq!(history["summary"]["rows"], 3);
    assert!(history["events"].as_array().is_some_and(|events| {
        events.iter().any(|event| {
            event["phase"] == "template_read"
                && event["status"] == "succeeded"
                && event["actual_engine"] == "calamine + namespace-aware OOXML supplement"
        }) && events.iter().any(|event| {
            event["phase"] == "output_save"
                && event["status"] == "succeeded"
                && event["actual_engine"] == "rust_xlsxwriter"
        })
    }));

    let summary_path = PathBuf::from(
        history["summary"]["path"]
            .as_str()
            .context("duration summary path is missing from history")?,
    );
    let summary = read_workbook(&summary_path)?;
    let sheet = summary
        .sheets
        .iter()
        .find(|sheet| sheet.name == "Branch_A")
        .context("duration summary output sheet is missing")?;
    assert_eq!(
        sheet.values.get("C1"),
        Some(&serde_json::json!({
            "type": "string",
            "value": "1 day, 3:15:20.123000"
        })),
        "interactive summary must format elapsed header values like Python timedelta"
    );
    assert_eq!(fs::read(&template)?, template_before, "template changed");
    assert_eq!(fs::read(&source)?, source_before, "source workbook changed");
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
        r#"6
1
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

fn create_elapsed_duration_template(path: &Path) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_name("Branch_A")?;
    worksheet.write_string(1, 0, "row start")?;
    worksheet.insert_note(1, 0, &rust_xlsxwriter::Note::new("行区域1"))?;
    worksheet.write_string(3, 0, "row end")?;
    worksheet.insert_note(3, 0, &rust_xlsxwriter::Note::new("行区域#1"))?;
    let elapsed_format = rust_xlsxwriter::Format::new().set_num_format("[h]:mm:ss.000");
    worksheet.write_number_with_format(0, 1, 1.1356495717592594, &elapsed_format)?;
    worksheet.insert_note(0, 1, &rust_xlsxwriter::Note::new("列区域1"))?;
    worksheet.insert_note(1, 1, &rust_xlsxwriter::Note::new("列区域#1"))?;
    workbook
        .save(path)
        .context("save interactive elapsed duration template")
}

fn create_elapsed_duration_source(path: &Path) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_name("Branch_A")?;
    for (row, amount) in [(1, 10.5), (2, 20.0), (3, 0.0)] {
        worksheet.write_number(row, 1, amount)?;
    }
    workbook
        .save(path)
        .context("save interactive elapsed duration source")
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

fn run_menu_v2(config: &Path, input: &str) -> Result<Output> {
    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("menu-v2")
        .arg("--config")
        .arg(config)
        .arg("--log-view")
        .arg("developer")
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .context("start PyTools Rust V2 menu")?;
    child
        .stdin
        .take()
        .context("V2 menu stdin missing")?
        .write_all(input.as_bytes())
        .context("send V2 menu selections")?;
    child
        .wait_with_output()
        .context("wait for PyTools Rust V2 menu")
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
