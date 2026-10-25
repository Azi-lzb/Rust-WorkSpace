use anyhow::{Context, Result};
use excel_editor::{
    BackendCapability, CapabilityEvidence, CapabilityStatus, EditOperation, EngineId, ExcelEditPlan,
};
use excel_reader::read_workbook;
use pytools_core::dedup_ops::AppendTargetPlan;
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
            "pytools-timeline-menu-{}-{nonce}-{serial}",
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
fn interactive_timeline_menu_routes_compare_wide_modes_and_precheck() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_timeline_config(root.path())?;
    let template = root.path().join("template.xlsx");
    let compare_source = root.path().join("sample compare source.xlsx");
    let wide_source = root.path().join("sample wide source.xlsx");
    create_wide_source(&template, "模板甲", 10.0)?;
    create_wide_source(&compare_source, "来源甲", 12.0)?;
    create_wide_source(&wide_source, "宽表甲", 17.5)?;

    let compare = run_menu(
        &config,
        &format!(
            "1
1
{}
{}
0
",
            template.display(),
            compare_source.display()
        ),
    )?;
    assert!(
        compare.status.success(),
        "timeline compare menu failed: {}",
        String::from_utf8_lossy(&compare.stderr)
    );
    let compare_stdout = String::from_utf8_lossy(&compare.stdout);
    assert!(compare_stdout.contains("1  时序工具"));
    assert!(compare_stdout.contains("1  表头路径比对"));
    assert!(compare_stdout.contains("表头路径比对模板"));
    assert!(compare_stdout.contains(r#""output""#));
    let output_dir = root.path().join("app").join("output");
    assert!(find_output(&output_dir, "compare_3_9_3_")?.is_file());

    let row_suffix = run_menu(
        &config,
        &format!(
            "1
3
{}
0
",
            wide_source.display()
        ),
    )?;
    assert!(
        row_suffix.status.success(),
        "row-suffix wide summary menu failed: {}",
        String::from_utf8_lossy(&row_suffix.stderr)
    );
    assert!(String::from_utf8_lossy(&row_suffix.stdout).contains("3  宽表汇总"));
    assert!(String::from_utf8_lossy(&row_suffix.stdout).contains(r#""output""#));
    let row_suffix_output = find_output(&output_dir, "行头加后缀")?;
    let row_suffix_dump = read_workbook(&row_suffix_output)?;
    let row_suffix_sheet = row_suffix_dump
        .sheets
        .iter()
        .find(|sheet| sheet.name == "WideRule")
        .context("row-suffix summary sheet missing")?;
    assert!(!row_suffix_sheet.values.is_empty());

    let no_suffix = run_menu(
        &config,
        &format!(
            "1
4
{}
0
",
            wide_source.display()
        ),
    )?;
    assert!(
        no_suffix.status.success(),
        "no-suffix wide summary menu failed: {}",
        String::from_utf8_lossy(&no_suffix.stderr)
    );
    assert!(String::from_utf8_lossy(&no_suffix.stdout).contains("4  宽表汇总"));
    let no_suffix_output = find_output(&output_dir, "行头不加后缀")?;
    let no_suffix_dump = read_workbook(&no_suffix_output)?;
    assert!(no_suffix_dump
        .sheets
        .iter()
        .any(|sheet| sheet.name == "WideRule" && !sheet.values.is_empty()));

    let precheck = run_menu(
        &config, "1
5
0
",
    )?;
    assert!(
        precheck.status.success(),
        "timeline precheck menu failed: {}",
        String::from_utf8_lossy(&precheck.stderr)
    );
    let precheck_stdout = String::from_utf8_lossy(&precheck.stdout);
    assert!(precheck_stdout.contains("5  预校验"));
    assert!(precheck_stdout.contains(r#""errors": 0"#));
    assert!(precheck_stdout.contains(r#""warnings": 0"#));
    Ok(())
}

#[test]
fn wide_summary_cli_filters_missing_and_unconfigured_extensions_before_reading() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_timeline_config(root.path())?;
    let source = root.path().join("sample valid wide.xlsx");
    let unsupported = root.path().join("sample unsupported.xlsb");
    let missing = root.path().join("sample missing.xlsx");
    create_wide_source(&source, "宽表甲", 17.5)?;
    fs::write(&unsupported, b"not a supported workbook")?;
    let source_before = fs::read(&source)?;
    let unsupported_before = fs::read(&unsupported)?;

    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("wide-summary")
        .arg("--config")
        .arg(&config)
        .arg("--log-view")
        .arg("developer")
        .arg(&source)
        .arg(&unsupported)
        .arg(&missing)
        .output()
        .context("run wide-summary CLI with filtered inputs")?;
    assert!(
        output.status.success(),
        "wide-summary CLI did not filter unsupported inputs: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let report: Value = serde_json::from_slice(&output.stdout)?;
    let summary_path = report["output"]
        .as_str()
        .map(PathBuf::from)
        .context("wide-summary CLI output path missing")?;
    assert!(summary_path.is_file());
    let run_log_path = report["run_log_path"]
        .as_str()
        .map(PathBuf::from)
        .context("wide-summary run log path missing")?;
    let history: Value = serde_json::from_slice(&fs::read(run_log_path)?)?;
    assert_eq!(history["total_inputs"], 1);
    assert_eq!(history["completed_inputs"], 1);
    assert_eq!(history["failed_inputs"], 0);
    assert!(history["events"]
        .as_array()
        .context("wide-summary events missing")?
        .iter()
        .all(|event| event["progress"]["total_inputs"] == 1));
    assert_eq!(fs::read(&source)?, source_before);
    assert_eq!(fs::read(&unsupported)?, unsupported_before);
    Ok(())
}

#[test]
fn wide_summary_cli_projects_elapsed_duration_as_numeric_for_both_row_modes() -> Result<()> {
    for command in ["wide-summary", "wide-summary-row-suffix"] {
        let root = TestDirectory::new()?;
        let config = create_timeline_config(root.path())?;
        let source = root.path().join("sample elapsed duration.xlsx");
        create_wide_source_with_elapsed_duration(&source, "累计时长", 1.123_456_789)?;

        let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
            .arg(command)
            .arg("--config")
            .arg(&config)
            .arg("--log-view")
            .arg("developer")
            .arg(&source)
            .output()
            .context("run elapsed-duration wide-summary CLI")?;
        assert!(
            output.status.success(),
            "{command} did not process elapsed duration: {}",
            String::from_utf8_lossy(&output.stderr)
        );
        let report: Value = serde_json::from_slice(&output.stdout)?;
        let summary_path = report["output"]
            .as_str()
            .map(PathBuf::from)
            .context("elapsed-duration summary output path missing")?;
        let dump = read_workbook(&summary_path)?;
        let summary = dump
            .sheets
            .iter()
            .find(|sheet| sheet.name == "WideRule")
            .context("elapsed-duration summary sheet missing")?;
        let cell = &summary.values["G2"];
        assert_eq!(cell["type"], "number");
        let actual = cell["value"]
            .as_f64()
            .context("duration serial is not numeric")?;
        assert!((actual - 1.123_456_793_981_481).abs() < 1e-15);

        let history_path = report["run_log_path"]
            .as_str()
            .map(PathBuf::from)
            .context("elapsed-duration run history missing")?;
        let history: Value = serde_json::from_slice(&fs::read(history_path)?)?;
        assert!(history["events"].as_array().is_some_and(|events| {
            events.iter().any(|event| {
                event["phase"] == "elapsed_duration_normalization"
                    && event["status"] == "completed"
                    && event["actual_engine"] == "pytools-core"
                    && event["stage_elapsed_ms"].as_u64().is_some()
            })
        }));
    }
    Ok(())
}

#[test]
fn wide_summary_cli_discovers_default_sources_for_both_row_modes() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_timeline_config(root.path())?;
    let input_dir = root.path().join("app").join("input");
    let nested_dir = input_dir.join("nested");
    fs::create_dir_all(&nested_dir)?;

    create_wide_source(
        &input_dir.join("sample root source.xlsx"),
        "默认发现根目录",
        21.0,
    )?;
    create_wide_source(
        &nested_dir.join("sample nested source.XLSX"),
        "默认发现子目录",
        31.0,
    )?;
    create_wide_source(
        &input_dir.join("~$sample locked source.xlsx"),
        "锁文件必须跳过",
        41.0,
    )?;
    fs::write(
        input_dir.join("sample unsupported source.xlsb"),
        b"not xlsb",
    )?;

    for command in ["wide-summary", "wide-summary-row-suffix"] {
        let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
            .arg(command)
            .arg("--config")
            .arg(&config)
            .arg("--log-view")
            .arg("developer")
            .output()
            .with_context(|| format!("run {command} with configured source discovery"))?;
        assert!(
            output.status.success(),
            "{command} default source discovery failed: {}",
            String::from_utf8_lossy(&output.stderr)
        );

        let report: Value = serde_json::from_slice(&output.stdout)?;
        let summary_path = report["output"]
            .as_str()
            .map(PathBuf::from)
            .with_context(|| format!("{command} output path missing"))?;
        let summary = read_workbook(&summary_path)?;
        assert_eq!(summary.sheets.len(), 1);
        assert!(summary.sheets[0].values.len() > 10);

        let history: Value = serde_json::from_slice(&fs::read(
            report["run_log_path"]
                .as_str()
                .with_context(|| format!("{command} run history path missing"))?,
        )?)?;
        assert_eq!(history["total_inputs"], 2);
        assert_eq!(history["completed_inputs"], 2);
        assert_eq!(history["failed_inputs"], 0);
        let events = history["events"]
            .as_array()
            .context("wide-summary events missing")?;
        assert!(events.iter().any(|event| {
            event["phase"] == "source_discovery"
                && event["status"] == "completed"
                && event["progress"]["total_inputs"] == 2
        }));
        assert!(events
            .iter()
            .enumerate()
            .all(|(index, event)| { event["sequence"].as_u64() == Some(index as u64 + 1) }));
    }
    Ok(())
}

#[test]
fn wide_summary_cli_resolves_relative_configured_directories_for_discovery() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_timeline_config_with_relative_dirs(root.path())?;
    let input_dir = root.path().join("app").join("input");
    let nested_dir = input_dir.join("nested");
    fs::create_dir_all(&nested_dir)?;
    create_wide_source(
        &input_dir.join("sample relative root source.xlsx"),
        "相对目录根",
        21.0,
    )?;
    create_wide_source(
        &nested_dir.join("sample relative nested source.xlsx"),
        "相对目录子",
        31.0,
    )?;
    create_wide_source(
        &input_dir.join("~$sample relative locked source.xlsx"),
        "锁文件必须跳过",
        41.0,
    )?;
    fs::write(
        input_dir.join("sample relative unsupported source.xlsb"),
        b"not xlsb",
    )?;

    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("wide-summary")
        .arg("--config")
        .arg(&config)
        .arg("--log-view")
        .arg("developer")
        .output()
        .context("run wide-summary with relative configured directories")?;
    assert!(
        output.status.success(),
        "relative directory wide-summary failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let report: Value = serde_json::from_slice(&output.stdout)?;
    let summary_path = report["output"]
        .as_str()
        .map(PathBuf::from)
        .with_context(|| "relative wide-summary output path missing")?;
    let output_dir = root.path().join("app").join("output");
    assert!(
        summary_path.starts_with(&output_dir),
        "summary must be written under the config-relative output directory: {summary_path:?}"
    );
    let summary = read_workbook(&summary_path)?;
    assert_eq!(summary.sheets.len(), 1);
    assert!(summary.sheets[0].values.len() > 10);

    let run_log_path = report["run_log_path"]
        .as_str()
        .map(PathBuf::from)
        .with_context(|| "relative wide-summary run log path missing")?;
    let log_dir = root.path().join("app").join("logs");
    assert!(
        run_log_path.starts_with(&log_dir),
        "run history must be written under the config-relative log directory: {run_log_path:?}"
    );
    let history: Value = serde_json::from_slice(&fs::read(&run_log_path)?)?;
    assert_eq!(history["total_inputs"], 2);
    assert_eq!(history["completed_inputs"], 2);
    assert_eq!(history["failed_inputs"], 0);

    Ok(())
}

#[test]
fn timeline_extract_cli_discovers_default_sources_recursively_and_skips_lock_files() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_timeline_config(root.path())?;
    let input_dir = root.path().join("app").join("input");
    let nested_dir = input_dir.join("nested");
    fs::create_dir_all(&nested_dir)?;

    let root_source = input_dir.join("sample root source.xlsx");
    let nested_source = nested_dir.join("sample nested source.XLSX");
    let lock_source = input_dir.join("~$sample locked source.xlsx");
    let unsupported_source = input_dir.join("sample unsupported source.xlsb");
    create_wide_source(&root_source, "自动发现根目录", 21.0)?;
    create_wide_source(&nested_source, "自动发现子目录", 31.0)?;
    create_wide_source(&lock_source, "锁文件必须跳过", 41.0)?;
    fs::write(&unsupported_source, b"unsupported workbook format")?;

    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("timeline-extract")
        .arg("--config")
        .arg(&config)
        .arg("--log-view")
        .arg("developer")
        .output()
        .context("run timeline-extract with configured input discovery")?;
    assert!(
        output.status.success(),
        "timeline-extract default discovery failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let report: Value = serde_json::from_slice(&output.stdout)?;
    assert_eq!(report["total_files"], 2);
    assert_eq!(report["processed_files"], 2);
    assert_eq!(report["failed_files"], 0);
    assert!(report["rows"].as_u64().is_some_and(|rows| rows > 0));
    let output_path = report["path"]
        .as_str()
        .map(PathBuf::from)
        .context("timeline extraction output path missing")?;
    let workbook = read_workbook(&output_path)?;
    let output_values = &workbook.sheets[0].values;
    for expected_source in ["sample root source.xlsx", "sample nested source.XLSX"] {
        assert!(
            output_values
                .values()
                .any(|cell| { cell.get("value").and_then(Value::as_str) == Some(expected_source) }),
            "default discovery output missing source {expected_source}"
        );
    }
    assert!(!output_values.values().any(|cell| {
        cell.get("value").and_then(Value::as_str) == Some("~$sample locked source.xlsx")
    }));

    let history: Value = serde_json::from_slice(&fs::read(
        report["run_log_path"]
            .as_str()
            .context("timeline extraction history path missing")?,
    )?)?;
    assert_eq!(history["progress"]["total_inputs"], 2);
    assert!(history["events"].as_array().is_some_and(|events| {
        events.iter().any(|event| {
            event["phase"] == "source_discovery"
                && event["status"] == "completed"
                && event["progress"]["total_inputs"] == 2
        })
    }));
    Ok(())
}

#[test]
fn timeline_extract_cli_with_empty_default_input_directory_skips_output() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_timeline_config(root.path())?;
    let output = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("timeline-extract")
        .arg("--config")
        .arg(&config)
        .arg("--log-view")
        .arg("developer")
        .output()
        .context("run timeline-extract with no configured source files")?;
    assert!(
        output.status.success(),
        "timeline-extract empty discovery failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );

    let report: Value = serde_json::from_slice(&output.stdout)?;
    assert_eq!(report["total_files"], 0);
    assert_eq!(report["rows"], 0);
    assert!(report["path"].is_null());
    assert!(!root.path().join("app").join("output").exists());
    Ok(())
}

#[test]
fn header_path_compare_menu_supports_user_developer_and_off_views() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_timeline_config(root.path())?;
    let template = root.path().join("compare template.xlsx");
    let source = root.path().join("sample compare source.xlsx");
    let broken = root.path().join("broken compare source.xlsx");
    create_wide_source(&template, "模板甲", 10.0)?;
    create_wide_source(&source, "来源甲", 12.0)?;
    fs::write(&broken, b"not an xlsx workbook")?;
    let log_dir = root.path().join("app").join("logs");

    let user = run_menu_with_view(
        &config,
        "user",
        &format!("1\n1\n{}\n{}\n0\n", template.display(), source.display()),
    )?;
    assert!(
        user.status.success(),
        "user-view compare failed: {}",
        String::from_utf8_lossy(&user.stderr)
    );
    let user_stderr = String::from_utf8_lossy(&user.stderr);
    assert!(user_stderr.contains("正在读取表头比对配置"));
    assert!(user_stderr.contains("表头比对模板读取完成"));
    assert!(user_stderr.contains("已处理 1/1 个源工作簿"));
    assert!(user_stderr.contains("表头路径比对结束"));
    assert!(user_stderr.contains("阶段用时"));
    assert!(!user_stderr.contains(&source.display().to_string()));
    assert!(!user_stderr.contains("来源甲"));
    let user_report = latest_timeline_compare_report(&user)?;
    assert!(user_report.get("output").is_some());
    assert!(user_report.get("run_history_path").is_none());
    assert!(!log_dir.exists() || run_history_paths(&log_dir)?.is_empty());

    let developer_sources = format!(
        "{};{};{}",
        source.display(),
        broken.display(),
        template.display()
    );
    let developer = run_menu_with_view(
        &config,
        "developer",
        &format!("1\n1\n{}\n{}\n0\n", template.display(), developer_sources),
    )?;
    assert!(
        developer.status.success(),
        "developer-view compare failed: {}",
        String::from_utf8_lossy(&developer.stderr)
    );
    let developer_report = latest_timeline_compare_report(&developer)?;
    let history_path = PathBuf::from(
        developer_report["run_history_path"]
            .as_str()
            .context("developer run-history path missing")?,
    );
    let history_bytes = fs::read(&history_path)?;
    let history_text = String::from_utf8_lossy(&history_bytes);
    assert!(!history_text.contains(&source.display().to_string()));
    assert!(!history_text.contains(&broken.display().to_string()));
    assert!(!history_text.contains("来源甲"));
    assert!(!history_text.contains("not an xlsx workbook"));
    let history: Value = serde_json::from_slice(&history_bytes)?;
    assert_eq!(history["feature_id"], "PT-1-1");
    assert_eq!(history["status"], "partial");
    assert_eq!(history["progress"]["total_inputs"], 3);
    assert_eq!(history["progress"]["completed_inputs"], 1);
    assert_eq!(history["progress"]["failed_inputs"], 1);
    assert_eq!(history["progress"]["skipped_inputs"], 1);
    for phase in [
        "configuration_read",
        "template_read",
        "source_read",
        "source_compare",
        "report_write",
        "task",
    ] {
        assert!(
            history["events"]
                .as_array()
                .is_some_and(|events| events.iter().any(|event| event["phase"] == phase)),
            "missing {phase} event"
        );
    }
    assert!(history["actual_engines"].as_array().is_some_and(|engines| {
        engines.iter().any(|engine| engine == "calamine")
            && engines.iter().any(|engine| engine == "rust_xlsxwriter")
    }));
    let developer_events = String::from_utf8_lossy(&developer.stderr)
        .lines()
        .filter_map(|line| serde_json::from_str::<Value>(line).ok())
        .collect::<Vec<_>>();
    assert!(developer_events.len() >= 10);
    assert!(developer_events.iter().all(|event| {
        event["stage_elapsed_ms"].is_number() && event["total_elapsed_ms"].is_number()
    }));

    let history_count_before_off = run_history_paths(&log_dir)?.len();
    let off = run_menu_with_view(
        &config,
        "off",
        &format!("1\n1\n{}\n{}\n0\n", template.display(), source.display()),
    )?;
    assert!(
        off.status.success(),
        "off-view compare failed: {}",
        String::from_utf8_lossy(&off.stderr)
    );
    assert!(String::from_utf8_lossy(&off.stderr).is_empty());
    assert!(latest_timeline_compare_report(&off)?
        .get("run_history_path")
        .is_none());
    assert_eq!(run_history_paths(&log_dir)?.len(), history_count_before_off);
    Ok(())
}

#[test]
fn header_path_compare_formats_date_headers_as_python_datetime_text() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_timeline_config(root.path())?;
    let template = root.path().join("sample template.xlsx");
    let source = root.path().join("sample source.xlsx");
    create_date_header_source(&template, 2024, 1, 31)?;
    create_date_header_source(&source, 2024, 2, 29)?;

    let input = format!("1\n1\n{}\n{}\n0\n", template.display(), source.display());
    let output = run_menu_with_view(&config, "off", &input)?;
    assert!(
        output.status.success(),
        "date-header compare failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report = latest_timeline_compare_report(&output)?;
    let result_path = PathBuf::from(
        report["output"]
            .as_str()
            .context("date-header comparison output path missing")?,
    );
    let workbook = read_workbook(result_path)?;
    let differences = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "表头比对结果")
        .context("date-header difference sheet missing")?;
    assert_eq!(
        differences.values["H2"]["value"],
        "存款_2024-02-29 00:00:00"
    );
    assert_eq!(
        differences.values["I2"]["value"],
        "存款_2024-01-31 00:00:00"
    );
    assert_eq!(
        differences.values["H3"]["value"],
        "存款_2024-02-29 00:00:00"
    );
    assert_eq!(
        differences.values["I4"]["value"],
        "存款_2024-01-31 00:00:00"
    );
    Ok(())
}

#[test]
fn header_path_compare_menu_records_terminal_failure_summary() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_timeline_config(root.path())?;
    let template = root.path().join("broken template.xlsx");
    let source = root.path().join("valid source.xlsx");
    fs::write(&template, b"not an xlsx workbook")?;
    create_wide_source(&source, "来源甲", 12.0)?;

    let user = run_menu_with_view(
        &config,
        "user",
        &format!("1\n1\n{}\n{}\n0\n", template.display(), source.display()),
    )?;
    assert!(!user.status.success());
    let user_stderr = String::from_utf8_lossy(&user.stderr);
    assert!(user_stderr.contains("表头比对模板读取失败"));
    assert!(user_stderr.contains("表头路径比对任务失败"));
    assert!(!user_stderr.contains(&template.display().to_string()));
    assert!(!user_stderr.contains("not an xlsx workbook"));
    assert!(!user_stderr.contains("calamine"));

    let log_dir = root.path().join("app").join("logs");
    assert!(!log_dir.exists() || run_history_paths(&log_dir)?.is_empty());

    let developer = run_menu_with_view(
        &config,
        "developer",
        &format!("1\n1\n{}\n{}\n0\n", template.display(), source.display()),
    )?;
    assert!(!developer.status.success());
    let history = read_single_run_history(&log_dir, "PT-1-1")?;
    assert_eq!(history["status"], "failed");
    assert_eq!(history["error_type"], "template_read_error");
    let history_text = serde_json::to_string(&history)?;
    assert!(!history_text.contains(&template.display().to_string()));
    assert!(!history_text.contains("not an xlsx workbook"));
    assert!(history["events"].as_array().is_some_and(|events| {
        events
            .iter()
            .any(|event| event["phase"] == "template_read" && event["status"] == "failed")
            && events.iter().any(|event| {
                event["phase"] == "task"
                    && event["status"] == "failed"
                    && event["error_type"] == "template_read_error"
            })
    }));
    Ok(())
}

#[test]
fn interactive_wide_summary_menu_commits_new_target_after_confirmation() -> Result<()> {
    let root = TestDirectory::new()?;
    let target = root.path().join("configured target.xlsx");
    let config = create_timeline_config_with_target(root.path(), Some(&target))?;
    let source = root.path().join("sample wide source.xlsx");
    create_wide_source(&source, "交互汇总甲", 23.5)?;

    let completed = run_menu(&config, &format!("1\n3\n{}\ny\n0\n", source.display()))?;
    assert!(
        completed.status.success(),
        "confirmed target append failed: {}",
        String::from_utf8_lossy(&completed.stderr)
    );
    assert!(target.is_file(), "confirmed new target was not published");
    assert!(String::from_utf8_lossy(&completed.stdout).contains("新建目标簿"));
    assert!(String::from_utf8_lossy(&completed.stdout).contains("目标写入完成"));

    let target_dump = read_workbook(&target)?;
    let target_sheet = target_dump
        .sheets
        .iter()
        .find(|sheet| sheet.name == "Append")
        .context("configured target worksheet was not created")?;
    assert!(target_sheet
        .values
        .values()
        .any(|cell| { cell.get("value").and_then(Value::as_str) == Some("交互汇总甲") }));

    let report = latest_menu_json_report(&completed)?;
    assert_eq!(report["target_append_status"], "succeeded");
    assert_eq!(
        report["target_append_executions"][0]["operation"],
        "created"
    );
    assert!(report["target_append_executions"][0]["engine"].is_null());
    let history_path = report["run_log_path"]
        .as_str()
        .context("wide summary run history path missing")?;
    let history_bytes = fs::read(history_path)?;
    let history: Value = serde_json::from_slice(&history_bytes)?;
    assert!(history["events"].as_array().is_some_and(|events| {
        events.iter().any(|event| {
            event["phase"] == "target_append_execute"
                && event["status"] == "succeeded"
                && event["actual_engine"] == "rust_xlsxwriter"
                && event["stage_elapsed_ms"].as_u64().is_some()
        })
    }));
    let history_text = String::from_utf8_lossy(&history_bytes);
    assert!(!history_text.contains(&source.display().to_string()));
    assert!(!history_text.contains(&target.display().to_string()));
    assert!(!history_text.contains("交互汇总甲"));
    Ok(())
}

#[test]
fn interactive_wide_summary_menu_preserves_negative_date_serial_in_new_target() -> Result<()> {
    assert_negative_date_serial_menu_target("3")
}

#[test]
fn interactive_wide_summary_no_suffix_menu_preserves_negative_date_serial_in_new_target(
) -> Result<()> {
    assert_negative_date_serial_menu_target("4")
}

fn assert_negative_date_serial_menu_target(menu_choice: &str) -> Result<()> {
    let root = TestDirectory::new()?;
    let target = root.path().join("configured target.xlsx");
    let config = create_timeline_config_with_target(root.path(), Some(&target))?;
    let config_before = fs::read(&config)?;
    let source = root.path().join("sample negative date source.xlsx");
    create_wide_source_with_negative_date(&source, "负序列分行")?;
    let source_before = fs::read(&source)?;

    let completed = run_menu(
        &config,
        &format!("1\n{menu_choice}\n{}\ny\n0\n", source.display()),
    )?;
    assert!(
        completed.status.success(),
        "wide-summary negative serial target write failed: {}",
        String::from_utf8_lossy(&completed.stderr)
    );
    assert!(target.is_file(), "negative serial target was not published");
    assert_eq!(fs::read(&source)?, source_before);
    assert_eq!(fs::read(&config)?, config_before);

    let target_dump = read_workbook(&target)?;
    let negative_serial_exists = target_dump.sheets.iter().any(|sheet| {
        sheet.values.iter().any(|(_, value)| {
            value.get("type").and_then(Value::as_str) == Some("datetime_serial")
                && value
                    .get("value")
                    .and_then(Value::as_f64)
                    .is_some_and(|serial| (serial + 0.5).abs() < 1e-9)
        })
    });
    assert!(
        negative_serial_exists,
        "negative cached Excel serial was not preserved: {:?}",
        target_dump.sheets
    );
    let report = latest_menu_json_report(&completed)?;
    assert_eq!(report["target_append_status"], "succeeded");
    assert_eq!(
        report["target_append_executions"][0]["operation"],
        "created"
    );
    let feature_id = if menu_choice == "3" {
        "PT-1-3"
    } else {
        "PT-1-4"
    };
    let history = read_single_run_history(&root.path().join("app").join("logs"), feature_id)?;
    assert_eq!(history["status"], "succeeded");
    Ok(())
}

#[test]
fn interactive_wide_summary_menu_does_not_commit_without_confirmation() -> Result<()> {
    let root = TestDirectory::new()?;
    let target = root.path().join("configured target.xlsx");
    let config = create_timeline_config_with_target(root.path(), Some(&target))?;
    let source = root.path().join("sample wide source.xlsx");
    create_wide_source(&source, "交互汇总乙", 24.5)?;

    let declined = run_menu(&config, &format!("1\n3\n{}\nn\n0\n", source.display()))?;
    assert!(
        declined.status.success(),
        "declining target append failed: {}",
        String::from_utf8_lossy(&declined.stderr)
    );
    assert!(!target.exists(), "declined new target was published");
    assert!(String::from_utf8_lossy(&declined.stdout).contains("已取消目标写入"));
    let report = latest_menu_json_report(&declined)?;
    assert_eq!(report["target_append_executions"][0]["status"], "cancelled");
    Ok(())
}

#[test]
fn interactive_wide_summary_existing_target_requires_capability_and_commits_safely() -> Result<()> {
    let root = TestDirectory::new()?;
    let target = root.path().join("configured target.xlsx");
    let config = create_timeline_config_with_target(root.path(), Some(&target))?;
    let first_source = root.path().join("sample first source.xlsx");
    let second_source = root.path().join("sample second source.xlsx");
    create_wide_source_with_negative_date(&first_source, "交互汇总甲")?;
    create_wide_source(&second_source, "交互汇总乙", 31.5)?;

    let created = run_menu(
        &config,
        &format!("1\n3\n{}\ny\n0\n", first_source.display()),
    )?;
    assert!(created.status.success());
    let original_target = fs::read(&target)?;

    let without_capability = run_menu(&config, &format!("1\n3\n{}\n0\n", second_source.display()))?;
    assert!(
        without_capability.status.success(),
        "plan-only existing target route failed: {}",
        String::from_utf8_lossy(&without_capability.stderr)
    );
    assert!(String::from_utf8_lossy(&without_capability.stdout).contains("[仅生成计划]"));
    assert_eq!(fs::read(&target)?, original_target);

    let planned = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["wide-summary", "--config"])
        .arg(&config)
        .args(["--log-view", "off"])
        .arg(&second_source)
        .output()?;
    assert!(
        planned.status.success(),
        "failed to prepare fixture capability plan: {}",
        String::from_utf8_lossy(&planned.stderr)
    );
    let planned_report: Value = serde_json::from_slice(&planned.stdout)?;
    let plan_path = planned_report["target_append_plans"][0]
        .as_str()
        .context("existing target plan path missing")?;
    let plan: AppendTargetPlan = serde_json::from_slice(&fs::read(plan_path)?)?;
    let edit_plan = plan.edit_plan.as_ref().context("edit plan missing")?;
    let capability = capability_for_menu_test(edit_plan);
    let capability_path = root.path().join("fixture-scoped-capability.json");
    fs::write(&capability_path, serde_json::to_vec_pretty(&capability)?)?;

    let committed = run_menu_with_capabilities(
        &config,
        "user",
        &capability_path,
        "synthetic-pt-1-wide-menu-umya",
        &format!("1\n3\n{}\ny\n0\n", second_source.display()),
    )?;
    assert!(
        committed.status.success(),
        "capability-gated menu commit failed: {}",
        String::from_utf8_lossy(&committed.stderr)
    );
    let updated_target = read_workbook(&target)?;
    let append_sheet = updated_target
        .sheets
        .iter()
        .find(|sheet| sheet.name == "Append")
        .context("updated target worksheet missing")?;
    assert!(append_sheet
        .values
        .values()
        .any(|cell| { cell.get("value").and_then(Value::as_str) == Some("交互汇总甲") }));
    assert!(append_sheet
        .values
        .values()
        .any(|cell| { cell.get("value").and_then(Value::as_str) == Some("交互汇总乙") }));
    let report = latest_menu_json_report(&committed)?;
    assert_eq!(report["target_append_status"], "succeeded");
    assert_eq!(
        report["target_append_executions"][0]["operation"],
        "replaced"
    );
    assert!(report["target_append_executions"][0]["engine"].is_null());
    let history: Value = serde_json::from_slice(&fs::read(
        report["run_log_path"]
            .as_str()
            .context("wide summary history missing")?,
    )?)?;
    assert!(history["actual_engines"]
        .as_array()
        .is_some_and(|engines| engines.iter().any(|engine| engine == "umya_spreadsheet")));
    assert!(history["events"].as_array().is_some_and(|events| {
        events.iter().any(|event| {
            event["phase"] == "target_append_execute"
                && event["status"] == "succeeded"
                && event["actual_engine"] == "umya_spreadsheet"
        })
    }));
    Ok(())
}

#[test]
fn wide_summary_menu_logging_views_record_progress_and_respect_off() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_timeline_config(root.path())?;
    let source = root.path().join("sample wide source.xlsx");
    create_wide_source(&source, "宽表甲", 17.5)?;

    let user = run_menu_with_view(&config, "user", &format!("1\n3\n{}\n0\n", source.display()))?;
    assert!(
        user.status.success(),
        "user-view wide summary failed: {}",
        String::from_utf8_lossy(&user.stderr)
    );
    let user_stderr = String::from_utf8_lossy(&user.stderr);
    assert!(user_stderr.contains("正在读取第 1/1 个工作簿"));
    assert!(!user_stderr.contains("\"phase\""));

    let log_dir = root.path().join("app").join("logs");
    let row_suffix_history = read_single_run_history(&log_dir, "PT-1-3")?;
    assert_eq!(row_suffix_history["status"], "succeeded");
    assert_eq!(row_suffix_history["completed_inputs"], 1);
    assert_eq!(row_suffix_history["failed_inputs"], 0);
    assert!(row_suffix_history["actual_engines"]
        .as_array()
        .is_some_and(|engines| engines.iter().any(|engine| engine == "calamine")
            && engines.iter().any(|engine| engine == "rust_xlsxwriter")));
    assert_history_has_phase(&row_suffix_history, "source_read", "completed", "calamine")?;
    assert_history_has_phase(&row_suffix_history, "rule_aggregation", "completed", "")?;
    assert_history_has_phase(
        &row_suffix_history,
        "summary_write",
        "completed",
        "rust_xlsxwriter",
    )?;
    let append_plan_event = row_suffix_history["events"]
        .as_array()
        .context("run history events must be an array")?
        .iter()
        .find(|event| event["phase"] == "target_append_plan" && event["status"] == "completed")
        .context("target append-plan completion event missing")?;
    assert!(append_plan_event["progress"]["current_input"].is_null());
    let history_text = serde_json::to_string(&row_suffix_history)?;
    assert!(!history_text.contains(&source.display().to_string()));
    assert!(!history_text.contains("宽表甲"));

    let developer = run_menu_with_view(
        &config,
        "developer",
        &format!("1\n4\n{}\n0\n", source.display()),
    )?;
    assert!(
        developer.status.success(),
        "developer-view wide summary failed: {}",
        String::from_utf8_lossy(&developer.stderr)
    );
    let developer_events = String::from_utf8_lossy(&developer.stderr)
        .lines()
        .filter_map(|line| serde_json::from_str::<Value>(line).ok())
        .collect::<Vec<_>>();
    assert!(developer_events.iter().any(|event| {
        event["phase"] == "summary_write"
            && event["status"] == "completed"
            && event["actual_engine"] == "rust_xlsxwriter"
            && event["components"]
                .as_array()
                .is_some_and(|components| components.iter().any(|value| value == "Rust"))
    }));
    assert!(developer_events.iter().all(|event| {
        let text = event.to_string();
        !text.contains(&source.display().to_string()) && !text.contains("宽表甲")
    }));
    assert!(developer_events.windows(2).all(|pair| {
        pair[0]["total_elapsed_ms"]
            .as_u64()
            .zip(pair[1]["total_elapsed_ms"].as_u64())
            .is_some_and(|(previous, next)| previous <= next)
    }));
    let no_suffix_history = read_single_run_history(&log_dir, "PT-1-4")?;
    assert_eq!(no_suffix_history["status"], "succeeded");
    assert_history_has_phase(&no_suffix_history, "source_read", "completed", "calamine")?;

    let history_count_before_off = run_history_paths(&log_dir)?.len();
    let off = run_menu_with_view(&config, "off", &format!("1\n3\n{}\n0\n", source.display()))?;
    assert!(
        off.status.success(),
        "off-view wide summary failed: {}",
        String::from_utf8_lossy(&off.stderr)
    );
    let off_stdout = String::from_utf8_lossy(&off.stdout);
    let off_stderr = String::from_utf8_lossy(&off.stderr);
    assert!(!off_stdout.contains("run_log_path"));
    assert!(!off_stdout.contains("run_id"));
    assert!(!off_stderr.contains("正在读取第"));
    assert_eq!(run_history_paths(&log_dir)?.len(), history_count_before_off);
    Ok(())
}

#[test]
fn wide_summary_menu_failure_history_identifies_read_stage_without_source_data() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_timeline_config(root.path())?;
    let source = root.path().join("broken source.xlsx");
    fs::write(&source, b"not an xlsx workbook")?;

    let failure = run_menu_with_view(&config, "user", &format!("1\n3\n{}\n0\n", source.display()))?;
    assert!(!failure.status.success());
    let user_stderr = String::from_utf8_lossy(&failure.stderr);
    assert!(user_stderr.contains("读取失败"));
    assert!(user_stderr.contains("宽表汇总失败"));
    assert!(!user_stderr.contains(&source.display().to_string()));
    assert!(!user_stderr.contains("calamine"));

    let log_dir = root.path().join("app").join("logs");
    let history = read_single_run_history(&log_dir, "PT-1-3")?;
    assert_eq!(history["status"], "failed");
    assert_eq!(history["failure"]["stage"], "source_read");
    assert_eq!(history["failure"]["input_index"], 1);
    assert_eq!(history["failure"]["error_type"], "workbook_read_error");
    assert_history_has_phase(&history, "source_read", "failed", "calamine")?;
    let history_text = serde_json::to_string(&history)?;
    assert!(!history_text.contains(&source.display().to_string()));
    assert!(!history_text.contains("not an xlsx workbook"));
    Ok(())
}

fn assert_history_has_phase(
    history: &Value,
    phase: &str,
    status: &str,
    actual_engine: &str,
) -> Result<()> {
    let events = history["events"]
        .as_array()
        .context("run history events must be an array")?;
    assert!(
        events.iter().any(|event| {
            event["phase"] == phase
                && event["status"] == status
                && (actual_engine.is_empty() || event["actual_engine"] == actual_engine)
        }),
        "missing {phase}/{status} history event"
    );
    Ok(())
}

fn run_history_paths(log_dir: &Path) -> Result<Vec<PathBuf>> {
    let paths = fs::read_dir(log_dir)?
        .map(|entry| entry.map(|entry| entry.path()))
        .collect::<std::io::Result<Vec<_>>>()?;
    Ok(paths
        .into_iter()
        .filter(|path| {
            path.extension()
                .is_some_and(|extension| extension == "json")
        })
        .collect())
}

fn read_single_run_history(log_dir: &Path, feature_id: &str) -> Result<Value> {
    let prefix = format!("{feature_id}_");
    let paths = run_history_paths(log_dir)?
        .into_iter()
        .filter(|path| {
            path.file_name()
                .is_some_and(|name| name.to_string_lossy().starts_with(&prefix))
        })
        .collect::<Vec<_>>();
    anyhow::ensure!(
        paths.len() == 1,
        "expected one {feature_id} history file, found {}",
        paths.len()
    );
    serde_json::from_slice(&fs::read(&paths[0])?).context("parse wide summary run history")
}

fn create_timeline_config(root: &Path) -> Result<PathBuf> {
    create_timeline_config_inner(root, None, false)
}

fn create_timeline_config_with_relative_dirs(root: &Path) -> Result<PathBuf> {
    create_timeline_config_inner(root, None, true)
}

fn create_timeline_config_with_target(root: &Path, target: Option<&Path>) -> Result<PathBuf> {
    create_timeline_config_inner(root, target, false)
}

fn create_timeline_config_inner(
    root: &Path,
    target: Option<&Path>,
    relative_dirs: bool,
) -> Result<PathBuf> {
    let config = root.join("app").join("config").join("config.xlsx");
    fs::create_dir_all(config.parent().context("config parent missing")?)?;
    let app = root.join("app");
    fs::create_dir_all(app.join("input"))?;
    let directory_values = if relative_dirs {
        ["input".to_owned(), "output".to_owned(), "logs".to_owned()]
    } else {
        [
            app.join("input").to_string_lossy().into_owned(),
            app.join("output").to_string_lossy().into_owned(),
            app.join("logs").to_string_lossy().into_owned(),
        ]
    };

    let mut workbook = rust_xlsxwriter::Workbook::new();
    {
        let global = workbook.add_worksheet();
        global.set_name("全局配置")?;
        global.write_string(0, 0, "配置项")?;
        global.write_string(0, 1, "值")?;
        let settings = [
            ("输入目录", directory_values[0].clone()),
            ("输出目录", directory_values[1].clone()),
            ("日志目录", directory_values[2].clone()),
            ("错误策略", "fail_fast".to_owned()),
            ("时序规则命中策略", "all_match".to_owned()),
        ];
        for (index, (key, value)) in settings.iter().enumerate() {
            global.write_string((index + 1) as u32, 0, *key)?;
            global.write_string((index + 1) as u32, 1, value)?;
        }
    }
    {
        let timeline = workbook.add_worksheet();
        timeline.set_name("时序提取规则")?;
        let headers = [
            "是否启用",
            "规则名称",
            "工作簿关键字",
            "工作表关键字",
            "行头列",
            "列表头行",
            "必含列头",
            "必含行头",
            "数据起始行",
            "数据结束行",
            "数据起始列",
            "数据结束列",
            "跳过关键字",
            "目标工作簿路径",
            "目标工作表",
            "启用目标写入",
            "set区域",
            "目标去重列",
            "是否输出行序号",
        ];
        for (column, header) in headers.iter().enumerate() {
            timeline.write_string(0, column as u16, *header)?;
        }
        let target_workbook = target
            .map(|path| path.to_string_lossy().into_owned())
            .unwrap_or_default();
        let target_sheet = if target.is_some() { "Append" } else { "" };
        let target_enabled = if target.is_some() { "是" } else { "否" };
        let values = [
            "是",
            "WideRule",
            "sample",
            "Report",
            "A",
            "3,4",
            "",
            "",
            "5",
            "",
            "B",
            "D",
            "",
            target_workbook.as_str(),
            target_sheet,
            target_enabled,
            "集团@B2",
            "",
            "是",
        ];
        for (column, value) in values.iter().enumerate() {
            timeline.write_string(1, column as u16, *value)?;
        }
    }
    {
        let path_maps = workbook.add_worksheet();
        path_maps.set_name("路径标准化映射")?;
        for (column, header) in [
            "是否启用",
            "映射名称",
            "适用规则名",
            "工作簿关键字",
            "工作表关键字",
            "作用对象",
            "匹配方式",
            "原始路径",
            "标准路径",
            "备注",
        ]
        .iter()
        .enumerate()
        {
            path_maps.write_string(0, column as u16, *header)?;
        }
    }
    workbook
        .save(&config)
        .context("save timeline menu configuration")?;
    Ok(config)
}

fn capability_for_menu_test(plan: &ExcelEditPlan) -> BackendCapability {
    let platform = "synthetic-pt-1-wide-menu-umya".to_owned();
    BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Passed,
        format: plan.format,
        operations: plan.operations.iter().map(EditOperation::kind).collect(),
        supported_features: plan.workbook.features_present.clone(),
        preserved: plan.preserve.clone(),
        platform: platform.clone(),
        evidence: Some(CapabilityEvidence {
            sample_id: "synthetic-pt-1-wide-menu-existing-target".to_owned(),
            report_path: "pytools-core/tests/timeline_menu_cli.rs".to_owned(),
            adapter_version: "umya-spreadsheet-3.1.1".to_owned(),
            engine_version: "umya-spreadsheet-3.1.1".to_owned(),
            platform,
        }),
    }
}

fn latest_menu_json_report(output: &Output) -> Result<Value> {
    String::from_utf8_lossy(&output.stdout)
        .lines()
        .rev()
        .find_map(|line| {
            serde_json::from_str::<Value>(line)
                .ok()
                .filter(|value| value.get("target_append_plans").is_some())
        })
        .context("interactive menu did not print a wide-summary report")
}

fn latest_timeline_compare_report(output: &Output) -> Result<Value> {
    let start = output
        .stdout
        .iter()
        .position(|byte| *byte == b'{')
        .context("interactive menu did not print a timeline comparison report")?;
    serde_json::Deserializer::from_slice(&output.stdout[start..])
        .into_iter::<Value>()
        .next()
        .context("timeline comparison JSON missing")?
        .context("decode timeline comparison JSON")
}

fn create_wide_source(path: &Path, organization: &str, amount: f64) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Report")?;
    sheet.write_string(0, 0, "报表")?;
    sheet.merge_range(1, 0, 1, 1, "集团A", &rust_xlsxwriter::Format::new())?;
    sheet.write_string(2, 0, "地区")?;
    sheet.merge_range(2, 1, 2, 2, "存款", &rust_xlsxwriter::Format::new())?;
    sheet.write_string(2, 3, "贷款")?;
    sheet.write_string(3, 0, "名称")?;
    sheet.write_string(3, 1, "余额")?;
    sheet.write_string(3, 2, "余额")?;
    sheet.write_string(3, 3, "本息")?;
    sheet.write_string(4, 0, organization)?;
    sheet.write_number(4, 1, amount)?;
    sheet.write_number(4, 2, amount + 1.0)?;
    sheet.write_number(4, 3, amount + 2.0)?;
    workbook.save(path).context("save timeline source workbook")
}

fn create_wide_source_with_elapsed_duration(
    path: &Path,
    organization: &str,
    serial: f64,
) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Report")?;
    sheet.write_string(0, 0, "报表")?;
    sheet.merge_range(1, 0, 1, 1, "集团A", &rust_xlsxwriter::Format::new())?;
    sheet.write_string(2, 0, "地区")?;
    sheet.merge_range(2, 1, 2, 2, "存款", &rust_xlsxwriter::Format::new())?;
    sheet.write_string(2, 3, "贷款")?;
    sheet.write_string(3, 0, "名称")?;
    sheet.write_string(3, 1, "余额")?;
    sheet.write_string(3, 2, "余额")?;
    sheet.write_string(3, 3, "本息")?;
    sheet.write_string(4, 0, organization)?;
    let elapsed_format = rust_xlsxwriter::Format::new().set_num_format("[h]:mm:ss.000");
    sheet.write_number_with_format(4, 1, serial, &elapsed_format)?;
    sheet.write_number(4, 2, 12.0)?;
    sheet.write_number(4, 3, 23.0)?;
    workbook
        .save(path)
        .context("save elapsed-duration timeline source workbook")
}

fn create_wide_source_with_negative_date(path: &Path, organization: &str) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Report")?;
    sheet.write_string(0, 0, "报表")?;
    sheet.merge_range(1, 0, 1, 1, "集团A", &rust_xlsxwriter::Format::new())?;
    sheet.write_string(2, 0, "地区")?;
    sheet.merge_range(2, 1, 2, 2, "存款", &rust_xlsxwriter::Format::new())?;
    sheet.write_string(2, 3, "贷款")?;
    sheet.write_string(3, 0, "名称")?;
    sheet.write_string(3, 1, "余额")?;
    sheet.write_string(3, 2, "余额")?;
    sheet.write_string(3, 3, "本息")?;
    sheet.write_string(4, 0, organization)?;
    let date_format = rust_xlsxwriter::Format::new().set_num_format("yyyy-mm-dd");
    sheet.write_number_with_format(4, 1, -0.5, &date_format)?;
    sheet.write_number(4, 2, 24.5)?;
    sheet.write_number(4, 3, 25.5)?;
    workbook
        .save(path)
        .context("save negative-date timeline source workbook")
}

fn create_date_header_source(path: &Path, year: u16, month: u8, day: u8) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Report")?;
    let date = rust_xlsxwriter::ExcelDateTime::from_ymd(year, month, day)?;
    let date_format = rust_xlsxwriter::Format::new().set_num_format("yyyy-mm-dd");
    sheet.write_string(2, 0, "地区")?;
    sheet.write_string(2, 1, "存款")?;
    sheet.write_string(3, 0, "名称")?;
    sheet.write_datetime_with_format(3, 1, &date, &date_format)?;
    sheet.write_string(4, 0, "甲支行")?;
    sheet.write_number(4, 1, 1.0)?;
    workbook
        .save(path)
        .context("save date-header timeline source")
}

fn find_output(output_dir: &Path, stem_contains: &str) -> Result<PathBuf> {
    fs::read_dir(output_dir)?
        .map(|entry| entry.map(|entry| entry.path()))
        .collect::<std::io::Result<Vec<_>>>()?
        .into_iter()
        .find(|path| {
            path.extension()
                .is_some_and(|extension| extension == "xlsx")
                && path
                    .file_name()
                    .is_some_and(|name| name.to_string_lossy().contains(stem_contains))
        })
        .with_context(|| {
            format!(
                "no XLSX output containing {stem_contains:?} in {}",
                output_dir.display()
            )
        })
}

fn run_menu(config: &Path, input: &str) -> Result<Output> {
    run_menu_with_view(config, "user", input)
}

fn run_menu_with_view(config: &Path, log_view: &str, input: &str) -> Result<Output> {
    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("menu")
        .arg("--config")
        .arg(config)
        .arg("--log-view")
        .arg(log_view)
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

fn run_menu_with_capabilities(
    config: &Path,
    log_view: &str,
    capabilities: &Path,
    platform: &str,
    input: &str,
) -> Result<Output> {
    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("menu")
        .arg("--config")
        .arg(config)
        .arg("--log-view")
        .arg(log_view)
        .arg("--capabilities")
        .arg(capabilities)
        .arg("--platform")
        .arg(platform)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .context("start interactive pytools menu with edit capability")?;
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
