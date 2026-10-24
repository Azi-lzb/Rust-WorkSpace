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
    let config = root.join("app").join("config").join("config.xlsx");
    fs::create_dir_all(config.parent().context("config parent missing")?)?;
    let app = root.join("app");
    fs::create_dir_all(app.join("input"))?;

    let mut workbook = rust_xlsxwriter::Workbook::new();
    {
        let global = workbook.add_worksheet();
        global.set_name("全局配置")?;
        global.write_string(0, 0, "配置项")?;
        global.write_string(0, 1, "值")?;
        let settings = [
            ("输入目录", app.join("input").to_string_lossy().into_owned()),
            (
                "输出目录",
                app.join("output").to_string_lossy().into_owned(),
            ),
            ("日志目录", app.join("logs").to_string_lossy().into_owned()),
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
            "",
            "",
            "否",
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
