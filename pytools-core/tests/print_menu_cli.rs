use anyhow::{Context, Result};
use serde::Deserialize;
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
            "pytools-print-menu-{}-{nonce}-{serial}",
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
fn interactive_print_menu_routes_static_report_and_precheck() -> Result<()> {
    let root = TestDirectory::new()?;
    let source = root.path().join("source.xlsx");
    let target = root.path().join("printed.xlsx");
    let config = create_print_config(root.path(), &source, &target)?;
    create_print_source(&source)?;
    let source_before = fs::read(&source)?;

    let menu = run_menu(&config, "3\n0\n0\n")?;
    assert!(menu.status.success());
    let menu_text = String::from_utf8_lossy(&menu.stdout);
    assert!(menu_text.contains("按配置生成静态打印报表"));
    assert!(!menu_text.contains("按批注"));
    assert!(!menu_text.contains("PDF"));

    let config_plan = run_menu(&config, "3\n1\n0\n")?;
    assert!(config_plan.status.success());
    let config_response = output_json(&config_plan)?;
    assert_eq!(config_response[0]["execution"], "plan_only");
    let config_plan_path = PathBuf::from(
        config_response[0]["plan_file"]
            .as_str()
            .context("config print plan path missing")?,
    );
    assert!(config_plan_path.is_file());
    assert!(!target.exists());
    assert!(!PathBuf::from(
        config_response[0]["report"]["output"]
            .as_str()
            .context("config print candidate path missing")?
    )
    .exists());

    let precheck = run_menu(&config, "3\n2\n0\n")?;
    assert!(precheck.status.success());
    let precheck_response = output_json(&precheck)?;
    assert_eq!(precheck_response["failed"], 0);
    assert_eq!(precheck_response["ok"], 1);

    assert_eq!(
        fs::read(&source)?,
        source_before,
        "source remains unchanged"
    );
    Ok(())
}

#[test]
fn interactive_print_menu_logs_all_actions_and_off_suppresses_progress() -> Result<()> {
    let root = TestDirectory::new()?;
    let source = root.path().join("source.xlsx");
    let target = root.path().join("printed.xlsx");
    let config = create_print_config(root.path(), &source, &target)?;
    create_print_source(&source)?;

    let cases = [
        ("PT-3-3", "config_plan", "3\n1\n0\n".to_owned()),
        ("PT-3-4", "config_precheck", "3\n2\n0\n".to_owned()),
    ];
    let log_dir = root.path().join("logs");
    for (index, (feature_id, action, input)) in cases.iter().enumerate() {
        let result = run_menu_with_view(&config, input, "developer")?;
        assert!(
            result.status.success(),
            "{feature_id} developer menu failed: {}",
            String::from_utf8_lossy(&result.stderr)
        );
        let stdout = String::from_utf8_lossy(&result.stdout);
        assert!(stdout.contains(&format!("PT-3 开始：{feature_id} {action}")));
        assert!(stdout.contains(&format!("PT-3 完成：{feature_id}")));
        let _ = output_json(&result)?;

        let histories = fs::read_dir(&log_dir)?
            .filter_map(|entry| entry.ok())
            .map(|entry| entry.path())
            .filter(|path| {
                path.file_name()
                    .and_then(|name| name.to_str())
                    .is_some_and(|name| name.starts_with(&format!("pytools-print-{action}-")))
            })
            .collect::<Vec<_>>();
        assert_eq!(histories.len(), 1, "one history per {feature_id}");
        let all_history_count = fs::read_dir(&log_dir)?.count();
        assert_eq!(all_history_count, index + 1);
        let history_text = fs::read_to_string(&histories[0])?;
        assert!(!history_text.contains(&source.to_string_lossy().to_string()));
        assert!(!history_text.contains(&config.to_string_lossy().to_string()));
        let history: Value = serde_json::from_str(&history_text)?;
        assert_eq!(history["feature_id"], *feature_id);
        assert_eq!(history["action"], *action);
        assert_eq!(history["status"], "succeeded");
        assert!(history["elapsed_ms"].as_u64().is_some());
        assert!(history["actual_engines"]
            .as_array()
            .is_some_and(|engines| engines.iter().any(|engine| engine == "calamine")));
        assert_eq!(history["excel_modifier_invoked"], false);
        assert_eq!(history["excel_writer_invoked"], false);
        match *action {
            "config_plan" => {
                assert_eq!(history["metrics"]["planned_groups"], 1);
                assert_eq!(history["metrics"]["plan_files"], 1);
                assert_eq!(history["plan_only"], true);
            }
            "config_precheck" => {
                assert_eq!(history["metrics"]["passed_tasks"], 1);
                assert_eq!(history["metrics"]["failed_tasks"], 0);
                assert_eq!(history["plan_only"], false);
            }
            _ => unreachable!(),
        }
    }

    let user_result = run_menu_with_view(&config, "3\n2\n0\n", "user")?;
    assert!(user_result.status.success());
    let user_output = String::from_utf8_lossy(&user_result.stdout);
    assert!(user_output.contains("开始处理：按配置预校验打印任务"));
    assert!(user_output.contains("处理完成：通过 1 项、警告 0 项、失败 0 项"));
    let history_count_before_off = fs::read_dir(&log_dir)?.count();

    let off_result = run_menu_with_view(&config, "3\n2\n0\n", "off")?;
    assert!(off_result.status.success());
    let off_output = String::from_utf8_lossy(&off_result.stdout);
    assert!(!off_output.contains("PT-3 开始："));
    assert!(!off_output.contains("PT-3 完成："));
    assert!(!off_output.contains("开始处理："));
    assert!(!off_output.contains("处理完成："));
    assert_eq!(fs::read_dir(&log_dir)?.count(), history_count_before_off);
    Ok(())
}

fn create_print_source(path: &Path) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Marked")?;
    sheet.write_string(0, 0, "Report")?;
    sheet.insert_note(0, 0, &rust_xlsxwriter::Note::new("打印输出 横向"))?;

    sheet.write_number_with_format(1, 1, 0.0, &rust_xlsxwriter::Format::new().set_bold())?;
    sheet.write_string(1, 2, "A")?;
    sheet.write_string(1, 3, "B")?;
    sheet.write_number(2, 1, 2.0)?;
    sheet.merge_range(2, 2, 2, 3, "merged", &rust_xlsxwriter::Format::new())?;
    sheet.write_number(3, 1, 3.0)?;
    sheet.write_string(3, 2, "E")?;
    sheet.write_string(3, 3, "F")?;
    sheet.insert_note(1, 1, &rust_xlsxwriter::Note::new("打印区域 1"))?;
    sheet.insert_note(3, 3, &rust_xlsxwriter::Note::new("打印区域#1"))?;

    sheet.write_string(1, 5, "second")?;
    sheet.write_string(2, 6, "block")?;
    sheet.insert_note(1, 5, &rust_xlsxwriter::Note::new("打印区域 2"))?;
    sheet.insert_note(2, 6, &rust_xlsxwriter::Note::new("打印区域#2"))?;
    sheet.set_row_height(1, 22.0)?;
    sheet.set_column_width(1, 14.0)?;
    workbook.add_worksheet().set_name("Plain")?;
    workbook.save(path)?;
    Ok(())
}

fn create_print_config(root: &Path, source: &Path, target: &Path) -> Result<PathBuf> {
    let config = root.join("config.xlsx");
    let mut workbook = rust_xlsxwriter::Workbook::new();
    {
        let global = workbook.add_worksheet();
        global.set_name("全局配置")?;
        global.write_string(0, 0, "配置项")?;
        global.write_string(0, 1, "值")?;
        for (row, (key, value)) in [
            ("输入目录", root.to_string_lossy().into_owned()),
            (
                "输出目录",
                root.join("output").to_string_lossy().into_owned(),
            ),
            ("日志目录", root.join("logs").to_string_lossy().into_owned()),
            ("错误策略", "fail_fast".to_owned()),
        ]
        .iter()
        .enumerate()
        {
            global.write_string((row + 1) as u32, 0, *key)?;
            global.write_string((row + 1) as u32, 1, value)?;
        }
    }
    {
        let settings = workbook.add_worksheet();
        settings.set_name("打印配置")?;
        let headers = [
            "是否启用",
            "打印模式",
            "源工作簿",
            "源工作表",
            "源工作表打印区域",
            "目标工作簿",
            "目标工作表",
            "FitToPagesWide",
            "FitToPagesTall",
            "打印方向",
            "水平居中",
            "垂直居中",
            "不输出批注",
            "零值不输出",
            "备注",
        ];
        for (column, header) in headers.iter().enumerate() {
            settings.write_string(0, column as u16, *header)?;
        }
        let values = [
            "是".to_owned(),
            "1".to_owned(),
            source.to_string_lossy().into_owned(),
            "Marked".to_owned(),
            "B2:D4".to_owned(),
            target.to_string_lossy().into_owned(),
            "Printed".to_owned(),
            "1".to_owned(),
            "1".to_owned(),
            "横向".to_owned(),
            "是".to_owned(),
            "是".to_owned(),
            "".to_owned(),
            "".to_owned(),
            "".to_owned(),
        ];
        for (column, value) in values.iter().enumerate() {
            settings.write_string(1, column as u16, value.as_str())?;
        }
    }
    workbook.save(&config)?;
    Ok(config)
}

fn run_menu(config: &Path, input: &str) -> Result<Output> {
    run_menu_with_view(config, input, "user")
}

fn run_menu_with_view(config: &Path, input: &str, log_view: &str) -> Result<Output> {
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

fn output_json(output: &Output) -> Result<Value> {
    let start = output
        .stdout
        .iter()
        .enumerate()
        .filter(|(_, byte)| matches!(byte, b'{' | b'['))
        .map(|(index, _)| index)
        .min()
        .context("CLI JSON output missing")?;
    let mut deserializer = serde_json::Deserializer::from_slice(&output.stdout[start..]);
    Value::deserialize(&mut deserializer).context("decode CLI JSON output")
}

#[test]
fn v2_print_menu_hides_removed_pdf_entry() -> Result<()> {
    let root = TestDirectory::new()?;
    let source = root.path().join("unused-source.xlsx");
    let target = root.path().join("must-not-be-created.xlsx");
    let config = create_print_config(root.path(), &source, &target)?;

    let output = run_menu(&config, "3\n0\n0\n")?;
    assert!(
        output.status.success(),
        "print menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("按配置生成静态打印报表"));
    assert!(!stdout.contains("PDF"));
    assert!(!stdout.contains("按批注"));
    assert!(
        !target.exists(),
        "PDF placeholder must not create an output"
    );
    Ok(())
}
