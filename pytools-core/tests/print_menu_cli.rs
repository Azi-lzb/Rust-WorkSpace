use anyhow::{Context, Result};
use excel_editor::{
    BackendCapability, CapabilityEvidence, CapabilityStatus, EditOperation, EngineId, ExcelEditPlan,
};
use excel_reader::read_workbook;
use pytools_core::print_ops::build_comment_print_plan;
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
const COMMENT_PRINT_PLATFORM: &str = "synthetic-pt-3-1-3-2-menu";

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
fn interactive_print_menu_routes_all_four_actions_plan_only() -> Result<()> {
    let root = TestDirectory::new()?;
    let source = root.path().join("source.xlsx");
    let target = root.path().join("printed.xlsx");
    let config = create_print_config(root.path(), &source, &target)?;
    create_print_source(&source)?;
    let source_before = fs::read(&source)?;

    for (choice, drop_comments, hide_zero) in [(1, false, false), (2, true, true)] {
        let output = run_menu(&config, &format!("3\n{choice}\n{}\n0\n", source.display()))?;
        assert!(
            output.status.success(),
            "print comment menu {choice} failed: {}",
            String::from_utf8_lossy(&output.stderr)
        );
        let response = output_json(&output)?;
        assert_eq!(response["matching_sheets"], 1);
        assert_eq!(response["marked_ranges"], 2);
        assert_eq!(response["execution"], "plan_only");
        let plan_path = PathBuf::from(
            response["plan_file"]
                .as_str()
                .context("comment print plan path missing")?,
        );
        assert!(plan_path.is_file());
        let plan: Value = serde_json::from_slice(&fs::read(plan_path)?)?;
        let imports = plan["operations"]
            .as_array()
            .context("comment print operations missing")?
            .iter()
            .filter(|operation| operation["operation"] == "import_range")
            .collect::<Vec<_>>();
        assert_eq!(imports.len(), 2);
        assert!(imports.iter().all(|operation| {
            operation["copy_comments"] == !drop_comments
                && (operation["style_mode"] == "basic") == drop_comments
                && operation["hide_zero"] == hide_zero
        }));
        assert!(!PathBuf::from(
            response["output"]
                .as_str()
                .context("comment print candidate path missing")?
        )
        .exists());
    }

    let config_plan = run_menu(&config, "3\n3\n0\n")?;
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

    let precheck = run_menu(&config, "3\n4\n0\n")?;
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
fn interactive_print_comment_menu_executes_after_capability_and_confirmation() -> Result<()> {
    let root = TestDirectory::new()?;
    for (choice, drop_comments, hide_zero) in [(1, false, false), (2, true, true)] {
        let source = root.path().join(format!("comment-print-{choice}.xlsx"));
        let target = root.path().join(format!("configured-target-{choice}.xlsx"));
        let config = create_print_config(root.path(), &source, &target)?;
        create_print_source(&source)?;
        let output_hint = root.path().join(format!("candidate-hint-{choice}.xlsx"));
        let plan = build_comment_print_plan(&source, &output_hint, drop_comments, hide_zero)?
            .plan
            .context("synthetic print fixture should produce a comment print plan")?;
        let capability_path = root.path().join(format!("print-capability-{choice}.json"));
        write_comment_print_capability(&plan, &capability_path)?;
        let source_before = fs::read(&source)?;

        let menu_input = format!("3\n{choice}\n{}\ny\n0\n", source.display());
        let result = run_menu_with_capability(&config, &capability_path, &menu_input)?;
        assert!(
            result.status.success(),
            "confirmed print mode {choice} failed: {}",
            String::from_utf8_lossy(&result.stderr)
        );
        let stdout = String::from_utf8_lossy(&result.stdout);
        assert!(stdout.contains("[确认]"));
        let response = menu_object_json(&result)?;
        assert_eq!(response["matching_sheets"], 1);
        assert_eq!(response["marked_ranges"], 2);
        assert_eq!(response["execution"]["engine"], "umya_spreadsheet");
        let output = PathBuf::from(
            response["output"]
                .as_str()
                .context("executed print output path missing")?,
        );
        assert!(stdout.contains(&output.to_string_lossy().to_string()));
        assert!(output.is_file());
        let workbook = read_workbook(&output)?;
        assert_eq!(
            workbook
                .sheets
                .iter()
                .map(|sheet| sheet.name.as_str())
                .collect::<Vec<_>>(),
            ["Marked_1", "Marked_2"]
        );
        if drop_comments {
            assert!(workbook
                .sheets
                .iter()
                .all(|sheet| sheet.comments.is_empty()));
            assert!(workbook.sheets[0]
                .values
                .get("A1")
                .is_none_or(|value| value["value"].is_null()));
        } else {
            assert!(workbook.sheets[0]
                .comments
                .get("A1")
                .is_some_and(|comment| comment.contains("打印区域 1")));
            assert_eq!(
                workbook.sheets[0]
                    .values
                    .get("A1")
                    .and_then(|value| value["value"].as_f64()),
                Some(0.0)
            );
        }
        assert_eq!(fs::read(&source)?, source_before);
    }
    Ok(())
}

#[test]
fn interactive_print_comment_menu_decline_preserves_plan_only_behavior() -> Result<()> {
    let root = TestDirectory::new()?;
    let source = root.path().join("declined-source.xlsx");
    let target = root.path().join("configured-target.xlsx");
    let config = create_print_config(root.path(), &source, &target)?;
    create_print_source(&source)?;
    let output_hint = root.path().join("declined-output-hint.xlsx");
    let plan = build_comment_print_plan(&source, &output_hint, false, false)?
        .plan
        .context("synthetic print fixture should produce a comment print plan")?;
    let capability_path = root.path().join("print-capability.json");
    write_comment_print_capability(&plan, &capability_path)?;
    let source_before = fs::read(&source)?;

    let menu_input = format!("3\n1\n{}\nn\n0\n", source.display());
    let result = run_menu_with_capability(&config, &capability_path, &menu_input)?;
    assert!(
        result.status.success(),
        "declined print plan failed: {}",
        String::from_utf8_lossy(&result.stderr)
    );
    assert!(String::from_utf8_lossy(&result.stdout).contains("[未执行]"));
    let response = menu_object_json(&result)?;
    assert_eq!(response["execution"], "plan_only");
    let output = PathBuf::from(
        response["output"]
            .as_str()
            .context("declined print output path missing")?,
    );
    assert!(!output.exists());
    let plan_path = PathBuf::from(
        response["plan_file"]
            .as_str()
            .context("declined print plan path missing")?,
    );
    assert!(plan_path.is_file());
    assert_eq!(fs::read(&source)?, source_before);
    Ok(())
}

fn write_comment_print_capability(plan: &ExcelEditPlan, path: &Path) -> Result<()> {
    let capability = BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Passed,
        format: plan.format,
        operations: plan.operations.iter().map(EditOperation::kind).collect(),
        supported_features: plan.workbook.features_present.clone(),
        preserved: plan.preserve.clone(),
        platform: COMMENT_PRINT_PLATFORM.to_owned(),
        evidence: Some(CapabilityEvidence {
            sample_id: "synthetic-pt-3-1-3-2-print-menu".to_owned(),
            report_path: "pytools-core/tests/print_menu_cli.rs".to_owned(),
            adapter_version: "umya-spreadsheet-3.1.1".to_owned(),
            engine_version: "umya-spreadsheet-3.1.1".to_owned(),
            platform: COMMENT_PRINT_PLATFORM.to_owned(),
        }),
    };
    fs::write(path, serde_json::to_vec_pretty(&capability)?)?;
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
    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("menu")
        .arg("--config")
        .arg(config)
        .arg("--log-view")
        .arg("user")
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

fn run_menu_with_capability(config: &Path, capability: &Path, input: &str) -> Result<Output> {
    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("menu")
        .arg("--config")
        .arg(config)
        .arg("--capabilities")
        .arg(capability)
        .arg("--platform")
        .arg(COMMENT_PRINT_PLATFORM)
        .arg("--log-view")
        .arg("user")
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .context("start interactive pytools menu with print capability")?;
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

fn menu_object_json(output: &Output) -> Result<Value> {
    let start = output
        .stdout
        .iter()
        .position(|byte| *byte == b'{')
        .context("menu JSON object missing")?;
    let mut deserializer = serde_json::Deserializer::from_slice(&output.stdout[start..]);
    Value::deserialize(&mut deserializer).context("decode menu JSON object")
}
