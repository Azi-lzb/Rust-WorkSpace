use anyhow::{Context, Result};
use excel_editor::{
    BackendCapability, CapabilityEvidence, CapabilityStatus, EditOperation, EngineId, ExcelEditPlan,
};
use pytools_core::dedup_ops::AppendTargetPlan;
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
            "pytools-dedup-menu-{}-{nonce}-{serial}",
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
fn interactive_dedup_menu_routes_all_six_actions_without_committing_edits() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_dedup_config(root.path())?;
    let source = root.path().join("source.xlsx");
    let target = root.path().join("target.xlsx");
    create_duplicate_source(&source)?;
    let source_before = fs::read(&source)?;

    let comment_check = run_menu(&config, &format!("2\n1\n{}\n0\n", source.display()))?;
    assert!(comment_check.status.success());
    let comment_check_output = String::from_utf8_lossy(&comment_check.stdout);
    assert!(comment_check_output.contains("开始处理：按批注检查重复并生成计划"));
    assert!(comment_check_output.contains("发现 1 条重复行，计划操作 3 项"));
    let comment_check_json = output_json(&comment_check)?;
    assert_eq!(comment_check_json["duplicate_rows"], 1);
    assert_eq!(comment_check_json["execution"], "plan_only");
    assert_eq!(
        comment_check_json["plan"]["operations"]
            .as_array()
            .unwrap()
            .len(),
        3
    );
    assert!(PathBuf::from(
        comment_check_json["plan_file"]
            .as_str()
            .context("comment check plan path missing")?
    )
    .is_file());
    assert!(!PathBuf::from(
        comment_check_json["plan"]["output"]
            .as_str()
            .context("comment check candidate path missing")?
    )
    .exists());

    let comment_delete = run_menu(&config, &format!("2\n2\n{}\n0\n", source.display()))?;
    assert!(comment_delete.status.success());
    let comment_delete_json = output_json(&comment_delete)?;
    assert_eq!(comment_delete_json["duplicate_rows"], 1);
    assert_eq!(comment_delete_json["execution"], "plan_only");
    let comment_plan = PathBuf::from(
        comment_delete_json["plan_file"]
            .as_str()
            .context("comment delete plan path missing")?,
    );
    assert!(comment_plan.is_file());
    assert!(!PathBuf::from(
        comment_delete_json["plan"]["output"]
            .as_str()
            .context("comment delete candidate path missing")?
    )
    .exists());

    let config_check = run_menu(&config, "2\n3\n0\n")?;
    assert!(config_check.status.success());
    let config_check_json = output_json(&config_check)?;
    assert_eq!(config_check_json[0]["report"]["duplicate_rows"], 1);
    assert_eq!(config_check_json[0]["execution"], "plan_only");
    assert!(PathBuf::from(
        config_check_json[0]["plan_file"]
            .as_str()
            .context("config check plan path missing")?
    )
    .is_file());
    assert_eq!(
        config_check_json[0]["report"]["plan"]["operations"]
            .as_array()
            .unwrap()
            .len(),
        3
    );
    assert!(!PathBuf::from(
        config_check_json[0]["report"]["plan"]["output"]
            .as_str()
            .context("config check candidate path missing")?
    )
    .exists());

    let config_delete = run_menu(&config, "2\n4\n0\n")?;
    assert!(config_delete.status.success());
    let config_delete_json = output_json(&config_delete)?;
    assert_eq!(config_delete_json[0]["report"]["duplicate_rows"], 1);
    assert_eq!(config_delete_json[0]["execution"], "plan_only");
    let config_delete_plan = PathBuf::from(
        config_delete_json[0]["plan_file"]
            .as_str()
            .context("config delete plan path missing")?,
    );
    assert!(config_delete_plan.is_file());
    assert!(!PathBuf::from(
        config_delete_json[0]["report"]["plan"]["output"]
            .as_str()
            .context("config delete candidate path missing")?
    )
    .exists());

    let append_plan = run_menu(&config, "2\n5\n0\n")?;
    assert!(append_plan.status.success());
    let append_json = output_json(&append_plan)?;
    assert_eq!(append_json[0]["target_existed"], false);
    assert_eq!(append_json[0]["execution"], "plan_only");
    let append_plan_file = PathBuf::from(
        append_json[0]["plan_file"]
            .as_str()
            .context("append plan path missing")?,
    );
    assert!(append_plan_file.is_file());
    assert!(!target.exists());
    assert!(!PathBuf::from(
        append_json[0]["candidate"]
            .as_str()
            .context("append candidate path missing")?
    )
    .exists());

    let precheck = run_menu(&config, "2\n6\n0\n")?;
    assert!(precheck.status.success());
    let precheck_json = output_json(&precheck)?;
    assert_eq!(precheck_json["fail"], 0);
    assert_eq!(precheck_json["ok"], 1);

    assert_eq!(
        fs::read(&source)?,
        source_before,
        "source remains unchanged"
    );
    let log_dir = root.path().join("logs");
    assert!(!log_dir.exists() || fs::read_dir(log_dir)?.next().is_none());
    Ok(())
}

#[test]
fn v2_config_check_without_capability_is_plan_only_and_preserves_source() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_dedup_config(root.path())?;
    let source = root.path().join("source.xlsx");
    create_plain_duplicate_source(&source)?;
    let source_before = fs::read(&source)?;

    let output = run_v2_menu(&config, "2\n1\n0\n", "off")?;
    assert!(
        output.status.success(),
        "V2 plan-only config duplicate check failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("[仅计划] 未提供已验收的 ExcelModifier capability/platform"));
    assert!(!stdout.contains("ExcelModifier 已调用"));

    let report = output_json_after_menu(&output)?;
    assert_eq!(report[0]["report"]["duplicate_rows"], 1);
    assert_eq!(report[0]["execution"], "plan_only");
    let candidate = PathBuf::from(
        report[0]["report"]["plan"]["output"]
            .as_str()
            .context("planned marked-copy path missing")?,
    );
    assert!(
        !candidate.exists(),
        "plan-only mode must not create a candidate"
    );
    assert_eq!(
        fs::read(&source)?,
        source_before,
        "source remains unchanged"
    );
    Ok(())
}
#[test]
fn v2_config_check_executes_whole_row_marking_to_a_safe_result_copy() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_dedup_config(root.path())?;
    let source = root.path().join("source.xlsx");
    let capability_path = root.path().join("test-capability.json");
    create_plain_duplicate_source(&source)?;
    let source_before = fs::read(&source)?;

    let seed_output = root.path().join("plan-seed");
    let seed = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["dedup-config-plan", "--config"])
        .arg(&config)
        .args(["--output-dir"])
        .arg(&seed_output)
        .args(["--action", "check"])
        .output()
        .context("build the synthetic PT-2-3 plan for its capability")?;
    assert!(
        seed.status.success(),
        "seed plan failed: {}",
        String::from_utf8_lossy(&seed.stderr)
    );
    let seed_report: Value = serde_json::from_slice(&seed.stdout)?;
    let plan: ExcelEditPlan = serde_json::from_value(seed_report[0]["report"]["plan"].clone())?;
    assert_eq!(seed_report[0]["report"]["duplicate_rows"], 1);
    assert_eq!(plan.operations.len(), 3, "the full duplicate row is marked");
    fs::write(
        &capability_path,
        serde_json::to_vec_pretty(&capability_for_feature(&plan, "pt-2-3"))?,
    )?;

    let output = run_v2_menu_with_capability(
        &config,
        "2\n1\ny\n0\n",
        &capability_path,
        "synthetic-pt-2-3-menu-test",
        "developer",
    )?;
    assert!(
        output.status.success(),
        "V2 config duplicate check failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("2-1 按配置检查重复数据"));
    assert!(stdout.contains("标记结果副本 1 个"));
    assert!(stdout.contains("ExcelModifier 已调用"));

    let report = output_json_after_menu(&output)?;
    assert_eq!(report[0]["report"]["duplicate_rows"], 1);
    assert_eq!(
        report[0]["report"]["plan"]["operations"]
            .as_array()
            .unwrap()
            .len(),
        3
    );
    assert_eq!(report[0]["execution"]["engine"], "umya_spreadsheet");
    let candidate = PathBuf::from(
        report[0]["report"]["plan"]["output"]
            .as_str()
            .context("marked candidate path missing")?,
    );
    assert!(candidate.is_file());
    assert_eq!(
        fs::read(&source)?,
        source_before,
        "source remains unchanged"
    );

    let marked_book = excel_editor::open_xlsx(&candidate)?;
    for address in ["A3", "B3", "C3"] {
        assert_eq!(
            style_fill(&marked_book, "Commented", address)?,
            Some("FFFF0000".to_owned()),
            "duplicate data row must be filled across the used row at {address}"
        );
    }
    for address in ["A2", "B2", "C2"] {
        assert_eq!(
            style_fill(&marked_book, "Commented", address)?,
            None,
            "first occurrence must remain unmarked at {address}"
        );
    }

    let log_dir = root.path().join("logs");
    let histories = fs::read_dir(&log_dir)?
        .filter_map(|entry| entry.ok().map(|entry| entry.path()))
        .filter(|path| {
            path.extension()
                .is_some_and(|extension| extension == "json")
        })
        .collect::<Vec<_>>();
    assert_eq!(histories.len(), 1);
    let history_text = fs::read_to_string(&histories[0])?;
    let history: Value = serde_json::from_str(&history_text)?;
    assert_eq!(history["feature_id"], "PT-2-3");
    assert_eq!(history["excel_modifier_invoked"], true);
    assert_eq!(history["plan_only"], false);
    assert!(history["actual_engines"]
        .as_array()
        .is_some_and(|engines| { engines.iter().any(|engine| engine == "umya_spreadsheet") }));
    assert!(!history_text.contains("source.xlsx"));
    assert!(!history_text.contains(&root.path().to_string_lossy().to_string()));
    Ok(())
}

#[test]
fn v2_config_delete_creates_a_deduplicated_copy_without_changing_source() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_dedup_config(root.path())?;
    let source = root.path().join("source.xlsx");
    let capability_path = root.path().join("test-delete-capability.json");
    create_plain_duplicate_source(&source)?;
    let source_before = fs::read(&source)?;

    let seed_output = root.path().join("delete-plan-seed");
    let seed = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["dedup-config-plan", "--config"])
        .arg(&config)
        .args(["--output-dir"])
        .arg(&seed_output)
        .args(["--action", "delete"])
        .output()
        .context("build the synthetic PT-2-4 plan for its capability")?;
    assert!(
        seed.status.success(),
        "seed delete plan failed: {}",
        String::from_utf8_lossy(&seed.stderr)
    );
    let seed_report: Value = serde_json::from_slice(&seed.stdout)?;
    let plan: ExcelEditPlan = serde_json::from_value(seed_report[0]["report"]["plan"].clone())?;
    assert_eq!(seed_report[0]["report"]["duplicate_rows"], 1);
    assert_eq!(
        plan.operations.len(),
        1,
        "one duplicate data row is deleted"
    );
    fs::write(
        &capability_path,
        serde_json::to_vec_pretty(&capability_for_feature(&plan, "pt-2-4"))?,
    )?;

    let output = run_v2_menu_with_capability(
        &config,
        "2\n2\ny\n0\n",
        &capability_path,
        "synthetic-pt-2-4-menu-test",
        "developer",
    )?;
    assert!(
        output.status.success(),
        "V2 config duplicate deletion failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("2-2 按配置删除重复数据"));
    assert!(stdout.contains("去重结果副本 1 个"));
    assert!(stdout.contains("ExcelModifier 已调用"));

    let report = output_json_after_menu(&output)?;
    assert_eq!(report[0]["report"]["duplicate_rows"], 1);
    assert_eq!(report[0]["execution"]["engine"], "umya_spreadsheet");
    let candidate = PathBuf::from(
        report[0]["report"]["plan"]["output"]
            .as_str()
            .context("deduplicated-copy path missing")?,
    );
    assert!(candidate.is_file());
    assert_eq!(
        fs::read(&source)?,
        source_before,
        "source remains unchanged"
    );

    let result = excel_reader::read_workbook(&candidate)?;
    let sheet = result
        .sheets
        .iter()
        .find(|sheet| sheet.name == "Commented")
        .context("deduplicated sheet missing")?;
    assert_eq!(sheet.values["A2"]["value"], "First");
    assert_eq!(sheet.values["A3"]["value"], "Other");
    assert!(!sheet.values.contains_key("A4"));
    assert_eq!(sheet.values["B3"]["value"], "B");
    Ok(())
}
#[test]
fn interactive_dedup_append_requires_confirmation_then_creates_a_new_target() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_dedup_config(root.path())?;
    let source = root.path().join("source.xlsx");
    let target = root.path().join("target.xlsx");
    create_duplicate_source(&source)?;
    let source_before = fs::read(&source)?;

    let output = run_menu_with_view(&config, "2\n5\nyes\n0\n", "developer")?;
    assert!(
        output.status.success(),
        "confirmed append failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report = output_json(&output)?;
    assert_eq!(report[0]["status"], "succeeded");
    assert_eq!(report[0]["execution"], "committed");
    assert_eq!(report[0]["target_existed"], false);
    assert_eq!(
        report[0]["execution_result"]["target_commit"]["operation"],
        "created"
    );
    assert!(target.is_file());
    let candidate = PathBuf::from(
        report[0]["candidate"]
            .as_str()
            .context("append candidate path missing")?,
    );
    assert!(candidate.is_file());
    assert_eq!(fs::read(&target)?, fs::read(&candidate)?);
    assert_eq!(fs::read(&source)?, source_before);

    let log_dir = root.path().join("logs");
    let histories = fs::read_dir(&log_dir)?
        .filter_map(|entry| entry.ok().map(|entry| entry.path()))
        .filter(|path| {
            path.extension()
                .is_some_and(|extension| extension == "json")
        })
        .collect::<Vec<_>>();
    assert_eq!(histories.len(), 1);
    let history_text = fs::read_to_string(&histories[0])?;
    let history: Value = serde_json::from_str(&history_text)?;
    assert_eq!(history["feature_id"], "PT-2-5");
    assert_eq!(history["status"], "succeeded");
    assert_eq!(history["excel_writer_invoked"], true);
    assert_eq!(history["excel_modifier_invoked"], false);
    assert_eq!(history["plan_only"], false);
    assert!(history["actual_engines"]
        .as_array()
        .is_some_and(|engines| engines.iter().any(|engine| engine == "rust_xlsxwriter")));
    assert!(history["stages"].as_array().is_some_and(|stages| stages
        .iter()
        .any(|stage| stage["name"] == "new_workbook_write_and_reopen")));
    assert!(!history_text.contains("source.xlsx"));
    assert!(!history_text.contains(&root.path().to_string_lossy().to_string()));
    Ok(())
}

#[test]
fn interactive_dedup_append_keeps_existing_target_plan_only_without_capability() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_dedup_config(root.path())?;
    let source = root.path().join("source.xlsx");
    let target = root.path().join("target.xlsx");
    create_duplicate_source(&source)?;
    create_initial_target(&target)?;
    let source_before = fs::read(&source)?;
    let target_before = fs::read(&target)?;

    let output = run_menu(&config, "2\n5\n0\n")?;
    assert!(
        output.status.success(),
        "plan-only append failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report = output_json(&output)?;
    assert_eq!(report[0]["target_existed"], true);
    assert_eq!(report[0]["status"], "planned_only");
    assert_eq!(
        report[0]["execution"],
        "requires ExcelModifier capability and platform"
    );
    assert!(PathBuf::from(
        report[0]["plan_file"]
            .as_str()
            .context("existing-target append plan path missing")?
    )
    .is_file());
    assert!(!PathBuf::from(
        report[0]["candidate"]
            .as_str()
            .context("append candidate path missing")?
    )
    .exists());
    assert_eq!(fs::read(&target)?, target_before);
    assert_eq!(fs::read(&source)?, source_before);
    Ok(())
}

#[test]
fn interactive_dedup_append_commits_existing_target_with_matching_capability() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_dedup_config(root.path())?;
    let source = root.path().join("source.xlsx");
    let target = root.path().join("target.xlsx");
    let capability_path = root.path().join("test-capability.json");
    create_duplicate_source(&source)?;
    create_initial_target(&target)?;
    let source_before = fs::read(&source)?;
    let target_before = fs::read(&target)?;

    let planned = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .args(["dedup-append-plan", "--config"])
        .arg(&config)
        .output()
        .context("plan existing-target append for its synthetic capability")?;
    assert!(
        planned.status.success(),
        "planning failed: {}",
        String::from_utf8_lossy(&planned.stderr)
    );
    let plan_report: Value = serde_json::from_slice(&planned.stdout)?;
    let plan_path = PathBuf::from(
        plan_report[0]["plan_file"]
            .as_str()
            .context("append plan path missing")?,
    );
    let plan: AppendTargetPlan = serde_json::from_slice(&fs::read(plan_path)?)?;
    let edit_plan = plan.edit_plan.as_ref().context("edit plan missing")?;
    fs::write(
        &capability_path,
        serde_json::to_vec_pretty(&capability_for(edit_plan))?,
    )?;

    let output = run_menu_with_capability(
        &config,
        "2\n5\nyes\n0\n",
        &capability_path,
        "synthetic-pt-2-5-menu-test",
    )?;
    assert!(
        output.status.success(),
        "confirmed existing-target append failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report = output_json(&output)?;
    assert_eq!(report[0]["target_existed"], true);
    assert_eq!(report[0]["status"], "succeeded");
    assert_eq!(report[0]["execution"], "committed");
    assert_eq!(report[0]["execution_result"]["engine"], "umya_spreadsheet");
    assert_eq!(
        report[0]["execution_result"]["target_commit"]["operation"],
        "replaced"
    );
    let candidate = PathBuf::from(
        report[0]["candidate"]
            .as_str()
            .context("append candidate path missing")?,
    );
    assert!(candidate.is_file());
    assert_eq!(fs::read(&target)?, fs::read(candidate)?);
    assert_ne!(fs::read(&target)?, target_before);
    assert_eq!(fs::read(&source)?, source_before);
    Ok(())
}

#[test]
fn interactive_dedup_menu_writes_developer_history_and_off_suppresses_it() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_dedup_config(root.path())?;
    let source = root.path().join("source.xlsx");
    create_duplicate_source(&source)?;
    let input = format!("2\n1\n{}\n0\n", source.display());

    let developer = run_menu_with_view(&config, &input, "developer")?;
    assert!(
        developer.status.success(),
        "developer menu failed: {}",
        String::from_utf8_lossy(&developer.stderr)
    );
    assert!(String::from_utf8_lossy(&developer.stdout).contains("PT-2 完成：PT-2-1"));
    let developer_report = output_json(&developer)?;
    assert_eq!(developer_report["duplicate_rows"], 1);

    let log_dir = root.path().join("logs");
    let histories = fs::read_dir(&log_dir)?
        .filter_map(|entry| entry.ok().map(|entry| entry.path()))
        .filter(|path| {
            path.extension()
                .is_some_and(|extension| extension == "json")
        })
        .collect::<Vec<_>>();
    assert_eq!(histories.len(), 1);
    let history_text = fs::read_to_string(&histories[0])?;
    let history: Value = serde_json::from_str(&history_text)?;
    assert_eq!(history["feature_id"], "PT-2-1");
    assert_eq!(history["action"], "comment_check");
    assert_eq!(history["status"], "succeeded");
    assert_eq!(history["metrics"]["duplicate_rows"], 1);
    assert_eq!(history["metrics"]["hit_sheets"], 1);
    assert_eq!(history["metrics"]["planned_operations"], 3);
    assert!(history["elapsed_ms"].as_u64().is_some());
    assert!(history["stages"]
        .as_array()
        .is_some_and(|stages| !stages.is_empty()));
    assert!(history["actual_engines"]
        .as_array()
        .is_some_and(|engines| engines
            .iter()
            .any(|engine| engine == "Calamine via excel-reader")));
    assert_eq!(history["excel_modifier_invoked"], false);
    assert_eq!(history["excel_writer_invoked"], false);
    assert!(!history_text.contains("source.xlsx"));
    assert!(!history_text.contains(&root.path().to_string_lossy().to_string()));

    let off = run_menu_with_view(&config, &input, "off")?;
    assert!(
        off.status.success(),
        "off menu failed: {}",
        String::from_utf8_lossy(&off.stderr)
    );
    let off_output = String::from_utf8_lossy(&off.stdout);
    assert!(!off_output.contains("开始处理："));
    assert!(!off_output.contains("PT-2 开始："));
    let history_count = fs::read_dir(&log_dir)?
        .filter_map(|entry| entry.ok())
        .filter(|entry| {
            entry
                .path()
                .extension()
                .is_some_and(|extension| extension == "json")
        })
        .count();
    assert_eq!(history_count, 1);
    Ok(())
}

fn create_duplicate_source(path: &Path) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Commented")?;
    sheet.write_string(0, 0, "源数据")?;
    sheet.write_string(0, 1, "ID")?;
    sheet.write_string(0, 2, "Value")?;
    sheet.insert_note(0, 0, &rust_xlsxwriter::Note::new("本表为源数据"))?;
    sheet.insert_note(0, 1, &rust_xlsxwriter::Note::new("标识列"))?;
    for (row, id, value) in [(1, "A", 10.0), (2, "A", 20.0), (3, "B", 30.0)] {
        sheet.write_string(row, 0, id)?;
        sheet.write_string(row, 1, id)?;
        sheet.write_number(row, 2, value)?;
    }
    workbook.save(path)?;
    Ok(())
}

fn create_plain_duplicate_source(path: &Path) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Commented")?;
    for (column, value) in ["Name", "ID", "Value"].iter().enumerate() {
        sheet.write_string(0, column as u16, *value)?;
    }
    for (row, name, id, value) in [
        (1, "First", "A", 10.0),
        (2, "Duplicate", "A", 20.0),
        (3, "Other", "B", 30.0),
    ] {
        sheet.write_string(row, 0, name)?;
        sheet.write_string(row, 1, id)?;
        sheet.write_number(row, 2, value)?;
    }
    workbook.save(path)?;
    Ok(())
}

fn create_initial_target(path: &Path) -> Result<()> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("Append")?;
    sheet.write_string(0, 0, "源数据")?;
    sheet.write_string(0, 1, "ID")?;
    sheet.write_string(0, 2, "Value")?;
    sheet.write_string(1, 0, "existing")?;
    sheet.write_string(1, 1, "existing")?;
    workbook.save(path)?;
    Ok(())
}

fn capability_for(plan: &ExcelEditPlan) -> BackendCapability {
    capability_for_feature(plan, "pt-2-5")
}

fn capability_for_feature(plan: &ExcelEditPlan, feature_id: &str) -> BackendCapability {
    let platform = format!("synthetic-{feature_id}-menu-test");
    BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Passed,
        format: plan.format,
        operations: plan.operations.iter().map(EditOperation::kind).collect(),
        supported_features: plan.workbook.features_present.clone(),
        preserved: plan.preserve.clone(),
        platform: platform.clone(),
        evidence: Some(CapabilityEvidence {
            sample_id: format!("synthetic-{feature_id}-dedup-menu"),
            report_path: "pytools-core/tests/dedup_menu_cli.rs".to_owned(),
            adapter_version: "umya-spreadsheet-3.1.1".to_owned(),
            engine_version: "umya-spreadsheet-3.1.1".to_owned(),
            platform: platform.clone(),
        }),
    }
}

fn create_dedup_config(root: &Path) -> Result<PathBuf> {
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
        let tasks = workbook.add_worksheet();
        tasks.set_name("去重追加数据配置")?;
        let headers = [
            "是否启用",
            "任务名",
            "源数据工作簿",
            "源数据工作表",
            "标识列序号",
            "目标工作簿",
            "目标工作表",
            "执行模式",
        ];
        for (column, header) in headers.iter().enumerate() {
            tasks.write_string(0, column as u16, *header)?;
        }
        let values = [
            "是".to_owned(),
            "DedupMenu".to_owned(),
            root.join("source.xlsx").to_string_lossy().into_owned(),
            "Commented".to_owned(),
            "2".to_owned(),
            root.join("target.xlsx").to_string_lossy().into_owned(),
            "Append".to_owned(),
            "追加".to_owned(),
        ];
        for (column, value) in values.iter().enumerate() {
            tasks.write_string(1, column as u16, value.as_str())?;
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

fn run_menu_with_capability(
    config: &Path,
    input: &str,
    capability: &Path,
    platform: &str,
) -> Result<Output> {
    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("menu")
        .arg("--config")
        .arg(config)
        .arg("--capabilities")
        .arg(capability)
        .arg("--platform")
        .arg(platform)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .context("start interactive pytools menu with capability evidence")?;
    child
        .stdin
        .take()
        .context("interactive stdin missing")?
        .write_all(input.as_bytes())
        .context("send interactive menu selections")?;
    child
        .wait_with_output()
        .context("wait for interactive pytools menu")
}

fn run_v2_menu(config: &Path, input: &str, log_view: &str) -> Result<Output> {
    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("menu-v2")
        .arg("--config")
        .arg(config)
        .arg("--log-view")
        .arg(log_view)
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
    child.wait_with_output().context("wait for V2 menu")
}
fn run_v2_menu_with_capability(
    config: &Path,
    input: &str,
    capability: &Path,
    platform: &str,
    log_view: &str,
) -> Result<Output> {
    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("menu-v2")
        .arg("--config")
        .arg(config)
        .arg("--capabilities")
        .arg(capability)
        .arg("--platform")
        .arg(platform)
        .arg("--log-view")
        .arg(log_view)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .context("start PyTools V2 menu with edit capability")?;
    child
        .stdin
        .take()
        .context("V2 menu stdin missing")?
        .write_all(input.as_bytes())
        .context("send V2 menu selections")?;
    child.wait_with_output().context("wait for V2 menu")
}

fn output_json(output: &Output) -> Result<Value> {
    let stdout = String::from_utf8_lossy(&output.stdout);
    let mut offset = 0;
    for line in stdout.split_inclusive('\n') {
        let leading_whitespace = line.len() - line.trim_start().len();
        let candidate = &line[leading_whitespace..];
        if candidate.starts_with('{') || candidate.starts_with('[') {
            let mut deserializer =
                serde_json::Deserializer::from_str(&stdout[offset + leading_whitespace..]);
            if let Ok(value) = Value::deserialize(&mut deserializer) {
                return Ok(value);
            }
        }
        offset += line.len();
    }
    anyhow::bail!("CLI JSON output missing")
}

fn output_json_after_menu(output: &Output) -> Result<Value> {
    let stdout = String::from_utf8_lossy(&output.stdout);
    let start = stdout
        .find("\n[\n")
        .map(|index| index + 1)
        .context("menu action JSON array missing")?;
    let mut deserializer = serde_json::Deserializer::from_str(&stdout[start..]);
    Value::deserialize(&mut deserializer).context("decode menu action JSON output")
}

fn style_fill(
    book: &excel_editor::Workbook,
    sheet_name: &str,
    address: &str,
) -> Result<Option<String>> {
    let sheet = book.sheet_by_name(sheet_name)?;
    Ok(sheet
        .cell(address)
        .and_then(|cell| cell.style().fill())
        .and_then(|fill| fill.pattern_fill())
        .and_then(|fill| fill.foreground_color())
        .map(|color| color.argb_str().to_ascii_uppercase()))
}
