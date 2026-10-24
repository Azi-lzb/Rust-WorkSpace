use anyhow::{Context, Result};
use excel_editor::{
    BackendCapability, CapabilityEvidence, CapabilityStatus, EditOperation, EngineId,
};
use excel_reader::read_workbook;
use pytools_core::batch_ops::build_village_split_plan;
use rust_xlsxwriter::Workbook;
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
            "pytools-batch-menu-{}-{nonce}-{serial}",
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
fn interactive_batch_menu_runs_region_sum_check_through_real_cli() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let fixture_dir = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/region_sum");
    let fixture_names = [
        "balanced.xlsx",
        "mismatch.xlsx",
        "no_target.xlsx",
        "non_numeric_missing.xlsx",
        "broken.xlsx",
    ];
    let inputs = fixture_names
        .iter()
        .map(|name| fixture_dir.join(name))
        .collect::<Vec<_>>();
    let original_inputs = inputs
        .iter()
        .map(|path| Ok((path.clone(), fs::read(path)?)))
        .collect::<Result<Vec<_>>>()?;
    let unsupported_xls = root.path().join("unsupported.xls");
    fs::write(&unsupported_xls, b"synthetic legacy workbook placeholder")?;
    let original_unsupported_xls = fs::read(&unsupported_xls)?;
    let input_list = inputs
        .iter()
        .map(|path| path.to_string_lossy())
        .collect::<Vec<_>>()
        .join(";")
        + ";"
        + &unsupported_xls.to_string_lossy();

    let output = run_menu(&config, &format!("7\n5\n1\n{input_list}\n0\n"))?;
    assert!(
        output.status.success(),
        "batch menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("7  一二批处理工具"));
    assert!(stdout.contains("5  地区总分校验"));
    assert!(stdout.contains("已选择文件 6 个"));
    assert!(stdout.contains("先显式转换为 .xlsx"));
    assert!(stdout.contains(r#""files": 4"#));
    assert!(stdout.contains(r#""sheets": 3"#));
    assert!(stdout.contains(r#""errors": 1"#));
    assert!(stdout.contains(r#""failed_files": 1"#));

    let output_dir = root.path().join("app/output");
    let report_path = fs::read_dir(&output_dir)?
        .filter_map(|entry| entry.ok().map(|entry| entry.path()))
        .find(|path| {
            path.file_name()
                .and_then(|name| name.to_str())
                .is_some_and(|name| {
                    name.starts_with("地区数值核对结果_") && name.ends_with(".xlsx")
                })
        })
        .context("region sum result workbook was not created")?;
    let report = read_workbook(&report_path)?;
    let sheet = report
        .sheets
        .iter()
        .find(|sheet| sheet.name == "核对结果")
        .context("region sum result sheet missing")?;
    assert_eq!(sheet.values["B2"]["value"], "balanced.xlsx");
    assert_eq!(sheet.values["D3"]["value"], "错误");
    assert_eq!(sheet.values["B5"]["value"], "已检查文件数: 4");

    let log_dir = root.path().join("app/logs");
    assert!(fs::read_dir(log_dir)?.next().is_some());
    for (path, original) in original_inputs {
        assert_eq!(
            fs::read(&path)?,
            original,
            "input changed: {}",
            path.display()
        );
    }
    assert_eq!(fs::read(&unsupported_xls)?, original_unsupported_xls);
    Ok(())
}

#[test]
fn interactive_village_split_menu_writes_plan_without_modifying_target() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let current = root.path().join("current.xlsx");
    let previous = root.path().join("previous.xlsx");
    let target = root.path().join("target.xlsx");
    make_village_period(&current, true)?;
    make_village_period(&previous, false)?;
    make_village_target(&target)?;
    let original_target = excel_editor::open_xlsx(&target)?;
    assert!(original_target
        .defined_names()
        .iter()
        .chain(
            original_target
                .sheet_collection()
                .iter()
                .flat_map(|sheet| sheet.defined_names()),
        )
        .any(|name| name.name() == "VillageTail"));
    let current_bytes = fs::read(&current)?;
    let previous_bytes = fs::read(&previous)?;
    let target_bytes = fs::read(&target)?;

    let input = village_menu_input(&current, &previous, &target);
    let output = run_menu(&config, &input)?;
    assert!(
        output.status.success(),
        "batch menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("7-1"));
    assert!(stdout.contains("[计划模式]"));
    assert!(stdout.contains(r#""feature_id": "PT-7-1""#));
    assert!(stdout.contains(r#""planned_files": 1"#));
    assert!(stdout.contains(r#""committed_files": 0"#));
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.contains("已完成业务计算并生成工作簿修改计划"));
    assert!(!stderr.contains("\"actual_engine\""));
    assert_eq!(fs::read(&current)?, current_bytes);
    assert_eq!(fs::read(&previous)?, previous_bytes);
    assert_eq!(
        fs::read(&target)?,
        target_bytes,
        "plan-only must be read-only"
    );

    let output_dir = root.path().join("app/output");
    let plan_file = find_file_recursively(&output_dir, ".plan.json")?
        .context("village split edit plan was not written")?;
    let plan: Value = serde_json::from_slice(&fs::read(plan_file)?)?;
    let operations = plan["operations"]
        .as_array()
        .context("plan operations missing")?;
    assert!(operations
        .iter()
        .any(|operation| operation["operation"] == "delete_rows"));
    assert!(operations
        .iter()
        .any(|operation| operation["operation"] == "insert_rows"));

    let history_path = find_file_recursively(&root.path().join("app/logs"), ".run.json")?
        .context("village split run history was not written")?;
    let history_bytes = fs::read(&history_path)?;
    let history: Value = serde_json::from_slice(&history_bytes)?;
    assert_eq!(history["product"], "pytools");
    assert_eq!(history["feature_id"], "PT-7-1");
    assert_eq!(history["status"], "succeeded");
    assert!(history["events"].as_array().is_some_and(|events| {
        events
            .iter()
            .any(|event| event["phase"] == "planning" && event["actual_engine"] == "calamine")
            && events
                .iter()
                .any(|event| event["phase"] == "engine_execution" && event["status"] == "skipped")
    }));
    assert!(!String::from_utf8_lossy(&history_bytes).contains("target.xlsx"));
    Ok(())
}

#[test]
fn interactive_village_split_menu_off_view_skips_run_history() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let current = root.path().join("current.xlsx");
    let previous = root.path().join("previous.xlsx");
    let target = root.path().join("target.xlsx");
    make_village_period(&current, true)?;
    make_village_period(&previous, false)?;
    make_village_target(&target)?;
    let target_bytes = fs::read(&target)?;

    let output = run_menu_with_arguments(
        &config,
        &village_menu_input(&current, &previous, &target),
        &["--log-view".to_owned(), "off".to_owned()],
    )?;
    assert!(
        output.status.success(),
        "batch menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    assert!(!String::from_utf8_lossy(&output.stderr).contains("运行记录已保存"));
    assert_eq!(fs::read(&target)?, target_bytes);
    let log_dir = root.path().join("app/logs");
    assert!(!log_dir.exists() || fs::read_dir(log_dir)?.next().is_none());
    Ok(())
}

#[test]
fn interactive_village_split_menu_executes_capability_and_safely_commits_target() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let current = root.path().join("current.xlsx");
    let previous = root.path().join("previous.xlsx");
    let target = root.path().join("target.xlsx");
    let capability_path = root.path().join("test-capability.json");
    make_village_period(&current, true)?;
    make_village_period(&previous, false)?;
    make_village_target(&target)?;
    let original_target = excel_editor::open_xlsx(&target)?;
    let original_branch_row_height = original_target
        .sheet_by_name("广东分机构本外币")?
        .row_dimension(4)
        .map(|row| row.height());
    let original_untouched_row_height = original_target
        .sheet_by_name("非目标工作表")?
        .row_dimension(5)
        .map(|row| row.height());
    let current_bytes = fs::read(&current)?;
    let previous_bytes = fs::read(&previous)?;
    let target_bytes = fs::read(&target)?;
    let plan = build_village_split_plan(
        &current,
        &previous,
        &target,
        &root.path().join("capability-candidate.xlsx"),
    )?
    .plan
    .context("village split plan missing")?;
    let capability = village_split_capability(&plan);
    fs::write(&capability_path, serde_json::to_vec_pretty(&capability)?)?;

    let input = village_menu_input(&current, &previous, &target);
    let output = run_menu_with_capability(
        &config,
        &input,
        &capability_path,
        "synthetic-pt-7-1-umya-menu-test",
        "developer",
    )?;
    assert!(
        output.status.success(),
        "batch menu failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stdout.contains("[执行模式]"));
    assert!(stdout.contains(r#""committed_files": 1"#));
    assert!(stdout.contains("umya_spreadsheet"));
    assert!(stderr.contains("\"actual_engine\":\"umya_spreadsheet\""));
    assert!(stderr.contains("\"phase\":\"safe_commit\""));
    assert_eq!(fs::read(&current)?, current_bytes);
    assert_eq!(fs::read(&previous)?, previous_bytes);
    assert_ne!(
        fs::read(&target)?,
        target_bytes,
        "target should be committed"
    );

    let workbook = read_workbook(&target)?;
    let branch = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "广东分机构本外币")
        .context("branch sheet missing after village split")?;
    assert_eq!(branch.values["A3"]["value"], "博罗长江村镇银行");
    assert_eq!(branch.values["A4"]["value"], "惠东惠民村镇银行");
    assert_eq!(branch.values["B3"]["value"], 100.0);
    assert_eq!(branch.values["B4"]["value"], 110.0);
    let untouched = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "非目标工作表")
        .context("untouched worksheet missing")?;
    assert_eq!(untouched.values["A2"]["value"], "保持不变");

    let preserved = excel_editor::open_xlsx(&target)?;
    let branch = preserved.sheet_by_name("广东分机构本外币")?;
    assert_eq!(branch.cell("J5").map(|cell| cell.formula()), Some("A5"));
    assert!(branch
        .merge_cells()
        .iter()
        .any(|range| range.range() == "K5:L5"));
    assert_eq!(
        branch.row_dimension(5).map(|row| row.height()),
        original_branch_row_height
    );
    assert!(branch
        .comments()
        .iter()
        .any(|comment| comment.coordinate().to_string() == "M5"));
    let defined_name = preserved
        .defined_names()
        .iter()
        .chain(
            preserved
                .sheet_collection()
                .iter()
                .flat_map(|sheet| sheet.defined_names()),
        )
        .find(|name| name.name() == "VillageTail")
        .context("village tail defined name missing")?;
    assert!(defined_name.address().ends_with("$J$5"));

    let untouched = preserved.sheet_by_name("非目标工作表")?;
    assert_eq!(
        untouched.cell("B2").map(|cell| cell.formula()),
        Some("'广东分机构本外币'!A5"),
        "cross-sheet formula references must follow the target row"
    );
    assert!(untouched
        .merge_cells()
        .iter()
        .any(|range| range.range() == "D4:E4"));
    assert_eq!(
        untouched.row_dimension(5).map(|row| row.height()),
        original_untouched_row_height
    );
    assert!(untouched
        .comments()
        .iter()
        .any(|comment| comment.coordinate().to_string() == "F2"));
    Ok(())
}

#[test]
fn interactive_institution_normalize_menu_supports_user_developer_and_off_views() -> Result<()> {
    let root = TestDirectory::new()?;
    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/institution_normalize.xlsx");
    let unsupported = root.path().join("unsupported.xlsm");
    fs::write(&unsupported, b"synthetic unverified XLSM placeholder")?;

    for view in ["user", "developer", "off"] {
        let view_root = root.path().join(view);
        let config = create_config(&view_root)?;
        let target = view_root.join("institution-target.xlsx");
        fs::copy(&fixture, &target)?;
        let target_before = fs::read(&target)?;
        let input_paths = if view == "user" {
            format!("{};{}", target.display(), unsupported.display())
        } else {
            target.display().to_string()
        };
        let input = format!("7\n2\n1\n{input_paths}\n0\n");
        let output =
            run_menu_with_arguments(&config, &input, &["--log-view".to_owned(), view.to_owned()])?;
        assert!(
            output.status.success(),
            "PT-7-2 menu ({view}) failed: {}",
            String::from_utf8_lossy(&output.stderr)
        );
        assert_eq!(
            fs::read(&target)?,
            target_before,
            "{view} plan mode modified input"
        );

        let stdout = String::from_utf8_lossy(&output.stdout);
        assert!(stdout.contains("2  机构名称规范化"));
        assert!(stdout.contains(r#""feature_id": "PT-7-2""#));
        assert!(stdout.contains(r#""mapped_rows": 3"#));
        let stderr = String::from_utf8_lossy(&output.stderr);
        if view == "off" {
            assert!(stderr.trim().is_empty(), "off view emitted logs: {stderr}");
            assert!(!stdout.contains("history_report"));
            assert!(!stdout.contains("\"events\""));
        } else {
            assert!(find_file_recursively(&view_root.join("app/logs"), ".json")?.is_some());
            assert!(
                find_file_recursively(&view_root.join("app/output/plans/PT-7-2"), ".json")?
                    .is_some()
            );
            if view == "user" {
                assert!(stderr.contains("正在处理 1/1"));
                assert!(stdout.contains("[跳过]"));
                assert!(!stdout.contains("calamine"));
            } else {
                assert!(stderr.contains("\"phase\":\"plan_build\""));
                assert!(stdout.contains("\"actual_engine\": null"));
            }
        }
        if view == "off" {
            let log_dir = view_root.join("app/logs");
            assert!(!log_dir.exists() || fs::read_dir(log_dir)?.next().is_none());
            assert!(
                find_file_recursively(&view_root.join("app/output/plans/PT-7-2"), ".json")?
                    .is_some()
            );
        }
    }
    Ok(())
}

#[test]
fn interactive_institution_normalize_menu_applies_only_with_temporary_capability() -> Result<()> {
    let root = TestDirectory::new()?;
    let config = create_config(root.path())?;
    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/institution_normalize.xlsx");
    let target = root.path().join("institution-target.xlsx");
    let candidate = root.path().join("plan-candidate.xlsx");
    let capability_path = root.path().join("temporary-test-capability.json");
    fs::copy(&fixture, &target)?;
    let original = fs::read(&target)?;
    let mapping = pytools_core::config::load_institution_mapping_ordered(&config)?;
    let profile = excel_editor::inspect_xlsx_profile(&target)?;
    let plan = pytools_core::institution_normalize::build_institution_normalize_plan(
        &target,
        &candidate,
        profile,
        "synthetic-pt-7-2-menu-capability",
        &mapping,
    )?
    .plan
    .context("institution normalization plan missing")?;
    fs::write(
        &capability_path,
        serde_json::to_vec_pretty(&institution_normalize_capability(&plan))?,
    )?;

    let input = format!("7\n2\n1\n{}\n0\n", target.display());
    let output = run_menu_with_capability(
        &config,
        &input,
        &capability_path,
        "synthetic-pt-7-2-menu-test",
        "developer",
    )?;
    assert!(
        output.status.success(),
        "PT-7-2 menu apply failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stdout.contains("[执行模式]"));
    assert!(stdout.contains(r#""modified_files": 1"#));
    assert!(stderr.contains("\"actual_engine\":\"umya_spreadsheet\""));
    assert_ne!(
        fs::read(&target)?,
        original,
        "capability mode did not commit target"
    );
    let workbook = read_workbook(&target)?;
    let branch = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "报表分机构本外币")
        .context("institution branch sheet missing after menu apply")?;
    assert_eq!(branch.values["A2"]["value"], "Exact Standard");
    let baseline: Value = serde_json::from_slice(&fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("probe/reports/institution_normalize_baseline.json"),
    )?)?;
    let expected_sheets = baseline["sheets"]
        .as_array()
        .context("institution baseline sheets missing")?;
    assert_eq!(workbook.sheets.len(), expected_sheets.len());
    for expected in expected_sheets {
        let name = expected["name"]
            .as_str()
            .context("baseline sheet name missing")?;
        let actual = workbook
            .sheets
            .iter()
            .find(|sheet| sheet.name == name)
            .with_context(|| format!("output sheet {name:?} missing"))?;
        assert_eq!(serde_json::to_value(&actual.values)?, expected["values"]);
        assert_eq!(
            serde_json::to_value(&actual.formulas)?,
            expected["formulas"]
        );
        assert_eq!(serde_json::to_value(&actual.merges)?, expected["merges"]);
        assert_eq!(
            serde_json::to_value(&actual.comments)?,
            expected["comments"]
        );
    }
    Ok(())
}

fn create_config(root: &Path) -> Result<PathBuf> {
    let app = root.join("app");
    let config = app.join("config/config.xlsx");
    fs::create_dir_all(config.parent().context("config parent missing")?)?;
    fs::create_dir_all(app.join("input"))?;
    let mut workbook = rust_xlsxwriter::Workbook::new();
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
    ];
    for (index, (key, value)) in settings.iter().enumerate() {
        global.write_string((index + 1) as u32, 0, *key)?;
        global.write_string((index + 1) as u32, 1, value)?;
    }
    let mapping = workbook.add_worksheet();
    mapping.set_name("机构映射表")?;
    mapping.write_string(0, 0, "原始机构")?;
    mapping.write_string(0, 1, "标准机构")?;
    for (row, (source, target)) in [
        ("AB", "Short Standard"),
        ("ABC Bank", "Exact Standard"),
        ("branch", "Branch Standard"),
    ]
    .iter()
    .enumerate()
    {
        mapping.write_string((row + 1) as u32, 0, *source)?;
        mapping.write_string((row + 1) as u32, 1, *target)?;
    }
    workbook.save(&config)?;
    Ok(config)
}

fn make_village_period(path: &Path, current: bool) -> Result<()> {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("村镇银行本外币")?;
    for (row, base) in [(4_u32, 100.0), (5_u32, 110.0)] {
        let factor = if current { 1.0 } else { 0.5 };
        sheet.write_number(row, 2, base * factor)?;
        sheet.write_string(row, 3, "存款备注")?;
        sheet.write_number(row, 4, row as f64 + 1.0)?;
        sheet.write_number(row, 86, (base + 100.0) * factor)?;
        sheet.write_string(row, 87, "贷款备注")?;
        sheet.write_number(row, 88, row as f64 + 2.0)?;
    }
    workbook.save(path)?;
    Ok(())
}

fn make_village_target(path: &Path) -> Result<()> {
    let mut workbook = Workbook::new();
    let branch = workbook.add_worksheet();
    branch.set_name("广东分机构本外币")?;
    for (row, name) in [(0_u32, "机构"), (1, "其他"), (2, "村镇银行"), (3, "尾行")] {
        branch.write_string(row, 0, name)?;
    }
    branch.write_number(2, 1, 9.0)?;
    branch.write_formula(3, 9, "=A4")?;
    branch.merge_range(
        3,
        10,
        3,
        11,
        "保留尾行合并区域",
        &rust_xlsxwriter::Format::new(),
    )?;
    branch.insert_note(3, 12, &rust_xlsxwriter::Note::new("保留尾行批注"))?;
    branch.set_row_height(3, 24.0)?;
    let untouched = workbook.add_worksheet();
    untouched.set_name("非目标工作表")?;
    untouched.write_string(0, 0, "标题")?;
    untouched.write_string(1, 0, "保持不变")?;
    untouched.write_formula(1, 1, "='广东分机构本外币'!A4")?;
    untouched.merge_range(
        3,
        3,
        3,
        4,
        "非目标合并区域",
        &rust_xlsxwriter::Format::new(),
    )?;
    untouched.insert_note(1, 5, &rust_xlsxwriter::Note::new("非目标批注"))?;
    untouched.set_row_height(4, 19.0)?;
    workbook.define_name("VillageTail", "'广东分机构本外币'!$J$4")?;
    workbook.save(path)?;
    Ok(())
}

fn village_split_capability(plan: &excel_editor::ExcelEditPlan) -> BackendCapability {
    BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Passed,
        format: plan.format,
        operations: plan.operations.iter().map(EditOperation::kind).collect(),
        supported_features: plan.workbook.features_present.clone(),
        preserved: plan.preserve.clone(),
        platform: "synthetic-pt-7-1-umya-menu-test".to_owned(),
        evidence: Some(CapabilityEvidence {
            sample_id: "synthetic-pt-7-1-village-split".to_owned(),
            report_path: "pytools-core/tests/batch_tools_menu_cli.rs".to_owned(),
            adapter_version: "umya-spreadsheet-3.1.1".to_owned(),
            engine_version: "umya-spreadsheet-3.1.1".to_owned(),
            platform: "synthetic-pt-7-1-umya-menu-test".to_owned(),
        }),
    }
}

fn institution_normalize_capability(plan: &excel_editor::ExcelEditPlan) -> BackendCapability {
    BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Passed,
        format: plan.format,
        operations: plan.operations.iter().map(EditOperation::kind).collect(),
        supported_features: plan.workbook.features_present.clone(),
        preserved: plan.preserve.clone(),
        platform: "synthetic-pt-7-2-menu-test".to_owned(),
        evidence: Some(CapabilityEvidence {
            sample_id: "synthetic-pt-7-2-institution-normalize-menu".to_owned(),
            report_path: "pytools-core/tests/batch_tools_menu_cli.rs".to_owned(),
            adapter_version: "umya-spreadsheet-3.1.1".to_owned(),
            engine_version: "umya-spreadsheet-3.1.1".to_owned(),
            platform: "synthetic-pt-7-2-menu-test".to_owned(),
        }),
    }
}

fn village_menu_input(current: &Path, previous: &Path, target: &Path) -> String {
    format!(
        "7\n1\n{}\n{}\n1\n{}\n0\n0\n",
        current.display(),
        previous.display(),
        target.display()
    )
}

fn find_file_recursively(root: &Path, suffix: &str) -> Result<Option<PathBuf>> {
    for entry in fs::read_dir(root).with_context(|| format!("read directory {}", root.display()))? {
        let path = entry?.path();
        if path.is_dir() {
            if let Some(found) = find_file_recursively(&path, suffix)? {
                return Ok(Some(found));
            }
        } else if path
            .file_name()
            .and_then(|name| name.to_str())
            .is_some_and(|name| name.ends_with(suffix))
        {
            return Ok(Some(path));
        }
    }
    Ok(None)
}

fn run_menu(config: &Path, input: &str) -> Result<Output> {
    run_menu_with_arguments(config, input, &[])
}

fn run_menu_with_capability(
    config: &Path,
    input: &str,
    capability: &Path,
    platform: &str,
    log_view: &str,
) -> Result<Output> {
    run_menu_with_arguments(
        config,
        input,
        &[
            "--capabilities".to_owned(),
            capability.to_string_lossy().into_owned(),
            "--platform".to_owned(),
            platform.to_owned(),
            "--log-view".to_owned(),
            log_view.to_owned(),
        ],
    )
}

fn run_menu_with_arguments(config: &Path, input: &str, extra: &[String]) -> Result<Output> {
    let mut child = Command::new(env!("CARGO_BIN_EXE_pytools"))
        .arg("menu")
        .arg("--config")
        .arg(config)
        .args(extra)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()?;
    child
        .stdin
        .take()
        .context("menu stdin missing")?
        .write_all(input.as_bytes())?;
    Ok(child.wait_with_output()?)
}
