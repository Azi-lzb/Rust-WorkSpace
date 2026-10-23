use anyhow::{bail, Context, Result};
use chrono::Local;
use pytools_core::archive_ops::{
    archive_by_date_copy, archive_by_date_preview, archive_by_type_copy, archive_by_type_preview,
    execute_archive_plan,
};
use pytools_core::batch_ops::{
    build_quarterly_rural_loan_plan, build_remove_foreign_plan, build_rural_loan_plan,
    build_village_split_plan,
};
use pytools_core::batch_rename::run_batch_rename_33;
use pytools_core::config::{load_extract_tasks, load_global};
use pytools_core::config_admin::{build_config_repair_plan, initialize_new_config, view_config};
use pytools_core::conversion::{
    convert_excel_files_in_source_folders, convert_excel_files_to, convert_one,
    convert_word_files_in_source_folders, convert_word_files_to,
};
use pytools_core::dedup_ops::{
    build_append_by_config_plans, build_comment_duplicate_plan, build_config_duplicate_plan,
    precheck_dedup_tasks, write_new_append_candidate, AppendTargetPlan, DuplicateAction,
};
use pytools_core::edit_plan::{apply_plan_file, apply_plan_with_capability_file};
use pytools_core::file_commit::{
    copy_workbook_to_new_target, create_target_from_candidate, replace_target_from_candidate,
    sha256_file,
};
use pytools_core::file_sort_prefix::{
    run_content_prefix_rename_46_xlsx_candidate, run_content_prefix_unrename_47_xlsx_candidate,
    run_file_sort_prefix_49, run_remove_file_sort_prefix_410,
};
use pytools_core::local_filename_map::run_local_filename_map_48_candidate;
use pytools_core::print_ops::{
    build_comment_print_plan, build_config_print_plans, precheck_print_tasks,
};
use pytools_core::quick::{open_config, open_latest_output, open_output_directory};
use pytools_core::region_sum::run_region_sum_check;
use pytools_core::sheet_extract::run_extract_sheets_candidate;
use pytools_core::sheet_rename::{build_sheet_rename_plan, run_batch_sheet_rename_candidate};
use pytools_core::submission_check::run_submission_check;
use pytools_core::summary_ops::{summarize_by_comments, summarize_used_ranges};
use pytools_core::timeline_ops::{
    build_wide_target_append_plans, prepare_wide_target_replacement_source,
    run_header_path_compare, run_timeline_precheck, run_wide_summary,
    validate_wide_target_replacement_candidate,
};
use std::fs::{self, File, OpenOptions};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::time::Instant;

fn main() {
    if let Err(error) = run() {
        eprintln!("pytools: {error:#}");
        std::process::exit(2);
    }
}

fn run() -> Result<()> {
    let mut arguments = std::env::args().skip(1);
    let Some(command) = arguments.next() else {
        print_usage();
        bail!("missing command");
    };
    if command == "--help" || command == "-h" {
        print_usage();
        return Ok(());
    }
    let arguments = arguments.collect::<Vec<_>>();
    if arguments
        .iter()
        .any(|argument| matches!(argument.as_str(), "--help" | "-h"))
    {
        print_usage();
        return Ok(());
    }
    match command.as_str() {
        "timeline-precheck" => run_timeline_precheck_cli(arguments),
        "timeline-compare" => run_timeline_compare_cli(arguments),
        "wide-summary" => run_wide_summary_cli(arguments, false),
        "wide-summary-row-suffix" => run_wide_summary_cli(arguments, true),
        "summary-used-range" => run_summary_used_range_cli(arguments),
        "summary-comments" => run_summary_comments_cli(arguments),
        "archive-date-preview" => run_archive_cli(arguments, "date-preview"),
        "archive-type-preview" => run_archive_cli(arguments, "type-preview"),
        "archive-date-copy" => run_archive_cli(arguments, "date-copy"),
        "archive-type-copy" => run_archive_cli(arguments, "type-copy"),
        "archive-execute-plan" => run_archive_cli(arguments, "plan"),
        "convert-excel" => run_convert_batch_cli(arguments, true),
        "convert-word" => run_convert_batch_cli(arguments, false),
        "convert-file" => run_convert_file_cli(arguments),
        "config-init" => run_config_init_cli(arguments),
        "config-repair-plan" => run_config_repair_plan_cli(arguments),
        "config-view" => run_config_view_cli(arguments),
        "dedup-comment-plan" => run_dedup_comment_plan(arguments),
        "dedup-config-plan" => run_dedup_config_plan(arguments),
        "dedup-precheck" => run_dedup_precheck(arguments),
        "dedup-append-plan" => run_dedup_append_plan(arguments),
        "dedup-append-execute" => run_dedup_append_execute(arguments),
        "apply-edit-plan" => run_apply_edit_plan(arguments),
        "print-comment-plan" => run_print_comment_plan(arguments),
        "print-config-precheck" => run_print_config_precheck(arguments),
        "print-config-plan" => run_print_config_plan(arguments),
        "village-split-plan" => run_village_split_plan(arguments),
        "remove-foreign-bank-plan" => run_remove_foreign_bank_plan(arguments),
        "remove-foreign-bank" => run_remove_foreign_bank_batch(arguments),
        "rural-loan-plan" => run_rural_loan_plan(arguments),
        "quarterly-rural-loan-plan" => run_quarterly_rural_loan_plan(arguments),
        "institution-normalize-plan" => run_institution_normalize_plan(arguments),
        "institution-normalize" => run_institution_normalize_batch(arguments),
        "fx-header-plan" => run_fx_header_plan(arguments),
        "rename-sheets-plan" => run_rename_sheets_plan(arguments),
        "rename-sheets" => run_rename_sheets(arguments),
        "submission-check" => run_submission_check_cli(arguments),
        "open-config" => run_quick_cli(arguments, "config"),
        "open-latest-output" => run_quick_cli(arguments, "latest"),
        "open-output-dir" => run_quick_cli(arguments, "output"),
        "extract-sheets" | "extract-sheets-area" => run_extract_sheets(arguments),
        "region-sum-check" => run_region_sum(arguments),
        "file-sort-prefix" => run_file_sort_prefix(arguments),
        "remove-file-sort-prefix" => run_remove_file_sort_prefix(arguments),
        "rename-by-content" => run_content_prefix_rename_46(arguments),
        "unrename-by-content" => run_content_prefix_unrename_47(arguments),
        "rename-local-map" => run_local_filename_map(arguments),
        "rename-batch" => run_batch_rename(arguments),
        _ => {
            print_usage();
            bail!("unsupported command {command:?}")
        }
    }
}

fn run_quick_cli(arguments: Vec<String>, action: &str) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let global = load_global(config)?;
    match action {
        "config" => open_config(
            options
                .values
                .get("config")
                .map(Path::new)
                .context("--config is required")?,
        )?,
        "latest" => {
            let path = open_latest_output(&global.output_dir)?;
            println!("{}", serde_json::json!({"latest":path}));
        }
        "output" => open_output_directory(&global.output_dir)?,
        _ => unreachable!(),
    };
    Ok(())
}

fn run_config_init_cli(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let path = required_path(&options, "path")?;
    let result = initialize_new_config(&path)?;
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}
fn run_config_view_cli(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let path = required_path(&options, "config")?;
    let result = view_config(path)?;
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn run_config_repair_plan_cli(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let input = required_path(&options, "input")?;
    let output = required_path(&options, "output")?;
    let capabilities = options.values.get("capabilities");
    let platform = options.values.get("platform");
    anyhow::ensure!(
        capabilities.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    let Some(plan) = build_config_repair_plan(&input, &output)? else {
        println!(
            "{{\"status\":\"already_complete\",\"input\":{}}}",
            serde_json::to_string(&input.to_string_lossy())?
        );
        return Ok(());
    };
    if let Some(plan_path) = options.values.get("plan-output") {
        fs::write(plan_path, serde_json::to_vec_pretty(&plan)?)?;
    }
    if let (Some(capabilities), Some(platform)) = (capabilities, platform) {
        let execution = apply_plan_with_capability_file(&plan, capabilities, platform)?;
        println!(
            "{}",
            serde_json::to_string_pretty(&serde_json::json!({
                "plan": plan,
                "plan_file": options.values.get("plan-output"),
                "execution": execution,
            }))?
        );
    } else {
        println!("{}", serde_json::to_string_pretty(&plan)?);
    }
    Ok(())
}

fn run_dedup_comment_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let input = required_path(&options, "input")?;
    let output = required_path(&options, "output")?;
    let action = parse_duplicate_action(
        options
            .values
            .get("action")
            .map(String::as_str)
            .unwrap_or("check"),
    )?;
    let report = build_comment_duplicate_plan(&input, &output, action)?;
    let capabilities = options.values.get("capabilities");
    let platform = options.values.get("platform");
    anyhow::ensure!(
        capabilities.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    let mut result = serde_json::json!({
        "workbook": report.workbook,
        "hit_sheets": report.hit_sheets,
        "duplicate_rows": report.duplicate_rows,
        "skipped_tasks": report.skipped_tasks,
        "plan": report.plan,
        "execution": if capabilities.is_some() { "not_needed" } else { "plan_only" },
    });
    let planned = result.get("plan").filter(|plan| !plan.is_null()).cloned();
    if let Some(plan_value) = planned {
        if let Some(plan_path) = options.values.get("plan-output") {
            if let Some(parent) = Path::new(plan_path)
                .parent()
                .filter(|path| !path.as_os_str().is_empty())
            {
                fs::create_dir_all(parent)?;
            }
            fs::write(plan_path, serde_json::to_vec_pretty(&plan_value)?)?;
            result["plan_file"] = serde_json::json!(plan_path);
        }
        if let (Some(capabilities), Some(platform)) = (capabilities, platform) {
            let plan: excel_editor::ExcelEditPlan = serde_json::from_value(plan_value)?;
            let execution = apply_plan_with_capability_file(&plan, capabilities, platform)?;
            result["execution"] = serde_json::json!(execution);
        }
    }
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn run_dedup_config_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let output_dir = options
        .values
        .get("output-dir")
        .map(PathBuf::from)
        .unwrap_or(load_global(&config)?.output_dir);
    let action = parse_duplicate_action(
        options
            .values
            .get("action")
            .map(String::as_str)
            .unwrap_or("check"),
    )?;
    let tasks = pytools_core::config::load_dedup_tasks(&config)?;
    let mut outputs = std::collections::BTreeMap::new();
    for task in tasks.iter().filter(|task| task.enabled) {
        if let Some(input) = &task.source_wb {
            outputs
                .entry(input.clone())
                .or_insert_with(|| unique_workbook_candidate(input, &output_dir, "dedup"));
        }
    }
    fs::create_dir_all(&output_dir)?;
    let reports = build_config_duplicate_plan(&tasks, &outputs, action)?;
    let capabilities = options.values.get("capabilities");
    let platform = options.values.get("platform");
    anyhow::ensure!(
        capabilities.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    let mut results = Vec::with_capacity(reports.len());
    for (index, report) in reports.iter().enumerate() {
        let mut plan_path = None;
        let mut execution = if capabilities.is_some() {
            serde_json::json!("not_needed")
        } else {
            serde_json::json!("plan_only")
        };
        if let Some(plan) = &report.plan {
            let path = output_dir.join(format!(
                "dedup-plan-{}-{index}.json",
                chrono::Local::now().format("%Y%m%d_%H%M%S")
            ));
            fs::write(&path, serde_json::to_vec_pretty(plan)?)?;
            plan_path = Some(path.clone());
            if let (Some(capabilities), Some(platform)) = (capabilities, platform) {
                execution = serde_json::json!(apply_plan_with_capability_file(
                    plan,
                    capabilities,
                    platform
                )?);
            }
        }
        results.push(serde_json::json!({
            "report": report,
            "plan_file": plan_path,
            "execution": execution,
        }));
    }
    println!("{}", serde_json::to_string_pretty(&results)?);
    Ok(())
}

fn run_dedup_precheck(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let tasks = pytools_core::config::load_dedup_tasks(config)?;
    let report = precheck_dedup_tasks(&tasks);
    println!("{}", serde_json::to_string_pretty(&report)?);
    if report.fail > 0 {
        bail!("dedup precheck found {} failure(s)", report.fail);
    }
    Ok(())
}

fn run_dedup_append_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let global = load_global(&config)?;
    let output_dir = options
        .values
        .get("output-dir")
        .map(PathBuf::from)
        .unwrap_or(global.output_dir);
    let tasks = pytools_core::config::load_dedup_tasks(&config)?;
    let reports = build_append_by_config_plans(&tasks, &output_dir)?;
    fs::create_dir_all(&output_dir)?;
    let timestamp = Local::now().format("%Y%m%d_%H%M%S");
    let mut output = Vec::new();
    for (index, report) in reports.iter().enumerate() {
        let plan_path = output_dir.join(format!("dedup-append-{timestamp}-{index}.json"));
        fs::write(&plan_path, serde_json::to_vec_pretty(report)?)?;
        output.push(serde_json::json!({
            "target": report.target,
            "candidate": report.candidate,
            "target_existed": report.target_existed,
            "task_ok": report.task_ok,
            "task_skipped": report.task_skipped,
            "appended_rows": report.appended_rows,
            "duplicate_rows_filtered": report.duplicate_rows_filtered,
            "merged_tasks": report.merged_tasks,
            "issues": report.issues,
            "plan_file": plan_path,
            "execution": if report.target_existed { "requires apply-edit-plan capability evidence" } else { "run dedup-append-execute" },
        }));
    }
    println!("{}", serde_json::to_string_pretty(&output)?);
    Ok(())
}

fn run_dedup_append_execute(arguments: Vec<String>) -> Result<()> {
    let run_started = Instant::now();
    let (commit_target, arguments) = extract_boolean_flag(arguments, "commit-target")?;
    let options = parse_cli_arguments(arguments)?;
    let plan_path = required_path(&options, "plan")?;
    let plan: AppendTargetPlan = serde_json::from_slice(&fs::read(&plan_path)?)
        .with_context(|| format!("decode append plan {}", plan_path.display()))?;
    if plan.target_existed {
        let edit_plan = plan
            .edit_plan
            .as_ref()
            .context("append plan does not contain an edit plan")?;
        edit_plan.validate().context("stage=edit_plan_validation")?;
        let capabilities = required_path(&options, "capabilities")?;
        let platform = options
            .values
            .get("platform")
            .context("--platform is required")?;
        let expected_target_hash = if commit_target {
            let expected = plan.target_sha256.as_deref().context(
                "append plan has no target SHA-256; regenerate the plan before committing",
            )?;
            let actual = sha256_file(&plan.target)?;
            anyhow::ensure!(
                actual == expected,
                "configured target changed after planning; refusing to execute the commit: {}",
                plan.target.display()
            );
            Some(expected.to_owned())
        } else {
            None
        };
        let staging = if plan.replacement_workbook.is_some() {
            Some(TemporaryFile::new(prepare_wide_target_replacement_source(
                &plan,
            )?))
        } else {
            None
        };
        let candidate_preexisted = plan.candidate.exists();
        let edit_started = Instant::now();
        let execution = match apply_plan_with_capability_file(edit_plan, capabilities, platform) {
            Ok(execution) => execution,
            Err(error) => {
                if !candidate_preexisted && plan.candidate.exists() {
                    let _ = fs::remove_file(&plan.candidate);
                }
                return Err(error).context("stage=excel_modifier");
            }
        };
        let edit_elapsed_ms = edit_started.elapsed().as_millis();
        let mut stages = vec![serde_json::json!({
            "name": "excel_modify",
            "status": "succeeded",
            "engine": execution.engine,
            "elapsed_ms": edit_elapsed_ms
        })];
        if let Some(staging) = &staging {
            let validation_started = Instant::now();
            if let Err(error) = validate_wide_target_replacement_candidate(&plan, staging.path()) {
                if !candidate_preexisted && plan.candidate.exists() {
                    fs::remove_file(&plan.candidate).with_context(|| {
                        format!(
                            "remove invalid wide target candidate {}",
                            plan.candidate.display()
                        )
                    })?;
                }
                return Err(error).context("stage=replacement_contract_validation");
            }
            stages.push(serde_json::json!({
                "name": "replacement_contract_validation",
                "status": "succeeded",
                "elapsed_ms": validation_started.elapsed().as_millis()
            }));
        }
        let mut output = serde_json::to_value(&execution)?;
        if staging.is_some() {
            output["replacement_validation"] = serde_json::Value::String("passed".to_owned());
        }
        if let Some(expected_target_hash) = expected_target_hash {
            let commit =
                replace_target_from_candidate(&plan.candidate, &plan.target, &expected_target_hash)
                    .context("stage=target_commit")?;
            stages.push(serde_json::json!({
                "name": "target_commit",
                "status": "succeeded",
                "elapsed_ms": commit.elapsed_ms
            }));
            output["target_commit"] = serde_json::to_value(commit)?;
        } else {
            stages.push(serde_json::json!({
                "name": "target_commit",
                "status": "not_requested"
            }));
        }
        output["execution_summary"] = serde_json::json!({
            "function": "append_target_execute",
            "status": "succeeded",
            "task_success_count": plan.task_ok,
            "task_skipped_count": plan.task_skipped,
            "appended_rows": plan.appended_rows,
            "duplicate_rows_filtered": plan.duplicate_rows_filtered,
            "elapsed_ms": run_started.elapsed().as_millis(),
            "stages": stages
        });
        println!("{}", serde_json::to_string_pretty(&output)?);
    } else {
        let write_started = Instant::now();
        let candidate =
            write_new_append_candidate(&plan).context("stage=new_workbook_write_and_reopen")?;
        let mut stages = vec![serde_json::json!({
            "name": "new_workbook_write_and_reopen",
            "status": "succeeded",
            "engine": "rust_xlsxwriter",
            "elapsed_ms": write_started.elapsed().as_millis()
        })];
        let mut output = serde_json::json!({
            "candidate": candidate,
            "candidate_engine": "rust_xlsxwriter",
            "status": "created_and_reopened"
        });
        if commit_target {
            let commit = create_target_from_candidate(&candidate, &plan.target)
                .context("stage=target_commit")?;
            stages.push(serde_json::json!({
                "name": "target_commit",
                "status": "succeeded",
                "elapsed_ms": commit.elapsed_ms
            }));
            output["target_commit"] = serde_json::to_value(commit)?;
            output["status"] = serde_json::Value::String("created_candidate_and_target".into());
        } else {
            stages.push(serde_json::json!({
                "name": "target_commit",
                "status": "not_requested"
            }));
        }
        output["execution_summary"] = serde_json::json!({
            "function": "append_target_execute",
            "status": "succeeded",
            "task_success_count": plan.task_ok,
            "task_skipped_count": plan.task_skipped,
            "appended_rows": plan.appended_rows,
            "duplicate_rows_filtered": plan.duplicate_rows_filtered,
            "elapsed_ms": run_started.elapsed().as_millis(),
            "stages": stages
        });
        println!("{}", serde_json::to_string_pretty(&output)?);
    }
    Ok(())
}

struct TemporaryFile(PathBuf);

impl TemporaryFile {
    fn new(path: PathBuf) -> Self {
        Self(path)
    }

    fn path(&self) -> &Path {
        &self.0
    }
}

impl Drop for TemporaryFile {
    fn drop(&mut self) {
        let _ = fs::remove_file(&self.0);
    }
}

fn run_apply_edit_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let plan = required_path(&options, "plan")?;
    let capabilities = required_path(&options, "capabilities")?;
    let platform = options
        .values
        .get("platform")
        .context("--platform is required")?;
    let report = apply_plan_file(plan, capabilities, platform)?;
    println!("{}", serde_json::to_string_pretty(&report)?);
    Ok(())
}

fn parse_duplicate_action(value: &str) -> Result<DuplicateAction> {
    match value {
        "check" => Ok(DuplicateAction::Check),
        "mark" => Ok(DuplicateAction::Mark),
        "delete" => Ok(DuplicateAction::Delete),
        _ => bail!("--action must be check, mark, or delete"),
    }
}

fn run_print_comment_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let input = required_path(&options, "input")?;
    let output = required_path(&options, "output")?;
    let drop_comments = options
        .values
        .get("drop-comments")
        .is_some_and(|value| parse_cli_bool(value));
    let hide_zero = options
        .values
        .get("hide-zero")
        .is_some_and(|value| parse_cli_bool(value));
    let report = build_comment_print_plan(&input, &output, drop_comments, hide_zero)?;
    let capabilities = options.values.get("capabilities");
    let platform = options.values.get("platform");
    anyhow::ensure!(
        capabilities.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    let mut result = serde_json::json!({
        "input": report.input,
        "output": report.output,
        "matching_sheets": report.matching_sheets,
        "marked_ranges": report.marked_ranges,
        "plan": report.plan,
        "execution": if capabilities.is_some() { "not_needed" } else { "plan_only" },
    });
    let planned = result.get("plan").filter(|plan| !plan.is_null()).cloned();
    if let Some(plan_value) = planned {
        if let Some(plan_path) = options.values.get("plan-output") {
            if let Some(parent) = Path::new(plan_path)
                .parent()
                .filter(|path| !path.as_os_str().is_empty())
            {
                fs::create_dir_all(parent)?;
            }
            fs::write(plan_path, serde_json::to_vec_pretty(&plan_value)?)?;
            result["plan_file"] = serde_json::json!(plan_path);
        }
        if let (Some(capabilities), Some(platform)) = (capabilities, platform) {
            let plan: excel_editor::ExcelEditPlan = serde_json::from_value(plan_value)?;
            let execution = apply_plan_with_capability_file(&plan, capabilities, platform)?;
            result["execution"] = serde_json::json!(execution);
        }
    }
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn parse_cli_bool(value: &str) -> bool {
    matches!(
        value.trim().to_ascii_lowercase().as_str(),
        "1" | "true" | "yes" | "y" | "是" | "on"
    )
}

fn run_print_config_precheck(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let tasks = pytools_core::config::load_print_tasks(config)?;
    let report = precheck_print_tasks(&tasks);
    println!("{}", serde_json::to_string_pretty(&report)?);
    if report.failed > 0 {
        bail!("print precheck found {} failure(s)", report.failed);
    }
    Ok(())
}

fn run_print_config_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let output_dir = match options.values.get("output-dir") {
        Some(path) => PathBuf::from(path),
        None => load_global(&config)?.output_dir,
    };
    let tasks = pytools_core::config::load_print_tasks(config)?;
    let reports = build_config_print_plans(&tasks)?;
    if !reports.is_empty() {
        fs::create_dir_all(&output_dir)?;
    }
    let capabilities = options.values.get("capabilities");
    let platform = options.values.get("platform");
    anyhow::ensure!(
        capabilities.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    let mut output = Vec::new();
    for (index, report) in reports.iter().enumerate() {
        let plan = report
            .plan
            .as_ref()
            .context("grouped print report has no edit plan")?;
        let path = output_dir.join(format!("print-target-{index}.plan.json"));
        fs::write(&path, serde_json::to_vec_pretty(plan)?)?;
        let execution = if let (Some(capabilities), Some(platform)) = (capabilities, platform) {
            let output_parent = plan
                .output
                .parent()
                .filter(|parent| !parent.as_os_str().is_empty())
                .unwrap_or_else(|| std::path::Path::new("."));
            fs::create_dir_all(output_parent).with_context(|| {
                format!("create print target directory {}", output_parent.display())
            })?;
            serde_json::json!(apply_plan_with_capability_file(
                plan,
                capabilities,
                platform
            )?)
        } else {
            serde_json::json!("plan_only")
        };
        output.push(serde_json::json!({
            "report": report,
            "plan_file": path,
            "execution": execution,
        }));
    }
    println!("{}", serde_json::to_string_pretty(&output)?);
    Ok(())
}

fn run_village_split_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let current = required_path(&options, "current")?;
    let previous = required_path(&options, "previous")?;
    let input = required_path(&options, "input")?;
    let output = required_path(&options, "output")?;
    let report = build_village_split_plan(&current, &previous, &input, &output)?;
    write_optional_plan(&options, &report.plan)?;
    let execution = execute_optional_plan(&options, report.plan.as_ref())?;
    let mut result = serde_json::to_value(report)?;
    result["execution"] = execution;
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn run_remove_foreign_bank_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let input = required_path(&options, "input")?;
    let output = required_path(&options, "output")?;
    let foreign = pytools_core::config::load_foreign_banks(config)?;
    let report = build_remove_foreign_plan(&input, &output, &foreign)?;
    write_optional_plan(&options, &report.plan)?;
    let execution = execute_optional_plan(&options, report.plan.as_ref())?;
    let mut result = serde_json::to_value(report)?;
    result["execution"] = execution;
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn run_remove_foreign_bank_batch(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    anyhow::ensure!(
        !options.positional.is_empty(),
        "provide at least one source workbook"
    );
    let capabilities = options.values.get("capabilities");
    let platform = options.values.get("platform");
    anyhow::ensure!(
        capabilities.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    let execute = capabilities.is_some();
    let foreign = pytools_core::config::load_foreign_banks(&config)?;
    let mut result = serde_json::json!({
        "files": 0,
        "ok_files": 0,
        "branch_sheets": 0,
        "deleted_rows": 0,
        "failed_files": 0,
        "planned_files": 0,
        "planned_branch_sheets": 0,
        "planned_deleted_rows": 0,
        "foreign_empty": foreign.is_empty(),
        "mode": if execute { "execute" } else { "plan_only" },
        "items": [],
    });
    if foreign.is_empty() {
        println!("{}", serde_json::to_string_pretty(&result)?);
        return Ok(());
    }

    result["files"] = serde_json::json!(options.positional.len());
    for input in options.positional {
        let started = Instant::now();
        let output = match foreign_bank_copy_path(&input) {
            Ok(path) => path,
            Err(error) => {
                result["failed_files"] =
                    serde_json::json!(result["failed_files"].as_u64().unwrap_or_default() + 1);
                result["items"]
                    .as_array_mut()
                    .unwrap()
                    .push(serde_json::json!({
                        "input": input,
                        "status": "failed",
                        "error_stage": "input_validation",
                        "error": format!("{error:#}"),
                        "elapsed_ms": started.elapsed().as_millis(),
                    }));
                continue;
            }
        };
        let extension_is_xlsx = input
            .extension()
            .and_then(|value| value.to_str())
            .is_some_and(|value| value.eq_ignore_ascii_case("xlsx"));
        if !extension_is_xlsx {
            result["failed_files"] =
                serde_json::json!(result["failed_files"].as_u64().unwrap_or_default() + 1);
            result["items"]
                .as_array_mut()
                .unwrap()
                .push(serde_json::json!({
                    "input": input,
                    "output": output,
                    "status": "failed",
                    "error_stage": "input_validation",
                    "error": "Rust row-removal candidate currently accepts .xlsx only",
                    "elapsed_ms": started.elapsed().as_millis(),
                }));
            continue;
        }
        let report = match build_remove_foreign_plan(&input, &output, &foreign) {
            Ok(report) => report,
            Err(error) => {
                result["failed_files"] =
                    serde_json::json!(result["failed_files"].as_u64().unwrap_or_default() + 1);
                result["items"]
                    .as_array_mut()
                    .unwrap()
                    .push(serde_json::json!({
                        "input": input,
                        "output": output,
                        "status": "failed",
                        "error_stage": "planning",
                        "error": format!("{error:#}"),
                        "elapsed_ms": started.elapsed().as_millis(),
                    }));
                continue;
            }
        };
        let planned_deleted_rows = report
            .plan
            .as_ref()
            .map(|plan| {
                plan.operations
                    .iter()
                    .map(|operation| match operation {
                        excel_editor::EditOperation::DeleteRows { count, .. } => *count as usize,
                        _ => 0,
                    })
                    .sum::<usize>()
            })
            .unwrap_or_default();
        result["planned_files"] =
            serde_json::json!(result["planned_files"].as_u64().unwrap_or_default() + 1);
        result["planned_branch_sheets"] = serde_json::json!(
            result["planned_branch_sheets"].as_u64().unwrap_or_default()
                + report.matched_sheets as u64
        );
        result["planned_deleted_rows"] = serde_json::json!(
            result["planned_deleted_rows"].as_u64().unwrap_or_default()
                + planned_deleted_rows as u64
        );

        let mut item = serde_json::json!({
            "input": input,
            "output": output,
            "status": if execute { "pending" } else if report.plan.is_some() { "plan_only" } else { "plan_only_no_changes" },
            "planned_branch_sheets": report.matched_sheets,
            "planned_deleted_rows": planned_deleted_rows,
            "branch_sheets": 0,
            "deleted_rows": 0,
            "plan": report.plan,
        });
        if !execute {
            item["execution"] = serde_json::json!("plan_only");
        } else {
            let execution: Result<serde_json::Value> = if let Some(plan) =
                item.get("plan").filter(|value| !value.is_null()).cloned()
            {
                let plan: excel_editor::ExcelEditPlan = serde_json::from_value(plan)?;
                apply_plan_with_capability_file(
                    &plan,
                    capabilities.expect("validated capability path"),
                    platform.expect("validated platform"),
                )
                .and_then(|report| serde_json::to_value(report).context("serialize edit result"))
            } else {
                copy_workbook_to_new_target(&input, &output).and_then(|report| {
                    serde_json::to_value(report).context("serialize copy result")
                })
            };
            match execution {
                Ok(execution) => {
                    result["ok_files"] =
                        serde_json::json!(result["ok_files"].as_u64().unwrap_or_default() + 1);
                    result["branch_sheets"] = serde_json::json!(
                        result["branch_sheets"].as_u64().unwrap_or_default()
                            + report.matched_sheets as u64
                    );
                    result["deleted_rows"] = serde_json::json!(
                        result["deleted_rows"].as_u64().unwrap_or_default()
                            + planned_deleted_rows as u64
                    );
                    item["status"] = serde_json::json!("written");
                    item["branch_sheets"] = serde_json::json!(report.matched_sheets);
                    item["deleted_rows"] = serde_json::json!(planned_deleted_rows);
                    item["execution"] = execution;
                }
                Err(error) => {
                    result["failed_files"] =
                        serde_json::json!(result["failed_files"].as_u64().unwrap_or_default() + 1);
                    item["status"] = serde_json::json!("failed");
                    item["error_stage"] = serde_json::json!("execution");
                    item["error"] = serde_json::json!(format!("{error:#}"));
                }
            }
        }
        item["elapsed_ms"] = serde_json::json!(started.elapsed().as_millis());
        result["items"].as_array_mut().unwrap().push(item);
    }
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn foreign_bank_copy_path(input: &Path) -> Result<PathBuf> {
    let stem = input
        .file_stem()
        .with_context(|| format!("source workbook has no file stem: {}", input.display()))?;
    let extension = input
        .extension()
        .with_context(|| format!("source workbook has no extension: {}", input.display()))?;
    let mut file_name = stem.to_os_string();
    file_name.push("(金融局).");
    file_name.push(extension);
    Ok(input.with_file_name(file_name))
}

fn run_rural_loan_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let current = required_path(&options, "current")?;
    let previous = required_path(&options, "previous")?;
    let output = required_path(&options, "output")?;
    let (report, summary) = build_rural_loan_plan(&current, &previous, &output)?;
    write_optional_plan(&options, &report.plan)?;
    let execution = execute_optional_plan(&options, report.plan.as_ref())?;
    println!(
        "{}",
        serde_json::to_string_pretty(
            &serde_json::json!({"edit":report,"summary":summary,"execution":execution})
        )?
    );
    Ok(())
}

fn run_quarterly_rural_loan_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let current = required_path(&options, "current")?;
    let previous = required_path(&options, "previous")?;
    let output = required_path(&options, "output")?;
    let (report, summary) = build_quarterly_rural_loan_plan(&current, &previous, &output)?;
    write_optional_plan(&options, &report.plan)?;
    let execution = execute_optional_plan(&options, report.plan.as_ref())?;
    println!(
        "{}",
        serde_json::to_string_pretty(
            &serde_json::json!({"edit":report,"summary":summary,"execution":execution})
        )?
    );
    Ok(())
}

fn run_institution_normalize_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let input = required_path(&options, "input")?;
    let output = required_path(&options, "output")?;
    let profile = excel_editor::inspect_xlsx_profile(&input)?;
    let mapping = pytools_core::config::load_institution_mapping_ordered(config)?;
    let report = pytools_core::institution_normalize::build_institution_normalize_plan(
        &input,
        &output,
        profile,
        "pytools-institution-normalize",
        &mapping,
    )?;
    write_optional_plan(&options, &report.plan)?;
    let execution = execute_optional_plan(&options, report.plan.as_ref())?;
    println!(
        "{}",
        serde_json::to_string_pretty(&serde_json::json!({
            "branch_sheets": report.branch_sheets,
            "mapped_sheets": report.mapped_sheets,
            "mapped_rows": report.mapped_rows,
            "mapping_empty": report.mapping_empty,
            "plan_written": report.plan.is_some(),
            "execution": execution,
        }))?
    );
    Ok(())
}

fn run_institution_normalize_batch(arguments: Vec<String>) -> Result<()> {
    let (apply, arguments) = extract_boolean_flag(arguments, "apply")?;
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let view = options
        .values
        .get("view")
        .map(String::as_str)
        .unwrap_or("user");
    anyhow::ensure!(
        matches!(view, "user" | "developer"),
        "--view must be either user or developer"
    );
    let plan_dir = options.values.get("plan-dir").map(PathBuf::from);
    let report_output = options
        .values
        .get("report-output")
        .map(PathBuf::from)
        .unwrap_or_else(|| {
            let nonce = std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .map(|duration| duration.as_nanos())
                .unwrap_or_default();
            PathBuf::from("logs").join(format!(
                "PT-7-2_{}_{}_{}.json",
                chrono::Local::now().format("%Y%m%dT%H%M%S"),
                std::process::id(),
                nonce
            ))
        });
    anyhow::ensure!(
        report_output
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("json")),
        "--report-output must use a .json extension"
    );
    anyhow::ensure!(
        !report_output.exists(),
        "run history file already exists; refusing to overwrite {}",
        report_output.display()
    );
    if let Some(parent) = report_output
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
    {
        fs::create_dir_all(parent)
            .with_context(|| format!("create history directory {}", parent.display()))?;
    }
    let capability = options.values.get("capabilities").map(PathBuf::from);
    let platform = options.values.get("platform").cloned();
    anyhow::ensure!(
        apply == capability.is_some() && apply == platform.is_some(),
        "--apply requires both --capabilities and --platform; omit all three for plan-only mode"
    );
    let run_options = pytools_core::institution_normalize::InstitutionNormalizeRunOptions {
        apply,
        capabilities: capability,
        platform,
        plan_dir,
    };
    let mut emitted_events = Vec::new();
    let report = pytools_core::institution_normalize::run_institution_normalize_files(
        &config,
        &options.positional,
        run_options,
        |event| {
            if view == "developer" {
                if let Ok(encoded) = serde_json::to_string(event) {
                    eprintln!("{encoded}");
                }
            } else {
                let current = event
                    .progress
                    .current_file
                    .unwrap_or(event.progress.processed_files);
                eprintln!(
                    "{}（进度 {current}/{}, 成功 {}，失败 {}）",
                    event.user_message,
                    event.progress.total_files,
                    event.progress.completed_files,
                    event.progress.failed_files
                );
            }
            emitted_events.push(event.clone());
        },
    )?;
    anyhow::ensure!(
        emitted_events.len() == report.events.len(),
        "internal progress event count mismatch"
    );

    write_json_report_create_new(&report_output, &report)?;
    let output = if view == "developer" {
        serde_json::to_value(&report)?
    } else {
        institution_normalize_user_view(&report)
    };
    let mut output = output;
    output["history_report"] = serde_json::Value::String(report_output.display().to_string());
    println!("{}", serde_json::to_string_pretty(&output)?);
    Ok(())
}

fn institution_normalize_user_view(
    report: &pytools_core::institution_normalize::InstitutionNormalizeRunReport,
) -> serde_json::Value {
    serde_json::json!({
        "schema_version": report.schema_version,
        "product": report.product,
        "feature_id": report.feature_id,
        "run_id": report.run_id,
        "started_at": report.started_at,
        "status": user_friendly_normalize_run_status(&report.status),
        "mode": if report.mode == "apply" { "执行修改模式" } else { "仅生成计划，未修改文件" },
        "failure": report.failure.as_ref().map(|failure| serde_json::json!({
            "stage": user_friendly_normalize_error_stage(&failure.stage),
            "message": failure.user_message
        })),
        "progress": report.events.last().map(|event| &event.progress),
        "summary": {
            "total_inputs": report.total_inputs,
            "completed_files": report.completed_files,
            "failed_files": report.failed_files,
            "modified_files": report.modified_files,
            "branch_sheets": report.branch_sheets,
            "mapped_sheets": report.mapped_sheets,
            "mapped_rows": report.mapped_rows,
            "mapping_empty": report.mapping_empty,
            "configuration_elapsed_ms": report.configuration_elapsed_ms,
            "elapsed_ms": report.elapsed_ms
        },
        "events": report.events.iter().map(|event| serde_json::json!({
            "sequence": event.sequence,
            "timestamp": event.timestamp,
            "phase": user_friendly_normalize_phase(&event.phase),
            "status": user_friendly_normalize_event_status(&event.status),
            "message": event.user_message,
            "progress": event.progress,
            "file_name": event.file_name,
            "elapsed_ms": event.elapsed_ms
        })).collect::<Vec<_>>(),
        "files": report.files.iter().map(|file| serde_json::json!({
            "file_name": file.file_name,
            "status": user_friendly_normalize_status(&file.status),
            "branch_sheets": file.branch_sheets,
            "mapped_sheets": file.mapped_sheets,
            "mapped_rows": file.mapped_rows,
            "operation_count": file.operation_count,
            "plan_saved": file.plan_path.is_some(),
            "stages": [
                {"name": "文件检查", "elapsed_ms": file.timings.source_fingerprint_ms},
                {"name": "工作簿结构检查", "elapsed_ms": file.timings.profile_inspection_ms},
                {"name": "业务规则与修改计划", "elapsed_ms": file.timings.plan_build_ms},
                {"name": "计划保存", "elapsed_ms": file.timings.plan_persistence_ms},
                {"name": "输入一致性复核", "elapsed_ms": file.timings.source_stability_check_ms},
                {"name": "修改引擎执行", "elapsed_ms": file.timings.engine_execution_ms},
                {"name": "安全提交", "elapsed_ms": file.timings.safe_commit_ms},
                {"name": "文件总耗时", "elapsed_ms": file.timings.total_ms}
            ],
            "error_stage": file.error_stage.as_deref().map(user_friendly_normalize_error_stage),
            "message": file.user_message,
            "cleanup_warning": file.cleanup_warning
        })).collect::<Vec<_>>()
    })
}

fn user_friendly_normalize_phase(phase: &str) -> &str {
    match phase {
        "source_fingerprint" => "文件安全检查",
        "profile_inspection" => "工作簿结构检查",
        "plan_build" => "规则检查与修改计划",
        "plan_persistence" => "保存修改计划",
        "source_stability_check" => "输入一致性复核",
        "engine_execution" => "应用修改",
        "safe_commit" => "验证并提交结果",
        "configuration_read" => "读取映射配置",
        "normalization" => "机构名称规范化",
        "file_processing" => "处理工作簿",
        "task" => "任务汇总",
        _ => "处理步骤",
    }
}

fn user_friendly_normalize_run_status(status: &str) -> &str {
    match status {
        "completed" => "已完成",
        "partial_failure" => "部分文件失败",
        "failed" => "失败",
        _ => "处理中",
    }
}

fn user_friendly_normalize_event_status(status: &str) -> &str {
    match status {
        "started" => "进行中",
        "completed" => "已完成",
        "partial_failure" => "部分文件失败",
        "failed" => "失败",
        "skipped" => "已跳过",
        _ => "处理中",
    }
}

fn user_friendly_normalize_status(status: &str) -> &str {
    match status {
        "planned" => "已生成计划，尚未修改",
        "applied" => "已修改",
        "no_changes" => "无需修改",
        "failed" => "失败",
        other => other,
    }
}

fn user_friendly_normalize_error_stage(stage: &str) -> &str {
    match stage {
        "configuration_read" => "读取映射配置",
        "format_check" => "文件格式检查",
        "input_check" => "输入文件检查",
        "source_fingerprint" | "source_stability_check" => "源文件安全检查",
        "profile_inspection" => "工作簿结构检查",
        "plan_build" => "规则计算与修改计划",
        "plan_persistence" => "修改计划保存",
        "engine_execution" => "工作簿修改",
        "safe_commit" => "安全提交",
        other => other,
    }
}

fn write_json_report_create_new<T: serde::Serialize>(path: &Path, report: &T) -> Result<()> {
    if let Some(parent) = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
    {
        fs::create_dir_all(parent)
            .with_context(|| format!("create report directory {}", parent.display()))?;
    }
    let mut output = OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(path)
        .with_context(|| format!("create run report {} without overwriting", path.display()))?;
    output
        .write_all(&serde_json::to_vec_pretty(report)?)
        .with_context(|| format!("write run report {}", path.display()))?;
    output
        .sync_all()
        .with_context(|| format!("sync run report {}", path.display()))?;
    Ok(())
}

fn run_fx_header_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let input = required_path(&options, "input")?;
    let output = required_path(&options, "output")?;
    let profile = excel_editor::inspect_xlsx_profile(&input)?;
    let report = pytools_core::fx_header::build_fx_header_plan(
        &input,
        &output,
        profile,
        "pytools-fx-header",
    )?;
    write_optional_plan(&options, &report.plan)?;
    let execution = execute_optional_plan(&options, report.plan.as_ref())?;
    println!(
        "{}",
        serde_json::to_string_pretty(&serde_json::json!({
            "fx_sheets": report.fx_sheets,
            "modified_sheets": report.modified_sheets,
            "plan_written": report.plan.is_some(),
            "execution": execution,
        }))?
    );
    Ok(())
}

fn write_optional_plan(
    options: &ParsedCli,
    plan: &Option<excel_editor::ExcelEditPlan>,
) -> Result<()> {
    if let (Some(path), Some(plan)) = (options.values.get("plan-output"), plan) {
        if let Some(parent) = Path::new(path).parent() {
            if !parent.as_os_str().is_empty() {
                fs::create_dir_all(parent)?;
            }
        }
        fs::write(path, serde_json::to_vec_pretty(plan)?)?;
    }
    Ok(())
}

fn execute_optional_plan(
    options: &ParsedCli,
    plan: Option<&excel_editor::ExcelEditPlan>,
) -> Result<serde_json::Value> {
    let capabilities = options.values.get("capabilities");
    let platform = options.values.get("platform");
    anyhow::ensure!(
        capabilities.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    let Some(plan) = plan else {
        return Ok(serde_json::json!("no_changes"));
    };
    match (capabilities, platform) {
        (Some(capabilities), Some(platform)) => Ok(serde_json::to_value(
            apply_plan_with_capability_file(plan, capabilities, platform)?,
        )?),
        _ => Ok(serde_json::json!("plan_only")),
    }
}

fn run_rename_sheets_plan(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let input = required_path(&options, "input")?;
    let output = required_path(&options, "output")?;
    let report = build_sheet_rename_plan(&config, &input, &output)?;
    write_optional_plan(&options, &report.plan)?;
    let execution = execute_optional_plan(&options, report.plan.as_ref())?;
    let mut result = serde_json::to_value(report)?;
    result["execution"] = execution;
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn run_rename_sheets(arguments: Vec<String>) -> Result<()> {
    let (recursive, arguments) = extract_boolean_flag(arguments, "recursive")?;
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let capabilities = required_path(&options, "capabilities")?;
    let platform = options
        .values
        .get("platform")
        .context("--platform is required when executing sheet renames")?;
    let report = run_batch_sheet_rename_candidate(
        &config,
        &options.positional,
        recursive,
        &capabilities,
        platform,
    )?;
    println!("{}", serde_json::to_string_pretty(&report)?);
    Ok(())
}

fn run_submission_check_cli(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let output_dir = options
        .values
        .get("output-dir")
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("output"));
    if options.positional.is_empty() {
        bail!("provide at least one file or directory path");
    }
    let report = run_submission_check(&config, &options.positional, &output_dir)?;
    println!("{}", serde_json::to_string_pretty(&report)?);
    Ok(())
}

fn run_convert_batch_cli(arguments: Vec<String>, excel: bool) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    if options.positional.is_empty() {
        bail!("provide at least one source file path");
    }
    let engine = options
        .values
        .get("engine")
        .map(String::as_str)
        .unwrap_or("auto");
    let target = options.values.get("to").cloned().unwrap_or_else(|| {
        pytools_core::config::load_global_value(
            &config,
            if excel {
                "excel目的格式"
            } else {
                "word目的格式"
            },
            if excel { "xlsx" } else { "docx" },
        )
        .unwrap_or_else(|_| if excel { "xlsx".into() } else { "docx".into() })
    });
    let result = match (excel, options.values.get("output-dir")) {
        (true, Some(output)) => {
            convert_excel_files_to(&options.positional, Path::new(output), &target, engine)?
        }
        (false, Some(output)) => {
            convert_word_files_to(&options.positional, Path::new(output), &target, engine)?
        }
        (true, None) => {
            convert_excel_files_in_source_folders(&options.positional, &target, engine)?
        }
        (false, None) => {
            convert_word_files_in_source_folders(&options.positional, &target, engine)?
        }
    };
    println!("{}", serde_json::to_string_pretty(&result)?);
    if result.failed > 0 {
        bail!("{} conversion(s) failed", result.failed);
    }
    Ok(())
}

fn run_convert_file_cli(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let input = required_path(&options, "input")?;
    let target = options.values.get("to").context("--to is required")?;
    let output = if let Some(value) = options.values.get("output") {
        PathBuf::from(value)
    } else {
        input.with_extension(target)
    };
    let engine = options
        .values
        .get("engine")
        .map(String::as_str)
        .unwrap_or("auto");
    let validation = convert_one(&input, &output, target, engine)?;
    println!("{validation}");
    Ok(())
}

fn run_timeline_precheck_cli(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let report = run_timeline_precheck(config)?;
    println!("{}", serde_json::to_string_pretty(&report)?);
    if report.errors > 0 {
        bail!("timeline precheck found {} error(s)", report.errors);
    }
    Ok(())
}

fn run_timeline_compare_cli(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let template = required_path(&options, "template")?;
    let global = load_global(&config)?;
    let rules = pytools_core::config::load_timeline_rules(&config)?;
    let output = run_header_path_compare(&template, &options.positional, &rules, &global)?;
    println!("{}", serde_json::json!({"output": output}));
    Ok(())
}

fn run_wide_summary_cli(arguments: Vec<String>, row_suffix: bool) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let config = required_path(&options, "config")?;
    let global = load_global(&config)?;
    let rules = pytools_core::config::load_timeline_rules(&config)?;
    let maps = pytools_core::config::load_path_maps(&config)?;
    let output = run_wide_summary(&options.positional, &rules, &maps, &global, row_suffix)?;
    let mut append_plans = Vec::new();
    if let Some(summary_path) = output.as_deref() {
        for (index, plan) in
            build_wide_target_append_plans(summary_path, &rules, &global.output_dir)?
                .into_iter()
                .enumerate()
        {
            let stem = summary_path
                .file_stem()
                .unwrap_or_default()
                .to_string_lossy();
            let plan_path = global
                .output_dir
                .join(format!("{stem}_target_append_{}.json", index + 1));
            fs::write(&plan_path, serde_json::to_vec_pretty(&plan)?)?;
            append_plans.push(plan_path);
        }
    }
    println!(
        "{}",
        serde_json::json!({"output": output, "target_append_plans": append_plans})
    );
    Ok(())
}

fn run_summary_used_range_cli(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let output_dir = options
        .values
        .get("output-dir")
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("output"));
    if options.positional.is_empty() {
        bail!("provide at least one workbook path");
    }
    let result = summarize_used_ranges(&options.positional, &output_dir)?;
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn run_summary_comments_cli(arguments: Vec<String>) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let template = required_path(&options, "template")?;
    let output_dir = options
        .values
        .get("output-dir")
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("output"));
    if options.positional.is_empty() {
        bail!("provide at least one source workbook path");
    }
    let result = summarize_by_comments(&template, &options.positional, &output_dir)?;
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn run_archive_cli(arguments: Vec<String>, operation: &str) -> Result<()> {
    let options = parse_cli_arguments(arguments)?;
    let output_dir = options
        .values
        .get("output-dir")
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("output"));
    let result = match operation {
        "date-preview" => archive_by_date_preview(
            &options.positional,
            &required_path(&options, "target-root")?,
            &output_dir,
            options.values.get("config").map(Path::new),
        )?,
        "type-preview" => archive_by_type_preview(
            &options.positional,
            &required_path(&options, "target-root")?,
            &output_dir,
            &required_path(&options, "config")?,
        )?,
        "date-copy" => archive_by_date_copy(
            &options.positional,
            &required_path(&options, "target-root")?,
            &output_dir,
            options.values.get("config").map(Path::new),
        )?,
        "type-copy" => archive_by_type_copy(
            &options.positional,
            &required_path(&options, "target-root")?,
            &output_dir,
            &required_path(&options, "config")?,
        )?,
        "plan" => execute_archive_plan(&required_path(&options, "plan")?, &output_dir)?,
        _ => unreachable!(),
    };
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

struct ParsedCli {
    values: std::collections::BTreeMap<String, String>,
    positional: Vec<PathBuf>,
}
fn parse_cli_arguments(arguments: Vec<String>) -> Result<ParsedCli> {
    let mut values = std::collections::BTreeMap::new();
    let mut positional = Vec::new();
    let mut args = arguments.into_iter();
    while let Some(value) = args.next() {
        if value == "--help" || value == "-h" {
            print_usage();
            return Ok(ParsedCli { values, positional });
        }
        if let Some(name) = value.strip_prefix("--") {
            let argument = args
                .next()
                .with_context(|| format!("--{name} requires a value"))?;
            values.insert(name.to_owned(), argument);
        } else if value.starts_with('-') {
            bail!("unknown option {value:?}");
        } else {
            positional.push(PathBuf::from(value));
        }
    }
    Ok(ParsedCli { values, positional })
}

fn extract_boolean_flag(arguments: Vec<String>, name: &str) -> Result<(bool, Vec<String>)> {
    let flag = format!("--{name}");
    let inline_prefix = format!("{flag}=");
    anyhow::ensure!(
        !arguments
            .windows(2)
            .any(|pair| { pair[0] == flag && matches!(pair[1].as_str(), "true" | "false") }),
        "{flag} is a boolean switch and does not take a value"
    );
    let mut found = false;
    let mut filtered = Vec::with_capacity(arguments.len());
    for argument in arguments {
        if argument == flag {
            anyhow::ensure!(!found, "{flag} may only be specified once");
            found = true;
        } else if argument.starts_with(&inline_prefix) {
            bail!("{flag} is a boolean switch and does not take a value");
        } else {
            filtered.push(argument);
        }
    }
    Ok((found, filtered))
}
fn required_path(options: &ParsedCli, name: &str) -> Result<PathBuf> {
    options
        .values
        .get(name)
        .map(PathBuf::from)
        .with_context(|| format!("--{name} is required"))
}

fn unique_workbook_candidate(input: &Path, output_dir: &Path, suffix: &str) -> PathBuf {
    use std::hash::{Hash, Hasher};
    let stem = input.file_stem().unwrap_or_default().to_string_lossy();
    let mut hasher = std::collections::hash_map::DefaultHasher::new();
    input.hash(&mut hasher);
    let base = format!("{stem}_{suffix}_{:08x}", hasher.finish() as u32);
    let mut candidate = output_dir.join(format!("{base}.xlsx"));
    let mut index = 1usize;
    while candidate.exists() {
        candidate = output_dir.join(format!("{base}_{index}.xlsx"));
        index += 1;
    }
    candidate
}

fn run_local_filename_map(arguments: Vec<String>) -> Result<()> {
    let mut arguments = arguments.into_iter();
    let mut config_path = None::<PathBuf>;
    let mut files = Vec::new();
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--config" => {
                config_path = Some(PathBuf::from(
                    arguments.next().context("--config requires a path")?,
                ));
            }
            "--help" | "-h" => {
                print_usage();
                return Ok(());
            }
            value if value.starts_with('-') => bail!("unknown option {value:?}"),
            path => files.push(PathBuf::from(path)),
        }
    }
    let config_path = config_path.context("--config is required")?;
    if files.is_empty() {
        bail!("provide at least one file path");
    }

    let result = run_local_filename_map_48_candidate(&config_path, &files)?;
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn run_batch_rename(arguments: Vec<String>) -> Result<()> {
    let mut arguments = arguments.into_iter();
    let mut config_path = None::<PathBuf>;
    let mut log_dir = None::<PathBuf>;
    let mut paths = Vec::new();
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--config" => {
                config_path = Some(PathBuf::from(
                    arguments.next().context("--config requires a path")?,
                ));
            }
            "--log-dir" => {
                log_dir = Some(PathBuf::from(
                    arguments.next().context("--log-dir requires a path")?,
                ));
            }
            "--help" | "-h" => {
                print_usage();
                return Ok(());
            }
            value if value.starts_with('-') => bail!("unknown option {value:?}"),
            path => paths.push(PathBuf::from(path)),
        }
    }

    let config_path = config_path.context("--config is required")?;
    let log_dir = log_dir.unwrap_or_else(|| PathBuf::from("logs"));
    let (log_path, mut log_file) = open_batch_rename_log(&log_dir)?;
    let result = run_batch_rename_33(&config_path, &paths)?;
    let log_warning = write_batch_rename_log(&mut log_file, &result.rows).err();
    let mut output = serde_json::to_value(result)?;
    if let Some(object) = output.as_object_mut() {
        object.insert(
            "log_path".to_owned(),
            serde_json::json!(log_path.to_string_lossy()),
        );
        if let Some(warning) = log_warning {
            object.insert(
                "log_warning".to_owned(),
                serde_json::json!(warning.to_string()),
            );
        }
    }
    println!("{}", serde_json::to_string_pretty(&output)?);
    Ok(())
}

fn open_batch_rename_log(log_dir: &Path) -> Result<(PathBuf, File)> {
    fs::create_dir_all(log_dir)
        .with_context(|| format!("create batch-rename log directory {}", log_dir.display()))?;
    let timestamp = Local::now().format("%Y%m%d_%H%M%S");
    let base = format!("3_3_batch_rename_{timestamp}");
    for index in 0usize.. {
        let name = if index == 0 {
            format!("{base}.log")
        } else {
            format!("{base}_{index}.log")
        };
        let path = log_dir.join(name);
        match OpenOptions::new().write(true).create_new(true).open(&path) {
            Ok(file) => return Ok((path, file)),
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(error) => {
                return Err(error).with_context(|| format!("create log {}", path.display()))
            }
        }
    }
    unreachable!("usize log suffix range exhausted")
}

fn write_batch_rename_log(
    file: &mut File,
    rows: &[pytools_core::batch_rename::BatchRenameRow],
) -> Result<()> {
    for row in rows {
        let (level, message) = if row.status == "renamed" {
            (
                "INFO",
                format!("ok old={} new={}", row.original_name, row.new_name),
            )
        } else {
            let file = if row.original_name.is_empty() {
                row.source_path.as_str()
            } else {
                row.original_name.as_str()
            };
            let level = if row.reason.starts_with("rename_failed") {
                "WARNING"
            } else {
                "INFO"
            };
            (level, format!("skip file={file} reason={}", row.reason))
        };
        let line = format!(
            "{} [{level}] {message}",
            Local::now().format("%Y-%m-%d %H:%M:%S")
        );
        writeln!(file, "{line}")?;
        eprintln!("{line}");
    }
    file.flush()?;
    Ok(())
}

fn run_content_prefix_rename_46(arguments: Vec<String>) -> Result<()> {
    let Some((config_path, files)) = parse_config_and_files(arguments)? else {
        return Ok(());
    };
    let result = run_content_prefix_rename_46_xlsx_candidate(&config_path, &files)?;
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn run_content_prefix_unrename_47(arguments: Vec<String>) -> Result<()> {
    let Some((config_path, files)) = parse_config_and_files(arguments)? else {
        return Ok(());
    };
    let result = run_content_prefix_unrename_47_xlsx_candidate(&config_path, &files)?;
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn parse_config_and_files(arguments: Vec<String>) -> Result<Option<(PathBuf, Vec<PathBuf>)>> {
    let mut arguments = arguments.into_iter();
    let mut config_path = None::<PathBuf>;
    let mut files = Vec::new();
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--config" => {
                config_path = Some(PathBuf::from(
                    arguments.next().context("--config requires a path")?,
                ));
            }
            "--help" | "-h" => {
                print_usage();
                return Ok(None);
            }
            value if value.starts_with('-') => bail!("unknown option {value:?}"),
            path => files.push(PathBuf::from(path)),
        }
    }

    let config_path = config_path.context("--config is required")?;
    if files.is_empty() {
        bail!("provide at least one file path");
    }
    Ok(Some((config_path, files)))
}

fn run_file_sort_prefix(arguments: Vec<String>) -> Result<()> {
    let mut arguments = arguments.into_iter();
    let mut config_path = None::<PathBuf>;
    let mut output_dir = None::<PathBuf>;
    let mut log_dir = None::<PathBuf>;
    let mut files = Vec::new();
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--config" => {
                config_path = Some(PathBuf::from(
                    arguments.next().context("--config requires a path")?,
                ));
            }
            "--output-dir" => {
                output_dir = Some(PathBuf::from(
                    arguments.next().context("--output-dir requires a path")?,
                ));
            }
            "--log-dir" => {
                log_dir = Some(PathBuf::from(
                    arguments.next().context("--log-dir requires a path")?,
                ));
            }
            "--help" | "-h" => {
                print_usage();
                return Ok(());
            }
            value if value.starts_with('-') => bail!("unknown option {value:?}"),
            path => files.push(PathBuf::from(path)),
        }
    }

    let config_path = config_path.context("--config is required")?;
    if files.is_empty() {
        bail!("provide at least one file path");
    }
    let output_dir = output_dir.unwrap_or_else(|| PathBuf::from("output"));
    let log_dir = log_dir.unwrap_or_else(|| PathBuf::from("logs"));
    let (log_path, mut log_file) = open_file_sort_log(&log_dir)?;

    let result = run_file_sort_prefix_49(&config_path, &files, &output_dir)?;
    let log_warning = write_file_sort_log(&mut log_file, &result.rows).err();
    let mut output = serde_json::to_value(result)?;
    output["log_path"] = serde_json::json!(log_path);
    if let Some(error) = log_warning {
        let warning = format!("could not write complete operation log: {error:#}");
        eprintln!("pytools: warning: {warning}");
        output["log_warning"] = serde_json::json!(warning);
    }
    println!("{}", serde_json::to_string_pretty(&output)?);
    Ok(())
}

fn open_file_sort_log(log_dir: &Path) -> Result<(PathBuf, File)> {
    fs::create_dir_all(log_dir)
        .with_context(|| format!("create log directory {}", log_dir.display()))?;
    let timestamp = Local::now().format("%Y%m%d_%H%M%S");
    let path = log_dir.join(format!("4_9_file_sort_prefix_{timestamp}.log"));
    let file = OpenOptions::new()
        .create(true)
        .append(true)
        .open(&path)
        .with_context(|| format!("open operation log {}", path.display()))?;
    Ok((path, file))
}

fn write_file_sort_log(
    file: &mut File,
    rows: &[pytools_core::file_sort_prefix::FileSortPrefixRow],
) -> Result<()> {
    for row in rows {
        let message = if let Some(config_row) = row.config_row {
            Some((
                "INFO",
                format!(
                    "ok old={} new={} row={} prefix={}",
                    row.original_name, row.new_name, config_row, row.sort_prefix
                ),
            ))
        } else if row.status == "失败" {
            Some((
                "WARNING",
                format!("failed file={} err={}", row.source_path, row.remark),
            ))
        } else {
            None
        };
        if let Some((level, message)) = message {
            let line = format!(
                "{} [{level}] {message}",
                Local::now().format("%Y-%m-%d %H:%M:%S")
            );
            writeln!(file, "{line}")?;
            eprintln!("{line}");
        }
    }
    file.flush()?;
    Ok(())
}

fn run_remove_file_sort_prefix(arguments: Vec<String>) -> Result<()> {
    let mut arguments = arguments.into_iter();
    let mut config_path = None::<PathBuf>;
    let mut output_dir = None::<PathBuf>;
    let mut log_dir = None::<PathBuf>;
    let mut files = Vec::new();
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--config" => {
                config_path = Some(PathBuf::from(
                    arguments.next().context("--config requires a path")?,
                ));
            }
            "--output-dir" => {
                output_dir = Some(PathBuf::from(
                    arguments.next().context("--output-dir requires a path")?,
                ));
            }
            "--log-dir" => {
                log_dir = Some(PathBuf::from(
                    arguments.next().context("--log-dir requires a path")?,
                ));
            }
            "--help" | "-h" => {
                print_usage();
                return Ok(());
            }
            value if value.starts_with('-') => bail!("unknown option {value:?}"),
            path => files.push(PathBuf::from(path)),
        }
    }

    let config_path = config_path.context("--config is required")?;
    let output_dir = output_dir.unwrap_or_else(|| PathBuf::from("output"));
    let log_dir = log_dir.unwrap_or_else(|| PathBuf::from("logs"));
    let (log_path, mut log_file) = open_file_sort_removal_log(&log_dir)?;

    let result = run_remove_file_sort_prefix_410(&config_path, &files, &output_dir)?;
    let log_warning = write_file_sort_removal_log(&mut log_file, &result.rows).err();
    let mut output = serde_json::to_value(result)?;
    output["log_path"] = serde_json::json!(log_path);
    if let Some(error) = log_warning {
        let warning = format!("could not write complete operation log: {error:#}");
        eprintln!("pytools: warning: {warning}");
        output["log_warning"] = serde_json::json!(warning);
    }
    println!("{}", serde_json::to_string_pretty(&output)?);
    Ok(())
}

fn open_file_sort_removal_log(log_dir: &Path) -> Result<(PathBuf, File)> {
    fs::create_dir_all(log_dir)
        .with_context(|| format!("create log directory {}", log_dir.display()))?;
    let timestamp = Local::now().format("%Y%m%d_%H%M%S");
    let path = log_dir.join(format!("4_10_remove_file_sort_prefix_{timestamp}.log"));
    let file = OpenOptions::new()
        .create(true)
        .append(true)
        .open(&path)
        .with_context(|| format!("open operation log {}", path.display()))?;
    Ok((path, file))
}

fn write_file_sort_removal_log(
    file: &mut File,
    rows: &[pytools_core::file_sort_prefix::FileSortPrefixRemovalRow],
) -> Result<()> {
    for row in rows {
        let message = if !row.new_name.is_empty() {
            Some((
                "INFO",
                format!("ok old={} new={}", row.original_name, row.new_name),
            ))
        } else if row.status == "失败" {
            Some((
                "WARNING",
                format!("failed file={} err={}", row.source_path, row.remark),
            ))
        } else {
            None
        };
        if let Some((level, message)) = message {
            let line = format!(
                "{} [{level}] {message}",
                Local::now().format("%Y-%m-%d %H:%M:%S")
            );
            writeln!(file, "{line}")?;
            eprintln!("{line}");
        }
    }
    file.flush()?;
    Ok(())
}

fn run_extract_sheets(arguments: Vec<String>) -> Result<()> {
    let mut arguments = arguments.into_iter();
    let mut config_path = None::<PathBuf>;
    let mut output_dir = None::<PathBuf>;
    let mut capability_path = None::<PathBuf>;
    let mut platform = None::<String>;
    let mut source_workbooks = Vec::new();
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--config" => {
                config_path = Some(PathBuf::from(
                    arguments.next().context("--config requires a path")?,
                ));
            }
            "--output-dir" => {
                output_dir = Some(PathBuf::from(
                    arguments.next().context("--output-dir requires a path")?,
                ));
            }
            "--capabilities" => {
                capability_path = Some(PathBuf::from(
                    arguments.next().context("--capabilities requires a path")?,
                ));
            }
            "--platform" => {
                platform = Some(arguments.next().context("--platform requires an id")?);
            }
            "--help" | "-h" => {
                print_usage();
                return Ok(());
            }
            value if value.starts_with('-') => bail!("unknown option {value:?}"),
            path => source_workbooks.push(PathBuf::from(path)),
        }
    }

    let config_path = config_path.context("--config is required")?;
    if source_workbooks.is_empty() {
        bail!("provide at least one source workbook path");
    }
    anyhow::ensure!(
        capability_path.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    let tasks = load_extract_tasks(&config_path).with_context(|| {
        format!(
            "load worksheet-extraction tasks from {}",
            config_path.display()
        )
    })?;
    let output_dir = match output_dir {
        Some(path) => path,
        None => {
            load_global(&config_path)
                .with_context(|| format!("load output directory from {}", config_path.display()))?
                .output_dir
        }
    };
    let result = run_extract_sheets_candidate(
        &source_workbooks,
        &tasks,
        &output_dir,
        false,
        None,
        capability_path.as_deref(),
        platform.as_deref(),
    )?;
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn run_region_sum(arguments: Vec<String>) -> Result<()> {
    let mut arguments = arguments.into_iter();
    let mut config_path = None::<PathBuf>;
    let mut output_dir = None::<PathBuf>;
    let mut log_dir = None::<PathBuf>;
    let mut source_workbooks = Vec::new();
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--config" => {
                config_path = Some(PathBuf::from(
                    arguments.next().context("--config requires a path")?,
                ));
            }
            "--output-dir" => {
                output_dir = Some(PathBuf::from(
                    arguments.next().context("--output-dir requires a path")?,
                ));
            }
            "--log-dir" => {
                log_dir = Some(PathBuf::from(
                    arguments.next().context("--log-dir requires a path")?,
                ));
            }
            "--help" | "-h" => {
                print_usage();
                return Ok(());
            }
            value if value.starts_with('-') => bail!("unknown option {value:?}"),
            path => source_workbooks.push(PathBuf::from(path)),
        }
    }

    if source_workbooks.is_empty() {
        bail!("provide at least one source workbook path");
    }

    let global = match config_path.as_ref() {
        Some(path) => Some(
            load_global(path)
                .with_context(|| format!("load global configuration from {}", path.display()))?,
        ),
        None => None,
    };
    let output_dir = output_dir
        .or_else(|| global.as_ref().map(|config| config.output_dir.clone()))
        .unwrap_or_else(|| PathBuf::from("output"));
    let log_dir = log_dir
        .or_else(|| global.as_ref().map(|config| config.log_dir.clone()))
        .unwrap_or_else(|| PathBuf::from("logs"));

    let result = run_region_sum_check(&source_workbooks, &output_dir, &log_dir)?;
    println!(
        "{}",
        serde_json::to_string_pretty(&serde_json::json!({
            "files": result.files,
            "sheets": result.sheets,
            "errors": result.errors,
            "failed_files": result.failed_files,
            "path": result.path.map(|path| path.to_string_lossy().into_owned()),
            "log_path": result.log_path.to_string_lossy(),
        }))?
    );
    Ok(())
}

fn print_usage() {
    eprintln!(
        "Usage:\n\
         timeline-precheck --config <workbook.xlsx>\n\
         timeline-compare --config <workbook.xlsx> --template <workbook.xlsx> <source.xlsx>...\n\
         wide-summary --config <workbook.xlsx> <source.xlsx>...\n\
         wide-summary-row-suffix --config <workbook.xlsx> <source.xlsx>...\n\
         wide-summary target writes produce append-plan JSON; execute with dedup-append-execute and matching capability evidence\n\
         summary-used-range [--output-dir <dir>] <workbook>...\n\
         summary-comments --template <workbook.xlsx> [--output-dir <dir>] <source.xlsx>...\n\
         dedup-comment-plan --input <workbook.xlsx> --output <candidate.xlsx> [--action check|mark|delete] [--plan-output <plan.json>] [--capabilities <json> --platform <id>]\n\
         dedup-config-plan --config <workbook.xlsx> [--output-dir <dir>] [--action check|mark|delete] [--capabilities <json> --platform <id>]\n\
         dedup-precheck --config <workbook.xlsx>\n\
         dedup-append-plan --config <workbook.xlsx> [--output-dir <dir>]\n\
         dedup-append-execute --plan <append-plan.json> [--capabilities <json> --platform <id>] [--commit-target]\n\
         print-comment-plan --input <workbook.xlsx> --output <candidate.xlsx> [--drop-comments true] [--hide-zero true] [--plan-output <plan.json>] [--capabilities <json> --platform <id>]\n\
         print-config-precheck --config <workbook.xlsx>\n\
         print-config-plan --config <workbook.xlsx> [--output-dir <plan-report-dir>] [--capabilities <json> --platform <id>]\n\
         apply-edit-plan --plan <plan.json> --capabilities <json> --platform <id>\n\
         config-init --path <new-config.xlsx>\n\
         config-repair-plan --input <config.xlsx> --output <candidate.xlsx> [--plan-output <plan.json>] [--capabilities <json> --platform <id>]\n\
         config-view --config <config.xlsx>\n\
         village-split-plan --current <xlsx> --previous <xlsx> --input <target.xlsx> --output <candidate.xlsx> [--plan-output <json>] [--capabilities <json> --platform <id>]\n\
         remove-foreign-bank-plan --config <config.xlsx> --input <xlsx> --output <candidate.xlsx> [--plan-output <json>] [--capabilities <json> --platform <id>]\n\
         remove-foreign-bank --config <config.xlsx> [--capabilities <json> --platform <id>] <source.xlsx>...\n\
         rural-loan-plan --current <xlsx> --previous <xlsx> --output <candidate.xlsx> [--plan-output <json>] [--capabilities <json> --platform <id>]\n\
         quarterly-rural-loan-plan --current <xlsx> --previous <xlsx> --output <candidate.xlsx> [--plan-output <json>] [--capabilities <json> --platform <id>]\n\
         institution-normalize-plan --config <config.xlsx> --input <xlsx> --output <candidate.xlsx> [--plan-output <json>] [--capabilities <json> --platform <id>]\n\
         institution-normalize --config <config.xlsx> --plan-dir <dir> [--view user|developer] [--report-output <json>] <workbook.xlsx>...\n\
         institution-normalize --config <config.xlsx> --apply --capabilities <json> --platform <id> [--view user|developer] [--report-output <json>] <workbook.xlsx>...\n\
         fx-header-plan --input <xlsx> --output <candidate.xlsx> [--plan-output <json>] [--capabilities <json> --platform <id>]\n\
         rename-sheets-plan --config <config.xlsx> --input <xlsx> --output <candidate.xlsx> [--plan-output <json>] [--capabilities <json> --platform <id>]\n\
         rename-sheets --config <config.xlsx> --capabilities <json> --platform <id> [--recursive] <file-or-directory>...\n\
         submission-check --config <config.xlsx> [--output-dir <dir>] <file-or-directory>...\n\
         archive-date-preview --target-root <dir> [--config <config.xlsx>] [--output-dir <dir>] <file-or-directory>...\n\
         archive-type-preview --config <config.xlsx> --target-root <dir> [--output-dir <dir>] <file-or-directory>...\n\
         archive-date-copy --target-root <dir> [--config <config.xlsx>] [--output-dir <dir>] <file-or-directory>...\n\
         archive-type-copy --config <config.xlsx> --target-root <dir> [--output-dir <dir>] <file-or-directory>...\n\
         archive-execute-plan --plan <archive-plan.xlsx> [--output-dir <dir>]\n\
         convert-excel --config <config.xlsx> [--to xlsx] [--engine auto|lo|native] [--output-dir <dir>] <file>...\n\
         convert-word --config <config.xlsx> [--to docx] [--engine auto|lo] [--output-dir <dir>] <file>...\n\
         convert-file --input <file> --to <extension> [--output <file>] [--engine auto|lo|native]\n\
         native = 按值重建（C01，丢弃样式/公式缓存）；Excel 公式文本保留但不重算，CSV 按文本写入\n\
         pytools extract-sheets --config <workbook.xlsx> [--output-dir <dir>] [--capabilities <json> --platform <id>] <source.xlsx>...\n\
         pytools region-sum-check [--config <workbook.xlsx>] [--output-dir <dir>] [--log-dir <dir>] <source.xlsx>...\n\
         pytools file-sort-prefix --config <workbook.xlsx> [--output-dir <dir>] [--log-dir <dir>] <file>...\n\
         pytools remove-file-sort-prefix --config <workbook.xlsx> [--output-dir <dir>] [--log-dir <dir>] [<file>...]\n\
         pytools rename-by-content --config <workbook.xlsx> <file>...\n\
         pytools unrename-by-content --config <workbook.xlsx> <file>...\n\
         pytools rename-local-map --config <workbook.xlsx> <file>...\n\
         pytools rename-batch --config <workbook.xlsx> [--log-dir <dir>] [<file-or-directory>...]\n\
         The region-sum command runs the Rust non-COM regional total check. Without --config,\n\
         output and logs default to ./output and ./logs. File-sort-prefix writes its XLSX report\n\
         under ./output unless --output-dir is supplied; its log defaults to ./logs. The\n\
         remove-file-sort-prefix command also writes under ./output and ./logs by default; it\n\
         accepts an empty file list to produce the source-compatible blank report. Content-prefix commands accept only .xlsx/.xlsm files; legacy .xls/.xlt COM reads are not part of the Rust path."
    );
}
