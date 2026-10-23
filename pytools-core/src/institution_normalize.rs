//! Candidate Rust implementation of pytools institution-name normalization.
//!
//! The workbook adapter currently accepts `.xlsx` only. It is not a default
//! edit backend until preservation parity is established on real templates.

use anyhow::{anyhow, bail, ensure, Context, Result};
use excel_editor::{
    CellValue as PlannedCellValue, EditOperation, ExcelEditPlan, PreservationRequirement,
    WorkbookFormat, WorkbookProfile,
};
use excel_reader::{read_workbook, CellValue as ReaderCellValue, SheetDump};
use serde::Serialize;
use serde_json::Value;
use std::{
    collections::{BTreeMap, BTreeSet},
    fs::{self, OpenOptions},
    io::Write,
    path::{Path, PathBuf},
    time::{Instant, SystemTime, UNIX_EPOCH},
};

#[cfg(test)]
use excel_editor::{open_xlsx, save_xlsx};
#[cfg(test)]
use rayon::prelude::*;
#[cfg(test)]
use std::collections::HashSet;

const BRANCH_SHEET_KEY: &str = "分机构";

#[cfg(test)]
#[derive(Debug, Default, Clone, PartialEq, Eq, Serialize)]
pub struct NormalizeReport {
    pub files: usize,
    pub branch_sheets: usize,
    pub mapped_sheets: usize,
    pub mapped_rows: usize,
    pub failed_files: usize,
    pub mapping_empty: bool,
    pub errors: Vec<String>,
}

#[derive(Debug, Default, Clone, PartialEq)]
pub struct NormalizePlanResult {
    pub branch_sheets: usize,
    pub mapped_sheets: usize,
    pub mapped_rows: usize,
    pub mapping_empty: bool,
    pub plan: Option<ExcelEditPlan>,
}

#[derive(Debug, Clone)]
pub struct InstitutionNormalizeRunOptions {
    pub apply: bool,
    pub capabilities: Option<PathBuf>,
    pub platform: Option<String>,
    pub plan_dir: Option<PathBuf>,
}

#[derive(Debug, Clone, Default, Serialize)]
pub struct InstitutionNormalizeProgress {
    pub total_files: usize,
    pub current_file: Option<usize>,
    pub processed_files: usize,
    pub completed_files: usize,
    pub failed_files: usize,
    pub modified_files: usize,
}

#[derive(Debug, Clone, Serialize)]
pub struct InstitutionNormalizeEvent {
    pub sequence: usize,
    pub timestamp: String,
    pub phase: String,
    pub status: String,
    pub user_message: String,
    pub progress: InstitutionNormalizeProgress,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub file_name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub elapsed_ms: Option<u128>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub components: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub actual_engine: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub diagnostic: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub route_diagnostic: Option<Value>,
}

#[derive(Debug, Clone, Default, Serialize)]
pub struct InstitutionNormalizeStageTimings {
    pub source_fingerprint_ms: u128,
    pub profile_inspection_ms: u128,
    pub plan_build_ms: u128,
    pub plan_persistence_ms: u128,
    pub source_stability_check_ms: u128,
    pub engine_execution_ms: u128,
    pub safe_commit_ms: u128,
    pub total_ms: u128,
}

#[derive(Debug, Clone, Serialize)]
pub struct InstitutionNormalizeFileResult {
    pub file_name: String,
    pub status: String,
    pub branch_sheets: usize,
    pub mapped_sheets: usize,
    pub mapped_rows: usize,
    pub operation_count: usize,
    pub plan_path: Option<PathBuf>,
    pub actual_engine: Option<String>,
    pub stage_components: BTreeMap<String, Vec<String>>,
    pub execution: Option<Value>,
    pub timings: InstitutionNormalizeStageTimings,
    pub error_stage: Option<String>,
    pub error_type: Option<String>,
    pub user_message: String,
    pub diagnostic: Option<String>,
    pub route_diagnostic: Option<Value>,
    pub cleanup_warning: Option<String>,
}

#[derive(Debug, Clone, Serialize)]
pub struct InstitutionNormalizeRunReport {
    pub schema_version: u32,
    pub product: String,
    pub feature_id: String,
    pub run_id: String,
    pub started_at: String,
    pub mode: String,
    pub status: String,
    pub total_inputs: usize,
    pub completed_files: usize,
    pub failed_files: usize,
    pub modified_files: usize,
    pub branch_sheets: usize,
    pub mapped_sheets: usize,
    pub mapped_rows: usize,
    pub mapping_empty: bool,
    pub configuration_elapsed_ms: u128,
    pub elapsed_ms: u128,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub failure: Option<InstitutionNormalizeRunFailure>,
    pub events: Vec<InstitutionNormalizeEvent>,
    pub files: Vec<InstitutionNormalizeFileResult>,
}

#[derive(Debug, Clone, Serialize)]
pub struct InstitutionNormalizeRunFailure {
    pub stage: String,
    pub error_type: String,
    pub user_message: String,
    pub diagnostic: String,
}

#[derive(Debug)]
struct InstitutionNormalizeFileFailure {
    stage: &'static str,
    error_type: &'static str,
    diagnostic: String,
    route_diagnostic: Option<Value>,
    actual_engine: Option<String>,
    execution: Option<Value>,
    timings: InstitutionNormalizeStageTimings,
    stage_components: BTreeMap<String, Vec<String>>,
}

impl InstitutionNormalizeFileFailure {
    fn attach_state(
        mut self,
        timings: &InstitutionNormalizeStageTimings,
        stage_components: &BTreeMap<String, Vec<String>>,
        actual_engine: Option<&str>,
    ) -> Self {
        self.timings = timings.clone();
        self.stage_components = stage_components.clone();
        self.actual_engine = actual_engine.map(str::to_owned);
        self
    }

    fn attach_execution(mut self, execution: Option<Value>) -> Self {
        self.execution = execution;
        self
    }
}

/// Run PT-7-2 against an explicit workbook list. Planning is the default; the
/// optional apply path requires a matching, evidence-backed ExcelModifier
/// capability and commits each successful workbook through file_commit.
pub fn run_institution_normalize_files<F>(
    config_path: impl AsRef<Path>,
    inputs: &[PathBuf],
    options: InstitutionNormalizeRunOptions,
    mut on_event: F,
) -> Result<InstitutionNormalizeRunReport>
where
    F: FnMut(&InstitutionNormalizeEvent),
{
    ensure!(!inputs.is_empty(), "provide at least one workbook path");
    ensure!(
        options.apply == options.capabilities.is_some()
            && options.apply == options.platform.is_some(),
        "--apply requires both --capabilities and --platform; omit all three for plan-only mode"
    );
    ensure!(
        options.apply || options.plan_dir.is_some(),
        "plan-only mode requires --plan-dir so the generated plans can be reviewed"
    );

    let started_at = chrono::Utc::now().to_rfc3339();
    let run_nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .context("system clock is before Unix epoch")?
        .as_nanos();
    let run_id = format!("pytools-pt-7-2-{}-{run_nonce}", std::process::id());
    let run_started = Instant::now();
    let mut report = InstitutionNormalizeRunReport {
        schema_version: 1,
        product: "pytools".to_owned(),
        feature_id: "PT-7-2".to_owned(),
        run_id: run_id.clone(),
        started_at,
        mode: if options.apply {
            "apply".to_owned()
        } else {
            "plan_only".to_owned()
        },
        status: "running".to_owned(),
        total_inputs: inputs.len(),
        completed_files: 0,
        failed_files: 0,
        modified_files: 0,
        branch_sheets: 0,
        mapped_sheets: 0,
        mapped_rows: 0,
        mapping_empty: false,
        configuration_elapsed_ms: 0,
        elapsed_ms: 0,
        failure: None,
        events: Vec::new(),
        files: Vec::new(),
    };

    emit_institution_normalize_event(
        &mut report,
        &mut on_event,
        "configuration_read",
        "started",
        "正在读取机构映射配置。".to_owned(),
        InstitutionNormalizeProgress {
            total_files: inputs.len(),
            ..Default::default()
        },
        None,
        None,
        vec!["Rust".to_owned(), "calamine".to_owned()],
        None,
        None,
        None,
        None,
    );
    let config_started = Instant::now();
    let mapping = match crate::config::load_institution_mapping_ordered(config_path.as_ref()) {
        Ok(mapping) => mapping,
        Err(error) => {
            report.configuration_elapsed_ms = config_started.elapsed().as_millis();
            report.status = "failed".to_owned();
            let error_type = "configuration_read_failed".to_owned();
            let user_message = "机构映射配置读取失败，未处理所选工作簿。".to_owned();
            let diagnostic = format!("PT-7-2 stage=configuration_read: {error:#}");
            report.failure = Some(InstitutionNormalizeRunFailure {
                stage: "configuration_read".to_owned(),
                error_type: error_type.clone(),
                user_message: user_message.clone(),
                diagnostic: diagnostic.clone(),
            });
            let config_elapsed_ms = report.configuration_elapsed_ms;
            emit_institution_normalize_event(
                &mut report,
                &mut on_event,
                "configuration_read",
                "failed",
                user_message,
                InstitutionNormalizeProgress {
                    total_files: inputs.len(),
                    ..Default::default()
                },
                None,
                Some(config_elapsed_ms),
                vec!["Rust".to_owned(), "calamine".to_owned()],
                None,
                Some(error_type.clone()),
                Some(diagnostic.clone()),
                None,
            );
            report.elapsed_ms = run_started.elapsed().as_millis();
            let elapsed_ms = report.elapsed_ms;
            let final_progress = current_institution_progress(&report, inputs.len());
            emit_institution_normalize_event(
                &mut report,
                &mut on_event,
                "task",
                "failed",
                "任务失败：映射配置未能读取，未处理工作簿。".to_owned(),
                final_progress,
                None,
                Some(elapsed_ms),
                vec!["Rust".to_owned()],
                None,
                Some(error_type),
                None,
                None,
            );
            return Ok(report);
        }
    };
    report.configuration_elapsed_ms = config_started.elapsed().as_millis();
    let configuration_elapsed_ms = report.configuration_elapsed_ms;
    emit_institution_normalize_event(
        &mut report,
        &mut on_event,
        "configuration_read",
        "completed",
        format!("机构映射配置已读取，共 {} 条有效映射。", mapping.len()),
        InstitutionNormalizeProgress {
            total_files: inputs.len(),
            ..Default::default()
        },
        None,
        Some(configuration_elapsed_ms),
        vec!["Rust".to_owned(), "calamine".to_owned()],
        None,
        None,
        None,
        None,
    );

    if mapping.is_empty() {
        report.mapping_empty = true;
        emit_institution_normalize_event(
            &mut report,
            &mut on_event,
            "normalization",
            "skipped",
            "机构映射表为空，未处理所选工作簿。".to_owned(),
            InstitutionNormalizeProgress {
                total_files: inputs.len(),
                ..Default::default()
            },
            None,
            None,
            vec!["Rust".to_owned()],
            None,
            None,
            None,
            None,
        );
        report.status = "completed".to_owned();
        report.elapsed_ms = run_started.elapsed().as_millis();
        let elapsed_ms = report.elapsed_ms;
        let final_progress = current_institution_progress(&report, inputs.len());
        emit_institution_normalize_event(
            &mut report,
            &mut on_event,
            "task",
            "completed",
            "任务结束：机构映射表为空，未处理工作簿。".to_owned(),
            final_progress,
            None,
            Some(elapsed_ms),
            vec!["Rust".to_owned()],
            None,
            None,
            None,
            None,
        );
        return Ok(report);
    }

    for (index, input) in inputs.iter().enumerate() {
        let file_name = input
            .file_name()
            .unwrap_or_default()
            .to_string_lossy()
            .into_owned();
        let mut progress = current_institution_progress(&report, inputs.len());
        progress.current_file = Some(index + 1);
        emit_institution_normalize_event(
            &mut report,
            &mut on_event,
            "file_processing",
            "started",
            format!("正在处理 {}/{}：{}", index + 1, inputs.len(), file_name),
            progress.clone(),
            Some(file_name.clone()),
            None,
            vec!["Rust".to_owned()],
            None,
            None,
            None,
            None,
        );

        let candidate = institution_normalize_candidate(input, &run_id, index + 1)?;
        let file_started = Instant::now();
        let outcome = build_and_optionally_apply_institution_plan(
            input,
            &candidate,
            index + 1,
            inputs.len(),
            &run_id,
            &mapping,
            &options,
            &mut report,
            &mut on_event,
            progress,
        );
        let mut file_result = match outcome {
            Ok(mut result) => {
                if candidate.exists() {
                    match fs::remove_file(&candidate) {
                        Ok(()) => {}
                        Err(error) => {
                            result.cleanup_warning =
                                Some(format!("could not remove temporary candidate: {error}"));
                        }
                    }
                }
                result
            }
            Err(failure) => {
                report.failed_files += 1;
                InstitutionNormalizeFileResult {
                    file_name: file_name.clone(),
                    status: "failed".to_owned(),
                    branch_sheets: 0,
                    mapped_sheets: 0,
                    mapped_rows: 0,
                    operation_count: 0,
                    plan_path: None,
                    actual_engine: failure.actual_engine.clone(),
                    stage_components: failure.stage_components.clone(),
                    execution: failure.execution.clone(),
                    timings: failure.timings.clone(),
                    error_stage: Some(failure.stage.to_owned()),
                    error_type: Some(failure.error_type.to_owned()),
                    user_message: user_message_for_normalize_failure(
                        failure.error_type,
                        &file_name,
                        index + 1,
                        inputs.len(),
                    ),
                    diagnostic: Some(failure.diagnostic),
                    route_diagnostic: failure.route_diagnostic,
                    cleanup_warning: None,
                }
            }
        };

        if file_result.status != "failed" {
            report.completed_files += 1;
            report.branch_sheets += file_result.branch_sheets;
            report.mapped_sheets += file_result.mapped_sheets;
            report.mapped_rows += file_result.mapped_rows;
            if file_result.status == "applied" {
                report.modified_files += 1;
            }
        }
        file_result.timings.total_ms = file_started.elapsed().as_millis();
        let event_status = if file_result.status == "failed" {
            "failed"
        } else {
            "completed"
        };
        let mut components = vec!["Rust".to_owned()];
        if let Some(engine) = file_result.actual_engine.as_deref() {
            components.push(engine.to_owned());
            components.push(engine_language(engine).to_owned());
        }
        let mut progress = current_institution_progress(&report, inputs.len());
        progress.current_file = Some(index + 1);
        emit_institution_normalize_event(
            &mut report,
            &mut on_event,
            "file_processing",
            event_status,
            file_result.user_message.clone(),
            progress,
            Some(file_name),
            Some(file_result.timings.total_ms),
            components,
            file_result.actual_engine.clone(),
            file_result.error_type.clone(),
            file_result.diagnostic.clone(),
            file_result.route_diagnostic.clone(),
        );
        report.files.push(file_result);
    }

    report.status = if report.failed_files == 0 {
        "completed".to_owned()
    } else {
        "partial_failure".to_owned()
    };
    report.elapsed_ms = run_started.elapsed().as_millis();
    let final_status = report.status.clone();
    let final_message = format!(
        "处理结束：成功 {} 个，失败 {} 个，修改 {} 个文件。",
        report.completed_files, report.failed_files, report.modified_files
    );
    let final_progress = current_institution_progress(&report, inputs.len());
    let elapsed_ms = report.elapsed_ms;
    emit_institution_normalize_event(
        &mut report,
        &mut on_event,
        "task",
        &final_status,
        final_message,
        final_progress,
        None,
        Some(elapsed_ms),
        vec!["Rust".to_owned()],
        None,
        None,
        None,
        None,
    );
    Ok(report)
}

fn build_and_optionally_apply_institution_plan(
    input: &Path,
    candidate: &Path,
    index: usize,
    total_files: usize,
    run_id: &str,
    mapping: &[(String, String)],
    options: &InstitutionNormalizeRunOptions,
    report: &mut InstitutionNormalizeRunReport,
    on_event: &mut impl FnMut(&InstitutionNormalizeEvent),
    progress: InstitutionNormalizeProgress,
) -> std::result::Result<InstitutionNormalizeFileResult, InstitutionNormalizeFileFailure> {
    let file_name = input
        .file_name()
        .unwrap_or_default()
        .to_string_lossy()
        .into_owned();
    let mut timings = InstitutionNormalizeStageTimings::default();
    let mut stage_components = BTreeMap::new();
    let mut actual_engine = None;
    if !has_extension(input, "xlsx") {
        return Err(normalize_file_failure(
            "format_check",
            "unsupported_format",
            "this PT-7-2 candidate currently supports .xlsx only; convert other formats before retrying",
            None,
        )
        .attach_state(&timings, &stage_components, actual_engine.as_deref()));
    }
    let metadata = fs::symlink_metadata(input).map_err(|error| {
        normalize_file_failure(
            "input_check",
            "input_unavailable",
            format!("inspect input workbook: {error}"),
            None,
        )
        .attach_state(&timings, &stage_components, actual_engine.as_deref())
    })?;
    if !metadata.is_file() || metadata.file_type().is_symlink() {
        return Err(normalize_file_failure(
            "input_check",
            "unsafe_input",
            "input must be a regular non-symlink file",
            None,
        )
        .attach_state(&timings, &stage_components, actual_engine.as_deref()));
    }

    let source_hash = if options.apply {
        let started = Instant::now();
        stage_components.insert(
            "source_fingerprint".to_owned(),
            vec!["Rust".to_owned(), "sha2".to_owned()],
        );
        emit_institution_stage(
            report,
            on_event,
            "source_fingerprint",
            "started",
            format!("正在校验源文件：{file_name}。"),
            &progress,
            &file_name,
            None,
            vec!["Rust".to_owned(), "sha2".to_owned()],
            None,
            None,
            None,
            None,
        );
        let hash_result = crate::file_commit::sha256_file(input);
        timings.source_fingerprint_ms = started.elapsed().as_millis();
        let hash = match hash_result {
            Ok(hash) => hash,
            Err(error) => {
                let failure = normalize_file_failure(
                    "source_fingerprint",
                    "input_unavailable",
                    format!("hash input workbook: {error:#}"),
                    None,
                )
                .attach_state(
                    &timings,
                    &stage_components,
                    actual_engine.as_deref(),
                );
                emit_institution_stage(
                    report,
                    on_event,
                    "source_fingerprint",
                    "failed",
                    format!("源文件安全检查失败：{file_name}。"),
                    &progress,
                    &file_name,
                    Some(timings.source_fingerprint_ms),
                    vec!["Rust".to_owned(), "sha2".to_owned()],
                    None,
                    Some(failure.error_type.to_owned()),
                    Some(failure.diagnostic.clone()),
                    None,
                );
                return Err(failure);
            }
        };
        emit_institution_stage(
            report,
            on_event,
            "source_fingerprint",
            "completed",
            format!("源文件安全检查完成：{file_name}。"),
            &progress,
            &file_name,
            Some(timings.source_fingerprint_ms),
            vec!["Rust".to_owned(), "sha2".to_owned()],
            None,
            None,
            None,
            None,
        );
        Some(hash)
    } else {
        None
    };

    let profile_started = Instant::now();
    stage_components.insert(
        "profile_inspection".to_owned(),
        vec!["Rust".to_owned(), "zip".to_owned(), "roxmltree".to_owned()],
    );
    emit_institution_stage(
        report,
        on_event,
        "profile_inspection",
        "started",
        format!("正在检查工作簿结构：{file_name}。"),
        &progress,
        &file_name,
        None,
        vec!["Rust".to_owned(), "zip".to_owned(), "roxmltree".to_owned()],
        None,
        None,
        None,
        None,
    );
    let profile_result = excel_editor::inspect_xlsx_profile(input);
    timings.profile_inspection_ms = profile_started.elapsed().as_millis();
    let profile = match profile_result {
        Ok(profile) => profile,
        Err(error) => {
            let failure = normalize_file_failure(
                "profile_inspection",
                "workbook_inspection_failed",
                format!("inspect workbook package: {error:#}"),
                None,
            )
            .attach_state(&timings, &stage_components, actual_engine.as_deref());
            emit_institution_stage(
                report,
                on_event,
                "profile_inspection",
                "failed",
                format!("工作簿结构检查失败：{file_name}。"),
                &progress,
                &file_name,
                Some(timings.profile_inspection_ms),
                vec!["Rust".to_owned(), "zip".to_owned(), "roxmltree".to_owned()],
                None,
                Some(failure.error_type.to_owned()),
                Some(failure.diagnostic.clone()),
                None,
            );
            return Err(failure);
        }
    };
    emit_institution_stage(
        report,
        on_event,
        "profile_inspection",
        "completed",
        format!("工作簿结构检查完成：{file_name}。"),
        &progress,
        &file_name,
        Some(timings.profile_inspection_ms),
        vec!["Rust".to_owned(), "zip".to_owned(), "roxmltree".to_owned()],
        None,
        None,
        None,
        None,
    );

    let plan_started = Instant::now();
    stage_components.insert(
        "plan_build".to_owned(),
        vec!["Rust".to_owned(), "calamine".to_owned()],
    );
    emit_institution_stage(
        report,
        on_event,
        "plan_build",
        "started",
        format!("正在按映射规则检查工作表：{file_name}。"),
        &progress,
        &file_name,
        None,
        vec!["Rust".to_owned(), "calamine".to_owned()],
        None,
        None,
        None,
        None,
    );
    let plan_result_result = build_institution_normalize_plan(
        input,
        candidate,
        profile,
        format!("{run_id}-{index}"),
        mapping,
    );
    timings.plan_build_ms = plan_started.elapsed().as_millis();
    let plan_result = match plan_result_result {
        Ok(plan_result) => plan_result,
        Err(error) => {
            let failure = normalize_file_failure(
                "plan_build",
                "planning_failed",
                format!("build workbook edit plan: {error:#}"),
                None,
            )
            .attach_state(&timings, &stage_components, actual_engine.as_deref());
            emit_institution_stage(
                report,
                on_event,
                "plan_build",
                "failed",
                format!("映射规则处理失败：{file_name}。"),
                &progress,
                &file_name,
                Some(timings.plan_build_ms),
                vec!["Rust".to_owned(), "calamine".to_owned()],
                None,
                Some(failure.error_type.to_owned()),
                Some(failure.diagnostic.clone()),
                None,
            );
            return Err(failure);
        }
    };
    emit_institution_stage(
        report,
        on_event,
        "plan_build",
        "completed",
        format!(
            "映射规则检查完成：{file_name}，命中 {} 行。",
            plan_result.mapped_rows
        ),
        &progress,
        &file_name,
        Some(timings.plan_build_ms),
        vec!["Rust".to_owned(), "calamine".to_owned()],
        None,
        None,
        None,
        None,
    );

    let Some(plan) = plan_result.plan else {
        return Ok(InstitutionNormalizeFileResult {
            file_name: file_name.clone(),
            status: "no_changes".to_owned(),
            branch_sheets: plan_result.branch_sheets,
            mapped_sheets: 0,
            mapped_rows: 0,
            operation_count: 0,
            plan_path: None,
            actual_engine: None,
            stage_components,
            execution: None,
            timings,
            error_stage: None,
            error_type: None,
            user_message: format!("完成 {index}/{total_files}：{file_name}，无需修改。"),
            diagnostic: None,
            route_diagnostic: None,
            cleanup_warning: None,
        });
    };

    let operation_count = plan.operations.len();
    let plan_path = if let Some(plan_dir) = options.plan_dir.as_deref() {
        let path = plan_dir.join(format!("{index:04}_institution_normalize.json"));
        stage_components.insert(
            "plan_persistence".to_owned(),
            vec!["Rust".to_owned(), "serde_json".to_owned()],
        );
        emit_institution_stage(
            report,
            on_event,
            "plan_persistence",
            "started",
            format!("正在保存修改计划：{file_name}。"),
            &progress,
            &file_name,
            None,
            vec!["Rust".to_owned(), "serde_json".to_owned()],
            None,
            None,
            None,
            None,
        );
        let write_started = Instant::now();
        let write_result = write_json_create_new(&path, &plan);
        timings.plan_persistence_ms = write_started.elapsed().as_millis();
        if let Err(error) = write_result {
            let failure = normalize_file_failure(
                "plan_persistence",
                "plan_output_failed",
                format!("write reviewable plan: {error:#}"),
                None,
            )
            .attach_state(&timings, &stage_components, actual_engine.as_deref());
            emit_institution_stage(
                report,
                on_event,
                "plan_persistence",
                "failed",
                format!("修改计划保存失败：{file_name}。"),
                &progress,
                &file_name,
                Some(timings.plan_persistence_ms),
                vec!["Rust".to_owned(), "serde_json".to_owned()],
                None,
                Some(failure.error_type.to_owned()),
                Some(failure.diagnostic.clone()),
                None,
            );
            return Err(failure);
        }
        emit_institution_stage(
            report,
            on_event,
            "plan_persistence",
            "completed",
            format!("修改计划已保存：{file_name}。"),
            &progress,
            &file_name,
            Some(timings.plan_persistence_ms),
            vec!["Rust".to_owned(), "serde_json".to_owned()],
            None,
            None,
            None,
            None,
        );
        Some(path)
    } else {
        None
    };

    if !options.apply {
        return Ok(InstitutionNormalizeFileResult {
            file_name: file_name.clone(),
            status: "planned".to_owned(),
            branch_sheets: plan_result.branch_sheets,
            mapped_sheets: plan_result.mapped_sheets,
            mapped_rows: plan_result.mapped_rows,
            operation_count,
            plan_path,
            actual_engine: None,
            stage_components,
            execution: None,
            timings,
            error_stage: None,
            error_type: None,
            user_message: format!(
                "已生成 {index}/{total_files}：{file_name} 的修改计划，涉及 {} 行；尚未修改原文件。",
                plan_result.mapped_rows
            ),
            diagnostic: None,
            route_diagnostic: None,
            cleanup_warning: None,
        });
    }

    let source_check_started = Instant::now();
    let expected_source_hash = source_hash.expect("apply mode computes source hash");
    stage_components.insert(
        "source_stability_check".to_owned(),
        vec!["Rust".to_owned(), "sha2".to_owned()],
    );
    emit_institution_stage(
        report,
        on_event,
        "source_stability_check",
        "started",
        format!("正在确认源文件未变化：{file_name}。"),
        &progress,
        &file_name,
        None,
        vec!["Rust".to_owned(), "sha2".to_owned()],
        None,
        None,
        None,
        None,
    );
    let current_hash_result = crate::file_commit::sha256_file(input);
    timings.source_stability_check_ms = source_check_started.elapsed().as_millis();
    let current_source_hash = match current_hash_result {
        Ok(hash) => hash,
        Err(error) => {
            let failure = normalize_file_failure(
                "source_stability_check",
                "input_unavailable",
                format!("recheck input workbook: {error:#}"),
                None,
            )
            .attach_state(&timings, &stage_components, actual_engine.as_deref());
            emit_institution_stage(
                report,
                on_event,
                "source_stability_check",
                "failed",
                format!("源文件一致性检查失败：{file_name}。"),
                &progress,
                &file_name,
                Some(timings.source_stability_check_ms),
                vec!["Rust".to_owned(), "sha2".to_owned()],
                None,
                Some(failure.error_type.to_owned()),
                Some(failure.diagnostic.clone()),
                None,
            );
            return Err(failure);
        }
    };
    if current_source_hash != expected_source_hash {
        let failure = normalize_file_failure(
            "source_stability_check",
            "input_changed",
            "input workbook changed while the modification plan was being built",
            None,
        )
        .attach_state(&timings, &stage_components, actual_engine.as_deref());
        emit_institution_stage(
            report,
            on_event,
            "source_stability_check",
            "failed",
            format!("源文件在计划期间发生变化，已停止修改：{file_name}。"),
            &progress,
            &file_name,
            Some(timings.source_stability_check_ms),
            vec!["Rust".to_owned(), "sha2".to_owned()],
            None,
            Some(failure.error_type.to_owned()),
            Some(failure.diagnostic.clone()),
            None,
        );
        return Err(failure);
    }
    emit_institution_stage(
        report,
        on_event,
        "source_stability_check",
        "completed",
        format!("源文件一致性检查通过：{file_name}。"),
        &progress,
        &file_name,
        Some(timings.source_stability_check_ms),
        vec!["Rust".to_owned(), "sha2".to_owned()],
        None,
        None,
        None,
        None,
    );

    let capability_file = options.capabilities.as_deref().expect("validated options");
    let platform = options.platform.as_deref().expect("validated options");
    let apply_started = Instant::now();
    stage_components.insert(
        "engine_execution".to_owned(),
        vec!["Rust".to_owned(), "ExcelModifier".to_owned()],
    );
    emit_institution_stage(
        report,
        on_event,
        "engine_execution",
        "started",
        format!("正在按已登记能力修改工作簿：{file_name}。"),
        &progress,
        &file_name,
        None,
        vec!["Rust".to_owned(), "ExcelModifier".to_owned()],
        None,
        None,
        None,
        None,
    );
    let execution_result =
        crate::edit_plan::apply_plan_with_capability_file(&plan, capability_file, platform);
    timings.engine_execution_ms = apply_started.elapsed().as_millis();
    let execution = execution_result.map_err(|error| {
        let modifier_error = error.downcast_ref::<excel_editor::ExcelModifierError>();
        let route_diagnostic = modifier_error.and_then(|error| serde_json::to_value(error).ok());
        let error_type = match modifier_error.map(|error| error.kind) {
            Some(
                excel_editor::ExcelModifierErrorKind::NoEligibleBackend
                | excel_editor::ExcelModifierErrorKind::AllBackendsRejectedBeforeWrite,
            ) => "no_eligible_backend",
            _ => "backend_execution_failed",
        };
        let failure = normalize_file_failure(
            "engine_execution",
            error_type,
            format!("ExcelModifier execution failed: {error:#}"),
            route_diagnostic,
        )
        .attach_state(&timings, &stage_components, actual_engine.as_deref());
        emit_institution_stage(
            report,
            on_event,
            "engine_execution",
            "failed",
            format!("工作簿修改失败：{file_name}；原文件未安全替换。"),
            &progress,
            &file_name,
            Some(timings.engine_execution_ms),
            vec!["Rust".to_owned(), "ExcelModifier".to_owned()],
            None,
            Some(failure.error_type.to_owned()),
            Some(failure.diagnostic.clone()),
            failure.route_diagnostic.clone(),
        );
        failure
    })?;
    actual_engine = Some(execution.engine.to_string());
    let execution_value = serde_json::to_value(&execution).ok();
    let mut execution_components = vec!["Rust".to_owned(), "ExcelModifier".to_owned()];
    execution_components.push(execution.engine.to_string());
    execution_components.push(engine_language(&execution.engine.to_string()).to_owned());
    stage_components.insert("engine_execution".to_owned(), execution_components);
    emit_institution_stage(
        report,
        on_event,
        "engine_execution",
        "completed",
        format!("工作簿修改引擎已执行：{file_name}。"),
        &progress,
        &file_name,
        Some(timings.engine_execution_ms),
        vec![
            "Rust".to_owned(),
            "ExcelModifier".to_owned(),
            execution.engine.to_string(),
            engine_language(&execution.engine.to_string()).to_owned(),
        ],
        actual_engine.clone(),
        None,
        None,
        None,
    );

    let commit_started = Instant::now();
    stage_components.insert(
        "safe_commit".to_owned(),
        vec!["Rust".to_owned(), "file_commit".to_owned()],
    );
    emit_institution_stage(
        report,
        on_event,
        "safe_commit",
        "started",
        format!("正在验证修改结果并安全提交：{file_name}。"),
        &progress,
        &file_name,
        None,
        vec!["Rust".to_owned(), "file_commit".to_owned()],
        actual_engine.clone(),
        None,
        None,
        None,
    );
    let commit_result =
        crate::file_commit::replace_target_from_candidate(candidate, input, &expected_source_hash);
    timings.safe_commit_ms = commit_started.elapsed().as_millis();
    let commit = commit_result.map_err(|error| {
        let mut failure = normalize_file_failure(
            "safe_commit",
            "commit_failed",
            format!("safe target replacement failed: {error:#}"),
            None,
        )
        .attach_state(&timings, &stage_components, actual_engine.as_deref())
        .attach_execution(execution_value.clone());
        if let Err(cleanup_error) = fs::remove_file(candidate) {
            if cleanup_error.kind() != std::io::ErrorKind::NotFound {
                failure.diagnostic.push_str(&format!(
                    "; validated candidate cleanup failed: {cleanup_error}"
                ));
            }
        }
        emit_institution_stage(
            report,
            on_event,
            "safe_commit",
            "failed",
            format!("安全提交失败：{file_name}；请检查原文件和备份。"),
            &progress,
            &file_name,
            Some(timings.safe_commit_ms),
            vec!["Rust".to_owned(), "file_commit".to_owned()],
            actual_engine.clone(),
            Some(failure.error_type.to_owned()),
            Some(failure.diagnostic.clone()),
            None,
        );
        failure
    })?;
    emit_institution_stage(
        report,
        on_event,
        "safe_commit",
        "completed",
        format!("修改结果已通过验证并安全提交：{file_name}。"),
        &progress,
        &file_name,
        Some(timings.safe_commit_ms),
        vec!["Rust".to_owned(), "file_commit".to_owned()],
        actual_engine.clone(),
        None,
        None,
        None,
    );
    let mut cleanup_warning = commit.cleanup_warning;

    Ok(InstitutionNormalizeFileResult {
        file_name: file_name.clone(),
        status: "applied".to_owned(),
        branch_sheets: plan_result.branch_sheets,
        mapped_sheets: plan_result.mapped_sheets,
        mapped_rows: plan_result.mapped_rows,
        operation_count,
        plan_path,
        actual_engine,
        stage_components,
        execution: execution_value,
        timings,
        error_stage: None,
        error_type: None,
        user_message: format!(
            "完成 {index}/{total_files}：{file_name}，修改 {} 行。",
            plan_result.mapped_rows
        ),
        diagnostic: None,
        route_diagnostic: None,
        cleanup_warning: cleanup_warning.take(),
    })
}

fn emit_institution_normalize_event<F>(
    report: &mut InstitutionNormalizeRunReport,
    on_event: &mut F,
    phase: &str,
    status: &str,
    user_message: String,
    progress: InstitutionNormalizeProgress,
    file_name: Option<String>,
    elapsed_ms: Option<u128>,
    components: Vec<String>,
    actual_engine: Option<String>,
    error_type: Option<String>,
    diagnostic: Option<String>,
    route_diagnostic: Option<Value>,
) where
    F: FnMut(&InstitutionNormalizeEvent),
{
    let event = InstitutionNormalizeEvent {
        sequence: report.events.len() + 1,
        timestamp: chrono::Utc::now().to_rfc3339(),
        phase: phase.to_owned(),
        status: status.to_owned(),
        user_message,
        progress,
        file_name,
        elapsed_ms,
        components,
        actual_engine,
        error_type,
        diagnostic,
        route_diagnostic,
    };
    on_event(&event);
    report.events.push(event);
}

#[allow(clippy::too_many_arguments)]
fn emit_institution_stage<F>(
    report: &mut InstitutionNormalizeRunReport,
    on_event: &mut F,
    phase: &str,
    status: &str,
    message: String,
    progress: &InstitutionNormalizeProgress,
    file_name: &str,
    elapsed_ms: Option<u128>,
    components: Vec<String>,
    actual_engine: Option<String>,
    error_type: Option<String>,
    diagnostic: Option<String>,
    route_diagnostic: Option<Value>,
) where
    F: FnMut(&InstitutionNormalizeEvent),
{
    emit_institution_normalize_event(
        report,
        on_event,
        phase,
        status,
        message,
        progress.clone(),
        Some(file_name.to_owned()),
        elapsed_ms,
        components,
        actual_engine,
        error_type,
        diagnostic,
        route_diagnostic,
    );
}

fn current_institution_progress(
    report: &InstitutionNormalizeRunReport,
    total_files: usize,
) -> InstitutionNormalizeProgress {
    InstitutionNormalizeProgress {
        total_files,
        current_file: None,
        processed_files: report.completed_files + report.failed_files,
        completed_files: report.completed_files,
        failed_files: report.failed_files,
        modified_files: report.modified_files,
    }
}

fn institution_normalize_candidate(input: &Path, run_id: &str, index: usize) -> Result<PathBuf> {
    let parent = input
        .parent()
        .filter(|path| !path.as_os_str().is_empty())
        .unwrap_or(Path::new("."));
    for attempt in 0..1000usize {
        let candidate = parent.join(format!(
            ".pytools-institution-normalize-{run_id}-{index}-{attempt}.xlsx"
        ));
        if !candidate.exists() {
            return Ok(candidate);
        }
    }
    bail!(
        "could not reserve a unique PT-7-2 candidate beside {}",
        input.display()
    )
}

fn write_json_create_new<T: Serialize>(path: &Path, value: &T) -> Result<()> {
    if let Some(parent) = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
    {
        fs::create_dir_all(parent)
            .with_context(|| format!("create plan directory {}", parent.display()))?;
    }
    let mut file = OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(path)
        .with_context(|| format!("create plan file {} without overwrite", path.display()))?;
    let result = (|| -> Result<()> {
        file.write_all(&serde_json::to_vec_pretty(value)?)
            .with_context(|| format!("write plan file {}", path.display()))?;
        file.sync_all()
            .with_context(|| format!("sync plan file {}", path.display()))?;
        Ok(())
    })();
    drop(file);
    if let Err(error) = result {
        let _ = fs::remove_file(path);
        return Err(error);
    }
    Ok(())
}

fn normalize_file_failure(
    stage: &'static str,
    error_type: &'static str,
    diagnostic: impl Into<String>,
    route_diagnostic: Option<Value>,
) -> InstitutionNormalizeFileFailure {
    InstitutionNormalizeFileFailure {
        stage,
        error_type,
        diagnostic: diagnostic.into(),
        route_diagnostic,
        actual_engine: None,
        execution: None,
        timings: InstitutionNormalizeStageTimings::default(),
        stage_components: BTreeMap::new(),
    }
}

fn user_message_for_normalize_failure(
    error_type: &str,
    file_name: &str,
    index: usize,
    total: usize,
) -> String {
    match error_type {
        "unsupported_format" => {
            format!("跳过 {index}/{total}：{file_name}，当前候选仅支持 .xlsx；原文件未修改。")
        }
        "no_eligible_backend" => format!(
            "失败 {index}/{total}：{file_name}，没有已验收的修改引擎可处理此文件；原文件未修改。"
        ),
        "backend_execution_failed" => {
            format!("失败 {index}/{total}：{file_name}，工作簿修改未通过验证；原文件未安全替换。")
        }
        _ => format!("失败 {index}/{total}：{file_name}，该文件未完成安全处理；请查看开发诊断。"),
    }
}

fn engine_language(engine: &str) -> &'static str {
    match engine {
        "openpyxl_worker" => "Python",
        "excelize_worker" => "Go",
        _ => "Rust",
    }
}

/// Normalize one in-memory worksheet matrix, preserving mapping insertion order.
///
/// Rows and columns use one-based Excel semantics at the API boundary: row 1 is
/// the header, and column A is preferred whenever it has data; otherwise B is
/// used. The supplied mapping is ordered because substring matches use the
/// first matching key after exact-match lookup.
pub fn normalize_branch_sheet_values(
    sheet_name: &str,
    rows: &mut [Vec<Option<String>>],
    mapping: &[(String, String)],
) -> usize {
    if !sheet_name.contains(BRANCH_SHEET_KEY) || mapping.is_empty() {
        return 0;
    }

    let last_a = last_nonblank_row(rows, 0);
    let last_b = last_nonblank_row(rows, 1);
    let last_row = last_a.max(last_b);
    if last_row < 2 {
        return 0;
    }

    let Some(target_column) = target_column(rows, last_row) else {
        return 0;
    };

    let mut changed = 0;
    for row in rows.iter_mut().take(last_row).skip(1) {
        let Some(value) = row.get_mut(target_column) else {
            continue;
        };
        let Some(raw) = value.as_deref() else {
            continue;
        };
        let original = raw.trim();
        if original.is_empty() {
            continue;
        }
        let normalized = normalize_name(original, mapping);
        if normalized != original {
            *value = Some(normalized);
            changed += 1;
        }
    }
    changed
}

/// Build an engine-independent cell-update plan for one `.xlsx` workbook.
///
/// `workbook_profile` must come from a complete, independently validated
/// package inventory before the plan is executed. This function preserves the
/// caller's inventory status; it does not claim that reading cell values is a
/// complete OOXML feature inventory. The existing umya candidate remains
/// The old direct umya candidate is compiled only into this module's tests;
/// production callers must route the resulting plan through ExcelModifier.
pub fn build_institution_normalize_plan(
    input: impl AsRef<Path>,
    output: impl AsRef<Path>,
    workbook_profile: WorkbookProfile,
    request_id: impl Into<String>,
    mapping: &[(String, String)],
) -> Result<NormalizePlanResult> {
    if mapping.is_empty() {
        return Ok(NormalizePlanResult {
            mapping_empty: true,
            ..NormalizePlanResult::default()
        });
    }

    let input = input.as_ref();
    let output = output.as_ref();
    if !has_extension(input, "xlsx") || !has_extension(output, "xlsx") {
        return Err(anyhow!(
            "institution normalization edit plans currently accept only .xlsx"
        ));
    }
    if workbook_profile.format != WorkbookFormat::Xlsx {
        return Err(anyhow!(
            "institution normalization edit plans require an XLSX workbook profile"
        ));
    }

    let dump = read_workbook(input).with_context(|| format!("read {}", input.display()))?;
    let mut result = NormalizePlanResult::default();
    let mut operations = Vec::new();

    for sheet in dump
        .sheets
        .iter()
        .filter(|sheet| sheet.name.contains(BRANCH_SHEET_KEY))
    {
        result.branch_sheets += 1;
        if sheet.formula_status.starts_with("read_error:") {
            return Err(anyhow!(
                "cannot plan institution normalization because formula text could not be read for worksheet {}: {}",
                sheet.name,
                sheet.formula_status
            ));
        }

        let max_row = sheet
            .values
            .keys()
            .chain(sheet.formulas.keys())
            .filter_map(|address| target_column_row(address))
            .map(|(_, row)| row)
            .max()
            .unwrap_or_default();
        if max_row < 2 {
            continue;
        }

        let mut rows = vec![vec![None, None]; max_row];
        for row in 1..=max_row {
            for column in 0..2 {
                let address = target_cell_address(column, row);
                rows[row - 1][column] = cell_python_string_from_dump(sheet, &address)?;
            }
        }
        let original_rows = rows.clone();
        let changed_rows = normalize_branch_sheet_values(&sheet.name, &mut rows, mapping);
        if changed_rows == 0 {
            continue;
        }

        result.mapped_sheets += 1;
        result.mapped_rows += changed_rows;
        for row_index in 1..rows.len() {
            for column_index in 0..2 {
                if rows[row_index][column_index] == original_rows[row_index][column_index] {
                    continue;
                }
                let value = rows[row_index][column_index]
                    .as_ref()
                    .expect("normalization only writes nonblank mapped values");
                operations.push(EditOperation::SetCellValue {
                    sheet: sheet.name.clone(),
                    cell: target_cell_address(column_index, row_index + 1),
                    value: PlannedCellValue::Text(value.clone()),
                });
            }
        }
    }

    if !operations.is_empty() {
        // openpyxl rewrites known package XML (including core properties), so
        // PT-7-2 asks for explicit semantic object preservation instead of the
        // raw UnmodifiedPackageParts contract. Unknown package parts remain a
        // pre-write rejection at the backend boundary.
        let preserve = BTreeSet::from([
            PreservationRequirement::CellValuesAndTypes,
            PreservationRequirement::FormulaText,
            PreservationRequirement::CommentsAndPositions,
            PreservationRequirement::Styles,
            PreservationRequirement::MergedRanges,
            PreservationRequirement::DefinedNames,
            PreservationRequirement::RowAndColumnDimensions,
            PreservationRequirement::PageSettings,
            PreservationRequirement::HeaderFooterText,
            PreservationRequirement::Relationships,
            PreservationRequirement::SheetIdentityAndOrder,
        ]);
        let mut preserve = preserve;
        if workbook_profile
            .features_present
            .contains(&excel_editor::WorkbookFeature::ConditionalFormatting)
        {
            preserve.insert(PreservationRequirement::ConditionalFormatting);
        }
        if workbook_profile
            .features_present
            .contains(&excel_editor::WorkbookFeature::DataValidations)
        {
            preserve.insert(PreservationRequirement::DataValidations);
        }
        let plan = ExcelEditPlan {
            request_id: request_id.into(),
            input: input.to_path_buf(),
            output: output.to_path_buf(),
            format: WorkbookFormat::Xlsx,
            workbook: workbook_profile,
            operations,
            preserve,
            drop_allowed: Default::default(),
        };
        plan.validate()
            .context("validate institution normalization edit plan")?;
        result.plan = Some(plan);
    }

    Ok(result)
}

fn target_column_row(address: &str) -> Option<(usize, usize)> {
    let first_digit = address.find(|character: char| character.is_ascii_digit())?;
    let column = match &address[..first_digit] {
        "A" => 0,
        "B" => 1,
        _ => return None,
    };
    let row = address[first_digit..].parse::<usize>().ok()?;
    (row > 0).then_some((column, row))
}

fn target_cell_address(column: usize, row: usize) -> String {
    format!("{}{}", if column == 0 { "A" } else { "B" }, row)
}

fn has_extension(path: &Path, expected: &str) -> bool {
    path.extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| extension.eq_ignore_ascii_case(expected))
}

fn cell_python_string_from_dump(sheet: &SheetDump, address: &str) -> Result<Option<String>> {
    if let Some(formula) = sheet.formulas.get(address) {
        return Ok(Some(format!("={formula}")));
    }
    let Some(value) = sheet.values.get(address) else {
        return Ok(None);
    };
    let value = ReaderCellValue::from_probe_json(value)
        .with_context(|| format!("decode {address} on worksheet {}", sheet.name))?;
    Ok(match value {
        ReaderCellValue::Empty => None,
        ReaderCellValue::Text(value)
        | ReaderCellValue::DateTimeIso(value)
        | ReaderCellValue::DurationIso(value) => Some(value),
        ReaderCellValue::Number(value) | ReaderCellValue::DateTimeSerial(value) => {
            Some(value.to_string())
        }
        ReaderCellValue::Integer(value) => Some(value.to_string()),
        ReaderCellValue::Boolean(value) => Some(if value { "True" } else { "False" }.to_owned()),
        ReaderCellValue::Error(value) => {
            return Err(anyhow!(
                "cannot plan institution normalization for error-valued cell {address} on worksheet {}: {value}",
                sheet.name
            ));
        }
    })
}

/// Process `.xlsx` workbooks in place. Inputs are processed concurrently when
/// their paths are unique; repeated paths are processed sequentially to avoid
/// racing writes to the same file. This legacy implementation is test-only so
/// production code cannot bypass ExcelModifier.
#[cfg(test)]
pub fn run_normalize_institution(
    target_workbooks: &[PathBuf],
    mapping: &[(String, String)],
) -> NormalizeReport {
    if mapping.is_empty() {
        return NormalizeReport {
            mapping_empty: true,
            ..NormalizeReport::default()
        };
    }

    let mut seen = HashSet::new();
    let unique_paths = target_workbooks.iter().all(|path| {
        let key = std::fs::canonicalize(path).unwrap_or_else(|_| path.clone());
        seen.insert(key)
    });
    let results = if unique_paths {
        target_workbooks
            .par_iter()
            .map(|path| (path, process_workbook(path, mapping)))
            .collect::<Vec<_>>()
    } else {
        target_workbooks
            .iter()
            .map(|path| (path, process_workbook(path, mapping)))
            .collect::<Vec<_>>()
    };

    let mut report = NormalizeReport::default();
    for (path, result) in results {
        match result {
            Ok((branch_sheets, mapped_sheets, mapped_rows)) => {
                report.files += 1;
                report.branch_sheets += branch_sheets;
                report.mapped_sheets += mapped_sheets;
                report.mapped_rows += mapped_rows;
            }
            Err(error) => {
                report.failed_files += 1;
                report.errors.push(format!("{}: {error:#}", path.display()));
            }
        }
    }
    report
}

#[cfg(test)]
fn process_workbook(path: &Path, mapping: &[(String, String)]) -> Result<(usize, usize, usize)> {
    if !path
        .extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
    {
        return Err(anyhow!(
            "当前机构名称规范化候选后端只接受 .xlsx，未验证其他格式"
        ));
    }

    let mut workbook = open_xlsx(path).with_context(|| format!("open {}", path.display()))?;
    let sheet_names = workbook
        .sheet_collection()
        .iter()
        .filter(|sheet| sheet.name().contains(BRANCH_SHEET_KEY))
        .map(|sheet| sheet.name().to_owned())
        .collect::<Vec<_>>();

    let mut branch_sheets = 0;
    let mut mapped_sheets = 0;
    let mut mapped_rows = 0;
    for sheet_name in sheet_names {
        branch_sheets += 1;
        let sheet = workbook
            .sheet_by_name_mut(&sheet_name)
            .with_context(|| format!("find worksheet {sheet_name}"))?;
        let original_column_numbers = sheet
            .column_dimensions()
            .iter()
            .map(|column| column.col_num())
            .collect::<HashSet<_>>();
        let max_row = sheet.highest_row() as usize;
        let mut values = (0..max_row)
            .map(|row_index| {
                (1_u32..=2)
                    .map(|column_index| {
                        sheet
                            .cell((column_index, row_index as u32 + 1))
                            .map(cell_python_string)
                    })
                    .collect::<Vec<_>>()
            })
            .collect::<Vec<_>>();
        let changed = normalize_branch_sheet_values(&sheet_name, &mut values, mapping);
        if changed == 0 {
            continue;
        }

        let target_column = target_column(&values, values.len()).unwrap();
        for (row_index, row) in values.iter().enumerate().skip(1) {
            let Some(value) = row.get(target_column).and_then(Option::as_ref) else {
                continue;
            };
            let original = sheet
                .cell((target_column as u32 + 1, row_index as u32 + 1))
                .map(cell_python_string)
                .unwrap_or_default();
            if *value != original {
                sheet
                    .cell_mut((target_column as u32 + 1, row_index as u32 + 1))
                    .set_value_string(value.clone());
            }
        }
        sheet
            .column_dimensions_mut()
            .retain(|column| original_column_numbers.contains(&column.col_num()));
        mapped_sheets += 1;
        mapped_rows += changed;
    }

    if mapped_rows > 0 {
        save_xlsx(&workbook, path).with_context(|| format!("save {}", path.display()))?;
    }
    Ok((branch_sheets, mapped_sheets, mapped_rows))
}

#[cfg(test)]
fn cell_python_string(cell: &excel_editor::Cell) -> String {
    if cell.is_formula() {
        format!("={}", cell.formula())
    } else if cell.data_type() == "b" {
        match cell.value().as_ref() {
            "TRUE" => "True".to_owned(),
            "FALSE" => "False".to_owned(),
            value => value.to_owned(),
        }
    } else {
        cell.value().into_owned()
    }
}

fn last_nonblank_row(rows: &[Vec<Option<String>>], column: usize) -> usize {
    rows.iter()
        .enumerate()
        .rev()
        .find(|(_, row)| {
            row.get(column)
                .and_then(Option::as_deref)
                .is_some_and(|value| !value.trim().is_empty())
        })
        .map_or(0, |(index, _)| index + 1)
}

fn target_column(rows: &[Vec<Option<String>>], last_row: usize) -> Option<usize> {
    let has_data = |column: usize| {
        rows.iter().take(last_row).skip(1).any(|row| {
            row.get(column)
                .and_then(Option::as_deref)
                .is_some_and(|value| !value.trim().is_empty())
        })
    };
    if has_data(0) {
        Some(0)
    } else if has_data(1) {
        Some(1)
    } else {
        None
    }
}

fn normalize_name(original: &str, mapping: &[(String, String)]) -> String {
    if let Some((_, standard)) = mapping.iter().find(|(key, _)| key == original) {
        return standard.clone();
    }
    let original_lower = original.to_lowercase();
    mapping
        .iter()
        .find(|(key, _)| !key.is_empty() && original_lower.contains(&key.to_lowercase()))
        .map_or_else(|| original.to_owned(), |(_, standard)| standard.clone())
}

#[cfg(test)]
mod tests {
    use super::*;
    use excel_reader::read_workbook;
    use serde_json::{json, Value};
    use std::time::{SystemTime, UNIX_EPOCH};

    #[test]
    fn workbook_normalization_matches_python_baseline_and_preserves_fixture_metadata() {
        let fixture_dir = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures");
        let baseline: Value = serde_json::from_str(include_str!(
            "../probe/reports/institution_normalize_baseline.json"
        ))
        .expect("Python baseline report must be valid JSON");
        let mapping = baseline["mapping"]
            .as_array()
            .unwrap()
            .iter()
            .map(|entry| {
                (
                    entry[0].as_str().unwrap().to_owned(),
                    entry[1].as_str().unwrap().to_owned(),
                )
            })
            .collect::<Vec<_>>();
        let source = fixture_dir.join("institution_normalize.xlsx");
        let unique = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let edited = std::env::temp_dir().join(format!(
            "pytools-institution-normalize-{}-{unique}.xlsx",
            std::process::id()
        ));
        std::fs::copy(&source, &edited).unwrap();

        let report = run_normalize_institution(std::slice::from_ref(&edited), &mapping);
        assert_eq!(
            report.files,
            baseline["stats"]["files"].as_u64().unwrap() as usize
        );
        assert_eq!(
            report.branch_sheets,
            baseline["stats"]["branch_sheets"].as_u64().unwrap() as usize
        );
        assert_eq!(
            report.mapped_sheets,
            baseline["stats"]["mapped_sheets"].as_u64().unwrap() as usize
        );
        assert_eq!(
            report.mapped_rows,
            baseline["stats"]["mapped_rows"].as_u64().unwrap() as usize
        );
        assert_eq!(report.failed_files, 0, "{:?}", report.errors);

        let dump = read_workbook(&edited).unwrap();
        let actual = dump
            .sheets
            .iter()
            .map(|sheet| {
                json!({
                    "name": sheet.name,
                    "values": sheet.values,
                    "formulas": sheet.formulas,
                    "merges": sheet.merges,
                    "comments": sheet.comments,
                })
            })
            .collect::<Vec<_>>();
        let expected = baseline["sheets"]
            .as_array()
            .unwrap()
            .iter()
            .map(|sheet| {
                json!({
                    "name": sheet["name"],
                    "values": sheet["values"],
                    "formulas": sheet["formulas"],
                    "merges": sheet["merges"],
                    "comments": sheet["comments"],
                })
            })
            .collect::<Vec<_>>();
        assert_eq!(json!(actual), json!(expected));
        std::fs::remove_file(edited).unwrap();
    }

    #[test]
    fn helper_uses_exact_match_then_first_case_insensitive_substring() {
        let mapping = vec![
            ("AB".to_owned(), "Short Standard".to_owned()),
            ("ABC Bank".to_owned(), "Exact Standard".to_owned()),
            ("branch".to_owned(), "Branch Standard".to_owned()),
        ];
        let mut rows = vec![
            vec![Some("机构".to_owned()), Some("名称".to_owned())],
            vec![Some("ABC Bank".to_owned()), Some("ignored".to_owned())],
            vec![Some(" abc bank office ".to_owned()), None],
            vec![Some("   ".to_owned()), None],
        ];
        assert_eq!(
            normalize_branch_sheet_values("报表分机构", &mut rows, &mapping),
            2
        );
        assert_eq!(rows[1][0].as_deref(), Some("Exact Standard"));
        assert_eq!(rows[2][0].as_deref(), Some("Short Standard"));

        let mut fallback_rows = vec![
            vec![Some("机构".to_owned()), Some("名称".to_owned())],
            vec![None, Some("Branch Office".to_owned())],
        ];
        assert_eq!(
            normalize_branch_sheet_values("分机构-备用", &mut fallback_rows, &mapping),
            1
        );
        assert_eq!(fallback_rows[1][1].as_deref(), Some("Branch Standard"));
        assert_eq!(
            normalize_branch_sheet_values("普通表", &mut rows, &mapping),
            0
        );
    }

    #[test]
    fn empty_mapping_skips_all_files() {
        let missing = PathBuf::from("missing.xlsx");
        let report = run_normalize_institution(&[missing], &[]);

        assert!(report.mapping_empty);
        assert_eq!(report.files, 0);
        assert_eq!(report.failed_files, 0);
    }

    #[test]
    fn repeated_paths_are_processed_sequentially_without_double_mapping() {
        let fixture =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/institution_normalize.xlsx");
        let unique = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let edited = std::env::temp_dir().join(format!(
            "pytools-institution-normalize-duplicate-{}-{unique}.xlsx",
            std::process::id()
        ));
        std::fs::copy(fixture, &edited).unwrap();
        let mapping = vec![("ABC Bank".to_owned(), "Exact Standard".to_owned())];

        let report = run_normalize_institution(&[edited.clone(), edited.clone()], &mapping);
        assert_eq!(report.files, 2);
        assert_eq!(report.branch_sheets, 6);
        assert_eq!(report.mapped_sheets, 1);
        assert_eq!(report.mapped_rows, 2);
        assert_eq!(report.failed_files, 0, "{:?}", report.errors);
        std::fs::remove_file(edited).unwrap();
    }
}
