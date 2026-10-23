use anyhow::{Context, Result};
use audit_core::{
    conditional_format::evaluate_xlsx_workbook,
    discovery::{
        detect_period_from_workbooks, organisation_from_name, source_workbooks, ScanDepth,
    },
    models::CopyRange,
    name_config::{
        load_s1_conditional_mappings, matches_named_range, named_areas_from_defined_names,
        resolve_s1_conditional_area_plan, CONDITIONAL_FORMAT_EXTRACT_FUNCTION,
        DEFAULT_CONDITIONAL_FEATURE_NAME,
    },
    s1_conditional_audit::{
        extract_conditional_issues, extract_conditional_issues_batch,
        S1ConditionalBatchAuditResult, S1ConditionalBatchInput, S1ConditionalBatchProgress,
        S1ConditionalMetadata,
    },
    s1_structure_compare::{
        compare_s1_source_structure, read_s1_structure_template_with_required_sheets,
    },
    s1_template::{recommend_template, recommend_template_for_workbooks, template_workbook_paths},
    s1_template_rules::load_s1_template_rule_plan,
    s2_service::{run_s2_audit, S2ProgressEvent, S2RunRequest},
    s2_summary::S2AuditSummary,
    web_bridge::serve_web_ui,
};
use clap::{Args, Parser, Subcommand};
use rayon::prelude::*;
use serde::{Deserialize, Serialize};
use std::{
    fs,
    path::{Path, PathBuf},
    time::Instant,
};

#[derive(Debug, Parser)]
#[command(
    name = "open-data-audit",
    version,
    about = "Rust implementation of Open-Data-Audit workflows"
)]
struct Cli {
    #[command(subcommand)]
    command: Command,
}

#[derive(Debug, Subcommand)]
enum Command {
    /// Run the S2 period-report audit and create its three-sheet result workbook.
    S2(S2Args),
    /// Recommend an S1 template from source and template workbook sheet names.
    TemplateSuggest(TemplateSuggestArgs),
    /// Evaluate the Rust-supported static conditional-format rules in an XLSX workbook.
    ConditionalEvaluate(ConditionalEvaluateArgs),
    /// Extract static conditional-format results into S1-F01 Audit issue records.
    ConditionalAudit(ConditionalAuditArgs),
    /// Run conditional-format issue extraction for a selected S1 workbook batch.
    ConditionalAuditBatch(ConditionalAuditBatchArgs),
    /// Discover an S1 input directory and run conditional-format issue extraction.
    ConditionalAuditDirectory(ConditionalAuditDirectoryArgs),
    /// Discover and validate an S1 template's formula rules without recalculation.
    S1TemplateRules(S1TemplateRulesArgs),
    /// Serve the bundled Audit frontend through the Rust API bridge on localhost.
    Serve(ServeArgs),
}

#[derive(Debug, Args)]
struct ConditionalEvaluateArgs {
    /// Input .xlsx/.xlsm workbook. Formula cells are read from stored caches; this command does not recalculate.
    #[arg(long)]
    workbook: PathBuf,

    /// Optional scan range in Sheet!A1:B20 form. Repeat for multiple ranges. Defaults to rule AppliesTo ranges.
    #[arg(long = "scan-area")]
    scan_areas: Vec<String>,
}

#[derive(Debug, Args)]
struct ConditionalAuditArgs {
    /// Workbook containing the values and conditional-format rules to audit.
    #[arg(long)]
    workbook: PathBuf,

    /// Original report path used for Audit issue identity. Defaults to --workbook.
    #[arg(long)]
    source_file: Option<PathBuf>,

    /// Reporting period recorded on the generated Audit issues.
    #[arg(long)]
    period: String,

    /// S1 run/batch identifier recorded on the generated Audit issues.
    #[arg(long)]
    batch_id: String,

    /// Audit timestamp recorded on the generated Audit issues.
    #[arg(long)]
    audit_time: String,

    /// Institution code when it was resolved by the calling workflow.
    #[arg(long, default_value = "")]
    org_code: String,

    /// Institution name when it was resolved by the calling workflow.
    #[arg(long, default_value = "")]
    org_name: String,

    /// Optional evaluation area in Sheet!A1:B20 form. Repeat for multiple areas.
    #[arg(long = "scan-area")]
    scan_areas: Vec<String>,

    /// Optional template structure area used to resolve check_field. Repeat as needed.
    #[arg(long = "structure-area")]
    structure_areas: Vec<String>,

    /// Template workbook whose configured defined names select condition and structure areas.
    #[arg(long)]
    template_workbook: Option<PathBuf>,

    /// Optional Audit config workbook containing the “模块化功能” mappings.
    #[arg(long)]
    mapping_config: Option<PathBuf>,

    /// Emit developer-view JSON progress events with actual engine and timings.
    #[arg(long, default_value_t = false)]
    developer_view: bool,
}

#[derive(Debug, Args)]
struct ConditionalAuditBatchArgs {
    /// JSON manifest containing selected workbooks and per-file institution metadata.
    #[arg(long)]
    manifest: PathBuf,

    /// Reporting period recorded on the generated Audit issues.
    #[arg(long)]
    period: String,

    /// S1 run/batch identifier recorded on the generated Audit issues.
    #[arg(long)]
    batch_id: String,

    /// Audit timestamp recorded on the generated Audit issues.
    #[arg(long)]
    audit_time: String,

    /// Optional evaluation area in Sheet!A1:B20 form. Repeat for multiple areas.
    #[arg(long = "scan-area")]
    scan_areas: Vec<String>,

    /// Optional template structure area used to resolve check_field. Repeat as needed.
    #[arg(long = "structure-area")]
    structure_areas: Vec<String>,

    /// Template workbook whose configured defined names select condition and structure areas.
    #[arg(long)]
    template_workbook: Option<PathBuf>,

    /// Optional Audit config workbook containing the “模块化功能” mappings.
    #[arg(long)]
    mapping_config: Option<PathBuf>,

    /// Emit developer-view JSON progress events with actual engine and timings.
    #[arg(long, default_value_t = false)]
    developer_view: bool,
}

#[derive(Debug, Args)]
struct ConditionalAuditDirectoryArgs {
    /// Directory containing source .xlsx workbooks.
    #[arg(long)]
    input_dir: PathBuf,

    /// Reporting period; when omitted, detect one from the source filenames/directories.
    #[arg(long)]
    period: Option<String>,

    /// S1 run/batch identifier recorded on generated Audit issues.
    #[arg(long)]
    batch_id: String,

    /// Audit timestamp recorded on generated Audit issues.
    #[arg(long)]
    audit_time: String,

    /// Recurse through all child directories when discovering source workbooks.
    #[arg(long, default_value_t = false)]
    recursive: bool,

    /// Restrict discovery to the selected source files; repeat the flag as needed.
    /// Relative paths are resolved from --input-dir.
    #[arg(long = "selected-file")]
    selected_files: Vec<PathBuf>,

    /// Optional evaluation area in Sheet!A1:B20 form. Repeat for multiple areas.
    #[arg(long = "scan-area")]
    scan_areas: Vec<String>,

    /// Optional template structure area used to resolve check_field. Repeat as needed.
    #[arg(long = "structure-area")]
    structure_areas: Vec<String>,

    /// Template workbook whose configured defined names select condition and structure areas.
    #[arg(long, conflicts_with = "template_dir")]
    template_workbook: Option<PathBuf>,

    /// Directory of templates used to recommend a workbook for this source directory.
    #[arg(long)]
    template_dir: Option<PathBuf>,

    /// Optional Audit config workbook containing the “模块化功能” mappings.
    #[arg(long)]
    mapping_config: Option<PathBuf>,

    /// Emit developer-view JSON progress events with actual engine and timings.
    #[arg(long, default_value_t = false)]
    developer_view: bool,
}

#[derive(Debug, Args)]
struct S1TemplateRulesArgs {
    /// S1 template workbook containing named formula ranges or an “审核规则” sheet.
    #[arg(long)]
    template_workbook: PathBuf,

    /// Optional Audit config workbook containing the “模块化功能” mappings.
    #[arg(long)]
    mapping_config: Option<PathBuf>,

    /// Include the complete rule definition in stdout for an in-process workflow consumer.
    #[arg(long, default_value_t = false)]
    include_rule_definitions: bool,

    /// Emit developer-view JSON progress events with reader and timing details.
    #[arg(long, default_value_t = false)]
    developer_view: bool,
}

#[derive(Debug, Deserialize)]
struct ConditionalAuditManifest {
    files: Vec<ConditionalAuditManifestFile>,
}

#[derive(Debug, Deserialize)]
struct ConditionalAuditManifestFile {
    workbook: PathBuf,
    #[serde(default)]
    source_file: Option<PathBuf>,
    #[serde(default)]
    org_code: String,
    #[serde(default)]
    org_name: String,
}

#[derive(Clone, Debug, Serialize)]
struct S1StructurePreflightFile {
    source_file: String,
    status: &'static str,
    matched_labels: usize,
    checked_labels: usize,
    score: f64,
    missing_sheets: Vec<String>,
    read_time_ms: Option<u128>,
    elapsed_ms: u128,
    diagnostic: Option<String>,
    error: Option<String>,
}

#[derive(Clone, Debug, Serialize)]
struct S1StructurePreflightSummary {
    status: &'static str,
    engine: &'static str,
    files_total: usize,
    files_accepted: usize,
    files_rejected: usize,
    template_areas: usize,
    template_labels: usize,
    template_read_time_ms: u128,
    elapsed_ms: u128,
    files: Vec<S1StructurePreflightFile>,
}

#[derive(Debug, Args)]
struct TemplateSuggestArgs {
    /// Directory containing .xlsx templates.
    #[arg(long)]
    template_dir: PathBuf,

    /// Directory containing source .xlsx workbooks.
    #[arg(long)]
    input_dir: PathBuf,

    /// Include files below the input directory.
    #[arg(long, default_value_t = false)]
    recursive: bool,
}

#[derive(Debug, Args)]
struct ServeArgs {
    /// Local HTTP port. The server always binds to 127.0.0.1.
    #[arg(long, default_value_t = 8750)]
    port: u16,

    /// S2 configuration workbook. Defaults to OPEN_DATA_AUDIT_S2_CONFIG or the workspace placeholder.
    #[arg(long)]
    config: Option<PathBuf>,
}

#[derive(Debug, Args)]
struct S2Args {
    /// Directory containing current-period report workbooks.
    #[arg(long)]
    current_dir: PathBuf,

    /// Directory containing previous-period report workbooks.
    #[arg(long)]
    previous_dir: PathBuf,

    /// S2 comparison configuration workbook.
    #[arg(long)]
    config: PathBuf,

    /// Directory in which to create the result workbook.
    #[arg(long)]
    output_dir: PathBuf,

    /// Optional external workbook used by the configured cross-check rules.
    #[arg(long)]
    external: Option<PathBuf>,
}

#[derive(Debug, Serialize)]
struct S2CliOutput {
    status: &'static str,
    report_path: PathBuf,
    summary: S2AuditSummary,
}

fn main() {
    if let Err(error) = run() {
        eprintln!("open-data-audit: {error:#}");
        std::process::exit(2);
    }
}

fn run() -> Result<()> {
    let cli = Cli::parse();
    match cli.command {
        Command::Serve(arguments) => {
            let config_path = arguments.config.unwrap_or_else(default_s2_config_path);
            serve_web_ui(arguments.port, config_path)?;
        }
        Command::S2(arguments) => {
            let request = S2RunRequest {
                current_directory: arguments.current_dir,
                previous_directory: arguments.previous_dir,
                config_path: arguments.config,
                output_directory: arguments.output_dir,
                external_workbook: arguments.external,
            };
            let mut on_progress = |event: S2ProgressEvent| {
                let output = serde_json::json!({"type": "progress", "event": event});
                eprintln!(
                    "{}",
                    serde_json::to_string(&output).expect("progress event is serializable")
                );
            };
            let result = run_s2_audit(&request, Some(&mut on_progress))?;
            let output = S2CliOutput {
                status: "completed",
                report_path: result.report_path,
                summary: result.analysis.summary,
            };
            println!("{}", serde_json::to_string_pretty(&output)?);
        }
        Command::TemplateSuggest(arguments) => {
            let depth = if arguments.recursive {
                ScanDepth::Full
            } else {
                ScanDepth::RootOnly
            };
            let suggestion =
                recommend_template(&arguments.template_dir, &arguments.input_dir, depth)?;
            println!("{}", serde_json::to_string_pretty(&suggestion)?);
        }
        Command::ConditionalEvaluate(arguments) => {
            let report = evaluate_xlsx_workbook(&arguments.workbook, &arguments.scan_areas)?;
            println!("{}", serde_json::to_string_pretty(&report)?);
        }
        Command::ConditionalAudit(arguments) => {
            let (scan_areas, structure_ranges) = resolve_s1_conditional_cli_areas(
                arguments.template_workbook.as_deref(),
                arguments.mapping_config.as_deref(),
                &arguments.scan_areas,
                &arguments.structure_areas,
                arguments.developer_view,
            )?;
            let source_file = arguments
                .source_file
                .unwrap_or_else(|| arguments.workbook.clone());
            let metadata = S1ConditionalMetadata {
                period: arguments.period,
                batch_id: arguments.batch_id,
                audit_time: arguments.audit_time,
                org_code: arguments.org_code,
                org_name: arguments.org_name,
            };
            let run_started = Instant::now();
            let started_at = chrono::Local::now().to_rfc3339();
            if arguments.developer_view {
                eprintln!(
                    "{}",
                    serde_json::to_string(&serde_json::json!({
                        "type": "progress",
                        "product": "open-data-audit",
                        "feature": "S1-F01",
                        "step": "conditional_format_issue_extraction",
                        "status": "started",
                        "batch_id": metadata.batch_id,
                        "started_at": started_at
                    }))?
                );
            } else {
                eprintln!("正在执行 S1-F01 条件格式检查…");
            }
            let (preflighted, structure_preflight) = prepare_s1_structure_inputs(
                vec![S1ConditionalBatchInput {
                    workbook_path: arguments.workbook.clone(),
                    source_file: source_file.clone(),
                    metadata: metadata.clone(),
                }],
                arguments.template_workbook.as_deref(),
                &scan_areas,
                &structure_ranges,
                arguments.developer_view,
            )?;
            if preflighted.is_empty() {
                anyhow::bail!("报送文件与模板表结构不匹配，已跳过条件格式审核");
            }
            let report = match extract_conditional_issues(
                &arguments.workbook,
                &source_file,
                &scan_areas,
                &structure_ranges,
                &metadata,
            ) {
                Ok(report) => report,
                Err(error) => {
                    if arguments.developer_view {
                        eprintln!(
                            "{}",
                            serde_json::to_string(&serde_json::json!({
                                "type": "progress",
                                "product": "open-data-audit",
                                "feature": "S1-F01",
                                "step": "conditional_format_issue_extraction",
                                "status": "failed",
                                "finished_at": chrono::Local::now().to_rfc3339(),
                                "elapsed_ms": run_started.elapsed().as_millis(),
                                "error": format!("{error:#}")
                            }))?
                        );
                    } else {
                        eprintln!("条件格式检查失败，未生成审核项。");
                    }
                    return Err(error);
                }
            };
            if arguments.developer_view {
                eprintln!(
                    "{}",
                    serde_json::to_string(&serde_json::json!({
                        "type": "progress",
                        "product": "open-data-audit",
                        "feature": "S1-F01",
                        "step": "conditional_format_issue_extraction",
                        "status": "completed",
                        "batch_id": metadata.batch_id,
                        "started_at": started_at,
                        "finished_at": chrono::Local::now().to_rfc3339(),
                        "elapsed_ms": run_started.elapsed().as_millis(),
                        "engine": report.evaluation.reader,
                        "rules_total": report.evaluation.rules_total,
                        "rules_supported": report.evaluation.rules_supported,
                        "scanned_cells": report.evaluation.scanned_cells,
                        "issues": report.issues.len(),
                        "diagnostics": report.evaluation.diagnostics.len(),
                        "read_time_ms": report.evaluation.read_time_ms,
                        "evaluation_time_ms": report.evaluation.evaluation_time_ms,
                        "issue_build_time_ms": report.issue_build_time_ms,
                        "total_time_ms": report.evaluation.total_time_ms + report.issue_build_time_ms
                    }))?
                );
            } else {
                eprintln!(
                    "条件格式检查完成：引擎 {}，{} 条审核项；读取 {} ms、判定 {} ms、组装 {} ms，实际耗时 {} ms。",
                    report.evaluation.reader,
                    report.issues.len(),
                    report.evaluation.read_time_ms,
                    report.evaluation.evaluation_time_ms,
                    report.issue_build_time_ms,
                    run_started.elapsed().as_millis()
                );
            }
            let mut output = serde_json::to_value(&report)?;
            if let Some(summary) = structure_preflight {
                output["structure_preflight"] = serde_json::to_value(summary)?;
            }
            println!("{}", serde_json::to_string_pretty(&output)?);
        }
        Command::ConditionalAuditBatch(arguments) => {
            run_conditional_audit_batch(arguments)?;
        }
        Command::ConditionalAuditDirectory(arguments) => {
            run_conditional_audit_directory(arguments)?;
        }
        Command::S1TemplateRules(arguments) => {
            run_s1_template_rules(arguments)?;
        }
    }
    Ok(())
}

fn run_s1_template_rules(arguments: S1TemplateRulesArgs) -> Result<()> {
    let started = Instant::now();
    let started_at = chrono::Local::now().to_rfc3339();
    if arguments.developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S1-F01",
                "step": "template_formula_rule_discovery",
                "status": "started",
                "engine": "calamine+namespace_aware_ooxml",
                "started_at": started_at
            }))?
        );
    } else {
        eprintln!("正在读取并检查 S1 模板公式规则…");
    }

    let plan = match load_s1_template_rule_plan(
        &arguments.template_workbook,
        arguments.mapping_config.as_deref(),
    ) {
        Ok(plan) => plan,
        Err(error) => {
            if arguments.developer_view {
                eprintln!(
                    "{}",
                    serde_json::to_string(&serde_json::json!({
                        "type": "progress",
                        "product": "open-data-audit",
                        "feature": "S1-F01",
                        "step": "template_formula_rule_discovery",
                        "status": "failed",
                        "engine": "calamine+namespace_aware_ooxml",
                        "finished_at": chrono::Local::now().to_rfc3339(),
                        "elapsed_ms": started.elapsed().as_millis(),
                        "error": format!("{error:#}")
                    }))?
                );
            } else {
                eprintln!("S1 模板公式规则检查失败：{error:#}");
            }
            return Err(error);
        }
    };

    if arguments.developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S1-F01",
                "step": "template_formula_rule_discovery",
                "status": "completed",
                "started_at": started_at,
                "finished_at": chrono::Local::now().to_rfc3339(),
                "elapsed_ms": plan.summary.elapsed_ms,
                "summary": plan.summary
            }))?
        );
    } else {
        eprintln!(
            "S1 模板规则检查完成：模式 {}，规则 {} 条（启用 {} 条），公式范围 {} 个，结构范围 {} 个，读取引擎 {}，读取 {} ms，合计 {} ms。",
            plan.summary.template_mode,
            plan.summary.rules_total,
            plan.summary.rules_enabled,
            plan.summary.formula_named_ranges,
            plan.summary.structure_ranges,
            plan.summary.engine,
            plan.summary.read_time_ms,
            plan.summary.elapsed_ms
        );
    }

    let mut output = serde_json::json!({
        "status": "completed",
        "summary": plan.summary
    });
    if arguments.include_rule_definitions {
        output["definition"] = serde_json::to_value(plan.definition)?;
    }
    println!("{}", serde_json::to_string_pretty(&output)?);
    Ok(())
}

fn run_conditional_audit_batch(arguments: ConditionalAuditBatchArgs) -> Result<()> {
    let (scan_areas, structure_ranges) = resolve_s1_conditional_cli_areas(
        arguments.template_workbook.as_deref(),
        arguments.mapping_config.as_deref(),
        &arguments.scan_areas,
        &arguments.structure_areas,
        arguments.developer_view,
    )?;
    let manifest: ConditionalAuditManifest = serde_json::from_str(
        &fs::read_to_string(&arguments.manifest)
            .with_context(|| format!("读取 S1 条件格式清单 {}", arguments.manifest.display()))?,
    )
    .with_context(|| format!("解析 S1 条件格式清单 {}", arguments.manifest.display()))?;
    if manifest.files.is_empty() {
        anyhow::bail!("S1 条件格式清单中没有工作簿");
    }
    let manifest_dir = arguments
        .manifest
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let inputs = manifest
        .files
        .into_iter()
        .map(|file| {
            let workbook_path = resolve_manifest_path(manifest_dir, file.workbook);
            let source_file = file
                .source_file
                .map(|path| resolve_manifest_path(manifest_dir, path))
                .unwrap_or_else(|| workbook_path.clone());
            S1ConditionalBatchInput {
                workbook_path,
                source_file,
                metadata: S1ConditionalMetadata {
                    period: arguments.period.clone(),
                    batch_id: arguments.batch_id.clone(),
                    audit_time: arguments.audit_time.clone(),
                    org_code: file.org_code,
                    org_name: file.org_name,
                },
            }
        })
        .collect::<Vec<_>>();
    let excluded_workbooks = excluded_s1_input_workbooks(
        arguments.template_workbook.as_deref(),
        arguments.mapping_config.as_deref(),
        None,
    )?;
    let inputs = inputs
        .into_iter()
        .filter(|input| {
            !excluded_workbooks
                .iter()
                .any(|excluded| same_file_path(&input.workbook_path, excluded))
        })
        .collect::<Vec<_>>();
    if inputs.is_empty() {
        anyhow::bail!("S1 条件格式清单中没有可审核的源工作簿");
    }
    run_conditional_audit_inputs(
        inputs,
        arguments.template_workbook.as_deref(),
        &scan_areas,
        &structure_ranges,
        &arguments.batch_id,
        arguments.developer_view,
    )
}

fn run_conditional_audit_directory(arguments: ConditionalAuditDirectoryArgs) -> Result<()> {
    let discovery_started = Instant::now();
    let discovery_started_at = chrono::Local::now().to_rfc3339();
    if arguments.developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S1-F01",
                "step": "source_workbook_discovery",
                "status": "started",
                "input_dir": arguments.input_dir,
                "recursive": arguments.recursive,
                "selected_files_requested": arguments.selected_files.len(),
                "started_at": discovery_started_at
            }))?
        );
    } else {
        eprintln!("正在扫描 S1-F01 源目录中的 XLSX 工作簿…");
    }

    let discovery_result = (|| -> Result<_> {
        let depth = if arguments.recursive {
            ScanDepth::Full
        } else {
            ScanDepth::RootOnly
        };
        let excluded_workbooks = excluded_s1_input_workbooks(
            arguments.template_workbook.as_deref(),
            arguments.mapping_config.as_deref(),
            arguments.template_dir.as_deref(),
        )?;
        let source_files = source_workbooks(&arguments.input_dir, depth)?
            .into_iter()
            .filter(|path| {
                !excluded_workbooks
                    .iter()
                    .any(|excluded| same_file_path(path, excluded))
            })
            .collect::<Vec<_>>();
        let available_source_files = source_files.len();
        let source_files = select_s1_source_workbooks(
            &arguments.input_dir,
            &source_files,
            &arguments.selected_files,
        )?;
        if source_files.is_empty() {
            anyhow::bail!("源数据目录中没有可审核的 .xlsx 文件（模板和映射配置不会作为源文件）");
        }
        let detection = detect_period_from_workbooks(&arguments.input_dir, &source_files);
        let explicit_period = arguments
            .period
            .as_deref()
            .map(str::trim)
            .filter(|period| !period.is_empty());
        let (period, period_source) = if let Some(period) = explicit_period {
            (period.to_owned(), "用户指定".to_owned())
        } else if detection.conflict {
            anyhow::bail!("{}；请显式指定 --period", detection.details);
        } else if detection.period.is_empty() {
            anyhow::bail!("{}；请使用 --period 指定数据期", detection.details);
        } else {
            (detection.period.clone(), detection.source.clone())
        };
        Ok((
            source_files,
            period,
            period_source,
            detection.details,
            available_source_files,
        ))
    })();

    let (source_files, period, period_source, detection_details, available_source_files) =
        match discovery_result {
            Ok(result) => result,
            Err(error) => {
                if arguments.developer_view {
                    eprintln!(
                        "{}",
                        serde_json::to_string(&serde_json::json!({
                            "type": "progress",
                            "product": "open-data-audit",
                            "feature": "S1-F01",
                            "step": "source_workbook_discovery",
                            "status": "failed",
                            "input_dir": arguments.input_dir,
                            "selected_files_requested": arguments.selected_files.len(),
                            "started_at": discovery_started_at,
                            "finished_at": chrono::Local::now().to_rfc3339(),
                            "elapsed_ms": discovery_started.elapsed().as_millis(),
                            "error": format!("{error:#}")
                        }))?
                    );
                } else {
                    eprintln!("S1-F01 源目录扫描失败：{error:#}");
                }
                return Err(error);
            }
        };

    if arguments.developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S1-F01",
                "step": "source_workbook_discovery",
                "status": "completed",
                "input_dir": arguments.input_dir,
                "available_source_files": available_source_files,
                "files_total": source_files.len(),
                "selected_files_requested": arguments.selected_files.len(),
                "selection_applied": !arguments.selected_files.is_empty(),
                "period": period,
                "period_source": period_source,
                "detection_details": detection_details,
                "started_at": discovery_started_at,
                "finished_at": chrono::Local::now().to_rfc3339(),
                "elapsed_ms": discovery_started.elapsed().as_millis()
            }))?
        );
    } else {
        if arguments.selected_files.is_empty() {
            eprintln!(
                "源目录发现 {} 个 XLSX；数据期 {}（{}）；筛选耗时 {} ms。",
                source_files.len(),
                period,
                period_source,
                discovery_started.elapsed().as_millis()
            );
        } else {
            eprintln!(
                "源目录有 {} 个可审核 XLSX，按所选文件筛出 {} 个；数据期 {}（{}）；筛选耗时 {} ms。",
                available_source_files,
                source_files.len(),
                period,
                period_source,
                discovery_started.elapsed().as_millis()
            );
        }
    }

    let template_workbook = if let Some(template_dir) = arguments.template_dir.as_deref() {
        Some(recommend_s1_directory_template(
            template_dir,
            &source_files,
            arguments.developer_view,
        )?)
    } else {
        arguments.template_workbook.clone()
    };
    let (scan_areas, structure_ranges) = resolve_s1_conditional_cli_areas(
        template_workbook.as_deref(),
        arguments.mapping_config.as_deref(),
        &arguments.scan_areas,
        &arguments.structure_areas,
        arguments.developer_view,
    )?;
    let inputs = source_files
        .into_iter()
        .map(|path| {
            let (org_code, org_name) = organisation_from_name(&path);
            S1ConditionalBatchInput {
                workbook_path: path.clone(),
                source_file: path,
                metadata: S1ConditionalMetadata {
                    period: period.clone(),
                    batch_id: arguments.batch_id.clone(),
                    audit_time: arguments.audit_time.clone(),
                    org_code,
                    org_name,
                },
            }
        })
        .collect::<Vec<_>>();

    run_conditional_audit_inputs(
        inputs,
        template_workbook.as_deref(),
        &scan_areas,
        &structure_ranges,
        &arguments.batch_id,
        arguments.developer_view,
    )
}

fn run_conditional_audit_inputs(
    inputs: Vec<S1ConditionalBatchInput>,
    template_workbook: Option<&Path>,
    scan_areas: &[String],
    structure_ranges: &[CopyRange],
    batch_id: &str,
    developer_view: bool,
) -> Result<()> {
    let batch_started = Instant::now();
    let started_at = chrono::Local::now().to_rfc3339();
    let (inputs, structure_preflight) = prepare_s1_structure_inputs(
        inputs,
        template_workbook,
        scan_areas,
        structure_ranges,
        developer_view,
    )?;
    if developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S1-F01",
                "step": "conditional_format_issue_extraction",
                "status": "started",
                "batch_id": batch_id,
                "files_total": inputs.len(),
                "started_at": started_at.as_str()
            }))?
        );
    } else {
        eprintln!("正在执行 S1-F01 条件格式检查：{} 个工作簿。", inputs.len());
    }

    let report = if inputs.is_empty() {
        Ok(S1ConditionalBatchAuditResult {
            status: "completed_with_diagnostics",
            issues: Vec::new(),
            files: Vec::new(),
            total_elapsed_ms: 0,
        })
    } else {
        extract_conditional_issues_batch(
            &inputs,
            scan_areas,
            structure_ranges,
            |progress: &S1ConditionalBatchProgress| {
                if developer_view {
                    if let Ok(event) = serde_json::to_string(&serde_json::json!({
                        "type": "progress",
                        "product": "open-data-audit",
                        "feature": "S1-F01",
                        "step": "conditional_format_issue_extraction",
                        "batch_id": batch_id,
                        "event": progress
                    })) {
                        eprintln!("{event}");
                    }
                } else {
                    match progress.status {
                    "started" => eprintln!(
                        "正在检查条件格式：{}（{}/{}）",
                        progress.source_file, progress.file_index, progress.files_total
                    ),
                    "completed" => eprintln!(
                        "{}：引擎 {}，规则 {}/{}，扫描 {} 格，触发 {} 项，未确定 {} 项；读取 {} ms、判定 {} ms、组装 {} ms、单文件总计 {} ms。",
                        progress.source_file,
                        progress.reader.unwrap_or("未知"),
                        progress.rules_supported.unwrap_or_default(),
                        progress.rules_total.unwrap_or_default(),
                        progress.scanned_cells.unwrap_or_default(),
                        progress.issue_count.unwrap_or_default(),
                        progress.diagnostics.unwrap_or_default(),
                        progress.read_time_ms.unwrap_or_default(),
                        progress.evaluation_time_ms.unwrap_or_default(),
                        progress.issue_build_time_ms.unwrap_or_default(),
                        progress.elapsed_ms.unwrap_or_default()
                    ),
                    "failed" => eprintln!(
                        "{}：条件格式检查失败（{} ms）：{}",
                        progress.source_file,
                        progress.elapsed_ms.unwrap_or_default(),
                        progress.error.as_deref().unwrap_or("未知错误")
                    ),
                    _ => {}
                }
                }
            },
        )
    };

    let mut report = match report {
        Ok(report) => report,
        Err(error) => {
            if developer_view {
                eprintln!(
                    "{}",
                    serde_json::to_string(&serde_json::json!({
                        "type": "progress",
                        "product": "open-data-audit",
                        "feature": "S1-F01",
                        "step": "conditional_format_issue_extraction",
                        "status": "failed",
                        "batch_id": batch_id,
                        "started_at": started_at.as_str(),
                        "finished_at": chrono::Local::now().to_rfc3339(),
                        "elapsed_ms": batch_started.elapsed().as_millis(),
                        "error": format!("{error:#}")
                    }))?
                );
            } else {
                eprintln!("S1-F01 条件格式批量审核失败，未输出汇总结果。");
            }
            return Err(error);
        }
    };
    report.total_elapsed_ms = batch_started.elapsed().as_millis();
    if structure_preflight
        .as_ref()
        .is_some_and(|summary| summary.files_rejected > 0)
    {
        report.status = "completed_with_diagnostics";
    }

    if developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S1-F01",
                "step": "conditional_format_issue_extraction",
                "status": "completed",
                "batch_id": batch_id,
                "started_at": started_at.as_str(),
                "finished_at": chrono::Local::now().to_rfc3339(),
                "elapsed_ms": batch_started.elapsed().as_millis(),
                "files": report.files.len(),
                "issues": report.issues.len(),
                "result_status": report.status
            }))?
        );
    } else {
        eprintln!(
            "条件格式批量检查完成：{} 个工作簿，形成 {} 条审核项，状态 {}，总耗时 {} ms。",
            report.files.len(),
            report.issues.len(),
            report.status,
            batch_started.elapsed().as_millis()
        );
    }
    let mut output = serde_json::to_value(&report)?;
    if let Some(summary) = structure_preflight {
        output["structure_preflight"] = serde_json::to_value(summary)?;
    }
    println!("{}", serde_json::to_string_pretty(&output)?);
    Ok(())
}

fn prepare_s1_structure_inputs(
    inputs: Vec<S1ConditionalBatchInput>,
    template_workbook: Option<&Path>,
    scan_areas: &[String],
    structure_ranges: &[CopyRange],
    developer_view: bool,
) -> Result<(
    Vec<S1ConditionalBatchInput>,
    Option<S1StructurePreflightSummary>,
)> {
    let Some(template_workbook) = template_workbook else {
        return Ok((inputs, None));
    };
    if structure_ranges.is_empty() {
        if developer_view {
            eprintln!(
                "{}",
                serde_json::to_string(&serde_json::json!({
                    "type": "progress",
                    "product": "open-data-audit",
                    "feature": "S1-F01",
                    "step": "structure_preflight",
                    "status": "skipped",
                    "reason": "template_has_no_mapped_structure_ranges"
                }))?
            );
        } else {
            eprintln!("未配置模板表结构区域，本次跳过结构预检。");
        }
        return Ok((inputs, None));
    }

    let started = Instant::now();
    let started_at = chrono::Local::now().to_rfc3339();
    if developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S1-F01",
                "step": "structure_preflight",
                "status": "started",
                "engine": "calamine",
                "files_total": inputs.len(),
                "template_workbook": template_workbook.file_name().map(|name| name.to_string_lossy()),
                "started_at": started_at
            }))?
        );
    } else {
        eprintln!("正在按模板检查 S1-F01 报送文件的固定表结构…");
    }

    let required_sheets = scan_areas
        .iter()
        .filter_map(|area| scan_area_sheet_name(area))
        .chain(
            structure_ranges
                .iter()
                .map(|range| range.sheet_name.clone()),
        )
        .collect::<Vec<_>>();
    let template = match read_s1_structure_template_with_required_sheets(
        template_workbook,
        structure_ranges,
        &required_sheets,
    ) {
        Ok(template) => template,
        Err(error) => {
            if developer_view {
                eprintln!(
                    "{}",
                    serde_json::to_string(&serde_json::json!({
                        "type": "progress",
                        "product": "open-data-audit",
                        "feature": "S1-F01",
                        "step": "structure_preflight",
                        "status": "failed",
                        "engine": "calamine",
                        "started_at": started_at,
                        "finished_at": chrono::Local::now().to_rfc3339(),
                        "elapsed_ms": started.elapsed().as_millis(),
                        "error": format!("{error:#}")
                    }))?
                );
            } else {
                eprintln!("无法读取模板表结构，已停止本次审核。");
            }
            return Err(error.context("S1-F01 表结构模板读取失败"));
        }
    };

    let outcomes = inputs
        .par_iter()
        .map(|input| {
            let file_started = Instant::now();
            (
                input.clone(),
                compare_s1_source_structure(&input.workbook_path, &template),
                file_started.elapsed().as_millis(),
            )
        })
        .collect::<Vec<_>>();
    let mut accepted = Vec::new();
    let mut files = Vec::with_capacity(outcomes.len());
    for (input, result, file_elapsed_ms) in outcomes {
        let source_file = input
            .source_file
            .file_name()
            .map(|name| name.to_string_lossy().into_owned())
            .unwrap_or_else(|| input.source_file.display().to_string());
        let (
            status,
            matched_labels,
            checked_labels,
            score,
            missing_sheets,
            read_time_ms,
            elapsed_ms,
            diagnostic,
            error,
        ) = match result {
            Ok(result) => {
                if result.matched {
                    accepted.push(input);
                }
                (
                    if result.matched {
                        "matched"
                    } else {
                        "rejected"
                    },
                    result.matched_labels,
                    result.checked_labels,
                    result.score,
                    result.missing_sheets,
                    Some(result.read_time_ms),
                    result.elapsed_ms,
                    result.diagnostic,
                    None,
                )
            }
            Err(error) => (
                "read_error",
                0,
                0,
                0.0,
                Vec::new(),
                None,
                file_elapsed_ms,
                None,
                Some(format!("{error:#}")),
            ),
        };
        let file = S1StructurePreflightFile {
            source_file: source_file.clone(),
            status,
            matched_labels,
            checked_labels,
            score,
            missing_sheets: missing_sheets.clone(),
            read_time_ms,
            elapsed_ms,
            diagnostic: diagnostic.clone(),
            error: error.as_ref().map(|_| "无法读取报送文件结构".to_owned()),
        };

        if developer_view {
            eprintln!(
                "{}",
                serde_json::to_string(&serde_json::json!({
                    "type": "progress",
                    "product": "open-data-audit",
                    "feature": "S1-F01",
                    "step": "structure_preflight_file",
                    "status": status,
                    "engine": "calamine",
                    "source_file": source_file,
                    "matched_labels": matched_labels,
                    "checked_labels": checked_labels,
                    "score": score,
                    "missing_sheets": missing_sheets,
                    "read_time_ms": read_time_ms,
                    "elapsed_ms": elapsed_ms,
                    "diagnostic": diagnostic,
                    "error": error
                }))?
            );
        } else if status == "rejected" {
            eprintln!(
                "{}：表结构未通过，匹配 {}/{}（{:.0}%），已跳过。{}{}",
                source_file,
                matched_labels,
                checked_labels,
                score * 100.0,
                if missing_sheets.is_empty() {
                    String::new()
                } else {
                    format!("缺少工作表：{}。", missing_sheets.join("、"))
                },
                diagnostic
                    .as_deref()
                    .map(|message| format!("{message}。"))
                    .unwrap_or_default()
            );
        } else if status == "read_error" {
            eprintln!("{}：无法读取表结构，已跳过。", source_file);
        }
        files.push(file);
    }

    let files_rejected = files.len() - accepted.len();
    let summary = S1StructurePreflightSummary {
        status: if files_rejected == 0 {
            "completed"
        } else {
            "completed_with_rejections"
        },
        engine: "calamine",
        files_total: files.len(),
        files_accepted: accepted.len(),
        files_rejected,
        template_areas: template.area_count,
        template_labels: template.label_count,
        template_read_time_ms: template.read_time_ms,
        elapsed_ms: started.elapsed().as_millis(),
        files,
    };
    if developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S1-F01",
                "step": "structure_preflight",
                "status": summary.status,
                "engine": summary.engine,
                "files_total": summary.files_total,
                "files_accepted": summary.files_accepted,
                "files_rejected": summary.files_rejected,
                "template_areas": summary.template_areas,
                "template_labels": summary.template_labels,
                "template_read_time_ms": summary.template_read_time_ms,
                "started_at": started_at,
                "finished_at": chrono::Local::now().to_rfc3339(),
                "elapsed_ms": summary.elapsed_ms
            }))?
        );
    } else {
        eprintln!(
            "表结构预检完成：通过 {} 个，跳过 {} 个；读取引擎 {}，模板读取 {} ms，预检 {} ms。",
            summary.files_accepted,
            summary.files_rejected,
            summary.engine,
            summary.template_read_time_ms,
            summary.elapsed_ms
        );
    }
    Ok((accepted, Some(summary)))
}

fn recommend_s1_directory_template(
    template_dir: &Path,
    source_files: &[PathBuf],
    developer_view: bool,
) -> Result<PathBuf> {
    let started = Instant::now();
    let started_at = chrono::Local::now().to_rfc3339();
    if developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S1-F01",
                "step": "template_recommendation",
                "status": "started",
                "engine": "namespace_aware_ooxml_sheet_directory",
                "source_workbooks": source_files.len(),
                "started_at": started_at
            }))?
        );
    } else {
        eprintln!("正在按所选源工作簿推荐 S1-F01 模板…");
    }

    let suggestion = match recommend_template_for_workbooks(template_dir, source_files) {
        Ok(suggestion) => suggestion,
        Err(error) => {
            if developer_view {
                eprintln!(
                    "{}",
                    serde_json::to_string(&serde_json::json!({
                        "type": "progress",
                        "product": "open-data-audit",
                        "feature": "S1-F01",
                        "step": "template_recommendation",
                        "status": "failed",
                        "engine": "namespace_aware_ooxml_sheet_directory",
                        "source_workbooks": source_files.len(),
                        "started_at": started_at,
                        "finished_at": chrono::Local::now().to_rfc3339(),
                        "elapsed_ms": started.elapsed().as_millis(),
                        "error": format!("{error:#}")
                    }))?
                );
            } else {
                eprintln!("S1-F01 模板推荐失败：{error:#}");
            }
            return Err(error.context("S1-F01 模板推荐失败"));
        }
    };

    let Some(template_path) = suggestion
        .template_path
        .as_ref()
        .filter(|_| suggestion.matched)
        .cloned()
    else {
        let error = anyhow::anyhow!(
            "模板自动推荐未能唯一确定模板：{}；请显式指定 --template-workbook",
            suggestion.details
        );
        if developer_view {
            eprintln!(
                "{}",
                serde_json::to_string(&serde_json::json!({
                    "type": "progress",
                    "product": "open-data-audit",
                    "feature": "S1-F01",
                    "step": "template_recommendation",
                    "status": "failed",
                    "engine": "namespace_aware_ooxml_sheet_directory",
                    "source_workbooks": source_files.len(),
                    "matched": suggestion.matched,
                    "confidence": suggestion.confidence,
                    "details": suggestion.details,
                    "alternatives": suggestion.alternatives,
                    "started_at": started_at,
                    "finished_at": chrono::Local::now().to_rfc3339(),
                    "elapsed_ms": started.elapsed().as_millis(),
                    "error": format!("{error:#}")
                }))?
            );
        } else {
            eprintln!("S1-F01 模板推荐未能确定唯一模板：{}", suggestion.details);
            eprintln!("请检查候选模板，或显式指定 --template-workbook。");
        }
        return Err(error);
    };

    if developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S1-F01",
                "step": "template_recommendation",
                "status": "completed",
                "engine": "namespace_aware_ooxml_sheet_directory",
                "source_workbooks": source_files.len(),
                "matched": suggestion.matched,
                "confidence": suggestion.confidence,
                "details": suggestion.details,
                "selected_template": template_path
                    .file_name()
                    .map(|name| name.to_string_lossy().into_owned()),
                "alternatives": suggestion.alternatives,
                "started_at": started_at,
                "finished_at": chrono::Local::now().to_rfc3339(),
                "elapsed_ms": started.elapsed().as_millis()
            }))?
        );
    } else {
        eprintln!(
            "S1-F01 模板推荐完成：{}；可信度 {:.0}%；引擎 namespace-aware OOXML 工作表目录读取，耗时 {} ms。",
            suggestion.details,
            suggestion.confidence * 100.0,
            started.elapsed().as_millis()
        );
    }
    Ok(template_path)
}

fn resolve_s1_conditional_cli_areas(
    template_workbook: Option<&Path>,
    mapping_config: Option<&Path>,
    scan_area_overrides: &[String],
    structure_area_overrides: &[String],
    developer_view: bool,
) -> Result<(Vec<String>, Vec<CopyRange>)> {
    let explicit_structure_ranges = structure_area_overrides
        .iter()
        .map(|area| parse_copy_range(area))
        .collect::<Result<Vec<_>>>()?;
    let Some(template_workbook) = template_workbook else {
        if mapping_config.is_some() {
            anyhow::bail!("--mapping-config 需要同时指定 --template-workbook");
        }
        return Ok((scan_area_overrides.to_vec(), explicit_structure_ranges));
    };

    let started = Instant::now();
    let started_at = chrono::Local::now().to_rfc3339();
    if developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S1-F01",
                "step": "conditional_area_mapping",
                "status": "started",
                "engine": "calamine+namespace_aware_ooxml",
                "template_workbook": template_workbook
                    .file_name()
                    .map(|name| name.to_string_lossy().into_owned()),
                "started_at": started_at
            }))?
        );
    }

    let resolution = (|| -> Result<_> {
        let config_path = effective_s1_mapping_config_path(mapping_config);
        let mappings = load_s1_conditional_mappings(config_path.as_deref(), template_workbook)?;
        let defined_names = excel_reader::read_defined_names(template_workbook)
            .with_context(|| format!("读取模板命名区域 {}", template_workbook.display()))?;
        let workbook_sheets = excel_reader::read_sheet_names(template_workbook)
            .with_context(|| format!("读取模板工作表 {}", template_workbook.display()))?;
        let areas = named_areas_from_defined_names(&defined_names);
        let unresolved_mappings = defined_names
            .iter()
            .filter(|record| {
                mappings
                    .iter()
                    .any(|mapping| matches_named_range(mapping, &record.name))
                    && !defined_name_is_fully_supported(record)
            })
            .count();
        let conditional_mapping = mappings
            .iter()
            .find(|mapping| {
                mapping.name == DEFAULT_CONDITIONAL_FEATURE_NAME
                    && mapping.feature_type == CONDITIONAL_FORMAT_EXTRACT_FUNCTION
            })
            .or_else(|| {
                mappings
                    .iter()
                    .find(|mapping| mapping.feature_type == CONDITIONAL_FORMAT_EXTRACT_FUNCTION)
            });
        let unresolved_conditional_names = defined_names
            .iter()
            .filter(|record| {
                conditional_mapping.is_some_and(|mapping| {
                    matches_named_range(mapping, &record.name)
                        && !defined_name_is_fully_supported(record)
                })
            })
            .map(|record| record.name.as_str())
            .collect::<Vec<_>>();
        if scan_area_overrides.is_empty() && !unresolved_conditional_names.is_empty() {
            anyhow::bail!(
                "模板中有 {} 个条件格式命名区域无法由共享 Defined Names reader 解析；为避免将扫描范围扩大到规则 AppliesTo，已停止本次审核",
                unresolved_conditional_names.len()
            );
        }
        let plan = resolve_s1_conditional_area_plan(&mappings, None, &areas, &workbook_sheets)?;
        let mapped_area_count = plan.scan_areas.len() + plan.structure_ranges.len();
        let scan_areas = if scan_area_overrides.is_empty() {
            plan.scan_areas
        } else {
            scan_area_overrides.to_vec()
        };
        let structure_ranges = if structure_area_overrides.is_empty() {
            plan.structure_ranges
        } else {
            explicit_structure_ranges.clone()
        };
        Ok((
            scan_areas,
            structure_ranges,
            mappings.len(),
            defined_names.len(),
            mapped_area_count,
            unresolved_mappings,
            config_path,
        ))
    })();

    match resolution {
        Ok((
            scan_areas,
            structure_ranges,
            mappings,
            defined_names,
            mapped_areas,
            unresolved,
            config_path,
        )) => {
            let source = if scan_area_overrides.is_empty() && structure_area_overrides.is_empty() {
                if scan_areas.is_empty() {
                    "conditional_format_applies_to"
                } else {
                    "template_defined_names"
                }
            } else if mapped_areas == 0 {
                "explicit_overrides"
            } else {
                "explicit_overrides_and_template_defined_names"
            };
            if developer_view {
                eprintln!(
                    "{}",
                    serde_json::to_string(&serde_json::json!({
                        "type": "progress",
                        "product": "open-data-audit",
                        "feature": "S1-F01",
                        "step": "conditional_area_mapping",
                        "status": "completed",
                        "engine": "calamine+namespace_aware_ooxml",
                        "area_source": source,
                        "template_workbook": template_workbook
                    .file_name()
                    .map(|name| name.to_string_lossy().into_owned()),
                        "mapping_config": config_path
                            .as_deref()
                            .and_then(Path::file_name)
                            .map(|name| name.to_string_lossy().into_owned()),
                        "mappings_loaded": mappings,
                        "defined_names_total": defined_names,
                        "mapped_named_areas": mapped_areas,
                        "unresolved_matching_names": unresolved,
                        "scan_areas": scan_areas.len(),
                        "structure_areas": structure_ranges.len(),
                        "started_at": started_at,
                        "finished_at": chrono::Local::now().to_rfc3339(),
                        "elapsed_ms": started.elapsed().as_millis()
                    }))?
                );
            } else {
                eprintln!(
                    "模板命名区域映射完成：引擎 calamine+namespace_aware_ooxml，条件范围 {} 个，结构范围 {} 个，匹配但未解析的名称 {} 个，耗时 {} ms。",
                    scan_areas.len(),
                    structure_ranges.len(),
                    unresolved,
                    started.elapsed().as_millis()
                );
            }
            Ok((scan_areas, structure_ranges))
        }
        Err(error) => {
            if developer_view {
                eprintln!(
                    "{}",
                    serde_json::to_string(&serde_json::json!({
                        "type": "progress",
                        "product": "open-data-audit",
                        "feature": "S1-F01",
                        "step": "conditional_area_mapping",
                        "status": "failed",
                        "engine": "calamine+namespace_aware_ooxml",
                        "template_workbook": template_workbook
                    .file_name()
                    .map(|name| name.to_string_lossy().into_owned()),
                        "started_at": started_at,
                        "finished_at": chrono::Local::now().to_rfc3339(),
                        "elapsed_ms": started.elapsed().as_millis(),
                        "error": format!("{error:#}")
                    }))?
                );
            } else {
                eprintln!("模板命名区域映射失败：{error:#}");
            }
            Err(error)
        }
    }
}

fn defined_name_is_fully_supported(record: &excel_reader::DefinedNameRecord) -> bool {
    let resolved = named_areas_from_defined_names(std::slice::from_ref(record));
    !resolved.is_empty() && resolved.len() == record.areas.len()
}

fn effective_s1_mapping_config_path(mapping_config: Option<&Path>) -> Option<PathBuf> {
    mapping_config
        .map(Path::to_path_buf)
        .or_else(|| std::env::var_os("OPEN_DATA_AUDIT_CONFIG").map(PathBuf::from))
}

fn excluded_s1_input_workbooks(
    template_workbook: Option<&Path>,
    mapping_config: Option<&Path>,
    template_dir: Option<&Path>,
) -> Result<Vec<PathBuf>> {
    let mut excluded = template_workbook
        .map(Path::to_path_buf)
        .into_iter()
        .collect::<Vec<_>>();
    if let Some(template_dir) = template_dir {
        excluded.extend(template_workbook_paths(template_dir)?);
    }
    excluded.extend(effective_s1_mapping_config_path(mapping_config));
    Ok(excluded)
}

fn select_s1_source_workbooks(
    input_dir: &Path,
    available_files: &[PathBuf],
    selected_files: &[PathBuf],
) -> Result<Vec<PathBuf>> {
    if selected_files.is_empty() {
        return Ok(available_files.to_vec());
    }

    let selected_paths = selected_files
        .iter()
        .map(|path| {
            if path.is_absolute() {
                path.clone()
            } else {
                input_dir.join(path)
            }
        })
        .collect::<Vec<_>>();
    if let Some(invalid) = selected_paths.iter().find(|selected| {
        !available_files
            .iter()
            .any(|available| same_file_path(available, selected))
    }) {
        anyhow::bail!(
            "选择的待审核文件不在源数据目录或当前扫描范围内：{}",
            invalid.display()
        );
    }

    Ok(available_files
        .iter()
        .filter(|available| {
            selected_paths
                .iter()
                .any(|selected| same_file_path(available, selected))
        })
        .cloned()
        .collect())
}

fn same_file_path(left: &Path, right: &Path) -> bool {
    let left = fs::canonicalize(left).unwrap_or_else(|_| left.to_path_buf());
    let right = fs::canonicalize(right).unwrap_or_else(|_| right.to_path_buf());
    #[cfg(windows)]
    {
        left.to_string_lossy()
            .eq_ignore_ascii_case(&right.to_string_lossy())
    }
    #[cfg(not(windows))]
    {
        left == right
    }
}

fn resolve_manifest_path(base: &Path, path: PathBuf) -> PathBuf {
    if path.is_absolute() {
        path
    } else {
        base.join(path)
    }
}

fn default_s2_config_path() -> PathBuf {
    if let Some(path) = std::env::var_os("OPEN_DATA_AUDIT_S2_CONFIG") {
        return PathBuf::from(path);
    }

    let current_directory = std::env::current_dir().unwrap_or_else(|_| PathBuf::from("."));
    let candidates = [
        current_directory.join("audit-core/config-real/2.报表采集系统_配置.xlsx"),
        current_directory.join("audit-core/config/默认配置/2.报表采集系统_配置.xlsx"),
        current_directory.join("config-real/2.报表采集系统_配置.xlsx"),
        current_directory.join("config/默认配置/2.报表采集系统_配置.xlsx"),
    ];
    candidates
        .into_iter()
        .find(|path| path.is_file())
        .unwrap_or_else(|| PathBuf::from("2.报表采集系统_配置.xlsx"))
}

fn parse_copy_range(specification: &str) -> Result<CopyRange> {
    let (sheet_name, address) = specification
        .split_once('!')
        .context("结构区域应使用 Sheet!A1:B20 格式")?;
    let sheet_name = decode_sheet_name(sheet_name.trim());
    let address = address.trim();
    if sheet_name.is_empty() || address.is_empty() {
        anyhow::bail!("结构区域应包含工作表名称和单元格范围");
    }
    Ok(CopyRange {
        sheet_name,
        address: address.to_owned(),
    })
}

fn scan_area_sheet_name(specification: &str) -> Option<String> {
    let (sheet_name, _) = specification.split_once('!')?;
    let sheet_name = decode_sheet_name(sheet_name.trim());
    (!sheet_name.is_empty()).then_some(sheet_name)
}

fn decode_sheet_name(value: &str) -> String {
    if let Some(quoted) = value
        .strip_prefix('\'')
        .and_then(|value| value.strip_suffix('\''))
    {
        quoted.replace("''", "'")
    } else {
        value.to_owned()
    }
}
