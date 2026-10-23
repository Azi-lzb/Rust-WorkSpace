use anyhow::{Context, Result};
use audit_core::{
    conditional_format::evaluate_xlsx_workbook,
    discovery::{detect_period, organisation_from_name, source_workbooks, ScanDepth},
    models::CopyRange,
    s1_conditional_audit::{
        extract_conditional_issues, extract_conditional_issues_batch, S1ConditionalBatchInput,
        S1ConditionalBatchProgress, S1ConditionalMetadata,
    },
    s1_template::recommend_template,
    s2_service::{run_s2_audit, S2ProgressEvent, S2RunRequest},
    s2_summary::S2AuditSummary,
    web_bridge::serve_web_ui,
};
use clap::{Args, Parser, Subcommand};
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

    /// Optional evaluation area in Sheet!A1:B20 form. Repeat for multiple areas.
    #[arg(long = "scan-area")]
    scan_areas: Vec<String>,

    /// Optional template structure area used to resolve check_field. Repeat as needed.
    #[arg(long = "structure-area")]
    structure_areas: Vec<String>,

    /// Emit developer-view JSON progress events with actual engine and timings.
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
            let structure_ranges = arguments
                .structure_areas
                .iter()
                .map(|area| parse_copy_range(area))
                .collect::<Result<Vec<_>>>()?;
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
            let report = match extract_conditional_issues(
                &arguments.workbook,
                &source_file,
                &arguments.scan_areas,
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
            println!("{}", serde_json::to_string_pretty(&report)?);
        }
        Command::ConditionalAuditBatch(arguments) => {
            run_conditional_audit_batch(arguments)?;
        }
        Command::ConditionalAuditDirectory(arguments) => {
            run_conditional_audit_directory(arguments)?;
        }
    }
    Ok(())
}

fn run_conditional_audit_batch(arguments: ConditionalAuditBatchArgs) -> Result<()> {
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
    run_conditional_audit_inputs(
        inputs,
        &arguments.scan_areas,
        &arguments.structure_areas,
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
        let source_files = source_workbooks(&arguments.input_dir, depth)?;
        if source_files.is_empty() {
            anyhow::bail!("源数据目录中没有可审核的 .xlsx 文件");
        }
        let detection = detect_period(&arguments.input_dir, depth)?;
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
        Ok((inputs, period, period_source, detection.details))
    })();

    let (inputs, period, period_source, detection_details) = match discovery_result {
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
                "files_total": inputs.len(),
                "period": period,
                "period_source": period_source,
                "detection_details": detection_details,
                "started_at": discovery_started_at,
                "finished_at": chrono::Local::now().to_rfc3339(),
                "elapsed_ms": discovery_started.elapsed().as_millis()
            }))?
        );
    } else {
        eprintln!(
            "源目录发现 {} 个 XLSX；数据期 {}（{}）；筛选耗时 {} ms。",
            inputs.len(),
            period,
            period_source,
            discovery_started.elapsed().as_millis()
        );
    }

    run_conditional_audit_inputs(
        inputs,
        &arguments.scan_areas,
        &arguments.structure_areas,
        &arguments.batch_id,
        arguments.developer_view,
    )
}

fn run_conditional_audit_inputs(
    inputs: Vec<S1ConditionalBatchInput>,
    scan_areas: &[String],
    structure_areas: &[String],
    batch_id: &str,
    developer_view: bool,
) -> Result<()> {
    let structure_ranges = structure_areas
        .iter()
        .map(|area| parse_copy_range(area))
        .collect::<Result<Vec<_>>>()?;
    let batch_started = Instant::now();
    let started_at = chrono::Local::now().to_rfc3339();
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

    let report = extract_conditional_issues_batch(
        &inputs,
        scan_areas,
        &structure_ranges,
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
    );

    let report = match report {
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
    println!("{}", serde_json::to_string_pretty(&report)?);
    Ok(())
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
    let sheet_name = sheet_name.trim().trim_matches('\'');
    let address = address.trim();
    if sheet_name.is_empty() || address.is_empty() {
        anyhow::bail!("结构区域应包含工作表名称和单元格范围");
    }
    Ok(CopyRange {
        sheet_name: sheet_name.to_owned(),
        address: address.to_owned(),
    })
}
