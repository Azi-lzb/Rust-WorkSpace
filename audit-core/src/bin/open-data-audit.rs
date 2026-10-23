use anyhow::{Context, Result};
use audit_core::{
    conditional_format::evaluate_xlsx_workbook,
    discovery::ScanDepth,
    models::CopyRange,
    s1_conditional_audit::{extract_conditional_issues, S1ConditionalMetadata},
    s1_template::recommend_template,
    s2_service::{run_s2_audit, S2ProgressEvent, S2RunRequest},
    s2_summary::S2AuditSummary,
    web_bridge::serve_web_ui,
};
use clap::{Args, Parser, Subcommand};
use serde::Serialize;
use std::{path::PathBuf, time::Instant};

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
                    "条件格式检查完成：形成 {} 条审核项，耗时 {} 毫秒。",
                    report.issues.len(),
                    run_started.elapsed().as_millis()
                );
            }
            println!("{}", serde_json::to_string_pretty(&report)?);
        }
    }
    Ok(())
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
