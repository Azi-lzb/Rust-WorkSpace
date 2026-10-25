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
    s1_formula_audit::{extract_s1_formula_issues, S1FormulaMetadata},
    s1_history::{
        process_s1_history, read_s1_history_workbook, write_s1_history_workbook, S1HistoryRequest,
        S1HistoryWorkbookReadSummary,
    },
    s1_issue_chain::{merge_s1_issue_sources, S1IssueFilePair},
    s1_region_summary::summarize_s1_named_regions,
    s1_result_workbook::write_s1_result_workbook,
    s1_structure_compare::{
        compare_s1_source_structure, read_s1_structure_template_with_required_sheets,
    },
    s1_summary_merge::merge_s1_workbook_tables,
    s1_template::{recommend_template, recommend_template_for_workbooks, template_workbook_paths},
    s1_template_rules::{assess_s1_template_health, load_s1_template_rule_plan},
    s1_validation_formula_scan::{
        issues_from_validation_formula_scan, validation_formula_scan_failure_issue,
        S1ValidationFormulaScanPlan,
    },
    s2_service::{run_s2_audit, s2_engine_for_stage, S2ProgressEvent, S2RunRequest},
    s2_summary::S2AuditSummary,
    s3_explanation_export::export_selected_explanations,
    s3_explanation_feedback::plan_explanation_feedback_import,
    s5_config::load_configuration,
    s5_config_check::validate_s5_config,
    s5_outputs::{write_run_outputs, S5OutputRequest, S5OutputSource},
    s5_period_recognition::{detect_s5_period, s5_file_period_hint},
    s5_retrieval::{S5CurrentSourceBinding, S5Retrieval, S5SourceDescriptor},
    s5_service::run_rules,
    s5_template_matching::match_s5_template,
    web_bridge::serve_web_ui,
};
use clap::{Args, Parser, Subcommand};
use rayon::prelude::*;
use serde::{Deserialize, Serialize};
use serde_json::json;
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
    /// Detect an S5 report period from selected workbooks and their source fields.
    S5DetectPeriod(S5DetectPeriodArgs),
    /// Add filename/folder period hints to selected S5 workbooks without opening them.
    S5FilePeriodHints(S5FilePeriodHintsArgs),
    /// Match selected S5 XLSX workbooks to one template by worksheet names.
    S5MatchTemplate(S5MatchTemplateArgs),
    /// Validate S5 configuration rules and current/full-scope template/external addresses.
    S5ConfigCheck(S5ConfigCheckArgs),
    /// Run the S5 configured audit against current and optional previous-period sources.
    S5AuditRun(S5AuditRunArgs),
    /// Run the S2 period-report audit and create its three-sheet result workbook.
    S2(S2Args),
    /// Export S3-F02 comparison rows selected for institution explanation.
    S3ExportExplanations(S3ExportExplanationsArgs),
    /// Match S3-F03 institution feedback and write a read-only ExcelEditPlan.
    #[command(name = "s3-plan-explanation-feedback-import")]
    S3PlanExplanationFeedbackImport(S3PlanExplanationFeedbackImportArgs),
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
    /// Extract standard S1 Issues from formula values already stored in an audit copy.
    S1FormulaAudit(S1FormulaAuditArgs),
    /// Run S1 formula and conditional issue extraction, merge per file, then classify against history.
    S1AuditDirectory(S1AuditDirectoryArgs),
    /// Classify S1 Issues against prior history and carry forward manual notes.
    S1HistoryClassify(S1HistoryClassifyArgs),
    /// Merge XLSX worksheet tables with matching first-row headers for S1-F02.
    S1MergeWorkbookTables(S1MergeWorkbookTablesArgs),
    /// Summarize S1-F02 configured arbitrary-row or fixed-row named regions.
    S1RegionSummary(S1RegionSummaryArgs),
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

#[derive(Debug, Args)]
struct S1FormulaAuditArgs {
    /// S1 template workbook containing named formula ranges or an “审核规则” sheet.
    #[arg(long)]
    template_workbook: PathBuf,

    /// Audit copy with stored formula result values. This command does not recalculate.
    #[arg(long)]
    audit_workbook: PathBuf,

    /// Original report path used for Audit issue identity. Defaults to --audit-workbook.
    #[arg(long)]
    source_file: Option<PathBuf>,

    /// Reporting period recorded on the generated Audit issues.
    #[arg(long)]
    period: String,

    /// S1 run/batch identifier recorded on the generated Audit issues.
    #[arg(long)]
    batch_id: String,

    /// Audit timestamp. Defaults to the current local time.
    #[arg(long)]
    audit_time: Option<String>,

    /// Institution code when it was resolved by the calling workflow.
    #[arg(long, default_value = "")]
    org_code: String,

    /// Institution name when it was resolved by the calling workflow.
    #[arg(long, default_value = "")]
    org_name: String,

    /// Optional Audit config workbook containing the “模块化功能” mappings.
    #[arg(long)]
    mapping_config: Option<PathBuf>,

    /// Emit developer-view JSON progress events with actual engines and timings.
    #[arg(long, default_value_t = false)]
    developer_view: bool,
}

#[derive(Debug, Args)]
struct S1AuditDirectoryArgs {
    /// Directory containing the source .xlsx workbooks.
    #[arg(long)]
    input_dir: PathBuf,

    /// Manifest mapping each structurally accepted source file to its saved formula-result audit copy.
    #[arg(long)]
    formula_manifest: PathBuf,

    /// S1 template workbook containing formula rules and conditional-area mappings.
    #[arg(long, conflicts_with = "template_dir")]
    template_workbook: Option<PathBuf>,

    /// Directory of S1 templates to match against the selected source workbooks.
    #[arg(long, conflicts_with = "template_workbook")]
    template_dir: Option<PathBuf>,

    /// Reporting period; when omitted, detect one from source filenames/directories.
    #[arg(long)]
    period: Option<String>,

    /// S1 run/batch identifier recorded on generated Audit issues.
    #[arg(long)]
    batch_id: String,

    /// Audit timestamp recorded on generated Audit issues. Defaults to local time.
    #[arg(long)]
    audit_time: Option<String>,

    /// Recurse through all child directories when discovering source workbooks.
    #[arg(long, default_value_t = false)]
    recursive: bool,

    /// Restrict discovery to selected source files; repeat the flag as needed.
    #[arg(long = "selected-file")]
    selected_files: Vec<PathBuf>,

    /// Optional Audit config workbook containing the “模块化功能” mappings.
    #[arg(long)]
    mapping_config: Option<PathBuf>,

    /// Optional XLSX/XLSM history workbook used for classification and note inheritance.
    #[arg(long)]
    history_workbook: Option<PathBuf>,

    /// Optional preferred history sheet, falling back to standard and legacy names.
    #[arg(long, requires = "history_workbook")]
    history_sheet: Option<String>,

    /// Optional path for a new XLSX containing the current S1-F01 result table.
    #[arg(long)]
    result_workbook: Option<PathBuf>,

    /// Optional path for a new standalone XLSX containing merged S1 history rows.
    /// The path must not exist; the input history workbook is never modified.
    #[arg(long)]
    history_result_workbook: Option<PathBuf>,

    /// Emit developer-view JSON progress events with actual engines and timings.
    #[arg(long, default_value_t = false)]
    developer_view: bool,
}

#[derive(Debug, Args)]
struct S1HistoryClassifyArgs {
    /// JSON document with period and current_issues; history may also be embedded.
    #[arg(long)]
    input: PathBuf,

    /// Optional XLSX/XLSM history workbook; when supplied it replaces JSON history.
    #[arg(long)]
    history_workbook: Option<PathBuf>,

    /// Optional preferred history sheet, falling back to the standard and legacy sheet names.
    #[arg(long, requires = "history_workbook")]
    history_sheet: Option<String>,

    /// Emit developer-view JSON progress events with stage timing details.
    #[arg(long, default_value_t = false)]
    developer_view: bool,
}

#[derive(Debug, Args)]
struct S1MergeWorkbookTablesArgs {
    /// Directory containing source .xlsx workbooks.
    #[arg(long)]
    input_dir: PathBuf,

    /// New XLSX output path. Existing files are never overwritten.
    #[arg(long)]
    output_workbook: PathBuf,

    /// Feature name used if an output sheet has no source sheet name.
    #[arg(long, default_value = "汇总表合并")]
    feature_name: String,

    /// Search child directories; when nested workbooks exist, root-level files are treated as prior summaries.
    #[arg(long, default_value_t = true, action = clap::ArgAction::Set)]
    recursive: bool,

    /// Emit developer-view JSON progress events with actual engines and timings.
    #[arg(long, default_value_t = false)]
    developer_view: bool,
}

#[derive(Debug, Args)]
struct S1RegionSummaryArgs {
    /// S1 template workbook containing named summary and header ranges.
    #[arg(long)]
    template_workbook: PathBuf,

    /// Directory containing source .xlsx workbooks.
    #[arg(long)]
    input_dir: PathBuf,

    /// New XLSX output path. Existing files are never overwritten.
    #[arg(long)]
    output_workbook: PathBuf,

    /// Optional workbook containing the “模块化功能” feature mappings.
    #[arg(long)]
    mapping_config: Option<PathBuf>,

    /// Optional XLSX workbook containing built-in S1 summary history sheets.
    #[arg(long)]
    history_workbook: Option<PathBuf>,

    /// Feature name from the mapping workbook; repeat to run multiple summaries.
    #[arg(long = "feature-name", required = true)]
    feature_names: Vec<String>,

    /// Search child directories; when nested workbooks exist, root-level files are treated as prior summaries.
    #[arg(long, default_value_t = true, action = clap::ArgAction::Set)]
    recursive: bool,

    /// Emit developer-view JSON progress events with actual engines and timings.
    #[arg(long, default_value_t = false)]
    developer_view: bool,
}

#[derive(Debug, Deserialize)]
struct S1AuditCopyManifest {
    files: Vec<S1AuditCopyManifestFile>,
}

#[derive(Debug, Deserialize)]
struct S1AuditCopyManifestFile {
    source_file: PathBuf,
    audit_workbook: PathBuf,
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

#[derive(Debug, Args)]
struct S5DetectPeriodArgs {
    /// S5 input directory used for the final report-period fallback.
    #[arg(long)]
    input_dir: PathBuf,

    /// Selected S5 input workbook. Repeat once per selected .xlsx/.xlsm file.
    #[arg(long = "input-file", required = true)]
    selected_files: Vec<PathBuf>,

    /// Emit developer-view JSON progress with the reader and elapsed time.
    #[arg(long, default_value_t = false)]
    developer_view: bool,
}

#[derive(Debug, Args)]
struct S5FilePeriodHintsArgs {
    /// Selected S5 input workbook. Repeat once per selected .xlsx/.xlsm file.
    #[arg(long = "input-file", required = true)]
    selected_files: Vec<PathBuf>,

    /// Emit developer-view JSON progress with the path parser and elapsed time.
    #[arg(long, default_value_t = false)]
    developer_view: bool,
}

#[derive(Debug, Args)]
struct S5MatchTemplateArgs {
    /// Directory containing the S5 XLSX templates.
    #[arg(long)]
    template_dir: PathBuf,

    /// Selected S5 source workbook. Repeat once per selected .xlsx/.xlsm file.
    #[arg(long = "input-file", required = true)]
    selected_files: Vec<PathBuf>,

    /// Emit developer-view JSON progress with the reader and elapsed time.
    #[arg(long, default_value_t = false)]
    developer_view: bool,
}

#[derive(Debug, Args)]
struct S5AuditRunArgs {
    /// S5 audit configuration workbook (审核规则 plus optional catalog sheets).
    #[arg(long)]
    config: PathBuf,

    /// Bound current-period source workbook.
    #[arg(long)]
    source: PathBuf,

    /// Explicit layout template used to project S5 numeric regions.
    #[arg(long)]
    template_path: Option<PathBuf>,

    /// Directory of layout templates for automatic numeric-region matching.
    #[arg(long)]
    template_dir: Option<PathBuf>,

    /// Logical source id recorded on retrieval evidence.
    #[arg(long, default_value = "CURRENT")]
    source_id: String,

    /// Business associated with the primary source for Python-compatible fallback selection.
    #[arg(long, default_value = "")]
    source_business: String,

    /// Additional current-period binding in SOURCE_ID=BUSINESS=FILE form; repeat as needed.
    #[arg(long = "current-source", value_name = "SOURCE_ID=BUSINESS=FILE")]
    current_sources: Vec<String>,

    /// Optional previous-period source workbook for V_PRE calls.
    #[arg(long)]
    previous_source: Option<PathBuf>,

    /// Logical source id for the selected previous-period workbook.
    /// Defaults to --source-id.
    #[arg(long)]
    previous_source_id: Option<String>,

    /// Reporting period of --previous-source. Required when that source is selected.
    #[arg(long, requires = "previous_source")]
    previous_period: Option<String>,

    /// Optional configured external source workbook for V_EXT calls.
    #[arg(long)]
    external_source: Option<PathBuf>,

    /// Logical source id for --external-source; required when that source is selected.
    #[arg(long, requires = "external_source")]
    external_source_id: Option<String>,

    /// Reporting period, e.g. 2026-07-31.
    #[arg(long)]
    period: String,

    /// Institution code recorded on results and evidence.
    #[arg(long, default_value = "")]
    org_code: String,

    /// Institution name recorded on results.
    #[arg(long, default_value = "")]
    org_name: String,

    /// Business name to execute. Repeat for multiple businesses; defaults to all.
    #[arg(long = "business")]
    businesses: Vec<String>,

    /// Directory that receives the run outputs (审核副本/审核结果.xlsx/运行总账.json).
    #[arg(long)]
    output_dir: PathBuf,

    /// Stable run identifier; defaults to a timestamp-based id.
    #[arg(long)]
    run_id: Option<String>,

    /// Emit developer-view JSON progress.
    #[arg(long, default_value_t = false)]
    developer_view: bool,
}

#[derive(Debug, Args)]
struct S5ConfigCheckArgs {
    /// S5 configuration workbook.
    #[arg(long)]
    config: PathBuf,

    /// Optional selected template workbook for current-scope address checks.
    #[arg(long)]
    template_path: Option<PathBuf>,

    /// Template directory for full-scope address checks.
    #[arg(long)]
    template_dir: Option<PathBuf>,

    /// Optional selected external workbook for V_EXT header checks.
    #[arg(long)]
    external_path: Option<PathBuf>,

    /// Business to check. Empty checks all enabled rules in the current scope.
    #[arg(long, default_value = "")]
    business: String,

    /// Check scope: current, current_business, all, or all_enabled.
    #[arg(long, default_value = "current")]
    scope: String,

    /// Emit developer-view JSON events with the reader engine and elapsed time.
    #[arg(long, default_value_t = false)]
    developer_view: bool,
}

#[derive(Debug, Args)]
struct S3ExportExplanationsArgs {
    /// S3 comparison workbook containing the output-explanation selection column.
    #[arg(long)]
    comparison_file: PathBuf,

    /// New XLSX path for the institution explanation workbook.
    #[arg(long)]
    output_workbook: PathBuf,

    /// Optional workbook containing the “机构地区参照” sheet.
    #[arg(long)]
    config_workbook: Option<PathBuf>,

    /// Emit developer-view progress with the actual engines and stage timings.
    #[arg(long, default_value_t = false)]
    developer_view: bool,
}

#[derive(Debug, Args)]
struct S3PlanExplanationFeedbackImportArgs {
    /// Existing .xlsx institution explanation workbook. It is read-only in this command.
    #[arg(long)]
    target_workbook: PathBuf,

    /// Institution feedback workbook; repeat this option for multiple files.
    #[arg(long = "feedback-workbook", required = true)]
    feedback_workbooks: Vec<PathBuf>,

    /// New .xlsx path reserved in the plan for a later capability-validated edit.
    #[arg(long)]
    candidate_output_workbook: PathBuf,

    /// New .json path where the ExcelEditPlan and source hashes will be written.
    #[arg(long)]
    plan_output: PathBuf,

    /// Emit developer-view progress with the reader, writer and stage timings.
    #[arg(long, default_value_t = false)]
    developer_view: bool,
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
        Command::S5DetectPeriod(arguments) => {
            run_s5_detect_period(arguments)?;
        }
        Command::S5FilePeriodHints(arguments) => {
            run_s5_file_period_hints(arguments)?;
        }
        Command::S5MatchTemplate(arguments) => {
            run_s5_match_template(arguments)?;
        }
        Command::S5ConfigCheck(arguments) => {
            run_s5_config_check(arguments)?;
        }
        Command::S5AuditRun(arguments) => {
            run_s5_audit_run(arguments)?;
        }
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
            let run_started = Instant::now();
            let mut active_stage = String::new();
            let mut stage_started = run_started;
            let mut on_progress = |event: S2ProgressEvent| {
                if active_stage != event.stage {
                    active_stage.clone_from(&event.stage);
                    stage_started = Instant::now();
                }
                let engine = s2_engine_for_stage(&event.stage);
                let output = serde_json::json!({
                    "type": "progress",
                    "event": event,
                    "engine": engine,
                    "elapsed_ms": run_started.elapsed().as_millis(),
                    "stage_elapsed_ms": stage_started.elapsed().as_millis()
                });
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
        Command::S3ExportExplanations(arguments) => {
            run_s3_export_explanations(arguments)?;
        }
        Command::S3PlanExplanationFeedbackImport(arguments) => {
            run_s3_plan_explanation_feedback_import(arguments)?;
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
        Command::S1FormulaAudit(arguments) => {
            run_s1_formula_audit(arguments)?;
        }
        Command::S1AuditDirectory(arguments) => {
            run_s1_audit_directory(arguments)?;
        }
        Command::S1HistoryClassify(arguments) => {
            run_s1_history_classify(arguments)?;
        }
        Command::S1MergeWorkbookTables(arguments) => {
            run_s1_merge_workbook_tables(arguments)?;
        }
        Command::S1RegionSummary(arguments) => {
            run_s1_region_summary(arguments)?;
        }
    }
    Ok(())
}

#[derive(Debug, Serialize)]
struct S5DetectPeriodOutput<'a> {
    status: &'static str,
    period: &'a str,
    source: &'a str,
    selected_file_count: usize,
    #[serde(skip_serializing_if = "Option::is_none")]
    reader: Option<&'a str>,
    elapsed_ms: u128,
}

#[derive(Debug, Serialize)]
struct S5FilePeriodHintItem {
    path: String,
    period_hint: String,
}

#[derive(Debug, Serialize)]
struct S5FilePeriodHintsOutput {
    status: &'static str,
    selected_file_count: usize,
    reader: &'static str,
    items: Vec<S5FilePeriodHintItem>,
    elapsed_ms: u128,
}

#[derive(Debug, Serialize)]
struct S5TemplateMatchOutput {
    status: &'static str,
    selected_file_count: usize,
    template_path: PathBuf,
    engine: &'static str,
    reader: &'static str,
    elapsed_ms: u128,
}

fn run_s5_file_period_hints(arguments: S5FilePeriodHintsArgs) -> Result<()> {
    for path in &arguments.selected_files {
        anyhow::ensure!(path.is_file(), "S5 输入文件不存在：{}", path.display());
        let extension = path
            .extension()
            .and_then(|value| value.to_str())
            .unwrap_or_default();
        anyhow::ensure!(
            extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm"),
            "S5 文件期次提示只支持 XLSX/XLSM：{}",
            path.display()
        );
    }

    let started = Instant::now();
    if arguments.developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S5-F01",
                "step": "file_period_hints",
                "status": "started",
                "selected_file_count": arguments.selected_files.len(),
                "engine": "rust_path_name_parser",
                "reader": "none",
                "started_at": chrono::Local::now().to_rfc3339()
            }))?
        );
    } else {
        eprintln!(
            "正在识别 S5 文件列表报告期提示：{} 个文件（仅解析文件名和目录，不读取工作簿）…",
            arguments.selected_files.len()
        );
    }

    let items = arguments
        .selected_files
        .iter()
        .map(|path| S5FilePeriodHintItem {
            path: path.to_string_lossy().into_owned(),
            period_hint: s5_file_period_hint(path),
        })
        .collect::<Vec<_>>();
    let output = S5FilePeriodHintsOutput {
        status: "completed",
        selected_file_count: items.len(),
        reader: "none",
        items,
        elapsed_ms: started.elapsed().as_millis(),
    };

    if arguments.developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S5-F01",
                "step": "file_period_hints",
                "status": "completed",
                "selected_file_count": output.selected_file_count,
                "engine": "rust_path_name_parser",
                "reader": output.reader,
                "elapsed_ms": output.elapsed_ms,
                "finished_at": chrono::Local::now().to_rfc3339()
            }))?
        );
        println!("{}", serde_json::to_string_pretty(&output)?);
    } else {
        eprintln!(
            "已完成 {} 个 S5 文件期次提示；引擎 Rust 文件名/目录解析（未读取工作簿），耗时 {} ms。",
            output.selected_file_count, output.elapsed_ms
        );
        println!("{}", serde_json::to_string_pretty(&output)?);
    }
    Ok(())
}

fn run_s5_match_template(arguments: S5MatchTemplateArgs) -> Result<()> {
    anyhow::ensure!(
        arguments.template_dir.is_dir(),
        "请选择存在的模板目录，或使用“更换”手动指定模板"
    );
    for path in &arguments.selected_files {
        anyhow::ensure!(path.is_file(), "S5 输入文件不存在：{}", path.display());
        let extension = path
            .extension()
            .and_then(|value| value.to_str())
            .unwrap_or_default();
        if !extension.eq_ignore_ascii_case("xlsx") && !extension.eq_ignore_ascii_case("xlsm") {
            anyhow::bail!("S5 模板匹配当前只支持 XLSX/XLSM：{}", path.display());
        }
    }

    let started = Instant::now();
    if arguments.developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S5-F01",
                "step": "template_matching",
                "status": "started",
                "selected_file_count": arguments.selected_files.len(),
                "template_dir": &arguments.template_dir,
                "engine": "rust_s5_sheet_set_matcher",
                "reader": "excel_reader_ooxml_sheet_names",
                "started_at": chrono::Local::now().to_rfc3339()
            }))?
        );
    } else {
        eprintln!(
            "正在按工作表结构匹配 S5 模板：{} 个文件（Rust + OOXML 工作表名称读取）…",
            arguments.selected_files.len()
        );
    }

    let matched = match_s5_template(&arguments.selected_files, &arguments.template_dir);
    let elapsed_ms = started.elapsed().as_millis();
    let template_path = match matched {
        Ok(path) => path,
        Err(error) => {
            if arguments.developer_view {
                eprintln!(
                    "{}",
                    serde_json::to_string(&serde_json::json!({
                        "type": "progress",
                        "product": "open-data-audit",
                        "feature": "S5-F01",
                        "step": "template_matching",
                        "status": "failed",
                        "selected_file_count": arguments.selected_files.len(),
                        "engine": "rust_s5_sheet_set_matcher",
                        "reader": "excel_reader_ooxml_sheet_names",
                        "elapsed_ms": elapsed_ms,
                        "error": format!("{error:#}"),
                        "finished_at": chrono::Local::now().to_rfc3339()
                    }))?
                );
            }
            return Err(error);
        }
    };

    let output = S5TemplateMatchOutput {
        status: "completed",
        selected_file_count: arguments.selected_files.len(),
        template_path,
        engine: "rust_s5_sheet_set_matcher",
        reader: "excel_reader_ooxml_sheet_names",
        elapsed_ms,
    };
    if arguments.developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S5-F01",
                "step": "template_matching",
                "status": "completed",
                "selected_file_count": output.selected_file_count,
                "engine": output.engine,
                "reader": output.reader,
                "elapsed_ms": output.elapsed_ms,
                "finished_at": chrono::Local::now().to_rfc3339()
            }))?
        );
        println!("{}", serde_json::to_string_pretty(&output)?);
    } else {
        eprintln!(
            "已为 {} 个 S5 文件匹配模板；引擎 Rust 工作表结构匹配，读取器 OOXML 工作表名称适配器，耗时 {} ms。",
            output.selected_file_count, output.elapsed_ms
        );
        println!("{}", serde_json::to_string_pretty(&output)?);
    }
    Ok(())
}

fn run_s5_detect_period(arguments: S5DetectPeriodArgs) -> Result<()> {
    anyhow::ensure!(
        arguments.input_dir.is_dir(),
        "S5 输入目录不存在或不是目录：{}",
        arguments.input_dir.display()
    );
    for path in &arguments.selected_files {
        anyhow::ensure!(path.is_file(), "S5 输入文件不存在：{}", path.display());
        let extension = path
            .extension()
            .and_then(|value| value.to_str())
            .unwrap_or_default();
        anyhow::ensure!(
            extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm"),
            "S5 报告期识别只支持 XLSX/XLSM：{}",
            path.display()
        );
    }

    let started = Instant::now();
    if arguments.developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S5-F01",
                "step": "report_period_recognition",
                "status": "started",
                "selected_file_count": arguments.selected_files.len(),
                "started_at": chrono::Local::now().to_rfc3339()
            }))?
        );
    } else {
        eprintln!(
            "正在识别 S5 报告期：{} 个已选工作簿…",
            arguments.selected_files.len()
        );
    }

    let result = match detect_s5_period(&arguments.selected_files, &arguments.input_dir) {
        Ok(result) => result,
        Err(error) => {
            if arguments.developer_view {
                eprintln!(
                    "{}",
                    serde_json::to_string(&serde_json::json!({
                        "type": "progress",
                        "product": "open-data-audit",
                        "feature": "S5-F01",
                        "step": "report_period_recognition",
                        "status": "failed",
                        "elapsed_ms": started.elapsed().as_millis(),
                        "error": format!("{error:#}")
                    }))?
                );
            } else {
                eprintln!("S5 报告期识别失败，未启动规则审核。");
            }
            return Err(error);
        }
    };

    let output = S5DetectPeriodOutput {
        status: "completed",
        period: &result.period,
        source: &result.source,
        selected_file_count: result.selected_file_count,
        reader: result.reader.as_deref(),
        elapsed_ms: started.elapsed().as_millis(),
    };
    if arguments.developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S5-F01",
                "step": "report_period_recognition",
                "status": "completed",
                "period": result.period,
                "source": result.source,
                "selected_file_count": result.selected_file_count,
                "reader": result.reader,
                "engine": result.reader.as_deref().unwrap_or("none"),
                "elapsed_ms": output.elapsed_ms,
                "finished_at": chrono::Local::now().to_rfc3339()
            }))?
        );
        println!("{}", serde_json::to_string_pretty(&output)?);
    } else {
        let engine = result.reader.as_deref().unwrap_or("无需读取工作簿");
        eprintln!(
            "已识别报告期 {}（{}）；引擎 {}，耗时 {} ms。",
            result.period, result.source, engine, output.elapsed_ms
        );
        let user_output = serde_json::json!({
            "status": output.status,
            "period": output.period,
            "source": output.source,
            "selected_file_count": output.selected_file_count
        });
        println!("{}", serde_json::to_string_pretty(&user_output)?);
    }
    Ok(())
}

fn run_s3_export_explanations(arguments: S3ExportExplanationsArgs) -> Result<()> {
    let started = Instant::now();
    let started_at = chrono::Local::now().to_rfc3339();
    if arguments.developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S3-F02",
                "step": "selected_explanation_export",
                "status": "started",
                "started_at": started_at
            }))?
        );
    } else {
        eprintln!("正在整理比较结果中标记的机构说明并生成工作簿…");
    }

    let report = match export_selected_explanations(
        &arguments.comparison_file,
        &arguments.output_workbook,
        arguments.config_workbook.as_deref(),
    ) {
        Ok(report) => report,
        Err(error) => {
            if arguments.developer_view {
                eprintln!(
                    "{}",
                    serde_json::to_string(&serde_json::json!({
                        "type": "progress",
                        "product": "open-data-audit",
                        "feature": "S3-F02",
                        "step": "selected_explanation_export",
                        "status": "failed",
                        "finished_at": chrono::Local::now().to_rfc3339(),
                        "elapsed_ms": started.elapsed().as_millis(),
                        "error": format!("{error:#}")
                    }))?
                );
            } else {
                eprintln!("说明文件导出失败，未报告成功结果。");
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
                "feature": "S3-F02",
                "step": "selected_explanation_export",
                "status": "completed",
                "started_at": started_at,
                "finished_at": chrono::Local::now().to_rfc3339(),
                "reader": report.reader,
                "writer": report.writer,
                "selected_rows": report.selected_rows,
                "comparison_read_time_ms": report.comparison_read_time_ms,
                "config_read_time_ms": report.config_read_time_ms,
                "workbook_write_time_ms": report.workbook_write_time_ms,
                "elapsed_ms": started.elapsed().as_millis()
            }))?
        );
    } else {
        eprintln!(
            "说明文件导出完成：{} 行；耗时 {} ms；输出：{}",
            report.selected_rows,
            started.elapsed().as_millis(),
            report.output_workbook.display()
        );
    }

    let mut output = serde_json::to_value(&report)?;
    if !arguments.developer_view {
        if let Some(object) = output.as_object_mut() {
            for key in [
                "reader",
                "writer",
                "comparison_read_time_ms",
                "config_read_time_ms",
                "workbook_write_time_ms",
                "total_time_ms",
            ] {
                object.remove(key);
            }
        }
    }
    println!("{}", serde_json::to_string_pretty(&output)?);
    Ok(())
}

fn run_s3_plan_explanation_feedback_import(
    arguments: S3PlanExplanationFeedbackImportArgs,
) -> Result<()> {
    let started = Instant::now();
    let started_at = chrono::Local::now().to_rfc3339();
    if arguments.developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S3-F03",
                "step": "feedback_matching_and_plan",
                "status": "started",
                "started_at": started_at
            }))?
        );
    } else {
        eprintln!("正在读取机构反馈并生成匹配计划；目标说明工作簿保持不变…");
    }

    let report = match plan_explanation_feedback_import(
        &arguments.target_workbook,
        &arguments.feedback_workbooks,
        &arguments.candidate_output_workbook,
        &arguments.plan_output,
    ) {
        Ok(report) => report,
        Err(error) => {
            if arguments.developer_view {
                eprintln!(
                    "{}",
                    serde_json::to_string(&json!({
                        "type": "progress",
                        "product": "open-data-audit",
                        "feature": "S3-F03",
                        "step": "feedback_matching_and_plan",
                        "status": "failed",
                        "finished_at": chrono::Local::now().to_rfc3339(),
                        "elapsed_ms": started.elapsed().as_millis(),
                        "error": format!("{error:#}")
                    }))?
                );
            } else {
                eprintln!(
                    "机构反馈导入计划生成失败（{} ms），未报告成功结果。",
                    started.elapsed().as_millis()
                );
            }
            return Err(error);
        }
    };

    if arguments.developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S3-F03",
                "step": "feedback_matching_and_plan",
                "status": "completed",
                "started_at": started_at,
                "finished_at": chrono::Local::now().to_rfc3339(),
                "reader": report.reader,
                "business_engine": report.business_engine,
                "conflict_writer": report.conflict_writer,
                "imported": report.imported,
                "unmatched": report.unmatched,
                "conflicts": report.conflicts,
                "target_read_time_ms": report.target_read_time_ms,
                "feedback_read_time_ms": report.feedback_read_time_ms,
                "conflict_write_time_ms": report.conflict_write_time_ms,
                "plan_write_time_ms": report.plan_write_time_ms,
                "elapsed_ms": started.elapsed().as_millis()
            }))?
        );
    } else {
        let conflict_writer = report.conflict_writer.unwrap_or("无需生成冲突清单");
        eprintln!(
            "导入匹配计划已生成：拟写入 {} 行，未匹配 {} 项，冲突 {} 项。目标工作簿未修改；读取器 {}，业务引擎 {}，冲突清单写入器 {}。目标读取 {} ms、反馈读取 {} ms、冲突表写入 {} ms、计划写入 {} ms；总耗时 {} ms。计划：{}。",
            report.imported,
            report.unmatched,
            report.conflicts,
            report.reader,
            report.business_engine,
            conflict_writer,
            report.target_read_time_ms,
            report.feedback_read_time_ms,
            report.conflict_write_time_ms,
            report.plan_write_time_ms,
            started.elapsed().as_millis(),
            report.plan_file.display(),
        );
    }

    let mut output = serde_json::to_value(&report)?;
    if !arguments.developer_view {
        if let Some(object) = output.as_object_mut() {
            for key in [
                "reader",
                "business_engine",
                "conflict_writer",
                "target_read_time_ms",
                "feedback_read_time_ms",
                "conflict_write_time_ms",
                "plan_write_time_ms",
                "total_time_ms",
            ] {
                object.remove(key);
            }
        }
    }
    println!("{}", serde_json::to_string_pretty(&output)?);
    Ok(())
}

fn run_s1_merge_workbook_tables(arguments: S1MergeWorkbookTablesArgs) -> Result<()> {
    let started = Instant::now();
    let started_at = chrono::Local::now().to_rfc3339();
    if arguments.developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S1-F02",
                "step": "workbook_table_merge",
                "status": "started",
                "started_at": started_at
            }))?
        );
    } else {
        eprintln!("正在发现并汇总表头一致的源工作表…");
    }

    let summary = match merge_s1_workbook_tables(
        &arguments.input_dir,
        &arguments.output_workbook,
        &arguments.feature_name,
        arguments.recursive,
    ) {
        Ok(summary) => summary,
        Err(error) => {
            if arguments.developer_view {
                eprintln!(
                    "{}",
                    serde_json::to_string(&serde_json::json!({
                        "type": "progress",
                        "product": "open-data-audit",
                        "feature": "S1-F02",
                        "step": "workbook_table_merge",
                        "status": "failed",
                        "finished_at": chrono::Local::now().to_rfc3339(),
                        "elapsed_ms": started.elapsed().as_millis(),
                        "error": format!("{error:#}")
                    }))?
                );
            } else {
                eprintln!("汇总失败，未生成结果工作簿。");
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
                "feature": "S1-F02",
                "step": "workbook_table_merge",
                "status": "completed",
                "started_at": started_at,
                "finished_at": chrono::Local::now().to_rfc3339(),
                "reader": summary.reader,
                "writer": summary.writer,
                "source_workbooks": summary.source_workbooks,
                "source_sheets": summary.source_sheets,
                "merged_sheets": summary.merged_sheets,
                "data_rows": summary.data_rows,
                "discovery_time_ms": summary.discovery_time_ms,
                "read_time_ms": summary.read_time_ms,
                "grouping_time_ms": summary.grouping_time_ms,
                "write_time_ms": summary.write_time_ms,
                "validation_time_ms": summary.validation_time_ms,
                "elapsed_ms": started.elapsed().as_millis()
            }))?
        );
    } else {
        eprintln!(
            "汇总完成：{} 个工作簿，{} 张结果表，{} 行；耗时 {} ms；输出：{}",
            summary.source_workbooks,
            summary.merged_sheets,
            summary.data_rows,
            started.elapsed().as_millis(),
            summary.output_path.display()
        );
    }
    let mut summary_output = serde_json::to_value(&summary)?;
    if !arguments.developer_view {
        if let Some(object) = summary_output.as_object_mut() {
            object.remove("reader");
            object.remove("writer");
        }
    }
    println!(
        "{}",
        serde_json::to_string_pretty(&serde_json::json!({
            "status": "completed",
            "summary": summary_output
        }))?
    );
    Ok(())
}

fn run_s1_region_summary(arguments: S1RegionSummaryArgs) -> Result<()> {
    let started = Instant::now();
    let started_at = chrono::Local::now().to_rfc3339();
    if arguments.developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S1-F02",
                "step": "named_region_summary",
                "status": "started",
                "started_at": started_at
            }))?
        );
    } else {
        eprintln!("正在读取汇总模板并汇总选定区域…");
    }

    let summary = match summarize_s1_named_regions(
        &arguments.template_workbook,
        &arguments.input_dir,
        &arguments.output_workbook,
        arguments.mapping_config.as_deref(),
        arguments.history_workbook.as_deref(),
        &arguments.feature_names,
        arguments.recursive,
    ) {
        Ok(summary) => summary,
        Err(error) => {
            if arguments.developer_view {
                eprintln!(
                    "{}",
                    serde_json::to_string(&serde_json::json!({
                        "type": "progress",
                        "product": "open-data-audit",
                        "feature": "S1-F02",
                        "step": "named_region_summary",
                        "status": "failed",
                        "finished_at": chrono::Local::now().to_rfc3339(),
                        "elapsed_ms": started.elapsed().as_millis(),
                        "error": format!("{error:#}")
                    }))?
                );
            } else {
                eprintln!("汇总失败，未生成结果工作簿。");
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
                "feature": "S1-F02",
                "step": "named_region_summary",
                "status": "completed",
                "started_at": started_at,
                "finished_at": chrono::Local::now().to_rfc3339(),
                "reader": summary.reader,
                "writer": summary.writer,
                "features": summary.features,
                "source_workbooks": summary.source_workbooks,
                "source_sheets": summary.source_sheets,
                "output_sheets": summary.output_sheets,
                "data_rows": summary.data_rows,
                "skipped_missing_sheets": summary.skipped_missing_sheets,
                "metadata_fields_left_blank": summary.metadata_fields_left_blank,
                "history_reader": summary.history_reader,
                "history_rows_matched": summary.history_rows_matched,
                "discovery_time_ms": summary.discovery_time_ms,
                "template_read_time_ms": summary.template_read_time_ms,
                "source_read_time_ms": summary.source_read_time_ms,
                "calculation_time_ms": summary.calculation_time_ms,
                "history_elapsed_ms": summary.history_elapsed_ms,
                "write_time_ms": summary.write_time_ms,
                "validation_time_ms": summary.validation_time_ms,
                "elapsed_ms": started.elapsed().as_millis()
            }))?
        );
    } else {
        let history_message = summary.history_reader.map_or_else(String::new, |_| {
            format!("；历史追加匹配 {} 行", summary.history_rows_matched)
        });
        eprintln!(
            "汇总完成：{} 个工作簿，{} 张结果表，{} 行{}；耗时 {} ms；输出：{}",
            summary.source_workbooks,
            summary.output_sheets,
            summary.data_rows,
            history_message,
            started.elapsed().as_millis(),
            summary.output_path.display()
        );
    }

    let mut summary_output = serde_json::to_value(&summary)?;
    if !arguments.developer_view {
        if let Some(object) = summary_output.as_object_mut() {
            object.remove("reader");
            object.remove("history_reader");
            object.remove("writer");
        }
    }
    println!(
        "{}",
        serde_json::to_string_pretty(&serde_json::json!({
            "status": "completed",
            "summary": summary_output
        }))?
    );
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
    let health = assess_s1_template_health(&plan, &plan.definition.structure_ranges);

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
                "summary": plan.summary,
                "template_health": health
            }))?
        );
    } else {
        eprintln!(
            "S1 模板规则检查完成：模式 {}，规则 {} 条（启用 {} 条），问题规则 {} 条，结构范围 {} 个，模板体检{}（#REF! 规则 {} 个），读取引擎 {}，读取 {} ms，合计 {} ms。",
            plan.summary.template_mode,
            plan.summary.rules_total,
            plan.summary.rules_enabled,
            health.result_rules,
            health.structure_ranges,
            if health.passed { "通过" } else { "未通过" },
            health.broken_reference_rules,
            plan.summary.engine,
            plan.summary.read_time_ms,
            plan.summary.elapsed_ms
        );
    }

    let mut output = serde_json::json!({
        "status": "completed",
        "summary": plan.summary,
        "template_health": health
    });
    if arguments.include_rule_definitions {
        output["definition"] = serde_json::to_value(plan.definition)?;
    }
    println!("{}", serde_json::to_string_pretty(&output)?);
    Ok(())
}

fn run_s1_formula_audit(arguments: S1FormulaAuditArgs) -> Result<()> {
    let started = Instant::now();
    let started_at = chrono::Local::now().to_rfc3339();
    if arguments.developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S1-F01",
                "step": "formula_result_issue_extraction",
                "status": "started",
                "template_reader": "calamine+namespace_aware_ooxml",
                "result_reader": "calamine",
                "batch_id": arguments.batch_id,
                "started_at": started_at
            }))?
        );
    } else {
        eprintln!("正在读取 S1 模板规则并提取已保存公式结果对应的审核项…");
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
                        "step": "formula_result_issue_extraction",
                        "status": "failed",
                        "stage": "template_rule_read",
                        "finished_at": chrono::Local::now().to_rfc3339(),
                        "elapsed_ms": started.elapsed().as_millis(),
                        "error": format!("{error:#}")
                    }))?
                );
            } else {
                eprintln!("S1 模板规则读取失败：{error:#}");
            }
            return Err(error);
        }
    };
    let metadata = S1FormulaMetadata {
        period: arguments.period,
        batch_id: arguments.batch_id,
        audit_time: arguments
            .audit_time
            .unwrap_or_else(|| chrono::Local::now().format("%Y-%m-%d %H:%M:%S").to_string()),
        org_code: arguments.org_code,
        org_name: arguments.org_name,
    };
    let result = match extract_s1_formula_issues(
        &plan.definition,
        &arguments.audit_workbook,
        arguments.source_file.as_deref(),
        &metadata,
    ) {
        Ok(result) => result,
        Err(error) => {
            if arguments.developer_view {
                eprintln!(
                    "{}",
                    serde_json::to_string(&serde_json::json!({
                        "type": "progress",
                        "product": "open-data-audit",
                        "feature": "S1-F01",
                        "step": "formula_result_issue_extraction",
                        "status": "failed",
                        "stage": "stored_result_read_or_issue_build",
                        "batch_id": metadata.batch_id,
                        "finished_at": chrono::Local::now().to_rfc3339(),
                        "elapsed_ms": started.elapsed().as_millis(),
                        "error": format!("{error:#}")
                    }))?
                );
            } else {
                eprintln!("S1 公式结果提取失败，未生成审核项：{error:#}");
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
                "step": "formula_result_issue_extraction",
                "status": "completed",
                "batch_id": metadata.batch_id,
                "started_at": started_at,
                "finished_at": chrono::Local::now().to_rfc3339(),
                "elapsed_ms": started.elapsed().as_millis(),
                "template_reader": plan.summary.engine,
                "template_read_time_ms": plan.summary.read_time_ms,
                "result_reader": result.summary.reader,
                "rules_total": result.summary.rules_total,
                "rules_enabled": result.summary.rules_enabled,
                "rules_selected": result.summary.rules_selected,
                "formula_result_values_nonempty": result.summary.formula_result_values_nonempty,
                "issues": result.summary.issue_count,
                "result_read_time_ms": result.summary.read_time_ms,
                "issue_build_time_ms": result.summary.issue_build_time_ms,
                "total_time_ms": started.elapsed().as_millis()
            }))?
        );
    } else {
        eprintln!(
            "S1 公式结果提取完成：模板规则 {} 条，选中 {} 条，读取器 {}，非空结果值 {} 格，生成审核项 {} 条；模板读取 {} ms、工作簿读取 {} ms、组装 {} ms、实际耗时 {} ms。未执行公式重算。",
            result.summary.rules_total,
            result.summary.rules_selected,
            result.summary.reader,
            result.summary.formula_result_values_nonempty,
            result.summary.issue_count,
            plan.summary.read_time_ms,
            result.summary.read_time_ms,
            result.summary.issue_build_time_ms,
            started.elapsed().as_millis()
        );
    }

    let output = serde_json::json!({
        "status": result.status,
        "template_summary": plan.summary,
        "summary": result.summary,
        "issues": result.issues
    });
    println!("{}", serde_json::to_string_pretty(&output)?);
    Ok(())
}

fn run_s1_audit_directory(arguments: S1AuditDirectoryArgs) -> Result<()> {
    let run_started = Instant::now();
    let started_at = chrono::Local::now().to_rfc3339();
    let input_dir = fs::canonicalize(&arguments.input_dir)
        .with_context(|| format!("读取 S1 输入目录 {}", arguments.input_dir.display()))?;
    anyhow::ensure!(
        input_dir.is_dir(),
        "S1 输入路径不是目录：{}",
        input_dir.display()
    );
    let result_workbook_path = arguments
        .result_workbook
        .as_deref()
        .map(prepare_s1_result_workbook_path)
        .transpose()?;
    let history_result_workbook_path = arguments
        .history_result_workbook
        .as_deref()
        .map(prepare_s1_history_result_workbook_path)
        .transpose()?;
    anyhow::ensure!(
        arguments.template_workbook.is_some() || arguments.template_dir.is_some(),
        "必须指定 --template-workbook 或 --template-dir"
    );
    let explicit_template_workbook = arguments
        .template_workbook
        .as_deref()
        .map(|path| {
            fs::canonicalize(path).with_context(|| format!("读取 S1 模板 {}", path.display()))
        })
        .transpose()?;
    let template_dir = arguments
        .template_dir
        .as_deref()
        .map(|path| {
            fs::canonicalize(path).with_context(|| format!("读取 S1 模板目录 {}", path.display()))
        })
        .transpose()?;
    if let Some(template_dir) = template_dir.as_deref() {
        anyhow::ensure!(
            template_dir.is_dir(),
            "S1 模板路径不是目录：{}",
            template_dir.display()
        );
    }
    let manifest_path = fs::canonicalize(&arguments.formula_manifest).with_context(|| {
        format!(
            "读取公式审核副本清单 {}",
            arguments.formula_manifest.display()
        )
    })?;
    let mapping_config = arguments.mapping_config.as_deref();

    let discovery_started = Instant::now();
    let depth = if arguments.recursive {
        ScanDepth::Full
    } else {
        ScanDepth::RootOnly
    };
    let excluded = excluded_s1_input_workbooks(
        explicit_template_workbook.as_deref(),
        mapping_config,
        template_dir.as_deref(),
    )?;
    let source_files = source_workbooks(&input_dir, depth)?
        .into_iter()
        .filter(|path| {
            !excluded
                .iter()
                .any(|candidate| same_file_path(path, candidate))
                && result_workbook_path
                    .as_ref()
                    .is_none_or(|output| !same_file_path(path, output))
                && history_result_workbook_path
                    .as_ref()
                    .is_none_or(|output| !same_file_path(path, output))
        })
        .collect::<Vec<_>>();
    let available_source_count = source_files.len();
    let source_files =
        select_s1_source_workbooks(&input_dir, &source_files, &arguments.selected_files)?;
    anyhow::ensure!(
        !source_files.is_empty(),
        "源数据目录中没有可审核的 .xlsx 文件（模板和映射配置不会作为源文件）"
    );
    let detection = detect_period_from_workbooks(&input_dir, &source_files);
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
    let audit_time = arguments
        .audit_time
        .clone()
        .unwrap_or_else(|| chrono::Local::now().format("%Y-%m-%d %H:%M:%S").to_string());
    if arguments.developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S1-F01",
                "step": "source_workbook_discovery",
                "status": "completed",
                "started_at": started_at,
                "finished_at": chrono::Local::now().to_rfc3339(),
                "files_discovered": available_source_count,
                "files_selected": source_files.len(),
                "selection_applied": !arguments.selected_files.is_empty(),
                "period": period,
                "period_source": period_source,
                "elapsed_ms": discovery_started.elapsed().as_millis()
            }))?
        );
    } else {
        eprintln!(
            "S1-F01 输入发现完成：可审核 XLSX {} 个，纳入本次审核 {} 个，数据期 {}（{}），耗时 {} ms。",
            available_source_count,
            source_files.len(),
            period,
            period_source,
            discovery_started.elapsed().as_millis()
        );
    }

    let (template_workbook, template_selection_mode) = if let Some(template_workbook) =
        explicit_template_workbook
    {
        (template_workbook, "explicit")
    } else {
        let template_dir = template_dir.as_deref().context("S1 模板目录参数丢失")?;
        let suggested =
            recommend_s1_directory_template(template_dir, &source_files, arguments.developer_view)?;
        let template_workbook = fs::canonicalize(&suggested)
            .with_context(|| format!("读取自动匹配的 S1 模板 {}", suggested.display()))?;
        (template_workbook, "automatic")
    };

    let template_started = Instant::now();
    let template_plan = load_s1_template_rule_plan(&template_workbook, mapping_config)
        .context("读取 S1 模板公式规则")?;
    let (scan_areas, structure_ranges) = resolve_s1_conditional_cli_areas(
        Some(&template_workbook),
        mapping_config,
        &[],
        &[],
        arguments.developer_view,
    )?;
    let health_started = Instant::now();
    let health_started_at = chrono::Local::now().to_rfc3339();
    let template_health = assess_s1_template_health(&template_plan, &structure_ranges);
    let health_error_count = template_health
        .items
        .iter()
        .filter(|item| item.level == "错误")
        .count();
    let health_warning_count = template_health
        .items
        .iter()
        .filter(|item| item.level == "警告")
        .count();
    if arguments.developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S1-F01",
                "step": "template_health_preflight",
                "status": if template_health.passed { "completed" } else { "failed" },
                "engine": template_plan.summary.engine,
                "started_at": health_started_at,
                "finished_at": chrono::Local::now().to_rfc3339(),
                "enabled_rules": template_health.enabled_rules,
                "result_rules": template_health.result_rules,
                "structure_ranges": template_health.structure_ranges,
                "broken_reference_rules": template_health.broken_reference_rules,
                "errors": health_error_count,
                "warnings": health_warning_count,
                "items": template_health.items,
                "elapsed_ms": health_started.elapsed().as_millis()
            }))?
        );
    } else if template_health.passed {
        eprintln!(
            "S1 模板体检通过：启用规则 {} 条，问题规则 {} 条，结构范围 {} 个，#REF! 规则 {} 个；耗时 {} ms。",
            template_health.enabled_rules,
            template_health.result_rules,
            template_health.structure_ranges,
            template_health.broken_reference_rules,
            health_started.elapsed().as_millis()
        );
    } else {
        eprintln!(
            "S1 模板体检未通过：存在 {} 个错误项、{} 个警告项；已停止审核。",
            health_error_count, health_warning_count
        );
    }
    anyhow::ensure!(
        template_health.passed,
        "S1 模板体检不通过；为避免遗漏公式审核或无法核对报送文件结构，已停止审核"
    );
    if arguments.developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S1-F01",
                "step": "template_rule_and_area_load",
                "status": "completed",
                "template_reader": template_plan.summary.engine,
                "formula_rules_enabled": template_plan.summary.rules_enabled,
                "formula_rules_selected": template_plan.summary.rules_enabled,
                "conditional_scan_areas": scan_areas.len(),
                "structure_ranges": structure_ranges.len(),
                "elapsed_ms": template_started.elapsed().as_millis()
            }))?
        );
    } else {
        eprintln!(
            "S1 模板规则准备完成：启用公式规则 {} 条，条件扫描区域 {} 个，结构区域 {} 个；耗时 {} ms。",
            template_plan.summary.rules_enabled,
            scan_areas.len(),
            structure_ranges.len(),
            template_started.elapsed().as_millis()
        );
    }

    let all_inputs = source_files
        .iter()
        .map(|path| S1ConditionalBatchInput {
            workbook_path: path.clone(),
            source_file: path.clone(),
            metadata: S1ConditionalMetadata {
                period: period.clone(),
                batch_id: arguments.batch_id.clone(),
                audit_time: audit_time.clone(),
                org_code: String::new(),
                org_name: String::new(),
            },
        })
        .collect::<Vec<_>>();
    let (accepted_inputs, structure_preflight) = prepare_s1_structure_inputs(
        all_inputs,
        Some(&template_workbook),
        &scan_areas,
        &structure_ranges,
        arguments.developer_view,
    )?;
    anyhow::ensure!(
        !accepted_inputs.is_empty(),
        "没有通过模板表结构预检的源文件；未运行公式或条件格式审核"
    );

    let formula_pairs = read_s1_audit_copy_manifest(&manifest_path, &input_dir, &accepted_inputs)?;
    let pair_by_source = formula_pairs
        .iter()
        .map(|pair| {
            (
                s1_path_identity(&pair.source_file),
                pair.audit_workbook.clone(),
            )
        })
        .collect::<std::collections::HashMap<_, _>>();

    let conditional_started = Instant::now();
    if arguments.developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S1-F01",
                "step": "conditional_format_issue_extraction",
                "status": "started",
                "engine": "calamine+namespace_aware_ooxml",
                "files_total": accepted_inputs.len(),
                "started_at": chrono::Local::now().to_rfc3339()
            }))?
        );
    } else {
        eprintln!("正在对通过结构预检的源文件执行条件格式审核…");
    }
    let conditional_report = extract_conditional_issues_batch(
        &accepted_inputs,
        &scan_areas,
        &structure_ranges,
        |progress: &S1ConditionalBatchProgress| {
            if arguments.developer_view {
                if let Ok(event) = serde_json::to_string(&serde_json::json!({
                    "type": "progress",
                    "product": "open-data-audit",
                    "feature": "S1-F01",
                    "step": "conditional_format_issue_extraction",
                    "batch_id": arguments.batch_id,
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
                        "{}：规则 {}/{}，扫描 {} 格，形成 {} 条审核项；读取 {} ms、计算 {} ms、总计 {} ms。",
                        progress.source_file,
                        progress.rules_supported.unwrap_or_default(),
                        progress.rules_total.unwrap_or_default(),
                        progress.scanned_cells.unwrap_or_default(),
                        progress.issue_count.unwrap_or_default(),
                        progress.read_time_ms.unwrap_or_default(),
                        progress.evaluation_time_ms.unwrap_or_default(),
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
    .context("执行 S1 条件格式审核")?;
    anyhow::ensure!(
        conditional_report.status == "completed",
        "条件格式审核包含未支持规则或未确认的业务上下文，已停止历史归类（{} 个文件诊断）",
        conditional_report
            .files
            .iter()
            .map(|file| file.diagnostics.len() + file.unresolved_indicator_count)
            .sum::<usize>()
    );
    if arguments.developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S1-F01",
                "step": "conditional_format_issue_extraction",
                "status": "completed",
                "engine": "calamine+namespace_aware_ooxml",
                "files": conditional_report.files.len(),
                "issues": conditional_report.issues.len(),
                "elapsed_ms": conditional_started.elapsed().as_millis()
            }))?
        );
    } else {
        eprintln!(
            "条件格式审核完成：{} 个工作簿，{} 条审核项；耗时 {} ms。",
            conditional_report.files.len(),
            conditional_report.issues.len(),
            conditional_started.elapsed().as_millis()
        );
    }

    let formula_started = Instant::now();
    let formula_results = accepted_inputs
        .par_iter()
        .map(|input| {
            let source_key = s1_path_identity(&input.source_file);
            let audit_workbook = pair_by_source
                .get(&source_key)
                .context("公式审核副本清单与结构预检结果不一致")?;
            let (org_code, org_name) = organisation_from_name(&input.source_file);
            let result = extract_s1_formula_issues(
                &template_plan.definition,
                audit_workbook,
                Some(&input.source_file),
                &S1FormulaMetadata {
                    period: period.clone(),
                    batch_id: arguments.batch_id.clone(),
                    audit_time: audit_time.clone(),
                    org_code,
                    org_name,
                },
            )
            .with_context(|| {
                format!(
                    "提取公式审核项失败：{}",
                    input
                        .source_file
                        .file_name()
                        .unwrap_or_default()
                        .to_string_lossy()
                )
            })?;
            Ok((result.issues, result.summary))
        })
        .collect::<Result<Vec<_>>>()?;
    let mut formula_issues = formula_results
        .iter()
        .flat_map(|(issues, _)| issues.iter().cloned())
        .collect::<Vec<_>>();
    let formula_result_issue_count = formula_issues.len();
    let formula_nonempty_results = formula_results
        .iter()
        .map(|(_, summary)| summary.formula_result_values_nonempty)
        .sum::<usize>();
    if arguments.developer_view {
        let per_file = accepted_inputs
            .iter()
            .zip(formula_results.iter())
            .map(|(input, (_, summary))| {
                serde_json::json!({
                    "source_file": input.source_file.file_name().map(|name| name.to_string_lossy()),
                    "reader": summary.reader,
                    "rules_selected": summary.rules_selected,
                    "formula_result_values_nonempty": summary.formula_result_values_nonempty,
                    "issues": summary.issue_count,
                    "read_time_ms": summary.read_time_ms,
                    "issue_build_time_ms": summary.issue_build_time_ms,
                    "total_time_ms": summary.total_time_ms
                })
            })
            .collect::<Vec<_>>();
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S1-F01",
                "step": "formula_result_issue_extraction",
                "status": "completed",
                "engine": "calamine",
                "files": formula_results.len(),
                "rules_enabled_per_file": template_plan.summary.rules_enabled,
                "formula_result_values_nonempty": formula_nonempty_results,
                "issues": formula_result_issue_count,
                "per_file": per_file,
                "elapsed_ms": formula_started.elapsed().as_millis(),
                "formula_recalculation": "not performed; saved values only"
            }))?
        );
    } else {
        eprintln!(
            "公式结果审核完成：{} 个工作簿，规则 {} 条/份，非空结果值 {} 格，形成 {} 条审核项；读取与组装耗时 {} ms。未执行公式重算。",
            formula_results.len(),
            template_plan.summary.rules_enabled,
            formula_nonempty_results,
            formula_result_issue_count,
            formula_started.elapsed().as_millis()
        );
    }

    let validation_scan_started = Instant::now();
    let validation_plan_result =
        S1ValidationFormulaScanPlan::load(std::slice::from_ref(&template_workbook));
    let (validation_plan, validation_plan_error) = match validation_plan_result {
        Ok(plan) => (Some(plan), None),
        Err(error) => (None, Some(format!("{error:#}"))),
    };
    let validation_template_read_time_ms = validation_plan
        .as_ref()
        .map(S1ValidationFormulaScanPlan::template_read_time_ms)
        .unwrap_or_default();
    let validation_scan_results = accepted_inputs
        .par_iter()
        .map(|input| {
            let source_key = s1_path_identity(&input.source_file);
            let audit_workbook = pair_by_source
                .get(&source_key)
                .context("校验公式扫描清单与结构预检结果不一致")?;
            let (org_code, org_name) = organisation_from_name(&input.source_file);
            let metadata = S1FormulaMetadata {
                period: period.clone(),
                batch_id: arguments.batch_id.clone(),
                audit_time: audit_time.clone(),
                org_code,
                org_name,
            };
            let scan_started = Instant::now();
            let scan_result = match (&validation_plan, validation_plan_error.as_deref()) {
                (Some(plan), _) => plan.scan(&input.source_file),
                (None, Some(error)) => Err(anyhow::anyhow!(error.to_owned())),
                (None, None) => Err(anyhow::anyhow!("未能准备校验公式扫描计划")),
            };
            match scan_result {
                Ok(report) => {
                    let mut issues = issues_from_validation_formula_scan(
                        &report,
                        &input.source_file,
                        audit_workbook,
                        &metadata,
                    );
                    let duplicate_issues_suppressed = issues.len()
                        - issues
                            .iter()
                            .filter(|candidate| {
                                !formula_issues.iter().any(|existing| {
                                    existing.issue_id == candidate.issue_id
                                        && existing.formula_result == candidate.formula_result
                                        && existing.source_file == candidate.source_file
                                })
                            })
                            .count();
                    issues.retain(|candidate| {
                        !formula_issues.iter().any(|existing| {
                            existing.issue_id == candidate.issue_id
                                && existing.formula_result == candidate.formula_result
                                && existing.source_file == candidate.source_file
                        })
                    });
                    Ok((
                        Some(report),
                        issues,
                        None,
                        duplicate_issues_suppressed,
                        scan_started.elapsed().as_millis(),
                    ))
                }
                Err(error) => Ok((
                    None,
                    vec![validation_formula_scan_failure_issue(
                        &input.source_file,
                        audit_workbook,
                        &metadata,
                        &format!("{error:#}"),
                    )],
                    Some(format!("{error:#}")),
                    0,
                    scan_started.elapsed().as_millis(),
                )),
            }
        })
        .collect::<Result<Vec<_>>>()?;
    let validation_scan_issue_count = validation_scan_results
        .iter()
        .map(|(_, issues, _, _, _)| issues.len())
        .sum::<usize>();
    let validation_scan_error_count = validation_scan_results
        .iter()
        .map(|(report, _, _, _, _)| {
            report
                .as_ref()
                .map(|report| report.errors.len())
                .unwrap_or_default()
        })
        .sum::<usize>();
    let validation_scan_formula_cells = validation_scan_results
        .iter()
        .map(|(report, _, _, _, _)| {
            report
                .as_ref()
                .map(|report| report.formula_cells_scanned)
                .unwrap_or_default()
        })
        .sum::<usize>();
    let validation_scan_missing_sheets = validation_scan_results
        .iter()
        .map(|(report, _, _, _, _)| {
            report
                .as_ref()
                .map(|report| report.missing_sheets.len())
                .unwrap_or_default()
        })
        .sum::<usize>();
    let validation_scan_failures = validation_scan_results
        .iter()
        .filter(|(_, _, error, _, _)| error.is_some())
        .count();
    let validation_scan_duplicates_suppressed = validation_scan_results
        .iter()
        .map(|(_, _, _, count, _)| count)
        .sum::<usize>();
    let validation_scan_per_file = accepted_inputs
        .iter()
        .zip(validation_scan_results.iter())
        .map(|(input, (report, issues, error, _, elapsed_ms))| {
            json!({
                "source_file": input.source_file.file_name().map(|name| name.to_string_lossy()),
                "status": if error.is_some() { "failed" } else { report.as_ref().map(|report| report.status).unwrap_or("skipped") },
                "formula_cells_scanned": report.as_ref().map(|report| report.formula_cells_scanned).unwrap_or_default(),
                "source_errors": report.as_ref().map(|report| report.errors.len()).unwrap_or_default(),
                "missing_sheets": report.as_ref().map(|report| report.missing_sheets.len()).unwrap_or_default(),
                "issues_added": issues.len(),
                "elapsed_ms": elapsed_ms
            })
        })
        .collect::<Vec<_>>();
    let validation_scan_error_details = accepted_inputs
        .iter()
        .zip(validation_scan_results.iter())
        .flat_map(|(input, (report, _, _, _, _))| {
            report
                .as_ref()
                .into_iter()
                .flat_map(|report| report.errors.iter())
                .map(|error| {
                    json!({
                        "source_file": input.source_file.file_name().map(|name| name.to_string_lossy()),
                        "sheet": error.sheet,
                        "cell": error.cell,
                        "value": error.value,
                        "formula": error.formula,
                        "origin": error.origin,
                        "template_paths": error.template_paths.iter().map(|path| Path::new(path).file_name().map(|name| name.to_string_lossy())).collect::<Vec<_>>(),
                        "formula_error_tokens": error.formula_error_tokens
                    })
                })
        })
        .collect::<Vec<_>>();
    for (_, issues, _, _, _) in &validation_scan_results {
        formula_issues.extend(issues.iter().cloned());
    }
    let formula_issue_count = formula_issues.len();
    let validation_formula_scan_summary = json!({
        "status": if validation_scan_failures == 0 { "completed" } else { "completed_with_diagnostics" },
        "reader": validation_plan.as_ref().map(|_| "calamine+namespace_aware_ooxml_defined_names").unwrap_or("calamine+namespace_aware_ooxml_defined_names"),
        "template_count": validation_plan.as_ref().map(S1ValidationFormulaScanPlan::template_count).unwrap_or(1),
        "formula_cells_scanned": validation_scan_formula_cells,
        "formula_errors": validation_scan_error_count,
        "errors": validation_scan_error_details,
        "missing_sheets": validation_scan_missing_sheets,
        "issues_added": validation_scan_issue_count,
        "duplicates_suppressed": validation_scan_duplicates_suppressed,
        "failures": validation_scan_failures,
        "template_read_time_ms": validation_template_read_time_ms,
        "elapsed_ms": validation_scan_started.elapsed().as_millis(),
        "per_file": validation_scan_per_file
    });
    if arguments.developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S1-F01",
                "step": "validation_formula_error_scan",
                "status": validation_formula_scan_summary["status"],
                "engine": "calamine+namespace_aware_ooxml_defined_names",
                "template_read_time_ms": validation_template_read_time_ms,
                "formula_cells_scanned": validation_scan_formula_cells,
                "formula_errors": validation_scan_error_count,
                "missing_sheets": validation_scan_missing_sheets,
                "issues_added": validation_scan_issue_count,
                "duplicates_suppressed": validation_scan_duplicates_suppressed,
                "failures": validation_scan_failures,
                "per_file": validation_scan_per_file,
                "elapsed_ms": validation_scan_started.elapsed().as_millis()
            }))?
        );
    } else {
        eprintln!(
            "S1 校验公式扫描完成：扫描公式 {} 个，发现错误 {} 项、缺失工作表 {} 项，新增审核项 {} 条；失败 {} 个文件，耗时 {} ms。",
            validation_scan_formula_cells,
            validation_scan_error_count,
            validation_scan_missing_sheets,
            validation_scan_issue_count,
            validation_scan_failures,
            validation_scan_started.elapsed().as_millis()
        );
    }

    let merge_started = Instant::now();
    let merged_issues =
        merge_s1_issue_sources(&formula_pairs, &formula_issues, &conditional_report.issues)
            .context("按源文件顺序合流公式与条件格式审核项")?;

    let history_started = Instant::now();
    let history_read = if let Some(history_path) = arguments.history_workbook.as_deref() {
        Some(
            read_s1_history_workbook(history_path, arguments.history_sheet.as_deref())
                .with_context(|| format!("读取 S1 历史工作簿 {}", history_path.display()))?,
        )
    } else {
        None
    };
    if let Some(loaded) = history_read.as_ref() {
        if arguments.developer_view {
            eprintln!(
                "{}",
                serde_json::to_string(&serde_json::json!({
                    "type": "progress",
                    "product": "open-data-audit",
                    "feature": "S1-F01",
                    "step": "history_workbook_read",
                    "status": loaded.summary.status,
                    "engine": loaded.summary.reader,
                    "sheet_name": loaded.summary.sheet_name,
                    "rows_total": loaded.summary.rows_total,
                    "rows_read": loaded.summary.rows_read,
                    "rows_skipped": loaded.summary.rows_skipped,
                    "elapsed_ms": loaded.summary.elapsed_ms
                }))?
            );
        } else {
            eprintln!(
                "S1 历史工作簿读取完成：状态 {}，工作表 {}，读取 {}/{} 行、跳过 {} 行；耗时 {} ms。",
                loaded.summary.status,
                loaded.summary.sheet_name.as_deref().unwrap_or("未选择"),
                loaded.summary.rows_read,
                loaded.summary.rows_total,
                loaded.summary.rows_skipped,
                loaded.summary.elapsed_ms
            );
        }
    }
    let history_rows = history_read
        .as_ref()
        .map(|loaded| loaded.issues.clone())
        .unwrap_or_default();
    let history_read_summary = history_read.map(|loaded| loaded.summary);
    let mut history_result = process_s1_history(S1HistoryRequest {
        period: period.clone(),
        batch_id: arguments.batch_id.clone(),
        audit_time: audit_time.clone(),
        audited_org_codes: None,
        current_issues: merged_issues,
        history: history_rows,
    });
    history_result.summary.elapsed_ms = history_started.elapsed().as_millis();

    if arguments.developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S1-F01",
                "step": "issue_merge_and_history_classification",
                "status": "completed",
                "engine": "rust_s1_issue_chain+rust_s1_history",
                "files": accepted_inputs.len(),
                "formula_issues": formula_issue_count,
                "conditional_issues": conditional_report.issues.len(),
                "current_issues": history_result.summary.current_issues,
                "resolved_issues": history_result.summary.resolved_issues,
                "history_rows": history_result.summary.history_rows,
                "merged_history_rows": history_result.summary.merged_history_rows,
                "elapsed_ms": merge_started.elapsed().as_millis()
            }))?
        );
    } else {
        eprintln!(
            "S1 Issue 合流与历史归类完成：公式 {} 条、条件格式 {} 条，本期 {} 条（新增 {}、连续 {}、再次出现 {}），已整改 {} 条；耗时 {} ms。",
            formula_issue_count,
            conditional_report.issues.len(),
            history_result.summary.current_issues,
            history_result.summary.new_issues,
            history_result.summary.continuous_issues,
            history_result.summary.recurring_issues,
            history_result.summary.resolved_issues,
            merge_started.elapsed().as_millis()
        );
    }

    let result_workbook_summary = if let Some(path) = result_workbook_path.as_deref() {
        let writer_started_at = chrono::Local::now().to_rfc3339();
        let writer_started = Instant::now();
        if arguments.developer_view {
            eprintln!(
                "{}",
                serde_json::to_string(&serde_json::json!({
                    "type": "progress",
                    "product": "open-data-audit",
                    "feature": "S1-F01",
                    "step": "result_workbook_write",
                    "status": "started",
                    "engine": "rust_xlsxwriter 0.101.0",
                    "started_at": writer_started_at,
                    "issues": history_result.current_issues.len()
                }))?
            );
        } else {
            eprintln!("正在生成 S1-F01 本期结果工作簿…");
        }
        match write_s1_result_workbook(path, None, &history_result.current_issues) {
            Ok(summary) => {
                if arguments.developer_view {
                    eprintln!(
                        "{}",
                        serde_json::to_string(&serde_json::json!({
                            "type": "progress",
                            "product": "open-data-audit",
                            "feature": "S1-F01",
                            "step": "result_workbook_write",
                            "status": "completed",
                            "engine": summary.engine,
                            "started_at": writer_started_at,
                            "finished_at": chrono::Local::now().to_rfc3339(),
                            "elapsed_ms": summary.elapsed_ms,
                            "issues": summary.issue_count,
                            "hyperlinks": summary.hyperlink_count,
                            "path": summary.path
                        }))?
                    );
                } else {
                    eprintln!(
                        "S1-F01 本期结果工作簿已生成：{} 条审核项，{} 个审核副本链接；耗时 {} ms；路径 {}。",
                        summary.issue_count,
                        summary.hyperlink_count,
                        summary.elapsed_ms,
                        summary.path.display()
                    );
                }
                Some(summary)
            }
            Err(error) => {
                if arguments.developer_view {
                    eprintln!(
                        "{}",
                        serde_json::to_string(&serde_json::json!({
                            "type": "progress",
                            "product": "open-data-audit",
                            "feature": "S1-F01",
                            "step": "result_workbook_write",
                            "status": "failed",
                            "engine": "rust_xlsxwriter 0.101.0",
                            "elapsed_ms": writer_started.elapsed().as_millis(),
                            "error": format!("{error:#}")
                        }))?
                    );
                } else {
                    eprintln!(
                        "S1-F01 本期结果工作簿生成失败；耗时 {} ms：{error:#}",
                        writer_started.elapsed().as_millis()
                    );
                }
                return Err(error.context("S1-F01 结果工作簿生成失败"));
            }
        }
    } else {
        None
    };

    let history_result_workbook_summary =
        if let Some(path) = history_result_workbook_path.as_deref() {
            let writer_started_at = chrono::Local::now().to_rfc3339();
            let writer_started = Instant::now();
            if arguments.developer_view {
                eprintln!(
                    "{}",
                    serde_json::to_string(&serde_json::json!({
                        "type": "progress",
                        "product": "open-data-audit",
                        "feature": "S1-F01",
                        "step": "history_result_workbook_write",
                        "status": "started",
                        "engine": "rust_xlsxwriter 0.101.0",
                        "started_at": writer_started_at,
                        "issues": history_result.merged_history.len()
                    }))?
                );
            } else {
                eprintln!("正在生成 S1-F01 历史结果工作簿…");
            }
            match write_s1_history_workbook(path, &history_result.merged_history) {
                Ok(summary) => {
                    if arguments.developer_view {
                        eprintln!(
                            "{}",
                            serde_json::to_string(&serde_json::json!({
                                "type": "progress",
                                "product": "open-data-audit",
                                "feature": "S1-F01",
                                "step": "history_result_workbook_write",
                                "status": "completed",
                                "engine": summary.engine,
                                "started_at": writer_started_at,
                                "finished_at": chrono::Local::now().to_rfc3339(),
                                "elapsed_ms": summary.elapsed_ms,
                                "issues": summary.issue_count,
                                "path": summary.path
                            }))?
                        );
                    } else {
                        eprintln!(
                            "S1-F01 历史结果工作簿已生成：{} 条历史审核项；耗时 {} ms；路径 {}。",
                            summary.issue_count,
                            summary.elapsed_ms,
                            summary.path.display()
                        );
                    }
                    Some(summary)
                }
                Err(error) => {
                    if arguments.developer_view {
                        eprintln!(
                            "{}",
                            serde_json::to_string(&serde_json::json!({
                                "type": "progress",
                                "product": "open-data-audit",
                                "feature": "S1-F01",
                                "step": "history_result_workbook_write",
                                "status": "failed",
                                "engine": "rust_xlsxwriter 0.101.0",
                                "elapsed_ms": writer_started.elapsed().as_millis(),
                                "error": format!("{error:#}")
                            }))?
                        );
                    } else {
                        eprintln!(
                            "S1-F01 历史结果工作簿生成失败；耗时 {} ms：{error:#}",
                            writer_started.elapsed().as_millis()
                        );
                    }
                    return Err(error.context("S1-F01 历史结果工作簿生成失败"));
                }
            }
        } else {
            None
        };

    let status = if structure_preflight
        .as_ref()
        .is_some_and(|summary| summary.files_rejected > 0)
    {
        "completed_with_diagnostics"
    } else {
        "completed"
    };
    let mut output = serde_json::json!({
        "status": status,
        "template_selection": {
            "mode": template_selection_mode,
            "template_file": template_workbook.file_name().map(|name| name.to_string_lossy())
        },
        "period": period,
        "batch_id": arguments.batch_id,
        "template_summary": template_plan.summary,
        "formula_summary": {
            "reader": "calamine",
            "files": formula_results.len(),
            "rules_enabled_per_file": template_plan.summary.rules_enabled,
            "formula_result_values_nonempty": formula_nonempty_results,
            "issues": formula_result_issue_count,
            "elapsed_ms": formula_started.elapsed().as_millis()
        },
        "validation_formula_scan": validation_formula_scan_summary,
        "conditional_summary": {
            "status": conditional_report.status,
            "files": conditional_report.files,
            "issues": conditional_report.issues.len(),
            "elapsed_ms": conditional_started.elapsed().as_millis()
        },
        "history_workbook_summary": history_read_summary,
        "summary": history_result.summary,
        "issues": history_result.current_issues,
        "resolved_issues": history_result.resolved_issues,
        "merged_history": history_result.merged_history,
        "elapsed_ms": run_started.elapsed().as_millis()
    });
    if let Some(summary) = result_workbook_summary {
        output["result_workbook"] = serde_json::to_value(summary)?;
    }
    if let Some(summary) = history_result_workbook_summary {
        output["history_result_workbook"] = serde_json::to_value(summary)?;
    }
    if let Some(preflight) = structure_preflight {
        output["structure_preflight"] = serde_json::to_value(preflight)?;
    }
    println!("{}", serde_json::to_string_pretty(&output)?);
    Ok(())
}

fn read_s1_audit_copy_manifest(
    manifest_path: &Path,
    input_dir: &Path,
    accepted_inputs: &[S1ConditionalBatchInput],
) -> Result<Vec<S1IssueFilePair>> {
    let manifest_text = fs::read_to_string(manifest_path)
        .with_context(|| format!("读取公式审核副本清单 {}", manifest_path.display()))?;
    let manifest: S1AuditCopyManifest = serde_json::from_str(&manifest_text)
        .with_context(|| format!("解析公式审核副本清单 {}", manifest_path.display()))?;
    anyhow::ensure!(
        !manifest.files.is_empty(),
        "公式审核副本清单中没有工作簿映射"
    );
    let manifest_dir = manifest_path.parent().unwrap_or_else(|| Path::new("."));
    let accepted_by_source = accepted_inputs
        .iter()
        .map(|input| {
            (
                s1_path_identity(&input.source_file),
                input.source_file.clone(),
            )
        })
        .collect::<std::collections::HashMap<_, _>>();
    let mut sources = std::collections::HashSet::new();
    let mut audits = std::collections::HashSet::new();
    let mut pairs = Vec::with_capacity(manifest.files.len());
    for entry in manifest.files {
        let source_path = resolve_manifest_path(input_dir, entry.source_file);
        let source_file = fs::canonicalize(&source_path)
            .with_context(|| format!("公式审核副本清单源文件不存在：{}", source_path.display()))?;
        let audit_path = resolve_manifest_path(manifest_dir, entry.audit_workbook);
        let audit_workbook = fs::canonicalize(&audit_path)
            .with_context(|| format!("公式审核副本不存在：{}", audit_path.display()))?;
        anyhow::ensure!(
            accepted_by_source.contains_key(&s1_path_identity(&source_file)),
            "公式审核副本清单包含未通过结构预检或未纳入本次审核的源文件：{}",
            source_file
                .file_name()
                .unwrap_or_default()
                .to_string_lossy()
        );
        anyhow::ensure!(
            sources.insert(s1_path_identity(&source_file)),
            "公式审核副本清单中存在重复源文件：{}",
            source_file
                .file_name()
                .unwrap_or_default()
                .to_string_lossy()
        );
        anyhow::ensure!(
            audits.insert(s1_path_identity(&audit_workbook)),
            "多个源文件不能共用同一个公式审核副本：{}",
            audit_workbook
                .file_name()
                .unwrap_or_default()
                .to_string_lossy()
        );
        anyhow::ensure!(
            audit_workbook
                .extension()
                .and_then(|extension| extension.to_str())
                .is_some_and(|extension| {
                    extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm")
                }),
            "公式审核副本必须是 .xlsx 或 .xlsm：{}",
            audit_workbook
                .file_name()
                .unwrap_or_default()
                .to_string_lossy()
        );
        pairs.push(S1IssueFilePair {
            source_file,
            audit_workbook,
        });
    }
    anyhow::ensure!(
        pairs.len() == accepted_inputs.len()
            && accepted_by_source
                .keys()
                .all(|source| sources.contains(source)),
        "公式审核副本清单必须且只能覆盖所有通过结构预检的源文件（通过预检 {} 个，清单映射 {} 个）",
        accepted_inputs.len(),
        pairs.len()
    );
    pairs.sort_by_key(|pair| {
        accepted_inputs
            .iter()
            .position(|input| same_file_path(&input.source_file, &pair.source_file))
            .unwrap_or(usize::MAX)
    });
    Ok(pairs)
}

fn s1_path_identity(path: &Path) -> String {
    let canonical = fs::canonicalize(path).unwrap_or_else(|_| path.to_path_buf());
    let value = canonical.to_string_lossy().replace('\\', "/");
    #[cfg(windows)]
    {
        value.to_lowercase()
    }
    #[cfg(not(windows))]
    {
        value
    }
}

fn run_s1_history_classify(arguments: S1HistoryClassifyArgs) -> Result<()> {
    let started = Instant::now();
    let started_at = chrono::Local::now().to_rfc3339();
    if arguments.developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S1-F01",
                "step": "history_issue_classification",
                "status": "started",
                "engine": "rust_s1_history",
                "started_at": started_at
            }))?
        );
    } else {
        eprintln!("正在按历史审核记录归类 S1-F01 审核项…");
    }

    let mut request: S1HistoryRequest = serde_json::from_str(
        &fs::read_to_string(&arguments.input)
            .with_context(|| format!("读取 S1 历史审核数据 {}", arguments.input.display()))?,
    )
    .with_context(|| format!("解析 S1 历史审核数据 {}", arguments.input.display()))?;
    let history_read_summary: Option<S1HistoryWorkbookReadSummary> = if let Some(history_path) =
        arguments.history_workbook.as_deref()
    {
        let history_engine = if history_path.is_file()
            && history_path
                .extension()
                .and_then(|extension| extension.to_str())
                .is_some_and(|extension| {
                    extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm")
                }) {
            "calamine"
        } else {
            "none"
        };
        let history_started = Instant::now();
        let history_started_at = chrono::Local::now().to_rfc3339();
        if arguments.developer_view {
            eprintln!(
                "{}",
                serde_json::to_string(&serde_json::json!({
                    "type": "progress",
                    "product": "open-data-audit",
                    "feature": "S1-F01",
                    "step": "history_workbook_read",
                    "status": "started",
                    "engine": history_engine,
                    "started_at": history_started_at
                }))?
            );
        } else {
            eprintln!("正在读取 S1 历史工作簿…");
        }
        let loaded =
            match read_s1_history_workbook(history_path, arguments.history_sheet.as_deref()) {
                Ok(loaded) => loaded,
                Err(error) => {
                    if arguments.developer_view {
                        eprintln!(
                            "{}",
                            serde_json::to_string(&serde_json::json!({
                                "type": "progress",
                                "product": "open-data-audit",
                                "feature": "S1-F01",
                                "step": "history_workbook_read",
                                "status": "failed",
                                "engine": history_engine,
                                "elapsed_ms": history_started.elapsed().as_millis(),
                                "error": format!("{error:#}")
                            }))?
                        );
                    } else {
                        eprintln!(
                            "S1 历史工作簿读取失败；引擎 {}，耗时 {} ms：{error:#}",
                            history_engine,
                            history_started.elapsed().as_millis()
                        );
                    }
                    return Err(error);
                }
            };
        request.history = loaded.issues;
        let summary = loaded.summary;
        if arguments.developer_view {
            eprintln!(
                "{}",
                serde_json::to_string(&serde_json::json!({
                    "type": "progress",
                    "product": "open-data-audit",
                    "feature": "S1-F01",
                    "step": "history_workbook_read",
                    "status": "completed",
                    "engine": summary.reader,
                    "sheet_name": summary.sheet_name,
                    "rows_total": summary.rows_total,
                    "rows_read": summary.rows_read,
                    "rows_skipped": summary.rows_skipped,
                    "started_at": history_started_at,
                    "finished_at": chrono::Local::now().to_rfc3339(),
                    "elapsed_ms": summary.elapsed_ms
                }))?
            );
        } else {
            match summary.status {
                    "missing_file" => eprintln!(
                        "S1 历史工作簿不存在，按空历史处理；未调用 Excel 读取引擎，耗时 {} ms。",
                        summary.elapsed_ms
                    ),
                    "sheet_missing" => eprintln!(
                        "S1 历史工作簿未找到支持的历史工作表，按空历史处理；引擎 {}，耗时 {} ms。",
                        summary.reader, summary.elapsed_ms
                    ),
                    _ => eprintln!(
                        "S1 历史工作簿读取完成：引擎 {}，工作表 {}，读取 {} 条、跳过 {} 条；耗时 {} ms。",
                        summary.reader,
                        summary.sheet_name.as_deref().unwrap_or("未知"),
                        summary.rows_read,
                        summary.rows_skipped,
                        summary.elapsed_ms
                    ),
                }
        }
        Some(summary)
    } else {
        None
    };
    let history_rows = request.history.len();
    let batch_id = request.batch_id.clone();
    let mut result = process_s1_history(request);
    result.summary.elapsed_ms = started.elapsed().as_millis();

    if arguments.developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S1-F01",
                "step": "history_issue_classification",
                "status": "completed",
                "engine": "rust_s1_history",
                "batch_id": batch_id,
                "started_at": started_at,
                "finished_at": chrono::Local::now().to_rfc3339(),
                "elapsed_ms": result.summary.elapsed_ms,
                "history_rows": history_rows,
                "current_issues": result.summary.current_issues,
                "new_issues": result.summary.new_issues,
                "continuous_issues": result.summary.continuous_issues,
                "recurring_issues": result.summary.recurring_issues,
                "resolved_issues": result.summary.resolved_issues,
                "merged_history_rows": result.summary.merged_history_rows,
                "previous_period": result.summary.previous_period
            }))?
        );
    } else {
        eprintln!(
            "S1-F01 历史归类完成：引擎 Rust，历史 {} 条，本期 {} 条（新增 {}、连续 {}、再次出现 {}），已整改 {} 条；耗时 {} ms。",
            result.summary.history_rows,
            result.summary.current_issues,
            result.summary.new_issues,
            result.summary.continuous_issues,
            result.summary.recurring_issues,
            result.summary.resolved_issues,
            result.summary.elapsed_ms
        );
    }

    let mut output = serde_json::to_value(&result)?;
    if let Some(summary) = history_read_summary {
        output["history_workbook_read"] = serde_json::to_value(summary)?;
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
            "表结构预检完成：通过 {} 个，跳过 {} 个；模板读取 {} ms，预检 {} ms。",
            summary.files_accepted,
            summary.files_rejected,
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
                    "模板命名区域映射完成：条件范围 {} 个，结构范围 {} 个，匹配但未解析的名称 {} 个，耗时 {} ms。",
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

fn prepare_s1_result_workbook_path(output_path: &Path) -> Result<PathBuf> {
    anyhow::ensure!(
        output_path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx")),
        "S1 结果工作簿输出路径必须使用 .xlsx"
    );
    let absolute = if output_path.is_absolute() {
        output_path.to_path_buf()
    } else {
        std::env::current_dir()?.join(output_path)
    };
    let parent = absolute
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or(Path::new("."));
    let file_name = absolute
        .file_name()
        .context("S1 结果工作簿输出路径缺少文件名")?;
    let resolved = if parent.is_dir() {
        fs::canonicalize(parent)?.join(file_name)
    } else {
        absolute
    };
    anyhow::ensure!(
        !resolved.exists(),
        "拒绝覆盖已有 S1 结果工作簿：{}",
        resolved.display()
    );
    Ok(resolved)
}

fn prepare_s1_history_result_workbook_path(output_path: &Path) -> Result<PathBuf> {
    anyhow::ensure!(
        output_path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx")),
        "S1 历史结果工作簿输出路径必须使用 .xlsx"
    );
    let absolute = if output_path.is_absolute() {
        output_path.to_path_buf()
    } else {
        std::env::current_dir()?.join(output_path)
    };
    let parent = absolute
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or(Path::new("."));
    let file_name = absolute
        .file_name()
        .context("S1 历史结果工作簿输出路径缺少文件名")?;
    let resolved = if parent.is_dir() {
        fs::canonicalize(parent)?.join(file_name)
    } else {
        absolute
    };
    anyhow::ensure!(
        !resolved.exists(),
        "拒绝覆盖已有 S1 历史结果工作簿：{}",
        resolved.display()
    );
    Ok(resolved)
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

fn run_s5_config_check(arguments: S5ConfigCheckArgs) -> Result<()> {
    let started = Instant::now();
    let template_path = arguments
        .template_path
        .as_deref()
        .map(|path| path.to_string_lossy().into_owned())
        .unwrap_or_default();
    let template_dir = arguments
        .template_dir
        .as_deref()
        .map(|path| path.to_string_lossy().into_owned())
        .unwrap_or_default();
    let external_path = arguments
        .external_path
        .as_deref()
        .map(|path| path.to_string_lossy().into_owned())
        .unwrap_or_default();
    let template_catalog_engine = if matches!(arguments.scope.as_str(), "all" | "all_enabled")
        && arguments.template_dir.is_some()
    {
        "OOXML workbook.xml via excel-reader"
    } else {
        "not_used"
    };
    if arguments.developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&json!({
                "type": "progress",
                "step": "s5_config_check_started",
                "engine": "Calamine via excel-reader",
                "template_catalog_engine": template_catalog_engine,
                "config": arguments.config.display().to_string(),
                "template_path": template_path.as_str(),
                "template_dir": template_dir.as_str(),
                "external_path": external_path.as_str(),
                "scope": arguments.scope.as_str(),
                "elapsed_ms": 0,
            }))?
        );
    }
    let result = validate_s5_config(
        &arguments.config,
        &template_path,
        &template_dir,
        "",
        &external_path,
        &arguments.business,
        &arguments.scope,
    )?;
    if arguments.developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&json!({
                "type": "progress",
                "step": "s5_config_check_completed",
                "engine": "Calamine via excel-reader",
                "template_catalog_engine": template_catalog_engine,
                "valid": result.get("valid").and_then(serde_json::Value::as_bool),
                "complete": result.get("complete").and_then(serde_json::Value::as_bool),
                "elapsed_ms": started.elapsed().as_millis(),
            }))?
        );
    }
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}

fn default_s5_template_dir() -> Option<PathBuf> {
    let base_data_dir = std::env::var_os("BASE_AUDIT_DATA_DIR")
        .filter(|value| !value.is_empty())
        .map(PathBuf::from)
        .or_else(|| {
            std::env::var_os("USERPROFILE")
                .or_else(|| std::env::var_os("HOME"))
                .map(|home| PathBuf::from(home).join(".audit-tools"))
        })?;
    Some(base_data_dir.join("s5").join("templates"))
}

struct S5CurrentSourceArgument {
    source_id: String,
    business: String,
    path: PathBuf,
}

fn parse_s5_current_source_argument(value: &str) -> Result<S5CurrentSourceArgument> {
    let mut parts = value.splitn(3, '=');
    let source_id = parts.next().unwrap_or_default().trim();
    let business = parts.next().unwrap_or_default().trim();
    let path = parts.next().unwrap_or_default();
    anyhow::ensure!(
        !source_id.is_empty() && !path.is_empty(),
        "--current-source must use SOURCE_ID=BUSINESS=FILE form"
    );
    Ok(S5CurrentSourceArgument {
        source_id: source_id.to_owned(),
        business: business.to_owned(),
        path: PathBuf::from(path),
    })
}

fn run_s5_audit_run(arguments: S5AuditRunArgs) -> Result<()> {
    anyhow::ensure!(
        arguments.config.is_file(),
        "S5 审核配置工作簿不存在：{}",
        arguments.config.display()
    );
    anyhow::ensure!(
        arguments.source.is_file(),
        "S5 本期数据源工作簿不存在：{}",
        arguments.source.display()
    );
    let additional_current_sources: Vec<S5CurrentSourceArgument> = arguments
        .current_sources
        .iter()
        .map(|value| parse_s5_current_source_argument(value))
        .collect::<Result<_>>()?;
    let mut current_source_ids = vec![arguments.source_id.trim().to_owned()];
    for source in &additional_current_sources {
        anyhow::ensure!(
            source.path.is_file(),
            "S5 附加本期数据源工作簿不存在：{}",
            source.path.display()
        );
        anyhow::ensure!(
            !current_source_ids.iter().any(|id| id == &source.source_id),
            "本期逻辑数据源 ID 重复：{}",
            source.source_id
        );
        current_source_ids.push(source.source_id.clone());
    }
    if let Some(template_path) = &arguments.template_path {
        anyhow::ensure!(
            template_path.is_file(),
            "S5 数字区域布局模板工作簿不存在：{}",
            template_path.display()
        );
    }
    if let Some(template_dir) = &arguments.template_dir {
        anyhow::ensure!(
            template_dir.is_dir(),
            "S5 数字区域布局模板目录不存在：{}",
            template_dir.display()
        );
    }
    let default_template_dir = if arguments.template_dir.is_none() {
        default_s5_template_dir()
    } else {
        None
    };
    let template_dir = arguments
        .template_dir
        .as_deref()
        .or(default_template_dir.as_deref());
    if let Some(previous_source) = &arguments.previous_source {
        anyhow::ensure!(
            previous_source.is_file(),
            "S5 上期数据源工作簿不存在：{}",
            previous_source.display()
        );
        anyhow::ensure!(
            arguments
                .previous_period
                .as_deref()
                .is_some_and(|period| !period.trim().is_empty()),
            "--previous-period must be supplied with --previous-source"
        );
    } else {
        anyhow::ensure!(
            arguments.previous_source_id.is_none(),
            "--previous-source-id requires --previous-source"
        );
    }
    if let Some(external_source) = &arguments.external_source {
        anyhow::ensure!(
            external_source.is_file(),
            "S5 外部数据源工作簿不存在：{}",
            external_source.display()
        );
        anyhow::ensure!(
            arguments
                .external_source_id
                .as_deref()
                .is_some_and(|source_id| !source_id.trim().is_empty()),
            "--external-source-id must be supplied with --external-source"
        );
    } else {
        anyhow::ensure!(
            arguments.external_source_id.is_none(),
            "--external-source-id requires --external-source"
        );
    }
    anyhow::ensure!(
        !arguments.period.trim().is_empty(),
        "运行上下文必须提供报告期"
    );

    let run_id = arguments
        .run_id
        .clone()
        .unwrap_or_else(|| format!("s5-{}", chrono::Local::now().format("%Y%m%d_%H%M%S")));
    let started = Instant::now();
    let mut source_roles = vec!["current"];
    if !additional_current_sources.is_empty() {
        source_roles.push("additional current");
    }
    if arguments.previous_source.is_some() {
        source_roles.push("previous");
    }
    if arguments.external_source.is_some() {
        source_roles.push("external");
    }
    if arguments.developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S5-F01",
                "step": "audit_run",
                "status": "started",
                "run_id": run_id,
                "period": arguments.period,
                "businesses": arguments.businesses,
                "source_roles": source_roles,
                "current_source_ids": current_source_ids,
                "template_path": arguments.template_path,
                "template_dir": template_dir,
                "numeric_region_template_engine": "Calamine via excel-reader + namespace-aware defined names",
                "previous_period": arguments.previous_period,
                "external_source_id": arguments.external_source_id,
                "engine": "rust_s5_configured_audit",
                "started_at": chrono::Local::now().to_rfc3339()
            }))?
        );
    } else {
        eprintln!(
            "正在执行 S5 配置化审核（Rust 配置加载、{}来源读取与规则求值{}）；报告期 {}",
            source_roles.join("、"),
            if arguments.template_path.is_some() || template_dir.is_some() {
                "、数字区域模板投影"
            } else {
                ""
            },
            arguments.period
        );
    }

    let configuration = load_configuration(&arguments.config, false, "")
        .map_err(|error| anyhow::anyhow!("{error}"))?;
    let previous_source_id = arguments
        .previous_source_id
        .clone()
        .unwrap_or_else(|| arguments.source_id.clone());
    let external_source_id = arguments.external_source_id.clone().unwrap_or_default();
    let previous_descriptor = arguments
        .previous_source
        .as_deref()
        .zip(arguments.previous_period.as_deref())
        .map(|(path, period)| S5SourceDescriptor {
            path,
            source_id: &previous_source_id,
            period: period.trim(),
            institution_code: &arguments.org_code,
            institution_name: &arguments.org_name,
        });
    let external_descriptor = arguments
        .external_source
        .as_deref()
        .map(|path| S5SourceDescriptor {
            path,
            source_id: &external_source_id,
            period: arguments.period.trim(),
            institution_code: &arguments.org_code,
            institution_name: &arguments.org_name,
        });
    let mut current_bindings = vec![S5CurrentSourceBinding {
        source: S5SourceDescriptor {
            path: &arguments.source,
            source_id: &arguments.source_id,
            period: arguments.period.trim(),
            institution_code: &arguments.org_code,
            institution_name: &arguments.org_name,
        },
        business: &arguments.source_business,
    }];
    current_bindings.extend(additional_current_sources.iter().map(|binding| {
        S5CurrentSourceBinding {
            source: S5SourceDescriptor {
                path: &binding.path,
                source_id: &binding.source_id,
                period: arguments.period.trim(),
                institution_code: &arguments.org_code,
                institution_name: &arguments.org_name,
            },
            business: &binding.business,
        }
    }));
    let mut retrieval = S5Retrieval::bind_with_current_sources_and_templates(
        configuration.clone(),
        &current_bindings,
        previous_descriptor,
        external_descriptor,
        arguments.period.trim(),
        arguments.template_path.as_deref(),
        template_dir,
    )
    .map_err(|failure| anyhow::anyhow!("{}：{}", failure.code, failure.message))?;
    let source_sha256_before = retrieval.source_sha256().to_owned();

    let (outcomes, status_counts) = run_rules(
        &configuration,
        &mut retrieval,
        arguments.period.trim(),
        &arguments.businesses,
        &arguments.org_code,
        &arguments.org_name,
    )
    .map_err(|error| anyhow::anyhow!("{error}"))?;

    let current_source_snapshots = retrieval.additional_current_source_snapshots();
    let mut additional_output_sources: Vec<S5OutputSource<'_>> = current_source_snapshots
        .iter()
        .map(|source| S5OutputSource {
            role: "current",
            path: &source.path,
            source_id: &source.source_id,
            sha256_before: &source.sha256,
        })
        .collect();
    if let (Some(path), Some(sha256_before)) = (
        arguments.previous_source.as_deref(),
        retrieval.previous_source_sha256(),
    ) {
        additional_output_sources.push(S5OutputSource {
            role: "previous",
            path,
            source_id: &previous_source_id,
            sha256_before,
        });
    }
    if let (Some(path), Some(sha256_before)) = (
        arguments.external_source.as_deref(),
        retrieval.external_source_sha256(),
    ) {
        additional_output_sources.push(S5OutputSource {
            role: "external",
            path,
            source_id: &external_source_id,
            sha256_before,
        });
    }
    let selected_template_bindings = retrieval.numeric_region_template_sources();
    let selected_output_templates: Vec<S5OutputSource<'_>> = selected_template_bindings
        .iter()
        .map(|template| S5OutputSource {
            role: "layout",
            path: &template.path,
            source_id: &template.source_id,
            sha256_before: &template.sha256,
        })
        .collect();
    let output_summary = write_run_outputs(
        &S5OutputRequest {
            run_id: &run_id,
            period: arguments.period.trim(),
            institution_code: &arguments.org_code,
            institution_name: &arguments.org_name,
            outcomes: &outcomes,
            status_counts: &status_counts,
            data_quality_warnings: retrieval.data_quality_warnings(),
            source_path: &arguments.source,
            source_id: &arguments.source_id,
            source_sha256_before: &source_sha256_before,
            additional_sources: &additional_output_sources,
            selected_templates: &selected_output_templates,
            config_path: &arguments.config,
            config_sha256_before: &configuration.sha256,
            config_version: &configuration.version,
        },
        &arguments.output_dir,
    )
    .map_err(|error| anyhow::anyhow!("{error}"))?;

    let elapsed_ms = started.elapsed().as_millis();
    let numeric_region_template_selections = retrieval.numeric_region_template_selections();
    if !arguments.developer_view && !numeric_region_template_selections.is_empty() {
        let selected_templates: Vec<&serde_json::Value> = numeric_region_template_selections
            .iter()
            .filter(|selection| selection["template"] != "兼容回退")
            .collect();
        let fallback_selections: Vec<&serde_json::Value> = numeric_region_template_selections
            .iter()
            .filter(|selection| selection["template"] == "兼容回退")
            .collect();
        let selection_elapsed_ms = |selections: &[&serde_json::Value]| -> u128 {
            selections
                .iter()
                .filter_map(|selection| selection["elapsed_ms"].as_u64())
                .map(u128::from)
                .sum()
        };
        let selected_template_count = selected_templates.len();
        let template_elapsed_ms = selection_elapsed_ms(&selected_templates);
        let fallback_elapsed_ms = selection_elapsed_ms(&fallback_selections);
        if selected_template_count > 0 {
            eprintln!(
                "数字区域模板投影完成：{} 个工作表，Calamine via excel-reader，耗时 {} ms",
                selected_template_count, template_elapsed_ms
            );
        }
        if !fallback_selections.is_empty() {
            eprintln!(
                "未匹配到数字区域布局模板：{} 个工作表按全表兼容范围检索，耗时 {} ms",
                fallback_selections.len(),
                fallback_elapsed_ms
            );
        }
    }
    let output = serde_json::json!({
        "status": "completed",
        "run_id": run_id,
        "period": arguments.period,
        "previous_period": arguments.previous_period,
        "external_source_id": arguments.external_source_id,
        "businesses": arguments.businesses,
        "source_roles": source_roles,
        "rule_results": outcomes.len(),
        "status_counts": status_counts,
        "output_files": {
            "result_workbook": output_summary.output_files.result_workbook,
            "run_ledger": output_summary.output_files.run_ledger,
            "audit_copies": output_summary.output_files.audit_copies,
            "configuration_copy": output_summary.output_files.configuration_copy,
        },
        "triggered_rows": output_summary.triggered_rows,
        "failure_rows": output_summary.failure_rows,
        "data_quality_warning_count": output_summary.data_quality_warning_count,
        "engine": "rust_s5_configured_audit",
        "elapsed_ms": elapsed_ms,
    });
    if arguments.developer_view {
        eprintln!(
            "{}",
            serde_json::to_string(&serde_json::json!({
                "type": "progress",
                "product": "open-data-audit",
                "feature": "S5-F01",
                "step": "audit_run",
                "status": "completed",
                "run_id": run_id,
                "elapsed_ms": elapsed_ms,
                "source_roles": source_roles,
                "numeric_region_templates": numeric_region_template_selections,
            }))?
        );
    }
    println!("{output}");
    Ok(())
}
