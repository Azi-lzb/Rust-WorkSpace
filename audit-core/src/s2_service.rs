//! File-based S2 report service composed from the separately parity-checked
//! configuration, import, analysis, and report-writing components.
//!
//! Synthetic end-to-end parity is covered and the service is exposed through
//! the `open-data-audit s2` CLI candidate. It is not yet a production default;
//! real business templates and target Linux packaging still need acceptance.

use crate::{
    s2_analysis::{build_analysis_result, first_records_by_logical_key, S2AnalysisResult},
    s2_config::{load_config, S2AuditConfig},
    s2_config_check::check_s2_config,
    s2_exporter::export_s2_report,
    s2_external::compare_external,
    s2_external_import::import_external_workbook,
    s2_import::{import_period_directory_with_progress, S2ImportResult},
    s2_period::compare_periods,
    s2_rule_engine::evaluate_rules,
    s2_validation::{
        validate_config_warnings, validate_external_warnings, validate_import, S2AuditFinding,
    },
};
use anyhow::{anyhow, Result};
use serde::{Deserialize, Serialize};
use std::{
    path::{Path, PathBuf},
    time::Instant,
};

#[derive(Clone, Debug, PartialEq)]
pub struct S2RunRequest {
    pub current_directory: PathBuf,
    pub previous_directory: PathBuf,
    pub config_path: PathBuf,
    pub output_directory: PathBuf,
    pub external_workbook: Option<PathBuf>,
}

#[derive(Clone, Debug, PartialEq, Deserialize, Serialize)]
pub struct S2ProgressEvent {
    pub stage: String,
    pub label: String,
    pub percent: f64,
    pub current: usize,
    pub total: usize,
    pub message: String,
}

#[derive(Clone, Debug, PartialEq)]
pub struct S2RunResult {
    pub analysis: S2AnalysisResult,
    pub report_path: PathBuf,
}

/// Return the actual implementation used for an S2 progress stage.
pub fn s2_engine_for_stage(stage: &str) -> Option<&'static str> {
    match stage {
        "config" | "current" | "previous" | "external_read" => Some("Calamine"),
        "validate" | "period" | "rules" | "external" | "complete" => Some("Rust audit-core"),
        "export" => Some("rust_xlsxwriter"),
        "save" => Some("rust_xlsxwriter + Calamine read-back"),
        _ => None,
    }
}

/// Return the source service's read-only configuration issue messages.
pub fn check_config(config_path: impl AsRef<Path>) -> Vec<String> {
    check_s2_config(config_path)
        .issues
        .into_iter()
        .map(|issue| issue.message)
        .collect()
}

/// Execute the S2 file-based flow without changing any input workbook.
///
/// The service emits completion only after the candidate report has been
/// written, reopened, validated, and published by `export_s2_report`.
pub fn run_s2_audit(
    request: &S2RunRequest,
    on_progress: Option<&mut dyn FnMut(S2ProgressEvent)>,
) -> Result<S2RunResult> {
    let mut on_progress = on_progress;
    emit_progress(
        &mut on_progress,
        "config",
        "检查审核配置",
        0.0,
        0,
        0,
        format!("读取配置文件：{}", file_name(&request.config_path)),
    );

    let config_report = check_s2_config(&request.config_path);
    if !config_report.passed {
        let details = config_report
            .issues
            .iter()
            .filter(|issue| issue.level == "blocking")
            .map(|issue| issue.message.as_str())
            .collect::<Vec<_>>()
            .join("；");
        return Err(anyhow!("审核配置存在阻断错误，不能执行报表规则：{details}"));
    }
    let config = load_config(&request.config_path)?;
    let import_config = config.import_config();

    emit_progress(
        &mut on_progress,
        "current",
        "读取本期报表",
        5.0,
        0,
        0,
        format!("本期目录：{}", request.current_directory.display()),
    );
    let current_import = {
        let mut report_progress = |current: usize, total: usize, message: &str| {
            emit_progress(
                &mut on_progress,
                "current",
                "读取本期报表",
                5.0 + 25.0 * current as f64 / total.max(1) as f64,
                current,
                total,
                message.to_owned(),
            );
        };
        import_period_directory_with_progress(
            &request.current_directory,
            &import_config,
            "本期",
            Some(&mut report_progress),
        )?
    };

    emit_progress(
        &mut on_progress,
        "previous",
        "读取上期报表",
        30.0,
        0,
        0,
        format!("上期目录：{}", request.previous_directory.display()),
    );
    let previous_import = {
        let mut report_progress = |current: usize, total: usize, message: &str| {
            emit_progress(
                &mut on_progress,
                "previous",
                "读取上期报表",
                30.0 + 25.0 * current as f64 / total.max(1) as f64,
                current,
                total,
                message.to_owned(),
            );
        };
        import_period_directory_with_progress(
            &request.previous_directory,
            &import_config,
            "上期",
            Some(&mut report_progress),
        )?
    };

    emit_progress(
        &mut on_progress,
        "validate",
        "检查导入数据",
        55.0,
        0,
        0,
        "正在整理本期与上期指标".to_owned(),
    );
    let current_records = first_records_by_logical_key(&current_import);
    let previous_records = first_records_by_logical_key(&previous_import);
    let mut quality_findings = initial_quality_findings(&config, &current_import, &previous_import);
    emit_progress(
        &mut on_progress,
        "validate",
        "检查导入数据",
        59.0,
        0,
        0,
        format!(
            "完成：本期 {} 项指标，上期 {} 项指标，{} 条导入/配置提示",
            current_records.len(),
            previous_records.len(),
            quality_findings.len()
        ),
    );

    emit_progress(
        &mut on_progress,
        "period",
        "执行环比规则",
        60.0,
        0,
        0,
        format!(
            "正在对比本期 {} 项与上期 {} 项指标",
            current_records.len(),
            previous_records.len()
        ),
    );
    let period_findings =
        compare_periods(&current_records, &previous_records, &config.period_config())?;
    emit_progress(
        &mut on_progress,
        "period",
        "执行环比规则",
        66.0,
        0,
        0,
        format!("完成：生成 {} 条环比结果", period_findings.len()),
    );

    emit_progress(
        &mut on_progress,
        "rules",
        "执行校验规则",
        67.0,
        0,
        0,
        format!("正在按配置检查本期 {} 项指标", current_records.len()),
    );
    let (rule_findings, rule_quality_findings) = evaluate_rules(
        &current_records,
        &previous_records,
        &config.rule_engine_config(),
    )?;
    emit_progress(
        &mut on_progress,
        "rules",
        "执行校验规则",
        72.0,
        0,
        0,
        format!(
            "完成：生成 {} 条校验结果，另有 {} 条规则数据提示",
            rule_findings.len(),
            rule_quality_findings.len()
        ),
    );

    let mut external_findings = Vec::new();
    let mut external_warning_count = 0;
    if let Some(external_path) = &request.external_workbook {
        emit_progress(
            &mut on_progress,
            "external_read",
            "读取大集中数据",
            73.0,
            0,
            0,
            format!("正在读取：{}", file_name(external_path)),
        );
        let started = Instant::now();
        let imported = import_external_workbook(external_path)?;
        let external_quality = validate_external_warnings(&imported.warnings);
        external_warning_count = external_quality.len();
        quality_findings.extend(external_quality);
        let reader_elapsed = if imported.reader_elapsed_seconds > 0.0 {
            imported.reader_elapsed_seconds
        } else {
            started.elapsed().as_secs_f64()
        };
        let mut reader_label = imported.reader_engine.clone();
        if !imported.reader_version.is_empty() {
            reader_label.push(' ');
            reader_label.push_str(&imported.reader_version);
        }
        emit_progress(
            &mut on_progress,
            "external_read",
            "读取大集中数据",
            76.0,
            0,
            0,
            format!(
                "完成：{reader_label}；读取 {} 条指标数据，机构参照 {} 条，耗时 {:.2} 秒；外部数据提示 {} 条",
                imported.data.values.len(),
                imported.external_org_map.len(),
                reader_elapsed,
                external_warning_count
            ),
        );
        emit_progress(
            &mut on_progress,
            "external",
            "执行外部核对",
            77.0,
            0,
            0,
            format!(
                "正在用 {} 条大集中指标数据核对报表",
                imported.data.values.len()
            ),
        );
        external_findings = compare_external(
            &current_records,
            &imported.data,
            &imported.external_org_map,
            &config.external_config(),
        )?;
        emit_progress(
            &mut on_progress,
            "external",
            "执行外部核对",
            79.0,
            0,
            0,
            format!("完成：生成 {} 条外部核对结果", external_findings.len()),
        );
    }

    let analysis = build_analysis_result(
        current_records,
        previous_records,
        period_findings,
        external_findings,
        rule_findings,
        quality_findings,
        rule_quality_findings,
    );
    emit_progress(
        &mut on_progress,
        "export",
        "生成审核结果",
        80.0,
        0,
        0,
        format!(
            "共 {} 条审核结果，按环比、外部核对、校验规则分类写入",
            analysis.findings.len()
        ),
    );

    let report_path = {
        let mut report_progress = |stage: &str, current: usize, total: usize| {
            let is_save = stage == "save";
            let percent = if is_save {
                95.0
            } else if total == 0 {
                80.0
            } else {
                80.0 + 14.0 * current as f64 / total as f64
            };
            let label = if is_save {
                "保存审核工作簿".to_owned()
            } else {
                format!("写入{stage}")
            };
            emit_progress(
                &mut on_progress,
                if is_save { "save" } else { "export" },
                &label,
                percent,
                current,
                total,
                String::new(),
            );
        };
        export_s2_report(
            &request.output_directory,
            &analysis.findings,
            config.hide_unchanged_period_rows,
            config.hide_matching_external_rows,
            Some(&mut report_progress),
        )?
    };

    let mut complete_message = "已生成三张分类审核表".to_owned();
    if external_warning_count > 0 {
        complete_message.push_str(&format!(
            "；外部数据提示 {external_warning_count} 条，详情见“校验规则”工作表"
        ));
    }
    emit_progress(
        &mut on_progress,
        "complete",
        "审核完成",
        100.0,
        0,
        0,
        complete_message,
    );

    Ok(S2RunResult {
        analysis,
        report_path,
    })
}

fn initial_quality_findings(
    config: &S2AuditConfig,
    current_import: &S2ImportResult,
    previous_import: &S2ImportResult,
) -> Vec<S2AuditFinding> {
    let mut findings = validate_config_warnings(&config.warnings);
    findings.extend(validate_import(current_import));
    findings.extend(validate_import(previous_import));
    findings
}

fn emit_progress(
    callback: &mut Option<&mut dyn FnMut(S2ProgressEvent)>,
    stage: &str,
    label: &str,
    percent: f64,
    current: usize,
    total: usize,
    message: String,
) {
    if let Some(callback) = callback.as_mut() {
        (**callback)(S2ProgressEvent {
            stage: stage.to_owned(),
            label: label.to_owned(),
            percent,
            current,
            total,
            message,
        });
    }
}

fn file_name(path: &Path) -> &str {
    path.file_name()
        .and_then(|name| name.to_str())
        .unwrap_or_default()
}

#[cfg(test)]
mod tests {
    use super::s2_engine_for_stage;

    #[test]
    fn progress_stages_report_the_engine_used_by_each_stage() {
        assert_eq!(s2_engine_for_stage("config"), Some("Calamine"));
        assert_eq!(s2_engine_for_stage("current"), Some("Calamine"));
        assert_eq!(s2_engine_for_stage("period"), Some("Rust audit-core"));
        assert_eq!(s2_engine_for_stage("external_read"), Some("Calamine"));
        assert_eq!(s2_engine_for_stage("export"), Some("rust_xlsxwriter"));
        assert_eq!(
            s2_engine_for_stage("save"),
            Some("rust_xlsxwriter + Calamine read-back")
        );
        assert_eq!(s2_engine_for_stage("unknown"), None);
    }
}
