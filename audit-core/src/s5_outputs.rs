//! S5 run outputs (`outputs.write_run_outputs` port, single-business shape).
//!
//! Produces the run directory: audit copies of the source workbook and the
//! configuration with a sha256 manifest, the 审核结果.xlsx result workbook
//! (triggered + failure rows, severity-classified error types, evidence-based
//! locations and hyperlinks to the copies, per-column precision formats, the
//! Python `_style_table` look), and the 运行总账.json ledger. Conditional
//! format results remain a separate future slice; dynamic external-source data
//! quality warnings are written to the baseline-compatible table and ledger.

use crate::s5_service::S5RuleOutcome;
use excel_writer::{Color, Format, FormatAlign, FormatUnderline, Workbook};
use serde_json::{json, Value as JsonValue};
use sha2::{Digest, Sha256};
use std::{
    collections::{BTreeMap, BTreeSet},
    fs,
    path::{Path, PathBuf},
};

#[derive(Clone, Copy)]
pub struct S5OutputSource<'a> {
    pub role: &'a str,
    pub path: &'a Path,
    pub source_id: &'a str,
    pub sha256_before: &'a str,
}

pub struct S5OutputRequest<'a> {
    pub run_id: &'a str,
    pub period: &'a str,
    pub institution_code: &'a str,
    pub institution_name: &'a str,
    pub outcomes: &'a [S5RuleOutcome],
    pub status_counts: &'a BTreeMap<String, usize>,
    pub data_quality_warnings: &'a [JsonValue],
    pub source_path: &'a Path,
    pub source_id: &'a str,
    pub source_sha256_before: &'a str,
    pub additional_sources: &'a [S5OutputSource<'a>],
    pub selected_templates: &'a [S5OutputSource<'a>],
    pub config_path: &'a Path,
    pub config_sha256_before: &'a str,
    pub config_version: &'a str,
}

#[derive(Clone, Debug, PartialEq, serde::Serialize)]
pub struct S5OutputManifestEntry {
    pub role: String,
    pub source_id: String,
    pub source_path: String,
    pub copy_path: String,
    pub sha256_before: String,
    pub sha256_after: String,
    pub sha256_copy: String,
    pub matches_snapshot: bool,
}

#[derive(Clone, Debug, PartialEq, serde::Serialize)]
pub struct S5OutputFiles {
    pub result_workbook: PathBuf,
    pub run_ledger: PathBuf,
    pub audit_copies: PathBuf,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub configuration_copy: Option<PathBuf>,
}

#[derive(Clone, Debug, PartialEq, serde::Serialize)]
pub struct S5WriteOutputsSummary {
    pub run_dir: PathBuf,
    pub manifest: Vec<S5OutputManifestEntry>,
    pub output_files: S5OutputFiles,
    pub run_status: String,
    pub input_integrity: String,
    pub configuration_integrity: String,
    pub status_counts: BTreeMap<String, usize>,
    pub triggered_rows: usize,
    pub failure_rows: usize,
    pub data_quality_warning_count: usize,
}

const RESULT_HEADERS: [&str; 12] = [
    "机构名称",
    "报表项目",
    "统一社会信用代码",
    "规则名称",
    "工作表名",
    "定位单元格",
    "错误类型",
    "校验指标",
    "描述",
    "本期值",
    "对比值",
    "差值",
];
const COLUMN_WIDTHS: [f64; 12] = [
    34.0, 34.0, 26.0, 28.0, 30.0, 20.0, 20.0, 30.0, 70.0, 18.0, 18.0, 18.0,
];

fn path_stem(path: &Path) -> String {
    path.file_stem()
        .map(|stem| stem.to_string_lossy().into_owned())
        .unwrap_or_default()
}
fn safe_name(value: &str) -> String {
    let cleaned: String = value
        .chars()
        .map(|character| {
            if character.is_alphanumeric() || "._-".contains(character) {
                character
            } else {
                '_'
            }
        })
        .collect();
    let trimmed = cleaned.trim_matches('_');
    if trimmed.is_empty() {
        "_".to_owned()
    } else {
        trimmed.to_owned()
    }
}

fn file_sha256(path: &Path) -> Result<String, String> {
    let bytes = fs::read(path)
        .map_err(|error| format!("读取文件计算 SHA-256 失败 {}：{error}", path.display()))?;
    let mut hasher = Sha256::new();
    hasher.update(&bytes);
    Ok(format!("{:x}", hasher.finalize()))
}

fn copy_identity(path: &Path) -> PathBuf {
    fs::canonicalize(path).unwrap_or_else(|_| path.to_path_buf())
}

fn capture_copied_snapshot(
    path: &Path,
    expected: &str,
    copy_path: &Path,
) -> Result<(String, String, bool), String> {
    let actual = file_sha256(path).unwrap_or_else(|_| "MISSING".to_owned());
    let copy = file_sha256(copy_path)?;
    let matches_snapshot = actual == expected && copy == actual;
    Ok((actual, copy, matches_snapshot))
}

fn record_input_hash(
    role: &str,
    source_id: &str,
    path: &Path,
    expected: &str,
    hashes: &mut BTreeMap<String, JsonValue>,
) -> bool {
    let actual = file_sha256(path).unwrap_or_else(|_| "MISSING".to_owned());
    hashes.insert(
        format!("{role}:{source_id}"),
        json!({"sha256_before": expected, "sha256_after": actual}),
    );
    actual == expected
}

fn count_outcome_statuses(outcomes: &[S5RuleOutcome]) -> BTreeMap<String, usize> {
    let mut counts = BTreeMap::new();
    for outcome in outcomes {
        *counts.entry(outcome.status.clone()).or_insert(0) += 1;
    }
    counts
}

fn downgrade_successful_outcomes(outcomes: &mut [S5RuleOutcome]) {
    for outcome in outcomes {
        if matches!(outcome.status.as_str(), "PASSED" | "TRIGGERED") {
            outcome.status = "CALCULATION_FAILED".to_owned();
            outcome.triggered = None;
            outcome.error_code = Some("INPUT_CHANGED_DURING_RUN".to_owned());
            outcome.error_message =
                "审核结果文件生成期间输入或配置哈希发生变化，结论不可用".to_owned();
        }
    }
}

fn failure_type(status: &str) -> Option<&'static str> {
    Some(match status {
        "RETRIEVAL_FAILED" => "取数失败",
        "STRUCTURE_ERROR" => "表结构错误",
        "CALCULATION_FAILED" => "计算失败",
        "CONFIG_ERROR" => "配置错误",
        _ => return None,
    })
}

fn error_type(item: &S5RuleOutcome) -> String {
    if let Some(failure) = failure_type(&item.status) {
        return failure.to_owned();
    }
    let level = item.severity.trim().to_lowercase();
    match level.as_str() {
        "条件格式" | "条件格式触发" | "条件格式错误" => "条件格式触发".to_owned(),
        "错误" | "硬性" | "硬性错误" | "校验硬性错误" | "error" | "hard" => {
            "校验硬性错误".to_owned()
        }
        "核实" | "警告" | "软性" | "软性错误" | "校验软性错误" | "warn" | "warning" | "soft" => {
            "校验软性错误".to_owned()
        }
        _ => item.severity.clone(),
    }
}

struct LocationRender {
    sheet: String,
    cells: String,
    link_evidence: Option<JsonValue>,
}

fn location(item: &S5RuleOutcome, copy_by_source: &BTreeMap<String, PathBuf>) -> LocationRender {
    if failure_type(&item.status).is_some() && !item.error_evidence.is_null() {
        let sheet = item
            .error_evidence
            .get("sheet")
            .and_then(JsonValue::as_str)
            .filter(|text| !text.is_empty())
            .unwrap_or(&item.sheet_name)
            .to_owned();
        let cell = item
            .error_evidence
            .get("cell")
            .and_then(JsonValue::as_str)
            .unwrap_or_default()
            .to_owned();
        return LocationRender {
            sheet,
            cells: cell,
            link_evidence: None,
        };
    }
    if !item.formula_cell.is_empty() {
        return LocationRender {
            sheet: item.sheet_name.clone(),
            cells: item.formula_cell.clone(),
            link_evidence: None,
        };
    }
    let evidence: Vec<&JsonValue> = item
        .evidence
        .iter()
        .filter(|entry| {
            copy_by_source.is_empty()
                || entry
                    .get("source_path")
                    .and_then(JsonValue::as_str)
                    .map(|path| copy_by_source.contains_key(path))
                    .unwrap_or(false)
        })
        .collect();
    let preferred: Vec<&JsonValue> = evidence
        .iter()
        .copied()
        .filter(|entry| {
            entry
                .get("sheet")
                .and_then(JsonValue::as_str)
                .map(|sheet| sheet == item.sheet_name)
                .unwrap_or(false)
        })
        .collect();
    let evidence: Vec<&JsonValue> = if preferred.is_empty() {
        evidence
    } else {
        preferred
    };
    if !evidence.is_empty() {
        let mut sheets: Vec<String> = Vec::new();
        for entry in &evidence {
            let sheet = entry
                .get("sheet")
                .and_then(JsonValue::as_str)
                .unwrap_or_default();
            if !sheets.iter().any(|existing| existing == sheet) {
                sheets.push(sheet.to_owned());
            }
        }
        let single_sheet = sheets.len() == 1;
        let mut cells: Vec<String> = Vec::new();
        for entry in &evidence {
            let cell = entry
                .get("cell")
                .and_then(JsonValue::as_str)
                .unwrap_or_default();
            let rendered = if single_sheet {
                cell.to_owned()
            } else {
                format!(
                    "{}!{cell}",
                    entry
                        .get("sheet")
                        .and_then(JsonValue::as_str)
                        .unwrap_or_default()
                )
            };
            if !cells.iter().any(|existing| *existing == rendered) {
                cells.push(rendered);
            }
        }
        return LocationRender {
            sheet: sheets.join("；"),
            cells: cells.join("；"),
            link_evidence: Some(evidence[0].clone()),
        };
    }
    let sheet = item
        .error_evidence
        .get("sheet")
        .and_then(JsonValue::as_str)
        .filter(|text| !text.is_empty())
        .unwrap_or(&item.sheet_name)
        .to_owned();
    let cell = item
        .error_evidence
        .get("cell")
        .and_then(JsonValue::as_str)
        .unwrap_or_default()
        .to_owned();
    LocationRender {
        sheet,
        cells: cell,
        link_evidence: None,
    }
}

fn description(item: &S5RuleOutcome) -> String {
    if item.status == "TRIGGERED" {
        return item.description.clone();
    }
    let base = failure_type(&item.status).unwrap_or(&item.status);
    let mut message = if item.error_message.is_empty() {
        base.to_owned()
    } else {
        item.error_message.clone()
    };
    if let Some(candidates) = item
        .error_evidence
        .get("cells")
        .and_then(JsonValue::as_array)
    {
        let cells: Vec<String> = candidates
            .iter()
            .map(|cell| {
                cell.as_str()
                    .map(str::to_owned)
                    .unwrap_or_else(|| cell.to_string())
            })
            .collect();
        if !cells.is_empty() {
            message.push_str("；候选单元格：");
            message.push_str(&cells.join("、"));
        }
    }
    message
}

fn relative_link(source: &Path, result: &Path, sheet: &str, cell: &str) -> String {
    let target = source
        .strip_prefix(result.parent().unwrap_or_else(|| Path::new("")))
        .map(|relative| relative.to_string_lossy().replace('\\', "/"))
        .unwrap_or_else(|_| source.to_string_lossy().replace('\\', "/"));
    format!("file:///{target}#'{sheet}'!{cell}")
}

fn numeric_text(value: &JsonValue) -> Option<f64> {
    value.as_f64()
}

/// Write the run outputs; mirrors `write_run_outputs` for the single-binding
/// current-source shape. Conditional-format results remain a separate slice.
pub fn write_run_outputs(
    request: &S5OutputRequest<'_>,
    output_root: &Path,
) -> Result<S5WriteOutputsSummary, String> {
    for source in request.additional_sources {
        if !matches!(source.role, "current" | "previous" | "external") {
            return Err(format!("不支持的附加来源角色：{}", source.role));
        }
    }
    for template in request.selected_templates {
        if template.role != "layout" {
            return Err(format!("不支持的布局模板角色：{}", template.role));
        }
    }
    let output_root = output_root.to_path_buf();
    fs::create_dir_all(&output_root).map_err(|error| format!("创建输出根目录失败：{error}"))?;
    let run_dir = output_root.join(safe_name(request.run_id));
    if run_dir.exists() {
        return Err(format!(
            "运行输出目录已存在，不覆盖已有结果：{}",
            run_dir.display()
        ));
    }
    fs::create_dir(&run_dir).map_err(|error| format!("创建运行目录失败：{error}"))?;
    let copies_dir = run_dir.join("审核副本");
    fs::create_dir(&copies_dir).map_err(|error| format!("创建审核副本目录失败：{error}"))?;

    let mut manifest: Vec<S5OutputManifestEntry> = Vec::new();
    let mut copy_by_source: BTreeMap<String, PathBuf> = BTreeMap::new();
    let mut copied_source_paths = BTreeSet::new();

    if request.source_path.is_file() {
        let source_copy = copies_dir.join(format!(
            "current_{}_{}{}",
            safe_name(request.source_id),
            safe_name(&path_stem(request.source_path)),
            request
                .source_path
                .extension()
                .map(|extension| format!(".{}", extension.to_string_lossy().to_lowercase()))
                .unwrap_or_default()
        ));
        fs::copy(request.source_path, &source_copy)
            .map_err(|error| format!("复制数据源失败：{error}"))?;
        let (sha_after, sha_copy, matches_snapshot) = capture_copied_snapshot(
            request.source_path,
            request.source_sha256_before,
            &source_copy,
        )?;
        manifest.push(S5OutputManifestEntry {
            role: "current".to_owned(),
            source_id: request.source_id.to_owned(),
            source_path: request.source_path.display().to_string(),
            copy_path: source_copy.display().to_string(),
            sha256_before: request.source_sha256_before.to_owned(),
            sha256_after: sha_after,
            sha256_copy: sha_copy,
            matches_snapshot,
        });
        copy_by_source.insert(
            request.source_path.display().to_string(),
            source_copy.clone(),
        );
        copied_source_paths.insert(copy_identity(request.source_path));
    }

    for source in request.additional_sources {
        if !source.path.is_file() {
            continue;
        }
        let extension = source
            .path
            .extension()
            .map(|extension| format!(".{}", extension.to_string_lossy().to_lowercase()))
            .unwrap_or_default();
        let source_copy = copies_dir.join(format!(
            "{}_{}_{}{}",
            safe_name(source.role),
            safe_name(source.source_id),
            safe_name(&path_stem(source.path)),
            extension
        ));
        fs::copy(source.path, &source_copy)
            .map_err(|error| format!("复制附加数据源失败：{error}"))?;
        let (sha_after, sha_copy, matches_snapshot) =
            capture_copied_snapshot(source.path, source.sha256_before, &source_copy)?;
        manifest.push(S5OutputManifestEntry {
            role: source.role.to_owned(),
            source_id: source.source_id.to_owned(),
            source_path: source.path.display().to_string(),
            copy_path: source_copy.display().to_string(),
            sha256_before: source.sha256_before.to_owned(),
            sha256_after: sha_after,
            sha256_copy: sha_copy,
            matches_snapshot,
        });
        copy_by_source.insert(source.path.display().to_string(), source_copy);
        copied_source_paths.insert(copy_identity(source.path));
    }

    for template in request.selected_templates {
        if !template.path.is_file() {
            continue;
        }
        if !copied_source_paths.insert(copy_identity(template.path)) {
            continue;
        }
        let extension = template
            .path
            .extension()
            .map(|extension| format!(".{}", extension.to_string_lossy().to_lowercase()))
            .unwrap_or_default();
        let template_copy = copies_dir.join(format!(
            "layout_{}_{}{}",
            safe_name(template.source_id),
            safe_name(&path_stem(template.path)),
            extension
        ));
        fs::copy(template.path, &template_copy)
            .map_err(|error| format!("复制数字区域布局模板失败：{error}"))?;
        let (sha_after, sha_copy, matches_snapshot) =
            capture_copied_snapshot(template.path, template.sha256_before, &template_copy)?;
        manifest.push(S5OutputManifestEntry {
            role: template.role.to_owned(),
            source_id: template.source_id.to_owned(),
            source_path: template.path.display().to_string(),
            copy_path: template_copy.display().to_string(),
            sha256_before: template.sha256_before.to_owned(),
            sha256_after: sha_after,
            sha256_copy: sha_copy,
            matches_snapshot,
        });
    }

    let config_copy = if request.config_path.is_file() {
        let path = copies_dir.join(format!(
            "配置_{}{}",
            safe_name(&path_stem(request.config_path)),
            request
                .config_path
                .extension()
                .map(|extension| format!(".{}", extension.to_string_lossy().to_lowercase()))
                .unwrap_or_default()
        ));
        fs::copy(request.config_path, &path).map_err(|error| format!("复制配置失败：{error}"))?;
        let (sha_after, sha_copy, matches_snapshot) =
            capture_copied_snapshot(request.config_path, request.config_sha256_before, &path)?;
        manifest.push(S5OutputManifestEntry {
            role: "configuration".to_owned(),
            source_id: "configuration".to_owned(),
            source_path: request.config_path.display().to_string(),
            copy_path: path.display().to_string(),
            sha256_before: request.config_sha256_before.to_owned(),
            sha256_after: sha_after,
            sha256_copy: sha_copy,
            matches_snapshot,
        });
        Some(path)
    } else {
        None
    };

    let mut input_hashes = BTreeMap::new();
    let mut inputs_unchanged = record_input_hash(
        "current",
        request.source_id,
        request.source_path,
        request.source_sha256_before,
        &mut input_hashes,
    );
    for source in request.additional_sources {
        inputs_unchanged &= record_input_hash(
            source.role,
            source.source_id,
            source.path,
            source.sha256_before,
            &mut input_hashes,
        );
    }
    for template in request.selected_templates {
        inputs_unchanged &= record_input_hash(
            template.role,
            template.source_id,
            template.path,
            template.sha256_before,
            &mut input_hashes,
        );
    }
    let configuration_after =
        file_sha256(request.config_path).unwrap_or_else(|_| "MISSING".to_owned());
    let configuration_unchanged = configuration_after == request.config_sha256_before;
    input_hashes.insert(
        "configuration".to_owned(),
        json!({
            "sha256_before": request.config_sha256_before,
            "sha256_after": configuration_after,
        }),
    );
    let integrity_changed = !inputs_unchanged || !configuration_unchanged;
    let input_integrity = if inputs_unchanged {
        "UNCHANGED"
    } else {
        "CHANGED_DURING_RUN"
    };
    let configuration_integrity = if configuration_unchanged {
        "UNCHANGED"
    } else {
        "CHANGED_DURING_RUN"
    };
    let run_status = if integrity_changed {
        "INPUT_CHANGED_DURING_RUN"
    } else {
        "COMPLETED"
    };
    let mut outcomes = request.outcomes.to_vec();
    let mut status_counts = request.status_counts.clone();
    if integrity_changed {
        downgrade_successful_outcomes(&mut outcomes);
        status_counts = count_outcome_statuses(&outcomes);
    }

    let result_path = run_dir.join("审核结果.xlsx");
    let mut triggered_rows = 0usize;
    let mut failure_rows = 0usize;
    {
        let mut workbook = Workbook::new();
        let header_format = Format::new()
            .set_font_name("微软雅黑")
            .set_font_size(10)
            .set_bold()
            .set_font_color(Color::RGB(0xFF_FF_FF))
            .set_background_color(Color::RGB(0x1F_4E_78))
            .set_align(FormatAlign::Center)
            .set_align(FormatAlign::VerticalCenter)
            .set_text_wrap();
        let body_format = Format::new()
            .set_font_name("微软雅黑")
            .set_font_size(10)
            .set_font_color(Color::RGB(0x22_22_22))
            .set_align(FormatAlign::Top)
            .set_text_wrap();
        let text_format = body_format.clone().set_num_format("@");
        let right_format = Format::new()
            .set_font_name("微软雅黑")
            .set_font_size(10)
            .set_font_color(Color::RGB(0x22_22_22))
            .set_align(FormatAlign::Right)
            .set_align(FormatAlign::Top);
        let hyperlink_format = Format::new()
            .set_hyperlink()
            .set_font_name("微软雅黑")
            .set_font_size(10)
            .set_font_color(Color::RGB(0x22_22_22))
            .set_align(FormatAlign::Top)
            .set_text_wrap()
            .set_underline(FormatUnderline::None);

        let worksheet = workbook
            .add_worksheet()
            .set_name("审核结果")
            .map_err(|error| error.to_string())?;
        worksheet.set_screen_gridlines(false);
        for (column, header) in RESULT_HEADERS.iter().enumerate() {
            worksheet
                .write_string_with_format(0, column as u16, *header, &header_format)
                .map_err(|error| error.to_string())?;
        }
        worksheet
            .set_row_height(0, 30.0)
            .map_err(|error| error.to_string())?;

        let mut row_number: u32 = 1;
        for item in &outcomes {
            let is_failure = failure_type(&item.status).is_some();
            if item.status != "TRIGGERED" && !is_failure {
                continue;
            }
            if is_failure {
                failure_rows += 1;
            } else {
                triggered_rows += 1;
            }
            let location = location(item, &copy_by_source);
            let report_item = item
                .evidence
                .iter()
                .find_map(|entry| {
                    entry
                        .get("report_item")
                        .and_then(JsonValue::as_str)
                        .filter(|text| !text.is_empty())
                        .map(str::to_owned)
                })
                .unwrap_or_default();
            let mut link_target: Option<String> = None;
            if let Some(entry) = &location.link_evidence {
                let source_path = entry
                    .get("source_path")
                    .and_then(JsonValue::as_str)
                    .unwrap_or_default();
                let sheet = entry
                    .get("sheet")
                    .and_then(JsonValue::as_str)
                    .unwrap_or_default();
                let cell = entry
                    .get("cell")
                    .and_then(JsonValue::as_str)
                    .unwrap_or_default();
                if let Some(copy) = copy_by_source.get(source_path) {
                    let safe_sheet = sheet.replace('\'', "''");
                    link_target = Some(relative_link(copy, &result_path, &safe_sheet, cell));
                }
            } else {
                let sheet = item
                    .error_evidence
                    .get("sheet")
                    .and_then(JsonValue::as_str)
                    .unwrap_or("");
                let cell = item
                    .error_evidence
                    .get("cell")
                    .and_then(JsonValue::as_str)
                    .unwrap_or("");
                if !cell.is_empty() {
                    let source_path = item
                        .error_evidence
                        .get("source_path")
                        .and_then(JsonValue::as_str)
                        .unwrap_or_default();
                    if let Some(copy) = copy_by_source.get(source_path) {
                        let safe_sheet = sheet.replace('\'', "''");
                        link_target = Some(relative_link(copy, &result_path, &safe_sheet, cell));
                    }
                }
            }
            let texts = [
                request.institution_name.to_owned(),
                report_item,
                request.institution_code.to_owned(),
                item.rule_name.clone(),
                location.sheet.clone(),
                location.cells.clone(),
                error_type(item),
                item.metric.clone(),
                description(item),
            ];
            for (column, value) in texts.iter().enumerate() {
                if column == 5 {
                    if let Some(target) = &link_target {
                        worksheet
                            .write_url_with_options(
                                row_number,
                                column as u16,
                                target.as_str(),
                                value.as_str(),
                                "",
                                Some(&hyperlink_format),
                            )
                            .map_err(|error| error.to_string())?;
                        continue;
                    }
                }
                // Only the credit-code column keeps the literal-text format in
                // the baseline; other text columns render with General.
                let format = if column == 2 {
                    &text_format
                } else {
                    &body_format
                };
                worksheet
                    .write_string_with_format(row_number, column as u16, value.as_str(), format)
                    .map_err(|error| error.to_string())?;
            }
            let value_fields = [
                (item.current_value.as_ref(), item.current_precision),
                (item.comparison_value.as_ref(), item.comparison_precision),
                (item.difference_value.as_ref(), item.difference_precision),
            ];
            for (offset, (value_option, precision)) in value_fields.iter().enumerate() {
                let column = 9 + offset as u16;
                let owned = value_option.cloned().unwrap_or(JsonValue::Null);
                let value = &owned;
                // The baseline styles every J/K/L cell right/top without wrap,
                // including empty ones.
                let format = precision
                    .map(|digits| {
                        right_format.clone().set_num_format(if digits == 0 {
                            "0".to_owned()
                        } else {
                            format!("0.{}", "0".repeat(digits.max(0) as usize))
                        })
                    })
                    .unwrap_or_else(|| right_format.clone());
                // Decimal values arrive as canonical strings through the
                // outcome JSON; numeric metrics render them as numbers like
                // the baseline's openpyxl writer does with Decimal cells.
                let number = if value.is_number() {
                    numeric_text(value)
                } else if precision.is_some() {
                    value.as_str().and_then(|text| text.parse::<f64>().ok())
                } else {
                    None
                };
                if let Some(number) = number {
                    worksheet
                        .write_number_with_format(row_number, column, number, &format)
                        .map_err(|error| error.to_string())?;
                } else {
                    let text = value.as_str().unwrap_or_default();
                    worksheet
                        .write_string_with_format(row_number, column, text, &format)
                        .map_err(|error| error.to_string())?;
                }
            }
            worksheet
                .set_row_height(row_number, 54.0)
                .map_err(|error| error.to_string())?;
            row_number += 1;
        }
        worksheet
            .set_freeze_panes(1, 0)
            .map_err(|error| error.to_string())?;
        worksheet
            .autofilter(
                0,
                0,
                row_number.saturating_sub(1),
                RESULT_HEADERS.len() as u16 - 1,
            )
            .map_err(|error| error.to_string())?;
        for (column, width) in COLUMN_WIDTHS.iter().enumerate() {
            // The pixel interface serializes the exact character width; the
            // character-unit interface adds a 5px serializer padding.
            worksheet
                .set_column_width_pixels(column as u16, ((*width as f64) * 7.0).round() as u32)
                .map_err(|error| error.to_string())?;
        }
        worksheet.set_landscape();
        worksheet.set_paper_size(8);
        worksheet.set_print_fit_to_pages(1, 0);
        worksheet
            .set_repeat_rows(0, 0)
            .map_err(|error| error.to_string())?;

        // The conditional-format sheet is always present in the baseline
        // (openpyxl branch); S5 runs in this slice produce it empty.
        let conditional = workbook
            .add_worksheet()
            .set_name("条件格式触发")
            .map_err(|error| error.to_string())?;
        conditional.set_screen_gridlines(false);
        let cf_headers = [
            "机构名称",
            "报表项目",
            "统一社会信用代码",
            "工作表名",
            "定位单元格",
            "单元格值",
            "判定状态",
            "条件格式类型",
            "规则公式",
            "应用范围",
            "优先级",
            "来源文件",
            "判定说明",
        ];
        let cf_widths = [
            34.0, 26.0, 22.0, 32.0, 16.0, 18.0, 18.0, 20.0, 70.0, 32.0, 12.0, 54.0, 60.0,
        ];
        for (column, header) in cf_headers.iter().enumerate() {
            conditional
                .write_string_with_format(0, column as u16, *header, &header_format)
                .map_err(|error| error.to_string())?;
        }
        conditional
            .set_row_height(0, 30.0)
            .map_err(|error| error.to_string())?;
        conditional
            .set_freeze_panes(1, 0)
            .map_err(|error| error.to_string())?;
        conditional
            .autofilter(0, 0, 0, cf_headers.len() as u16 - 1)
            .map_err(|error| error.to_string())?;
        for (column, width) in cf_widths.iter().enumerate() {
            conditional
                .set_column_width_pixels(column as u16, ((*width as f64) * 7.0).round() as u32)
                .map_err(|error| error.to_string())?;
        }
        conditional.set_landscape();
        conditional.set_paper_size(8);
        conditional.set_print_fit_to_pages(1, 0);
        conditional
            .set_repeat_rows(0, 0)
            .map_err(|error| error.to_string())?;

        if !request.data_quality_warnings.is_empty() {
            let quality = workbook
                .add_worksheet()
                .set_name("数据质量提示")
                .map_err(|error| error.to_string())?;
            quality.set_screen_gridlines(false);
            let warning_headers = [
                "机构名称",
                "统一社会信用代码",
                "来源文件",
                "工作表",
                "单元格/行",
                "提示",
            ];
            let warning_widths = [34.0, 26.0, 45.0, 24.0, 18.0, 90.0];
            for (column, header) in warning_headers.iter().enumerate() {
                quality
                    .write_string_with_format(0, column as u16, *header, &header_format)
                    .map_err(|error| error.to_string())?;
            }
            quality
                .set_row_height(0, 30.0)
                .map_err(|error| error.to_string())?;
            for (row_index, warning) in request.data_quality_warnings.iter().enumerate() {
                let values = [
                    "institution_name",
                    "institution_code",
                    "source_file",
                    "sheet",
                    "location",
                    "message",
                ];
                for (column, key) in values.iter().enumerate() {
                    let text = warning
                        .get(*key)
                        .and_then(JsonValue::as_str)
                        .unwrap_or_default();
                    let format = if column == 1 {
                        &text_format
                    } else {
                        &body_format
                    };
                    quality
                        .write_string_with_format(row_index as u32 + 1, column as u16, text, format)
                        .map_err(|error| error.to_string())?;
                }
                quality
                    .set_row_height(row_index as u32 + 1, 54.0)
                    .map_err(|error| error.to_string())?;
            }
            quality
                .set_freeze_panes(1, 0)
                .map_err(|error| error.to_string())?;
            quality
                .autofilter(
                    0,
                    0,
                    request.data_quality_warnings.len() as u32,
                    warning_headers.len() as u16 - 1,
                )
                .map_err(|error| error.to_string())?;
            for (column, width) in warning_widths.iter().enumerate() {
                quality
                    .set_column_width_pixels(column as u16, ((*width as f64) * 7.0).round() as u32)
                    .map_err(|error| error.to_string())?;
            }
            quality.set_landscape();
            quality.set_paper_size(8);
            quality.set_print_fit_to_pages(1, 0);
            quality
                .set_repeat_rows(0, 0)
                .map_err(|error| error.to_string())?;
        }
        workbook
            .save(&result_path)
            .map_err(|error| error.to_string())?;
    }

    let ledger_path = run_dir.join("运行总账.json");
    let output_files = S5OutputFiles {
        result_workbook: result_path.clone(),
        run_ledger: ledger_path.clone(),
        audit_copies: copies_dir.clone(),
        configuration_copy: config_copy.clone(),
    };
    let output_files_json = serde_json::to_value(&output_files)
        .map_err(|error| format!("序列化输出文件清单失败：{error}"))?;
    let ledger = json!({
        "run_id": request.run_id,
        "period": request.period,
        "institution_code": request.institution_code,
        "institution_name": request.institution_name,
        "configuration_version": request.config_version,
        "configuration_sha256": request.config_sha256_before,
        "run_status": run_status,
        "input_integrity": input_integrity,
        "configuration_integrity": configuration_integrity,
        "input_hashes": input_hashes,
        "status_counts": status_counts,
        "data_quality_warnings": request.data_quality_warnings,
        "audit_copy_manifest": manifest,
        "output_files": output_files_json,
    });
    fs::write(
        &ledger_path,
        serde_json::to_vec_pretty(&ledger).map_err(|error| error.to_string())?,
    )
    .map_err(|error| format!("写入运行总账失败：{error}"))?;

    Ok(S5WriteOutputsSummary {
        run_dir,
        manifest,
        output_files,
        run_status: run_status.to_owned(),
        input_integrity: input_integrity.to_owned(),
        configuration_integrity: configuration_integrity.to_owned(),
        status_counts,
        triggered_rows,
        failure_rows,
        data_quality_warning_count: request.data_quality_warnings.len(),
    })
}

#[cfg(test)]
mod tests {
    use super::{
        count_outcome_statuses, downgrade_successful_outcomes, file_sha256, write_run_outputs,
        S5OutputRequest, S5OutputSource,
    };
    use crate::s5_service::S5RuleOutcome;
    use std::{
        collections::BTreeMap,
        fs,
        path::PathBuf,
        time::{SystemTime, UNIX_EPOCH},
    };

    fn test_outcome(status: &str) -> S5RuleOutcome {
        S5RuleOutcome {
            status: status.to_owned(),
            triggered: match status {
                "TRIGGERED" => Some(true),
                "PASSED" => Some(false),
                _ => None,
            },
            error_code: (status == "RETRIEVAL_FAILED").then(|| "SOURCE_READ_FAILED".to_owned()),
            ..S5RuleOutcome::default()
        }
    }

    #[test]
    fn hash_integrity_downgrade_matches_python_result_mutation() {
        let mut outcomes = vec![
            test_outcome("TRIGGERED"),
            test_outcome("PASSED"),
            test_outcome("RETRIEVAL_FAILED"),
        ];
        downgrade_successful_outcomes(&mut outcomes);
        assert_eq!(
            outcomes
                .iter()
                .map(|item| item.status.as_str())
                .collect::<Vec<_>>(),
            vec![
                "CALCULATION_FAILED",
                "CALCULATION_FAILED",
                "RETRIEVAL_FAILED"
            ]
        );
        for outcome in &outcomes[..2] {
            assert_eq!(outcome.triggered, None);
            assert_eq!(
                outcome.error_code.as_deref(),
                Some("INPUT_CHANGED_DURING_RUN")
            );
            assert_eq!(
                outcome.error_message,
                "审核结果文件生成期间输入或配置哈希发生变化，结论不可用"
            );
        }
        assert_eq!(
            outcomes[2].error_code.as_deref(),
            Some("SOURCE_READ_FAILED")
        );
    }

    #[test]
    fn output_writer_downgrades_successes_when_current_input_changed() {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let root = std::env::temp_dir().join(format!(
            "audit_s5_snapshot_guard_{}_{}",
            std::process::id(),
            nonce
        ));
        fs::create_dir_all(&root).expect("create test root");
        let source = root.join("source.xlsx");
        let config = root.join("config.xlsx");
        let output = root.join("output");
        fs::write(&source, b"source snapshot changed").expect("write source");
        fs::write(&config, b"config snapshot").expect("write config");
        let config_sha = file_sha256(&config).expect("hash config");
        let source_sha = "0".repeat(64);
        let outcomes = vec![
            test_outcome("TRIGGERED"),
            test_outcome("PASSED"),
            test_outcome("RETRIEVAL_FAILED"),
        ];
        let counts = count_outcome_statuses(&outcomes);
        let request = S5OutputRequest {
            run_id: "changed-input",
            period: "2026-07-31",
            institution_code: "123",
            institution_name: "sample",
            outcomes: &outcomes,
            status_counts: &counts,
            data_quality_warnings: &[],
            source_path: &source,
            source_id: "CURRENT",
            source_sha256_before: &source_sha,
            additional_sources: &[],
            selected_templates: &[],
            config_path: &config,
            config_sha256_before: &config_sha,
            config_version: "test",
        };

        let summary = write_run_outputs(&request, &output).expect("write downgraded result");
        assert_eq!(summary.run_status, "INPUT_CHANGED_DURING_RUN");
        assert_eq!(summary.input_integrity, "CHANGED_DURING_RUN");
        assert_eq!(summary.configuration_integrity, "UNCHANGED");
        assert_eq!(summary.status_counts["CALCULATION_FAILED"], 2);
        assert_eq!(summary.status_counts["RETRIEVAL_FAILED"], 1);
        assert_eq!(summary.triggered_rows, 0);
        assert_eq!(summary.failure_rows, 3);
        assert!(summary.output_files.result_workbook.is_file());
        let ledger: serde_json::Value = serde_json::from_slice(
            &fs::read(&summary.output_files.run_ledger).expect("read run ledger"),
        )
        .expect("parse run ledger");
        assert_eq!(ledger["run_status"], "INPUT_CHANGED_DURING_RUN");
        assert_eq!(ledger["input_integrity"], "CHANGED_DURING_RUN");
        assert_eq!(ledger["configuration_integrity"], "UNCHANGED");
        assert_eq!(ledger["status_counts"]["CALCULATION_FAILED"], 2);
        assert_eq!(
            ledger["input_hashes"]["current:CURRENT"]["sha256_before"],
            source_sha
        );
        assert_eq!(
            ledger["input_hashes"]["current:CURRENT"]["sha256_after"],
            file_sha256(&source).expect("hash changed source")
        );
        assert_eq!(ledger["audit_copy_manifest"][0]["matches_snapshot"], false);
        fs::remove_dir_all(root).expect("remove test root");
    }

    #[test]
    fn output_writer_records_missing_inputs_without_copying_them() {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let root = std::env::temp_dir().join(format!(
            "audit_s5_missing_inputs_downgrade_{}_{}",
            std::process::id(),
            nonce
        ));
        fs::create_dir_all(&root).expect("create test root");
        let current = root.join("current.xlsx");
        let previous = root.join("previous.xlsx");
        let template = root.join("layout.xlsx");
        let config = root.join("config.xlsx");
        let output = root.join("output");
        fs::write(&current, b"current snapshot").expect("write current source");
        fs::write(&previous, b"previous snapshot").expect("write previous source");
        fs::write(&template, b"layout snapshot").expect("write layout template");
        fs::write(&config, b"configuration snapshot").expect("write config");
        let current_sha = file_sha256(&current).expect("hash current");
        let previous_sha = file_sha256(&previous).expect("hash previous");
        let template_sha = file_sha256(&template).expect("hash layout");
        let config_sha = file_sha256(&config).expect("hash config");
        let outcomes = vec![test_outcome("TRIGGERED")];
        let counts = count_outcome_statuses(&outcomes);
        let additional_sources = [S5OutputSource {
            role: "previous",
            path: &previous,
            source_id: "PREVIOUS",
            sha256_before: &previous_sha,
        }];
        let selected_templates = [S5OutputSource {
            role: "layout",
            path: &template,
            source_id: "LAYOUT",
            sha256_before: &template_sha,
        }];
        let request = S5OutputRequest {
            run_id: "missing-inputs",
            period: "2026-07-31",
            institution_code: "123",
            institution_name: "sample",
            outcomes: &outcomes,
            status_counts: &counts,
            data_quality_warnings: &[],
            source_path: &current,
            source_id: "CURRENT",
            source_sha256_before: &current_sha,
            additional_sources: &additional_sources,
            selected_templates: &selected_templates,
            config_path: &config,
            config_sha256_before: &config_sha,
            config_version: "test",
        };
        fs::remove_file(&current).expect("remove current source");
        fs::remove_file(&previous).expect("remove previous source");
        fs::remove_file(&template).expect("remove layout template");
        fs::remove_file(&config).expect("remove config");

        let summary = write_run_outputs(&request, &output).expect("write downgraded result");
        assert_eq!(summary.run_status, "INPUT_CHANGED_DURING_RUN");
        assert_eq!(summary.input_integrity, "CHANGED_DURING_RUN");
        assert_eq!(summary.configuration_integrity, "CHANGED_DURING_RUN");
        assert_eq!(summary.status_counts["CALCULATION_FAILED"], 1);
        assert_eq!(summary.triggered_rows, 0);
        assert_eq!(summary.failure_rows, 1);
        assert!(summary.output_files.configuration_copy.is_none());
        assert!(summary.output_files.result_workbook.is_file());
        assert!(summary.manifest.is_empty());

        let ledger: serde_json::Value = serde_json::from_slice(
            &fs::read(&summary.output_files.run_ledger).expect("read run ledger"),
        )
        .expect("parse run ledger");
        for key in ["current:CURRENT", "previous:PREVIOUS", "layout:LAYOUT"] {
            assert_eq!(ledger["input_hashes"][key]["sha256_after"], "MISSING");
        }
        assert_eq!(
            ledger["input_hashes"]["configuration"]["sha256_after"],
            "MISSING"
        );
        assert!(ledger["output_files"].get("configuration_copy").is_none());
        fs::remove_dir_all(root).expect("remove test root");
    }

    #[test]
    fn output_writer_downgrades_successes_when_previous_input_changed() {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let root = std::env::temp_dir().join(format!(
            "audit_s5_previous_snapshot_guard_{}_{}",
            std::process::id(),
            nonce
        ));
        fs::create_dir_all(&root).expect("create test root");
        let current = root.join("current.xlsx");
        let previous = root.join("previous.xlsx");
        let config = root.join("config.xlsx");
        let output = root.join("output");
        fs::write(&current, b"current source").expect("write current source");
        fs::write(&previous, b"previous source changed").expect("write previous source");
        fs::write(&config, b"config snapshot").expect("write config");
        let config_sha = file_sha256(&config).expect("hash config");
        let current_sha = file_sha256(&current).expect("hash current");
        let stale_previous_sha = "0".repeat(64);
        let outcomes = vec![test_outcome("TRIGGERED")];
        let counts = count_outcome_statuses(&outcomes);
        let additional_sources = [S5OutputSource {
            role: "previous",
            path: &previous,
            source_id: "PREVIOUS",
            sha256_before: &stale_previous_sha,
        }];
        let request = S5OutputRequest {
            run_id: "changed-previous-input",
            period: "2026-07-31",
            institution_code: "123",
            institution_name: "sample",
            outcomes: &outcomes,
            status_counts: &counts,
            data_quality_warnings: &[],
            source_path: &current,
            source_id: "CURRENT",
            source_sha256_before: &current_sha,
            additional_sources: &additional_sources,
            selected_templates: &[],
            config_path: &config,
            config_sha256_before: &config_sha,
            config_version: "test",
        };

        let summary = write_run_outputs(&request, &output).expect("write downgraded result");
        assert_eq!(summary.run_status, "INPUT_CHANGED_DURING_RUN");
        assert_eq!(summary.input_integrity, "CHANGED_DURING_RUN");
        assert_eq!(summary.configuration_integrity, "UNCHANGED");
        assert_eq!(summary.status_counts["CALCULATION_FAILED"], 1);
        assert_eq!(summary.triggered_rows, 0);
        assert_eq!(summary.failure_rows, 1);
        let ledger: serde_json::Value = serde_json::from_slice(
            &fs::read(&summary.output_files.run_ledger).expect("read run ledger"),
        )
        .expect("parse run ledger");
        assert_eq!(
            ledger["input_hashes"]["previous:PREVIOUS"]["sha256_before"],
            stale_previous_sha
        );
        assert_eq!(
            ledger["input_hashes"]["previous:PREVIOUS"]["sha256_after"],
            file_sha256(&previous).expect("hash changed previous source")
        );
        assert_eq!(ledger["audit_copy_manifest"][1]["matches_snapshot"], false);
        fs::remove_dir_all(root).expect("remove test root");
    }

    #[test]
    fn output_writer_downgrades_successes_when_configuration_changed() {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let root = std::env::temp_dir().join(format!(
            "audit_s5_config_snapshot_downgrade_{}_{}",
            std::process::id(),
            nonce
        ));
        fs::create_dir_all(&root).expect("create test root");
        let source = root.join("source.xlsx");
        let config = root.join("config.xlsx");
        let output = root.join("output");
        fs::write(&source, b"source snapshot").expect("write source");
        fs::write(&config, b"configuration changed").expect("write changed config");
        let source_sha = file_sha256(&source).expect("hash source");
        let stale_config_sha = "0".repeat(64);
        let outcomes = vec![test_outcome("TRIGGERED")];
        let counts = count_outcome_statuses(&outcomes);
        let request = S5OutputRequest {
            run_id: "changed-configuration",
            period: "2026-07-31",
            institution_code: "123",
            institution_name: "sample",
            outcomes: &outcomes,
            status_counts: &counts,
            data_quality_warnings: &[],
            source_path: &source,
            source_id: "CURRENT",
            source_sha256_before: &source_sha,
            additional_sources: &[],
            selected_templates: &[],
            config_path: &config,
            config_sha256_before: &stale_config_sha,
            config_version: "test",
        };

        let summary = write_run_outputs(&request, &output).expect("write downgraded result");
        assert_eq!(summary.run_status, "INPUT_CHANGED_DURING_RUN");
        assert_eq!(summary.input_integrity, "UNCHANGED");
        assert_eq!(summary.configuration_integrity, "CHANGED_DURING_RUN");
        assert_eq!(summary.status_counts["CALCULATION_FAILED"], 1);
        assert_eq!(summary.triggered_rows, 0);
        assert_eq!(summary.failure_rows, 1);
        let ledger: serde_json::Value = serde_json::from_slice(
            &fs::read(&summary.output_files.run_ledger).expect("read run ledger"),
        )
        .expect("parse run ledger");
        assert_eq!(ledger["input_integrity"], "UNCHANGED");
        assert_eq!(ledger["configuration_integrity"], "CHANGED_DURING_RUN");
        assert_eq!(
            ledger["input_hashes"]["configuration"]["sha256_before"],
            stale_config_sha
        );
        assert_eq!(
            ledger["input_hashes"]["configuration"]["sha256_after"],
            file_sha256(&config).expect("hash changed configuration")
        );
        assert_eq!(ledger["audit_copy_manifest"][1]["matches_snapshot"], false);
        fs::remove_dir_all(root).expect("remove test root");
    }

    #[test]
    fn output_writer_copies_selected_layout_template_and_records_snapshot() {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        let root = std::env::temp_dir().join(format!(
            "audit_s5_layout_copy_{}_{}",
            std::process::id(),
            nonce
        ));
        fs::create_dir_all(&root).expect("create test root");
        let current = root.join("current.xlsx");
        let template = root.join("template.xlsx");
        let config = root.join("config.xlsx");
        let output = root.join("output");
        fs::write(&current, b"current workbook snapshot").expect("write current source");
        fs::write(&template, b"selected layout snapshot").expect("write selected template");
        fs::write(&config, b"configuration snapshot").expect("write config");
        let current_sha = file_sha256(&current).expect("hash current source");
        let template_sha = file_sha256(&template).expect("hash selected template");
        let config_sha = file_sha256(&config).expect("hash config");
        let counts = BTreeMap::new();
        let selected_templates = [S5OutputSource {
            role: "layout",
            path: &template,
            source_id: "LAYOUT_0123456789ab",
            sha256_before: &template_sha,
        }];
        let request = S5OutputRequest {
            run_id: "selected-layout-template",
            period: "2026-07-31",
            institution_code: "123",
            institution_name: "sample",
            outcomes: &[],
            status_counts: &counts,
            data_quality_warnings: &[],
            source_path: &current,
            source_id: "CURRENT",
            source_sha256_before: &current_sha,
            additional_sources: &[],
            selected_templates: &selected_templates,
            config_path: &config,
            config_sha256_before: &config_sha,
            config_version: "test",
        };

        let summary = write_run_outputs(&request, &output).expect("write run outputs");
        assert_eq!(summary.run_status, "COMPLETED");
        assert_eq!(summary.input_integrity, "UNCHANGED");
        assert_eq!(summary.configuration_integrity, "UNCHANGED");
        assert_eq!(
            summary
                .manifest
                .iter()
                .map(|entry| entry.role.as_str())
                .collect::<Vec<_>>(),
            vec!["current", "layout", "configuration"]
        );
        let layout_entry = summary
            .manifest
            .iter()
            .find(|entry| entry.role == "layout")
            .expect("layout manifest entry");
        let layout_copy = PathBuf::from(&layout_entry.copy_path);
        assert_eq!(
            layout_copy.file_name().unwrap(),
            "layout_LAYOUT_0123456789ab_template.xlsx"
        );
        assert_eq!(
            fs::read(layout_copy).expect("read copied template"),
            b"selected layout snapshot"
        );
        assert!(layout_entry.matches_snapshot);
        let ledger: serde_json::Value = serde_json::from_slice(
            &fs::read(&summary.output_files.run_ledger).expect("read run ledger"),
        )
        .expect("parse run ledger");
        assert_eq!(ledger["run_status"], "COMPLETED");
        assert_eq!(ledger["input_integrity"], "UNCHANGED");
        assert_eq!(ledger["configuration_integrity"], "UNCHANGED");
        assert_eq!(
            ledger["input_hashes"]["current:CURRENT"]["sha256_before"],
            current_sha
        );
        assert_eq!(
            ledger["input_hashes"]["layout:LAYOUT_0123456789ab"]["sha256_before"],
            template_sha
        );
        assert_eq!(
            ledger["input_hashes"]["configuration"]["sha256_before"],
            config_sha
        );
        fs::remove_dir_all(root).expect("remove test root");
    }
}
