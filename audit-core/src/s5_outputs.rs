//! S5 run outputs (`outputs.write_run_outputs` port, single-business shape).
//!
//! Produces the run directory: audit copies of the source workbook and the
//! configuration with a sha256 manifest, the 审核结果.xlsx result workbook
//! (triggered + failure rows, severity-classified error types, evidence-based
//! locations and hyperlinks to the copies, per-column precision formats, the
//! Python `_style_table` look), and the 运行总账.json ledger. Conditional
//! format and data-quality tables belong to slices that produce those inputs.

use crate::s5_service::S5RuleOutcome;
use excel_writer::{Color, Format, FormatAlign, FormatUnderline, Workbook};
use serde_json::{json, Value as JsonValue};
use sha2::{Digest, Sha256};
use std::{
    collections::BTreeMap,
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
    pub source_path: &'a Path,
    pub source_id: &'a str,
    pub source_sha256_before: &'a str,
    pub additional_sources: &'a [S5OutputSource<'a>],
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
    pub configuration_copy: PathBuf,
}

#[derive(Clone, Debug, PartialEq, serde::Serialize)]
pub struct S5WriteOutputsSummary {
    pub run_dir: PathBuf,
    pub manifest: Vec<S5OutputManifestEntry>,
    pub output_files: S5OutputFiles,
    pub triggered_rows: usize,
    pub failure_rows: usize,
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

fn verify_input_snapshot(path: &Path, expected: &str, label: &str) -> Result<(), String> {
    let actual = file_sha256(path)?;
    if actual != expected {
        return Err(format!("{label}在审核期间发生变化，已停止生成结果"));
    }
    Ok(())
}

fn verify_copied_snapshot(
    run_dir: &Path,
    label: &str,
    expected: &str,
    actual: Result<String, String>,
    copy: Result<String, String>,
) -> Result<(String, String), String> {
    let result = (|| {
        let actual = actual?;
        let copy = copy?;
        if actual != expected || copy != actual {
            return Err(format!(
                "{label}在审核期间发生变化或审核副本校验失败，已停止生成结果"
            ));
        }
        Ok((actual, copy))
    })();
    match result {
        Ok(hashes) => Ok(hashes),
        Err(message) => {
            let cleanup = fs::remove_dir_all(run_dir);
            match cleanup {
                Ok(()) => Err(message),
                Err(error) => Err(format!(
                    "{message}；未验证输出目录清理失败 {}：{error}",
                    run_dir.display()
                )),
            }
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
/// current-source shape without conditional-format or data-quality tables.
pub fn write_run_outputs(
    request: &S5OutputRequest<'_>,
    output_root: &Path,
) -> Result<S5WriteOutputsSummary, String> {
    verify_input_snapshot(
        request.source_path,
        request.source_sha256_before,
        "本期数据源",
    )?;
    verify_input_snapshot(
        request.config_path,
        request.config_sha256_before,
        "审核配置",
    )?;
    for source in request.additional_sources {
        let label = match source.role {
            "previous" => "上期数据源",
            "external" => "外部数据源",
            _ => return Err(format!("不支持的附加来源角色：{}", source.role)),
        };
        verify_input_snapshot(source.path, source.sha256_before, label)?;
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
    let (sha_after, sha_copy) = verify_copied_snapshot(
        &run_dir,
        "本期数据源",
        request.source_sha256_before,
        file_sha256(request.source_path),
        file_sha256(&source_copy),
    )?;
    manifest.push(S5OutputManifestEntry {
        role: "current".to_owned(),
        source_id: request.source_id.to_owned(),
        source_path: request.source_path.display().to_string(),
        copy_path: source_copy.display().to_string(),
        sha256_before: request.source_sha256_before.to_owned(),
        sha256_after: sha_after,
        sha256_copy: sha_copy,
        matches_snapshot: true,
    });
    copy_by_source.insert(
        request.source_path.display().to_string(),
        source_copy.clone(),
    );

    for source in request.additional_sources {
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
        let label = match source.role {
            "previous" => "上期数据源",
            "external" => "外部数据源",
            _ => return Err(format!("不支持的附加来源角色：{}", source.role)),
        };
        let (sha_after, sha_copy) = verify_copied_snapshot(
            &run_dir,
            label,
            source.sha256_before,
            file_sha256(source.path),
            file_sha256(&source_copy),
        )?;
        manifest.push(S5OutputManifestEntry {
            role: source.role.to_owned(),
            source_id: source.source_id.to_owned(),
            source_path: source.path.display().to_string(),
            copy_path: source_copy.display().to_string(),
            sha256_before: source.sha256_before.to_owned(),
            sha256_after: sha_after,
            sha256_copy: sha_copy,
            matches_snapshot: true,
        });
        copy_by_source.insert(source.path.display().to_string(), source_copy);
    }

    let config_copy = copies_dir.join(format!(
        "配置_{}{}",
        safe_name(&path_stem(request.config_path)),
        request
            .config_path
            .extension()
            .map(|extension| format!(".{}", extension.to_string_lossy().to_lowercase()))
            .unwrap_or_default()
    ));
    fs::copy(request.config_path, &config_copy)
        .map_err(|error| format!("复制配置失败：{error}"))?;
    let (sha_after, sha_copy) = verify_copied_snapshot(
        &run_dir,
        "审核配置",
        request.config_sha256_before,
        file_sha256(request.config_path),
        file_sha256(&config_copy),
    )?;
    manifest.push(S5OutputManifestEntry {
        role: "configuration".to_owned(),
        source_id: "configuration".to_owned(),
        source_path: request.config_path.display().to_string(),
        copy_path: config_copy.display().to_string(),
        sha256_before: request.config_sha256_before.to_owned(),
        sha256_after: sha_after,
        sha256_copy: sha_copy,
        matches_snapshot: true,
    });

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
        for item in request.outcomes {
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
    let ledger = json!({
        "run_id": request.run_id,
        "period": request.period,
        "institution_code": request.institution_code,
        "institution_name": request.institution_name,
        "configuration_version": request.config_version,
        "configuration_sha256": request.config_sha256_before,
        "run_status": "COMPLETED",
        "status_counts": request.status_counts,
        "audit_copy_manifest": manifest,
        "output_files": {
            "result_workbook": output_files.result_workbook.display().to_string(),
            "run_ledger": output_files.run_ledger.display().to_string(),
            "audit_copies": output_files.audit_copies.display().to_string(),
            "configuration_copy": output_files.configuration_copy.display().to_string(),
        },
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
        triggered_rows,
        failure_rows,
    })
}

#[cfg(test)]
mod tests {
    use super::{file_sha256, write_run_outputs, S5OutputRequest, S5OutputSource};
    use std::{
        collections::BTreeMap,
        fs,
        time::{SystemTime, UNIX_EPOCH},
    };

    #[test]
    fn output_writer_rejects_changed_inputs_before_creating_run_outputs() {
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
        let counts = BTreeMap::new();
        let request = S5OutputRequest {
            run_id: "changed-input",
            period: "2026-07-31",
            institution_code: "123",
            institution_name: "sample",
            outcomes: &[],
            status_counts: &counts,
            source_path: &source,
            source_id: "CURRENT",
            source_sha256_before: &source_sha,
            additional_sources: &[],
            config_path: &config,
            config_sha256_before: &config_sha,
            config_version: "test",
        };

        let error = write_run_outputs(&request, &output)
            .expect_err("changed source must not produce a completed output run");
        assert!(error.contains("本期数据源在审核期间发生变化"));
        assert!(
            !output.exists(),
            "rejected input must not create a run directory"
        );
        fs::remove_dir_all(root).expect("remove test root");
    }
    #[test]
    fn output_writer_rejects_changed_previous_input_before_creating_run_outputs() {
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
        let counts = BTreeMap::new();
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
            outcomes: &[],
            status_counts: &counts,
            source_path: &current,
            source_id: "CURRENT",
            source_sha256_before: &current_sha,
            additional_sources: &additional_sources,
            config_path: &config,
            config_sha256_before: &config_sha,
            config_version: "test",
        };

        let error = write_run_outputs(&request, &output)
            .expect_err("changed previous source must not produce a completed output run");
        assert!(error.contains("上期数据源在审核期间发生变化"));
        assert!(
            !output.exists(),
            "rejected previous input must not create a run directory"
        );
        fs::remove_dir_all(root).expect("remove test root");
    }
}
