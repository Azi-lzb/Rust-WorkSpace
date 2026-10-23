//! S1 template recommendation migrated from `base_audit.discovery`.
//!
//! This covers workbook discovery and worksheet-set matching only. It does
//! not inspect headers, rule areas, formulas, or run an audit.

use crate::discovery::{source_workbooks, ScanDepth};
use anyhow::Result;
use excel_reader::read_xlsx_sheet_names_ooxml;
use serde::Serialize;
use std::{
    collections::{BTreeMap, BTreeSet},
    fs,
    path::{Path, PathBuf},
};

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TemplateProfile {
    pub path: PathBuf,
    pub sheet_names: Vec<String>,
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct TemplateSuggestion {
    pub template_path: Option<PathBuf>,
    pub confidence: f64,
    pub matched: bool,
    pub details: String,
    pub alternatives: Vec<(String, f64)>,
}

/// Recommend a template using the source project's worksheet-set contract.
///
/// The directory variant scans xlsx files and reads their sheet names with
/// the shared namespace-aware OOXML metadata adapter. It does not create the
/// source project's writable profile index, so probing real template
/// directories stays read-only.
pub fn recommend_template(
    template_dir: &Path,
    input_dir: &Path,
    depth: ScanDepth,
) -> Result<TemplateSuggestion> {
    let source_files = source_workbooks(input_dir, depth)?;
    if source_files.is_empty() {
        return Ok(unmatched("源数据目录中没有可识别的 .xlsx 文件"));
    }

    let mut template_paths = if template_dir.is_dir() {
        fs::read_dir(template_dir)?
            .filter_map(std::result::Result::ok)
            .map(|entry| entry.path())
            .filter(|path| {
                path.extension()
                    .and_then(|extension| extension.to_str())
                    .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
                    && !path
                        .file_name()
                        .and_then(|name| name.to_str())
                        .is_some_and(|name| name.starts_with("~$"))
            })
            .collect::<Vec<_>>()
    } else {
        Vec::new()
    };
    template_paths.sort();

    let templates = template_paths
        .into_iter()
        .map(|path| TemplateProfile {
            sheet_names: read_xlsx_sheet_names_ooxml(&path).unwrap_or_default(),
            path,
        })
        .collect::<Vec<_>>();
    if templates.is_empty() {
        return Ok(unmatched("模板文件目录中没有正式模板"));
    }

    let source_sheets = source_files
        .iter()
        .map(|path| read_xlsx_sheet_names_ooxml(path).unwrap_or_default())
        .collect::<Vec<_>>();
    Ok(recommend_from_sheet_sets(&templates, &source_sheets))
}

/// Pure worksheet-set recommendation logic, suitable for parity probes.
pub fn recommend_from_sheet_sets(
    templates: &[TemplateProfile],
    source_workbooks: &[Vec<String>],
) -> TemplateSuggestion {
    if source_workbooks.is_empty() {
        return unmatched("源数据目录中没有可识别的 .xlsx 文件");
    }
    if templates.is_empty() {
        return unmatched("模板文件目录中没有正式模板");
    }

    let mut per_file_choices: BTreeMap<PathBuf, Vec<f64>> = BTreeMap::new();
    for source_sheets in source_workbooks {
        let ranked = ranked_templates(templates, source_sheets);
        let (top_score, top_template) = &ranked[0];
        let covered = all_source_sheets_in_template(&top_template.sheet_names, source_sheets);
        let margin = if ranked.len() > 1 {
            top_score - ranked[1].0
        } else {
            f64::INFINITY
        };
        if (*top_score >= 0.90 || covered) && (ranked.len() == 1 || margin >= 0.04) {
            per_file_choices
                .entry(top_template.path.clone())
                .or_default()
                .push(if covered { 1.0 } else { *top_score });
        }
    }

    if per_file_choices.len() > 1 {
        return unmatched("源数据目录可能包含多种报表，请按报表类型分目录后再审核");
    }

    if let Some((selected_path, selected_scores)) = per_file_choices.into_iter().next() {
        let selected = templates
            .iter()
            .find(|profile| profile.path == selected_path)
            .expect("a selected template must be in the profile list");
        let confidence = selected_scores.iter().sum::<f64>() / selected_scores.len() as f64;
        let match_detail = if confidence >= 1.0 {
            "源文件工作表全部包含于模板".to_owned()
        } else {
            format!("工作表集合匹配率 {:.0}%", confidence * 100.0)
        };
        let details = format!(
            "自动匹配：{}（{} 个源文件；{}；已忽略未能识别的非报送工作簿）",
            file_name(&selected.path),
            selected_scores.len(),
            match_detail
        );
        return TemplateSuggestion {
            template_path: Some(selected.path.clone()),
            confidence,
            matched: true,
            details,
            alternatives: vec![(file_name(&selected.path), round3(confidence))],
        };
    }

    let source_sheets = &source_workbooks[0];
    let ranked = ranked_templates(templates, source_sheets);
    let (top_score, top_template) = &ranked[0];
    let runner_up = ranked.get(1).map_or(0.0, |candidate| candidate.0);
    let top_covered = all_source_sheets_in_template(&top_template.sheet_names, source_sheets);
    let score_gap = top_score - runner_up;
    let structure_decisive = (*top_score >= 0.95 && (ranked.len() == 1 || score_gap >= 0.20))
        || (top_covered && (ranked.len() == 1 || score_gap >= 0.20));
    let unique_enough = score_gap >= 0.05 || structure_decisive;
    let matched = (*top_score >= 0.90 || top_covered) && unique_enough;
    let alternatives = ranked
        .iter()
        .take(3)
        .map(|(score, profile)| (file_name(&profile.path), round3(*score)))
        .collect::<Vec<_>>();

    if matched {
        let details = if top_covered {
            format!(
                "自动匹配：{}（源文件工作表均包含于模板；模板中未出现在源文件里的其他工作表不影响匹配）",
                file_name(&top_template.path)
            )
        } else {
            format!(
                "自动匹配：{}（工作表集合匹配率 {:.0}%，取最高匹配率候选；前三候选：{}）",
                file_name(&top_template.path),
                top_score * 100.0,
                candidate_text(&ranked)
            )
        };
        return TemplateSuggestion {
            template_path: Some(top_template.path.clone()),
            confidence: if top_covered { 1.0 } else { *top_score },
            matched: true,
            details,
            alternatives,
        };
    }

    let details = if *top_score >= 0.80 && !unique_enough {
        format!(
            "有多个相近模板，已停止自动选择，请手动确认。候选：{}",
            candidate_text(&ranked)
        )
    } else {
        "未找到可信度足够的模板，请手动选择".to_owned()
    };
    TemplateSuggestion {
        template_path: None,
        confidence: *top_score,
        matched: false,
        details,
        alternatives,
    }
}

pub fn sheet_score(template_sheets: &[String], source_sheets: &[String]) -> f64 {
    let expected = template_sheets
        .iter()
        .filter(|name| name.as_str() != "审核规则")
        .cloned()
        .collect::<BTreeSet<_>>();
    let actual = source_sheets.iter().cloned().collect::<BTreeSet<_>>();
    if expected.is_empty() || actual.is_empty() {
        return 0.0;
    }
    let common = expected.intersection(&actual).count() as f64;
    let source_coverage = common / actual.len() as f64;
    let template_precision = common / expected.len() as f64;
    0.8 * source_coverage + 0.2 * template_precision
}

pub fn all_source_sheets_in_template(template_sheets: &[String], source_sheets: &[String]) -> bool {
    let template_data = template_sheets
        .iter()
        .filter(|name| !matches!(name.as_str(), "审核规则" | "兜底模板"))
        .collect::<BTreeSet<_>>();
    let source_data = source_sheets
        .iter()
        .filter(|name| !matches!(name.as_str(), "审核规则" | "兜底模板"))
        .collect::<BTreeSet<_>>();
    !source_data.is_empty() && source_data.is_subset(&template_data)
}

fn ranked_templates<'a>(
    templates: &'a [TemplateProfile],
    source_sheets: &[String],
) -> Vec<(f64, &'a TemplateProfile)> {
    let mut ranked = templates
        .iter()
        .map(|profile| (sheet_score(&profile.sheet_names, source_sheets), profile))
        .collect::<Vec<_>>();
    ranked.sort_by(|left, right| {
        right
            .0
            .total_cmp(&left.0)
            .then_with(|| left.1.path.cmp(&right.1.path))
    });
    ranked
}

fn candidate_text(ranked: &[(f64, &TemplateProfile)]) -> String {
    ranked
        .iter()
        .take(3)
        .map(|(score, profile)| {
            format!(
                "{}（工作表集合 {:.0}%）",
                file_name(&profile.path),
                score * 100.0
            )
        })
        .collect::<Vec<_>>()
        .join("；")
}

fn file_name(path: &Path) -> String {
    path.file_name()
        .map(|name| name.to_string_lossy().into_owned())
        .unwrap_or_else(|| path.display().to_string())
}

fn round3(value: f64) -> f64 {
    (value * 1000.0).round() / 1000.0
}

fn unmatched(details: &str) -> TemplateSuggestion {
    TemplateSuggestion {
        template_path: None,
        confidence: 0.0,
        matched: false,
        details: details.to_owned(),
        alternatives: Vec::new(),
    }
}
