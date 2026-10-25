//! S5 business identification and the per-file classification list.
//!
//! Faithful Rust port of
//! `s5_configured_statistics.recognition.match_s5_business` and
//! `classify_s5_files`. Business identification reads the loaded
//! [`Configuration`] (from [`crate::s5_config`]); classification reuses the
//! parity-proven sheet scoring from [`crate::s5_template_matching`] and the
//! period hint from [`crate::s5_period_recognition`]. Error messages mirror
//! the Python source verbatim.
//!
//! Reading delta: Python `read_xlsx_sheet_names` swallows file-level errors
//! but propagates XML parse errors in `match_s5_business`; this port maps all
//! read failures to an empty sheet set (real corpora contain readable
//! packages). Python's `TemplateCatalog` also persists a JSON index beside the
//! template directory; this port never writes one, which is behavior-neutral
//! for the produced rows.

use crate::s5_config::{load_configuration, Configuration};
use crate::s5_period_recognition::s5_file_period_hint;
use crate::s5_template_matching::{all_source_sheets_in_template, sheet_score};
use excel_reader::read_xlsx_sheet_names_ooxml;
use std::collections::BTreeSet;
use std::path::{Path, PathBuf};

#[derive(Clone, Debug, PartialEq, serde::Serialize)]
pub struct S5FileClassification {
    pub name: String,
    pub template: String,
    pub match_status: String,
    pub business: String,
    pub period: String,
    pub match_detail: String,
    pub matched_sheets: usize,
    pub source_sheet_count: usize,
    pub template_sheet_count: usize,
}

fn read_sheet_names(path: &Path) -> Vec<String> {
    read_xlsx_sheet_names_ooxml(path).unwrap_or_default()
}

fn sheet_name_set(path: &Path) -> BTreeSet<String> {
    read_sheet_names(path).into_iter().collect()
}

fn path_stem(path: &Path) -> String {
    path.file_stem()
        .map(|value| value.to_string_lossy().into_owned())
        .unwrap_or_default()
}

fn path_name(path: &Path) -> String {
    path.file_name()
        .map(|value| value.to_string_lossy().into_owned())
        .unwrap_or_default()
}

fn parent_name(path: &Path) -> String {
    path.parent()
        .and_then(|value| value.file_name())
        .map(|value| value.to_string_lossy().into_owned())
        .unwrap_or_default()
}

/// Mirror of `match_s5_business`: identify the one business whose required
/// worksheet set fits every source file, with a longest text-match fallback.
pub fn match_s5_business(
    configuration: &Configuration,
    files: &[PathBuf],
    template_path: &Path,
) -> Result<String, String> {
    let sheet_sets: Vec<BTreeSet<String>> = files.iter().map(|path| sheet_name_set(path)).collect();
    let mut businesses: Vec<String> = configuration
        .rules
        .iter()
        .map(|rule| rule.business.clone())
        .collect::<BTreeSet<_>>()
        .into_iter()
        .collect();
    businesses.sort();
    let mut candidates: Vec<String> = Vec::new();
    for business in &businesses {
        let mut sheets: BTreeSet<String> = configuration
            .rules
            .iter()
            .filter(|rule| &rule.business == business && !rule.default_sheet.is_empty())
            .map(|rule| rule.default_sheet.clone())
            .collect();
        for metric in &configuration.metrics {
            if &metric.business == business && !metric.sheet.is_empty() {
                sheets.insert(metric.sheet.clone());
            }
        }
        if sheets.is_empty() || sheet_sets.iter().all(|source| sheets.is_subset(source)) {
            candidates.push(business.clone());
        }
    }
    if candidates.len() == 1 {
        return Ok(candidates.remove(0));
    }
    let mut text = path_stem(template_path);
    for path in files {
        text.push(' ');
        text.push_str(&parent_name(path));
        text.push(' ');
        text.push_str(&path_stem(path));
    }
    let mut named: Vec<String> = candidates
        .iter()
        .filter(|name| !name.is_empty() && text.contains(name.as_str()))
        .cloned()
        .collect();
    if !named.is_empty() {
        let longest = named
            .iter()
            .map(|name| name.chars().count())
            .max()
            .unwrap_or(0);
        named.retain(|name| name.chars().count() == longest);
    }
    if named.len() == 1 {
        return Ok(named.remove(0));
    }
    if candidates.is_empty() {
        return Err("源文件工作表与 S5 配置中的业务不对应，请核实配置与匹配模板".to_owned());
    }
    Err(format!(
        "无法唯一识别业务：{}；请完善配置的业务/工作表对应关系或在文件名中标明业务",
        candidates.join("、")
    ))
}

/// Mirror of `classify_s5_files`: build the S5 per-file worksheet-match,
/// business, and date-hint rows. `template_path` (explicit) takes precedence
/// over `template_dir` (catalog without index persistence). The `input_dir`
/// parameter of the Python signature is ignored by the source hint logic and
/// therefore omitted here.
pub fn classify_s5_files(
    files: &[PathBuf],
    template_dir: Option<&Path>,
    template_path: Option<&Path>,
    config_path: Option<&Path>,
    selected_files: &[PathBuf],
    business: &str,
    period: &str,
) -> Vec<S5FileClassification> {
    let chosen_path: Option<PathBuf> = template_path
        .filter(|path| !path.as_os_str().is_empty())
        .map(Path::to_path_buf)
        .filter(|path| path.is_file());
    let profiles: Vec<(PathBuf, Vec<String>)> = if let Some(chosen) = &chosen_path {
        let sheet_names = read_sheet_names(chosen);
        vec![(chosen.clone(), sheet_names)]
    } else if let Some(directory) = template_dir.filter(|dir| dir.is_dir()) {
        let mut paths: Vec<PathBuf> = std::fs::read_dir(directory)
            .into_iter()
            .flatten()
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
            .collect();
        paths.sort();
        paths
            .into_iter()
            .map(|path| {
                let canonical = std::fs::canonicalize(&path).unwrap_or_else(|_| path.clone());
                let sheet_names = read_sheet_names(&path);
                (canonical, sheet_names)
            })
            .collect()
    } else {
        Vec::new()
    };

    let configuration: Option<Configuration> = config_path
        .filter(|path| path.is_file())
        .and_then(|path| load_configuration(path, false, "").ok());

    let selected: BTreeSet<String> = selected_files
        .iter()
        .map(|path| resolve_display(path))
        .collect();
    let mut rows = Vec::with_capacity(files.len());
    for path in files {
        let resolved = resolve_display(path);
        let source_sheet_set = sheet_name_set(path);
        let source_sheet_list: Vec<String> = source_sheet_set.iter().cloned().collect();
        let mut ranked: Vec<(f64, &PathBuf, &Vec<String>)> = profiles
            .iter()
            .map(|(profile_path, profile_sheets)| {
                (
                    sheet_score(profile_sheets, &source_sheet_list),
                    profile_path,
                    profile_sheets,
                )
            })
            .collect();
        // Python sorts on the score only with a stable sort, so equal scores
        // keep the catalog's ascending-path order.
        ranked.sort_by(|left, right| right.0.total_cmp(&left.0));

        let mut profile_path = String::new();
        let mut template_sheets: BTreeSet<String> = BTreeSet::new();
        let mut ambiguous = false;
        if let Some((score, candidate_path, candidate_sheets)) = ranked.first() {
            let score = *score;
            let gap = if ranked.len() > 1 {
                score - ranked[1].0
            } else {
                score
            };
            ambiguous = ranked.len() > 1 && gap < 0.05;
            let covered = all_source_sheets_in_template(candidate_sheets, &source_sheet_list);
            if chosen_path.is_some() || (!ambiguous && (score >= 0.90 || covered)) {
                profile_path = path_text(candidate_path);
                template_sheets = candidate_sheets.iter().cloned().collect();
            } else if !ambiguous && score > 0.0 {
                profile_path = path_text(candidate_path);
                template_sheets = candidate_sheets.iter().cloned().collect();
            } else {
                template_sheets = candidate_sheets.iter().cloned().collect();
            }
        }

        let mut expected: BTreeSet<String> = template_sheets
            .iter()
            .filter(|name| name.as_str() != "审核规则")
            .cloned()
            .collect();
        let source_data_sheets: BTreeSet<String> = source_sheet_set
            .iter()
            .filter(|name| !matches!(name.as_str(), "审核规则" | "兜底模板"))
            .cloned()
            .collect();
        expected.remove("兜底模板");
        let common: BTreeSet<String> = expected
            .intersection(&source_data_sheets)
            .cloned()
            .collect();
        let missing: Vec<String> = expected.difference(&source_data_sheets).cloned().collect();
        let extra: Vec<String> = source_data_sheets.difference(&expected).cloned().collect();
        let match_status = if source_sheet_set.is_empty() || common.is_empty() {
            "unmatched"
        } else if ambiguous || profile_path.is_empty() || !extra.is_empty() {
            "partial"
        } else {
            "matched"
        };

        let mut row_business = String::new();
        if selected.contains(&resolved) && !business.is_empty() {
            row_business = business.to_owned();
        } else if let Some(configuration) = &configuration {
            let single_file = vec![path.clone()];
            let template_for_business = if profile_path.is_empty() {
                Path::new("")
            } else {
                Path::new(&profile_path)
            };
            if let Ok(identified) =
                match_s5_business(configuration, &single_file, template_for_business)
            {
                row_business = identified;
            }
        }
        let mut row_period = s5_file_period_hint(path);
        if row_period.is_empty() && selected.contains(&resolved) {
            row_period = period.to_owned();
        }
        let mut detail = if !profile_path.is_empty() {
            format!("模板：{}", basename_of(&profile_path))
        } else if ranked.first().is_some_and(|(score, _, _)| *score > 0.0) {
            format!("候选模板：{}", basename_of(path_text(ranked[0].1).as_str()))
        } else {
            "尚未识别模板".to_owned()
        };
        if !source_data_sheets.is_empty() {
            detail.push_str(&format!(
                "；源工作表匹配 {}/{} 张",
                common.len(),
                source_data_sheets.len()
            ));
        }
        if !missing.is_empty() {
            detail.push_str(&format!(
                "；模板中其他工作表（源文件未含，不影响匹配）：{}",
                missing.join("、")
            ));
        }
        if !extra.is_empty() {
            detail.push_str(&format!("；多出：{}", extra.join("、")));
        }
        rows.push(S5FileClassification {
            name: path_name(path),
            template: if profile_path.is_empty() {
                "未识别".to_owned()
            } else {
                basename_of(&profile_path)
            },
            match_status: match_status.to_owned(),
            business: row_business,
            period: row_period,
            match_detail: detail,
            matched_sheets: common.len(),
            source_sheet_count: source_data_sheets.len(),
            template_sheet_count: expected.len(),
        });
    }
    rows
}

fn path_text(path: &Path) -> String {
    path.to_string_lossy().into_owned()
}

fn resolve_display(path: &Path) -> String {
    std::fs::canonicalize(path)
        .map(|resolved| path_text(&resolved))
        .unwrap_or_else(|_| path_text(path))
}

fn basename_of(text: &str) -> String {
    text.rsplit(['/', '\\']).next().unwrap_or(text).to_owned()
}
