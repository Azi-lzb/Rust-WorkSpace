//! S5 template matching migrated from
//! `base_audit.systems.s5_configured_statistics.recognition.match_s5_template`.
//!
//! This module matches selected XLSX workbooks by worksheet-set compatibility.
//! It does not load S5 rules, infer a business, or run an audit.

use anyhow::{bail, ensure, Context, Result};
use excel_reader::read_xlsx_sheet_names_ooxml;
use std::{collections::BTreeSet, path::Path, path::PathBuf};

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct S5TemplateProfile {
    pub path: PathBuf,
    pub sheet_names: Vec<String>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct S5SourceWorkbook {
    pub path: PathBuf,
    pub sheet_names: Vec<String>,
}

/// Match the selected XLSX files to one template in a direct template folder.
/// The source Python index is a read/write cache; Rust reads the current
/// worksheet metadata and leaves that index and the source workbooks untouched.
pub fn match_s5_template(files: &[PathBuf], template_dir: &Path) -> Result<PathBuf> {
    ensure!(
        template_dir.is_dir(),
        "请选择存在的模板目录，或使用“更换”手动指定模板"
    );

    let mut paths = std::fs::read_dir(template_dir)
        .with_context(|| format!("列出 S5 模板目录：{}", template_dir.display()))?
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
        .collect::<Vec<_>>();
    paths.sort();
    ensure!(
        !paths.is_empty(),
        "模板目录中没有可用模板，请选择正确的模板目录或手动更换模板"
    );

    let profiles = paths
        .into_iter()
        .map(|path| {
            let canonical_path = std::fs::canonicalize(&path)
                .with_context(|| format!("解析 S5 模板路径：{}", path.display()))?;
            let sheet_names = read_xlsx_sheet_names_ooxml(&path)
                .with_context(|| format!("读取 S5 模板工作表名称：{}", path.display()))?;
            Ok(S5TemplateProfile {
                path: canonical_path,
                sheet_names,
            })
        })
        .collect::<Result<Vec<_>>>()?;

    let source_workbooks = files
        .iter()
        .map(|path| {
            let sheet_names = read_xlsx_sheet_names_ooxml(path)
                .with_context(|| format!("读取 S5 源文件工作表名称：{}", path.display()))?;
            Ok(S5SourceWorkbook {
                path: path.clone(),
                sheet_names,
            })
        })
        .collect::<Result<Vec<_>>>()?;

    match_s5_template_from_sheet_sets(&profiles, &source_workbooks)
}

/// Apply the source S5 per-file score, coverage, ambiguity, and batch
/// consistency rules to already-read worksheet names.
pub fn match_s5_template_from_sheet_sets(
    profiles: &[S5TemplateProfile],
    source_workbooks: &[S5SourceWorkbook],
) -> Result<PathBuf> {
    ensure!(
        !profiles.is_empty(),
        "模板目录中没有可用模板，请选择正确的模板目录或手动更换模板"
    );
    ensure!(!source_workbooks.is_empty(), "请至少选择一个 S5 源文件");

    let mut selected = Vec::with_capacity(source_workbooks.len());
    for source in source_workbooks {
        let mut ranked = profiles
            .iter()
            .map(|profile| {
                (
                    sheet_score(&profile.sheet_names, &source.sheet_names),
                    profile,
                )
            })
            .collect::<Vec<_>>();
        // Python sorts (score, path) in reverse order: ties choose the
        // lexicographically later resolved path before ambiguity is reported.
        ranked.sort_by(|left, right| {
            right
                .0
                .total_cmp(&left.0)
                .then_with(|| path_text(&right.1.path).cmp(&path_text(&left.1.path)))
        });

        let (top_score, top_profile) = &ranked[0];
        let covered = all_source_sheets_in_template(&top_profile.sheet_names, &source.sheet_names);
        if *top_score < 0.9 && !covered {
            bail!(
                "{}：没有匹配到模板，请核实模板目录或手动更换模板",
                file_name(&source.path)
            );
        }
        if ranked.len() > 1 && top_score - ranked[1].0 < 0.05 {
            bail!(
                "{}：有多个相近模板，请使用“更换”明确指定模板",
                file_name(&source.path)
            );
        }
        selected.push(top_profile.path.clone());
    }

    if selected.iter().collect::<BTreeSet<_>>().len() != 1 {
        bail!("勾选文件匹配到多种业务模板，请一次只保留一个业务的文件");
    }
    Ok(selected.remove(0))
}

pub(crate) fn sheet_score(template_sheets: &[String], source_sheets: &[String]) -> f64 {
    let expected = template_sheets
        .iter()
        .filter(|name| name.as_str() != "审核规则")
        .collect::<BTreeSet<_>>();
    let actual = source_sheets.iter().collect::<BTreeSet<_>>();
    if expected.is_empty() || actual.is_empty() {
        return 0.0;
    }
    let common = expected.intersection(&actual).count() as f64;
    let source_coverage = common / actual.len() as f64;
    let template_precision = common / expected.len() as f64;
    0.8 * source_coverage + 0.2 * template_precision
}

pub(crate) fn all_source_sheets_in_template(
    template_sheets: &[String],
    source_sheets: &[String],
) -> bool {
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

fn path_text(path: &Path) -> String {
    path.to_string_lossy().into_owned()
}

fn file_name(path: &Path) -> String {
    path.file_name()
        .map(|name| name.to_string_lossy().into_owned())
        .unwrap_or_else(|| path.display().to_string())
}

#[cfg(test)]
mod tests {
    use super::{match_s5_template_from_sheet_sets, S5SourceWorkbook, S5TemplateProfile};
    use std::path::PathBuf;

    fn profile(path: &str, sheets: &[&str]) -> S5TemplateProfile {
        S5TemplateProfile {
            path: PathBuf::from(path),
            sheet_names: sheets.iter().map(|sheet| (*sheet).to_owned()).collect(),
        }
    }

    fn source(path: &str, sheets: &[&str]) -> S5SourceWorkbook {
        S5SourceWorkbook {
            path: PathBuf::from(path),
            sheet_names: sheets.iter().map(|sheet| (*sheet).to_owned()).collect(),
        }
    }

    #[test]
    fn matches_covered_source_sheets_and_ignores_helper_sheets_for_coverage() {
        let profiles = [profile(
            "template.xlsx",
            &["汇总", "机构明细", "审核规则", "兜底模板"],
        )];
        let sources = [source("source.xlsx", &["汇总", "审核规则"])];
        assert_eq!(
            match_s5_template_from_sheet_sets(&profiles, &sources).unwrap(),
            PathBuf::from("template.xlsx")
        );
    }

    #[test]
    fn reports_near_tied_templates_as_ambiguous() {
        let profiles = [
            profile("a.xlsx", &["汇总", "机构明细"]),
            profile("b.xlsx", &["汇总", "机构明细"]),
        ];
        let sources = [source("source.xlsx", &["汇总", "机构明细"])];
        let error = match_s5_template_from_sheet_sets(&profiles, &sources).unwrap_err();
        assert!(format!("{error:#}").contains("有多个相近模板"));
    }

    #[test]
    fn rejects_unmatched_and_empty_source_sheets() {
        let profiles = [profile("template.xlsx", &["汇总"])];
        let sources = [source("source.xlsx", &["完全不同"])];
        let error = match_s5_template_from_sheet_sets(&profiles, &sources).unwrap_err();
        assert!(format!("{error:#}").contains("没有匹配到模板"));

        let empty = [source("empty.xlsx", &[])];
        let error = match_s5_template_from_sheet_sets(&profiles, &empty).unwrap_err();
        assert!(format!("{error:#}").contains("没有匹配到模板"));
    }

    #[test]
    fn rejects_a_batch_that_matches_different_templates() {
        let profiles = [
            profile("a.xlsx", &["汇总", "规则A"]),
            profile("b.xlsx", &["汇总", "规则B"]),
        ];
        let sources = [
            source("source_a.xlsx", &["汇总", "规则A"]),
            source("source_b.xlsx", &["汇总", "规则B"]),
        ];
        let error = match_s5_template_from_sheet_sets(&profiles, &sources).unwrap_err();
        assert!(format!("{error:#}").contains("匹配到多种业务模板"));
    }
}
