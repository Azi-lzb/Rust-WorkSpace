//! Filename-based submission completeness report (conversion menu 4-8).

use anyhow::{Context, Result};
use chrono::Local;
use excel_reader::{read_workbook, SheetDump};
use excel_writer::{Format, Workbook};
use serde::Serialize;
use std::{
    collections::HashMap,
    fs,
    path::{Path, PathBuf},
};

#[derive(Clone, Debug, Serialize)]
pub struct SubmissionCheckReport {
    pub output: PathBuf,
    pub rules: usize,
    pub files: usize,
    pub submitted: usize,
    pub missing: usize,
    pub unmatched: usize,
    pub multi: usize,
}
#[derive(Clone)]
struct Rule {
    row: usize,
    enabled: String,
    date: String,
    area: String,
    institution: String,
    extra: String,
    remark: String,
    groups: Vec<Vec<String>>,
}

pub fn run_submission_check(
    config: &Path,
    inputs: &[PathBuf],
    output_dir: &Path,
) -> Result<SubmissionCheckReport> {
    let rules = load_rules(config)?;
    anyhow::ensure!(!rules.is_empty(), "提交检查配置 没有启用的检查项。");
    let files = collect_files(inputs);
    let mut hits = HashMap::<PathBuf, Vec<usize>>::new();
    for file in &files {
        hits.insert(file.clone(), Vec::new());
    }
    let mut rows = Vec::new();
    for (rule_index, rule) in rules.iter().enumerate() {
        let matched = files
            .iter()
            .filter(|path| {
                matches_rule(
                    path.file_name()
                        .unwrap_or_default()
                        .to_string_lossy()
                        .as_ref(),
                    rule,
                )
            })
            .collect::<Vec<_>>();
        for file in &matched {
            hits.get_mut(*file).expect("known file").push(rule_index);
        }
        rows.push(vec![
            rule.enabled.clone(),
            rule.date.clone(),
            rule.area.clone(),
            rule.institution.clone(),
            rule.extra.clone(),
            matched.len().to_string(),
            matched
                .iter()
                .map(|p| {
                    p.file_name()
                        .unwrap_or_default()
                        .to_string_lossy()
                        .into_owned()
                })
                .collect::<Vec<_>>()
                .join("；"),
            rule.remark.clone(),
            rule.row.to_string(),
        ]);
    }
    let missing_rules = rows
        .iter()
        .filter(|row| row[5] == "0")
        .cloned()
        .collect::<Vec<_>>();
    let unmatched = files
        .iter()
        .filter(|path| hits.get(*path).is_some_and(Vec::is_empty))
        .map(|path| {
            vec![
                path.file_name()
                    .unwrap_or_default()
                    .to_string_lossy()
                    .into_owned(),
                path.display().to_string(),
                "未命中任何启用配置".to_owned(),
            ]
        })
        .collect::<Vec<_>>();
    let multi = files
        .iter()
        .filter_map(|path| {
            let matched = hits.get(path)?;
            if matched.len() <= 1 {
                return None;
            }
            Some(vec![
                path.file_name()
                    .unwrap_or_default()
                    .to_string_lossy()
                    .into_owned(),
                path.display().to_string(),
                matched
                    .iter()
                    .map(|rule_index| {
                        let rule = &rules[*rule_index];
                        format!(
                            "第{}行({}|{}|{}|{})",
                            rule.row, rule.date, rule.area, rule.institution, rule.extra
                        )
                    })
                    .collect::<Vec<_>>()
                    .join("；"),
            ])
        })
        .collect::<Vec<_>>();
    fs::create_dir_all(output_dir)?;
    let output = unique_output(output_dir)?;
    let mut workbook = Workbook::new();
    write_table(
        &mut workbook,
        "检查结果",
        &[
            "是否启用",
            "日期关键词",
            "县区关键词",
            "机构关键词",
            "备用关键词",
            "命中次数",
            "命中的文件名",
            "备注",
            "配置行",
        ],
        &rows,
        &[5, 8],
    )?;
    write_table(
        &mut workbook,
        "缺失清单",
        &[
            "是否启用",
            "日期关键词",
            "县区关键词",
            "机构关键词",
            "备用关键词",
            "命中次数",
            "命中的文件名",
            "备注",
            "配置行",
        ],
        &missing_rules,
        &[5, 8],
    )?;
    write_table(
        &mut workbook,
        "未识别文件",
        &["文件名", "源文件路径", "备注"],
        &unmatched,
        &[],
    )?;
    write_table(
        &mut workbook,
        "多重命中文件",
        &["文件名", "源文件路径", "命中配置"],
        &multi,
        &[],
    )?;
    workbook
        .save(&output)
        .with_context(|| format!("save submission check {}", output.display()))?;
    let submitted = rows.iter().filter(|row| row[5] != "0").count();
    let missing = rows.len() - submitted;
    Ok(SubmissionCheckReport {
        output,
        rules: rules.len(),
        files: files.len(),
        submitted,
        missing,
        unmatched: unmatched.len(),
        multi: multi.len(),
    })
}

fn load_rules(path: &Path) -> Result<Vec<Rule>> {
    let dump = read_workbook(path)?;
    let sheet = dump
        .sheets
        .iter()
        .find(|sheet| sheet.name == "提交检查配置")
        .context("提交检查配置 缺少工作表")?;
    let headers = header_index(sheet);
    let mut rules = Vec::new();
    for row in 2..=used_rows(sheet) {
        let enabled = field(sheet, &headers, row, "是否启用");
        if !matches!(
            enabled.to_lowercase().as_str(),
            "1" | "y" | "yes" | "true" | "是" | "启用"
        ) {
            continue;
        }
        let date = field(sheet, &headers, row, "日期关键词");
        let area = field(sheet, &headers, row, "县区关键词");
        let institution = field(sheet, &headers, row, "机构关键词");
        let extra = field(sheet, &headers, row, "备用关键词");
        let groups = [&date, &area, &institution, &extra]
            .iter()
            .map(|v| split_tokens(v))
            .collect::<Vec<_>>();
        rules.push(Rule {
            row,
            enabled,
            date,
            area,
            institution,
            extra,
            remark: field(sheet, &headers, row, "备注"),
            groups,
        });
    }
    Ok(rules)
}
fn matches_rule(file_name: &str, rule: &Rule) -> bool {
    let name = file_name.to_lowercase();
    let mut has = false;
    for group in &rule.groups {
        if group.is_empty() {
            continue;
        }
        has = true;
        if !group
            .iter()
            .any(|token| name.contains(&token.to_lowercase()))
        {
            return false;
        }
    }
    has
}
fn split_tokens(value: &str) -> Vec<String> {
    value
        .replace('；', ";")
        .replace(',', ";")
        .split(';')
        .map(str::trim)
        .filter(|v| !v.is_empty())
        .map(str::to_owned)
        .collect()
}
fn collect_files(inputs: &[PathBuf]) -> Vec<PathBuf> {
    let mut out = Vec::new();
    let mut seen = std::collections::BTreeSet::new();
    for path in inputs {
        if path.is_dir() {
            walk(path, &mut out);
        } else if path.is_file() {
            out.push(path.clone());
        }
    }
    out.retain(|path| {
        !path
            .file_name()
            .unwrap_or_default()
            .to_string_lossy()
            .starts_with("~$")
    });
    out.retain(|path| seen.insert(path.canonicalize().unwrap_or_else(|_| path.clone())));
    out
}
fn walk(root: &Path, out: &mut Vec<PathBuf>) {
    let Ok(entries) = fs::read_dir(root) else {
        return;
    };
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            walk(&path, out);
        } else if path.is_file() {
            out.push(path);
        }
    }
}
fn write_table(
    workbook: &mut Workbook,
    name: &str,
    headers: &[&str],
    rows: &[Vec<String>],
    numeric_columns: &[usize],
) -> Result<()> {
    let sheet = workbook.add_worksheet();
    sheet.set_name(name)?;
    let header_format = Format::new().set_bold();
    for (col, header) in headers.iter().enumerate() {
        sheet.write_string_with_format(0, col as u16, *header, &header_format)?;
    }
    for (row, values) in rows.iter().enumerate() {
        for (col, value) in values.iter().enumerate() {
            if numeric_columns.contains(&col) {
                let number = value.parse::<f64>().with_context(|| {
                    format!("submission report column {} must be numeric", col + 1)
                })?;
                sheet.write_number((row + 1) as u32, col as u16, number)?;
            } else {
                sheet.write_string((row + 1) as u32, col as u16, value)?;
            }
        }
    }
    Ok(())
}
fn unique_output(dir: &Path) -> Result<PathBuf> {
    let stamp = Local::now().format("%Y%m%d_%H%M%S");
    for index in 0..10000 {
        let suffix = if index == 0 {
            String::new()
        } else {
            format!("_{index}")
        };
        let path = dir.join(format!("{stamp}_提交缺失检查{suffix}.xlsx"));
        if !path.exists() {
            return Ok(path);
        }
    }
    anyhow::bail!("cannot reserve submission report path")
}
fn header_index(sheet: &SheetDump) -> HashMap<String, usize> {
    let mut out = HashMap::new();
    for (address, value) in &sheet.values {
        if parse_a1(address).is_some_and(|(row, _)| row == 1) {
            out.insert(value_text(value), parse_a1(address).expect("checked").1);
        }
    }
    out
}
fn field(sheet: &SheetDump, headers: &HashMap<String, usize>, row: usize, name: &str) -> String {
    headers
        .get(name)
        .map(|col| {
            value_text(
                sheet
                    .values
                    .get(&format!("{}{}", column_letters(*col), row))
                    .unwrap_or(&serde_json::Value::Null),
            )
        })
        .unwrap_or_default()
}
fn value_text(value: &serde_json::Value) -> String {
    match value.get("type").and_then(serde_json::Value::as_str) {
        Some("string" | "datetime_iso" | "duration_iso" | "error") => value
            .get("value")
            .and_then(serde_json::Value::as_str)
            .unwrap_or_default()
            .trim()
            .into(),
        Some("number" | "integer") => value
            .get("value")
            .map(ToString::to_string)
            .unwrap_or_default(),
        Some("boolean") => value
            .get("value")
            .and_then(serde_json::Value::as_bool)
            .map(|v| if v { "True" } else { "False" }.into())
            .unwrap_or_default(),
        _ => String::new(),
    }
}
fn used_rows(sheet: &SheetDump) -> usize {
    sheet
        .values
        .keys()
        .filter_map(|address| parse_a1(address).map(|value| value.0))
        .max()
        .unwrap_or(1)
}
fn parse_a1(address: &str) -> Option<(usize, usize)> {
    let split = address.find(|c: char| c.is_ascii_digit())?;
    let col = address[..split].bytes().try_fold(0usize, |acc, b| {
        acc.checked_mul(26)?
            .checked_add((b.to_ascii_uppercase() - b'A' + 1) as usize)
    })?;
    Some((address[split..].parse().ok()?, col))
}
fn column_letters(mut col: usize) -> String {
    let mut value = String::new();
    while col > 0 {
        value.insert(0, (b'A' + ((col - 1) % 26) as u8) as char);
        col = (col - 1) / 26;
    }
    value
}
