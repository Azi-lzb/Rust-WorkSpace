//! Ordered substring replacement for filenames, migrated from pytools operation 4-8.

use anyhow::{anyhow, Result};
use excel_reader::{read_workbook, SheetDump};
use serde::Serialize;
use std::path::{Path, PathBuf};

const CONFIG_SHEET: &str = "重命名配置";
const SOURCE_HEADER: &str = "原文件名片段";
const TARGET_HEADER: &str = "新文件名片段";

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct LocalFilenameMapRow {
    pub source_path: String,
    pub original_name: String,
    pub new_name: String,
    pub status: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct LocalFilenameMapRunReport {
    pub ok: usize,
    pub skip: usize,
    pub rules: usize,
    pub rows: Vec<LocalFilenameMapRow>,
}

/// Execute operation 4-8 using ordered M:N filename fragment mappings.
///
/// The configuration is read through the shared Calamine reader. The function
/// accepts regular files of any extension, does not open or rewrite them, and
/// reserves rename targets without replacing existing files.
pub fn run_local_filename_map_48_candidate(
    config_path: impl AsRef<Path>,
    files: &[PathBuf],
) -> Result<LocalFilenameMapRunReport> {
    let mappings = load_mappings(config_path.as_ref())?;
    if mappings.is_empty() {
        return Err(anyhow!("重命名配置 缺少 M/N 列局部文件名映射。"));
    }

    let mut ok = 0;
    let mut skip = 0;
    let mut rows = Vec::with_capacity(files.len());
    for source in files {
        let mut row = LocalFilenameMapRow {
            source_path: source.to_string_lossy().into_owned(),
            original_name: source
                .file_name()
                .map(|name| name.to_string_lossy().into_owned())
                .unwrap_or_default(),
            new_name: String::new(),
            status: String::new(),
        };
        if !source.exists() || !source.is_file() {
            skip += 1;
            row.status = "跳过：非文件或文件不存在".to_owned();
            rows.push(row);
            continue;
        }

        let Some(file_name) = source.file_name().and_then(|name| name.to_str()) else {
            skip += 1;
            row.status = "跳过：文件名不是 UTF-8".to_owned();
            rows.push(row);
            continue;
        };
        let new_name = apply_mapping(file_name, &mappings);
        if new_name == row.original_name {
            skip += 1;
            row.status = "跳过：没有匹配片段".to_owned();
            rows.push(row);
            continue;
        }
        if !is_single_filename(&new_name) {
            skip += 1;
            row.status = "跳过：改名失败".to_owned();
            rows.push(row);
            continue;
        }

        match crate::file_sort_prefix::move_to_available_path(
            source,
            &source.with_file_name(&new_name),
        ) {
            Ok(target) => {
                ok += 1;
                row.new_name = target
                    .file_name()
                    .map(|name| name.to_string_lossy().into_owned())
                    .unwrap_or_default();
                row.status = "已按局部映射改名".to_owned();
            }
            Err(_) => {
                skip += 1;
                row.status = "跳过：改名失败".to_owned();
            }
        }
        rows.push(row);
    }

    Ok(LocalFilenameMapRunReport {
        ok,
        skip,
        rules: mappings.len(),
        rows,
    })
}

fn load_mappings(path: &Path) -> Result<Vec<(String, String)>> {
    let workbook = read_workbook(path)?;
    let sheet = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == CONFIG_SHEET)
        .ok_or_else(|| anyhow!("重命名配置 缺少 M/N 列局部文件名映射。"))?;

    let source_column = header_column(sheet, SOURCE_HEADER);
    let target_column = header_column(sheet, TARGET_HEADER);
    let (Some(source_column), Some(target_column)) = (source_column, target_column) else {
        return Err(anyhow!("重命名配置 缺少 M/N 列局部文件名映射。"));
    };

    let max_row = sheet
        .values
        .keys()
        .filter_map(|address| parse_cell_address(address).map(|(row, _)| row))
        .max()
        .unwrap_or(1);
    let mappings = (2..=max_row)
        .filter_map(|row| {
            let source = field_at(sheet, row, source_column);
            if source.is_empty() {
                return None;
            }
            Some((source, field_at(sheet, row, target_column)))
        })
        .collect();
    Ok(mappings)
}

fn header_column(sheet: &SheetDump, expected: &str) -> Option<usize> {
    sheet
        .values
        .iter()
        .filter_map(|(address, value)| {
            let (row, column) = parse_cell_address(address)?;
            (row == 1 && value_text(value) == expected).then_some(column)
        })
        .min()
}

fn field_at(sheet: &SheetDump, row: usize, column: usize) -> String {
    sheet
        .values
        .get(&cell_address(row, column))
        .map(value_text)
        .unwrap_or_default()
}

fn value_text(value: &serde_json::Value) -> String {
    let text = match value["type"].as_str().unwrap_or_default() {
        "string" | "datetime_iso" | "duration_iso" | "error" => {
            value["value"].as_str().unwrap_or_default().to_owned()
        }
        "integer" | "number" | "datetime_serial" => value["value"].to_string(),
        "boolean" => match value["value"].as_bool() {
            Some(true) => "True".to_owned(),
            Some(false) => "False".to_owned(),
            None => String::new(),
        },
        _ => String::new(),
    };
    text.trim().to_owned()
}

fn parse_cell_address(address: &str) -> Option<(usize, usize)> {
    let split = address.find(|character: char| character.is_ascii_digit())?;
    let (column, row) = address.split_at(split);
    let column = column.chars().try_fold(0_usize, |value, character| {
        let letter = character.to_ascii_uppercase();
        if !letter.is_ascii_uppercase() {
            return None;
        }
        value
            .checked_mul(26)?
            .checked_add((letter as u8 - b'A' + 1) as usize)
    })?;
    Some((row.parse().ok()?, column))
}

fn cell_address(row: usize, column: usize) -> String {
    let mut column = column;
    let mut label = String::new();
    while column > 0 {
        column -= 1;
        label.insert(0, (b'A' + (column % 26) as u8) as char);
        column /= 26;
    }
    format!("{label}{row}")
}

fn apply_mapping(name: &str, mappings: &[(String, String)]) -> String {
    let (stem, suffix) = stem_and_suffix(name);
    let mut new_stem = stem.to_owned();
    for (source, target) in mappings {
        new_stem = new_stem.replace(source, target);
    }
    format!("{new_stem}{suffix}")
}

fn stem_and_suffix(name: &str) -> (&str, &str) {
    match name.rfind('.') {
        Some(index) if index > 0 && index + 1 < name.len() => (&name[..index], &name[index..]),
        _ => (name, ""),
    }
}

fn is_single_filename(name: &str) -> bool {
    !name.is_empty() && name != "." && name != ".." && !name.contains('/') && !name.contains('\\')
}

#[cfg(test)]
mod tests {
    use super::apply_mapping;

    #[test]
    fn mappings_apply_in_order_to_stem_and_preserve_suffix() {
        let mappings = vec![
            ("aa".to_owned(), "bb".to_owned()),
            ("bb".to_owned(), "cc".to_owned()),
        ];
        assert_eq!(
            apply_mapping("file_aa_aa.docx", &mappings),
            "file_cc_cc.docx"
        );
        assert_eq!(apply_mapping("unmatched.pdf", &mappings), "unmatched.pdf");
    }

    #[test]
    fn filename_suffix_matches_python_pathlib_for_trailing_dot_and_dotfile() {
        let mappings = vec![("aa".to_owned(), "bb".to_owned())];
        assert_eq!(apply_mapping("file.aa.", &mappings), "file.bb.");
        assert_eq!(apply_mapping(".hidden_aa", &mappings), ".hidden_bb");
    }
}
