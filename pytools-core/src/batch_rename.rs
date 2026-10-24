//! Rust implementation of PyCode's configuration-driven batch file rename (3-3).

use anyhow::{bail, Context, Result};
use excel_reader::{read_workbook, SheetDump};
use serde::Serialize;
use serde_json::Value;
use std::path::{Path, PathBuf};

const RENAME_SHEET: &str = "重命名配置";

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct BatchRenameRow {
    pub source_path: String,
    pub original_name: String,
    pub new_name: String,
    pub status: String,
    pub reason: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct BatchRenameResult {
    pub ok: usize,
    pub skip: usize,
    pub rows: Vec<BatchRenameRow>,
}

#[derive(Debug, Default)]
struct RenameMaps {
    short_to_full: Vec<(String, String)>,
    full_to_code: Vec<(String, String)>,
    settings: Vec<(String, String)>,
}

/// Rename paths using the original 3-3 configuration contract.
///
/// The configuration workbook is read through the shared Calamine adapter. Each
/// input path is renamed within its parent directory; directories are supported
/// because the source operation accepts both files and folders. Existing names
/// are never intentionally selected as destinations: collisions receive `_1`,
/// `_2`, ... before the original extension.
pub fn run_batch_rename(
    config_path: impl AsRef<Path>,
    paths: &[PathBuf],
) -> Result<BatchRenameResult> {
    let config_path = config_path.as_ref();
    let workbook = read_workbook(config_path)
        .with_context(|| format!("read rename configuration {}", config_path.display()))?;
    let sheet = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == RENAME_SHEET)
        .with_context(|| format!("configuration workbook has no {RENAME_SHEET:?} sheet"))?;
    let maps = load_rename_maps(sheet)?;

    let date_text = setting(&maps.settings, "数据日期");
    let report_name = setting(&maps.settings, "报表名称");
    if maps.short_to_full.is_empty()
        || maps.full_to_code.is_empty()
        || date_text.is_empty()
        || report_name.is_empty()
    {
        bail!("重命名配置 缺少必要配置：简称/全称、全称(代码映射)/代码、数据日期、报表名称。");
    }

    let mut result = BatchRenameResult {
        ok: 0,
        skip: 0,
        rows: Vec::with_capacity(paths.len()),
    };
    for path in paths {
        let original_name = path
            .file_name()
            .map(|name| name.to_string_lossy().into_owned())
            .unwrap_or_default();
        let source_path = path.to_string_lossy().into_owned();
        let row = rename_one(
            path,
            &original_name,
            &source_path,
            &maps,
            &date_text,
            &report_name,
        );
        if row.status == "renamed" {
            result.ok += 1;
        } else {
            result.skip += 1;
        }
        result.rows.push(row);
    }
    Ok(result)
}

fn rename_one(
    path: &Path,
    original_name: &str,
    source_path: &str,
    maps: &RenameMaps,
    date_text: &str,
    report_name: &str,
) -> BatchRenameRow {
    let skip = |reason: String| BatchRenameRow {
        source_path: source_path.to_owned(),
        original_name: original_name.to_owned(),
        new_name: String::new(),
        status: "skipped".to_owned(),
        reason,
    };

    if !path.exists() {
        return skip("not_exists".to_owned());
    }

    let lowered_name = original_name.to_lowercase();
    let Some((_, full_name)) = maps
        .short_to_full
        .iter()
        .find(|(short_name, _)| lowered_name.contains(&short_name.to_lowercase()))
    else {
        return skip("short_name_not_matched".to_owned());
    };
    let Some((_, code)) = maps.full_to_code.iter().find(|(full, _)| full == full_name) else {
        return skip(format!("full_name_no_code full={full_name}"));
    };

    let suffix = path
        .extension()
        .map(|extension| format!(".{}", extension.to_string_lossy()))
        .unwrap_or_default();
    let desired_name = format!("{date_text} {code} {full_name} {report_name}{suffix}");
    let desired_path = path.with_file_name(&desired_name);

    if same_resolved_path(&desired_path, path) {
        return skip("name_same".to_owned());
    }

    let target_path = next_available_path(&desired_path);
    match std::fs::rename(path, &target_path) {
        Ok(()) => BatchRenameRow {
            source_path: source_path.to_owned(),
            original_name: original_name.to_owned(),
            new_name: target_path
                .file_name()
                .map(|name| name.to_string_lossy().into_owned())
                .unwrap_or_default(),
            status: "renamed".to_owned(),
            reason: String::new(),
        },
        Err(error) => skip(format!("rename_failed err={error}")),
    }
}

fn next_available_path(path: &Path) -> PathBuf {
    if !path.exists() {
        return path.to_path_buf();
    }
    let parent = path.parent().unwrap_or_else(|| Path::new("."));
    let stem = path
        .file_stem()
        .map(|value| value.to_string_lossy().into_owned())
        .unwrap_or_default();
    let suffix = path
        .extension()
        .map(|value| format!(".{}", value.to_string_lossy()))
        .unwrap_or_default();
    for index in 1usize.. {
        let candidate = parent.join(format!("{stem}_{index}{suffix}"));
        if !candidate.exists() {
            return candidate;
        }
    }
    unreachable!("usize collision suffix range exhausted")
}

fn same_resolved_path(left: &Path, right: &Path) -> bool {
    match (std::fs::canonicalize(left), std::fs::canonicalize(right)) {
        (Ok(left), Ok(right)) => left == right,
        _ => left == right,
    }
}

fn load_rename_maps(sheet: &SheetDump) -> Result<RenameMaps> {
    let mut columns = std::collections::BTreeMap::<String, usize>::new();
    for (address, value) in &sheet.values {
        let Some((column, row)) = parse_cell_address(address) else {
            continue;
        };
        if row != 1 {
            continue;
        }
        let header = cell_raw_text(Some(value));
        if !header.is_empty() {
            // pandas keeps the first duplicate header's meaning and disambiguates
            // later duplicate labels; all required labels are looked up by name.
            columns.entry(header).or_insert(column);
        }
    }

    let max_row = sheet
        .values
        .keys()
        .filter_map(|address| parse_cell_address(address).map(|(_, row)| row))
        .max()
        .unwrap_or(1);
    let mut maps = RenameMaps::default();
    for row in 2..=max_row {
        let get = |header: &str| {
            columns
                .get(header)
                .and_then(|column| sheet.values.get(&cell_address(*column, row)))
                .map(|value| cell_text(Some(value)))
                .unwrap_or_default()
                .trim()
                .to_owned()
        };
        let short = get("简称");
        let full = get("全称");
        let code = get("代码");
        let code_full = get("全称(代码映射)");
        let key = get("键");
        let value = get("值");
        if !short.is_empty() && !full.is_empty() {
            ordered_set(&mut maps.short_to_full, short, full);
        }
        if !code_full.is_empty() && !code.is_empty() {
            ordered_set(&mut maps.full_to_code, code_full, code);
        }
        if !key.is_empty() && !value.is_empty() {
            ordered_set(&mut maps.settings, key, value);
        }
    }
    Ok(maps)
}

fn ordered_set(entries: &mut Vec<(String, String)>, key: String, value: String) {
    if let Some((_, current)) = entries.iter_mut().find(|(existing, _)| existing == &key) {
        *current = value;
    } else {
        entries.push((key, value));
    }
}

fn setting(settings: &[(String, String)], key: &str) -> String {
    settings
        .iter()
        .find(|(name, _)| name == key)
        .map(|(_, value)| value.clone())
        .unwrap_or_default()
}

fn cell_text(value: Option<&Value>) -> String {
    cell_raw_text(value).trim().to_owned()
}

fn cell_raw_text(value: Option<&Value>) -> String {
    let Some(value) = value else {
        return String::new();
    };
    let Some(kind) = value.get("type").and_then(Value::as_str) else {
        return value.as_str().unwrap_or_default().to_owned();
    };
    let Some(data) = value.get("value") else {
        return String::new();
    };
    match kind {
        "string" | "datetime_iso" | "duration_iso" | "error" => {
            data.as_str().unwrap_or_default().to_owned()
        }
        "boolean" => data
            .as_bool()
            .map(|value| if value { "True" } else { "False" }.to_owned())
            .unwrap_or_default(),
        "integer" => data
            .as_i64()
            .map(|number| number.to_string())
            .unwrap_or_default(),
        "number" | "datetime_serial" => data
            .as_f64()
            .filter(|number| number.is_finite())
            .map(|number| {
                if number.fract() == 0.0 {
                    format!("{number:.1}")
                } else {
                    number.to_string()
                }
            })
            .unwrap_or_default(),
        _ => String::new(),
    }
}

fn parse_cell_address(address: &str) -> Option<(usize, usize)> {
    let split = address.bytes().take_while(u8::is_ascii_alphabetic).count();
    if split == 0 || split == address.len() {
        return None;
    }
    let mut column = 0usize;
    for byte in address[..split].bytes() {
        column = column.checked_mul(26)? + usize::from(byte.to_ascii_uppercase() - b'A' + 1);
    }
    let row = address[split..].parse::<usize>().ok()?;
    (row > 0).then_some((column, row))
}

fn cell_address(mut column: usize, row: usize) -> String {
    let mut letters = String::new();
    while column > 0 {
        let remainder = (column - 1) % 26;
        letters.insert(0, char::from(b'A' + remainder as u8));
        column = (column - 1) / 26;
    }
    format!("{letters}{row}")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn ordered_mapping_updates_value_without_changing_match_priority() {
        let mut entries = Vec::new();
        ordered_set(&mut entries, "建".to_owned(), "第一版".to_owned());
        ordered_set(&mut entries, "建行".to_owned(), "第二项".to_owned());
        ordered_set(&mut entries, "建".to_owned(), "后续覆盖".to_owned());

        assert_eq!(
            entries,
            [
                ("建".to_owned(), "后续覆盖".to_owned()),
                ("建行".to_owned(), "第二项".to_owned())
            ]
        );
        assert_eq!(
            entries
                .iter()
                .find(|(short, _)| "客户_建行.xlsx"
                    .to_lowercase()
                    .contains(&short.to_lowercase()))
                .map(|(_, full)| full.as_str()),
            Some("后续覆盖")
        );
    }

    #[test]
    fn cell_coordinates_round_trip_for_multiletter_columns() {
        for column in [1, 26, 27, 52, 53, 702, 703] {
            let address = cell_address(column, 12);
            assert_eq!(parse_cell_address(&address), Some((column, 12)));
        }
    }
}

/// Compatibility alias for the original Python 3-3 operation identifier.
#[deprecated(note = "use `run_batch_rename`")]
pub fn run_batch_rename_33(
    config_path: impl AsRef<Path>,
    paths: &[PathBuf],
) -> Result<BatchRenameResult> {
    run_batch_rename(config_path, paths)
}
