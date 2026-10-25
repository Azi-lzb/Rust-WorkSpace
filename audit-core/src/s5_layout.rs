//! S5 layout region discovery (foundation for address scanning and retrieval).
//!
//! Faithful Rust port of the pure defined-name half of
//! `s1audit_configured.layout`: `NamedRegion`, `region_number`,
//! `named_regions`, `numeric_regions`, `inferred_numeric_regions`,
//! `paired_column_region`, and the structure-region grouping used by the
//! source workbook (`sources.py` `_load_structure_regions`). The sequence
//! matching inside `scan_rule_paths` belongs to the address-scan slice and is
//! not ported here. Region references are composed exactly like the Python
//! `NamedRegion.reference` so real-template parity can compare them verbatim.

use excel_reader::{read_defined_names, read_workbook, DefinedNameScope, WorkbookDump};
use serde::Serialize;
use std::collections::BTreeMap;
use std::path::Path;

pub const STRUCTURE_PREFIXES: [&str; 3] = ["表结构区域", "行索引区域", "列索引区域"];
const INFERRED_UNBOUNDED: i64 = 1_000_000_000;

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct NamedRegion {
    pub name: String,
    pub sheet: String,
    pub reference: String,
    pub min_row: usize,
    pub min_col: usize,
    pub max_row: usize,
    pub max_col: usize,
}

/// Mirror of layout.region_number: trailing `_<digits>(#<digits>)?`.
pub fn region_number(name: &str) -> i64 {
    let without_hash = match name.rfind('#') {
        Some(position)
            if name[position + 1..]
                .bytes()
                .all(|byte| byte.is_ascii_digit()) =>
        {
            &name[..position]
        }
        _ => name,
    };
    let underscore = without_hash.rfind('_');
    let Some(underscore) = underscore else {
        return INFERRED_UNBOUNDED;
    };
    let digits = &without_hash[underscore + 1..];
    if !digits.is_empty() && digits.bytes().all(|byte| byte.is_ascii_digit()) {
        digits.parse::<i64>().unwrap_or(INFERRED_UNBOUNDED)
    } else {
        INFERRED_UNBOUNDED
    }
}

fn parse_bounds(reference: &str) -> Option<(usize, usize, usize, usize)> {
    let reference = reference.replace('$', "");
    let mut parts = reference.split(':');
    let (min_col, min_row) = parse_cell(parts.next()?)?;
    let (max_col, max_row) = match parts.next() {
        Some(last) => parse_cell(last)?,
        None => (min_col, min_row),
    };
    if parts.next().is_some() || min_row == 0 || min_col == 0 {
        return None;
    }
    Some((min_row, min_col, max_row, max_col))
}

fn parse_cell(address: &str) -> Option<(usize, usize)> {
    let mut row_part = String::new();
    let mut column = 0usize;
    for character in address.chars() {
        if character.is_ascii_alphabetic() {
            column = column * 26 + (character.to_ascii_uppercase() as usize - 'A' as usize + 1);
        } else if character.is_ascii_digit() {
            row_part.push(character);
        } else {
            return None;
        }
    }
    let row = row_part.parse::<usize>().ok()?;
    (row > 0 && column > 0).then_some((column, row))
}

/// Mirror of layout.named_regions over the workbook's defined names.
pub fn named_regions(path: &Path, prefixes: &[&str]) -> Vec<NamedRegion> {
    let records = read_defined_names(path).unwrap_or_default();
    let sheet_names = excel_reader::read_sheet_names(path).unwrap_or_default();
    let mut regions: BTreeMap<(String, String, usize, usize, usize, usize), NamedRegion> =
        BTreeMap::new();
    for record in &records {
        if !prefixes
            .iter()
            .any(|prefix| record.name.starts_with(prefix))
        {
            continue;
        }
        for (index, area) in record.areas.iter().enumerate() {
            let mut sheet = area.sheet_name.clone();
            if !sheet_names.iter().any(|name| name == &sheet)
                && sheet_names
                    .iter()
                    .any(|name| name == &sheet.replace("''", "'"))
            {
                sheet = sheet.replace("''", "'");
            }
            if !sheet_names.iter().any(|name| name == &sheet) {
                continue;
            }
            let Some((min_row, min_col, max_row, max_col)) = parse_bounds(&area.reference) else {
                continue;
            };
            let display_name = if record.areas.len() == 1 {
                record.name.clone()
            } else {
                format!("{}#{}", record.name, index + 1)
            };
            let key = (
                display_name.clone(),
                sheet.clone(),
                min_row,
                min_col,
                max_row,
                max_col,
            );
            regions.entry(key).or_insert_with(|| NamedRegion {
                name: display_name,
                reference: format!("{sheet}!{}", area.reference),
                sheet,
                min_row,
                min_col,
                max_row,
                max_col,
            });
        }
    }
    let mut list: Vec<NamedRegion> = regions.into_values().collect();
    list.sort_by(|left, right| {
        left.sheet
            .cmp(&right.sheet)
            .then(region_number(&left.name).cmp(&region_number(&right.name)))
            .then(left.min_row.cmp(&right.min_row))
            .then(left.min_col.cmp(&right.min_col))
    });
    list
}

/// Mirror of layout.inferred_numeric_regions' geometric row/column pairing.
pub(crate) fn paired_column_region<'a>(
    region: &NamedRegion,
    axes: &'a [NamedRegion],
) -> Option<&'a NamedRegion> {
    let mut candidates: Vec<&NamedRegion> = axes
        .iter()
        .filter(|item| {
            item.sheet == region.sheet
                && item.max_col > region.max_col
                && item.max_row < region.max_row
                && item.min_row <= region.min_row
                && !item.name.starts_with("行索引区域")
        })
        .collect();
    let numbered: Vec<&NamedRegion> = candidates
        .iter()
        .copied()
        .filter(|item| region_number(&item.name) == region_number(&region.name) + 1)
        .collect();
    if !numbered.is_empty() {
        candidates = numbered;
    }
    candidates.into_iter().min_by_key(|item| {
        (
            item.min_col,
            (item.min_row as i64 - region.min_row as i64).abs(),
        )
    })
}

fn cell_has_value(dump: &WorkbookDump, sheet: &str, row: usize, column: usize) -> bool {
    dump.sheets
        .iter()
        .find(|candidate| candidate.name == sheet)
        .map(|sheet| {
            let mut letters = String::new();
            let mut rest = column;
            while rest > 0 {
                let digit = ((rest - 1) % 26) as u8;
                letters.insert(0, (b'A' + digit) as char);
                rest = (rest - 1) / 26;
            }
            sheet.values.contains_key(&format!("{letters}{row}"))
        })
        .unwrap_or(false)
}

/// Mirror of layout.inferred_numeric_regions.
pub fn inferred_numeric_regions(path: &Path) -> (Vec<NamedRegion>, Vec<NamedRegion>) {
    let axes = named_regions(path, &STRUCTURE_PREFIXES);
    let dump = read_workbook(path).unwrap_or_else(|_| WorkbookDump {
        schema_version: 0,
        reader: "",
        elapsed_ms: 0,
        sheets: Vec::new(),
    });
    let mut results: BTreeMap<(String, usize, usize, usize, usize), NamedRegion> = BTreeMap::new();
    let mut result_order: Vec<(String, usize, usize, usize, usize)> = Vec::new();
    let mut used_bounds: std::collections::BTreeSet<(String, usize, usize, usize, usize)> =
        std::collections::BTreeSet::new();
    let bounds = |region: &NamedRegion| {
        (
            region.sheet.clone(),
            region.min_row,
            region.min_col,
            region.max_row,
            region.max_col,
        )
    };
    for row_axis in &axes {
        if row_axis.name.starts_with("列索引区域") {
            continue;
        }
        let mut column_axis = paired_column_region(row_axis, &axes).cloned();
        if column_axis.is_none() {
            let numbered: Vec<&NamedRegion> = axes
                .iter()
                .filter(|item| {
                    region_number(&item.name) == region_number(&row_axis.name) + 1
                        && item.min_col > row_axis.max_col
                })
                .collect();
            if let Some(candidate) = numbered
                .iter()
                .copied()
                .min_by_key(|item| (item.min_col, item.min_row))
            {
                let has_target_headers = (candidate.min_row..=candidate.max_row).any(|row| {
                    (candidate.min_col..=candidate.max_col)
                        .any(|column| cell_has_value(&dump, &row_axis.sheet, row, column))
                });
                if has_target_headers {
                    used_bounds.insert(bounds(&candidate));
                    column_axis = Some(NamedRegion {
                        name: candidate.name.clone(),
                        sheet: row_axis.sheet.clone(),
                        reference: format!(
                            "{}!{}{}:{}{}",
                            row_axis.sheet,
                            column_letter(candidate.min_col),
                            candidate.min_row,
                            column_letter(candidate.max_col),
                            candidate.max_row
                        ),
                        min_row: candidate.min_row,
                        min_col: candidate.min_col,
                        max_row: candidate.max_row,
                        max_col: candidate.max_col,
                    });
                }
            }
        }
        let Some(column_axis) = column_axis else {
            continue;
        };
        let min_row = row_axis.min_row.max(column_axis.max_row + 1);
        let min_col = (row_axis.max_col + 1).max(column_axis.min_col);
        if min_row > row_axis.max_row || min_col > column_axis.max_col {
            continue;
        }
        let key = (
            row_axis.sheet.clone(),
            min_row,
            min_col,
            row_axis.max_row,
            column_axis.max_col,
        );
        if !results.contains_key(&key) {
            let name = format!("数字区域_推导_{:03}", results.len() + 1);
            let region = NamedRegion {
                name: name.clone(),
                sheet: row_axis.sheet.clone(),
                reference: format!(
                    "{}!{}{}:{}{}",
                    row_axis.sheet,
                    column_letter(min_col),
                    min_row,
                    column_letter(column_axis.max_col),
                    row_axis.max_row
                ),
                min_row,
                min_col,
                max_row: row_axis.max_row,
                max_col: column_axis.max_col,
            };
            result_order.push(key.clone());
            results.insert(key, region);
        }
        used_bounds.insert(bounds(row_axis));
        used_bounds.insert(bounds(&column_axis));
    }
    let results: Vec<NamedRegion> = result_order
        .iter()
        .map(|key| results[key].clone())
        .collect();
    let missing: Vec<NamedRegion> = axes
        .iter()
        .filter(|region| !used_bounds.contains(&bounds(region)))
        .cloned()
        .collect();
    (results, missing)
}

fn column_letter(mut column: usize) -> String {
    let mut letters = Vec::new();
    while column > 0 {
        let digit = ((column - 1) % 26) as u8;
        letters.insert(0, (b'A' + digit) as char);
        column = (column - 1) / 26;
    }
    letters.into_iter().collect()
}

/// Mirror of layout.numeric_regions: per-sheet priority across declared and
/// inferred sources, with the origin list joined by 、.
pub fn numeric_regions(path: &Path) -> (Vec<NamedRegion>, String) {
    let primary = named_regions(path, &["数字区域", "数值区域"]);
    let compatible = named_regions(path, &["条件格式区域"]);
    let (inferred, _) = inferred_numeric_regions(path);
    let sheet_names = excel_reader::read_sheet_names(path).unwrap_or_default();
    let mut regions: Vec<NamedRegion> = Vec::new();
    let mut origins: Vec<String> = Vec::new();
    for sheet in &sheet_names {
        for (candidates, origin) in [
            (&primary, "数字区域/数值区域"),
            (&compatible, "条件格式区域（兼容现有模板）"),
            (&inferred, "表结构区域交叉推导"),
        ] {
            let selected: Vec<&NamedRegion> = candidates
                .iter()
                .filter(|region| &region.sheet == sheet)
                .collect();
            if !selected.is_empty() {
                regions.extend(selected.into_iter().cloned());
                if !origins.iter().any(|existing| existing == origin) {
                    origins.push(origin.to_owned());
                }
                break;
            }
        }
    }
    (regions, origins.join("、"))
}

/// Mirror of `sources._load_structure_regions`: structure/row-index/column-index
/// regions grouped by sheet and sorted by (region number, min column, min row,
/// name).
pub fn structure_regions(path: &Path) -> BTreeMap<String, Vec<NamedRegion>> {
    let mut grouped: BTreeMap<String, Vec<NamedRegion>> = BTreeMap::new();
    for region in named_regions(path, &STRUCTURE_PREFIXES) {
        grouped
            .entry(region.sheet.clone())
            .or_default()
            .push(region);
    }
    for regions in grouped.values_mut() {
        regions.sort_by(|left, right| {
            region_number(&left.name)
                .cmp(&region_number(&right.name))
                .then(left.min_col.cmp(&right.min_col))
                .then(left.min_row.cmp(&right.min_row))
                .then(left.name.cmp(&right.name))
        });
    }
    grouped
}

/// Sheet-count probe used by parity reporting only.
pub fn defined_name_count(path: &Path) -> usize {
    read_defined_names(path)
        .map(|records| records.len())
        .unwrap_or(0)
}

#[allow(dead_code)]
fn scope_marker(_: DefinedNameScope) {}
