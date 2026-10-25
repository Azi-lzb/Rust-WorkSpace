//! S5 template address scanning (`layout.scan_rule_paths` port).
//!
//! Port of the sequence-matching half of `s1audit_configured.layout`:
//! `_anchor`, `_sequence`, `row_path_candidates`, `_contextual_row_hits`,
//! `_starts`, `in_numeric_regions`, and the `scan_rule_paths` main loop. The
//! scanned book is read through the shared calamine reader; merged cells map
//! to their anchors exactly like the Python `_merged_by_row` bookkeeping.
//!
//! Data regions follow the unbound-source behavior of
//! `numeric_scope.resolve_numeric_regions` (no template binding and no
//! template directory → empty list → whole-sheet fallback); the template
//! projection machinery belongs to the source-reading slice.

use crate::s5_layout::{NamedRegion, STRUCTURE_PREFIXES};
use excel_reader::{read_workbook, WorkbookDump};
use serde::Serialize;
use std::collections::{BTreeMap, BTreeSet};
use std::path::{Path, PathBuf};

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct ScanPairLocation {
    pub row: usize,
    pub column: usize,
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct ScanTable {
    pub region: crate::s5_layout::NamedRegion,
    pub column_region: Option<crate::s5_layout::NamedRegion>,
    pub row_hits: BTreeMap<Vec<String>, Vec<(usize, usize, usize)>>,
    pub column_hits: BTreeMap<Vec<String>, Vec<usize>>,
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct ScanResult {
    /// `(row_path, column_path) -> sorted (row, column)` intersections.
    pub locations: BTreeMap<(Vec<String>, Vec<String>), Vec<ScanPairLocation>>,
    pub row_locations: BTreeMap<Vec<String>, Vec<usize>>,
    pub column_locations: BTreeMap<Vec<String>, Vec<usize>>,
    pub tables: Vec<ScanTable>,
}

pub struct ScanBook {
    path: PathBuf,
    dump: WorkbookDump,
    /// (sheet, row) -> merged spans (min_col, max_col, anchor_row, anchor_col).
    merged_by_row: BTreeMap<(String, usize), Vec<(usize, usize, usize, usize)>>,
    sheet_names: Vec<String>,
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

fn column_letter(mut column: usize) -> String {
    let mut letters = Vec::new();
    while column > 0 {
        let digit = ((column - 1) % 26) as u8;
        letters.insert(0, (b'A' + digit) as char);
        column = (column - 1) / 26;
    }
    letters.into_iter().collect()
}

fn address_of(row: usize, column: usize) -> String {
    format!("{}{row}", column_letter(column))
}

impl ScanBook {
    pub fn open(path: &Path) -> Result<Self, String> {
        let dump = read_workbook(path).map_err(|error| error.to_string())?;
        let sheet_names = dump.sheets.iter().map(|sheet| sheet.name.clone()).collect();
        let mut merged_by_row: BTreeMap<(String, usize), Vec<(usize, usize, usize, usize)>> =
            BTreeMap::new();
        for sheet in &dump.sheets {
            for merge in &sheet.merges {
                let Some((min_row, min_col, max_row, max_col)) = parse_bounds(merge) else {
                    continue;
                };
                for row in min_row..=max_row {
                    merged_by_row
                        .entry((sheet.name.clone(), row))
                        .or_default()
                        .push((min_col, max_col, min_row, min_col));
                }
            }
        }
        Ok(Self {
            path: path.to_path_buf(),
            dump,
            merged_by_row,
            sheet_names,
        })
    }

    pub fn sheet_names(&self) -> &[String] {
        &self.sheet_names
    }

    pub fn path(&self) -> &Path {
        &self.path
    }

    /// Whether the cell carries a formula (Python `data_type == "f"`).
    pub fn sheet_has_formula(&self, sheet: &str, row: usize, column: usize) -> bool {
        self.dump
            .sheets
            .iter()
            .find(|candidate| candidate.name == sheet)
            .map(|sheet_dump| sheet_dump.formulas.contains_key(&address_of(row, column)))
            .unwrap_or(false)
    }

    /// Stored (cached) cell content as `(type, text)`; numeric cells render
    /// through their shortest JSON text.
    pub fn stored_value(&self, sheet: &str, row: usize, column: usize) -> (String, String) {
        self.dump
            .sheets
            .iter()
            .find(|candidate| candidate.name == sheet)
            .and_then(|sheet_dump| sheet_dump.values.get(&address_of(row, column)))
            .map(|value| {
                let kind = value
                    .get("type")
                    .and_then(serde_json::Value::as_str)
                    .unwrap_or_default()
                    .to_owned();
                let text = match value.get("value") {
                    Some(serde_json::Value::String(text)) => text.trim().to_owned(),
                    Some(serde_json::Value::Number(number)) => number.to_string(),
                    Some(serde_json::Value::Bool(flag)) => {
                        if *flag {
                            "True".to_owned()
                        } else {
                            "False".to_owned()
                        }
                    }
                    _ => String::new(),
                };
                (kind, text)
            })
            .unwrap_or_else(|| ("empty".to_owned(), String::new()))
    }

    /// Cell count probe for diagnostics.
    pub fn sheet_value_count(&self, sheet: &str) -> usize {
        self.dump
            .sheets
            .iter()
            .find(|candidate| candidate.name == sheet)
            .map(|sheet_dump| sheet_dump.values.len())
            .unwrap_or(0)
    }

    pub fn sheet_max_row(&self, sheet: &str) -> usize {
        self.sheet_extent(sheet).0
    }

    pub fn sheet_max_col(&self, sheet: &str) -> usize {
        self.sheet_extent(sheet).1
    }

    fn sheet_extent(&self, sheet: &str) -> (usize, usize) {
        let Some(sheet_dump) = self
            .dump
            .sheets
            .iter()
            .find(|candidate| candidate.name == sheet)
        else {
            return (0, 0);
        };
        let mut max_row = 0usize;
        let mut max_col = 0usize;
        for address in sheet_dump.values.keys().chain(sheet_dump.formulas.keys()) {
            if let Some((column, row)) = parse_cell(address) {
                max_row = max_row.max(row);
                max_col = max_col.max(column);
            }
        }
        (max_row, max_col)
    }

    fn anchor(&self, sheet: &str, row: usize, column: usize) -> (usize, usize) {
        if let Some(spans) = self.merged_by_row.get(&(sheet.to_owned(), row)) {
            for (min_col, max_col, anchor_row, anchor_col) in spans {
                if *min_col <= column && column <= *max_col {
                    return (*anchor_row, *anchor_col);
                }
            }
        }
        (row, column)
    }

    /// Mirror of `cell_raw`: the merged non-anchor rule blanks the cell, the
    /// anchor keeps its stored text.
    fn cell_label(&self, sheet: &str, row: usize, column: usize) -> String {
        let Some(sheet_dump) = self
            .dump
            .sheets
            .iter()
            .find(|candidate| candidate.name == sheet)
        else {
            return String::new();
        };
        if let Some(spans) = self.merged_by_row.get(&(sheet.to_owned(), row)) {
            for (min_col, max_col, anchor_row, anchor_col) in spans {
                if *min_col <= column
                    && column <= *max_col
                    && (row, column) != (*anchor_row, *anchor_col)
                {
                    return String::new();
                }
            }
        }
        sheet_dump
            .values
            .get(&address_of(row, column))
            .map(|value| {
                value
                    .get("value")
                    .and_then(serde_json::Value::as_str)
                    .map(str::trim)
                    .unwrap_or_default()
                    .to_owned()
            })
            .unwrap_or_default()
    }

    /// Mirror of `_sequence` over one row's column span.
    pub fn row_sequence(
        &self,
        sheet: &str,
        row: usize,
        min_col: usize,
        max_col: usize,
    ) -> Vec<(String, usize)> {
        let mut result = Vec::new();
        let mut previous: Option<(usize, usize)> = None;
        for (offset, column) in (min_col..=max_col).enumerate() {
            let anchor = self.anchor(sheet, row, column);
            let label = self.cell_label(sheet, anchor.0, anchor.1);
            if previous != Some(anchor) && !label.is_empty() {
                result.push((label, offset));
            }
            previous = Some(anchor);
        }
        result
    }
}

#[derive(Clone, Debug, Default)]
pub struct RowPathEntry {
    /// `(label, column, anchor_row, anchor_col)` in column order.
    pub tokens: Vec<(String, usize, usize, usize)>,
    pub path: Vec<String>,
    pub aliases: Vec<Vec<String>>,
    pub primary: Vec<String>,
}

/// Mirror of `row_path_candidates` including the group-qualified alias
/// proposals for repeated horizontal paths.
pub fn row_path_candidates(
    book: &ScanBook,
    sheet: &str,
    region: &NamedRegion,
    candidate_rows: Option<&BTreeSet<usize>>,
) -> BTreeMap<usize, RowPathEntry> {
    let rows: Vec<usize> = match candidate_rows {
        Some(rows) => rows.iter().copied().collect(),
        None => (region.min_row..=region.max_row).collect(),
    };
    let mut entries: BTreeMap<usize, RowPathEntry> = BTreeMap::new();
    let mut anchor_occurrences: BTreeMap<(usize, String), BTreeSet<(usize, usize)>> =
        BTreeMap::new();
    let mut path_counts: BTreeMap<Vec<String>, usize> = BTreeMap::new();
    for row in &rows {
        let mut tokens: Vec<(String, usize, usize, usize)> = Vec::new();
        let mut previous: Option<(usize, usize)> = None;
        for column in region.min_col..=region.max_col {
            let anchor = book.anchor(sheet, *row, column);
            let raw = book.cell_label(sheet, *row, column);
            if previous != Some(anchor) && !raw.is_empty() {
                tokens.push((raw.clone(), column, anchor.0, anchor.1));
                anchor_occurrences
                    .entry((column, raw.clone()))
                    .or_default()
                    .insert(anchor);
            }
            previous = Some(anchor);
        }
        let path: Vec<String> = tokens
            .iter()
            .map(|(label, _, _, _)| label.clone())
            .collect();
        if !path.is_empty() {
            *path_counts.entry(path.clone()).or_default() += 1;
        }
        entries.insert(
            *row,
            RowPathEntry {
                tokens,
                path,
                aliases: Vec::new(),
                primary: Vec::new(),
            },
        );
    }

    let duplicate_paths: std::collections::BTreeSet<Vec<String>> = path_counts
        .iter()
        .filter(|(_, count)| **count > 1)
        .map(|(path, _)| path.clone())
        .collect();
    let mut duplicate_rows: Vec<usize> = entries
        .iter()
        .filter(|(_, item)| duplicate_paths.contains(&item.path))
        .map(|(row, _)| *row)
        .collect();
    duplicate_rows.sort();
    let mut runs: Vec<Vec<usize>> = Vec::new();
    for row in duplicate_rows {
        match runs.last_mut() {
            Some(run) if *run.last().expect("runs hold rows") + 1 == row => run.push(row),
            _ => runs.push(vec![row]),
        }
    }
    let mut run_groups: BTreeMap<Vec<String>, Vec<Vec<usize>>> = BTreeMap::new();
    for run in runs {
        let signature: Vec<String> = run
            .iter()
            .filter_map(|row| entries.get(row).map(|entry| entry.path.clone()))
            .next()
            .unwrap_or_default();
        run_groups.entry(signature).or_default().push(run);
    }

    let mut proposals: BTreeMap<usize, Vec<String>> = BTreeMap::new();
    for (signature, matching_runs) in &run_groups {
        if matching_runs.len() < 2 || signature.is_empty() {
            continue;
        }
        for run in matching_runs {
            let Some(first_entry) = entries.get(&run[0]) else {
                continue;
            };
            let repeated_tokens: Vec<&(String, usize, usize, usize)> = first_entry
                .tokens
                .iter()
                .filter(|(label, column, _, _)| {
                    anchor_occurrences
                        .get(&(*column, label.clone()))
                        .map(|occurrences| occurrences.len() > 1)
                        .unwrap_or(false)
                })
                .collect();
            if repeated_tokens.is_empty() {
                continue;
            }
            let repeated = repeated_tokens.last().expect("checked non-empty");
            let (label, column, _, _) = (repeated.0.clone(), repeated.1, repeated.2, repeated.3);
            let mut parent: Option<String> = None;
            let mut parent_row_cursor = run[0];
            while parent_row_cursor > region.min_row {
                parent_row_cursor -= 1;
                let Some(parent_entry) = entries.get(&parent_row_cursor) else {
                    continue;
                };
                let raw = book.cell_label(sheet, parent_row_cursor, column);
                let parent_label = raw.trim().to_owned();
                if parent_label.is_empty() {
                    continue;
                }
                let parent_anchor = book.anchor(sheet, parent_row_cursor, column);
                if anchor_occurrences
                    .get(&(column, parent_label.clone()))
                    .map(|occurrences| occurrences.len() != 1)
                    .unwrap_or(true)
                {
                    continue;
                }
                let parent_path = &parent_entry.path;
                if parent_path.is_empty() || path_counts.get(parent_path).copied().unwrap_or(0) != 1
                {
                    continue;
                }
                parent = Some(parent_label);
                let _ = parent_anchor;
                break;
            }
            let Some(parent) = parent else { continue };
            for row in run {
                let Some(entry) = entries.get_mut(row) else {
                    continue;
                };
                let token_index =
                    entry
                        .tokens
                        .iter()
                        .position(|(token_label, token_column, _, _)| {
                            *token_column == column
                                && anchor_occurrences
                                    .get(&(column, token_label.clone()))
                                    .map(|occurrences| occurrences.len() > 1)
                                    .unwrap_or(false)
                        });
                let Some(token_index) = token_index else {
                    continue;
                };
                let mut path = entry.path.clone();
                path.insert(token_index, parent.clone());
                proposals.insert(*row, path);
            }
        }
    }

    let mut proposal_counts: BTreeMap<Vec<String>, usize> = BTreeMap::new();
    for path in proposals.values() {
        *proposal_counts.entry(path.clone()).or_default() += 1;
    }
    let base_paths: std::collections::BTreeSet<Vec<String>> =
        entries.values().map(|entry| entry.path.clone()).collect();
    for (row, entry) in entries.iter_mut() {
        let base = entry.path.clone();
        let proposed = proposals.get(row).cloned();
        let unique = proposed
            .as_ref()
            .map(|path| {
                proposal_counts.get(path).copied().unwrap_or(0) == 1 && !base_paths.contains(path)
            })
            .unwrap_or(false);
        entry.aliases = match (&proposed, unique) {
            (Some(proposed), true) => vec![base.clone(), proposed.clone()],
            _ if !base.is_empty() => vec![base.clone()],
            _ => Vec::new(),
        };
        entry.primary = if unique {
            proposed.unwrap_or_else(|| base.clone())
        } else {
            base
        };
    }
    entries
}

fn starts(sequence: &[(String, usize)], path: &[String]) -> Vec<usize> {
    let mut starts = Vec::new();
    if path.len() > sequence.len() {
        return starts;
    }
    for index in 0..=(sequence.len() - path.len()) {
        if sequence[index..index + path.len()]
            .iter()
            .zip(path.iter())
            .all(|((value, _), expected)| value == expected)
        {
            starts.push(index);
        }
    }
    starts
}

/// Mirror of `_contextual_row_hits`.
pub fn contextual_row_hits(
    book: &ScanBook,
    sheet: &str,
    region: &NamedRegion,
    paths: &[Vec<String>],
    candidate_rows: &BTreeSet<usize>,
) -> BTreeMap<Vec<String>, Vec<(usize, usize, usize)>> {
    let rows: Vec<usize> = candidate_rows.iter().copied().collect();
    if rows.is_empty() {
        return BTreeMap::new();
    }
    let mut base_to_paths: BTreeMap<Vec<String>, Vec<(Vec<String>, usize, String)>> =
        BTreeMap::new();
    for path in paths {
        for parent_index in 0..path.len().saturating_sub(1) {
            let mut base = path.clone();
            base.remove(parent_index);
            if !base.is_empty() {
                base_to_paths.entry(base).or_default().push((
                    path.clone(),
                    parent_index,
                    path[parent_index].clone(),
                ));
            }
        }
    }
    if base_to_paths.is_empty() {
        return BTreeMap::new();
    }

    let mut sequences: BTreeMap<usize, Vec<(String, usize)>> = BTreeMap::new();
    let mut anchors_by_column_label: BTreeMap<(usize, String), BTreeSet<(usize, usize)>> =
        BTreeMap::new();
    for row in &rows {
        let sequence = book.row_sequence(sheet, *row, region.min_col, region.max_col);
        for column in region.min_col..=region.max_col {
            let anchor = book.anchor(sheet, *row, column);
            let label = book.cell_label(sheet, anchor.0, anchor.1);
            if !label.is_empty() {
                anchors_by_column_label
                    .entry((column, label))
                    .or_default()
                    .insert(anchor);
            }
        }
        sequences.insert(*row, sequence);
    }

    let mut hits_by_base: BTreeMap<Vec<String>, Vec<(usize, usize)>> = BTreeMap::new();
    for (row, sequence) in &sequences {
        for base in base_to_paths.keys() {
            for start in starts(sequence, base) {
                hits_by_base
                    .entry(base.clone())
                    .or_default()
                    .push((*row, start));
            }
        }
    }

    let mut result: BTreeMap<Vec<String>, BTreeSet<(usize, usize, usize)>> = BTreeMap::new();
    let eligible: std::collections::BTreeSet<usize> = rows.iter().copied().collect();
    for (base, candidates) in &hits_by_base {
        let distinct_rows: std::collections::BTreeSet<usize> =
            candidates.iter().map(|(row, _)| *row).collect();
        if distinct_rows.len() < 2 {
            continue;
        }
        for (path, parent_index, expected_parent) in &base_to_paths[base] {
            for (row, start) in candidates {
                let child_index = start + parent_index;
                let Some(sequence) = sequences.get(row) else {
                    continue;
                };
                if child_index >= sequence.len() {
                    continue;
                }
                let child_column = region.min_col + sequence[child_index].1;
                let mut nearest_unique = String::new();
                let mut parent_cursor = *row;
                while parent_cursor > region.min_row {
                    parent_cursor -= 1;
                    if !eligible.contains(&parent_cursor) {
                        continue;
                    }
                    let anchor = book.anchor(sheet, parent_cursor, child_column);
                    let label = book.cell_label(sheet, anchor.0, anchor.1);
                    if !label.is_empty()
                        && anchors_by_column_label
                            .get(&(child_column, label.clone()))
                            .map(|occurrences| occurrences.len() == 1)
                            .unwrap_or(false)
                    {
                        nearest_unique = label;
                        break;
                    }
                }
                if nearest_unique != *expected_parent {
                    continue;
                }
                let first_column = region.min_col + sequence[*start].1;
                let last_column = region.min_col + sequence[start + base.len() - 1].1;
                result
                    .entry(path.clone())
                    .or_default()
                    .insert((*row, first_column, last_column));
            }
        }
    }
    result
        .into_iter()
        .map(|(path, matches)| {
            let mut sorted: Vec<(usize, usize, usize)> = matches.into_iter().collect();
            sorted.sort();
            (path, sorted)
        })
        .collect()
}

/// Mirror of `in_numeric_regions` (anchor-aware membership).
pub fn in_numeric_regions(
    book: &ScanBook,
    sheet: &str,
    row: usize,
    column: usize,
    regions: &[NamedRegion],
) -> bool {
    let (anchor_row, anchor_column) = book.anchor(sheet, row, column);
    regions.iter().any(|region| {
        region.min_row <= row
            && row <= region.max_row
            && region.min_col <= column
            && column <= region.max_col
            && region.min_row <= anchor_row
            && anchor_row <= region.max_row
            && region.min_col <= anchor_column
            && anchor_column <= region.max_col
    })
}

fn projected_column_region(
    book: &ScanBook,
    path: &Path,
    region: &NamedRegion,
    regions: &[NamedRegion],
) -> Option<NamedRegion> {
    if let Some(paired) = crate::s5_layout::paired_column_region(region, regions) {
        return Some(paired.clone());
    }
    let axis_regions = crate::s5_layout::named_regions(path, &STRUCTURE_PREFIXES);
    let candidates: Vec<&NamedRegion> = axis_regions
        .iter()
        .filter(|item| {
            region_number_of(item) == region_number_of(region) + 1 && item.min_col > region.max_col
        })
        .collect();
    let candidate = candidates
        .iter()
        .copied()
        .min_by_key(|item| (item.min_col, item.min_row))?;
    let has_content = (candidate.min_row..=candidate.max_row).any(|row| {
        (candidate.min_col..=candidate.max_col)
            .any(|column| !book.cell_label(&region.sheet, row, column).is_empty())
    });
    if !has_content {
        return None;
    }
    Some(NamedRegion {
        name: candidate.name.clone(),
        sheet: region.sheet.clone(),
        reference: format!(
            "{}!{}{}:{}{}",
            region.sheet,
            column_letter(candidate.min_col),
            candidate.min_row,
            column_letter(candidate.max_col),
            candidate.max_row
        ),
        min_row: candidate.min_row,
        min_col: candidate.min_col,
        max_row: candidate.max_row,
        max_col: candidate.max_col,
    })
}

fn region_number_of(region: &NamedRegion) -> i64 {
    crate::s5_layout::region_number(&region.name)
}

/// Mirror of `scan_rule_paths` with unbound-source data regions (empty).
pub fn scan_rule_paths(
    book: &ScanBook,
    template_path: &Path,
    sheet_name: &str,
    key_pairs: &[(Vec<String>, Vec<String>)],
) -> ScanResult {
    let structure_regions = crate::s5_layout::structure_regions(template_path)
        .get(sheet_name)
        .cloned()
        .unwrap_or_default();
    let regions: Vec<NamedRegion> = if structure_regions.is_empty() {
        let (max_row, max_col) = (
            book.sheet_max_row(sheet_name),
            book.sheet_max_col(sheet_name),
        );
        vec![NamedRegion {
            name: "全工作表回退".to_owned(),
            sheet: sheet_name.to_owned(),
            reference: format!(
                "{sheet_name}!{}1:{}{max_row}",
                column_letter(1),
                column_letter(max_col.max(1))
            ),
            min_row: 1,
            min_col: 1,
            max_row: max_row.max(1),
            max_col: max_col.max(1),
        }]
    } else {
        structure_regions
    };
    // Unbound template books carry no 数字区域, so data_regions is empty and
    // every region row is eligible (whole-sheet compatibility fallback).
    let data_regions: Vec<NamedRegion> = Vec::new();

    let row_paths: std::collections::BTreeSet<Vec<String>> = key_pairs
        .iter()
        .map(|(row_path, _)| row_path.clone())
        .collect();
    let mut locations: BTreeMap<(Vec<String>, Vec<String>), BTreeSet<(usize, usize)>> =
        BTreeMap::new();
    let mut row_locations: BTreeMap<Vec<String>, BTreeSet<usize>> = BTreeMap::new();
    let mut column_locations: BTreeMap<Vec<String>, BTreeSet<usize>> = BTreeMap::new();
    let mut tables: Vec<ScanTable> = Vec::new();

    for region in &regions {
        if region.name.starts_with("列索引区域") {
            continue;
        }
        let eligible_rows: BTreeSet<usize> = if data_regions.is_empty() {
            (region.min_row..=region.max_row).collect()
        } else {
            data_regions
                .iter()
                .filter(|data| data.sheet == sheet_name)
                .flat_map(|data| data.min_row..=data.max_row)
                .filter(|row| (region.min_row..=region.max_row).contains(row))
                .collect()
        };
        let path_index = row_path_candidates(book, sheet_name, region, Some(&eligible_rows));
        let contextual_hits = contextual_row_hits(
            book,
            sheet_name,
            region,
            &row_paths.iter().cloned().collect::<Vec<_>>(),
            &eligible_rows,
        );
        let mut row_hits: BTreeMap<Vec<String>, Vec<(usize, usize, usize)>> = BTreeMap::new();
        for (row, item) in &path_index {
            let sequence = book.row_sequence(sheet_name, *row, region.min_col, region.max_col);
            if sequence.is_empty() {
                continue;
            }
            for path in &row_paths {
                if item.aliases.iter().any(|alias| alias == path) {
                    let first_column =
                        region.min_col + sequence.first().expect("nonempty sequence").1;
                    let last_column =
                        region.min_col + sequence.last().expect("nonempty sequence").1;
                    row_hits.entry(path.clone()).or_default().push((
                        *row,
                        first_column,
                        last_column,
                    ));
                    row_locations.entry(path.clone()).or_default().insert(*row);
                    continue;
                }
                for start in starts(&sequence, path) {
                    let first_column = region.min_col + sequence[start].1;
                    let last_column = region.min_col + sequence[start + path.len() - 1].1;
                    row_hits.entry(path.clone()).or_default().push((
                        *row,
                        first_column,
                        last_column,
                    ));
                    row_locations.entry(path.clone()).or_default().insert(*row);
                }
            }
        }
        for (path, matches) in &contextual_hits {
            for (row, first_column, last_column) in matches {
                row_hits
                    .entry(path.clone())
                    .or_default()
                    .push((*row, *first_column, *last_column));
                row_locations.entry(path.clone()).or_default().insert(*row);
            }
        }
        if row_hits.is_empty() {
            continue;
        }
        let column_region = projected_column_region(book, template_path, region, &regions);
        let search = column_region.clone().unwrap_or_else(|| region.clone());
        let mut column_hits: BTreeMap<Vec<String>, BTreeSet<usize>> = BTreeMap::new();
        let mut cache: BTreeMap<(usize, usize), Vec<(String, usize)>> = BTreeMap::new();
        for (row_path, column_path) in key_pairs {
            for (row, _, last_column) in row_hits.get(row_path).cloned().unwrap_or_default() {
                let first_column = if column_region.is_some() {
                    search.min_col.max(region.max_col + 1)
                } else {
                    search.min_col.max(last_column + 1)
                };
                let last_row = search.max_row.min(row.saturating_sub(1));
                for column in first_column..=search.max_col {
                    if !data_regions.is_empty()
                        && !in_numeric_regions(book, sheet_name, row, column, &data_regions)
                    {
                        continue;
                    }
                    let cache_key = (column, last_row);
                    if !cache.contains_key(&cache_key) {
                        let mut sequence: Vec<(String, usize)> = Vec::new();
                        let mut previous: Option<(usize, usize)> = None;
                        for (offset, scan_row) in (search.min_row..=last_row).enumerate() {
                            let anchor = book.anchor(sheet_name, scan_row, column);
                            let label = book.cell_label(sheet_name, anchor.0, anchor.1);
                            if previous != Some(anchor) && !label.is_empty() {
                                sequence.push((label, offset));
                            }
                            previous = Some(anchor);
                        }
                        cache.insert(cache_key, sequence);
                    }
                    let sequence = &cache[&cache_key];
                    if !starts(sequence, column_path).is_empty() {
                        column_hits
                            .entry(column_path.clone())
                            .or_default()
                            .insert(column);
                        column_locations
                            .entry(column_path.clone())
                            .or_default()
                            .insert(column);
                        locations
                            .entry((row_path.clone(), column_path.clone()))
                            .or_default()
                            .insert((row, column));
                    }
                }
            }
        }
        tables.push(ScanTable {
            region: region.clone(),
            column_region,
            row_hits: row_hits.clone(),
            column_hits: column_hits
                .into_iter()
                .map(|(path, columns)| (path, columns.into_iter().collect()))
                .collect(),
        });
    }

    ScanResult {
        locations: locations
            .into_iter()
            .map(|(key, set)| {
                let mut sorted: Vec<ScanPairLocation> = set
                    .into_iter()
                    .map(|(row, column)| ScanPairLocation { row, column })
                    .collect();
                sorted.sort_by(|left, right| {
                    left.row
                        .cmp(&right.row)
                        .then(left.column.cmp(&right.column))
                });
                (key, sorted)
            })
            .collect(),
        row_locations: row_locations
            .into_iter()
            .map(|(key, set)| (key, set.into_iter().collect()))
            .collect(),
        column_locations: column_locations
            .into_iter()
            .map(|(key, set)| (key, set.into_iter().collect()))
            .collect(),
        tables,
    }
}
