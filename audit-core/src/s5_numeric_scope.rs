//! Template-bound numeric-region projection for S5 source retrieval.
//!
//! Mirrors `numeric_scope.resolve_numeric_regions` and
//! `numeric_scope.project_numeric_regions`: named numeric rectangles are
//! projected from a selected layout template by matching the complete merged
//! header block. When no template context was supplied and no registered
//! layout matches, callers retain the legacy whole-sheet scan behavior.

use crate::s5_layout::{self, NamedRegion};
use crate::s5_scan::ScanBook;
use serde_json::{json, Value};
use sha2::{Digest, Sha256};
use std::{
    collections::{BTreeMap, BTreeSet},
    fs,
    path::{Path, PathBuf},
    time::Instant,
};

#[derive(Clone, Debug)]
pub(crate) struct NumericScopeError {
    pub code: String,
    pub message: String,
    pub evidence: Value,
}

impl NumericScopeError {
    fn new(code: &str, message: impl Into<String>, evidence: Value) -> Self {
        Self {
            code: code.to_owned(),
            message: message.into(),
            evidence,
        }
    }
}

#[derive(Clone, Debug)]
struct NumericRegionLayout {
    region: NamedRegion,
    top: usize,
    left: usize,
    height: usize,
    width: usize,
    pattern: Vec<(usize, usize, String)>,
    first_row_column: usize,
    last_row_column: usize,
    last_path: Vec<String>,
}

struct TemplateMetadata {
    sha256: String,
    numeric_sheets: BTreeSet<String>,
    layouts: Vec<NumericRegionLayout>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct SelectedNumericTemplate {
    pub path: PathBuf,
    pub sha256: String,
}

#[derive(Clone, Debug)]
struct TemplateSelection {
    path: Option<PathBuf>,
    elapsed_ms: u128,
}

/// Resolves template numeric rectangles once per source sheet and caches the
/// projected result for the lifetime of an S5 retrieval run.
pub(crate) struct NumericScopeResolver {
    explicit_template_path: Option<PathBuf>,
    template_dir: Option<PathBuf>,
    templates: BTreeMap<PathBuf, TemplateMetadata>,
    projected_by_sheet: BTreeMap<String, Vec<NamedRegion>>,
    selections_by_sheet: BTreeMap<String, TemplateSelection>,
    selected_templates: Vec<SelectedNumericTemplate>,
}

impl NumericScopeResolver {
    pub(crate) fn new(explicit_template_path: Option<&Path>, template_dir: Option<&Path>) -> Self {
        Self {
            explicit_template_path: explicit_template_path.map(Path::to_path_buf),
            template_dir: template_dir.map(Path::to_path_buf),
            templates: BTreeMap::new(),
            projected_by_sheet: BTreeMap::new(),
            selections_by_sheet: BTreeMap::new(),
            selected_templates: Vec::new(),
        }
    }

    pub(crate) fn regions_for_sheet(
        &mut self,
        source: &ScanBook,
        sheet_name: &str,
    ) -> Result<Vec<NamedRegion>, NumericScopeError> {
        if let Some(regions) = self.projected_by_sheet.get(sheet_name) {
            return Ok(regions.clone());
        }

        let started = Instant::now();
        let (regions, selected_path, selected_sha256) =
            if let Some(path) = self.explicit_template_path.clone() {
                let path = resolve_template_path(&path)?;
                let metadata = self.template_metadata(&path)?;
                let sha256 = metadata.sha256.clone();
                let projected = project_numeric_regions(source, metadata, sheet_name)?;
                if projected.is_empty() {
                    return Err(NumericScopeError::new(
                        "NUMERIC_REGION_NOT_FOUND",
                        "选定模板缺少该工作表可定位的数字区域",
                        json!({
                            "sheet": sheet_name,
                            "template": file_name(&path),
                        }),
                    ));
                }
                (projected, Some(path), Some(sha256))
            } else if let Some(template_dir) = self.template_dir.clone() {
                let selected = self.select_template(source, &template_dir, sheet_name)?;
                if let Some(path) = selected
                    .map(|path| resolve_template_path(&path))
                    .transpose()?
                {
                    let metadata = self.template_metadata(&path)?;
                    let sha256 = metadata.sha256.clone();
                    let projected = project_numeric_regions(source, metadata, sheet_name)?;
                    if projected.is_empty() {
                        return Err(NumericScopeError::new(
                            "NUMERIC_REGION_NOT_FOUND",
                            "布局模板缺少可定位的数字区域表头",
                            json!({
                                "sheet": sheet_name,
                                "template": file_name(&path),
                            }),
                        ));
                    }
                    (projected, Some(path), Some(sha256))
                } else {
                    (Vec::new(), None, None)
                }
            } else {
                (Vec::new(), None, None)
            };

        self.projected_by_sheet
            .insert(sheet_name.to_owned(), regions.clone());
        self.selections_by_sheet.insert(
            sheet_name.to_owned(),
            TemplateSelection {
                path: selected_path.clone(),
                elapsed_ms: started.elapsed().as_millis(),
            },
        );
        if let (Some(path), Some(sha256)) = (selected_path, selected_sha256) {
            if !self
                .selected_templates
                .iter()
                .any(|template| template.path == path)
            {
                self.selected_templates
                    .push(SelectedNumericTemplate { path, sha256 });
            }
        }
        Ok(regions)
    }

    pub(crate) fn selected_templates(&self) -> &[SelectedNumericTemplate] {
        &self.selected_templates
    }

    pub(crate) fn selection_records(&self, role: &str) -> Vec<Value> {
        self.selections_by_sheet
            .iter()
            .map(|(sheet, selection)| {
                json!({
                    "source_role": role,
                    "sheet": sheet,
                    "template": selection.path.as_deref().map(file_name).unwrap_or_else(|| "兼容回退".to_owned()),
                    "engine": if selection.path.is_some() {
                        "Calamine via excel-reader + namespace-aware defined names"
                    } else {
                        "whole-sheet compatibility fallback"
                    },
                    "elapsed_ms": selection.elapsed_ms,
                })
            })
            .collect()
    }

    fn template_metadata(&mut self, path: &Path) -> Result<&TemplateMetadata, NumericScopeError> {
        let key = path.to_path_buf();
        if !self.templates.contains_key(&key) {
            let metadata = load_template_metadata(path)?;
            self.templates.insert(key.clone(), metadata);
        }
        Ok(&self.templates[&key])
    }

    fn select_template(
        &mut self,
        source: &ScanBook,
        template_dir: &Path,
        sheet_name: &str,
    ) -> Result<Option<PathBuf>, NumericScopeError> {
        // Python's Path.glob returns no candidates when the default template
        // directory is absent; explicit CLI directories are validated before
        // the resolver is constructed.
        if !template_dir.is_dir() {
            return Ok(None);
        }
        let entries = fs::read_dir(template_dir).map_err(|error| {
            NumericScopeError::new(
                "RETRIEVAL_FAILED",
                format!("数字区域布局模板目录无法读取：{error}"),
                json!({ "template_dir": file_name(template_dir) }),
            )
        })?;
        let source_sheets: BTreeSet<&str> =
            source.sheet_names().iter().map(String::as_str).collect();
        let mut candidates = Vec::new();
        let mut paths = entries
            .map(|entry| {
                entry.map(|entry| entry.path()).map_err(|error| {
                    NumericScopeError::new(
                        "RETRIEVAL_FAILED",
                        format!("数字区域布局模板目录无法读取：{error}"),
                        json!({ "template_dir": file_name(template_dir) }),
                    )
                })
            })
            .collect::<Result<Vec<_>, _>>()?;
        paths.sort();
        for path in paths {
            if path
                .file_name()
                .and_then(|name| name.to_str())
                .is_some_and(|name| name.starts_with("~$"))
            {
                continue;
            }
            let extension = path
                .extension()
                .and_then(|value| value.to_str())
                .unwrap_or_default()
                .to_ascii_lowercase();
            if !extension.starts_with("xls") || !path.is_file() {
                continue;
            }
            let metadata = self.template_metadata(&path)?;
            if metadata.numeric_sheets.contains(sheet_name)
                && metadata
                    .numeric_sheets
                    .iter()
                    .all(|name| source_sheets.contains(name.as_str()))
            {
                candidates.push((metadata.numeric_sheets.len(), path));
            }
        }
        let Some(best_score) = candidates.iter().map(|(score, _)| *score).max() else {
            return Ok(None);
        };
        let best: Vec<PathBuf> = candidates
            .into_iter()
            .filter_map(|(score, path)| (score == best_score).then_some(path))
            .collect();
        if best.len() != 1 {
            return Err(NumericScopeError::new(
                "NUMERIC_REGION_AMBIGUOUS",
                "匹配到多个数字区域布局模板，请在源文件定义数字区域",
                json!({ "sheet": sheet_name }),
            ));
        }
        Ok(best.into_iter().next())
    }
}

fn resolve_template_path(path: &Path) -> Result<PathBuf, NumericScopeError> {
    let canonical = fs::canonicalize(path).map_err(|error| {
        NumericScopeError::new(
            "RETRIEVAL_FAILED",
            format!("数字区域布局模板无法读取：{}：{error}", file_name(path)),
            json!({ "template": file_name(path) }),
        )
    })?;
    #[cfg(windows)]
    {
        let canonical_text = canonical.to_string_lossy();
        if let Some(unc_path) = canonical_text.strip_prefix(r"\\?\UNC\") {
            return Ok(PathBuf::from(format!(r"\\{}", unc_path)));
        }
        if let Some(drive_path) = canonical_text.strip_prefix(r"\\?\") {
            return Ok(PathBuf::from(drive_path));
        }
    }
    Ok(canonical)
}

fn load_template_metadata(path: &Path) -> Result<TemplateMetadata, NumericScopeError> {
    let bytes = fs::read(path).map_err(|error| {
        NumericScopeError::new(
            "RETRIEVAL_FAILED",
            format!("数字区域布局模板无法读取：{}：{error}", file_name(path)),
            json!({ "template": file_name(path) }),
        )
    })?;
    let sha256 = format!("{:x}", Sha256::digest(bytes));
    let book = ScanBook::open(path).map_err(|error| {
        NumericScopeError::new(
            "RETRIEVAL_FAILED",
            format!("数字区域布局模板无法读取：{}：{error}", file_name(path)),
            json!({ "template": file_name(path) }),
        )
    })?;
    let (regions, _) = s5_layout::numeric_regions(path);
    let numeric_sheets: BTreeSet<String> =
        regions.iter().map(|region| region.sheet.clone()).collect();
    let structure = s5_layout::structure_regions(path);
    let axes: Vec<&NamedRegion> = structure.values().flat_map(|items| items.iter()).collect();
    let layouts = regions
        .into_iter()
        .filter_map(|region| numeric_region_layout(&book, &axes, region))
        .collect();
    Ok(TemplateMetadata {
        sha256,
        numeric_sheets,
        layouts,
    })
}

fn numeric_region_layout(
    book: &ScanBook,
    axes: &[&NamedRegion],
    region: NamedRegion,
) -> Option<NumericRegionLayout> {
    let (row_axis, column_axis) = numeric_header_regions(book, axes, &region).ok()?;
    let top = column_axis.min_row;
    let left = row_axis.min_col.min(column_axis.min_col);
    let height = region.min_row.checked_sub(top)?;
    let width = region.max_col.checked_sub(left)?.checked_add(1)?;
    if height.saturating_mul(width) > 10_000 {
        return None;
    }
    let mut pattern = Vec::with_capacity(height.saturating_mul(width));
    for row_offset in 0..height {
        for column_offset in 0..width {
            pattern.push((
                row_offset,
                column_offset,
                book.merged_label(&region.sheet, top + row_offset, left + column_offset),
            ));
        }
    }
    if !pattern.iter().any(|(_, _, label)| !label.is_empty()) {
        return None;
    }
    let last_path = (row_axis.min_col..=row_axis.max_col)
        .map(|column| book.merged_label(&region.sheet, region.max_row, column))
        .collect();
    Some(NumericRegionLayout {
        region,
        top,
        left,
        height,
        width,
        pattern,
        first_row_column: row_axis.min_col,
        last_row_column: row_axis.max_col,
        last_path,
    })
}

fn numeric_header_regions<'a>(
    book: &ScanBook,
    axes: &[&'a NamedRegion],
    data: &NamedRegion,
) -> Result<(NamedRegion, NamedRegion), ()> {
    let rows: Vec<&NamedRegion> = axes
        .iter()
        .copied()
        .filter(|region| {
            region.sheet == data.sheet
                && region.max_col < data.min_col
                && region.min_row <= data.min_row
                && region.max_row >= data.max_row
                && !region.name.starts_with("列索引区域")
        })
        .collect();
    let mut columns: Vec<&NamedRegion> = axes
        .iter()
        .copied()
        .filter(|region| {
            region.sheet == data.sheet
                && region.max_row < data.min_row
                && region.min_col <= data.min_col
                && region.max_col >= data.max_col
                && !region.name.starts_with("行索引区域")
        })
        .collect();
    let row_axis = choose_axis(&rows, "行索引区域", data, true)?;
    if columns.is_empty() {
        let expected_number = s5_layout::region_number(&row_axis.name).saturating_add(1);
        if let Some(candidate) = axes
            .iter()
            .copied()
            .filter(|region| {
                s5_layout::region_number(&region.name) == expected_number
                    && region.min_col <= data.min_col
                    && region.max_col >= data.max_col
                    && region.max_row < data.min_row
            })
            .min_by_key(|region| (region.min_row, region.min_col))
        {
            let has_values = (candidate.min_row..=candidate.max_row).any(|row| {
                (candidate.min_col..=candidate.max_col)
                    .any(|column| book.cell_has_stored_value(&data.sheet, row, column))
            });
            if has_values {
                columns.push(candidate);
            }
        }
    }
    let column_axis = choose_axis(&columns, "列索引区域", data, false)?;
    Ok(((*row_axis).clone(), (*column_axis).clone()))
}

fn choose_axis<'a>(
    candidates: &[&'a NamedRegion],
    explicit_prefix: &str,
    data: &NamedRegion,
    row_axis: bool,
) -> Result<&'a NamedRegion, ()> {
    let explicit: Vec<&NamedRegion> = candidates
        .iter()
        .copied()
        .filter(|region| region.name.starts_with(explicit_prefix))
        .collect();
    let selected = if explicit.is_empty() {
        candidates
    } else {
        explicit.as_slice()
    };
    let mut unique_bounds = BTreeMap::new();
    for region in selected {
        unique_bounds.insert(
            (
                region.min_row,
                region.min_col,
                region.max_row,
                region.max_col,
            ),
            *region,
        );
    }
    let distance = |region: &&NamedRegion| {
        if row_axis {
            data.min_col.saturating_sub(region.max_col)
        } else {
            data.min_row.saturating_sub(region.max_row)
        }
    };
    let Some(nearest) = unique_bounds.values().map(|region| distance(region)).min() else {
        return Err(());
    };
    let nearest: Vec<&NamedRegion> = unique_bounds
        .values()
        .copied()
        .filter(|region| distance(region) == nearest)
        .collect();
    if nearest.len() == 1 {
        Ok(nearest[0])
    } else {
        Err(())
    }
}

fn project_numeric_regions(
    source: &ScanBook,
    template: &TemplateMetadata,
    sheet_name: &str,
) -> Result<Vec<NamedRegion>, NumericScopeError> {
    let mut projected = Vec::new();
    for layout in template
        .layouts
        .iter()
        .filter(|layout| layout.region.sheet == sheet_name)
    {
        if !source.sheet_names().iter().any(|name| name == sheet_name) {
            continue;
        }
        let Some((anchor_row, anchor_column, anchor_label)) = layout
            .pattern
            .iter()
            .find(|(_, _, label)| !label.is_empty())
            .map(|(row, column, label)| (*row, *column, label.as_str()))
        else {
            continue;
        };
        let source_max_row = source.sheet_max_row(sheet_name);
        let source_max_column = source.sheet_max_col(sheet_name);
        if source_max_row < layout.height || source_max_column < layout.width {
            return Err(projection_failure(
                "NUMERIC_REGION_NOT_FOUND",
                "模板数字区域的完整表头无法唯一匹配源文件，请在源文件定义数字区域或更新布局模板",
                sheet_name,
                &layout.region.name,
            ));
        }
        let max_top = source_max_row - layout.height + 1;
        let max_left = source_max_column - layout.width + 1;
        let mut candidates = Vec::new();
        for row in 1..=max_top {
            for column in 1..=max_left {
                if source.merged_label(sheet_name, row + anchor_row, column + anchor_column)
                    != anchor_label
                {
                    continue;
                }
                if layout
                    .pattern
                    .iter()
                    .all(|(row_offset, column_offset, label)| {
                        source.merged_label(sheet_name, row + *row_offset, column + *column_offset)
                            == *label
                    })
                {
                    candidates.push((row, column));
                }
            }
        }
        if candidates.len() != 1 {
            let (code, message) = if candidates.is_empty() {
                (
                    "NUMERIC_REGION_NOT_FOUND",
                    "模板数字区域的完整表头无法唯一匹配源文件，请在源文件定义数字区域或更新布局模板",
                )
            } else {
                (
                    "NUMERIC_REGION_AMBIGUOUS",
                    "模板数字区域的完整表头无法唯一匹配源文件，请在源文件定义数字区域或更新布局模板",
                )
            };
            return Err(projection_failure(
                code,
                message,
                sheet_name,
                &layout.region.name,
            ));
        }
        let (matched_top, matched_left) = candidates[0];
        let row_delta = matched_top as isize - layout.top as isize;
        let column_delta = matched_left as isize - layout.left as isize;
        let min_row = shift_coordinate(layout.region.min_row, row_delta).ok_or_else(|| {
            projection_failure(
                "NUMERIC_REGION_NOT_FOUND",
                "模板数字区域投影坐标超出源工作表范围",
                sheet_name,
                &layout.region.name,
            )
        })?;
        let mut max_row = shift_coordinate(layout.region.max_row, row_delta).ok_or_else(|| {
            projection_failure(
                "NUMERIC_REGION_NOT_FOUND",
                "模板数字区域投影坐标超出源工作表范围",
                sheet_name,
                &layout.region.name,
            )
        })?;
        let min_column =
            shift_coordinate(layout.region.min_col, column_delta).ok_or_else(|| {
                projection_failure(
                    "NUMERIC_REGION_NOT_FOUND",
                    "模板数字区域投影坐标超出源工作表范围",
                    sheet_name,
                    &layout.region.name,
                )
            })?;
        let max_column =
            shift_coordinate(layout.region.max_col, column_delta).ok_or_else(|| {
                projection_failure(
                    "NUMERIC_REGION_NOT_FOUND",
                    "模板数字区域投影坐标超出源工作表范围",
                    sheet_name,
                    &layout.region.name,
                )
            })?;
        if layout.last_path.iter().any(|label| !label.is_empty()) {
            let first_column =
                shift_coordinate(layout.first_row_column, column_delta).ok_or_else(|| {
                    projection_failure(
                        "NUMERIC_REGION_NOT_FOUND",
                        "模板数字区域投影坐标超出源工作表范围",
                        sheet_name,
                        &layout.region.name,
                    )
                })?;
            let last_column =
                shift_coordinate(layout.last_row_column, column_delta).ok_or_else(|| {
                    projection_failure(
                        "NUMERIC_REGION_NOT_FOUND",
                        "模板数字区域投影坐标超出源工作表范围",
                        sheet_name,
                        &layout.region.name,
                    )
                })?;
            let endings: Vec<usize> = (min_row..=source.sheet_max_row(sheet_name))
                .filter(|row| {
                    (first_column..=last_column)
                        .map(|column| source.merged_label(sheet_name, *row, column))
                        .eq(layout.last_path.iter().cloned())
                })
                .collect();
            if endings.len() == 1 {
                max_row = endings[0];
            }
        }
        projected.push(NamedRegion {
            name: layout.region.name.clone(),
            sheet: sheet_name.to_owned(),
            reference: format!("{sheet_name}!R{min_row}C{min_column}:R{max_row}C{max_column}"),
            min_row,
            min_col: min_column,
            max_row,
            max_col: max_column,
        });
    }
    Ok(projected)
}

fn projection_failure(
    code: &str,
    message: &str,
    sheet_name: &str,
    region_name: &str,
) -> NumericScopeError {
    NumericScopeError::new(
        code,
        message,
        json!({ "sheet": sheet_name, "region": region_name }),
    )
}

fn shift_coordinate(value: usize, delta: isize) -> Option<usize> {
    if delta < 0 {
        value.checked_sub((-delta) as usize)
    } else {
        value.checked_add(delta as usize)
    }
}

fn file_name(path: &Path) -> String {
    path.file_name()
        .map(|name| name.to_string_lossy().into_owned())
        .unwrap_or_else(|| path.display().to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use sha2::{Digest, Sha256};

    fn file_sha256(path: &Path) -> String {
        let bytes = fs::read(path).expect("read sample for before/after integrity check");
        format!("{:x}", Sha256::digest(bytes))
    }

    #[test]
    #[ignore = "manual read-only sample projection; set AUDIT_S5_REAL_SOURCE and AUDIT_S5_REAL_TEMPLATE"]
    fn projects_numeric_regions_from_real_workbook_pair() {
        let source_path = PathBuf::from(
            std::env::var_os("AUDIT_S5_REAL_SOURCE")
                .expect("AUDIT_S5_REAL_SOURCE must point to a read-only XLSX sample"),
        );
        let template_path = PathBuf::from(
            std::env::var_os("AUDIT_S5_REAL_TEMPLATE")
                .expect("AUDIT_S5_REAL_TEMPLATE must point to a read-only XLSX template"),
        );
        let source_sha_before = file_sha256(&source_path);
        let template_sha_before = file_sha256(&template_path);
        let template_regions = s5_layout::numeric_regions(&template_path).0;
        let sheet_names: BTreeSet<String> = template_regions
            .iter()
            .map(|region| region.sheet.clone())
            .collect();
        let source = ScanBook::open(&source_path).expect("read real sample source");
        let mut resolver = NumericScopeResolver::new(Some(&template_path), None);
        let started = Instant::now();
        let mut projected = Vec::new();
        for sheet in sheet_names {
            projected.extend(
                resolver
                    .regions_for_sheet(&source, &sheet)
                    .expect("project real template regions"),
            );
        }
        assert_eq!(template_regions.len(), 4);
        assert_eq!(projected.len(), 4);
        assert_eq!(source_sha_before, file_sha256(&source_path));
        assert_eq!(template_sha_before, file_sha256(&template_path));

        let coordinates: Vec<Value> = projected
            .iter()
            .enumerate()
            .map(|(index, region)| {
                json!({
                    "region": index + 1,
                    "min_row": region.min_row,
                    "min_col": region.min_col,
                    "max_row": region.max_row,
                    "max_col": region.max_col,
                })
            })
            .collect();
        println!(
            "{}",
            json!({
                "template_numeric_region_count": template_regions.len(),
                "projected_region_count": projected.len(),
                "projected_coordinates": coordinates,
                "source_unchanged": true,
                "template_unchanged": true,
                "engine": "Calamine via excel-reader + namespace-aware defined names",
                "elapsed_ms": started.elapsed().as_millis(),
            })
        );
    }
}
