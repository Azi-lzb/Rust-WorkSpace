//! Pure selection and area-extraction rules for the worksheet extraction flow.
//!
//! File reading and workbook writing remain in their respective adapters; this
//! module keeps the source feature's keyword, range, stitching, and compaction
//! behavior independently testable.

use crate::config::ExtractTask;
use crate::edit_plan::apply_plan_with_capability_file;
use crate::file_commit::{replace_target_from_candidate, sha256_file, verify_planning_snapshot};
use crate::summary_ops::{
    excel_error_value, patch_existing_error_cells, PendingErrorCell, SummaryProgress,
    SummaryRunEvent,
};
use anyhow::{bail, ensure, Context, Result};
use chrono::Local;
use excel_editor::{
    inspect_xlsx_profile, EditOperation, ExcelEditPlan, PreservationRequirement, RangeStyleMode,
    WorkbookFormat,
};
use excel_reader::{
    read_value_columns, read_value_columns_with_iso_dates, read_workbook,
    read_xlsx_cell_number_formats, read_xlsx_worksheet_extents, uses_1904_date_system,
    CellNumberFormatRequest, CellValue, SheetValues,
};
use excel_writer::{ExcelDateTime, Format, Workbook};
use serde::Serialize;
use serde_json::{json, Value};
use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::path::{Path, PathBuf};
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};

pub type CellMatrix = Vec<Vec<Option<Value>>>;

const EXCEL_MAX_COLUMNS: usize = 16_384;

/// Statistics from a new-workbook worksheet extraction candidate.
#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize)]
pub struct ExtractSheetsAreaResult {
    pub files: usize,
    pub extracted_sheets: usize,
    pub output_files: usize,
    pub failed_files: usize,
    pub paths: Vec<String>,
}

struct SparseSheet {
    name: String,
    max_row: usize,
    max_column: usize,
    values: BTreeMap<(usize, usize), Value>,
}

struct ExpectedSheet {
    name: String,
    max_row: usize,
    max_column: usize,
    non_empty_cells: usize,
    full_sheet: bool,
}

struct PendingWorkbook {
    path: PathBuf,
    existing_sha256: Option<String>,
    workbook: Workbook,
    sheets: Vec<ExpectedSheet>,
    existing_sheet_names: Vec<String>,
    sheet_names: Vec<String>,
    import_operations: Vec<EditOperation>,
    error_cells: Vec<PendingErrorCell>,
}

struct StagedWorkbook {
    stage_path: PathBuf,
    output_path: PathBuf,
    existing_sha256: Option<String>,
}

struct ExtractSheetsProgressEmitter<'a> {
    started: Instant,
    sequence: usize,
    progress: SummaryProgress,
    callback: &'a mut dyn FnMut(SummaryRunEvent),
}

impl ExtractSheetsProgressEmitter<'_> {
    fn emit(
        &mut self,
        phase: &str,
        status: &str,
        user_message: impl Into<String>,
        file: Option<&Path>,
        stage_elapsed: Option<Duration>,
        components: &[&str],
        actual_engine: Option<&str>,
        error_type: Option<&str>,
    ) {
        self.sequence += 1;
        (self.callback)(SummaryRunEvent {
            sequence: self.sequence,
            timestamp: Local::now().to_rfc3339(),
            phase: phase.to_owned(),
            status: status.to_owned(),
            user_message: user_message.into(),
            progress: self.progress.clone(),
            file_name: file
                .and_then(Path::file_name)
                .map(|name| name.to_string_lossy().into_owned()),
            stage_elapsed_ms: stage_elapsed.map(|elapsed| elapsed.as_millis()),
            total_elapsed_ms: self.started.elapsed().as_millis(),
            components: components
                .iter()
                .map(|component| (*component).to_owned())
                .collect(),
            actual_engine: actual_engine.map(str::to_owned),
            error_type: error_type.map(str::to_owned),
        });
    }
}

const FULL_AREA_WORDS: &[&str] = &["", "整表", "全部", "全表"];

pub fn normalize_range_spec(text: &str) -> String {
    let mut normalized = text.replace([' ', '\t'], "");
    for separator in ['，', '、', '；', ';'] {
        normalized = normalized.replace(separator, ",");
    }
    normalized.replace('：', ":")
}

pub fn token_to_index(token: &str, is_row: bool) -> usize {
    let token = token.trim();
    if token.is_empty() {
        return 0;
    }
    if token.bytes().all(|byte| byte.is_ascii_digit()) {
        return token.parse().unwrap_or(0);
    }
    if is_row || !token.bytes().all(|byte| byte.is_ascii_alphabetic()) {
        return 0;
    }
    column_number(token)
}

/// Parse the legacy row/column index syntax retained by the source helpers.
pub fn parse_indexes(spec: &str, max_value: usize, is_row: bool) -> Vec<usize> {
    if spec.is_empty() || max_value == 0 {
        return (1..=max_value).collect();
    }

    let mut result = Vec::new();
    for segment in normalize_range_spec(spec).split(',') {
        let segment = segment.trim();
        if segment.is_empty() {
            continue;
        }
        if let Some((start, end)) = segment.split_once(':') {
            let mut start = token_to_index(start, is_row);
            let mut end = token_to_index(end, is_row);
            if start == 0 || end == 0 {
                continue;
            }
            if start > end {
                std::mem::swap(&mut start, &mut end);
            }
            let lower = start.max(1);
            let upper = end.min(max_value);
            if lower <= upper {
                result.extend(lower..=upper);
            }
        } else {
            let index = token_to_index(segment, is_row);
            if (1..=max_value).contains(&index) {
                result.push(index);
            }
        }
    }

    if result.is_empty() {
        return (1..=max_value).collect();
    }
    let mut seen = BTreeSet::new();
    result.retain(|value| seen.insert(*value));
    result
}

pub fn build_exact_name(keywords: &[String]) -> String {
    keywords
        .iter()
        .filter(|keyword| !keyword.trim().is_empty())
        .cloned()
        .collect::<Vec<_>>()
        .join("-")
}

pub fn sheet_matches_all(name: &str, keywords: &[String]) -> bool {
    let name = name.to_lowercase();
    keywords
        .iter()
        .map(|keyword| keyword.trim())
        .filter(|keyword| !keyword.is_empty())
        .all(|keyword| name.contains(&keyword.to_lowercase()))
}

/// Return sheet indices in source order, except an exact joined-name match is
/// promoted to the first result just like the source implementation.
pub fn find_matched_sheets(sheet_names: &[String], keywords: &[String]) -> Vec<usize> {
    let exact = build_exact_name(keywords);
    let exact_index = (!exact.is_empty())
        .then(|| sheet_names.iter().position(|name| name == &exact))
        .flatten();
    let mut matched = Vec::new();
    if let Some(index) = exact_index {
        matched.push(index);
    }
    for (index, name) in sheet_names.iter().enumerate() {
        if Some(index) != exact_index && sheet_matches_all(name, keywords) {
            matched.push(index);
        }
    }
    matched
}

pub fn split_keyword_groups(text: &str) -> Vec<Vec<String>> {
    text.replace('；', ";")
        .split(';')
        .map(|group| {
            group
                .replace('，', ",")
                .split([',', '|'])
                .map(str::trim)
                .filter(|option| !option.is_empty())
                .map(str::to_owned)
                .collect::<Vec<_>>()
        })
        .filter(|group| !group.is_empty())
        .collect()
}

pub fn name_matches_keyword_options(name: &str, keyword_text: &str) -> bool {
    let groups = split_keyword_groups(keyword_text);
    let name = name.to_lowercase();
    groups.iter().all(|group| {
        group
            .iter()
            .any(|option| name.contains(&option.to_lowercase()))
    })
}

pub fn workbook_matches(source_file_name: &str, task: &ExtractTask) -> bool {
    name_matches_keyword_options(source_file_name, &task.workbook_keywords)
}

pub fn name_matches_exclude(name: &str, exclude_keywords: &str) -> bool {
    !exclude_keywords.trim().is_empty() && name_matches_keyword_options(name, exclude_keywords)
}

pub fn find_matched_sheets_for_task(
    sheet_names: &[String],
    sheet_keywords: &str,
    exclude_keywords: &str,
) -> Vec<usize> {
    let mut seen = BTreeSet::new();
    sheet_names
        .iter()
        .enumerate()
        .filter_map(|(index, name)| {
            if !seen.insert(name) {
                return None;
            }
            if name_matches_exclude(name, exclude_keywords)
                || !name_matches_keyword_options(name, sheet_keywords)
            {
                None
            } else {
                Some(index)
            }
        })
        .collect()
}

pub fn sanitize_sheet_name(name: &str) -> String {
    name.chars()
        .map(|character| {
            if matches!(character, ':' | '\\' | '/' | '?' | '*' | '[' | ']') {
                '_'
            } else {
                character
            }
        })
        .take(31)
        .collect()
}

pub fn unique_sheet_name(used_names: &[String], base: &str, fallback_time: &str) -> String {
    let base = sanitize_sheet_name(base);
    if !used_names.iter().any(|name| name == &base) {
        return base;
    }
    let root = base.chars().take(28).collect::<String>();
    for index in 2..=100 {
        let candidate = format!("{root}_{index}");
        if !used_names.iter().any(|name| name == &candidate) {
            return candidate;
        }
    }
    format!("{root}_{fallback_time}")
}

pub fn is_full_area_spec(spec: &str) -> bool {
    FULL_AREA_WORDS.contains(&spec.trim())
}

pub fn split_area_specs(area_spec: &str) -> Vec<String> {
    let text = area_spec.trim();
    if is_full_area_spec(text) {
        return vec![String::new()];
    }
    let parts = text
        .replace('；', ";")
        .split(';')
        .map(str::trim)
        .filter(|part| !part.is_empty())
        .map(str::to_owned)
        .collect::<Vec<_>>();
    if parts.is_empty() {
        vec![String::new()]
    } else {
        parts
    }
}

/// Parse an A1 cell/range, whole column range, whole row range, or full-sheet
/// marker. Returned row/column indices are one-based to mirror the source.
pub fn parse_area_indexes(
    area_spec: &str,
    max_row: usize,
    max_col: usize,
) -> Option<(Vec<usize>, Vec<usize>)> {
    let area = area_spec.trim().replace([' ', '\t'], "").replace('：', ":");
    if is_full_area_spec(&area) {
        return Some(((1..=max_row).collect(), (1..=max_col).collect()));
    }
    if max_row == 0 || max_col == 0 {
        return Some((Vec::new(), Vec::new()));
    }

    if let Some((start, end)) = area.split_once(':') {
        if let (Some((start_row, start_col)), Some((end_row, end_col))) =
            (parse_cell(start), parse_cell(end))
        {
            let (row_start, row_end) = ordered(start_row, end_row);
            let (col_start, col_end) = ordered(start_col, end_col);
            return Some((
                bounded_range(row_start, row_end, max_row),
                bounded_range(col_start, col_end, max_col),
            ));
        }
        if start.bytes().all(|byte| byte.is_ascii_alphabetic())
            && end.bytes().all(|byte| byte.is_ascii_alphabetic())
            && !start.is_empty()
            && !end.is_empty()
        {
            let (col_start, col_end) = ordered(column_number(start), column_number(end));
            return Some((
                (1..=max_row).collect(),
                bounded_range(col_start, col_end, max_col),
            ));
        }
        if start.bytes().all(|byte| byte.is_ascii_digit())
            && end.bytes().all(|byte| byte.is_ascii_digit())
            && !start.is_empty()
            && !end.is_empty()
        {
            let (row_start, row_end) = ordered(start.parse().ok()?, end.parse().ok()?);
            return Some((
                bounded_range(row_start, row_end, max_row),
                (1..=max_col).collect(),
            ));
        }
        return None;
    }

    if let Some((row, column)) = parse_cell(&area) {
        if row <= max_row && column <= max_col {
            return Some((vec![row], vec![column]));
        }
        return Some((Vec::new(), Vec::new()));
    }
    None
}

pub fn extract_area_values(data: &CellMatrix, area_spec: &str) -> Option<CellMatrix> {
    if data.is_empty() {
        return Some(Vec::new());
    }
    let max_row = data.len();
    let max_col = data.iter().map(Vec::len).max().unwrap_or(0);
    if max_col == 0 {
        return Some(Vec::new());
    }
    let (rows, columns) = parse_area_indexes(area_spec, max_row, max_col)?;
    Some(
        rows.into_iter()
            .map(|row| {
                columns
                    .iter()
                    .map(|column| {
                        data.get(row - 1)
                            .and_then(|source| source.get(column - 1))
                            .cloned()
                            .unwrap_or(None)
                    })
                    .collect()
            })
            .collect(),
    )
}

pub fn stitch_area_values(area_values: &[CellMatrix], stitch_mode: &str) -> CellMatrix {
    if area_values.is_empty() {
        return Vec::new();
    }
    if stitch_mode.trim() == "纵向追加" {
        return area_values.iter().flatten().cloned().collect();
    }

    let max_rows = area_values.iter().map(Vec::len).max().unwrap_or(0);
    let widths = area_values
        .iter()
        .map(|values| values.iter().map(Vec::len).max().unwrap_or(0))
        .collect::<Vec<_>>();
    (0..max_rows)
        .map(|row_index| {
            let mut row = Vec::new();
            for (values, width) in area_values.iter().zip(&widths) {
                if let Some(source_row) = values.get(row_index) {
                    row.extend(source_row.iter().cloned());
                    row.extend(std::iter::repeat_n(
                        None,
                        width.saturating_sub(source_row.len()),
                    ));
                } else {
                    row.extend(std::iter::repeat_n(None, *width));
                }
            }
            row
        })
        .collect()
}

pub fn extract_configured_values(
    data: &CellMatrix,
    area_spec: &str,
    stitch_mode: &str,
) -> Option<CellMatrix> {
    let mut areas = Vec::new();
    for part in split_area_specs(area_spec) {
        areas.push(extract_area_values(data, &part)?);
    }
    if areas.len() == 1 {
        areas.into_iter().next()
    } else {
        Some(stitch_area_values(&areas, stitch_mode))
    }
}

pub fn compact(values: &CellMatrix) -> CellMatrix {
    let rows = values
        .iter()
        .filter(|row| row.iter().any(|value| !is_empty_cell(value.as_ref())))
        .collect::<Vec<_>>();
    if rows.is_empty() {
        return Vec::new();
    }
    let max_columns = rows.iter().map(|row| row.len()).max().unwrap_or(0);
    let keep_columns = (0..max_columns)
        .map(|column| {
            rows.iter().any(|row| {
                row.get(column)
                    .is_some_and(|value| !is_empty_cell(value.as_ref()))
            })
        })
        .collect::<Vec<_>>();
    rows.into_iter()
        .map(|row| {
            keep_columns
                .iter()
                .enumerate()
                .filter_map(|(column, keep)| keep.then(|| row.get(column).cloned().unwrap_or(None)))
                .collect()
        })
        .collect()
}

pub fn is_blank_spec(spec: &str) -> bool {
    spec.trim().is_empty()
}

pub fn is_full_sheet_task(task: &ExtractTask) -> bool {
    is_full_area_spec(&task.area_spec)
}

pub fn output_path_for_task(task: &ExtractTask, output_dir: &Path, default_name: &str) -> PathBuf {
    let folder = task.output_folder.trim();
    let base_dir = if folder.is_empty() {
        output_dir.to_path_buf()
    } else {
        let folder = Path::new(folder);
        if folder.is_absolute() {
            folder.to_path_buf()
        } else {
            output_dir.join(folder)
        }
    };
    let name = if task.output_name.trim().is_empty() {
        default_name.trim()
    } else {
        task.output_name.trim()
    };
    let mut path = PathBuf::from(name);
    match path_suffix(&path) {
        Some(extension) if extension.eq_ignore_ascii_case("xlsx") => {}
        Some(_) => {
            path.set_extension("xlsx");
        }
        None => {
            if let Some(file_name) = path
                .file_name()
                .map(|name| name.to_string_lossy().into_owned())
            {
                path.set_file_name(format!("{file_name}.xlsx"));
            } else {
                path.set_extension("xlsx");
            }
        }
    }
    if !path.is_absolute() {
        base_dir.join(path)
    } else {
        path
    }
}

/// Run extraction against XLSX outputs. Full-sheet tasks use an evidence-gated
/// `ImportRange` batch per output workbook; area-only tasks use the Rust
/// reader/writer path. Existing outputs are appended only when requested and
/// when an ExcelModifier capability matches the workbook and operation batch.
pub fn run_extract_sheets_area_candidate(
    source_workbooks: &[PathBuf],
    tasks: &[ExtractTask],
    output_dir: &Path,
    append_existing: bool,
    default_output_name: Option<&str>,
) -> Result<ExtractSheetsAreaResult> {
    run_extract_sheets_candidate(
        source_workbooks,
        tasks,
        output_dir,
        append_existing,
        default_output_name,
        None,
        None,
    )
}

/// Run ordinary worksheet extraction. Area-only tasks need only the shared
/// reader/writer path. Full-sheet tasks additionally require a matching,
/// evidence-backed ExcelModifier capability for XLSX `ImportRange`.
pub fn run_extract_sheets_candidate(
    source_workbooks: &[PathBuf],
    tasks: &[ExtractTask],
    output_dir: &Path,
    append_existing: bool,
    default_output_name: Option<&str>,
    capability_path: Option<&Path>,
    platform: Option<&str>,
) -> Result<ExtractSheetsAreaResult> {
    run_extract_sheets_candidate_with_progress(
        source_workbooks,
        tasks,
        output_dir,
        append_existing,
        default_output_name,
        capability_path,
        platform,
        |_| {},
    )
}

pub fn run_extract_sheets_candidate_with_progress<F>(
    source_workbooks: &[PathBuf],
    tasks: &[ExtractTask],
    output_dir: &Path,
    append_existing: bool,
    default_output_name: Option<&str>,
    capability_path: Option<&Path>,
    platform: Option<&str>,
    mut on_event: F,
) -> Result<ExtractSheetsAreaResult>
where
    F: FnMut(SummaryRunEvent),
{
    let mut emitter = ExtractSheetsProgressEmitter {
        started: Instant::now(),
        sequence: 0,
        progress: SummaryProgress {
            total_inputs: source_workbooks.len(),
            total_files: source_workbooks.len(),
            ..SummaryProgress::default()
        },
        callback: &mut on_event,
    };
    emitter.emit(
        "run",
        "started",
        format!(
            "开始工作表提取：{} 个源文件，{} 条启用任务",
            source_workbooks.len(),
            tasks.len()
        ),
        None,
        None,
        &["pytools-core"],
        None,
        None,
    );
    let mut result = ExtractSheetsAreaResult::default();
    if tasks.is_empty() {
        emitter.emit(
            "run",
            "succeeded",
            "没有启用的工作表提取任务",
            None,
            Some(emitter.started.elapsed()),
            &["pytools-core"],
            None,
            None,
        );
        return Ok(result);
    }
    ensure!(
        capability_path.is_some() == platform.is_some(),
        "--capabilities and --platform must be provided together"
    );
    if tasks.iter().any(is_full_sheet_task) && capability_path.is_none() {
        bail!(
            "full-sheet extraction requires an accepted XLSX ImportRange capability; provide --capabilities and --platform"
        );
    }

    for (index, source_path) in source_workbooks.iter().enumerate() {
        emitter.progress.current_file = Some(index + 1);
        if !source_path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| {
                extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm")
            })
        {
            let file_name = source_path
                .file_name()
                .map(|name| name.to_string_lossy().into_owned())
                .unwrap_or_else(|| source_path.display().to_string());
            let message = format!(
                "工作表区域提取仅接受 .xlsx/.xlsm；整表样式导入仍仅接受 .xlsx：{file_name}"
            );
            emitter.emit(
                "input_filter",
                "failed",
                message.clone(),
                Some(source_path),
                None,
                &["pytools-core"],
                None,
                Some("unsupported_input_format"),
            );
            bail!("{message}");
        }
    }

    let generated_name;
    let default_name = match default_output_name {
        Some(name) => name,
        None => {
            generated_name = format!("提取结果_{}", Local::now().format("%Y%m%d_%H%M%S"));
            &generated_name
        }
    };
    let output_dir_started = Instant::now();
    if let Err(error) = fs::create_dir_all(output_dir) {
        emitter.emit(
            "prepare_output",
            "failed",
            format!("无法准备工作表提取输出目录：{error}"),
            None,
            Some(output_dir_started.elapsed()),
            &["pytools-core", "filesystem"],
            None,
            Some("output_directory_error"),
        );
        return Err(error)
            .with_context(|| format!("create output directory {}", output_dir.display()));
    }
    emitter.emit(
        "prepare_output",
        "succeeded",
        "工作表提取输出目录已准备",
        None,
        Some(output_dir_started.elapsed()),
        &["pytools-core", "filesystem"],
        None,
        None,
    );

    let mut output_indexes = BTreeMap::<String, usize>::new();
    let mut outputs = Vec::<PendingWorkbook>::new();
    let fallback_time = Local::now().format("%H%M%S").to_string();

    for (source_index, source_path) in source_workbooks.iter().enumerate() {
        emitter.progress.current_file = Some(source_index + 1);
        let source_started = Instant::now();
        emitter.emit(
            "source_read",
            "started",
            format!(
                "正在读取源工作簿 [{}/{}]",
                source_index + 1,
                source_workbooks.len()
            ),
            Some(source_path),
            None,
            &["pytools-core", "excel-reader"],
            None,
            None,
        );
        let source_sheets = match read_value_columns_with_iso_dates(source_path, EXCEL_MAX_COLUMNS)
        {
            Ok(sheets) => {
                result.files += 1;
                emitter.emit(
                    "source_read",
                    "succeeded",
                    "源工作簿读取完成",
                    Some(source_path),
                    Some(source_started.elapsed()),
                    &["pytools-core", "excel-reader", "calamine"],
                    Some("calamine"),
                    None,
                );
                sheets
            }
            Err(error) => {
                result.failed_files += 1;
                emitter.progress.processed_files = source_index + 1;
                emitter.progress.completed_files = source_index + 1;
                emitter.progress.failed_files = result.failed_files;
                emitter.emit(
                    "source_read",
                    "failed",
                    format!("源工作簿读取失败：{error}"),
                    Some(source_path),
                    Some(source_started.elapsed()),
                    &["pytools-core", "excel-reader", "calamine"],
                    Some("calamine"),
                    Some("source_read_error"),
                );
                continue;
            }
        };

        let source_file_name = source_path
            .file_name()
            .map(|name| name.to_string_lossy().into_owned())
            .unwrap_or_default();
        let sheet_names = source_sheets
            .iter()
            .map(|sheet| sheet.name.clone())
            .collect::<Vec<_>>();
        let subday_date_cells = read_subday_date_cells(source_path, &source_sheets)?;
        let mut sparse_sheets = source_sheets
            .iter()
            .map(|sheet| sparse_sheet(sheet, subday_date_cells.get(&sheet.name)))
            .collect::<Result<Vec<_>>>()?;
        let matching_area_sheet_indexes = tasks
            .iter()
            .filter(|task| {
                !is_full_sheet_task(task)
                    && workbook_matches(&source_file_name, task)
                    && !name_matches_exclude(&source_file_name, &task.exclude_keywords)
            })
            .flat_map(|task| {
                find_matched_sheets_for_task(
                    &sheet_names,
                    &task.sheet_keywords,
                    &task.exclude_keywords,
                )
            })
            .collect::<BTreeSet<_>>();
        if !matching_area_sheet_indexes.is_empty() {
            let merge_compatibility_started = Instant::now();
            let merged_child_values_suppressed = suppress_openpyxl_merged_child_values(
                source_path,
                &mut sparse_sheets,
                &matching_area_sheet_indexes,
            )?;
            emitter.emit(
                "merge_value_compatibility",
                "succeeded",
                format!(
                    "已按旧 Python 合并单元格读取语义忽略 {merged_child_values_suppressed} 个合并子单元格值"
                ),
                Some(source_path),
                Some(merge_compatibility_started.elapsed()),
                &["pytools-core", "excel-reader", "namespace-aware OOXML"],
                Some("Calamine + namespace-aware OOXML"),
                None,
            );
        }
        let extracted_sheets_before = result.extracted_sheets;
        let has_matching_full_sheet_task = tasks.iter().any(|task| {
            is_full_sheet_task(task)
                && workbook_matches(&source_file_name, task)
                && !name_matches_exclude(&source_file_name, &task.exclude_keywords)
        });
        if has_matching_full_sheet_task
            && source_path
                .extension()
                .and_then(|extension| extension.to_str())
                .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
        {
            let extents = read_xlsx_worksheet_extents(source_path).with_context(|| {
                format!("read stored-cell extents from {}", source_path.display())
            })?;
            emitter.emit(
                "worksheet_extent_read",
                "succeeded",
                "已读取整表导入所需的 OOXML 工作表范围",
                Some(source_path),
                Some(source_started.elapsed()),
                &["pytools-core", "excel-reader"],
                Some("excel-reader OOXML extent inventory"),
                None,
            );
            ensure!(
                extents.len() == sparse_sheets.len(),
                "OOXML extent inventory disagrees with Calamine sheet count in {}",
                source_path.display()
            );
            for (sheet, extent) in sparse_sheets.iter_mut().zip(extents) {
                ensure!(
                    sheet.name == extent.name,
                    "OOXML extent inventory disagrees with Calamine sheet order in {}",
                    source_path.display()
                );
                sheet.max_row = sheet.max_row.max(extent.max_row);
                sheet.max_column = sheet.max_column.max(extent.max_column);
            }
        }

        for task in tasks {
            if !workbook_matches(&source_file_name, task)
                || name_matches_exclude(&source_file_name, &task.exclude_keywords)
            {
                continue;
            }
            let matched_indexes = find_matched_sheets_for_task(
                &sheet_names,
                &task.sheet_keywords,
                &task.exclude_keywords,
            );
            for sheet_index in matched_indexes {
                let source_sheet = &sparse_sheets[sheet_index];
                let full_sheet = is_full_sheet_task(task);
                let values = if full_sheet {
                    None
                } else {
                    let Some(values) = extract_sparse_configured_values(
                        source_sheet,
                        &task.area_spec,
                        &task.stitch_mode,
                    ) else {
                        continue;
                    };
                    let values = compact(&values);
                    if values.is_empty() {
                        continue;
                    }
                    Some(values)
                };
                if full_sheet {
                    ensure!(
                        source_path
                            .extension()
                            .and_then(|extension| extension.to_str())
                            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx")),
                        "full-sheet extraction requires XLSX input for style import; convert {} to .xlsx first",
                        source_path.display()
                    );
                }

                let output_path = output_path_for_task(task, output_dir, default_name);
                if normalized_path_key(&output_path)? == normalized_path_key(source_path)? {
                    bail!(
                        "refuse to replace source workbook with extraction output: {}",
                        source_path.display()
                    );
                }
                let key = normalized_path_key(&output_path)?;
                let output_index = match output_indexes.get(&key).copied() {
                    Some(index) => index,
                    None => {
                        let append_to_existing = append_existing && output_path.exists();
                        let (existing_sha256, existing_sheet_names) = if append_to_existing {
                            ensure!(
                                    capability_path.is_some() && platform.is_some(),
                                    "appending to an existing extraction output requires --capabilities and --platform"
                                );
                            let expected_sha256 = sha256_file(&output_path).with_context(|| {
                                format!(
                                    "snapshot existing extraction output {}",
                                    output_path.display()
                                )
                            })?;
                            let profile =
                                inspect_xlsx_profile(&output_path).with_context(|| {
                                    format!(
                                        "inventory existing extraction output {}",
                                        output_path.display()
                                    )
                                })?;
                            ensure!(
                                profile.format == WorkbookFormat::Xlsx
                                    && profile.inventory_complete,
                                "existing extraction output has an incomplete XLSX inventory: {}",
                                output_path.display()
                            );
                            let sheets =
                                read_value_columns_with_iso_dates(&output_path, EXCEL_MAX_COLUMNS)
                                    .with_context(|| {
                                        format!(
                                            "read existing extraction output {}",
                                            output_path.display()
                                        )
                                    })?;
                            verify_planning_snapshot(&output_path, Some(expected_sha256.as_str()))?;
                            (
                                Some(expected_sha256),
                                sheets.into_iter().map(|sheet| sheet.name).collect(),
                            )
                        } else {
                            (None, Vec::new())
                        };
                        let mut sheet_names = existing_sheet_names.clone();
                        if !append_to_existing {
                            // The PyCode path starts a fresh openpyxl Workbook, whose
                            // default "Sheet" reserves that name until imported sheets
                            // have been added and the empty default page is removed.
                            sheet_names.push("Sheet".to_owned());
                        }
                        let index = outputs.len();
                        output_indexes.insert(key, index);
                        outputs.push(PendingWorkbook {
                            path: output_path,
                            existing_sha256,
                            workbook: Workbook::new(),
                            sheets: Vec::new(),
                            sheet_names,
                            existing_sheet_names,
                            import_operations: Vec::new(),
                            error_cells: Vec::new(),
                        });
                        index
                    }
                };
                let pending = &mut outputs[output_index];
                let sheet_name =
                    unique_sheet_name(&pending.sheet_names, &source_sheet.name, &fallback_time);
                let worksheet = pending.workbook.add_worksheet();
                worksheet
                    .set_name(&sheet_name)
                    .with_context(|| format!("name extracted worksheet {sheet_name:?}"))?;

                let (max_row, max_column, non_empty_cells, import_operation) = if full_sheet {
                    let max_row = source_sheet.max_row.max(1);
                    let max_column = source_sheet.max_column.max(1);
                    ensure!(
                        max_row <= 1_048_576 && max_column <= EXCEL_MAX_COLUMNS,
                        "source worksheet {:?} exceeds Excel's row or column limit",
                        source_sheet.name
                    );
                    let source_range = format!("A1:{}{}", column_name(max_column), max_row);
                    let import_operation = EditOperation::ImportRange {
                        source_workbook: source_path.clone(),
                        source_sheet: source_sheet.name.clone(),
                        source_range,
                        target_sheet: sheet_name.clone(),
                        target_start: "A1".to_owned(),
                        copy_values: true,
                        copy_formulas: false,
                        style_mode: RangeStyleMode::Full,
                        copy_comments: false,
                        copy_merges: true,
                        copy_row_heights: true,
                        copy_column_widths: true,
                        copy_row_hidden: true,
                        copy_column_hidden: true,
                        copy_outline_levels: true,
                        copy_all_row_dimensions: true,
                        copy_all_column_dimensions: true,
                        materialize_blank_cells: true,
                        hide_zero: false,
                    };
                    let non_empty_cells = source_sheet
                        .values
                        .values()
                        .filter(|value| !is_empty_cell(Some(value)))
                        .count();
                    (max_row, max_column, non_empty_cells, Some(import_operation))
                } else {
                    let values = values.as_ref().expect("area extraction has values");
                    let mut non_empty_cells = 0;
                    let mut max_row = 0;
                    let mut max_column = 0;
                    let datetime_format = Format::new().set_num_format("yyyy-mm-dd h:mm:ss");
                    let duration_format = Format::new().set_num_format("[hh]:mm:ss");
                    let time_format = Format::new().set_num_format("h:mm:ss");
                    for (row_index, row) in values.iter().enumerate() {
                        for (column_index, value) in row.iter().enumerate() {
                            let Some(value) = value.as_ref() else {
                                continue;
                            };
                            if is_empty_cell(Some(value)) {
                                continue;
                            }
                            let row = u32::try_from(row_index)
                                .context("extracted worksheet exceeds Excel row index range")?;
                            let column = u16::try_from(column_index)
                                .context("extracted worksheet exceeds Excel column index range")?;
                            let reader_error = write_extracted_value(
                                worksheet,
                                row,
                                column,
                                value,
                                &datetime_format,
                                &duration_format,
                                &time_format,
                            )
                            .with_context(|| {
                                format!(
                                    "write extracted value at {}{} from worksheet {:?}",
                                    column_name(column_index + 1),
                                    row_index + 1,
                                    source_sheet.name
                                )
                            })?;
                            if let Some(reader_value) = reader_error {
                                let address =
                                    format!("{}{}", column_name(column_index + 1), row_index + 1);
                                pending.error_cells.push(PendingErrorCell {
                                    sheet_name: sheet_name.clone(),
                                    address,
                                    excel_value: excel_error_value(&reader_value)?,
                                    reader_value,
                                });
                            }
                            non_empty_cells += 1;
                            max_row = max_row.max(row_index + 1);
                            max_column = max_column.max(column_index + 1);
                        }
                    }
                    ensure!(
                        non_empty_cells > 0,
                        "compacted extraction unexpectedly contains no non-empty cells"
                    );
                    (max_row, max_column, non_empty_cells, None)
                };
                if let Some(operation) = import_operation {
                    pending.import_operations.push(operation);
                }
                pending.sheet_names.push(sheet_name.clone());
                pending.sheets.push(ExpectedSheet {
                    name: sheet_name,
                    max_row,
                    max_column,
                    non_empty_cells,
                    full_sheet,
                });
                result.extracted_sheets += 1;
            }
        }
        emitter.progress.processed_files = source_index + 1;
        emitter.progress.completed_files = source_index + 1;
        emitter.progress.failed_files = result.failed_files;
        emitter.progress.processed_sheets = result.extracted_sheets;
        emitter.emit(
            "source_extract",
            "succeeded",
            format!(
                "源工作簿处理完成，本文件新增 {} 张提取工作表",
                result.extracted_sheets - extracted_sheets_before
            ),
            Some(source_path),
            Some(source_started.elapsed()),
            &["pytools-core"],
            None,
            None,
        );
    }

    let mut staged = Vec::<StagedWorkbook>::new();
    let mut temporary_files = Vec::<PathBuf>::new();
    let stage_result = (|| -> Result<()> {
        for pending in &mut outputs {
            let output_started = Instant::now();
            emitter.emit(
                "output_write",
                "started",
                "正在生成提取结果工作簿",
                Some(&pending.path),
                None,
                &["pytools-core", "excel-writer"],
                None,
                None,
            );
            if pending.existing_sha256.is_none() {
                ensure!(
                    !pending.path.exists(),
                    "refuse to overwrite existing extraction output {}",
                    pending.path.display()
                );
            }
            let parent = pending.path.parent().unwrap_or(output_dir);
            fs::create_dir_all(parent)
                .with_context(|| format!("create output directory {}", parent.display()))?;
            let stage_path = unique_stage_path(&pending.path)?;
            staged.push(StagedWorkbook {
                stage_path: stage_path.clone(),
                output_path: pending.path.clone(),
                existing_sha256: pending.existing_sha256.clone(),
            });
            temporary_files.push(stage_path.clone());
            if let Some(expected_sha256) = pending.existing_sha256.as_deref() {
                let capability_path = capability_path.context(
                    "appending to an existing extraction output requires a capability file",
                )?;
                let platform = platform
                    .context("appending to an existing extraction output requires a platform id")?;
                let skeleton_path = unique_stage_path(&pending.path)?;
                temporary_files.push(skeleton_path.clone());
                let skeleton_write_started = Instant::now();
                if let Err(error) = pending.workbook.save(&skeleton_path) {
                    emitter.emit(
                        "skeleton_write",
                        "failed",
                        format!("待追加工作表写入失败：{error}"),
                        Some(&pending.path),
                        Some(skeleton_write_started.elapsed()),
                        &["pytools-core", "excel-writer", "rust_xlsxwriter"],
                        Some("rust_xlsxwriter"),
                        Some("skeleton_write_error"),
                    );
                    return Err(error).with_context(|| {
                        format!("save extraction skeleton {}", skeleton_path.display())
                    });
                }
                emitter.emit(
                    "skeleton_write",
                    "succeeded",
                    "待追加工作表已写入临时簿",
                    Some(&pending.path),
                    Some(skeleton_write_started.elapsed()),
                    &["pytools-core", "excel-writer", "rust_xlsxwriter"],
                    Some("rust_xlsxwriter"),
                    None,
                );

                let ready_path = if pending.import_operations.is_empty() {
                    skeleton_path.clone()
                } else {
                    let ready_path = unique_stage_path(&pending.path)?;
                    temporary_files.push(ready_path.clone());
                    let modifier_started = Instant::now();
                    let engine = match apply_full_sheet_imports(
                        &skeleton_path,
                        &ready_path,
                        &pending.import_operations,
                        capability_path,
                        platform,
                    ) {
                        Ok(engine) => engine,
                        Err(error) => {
                            emitter.emit(
                                "excel_modifier",
                                "failed",
                                format!("整表导入修改失败：{error:#}"),
                                Some(&pending.path),
                                Some(modifier_started.elapsed()),
                                &["pytools-core", "excel-editor"],
                                None,
                                Some("excel_modifier_error"),
                            );
                            return Err(error);
                        }
                    };
                    emitter.emit(
                        "excel_modifier",
                        "succeeded",
                        "ExcelModifier 已完成整表导入",
                        Some(&pending.path),
                        Some(modifier_started.elapsed()),
                        &["pytools-core", "excel-editor"],
                        Some(&engine),
                        None,
                    );
                    fs::remove_file(&skeleton_path).with_context(|| {
                        format!("remove extraction skeleton {}", skeleton_path.display())
                    })?;
                    ready_path
                };

                let modifier_started = Instant::now();
                let (modifier_engine, append_operations) = match apply_append_imports(
                    &pending.path,
                    &ready_path,
                    &stage_path,
                    expected_sha256,
                    &pending.sheets,
                    capability_path,
                    platform,
                ) {
                    Ok(execution) => execution,
                    Err(error) => {
                        emitter.emit(
                            "excel_modifier",
                            "failed",
                            format!("追加提取工作表失败：{error:#}"),
                            Some(&pending.path),
                            Some(modifier_started.elapsed()),
                            &["pytools-core", "excel-editor"],
                            None,
                            Some("excel_modifier_error"),
                        );
                        return Err(error);
                    }
                };
                emitter.emit(
                    "excel_modifier",
                    "succeeded",
                    "ExcelModifier 已把提取工作表追加到既有输出簿候选",
                    Some(&pending.path),
                    Some(modifier_started.elapsed()),
                    &["pytools-core", "excel-editor"],
                    Some(&modifier_engine),
                    None,
                );
                let validation_started = Instant::now();
                if let Err(error) = validate_full_sheet_imports(&append_operations, &stage_path) {
                    emitter.emit(
                        "append_validation",
                        "failed",
                        format!("追加工作表验证失败：{error:#}"),
                        Some(&pending.path),
                        Some(validation_started.elapsed()),
                        &["pytools-core", "excel-editor", "excel-reader"],
                        Some("calamine"),
                        Some("append_validation_error"),
                    );
                    return Err(error);
                }
                fs::remove_file(&ready_path).with_context(|| {
                    format!("remove appended extraction source {}", ready_path.display())
                })?;
                emitter.emit(
                    "append_validation",
                    "succeeded",
                    "追加工作表的值和类型验证通过",
                    Some(&pending.path),
                    Some(validation_started.elapsed()),
                    &["pytools-core", "excel-editor", "excel-reader"],
                    Some("calamine"),
                    None,
                );
            } else if pending.import_operations.is_empty() {
                if let Err(error) = pending.workbook.save(&stage_path) {
                    emitter.emit(
                        "output_write",
                        "failed",
                        format!("提取结果工作簿写入失败：{error}"),
                        Some(&pending.path),
                        Some(output_started.elapsed()),
                        &["pytools-core", "excel-writer", "rust_xlsxwriter"],
                        Some("rust_xlsxwriter"),
                        Some("output_write_error"),
                    );
                    return Err(error)
                        .with_context(|| format!("save staged workbook {}", stage_path.display()));
                }
                emitter.emit(
                    "output_write",
                    "succeeded",
                    "提取结果工作簿已写入",
                    Some(&pending.path),
                    Some(output_started.elapsed()),
                    &["pytools-core", "excel-writer", "rust_xlsxwriter"],
                    Some("rust_xlsxwriter"),
                    None,
                );
            } else {
                let capability_path =
                    capability_path.context("full-sheet extraction requires a capability file")?;
                let platform = platform.context("full-sheet extraction requires a platform id")?;
                let skeleton_path = unique_stage_path(&pending.path)?;
                temporary_files.push(skeleton_path.clone());
                let skeleton_write_started = Instant::now();
                if let Err(error) = pending.workbook.save(&skeleton_path) {
                    emitter.emit(
                        "skeleton_write",
                        "failed",
                        format!("整表导入骨架写入失败：{error}"),
                        Some(&pending.path),
                        Some(skeleton_write_started.elapsed()),
                        &["pytools-core", "excel-writer", "rust_xlsxwriter"],
                        Some("rust_xlsxwriter"),
                        Some("skeleton_write_error"),
                    );
                    return Err(error).with_context(|| {
                        format!("save extraction skeleton {}", skeleton_path.display())
                    });
                }
                emitter.emit(
                    "skeleton_write",
                    "succeeded",
                    "整表导入骨架已写入",
                    Some(&pending.path),
                    Some(skeleton_write_started.elapsed()),
                    &["pytools-core", "excel-writer", "rust_xlsxwriter"],
                    Some("rust_xlsxwriter"),
                    None,
                );
                let mut workbook_profile = inspect_xlsx_profile(&skeleton_path)?;
                let mut imported_source_profiles =
                    BTreeMap::<String, excel_editor::WorkbookProfile>::new();
                for operation in &pending.import_operations {
                    if let EditOperation::ImportRange {
                        source_workbook, ..
                    } = operation
                    {
                        let source_key = normalized_path_key(source_workbook)?;
                        let source_profile =
                            if let Some(profile) = imported_source_profiles.get(&source_key) {
                                profile.clone()
                            } else {
                                let profile =
                                    inspect_xlsx_profile(source_workbook).with_context(|| {
                                        format!(
                                            "inventory full-sheet extraction source {}",
                                            source_workbook.display()
                                        )
                                    })?;
                                imported_source_profiles.insert(source_key, profile.clone());
                                profile
                            };
                        ensure!(
                            source_profile.inventory_complete,
                            "full-sheet extraction source has an incomplete OOXML inventory: {}",
                            source_workbook.display()
                        );
                        workbook_profile
                            .features_present
                            .extend(source_profile.features_present);
                    }
                }
                let plan = ExcelEditPlan {
                    request_id: format!(
                        "pytools-extract-sheets-{}",
                        Local::now().format("%Y%m%d%H%M%S%f")
                    ),
                    input: skeleton_path.clone(),
                    output: stage_path.clone(),
                    format: WorkbookFormat::Xlsx,
                    workbook: workbook_profile,
                    operations: std::mem::take(&mut pending.import_operations),
                    preserve: [
                        PreservationRequirement::CellValuesAndTypes,
                        PreservationRequirement::Styles,
                        PreservationRequirement::MergedRanges,
                        PreservationRequirement::RowAndColumnDimensions,
                        PreservationRequirement::UnmodifiedPackageParts,
                        PreservationRequirement::SheetIdentityAndOrder,
                    ]
                    .into_iter()
                    .collect(),
                    drop_allowed: Default::default(),
                };
                plan.validate()?;
                let import_operations = plan.operations.clone();
                let modifier_started = Instant::now();
                let execution =
                    match apply_plan_with_capability_file(&plan, capability_path, platform) {
                        Ok(execution) => execution,
                        Err(error) => {
                            emitter.emit(
                                "excel_modifier",
                                "failed",
                                format!("整表导入修改失败：{error:#}"),
                                Some(&pending.path),
                                Some(modifier_started.elapsed()),
                                &["pytools-core", "excel-editor"],
                                None,
                                Some("excel_modifier_error"),
                            );
                            return Err(error).with_context(|| {
                                "apply full-sheet extraction through ExcelModifier"
                            });
                        }
                    };
                let modifier_engine = execution.engine.to_string();
                emitter.emit(
                    "excel_modifier",
                    "succeeded",
                    "ExcelModifier 已完成整表导入",
                    Some(&pending.path),
                    Some(modifier_started.elapsed()),
                    &["pytools-core", "excel-editor"],
                    Some(&modifier_engine),
                    None,
                );
                let validation_started = Instant::now();
                if let Err(error) = validate_full_sheet_imports(&import_operations, &stage_path) {
                    emitter.emit(
                        "full_sheet_validation",
                        "failed",
                        format!("整表导入结果验证失败：{error:#}"),
                        Some(&pending.path),
                        Some(validation_started.elapsed()),
                        &["pytools-core", "excel-editor", "calamine"],
                        Some("calamine"),
                        Some("full_sheet_validation_error"),
                    );
                    return Err(error);
                }
                emitter.emit(
                    "full_sheet_validation",
                    "succeeded",
                    "整表导入的值与结构验证通过",
                    Some(&pending.path),
                    Some(validation_started.elapsed()),
                    &["pytools-core", "excel-editor", "calamine"],
                    Some("calamine"),
                    None,
                );
                fs::remove_file(&skeleton_path).with_context(|| {
                    format!("remove extraction skeleton {}", skeleton_path.display())
                })?;
            }
            let mut validated_stage_path = stage_path.clone();
            if !pending.error_cells.is_empty() {
                let patch_started = Instant::now();
                emitter.emit(
                    "ooxml_error_patch",
                    "started",
                    format!("正在保留 {} 个 Excel 错误值", pending.error_cells.len()),
                    Some(&pending.path),
                    None,
                    &["pytools-core", "ooxml-patch"],
                    Some("OOXML Patch"),
                    None,
                );
                let patched_path = unique_stage_path(&pending.path)?;
                temporary_files.push(patched_path.clone());
                if let Err(error) =
                    patch_existing_error_cells(&stage_path, &patched_path, &pending.error_cells)
                {
                    emitter.emit(
                        "ooxml_error_patch",
                        "failed",
                        format!("提取结果中的 Excel 错误值保留失败：{error:#}"),
                        Some(&pending.path),
                        Some(patch_started.elapsed()),
                        &["pytools-core", "ooxml-patch"],
                        Some("OOXML Patch"),
                        Some("ooxml_error_patch_error"),
                    );
                    return Err(error);
                }
                fs::remove_file(&stage_path).with_context(|| {
                    format!(
                        "remove pre-patch extraction candidate {}",
                        stage_path.display()
                    )
                })?;
                staged
                    .last_mut()
                    .context("staged extraction workbook is missing")?
                    .stage_path = patched_path.clone();
                validated_stage_path = patched_path;
                emitter.emit(
                    "ooxml_error_patch",
                    "succeeded",
                    "提取结果中的 Excel 错误值已保留并通过回读验证",
                    Some(&pending.path),
                    Some(patch_started.elapsed()),
                    &["pytools-core", "ooxml-patch"],
                    Some("OOXML Patch"),
                    None,
                );
            }
            let validation_started = Instant::now();
            if let Err(error) = validate_staged_workbook_with_prefix(
                &validated_stage_path,
                &pending.existing_sheet_names,
                &pending.sheets,
            ) {
                emitter.emit(
                    "output_validation",
                    "failed",
                    format!("提取输出簿验证失败：{error:#}"),
                    Some(&pending.path),
                    Some(validation_started.elapsed()),
                    &["pytools-core", "excel-reader", "calamine"],
                    Some("calamine"),
                    Some("output_validation_error"),
                );
                return Err(error);
            }
            emitter.emit(
                "output_validation",
                "succeeded",
                "提取输出簿的工作表名称、尺寸和非空单元格计数验证通过",
                Some(&pending.path),
                Some(validation_started.elapsed()),
                &["pytools-core", "excel-reader", "calamine"],
                Some("calamine"),
                None,
            );
        }
        Ok(())
    })();
    if let Err(error) = stage_result {
        remove_staged_files(&staged);
        remove_temp_files(&temporary_files);
        return Err(error);
    }

    let mut committed_new_outputs = Vec::<PathBuf>::new();
    for workbook in &staged {
        let stage_path = &workbook.stage_path;
        let output_path = &workbook.output_path;
        let commit_elapsed = if let Some(expected_sha256) = workbook.existing_sha256.as_deref() {
            let commit = replace_target_from_candidate(stage_path, output_path, expected_sha256);
            let commit = match commit {
                Ok(commit) => commit,
                Err(error) => {
                    remove_staged_files(&staged);
                    remove_temp_files(&temporary_files);
                    for committed_path in &committed_new_outputs {
                        let _ = fs::remove_file(committed_path);
                    }
                    return Err(error).with_context(|| {
                        format!(
                            "safely append extracted worksheets to {}",
                            output_path.display()
                        )
                    });
                }
            };
            if let Err(error) = fs::remove_file(stage_path) {
                emitter.emit(
                    "output_cleanup",
                    "warning",
                    format!("追加结果已提交；临时候选清理失败：{error}"),
                    Some(stage_path),
                    None,
                    &["pytools-core", "filesystem"],
                    None,
                    Some("temporary_cleanup_error"),
                );
            }
            commit.elapsed_ms
        } else {
            if output_path.exists() {
                remove_staged_files(&staged);
                remove_temp_files(&temporary_files);
                for committed_path in &committed_new_outputs {
                    let _ = fs::remove_file(committed_path);
                }
                bail!(
                    "refuse to overwrite extraction output created during run: {}",
                    output_path.display()
                );
            }
            if let Err(error) = fs::hard_link(stage_path, output_path) {
                remove_staged_files(&staged);
                remove_temp_files(&temporary_files);
                for committed_path in &committed_new_outputs {
                    let _ = fs::remove_file(committed_path);
                }
                return Err(error).with_context(|| {
                    format!(
                        "commit extraction output without replacing an existing file {}",
                        output_path.display()
                    )
                });
            }
            if let Err(error) = fs::remove_file(stage_path) {
                let _ = fs::remove_file(output_path);
                remove_staged_files(&staged);
                remove_temp_files(&temporary_files);
                for committed_path in &committed_new_outputs {
                    let _ = fs::remove_file(committed_path);
                }
                return Err(error)
                    .with_context(|| format!("remove staged workbook {}", stage_path.display()));
            }
            committed_new_outputs.push(output_path.clone());
            0
        };
        emitter.emit(
            "output_commit",
            "succeeded",
            if workbook.existing_sha256.is_some() {
                "提取工作表已追加并安全替换既有输出簿"
            } else {
                "提取结果已安全提交到目标路径"
            },
            Some(output_path),
            Some(Duration::from_millis(commit_elapsed as u64)),
            &["pytools-core", "file_commit"],
            if workbook.existing_sha256.is_some() {
                Some("file_commit atomic replacement")
            } else {
                Some("hard_link create-only publication")
            },
            None,
        );
        result
            .paths
            .push(output_path.to_string_lossy().into_owned());
        result.output_files += 1;
    }

    emitter.progress.processed_files = source_workbooks.len();
    emitter.progress.completed_files = source_workbooks.len();
    emitter.progress.failed_files = result.failed_files;
    emitter.progress.processed_sheets = result.extracted_sheets;
    emitter.emit(
        "run",
        "succeeded",
        format!(
            "工作表提取完成：读取 {} 个文件，提取 {} 张工作表，输出 {} 个文件，失败 {} 个",
            result.files, result.extracted_sheets, result.output_files, result.failed_files
        ),
        None,
        Some(emitter.started.elapsed()),
        &["pytools-core"],
        None,
        None,
    );

    Ok(result)
}

fn sparse_sheet(
    sheet: &SheetValues,
    subday_date_cells: Option<&BTreeMap<String, f64>>,
) -> Result<SparseSheet> {
    ensure!(
        sheet.max_column <= EXCEL_MAX_COLUMNS,
        "worksheet {:?} exceeds Excel's column limit",
        sheet.name
    );
    let mut values = BTreeMap::new();
    for (address, value) in &sheet.values {
        let (row, column) = parse_cell(address)
            .with_context(|| format!("invalid Calamine cell coordinate {address:?}"))?;
        ensure!(
            row <= sheet.max_row && column <= sheet.max_column,
            "Calamine cell {address} falls outside worksheet dimensions"
        );
        let projected = if let Some(serial) = subday_date_cells
            .and_then(|cells| cells.get(address))
            .copied()
        {
            json!({
                "type": "pytools_time_only",
                "value": serial
            })
        } else {
            serde_json::to_value(value)?
        };
        values.insert((row, column), projected);
    }
    Ok(SparseSheet {
        name: sheet.name.clone(),
        max_row: sheet.max_row,
        max_column: sheet.max_column,
        values,
    })
}

fn read_subday_date_cells(
    path: &Path,
    sheets: &[SheetValues],
) -> Result<BTreeMap<String, BTreeMap<String, f64>>> {
    let is_ooxml = path
        .extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| {
            extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm")
        });
    if !is_ooxml {
        return Ok(BTreeMap::new());
    }

    let has_epoch_candidate = sheets.iter().any(|sheet| {
        sheet.values.values().any(
            |value| matches!(value, CellValue::DateTimeIso(iso) if is_excel_epoch_candidate(iso)),
        )
    });
    if !has_epoch_candidate {
        return Ok(BTreeMap::new());
    }

    let date_1904 = uses_1904_date_system(path)?;
    let (epoch_date, next_day) = excel_epoch_dates(date_1904);
    let raw_sheets = read_value_columns(path, EXCEL_MAX_COLUMNS)
        .with_context(|| format!("read raw epoch-cell serials from {}", path.display()))?
        .into_iter()
        .map(|sheet| (sheet.name, sheet.values))
        .collect::<BTreeMap<_, _>>();
    let mut requests = Vec::new();
    let mut time_serials = Vec::new();
    for sheet in sheets {
        for (address, value) in &sheet.values {
            if let CellValue::DateTimeIso(iso) = value {
                let date = iso
                    .split_once('T')
                    .or_else(|| iso.split_once(' '))
                    .map_or(iso.as_str(), |(date, _)| date);
                if date != epoch_date && date != next_day {
                    continue;
                }
                let Some(CellValue::DateTimeSerial(raw_serial)) = raw_sheets
                    .get(&sheet.name)
                    .and_then(|values| values.get(address))
                else {
                    continue;
                };
                let Some(time_serial) = openpyxl_subday_time_serial(*raw_serial) else {
                    continue;
                };
                requests.push(CellNumberFormatRequest {
                    sheet_name: sheet.name.clone(),
                    cell_reference: address.clone(),
                });
                time_serials.push(time_serial);
            }
        }
    }
    if requests.is_empty() {
        return Ok(BTreeMap::new());
    }

    let formats = read_xlsx_cell_number_formats(path, &requests)
        .with_context(|| format!("read epoch-cell number formats from {}", path.display()))?;
    ensure!(
        requests.len() == formats.len(),
        "cell number-format reader returned an unexpected result count"
    );

    let mut cells_by_sheet = BTreeMap::<String, BTreeMap<String, f64>>::new();
    for ((request, format), time_serial) in requests.iter().zip(formats).zip(time_serials) {
        let Some(format_code) = format.format_code.as_deref() else {
            continue;
        };
        if is_openpyxl_date_format(format_code) && !is_openpyxl_timedelta_format(format_code) {
            cells_by_sheet
                .entry(request.sheet_name.clone())
                .or_default()
                .insert(format.cell_reference, time_serial);
        }
    }
    Ok(cells_by_sheet)
}

fn is_excel_epoch_candidate(value: &str) -> bool {
    let date = value
        .split_once('T')
        .or_else(|| value.split_once(' '))
        .map_or(value, |(date, _)| date);
    matches!(
        date,
        "1899-12-31" | "1900-01-01" | "1904-01-01" | "1904-01-02"
    )
}

fn excel_epoch_dates(date_1904: bool) -> (&'static str, &'static str) {
    if date_1904 {
        ("1904-01-01", "1904-01-02")
    } else {
        ("1899-12-31", "1900-01-01")
    }
}

fn openpyxl_subday_time_serial(serial: f64) -> Option<f64> {
    const MILLISECONDS_PER_DAY: f64 = 86_400_000.0;
    if !serial.is_finite() || !(0.0..1.0).contains(&serial) {
        return None;
    }
    let milliseconds = (serial * 86_400.0 * 1_000.0).round_ties_even();
    if !(0.0..MILLISECONDS_PER_DAY).contains(&milliseconds) {
        return None;
    }
    Some(milliseconds / MILLISECONDS_PER_DAY)
}

fn is_openpyxl_date_format(format_code: &str) -> bool {
    let first_section = format_code.split(';').next().unwrap_or_default();
    let chars = strip_openpyxl_date_format_literals(first_section)
        .chars()
        .collect::<Vec<_>>();
    chars.iter().enumerate().any(|(index, character)| {
        matches!(character.to_ascii_lowercase(), 'd' | 'm' | 'h' | 'y' | 's')
            && index
                .checked_sub(1)
                .and_then(|previous| chars.get(previous))
                .is_none_or(|previous| !matches!(previous, '_' | '\\'))
    })
}

fn strip_openpyxl_date_format_literals(format_code: &str) -> String {
    let chars = format_code.chars().collect::<Vec<_>>();
    let mut stripped = String::with_capacity(format_code.len());
    let mut index = 0;
    while index < chars.len() {
        match chars[index] {
            '"' => {
                index += 1;
                while index < chars.len() && chars[index] != '"' {
                    index += 1;
                }
                if index < chars.len() {
                    index += 1;
                }
            }
            '[' => {
                let Some(relative_end) = chars[index + 1..]
                    .iter()
                    .position(|character| *character == ']')
                else {
                    stripped.push(chars[index]);
                    index += 1;
                    continue;
                };
                let end = index + 1 + relative_end;
                let bracket_code = chars[index + 1..end]
                    .iter()
                    .collect::<String>()
                    .to_ascii_lowercase();
                if matches!(bracket_code.as_str(), "h" | "hh" | "m" | "mm" | "s" | "ss") {
                    stripped.extend(chars[index..=end].iter());
                }
                index = end + 1;
            }
            character => {
                stripped.push(character);
                index += 1;
            }
        }
    }
    stripped
}

fn is_openpyxl_timedelta_format(format_code: &str) -> bool {
    let first_section = format_code.split(';').next().unwrap_or_default();
    let chars = first_section.chars().collect::<Vec<_>>();
    let mut index = 0;
    while index < chars.len() {
        if chars[index] != '[' {
            index += 1;
            continue;
        }
        let Some(relative_end) = chars[index + 1..]
            .iter()
            .position(|character| *character == ']')
        else {
            return false;
        };
        let end = index + 1 + relative_end;
        let bracket_code = chars[index + 1..end]
            .iter()
            .collect::<String>()
            .to_ascii_lowercase();
        if matches!(bracket_code.as_str(), "h" | "hh" | "m" | "mm" | "s" | "ss") {
            return true;
        }
        index = end + 1;
    }
    false
}

fn extract_sparse_configured_values(
    sheet: &SparseSheet,
    area_spec: &str,
    stitch_mode: &str,
) -> Option<CellMatrix> {
    let mut areas = Vec::new();
    for area_spec in split_area_specs(area_spec) {
        if sheet.max_row == 0 || sheet.max_column == 0 {
            areas.push(Vec::new());
            continue;
        }
        let (rows, columns) =
            parse_area_indexes(area_spec.as_str(), sheet.max_row, sheet.max_column)?;
        let mut values = Vec::with_capacity(rows.len());
        for row in rows {
            let mut row_values = Vec::with_capacity(columns.len());
            for column in &columns {
                row_values.push(sheet.values.get(&(row, *column)).cloned());
            }
            values.push(row_values);
        }
        areas.push(values);
    }
    if areas.len() == 1 {
        areas.into_iter().next()
    } else {
        Some(stitch_area_values(&areas, stitch_mode))
    }
}

fn write_extracted_value(
    worksheet: &mut excel_writer::Worksheet,
    row: u32,
    column: u16,
    value: &Value,
    datetime_format: &Format,
    duration_format: &Format,
    time_format: &Format,
) -> Result<Option<String>> {
    if value.get("type").and_then(Value::as_str) == Some("pytools_time_only") {
        let serial = value
            .get("value")
            .and_then(Value::as_f64)
            .context("time-only extraction marker has no numeric serial")?;
        ensure!(
            serial.is_finite() && (0.0..1.0).contains(&serial),
            "time-only extraction serial is outside one day"
        );
        worksheet.write_number_with_format(row, column, serial, time_format)?;
        return Ok(None);
    }
    let cell_value = CellValue::from_probe_json(value)?;
    match cell_value {
        CellValue::Text(text) => {
            worksheet.write_string(row, column, text)?;
        }
        CellValue::Number(number) if number.is_finite() => {
            worksheet.write_number(row, column, number)?;
        }
        CellValue::Number(_) => bail!("non-finite numeric cell is unsupported by XLSX output"),
        CellValue::Integer(number) => {
            const MAX_EXACT_F64_INTEGER: i64 = 9_007_199_254_740_992;
            ensure!(
                (-MAX_EXACT_F64_INTEGER..=MAX_EXACT_F64_INTEGER).contains(&number),
                "integer cell exceeds exact XLSX numeric precision"
            );
            worksheet.write_number(row, column, number as f64)?;
        }
        CellValue::Boolean(value) => {
            worksheet.write_boolean(row, column, value)?;
        }
        CellValue::Empty => {}
        CellValue::DateTimeIso(value) => {
            if let Some(serial) = parse_openpyxl_iso_duration_serial(&value)? {
                let serial = normalize_openpyxl_duration_serial(serial)?;
                worksheet.write_number_with_format(row, column, serial, duration_format)?;
            } else {
                let datetime = parse_extracted_calendar_datetime(&value)?;
                worksheet.write_datetime_with_format(row, column, &datetime, datetime_format)?;
            }
        }
        CellValue::DateTimeSerial(serial) => {
            // This flow reads through `read_value_columns_with_iso_dates`: its
            // calendar values are DateTimeIso, while Calamine's elapsed
            // durations remain DateTimeSerial.
            let serial = normalize_openpyxl_duration_serial(serial)?;
            worksheet.write_number_with_format(row, column, serial, duration_format)?;
        }
        CellValue::DurationIso(_) => {
            bail!("ISO-typed duration values are unsupported by the area-only writer; output was not committed")
        }
        CellValue::Error(reader_value) => {
            worksheet.write_string(row, column, excel_error_value(&reader_value)?)?;
            return Ok(Some(reader_value));
        }
    }
    Ok(None)
}

fn parse_extracted_calendar_datetime(value: &str) -> Result<ExcelDateTime> {
    ensure!(
        value
            .bytes()
            .all(|byte| byte.is_ascii_digit() || matches!(byte, b'-' | b':' | b'.' | b'T' | b' ')),
        "unsupported date/time text; output was not committed"
    );
    let date = value
        .split_once('T')
        .or_else(|| value.split_once(' '))
        .map_or(value, |(date, _)| date);
    ensure!(
        date.len() == 10
            && date.bytes().enumerate().all(|(index, byte)| {
                if matches!(index, 4 | 7) {
                    byte == b'-'
                } else {
                    byte.is_ascii_digit()
                }
            }),
        "unsupported date/time text; output was not committed"
    );
    ensure!(
        date != "1899-12-31" && date != "1904-01-01",
        "date/time uses an Excel epoch day that is ambiguous with a time-only value; output was not committed"
    );
    ExcelDateTime::parse_from_str(value).context("parse extracted Excel calendar date/time")
}

fn parse_openpyxl_iso_duration_serial(value: &str) -> Result<Option<f64>> {
    const HOURS_IN_MICROSECONDS: i128 = 3_600_000_000;
    const MINUTES_IN_MICROSECONDS: i128 = 60_000_000;
    const MICROSECONDS_PER_SECOND: i128 = 1_000_000;
    const MICROSECONDS_PER_DAY: i128 = 86_400_000_000;

    let Some(mut remaining) = value.strip_prefix("PT") else {
        return Ok(None);
    };
    let mut total_microseconds = 0_i128;
    let mut found_component = false;

    if let Some((component, rest)) = take_iso_duration_integer_component(remaining, 'H') {
        let hours = component
            .parse::<i128>()
            .context("parse ISO duration hours")?;
        total_microseconds = total_microseconds
            .checked_add(
                hours
                    .checked_mul(HOURS_IN_MICROSECONDS)
                    .context("ISO duration hours overflow")?,
            )
            .context("ISO duration overflow")?;
        remaining = rest;
        found_component = true;
    }

    if let Some((component, rest)) = take_iso_duration_integer_component(remaining, 'M') {
        let minutes = component
            .parse::<i128>()
            .context("parse ISO duration minutes")?;
        total_microseconds = total_microseconds
            .checked_add(
                minutes
                    .checked_mul(MINUTES_IN_MICROSECONDS)
                    .context("ISO duration minutes overflow")?,
            )
            .context("ISO duration overflow")?;
        remaining = rest;
        found_component = true;
    }

    if let Some((component, rest)) = take_iso_duration_seconds_component(remaining)? {
        let (whole_seconds, fractional_seconds) = component
            .split_once('.')
            .map_or((component, None), |(whole, fraction)| {
                (whole, Some(fraction))
            });
        let whole_seconds = whole_seconds
            .parse::<i128>()
            .context("parse ISO duration seconds")?;
        let mut component_microseconds = whole_seconds
            .checked_mul(MICROSECONDS_PER_SECOND)
            .context("ISO duration seconds overflow")?;
        if let Some(fraction) = fractional_seconds {
            let fraction_digits = fraction.len();
            let fraction = fraction
                .parse::<i128>()
                .context("parse ISO duration fractional seconds")?;
            let scale = match fraction_digits {
                1 => 100_000,
                2 => 10_000,
                3 => 1_000,
                _ => bail!("ISO duration seconds support at most 3 decimal places"),
            };
            component_microseconds = component_microseconds
                .checked_add(fraction * scale)
                .context("ISO duration fractional seconds overflow")?;
        }
        total_microseconds = total_microseconds
            .checked_add(component_microseconds)
            .context("ISO duration overflow")?;
        remaining = rest;
        found_component = true;
    }

    ensure!(
        found_component && remaining.is_empty(),
        "unsupported ISO duration text; output was not committed"
    );

    let days = total_microseconds / MICROSECONDS_PER_DAY;
    let remainder = total_microseconds % MICROSECONDS_PER_DAY;
    let total_seconds = days as f64 * 86_400.0
        + (remainder / MICROSECONDS_PER_SECOND) as f64
        + (remainder % MICROSECONDS_PER_SECOND) as f64 / MICROSECONDS_PER_SECOND as f64;
    let serial = total_seconds / 86_400.0;
    ensure!(serial.is_finite(), "ISO duration exceeds numeric range");
    Ok(Some(serial))
}

fn take_iso_duration_integer_component(value: &str, unit: char) -> Option<(&str, &str)> {
    let digit_count = value.bytes().take_while(u8::is_ascii_digit).count();
    if digit_count == 0 || !value[digit_count..].starts_with(unit) {
        return None;
    }
    Some((
        &value[..digit_count],
        &value[digit_count + unit.len_utf8()..],
    ))
}

fn take_iso_duration_seconds_component(value: &str) -> Result<Option<(&str, &str)>> {
    let digit_count = value.bytes().take_while(u8::is_ascii_digit).count();
    if digit_count == 0 {
        return Ok(None);
    }
    let mut component_end = digit_count;
    if value[digit_count..].starts_with('.') {
        let fraction_start = digit_count + 1;
        let fraction_count = value[fraction_start..]
            .bytes()
            .take_while(u8::is_ascii_digit)
            .count();
        ensure!(
            (1..=3).contains(&fraction_count),
            "ISO duration seconds support 1 to 3 decimal places"
        );
        component_end = fraction_start + fraction_count;
    }
    if !value[component_end..].starts_with('S') {
        return Ok(None);
    }
    Ok(Some((&value[..component_end], &value[component_end + 1..])))
}

fn normalize_openpyxl_duration_serial(value: f64) -> Result<f64> {
    const MAX_TIMEDELTA_DAYS: f64 = 999_999_999.0;
    const MICROSECONDS_PER_DAY: f64 = 86_400_000_000.0;
    const MICROSECONDS_PER_SECOND: u64 = 1_000_000;
    const MICROSECONDS_PER_MILLISECOND: u64 = 1_000;
    const SECONDS_PER_DAY: u64 = 86_400;

    ensure!(
        value.is_finite(),
        "non-finite elapsed duration is unsupported"
    );
    let mut days = value.floor();
    ensure!(
        (-MAX_TIMEDELTA_DAYS..=MAX_TIMEDELTA_DAYS).contains(&days),
        "elapsed duration exceeds openpyxl timedelta range"
    );

    // Python timedelta first rounds fractional days to microseconds. Openpyxl
    // then rounds that remainder to milliseconds when reading an Excel
    // duration. Preserve both ties-to-even steps, including negative values.
    let fractional_microseconds = ((value - days) * MICROSECONDS_PER_DAY).round_ties_even();
    ensure!(
        fractional_microseconds.is_finite()
            && (0.0..=MICROSECONDS_PER_DAY).contains(&fractional_microseconds),
        "elapsed duration could not be normalized"
    );
    let mut fractional_microseconds = fractional_microseconds as u64;
    if fractional_microseconds == MICROSECONDS_PER_DAY as u64 {
        days += 1.0;
        fractional_microseconds = 0;
    }

    let mut seconds = fractional_microseconds / MICROSECONDS_PER_SECOND;
    let microsecond_remainder = fractional_microseconds % MICROSECONDS_PER_SECOND;
    let mut rounded_microseconds = (microsecond_remainder as f64
        / MICROSECONDS_PER_MILLISECOND as f64)
        .round_ties_even() as u64
        * MICROSECONDS_PER_MILLISECOND;
    if rounded_microseconds == MICROSECONDS_PER_SECOND {
        seconds += 1;
        rounded_microseconds = 0;
    }
    if seconds == SECONDS_PER_DAY {
        days += 1.0;
        seconds = 0;
    }
    ensure!(
        (-MAX_TIMEDELTA_DAYS..=MAX_TIMEDELTA_DAYS).contains(&days),
        "rounded elapsed duration exceeds openpyxl timedelta range"
    );

    let normalized = days
        + (seconds as f64 + rounded_microseconds as f64 / MICROSECONDS_PER_SECOND as f64)
            / SECONDS_PER_DAY as f64;
    ensure!(
        normalized.is_finite(),
        "normalized elapsed duration is non-finite"
    );
    Ok(normalized)
}

#[cfg(test)]
mod duration_tests {
    use super::{
        excel_epoch_dates, is_excel_epoch_candidate, is_openpyxl_date_format,
        is_openpyxl_timedelta_format, normalize_openpyxl_duration_serial,
        openpyxl_subday_time_serial, parse_extracted_calendar_datetime,
        parse_openpyxl_iso_duration_serial,
    };

    const MILLISECOND_IN_DAYS: f64 = 1.0 / 86_400_000.0;

    #[test]
    fn elapsed_durations_follow_openpyxl_millisecond_ties_to_even() {
        assert_eq!(normalize_openpyxl_duration_serial(1.5).unwrap(), 1.5);
        assert_eq!(
            normalize_openpyxl_duration_serial(0.5 * MILLISECOND_IN_DAYS).unwrap(),
            0.0
        );
        assert_eq!(
            normalize_openpyxl_duration_serial(1.5 * MILLISECOND_IN_DAYS).unwrap(),
            2.0 * MILLISECOND_IN_DAYS
        );
        assert_eq!(
            normalize_openpyxl_duration_serial(-0.5 * MILLISECOND_IN_DAYS).unwrap(),
            0.0
        );
        let negative_millisecond_tie =
            normalize_openpyxl_duration_serial(-1.5 * MILLISECOND_IN_DAYS).unwrap();
        assert!((negative_millisecond_tie + 2.0 * MILLISECOND_IN_DAYS).abs() < 1e-15);
    }

    #[test]
    fn elapsed_durations_reject_values_outside_openpyxl_timedelta_range() {
        assert!(normalize_openpyxl_duration_serial(f64::INFINITY).is_err());
        assert!(normalize_openpyxl_duration_serial(-1_000_000_000.0).is_err());
        assert!(normalize_openpyxl_duration_serial(1_000_000_000.0).is_err());
    }

    #[test]
    fn openpyxl_iso_duration_subset_matches_supported_pt_components() {
        let serial = parse_openpyxl_iso_duration_serial("PT1H30M15.125S")
            .unwrap()
            .unwrap();
        let expected = 5_415.125 / 86_400.0;
        assert!((normalize_openpyxl_duration_serial(serial).unwrap() - expected).abs() < 1e-15);

        let zero = parse_openpyxl_iso_duration_serial("PT0S").unwrap().unwrap();
        assert_eq!(normalize_openpyxl_duration_serial(zero).unwrap(), 0.0);
        assert_eq!(
            parse_openpyxl_iso_duration_serial("2026-10-26T12:00:00").unwrap(),
            None
        );
        assert!(parse_openpyxl_iso_duration_serial("PT").is_err());
        assert!(parse_openpyxl_iso_duration_serial("PT1.1234S").is_err());
        assert!(parse_openpyxl_iso_duration_serial("PT1H2M3Sx").is_err());
    }

    #[test]
    fn epoch_candidate_dates_are_scoped_to_the_workbook_date_system() {
        assert_eq!(excel_epoch_dates(false), ("1899-12-31", "1900-01-01"));
        assert_eq!(excel_epoch_dates(true), ("1904-01-01", "1904-01-02"));
        assert!(is_excel_epoch_candidate("1899-12-31T00:00:00.000"));
        assert!(is_excel_epoch_candidate("1900-01-01T00:00:00.000"));
        assert!(is_excel_epoch_candidate("1904-01-01T00:00:00.000"));
        assert!(is_excel_epoch_candidate("1904-01-02T00:00:00.000"));
        assert!(!is_excel_epoch_candidate("1904-01-03T00:00:00.000"));
    }

    #[test]
    fn date_format_detection_ignores_literals_and_elapsed_tokens() {
        assert!(is_openpyxl_date_format("[$-409]h:mm"));
        assert!(is_openpyxl_date_format("yyyy-mm-dd"));
        assert!(!is_openpyxl_date_format("General"));
        assert!(!is_openpyxl_date_format("0.00 \"days\""));
        assert!(!is_openpyxl_date_format(r"0 \\m"));
        assert!(!is_openpyxl_date_format("[Red]0.00"));
        assert!(is_openpyxl_timedelta_format("[h]:mm:ss"));
        assert!(!is_openpyxl_timedelta_format("h:mm:ss"));
    }

    #[test]
    fn epoch_calendar_dates_remain_rejected_when_time_metadata_is_unavailable() {
        for value in ["1899-12-31T00:00:00.000", "1904-01-01T00:00:00.000"] {
            let error = parse_extracted_calendar_datetime(value)
                .err()
                .expect("epoch date must be rejected");
            assert!(format!("{error:#}").contains("ambiguous with a time-only value"));
        }
        assert!(parse_extracted_calendar_datetime("1900-01-01T00:00:00.000").is_ok());
        assert!(parse_extracted_calendar_datetime("1904-01-02T00:00:00.000").is_ok());
    }

    #[test]
    fn subday_time_serial_matches_openpyxl_millisecond_ties_to_even() {
        const MILLISECONDS_PER_DAY: f64 = 86_400_000.0;
        assert_eq!(openpyxl_subday_time_serial(0.0), Some(0.0));
        assert_eq!(
            openpyxl_subday_time_serial(0.5 / MILLISECONDS_PER_DAY),
            Some(0.0)
        );
        assert_eq!(
            openpyxl_subday_time_serial(1.5 / MILLISECONDS_PER_DAY),
            Some(1.0 / MILLISECONDS_PER_DAY)
        );
        assert_eq!(
            openpyxl_subday_time_serial(86_399_999.5 / MILLISECONDS_PER_DAY),
            Some(86_399_999.0 / MILLISECONDS_PER_DAY)
        );
        assert_eq!(openpyxl_subday_time_serial(1.0), None);
        assert_eq!(openpyxl_subday_time_serial(-0.5), None);
        assert_eq!(openpyxl_subday_time_serial(f64::NAN), None);
    }
}
fn validate_staged_workbook_with_prefix(
    path: &Path,
    existing_sheet_names: &[String],
    expected_sheets: &[ExpectedSheet],
) -> Result<()> {
    let actual_sheets = read_value_columns_with_iso_dates(path, EXCEL_MAX_COLUMNS)
        .with_context(|| format!("reopen staged workbook {}", path.display()))?;
    let actual_extents = read_xlsx_worksheet_extents(path)
        .with_context(|| format!("read staged worksheet extents {}", path.display()))?;
    ensure!(
        actual_sheets.len() == existing_sheet_names.len() + expected_sheets.len(),
        "staged workbook has an unexpected worksheet count"
    );
    ensure!(
        actual_extents.len() == existing_sheet_names.len() + expected_sheets.len(),
        "staged workbook has an unexpected worksheet extent count"
    );
    for ((actual, extent), expected_name) in actual_sheets
        .iter()
        .zip(&actual_extents)
        .zip(existing_sheet_names)
    {
        ensure!(
            actual.name == *expected_name && extent.name == *expected_name,
            "staged workbook changed an existing worksheet identity or order"
        );
    }
    let appended_sheets = actual_sheets
        .iter()
        .skip(existing_sheet_names.len())
        .zip(actual_extents.iter().skip(existing_sheet_names.len()));
    for ((actual, extent), expected) in appended_sheets.zip(expected_sheets) {
        ensure!(
            actual.name == expected.name && extent.name == expected.name,
            "staged worksheet name changed"
        );
        ensure!(
            extent.max_row == expected.max_row && extent.max_column == expected.max_column,
            "staged worksheet stored-cell dimensions changed for {:?}: expected {}x{}, got {}x{}",
            expected.name,
            expected.max_row,
            expected.max_column,
            extent.max_row,
            extent.max_column
        );
        let actual_non_empty = actual
            .values
            .values()
            .filter(|value| !matches!(value, CellValue::Text(text) if text.trim().is_empty()))
            .count();
        ensure!(
            actual_non_empty == expected.non_empty_cells,
            "staged worksheet cell count changed for {:?}",
            expected.name
        );
    }
    Ok(())
}

fn apply_full_sheet_imports(
    skeleton_path: &Path,
    output_path: &Path,
    operations: &[EditOperation],
    capability_path: &Path,
    platform: &str,
) -> Result<String> {
    let mut workbook_profile = inspect_xlsx_profile(skeleton_path)?;
    let mut imported_source_profiles = BTreeMap::<String, excel_editor::WorkbookProfile>::new();
    for operation in operations {
        if let EditOperation::ImportRange {
            source_workbook, ..
        } = operation
        {
            let source_key = normalized_path_key(source_workbook)?;
            let source_profile = if let Some(profile) = imported_source_profiles.get(&source_key) {
                profile.clone()
            } else {
                let profile = inspect_xlsx_profile(source_workbook).with_context(|| {
                    format!(
                        "inventory full-sheet extraction source {}",
                        source_workbook.display()
                    )
                })?;
                imported_source_profiles.insert(source_key, profile.clone());
                profile
            };
            ensure!(
                source_profile.inventory_complete,
                "full-sheet extraction source has an incomplete OOXML inventory: {}",
                source_workbook.display()
            );
            workbook_profile
                .features_present
                .extend(source_profile.features_present);
        }
    }
    let plan = ExcelEditPlan {
        request_id: format!(
            "pytools-extract-sheets-{}",
            Local::now().format("%Y%m%d%H%M%S%f")
        ),
        input: skeleton_path.to_path_buf(),
        output: output_path.to_path_buf(),
        format: WorkbookFormat::Xlsx,
        workbook: workbook_profile,
        operations: operations.to_vec(),
        preserve: [
            PreservationRequirement::CellValuesAndTypes,
            PreservationRequirement::Styles,
            PreservationRequirement::MergedRanges,
            PreservationRequirement::RowAndColumnDimensions,
            PreservationRequirement::UnmodifiedPackageParts,
            PreservationRequirement::SheetIdentityAndOrder,
        ]
        .into_iter()
        .collect(),
        drop_allowed: Default::default(),
    };
    plan.validate()?;
    let execution = apply_plan_with_capability_file(&plan, capability_path, platform)
        .context("apply full-sheet extraction through ExcelModifier")?;
    validate_full_sheet_imports(operations, output_path)?;
    Ok(execution.engine.to_string())
}

fn append_import_operations(source_path: &Path, sheets: &[ExpectedSheet]) -> Vec<EditOperation> {
    let mut operations = Vec::with_capacity(sheets.len() * 2);
    for sheet in sheets {
        operations.push(EditOperation::CreateSheet {
            name: sheet.name.clone(),
        });
        operations.push(EditOperation::ImportRange {
            source_workbook: source_path.to_path_buf(),
            source_sheet: sheet.name.clone(),
            source_range: format!(
                "A1:{}{}",
                column_name(sheet.max_column.max(1)),
                sheet.max_row.max(1)
            ),
            target_sheet: sheet.name.clone(),
            target_start: "A1".to_owned(),
            copy_values: true,
            copy_formulas: false,
            style_mode: if sheet.full_sheet {
                RangeStyleMode::Full
            } else {
                RangeStyleMode::None
            },
            copy_comments: false,
            copy_merges: sheet.full_sheet,
            copy_row_heights: sheet.full_sheet,
            copy_column_widths: sheet.full_sheet,
            copy_row_hidden: sheet.full_sheet,
            copy_column_hidden: sheet.full_sheet,
            copy_outline_levels: sheet.full_sheet,
            copy_all_row_dimensions: sheet.full_sheet,
            copy_all_column_dimensions: sheet.full_sheet,
            materialize_blank_cells: sheet.full_sheet,
            hide_zero: false,
        });
    }
    operations
}

fn apply_append_imports(
    existing_path: &Path,
    source_path: &Path,
    output_path: &Path,
    expected_sha256: &str,
    sheets: &[ExpectedSheet],
    capability_path: &Path,
    platform: &str,
) -> Result<(String, Vec<EditOperation>)> {
    verify_planning_snapshot(existing_path, Some(expected_sha256))?;
    let mut workbook_profile = inspect_xlsx_profile(existing_path).with_context(|| {
        format!(
            "inventory existing extraction output {}",
            existing_path.display()
        )
    })?;
    ensure!(
        workbook_profile.format == WorkbookFormat::Xlsx && workbook_profile.inventory_complete,
        "existing extraction output has an incomplete XLSX inventory: {}",
        existing_path.display()
    );
    let source_profile = inspect_xlsx_profile(source_path)
        .with_context(|| format!("inventory new extraction sheets {}", source_path.display()))?;
    ensure!(
        source_profile.inventory_complete,
        "new extraction sheets have an incomplete OOXML inventory: {}",
        source_path.display()
    );
    workbook_profile
        .features_present
        .extend(source_profile.features_present);
    let operations = append_import_operations(source_path, sheets);
    let plan = ExcelEditPlan {
        request_id: format!(
            "pytools-extract-sheets-append-{}",
            Local::now().format("%Y%m%d%H%M%S%f")
        ),
        input: existing_path.to_path_buf(),
        output: output_path.to_path_buf(),
        format: WorkbookFormat::Xlsx,
        workbook: workbook_profile,
        operations: operations.clone(),
        preserve: [
            PreservationRequirement::CellValuesAndTypes,
            PreservationRequirement::FormulaText,
            PreservationRequirement::CommentsAndPositions,
            PreservationRequirement::Styles,
            PreservationRequirement::MergedRanges,
            PreservationRequirement::ConditionalFormatting,
            PreservationRequirement::DataValidations,
            PreservationRequirement::DefinedNames,
            PreservationRequirement::RowAndColumnDimensions,
            PreservationRequirement::PageSettings,
            PreservationRequirement::HeaderFooterText,
            PreservationRequirement::Relationships,
            PreservationRequirement::UnmodifiedPackageParts,
            PreservationRequirement::SheetIdentityAndOrder,
            PreservationRequirement::FreezePanes,
            PreservationRequirement::AutoFilter,
        ]
        .into_iter()
        .collect(),
        drop_allowed: Default::default(),
    };
    plan.validate()?;
    let execution = apply_plan_with_capability_file(&plan, capability_path, platform)
        .context("append extracted worksheets through ExcelModifier")?;
    Ok((execution.engine.to_string(), operations))
}

fn validate_full_sheet_imports(operations: &[EditOperation], output: &Path) -> Result<()> {
    let output_dump = read_workbook(output)
        .with_context(|| format!("reopen full-sheet candidate {}", output.display()))?;
    for operation in operations {
        let EditOperation::ImportRange {
            source_workbook,
            source_sheet,
            target_sheet,
            copy_values,
            copy_formulas,
            copy_comments,
            copy_merges,
            ..
        } = operation
        else {
            continue;
        };
        ensure!(
            *copy_values && !copy_formulas,
            "full-sheet extraction must use cached values only"
        );
        ensure!(
            !copy_comments,
            "ordinary full-sheet extraction must not copy comments"
        );
        let source_dump = read_workbook(source_workbook)
            .with_context(|| format!("reopen extraction source {}", source_workbook.display()))?;
        let source = source_dump
            .sheets
            .iter()
            .find(|sheet| sheet.name == *source_sheet)
            .with_context(|| format!("source worksheet {source_sheet:?} disappeared"))?;
        let target = output_dump
            .sheets
            .iter()
            .find(|sheet| sheet.name == *target_sheet)
            .with_context(|| format!("extracted worksheet {target_sheet:?} is missing"))?;
        ensure!(
            source.values.len() == target.values.len()
                && source.values.iter().all(|(address, value)| {
                    target
                        .values
                        .get(address)
                        .is_some_and(|target_value| cell_values_match(value, target_value))
                }),
            "full-sheet extraction changed stored values or cell types in {source_sheet:?}"
        );
        ensure!(
            target.formulas.is_empty(),
            "full-sheet extraction unexpectedly retained formula text in {target_sheet:?}"
        );
        ensure!(
            target.comments.is_empty(),
            "ordinary full-sheet extraction unexpectedly copied comments into {target_sheet:?}"
        );
        if *copy_merges {
            let mut source_merges = source.merges.clone();
            let mut target_merges = target.merges.clone();
            source_merges.sort();
            target_merges.sort();
            ensure!(
                source_merges == target_merges,
                "full-sheet extraction changed merged ranges in {source_sheet:?}"
            );
        }
    }
    Ok(())
}

fn cell_values_match(source: &Value, target: &Value) -> bool {
    let (Ok(source), Ok(target)) = (
        CellValue::from_probe_json(source),
        CellValue::from_probe_json(target),
    ) else {
        return false;
    };
    match (source, target) {
        (CellValue::Number(left), CellValue::Number(right))
        | (CellValue::DateTimeSerial(left), CellValue::DateTimeSerial(right)) => {
            if !left.is_finite() || !right.is_finite() {
                return left == right;
            }
            let scale = left.abs().max(right.abs()).max(1.0);
            (left - right).abs() <= scale * 1e-15
        }
        (CellValue::Text(left), CellValue::Text(right))
        | (CellValue::DateTimeIso(left), CellValue::DateTimeIso(right))
        | (CellValue::DurationIso(left), CellValue::DurationIso(right))
        | (CellValue::Error(left), CellValue::Error(right)) => left == right,
        (CellValue::Integer(left), CellValue::Integer(right)) => left == right,
        (CellValue::Boolean(left), CellValue::Boolean(right)) => left == right,
        (CellValue::Empty, CellValue::Empty) => true,
        _ => false,
    }
}

fn normalized_path_key(path: &Path) -> Result<String> {
    let absolute = if path.is_absolute() {
        path.to_path_buf()
    } else {
        std::env::current_dir()?.join(path)
    };
    let mut normalized = PathBuf::new();
    for component in absolute.components() {
        match component {
            std::path::Component::CurDir => {}
            std::path::Component::ParentDir => {
                normalized.pop();
            }
            other => normalized.push(other.as_os_str()),
        }
    }
    let normalized = normalized.to_string_lossy().into_owned();
    Ok(if cfg!(windows) {
        normalized.to_lowercase()
    } else {
        normalized
    })
}

fn unique_stage_path(output_path: &Path) -> Result<PathBuf> {
    let parent = output_path.parent().context("output path has no parent")?;
    let stem = output_path
        .file_stem()
        .context("output path has no file name")?
        .to_string_lossy();
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .context("system clock is before Unix epoch")?
        .as_nanos();
    Ok(parent.join(format!(
        ".{stem}.{}.{}.stage.xlsx",
        std::process::id(),
        nonce
    )))
}

fn remove_staged_files(staged: &[StagedWorkbook]) {
    for workbook in staged {
        let _ = fs::remove_file(&workbook.stage_path);
    }
}

fn remove_temp_files(paths: &[PathBuf]) {
    for path in paths {
        let _ = fs::remove_file(path);
    }
}

fn column_name(mut column: usize) -> String {
    let mut name = String::new();
    while column > 0 {
        column -= 1;
        name.insert(0, (b'A' + (column % 26) as u8) as char);
        column /= 26;
    }
    name
}

fn path_suffix(path: &Path) -> Option<&str> {
    let name = path.file_name()?.to_str()?;
    let dot = name.rfind('.')?;
    if dot == 0 || dot + 1 == name.len() {
        None
    } else {
        Some(&name[dot + 1..])
    }
}

fn is_empty_cell(value: Option<&Value>) -> bool {
    match value {
        None | Some(Value::Null) => true,
        Some(Value::String(text)) => text.trim().is_empty(),
        Some(Value::Object(object))
            if object.get("type").and_then(Value::as_str) == Some("empty") =>
        {
            true
        }
        Some(Value::Object(object))
            if object.get("type").and_then(Value::as_str) == Some("string") =>
        {
            object
                .get("value")
                .and_then(Value::as_str)
                .map_or(true, |text| text.trim().is_empty())
        }
        Some(_) => false,
    }
}

fn suppress_openpyxl_merged_child_values(
    source_path: &Path,
    sheets: &mut [SparseSheet],
    selected_sheet_indexes: &BTreeSet<usize>,
) -> Result<usize> {
    let workbook = read_workbook(source_path)
        .with_context(|| format!("read merged ranges from {}", source_path.display()))?;
    ensure!(
        workbook.sheets.len() == sheets.len(),
        "merged-range inventory disagrees with worksheet count in {}",
        source_path.display()
    );
    let mut suppressed = 0;
    for &sheet_index in selected_sheet_indexes {
        let sheet = sheets.get_mut(sheet_index).with_context(|| {
            format!(
                "selected worksheet index {sheet_index} is missing from {}",
                source_path.display()
            )
        })?;
        let metadata = workbook.sheets.get(sheet_index).with_context(|| {
            format!(
                "merged-range metadata is missing worksheet index {sheet_index} in {}",
                source_path.display()
            )
        })?;
        ensure!(
            metadata.name == sheet.name,
            "merged-range inventory worksheet order disagrees at index {sheet_index} in {}",
            source_path.display()
        );
        ensure!(
            metadata.merge_status == "read_by_calamine",
            "cannot reproduce Python merged-cell semantics for worksheet {:?} in {}: {}",
            sheet.name,
            source_path.display(),
            metadata.merge_status
        );
        suppressed += suppress_merged_child_values(sheet, &metadata.merges)?;
    }
    Ok(suppressed)
}

fn suppress_merged_child_values(
    sheet: &mut SparseSheet,
    merged_ranges: &[String],
) -> Result<usize> {
    let mut bounds = Vec::with_capacity(merged_ranges.len());
    for range in merged_ranges {
        let (start, end) = range
            .split_once(':')
            .with_context(|| format!("invalid merged range {range:?} on {:?}", sheet.name))?;
        let (start_row, start_column) = parse_cell(start)
            .with_context(|| format!("invalid merged range start {start:?} on {:?}", sheet.name))?;
        let (end_row, end_column) = parse_cell(end)
            .with_context(|| format!("invalid merged range end {end:?} on {:?}", sheet.name))?;
        ensure!(
            start_row <= end_row && start_column <= end_column,
            "reversed merged range {range:?} on {:?}",
            sheet.name
        );
        bounds.push((start_row, start_column, end_row, end_column));
    }
    if bounds.is_empty() {
        return Ok(0);
    }
    let children = sheet
        .values
        .keys()
        .copied()
        .filter(|&(row, column)| {
            bounds
                .iter()
                .any(|&(start_row, start_column, end_row, end_column)| {
                    row >= start_row
                        && row <= end_row
                        && column >= start_column
                        && column <= end_column
                        && (row != start_row || column != start_column)
                })
        })
        .collect::<Vec<_>>();
    for child in &children {
        sheet.values.remove(child);
    }
    Ok(children.len())
}

fn parse_cell(token: &str) -> Option<(usize, usize)> {
    let letters = token
        .bytes()
        .take_while(|byte| byte.is_ascii_alphabetic())
        .count();
    if letters == 0 || letters == token.len() {
        return None;
    }
    let row = token[letters..].parse::<usize>().ok()?;
    if row == 0 {
        return None;
    }
    let column = column_number(&token[..letters]);
    (column > 0).then_some((row, column))
}

fn column_number(letters: &str) -> usize {
    if letters.is_empty() || letters.len() > 3 {
        return 0;
    }
    let column = letters.bytes().fold(0usize, |column, letter| {
        column
            .saturating_mul(26)
            .saturating_add((letter.to_ascii_uppercase() - b'A' + 1) as usize)
    });
    if column > 18_278 {
        0
    } else {
        column
    }
}

#[cfg(test)]
mod merged_child_values_tests {
    use super::*;

    #[test]
    fn suppresses_stored_values_in_merged_children_but_keeps_anchor() -> Result<()> {
        let mut values = BTreeMap::new();
        values.insert((1, 1), json!({"type":"string", "value":"anchor"}));
        values.insert((1, 2), json!({"type":"string", "value":"child"}));
        values.insert((2, 1), json!({"type":"string", "value":"child"}));
        values.insert((2, 2), json!({"type":"string", "value":"child"}));
        values.insert((3, 2), json!({"type":"string", "value":"outside"}));
        let mut sheet = SparseSheet {
            name: "Merged".to_owned(),
            max_row: 3,
            max_column: 2,
            values,
        };

        let suppressed = suppress_merged_child_values(&mut sheet, &["A1:B2".to_owned()])?;

        assert_eq!(suppressed, 3);
        assert_eq!(sheet.values.len(), 2);
        assert!(sheet.values.contains_key(&(1, 1)));
        assert!(sheet.values.contains_key(&(3, 2)));
        Ok(())
    }
}

fn ordered<T: Ord>(left: T, right: T) -> (T, T) {
    if left <= right {
        (left, right)
    } else {
        (right, left)
    }
}

fn bounded_range(start: usize, end: usize, max: usize) -> Vec<usize> {
    let lower = start.max(1);
    let upper = end.min(max);
    if lower > upper {
        Vec::new()
    } else {
        (lower..=upper).collect()
    }
}
