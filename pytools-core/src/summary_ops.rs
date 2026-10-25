//! Worksheet used-range and comment-template summary workflows.

use anyhow::{Context, Result};
use chrono::Local;
use excel_reader::{
    read_value_columns_with_iso_dates, read_workbook,
    read_workbook_with_unknown_error_tokens_as_text, read_xlsx_worksheet_extents, CellValue,
    SheetDump, WorksheetExtent,
};
use excel_writer::{Format, Workbook, Worksheet};
use serde::Serialize;
use std::{
    collections::{BTreeMap, BTreeSet},
    fs::{self, OpenOptions},
    io::{Read, Write},
    path::{Path, PathBuf},
    time::{Duration, Instant, SystemTime, UNIX_EPOCH},
};
use zip::{write::SimpleFileOptions, ZipArchive, ZipWriter};

const SPREADSHEET_MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const OFFICE_REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const COMMENT_READ_COMPONENTS: &[&str] = &[
    "pytools-core",
    "excel-reader",
    "calamine",
    "namespace-aware OOXML supplement",
];
const COMMENT_TEMPLATE_READ_COMPONENTS: &[&str] = &[
    "pytools-core",
    "excel-reader",
    "calamine",
    "Calamine ISO date projection",
    "namespace-aware OOXML supplement",
];
const USED_RANGE_READ_COMPONENTS: &[&str] = &["pytools-core", "excel-reader", "calamine"];
const USED_RANGE_COMPAT_READ_COMPONENTS: &[&str] = &[
    "pytools-core",
    "excel-reader",
    "calamine",
    "OOXML error-token normalization",
];

#[derive(Clone, Debug, Serialize)]
pub struct SummaryResult {
    pub saved: bool,
    pub rows: usize,
    pub files: usize,
    pub path: Option<PathBuf>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub execution: Option<SummaryExecution>,
}

#[derive(Clone, Debug, Serialize)]
pub struct SummaryExecution {
    pub reader_engine: Option<&'static str>,
    pub writer_engine: Option<&'static str>,
    pub read_elapsed_ms: u128,
    pub write_elapsed_ms: u128,
    pub total_elapsed_ms: u128,
    pub ooxml_error_cells_patched: usize,
    #[serde(default, skip_serializing_if = "is_zero")]
    pub unknown_error_token_cells_normalized: usize,
}

fn is_zero(value: &usize) -> bool {
    *value == 0
}

#[derive(Clone, Debug, Default, Serialize)]
pub struct SummaryProgress {
    pub total_inputs: usize,
    pub total_files: usize,
    pub current_file: Option<usize>,
    pub processed_files: usize,
    pub completed_files: usize,
    pub failed_files: usize,
    pub processed_rows: usize,
    pub total_rows: usize,
    #[serde(default, skip_serializing_if = "is_zero")]
    pub processed_sheets: usize,
}

#[derive(Clone, Debug, Serialize)]
pub struct SummaryRunEvent {
    pub sequence: usize,
    pub timestamp: String,
    pub phase: String,
    pub status: String,
    pub user_message: String,
    pub progress: SummaryProgress,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub file_name: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stage_elapsed_ms: Option<u128>,
    pub total_elapsed_ms: u128,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub components: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub actual_engine: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error_type: Option<String>,
}

struct SummaryProgressEmitter<'a> {
    started: Instant,
    sequence: usize,
    callback: Option<&'a mut dyn FnMut(SummaryRunEvent)>,
}

impl SummaryProgressEmitter<'_> {
    fn emit(
        &mut self,
        phase: &str,
        status: &str,
        user_message: impl Into<String>,
        progress: &SummaryProgress,
        file: Option<&Path>,
        stage_elapsed: Option<Duration>,
        components: &[&str],
        actual_engine: Option<&str>,
        error_type: Option<&str>,
    ) {
        let Some(callback) = self.callback.as_mut() else {
            return;
        };
        self.sequence += 1;
        let event = SummaryRunEvent {
            sequence: self.sequence,
            timestamp: Local::now().to_rfc3339(),
            phase: phase.to_owned(),
            status: status.to_owned(),
            user_message: user_message.into(),
            progress: progress.clone(),
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
        };
        (**callback)(event);
    }
}

#[derive(Clone, Debug)]
struct PendingErrorCell {
    address: String,
    reader_value: String,
    excel_value: &'static str,
}

pub fn summarize_used_ranges(sources: &[PathBuf], output_dir: &Path) -> Result<SummaryResult> {
    summarize_used_ranges_inner(sources, output_dir, None)
}

pub fn summarize_used_ranges_with_progress<F>(
    sources: &[PathBuf],
    output_dir: &Path,
    mut on_event: F,
) -> Result<SummaryResult>
where
    F: FnMut(SummaryRunEvent),
{
    summarize_used_ranges_inner(sources, output_dir, Some(&mut on_event))
}

fn summarize_used_ranges_inner(
    sources: &[PathBuf],
    output_dir: &Path,
    callback: Option<&mut dyn FnMut(SummaryRunEvent)>,
) -> Result<SummaryResult> {
    let started = Instant::now();
    let mut emitter = SummaryProgressEmitter {
        started,
        sequence: 0,
        callback,
    };
    let mut progress = SummaryProgress {
        total_inputs: sources.len(),
        total_files: sources
            .iter()
            .filter(|path| supported(path) && path.is_file())
            .count(),
        ..SummaryProgress::default()
    };
    emitter.emit(
        "run",
        "started",
        format!(
            "开始按使用区域汇总：{} 个输入，{} 个可用工作簿",
            progress.total_inputs, progress.total_files
        ),
        &progress,
        None,
        None,
        &["pytools-core"],
        None,
        None,
    );
    if let Err(error) = fs::create_dir_all(output_dir) {
        emitter.emit(
            "prepare_output",
            "failed",
            "无法准备汇总输出目录",
            &progress,
            None,
            None,
            &["pytools-core"],
            None,
            Some("output_directory_error"),
        );
        return Err(error)
            .with_context(|| format!("create summary output directory {}", output_dir.display()));
    }
    let mut workbooks = Vec::new();
    let mut max_col = 0usize;
    let mut rows_written = 0usize;
    let mut files = 0usize;
    let mut read_elapsed = Duration::ZERO;
    for (input_index, path) in sources.iter().enumerate() {
        if !supported(path) {
            emitter.emit(
                "input_filter",
                "skipped",
                format!("忽略不支持的输入 [{}/{}]", input_index + 1, sources.len()),
                &progress,
                Some(path),
                None,
                &["pytools-core"],
                None,
                None,
            );
            continue;
        }
        if !path.is_file() {
            emitter.emit(
                "input_filter",
                "skipped",
                format!("忽略不存在的输入 [{}/{}]", input_index + 1, sources.len()),
                &progress,
                Some(path),
                None,
                &["pytools-core"],
                None,
                None,
            );
            continue;
        }
        let current_file = files + 1;
        let file_name = path
            .file_name()
            .map(|name| name.to_string_lossy().into_owned())
            .unwrap_or_default();
        progress.current_file = Some(current_file);
        let stage_started = Instant::now();
        emitter.emit(
            "extent_scan",
            "started",
            format!(
                "正在扫描工作簿范围 [{current_file}/{}]：{file_name}",
                progress.total_files
            ),
            &progress,
            Some(path),
            None,
            &["pytools-core", "excel-reader"],
            Some("excel-reader::read_xlsx_worksheet_extents"),
            None,
        );
        let read_started = Instant::now();
        let extents = match read_xlsx_worksheet_extents(path) {
            Ok(extents) => extents,
            Err(error) => {
                progress.failed_files += 1;
                emitter.emit(
                    "extent_scan",
                    "failed",
                    format!(
                        "工作簿范围扫描失败 [{current_file}/{}]：{file_name}",
                        progress.total_files
                    ),
                    &progress,
                    Some(path),
                    Some(stage_started.elapsed()),
                    &["pytools-core", "excel-reader"],
                    Some("excel-reader::read_xlsx_worksheet_extents"),
                    Some("extent_scan_error"),
                );
                return Err(error).with_context(|| {
                    format!("read used-range summary extents {}", path.display())
                });
            }
        };
        let bounds_by_sheet = extents
            .iter()
            .map(|extent| (extent.name.clone(), summary_bounds(extent)))
            .collect::<BTreeMap<_, _>>();
        read_elapsed += read_started.elapsed();
        emitter.emit(
            "extent_scan",
            "succeeded",
            format!(
                "范围扫描完成 [{current_file}/{}]：{file_name}",
                progress.total_files
            ),
            &progress,
            Some(path),
            Some(stage_started.elapsed()),
            &["pytools-core", "excel-reader"],
            Some("excel-reader::read_xlsx_worksheet_extents"),
            None,
        );
        files += 1;
        for (min_row, _min_col, max_row, max_col_sheet) in bounds_by_sheet.values() {
            max_col = max_col.max(*max_col_sheet);
            if *min_row > 0 {
                rows_written += max_row.saturating_sub(*min_row) + 1;
            }
        }
        progress.processed_files = files;
        progress.total_rows = rows_written;
        workbooks.push((path.clone(), bounds_by_sheet));
    }
    if files == 0 || rows_written == 0 || max_col == 0 {
        let result = SummaryResult {
            saved: false,
            rows: 0,
            files,
            path: None,
            execution: Some(SummaryExecution {
                reader_engine: None,
                writer_engine: None,
                read_elapsed_ms: read_elapsed.as_millis(),
                write_elapsed_ms: 0,
                total_elapsed_ms: started.elapsed().as_millis(),
                ooxml_error_cells_patched: 0,
                unknown_error_token_cells_normalized: 0,
            }),
        };
        emitter.emit(
            "run",
            "skipped",
            "没有可汇总的工作簿或使用区域",
            &progress,
            None,
            None,
            &["pytools-core"],
            None,
            None,
        );
        return Ok(result);
    }
    fs::create_dir_all(output_dir)?;
    let path = unique_output(
        output_dir,
        &format!(
            "{}_按使用区域汇总（全量）",
            Local::now().format("%Y%m%d_%H%M%S")
        ),
    )?;
    let write_started = std::time::Instant::now();
    emitter.emit(
        "output_workbook",
        "started",
        "正在创建汇总工作簿",
        &progress,
        None,
        None,
        &["pytools-core", "excel-writer"],
        Some("rust_xlsxwriter"),
        None,
    );
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("汇总")?;
    let header_format = Format::new().set_bold();
    sheet.write_string_with_format(0, 0, "工作簿", &header_format)?;
    sheet.write_string_with_format(0, 1, "工作表", &header_format)?;
    sheet.write_string_with_format(0, 2, "行号", &header_format)?;
    for col in 1..=max_col {
        sheet.write_string_with_format(0, (col + 2) as u16, &format!("列{col}"), &header_format)?;
    }
    sheet.set_freeze_panes(1, 0)?;
    sheet.set_margins(0.75, 0.75, 1.0, 1.0, 0.5, 0.5);
    let mut out_row = 1u32;
    let mut error_cells = Vec::new();
    let mut unknown_error_token_cells_normalized = 0usize;
    let mut read_elapsed_during_write = Duration::ZERO;
    for (file_index, (path, bounds_by_sheet)) in workbooks.into_iter().enumerate() {
        let current_file = file_index + 1;
        let file_name = path
            .file_name()
            .map(|name| name.to_string_lossy().into_owned())
            .unwrap_or_default();
        progress.current_file = Some(current_file);
        let stage_started = Instant::now();
        emitter.emit(
            "value_read",
            "started",
            format!(
                "正在读取工作簿数据 [{current_file}/{}]：{file_name}",
                progress.total_files
            ),
            &progress,
            Some(&path),
            None,
            &["pytools-core", "excel-reader", "calamine"],
            Some("calamine"),
            None,
        );
        let read_started = Instant::now();
        let read_result = match read_workbook_with_unknown_error_tokens_as_text(&path) {
            Ok(read_result) => read_result,
            Err(error) => {
                progress.failed_files += 1;
                emitter.emit(
                    "value_read",
                    "failed",
                    format!(
                        "工作簿读取失败 [{current_file}/{}]：{file_name}",
                        progress.total_files
                    ),
                    &progress,
                    Some(&path),
                    Some(stage_started.elapsed()),
                    USED_RANGE_READ_COMPONENTS,
                    Some("calamine"),
                    Some("value_read_error"),
                );
                return Err(error)
                    .with_context(|| format!("read used-range summary source {}", path.display()));
            }
        };
        let compatibility_normalization_used = read_result.compatibility_normalization_used;
        let normalized_cell_count = read_result.normalized_cell_count;
        let dump = read_result.workbook;
        read_elapsed_during_write += read_started.elapsed();
        let (reader_components, reader_engine) = if compatibility_normalization_used {
            unknown_error_token_cells_normalized += normalized_cell_count;
            emitter.emit(
                "value_read_compatibility",
                "succeeded",
                format!("已按文本兼容读取 {normalized_cell_count} 个未知错误值"),
                &progress,
                Some(&path),
                Some(read_started.elapsed()),
                USED_RANGE_COMPAT_READ_COMPONENTS,
                Some("calamine + OOXML error-token normalization"),
                None,
            );
            (
                USED_RANGE_COMPAT_READ_COMPONENTS,
                "calamine + OOXML error-token normalization",
            )
        } else {
            (USED_RANGE_READ_COMPONENTS, "calamine")
        };
        emitter.emit(
            "value_read",
            "succeeded",
            format!(
                "工作簿读取完成 [{current_file}/{}]：{file_name}",
                progress.total_files
            ),
            &progress,
            Some(&path),
            Some(stage_started.elapsed()),
            reader_components,
            Some(reader_engine),
            None,
        );
        let row_stage_started = Instant::now();
        let base = path
            .file_stem()
            .map(|s| s.to_string_lossy().into_owned())
            .unwrap_or_default();
        let mut seen_sheets = BTreeSet::new();
        for ws in &dump.sheets {
            let (min_row, min_col, max_row, max_col) =
                *bounds_by_sheet.get(&ws.name).with_context(|| {
                    format!("worksheet {:?} is missing from the extent scan", ws.name)
                })?;
            let merged_ranges = parse_merged_ranges(&ws.merges)
                .with_context(|| format!("parse merge regions for worksheet {:?}", ws.name))?;
            seen_sheets.insert(ws.name.clone());
            if min_row == 0 || max_row < min_row || max_col < min_col {
                continue;
            }
            for row in min_row..=max_row {
                sheet.write_string(out_row, 0, &base)?;
                sheet.write_string(out_row, 1, &ws.name)?;
                sheet.write_number(out_row, 2, row as f64)?;
                for col in 1..=max_col {
                    if col >= min_col {
                        let address = format!("{}{}", column_letters(col), row);
                        if let Some(value) = ws.values.get(&address) {
                            if is_merged_non_anchor(&merged_ranges, row, col) {
                                continue;
                            }
                            if value.get("type").and_then(serde_json::Value::as_str)
                                == Some("error")
                            {
                                let reader_value = value
                                    .get("value")
                                    .and_then(serde_json::Value::as_str)
                                    .unwrap_or_default()
                                    .to_owned();
                                error_cells.push(PendingErrorCell {
                                    address: format!("{}{}", column_letters(col + 3), out_row + 1),
                                    excel_value: excel_error_value(&reader_value)?,
                                    reader_value,
                                });
                            } else {
                                write_probe_value(sheet, out_row, (col + 2) as u16, value)?;
                            }
                        }
                    }
                }
                out_row += 1;
                progress.processed_rows = out_row.saturating_sub(1) as usize;
                if progress.processed_rows > 0 && progress.processed_rows % 5_000 == 0 {
                    emitter.emit(
                        "row_write",
                        "running",
                        format!(
                            "已汇总 {} / {} 行",
                            progress.processed_rows, progress.total_rows
                        ),
                        &progress,
                        Some(&path),
                        Some(row_stage_started.elapsed()),
                        &["pytools-core", "rust_xlsxwriter"],
                        Some("rust_xlsxwriter"),
                        None,
                    );
                }
            }
        }
        anyhow::ensure!(
            seen_sheets.len() == bounds_by_sheet.len(),
            "worksheet set changed between extent scan and value read for {}",
            path.display()
        );
        progress.completed_files += 1;
        emitter.emit(
            "row_write",
            "succeeded",
            format!(
                "工作簿汇总完成 [{current_file}/{}]：{file_name}",
                progress.total_files
            ),
            &progress,
            Some(&path),
            Some(row_stage_started.elapsed()),
            &["pytools-core", "rust_xlsxwriter"],
            Some("rust_xlsxwriter"),
            None,
        );
    }
    if error_cells.is_empty() {
        let save_started = Instant::now();
        emitter.emit(
            "output_save",
            "started",
            "正在保存汇总工作簿",
            &progress,
            None,
            None,
            &["rust_xlsxwriter"],
            Some("rust_xlsxwriter"),
            None,
        );
        if let Err(error) = workbook.save(&path) {
            emitter.emit(
                "output_save",
                "failed",
                "汇总工作簿保存失败",
                &progress,
                None,
                Some(save_started.elapsed()),
                &["rust_xlsxwriter"],
                Some("rust_xlsxwriter"),
                Some("output_save_error"),
            );
            return Err(error)
                .with_context(|| format!("save used-range summary {}", path.display()));
        }
        emitter.emit(
            "output_save",
            "succeeded",
            "汇总工作簿已保存",
            &progress,
            None,
            Some(save_started.elapsed()),
            &["rust_xlsxwriter"],
            Some("rust_xlsxwriter"),
            None,
        );
    } else {
        let save_started = Instant::now();
        emitter.emit(
            "output_save",
            "started",
            format!("正在保存汇总结果和 {} 个错误单元格", error_cells.len()),
            &progress,
            None,
            None,
            &["rust_xlsxwriter"],
            Some("rust_xlsxwriter"),
            None,
        );
        let staging_path = staging_path_for(&path)?;
        if let Err(error) = workbook.save(&staging_path) {
            let _ = fs::remove_file(&staging_path);
            emitter.emit(
                "output_save",
                "failed",
                "汇总工作簿保存失败",
                &progress,
                None,
                Some(save_started.elapsed()),
                &["rust_xlsxwriter"],
                Some("rust_xlsxwriter"),
                Some("output_save_error"),
            );
            return Err(error).with_context(|| {
                format!("save staged used-range summary {}", staging_path.display())
            });
        }
        emitter.emit(
            "output_save",
            "succeeded",
            "汇总工作簿已保存",
            &progress,
            None,
            Some(save_started.elapsed()),
            &["rust_xlsxwriter"],
            Some("rust_xlsxwriter"),
            None,
        );
        let patch_started = Instant::now();
        emitter.emit(
            "ooxml_error_patch",
            "started",
            format!("正在保留 {} 个 Excel 错误值", error_cells.len()),
            &progress,
            None,
            None,
            &["ooxml-patch"],
            Some("OOXML Patch"),
            None,
        );
        let patch_result = patch_error_cells(&staging_path, &path, &error_cells);
        let _ = fs::remove_file(&staging_path);
        if let Err(error) = patch_result {
            emitter.emit(
                "ooxml_error_patch",
                "failed",
                "Excel 错误值保留失败",
                &progress,
                None,
                Some(patch_started.elapsed()),
                &["ooxml-patch"],
                Some("OOXML Patch"),
                Some("ooxml_patch_error"),
            );
            return Err(error).with_context(|| {
                format!("preserve Excel error values in summary {}", path.display())
            });
        }
        emitter.emit(
            "ooxml_error_patch",
            "succeeded",
            "Excel 错误值已保留",
            &progress,
            None,
            Some(patch_started.elapsed()),
            &["ooxml-patch"],
            Some("OOXML Patch"),
            None,
        );
    }
    let read_elapsed_ms = (read_elapsed + read_elapsed_during_write).as_millis();
    let write_elapsed = write_started
        .elapsed()
        .saturating_sub(read_elapsed_during_write);
    emitter.emit(
        "output_workbook",
        "succeeded",
        "汇总工作簿生成完成",
        &progress,
        None,
        Some(write_elapsed),
        &["pytools-core", "excel-writer"],
        Some("rust_xlsxwriter"),
        None,
    );
    let write_elapsed_ms = write_elapsed.as_millis();
    let result = SummaryResult {
        saved: true,
        rows: out_row.saturating_sub(1) as usize,
        files,
        path: Some(path),
        execution: Some(SummaryExecution {
            reader_engine: Some(if unknown_error_token_cells_normalized > 0 {
                "calamine + OOXML error-token normalization"
            } else {
                "calamine"
            }),
            writer_engine: Some("rust_xlsxwriter"),
            read_elapsed_ms,
            write_elapsed_ms,
            total_elapsed_ms: started.elapsed().as_millis(),
            ooxml_error_cells_patched: error_cells.len(),
            unknown_error_token_cells_normalized,
        }),
    };
    emitter.emit(
        "run",
        "succeeded",
        format!(
            "按使用区域汇总完成：{} 个工作簿，{} 行",
            result.files, result.rows
        ),
        &progress,
        None,
        Some(started.elapsed()),
        &["pytools-core"],
        None,
        None,
    );
    Ok(result)
}
#[derive(Default)]
struct TemplateSpec {
    row_ranges: Vec<(usize, usize, usize, usize)>,
    col_ranges: Vec<(usize, usize, usize, usize)>,
    set_names: Vec<String>,
    set_addresses: Vec<String>,
    headers: Vec<String>,
}

pub fn summarize_by_comments(
    template_path: &Path,
    sources: &[PathBuf],
    output_dir: &Path,
) -> Result<SummaryResult> {
    summarize_by_comments_inner(template_path, sources, output_dir, None)
}

pub fn summarize_by_comments_with_progress<F>(
    template_path: &Path,
    sources: &[PathBuf],
    output_dir: &Path,
    mut on_event: F,
) -> Result<SummaryResult>
where
    F: FnMut(SummaryRunEvent),
{
    summarize_by_comments_inner(template_path, sources, output_dir, Some(&mut on_event))
}

fn summarize_by_comments_inner(
    template_path: &Path,
    sources: &[PathBuf],
    output_dir: &Path,
    callback: Option<&mut dyn FnMut(SummaryRunEvent)>,
) -> Result<SummaryResult> {
    let started = Instant::now();
    let mut emitter = SummaryProgressEmitter {
        started,
        sequence: 0,
        callback,
    };
    let mut progress = SummaryProgress {
        total_inputs: sources.len(),
        total_files: sources
            .iter()
            .filter(|path| supported(path) && path.is_file())
            .count(),
        ..SummaryProgress::default()
    };
    emitter.emit(
        "run",
        "started",
        format!(
            "开始按批注汇总：{} 个输入，{} 个可用工作簿",
            progress.total_inputs, progress.total_files
        ),
        &progress,
        None,
        None,
        &["pytools-core"],
        None,
        None,
    );

    let output_started = Instant::now();
    emitter.emit(
        "prepare_output",
        "started",
        "正在准备汇总输出目录",
        &progress,
        None,
        None,
        &["pytools-core"],
        None,
        None,
    );
    if let Err(error) = fs::create_dir_all(output_dir) {
        emitter.emit(
            "prepare_output",
            "failed",
            "无法准备汇总输出目录",
            &progress,
            None,
            Some(output_started.elapsed()),
            &["pytools-core"],
            None,
            Some("output_directory_error"),
        );
        return Err(error)
            .with_context(|| format!("create summary output directory {}", output_dir.display()));
    }
    emitter.emit(
        "prepare_output",
        "succeeded",
        "汇总输出目录已准备",
        &progress,
        None,
        Some(output_started.elapsed()),
        &["pytools-core"],
        None,
        None,
    );

    let validation_started = Instant::now();
    emitter.emit(
        "template_validation",
        "started",
        "正在检查批注汇总模板格式",
        &progress,
        Some(template_path),
        None,
        &["pytools-core"],
        None,
        None,
    );
    if !is_xlsx_or_xlsm(template_path) {
        emitter.emit(
            "template_validation",
            "failed",
            "批注汇总模板格式不受支持",
            &progress,
            Some(template_path),
            Some(validation_started.elapsed()),
            &["pytools-core"],
            None,
            Some("template_format_error"),
        );
        anyhow::bail!("summary template must be .xlsx or .xlsm");
    }
    emitter.emit(
        "template_validation",
        "succeeded",
        "批注汇总模板格式检查完成",
        &progress,
        Some(template_path),
        Some(validation_started.elapsed()),
        &["pytools-core"],
        None,
        None,
    );

    let mut read_elapsed = Duration::ZERO;
    let template_read_started = Instant::now();
    emitter.emit(
        "template_read",
        "started",
        "正在读取批注汇总模板",
        &progress,
        Some(template_path),
        None,
        COMMENT_TEMPLATE_READ_COMPONENTS,
        Some("calamine + namespace-aware OOXML supplement"),
        None,
    );
    let template = match read_workbook(template_path).and_then(|mut template| {
        overlay_template_iso_dates(template_path, &mut template.sheets)?;
        Ok(template)
    }) {
        Ok(template) => template,
        Err(error) => {
            emitter.emit(
                "template_read",
                "failed",
                "批注汇总模板读取失败",
                &progress,
                Some(template_path),
                Some(template_read_started.elapsed()),
                COMMENT_TEMPLATE_READ_COMPONENTS,
                Some("calamine + namespace-aware OOXML supplement"),
                Some("template_read_error"),
            );
            return Err(error)
                .with_context(|| format!("read summary template {}", template_path.display()));
        }
    };
    read_elapsed += template_read_started.elapsed();
    emitter.emit(
        "template_read",
        "succeeded",
        "批注汇总模板读取完成",
        &progress,
        Some(template_path),
        Some(template_read_started.elapsed()),
        COMMENT_TEMPLATE_READ_COMPONENTS,
        Some("calamine + namespace-aware OOXML supplement"),
        None,
    );

    let rules_started = Instant::now();
    emitter.emit(
        "template_rules",
        "started",
        "正在解析模板中的行列区域和表头",
        &progress,
        Some(template_path),
        None,
        &["pytools-core"],
        None,
        None,
    );
    let mut specs = BTreeMap::new();
    let mut spec_order = Vec::new();
    for sheet in &template.sheets {
        let comments = &sheet.comments;
        if comments.is_empty() {
            continue;
        }
        let row_ranges = parse_comment_ranges(comments, "行区域");
        let col_ranges = parse_comment_ranges(comments, "列区域");
        if row_ranges.is_empty() || col_ranges.is_empty() {
            continue;
        }
        let mut spec = TemplateSpec {
            row_ranges,
            col_ranges,
            ..Default::default()
        };
        for (addr, text) in ordered_comments(comments) {
            if let Some(name) = parse_set_name(text) {
                spec.set_names.push(name);
                spec.set_addresses.push(addr.to_owned());
            }
        }
        spec.headers = build_column_headers(sheet, &spec.col_ranges);
        spec_order.push(sheet.name.clone());
        specs.insert(sheet.name.clone(), spec);
    }
    if specs.is_empty() {
        emitter.emit(
            "template_rules",
            "failed",
            "模板中没有可用的批注区域",
            &progress,
            Some(template_path),
            Some(rules_started.elapsed()),
            &["pytools-core"],
            None,
            Some("template_rules_error"),
        );
        anyhow::bail!("模板未识别到可用批注区域（需包含 行区域N/#N 与 列区域N/#N）。");
    }
    let mut data: BTreeMap<String, Vec<Vec<serde_json::Value>>> = BTreeMap::new();
    let mut headers_by_sheet = BTreeMap::new();
    let mut total_rows = 0;
    for name in &spec_order {
        let spec = &specs[name];
        let set_headers = disambiguate(&spec.set_names);
        let col_headers = disambiguate(&spec.headers);
        let mut header = vec!["工作簿".to_owned(), "工作表".to_owned()];
        header.extend(set_headers);
        header.extend(col_headers);
        header.push("行号".into());
        headers_by_sheet.insert(name.clone(), header);
        data.insert(name.clone(), Vec::new());
    }
    emitter.emit(
        "template_rules",
        "succeeded",
        format!("模板区域解析完成：{} 个工作表", spec_order.len()),
        &progress,
        Some(template_path),
        Some(rules_started.elapsed()),
        &["pytools-core"],
        None,
        None,
    );
    let mut files_hit = 0usize;
    for (input_index, source) in sources.iter().enumerate() {
        if !supported(source) || !source.is_file() {
            let reason = if supported(source) {
                "输入文件不存在"
            } else {
                "输入格式不受支持"
            };
            emitter.emit(
                "input_filter",
                "skipped",
                format!("跳过{reason} [{}/{}]", input_index + 1, sources.len()),
                &progress,
                Some(source),
                None,
                &["pytools-core"],
                None,
                None,
            );
            continue;
        }
        let current_file = files_hit + progress.failed_files + 1;
        progress.current_file = Some(current_file);
        let source_read_started = Instant::now();
        emitter.emit(
            "source_read",
            "started",
            format!(
                "正在读取源工作簿 [{current_file}/{}]：{}",
                progress.total_files,
                source.file_name().unwrap_or_default().to_string_lossy()
            ),
            &progress,
            Some(source),
            None,
            COMMENT_READ_COMPONENTS,
            Some("calamine + namespace-aware OOXML supplement"),
            None,
        );
        let read_started = Instant::now();
        let dump = match read_workbook(source) {
            Ok(dump) => dump,
            Err(error) => {
                progress.failed_files += 1;
                emitter.emit(
                    "source_read",
                    "failed",
                    format!(
                        "源工作簿读取失败 [{current_file}/{}]：{}",
                        progress.total_files,
                        source.file_name().unwrap_or_default().to_string_lossy()
                    ),
                    &progress,
                    Some(source),
                    Some(source_read_started.elapsed()),
                    COMMENT_READ_COMPONENTS,
                    Some("calamine + namespace-aware OOXML supplement"),
                    Some("source_read_error"),
                );
                return Err(error)
                    .with_context(|| format!("read comment-summary source {}", source.display()));
            }
        };
        read_elapsed += read_started.elapsed();
        files_hit += 1;
        progress.processed_files = files_hit;
        emitter.emit(
            "source_read",
            "succeeded",
            format!(
                "源工作簿读取完成 [{current_file}/{}]：{}",
                progress.total_files,
                source.file_name().unwrap_or_default().to_string_lossy()
            ),
            &progress,
            Some(source),
            Some(source_read_started.elapsed()),
            COMMENT_READ_COMPONENTS,
            Some("calamine + namespace-aware OOXML supplement"),
            None,
        );

        let row_stage_started = Instant::now();
        emitter.emit(
            "row_extract",
            "started",
            format!(
                "正在按模板匹配并提取区域数据 [{current_file}/{}]",
                progress.total_files
            ),
            &progress,
            Some(source),
            None,
            &["pytools-core"],
            None,
            None,
        );
        let applicable_sheets = dump
            .sheets
            .iter()
            .filter_map(|sheet| {
                let spec_key = if specs.contains_key(&sheet.name) {
                    Some(sheet.name.as_str())
                } else if specs.contains_key("模板") {
                    Some("模板")
                } else {
                    None
                }?;
                Some((sheet, spec_key.to_owned()))
            })
            .collect::<Vec<_>>();
        let source_row_count = applicable_sheets
            .iter()
            .map(|(_, spec_key)| {
                specs[spec_key]
                    .row_ranges
                    .iter()
                    .map(|(start, end, _, _)| end.saturating_sub(*start) + 1)
                    .sum::<usize>()
            })
            .sum::<usize>();
        progress.total_rows += source_row_count;
        let book_name = source
            .file_stem()
            .map(|s| s.to_string_lossy().into_owned())
            .unwrap_or_default();
        for (sheet, spec_key) in applicable_sheets {
            let spec = &specs[&spec_key];
            let values = data
                .get_mut(&spec_key)
                .expect("spec output table initialized");
            for (start, end, _, _) in &spec.row_ranges {
                for row in *start..=*end {
                    let mut output = vec![
                        serde_json::Value::String(book_name.clone()),
                        serde_json::Value::String(sheet.name.clone()),
                    ];
                    for address in &spec.set_addresses {
                        output.push(cell_at_or_merged(sheet, address));
                    }
                    for (_, _, start_col, end_col) in &spec.col_ranges {
                        for col in *start_col..=*end_col {
                            output.push(cell_at_or_merged(
                                sheet,
                                &format!("{}{}", column_letters(col), row),
                            ));
                        }
                    }
                    output.push(serde_json::Value::from(row as u64));
                    values.push(output);
                    total_rows += 1;
                    progress.processed_rows += 1;
                    if progress.processed_rows > 0 && progress.processed_rows % 5_000 == 0 {
                        emitter.emit(
                            "row_extract",
                            "running",
                            format!(
                                "已提取 {} / {} 行",
                                progress.processed_rows, progress.total_rows
                            ),
                            &progress,
                            Some(source),
                            Some(row_stage_started.elapsed()),
                            &["pytools-core"],
                            None,
                            None,
                        );
                    }
                }
            }
        }
        progress.completed_files += 1;
        emitter.emit(
            "row_extract",
            "succeeded",
            format!(
                "源工作簿区域提取完成 [{current_file}/{}]：{} 行",
                progress.total_files, source_row_count
            ),
            &progress,
            Some(source),
            Some(row_stage_started.elapsed()),
            &["pytools-core"],
            None,
            None,
        );
    }
    if total_rows == 0 {
        let result = SummaryResult {
            saved: false,
            rows: 0,
            files: files_hit,
            path: None,
            execution: Some(SummaryExecution {
                reader_engine: Some("calamine"),
                writer_engine: None,
                read_elapsed_ms: read_elapsed.as_millis(),
                write_elapsed_ms: 0,
                total_elapsed_ms: started.elapsed().as_millis(),
                ooxml_error_cells_patched: 0,
                unknown_error_token_cells_normalized: 0,
            }),
        };
        emitter.emit(
            "run",
            "skipped",
            "没有可汇总的批注区域数据",
            &progress,
            None,
            Some(started.elapsed()),
            &["pytools-core"],
            None,
            None,
        );
        return Ok(result);
    }
    let output_prepare_started = Instant::now();
    emitter.emit(
        "output_prepare",
        "started",
        "正在准备汇总工作簿输出位置",
        &progress,
        None,
        None,
        &["pytools-core"],
        None,
        None,
    );
    if let Err(error) = fs::create_dir_all(output_dir) {
        emitter.emit(
            "output_prepare",
            "failed",
            "无法准备汇总工作簿输出位置",
            &progress,
            None,
            Some(output_prepare_started.elapsed()),
            &["pytools-core"],
            None,
            Some("output_directory_error"),
        );
        return Err(error)
            .with_context(|| format!("create summary output directory {}", output_dir.display()));
    }
    let path = match unique_output(
        output_dir,
        &format!(
            "{}_按批注汇总（模板+源文件）",
            Local::now().format("%Y%m%d_%H%M%S")
        ),
    ) {
        Ok(path) => path,
        Err(error) => {
            emitter.emit(
                "output_prepare",
                "failed",
                "无法生成汇总工作簿文件名",
                &progress,
                None,
                Some(output_prepare_started.elapsed()),
                &["pytools-core"],
                None,
                Some("output_path_error"),
            );
            return Err(error);
        }
    };
    emitter.emit(
        "output_prepare",
        "succeeded",
        "汇总工作簿输出位置已准备",
        &progress,
        None,
        Some(output_prepare_started.elapsed()),
        &["pytools-core"],
        None,
        None,
    );

    let write_started = Instant::now();
    emitter.emit(
        "output_workbook",
        "started",
        "正在生成批注汇总工作簿",
        &progress,
        None,
        None,
        &["pytools-core", "excel-writer"],
        Some("rust_xlsxwriter"),
        None,
    );
    let workbook_result = (|| -> Result<Workbook> {
        let mut workbook = Workbook::new();
        let header_format = Format::new().set_bold();
        let output_sheet_names = comment_summary_sheet_names(&spec_order);
        for (name, sheet_name) in spec_order.iter().zip(&output_sheet_names) {
            let rows = &data[name];
            let ws = workbook.add_worksheet();
            ws.set_name(sheet_name)?;
            ws.set_margins(0.75, 0.75, 1.0, 1.0, 0.5, 0.5);
            ws.set_formula_result_default("");
            for (col, header) in headers_by_sheet[name].iter().enumerate() {
                if header.starts_with('=') {
                    ws.write_formula_with_format(0, col as u16, header.as_str(), &header_format)?;
                } else {
                    ws.write_string_with_format(0, col as u16, header, &header_format)?;
                }
            }
            for (row, values) in rows.iter().enumerate() {
                for (col, value) in values.iter().enumerate() {
                    write_json(ws, (row + 1) as u32, col as u16, value)?;
                }
            }
            if !rows.is_empty() {
                ws.set_freeze_panes(1, 0)?;
            }
        }
        Ok(workbook)
    })();
    let mut workbook = match workbook_result {
        Ok(workbook) => workbook,
        Err(error) => {
            emitter.emit(
                "output_workbook",
                "failed",
                "批注汇总工作簿生成失败",
                &progress,
                None,
                Some(write_started.elapsed()),
                &["pytools-core", "excel-writer"],
                Some("rust_xlsxwriter"),
                Some("output_workbook_error"),
            );
            return Err(error).context("build comment summary workbook");
        }
    };
    emitter.emit(
        "output_workbook",
        "succeeded",
        "批注汇总工作簿已生成",
        &progress,
        None,
        Some(write_started.elapsed()),
        &["pytools-core", "excel-writer"],
        Some("rust_xlsxwriter"),
        None,
    );

    let save_started = Instant::now();
    emitter.emit(
        "output_save",
        "started",
        "正在保存批注汇总工作簿",
        &progress,
        None,
        None,
        &["rust_xlsxwriter"],
        Some("rust_xlsxwriter"),
        None,
    );
    if let Err(error) = workbook.save(&path) {
        emitter.emit(
            "output_save",
            "failed",
            "批注汇总工作簿保存失败",
            &progress,
            None,
            Some(save_started.elapsed()),
            &["rust_xlsxwriter"],
            Some("rust_xlsxwriter"),
            Some("output_save_error"),
        );
        return Err(error).with_context(|| format!("save comment summary {}", path.display()));
    }
    emitter.emit(
        "output_save",
        "succeeded",
        "批注汇总工作簿已保存",
        &progress,
        None,
        Some(save_started.elapsed()),
        &["rust_xlsxwriter"],
        Some("rust_xlsxwriter"),
        None,
    );

    progress.processed_rows = total_rows;
    let write_elapsed = write_started.elapsed();
    let result = SummaryResult {
        saved: true,
        rows: total_rows,
        files: files_hit,
        path: Some(path),
        execution: Some(SummaryExecution {
            reader_engine: Some("calamine"),
            writer_engine: Some("rust_xlsxwriter"),
            read_elapsed_ms: read_elapsed.as_millis(),
            write_elapsed_ms: write_elapsed.as_millis(),
            total_elapsed_ms: started.elapsed().as_millis(),
            ooxml_error_cells_patched: 0,
            unknown_error_token_cells_normalized: 0,
        }),
    };
    emitter.emit(
        "run",
        "succeeded",
        format!(
            "按批注汇总完成：{} 个工作簿，{} 行",
            result.files, result.rows
        ),
        &progress,
        None,
        Some(started.elapsed()),
        &["pytools-core"],
        None,
        None,
    );
    Ok(result)
}

fn supported(path: &Path) -> bool {
    path.extension()
        .and_then(|s| s.to_str())
        .is_some_and(|s| matches!(s.to_ascii_lowercase().as_str(), "xlsx" | "xlsm"))
}
fn is_xlsx_or_xlsm(path: &Path) -> bool {
    path.extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| matches!(extension.to_ascii_lowercase().as_str(), "xlsx" | "xlsm"))
}

fn excel_error_value(reader_value: &str) -> Result<&'static str> {
    match reader_value {
        "Div0" => Ok("#DIV/0!"),
        "NA" => Ok("#N/A"),
        "Name" => Ok("#NAME?"),
        "Null" => Ok("#NULL!"),
        "Num" => Ok("#NUM!"),
        "Ref" => Ok("#REF!"),
        "Value" => Ok("#VALUE!"),
        other => anyhow::bail!("unsupported Calamine error value {other:?} in used-range summary"),
    }
}

fn staging_path_for(output: &Path) -> Result<PathBuf> {
    let file_name = output
        .file_name()
        .context("summary output path has no filename")?
        .to_string_lossy();
    for _ in 0..1000 {
        let sequence = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let candidate = output.with_file_name(format!(
            ".{file_name}.{}.{}.staging",
            std::process::id(),
            sequence
        ));
        if !candidate.exists() {
            return Ok(candidate);
        }
    }
    anyhow::bail!("cannot reserve a staging file beside {}", output.display())
}

fn patch_error_cells(source: &Path, output: &Path, errors: &[PendingErrorCell]) -> Result<()> {
    anyhow::ensure!(
        !errors.is_empty(),
        "no error cells requested for OOXML patch"
    );
    let mut output_created = false;
    let result = (|| -> Result<()> {
        let (worksheet_part, replacement) = {
            let mut archive = ZipArchive::new(fs::File::open(source)?)
                .with_context(|| format!("open staged XLSX {}", source.display()))?;
            let part = resolve_summary_worksheet_part(&mut archive)?;
            let original = read_zip_part(&mut archive, &part)?;
            let replacement = patch_worksheet_error_cells(&original, errors)?;
            (part, replacement)
        };

        let target = OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(output)
            .with_context(|| format!("create patched summary {}", output.display()))?;
        output_created = true;
        let mut writer = ZipWriter::new(target);
        let mut archive = ZipArchive::new(fs::File::open(source)?)?;
        let mut patched_parts = 0usize;
        for index in 0..archive.len() {
            let mut entry = archive.by_index(index)?;
            let name = entry.name().to_owned();
            let mut options = SimpleFileOptions::default().compression_method(entry.compression());
            if let Some(modified) = entry.last_modified() {
                options = options.last_modified_time(modified);
            }
            if entry.is_dir() {
                writer.add_directory(name, options)?;
                continue;
            }
            writer.start_file(&name, options)?;
            if name == worksheet_part {
                writer.write_all(&replacement)?;
                patched_parts += 1;
            } else {
                std::io::copy(&mut entry, &mut writer)?;
            }
        }
        anyhow::ensure!(
            patched_parts == 1,
            "expected one summary worksheet part, patched {patched_parts}"
        );
        writer.finish()?;
        validate_error_cells(output, errors)?;
        Ok(())
    })();
    if result.is_err() && output_created {
        let _ = fs::remove_file(output);
    }
    result
}

fn resolve_summary_worksheet_part(archive: &mut ZipArchive<fs::File>) -> Result<String> {
    let workbook = read_zip_part(archive, "xl/workbook.xml")?;
    let relationships = read_zip_part(archive, "xl/_rels/workbook.xml.rels")?;
    let workbook_text = std::str::from_utf8(&workbook).context("workbook.xml must be UTF-8")?;
    let relationships_text =
        std::str::from_utf8(&relationships).context("workbook relationships must be UTF-8")?;
    let workbook_doc = roxmltree::Document::parse(workbook_text).context("parse workbook.xml")?;
    let relationships_doc =
        roxmltree::Document::parse(relationships_text).context("parse workbook relationships")?;
    let sheet = workbook_doc
        .root()
        .descendants()
        .find(|node| {
            is_xml_element(*node, SPREADSHEET_MAIN_NS, "sheet")
                && node.attribute("name") == Some("汇总")
        })
        .context("generated summary worksheet '汇总' is missing")?;
    let relationship_id = sheet
        .attribute((OFFICE_REL_NS, "id"))
        .context("summary worksheet relationship id is missing")?;
    let relationship = relationships_doc
        .root()
        .descendants()
        .find(|node| {
            is_xml_element(*node, PACKAGE_REL_NS, "Relationship")
                && node.attribute("Id") == Some(relationship_id)
        })
        .context("summary worksheet relationship is missing")?;
    anyhow::ensure!(
        relationship.attribute("TargetMode") != Some("External"),
        "summary worksheet relationship cannot be external"
    );
    let relationship_type = relationship
        .attribute("Type")
        .context("summary worksheet relationship type is missing")?;
    anyhow::ensure!(
        relationship_type.ends_with("/worksheet"),
        "summary worksheet relationship targets a non-worksheet part"
    );
    let target = relationship
        .attribute("Target")
        .context("summary worksheet relationship target is missing")?;
    normalize_package_target(target)
}

fn normalize_package_target(target: &str) -> Result<String> {
    anyhow::ensure!(
        !target.contains('\\'),
        "OOXML relationship target uses an invalid separator"
    );
    let joined = if target.starts_with('/') {
        target.trim_start_matches('/').to_owned()
    } else {
        format!("xl/{target}")
    };
    let mut parts = Vec::new();
    for part in joined.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                anyhow::ensure!(
                    parts.pop().is_some(),
                    "OOXML relationship target escapes the package root"
                );
            }
            value => parts.push(value),
        }
    }
    anyhow::ensure!(!parts.is_empty(), "OOXML relationship target is empty");
    Ok(parts.join("/"))
}

fn read_zip_part(archive: &mut ZipArchive<fs::File>, name: &str) -> Result<Vec<u8>> {
    let mut entry = archive
        .by_name(name)
        .with_context(|| format!("XLSX package is missing {name}"))?;
    let mut bytes = Vec::new();
    entry
        .read_to_end(&mut bytes)
        .with_context(|| format!("read XLSX package part {name}"))?;
    Ok(bytes)
}

fn patch_worksheet_error_cells(xml: &[u8], errors: &[PendingErrorCell]) -> Result<Vec<u8>> {
    let source = std::str::from_utf8(xml).context("worksheet XML must be UTF-8")?;
    let document = roxmltree::Document::parse(source).context("parse summary worksheet XML")?;
    let root = document.root_element();
    anyhow::ensure!(
        is_xml_element(root, SPREADSHEET_MAIN_NS, "worksheet"),
        "expected a SpreadsheetML worksheet"
    );
    let sheet_data = root
        .children()
        .find(|node| is_xml_element(*node, SPREADSHEET_MAIN_NS, "sheetData"))
        .context("summary worksheet has no sheetData")?;
    let prefix = root.lookup_prefix(SPREADSHEET_MAIN_NS);
    let cell_tag = xml_local_name(prefix, "c");
    let value_tag = xml_local_name(prefix, "v");
    let mut edits = Vec::with_capacity(errors.len());
    let mut seen = BTreeSet::new();

    for error in errors {
        anyhow::ensure!(
            seen.insert(error.address.as_str()),
            "duplicate error cell {}",
            error.address
        );
        let (row_number, column_number) = parse_a1(&error.address)
            .with_context(|| format!("invalid summary error cell {}", error.address))?;
        let row_text = row_number.to_string();
        let row = sheet_data
            .children()
            .find(|node| {
                is_xml_element(*node, SPREADSHEET_MAIN_NS, "row")
                    && node.attribute("r") == Some(row_text.as_str())
            })
            .with_context(|| format!("summary output row {row_number} is missing"))?;
        let mut last_cell = None;
        let mut insertion = None;
        for cell in row
            .children()
            .filter(|node| is_xml_element(*node, SPREADSHEET_MAIN_NS, "c"))
        {
            let cell_address = cell
                .attribute("r")
                .context("summary output contains a cell without an address")?;
            let (_, cell_column) = parse_a1(cell_address)
                .with_context(|| format!("invalid summary output cell {cell_address}"))?;
            anyhow::ensure!(
                cell_address != error.address,
                "summary error cell {} already exists",
                error.address
            );
            if cell_column > column_number {
                insertion = Some(cell.range().start);
                break;
            }
            last_cell = Some(cell);
        }
        let offset = insertion
            .or_else(|| last_cell.map(|cell| cell.range().end))
            .with_context(|| format!("summary output row {row_number} has no cell anchor"))?;
        let markup = format!(
            "<{cell_tag} r=\"{}\" t=\"e\"><{value_tag}>{}</{value_tag}></{cell_tag}>",
            error.address, error.excel_value
        );
        edits.push((offset, column_number, markup));
    }

    edits.sort_by(|left, right| right.0.cmp(&left.0).then_with(|| right.1.cmp(&left.1)));
    let mut patched = source.to_owned();
    for (offset, _, markup) in edits {
        patched.insert_str(offset, &markup);
    }
    roxmltree::Document::parse(&patched).context("validate patched summary worksheet XML")?;
    Ok(patched.into_bytes())
}

fn validate_error_cells(workbook: &Path, errors: &[PendingErrorCell]) -> Result<()> {
    let dump = read_workbook(workbook)
        .with_context(|| format!("reopen patched summary workbook {}", workbook.display()))?;
    let sheet = dump
        .sheets
        .iter()
        .find(|sheet| sheet.name == "汇总")
        .context("reopened summary worksheet is missing")?;
    for error in errors {
        let cell = sheet
            .values
            .get(&error.address)
            .with_context(|| format!("patched summary error cell {} is missing", error.address))?;
        anyhow::ensure!(
            cell.get("type").and_then(serde_json::Value::as_str) == Some("error")
                && cell.get("value").and_then(serde_json::Value::as_str)
                    == Some(error.reader_value.as_str()),
            "patched summary error cell {} did not round trip: {cell}",
            error.address
        );
    }
    Ok(())
}

fn is_xml_element(node: roxmltree::Node<'_, '_>, namespace: &str, local_name: &str) -> bool {
    node.is_element()
        && node.tag_name().name() == local_name
        && node.tag_name().namespace() == Some(namespace)
}

fn xml_local_name(prefix: Option<&str>, name: &str) -> String {
    match prefix {
        Some(prefix) => format!("{prefix}:{name}"),
        None => name.to_owned(),
    }
}
fn summary_bounds(extent: &WorksheetExtent) -> (usize, usize, usize, usize) {
    if extent.max_row == 0 || extent.max_column == 0 {
        return (1, 1, 1, 1);
    }
    (
        extent.min_row.max(1),
        extent.min_column.max(1),
        extent.max_row,
        extent.max_column,
    )
}
fn parse_merged_ranges(merges: &[String]) -> Result<Vec<(usize, usize, usize, usize)>> {
    merges
        .iter()
        .map(|reference| {
            let normalized = reference.replace('$', "");
            let (start, end) = normalized
                .split_once(':')
                .unwrap_or((normalized.as_str(), normalized.as_str()));
            let (start_row, start_column) =
                parse_a1(start).with_context(|| format!("invalid merged range start {start:?}"))?;
            let (end_row, end_column) =
                parse_a1(end).with_context(|| format!("invalid merged range end {end:?}"))?;
            anyhow::ensure!(
                start_row <= end_row && start_column <= end_column,
                "merged range has reversed bounds: {reference:?}"
            );
            Ok((start_row, start_column, end_row, end_column))
        })
        .collect()
}
fn is_merged_non_anchor(
    merges: &[(usize, usize, usize, usize)],
    row: usize,
    column: usize,
) -> bool {
    merges
        .iter()
        .any(|(min_row, min_column, max_row, max_column)| {
            row >= *min_row
                && row <= *max_row
                && column >= *min_column
                && column <= *max_column
                && (row != *min_row || column != *min_column)
        })
}
fn parse_a1(address: &str) -> Option<(usize, usize)> {
    let index = address.find(|c: char| c.is_ascii_digit())?;
    let col = address[..index].chars().try_fold(0usize, |v, c| {
        v.checked_mul(26)?
            .checked_add(c.to_ascii_uppercase() as usize - 'A' as usize + 1)
    })?;
    Some((address[index..].parse().ok()?, col))
}
fn column_letters(mut col: usize) -> String {
    let mut out = String::new();
    while col > 0 {
        out.insert(0, (b'A' + ((col - 1) % 26) as u8) as char);
        col = (col - 1) / 26;
    }
    out
}
fn write_probe_value(
    ws: &mut Worksheet,
    row: u32,
    col: u16,
    value: &serde_json::Value,
) -> Result<()> {
    let kind = value
        .get("type")
        .and_then(serde_json::Value::as_str)
        .unwrap_or("");
    let v = value.get("value").unwrap_or(&serde_json::Value::Null);
    match kind {
        "string" | "error" => {
            ws.write_string(row, col, v.as_str().unwrap_or_default())?;
        }
        "duration_iso" => {
            let text = v.as_str().unwrap_or_default();
            if let Some(serial) = iso_duration_serial(text) {
                let format = Format::new().set_num_format("[h]:mm:ss");
                ws.write_number_with_format(row, col, serial, &format)?;
            } else {
                ws.write_string(row, col, text)?;
            }
        }
        "integer" => {
            ws.write_number(row, col, v.as_i64().unwrap_or_default() as f64)?;
        }
        "datetime_iso" => {
            let text = v.as_str().unwrap_or_default();
            if let Some(serial) = iso_datetime_serial(text) {
                let format = date_format(serial);
                ws.write_number_with_format(row, col, serial, &format)?;
            } else {
                ws.write_string(row, col, text)?;
            }
        }
        "number" => {
            if let Some(n) = v.as_f64() {
                ws.write_number(row, col, n)?;
            }
        }
        "datetime_serial" => {
            if let Some(n) = v.as_f64() {
                let format = date_format(n);
                ws.write_number_with_format(row, col, n, &format)?;
            }
        }
        "boolean" => {
            ws.write_boolean(row, col, v.as_bool().unwrap_or(false))?;
        }
        _ => {}
    }
    Ok(())
}
fn write_json(ws: &mut Worksheet, row: u32, col: u16, value: &serde_json::Value) -> Result<()> {
    if value
        .get("type")
        .and_then(serde_json::Value::as_str)
        .is_some()
    {
        write_probe_value(ws, row, col, value)?;
    } else if value.is_string() {
        ws.write_string(row, col, value.as_str().unwrap_or_default())?;
    } else if value.is_boolean() {
        ws.write_boolean(row, col, value.as_bool().unwrap_or(false))?;
    } else if let Some(number) = value.as_f64() {
        ws.write_number(row, col, number)?;
    }
    Ok(())
}
fn unique_output(dir: &Path, stem: &str) -> Result<PathBuf> {
    for i in 0..10000 {
        let suffix = if i == 0 {
            String::new()
        } else {
            format!("_{i}")
        };
        let path = dir.join(format!("{stem}{suffix}.xlsx"));
        if !path.exists() {
            return Ok(path);
        }
    }
    anyhow::bail!("cannot reserve unique output under {}", dir.display())
}
fn parse_comment_ranges(
    comments: &BTreeMap<String, String>,
    keyword: &str,
) -> Vec<(usize, usize, usize, usize)> {
    let ordered = ordered_comments(comments);
    // PyCode scans comments in row-major order and overwrites each boundary,
    // so the last matching cell wins when a marker is duplicated.
    let mut ids = BTreeSet::new();
    for (_, text) in &ordered {
        if text.contains(keyword) {
            if let Some(id) = first_number(text) {
                ids.insert(id);
            }
        }
    }
    let mut out = Vec::new();
    for id in ids {
        let start = ordered
            .iter()
            .rev()
            .find(|(_, text)| {
                text.replace(' ', "").contains(&format!("{keyword}{id}"))
                    && !text.replace(' ', "").contains(&format!("{keyword}#{id}"))
            })
            .map(|(a, _)| *a);
        let end = ordered
            .iter()
            .rev()
            .find(|(_, text)| text.replace(' ', "").contains(&format!("{keyword}#{id}")))
            .map(|(a, _)| *a);
        if let (Some(a), Some(b)) = (start, end) {
            if let (Some((r1, c1)), Some((r2, c2))) = (parse_a1(a), parse_a1(b)) {
                out.push((r1.min(r2), r1.max(r2), c1.min(c2), c1.max(c2)));
            }
        }
    }
    out
}
fn ordered_comments<'a>(comments: &'a BTreeMap<String, String>) -> Vec<(&'a String, &'a String)> {
    let mut ordered = comments.iter().collect::<Vec<_>>();
    ordered.sort_by_key(|(address, _)| parse_a1(address).unwrap_or((usize::MAX, usize::MAX)));
    ordered
}

fn first_number(text: &str) -> Option<usize> {
    let normalized = text.replace('#', "");
    let start = normalized.find(|character: char| character.is_ascii_digit())?;
    normalized[start..]
        .chars()
        .take_while(char::is_ascii_digit)
        .collect::<String>()
        .parse()
        .ok()
}
fn parse_set_name(text: &str) -> Option<String> {
    let lower = text.to_lowercase();
    let ascii = lower.find("set(").map(|idx| (idx, 4));
    let fullwidth = lower.find("set（").map(|idx| (idx, "set（".len()));
    let (idx, marker_len) = match (ascii, fullwidth) {
        (Some(left), Some(right)) => {
            if left.0 <= right.0 {
                left
            } else {
                right
            }
        }
        (Some(value), None) | (None, Some(value)) => value,
        (None, None) => return None,
    };
    let rest = &text[idx + marker_len..];
    let end = rest.find([')', '）'])?;
    let name = &rest[..end];
    if name.is_empty() {
        return None;
    }
    Some(name.trim().to_owned())
}
fn overlay_template_iso_dates(path: &Path, sheets: &mut [SheetDump]) -> Result<()> {
    let max_columns = sheets
        .iter()
        .flat_map(|sheet| sheet.values.keys())
        .filter_map(|address| parse_a1(address).map(|(_, column)| column))
        .max()
        .unwrap_or(1)
        .max(1);
    for date_sheet in read_value_columns_with_iso_dates(path, max_columns)
        .context("read ISO date values from summary template")?
    {
        let Some(sheet) = sheets
            .iter_mut()
            .find(|sheet| sheet.name == date_sheet.name)
        else {
            continue;
        };
        for (address, value) in date_sheet.values {
            if let CellValue::DateTimeIso(value) = value {
                sheet.values.insert(
                    address,
                    serde_json::json!({"type":"datetime_iso", "value":value}),
                );
            }
        }
    }
    Ok(())
}

fn build_column_headers(sheet: &SheetDump, ranges: &[(usize, usize, usize, usize)]) -> Vec<String> {
    let mut out = Vec::new();
    for (sr, er, sc, ec) in ranges {
        for col in *sc..=*ec {
            let upper =
                template_cell_at_or_merged(sheet, &format!("{}{}", column_letters(col), sr));
            let lower =
                template_cell_at_or_merged(sheet, &format!("{}{}", column_letters(col), er));
            let a = probe_text(&upper);
            let b = probe_text(&lower);
            let text = if *sr < *er && !a.is_empty() && !b.is_empty() && a != b {
                format!("{a}_{b}")
            } else if !a.is_empty() {
                a
            } else if !b.is_empty() {
                b
            } else {
                column_letters(col)
            };
            out.push(text);
        }
    }
    out
}
fn cell_at_or_merged(sheet: &SheetDump, address: &str) -> serde_json::Value {
    let Some((row, col)) = parse_a1(address) else {
        return serde_json::Value::Null;
    };
    for range in &sheet.merges {
        let Some((start, end)) = range.split_once(':') else {
            continue;
        };
        if let (Some((r1, c1)), Some((r2, c2))) = (parse_a1(start), parse_a1(end)) {
            if r1.min(r2) <= row && row <= r1.max(r2) && c1.min(c2) <= col && col <= c1.max(c2) {
                return sheet
                    .values
                    .get(&format!("{}{}", column_letters(c1.min(c2)), r1.min(r2)))
                    .cloned()
                    .unwrap_or(serde_json::Value::Null);
            }
        }
    }
    sheet
        .values
        .get(address)
        .cloned()
        .unwrap_or(serde_json::Value::Null)
}

fn template_cell_at_or_merged(sheet: &SheetDump, address: &str) -> serde_json::Value {
    let Some((row, col)) = parse_a1(address) else {
        return serde_json::Value::Null;
    };
    let mut anchor = address.to_owned();
    for range in &sheet.merges {
        let Some((start, end)) = range.split_once(':') else {
            continue;
        };
        if let (Some((r1, c1)), Some((r2, c2))) = (parse_a1(start), parse_a1(end)) {
            if r1.min(r2) <= row && row <= r1.max(r2) && c1.min(c2) <= col && col <= c1.max(c2) {
                anchor = format!("{}{}", column_letters(c1.min(c2)), r1.min(r2));
                break;
            }
        }
    }
    if let Some(formula) = sheet.formulas.get(&anchor) {
        let formula = formula.strip_prefix('=').unwrap_or(formula);
        return serde_json::json!({
            "type": "string",
            "value": format!("={formula}")
        });
    }
    sheet
        .values
        .get(&anchor)
        .cloned()
        .unwrap_or(serde_json::Value::Null)
}

fn date_format(_serial: f64) -> Format {
    // openpyxl's data_only read returns date-formatted XLSX cells as midnight
    // datetimes. Appending those values to a new workbook uses this format.
    Format::new().set_num_format("yyyy-mm-dd h:mm:ss")
}

fn iso_datetime_serial(text: &str) -> Option<f64> {
    use chrono::{DateTime, NaiveDate, NaiveDateTime};

    let datetime = DateTime::parse_from_rfc3339(text)
        .map(|datetime| datetime.naive_utc())
        .or_else(|_| NaiveDateTime::parse_from_str(text, "%Y-%m-%dT%H:%M:%S%.f"))
        .or_else(|_| NaiveDateTime::parse_from_str(text, "%Y-%m-%d %H:%M:%S%.f"))
        .or_else(|_| {
            NaiveDate::parse_from_str(text, "%Y-%m-%d")
                .map(|date| date.and_hms_opt(0, 0, 0).expect("valid midnight"))
        })
        .ok()?;
    let epoch = NaiveDate::from_ymd_opt(1899, 12, 30)?.and_hms_opt(0, 0, 0)?;
    let millis = datetime.signed_duration_since(epoch).num_milliseconds();
    Some(millis as f64 / 86_400_000.0)
}

fn iso_duration_serial(text: &str) -> Option<f64> {
    let text = text.strip_prefix('P')?;
    let mut in_time = false;
    let mut number = String::new();
    let mut seconds = 0.0;
    for character in text.chars() {
        if character == 'T' {
            in_time = true;
            continue;
        }
        if character.is_ascii_digit() || character == '.' || character == '-' {
            number.push(character);
            continue;
        }
        let value = number.parse::<f64>().ok()?;
        number.clear();
        let multiplier = match (in_time, character) {
            (false, 'D') => 86_400.0,
            (true, 'H') => 3_600.0,
            (true, 'M') => 60.0,
            (true, 'S') => 1.0,
            _ => return None,
        };
        seconds += value * multiplier;
    }
    if !number.is_empty() {
        return None;
    }
    Some(seconds / 86_400.0)
}
fn probe_text(value: &serde_json::Value) -> String {
    if value.is_null() {
        String::new()
    } else {
        let cell_type = value.get("type").and_then(serde_json::Value::as_str);
        let cell_value = value.get("value").unwrap_or(value);
        match cell_type {
            Some("empty") => String::new(),
            Some("string") | Some("duration_iso") | Some("error") => {
                cell_value.as_str().unwrap_or_default().trim().to_owned()
            }
            Some("boolean") => python_bool_text(cell_value.as_bool().unwrap_or(false)),
            Some("number") | Some("integer") => python_number_text(cell_value),
            Some("datetime_iso") => {
                python_datetime_iso_text(cell_value.as_str().unwrap_or_default())
            }
            Some("datetime_serial") => cell_value
                .as_f64()
                .and_then(python_datetime_serial_text)
                .unwrap_or_default(),
            _ if let Some(text) = value.as_str() => text.trim().to_owned(),
            _ if let Some(boolean) = value.as_bool() => python_bool_text(boolean),
            _ => value.to_string(),
        }
    }
}

fn python_bool_text(value: bool) -> String {
    if value {
        "True".to_owned()
    } else {
        "False".to_owned()
    }
}

fn python_number_text(value: &serde_json::Value) -> String {
    if let Some(integer) = value.as_i64() {
        return integer.to_string();
    }
    if let Some(integer) = value.as_u64() {
        return integer.to_string();
    }
    if let Some(number) = value.as_f64() {
        if number.is_finite() && number.fract() == 0.0 && number.abs() <= 9_007_199_254_740_991.0 {
            // Excel stores ordinary integral numeric values without a decimal
            // point, which openpyxl reloads as Python int even when Calamine
            // exposes the value as f64.
            return format!("{number:.0}");
        }
        return python_float_text(number);
    }
    value.to_string()
}

pub(crate) fn python_float_text(number: f64) -> String {
    let absolute = number.abs();
    if absolute >= 1e16 || (absolute != 0.0 && absolute < 1e-4) {
        let representation = format!("{number:?}");
        if let Some((mantissa, exponent)) = representation.split_once('e') {
            if let Ok(exponent) = exponent.parse::<i32>() {
                let mantissa = mantissa.strip_suffix(".0").unwrap_or(mantissa);
                let sign = if exponent < 0 { '-' } else { '+' };
                let magnitude = format!("{:02}", exponent.abs());
                return format!("{mantissa}e{sign}{magnitude}");
            }
        }
        return representation;
    }

    let mut representation = number.to_string();
    if !representation.contains('.') && !representation.contains('e') {
        representation.push_str(".0");
    }
    representation
}

pub(crate) fn python_datetime_iso_text(value: &str) -> String {
    use chrono::{NaiveDate, NaiveDateTime};

    let datetime = NaiveDateTime::parse_from_str(value, "%Y-%m-%dT%H:%M:%S%.f")
        .or_else(|_| NaiveDateTime::parse_from_str(value, "%Y-%m-%d %H:%M:%S%.f"));
    if let Ok(datetime) = datetime {
        return python_naive_datetime_text(datetime);
    }
    if let Ok(date) = NaiveDate::parse_from_str(value, "%Y-%m-%d") {
        return format!("{} 00:00:00", date.format("%Y-%m-%d"));
    }
    value.replace('T', " ")
}

pub(crate) fn python_datetime_serial_text(serial: f64) -> Option<String> {
    use chrono::{Duration, NaiveDate, NaiveTime};

    if !serial.is_finite() {
        return None;
    }
    let milliseconds = (serial * 86_400_000.0).round();
    if milliseconds < i64::MIN as f64 || milliseconds > i64::MAX as f64 {
        return None;
    }
    let milliseconds = milliseconds as i64;
    const DAY_MILLISECONDS: i64 = 86_400_000;
    let days = milliseconds.div_euclid(DAY_MILLISECONDS);
    let day_milliseconds = milliseconds.rem_euclid(DAY_MILLISECONDS);
    let time = NaiveTime::from_num_seconds_from_midnight_opt(
        (day_milliseconds / 1_000) as u32,
        ((day_milliseconds % 1_000) * 1_000_000) as u32,
    )?;

    // openpyxl returns a time object for serials in [0, 1). For modern dates,
    // Excel's 1900-system epoch is 1899-12-30 (including its leap-year quirk).
    if (0..DAY_MILLISECONDS).contains(&milliseconds) {
        return Some(python_time_text(time));
    }
    let epoch = if days < 60 {
        NaiveDate::from_ymd_opt(1900, 1, 1)?
    } else {
        NaiveDate::from_ymd_opt(1899, 12, 30)?
    };
    let date = epoch
        .and_time(time)
        .checked_add_signed(Duration::days(if days < 60 {
            days.saturating_sub(1)
        } else {
            days
        }))?;
    Some(python_naive_datetime_text(date))
}

fn python_naive_datetime_text(datetime: chrono::NaiveDateTime) -> String {
    let mut text = datetime.format("%Y-%m-%d %H:%M:%S").to_string();
    let micros = datetime.and_utc().timestamp_subsec_micros();
    if micros != 0 {
        text.push('.');
        text.push_str(&format!("{micros:06}"));
    }
    text
}

fn python_time_text(time: chrono::NaiveTime) -> String {
    use chrono::Timelike;

    let mut text = time.format("%H:%M:%S").to_string();
    let micros = time.nanosecond() / 1_000;
    if micros != 0 {
        text.push('.');
        text.push_str(&format!("{micros:06}"));
    }
    text
}
fn disambiguate(names: &[String]) -> Vec<String> {
    let mut counts = std::collections::HashMap::new();
    names
        .iter()
        .map(|name| {
            let count = counts.entry(name.clone()).or_insert(0usize);
            *count += 1;
            if *count == 1 {
                name.clone()
            } else {
                format!("{name}_{}", count)
            }
        })
        .collect()
}

fn comment_summary_sheet_names(names: &[String]) -> Vec<String> {
    let mut used = BTreeSet::new();
    names
        .iter()
        .map(|name| {
            let mut sanitized = name
                .trim()
                .chars()
                .map(|character| {
                    if matches!(character, ':' | '\\' | '/' | '?' | '*' | '[' | ']') {
                        '_'
                    } else {
                        character
                    }
                })
                .take(31)
                .collect::<String>();
            if sanitized.is_empty() {
                sanitized = "Sheet".to_owned();
            }
            if used.insert(sanitized.clone()) {
                return sanitized;
            }

            let base = sanitized.chars().take(28).collect::<String>();
            for suffix in 2usize.. {
                let candidate = format!("{base}_{suffix}")
                    .chars()
                    .take(31)
                    .collect::<String>();
                if used.insert(candidate.clone()) {
                    return candidate;
                }
            }
            unreachable!("worksheet names are bounded by Excel's 31-character limit")
        })
        .collect()
}

#[cfg(test)]
mod patch_error_cells_tests {
    use super::*;

    #[test]
    fn failed_patch_does_not_remove_a_preexisting_output() -> Result<()> {
        let unique = SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos();
        let root = std::env::temp_dir().join(format!(
            "pytools-summary-patch-{}-{unique}",
            std::process::id()
        ));
        fs::create_dir(&root)?;
        let source = root.join("invalid.xlsx");
        let output = root.join("existing.xlsx");
        fs::write(&source, b"not an xlsx archive")?;
        fs::write(&output, b"existing output")?;
        let errors = [PendingErrorCell {
            address: "A1".to_owned(),
            reader_value: "Div0".to_owned(),
            excel_value: "#DIV/0!",
        }];

        let result = patch_error_cells(&source, &output, &errors);
        let output_bytes = fs::read(&output);
        fs::remove_dir_all(&root)?;

        assert!(result.is_err());
        assert_eq!(output_bytes?, b"existing output");
        Ok(())
    }
}

#[cfg(test)]
mod comment_region_tests {
    use super::*;

    #[test]
    fn empty_set_capture_is_ignored_but_whitespace_capture_is_retained() {
        assert_eq!(parse_set_name("set()"), None);
        assert_eq!(parse_set_name("set( )"), Some(String::new()));
        assert_eq!(parse_set_name("SET（区域）"), Some("区域".to_owned()));
    }

    #[test]
    fn duplicate_region_markers_use_last_row_major_boundary() {
        let comments = BTreeMap::from([
            ("A2".to_owned(), "行区域1".to_owned()),
            ("A3".to_owned(), "行区域1".to_owned()),
            ("A4".to_owned(), "行区域#1".to_owned()),
            ("A5".to_owned(), "行区域#1".to_owned()),
        ]);

        assert_eq!(
            parse_comment_ranges(&comments, "行区域"),
            vec![(3, 5, 1, 1)]
        );
    }

    #[test]
    fn template_header_text_matches_python_for_empty_and_boolean_values() {
        assert_eq!(probe_text(&serde_json::Value::Null), "");
        assert_eq!(probe_text(&serde_json::json!(true)), "True");
        assert_eq!(probe_text(&serde_json::json!(false)), "False");
        assert_eq!(
            probe_text(&serde_json::json!({"type":"number", "value":10.0})),
            "10"
        );
        assert_eq!(
            probe_text(&serde_json::json!({"type":"number", "value":10.5})),
            "10.5"
        );
        assert_eq!(
            probe_text(&serde_json::json!({"type":"number", "value":1e-5})),
            "1e-05"
        );
        assert_eq!(
            probe_text(&serde_json::json!({"type":"number", "value":1e-7})),
            "1e-07"
        );
        assert_eq!(
            probe_text(&serde_json::json!({"type":"number", "value":1e-4})),
            "0.0001"
        );
        assert_eq!(
            probe_text(&serde_json::json!({"type":"number", "value":1e20})),
            "1e+20"
        );
        assert_eq!(
            probe_text(&serde_json::json!({
                "type":"datetime_iso",
                "value":"2026-06-30T00:00:00.123"
            })),
            "2026-06-30 00:00:00.123000"
        );
        assert_eq!(
            probe_text(&serde_json::json!({"type":"datetime_serial", "value":46203.0})),
            "2026-06-30 00:00:00"
        );
    }

    #[test]
    fn comment_summary_output_sheet_names_match_python_normalization() {
        let names = vec![
            "  Sales  ".to_owned(),
            "Sales".to_owned(),
            "Regional/summary".to_owned(),
            "   ".to_owned(),
            "A".repeat(32),
            "A".repeat(32),
        ];

        assert_eq!(
            comment_summary_sheet_names(&names),
            vec![
                "Sales".to_owned(),
                "Sales_2".to_owned(),
                "Regional_summary".to_owned(),
                "Sheet".to_owned(),
                "A".repeat(31),
                format!("{}_2", "A".repeat(28)),
            ]
        );
    }
}

#[cfg(test)]
mod merged_cell_tests {
    use super::*;

    #[test]
    fn merged_range_non_anchor_cells_are_omitted() -> Result<()> {
        let merges = parse_merged_ranges(&["$A$1:$AG$1".to_owned(), "C3:D4".to_owned()])?;

        assert!(!is_merged_non_anchor(&merges, 1, 1));
        assert!(is_merged_non_anchor(&merges, 1, 2));
        assert!(is_merged_non_anchor(&merges, 1, 33));
        assert!(!is_merged_non_anchor(&merges, 1, 34));
        assert!(!is_merged_non_anchor(&merges, 3, 3));
        assert!(is_merged_non_anchor(&merges, 3, 4));
        assert!(is_merged_non_anchor(&merges, 4, 3));
        assert!(is_merged_non_anchor(&merges, 4, 4));
        Ok(())
    }
}
