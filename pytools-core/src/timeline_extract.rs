//! Rust implementation of the compact timeline extraction workflow (PT-1-2).

use crate::{
    config::{GlobalConfig, PathMapRule, TimelineRule},
    rule_engine::{extract_cells_from_matrix, pick_rules_for_workbook, ExtractedCell},
    workbook_extract::{matrix_from_dump, WorkbookCell},
};
use anyhow::{anyhow, Context, Result};
use chrono::{Local, NaiveDate, NaiveDateTime};
use excel_reader::read_workbook;
use excel_writer::{Format, Workbook, Worksheet};
use rayon::prelude::*;
use serde::Serialize;
use std::{
    collections::HashSet,
    fs::{self, OpenOptions},
    path::{Path, PathBuf},
    sync::mpsc,
    time::{Duration, Instant},
};

const OUTPUT_SHEET: &str = "时序提取结果";
const OUTPUT_HEADERS: [&str; 7] = [
    "源文件",
    "工作表名",
    "规则名称",
    "数据日期",
    "行头路径",
    "列头路径",
    "数值",
];

#[derive(Clone, Debug, Default, Serialize)]
pub struct TimelineExtractProgress {
    pub total_inputs: usize,
    pub processed_inputs: usize,
    pub completed_inputs: usize,
    pub failed_inputs: usize,
    pub current_input: Option<usize>,
    pub matched_rule_sheets: usize,
    pub extracted_rows: usize,
    pub output_rows: usize,
}

#[derive(Clone, Debug, Serialize)]
pub struct TimelineExtractRunEvent {
    pub sequence: usize,
    pub timestamp: String,
    pub feature_id: &'static str,
    pub phase: String,
    pub status: String,
    pub user_message: String,
    pub progress: TimelineExtractProgress,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub input_index: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stage_elapsed_ms: Option<u128>,
    pub total_elapsed_ms: u128,
    pub components: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub actual_engine: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error_type: Option<String>,
}

#[derive(Clone, Debug, Serialize)]
pub struct TimelineExtractResult {
    pub total_files: usize,
    pub processed_files: usize,
    pub files: usize,
    pub failed_files: usize,
    pub matched_rule_sheets: usize,
    pub rows: usize,
    pub configured_target_rules: usize,
    pub target_writes_deferred: bool,
    pub path: Option<PathBuf>,
    pub read_elapsed_ms: u128,
    pub write_elapsed_ms: u128,
    pub total_elapsed_ms: u128,
}

struct FileExtraction {
    cells: Vec<ExtractedCell<WorkbookCell>>,
    matched_rule_sheets: usize,
    elapsed: Duration,
}

struct EventEmitter<'a> {
    started: Instant,
    sequence: usize,
    progress: TimelineExtractProgress,
    callback: Option<&'a mut dyn FnMut(TimelineExtractRunEvent)>,
}

impl EventEmitter<'_> {
    fn emit(
        &mut self,
        phase: &str,
        status: &str,
        message: impl Into<String>,
        input_index: Option<usize>,
        elapsed: Option<Duration>,
        components: &[&str],
        engine: Option<&str>,
        error_type: Option<&str>,
    ) {
        let Some(callback) = self.callback.as_mut() else {
            return;
        };
        self.sequence += 1;
        let mut recorded_components = vec!["Rust".to_owned()];
        recorded_components.extend(components.iter().map(|component| (*component).to_owned()));
        callback(TimelineExtractRunEvent {
            sequence: self.sequence,
            timestamp: Local::now().to_rfc3339(),
            feature_id: "PT-1-2",
            phase: phase.to_owned(),
            status: status.to_owned(),
            user_message: message.into(),
            progress: self.progress.clone(),
            input_index,
            stage_elapsed_ms: elapsed.map(|value| value.as_millis()),
            total_elapsed_ms: self.started.elapsed().as_millis(),
            components: recorded_components,
            actual_engine: engine.map(str::to_owned),
            error_type: error_type.map(str::to_owned),
        });
    }
}

fn record_file_extraction(
    emitter: &mut EventEmitter<'_>,
    index: usize,
    total: usize,
    result: &Result<FileExtraction>,
) {
    emitter.progress.processed_inputs += 1;
    emitter.progress.current_input = Some(index + 1);
    match result {
        Ok(file) => {
            emitter.progress.completed_inputs += 1;
            emitter.progress.matched_rule_sheets += file.matched_rule_sheets;
            emitter.progress.extracted_rows += file.cells.len();
            emitter.emit(
                "workbook_extract",
                "completed",
                format!(
                    "已处理 {}/{} 个工作簿，匹配 {} 个工作表",
                    emitter.progress.processed_inputs, total, file.matched_rule_sheets
                ),
                Some(index + 1),
                Some(file.elapsed),
                &["pytools-core", "excel-reader", "calamine", "rule_engine"],
                Some("calamine"),
                None,
            );
        }
        Err(_) => {
            emitter.progress.failed_inputs += 1;
            emitter.emit(
                "workbook_extract",
                "failed",
                format!("第 {}/{} 个工作簿读取或抽取失败", index + 1, total),
                Some(index + 1),
                None,
                &["pytools-core", "excel-reader", "calamine", "rule_engine"],
                Some("calamine"),
                Some("workbook_extract_error"),
            );
        }
    }
}

/// Extract configured row/column paths from supported source workbooks and
/// write the compact seven-column result used by the original `run_timeline_slim`.
pub fn run_timeline_extract(
    source_paths: &[PathBuf],
    rules: &[TimelineRule],
    path_maps: &[PathMapRule],
    global: &GlobalConfig,
) -> Result<TimelineExtractResult> {
    run_timeline_extract_inner(source_paths, rules, path_maps, global, None)
}

pub fn run_timeline_extract_with_progress<F>(
    source_paths: &[PathBuf],
    rules: &[TimelineRule],
    path_maps: &[PathMapRule],
    global: &GlobalConfig,
    mut on_event: F,
) -> Result<TimelineExtractResult>
where
    F: FnMut(TimelineExtractRunEvent),
{
    run_timeline_extract_inner(source_paths, rules, path_maps, global, Some(&mut on_event))
}

fn run_timeline_extract_inner(
    source_paths: &[PathBuf],
    rules: &[TimelineRule],
    path_maps: &[PathMapRule],
    global: &GlobalConfig,
    callback: Option<&mut dyn FnMut(TimelineExtractRunEvent)>,
) -> Result<TimelineExtractResult> {
    let started = Instant::now();
    let supported_extensions = global
        .source_exts
        .iter()
        .map(|value| normalize_extension(value))
        .chain(std::iter::once(".xls".to_owned()))
        .collect::<HashSet<_>>();
    let sources = source_paths
        .iter()
        .filter(|path| {
            path.is_file()
                && path
                    .extension()
                    .map(|extension| {
                        supported_extensions
                            .contains(&format!(".{}", extension.to_string_lossy().to_lowercase()))
                    })
                    .unwrap_or(false)
        })
        .cloned()
        .collect::<Vec<_>>();
    let mut emitter = EventEmitter {
        started,
        sequence: 0,
        progress: TimelineExtractProgress {
            total_inputs: sources.len(),
            ..TimelineExtractProgress::default()
        },
        callback,
    };
    emitter.emit(
        "run",
        "started",
        format!(
            "开始时序提取：{} 个输入，{} 条启用规则",
            sources.len(),
            rules.iter().filter(|rule| rule.enabled).count()
        ),
        None,
        None,
        &["pytools-core"],
        None,
        None,
    );

    let read_started = Instant::now();
    let source_count = sources.len();
    let mut results = (0..sources.len())
        .map(|_| None)
        .collect::<Vec<Option<Result<FileExtraction>>>>();
    if global.error_policy == "fail_fast" {
        for (index, path) in sources.iter().enumerate() {
            let result = extract_one_file(path, rules, path_maps, &global.timeline_rule_match_mode);
            record_file_extraction(&mut emitter, index, source_count, &result);
            let failed = result.is_err();
            results[index] = Some(result);
            if failed {
                break;
            }
        }
    } else {
        let (sender, receiver) = mpsc::channel::<(usize, Result<FileExtraction>)>();
        std::thread::scope(|scope| {
            let worker_sender = sender.clone();
            let sources_ref = &sources;
            let rules_ref = rules;
            let path_maps_ref = path_maps;
            let rule_match_mode = global.timeline_rule_match_mode.as_str();
            scope.spawn(move || {
                sources_ref
                    .par_iter()
                    .enumerate()
                    .for_each(|(index, path)| {
                        let result =
                            extract_one_file(path, rules_ref, path_maps_ref, rule_match_mode);
                        let _ = worker_sender.send((index, result));
                    });
            });
            drop(sender);
            for (index, result) in receiver {
                record_file_extraction(&mut emitter, index, source_count, &result);
                results[index] = Some(result);
            }
        });
    }
    let read_elapsed = read_started.elapsed();

    let mut cells = Vec::new();
    let mut matched_rule_sheets = 0usize;
    let mut first_error = None;
    let mut seen = HashSet::<(String, String, String, String, String, String)>::new();
    for (index, result) in results.into_iter().enumerate() {
        let Some(result) = result else {
            first_error.get_or_insert_with(|| {
                anyhow!("workbook extraction worker did not return a result")
            });
            continue;
        };
        match result {
            Ok(file) => {
                matched_rule_sheets += file.matched_rule_sheets;
                for cell in file.cells {
                    let key = (
                        cell.source_file.clone(),
                        cell.sheet_name.clone(),
                        cell.data_date.clone(),
                        cell.row_path.clone(),
                        cell.col_path.clone(),
                        python_value_text(&cell.value),
                    );
                    if seen.insert(key) {
                        cells.push(cell);
                    }
                }
            }
            Err(error) => {
                first_error.get_or_insert_with(|| {
                    error.context(format!(
                        "process source workbook {}",
                        sources[index].display()
                    ))
                });
            }
        }
    }
    if global.error_policy == "fail_fast" {
        if let Some(error) = first_error {
            emitter.emit(
                "run",
                "failed",
                "时序提取失败",
                None,
                Some(read_elapsed),
                &["pytools-core"],
                None,
                Some("workbook_extract_error"),
            );
            return Err(error);
        }
    }

    let write_started = Instant::now();
    let path = if cells.is_empty() {
        None
    } else {
        emitter.progress.output_rows = cells.len();
        emitter.emit(
            "workbook_write",
            "started",
            "正在写入时序提取结果",
            None,
            None,
            &["pytools-core", "excel-writer", "rust_xlsxwriter"],
            None,
            None,
        );
        Some(write_output_workbook(&cells, &global.output_dir)?)
    };
    let write_elapsed = write_started.elapsed();
    emitter.emit(
        "workbook_write",
        if path.is_some() {
            "completed"
        } else {
            "skipped"
        },
        if path.is_some() {
            format!("时序提取完成，输出 {} 行", cells.len())
        } else {
            "无匹配结果，不生成输出工作簿".to_owned()
        },
        None,
        Some(write_elapsed),
        &["pytools-core", "excel-writer", "rust_xlsxwriter"],
        path.as_ref().map(|_| "rust_xlsxwriter"),
        None,
    );
    let output_rule_names = cells
        .iter()
        .map(|cell| cell.rule_name.as_str())
        .collect::<HashSet<_>>();
    let configured_target_rules = rules
        .iter()
        .filter(|rule| {
            rule.target_write_enabled
                && rule.target_wb_path.is_some()
                && rule
                    .target_sheet
                    .as_ref()
                    .is_some_and(|sheet| !sheet.trim().is_empty())
                && output_rule_names.contains(rule.name.as_str())
        })
        .count();
    let target_writes_deferred = configured_target_rules > 0 && !cells.is_empty();
    let partial = emitter.progress.failed_inputs > 0 || target_writes_deferred;
    let end_message = if target_writes_deferred {
        format!(
            "时序提取生成 {} 行；{} 条目标写入规则尚待 PT-1-2 目标簿语义验收",
            cells.len(),
            configured_target_rules
        )
    } else if emitter.progress.failed_inputs > 0 {
        format!(
            "时序提取部分完成：{} 个工作簿失败，仍输出 {} 行",
            emitter.progress.failed_inputs,
            cells.len()
        )
    } else {
        format!("时序提取结束：输出 {} 行", cells.len())
    };
    emitter.emit(
        "run",
        if partial { "partial" } else { "completed" },
        end_message,
        None,
        Some(started.elapsed()),
        &["pytools-core"],
        None,
        None,
    );
    Ok(TimelineExtractResult {
        total_files: emitter.progress.total_inputs,
        processed_files: emitter.progress.processed_inputs,
        files: emitter.progress.completed_inputs,
        failed_files: emitter.progress.failed_inputs,
        matched_rule_sheets,
        rows: cells.len(),
        configured_target_rules,
        target_writes_deferred,
        path,
        read_elapsed_ms: read_elapsed.as_millis(),
        write_elapsed_ms: write_elapsed.as_millis(),
        total_elapsed_ms: started.elapsed().as_millis(),
    })
}

fn extract_one_file(
    path: &Path,
    rules: &[TimelineRule],
    path_maps: &[PathMapRule],
    rule_match_mode: &str,
) -> Result<FileExtraction> {
    let started = Instant::now();
    let filename = path
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or_else(|| anyhow!("source workbook has no valid file name"))?;
    let matched_rules = pick_rules_for_workbook(rules, filename, rule_match_mode);
    if matched_rules.is_empty() {
        return Ok(FileExtraction {
            cells: Vec::new(),
            matched_rule_sheets: 0,
            elapsed: started.elapsed(),
        });
    }
    let workbook =
        read_workbook(path).with_context(|| format!("read source workbook {}", path.display()))?;
    let mut cells = Vec::new();
    let mut matched_rule_sheets = 0usize;
    for rule in matched_rules {
        for sheet in &workbook.sheets {
            if !crate::path_map::match_all_keywords(&sheet.name, &rule.sheet_keyword) {
                continue;
            }
            let (mut matrix, merges) = matrix_from_dump(sheet)?;
            clear_non_anchor_merge_values(&mut matrix, &merges);
            for row in &mut matrix {
                for value in row {
                    if matches!(value, Some(WorkbookCell::Error(_))) {
                        *value = None;
                    }
                }
            }
            let extracted = extract_cells_from_matrix(
                &matrix,
                &merges,
                filename,
                rule,
                &sheet.name,
                path_maps,
                true,
                &Default::default(),
            )
            .map_err(|error| anyhow!(error))
            .with_context(|| {
                format!(
                    "extract rule {:?} from worksheet {:?}",
                    rule.name, sheet.name
                )
            })?;
            if !extracted.is_empty() {
                matched_rule_sheets += 1;
                cells.extend(extracted);
            }
        }
    }
    Ok(FileExtraction {
        cells,
        matched_rule_sheets,
        elapsed: started.elapsed(),
    })
}

fn clear_non_anchor_merge_values(
    matrix: &mut [Vec<Option<WorkbookCell>>],
    merges: &[crate::rule_engine::MergeRange],
) {
    for merge in merges {
        for row in merge.min_row..=merge.max_row {
            for column in merge.min_col..=merge.max_col {
                if row == merge.min_row && column == merge.min_col {
                    continue;
                }
                if let Some(cell) = matrix.get_mut(row).and_then(|cells| cells.get_mut(column)) {
                    *cell = None;
                }
            }
        }
    }
}

fn normalize_extension(extension: &str) -> String {
    let extension = extension.trim().to_lowercase();
    if extension.starts_with('.') {
        extension
    } else {
        format!(".{extension}")
    }
}

fn python_value_text(value: &WorkbookCell) -> String {
    match value {
        WorkbookCell::Text(value)
        | WorkbookCell::Error(value)
        | WorkbookCell::DurationIso(value) => value.clone(),
        WorkbookCell::Integer(value) => value.to_string(),
        WorkbookCell::Number(value) => crate::summary_ops::python_float_text(*value),
        WorkbookCell::Boolean(value) => if *value { "True" } else { "False" }.to_owned(),
        WorkbookCell::DateTimeSerial(serial) => {
            crate::summary_ops::python_datetime_serial_text(*serial)
                .unwrap_or_else(|| crate::summary_ops::python_float_text(*serial))
        }
        WorkbookCell::DateTimeIso(value) => crate::summary_ops::python_datetime_iso_text(value),
    }
}

fn write_output_workbook(
    rows: &[ExtractedCell<WorkbookCell>],
    output_dir: &Path,
) -> Result<PathBuf> {
    fs::create_dir_all(output_dir)
        .with_context(|| format!("create timeline output directory {}", output_dir.display()))?;
    let output_path = reserve_output_path(output_dir)?;
    let write_result = (|| -> Result<()> {
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();
        worksheet.set_name(OUTPUT_SHEET)?;
        let header_format = Format::new().set_bold();
        for (column, header) in OUTPUT_HEADERS.iter().enumerate() {
            worksheet.write_string_with_format(0, column as u16, *header, &header_format)?;
        }
        for (index, cell) in rows.iter().enumerate() {
            let row = u32::try_from(index + 1).context("timeline output exceeds XLSX row limit")?;
            worksheet.write_string(row, 0, &cell.source_file)?;
            worksheet.write_string(row, 1, &cell.sheet_name)?;
            worksheet.write_string(row, 2, &cell.rule_name)?;
            write_data_date(worksheet, row, 3, &cell.data_date)?;
            worksheet.write_string(row, 4, &cell.row_path)?;
            worksheet.write_string(row, 5, &cell.col_path)?;
            write_value(worksheet, row, 6, &cell.value)?;
        }
        workbook.save(&output_path).with_context(|| {
            format!("save timeline extraction output {}", output_path.display())
        })?;
        Ok(())
    })();
    if let Err(error) = write_result {
        let _ = fs::remove_file(&output_path);
        return Err(error);
    }
    Ok(output_path)
}

fn reserve_output_path(output_dir: &Path) -> Result<PathBuf> {
    let timestamp = Local::now().format("%Y%m%d_%H%M%S");
    for suffix in 1..=10_000usize {
        let label = if suffix == 1 {
            String::new()
        } else {
            format!("_{}", suffix)
        };
        let candidate = output_dir.join(format!("{timestamp}_时序提取（快速精简）{label}.xlsx"));
        match OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&candidate)
        {
            Ok(_) => return Ok(candidate),
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(error) => {
                return Err(error).with_context(|| {
                    format!("reserve timeline output path {}", candidate.display())
                });
            }
        }
    }
    Err(anyhow!("could not reserve a unique timeline output path"))
}

fn write_data_date(worksheet: &mut Worksheet, row: u32, column: u16, value: &str) -> Result<()> {
    let format = Format::new().set_num_format("yyyy/m/d");
    if let Some(serial) = date_to_excel_serial(value) {
        worksheet.write_number_with_format(row, column, serial, &format)?;
    } else if value.is_empty() {
        worksheet.write_blank(row, column, &format)?;
    } else {
        worksheet.write_string_with_format(row, column, value, &format)?;
    }
    Ok(())
}

fn write_value(
    worksheet: &mut Worksheet,
    row: u32,
    column: u16,
    value: &WorkbookCell,
) -> Result<()> {
    match value {
        WorkbookCell::Text(value) | WorkbookCell::Error(value) => {
            worksheet.write_string(row, column, value)?;
        }
        WorkbookCell::Integer(value) => {
            worksheet.write_number(row, column, *value as f64)?;
        }
        WorkbookCell::Number(value) => {
            worksheet.write_number(row, column, *value)?;
        }
        WorkbookCell::Boolean(value) => {
            worksheet.write_boolean(row, column, *value)?;
        }
        WorkbookCell::DateTimeSerial(value) => {
            let format = Format::new().set_num_format("YYYY-MM-DD HH:MM:SS");
            worksheet.write_number_with_format(row, column, *value, &format)?;
        }
        WorkbookCell::DateTimeIso(value) => {
            if let Some(serial) = datetime_to_excel_serial(value) {
                let format = Format::new().set_num_format("YYYY-MM-DD HH:MM:SS");
                worksheet.write_number_with_format(row, column, serial, &format)?;
            } else {
                worksheet.write_string(row, column, value)?;
            }
        }
        WorkbookCell::DurationIso(value) => {
            if let Some(serial) = iso_duration_serial(value) {
                let format = Format::new().set_num_format("[h]:mm:ss.000");
                worksheet.write_number_with_format(row, column, serial, &format)?;
            } else {
                worksheet.write_string(row, column, value)?;
            }
        }
    }
    Ok(())
}

fn date_to_excel_serial(value: &str) -> Option<f64> {
    let date = NaiveDate::parse_from_str(value, "%Y-%m-%d").ok()?;
    let epoch = NaiveDate::from_ymd_opt(1899, 12, 30)?;
    Some(date.signed_duration_since(epoch).num_days() as f64)
}

fn datetime_to_excel_serial(value: &str) -> Option<f64> {
    let datetime = chrono::DateTime::parse_from_rfc3339(value)
        .map(|value| value.naive_utc())
        .or_else(|_| NaiveDateTime::parse_from_str(value, "%Y-%m-%dT%H:%M:%S%.f"))
        .or_else(|_| NaiveDateTime::parse_from_str(value, "%Y-%m-%d %H:%M:%S%.f"))
        .or_else(|_| {
            NaiveDate::parse_from_str(value, "%Y-%m-%d")
                .map(|date| date.and_hms_opt(0, 0, 0).expect("valid midnight"))
        })
        .ok()?;
    let epoch = NaiveDate::from_ymd_opt(1899, 12, 30)?.and_hms_opt(0, 0, 0)?;
    Some(datetime.signed_duration_since(epoch).num_milliseconds() as f64 / 86_400_000.0)
}

fn iso_duration_serial(value: &str) -> Option<f64> {
    let text = value.strip_prefix('P')?;
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
        let amount = number.parse::<f64>().ok()?;
        number.clear();
        let multiplier = match (in_time, character) {
            (false, 'D') => 86_400.0,
            (true, 'H') => 3_600.0,
            (true, 'M') => 60.0,
            (true, 'S') => 1.0,
            _ => return None,
        };
        seconds += amount * multiplier;
    }
    if number.is_empty() {
        Some(seconds / 86_400.0)
    } else {
        None
    }
}
