//! S5 external-workbook layout discovery for the Python `external_layout` path.
//!
//! Ordinary and date-keyed external report layouts, institution mapping, and
//! source-layout quality warnings for the S5 V_EXT retrieval path.

use crate::s5_scan::ScanBook;
use excel_reader::{read_value_columns_with_iso_dates, CellValue};
use std::collections::{BTreeMap, BTreeSet};

const MAPPING_SHEETS: [&str; 2] = ["机构参照表", "机构参照"];
const IDENTITY_LABELS: [&str; 8] = [
    "统一社会信用代码",
    "法人机构统一社会信用代码",
    "统一信用代码",
    "机构代码",
    "机构编码",
    "机构编号",
    "金融机构代码",
    "金融机构编码",
];
const NAME_LABELS: [&str; 4] = ["法人机构名称", "机构名称", "机构全称", "法人名称"];
const TARGET_LABELS: [&str; 5] = ["报表项目", "报表名称", "项目名称", "项目编码", "行项目"];
const PLACEHOLDER_HEADERS: [&str; 7] = ["null", "none", "nan", "n/a", "#n/a", "-", "--"];
const LEGACY_MAPPING_SHEET: &str = "参照表";
const EXTERNAL_METADATA_HEADERS: [&str; 8] = [
    "数据日期",
    "报表名称",
    "报表代码",
    "版本号",
    "锁定信息",
    "报表组",
    "行顺序号",
    "行名称",
];

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ExternalLayoutFailure {
    pub code: String,
    pub message: String,
}

impl ExternalLayoutFailure {
    fn new(code: &str, message: impl Into<String>) -> Self {
        Self {
            code: code.to_owned(),
            message: message.into(),
        }
    }
}

#[derive(Clone, Debug)]
struct InstitutionMapping {
    code: String,
    name: String,
    report_item: String,
    disabled: bool,
    row: usize,
    code_column: usize,
    name_column: usize,
    target_column: usize,
}

#[derive(Clone, Debug)]
struct ExternalColumn {
    sheet: String,
    column: usize,
    header: String,
    rows: BTreeMap<String, Vec<usize>>,
    data_date_column: Option<usize>,
    row_dates: BTreeMap<usize, String>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ExternalLayoutWarning {
    pub sheet: String,
    pub location: String,
    pub message: String,
}

#[derive(Clone, Debug)]
pub struct ExternalLookup {
    pub mapping_sheet: String,
    pub mapping_cell: String,
    pub mapping_key: String,
    pub mapping_method: String,
    pub report_item: String,
    pub sheet: String,
    pub row: usize,
    pub column: usize,
}

#[derive(Clone, Debug)]
pub struct ExternalLayout {
    mapping_sheet: String,
    mapping_rows: Vec<InstitutionMapping>,
    columns: Vec<ExternalColumn>,
    pub warnings: Vec<ExternalLayoutWarning>,
}

pub fn normalize_external_label(value: &str) -> String {
    let compact: String = value
        .chars()
        .filter(|character| !character.is_whitespace())
        .collect();
    compact
        .trim_end_matches(|character| character == ':' || character == '：')
        .to_lowercase()
}

impl ExternalLayout {
    pub fn has_time_series_header(book: &ScanBook) -> bool {
        book.sheet_names().iter().any(|sheet| {
            !MAPPING_SHEETS.contains(&sheet.as_str())
                && sheet != LEGACY_MAPPING_SHEET
                && time_series_header(book, sheet).is_some()
        })
    }

    pub fn metric_header_match_count(&self, metric_header: &str, sheet_filter: &str) -> usize {
        let metric = normalize_external_label(metric_header);
        self.columns
            .iter()
            .filter(|column| {
                normalize_external_label(&column.header) == metric
                    && (sheet_filter.is_empty() || column.sheet == sheet_filter)
            })
            .count()
    }

    pub fn discover(book: &ScanBook) -> Result<Self, ExternalLayoutFailure> {
        let (mapping_sheet, mapping_rows) = discover_institution_mapping(book)?;
        let targets: BTreeSet<String> = mapping_rows
            .iter()
            .filter(|entry| !entry.report_item.is_empty())
            .map(|entry| normalize_external_label(&entry.report_item))
            .collect();
        if targets.is_empty() {
            return Err(ExternalLayoutFailure::new(
                "MAPPING_SHEET_MISSING",
                "外部文件中未找到包含机构代码和报表项目的映射表",
            ));
        }

        let mapping_sheets: BTreeSet<&str> = MAPPING_SHEETS.into_iter().collect();
        let time_series_headers: BTreeMap<String, (usize, usize, usize)> = book
            .sheet_names()
            .iter()
            .filter(|sheet| {
                !mapping_sheets.contains(sheet.as_str()) && sheet.as_str() != LEGACY_MAPPING_SHEET
            })
            .filter_map(|sheet| {
                time_series_header(book, sheet).map(|header| (sheet.clone(), header))
            })
            .collect();
        let iso_dates = if time_series_headers.is_empty() {
            BTreeMap::new()
        } else {
            read_value_columns_with_iso_dates(book.path(), usize::MAX)
                .map_err(|error| {
                    ExternalLayoutFailure::new(
                        "EXTERNAL_DATE_READ_FAILED",
                        format!("无法读取外部时序数据的 Excel 日期值：{error}"),
                    )
                })?
                .into_iter()
                .map(|sheet| {
                    let values = sheet
                        .values
                        .into_iter()
                        .filter_map(|(address, value)| match value {
                            CellValue::DateTimeIso(value) => Some((address, value)),
                            _ => None,
                        })
                        .collect();
                    (sheet.name, values)
                })
                .collect()
        };
        let mut columns = Vec::new();
        let mut warnings = Vec::new();
        for sheet in book.sheet_names() {
            if mapping_sheets.contains(sheet.as_str()) || sheet == LEGACY_MAPPING_SHEET {
                continue;
            }
            let mut key_rows: BTreeMap<usize, BTreeMap<String, Vec<usize>>> = BTreeMap::new();
            let time_series = time_series_headers.get(sheet).copied();
            let mut explicit_header_rows = BTreeMap::new();
            if let Some((header_row, _date_column, key_column)) = time_series {
                explicit_header_rows.insert(key_column, header_row);
                let rows = key_rows.entry(key_column).or_default();
                for row in header_row + 1..=book.sheet_max_row(sheet) {
                    if book.sheet_has_formula(sheet, row, key_column) {
                        continue;
                    }
                    let (kind, raw) = book.stored_value(sheet, row, key_column);
                    if kind == "string" {
                        let label = normalize_external_label(&raw);
                        if targets.contains(&label) {
                            rows.entry(label).or_default().push(row);
                        }
                    }
                }
            } else {
                for row in 1..=book.sheet_max_row(sheet) {
                    for column in 1..=book.sheet_max_col(sheet) {
                        let (kind, raw) = book.stored_value(sheet, row, column);
                        if kind != "string" || book.sheet_has_formula(sheet, row, column) {
                            continue;
                        }
                        let label = normalize_external_label(&raw);
                        if targets.contains(&label) {
                            key_rows
                                .entry(column)
                                .or_default()
                                .entry(label)
                                .or_default()
                                .push(row);
                        }
                    }
                }

                for row in 1..=book.sheet_max_row(sheet).min(30) {
                    let headers: Vec<String> = (1..=book.sheet_max_col(sheet))
                        .map(|column| {
                            normalize_external_label(&book.cell_raw_text(sheet, row, column))
                        })
                        .collect();
                    let label_columns: Vec<usize> = headers
                        .iter()
                        .enumerate()
                        .filter_map(|(index, header)| {
                            ["行名称", "报表项目"]
                                .contains(&header.as_str())
                                .then_some(index + 1)
                        })
                        .collect();
                    if label_columns.len() == 1 {
                        let key_column = label_columns[0];
                        if headers
                            .iter()
                            .skip(key_column)
                            .any(|header| !header.is_empty())
                        {
                            explicit_header_rows.entry(key_column).or_insert(row);
                            key_rows.entry(key_column).or_default();
                        }
                    }
                }
            }

            let key_columns: BTreeSet<usize> = key_rows
                .keys()
                .chain(explicit_header_rows.keys())
                .copied()
                .collect();
            for key_column in key_columns {
                let item_rows = key_rows.get(&key_column).cloned().unwrap_or_default();
                let first_data_row = item_rows
                    .values()
                    .flat_map(|rows| rows.iter().copied())
                    .min()
                    .unwrap_or_else(|| {
                        explicit_header_rows.get(&key_column).copied().unwrap_or(1) + 1
                    });
                let candidates: Vec<usize> = (1..first_data_row.min(31))
                    .filter(|row| {
                        ((key_column + 1)..=book.sheet_max_col(sheet)).any(|column| {
                            let (kind, raw) = book.stored_value(sheet, *row, column);
                            kind == "string"
                                && !book.sheet_has_formula(sheet, *row, column)
                                && !normalize_external_label(&raw).is_empty()
                        })
                    })
                    .collect();
                let header_row = if let Some(row) = explicit_header_rows.get(&key_column) {
                    *row
                } else if let Some(row) = candidates.last() {
                    *row
                } else {
                    continue;
                };

                let date_column = time_series.map(|(_, date_column, _)| date_column);
                let mut row_dates = BTreeMap::new();
                if let Some(date_column) = date_column {
                    for row in item_rows.values().flatten() {
                        row_dates.insert(
                            *row,
                            external_date_value(book, &iso_dates, sheet, *row, date_column),
                        );
                    }
                }

                let mut metric_columns = Vec::new();
                for column in (key_column + 1)..=book.sheet_max_col(sheet) {
                    if book.sheet_has_formula(sheet, header_row, column) {
                        continue;
                    }
                    let raw = book.merged_raw_text(sheet, header_row, column);
                    if raw.trim().is_empty() || raw.starts_with('=') {
                        continue;
                    }
                    let normalized = normalize_external_label(&raw);
                    if date_column.is_some()
                        && EXTERNAL_METADATA_HEADERS.contains(&normalized.as_str())
                    {
                        continue;
                    }
                    if PLACEHOLDER_HEADERS.contains(&normalized.as_str()) {
                        continue;
                    }
                    let header = raw.trim().to_owned();
                    metric_columns.push((column, header.clone()));
                    columns.push(ExternalColumn {
                        sheet: sheet.clone(),
                        column,
                        header,
                        rows: item_rows.clone(),
                        data_date_column: date_column,
                        row_dates: row_dates.clone(),
                    });
                }
                if metric_columns.is_empty() {
                    continue;
                }

                let data_start = header_row + 1;
                let mut row_has_values = BTreeMap::new();
                let mut last_nonempty_row = data_start.saturating_sub(1);
                for row in data_start..=book.sheet_max_row(sheet) {
                    let present = (1..=book.sheet_max_col(sheet))
                        .any(|column| has_value(book, sheet, row, column));
                    row_has_values.insert(row, present);
                    if present {
                        last_nonempty_row = row;
                    }
                }
                for row in data_start..=last_nonempty_row {
                    let raw_label = book.cell_raw_text(sheet, row, key_column);
                    let label = normalize_external_label(&raw_label);
                    let populated_metrics: Vec<&str> = metric_columns
                        .iter()
                        .filter_map(|(column, title)| {
                            has_value(book, sheet, row, *column).then_some(title.as_str())
                        })
                        .collect();
                    if label.is_empty() {
                        if !populated_metrics.is_empty() {
                            let titles = populated_metrics
                                .iter()
                                .map(|title| format!("“{title}”"))
                                .collect::<Vec<_>>()
                                .join("、");
                            push_warning(
                                &mut warnings,
                                sheet,
                                row,
                                format!(
                                    "行名称为空，但指标列{titles}有数据；本行无法匹配报表项目，未参与外部取数。请检查行名称是否被误删。"
                                ),
                            );
                        } else if !row_has_values.get(&row).copied().unwrap_or(false)
                            && row < last_nonempty_row
                        {
                            push_warning(
                                &mut warnings,
                                sheet,
                                row,
                                "行名称及指标数据均为空，已忽略该行；如属导出空行，可忽略。",
                            );
                        }
                    }
                    if let Some(date_column) = date_column {
                        if targets.contains(&label)
                            && row_dates.get(&row).map_or(true, String::is_empty)
                        {
                            let raw_date = book.cell_raw_text(sheet, row, date_column);
                            let date_display = if raw_date.trim().is_empty() {
                                "空白".to_owned()
                            } else {
                                format!("“{}”", raw_date.trim())
                            };
                            push_warning(
                                &mut warnings,
                                sheet,
                                row,
                                format!(
                                    "行名称“{}”对应的数据日期{}为空或无效，本行未参与时序匹配；请检查数据日期是否误删或格式错误。",
                                    raw_label.trim(), date_display
                                ),
                            );
                        }
                    }
                }
            }
        }
        if columns.is_empty() {
            return Err(ExternalLayoutFailure::new(
                "EXTERNAL_LAYOUT_MISSING",
                "未找到机构报表项目行及其右侧的外部指标表头",
            ));
        }
        Ok(Self {
            mapping_sheet,
            mapping_rows,
            columns,
            warnings,
        })
    }

    pub fn resolve(
        &self,
        institution_code: &str,
        institution_name: &str,
        metric_header: &str,
        sheet_filter: &str,
        data_date: &str,
    ) -> Result<ExternalLookup, ExternalLayoutFailure> {
        let (entry, mapping_method) = self.institution_entry(institution_code, institution_name)?;
        let metric = normalize_external_label(metric_header);
        let metric_columns: Vec<&ExternalColumn> = self
            .columns
            .iter()
            .filter(|column| {
                normalize_external_label(&column.header) == metric
                    && (sheet_filter.is_empty() || column.sheet == sheet_filter)
            })
            .collect();
        if metric_columns.is_empty() {
            return Err(ExternalLayoutFailure::new(
                "METRIC_NOT_FOUND",
                "外部文件中不存在 V_EXT 指定的指标表头",
            ));
        }
        if metric_columns.len() != 1 {
            return Err(ExternalLayoutFailure::new(
                "METRIC_AMBIGUOUS",
                "外部指标表头重复，不能唯一定位；可通过 sheet= 指定工作表",
            ));
        }
        let column = metric_columns[0];
        let mut row_matches = column
            .rows
            .get(&normalize_external_label(&entry.report_item))
            .cloned()
            .unwrap_or_default();
        if column.data_date_column.is_some() {
            let selected_date = crate::s2_external::normalize_external_date(data_date);
            if selected_date.is_empty() {
                return Err(ExternalLayoutFailure::new(
                    "EXTERNAL_DATE_INVALID",
                    "本期报告期不是有效的完整日期，无法匹配外部时序数据",
                ));
            }
            row_matches.retain(|row| {
                column
                    .row_dates
                    .get(row)
                    .is_some_and(|row_date| row_date == &selected_date)
            });
            if row_matches.is_empty() {
                return Err(ExternalLayoutFailure::new(
                    "EXTERNAL_DATE_NOT_FOUND",
                    format!(
                        "外部时序数据中，报表项目“{}”没有本期日期 {} 的记录",
                        entry.report_item, selected_date
                    ),
                ));
            }
        }
        if row_matches.is_empty() {
            return Err(ExternalLayoutFailure::new(
                "REPORT_ITEM_NOT_FOUND",
                "外部指标表中不存在机构映射得到的报表项目",
            ));
        }
        if row_matches.len() != 1 {
            return Err(ExternalLayoutFailure::new(
                "REPORT_ITEM_DUPLICATE",
                "外部指标表中机构映射得到的报表项目重复",
            ));
        }
        let mapping_column = if mapping_method == "统一社会信用代码" {
            entry.code_column
        } else {
            entry.name_column
        };
        let mapping_key = if mapping_method == "统一社会信用代码" {
            entry.code.clone()
        } else {
            entry.name.clone()
        };
        Ok(ExternalLookup {
            mapping_sheet: self.mapping_sheet.clone(),
            mapping_cell: format!(
                "{}{}:{}{}",
                column_letter(mapping_column),
                entry.row,
                column_letter(entry.target_column),
                entry.row
            ),
            mapping_key,
            mapping_method,
            report_item: entry.report_item.clone(),
            sheet: column.sheet.clone(),
            row: row_matches[0],
            column: column.column,
        })
    }

    fn institution_entry(
        &self,
        institution_code: &str,
        institution_name: &str,
    ) -> Result<(&InstitutionMapping, String), ExternalLayoutFailure> {
        let code_key = normalize_external_label(institution_code);
        let name_key = normalize_external_label(institution_name);
        let mut entries: Vec<&InstitutionMapping> = self
            .mapping_rows
            .iter()
            .filter(|entry| {
                !code_key.is_empty() && normalize_external_label(&entry.code) == code_key
            })
            .collect();
        let mut method = "统一社会信用代码";
        if !entries.is_empty() && entries.iter().any(|entry| entry.disabled) {
            return Err(ExternalLayoutFailure::new(
                "INSTITUTION_DISABLED",
                "当前机构在外部参照表中标为禁用，请先核对禁用设置",
            ));
        }
        if !entries.iter().any(|entry| !entry.report_item.is_empty()) {
            entries.clear();
        }
        if entries.is_empty() {
            entries = self
                .mapping_rows
                .iter()
                .filter(|entry| {
                    !name_key.is_empty() && normalize_external_label(&entry.name) == name_key
                })
                .collect();
            method = "机构名称";
            if entries.iter().any(|entry| entry.disabled) {
                return Err(ExternalLayoutFailure::new(
                    "INSTITUTION_DISABLED",
                    "当前机构在外部参照表中标为禁用，请先核对禁用设置",
                ));
            }
            if !entries.iter().any(|entry| !entry.report_item.is_empty()) {
                entries.clear();
            }
        }
        if entries.is_empty() {
            return Err(ExternalLayoutFailure::new(
                "IDENTITY_NOT_MAPPED",
                "信用代码及机构名称均未在外部文件中匹配到报表项目",
            ));
        }

        let codes: BTreeSet<String> = entries
            .iter()
            .filter(|entry| !entry.code.is_empty())
            .map(|entry| normalize_external_label(&entry.code))
            .collect();
        let targets: BTreeSet<String> = entries
            .iter()
            .filter(|entry| !entry.report_item.is_empty())
            .map(|entry| normalize_external_label(&entry.report_item))
            .collect();
        if codes.len() > 1 || targets.len() > 1 {
            return Err(ExternalLayoutFailure::new(
                "INSTITUTION_REFERENCE_CONFLICT",
                "机构参照匹配到多个信用代码或报表项目，请核对外部文件各参照表中的映射",
            ));
        }
        let mut seen = BTreeSet::new();
        for entry in &entries {
            let key = if method == "机构名称" {
                normalize_external_label(&entry.name)
            } else {
                normalize_external_label(&entry.code)
            };
            if !seen.insert(key) {
                return Err(ExternalLayoutFailure::new(
                    "IDENTITY_MAPPING_AMBIGUOUS",
                    "外部文件同一参照表中机构匹配多行，请保留唯一机构记录",
                ));
            }
        }
        let entry = entries
            .into_iter()
            .find(|entry| !entry.report_item.is_empty())
            .expect("filtered institution entries have a report item");
        Ok((entry, method.to_owned()))
    }
}

fn discover_institution_mapping(
    book: &ScanBook,
) -> Result<(String, Vec<InstitutionMapping>), ExternalLayoutFailure> {
    for sheet in MAPPING_SHEETS {
        if !book.sheet_names().iter().any(|name| name == sheet) {
            continue;
        }
        let mut header = None;
        for row in 1..=book.sheet_max_row(sheet).min(30) {
            let labels: Vec<String> = (1..=book.sheet_max_col(sheet))
                .map(|column| normalize_external_label(&book.cell_raw_text(sheet, row, column)))
                .collect();
            let code_columns: Vec<usize> = labels
                .iter()
                .enumerate()
                .filter_map(|(index, label)| {
                    IDENTITY_LABELS
                        .iter()
                        .any(|candidate| normalize_external_label(candidate) == *label)
                        .then_some(index + 1)
                })
                .collect();
            let name_columns: Vec<usize> = labels
                .iter()
                .enumerate()
                .filter_map(|(index, label)| {
                    NAME_LABELS
                        .iter()
                        .any(|candidate| normalize_external_label(candidate) == *label)
                        .then_some(index + 1)
                })
                .collect();
            let target_columns: Vec<usize> = labels
                .iter()
                .enumerate()
                .filter_map(|(index, label)| {
                    TARGET_LABELS
                        .iter()
                        .any(|candidate| normalize_external_label(candidate) == *label)
                        .then_some(index + 1)
                })
                .collect();
            if code_columns.is_empty() || (name_columns.is_empty() && target_columns.is_empty()) {
                continue;
            }
            if code_columns.len() != 1 || name_columns.len() > 1 || target_columns.len() > 1 {
                return Err(ExternalLayoutFailure::new(
                    "REFERENCE_HEADERS_AMBIGUOUS",
                    "外部机构参照表头重复，请保留唯一代码、名称和报表项目字段",
                ));
            }
            let disabled_column = labels
                .iter()
                .position(|label| label == "禁用")
                .map(|index| index + 1);
            header = Some((
                row,
                code_columns[0],
                name_columns.first().copied().unwrap_or(0),
                target_columns.first().copied().unwrap_or(0),
                disabled_column,
            ));
            break;
        }
        let Some((header_row, code_column, name_column, target_column, disabled_column)) = header
        else {
            continue;
        };
        let mut entries = Vec::new();
        for row in (header_row + 1)..=book.sheet_max_row(sheet) {
            let code = mapping_text(book, sheet, row, code_column)?;
            let name = if name_column == 0 {
                String::new()
            } else {
                mapping_text(book, sheet, row, name_column)?
            };
            let report_item = if target_column == 0 {
                String::new()
            } else {
                mapping_text(book, sheet, row, target_column)?
            };
            if code.is_empty() && name.is_empty() {
                continue;
            }
            let disabled = if let Some(column) = disabled_column {
                let flag = mapping_text(book, sheet, row, column)?;
                let normalized = normalize_external_label(&flag);
                if ![
                    "", "否", "0", "false", "no", "是", "1", "true", "yes", "禁用",
                ]
                .contains(&normalized.as_str())
                {
                    return Err(ExternalLayoutFailure::new(
                        "REFERENCE_FLAG_INVALID",
                        "机构参照表禁用字段应填写是或否",
                    ));
                }
                ["是", "1", "true", "yes", "禁用"].contains(&normalized.as_str())
            } else {
                false
            };
            entries.push(InstitutionMapping {
                code,
                name,
                report_item,
                disabled,
                row,
                code_column,
                name_column,
                target_column,
            });
        }
        if entries.iter().any(|entry| !entry.report_item.is_empty()) {
            return Ok((sheet.to_owned(), entries));
        }
    }
    Err(ExternalLayoutFailure::new(
        "MAPPING_SHEET_MISSING",
        "外部文件中未找到包含机构代码和报表项目的映射表",
    ))
}

fn mapping_text(
    book: &ScanBook,
    sheet: &str,
    row: usize,
    column: usize,
) -> Result<String, ExternalLayoutFailure> {
    if book.sheet_has_formula(sheet, row, column) {
        return Err(ExternalLayoutFailure::new(
            "REFERENCE_FORMULA_UNVERIFIED",
            "外部机构映射包含未验证公式",
        ));
    }
    Ok(book.stored_value(sheet, row, column).1.trim().to_owned())
}

fn has_value(book: &ScanBook, sheet: &str, row: usize, column: usize) -> bool {
    if book.sheet_has_formula(sheet, row, column) {
        return true;
    }
    let (kind, value) = book.stored_value(sheet, row, column);
    kind != "empty" && (kind != "string" || !value.trim().is_empty())
}

fn time_series_header(book: &ScanBook, sheet: &str) -> Option<(usize, usize, usize)> {
    for row in 1..=book.sheet_max_row(sheet).min(20) {
        let headers: Vec<String> = (1..=book.sheet_max_col(sheet))
            .map(|column| normalize_external_label(&book.cell_raw_text(sheet, row, column)))
            .collect();
        let date_columns: Vec<usize> = headers
            .iter()
            .enumerate()
            .filter_map(|(index, header)| (header == "数据日期").then_some(index + 1))
            .collect();
        let label_columns: Vec<usize> = headers
            .iter()
            .enumerate()
            .filter_map(|(index, header)| (header == "行名称").then_some(index + 1))
            .collect();
        if date_columns.len() == 1 && label_columns.len() == 1 {
            return Some((row, date_columns[0], label_columns[0]));
        }
    }
    None
}

fn external_date_value(
    book: &ScanBook,
    iso_dates: &BTreeMap<String, BTreeMap<String, String>>,
    sheet: &str,
    row: usize,
    column: usize,
) -> String {
    if book.sheet_has_formula(sheet, row, column) {
        return String::new();
    }
    let (kind, raw) = book.stored_value(sheet, row, column);
    if kind == "datetime_serial" {
        let address = format!("{}{row}", column_letter(column));
        return iso_dates
            .get(sheet)
            .and_then(|values| values.get(&address))
            .map(|value| crate::s2_external::normalize_external_date(value))
            .unwrap_or_default();
    }
    if matches!(
        kind.as_str(),
        "string" | "datetime_iso" | "integer" | "number"
    ) {
        return crate::s2_external::normalize_external_date(&raw);
    }
    String::new()
}

fn push_warning(
    warnings: &mut Vec<ExternalLayoutWarning>,
    sheet: &str,
    row: usize,
    message: impl Into<String>,
) {
    let warning = ExternalLayoutWarning {
        sheet: sheet.to_owned(),
        location: format!("第{row}行"),
        message: message.into(),
    };
    if !warnings.contains(&warning) {
        warnings.push(warning);
    }
}

fn column_letter(mut column: usize) -> String {
    let mut letters = Vec::new();
    while column > 0 {
        letters.insert(0, (b'A' + ((column - 1) % 26) as u8) as char);
        column = (column - 1) / 26;
    }
    letters.into_iter().collect()
}
