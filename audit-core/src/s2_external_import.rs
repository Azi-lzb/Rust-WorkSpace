//! S2 external workbook parsing migrated from the read-only Python importer.
//!
//! The parser accepts the source project's `.xlsx` worksheet contract and uses
//! the shared Rust Calamine adapter. It does not select or launch a Python
//! reader. Parsing warnings and the external row-label map are returned beside
//! the comparison input so a later service layer can preserve the source flow.

use crate::{
    s2_external::{
        normalize_external_date, python_float_text, S2ExternalData, S2ExternalValue,
        S2ExternalValueKey,
    },
    s2_import::{
        normalize_fast_value, parse_cell_reference, python_code, python_string, python_truthy,
        ImportedValue, S2SourceRef,
    },
    unit_conversion::{decimal_number_from_cell, AmountSourceFormat, DecimalTuple},
};
use anyhow::{bail, Context, Result};
use excel_reader::{read_value_columns_with_iso_dates, CellValue, SheetValues};
use std::{
    collections::{BTreeMap, HashMap, HashSet},
    path::{Path, PathBuf},
};

const EXTERNAL_SHEET: &str = "集中系统数据";
const CALAMINE_VERSION: &str = "0.37.0";
const EXCEL_MAX_COLUMNS: usize = 16_384;
const EXTERNAL_METADATA_HEADERS: &[&str] = &[
    "数据日期",
    "报表名称",
    "报表代码",
    "版本号",
    "锁定信息",
    "报表组",
    "行顺序号",
    "行名称",
];

/// Parsed data returned by the S2 external workbook import operation.
#[derive(Clone, Debug)]
pub struct S2ExternalImportResult {
    pub data: S2ExternalData,
    pub external_org_map: HashMap<String, String>,
    pub warnings: Vec<String>,
    pub reader_engine: String,
    pub reader_version: String,
    pub reader_elapsed_seconds: f64,
}

/// Read the S2 external workbook using the pure Rust Calamine adapter.
///
/// The source importer intentionally accepts only `.xlsx`: its organization
/// mapping contract depends on the `机构参照表` / legacy `参照表` worksheet.
pub fn import_external_workbook(path: impl AsRef<Path>) -> Result<S2ExternalImportResult> {
    let path = path.as_ref();
    if !path.is_file() {
        bail!("大集中文件不存在：{}", path.display());
    }
    let is_xlsx = path
        .extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"));
    if !is_xlsx {
        bail!(
            "大集中数据仅支持 .xlsx 工作簿：{}（.xls/.csv 会丢失「机构参照表/参照表」机构映射，不支持）",
            file_name(path)
        );
    }

    let started = std::time::Instant::now();
    let sheets = read_value_columns_with_iso_dates(path, EXCEL_MAX_COLUMNS)
        .with_context(|| format!("读取大集中工作簿失败：{}", path.display()))?;
    let main = sheets
        .iter()
        .find(|sheet| sheet.name == EXTERNAL_SHEET)
        .with_context(|| "大集中文件缺少“集中系统数据”工作表")?;
    let matrix = SparseSheet::from_sheet(main);
    let first_header_row = (1..=matrix.max_row.unwrap_or(0)).find(|row| {
        matrix
            .row_cells(*row)
            .any(|(_, value)| matches!(value, CellValue::Text(text) if !text.trim().is_empty()))
    });
    let Some(first_header_row) = first_header_row else {
        bail!("集中系统数据未找到表头");
    };

    let first_twenty = 1..=matrix.max_row.unwrap_or(0).min(20);
    let time_header_row = first_twenty
        .clone()
        .find(|row| contains_headers(&matrix.headers(*row), &["数据日期", "行名称"]));
    let plain_header_row = first_twenty.clone().find(|row| {
        matrix
            .headers(*row)
            .into_iter()
            .take(4)
            .any(|header| header == "行名称" || header == "报表项目")
    });

    let source_file = resolved_path(path);
    let mut data = S2ExternalData::default();
    let mut warnings = Vec::new();
    if let Some(header_row) = time_header_row {
        read_time_series(&matrix, header_row, &source_file, &mut data, &mut warnings)?;
    } else {
        if let Some(partial_header_row) = first_twenty.clone().find(|row| {
            matrix
                .headers(*row)
                .iter()
                .any(|header| header == "数据日期")
        }) {
            bail!(
                "集中系统数据第{partial_header_row}行：时序表头不完整，必须同时包含“数据日期”和“行名称”"
            );
        }
        read_plain(
            &matrix,
            plain_header_row.unwrap_or(first_header_row),
            &source_file,
            &mut data,
            &mut warnings,
        );
    }

    let reference = sheets
        .iter()
        .find(|sheet| sheet.name == "机构参照表")
        .or_else(|| sheets.iter().find(|sheet| sheet.name == "参照表"));
    let external_org_map = match reference {
        Some(sheet) => read_external_org_map(&SparseSheet::from_sheet(sheet), &sheet.name)?,
        None => HashMap::new(),
    };

    Ok(S2ExternalImportResult {
        data,
        external_org_map,
        warnings,
        reader_engine: "Calamine".to_owned(),
        reader_version: CALAMINE_VERSION.to_owned(),
        reader_elapsed_seconds: started.elapsed().as_secs_f64(),
    })
}

#[derive(Default)]
struct SparseSheet {
    cells: BTreeMap<(usize, usize), CellValue>,
    max_row: Option<usize>,
    max_column: Option<usize>,
}

impl SparseSheet {
    fn from_sheet(sheet: &SheetValues) -> Self {
        let mut result = Self::default();
        for (address, value) in &sheet.values {
            let Some((row, column)) = parse_cell_reference(address) else {
                continue;
            };
            let value = normalize_fast_value(value, AmountSourceFormat::XlsxLike);
            if matches!(value, CellValue::Empty) {
                continue;
            }
            result.max_row = Some(result.max_row.unwrap_or(0).max(row));
            result.max_column = Some(result.max_column.unwrap_or(0).max(column));
            result.cells.insert((row, column), value);
        }
        result
    }

    fn get(&self, row: usize, column: usize) -> Option<&CellValue> {
        self.cells.get(&(row, column))
    }

    fn row_cells(&self, row: usize) -> impl Iterator<Item = (usize, &CellValue)> {
        self.cells
            .range((row, 1)..=(row, usize::MAX))
            .map(|((_, column), value)| (*column, value))
    }

    fn headers(&self, row: usize) -> Vec<String> {
        (1..=self.max_column.unwrap_or(0))
            .map(|column| external_header(self.get(row, column)))
            .collect()
    }

    fn has_external_value(&self, row: usize, column: usize) -> bool {
        self.get(row, column).is_some_and(cell_has_external_value)
    }

    fn row_has_external_value(&self, row: usize) -> bool {
        self.row_cells(row)
            .any(|(_, value)| cell_has_external_value(value))
    }

    fn last_nonempty_row(&self, start: usize) -> Option<usize> {
        self.cells
            .iter()
            .filter(|((row, _), value)| *row >= start && cell_has_external_value(value))
            .map(|((row, _), _)| *row)
            .max()
    }
}

fn read_time_series(
    matrix: &SparseSheet,
    header_row: usize,
    source_file: &str,
    data: &mut S2ExternalData,
    warnings: &mut Vec<String>,
) -> Result<()> {
    let headers = matrix.headers(header_row);
    for field in ["数据日期", "行名称"] {
        if headers
            .iter()
            .filter(|header| header.as_str() == field)
            .count()
            != 1
        {
            bail!("集中系统数据第{header_row}行：时序格式必须有且只有一列“{field}”");
        }
    }
    let date_column = headers
        .iter()
        .position(|header| header == "数据日期")
        .expect("validated date column")
        + 1;
    let label_column = headers
        .iter()
        .position(|header| header == "行名称")
        .expect("validated label column")
        + 1;
    let mut indicators = Vec::new();
    let mut seen_names = HashSet::new();
    for (index, title) in headers.into_iter().enumerate() {
        if title.is_empty()
            || title.eq_ignore_ascii_case("null")
            || EXTERNAL_METADATA_HEADERS.contains(&title.as_str())
        {
            continue;
        }
        if !seen_names.insert(title.clone()) {
            bail!("集中系统数据第{header_row}行：时序格式的指标列“{title}”重复，无法确定取数列");
        }
        indicators.push((index + 1, title.clone()));
        data.indicator_names.insert(title);
    }
    if indicators.is_empty() {
        bail!("集中系统数据第{header_row}行：时序格式缺少指标列");
    }
    data.time_series = true;

    let last_nonempty_row = matrix.last_nonempty_row(header_row + 1);
    let mut seen_rows: HashMap<(String, String), usize> = HashMap::new();
    for row_number in header_row + 1..=matrix.max_row.unwrap_or(header_row) {
        let label = truthy_string(matrix.get(row_number, label_column));
        let populated_indicators = indicators
            .iter()
            .filter(|(column, _)| matrix.has_external_value(row_number, *column))
            .map(|(_, title)| title.as_str())
            .collect::<Vec<_>>();
        if label.is_empty() {
            if !populated_indicators.is_empty() {
                let titles = populated_indicators
                    .iter()
                    .map(|title| format!("“{title}”"))
                    .collect::<Vec<_>>()
                    .join("、");
                push_warning(
                    warnings,
                    format!(
                        "集中系统数据第{row_number}行：行名称为空，但指标列{titles}有数据；本行无法匹配报表项目，未参与外部核对。请检查行名称是否被误删。"
                    ),
                );
            } else if !matrix.row_has_external_value(row_number)
                && last_nonempty_row.is_some_and(|last| row_number < last)
            {
                push_warning(
                    warnings,
                    format!(
                        "集中系统数据第{row_number}行：行名称及指标数据均为空，已忽略该行；如属导出空行，可忽略。"
                    ),
                );
            }
            continue;
        }

        let raw_date = matrix.get(row_number, date_column);
        let data_date = normalize_date_cell(raw_date);
        if data_date.is_empty() {
            let date_display = if raw_date.is_none_or(cell_is_blank) {
                "空白".to_owned()
            } else {
                format!(
                    "“{}”",
                    external_value_preview(raw_date.expect("checked"), 80)
                )
            };
            push_warning(
                warnings,
                format!(
                    "集中系统数据第{row_number}行、行名称“{label}”：数据日期{date_display}为空或无效，本行未参与时序匹配；请检查数据日期是否误删或格式错误。"
                ),
            );
            continue;
        }

        let row_key = (data_date.clone(), label.clone());
        if let Some(first_row) = seen_rows.insert(row_key.clone(), row_number) {
            bail!(
                "集中系统数据第{first_row}、{row_number}行：日期 {data_date}、报表项目“{label}”重复，请保留本轮使用的唯一记录，不自动合计或覆盖不同报表版本"
            );
        }
        data.available_dates.insert(data_date.clone());
        data.date_labels.insert(row_key);
        data.row_labels.insert(label.clone());
        let source = S2SourceRef {
            source_file: source_file.to_owned(),
            source_sheet: EXTERNAL_SHEET.to_owned(),
            source_row: row_number,
        };
        let mut blank_titles = Vec::new();
        let mut non_numeric_titles = Vec::new();
        for (column, title) in &indicators {
            let raw_value = matrix.get(row_number, *column);
            let value = if raw_value.is_none_or(|value| !cell_has_external_value(value)) {
                data.blank_values.insert(S2ExternalValueKey::TimeSeries(
                    data_date.clone(),
                    label.clone(),
                    title.clone(),
                ));
                blank_titles.push(format!("“{title}”为空白"));
                DecimalTuple {
                    sign: 0,
                    digits: vec![0],
                    exponent: 0,
                }
            } else if let Some(number) = decimal_from_cell(raw_value) {
                number
            } else {
                non_numeric_titles.push(format!(
                    "“{title}”内容“{}”不是数值",
                    external_value_preview(raw_value.expect("checked"), 80)
                ));
                DecimalTuple {
                    sign: 0,
                    digits: vec![0],
                    exponent: 0,
                }
            };
            data.values.push(S2ExternalValue {
                data_date: data_date.clone(),
                row_label: label.clone(),
                indicator_name: title.clone(),
                value: ImportedValue::Decimal(value),
                source: source.clone(),
                fallback_reason: String::new(),
            });
        }
        add_row_value_warning(
            warnings,
            row_number,
            &label,
            blank_titles,
            non_numeric_titles,
        );
    }
    Ok(())
}

fn read_plain(
    matrix: &SparseSheet,
    header_row: usize,
    source_file: &str,
    data: &mut S2ExternalData,
    warnings: &mut Vec<String>,
) {
    let headers = matrix.headers(header_row);
    let label_column = headers
        .iter()
        .take(4)
        .position(|header| header == "行名称" || header == "报表项目")
        .map(|index| index + 1)
        .or_else(|| {
            (1..=matrix.max_column.unwrap_or(0).min(4)).find(|column| {
                truthy_string(matrix.get(header_row, *column)).is_empty()
                    && (header_row + 1..=matrix.max_row.unwrap_or(header_row).min(header_row + 4))
                        .any(|row| {
                            matches!(matrix.get(row, *column), Some(CellValue::Text(text)) if !text.trim().is_empty())
                        })
            })
        });

    let Some(label_column) = label_column else {
        data.missing_label_fallback = Some(S2ExternalValue {
            data_date: String::new(),
            row_label: String::new(),
            indicator_name: String::new(),
            value: ImportedValue::Decimal(DecimalTuple {
                sign: 0,
                digits: vec![0],
                exponent: 0,
            }),
            source: S2SourceRef {
                source_file: source_file.to_owned(),
                source_sheet: EXTERNAL_SHEET.to_owned(),
                source_row: header_row,
            },
            fallback_reason: "机构行标签列未识别，外部数按0参与比较".to_owned(),
        });
        return;
    };

    let indicators = headers
        .iter()
        .enumerate()
        .filter_map(|(index, _)| {
            let column = index + 1;
            // The Python ordinary-layout path uses `str(title or "").strip()`
            // here, unlike its time-series header normalization. Numeric zero
            // and boolean false therefore do not create indicator columns.
            let title = truthy_string(matrix.get(header_row, column));
            (!title.is_empty()
                && column != label_column
                && !EXTERNAL_METADATA_HEADERS.contains(&title.as_str()))
            .then_some((column, title))
        })
        .collect::<Vec<_>>();
    data.indicator_names
        .extend(indicators.iter().map(|(_, title)| title.clone()));

    let last_nonempty_row = matrix.last_nonempty_row(header_row + 1);
    let mut final_values = BTreeMap::<(String, String), S2ExternalValue>::new();
    for row_number in header_row + 1..=matrix.max_row.unwrap_or(header_row) {
        let label = truthy_string(matrix.get(row_number, label_column));
        let populated_indicators = indicators
            .iter()
            .filter(|(column, _)| matrix.has_external_value(row_number, *column))
            .map(|(_, title)| title.as_str())
            .collect::<Vec<_>>();
        if label.is_empty() {
            if !populated_indicators.is_empty() {
                let titles = populated_indicators
                    .iter()
                    .map(|title| format!("“{title}”"))
                    .collect::<Vec<_>>()
                    .join("、");
                push_warning(
                    warnings,
                    format!(
                        "集中系统数据第{row_number}行：行名称为空，但指标列{titles}有数据；本行无法匹配报表项目，未参与外部核对。请检查行名称是否被误删。"
                    ),
                );
            } else if !matrix.row_has_external_value(row_number)
                && last_nonempty_row.is_some_and(|last| row_number < last)
            {
                push_warning(
                    warnings,
                    format!(
                        "集中系统数据第{row_number}行：行名称及指标数据均为空，已忽略该行；如属导出空行，可忽略。"
                    ),
                );
            }
            continue;
        }

        data.row_labels.insert(label.clone());
        let source = S2SourceRef {
            source_file: source_file.to_owned(),
            source_sheet: EXTERNAL_SHEET.to_owned(),
            source_row: row_number,
        };
        let mut blank_titles = Vec::new();
        let mut non_numeric_titles = Vec::new();
        for (column, title) in &indicators {
            let raw_value = matrix.get(row_number, *column);
            let value = if raw_value.is_none_or(|value| !cell_has_external_value(value)) {
                data.blank_values
                    .insert(S2ExternalValueKey::Plain(label.clone(), title.clone()));
                blank_titles.push(format!("“{title}”为空白"));
                DecimalTuple {
                    sign: 0,
                    digits: vec![0],
                    exponent: 0,
                }
            } else if let Some(number) = decimal_from_cell(raw_value) {
                number
            } else {
                non_numeric_titles.push(format!(
                    "“{title}”内容“{}”不是数值",
                    external_value_preview(raw_value.expect("checked"), 80)
                ));
                DecimalTuple {
                    sign: 0,
                    digits: vec![0],
                    exponent: 0,
                }
            };
            final_values.insert(
                (label.clone(), title.clone()),
                S2ExternalValue {
                    data_date: String::new(),
                    row_label: label.clone(),
                    indicator_name: title.clone(),
                    value: ImportedValue::Decimal(value),
                    source: source.clone(),
                    fallback_reason: String::new(),
                },
            );
        }
        add_row_value_warning(
            warnings,
            row_number,
            &label,
            blank_titles,
            non_numeric_titles,
        );
    }
    data.values.extend(final_values.into_values());
}

fn read_external_org_map(
    matrix: &SparseSheet,
    sheet_name: &str,
) -> Result<HashMap<String, String>> {
    if matrix.max_row.is_none() {
        return Ok(HashMap::new());
    }
    let headers = (1..=matrix.max_column.unwrap_or(0))
        .map(|column| truthy_string(matrix.get(1, column)))
        .collect::<Vec<_>>();
    let code_column = ["社会信用代码", "统一社会信用代码"]
        .iter()
        .find_map(|alias| {
            headers
                .iter()
                .position(|header| header == alias)
                .map(|index| index + 1)
        });
    let report_item_column = headers
        .iter()
        .position(|header| header == "报表项目")
        .map(|index| index + 1);
    let (Some(code_column), Some(report_item_column)) = (code_column, report_item_column) else {
        bail!("大集中文件“{sheet_name}”缺少必要表头：社会信用代码（或统一社会信用代码）、报表项目");
    };
    let disabled_column = headers
        .iter()
        .position(|header| header == "禁用")
        .map(|index| index + 1);
    let mut org_map = HashMap::new();
    for row in 2..=matrix.max_row.unwrap_or(1) {
        if disabled_column.is_some_and(|column| is_yes(matrix.get(row, column))) {
            continue;
        }
        let code = python_code(matrix.get(row, code_column));
        let report_item = truthy_string(matrix.get(row, report_item_column));
        if !code.is_empty() && !report_item.is_empty() {
            org_map.insert(code, report_item);
        }
    }
    Ok(org_map)
}

fn contains_headers(headers: &[String], required: &[&str]) -> bool {
    required
        .iter()
        .all(|required| headers.contains(&required.to_string()))
}

fn external_header(value: Option<&CellValue>) -> String {
    let text = value.map(external_python_string).unwrap_or_default();
    text.trim().trim_start_matches('\u{feff}').trim().to_owned()
}

fn truthy_string(value: Option<&CellValue>) -> String {
    value
        .filter(|value| python_truthy(value))
        .map(external_python_string)
        .unwrap_or_default()
        .trim()
        .to_owned()
}

fn external_python_string(value: &CellValue) -> String {
    match value {
        CellValue::Number(value) => python_float_text(*value),
        CellValue::DateTimeIso(value) => python_datetime_text(value),
        _ => python_string(value),
    }
}

fn python_datetime_text(value: &str) -> String {
    let Some((date, time)) = value.split_once('T') else {
        return value.to_owned();
    };
    let Some((clock, milliseconds)) = time.split_once('.') else {
        return format!("{date} {time}");
    };
    if milliseconds.chars().all(|character| character == '0') {
        format!("{date} {clock}")
    } else {
        format!("{date} {clock}.{milliseconds}000")
    }
}

fn normalize_date_cell(value: Option<&CellValue>) -> String {
    let Some(value) = value else {
        return String::new();
    };
    match value {
        CellValue::Empty | CellValue::Boolean(_) | CellValue::DateTimeSerial(_) => String::new(),
        CellValue::Number(value) if value.is_finite() && value.fract() == 0.0 => {
            normalize_external_date(&format!("{value:.0}"))
        }
        CellValue::Integer(value) => normalize_external_date(&value.to_string()),
        CellValue::Text(value) | CellValue::DateTimeIso(value) => normalize_external_date(value),
        CellValue::Number(_) | CellValue::DurationIso(_) | CellValue::Error(_) => String::new(),
    }
}

fn cell_is_blank(value: &CellValue) -> bool {
    matches!(value, CellValue::Empty)
        || matches!(value, CellValue::Text(text) if text.trim().is_empty())
}

fn cell_has_external_value(value: &CellValue) -> bool {
    !cell_is_blank(value)
}

fn decimal_from_cell(value: Option<&CellValue>) -> Option<DecimalTuple> {
    decimal_number_from_cell(value, AmountSourceFormat::XlsxLike)
        .ok()
        .flatten()
}

fn external_value_preview(value: &CellValue, limit: usize) -> String {
    let text = external_python_string(value).replace(['\r', '\n'], " ");
    let mut characters = text.chars();
    let prefix = characters.by_ref().take(limit).collect::<String>();
    if characters.next().is_none() {
        prefix
    } else {
        let keep = limit.saturating_sub(3);
        format!("{}...", prefix.chars().take(keep).collect::<String>())
    }
}

fn add_row_value_warning(
    warnings: &mut Vec<String>,
    row_number: usize,
    label: &str,
    blank_titles: Vec<String>,
    non_numeric_titles: Vec<String>,
) {
    if blank_titles.is_empty() && non_numeric_titles.is_empty() {
        return;
    }
    let details = [blank_titles, non_numeric_titles]
        .into_iter()
        .flatten()
        .collect::<Vec<_>>()
        .join("；");
    push_warning(
        warnings,
        format!(
            "集中系统数据第{row_number}行、报表项目“{label}”：{details}，已按 0 参与外部核对；请核查是否误删或改动了数据。"
        ),
    );
}

fn push_warning(warnings: &mut Vec<String>, message: String) {
    if !warnings.contains(&message) {
        warnings.push(message);
    }
}

fn is_yes(value: Option<&CellValue>) -> bool {
    let text = truthy_string(value).to_lowercase();
    matches!(text.as_str(), "是" | "1" | "true" | "y" | "yes")
}

fn resolved_path(path: &Path) -> String {
    path.canonicalize()
        .unwrap_or_else(|_| PathBuf::from(path))
        .to_string_lossy()
        .into_owned()
}

fn file_name(path: &Path) -> String {
    path.file_name()
        .map(|name| name.to_string_lossy().into_owned())
        .unwrap_or_else(|| path.display().to_string())
}
