//! S1-F02 workbook-table merge implemented with Rust-owned values and output.
//!
//! This covers only the configured "汇总表合并" operation. It reads stored
//! values from XLSX workbooks, groups sheets by their first non-empty row, and
//! creates a new value-only workbook. Formula-bearing sheets are refused until
//! the formula recalculation/cache contract for this operation is validated.

use anyhow::{bail, ensure, Context, Result};
use excel_reader::{read_workbook, read_xlsx_worksheet_extents, CellValue, SheetDump};
use excel_writer::{Color, Format, Workbook};
use serde::Serialize;
use std::{
    collections::{BTreeMap, HashMap, HashSet},
    fs,
    path::{Path, PathBuf},
    time::{Instant, SystemTime, UNIX_EPOCH},
};

const READER_ENGINE: &str = "calamine+namespace_aware_ooxml";
const WRITER_ENGINE: &str = "rust_xlsxwriter 0.101.0";
const MAX_ROWS: usize = 1_048_576;
const MAX_COLUMNS: usize = 16_384;

#[derive(Clone, Debug, Serialize)]
pub struct S1SummaryMergeSheetSummary {
    pub sheet_name: String,
    /// Number of data rows, excluding the header.
    pub row_count: usize,
    pub column_count: usize,
    pub source_sheets: Vec<String>,
}

#[derive(Clone, Debug, Serialize)]
pub struct S1SummaryMergeSummary {
    pub output_path: PathBuf,
    pub source_workbooks: usize,
    pub source_sheets: usize,
    pub merged_sheets: usize,
    pub data_rows: usize,
    pub reader: &'static str,
    pub writer: &'static str,
    pub discovery_time_ms: u128,
    pub read_time_ms: u128,
    pub grouping_time_ms: u128,
    pub write_time_ms: u128,
    pub validation_time_ms: u128,
    pub elapsed_ms: u128,
    pub sheets: Vec<S1SummaryMergeSheetSummary>,
}

#[derive(Clone, Debug)]
struct SummaryGroup {
    headers: Vec<CellValue>,
    rows: Vec<Vec<CellValue>>,
    source_sheets: Vec<String>,
    output_name: String,
}

/// Merge source worksheet tables whose first non-empty rows have equal values.
///
/// This intentionally refuses formula-bearing workbooks. The source COM flow
/// opens workbooks in Excel, which may recalculate formulas; Calamine only
/// exposes stored formula caches. No formula cache is treated as a fresh
/// calculation here.
pub fn merge_s1_workbook_tables(
    input_dir: impl AsRef<Path>,
    output_path: impl AsRef<Path>,
    feature_name: &str,
    recursive: bool,
) -> Result<S1SummaryMergeSummary> {
    let started = Instant::now();
    let input_dir = fs::canonicalize(input_dir.as_ref()).with_context(|| {
        format!(
            "读取 S1-F02 汇总源目录失败：{}",
            input_dir.as_ref().display()
        )
    })?;
    ensure!(input_dir.is_dir(), "S1-F02 汇总源路径不是目录");

    let output_path = absolute_path(output_path.as_ref())?;
    ensure!(
        output_path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx")),
        "S1-F02 汇总输出必须使用 .xlsx"
    );
    ensure!(
        !output_path.exists(),
        "拒绝覆盖已有 S1-F02 汇总工作簿：{}",
        output_path.display()
    );

    let discovery_started = Instant::now();
    let sources = discover_source_workbooks(&input_dir, recursive)?;
    ensure!(!sources.is_empty(), "源数据目录中没有可汇总的 .xlsx 文件");
    let discovery_time_ms = discovery_started.elapsed().as_millis();

    let mut read_time_ms = 0_u128;
    let mut grouping_time_ms = 0_u128;
    let mut source_sheet_count = 0_usize;
    let mut groups: Vec<SummaryGroup> = Vec::new();
    let mut group_indices: HashMap<Vec<String>, usize> = HashMap::new();

    for source_path in &sources {
        let read_started = Instant::now();
        let workbook = read_workbook(source_path)
            .with_context(|| format!("读取 S1-F02 汇总源工作簿失败：{}", source_path.display()))?;
        let extents = read_xlsx_worksheet_extents(source_path)
            .with_context(|| format!("读取 S1-F02 工作表范围失败：{}", source_path.display()))?;
        read_time_ms += read_started.elapsed().as_millis();

        let grouping_started = Instant::now();
        let extents_by_sheet = extents
            .into_iter()
            .map(|extent| (extent.name.clone(), extent))
            .collect::<HashMap<_, _>>();
        for sheet in &workbook.sheets {
            source_sheet_count += 1;
            ensure!(
                sheet.formula_status == "read_by_calamine",
                "无法确认工作表“{}”是否含公式，停止 S1-F02 汇总",
                sheet.name
            );
            if !sheet.formulas.is_empty() {
                bail!(
                    "S1-F02 汇总暂不处理含公式的工作表；工作表“{}”含公式单元格，需先明确本轮计算语义",
                    sheet.name
                );
            }
            let Some(extent) = extents_by_sheet.get(&sheet.name) else {
                bail!("S1-F02 缺少工作表“{}”的 OOXML 单元格范围", sheet.name);
            };
            if extent.min_row == 0 || extent.min_column == 0 {
                continue;
            }
            ensure!(
                extent.max_row <= MAX_ROWS && extent.max_column <= MAX_COLUMNS,
                "工作表“{}”超过 XLSX 的最大行列数",
                sheet.name
            );

            let grid = decode_sheet_values(sheet)?;
            let rows = nonempty_rows(
                extent.min_row,
                extent.max_row,
                extent.min_column,
                extent.max_column,
                &grid,
            )?;
            let Some((header_values, data_rows)) = rows.split_first() else {
                continue;
            };
            if header_values.iter().all(is_empty_value) {
                continue;
            }
            let signature = header_values
                .iter()
                .map(header_signature_value)
                .collect::<Result<Vec<_>>>()?;
            let group_index = if let Some(index) = group_indices.get(&signature) {
                *index
            } else {
                let index = groups.len();
                group_indices.insert(signature.clone(), index);
                groups.push(SummaryGroup {
                    headers: header_values.clone(),
                    rows: Vec::new(),
                    source_sheets: Vec::new(),
                    output_name: String::new(),
                });
                index
            };
            let group = &mut groups[group_index];
            if !group.source_sheets.iter().any(|name| name == &sheet.name) {
                group.source_sheets.push(sheet.name.clone());
            }
            let source_name = source_path
                .file_name()
                .and_then(|name| name.to_str())
                .unwrap_or_default()
                .to_owned();
            for row in data_rows {
                let mut output_row = Vec::with_capacity(row.len() + 2);
                output_row.push(CellValue::Text(source_name.clone()));
                output_row.push(CellValue::Text(sheet.name.clone()));
                output_row.extend(row.iter().cloned());
                group.rows.push(output_row);
            }
        }
        grouping_time_ms += grouping_started.elapsed().as_millis();
    }

    let mut used_names = HashSet::new();
    for group in &mut groups {
        group.output_name = output_sheet_name(feature_name, &group.source_sheets, &mut used_names);
        ensure!(
            group.rows.len() + 1 <= MAX_ROWS && group.headers.len() + 2 <= MAX_COLUMNS,
            "S1-F02 输出超过 XLSX 的最大行列数"
        );
    }

    let output_parent = output_path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or(Path::new("."));
    fs::create_dir_all(output_parent)
        .with_context(|| format!("创建 S1-F02 输出目录失败：{}", output_parent.display()))?;
    let writer_temp = temporary_path(&output_path)?;
    let write_started = Instant::now();
    let write_result = (|| -> Result<()> {
        let mut workbook = Workbook::new();
        let value_format = Format::new().set_num_format("@");
        let header_format = Format::new()
            .set_num_format("@")
            .set_bold()
            .set_background_color(Color::RGB(0xD3_EA_D9));

        if groups.is_empty() {
            workbook.add_worksheet();
        } else {
            // The Windows Excel COM baseline serializes output sheets in
            // reverse group order. Keep its observed workbook order.
            for group in groups.iter().rev() {
                let worksheet = workbook.add_worksheet();
                worksheet.set_name(&group.output_name).with_context(|| {
                    format!("设置 S1-F02 输出工作表名失败：{}", group.output_name)
                })?;
                let mut header = vec![
                    CellValue::Text("来源文件".to_owned()),
                    CellValue::Text("来源工作表".to_owned()),
                ];
                header.extend(group.headers.iter().cloned());
                write_row(worksheet, 0, &header, &header_format)?;
                for (row_index, row) in group.rows.iter().enumerate() {
                    write_row(worksheet, row_index as u32 + 1, row, &value_format)?;
                }
                let column_count = header.len() as u16;
                worksheet
                    .autofilter(0, 0, group.rows.len() as u32, column_count - 1)
                    .context("设置 S1-F02 汇总筛选")?;
                // The Python COM writer requests FreezePanes, but the saved
                // workbook has no frozen panes on the tested Windows engine.
                worksheet.set_autofit_max_width(280).autofit();
            }
        }
        workbook
            .save(&writer_temp)
            .context("写入 S1-F02 汇总临时工作簿")?;
        Ok(())
    })();
    let write_time_ms = write_started.elapsed().as_millis();
    if let Err(error) = write_result {
        let _ = fs::remove_file(&writer_temp);
        return Err(error);
    }

    let validation_started = Instant::now();
    if let Err(error) = validate_summary_workbook(&writer_temp, &groups) {
        let _ = fs::remove_file(&writer_temp);
        return Err(error);
    }
    let validation_time_ms = validation_started.elapsed().as_millis();
    if let Err(error) = fs::rename(&writer_temp, &output_path) {
        let _ = fs::remove_file(&writer_temp);
        return Err(error).with_context(|| {
            format!(
                "提交 S1-F02 汇总工作簿失败：{} → {}",
                writer_temp.display(),
                output_path.display()
            )
        });
    }

    let sheets = groups
        .iter()
        .rev()
        .map(|group| S1SummaryMergeSheetSummary {
            sheet_name: group.output_name.clone(),
            row_count: group.rows.len(),
            column_count: group.headers.len() + 2,
            source_sheets: group.source_sheets.clone(),
        })
        .collect::<Vec<_>>();
    let data_rows = sheets.iter().map(|sheet| sheet.row_count).sum();
    Ok(S1SummaryMergeSummary {
        output_path,
        source_workbooks: sources.len(),
        source_sheets: source_sheet_count,
        merged_sheets: groups.len(),
        data_rows,
        reader: READER_ENGINE,
        writer: WRITER_ENGINE,
        discovery_time_ms,
        read_time_ms,
        grouping_time_ms,
        write_time_ms,
        validation_time_ms,
        elapsed_ms: started.elapsed().as_millis(),
        sheets,
    })
}

fn discover_source_workbooks(input_dir: &Path, recursive: bool) -> Result<Vec<PathBuf>> {
    fn visit(directory: &Path, recursive: bool, files: &mut Vec<PathBuf>) -> Result<()> {
        let entries = fs::read_dir(directory)
            .with_context(|| format!("扫描 S1-F02 汇总目录失败：{}", directory.display()))?;
        for entry in entries {
            let entry = entry?;
            let file_type = entry.file_type()?;
            let path = entry.path();
            if file_type.is_dir() && recursive {
                visit(&path, recursive, files)?;
            } else if file_type.is_file()
                && path
                    .extension()
                    .and_then(|extension| extension.to_str())
                    .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
                && !path
                    .file_name()
                    .and_then(|name| name.to_str())
                    .unwrap_or_default()
                    .starts_with("~$")
                && !path
                    .file_stem()
                    .and_then(|stem| stem.to_str())
                    .unwrap_or_default()
                    .contains("_审核版")
            {
                files.push(path);
            }
        }
        Ok(())
    }

    let mut files = Vec::new();
    visit(input_dir, recursive, &mut files)?;
    files.sort();
    if recursive {
        let nested = files
            .iter()
            .filter(|path| path.parent().is_some_and(|parent| parent != input_dir))
            .cloned()
            .collect::<Vec<_>>();
        if !nested.is_empty() {
            return Ok(nested);
        }
    }
    Ok(files)
}

fn decode_sheet_values(sheet: &SheetDump) -> Result<BTreeMap<(usize, usize), CellValue>> {
    sheet
        .values
        .iter()
        .map(|(address, value)| {
            let coordinate = parse_cell_address(address)
                .with_context(|| format!("解析 S1-F02 工作表“{}”的单元格地址失败", sheet.name))?;
            let cell = CellValue::from_probe_json(value)
                .with_context(|| format!("读取 S1-F02 工作表“{}”的单元格类型失败", sheet.name))?;
            ensure!(
                !matches!(
                    &cell,
                    CellValue::Error(_) | CellValue::DateTimeIso(_) | CellValue::DurationIso(_)
                ),
                "工作表“{}”含无法按 Excel Value2 契约保真的单元格类型",
                sheet.name
            );
            ensure!(
                !matches!(
                    &cell,
                    CellValue::Number(number) | CellValue::DateTimeSerial(number)
                        if !number.is_finite()
                ),
                "工作表“{}”含非有限数字，无法生成 S1-F02 汇总",
                sheet.name
            );
            Ok((coordinate, cell))
        })
        .collect()
}

fn nonempty_rows(
    first_row: usize,
    last_row: usize,
    first_column: usize,
    last_column: usize,
    values: &BTreeMap<(usize, usize), CellValue>,
) -> Result<Vec<Vec<CellValue>>> {
    let width = last_column - first_column + 1;
    let mut rows = Vec::new();
    for row_number in first_row..=last_row {
        let row = (first_column..=last_column)
            .map(|column| {
                values
                    .get(&(row_number, column))
                    .cloned()
                    .unwrap_or(CellValue::Empty)
            })
            .collect::<Vec<_>>();
        ensure!(row.len() == width, "S1-F02 内部表格行宽不一致");
        if row.iter().any(|value| !is_empty_value(value)) {
            rows.push(row);
        }
    }
    Ok(rows)
}

fn header_signature_value(value: &CellValue) -> Result<String> {
    Ok(match value {
        CellValue::Empty => String::new(),
        CellValue::Text(value) => value.trim().to_owned(),
        CellValue::Number(value) | CellValue::DateTimeSerial(value) => python_float(*value),
        CellValue::Integer(value) => python_float(*value as f64),
        CellValue::Boolean(true) => "True".to_owned(),
        CellValue::Boolean(false) => "False".to_owned(),
        CellValue::Error(_) | CellValue::DateTimeIso(_) | CellValue::DurationIso(_) => {
            bail!("S1-F02 表头含无法按 Excel Value2 契约比较的类型")
        }
    })
}

fn python_float(value: f64) -> String {
    let mut result = value.to_string();
    if !result.contains('.') && !result.contains('e') && !result.contains('E') {
        result.push_str(".0");
    }
    result
}

fn write_row(
    worksheet: &mut excel_writer::Worksheet,
    row_number: u32,
    values: &[CellValue],
    format: &Format,
) -> Result<()> {
    for (column, value) in values.iter().enumerate() {
        let column = column as u16;
        match value {
            CellValue::Empty => {
                worksheet.write_blank(row_number, column, format)?;
            }
            CellValue::Text(value) => {
                worksheet.write_string_with_format(row_number, column, value, format)?;
            }
            CellValue::Number(value) | CellValue::DateTimeSerial(value) => {
                worksheet.write_number_with_format(row_number, column, *value, format)?;
            }
            CellValue::Integer(value) => {
                worksheet.write_number_with_format(row_number, column, *value as f64, format)?;
            }
            CellValue::Boolean(value) => {
                worksheet.write_boolean_with_format(row_number, column, *value, format)?;
            }
            CellValue::Error(_) | CellValue::DateTimeIso(_) | CellValue::DurationIso(_) => {
                bail!("S1-F02 汇总遇到无法按 Excel Value2 契约写入的单元格类型")
            }
        }
    }
    Ok(())
}

fn validate_summary_workbook(path: &Path, groups: &[SummaryGroup]) -> Result<()> {
    let actual = read_workbook(path).context("重新打开 S1-F02 汇总工作簿失败")?;
    let expected_names = if groups.is_empty() {
        vec!["Sheet1".to_owned()]
    } else {
        groups
            .iter()
            .rev()
            .map(|group| group.output_name.clone())
            .collect()
    };
    let actual_names = actual
        .sheets
        .iter()
        .map(|sheet| sheet.name.clone())
        .collect::<Vec<_>>();
    ensure!(
        actual_names == expected_names,
        "S1-F02 输出工作表名称或顺序不匹配"
    );

    for (index, group) in groups.iter().rev().enumerate() {
        let sheet = &actual.sheets[index];
        let mut expected = BTreeMap::new();
        let mut headers = vec![
            CellValue::Text("来源文件".to_owned()),
            CellValue::Text("来源工作表".to_owned()),
        ];
        headers.extend(group.headers.iter().cloned());
        for (column, value) in headers.iter().enumerate() {
            if !is_empty_value(value) {
                expected.insert((1, column + 1), value.clone());
            }
        }
        for (row_index, row) in group.rows.iter().enumerate() {
            for (column, value) in row.iter().enumerate() {
                if !is_empty_value(value) {
                    expected.insert((row_index + 2, column + 1), value.clone());
                }
            }
        }
        let actual_values = sheet
            .values
            .iter()
            .map(|(address, value)| {
                Ok((
                    parse_cell_address(address)?,
                    CellValue::from_probe_json(value)?,
                ))
            })
            .collect::<Result<BTreeMap<_, _>>>()?;
        ensure!(
            expected.len() == actual_values.len(),
            "S1-F02 输出工作表“{}”的非空单元格数量不匹配",
            group.output_name
        );
        for (coordinate, expected_value) in expected {
            let Some(actual_value) = actual_values.get(&coordinate) else {
                bail!("S1-F02 输出工作表“{}”缺少单元格", group.output_name);
            };
            ensure!(
                values_equal(&expected_value, actual_value),
                "S1-F02 输出工作表“{}”的单元格值或类型不匹配",
                group.output_name
            );
        }
        ensure!(sheet.formulas.is_empty(), "S1-F02 输出意外包含公式");
    }
    Ok(())
}

fn values_equal(expected: &CellValue, actual: &CellValue) -> bool {
    match (expected, actual) {
        (CellValue::Empty, CellValue::Empty) => true,
        (CellValue::Text(left), CellValue::Text(right)) => left == right,
        (CellValue::Boolean(left), CellValue::Boolean(right)) => left == right,
        (CellValue::Integer(left), CellValue::Integer(right)) => left == right,
        (left, right) => match (numeric_value(left), numeric_value(right)) {
            (Some(left), Some(right)) => floats_equal(left, right),
            _ => false,
        },
    }
}

fn numeric_value(value: &CellValue) -> Option<f64> {
    match value {
        CellValue::Number(value) | CellValue::DateTimeSerial(value) => Some(*value),
        CellValue::Integer(value) => Some(*value as f64),
        _ => None,
    }
}

fn is_empty_value(value: &CellValue) -> bool {
    matches!(value, CellValue::Empty) || matches!(value, CellValue::Text(text) if text.is_empty())
}

fn floats_equal(left: f64, right: f64) -> bool {
    left == right || (left - right).abs() <= left.abs().max(right.abs()).max(1.0) * 1e-15
}

fn parse_cell_address(address: &str) -> Result<(usize, usize)> {
    let address = address.trim().trim_start_matches('$');
    let mut column = 0_usize;
    let mut row = 0_usize;
    let mut saw_row = false;
    for character in address.chars() {
        if character.is_ascii_alphabetic() && !saw_row {
            column = column
                .checked_mul(26)
                .and_then(|value| {
                    value.checked_add(character.to_ascii_uppercase() as usize - 'A' as usize + 1)
                })
                .context("单元格列号无效")?;
        } else if character.is_ascii_digit() {
            saw_row = true;
            row = row
                .checked_mul(10)
                .and_then(|value| value.checked_add(character.to_digit(10).unwrap() as usize))
                .context("单元格行号无效")?;
        } else {
            bail!("单元格地址格式无效");
        }
    }
    ensure!(row > 0 && column > 0, "单元格地址缺少行号或列号");
    Ok((row, column))
}

fn output_sheet_name(
    feature_name: &str,
    source_sheets: &[String],
    used: &mut HashSet<String>,
) -> String {
    let base_source = source_sheets
        .first()
        .map(String::as_str)
        .unwrap_or(feature_name);
    let mut base = base_source
        .chars()
        .map(|character| {
            if "[]:*?/\\".contains(character) {
                '_'
            } else {
                character
            }
        })
        .collect::<String>()
        .trim()
        .chars()
        .take(31)
        .collect::<String>();
    if base.is_empty() {
        base = "汇总".to_owned();
    }
    let mut candidate = base.clone();
    let mut suffix = 2_usize;
    while !used.insert(candidate.to_lowercase()) {
        let marker = format!("_{suffix}");
        let keep = 31_usize.saturating_sub(marker.chars().count());
        candidate = format!("{}{}", base.chars().take(keep).collect::<String>(), marker);
        suffix += 1;
    }
    candidate
}

fn absolute_path(path: &Path) -> Result<PathBuf> {
    if path.is_absolute() {
        Ok(path.to_path_buf())
    } else {
        Ok(std::env::current_dir()?.join(path))
    }
}

fn temporary_path(output_path: &Path) -> Result<PathBuf> {
    let parent = output_path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or(Path::new("."));
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .context("读取系统时间失败")?
        .as_nanos();
    Ok(parent.join(format!(
        ".{}.{}-{nonce}.tmp.xlsx",
        output_path
            .file_stem()
            .and_then(|stem| stem.to_str())
            .unwrap_or("s1-summary"),
        std::process::id()
    )))
}
