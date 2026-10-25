//! Export user-selected S3 comparison rows to the eight-column explanation
//! workbook used for institution feedback.
//!
//! The source comparison workbook is read through the shared Calamine reader.
//! This module only creates a new workbook; it never modifies the comparison
//! file or the optional institution/region reference workbook.

use anyhow::{bail, ensure, Context, Result};
use chrono::Local;
use excel_reader::{read_value_columns, read_workbook, CellValue, SheetValues};
use excel_writer::{Color, ExcelDateTime, Format, FormatAlign, Workbook, Worksheet};
use serde::Serialize;
use std::{
    collections::{BTreeMap, HashMap, HashSet},
    fs::{self, OpenOptions},
    io::ErrorKind,
    path::{Path, PathBuf},
    sync::atomic::{AtomicU64, Ordering},
    time::Instant,
};

const REQUIRED_HEADERS: [&str; 10] = [
    "机构类代码",
    "机构类名称",
    "地区代码",
    "地区名称",
    "指标代码",
    "指标名称",
    "数据值",
    "上期值",
    "输出说明",
    "说明内容",
];
const SELECTED_VALUES: [&str; 7] = ["是", "y", "yes", "true", "1", "√", "✓"];
const ORG_REFERENCE_SHEET: &str = "机构地区参照";
const ENGINE_READER: &str = "calamine";
const ENGINE_WRITER: &str = "rust_xlsxwriter";
const TITLE: &str = "xx说明";
const CHANGE_REASON_RULE: &str = "比较结果“输出说明”列填写“是”";
const BODY_HEADERS: [&str; 8] = [
    "机构代码",
    "机构名称",
    "指标代码",
    "指标名称",
    "数据值",
    "上期值",
    "增减额",
    "变动原因",
];
const COLUMN_WIDTHS: [u32; 8] = [14, 26, 14, 36, 14, 14, 14, 48];

static TEMP_COUNTER: AtomicU64 = AtomicU64::new(0);

#[derive(Clone, Debug, Serialize)]
pub struct S3ExplanationExportReport {
    pub status: &'static str,
    pub output_workbook: PathBuf,
    pub selected_rows: usize,
    pub reader: &'static str,
    pub writer: &'static str,
    pub comparison_read_time_ms: u128,
    pub config_read_time_ms: u128,
    pub workbook_write_time_ms: u128,
    pub total_time_ms: u128,
}

#[derive(Clone, Debug)]
struct ExplanationRow {
    values: Vec<CellValue>,
}

/// Export rows marked in the comparison result's `输出说明` column.
///
/// The layout and filtering rules follow Python's
/// `s3_central_statistics.explanation_exporter.export_selected_explanations`.
/// A supplied output path must be new; the input workbooks are never changed.
pub fn export_selected_explanations(
    comparison_file: impl AsRef<Path>,
    output_path: impl AsRef<Path>,
    config_path: Option<&Path>,
) -> Result<S3ExplanationExportReport> {
    let started = Instant::now();
    let comparison_file = comparison_file.as_ref();
    let output_path = output_path.as_ref();
    ensure_supported_input(comparison_file)?;
    ensure!(
        output_path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx")),
        "S3-F02 说明文件输出必须使用 .xlsx"
    );
    ensure!(
        !output_path.exists(),
        "说明文件输出路径已存在，未覆盖：{}",
        output_path.display()
    );

    let comparison_started = Instant::now();
    let sheets = read_value_columns(comparison_file, usize::MAX)
        .with_context(|| format!("读取比较结果：{}", comparison_file.display()))?;
    let comparison_read_time_ms = comparison_started.elapsed().as_millis();
    let source_sheet = sheets
        .iter()
        .find(|sheet| header_index(sheet).contains_key("指标代码"))
        .context("比较结果中未找到包含“指标代码”的结果工作表")?;
    let headers = header_index(source_sheet);
    let missing = REQUIRED_HEADERS
        .iter()
        .filter(|header| !headers.contains_key(**header))
        .copied()
        .collect::<Vec<_>>();
    ensure!(
        missing.is_empty(),
        "比较结果缺少说明导出所需列：{}",
        missing.join("、")
    );

    let unit_column = headers["数据值"];
    let unit_header = python_string_or_empty(&cell(source_sheet, 1, unit_column));
    let unit_header = if unit_header.is_empty() {
        "本期金额".to_owned()
    } else {
        unit_header
    };
    let change_header = header_names_in_order(source_sheet)
        .into_iter()
        .find(|name| name.starts_with("增减额"))
        .unwrap_or_else(|| "增减额".to_owned());
    let change_column = headers.get(&change_header).copied();

    let (org_reference, config_read_time_ms) = load_org_reference(config_path)?;
    let mut regions_by_org = HashMap::<String, HashSet<String>>::new();
    for row_number in 2..=source_sheet.max_row {
        let org_code = clean_code(&cell(source_sheet, row_number, headers["机构类代码"]));
        let region_code = clean_code(&cell(source_sheet, row_number, headers["地区代码"]));
        if !org_code.is_empty() && !region_code.is_empty() {
            regions_by_org
                .entry(org_code)
                .or_default()
                .insert(region_code);
        }
    }

    let mut rows = Vec::new();
    for row_number in 2..=source_sheet.max_row {
        let selected = cell(source_sheet, row_number, headers["输出说明"]);
        if !is_selected(&selected) {
            continue;
        }
        let Some(change_column) = change_column else {
            bail!("比较结果缺少增减额列，无法导出已选说明行");
        };

        let org_code_value = cell(source_sheet, row_number, headers["机构类代码"]);
        let org_code = clean_code(&org_code_value);
        let org_name_value = cell(source_sheet, row_number, headers["机构类名称"]);
        let org_name = python_string_or_empty(&org_name_value).trim().to_owned();
        let region_code = clean_code(&cell(source_sheet, row_number, headers["地区代码"]));
        let region_name =
            python_string_or_empty(&cell(source_sheet, row_number, headers["地区名称"]))
                .trim()
                .to_owned();

        let base_name = org_reference
            .get(&(org_code.clone(), region_code.clone()))
            .or_else(|| org_reference.get(&(org_code.clone(), "*".to_owned())))
            .cloned()
            .unwrap_or_else(|| org_name.clone());
        let display_name = if region_name.is_empty() {
            base_name
        } else {
            format!("{base_name}（{region_name}）")
        };
        let is_village_bank = org_code.starts_with("7020") || org_name.contains("村镇银行");
        let multi_region = regions_by_org
            .get(&org_code)
            .is_some_and(|regions| regions.len() > 1);
        let display_code = if !region_code.is_empty() && (is_village_bank || multi_region) {
            format!("{org_code}{region_code}")
        } else {
            org_code
        };

        rows.push(ExplanationRow {
            values: vec![
                CellValue::Text(display_code),
                CellValue::Text(display_name),
                cell(source_sheet, row_number, headers["指标代码"]),
                cell(source_sheet, row_number, headers["指标名称"]),
                cell(source_sheet, row_number, headers["数据值"]),
                cell(source_sheet, row_number, headers["上期值"]),
                cell(source_sheet, row_number, change_column),
                cell(source_sheet, row_number, headers["说明内容"]),
            ],
        });
    }

    let output_headers = [
        BODY_HEADERS[0].to_owned(),
        BODY_HEADERS[1].to_owned(),
        BODY_HEADERS[2].to_owned(),
        BODY_HEADERS[3].to_owned(),
        unit_header,
        BODY_HEADERS[5].to_owned(),
        change_header,
        BODY_HEADERS[7].to_owned(),
    ];
    let output_parent = output_path
        .parent()
        .filter(|path| !path.as_os_str().is_empty());
    if let Some(parent) = output_parent {
        fs::create_dir_all(parent)
            .with_context(|| format!("创建说明文件目录：{}", parent.display()))?;
    }

    let workbook_write_started = Instant::now();
    let temporary_path = reserve_temporary_path(output_path)?;
    let write_result = write_workbook(&temporary_path, comparison_file, &output_headers, &rows)
        .and_then(|()| validate_output(&temporary_path, rows.len()));
    if let Err(error) = write_result {
        let _ = fs::remove_file(&temporary_path);
        return Err(error);
    }
    // A hard-link publish is atomic on the same filesystem and fails if another
    // process created the requested output path after the initial existence check.
    if let Err(error) = fs::hard_link(&temporary_path, output_path) {
        let _ = fs::remove_file(&temporary_path);
        return Err(error).with_context(|| format!("发布说明文件失败：{}", output_path.display()));
    }
    fs::remove_file(&temporary_path)
        .with_context(|| format!("清理临时说明文件：{}", temporary_path.display()))?;
    let workbook_write_time_ms = workbook_write_started.elapsed().as_millis();

    Ok(S3ExplanationExportReport {
        status: "completed",
        output_workbook: output_path.to_path_buf(),
        selected_rows: rows.len(),
        reader: ENGINE_READER,
        writer: ENGINE_WRITER,
        comparison_read_time_ms,
        config_read_time_ms,
        workbook_write_time_ms,
        total_time_ms: started.elapsed().as_millis(),
    })
}

fn ensure_supported_input(path: &Path) -> Result<()> {
    ensure!(path.is_file(), "比较结果文件不存在：{}", path.display());
    ensure!(
        path.extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| {
                extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm")
            }),
        "S3-F02 比较结果目前只支持 .xlsx 和 .xlsm"
    );
    Ok(())
}

fn header_index(sheet: &SheetValues) -> BTreeMap<String, usize> {
    let mut headers = BTreeMap::new();
    for column in 1..=sheet.max_column {
        let value = cell(sheet, 1, column);
        if matches!(&value, CellValue::Empty)
            || matches!(&value, CellValue::Text(text) if text.is_empty())
        {
            continue;
        }
        let name = python_string(&value).trim().to_owned();
        if !name.is_empty() {
            headers.insert(name, column);
        }
    }
    headers
}

fn header_names_in_order(sheet: &SheetValues) -> Vec<String> {
    let mut seen = HashSet::new();
    let mut ordered = Vec::new();
    for column in 1..=sheet.max_column {
        let value = cell(sheet, 1, column);
        if matches!(&value, CellValue::Empty)
            || matches!(&value, CellValue::Text(text) if text.is_empty())
        {
            continue;
        }
        let name = python_string(&value).trim().to_owned();
        if !name.is_empty() && seen.insert(name.clone()) {
            ordered.push(name);
        }
    }
    ordered
}

fn load_org_reference(
    config_path: Option<&Path>,
) -> Result<(HashMap<(String, String), String>, u128)> {
    let Some(config_path) = config_path.filter(|path| path.is_file()) else {
        return Ok((HashMap::new(), 0));
    };
    let started = Instant::now();
    let sheets = read_value_columns(config_path, usize::MAX)
        .with_context(|| format!("读取机构地区参照配置：{}", config_path.display()))?;
    let elapsed_ms = started.elapsed().as_millis();
    let Some(sheet) = sheets
        .iter()
        .find(|sheet| sheet.name == ORG_REFERENCE_SHEET)
    else {
        return Ok((HashMap::new(), elapsed_ms));
    };
    let headers = header_index(sheet);
    let required = ["机构类代码", "地区代码", "机构名称"];
    if required.iter().any(|name| !headers.contains_key(*name)) {
        return Ok((HashMap::new(), elapsed_ms));
    }
    let max_required_column = required
        .iter()
        .filter_map(|name| headers.get(*name))
        .copied()
        .max()
        .unwrap_or_default();
    let mut mapping = HashMap::new();
    let mut names_by_org = HashMap::<String, HashSet<String>>::new();
    for row_number in 2..=sheet.max_row {
        if max_required_column > sheet.max_column {
            continue;
        }
        let org_code = clean_code(&cell(sheet, row_number, headers["机构类代码"]));
        let region_code = clean_code(&cell(sheet, row_number, headers["地区代码"]));
        let name = python_string_or_empty(&cell(sheet, row_number, headers["机构名称"]))
            .trim()
            .to_owned();
        if org_code.is_empty() || name.is_empty() {
            continue;
        }
        mapping.insert((org_code.clone(), region_code), name.clone());
        names_by_org.entry(org_code).or_default().insert(name);
    }
    for (org_code, names) in names_by_org {
        if names.len() == 1 {
            if let Some(name) = names.into_iter().next() {
                mapping.entry((org_code, "*".to_owned())).or_insert(name);
            }
        }
    }
    Ok((mapping, elapsed_ms))
}

fn is_selected(value: &CellValue) -> bool {
    let text = python_string_or_empty(value).trim().to_lowercase();
    SELECTED_VALUES.contains(&text.as_str())
}

fn clean_code(value: &CellValue) -> String {
    python_string_or_empty(value)
        .replace('\'', "")
        .trim()
        .to_owned()
}

fn python_string_or_empty(value: &CellValue) -> String {
    if !python_truthy(value) {
        return String::new();
    }
    match value {
        CellValue::Empty => String::new(),
        CellValue::Text(value) => value.clone(),
        CellValue::Number(value) | CellValue::DateTimeSerial(value) => format!("{value:?}"),
        CellValue::Integer(value) => value.to_string(),
        CellValue::Boolean(value) => if *value { "True" } else { "False" }.to_owned(),
        CellValue::DateTimeIso(value) | CellValue::DurationIso(value) => value.clone(),
        CellValue::Error(value) => format_excel_error(value),
    }
}

fn python_string(value: &CellValue) -> String {
    match value {
        CellValue::Empty => String::new(),
        CellValue::Text(value) => value.clone(),
        CellValue::Number(value) | CellValue::DateTimeSerial(value) => format!("{value:?}"),
        CellValue::Integer(value) => value.to_string(),
        CellValue::Boolean(value) => if *value { "True" } else { "False" }.to_owned(),
        CellValue::DateTimeIso(value) | CellValue::DurationIso(value) => value.clone(),
        CellValue::Error(value) => format_excel_error(value),
    }
}

fn python_truthy(value: &CellValue) -> bool {
    match value {
        CellValue::Empty => false,
        CellValue::Text(value) => !value.is_empty(),
        CellValue::Number(value) | CellValue::DateTimeSerial(value) => *value != 0.0,
        CellValue::Integer(value) => *value != 0,
        CellValue::Boolean(value) => *value,
        CellValue::DateTimeIso(value) | CellValue::DurationIso(value) => !value.is_empty(),
        CellValue::Error(value) => !value.is_empty(),
    }
}

fn format_excel_error(value: &str) -> String {
    let normalized = value
        .trim()
        .trim_start_matches('#')
        .trim_end_matches(['!', '?'])
        .replace('_', "")
        .replace('-', "")
        .replace(' ', "")
        .to_ascii_lowercase();
    match normalized.as_str() {
        "ref" => "#REF!",
        "value" => "#VALUE!",
        "na" | "notavailable" => "#N/A",
        "div0" | "divzero" => "#DIV/0!",
        "name" => "#NAME?",
        "num" => "#NUM!",
        "null" => "#NULL!",
        "spill" => "#SPILL!",
        "calc" => "#CALC!",
        _ => value,
    }
    .to_owned()
}

fn cell(sheet: &SheetValues, row: usize, column: usize) -> CellValue {
    sheet
        .values
        .get(&cell_address(row, column))
        .cloned()
        .unwrap_or(CellValue::Empty)
}

fn cell_address(row: usize, column: usize) -> String {
    let mut column = column;
    let mut label = String::new();
    while column > 0 {
        let remainder = (column - 1) % 26;
        label.insert(0, char::from(b'A' + remainder as u8));
        column = (column - 1) / 26;
    }
    format!("{label}{row}")
}

fn reserve_temporary_path(output: &Path) -> Result<PathBuf> {
    let parent = output
        .parent()
        .filter(|path| !path.as_os_str().is_empty())
        .unwrap_or(Path::new("."));
    let stem = output
        .file_stem()
        .and_then(|value| value.to_str())
        .unwrap_or("explanation");
    for _ in 0..100 {
        let counter = TEMP_COUNTER.fetch_add(1, Ordering::Relaxed);
        let candidate = parent.join(format!(
            ".{stem}.{}.{}.tmp.xlsx",
            std::process::id(),
            counter
        ));
        match OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&candidate)
        {
            Ok(file) => {
                drop(file);
                return Ok(candidate);
            }
            Err(error) if error.kind() == ErrorKind::AlreadyExists => continue,
            Err(error) => {
                return Err(error)
                    .with_context(|| format!("创建说明文件临时路径：{}", candidate.display()));
            }
        }
    }
    bail!("无法保留唯一的说明文件临时路径")
}

fn write_workbook(
    path: &Path,
    comparison_file: &Path,
    headers: &[String; 8],
    rows: &[ExplanationRow],
) -> Result<()> {
    let mut workbook = Workbook::new();
    let title_format = Format::new()
        .set_bold()
        .set_font_size(14.0)
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::VerticalCenter);
    let header_format = Format::new()
        .set_bold()
        .set_font_color(Color::RGB(0xFF_FF_FF))
        .set_background_color(Color::RGB(0x1F_4E_78))
        .set_align(FormatAlign::VerticalCenter)
        .set_text_wrap();
    let info_header_format = Format::new()
        .set_bold()
        .set_font_color(Color::RGB(0xFF_FF_FF))
        .set_background_color(Color::RGB(0x1F_4E_78));
    let amount_format = Format::new().set_num_format("0.00");
    let date_format = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss");

    let explanation = workbook.add_worksheet();
    explanation
        .set_name("说明文件")
        .context("设置说明工作表名称失败")?;
    explanation
        .merge_range(0, 0, 0, 7, TITLE, &title_format)
        .context("合并说明文件标题单元格失败")?;
    for (column, header) in headers.iter().enumerate() {
        explanation.write_string_with_format(1, column as u16, header, &header_format)?;
    }
    for (index, row) in rows.iter().enumerate() {
        let output_row = index as u32 + 2;
        for (column, value) in row.values.iter().enumerate() {
            let number_format = (4..=6).contains(&column).then_some(&amount_format);
            write_cell(
                explanation,
                output_row,
                column as u16,
                value,
                number_format,
                &date_format,
            )?;
        }
    }
    explanation
        .set_freeze_panes(2, 0)
        .context("冻结说明文件前两行失败")?;
    let last_filter_row = rows.len().saturating_add(1).max(2) as u32;
    explanation
        .autofilter(1, 0, last_filter_row, 7)
        .context("设置说明文件筛选范围失败")?;
    for (column, width) in COLUMN_WIDTHS.iter().enumerate() {
        explanation
            .set_column_width_pixels(column as u16, width * 7)
            .with_context(|| format!("设置说明文件第 {} 列宽度失败", column + 1))?;
    }

    let info = workbook.add_worksheet();
    info.set_name("生成信息")
        .context("设置生成信息工作表名称失败")?;
    info.write_string_with_format(0, 0, "项目", &info_header_format)?;
    info.write_string_with_format(0, 1, "内容", &info_header_format)?;
    let exported_at = Local::now().format("%Y-%m-%d %H:%M:%S").to_string();
    info.write_string(1, 0, "来源比较文件")?;
    info.write_string(1, 1, comparison_file.to_string_lossy().into_owned())?;
    info.write_string(2, 0, "导出时间")?;
    info.write_string(2, 1, &exported_at)?;
    info.write_string(3, 0, "选定指标数")?;
    info.write_number(3, 1, rows.len() as f64)?;
    info.write_string(4, 0, "选取规则")?;
    info.write_string(4, 1, CHANGE_REASON_RULE)?;
    info.set_column_width_pixels(0, 18 * 7)?;
    info.set_column_width_pixels(1, 90 * 7)?;
    workbook.save(path).context("保存 S3-F02 说明 XLSX 失败")
}

fn write_cell(
    worksheet: &mut Worksheet,
    row: u32,
    column: u16,
    value: &CellValue,
    amount_format: Option<&Format>,
    date_format: &Format,
) -> Result<()> {
    match value {
        CellValue::Empty => {
            worksheet.write_blank(row, column, &Format::new())?;
        }
        CellValue::Text(value) => {
            worksheet.write_string(row, column, value)?;
        }
        CellValue::Number(value) if amount_format.is_some() => {
            worksheet.write_number_with_format(row, column, *value, amount_format.unwrap())?;
        }
        CellValue::Number(value) => {
            worksheet.write_number(row, column, *value)?;
        }
        CellValue::Integer(value) => {
            worksheet.write_number(row, column, *value as f64)?;
        }
        CellValue::Boolean(value) => {
            worksheet.write_boolean(row, column, *value)?;
        }
        CellValue::DateTimeSerial(value) => {
            worksheet.write_number_with_format(row, column, *value, date_format)?;
        }
        CellValue::DateTimeIso(value) => {
            let date = ExcelDateTime::parse_from_str(value)
                .with_context(|| format!("无法写入说明行日期值：{value}"))?;
            worksheet.write_datetime_with_format(row, column, date, date_format)?;
        }
        CellValue::DurationIso(value) => {
            worksheet.write_string(row, column, value)?;
        }
        CellValue::Error(value) => {
            worksheet.write_string(row, column, &format_excel_error(value))?;
        }
    }
    Ok(())
}

fn validate_output(path: &Path, selected_rows: usize) -> Result<()> {
    let sheets = read_value_columns(path, 8).context("独立读回 S3-F02 说明工作簿失败")?;
    ensure!(
        sheets
            .iter()
            .map(|sheet| sheet.name.as_str())
            .collect::<Vec<_>>()
            == ["说明文件", "生成信息"],
        "S3-F02 说明工作簿工作表名称或顺序错误"
    );
    let explanation = &sheets[0];
    ensure!(
        explanation.max_row == selected_rows + 2,
        "S3-F02 说明工作簿行数校验失败：期望 {}，实际 {}",
        selected_rows + 2,
        explanation.max_row
    );
    ensure!(explanation.max_column == 8, "S3-F02 说明工作簿列数不是 8");
    ensure!(
        cell(explanation, 1, 1) == CellValue::Text(TITLE.to_owned()),
        "S3-F02 说明工作簿标题读回不一致"
    );
    let metadata = &sheets[1];
    ensure!(
        metadata.max_row == 5 && metadata.max_column == 2,
        "S3-F02 生成信息工作表结构错误"
    );
    let package = read_workbook(path).context("验证 S3-F02 说明工作簿合并区域失败")?;
    ensure!(
        package.sheets.len() == 2,
        "S3-F02 说明工作簿工作表数量不是 2"
    );
    ensure!(
        package.sheets[0].merges.iter().any(|area| area == "A1:H1"),
        "S3-F02 说明工作簿缺少 A1:H1 标题合并"
    );
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::{cell_address, is_selected, python_string_or_empty};
    use excel_reader::CellValue;

    #[test]
    fn a1_column_addresses_cover_excel_column_boundaries() {
        assert_eq!(cell_address(1, 1), "A1");
        assert_eq!(cell_address(27, 26), "Z27");
        assert_eq!(cell_address(10, 27), "AA10");
    }

    #[test]
    fn explanation_markers_follow_python_text_coercion() {
        for value in [
            CellValue::Text("是".to_owned()),
            CellValue::Text(" YES ".to_owned()),
            CellValue::Text("√".to_owned()),
            CellValue::Text("✓".to_owned()),
            CellValue::Integer(1),
            CellValue::Boolean(true),
        ] {
            assert!(is_selected(&value));
        }
        for value in [
            CellValue::Text("否".to_owned()),
            CellValue::Text("1.0".to_owned()),
            CellValue::Integer(0),
            CellValue::Boolean(false),
            CellValue::Empty,
        ] {
            assert!(!is_selected(&value));
        }
        assert_eq!(python_string_or_empty(&CellValue::Number(1.0)), "1.0");
    }
}
