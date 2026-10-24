//! S1-F02 named-range summaries for formula-free XLSX input.
//!
//! This migrates the ordinary `任意行汇总` and `固定行汇总` value path. It
//! intentionally rejects fallback-template, extra metadata, history-enrichment,
//! and formula-bearing cases until each behavior has an Audit-specific parity
//! contract.

use crate::{
    name_config::{
        load_s1_region_summary_mappings, resolve_named_ranges_for_mappings, NamedRangeArea,
        FIXED_ROW_SUMMARY_FUNCTION, USED_RANGE_SUMMARY_FUNCTION,
    },
    template::normalize_template_name,
};
use anyhow::{bail, ensure, Context, Result};
use excel_reader::{
    read_defined_names, read_workbook, read_xlsx_worksheet_extents, CellValue,
    DefinedNameParseStatus, DefinedNameRecord, SheetDump, WorkbookDump, WorksheetExtent,
};
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
pub struct S1RegionSummarySheetSummary {
    pub feature_name: String,
    pub sheet_name: String,
    pub row_count: usize,
    pub column_count: usize,
}

#[derive(Clone, Debug, Serialize)]
pub struct S1RegionSummarySummary {
    pub output_path: PathBuf,
    pub features: usize,
    pub source_workbooks: usize,
    pub source_sheets: usize,
    pub output_sheets: usize,
    pub data_rows: usize,
    pub skipped_missing_sheets: usize,
    pub metadata_fields_left_blank: usize,
    pub reader: &'static str,
    pub writer: &'static str,
    pub discovery_time_ms: u128,
    pub template_read_time_ms: u128,
    pub source_read_time_ms: u128,
    pub calculation_time_ms: u128,
    pub write_time_ms: u128,
    pub validation_time_ms: u128,
    pub elapsed_ms: u128,
    pub sheets: Vec<S1RegionSummarySheetSummary>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum SummaryMode {
    UsedRange,
    FixedRows,
}

#[derive(Clone, Debug)]
struct RegionGroup {
    feature_name: String,
    template_sheet_name: String,
    output_sheet_name: String,
    extra_fields: Vec<SummaryField>,
    headers: Vec<String>,
    rows: Vec<Vec<CellValue>>,
}

#[derive(Clone, Debug)]
struct SummaryField {
    name: String,
    range: NamedRangeArea,
    bounds: Bounds,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
struct Bounds {
    top: usize,
    left: usize,
    bottom: usize,
    right: usize,
}

/// Summarize configured named regions into a new XLSX workbook.
///
/// `mapping_config` is read-only. A missing mapping workbook uses the source
/// project's built-in `任意行汇总区域` / `固定行汇总区域` mappings. Output is
/// written to a temporary workbook, validated by the Rust reader, and only
/// then moved to the requested new path.
pub fn summarize_s1_named_regions(
    template_path: impl AsRef<Path>,
    input_dir: impl AsRef<Path>,
    output_path: impl AsRef<Path>,
    mapping_config: Option<&Path>,
    feature_names: &[String],
    recursive: bool,
) -> Result<S1RegionSummarySummary> {
    let started = Instant::now();
    let template_path = fs::canonicalize(template_path.as_ref()).with_context(|| {
        format!(
            "读取 S1-F02 汇总模板失败：{}",
            template_path.as_ref().display()
        )
    })?;
    ensure!(template_path.is_file(), "S1-F02 汇总模板路径不是文件");
    ensure!(
        template_path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx")),
        "S1-F02 命名区域汇总仅支持 .xlsx 模板"
    );
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

    let template_started = Instant::now();
    let template = read_workbook(&template_path)
        .with_context(|| format!("读取 S1-F02 汇总模板值失败：{}", template_path.display()))?;
    let defined_names = read_defined_names(&template_path)
        .with_context(|| format!("读取 S1-F02 模板命名区域失败：{}", template_path.display()))?;
    let feature_mappings = load_s1_region_summary_mappings(mapping_config, &template_path)?;
    let selected_mappings = feature_mappings
        .into_iter()
        .filter(|mapping| {
            feature_names.is_empty() || feature_names.iter().any(|name| name == &mapping.name)
        })
        .collect::<Vec<_>>();
    ensure!(
        !selected_mappings.is_empty(),
        "配置中没有匹配的 S1-F02 任意行/固定行汇总功能"
    );
    ensure!(
        !template
            .sheets
            .iter()
            .any(|sheet| sheet.name.eq_ignore_ascii_case("兜底模板")),
        "S1-F02 模板含“兜底模板”工作表；当前 Rust 命名区域汇总尚未迁移兜底汇总"
    );
    let areas = crate::name_config::named_areas_from_defined_names(&defined_names);
    let (global_summary_fields, local_summary_fields) = parse_summary_fields(&defined_names)?;
    let template_sheet_names = template
        .sheets
        .iter()
        .map(|sheet| sheet.name.clone())
        .collect::<Vec<_>>();
    let template_sheets = template
        .sheets
        .iter()
        .map(|sheet| (sheet.name.to_lowercase(), sheet))
        .collect::<HashMap<_, _>>();
    let template_read_time_ms = template_started.elapsed().as_millis();

    let mut source_workbooks =
        Vec::<(&PathBuf, WorkbookDump, Vec<WorksheetExtent>)>::with_capacity(sources.len());
    let mut source_read_time_ms = 0_u128;
    let mut source_sheet_count = 0_usize;
    for source_path in &sources {
        let read_started = Instant::now();
        let workbook = read_workbook(source_path)
            .with_context(|| format!("读取 S1-F02 汇总源工作簿失败：{}", source_path.display()))?;
        let extents = read_xlsx_worksheet_extents(source_path)
            .with_context(|| format!("读取 S1-F02 源工作表范围失败：{}", source_path.display()))?;
        source_read_time_ms += read_started.elapsed().as_millis();
        source_sheet_count += workbook.sheets.len();
        for sheet in &workbook.sheets {
            ensure!(
                sheet.formula_status == "read_by_calamine",
                "无法确认源工作表“{}”是否含公式，停止 S1-F02 汇总",
                sheet.name
            );
            if !sheet.formulas.is_empty() {
                bail!(
                    "S1-F02 命名区域汇总暂不处理含公式的工作表；工作表“{}”含公式单元格，需先明确本轮计算语义",
                    sheet.name
                );
            }
        }
        source_workbooks.push((source_path, workbook, extents));
    }

    let mut groups = Vec::<RegionGroup>::new();
    let mut feature_count = 0;
    let mut skipped_missing_sheets = 0;
    let mut metadata_fields_left_blank = 0;
    let mut calculation_time_ms = 0_u128;

    for mapping in &selected_mappings {
        let mode = match mapping.feature_type.as_str() {
            USED_RANGE_SUMMARY_FUNCTION => SummaryMode::UsedRange,
            FIXED_ROW_SUMMARY_FUNCTION => SummaryMode::FixedRows,
            _ => continue,
        };
        feature_count += 1;

        let data_ranges = resolve_named_ranges_for_mappings(
            std::slice::from_ref(mapping),
            &areas,
            &template_sheet_names,
        );
        ensure!(
            !data_ranges.is_empty(),
            "汇总功能“{}”未找到命名区域：{}",
            mapping.name,
            mapping.range_names.join("、")
        );

        let mut header_mapping = mapping.clone();
        header_mapping.range_names = vec!["表头区域".to_owned()];
        let header_ranges =
            resolve_named_ranges_for_mappings(&[header_mapping], &areas, &template_sheet_names);
        ensure!(
            !header_ranges.is_empty(),
            "汇总功能“{}”缺少命名区域“表头区域”",
            mapping.name
        );

        let mut feature_group_indices = HashMap::<String, usize>::new();
        for data_range in &data_ranges {
            let data_bounds = parse_bounds(&data_range.address)
                .with_context(|| format!("S1-F02 数据命名区域无效：{}", data_range.address))?;
            ensure!(
                data_bounds.bottom <= MAX_ROWS && data_bounds.right <= MAX_COLUMNS,
                "S1-F02 数据命名区域超过 XLSX 的最大行列数"
            );
            let mut matching_headers = Vec::new();
            for header in header_ranges.iter().filter(|header| {
                header
                    .sheet_name
                    .eq_ignore_ascii_case(&data_range.sheet_name)
            }) {
                let header_bounds = parse_bounds(&header.address)
                    .with_context(|| format!("S1-F02 表头命名区域无效：{}", header.address))?;
                ensure!(
                    header_bounds.bottom <= MAX_ROWS && header_bounds.right <= MAX_COLUMNS,
                    "S1-F02 表头命名区域超过 XLSX 的最大行列数"
                );
                let left = data_bounds.left.max(header_bounds.left);
                let right = data_bounds.right.min(header_bounds.right);
                if left <= right {
                    matching_headers.push((header, header_bounds, left, right));
                }
            }
            ensure!(
                !matching_headers.is_empty(),
                "汇总功能“{}”在工作表“{}”缺少相交的“表头区域”",
                mapping.name,
                data_range.sheet_name
            );
            let mut chosen = matching_headers[0];
            for candidate in matching_headers.iter().copied().skip(1) {
                if candidate.3 - candidate.2 > chosen.3 - chosen.2 {
                    chosen = candidate;
                }
            }
            let (_, header_bounds, left, right) = chosen;
            let template_sheet = template_sheets
                .get(&data_range.sheet_name.to_lowercase())
                .with_context(|| format!("模板缺少工作表：{}", data_range.sheet_name))?;
            ensure!(
                template_sheet.merge_status == "read_by_calamine",
                "无法确认模板工作表“{}”的合并单元格，停止 S1-F02 汇总",
                template_sheet.name
            );
            let template_grid = decode_sheet_values(template_sheet)?;
            for formula_address in template_sheet.formulas.keys() {
                let (formula_row, formula_column) = parse_cell(formula_address)
                    .with_context(|| format!("模板公式坐标无效：{formula_address}"))?;
                ensure!(
                    formula_row < header_bounds.top
                        || formula_row > header_bounds.bottom
                        || formula_column < left
                        || formula_column > right,
                    "S1-F02 表头区域含公式单元格 {formula_address}，需先明确公式缓存语义"
                );
            }
            let header_values =
                read_header_values(template_sheet, &template_grid, header_bounds, left, right)?;
            let extra_fields = summary_fields_for_sheet(
                &global_summary_fields,
                &local_summary_fields,
                &data_range.sheet_name,
            )?;

            let group_key = data_range.sheet_name.to_lowercase();
            let group_index = if let Some(index) = feature_group_indices.get(&group_key) {
                *index
            } else {
                let index = groups.len();
                feature_group_indices.insert(group_key, index);
                groups.push(RegionGroup {
                    feature_name: mapping.name.clone(),
                    template_sheet_name: data_range.sheet_name.clone(),
                    output_sheet_name: String::new(),
                    extra_fields: extra_fields.clone(),
                    headers: header_values,
                    rows: Vec::new(),
                });
                index
            };
            let group = &mut groups[group_index];
            ensure!(
                group
                    .extra_fields
                    .iter()
                    .map(|field| &field.name)
                    .collect::<Vec<_>>()
                    == extra_fields
                        .iter()
                        .map(|field| &field.name)
                        .collect::<Vec<_>>(),
                "汇总功能“{}”的附加字段定义不一致",
                mapping.name
            );
            ensure!(
                group.headers.len() == right - left + 1,
                "汇总功能“{}”的表头区域列数不一致",
                mapping.name
            );

            let data_top = data_bounds.top;
            let data_bottom = data_bounds.bottom;
            for (source_path, source_workbook, source_extents) in &source_workbooks {
                let Some(source_sheet) = source_workbook
                    .sheets
                    .iter()
                    .find(|sheet| sheet.name.eq_ignore_ascii_case(&data_range.sheet_name))
                else {
                    skipped_missing_sheets += 1;
                    continue;
                };
                let extent = source_extents
                    .iter()
                    .find(|extent| extent.name.eq_ignore_ascii_case(&source_sheet.name));
                let (first_row, last_row) = match mode {
                    SummaryMode::FixedRows => (data_top, data_bottom),
                    SummaryMode::UsedRange => {
                        let Some(extent) = extent else {
                            continue;
                        };
                        if extent.max_row == 0 {
                            continue;
                        }
                        (data_top.max(extent.min_row), extent.max_row)
                    }
                };
                if last_row < first_row {
                    continue;
                }
                ensure!(
                    last_row <= MAX_ROWS && right <= MAX_COLUMNS,
                    "S1-F02 汇总区域超过 XLSX 的最大行列数"
                );

                let calculation_started = Instant::now();
                let source_grid = decode_sheet_values(source_sheet)?;
                let mut source_grids =
                    HashMap::<String, Option<BTreeMap<(usize, usize), CellValue>>>::new();
                source_grids.insert(source_sheet.name.to_lowercase(), Some(source_grid.clone()));
                let mut extra_values = Vec::with_capacity(group.extra_fields.len());
                for field in &group.extra_fields {
                    let metadata_sheet_key = field.range.sheet_name.to_lowercase();
                    if !source_grids.contains_key(&metadata_sheet_key) {
                        let metadata_sheet = source_workbook
                            .sheets
                            .iter()
                            .find(|sheet| sheet.name.eq_ignore_ascii_case(&field.range.sheet_name));
                        let metadata_grid = metadata_sheet.map(decode_sheet_values).transpose()?;
                        source_grids.insert(metadata_sheet_key.clone(), metadata_grid);
                    }
                    match source_grids
                        .get(&metadata_sheet_key)
                        .and_then(Option::as_ref)
                    {
                        Some(metadata_grid) => extra_values.push(cell_value(
                            metadata_grid,
                            field.bounds.top,
                            field.bounds.left,
                        )),
                        None => {
                            metadata_fields_left_blank += 1;
                            extra_values.push(CellValue::Empty);
                        }
                    }
                }
                let source_stem = source_path
                    .file_stem()
                    .and_then(|name| name.to_str())
                    .unwrap_or_default();
                let source_name = normalize_template_name(source_stem);
                for row_number in first_row..=last_row {
                    let row = (left..=right)
                        .map(|column| cell_value(&source_grid, row_number, column))
                        .collect::<Vec<_>>();
                    if row.iter().any(|value| !is_empty_value(value)) {
                        let mut output_row =
                            Vec::with_capacity(row.len() + group.extra_fields.len() + 2);
                        output_row.push(CellValue::Text(source_name.clone()));
                        output_row.push(CellValue::Text(source_sheet.name.clone()));
                        output_row.extend(extra_values.iter().cloned());
                        output_row.extend(row);
                        group.rows.push(output_row);
                    }
                }
                calculation_time_ms += calculation_started.elapsed().as_millis();
            }
        }
    }
    ensure!(feature_count > 0, "配置中没有可执行的任意行/固定行汇总映射");

    let mut used_sheet_names = HashSet::new();
    for group in &mut groups {
        group.output_sheet_name =
            output_sheet_name(&group.template_sheet_name, &mut used_sheet_names);
        ensure!(
            group.rows.len() + 1 <= MAX_ROWS
                && group.headers.len() + group.extra_fields.len() + 2 <= MAX_COLUMNS,
            "S1-F02 汇总输出超过 XLSX 的最大行列数"
        );
    }

    let output_parent = output_path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or(Path::new("."));
    fs::create_dir_all(output_parent)
        .with_context(|| format!("创建 S1-F02 汇总输出目录失败：{}", output_parent.display()))?;
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
            workbook
                .add_worksheet()
                .set_name("汇总")
                .context("设置 S1-F02 空汇总工作表名")?;
        } else {
            // Windows Excel COM serializes the output sheets in reverse
            // creation order for the Python baseline used by this workflow.
            for group in groups.iter().rev() {
                let worksheet = workbook.add_worksheet();
                worksheet
                    .set_name(&group.output_sheet_name)
                    .with_context(|| {
                        format!("设置 S1-F02 输出工作表名失败：{}", group.output_sheet_name)
                    })?;
                let mut headers = vec![
                    CellValue::Text("来源文件".to_owned()),
                    CellValue::Text("来源工作表".to_owned()),
                ];
                headers.extend(
                    group
                        .extra_fields
                        .iter()
                        .map(|field| CellValue::Text(field.name.clone())),
                );
                headers.extend(group.headers.iter().cloned().map(CellValue::Text));
                write_row(worksheet, 0, &headers, &header_format)?;
                for (row_index, row) in group.rows.iter().enumerate() {
                    write_row(worksheet, row_index as u32 + 1, row, &value_format)?;
                }
                worksheet
                    .autofilter(0, 0, group.rows.len() as u32, (headers.len() - 1) as u16)
                    .context("设置 S1-F02 汇总筛选")?;
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
        .map(|group| S1RegionSummarySheetSummary {
            feature_name: group.feature_name.clone(),
            sheet_name: group.output_sheet_name.clone(),
            row_count: group.rows.len(),
            column_count: group.headers.len() + group.extra_fields.len() + 2,
        })
        .collect::<Vec<_>>();
    let data_rows = sheets.iter().map(|sheet| sheet.row_count).sum();
    Ok(S1RegionSummarySummary {
        output_path,
        features: feature_count,
        source_workbooks: sources.len(),
        source_sheets: source_sheet_count,
        output_sheets: sheets.len(),
        data_rows,
        skipped_missing_sheets,
        metadata_fields_left_blank,
        reader: READER_ENGINE,
        writer: WRITER_ENGINE,
        discovery_time_ms,
        template_read_time_ms,
        source_read_time_ms,
        calculation_time_ms,
        write_time_ms,
        validation_time_ms,
        elapsed_ms: started.elapsed().as_millis(),
        sheets,
    })
}

fn parse_summary_fields(
    records: &[DefinedNameRecord],
) -> Result<(Vec<SummaryField>, HashMap<String, Vec<SummaryField>>)> {
    let mut global_fields = Vec::new();
    let mut local_fields = HashMap::<String, Vec<SummaryField>>::new();
    for record in records {
        let (prefix, is_global) = if record.name.starts_with("全局单元格区域.") {
            ("全局单元格区域.", true)
        } else if record.name.starts_with("单元格区域.") {
            ("单元格区域.", false)
        } else {
            continue;
        };
        let field_name = record.name[prefix.len()..].trim();
        if field_name.is_empty() {
            continue;
        }
        ensure!(
            matches!(&record.parse_status, DefinedNameParseStatus::Areas),
            "S1-F02 附加字段命名区域“{}”无法解析为工作表单元格：{:?}",
            record.name,
            record.parse_status
        );
        let areas =
            crate::name_config::named_areas_from_defined_names(std::slice::from_ref(record));
        ensure!(
            !areas.is_empty(),
            "S1-F02 附加字段命名区域“{}”没有可用的工作表单元格",
            record.name
        );
        for range in areas {
            let bounds = parse_bounds(&range.address)
                .with_context(|| format!("S1-F02 附加字段命名区域无效：{}", range.address))?;
            ensure!(
                bounds.top == bounds.bottom && bounds.left == bounds.right,
                "S1-F02 附加字段命名区域“{}”必须只定位一个单元格",
                record.name
            );
            let field = SummaryField {
                name: field_name.to_owned(),
                range: range.clone(),
                bounds,
            };
            if is_global {
                global_fields.push(field);
            } else {
                local_fields
                    .entry(range.sheet_name.to_lowercase())
                    .or_default()
                    .push(field);
            }
        }
    }
    sort_summary_fields(&mut global_fields);
    for fields in local_fields.values_mut() {
        sort_summary_fields(fields);
    }
    Ok((global_fields, local_fields))
}

fn sort_summary_fields(fields: &mut [SummaryField]) {
    fields.sort_by(|left, right| {
        (left.bounds.top, left.bounds.left, &left.name).cmp(&(
            right.bounds.top,
            right.bounds.left,
            &right.name,
        ))
    });
}

fn summary_fields_for_sheet(
    global_fields: &[SummaryField],
    local_fields: &HashMap<String, Vec<SummaryField>>,
    sheet_name: &str,
) -> Result<Vec<SummaryField>> {
    let mut fields = global_fields.to_vec();
    if let Some(local) = local_fields.get(&sheet_name.to_lowercase()) {
        fields.extend(local.iter().cloned());
    }
    let mut seen = HashSet::new();
    for field in &fields {
        let key = field.name.trim().to_lowercase();
        ensure!(
            seen.insert(key),
            "工作表“{}”存在重复的单元格区域字段名：{}",
            sheet_name,
            field.name
        );
    }
    Ok(fields)
}

fn discover_source_workbooks(input_dir: &Path, recursive: bool) -> Result<Vec<PathBuf>> {
    fn visit(directory: &Path, recursive: bool, files: &mut Vec<PathBuf>) -> Result<()> {
        let entries = fs::read_dir(directory)
            .with_context(|| format!("扫描 S1-F02 汇总目录失败：{}", directory.display()))?;
        for entry in entries {
            let entry = entry?;
            let path = entry.path();
            let kind = entry.file_type()?;
            if kind.is_dir() && recursive {
                visit(&path, recursive, files)?;
            } else if kind.is_file()
                && path
                    .extension()
                    .and_then(|value| value.to_str())
                    .is_some_and(|value| value.eq_ignore_ascii_case("xlsx"))
                && !path
                    .file_name()
                    .and_then(|value| value.to_str())
                    .unwrap_or_default()
                    .starts_with("~$")
                && !path
                    .file_stem()
                    .and_then(|value| value.to_str())
                    .unwrap_or_default()
                    .contains("_审核版")
            {
                files.push(path);
            }
        }
        Ok(())
    }

    let mut found = Vec::new();
    visit(input_dir, recursive, &mut found)?;
    found.sort();
    let sources = if recursive {
        let nested = found.iter().any(|path| path.parent() != Some(input_dir));
        if nested {
            found
                .into_iter()
                .filter(|path| path.parent() != Some(input_dir))
                .collect()
        } else {
            found
        }
    } else {
        found
    };
    Ok(sources)
}

fn decode_sheet_values(sheet: &SheetDump) -> Result<BTreeMap<(usize, usize), CellValue>> {
    sheet
        .values
        .iter()
        .map(|(address, value)| Ok((parse_cell(address)?, CellValue::from_probe_json(value)?)))
        .collect()
}

fn parse_bounds(address: &str) -> Result<Bounds> {
    let normalized = address.replace('$', "");
    ensure!(
        !normalized.ends_with(':'),
        "命名区域不是有效矩形：{address}"
    );
    let mut parts = normalized.split(':');
    let first_address = parts.next().context("命名区域地址为空")?;
    let first = parse_cell(first_address)?;
    let last_address = parts
        .next()
        .filter(|address| !address.is_empty())
        .unwrap_or(first_address);
    let last = parse_cell(last_address)?;
    ensure!(parts.next().is_none(), "命名区域不是单个矩形");
    Ok(Bounds {
        top: first.0.min(last.0),
        left: first.1.min(last.1),
        bottom: first.0.max(last.0),
        right: first.1.max(last.1),
    })
}

fn parse_cell(address: &str) -> Result<(usize, usize)> {
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
                .context("命名区域列号无效")?;
        } else if character.is_ascii_digit() {
            saw_row = true;
            row = row
                .checked_mul(10)
                .and_then(|value| value.checked_add(character.to_digit(10).unwrap() as usize))
                .context("命名区域行号无效")?;
        } else {
            bail!("单元格地址格式无效：{address}");
        }
    }
    ensure!(row > 0 && column > 0, "单元格地址缺少行号或列号：{address}");
    Ok((row, column))
}

fn cell_value(grid: &BTreeMap<(usize, usize), CellValue>, row: usize, column: usize) -> CellValue {
    grid.get(&(row, column))
        .cloned()
        .unwrap_or(CellValue::Empty)
}

fn read_header_values(
    sheet: &SheetDump,
    grid: &BTreeMap<(usize, usize), CellValue>,
    header_bounds: Bounds,
    left: usize,
    right: usize,
) -> Result<Vec<String>> {
    let mut rows = Vec::<Vec<CellValue>>::new();
    for row in header_bounds.top..=header_bounds.bottom {
        let mut values = (left..=right)
            .map(|column| cell_value(grid, row, column))
            .collect::<Vec<_>>();
        for merge in &sheet.merges {
            let merge_bounds = parse_bounds(merge)
                .with_context(|| format!("模板工作表包含无法识别的合并范围：{merge}"))?;
            if row < merge_bounds.top || row > merge_bounds.bottom {
                continue;
            }
            let anchor_value = cell_value(grid, merge_bounds.top, merge_bounds.left);
            if is_empty_value(&anchor_value) {
                continue;
            }
            for column in left..=right {
                if column >= merge_bounds.left
                    && column <= merge_bounds.right
                    && is_empty_value(&values[column - left])
                {
                    values[column - left] = anchor_value.clone();
                }
            }
        }
        rows.push(values);
    }

    let width = right.saturating_sub(left) + 1;
    (0..width)
        .map(|column| {
            rows.iter()
                .map(|row| header_text(row.get(column).unwrap_or(&CellValue::Empty)))
                .collect::<Result<Vec<_>>>()
                .map(|parts| {
                    parts
                        .into_iter()
                        .filter(|part| !part.is_empty())
                        .collect::<Vec<_>>()
                        .join("_")
                })
        })
        .collect()
}

fn header_text(value: &CellValue) -> Result<String> {
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
    row: u32,
    values: &[CellValue],
    format: &Format,
) -> Result<()> {
    for (column, value) in values.iter().enumerate() {
        let column = column as u16;
        match value {
            CellValue::Empty => {
                worksheet.write_blank(row, column, format)?;
            }
            CellValue::Text(value) => {
                worksheet.write_string_with_format(row, column, value, format)?;
            }
            CellValue::Number(value) | CellValue::DateTimeSerial(value) => {
                worksheet.write_number_with_format(row, column, *value, format)?;
            }
            CellValue::Integer(value) => {
                worksheet.write_number_with_format(row, column, *value as f64, format)?;
            }
            CellValue::Boolean(value) => {
                worksheet.write_boolean_with_format(row, column, *value, format)?;
            }
            CellValue::Error(_) | CellValue::DateTimeIso(_) | CellValue::DurationIso(_) => {
                bail!("S1-F02 汇总遇到无法按 Excel Value2 契约写入的单元格类型")
            }
        }
    }
    Ok(())
}

fn validate_summary_workbook(path: &Path, groups: &[RegionGroup]) -> Result<()> {
    let actual = read_workbook(path).context("重新打开 S1-F02 命名区域汇总工作簿失败")?;
    let actual_names = actual
        .sheets
        .iter()
        .map(|sheet| sheet.name.as_str())
        .collect::<Vec<_>>();
    let expected_names = if groups.is_empty() {
        vec!["汇总"]
    } else {
        groups
            .iter()
            .rev()
            .map(|group| group.output_sheet_name.as_str())
            .collect()
    };
    ensure!(
        actual_names == expected_names,
        "S1-F02 输出工作表名称或顺序不匹配"
    );

    for (sheet, group) in actual.sheets.iter().zip(groups.iter().rev()) {
        let mut expected = BTreeMap::<(usize, usize), CellValue>::new();
        let mut headers = vec![
            CellValue::Text("来源文件".to_owned()),
            CellValue::Text("来源工作表".to_owned()),
        ];
        headers.extend(
            group
                .extra_fields
                .iter()
                .map(|field| CellValue::Text(field.name.clone())),
        );
        headers.extend(group.headers.iter().cloned().map(CellValue::Text));
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
            .map(|(address, value)| Ok((parse_cell(address)?, CellValue::from_probe_json(value)?)))
            .collect::<Result<BTreeMap<_, _>>>()?;
        ensure!(
            expected.len() == actual_values.len(),
            "S1-F02 输出工作表“{}”的非空单元格数量不匹配",
            group.output_sheet_name
        );
        for (coordinate, expected_value) in expected {
            let Some(actual_value) = actual_values.get(&coordinate) else {
                bail!("S1-F02 输出工作表“{}”缺少单元格", group.output_sheet_name);
            };
            ensure!(
                values_equal(&expected_value, actual_value),
                "S1-F02 输出工作表“{}”的单元格值或类型不匹配",
                group.output_sheet_name
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
            (Some(left), Some(right)) => {
                left == right
                    || (left - right).abs() <= left.abs().max(right.abs()).max(1.0) * 1e-15
            }
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
    matches!(value, CellValue::Empty) || matches!(value, CellValue::Text(value) if value.is_empty())
}

fn output_sheet_name(base: &str, used: &mut HashSet<String>) -> String {
    let mut candidate = base
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
    if candidate.is_empty() {
        candidate = "汇总".to_owned();
    }
    let root = candidate.clone();
    let mut suffix = 2_usize;
    while !used.insert(candidate.to_lowercase()) {
        let marker = format!("_{suffix}");
        let keep = 31_usize.saturating_sub(marker.chars().count());
        candidate = format!("{}{}", root.chars().take(keep).collect::<String>(), marker);
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
            .and_then(|value| value.to_str())
            .unwrap_or("s1-region-summary"),
        std::process::id()
    )))
}

#[cfg(test)]
mod tests {
    use super::*;
    use excel_reader::{DefinedNameArea, DefinedNameScope};

    fn area_record(
        name: &str,
        scope: DefinedNameScope,
        sheet: &str,
        address: &str,
    ) -> DefinedNameRecord {
        DefinedNameRecord {
            name: name.to_owned(),
            scope,
            raw_reference: format!("'{sheet}'!{address}"),
            parse_status: DefinedNameParseStatus::Areas,
            areas: vec![DefinedNameArea {
                sheet_name: sheet.to_owned(),
                sheet_kind: "worksheet".to_owned(),
                reference: address.to_owned(),
            }],
        }
    }

    #[test]
    fn metadata_fields_require_single_cell_areas() {
        let record = area_record(
            "全局单元格区域.批次",
            DefinedNameScope::Workbook,
            "Meta",
            "$A$1:$B$1",
        );
        let error = parse_summary_fields(&[record]).unwrap_err().to_string();
        assert!(error.contains("必须只定位一个单元格"), "{error}");
    }

    #[test]
    fn metadata_fields_sort_global_fields_then_local_fields_and_reject_duplicates() {
        let records = vec![
            area_record(
                "全局单元格区域.后批次",
                DefinedNameScope::Workbook,
                "Meta",
                "$B$2",
            ),
            area_record(
                "全局单元格区域.前批次",
                DefinedNameScope::Workbook,
                "Meta",
                "$A$1",
            ),
            area_record(
                "单元格区域.分组",
                DefinedNameScope::Worksheet {
                    local_sheet_id: 0,
                    sheet_name: Some("Data".to_owned()),
                    sheet_kind: Some("worksheet".to_owned()),
                },
                "Data",
                "$F$1",
            ),
        ];
        let (global, local) = parse_summary_fields(&records).expect("parse metadata fields");
        let fields = summary_fields_for_sheet(&global, &local, "data").expect("select metadata");
        assert_eq!(
            fields
                .iter()
                .map(|field| field.name.as_str())
                .collect::<Vec<_>>(),
            ["前批次", "后批次", "分组"]
        );

        let mut duplicated = records;
        duplicated.push(area_record(
            "单元格区域.前批次",
            DefinedNameScope::Worksheet {
                local_sheet_id: 0,
                sheet_name: Some("Data".to_owned()),
                sheet_kind: Some("worksheet".to_owned()),
            },
            "Data",
            "$G$1",
        ));
        let (global, local) = parse_summary_fields(&duplicated).expect("parse duplicate fields");
        let error = summary_fields_for_sheet(&global, &local, "Data")
            .unwrap_err()
            .to_string();
        assert!(error.contains("存在重复的单元格区域字段名"), "{error}");
    }
}
