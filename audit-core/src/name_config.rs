//! Pure value helpers from `base_audit.name_config`.
//!
//! Configuration values are read through the shared Rust Excel reader; this
//! module applies Audit feature-selection and named-range mapping semantics and
//! never modifies the configuration workbook.

use crate::models::CopyRange;
use anyhow::{bail, Context, Result};
use excel_reader::{CellValue, SheetDump};
use serde::{Deserialize, Serialize};
use std::{
    collections::{BTreeMap, BTreeSet},
    path::Path,
};

const MAPPING_SHEET: &str = "模块化功能";
const MAPPING_HEADERS: [&str; 6] = ["功能名", "执行模块", "命名区域名", "输入", "输出", "备注"];
const CONDITIONAL_FORMAT_EXTRACT_FUNCTION: &str = "汇总_条件格式结果提取";
const STRUCTURE_COMPARE_FUNCTION: &str = "核对_表结构比对";

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct FeatureMapping {
    pub name: String,
    pub feature_type: String,
    pub range_names: Vec<String>,
    pub workbook_limited: bool,
    pub workbook_keyword: String,
    pub remark: String,
    #[serde(default)]
    pub output: String,
    #[serde(default)]
    pub input_note: String,
}

/// One already-resolved destination from a workbook or worksheet-scoped name.
///
/// OOXML parsing and reference validation belong to the shared Excel reader;
/// this Audit type carries only the values needed for business mapping.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct NamedRangeArea {
    pub name: String,
    pub sheet_name: String,
    pub address: String,
}

/// Split a delimited named-range cell using the source project's separators.
pub fn split_names(value: &str) -> Vec<String> {
    value
        .split(|character| matches!(character, '、' | '，' | ',' | '；' | ';' | '\n'))
        .map(str::trim)
        .filter(|part| !part.is_empty())
        .map(str::to_owned)
        .collect()
}

/// Interpret the source project's affirmative workbook-limit values.
pub fn is_limited(value: &str) -> bool {
    matches!(
        value.to_lowercase().as_str(),
        "是" | "y" | "yes" | "true" | "1"
    )
}

/// Return mappings whose feature type matches exactly.
pub fn features_of_type<'a>(
    features: &'a [FeatureMapping],
    feature_type: &str,
) -> Vec<&'a FeatureMapping> {
    features
        .iter()
        .filter(|feature| feature.feature_type == feature_type)
        .collect()
}

/// Match an exact named range or an underscore-suffixed variant.
pub fn matches_named_range(mapping: &FeatureMapping, candidate: &str) -> bool {
    let simple = candidate
        .rsplit('!')
        .next()
        .unwrap_or(candidate)
        .trim()
        .trim_matches('\'');
    let folded = simple.to_lowercase();

    mapping.range_names.iter().any(|name| {
        let name = name.to_lowercase();
        folded == name || folded.starts_with(&format!("{name}_"))
    })
}

/// Resolve already-parsed defined-name destinations for the requested Audit
/// mappings, matching the source project's `named_ranges` behavior.
///
/// Workbook-scoped and worksheet-scoped names are both accepted. Unmatched
/// names and destinations pointing at sheets absent from the input workbook
/// are ignored. The reader must omit unresolved or invalid OOXML references;
/// this function does not parse Excel references itself.
pub fn resolve_named_ranges_for_mappings(
    mappings: &[FeatureMapping],
    areas: &[NamedRangeArea],
    workbook_sheets: &[String],
) -> Vec<CopyRange> {
    let mut seen = BTreeSet::new();
    let mut resolved = Vec::new();

    for area in areas {
        if !mappings
            .iter()
            .any(|mapping| matches_named_range(mapping, &area.name))
            || !workbook_sheets
                .iter()
                .any(|sheet_name| sheet_name == &area.sheet_name)
        {
            continue;
        }

        let address = area.address.replace('$', "").trim().to_owned();
        if address.is_empty() {
            continue;
        }
        let identity = (area.sheet_name.clone(), address.to_ascii_uppercase());
        if seen.insert(identity) {
            resolved.push(CopyRange {
                sheet_name: area.sheet_name.clone(),
                address,
            });
        }
    }

    resolved
}

/// Load the S1 conditional-format and structure mappings from the configured
/// module sheet, using the source project's defaults when that sheet is
/// absent or contains no applicable rows.
pub fn load_s1_conditional_mappings(
    config_path: Option<&Path>,
    template_path: &Path,
) -> Result<Vec<FeatureMapping>> {
    let Some(config_path) = config_path.filter(|path| path.is_file()) else {
        return Ok(default_s1_conditional_mappings());
    };

    let workbook = excel_reader::read_workbook(config_path)
        .with_context(|| format!("读取 Audit 功能映射配置 {}", config_path.display()))?;
    let Some(sheet) = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == MAPPING_SHEET)
    else {
        return Ok(default_s1_conditional_mappings());
    };

    let rows = mapping_sheet_rows(sheet)?;
    let template_name = template_path
        .file_name()
        .map(|name| name.to_string_lossy().into_owned())
        .unwrap_or_default();
    parse_s1_conditional_mappings(&rows, &template_name)
}

fn parse_s1_conditional_mappings(
    rows: &[Vec<String>],
    template_name: &str,
) -> Result<Vec<FeatureMapping>> {
    let Some(headers) = rows.first() else {
        return Ok(default_s1_conditional_mappings());
    };
    if headers.iter().all(String::is_empty) {
        return Ok(default_s1_conditional_mappings());
    }

    let mut positions = BTreeMap::new();
    for header in MAPPING_HEADERS {
        let Some(position) = headers.iter().position(|value| value == header) else {
            bail!("config.xlsx 的“{MAPPING_SHEET}”缺少列：{header}");
        };
        positions.insert(header, position);
    }
    for optional_header in ["是否限定工作簿", "工作簿关键字"] {
        if let Some(position) = headers.iter().position(|value| value == optional_header) {
            positions.insert(optional_header, position);
        }
    }

    let template_name = template_name.to_lowercase();
    let mut selected: Vec<(u8, FeatureMapping)> = Vec::new();
    for row in rows.iter().skip(1) {
        let value = |header: &str| -> String {
            positions
                .get(header)
                .and_then(|index| row.get(*index))
                .cloned()
                .unwrap_or_default()
        };

        let name = value("功能名");
        let feature_type = value("执行模块");
        if name.is_empty()
            || !matches!(
                feature_type.as_str(),
                CONDITIONAL_FORMAT_EXTRACT_FUNCTION | STRUCTURE_COMPARE_FUNCTION
            )
        {
            continue;
        }

        let workbook_limited = is_limited(&value("是否限定工作簿"));
        let workbook_keyword = value("工作簿关键字");
        if workbook_limited
            && (workbook_keyword.is_empty()
                || !template_name.contains(&workbook_keyword.to_lowercase()))
        {
            continue;
        }

        let mapping = FeatureMapping {
            name,
            feature_type,
            range_names: split_names(&value("命名区域名")),
            workbook_limited,
            workbook_keyword,
            remark: value("备注"),
            // The source loader constructs this six-field variant and leaves
            // the dataclass's input/output metadata at their empty defaults.
            output: String::new(),
            input_note: String::new(),
        };
        let priority = u8::from(mapping.workbook_limited) + 1;
        if let Some((existing_priority, existing)) = selected.iter_mut().find(|(_, existing)| {
            existing.name == mapping.name && existing.feature_type == mapping.feature_type
        }) {
            if priority >= *existing_priority {
                *existing_priority = priority;
                *existing = mapping;
            }
        } else {
            selected.push((priority, mapping));
        }
    }

    if selected.is_empty() {
        Ok(default_s1_conditional_mappings())
    } else {
        Ok(selected.into_iter().map(|(_, mapping)| mapping).collect())
    }
}

fn mapping_sheet_rows(sheet: &SheetDump) -> Result<Vec<Vec<String>>> {
    let mut cells: BTreeMap<usize, BTreeMap<usize, String>> = BTreeMap::new();
    let mut max_column = 0usize;
    for (address, encoded) in &sheet.values {
        let (row, column) = parse_configuration_cell_address(address)
            .with_context(|| format!("配置工作表包含无效单元格坐标：{address}"))?;
        let value = CellValue::from_probe_json(encoded)
            .with_context(|| format!("解码配置单元格 {address}"))?;
        let text = configuration_cell_text(value);
        if !text.is_empty() {
            max_column = max_column.max(column);
            cells.entry(row).or_default().insert(column, text);
        }
    }
    if max_column == 0 || !cells.contains_key(&1) {
        return Ok(Vec::new());
    }

    Ok(cells
        .into_iter()
        .map(|(_, row)| {
            let mut values = vec![String::new(); max_column];
            for (column, value) in row {
                values[column - 1] = value;
            }
            values
        })
        .collect())
}

fn configuration_cell_text(value: CellValue) -> String {
    match value {
        CellValue::Empty => String::new(),
        CellValue::Text(value) => value.trim().to_owned(),
        CellValue::Number(value) | CellValue::DateTimeSerial(value) => value.to_string(),
        CellValue::Integer(value) => value.to_string(),
        CellValue::Boolean(value) => {
            if value {
                "True".to_owned()
            } else {
                "False".to_owned()
            }
        }
        CellValue::DateTimeIso(value) | CellValue::DurationIso(value) | CellValue::Error(value) => {
            value.trim().to_owned()
        }
    }
}

fn parse_configuration_cell_address(address: &str) -> Option<(usize, usize)> {
    let address = address.replace('$', "");
    let split = address.find(|character: char| character.is_ascii_digit())?;
    let (column, row) = address.split_at(split);
    if column.is_empty()
        || !column.bytes().all(|byte| byte.is_ascii_alphabetic())
        || row.is_empty()
        || !row.bytes().all(|byte| byte.is_ascii_digit())
    {
        return None;
    }
    let column = column.bytes().try_fold(0usize, |value, byte| {
        value
            .checked_mul(26)?
            .checked_add(usize::from(byte.to_ascii_uppercase() - b'A' + 1))
    })?;
    let row = row.parse::<usize>().ok()?;
    (column > 0 && row > 0).then_some((row, column))
}

fn default_s1_conditional_mappings() -> Vec<FeatureMapping> {
    vec![
        FeatureMapping {
            name: "条件格式结果提取".to_owned(),
            feature_type: CONDITIONAL_FORMAT_EXTRACT_FUNCTION.to_owned(),
            range_names: vec!["条件格式区域".to_owned()],
            workbook_limited: false,
            workbook_keyword: String::new(),
            remark: "提取命名区域内被条件格式实际改变填充色的单元格；单元格批注作为校验类型，条件格式公式作为描述，表结构区域用于定位校验指标".to_owned(),
            output: "审核结果集".to_owned(),
            input_note: "源数据目录的条件格式区域".to_owned(),
        },
        FeatureMapping {
            name: "表结构比对".to_owned(),
            feature_type: STRUCTURE_COMPARE_FUNCTION.to_owned(),
            range_names: vec!["表结构区域".to_owned()],
            workbook_limited: false,
            workbook_keyword: String::new(),
            remark: "逐格精确核对固定表结构".to_owned(),
            output: "运行日志".to_owned(),
            input_note: "源数据目录、模板表结构区域".to_owned(),
        },
    ]
}
