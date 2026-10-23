//! First Rust vertical slice of the S2 source workbook importer.
//!
//! The public model is business-specific and stays in `audit-core`; workbook
//! access is delegated to the shared `excel-reader` Calamine adapter.

use crate::{
    report_source::report_source_identity,
    unit_conversion::{convert_amount_from_cell, AmountSourceFormat, DecimalTuple},
};
use anyhow::{bail, Context, Result};
use excel_reader::{read_value_columns, CellValue, SheetValues};
use serde::{Deserialize, Serialize};
use std::{
    collections::{BTreeMap, HashMap},
    fs,
    path::{Path, PathBuf},
};

const AMOUNT_DATA_TYPES: &[&str] = &["余额", "累发"];
const MAX_EXACT_XLSX_INTEGER: f64 = 9_007_199_254_740_992.0;

#[derive(Clone, Debug, Deserialize, Default)]
pub struct S2ImportConfig {
    #[serde(default)]
    pub indicators: HashMap<String, S2IndicatorConfig>,
    #[serde(default)]
    pub institutions_by_id: HashMap<String, S2InstitutionProfile>,
    #[serde(default)]
    pub institutions_by_name: HashMap<String, S2InstitutionProfile>,
    #[serde(default)]
    pub unit_settings: S2UnitSettings,
}

#[derive(Clone, Debug, Deserialize, Default)]
pub struct S2IndicatorConfig {
    #[serde(default)]
    pub indicator_name: String,
    #[serde(default)]
    pub form_code: String,
    #[serde(default)]
    pub data_type: String,
}

#[derive(Clone, Debug, Deserialize, Default)]
pub struct S2InstitutionProfile {
    #[serde(default)]
    pub institution_id: String,
    #[serde(default)]
    pub institution_name: String,
    #[serde(default)]
    pub social_credit_code: String,
    #[serde(default)]
    pub institution_type: String,
    #[serde(default)]
    pub handling_branch: String,
    #[serde(default)]
    pub region: String,
}

#[derive(Clone, Debug, Deserialize)]
pub struct S2UnitSettings {
    #[serde(default = "default_yuan")]
    pub source_unit: String,
    #[serde(default = "default_yuan")]
    pub output_file_unit: String,
}

impl Default for S2UnitSettings {
    fn default() -> Self {
        Self {
            source_unit: default_yuan(),
            output_file_unit: default_yuan(),
        }
    }
}

fn default_yuan() -> String {
    "元".to_owned()
}

/// JSON-stable representation of values surfaced by the S2 import contract.
/// Decimal amounts retain their exact base-ten tuple.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(tag = "type", content = "value")]
pub enum ImportedValue {
    #[serde(rename = "empty")]
    Empty,
    #[serde(rename = "string")]
    Text(String),
    #[serde(rename = "number")]
    Number(f64),
    #[serde(rename = "integer")]
    Integer(i64),
    #[serde(rename = "boolean")]
    Boolean(bool),
    #[serde(rename = "datetime_serial")]
    DateTimeSerial(f64),
    #[serde(rename = "datetime_iso")]
    DateTimeIso(String),
    #[serde(rename = "duration_iso")]
    DurationIso(String),
    #[serde(rename = "error")]
    Error(String),
    #[serde(rename = "decimal")]
    Decimal(DecimalTuple),
}

impl From<CellValue> for ImportedValue {
    fn from(value: CellValue) -> Self {
        match value {
            CellValue::Empty => Self::Empty,
            CellValue::Text(value) => Self::Text(value),
            CellValue::Number(value) => Self::Number(value),
            CellValue::Integer(value) => Self::Integer(value),
            CellValue::Boolean(value) => Self::Boolean(value),
            CellValue::DateTimeSerial(value) => Self::DateTimeSerial(value),
            CellValue::DateTimeIso(value) => Self::DateTimeIso(value),
            CellValue::DurationIso(value) => Self::DurationIso(value),
            CellValue::Error(value) => Self::Error(value),
        }
    }
}

#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
pub struct S2SourceRef {
    pub source_file: String,
    pub source_sheet: String,
    pub source_row: usize,
}

#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
pub struct IndicatorRecord {
    pub institution_id: String,
    pub institution_name: String,
    pub social_credit_code: String,
    pub institution_type: String,
    pub handling_branch: String,
    pub region: String,
    pub form_code: String,
    pub indicator_code: String,
    pub indicator_name: String,
    pub period: String,
    pub source_date: String,
    pub value: ImportedValue,
    pub value_type: String,
    pub source_unit: String,
    pub normalized_value: ImportedValue,
    pub normalized_unit: String,
    pub source: S2SourceRef,
}

impl IndicatorRecord {
    pub fn logical_key(&self) -> (String, String, String) {
        (
            self.institution_id.clone(),
            self.indicator_code.clone(),
            self.period.clone(),
        )
    }
}

#[derive(Clone, Debug, Default, PartialEq, Serialize)]
pub struct S2ImportResult {
    pub records: Vec<IndicatorRecord>,
    pub duplicates: Vec<(IndicatorRecord, IndicatorRecord)>,
    pub warnings: Vec<String>,
}

impl S2ImportResult {
    /// Keep the first record for each logical key; duplicates remain explicit.
    pub fn by_logical_key(&self) -> BTreeMap<(String, String, String), &IndicatorRecord> {
        let mut result = BTreeMap::new();
        for record in &self.records {
            result.entry(record.logical_key()).or_insert(record);
        }
        result
    }
}

#[derive(Clone, Debug)]
struct PendingRecord {
    path: PathBuf,
    sheet: String,
    row: usize,
    raw_id: String,
    raw_name: String,
    value: Option<CellValue>,
    code: String,
    name: String,
    form_code: String,
    data_date: String,
    source_format: AmountSourceFormat,
}

/// Import one S2 period directory using Calamine's stored-value view.
///
/// This is a candidate Rust path, not a production default. It intentionally
/// preserves the source importer's `.xls`/`.xlsx` discovery contract and does
/// not try another reader when Calamine rejects a file.
pub fn import_period_directory(
    directory: impl AsRef<Path>,
    config: &S2ImportConfig,
    label: &str,
) -> Result<S2ImportResult> {
    import_period_directory_with_progress(directory, config, label, None)
}

/// Import one S2 period directory and report the same file-count progress
/// boundaries as the source Python importer.
pub fn import_period_directory_with_progress(
    directory: impl AsRef<Path>,
    config: &S2ImportConfig,
    label: &str,
    mut on_progress: Option<&mut dyn FnMut(usize, usize, &str)>,
) -> Result<S2ImportResult> {
    let directory = directory.as_ref();
    if !directory.is_dir() {
        bail!("{label}目录不存在：{}", directory.display());
    }

    let mut files = Vec::new();
    for entry in fs::read_dir(directory).with_context(|| format!("read {}", directory.display()))? {
        let entry = entry?;
        let path = entry.path();
        let Some(file_name) = path.file_name().and_then(|value| value.to_str()) else {
            continue;
        };
        let extension = path
            .extension()
            .and_then(|value| value.to_str())
            .unwrap_or_default()
            .to_ascii_lowercase();
        if path.is_file()
            && matches!(extension.as_str(), "xls" | "xlsx")
            && !file_name.starts_with("~$")
        {
            files.push(path);
        }
    }
    files.sort();
    if files.is_empty() {
        bail!(
            "{label}目录没有可读取的 .xls/.xlsx 文件：{}",
            directory.display()
        );
    }

    let mut result = S2ImportResult::default();
    let mut pending = Vec::new();
    let mut resolved_paths = BTreeMap::new();

    let total_files = files.len();
    for (index, path) in files.iter().enumerate() {
        if let Some(callback) = on_progress.as_mut() {
            let name = file_name(path);
            (**callback)(index, total_files, &name);
        }
        let file_identity = report_source_identity(&path, "");
        if !file_identity.recognized {
            result
                .warnings
                .push(format!("无法识别文件名，已跳过：{}", file_name(&path)));
            continue;
        }

        let canonical = fs::canonicalize(&path)
            .with_context(|| format!("resolve source path {}", path.display()))?;
        resolved_paths.insert(path.clone(), canonical.display().to_string());
        let source_format = if path
            .extension()
            .and_then(|value| value.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xls"))
        {
            AmountSourceFormat::Xls
        } else {
            AmountSourceFormat::XlsxLike
        };
        let workbook = read_value_columns(&path, 3)
            .with_context(|| format!("{}报送工作簿读取失败：{}", label, path.display()))?;

        for sheet in workbook {
            let identity = report_source_identity(&path, &sheet.name);
            let max_row = max_used_row(&sheet);
            for row in 4..=max_row {
                let code_cell = read_cell(&sheet, &format!("A{row}"))?;
                let code = python_code(code_cell.as_ref());
                if code.is_empty() {
                    continue;
                }
                let name_cell = read_cell(&sheet, &format!("B{row}"))?;
                let name = name_cell
                    .as_ref()
                    .map(python_string)
                    .unwrap_or_default()
                    .trim()
                    .to_owned();
                let value = read_cell(&sheet, &format!("C{row}"))?;
                pending.push(PendingRecord {
                    path: path.clone(),
                    sheet: sheet.name.clone(),
                    row,
                    raw_id: identity.institution_id.clone(),
                    raw_name: identity.institution_name.clone(),
                    value,
                    code,
                    name,
                    form_code: identity.form_code.clone(),
                    data_date: identity.data_date.clone(),
                    source_format,
                });
            }
        }
    }

    if let Some(callback) = on_progress.as_mut() {
        (**callback)(total_files, total_files, &format!("正在整理{label}指标"));
    }

    let mut seen: BTreeMap<(String, String, String), IndicatorRecord> = BTreeMap::new();
    for item in pending {
        let profile = config.institutions_by_id.get(&item.raw_id).or_else(|| {
            let candidate = config.institutions_by_name.get(&item.raw_name)?;
            if item.raw_id.is_empty()
                || candidate.social_credit_code.is_empty()
                || candidate.social_credit_code == item.raw_id
            {
                Some(candidate)
            } else {
                None
            }
        });

        // File/sheet identity wins; configuration fills only missing fields.
        let institution_id = if !item.raw_id.is_empty() {
            item.raw_id.clone()
        } else if let Some(profile) = profile {
            profile.institution_id.clone()
        } else {
            item.raw_name.clone()
        };
        let institution_name = if !item.raw_name.is_empty() {
            item.raw_name.clone()
        } else if let Some(profile) = profile {
            profile.institution_name.clone()
        } else {
            item.raw_id.clone()
        };
        let indicator = config.indicators.get(&item.code);
        let expected_form = indicator.map_or("", |config| config.form_code.as_str());
        if !expected_form.is_empty()
            && !item.form_code.is_empty()
            && expected_form != item.form_code
        {
            result.warnings.push(format!(
                "表单归属不一致：{}/{} 第{}行，声明表单 {}，指标 {} 配置归属 {}",
                file_name(&item.path),
                item.sheet,
                item.row,
                item.form_code,
                item.code,
                expected_form
            ));
        }

        let data_type = indicator.map_or("", |config| config.data_type.as_str());
        let is_text = data_type == "文字";
        let is_amount = AMOUNT_DATA_TYPES.contains(&data_type);
        let normalized_raw = item
            .value
            .as_ref()
            .map(|value| normalize_fast_value(value, item.source_format))
            .unwrap_or(CellValue::Empty);

        let normalized_value = if is_text {
            ImportedValue::Text(if python_truthy(&normalized_raw) {
                python_string(&normalized_raw).trim().to_owned()
            } else {
                String::new()
            })
        } else if is_amount {
            match convert_amount_from_cell(
                item.value.as_ref(),
                item.source_format,
                &config.unit_settings.source_unit,
                &config.unit_settings.output_file_unit,
            )
            .with_context(|| {
                format!(
                    "{}/{} 第{}行指标 {} 无法转换金额到输出单位",
                    file_name(&item.path),
                    item.sheet,
                    item.row,
                    item.code
                )
            })? {
                Some(value) => ImportedValue::Decimal(value),
                None => normalized_raw.clone().into(),
            }
        } else if let Some(value) = python_number(&normalized_raw) {
            ImportedValue::Number(value)
        } else {
            normalized_raw.clone().into()
        };
        let normalized_unit = if is_amount && matches!(&normalized_value, ImportedValue::Decimal(_))
        {
            config.unit_settings.output_file_unit.clone()
        } else {
            String::new()
        };

        let record = IndicatorRecord {
            institution_id,
            institution_name,
            social_credit_code: if !item.raw_id.is_empty() {
                item.raw_id.clone()
            } else {
                profile.map_or_else(String::new, |profile| profile.social_credit_code.clone())
            },
            institution_type: profile
                .map_or_else(String::new, |profile| profile.institution_type.clone()),
            handling_branch: profile
                .map_or_else(String::new, |profile| profile.handling_branch.clone()),
            region: profile.map_or_else(String::new, |profile| profile.region.clone()),
            form_code: if !item.form_code.is_empty() {
                item.form_code.clone()
            } else {
                expected_form.to_owned()
            },
            indicator_code: item.code.clone(),
            indicator_name: if !item.name.is_empty() {
                item.name.clone()
            } else {
                indicator.map_or_else(String::new, |config| config.indicator_name.clone())
            },
            period: period(&item.data_date),
            source_date: item.data_date,
            value: normalized_raw.into(),
            value_type: if !data_type.is_empty() {
                data_type.to_owned()
            } else if is_text {
                "文字".to_owned()
            } else {
                "数值".to_owned()
            },
            source_unit: if is_amount {
                config.unit_settings.source_unit.clone()
            } else {
                "原始单位".to_owned()
            },
            normalized_value,
            normalized_unit,
            source: S2SourceRef {
                source_file: resolved_paths
                    .get(&item.path)
                    .cloned()
                    .unwrap_or_else(|| item.path.display().to_string()),
                source_sheet: item.sheet,
                source_row: item.row,
            },
        };
        let key = record.logical_key();
        if let Some(previous) = seen.get(&key) {
            result.duplicates.push((previous.clone(), record));
            continue;
        }
        seen.insert(key, record.clone());
        result.records.push(record);
    }
    Ok(result)
}

fn read_cell(sheet: &SheetValues, address: &str) -> Result<Option<CellValue>> {
    Ok(sheet.values.get(address).cloned())
}

fn max_used_row(sheet: &SheetValues) -> usize {
    sheet
        .values
        .keys()
        .filter_map(|address| parse_cell_reference(address))
        .filter(|(_, column)| *column <= 3)
        .map(|(row, _)| row)
        .max()
        .unwrap_or(0)
}

pub(crate) fn parse_cell_reference(address: &str) -> Option<(usize, usize)> {
    let mut letters = String::new();
    let mut digits = String::new();
    for character in address.chars() {
        if character.is_ascii_alphabetic() && digits.is_empty() {
            letters.push(character.to_ascii_uppercase());
        } else if character.is_ascii_digit() {
            digits.push(character);
        } else {
            return None;
        }
    }
    if letters.is_empty() || digits.is_empty() {
        return None;
    }
    let column = letters.chars().try_fold(0_usize, |value, character| {
        value
            .checked_mul(26)?
            .checked_add((character as usize) - ('A' as usize) + 1)
    })?;
    let row = digits.parse::<usize>().ok()?;
    (row > 0).then_some((row, column))
}

pub(crate) fn normalize_fast_value(
    value: &CellValue,
    source_format: AmountSourceFormat,
) -> CellValue {
    match (source_format, value) {
        (AmountSourceFormat::Xls, CellValue::Boolean(value)) => {
            CellValue::Integer(i64::from(*value))
        }
        (AmountSourceFormat::Xls, CellValue::Integer(value)) => CellValue::Number(*value as f64),
        (AmountSourceFormat::XlsxLike, CellValue::Number(value))
            if value.is_finite()
                && value.fract() == 0.0
                && value.abs() <= MAX_EXACT_XLSX_INTEGER =>
        {
            CellValue::Integer(*value as i64)
        }
        (AmountSourceFormat::XlsxLike, CellValue::Text(value)) if value.is_empty() => {
            CellValue::Empty
        }
        (_, value) => value.clone(),
    }
}

fn python_number(value: &CellValue) -> Option<f64> {
    match value {
        CellValue::Empty | CellValue::Boolean(_) => None,
        CellValue::Integer(value) => Some(*value as f64),
        CellValue::Number(value) => Some(*value),
        CellValue::Text(value) => value.trim().replace(',', "").parse::<f64>().ok(),
        CellValue::DateTimeSerial(_)
        | CellValue::DateTimeIso(_)
        | CellValue::DurationIso(_)
        | CellValue::Error(_) => None,
    }
}

pub(crate) fn python_truthy(value: &CellValue) -> bool {
    match value {
        CellValue::Empty => false,
        CellValue::Text(value) => !value.is_empty(),
        CellValue::Number(value) => *value != 0.0,
        CellValue::Integer(value) => *value != 0,
        CellValue::Boolean(value) => *value,
        CellValue::DateTimeSerial(_)
        | CellValue::DateTimeIso(_)
        | CellValue::DurationIso(_)
        | CellValue::Error(_) => true,
    }
}

pub(crate) fn python_string(value: &CellValue) -> String {
    match value {
        CellValue::Empty => "".to_owned(),
        CellValue::Text(value) => value.clone(),
        CellValue::Number(value) => format!("{value:?}"),
        CellValue::Integer(value) => value.to_string(),
        CellValue::Boolean(value) => if *value { "True" } else { "False" }.to_owned(),
        CellValue::DateTimeSerial(value) => format!("{value:?}"),
        CellValue::DateTimeIso(value) | CellValue::DurationIso(value) => value.clone(),
        CellValue::Error(value) => value.clone(),
    }
}

pub(crate) fn python_code(value: Option<&CellValue>) -> String {
    let Some(value) = value else {
        return String::new();
    };
    match value {
        CellValue::Number(value) if value.is_finite() && value.fract() == 0.0 => {
            return python_integer_from_float(*value);
        }
        CellValue::Integer(value) if *value != 0 => return value.to_string(),
        CellValue::Integer(0) | CellValue::Boolean(false) | CellValue::Empty => {
            return String::new()
        }
        _ => {}
    }
    let text = if python_truthy(value) {
        python_string(value).trim().to_owned()
    } else {
        String::new()
    };
    if let Some((integer, fraction)) = text.split_once('.') {
        if !integer.is_empty()
            && integer.chars().all(|character| character.is_ascii_digit())
            && !fraction.is_empty()
            && fraction.chars().all(|character| character == '0')
        {
            if let Ok(number) = text.parse::<f64>() {
                if number.is_finite() && number.fract() == 0.0 {
                    return python_integer_from_float(number);
                }
            }
        }
    }
    text
}

fn python_integer_from_float(value: f64) -> String {
    // Python int(float) has arbitrary precision. Formatting the integral f64
    // directly avoids Rust's saturating float-to-i64 cast at the 2^63 edge.
    if value == 0.0 {
        "0".to_owned()
    } else {
        format!("{value:.0}")
    }
}

fn period(value: &str) -> String {
    let digits: String = value.chars().filter(char::is_ascii_digit).collect();
    if digits.len() >= 6 {
        format!("{}-{}", &digits[..4], &digits[4..6])
    } else {
        value.to_owned()
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

    #[test]
    fn source_identity_matches_supported_filename_shapes() {
        let standard =
            report_source_identity(Path::new("9144#2026-06-30#x#20202#甲银行.xlsx"), "20202");
        assert!(standard.recognized);
        assert_eq!(standard.institution_id, "9144");
        assert_eq!(standard.institution_name, "甲银行");
        assert_eq!(standard.form_code, "20202");
        assert_eq!(standard.data_date, "2026-06-30");

        let reports = report_source_identity(
            Path::new("reports#9144#2026-06-30#01#甲银行.xlsx"),
            "Sheet1",
        );
        assert_eq!(reports.institution_id, "9144");
        assert_eq!(reports.institution_name, "甲银行");

        let banks = report_source_identity(Path::new("banks#2026-06-30#01#20202.xlsx"), "9144");
        assert_eq!(banks.institution_id, "9144");
        assert_eq!(banks.institution_name, "");
        assert_eq!(banks.form_code, "20202");
    }

    #[test]
    fn normalization_preserves_python_xls_and_xlsx_value_shapes() {
        assert_eq!(
            normalize_fast_value(&CellValue::Boolean(true), AmountSourceFormat::Xls),
            CellValue::Integer(1)
        );
        assert_eq!(
            normalize_fast_value(&CellValue::Integer(42), AmountSourceFormat::Xls),
            CellValue::Number(42.0)
        );
        assert_eq!(
            normalize_fast_value(&CellValue::Number(42.0), AmountSourceFormat::XlsxLike),
            CellValue::Integer(42)
        );
        assert_eq!(period("2026-06-30"), "2026-06");
        assert_eq!(
            python_code(Some(&CellValue::Text("20201028.0".to_owned()))),
            "20201028"
        );
        assert_eq!(python_code(Some(&CellValue::Integer(0))), "");
    }

    #[test]
    fn numeric_codes_above_i64_max_match_python_integer_conversion() {
        let beyond_i64 = 2_f64.powi(63);
        assert_eq!(
            python_code(Some(&CellValue::Number(beyond_i64))),
            "9223372036854775808"
        );
        assert_eq!(
            python_code(Some(&CellValue::Text("9223372036854775808.0".to_owned()))),
            "9223372036854775808"
        );
    }
}
