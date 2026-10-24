//! S1 validation-formula error scan over template-defined validation areas.
//!
//! The template is the authority for scan coordinates. Formula text and
//! defined names are read without recalculation; source values come from the
//! stored workbook values exposed by the Rust reader.

use crate::{
    models::Issue, s1_formula_audit::S1FormulaMetadata, template::normalize_template_name,
};
use anyhow::{ensure, Context, Result};
use excel_reader::{
    read_defined_names, read_workbook, CellValue, DefinedNameArea, DefinedNameParseStatus,
    SheetDump, WorkbookDump,
};
use serde::Serialize;
use serde_json::Value;
use std::{
    collections::{BTreeMap, BTreeSet, HashMap},
    fs,
    path::{Path, PathBuf},
    time::Instant,
};

const READER: &str = "calamine+namespace_aware_ooxml_defined_names";
const VALIDATION_NAME_PREFIX: &str = "校验区域";
const EXCEL_ERROR_TOKENS: [&str; 10] = [
    "#NULL!",
    "#DIV/0!",
    "#VALUE!",
    "#REF!",
    "#NAME?",
    "#NUM!",
    "#N/A",
    "#SPILL!",
    "#CALC!",
    "#GETTING_DATA",
];

#[derive(Clone, Debug, Serialize)]
pub struct S1ValidationFormulaError {
    pub sheet: String,
    pub cell: String,
    pub value: String,
    pub formula: String,
    pub source: String,
    pub origin: String,
    pub template_paths: Vec<String>,
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub formula_error_tokens: Vec<String>,
}

#[derive(Clone, Debug, Serialize)]
pub struct S1ValidationFormulaScanReport {
    pub status: &'static str,
    pub reader: &'static str,
    pub reason: String,
    pub template_count: usize,
    pub formula_cells_scanned: usize,
    pub missing_sheets: Vec<String>,
    pub errors: Vec<S1ValidationFormulaError>,
    pub read_time_ms: u128,
    pub elapsed_ms: u128,
}

#[derive(Clone, Debug)]
struct ValidationFormulaCell {
    sheet: String,
    cell: String,
    formula: String,
}

#[derive(Clone, Debug)]
pub struct S1ValidationFormulaScanPlan {
    template_paths: Vec<PathBuf>,
    formulas: BTreeMap<(String, String), String>,
    template_errors: Vec<S1ValidationFormulaError>,
    template_read_time_ms: u128,
}

impl S1ValidationFormulaScanPlan {
    /// Load formula coordinates only from `校验区域*` names in the selected
    /// layout templates. Unsupported name expressions are ignored just as in
    /// the Python baseline, whose named-region helper skips non-area names.
    pub fn load(template_paths: &[PathBuf]) -> Result<Self> {
        let started = Instant::now();
        let mut unique_paths = Vec::new();
        let mut seen_paths = BTreeSet::new();
        for path in template_paths {
            let resolved = fs::canonicalize(path)
                .with_context(|| format!("读取校验公式布局模板 {}", path.display()))?;
            if seen_paths.insert(path_key(&resolved)) {
                unique_paths.push(resolved);
            }
        }

        let mut formulas = BTreeMap::<(String, String), String>::new();
        let mut template_errors = Vec::<S1ValidationFormulaError>::new();
        let mut error_indices = HashMap::<(String, String, String), usize>::new();
        for template_path in &unique_paths {
            let names = read_defined_names(template_path)
                .with_context(|| format!("读取模板命名区域 {}", template_path.display()))?;
            let workbook = read_workbook(template_path)
                .with_context(|| format!("读取模板校验公式 {}", template_path.display()))?;
            let areas = validation_areas(&names);
            let validation_cells = formulas_in_areas(&workbook, &areas);

            for cell in validation_cells {
                let key = (cell.sheet.clone(), cell.cell.clone());
                if let Some(previous) = formulas.get(&key) {
                    ensure!(
                        previous == &cell.formula,
                        "多个布局模板对校验公式定义不一致：{}!{}",
                        cell.sheet,
                        cell.cell
                    );
                } else {
                    formulas.insert(key.clone(), cell.formula.clone());
                }

                let tokens = formula_error_tokens(&cell.formula);
                let cached_value = workbook
                    .sheets
                    .iter()
                    .find(|sheet| sheet.name == cell.sheet)
                    .and_then(|sheet| read_cell_value(sheet, &cell.cell));
                let cached_error = cached_value.as_ref().and_then(excel_error_value);
                if tokens.is_empty() && cached_error.is_none() {
                    continue;
                }
                let value = tokens.first().cloned().or(cached_error).unwrap_or_default();
                let error = S1ValidationFormulaError {
                    sheet: cell.sheet,
                    cell: cell.cell,
                    value,
                    formula: cell.formula,
                    source: String::new(),
                    origin: "template_formula".to_owned(),
                    template_paths: vec![template_path.to_string_lossy().into_owned()],
                    formula_error_tokens: tokens,
                };
                merge_scan_error(&mut template_errors, &mut error_indices, &error, false);
            }
        }

        Ok(Self {
            template_paths: unique_paths,
            formulas,
            template_errors,
            template_read_time_ms: started.elapsed().as_millis(),
        })
    }

    /// Compare the formula cells selected by the template with stored source
    /// values. No formula calculation or workbook modification is performed.
    pub fn scan(&self, source_path: &Path) -> Result<S1ValidationFormulaScanReport> {
        let started = Instant::now();
        if self.template_paths.is_empty() {
            return Ok(S1ValidationFormulaScanReport {
                status: "SKIPPED",
                reader: READER,
                reason: "本次规则没有读取布局模板；源工作簿也不提供可信的校验区域定义".to_owned(),
                template_count: 0,
                formula_cells_scanned: 0,
                missing_sheets: Vec::new(),
                errors: Vec::new(),
                read_time_ms: 0,
                elapsed_ms: started.elapsed().as_millis(),
            });
        }

        let read_started = Instant::now();
        let source = read_workbook(source_path)
            .with_context(|| format!("读取源工作簿校验公式结果 {}", source_path.display()))?;
        let read_time_ms = read_started.elapsed().as_millis();
        let source_display = resolved_path(source_path);
        let mut errors = self.template_errors.clone();
        let mut error_indices = errors
            .iter()
            .enumerate()
            .map(|(index, error)| {
                (
                    (error.sheet.clone(), error.cell.clone(), error.value.clone()),
                    index,
                )
            })
            .collect::<HashMap<_, _>>();
        let mut missing_sheets = BTreeSet::new();

        for ((sheet_name, cell_address), formula) in &self.formulas {
            let Some(sheet) = source.sheets.iter().find(|sheet| sheet.name == *sheet_name) else {
                missing_sheets.insert(sheet_name.clone());
                continue;
            };
            let Some(value) =
                read_cell_value(sheet, cell_address).and_then(|value| excel_error_value(&value))
            else {
                continue;
            };
            let error = S1ValidationFormulaError {
                sheet: sheet_name.clone(),
                cell: cell_address.clone(),
                value,
                formula: formula.clone(),
                source: source_display.clone(),
                origin: "source_cached_value".to_owned(),
                template_paths: Vec::new(),
                formula_error_tokens: Vec::new(),
            };
            merge_scan_error(&mut errors, &mut error_indices, &error, true);
        }

        for error in &mut errors {
            if error.source.is_empty() {
                error.source = source_display.clone();
            }
        }

        Ok(S1ValidationFormulaScanReport {
            status: "COMPLETED",
            reader: READER,
            reason: String::new(),
            template_count: self.template_paths.len(),
            formula_cells_scanned: self.formulas.len(),
            missing_sheets: missing_sheets.into_iter().collect(),
            errors,
            read_time_ms,
            elapsed_ms: started.elapsed().as_millis(),
        })
    }

    pub fn template_read_time_ms(&self) -> u128 {
        self.template_read_time_ms
    }

    pub fn template_count(&self) -> usize {
        self.template_paths.len()
    }
}

/// Convert scanner results into the same issue stream that enters S1 history
/// classification and result-workbook output.
pub fn issues_from_validation_formula_scan(
    report: &S1ValidationFormulaScanReport,
    source_path: &Path,
    audit_path: &Path,
    metadata: &S1FormulaMetadata,
) -> Vec<Issue> {
    if report.status != "COMPLETED" {
        return Vec::new();
    }

    let source_display = resolved_path(source_path);
    let audit_display = resolved_path(audit_path);
    let mut issues = report
        .errors
        .iter()
        .map(|error| {
            let message = match error.origin.as_str() {
                "template_formula" => format!(
                    "模板校验公式本身含 Excel 错误值或引用 {}；原公式：{}",
                    error.value, error.formula
                ),
                "template_formula_and_source_cached_value" => format!(
                    "模板校验公式及源文件结果均包含 Excel 错误值 {}；原公式：{}",
                    error.value, error.formula
                ),
                _ => format!(
                    "模板校验公式在源文件中返回 Excel 错误值 {}；原公式：{}",
                    error.value, error.formula
                ),
            };
            let issue_id = formula_error_issue_id(source_path, &error.sheet, &error.cell);
            Issue {
                issue_id: issue_id.clone(),
                period: metadata.period.clone(),
                batch_id: metadata.batch_id.clone(),
                audit_time: metadata.audit_time.clone(),
                triggered: true,
                status: String::new(),
                first_seen_period: String::new(),
                previous_seen_period: String::new(),
                consecutive_count: 1,
                org_code: metadata.org_code.clone(),
                org_name: metadata.org_name.clone(),
                report_code: String::new(),
                sheet_name: error.sheet.clone(),
                rule_id: format!("TEMPLATE_FORMULA_ERROR:{}!{}", error.sheet, error.cell),
                severity: "错误".to_owned(),
                formula_cell: error.cell.clone(),
                target_cell: error.cell.clone(),
                target_value: Value::String(error.value.clone()),
                formula_result: Value::String(error.value.clone()),
                message: message.clone(),
                source_file: source_display.clone(),
                audit_file: audit_display.clone(),
                check_field: "模板校验公式结果".to_owned(),
                check_type: "校验公式错误".to_owned(),
                comparison_value: Value::String(String::new()),
                reference_value: Value::String(String::new()),
                difference_value: Value::String(String::new()),
                detail: format!(
                    "来源：{}；模板：{}；公式：{}",
                    error.origin,
                    error.template_paths.join("、"),
                    error.formula
                ),
                institution_feedback: String::new(),
                auditor_opinion: String::new(),
            }
        })
        .collect::<Vec<_>>();

    for sheet_name in &report.missing_sheets {
        let issue_id = scan_issue_id(source_path, sheet_name, "模板校验公式扫描失败");
        let message = format!("本次源文件缺少模板校验公式所在工作表：{sheet_name}");
        issues.push(Issue {
            issue_id: issue_id.clone(),
            period: metadata.period.clone(),
            batch_id: metadata.batch_id.clone(),
            audit_time: metadata.audit_time.clone(),
            triggered: true,
            status: String::new(),
            first_seen_period: String::new(),
            previous_seen_period: String::new(),
            consecutive_count: 1,
            org_code: metadata.org_code.clone(),
            org_name: metadata.org_name.clone(),
            report_code: String::new(),
            sheet_name: sheet_name.clone(),
            rule_id: format!("TEMPLATE_FORMULA_SCAN:{sheet_name}"),
            severity: "错误".to_owned(),
            formula_cell: String::new(),
            target_cell: String::new(),
            target_value: Value::String(String::new()),
            formula_result: Value::String(String::new()),
            message: message.clone(),
            source_file: source_display.clone(),
            audit_file: audit_display.clone(),
            check_field: "模板校验公式".to_owned(),
            check_type: "校验公式扫描失败".to_owned(),
            comparison_value: Value::String(String::new()),
            reference_value: Value::String(String::new()),
            difference_value: Value::String(String::new()),
            detail: message,
            institution_feedback: String::new(),
            auditor_opinion: String::new(),
        });
    }
    issues
}

pub fn validation_formula_scan_failure_issue(
    source_path: &Path,
    audit_path: &Path,
    metadata: &S1FormulaMetadata,
    error: &str,
) -> Issue {
    let issue_id = scan_issue_id(source_path, "", "模板校验公式扫描失败");
    let message = format!("模板校验公式扫描失败：{error}");
    Issue {
        issue_id: issue_id.clone(),
        period: metadata.period.clone(),
        batch_id: metadata.batch_id.clone(),
        audit_time: metadata.audit_time.clone(),
        triggered: true,
        status: String::new(),
        first_seen_period: String::new(),
        previous_seen_period: String::new(),
        consecutive_count: 1,
        org_code: metadata.org_code.clone(),
        org_name: metadata.org_name.clone(),
        report_code: String::new(),
        sheet_name: String::new(),
        rule_id: "TEMPLATE_FORMULA_SCAN:FAILED".to_owned(),
        severity: "错误".to_owned(),
        formula_cell: String::new(),
        target_cell: String::new(),
        target_value: Value::String(String::new()),
        formula_result: Value::String(String::new()),
        message: message.clone(),
        source_file: resolved_path(source_path),
        audit_file: resolved_path(audit_path),
        check_field: "模板校验公式".to_owned(),
        check_type: "校验公式扫描失败".to_owned(),
        comparison_value: Value::String(String::new()),
        reference_value: Value::String(String::new()),
        difference_value: Value::String(String::new()),
        detail: message,
        institution_feedback: String::new(),
        auditor_opinion: String::new(),
    }
}

fn validation_areas(names: &[excel_reader::DefinedNameRecord]) -> Vec<DefinedNameArea> {
    names
        .iter()
        .filter(|name| {
            name.name.starts_with(VALIDATION_NAME_PREFIX)
                && name.parse_status == DefinedNameParseStatus::Areas
        })
        .flat_map(|name| name.areas.iter())
        .filter(|area| area.sheet_kind == "worksheet")
        .cloned()
        .collect()
}

fn formulas_in_areas(
    workbook: &WorkbookDump,
    areas: &[DefinedNameArea],
) -> Vec<ValidationFormulaCell> {
    let bounds = areas
        .iter()
        .filter_map(|area| parse_bounds(&area.reference).map(|bounds| (&area.sheet_name, bounds)))
        .collect::<Vec<_>>();
    let mut formulas = BTreeMap::<(String, String), String>::new();
    for sheet in &workbook.sheets {
        for (address, formula) in &sheet.formulas {
            let Some((row, column)) = parse_cell_address(address) else {
                continue;
            };
            if bounds.iter().any(|(sheet_name, range)| {
                *sheet_name == &sheet.name
                    && (range.min_row..=range.max_row).contains(&row)
                    && (range.min_column..=range.max_column).contains(&column)
            }) {
                let formula = if formula.starts_with('=') {
                    formula.clone()
                } else {
                    format!("={formula}")
                };
                formulas.insert((sheet.name.clone(), address.clone()), formula);
            }
        }
    }
    formulas
        .into_iter()
        .map(|((sheet, cell), formula)| ValidationFormulaCell {
            sheet,
            cell,
            formula,
        })
        .collect()
}

#[derive(Clone, Copy)]
struct CellBounds {
    min_row: usize,
    min_column: usize,
    max_row: usize,
    max_column: usize,
}

fn parse_bounds(reference: &str) -> Option<CellBounds> {
    let reference = reference.replace('$', "");
    let mut parts = reference.split(':');
    let (min_row, min_column) = parse_cell_address(parts.next()?)?;
    let (max_row, max_column) = match parts.next() {
        Some(last) => parse_cell_address(last)?,
        None => (min_row, min_column),
    };
    if parts.next().is_some() || min_row > max_row || min_column > max_column {
        return None;
    }
    Some(CellBounds {
        min_row,
        min_column,
        max_row,
        max_column,
    })
}

fn parse_cell_address(value: &str) -> Option<(usize, usize)> {
    let value = value.replace('$', "");
    let split = value.find(|character: char| character.is_ascii_digit())?;
    let (column, row) = value.split_at(split);
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

fn read_cell_value(sheet: &SheetDump, address: &str) -> Option<CellValue> {
    sheet
        .values
        .get(address)
        .and_then(|value| CellValue::from_probe_json(value).ok())
}

fn excel_error_value(value: &CellValue) -> Option<String> {
    match value {
        CellValue::Error(value) => canonical_excel_error(value).map(str::to_owned),
        CellValue::Text(value)
            if EXCEL_ERROR_TOKENS
                .iter()
                .any(|token| value.trim().eq_ignore_ascii_case(token)) =>
        {
            Some(value.clone())
        }
        _ => None,
    }
}

fn canonical_excel_error(value: &str) -> Option<&'static str> {
    let normalized = value
        .trim()
        .trim_start_matches('#')
        .chars()
        .filter(|character| character.is_ascii_alphanumeric())
        .flat_map(char::to_lowercase)
        .collect::<String>();
    match normalized.as_str() {
        "null" => Some("#NULL!"),
        "div0" | "divzero" => Some("#DIV/0!"),
        "value" => Some("#VALUE!"),
        "ref" => Some("#REF!"),
        "name" => Some("#NAME?"),
        "num" => Some("#NUM!"),
        "na" | "notavailable" => Some("#N/A"),
        "spill" => Some("#SPILL!"),
        "calc" => Some("#CALC!"),
        "gettingdata" => Some("#GETTING_DATA"),
        _ => None,
    }
}

/// Return literal Excel error tokens from formula text, excluding quoted
/// string literals and retaining the first-seen order without duplicates.
pub fn formula_error_tokens(formula: &str) -> Vec<String> {
    let bytes = formula.as_bytes();
    let mut tokens = Vec::new();
    let mut seen = BTreeSet::new();
    let mut index = 0;
    while index < bytes.len() {
        if bytes[index] == b'"' {
            index += 1;
            while index < bytes.len() {
                if bytes[index] != b'"' {
                    index += 1;
                    continue;
                }
                if bytes.get(index + 1) == Some(&b'"') {
                    index += 2;
                    continue;
                }
                index += 1;
                break;
            }
            continue;
        }
        if bytes[index] == b'#' {
            let token = EXCEL_ERROR_TOKENS.iter().find(|token| {
                let end = index + token.len();
                end <= bytes.len() && bytes[index..end].eq_ignore_ascii_case(token.as_bytes())
            });
            if let Some(token) = token {
                if seen.insert(*token) {
                    tokens.push((*token).to_owned());
                }
                index += token.len();
                continue;
            }
        }
        index += 1;
    }
    tokens
}

fn merge_scan_error(
    errors: &mut Vec<S1ValidationFormulaError>,
    indices: &mut HashMap<(String, String, String), usize>,
    incoming: &S1ValidationFormulaError,
    is_source_error: bool,
) {
    let key = (
        incoming.sheet.clone(),
        incoming.cell.clone(),
        incoming.value.clone(),
    );
    if let Some(index) = indices.get(&key).copied() {
        let existing = &mut errors[index];
        if is_source_error && existing.origin == "template_formula" {
            existing.origin = "template_formula_and_source_cached_value".to_owned();
        }
        for path in &incoming.template_paths {
            if !existing.template_paths.contains(path) {
                existing.template_paths.push(path.clone());
            }
        }
        if existing.formula_error_tokens.is_empty() {
            existing
                .formula_error_tokens
                .clone_from(&incoming.formula_error_tokens);
        }
    } else {
        indices.insert(key, errors.len());
        errors.push(incoming.clone());
    }
}

fn formula_error_issue_id(source_path: &Path, sheet: &str, cell: &str) -> String {
    let source_stem = source_path
        .file_stem()
        .map(|stem| stem.to_string_lossy().into_owned())
        .unwrap_or_default();
    let source_name = normalize_template_name(&source_stem);
    [
        if source_name.is_empty() {
            "工作簿未识别".to_owned()
        } else {
            source_name
        },
        if sheet.is_empty() {
            "工作表未识别".to_owned()
        } else {
            sheet.to_owned()
        },
        if cell.is_empty() {
            "公式单元格未识别".to_owned()
        } else {
            cell.to_owned()
        },
        "校验指标未识别".to_owned(),
    ]
    .join("｜")
}

fn scan_issue_id(source_path: &Path, sheet: &str, check: &str) -> String {
    let source_stem = source_path
        .file_stem()
        .map(|stem| stem.to_string_lossy().into_owned())
        .unwrap_or_default();
    let source_name = normalize_template_name(&source_stem);
    [
        if source_name.is_empty() {
            "工作簿未识别".to_owned()
        } else {
            source_name
        },
        if sheet.is_empty() {
            "工作表未识别".to_owned()
        } else {
            sheet.to_owned()
        },
        "公式单元格未识别".to_owned(),
        check.to_owned(),
    ]
    .join("｜")
}

fn resolved_path(path: &Path) -> String {
    if let Ok(path) = fs::canonicalize(path) {
        return path.display().to_string();
    }
    if path.is_absolute() {
        return path.display().to_string();
    }
    std::env::current_dir()
        .map(|directory| directory.join(path).display().to_string())
        .unwrap_or_else(|_| PathBuf::from(path).display().to_string())
}

fn path_key(path: &Path) -> String {
    let value = path.to_string_lossy().replace('\\', "/");
    #[cfg(windows)]
    {
        value.to_lowercase()
    }
    #[cfg(not(windows))]
    {
        value
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use excel_writer::Workbook;
    use std::{fs, time::SystemTime};

    fn temporary_directory(label: &str) -> PathBuf {
        let nonce = SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .expect("system clock after Unix epoch")
            .as_nanos();
        let path = std::env::temp_dir().join(format!(
            "audit_s1_validation_formula_scan_{label}_{}_{}",
            std::process::id(),
            nonce
        ));
        fs::create_dir_all(&path).expect("create temporary directory");
        path
    }

    fn template(path: &Path) {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.set_name("校验表").expect("name template sheet");
        sheet
            .write_formula(0, 0, "=1+1")
            .expect("write formula in validation area");
        sheet
            .write_formula(0, 1, "=1/0")
            .expect("write formula outside validation area");
        workbook
            .define_name("校验区域_001", "='校验表'!$A$1:$A$1")
            .expect("define validation range");
        workbook.save(path).expect("save template");
    }

    fn source(path: &Path, include_validation_sheet: bool) {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        if include_validation_sheet {
            sheet.set_name("校验表").expect("name validation sheet");
            sheet
                .write_string(0, 0, "#DIV/0!")
                .expect("write source error value");
            sheet
                .write_string(0, 1, "#REF!")
                .expect("write outside-range error value");
        } else {
            sheet.set_name("源数据").expect("name source sheet");
        }
        workbook.save(path).expect("save source workbook");
    }

    #[test]
    fn formula_error_tokens_ignore_quoted_strings_and_deduplicate() {
        assert_eq!(
            formula_error_tokens("=IF(A1=\"#N/A\", #ref!, \"#DIV/0!\"&#REF!)"),
            ["#REF!"]
        );
        assert_eq!(
            formula_error_tokens("=\"escaped \"\"#N/A\"\" text\"+#VALUE!"),
            ["#VALUE!"]
        );
    }

    #[test]
    fn scans_only_source_cells_selected_by_template_validation_names() {
        let directory = temporary_directory("range");
        let template_path = directory.join("template.xlsx");
        let source_path = directory.join("source.xlsx");
        template(&template_path);
        source(&source_path, true);

        let plan = S1ValidationFormulaScanPlan::load(&[template_path]).expect("load scan plan");
        let report = plan.scan(&source_path).expect("scan source values");
        assert_eq!(report.status, "COMPLETED");
        assert_eq!(report.template_count, 1);
        assert_eq!(report.formula_cells_scanned, 1);
        assert!(report.missing_sheets.is_empty());
        assert_eq!(report.errors.len(), 1);
        assert_eq!(report.errors[0].sheet, "校验表");
        assert_eq!(report.errors[0].cell, "A1");
        assert_eq!(report.errors[0].value, "#DIV/0!");
        assert_eq!(report.errors[0].origin, "source_cached_value");
        fs::remove_dir_all(directory).expect("remove temporary fixtures");
    }

    #[test]
    fn reports_missing_validation_sheets_and_skips_without_templates() {
        let directory = temporary_directory("missing");
        let template_path = directory.join("template.xlsx");
        let source_path = directory.join("source.xlsx");
        template(&template_path);
        source(&source_path, false);

        let plan = S1ValidationFormulaScanPlan::load(&[template_path]).expect("load scan plan");
        let report = plan.scan(&source_path).expect("scan source values");
        assert_eq!(report.missing_sheets, ["校验表"]);

        let skipped = S1ValidationFormulaScanPlan::load(&[])
            .expect("empty scan plan")
            .scan(&source_path)
            .expect("skipped report");
        assert_eq!(skipped.status, "SKIPPED");
        assert_eq!(skipped.formula_cells_scanned, 0);
        fs::remove_dir_all(directory).expect("remove temporary fixtures");
    }

    #[test]
    fn reports_literal_template_error_tokens_but_ignores_quoted_formula_text() {
        let directory = temporary_directory("template_error");
        let template_path = directory.join("template.xlsx");
        let source_path = directory.join("source.xlsx");
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.set_name("校验表").expect("name template sheet");
        sheet
            .write_formula(0, 0, "=IF(A1=\"#N/A\",#REF!,0)")
            .expect("write literal formula error token");
        sheet
            .write_formula(1, 0, "=\"#DIV/0!\"")
            .expect("write quoted formula error token");
        workbook
            .define_name("校验区域_001", "='校验表'!$A$1:$A$2")
            .expect("define validation range");
        workbook.save(&template_path).expect("save template");
        let mut source = Workbook::new();
        source
            .add_worksheet()
            .set_name("校验表")
            .expect("name source sheet");
        source.save(&source_path).expect("save source workbook");

        let plan = S1ValidationFormulaScanPlan::load(&[template_path]).expect("load scan plan");
        let report = plan.scan(&source_path).expect("scan source values");
        assert_eq!(report.errors.len(), 1);
        assert_eq!(report.errors[0].value, "#REF!");
        assert_eq!(report.errors[0].origin, "template_formula");
        assert_eq!(report.errors[0].formula_error_tokens, ["#REF!"]);
        fs::remove_dir_all(directory).expect("remove temporary fixtures");
    }
}
