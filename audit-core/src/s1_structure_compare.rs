//! S1 template structure checks used to screen source workbooks before audit.

use crate::models::CopyRange;
use anyhow::{Context, Result, bail};
use excel_reader::{CellValue, SheetDump, read_workbook};
use serde::Serialize;
use std::{
    collections::{BTreeMap, BTreeSet},
    path::Path,
    time::Instant,
};

const STRUCTURE_MATCH_THRESHOLD: f64 = 0.90;

#[derive(Clone, Debug)]
enum StructureValue {
    Formula(String),
    Cell(CellValue),
}

#[derive(Clone, Debug)]
struct ExpectedCell {
    sheet_name: String,
    row: usize,
    column: usize,
    value: StructureValue,
}

/// Fixed, nonblank labels read from a template's configured structure ranges.
#[derive(Clone, Debug)]
pub struct S1StructureTemplate {
    required_sheets: Vec<String>,
    expected_cells: Vec<ExpectedCell>,
    pub area_count: usize,
    pub label_count: usize,
    pub read_time_ms: u128,
}

/// Value-free result for one source workbook.
#[derive(Clone, Debug, Serialize)]
pub struct S1StructureMatch {
    pub matched: bool,
    pub score: f64,
    pub matched_labels: usize,
    pub checked_labels: usize,
    pub missing_sheets: Vec<String>,
    pub read_time_ms: u128,
    pub elapsed_ms: u128,
    pub diagnostic: Option<String>,
}

/// Read the template labels covered by its mapped structure ranges once.
pub fn read_s1_structure_template(
    template_path: impl AsRef<Path>,
    ranges: &[CopyRange],
) -> Result<S1StructureTemplate> {
    read_s1_structure_template_with_required_sheets(template_path, ranges, &[])
}

/// Read structure labels and retain additional mapped sheets required by the
/// conditional-format scan, even when they have no structure labels.
pub fn read_s1_structure_template_with_required_sheets(
    template_path: impl AsRef<Path>,
    ranges: &[CopyRange],
    required_sheets: &[String],
) -> Result<S1StructureTemplate> {
    if ranges.is_empty() {
        bail!("模板没有可用的“表结构区域”命名区域，请先维护模板");
    }
    let workbook = read_workbook(template_path.as_ref())
        .with_context(|| format!("读取 S1 表结构模板 {}", template_path.as_ref().display()))?;
    let mut required_sheets = required_sheets.iter().cloned().collect::<BTreeSet<_>>();
    let mut expected_cells = Vec::new();

    for range in ranges {
        let sheet = workbook
            .sheets
            .iter()
            .find(|sheet| sheet.name == range.sheet_name)
            .with_context(|| {
                format!("表结构区域引用了模板中不存在的工作表：{}", range.sheet_name)
            })?;
        let bounds = parse_range(&range.address)
            .with_context(|| format!("无效的表结构区域：{}!{}", range.sheet_name, range.address))?;
        required_sheets.insert(range.sheet_name.clone());
        let mut area_cells = BTreeMap::new();
        collect_expected_cells(sheet, bounds, &mut area_cells)?;
        expected_cells.extend(
            area_cells
                .into_iter()
                .map(|((row, column), value)| ExpectedCell {
                    sheet_name: range.sheet_name.clone(),
                    row,
                    column,
                    value,
                }),
        );
    }

    Ok(S1StructureTemplate {
        required_sheets: required_sheets.into_iter().collect(),
        label_count: expected_cells.len(),
        expected_cells,
        area_count: ranges.len(),
        read_time_ms: workbook.elapsed_ms,
    })
}

/// Match one source workbook to the template's nonblank fixed structure cells.
///
/// The acceptance threshold and minimum-match rule follow Audit's Python
/// `validate_source_xlsx`: 90% with at least one match for fewer than five
/// checked labels, otherwise at least three matches. Missing required sheets
/// fail immediately. Merged source cells are read from their top-left anchor.
pub fn compare_s1_source_structure(
    source_path: impl AsRef<Path>,
    template: &S1StructureTemplate,
) -> Result<S1StructureMatch> {
    let started = Instant::now();
    let source_path = source_path.as_ref();
    let workbook = read_workbook(source_path)
        .with_context(|| format!("读取 S1 表结构报送文件 {}", source_path.display()))?;
    let sheet_names = workbook
        .sheets
        .iter()
        .map(|sheet| sheet.name.as_str())
        .collect::<BTreeSet<_>>();
    let missing_sheets = template
        .required_sheets
        .iter()
        .filter(|name| !sheet_names.contains(name.as_str()))
        .cloned()
        .collect::<Vec<_>>();

    if !missing_sheets.is_empty() {
        return Ok(S1StructureMatch {
            matched: false,
            score: 0.0,
            matched_labels: 0,
            checked_labels: 0,
            missing_sheets,
            read_time_ms: workbook.elapsed_ms,
            elapsed_ms: started.elapsed().as_millis(),
            diagnostic: None,
        });
    }

    if template.expected_cells.is_empty() {
        return Ok(S1StructureMatch {
            matched: false,
            score: 0.0,
            matched_labels: 0,
            checked_labels: 0,
            missing_sheets: Vec::new(),
            read_time_ms: workbook.elapsed_ms,
            elapsed_ms: started.elapsed().as_millis(),
            diagnostic: Some("模板表结构区域没有可比对的固定值".to_owned()),
        });
    }

    let sheets = workbook
        .sheets
        .iter()
        .map(|sheet| (sheet.name.as_str(), sheet))
        .collect::<BTreeMap<_, _>>();
    let merge_anchors = build_merge_anchor_map(template, &sheets);
    let mut matched_labels = 0usize;
    for expected in &template.expected_cells {
        let sheet = sheets[expected.sheet_name.as_str()];
        let (row, column) = merge_anchors
            .get(&(expected.sheet_name.clone(), expected.row, expected.column))
            .copied()
            .unwrap_or((expected.row, expected.column));
        let actual = value_at(sheet, row, column)?;
        if structure_values_equal(&expected.value, actual.as_ref()) {
            matched_labels += 1;
        }
    }

    let checked_labels = template.expected_cells.len();
    let score = matched_labels as f64 / checked_labels as f64;
    let enough_matches = if checked_labels >= 5 {
        matched_labels >= 3
    } else {
        matched_labels >= 1
    };
    Ok(S1StructureMatch {
        matched: score >= STRUCTURE_MATCH_THRESHOLD && enough_matches,
        score,
        matched_labels,
        checked_labels,
        missing_sheets: Vec::new(),
        read_time_ms: workbook.elapsed_ms,
        elapsed_ms: started.elapsed().as_millis(),
        diagnostic: None,
    })
}

fn collect_expected_cells(
    sheet: &SheetDump,
    bounds: (usize, usize, usize, usize),
    expected: &mut BTreeMap<(usize, usize), StructureValue>,
) -> Result<()> {
    // Match the Python/openpyxl structure baseline: merged cells other than
    // the top-left anchor read as blank, even if the OOXML part stores payload
    // in those coordinates. Do not count that hidden content as a
    // required template label.
    let merged_ranges = sheet
        .merges
        .iter()
        .filter_map(|address| parse_range(address))
        .collect::<Vec<_>>();
    let mut cells = BTreeMap::<(usize, usize), StructureValue>::new();
    for (address, value) in &sheet.values {
        let Some((row, column)) = parse_cell(address) else {
            continue;
        };
        if within(row, column, bounds) && !is_merged_non_anchor_cell(row, column, &merged_ranges) {
            let value = CellValue::from_probe_json(value)
                .with_context(|| format!("解码模板结构单元格 {address}"))?;
            if !is_blank(&value) {
                cells.insert((row, column), StructureValue::Cell(value));
            }
        }
    }
    for (address, formula) in &sheet.formulas {
        let Some((row, column)) = parse_cell(address) else {
            continue;
        };
        if within(row, column, bounds) && !is_merged_non_anchor_cell(row, column, &merged_ranges) {
            cells.insert(
                (row, column),
                StructureValue::Formula(normalize_formula(formula)),
            );
        }
    }
    for (coordinate, value) in cells {
        expected.insert(coordinate, value);
    }
    Ok(())
}

fn is_merged_non_anchor_cell(
    row: usize,
    column: usize,
    merged_ranges: &[(usize, usize, usize, usize)],
) -> bool {
    merged_ranges
        .iter()
        .any(|bounds| within(row, column, *bounds) && (row, column) != (bounds.1, bounds.0))
}

fn value_at(sheet: &SheetDump, row: usize, column: usize) -> Result<Option<StructureValue>> {
    let address = cell_address(row, column);
    if let Some(formula) = sheet.formulas.get(&address) {
        return Ok(Some(StructureValue::Formula(normalize_formula(formula))));
    }
    sheet
        .values
        .get(&address)
        .map(|value| {
            CellValue::from_probe_json(value)
                .map(StructureValue::Cell)
                .with_context(|| format!("解码报送结构单元格 {address}"))
        })
        .transpose()
}

fn build_merge_anchor_map(
    template: &S1StructureTemplate,
    sheets: &BTreeMap<&str, &SheetDump>,
) -> BTreeMap<(String, usize, usize), (usize, usize)> {
    let mut result = BTreeMap::new();
    let requested = template
        .expected_cells
        .iter()
        .map(|cell| (cell.sheet_name.as_str(), cell.row, cell.column))
        .collect::<Vec<_>>();
    for (sheet_name, sheet) in sheets {
        for merge in &sheet.merges {
            let Some(bounds) = parse_range(merge) else {
                continue;
            };
            let anchor = (bounds.1, bounds.0);
            for (requested_sheet, row, column) in &requested {
                if *requested_sheet == *sheet_name && within(*row, *column, bounds) {
                    result.insert(((*sheet_name).to_owned(), *row, *column), anchor);
                }
            }
        }
    }
    result
}

fn structure_values_equal(expected: &StructureValue, actual: Option<&StructureValue>) -> bool {
    let Some(actual) = actual else {
        return false;
    };
    match (expected, actual) {
        (StructureValue::Formula(left), StructureValue::Formula(right)) => left == right,
        (StructureValue::Cell(left), StructureValue::Cell(right)) => cell_values_equal(left, right),
        _ => false,
    }
}

fn cell_values_equal(left: &CellValue, right: &CellValue) -> bool {
    match (left, right) {
        (CellValue::Text(left), CellValue::Text(right)) => left == right,
        (CellValue::Integer(left), CellValue::Integer(right)) => left == right,
        (CellValue::Integer(left), CellValue::Number(right)) => integer_number_equal(*left, *right),
        (CellValue::Number(left), CellValue::Integer(right)) => integer_number_equal(*right, *left),
        (CellValue::Number(left), CellValue::Number(right))
        | (CellValue::DateTimeSerial(left), CellValue::DateTimeSerial(right)) => left == right,
        (CellValue::Boolean(left), CellValue::Boolean(right)) => left == right,
        (CellValue::DateTimeIso(left), CellValue::DateTimeIso(right))
        | (CellValue::DurationIso(left), CellValue::DurationIso(right))
        | (CellValue::Error(left), CellValue::Error(right)) => left == right,
        _ => false,
    }
}

fn integer_number_equal(integer: i64, number: f64) -> bool {
    const I64_UPPER_EXCLUSIVE: f64 = 9_223_372_036_854_775_808.0;
    number.is_finite()
        && number.fract() == 0.0
        && number >= i64::MIN as f64
        && number < I64_UPPER_EXCLUSIVE
        && number as i64 == integer
}

fn is_blank(value: &CellValue) -> bool {
    matches!(value, CellValue::Empty)
        || matches!(value, CellValue::Text(text) if text.trim().is_empty())
}

fn normalize_formula(formula: &str) -> String {
    if formula.starts_with('=') {
        formula.to_owned()
    } else {
        format!("={formula}")
    }
}

fn parse_range(value: &str) -> Option<(usize, usize, usize, usize)> {
    let value = value.replace('$', "");
    let mut parts = value.split(':');
    let first = parse_cell(parts.next()?)?;
    let second = match parts.next() {
        Some(value) => parse_cell(value)?,
        None => first,
    };
    if parts.next().is_some() || first.0 > second.0 || first.1 > second.1 {
        return None;
    }
    Some((first.1, first.0, second.1, second.0))
}

fn parse_cell(value: &str) -> Option<(usize, usize)> {
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

fn within(row: usize, column: usize, bounds: (usize, usize, usize, usize)) -> bool {
    row >= bounds.1 && row <= bounds.3 && column >= bounds.0 && column <= bounds.2
}

fn cell_address(row: usize, column: usize) -> String {
    let mut column = column;
    let mut letters = String::new();
    while column > 0 {
        let remainder = (column - 1) % 26;
        letters.insert(0, char::from(b'A' + remainder as u8));
        column = (column - 1) / 26;
    }
    format!("{letters}{row}")
}

#[cfg(test)]
mod tests {
    use super::*;
    use excel_writer::{Format, Workbook};
    use std::{
        fs,
        path::PathBuf,
        time::{SystemTime, UNIX_EPOCH},
    };

    fn temp_dir() -> PathBuf {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock after Unix epoch")
            .as_nanos();
        let directory = std::env::temp_dir().join(format!(
            "audit_s1_structure_{}_{}",
            std::process::id(),
            nonce
        ));
        fs::create_dir_all(&directory).expect("create structure test directory");
        directory
    }

    fn write_workbook(path: &Path, values: &[(&str, &str)]) {
        write_workbook_named(path, "Report", values);
    }

    fn write_workbook_named(path: &Path, sheet_name: &str, values: &[(&str, &str)]) {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.set_name(sheet_name).expect("set worksheet name");
        for (address, value) in values {
            let (row, column) = parse_cell(address).expect("valid test cell address");
            sheet
                .write_string((row - 1) as u32, (column - 1) as u16, *value)
                .expect("write structure label");
        }
        workbook.save(path).expect("save structure workbook");
    }

    #[test]
    fn template_structure_ignores_payload_in_merged_non_anchor_cells() {
        let sheet = SheetDump {
            name: "Report".to_owned(),
            values: BTreeMap::from([
                (
                    "B1".to_owned(),
                    serde_json::to_value(CellValue::Text("first anchor".to_owned()))
                        .expect("serialize first anchor"),
                ),
                (
                    "C1".to_owned(),
                    serde_json::to_value(CellValue::Text("hidden merged value".to_owned()))
                        .expect("serialize merged non-anchor value"),
                ),
                (
                    "E1".to_owned(),
                    serde_json::to_value(CellValue::Text("second anchor".to_owned()))
                        .expect("serialize second anchor"),
                ),
            ]),
            formulas: BTreeMap::from([
                ("D1".to_owned(), "=1+2".to_owned()),
                ("F1".to_owned(), "=3+4".to_owned()),
            ]),
            formula_status: "available".to_owned(),
            merges: vec!["B1:C1".to_owned(), "E1:F1".to_owned()],
            merge_status: "available".to_owned(),
            comments: BTreeMap::new(),
            comments_status: "not_requested".to_owned(),
            conditional_formatting: Vec::new(),
            conditional_formatting_status: "not_requested".to_owned(),
            data_validations: Vec::new(),
            data_validations_status: "not_requested".to_owned(),
        };
        let mut expected = BTreeMap::new();

        collect_expected_cells(
            &sheet,
            parse_range("B1:F1").expect("valid area"),
            &mut expected,
        )
        .expect("collect visible structure cells");

        assert_eq!(
            expected.keys().copied().collect::<Vec<_>>(),
            vec![(1, 2), (1, 4), (1, 5)]
        );
        assert!(matches!(
            expected.get(&(1, 4)),
            Some(StructureValue::Formula(formula)) if formula == "=1+2"
        ));

        let mut non_anchor_only = BTreeMap::new();
        collect_expected_cells(
            &sheet,
            parse_range("C1:C1").expect("valid non-anchor-only area"),
            &mut non_anchor_only,
        )
        .expect("ignore a merged non-anchor when its anchor is outside the area");
        assert!(non_anchor_only.is_empty());
    }

    #[test]
    fn accepts_at_ninety_percent_when_minimum_match_count_is_met() {
        let directory = temp_dir();
        let template_path = directory.join("template.xlsx");
        let source_path = directory.join("source.xlsx");
        let template_values = [
            ("A1", "label one"),
            ("B1", "label two"),
            ("C1", "label three"),
            ("D1", "label four"),
            ("E1", "label five"),
            ("F1", "label six"),
            ("G1", "label seven"),
            ("H1", "label eight"),
            ("I1", "label nine"),
            ("J1", "label ten"),
        ];
        write_workbook(&template_path, &template_values);
        let mut source_values = template_values.to_vec();
        source_values[9] = ("J1", "different");
        write_workbook(&source_path, &source_values);
        let template = read_s1_structure_template(
            &template_path,
            &[CopyRange {
                sheet_name: "Report".to_owned(),
                address: "A1:J1".to_owned(),
            }],
        )
        .expect("read template labels");

        let result =
            compare_s1_source_structure(&source_path, &template).expect("compare source labels");
        assert!(result.matched);
        assert_eq!(result.matched_labels, 9);
        assert_eq!(result.checked_labels, 10);
        assert_eq!(result.score, 0.9);
        fs::remove_dir_all(directory).expect("remove structure test directory");
    }

    #[test]
    fn rejects_missing_sheet_and_blank_template_structure() {
        let directory = temp_dir();
        let template_path = directory.join("template.xlsx");
        let source_path = directory.join("source.xlsx");
        write_workbook(&template_path, &[("A1", "label")]);
        write_workbook_named(&source_path, "Other", &[]);
        let template = read_s1_structure_template(
            &template_path,
            &[CopyRange {
                sheet_name: "Report".to_owned(),
                address: "A1".to_owned(),
            }],
        )
        .expect("read template labels");
        let result =
            compare_s1_source_structure(&source_path, &template).expect("compare source labels");
        assert!(!result.matched);
        assert_eq!(result.missing_sheets, vec!["Report"]);

        let empty_template_path = directory.join("empty-template.xlsx");
        let empty_source_path = directory.join("empty-source.xlsx");
        write_workbook(&empty_template_path, &[]);
        write_workbook(&empty_source_path, &[]);
        let empty_template = read_s1_structure_template(
            &empty_template_path,
            &[CopyRange {
                sheet_name: "Report".to_owned(),
                address: "A1".to_owned(),
            }],
        )
        .expect("read blank template area");
        let result = compare_s1_source_structure(&empty_source_path, &empty_template)
            .expect("compare blank template");
        assert!(!result.matched);
        assert_eq!(result.checked_labels, 0);
        assert!(result.diagnostic.is_some());
        fs::remove_dir_all(directory).expect("remove structure test directory");
    }

    #[test]
    fn reads_values_from_source_merge_anchor() {
        let directory = temp_dir();
        let template_path = directory.join("template.xlsx");
        let source_path = directory.join("source.xlsx");
        write_workbook(&template_path, &[("B1", "merged heading")]);
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.set_name("Report").expect("set worksheet name");
        sheet
            .merge_range(0, 0, 0, 1, "merged heading", &Format::new())
            .expect("merge source heading");
        workbook.save(&source_path).expect("save merged source");
        let template = read_s1_structure_template(
            &template_path,
            &[CopyRange {
                sheet_name: "Report".to_owned(),
                address: "B1".to_owned(),
            }],
        )
        .expect("read template labels");

        let result =
            compare_s1_source_structure(&source_path, &template).expect("compare merged source");
        assert!(result.matched);
        assert_eq!(result.matched_labels, 1);
        fs::remove_dir_all(directory).expect("remove structure test directory");
    }

    #[test]
    fn counts_overlapping_configured_areas_independently() {
        let directory = temp_dir();
        let template_path = directory.join("template.xlsx");
        let source_path = directory.join("source.xlsx");
        let values = [("A1", "first"), ("B1", "shared"), ("C1", "last")];
        write_workbook(&template_path, &values);
        write_workbook(&source_path, &values);
        let ranges = [
            CopyRange {
                sheet_name: "Report".to_owned(),
                address: "A1:B1".to_owned(),
            },
            CopyRange {
                sheet_name: "Report".to_owned(),
                address: "B1:C1".to_owned(),
            },
        ];
        let template = read_s1_structure_template(&template_path, &ranges)
            .expect("read overlapping structure areas");
        let result = compare_s1_source_structure(&source_path, &template)
            .expect("compare overlapping structure areas");

        assert_eq!(template.label_count, 4);
        assert_eq!(result.checked_labels, 4);
        assert_eq!(result.matched_labels, 4);
        assert!(result.matched);
        fs::remove_dir_all(directory).expect("remove structure test directory");
    }

    #[test]
    fn requires_mapped_scan_sheets_even_without_structure_labels() {
        let directory = temp_dir();
        let template_path = directory.join("template.xlsx");
        let source_path = directory.join("source.xlsx");
        write_workbook(&template_path, &[("A1", "fixed label")]);
        write_workbook(&source_path, &[("A1", "fixed label")]);
        let structure_ranges = [CopyRange {
            sheet_name: "Report".to_owned(),
            address: "A1".to_owned(),
        }];
        let required_sheets = ["ConditionalScan".to_owned()];
        let template = read_s1_structure_template_with_required_sheets(
            &template_path,
            &structure_ranges,
            &required_sheets,
        )
        .expect("read structure template with required scan sheet");
        let result = compare_s1_source_structure(&source_path, &template)
            .expect("compare required scan sheets");

        assert!(!result.matched);
        assert_eq!(result.missing_sheets, vec!["ConditionalScan"]);
        assert_eq!(result.checked_labels, 0);
        fs::remove_dir_all(directory).expect("remove structure test directory");
    }

    #[test]
    fn compares_formula_text_instead_of_formula_cache_values() {
        let directory = temp_dir();
        let template_path = directory.join("template.xlsx");
        let source_path = directory.join("source.xlsx");
        for (path, formula) in [(&template_path, "=1+2"), (&source_path, "=1+3")] {
            let mut workbook = Workbook::new();
            let sheet = workbook.add_worksheet();
            sheet
                .set_name("Report")
                .expect("set formula test worksheet name");
            sheet
                .write_formula(0, 0, formula)
                .expect("write structure formula");
            workbook.save(path).expect("save formula test workbook");
        }
        let template = read_s1_structure_template(
            &template_path,
            &[CopyRange {
                sheet_name: "Report".to_owned(),
                address: "A1".to_owned(),
            }],
        )
        .expect("read formula structure label");
        let result = compare_s1_source_structure(&source_path, &template)
            .expect("compare formula structure label");

        assert!(!result.matched);
        assert_eq!(result.checked_labels, 1);
        assert_eq!(result.matched_labels, 0);
        fs::remove_dir_all(directory).expect("remove structure test directory");
    }
}
