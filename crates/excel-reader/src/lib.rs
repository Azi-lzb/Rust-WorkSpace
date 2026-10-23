use anyhow::{bail, ensure, Context, Result};
use calamine::{open_workbook_auto, Data, Reader, Sheets};
use roxmltree::Document;
use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use std::{collections::BTreeMap, fs, io::Read, path::Path};
use zip::ZipArchive;

const SPREADSHEET_MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const OFFICE_REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

/// Engine-independent representation of a cell value returned by the reader.
/// The tagged JSON form is part of the existing WorkbookDump probe contract.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(tag = "type", content = "value")]
pub enum CellValue {
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
}

impl From<&Data> for CellValue {
    fn from(data: &Data) -> Self {
        match data {
            Data::Empty => Self::Empty,
            Data::String(value) => Self::Text(value.clone()),
            Data::Float(value) => Self::Number(*value),
            Data::Int(value) => Self::Integer(*value),
            Data::Bool(value) => Self::Boolean(*value),
            Data::DateTime(value) => Self::DateTimeSerial(value.as_f64()),
            Data::DateTimeIso(value) => Self::DateTimeIso(value.clone()),
            Data::DurationIso(value) => Self::DurationIso(value.clone()),
            Data::Error(value) => Self::Error(format!("{value:?}")),
        }
    }
}

impl CellValue {
    /// Decode the tagged JSON representation emitted by `WorkbookDump`.
    /// Non-finite numbers are emitted by `json!` with a null value; decode those
    /// as NaN so callers can consistently classify them as non-finite.
    pub fn from_probe_json(value: &Value) -> Result<Self> {
        if value.get("type").and_then(Value::as_str) == Some("number")
            && value.get("value").is_some_and(Value::is_null)
        {
            return Ok(Self::Number(f64::NAN));
        }
        serde_json::from_value(value.clone()).context("decode workbook cell JSON")
    }

    fn to_probe_json(self) -> Option<Value> {
        Some(match self {
            Self::Empty => return None,
            Self::Text(value) => json!({"type":"string", "value":value}),
            Self::Number(value) => json!({"type":"number", "value":value}),
            Self::Integer(value) => json!({"type":"integer", "value":value}),
            Self::Boolean(value) => json!({"type":"boolean", "value":value}),
            Self::DateTimeSerial(value) => json!({"type":"datetime_serial", "value":value}),
            Self::DateTimeIso(value) => json!({"type":"datetime_iso", "value":value}),
            Self::DurationIso(value) => json!({"type":"duration_iso", "value":value}),
            Self::Error(value) => json!({"type":"error", "value":value}),
        })
    }
}

#[derive(Debug, Serialize)]
pub struct SheetDump {
    pub name: String,
    pub values: BTreeMap<String, Value>,
    pub formulas: BTreeMap<String, String>,
    pub formula_status: String,
    pub merges: Vec<String>,
    pub merge_status: String,
    pub comments: BTreeMap<String, String>,
    pub comments_status: String,
    pub conditional_formatting: Vec<Value>,
    pub conditional_formatting_status: String,
    pub data_validations: Vec<Value>,
    pub data_validations_status: String,
}

#[derive(Debug, Serialize)]
pub struct WorkbookDump {
    pub schema_version: u32,
    pub reader: &'static str,
    pub elapsed_ms: u128,
    pub sheets: Vec<SheetDump>,
}

/// A lightweight projection of stored values from a sheet.
///
/// Values use absolute A1 coordinates and preserve the same typed `CellValue`
/// interpretation as `WorkbookDump`. The calamine `Range` is still materialized
/// for the used sheet, but callers can avoid serializing unrelated columns and
/// reading OOXML metadata they do not need.
#[derive(Clone, Debug, PartialEq)]
pub struct SheetValues {
    pub name: String,
    pub values: BTreeMap<String, CellValue>,
    /// Absolute worksheet extent reported by Calamine, in one-based rows and columns.
    /// Empty worksheets report `(0, 0)`; callers that need application-style defaults
    /// should apply those at their own contract boundary.
    pub max_row: usize,
    pub max_column: usize,
}

/// Cell and merged-range bounds stored in one XLSX worksheet part. This is a
/// small OOXML metadata projection used when Calamine's value range does not
/// include styled blank cells or merge-only coordinates. Zero min/max values
/// mean that the worksheet contains no stored cells and no merged ranges.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct WorksheetExtent {
    pub name: String,
    pub min_row: usize,
    pub min_column: usize,
    pub max_row: usize,
    pub max_column: usize,
}

#[derive(Default)]
struct OoxmlSheetMetadata {
    comments: BTreeMap<String, String>,
    conditional_formatting: Vec<Value>,
    data_validations: Vec<Value>,
}

/// The three SpreadsheetML worksheet header variants.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Ord, PartialOrd)]
pub enum HeaderVariant {
    Odd,
    Even,
    First,
}

/// Plain text from one worksheet header variant, including Excel control codes.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct WorksheetHeaderText {
    pub sheet_name: String,
    pub variant: HeaderVariant,
    pub text: String,
}

/// Header text inventory for one worksheet, in workbook sheet order.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct WorksheetHeaderInventory {
    pub sheet_name: String,
    pub headers: Vec<WorksheetHeaderText>,
}

/// Read odd/even/first-page header text from an OOXML workbook.
///
/// This is an OOXML metadata read, separate from Calamine cell-value reading.
/// Rich/nested header markup is rejected so callers cannot mistake a partial
/// text projection for the original header contract.
pub fn read_header_inventory(path: impl AsRef<Path>) -> Result<Vec<WorksheetHeaderInventory>> {
    let path = path.as_ref();
    let is_ooxml = path
        .extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| {
            matches!(
                extension.to_ascii_lowercase().as_str(),
                "xlsx" | "xlsm" | "xlam"
            )
        });
    ensure!(
        is_ooxml,
        "header metadata reader only supports XLSX/XLSM/XLAM"
    );

    let mut zip = ZipArchive::new(fs::File::open(path)?)?;
    let worksheets = ooxml_worksheet_parts(&mut zip)?;
    let mut inventory = Vec::with_capacity(worksheets.len());
    for (sheet_name, sheet_part) in worksheets {
        let xml = zip_read(&mut zip, &sheet_part)?;
        let document =
            Document::parse(&xml).with_context(|| format!("parse worksheet part {sheet_part}"))?;
        let mut headers = Vec::new();
        if let Some(header_footer) = document.root_element().descendants().find(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                && node.tag_name().name() == "headerFooter"
        }) {
            let mut seen = std::collections::BTreeSet::new();
            for header in header_footer.children().filter(|node| node.is_element()) {
                if header.tag_name().namespace() != Some(SPREADSHEET_MAIN_NS) {
                    continue;
                }
                let variant = match header.tag_name().name() {
                    "oddHeader" => HeaderVariant::Odd,
                    "evenHeader" => HeaderVariant::Even,
                    "firstHeader" => HeaderVariant::First,
                    _ => continue,
                };
                ensure!(
                    seen.insert(variant),
                    "worksheet {sheet_name:?} contains duplicate {variant:?} header elements"
                );
                if header.children().any(|child| child.is_element()) {
                    bail!(
                        "worksheet {sheet_name:?} contains unsupported nested markup in {variant:?} header"
                    );
                }
                let text = header
                    .children()
                    .filter(|child| child.is_text())
                    .filter_map(|child| child.text())
                    .collect();
                headers.push(WorksheetHeaderText {
                    sheet_name: sheet_name.clone(),
                    variant,
                    text,
                });
            }
        }
        inventory.push(WorksheetHeaderInventory {
            sheet_name,
            headers,
        });
    }
    Ok(inventory)
}

fn value_json(data: &Data) -> Option<Value> {
    CellValue::from(data).to_probe_json()
}

fn column_letters(mut col: usize) -> String {
    let mut out = String::new();
    loop {
        let rem = col % 26;
        out.insert(0, (b'A' + rem as u8) as char);
        if col < 26 {
            break;
        }
        col = col / 26 - 1;
    }
    out
}

fn coord(row: usize, col: usize) -> String {
    format!("{}{}", column_letters(col), row + 1)
}

fn dimensions_to_range(start: (u32, u32), end: (u32, u32)) -> String {
    format!(
        "{}{}:{}{}",
        column_letters(start.1 as usize),
        start.0 + 1,
        column_letters(end.1 as usize),
        end.0 + 1
    )
}

fn merges_for(
    workbook: &mut Sheets<std::io::BufReader<fs::File>>,
    name: &str,
) -> Result<(Vec<String>, &'static str)> {
    match workbook {
        Sheets::Xlsx(reader) => Ok((
            reader
                .merge_cells_by_sheet_name(name)?
                .iter()
                .map(|d| dimensions_to_range(d.start, d.end))
                .collect(),
            "read_by_calamine",
        )),
        Sheets::Xls(reader) => Ok((
            reader
                .merge_cells_by_sheet_name(name)?
                .iter()
                .map(|d| dimensions_to_range(d.start, d.end))
                .collect(),
            "read_by_calamine",
        )),
        Sheets::Xlsb(_) => Ok((
            Vec::new(),
            "unsupported_api: Calamine 0.37.0 exposes no xlsb merge reader",
        )),
        Sheets::Ods(_) => Ok((Vec::new(), "out_of_scope_format")),
    }
}

pub fn read_workbook(path: impl AsRef<Path>) -> Result<WorkbookDump> {
    let path = path.as_ref();
    validate_legacy_xls_signature(path)?;
    let started = std::time::Instant::now();
    let mut workbook =
        open_workbook_auto(path).with_context(|| format!("open {}", path.display()))?;
    let names = workbook.sheet_names().to_vec();
    let is_ooxml = path
        .extension()
        .and_then(|e| e.to_str())
        .map(|e| matches!(e.to_ascii_lowercase().as_str(), "xlsx" | "xlsm" | "xlam"))
        .unwrap_or(false);
    let metadata_by_sheet = if is_ooxml {
        extract_ooxml_metadata(path)?
    } else {
        BTreeMap::new()
    };
    let mut sheets = Vec::new();
    for name in names {
        let range = workbook
            .worksheet_range(&name)
            .with_context(|| format!("read values for {name}"))?;
        let (start_row, start_col) = range.start().unwrap_or((0, 0));
        let mut values = BTreeMap::new();
        for (local_row, local_col, value) in range.cells() {
            if let Some(v) = value_json(value) {
                values.insert(
                    coord(
                        start_row as usize + local_row,
                        start_col as usize + local_col,
                    ),
                    v,
                );
            }
        }
        let (formulas, formula_status) = match workbook.worksheet_formula(&name) {
            Ok(range) => {
                let (fr, fc) = range.start().unwrap_or((0, 0));
                (
                    range
                        .cells()
                        .map(|(r, c, v)| (coord(fr as usize + r, fc as usize + c), v.to_owned()))
                        .collect(),
                    "read_by_calamine".to_owned(),
                )
            }
            Err(error) => (BTreeMap::new(), format!("read_error:{error}")),
        };
        let (merges, merge_status) = merges_for(&mut workbook, &name)
            .with_context(|| format!("read merge regions for {name}"))?;
        let (
            comments,
            comments_status,
            conditional_formatting,
            conditional_formatting_status,
            data_validations,
            data_validations_status,
        ) = if is_ooxml {
            let metadata = metadata_by_sheet
                .get(&name)
                .with_context(|| format!("OOXML metadata missing for sheet {name}"))?;
            (
                metadata.comments.clone(),
                "read_by_namespace_aware_ooxml_supplement".to_owned(),
                metadata.conditional_formatting.clone(),
                "read_by_namespace_aware_ooxml_supplement".to_owned(),
                metadata.data_validations.clone(),
                "read_by_namespace_aware_ooxml_supplement".to_owned(),
            )
        } else {
            (
                BTreeMap::new(),
                "unsupported_api: comment adapter only supports OOXML xlsx/xlsm".to_owned(),
                Vec::new(),
                "unsupported_api: conditional-formatting adapter only supports OOXML xlsx/xlsm"
                    .to_owned(),
                Vec::new(),
                "unsupported_api: data-validation adapter only supports OOXML xlsx/xlsm".to_owned(),
            )
        };
        sheets.push(SheetDump {
            name,
            values,
            formulas,
            formula_status,
            merges,
            merge_status: merge_status.to_owned(),
            comments,
            comments_status,
            conditional_formatting,
            conditional_formatting_status,
            data_validations,
            data_validations_status,
        });
    }
    Ok(WorkbookDump {
        schema_version: 3,
        reader: "calamine",
        elapsed_ms: started.elapsed().as_millis(),
        sheets,
    })
}

/// Read worksheet names without reading worksheet cell values.
///
/// This is intended for workflows that identify workbook/sheet-level items
/// from file metadata alone. The reader and supported file formats follow the
/// same Calamine dispatch and legacy XLS signature checks as the value APIs.
pub fn read_sheet_names(path: impl AsRef<Path>) -> Result<Vec<String>> {
    let path = path.as_ref();
    validate_legacy_xls_signature(path)?;
    let workbook = open_workbook_auto(path).with_context(|| format!("open {}", path.display()))?;
    let names = workbook.sheet_names().to_vec();
    ensure!(
        !names.is_empty(),
        "workbook has no worksheets: {}",
        path.display()
    );
    Ok(names)
}

/// Return the OOXML workbook date epoch used for numeric date cells.
///
/// This is intentionally separate from value decoding so callers that compare
/// date-formatted cells with libraries exposing calendar values can normalize
/// the stored serial correctly. Non-OOXML workbooks use the 1900 system in the
/// current adapter contract and return `false`.
pub fn uses_1904_date_system(path: impl AsRef<Path>) -> Result<bool> {
    let path = path.as_ref();
    let is_ooxml = path
        .extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| {
            matches!(
                extension.to_ascii_lowercase().as_str(),
                "xlsx" | "xlsm" | "xlam"
            )
        });
    if !is_ooxml {
        return Ok(false);
    }
    let mut zip = ZipArchive::new(fs::File::open(path).context("open OOXML workbook")?)
        .context("open OOXML ZIP package")?;
    let xml =
        zip_read(&mut zip, "xl/workbook.xml").context("read workbook date-system metadata")?;
    let document = Document::parse(&xml).context("parse workbook date-system metadata")?;
    let root = document.root_element();
    ensure!(
        root.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
            && root.tag_name().name() == "workbook",
        "xl/workbook.xml has an unexpected root element"
    );
    let Some(properties) = root.descendants().find(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
            && node.tag_name().name() == "workbookPr"
    }) else {
        return Ok(false);
    };
    Ok(matches!(
        properties
            .attribute("date1904")
            .unwrap_or_default()
            .to_ascii_lowercase()
            .as_str(),
        "1" | "true" | "on"
    ))
}

/// Read the worksheet directory from the standard XLSX `xl/workbook.xml`
/// part using a namespace-aware OOXML parser. This metadata-only adapter does
/// not read cell values, and it deliberately does not fall back to Calamine
/// when the workbook part is missing or malformed.
pub fn read_xlsx_sheet_names_ooxml(path: impl AsRef<Path>) -> Result<Vec<String>> {
    let path = path.as_ref();
    ensure!(
        path.extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx")),
        "OOXML sheet-name adapter only supports .xlsx"
    );

    let mut zip = ZipArchive::new(fs::File::open(path).context("open XLSX package")?)
        .context("open XLSX ZIP package")?;
    let xml =
        zip_read(&mut zip, "xl/workbook.xml").context("read xl/workbook.xml for sheet names")?;
    let document = Document::parse(&xml).context("parse xl/workbook.xml")?;
    let root = document.root_element();
    ensure!(
        root.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
            && root.tag_name().name() == "workbook",
        "xl/workbook.xml has an unexpected root element"
    );

    let names = root
        .descendants()
        .filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                && node.tag_name().name() == "sheet"
        })
        .filter_map(|node| node.attribute("name"))
        .map(str::to_owned)
        .collect::<Vec<_>>();
    ensure!(!names.is_empty(), "workbook has no worksheets");
    Ok(names)
}

/// Read actual cell and merged-range bounds from an OOXML workbook. The
/// `<dimension>` hint is deliberately ignored because some writers leave it
/// stale or broader than the stored cell set. XLSX and XLSM are accepted for
/// this read-only metadata projection; VBA parts are not opened or changed.
pub fn read_xlsx_worksheet_extents(path: impl AsRef<Path>) -> Result<Vec<WorksheetExtent>> {
    let path = path.as_ref();
    ensure!(
        path.extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| {
                extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm")
            }),
        "worksheet extent reader only supports .xlsx and .xlsm"
    );
    let mut zip = ZipArchive::new(fs::File::open(path).context("open XLSX package")?)
        .context("open XLSX ZIP package")?;
    let worksheet_parts = ooxml_worksheet_parts(&mut zip)?;
    let mut seen = std::collections::BTreeSet::new();
    let mut extents = Vec::with_capacity(worksheet_parts.len());
    for (name, part) in worksheet_parts {
        ensure!(
            seen.insert(name.clone()),
            "duplicate worksheet name {name:?}"
        );
        let xml =
            zip_read(&mut zip, &part).with_context(|| format!("read worksheet part {part}"))?;
        let document =
            Document::parse(&xml).with_context(|| format!("parse worksheet part {part}"))?;
        let root = document.root_element();
        ensure!(
            root.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                && root.tag_name().name() == "worksheet",
            "worksheet part {part} has an unexpected root element"
        );
        let mut extent = WorksheetExtent {
            name,
            ..WorksheetExtent::default()
        };
        for cell in root.descendants().filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                && node.tag_name().name() == "c"
        }) {
            let reference = cell
                .attribute("r")
                .with_context(|| format!("cell in {part} has no A1 reference"))?;
            let (row, column) = parse_ooxml_cell_reference(reference)
                .with_context(|| format!("invalid cell reference {reference:?} in {part}"))?;
            include_extent_coordinate(&mut extent, row, column);
        }
        for merged in root.descendants().filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
                && node.tag_name().name() == "mergeCell"
        }) {
            let reference = merged
                .attribute("ref")
                .with_context(|| format!("merged range in {part} has no reference"))?;
            let (row1, column1, row2, column2) = parse_ooxml_range_reference(reference)
                .with_context(|| format!("invalid merged range {reference:?} in {part}"))?;
            include_extent_coordinate(&mut extent, row1, column1);
            include_extent_coordinate(&mut extent, row2, column2);
        }
        extents.push(extent);
    }
    Ok(extents)
}

fn include_extent_coordinate(extent: &mut WorksheetExtent, row: usize, column: usize) {
    extent.min_row = if extent.min_row == 0 {
        row
    } else {
        extent.min_row.min(row)
    };
    extent.min_column = if extent.min_column == 0 {
        column
    } else {
        extent.min_column.min(column)
    };
    extent.max_row = extent.max_row.max(row);
    extent.max_column = extent.max_column.max(column);
}

fn parse_ooxml_cell_reference(reference: &str) -> Result<(usize, usize)> {
    let reference = reference
        .rsplit_once('!')
        .map(|(_, cell)| cell)
        .unwrap_or(reference)
        .trim()
        .replace('$', "");
    let split = reference
        .find(|character: char| character.is_ascii_digit())
        .context("A1 reference has no row number")?;
    let (letters, digits) = reference.split_at(split);
    ensure!(
        !letters.is_empty() && letters.bytes().all(|byte| byte.is_ascii_alphabetic()),
        "A1 reference has an invalid column"
    );
    ensure!(
        !digits.is_empty() && digits.bytes().all(|byte| byte.is_ascii_digit()),
        "A1 reference has an invalid row"
    );
    let mut column = 0usize;
    for letter in letters.bytes() {
        column = column
            .checked_mul(26)
            .and_then(|value| {
                value.checked_add(usize::from(letter.to_ascii_uppercase() - b'A' + 1))
            })
            .context("A1 column index overflows")?;
    }
    let row = digits.parse::<usize>().context("parse A1 row number")?;
    ensure!(row > 0 && column > 0, "A1 coordinates are 1-based");
    Ok((row, column))
}

fn parse_ooxml_range_reference(reference: &str) -> Result<(usize, usize, usize, usize)> {
    let (first, last) = reference.split_once(':').unwrap_or((reference, reference));
    let (row1, column1) = parse_ooxml_cell_reference(first)?;
    let (row2, column2) = parse_ooxml_cell_reference(last)?;
    Ok((row1, column1, row2, column2))
}

/// Read stored values from the first `max_columns` columns of each worksheet.
///
/// This is a value-only API: it does not read formula text or OOXML metadata.
/// Formula cells are represented by the value available from Calamine's
/// worksheet range (normally the cached result); callers that require formula
/// text must use `read_workbook` and inspect `SheetDump::formulas` separately.
pub fn read_value_columns(path: impl AsRef<Path>, max_columns: usize) -> Result<Vec<SheetValues>> {
    read_value_columns_impl(path.as_ref(), max_columns, false)
}

/// Read stored values while representing Calamine date-formatted cells as ISO
/// datetime strings. This preserves the workbook's 1900/1904 date epoch for
/// callers that need calendar dates rather than raw serial values.
pub fn read_value_columns_with_iso_dates(
    path: impl AsRef<Path>,
    max_columns: usize,
) -> Result<Vec<SheetValues>> {
    read_value_columns_impl(path.as_ref(), max_columns, true)
}

fn read_value_columns_impl(
    path: &Path,
    max_columns: usize,
    iso_dates: bool,
) -> Result<Vec<SheetValues>> {
    anyhow::ensure!(max_columns > 0, "max_columns must be greater than zero");
    validate_legacy_xls_signature(path)?;
    let mut workbook =
        open_workbook_auto(path).with_context(|| format!("open {}", path.display()))?;
    let names = workbook.sheet_names().to_vec();
    let mut sheets = Vec::with_capacity(names.len());

    for name in names {
        let range = workbook
            .worksheet_range(&name)
            .with_context(|| format!("read values for {name}"))?;
        let (start_row, start_col) = range.start().unwrap_or((0, 0));
        let (max_row, max_column) = range
            .end()
            .map(|(row, column)| (row as usize + 1, column as usize + 1))
            .unwrap_or_default();
        let mut values = BTreeMap::new();
        for (local_row, local_col, value) in range.cells() {
            let absolute_col = start_col as usize + local_col;
            if absolute_col >= max_columns {
                continue;
            }
            if !matches!(value, Data::Empty) {
                let cell_value = match value {
                    Data::DateTime(datetime) if iso_dates && datetime.is_datetime() => {
                        let (year, month, day, hour, minute, second, millisecond) =
                            datetime.to_ymd_hms_milli();
                        CellValue::DateTimeIso(format!(
                            "{year:04}-{month:02}-{day:02}T{hour:02}:{minute:02}:{second:02}.{millisecond:03}"
                        ))
                    }
                    other => CellValue::from(other),
                };
                values.insert(
                    coord(start_row as usize + local_row, absolute_col),
                    cell_value,
                );
            }
        }
        sheets.push(SheetValues {
            name,
            values,
            max_row,
            max_column,
        });
    }
    Ok(sheets)
}

fn validate_legacy_xls_signature(path: &Path) -> Result<()> {
    let is_xls = path
        .extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| extension.eq_ignore_ascii_case("xls"));
    if !is_xls {
        return Ok(());
    }

    const CFB_OLE_SIGNATURE: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    let mut file = fs::File::open(path).with_context(|| format!("open {}", path.display()))?;
    let mut signature = [0_u8; 8];
    if file.read_exact(&mut signature).is_err() || signature != CFB_OLE_SIGNATURE {
        anyhow::bail!(
            "格式不支持：.xls 文件的 CFB/OLE 签名无效（{}）",
            path.display()
        );
    }
    Ok(())
}

fn zip_read(zip: &mut ZipArchive<fs::File>, name: &str) -> Result<String> {
    let mut file = zip
        .by_name(name)
        .with_context(|| format!("missing ZIP part {name}"))?;
    let mut text = String::new();
    file.read_to_string(&mut text)?;
    Ok(text)
}

fn norm_part(base_dir: &str, target: &str) -> String {
    let combined = if target.starts_with('/') {
        target.trim_start_matches('/').to_owned()
    } else if base_dir.is_empty() {
        target.to_owned()
    } else {
        format!("{base_dir}/{target}")
    };
    let mut parts: Vec<&str> = Vec::new();
    for p in combined.split('/') {
        match p {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            _ => parts.push(p),
        }
    }
    parts.join("/")
}

fn relationships_part(part: &str) -> String {
    let (parent, file_name) = part.rsplit_once('/').unwrap_or(("", part));
    if parent.is_empty() {
        format!("_rels/{file_name}.rels")
    } else {
        format!("{parent}/_rels/{file_name}.rels")
    }
}

fn ooxml_worksheet_parts(zip: &mut ZipArchive<fs::File>) -> Result<Vec<(String, String)>> {
    let root_rels_xml = zip_read(zip, "_rels/.rels")?;
    let root_rels_doc = Document::parse(&root_rels_xml)?;
    let workbook_target = root_rels_doc
        .descendants()
        .filter(|node| {
            node.is_element()
                && node.tag_name().namespace() == Some(PACKAGE_REL_NS)
                && node.tag_name().name() == "Relationship"
        })
        .find_map(|relationship| {
            let is_workbook = relationship
                .attribute("Type")
                .is_some_and(|kind| kind.ends_with("/officeDocument"));
            let is_external = relationship
                .attribute("TargetMode")
                .is_some_and(|mode| mode.eq_ignore_ascii_case("External"));
            if is_workbook && !is_external {
                relationship.attribute("Target")
            } else {
                None
            }
        })
        .context("package relationships do not identify an internal officeDocument part")?;
    let workbook_part = norm_part("", workbook_target);
    let workbook_xml = zip_read(zip, &workbook_part)?;
    let workbook_rels_xml = zip_read(zip, &relationships_part(&workbook_part))?;
    let workbook_doc = Document::parse(&workbook_xml)?;
    let rels_doc = Document::parse(&workbook_rels_xml)?;
    let workbook_parent = workbook_part
        .rsplit_once('/')
        .map(|(parent, _)| parent)
        .unwrap_or("");

    let mut rel_targets = BTreeMap::new();
    for relationship in rels_doc.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(PACKAGE_REL_NS)
            && node.tag_name().name() == "Relationship"
    }) {
        if let (Some(id), Some(kind), Some(target)) = (
            relationship.attribute("Id"),
            relationship.attribute("Type"),
            relationship.attribute("Target"),
        ) {
            let is_external = relationship
                .attribute("TargetMode")
                .is_some_and(|mode| mode.eq_ignore_ascii_case("External"));
            rel_targets.insert(
                id.to_owned(),
                (kind.to_owned(), target.to_owned(), is_external),
            );
        }
    }

    let mut worksheets = Vec::new();
    let mut names = std::collections::BTreeSet::new();
    let mut parts = std::collections::BTreeSet::new();
    for sheet in workbook_doc.descendants().filter(|node| {
        node.is_element()
            && node.tag_name().namespace() == Some(SPREADSHEET_MAIN_NS)
            && node.tag_name().name() == "sheet"
    }) {
        let name = sheet
            .attribute("name")
            .context("workbook sheet has no name")?;
        let id = sheet
            .attribute((OFFICE_REL_NS, "id"))
            .with_context(|| format!("workbook sheet {name:?} has no relationship id"))?;
        let (kind, target, is_external) = rel_targets
            .get(id)
            .with_context(|| format!("workbook sheet {name:?} has an unresolved relationship"))?;
        if !kind.ends_with("/worksheet") {
            continue;
        }
        ensure!(
            !*is_external,
            "workbook worksheet {name:?} has an external relationship"
        );
        let part = norm_part(workbook_parent, target);
        ensure!(
            !part.is_empty(),
            "workbook worksheet {name:?} has an empty part path"
        );
        ensure!(
            names.insert(name.to_owned()),
            "duplicate worksheet name {name:?}"
        );
        ensure!(
            parts.insert(part.clone()),
            "multiple worksheet names resolve to part {part}"
        );
        worksheets.push((name.to_owned(), part));
    }
    ensure!(
        !worksheets.is_empty(),
        "workbook contains no resolvable worksheet relationships"
    );
    Ok(worksheets)
}

fn xml_element_value(node: roxmltree::Node<'_, '_>) -> Value {
    let mut attributes = BTreeMap::new();
    for attribute in node.attributes() {
        let key = match attribute.namespace() {
            Some(namespace) => format!("{{{namespace}}}{}", attribute.name()),
            None => attribute.name().to_owned(),
        };
        attributes.insert(key, attribute.value().to_owned());
    }

    let content = node
        .children()
        .filter_map(|child| {
            if child.is_element() {
                Some(xml_element_value(child))
            } else {
                child
                    .text()
                    .filter(|text| !text.trim().is_empty())
                    .map(|text| json!({"text": text}))
            }
        })
        .collect::<Vec<_>>();

    json!({
        "namespace": node.tag_name().namespace().unwrap_or(""),
        "name": node.tag_name().name(),
        "attributes": attributes,
        "content": content,
    })
}

// Comments, conditional formatting, and data validations are not exposed by Calamine. This
// namespace-aware supplement follows package relationships and records normalized XML structure.
fn extract_ooxml_metadata(path: &Path) -> Result<BTreeMap<String, OoxmlSheetMetadata>> {
    let file = fs::File::open(path)?;
    let mut zip = ZipArchive::new(file)?;
    let worksheets = ooxml_worksheet_parts(&mut zip)?;
    let mut result = BTreeMap::new();
    for (sheet_name, sheet_part) in worksheets {
        let sheet_xml = zip_read(&mut zip, &sheet_part)?;
        let sheet_doc = Document::parse(&sheet_xml)
            .with_context(|| format!("parse worksheet part {sheet_part}"))?;
        let conditional_formatting = sheet_doc
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "conditionalFormatting")
            .map(xml_element_value)
            .collect();
        let data_validations = sheet_doc
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "dataValidations")
            .map(xml_element_value)
            .collect();

        let parent = sheet_part
            .rsplit_once('/')
            .map(|(parent, _)| parent)
            .unwrap_or("");
        let rels_part = relationships_part(&sheet_part);
        let mut comment_parts = Vec::new();
        if let Ok(sheet_rels_xml) = zip_read(&mut zip, &rels_part) {
            let sheet_rels_doc = Document::parse(&sheet_rels_xml)?;
            for rel in sheet_rels_doc
                .descendants()
                .filter(|n| n.is_element() && n.tag_name().name() == "Relationship")
            {
                let ty = rel.attribute("Type").unwrap_or_default();
                if ty.ends_with("/comments") {
                    if let Some(target) = rel.attribute("Target") {
                        comment_parts.push(norm_part(parent, target));
                    }
                }
            }
        }
        let mut comments = BTreeMap::new();
        for part in comment_parts {
            let xml = zip_read(&mut zip, &part)?;
            let doc = Document::parse(&xml)?;
            for comment in doc
                .descendants()
                .filter(|n| n.is_element() && n.tag_name().name() == "comment")
            {
                let Some(reference) = comment.attribute("ref") else {
                    continue;
                };
                let mut text = String::new();
                for t in comment
                    .descendants()
                    .filter(|n| n.is_element() && n.tag_name().name() == "t")
                {
                    if let Some(s) = t.text() {
                        text.push_str(s);
                    }
                }
                comments.insert(reference.to_owned(), text);
            }
        }
        result.insert(
            sheet_name.to_owned(),
            OoxmlSheetMetadata {
                comments,
                conditional_formatting,
                data_validations,
            },
        );
    }
    Ok(result)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;
    use std::path::PathBuf;
    use zip::{write::SimpleFileOptions, ZipWriter};

    fn nonstandard_header_fixture() -> Result<PathBuf> {
        let path = std::env::temp_dir().join(format!(
            "excel-reader-header-{}-{}.xlsx",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)?
                .as_nanos()
        ));
        let package_relationships = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="workbook" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="custom/books/main.xml"/></Relationships>"#
        );
        let workbook = format!(
            r#"<s:workbook xmlns:s="{SPREADSHEET_MAIN_NS}" xmlns:r="{OFFICE_REL_NS}"><s:sheets><s:sheet name="外汇报表" sheetId="1" r:id="worksheet"/><s:sheet name="图表页" sheetId="2" r:id="chartsheet"/></s:sheets></s:workbook>"#
        );
        let workbook_relationships = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="worksheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="../worksheets/fx.xml"/><Relationship Id="chartsheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="../charts/chart1.xml"/></Relationships>"#
        );
        let worksheet = format!(
            r#"<s:worksheet xmlns:s="{SPREADSHEET_MAIN_NS}"><s:sheetData/><s:headerFooter differentOddEven="1" differentFirst="1"><s:oddHeader>&amp;L单位：万元</s:oddHeader><s:evenHeader>&amp;C第二页</s:evenHeader><s:firstHeader/></s:headerFooter></s:worksheet>"#
        );
        let entries = [
            ("_rels/.rels", package_relationships),
            ("custom/books/main.xml", workbook),
            ("custom/books/_rels/main.xml.rels", workbook_relationships),
            ("custom/worksheets/fx.xml", worksheet),
            ("custom/charts/chart1.xml", "<chart/>".to_owned()),
        ];
        let mut archive = ZipWriter::new(fs::File::create(&path)?);
        for (name, data) in entries {
            archive.start_file(name, SimpleFileOptions::default())?;
            archive.write_all(data.as_bytes())?;
        }
        archive.finish()?;
        Ok(path)
    }

    fn sheet_name_ooxml_fixture() -> Result<PathBuf> {
        let path = std::env::temp_dir().join(format!(
            "excel-reader-sheet-names-{}-{}.xlsx",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)?
                .as_nanos()
        ));
        let workbook = format!(
            r#"<s:workbook xmlns:s="{SPREADSHEET_MAIN_NS}" xmlns:x="urn:unrelated"><s:sheets><s:sheet name="单位贷款" sheetId="1"/><s:sheet name="图表页" sheetId="2"/><x:sheet name="must-be-ignored" sheetId="3"/></s:sheets></s:workbook>"#
        );
        let mut archive = ZipWriter::new(fs::File::create(&path)?);
        archive.start_file("xl/workbook.xml", SimpleFileOptions::default())?;
        archive.write_all(workbook.as_bytes())?;
        archive.finish()?;
        Ok(path)
    }

    fn worksheet_extent_fixture() -> Result<PathBuf> {
        let path = std::env::temp_dir().join(format!(
            "excel-reader-worksheet-extents-{}-{}.xlsx",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)?
                .as_nanos()
        ));
        let package_rels = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
        );
        let workbook = format!(
            r#"<s:workbook xmlns:s="{SPREADSHEET_MAIN_NS}" xmlns:r="{OFFICE_REL_NS}"><s:sheets><s:sheet name="Layout" sheetId="1" r:id="rId1"/></s:sheets></s:workbook>"#
        );
        let workbook_rels = format!(
            r#"<Relationships xmlns="{PACKAGE_REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#
        );
        let worksheet = format!(
            r#"<s:worksheet xmlns:s="{SPREADSHEET_MAIN_NS}"><s:dimension ref="A1:XFD1048576"/><s:sheetData><s:row r="3"><s:c r="B3"><s:v>4</s:v></s:c></s:row><s:row r="8"><s:c r="G8" s="1"/></s:row></s:sheetData><s:mergeCells count="1"><s:mergeCell ref="C10:D12"/></s:mergeCells></s:worksheet>"#
        );
        let mut archive = ZipWriter::new(fs::File::create(&path)?);
        for (name, xml) in [
            ("_rels/.rels", package_rels),
            ("xl/workbook.xml", workbook),
            ("xl/_rels/workbook.xml.rels", workbook_rels),
            ("xl/worksheets/sheet1.xml", worksheet),
        ] {
            archive.start_file(name, SimpleFileOptions::default())?;
            archive.write_all(xml.as_bytes())?;
        }
        archive.finish()?;
        Ok(path)
    }

    #[test]
    fn cell_value_json_shape_remains_compatible_and_typed() {
        let cases = [
            (
                Data::String("001.20".to_owned()),
                json!({"type":"string", "value":"001.20"}),
            ),
            (Data::Float(1.0), json!({"type":"number", "value":1.0})),
            (Data::Int(42), json!({"type":"integer", "value":42})),
            (Data::Bool(true), json!({"type":"boolean", "value":true})),
        ];

        for (data, expected) in cases {
            assert_eq!(value_json(&data), Some(expected));
        }
        assert_eq!(value_json(&Data::Empty), None);
        assert_eq!(
            value_json(&Data::Float(f64::NAN)),
            Some(json!({"type":"number", "value":null}))
        );

        let decoded = CellValue::from_probe_json(&json!({
            "type": "number",
            "value": 12.5
        }))
        .unwrap();
        assert_eq!(decoded, CellValue::Number(12.5));
        let non_finite = CellValue::from_probe_json(&json!({
            "type": "number",
            "value": null
        }))
        .unwrap();
        assert!(matches!(non_finite, CellValue::Number(value) if value.is_nan()));
    }

    #[test]
    fn ooxml_sheet_name_adapter_matches_namespace_and_workbook_order() -> Result<()> {
        let fixture = sheet_name_ooxml_fixture()?;
        let names = read_xlsx_sheet_names_ooxml(&fixture)?;
        fs::remove_file(fixture)?;
        assert_eq!(names, vec!["单位贷款".to_owned(), "图表页".to_owned()]);
        Ok(())
    }

    #[test]
    fn worksheet_extents_include_styled_blanks_and_merge_coordinates_not_dimension_hint(
    ) -> Result<()> {
        let fixture = worksheet_extent_fixture()?;
        let extents = read_xlsx_worksheet_extents(&fixture)?;
        fs::remove_file(fixture)?;
        assert_eq!(
            extents,
            vec![WorksheetExtent {
                name: "Layout".to_owned(),
                min_row: 3,
                min_column: 2,
                max_row: 12,
                max_column: 7,
            }]
        );
        Ok(())
    }

    #[test]
    fn ooxml_sheet_name_adapter_rejects_other_formats_without_fallback() -> Result<()> {
        let path = std::env::temp_dir().join(format!(
            "excel-reader-sheet-names-{}.xlsm",
            std::process::id()
        ));
        fs::write(&path, b"not an xlsx")?;
        let error = read_xlsx_sheet_names_ooxml(&path).expect_err("reject unsupported format");
        fs::remove_file(path)?;
        assert!(error.to_string().contains("only supports .xlsx"));
        Ok(())
    }

    #[test]
    fn value_column_projection_matches_dump_and_keeps_absolute_coordinates() {
        let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../audit-core/tests/fixtures/synthetic_probe.xlsx");
        let dump = read_workbook(&fixture).expect("read full probe fixture");
        let projected = read_value_columns(&fixture, 3).expect("read first three columns");

        assert_eq!(projected.len(), dump.sheets.len());
        for (projection, sheet) in projected.iter().zip(&dump.sheets) {
            assert_eq!(projection.name, sheet.name);
            let expected: BTreeMap<_, _> = sheet
                .values
                .iter()
                .filter(|(address, _)| {
                    address
                        .chars()
                        .take_while(|character| character.is_ascii_alphabetic())
                        .fold(0_usize, |column, character| {
                            column * 26
                                + (character.to_ascii_uppercase() as usize - 'A' as usize + 1)
                        })
                        <= 3
                })
                .map(|(address, value)| {
                    (
                        address.clone(),
                        CellValue::from_probe_json(value).expect("decode dumped cell value"),
                    )
                })
                .collect();
            assert_eq!(projection.values, expected);
        }
    }

    #[test]
    fn invalid_legacy_xls_signature_is_rejected_as_unsupported_format() {
        let path = std::env::temp_dir().join(format!(
            "excel-reader-invalid-xls-{}.xls",
            std::process::id()
        ));
        fs::write(&path, b"<html>not an xls</html>").expect("write temporary pseudo XLS");
        let error = read_value_columns(&path, 3).expect_err("reject invalid CFB/OLE signature");
        let _ = fs::remove_file(&path);
        assert!(error.to_string().contains("格式不支持"));
        assert!(error.to_string().contains("CFB/OLE"));
    }

    #[test]
    fn normalized_ooxml_keeps_namespaces_attributes_and_text() {
        let xml = r#"<x:conditionalFormatting xmlns:x="urn:example:sheet" sqref="A1">
            <x:cfRule type="expression"><x:formula>A1&gt;0</x:formula></x:cfRule>
        </x:conditionalFormatting>"#;
        let document = Document::parse(xml).unwrap();
        let value = xml_element_value(document.root_element());

        assert_eq!(value["namespace"], "urn:example:sheet");
        assert_eq!(value["attributes"]["sqref"], "A1");
        assert_eq!(value["content"][0]["name"], "cfRule");
        assert_eq!(
            value["content"][0]["content"][0]["content"][0]["text"],
            "A1>0"
        );
    }

    #[test]
    fn relationship_targets_are_resolved_from_their_source_part() {
        assert_eq!(
            norm_part("custom/worksheets", "../comments/comments1.xml"),
            "custom/comments/comments1.xml"
        );
        assert_eq!(norm_part("", "/custom/workbook.xml"), "custom/workbook.xml");
        assert_eq!(
            relationships_part("custom/worksheets/sheet1.xml"),
            "custom/worksheets/_rels/sheet1.xml.rels"
        );
    }

    #[test]
    fn header_inventory_resolves_custom_workbook_part_and_excludes_chartsheets() -> Result<()> {
        let fixture = nonstandard_header_fixture()?;
        let inventory = read_header_inventory(&fixture)?;

        assert_eq!(inventory.len(), 1);
        assert_eq!(inventory[0].sheet_name, "外汇报表");
        assert_eq!(
            inventory[0].headers,
            vec![
                WorksheetHeaderText {
                    sheet_name: "外汇报表".to_owned(),
                    variant: HeaderVariant::Odd,
                    text: "&L单位：万元".to_owned(),
                },
                WorksheetHeaderText {
                    sheet_name: "外汇报表".to_owned(),
                    variant: HeaderVariant::Even,
                    text: "&C第二页".to_owned(),
                },
                WorksheetHeaderText {
                    sheet_name: "外汇报表".to_owned(),
                    variant: HeaderVariant::First,
                    text: String::new(),
                },
            ]
        );

        fs::remove_file(fixture)?;
        Ok(())
    }
}
