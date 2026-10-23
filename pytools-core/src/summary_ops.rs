//! Worksheet used-range and comment-template summary workflows.

use anyhow::{Context, Result};
use chrono::Local;
use excel_reader::{read_workbook, read_xlsx_worksheet_extents, SheetDump, WorksheetExtent};
use excel_writer::{Format, Workbook, Worksheet};
use serde::Serialize;
use std::{
    collections::{BTreeMap, BTreeSet},
    fs,
    path::{Path, PathBuf},
};

#[derive(Clone, Debug, Serialize)]
pub struct SummaryResult {
    pub saved: bool,
    pub rows: usize,
    pub files: usize,
    pub path: Option<PathBuf>,
}

pub fn summarize_used_ranges(sources: &[PathBuf], output_dir: &Path) -> Result<SummaryResult> {
    fs::create_dir_all(output_dir)
        .with_context(|| format!("create summary output directory {}", output_dir.display()))?;
    let mut workbooks = Vec::new();
    let mut max_col = 0usize;
    let mut rows_written = 0usize;
    let mut files = 0usize;
    for path in sources {
        if !supported(path) {
            continue;
        }
        if !path.is_file() {
            continue;
        }
        let dump = read_workbook(path)
            .with_context(|| format!("read used-range summary source {}", path.display()))?;
        let extents = if is_xlsx_or_xlsm(path) {
            read_xlsx_worksheet_extents(path)?
        } else {
            Vec::new()
        };
        let bounds_by_sheet = dump
            .sheets
            .iter()
            .map(|sheet| {
                let extent = extents.iter().find(|extent| extent.name == sheet.name);
                (sheet.name.clone(), summary_bounds(sheet, extent))
            })
            .collect::<BTreeMap<_, _>>();
        files += 1;
        for sheet in &dump.sheets {
            let (min_row, _min_col, max_row, max_col_sheet) = bounds_by_sheet[&sheet.name];
            max_col = max_col.max(max_col_sheet);
            if min_row > 0 {
                rows_written += max_row - min_row + 1;
            }
        }
        workbooks.push((path.clone(), dump, bounds_by_sheet));
    }
    if files == 0 || rows_written == 0 || max_col == 0 {
        return Ok(SummaryResult {
            saved: false,
            rows: 0,
            files,
            path: None,
        });
    }
    fs::create_dir_all(output_dir)?;
    let path = unique_output(
        output_dir,
        &format!(
            "{}_按使用区域汇总（全量）",
            Local::now().format("%Y%m%d_%H%M%S")
        ),
    )?;
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet.set_name("汇总")?;
    let header_format = Format::new().set_bold();
    sheet.write_string_with_format(0, 0, "工作簿", &header_format)?;
    sheet.write_string_with_format(0, 1, "工作表", &header_format)?;
    sheet.write_string_with_format(0, 2, "行号", &header_format)?;
    for col in 1..=max_col {
        sheet.write_string_with_format(0, (col + 2) as u16, &format!("列{col}"), &header_format)?;
    }
    sheet.set_freeze_panes(1, 0)?;
    let mut out_row = 1u32;
    for (path, dump, bounds_by_sheet) in workbooks {
        let base = path
            .file_stem()
            .map(|s| s.to_string_lossy().into_owned())
            .unwrap_or_default();
        for ws in &dump.sheets {
            let (min_row, min_col, max_row, max_col) = bounds_by_sheet[&ws.name];
            if min_row == 0 || max_row < min_row || max_col < min_col {
                continue;
            }
            for row in min_row..=max_row {
                sheet.write_string(out_row, 0, &base)?;
                sheet.write_string(out_row, 1, &ws.name)?;
                sheet.write_number(out_row, 2, row as f64)?;
                for col in 1..=max_col {
                    if col >= min_col {
                        if let Some(value) =
                            ws.values.get(&format!("{}{}", column_letters(col), row))
                        {
                            write_probe_value(sheet, out_row, (col + 2) as u16, value)?;
                        }
                    }
                }
                out_row += 1;
            }
        }
    }
    workbook
        .save(&path)
        .with_context(|| format!("save used-range summary {}", path.display()))?;
    Ok(SummaryResult {
        saved: true,
        rows: out_row.saturating_sub(1) as usize,
        files,
        path: Some(path),
    })
}

#[derive(Default)]
struct TemplateSpec {
    row_ranges: Vec<(usize, usize, usize, usize)>,
    col_ranges: Vec<(usize, usize, usize, usize)>,
    set_names: Vec<String>,
    set_addresses: Vec<String>,
    headers: Vec<String>,
}

pub fn summarize_by_comments(
    template_path: &Path,
    sources: &[PathBuf],
    output_dir: &Path,
) -> Result<SummaryResult> {
    fs::create_dir_all(output_dir)
        .with_context(|| format!("create summary output directory {}", output_dir.display()))?;
    anyhow::ensure!(
        is_xlsx_or_xlsm(template_path),
        "summary template must be .xlsx or .xlsm"
    );
    let template = read_workbook(template_path)
        .with_context(|| format!("read summary template {}", template_path.display()))?;
    let mut specs = BTreeMap::new();
    let mut spec_order = Vec::new();
    for sheet in &template.sheets {
        let comments = &sheet.comments;
        if comments.is_empty() {
            continue;
        }
        let row_ranges = parse_comment_ranges(comments, "行区域");
        let col_ranges = parse_comment_ranges(comments, "列区域");
        if row_ranges.is_empty() || col_ranges.is_empty() {
            continue;
        }
        let mut spec = TemplateSpec {
            row_ranges,
            col_ranges,
            ..Default::default()
        };
        for (addr, text) in ordered_comments(comments) {
            if let Some(name) = parse_set_name(text) {
                spec.set_names.push(name);
                spec.set_addresses.push(addr.to_owned());
            }
        }
        spec.headers = build_column_headers(sheet, &spec.col_ranges);
        spec_order.push(sheet.name.clone());
        specs.insert(sheet.name.clone(), spec);
    }
    if specs.is_empty() {
        anyhow::bail!("template has no supported comment-marked row/column regions");
    }
    let mut data: BTreeMap<String, Vec<Vec<serde_json::Value>>> = BTreeMap::new();
    let mut headers_by_sheet = BTreeMap::new();
    let mut total_rows = 0;
    for name in &spec_order {
        let spec = &specs[name];
        let set_headers = disambiguate(&spec.set_names);
        let col_headers = disambiguate(&spec.headers);
        let mut header = vec!["工作簿".to_owned(), "工作表".to_owned()];
        header.extend(set_headers);
        header.extend(col_headers);
        header.push("行号".into());
        headers_by_sheet.insert(name.clone(), header);
        data.insert(name.clone(), Vec::new());
    }
    let mut files_hit = 0usize;
    for source in sources {
        if !supported(source) || !source.is_file() {
            continue;
        }
        let dump = read_workbook(source)
            .with_context(|| format!("read comment-summary source {}", source.display()))?;
        files_hit += 1;
        let book_name = source
            .file_stem()
            .map(|s| s.to_string_lossy().into_owned())
            .unwrap_or_default();
        for sheet in &dump.sheets {
            let spec_key = if specs.contains_key(&sheet.name) {
                Some(sheet.name.as_str())
            } else if specs.contains_key("模板") {
                Some("模板")
            } else {
                None
            };
            let Some(spec_key) = spec_key else { continue };
            let spec = &specs[spec_key];
            let values = data
                .get_mut(spec_key)
                .expect("spec output table initialized");
            for (start, end, _, _) in &spec.row_ranges {
                for row in *start..=*end {
                    let mut output = vec![
                        serde_json::Value::String(book_name.clone()),
                        serde_json::Value::String(sheet.name.clone()),
                    ];
                    for address in &spec.set_addresses {
                        output.push(cell_at_or_merged(sheet, address));
                    }
                    for (_, _, start_col, end_col) in &spec.col_ranges {
                        for col in *start_col..=*end_col {
                            output.push(cell_at_or_merged(
                                sheet,
                                &format!("{}{}", column_letters(col), row),
                            ));
                        }
                    }
                    output.push(serde_json::Value::from(row as u64));
                    values.push(output);
                    total_rows += 1;
                }
            }
        }
    }
    if total_rows == 0 {
        return Ok(SummaryResult {
            saved: false,
            rows: 0,
            files: files_hit,
            path: None,
        });
    }
    fs::create_dir_all(output_dir)?;
    let path = unique_output(
        output_dir,
        &format!(
            "{}_按批注汇总（模板+源文件）",
            Local::now().format("%Y%m%d_%H%M%S")
        ),
    )?;
    let mut workbook = Workbook::new();
    let header_format = Format::new().set_bold();
    for name in &spec_order {
        let rows = &data[name];
        let sheet_name = crate::sheet_extract::sanitize_sheet_name(&name);
        let ws = workbook.add_worksheet();
        ws.set_name(&sheet_name)?;
        for (col, header) in headers_by_sheet[name].iter().enumerate() {
            ws.write_string_with_format(0, col as u16, header, &header_format)?;
        }
        for (row, values) in rows.iter().enumerate() {
            for (col, value) in values.iter().enumerate() {
                write_json(ws, (row + 1) as u32, col as u16, value)?;
            }
        }
        if !rows.is_empty() {
            ws.set_freeze_panes(1, 0)?;
        }
    }
    workbook
        .save(&path)
        .with_context(|| format!("save comment summary {}", path.display()))?;
    Ok(SummaryResult {
        saved: true,
        rows: total_rows,
        files: files_hit,
        path: Some(path),
    })
}

fn supported(path: &Path) -> bool {
    path.extension()
        .and_then(|s| s.to_str())
        .is_some_and(|s| matches!(s.to_ascii_lowercase().as_str(), "xlsx" | "xlsm"))
}
fn is_xlsx_or_xlsm(path: &Path) -> bool {
    path.extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| matches!(extension.to_ascii_lowercase().as_str(), "xlsx" | "xlsm"))
}
fn bounds(sheet: &SheetDump) -> (usize, usize, usize, usize) {
    let mut min_row = usize::MAX;
    let mut min_col = usize::MAX;
    let mut max_row = 0;
    let mut max_col = 0;
    for address in sheet.values.keys() {
        if let Some((row, col)) = parse_a1(address) {
            min_row = min_row.min(row);
            min_col = min_col.min(col);
            max_row = max_row.max(row);
            max_col = max_col.max(col);
        }
    }
    if max_row == 0 {
        (0, 0, 0, 0)
    } else {
        (min_row, min_col, max_row, max_col)
    }
}

fn summary_bounds(
    sheet: &SheetDump,
    extent: Option<&WorksheetExtent>,
) -> (usize, usize, usize, usize) {
    if let Some(extent) = extent {
        if extent.max_row == 0 || extent.max_column == 0 {
            return (1, 1, 1, 1);
        }
        return (
            extent.min_row.max(1),
            extent.min_column.max(1),
            extent.max_row,
            extent.max_column,
        );
    }
    let value_bounds = bounds(sheet);
    if value_bounds.0 == 0 {
        (1, 1, 1, 1)
    } else {
        value_bounds
    }
}
fn parse_a1(address: &str) -> Option<(usize, usize)> {
    let index = address.find(|c: char| c.is_ascii_digit())?;
    let col = address[..index].chars().try_fold(0usize, |v, c| {
        v.checked_mul(26)?
            .checked_add(c.to_ascii_uppercase() as usize - 'A' as usize + 1)
    })?;
    Some((address[index..].parse().ok()?, col))
}
fn column_letters(mut col: usize) -> String {
    let mut out = String::new();
    while col > 0 {
        out.insert(0, (b'A' + ((col - 1) % 26) as u8) as char);
        col = (col - 1) / 26;
    }
    out
}
fn write_probe_value(
    ws: &mut Worksheet,
    row: u32,
    col: u16,
    value: &serde_json::Value,
) -> Result<()> {
    let kind = value
        .get("type")
        .and_then(serde_json::Value::as_str)
        .unwrap_or("");
    let v = value.get("value").unwrap_or(&serde_json::Value::Null);
    match kind {
        "string" | "error" => {
            ws.write_string(row, col, v.as_str().unwrap_or_default())?;
        }
        "duration_iso" => {
            let text = v.as_str().unwrap_or_default();
            if let Some(serial) = iso_duration_serial(text) {
                let format = Format::new().set_num_format("[h]:mm:ss");
                ws.write_number_with_format(row, col, serial, &format)?;
            } else {
                ws.write_string(row, col, text)?;
            }
        }
        "integer" => {
            ws.write_number(row, col, v.as_i64().unwrap_or_default() as f64)?;
        }
        "datetime_iso" => {
            let text = v.as_str().unwrap_or_default();
            if let Some(serial) = iso_datetime_serial(text) {
                let format = date_format(serial);
                ws.write_number_with_format(row, col, serial, &format)?;
            } else {
                ws.write_string(row, col, text)?;
            }
        }
        "number" => {
            if let Some(n) = v.as_f64() {
                ws.write_number(row, col, n)?;
            }
        }
        "datetime_serial" => {
            if let Some(n) = v.as_f64() {
                let format = date_format(n);
                ws.write_number_with_format(row, col, n, &format)?;
            }
        }
        "boolean" => {
            ws.write_boolean(row, col, v.as_bool().unwrap_or(false))?;
        }
        _ => {}
    }
    Ok(())
}
fn write_json(ws: &mut Worksheet, row: u32, col: u16, value: &serde_json::Value) -> Result<()> {
    if value
        .get("type")
        .and_then(serde_json::Value::as_str)
        .is_some()
    {
        write_probe_value(ws, row, col, value)?;
    } else if value.is_string() {
        ws.write_string(row, col, value.as_str().unwrap_or_default())?;
    } else if value.is_boolean() {
        ws.write_boolean(row, col, value.as_bool().unwrap_or(false))?;
    } else if let Some(number) = value.as_f64() {
        ws.write_number(row, col, number)?;
    }
    Ok(())
}
fn unique_output(dir: &Path, stem: &str) -> Result<PathBuf> {
    for i in 0..10000 {
        let suffix = if i == 0 {
            String::new()
        } else {
            format!("_{i}")
        };
        let path = dir.join(format!("{stem}{suffix}.xlsx"));
        if !path.exists() {
            return Ok(path);
        }
    }
    anyhow::bail!("cannot reserve unique output under {}", dir.display())
}
fn parse_comment_ranges(
    comments: &BTreeMap<String, String>,
    keyword: &str,
) -> Vec<(usize, usize, usize, usize)> {
    let mut ids = BTreeSet::new();
    for (_, text) in ordered_comments(comments) {
        if text.contains(keyword) {
            if let Some(id) = first_number(text) {
                ids.insert(id);
            }
        }
    }
    let mut out = Vec::new();
    for id in ids {
        let ordered = ordered_comments(comments);
        let start = ordered
            .iter()
            .find(|(_, text)| {
                text.replace(' ', "").contains(&format!("{keyword}{id}"))
                    && !text.replace(' ', "").contains(&format!("{keyword}#{id}"))
            })
            .map(|(a, _)| *a);
        let end = ordered
            .iter()
            .find(|(_, text)| text.replace(' ', "").contains(&format!("{keyword}#{id}")))
            .map(|(a, _)| *a);
        if let (Some(a), Some(b)) = (start, end) {
            if let (Some((r1, c1)), Some((r2, c2))) = (parse_a1(a), parse_a1(b)) {
                out.push((r1.min(r2), r1.max(r2), c1.min(c2), c1.max(c2)));
            }
        }
    }
    out
}
fn ordered_comments<'a>(comments: &'a BTreeMap<String, String>) -> Vec<(&'a String, &'a String)> {
    let mut ordered = comments.iter().collect::<Vec<_>>();
    ordered.sort_by_key(|(address, _)| parse_a1(address).unwrap_or((usize::MAX, usize::MAX)));
    ordered
}

fn first_number(text: &str) -> Option<usize> {
    let normalized = text.replace('#', "");
    let start = normalized.find(|character: char| character.is_ascii_digit())?;
    normalized[start..]
        .chars()
        .take_while(char::is_ascii_digit)
        .collect::<String>()
        .parse()
        .ok()
}
fn parse_set_name(text: &str) -> Option<String> {
    let lower = text.to_lowercase();
    let ascii = lower.find("set(").map(|idx| (idx, 4));
    let fullwidth = lower.find("set（").map(|idx| (idx, "set（".len()));
    let (idx, marker_len) = match (ascii, fullwidth) {
        (Some(left), Some(right)) => {
            if left.0 <= right.0 {
                left
            } else {
                right
            }
        }
        (Some(value), None) | (None, Some(value)) => value,
        (None, None) => return None,
    };
    let rest = &text[idx + marker_len..];
    let end = rest.find([')', '）'])?;
    Some(rest[..end].trim().to_owned())
}
fn build_column_headers(sheet: &SheetDump, ranges: &[(usize, usize, usize, usize)]) -> Vec<String> {
    let mut out = Vec::new();
    for (sr, er, sc, ec) in ranges {
        for col in *sc..=*ec {
            let upper = cell_at_or_merged(sheet, &format!("{}{}", column_letters(col), sr));
            let lower = cell_at_or_merged(sheet, &format!("{}{}", column_letters(col), er));
            let a = probe_text(&probe_to_value(&upper));
            let b = probe_text(&probe_to_value(&lower));
            let text = if *sr < *er && !a.is_empty() && !b.is_empty() && a != b {
                format!("{a}_{b}")
            } else if !a.is_empty() {
                a
            } else if !b.is_empty() {
                b
            } else {
                column_letters(col)
            };
            out.push(text);
        }
    }
    out
}
fn cell_at_or_merged(sheet: &SheetDump, address: &str) -> serde_json::Value {
    if let Some(value) = sheet.values.get(address) {
        return value.clone();
    }
    let Some((row, col)) = parse_a1(address) else {
        return serde_json::Value::Null;
    };
    for range in &sheet.merges {
        let Some((start, end)) = range.split_once(':') else {
            continue;
        };
        if let (Some((r1, c1)), Some((r2, c2))) = (parse_a1(start), parse_a1(end)) {
            if r1.min(r2) <= row && row <= r1.max(r2) && c1.min(c2) <= col && col <= c1.max(c2) {
                return sheet
                    .values
                    .get(&format!("{}{}", column_letters(c1.min(c2)), r1.min(r2)))
                    .cloned()
                    .unwrap_or(serde_json::Value::Null);
            }
        }
    }
    serde_json::Value::Null
}

fn date_format(serial: f64) -> Format {
    if serial.fract().abs() < f64::EPSILON {
        Format::new().set_num_format("yyyy-mm-dd")
    } else {
        Format::new().set_num_format("yyyy-mm-dd hh:mm:ss")
    }
}

fn iso_datetime_serial(text: &str) -> Option<f64> {
    use chrono::{DateTime, NaiveDate, NaiveDateTime};

    let datetime = DateTime::parse_from_rfc3339(text)
        .map(|datetime| datetime.naive_utc())
        .or_else(|_| NaiveDateTime::parse_from_str(text, "%Y-%m-%dT%H:%M:%S%.f"))
        .or_else(|_| NaiveDateTime::parse_from_str(text, "%Y-%m-%d %H:%M:%S%.f"))
        .or_else(|_| {
            NaiveDate::parse_from_str(text, "%Y-%m-%d")
                .map(|date| date.and_hms_opt(0, 0, 0).expect("valid midnight"))
        })
        .ok()?;
    let epoch = NaiveDate::from_ymd_opt(1899, 12, 30)?.and_hms_opt(0, 0, 0)?;
    let millis = datetime.signed_duration_since(epoch).num_milliseconds();
    Some(millis as f64 / 86_400_000.0)
}

fn iso_duration_serial(text: &str) -> Option<f64> {
    let text = text.strip_prefix('P')?;
    let mut in_time = false;
    let mut number = String::new();
    let mut seconds = 0.0;
    for character in text.chars() {
        if character == 'T' {
            in_time = true;
            continue;
        }
        if character.is_ascii_digit() || character == '.' || character == '-' {
            number.push(character);
            continue;
        }
        let value = number.parse::<f64>().ok()?;
        number.clear();
        let multiplier = match (in_time, character) {
            (false, 'D') => 86_400.0,
            (true, 'H') => 3_600.0,
            (true, 'M') => 60.0,
            (true, 'S') => 1.0,
            _ => return None,
        };
        seconds += value * multiplier;
    }
    if !number.is_empty() {
        return None;
    }
    Some(seconds / 86_400.0)
}
fn probe_to_value(value: &serde_json::Value) -> serde_json::Value {
    match value.get("type").and_then(serde_json::Value::as_str) {
        Some("string") | Some("datetime_iso") | Some("duration_iso") | Some("error") => {
            serde_json::Value::String(
                value
                    .get("value")
                    .and_then(serde_json::Value::as_str)
                    .unwrap_or_default()
                    .to_owned(),
            )
        }
        Some("boolean") => value.get("value").cloned().unwrap_or_default(),
        Some("number") | Some("integer") | Some("datetime_serial") => {
            value.get("value").cloned().unwrap_or_default()
        }
        _ => serde_json::Value::Null,
    }
}
fn probe_text(value: &serde_json::Value) -> String {
    if let Some(s) = value.as_str() {
        s.trim().to_owned()
    } else {
        value.to_string()
    }
}
fn disambiguate(names: &[String]) -> Vec<String> {
    let mut counts = std::collections::HashMap::new();
    names
        .iter()
        .map(|name| {
            let count = counts.entry(name.clone()).or_insert(0usize);
            *count += 1;
            if *count == 1 {
                name.clone()
            } else {
                format!("{name}_{}", count)
            }
        })
        .collect()
}
