//! CSV import and normalization for the Audit S3 central-statistics workflow.
//!
//! This module currently covers the legacy CSV path only. XLS/XLSX dispatch is
//! intentionally left to a later migration step that can use the shared
//! `excel-reader` adapter and its own workbook parity cases.

use anyhow::{bail, Context, Result};
use chrono::{NaiveDate, NaiveDateTime};
use encoding_rs::GBK;
use serde::{Deserialize, Serialize};
use std::collections::{BTreeMap, BTreeSet};

pub const EXPECTED_HEADERS: [&str; 14] = [
    "业务类",
    "数据日期",
    "机构类代码",
    "机构类名称",
    "地区代码",
    "地区名称",
    "指标顺序码",
    "指标代码",
    "指标名称",
    "数据属性",
    "币种",
    "频度",
    "批次",
    "数据值",
];

/// Value semantics from the Python `CentralValue` union: number, original
/// abnormal text, or an empty value.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(untagged)]
pub enum CentralValue {
    Number(f64),
    Text(String),
    Empty,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct CentralRecord {
    pub biz_class: String,
    pub record_date: String,
    pub org_code: String,
    pub org_name: String,
    pub region_code: String,
    pub region_name: String,
    pub order_code: String,
    pub indicator: String,
    pub indicator_name: String,
    pub data_attr: String,
    pub currency: String,
    pub frequency: String,
    pub batch: String,
    pub value: CentralValue,
}

impl CentralRecord {
    pub fn key(&self) -> String {
        format!(
            "{}{}{}{}{}{}{}{}",
            self.biz_class,
            self.org_code,
            self.region_code,
            self.indicator,
            self.data_attr,
            self.currency,
            self.frequency,
            self.batch
        )
    }

    pub fn attr_key(&self) -> String {
        format!(
            "{}{}{}{}{}",
            self.indicator, self.data_attr, self.currency, self.frequency, self.batch
        )
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct ImportIssue {
    pub row: usize,
    pub message: String,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct CentralDataset {
    pub records: Vec<CentralRecord>,
    pub issues: Vec<ImportIssue>,
    pub record_date: String,
    pub key_index: BTreeMap<String, CentralRecord>,
    pub freq_dates: BTreeMap<String, String>,
}

impl Default for CentralDataset {
    fn default() -> Self {
        Self {
            records: Vec::new(),
            issues: Vec::new(),
            record_date: String::new(),
            key_index: BTreeMap::new(),
            freq_dates: BTreeMap::new(),
        }
    }
}

impl CentralDataset {
    fn build_index(&mut self) {
        self.key_index.clear();
        for record in &self.records {
            // Match the source dictionary behavior: a later duplicate key wins.
            self.key_index.insert(record.key(), record.clone());
        }
    }
}

/// Parse one CSV input using the original encoding order (strict GBK, then
/// UTF-8 with optional BOM, then GBK replacement decoding) and build the S3
/// import dataset. `unit_factor` and exempt indicator codes retain the legacy
/// float conversion behavior for this migration slice.
pub fn parse_central_csv(
    file_name: &str,
    bytes: &[u8],
    unit_factor: f64,
    exempt_indicators: &BTreeSet<String>,
) -> Result<CentralDataset> {
    let text = decode_legacy_csv(bytes);
    // Python's csv.reader retains an empty physical record, while the Rust csv
    // crate skips it. Turn only empty records outside quoted fields into an
    // explicit empty CSV field so subsequent issue row numbers remain stable.
    let text = preserve_empty_records(&text);
    let mut reader = csv::ReaderBuilder::new()
        .has_headers(false)
        .flexible(true)
        .from_reader(text.as_bytes());
    let mut rows = Vec::new();
    for record in reader.records() {
        let record = record.with_context(|| format!("CSV 格式错误：{file_name}"))?;
        rows.push(record.iter().map(str::to_owned).collect::<Vec<_>>());
    }
    build_dataset(&rows, file_name, unit_factor, exempt_indicators)
}

fn preserve_empty_records(text: &str) -> String {
    let bytes = text.as_bytes();
    let mut output = Vec::with_capacity(bytes.len());
    let mut line_start = 0;
    let mut cursor = 0;
    let mut in_quotes = false;
    let mut field_start = true;

    while cursor < bytes.len() {
        if bytes[cursor] != b'\r' && bytes[cursor] != b'\n' {
            if bytes[cursor] == b'"' {
                if in_quotes && bytes.get(cursor + 1) == Some(&b'"') {
                    cursor += 2;
                    continue;
                }
                if in_quotes {
                    in_quotes = false;
                } else if field_start {
                    in_quotes = true;
                    field_start = false;
                }
            } else if !in_quotes {
                field_start = bytes[cursor] == b',';
            }
            cursor += 1;
            continue;
        }

        if cursor == line_start && !in_quotes {
            output.extend_from_slice(b"\"\"");
        }
        output.extend_from_slice(&bytes[line_start..cursor]);
        if bytes[cursor] == b'\r' && bytes.get(cursor + 1) == Some(&b'\n') {
            output.extend_from_slice(b"\r\n");
            cursor += 2;
        } else {
            output.push(bytes[cursor]);
            cursor += 1;
        }
        line_start = cursor;
        if !in_quotes {
            field_start = true;
        }
    }

    output.extend_from_slice(&bytes[line_start..]);
    // The byte operations above preserve valid UTF-8 boundaries because all
    // inserted bytes are ASCII and source bytes are copied unchanged.
    String::from_utf8(output).expect("CSV newline compatibility keeps UTF-8 valid")
}

fn decode_legacy_csv(bytes: &[u8]) -> String {
    let (gbk, had_gbk_errors) = GBK.decode_without_bom_handling(bytes);
    if !had_gbk_errors {
        return gbk.into_owned();
    }

    let utf8 = bytes.strip_prefix(&[0xEF, 0xBB, 0xBF]).unwrap_or(bytes);
    if let Ok(text) = std::str::from_utf8(utf8) {
        return text.to_owned();
    }

    let (replacement, _, _) = GBK.decode(bytes);
    replacement.into_owned()
}

fn build_dataset(
    rows: &[Vec<String>],
    file_name: &str,
    unit_factor: f64,
    exempt_indicators: &BTreeSet<String>,
) -> Result<CentralDataset> {
    let Some(first_row) = rows.first() else {
        bail!("导数文件为空：{file_name}");
    };
    if !header_matches(first_row) {
        bail!("表头与大集中导数 14 列不符：{file_name}");
    }

    let mut dataset = CentralDataset::default();
    for (offset, input_fields) in rows.iter().enumerate().skip(1) {
        if !input_fields.iter().any(|cell| !cell.trim().is_empty()) {
            continue;
        }
        let mut fields = input_fields.clone();
        fields.resize(fields.len().max(EXPECTED_HEADERS.len()), String::new());

        let record_date = parse_central_date(&fields[1]);
        if record_date.is_empty() {
            dataset.issues.push(ImportIssue {
                row: offset + 1,
                message: format!("数据日期无法解析：{}", python_repr(&fields[1])),
            });
        }

        let raw_value = fields[13].trim();
        let indicator = clean_code(&fields[7]);
        let value = if raw_value.is_empty() {
            CentralValue::Empty
        } else if let Some(number) = parse_number(raw_value) {
            if number == 0.0 && is_abnormal_number(raw_value) {
                dataset.issues.push(ImportIssue {
                    row: offset + 1,
                    message: format!("数据值非纯数值：{}", python_repr(raw_value)),
                });
                CentralValue::Text(raw_value.to_owned())
            } else if exempt_indicators.contains(&indicator) {
                CentralValue::Number(number)
            } else {
                CentralValue::Number(number * unit_factor)
            }
        } else {
            dataset.issues.push(ImportIssue {
                row: offset + 1,
                message: format!("数据值非纯数值：{}", python_repr(raw_value)),
            });
            CentralValue::Text(raw_value.to_owned())
        };

        dataset.records.push(CentralRecord {
            biz_class: fields[0].trim().to_owned(),
            record_date,
            org_code: clean_code(&fields[2]),
            org_name: fields[3].trim().to_owned(),
            region_code: clean_code(&fields[4]),
            region_name: fields[5].trim().to_owned(),
            order_code: clean_code(&fields[6]),
            indicator,
            indicator_name: fields[8].trim().to_owned(),
            data_attr: fields[9].trim().to_owned(),
            currency: fields[10].trim().to_owned(),
            frequency: fields[11].trim().to_owned(),
            batch: fields[12].trim().to_owned(),
            value,
        });
    }

    let dates: BTreeSet<_> = dataset
        .records
        .iter()
        .filter(|record| !record.record_date.is_empty())
        .map(|record| record.record_date.clone())
        .collect();
    if dates.len() > 1 {
        let dates = dates
            .iter()
            .map(|date| python_repr(date))
            .collect::<Vec<_>>()
            .join(", ");
        bail!("{file_name}：数据日期不一致（[{dates}]）");
    }
    dataset.record_date = dates.into_iter().next().unwrap_or_default();
    dataset.build_index();
    Ok(dataset)
}

fn header_matches(row: &[String]) -> bool {
    row.len() >= EXPECTED_HEADERS.len()
        && row
            .iter()
            .take(EXPECTED_HEADERS.len())
            .map(|cell| cell.trim())
            .eq(EXPECTED_HEADERS)
}

fn clean_code(value: &str) -> String {
    value.replace('\'', "").trim().to_owned()
}

fn parse_central_date(value: &str) -> String {
    let text = value.trim().trim_start_matches('\'');
    let date = NaiveDate::parse_from_str(text, "%Y-%m-%d")
        .or_else(|_| NaiveDate::parse_from_str(text, "%Y/%m/%d"))
        .or_else(|_| NaiveDate::parse_from_str(text, "%Y%m%d"))
        .map(|date| date.format("%Y-%m-%d").to_string())
        .or_else(|_| {
            NaiveDateTime::parse_from_str(text, "%Y-%m-%d %H:%M:%S")
                .map(|moment| moment.date().format("%Y-%m-%d").to_string())
        });
    date.unwrap_or_default()
}

/// Match Python's `parse_number`: try the longest prefix accepted by float().
fn parse_number(value: &str) -> Option<f64> {
    let mut ends: Vec<_> = value
        .char_indices()
        .map(|(index, character)| index + character.len_utf8())
        .collect();
    ends.reverse();
    ends.into_iter()
        .find_map(|end| value[..end].trim().parse::<f64>().ok())
}

fn is_abnormal_number(text: &str) -> bool {
    if text.is_empty() {
        return true;
    }
    let Some(number) = parse_number(text) else {
        return true;
    };
    if number == 0.0 {
        let stripped = text.trim_start_matches(['+', '−', '-']);
        return !stripped.starts_with(['0', '.']);
    }
    false
}

fn python_repr(value: &str) -> String {
    let quote = if value.contains('\'') && !value.contains('"') {
        '"'
    } else {
        '\''
    };
    let mut output = String::with_capacity(value.len() + 2);
    output.push(quote);
    for character in value.chars() {
        match character {
            '\\' => output.push_str("\\\\"),
            '\n' => output.push_str("\\n"),
            '\r' => output.push_str("\\r"),
            '\t' => output.push_str("\\t"),
            character if character == quote => {
                output.push('\\');
                output.push(character);
            }
            character => output.push(character),
        }
    }
    output.push(quote);
    output
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde::Deserialize;
    use serde_json::Value;

    #[derive(Deserialize)]
    struct CaseFile {
        cases: Vec<Case>,
    }

    #[derive(Deserialize)]
    struct Case {
        id: String,
        file_name: String,
        encoding: String,
        csv_text: String,
        unit_factor: f64,
        exempt_indicators: BTreeSet<String>,
    }

    fn encode_case(case: &Case) -> Vec<u8> {
        match case.encoding.as_str() {
            "gbk" => {
                let (bytes, _, had_errors) = GBK.encode(&case.csv_text);
                assert!(
                    !had_errors,
                    "fixture {} must be representable in GBK",
                    case.id
                );
                bytes.into_owned()
            }
            "utf-8" => case.csv_text.as_bytes().to_vec(),
            "utf-8-sig" => {
                let mut bytes = vec![0xEF, 0xBB, 0xBF];
                bytes.extend_from_slice(case.csv_text.as_bytes());
                bytes
            }
            encoding => panic!("unsupported test encoding {encoding}"),
        }
    }

    #[test]
    fn central_csv_dataset_matches_python_baseline() {
        let cases: CaseFile =
            serde_json::from_str(include_str!("../tests/fixtures/central_csv_cases.json"))
                .expect("valid CSV parity cases");
        let baseline: Value = serde_json::from_str(include_str!(
            "../probe/reports/central_csv_python_baseline.json"
        ))
        .expect("valid Python CSV baseline");
        let expected_cases = baseline["cases"].as_array().unwrap();
        assert_eq!(cases.cases.len(), expected_cases.len());

        for (case, expected) in cases.cases.iter().zip(expected_cases) {
            assert_eq!(case.id, expected["id"].as_str().unwrap());
            let actual = parse_central_csv(
                &case.file_name,
                &encode_case(case),
                case.unit_factor,
                &case.exempt_indicators,
            )
            .unwrap_or_else(|error| panic!("{}: {error:#}", case.id));
            let actual = serde_json::to_value(actual).unwrap();
            assert_eq!(
                actual, expected["dataset"],
                "Python parity case {}",
                case.id
            );
        }
    }

    #[test]
    fn rejects_empty_file_and_invalid_header() {
        let exempt = BTreeSet::new();
        let empty = parse_central_csv("empty.csv", b"", 1.0, &exempt).unwrap_err();
        assert!(empty.to_string().contains("导数文件为空"));

        let wrong_header = parse_central_csv("wrong.csv", b"a,b,c\n", 1.0, &exempt).unwrap_err();
        assert!(wrong_header.to_string().contains("14 列不符"));
    }
}
