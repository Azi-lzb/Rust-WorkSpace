//! Pure template identity, comment, formula-result, and rule-row parsing logic.
//!
//! These functions are migrated from `base_audit.template`. They accept values
//! already read by the caller; workbook access and audit execution remain outside
//! this module.

use crate::models::AuditRule;
use serde::{Deserialize, Serialize};
use serde_json::Value;
use std::fmt;

pub const REQUIRED_HEADERS: [&str; 8] = [
    "规则编号",
    "启用",
    "报表代码",
    "工作表",
    "公式单元格",
    "定位单元格",
    "级别",
    "问题说明",
];
pub const ALLOWED_SEVERITIES: [&str; 3] = ["错误", "核实", "提示"];

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TemplateError {
    message: String,
}

impl TemplateError {
    fn new(message: impl Into<String>) -> Self {
        Self {
            message: message.into(),
        }
    }
}

impl fmt::Display for TemplateError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.message)
    }
}

impl std::error::Error for TemplateError {}

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct ParsedFormulaResult {
    pub severity: String,
    pub indicator: String,
    pub value: String,
    pub message: String,
    pub comparison_value: String,
    pub reference_value: String,
    pub difference_value: String,
    pub detail: String,
    pub check_type: String,
}

/// Remove leading template marks and changing `20YYMM`/`20YYMMDD` date tokens.
pub fn normalize_template_name(value: &str) -> String {
    let mut text = value.trim().to_owned();
    while text.starts_with('!') || text.starts_with('！') {
        let first = text.chars().next().unwrap();
        text.drain(..first.len_utf8());
    }
    let text = remove_date_tokens(&text);
    text.trim_end_matches(|character| matches!(character, ' ' | '_' | '.' | '-'))
        .to_owned()
}

/// Parse and validate the first explicit rule ID in a comment.
pub fn parse_comment_rule_id(comment_text: &str) -> Result<String, TemplateError> {
    let chars: Vec<char> = comment_text.chars().collect();
    for start in 0..chars.len() {
        if let Some((rule_id, _end)) = rule_id_match(&chars, start) {
            if valid_rule_id(&rule_id) {
                return Ok(rule_id);
            }
            return Err(TemplateError::new(format!("规则编号无效：{rule_id}")));
        }
    }
    Ok(String::new())
}

/// Remove comment prefixes, rule IDs, and template region markers.
pub fn clean_rule_comment(comment_text: &str) -> String {
    let text = strip_leading_comment_source(comment_text);
    let text = remove_rule_id_markers(&text);
    let text = remove_region_markers(&text);
    text.lines()
        .map(|line| line.trim_matches(|character| " ：:，,".contains(character)))
        .filter(|line| !line.trim().is_empty())
        .collect::<Vec<_>>()
        .join("\n")
}

/// Parse the positional fields emitted by a template formula result.
pub fn parse_formula_result(
    value: &Value,
    default_severity: Option<&str>,
    default_message: Option<&str>,
) -> ParsedFormulaResult {
    let text = py_text(value).replace('｜', "|");
    let parts = text.split('|').map(str::trim).collect::<Vec<_>>();
    let marker = parts.first().copied().unwrap_or_default();
    let severity = if ["错误", "硬性", "错"]
        .iter()
        .any(|word| marker.contains(word))
    {
        "错误"
    } else if marker.contains("软性") {
        "核实"
    } else {
        default_severity.unwrap_or("错误")
    };
    let slot = |index: usize| parts.get(index).copied().unwrap_or_default().to_owned();
    let detail = slot(3);
    let difference_value = if parts.len() > 6 {
        parts[6..].join("|").trim().to_owned()
    } else {
        String::new()
    };
    ParsedFormulaResult {
        severity: severity.to_owned(),
        check_type: slot(1),
        indicator: slot(2),
        message: if detail.is_empty() {
            default_message.unwrap_or_default().to_owned()
        } else {
            detail.clone()
        },
        value: slot(4),
        comparison_value: slot(5),
        reference_value: String::new(),
        difference_value,
        detail,
    }
}

/// Convert workbook rule records into normalized AuditRule values.
///
/// This is a value-level parser; the caller is responsible for obtaining rows
/// from the template using an accepted workbook reader.
pub fn parse_rule_rows(rows: &[Value]) -> Result<Vec<AuditRule>, TemplateError> {
    let mut rules = Vec::new();
    let mut seen = std::collections::BTreeSet::new();
    let mut errors = Vec::new();

    for (index, row) in rows.iter().enumerate() {
        let row_number = index + 2;
        let rule_id = row_text(row, "规则编号");
        if rule_id.is_empty() {
            continue;
        }
        if !seen.insert(rule_id.clone()) {
            errors.push(format!("第{row_number}行规则编号重复：{rule_id}"));
            continue;
        }

        let enabled_text = row_text(row, "启用");
        let enabled = !["否", "0", "False", "false", "停用"].contains(&enabled_text.as_str());
        let sheet_name = row_text(row, "工作表");
        let formula_cell = row_text(row, "公式单元格").replace('$', "").to_uppercase();
        let target_cell = row_text(row, "定位单元格").replace('$', "").to_uppercase();
        let severity = {
            let value = row_text(row, "级别");
            if value.is_empty() {
                "错误".to_owned()
            } else {
                value
            }
        };

        if sheet_name.is_empty() {
            errors.push(format!("第{row_number}行未填写工作表：{rule_id}"));
        }
        if !valid_cell_reference(&formula_cell) {
            errors.push(format!(
                "第{row_number}行公式单元格无效：{}",
                if formula_cell.is_empty() {
                    "空"
                } else {
                    &formula_cell
                }
            ));
        }
        if !valid_cell_reference(&target_cell) {
            errors.push(format!(
                "第{row_number}行定位单元格无效：{}",
                if target_cell.is_empty() {
                    "空"
                } else {
                    &target_cell
                }
            ));
        }
        if !ALLOWED_SEVERITIES.contains(&severity.as_str()) {
            errors.push(format!("第{row_number}行级别无效：{severity}"));
        }

        rules.push(AuditRule {
            rule_id,
            enabled,
            report_code: row_text(row, "报表代码"),
            sheet_name,
            formula_cell: formula_cell.clone(),
            target_cell,
            severity,
            message: row_text(row, "问题说明"),
            copy_range: formula_cell,
            result_mode: "nonempty".to_owned(),
            value_from_result: false,
        });
    }

    if !errors.is_empty() {
        return Err(TemplateError::new(errors.join("；")));
    }
    if rules.is_empty() {
        return Err(TemplateError::new("审核规则表中没有可读取的规则"));
    }
    Ok(rules)
}

fn row_text(row: &Value, key: &str) -> String {
    row.get(key).map(py_text).unwrap_or_default()
}

fn py_text(value: &Value) -> String {
    match value {
        Value::Null => String::new(),
        Value::String(text) => text.trim().to_owned(),
        Value::Bool(true) => "True".to_owned(),
        Value::Bool(false) => "False".to_owned(),
        Value::Number(number) => number.to_string().trim().to_owned(),
        Value::Array(_) | Value::Object(_) => value.to_string().trim().to_owned(),
    }
}

fn valid_cell_reference(value: &str) -> bool {
    let chars: Vec<char> = value.chars().collect();
    let mut index = usize::from(chars.first() == Some(&'$'));
    let letters_start = index;
    while chars
        .get(index)
        .is_some_and(|character| character.is_ascii_alphabetic())
    {
        index += 1;
    }
    let letter_count = index - letters_start;
    if !(1..=3).contains(&letter_count) {
        return false;
    }
    if chars.get(index) == Some(&'$') {
        index += 1;
    }
    if !chars
        .get(index)
        .is_some_and(|character| ('1'..='9').contains(character))
    {
        return false;
    }
    index += 1;
    while chars
        .get(index)
        .is_some_and(|character| character.is_ascii_digit())
    {
        index += 1;
    }
    index == chars.len()
}

fn valid_rule_id(value: &str) -> bool {
    let length = value.chars().count();
    (2..=64).contains(&length)
        && value
            .chars()
            .all(|character| character.is_alphanumeric() || matches!(character, '_' | '.' | '-'))
}

fn rule_id_match(chars: &[char], start: usize) -> Option<(String, usize)> {
    let token_len = if chars.get(start..start + 4) == Some(&['规', '则', '编', '号']) {
        4
    } else if chars.get(start) == Some(&'规')
        && chars.get(start + 1) == Some(&'则')
        && chars
            .get(start + 2)
            .is_some_and(|character| character.eq_ignore_ascii_case(&'i'))
        && chars
            .get(start + 3)
            .is_some_and(|character| character.eq_ignore_ascii_case(&'d'))
    {
        4
    } else {
        return None;
    };
    let mut index = start + token_len;
    while chars
        .get(index)
        .is_some_and(|character| character.is_whitespace())
    {
        index += 1;
    }
    if !matches!(chars.get(index), Some(':' | '：')) {
        return None;
    }
    index += 1;
    while chars
        .get(index)
        .is_some_and(|character| character.is_whitespace())
    {
        index += 1;
    }
    let value_start = index;
    while chars.get(index).is_some_and(|character| {
        !character.is_whitespace() && !matches!(character, '，' | ',' | '；' | ';')
    }) {
        index += 1;
    }
    (index > value_start).then(|| (chars[value_start..index].iter().collect(), index))
}

fn strip_leading_comment_source(text: &str) -> String {
    let chars: Vec<char> = text.chars().collect();
    let first_line_end = chars
        .iter()
        .position(|character| *character == '\n')
        .unwrap_or(chars.len());
    let colon = (1..=30.min(first_line_end))
        .rev()
        .find(|index| chars.get(*index) == Some(&':'));
    let mut start = colon.map(|index| index + 1).unwrap_or(0);
    while chars
        .get(start)
        .is_some_and(|character| character.is_whitespace())
    {
        start += 1;
    }
    chars[start..].iter().collect()
}

fn remove_rule_id_markers(text: &str) -> String {
    let chars: Vec<char> = text.chars().collect();
    let mut output = String::new();
    let mut index = 0;
    while index < chars.len() {
        if let Some((_rule_id, end)) = rule_id_match(&chars, index) {
            index = end;
        } else {
            output.push(chars[index]);
            index += 1;
        }
    }
    output
}

fn remove_region_markers(text: &str) -> String {
    const MARKERS: [&str; 2] = ["本地校验区域", "联表校验区域"];
    let chars: Vec<char> = text.chars().collect();
    let mut output = String::new();
    let mut index = 0;
    while index < chars.len() {
        let marker = MARKERS.iter().find(|marker| {
            let marker_chars = marker.chars().collect::<Vec<_>>();
            chars.get(index..index + marker_chars.len()) == Some(marker_chars.as_slice())
        });
        if let Some(marker) = marker {
            let mut end = index + marker.chars().count();
            if chars.get(end) == Some(&'#') {
                end += 1;
            }
            let digits_start = end;
            while chars
                .get(end)
                .is_some_and(|character| is_python_decimal_digit(*character))
            {
                end += 1;
            }
            if end > digits_start {
                index = end;
                continue;
            }
        }
        output.push(chars[index]);
        index += 1;
    }
    output
}

// Python's `\d` matches Unicode Decimal_Number, not every Unicode numeric
// character. This is the Unicode 15.0 Nd range table used by the Python 3.12
// parity baseline; keeping it explicit makes the Rust behavior deterministic.
const UNICODE_DECIMAL_RANGES: &[(u32, u32)] = &[
    (0x0660, 0x0669),
    (0x06F0, 0x06F9),
    (0x07C0, 0x07C9),
    (0x0966, 0x096F),
    (0x09E6, 0x09EF),
    (0x0A66, 0x0A6F),
    (0x0AE6, 0x0AEF),
    (0x0B66, 0x0B6F),
    (0x0BE6, 0x0BEF),
    (0x0C66, 0x0C6F),
    (0x0CE6, 0x0CEF),
    (0x0D66, 0x0D6F),
    (0x0DE6, 0x0DEF),
    (0x0E50, 0x0E59),
    (0x0ED0, 0x0ED9),
    (0x0F20, 0x0F29),
    (0x1040, 0x1049),
    (0x1090, 0x1099),
    (0x17E0, 0x17E9),
    (0x1810, 0x1819),
    (0x1946, 0x194F),
    (0x19D0, 0x19D9),
    (0x1A80, 0x1A89),
    (0x1A90, 0x1A99),
    (0x1B50, 0x1B59),
    (0x1BB0, 0x1BB9),
    (0x1C40, 0x1C49),
    (0x1C50, 0x1C59),
    (0xA620, 0xA629),
    (0xA8D0, 0xA8D9),
    (0xA900, 0xA909),
    (0xA9D0, 0xA9D9),
    (0xA9F0, 0xA9F9),
    (0xAA50, 0xAA59),
    (0xABF0, 0xABF9),
    (0xFF10, 0xFF19),
    (0x104A0, 0x104A9),
    (0x10D30, 0x10D39),
    (0x11066, 0x1106F),
    (0x110F0, 0x110F9),
    (0x11136, 0x1113F),
    (0x111D0, 0x111D9),
    (0x112F0, 0x112F9),
    (0x11450, 0x11459),
    (0x114D0, 0x114D9),
    (0x11650, 0x11659),
    (0x116C0, 0x116C9),
    (0x11730, 0x11739),
    (0x118E0, 0x118E9),
    (0x11950, 0x11959),
    (0x11C50, 0x11C59),
    (0x11D50, 0x11D59),
    (0x11DA0, 0x11DA9),
    (0x11F50, 0x11F59),
    (0x16A60, 0x16A69),
    (0x16AC0, 0x16AC9),
    (0x16B50, 0x16B59),
    (0x1D7CE, 0x1D7FF),
    (0x1E140, 0x1E149),
    (0x1E2F0, 0x1E2F9),
    (0x1E4F0, 0x1E4F9),
    (0x1E950, 0x1E959),
    (0x1FBF0, 0x1FBF9),
];

fn is_python_decimal_digit(character: char) -> bool {
    if character.is_ascii_digit() {
        return true;
    }
    let codepoint = character as u32;
    UNICODE_DECIMAL_RANGES
        .iter()
        .any(|(start, end)| (*start..=*end).contains(&codepoint))
}

fn remove_date_tokens(value: &str) -> String {
    let chars: Vec<char> = value.chars().collect();
    let mut output = String::new();
    let mut index = 0;
    while index < chars.len() {
        if chars.get(index) == Some(&'2')
            && chars.get(index + 1) == Some(&'0')
            && chars
                .get(index + 2)
                .is_some_and(|character| is_python_decimal_digit(*character))
            && chars
                .get(index + 3)
                .is_some_and(|character| is_python_decimal_digit(*character))
        {
            let mut end = index + 4;
            let mut groups = 0;
            while groups < 2 {
                let mut candidate_end = end;
                if chars
                    .get(candidate_end)
                    .is_some_and(|character| matches!(character, '-' | '.' | '_'))
                {
                    candidate_end += 1;
                }
                if chars
                    .get(candidate_end)
                    .is_some_and(|character| is_python_decimal_digit(*character))
                    && chars
                        .get(candidate_end + 1)
                        .is_some_and(|character| is_python_decimal_digit(*character))
                {
                    end = candidate_end + 2;
                    groups += 1;
                } else {
                    break;
                }
            }
            if groups > 0 {
                index = end;
                continue;
            }
        }
        output.push(chars[index]);
        index += 1;
    }
    output
}
