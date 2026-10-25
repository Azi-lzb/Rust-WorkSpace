//! S5 configuration validation entry (`validate_s5_config` port, config level).
//!
//! Faithful port of the source project's
//! `s5_configured_statistics.service.validate_s5_config` for the
//! configuration-level checks: scope handling, business inference from a
//! matched template, configuration loading with error mapping, rule/business
//! counters, enabled-rule reference extraction, and the no-template /
//! no-external-file `not_checked` outcomes.
//!
//! Port boundary: the template/external workbook address scanning (including
//! the `all`/`all_enabled` full-scope branch with its template catalog) is not
//! migrated yet; calls that would enter those branches fail with an explicit
//! not-migrated error instead of silently returning a check result.

use crate::s5_config::{load_configuration, Configuration};
use crate::s5_config_dsl::{calls_in, parse_dictionary, parse_expression, Node};
use crate::s5_scan::{contextual_row_hits, row_path_candidates, scan_rule_paths, ScanBook};
use excel_reader::read_workbook;
use serde_json::{json, Value};
use std::collections::BTreeMap;
use std::path::Path;

const NOT_MIGRATED_SCAN: &str =
    "S5 配置校验的模板/外部地址扫描尚未迁移，请使用不含模板/外部文件的配置级检查";

fn configuration_error_to_json(message: &str) -> Value {
    json!({"message": message, "locations": []})
}

/// Mirror of `_configured_business_sheets`: lightweight business-to-sheet
/// hints read without parsing rule expressions.
fn configured_business_sheets(config_path: &Path) -> BTreeMap<String, BTreeSetText> {
    let mut result: BTreeMap<String, BTreeSetText> = BTreeMap::new();
    let Ok(dump) = read_workbook(config_path) else {
        return result;
    };
    let Some(sheet) = dump.sheets.iter().find(|sheet| sheet.name == "审核规则") else {
        return result;
    };
    let headers = sheet_header_names(sheet);
    let business_header = ["业务名称", "业务类型"]
        .into_iter()
        .find(|name| headers.iter().any(|header| header == name));
    let sheet_header = ["工作表名称", "默认Sheet"]
        .into_iter()
        .find(|name| headers.iter().any(|header| header == name));
    let enabled_header = headers.iter().find(|header| *header == "是否启用");
    if let (Some(business_header), Some(sheet_header)) = (business_header, sheet_header) {
        let business_col = headers
            .iter()
            .position(|header| header == business_header)
            .unwrap();
        let sheet_col = headers
            .iter()
            .position(|header| header == sheet_header)
            .unwrap();
        let enabled_col = enabled_header
            .and_then(|header| headers.iter().position(|candidate| candidate == header));
        for row in 2..=sheet_max_row(sheet) {
            let business = sheet_cell_text(sheet, business_col + 1, row);
            let sheet_name = sheet_cell_text(sheet, sheet_col + 1, row);
            if business.is_empty() || sheet_name.is_empty() {
                continue;
            }
            let enabled = enabled_col
                .map(|column| sheet_cell_text(sheet, column + 1, row).to_lowercase())
                .unwrap_or_else(|| "是".to_owned());
            if matches!(
                enabled.as_str(),
                "是" | "是的" | "yes" | "y" | "true" | "1" | "启用" | "已确认"
            ) {
                result.entry(business).or_default().insert(sheet_name);
            }
        }
    }
    if let Some(metric_sheet) = dump
        .sheets
        .iter()
        .find(|sheet| sheet.name == "模板指标目录")
    {
        let headers = sheet_header_names(metric_sheet);
        if let (Some(business_position), Some(sheet_position)) = (
            headers.iter().position(|header| header == "业务类型"),
            headers.iter().position(|header| header == "工作表名称"),
        ) {
            for row in 2..=sheet_max_row(metric_sheet) {
                let business = sheet_cell_text(metric_sheet, business_position + 1, row);
                let sheet_name = sheet_cell_text(metric_sheet, sheet_position + 1, row);
                if !business.is_empty()
                    && !sheet_name.is_empty()
                    && result.get(&business).is_none_or(|set| set.is_empty())
                {
                    result.entry(business).or_default().insert(sheet_name);
                }
            }
        }
    }
    result
}

type BTreeSetText = std::collections::BTreeSet<String>;

fn sheet_header_names(sheet: &excel_reader::SheetDump) -> Vec<String> {
    let mut max_column = 0usize;
    for address in sheet.values.keys() {
        if let Some((row, column)) = parse_cell(address) {
            if row == 1 && column > max_column {
                max_column = column;
            }
        }
    }
    (1..=max_column)
        .map(|column| {
            sheet
                .values
                .get(&format!("{}1", column_letter(column)))
                .map(cell_text)
                .unwrap_or_default()
        })
        .collect()
}

fn sheet_max_row(sheet: &excel_reader::SheetDump) -> usize {
    let mut max_row = 0usize;
    for address in sheet.values.keys() {
        if let Some((row, _)) = parse_cell(address) {
            if row > max_row {
                max_row = row;
            }
        }
    }
    max_row
}

fn sheet_cell_text(sheet: &excel_reader::SheetDump, column: usize, row: usize) -> String {
    sheet
        .values
        .get(&format!("{}{row}", column_letter(column)))
        .map(cell_text)
        .unwrap_or_default()
}

fn column_letter(mut column: usize) -> String {
    let mut letters = Vec::new();
    while column > 0 {
        let digit = (column - 1) % 26;
        letters.push((b'A' + digit as u8) as char);
        column = (column - 1) / 26;
    }
    letters.into_iter().rev().collect()
}

fn parse_cell(address: &str) -> Option<(usize, usize)> {
    let mut row_part = String::new();
    let mut column = 0usize;
    for character in address.chars() {
        if character.is_ascii_alphabetic() {
            column = column * 26 + (character.to_ascii_uppercase() as usize - 'A' as usize + 1);
        } else if character.is_ascii_digit() {
            row_part.push(character);
        } else {
            return None;
        }
    }
    let row = row_part.parse::<usize>().ok()?;
    (row > 0 && column > 0).then_some((row, column))
}

fn cell_text(value: &serde_json::Value) -> String {
    let kind = value.get("type").and_then(serde_json::Value::as_str);
    let raw = value.get("value");
    match kind {
        Some("string") => raw
            .and_then(serde_json::Value::as_str)
            .unwrap_or_default()
            .trim()
            .to_owned(),
        Some("number") => {
            let number = raw.and_then(serde_json::Value::as_f64).unwrap_or(0.0);
            if number.fract() == 0.0 && number.abs() < 1e15 {
                format!("{}", number as i64)
            } else {
                format!("{number}")
            }
        }
        Some("boolean") => {
            if raw.and_then(serde_json::Value::as_bool) == Some(true) {
                "True".to_owned()
            } else {
                "False".to_owned()
            }
        }
        _ => String::new(),
    }
}

/// Mirror of `_infer_business_from_template`.
fn infer_business_from_template(
    config_path: &Path,
    template_path: &Path,
) -> Result<String, String> {
    let template_path = template_path.to_path_buf();
    if !template_path.is_file() {
        return Err(format!("匹配模板不存在：{}", template_path.display()));
    }
    let sheets = excel_reader::read_xlsx_sheet_names_ooxml(&template_path).map_err(|_| {
        format!(
            "无法读取匹配模板的工作表清单：{}",
            template_path
                .file_name()
                .map(|name| name.to_string_lossy().into_owned())
                .unwrap_or_default()
        )
    })?;
    let sheets: std::collections::BTreeSet<String> = sheets.into_iter().collect();
    if sheets.is_empty() {
        return Err(format!(
            "无法读取匹配模板的工作表清单：{}",
            template_path
                .file_name()
                .map(|name| name.to_string_lossy().into_owned())
                .unwrap_or_default()
        ));
    }
    let configured = configured_business_sheets(config_path);
    let mut scored: Vec<(f64, String)> = Vec::new();
    for (business, required) in &configured {
        if required.is_empty() {
            continue;
        }
        let overlap = required.intersection(&sheets).count();
        if overlap > 0 {
            scored.push((overlap as f64 / required.len() as f64, business.clone()));
        }
    }
    if scored.is_empty() {
        return Err(format!(
            "模板“{}”的工作表无法对应到配置业务，请核实模板和业务工作表配置",
            template_path
                .file_name()
                .map(|name| name.to_string_lossy().into_owned())
                .unwrap_or_default()
        ));
    }
    let best = scored
        .iter()
        .map(|(score, _)| *score)
        .fold(f64::NEG_INFINITY, f64::max);
    let mut candidates: Vec<String> = scored
        .iter()
        .filter(|(score, _)| *score == best)
        .map(|(_, business)| business.clone())
        .collect();
    if candidates.len() > 1 {
        let stem = template_path
            .file_stem()
            .map(|value| value.to_string_lossy().into_owned())
            .unwrap_or_default();
        let name_matches: Vec<String> = candidates
            .iter()
            .filter(|business| stem.contains(business.as_str()))
            .cloned()
            .collect();
        if name_matches.len() == 1 {
            candidates = name_matches;
        }
    }
    if candidates.len() != 1 {
        let mut sorted_candidates = candidates.clone();
        sorted_candidates.sort();
        return Err(format!(
            "模板“{}”同时匹配多个业务（{}），无法唯一确定当前业务",
            template_path
                .file_name()
                .map(|name| name.to_string_lossy().into_owned())
                .unwrap_or_default(),
            sorted_candidates.join("、")
        ));
    }
    Ok(candidates.remove(0))
}

#[derive(Clone, Debug, PartialEq)]
struct TemplateReference {
    rule_id: String,
    revision: i64,
    field: String,
    source: String,
}

#[derive(Clone, Debug, PartialEq)]
struct DynamicReference {
    business: String,
    sheet: String,
}

struct RuleReferences {
    /// First-seen key order mirrors Python's defaultdict iteration.
    template_refs: Vec<(
        (String, String, Vec<String>, Vec<String>),
        Vec<TemplateReference>,
    )>,
    external_reference_count: usize,
    dynamic_refs: Vec<DynamicReference>,
}

fn path_argument(node: &Node) -> Vec<String> {
    match node {
        Node::Literal(crate::s5_config_dsl::LiteralNode {
            value: crate::s5_config_dsl::LiteralValue::String(value),
            kind: "string",
        }) => crate::s5_config::split_at_path(value),
        Node::Tuple(tuple) => tuple
            .values
            .iter()
            .map(|value| value.trim().to_owned())
            .collect(),
        _ => Vec::new(),
    }
}

fn literal_string(node: &Node) -> Option<&str> {
    match node {
        Node::Literal(crate::s5_config_dsl::LiteralNode {
            value: crate::s5_config_dsl::LiteralValue::String(value),
            kind: "string",
        }) => Some(value),
        _ => None,
    }
}

fn rule_formula_nodes(rule: &crate::s5_config::RuleDefinition) -> Vec<(String, Node)> {
    let mut nodes = Vec::new();
    if let Ok(dictionary) = parse_dictionary(if rule.value_dictionary.is_empty() {
        None
    } else {
        Some(rule.value_dictionary.as_str())
    }) {
        for (_, node) in &dictionary {
            nodes.push(("取值字典".to_owned(), node.clone()));
        }
    }
    for (field, text) in [
        ("审核表达式", &rule.audit_expression),
        ("当前值表达式", &rule.current_expression),
        ("对比值表达式", &rule.comparison_expression),
        ("差值表达式", &rule.difference_expression),
        ("适用表达式", &rule.applicable_expression),
    ] {
        if !text.is_empty() {
            if let Ok(node) = parse_expression(text) {
                nodes.push((field.to_owned(), node));
            }
        }
    }
    nodes
}

/// Mirror of `_rule_references` (counts and shapes only; locations are part of
/// the not-yet-migrated scan stage).
fn rule_references(configuration: &Configuration, business: &str) -> RuleReferences {
    let mut template_refs: Vec<(
        (String, String, Vec<String>, Vec<String>),
        Vec<TemplateReference>,
    )> = Vec::new();
    let mut external_reference_count = 0usize;
    let mut dynamic_refs: Vec<DynamicReference> = Vec::new();
    for rule in configuration
        .rules
        .iter()
        .filter(|rule| rule.enabled && (business.is_empty() || rule.business == business))
    {
        for (field, node) in rule_formula_nodes(rule) {
            for call in calls_in(&node) {
                if call.name == "V_CUR" || call.name == "V_PRE" {
                    let (mut sheet, row_node, column_node) = if call.arguments.len() == 3 {
                        (
                            literal_string(&call.arguments[0])
                                .map(str::trim)
                                .unwrap_or_default()
                                .to_owned(),
                            Some(&call.arguments[1]),
                            Some(&call.arguments[2]),
                        )
                    } else if call.arguments.len() == 2 {
                        let sheet = call
                            .keywords
                            .iter()
                            .find(|(key, _)| key == "sheet")
                            .map(|(_, value)| value.clone())
                            .unwrap_or_default();
                        (sheet, Some(&call.arguments[0]), Some(&call.arguments[1]))
                    } else {
                        (String::new(), None, None)
                    };
                    if sheet.is_empty() {
                        sheet = rule.default_sheet.clone();
                    }
                    let is_dynamic = row_node
                        .and_then(literal_string)
                        .is_some_and(|value| value.trim() == "@当前行");
                    if is_dynamic {
                        dynamic_refs.push(DynamicReference {
                            business: rule.business.clone(),
                            sheet,
                        });
                        continue;
                    }
                    let row_path = row_node.map(path_argument).unwrap_or_default();
                    let column_path = column_node.map(path_argument).unwrap_or_default();
                    if sheet.is_empty() || row_path.is_empty() || column_path.is_empty() {
                        continue;
                    }
                    let key = (rule.business.clone(), sheet, row_path, column_path);
                    let reference = TemplateReference {
                        rule_id: rule.rule_id.clone(),
                        revision: rule.revision,
                        field: field.clone(),
                        source: call.name.clone(),
                    };
                    let entry = template_refs
                        .iter_mut()
                        .find(|(existing_key, _)| *existing_key == key);
                    match entry {
                        Some((_, owners)) => {
                            if !owners.contains(&reference) {
                                owners.push(reference);
                            }
                        }
                        None => template_refs.push((key, vec![reference])),
                    }
                } else if call.name == "V_EXT" {
                    let metric = call
                        .arguments
                        .first()
                        .and_then(literal_string)
                        .map(str::trim)
                        .unwrap_or_default();
                    if !metric.is_empty() {
                        external_reference_count += 1;
                    }
                }
            }
        }
    }
    RuleReferences {
        template_refs,
        external_reference_count,
        dynamic_refs,
    }
}

/// Validate one S5 audit configuration at the configuration level.
///
/// `template_path`, `template_dir`, `template_index_path`, and `external_path`
/// follow the Python signature; providing a template or external file selects
/// the not-yet-migrated workbook scan and fails with [`NOT_MIGRATED_SCAN`].
pub fn validate_s5_config(
    config_path: &Path,
    template_path: &str,
    template_dir: &str,
    template_index_path: &str,
    external_path: &str,
    business: &str,
    scope: &str,
) -> Result<Value, crate::s5_config::LoadError> {
    let _ = (template_dir, template_index_path);
    let full_scope = scope == "all" || scope == "all_enabled";
    let strict_current = scope == "current_business";
    if !matches!(
        scope,
        "current" | "current_business" | "all" | "all_enabled"
    ) {
        return Err(crate::s5_config::LoadError::configuration(format!(
            "未知的 S5 配置检查范围：{scope}"
        )));
    }
    let scope_label = if full_scope {
        "全量启用配置"
    } else {
        "当前业务启用配置"
    };

    let mut business = business.trim().to_owned();
    if !full_scope && business.is_empty() && !template_path.is_empty() {
        match infer_business_from_template(config_path, Path::new(template_path)) {
            Ok(inferred) => business = inferred,
            Err(inference_error) => {
                if strict_current {
                    return Ok(json!({
                        "valid": false, "complete": false, "scope": scope,
                        "scope_label": scope_label, "business": "",
                        "message": inference_error,
                        "rules": 0, "enabled_rules": 0, "disabled_rules": 0,
                        "checks": {
                            "expressions": {"status": "not_checked"},
                            "dictionary_keys": {"status": "not_checked"},
                            "template_addresses": {"status": "not_checked"},
                            "external_addresses": {"status": "not_checked"},
                        },
                        "warnings": [inference_error],
                        "configuration_errors": [configuration_error_to_json(&inference_error)],
                    }));
                }
            }
        }
    }
    if strict_current && (business.is_empty() || template_path.is_empty()) {
        let message = "当前业务检查需要先识别或选择匹配模板，再根据模板确定业务。";
        return Ok(json!({
            "valid": false, "complete": false, "scope": scope,
            "scope_label": scope_label, "business": business,
            "message": message,
            "rules": 0, "enabled_rules": 0, "disabled_rules": 0,
            "checks": {
                "expressions": {"status": "not_checked"},
                "dictionary_keys": {"status": "not_checked"},
                "template_addresses": {"status": "not_checked"},
                "external_addresses": {"status": "not_checked"},
            },
            "warnings": [message],
            "configuration_errors": [configuration_error_to_json(message)],
        }));
    }
    if !external_path.is_empty() || full_scope {
        return Err(crate::s5_config::LoadError::configuration(
            NOT_MIGRATED_SCAN,
        ));
    }

    let configuration = match load_configuration(config_path, false, &business) {
        Ok(configuration) => configuration,
        Err(error) => {
            return Ok(json!({
                "valid": false, "complete": false, "scope": scope,
                "scope_label": scope_label, "business": business,
                "message": "S5 配置语法或字典引用检查未通过",
                "rules": 0, "enabled_rules": 0, "disabled_rules": 0,
                "checks": {
                    "expressions": {"status": "failed"},
                    "dictionary_keys": {"status": "failed"},
                    "template_addresses": {"status": "not_checked"},
                    "external_addresses": {"status": "not_checked"},
                },
                "warnings": [],
                "configuration_errors": [configuration_error_to_json(&error.message)],
            }));
        }
    };
    let rules_by_business = |name: &str| -> usize {
        configuration
            .rules
            .iter()
            .filter(|rule| rule.business == name)
            .count()
    };
    let enabled_by_business = |name: &str| -> usize {
        configuration
            .rules
            .iter()
            .filter(|rule| rule.business == name && rule.enabled)
            .count()
    };
    let enabled_rules = configuration
        .rules
        .iter()
        .filter(|rule| rule.enabled)
        .count();
    let all_businesses: Vec<String> = if full_scope {
        let mut names: Vec<String> = configuration
            .rules
            .iter()
            .filter(|rule| rule.enabled)
            .map(|rule| rule.business.clone())
            .collect::<std::collections::BTreeSet<_>>()
            .into_iter()
            .collect();
        names.sort();
        names
    } else {
        let mut names: Vec<String> = configuration
            .rules
            .iter()
            .map(|rule| rule.business.clone())
            .collect::<std::collections::BTreeSet<_>>()
            .into_iter()
            .collect();
        names.sort();
        names
    };
    if enabled_rules == 0 {
        let warning = if !business.is_empty() && !full_scope {
            if !configuration.rules.is_empty() {
                format!("业务“{business}”的规则全部禁用；本次未执行表达式、字典键引用以及模板/外部文件地址检查。")
            } else {
                format!("配置中没有业务“{business}”的规则；本次未执行表达式、字典键引用以及模板/外部文件地址检查。")
            }
        } else {
            "当前没有启用规则；表达式、字典键引用以及模板/外部文件地址检查均未执行。启用需要检查的规则后再运行检查。".to_owned()
        };
        let empty_message = if !configuration.rules.is_empty() {
            format!(
                "配置文件结构可读取，共 {} 条规则，全部禁用。{warning}",
                configuration.rules.len()
            )
        } else {
            format!("配置文件结构可读取，{warning}")
        };
        return Ok(json!({
            "valid": true, "complete": false, "scope": scope,
            "scope_label": scope_label, "business": business,
            "message": empty_message,
            "version": configuration.version,
            "version_status": configuration.version_status,
            "rules": configuration.rules.len(),
            "enabled_rules": 0,
            "disabled_rules": configuration.rules.len(),
            "metrics": configuration.metrics.len(),
            "external_metrics": configuration.external_metrics.len(),
            "source_contracts": configuration.source_contracts.len(),
            "applicability_scopes": configuration.applicability_scopes.len(),
            "businesses": all_businesses.iter().map(|name| json!({
                "name": name, "rules": rules_by_business(name), "enabled_rules": 0,
            })).collect::<Vec<_>>(),
            "checks": {
                "expressions": {"status": "not_checked", "rules": 0},
                "dictionary_keys": {"status": "not_checked", "rules": 0},
                "template_addresses": {"status": "not_checked", "checked": 0, "resolved": 0,
                                        "errors": [], "message": "没有启用规则，未扫描模板地址"},
                "external_addresses": {"status": "not_checked", "checked": 0, "resolved": 0,
                                       "errors": [], "message": "没有启用规则，未扫描外部文件地址"},
            },
            "warnings": [warning],
            "configuration_errors": [],
        }));
    }

    let references = rule_references(&configuration, &business);
    let dynamic_for_business = references
        .dynamic_refs
        .iter()
        .filter(|item| item.business == business)
        .count();

    let template_check =
        if references.template_refs.is_empty() && references.dynamic_refs.is_empty() {
            json!({
                "status": "no_references", "business": business, "checked": 0,
                "resolved": 0, "errors": [], "dynamic_references": dynamic_for_business,
            })
        } else if template_path.is_empty() {
            json!({
                "status": "not_checked", "business": business, "checked": 0,
                "resolved": 0, "errors": [], "dynamic_references": dynamic_for_business,
                "message": "尚未选择或识别匹配模板，无法检查模板取数地址",
            })
        } else {
            check_template_addresses(
                &configuration,
                Path::new(template_path),
                &business,
                &references,
            )?
        };
    let template_file_name = if template_path.is_empty() {
        "".to_owned()
    } else {
        std::path::Path::new(template_path)
            .file_name()
            .map(|name| name.to_string_lossy().into_owned())
            .unwrap_or_default()
    };
    let business_checks = json!([{
        "name": business,
        "rules": if business.is_empty() {
            configuration.rules.len()
        } else {
            rules_by_business(&business)
        },
        "enabled_rules": if business.is_empty() {
            enabled_rules
        } else {
            enabled_by_business(&business)
        },
        "template": template_file_name,
        "template_addresses": template_check.as_object().map(|object| {
            object.iter().filter(|(key, _)| key.as_str() != "errors").map(|(key, value)| (key.clone(), value.clone())).collect::<serde_json::Map<String, Value>>()
        }).unwrap_or_default(),
    }]);
    let external_check = if references.external_reference_count == 0 {
        json!({"status": "no_references", "checked": 0, "resolved": 0, "errors": []})
    } else {
        json!({
            "status": "not_checked", "checked": 0, "resolved": 0, "errors": [],
            "message": "配置引用 V_EXT，但尚未选择外部文件，无法检查指标表头地址",
        })
    };
    let template_status_text = template_check
        .get("status")
        .and_then(Value::as_str)
        .unwrap_or_default();
    let external_status_text = external_check
        .get("status")
        .and_then(Value::as_str)
        .unwrap_or_default();
    let checks_complete = !matches!(template_status_text, "not_checked" | "partial")
        && !matches!(external_status_text, "not_checked" | "partial");
    let valid = true;

    let mut warnings: Vec<String> = Vec::new();
    if matches!(template_status_text, "not_checked" | "partial") {
        if let Some(message) = template_check.get("message").and_then(Value::as_str) {
            warnings.push(message.to_owned());
        }
    }
    if matches!(external_status_text, "not_checked" | "partial") {
        if let Some(message) = external_check.get("message").and_then(Value::as_str) {
            warnings.push(message.to_owned());
        }
    }
    let template_message = template_check.get("message").and_then(Value::as_str);
    let external_message = external_check.get("message").and_then(Value::as_str);
    let mut message_parts = vec![
        if valid {
            "配置语法与字典引用检查通过".to_owned()
        } else {
            "配置检查发现问题".to_owned()
        },
        format!(
            "已用正式审核表达式解析器检查{scope_label}中的 {enabled_rules} 条启用规则，并校验表达式变量均能在对应取值字典键中找到"
        ),
    ];
    if !business.is_empty() && !full_scope {
        message_parts.push(format!("当前业务：{business}"));
    }
    for message in [template_message, external_message].into_iter().flatten() {
        message_parts.push(message.to_owned());
    }
    for warning in &warnings {
        message_parts.push(warning.clone());
    }

    Ok(json!({
        "valid": valid,
        "complete": checks_complete,
        "scope": scope,
        "scope_label": scope_label,
        "business": if full_scope { "".to_owned() } else { business.clone() },
        "message": message_parts.join("\n"),
        "version": configuration.version,
        "version_status": configuration.version_status,
        "rules": configuration.rules.len(),
        "enabled_rules": enabled_rules,
        "disabled_rules": configuration.rules.iter().filter(|rule| !rule.enabled).count(),
        "metrics": configuration.metrics.len(),
        "external_metrics": configuration.external_metrics.len(),
        "source_contracts": configuration.source_contracts.len(),
        "applicability_scopes": configuration.applicability_scopes.len(),
        "businesses": all_businesses.iter().map(|name| json!({
            "name": name,
            "rules": rules_by_business(name),
            "enabled_rules": enabled_by_business(name),
        })).collect::<Vec<_>>(),
        "business_checks": business_checks,
        "checks": {
            "expressions": {"status": "passed", "rules": enabled_rules},
            "dictionary_keys": {"status": "passed", "rules": enabled_rules},
            "template_addresses": template_check,
            "external_addresses": external_check,
        },
        "warnings": warnings,
        "configuration_errors": [],
    }))
}

/// Mirror of `_configuration_rule_locations`: stable rule identity to its
/// physical Excel row plus the header column index map.
fn configuration_rule_locations(
    config_path: &Path,
) -> Result<
    BTreeMap<(String, i64), Vec<(usize, BTreeMap<String, usize>)>>,
    crate::s5_config::LoadError,
> {
    let dump = read_workbook(config_path).map_err(|error| {
        crate::s5_config::LoadError::configuration(format!("配置工作簿无法读取：{error}"))
    })?;
    let Some(sheet) = dump.sheets.iter().find(|sheet| sheet.name == "审核规则") else {
        return Ok(BTreeMap::new());
    };
    let headers = check_sheet_headers(sheet);
    let Some(id_header) = ["规则ID", "规则编号", "规则序号"]
        .into_iter()
        .find(|name| headers.iter().any(|header| header == name))
    else {
        return Ok(BTreeMap::new());
    };
    let id_column = headers
        .iter()
        .position(|header| header == id_header)
        .unwrap()
        + 1;
    let revision_column = headers
        .iter()
        .position(|header| header == "规则修订")
        .map(|position| position + 1);
    let mut columns: BTreeMap<String, usize> = BTreeMap::new();
    for (index, header) in headers.iter().enumerate() {
        if !header.is_empty() {
            columns.insert(header.clone(), index + 1);
        }
    }
    columns.entry("规则ID".to_owned()).or_insert(id_column);
    let mut locations: BTreeMap<(String, i64), Vec<(usize, BTreeMap<String, usize>)>> =
        BTreeMap::new();
    for row in 2..=check_sheet_extent_rows(sheet) {
        let rule_id = check_sheet_cell_text(sheet, id_column, row);
        if rule_id.is_empty() {
            continue;
        }
        let revision = revision_column
            .map(|column| check_sheet_cell_text(sheet, column, row))
            .and_then(|text| text.parse::<i64>().ok())
            .unwrap_or(1);
        locations
            .entry((rule_id, revision))
            .or_default()
            .push((row, columns.clone()));
    }
    Ok(locations)
}

fn check_sheet_headers(sheet: &excel_reader::SheetDump) -> Vec<String> {
    let mut max_column = 0usize;
    for address in sheet.values.keys() {
        if let Some((row, column)) = parse_sheet_cell(address) {
            if row == 1 && column > max_column {
                max_column = column;
            }
        }
    }
    (1..=max_column)
        .map(|column| {
            sheet
                .values
                .get(&format!("{}1", check_column_letter(column)))
                .map(check_sheet_cell_text_value)
                .unwrap_or_default()
        })
        .collect()
}

fn check_sheet_extent_rows(sheet: &excel_reader::SheetDump) -> usize {
    let mut max_row = 0usize;
    for address in sheet.values.keys() {
        if let Some((row, _)) = parse_sheet_cell(address) {
            max_row = max_row.max(row);
        }
    }
    max_row
}

fn check_column_letter(mut column: usize) -> String {
    let mut letters = Vec::new();
    while column > 0 {
        let digit = ((column - 1) % 26) as u8;
        letters.insert(0, (b'A' + digit) as char);
        column = (column - 1) / 26;
    }
    letters.into_iter().collect()
}

fn parse_sheet_cell(address: &str) -> Option<(usize, usize)> {
    let mut row_part = String::new();
    let mut column = 0usize;
    for character in address.chars() {
        if character.is_ascii_alphabetic() {
            column = column * 26 + (character.to_ascii_uppercase() as usize - 'A' as usize + 1);
        } else if character.is_ascii_digit() {
            row_part.push(character);
        } else {
            return None;
        }
    }
    let row = row_part.parse::<usize>().ok()?;
    (row > 0 && column > 0).then_some((row, column))
}

fn check_sheet_cell_text_value(value: &serde_json::Value) -> String {
    let kind = value.get("type").and_then(serde_json::Value::as_str);
    let raw = value.get("value");
    match kind {
        Some("string") => raw
            .and_then(serde_json::Value::as_str)
            .unwrap_or_default()
            .trim()
            .to_owned(),
        Some("number") => {
            let number = raw.and_then(serde_json::Value::as_f64).unwrap_or(0.0);
            if number.fract() == 0.0 && number.abs() < 1e15 {
                format!("{}", number as i64)
            } else {
                format!("{number}")
            }
        }
        Some("boolean") => {
            if raw.and_then(serde_json::Value::as_bool) == Some(true) {
                "True".to_owned()
            } else {
                "False".to_owned()
            }
        }
        _ => String::new(),
    }
}

fn check_sheet_cell_text(sheet: &excel_reader::SheetDump, column: usize, row: usize) -> String {
    sheet
        .values
        .get(&format!("{}{row}", check_column_letter(column)))
        .map(check_sheet_cell_text_value)
        .unwrap_or_default()
}

fn config_locations(
    owners: &[TemplateReference],
    row_map: &BTreeMap<(String, i64), Vec<(usize, BTreeMap<String, usize>)>>,
) -> Vec<Value> {
    let mut locations = Vec::new();
    for reference in owners {
        let Some(rows) = row_map.get(&(reference.rule_id.clone(), reference.revision)) else {
            continue;
        };
        for (row, columns) in rows {
            let column = columns.get(&reference.field).copied();
            locations.push(json!({
                "sheet": "审核规则",
                "row": row,
                "column": column,
                "cell": column
                    .map(|column| format!("{}{row}", check_column_letter(column)))
                    .unwrap_or_default(),
                "field": reference.field,
                "rule_id": reference.rule_id,
            }));
        }
    }
    locations
}

/// Mirror of `_check_template_addresses` for the current business with a
/// selected template workbook. External-file and full-scope branches remain
/// with the not-migrated boundary above.
fn check_template_addresses(
    configuration: &Configuration,
    template_path: &Path,
    business: &str,
    references: &RuleReferences,
) -> Result<Value, crate::s5_config::LoadError> {
    let dynamic: Vec<&DynamicReference> = references
        .dynamic_refs
        .iter()
        .filter(|item| item.business == business)
        .collect();
    if references.template_refs.is_empty() && dynamic.is_empty() {
        return Ok(json!({
            "status": "no_references", "business": business, "checked": 0,
            "resolved": 0, "errors": [], "dynamic_references": dynamic.len(),
        }));
    }
    if !template_path.is_file() {
        return Ok(json!({
            "status": "failed", "business": business, "checked": 0, "resolved": 0,
            "errors": [{"message": format!("匹配模板不存在：{}", template_path.display()), "locations": []}],
            "dynamic_references": dynamic.len(),
        }));
    }
    if business.is_empty() {
        return Ok(json!({
            "status": "not_checked", "business": "", "checked": 0, "resolved": 0,
            "errors": [], "dynamic_references": dynamic.len(),
            "message": "已选择匹配模板，但尚未识别审核业务；不能确定应使用哪组字典地址",
        }));
    }
    let row_map = configuration_rule_locations(&configuration_source_path()).unwrap_or_default();

    let mut grouped_order: Vec<String> = Vec::new();
    let mut grouped: BTreeMap<String, Vec<(Vec<String>, Vec<String>)>> = BTreeMap::new();
    let mut references_by_key: BTreeMap<
        (String, String, Vec<String>, Vec<String>),
        Vec<TemplateReference>,
    > = BTreeMap::new();
    for (key, owners) in &references.template_refs {
        if key.0 != business {
            continue;
        }
        if !grouped_order.iter().any(|sheet| sheet == &key.1) {
            grouped_order.push(key.1.clone());
        }
        let pairs = grouped.entry(key.1.clone()).or_default();
        if !pairs
            .iter()
            .any(|(row, column)| row == &key.2 && column == &key.3)
        {
            pairs.push((key.2.clone(), key.3.clone()));
        }
        references_by_key.insert(key.clone(), owners.clone());
    }

    let mut errors: Vec<Value> = Vec::new();
    let mut checked = 0usize;
    let mut resolved = 0usize;
    let scan = match ScanBook::open(template_path) {
        Ok(scan) => scan,
        Err(error) => {
            errors.push(json!({
                "message": format!("匹配模板无法读取或扫描：{error}"),
                "locations": [],
            }));
            return Ok(json!({
                "status": "failed", "business": business, "checked": 0, "resolved": 0,
                "errors": errors, "dynamic_references": dynamic.len(),
            }));
        }
    };
    let scan_reference = &scan;
    for sheet_name in &grouped_order {
        let Some(key_pairs) = grouped.get(sheet_name) else {
            continue;
        };
        let book = scan_reference;
        if !book.sheet_names().iter().any(|name| name == sheet_name) {
            let mut pair_locations: Vec<Value> = Vec::new();
            for (row_path, column_path) in key_pairs {
                checked += 1;
                let owners = references_by_key
                    .get(&(
                        business.to_owned(),
                        sheet_name.clone(),
                        row_path.clone(),
                        column_path.clone(),
                    ))
                    .cloned()
                    .unwrap_or_default();
                pair_locations.extend(config_locations(&owners, &row_map));
            }
            errors.push(json!({
                "message": format!("匹配模板缺少工作表：{sheet_name}"),
                "locations": pair_locations,
            }));
            continue;
        }
        let scan_result = scan_rule_paths(book, template_path, sheet_name, key_pairs);
        for (row_path, column_path) in key_pairs {
            checked += 1;
            let owners = references_by_key
                .get(&(
                    business.to_owned(),
                    sheet_name.clone(),
                    row_path.clone(),
                    column_path.clone(),
                ))
                .cloned()
                .unwrap_or_default();
            let key = (row_path.clone(), column_path.clone());
            let matches = scan_result.locations.get(&key);
            let match_count = matches.map(|locations| locations.len()).unwrap_or(0);
            if match_count == 1 {
                resolved += 1;
                continue;
            }
            let reason = if match_count > 1 {
                let cells: Vec<String> = matches
                    .into_iter()
                    .flatten()
                    .take(10)
                    .map(|location| {
                        format!("{}{}", check_column_letter(location.column), location.row)
                    })
                    .collect();
                format!("匹配到多个单元格：{}", cells.join(", "))
            } else {
                let row_found = scan_result
                    .row_locations
                    .get(row_path)
                    .map(|rows| !rows.is_empty())
                    .unwrap_or(false);
                let column_found = scan_result
                    .column_locations
                    .get(column_path)
                    .map(|columns| !columns.is_empty())
                    .unwrap_or(false);
                if !row_found {
                    "未找到完整行路径".to_owned()
                } else if !column_found {
                    "未找到完整列路径".to_owned()
                } else {
                    "行路径和列路径存在，但没有唯一有效交点".to_owned()
                }
            };
            errors.push(json!({
                "message": format!(
                    "模板取值地址无法唯一定位：{} / {} / {}；{}",
                    sheet_name,
                    row_path.join("@"),
                    column_path.join("@"),
                    reason
                ),
                "locations": config_locations(&owners, &row_map),
            }));
        }
    }
    let status = if !errors.is_empty() {
        "failed"
    } else if !dynamic.is_empty() {
        "partial"
    } else {
        "passed"
    };
    let mut message =
        format!("业务“{business}”检查 {checked} 个静态模板取数地址，成功定位 {resolved} 个");
    if !dynamic.is_empty() {
        message.push_str(&format!(
            "；另有 {} 个 @当前行 地址尚未展开检查，需结合适用范围核对",
            dynamic.len()
        ));
    }
    Ok(json!({
        "status": status, "business": business, "checked": checked,
        "resolved": resolved, "errors": errors, "dynamic_references": dynamic.len(),
        "message": message,
    }))
}

thread_local! {
    static CONFIG_SOURCE_PATH: std::cell::RefCell<Option<std::path::PathBuf>> =
        const { std::cell::RefCell::new(None) };
}

/// Record the configuration path for the diagnostics that re-read the rule
/// sheet (the parity runner sets it before validation).
pub fn set_configuration_source_path(path: &Path) {
    CONFIG_SOURCE_PATH.with(|cell| {
        *cell.borrow_mut() = Some(path.to_path_buf());
    });
}

fn configuration_source_path() -> std::path::PathBuf {
    CONFIG_SOURCE_PATH
        .with(|cell| cell.borrow().clone())
        .unwrap_or_default()
}
