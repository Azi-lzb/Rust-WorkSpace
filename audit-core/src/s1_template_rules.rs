//! S1 template rule discovery and validation for the Rust audit workflow.
//!
//! The reader owns stored values, formula text, comments, and defined names.
//! This module turns those workbook facts into Audit's immutable template
//! definition; it never recalculates formulas or modifies the workbook.

use crate::{
    models::{AuditRule, CopyRange, TemplateDefinition, PreflightItem},
    name_config::{
        load_s1_conditional_mappings, load_s1_formula_mappings, named_areas_from_defined_names,
        resolve_named_ranges_for_mappings, FeatureMapping, FORMULA_ISSUE_EXTRACT_FUNCTION,
        STRUCTURE_COMPARE_FUNCTION,
    },
    template::{
        clean_rule_comment, normalize_template_name, parse_comment_rule_id, parse_rule_rows,
        REQUIRED_HEADERS,
    },
};
use anyhow::{bail, ensure, Context, Result};
use excel_reader::{CellValue, DefinedNameRecord, SheetDump, WorkbookDump};
use serde::Serialize;
use serde_json::{Map, Value};
use std::{
    collections::{BTreeMap, BTreeSet},
    path::Path,
    time::Instant,
};

const RESULT_KEYWORDS: [&str; 4] = ["错误", "硬性", "错", "软性"];
const ENGINE: &str = "calamine+namespace_aware_ooxml";

#[derive(Clone, Debug, Serialize)]
pub struct S1TemplateRuleSummary {
    pub template_mode: &'static str,
    pub engine: &'static str,
    pub formula_mapping_configured: bool,
    pub formula_named_ranges: usize,
    pub copy_ranges: usize,
    pub rules_total: usize,
    pub rules_enabled: usize,
    pub formula_cells: usize,
    pub rule_comments: usize,
    pub explicit_comment_ids: usize,
    pub keyword_rules: usize,
    pub error_only_rules: usize,
    pub structure_ranges: usize,
    pub extraction_ranges: usize,
    pub read_time_ms: u128,
    pub elapsed_ms: u128,
}

#[derive(Clone, Debug)]
pub struct S1TemplateRulePlan {
    pub definition: TemplateDefinition,
    pub summary: S1TemplateRuleSummary,
    formula_reference_errors: Vec<String>,
}

#[derive(Clone, Debug, Serialize)]
pub struct S1TemplateHealthReport {
    pub passed: bool,
    pub enabled_rules: usize,
    pub result_rules: usize,
    pub structure_ranges: usize,
    pub broken_reference_rules: usize,
    pub items: Vec<PreflightItem>,
}

/// Assess the static S1 template health checks used by the Audit workflow.
///
/// `effective_structure_ranges` may contain CLI-resolved overrides. Formula
/// references are read from the already-loaded Rust template snapshot; this
/// function never recalculates formulas or opens a workbook for modification.
pub fn assess_s1_template_health(
    plan: &S1TemplateRulePlan,
    effective_structure_ranges: &[CopyRange],
) -> S1TemplateHealthReport {
    let enabled_rules = plan
        .definition
        .rules
        .iter()
        .filter(|rule| rule.enabled)
        .collect::<Vec<_>>();
    let result_rules = enabled_rules
        .iter()
        .filter(|rule| rule.result_mode != "error_only")
        .count();
    let structure_ranges = effective_structure_ranges.len();
    let mut items = vec![
        PreflightItem {
            category: "规则".to_owned(),
            level: if enabled_rules.is_empty() {
                "错误"
            } else {
                "提示"
            }
            .to_owned(),
            status: if enabled_rules.is_empty() {
                "需处理"
            } else {
                "信息"
            }
            .to_owned(),
            sheet_name: String::new(),
            location: String::new(),
            message: if enabled_rules.is_empty() {
                "模板没有启用的校验规则".to_owned()
            } else {
                format!(
                    "启用公式 {} 个，其中问题规则 {} 个、公式健康检查 {} 个",
                    enabled_rules.len(),
                    result_rules,
                    enabled_rules.len() - result_rules
                )
            },
        },
        PreflightItem {
            category: "规则".to_owned(),
            level: if result_rules == 0 {
                "错误"
            } else {
                "提示"
            }
            .to_owned(),
            status: if result_rules == 0 {
                "需处理"
            } else {
                "信息"
            }
            .to_owned(),
            sheet_name: String::new(),
            location: String::new(),
            message: if result_rules == 0 {
                "没有识别到会返回问题标记的公式，审核结果可能始终为零".to_owned()
            } else {
                "问题规则识别正常".to_owned()
            },
        },
        PreflightItem {
            category: "表结构区域".to_owned(),
            level: if structure_ranges == 0 {
                "错误"
            } else {
                "提示"
            }
            .to_owned(),
            status: if structure_ranges == 0 {
                "需处理"
            } else {
                "通过"
            }
            .to_owned(),
            sheet_name: String::new(),
            location: String::new(),
            message: if structure_ranges == 0 {
                "模板没有“表结构区域”命名区域，无法精确核对报送文件结构".to_owned()
            } else {
                format!("已设置 {structure_ranges} 个“表结构区域”命名区域")
            },
        },
    ];

    if !plan.formula_reference_errors.is_empty() {
        items.push(PreflightItem {
            category: "公式".to_owned(),
            level: "警告".to_owned(),
            status: "关注".to_owned(),
            sheet_name: String::new(),
            location: plan
                .formula_reference_errors
                .iter()
                .take(10)
                .cloned()
                .collect::<Vec<_>>()
                .join("、"),
            message: format!(
                "发现 {} 个公式含 #REF! 引用；实际计算触发错误时会写入问题结果",
                plan.formula_reference_errors.len()
            ),
        });
    }

    let passed = !items.iter().any(|item| item.level == "错误");
    items.insert(
        0,
        PreflightItem {
            category: "总体".to_owned(),
            level: if passed { "提示" } else { "错误" }.to_owned(),
            status: if passed { "通过" } else { "不通过" }.to_owned(),
            sheet_name: String::new(),
            location: String::new(),
            message: if passed {
                "模板结构可用于审核".to_owned()
            } else {
                "模板存在必须处理的问题".to_owned()
            },
        },
    );

    S1TemplateHealthReport {
        passed,
        enabled_rules: enabled_rules.len(),
        result_rules,
        structure_ranges,
        broken_reference_rules: plan.formula_reference_errors.len(),
        items,
    }
}

/// Read and validate one S1 template using the configured feature mappings.
///
/// Formula text is inspected as stored. This function does not calculate or
/// refresh formula caches and does not open the workbook for modification.
pub fn load_s1_template_rule_plan(
    template_path: impl AsRef<Path>,
    mapping_config: Option<&Path>,
) -> Result<S1TemplateRulePlan> {
    let started = Instant::now();
    let template_path = template_path.as_ref();
    ensure!(
        template_path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| {
                extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm")
            }),
        "S1 公式模板规则读取目前支持 .xlsx 和 .xlsm"
    );

    let formula_mappings = load_s1_formula_mappings(mapping_config, template_path)?;
    let conditional_mappings = load_s1_conditional_mappings(mapping_config, template_path)?;
    let read_started = Instant::now();
    let workbook = excel_reader::read_workbook(template_path).context("读取 S1 模板工作簿")?;
    let defined_names =
        excel_reader::read_defined_names(template_path).context("读取 S1 模板命名区域")?;
    build_s1_template_rule_plan(
        template_path,
        &formula_mappings,
        &conditional_mappings,
        &workbook,
        &defined_names,
        read_started.elapsed().as_millis(),
        started,
    )
}

fn build_s1_template_rule_plan(
    template_path: &Path,
    formula_mappings: &[FeatureMapping],
    conditional_mappings: &[FeatureMapping],
    workbook: &WorkbookDump,
    defined_names: &[DefinedNameRecord],
    read_time_ms: u128,
    started: Instant,
) -> Result<S1TemplateRulePlan> {
    let workbook_sheets = workbook
        .sheets
        .iter()
        .map(|sheet| sheet.name.clone())
        .collect::<Vec<_>>();
    let areas = named_areas_from_defined_names(defined_names);
    let formula_named_ranges =
        resolve_named_ranges_for_mappings(formula_mappings, &areas, &workbook_sheets);
    let structure_mappings = conditional_mappings
        .iter()
        .filter(|mapping| mapping.feature_type == STRUCTURE_COMPARE_FUNCTION)
        .cloned()
        .collect::<Vec<_>>();
    let structure_ranges =
        resolve_named_ranges_for_mappings(&structure_mappings, &areas, &workbook_sheets);
    let extraction_mappings = formula_mappings
        .iter()
        .filter(|mapping| mapping.feature_type == FORMULA_ISSUE_EXTRACT_FUNCTION)
        .cloned()
        .collect::<Vec<_>>();
    let extraction_ranges =
        resolve_named_ranges_for_mappings(&extraction_mappings, &areas, &workbook_sheets);

    let (rules, copy_ranges, template_mode, formula_cells, rule_comments, explicit_comment_ids) =
        if let Some(rule_sheet) = workbook
            .sheets
            .iter()
            .find(|sheet| sheet.name == "审核规则")
        {
            let rules = read_structured_rules(rule_sheet, workbook)?;
            let formula_cells = rules.iter().filter(|rule| rule.enabled).count();
            let copy_ranges = rules
                .iter()
                .filter(|rule| rule.enabled)
                .map(|rule| CopyRange {
                    sheet_name: rule.sheet_name.clone(),
                    address: rule.copy_range.clone(),
                })
                .collect();
            (rules, copy_ranges, "structured", formula_cells, 0, 0)
        } else {
            if !formula_mappings.is_empty() && formula_named_ranges.is_empty() {
                bail!("公式校验功能对应的命名区域中没有可用区域");
            }
            let (rules, formula_cells, rule_comments, explicit_comment_ids) =
                discover_named_formula_rules(workbook, &formula_named_ranges, template_path)?;
            if !formula_mappings.is_empty() && rules.is_empty() {
                bail!("公式校验功能对应的命名区域中没有公式");
            }
            (
                rules,
                formula_named_ranges.clone(),
                "named_formula_ranges",
                formula_cells,
                rule_comments,
                explicit_comment_ids,
            )
        };

    let rules_enabled = rules.iter().filter(|rule| rule.enabled).count();
    let keyword_rules = rules
        .iter()
        .filter(|rule| rule.result_mode == "keyword")
        .count();
    let error_only_rules = rules
        .iter()
        .filter(|rule| rule.result_mode == "error_only")
        .count();
    let definition = TemplateDefinition {
        rules,
        copy_ranges: copy_ranges.clone(),
        structured: template_mode == "structured",
        structure_ranges: structure_ranges.clone(),
        extraction_ranges: extraction_ranges.clone(),
    };
    let formula_reference_errors = enabled_formula_reference_errors(&definition, workbook);
    let summary = S1TemplateRuleSummary {
        template_mode,
        engine: ENGINE,
        formula_mapping_configured: !formula_mappings.is_empty(),
        formula_named_ranges: if template_mode == "structured" {
            0
        } else {
            formula_named_ranges.len()
        },
        copy_ranges: copy_ranges.len(),
        rules_total: definition.rules.len(),
        rules_enabled,
        formula_cells,
        rule_comments,
        explicit_comment_ids,
        keyword_rules,
        error_only_rules,
        structure_ranges: structure_ranges.len(),
        extraction_ranges: extraction_ranges.len(),
        read_time_ms: read_time_ms.max(workbook.elapsed_ms),
        elapsed_ms: started.elapsed().as_millis(),
    };
    Ok(S1TemplateRulePlan {
        definition,
        summary,
        formula_reference_errors,
    })
}

fn enabled_formula_reference_errors(
    definition: &TemplateDefinition,
    workbook: &WorkbookDump,
) -> Vec<String> {
    let sheets = workbook
        .sheets
        .iter()
        .map(|sheet| (sheet.name.as_str(), sheet))
        .collect::<BTreeMap<_, _>>();
    definition
        .rules
        .iter()
        .filter(|rule| rule.enabled)
        .filter_map(|rule| {
            let formula = sheets
                .get(rule.sheet_name.as_str())?
                .formulas
                .get(&rule.formula_cell)?;
            formula
                .to_uppercase()
                .contains("#REF!")
                .then(|| format!("{}!{}", rule.sheet_name, rule.formula_cell))
        })
        .collect()
}

fn read_structured_rules(
    rule_sheet: &SheetDump,
    workbook: &WorkbookDump,
) -> Result<Vec<AuditRule>> {
    let rows = sheet_rows_as_text(rule_sheet)?;
    let Some(header_row) = rows.first() else {
        bail!("模板“审核规则”工作表为空");
    };
    let headers = header_row
        .iter()
        .map(|value| value.trim().to_owned())
        .collect::<Vec<_>>();
    let missing = REQUIRED_HEADERS
        .iter()
        .filter(|required| !headers.iter().any(|header| header.as_str() == **required))
        .copied()
        .collect::<Vec<_>>();
    if !missing.is_empty() {
        bail!("审核规则表缺少字段：{}", missing.join("、"));
    }

    let max_column = headers.len();
    let records = rows
        .iter()
        .skip(1)
        .map(|row| {
            let mut record = Map::new();
            for (column, header) in headers.iter().take(max_column).enumerate() {
                let value = row.get(column).cloned().unwrap_or_default();
                record.insert(header.clone(), Value::String(value));
            }
            Value::Object(record)
        })
        .collect::<Vec<_>>();
    let rules = parse_rule_rows(&records).map_err(anyhow::Error::new)?;
    let sheets = workbook
        .sheets
        .iter()
        .map(|sheet| (sheet.name.as_str(), sheet))
        .collect::<BTreeMap<_, _>>();
    let mut errors = Vec::new();
    for rule in rules.iter().filter(|rule| rule.enabled) {
        let Some(sheet) = sheets.get(rule.sheet_name.as_str()) else {
            errors.push(format!(
                "{}：{}!{}没有公式",
                rule.rule_id, rule.sheet_name, rule.formula_cell
            ));
            continue;
        };
        if !sheet
            .formulas
            .get(&rule.formula_cell)
            .is_some_and(|formula| !formula.trim().is_empty())
        {
            errors.push(format!(
                "{}：{}!{}没有公式",
                rule.rule_id, rule.sheet_name, rule.formula_cell
            ));
        }
    }
    if !errors.is_empty() {
        bail!("{}", errors.join("；"));
    }
    Ok(rules)
}

fn sheet_rows_as_text(sheet: &SheetDump) -> Result<Vec<Vec<String>>> {
    let mut values = BTreeMap::<(usize, usize), String>::new();
    let mut max_row = 0usize;
    let mut max_column = 0usize;
    for (address, value) in &sheet.values {
        let Some((row, column)) = parse_cell_address(address) else {
            continue;
        };
        let value = CellValue::from_probe_json(value)
            .with_context(|| format!("解码审核规则表单元格 {address}"))?;
        let text = cell_text(value);
        if !text.is_empty() {
            max_row = max_row.max(row);
            max_column = max_column.max(column);
            values.insert((row, column), text);
        }
    }
    for (address, formula) in &sheet.formulas {
        let Some((row, column)) = parse_cell_address(address) else {
            continue;
        };
        max_row = max_row.max(row);
        max_column = max_column.max(column);
        values.insert((row, column), normalize_formula(formula));
    }
    let mut rows = vec![vec![String::new(); max_column]; max_row];
    for ((row, column), value) in values {
        rows[row - 1][column - 1] = value;
    }
    Ok(rows)
}

fn cell_text(value: CellValue) -> String {
    match value {
        CellValue::Empty => String::new(),
        CellValue::Text(value) => value.trim().to_owned(),
        CellValue::Number(value) | CellValue::DateTimeSerial(value) => value.to_string(),
        CellValue::Integer(value) => value.to_string(),
        CellValue::Boolean(true) => "True".to_owned(),
        CellValue::Boolean(false) => "False".to_owned(),
        CellValue::DateTimeIso(value) | CellValue::DurationIso(value) | CellValue::Error(value) => {
            value.trim().to_owned()
        }
    }
}

fn discover_named_formula_rules(
    workbook: &WorkbookDump,
    ranges: &[CopyRange],
    template_path: &Path,
) -> Result<(Vec<AuditRule>, usize, usize, usize)> {
    let sheets = workbook
        .sheets
        .iter()
        .map(|sheet| (sheet.name.as_str(), sheet))
        .collect::<BTreeMap<_, _>>();
    let report_code = template_path
        .file_stem()
        .map(|stem| stem.to_string_lossy().into_owned())
        .map(|stem| normalize_template_name(&stem))
        .unwrap_or_default();
    let mut rules = Vec::new();
    let mut seen_cells = BTreeSet::<(String, String)>::new();
    let mut seen_rule_ids = BTreeMap::<String, (String, String)>::new();
    let mut formula_cells = 0usize;
    let mut rule_comments = 0usize;
    let mut explicit_comment_ids = 0usize;

    for range in ranges {
        let sheet = sheets.get(range.sheet_name.as_str()).with_context(|| {
            format!(
                "公式命名区域引用了模板中不存在的工作表：{}",
                range.sheet_name
            )
        })?;
        let bounds = parse_range(&range.address).with_context(|| {
            format!("无效的公式校验区域：{}!{}", range.sheet_name, range.address)
        })?;
        let mut formula_positions = sheet
            .formulas
            .iter()
            .filter_map(|(address, formula)| {
                let (row, column) = parse_cell_address(address)?;
                (within(row, column, bounds) && !formula.trim().is_empty())
                    .then(|| (row, column, format_cell_address(row, column), formula))
            })
            .collect::<Vec<_>>();
        formula_positions.sort_by_key(|(row, column, _, _)| (*row, *column));

        for (_, _, coordinate, formula) in formula_positions {
            let identity = (range.sheet_name.clone(), coordinate.clone());
            if !seen_cells.insert(identity.clone()) {
                continue;
            }
            formula_cells += 1;
            let raw_comment = sheet.comments.get(&coordinate).cloned().unwrap_or_default();
            if !raw_comment.trim().is_empty() {
                rule_comments += 1;
            }
            let explicit_id = parse_comment_rule_id(&raw_comment)
                .map_err(anyhow::Error::new)
                .with_context(|| {
                    format!("解析公式规则批注：{}!{}", range.sheet_name, coordinate)
                })?;
            if !explicit_id.is_empty() {
                explicit_comment_ids += 1;
            }
            let rule_id = if explicit_id.is_empty() {
                format!("{}-{coordinate}", range.sheet_name)
            } else {
                explicit_id
            };
            if let Some(previous) = seen_rule_ids.get(&rule_id) {
                if previous != &identity {
                    bail!(
                        "规则编号“{rule_id}”重复：{}!{}、{}!{}",
                        previous.0,
                        previous.1,
                        identity.0,
                        identity.1
                    );
                }
            } else {
                seen_rule_ids.insert(rule_id.clone(), identity.clone());
            }
            rules.push(AuditRule {
                rule_id,
                enabled: true,
                report_code: report_code.clone(),
                sheet_name: range.sheet_name.clone(),
                formula_cell: coordinate.clone(),
                target_cell: coordinate,
                severity: "错误".to_owned(),
                message: clean_rule_comment(&raw_comment),
                copy_range: range.address.clone(),
                result_mode: if RESULT_KEYWORDS
                    .iter()
                    .any(|keyword| formula.contains(keyword))
                {
                    "keyword".to_owned()
                } else {
                    "error_only".to_owned()
                },
                value_from_result: true,
            });
        }
    }
    Ok((rules, formula_cells, rule_comments, explicit_comment_ids))
}

#[derive(Clone, Copy)]
struct CellBounds {
    min_row: usize,
    min_column: usize,
    max_row: usize,
    max_column: usize,
}

fn parse_range(value: &str) -> Option<CellBounds> {
    let value = value.replace('$', "");
    let mut parts = value.split(':');
    let (first_row, first_column) = parse_cell_address(parts.next()?)?;
    let (last_row, last_column) = match parts.next() {
        Some(last) => parse_cell_address(last)?,
        None => (first_row, first_column),
    };
    if parts.next().is_some() || first_row > last_row || first_column > last_column {
        return None;
    }
    Some(CellBounds {
        min_row: first_row,
        min_column: first_column,
        max_row: last_row,
        max_column: last_column,
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

fn format_cell_address(row: usize, column: usize) -> String {
    let mut remaining = column;
    let mut letters = Vec::new();
    while remaining > 0 {
        let index = (remaining - 1) % 26;
        letters.push((b'A' + index as u8) as char);
        remaining = (remaining - 1) / 26;
    }
    letters.reverse();
    format!("{}{row}", letters.into_iter().collect::<String>())
}

fn within(row: usize, column: usize, bounds: CellBounds) -> bool {
    (bounds.min_row..=bounds.max_row).contains(&row)
        && (bounds.min_column..=bounds.max_column).contains(&column)
}

fn normalize_formula(formula: &str) -> String {
    if formula.starts_with('=') {
        formula.to_owned()
    } else {
        format!("={formula}")
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use excel_writer::Workbook;
    use std::{
        fs,
        path::PathBuf,
        time::{SystemTime, UNIX_EPOCH},
    };

    fn temporary_workbook_path(label: &str) -> PathBuf {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("system clock after Unix epoch")
            .as_nanos();
        std::env::temp_dir().join(format!(
            "audit_s1_template_rules_{label}_{}_{}.xlsx",
            std::process::id(),
            nonce
        ))
    }

    #[test]
    fn discovers_named_formula_rules_in_row_major_order() {
        let path = temporary_workbook_path("named");
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();
        worksheet
            .set_name("检查表")
            .expect("set rule worksheet name");
        worksheet
            .write_formula(0, 0, "=IF(TRUE,\"错误|字段|指标|说明\",\"错误\")")
            .expect("write first rule formula");
        worksheet
            .write_formula(2, 0, "=IF(TRUE,\"通过\",\"通过\")")
            .expect("write second rule formula");
        workbook
            .define_name("校验区域", "='检查表'!$A$1:$A$3")
            .expect("define formula range");
        workbook
            .define_name("表结构区域", "='检查表'!$B$1")
            .expect("define structure range");
        workbook.save(&path).expect("save formula template");

        let plan = load_s1_template_rule_plan(&path, None).expect("discover formula rules");
        let health = assess_s1_template_health(&plan, &plan.definition.structure_ranges);
        fs::remove_file(&path).expect("remove test template");

        assert_eq!(plan.summary.template_mode, "named_formula_ranges");
        assert_eq!(plan.summary.formula_named_ranges, 1);
        assert_eq!(plan.summary.formula_cells, 2);
        assert_eq!(plan.definition.rules.len(), 2);
        assert_eq!(plan.definition.rules[0].formula_cell, "A1");
        assert_eq!(plan.definition.rules[0].result_mode, "keyword");
        assert_eq!(plan.definition.rules[1].formula_cell, "A3");
        assert_eq!(plan.definition.rules[1].result_mode, "error_only");
        assert!(health.passed);
        assert_eq!(health.enabled_rules, 2);
        assert_eq!(health.result_rules, 1);
        assert_eq!(health.structure_ranges, 1);
        assert_eq!(health.broken_reference_rules, 0);
    }

    #[test]
    fn template_health_blocks_missing_problem_rules_and_structure_ranges() {
        let path = temporary_workbook_path("unhealthy");
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();
        worksheet
            .set_name("检查表")
            .expect("set rule worksheet name");
        worksheet
            .write_formula(0, 0, "=#REF!+1")
            .expect("write broken formula");
        workbook
            .define_name("校验区域", "='检查表'!$A$1")
            .expect("define formula range");
        workbook
            .save(&path)
            .expect("save unhealthy formula template");

        let plan = load_s1_template_rule_plan(&path, None).expect("load unhealthy template");
        let health = assess_s1_template_health(&plan, &plan.definition.structure_ranges);
        fs::remove_file(&path).expect("remove unhealthy template");

        assert!(!health.passed);
        assert_eq!(health.enabled_rules, 1);
        assert_eq!(health.result_rules, 0);
        assert_eq!(health.structure_ranges, 0);
        assert_eq!(health.broken_reference_rules, 1);
        assert_eq!(health.items[0].status, "不通过");
        assert!(health.items.iter().any(|item| {
            item.category == "规则" && item.message.contains("问题标记") && item.level == "错误"
        }));
        assert!(health.items.iter().any(|item| {
            item.category == "公式" && item.location == "检查表!A1" && item.level == "警告"
        }));
    }

    #[test]
    fn loads_and_validates_structured_rule_table() {
        let path = temporary_workbook_path("structured");
        let mut workbook = Workbook::new();
        {
            let worksheet = workbook.add_worksheet();
            worksheet
                .set_name("审核规则")
                .expect("set structured rules sheet");
            for (column, header) in REQUIRED_HEADERS.iter().enumerate() {
                worksheet
                    .write_string(0, column as u16, *header)
                    .expect("write rule header");
            }
            for (column, value) in [
                "R-01",
                "是",
                "报表甲",
                "检查表",
                "A1",
                "B1",
                "错误",
                "核对错误项",
            ]
            .iter()
            .enumerate()
            {
                worksheet
                    .write_string(1, column as u16, *value)
                    .expect("write structured rule");
            }
        }
        let worksheet = workbook.add_worksheet();
        worksheet
            .set_name("检查表")
            .expect("set formula sheet name");
        worksheet
            .write_formula(0, 0, "=1+1")
            .expect("write referenced rule formula");
        workbook.save(&path).expect("save structured template");

        let plan = load_s1_template_rule_plan(&path, None).expect("load structured rules");
        fs::remove_file(&path).expect("remove test template");

        assert!(plan.definition.structured);
        assert_eq!(plan.summary.template_mode, "structured");
        assert_eq!(plan.definition.rules.len(), 1);
        assert_eq!(plan.definition.rules[0].rule_id, "R-01");
        assert_eq!(plan.definition.rules[0].formula_cell, "A1");
        assert_eq!(plan.definition.rules[0].result_mode, "nonempty");
        assert!(!plan.definition.rules[0].value_from_result);
        assert_eq!(plan.definition.copy_ranges.len(), 1);
    }

    #[test]
    fn rejects_structured_rules_that_reference_non_formula_cells() {
        let path = temporary_workbook_path("missing_formula");
        let mut workbook = Workbook::new();
        {
            let worksheet = workbook.add_worksheet();
            worksheet
                .set_name("审核规则")
                .expect("set structured rules sheet");
            for (column, header) in REQUIRED_HEADERS.iter().enumerate() {
                worksheet
                    .write_string(0, column as u16, *header)
                    .expect("write rule header");
            }
            for (column, value) in [
                "R-02",
                "是",
                "报表甲",
                "检查表",
                "A1",
                "B1",
                "错误",
                "核对错误项",
            ]
            .iter()
            .enumerate()
            {
                worksheet
                    .write_string(1, column as u16, *value)
                    .expect("write structured rule");
            }
        }
        let worksheet = workbook.add_worksheet();
        worksheet
            .set_name("检查表")
            .expect("set formula sheet name");
        worksheet
            .write_string(0, 0, "普通值")
            .expect("write non-formula value");
        workbook
            .save(&path)
            .expect("save invalid structured template");

        let error =
            load_s1_template_rule_plan(&path, None).expect_err("non-formula references must fail");
        fs::remove_file(&path).expect("remove test template");
        assert!(format!("{error:#}").contains("R-02：检查表!A1没有公式"));
    }
}
