//! S3-F01 two-period comparison: configuration loading and base record pairing.
//!
//! This is the first Rust slice of the legacy `run_comparison` pipeline. It
//! keeps the original eight-field record identity and the five/eight-segment
//! expression-rule provenance so later rule execution can enrich the same rows.

use crate::central_csv::{CentralDataset, CentralRecord, CentralValue};
use anyhow::{Context, Result, ensure};
use chrono::{Datelike, Duration, NaiveDate};
use excel_reader::{CellValue, SheetValues, read_value_columns};
use excel_writer::{Color, Format, FormatAlign, Workbook, Worksheet};
use regex::Regex;
use std::{
    collections::{BTreeMap, BTreeSet, HashMap},
    fs::{self, OpenOptions},
    path::{Path, PathBuf},
    sync::atomic::{AtomicU64, Ordering},
    sync::{Mutex, OnceLock},
    time::Instant,
};

const DBL_EPSILON: f64 = 2.22044604925031e-13;
const REMARK_CURRENT_ONLY: &str = "本期有，上期无";
const REMARK_PREVIOUS_ONLY: &str = "本期无，上期有";
const REMARK_DIV_ZERO: &str = "除数为0";
const REMARK_SHRINK_100: &str = "缩小100倍以上";
const S3_RESULT_COLUMN_WIDTHS: [u16; 22] = [
    10, 12, 10, 24, 10, 10, 12, 10, 32, 9, 9, 6, 6, 14, 14, 14, 12, 24, 24, 24, 40, 12,
];

static TEMP_OUTPUT_COUNTER: AtomicU64 = AtomicU64::new(0);

#[derive(Clone, Debug, Default, PartialEq)]
pub struct S3AlertBand {
    pub lower: String,
    pub upper: String,
    pub remark_text: String,
    pub fill_color: String,
    pub fill_entire_row: String,
    pub amount_threshold: String,
    pub explanation: String,
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct S3IndicatorReference {
    pub form_code: String,
    pub row_number: String,
    pub alias: String,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum S3ExpressionSchema {
    FiveSegment,
    LegacyEightSegment,
}

#[derive(Clone, Debug, PartialEq)]
pub struct S3ExpressionRule {
    pub schema: S3ExpressionSchema,
    pub fields: BTreeMap<String, String>,
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct S3ExecutionCompareConfig {
    pub source_unit: String,
    pub target_unit: String,
    pub run_parameters: BTreeMap<String, String>,
    pub exempt_indicators: BTreeSet<String>,
    pub alert_bands: Vec<S3AlertBand>,
    pub indicator_references: BTreeMap<String, S3IndicatorReference>,
    /// Raw named fields are retained to match the existing workbook contract.
    /// The execution layer can implement each action without lossy remapping.
    pub action_rules: Vec<BTreeMap<String, String>>,
    pub expression_rules: Vec<S3ExpressionRule>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct S3ExecutionCompareRow {
    pub record: CentralRecord,
    pub order_code_out: String,
    pub indicator_name_out: String,
    pub previous_value: CentralValue,
    pub change: CentralValue,
    pub ratio: CentralValue,
    pub remark: String,
    pub need_explanation: String,
    pub explanation_content: String,
    pub calculation_process: String,
    pub fill_color: Option<u8>,
    pub fill_scope: String,
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct S3ExpressionEvaluationReport {
    pub enabled_rules: usize,
    pub evaluated_rules: usize,
    pub triggered_rules: usize,
    pub execution_errors: usize,
    pub skipped_frequency_rules: Vec<String>,
    pub issues: Vec<S3ExpressionEvaluationIssue>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct S3ExpressionEvaluationIssue {
    pub rule_id: String,
    pub indicator: String,
    pub detail: String,
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct S3ActionEvaluationReport {
    pub enabled_rules: usize,
    pub applicable_rules: usize,
    pub filtered_rules: usize,
    pub evaluated_rules: usize,
    pub triggered_rules: usize,
}

#[derive(Clone, Debug, PartialEq)]
pub struct S3ExecutionCompareReport {
    pub output_path: PathBuf,
    pub rows: usize,
    pub current_records: usize,
    pub previous_records: usize,
    pub enabled_action_rules: usize,
    pub applicable_action_rules: usize,
    pub triggered_action_rules: usize,
    pub enabled_expression_rules: usize,
    pub triggered_expression_rules: usize,
    pub expression_errors: usize,
    pub config_reader: &'static str,
    pub input_reader: String,
    pub writer: &'static str,
    pub config_read_time_ms: u128,
    pub input_read_time_ms: u128,
    pub calculation_time_ms: u128,
    pub workbook_write_time_ms: u128,
    pub total_time_ms: u128,
}

#[derive(Clone, Debug)]
struct CompiledExpressionRule {
    rule_id: String,
    description: String,
    inverted: bool,
    tolerance: f64,
    soft: bool,
    display_expression: String,
    expression: String,
    variants: Vec<ExpressionVariant>,
    organization_scope: Option<Regex>,
    region_scope: Option<Regex>,
    required_frequency_keys: BTreeSet<String>,
}

#[derive(Clone, Debug)]
struct ExpressionVariant {
    segments: Vec<String>,
    organization_pattern: Option<Regex>,
    region_pattern: Option<Regex>,
}

#[derive(Clone, Debug)]
struct ExpressionReference {
    start: usize,
    end: usize,
    current_period: bool,
    segments: Vec<String>,
}

#[derive(Clone, Debug)]
struct CompiledActionRule {
    rule_id: String,
    action: String,
    codes: Vec<String>,
    data_attr: String,
    currency: String,
    frequency: String,
    batch: String,
    scene: String,
    rmb_threshold: f64,
    usd_threshold: f64,
    amplitude_threshold: f64,
    target_indicator: String,
    detail: String,
    note: String,
    cumulative: bool,
}

#[derive(Clone, Copy, Debug)]
struct ActionFacts {
    current: f64,
    previous: f64,
    change: f64,
    ratio: f64,
    factor: f64,
    is_usd: bool,
}

/// Read the source product's 3.0 common workbook and 3.1 execution workbook.
/// The function preserves rule rows as configuration values; it does not
/// silently execute an unsupported expression or action.
pub fn load_execution_compare_config(
    common_path: impl AsRef<Path>,
    feature_path: impl AsRef<Path>,
) -> Result<S3ExecutionCompareConfig> {
    let common_path = common_path.as_ref();
    let feature_path = feature_path.as_ref();
    ensure_xlsx(common_path, "3.0 大集中通用配置")?;
    ensure_xlsx(feature_path, "3.1 大集中执行比较配置")?;

    let common_sheets = read_value_columns(common_path, usize::MAX)
        .with_context(|| format!("读取 3.0 大集中通用配置：{}", common_path.display()))?;
    let feature_sheets = read_value_columns(feature_path, usize::MAX)
        .with_context(|| format!("读取 3.1 大集中执行比较配置：{}", feature_path.display()))?;

    let mut config = S3ExecutionCompareConfig {
        source_unit: "元".to_owned(),
        target_unit: "亿元".to_owned(),
        ..S3ExecutionCompareConfig::default()
    };

    let run_sheet =
        find_sheet(&common_sheets, "运行参数").context("3.0 大集中通用配置缺少“运行参数”工作表")?;
    config.run_parameters = read_parameter_rows(run_sheet)?;
    if let Some(unit) = config.run_parameters.get("默认源数据单位") {
        if !unit.trim().is_empty() {
            config.source_unit = unit.trim().to_owned();
        }
    }
    if let Some(unit) = config.run_parameters.get("默认目标单位") {
        if !unit.trim().is_empty() {
            config.target_unit = unit.trim().to_owned();
        }
    }

    let exception_sheet = find_sheet(&common_sheets, "单位换算例外")
        .context("3.0 大集中通用配置缺少“单位换算例外”工作表")?;
    for row in read_rows(exception_sheet, &["指标代码", "启用"])? {
        let code = row.get("指标代码").map(String::as_str).unwrap_or_default();
        if !code.trim().is_empty()
            && enabled(row.get("启用").map(String::as_str).unwrap_or_default())
        {
            config
                .exempt_indicators
                .insert(code.replace('\'', "").trim().to_owned());
        }
    }

    let alert_sheet =
        find_sheet(&common_sheets, "环比警戒").context("3.0 大集中通用配置缺少“环比警戒”工作表")?;
    config.alert_bands = read_rows(
        alert_sheet,
        &[
            "下限",
            "上限",
            "备注文字",
            "填充颜色",
            "是否整行填充",
            "金额变动阈值",
            "是否说明",
        ],
    )?
    .into_iter()
    .map(|row| S3AlertBand {
        lower: row["下限"].clone(),
        upper: row["上限"].clone(),
        remark_text: row["备注文字"].clone(),
        fill_color: row["填充颜色"].clone(),
        fill_entire_row: row["是否整行填充"].clone(),
        amount_threshold: row["金额变动阈值"].clone(),
        explanation: row["是否说明"].clone(),
    })
    .collect();

    let indicator_sheet = find_sheet(&feature_sheets, "指标参照")
        .context("3.1 大集中执行比较配置缺少“指标参照”工作表")?;
    for row in read_rows(
        indicator_sheet,
        &["指标代码", "表单代码", "行序号", "指标别名"],
    )? {
        let code = row["指标代码"].replace('\'', "").trim().to_owned();
        if !code.is_empty() {
            config.indicator_references.insert(
                code,
                S3IndicatorReference {
                    form_code: row["表单代码"].trim().to_owned(),
                    row_number: row["行序号"].trim().to_owned(),
                    alias: row["指标别名"].trim().to_owned(),
                },
            );
        }
    }

    let action_sheet = find_sheet(&feature_sheets, "规则动作")
        .context("3.1 大集中执行比较配置缺少“规则动作”工作表")?;
    config.action_rules = read_rows(
        action_sheet,
        &[
            "规则编号",
            "规则说明",
            "规则动作",
            "指标代码",
            "数据属性",
            "币种",
            "比较指标代码",
            "人民币阈值(亿元)",
            "美元阈值(亿美元)",
            "变幅阈值(%)",
            "来源分组",
            "适用表单",
            "适用频度",
            "适用批次",
            "适用场景",
            "详细说明",
            "启用",
            "停用原因",
            "备注",
        ],
    )?;

    let mut saw_expression_sheet = false;
    for (name, schema) in [
        ("表达式校验（兼容）", S3ExpressionSchema::LegacyEightSegment),
        ("表达式校验", S3ExpressionSchema::FiveSegment),
    ] {
        let Some(sheet) = find_sheet(&feature_sheets, name) else {
            continue;
        };
        saw_expression_sheet = true;
        let rows = read_rows(
            sheet,
            &[
                "规则编号",
                "规则说明",
                "校验表达式",
                "是否取反",
                "容差值(万元)",
                "启用",
                "停用原因",
            ],
        )?;
        config.expression_rules.extend(
            rows.into_iter()
                .map(|fields| S3ExpressionRule { schema, fields }),
        );
    }
    ensure!(
        saw_expression_sheet,
        "3.1 大集中执行比较配置缺少表达式校验工作表"
    );
    Ok(config)
}

/// Pair current and previous records using the legacy eight-dimension key.
/// Rule actions and expression enrichment are intentionally separate so this
/// base stage cannot report an unsupported rule as a successful check.
pub fn build_base_comparison(
    current: &CentralDataset,
    previous: &CentralDataset,
    config: &S3ExecutionCompareConfig,
) -> Vec<S3ExecutionCompareRow> {
    let digits = round_digits(&config.target_unit);
    let mut previous_index = BTreeMap::new();
    for record in &previous.records {
        previous_index.insert(record.key(), record);
    }
    let mut seen = BTreeSet::new();
    let mut rows = Vec::with_capacity(current.records.len() + previous.records.len());

    for record in &current.records {
        let key = record.key();
        seen.insert(key.clone());
        let previous_record = previous_index.get(&key).copied();
        let mut row = base_row(record.clone(), config);
        if let Some(previous_record) = previous_record {
            row.previous_value = previous_record.value.clone();
            let current_number = value_number(&record.value).unwrap_or(0.0);
            let previous_number = value_number(&previous_record.value).unwrap_or(0.0);
            if (current_number - previous_number).abs() > DBL_EPSILON {
                let change = excel_round(current_number - previous_number, digits);
                row.change = CentralValue::Number(change);
                if previous_number != 0.0 {
                    let ratio = change / previous_number;
                    row.ratio = CentralValue::Number(ratio);
                    apply_alert_band(&mut row, ratio, change, config);
                } else {
                    row.ratio = CentralValue::Text(REMARK_DIV_ZERO.to_owned());
                }
            }
        } else {
            row.change = value_number(&record.value)
                .map(CentralValue::Number)
                .unwrap_or(CentralValue::Empty);
            row.remark = REMARK_CURRENT_ONLY.to_owned();
            row.fill_color = Some(3);
            row.fill_scope = "row".to_owned();
        }
        rows.push(row);
    }

    for previous_record in &previous.records {
        let key = previous_record.key();
        if !seen.insert(key) {
            continue;
        }
        let mut ghost = previous_record.clone();
        ghost.record_date = if current.record_date.is_empty() {
            previous.record_date.clone()
        } else {
            current.record_date.clone()
        };
        ghost.value = CentralValue::Empty;
        let mut row = base_row(ghost, config);
        let previous_number = value_number(&previous_record.value).unwrap_or(0.0);
        row.previous_value = CentralValue::Number(previous_number);
        row.change = CentralValue::Number(-previous_number);
        row.ratio = CentralValue::Number(-1.0);
        row.remark = REMARK_PREVIOUS_ONLY.to_owned();
        row.fill_color = Some(6);
        row.fill_scope = "row".to_owned();
        rows.push(row);
    }

    rows.sort_by(|left, right| {
        left.record
            .org_code
            .cmp(&right.record.org_code)
            .then_with(|| left.record.region_code.cmp(&right.record.region_code))
            .then_with(|| left.order_code_out.cmp(&right.order_code_out))
    });
    rows
}

/// Apply one selected 3.1 expression grammar to already paired rows.
///
/// The Python implementation defaults to the legacy eight-segment sheet;
/// callers must explicitly select the five-segment sheet when that migration
/// mode is requested. Compile failures stop the selected rule set. Runtime
/// evaluation failures are recorded and conservatively treated as triggered,
/// matching the legacy audit behavior rather than silently passing a row.
pub fn apply_expression_rules(
    rows: &mut [S3ExecutionCompareRow],
    current: &CentralDataset,
    previous: &CentralDataset,
    config: &S3ExecutionCompareConfig,
    schema: S3ExpressionSchema,
) -> Result<S3ExpressionEvaluationReport> {
    let mut report = S3ExpressionEvaluationReport::default();
    let compiled = compile_expression_rules(config, schema, &current.record_date)?;
    report.enabled_rules = compiled.len();
    if compiled.is_empty() {
        return Ok(report);
    }

    if schema == S3ExpressionSchema::FiveSegment {
        validate_five_segment_key_uniqueness(current, "本期")?;
        validate_five_segment_key_uniqueness(previous, "上期")?;
    }

    let current_index = record_index(current);
    let previous_index = record_index(previous);
    let available_frequency_keys = frequency_keys(current, previous);
    let mut by_indicator = BTreeMap::<String, Vec<usize>>::new();
    for (index, rule) in compiled.iter().enumerate() {
        if !rule
            .required_frequency_keys
            .iter()
            .all(|key| available_frequency_keys.contains(key))
        {
            report.skipped_frequency_rules.push(rule.rule_id.clone());
            continue;
        }
        for variant in &rule.variants {
            if let Some(indicator) = variant.segments.get(3) {
                by_indicator
                    .entry(indicator.replace('\'', "").trim().to_owned())
                    .or_default()
                    .push(index);
            }
        }
    }

    let unit_factor = crate::s3_current_value_check::s3_target_unit_factor(&config.target_unit);
    let soft_color = config
        .run_parameters
        .get("复杂校验软性颜色")
        .and_then(|value| parse_config_int(value))
        .and_then(|value| u8::try_from(value).ok())
        .unwrap_or(46);
    let mut soft_checked = BTreeMap::<String, BTreeSet<String>>::new();

    for row in rows {
        let Some(rule_indexes) = by_indicator.get(&row.record.indicator) else {
            continue;
        };
        let mut seen = BTreeSet::new();
        for rule_index in rule_indexes {
            if !seen.insert(*rule_index) {
                continue;
            }
            let rule = &compiled[*rule_index];
            if !rule
                .variants
                .iter()
                .any(|variant| expression_variant_matches(variant, &row.record))
            {
                continue;
            }
            if rule
                .organization_scope
                .as_ref()
                .is_some_and(|pattern| !pattern.is_match(&row.record.org_code))
                || rule
                    .region_scope
                    .as_ref()
                    .is_some_and(|pattern| !pattern.is_match(&row.record.region_code))
            {
                continue;
            }

            let row_key = row.record.key();
            if rule.soft
                && soft_checked
                    .get(&rule.expression)
                    .is_some_and(|registered| registered.contains(&row_key))
            {
                continue;
            }

            report.evaluated_rules += 1;
            let mut registrations = if rule.soft {
                Some(soft_checked.entry(rule.expression.clone()).or_default())
            } else {
                None
            };
            let substituted = substitute_expression_references(
                &rule.expression,
                &row.record,
                &current_index,
                &previous_index,
                &config.exempt_indicators,
                unit_factor,
                registrations.as_deref_mut(),
            )
            .replace(
                "Thd",
                &crate::s3_current_value_check::s3_python_float_repr(rule.tolerance),
            )
            .replace(
                "thd",
                &crate::s3_current_value_check::s3_python_float_repr(rule.tolerance),
            )
            .replace(' ', "")
            .to_lowercase();
            if let Some(registered) = registrations.as_deref_mut() {
                registered.insert(row_key);
            }

            let (hit, failure_reason) =
                match crate::s3_current_value_check::evaluate_s3_business_expression(
                    &substituted,
                    rule.tolerance,
                ) {
                    Ok(value) => match value.as_boolean() {
                        Some(boolean) => (boolean != rule.inverted, None),
                        None => (
                            true,
                            Some(format!(
                                "结果并不是逻辑值:{}",
                                python_cstr(value.as_number())
                            )),
                        ),
                    },
                    Err(error) => (true, Some(error)),
                };

            if let Some(reason) = failure_reason.as_ref() {
                report.execution_errors += 1;
                report.issues.push(S3ExpressionEvaluationIssue {
                    rule_id: rule.rule_id.clone(),
                    indicator: row.record.indicator.clone(),
                    detail: reason.clone(),
                });
            }
            if !hit {
                continue;
            }
            report.triggered_rules += 1;
            append_unique(&mut row.need_explanation, &rule.description);
            let mut process = format!(
                "计算过程：【{}】->【{}】",
                rule.display_expression, substituted
            );
            if let Some(reason) = failure_reason {
                let label = if reason.starts_with("结果并不是逻辑值:") {
                    "校验"
                } else {
                    "校验时发生错误"
                };
                process = format!(
                    " 计算【{}】{}--{}。{}",
                    rule.description, label, reason, process
                );
            }
            process = process.replace("左中", "[").replace("右中", "]");
            append_joined(&mut row.calculation_process, &process);
            if rule.soft && row.fill_scope != "row" {
                row.fill_color = Some(soft_color);
                row.fill_scope = "remark".to_owned();
            }
        }
    }
    report.skipped_frequency_rules.sort();
    report.skipped_frequency_rules.dedup();
    Ok(report)
}

/// Apply the 3.1 V3 action table to paired comparison rows.
/// Enabled actions that cannot be represented by the current Rust dispatcher
/// stop before any row is reported as clean.
pub fn apply_action_rules(
    rows: &mut [S3ExecutionCompareRow],
    current: &CentralDataset,
    previous: &CentralDataset,
    config: &S3ExecutionCompareConfig,
) -> Result<S3ActionEvaluationReport> {
    let (rules, enabled_rules) = prepare_action_rules(config, current, previous)?;
    let mut report = S3ActionEvaluationReport {
        enabled_rules,
        applicable_rules: rules.len(),
        filtered_rules: enabled_rules.saturating_sub(rules.len()),
        ..S3ActionEvaluationReport::default()
    };
    if rules.is_empty() {
        return Ok(report);
    }

    let mut special_by_code = BTreeMap::<String, Vec<usize>>::new();
    let mut cumulative_by_key = BTreeMap::<(String, String, String), Vec<usize>>::new();
    let mut reverse_correspondence = BTreeMap::<String, Vec<usize>>::new();
    for (index, rule) in rules.iter().enumerate() {
        if rule.cumulative {
            if let Some(code) = rule.codes.first() {
                cumulative_by_key
                    .entry((code.clone(), rule.data_attr.clone(), rule.frequency.clone()))
                    .or_default()
                    .push(index);
            }
        } else {
            for code in &rule.codes {
                special_by_code.entry(code.clone()).or_default().push(index);
            }
            if rule.action == "指标需要对应存在。" && !rule.target_indicator.is_empty() {
                reverse_correspondence
                    .entry(rule.target_indicator.clone())
                    .or_default()
                    .push(index);
            }
        }
    }

    let current_index = record_index(current);
    let previous_index = record_index(previous);
    let unit = crate::s3_current_value_check::s3_target_unit_factor(&config.target_unit)
        / crate::s3_current_value_check::s3_target_unit_factor("亿元");
    for row in rows {
        let current_number = number_or_zero(&row.record.value);
        let previous_number = number_or_zero(&row.previous_value);
        let change = number_or_zero(&row.change);
        let ratio = number_or_zero(&row.ratio);
        let factor = if config.exempt_indicators.contains(&row.record.indicator) {
            1.0
        } else {
            unit
        };
        let facts = ActionFacts {
            current: current_number,
            previous: previous_number,
            change,
            ratio,
            factor,
            is_usd: row.record.currency == "美元合计",
        };
        let mut messages = Vec::new();
        let cumulative_key = (
            row.record.indicator.clone(),
            row.record.data_attr.clone(),
            row.record.frequency.clone(),
        );
        if let Some(indices) = cumulative_by_key.get(&cumulative_key) {
            let message = cumulative_action_message(
                row,
                current.record_date.as_str(),
                previous.record_date.as_str(),
            );
            if !message.is_empty() {
                for index in indices {
                    let rule = &rules[*index];
                    report.evaluated_rules += 1;
                    report.triggered_rules += 1;
                    messages.push(format_action_message(rule, &message, &facts));
                }
            }
        }

        if let Some(indices) = special_by_code.get(&row.record.indicator) {
            for index in indices {
                let rule = &rules[*index];
                if !action_rule_applies(rule, &row.record, current.record_date.as_str()) {
                    continue;
                }
                report.evaluated_rules += 1;
                let hit = execute_special_action(
                    rule,
                    row,
                    &facts,
                    &current_index,
                    &previous_index,
                    &current.record_date,
                    &previous.record_date,
                    config,
                );
                if !hit.is_empty() {
                    report.triggered_rules += 1;
                    messages.push(format_action_message(rule, &hit, &facts));
                }
            }
        }

        if facts.current != 0.0 {
            if let Some(indices) = reverse_correspondence.get(&row.record.indicator) {
                for index in indices {
                    let rule = &rules[*index];
                    if !action_rule_applies(rule, &row.record, current.record_date.as_str()) {
                        continue;
                    }
                    for source in &rule.codes {
                        if source == &row.record.indicator {
                            continue;
                        }
                        let source_key = record_key_with_indicator(&row.record, source);
                        if current_index.contains_key(&source_key) {
                            continue;
                        }
                        report.evaluated_rules += 1;
                        report.triggered_rules += 1;
                        messages.push(format_action_message(
                            rule,
                            &format!("{}有数,{source}无数,请核实。", row.record.indicator),
                            &facts,
                        ));
                    }
                }
            }
        }

        if messages.is_empty() {
            continue;
        }
        for message in &messages {
            append_unique(&mut row.need_explanation, message);
        }
        if messages.iter().all(|message| message.contains("软性")) && row.fill_scope != "row" {
            let color = config
                .run_parameters
                .get("复杂校验软性颜色")
                .and_then(|value| parse_config_int(value))
                .and_then(|value| u8::try_from(value).ok())
                .unwrap_or(46);
            row.fill_color = Some(color);
            row.fill_scope = "remark".to_owned();
        }
    }
    Ok(report)
}

/// Execute the complete currently migrated portion of S3-F01 from business
/// inputs through a new Python-contract comparison workbook. Only the chosen
/// five/eight-segment expression sheet is applied, matching the legacy mode
/// selector; the default CLI mode remains eight-segment.
pub fn run_execution_compare_files(
    current_paths: &[PathBuf],
    previous_paths: &[PathBuf],
    common_config_path: &Path,
    feature_config_path: &Path,
    output_path: &Path,
    schema: S3ExpressionSchema,
    mut on_step: impl FnMut(&str),
) -> Result<S3ExecutionCompareReport> {
    let started = Instant::now();
    ensure!(
        !current_paths.is_empty(),
        "S3-F01 执行比较：请先选择本期数据"
    );
    ensure!(
        !previous_paths.is_empty(),
        "S3-F01 执行比较：请先选择上期数据"
    );
    let output_path =
        crate::s3_current_value_check::resolve_new_xlsx_path(output_path, "S3-F01 比较结果")?;
    let mut all_input_paths = current_paths.to_vec();
    all_input_paths.extend_from_slice(previous_paths);
    let input_reader = crate::s3_current_value_check::input_reader_for_paths(&all_input_paths);

    on_step("读取 3.0/3.1 配置");
    let config_started = Instant::now();
    let config = load_execution_compare_config(common_config_path, feature_config_path)?;
    let config_read_time_ms = config_started.elapsed().as_millis();
    let conversion_factor =
        crate::s3_current_value_check::s3_target_unit_factor(&config.source_unit)
            / crate::s3_current_value_check::s3_target_unit_factor(&config.target_unit);

    on_step("读取本期/上期导数");
    let input_started = Instant::now();
    let current = crate::s3_current_value_check::load_central_dataset(
        current_paths,
        conversion_factor,
        &config.exempt_indicators,
    )
    .context("读取 S3-F01 本期数据失败")?;
    let previous = crate::s3_current_value_check::load_central_dataset(
        previous_paths,
        conversion_factor,
        &config.exempt_indicators,
    )
    .context("读取 S3-F01 上期数据失败")?;
    ensure!(
        current.issues.is_empty() && previous.issues.is_empty(),
        "S3-F01 数据导入存在 {} 项本期问题、{} 项上期问题，已停止生成比较结果；请先检查导入错误",
        current.issues.len(),
        previous.issues.len()
    );
    let current_records = current.records.len();
    let previous_records = previous.records.len();
    let input_read_time_ms = input_started.elapsed().as_millis();

    on_step("执行八维配对、环比和规则审核");
    let calculation_started = Instant::now();
    let mut rows = build_base_comparison(&current, &previous, &config);
    let expression_report =
        apply_expression_rules(&mut rows, &current, &previous, &config, schema)?;
    let action_report = apply_action_rules(&mut rows, &current, &previous, &config)?;
    let calculation_time_ms = calculation_started.elapsed().as_millis();

    on_step("写入并验证比较结果工作簿");
    let workbook_started = Instant::now();
    write_execution_compare_workbook(
        &output_path,
        &rows,
        &config.target_unit,
        &current.record_date,
    )?;
    let workbook_write_time_ms = workbook_started.elapsed().as_millis();
    on_step("比较结果已生成");

    Ok(S3ExecutionCompareReport {
        output_path,
        rows: rows.len(),
        current_records,
        previous_records,
        enabled_action_rules: action_report.enabled_rules,
        applicable_action_rules: action_report.applicable_rules,
        triggered_action_rules: action_report.triggered_rules,
        enabled_expression_rules: expression_report.enabled_rules,
        triggered_expression_rules: expression_report.triggered_rules,
        expression_errors: expression_report.execution_errors,
        config_reader: "calamine",
        input_reader,
        writer: "rust_xlsxwriter",
        config_read_time_ms,
        input_read_time_ms,
        calculation_time_ms,
        workbook_write_time_ms,
        total_time_ms: started.elapsed().as_millis(),
    })
}

const S3_RESULT_HEADERS: [&str; 21] = [
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
    "上期值",
    "增减额",
    "环比",
    "备注",
    "是否说明",
    "说明内容",
    "计算过程",
];

fn write_execution_compare_workbook(
    output_path: &Path,
    rows: &[S3ExecutionCompareRow],
    target_unit: &str,
    record_date: &str,
) -> Result<()> {
    ensure!(
        rows.len() <= 1_048_575,
        "S3-F01 比较结果超过 Excel 的最大数据行数"
    );
    ensure!(
        output_path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx")),
        "S3-F01 比较结果必须输出为 .xlsx"
    );
    ensure!(
        !output_path.exists(),
        "拒绝覆盖已有 S3-F01 比较结果：{}",
        output_path.display()
    );
    let parent = output_path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or(Path::new("."));
    ensure!(
        parent.is_dir(),
        "S3-F01 比较结果目录不存在：{}",
        parent.display()
    );

    let sheet_name = execution_result_sheet_name(record_date)?;
    let headers = execution_result_headers(target_unit);
    let temporary_path = reserve_execution_result_temporary_path(output_path)?;
    let write_result = (|| -> Result<()> {
        {
            let mut workbook = Workbook::new();
            let worksheet = workbook.add_worksheet();
            worksheet
                .set_name(&sheet_name)
                .context("设置 S3-F01 比较结果工作表名称失败")?;

            let header_format = Format::new()
                .set_bold()
                .set_font_color(Color::RGB(0xFF_FF_FF))
                .set_background_color(Color::RGB(0x1F_4E_78))
                .set_align(FormatAlign::VerticalCenter)
                .set_text_wrap();

            for (column, header) in headers.iter().enumerate() {
                worksheet
                    .write_string_with_format(0, column as u16, header, &header_format)
                    .with_context(|| format!("写入 S3-F01 比较结果表头第 {} 列", column + 1))?;
            }

            for (index, row) in rows.iter().enumerate() {
                let output_row = index as u32 + 1;
                let record = &row.record;
                let fill_color = row
                    .fill_color
                    .filter(|_| !row.fill_scope.is_empty())
                    .and_then(vba_color_index_rgb);
                let text_values = [
                    (0, record.biz_class.as_str()),
                    (1, record.record_date.as_str()),
                    (2, record.org_code.as_str()),
                    (3, record.org_name.as_str()),
                    (4, record.region_code.as_str()),
                    (5, record.region_name.as_str()),
                    (
                        6,
                        if row.order_code_out.is_empty() {
                            record.indicator.as_str()
                        } else {
                            row.order_code_out.as_str()
                        },
                    ),
                    (7, record.indicator.as_str()),
                    (
                        8,
                        if row.indicator_name_out.is_empty() {
                            record.indicator_name.as_str()
                        } else {
                            row.indicator_name_out.as_str()
                        },
                    ),
                    (9, record.data_attr.as_str()),
                    (10, record.currency.as_str()),
                    (11, record.frequency.as_str()),
                    (12, record.batch.as_str()),
                    (17, row.remark.as_str()),
                    (18, row.need_explanation.as_str()),
                    (19, row.explanation_content.as_str()),
                    (20, row.calculation_process.as_str()),
                ];
                for (column, value) in text_values {
                    let applies_fill = fill_color.is_some()
                        && ((row.fill_scope == "row" && column <= 17)
                            || (row.fill_scope != "row" && column == 17));
                    if column == 1 || applies_fill {
                        let mut format = Format::new();
                        if column == 1 {
                            format = format.set_num_format("yyyy/m/d");
                        }
                        if applies_fill {
                            format = format.set_background_color(Color::RGB(fill_color.unwrap()));
                        }
                        worksheet.write_string_with_format(output_row, column, value, &format)?;
                    } else {
                        worksheet.write_string(output_row, column, value)?;
                    }
                }
                write_central_value(
                    worksheet,
                    output_row,
                    13,
                    &record.value,
                    Some("0.00"),
                    fill_color.filter(|_| row.fill_scope == "row"),
                )?;
                write_central_value(
                    worksheet,
                    output_row,
                    14,
                    &row.previous_value,
                    Some("0.00"),
                    fill_color.filter(|_| row.fill_scope == "row"),
                )?;
                write_central_value(
                    worksheet,
                    output_row,
                    15,
                    &row.change,
                    Some("0.00"),
                    fill_color.filter(|_| row.fill_scope == "row"),
                )?;
                write_central_value(
                    worksheet,
                    output_row,
                    16,
                    &row.ratio,
                    Some("0.00%"),
                    fill_color.filter(|_| row.fill_scope == "row"),
                )?;
                worksheet.write_blank(output_row, 21, &Format::new())?;
            }

            worksheet
                .set_freeze_panes(1, 0)
                .context("冻结 S3-F01 比较结果表头失败")?;
            worksheet
                .autofilter(0, 0, rows.len().max(1) as u32, 21)
                .context("设置 S3-F01 比较结果筛选范围失败")?;
            for (column, width) in S3_RESULT_COLUMN_WIDTHS.iter().enumerate() {
                worksheet
                    .set_column_width_pixels(column as u16, u32::from(*width) * 7)
                    .with_context(|| format!("设置 S3-F01 比较结果第 {} 列宽度", column + 1))?;
            }
            workbook
                .save(&temporary_path)
                .context("保存 S3-F01 比较结果临时工作簿失败")?;
        }

        validate_execution_compare_workbook(&temporary_path, &sheet_name, &headers, rows)?;
        fs::hard_link(&temporary_path, output_path)
            .with_context(|| format!("发布 S3-F01 比较结果失败：{}", output_path.display()))?;
        Ok(())
    })();
    if write_result.is_err() {
        let _ = fs::remove_file(&temporary_path);
    } else {
        fs::remove_file(&temporary_path)
            .with_context(|| format!("清理 S3-F01 临时工作簿：{}", temporary_path.display()))?;
    }
    write_result
}

fn execution_result_headers(target_unit: &str) -> Vec<String> {
    let mut headers = S3_RESULT_HEADERS
        .iter()
        .map(|header| (*header).to_owned())
        .collect::<Vec<_>>();
    headers[15] = format!("增减额({target_unit})");
    headers.push("输出说明".to_owned());
    headers
}

fn execution_result_sheet_name(record_date: &str) -> Result<String> {
    let Some((year, month)) = record_date.split_once('-') else {
        return Ok("对比结果".to_owned());
    };
    let month = month
        .split('-')
        .next()
        .unwrap_or_default()
        .parse::<u32>()
        .with_context(|| format!("S3-F01 数据日期格式无法生成工作表名称：{record_date}"))?;
    Ok(format!("{year}年{month}月对比结果"))
}

fn write_central_value(
    worksheet: &mut Worksheet,
    row: u32,
    column: u16,
    value: &CentralValue,
    number_format: Option<&str>,
    fill_color: Option<u32>,
) -> Result<()> {
    match value {
        CentralValue::Number(value) => {
            ensure!(value.is_finite(), "S3-F01 比较结果包含非有限数值");
            if let Some(format) = number_format {
                let mut cell_format = Format::new().set_num_format(format);
                if let Some(color) = fill_color {
                    cell_format = cell_format.set_background_color(Color::RGB(color));
                }
                worksheet.write_number_with_format(row, column, *value, &cell_format)?;
            } else if let Some(color) = fill_color {
                worksheet.write_number_with_format(
                    row,
                    column,
                    *value,
                    &Format::new().set_background_color(Color::RGB(color)),
                )?;
            } else {
                worksheet.write_number(row, column, *value)?;
            }
        }
        CentralValue::Text(value) => {
            let mut format = Format::new();
            if let Some(number_format) = number_format {
                format = format.set_num_format(number_format);
            }
            if let Some(color) = fill_color {
                format = format.set_background_color(Color::RGB(color));
            }
            if number_format.is_some() || fill_color.is_some() {
                worksheet.write_string_with_format(row, column, value, &format)?;
            } else {
                worksheet.write_string(row, column, value)?;
            }
        }
        CentralValue::Empty => {
            let mut format = Format::new();
            if let Some(number_format) = number_format {
                format = format.set_num_format(number_format);
            }
            if let Some(color) = fill_color {
                format = format.set_background_color(Color::RGB(color));
            }
            worksheet.write_blank(row, column, &format)?;
        }
    }
    Ok(())
}

fn vba_color_index_rgb(index: u8) -> Option<u32> {
    const PALETTE: [u32; 55] = [
        0x00_00_00, 0xFF_FF_FF, 0xFF_00_00, 0x00_FF_00, 0x00_00_FF, 0xFF_FF_00, 0xFF_00_FF,
        0x00_FF_FF, 0x80_00_00, 0x00_80_00, 0x00_00_80, 0x80_80_00, 0x80_00_80, 0x00_80_80,
        0xC0_C0_C0, 0x80_80_80, 0x99_99_FF, 0x99_33_66, 0xFF_FF_CC, 0xCC_FF_FF, 0x66_00_66,
        0xFF_80_80, 0x00_66_CC, 0xCC_CC_FF, 0x00_00_80, 0xFF_00_FF, 0xFF_FF_00, 0x00_FF_FF,
        0x80_00_80, 0x80_00_80, 0x00_80_80, 0x00_00_FF, 0x00_CC_FF, 0xCC_FF_FF, 0xCC_FF_CC,
        0xFF_FF_99, 0x99_CC_FF, 0xFF_99_CC, 0xCC_99_FF, 0xFF_CC_99, 0x33_66_FF, 0x33_CC_CC,
        0x99_CC_00, 0xFF_CC_00, 0xFF_99_00, 0xFF_66_00, 0x80_80_80, 0x96_96_96, 0x80_80_00,
        0x00_CC_99, 0x00_33_99, 0x66_66_99, 0x80_80_80, 0x99_33_66, 0x33_33_00,
    ];
    index
        .checked_sub(1)
        .and_then(|index| PALETTE.get(index as usize).copied())
}

fn reserve_execution_result_temporary_path(output_path: &Path) -> Result<PathBuf> {
    let parent = output_path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or(Path::new("."));
    let stem = output_path
        .file_stem()
        .and_then(|stem| stem.to_str())
        .unwrap_or("s3-comparison");
    for _ in 0..100 {
        let counter = TEMP_OUTPUT_COUNTER.fetch_add(1, Ordering::Relaxed);
        let temporary = parent.join(format!(
            ".{stem}.{}.{}.tmp.xlsx",
            std::process::id(),
            counter
        ));
        match OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&temporary)
        {
            Ok(file) => {
                drop(file);
                return Ok(temporary);
            }
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(error) => {
                return Err(error).with_context(|| {
                    format!("创建 S3-F01 临时工作簿路径失败：{}", temporary.display())
                });
            }
        }
    }
    anyhow::bail!("无法为 S3-F01 比较结果分配唯一临时路径")
}

fn validate_execution_compare_workbook(
    path: &Path,
    expected_sheet_name: &str,
    headers: &[String],
    rows: &[S3ExecutionCompareRow],
) -> Result<()> {
    let sheets = read_value_columns(path, 22).context("独立读回 S3-F01 比较结果失败")?;
    ensure!(sheets.len() == 1, "S3-F01 比较结果应只有一个工作表");
    let sheet = &sheets[0];
    ensure!(
        sheet.name == expected_sheet_name,
        "S3-F01 比较结果工作表名称不匹配"
    );
    ensure!(
        sheet.max_row == rows.len() + 1 && sheet.max_column == 22,
        "S3-F01 比较结果行列数不匹配：实际 {} 行、{} 列",
        sheet.max_row,
        sheet.max_column
    );
    for (column, header) in headers.iter().enumerate() {
        ensure!(
            cell_text(sheet, 1, column + 1) == *header,
            "S3-F01 比较结果第 {} 列表头读回不匹配",
            column + 1
        );
    }
    for (index, row) in rows.iter().enumerate() {
        let output_row = index + 2;
        let record = &row.record;
        let text_values = [
            (1, record.biz_class.as_str()),
            (2, record.record_date.as_str()),
            (3, record.org_code.as_str()),
            (4, record.org_name.as_str()),
            (5, record.region_code.as_str()),
            (6, record.region_name.as_str()),
            (
                7,
                if row.order_code_out.is_empty() {
                    record.indicator.as_str()
                } else {
                    row.order_code_out.as_str()
                },
            ),
            (8, record.indicator.as_str()),
            (
                9,
                if row.indicator_name_out.is_empty() {
                    record.indicator_name.as_str()
                } else {
                    row.indicator_name_out.as_str()
                },
            ),
            (10, record.data_attr.as_str()),
            (11, record.currency.as_str()),
            (12, record.frequency.as_str()),
            (13, record.batch.as_str()),
            (18, row.remark.as_str()),
            (19, row.need_explanation.as_str()),
            (20, row.explanation_content.as_str()),
            (21, row.calculation_process.as_str()),
        ];
        for (column, expected) in text_values {
            ensure!(
                cell_text(sheet, output_row, column) == expected,
                "S3-F01 比较结果第 {} 行第 {} 列文本读回不匹配",
                output_row,
                column
            );
        }
        for (column, expected) in [
            (14, &record.value),
            (15, &row.previous_value),
            (16, &row.change),
            (17, &row.ratio),
        ] {
            let actual = sheet
                .values
                .get(&format!("{}{}", column_name(column), output_row))
                .unwrap_or(&CellValue::Empty);
            ensure!(
                central_value_matches_cell(expected, actual),
                "S3-F01 比较结果第 {} 行第 {} 列值读回不匹配",
                output_row,
                column
            );
        }
        ensure!(
            sheet
                .values
                .get(&format!("V{output_row}"))
                .unwrap_or(&CellValue::Empty)
                == &CellValue::Empty,
            "S3-F01 比较结果第 {} 行“输出说明”应为空",
            output_row
        );
    }
    let workbook = excel_reader::read_workbook(path).context("校验 S3-F01 结果公式元数据失败")?;
    ensure!(
        workbook.sheets.len() == 1 && workbook.sheets[0].formulas.is_empty(),
        "S3-F01 比较结果不应包含公式"
    );
    Ok(())
}

fn central_value_matches_cell(expected: &CentralValue, actual: &CellValue) -> bool {
    match (expected, actual) {
        (CentralValue::Empty, CellValue::Empty) => true,
        (CentralValue::Text(expected), CellValue::Text(actual)) => expected == actual,
        (CentralValue::Number(expected), CellValue::Number(actual)) => {
            (expected - actual).abs() <= f64::EPSILON * expected.abs().max(1.0)
        }
        (CentralValue::Number(expected), CellValue::Integer(actual)) => {
            (*expected - *actual as f64).abs() <= f64::EPSILON * expected.abs().max(1.0)
        }
        _ => false,
    }
}

fn compile_expression_rules(
    config: &S3ExecutionCompareConfig,
    schema: S3ExpressionSchema,
    current_date: &str,
) -> Result<Vec<CompiledExpressionRule>> {
    let legacy_bindings = config
        .expression_rules
        .iter()
        .filter(|rule| rule.schema == S3ExpressionSchema::LegacyEightSegment)
        .map(|rule| {
            let fields = &rule.fields;
            let id = field(fields, &["规则编号"]);
            let expression =
                normalized_expression(field(fields, &["校验表达式", "校验规则", "表达式"]));
            let binding = expression_references(&expression)
                .into_iter()
                .map(|reference| reference.segments.first().cloned().unwrap_or_default())
                .collect::<Vec<_>>();
            (id.to_owned(), binding)
        })
        .filter(|(id, _)| !id.trim().is_empty())
        .collect::<BTreeMap<_, _>>();

    let mut result = Vec::new();
    for source_rule in config
        .expression_rules
        .iter()
        .filter(|rule| rule.schema == schema)
    {
        let fields = &source_rule.fields;
        if !enabled(field(fields, &["启用"])) {
            continue;
        }
        let rule_id = field(fields, &["规则编号"]).trim();
        ensure!(!rule_id.is_empty(), "3.1 启用的表达式规则缺少规则编号");
        let display_expression = field(fields, &["校验表达式", "校验规则", "表达式"]);
        let normalized = normalized_expression(display_expression);
        ensure!(
            !normalized.trim().is_empty(),
            "3.1 表达式规则 {rule_id} 的校验表达式为空"
        );
        let description = field(fields, &["规则说明", "校验描述", "校验说明"])
            .trim()
            .to_owned();
        let tolerance_text = field(fields, &["容差值(万元)", "Thd值(万元)"]);
        let tolerance = if tolerance_text.trim().is_empty() {
            0.01
        } else {
            parse_config_number(tolerance_text).with_context(|| {
                format!("3.1 表达式规则 {rule_id} 的容差不是数值：{tolerance_text}")
            })? * 10_000.0
        };
        let source_group = field(fields, &["来源分组", "来源"]).trim();
        if (source_group == "结转") != current_date.ends_with("-01-01") {
            continue;
        }

        let inverted_text = field(fields, &["是否取反", "取反标识"]);
        let inverted = matches!(
            inverted_text.trim().to_ascii_lowercase().as_str(),
            "是" | "1" | "true" | "y"
        );
        let (expression, variants, organization_scope, region_scope) = match schema {
            S3ExpressionSchema::LegacyEightSegment => {
                let references = expression_references(&normalized);
                validate_reference_shapes(&references, 8, rule_id, true)?;
                let variants = compile_variants(references, rule_id, true)?;
                (normalized.clone(), variants, None, None)
            }
            S3ExpressionSchema::FiveSegment => {
                let bindings = legacy_bindings
                    .get(rule_id)
                    .with_context(|| format!("五段式规则 {rule_id} 缺少同编号八段式绑定规则"))?;
                let references = expression_references(&normalized);
                ensure!(
                    bindings.len() == references.len(),
                    "五段式规则 {rule_id} 与八段式绑定 token 数不一致：{} 对 {}",
                    references.len(),
                    bindings.len()
                );
                let expanded =
                    expand_five_segment_expression(&normalized, &references, bindings)
                        .with_context(|| format!("五段式规则 {rule_id} 的 token 格式无效"))?;
                let expanded_references = expression_references(&expanded);
                validate_reference_shapes(&expanded_references, 8, rule_id, false)?;
                let variants = compile_variants(expanded_references, rule_id, false)?;
                let organization_scope =
                    compile_scope(field(fields, &["机构类代码"]), rule_id, "机构类代码")?;
                let region_scope =
                    compile_scope(field(fields, &["地区代码"]), rule_id, "地区代码")?;
                (expanded, variants, organization_scope, region_scope)
            }
        };
        ensure!(
            variants.iter().any(|variant| !variant.segments.is_empty()),
            "3.1 表达式规则 {rule_id} 缺少可挂接的本期指标 token"
        );
        let required_frequency_keys = expression_references(&expression)
            .iter()
            .map(|reference| format!("{}{}", reference.segments[6], reference.segments[7]))
            .collect();
        let soft = description.contains("软性");
        result.push(CompiledExpressionRule {
            rule_id: rule_id.to_owned(),
            description,
            inverted,
            tolerance,
            soft,
            display_expression: normalized,
            expression,
            variants,
            organization_scope,
            region_scope,
            required_frequency_keys,
        });
    }
    Ok(result)
}

fn compile_variants(
    references: Vec<ExpressionReference>,
    rule_id: &str,
    legacy_regex: bool,
) -> Result<Vec<ExpressionVariant>> {
    let mut variants = Vec::new();
    for reference in references
        .into_iter()
        .filter(|reference| reference.current_period)
    {
        ensure!(
            reference.segments.len() == 8,
            "3.1 表达式规则 {rule_id} 存在非八段 token：{} 段",
            reference.segments.len()
        );
        ensure!(
            !reference.segments[3].replace('\'', "").trim().is_empty(),
            "3.1 表达式规则 {rule_id} 的本期 token 缺少指标代码"
        );
        let organization_text = if legacy_regex {
            reference.segments[1]
                .replace("左中", "[")
                .replace("右中", "]")
        } else {
            reference.segments[1].clone()
        };
        let region_text = if legacy_regex {
            reference.segments[2]
                .replace("左中", "[")
                .replace("右中", "]")
        } else {
            reference.segments[2].clone()
        };
        let organization_pattern = compile_scope(&organization_text, rule_id, "token 机构类")?;
        let region_pattern = compile_scope(&region_text, rule_id, "token 地区")?;
        let variant = ExpressionVariant {
            segments: reference.segments,
            organization_pattern,
            region_pattern,
        };
        if !variants
            .iter()
            .any(|existing: &ExpressionVariant| existing.segments == variant.segments)
        {
            variants.push(variant);
        }
    }
    ensure!(
        !variants.is_empty(),
        "3.1 表达式规则 {rule_id} 没有本期方括号 token"
    );
    Ok(variants)
}

fn validate_reference_shapes(
    references: &[ExpressionReference],
    expected_segments: usize,
    rule_id: &str,
    legacy_regex: bool,
) -> Result<()> {
    for reference in references {
        ensure!(
            reference.segments.len() == expected_segments,
            "3.1 表达式规则 {rule_id} 的 token 必须为 {expected_segments} 段，实际为 {} 段",
            reference.segments.len()
        );
        ensure!(
            !reference.segments[3].replace('\'', "").trim().is_empty(),
            "3.1 表达式规则 {rule_id} 的 token 缺少指标代码"
        );
        for (position, field_name) in [(1, "机构类"), (2, "地区")] {
            let pattern = if legacy_regex {
                reference.segments[position]
                    .replace("左中", "[")
                    .replace("右中", "]")
            } else {
                reference.segments[position].clone()
            };
            let _ = compile_scope(&pattern, rule_id, &format!("token {field_name}"))?;
        }
    }
    Ok(())
}

fn compile_scope(pattern: &str, rule_id: &str, field_name: &str) -> Result<Option<Regex>> {
    let pattern = pattern.trim();
    if pattern.is_empty() {
        return Ok(None);
    }
    static CACHE: OnceLock<Mutex<HashMap<String, std::result::Result<Regex, String>>>> =
        OnceLock::new();
    let cache = CACHE.get_or_init(|| Mutex::new(HashMap::new()));
    let compiled = cache
        .lock()
        .expect("expression regex cache lock poisoned")
        .entry(pattern.to_owned())
        .or_insert_with(|| Regex::new(pattern).map_err(|error| error.to_string()))
        .clone();
    compiled
        .map(Some)
        .map_err(anyhow::Error::msg)
        .with_context(|| format!("3.1 表达式规则 {rule_id} 的{field_name}正则无效：{pattern}"))
}

fn expand_five_segment_expression(
    expression: &str,
    references: &[ExpressionReference],
    bindings: &[String],
) -> Result<String> {
    ensure!(
        references.len() == bindings.len(),
        "五段式 token 绑定数量不一致"
    );
    let mut output = String::with_capacity(expression.len() + references.len() * 16);
    let mut cursor = 0;
    for (reference, business_class) in references.iter().zip(bindings) {
        ensure!(
            reference.segments.len() == 5,
            "五段式 token 必须恰好包含五段"
        );
        output.push_str(&expression[cursor..reference.start]);
        let closing = if reference.current_period { ']' } else { '}' };
        output.push(if reference.current_period { '[' } else { '{' });
        output.push_str(business_class);
        output.push_str(",,");
        output.push(',');
        output.push_str(&reference.segments.join(","));
        output.push(closing);
        cursor = reference.end;
    }
    output.push_str(&expression[cursor..]);
    Ok(output)
}

fn expression_references(expression: &str) -> Vec<ExpressionReference> {
    static PATTERN: OnceLock<Regex> = OnceLock::new();
    let pattern = PATTERN
        .get_or_init(|| Regex::new(r"\[[^\[\]]*\]|\{[^{}]*\}").expect("static reference regex"));
    pattern
        .find_iter(expression)
        .map(|matched| {
            let token = matched.as_str();
            ExpressionReference {
                start: matched.start(),
                end: matched.end(),
                current_period: token.starts_with('['),
                segments: token[1..token.len() - 1]
                    .split(',')
                    .map(|segment| segment.trim().to_owned())
                    .collect(),
            }
        })
        .collect()
}

fn normalized_expression(expression: &str) -> String {
    expression
        .replace('【', "[")
        .replace('】', "]")
        .replace('（', "(")
        .replace('）', ")")
        .replace('｛', "{")
        .replace('｝', "}")
        .replace('，', ",")
        .replace('\u{3000}', " ")
}

fn field<'a>(fields: &'a BTreeMap<String, String>, names: &[&str]) -> &'a str {
    names
        .iter()
        .find_map(|name| fields.get(*name).map(String::as_str))
        .unwrap_or_default()
}

fn expression_variant_matches(variant: &ExpressionVariant, record: &CentralRecord) -> bool {
    let actual = [
        &record.biz_class,
        &record.org_code,
        &record.region_code,
        &record.indicator,
        &record.data_attr,
        &record.currency,
        &record.frequency,
        &record.batch,
    ];
    for position in [0, 4, 5, 6, 7] {
        let expected = variant.segments[position].trim();
        if !expected.is_empty() && expected != actual[position] {
            return false;
        }
    }
    variant
        .organization_pattern
        .as_ref()
        .is_none_or(|pattern| pattern.is_match(&record.org_code))
        && variant
            .region_pattern
            .as_ref()
            .is_none_or(|pattern| pattern.is_match(&record.region_code))
}

fn validate_five_segment_key_uniqueness(dataset: &CentralDataset, period: &str) -> Result<()> {
    let mut keys =
        BTreeMap::<(String, String, String, String, String, String, String), usize>::new();
    for record in &dataset.records {
        let key = (
            record.org_code.clone(),
            record.region_code.clone(),
            record.indicator.clone(),
            record.data_attr.clone(),
            record.currency.clone(),
            record.frequency.clone(),
            record.batch.clone(),
        );
        *keys.entry(key).or_default() += 1;
    }
    let conflicts = keys
        .iter()
        .filter(|(_, count)| **count > 1)
        .collect::<Vec<_>>();
    ensure!(
        conflicts.is_empty(),
        "五段式取数键不唯一（{period}共 {} 组），无法确定应使用哪条数据；首组指标 {}（机构 {}，地区 {}）出现 {} 条",
        conflicts.len(),
        conflicts[0].0.2,
        conflicts[0].0.0,
        conflicts[0].0.1,
        conflicts[0].1
    );
    Ok(())
}

fn record_index(dataset: &CentralDataset) -> BTreeMap<String, &CentralRecord> {
    let mut index = BTreeMap::new();
    for record in &dataset.records {
        index.insert(record.key(), record);
    }
    index
}

fn prepare_action_rules(
    config: &S3ExecutionCompareConfig,
    current: &CentralDataset,
    previous: &CentralDataset,
) -> Result<(Vec<CompiledActionRule>, usize)> {
    let available_frequency_keys = frequency_keys(current, previous);
    let january_first = current.record_date.ends_with("-01-01");
    let mut enabled_count = 0;
    let mut rules = Vec::new();
    for fields in &config.action_rules {
        if !enabled(field(fields, &["启用"])) {
            continue;
        }
        enabled_count += 1;
        let rule_id = field(fields, &["规则编号"]).trim();
        ensure!(!rule_id.is_empty(), "3.1 启用的规则动作缺少规则编号");
        let action = field(fields, &["规则动作"]).trim().to_owned();
        let cumulative = field(fields, &["来源分组"]).trim() == "当年累计";
        let raw_code = field(fields, &["指标代码"]).trim();
        ensure!(!raw_code.is_empty(), "3.1 规则动作 {rule_id} 缺少指标代码");
        let codes = if cumulative {
            vec![raw_code.replace('\'', "").trim().to_owned()]
        } else {
            raw_code
                .split('/')
                .map(|value| value.replace('\'', "").trim().to_owned())
                .filter(|value| !value.is_empty())
                .collect::<Vec<_>>()
        };
        ensure!(!codes.is_empty(), "3.1 规则动作 {rule_id} 缺少有效指标代码");
        if !cumulative {
            ensure!(
                action_supported(&action),
                "3.1 规则动作 {rule_id} 使用尚未支持的动作：{action}"
            );
        }
        let target_indicator = field(fields, &["比较指标代码"])
            .replace('\'', "")
            .trim()
            .to_owned();
        let requires_target = action == "指标需要对应存在。"
            || action == "指标应相等核查。"
            || action.contains("相除")
            || action.contains("太小核查")
            || action.contains("相减");
        ensure!(
            !requires_target || !target_indicator.is_empty(),
            "3.1 规则动作 {rule_id}（{action}）缺少比较指标代码"
        );

        let frequency = field(fields, &["适用频度"]).trim().to_owned();
        let batch = field(fields, &["适用批次"]).trim().to_owned();
        let scene = field(fields, &["适用场景"]).trim().to_owned();
        if scene == "结转" && !january_first {
            continue;
        }
        if !cumulative && scene != "结转" && january_first {
            continue;
        }
        if !cumulative && !frequency.is_empty() {
            let key = if batch.is_empty() {
                frequency.clone()
            } else {
                format!("{frequency}{batch}")
            };
            if !available_frequency_keys.contains(&key) {
                continue;
            }
        }
        rules.push(CompiledActionRule {
            rule_id: rule_id.to_owned(),
            action: if cumulative && action.is_empty() {
                "当年累计指标比上期不应减少。".to_owned()
            } else {
                action
            },
            codes,
            data_attr: field(fields, &["数据属性"]).trim().to_owned(),
            currency: field(fields, &["币种"]).trim().to_owned(),
            frequency,
            batch,
            scene,
            rmb_threshold: parse_config_number(field(fields, &["人民币阈值(亿元)"])).unwrap_or(0.0),
            usd_threshold: parse_config_number(field(fields, &["美元阈值(亿美元)"])).unwrap_or(0.0),
            amplitude_threshold: parse_config_number(field(fields, &["变幅阈值(%)"]))
                .unwrap_or(0.0),
            target_indicator,
            detail: field(fields, &["详细说明"]).trim().to_owned(),
            note: field(fields, &["备注"]).trim().to_owned(),
            cumulative,
        });
    }
    Ok((rules, enabled_count))
}

fn action_supported(action: &str) -> bool {
    matches!(
        action,
        "指标不应有数，需说明。"
            | "指标有变动需说明。"
            | "指标变动或变幅超过阈值需说明。"
            | "指标变幅异常需说明。"
            | "当年累计指标比上期不应减少。"
            | "历史累计指标比上期不应减少。"
            | "指标比上期增加需说明。"
            | "指标比上期减少需说明。"
            | "指标新增需说明。"
            | "指标结清需说明。"
            | "指标为负数需核实。"
            | "指标应为整数。"
            | "指标应能被某数整除。"
            | "指标本期余额应小于（上期余额+本期发生额）。"
            | "指标本期余额应小于（上期余额+本期当年发生额-上期当年累计发生额）。"
            | "指标需要对应存在。"
            | "指标应相等核查。"
    ) || action.contains("相除")
        || action.contains("太小核查")
        || action.contains("相减")
}

fn action_rule_applies(
    rule: &CompiledActionRule,
    record: &CentralRecord,
    current_date: &str,
) -> bool {
    if !rule.frequency.is_empty() && rule.frequency != record.frequency {
        return false;
    }
    if !rule.batch.is_empty() && rule.batch != record.batch {
        return false;
    }
    if !rule.data_attr.is_empty()
        && rule.data_attr != "0"
        && !rule.data_attr.contains(&record.data_attr)
    {
        return false;
    }
    if !rule.currency.is_empty() && rule.currency != record.currency {
        return false;
    }
    if rule.scene == "结转" && !current_date.ends_with("-01-01") {
        return false;
    }
    true
}

fn number_or_zero(value: &CentralValue) -> f64 {
    match value {
        CentralValue::Number(value) if value.is_finite() => *value,
        CentralValue::Number(_) | CentralValue::Text(_) | CentralValue::Empty => 0.0,
    }
}

fn cumulative_action_message(
    row: &S3ExecutionCompareRow,
    current_date: &str,
    previous_date: &str,
) -> String {
    let (Some(current), Some(previous)) =
        (parse_iso_date(current_date), parse_iso_date(previous_date))
    else {
        return if number_or_zero(&row.change) < 0.0 {
            " 当年累计指标比上期不应减少，需说明。".to_owned()
        } else {
            String::new()
        };
    };
    let change = number_or_zero(&row.change);
    if number_or_zero(&row.record.value) < 0.0 {
        return "发生额指标不应为负数，需说明。".to_owned();
    }
    if current.month() == 1 && current.day() == 1 {
        return String::new();
    }
    if current.year() == previous.year() && change < 0.0 {
        return " 当年累计指标比上期不应减少，需说明。".to_owned();
    }
    if current.year() != previous.year() {
        if current.month() >= previous.month() && change < 0.0 {
            return format!(
                "{}数据比{}的小，请核实。",
                row.record.indicator_name,
                &previous_date[..previous_date.len().min(7)]
            );
        }
        if current.month() < previous.month() && change >= 0.0 {
            return format!(
                "{}数据 >= {}的，请核实。",
                row.record.indicator_name,
                &previous_date[..previous_date.len().min(7)]
            );
        }
    }
    String::new()
}

fn format_action_message(
    rule: &CompiledActionRule,
    hit_message: &str,
    facts: &ActionFacts,
) -> String {
    let amount = if facts.is_usd && rule.usd_threshold != 0.0 {
        rule.usd_threshold
    } else {
        rule.rmb_threshold
    };
    let amount_unit = if facts.is_usd && (rule.usd_threshold != 0.0 || rule.rmb_threshold == 0.0) {
        "亿美元"
    } else {
        "亿元"
    };
    let amount_text = format!("阈值{}{amount_unit}", format_config_number(amount));
    let amplitude_text = format!(
        "变动幅度{}%",
        format_config_number(rule.amplitude_threshold * 100.0)
    );
    let threshold_text = if rule.action == "指标变幅异常需说明。" {
        amplitude_text
    } else if rule.action == "指标有变动需说明。" || rule.action == "指标变动或变幅超过阈值需说明。"
    {
        let mut parts = Vec::new();
        if rule.rmb_threshold != 0.0 || rule.usd_threshold != 0.0 || rule.amplitude_threshold == 0.0
        {
            parts.push(amount_text);
        }
        if rule.amplitude_threshold != 0.0 {
            parts.push(amplitude_text);
        }
        parts.join("、")
    } else if rule.action == "指标应能被某数整除。" {
        format!("阈值{}（除数）", format_config_number(rule.rmb_threshold))
    } else if rule.action.contains("相除") || rule.action.contains("太小核查") {
        format!("比值阈值{}", format_config_number(amount))
    } else if rule.action == "指标新增需说明。" || rule.action == "指标结清需说明。"
    {
        format!("阈值{}亿元", format_config_number(rule.rmb_threshold))
    } else if rule.action == "指标比上期增加需说明。"
        || rule.action == "指标比上期减少需说明。"
        || rule.action.contains("相减")
    {
        amount_text
    } else {
        "阈值不适用".to_owned()
    };
    let mut note = rule.note.clone();
    if hit_message != rule.action {
        let detail = format!("判定：{}", hit_message.trim());
        note = if note.is_empty() {
            detail
        } else {
            format!("{note}；{detail}")
        };
    }
    let head = format!(
        "{}-{}-{}",
        if rule.rule_id.is_empty() {
            "未编号"
        } else {
            &rule.rule_id
        },
        rule.action,
        threshold_text
    );
    if note.is_empty() {
        head
    } else {
        format!("{head}-{note}")
    }
}

#[allow(clippy::too_many_arguments)]
fn execute_special_action(
    rule: &CompiledActionRule,
    row: &S3ExecutionCompareRow,
    facts: &ActionFacts,
    current_index: &BTreeMap<String, &CentralRecord>,
    previous_index: &BTreeMap<String, &CentralRecord>,
    current_date: &str,
    previous_date: &str,
    config: &S3ExecutionCompareConfig,
) -> String {
    let action = rule.action.as_str();
    let record = &row.record;
    let amount_threshold = if facts.is_usd && rule.usd_threshold != 0.0 {
        rule.usd_threshold
    } else {
        rule.rmb_threshold
    };
    if action == "指标不应有数，需说明。" {
        return if matches!(&record.value, CentralValue::Empty)
            || matches!(&record.value, CentralValue::Text(value) if value.is_empty())
        {
            String::new()
        } else {
            action.to_owned()
        };
    }
    if action == "指标有变动需说明。" {
        let changed_amount = facts.change.abs() * facts.factor;
        if rule.amplitude_threshold != 0.0 {
            let amount_hit = if amount_threshold == 0.0 {
                facts.change != 0.0
            } else {
                changed_amount >= amount_threshold
            };
            return if amount_hit && facts.ratio.abs() > rule.amplitude_threshold {
                action.to_owned()
            } else {
                String::new()
            };
        }
        let hit = if rule.rmb_threshold == 0.0 && rule.usd_threshold == 0.0 {
            facts.change != 0.0
        } else {
            facts.change != 0.0 && changed_amount >= amount_threshold
        };
        return if hit {
            action.to_owned()
        } else {
            String::new()
        };
    }
    if action == "指标变动或变幅超过阈值需说明。" {
        let value_hit =
            amount_threshold > 0.0 && facts.current.abs() * facts.factor >= amount_threshold;
        let range_hit =
            rule.amplitude_threshold > 0.0 && facts.ratio.abs() > rule.amplitude_threshold;
        return if value_hit || range_hit {
            action.to_owned()
        } else {
            String::new()
        };
    }
    if action == "指标变幅异常需说明。" {
        return if rule.amplitude_threshold > 0.0 && facts.ratio.abs() >= rule.amplitude_threshold {
            format!("{:.2}%，{action}", facts.ratio * 100.0)
        } else {
            String::new()
        };
    }
    if action == "当年累计指标比上期不应减少。" {
        let active = parse_iso_date(current_date).is_some_and(|current| {
            parse_iso_date(previous_date).is_some_and(|previous| {
                current.year() != previous.year() || (previous.month() == 1 && previous.day() == 1)
            })
        });
        if active {
            return cumulative_action_message(row, current_date, previous_date);
        }
        return if facts.change < 0.0 {
            action.to_owned()
        } else {
            String::new()
        };
    }
    if action == "历史累计指标比上期不应减少。" {
        return if facts.current < 0.0 || facts.change < 0.0 {
            action.to_owned()
        } else {
            String::new()
        };
    }
    if action == "指标比上期增加需说明。" || action == "指标比上期减少需说明。"
    {
        let scaled = facts.change * facts.factor;
        let direction =
            (action.contains("增加") && scaled > 0.0) || (action.contains("减少") && scaled < 0.0);
        return if direction && scaled.abs() > amount_threshold {
            action.to_owned()
        } else {
            String::new()
        };
    }
    if action == "指标新增需说明。" || action == "指标结清需说明。" {
        let hit = if action.contains("新增") {
            facts.current > 0.0 && facts.previous == 0.0
        } else {
            facts.current == 0.0 && facts.previous > 0.0
        };
        return if hit && facts.change.abs() * facts.factor >= rule.rmb_threshold.abs() {
            action.to_owned()
        } else {
            String::new()
        };
    }
    if action == "指标为负数需核实。" {
        return if facts.current < 0.0 {
            action.to_owned()
        } else {
            String::new()
        };
    }
    if action == "指标应为整数。" {
        return if facts.current.abs().fract() == 0.0 {
            String::new()
        } else {
            action.to_owned()
        };
    }
    if action == "指标应能被某数整除。" {
        let divisor = rule.rmb_threshold.trunc() as i64;
        if divisor == 0 {
            return "除数为0，特殊指标表设置有误。".to_owned();
        }
        let factor = if config.exempt_indicators.contains(&record.indicator) {
            1.0
        } else {
            crate::s3_current_value_check::s3_target_unit_factor(&config.target_unit)
        };
        let quotient = facts.current.abs() * factor / divisor as f64;
        return if (quotient - round_ties_even(quotient)).abs() > 1e-10 {
            format!("指标应能被{divisor}整除。{}", rule.detail)
        } else {
            String::new()
        };
    }
    if action == "指标本期余额应小于（上期余额+本期发生额）。" {
        return balance_and_flow_message(
            row,
            current_index,
            previous_index,
            current_date,
            previous_date,
            config,
            false,
        );
    }
    if action == "指标本期余额应小于（上期余额+本期当年发生额-上期当年累计发生额）。"
    {
        return balance_and_flow_message(
            row,
            current_index,
            previous_index,
            current_date,
            previous_date,
            config,
            true,
        );
    }
    if action == "指标需要对应存在。" || action == "指标应相等核查。" {
        let opponent = opponent_value(row, rule, current_index, action);
        if action == "指标需要对应存在。" {
            if facts.current != 0.0 && opponent == 0.0 {
                return format!("{}有数，对应指标无数，需说明。", record.indicator_name);
            }
            if facts.current == 0.0 && opponent != 0.0 {
                return format!("{}无数，对应指标有数，需说明。", record.indicator_name);
            }
            return String::new();
        }
        let difference = facts.current - opponent;
        return if !rule.target_indicator.is_empty() && difference.abs() > DBL_EPSILON {
            format!(
                "{}不等于{}偏差{:.4},请核实。{}",
                record.indicator_name, rule.target_indicator, difference, rule.detail
            )
        } else {
            String::new()
        };
    }
    if action.contains("相除") || action.contains("太小核查") {
        let opponent = opponent_value(row, rule, current_index, action);
        let opponent_factor = if config.exempt_indicators.contains(&rule.target_indicator) {
            1.0
        } else {
            facts.factor
        };
        if opponent == 0.0 {
            return if facts.current != 0.0 {
                format!(
                    "{}有数，比较指标{}为空或为0，需核实。",
                    record.indicator_name, rule.target_indicator
                )
            } else {
                String::new()
            };
        }
        let quotient = (facts.current * facts.factor / (opponent * opponent_factor)).abs();
        let limit = if facts.is_usd && rule.usd_threshold != 0.0 {
            rule.usd_threshold
        } else {
            rule.rmb_threshold
        };
        if action.contains("太小") && quotient < limit {
            return format!(
                "{}={quotient:.4},小于{limit}，需说明。",
                record.indicator_name
            );
        }
        if !action.contains("太小") && quotient > limit {
            return format!(
                "{}={quotient:.4},超过{limit}，需说明。",
                record.indicator_name
            );
        }
        return String::new();
    }
    if action.contains("相减") {
        let opponent = opponent_value(row, rule, current_index, action);
        let opponent_factor = if config.exempt_indicators.contains(&rule.target_indicator) {
            1.0
        } else {
            facts.factor
        };
        let limit = if facts.is_usd && rule.usd_threshold != 0.0 {
            rule.usd_threshold
        } else {
            rule.rmb_threshold
        };
        return if facts.current * facts.factor - opponent * opponent_factor < limit.abs() {
            action.to_owned()
        } else {
            String::new()
        };
    }
    String::new()
}

fn opponent_value(
    row: &S3ExecutionCompareRow,
    rule: &CompiledActionRule,
    current_index: &BTreeMap<String, &CentralRecord>,
    action: &str,
) -> f64 {
    if rule.target_indicator.is_empty() {
        return 0.0;
    }
    let mut currency = row.record.currency.as_str();
    if currency == "美元合计"
        && ((action == "指标需要对应存在。" && rule.usd_threshold != 0.0)
            || ((action.contains("相除") || action.contains("太小核查"))
                && rule.usd_threshold == 999.0))
    {
        currency = "人民币";
    }
    let key = record_key_with_values(
        &row.record,
        &rule.target_indicator,
        &row.record.data_attr,
        currency,
    );
    current_index
        .get(&key)
        .map(|record| number_or_zero(&record.value))
        .unwrap_or(0.0)
}

fn record_key_with_indicator(record: &CentralRecord, indicator: &str) -> String {
    record_key_with_values(record, indicator, &record.data_attr, &record.currency)
}

fn record_key_with_values(
    record: &CentralRecord,
    indicator: &str,
    data_attr: &str,
    currency: &str,
) -> String {
    [
        record.biz_class.as_str(),
        record.org_code.as_str(),
        record.region_code.as_str(),
        indicator,
        data_attr,
        currency,
        record.frequency.as_str(),
        record.batch.as_str(),
    ]
    .concat()
}

fn balance_and_flow_message(
    row: &S3ExecutionCompareRow,
    current_index: &BTreeMap<String, &CentralRecord>,
    previous_index: &BTreeMap<String, &CentralRecord>,
    current_date: &str,
    previous_date: &str,
    config: &S3ExecutionCompareConfig,
    accumulated: bool,
) -> String {
    let (Some(current), Some(previous)) =
        (parse_iso_date(current_date), parse_iso_date(previous_date))
    else {
        return "不进行余额与发生额校验，因数据日期无法识别。".to_owned();
    };
    let record = &row.record;
    let counterpart_attr = if record.data_attr == "余额" {
        "发生额"
    } else {
        "余额"
    };
    let months = match record.frequency.as_str() {
        "季" => 3,
        "半年" => 6,
        "年" => 12,
        _ => 1,
    };
    let continuous = if accumulated {
        current.year() == previous.year()
            || (current.year() - previous.year() == 1
                && previous.month() == 12
                && previous.day() == 31)
    } else {
        previous == previous_month_end(current, months)
            || (previous.month() == 1
                && previous.day() == 1
                && previous == period_start(current, months))
    };
    if !continuous {
        let label = if accumulated {
            "余额与当年累计发生额"
        } else {
            "余额与本期发生额"
        };
        let discontinuity = if accumulated {
            "年份不连续"
        } else {
            "不连续"
        };
        return format!("不进行{label}校验，因{current_date}与{previous_date}{discontinuity}。");
    }

    let counterpart = |data_attr: &str, index: &BTreeMap<String, &CentralRecord>| {
        let key = record_key_with_values(record, &record.indicator, data_attr, &record.currency);
        index
            .get(&key)
            .map(|value| number_or_zero(&value.value))
            .unwrap_or(0.0)
    };
    let (current_balance, previous_balance, current_flow, mut previous_flow) =
        if counterpart_attr == "发生额" {
            (
                number_or_zero(&record.value),
                number_or_zero(&row.previous_value),
                counterpart("发生额", current_index),
                if accumulated {
                    counterpart("发生额", previous_index)
                } else {
                    0.0
                },
            )
        } else {
            (
                counterpart("余额", current_index),
                counterpart("余额", previous_index),
                number_or_zero(&record.value),
                if accumulated {
                    number_or_zero(&row.previous_value)
                } else {
                    0.0
                },
            )
        };
    if accumulated
        && ((previous.month() == 12 && previous.day() == 31)
            || (previous.month() == 1 && previous.day() == 1))
    {
        previous_flow = 0.0;
    }
    let to_yuan = if config.exempt_indicators.contains(&record.indicator) {
        1.0
    } else {
        crate::s3_current_value_check::s3_target_unit_factor(&config.target_unit)
    };
    let current_balance = current_balance * to_yuan;
    let previous_balance = previous_balance * to_yuan;
    let current_flow = current_flow * to_yuan;
    let previous_flow = previous_flow * to_yuan;
    let maximum = current_balance
        .max(previous_balance)
        .max(current_flow)
        .max(previous_flow);
    let delta = if accumulated {
        previous_balance + current_flow - previous_flow - current_balance
    } else {
        previous_balance + current_flow - current_balance
    };
    let mut message = if python_round_fixed(delta, 2) >= 0.0 {
        String::new()
    } else if accumulated && python_round_fixed(current_flow - previous_flow, 10) == 0.0 {
        "本期无发生额，余额增加了，请核实".to_owned()
    } else {
        balance_error_message(delta, maximum)
    };
    if !accumulated && current_flow < 0.0 {
        message = format!("发生额指标不应为负数。{message}");
    }
    message
}

fn balance_error_message(delta: f64, maximum: f64) -> String {
    if python_round_fixed(delta, 2) >= 0.0 {
        return String::new();
    }
    let absolute = delta.abs();
    if python_round_fixed(absolute, 2) < 150.0 {
        return "精度误差：上期余额+本期发生额<本期余额，偏差150元内。".to_owned();
    }
    if maximum < 10_000.0 {
        if python_round_fixed(absolute, 2) < maximum {
            return "精度误差：上期余额+本期发生额<本期余额，超过150元，未超过1万元。".to_owned();
        }
        return "上期余额+本期发生额>本期余额，超过该业务量万分之一。".to_owned();
    }
    if python_round_fixed(absolute, 2) < 10_000.0 {
        return "精度误差：上期余额+本期发生额<本期余额，超过150元，未超过10000元。".to_owned();
    }
    "上期余额+本期发生额>本期余额，超过10000元。".to_owned()
}

fn parse_iso_date(value: &str) -> Option<NaiveDate> {
    NaiveDate::parse_from_str(value, "%Y-%m-%d").ok()
}

fn previous_month_end(current: NaiveDate, months: i32) -> NaiveDate {
    first_day_of_period(current, months) - Duration::days(1)
}

fn period_start(current: NaiveDate, months: i32) -> NaiveDate {
    first_day_of_period(current, months)
}

fn first_day_of_period(current: NaiveDate, months: i32) -> NaiveDate {
    let mut year = current.year();
    let mut month = current.month() as i32 - months + 1;
    while month <= 0 {
        month += 12;
        year -= 1;
    }
    NaiveDate::from_ymd_opt(year, month as u32, 1).expect("calculated month is valid")
}

fn python_round_fixed(value: f64, digits: i32) -> f64 {
    let factor = 10_f64.powi(digits);
    round_ties_even(value * factor) / factor
}

fn round_ties_even(value: f64) -> f64 {
    let truncated = value.trunc();
    let fraction = value - truncated;
    if fraction.abs() < 0.5 {
        truncated
    } else if fraction.abs() > 0.5 || truncated % 2.0 != 0.0 {
        truncated + fraction.signum()
    } else {
        truncated
    }
}

fn frequency_keys(current: &CentralDataset, previous: &CentralDataset) -> BTreeSet<String> {
    let mut keys = BTreeSet::from([String::new()]);
    for dataset in [current, previous] {
        for record in &dataset.records {
            keys.insert(record.frequency.clone());
            keys.insert(format!("{}{}", record.frequency, record.batch));
        }
    }
    keys
}

fn substitute_expression_references(
    expression: &str,
    current_row: &CentralRecord,
    current_index: &BTreeMap<String, &CentralRecord>,
    previous_index: &BTreeMap<String, &CentralRecord>,
    exempt_indicators: &BTreeSet<String>,
    unit_factor: f64,
    mut registered: Option<&mut BTreeSet<String>>,
) -> String {
    let references = expression_references(expression);
    let mut output = String::with_capacity(expression.len());
    let mut cursor = 0;
    for reference in references {
        output.push_str(&expression[cursor..reference.start]);
        let key = expression_target_key(&reference.segments, current_row);
        if let Some(keys) = registered.as_deref_mut() {
            keys.insert(key.clone());
        }
        let index = if reference.current_period {
            current_index
        } else {
            previous_index
        };
        let replacement = match index.get(&key) {
            None => "0".to_owned(),
            Some(record) => {
                let value = match record.value {
                    CentralValue::Number(value) if value.is_finite() => value,
                    _ => 0.0,
                };
                if value == 0.0 {
                    "0.01".to_owned()
                } else {
                    let indicator = reference
                        .segments
                        .get(3)
                        .map(|value| value.replace('\'', ""))
                        .unwrap_or_default();
                    let factor = if exempt_indicators.contains(&indicator) {
                        1.0
                    } else {
                        unit_factor
                    };
                    python_cstr(value * factor)
                }
            }
        };
        output.push_str(&replacement);
        cursor = reference.end;
    }
    output.push_str(&expression[cursor..]);
    output
}

fn expression_target_key(segments: &[String], current_row: &CentralRecord) -> String {
    let mut segments = segments.to_vec();
    while segments.len() < 8 {
        segments.push(String::new());
    }
    let inherited = [
        current_row.biz_class.as_str(),
        current_row.org_code.as_str(),
        current_row.region_code.as_str(),
        current_row.indicator.as_str(),
        current_row.data_attr.as_str(),
        current_row.currency.as_str(),
        current_row.frequency.as_str(),
        current_row.batch.as_str(),
    ];
    let mut values = Vec::with_capacity(8);
    for position in 0..8 {
        let configured = segments[position].trim();
        let resolved = if configured.is_empty()
            || ((position == 1 || position == 2) && configured.contains('^'))
        {
            inherited[position].to_owned()
        } else if position == 1 || position == 2 || position == 3 {
            configured.replace('\'', "")
        } else {
            configured.to_owned()
        };
        values.push(resolved);
    }
    values.concat()
}

fn python_cstr(value: f64) -> String {
    if value == 0.0 {
        return "0".to_owned();
    }
    if !value.is_finite() {
        return value.to_string();
    }
    let scientific = format!("{value:.14e}");
    let (mantissa, exponent) = scientific.split_once('e').unwrap_or((&scientific, "0"));
    let exponent = exponent.parse::<i32>().unwrap_or(0);
    if !(-4..15).contains(&exponent) {
        let mantissa = mantissa.trim_end_matches('0').trim_end_matches('.');
        return format!("{mantissa}e{exponent:+03}");
    }
    let decimal_places = usize::try_from(14 - exponent).unwrap_or(0);
    format!("{value:.decimal_places$}")
        .trim_end_matches('0')
        .trim_end_matches('.')
        .to_owned()
}

fn append_joined(target: &mut String, value: &str) {
    if target.is_empty() {
        target.push_str(value);
    } else {
        target.push_str("    ||    ");
        target.push_str(value);
    }
}

fn append_unique(target: &mut String, value: &str) {
    if value.is_empty() || target.split("    ||    ").any(|existing| existing == value) {
        return;
    }
    append_joined(target, value);
}

fn base_row(record: CentralRecord, config: &S3ExecutionCompareConfig) -> S3ExecutionCompareRow {
    let reference = config.indicator_references.get(&record.indicator);
    let order_code_out = reference.map_or_else(
        || record.indicator.clone(),
        |reference| {
            let row_number = reference
                .row_number
                .parse::<f64>()
                .ok()
                .filter(|value| value.is_finite())
                .map(|value| value as i64)
                .unwrap_or(0);
            format!("{}_{row_number:04}", reference.form_code)
        },
    );
    let indicator_name_out = reference
        .map(|reference| reference.alias.trim())
        .filter(|alias| !alias.is_empty())
        .unwrap_or(&record.indicator_name)
        .to_owned();
    S3ExecutionCompareRow {
        record,
        order_code_out,
        indicator_name_out,
        previous_value: CentralValue::Empty,
        change: CentralValue::Empty,
        ratio: CentralValue::Empty,
        remark: String::new(),
        need_explanation: String::new(),
        explanation_content: String::new(),
        calculation_process: String::new(),
        fill_color: None,
        fill_scope: String::new(),
    }
}

fn apply_alert_band(
    row: &mut S3ExecutionCompareRow,
    ratio: f64,
    change: f64,
    config: &S3ExecutionCompareConfig,
) {
    if vba_round(ratio, 6) == 0.0 {
        return;
    }
    let Some(band) = matching_alert_band(ratio, &config.alert_bands) else {
        row.remark = REMARK_SHRINK_100.to_owned();
        row.fill_color = Some(10);
        row.fill_scope = "row".to_owned();
        return;
    };
    row.remark = format_alert_remark(band);
    if let Some(color) = parse_config_int(&band.fill_color) {
        if color != 0 {
            row.fill_color = u8::try_from(color).ok();
            row.fill_scope = if truthy(&band.fill_entire_row) {
                "row".to_owned()
            } else {
                "remark".to_owned()
            };
        }
    }
    if let Some(threshold) = parse_config_number(&band.amount_threshold) {
        let configured = band.explanation.trim();
        if !configured.is_empty() && !matches_disabled(configured) && change.abs() > threshold.abs()
        {
            let configured = if matches!(configured, "是" | "1" | "true" | "True") {
                row.remark.as_str()
            } else {
                configured
            };
            row.need_explanation = configured
                .replace("{金额变动阈值}", &format_config_number(threshold))
                .replace("{目标单位}", &config.target_unit)
                .replace("{备注文字}", &row.remark);
        }
    }
}

fn matching_alert_band(ratio: f64, bands: &[S3AlertBand]) -> Option<&S3AlertBand> {
    let percent = ratio * 100.0;
    bands.iter().find(|band| {
        let Some(lower) = parse_optional_bound(&band.lower) else {
            return false;
        };
        let Some(upper) = parse_optional_bound(&band.upper) else {
            return false;
        };
        if lower.is_some_and(|lower| round_ten(percent - lower) < 0.0) {
            return false;
        }
        upper.is_none_or(|upper| round_ten(percent - upper) < 0.0)
    })
}

fn parse_optional_bound(value: &str) -> Option<Option<f64>> {
    if value.trim().is_empty() {
        Some(None)
    } else {
        parse_config_number(value).map(Some)
    }
}

fn format_alert_remark(band: &S3AlertBand) -> String {
    let manual = band.remark_text.trim();
    let active = !manual.is_empty()
        || (!band.fill_color.trim().is_empty() && !matches_disabled(&band.fill_color))
        || (!band.fill_entire_row.trim().is_empty() && !matches_disabled(&band.fill_entire_row))
        || !band.amount_threshold.trim().is_empty()
        || !band.explanation.trim().is_empty();
    if !active {
        return String::new();
    }
    let lower = parse_config_number(&band.lower);
    let upper = parse_optional_bound(&band.upper).flatten();
    let automatic = match (lower, upper) {
        (None, Some(upper)) if upper <= 0.0 => format!("降幅<{}%", format_config_number(upper)),
        (None, Some(upper)) => format!("变幅<{}%", format_config_number(upper)),
        (Some(lower), None) if lower >= 0.0 => format!("增幅≥{}%", format_config_number(lower)),
        (Some(lower), None) => format!("变幅≥{}%", format_config_number(lower)),
        (Some(lower), Some(upper)) => {
            let prefix = if lower >= 0.0 {
                "增幅"
            } else if upper <= 0.0 {
                "降幅"
            } else {
                "变幅"
            };
            format!(
                "{prefix}[{}%,{}%)",
                format_config_number(lower),
                format_config_number(upper)
            )
        }
        (None, None) => String::new(),
    };
    match (automatic.is_empty(), manual.is_empty()) {
        (false, false) => format!("{automatic} | {manual}"),
        (false, true) => automatic,
        (true, false) => manual.to_owned(),
        (true, true) => String::new(),
    }
}

fn read_parameter_rows(sheet: &SheetValues) -> Result<BTreeMap<String, String>> {
    let columns = header_columns(sheet);
    for required in ["参数", "值"] {
        ensure!(
            columns.contains_key(required),
            "“运行参数”工作表缺少“{required}”列"
        );
    }
    let mut result = BTreeMap::new();
    for row in 2..=sheet.max_row {
        let name = cell_text(sheet, row, columns["参数"]);
        if !name.trim().is_empty() {
            result.insert(
                name.trim().to_owned(),
                cell_text(sheet, row, columns["值"]).trim().to_owned(),
            );
        }
    }
    Ok(result)
}

fn read_rows(sheet: &SheetValues, required: &[&str]) -> Result<Vec<BTreeMap<String, String>>> {
    let columns = header_columns(sheet);
    let missing = required
        .iter()
        .filter(|header| !columns.contains_key(**header))
        .copied()
        .collect::<Vec<_>>();
    ensure!(
        missing.is_empty(),
        "工作表“{}”缺少配置列：{}",
        sheet.name,
        missing.join("、")
    );
    let mut rows = Vec::new();
    for row in 2..=sheet.max_row {
        let values = columns
            .iter()
            .map(|(name, column)| (name.clone(), cell_text(sheet, row, *column)))
            .collect::<BTreeMap<_, _>>();
        if values.values().all(|value| value.trim().is_empty()) {
            continue;
        }
        rows.push(values);
    }
    Ok(rows)
}

fn header_columns(sheet: &SheetValues) -> HashMap<String, usize> {
    let mut columns = HashMap::new();
    for column in 1..=sheet.max_column {
        let value = cell_text(sheet, 1, column);
        if !value.trim().is_empty() {
            columns.insert(value.trim().to_owned(), column);
        }
    }
    columns
}

fn cell_text(sheet: &SheetValues, row: usize, column: usize) -> String {
    let address = format!("{}{}", column_name(column), row);
    match sheet.values.get(&address).unwrap_or(&CellValue::Empty) {
        CellValue::Empty => String::new(),
        CellValue::Text(value) => value.clone(),
        CellValue::Number(value) => format!("{value:?}"),
        CellValue::Integer(value) => value.to_string(),
        CellValue::Boolean(value) => value.to_string(),
        CellValue::DateTimeSerial(value) => format!("{value:?}"),
        CellValue::DateTimeIso(value) | CellValue::DurationIso(value) => value.clone(),
        CellValue::Error(value) => value.clone(),
    }
}

fn column_name(mut column: usize) -> String {
    let mut name = String::new();
    while column > 0 {
        let remainder = (column - 1) % 26;
        name.push((b'A' + remainder as u8) as char);
        column = (column - 1) / 26;
    }
    name.chars().rev().collect()
}

fn find_sheet<'a>(sheets: &'a [SheetValues], name: &str) -> Option<&'a SheetValues> {
    sheets.iter().find(|sheet| sheet.name == name)
}

fn ensure_xlsx(path: &Path, description: &str) -> Result<()> {
    ensure!(
        path.is_file(),
        "{description}文件不存在：{}",
        path.display()
    );
    ensure!(
        path.extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx")),
        "{description}目前只支持 XLSX：{}",
        path.display()
    );
    Ok(())
}

fn enabled(value: &str) -> bool {
    !matches_disabled(value)
}

fn matches_disabled(value: &str) -> bool {
    let normalized = value.trim().to_ascii_lowercase();
    matches!(normalized.as_str(), "否" | "0" | "false" | "n" | "no")
        || normalized.parse::<f64>().is_ok_and(|number| number == 0.0)
}

fn truthy(value: &str) -> bool {
    !matches_disabled(value) && !value.trim().is_empty()
}

fn parse_config_number(value: &str) -> Option<f64> {
    value
        .trim()
        .parse::<f64>()
        .ok()
        .filter(|value| value.is_finite())
}

fn parse_config_int(value: &str) -> Option<i64> {
    parse_config_number(value).map(|value| value as i64)
}

fn format_config_number(value: f64) -> String {
    if value.fract() == 0.0 {
        format!("{value:.0}")
    } else {
        value.to_string()
    }
}

fn value_number(value: &CentralValue) -> Option<f64> {
    match value {
        CentralValue::Number(value) if value.is_finite() => Some(*value),
        _ => None,
    }
}

fn excel_round(value: f64, digits: i32) -> f64 {
    let factor = 10_f64.powi(digits);
    let scaled = value * factor;
    let truncated = scaled.trunc();
    let fraction = scaled - truncated;
    let rounded = if scaled >= 0.0 {
        truncated + if fraction >= 0.5 { 1.0 } else { 0.0 }
    } else {
        truncated - if fraction.abs() >= 0.5 { 1.0 } else { 0.0 }
    };
    rounded / factor
}

fn vba_round(value: f64, digits: i32) -> f64 {
    let factor = 10_f64.powi(digits);
    let scaled = value * factor;
    let truncated = scaled.trunc();
    let fraction = scaled - truncated;
    let rounded = if fraction > 0.5 {
        truncated + 1.0
    } else if fraction < -0.5 {
        truncated - 1.0
    } else if fraction == 0.5 {
        if truncated % 2.0 == 0.0 {
            truncated + 1.0
        } else {
            truncated
        }
    } else if fraction == -0.5 {
        if truncated % 2.0 == 0.0 {
            truncated
        } else {
            truncated - 1.0
        }
    } else {
        truncated
    };
    rounded / factor
}

fn round_ten(value: f64) -> f64 {
    (value * 10_f64.powi(10)).round_ties_even() / 10_f64.powi(10)
}

fn round_digits(target_unit: &str) -> i32 {
    match target_unit {
        "元" => 2,
        "十元" => 3,
        "百元" => 4,
        "千元" => 5,
        "万元" => 6,
        "十万元" => 7,
        "百万元" => 8,
        "千万" | "千万元" => 9,
        "亿元" => 10,
        _ => 2,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn numeric_zero_config_flags_are_disabled() {
        for value in ["0", "0.0", "-0.0", "0.000"] {
            assert!(matches_disabled(value), "{value} must be disabled");
            assert!(!truthy(value), "{value} must not be truthy");
        }
        for value in ["否", "false"] {
            assert!(matches_disabled(value));
            assert!(!truthy(value));
        }
        assert!(!matches_disabled(""));
        assert!(!truthy(""));
        for value in ["1", "1.0", "6.0"] {
            assert!(!matches_disabled(value));
            assert!(truthy(value));
        }
    }

    fn record(
        org: &str,
        region: &str,
        indicator: &str,
        frequency: &str,
        batch: &str,
        value: CentralValue,
    ) -> CentralRecord {
        CentralRecord {
            biz_class: "人民币".to_owned(),
            record_date: "2026-08-31".to_owned(),
            org_code: org.to_owned(),
            org_name: format!("机构 {org}"),
            region_code: region.to_owned(),
            region_name: format!("地区 {region}"),
            order_code: indicator.to_owned(),
            indicator: indicator.to_owned(),
            indicator_name: format!("指标 {indicator}"),
            data_attr: "余额".to_owned(),
            currency: "人民币".to_owned(),
            frequency: frequency.to_owned(),
            batch: batch.to_owned(),
            value,
        }
    }

    fn dataset(records: Vec<CentralRecord>, date: &str) -> CentralDataset {
        CentralDataset {
            records,
            record_date: date.to_owned(),
            ..CentralDataset::default()
        }
    }

    fn expression_rule(
        schema: S3ExpressionSchema,
        id: &str,
        expression: &str,
        enabled_value: &str,
    ) -> S3ExpressionRule {
        let fields = [
            ("规则编号", id),
            ("规则说明", "核验表达式命中"),
            ("校验表达式", expression),
            ("是否取反", "否"),
            ("容差值(万元)", ""),
            ("启用", enabled_value),
            ("来源分组", "自定义"),
        ]
        .into_iter()
        .map(|(name, value)| (name.to_owned(), value.to_owned()))
        .collect();
        S3ExpressionRule { schema, fields }
    }

    fn action_rule(
        id: &str,
        action: &str,
        code: &str,
        threshold: &str,
    ) -> BTreeMap<String, String> {
        [
            ("规则编号", id),
            ("规则动作", action),
            ("指标代码", code),
            ("数据属性", "余额"),
            ("币种", "人民币"),
            ("人民币阈值(亿元)", threshold),
            ("来源分组", "自定义"),
            ("启用", "是"),
        ]
        .into_iter()
        .map(|(name, value)| (name.to_owned(), value.to_owned()))
        .collect()
    }

    #[test]
    fn base_pairing_keeps_all_dimensions_and_adds_current_previous_only_rows() {
        let current = dataset(
            vec![
                record(
                    "0001",
                    "1234567",
                    "A",
                    "月",
                    "1",
                    CentralValue::Number(200.0),
                ),
                record(
                    "0001",
                    "1234567",
                    "A",
                    "季",
                    "1",
                    CentralValue::Number(70.0),
                ),
                record(
                    "0001",
                    "1234567",
                    "B",
                    "月",
                    "1",
                    CentralValue::Number(10.0),
                ),
            ],
            "2026-08-31",
        );
        let previous = dataset(
            vec![
                record(
                    "0001",
                    "1234567",
                    "A",
                    "月",
                    "1",
                    CentralValue::Number(100.0),
                ),
                record(
                    "0001",
                    "1234567",
                    "A",
                    "季",
                    "2",
                    CentralValue::Number(50.0),
                ),
                record("0001", "1234567", "C", "月", "1", CentralValue::Number(8.0)),
            ],
            "2026-07-31",
        );
        let config = S3ExecutionCompareConfig {
            target_unit: "亿元".to_owned(),
            ..S3ExecutionCompareConfig::default()
        };

        let rows = build_base_comparison(&current, &previous, &config);
        assert_eq!(rows.len(), 5);
        let same_key = rows
            .iter()
            .find(|row| row.record.indicator == "A" && row.record.frequency == "月")
            .unwrap();
        assert_eq!(same_key.previous_value, CentralValue::Number(100.0));
        assert_eq!(same_key.change, CentralValue::Number(100.0));
        assert_eq!(same_key.ratio, CentralValue::Number(1.0));

        let moved_batch = rows
            .iter()
            .find(|row| {
                row.record.indicator == "A"
                    && row.record.frequency == "季"
                    && row.record.batch == "1"
            })
            .unwrap();
        assert_eq!(moved_batch.previous_value, CentralValue::Empty);
        assert_eq!(moved_batch.remark, REMARK_CURRENT_ONLY);

        let previous_only = rows.iter().find(|row| row.record.indicator == "C").unwrap();
        assert_eq!(previous_only.record.record_date, "2026-08-31");
        assert_eq!(previous_only.record.value, CentralValue::Empty);
        assert_eq!(previous_only.change, CentralValue::Number(-8.0));
        assert_eq!(previous_only.ratio, CentralValue::Number(-1.0));
        assert_eq!(previous_only.remark, REMARK_PREVIOUS_ONLY);
    }

    #[test]
    fn base_pairing_uses_indicator_reference_and_alert_configuration() {
        let current = dataset(
            vec![record(
                "0001",
                "1234567",
                "A",
                "月",
                "1",
                CentralValue::Number(200.0),
            )],
            "2026-08-31",
        );
        let previous = dataset(
            vec![record(
                "0001",
                "1234567",
                "A",
                "月",
                "1",
                CentralValue::Number(100.0),
            )],
            "2026-07-31",
        );
        let mut config = S3ExecutionCompareConfig {
            target_unit: "亿元".to_owned(),
            ..S3ExecutionCompareConfig::default()
        };
        config.indicator_references.insert(
            "A".to_owned(),
            S3IndicatorReference {
                form_code: "F1".to_owned(),
                row_number: "7".to_owned(),
                alias: "别名".to_owned(),
            },
        );
        config.alert_bands.push(S3AlertBand {
            lower: "50".to_owned(),
            upper: "150".to_owned(),
            remark_text: "核实".to_owned(),
            fill_color: "6".to_owned(),
            fill_entire_row: "1".to_owned(),
            amount_threshold: "50".to_owned(),
            explanation: "是".to_owned(),
        });

        let rows = build_base_comparison(&current, &previous, &config);
        assert_eq!(rows[0].order_code_out, "F1_0007");
        assert_eq!(rows[0].indicator_name_out, "别名");
        assert!(rows[0].remark.contains("[50%,150%)"));
        assert_eq!(rows[0].fill_color, Some(6));
        assert_eq!(rows[0].fill_scope, "row");
        assert!(!rows[0].need_explanation.is_empty());
    }

    #[test]
    fn config_truthiness_matches_enabled_flag_default() {
        assert!(enabled(""));
        assert!(enabled("是"));
        assert!(!enabled("否"));
        assert!(!enabled("false"));
    }

    #[test]
    fn alerts_allow_an_open_lower_bound_and_previous_only_text_defaults_to_zero() {
        let band = S3AlertBand {
            lower: String::new(),
            upper: "50".to_owned(),
            remark_text: "降幅区间".to_owned(),
            ..S3AlertBand::default()
        };
        assert!(matching_alert_band(-0.25, std::slice::from_ref(&band)).is_some());

        let current = dataset(Vec::new(), "2026-08-31");
        let previous = dataset(
            vec![record(
                "0001",
                "1234567",
                "TEXT",
                "月",
                "1",
                CentralValue::Text("非数值".to_owned()),
            )],
            "2026-07-31",
        );
        let rows = build_base_comparison(&current, &previous, &S3ExecutionCompareConfig::default());
        assert_eq!(rows[0].previous_value, CentralValue::Number(0.0));
        assert_eq!(rows[0].change, CentralValue::Number(-0.0));
        assert_eq!(rows[0].ratio, CentralValue::Number(-1.0));
    }

    #[test]
    fn legacy_eight_segment_expression_resolves_current_and_previous_dimensions() {
        let current = dataset(
            vec![record(
                "0001",
                "1100000",
                "A",
                "月",
                "1",
                CentralValue::Number(30.0),
            )],
            "2026-08-31",
        );
        let previous = dataset(
            vec![record(
                "0001",
                "1100000",
                "B",
                "月",
                "1",
                CentralValue::Number(20.0),
            )],
            "2026-07-31",
        );
        let mut config = S3ExecutionCompareConfig {
            target_unit: "元".to_owned(),
            ..S3ExecutionCompareConfig::default()
        };
        config.expression_rules.push(expression_rule(
            S3ExpressionSchema::LegacyEightSegment,
            "R-1",
            "[人民币,0001,1100000,A,余额,人民币,月,1]>{人民币,0001,1100000,B,余额,人民币,月,1}",
            "是",
        ));
        let mut rows = build_base_comparison(&current, &previous, &config);

        let report = apply_expression_rules(
            &mut rows,
            &current,
            &previous,
            &config,
            S3ExpressionSchema::LegacyEightSegment,
        )
        .unwrap();

        let row = rows.iter().find(|row| row.record.indicator == "A").unwrap();
        assert_eq!(report.enabled_rules, 1);
        assert_eq!(report.triggered_rules, 1);
        assert!(row.need_explanation.contains("核验表达式命中"));
        assert!(row.calculation_process.contains("【30>20】"));
    }

    #[test]
    fn five_segment_expression_uses_legacy_business_binding_and_rule_scope() {
        let current = dataset(
            vec![record(
                "0001",
                "1100000",
                "A",
                "月",
                "1",
                CentralValue::Number(30.0),
            )],
            "2026-08-31",
        );
        let previous = dataset(
            vec![record(
                "0001",
                "1100000",
                "B",
                "月",
                "1",
                CentralValue::Number(20.0),
            )],
            "2026-07-31",
        );
        let legacy = expression_rule(
            S3ExpressionSchema::LegacyEightSegment,
            "R-1",
            "[人民币,,,A,余额,人民币,月,1]>{人民币,,,B,余额,人民币,月,1}",
            "否",
        );
        let mut five = expression_rule(
            S3ExpressionSchema::FiveSegment,
            "R-1",
            "[A,余额,人民币,月,1]>{B,余额,人民币,月,1}",
            "是",
        );
        five.fields
            .insert("机构类代码".to_owned(), "^0001$".to_owned());
        five.fields
            .insert("地区代码".to_owned(), "^1100000$".to_owned());
        let config = S3ExecutionCompareConfig {
            target_unit: "元".to_owned(),
            expression_rules: vec![legacy, five],
            ..S3ExecutionCompareConfig::default()
        };
        let mut rows = build_base_comparison(&current, &previous, &config);

        let report = apply_expression_rules(
            &mut rows,
            &current,
            &previous,
            &config,
            S3ExpressionSchema::FiveSegment,
        )
        .unwrap();

        let row = rows.iter().find(|row| row.record.indicator == "A").unwrap();
        assert_eq!(report.enabled_rules, 1);
        assert_eq!(report.triggered_rules, 1);
        assert!(row.calculation_process.contains("【30>20】"));
    }

    #[test]
    fn five_segment_expression_rejects_non_unique_dimension_keys() {
        let first = record(
            "0001",
            "1100000",
            "A",
            "月",
            "1",
            CentralValue::Number(30.0),
        );
        let mut duplicate = first.clone();
        duplicate.biz_class = "外币".to_owned();
        let current = dataset(vec![first, duplicate], "2026-08-31");
        let previous = dataset(Vec::new(), "2026-07-31");
        let config = S3ExecutionCompareConfig {
            target_unit: "元".to_owned(),
            expression_rules: vec![
                expression_rule(
                    S3ExpressionSchema::LegacyEightSegment,
                    "R-1",
                    "[人民币,,,A,余额,人民币,月,1]>0",
                    "否",
                ),
                expression_rule(
                    S3ExpressionSchema::FiveSegment,
                    "R-1",
                    "[A,余额,人民币,月,1]>0",
                    "是",
                ),
            ],
            ..S3ExecutionCompareConfig::default()
        };
        let mut rows = build_base_comparison(&current, &previous, &config);

        let error = apply_expression_rules(
            &mut rows,
            &current,
            &previous,
            &config,
            S3ExpressionSchema::FiveSegment,
        )
        .unwrap_err()
        .to_string();
        assert!(error.contains("五段式取数键不唯一"));
    }

    #[test]
    fn action_rules_apply_thresholds_and_filter_unavailable_frequency() {
        let current = dataset(
            vec![record(
                "0001",
                "1100000",
                "A",
                "月",
                "1",
                CentralValue::Number(30.0),
            )],
            "2026-08-31",
        );
        let previous = dataset(
            vec![record(
                "0001",
                "1100000",
                "A",
                "月",
                "1",
                CentralValue::Number(20.0),
            )],
            "2026-07-31",
        );
        let mut unavailable = action_rule("R-2", "指标为负数需核实。", "A", "");
        unavailable.insert("适用频度".to_owned(), "季".to_owned());
        let config = S3ExecutionCompareConfig {
            target_unit: "亿元".to_owned(),
            action_rules: vec![
                action_rule("R-1", "指标有变动需说明。", "A", "5"),
                unavailable,
            ],
            ..S3ExecutionCompareConfig::default()
        };
        let mut rows = build_base_comparison(&current, &previous, &config);

        let report = apply_action_rules(&mut rows, &current, &previous, &config).unwrap();

        let row = rows.iter().find(|row| row.record.indicator == "A").unwrap();
        assert_eq!(report.enabled_rules, 2);
        assert_eq!(report.applicable_rules, 1);
        assert_eq!(report.filtered_rules, 1);
        assert_eq!(report.triggered_rules, 1);
        assert!(
            row.need_explanation
                .contains("R-1-指标有变动需说明。-阈值5亿元")
        );
    }

    #[test]
    fn action_rule_matching_negative_current_value_is_reported() {
        let current = dataset(
            vec![record(
                "0001",
                "1100000",
                "A",
                "月",
                "1",
                CentralValue::Number(-1.0),
            )],
            "2026-08-31",
        );
        let previous = dataset(Vec::new(), "2026-07-31");
        let config = S3ExecutionCompareConfig {
            target_unit: "亿元".to_owned(),
            action_rules: vec![action_rule("R-NEG", "指标为负数需核实。", "A", "")],
            ..S3ExecutionCompareConfig::default()
        };
        let mut rows = build_base_comparison(&current, &previous, &config);

        let report = apply_action_rules(&mut rows, &current, &previous, &config).unwrap();

        assert_eq!(report.triggered_rules, 1);
        assert!(rows[0].need_explanation.contains("指标为负数需核实。"));
    }

    #[test]
    fn expression_substitution_uses_python_fifteen_significant_digit_text() {
        assert_eq!(python_cstr(0.00001), "1e-05");
        assert_eq!(python_cstr(1e16), "1e+16");
        assert_eq!(python_cstr(3_204_999_999.999_999_5), "3205000000");
    }

    #[test]
    fn execution_compare_writer_preserves_result_values_and_refuses_overwrite() -> Result<()> {
        use std::io::Read as _;

        let counter = TEMP_OUTPUT_COUNTER.fetch_add(1, Ordering::Relaxed);
        let output = std::env::temp_dir().join(format!(
            "audit-s3-f1-writer-{}-{counter}.xlsx",
            std::process::id()
        ));
        let row = S3ExecutionCompareRow {
            record: record(
                "0001",
                "1100000",
                "A-01",
                "月",
                "1",
                CentralValue::Number(125.5),
            ),
            order_code_out: "0007".to_owned(),
            indicator_name_out: "测试指标输出名".to_owned(),
            previous_value: CentralValue::Number(100.0),
            change: CentralValue::Number(25.5),
            ratio: CentralValue::Number(0.255),
            remark: "需说明".to_owned(),
            need_explanation: "是".to_owned(),
            explanation_content: "测试说明".to_owned(),
            calculation_process: "125.5 - 100.0".to_owned(),
            fill_color: Some(6),
            fill_scope: "row".to_owned(),
        };
        let blank_row = S3ExecutionCompareRow {
            record: record("0001", "1100000", "A-02", "月", "1", CentralValue::Empty),
            order_code_out: "0008".to_owned(),
            indicator_name_out: String::new(),
            previous_value: CentralValue::Empty,
            change: CentralValue::Empty,
            ratio: CentralValue::Text(REMARK_DIV_ZERO.to_owned()),
            remark: String::new(),
            need_explanation: String::new(),
            explanation_content: String::new(),
            calculation_process: String::new(),
            fill_color: None,
            fill_scope: String::new(),
        };

        write_execution_compare_workbook(&output, &[row, blank_row], "亿元", "2026-08-31")?;
        let sheets = read_value_columns(&output, 22)?;
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].name, "2026年8月对比结果");
        assert_eq!(sheets[0].max_row, 3);
        assert_eq!(sheets[0].max_column, 22);
        assert_eq!(cell_text(&sheets[0], 1, 16), "增减额(亿元)");
        assert_eq!(cell_text(&sheets[0], 1, 22), "输出说明");
        assert_eq!(
            sheets[0].values.get("G2"),
            Some(&CellValue::Text("0007".to_owned()))
        );
        assert!(central_value_matches_cell(
            &CentralValue::Number(125.5),
            sheets[0].values.get("N2").unwrap_or(&CellValue::Empty)
        ));
        assert!(central_value_matches_cell(
            &CentralValue::Number(0.255),
            sheets[0].values.get("Q2").unwrap_or(&CellValue::Empty)
        ));
        assert_eq!(
            sheets[0].values.get("V2").unwrap_or(&CellValue::Empty),
            &CellValue::Empty
        );

        let mut package = zip::ZipArchive::new(std::fs::File::open(&output)?)?;
        let mut sheet_xml = String::new();
        package
            .by_name("xl/worksheets/sheet1.xml")?
            .read_to_string(&mut sheet_xml)?;
        let sheet_document = roxmltree::Document::parse(&sheet_xml)?;
        let main_namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        let pane = sheet_document
            .descendants()
            .find(|node| node.tag_name().name() == "pane")
            .context("writer output must contain a freeze pane")?;
        assert_eq!(pane.attribute("topLeftCell"), Some("A2"));
        let auto_filter = sheet_document
            .descendants()
            .find(|node| node.tag_name().name() == "autoFilter")
            .context("writer output must contain an autofilter")?;
        assert_eq!(auto_filter.attribute("ref"), Some("A1:V3"));

        let mut styles_xml = String::new();
        package
            .by_name("xl/styles.xml")?
            .read_to_string(&mut styles_xml)?;
        let styles_document = roxmltree::Document::parse(&styles_xml)?;
        let cell_xfs = styles_document
            .descendants()
            .find(|node| {
                node.tag_name().name() == "cellXfs"
                    && node.tag_name().namespace() == Some(main_namespace)
            })
            .context("writer output must contain cellXfs")?;
        let cell_style = |address: &str| -> Result<roxmltree::Node<'_, '_>> {
            let cell = sheet_document
                .descendants()
                .find(|node| node.tag_name().name() == "c" && node.attribute("r") == Some(address))
                .with_context(|| format!("writer output must contain cell {address}"))?;
            let style_index = cell
                .attribute("s")
                .unwrap_or("0")
                .parse::<usize>()
                .with_context(|| format!("cell {address} style index must be numeric"))?;
            cell_xfs
                .children()
                .filter(|node| node.is_element())
                .nth(style_index)
                .with_context(|| format!("cell {address} style index must be in bounds"))
        };
        let amount_style = cell_style("N2")?;
        let amount_fill_id = amount_style
            .attribute("fillId")
            .context("filled result row must have a fillId")?
            .parse::<usize>()?;
        assert!(amount_fill_id > 0);
        let ratio_style = cell_style("Q2")?;
        let ratio_num_format = ratio_style
            .attribute("numFmtId")
            .context("ratio cell must have a numFmtId")?;
        let ratio_num_format = styles_document
            .descendants()
            .find(|node| {
                node.tag_name().name() == "numFmt"
                    && node.attribute("numFmtId") == Some(ratio_num_format)
            })
            .context("ratio cell must use a registered number format")?;
        assert_eq!(ratio_num_format.attribute("formatCode"), Some("0.00%"));
        let custom_num_format = |address: &str| -> Result<String> {
            let style = cell_style(address)?;
            let num_format_id = style
                .attribute("numFmtId")
                .with_context(|| format!("cell {address} must retain its number format"))?;
            let number_format = styles_document
                .descendants()
                .find(|node| {
                    node.tag_name().name() == "numFmt"
                        && node.attribute("numFmtId") == Some(num_format_id)
                })
                .with_context(|| format!("cell {address} must use a registered number format"))?;
            Ok(number_format
                .attribute("formatCode")
                .unwrap_or_default()
                .to_owned())
        };
        for address in ["N3", "O3", "P3"] {
            assert_eq!(custom_num_format(address)?, "0.00");
        }
        assert_eq!(custom_num_format("Q3")?, "0.00%");
        assert!(write_execution_compare_workbook(&output, &[], "亿元", "2026-08-31").is_err());
        let _ = std::fs::remove_file(output);
        Ok(())
    }

    #[test]
    #[ignore = "requires AUDIT_S3_F1_COMMON_CONFIG and AUDIT_S3_F1_FEATURE_CONFIG"]
    fn loads_real_s3_f1_configuration_without_mutating_it() -> Result<()> {
        let common = std::env::var_os("AUDIT_S3_F1_COMMON_CONFIG")
            .context("AUDIT_S3_F1_COMMON_CONFIG is required")?;
        let feature = std::env::var_os("AUDIT_S3_F1_FEATURE_CONFIG")
            .context("AUDIT_S3_F1_FEATURE_CONFIG is required")?;
        let config = load_execution_compare_config(common, feature)?;
        assert!(!config.alert_bands.is_empty());
        assert!(!config.indicator_references.is_empty());
        assert!(config.action_rules.len() >= 1_000);
        assert!(config.expression_rules.len() >= 1_000);
        assert!(
            config
                .expression_rules
                .iter()
                .any(|rule| rule.schema == S3ExpressionSchema::FiveSegment)
        );
        assert!(
            config
                .expression_rules
                .iter()
                .any(|rule| rule.schema == S3ExpressionSchema::LegacyEightSegment)
        );
        let five_rules =
            compile_expression_rules(&config, S3ExpressionSchema::FiveSegment, "2026-08-31")?;
        let eight_rules = compile_expression_rules(
            &config,
            S3ExpressionSchema::LegacyEightSegment,
            "2026-08-31",
        )?;
        assert!(!five_rules.is_empty());
        assert!(!eight_rules.is_empty());
        let (action_rules, enabled_actions) = prepare_action_rules(
            &config,
            &dataset(Vec::new(), "2026-08-31"),
            &dataset(Vec::new(), "2026-07-31"),
        )?;
        assert_eq!(enabled_actions, 4_246);
        eprintln!(
            "loaded S3-F1 config: alerts={}, indicator refs={}, action rules={} (enabled={}, applicable with empty data={}), expression rules={}, compilable five/eight={}/{}",
            config.alert_bands.len(),
            config.indicator_references.len(),
            config.action_rules.len(),
            enabled_actions,
            action_rules.len(),
            config.expression_rules.len(),
            five_rules.len(),
            eight_rules.len()
        );
        Ok(())
    }
}
