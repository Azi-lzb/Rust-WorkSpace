//! Read-only structural and runtime validation for the S2 configuration workbook.
//!
//! This mirrors the source product's `config_check.check_s2_config` contract. It
//! reads stored values through the shared Calamine adapter and delegates runtime
//! parsing to `s2_config::load_config`; it never writes the source workbook.

use crate::s2_config::{cell_code, cell_text, cell_yes, load_config, severity};
use excel_reader::{read_value_columns_with_iso_dates, CellValue, SheetValues};
use serde::Serialize;
use std::collections::{HashMap, HashSet};
use std::path::Path;

const REQUIRED_SHEETS: &[&str] = &[
    "指标参照",
    "机构参照",
    "环比策略",
    "校验规则",
    "外部核对规则",
    "通用设置",
];
const DATA_TYPES: &[&str] = &["余额", "累发", "百分数", "个数", "文字"];
const SEVERITIES: &[&str] = &["", "硬性", "软性"];
const LEGACY_SEVERITIES: &[(&str, &str)] = &[
    ("错误", "硬性"),
    ("严重", "硬性"),
    ("核实", "软性"),
    ("提示", "软性"),
    ("关注", "软性"),
];
const CHECK_DEFINITIONS: &[(&str, &str)] = &[
    ("file_read", "配置文件可读取"),
    ("required_sheets", "必需工作表齐全"),
    ("indicators", "指标参照字段与记录"),
    ("period_strategies", "环比策略区间与指标引用"),
    ("runtime_parse", "全部配置可被正式运行时解析"),
];
const CHECK_LIMITATIONS: &[&str] = &[
    "此项检查只读取 S2 配置工作簿，不读取本期、上期报表或大集中外部数据文件。",
    "此项检查不实际执行审核规则，也不判断报表结果是否符合业务预期。",
    "检查通过表示配置结构和内容可由当前运行时读取，不代表真实数据审核已经通过。",
];

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct S2ConfigCheckIssue {
    pub level: String,
    pub message: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub sheet: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub row: Option<usize>,
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct S2ConfigSheetStat {
    pub name: String,
    pub rows: usize,
    pub columns: usize,
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct S2ConfigCheckStep {
    pub id: String,
    pub name: String,
    pub status: String,
    pub detail: String,
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct S2ConfigCheckReport {
    pub path: String,
    pub passed: bool,
    pub blocking_errors: usize,
    pub warnings: usize,
    pub issues: Vec<S2ConfigCheckIssue>,
    pub sheet_stats: Vec<S2ConfigSheetStat>,
    pub indicator_count: usize,
    pub strategy_count: usize,
    pub message: String,
    pub checks: Vec<S2ConfigCheckStep>,
    pub limitations: Vec<String>,
}

#[derive(Clone, Debug)]
struct StrategyRow {
    id: String,
    lower: Option<f64>,
    upper: Option<f64>,
    disabled: bool,
    data_type: String,
    row: usize,
}

/// Inspect an S2 configuration workbook without changing it.
pub fn check_s2_config(path: impl AsRef<Path>) -> S2ConfigCheckReport {
    let target = path.as_ref();
    let mut report = S2ConfigCheckReport {
        path: target.to_string_lossy().into_owned(),
        passed: false,
        blocking_errors: 0,
        warnings: 0,
        issues: Vec::new(),
        sheet_stats: Vec::new(),
        indicator_count: 0,
        strategy_count: 0,
        message: String::new(),
        checks: CHECK_DEFINITIONS
            .iter()
            .map(|(id, name)| S2ConfigCheckStep {
                id: (*id).to_owned(),
                name: (*name).to_owned(),
                status: "not_run".to_owned(),
                detail: "未执行（前置检查未完成）。".to_owned(),
            })
            .collect(),
        limitations: CHECK_LIMITATIONS
            .iter()
            .map(|item| (*item).to_owned())
            .collect(),
    };

    if !target.is_file() {
        push_issue(
            &mut report,
            "blocking",
            format!("配置文件不存在：{}", target.display()),
            None,
            None,
        );
        set_check(
            &mut report,
            "file_read",
            "failed",
            "配置文件不存在，无法读取。",
        );
        report.message = blocked_message(report.blocking_errors);
        return report;
    }

    let sheets = match read_value_columns_with_iso_dates(target, usize::MAX) {
        Ok(sheets) => sheets,
        Err(error) => {
            push_issue(
                &mut report,
                "blocking",
                format!("配置文件无法读取：{error}"),
                None,
                None,
            );
            set_check(
                &mut report,
                "file_read",
                "failed",
                "配置文件无法打开，详见下方问题明细。",
            );
            report.message = blocked_message(report.blocking_errors);
            return report;
        }
    };

    set_check(
        &mut report,
        "file_read",
        "passed",
        "配置工作簿已成功打开，并以只读方式检查。",
    );
    let sheet_names = sheets
        .iter()
        .map(|sheet| sheet.name.as_str())
        .collect::<HashSet<_>>();
    let required_start = report.issues.len();
    for required in REQUIRED_SHEETS {
        if !sheet_names.contains(required) {
            push_issue(
                &mut report,
                "blocking",
                format!("缺少工作表：{required}"),
                None,
                None,
            );
        }
    }
    report.sheet_stats = sheets.iter().map(sheet_stat).collect();
    finish_check(
        &mut report,
        "required_sheets",
        required_start,
        format!("已确认 {} 张必需工作表存在", REQUIRED_SHEETS.len()),
    );
    if report.blocking_errors > 0 {
        report.message = blocked_message(report.blocking_errors);
        return report;
    }

    let referenced_groups = check_indicators(&mut report, &sheets);
    check_period_strategies(&mut report, &sheets, &referenced_groups);
    check_runtime_parse(&mut report, &sheets, target);

    report.passed = report.blocking_errors == 0;
    report.message = if report.passed {
        "审核配置检查通过，可执行报表规则。".to_owned()
    } else {
        blocked_message(report.blocking_errors)
    };
    report
}

/// Format the structured check report for the existing UI text contract.
pub fn format_config_report(report: &S2ConfigCheckReport) -> String {
    let mut lines = vec![report.message.clone()];
    lines.push(format!("配置文件：{}", report.path));
    lines.push(format!(
        "指标参照：{} 条；环比策略：{} 条",
        report.indicator_count, report.strategy_count
    ));
    lines.push("实际检查项目：".to_owned());
    for check in &report.checks {
        let label = match check.status.as_str() {
            "passed" => "通过",
            "warning" => "有警告",
            "failed" => "未通过",
            _ => "未检查",
        };
        lines.push(format!("[{label}] {}：{}", check.name, check.detail));
    }
    if !report.limitations.is_empty() {
        lines.push("检查范围说明：".to_owned());
        lines.extend(report.limitations.iter().map(|item| format!("- {item}")));
    }
    for stat in &report.sheet_stats {
        lines.push(format!(
            "工作表 {}：{} 行，{} 列",
            stat.name, stat.rows, stat.columns
        ));
    }
    for issue in &report.issues {
        let mut location = issue.sheet.clone().unwrap_or_default();
        if let Some(row) = issue.row {
            location.push_str(&format!(" 第{row}行"));
        }
        let prefix = if issue.level == "blocking" {
            "阻断"
        } else {
            "警告"
        };
        lines.push(format!(
            "[{prefix}] {}{}",
            if location.is_empty() {
                String::new()
            } else {
                format!("{location}：")
            },
            issue.message
        ));
    }
    lines.join("\n")
}

fn check_indicators(report: &mut S2ConfigCheckReport, sheets: &[SheetValues]) -> HashSet<String> {
    let start = report.issues.len();
    let indicator_sheet = find_sheet(sheets, "指标参照").expect("required sheet checked");
    let indicator_rows = dense_rows(indicator_sheet);
    let headers = header_map(&indicator_rows);
    let mut inline_units = ["源数据单位", "大集中数据单位"]
        .into_iter()
        .filter(|name| headers.contains_key(*name))
        .collect::<Vec<_>>();
    inline_units.sort_unstable();
    if !inline_units.is_empty() {
        push_issue(
            report,
            "blocking",
            format!(
                "单位统一配置应放在本工作簿的“通用设置”表；请从“指标参照”删除旧字段：{}",
                inline_units.join("、")
            ),
            Some("指标参照"),
            None,
        );
    }
    for required in [
        "指标代码",
        "指标名称",
        "表单代码",
        "数据属性",
        "环比启用",
        "环比策略组",
        "最小变动值",
        "禁用",
    ] {
        if !headers.contains_key(required) {
            push_issue(
                report,
                "blocking",
                format!("“指标参照”缺少字段：{required}"),
                Some("指标参照"),
                None,
            );
        }
    }

    let mut seen_codes = HashSet::new();
    let mut referenced_groups = HashSet::new();
    for (row_index, row) in indicator_rows.iter().enumerate().skip(1) {
        let row_number = row_index + 1;
        let code = cell_code(cell_at(row, &headers, "指标代码"));
        if code.is_empty() {
            continue;
        }
        report.indicator_count += 1;
        if !seen_codes.insert(code.clone()) {
            push_issue(
                report,
                "blocking",
                format!("指标代码重复：{code}"),
                Some("指标参照"),
                Some(row_number),
            );
        }
        let data_type = cell_text(cell_at(row, &headers, "数据属性"));
        if !DATA_TYPES.contains(&data_type.as_str()) {
            push_issue(
                report,
                "blocking",
                format!("指标 {code} 数据属性无效：{}", empty_label(&data_type)),
                Some("指标参照"),
                Some(row_number),
            );
        }
        let enabled_raw = cell_at(row, &headers, "环比启用");
        let enabled_text = cell_text(enabled_raw);
        if !["是", "否"].contains(&enabled_text.as_str()) {
            push_issue(
                report,
                "blocking",
                format!(
                    "指标 {code} 环比启用必须为“是”或“否”：{}",
                    cell_repr(enabled_raw)
                ),
                Some("指标参照"),
                Some(row_number),
            );
        }
        let disabled_raw = cell_at(row, &headers, "禁用");
        let disabled_text = cell_text(disabled_raw);
        if !["是", "否", ""].contains(&disabled_text.as_str()) {
            push_issue(
                report,
                "blocking",
                format!(
                    "指标 {code} 禁用必须为“是”或“否”：{}",
                    cell_repr(disabled_raw)
                ),
                Some("指标参照"),
                Some(row_number),
            );
        }
        if !cell_yes(enabled_raw) {
            continue;
        }
        let group = cell_text(cell_at(row, &headers, "环比策略组"));
        let min_raw = cell_at(row, &headers, "最小变动值");
        if data_type == "文字" {
            if group != "不适用" || cell_text(min_raw) != "不适用" {
                push_issue(
                    report,
                    "blocking",
                    format!("文字指标 {code} 的环比策略组和最小变动值必须均为“不适用”"),
                    Some("指标参照"),
                    Some(row_number),
                );
            }
        } else {
            if group.is_empty() || group == "不适用" {
                push_issue(
                    report,
                    "blocking",
                    format!("数值指标 {code} 未配置有效环比策略组"),
                    Some("指标参照"),
                    Some(row_number),
                );
            } else {
                referenced_groups.insert(group);
            }
            let min_value = checker_number(min_raw);
            if !min_value.is_some_and(|value| value.is_finite() && value >= 0.0) {
                push_issue(
                    report,
                    "blocking",
                    format!(
                        "数值指标 {code} 最小变动值必须为非负数：{}",
                        cell_repr(min_raw)
                    ),
                    Some("指标参照"),
                    Some(row_number),
                );
            }
        }
    }
    finish_check(
        report,
        "indicators",
        start,
        format!(
            "已检查 {} 条指标记录的必需字段、重复代码、属性、启用值及环比设置",
            report.indicator_count
        ),
    );
    referenced_groups
}

fn check_period_strategies(
    report: &mut S2ConfigCheckReport,
    sheets: &[SheetValues],
    referenced_groups: &HashSet<String>,
) {
    let start = report.issues.len();
    let strategy_sheet = find_sheet(sheets, "环比策略").expect("required sheet checked");
    let rows = dense_rows(strategy_sheet);
    let headers = header_map(&rows);
    for required in [
        "策略组",
        "规则编号",
        "数据属性",
        "下限",
        "上限",
        "审核级别",
        "分类说明",
        "禁用",
    ] {
        if !headers.contains_key(required) {
            push_issue(
                report,
                "blocking",
                format!("“环比策略”缺少字段：{required}"),
                Some("环比策略"),
                None,
            );
        }
    }
    let mut group_positions = HashMap::<String, usize>::new();
    let mut groups = Vec::<(String, Vec<StrategyRow>)>::new();
    let mut seen_pairs = HashSet::new();
    for (row_index, row) in rows.iter().enumerate().skip(1) {
        let group = cell_text(cell_at(row, &headers, "策略组"));
        let rule_id = cell_text(cell_at(row, &headers, "规则编号"));
        if group.is_empty() && rule_id.is_empty() {
            continue;
        }
        report.strategy_count += 1;
        let row_number = row_index + 1;
        if group.is_empty() || rule_id.is_empty() {
            push_issue(
                report,
                "blocking",
                "策略组与规则编号不能为空",
                Some("环比策略"),
                Some(row_number),
            );
            continue;
        }
        if !seen_pairs.insert((group.clone(), rule_id.clone())) {
            push_issue(
                report,
                "blocking",
                format!("策略组+规则编号重复：{group}/{rule_id}"),
                Some("环比策略"),
                Some(row_number),
            );
        }
        let data_type = cell_text(cell_at(row, &headers, "数据属性"));
        if !["余额", "累发", "百分数", "个数"].contains(&data_type.as_str()) {
            push_issue(
                report,
                "blocking",
                format!(
                    "策略 {group}/{rule_id} 数据属性无效：{}",
                    empty_label(&data_type)
                ),
                Some("环比策略"),
                Some(row_number),
            );
        }
        let lower = checker_number(cell_at(row, &headers, "下限"));
        let upper = checker_number(cell_at(row, &headers, "上限"));
        if lower.is_some_and(|value| !value.is_finite())
            || upper.is_some_and(|value| !value.is_finite())
        {
            push_issue(
                report,
                "blocking",
                format!("策略 {group}/{rule_id} 边界必须是有限数字或空白"),
                Some("环比策略"),
                Some(row_number),
            );
        }
        if lower
            .zip(upper)
            .is_some_and(|(lower, upper)| upper <= lower)
        {
            push_issue(
                report,
                "blocking",
                format!("策略 {group}/{rule_id} 上限必须大于下限"),
                Some("环比策略"),
                Some(row_number),
            );
        }
        let severity_text = severity(&cell_text(cell_at(row, &headers, "审核级别")));
        if !SEVERITIES.contains(&severity_text.as_str()) {
            push_issue(
                report,
                "blocking",
                format!("策略 {group}/{rule_id} 审核级别无效：{severity_text}"),
                Some("环比策略"),
                Some(row_number),
            );
        }
        let disabled_raw = cell_at(row, &headers, "禁用");
        let disabled_text = cell_text(disabled_raw);
        if !["是", "否", ""].contains(&disabled_text.as_str()) {
            push_issue(
                report,
                "blocking",
                format!(
                    "策略 {group}/{rule_id} 禁用必须为“是”或“否”：{}",
                    cell_repr(disabled_raw)
                ),
                Some("环比策略"),
                Some(row_number),
            );
        }
        let position = *group_positions.entry(group.clone()).or_insert_with(|| {
            groups.push((group.clone(), Vec::new()));
            groups.len() - 1
        });
        groups[position].1.push(StrategyRow {
            id: rule_id,
            lower,
            upper,
            disabled: cell_yes(disabled_raw),
            data_type,
            row: row_number,
        });
    }

    for (group, group_rows) in &groups {
        let mut types = group_rows
            .iter()
            .map(|row| row.data_type.as_str())
            .collect::<HashSet<_>>()
            .into_iter()
            .collect::<Vec<_>>();
        types.sort_unstable();
        if types.len() != 1 {
            push_issue(
                report,
                "blocking",
                format!(
                    "环比策略组 {group} 的数据属性必须一致：{}",
                    types.join("、")
                ),
                Some("环比策略"),
                None,
            );
        }
        check_ranges(report, group_rows, group);
    }

    let group_names = group_positions.keys().cloned().collect::<HashSet<_>>();
    let mut missing_groups = referenced_groups
        .difference(&group_names)
        .cloned()
        .collect::<Vec<_>>();
    missing_groups.sort_unstable();
    for group in missing_groups {
        push_issue(
            report,
            "blocking",
            format!("指标参照引用了不存在的环比策略组：{group}"),
            Some("指标参照"),
            None,
        );
    }
    let mut unused_groups = group_positions
        .keys()
        .filter(|group| !referenced_groups.contains(*group))
        .cloned()
        .collect::<Vec<_>>();
    unused_groups.sort_unstable();
    for group in unused_groups {
        push_issue(
            report,
            "warning",
            format!("环比策略组 {group} 当前没有被任何启用指标引用"),
            Some("环比策略"),
            None,
        );
    }

    finish_check(
        report,
        "period_strategies",
        start,
        format!(
            "已检查 {} 条策略记录的编号、区间覆盖/空档/重叠及指标组引用",
            report.strategy_count
        ),
    );
}

fn check_runtime_parse(report: &mut S2ConfigCheckReport, sheets: &[SheetValues], path: &Path) {
    let start = report.issues.len();
    warn_legacy_severity_values(report, sheets);
    match load_config(path) {
        Err(error) => push_issue(
            report,
            "blocking",
            format!("配置无法进入运行时解析：{error}"),
            None,
            None,
        ),
        Ok(config) => {
            for warning in config.warnings {
                let level = if ["已跳过", "不存在", "无效", "暂不支持"]
                    .iter()
                    .any(|token| warning.contains(token))
                {
                    "blocking"
                } else {
                    "warning"
                };
                push_issue(report, level, warning, None, None);
            }
        }
    }
    finish_check(
        report,
        "runtime_parse",
        start,
        "指标参照、机构参照、校验规则、外部核对规则和通用设置均已交给正式配置解析器检查",
    );
}

fn warn_legacy_severity_values(report: &mut S2ConfigCheckReport, sheets: &[SheetValues]) {
    for (sheet_name, header_name) in [
        ("环比策略", "审核级别"),
        ("校验规则", "级别"),
        ("外部核对规则", "级别"),
    ] {
        let Some(sheet) = find_sheet(sheets, sheet_name) else {
            continue;
        };
        let rows = dense_rows(sheet);
        let headers = header_map(&rows);
        if !headers.contains_key(header_name) {
            continue;
        }
        for (row_index, row) in rows.iter().enumerate().skip(1) {
            let row_number = row_index + 1;
            let raw = cell_text(cell_at(row, &headers, header_name));
            if let Some((_, mapped)) = LEGACY_SEVERITIES.iter().find(|(old, _)| *old == raw) {
                push_issue(
                    report,
                    "warning",
                    format!(
                        "旧审核级别“{raw}”兼容映射为“{mapped}”；建议人工改为“{mapped}”。建议人工复核，程序未自动修改配置。"
                    ),
                    Some(sheet_name),
                    Some(row_number),
                );
            }
        }
    }
}

fn check_ranges(report: &mut S2ConfigCheckReport, rows: &[StrategyRow], group: &str) {
    let mut enabled = rows
        .iter()
        .filter(|row| !row.disabled)
        .cloned()
        .collect::<Vec<_>>();
    if enabled.is_empty() {
        push_issue(
            report,
            "blocking",
            format!("环比策略组 {group} 没有启用的策略行"),
            Some("环比策略"),
            None,
        );
        return;
    }
    enabled.sort_by(|left, right| compare_lower(left.lower, right.lower));
    if enabled[0].lower.is_some() {
        push_issue(
            report,
            "blocking",
            format!("环比策略组 {group} 左侧没有覆盖负无穷（首行下限应为空）"),
            Some("环比策略"),
            Some(enabled[0].row),
        );
    }
    if enabled[enabled.len() - 1].upper.is_some() {
        push_issue(
            report,
            "blocking",
            format!("环比策略组 {group} 右侧没有覆盖正无穷（末行上限应为空）"),
            Some("环比策略"),
            Some(enabled[enabled.len() - 1].row),
        );
    }
    for pair in enabled.windows(2) {
        let left = &pair[0];
        let right = &pair[1];
        match (left.upper, right.lower) {
            (None, _) | (_, None) => push_issue(
                report,
                "blocking",
                format!("环比策略组 {group} 存在重叠：{} 与 {}", left.id, right.id),
                Some("环比策略"),
                Some(right.row),
            ),
            (Some(left_upper), Some(right_lower)) if left_upper < right_lower => push_issue(
                report,
                "blocking",
                format!("环比策略组 {group} 存在空档：{} 与 {}", left.id, right.id),
                Some("环比策略"),
                Some(right.row),
            ),
            (Some(left_upper), Some(right_lower)) if left_upper > right_lower => push_issue(
                report,
                "blocking",
                format!("环比策略组 {group} 存在重叠：{} 与 {}", left.id, right.id),
                Some("环比策略"),
                Some(right.row),
            ),
            _ => {}
        }
    }
}

fn compare_lower(left: Option<f64>, right: Option<f64>) -> std::cmp::Ordering {
    match (left, right) {
        (None, None) => std::cmp::Ordering::Equal,
        (None, Some(_)) => std::cmp::Ordering::Less,
        (Some(_), None) => std::cmp::Ordering::Greater,
        (Some(left), Some(right)) => left
            .partial_cmp(&right)
            .unwrap_or(std::cmp::Ordering::Equal),
    }
}

fn sheet_stat(sheet: &SheetValues) -> S2ConfigSheetStat {
    let (max_row, max_column) = inferred_extent(sheet);
    S2ConfigSheetStat {
        name: sheet.name.clone(),
        rows: max_row.saturating_sub(1),
        columns: max_column.max(1),
    }
}

fn inferred_extent(sheet: &SheetValues) -> (usize, usize) {
    let mut max_row = sheet.max_row;
    let mut max_column = sheet.max_column;
    for address in sheet.values.keys() {
        if let Some((row, column)) = parse_coordinate(address) {
            max_row = max_row.max(row);
            max_column = max_column.max(column);
        }
    }
    (max_row, max_column)
}

fn dense_rows(sheet: &SheetValues) -> Vec<Vec<CellValue>> {
    // Validation skips rows whose cells are all empty, so only materialize the
    // populated extent. Sheet dimensions are retained separately for UI stats;
    // they must not cause a large blank declared range to allocate a dense grid.
    let (max_row, max_column) = populated_extent(sheet);
    if max_row == 0 || max_column == 0 {
        return Vec::new();
    }
    let mut rows = vec![vec![CellValue::Empty; max_column]; max_row];
    for (address, value) in &sheet.values {
        if let Some((row, column)) = parse_coordinate(address) {
            if row > 0 && column > 0 && row <= max_row && column <= max_column {
                rows[row - 1][column - 1] = value.clone();
            }
        }
    }
    rows
}

fn populated_extent(sheet: &SheetValues) -> (usize, usize) {
    let mut max_row = 0;
    let mut max_column = 0;
    for address in sheet.values.keys() {
        if let Some((row, column)) = parse_coordinate(address) {
            max_row = max_row.max(row);
            max_column = max_column.max(column);
        }
    }
    (max_row, max_column)
}

fn parse_coordinate(address: &str) -> Option<(usize, usize)> {
    let mut column = 0_usize;
    let mut row_start = 0;
    for (index, ch) in address.char_indices() {
        if ch.is_ascii_alphabetic() {
            column = column
                .checked_mul(26)?
                .checked_add((ch.to_ascii_uppercase() as usize) - ('A' as usize) + 1)?;
            row_start = index + ch.len_utf8();
        } else {
            break;
        }
    }
    if row_start == 0 || row_start >= address.len() {
        return None;
    }
    let row = address[row_start..].parse::<usize>().ok()?;
    (row > 0 && column > 0).then_some((row, column))
}

fn find_sheet<'a>(sheets: &'a [SheetValues], name: &str) -> Option<&'a SheetValues> {
    sheets.iter().find(|sheet| sheet.name == name)
}

fn header_map(rows: &[Vec<CellValue>]) -> HashMap<String, usize> {
    let mut headers = HashMap::new();
    if let Some(header) = rows.first() {
        for (index, value) in header.iter().enumerate() {
            let name = cell_text(Some(value));
            if !name.is_empty() {
                headers.insert(name, index);
            }
        }
    }
    headers
}

fn cell_at<'a>(
    row: &'a [CellValue],
    headers: &HashMap<String, usize>,
    name: &str,
) -> Option<&'a CellValue> {
    headers.get(name).and_then(|index| row.get(*index))
}

fn checker_number(value: Option<&CellValue>) -> Option<f64> {
    match value {
        None | Some(CellValue::Empty) => None,
        Some(CellValue::Text(text)) if text.is_empty() => None,
        Some(CellValue::Number(value)) => Some(*value),
        Some(CellValue::Integer(value)) => Some(*value as f64),
        Some(CellValue::Boolean(value)) => Some(f64::from(*value)),
        Some(CellValue::Text(value)) => value.trim().parse().ok(),
        Some(
            CellValue::DateTimeSerial(_)
            | CellValue::DateTimeIso(_)
            | CellValue::DurationIso(_)
            | CellValue::Error(_),
        ) => None,
    }
}

fn cell_repr(value: Option<&CellValue>) -> String {
    match value {
        None | Some(CellValue::Empty) => "None".to_owned(),
        Some(CellValue::Text(value)) | Some(CellValue::Error(value)) => python_string_repr(value),
        Some(CellValue::Number(value) | CellValue::DateTimeSerial(value)) => {
            crate::s2_expression::python_float_repr(*value)
        }
        Some(CellValue::Integer(value)) => value.to_string(),
        Some(CellValue::Boolean(value)) => if *value { "True" } else { "False" }.to_owned(),
        Some(CellValue::DateTimeIso(value)) => crate::s2_config::python_datetime_repr(value),
        Some(CellValue::DurationIso(value)) => python_string_repr(value),
    }
}

fn python_string_repr(value: &str) -> String {
    let quote = if value.contains('\'') && !value.contains('"') {
        '"'
    } else {
        '\''
    };
    let mut result = String::with_capacity(value.len() + 2);
    result.push(quote);
    for character in value.chars() {
        match character {
            '\\' => result.push_str("\\\\"),
            '\n' => result.push_str("\\n"),
            '\r' => result.push_str("\\r"),
            '\t' => result.push_str("\\t"),
            character if character == quote => {
                result.push('\\');
                result.push(character);
            }
            character if character.is_control() => {
                result.push_str(&format!("\\u{{{:x}}}", character as u32));
            }
            character => result.push(character),
        }
    }
    result.push(quote);
    result
}

fn empty_label(value: &str) -> &str {
    if value.is_empty() {
        "空白"
    } else {
        value
    }
}

fn push_issue(
    report: &mut S2ConfigCheckReport,
    level: &str,
    message: impl Into<String>,
    sheet: Option<&str>,
    row: Option<usize>,
) {
    report.issues.push(S2ConfigCheckIssue {
        level: level.to_owned(),
        message: message.into(),
        sheet: sheet.map(str::to_owned),
        row,
    });
    if level == "blocking" {
        report.blocking_errors += 1;
    } else {
        report.warnings += 1;
    }
}

fn set_check(report: &mut S2ConfigCheckReport, id: &str, status: &str, detail: &str) {
    let check = report
        .checks
        .iter_mut()
        .find(|check| check.id == id)
        .expect("known S2 configuration check id");
    check.status = status.to_owned();
    check.detail = detail.to_owned();
}

fn finish_check(
    report: &mut S2ConfigCheckReport,
    id: &str,
    issue_start: usize,
    success_detail: impl Into<String>,
) {
    let issues = &report.issues[issue_start..];
    let blocking = issues
        .iter()
        .filter(|issue| issue.level == "blocking")
        .count();
    let warnings = issues
        .iter()
        .filter(|issue| issue.level == "warning")
        .count();
    let success_detail = success_detail.into();
    if blocking > 0 {
        set_check(
            report,
            id,
            "failed",
            &format!("发现 {blocking} 个阻断问题，详见下方问题明细。"),
        );
    } else if warnings > 0 {
        set_check(
            report,
            id,
            "warning",
            &format!("{success_detail}；另有 {warnings} 条警告，详见下方问题明细。"),
        );
    } else {
        set_check(report, id, "passed", &success_detail);
    }
}

fn blocked_message(blocking_errors: usize) -> String {
    format!("审核配置存在 {blocking_errors} 个阻断错误，不能执行。")
}
