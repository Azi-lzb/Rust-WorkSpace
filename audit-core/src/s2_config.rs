//! Rust loader for the S2 report-collection configuration workbook.
//!
//! Workbook cells come from the shared Calamine reader. This module parses
//! configuration values only; it does not write or repair configuration files.

use crate::{
    s2_expression::python_float_repr,
    s2_external::{
        S2ExternalConfig, S2ExternalIndicatorConfig, S2ExternalInstitutionProfile, S2ExternalRule,
        S2ExternalUnitSettings,
    },
    s2_import::{
        S2ImportConfig, S2IndicatorConfig, S2InstitutionProfile, S2SourceRef, S2UnitSettings,
    },
    s2_period::{S2PeriodBandConfig, S2PeriodConfig, S2PeriodIndicatorConfig},
    s2_rule_engine::{S2Rule, S2RuleEngineConfig},
};
use anyhow::{anyhow, bail, Result};
use chrono::{Datelike, NaiveDateTime, Timelike};
use excel_reader::{read_value_columns_with_iso_dates, CellValue, SheetValues};
use serde::Serialize;
use std::{
    collections::{HashMap, HashSet},
    fs,
    path::Path,
};

const GENERAL_SETTINGS_SHEET: &str = "通用设置";
const GENERAL_SETTINGS_FIELDS: &[&str] = &[
    "源数据单位",
    "外部文件单位",
    "输出文件单位",
    "换算范围",
    "是否隐藏无变动指标",
    "是否隐藏与大集中一致指标",
    "文字变动默认级别",
    "本期缺失默认级别",
    "上期缺失默认级别",
];
const EXTERNAL_RULE_HEADERS: &[&str] = &[
    "规则编号",
    "指标代码",
    "比较方式",
    "外部指标名称",
    "容差",
    "级别",
    "启用",
];
const EXTERNAL_COMPARISON_METHODS: &[&str] =
    &["等于", "不等于", "大于", "大于等于", "小于", "小于等于"];

const INDICATOR_ALIASES: &[(&str, &str)] = &[
    ("指标代码", "code"),
    ("指标名称", "name"),
    ("表单代码", "form"),
    ("数据属性", "type"),
    ("环比启用", "period_enabled"),
    ("环比策略组", "period_group"),
    ("最小变动值", "min_change"),
    ("禁用", "disabled"),
    ("备注", "note"),
];
const EXTERNAL_RULE_ALIASES: &[(&str, &str)] = &[
    ("规则编号", "id"),
    ("指标代码", "indicator"),
    ("外部指标名称", "external_name"),
    ("比较方式", "method"),
    ("容差", "tolerance"),
    ("级别", "severity"),
    ("启用", "enabled"),
];
const INSTITUTION_ALIASES: &[(&str, &str)] = &[
    ("机构名称", "name"),
    ("社会信用代码", "code"),
    ("机构代码", "code"),
    ("机构类别", "type"),
    ("承接行", "branch"),
    ("地区", "region"),
    ("归属行", "region"),
    ("报表项目", "report_item"),
    ("禁用", "disabled"),
];
const BAND_ALIASES: &[(&str, &str)] = &[
    ("规则编号", "id"),
    ("下限", "lower"),
    ("上限", "upper"),
    ("分类说明", "desc"),
    ("级别", "severity"),
    ("禁用", "disabled"),
];
const PERIOD_ALIASES: &[(&str, &str)] = &[
    ("策略组", "group"),
    ("规则编号", "id"),
    ("数据属性", "type"),
    ("下限", "lower"),
    ("上限", "upper"),
    ("审核级别", "severity"),
    ("级别", "severity"),
    ("分类说明", "desc"),
    ("禁用", "disabled"),
];
const RULE_ALIASES: &[(&str, &str)] = &[
    ("规则编号", "id"),
    ("类型", "type"),
    ("描述", "desc"),
    ("规则内容", "expression"),
    ("Thd值(万元)", "threshold"),
    ("Thd值", "threshold"),
    ("阈值(万元)", "threshold"),
    ("取反", "invert"),
    ("取反标识", "invert"),
    ("级别", "severity"),
    ("禁用", "disabled"),
    ("备注", "note"),
];

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct S2ConfigUnitSettings {
    pub source_unit: String,
    pub external_file_unit: String,
    pub output_file_unit: String,
}

impl Default for S2ConfigUnitSettings {
    fn default() -> Self {
        Self {
            source_unit: "元".to_owned(),
            external_file_unit: "元".to_owned(),
            output_file_unit: "元".to_owned(),
        }
    }
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct S2ConfigIndicator {
    pub indicator_code: String,
    pub indicator_name: String,
    pub form_code: String,
    pub data_type: String,
    pub disabled: bool,
    pub note: String,
    pub period_enabled: bool,
    pub period_strategy_group: String,
    pub min_change_value: Option<f64>,
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct S2ConfigInstitution {
    pub institution_id: String,
    pub institution_name: String,
    pub social_credit_code: String,
    pub institution_type: String,
    pub handling_branch: String,
    pub region: String,
    pub report_item: String,
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct S2AuditConfig {
    pub indicators: HashMap<String, S2ConfigIndicator>,
    pub institutions_by_id: HashMap<String, S2ConfigInstitution>,
    pub institutions_by_name: HashMap<String, S2ConfigInstitution>,
    pub bands: Vec<S2PeriodBandConfig>,
    pub external_rules: Vec<S2ExternalRule>,
    pub rules: Vec<S2Rule>,
    pub warnings: Vec<String>,
    pub unit_settings: S2ConfigUnitSettings,
    pub hide_unchanged_period_rows: bool,
    pub hide_matching_external_rows: bool,
    pub text_change_severity: String,
    pub current_missing_severity: String,
    pub previous_missing_severity: String,
}

impl Default for S2AuditConfig {
    fn default() -> Self {
        Self {
            indicators: HashMap::new(),
            institutions_by_id: HashMap::new(),
            institutions_by_name: HashMap::new(),
            bands: Vec::new(),
            external_rules: Vec::new(),
            rules: Vec::new(),
            warnings: Vec::new(),
            unit_settings: S2ConfigUnitSettings::default(),
            hide_unchanged_period_rows: false,
            hide_matching_external_rows: false,
            text_change_severity: "硬性".to_owned(),
            current_missing_severity: "硬性".to_owned(),
            previous_missing_severity: "硬性".to_owned(),
        }
    }
}

impl S2AuditConfig {
    pub fn import_config(&self) -> S2ImportConfig {
        S2ImportConfig {
            indicators: self
                .indicators
                .iter()
                .map(|(code, item)| {
                    (
                        code.clone(),
                        S2IndicatorConfig {
                            indicator_name: item.indicator_name.clone(),
                            form_code: item.form_code.clone(),
                            data_type: item.data_type.clone(),
                        },
                    )
                })
                .collect(),
            institutions_by_id: self
                .institutions_by_id
                .iter()
                .map(|(code, item)| (code.clone(), import_profile(item)))
                .collect(),
            institutions_by_name: self
                .institutions_by_name
                .iter()
                .map(|(name, item)| (name.clone(), import_profile(item)))
                .collect(),
            unit_settings: S2UnitSettings {
                source_unit: self.unit_settings.source_unit.clone(),
                output_file_unit: self.unit_settings.output_file_unit.clone(),
            },
        }
    }

    pub fn period_config(&self) -> S2PeriodConfig {
        S2PeriodConfig {
            indicators: self
                .indicators
                .iter()
                .map(|(code, item)| {
                    (
                        code.clone(),
                        S2PeriodIndicatorConfig {
                            data_type: item.data_type.clone(),
                            disabled: item.disabled,
                            period_enabled: item.period_enabled,
                            period_strategy_group: item.period_strategy_group.clone(),
                            min_change_value: item.min_change_value,
                        },
                    )
                })
                .collect(),
            bands: self.bands.clone(),
            text_change_severity: self.text_change_severity.clone(),
            current_missing_severity: self.current_missing_severity.clone(),
            previous_missing_severity: self.previous_missing_severity.clone(),
        }
    }

    pub fn rule_engine_config(&self) -> S2RuleEngineConfig {
        S2RuleEngineConfig {
            indicators: self
                .indicators
                .iter()
                .map(|(code, item)| {
                    (
                        code.clone(),
                        S2IndicatorConfig {
                            indicator_name: item.indicator_name.clone(),
                            form_code: item.form_code.clone(),
                            data_type: item.data_type.clone(),
                        },
                    )
                })
                .collect(),
            rules: self.rules.clone(),
            unit_settings: S2UnitSettings {
                source_unit: self.unit_settings.source_unit.clone(),
                output_file_unit: self.unit_settings.output_file_unit.clone(),
            },
        }
    }

    pub fn external_config(&self) -> S2ExternalConfig {
        S2ExternalConfig {
            indicators: self
                .indicators
                .iter()
                .map(|(code, item)| {
                    (
                        code.clone(),
                        S2ExternalIndicatorConfig {
                            data_type: item.data_type.clone(),
                        },
                    )
                })
                .collect(),
            institutions_by_id: self
                .institutions_by_id
                .iter()
                .map(|(code, item)| {
                    (
                        code.clone(),
                        S2ExternalInstitutionProfile {
                            report_item: item.report_item.clone(),
                        },
                    )
                })
                .collect(),
            institutions_by_name: self
                .institutions_by_name
                .iter()
                .map(|(name, item)| {
                    (
                        name.clone(),
                        S2ExternalInstitutionProfile {
                            report_item: item.report_item.clone(),
                        },
                    )
                })
                .collect(),
            external_rules: self.external_rules.clone(),
            unit_settings: S2ExternalUnitSettings {
                external_file_unit: self.unit_settings.external_file_unit.clone(),
                output_file_unit: self.unit_settings.output_file_unit.clone(),
            },
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
struct GeneralSettingsResult {
    units: S2ConfigUnitSettings,
    hide_unchanged_period_rows: bool,
    hide_matching_external_rows: bool,
    text_change_severity: String,
    current_missing_severity: String,
    previous_missing_severity: String,
    missing_row_warnings: Vec<String>,
}

/// Load the complete S2 runtime configuration from stored worksheet values.
pub fn load_config(path: impl AsRef<Path>) -> Result<S2AuditConfig> {
    let path = path.as_ref();
    let sheets = read_config_sheets(path)?;
    let resolved = resolved_path(path)?;
    let mut config = S2AuditConfig::default();
    let general = load_general_settings(&sheets, path)?;
    config.unit_settings = general.units;
    config.hide_unchanged_period_rows = general.hide_unchanged_period_rows;
    config.hide_matching_external_rows = general.hide_matching_external_rows;
    config.text_change_severity = general.text_change_severity;
    config.current_missing_severity = general.current_missing_severity;
    config.previous_missing_severity = general.previous_missing_severity;
    config.warnings.extend(general.missing_row_warnings);

    let indicator_sheet =
        sheet(&sheets, "指标参照").ok_or_else(|| anyhow!("比较配置缺少“指标参照”"))?;
    let indicator_rows = rows(indicator_sheet);
    let headers = first_row(&indicator_rows);
    let inline_unit_headers = headers
        .iter()
        .map(|cell| cell_text(Some(cell)))
        .filter(|name| matches!(name.as_str(), "源数据单位" | "大集中数据单位"))
        .collect::<HashSet<_>>();
    if !inline_unit_headers.is_empty() {
        let mut names = inline_unit_headers.into_iter().collect::<Vec<_>>();
        names.sort();
        bail!(
            "单位统一配置应位于本工作簿的“{GENERAL_SETTINGS_SHEET}”工作表；请从“指标参照”删除旧字段：{}",
            names.join("、")
        );
    }
    let (mapping, data_rows) = row_mapping(&indicator_rows, INDICATOR_ALIASES);
    if !mapping.contains_key("code") {
        bail!("“指标参照”缺少“指标代码”");
    }
    for (row_index, row) in data_rows {
        let code = cell_code(get(row, &mapping, "code"));
        if code.is_empty() {
            continue;
        }
        if config.indicators.contains_key(&code) {
            config
                .warnings
                .push(format!("指标参照存在重复指标代码：{code}"));
            continue;
        }
        let form_code = cell_code(get(row, &mapping, "form"));
        if form_code.is_empty() {
            config.warnings.push(format!(
                "指标参照第{}行指标 {} 未填写表单代码",
                row_index + 1,
                code
            ));
        }
        config.indicators.insert(
            code.clone(),
            S2ConfigIndicator {
                indicator_code: code,
                indicator_name: cell_text(get(row, &mapping, "name")),
                form_code,
                data_type: cell_text(get(row, &mapping, "type")),
                disabled: cell_yes(get(row, &mapping, "disabled")),
                note: cell_text(get(row, &mapping, "note")),
                period_enabled: cell_enabled(get(row, &mapping, "period_enabled")),
                period_strategy_group: cell_text(get(row, &mapping, "period_group")),
                min_change_value: cell_number(get(row, &mapping, "min_change")),
            },
        );
    }

    load_external_rules(&sheets, &resolved, &mut config)?;
    load_institutions(&sheets, &mut config);
    load_bands(&sheets, &mut config);
    load_rules(&sheets, &resolved, &mut config);
    Ok(config)
}

/// Read only the workbook's shared unit settings.
pub fn load_unit_settings(path: impl AsRef<Path>) -> Result<S2ConfigUnitSettings> {
    let path = path.as_ref();
    let sheets = read_config_sheets(path)?;
    Ok(load_general_settings(&sheets, path)?.units)
}

fn read_config_sheets(path: &Path) -> Result<Vec<SheetValues>> {
    if !path.is_file() {
        bail!("V2 比较配置不存在：{}", path.display());
    }
    read_value_columns_with_iso_dates(path, usize::MAX)
}

fn resolved_path(path: &Path) -> Result<String> {
    Ok(fs::canonicalize(path)?.to_string_lossy().into_owned())
}

fn import_profile(profile: &S2ConfigInstitution) -> S2InstitutionProfile {
    S2InstitutionProfile {
        institution_id: profile.institution_id.clone(),
        institution_name: profile.institution_name.clone(),
        social_credit_code: profile.social_credit_code.clone(),
        institution_type: profile.institution_type.clone(),
        handling_branch: profile.handling_branch.clone(),
        region: profile.region.clone(),
    }
}

fn load_general_settings(sheets: &[SheetValues], path: &Path) -> Result<GeneralSettingsResult> {
    let Some(sheet_data) = sheet(sheets, GENERAL_SETTINGS_SHEET) else {
        bail!(
            "报表采集配置“{}”缺少必需工作表“{}”；请在该配置工作簿中维护单位及显示设置。",
            file_name(path),
            GENERAL_SETTINGS_SHEET
        );
    };
    let sheet_rows = rows(sheet_data);
    let header = first_row(&sheet_rows);
    let expected = ["设置项", "设置值", "说明"];
    let actual = (0..3)
        .map(|index| cell_text(header.get(index)))
        .collect::<Vec<_>>();
    if actual.iter().map(String::as_str).collect::<Vec<_>>() != expected {
        bail!(
            "“{}”工作表表头必须为：A1=设置项、B1=设置值、C1=说明。",
            GENERAL_SETTINGS_SHEET
        );
    }

    let allowed = GENERAL_SETTINGS_FIELDS
        .iter()
        .copied()
        .collect::<HashSet<_>>();
    let mut configured = HashMap::<String, (usize, String)>::new();
    for (row_index, row) in sheet_rows.iter().enumerate().skip(1) {
        let key = cell_text(row.get(0));
        let value = cell_text(row.get(1));
        if key.is_empty() && value.is_empty() {
            continue;
        }
        let row_number = row_index + 1;
        if key.is_empty() {
            bail!(
                "“{}”工作表第{}行设置项不能为空。",
                GENERAL_SETTINGS_SHEET,
                row_number
            );
        }
        if !allowed.contains(key.as_str()) {
            bail!(
                "“{}”工作表第{}行设置项无效：{}",
                GENERAL_SETTINGS_SHEET,
                row_number,
                key
            );
        }
        if configured.contains_key(&key) {
            bail!("“{}”工作表设置项重复：{}", GENERAL_SETTINGS_SHEET, key);
        }
        configured.insert(key, (row_number, value));
    }

    let default_display = HashMap::from([
        ("源数据单位", "元"),
        ("外部文件单位", "元"),
        ("输出文件单位", "元"),
        ("是否隐藏无变动指标", "否"),
        ("是否隐藏与大集中一致指标", "否"),
        ("文字变动默认级别", "硬性"),
        ("本期缺失默认级别", "硬性"),
        ("上期缺失默认级别", "硬性"),
    ]);
    let missing_row_warnings = GENERAL_SETTINGS_FIELDS
        .iter()
        .filter(|name| **name != "换算范围" && !configured.contains_key(**name))
        .map(|name| {
            format!(
                "通用设置缺少设置项“{}”，已按默认“{}”执行；请补齐该行后重跑。",
                name, default_display[name]
            )
        })
        .collect::<Vec<_>>();

    let mut units = S2ConfigUnitSettings::default();
    for (name, field) in [
        ("源数据单位", "source"),
        ("外部文件单位", "external"),
        ("输出文件单位", "output"),
    ] {
        if let Some((row_number, value)) = configured.get(name) {
            if !matches!(value.as_str(), "元" | "万元" | "亿元") {
                bail!(
                    "“{}”工作表第{}行“{}”必须为元、万元或亿元：{}",
                    GENERAL_SETTINGS_SHEET,
                    row_number,
                    name,
                    if value.is_empty() { "空白" } else { value }
                );
            }
            match field {
                "source" => units.source_unit = value.clone(),
                "external" => units.external_file_unit = value.clone(),
                "output" => units.output_file_unit = value.clone(),
                _ => unreachable!(),
            }
        }
    }

    let mut booleans = [
        ("是否隐藏无变动指标", false),
        ("是否隐藏与大集中一致指标", false),
    ];
    for (name, result) in &mut booleans {
        if let Some((row_number, value)) = configured.get(*name) {
            if !matches!(value.as_str(), "是" | "否") {
                bail!(
                    "“{}”工作表第{}行“{}”必须为“是”或“否”：{}",
                    GENERAL_SETTINGS_SHEET,
                    row_number,
                    name,
                    if value.is_empty() { "空白" } else { value }
                );
            }
            *result = value == "是";
        }
    }

    let mut severities = [
        ("文字变动默认级别", "硬性"),
        ("本期缺失默认级别", "硬性"),
        ("上期缺失默认级别", "硬性"),
    ];
    for (name, result) in &mut severities {
        if let Some((row_number, value)) = configured.get(*name) {
            if !matches!(value.as_str(), "硬性" | "软性") {
                bail!(
                    "“{}”工作表第{}行“{}”必须为硬性或软性：{}",
                    GENERAL_SETTINGS_SHEET,
                    row_number,
                    name,
                    if value.is_empty() { "空白" } else { value }
                );
            }
            *result = value;
        }
    }

    Ok(GeneralSettingsResult {
        units,
        hide_unchanged_period_rows: booleans[0].1,
        hide_matching_external_rows: booleans[1].1,
        text_change_severity: severities[0].1.to_owned(),
        current_missing_severity: severities[1].1.to_owned(),
        previous_missing_severity: severities[2].1.to_owned(),
        missing_row_warnings,
    })
}

fn load_external_rules(
    sheets: &[SheetValues],
    resolved: &str,
    config: &mut S2AuditConfig,
) -> Result<()> {
    let external_sheet =
        sheet(sheets, "外部核对规则").ok_or_else(|| anyhow!("比较配置缺少“外部核对规则”"))?;
    let external_rows = rows(external_sheet);
    let mut actual_headers = first_row(&external_rows)
        .iter()
        .map(|cell| cell_text(Some(cell)))
        .collect::<Vec<_>>();
    while actual_headers.last().is_some_and(String::is_empty) {
        actual_headers.pop();
    }
    if actual_headers
        .iter()
        .map(String::as_str)
        .collect::<Vec<_>>()
        != EXTERNAL_RULE_HEADERS
    {
        bail!(
            "“外部核对规则”表头必须依次为：{}；不再使用“外部系统”列。",
            EXTERNAL_RULE_HEADERS.join("、")
        );
    }
    let (mapping, data_rows) = row_mapping(&external_rows, EXTERNAL_RULE_ALIASES);
    if first_row(&external_rows)
        .iter()
        .any(|cell| cell_text(Some(cell)) == "大集中数据单位")
    {
        bail!(
            "单位统一配置应位于本工作簿的“{}”工作表；请从“外部核对规则”删除“大集中数据单位”字段",
            GENERAL_SETTINGS_SHEET
        );
    }
    for (row_index, row) in data_rows {
        let rule_id = cell_text(get(row, &mapping, "id"));
        let indicator_code = cell_code(get(row, &mapping, "indicator"));
        if rule_id.is_empty() && indicator_code.is_empty() {
            continue;
        }
        let row_number = row_index + 1;
        if rule_id.is_empty() || indicator_code.is_empty() {
            config.warnings.push(format!(
                "外部核对规则第{}行缺少规则编号或指标代码，已跳过",
                row_number
            ));
            continue;
        }
        if !config.indicators.contains_key(&indicator_code) {
            config.warnings.push(format!(
                "外部核对规则 {} 引用不存在指标：{}，已跳过",
                rule_id, indicator_code
            ));
            continue;
        }
        let external_indicator_name = cell_text(get(row, &mapping, "external_name"));
        let method = cell_text(get(row, &mapping, "method"));
        let tolerance = cell_number(get(row, &mapping, "tolerance"));
        let mut severity = severity(&cell_text(get(row, &mapping, "severity")));
        let Some(tolerance) = tolerance.filter(|value| value.is_finite() && *value >= 0.0) else {
            config.warnings.push(format!(
                "外部核对规则第{}行 {} 配置无效：比较方式须为等于/不等于/大于/大于等于/小于/小于等于，容差须为非负有限数，外部指标名称不能为空；已跳过",
                row_number, rule_id
            ));
            continue;
        };
        if external_indicator_name.is_empty()
            || !EXTERNAL_COMPARISON_METHODS.contains(&method.as_str())
        {
            config.warnings.push(format!(
                "外部核对规则第{}行 {} 配置无效：比较方式须为等于/不等于/大于/大于等于/小于/小于等于，容差须为非负有限数，外部指标名称不能为空；已跳过",
                row_number, rule_id
            ));
            continue;
        }
        if method != "等于" && tolerance != 0.0 {
            config.warnings.push(format!(
                "外部核对规则第{}行 {} 仅“等于”使用容差；其他比较方式容差必须为0，已跳过",
                row_number, rule_id
            ));
            continue;
        }
        if !matches!(severity.as_str(), "硬性" | "软性") {
            config.warnings.push(format!(
                "外部核对规则 {} 级别无效：{}，按硬性处理",
                rule_id, severity
            ));
            severity = "硬性".to_owned();
        }
        config.external_rules.push(S2ExternalRule {
            external_system: "大集中".to_owned(),
            rule_id,
            indicator_code,
            external_indicator_name,
            comparison_method: method,
            tolerance,
            severity,
            enabled: cell_enabled(get(row, &mapping, "enabled")),
            source: Some(S2SourceRef {
                source_file: resolved.to_owned(),
                source_sheet: "外部核对规则".to_owned(),
                source_row: row_number,
            }),
        });
    }
    Ok(())
}

fn load_institutions(sheets: &[SheetValues], config: &mut S2AuditConfig) {
    let Some(institution_sheet) = sheet(sheets, "机构参照") else {
        return;
    };
    let sheet_rows = rows(institution_sheet);
    let (mapping, data_rows) = row_mapping(&sheet_rows, INSTITUTION_ALIASES);
    for (_, row) in data_rows {
        if cell_yes(get(row, &mapping, "disabled")) {
            continue;
        }
        let name = cell_text(get(row, &mapping, "name"));
        let code = cell_code(get(row, &mapping, "code"));
        if name.is_empty() && code.is_empty() {
            continue;
        }
        let profile = S2ConfigInstitution {
            institution_id: if code.is_empty() {
                name.clone()
            } else {
                code.clone()
            },
            institution_name: if name.is_empty() {
                code.clone()
            } else {
                name.clone()
            },
            social_credit_code: code.clone(),
            institution_type: cell_text(get(row, &mapping, "type")),
            handling_branch: cell_text(get(row, &mapping, "branch")),
            region: cell_text(get(row, &mapping, "region")),
            report_item: cell_text(get(row, &mapping, "report_item")),
        };
        if !code.is_empty() {
            config.institutions_by_id.insert(code, profile.clone());
        }
        if !name.is_empty() {
            config.institutions_by_name.insert(name, profile);
        }
    }
}

fn load_bands(sheets: &[SheetValues], config: &mut S2AuditConfig) {
    let strategy_sheet = sheet(sheets, "环比策略");
    let (band_sheet, is_strategy_sheet) = match strategy_sheet {
        Some(value) => (value, true),
        None => match sheet(sheets, "环比规则") {
            Some(value) => (value, false),
            None => return,
        },
    };
    let sheet_rows = rows(band_sheet);
    let aliases = if is_strategy_sheet {
        PERIOD_ALIASES
    } else {
        BAND_ALIASES
    };
    let (mapping, data_rows) = row_mapping(&sheet_rows, aliases);
    for (row_index, row) in data_rows {
        let lower = cell_number(get(row, &mapping, "lower"));
        let upper = cell_number(get(row, &mapping, "upper"));
        if lower.is_none() && upper.is_none() && !is_strategy_sheet {
            continue;
        }
        let mut band = S2PeriodBandConfig {
            rule_id: {
                let rule_id = cell_text(get(row, &mapping, "id"));
                if rule_id.is_empty() {
                    format!("B{:03}", row_index)
                } else {
                    rule_id
                }
            },
            lower,
            upper,
            description: cell_text(get(row, &mapping, "desc")),
            severity: cell_text(get(row, &mapping, "severity")),
            disabled: cell_yes(get(row, &mapping, "disabled")),
            strategy_group: if is_strategy_sheet {
                cell_text(get(row, &mapping, "group"))
            } else {
                String::new()
            },
            data_type: if is_strategy_sheet {
                cell_text(get(row, &mapping, "type"))
            } else {
                String::new()
            },
        };
        if band
            .lower
            .zip(band.upper)
            .is_some_and(|(low, high)| high <= low)
        {
            config.warnings.push(format!(
                "环比规则 {} 上限必须大于下限，已跳过",
                band.rule_id
            ));
            continue;
        }
        if !band.severity.is_empty()
            && !matches!(severity(&band.severity).as_str(), "硬性" | "软性")
        {
            config.warnings.push(format!(
                "环比规则 {} 级别无效：{}，按软性处理",
                band.rule_id, band.severity
            ));
            band.severity = "软性".to_owned();
        } else if !band.severity.is_empty() {
            band.severity = severity(&band.severity);
        }
        validate_band_text(config, &band);
        config.bands.push(band);
    }
    config.bands.sort_by(|left, right| {
        left.strategy_group
            .cmp(&right.strategy_group)
            .then_with(|| match (left.lower, right.lower) {
                (None, None) => std::cmp::Ordering::Equal,
                (None, Some(_)) => std::cmp::Ordering::Less,
                (Some(_), None) => std::cmp::Ordering::Greater,
                (Some(left), Some(right)) => left
                    .partial_cmp(&right)
                    .unwrap_or(std::cmp::Ordering::Equal),
            })
    });
    for pair in config.bands.windows(2) {
        let [left, right] = pair else { continue };
        if left.strategy_group != right.strategy_group {
            continue;
        }
        if right.lower.is_none() {
            config.warnings.push(format!(
                "环比规则重叠：{} 与 {}（右档下限为空=无穷小）",
                left.rule_id, right.rule_id
            ));
            continue;
        }
        if left.upper.is_some_and(|upper| upper > right.lower.unwrap()) {
            config.warnings.push(format!(
                "环比规则重叠：{} 与 {}",
                left.rule_id, right.rule_id
            ));
        }
        if left.upper.is_some_and(|upper| upper < right.lower.unwrap()) {
            config.warnings.push(format!(
                "环比规则空洞：{} 与 {}",
                left.rule_id, right.rule_id
            ));
        }
    }
}

fn validate_band_text(config: &mut S2AuditConfig, band: &S2PeriodBandConfig) {
    let text = band.description.replace(' ', "");
    if !text.contains("降幅") {
        return;
    }
    let Some((left, right)) = parse_decline_description(&text) else {
        return;
    };
    if left < 0.0 && right > 0.0 {
        config.warnings.push(format!(
            "环比规则“{}”描述疑似正负号异常：{}；V2 将只按下限/上限数值执行，未自动修改配置。",
            if band.rule_id.is_empty() {
                &band.description
            } else {
                &band.rule_id
            },
            band.description
        ));
    }
}

fn parse_decline_description(text: &str) -> Option<(f64, f64)> {
    let chars = text.chars().collect::<Vec<_>>();
    for start in 0..chars.len() {
        if chars[start] != '(' {
            continue;
        }
        let mut index = start + 1;
        let first_start = index;
        if chars.get(index) == Some(&'-') {
            index += 1;
        }
        let mut digits = 0;
        while chars.get(index).is_some_and(|ch| ch.is_ascii_digit()) {
            index += 1;
            digits += 1;
        }
        if chars.get(index) == Some(&'.') {
            index += 1;
            while chars.get(index).is_some_and(|ch| ch.is_ascii_digit()) {
                index += 1;
                digits += 1;
            }
        }
        if digits == 0 || chars.get(index) != Some(&'%') {
            continue;
        }
        let first = chars[first_start..index].iter().collect::<String>();
        index += 1;
        if !matches!(chars.get(index), Some(',' | '，')) {
            continue;
        }
        index += 1;
        let second_start = index;
        if chars.get(index) == Some(&'-') {
            index += 1;
        }
        let mut second_digits = 0;
        while chars.get(index).is_some_and(|ch| ch.is_ascii_digit()) {
            index += 1;
            second_digits += 1;
        }
        if chars.get(index) == Some(&'.') {
            index += 1;
            while chars.get(index).is_some_and(|ch| ch.is_ascii_digit()) {
                index += 1;
                second_digits += 1;
            }
        }
        if second_digits == 0 || chars.get(index) != Some(&'%') {
            continue;
        }
        let second = chars[second_start..index].iter().collect::<String>();
        return Some((first.parse().ok()?, second.parse().ok()?));
    }
    None
}

fn load_rules(sheets: &[SheetValues], resolved: &str, config: &mut S2AuditConfig) {
    let Some(rule_sheet) = sheet(sheets, "校验规则") else {
        return;
    };
    let sheet_rows = rows(rule_sheet);
    let (mapping, data_rows) = row_mapping(&sheet_rows, RULE_ALIASES);
    for (row_index, row) in data_rows {
        let expression = cell_text(get(row, &mapping, "expression"));
        if expression.is_empty() {
            continue;
        }
        let rule_id = {
            let id = cell_text(get(row, &mapping, "id"));
            if id.is_empty() {
                format!("ROW{}", row_index + 1)
            } else {
                id
            }
        };
        let mut rule_severity = severity(&cell_text(get(row, &mapping, "severity")));
        if rule_severity.is_empty() {
            rule_severity = "软性".to_owned();
        } else if !matches!(rule_severity.as_str(), "硬性" | "软性") {
            config.warnings.push(format!(
                "规则 {} 级别无效：{}，按软性处理",
                rule_id, rule_severity
            ));
            rule_severity = "软性".to_owned();
        }
        let mut rule_type = cell_text(get(row, &mapping, "type"));
        if rule_type.is_empty() {
            rule_type = "表达式".to_owned();
        }
        if !matches!(
            rule_type.as_str(),
            "表达式" | "累计不降(当年)" | "累计不降(历史)"
        ) {
            config
                .warnings
                .push(format!("规则 {} 类型暂不支持：{}", rule_id, rule_type));
            continue;
        }
        let unknown = unknown_numeric_codes(&expression)
            .into_iter()
            .filter(|code| !config.indicators.contains_key(code))
            .collect::<Vec<_>>();
        if !unknown.is_empty() {
            config.warnings.push(format!(
                "规则 {} 引用不存在指标：{}",
                rule_id,
                unknown.join("、")
            ));
        }
        config.rules.push(S2Rule {
            rule_id,
            rule_type,
            description: {
                let description = cell_text(get(row, &mapping, "desc"));
                if description.is_empty() {
                    expression.clone()
                } else {
                    description
                }
            },
            expression,
            severity: rule_severity,
            threshold: cell_number(get(row, &mapping, "threshold")),
            invert: cell_yes(get(row, &mapping, "invert")),
            disabled: cell_yes(get(row, &mapping, "disabled")),
            note: cell_text(get(row, &mapping, "note")),
            source: Some(S2SourceRef {
                source_file: resolved.to_owned(),
                source_sheet: "校验规则".to_owned(),
                source_row: row_index + 1,
            }),
        });
    }
}

fn unknown_numeric_codes(expression: &str) -> Vec<String> {
    let chars = expression.chars().collect::<Vec<_>>();
    let mut codes = Vec::new();
    let mut index = 0;
    while index < chars.len() {
        if !matches!(chars[index], '[' | '{') {
            index += 1;
            continue;
        }
        let start = index + 1;
        let Some(relative_end) = chars[start..].iter().position(|ch| matches!(ch, ']' | '}'))
        else {
            break;
        };
        let end = start + relative_end;
        if end > start {
            let code = chars[start..end].iter().collect::<String>();
            let code = code.trim();
            if !code.is_empty() && code.chars().all(char::is_numeric) {
                codes.push(code.to_owned());
            }
        }
        index = end + 1;
    }
    codes
}

fn sheet<'a>(sheets: &'a [SheetValues], name: &str) -> Option<&'a SheetValues> {
    sheets.iter().find(|item| item.name == name)
}

fn rows(sheet: &SheetValues) -> Vec<Vec<CellValue>> {
    let mut coordinates = Vec::with_capacity(sheet.values.len());
    let mut max_row = 0;
    let mut max_column = 0;
    for (address, value) in &sheet.values {
        if let Some((row, column)) = parse_coordinate(address) {
            max_row = max_row.max(row);
            max_column = max_column.max(column);
            coordinates.push((row, column, value.clone()));
        }
    }
    if coordinates.is_empty() {
        return Vec::new();
    }
    let mut rows = vec![vec![CellValue::Empty; max_column + 1]; max_row + 1];
    for (row, column, value) in coordinates {
        rows[row][column] = value;
    }
    rows
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
    (row > 0 && column > 0).then_some((row - 1, column - 1))
}

fn first_row(rows: &[Vec<CellValue>]) -> &[CellValue] {
    rows.first().map(Vec::as_slice).unwrap_or_default()
}

fn row_mapping<'a>(
    rows: &'a [Vec<CellValue>],
    aliases: &[(&str, &str)],
) -> (HashMap<String, usize>, Vec<(usize, &'a [CellValue])>) {
    let mut mapping = HashMap::new();
    for (column, cell) in first_row(rows).iter().enumerate() {
        let header = cell_text(Some(cell));
        if let Some((_, field)) = aliases.iter().find(|(name, _)| *name == header) {
            mapping.entry((*field).to_owned()).or_insert(column);
        }
    }
    let data_rows = rows
        .iter()
        .enumerate()
        .skip(1)
        .map(|(index, row)| (index, row.as_slice()))
        .collect();
    (mapping, data_rows)
}

fn get<'a>(
    row: &'a [CellValue],
    mapping: &HashMap<String, usize>,
    field: &str,
) -> Option<&'a CellValue> {
    mapping.get(field).and_then(|index| row.get(*index))
}

pub(crate) fn cell_truthy(value: Option<&CellValue>) -> bool {
    match value {
        None | Some(CellValue::Empty) => false,
        Some(CellValue::Text(value)) => !value.is_empty(),
        Some(CellValue::Number(value) | CellValue::DateTimeSerial(value)) => *value != 0.0,
        Some(CellValue::Integer(value)) => *value != 0,
        Some(CellValue::Boolean(value)) => *value,
        Some(CellValue::DateTimeIso(_) | CellValue::DurationIso(_) | CellValue::Error(_)) => true,
    }
}

pub(crate) fn cell_text(value: Option<&CellValue>) -> String {
    if !cell_truthy(value) {
        return String::new();
    }
    match value.expect("truthy cell is present") {
        CellValue::Empty => String::new(),
        CellValue::Text(value) => value.trim().to_owned(),
        CellValue::Number(value) => python_float_repr(*value).trim().to_owned(),
        CellValue::Integer(value) => value.to_string(),
        CellValue::Boolean(value) => if *value { "True" } else { "False" }.to_owned(),
        CellValue::DateTimeSerial(value) => format!("{value:?}"),
        CellValue::DateTimeIso(value) => python_datetime_text(value),
        CellValue::DurationIso(value) | CellValue::Error(value) => value.trim().to_owned(),
    }
}

pub(crate) fn cell_code(value: Option<&CellValue>) -> String {
    match value {
        Some(CellValue::Number(number)) if number.is_finite() && number.fract() == 0.0 => {
            return python_integer_from_float(*number);
        }
        _ => {}
    }
    let text = cell_text(value);
    if let Some((integer, fraction)) = text.split_once('.') {
        if !integer.is_empty()
            && integer.chars().all(|ch| ch.is_ascii_digit())
            && !fraction.is_empty()
            && fraction.chars().all(|ch| ch == '0')
        {
            if let Ok(number) = text.parse::<f64>() {
                if number.is_finite() && number.fract() == 0.0 {
                    return python_integer_from_float(number);
                }
            }
        }
    }
    text
}

fn python_integer_from_float(value: f64) -> String {
    if value == 0.0 {
        "0".to_owned()
    } else {
        format!("{value:.0}")
    }
}

pub(crate) fn cell_yes(value: Option<&CellValue>) -> bool {
    matches!(
        cell_text(value).to_lowercase().as_str(),
        "是" | "1" | "true" | "y" | "yes"
    )
}

fn cell_enabled(value: Option<&CellValue>) -> bool {
    match value {
        None | Some(CellValue::Empty) => true,
        Some(CellValue::Text(text)) if text.is_empty() => true,
        _ => cell_yes(value),
    }
}

pub(crate) fn cell_number(value: Option<&CellValue>) -> Option<f64> {
    match value {
        None | Some(CellValue::Empty) => None,
        Some(CellValue::Text(value)) if value.is_empty() => None,
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

fn parse_datetime_iso(value: &str) -> Option<NaiveDateTime> {
    NaiveDateTime::parse_from_str(value, "%Y-%m-%dT%H:%M:%S%.f")
        .or_else(|_| NaiveDateTime::parse_from_str(value, "%Y-%m-%dT%H:%M:%S"))
        .ok()
}

fn python_datetime_text(value: &str) -> String {
    let Some(datetime) = parse_datetime_iso(value) else {
        return value.to_owned();
    };
    let text = datetime.format("%Y-%m-%d %H:%M:%S").to_string();
    let micros = datetime.nanosecond() / 1_000;
    if micros == 0 {
        text
    } else {
        format!("{text}.{micros:06}")
    }
}

pub(crate) fn python_datetime_repr(value: &str) -> String {
    let Some(datetime) = parse_datetime_iso(value) else {
        return value.to_owned();
    };
    let mut parts = vec![
        datetime.year().to_string(),
        datetime.month().to_string(),
        datetime.day().to_string(),
        datetime.hour().to_string(),
        datetime.minute().to_string(),
    ];
    let micros = datetime.nanosecond() / 1_000;
    if datetime.second() != 0 || micros != 0 {
        parts.push(datetime.second().to_string());
    }
    if micros != 0 {
        parts.push(micros.to_string());
    }
    format!("datetime.datetime({})", parts.join(", "))
}

pub(crate) fn severity(value: &str) -> String {
    match value {
        "错误" | "严重" => "硬性".to_owned(),
        "核实" | "提示" | "关注" => "软性".to_owned(),
        _ => value.to_owned(),
    }
}

pub(crate) fn file_name(path: &Path) -> String {
    path.file_name()
        .map(|name| name.to_string_lossy().into_owned())
        .unwrap_or_else(|| path.display().to_string())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn numeric_code_text_and_boolean_helpers_match_config_cells() {
        assert_eq!(cell_code(Some(&CellValue::Number(20201028.0))), "20201028");
        assert_eq!(
            cell_code(Some(&CellValue::Text("20201028.0".to_owned()))),
            "20201028"
        );
        assert_eq!(cell_code(Some(&CellValue::Integer(0))), "");
        assert!(cell_yes(Some(&CellValue::Text("YES".to_owned()))));
        assert!(cell_enabled(Some(&CellValue::Empty)));
        assert!(!cell_enabled(Some(&CellValue::Text("否".to_owned()))));
    }

    #[test]
    fn unicode_decline_warning_parser_reads_only_matching_percentage_pair() {
        assert_eq!(
            parse_decline_description("降幅(-80%,50%)"),
            Some((-80.0, 50.0))
        );
        assert_eq!(
            parse_decline_description("降幅(-80%,-50%)"),
            Some((-80.0, -50.0))
        );
        assert_eq!(
            parse_decline_description("增幅(-80%,50%)"),
            Some((-80.0, 50.0))
        );
    }

    #[test]
    fn unknown_numeric_codes_follow_bracket_token_order() {
        assert_eq!(
            unknown_numeric_codes("[20201001] + {20201002} + [alpha]"),
            ["20201001", "20201002"]
        );
    }
}
