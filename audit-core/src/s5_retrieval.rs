//! S5 V_CUR/V_PRE retrieval core (single-binding, no-contract path).
//!
//! Faithful Rust port of the retrieval chain in
//! `s1audit_configured.sources` for the common configuration shape used by
//! real runs: derived metric specs (no 模板指标目录 sheet), one bound source
//! workbook per role, no 数据源契约 (period/institution gates on the binding
//! instead), and `_template_index_from_paths` (rule paths indexed directly via
//! `scan_rule_paths`). Unit conversion, type checks, blank rules, error codes,
//! and evidence construction mirror `_convert_value`/`_value_with_unit`.
//!
//! Retrieval always produces decimal-storage values (Python
//! `_convert_value`'s `EvalValue.number` default); the evaluate runtime
//! applies `_with_numeric_mode` when variables resolve.

use crate::s5_config::Configuration;
use crate::s5_eval::{EvalArgument, EvalErrorKind, EvalValue};
use crate::s5_scan::{in_numeric_regions, row_path_candidates, scan_rule_paths, ScanBook};
use serde_json::{json, Value as JsonValue};
use sha2::{Digest, Sha256};
use std::collections::BTreeMap;
use std::path::{Path, PathBuf};

#[derive(Clone, Debug, PartialEq)]
pub struct RetrievalFailure {
    pub code: String,
    pub message: String,
    pub evidence: JsonValue,
}

impl RetrievalFailure {
    fn new(code: &str, message: impl Into<String>) -> Self {
        Self {
            code: code.to_owned(),
            message: message.into(),
            evidence: json!({}),
        }
    }

    fn with_evidence(mut self, sheet: &str, cell: &str) -> Self {
        self.evidence = json!({"sheet": sheet, "cell": cell});
        self
    }

    /// Mirror of `SourceRepository.retrieve`'s failure enrichment: the
    /// function name and the resolved sheet (explicit argument, keyword, or
    /// rule default) are added unless already present.
    pub fn enriched_for(
        mut self,
        function_name: &str,
        args: &[EvalArgument],
        keywords: &BTreeMap<String, String>,
        rule_default_sheet: &str,
    ) -> Self {
        let evidence = self.evidence.take();
        let mut object = match evidence {
            JsonValue::Object(map) => map,
            _ => serde_json::Map::new(),
        };
        object
            .entry("function".to_owned())
            .or_insert_with(|| json!(function_name));
        if matches!(function_name, "V_CUR" | "V_PRE") {
            let explicit = args.first().and_then(|argument| match argument {
                EvalArgument::Text(text) if args.len() == 3 => Some(text.clone()),
                _ => None,
            });
            let sheet = explicit
                .or_else(|| keywords.get("sheet").cloned())
                .unwrap_or_else(|| rule_default_sheet.to_owned());
            object
                .entry("sheet".to_owned())
                .or_insert_with(|| json!(sheet));
        }
        self.evidence = JsonValue::Object(object);
        self
    }

    pub fn to_eval_error(self) -> crate::s5_eval::EvalError {
        crate::s5_eval::EvalError {
            kind: EvalErrorKind::Evaluation,
            message: self.message,
            code: Some(self.code),
            evidence: self.evidence,
        }
    }
}

const AMOUNT_FACTORS: [(&str, i64); 3] = [("元", 1), ("万元", 10000), ("亿元", 100000000)];

fn amount_factor(unit: &str) -> Option<i64> {
    AMOUNT_FACTORS
        .iter()
        .find(|(name, _)| *name == unit)
        .map(|(_, factor)| *factor)
}

/// Single-binding runtime context for the current/previous source.
pub struct S5Retrieval {
    configuration: Configuration,
    source_path: PathBuf,
    source_id: String,
    source_period: String,
    institution_code: String,
    sha256: String,
    book: ScanBook,
    index_cache: BTreeMap<String, crate::s5_scan::ScanResult>,
}

impl S5Retrieval {
    /// Bind one source workbook for `role` ("current" or "previous") with the
    /// no-contract gates from `_validate_binding_context`.
    pub fn bind(
        configuration: Configuration,
        role: &'static str,
        source_path: &Path,
        source_id: &str,
        source_period: &str,
        institution_code: &str,
        audit_period: &str,
    ) -> Result<Self, RetrievalFailure> {
        if source_period.is_empty() {
            return Err(RetrievalFailure::new(
                "PERIOD_CONTRACT_MISSING",
                "来源未提供可验证的报告期",
            ));
        }
        if matches!(role, "current" | "external") && source_period != audit_period {
            return Err(RetrievalFailure::new(
                "PERIOD_CONTRACT_MISMATCH",
                "本期来源期间与本次审核期间不一致",
            ));
        }
        if role == "previous" && source_period == audit_period {
            return Err(RetrievalFailure::new(
                "PERIOD_CONTRACT_MISMATCH",
                "上期来源期间与本期相同",
            ));
        }
        if matches!(role, "current" | "previous") && institution_code.is_empty() {
            return Err(RetrievalFailure::new(
                "INSTITUTION_CONTRACT_MISSING",
                "本期/上期来源未提供可验证的机构身份",
            ));
        }
        let payload = std::fs::read(source_path)
            .map_err(|_| RetrievalFailure::new("SOURCE_NOT_FOUND", "数据源文件无法读取"))?;
        let mut hasher = Sha256::new();
        hasher.update(&payload);
        let sha256 = format!("{:x}", hasher.finalize());
        let book = ScanBook::open(source_path)
            .map_err(|error| RetrievalFailure::new("SOURCE_UNREADABLE", error))?;
        Ok(Self {
            configuration,
            source_path: source_path.to_path_buf(),
            source_id: source_id.to_owned(),
            source_period: source_period.to_owned(),
            institution_code: institution_code.to_owned(),
            sha256,
            book,
            index_cache: BTreeMap::new(),
        })
    }

    fn metrics_for_path(
        &self,
        business: &str,
        row_path: &[String],
        column_path: &[String],
        sheet_name: &str,
    ) -> Vec<crate::s5_config::MetricSpec> {
        self.configuration
            .metrics
            .iter()
            .filter(|metric| {
                metric.business == business
                    && metric.row_path == row_path
                    && metric.column_path == column_path
            })
            .filter(|metric| sheet_name.is_empty() || metric.sheet == sheet_name)
            .cloned()
            .collect()
    }

    /// Mirror of `_retrieve_template` for `V_CUR`/`V_PRE` without contracts.
    pub fn retrieve(
        &mut self,
        function_name: &str,
        args: &[EvalArgument],
        keywords: &BTreeMap<String, String>,
        rule_business: &str,
        rule_default_sheet: &str,
        rule_default_source: &str,
    ) -> Result<EvalValue, RetrievalFailure> {
        self.retrieve_inner(
            function_name,
            args,
            keywords,
            rule_business,
            rule_default_sheet,
            rule_default_source,
        )
        .map_err(|failure| failure.enriched_for(function_name, args, keywords, rule_default_sheet))
    }

    fn retrieve_inner(
        &mut self,
        function_name: &str,
        args: &[EvalArgument],
        keywords: &BTreeMap<String, String>,
        rule_business: &str,
        rule_default_sheet: &str,
        rule_default_source: &str,
    ) -> Result<EvalValue, RetrievalFailure> {
        let role = match function_name {
            "V_CUR" => "current",
            "V_PRE" => {
                // This port binds a single current source; the previous-source
                // map is always empty, matching Python's unbound behavior.
                return Err(RetrievalFailure::new(
                    "SOURCE_NOT_SELECTED",
                    format!("没有找到业务 {rule_business} 对应的previous数据源"),
                ));
            }
            "V_EXT" => {
                return Err(RetrievalFailure::new(
                    "SOURCE_NOT_SELECTED",
                    format!("没有找到业务 {rule_business} 对应的external数据源"),
                ));
            }
            _ => return Err(RetrievalFailure::new("SOURCE_NOT_FOUND", "未知取值函数")),
        };
        // Single-binding resolution: the requested source (when given) must be
        // the bound one; otherwise the single binding serves every business.
        let _ = (rule_default_source, keywords.get("source"));

        let (positional_sheet, row_argument, column_argument) = match args {
            [EvalArgument::Text(sheet), row, column] => {
                (sheet.trim().to_owned(), Some(row), Some(column))
            }
            [row, column] => (String::new(), Some(row), Some(column)),
            _ => (String::new(), None, None),
        };
        let (Some(row_argument), Some(column_argument)) = (row_argument, column_argument) else {
            return Err(RetrievalFailure::new(
                "VALUE_TYPE_INVALID",
                "取值函数需要行路径与列路径",
            ));
        };
        let row_path = as_path(row_argument)?;
        let column_path = as_path(column_argument)?;
        let sheet_name = if !positional_sheet.is_empty() {
            positional_sheet
        } else if let Some(sheet) = keywords.get("sheet") {
            sheet.clone()
        } else {
            rule_default_sheet.to_owned()
        };

        let mut candidates =
            self.metrics_for_path(rule_business, &row_path, &column_path, &sheet_name);
        if candidates.is_empty() {
            candidates = self.metrics_for_path(rule_business, &row_path, &column_path, "");
        }
        if candidates.is_empty() {
            return Err(RetrievalFailure::new(
                "CATALOG_METRIC_NOT_FOUND",
                "模板指标目录中不存在该业务、Sheet 和行列路径",
            ));
        }
        if candidates.len() > 1 {
            return Err(RetrievalFailure::new(
                "METRIC_AMBIGUOUS",
                "模板指标路径匹配多条目录记录",
            ));
        }
        let metric = candidates.remove(0);
        if !metric.confirmed {
            return Err(RetrievalFailure::new(
                "METRIC_NOT_CONFIRMED",
                "目标指标口径尚未确认",
            ));
        }
        if metric.metric_type != "text"
            && (metric.source_unit.is_empty() || metric.standard_unit.is_empty())
        {
            return Err(RetrievalFailure::new(
                "UNIT_CONTRACT_MISSING",
                "目标指标缺少来源单位或标准单位",
            ));
        }
        let resolved_sheet = if sheet_name.is_empty() {
            metric.sheet.clone()
        } else {
            sheet_name
        };
        if resolved_sheet != metric.sheet {
            return Err(RetrievalFailure::new(
                "METRIC_NOT_FOUND",
                "默认或显式工作表与指标目录不一致",
            ));
        }
        if !self
            .book
            .sheet_names()
            .iter()
            .any(|name| name == &resolved_sheet)
        {
            return Err(RetrievalFailure::new(
                "SHEET_NOT_FOUND",
                "源工作簿中不存在指标目录指定的工作表",
            )
            .with_evidence(&resolved_sheet, ""));
        }
        let (raw_value, address) = self.locate_template_value(&metric, &resolved_sheet, role)?;

        self.value_with_unit(
            &raw_value,
            &metric,
            &resolved_sheet,
            &address,
            &metric.source_unit,
            &metric.standard_unit,
            &metric.storage_scale,
            metric.allow_blank,
            metric.precision,
            &metric.layout_version,
            &metric.evidence,
        )
    }

    fn template_index(
        &mut self,
        metric: &crate::s5_config::MetricSpec,
        role: &str,
    ) -> &crate::s5_scan::ScanResult {
        let cache_key = format!("{}|{}|{}", role, metric.business, metric.sheet);
        if !self.index_cache.contains_key(&cache_key) {
            let key_pairs: Vec<(Vec<String>, Vec<String>)> = self
                .configuration
                .metrics
                .iter()
                .filter(|item| item.business == metric.business && item.sheet == metric.sheet)
                .map(|item| (item.row_path.clone(), item.column_path.clone()))
                .collect();
            let result = scan_rule_paths(&self.book, &self.source_path, &metric.sheet, &key_pairs);
            self.index_cache.insert(cache_key.clone(), result);
        }
        &self.index_cache[&cache_key]
    }

    /// Mirror of `_locate_template_value` (paths branch) plus `value()` gates.
    fn locate_template_value(
        &mut self,
        metric: &crate::s5_config::MetricSpec,
        sheet_name: &str,
        role: &str,
    ) -> Result<(String, String), RetrievalFailure> {
        let key = (metric.row_path.clone(), metric.column_path.clone());
        let (matches, row_found, column_found) = {
            let index = self.template_index(metric, role);
            let matches = index.locations.get(&key).cloned().unwrap_or_default();
            let row_found = index
                .row_locations
                .get(&metric.row_path)
                .map(|rows| !rows.is_empty())
                .unwrap_or(false);
            let column_found = index
                .column_locations
                .get(&metric.column_path)
                .map(|columns| !columns.is_empty())
                .unwrap_or(false);
            (matches, row_found, column_found)
        };
        if matches.is_empty() {
            if !row_found {
                return Err(RetrievalFailure::new(
                    "ROW_PATH_NOT_FOUND",
                    "工作表中未找到配置的完整行路径",
                ));
            }
            if !column_found {
                return Err(RetrievalFailure::new(
                    "COLUMN_PATH_NOT_FOUND",
                    "工作表中未找到配置的完整列路径",
                ));
            }
            return Err(RetrievalFailure::new(
                "STRUCTURE_ERROR",
                "完整行路径与列路径没有有效交点",
            ));
        }
        if matches.len() > 1 {
            let cells: Vec<String> = matches
                .iter()
                .take(10)
                .map(|location| format!("{}{}", column_letter_local(location.column), location.row))
                .collect();
            let mut failure =
                RetrievalFailure::new("METRIC_AMBIGUOUS", "完整行列路径匹配多个单元格");
            failure.evidence = json!({"cells": cells});
            return Err(failure);
        }
        let (row, column) = (matches[0].row, matches[0].column);
        // `value()` gates: formula rejection and error cells.
        if self.book_has_formula(sheet_name, row, column) {
            return Err(RetrievalFailure::new(
                "SOURCE_FORMULA_UNVERIFIED",
                "源数据目标单元格包含公式，未确认其重算缓存，不能作为审核输入",
            )
            .with_evidence(
                sheet_name,
                &format!("{}{}", column_letter_local(column), row),
            ));
        }
        let stored = self.book_stored_value(sheet_name, row, column);
        if stored.0.as_str() == "error" {
            return Err(RetrievalFailure::new(
                "VALUE_TYPE_INVALID",
                "源数据单元格包含 Excel 错误值",
            )
            .with_evidence(
                sheet_name,
                &format!("{}{}", column_letter_local(column), row),
            ));
        }
        Ok((stored.1, format!("{}{}", column_letter_local(column), row)))
    }

    fn book_has_formula(&self, sheet: &str, row: usize, column: usize) -> bool {
        self.book.sheet_has_formula(sheet, row, column)
    }

    fn book_stored_value(&self, sheet: &str, row: usize, column: usize) -> (String, String) {
        self.book.stored_value(sheet, row, column)
    }

    /// Mirror of `_value_with_unit` + `_convert_value`.
    #[allow(clippy::too_many_arguments)]
    #[allow(clippy::too_many_arguments)]
    fn value_with_unit(
        &self,
        raw_value: &str,
        metric: &crate::s5_config::MetricSpec,
        sheet: &str,
        cell: &str,
        source_unit: &str,
        standard_unit: &str,
        storage_scale: &str,
        allow_blank: bool,
        precision: i64,
        layout_version: &str,
        definition_evidence: &str,
    ) -> Result<EvalValue, RetrievalFailure> {
        let evidence = json!({
            "source_id": self.source_id,
            "source_path": self.source_path.display().to_string(),
            "source_sha256": self.sha256,
            "sheet": sheet,
            "cell": cell,
            "raw_value": raw_value,
            "source_unit": source_unit,
            "standard_unit": standard_unit,
            "source_period": self.source_period,
            "source_contract_version": "",
            "metric_layout_version": layout_version,
            "metric_definition_evidence": definition_evidence,
            "institution_code": self.institution_code,
            "precision": precision,
            "mapping_sheet": "",
            "mapping_cell": "",
            "mapping_key": "",
            "report_item": "",
            "metric_header": "",
            "mapping_method": "",
        });
        convert_value(
            raw_value,
            &metric.data_type,
            &metric.metric_type,
            source_unit,
            standard_unit,
            storage_scale,
            allow_blank,
            precision,
            evidence,
        )
    }
}

fn as_path(argument: &EvalArgument) -> Result<Vec<String>, RetrievalFailure> {
    match argument {
        EvalArgument::Text(value) => {
            let parts = crate::s5_config::split_at_path(value);
            if parts.is_empty() || parts.iter().any(|part| part.is_empty()) {
                return Err(RetrievalFailure::new(
                    "VALUE_TYPE_INVALID",
                    "行路径不能为空",
                ));
            }
            Ok(parts)
        }
        EvalArgument::Tuple(values) => {
            let parts: Vec<String> = values.iter().map(|value| value.trim().to_owned()).collect();
            if parts.is_empty() || parts.iter().any(|part| part.is_empty()) {
                return Err(RetrievalFailure::new(
                    "VALUE_TYPE_INVALID",
                    "行路径不能为空",
                ));
            }
            Ok(parts)
        }
    }
}

fn column_letter_local(mut column: usize) -> String {
    let mut letters = Vec::new();
    while column > 0 {
        let digit = ((column - 1) % 26) as u8;
        letters.insert(0, (b'A' + digit) as char);
        column = (column - 1) / 26;
    }
    letters.into_iter().collect()
}

/// Mirror of `_convert_value`.
#[allow(clippy::too_many_arguments)]
fn convert_value(
    raw_value: &str,
    data_type: &str,
    metric_type: &str,
    source_unit: &str,
    standard_unit: &str,
    storage_scale: &str,
    allow_blank: bool,
    precision: i64,
    evidence: JsonValue,
) -> Result<EvalValue, RetrievalFailure> {
    if raw_value.trim().is_empty() {
        if !allow_blank {
            let mut failure = RetrievalFailure::new("CELL_EMPTY", "目标单元格为空");
            failure.evidence = json!({
                "sheet": evidence.get("sheet").cloned().unwrap_or(json!(null)),
                "cell": evidence.get("cell").cloned().unwrap_or(json!(null)),
            });
            return Err(failure);
        }
        return Ok(blank_value(standard_unit, metric_type, evidence));
    }
    if matches!(data_type, "text" | "string") {
        return Ok(EvalValue {
            value: crate::s5_eval::EvalScalar::Text(raw_value.to_owned()),
            kind: "string",
            unit: None,
            metric_type: "text".to_owned(),
            evidence: vec![evidence],
            is_literal: false,
            precision: None,
            numeric_mode: "decimal".to_owned(),
        });
    }
    if !matches!(data_type, "number" | "decimal" | "integer" | "numeric") {
        let mut failure =
            RetrievalFailure::new("VALUE_TYPE_INVALID", "指标目录中的数据类型不受支持");
        failure.evidence = evidence;
        return Err(failure);
    }
    // Source cells arrive as stored text from the shared reader; numeric
    // validation mirrors Python's isinstance(value, (int, float, Decimal)).
    let number = crate::s5_eval::parse_decimal28(raw_value).map_err(|_| {
        let mut failure =
            RetrievalFailure::new("VALUE_TYPE_INVALID", "指标要求数值，但源单元格不是数值");
        failure.evidence = evidence.clone();
        failure
    })?;
    if data_type == "integer" && !number.is_integral() {
        let mut failure =
            RetrievalFailure::new("VALUE_TYPE_INVALID", "指标要求整数，但源单元格包含小数");
        failure.evidence = evidence;
        return Err(failure);
    }
    let mut conversion = String::new();
    let mut number = number;
    if metric_type == "amount" {
        let (Some(source_factor), Some(standard_factor)) =
            (amount_factor(source_unit), amount_factor(standard_unit))
        else {
            let mut failure = RetrievalFailure::new(
                "UNIT_CONTRACT_MISSING",
                "金额指标单位必须明确为元/万元/亿元",
            );
            failure.evidence = evidence;
            return Err(failure);
        };
        let converted = number
            .multiply(&crate::s5_eval::Decimal28::from_i64(source_factor))
            .divide(&crate::s5_eval::Decimal28::from_i64(standard_factor));
        conversion = format!(
            "{source_unit}→{standard_unit}:×{}",
            source_factor as f64 / standard_factor as f64
        );
        number = converted;
    } else if metric_type == "rate" {
        if storage_scale == "fraction" && ["百分比", "%", "百分点"].contains(&standard_unit) {
            number = number.multiply(&crate::s5_eval::Decimal28::from_i64(100));
            conversion = "比例存储→百分比:×100".to_owned();
        } else if storage_scale == "percent" && standard_unit == "比例" {
            number = number.divide(&crate::s5_eval::Decimal28::from_i64(100));
            conversion = "百分比存储→比例:÷100".to_owned();
        } else if source_unit != standard_unit {
            if ["比例", "fraction"].contains(&source_unit)
                && ["百分比", "%", "百分点"].contains(&standard_unit)
            {
                number = number.multiply(&crate::s5_eval::Decimal28::from_i64(100));
                conversion = "比例→百分比:×100".to_owned();
            } else if ["百分比", "%", "百分点"].contains(&source_unit) && standard_unit == "比例"
            {
                number = number.divide(&crate::s5_eval::Decimal28::from_i64(100));
                conversion = "百分比→比例:÷100".to_owned();
            } else if ["百分比", "%", "百分点"].contains(&source_unit)
                && ["百分比", "%", "百分点"].contains(&standard_unit)
            {
                conversion = "比例单位别名规范化:数值不变".to_owned();
            } else {
                let mut failure = RetrievalFailure::new(
                    "UNIT_CONTRACT_MISSING",
                    "比例指标来源单位与标准单位不兼容",
                );
                failure.evidence = evidence;
                return Err(failure);
            }
        }
    } else if metric_type == "count" && source_unit != standard_unit {
        let mut failure =
            RetrievalFailure::new("UNIT_CONTRACT_MISSING", "户数/笔数不得转换为其他单位");
        failure.evidence = evidence;
        return Err(failure);
    } else if metric_type == "number" && source_unit != standard_unit {
        let mut failure =
            RetrievalFailure::new("UNIT_CONTRACT_MISSING", "普通数值指标不允许隐式换算");
        failure.evidence = evidence;
        return Err(failure);
    }
    let mut new_evidence = evidence.clone();
    if let Some(object) = new_evidence.as_object_mut() {
        object.insert("conversion".to_owned(), json!(conversion));
    }
    Ok(EvalValue {
        value: crate::s5_eval::EvalScalar::Number(crate::s5_eval::NumberValue::Decimal(number)),
        kind: "number",
        unit: Some(standard_unit.to_owned()),
        metric_type: metric_type.to_owned(),
        evidence: vec![new_evidence],
        is_literal: false,
        precision: Some(precision),
        numeric_mode: "decimal".to_owned(),
    })
}

fn blank_value(unit: &str, metric_type: &str, evidence: JsonValue) -> EvalValue {
    EvalValue {
        value: crate::s5_eval::EvalScalar::Blank,
        kind: "blank",
        unit: Some(unit.to_owned()),
        metric_type: metric_type.to_owned(),
        evidence: vec![evidence],
        is_literal: false,
        precision: None,
        numeric_mode: "decimal".to_owned(),
    }
}
