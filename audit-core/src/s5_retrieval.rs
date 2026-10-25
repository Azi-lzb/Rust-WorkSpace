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

use crate::s5_config::{Configuration, SourceContract};
use crate::s5_eval::{EvalArgument, EvalErrorKind, EvalValue};
use crate::s5_external_layout::ExternalLayout;
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

/// A source explicitly selected for one S5 period role.
pub struct S5SourceDescriptor<'a> {
    pub path: &'a Path,
    pub source_id: &'a str,
    pub period: &'a str,
    pub institution_code: &'a str,
    pub institution_name: &'a str,
}

struct BoundSource {
    source_path: PathBuf,
    source_id: String,
    source_period: String,
    institution_code: String,
    institution_name: String,
    sha256: String,
    book: ScanBook,
    index_cache: BTreeMap<String, crate::s5_scan::ScanResult>,
}

impl BoundSource {
    fn bind(
        descriptor: S5SourceDescriptor<'_>,
        role: &'static str,
        audit_period: &str,
    ) -> Result<Self, RetrievalFailure> {
        if descriptor.period.is_empty() {
            return Err(RetrievalFailure::new(
                "PERIOD_CONTRACT_MISSING",
                "来源未提供可验证的报告期",
            ));
        }
        if matches!(role, "current" | "external") && descriptor.period != audit_period {
            return Err(RetrievalFailure::new(
                "PERIOD_CONTRACT_MISMATCH",
                if role == "external" {
                    "外部来源期间与本次审核期间不一致"
                } else {
                    "本期来源期间与本次审核期间不一致"
                },
            ));
        }
        if role == "previous" && descriptor.period == audit_period {
            return Err(RetrievalFailure::new(
                "PERIOD_CONTRACT_MISMATCH",
                "上期来源期间与本期相同",
            ));
        }
        if matches!(role, "current" | "previous") && descriptor.institution_code.is_empty() {
            return Err(RetrievalFailure::new(
                "INSTITUTION_CONTRACT_MISSING",
                "本期/上期来源未提供可验证的机构身份",
            ));
        }
        let payload = std::fs::read(descriptor.path)
            .map_err(|_| RetrievalFailure::new("SOURCE_NOT_FOUND", "数据源文件无法读取"))?;
        let mut hasher = Sha256::new();
        hasher.update(&payload);
        let sha256 = format!("{:x}", hasher.finalize());
        let book = ScanBook::open(descriptor.path)
            .map_err(|error| RetrievalFailure::new("SOURCE_UNREADABLE", error))?;
        Ok(Self {
            source_path: descriptor.path.to_path_buf(),
            source_id: descriptor.source_id.to_owned(),
            source_period: descriptor.period.to_owned(),
            institution_code: descriptor.institution_code.to_owned(),
            institution_name: descriptor.institution_name.to_owned(),
            sha256,
            book,
            index_cache: BTreeMap::new(),
        })
    }
}

/// Runtime context for one current source and an optional selected previous source.
pub struct S5Retrieval {
    configuration: Configuration,
    current: BoundSource,
    previous: Option<BoundSource>,
    external: Option<BoundSource>,
    external_layout: Option<ExternalLayout>,
    external_has_time_series: Option<bool>,
    data_quality_warnings: Vec<JsonValue>,
}

impl S5Retrieval {
    /// Bind one current source with the no-contract gates.
    pub fn bind(
        configuration: Configuration,
        role: &'static str,
        source_path: &Path,
        source_id: &str,
        source_period: &str,
        institution_code: &str,
        audit_period: &str,
    ) -> Result<Self, RetrievalFailure> {
        let current = BoundSource::bind(
            S5SourceDescriptor {
                path: source_path,
                source_id,
                period: source_period,
                institution_code,
                institution_name: "",
            },
            role,
            audit_period,
        )?;
        Ok(Self {
            configuration,
            current,
            previous: None,
            external: None,
            external_layout: None,
            external_has_time_series: None,
            data_quality_warnings: Vec::new(),
        })
    }

    /// Bind current and optional previous workbooks using the no-contract
    /// gates mirrored from `_validate_binding_context`.
    pub fn bind_with_previous(
        configuration: Configuration,
        current: S5SourceDescriptor<'_>,
        previous: Option<S5SourceDescriptor<'_>>,
        audit_period: &str,
    ) -> Result<Self, RetrievalFailure> {
        Self::bind_with_sources(configuration, current, previous, None, audit_period)
    }

    /// Bind the configured current, previous, and one external source using
    /// the currently migrated single-source-per-role path.
    pub fn bind_with_sources(
        configuration: Configuration,
        current: S5SourceDescriptor<'_>,
        previous: Option<S5SourceDescriptor<'_>>,
        external: Option<S5SourceDescriptor<'_>>,
        audit_period: &str,
    ) -> Result<Self, RetrievalFailure> {
        let current = BoundSource::bind(current, "current", audit_period)?;
        let previous = previous
            .map(|descriptor| BoundSource::bind(descriptor, "previous", audit_period))
            .transpose()?;
        let external = external
            .map(|descriptor| BoundSource::bind(descriptor, "external", audit_period))
            .transpose()?;
        if previous
            .as_ref()
            .is_some_and(|previous| previous.institution_code != current.institution_code)
        {
            return Err(RetrievalFailure::new(
                "INSTITUTION_MISMATCH",
                "上期来源机构身份与本次审核机构不一致",
            ));
        }
        Ok(Self {
            configuration,
            current,
            previous,
            external,
            external_layout: None,
            external_has_time_series: None,
            data_quality_warnings: Vec::new(),
        })
    }

    /// SHA-256 captured when the current source was opened for this run.
    pub fn source_sha256(&self) -> &str {
        &self.current.sha256
    }

    /// SHA-256 captured for the selected previous source, when present.
    pub fn previous_source_sha256(&self) -> Option<&str> {
        self.previous.as_ref().map(|source| source.sha256.as_str())
    }

    /// SHA-256 captured for the selected external source, when present.
    pub fn external_source_sha256(&self) -> Option<&str> {
        self.external.as_ref().map(|source| source.sha256.as_str())
    }

    /// Dynamic external-source quality warnings to include in the result and ledger.
    pub fn data_quality_warnings(&self) -> &[JsonValue] {
        &self.data_quality_warnings
    }

    fn metrics_for_path(
        configuration: &Configuration,
        business: &str,
        row_path: &[String],
        column_path: &[String],
        sheet_name: &str,
    ) -> Vec<crate::s5_config::MetricSpec> {
        configuration
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
        if function_name == "V_EXT" {
            return self.retrieve_external(args, keywords, rule_business);
        }
        let role = match function_name {
            "V_CUR" => "current",
            "V_PRE" => "previous",
            _ => return Err(RetrievalFailure::new("SOURCE_NOT_FOUND", "未知取值函数")),
        };
        let source = if role == "current" {
            &mut self.current
        } else {
            self.previous.as_mut().ok_or_else(|| {
                RetrievalFailure::new(
                    "SOURCE_NOT_SELECTED",
                    format!("没有找到业务 {rule_business} 对应的previous数据源"),
                )
            })?
        };
        if role == "previous" {
            let requested_source = keywords
                .get("source")
                .filter(|source_id| !source_id.trim().is_empty())
                .map(String::as_str)
                .unwrap_or(rule_default_source);
            if !requested_source.is_empty() && requested_source != source.source_id {
                return Err(RetrievalFailure::new(
                    "SOURCE_NOT_SELECTED",
                    format!("未选择或未登记逻辑数据源：{requested_source}"),
                ));
            }
        }

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

        let mut candidates = Self::metrics_for_path(
            &self.configuration,
            rule_business,
            &row_path,
            &column_path,
            &sheet_name,
        );
        if candidates.is_empty() {
            candidates = Self::metrics_for_path(
                &self.configuration,
                rule_business,
                &row_path,
                &column_path,
                "",
            );
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
        if !source
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
        let (raw_value, address) = Self::locate_template_value(
            &self.configuration,
            source,
            &metric,
            &resolved_sheet,
            role,
        )?;

        Self::value_with_unit(
            source,
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

    fn retrieve_external(
        &mut self,
        args: &[EvalArgument],
        keywords: &BTreeMap<String, String>,
        rule_business: &str,
    ) -> Result<EvalValue, RetrievalFailure> {
        let source_id = self
            .external
            .as_ref()
            .ok_or_else(|| {
                RetrievalFailure::new(
                    "SOURCE_NOT_SELECTED",
                    format!("没有找到业务 {rule_business} 对应的external数据源"),
                )
            })?
            .source_id
            .clone();
        if let Some(requested_source) = keywords
            .get("source")
            .filter(|source_id| !source_id.trim().is_empty())
            .filter(|requested| requested.as_str() != source_id)
        {
            return Err(RetrievalFailure::new(
                "SOURCE_NOT_SELECTED",
                format!("未选择或未登记逻辑数据源：{requested_source}"),
            ));
        }
        let [EvalArgument::Text(metric_argument)] = args else {
            return Err(RetrievalFailure::new(
                "VALUE_TYPE_INVALID",
                "外部取值函数需要一个指标标题",
            ));
        };
        let metric_header = metric_argument.trim();
        let metric_matches: Vec<_> = self
            .configuration
            .external_metrics
            .iter()
            .filter(|metric| metric.source_id == source_id && metric.metric_header == metric_header)
            .collect();
        if metric_matches.len() != 1 {
            let code = if metric_matches.is_empty() {
                "METRIC_NOT_FOUND"
            } else {
                "METRIC_AMBIGUOUS"
            };
            return Err(RetrievalFailure::new(
                code,
                "外部指标目录未能唯一找到被引用的指标",
            ));
        }
        let metric = metric_matches[0].clone();
        if !metric.confirmed {
            return Err(RetrievalFailure::new(
                "METRIC_NOT_CONFIRMED",
                "外部指标口径尚未确认",
            ));
        }
        let contract = self
            .configuration
            .source_contracts
            .iter()
            .find(|contract| contract.source_id == source_id)
            .cloned();
        let has_time_series_layout = if let Some(value) = self.external_has_time_series {
            value
        } else {
            let value = self
                .external
                .as_ref()
                .is_some_and(|source| ExternalLayout::has_time_series_header(&source.book));
            self.external_has_time_series = Some(value);
            value
        };
        let Some(contract) = contract.filter(|_| !has_time_series_layout) else {
            let sheet_filter = keywords.get("sheet").map(String::as_str).unwrap_or("");
            return self.retrieve_external_from_layout(&metric, metric_header, sheet_filter);
        };
        let source = self
            .external
            .as_ref()
            .expect("external source checked above");
        if contract.mapping_sheet == "参照表" {
            return Err(RetrievalFailure::new(
                "MAPPING_SHEET_MISSING",
                "旧参照表已停用，请将机构映射统一维护在外部文件的机构参照表",
            ));
        }
        if !source
            .book
            .sheet_names()
            .iter()
            .any(|name| name == &contract.mapping_sheet)
        {
            return Err(RetrievalFailure::new(
                "MAPPING_SHEET_MISSING",
                "外部数据源缺少配置的机构映射工作表",
            ));
        }
        let source_period = self.validate_external_contract(source, &contract)?;
        let identity_column = Self::external_header_column(
            source,
            &contract.mapping_sheet,
            contract.mapping_header_row,
            &contract.mapping_identity_field,
        )?;
        let target_column = Self::external_header_column(
            source,
            &contract.mapping_sheet,
            contract.mapping_header_row,
            &contract.mapping_target_field,
        )?;
        if self.current.institution_code.is_empty() {
            return Err(RetrievalFailure::new(
                "INSTITUTION_CONTRACT_MISSING",
                "运行上下文没有当前机构代码，不能执行外部映射",
            ));
        }
        let mapping_rows: Vec<(usize, String)> = (contract.mapping_data_start_row as usize
            ..=source.book.sheet_max_row(&contract.mapping_sheet))
            .filter_map(|row| {
                let identity =
                    source
                        .book
                        .cell_raw_text(&contract.mapping_sheet, row, identity_column);
                (identity.trim() == self.current.institution_code).then(|| {
                    let report_item =
                        source
                            .book
                            .cell_raw_text(&contract.mapping_sheet, row, target_column);
                    (row, report_item.trim().to_owned())
                })
            })
            .collect();
        if mapping_rows.is_empty() {
            return Err(RetrievalFailure::new(
                "IDENTITY_NOT_MAPPED",
                "当前机构未在外部映射表中找到",
            ));
        }
        if mapping_rows.len() != 1 || mapping_rows[0].1.is_empty() {
            return Err(RetrievalFailure::new(
                "IDENTITY_MAPPING_AMBIGUOUS",
                "当前机构映射缺失或匹配多行，无法唯一确定报表项目",
            ));
        }
        let (mapping_row, report_item) = &mapping_rows[0];
        let sheet_name = keywords
            .get("sheet")
            .filter(|sheet| !sheet.trim().is_empty())
            .map(|sheet| sheet.as_str())
            .unwrap_or(&contract.default_sheet);
        if sheet_name.is_empty()
            || !source
                .book
                .sheet_names()
                .iter()
                .any(|name| name == sheet_name)
        {
            return Err(RetrievalFailure::new(
                "SHEET_NOT_FOUND",
                "外部数据源缺少配置指定的数据工作表",
            ));
        }
        let data_row = (contract.data_start_row as usize..=source.book.sheet_max_row(sheet_name))
            .filter(|row| {
                source
                    .book
                    .cell_raw_text(sheet_name, *row, contract.data_key_column as usize)
                    .trim()
                    == report_item
            })
            .collect::<Vec<_>>();
        if data_row.is_empty() {
            return Err(RetrievalFailure::new(
                "REPORT_ITEM_NOT_FOUND",
                "外部数据表中不存在映射得到的报表项目",
            ));
        }
        if data_row.len() > 1 {
            return Err(RetrievalFailure::new(
                "REPORT_ITEM_DUPLICATE",
                "外部数据表中映射得到的报表项目重复",
            ));
        }
        let metric_columns: Vec<usize> = (contract.metric_first_column as usize
            ..=source.book.sheet_max_col(sheet_name))
            .filter(|column| {
                source
                    .book
                    .cell_raw_text(sheet_name, contract.metric_header_row as usize, *column)
                    .trim()
                    == metric_header
            })
            .collect();
        if metric_columns.is_empty() {
            return Err(RetrievalFailure::new(
                "METRIC_NOT_FOUND",
                "外部数据表中不存在配置指标标题",
            ));
        }
        if metric_columns.len() > 1 {
            return Err(RetrievalFailure::new(
                "METRIC_AMBIGUOUS",
                "外部数据表的指标标题重复",
            ));
        }
        let (row, column) = (data_row[0], metric_columns[0]);
        if source.book.sheet_has_formula(sheet_name, row, column) {
            return Err(RetrievalFailure::new(
                "SOURCE_FORMULA_UNVERIFIED",
                "源数据目标单元格包含公式，未确认其重算缓存，不能作为审核输入",
            )
            .with_evidence(
                sheet_name,
                &format!("{}{}", column_letter_local(column), row),
            ));
        }
        let (raw_type, raw_value) = source.book.stored_value(sheet_name, row, column);
        if raw_type == "error" {
            return Err(RetrievalFailure::new(
                "VALUE_TYPE_INVALID",
                "源数据单元格包含 Excel 错误值",
            )
            .with_evidence(
                sheet_name,
                &format!("{}{}", column_letter_local(column), row),
            ));
        }
        let mapping_cell = format!(
            "{}{}:{}{}",
            column_letter_local(identity_column),
            mapping_row,
            column_letter_local(target_column),
            mapping_row
        );
        let evidence = json!({
            "source_id": source.source_id,
            "source_path": source.source_path.display().to_string(),
            "source_sha256": source.sha256,
            "sheet": sheet_name,
            "cell": format!("{}{}", column_letter_local(column), row),
            "raw_value": raw_value,
            "source_unit": metric.source_unit,
            "standard_unit": metric.standard_unit,
            "source_period": source_period,
            "source_contract_version": contract.version,
            "metric_layout_version": "",
            "metric_definition_evidence": metric.evidence,
            "institution_code": self.current.institution_code,
            "precision": metric.precision,
            "mapping_sheet": contract.mapping_sheet,
            "mapping_cell": mapping_cell,
            "mapping_key": self.current.institution_code,
            "report_item": report_item,
            "metric_header": metric_header,
            "mapping_method": "",
        });
        convert_value(
            &raw_value,
            &metric.data_type,
            &metric.metric_type,
            &metric.source_unit,
            &metric.standard_unit,
            &metric.storage_scale,
            metric.allow_blank,
            metric.precision,
            evidence,
        )
    }

    fn retrieve_external_from_layout(
        &mut self,
        metric: &crate::s5_config::ExternalMetricSpec,
        metric_header: &str,
        sheet_filter: &str,
    ) -> Result<EvalValue, RetrievalFailure> {
        let source = self
            .external
            .as_ref()
            .expect("external source checked above");
        let source_period = source.source_period.clone();
        if self.external_layout.is_none() {
            let layout = ExternalLayout::discover(&source.book)
                .map_err(|failure| RetrievalFailure::new(&failure.code, failure.message))?;
            let source_path = source.source_path.clone();
            let warnings: Vec<JsonValue> = layout
                .warnings
                .iter()
                .map(|warning| {
                    external_data_warning(
                        &source_path,
                        "",
                        "",
                        &warning.sheet,
                        &warning.location,
                        &warning.message,
                    )
                })
                .collect();
            self.external_layout = Some(layout);
            for warning in warnings {
                push_data_quality_warning(&mut self.data_quality_warnings, warning);
            }
        }
        let layout = self
            .external_layout
            .as_ref()
            .expect("layout is populated before retrieval");
        let lookup = layout
            .resolve(
                &self.current.institution_code,
                &self.current.institution_name,
                metric_header,
                sheet_filter,
                &source_period,
            )
            .map_err(|failure| RetrievalFailure::new(&failure.code, failure.message))?;
        let source = self
            .external
            .as_ref()
            .expect("external source checked above");
        let cell = format!("{}{}", column_letter_local(lookup.column), lookup.row);
        let formula_cell = source
            .book
            .sheet_has_formula(&lookup.sheet, lookup.row, lookup.column);
        let (raw_type, raw_value) = if formula_cell {
            (
                "string".to_owned(),
                source
                    .book
                    .cell_raw_text(&lookup.sheet, lookup.row, lookup.column),
            )
        } else {
            source
                .book
                .stored_value(&lookup.sheet, lookup.row, lookup.column)
        };
        let mut conversion_value = raw_value.clone();
        if matches!(
            metric.data_type.as_str(),
            "number" | "decimal" | "integer" | "numeric"
        ) {
            let normalized_value = raw_value.trim().replace(',', "");
            let issue = if raw_value.trim().is_empty() {
                Some("指标单元格为空白".to_owned())
            } else if !matches!(raw_type.as_str(), "number" | "integer" | "string") && !formula_cell
            {
                Some(format!(
                    "指标单元格内容“{}”不是数值",
                    external_value_preview(&raw_value)
                ))
            } else if crate::s5_eval::parse_decimal28(&normalized_value).is_err() {
                Some(format!(
                    "指标单元格内容“{}”不是数值",
                    external_value_preview(&raw_value)
                ))
            } else {
                None
            };
            if let Some(issue) = issue {
                let warning = external_data_warning(
                    &source.source_path,
                    &self.current.institution_name,
                    &self.current.institution_code,
                    &lookup.sheet,
                    &cell,
                    &format!(
                        "报表项目“{}”的指标“{}”{}，已按 0 参与审核；请核查是否误删或改动了数据。",
                        lookup.report_item, metric_header, issue
                    ),
                );
                push_data_quality_warning(&mut self.data_quality_warnings, warning);
                conversion_value = "0".to_owned();
            } else {
                conversion_value = normalized_value;
            }
        } else if matches!(metric.data_type.as_str(), "text" | "string")
            && !raw_value.is_empty()
            && raw_type != "string"
        {
            return Err(RetrievalFailure::new(
                "VALUE_TYPE_INVALID",
                "指标要求文本值，但源单元格不是文本",
            )
            .with_evidence(&lookup.sheet, &cell));
        }
        let evidence = json!({
            "source_id": source.source_id,
            "source_path": source.source_path.display().to_string(),
            "source_sha256": source.sha256,
            "sheet": lookup.sheet,
            "cell": cell,
            "raw_value": raw_value,
            "source_unit": metric.source_unit,
            "standard_unit": metric.standard_unit,
            "source_period": source.source_period,
            "source_contract_version": "",
            "metric_layout_version": "",
            "metric_definition_evidence": metric.evidence,
            "institution_code": self.current.institution_code,
            "precision": metric.precision,
            "mapping_sheet": lookup.mapping_sheet,
            "mapping_cell": lookup.mapping_cell,
            "mapping_key": lookup.mapping_key,
            "report_item": lookup.report_item,
            "metric_header": metric_header,
            "mapping_method": lookup.mapping_method,
        });
        convert_value(
            &conversion_value,
            &metric.data_type,
            &metric.metric_type,
            &metric.source_unit,
            &metric.standard_unit,
            &metric.storage_scale,
            metric.allow_blank,
            metric.precision,
            evidence,
        )
    }

    fn external_header_column(
        source: &BoundSource,
        sheet_name: &str,
        header_row: i64,
        header: &str,
    ) -> Result<usize, RetrievalFailure> {
        if header.is_empty() {
            return Err(RetrievalFailure::new(
                "SOURCE_CONTRACT_MISSING",
                "外部映射字段名未配置",
            ));
        }
        let matches: Vec<usize> = (1..=source.book.sheet_max_col(sheet_name))
            .filter(|column| {
                source
                    .book
                    .cell_raw_text(sheet_name, header_row as usize, *column)
                    .trim()
                    == header
            })
            .collect();
        if matches.is_empty() {
            return Err(RetrievalFailure::new(
                "MAPPING_FIELD_NOT_FOUND",
                "外部映射表缺少配置字段",
            ));
        }
        if matches.len() > 1 {
            return Err(RetrievalFailure::new(
                "MAPPING_FIELD_AMBIGUOUS",
                "外部映射表配置字段重复",
            ));
        }
        Ok(matches[0])
    }

    fn external_identity_value(
        source: &BoundSource,
        contract: &SourceContract,
        label: &str,
    ) -> Result<String, RetrievalFailure> {
        if contract.identity_sheet.is_empty() || label.is_empty() {
            return Ok(String::new());
        }
        if !source
            .book
            .sheet_names()
            .iter()
            .any(|name| name == &contract.identity_sheet)
        {
            return Err(RetrievalFailure::new(
                "IDENTITY_NOT_FOUND",
                "身份契约指定的工作表不存在",
            ));
        }
        let mut matches = std::collections::BTreeSet::new();
        let max_row = source.book.sheet_max_row(&contract.identity_sheet);
        let max_column = source.book.sheet_max_col(&contract.identity_sheet);
        for row in 1..=max_row {
            for column in 1..=max_column {
                if source
                    .book
                    .merged_raw_text(&contract.identity_sheet, row, column)
                    .trim()
                    != label
                {
                    continue;
                }
                let right = if column < max_column {
                    source
                        .book
                        .merged_raw_text(&contract.identity_sheet, row, column + 1)
                } else {
                    String::new()
                };
                let below = if row < max_row {
                    source
                        .book
                        .merged_raw_text(&contract.identity_sheet, row + 1, column)
                } else {
                    String::new()
                };
                let candidate = if right.trim().is_empty() {
                    below
                } else {
                    right
                };
                if !candidate.trim().is_empty() {
                    matches.insert(candidate.trim().to_owned());
                }
            }
        }
        if matches.len() > 1 {
            return Err(RetrievalFailure::new(
                "IDENTITY_AMBIGUOUS",
                "来源文件中的身份字段存在多项不同值",
            ));
        }
        Ok(matches.into_iter().next().unwrap_or_default())
    }

    fn validate_external_contract(
        &self,
        source: &BoundSource,
        contract: &SourceContract,
    ) -> Result<String, RetrievalFailure> {
        let identity =
            Self::external_identity_value(source, contract, &contract.institution_code_label)?;
        let code = if source.institution_code.is_empty() {
            identity.clone()
        } else {
            source.institution_code.clone()
        };
        if !source.institution_code.is_empty()
            && !identity.is_empty()
            && source.institution_code != identity
        {
            return Err(RetrievalFailure::new(
                "INSTITUTION_MISMATCH",
                "文件机构身份与运行选择的机构不一致",
            ));
        }
        if !code.is_empty()
            && !self.current.institution_code.is_empty()
            && code != self.current.institution_code
        {
            return Err(RetrievalFailure::new(
                "INSTITUTION_MISMATCH",
                "来源文件机构身份与本次审核机构不一致",
            ));
        }
        if contract.institution_required && code.is_empty() {
            return Err(RetrievalFailure::new(
                "INSTITUTION_CONTRACT_MISSING",
                "来源契约要求机构身份，但文件或运行参数均未提供",
            ));
        }
        let name_from_file =
            Self::external_identity_value(source, contract, &contract.institution_name_label)?;
        if !source.institution_name.is_empty()
            && !name_from_file.is_empty()
            && source.institution_name != name_from_file
        {
            return Err(RetrievalFailure::new(
                "INSTITUTION_MISMATCH",
                "文件机构名称与运行绑定机构名称不一致",
            ));
        }
        if !self.current.institution_name.is_empty()
            && !name_from_file.is_empty()
            && self.current.institution_name != name_from_file
        {
            return Err(RetrievalFailure::new(
                "INSTITUTION_MISMATCH",
                "来源文件机构名称与本次审核机构不一致",
            ));
        }
        let period_from_file =
            Self::external_identity_value(source, contract, &contract.period_label)?;
        if !source.source_period.is_empty()
            && !contract.period_label.is_empty()
            && !period_from_file.is_empty()
            && source.source_period != period_from_file
        {
            return Err(RetrievalFailure::new(
                "PERIOD_CONTRACT_MISMATCH",
                "文件期间与运行绑定期间不一致",
            ));
        }
        let period = if source.source_period.is_empty() {
            period_from_file
        } else {
            source.source_period.clone()
        };
        if !contract.expected_period.is_empty()
            && !period.is_empty()
            && contract.expected_period != period
        {
            return Err(RetrievalFailure::new(
                "PERIOD_CONTRACT_MISMATCH",
                "来源期间与数据源契约要求不一致",
            ));
        }
        if contract.period_required && period.is_empty() {
            return Err(RetrievalFailure::new(
                "PERIOD_CONTRACT_MISSING",
                "来源契约要求报告期，但文件未提供可验证期间",
            ));
        }
        if period.is_empty() {
            return Err(RetrievalFailure::new(
                "PERIOD_CONTRACT_MISSING",
                "来源未提供可验证的报告期",
            ));
        }
        if !self.current.source_period.is_empty() && period != self.current.source_period {
            return Err(RetrievalFailure::new(
                "PERIOD_CONTRACT_MISMATCH",
                "外部来源期间与本次审核期间不一致",
            ));
        }
        Ok(period)
    }

    fn template_index<'a>(
        configuration: &Configuration,
        source: &'a mut BoundSource,
        metric: &crate::s5_config::MetricSpec,
        role: &str,
    ) -> &'a crate::s5_scan::ScanResult {
        let cache_key = format!("{}|{}|{}", role, metric.business, metric.sheet);
        if !source.index_cache.contains_key(&cache_key) {
            let key_pairs: Vec<(Vec<String>, Vec<String>)> = configuration
                .metrics
                .iter()
                .filter(|item| item.business == metric.business && item.sheet == metric.sheet)
                .map(|item| (item.row_path.clone(), item.column_path.clone()))
                .collect();
            let result =
                scan_rule_paths(&source.book, &source.source_path, &metric.sheet, &key_pairs);
            source.index_cache.insert(cache_key.clone(), result);
        }
        &source.index_cache[&cache_key]
    }

    /// Mirror of `_locate_template_value` (paths branch) plus `value()` gates.
    fn locate_template_value(
        configuration: &Configuration,
        source: &mut BoundSource,
        metric: &crate::s5_config::MetricSpec,
        sheet_name: &str,
        role: &str,
    ) -> Result<(String, String), RetrievalFailure> {
        let key = (metric.row_path.clone(), metric.column_path.clone());
        let (matches, row_found, column_found) = {
            let index = Self::template_index(configuration, source, metric, role);
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
        if source.book.sheet_has_formula(sheet_name, row, column) {
            return Err(RetrievalFailure::new(
                "SOURCE_FORMULA_UNVERIFIED",
                "源数据目标单元格包含公式，未确认其重算缓存，不能作为审核输入",
            )
            .with_evidence(
                sheet_name,
                &format!("{}{}", column_letter_local(column), row),
            ));
        }
        let stored = source.book.stored_value(sheet_name, row, column);
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

    /// Mirror of `_value_with_unit` + `_convert_value`.
    #[allow(clippy::too_many_arguments)]
    fn value_with_unit(
        source: &BoundSource,
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
            "source_id": source.source_id,
            "source_path": source.source_path.display().to_string(),
            "source_sha256": source.sha256,
            "sheet": sheet,
            "cell": cell,
            "raw_value": raw_value,
            "source_unit": source_unit,
            "standard_unit": standard_unit,
            "source_period": source.source_period,
            "source_contract_version": "",
            "metric_layout_version": layout_version,
            "metric_definition_evidence": definition_evidence,
            "institution_code": source.institution_code,
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

fn external_data_warning(
    source_path: &Path,
    institution_name: &str,
    institution_code: &str,
    sheet: &str,
    location: &str,
    message: &str,
) -> JsonValue {
    json!({
        "institution_name": institution_name,
        "institution_code": institution_code,
        "source_file": source_path
            .file_name()
            .map(|name| name.to_string_lossy().into_owned())
            .unwrap_or_default(),
        "sheet": sheet,
        "location": location,
        "message": message,
    })
}

fn push_data_quality_warning(warnings: &mut Vec<JsonValue>, warning: JsonValue) {
    if !warnings.contains(&warning) {
        warnings.push(warning);
    }
}

fn external_value_preview(value: &str) -> String {
    let normalized = value.replace('\r', " ").replace('\n', " ");
    let mut characters = normalized.chars();
    let preview: String = characters.by_ref().take(77).collect();
    if characters.next().is_some() {
        format!("{preview}...")
    } else {
        preview
    }
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
