//! Audit-specific data models migrated from `base_audit.models`.

use serde::{Deserialize, Serialize};
use serde_json::Value;
use std::{collections::BTreeMap, path::PathBuf};

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct AuditRule {
    pub rule_id: String,
    pub enabled: bool,
    pub report_code: String,
    pub sheet_name: String,
    pub formula_cell: String,
    pub target_cell: String,
    pub severity: String,
    pub message: String,
    #[serde(default)]
    pub copy_range: String,
    #[serde(default = "default_result_mode")]
    pub result_mode: String,
    #[serde(default)]
    pub value_from_result: bool,
}

fn default_result_mode() -> String {
    "nonempty".to_owned()
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct CopyRange {
    pub sheet_name: String,
    pub address: String,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct TemplateDefinition {
    pub rules: Vec<AuditRule>,
    pub copy_ranges: Vec<CopyRange>,
    pub structured: bool,
    #[serde(default)]
    pub structure_ranges: Vec<CopyRange>,
    #[serde(default)]
    pub extraction_ranges: Vec<CopyRange>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct PreflightItem {
    pub category: String,
    pub level: String,
    pub status: String,
    pub sheet_name: String,
    pub location: String,
    pub message: String,
}

/// One fixed template label compared with its source counterpart.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct StructureCheck {
    pub sheet_name: String,
    pub cell_address: String,
    pub expected_value: Value,
    pub actual_value: Value,
    pub matched: bool,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct SourceMatch {
    pub matched: bool,
    pub score: f64,
    pub matched_labels: usize,
    pub checked_labels: usize,
    #[serde(default)]
    pub missing_sheets: Vec<String>,
    #[serde(default)]
    pub details: String,
    #[serde(default)]
    pub structure_checks: Vec<StructureCheck>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct PreflightRunResult {
    pub report_path: Option<PathBuf>,
    pub total_files: i64,
    pub matched_files: i64,
    pub template_warnings: i64,
    #[serde(default)]
    pub source_matches: Vec<(PathBuf, SourceMatch)>,
    #[serde(default)]
    pub batch_id: Option<String>,
    #[serde(default)]
    pub log_path: Option<PathBuf>,
}

impl PreflightRunResult {
    pub fn failed_files(&self) -> i64 {
        self.total_files - self.matched_files
    }

    pub fn summary_text(&self) -> String {
        let mut text = format!(
            "审核前检查完成：通过 {} 个，不通过 {} 个，模板提示 {} 项。",
            self.matched_files,
            self.failed_files(),
            self.template_warnings
        );
        if let Some(path) = &self.report_path {
            text.push_str(&format!("\n检查报告：{}", path.display()));
        }
        if let Some(path) = &self.log_path {
            text.push_str(&format!("\n运行日志：{}", path.display()));
        }
        text
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct Issue {
    pub issue_id: String,
    pub period: String,
    pub batch_id: String,
    pub audit_time: String,
    pub triggered: bool,
    pub status: String,
    pub first_seen_period: String,
    pub previous_seen_period: String,
    pub consecutive_count: i64,
    pub org_code: String,
    pub org_name: String,
    pub report_code: String,
    pub sheet_name: String,
    pub rule_id: String,
    pub severity: String,
    pub formula_cell: String,
    pub target_cell: String,
    pub target_value: Value,
    pub formula_result: Value,
    pub message: String,
    pub source_file: String,
    pub audit_file: String,
    #[serde(default)]
    pub check_field: String,
    #[serde(default)]
    pub check_type: String,
    #[serde(default = "empty_string_value")]
    pub comparison_value: Value,
    #[serde(default = "empty_string_value")]
    pub reference_value: Value,
    #[serde(default = "empty_string_value")]
    pub difference_value: Value,
    #[serde(default)]
    pub detail: String,
    #[serde(default)]
    pub institution_feedback: String,
    #[serde(default)]
    pub auditor_opinion: String,
}

fn empty_string_value() -> Value {
    Value::String(String::new())
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct FileAuditResult {
    pub source_path: PathBuf,
    pub audit_path: Option<PathBuf>,
    pub org_code: String,
    pub org_name: String,
    #[serde(default)]
    pub issues: Vec<Issue>,
    #[serde(default)]
    pub error: String,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct AuditRunResult {
    pub batch_id: String,
    pub period: String,
    pub output_dir: PathBuf,
    pub current_issues: Vec<Issue>,
    pub resolved_issues: Vec<Issue>,
    pub files: Vec<FileAuditResult>,
    #[serde(default)]
    pub summary_path: Option<PathBuf>,
    #[serde(default)]
    pub history_path: Option<PathBuf>,
    #[serde(default)]
    pub preflight_path: Option<PathBuf>,
    #[serde(default)]
    pub performance_lines: Vec<String>,
    #[serde(default)]
    pub log_path: Option<PathBuf>,
    #[serde(default)]
    pub copies: BTreeMap<String, PathBuf>,
}

impl AuditRunResult {
    pub fn successful_files(&self) -> usize {
        self.files
            .iter()
            .filter(|item| item.error.is_empty())
            .count()
    }

    pub fn failed_files(&self) -> usize {
        self.files
            .iter()
            .filter(|item| !item.error.is_empty())
            .count()
    }

    pub fn summary_text(&self) -> String {
        let mut parts = vec![
            format!(
                "审核完成：成功 {} 个文件，失败 {} 个文件。",
                self.successful_files(),
                self.failed_files()
            ),
            format!("本期审核结果 {} 条。", self.current_issues.len()),
        ];
        if let Some(path) = &self.summary_path {
            parts.push(format!("问题汇总：{}", path.display()));
        }
        if let Some(path) = &self.history_path {
            parts.push(format!("历史库：{}", path.display()));
        }
        if let Some(path) = &self.preflight_path {
            parts.push(format!("审核前检查：{}", path.display()));
        }
        if let Some(path) = &self.log_path {
            parts.push(format!("运行日志：{}", path.display()));
        }
        let skipped: Vec<_> = self
            .files
            .iter()
            .filter(|item| item.error.starts_with("报送文件与模板不匹配："))
            .filter_map(|item| item.source_path.file_name())
            .map(|name| name.to_string_lossy().into_owned())
            .collect();
        if !skipped.is_empty() {
            parts.push(format!(
                "已跳过 {} 个结构不匹配文件：{}",
                skipped.len(),
                skipped.join("、")
            ));
        }
        if !self.performance_lines.is_empty() {
            parts.push(format!("耗时分解：{}", self.performance_lines.join("；")));
        }
        parts.join("\n")
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::Value;

    fn baseline() -> Value {
        serde_json::from_str(include_str!("../probe/reports/models_python_baseline.json"))
            .expect("valid Python data-model baseline")
    }

    fn assert_round_trip<T>(value: &Value)
    where
        T: for<'de> Deserialize<'de> + Serialize,
    {
        let decoded: T = serde_json::from_value(value.clone()).expect("decode baseline model");
        assert_eq!(serde_json::to_value(decoded).unwrap(), *value);
    }

    fn assert_constructor_defaults<T>(record: &Value)
    where
        T: for<'de> Deserialize<'de> + Serialize,
    {
        let decoded: T = serde_json::from_value(record["args"].clone())
            .expect("decode required Python dataclass arguments");
        assert_eq!(serde_json::to_value(decoded).unwrap(), record["value"]);
    }

    #[test]
    fn audit_models_round_trip_python_dataclass_baseline() {
        let baseline = baseline();
        assert_round_trip::<AuditRule>(&baseline["audit_rule"]);
        assert_round_trip::<AuditRule>(&baseline["audit_rule_defaults"]);
        assert_round_trip::<CopyRange>(&baseline["copy_range"]);
        assert_round_trip::<TemplateDefinition>(&baseline["template_definition"]);
        assert_round_trip::<PreflightItem>(&baseline["preflight_item"]);
        assert_round_trip::<StructureCheck>(&baseline["structure_check"]);
        assert_round_trip::<SourceMatch>(&baseline["source_match"]);
        assert_round_trip::<Issue>(&baseline["issue"]);
        assert_round_trip::<FileAuditResult>(&baseline["file_audit_result"]);

        let preflight = &baseline["preflight_run_result"];
        let result: PreflightRunResult =
            serde_json::from_value(preflight["value"].clone()).unwrap();
        assert_eq!(serde_json::to_value(&result).unwrap(), preflight["value"]);
        assert_eq!(result.failed_files(), preflight["failed_files"]);
        assert_eq!(result.summary_text(), preflight["summary_text"]);

        let audit = &baseline["audit_run_result"];
        let result: AuditRunResult = serde_json::from_value(audit["value"].clone()).unwrap();
        assert_eq!(serde_json::to_value(&result).unwrap(), audit["value"]);
        assert_eq!(result.successful_files(), audit["successful_files"]);
        assert_eq!(result.failed_files(), audit["failed_files"]);
        assert_eq!(result.summary_text(), audit["summary_text"]);
    }

    #[test]
    fn python_dataclass_defaults_are_preserved() {
        let defaults = baseline()["constructor_defaults"].clone();
        assert_constructor_defaults::<AuditRule>(&defaults["audit_rule"]);
        assert_constructor_defaults::<TemplateDefinition>(&defaults["template_definition"]);
        assert_constructor_defaults::<SourceMatch>(&defaults["source_match"]);
        assert_constructor_defaults::<PreflightRunResult>(&defaults["preflight_run_result"]);
        assert_constructor_defaults::<Issue>(&defaults["issue"]);
        assert_constructor_defaults::<FileAuditResult>(&defaults["file_audit_result"]);
        assert_constructor_defaults::<AuditRunResult>(&defaults["audit_run_result"]);
    }
}
