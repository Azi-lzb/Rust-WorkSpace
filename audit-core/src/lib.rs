//! Open-Data-Audit project package.
//!
//! Shared spreadsheet operations are provided by the workspace crates. This
//! package owns only Audit-specific probes, fixtures, rules, and future business
//! workflows.

pub mod central_csv;
pub mod conditional_format;
pub mod discovery;
pub mod models;
pub mod name_config;
pub mod period_pairs;
pub mod report_source;
pub mod s1_conditional_audit;
pub mod s1_formula_audit;
pub mod s1_formula_copy;
pub mod s1_history;
pub mod s1_issue_chain;
pub mod s1_region_summary;
pub mod s1_result_workbook;
pub mod s1_structure_compare;
pub mod s1_summary_merge;
pub mod s1_template;
pub mod s1_template_inspection;
pub mod s1_template_rules;
pub mod s2_analysis;
pub mod s2_config;
pub mod s2_config_check;
pub mod s2_exporter;
mod s2_expression;
pub mod s2_external;
pub mod s2_external_import;
pub mod s2_import;
pub mod s2_period;
pub mod s2_rule_engine;
pub mod s2_service;
pub mod s2_summary;
pub mod s2_validation;
pub mod s3_explanation_export;
pub mod s5_period_recognition;
pub mod s5_template_matching;
pub mod template;
pub mod unit_conversion;
pub mod web_bridge;
