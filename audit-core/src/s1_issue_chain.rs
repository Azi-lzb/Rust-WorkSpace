//! Combine formula-result and conditional-format issues in source-workbook order.
//!
//! The S1 workflow emits formula issues from the calculated audit copy and
//! conditional-format issues from the source workbook. History classification
//! must receive both issue families for every accepted source before it can
//! infer that a prior issue was resolved.

use crate::models::Issue;
use anyhow::{Result, ensure};
use std::{collections::HashMap, fs, path::PathBuf};

#[derive(Clone, Debug)]
pub struct S1IssueFilePair {
    pub source_file: PathBuf,
    pub audit_workbook: PathBuf,
}

/// Merge the two S1 issue families with the same per-file ordering used by
/// Python's handle_results_merge: formula issues first, then conditional
/// issues. Conditional issues are rebound to the calculated audit copy for
/// later workbook navigation.
pub fn merge_s1_issue_sources(
    files: &[S1IssueFilePair],
    formula_issues: &[Issue],
    conditional_issues: &[Issue],
) -> Result<Vec<Issue>> {
    ensure!(!files.is_empty(), "S1 Issue 合流至少需要一个工作簿");

    let mut source_to_audit = HashMap::<String, String>::new();
    let mut audit_to_source = HashMap::<String, String>::new();
    for file in files {
        let source = path_key(&file.source_file);
        let audit = path_key(&file.audit_workbook);
        ensure!(
            source != audit,
            "S1 公式审核副本必须与源工作簿分离：{}",
            file.source_file.display()
        );
        ensure!(
            source_to_audit
                .insert(source.clone(), audit.clone())
                .is_none(),
            "S1 Issue 合流清单中存在重复源文件：{}",
            file.source_file.display()
        );
        ensure!(
            audit_to_source.insert(audit, source).is_none(),
            "S1 Issue 合流清单中存在重复审核副本：{}",
            file.audit_workbook.display()
        );
    }

    let mut formulas_by_source = HashMap::<String, Vec<&Issue>>::new();
    for issue in formula_issues {
        let source = path_key(&PathBuf::from(&issue.source_file));
        let audit = path_key(&PathBuf::from(&issue.audit_file));
        ensure!(
            source_to_audit.get(&source) == Some(&audit),
            "公式审核项不属于 S1 Issue 合流清单中的源文件/审核副本配对：{}",
            issue.issue_id
        );
        formulas_by_source.entry(source).or_default().push(issue);
    }

    let mut conditionals_by_source = HashMap::<String, Vec<&Issue>>::new();
    for issue in conditional_issues {
        let source = path_key(&PathBuf::from(&issue.source_file));
        ensure!(
            source_to_audit.contains_key(&source),
            "条件格式审核项不属于 S1 Issue 合流清单：{}",
            issue.issue_id
        );
        conditionals_by_source
            .entry(source)
            .or_default()
            .push(issue);
    }

    let mut merged = Vec::with_capacity(formula_issues.len() + conditional_issues.len());
    for file in files {
        let source = path_key(&file.source_file);
        if let Some(issues) = formulas_by_source.get(&source) {
            merged.extend(issues.iter().map(|issue| (*issue).clone()));
        }
        if let Some(issues) = conditionals_by_source.get(&source) {
            merged.extend(issues.iter().map(|issue| {
                let mut issue = (*issue).clone();
                issue.audit_file = file.audit_workbook.to_string_lossy().into_owned();
                issue
            }));
        }
    }
    Ok(merged)
}

fn path_key(path: &PathBuf) -> String {
    let normalized = fs::canonicalize(path).unwrap_or_else(|_| path.clone());
    let value = normalized.to_string_lossy().replace('\\', "/");
    #[cfg(windows)]
    {
        value.to_lowercase()
    }
    #[cfg(not(windows))]
    {
        value
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    fn issue(issue_id: &str, source: &str, audit: &str) -> Issue {
        Issue {
            issue_id: issue_id.to_owned(),
            period: "2026-07".to_owned(),
            batch_id: "batch".to_owned(),
            audit_time: "2026-10-24 12:00:00".to_owned(),
            triggered: true,
            status: String::new(),
            first_seen_period: String::new(),
            previous_seen_period: String::new(),
            consecutive_count: 1,
            org_code: "1001".to_owned(),
            org_name: "机构甲".to_owned(),
            report_code: String::new(),
            sheet_name: "报表".to_owned(),
            rule_id: issue_id.to_owned(),
            severity: "错误".to_owned(),
            formula_cell: "A1".to_owned(),
            target_cell: "A1".to_owned(),
            target_value: json!(1),
            formula_result: json!("错误"),
            message: issue_id.to_owned(),
            source_file: source.to_owned(),
            audit_file: audit.to_owned(),
            check_field: String::new(),
            check_type: String::new(),
            comparison_value: json!(""),
            reference_value: json!(""),
            difference_value: json!(""),
            detail: String::new(),
            institution_feedback: String::new(),
            auditor_opinion: String::new(),
        }
    }

    #[test]
    fn merges_formula_then_conditional_issues_per_source_and_rebinds_audit_copy() {
        let first = S1IssueFilePair {
            source_file: PathBuf::from("source-a.xlsx"),
            audit_workbook: PathBuf::from("audit-a.xlsx"),
        };
        let second = S1IssueFilePair {
            source_file: PathBuf::from("source-b.xlsx"),
            audit_workbook: PathBuf::from("audit-b.xlsx"),
        };
        let formulas = vec![
            issue("formula-a", "source-a.xlsx", "audit-a.xlsx"),
            issue("formula-b", "source-b.xlsx", "audit-b.xlsx"),
        ];
        let conditionals = vec![
            issue("condition-a", "source-a.xlsx", "source-a.xlsx"),
            issue("condition-b", "source-b.xlsx", "source-b.xlsx"),
        ];

        let merged = merge_s1_issue_sources(&[first, second], &formulas, &conditionals).unwrap();
        assert_eq!(
            merged
                .iter()
                .map(|issue| issue.issue_id.as_str())
                .collect::<Vec<_>>(),
            ["formula-a", "condition-a", "formula-b", "condition-b"]
        );
        assert_eq!(merged[1].audit_file, "audit-a.xlsx");
        assert_eq!(merged[3].audit_file, "audit-b.xlsx");
    }

    #[test]
    fn rejects_formula_issues_with_an_unmatched_or_incorrect_audit_copy() {
        let files = [S1IssueFilePair {
            source_file: PathBuf::from("source-a.xlsx"),
            audit_workbook: PathBuf::from("audit-a.xlsx"),
        }];
        let error = merge_s1_issue_sources(
            &files,
            &[issue("formula-a", "source-a.xlsx", "wrong-audit.xlsx")],
            &[],
        )
        .expect_err("formula issue must match the source/audit pair");
        assert!(error.to_string().contains("源文件/审核副本配对"));
    }

    #[test]
    fn rejects_a_source_workbook_used_as_its_own_formula_copy() {
        let source = PathBuf::from("source.xlsx");
        let error = merge_s1_issue_sources(
            &[S1IssueFilePair {
                source_file: source.clone(),
                audit_workbook: source,
            }],
            &[],
            &[],
        )
        .expect_err("source and audit copy must be distinct");
        assert!(format!("{error:#}").contains("必须与源工作簿分离"));
    }
}
