//! Manual S5 configuration-DSL parity runner.
//!
//! Reads JSONL work items (one enabled audit rule's DSL fields) from
//! `AUDIT_S5_DSL_PARITY_INPUT`, applies the same loader-level DSL check
//! sequence as the source project's `load_configuration` rule loop, and
//! writes JSONL outcomes to `AUDIT_S5_DSL_PARITY_REPORT`. This is a manual
//! parity harness only; it is not a production path and carries no capability
//! status.

use audit_core::s5_config_dsl::{
    parse_dictionary, parse_expression, static_kind, validate_dictionary_dependencies,
    validate_static_types, variables_in, DslError, Node,
};
use serde::{Deserialize, Serialize};
use std::{fs, path::PathBuf};

#[derive(Deserialize)]
struct RuleItem {
    id: u64,
    dictionary: Option<String>,
    /// `(field_name, expression_text)` pairs in the loader's evaluation order;
    /// `审核表达式` always comes first.
    expressions: Vec<(String, String)>,
}

#[derive(Serialize)]
struct RuleOutcome {
    id: u64,
    status: &'static str,
    error_kind: Option<&'static str>,
    message: Option<String>,
}

fn loader_kind_check(field_name: &str, kind: Option<&str>) -> Result<(), DslError> {
    if matches!(field_name, "审核表达式" | "适用表达式") && !matches!(kind, None | Some("boolean"))
    {
        return Err(DslError::configuration(format!(
            "{field_name} 必须返回布尔值"
        )));
    }
    if field_name == "差值表达式" && !matches!(kind, None | Some("number")) {
        return Err(DslError::configuration("差值表达式必须返回数值"));
    }
    Ok(())
}

/// Mirrors the DSL segment of the source `load_configuration` rule loop for
/// one rule row: dictionary parse, dependency check, per-node static types,
/// expression parses, static types plus field kind constraints, then
/// undefined-variable checks.
fn check_rule_item(item: &RuleItem) -> Result<(), DslError> {
    let dictionary = parse_dictionary(item.dictionary.as_deref())?;
    validate_dictionary_dependencies(&dictionary)?;
    for (_, node) in &dictionary {
        validate_static_types(node, false)?;
    }
    let mut expression_nodes: Vec<(String, Node)> = Vec::new();
    for (field_name, text) in &item.expressions {
        if text.trim().is_empty() {
            continue;
        }
        expression_nodes.push((field_name.clone(), parse_expression(text)?));
    }
    for (field_name, node) in &expression_nodes {
        validate_static_types(node, false)?;
        loader_kind_check(field_name, static_kind(node))?;
    }
    for (field_name, node) in &expression_nodes {
        let undefined: Vec<String> = variables_in(node)
            .into_iter()
            .filter(|name| !dictionary.iter().any(|(key, _)| key == name))
            .collect();
        if !undefined.is_empty() {
            return Err(DslError::configuration(format!(
                "{field_name} 引用未定义变量：{}",
                undefined.join(", ")
            )));
        }
    }
    Ok(())
}

#[test]
#[ignore = "manual parity runner; requires AUDIT_S5_DSL_PARITY_INPUT/REPORT"]
fn runs_s5_config_dsl_parity_items() {
    let input_path: PathBuf = std::env::var_os("AUDIT_S5_DSL_PARITY_INPUT")
        .expect("AUDIT_S5_DSL_PARITY_INPUT is missing")
        .into();
    let report_path: PathBuf = std::env::var_os("AUDIT_S5_DSL_PARITY_REPORT")
        .expect("AUDIT_S5_DSL_PARITY_REPORT is missing")
        .into();

    let raw = fs::read_to_string(&input_path).expect("read parity input");
    let outcomes: Vec<RuleOutcome> = raw
        .lines()
        .filter(|line| !line.trim().is_empty())
        .map(|line| {
            let item: RuleItem = serde_json::from_str(line).expect("parse parity item");
            let (status, error_kind, message) = match check_rule_item(&item) {
                Ok(()) => ("ok", None, None),
                Err(error) => {
                    let kind = match error.kind {
                        audit_core::s5_config_dsl::DslErrorKind::Parse => "parse",
                        audit_core::s5_config_dsl::DslErrorKind::Configuration => "configuration",
                    };
                    ("error", Some(kind), Some(error.message))
                }
            };
            RuleOutcome {
                id: item.id,
                status,
                error_kind,
                message,
            }
        })
        .collect();

    let mut lines = Vec::new();
    for outcome in &outcomes {
        lines.push(serde_json::to_string(outcome).expect("serialize outcome"));
    }
    lines.push(String::new());
    fs::write(&report_path, lines.join("\n")).expect("write parity outcomes");
    println!("AUDIT_S5_DSL_PARITY_OUTCOMES={}", outcomes.len());
}
