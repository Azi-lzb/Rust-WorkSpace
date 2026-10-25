//! Manual S5 rule-evaluation parity runner.
//!
//! Reads JSONL requests (`AUDIT_S5_RULE_INPUT`: config/source paths plus rule
//! identifiers), evaluates each rule through the Rust service port, and writes
//! outcomes JSONL (`AUDIT_S5_RULE_REPORT`). Manual parity harness only.

use audit_core::s5_config::load_configuration;
use audit_core::s5_retrieval::S5Retrieval;
use audit_core::s5_service::evaluate_rule;
use serde::Deserialize;
use serde_json::json;
use std::{fs, path::PathBuf};

#[derive(Deserialize)]
struct RuleRequest {
    id: u64,
    config: PathBuf,
    source: PathBuf,
    period: String,
    institution_code: String,
    institution_name: String,
    source_id: String,
    rule_ids: Vec<String>,
}

#[test]
#[ignore = "manual parity runner; requires AUDIT_S5_RULE_INPUT/REPORT"]
fn runs_s5_rule_parity_items() {
    let input_path: PathBuf = std::env::var_os("AUDIT_S5_RULE_INPUT")
        .expect("AUDIT_S5_RULE_INPUT is missing")
        .into();
    let report_path: PathBuf = std::env::var_os("AUDIT_S5_RULE_REPORT")
        .expect("AUDIT_S5_RULE_REPORT is missing")
        .into();

    let raw = fs::read_to_string(&input_path).expect("read rule input");
    let outcomes: Vec<serde_json::Value> = raw
        .lines()
        .filter(|line| !line.trim().is_empty())
        .map(|line| {
            let request: RuleRequest = serde_json::from_str(line).expect("parse rule request");
            let configuration =
                load_configuration(&request.config, false, "").expect("load configuration");
            let rules: Vec<_> = configuration
                .rules
                .iter()
                .filter(|rule| request.rule_ids.contains(&rule.rule_id))
                .cloned()
                .collect();
            let mut retrieval = match S5Retrieval::bind(
                configuration,
                "current",
                &request.source,
                &request.source_id,
                &request.period,
                &request.institution_code,
                &request.period,
            ) {
                Ok(retrieval) => retrieval,
                Err(failure) => {
                    return json!({
                        "id": request.id, "status": "bind_error",
                        "code": failure.code, "message": failure.message,
                    });
                }
            };
            let mut results = Vec::new();
            for rule in &rules {
                let outcome = evaluate_rule(
                    rule,
                    &mut retrieval,
                    "PARITY",
                    &request.period,
                    &request.institution_code,
                    &request.institution_name,
                    "",
                    None,
                );
                results.push(json!({
                    "rule_id": rule.rule_id,
                    "outcome": outcome,
                }));
            }
            json!({"id": request.id, "status": "ok", "results": results})
        })
        .collect();

    let mut lines: Vec<String> = outcomes.iter().map(|outcome| outcome.to_string()).collect();
    lines.push(String::new());
    fs::write(&report_path, lines.join("\n")).expect("write rule outcomes");
    println!("AUDIT_S5_RULE_OUTCOMES={}", outcomes.len());
}
