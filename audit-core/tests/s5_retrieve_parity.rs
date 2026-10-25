//! Manual S5 retrieval parity runner.
//!
//! Reads JSONL retrieve requests (`AUDIT_S5_RETRIEVE_INPUT`), runs the Rust
//! `S5Retrieval` core, and writes outcomes JSONL (`AUDIT_S5_RETRIEVE_REPORT`).
//! Manual parity harness only.

use audit_core::s5_eval::{eval_value_to_json, EvalArgument};
use audit_core::s5_retrieval::S5Retrieval;
use serde::Deserialize;
use serde_json::json;
use std::{collections::BTreeMap, fs, path::PathBuf};

#[derive(Deserialize)]
struct CallInput {
    name: String,
    #[serde(default)]
    args: Vec<serde_json::Value>,
    #[serde(default)]
    keywords: BTreeMap<String, String>,
}

#[derive(Deserialize)]
struct RetrieveRequest {
    id: u64,
    source: PathBuf,
    #[serde(default = "default_source_id")]
    source_id: String,
    #[serde(default = "default_period")]
    period: String,
    #[serde(default = "default_institution")]
    institution_code: String,
    #[serde(default)]
    business: String,
    #[serde(default)]
    default_sheet: String,
    calls: Vec<CallInput>,
}

fn default_source_id() -> String {
    "CURRENT".to_owned()
}

fn default_period() -> String {
    "2026-07-31".to_owned()
}

fn default_institution() -> String {
    "INST".to_owned()
}

fn argument_value(value: &serde_json::Value) -> EvalArgument {
    match value {
        serde_json::Value::Array(values) => EvalArgument::Tuple(
            values
                .iter()
                .map(|item| item.as_str().unwrap_or_default().to_owned())
                .collect(),
        ),
        serde_json::Value::String(text) => EvalArgument::Text(text.clone()),
        other => EvalArgument::Text(other.to_string()),
    }
}

#[test]
#[ignore = "manual parity runner; requires AUDIT_S5_RETRIEVE_INPUT/REPORT"]
fn runs_s5_retrieve_parity_items() {
    let input_path: PathBuf = std::env::var_os("AUDIT_S5_RETRIEVE_INPUT")
        .expect("AUDIT_S5_RETRIEVE_INPUT is missing")
        .into();
    let report_path: PathBuf = std::env::var_os("AUDIT_S5_RETRIEVE_REPORT")
        .expect("AUDIT_S5_RETRIEVE_REPORT is missing")
        .into();

    let raw = fs::read_to_string(&input_path).expect("read retrieve input");
    let outcomes: Vec<serde_json::Value> = raw
        .lines()
        .filter(|line| !line.trim().is_empty())
        .map(|line| {
            let request: RetrieveRequest = serde_json::from_str(line).expect("parse request");
            let configuration = match audit_core::s5_config::load_configuration(
                &request
                    .source
                    .parent()
                    .map(|parent| parent.join("config.xlsx"))
                    .unwrap_or_else(|| PathBuf::from("config.xlsx")),
                false,
                "",
            ) {
                Ok(configuration) => configuration,
                Err(error) => {
                    return json!({
                        "id": request.id, "status": "setup_error",
                        "message": error.message,
                    });
                }
            };
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
            let mut call_outcomes = Vec::new();
            for call in &request.calls {
                let arguments: Vec<EvalArgument> = call.args.iter().map(argument_value).collect();
                let keywords: BTreeMap<String, String> = call
                    .keywords
                    .iter()
                    .map(|(k, v)| (k.clone(), v.clone()))
                    .collect();
                let outcome = match retrieval.retrieve(
                    &call.name,
                    &arguments,
                    &keywords,
                    &request.business,
                    &request.default_sheet,
                    "",
                ) {
                    Ok(value) => json!({
                        "status": "ok",
                        "result": eval_value_to_json(&value),
                        "evidence": value.evidence,
                    }),
                    Err(failure) => json!({
                        "status": "error",
                        "code": failure.code,
                        "message": failure.message,
                        "evidence": failure.evidence,
                    }),
                };
                call_outcomes.push(outcome);
            }
            json!({"id": request.id, "status": "ok", "calls": call_outcomes})
        })
        .collect();

    let mut lines: Vec<String> = outcomes.iter().map(|outcome| outcome.to_string()).collect();
    lines.push(String::new());
    fs::write(&report_path, lines.join("\n")).expect("write retrieve outcomes");
    println!("AUDIT_S5_RETRIEVE_OUTCOMES={}", outcomes.len());
}
