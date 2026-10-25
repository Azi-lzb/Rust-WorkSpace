//! Manual S5 evaluation-runtime parity runner.
//!
//! Reads JSONL cases (`AUDIT_S5_EVAL_INPUT`) with an expression, variable
//! values, deterministic call-value entries, and a numeric mode; evaluates
//! through the Rust port; writes outcomes JSONL (`AUDIT_S5_EVAL_REPORT`).
//! Manual parity harness only.

use audit_core::s5_config_dsl::parse_expression;
use audit_core::s5_eval::{eval_value_to_json, evaluate_with_maps, EvalValue};
use serde::Deserialize;
use serde_json::json;
use std::{collections::BTreeMap, fs, path::PathBuf};

#[derive(Deserialize)]
struct ScalarInput {
    #[serde(default)]
    r#type: String,
    #[serde(default)]
    value: serde_json::Value,
    #[serde(default)]
    unit: Option<String>,
    #[serde(default)]
    metric_type: String,
    #[serde(default)]
    is_literal: bool,
    #[serde(default)]
    precision: Option<i64>,
}

#[derive(Deserialize)]
struct EvalCase {
    id: u64,
    expression: String,
    #[serde(default = "default_mode")]
    mode: String,
    #[serde(default)]
    variables: BTreeMap<String, ScalarInput>,
    #[serde(default)]
    calls: BTreeMap<String, ScalarInput>,
}

fn default_mode() -> String {
    "excel_binary64".to_owned()
}

fn to_eval_value(input: &ScalarInput, mode: &str) -> Result<EvalValue, String> {
    let unit = input.unit.clone();
    let metric_type = input.metric_type.clone();
    match input.r#type.as_str() {
        "number" => {
            let text = match &input.value {
                serde_json::Value::String(text) => text.clone(),
                serde_json::Value::Number(number) => number.to_string(),
                _ => return Err("number input needs a string or numeric value".to_owned()),
            };
            EvalValue::number_from_text(&text, mode)
                .map_err(|error| error.message)
                .map(|mut value| {
                    value.unit = unit;
                    value.metric_type = metric_type;
                    value.is_literal = input.is_literal;
                    value.precision = input.precision;
                    value
                })
        }
        "string" => Ok(EvalValue::string(
            input.value.as_str().unwrap_or_default().to_owned(),
            input.is_literal,
        )),
        "boolean" => Ok(EvalValue::boolean(
            input.value.as_bool().unwrap_or(false),
            Vec::new(),
            input.is_literal,
        )),
        "blank" => Ok(EvalValue {
            value: audit_core::s5_eval::EvalScalar::Blank,
            kind: "blank",
            unit,
            metric_type,
            evidence: Vec::new(),
            is_literal: input.is_literal,
            precision: input.precision,
            numeric_mode: mode.to_owned(),
        }),
        other => Err(format!("unsupported input type {other}")),
    }
}

#[test]
#[ignore = "manual parity runner; requires AUDIT_S5_EVAL_INPUT/REPORT"]
fn runs_s5_eval_parity_items() {
    let input_path: PathBuf = std::env::var_os("AUDIT_S5_EVAL_INPUT")
        .expect("AUDIT_S5_EVAL_INPUT is missing")
        .into();
    let report_path: PathBuf = std::env::var_os("AUDIT_S5_EVAL_REPORT")
        .expect("AUDIT_S5_EVAL_REPORT is missing")
        .into();

    let raw = fs::read_to_string(&input_path).expect("read eval input");
    let outcomes: Vec<serde_json::Value> = raw
        .lines()
        .filter(|line| !line.trim().is_empty())
        .map(|line| {
            let case: EvalCase = serde_json::from_str(line).expect("parse eval case");
            let node = match parse_expression(&case.expression) {
                Ok(node) => node,
                Err(error) => {
                    return json!({
                        "id": case.id,
                        "status": "error",
                        "error_kind": match error.kind {
                            audit_core::s5_config_dsl::DslErrorKind::Parse => "parse",
                            audit_core::s5_config_dsl::DslErrorKind::Configuration => "configuration",
                        },
                        "message": error.message,
                    });
                }
            };
            let mut variables = std::collections::BTreeMap::new();
            let mut ok = true;
            for (name, input) in &case.variables {
                match to_eval_value(input, &case.mode) {
                    Ok(value) => {
                        variables.insert(name.clone(), value);
                    }
                    Err(message) => {
                        ok = false;
                        return json!({
                            "id": case.id, "status": "error",
                            "error_kind": "configuration", "message": message,
                        });
                    }
                }
            }
            let _ = ok;
            let mut calls = std::collections::BTreeMap::new();
            for (key, input) in &case.calls {
                match to_eval_value(input, &case.mode) {
                    Ok(value) => {
                        calls.insert(key.clone(), value);
                    }
                    Err(message) => {
                        return json!({
                            "id": case.id, "status": "error",
                            "error_kind": "configuration", "message": message,
                        });
                    }
                }
            }
            match evaluate_with_maps(&node, &variables, &calls, &case.mode) {
                Ok(value) => json!({
                    "id": case.id,
                    "status": "ok",
                    "result": eval_value_to_json(&value),
                }),
                Err(error) => json!({
                    "id": case.id,
                    "status": "error",
                    "error_kind": match error.kind {
                        audit_core::s5_eval::EvalErrorKind::Evaluation => "evaluation",
                        audit_core::s5_eval::EvalErrorKind::Configuration => "configuration",
                    },
                    "message": error.message,
                }),
            }
        })
        .collect();

    let mut lines: Vec<String> = outcomes.iter().map(|outcome| outcome.to_string()).collect();
    lines.push(String::new());
    fs::write(&report_path, lines.join("\n")).expect("write eval outcomes");
    println!("AUDIT_S5_EVAL_OUTCOMES={}", outcomes.len());
}
