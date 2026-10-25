//! Manual S5 config-check parity runner.
//!
//! Reads one JSON request (`AUDIT_S5_CHECK_INPUT`) with the configuration
//! path and check parameters, runs the Rust `validate_s5_config` port, and
//! writes the report JSON (`AUDIT_S5_CHECK_REPORT`). Manual parity harness
//! only.

use audit_core::s5_config_check::validate_s5_config;
use serde::Deserialize;
use std::{fs, path::PathBuf};

#[derive(Deserialize)]
struct CheckRequest {
    config: PathBuf,
    #[serde(default)]
    template_path: String,
    #[serde(default)]
    template_dir: String,
    #[serde(default)]
    template_index_path: String,
    #[serde(default)]
    external_path: String,
    #[serde(default)]
    business: String,
    #[serde(default = "default_scope")]
    scope: String,
}

fn default_scope() -> String {
    "current".to_owned()
}

#[test]
#[ignore = "manual parity runner; requires AUDIT_S5_CHECK_INPUT/REPORT"]
fn runs_s5_config_check_parity_item() {
    let input_path: PathBuf = std::env::var_os("AUDIT_S5_CHECK_INPUT")
        .expect("AUDIT_S5_CHECK_INPUT is missing")
        .into();
    let report_path: PathBuf = std::env::var_os("AUDIT_S5_CHECK_REPORT")
        .expect("AUDIT_S5_CHECK_REPORT is missing")
        .into();

    let request: CheckRequest =
        serde_json::from_str(&fs::read_to_string(&input_path).expect("read check request"))
            .expect("parse check request");
    audit_core::s5_config_check::set_configuration_source_path(&request.config);
    let outcome = match validate_s5_config(
        &request.config,
        &request.template_path,
        &request.template_dir,
        &request.template_index_path,
        &request.external_path,
        &request.business,
        &request.scope,
    ) {
        Ok(report) => serde_json::json!({"status": "ok", "report": report}),
        Err(error) => serde_json::json!({
            "status": "error",
            "error_kind": match error.kind {
                audit_core::s5_config_dsl::DslErrorKind::Parse => "parse",
                audit_core::s5_config_dsl::DslErrorKind::Configuration => "configuration",
            },
            "message": error.message,
        }),
    };
    fs::write(
        &report_path,
        serde_json::to_vec_pretty(&outcome).expect("serialize check outcome"),
    )
    .expect("write check outcome");
    println!(
        "AUDIT_S5_CHECK_STATUS={}",
        outcome
            .get("status")
            .and_then(serde_json::Value::as_str)
            .unwrap_or("unknown")
    );
}
