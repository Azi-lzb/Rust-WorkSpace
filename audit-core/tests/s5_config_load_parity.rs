//! Manual S5 configuration-load parity runner.
//!
//! Loads the real audit configuration workbook through the Rust
//! `s5_config::load_configuration` port and serializes the resulting
//! `Configuration` for field-by-field comparison with the Python baseline.
//! Manual parity harness only; not a production path.

use audit_core::s5_config;
use serde_json::json;
use std::{fs, path::PathBuf};

#[test]
#[ignore = "manual parity runner; requires AUDIT_S5_CONFIG_LOAD_INPUT/REPORT"]
fn runs_s5_config_load_parity_item() {
    let input_path: PathBuf = std::env::var_os("AUDIT_S5_CONFIG_LOAD_INPUT")
        .expect("AUDIT_S5_CONFIG_LOAD_INPUT is missing")
        .into();
    let report_path: PathBuf = std::env::var_os("AUDIT_S5_CONFIG_LOAD_REPORT")
        .expect("AUDIT_S5_CONFIG_LOAD_REPORT is missing")
        .into();

    let payload = match s5_config::load_configuration(&input_path, false, "") {
        Ok(configuration) => json!({
            "status": "ok",
            "configuration": serde_json::to_value(&configuration)
                .expect("serialize configuration"),
        }),
        Err(error) => json!({
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
        serde_json::to_vec_pretty(&payload).expect("serialize parity payload"),
    )
    .expect("write parity payload");
    println!(
        "AUDIT_S5_CONFIG_LOAD_STATUS={}",
        payload
            .get("status")
            .and_then(serde_json::Value::as_str)
            .unwrap_or("unknown")
    );
}
