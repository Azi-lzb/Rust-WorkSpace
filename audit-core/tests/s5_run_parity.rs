//! Manual S5 run-orchestration parity runner.
//!
//! Reads one JSON request (`AUDIT_S5_RUN_INPUT`) and writes the rule results
//! plus status counts (`AUDIT_S5_RUN_REPORT`). Manual parity harness only.

use audit_core::s5_config::load_configuration;
use audit_core::s5_retrieval::S5Retrieval;
use audit_core::s5_service::run_rules;
use serde::Deserialize;
use serde_json::json;
use std::{fs, path::PathBuf};

#[derive(Deserialize)]
struct RunRequest {
    config: PathBuf,
    source: PathBuf,
    period: String,
    institution_code: String,
    institution_name: String,
    source_id: String,
    businesses: Vec<String>,
}

#[test]
#[ignore = "manual parity runner; requires AUDIT_S5_RUN_INPUT/REPORT"]
fn runs_s5_run_parity_item() {
    let input_path: PathBuf = std::env::var_os("AUDIT_S5_RUN_INPUT")
        .expect("AUDIT_S5_RUN_INPUT is missing")
        .into();
    let report_path: PathBuf = std::env::var_os("AUDIT_S5_RUN_REPORT")
        .expect("AUDIT_S5_RUN_REPORT is missing")
        .into();

    let request: RunRequest =
        serde_json::from_str(&fs::read_to_string(&input_path).expect("read run request"))
            .expect("parse run request");
    let configuration = load_configuration(&request.config, false, "").expect("load configuration");
    let mut retrieval = S5Retrieval::bind(
        configuration.clone(),
        "current",
        &request.source,
        &request.source_id,
        &request.period,
        &request.institution_code,
        &request.period,
    )
    .expect("bind source");
    let (outcomes, counts) = run_rules(
        &configuration,
        &mut retrieval,
        &request.period,
        &request.businesses,
        &request.institution_code,
        &request.institution_name,
    )
    .expect("run rules");
    let payload = json!({
        "run_status": "COMPLETED",
        "status_counts": counts,
        "results": outcomes,
    });
    fs::write(
        &report_path,
        serde_json::to_vec_pretty(&payload).expect("serialize run payload"),
    )
    .expect("write run payload");
    println!("AUDIT_S5_RUN_RESULTS={}", outcomes.len());
}
