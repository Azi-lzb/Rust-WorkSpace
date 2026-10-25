//! Manual S5 outputs parity runner.
//!
//! Reads one JSON request (`AUDIT_S5_OUTPUTS_INPUT`), runs the rules and
//! writes run outputs, then reports a sanitized summary
//! (`AUDIT_S5_OUTPUTS_REPORT`). Manual parity harness only.

use audit_core::s5_config::load_configuration;
use audit_core::s5_outputs::{write_run_outputs, S5OutputRequest};
use audit_core::s5_retrieval::S5Retrieval;
use audit_core::s5_service::run_rules;
use serde::Deserialize;
use serde_json::json;
use std::{fs, path::PathBuf};

#[derive(Deserialize)]
struct OutputsRequest {
    config: PathBuf,
    source: PathBuf,
    period: String,
    institution_code: String,
    institution_name: String,
    source_id: String,
    run_id: String,
    output_root: PathBuf,
}

#[test]
#[ignore = "manual parity runner; requires AUDIT_S5_OUTPUTS_INPUT/REPORT"]
fn runs_s5_outputs_parity_item() {
    let input_path: PathBuf = std::env::var_os("AUDIT_S5_OUTPUTS_INPUT")
        .expect("AUDIT_S5_OUTPUTS_INPUT is missing")
        .into();
    let report_path: PathBuf = std::env::var_os("AUDIT_S5_OUTPUTS_REPORT")
        .expect("AUDIT_S5_OUTPUTS_REPORT is missing")
        .into();

    let request: OutputsRequest =
        serde_json::from_str(&fs::read_to_string(&input_path).expect("read outputs request"))
            .expect("parse outputs request");
    let configuration = load_configuration(&request.config, false, "").expect("load configuration");
    let source_sha = fs::read(&request.source)
        .map(|bytes| {
            use sha2::{Digest, Sha256};
            let mut hasher = Sha256::new();
            hasher.update(&bytes);
            format!("{:x}", hasher.finalize())
        })
        .unwrap_or_default();
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
        &["单位贷款".to_owned()],
        &request.institution_code,
        &request.institution_name,
    )
    .expect("run rules");
    let output_request = S5OutputRequest {
        run_id: &request.run_id,
        period: &request.period,
        institution_code: &request.institution_code,
        institution_name: &request.institution_name,
        outcomes: &outcomes,
        status_counts: &counts,
        source_path: &request.source,
        source_id: &request.source_id,
        source_sha256_before: &source_sha,
        config_path: &request.config,
        config_sha256_before: &configuration.sha256,
        config_version: &configuration.version,
    };
    let summary = write_run_outputs(&output_request, &request.output_root)
        .unwrap_or_else(|error| panic!("write outputs failed: {error}"));

    let payload = json!({
        "status": "ok",
        "run_dir": summary.run_dir.display().to_string(),
        "triggered_rows": summary.triggered_rows,
        "failure_rows": summary.failure_rows,
        "manifest_entry_count": summary.manifest.len(),
        "manifest_roles": summary.manifest.iter().map(|entry| entry.role.clone()).collect::<Vec<_>>(),
        "manifest_snapshots": summary.manifest.iter().map(|entry| entry.matches_snapshot).collect::<Vec<_>>(),
        "output_files": {
            "result_workbook": summary.output_files.result_workbook.display().to_string(),
            "run_ledger": summary.output_files.run_ledger.display().to_string(),
            "audit_copies": summary.output_files.audit_copies.display().to_string(),
            "configuration_copy": summary.output_files.configuration_copy.display().to_string(),
        },
    });
    fs::write(
        &report_path,
        serde_json::to_vec_pretty(&payload).expect("serialize outputs summary"),
    )
    .expect("write outputs summary");
    println!("AUDIT_S5_OUTPUTS_STATUS=ok");
}
