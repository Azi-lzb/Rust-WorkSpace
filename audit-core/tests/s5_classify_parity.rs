//! Manual S5 business-identification/classification parity runner.
//!
//! Reads one JSON request (`AUDIT_S5_CLASSIFY_INPUT`) describing the file
//! list, template sources, configuration path, selection, and hints; runs the
//! Rust `classify_s5_files` port; and writes the rows JSON
//! (`AUDIT_S5_CLASSIFY_REPORT`). Manual parity harness only.

use audit_core::s5_business::classify_s5_files;
use serde::Deserialize;
use std::{fs, path::PathBuf};

#[derive(Deserialize)]
struct ClassifyRequest {
    files: Vec<PathBuf>,
    #[serde(default)]
    template_dir: Option<PathBuf>,
    #[serde(default)]
    template_path: Option<PathBuf>,
    #[serde(default)]
    config_path: Option<PathBuf>,
    #[serde(default)]
    selected_files: Vec<PathBuf>,
    #[serde(default)]
    business: String,
    #[serde(default)]
    period: String,
}

#[test]
#[ignore = "manual parity runner; requires AUDIT_S5_CLASSIFY_INPUT/REPORT"]
fn runs_s5_classify_parity_items() {
    let input_path: PathBuf = std::env::var_os("AUDIT_S5_CLASSIFY_INPUT")
        .expect("AUDIT_S5_CLASSIFY_INPUT is missing")
        .into();
    let report_path: PathBuf = std::env::var_os("AUDIT_S5_CLASSIFY_REPORT")
        .expect("AUDIT_S5_CLASSIFY_REPORT is missing")
        .into();

    let request: ClassifyRequest =
        serde_json::from_str(&fs::read_to_string(&input_path).expect("read classify request"))
            .expect("parse classify request");
    let rows = classify_s5_files(
        &request.files,
        request.template_dir.as_deref(),
        request.template_path.as_deref(),
        request.config_path.as_deref(),
        &request.selected_files,
        &request.business,
        &request.period,
    );
    let payload = serde_json::to_vec_pretty(&rows).expect("serialize classification rows");
    fs::write(&report_path, payload).expect("write classification rows");
    println!("AUDIT_S5_CLASSIFY_ROWS={}", rows.len());
}
