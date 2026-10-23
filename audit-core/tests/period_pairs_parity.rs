use audit_core::period_pairs::{list_period_pairs, report_source_warning};
use serde_json::{json, Value};
use std::{
    fs,
    path::{Path, PathBuf},
    time::{SystemTime, UNIX_EPOCH},
};

fn workspace_path(relative: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join(relative)
}

fn normalize_path(value: &mut Value, key: &str) {
    let path = value[key].as_str().unwrap_or_default();
    if !path.is_empty() {
        let basename = path.rsplit(['/', '\\']).next().unwrap_or(path);
        value[key] = json!(basename);
    }
}

#[test]
fn period_pair_preview_matches_read_only_python_baseline() {
    let cases: Value = serde_json::from_slice(
        &fs::read(workspace_path("tests/fixtures/period_pairs_cases.json"))
            .expect("read period-pair cases"),
    )
    .expect("decode period-pair cases");
    let baseline: Value = serde_json::from_slice(
        &fs::read(workspace_path(
            "probe/reports/period_pairs_python_baseline.json",
        ))
        .expect("read Python period-pair baseline"),
    )
    .expect("decode Python period-pair baseline");

    let suffix = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("clock")
        .as_nanos();
    let temp_root = std::env::temp_dir().join(format!("audit-period-pairs-parity-{suffix}"));
    let fixture_root = workspace_path("tests/fixtures");
    let mut results = Vec::new();

    for case in cases["cases"].as_array().expect("cases array") {
        let case_root = temp_root.join(case["id"].as_str().expect("case id"));
        let current = case_root.join("current");
        let previous = case_root.join("previous");
        fs::create_dir_all(&current).expect("create current fixture directory");
        fs::create_dir_all(&previous).expect("create previous fixture directory");
        for (side, directory) in [("current", &current), ("previous", &previous)] {
            for item in case[side].as_array().expect("side file list") {
                fs::copy(
                    fixture_root.join(item["fixture"].as_str().expect("fixture name")),
                    directory.join(item["name"].as_str().expect("target name")),
                )
                .expect("copy workbook fixture");
            }
        }

        let mut pairs = serde_json::to_value(
            list_period_pairs(Some(&current), Some(&previous)).expect("build Rust pairs"),
        )
        .expect("serialize Rust pairs");
        for pair in pairs.as_array_mut().expect("pairs array") {
            normalize_path(pair, "curPath");
            normalize_path(pair, "prePath");
        }
        results.push(json!({
            "id": case["id"],
            "warning": report_source_warning(Some(&current), Some(&previous)),
            "pairs": pairs
        }));
    }
    let _ = fs::remove_dir_all(&temp_root);

    assert_eq!(json!(results), baseline["results"]);
}
