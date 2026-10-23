use audit_core::s2_import::{import_period_directory, S2ImportConfig};
use serde_json::Value;
use std::{fs, path::Path};

fn workspace_path(relative: &str) -> std::path::PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join(relative)
}

fn normalize_source_paths(value: &mut Value) {
    let records = value["records"]
        .as_array_mut()
        .expect("records array in Rust result");
    for record in records {
        normalize_record_path(record);
    }
    let duplicates = value["duplicates"]
        .as_array_mut()
        .expect("duplicates array in Rust result");
    for duplicate in duplicates {
        for record in duplicate.as_array_mut().expect("duplicate record pair") {
            normalize_record_path(record);
        }
    }
}

fn normalize_record_path(record: &mut Value) {
    let path = record["source"]["source_file"]
        .as_str()
        .expect("source_file string")
        .rsplit(['/', '\\'])
        .next()
        .expect("basename");
    record["source"]["source_file"] = Value::String(path.to_owned());
}

#[test]
fn s2_directory_import_matches_read_only_python_importer_baseline() {
    let fixture_dir = workspace_path("tests/fixtures/s2_import");
    let config: S2ImportConfig = serde_json::from_slice(
        &fs::read(fixture_dir.join("config.json")).expect("read synthetic S2 config"),
    )
    .expect("decode S2 config");
    let result = import_period_directory(&fixture_dir, &config, "本期")
        .expect("import synthetic S2 source directory");
    let mut actual = serde_json::to_value(result).expect("serialize Rust import result");
    normalize_source_paths(&mut actual);

    let baseline: Value = serde_json::from_slice(
        &fs::read(workspace_path(
            "probe/reports/s2_import_python_baseline.json",
        ))
        .expect("read source Python importer baseline"),
    )
    .expect("decode Python baseline");
    assert_eq!(actual, baseline["result"]);
}
