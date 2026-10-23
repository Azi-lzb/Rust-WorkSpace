use audit_core::s2_external::{compare_external, S2ExternalConfig, S2ExternalData};
use audit_core::s2_import::IndicatorRecord;
use serde::Deserialize;
use serde_json::Value;
use std::{collections::HashMap, fs, path::Path};

fn workspace_path(relative: &str) -> std::path::PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join(relative)
}

#[derive(Deserialize)]
struct FixtureRoot {
    cases: Vec<FixtureCase>,
}

#[derive(Deserialize)]
struct FixtureCase {
    id: String,
    current: Vec<IndicatorRecord>,
    external_org_map: HashMap<String, String>,
    external: S2ExternalData,
    config: S2ExternalConfig,
}

fn normalize_source_paths(value: &mut Value) {
    match value {
        Value::Array(items) => {
            for item in items {
                normalize_source_paths(item);
            }
        }
        Value::Object(fields) => {
            if let Some(Value::String(path)) = fields.get_mut("source_file") {
                *path = path
                    .rsplit(['/', '\\'])
                    .next()
                    .expect("source filename")
                    .to_owned();
            }
            for field in fields.values_mut() {
                normalize_source_paths(field);
            }
        }
        _ => {}
    }
}

fn assert_json_equal(actual: &Value, expected: &Value, path: &str) {
    match (actual, expected) {
        (Value::Object(actual), Value::Object(expected)) => {
            assert_eq!(
                actual.keys().collect::<Vec<_>>(),
                expected.keys().collect::<Vec<_>>(),
                "keys differ at {path}"
            );
            for (key, actual_value) in actual {
                assert_json_equal(actual_value, &expected[key], &format!("{path}.{key}"));
            }
        }
        (Value::Array(actual), Value::Array(expected)) => {
            assert_eq!(
                actual.len(),
                expected.len(),
                "array length differs at {path}"
            );
            for (index, (actual_value, expected_value)) in actual.iter().zip(expected).enumerate() {
                assert_json_equal(actual_value, expected_value, &format!("{path}[{index}]"));
            }
        }
        _ => assert_eq!(actual, expected, "value differs at {path}"),
    }
}

#[test]
fn s2_external_comparison_matches_read_only_python_baseline() {
    let fixture: FixtureRoot = serde_json::from_slice(
        &fs::read(workspace_path("tests/fixtures/s2_external/cases.json"))
            .expect("read S2 external fixture"),
    )
    .expect("decode S2 external fixture");
    let baseline: Value = serde_json::from_slice(
        &fs::read(workspace_path(
            "probe/reports/s2_external_python_baseline.json",
        ))
        .expect("read Python external baseline"),
    )
    .expect("decode Python external baseline");

    let mut actual = serde_json::Map::new();
    for case in fixture.cases {
        let result = match compare_external(
            &case.current,
            &case.external,
            &case.external_org_map,
            &case.config,
        ) {
            Ok(findings) => serde_json::json!({"findings": findings}),
            Err(error) => serde_json::json!({"error": error.to_string()}),
        };
        actual.insert(case.id, result);
    }
    let mut actual = Value::Object(actual);
    normalize_source_paths(&mut actual);
    assert_json_equal(&actual, &baseline["result"], "result");
}
