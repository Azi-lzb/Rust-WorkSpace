use audit_core::s2_config::{load_config, load_unit_settings, S2AuditConfig};
use serde::Deserialize;
use serde_json::Value;
use std::{fs, path::Path};

fn workspace_path(relative: &str) -> std::path::PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join(relative)
}

#[derive(Deserialize)]
struct FixtureIndex {
    cases: Vec<FixtureCase>,
}

#[derive(Deserialize)]
struct FixtureCase {
    id: String,
    file: String,
}

fn normalize_sources(value: &mut Value) {
    match value {
        Value::Array(items) => {
            for item in items {
                normalize_sources(item);
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
                normalize_sources(field);
            }
        }
        _ => {}
    }
}

fn normalize_error(error: &str, path: &Path) -> String {
    let filename = path
        .file_name()
        .and_then(|name| name.to_str())
        .unwrap_or_default();
    error
        .replace(&path.to_string_lossy().to_string(), filename)
        .replace(&path.to_string_lossy().replace('/', "\\"), filename)
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
fn s2_config_loading_matches_read_only_python_baseline() {
    let fixture_dir = workspace_path("tests/fixtures/s2_config");
    let index: FixtureIndex = serde_json::from_slice(
        &fs::read(fixture_dir.join("cases.json")).expect("read S2 config fixture index"),
    )
    .expect("decode S2 config fixture index");
    let baseline: Value = serde_json::from_slice(
        &fs::read(workspace_path(
            "probe/reports/s2_config_python_baseline.json",
        ))
        .expect("read S2 config Python baseline"),
    )
    .expect("decode S2 config Python baseline");

    let mut actual = serde_json::Map::new();
    for case in index.cases {
        let path = fixture_dir.join(&case.file);
        let mut result = serde_json::Map::new();
        match load_config(&path) {
            Ok(config) => {
                result.insert(
                    "config".to_owned(),
                    serde_json::to_value(config).expect("serialize loaded S2 config"),
                );
            }
            Err(error) => {
                result.insert(
                    "config_error".to_owned(),
                    Value::String(normalize_error(&error.to_string(), &path)),
                );
            }
        }
        match load_unit_settings(&path) {
            Ok(settings) => {
                result.insert(
                    "unit_settings".to_owned(),
                    serde_json::to_value(settings).expect("serialize S2 unit settings"),
                );
            }
            Err(error) => {
                result.insert(
                    "unit_settings_error".to_owned(),
                    Value::String(normalize_error(&error.to_string(), &path)),
                );
            }
        }
        actual.insert(case.id, Value::Object(result));
    }
    let mut actual = Value::Object(actual);
    normalize_sources(&mut actual);
    assert_json_equal(&actual, &baseline["result"], "result");
}

#[test]
fn loaded_configuration_projects_into_existing_s2_engine_contracts() {
    let path = workspace_path("tests/fixtures/s2_config/comprehensive.xlsx");
    let config: S2AuditConfig = load_config(path).expect("load comprehensive config fixture");

    let import = config.import_config();
    assert_eq!(import.indicators["20202001"].indicator_name, "存款余额");
    assert_eq!(import.institutions_by_id["9144"].institution_id, "9144");
    assert_eq!(import.unit_settings.source_unit, "亿元");

    let periods = config.period_config();
    assert_eq!(
        periods.indicators["20202001"].period_strategy_group,
        "BAL01"
    );
    assert_eq!(periods.bands.len(), config.bands.len());
    assert_eq!(periods.current_missing_severity, "软性");

    let rules = config.rule_engine_config();
    assert_eq!(rules.rules.len(), config.rules.len());
    assert_eq!(rules.unit_settings.output_file_unit, "元");
    assert_eq!(rules.rules[0].rule_id, "R001");
    assert_eq!(rules.rules[0].source.as_ref().unwrap().source_row, 2);

    let external = config.external_config();
    assert_eq!(external.external_rules.len(), config.external_rules.len());
    assert_eq!(external.external_rules[0].external_system, "大集中");
    assert_eq!(
        external.external_rules[0]
            .source
            .as_ref()
            .unwrap()
            .source_row,
        2
    );
    assert_eq!(external.unit_settings.external_file_unit, "万元");
}
