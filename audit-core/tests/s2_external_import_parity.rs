use audit_core::s2_external::{S2ExternalData, S2ExternalValueKey};
use audit_core::s2_external_import::{import_external_workbook, S2ExternalImportResult};
use serde::Deserialize;
use serde_json::{json, Value};
use std::{fs, path::Path};

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
    file: String,
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

fn sorted_strings(values: impl Iterator<Item = String>) -> Vec<String> {
    let mut values = values.collect::<Vec<_>>();
    values.sort();
    values
}

fn canonical_data(data: &S2ExternalData) -> Value {
    let mut values = data
        .values
        .iter()
        .map(|item| serde_json::to_value(item).expect("serialize external value"))
        .collect::<Vec<_>>();
    for item in &mut values {
        normalize_source_paths(item);
    }
    values.sort_by_key(|item| {
        (
            item["data_date"].as_str().unwrap_or_default().to_owned(),
            item["row_label"].as_str().unwrap_or_default().to_owned(),
            item["indicator_name"]
                .as_str()
                .unwrap_or_default()
                .to_owned(),
        )
    });

    let mut date_labels = data
        .date_labels
        .iter()
        .map(|(date, label)| vec![date.clone(), label.clone()])
        .collect::<Vec<_>>();
    date_labels.sort();
    let mut blank_values = data
        .blank_values
        .iter()
        .map(|key| match key {
            S2ExternalValueKey::Plain(label, indicator) => vec![label.clone(), indicator.clone()],
            S2ExternalValueKey::TimeSeries(date, label, indicator) => {
                vec![date.clone(), label.clone(), indicator.clone()]
            }
        })
        .collect::<Vec<_>>();
    blank_values.sort();

    let mut fallback =
        serde_json::to_value(&data.missing_label_fallback).expect("serialize external fallback");
    normalize_source_paths(&mut fallback);
    json!({
        "time_series": data.time_series,
        "values": values,
        "available_dates": sorted_strings(data.available_dates.iter().cloned()),
        "date_labels": date_labels,
        "row_labels": sorted_strings(data.row_labels.iter().cloned()),
        "indicator_names": sorted_strings(data.indicator_names.iter().cloned()),
        "blank_values": blank_values,
        "missing_label_fallback": fallback,
    })
}

fn actual_result(result: Result<S2ExternalImportResult, anyhow::Error>) -> Value {
    match result {
        Ok(result) => {
            let mut payload = canonical_data(&result.data);
            normalize_source_paths(&mut payload);
            payload["warnings"] = json!(result.warnings);
            payload["external_org_map"] =
                serde_json::to_value(result.external_org_map).expect("serialize organization map");
            json!({"result": payload})
        }
        Err(error) => json!({"error": error.to_string()}),
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
fn s2_external_workbook_import_matches_read_only_openpyxl_baseline() {
    let fixtures: FixtureRoot = serde_json::from_slice(
        &fs::read(workspace_path(
            "tests/fixtures/s2_external_import/cases.json",
        ))
        .expect("read S2 external workbook fixtures"),
    )
    .expect("decode S2 external workbook fixtures");
    let baseline: Value = serde_json::from_slice(
        &fs::read(workspace_path(
            "probe/reports/s2_external_import_python_baseline.json",
        ))
        .expect("read Python external workbook baseline"),
    )
    .expect("decode Python external workbook baseline");

    for case in fixtures.cases {
        let file = workspace_path(&format!("tests/fixtures/s2_external_import/{}", case.file));
        let actual = actual_result(import_external_workbook(file));
        assert_json_equal(
            &actual,
            &baseline["result"][&case.id],
            &format!("result.{}", case.id),
        );
    }
}
