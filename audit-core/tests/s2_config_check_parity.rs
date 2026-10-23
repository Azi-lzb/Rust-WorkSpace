use audit_core::s2_config_check::{check_s2_config, format_config_report};
use serde::Deserialize;
use serde_json::Value;
use std::{collections::BTreeMap, fs, path::Path};

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

fn normalize_paths(value: &mut Value, path: &Path) {
    match value {
        Value::String(text) => *text = normalize_path_text(text, path),
        Value::Array(values) => {
            for value in values {
                normalize_paths(value, path);
            }
        }
        Value::Object(fields) => {
            for value in fields.values_mut() {
                normalize_paths(value, path);
            }
        }
        _ => {}
    }
}

fn normalize_path_text(text: &str, path: &Path) -> String {
    let filename = path
        .file_name()
        .and_then(|name| name.to_str())
        .unwrap_or_default();
    text.replace(&path.to_string_lossy().to_string(), filename)
        .replace(&path.to_string_lossy().replace('/', "\\"), filename)
}

fn assert_json_equal(actual: &Value, expected: &Value, path: &str) {
    match (actual, expected) {
        (Value::Object(actual), Value::Object(expected)) => {
            let actual_keys = actual.keys().collect::<Vec<_>>();
            let expected_keys = expected.keys().collect::<Vec<_>>();
            assert_eq!(actual_keys, expected_keys, "keys differ at {path}");
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
fn s2_configuration_checks_match_read_only_python_baseline() {
    let fixture_dir = workspace_path("tests/fixtures/s2_config");
    let index: FixtureIndex = serde_json::from_slice(
        &fs::read(fixture_dir.join("check_cases.json")).expect("read S2 checker fixture index"),
    )
    .expect("decode S2 checker fixture index");
    let baseline: Value = serde_json::from_slice(
        &fs::read(workspace_path(
            "probe/reports/s2_config_check_python_baseline.json",
        ))
        .expect("read S2 checker Python baseline"),
    )
    .expect("decode S2 checker Python baseline");
    let case_count = index.cases.len();

    let mut reports = BTreeMap::<String, Value>::new();
    for case in index.cases {
        let path = fixture_dir.join(&case.file);
        let before = path
            .exists()
            .then(|| fs::read(&path).expect("read fixture before check"));
        let report = check_s2_config(&path);
        let formatted = format_config_report(&report);
        let mut actual_report = serde_json::to_value(&report).expect("serialize S2 checker report");
        normalize_paths(&mut actual_report, &path);
        let actual_formatted = normalize_path_text(&formatted, &path);
        let actual = serde_json::json!({
            "report": actual_report,
            "formatted": actual_formatted,
        });
        assert_json_equal(
            &actual,
            &baseline["result"][&case.id],
            &format!("result.{}", case.id),
        );
        if let Some(before) = before {
            assert_eq!(
                fs::read(&path).expect("read fixture after check"),
                before,
                "checker modified fixture {}",
                case.file
            );
        }
        reports.insert(case.id, actual["report"].clone());
    }

    assert_eq!(reports["check_valid"]["passed"], true);
    assert_eq!(reports["check_legacy_severity"]["passed"], true);
    assert_eq!(reports["check_legacy_severity"]["warnings"], 1);
    let gap_report = &reports["check_gap_duplicate"];
    assert_eq!(gap_report["passed"], false);
    assert!(gap_report["issues"]
        .as_array()
        .unwrap()
        .iter()
        .any(|issue| { issue["message"].as_str().unwrap().contains("存在空档") }));
    assert!(gap_report["issues"]
        .as_array()
        .unwrap()
        .iter()
        .any(|issue| {
            issue["message"]
                .as_str()
                .unwrap()
                .contains("策略组+规则编号重复")
        }));
    let datetime_report = &reports["check_datetime_indicator_flag"];
    assert!(datetime_report["issues"]
        .as_array()
        .unwrap()
        .iter()
        .any(|issue| {
            issue["message"]
                .as_str()
                .unwrap()
                .contains("datetime.datetime(2024, 1, 2, 0, 0)")
        }));
    let missing = &reports["missing_file"];
    assert_eq!(missing["checks"][0]["status"], "failed");
    assert!(missing["checks"].as_array().unwrap()[1..]
        .iter()
        .all(|check| check["status"] == "not_run"));

    let cases = reports
        .iter()
        .map(|(id, report)| {
            serde_json::json!({
                "id": id,
                "status": "passed",
                "business_check_passed": report["passed"],
                "blocking_errors": report["blocking_errors"],
                "warnings": report["warnings"],
            })
        })
        .collect::<Vec<_>>();
    let parity_report = serde_json::json!({
        "schema_version": 1,
        "status": "passed",
        "baseline": "s2_config_check_python_baseline.json",
        "source_sha256": baseline["source"]["source_sha256"],
        "fixture": {
            "path": "tests/fixtures/s2_config/check_cases.json",
            "case_count": case_count,
        },
        "coverage": {
            "structured_report": "passed",
            "formatted_report": "passed",
            "input_fixtures_unchanged": "passed",
        },
        "cases": cases,
        "scope": "Synthetic S2 configuration workbooks only; does not establish parity for real business configurations, service orchestration, report writing, or target Linux systems.",
    });
    let report_path = workspace_path("probe/reports/s2_config_check_parity.json");
    fs::create_dir_all(report_path.parent().expect("report parent"))
        .expect("create S2 checker report directory");
    fs::write(
        report_path,
        serde_json::to_vec_pretty(&parity_report).expect("serialize S2 checker parity report"),
    )
    .expect("write S2 checker parity report");
}
