use audit_core::{
    s2_import::{ImportedValue, IndicatorRecord, S2SourceRef},
    s2_rule_engine::{evaluate_rules, S2RuleEngineConfig},
};
use serde::Deserialize;
use serde_json::Value;
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
    current: Vec<RecordInput>,
    previous: Vec<RecordInput>,
    config: S2RuleEngineConfig,
}

#[derive(Deserialize)]
#[serde(default)]
struct RecordInput {
    institution_id: String,
    institution_name: String,
    social_credit_code: String,
    institution_type: String,
    handling_branch: String,
    region: String,
    form_code: String,
    indicator_code: String,
    indicator_name: String,
    period: String,
    source_date: String,
    value_type: String,
    source_unit: String,
    normalized_unit: String,
    normalized_value: ImportedValue,
    source_file: String,
    source_sheet: String,
    source_row: usize,
}

impl Default for RecordInput {
    fn default() -> Self {
        Self {
            institution_id: "9144".to_owned(),
            institution_name: "甲银行".to_owned(),
            social_credit_code: "C9144".to_owned(),
            institution_type: "农商行".to_owned(),
            handling_branch: "一部".to_owned(),
            region: "示例地区".to_owned(),
            form_code: "20202".to_owned(),
            indicator_code: String::new(),
            indicator_name: String::new(),
            period: String::new(),
            source_date: String::new(),
            value_type: "数值".to_owned(),
            source_unit: "元".to_owned(),
            normalized_unit: "万元".to_owned(),
            normalized_value: ImportedValue::Empty,
            source_file: "synthetic.xlsx".to_owned(),
            source_sheet: String::new(),
            source_row: 1,
        }
    }
}

impl RecordInput {
    fn into_record(self) -> IndicatorRecord {
        let source_date = if self.source_date.is_empty() {
            format!("{}-30", self.period)
        } else {
            self.source_date
        };
        let indicator_name = if self.indicator_name.is_empty() {
            self.indicator_code.clone()
        } else {
            self.indicator_name
        };
        let source_sheet = if self.source_sheet.is_empty() {
            self.form_code.clone()
        } else {
            self.source_sheet
        };
        IndicatorRecord {
            institution_id: self.institution_id,
            institution_name: self.institution_name,
            social_credit_code: self.social_credit_code,
            institution_type: self.institution_type,
            handling_branch: self.handling_branch,
            region: self.region,
            form_code: self.form_code,
            indicator_code: self.indicator_code,
            indicator_name,
            period: self.period,
            source_date,
            value: self.normalized_value.clone(),
            value_type: self.value_type,
            source_unit: self.source_unit,
            normalized_value: self.normalized_value,
            normalized_unit: self.normalized_unit,
            source: S2SourceRef {
                source_file: self.source_file,
                source_sheet,
                source_row: self.source_row,
            },
        }
    }
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
fn s2_rule_engine_matches_read_only_python_baseline() {
    let fixture: FixtureRoot = serde_json::from_slice(
        &fs::read(workspace_path("tests/fixtures/s2_rule_engine/cases.json"))
            .expect("read S2 rule-engine fixture"),
    )
    .expect("decode S2 rule-engine fixture");
    let baseline: Value = serde_json::from_slice(
        &fs::read(workspace_path(
            "probe/reports/s2_rule_engine_python_baseline.json",
        ))
        .expect("read Python rule-engine baseline"),
    )
    .expect("decode Python rule-engine baseline");

    let mut actual = serde_json::Map::new();
    for case in fixture.cases {
        let current = case
            .current
            .into_iter()
            .map(RecordInput::into_record)
            .collect::<Vec<_>>();
        let previous = case
            .previous
            .into_iter()
            .map(RecordInput::into_record)
            .collect::<Vec<_>>();
        let result = match evaluate_rules(&current, &previous, &case.config) {
            Ok((findings, quality)) => serde_json::json!({
                "findings": findings,
                "quality": quality,
            }),
            Err(error) => serde_json::json!({"error": error.to_string()}),
        };
        actual.insert(case.id, result);
    }
    let mut actual = Value::Object(actual);
    normalize_source_paths(&mut actual);
    assert_json_equal(&actual, &baseline["result"], "result");
}
