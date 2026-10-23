use audit_core::name_config::{
    features_of_type, is_limited, matches_named_range, split_names, FeatureMapping,
};
use serde_json::{json, Value};

fn cases() -> Value {
    serde_json::from_str(include_str!("fixtures/name_config_helper_cases.json"))
        .expect("valid name config helper cases")
}

fn baseline() -> Value {
    serde_json::from_str(include_str!(
        "../probe/reports/name_config_helpers_python_baseline.json"
    ))
    .expect("valid name config helper Python baseline")
}

fn mapping(value: &Value) -> FeatureMapping {
    serde_json::from_value(value.clone()).expect("valid feature mapping")
}

#[test]
fn name_config_helpers_match_python_baseline() {
    let cases = cases();
    let baseline = baseline();
    let mappings = cases["feature_mappings"]
        .as_array()
        .unwrap()
        .iter()
        .map(mapping)
        .collect::<Vec<_>>();

    let mapping_values = mappings
        .iter()
        .map(serde_json::to_value)
        .collect::<Result<Vec<_>, _>>()
        .unwrap();
    assert_eq!(json!(mapping_values), baseline["feature_mappings"]);

    for (case, expected) in cases["split_names"]
        .as_array()
        .unwrap()
        .iter()
        .zip(baseline["split_names"].as_array().unwrap())
    {
        let input = case.as_str().unwrap();
        assert_eq!(
            json!(split_names(input)),
            expected["output"],
            "split {input:?}"
        );
    }

    for (case, expected) in cases["limited_values"]
        .as_array()
        .unwrap()
        .iter()
        .zip(baseline["limited_values"].as_array().unwrap())
    {
        let input = case.as_str().unwrap();
        assert_eq!(
            json!(is_limited(input)),
            expected["output"],
            "limit {input:?}"
        );
    }

    for (feature_type, expected) in cases["features_of_type"]
        .as_array()
        .unwrap()
        .iter()
        .zip(baseline["features_of_type"].as_array().unwrap())
    {
        let feature_type = feature_type.as_str().unwrap();
        let indices = features_of_type(&mappings, feature_type)
            .iter()
            .filter_map(|feature| {
                mappings
                    .iter()
                    .position(|item| std::ptr::eq(item, *feature))
            })
            .collect::<Vec<_>>();
        assert_eq!(json!(indices), expected["indices"], "type {feature_type:?}");
    }

    for (case, expected) in cases["named_range_cases"]
        .as_array()
        .unwrap()
        .iter()
        .zip(baseline["named_range_cases"].as_array().unwrap())
    {
        let mapping_index = case["mapping_index"].as_u64().unwrap() as usize;
        let candidate = case["candidate"].as_str().unwrap();
        assert_eq!(
            json!(matches_named_range(&mappings[mapping_index], candidate)),
            expected["output"],
            "named range {candidate:?}"
        );
    }
}
