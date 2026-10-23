use audit_core::template::{
    clean_rule_comment, normalize_template_name, parse_comment_rule_id, parse_formula_result,
    parse_rule_rows,
};
use serde_json::{json, Value};

fn cases() -> Value {
    serde_json::from_str(include_str!("fixtures/template_parser_cases.json"))
        .expect("valid template parser cases")
}

fn baseline() -> Value {
    serde_json::from_str(include_str!(
        "../probe/reports/template_parser_python_baseline.json"
    ))
    .expect("valid template parser Python baseline")
}

#[test]
fn template_parser_helpers_match_python_baseline() {
    let cases = cases();
    let baseline = baseline();

    for (case, expected) in cases["normalize_names"]
        .as_array()
        .unwrap()
        .iter()
        .zip(baseline["normalize_names"].as_array().unwrap())
    {
        assert_eq!(
            normalize_template_name(case.as_str().unwrap()),
            expected["output"].as_str().unwrap(),
            "normalize name {:?}",
            case
        );
    }

    for (case, expected) in cases["comment_cases"]
        .as_array()
        .unwrap()
        .iter()
        .zip(baseline["comment_cases"].as_array().unwrap())
    {
        let text = case["text"].as_str().unwrap();
        let rule_id = match parse_comment_rule_id(text) {
            Ok(value) => json!({"value": value}),
            Err(error) => json!({"error": error.to_string()}),
        };
        assert_eq!(
            rule_id, expected["rule_id"],
            "comment case {}",
            case["name"]
        );
        assert_eq!(
            clean_rule_comment(text),
            expected["cleaned"],
            "clean comment case {}",
            case["name"]
        );
    }

    for (case, expected) in cases["formula_cases"]
        .as_array()
        .unwrap()
        .iter()
        .zip(baseline["formula_cases"].as_array().unwrap())
    {
        let result = parse_formula_result(
            &case["value"],
            case.get("default_severity").and_then(Value::as_str),
            case.get("default_message").and_then(Value::as_str),
        );
        assert_eq!(
            serde_json::to_value(result).unwrap(),
            expected["result"],
            "formula case {}",
            case["name"]
        );
    }

    for (case, expected) in cases["rule_row_cases"]
        .as_array()
        .unwrap()
        .iter()
        .zip(baseline["rule_row_cases"].as_array().unwrap())
    {
        let result = match parse_rule_rows(case["rows"].as_array().unwrap()) {
            Ok(rules) => json!({"rules": rules}),
            Err(error) => json!({"error": error.to_string()}),
        };
        assert_eq!(result, expected["result"], "rule case {}", case["name"]);
    }
}
