use audit_core::unit_conversion::{
    as_decimal, convert_amount, convert_amount_from_cell, decimal_number_from_cell,
    is_amount_data_type, to_yuan, AmountSourceFormat, DecimalTuple, UnitConversionError,
};
use excel_reader::CellValue;
use serde_json::{json, Value};

fn fixture() -> Value {
    serde_json::from_str(include_str!("fixtures/unit_conversion_cases.json"))
        .expect("valid unit conversion fixture")
}

fn baseline() -> Value {
    serde_json::from_str(include_str!(
        "../probe/reports/unit_conversion_python_baseline.json"
    ))
    .expect("valid Python unit conversion baseline")
}

fn outcome(result: Result<DecimalTuple, UnitConversionError>) -> Value {
    match result {
        Ok(value) => json!({ "status": "ok", "tuple": value }),
        Err(error) => json!({
            "status": "error",
            "type": "ValueError",
            "message": error.to_string(),
        }),
    }
}

fn optional_outcome(result: Result<Option<DecimalTuple>, UnitConversionError>) -> Value {
    match result {
        Ok(Some(value)) => json!({ "status": "ok", "tuple": value }),
        Ok(None) => json!({ "status": "skipped" }),
        Err(error) => json!({
            "status": "error",
            "type": "ValueError",
            "message": error.to_string(),
        }),
    }
}

fn source_format(value: &str) -> AmountSourceFormat {
    match value {
        "xls" => AmountSourceFormat::Xls,
        "xlsx" => AmountSourceFormat::XlsxLike,
        other => panic!("unknown source format: {other}"),
    }
}

#[test]
fn s2_amount_unit_helpers_match_python_decimal_baseline() {
    let fixture = fixture();
    let baseline = baseline();

    let amount_type_cases = fixture["amount_data_type_cases"].as_array().unwrap();
    let amount_type_baseline = baseline["amount_data_type_cases"].as_array().unwrap();
    assert_eq!(amount_type_cases.len(), amount_type_baseline.len());
    for (case, expected) in amount_type_cases.iter().zip(amount_type_baseline) {
        assert_eq!(
            json!({
                "id": case["id"],
                "value": case["value"],
                "output": is_amount_data_type(case["value"].as_str().unwrap()),
            }),
            *expected,
            "amount data type case {}",
            case["id"]
        );
    }

    let parse_cases = fixture["as_decimal_cases"].as_array().unwrap();
    let parse_baseline = baseline["as_decimal_cases"].as_array().unwrap();
    assert_eq!(parse_cases.len(), parse_baseline.len());
    for (case, expected) in parse_cases.iter().zip(parse_baseline) {
        let value = case["value"].as_str().unwrap();
        assert_eq!(
            json!({
                "id": case["id"],
                "value": case["value"],
                "output": outcome(as_decimal(value)),
            }),
            *expected,
            "decimal parse case {}",
            case["id"]
        );
    }

    let conversion_cases = fixture["conversion_cases"].as_array().unwrap();
    let conversion_baseline = baseline["conversion_cases"].as_array().unwrap();
    assert_eq!(conversion_cases.len(), conversion_baseline.len());
    for (case, expected) in conversion_cases.iter().zip(conversion_baseline) {
        let value = case["value"].as_str().unwrap();
        let source_unit = case["source_unit"].as_str().unwrap();
        let target_unit = case["target_unit"].as_str().unwrap();
        assert_eq!(
            json!({
                "id": case["id"],
                "value": case["value"],
                "source_unit": case["source_unit"],
                "target_unit": case["target_unit"],
                "output": outcome(convert_amount(value, source_unit, target_unit)),
            }),
            *expected,
            "amount conversion case {}",
            case["id"]
        );
    }

    let to_yuan_cases = fixture["to_yuan_cases"].as_array().unwrap();
    let to_yuan_baseline = baseline["to_yuan_cases"].as_array().unwrap();
    assert_eq!(to_yuan_cases.len(), to_yuan_baseline.len());
    for (case, expected) in to_yuan_cases.iter().zip(to_yuan_baseline) {
        let value = case["value"].as_str().unwrap();
        let unit = case["unit"].as_str().unwrap();
        assert_eq!(
            json!({
                "id": case["id"],
                "value": case["value"],
                "unit": case["unit"],
                "output": outcome(to_yuan(value, unit)),
            }),
            *expected,
            "to_yuan case {}",
            case["id"]
        );
    }

    let cell_cases = fixture["reader_cell_cases"].as_array().unwrap();
    let cell_baseline = baseline["reader_cell_cases"].as_array().unwrap();
    assert_eq!(cell_cases.len(), cell_baseline.len());
    for (case, expected) in cell_cases.iter().zip(cell_baseline) {
        let cell: Option<CellValue> = if case["cell"].is_null() {
            None
        } else {
            Some(CellValue::from_probe_json(&case["cell"]).unwrap())
        };
        let format = source_format(case["source_format"].as_str().unwrap());
        let decimal_output = optional_outcome(decimal_number_from_cell(cell.as_ref(), format));
        let conversion_output = optional_outcome(convert_amount_from_cell(
            cell.as_ref(),
            format,
            case["source_unit"].as_str().unwrap(),
            case["target_unit"].as_str().unwrap(),
        ));

        assert_eq!(
            decimal_output, expected["decimal_output"],
            "cell parse case {}",
            case["id"]
        );
        assert_eq!(
            conversion_output, expected["conversion_output"],
            "cell conversion case {}",
            case["id"]
        );
    }
}
