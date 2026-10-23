use audit_core::{
    s2_analysis::analyze_imports,
    s2_config::{S2AuditConfig, S2ConfigIndicator, S2ConfigInstitution, S2ConfigUnitSettings},
    s2_external::S2ExternalRule,
    s2_external_import::S2ExternalImportResult,
    s2_import::{ImportedValue, IndicatorRecord, S2ImportResult, S2SourceRef},
    s2_period::S2PeriodBandConfig,
    s2_rule_engine::S2Rule,
};
use serde::Deserialize;
use serde_json::{Map, Value};
use std::{collections::HashMap, fs, path::Path};

#[derive(Deserialize)]
struct FixtureRoot {
    config: ConfigInput,
    cases: Vec<CaseInput>,
}

#[derive(Deserialize)]
struct ConfigInput {
    #[serde(default)]
    indicators: HashMap<String, IndicatorInput>,
    #[serde(default)]
    institutions_by_id: HashMap<String, InstitutionInput>,
    #[serde(default)]
    institutions_by_name: HashMap<String, InstitutionInput>,
    #[serde(default)]
    bands: Vec<S2PeriodBandConfig>,
    #[serde(default)]
    external_rules: Vec<S2ExternalRule>,
    #[serde(default)]
    rules: Vec<S2Rule>,
    #[serde(default)]
    warnings: Vec<String>,
    #[serde(default)]
    unit_settings: UnitSettingsInput,
    #[serde(default)]
    hide_unchanged_period_rows: bool,
    #[serde(default)]
    hide_matching_external_rows: bool,
    #[serde(default = "default_severity")]
    text_change_severity: String,
    #[serde(default = "default_severity")]
    current_missing_severity: String,
    #[serde(default = "default_severity")]
    previous_missing_severity: String,
}

#[derive(Deserialize)]
struct IndicatorInput {
    #[serde(default)]
    indicator_name: String,
    #[serde(default)]
    form_code: String,
    #[serde(default)]
    data_type: String,
    #[serde(default)]
    disabled: bool,
    #[serde(default)]
    note: String,
    #[serde(default = "default_true")]
    period_enabled: bool,
    #[serde(default)]
    period_strategy_group: String,
    #[serde(default)]
    min_change_value: Option<f64>,
}

#[derive(Deserialize)]
struct InstitutionInput {
    #[serde(default)]
    institution_id: String,
    #[serde(default)]
    institution_name: String,
    #[serde(default)]
    social_credit_code: String,
    #[serde(default)]
    institution_type: String,
    #[serde(default)]
    handling_branch: String,
    #[serde(default)]
    region: String,
    #[serde(default)]
    report_item: String,
}

#[derive(Deserialize)]
struct UnitSettingsInput {
    #[serde(default = "default_yuan")]
    source_unit: String,
    #[serde(default = "default_yuan")]
    external_file_unit: String,
    #[serde(default = "default_yuan")]
    output_file_unit: String,
}

#[derive(Deserialize)]
struct CaseInput {
    id: String,
    #[serde(default)]
    current: Vec<RecordInput>,
    #[serde(default)]
    previous: Vec<RecordInput>,
    #[serde(default)]
    current_duplicates: Vec<DuplicateInput>,
    #[serde(default)]
    current_warnings: Vec<String>,
    #[serde(default)]
    previous_warnings: Vec<String>,
    external: Option<ExternalInput>,
}

#[derive(Deserialize)]
struct RecordInput {
    indicator_code: String,
    period: String,
    normalized_value: ImportedValue,
    #[serde(default)]
    institution_id: String,
    #[serde(default)]
    institution_name: String,
    #[serde(default)]
    social_credit_code: String,
    #[serde(default)]
    institution_type: String,
    #[serde(default)]
    handling_branch: String,
    #[serde(default)]
    region: String,
    #[serde(default)]
    form_code: String,
    #[serde(default)]
    indicator_name: String,
    #[serde(default)]
    source_date: String,
    #[serde(default)]
    value_type: String,
    #[serde(default)]
    source_unit: String,
    #[serde(default)]
    normalized_unit: String,
    #[serde(default)]
    source_file: String,
    #[serde(default)]
    source_sheet: String,
    #[serde(default)]
    source_row: usize,
}

#[derive(Deserialize)]
struct DuplicateInput {
    first: RecordInput,
    duplicate: RecordInput,
}

#[derive(Deserialize)]
struct ExternalInput {
    data: audit_core::s2_external::S2ExternalData,
    #[serde(default)]
    external_org_map: HashMap<String, String>,
    #[serde(default)]
    warnings: Vec<String>,
}

fn default_severity() -> String {
    "硬性".to_owned()
}

fn default_true() -> bool {
    true
}

fn default_yuan() -> String {
    "元".to_owned()
}

impl Default for UnitSettingsInput {
    fn default() -> Self {
        Self {
            source_unit: default_yuan(),
            external_file_unit: default_yuan(),
            output_file_unit: default_yuan(),
        }
    }
}

fn make_config(input: ConfigInput) -> S2AuditConfig {
    let indicators = input
        .indicators
        .into_iter()
        .map(|(code, item)| {
            let value = S2ConfigIndicator {
                indicator_code: code.clone(),
                indicator_name: item.indicator_name,
                form_code: item.form_code,
                data_type: item.data_type,
                disabled: item.disabled,
                note: item.note,
                period_enabled: item.period_enabled,
                period_strategy_group: item.period_strategy_group,
                min_change_value: item.min_change_value,
            };
            (code, value)
        })
        .collect();
    let institutions_by_id = input
        .institutions_by_id
        .into_iter()
        .map(|(code, item)| (code, institution(item)))
        .collect();
    let institutions_by_name = input
        .institutions_by_name
        .into_iter()
        .map(|(name, item)| (name, institution(item)))
        .collect();

    S2AuditConfig {
        indicators,
        institutions_by_id,
        institutions_by_name,
        bands: input.bands,
        external_rules: input.external_rules,
        rules: input.rules,
        warnings: input.warnings,
        unit_settings: S2ConfigUnitSettings {
            source_unit: input.unit_settings.source_unit,
            external_file_unit: input.unit_settings.external_file_unit,
            output_file_unit: input.unit_settings.output_file_unit,
        },
        hide_unchanged_period_rows: input.hide_unchanged_period_rows,
        hide_matching_external_rows: input.hide_matching_external_rows,
        text_change_severity: input.text_change_severity,
        current_missing_severity: input.current_missing_severity,
        previous_missing_severity: input.previous_missing_severity,
    }
}

fn institution(input: InstitutionInput) -> S2ConfigInstitution {
    S2ConfigInstitution {
        institution_id: input.institution_id,
        institution_name: input.institution_name,
        social_credit_code: input.social_credit_code,
        institution_type: input.institution_type,
        handling_branch: input.handling_branch,
        region: input.region,
        report_item: input.report_item,
    }
}

fn record(input: RecordInput) -> IndicatorRecord {
    let period = input.period;
    let form_code = if input.form_code.is_empty() {
        "20202".to_owned()
    } else {
        input.form_code
    };
    let source_date = if input.source_date.is_empty() {
        format!("{period}-30")
    } else {
        input.source_date
    };
    let value_type = if input.value_type.is_empty() {
        "数值".to_owned()
    } else {
        input.value_type
    };
    let source_file = if input.source_file.is_empty() {
        "synthetic.xlsx".to_owned()
    } else {
        input.source_file
    };
    let source_sheet = if input.source_sheet.is_empty() {
        form_code.clone()
    } else {
        input.source_sheet
    };
    let source_row = input.source_row.max(1);
    let institution_id = if input.institution_id.is_empty() {
        "9144".to_owned()
    } else {
        input.institution_id
    };
    let institution_name = if input.institution_name.is_empty() {
        "甲银行".to_owned()
    } else {
        input.institution_name
    };
    let social_credit_code = if input.social_credit_code.is_empty() {
        "C9144".to_owned()
    } else {
        input.social_credit_code
    };
    let institution_type = if input.institution_type.is_empty() {
        "农商行".to_owned()
    } else {
        input.institution_type
    };
    let handling_branch = if input.handling_branch.is_empty() {
        "一部".to_owned()
    } else {
        input.handling_branch
    };
    let region = if input.region.is_empty() {
        "示例地区".to_owned()
    } else {
        input.region
    };
    let indicator_name = if input.indicator_name.is_empty() {
        input.indicator_code.clone()
    } else {
        input.indicator_name
    };
    let source_unit = if input.source_unit.is_empty() {
        "元".to_owned()
    } else {
        input.source_unit
    };
    let normalized_unit = if input.normalized_unit.is_empty() {
        "万元".to_owned()
    } else {
        input.normalized_unit
    };
    let value = input.normalized_value;
    IndicatorRecord {
        institution_id,
        institution_name,
        social_credit_code,
        institution_type,
        handling_branch,
        region,
        form_code: form_code.clone(),
        indicator_code: input.indicator_code,
        indicator_name,
        period,
        source_date,
        value: value.clone(),
        value_type,
        source_unit,
        normalized_value: value,
        normalized_unit,
        source: S2SourceRef {
            source_file,
            source_sheet,
            source_row,
        },
    }
}

fn import_result(
    records: Vec<RecordInput>,
    duplicates: Vec<DuplicateInput>,
    warnings: Vec<String>,
) -> S2ImportResult {
    S2ImportResult {
        records: records.into_iter().map(record).collect(),
        duplicates: duplicates
            .into_iter()
            .map(|pair| (record(pair.first), record(pair.duplicate)))
            .collect(),
        warnings,
    }
}

fn workspace_path(relative: &str) -> std::path::PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join(relative)
}

fn collect_json_differences(
    actual: &Value,
    expected: &Value,
    path: &str,
    differences: &mut Vec<String>,
) {
    if differences.len() >= 40 {
        return;
    }
    match (actual, expected) {
        (Value::Object(actual), Value::Object(expected)) => {
            for key in actual.keys().chain(expected.keys()) {
                if actual.get(key) != expected.get(key) {
                    match (actual.get(key), expected.get(key)) {
                        (Some(actual), Some(expected)) => collect_json_differences(
                            actual,
                            expected,
                            &format!("{path}.{key}"),
                            differences,
                        ),
                        (actual, expected) => differences.push(format!(
                            "{path}.{key}: Rust={actual:?}; Python={expected:?}"
                        )),
                    }
                    if differences.len() >= 40 {
                        return;
                    }
                }
            }
        }
        (Value::Array(actual), Value::Array(expected)) => {
            if actual.len() != expected.len() {
                differences.push(format!(
                    "{path}.length: Rust={}; Python={}",
                    actual.len(),
                    expected.len()
                ));
            }
            for (index, (actual, expected)) in actual.iter().zip(expected).enumerate() {
                collect_json_differences(
                    actual,
                    expected,
                    &format!("{path}[{index}]"),
                    differences,
                );
                if differences.len() >= 40 {
                    return;
                }
            }
        }
        _ if actual != expected => {
            differences.push(format!("{path}: Rust={actual:?}; Python={expected:?}"));
        }
        _ => {}
    }
}

#[test]
fn in_memory_analysis_matches_read_only_python_service_components() {
    let fixture: FixtureRoot = serde_json::from_slice(
        &fs::read(workspace_path("tests/fixtures/s2_analysis/cases.json"))
            .expect("read S2 analysis fixture"),
    )
    .expect("decode S2 analysis fixture");
    let baseline: Value = serde_json::from_slice(
        &fs::read(workspace_path(
            "probe/reports/s2_analysis_python_baseline.json",
        ))
        .expect("read S2 analysis Python baseline"),
    )
    .expect("decode S2 analysis Python baseline");
    let config = make_config(fixture.config);

    let mut results = Map::new();
    let mut case_reports = Vec::new();
    for case in fixture.cases {
        let current = import_result(case.current, case.current_duplicates, case.current_warnings);
        let previous = import_result(case.previous, Vec::new(), case.previous_warnings);
        let external = case.external.map(|input| S2ExternalImportResult {
            data: input.data,
            external_org_map: input.external_org_map,
            warnings: input.warnings,
            reader_engine: "fixture".to_owned(),
            reader_version: "fixture".to_owned(),
            reader_elapsed_seconds: 0.0,
        });
        let actual = analyze_imports(&current, &previous, &config, external.as_ref())
            .unwrap_or_else(|error| panic!("analysis failed for {}: {error:#}", case.id));
        let actual_value = serde_json::to_value(actual).expect("serialize S2 analysis result");
        let expected_value = &baseline["result"][&case.id];
        if &actual_value != expected_value {
            let mut differences = Vec::new();
            collect_json_differences(&actual_value, expected_value, "$", &mut differences);
            panic!(
                "S2 in-memory analysis differs from Python baseline for {}:\n{}",
                case.id,
                differences.join("\n")
            );
        }
        results.insert(case.id.clone(), actual_value);
        case_reports.push(serde_json::json!({"case": case.id, "status": "passed"}));
    }

    if let Some(path) = std::env::var_os("AUDIT_S2_ANALYSIS_PARITY_REPORT") {
        let report = serde_json::json!({
            "status": "passed",
            "case_count": case_reports.len(),
            "python_baseline": baseline["source"],
            "result": Value::Object(results),
            "cases": case_reports,
            "scope": "S2 in-memory analysis orchestration only; configuration checking/loading, workbook imports, progress callbacks, XLSX export, real templates, target Linux packaging, and production routing are not covered."
        });
        fs::write(
            path,
            format!("{}\n", serde_json::to_string_pretty(&report).unwrap()),
        )
        .expect("write S2 analysis parity report");
    }
}
