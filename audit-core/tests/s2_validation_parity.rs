use audit_core::{
    s2_import::{import_period_directory, IndicatorRecord, S2ImportConfig, S2SourceRef},
    s2_validation::{validate_config_warnings, validate_external_warnings, validate_import},
};
use serde_json::Value;
use std::{fs, path::Path};

fn workspace_path(relative: &str) -> std::path::PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join(relative)
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
            if let Some(Value::String(trace)) = fields.get_mut("calculation_trace") {
                *trace = normalize_trace_paths(trace);
            }
            for field in fields.values_mut() {
                normalize_source_paths(field);
            }
        }
        _ => {}
    }
}

fn normalize_trace_paths(trace: &str) -> String {
    if trace.is_empty() {
        return String::new();
    }
    trace
        .split('；')
        .map(|segment| {
            let (label, detail) = segment.split_once('：').expect("trace label");
            let (location, row) = detail.split_once("!第").expect("trace row");
            let (source_path, sheet) = location.rsplit_once('/').expect("trace sheet");
            let basename = source_path
                .rsplit(['/', '\\'])
                .next()
                .expect("source basename");
            format!("{label}：{basename}/{sheet}!第{row}")
        })
        .collect::<Vec<_>>()
        .join("；")
}

#[test]
fn s2_quality_findings_match_read_only_python_baseline() {
    let fixture_dir = workspace_path("tests/fixtures/s2_import");
    let config: S2ImportConfig = serde_json::from_slice(
        &fs::read(fixture_dir.join("config.json")).expect("read synthetic S2 config"),
    )
    .expect("decode S2 config");
    let mut import_result =
        import_period_directory(&fixture_dir, &config, "本期").expect("import S2 fixture");

    let mut missing_identity: IndicatorRecord = import_result
        .records
        .last()
        .expect("fixture has a record")
        .clone();
    missing_identity.institution_id.clear();
    missing_identity.form_code.clear();
    missing_identity.indicator_code = "MISSING-IDENTITY".to_owned();
    missing_identity.source = S2SourceRef {
        source_file: "synthetic-missing.xlsx".to_owned(),
        source_sheet: "Synthetic".to_owned(),
        source_row: 99,
    };
    import_result.records.push(missing_identity);

    let config_warnings = vec!["合成配置缺少字段".to_owned(), "合成配置存在警告".to_owned()];
    let external_warnings = vec!["外部数据缺失".to_owned(), "外部值不是数值".to_owned()];
    let mut actual = serde_json::json!({
        "import_findings": validate_import(&import_result),
        "config_findings": validate_config_warnings(&config_warnings),
        "external_findings": validate_external_warnings(&external_warnings),
    });
    normalize_source_paths(&mut actual);

    let baseline: Value = serde_json::from_slice(
        &fs::read(workspace_path(
            "probe/reports/s2_validation_python_baseline.json",
        ))
        .expect("read Python validation baseline"),
    )
    .expect("decode Python validation baseline");
    assert_eq!(actual, baseline["result"]);
}
