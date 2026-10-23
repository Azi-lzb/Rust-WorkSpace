use audit_core::{s2_exporter::export_s2_report, s2_validation::S2AuditFinding};
use serde::Deserialize;
use serde_json::Value;
use std::{
    fs,
    path::PathBuf,
    time::{SystemTime, UNIX_EPOCH},
};

#[derive(Deserialize)]
struct FixtureRoot {
    findings: Vec<S2AuditFinding>,
    hide_unchanged_period_rows: bool,
    hide_matching_external_rows: bool,
}

#[test]
fn s2_report_writer_emits_three_valid_sheets_and_keeps_formula_policy() {
    let fixture: FixtureRoot =
        serde_json::from_str(include_str!("fixtures/s2_exporter/cases.json")).unwrap();
    let probe_root = std::env::var_os("AUDIT_S2_EXPORT_PROBE_DIR").map(PathBuf::from);
    let test_dir = unique_test_dir(probe_root.as_deref());
    fs::create_dir_all(&test_dir).unwrap();

    let mut events = Vec::new();
    let mut on_progress = |sheet: &str, current: usize, total: usize| {
        events.push((sheet.to_owned(), current, total));
    };
    let result_path = export_s2_report(
        &test_dir,
        &fixture.findings,
        fixture.hide_unchanged_period_rows,
        fixture.hide_matching_external_rows,
        Some(&mut on_progress),
    )
    .unwrap();

    assert!(result_path
        .file_name()
        .unwrap()
        .to_string_lossy()
        .starts_with("报表审核清单_"));
    let workbook = excel_reader::read_workbook(&result_path).unwrap();
    assert_eq!(
        workbook
            .sheets
            .iter()
            .map(|sheet| sheet.name.as_str())
            .collect::<Vec<_>>(),
        ["环比规则", "外部核对规则", "校验规则"]
    );
    assert!(workbook.sheets[2].formulas.contains_key("J2"));
    assert!(!workbook.sheets[1].formulas.contains_key("F2"));
    let report_item = workbook.sheets[1].values.get("F2").unwrap();
    assert_eq!(
        report_item.get("type").and_then(Value::as_str),
        Some("string")
    );
    assert_eq!(
        report_item.get("value").and_then(Value::as_str),
        Some("=SUM(1,2)")
    );
    let current_zero = workbook.sheets[0].values.get("L2").unwrap();
    assert_eq!(current_zero.get("value").and_then(Value::as_f64), Some(0.0));
    assert_eq!(events.last().unwrap(), &("save".to_owned(), 5, 5));

    if probe_root.is_none() {
        fs::remove_dir_all(test_dir).unwrap();
    }
}

fn unique_test_dir(base: Option<&std::path::Path>) -> PathBuf {
    let stamp = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap()
        .as_nanos();
    let name = format!("s2-exporter-{}-{stamp}", std::process::id());
    base.map(|base| base.join(&name))
        .unwrap_or_else(|| std::env::temp_dir().join(name))
}
