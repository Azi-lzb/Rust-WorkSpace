use audit_core::s1_template::{recommend_from_sheet_sets, TemplateProfile};
use serde_json::{json, Value};
use std::{fs, path::PathBuf};

#[test]
fn s1_template_recommendation_matches_read_only_python_baseline() {
    let cases: Value =
        serde_json::from_str(include_str!("fixtures/template_recommendation_cases.json"))
            .expect("valid S1 template recommendation cases");
    let baseline: Value = serde_json::from_str(include_str!(
        "../probe/reports/template_recommendation_python_baseline.json"
    ))
    .expect("valid Python template recommendation baseline");

    let cases = cases["cases"].as_array().unwrap();
    let expected = baseline["cases"].as_array().unwrap();
    assert_eq!(cases.len(), expected.len());

    for (case, expected) in cases.iter().zip(expected) {
        assert_eq!(case["id"], expected["id"]);
        let templates = case["templates"]
            .as_array()
            .unwrap()
            .iter()
            .map(|item| TemplateProfile {
                path: PathBuf::from(item["name"].as_str().unwrap()),
                sheet_names: item["sheets"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .map(|sheet| sheet.as_str().unwrap().to_owned())
                    .collect(),
            })
            .collect::<Vec<_>>();
        let sources = case["sources"]
            .as_array()
            .unwrap()
            .iter()
            .map(|sheets| {
                sheets
                    .as_array()
                    .unwrap()
                    .iter()
                    .map(|sheet| sheet.as_str().unwrap().to_owned())
                    .collect::<Vec<_>>()
            })
            .collect::<Vec<_>>();
        let suggestion = recommend_from_sheet_sets(&templates, &sources);
        let actual = json!({
            "template": suggestion.template_path.as_ref().and_then(|path| {
                path.file_name().map(|name| name.to_string_lossy().into_owned())
            }),
            "confidence": suggestion.confidence,
            "matched": suggestion.matched,
            "details": suggestion.details,
            "alternatives": suggestion.alternatives,
        });
        assert_eq!(actual, expected["result"], "case {}", case["id"]);
    }

    let report = json!({
        "schema_version": 1,
        "suite": "audit.s1_template_recommendation",
        "source_project": "Open-Data-Audit",
        "source_functions": ["base_audit.discovery.recommend_template"],
        "rust_functions": [
            "audit_core::s1_template::recommend_template",
            "audit_core::s1_template::recommend_template_for_workbooks",
            "audit_core::s1_template::recommend_from_sheet_sets"
        ],
        "source_baseline": "probe/reports/template_recommendation_python_baseline.json",
        "status": "保留（模板推荐局部）",
        "parity": {"cases": cases.len(), "passed": cases.len(), "failed": 0},
        "items": [
            {"item": "sheet-set score and source-sheet subset contract", "status": "保留"},
            {"item": "mixed report types stop automatic selection", "status": "保留"},
            {"item": "near-tied candidates require manual selection", "status": "保留"},
            {
                "item": "directory audit recommends from its discovered and selected source workbooks",
                "status": "合成 CLI 集成通过；歧义时停止自动审核"
            },
            {
                "item": "config-real plus 2026-07-31 source-tree smoke (68 source workbooks, 8 templates)",
                "status": "保留"
            },
            {"item": "other real template sets and reporting periods", "status": "未验证"},
            {"item": "complete S1-F01 audit workflow", "status": "未验证"},
            {"item": "performance and target Linux", "status": "未测量"}
        ],
        "verification": {
            "rust": "compares all synthetic fixture cases with the read-only Python baseline; directory CLI consumes the same selected source list",
            "real_smoke": "probe/reports/template_recommendation_real_smoke.json"
        }
    });
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("probe/reports/template_recommendation_parity.json");
    fs::write(path, serde_json::to_vec_pretty(&report).unwrap()).unwrap();
}
