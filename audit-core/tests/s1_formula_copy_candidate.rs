//! Manual, direct-backend candidate probe for the first S1 formula-copy scope.
//!
//! This deliberately bypasses ExcelModifier only to measure the Umya ImportRange
//! candidate. The capability is Unverified and this test must never be treated
//! as a production route or as feature acceptance.

use audit_core::s1_formula_copy::build_s1_formula_copy_edit_plan;
use excel_editor::{
    BackendCapability, BackendErrorKind, CapabilityStatus, EditBackend, EditOperation,
    EditOperationKind, EngineId, OpenpyxlWorkerConfig, OpenpyxlXlsxBackend, OpenpyxlXlsxValidator,
    OutputValidator, UmyaXlsxBackend, WorkbookFeature, WorkbookFormat,
};
use std::{
    collections::BTreeSet,
    fs,
    path::{Path, PathBuf},
    time::Instant,
};

fn required_path(variable: &str) -> PathBuf {
    std::env::var_os(variable)
        .map(PathBuf::from)
        .unwrap_or_else(|| panic!("required candidate input {variable} is missing"))
}

fn openpyxl_worker_path() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .expect("audit-core has a workspace parent")
        .join("crates/excel-editor/python/openpyxl_worker.py")
}

fn readback_failure_category(message: &str) -> &'static str {
    let message = message.to_ascii_lowercase();
    if message.contains("rust unmodified-style readback mismatch") {
        "unmodified_style_semantic_mismatch"
    } else if message.contains("rust style readback mismatch") {
        "imported_style_semantic_mismatch"
    } else if message.contains("unmodified cell value changed") {
        "unmodified_cell_value_changed"
    } else if message.contains("unmodified formula changed") {
        "unmodified_formula_changed"
    } else if message.contains("comment") {
        "comment_preservation_mismatch"
    } else if message.contains("style") {
        "style_preservation_mismatch"
    } else if message.contains("merge") {
        "merged_range_preservation_mismatch"
    } else if message.contains("calculation") {
        "calculation_properties_mismatch"
    } else if message.contains("relationship") || message.contains("defined name") {
        "ooxml_metadata_preservation_mismatch"
    } else {
        "other_readback_validation_failure"
    }
}

fn readback_failure_location<'a>(message: &'a str, marker: &str) -> Option<(&'a str, &'a str)> {
    let failure = message.rsplit_once(marker)?.1;
    let location = failure.split(": expected ").next()?;
    location.rsplit_once('!')
}

fn rust_cell_value_diagnostic(value: Option<&serde_json::Value>) -> serde_json::Value {
    let Some(value) = value else {
        return serde_json::json!({
            "present": false,
            "json_type": "missing",
            "value_class": "missing",
        });
    };
    let json_type = value
        .get("type")
        .and_then(serde_json::Value::as_str)
        .unwrap_or("unknown");
    let raw = value.get("value");
    let string_value = raw.and_then(serde_json::Value::as_str);
    let empty_string = json_type == "string" && string_value == Some("");
    let whitespace_only =
        string_value.is_some_and(|value| !value.is_empty() && value.trim().is_empty());
    let value_class = match json_type {
        "string" if empty_string => "empty_string",
        "string" => "nonempty_string",
        "number" | "datetime_serial" => "number",
        "boolean" => "boolean",
        "error" => "error",
        "datetime_iso" | "duration_iso" => "textual_datetime",
        _ => "other",
    };
    serde_json::json!({
        "present": true,
        "json_type": json_type,
        "value_class": value_class,
        "empty_string": empty_string,
        "string_length_chars": string_value.map(|value| value.chars().count()),
        "whitespace_only": whitespace_only,
    })
}

fn safe_readback_diagnostic(
    message: &str,
    input: &Path,
    output: &Path,
) -> Option<serde_json::Value> {
    if let Some((sheet_name, cell)) =
        readback_failure_location(message, "unmodified cell value changed at ")
    {
        let Ok(source) = excel_reader::read_workbook(input) else {
            return Some(serde_json::json!({
                "kind": "unmodified_cell_value_changed",
                "diagnostic_status": "source_reader_failed",
            }));
        };
        let Ok(output) = excel_reader::read_workbook(output) else {
            return Some(serde_json::json!({
                "kind": "unmodified_cell_value_changed",
                "diagnostic_status": "output_reader_failed",
            }));
        };
        let Some(sheet_ordinal) = source
            .sheets
            .iter()
            .position(|sheet| sheet.name == sheet_name)
            .map(|index| index + 1)
        else {
            return Some(serde_json::json!({
                "kind": "unmodified_cell_value_changed",
                "diagnostic_status": "source_sheet_not_found",
            }));
        };
        let normalized_cell = cell.trim().replace('$', "").to_ascii_uppercase();
        let source_value = source
            .sheets
            .get(sheet_ordinal - 1)?
            .values
            .get(&normalized_cell);
        let output_value = output
            .sheets
            .get(sheet_ordinal - 1)?
            .values
            .get(&normalized_cell);
        return Some(serde_json::json!({
            "kind": "unmodified_cell_value_changed",
            "diagnostic_status": "readback_values_available",
            "sheet_ordinal": sheet_ordinal,
            "cell": normalized_cell,
            "rust_source_json": rust_cell_value_diagnostic(source_value),
            "rust_output_json": rust_cell_value_diagnostic(output_value),
        }));
    }

    if let Some((sheet_name, cell)) =
        readback_failure_location(message, "Rust style readback mismatch at ")
    {
        let Ok(source) = excel_reader::read_workbook(input) else {
            return Some(serde_json::json!({
                "kind": "imported_style_semantic_mismatch",
                "diagnostic_status": "source_reader_failed",
            }));
        };
        let Some(sheet_ordinal) = source
            .sheets
            .iter()
            .position(|sheet| sheet.name == sheet_name)
            .map(|index| index + 1)
        else {
            return Some(serde_json::json!({
                "kind": "imported_style_semantic_mismatch",
                "diagnostic_status": "source_sheet_not_found",
            }));
        };
        let details = message.rsplit_once("Rust style readback mismatch at ")?.1;
        let style_details = details.split_once(": expected ")?.1;
        let style_pair = style_details.rsplit_once(", got ");
        return Some(serde_json::json!({
            "kind": "imported_style_semantic_mismatch",
            "diagnostic_status": if style_pair.is_some() { "style_signatures_available" } else { "style_signatures_unavailable" },
            "sheet_ordinal": sheet_ordinal,
            "cell": cell.trim().replace('$', "").to_ascii_uppercase(),
            "rust_expected_style": style_pair.map(|(expected, _)| expected.trim()),
            "rust_actual_style": style_pair.map(|(_, actual)| actual.trim()),
        }));
    }
    None
}

#[test]
#[ignore = "manual candidate probe; direct Unverified backend is not a production route"]
fn probes_s1_formula_copy_import_ranges_without_claiming_capability() {
    let target = required_path("AUDIT_S1_COPY_CANDIDATE_INPUT");
    let template = required_path("AUDIT_S1_COPY_CANDIDATE_TEMPLATE");
    let output = required_path("AUDIT_S1_COPY_CANDIDATE_OUTPUT");
    let report_path = required_path("AUDIT_S1_COPY_CANDIDATE_REPORT");

    assert!(target.is_file(), "candidate target workbook is unavailable");
    assert!(
        template.is_file(),
        "candidate template workbook is unavailable"
    );
    assert!(!output.exists(), "candidate output path must be new");
    let mapping_config = std::env::var_os("AUDIT_S1_COPY_CANDIDATE_MAPPING").map(PathBuf::from);
    let prepared =
        build_s1_formula_copy_edit_plan(&target, &template, &output, mapping_config.as_deref())
            .expect("build S1 formula-copy business plan from template rules");
    let template_profile = prepared.template_profile;
    let mut plan = prepared.edit_plan;
    let range_import_count = plan
        .operations
        .iter()
        .filter(|operation| operation.kind() == EditOperationKind::ImportRange)
        .count();
    let set_comment_count = plan
        .operations
        .iter()
        .filter(|operation| operation.kind() == EditOperationKind::SetComment)
        .count();
    let legacy_empty_author_compat_count = plan
        .operations
        .iter()
        .filter(|operation| {
            matches!(
                operation,
                EditOperation::SetComment {
                    author: Some(author),
                    ..
                } if author == "None"
            )
        })
        .count();
    assert_eq!(
        range_import_count, 4,
        "candidate scope requires four ranges"
    );
    let include_calculation_properties = std::env::var("AUDIT_S1_COPY_CANDIDATE_CALC_PROPERTIES")
        .map(|value| value != "0")
        .unwrap_or(true);
    if !include_calculation_properties {
        plan.operations.retain(|operation| {
            !matches!(
                operation,
                EditOperation::SetWorkbookCalculationProperties { .. }
            )
        });
    }
    let workbook_profile = plan.workbook.clone();
    let candidate_engine =
        std::env::var("AUDIT_S1_COPY_CANDIDATE_ENGINE").unwrap_or_else(|_| "umya".to_owned());
    let (engine, supported_features) = match candidate_engine.as_str() {
        "umya" => (EngineId::UmyaSpreadsheet, BTreeSet::new()),
        "openpyxl-worker" => {
            let supported_features = workbook_profile
                .features_present
                .union(&template_profile.features_present)
                .copied()
                .collect::<BTreeSet<WorkbookFeature>>();
            (EngineId::OpenpyxlWorker, supported_features)
        }
        _ => panic!("candidate engine must be umya or openpyxl-worker"),
    };
    let capability = BackendCapability {
        engine,
        status: CapabilityStatus::Unverified,
        format: WorkbookFormat::Xlsx,
        operations: [
            EditOperationKind::ImportRange,
            EditOperationKind::SetComment,
            EditOperationKind::SetWorkbookCalculationProperties,
        ]
        .into_iter()
        .collect(),
        supported_features,
        preserved: BTreeSet::new(),
        platform: "audit-s1-manual-candidate".to_owned(),
        evidence: None,
    };
    // This is intentionally a direct candidate invocation. ExcelModifier must
    // continue rejecting this Unverified capability until real-template and
    // target-platform preservation evidence is reviewed and registered.
    let started = Instant::now();
    let include_safe_diagnostics =
        std::env::var("AUDIT_S1_COPY_SAFE_DIAGNOSTICS").as_deref() == Ok("1");
    let (apply_result, output_readback_failure_category, output_readback_diagnostic) =
        match candidate_engine.as_str() {
            "umya" => {
                let backend =
                    UmyaXlsxBackend::new(capability).expect("construct unverified Umya candidate");
                (backend.apply(&plan), None, None)
            }
            "openpyxl-worker" => {
                let python = required_path("RUST_WORKSPACE_OPENPYXL_PYTHON");
                let config = OpenpyxlWorkerConfig::new(python, openpyxl_worker_path())
                    .expect("construct isolated test worker config");
                let backend = OpenpyxlXlsxBackend::new(capability, config)
                    .expect("construct unverified openpyxl candidate");
                let result = backend.apply(&plan);
                let readback_failure = if result.is_ok() && output.is_file() {
                    OpenpyxlXlsxValidator
                        .validate(&plan, &output)
                        .err()
                        .map(|error| {
                            let message = format!("{error:#}");
                            let category = readback_failure_category(&message);
                            let diagnostic = include_safe_diagnostics
                                .then(|| safe_readback_diagnostic(&message, &target, &output))
                                .flatten();
                            (category, diagnostic)
                        })
                } else {
                    None
                };
                match readback_failure {
                    Some((category, diagnostic)) => (result, Some(category), diagnostic),
                    None => (result, None, None),
                }
            }
            _ => unreachable!("candidate engine validated above"),
        };
    let output_created = Path::new(&output).is_file();
    let apply_failure_kind = apply_result
        .as_ref()
        .err()
        .map(|failure| match failure.kind {
            BackendErrorKind::UnsupportedBeforeWrite => "unsupported_before_write",
            BackendErrorKind::Failed => "backend_failed",
        });
    let apply_failure_category = apply_result.as_ref().err().map(|failure| {
        let message = failure.message.to_ascii_lowercase();
        if message.contains("unverified workbook structure") {
            "unverified_target_workbook_structure"
        } else if message.contains("unverified range import source structure") {
            "unverified_template_workbook_structure"
        } else if message.contains("does not support the complete operation batch") {
            "complete_operation_batch_unsupported"
        } else if message.contains("ImportRange source contains workbook features") {
            "unverified_template_workbook_features"
        } else if message.contains("dimension-copy flags are not enabled") {
            "unsupported_range_dimension_copy"
        } else if message.contains("formula_cache_not_preserved")
            || message.contains("cached formula result")
        {
            "formula_cache_not_preserved"
        } else if message.contains("unknown_package_parts")
            || message.contains("inventory is incomplete")
        {
            "unverified_target_workbook_structure"
        } else if message.contains("unsupported_import_source") {
            "unsupported_template_workbook_structure"
        } else if message.contains("does not exist") {
            "missing_candidate_input"
        } else if message.contains("supports xlsx only") {
            "unsupported_input_format"
        } else if message.contains("open xlsx") {
            "target_open_failed"
        } else if message.contains("apply import_range") {
            "range_import_failed"
        } else if message.contains("save xlsx") {
            "candidate_save_failed"
        } else {
            "other_candidate_failure"
        }
    });
    let status = if apply_result.is_ok() && output_created {
        "written_unverified_candidate"
    } else {
        "candidate_backend_rejected_or_failed"
    };
    let candidate_report = serde_json::json!({
        "status": status,
        "range_import_count": range_import_count,
        "set_comment_count": set_comment_count,
        "legacy_empty_author_compat_count": legacy_empty_author_compat_count,
        "operation_count": plan.operations.len(),
        "range_resolution": "rust_s1_template_rule_plan",
        "engine": match engine {
            EngineId::OoxmlPatch => "ooxml-patch",
            EngineId::UmyaSpreadsheet => "umya-spreadsheet",
            EngineId::ExcelizeWorker => "excelize-worker",
            EngineId::OpenpyxlWorker => "openpyxl-worker",
        },
        "capability_status": "unverified",
        "production_route_eligible": false,
        "calculation_properties_operation_present": include_calculation_properties,
        "formula_recalculation_performed": false,
        "target_inventory_complete": workbook_profile.inventory_complete,
        "output_created": output_created,
        "output_readback_failed": output_readback_failure_category.is_some(),
        "output_readback_failure_category": output_readback_failure_category,
        "output_readback_diagnostic": output_readback_diagnostic,
        "apply_failure_kind": apply_failure_kind,
        "apply_failure_category": apply_failure_category,
        "elapsed_ms": started.elapsed().as_millis(),
    });
    fs::write(
        report_path,
        serde_json::to_vec_pretty(&candidate_report).expect("serialize sanitized candidate report"),
    )
    .expect("write sanitized candidate report");
    println!("AUDIT_S1_COPY_CANDIDATE={candidate_report}");
}
