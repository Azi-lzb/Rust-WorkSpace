//! Manual, direct-backend candidate probe for the first S1 formula-copy scope.
//!
//! This deliberately probes an Unverified range-import candidate directly,
//! without routing it through ExcelModifier. It must never be treated as a
//! production route or as feature acceptance.

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

fn candidate_apply_failure_category(message: &str) -> &'static str {
    let message = message.to_ascii_lowercase();
    if message.contains("unverified workbook structure")
        || message.contains("source workbook contains unverified ooxml parts")
        || message.contains("unknown_package_parts")
        || message.contains("inventory is incomplete")
    {
        "unverified_target_workbook_structure"
    } else if message.contains("unverified range import source structure") {
        "unverified_template_workbook_structure"
    } else if message
        .contains("exceleditplan workbook profile differs from the source package inventory")
    {
        "target_workbook_profile_mismatch"
    } else if message.contains("cannot preserve stored content in merged non-anchor cell") {
        "target_merged_non_anchor_payload"
    } else if message
        .contains("cannot preserve importrange source content in merged non-anchor cell")
    {
        "source_merged_non_anchor_payload"
    } else if message.contains("does not support the complete operation batch")
        || message.contains("does not support the complete batch")
        || message.contains("unsupported_operation")
    {
        "complete_operation_batch_unsupported"
    } else if message.contains("importrange source contains workbook features") {
        "unverified_template_workbook_features"
    } else if message.contains("dimension-copy flags are not enabled") {
        "unsupported_range_dimension_copy"
    } else if message.contains("formula_cache_not_preserved")
        || message.contains("cached formula result")
    {
        "formula_cache_not_preserved"
    } else if message.contains("unsupported_import_source") {
        "unsupported_template_workbook_structure"
    } else if message.contains("does not exist") {
        "missing_candidate_input"
    } else if message.contains("supports xlsx only")
        || message.contains("accepts xlsx only")
        || message.contains("unsupported_format")
    {
        "unsupported_input_format"
    } else if message.contains("open xlsx") {
        "target_open_failed"
    } else if message.contains("apply import_range") {
        "range_import_failed"
    } else if message.contains("save xlsx") {
        "candidate_save_failed"
    } else if message.contains("bundled python/openpyxl worker")
        || message.contains("openpyxl worker is unavailable")
        || message.contains("start bundled openpyxl worker")
    {
        "worker_unavailable"
    } else {
        "other_candidate_failure"
    }
}

#[test]
fn candidate_apply_failure_categories_are_path_free_and_case_insensitive() {
    let cases = [
        (
            "openpyxl cannot preserve stored content in merged non-anchor cell Sheet1!D5",
            "target_merged_non_anchor_payload",
        ),
        (
            "openpyxl cannot preserve ImportRange source content in merged non-anchor cell Sheet1!A10",
            "source_merged_non_anchor_payload",
        ),
        (
            "ImportRange source contains workbook features outside the openpyxl capability evidence",
            "unverified_template_workbook_features",
        ),
        (
            "ExcelEditPlan workbook profile differs from the source package inventory",
            "target_workbook_profile_mismatch",
        ),
        (
            "formula_cache_not_preserved: cached formula result would be discarded",
            "formula_cache_not_preserved",
        ),
    ];

    for (message, expected) in cases {
        assert_eq!(candidate_apply_failure_category(message), expected);
    }
}

fn readback_failure_location<'a>(message: &'a str, marker: &str) -> Option<(&'a str, &'a str)> {
    let failure = message.rsplit_once(marker)?.1;
    let location = failure.split(": expected ").next()?;
    location.rsplit_once('!')
}

fn readback_failure_location_token<'a>(
    message: &'a str,
    marker: &str,
) -> Option<(&'a str, &'a str)> {
    let failure = message.rsplit_once(marker)?.1;
    let location = failure.split_whitespace().next()?;
    let (sheet, cell) = location.rsplit_once('!')?;
    let cell = cell.trim_end_matches(|character: char| matches!(character, ':' | ',' | ';'));
    Some((sheet, cell))
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
    template: &Path,
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

    if let Some((sheet_name, cell)) = readback_failure_location_token(
        message,
        "cannot determine inherited row/column style for absent ImportRange source cell ",
    ) {
        let template_sheet_ordinal =
            excel_reader::read_workbook(template)
                .ok()
                .and_then(|workbook| {
                    workbook
                        .sheets
                        .iter()
                        .position(|sheet| sheet.name == sheet_name)
                        .map(|index| index + 1)
                });
        return Some(serde_json::json!({
            "kind": "implicit_import_source_style_unresolved",
            "diagnostic_status": "failure_location_available",
            "phase": "resolve_implicit_import_source_style",
            "template_sheet_ordinal": template_sheet_ordinal,
            "cell": cell.trim().replace('$', "").to_ascii_uppercase(),
        }));
    }

    let style_failure = [
        (
            "Rust unmodified-style readback mismatch at ",
            "unmodified_style_semantic_mismatch",
            "unmodified_workbook_style_readback",
        ),
        (
            "Rust unmodified merged-cell border readback mismatch at ",
            "unmodified_merged_border_mismatch",
            "unmodified_merged_border_readback",
        ),
        (
            "Rust CopyRange style readback mismatch at ",
            "copied_style_semantic_mismatch",
            "copy_range_style_readback",
        ),
    ]
    .into_iter()
    .find_map(|(marker, kind, phase)| {
        readback_failure_location_token(message, marker).map(|location| (location, kind, phase))
    });
    if let Some(((sheet_name, cell), kind, phase)) = style_failure {
        let sheet_ordinal = excel_reader::read_workbook(input)
            .ok()
            .and_then(|workbook| {
                workbook
                    .sheets
                    .iter()
                    .position(|sheet| sheet.name == sheet_name)
                    .map(|index| index + 1)
            });
        return Some(serde_json::json!({
            "kind": kind,
            "diagnostic_status": "failure_location_available",
            "phase": phase,
            "sheet_ordinal": sheet_ordinal,
            "cell": cell.trim().replace('$', "").to_ascii_uppercase(),
        }));
    }

    let lower_message = message.to_ascii_lowercase();
    let phase = if lower_message.contains("open source workbook for style readback") {
        "open_source_for_style_readback"
    } else if lower_message.contains("open openpyxl output with umya for independent rust readback")
    {
        "open_candidate_for_independent_style_readback"
    } else if lower_message.contains("style-preservation output sheet") {
        "find_output_sheet_for_style_preservation"
    } else if lower_message.contains("cannot determine inherited row/column style") {
        "resolve_implicit_import_source_style"
    } else if lower_message.contains("open importrange styles from") {
        "open_import_range_source_styles"
    } else if lower_message.contains("style") {
        "other_style_preservation_validation"
    } else {
        return None;
    };
    Some(serde_json::json!({
        "kind": "style_preservation_mismatch",
        "diagnostic_status": "safe_failure_phase_available",
        "phase": phase,
        "failure_category": readback_failure_category(message),
    }))
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
                                .then(|| {
                                    safe_readback_diagnostic(&message, &target, &template, &output)
                                })
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
    let apply_failure_category = apply_result
        .as_ref()
        .err()
        .map(|failure| candidate_apply_failure_category(&failure.message));
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
