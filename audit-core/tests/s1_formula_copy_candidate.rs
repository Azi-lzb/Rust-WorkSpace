//! Manual, direct-backend candidate probe for the first S1 formula-copy scope.
//!
//! This deliberately bypasses ExcelModifier only to measure the Umya ImportRange
//! candidate. The capability is Unverified and this test must never be treated
//! as a production route or as feature acceptance.

use audit_core::s1_template_rules::load_s1_template_rule_plan;
use excel_editor::{
    inspect_xlsx_profile, BackendCapability, BackendErrorKind, CapabilityStatus, EditBackend,
    EditOperation, EditOperationKind, EngineId, ExcelEditPlan, OpenpyxlWorkerConfig,
    OpenpyxlXlsxBackend, OpenpyxlXlsxValidator, OutputValidator, PreservationRequirement,
    RangeStyleMode, UmyaXlsxBackend, WorkbookFeature, WorkbookFormat,
};
use std::{
    collections::{BTreeMap, BTreeSet},
    fs,
    path::{Path, PathBuf},
    time::Instant,
};

fn required_path(variable: &str) -> PathBuf {
    std::env::var_os(variable)
        .map(PathBuf::from)
        .unwrap_or_else(|| panic!("required candidate input {variable} is missing"))
}

fn target_start(address: &str) -> Option<String> {
    let address = address.rsplit('!').next()?.trim();
    let first_cell = address.split(':').next()?.replace('$', "");
    (!first_cell.is_empty()).then_some(first_cell)
}

fn openpyxl_worker_path() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .expect("audit-core has a workspace parent")
        .join("crates/excel-editor/python/openpyxl_worker.py")
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
    let template_plan = load_s1_template_rule_plan(&template, mapping_config.as_deref())
        .expect("resolve formula-copy ranges through Rust S1 template rules");
    let ranges = template_plan.definition.copy_ranges;
    assert_eq!(ranges.len(), 4, "candidate scope requires four ranges");
    let range_import_count = ranges.len();

    let mut operations: Vec<EditOperation> = ranges
        .into_iter()
        .map(|range| {
            let start = target_start(&range.address)
                .expect("copy range must begin with a valid cell address");
            EditOperation::ImportRange {
                source_workbook: template.clone(),
                source_sheet: range.sheet_name.clone(),
                source_range: range.address,
                target_sheet: range.sheet_name,
                target_start: start,
                copy_values: true,
                copy_formulas: true,
                style_mode: RangeStyleMode::Full,
                copy_comments: true,
                copy_merges: false,
                copy_row_heights: false,
                copy_column_widths: false,
                copy_row_hidden: false,
                copy_column_hidden: false,
                copy_outline_levels: false,
                copy_all_row_dimensions: false,
                copy_all_column_dimensions: false,
                materialize_blank_cells: true,
                hide_zero: false,
            }
        })
        .collect();
    let include_calculation_properties = std::env::var("AUDIT_S1_COPY_CANDIDATE_CALC_PROPERTIES")
        .map(|value| value != "0")
        .unwrap_or(true);
    if include_calculation_properties {
        operations.push(EditOperation::SetWorkbookCalculationProperties {
            full_calc_on_load: Some(true),
            force_full_calc: Some(true),
        });
    }

    let workbook_profile = inspect_xlsx_profile(&target).expect("inspect candidate target profile");
    let candidate_engine =
        std::env::var("AUDIT_S1_COPY_CANDIDATE_ENGINE").unwrap_or_else(|_| "umya".to_owned());
    let (engine, supported_features) = match candidate_engine.as_str() {
        "umya" => (EngineId::UmyaSpreadsheet, BTreeSet::new()),
        "openpyxl-worker" => {
            let source_profile =
                inspect_xlsx_profile(&template).expect("inspect formula template profile");
            let supported_features = workbook_profile
                .features_present
                .union(&source_profile.features_present)
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
            EditOperationKind::SetWorkbookCalculationProperties,
        ]
        .into_iter()
        .collect(),
        supported_features,
        preserved: BTreeSet::new(),
        platform: "audit-s1-manual-candidate".to_owned(),
        evidence: None,
    };
    let plan = ExcelEditPlan {
        request_id: "audit-s1-formula-copy-candidate".to_owned(),
        input: target.clone(),
        output: output.clone(),
        format: WorkbookFormat::Xlsx,
        workbook: workbook_profile.clone(),
        operations,
        preserve: [
            PreservationRequirement::CellValuesAndTypes,
            PreservationRequirement::FormulaText,
            PreservationRequirement::CommentsAndPositions,
            PreservationRequirement::Styles,
            PreservationRequirement::MergedRanges,
            PreservationRequirement::ConditionalFormatting,
            PreservationRequirement::DataValidations,
            PreservationRequirement::DefinedNames,
            PreservationRequirement::RowAndColumnDimensions,
            PreservationRequirement::PageSettings,
            PreservationRequirement::HeaderFooterText,
            PreservationRequirement::Relationships,
            PreservationRequirement::UnmodifiedPackageParts,
            PreservationRequirement::SheetIdentityAndOrder,
        ]
        .into_iter()
        .collect(),
        drop_allowed: BTreeMap::new(),
    };

    // This is intentionally a direct candidate invocation. ExcelModifier must
    // continue rejecting this Unverified capability until real-template and
    // target-platform preservation evidence is reviewed and registered.
    let started = Instant::now();
    let (apply_result, output_readback_failure_category) = match candidate_engine.as_str() {
        "umya" => {
            let backend =
                UmyaXlsxBackend::new(capability).expect("construct unverified Umya candidate");
            (backend.apply(&plan), None)
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
                    .map(|_| "rust_readback_validation_failed")
            } else {
                None
            };
            (result, readback_failure)
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
        "operation_count": range_import_count + if include_calculation_properties { 1 } else { 0 },
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
