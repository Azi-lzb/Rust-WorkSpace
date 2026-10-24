//! Manual, direct-backend candidate probe for the first S1 formula-copy scope.
//!
//! This deliberately bypasses ExcelModifier only to measure the Umya ImportRange
//! candidate. The capability is Unverified and this test must never be treated
//! as a production route or as feature acceptance.

use excel_editor::{
    inspect_xlsx_profile, BackendCapability, CapabilityStatus, EditBackend, EditOperation,
    EditOperationKind, EngineId, ExcelEditPlan, PreservationRequirement, RangeStyleMode,
    UmyaXlsxBackend, WorkbookFormat,
};
use serde::Deserialize;
use std::{
    collections::{BTreeMap, BTreeSet},
    fs,
    path::{Path, PathBuf},
    time::Instant,
};

#[derive(Debug, Deserialize)]
struct CopyRangeSpec {
    sheet_name: String,
    address: String,
}

#[derive(Debug, Deserialize)]
struct CopyRangesFile {
    ranges: Vec<CopyRangeSpec>,
}

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

#[test]
#[ignore = "manual candidate probe; direct Unverified backend is not a production route"]
fn probes_s1_formula_copy_import_ranges_without_claiming_capability() {
    let target = required_path("AUDIT_S1_COPY_CANDIDATE_INPUT");
    let template = required_path("AUDIT_S1_COPY_CANDIDATE_TEMPLATE");
    let output = required_path("AUDIT_S1_COPY_CANDIDATE_OUTPUT");
    let ranges_path = required_path("AUDIT_S1_COPY_CANDIDATE_RANGES");
    let report_path = required_path("AUDIT_S1_COPY_CANDIDATE_REPORT");

    assert!(target.is_file(), "candidate target workbook is unavailable");
    assert!(
        template.is_file(),
        "candidate template workbook is unavailable"
    );
    assert!(!output.exists(), "candidate output path must be new");
    let ranges: CopyRangesFile =
        serde_json::from_slice(&fs::read(ranges_path).expect("read range definitions"))
            .expect("parse range definitions");
    assert_eq!(
        ranges.ranges.len(),
        4,
        "candidate scope requires four ranges"
    );

    let mut operations: Vec<EditOperation> = ranges
        .ranges
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
    operations.push(EditOperation::SetWorkbookCalculationProperties {
        full_calc_on_load: Some(true),
        force_full_calc: Some(true),
    });

    let capability = BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Unverified,
        format: WorkbookFormat::Xlsx,
        operations: [
            EditOperationKind::ImportRange,
            EditOperationKind::SetWorkbookCalculationProperties,
        ]
        .into_iter()
        .collect(),
        supported_features: BTreeSet::new(),
        preserved: BTreeSet::new(),
        platform: "audit-s1-manual-candidate".to_owned(),
        evidence: None,
    };
    let backend = UmyaXlsxBackend::new(capability).expect("construct unverified candidate backend");
    let workbook_profile = inspect_xlsx_profile(&target).expect("inspect candidate target profile");
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
    let apply_result = backend.apply(&plan);
    let output_created = Path::new(&output).is_file();
    let status = if apply_result.is_ok() && output_created {
        "written_unverified_candidate"
    } else {
        "candidate_backend_rejected_or_failed"
    };
    let candidate_report = serde_json::json!({
        "status": status,
        "range_import_count": 4,
        "operation_count": 5,
        "engine": "umya-spreadsheet",
        "capability_status": "unverified",
        "production_route_eligible": false,
        "calculation_properties_operation_present": true,
        "formula_recalculation_performed": false,
        "target_inventory_complete": workbook_profile.inventory_complete,
        "output_created": output_created,
        "elapsed_ms": started.elapsed().as_millis(),
    });
    fs::write(
        report_path,
        serde_json::to_vec_pretty(&candidate_report).expect("serialize sanitized candidate report"),
    )
    .expect("write sanitized candidate report");
    println!("AUDIT_S1_COPY_CANDIDATE={candidate_report}");
}
