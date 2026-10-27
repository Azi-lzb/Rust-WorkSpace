use anyhow::{Context, Result};
use excel_editor::{
    inspect_xlsx_profile, BackendCapability, BackendErrorKind, CapabilityStatus, EditBackend,
    EditOperation, EditOperationKind, EngineId, ExcelEditPlan, ExcelModifier, OutputValidator,
    PreservationRequirement, SheetVisibility, UmyaXlsxBackend, UmyaXlsxValidator, WorkbookFormat,
};
use std::{
    fs,
    path::PathBuf,
    sync::Arc,
    time::{SystemTime, UNIX_EPOCH},
};

struct TestDir(PathBuf);

impl TestDir {
    fn new() -> Self {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("system time is after epoch")
            .as_nanos();
        let path = std::env::temp_dir().join(format!("excel-editor-s3-f3-{nonce}"));
        fs::create_dir_all(&path).expect("create test directory");
        Self(path)
    }

    fn path(&self, name: &str) -> PathBuf {
        self.0.join(name)
    }
}

impl Drop for TestDir {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

#[test]
fn excel_modifier_applies_s3_f3_structural_and_format_operations() -> Result<()> {
    let dir = TestDir::new();
    let input = dir.path("input.xlsx");
    let output = dir.path("output.xlsx");

    let mut workbook = umya_spreadsheet::new_file();
    let target = workbook.sheet_by_name_mut("Sheet1")?;
    target.set_name("表单");
    target.cell_mut("A1").set_value_string("保留 A");
    target.cell_mut("B1").set_value_string("删除 B");
    target.cell_mut("C1").set_value_number(12.5);
    target.cell_mut("D1").set_value_string("保留 D");
    target.cell_mut("E1").set_formula("SUM(C1:D1)+$C$1");
    target.cell_mut("B2").set_formula("B1");
    target.add_merge_cells("C3:D3");
    target.row_dimension_mut(4).set_hidden(false);
    let mut comment = umya_spreadsheet::Comment::default();
    comment.new_comment("D2");
    comment.set_author("test");
    comment.set_text_string("column shift");
    target.add_comments(comment);

    let other = workbook
        .new_sheet("引用")
        .context("create reference sheet")?;
    other.cell_mut("A1").set_formula("'表单'!$C$1+C1");
    other
        .cell_mut("B1")
        .set_value_string("other sheet unchanged");
    workbook.new_sheet("过期配置")?;
    umya_spreadsheet::writer::xlsx::write(&workbook, &input)?;
    assert!(excel_reader::read_defined_names(&input)?.is_empty());
    let source_bytes = fs::read(&input)?;

    let profile = inspect_xlsx_profile(&input)?;
    let platform = "synthetic-s3-f3-edit-test";
    let plan = ExcelEditPlan {
        request_id: "s3-f3-edit-test".to_owned(),
        input: input.clone(),
        output: output.clone(),
        format: WorkbookFormat::Xlsx,
        workbook: profile.clone(),
        operations: vec![
            EditOperation::DeleteSheet {
                name: "过期配置".to_owned(),
            },
            EditOperation::DeleteColumn {
                sheet: "表单".to_owned(),
                column_index: 2,
            },
            EditOperation::SetCellNumberFormat {
                sheet: "表单".to_owned(),
                cell: "B1".to_owned(),
                format_code: "0.00%".to_owned(),
            },
            EditOperation::SetRowHidden {
                sheet: "表单".to_owned(),
                row_index: 5,
                hidden: true,
            },
            EditOperation::SetSheetVisibility {
                sheet: "表单".to_owned(),
                visibility: SheetVisibility::Hidden,
            },
        ],
        preserve: [
            PreservationRequirement::CellValuesAndTypes,
            PreservationRequirement::FormulaText,
            PreservationRequirement::CommentsAndPositions,
            PreservationRequirement::MergedRanges,
            PreservationRequirement::DefinedNames,
            PreservationRequirement::RowAndColumnDimensions,
            PreservationRequirement::SheetIdentityAndOrder,
        ]
        .into_iter()
        .collect(),
        drop_allowed: Default::default(),
    };
    let capability = BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Unverified,
        format: WorkbookFormat::Xlsx,
        operations: [
            EditOperationKind::DeleteSheet,
            EditOperationKind::DeleteColumn,
            EditOperationKind::SetCellNumberFormat,
            EditOperationKind::SetRowHidden,
            EditOperationKind::SetSheetVisibility,
        ]
        .into_iter()
        .collect(),
        supported_features: profile.features_present.clone(),
        preserved: plan.preserve.clone(),
        platform: platform.to_owned(),
        evidence: None,
    };
    let backend = Arc::new(UmyaXlsxBackend::new(capability)?);
    let validator: Arc<dyn OutputValidator> = Arc::new(UmyaXlsxValidator);
    let modifier = ExcelModifier::new(platform, vec![backend], validator)?;
    assert!(
        modifier.execute(&plan).is_err(),
        "unverified test capability must not be routable"
    );
    let backend = UmyaXlsxBackend::new(BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Unverified,
        format: WorkbookFormat::Xlsx,
        operations: [
            EditOperationKind::DeleteColumn,
            EditOperationKind::SetCellNumberFormat,
            EditOperationKind::SetRowHidden,
            EditOperationKind::SetSheetVisibility,
            EditOperationKind::DeleteSheet,
        ]
        .into_iter()
        .collect(),
        supported_features: profile.features_present.clone(),
        preserved: plan.preserve.clone(),
        platform: platform.to_owned(),
        evidence: None,
    })?;
    backend
        .apply(&plan)
        .map_err(|failure| anyhow::anyhow!("{:?}: {}", failure.kind, failure.message))?;
    UmyaXlsxValidator.validate(&plan, &output)?;
    assert_eq!(fs::read(&input)?, source_bytes, "source workbook changed");

    let actual = excel_editor::open_xlsx(&output)?;
    assert!(excel_reader::read_defined_names(&output)?.is_empty());
    let target = actual.sheet_by_name("表单")?;
    assert_eq!(
        target.cell("A1").map(|cell| cell.value().to_string()),
        Some("保留 A".to_owned())
    );
    assert_eq!(
        target.cell("B1").map(|cell| cell.value().to_string()),
        Some("12.5".to_owned())
    );
    assert_eq!(
        target.cell("C1").map(|cell| cell.value().to_string()),
        Some("保留 D".to_owned())
    );
    assert_eq!(
        target.cell("D1").map(|cell| cell.formula()),
        Some("SUM(B1:C1)+$B$1")
    );
    assert!(!target.cell("B2").is_some_and(|cell| cell.is_formula()));
    assert_eq!(target.merge_cells()[0].range(), "B3:C3");
    assert_eq!(target.comments()[0].coordinate().to_string(), "C2");
    assert!(target.row_dimension(5).is_some_and(|row| row.hidden()));
    assert_eq!(
        target
            .cell("B1")
            .and_then(|cell| cell.style().number_format())
            .map(|format| format.format_code()),
        Some("0.00%")
    );
    assert_eq!(visibility_of(target), SheetVisibility::Hidden);
    let other = actual.sheet_by_name("引用")?;
    assert_eq!(
        other.cell("B1").map(|cell| cell.value().to_string()),
        Some("other sheet unchanged".to_owned())
    );
    assert_eq!(
        other.cell("A1").map(|cell| cell.formula()),
        Some("'表单'!$B$1+C1")
    );
    assert!(actual.sheet_by_name("过期配置").is_err());
    assert_eq!(
        visibility_of(actual.active_sheet()),
        SheetVisibility::Visible,
        "active worksheet must remain visible"
    );
    Ok(())
}

#[test]
fn column_deletion_refuses_defined_names_before_writing() -> Result<()> {
    let dir = TestDir::new();
    let input = dir.path("input_with_name.xlsx");
    let output = dir.path("output.xlsx");
    let mut workbook = umya_spreadsheet::new_file();
    let sheet = workbook.sheet_by_name_mut("Sheet1")?;
    sheet.cell_mut("A1").set_value_string("keep");
    sheet.cell_mut("B1").set_value_string("remove");
    let mut name = umya_spreadsheet::DefinedName::default();
    name.set_name("NamedCell").set_address("Sheet1!$A$1");
    workbook.add_defined_names(name);
    umya_spreadsheet::writer::xlsx::write(&workbook, &input)?;
    assert!(!excel_reader::read_defined_names(&input)?.is_empty());

    let profile = inspect_xlsx_profile(&input)?;
    let plan = ExcelEditPlan {
        request_id: "delete-column-with-defined-name".to_owned(),
        input: input.clone(),
        output: output.clone(),
        format: WorkbookFormat::Xlsx,
        workbook: profile.clone(),
        operations: vec![EditOperation::DeleteColumn {
            sheet: "Sheet1".to_owned(),
            column_index: 2,
        }],
        preserve: [PreservationRequirement::DefinedNames]
            .into_iter()
            .collect(),
        drop_allowed: Default::default(),
    };
    let backend = UmyaXlsxBackend::new(BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Unverified,
        format: WorkbookFormat::Xlsx,
        operations: [EditOperationKind::DeleteColumn].into_iter().collect(),
        supported_features: profile.features_present,
        preserved: plan.preserve.clone(),
        platform: "synthetic-s3-f3-edit-test".to_owned(),
        evidence: None,
    })?;

    let failure = backend
        .apply(&plan)
        .expect_err("named workbooks must be refused before write");
    assert_eq!(failure.kind, BackendErrorKind::UnsupportedBeforeWrite);
    assert!(
        !output.exists(),
        "preflight refusal must not create an output file"
    );
    Ok(())
}

#[test]
fn column_deletion_refuses_formula_references_to_the_removed_column() -> Result<()> {
    let dir = TestDir::new();
    let input = dir.path("input_with_cross_sheet_reference.xlsx");
    let output = dir.path("output.xlsx");
    let mut workbook = umya_spreadsheet::new_file();
    workbook
        .sheet_by_name_mut("Sheet1")?
        .cell_mut("B1")
        .set_value_number(1.0);
    let references = workbook.new_sheet("References")?;
    references.cell_mut("A1").set_formula("Sheet1!B1");
    umya_spreadsheet::writer::xlsx::write(&workbook, &input)?;

    let profile = inspect_xlsx_profile(&input)?;
    let plan = ExcelEditPlan {
        request_id: "delete-column-with-dependent-formula".to_owned(),
        input: input.clone(),
        output: output.clone(),
        format: WorkbookFormat::Xlsx,
        workbook: profile.clone(),
        operations: vec![EditOperation::DeleteColumn {
            sheet: "Sheet1".to_owned(),
            column_index: 2,
        }],
        preserve: [PreservationRequirement::FormulaText].into_iter().collect(),
        drop_allowed: Default::default(),
    };
    let backend = UmyaXlsxBackend::new(BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Unverified,
        format: WorkbookFormat::Xlsx,
        operations: [EditOperationKind::DeleteColumn].into_iter().collect(),
        supported_features: profile.features_present,
        preserved: plan.preserve.clone(),
        platform: "synthetic-s3-f3-edit-test".to_owned(),
        evidence: None,
    })?;

    let failure = backend
        .apply(&plan)
        .expect_err("dependent formulas must not be silently rewritten to #REF");
    assert_eq!(failure.kind, BackendErrorKind::UnsupportedBeforeWrite);
    assert!(
        !output.exists(),
        "preflight refusal must not create an output file"
    );
    Ok(())
}

#[test]
#[ignore = "uses the checked-in Audit 3.3 workbook as a real-template smoke test"]
fn real_audit_s3_f3_config_supports_the_narrow_edit_batch() -> Result<()> {
    let dir = TestDir::new();
    let input = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../audit-core/config/默认配置/3.3大集中指标比较拆分_配置.xlsx");
    let output = dir.path("s3_f3_output.xlsx");
    anyhow::ensure!(input.is_file(), "real Audit 3.3 config is missing");
    anyhow::ensure!(
        excel_reader::read_defined_names(&input)?.is_empty(),
        "the S3-F3 template now has defined names; review the backend's conservative refusal"
    );
    let source_bytes = fs::read(&input)?;
    let profile = inspect_xlsx_profile(&input)?;
    let operations = vec![
        EditOperation::DeleteSheet {
            name: "使用说明".to_owned(),
        },
        EditOperation::DeleteColumn {
            sheet: "A1411".to_owned(),
            column_index: 7,
        },
        EditOperation::SetCellNumberFormat {
            sheet: "A1411".to_owned(),
            cell: "B2".to_owned(),
            format_code: "0.00%".to_owned(),
        },
        EditOperation::SetRowHidden {
            sheet: "A1411".to_owned(),
            row_index: 5,
            hidden: true,
        },
        EditOperation::SetSheetVisibility {
            sheet: "A1411".to_owned(),
            visibility: SheetVisibility::Hidden,
        },
    ];
    let plan = ExcelEditPlan {
        request_id: "real-s3-f3-smoke".to_owned(),
        input: input.clone(),
        output: output.clone(),
        format: WorkbookFormat::Xlsx,
        workbook: profile.clone(),
        operations: operations.clone(),
        preserve: [
            PreservationRequirement::CellValuesAndTypes,
            PreservationRequirement::FormulaText,
            PreservationRequirement::CommentsAndPositions,
            PreservationRequirement::MergedRanges,
            PreservationRequirement::DefinedNames,
            PreservationRequirement::RowAndColumnDimensions,
            PreservationRequirement::SheetIdentityAndOrder,
        ]
        .into_iter()
        .collect(),
        drop_allowed: Default::default(),
    };
    let backend = UmyaXlsxBackend::new(BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Unverified,
        format: WorkbookFormat::Xlsx,
        operations: operations.iter().map(EditOperation::kind).collect(),
        supported_features: profile.features_present,
        preserved: plan.preserve.clone(),
        platform: "windows-x86_64-real-template-smoke".to_owned(),
        evidence: None,
    })?;

    backend
        .apply(&plan)
        .map_err(|failure| anyhow::anyhow!("{:?}: {}", failure.kind, failure.message))?;
    UmyaXlsxValidator.validate(&plan, &output)?;
    assert_eq!(fs::read(&input)?, source_bytes, "source template changed");
    assert!(excel_editor::open_xlsx(&output)?
        .sheet_by_name("使用说明")
        .is_err());
    Ok(())
}

fn visibility_of(sheet: &umya_spreadsheet::Worksheet) -> SheetVisibility {
    match sheet.state() {
        umya_spreadsheet::SheetStateValues::Visible => SheetVisibility::Visible,
        umya_spreadsheet::SheetStateValues::Hidden => SheetVisibility::Hidden,
        umya_spreadsheet::SheetStateValues::VeryHidden => SheetVisibility::VeryHidden,
    }
}
