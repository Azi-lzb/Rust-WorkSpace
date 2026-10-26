use anyhow::{Context, Result};
use excel_editor::{
    inspect_xlsx_profile, BackendCapability, CapabilityStatus, EditBackend, EditOperation,
    EditOperationKind, EngineId, ExcelEditPlan, RangeStyleMode, UmyaXlsxBackend, WorkbookFormat,
};
use std::{
    collections::BTreeSet,
    fs,
    path::{Path, PathBuf},
    time::{SystemTime, UNIX_EPOCH},
};

fn temp_path(label: &str) -> PathBuf {
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("clock after epoch")
        .as_nanos();
    std::env::temp_dir().join(format!(
        "excel_editor_formula_copy_{}_{}_{}",
        std::process::id(),
        nonce,
        label
    ))
}

fn write_source(path: &Path) -> Result<()> {
    let mut workbook = umya_spreadsheet::new_file();
    let source = workbook
        .get_sheet_by_name_mut("Sheet1")
        .context("default worksheet missing")?;
    source.set_name("Source");
    source.cell_mut("A1").set_value("untouched");
    source.cell_mut("B2").set_formula("SUM(C2:D2)");
    source.cell_mut("B3").set_formula("SUM($C$2:D2)");
    source.cell_mut("B4").set_formula("$C2+D$2");
    source.cell_mut("B5").set_formula("\"C5\"&C5");
    source.cell_mut("B6").set_formula("'Other Sheet'!$A$1+C6");
    source.cell_mut("B7").set_formula("SUM(Table1[Amount])");
    source
        .cell_mut("B8")
        .set_formula("'[External.xlsx]Sheet1'!A1");
    workbook.new_sheet("Other Sheet")?;
    umya_spreadsheet::writer::xlsx::write(&workbook, path)
        .with_context(|| format!("write source fixture {}", path.display()))
}

fn write_target(path: &Path) -> Result<()> {
    let mut workbook = umya_spreadsheet::new_file();
    workbook
        .get_sheet_by_name_mut("Sheet1")
        .context("default worksheet missing")?
        .set_name("Target");
    workbook.new_sheet("Other Sheet")?;
    umya_spreadsheet::writer::xlsx::write(&workbook, path)
        .with_context(|| format!("write target fixture {}", path.display()))
}

fn backend(operations: impl IntoIterator<Item = EditOperationKind>) -> Result<UmyaXlsxBackend> {
    Ok(UmyaXlsxBackend::new(BackendCapability {
        engine: EngineId::UmyaSpreadsheet,
        status: CapabilityStatus::Unverified,
        format: WorkbookFormat::Xlsx,
        operations: operations.into_iter().collect(),
        supported_features: BTreeSet::new(),
        preserved: BTreeSet::new(),
        platform: "test-only-direct-backend".to_owned(),
        evidence: None,
    })?)
}

fn plan(input: &Path, output: &Path, operations: Vec<EditOperation>) -> Result<ExcelEditPlan> {
    Ok(ExcelEditPlan {
        request_id: "formula-copy-translation-test".to_owned(),
        input: input.to_path_buf(),
        output: output.to_path_buf(),
        format: WorkbookFormat::Xlsx,
        workbook: inspect_xlsx_profile(input)?,
        operations,
        preserve: BTreeSet::new(),
        drop_allowed: Default::default(),
    })
}

fn formulas(path: &Path) -> Result<BTreeSet<(String, String)>> {
    let dump = excel_reader::read_workbook(path)?;
    Ok(dump
        .sheets
        .into_iter()
        .flat_map(|sheet| {
            sheet
                .formulas
                .into_iter()
                .map(move |(cell, formula)| (cell, formula.trim_start_matches('=').to_owned()))
        })
        .collect())
}

#[test]
fn same_workbook_translation_adjusts_relative_absolute_mixed_and_sheet_refs() -> Result<()> {
    let input = temp_path("input.xlsx");
    let output = temp_path("output.xlsx");
    write_source(&input)?;
    let source_before = fs::read(&input)?;
    let copy_plan = plan(
        &input,
        &output,
        vec![EditOperation::CopyRangeWithFormulaTranslation {
            source_sheet: "Source".to_owned(),
            source_range: "B2:B6".to_owned(),
            target_sheet: "Source".to_owned(),
            target_start: "B10".to_owned(),
            copy_values: true,
            copy_formulas: true,
            copy_styles: false,
            copy_comments: false,
        }],
    )?;
    copy_plan.validate()?;
    backend([EditOperationKind::CopyRangeWithFormulaTranslation])?
        .apply(&copy_plan)
        .map_err(|failure| anyhow::anyhow!(failure.message))?;

    let output_formulas = formulas(&output)?;
    assert!(output_formulas.contains(&("B10".to_owned(), "SUM(C10:D10)".to_owned())));
    assert!(output_formulas.contains(&("B11".to_owned(), "SUM($C$2:D10)".to_owned())));
    assert!(output_formulas.contains(&("B12".to_owned(), "$C10+D$2".to_owned())));
    assert!(output_formulas.contains(&("B13".to_owned(), "\"C5\"&C13".to_owned())));
    assert!(output_formulas.contains(&("B14".to_owned(), "'Other Sheet'!$A$1+C14".to_owned())));
    assert!(output_formulas.contains(&("B2".to_owned(), "SUM(C2:D2)".to_owned())));
    let output_dump = excel_reader::read_workbook(&output)?;
    let unchanged_value = output_dump.sheets[0].values.get("A1");
    assert_eq!(
        unchanged_value,
        Some(&serde_json::json!({"type":"string", "value":"untouched"}))
    );
    assert_eq!(fs::read(&input)?, source_before);

    let _ = fs::remove_file(input);
    let _ = fs::remove_file(output);
    Ok(())
}

#[test]
fn existing_copy_range_keeps_formula_text_verbatim() -> Result<()> {
    let input = temp_path("raw-input.xlsx");
    let output = temp_path("raw-output.xlsx");
    write_source(&input)?;
    let plan = plan(
        &input,
        &output,
        vec![EditOperation::CopyRange {
            source_sheet: "Source".to_owned(),
            source_range: "B2:B2".to_owned(),
            target_sheet: "Source".to_owned(),
            target_start: "B10".to_owned(),
            copy_values: true,
            copy_formulas: true,
            copy_styles: false,
            copy_comments: false,
        }],
    )?;
    plan.validate()?;
    backend([EditOperationKind::CopyRange])?
        .apply(&plan)
        .map_err(|failure| anyhow::anyhow!(failure.message))?;
    assert!(formulas(&output)?.contains(&("B10".to_owned(), "SUM(C2:D2)".to_owned())));

    let _ = fs::remove_file(input);
    let _ = fs::remove_file(output);
    Ok(())
}

#[test]
fn import_range_translation_works_across_workbooks_and_rejects_structured_refs() -> Result<()> {
    let source = temp_path("import-source.xlsx");
    let target = temp_path("import-target.xlsx");
    let output = temp_path("import-output.xlsx");
    write_source(&source)?;
    write_target(&target)?;
    let import_plan = plan(
        &target,
        &output,
        vec![EditOperation::ImportRangeWithFormulaTranslation {
            source_workbook: source.clone(),
            source_sheet: "Source".to_owned(),
            source_range: "B2:B6".to_owned(),
            target_sheet: "Target".to_owned(),
            target_start: "B10".to_owned(),
            copy_values: true,
            copy_formulas: true,
            style_mode: RangeStyleMode::None,
            copy_comments: false,
            copy_merges: false,
            copy_row_heights: false,
            copy_column_widths: false,
            copy_row_hidden: false,
            copy_column_hidden: false,
            copy_outline_levels: false,
            copy_all_row_dimensions: false,
            copy_all_column_dimensions: false,
            materialize_blank_cells: false,
            hide_zero: false,
        }],
    )?;
    import_plan.validate()?;
    backend([EditOperationKind::ImportRangeWithFormulaTranslation])?
        .apply(&import_plan)
        .map_err(|failure| anyhow::anyhow!(failure.message))?;
    let output_formulas = formulas(&output)?;
    assert!(output_formulas.contains(&("B10".to_owned(), "SUM(C10:D10)".to_owned())));
    assert!(output_formulas.contains(&("B11".to_owned(), "SUM($C$2:D10)".to_owned())));
    assert!(output_formulas.contains(&("B12".to_owned(), "$C10+D$2".to_owned())));
    assert!(output_formulas.contains(&("B13".to_owned(), "\"C5\"&C13".to_owned())));
    assert!(output_formulas.contains(&("B14".to_owned(), "'Other Sheet'!$A$1+C14".to_owned())));

    let raw_output = temp_path("raw-import-output.xlsx");
    let raw_import = plan(
        &target,
        &raw_output,
        vec![EditOperation::ImportRange {
            source_workbook: source.clone(),
            source_sheet: "Source".to_owned(),
            source_range: "B2:B2".to_owned(),
            target_sheet: "Target".to_owned(),
            target_start: "B10".to_owned(),
            copy_values: true,
            copy_formulas: true,
            style_mode: RangeStyleMode::None,
            copy_comments: false,
            copy_merges: false,
            copy_row_heights: false,
            copy_column_widths: false,
            copy_row_hidden: false,
            copy_column_hidden: false,
            copy_outline_levels: false,
            copy_all_row_dimensions: false,
            copy_all_column_dimensions: false,
            materialize_blank_cells: false,
            hide_zero: false,
        }],
    )?;
    raw_import.validate()?;
    backend([EditOperationKind::ImportRange])?
        .apply(&raw_import)
        .map_err(|failure| anyhow::anyhow!(failure.message))?;
    assert!(formulas(&raw_output)?.contains(&("B10".to_owned(), "SUM(C2:D2)".to_owned())));

    let rejected = plan(
        &target,
        &temp_path("unsupported-output.xlsx"),
        vec![EditOperation::ImportRangeWithFormulaTranslation {
            source_workbook: source.clone(),
            source_sheet: "Source".to_owned(),
            source_range: "B7:B7".to_owned(),
            target_sheet: "Target".to_owned(),
            target_start: "B20".to_owned(),
            copy_values: true,
            copy_formulas: true,
            style_mode: RangeStyleMode::None,
            copy_comments: false,
            copy_merges: false,
            copy_row_heights: false,
            copy_column_widths: false,
            copy_row_hidden: false,
            copy_column_hidden: false,
            copy_outline_levels: false,
            copy_all_row_dimensions: false,
            copy_all_column_dimensions: false,
            materialize_blank_cells: false,
            hide_zero: false,
        }],
    )?;
    let error = backend([EditOperationKind::ImportRangeWithFormulaTranslation])?
        .apply(&rejected)
        .expect_err("structured-reference formulas must be refused");
    assert!(error
        .message
        .contains("does not support structured/external"));
    assert!(!rejected.output.exists());

    let external_rejected = plan(
        &target,
        &temp_path("external-output.xlsx"),
        vec![EditOperation::ImportRangeWithFormulaTranslation {
            source_workbook: source.clone(),
            source_sheet: "Source".to_owned(),
            source_range: "B8:B8".to_owned(),
            target_sheet: "Target".to_owned(),
            target_start: "B20".to_owned(),
            copy_values: true,
            copy_formulas: true,
            style_mode: RangeStyleMode::None,
            copy_comments: false,
            copy_merges: false,
            copy_row_heights: false,
            copy_column_widths: false,
            copy_row_hidden: false,
            copy_column_hidden: false,
            copy_outline_levels: false,
            copy_all_row_dimensions: false,
            copy_all_column_dimensions: false,
            materialize_blank_cells: false,
            hide_zero: false,
        }],
    )?;
    let error = backend([EditOperationKind::ImportRangeWithFormulaTranslation])?
        .apply(&external_rejected)
        .expect_err("external-workbook formulas must be refused");
    assert!(error.message.contains("external workbook references"));
    assert!(!external_rejected.output.exists());

    let _ = fs::remove_file(source);
    let _ = fs::remove_file(target);
    let _ = fs::remove_file(output);
    let _ = fs::remove_file(raw_output);
    Ok(())
}
