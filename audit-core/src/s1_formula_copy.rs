//! S1 formula-copy business planning.
//!
//! This module translates the resolved Audit template ranges into one
//! workbook-level `ExcelEditPlan`. It does not pick an unverified editor or
//! recalculate formulas; those decisions remain with ExcelModifier and the
//! explicitly requested formula sidecar respectively.

use crate::s1_template_rules::{load_s1_template_rule_plan, S1TemplateRuleSummary};
use anyhow::{ensure, Context, Result};
use excel_editor::{
    inspect_xlsx_profile, EditOperation, ExcelEditPlan, PreservationRequirement, RangeStyleMode,
    WorkbookProfile,
};
use std::{collections::BTreeMap, fs, path::Path};

#[derive(Clone, Debug)]
pub struct S1FormulaCopyEditPlan {
    pub edit_plan: ExcelEditPlan,
    /// Kept separate because an ImportRange's source workbook also has to
    /// match the selected backend's verified feature set.
    pub template_profile: WorkbookProfile,
    pub template_summary: S1TemplateRuleSummary,
}

/// Build a non-destructive, workbook-level plan for copying formula ranges
/// from one S1 template into a new audit-workbook copy.
///
/// The output must not exist. This function only reads the target/template and
/// creates an edit plan; a production caller must still route it through
/// ExcelModifier and a Passed capability for the exact operation batch.
pub fn build_s1_formula_copy_edit_plan(
    target_workbook: impl AsRef<Path>,
    template_workbook: impl AsRef<Path>,
    output_workbook: impl AsRef<Path>,
    mapping_config: Option<&Path>,
) -> Result<S1FormulaCopyEditPlan> {
    let target_workbook = target_workbook.as_ref();
    let template_workbook = template_workbook.as_ref();
    let output_workbook = output_workbook.as_ref();

    ensure!(target_workbook.is_file(), "S1 公式复制目标工作簿不存在");
    ensure!(template_workbook.is_file(), "S1 公式模板工作簿不存在");
    ensure!(
        xlsx_extension(target_workbook),
        "S1 公式复制目标目前仅支持 .xlsx"
    );
    ensure!(
        xlsx_extension(template_workbook),
        "S1 公式复制模板目前仅支持 .xlsx"
    );
    ensure!(
        xlsx_extension(output_workbook),
        "S1 公式复制输出目前仅支持 .xlsx"
    );
    ensure!(
        !output_workbook.exists(),
        "S1 公式复制输出文件已存在，拒绝覆盖"
    );
    let output_parent = output_workbook
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    ensure!(output_parent.is_dir(), "S1 公式复制输出目录不存在");

    let target_canonical = fs::canonicalize(target_workbook).context("规范化 S1 目标工作簿路径")?;
    let template_canonical =
        fs::canonicalize(template_workbook).context("规范化 S1 模板工作簿路径")?;
    ensure!(
        !same_path(&target_canonical, &template_canonical),
        "S1 公式复制目标与模板必须是不同工作簿"
    );

    let template = load_s1_template_rule_plan(template_workbook, mapping_config)
        .context("解析 S1 公式复制模板规则")?;
    ensure!(
        !template.definition.copy_ranges.is_empty(),
        "S1 模板没有可复制的公式区域"
    );
    let target_profile = inspect_xlsx_profile(target_workbook).context("检查 S1 目标工作簿结构")?;
    let template_profile =
        inspect_xlsx_profile(template_workbook).context("检查 S1 公式模板结构")?;

    let mut operations = Vec::with_capacity(template.definition.copy_ranges.len() + 1);
    for range in &template.definition.copy_ranges {
        let target_start = range_target_start(&range.address)
            .with_context(|| format!("解析公式复制目标坐标 {}", range.address))?;
        operations.push(EditOperation::ImportRange {
            source_workbook: template_canonical.clone(),
            source_sheet: range.sheet_name.clone(),
            source_range: range.address.clone(),
            target_sheet: range.sheet_name.clone(),
            target_start,
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
            // The S1 copy contract overwrites the complete mapped rectangle,
            // including blank source cells.
            materialize_blank_cells: true,
            hide_zero: false,
        });
    }
    operations.push(EditOperation::SetWorkbookCalculationProperties {
        full_calc_on_load: Some(true),
        force_full_calc: Some(true),
    });

    let edit_plan = ExcelEditPlan {
        request_id: "audit-s1-formula-copy".to_owned(),
        input: target_workbook.to_owned(),
        output: output_workbook.to_owned(),
        format: target_profile.format,
        workbook: target_profile,
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
    edit_plan
        .validate()
        .context("校验 S1 公式复制 ExcelEditPlan")?;

    Ok(S1FormulaCopyEditPlan {
        edit_plan,
        template_profile,
        template_summary: template.summary,
    })
}

fn xlsx_extension(path: &Path) -> bool {
    path.extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
}

#[cfg(windows)]
fn same_path(left: &Path, right: &Path) -> bool {
    left.to_string_lossy()
        .eq_ignore_ascii_case(&right.to_string_lossy())
}

#[cfg(not(windows))]
fn same_path(left: &Path, right: &Path) -> bool {
    left == right
}

fn range_target_start(address: &str) -> Result<String> {
    let range = address.rsplit('!').next().unwrap_or(address).trim();
    let first_cell = range
        .split(':')
        .next()
        .unwrap_or(range)
        .trim()
        .replace('$', "");
    let split = first_cell
        .find(|character: char| character.is_ascii_digit())
        .context("区域没有有效的起始单元格")?;
    let (column, row) = first_cell.split_at(split);
    ensure!(
        !column.is_empty()
            && column.bytes().all(|byte| byte.is_ascii_alphabetic())
            && !row.is_empty()
            && row.bytes().all(|byte| byte.is_ascii_digit())
            && row.parse::<u32>().is_ok_and(|row| row > 0),
        "区域起始单元格坐标无效"
    );
    Ok(format!("{}{}", column.to_ascii_uppercase(), row))
}

#[cfg(test)]
mod tests {
    use super::*;
    use excel_editor::{EditOperationKind, WorkbookFormat};
    use excel_writer::Workbook;
    use std::{
        fs,
        path::PathBuf,
        time::{SystemTime, UNIX_EPOCH},
    };

    fn temporary_directory() -> PathBuf {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("system clock after Unix epoch")
            .as_nanos();
        let path = std::env::temp_dir().join(format!(
            "audit_s1_formula_copy_plan_{}_{}",
            std::process::id(),
            nonce
        ));
        fs::create_dir(&path).expect("create formula-copy test directory");
        path
    }

    fn write_template(path: &Path) {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.set_name("检查表").expect("name formula sheet");
        sheet
            .write_formula(0, 0, "=IF(TRUE,\"错误|字段|指标|说明\",\"通过\")")
            .expect("write formula");
        workbook
            .define_name("校验区域", "='检查表'!$A$1:$B$2")
            .expect("define formula-copy range");
        workbook.save(path).expect("save formula template");
    }

    fn write_target(path: &Path) {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.set_name("检查表").expect("name target sheet");
        sheet
            .write_string(0, 0, "target value")
            .expect("write target value");
        workbook.save(path).expect("save target workbook");
    }

    #[test]
    fn builds_one_protected_workbook_plan_for_template_ranges() {
        let directory = temporary_directory();
        let target = directory.join("target.xlsx");
        let template = directory.join("template.xlsx");
        let output = directory.join("output.xlsx");
        write_target(&target);
        write_template(&template);

        let prepared = build_s1_formula_copy_edit_plan(&target, &template, &output, None)
            .expect("build S1 formula-copy plan");

        assert_eq!(prepared.edit_plan.format, WorkbookFormat::Xlsx);
        assert_eq!(prepared.template_summary.copy_ranges, 1);
        assert_eq!(prepared.edit_plan.operations.len(), 2);
        assert_eq!(
            prepared
                .edit_plan
                .operations
                .iter()
                .filter(|operation| operation.kind() == EditOperationKind::ImportRange)
                .count(),
            1
        );
        assert!(matches!(
            &prepared.edit_plan.operations[0],
            EditOperation::ImportRange {
                source_range,
                target_start,
                copy_values: true,
                copy_formulas: true,
                copy_comments: true,
                materialize_blank_cells: true,
                style_mode: RangeStyleMode::Full,
                ..
            } if source_range == "A1:B2" && target_start == "A1"
        ));
        assert!(matches!(
            prepared.edit_plan.operations.last(),
            Some(EditOperation::SetWorkbookCalculationProperties {
                full_calc_on_load: Some(true),
                force_full_calc: Some(true),
            })
        ));
        assert!(!output.exists(), "plan construction must not write output");

        fs::remove_dir_all(directory).expect("remove formula-copy test directory");
    }

    #[test]
    fn refuses_to_overwrite_an_existing_formula_copy_output() {
        let directory = temporary_directory();
        let target = directory.join("target.xlsx");
        let template = directory.join("template.xlsx");
        let output = directory.join("output.xlsx");
        write_target(&target);
        write_template(&template);
        fs::write(&output, b"keep existing output").expect("create existing output sentinel");

        let error = build_s1_formula_copy_edit_plan(&target, &template, &output, None)
            .expect_err("existing output must not be replaced");

        assert!(format!("{error:#}").contains("拒绝覆盖"));
        assert_eq!(
            fs::read(&output).expect("read existing output sentinel"),
            b"keep existing output"
        );
        fs::remove_dir_all(directory).expect("remove formula-copy test directory");
    }
}
