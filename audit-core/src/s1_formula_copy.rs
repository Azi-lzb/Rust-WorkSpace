//! S1 formula-copy business planning.
//!
//! This module translates the resolved Audit template ranges into one
//! workbook-level `ExcelEditPlan`. It does not pick an unverified editor or
//! recalculate formulas; those decisions remain with ExcelModifier and the
//! explicitly requested formula sidecar respectively.

use crate::{
    models::CopyRange,
    s1_template_rules::{load_s1_template_rule_plan, S1TemplateRuleSummary},
};
use anyhow::{ensure, Context, Result};
use excel_editor::{
    inspect_xlsx_profile, EditOperation, ExcelEditPlan, PreservationRequirement, RangeStyleMode,
    WorkbookProfile,
};
use excel_reader::{read_ooxml_comments, WorkbookComment};
use std::{
    collections::{BTreeMap, BTreeSet},
    fs,
    path::Path,
};

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
    let target_sheet_names = excel_reader::read_xlsx_sheet_names_ooxml(target_workbook)
        .context("读取 S1 目标工作簿工作表名")?;
    for range in &template.definition.copy_ranges {
        ensure!(
            target_sheet_names
                .iter()
                .any(|sheet_name| sheet_name == &range.sheet_name),
            "S1 公式复制目标工作簿缺少模板映射工作表“{}”",
            range.sheet_name
        );
    }
    let template_profile =
        inspect_xlsx_profile(template_workbook).context("检查 S1 公式模板结构")?;
    let template_comments =
        read_ooxml_comments(template_workbook).context("读取 S1 公式模板批注作者和文本")?;

    let mut operations = Vec::with_capacity(template.definition.copy_ranges.len() + 1);
    let mut copied_ranges = BTreeSet::new();
    for range in &template.definition.copy_ranges {
        let identity = (range.sheet_name.clone(), range.address.to_ascii_uppercase());
        if !copied_ranges.insert(identity) {
            continue;
        }
        let target_start = range_target_start(&range.address)
            .with_context(|| format!("解析公式复制目标坐标 {}", range.address))?;
        let comment_operations = comment_operations_for_range(
            range,
            &target_start,
            template_comments.get(&range.sheet_name),
        )?;
        operations.push(EditOperation::ImportRange {
            source_workbook: template_canonical.clone(),
            source_sheet: range.sheet_name.clone(),
            source_range: range.address.clone(),
            target_sheet: range.sheet_name.clone(),
            target_start,
            copy_values: true,
            copy_formulas: true,
            style_mode: RangeStyleMode::Full,
            // Copy nonempty source comments as per-cell SetComment operations
            // below. This preserves existing target comments where the source
            // template has no comment at that coordinate.
            copy_comments: false,
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
        operations.extend(comment_operations);
    }
    operations.push(EditOperation::SetWorkbookCalculationProperties {
        full_calc_on_load: Some(true),
        force_full_calc: Some(true),
    });

    // Existing values in merged non-anchor cells are still workbook data.
    // Do not inherit the legacy Python save behavior that drops them. A
    // backend that cannot preserve these values must reject this plan.
    let drop_allowed: BTreeMap<String, BTreeSet<PreservationRequirement>> = BTreeMap::new();

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
        drop_allowed,
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
    let range = range_without_sheet(address).replace('$', "");
    let first_cell = range.split(':').next().unwrap_or(&range).trim();
    let (row, column) = parse_a1_cell(first_cell).context("区域没有有效的起始单元格")?;
    Ok(a1_cell(row, column))
}

fn comment_operations_for_range(
    range: &CopyRange,
    target_start: &str,
    comments: Option<&BTreeMap<String, WorkbookComment>>,
) -> Result<Vec<EditOperation>> {
    let Some(comments) = comments else {
        return Ok(Vec::new());
    };
    let mut operations = Vec::new();
    for (source_cell, comment) in comments {
        let Some(target_cell) =
            map_source_cell_to_target(&range.address, target_start, source_cell)?
        else {
            continue;
        };
        let author = comment
            .author
            .as_deref()
            .context("S1 模板批注作者无法解析")?;
        operations.push(EditOperation::SetComment {
            sheet: range.sheet_name.clone(),
            cell: target_cell,
            text: Some(comment.text.clone()),
            // openpyxl's legacy copy path exposes an empty XML author as the
            // string "None". Preserve that output while rejecting unresolved
            // authorId references above.
            author: Some(if author.is_empty() {
                "None".to_owned()
            } else {
                author.to_owned()
            }),
        });
    }
    Ok(operations)
}

fn map_source_cell_to_target(
    source_range: &str,
    target_start: &str,
    source_cell: &str,
) -> Result<Option<String>> {
    let (source_start_row, source_start_column, source_end_row, source_end_column) =
        parse_a1_range(source_range)?;
    let (target_start_row, target_start_column) = parse_a1_cell(target_start)?;
    let (source_row, source_column) = parse_a1_cell(source_cell)?;
    if source_row < source_start_row
        || source_row > source_end_row
        || source_column < source_start_column
        || source_column > source_end_column
    {
        return Ok(None);
    }
    let target_row = target_start_row
        .checked_add(source_row - source_start_row)
        .context("批注目标行坐标溢出")?;
    let target_column = target_start_column
        .checked_add(source_column - source_start_column)
        .context("批注目标列坐标溢出")?;
    Ok(Some(a1_cell(target_row, target_column)))
}

fn parse_a1_range(address: &str) -> Result<(u32, u32, u32, u32)> {
    let range = range_without_sheet(address).replace('$', "");
    let (start, end) = range.split_once(':').unwrap_or((&range, &range));
    let (start_row, start_column) = parse_a1_cell(start.trim())?;
    let (end_row, end_column) = parse_a1_cell(end.trim())?;
    ensure!(
        start_row <= end_row && start_column <= end_column,
        "公式复制区域边界无效"
    );
    Ok((start_row, start_column, end_row, end_column))
}

fn range_without_sheet(address: &str) -> &str {
    address.rsplit('!').next().unwrap_or(address).trim()
}

fn parse_a1_cell(address: &str) -> Result<(u32, u32)> {
    let mut row = 0u32;
    let mut column = 0u32;
    let mut reading_row = false;
    for character in address.trim().chars() {
        if character == '$' {
            continue;
        }
        if character.is_ascii_alphabetic() && !reading_row {
            column = column
                .checked_mul(26)
                .and_then(|value| {
                    value.checked_add(character.to_ascii_uppercase() as u32 - 'A' as u32 + 1)
                })
                .context("单元格列坐标溢出")?;
        } else if character.is_ascii_digit() {
            reading_row = true;
            row = row
                .checked_mul(10)
                .and_then(|value| value.checked_add(character as u32 - '0' as u32))
                .context("单元格行坐标溢出")?;
        } else {
            anyhow::bail!("无效 A1 单元格坐标");
        }
    }
    ensure!(reading_row && row > 0 && column > 0, "无效 A1 单元格坐标");
    Ok((row, column))
}

fn a1_cell(row: u32, mut column: u32) -> String {
    let mut letters = Vec::new();
    while column > 0 {
        let digit = (column - 1) % 26;
        letters.push((b'A' + digit as u8) as char);
        column = (column - 1) / 26;
    }
    format!("{}{}", letters.into_iter().rev().collect::<String>(), row)
}

#[cfg(test)]
mod tests {
    use super::*;
    use excel_editor::{EditOperationKind, WorkbookFormat};
    use excel_reader::WorkbookComment;
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

    fn write_structured_template_with_duplicate_ranges(path: &Path) {
        let mut workbook = Workbook::new();
        {
            let rules = workbook.add_worksheet();
            rules
                .set_name("审核规则")
                .expect("name structured rule worksheet");
            for (column, header) in crate::template::REQUIRED_HEADERS.iter().enumerate() {
                rules
                    .write_string(0, column as u16, *header)
                    .expect("write structured rule header");
            }
            for (row, rule_id) in [(1, "R-01"), (2, "R-02")] {
                for (column, value) in [
                    rule_id,
                    "是",
                    "报表甲",
                    "检查表",
                    "A1",
                    "B1",
                    "错误",
                    "核对错误项",
                ]
                .iter()
                .enumerate()
                {
                    rules
                        .write_string(row, column as u16, *value)
                        .expect("write structured rule row");
                }
            }
        }
        let formulas = workbook.add_worksheet();
        formulas.set_name("检查表").expect("name formula worksheet");
        formulas
            .write_formula(0, 0, "=IF(TRUE,\"错误|字段|指标|说明\",\"通过\")")
            .expect("write structured formula");
        workbook
            .save(path)
            .expect("save duplicate-range structured template");
    }

    fn write_target(path: &Path) {
        write_target_with_sheet(path, "检查表");
    }

    fn write_target_with_sheet(path: &Path, sheet_name: &str) {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_worksheet();
        sheet.set_name(sheet_name).expect("name target sheet");
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
        assert!(
            prepared.edit_plan.drop_allowed.is_empty(),
            "S1 formula copy must not allow a backend to drop merged-cell payloads"
        );
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
                copy_comments: false,
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
    fn refuses_formula_copy_when_target_lacks_a_mapped_template_sheet() {
        let directory = temporary_directory();
        let target = directory.join("target.xlsx");
        let template = directory.join("template.xlsx");
        let output = directory.join("output.xlsx");
        write_target_with_sheet(&target, "其他工作表");
        write_template(&template);

        let error = build_s1_formula_copy_edit_plan(&target, &template, &output, None)
            .expect_err("missing target sheet must stop plan construction");

        assert!(
            format!("{error:#}").contains("目标工作簿缺少模板映射工作表"),
            "{error:#}"
        );
        assert!(
            !output.exists(),
            "preflight must not create candidate output"
        );
        fs::remove_dir_all(directory).expect("remove formula-copy test directory");
    }

    #[test]
    fn deduplicates_identical_structured_formula_copy_ranges() {
        let directory = temporary_directory();
        let target = directory.join("target.xlsx");
        let template = directory.join("template.xlsx");
        let output = directory.join("output.xlsx");
        write_target(&target);
        write_structured_template_with_duplicate_ranges(&template);

        let prepared = build_s1_formula_copy_edit_plan(&target, &template, &output, None)
            .expect("build plan with duplicate structured ranges");
        let import_range_count = prepared
            .edit_plan
            .operations
            .iter()
            .filter(|operation| operation.kind() == EditOperationKind::ImportRange)
            .count();

        assert_eq!(prepared.template_summary.copy_ranges, 2);
        assert_eq!(import_range_count, 1);
        assert_eq!(prepared.edit_plan.operations.len(), 2);
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

    #[test]
    fn maps_only_template_comments_in_copy_range_and_keeps_explicit_author() {
        let range = CopyRange {
            sheet_name: "检查表".to_owned(),
            address: "$C$4:$D$6".to_owned(),
        };
        let comments = BTreeMap::from([
            (
                "C4".to_owned(),
                WorkbookComment {
                    text: "top-left".to_owned(),
                    author: Some("reviewer".to_owned()),
                },
            ),
            (
                "D6".to_owned(),
                WorkbookComment {
                    text: "bottom-right".to_owned(),
                    author: Some("reviewer".to_owned()),
                },
            ),
            (
                "D5".to_owned(),
                WorkbookComment {
                    text: "empty-author".to_owned(),
                    author: Some(String::new()),
                },
            ),
            (
                "B3".to_owned(),
                WorkbookComment {
                    text: "outside".to_owned(),
                    author: Some("reviewer".to_owned()),
                },
            ),
        ]);

        let operations = comment_operations_for_range(&range, "F9", Some(&comments))
            .expect("create mapped template comment operations");

        assert_eq!(operations.len(), 3);
        assert!(matches!(
            &operations[0],
            EditOperation::SetComment {
                sheet,
                cell,
                text: Some(text),
                author: Some(author),
            } if sheet == "检查表" && cell == "F9" && text == "top-left" && author == "reviewer"
        ));
        assert!(matches!(
            &operations[1],
            EditOperation::SetComment {
                sheet,
                cell,
                text: Some(text),
                author: Some(author),
            } if sheet == "检查表" && cell == "G10" && text == "empty-author" && author == "None"
        ));
        assert!(matches!(
            &operations[2],
            EditOperation::SetComment {
                sheet,
                cell,
                text: Some(text),
                author: Some(author),
            } if sheet == "检查表" && cell == "G11" && text == "bottom-right" && author == "reviewer"
        ));
    }

    #[test]
    fn emits_no_comment_operation_when_template_has_no_comment_in_range() {
        let range = CopyRange {
            sheet_name: "检查表".to_owned(),
            address: "A1:B2".to_owned(),
        };
        let comments = BTreeMap::from([(
            "C3".to_owned(),
            WorkbookComment {
                text: "outside".to_owned(),
                author: Some("reviewer".to_owned()),
            },
        )]);

        let operations = comment_operations_for_range(&range, "A1", Some(&comments))
            .expect("skip comments outside the copied range");

        assert!(operations.is_empty());
    }

    #[test]
    fn rejects_an_in_range_comment_with_unresolved_author() {
        let range = CopyRange {
            sheet_name: "检查表".to_owned(),
            address: "A1:B2".to_owned(),
        };
        let comments = BTreeMap::from([(
            "A1".to_owned(),
            WorkbookComment {
                text: "note".to_owned(),
                author: None,
            },
        )]);

        let error = comment_operations_for_range(&range, "A1", Some(&comments))
            .expect_err("unresolved authorId must not be replaced with a default");

        assert!(format!("{error:#}").contains("批注作者无法解析"));
    }
}
