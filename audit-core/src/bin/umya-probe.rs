use anyhow::{anyhow, Context, Result};
use clap::Parser;
use excel_editor::BorderStyleValues;
use ooxml_patch::{patch_column_dimensions, reject_unverified_parts, ColumnWidthUpdate};
use serde::Serialize;
use std::{fs, path::PathBuf};

#[derive(Parser)]
#[command(about = "Isolated umya-spreadsheet workbook modification probe")]
struct Args {
    #[arg(long)]
    input: PathBuf,
    #[arg(long)]
    output_dir: PathBuf,
    #[arg(long)]
    report: PathBuf,
}

#[derive(Serialize)]
struct Item {
    name: &'static str,
    status: &'static str,
    detail: String,
}

#[derive(Serialize)]
struct Report {
    schema_version: u32,
    library: &'static str,
    library_version: &'static str,
    input: String,
    output: String,
    patched_output: Option<String>,
    column_patch_status: String,
    elapsed_ms: u128,
    umya_library_elapsed_ms: u128,
    ooxml_patch_elapsed_ms: u128,
    items: Vec<Item>,
}

fn item(
    items: &mut Vec<Item>,
    name: &'static str,
    status: &'static str,
    detail: impl Into<String>,
) {
    items.push(Item {
        name,
        status,
        detail: detail.into(),
    });
}

fn main() -> Result<()> {
    let args = Args::parse();
    let started = std::time::Instant::now();
    let input = args.input.canonicalize().context("resolve input path")?;
    reject_unverified_parts(&input)?;
    fs::create_dir_all(&args.output_dir)?;
    if let Some(parent) = args.report.parent() {
        fs::create_dir_all(parent)?;
    }

    let stem = input
        .file_stem()
        .and_then(|x| x.to_str())
        .unwrap_or("workbook");
    let output = args.output_dir.join(format!("{stem}_umya.xlsx"));
    let patched_output = args.output_dir.join(format!("{stem}_umya_patched.xlsx"));
    let mut book =
        excel_editor::open_xlsx(&input).with_context(|| format!("read {}", input.display()))?;
    let sheet_names = book
        .sheet_collection()
        .iter()
        .map(|s| s.name().to_owned())
        .collect::<Vec<_>>();
    let sheet_name = sheet_names
        .first()
        .ok_or_else(|| anyhow!("workbook has no worksheets"))?
        .clone();
    let known_synthetic =
        input.file_name().and_then(|x| x.to_str()) == Some("synthetic_probe.xlsx");
    let mut items = Vec::new();

    {
        let sheet = book.sheet_by_name_mut(&sheet_name)?;
        sheet.cell_mut("A1").set_value("umya-probe ordinary value");
        item(
            &mut items,
            "ordinary_cell",
            "success",
            "A1 set to probe marker",
        );

        if sheet.merge_cells().iter().any(|r| r.range() == "B2:C2") {
            sheet.cell_mut("B2").set_value("umya-probe merged anchor");
            item(
                &mut items,
                "merged_anchor",
                "success",
                "B2 changed; B2:C2 exists",
            );
        } else if known_synthetic {
            item(
                &mut items,
                "merged_anchor",
                "failed",
                "expected synthetic merge B2:C2 was absent",
            );
        } else {
            item(
                &mut items,
                "merged_anchor",
                "not_applicable",
                "worksheet has no B2:C2 merge",
            );
        }

        let style = sheet.style_mut("C4");
        style.font_mut().set_bold(true);
        style.set_background_color("FFFFFF00");
        style
            .borders_mut()
            .left_mut()
            .set_style(BorderStyleValues::Thin);
        style.number_format_mut().set_format_code("0.00");
        item(
            &mut items,
            "font_fill_border_number_format",
            "success",
            "C4 bold, yellow fill, thin left border, number format 0.00",
        );

        sheet.cell_mut("D4").set_formula("SUM(A1:A3)");
        item(
            &mut items,
            "formula_text",
            "success",
            "D4 formula text set; no recalculation requested",
        );
    }

    let name_found = book
        .defined_names()
        .iter()
        .any(|n| n.name() == "SyntheticNamedCell")
        || book
            .sheet_by_name(&sheet_name)?
            .defined_names()
            .iter()
            .any(|n| n.name() == "SyntheticNamedCell");
    if name_found {
        book.sheet_by_name_mut(&sheet_name)?
            .cell_mut("E5")
            .set_value("umya-probe named cell");
        item(
            &mut items,
            "defined_name_cell",
            "success",
            "SyntheticNamedCell target E5 value changed",
        );
    } else if known_synthetic {
        item(
            &mut items,
            "defined_name_cell",
            "failed",
            "expected SyntheticNamedCell was absent",
        );
    } else {
        item(
            &mut items,
            "defined_name_cell",
            "not_applicable",
            "SyntheticNamedCell was absent",
        );
    }

    {
        let sheet = book.sheet_by_name_mut(&sheet_name)?;
        if !sheet.conditional_formatting_collection().is_empty() {
            sheet
                .cell_mut("F2")
                .set_value("umya-probe conditional formatting cell");
            item(
                &mut items,
                "conditional_formatting_region",
                "success",
                "F2 value changed; conditional formatting collection is readable",
            );
        } else if known_synthetic {
            item(
                &mut items,
                "conditional_formatting_region",
                "failed",
                "expected conditional formatting was absent",
            );
        } else {
            item(
                &mut items,
                "conditional_formatting_region",
                "not_applicable",
                "worksheet has no conditional formatting",
            );
        }

        let dv_present = sheet
            .data_validations()
            .map(|d| !d.data_validation_list().is_empty())
            .unwrap_or(false);
        if dv_present {
            sheet.cell_mut("G2").set_value("umya-probe validation cell");
            item(
                &mut items,
                "data_validation_region",
                "success",
                "G2 value changed; validation collection is readable",
            );
        } else if known_synthetic {
            item(
                &mut items,
                "data_validation_region",
                "failed",
                "expected data validation was absent",
            );
        } else {
            item(
                &mut items,
                "data_validation_region",
                "not_applicable",
                "worksheet has no data validation",
            );
        }

        let comment = sheet
            .comments()
            .iter()
            .any(|c| c.coordinate().get_coordinate() == "H2");
        if comment {
            sheet.cell_mut("H2").set_value("umya-probe comment cell");
            item(
                &mut items,
                "comment_cell",
                "success",
                "H2 changed; existing comment object left untouched",
            );
        } else if known_synthetic {
            item(
                &mut items,
                "comment_cell",
                "failed",
                "expected comment at H2 was absent",
            );
        } else {
            item(
                &mut items,
                "comment_cell",
                "not_applicable",
                "worksheet has no comment at H2",
            );
        }

        sheet.row_dimension_mut(3).set_height(30.0);
        sheet.column_dimension_mut("I").set_width(25.0);
        sheet.row_dimension_mut(2).set_hidden(true);
        item(
            &mut items,
            "row_height_column_width_hidden_row",
            "success",
            "row 3 height 30; column I width 25; row 2 hidden",
        );
    }

    let print_area_exists = book
        .sheet_by_name(&sheet_name)?
        .defined_names()
        .iter()
        .any(|n| n.name() == "_xlnm.Print_Area")
        || book
            .defined_names()
            .iter()
            .any(|n| n.name() == "_xlnm.Print_Area");
    if print_area_exists {
        if let Some(name) = book
            .sheet_by_name_mut(&sheet_name)?
            .defined_names_mut()
            .iter_mut()
            .find(|n| n.name() == "_xlnm.Print_Area")
        {
            name.set_address(format!("'{}'!$A$1:$D$10", sheet_name.replace('\'', "''")));
            name.set_local_sheet_id(0);
        } else if let Some(name) = book
            .defined_names_mut()
            .iter_mut()
            .find(|n| n.name() == "_xlnm.Print_Area")
        {
            name.set_address(format!("'{}'!$A$1:$D$10", sheet_name.replace('\'', "''")));
            name.set_local_sheet_id(0);
        }
        item(
            &mut items,
            "print_area",
            "success",
            "_xlnm.Print_Area set to A1:D10",
        );
    } else if known_synthetic {
        item(
            &mut items,
            "print_area",
            "failed",
            "expected _xlnm.Print_Area was absent",
        );
    } else {
        item(
            &mut items,
            "print_area",
            "not_applicable",
            "worksheet did not have a print-area defined name",
        );
    }

    excel_editor::save_xlsx(&book, &output)
        .with_context(|| format!("write {}", output.display()))?;
    let umya_library_elapsed_ms = started.elapsed().as_millis();
    let patch_started = std::time::Instant::now();
    let update = ColumnWidthUpdate::new(sheet_name.clone(), 9, 25.0)?;
    let column_patch_status = match patch_column_dimensions(
        &input,
        &output,
        &patched_output,
        &[update],
    ) {
        Ok(_) => {
            item(&mut items, "targeted_ooxml_column_restore", "success", "restored original SpreadsheetML <cols> definitions and applied only the requested I=25 width");
            "success".to_owned()
        }
        Err(err) => {
            item(
                &mut items,
                "targeted_ooxml_column_restore",
                "failed",
                err.to_string(),
            );
            "failed".to_owned()
        }
    };
    let ooxml_patch_elapsed_ms = patch_started.elapsed().as_millis();
    let report = Report {
        schema_version: 1,
        library: "umya-spreadsheet",
        library_version: "3.1.1",
        input: input.display().to_string(),
        output: output.display().to_string(),
        patched_output: (column_patch_status == "success")
            .then(|| patched_output.display().to_string()),
        column_patch_status,
        elapsed_ms: started.elapsed().as_millis(),
        umya_library_elapsed_ms,
        ooxml_patch_elapsed_ms,
        items,
    };
    fs::write(&args.report, serde_json::to_vec_pretty(&report)?)?;
    println!("{}", args.report.display());
    Ok(())
}
