//! Manual S5 layout-region parity runner.
//!
//! Loads the workbook identified by `AUDIT_S5_LAYOUT_INPUT` and writes the
//! Rust region discovery results (`AUDIT_S5_LAYOUT_REPORT`). Manual parity
//! harness only.

use audit_core::s5_layout;
use serde::Deserialize;
use serde_json::json;
use std::{fs, path::PathBuf};

#[derive(Deserialize)]
struct LayoutRequest {
    workbook: PathBuf,
}

fn region_json(region: &s5_layout::NamedRegion) -> serde_json::Value {
    json!({
        "name": region.name,
        "sheet": region.sheet,
        "reference": region.reference,
        "min_row": region.min_row,
        "min_col": region.min_col,
        "max_row": region.max_row,
        "max_col": region.max_col,
    })
}

#[test]
#[ignore = "manual parity runner; requires AUDIT_S5_LAYOUT_INPUT/REPORT"]
fn runs_s5_layout_parity_item() {
    let input_path: PathBuf = std::env::var_os("AUDIT_S5_LAYOUT_INPUT")
        .expect("AUDIT_S5_LAYOUT_INPUT is missing")
        .into();
    let report_path: PathBuf = std::env::var_os("AUDIT_S5_LAYOUT_REPORT")
        .expect("AUDIT_S5_LAYOUT_REPORT is missing")
        .into();

    let request: LayoutRequest =
        serde_json::from_str(&fs::read_to_string(&input_path).expect("read layout request"))
            .expect("parse layout request");
    let (numeric, origins) = s5_layout::numeric_regions(&request.workbook);
    let (inferred, missing) = s5_layout::inferred_numeric_regions(&request.workbook);
    let structure = s5_layout::structure_regions(&request.workbook);
    let payload = json!({
        "named_structure_axes": s5_layout::named_regions(
            &request.workbook,
            &s5_layout::STRUCTURE_PREFIXES,
        )
        .iter()
        .map(region_json)
        .collect::<Vec<_>>(),
        "numeric_regions": numeric.iter().map(region_json).collect::<Vec<_>>(),
        "numeric_origins": origins,
        "inferred_results": inferred.iter().map(region_json).collect::<Vec<_>>(),
        "inferred_missing": missing.iter().map(region_json).collect::<Vec<_>>(),
        "structure_by_sheet": structure
            .iter()
            .map(|(sheet, regions)| {
                (
                    sheet.clone(),
                    serde_json::Value::Array(
                        regions.iter().map(region_json).collect::<Vec<_>>(),
                    ),
                )
            })
            .collect::<serde_json::Map<String, serde_json::Value>>(),
    });
    fs::write(
        &report_path,
        serde_json::to_vec_pretty(&payload).expect("serialize layout payload"),
    )
    .expect("write layout payload");
    println!("AUDIT_S5_LAYOUT_STATUS=ok");
}
