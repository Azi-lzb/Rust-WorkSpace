use anyhow::Result;
use pytools_core::batch_ops::{build_quarterly_rural_loan_plan, build_rural_loan_plan};
use rust_xlsxwriter::Workbook;
use std::{fs, path::PathBuf};

struct TestDir(PathBuf);

impl TestDir {
    fn new() -> Self {
        let path = std::env::temp_dir().join(format!(
            "pytools-rural-loan-regressions-{}-{}",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
        fs::create_dir_all(&path).unwrap();
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

#[derive(Clone)]
enum Cell {
    Text(&'static str),
    Number(f64),
}

fn write_book(path: &std::path::Path, sheets: Vec<(&str, Vec<Vec<Cell>>)>) -> Result<()> {
    let mut workbook = Workbook::new();
    for (name, rows) in sheets {
        let worksheet = workbook.add_worksheet();
        worksheet.set_name(name)?;
        for (row, values) in rows.iter().enumerate() {
            for (column, value) in values.iter().enumerate() {
                match value {
                    Cell::Text(text) => {
                        worksheet.write_string(row as u32, column as u16, *text)?;
                    }
                    Cell::Number(number) => {
                        worksheet.write_number(row as u32, column as u16, *number)?;
                    }
                }
            }
        }
    }
    workbook.save(path)?;
    Ok(())
}

fn institution_sheet(entries: &[(&'static str, f64, f64)]) -> Vec<Vec<Cell>> {
    let mut rows = vec![
        vec![
            Cell::Text("序号"),
            Cell::Text("机构名称"),
            Cell::Text("涉农贷款余额"),
            Cell::Text("旧值"),
        ],
        vec![
            Cell::Text("序号"),
            Cell::Text("机构名称"),
            Cell::Text("涉农贷款余额"),
            Cell::Text("旧值"),
        ],
    ];
    rows.extend(
        entries
            .iter()
            .enumerate()
            .map(|(index, (name, value, old))| {
                vec![
                    Cell::Number((index + 1) as f64),
                    Cell::Text(name),
                    Cell::Number(*value),
                    Cell::Number(*old),
                ]
            }),
    );
    rows
}

#[test]
fn rural_loan_plan_preserves_unrounded_difference() -> Result<()> {
    let dir = TestDir::new();
    let current = dir.path("current.xlsx");
    let previous = dir.path("previous.xlsx");
    write_book(
        &current,
        vec![
            (
                "惠州市涉农贷款分机构",
                institution_sheet(&[("甲行", 14.123456789, 999.0)]),
            ),
            (
                "本外币涉农贷款分地区",
                vec![
                    vec![Cell::Text("地区"), Cell::Text("涉农贷款比上月")],
                    vec![Cell::Text("惠州市"), Cell::Number(5.123456789)],
                ],
            ),
        ],
    )?;
    write_book(
        &previous,
        vec![(
            "惠州市涉农贷款分机构",
            institution_sheet(&[("甲行", 9.0, 0.0)]),
        )],
    )?;

    let (report, summary) =
        build_rural_loan_plan(&current, &previous, &dir.path("candidate.xlsx"))?;
    let plan = report.plan.expect("rural loan plan");
    let value = plan
        .operations
        .iter()
        .find_map(|operation| match operation {
            excel_editor::EditOperation::SetCellValue {
                cell,
                value: excel_editor::CellValue::Decimal(value),
                ..
            } if cell == "D3" => Some(value),
            _ => None,
        })
        .expect("D3 difference operation");
    let actual = value.parse::<f64>().expect("decimal value");
    assert!((actual - 5.123456789).abs() <= 1e-15);
    assert!((summary.fill_total - actual).abs() <= 1e-15);
    Ok(())
}

#[test]
fn rural_loan_reports_and_plan_follow_institution_row_order() -> Result<()> {
    let dir = TestDir::new();
    let current = dir.path("current.xlsx");
    let previous = dir.path("previous.xlsx");
    write_book(
        &current,
        vec![
            (
                "惠州市涉农贷款分机构",
                institution_sheet(&[
                    ("机构乙", 15.0, 999.0),
                    ("未匹配甲", 4.0, 998.0),
                    ("机构甲", 11.0, 997.0),
                    ("未匹配乙", 5.0, 996.0),
                ]),
            ),
            (
                "本外币涉农贷款分地区",
                vec![
                    vec![Cell::Text("地区"), Cell::Text("涉农贷款比上月")],
                    vec![Cell::Text("惠州市"), Cell::Number(15.0)],
                ],
            ),
        ],
    )?;
    write_book(
        &previous,
        vec![(
            "惠州市涉农贷款分机构",
            institution_sheet(&[
                ("机构乙", 10.0, 0.0),
                ("上期乙", 3.0, 0.0),
                ("机构甲", 1.0, 0.0),
                ("上期甲", 5.0, 0.0),
            ]),
        )],
    )?;

    let (report, summary) =
        build_rural_loan_plan(&current, &previous, &dir.path("candidate.xlsx"))?;
    assert_eq!(summary.unmatched_current, ["未匹配甲", "未匹配乙"]);
    assert_eq!(summary.previous_only, ["上期乙", "上期甲"]);
    let plan = report.plan.expect("rural-loan plan missing");
    let edited_cells = plan
        .operations
        .iter()
        .filter_map(|operation| match operation {
            excel_editor::EditOperation::SetCellValue { cell, .. } => Some(cell.as_str()),
            _ => None,
        })
        .collect::<Vec<_>>();
    assert_eq!(edited_cells, ["D3", "D5"]);
    Ok(())
}

#[test]
fn quarterly_summary_keeps_invalid_regions_and_fails_overall_check() -> Result<()> {
    let dir = TestDir::new();
    let current = dir.path("current.xlsx");
    let previous = dir.path("previous.xlsx");
    write_book(
        &current,
        vec![
            (
                "惠州市涉农贷款分机构",
                institution_sheet(&[("甲行", 14.123456789, 999.0)]),
            ),
            (
                "江门市涉农贷款分机构",
                institution_sheet(&[("江门行", 8.0, 777.0)]),
            ),
            (
                "珠海市涉农贷款分机构",
                institution_sheet(&[("重复行", 4.0, 777.0), ("重复行", 5.0, 888.0)]),
            ),
            (
                "深圳市涉农贷款分机构",
                institution_sheet(&[("深行", 5.0, 999.0)]),
            ),
            (
                "涉农贷款分地区统计表",
                vec![
                    vec![Cell::Text("地区"), Cell::Text("涉农贷款比上月")],
                    vec![Cell::Text("惠州市"), Cell::Number(5.123456789)],
                    vec![Cell::Text("江门市"), Cell::Number(8.0)],
                    vec![Cell::Text("珠海市"), Cell::Number(0.0)],
                ],
            ),
        ],
    )?;
    write_book(
        &previous,
        vec![
            (
                "惠州市涉农贷款分机构",
                institution_sheet(&[("甲行", 9.0, 888.0)]),
            ),
            (
                "珠海市涉农贷款分机构",
                institution_sheet(&[("重复行", 3.0, 999.0)]),
            ),
            (
                "深圳市涉农贷款分机构",
                institution_sheet(&[("深行", 3.0, 999.0)]),
            ),
        ],
    )?;

    let (report, summary) =
        build_quarterly_rural_loan_plan(&current, &previous, &dir.path("candidate.xlsx"))?;
    assert_eq!(summary.regions.len(), 4);
    assert!(!summary.consistent);
    assert_eq!(summary.error, None);
    assert_eq!(summary.regions[0].error, None);
    assert_eq!(
        summary.regions[1].error.as_deref(),
        Some("上期文件缺少江门市分机构表")
    );
    assert_eq!(
        summary.regions[2].error.as_deref(),
        Some("分机构表存在重复机构名称：重复行")
    );
    assert_eq!(
        summary.regions[3].error.as_deref(),
        Some("“涉农贷款分地区统计表”未找到地区行：深圳市")
    );
    assert_eq!(report.changed_cells, 2);
    assert_eq!(report.plan.as_ref().unwrap().operations.len(), 2);
    Ok(())
}

#[test]
fn quarterly_rural_loan_reports_follow_institution_row_order() -> Result<()> {
    let dir = TestDir::new();
    let current = dir.path("current.xlsx");
    let previous = dir.path("previous.xlsx");
    write_book(
        &current,
        vec![
            (
                "江门市涉农贷款分机构",
                institution_sheet(&[
                    ("机构乙", 15.0, 999.0),
                    ("未匹配甲", 4.0, 998.0),
                    ("机构甲", 11.0, 997.0),
                    ("未匹配乙", 5.0, 996.0),
                ]),
            ),
            (
                "涉农贷款分地区统计表",
                vec![
                    vec![Cell::Text("地区"), Cell::Text("涉农贷款比上月")],
                    vec![Cell::Text("江门市"), Cell::Number(15.0)],
                ],
            ),
        ],
    )?;
    write_book(
        &previous,
        vec![(
            "江门市涉农贷款分机构",
            institution_sheet(&[
                ("机构乙", 10.0, 0.0),
                ("上期乙", 3.0, 0.0),
                ("机构甲", 1.0, 0.0),
                ("上期甲", 5.0, 0.0),
            ]),
        )],
    )?;

    let (report, summary) =
        build_quarterly_rural_loan_plan(&current, &previous, &dir.path("candidate.xlsx"))?;
    let region = &summary.regions[0];
    assert_eq!(region.unmatched_current, ["未匹配甲", "未匹配乙"]);
    assert_eq!(region.previous_only, ["上期乙", "上期甲"]);
    assert_eq!(report.changed_cells, 2);
    let plan = report.plan.expect("quarterly rural-loan plan missing");
    let edited_cells = plan
        .operations
        .iter()
        .filter_map(|operation| match operation {
            excel_editor::EditOperation::SetCellValue { cell, .. } => Some(cell.as_str()),
            _ => None,
        })
        .collect::<Vec<_>>();
    assert_eq!(edited_cells, ["D3", "D5"]);
    Ok(())
}
