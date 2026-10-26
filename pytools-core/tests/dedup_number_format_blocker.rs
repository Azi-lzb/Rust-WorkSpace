use anyhow::Result;
use pytools_core::{
    config::DedupTask,
    dedup_ops::{
        build_config_duplicate_plan, write_config_duplicate_delete_candidate, DuplicateAction,
    },
};
use rust_xlsxwriter::{Format, Workbook};
use std::{
    collections::BTreeMap,
    fs,
    path::PathBuf,
    time::{SystemTime, UNIX_EPOCH},
};

struct TestDir(PathBuf);

impl TestDir {
    fn new() -> Self {
        let path = std::env::temp_dir().join(format!(
            "pytools-number-format-repro-{}-{}",
            std::process::id(),
            SystemTime::now()
                .duration_since(UNIX_EPOCH)
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

#[test]
fn configured_delete_preserves_numeric_zero_padding_format() -> Result<()> {
    let dir = TestDir::new();
    let input = dir.path("input.xlsx");
    let output = dir.path("output.xlsx");

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_name("Data")?;
    let code_format = Format::new().set_num_format("000000");
    worksheet.write_string(0, 0, "机构代码")?;
    worksheet.write_string(0, 1, "复合键")?;
    worksheet.write_number_with_format(1, 0, 123.0, &code_format)?;
    worksheet.write_string(1, 1, "同键")?;
    worksheet.write_number_with_format(2, 0, 456.0, &code_format)?;
    worksheet.write_string(2, 1, "同键")?;
    worksheet.write_number_with_format(3, 0, 789.0, &code_format)?;
    worksheet.write_string(3, 1, "唯一键")?;
    workbook.save(&input)?;
    let input_bytes = fs::read(&input)?;

    let read_format = |path: &PathBuf, address: &str| -> Result<Option<String>> {
        let book = excel_editor::open_xlsx(path)?;
        let format = book
            .sheet_by_name("Data")?
            .cell(address)
            .and_then(|cell| cell.style().number_format())
            .map(|format| format.format_code().to_owned());
        Ok(format)
    };
    assert_eq!(read_format(&input, "A2")?.as_deref(), Some("000000"));
    assert_eq!(read_format(&input, "A4")?.as_deref(), Some("000000"));

    let task = DedupTask {
        row_no: 2,
        enabled: true,
        task_name: "synthetic".into(),
        source_wb: Some(input.clone()),
        source_ws: "Data".into(),
        key_cols: vec![2],
        target_wb: None,
        target_ws: String::new(),
        write_mode: String::new(),
    };
    let outputs = BTreeMap::from([(input.clone(), output.clone())]);
    let reports = build_config_duplicate_plan(&[task], &outputs, DuplicateAction::Delete)?;
    let plan = reports[0].plan.as_ref();
    write_config_duplicate_delete_candidate(&input, &output, plan)?;

    assert_eq!(read_format(&output, "A2")?.as_deref(), Some("000000"));
    assert_eq!(read_format(&output, "A3")?.as_deref(), Some("000000"));
    assert_eq!(fs::read(&input)?, input_bytes);
    Ok(())
}
