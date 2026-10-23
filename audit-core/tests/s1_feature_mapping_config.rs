use audit_core::name_config::load_s1_conditional_mappings;
use excel_writer::Workbook;
use std::{
    fs,
    path::{Path, PathBuf},
    time::{SystemTime, UNIX_EPOCH},
};

fn temporary_workbook_path(label: &str) -> PathBuf {
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system clock after Unix epoch")
        .as_nanos();
    std::env::temp_dir().join(format!(
        "open_data_audit_{label}_{}_{}.xlsx",
        std::process::id(),
        nonce
    ))
}

fn write_mapping_config(rows: &[Vec<&str>], sheet_name: &str) -> PathBuf {
    let path = temporary_workbook_path("s1_mapping_config");
    let mut workbook = Workbook::new();
    {
        let worksheet = workbook.add_worksheet();
        worksheet
            .set_name(sheet_name)
            .expect("set synthetic configuration sheet name");
        for (row_index, row) in rows.iter().enumerate() {
            for (column_index, value) in row.iter().enumerate() {
                if !value.is_empty() {
                    worksheet
                        .write_string(row_index as u32, column_index as u16, *value)
                        .expect("write synthetic configuration value");
                }
            }
        }
    }
    workbook.save(&path).expect("save synthetic configuration");
    path
}

fn feature_rows() -> Vec<Vec<&'static str>> {
    vec![
        vec![
            "功能名",
            "执行模块",
            "命名区域名",
            "输入",
            "输出",
            "备注",
            "是否限定工作簿",
            "工作簿关键字",
        ],
        vec![
            "条件格式结果提取",
            "汇总_条件格式结果提取",
            "区域一、区域二",
            "源目录",
            "通用结果",
            "通用映射",
            "否",
            "",
        ],
        vec![
            "条件格式结果提取",
            "汇总_条件格式结果提取",
            "限定区域",
            "模板来源",
            "限定结果",
            "限定映射",
            "是",
            "模板甲",
        ],
        vec![
            "条件格式结果提取",
            "汇总_条件格式结果提取",
            "后置通用区域",
            "后置来源",
            "后置结果",
            "后置通用映射",
            "否",
            "",
        ],
        vec![
            "表结构比对",
            "核对_表结构比对",
            "表结构区域",
            "源目录、模板表结构区域",
            "运行日志",
            "结构映射",
            "否",
            "",
        ],
        vec![
            "忽略不支持模块",
            "未登记的模块",
            "忽略区域",
            "",
            "",
            "",
            "否",
            "",
        ],
        vec![
            "忽略其他模板",
            "汇总_条件格式结果提取",
            "错误限定区域",
            "",
            "",
            "",
            "是",
            "其他模板",
        ],
    ]
}

#[test]
fn loads_conditional_and_structure_mappings_with_template_priority() {
    let config = write_mapping_config(&feature_rows(), "模块化功能");
    let template = Path::new("机构模板甲_20260731.xlsx");

    let mappings = load_s1_conditional_mappings(Some(&config), template)
        .expect("load S1 feature mappings from Rust reader");
    fs::remove_file(config).expect("remove temporary configuration");

    assert_eq!(mappings.len(), 2);
    assert_eq!(mappings[0].name, "条件格式结果提取");
    assert_eq!(mappings[0].feature_type, "汇总_条件格式结果提取");
    assert_eq!(mappings[0].range_names, vec!["限定区域".to_owned()]);
    assert!(mappings[0].workbook_limited);
    assert_eq!(mappings[0].workbook_keyword, "模板甲");
    // The source loader constructs non-default mappings with six fields; its
    // dataclass defaults input/output metadata to empty strings.
    assert!(mappings[0].input_note.is_empty());
    assert!(mappings[0].output.is_empty());
    assert_eq!(mappings[0].remark, "限定映射");
    assert_eq!(mappings[1].name, "表结构比对");
    assert_eq!(mappings[1].range_names, vec!["表结构区域".to_owned()]);
}

#[test]
fn missing_or_empty_mapping_sheet_uses_s1_defaults() {
    let template = Path::new("default-template.xlsx");
    let absent = load_s1_conditional_mappings(None, template).expect("load default mappings");
    assert_eq!(absent.len(), 2);
    assert_eq!(absent[0].range_names, vec!["条件格式区域".to_owned()]);
    assert_eq!(absent[1].range_names, vec!["表结构区域".to_owned()]);

    let rows = vec![
        vec!["功能名", "执行模块", "命名区域名", "输入", "输出", "备注"],
        vec!["不适用项", "未知模块", "其他区域", "", "", ""],
    ];
    let missing_sheet_config = write_mapping_config(&rows, "其他工作表");
    let missing_sheet = load_s1_conditional_mappings(Some(&missing_sheet_config), template)
        .expect("missing module sheet falls back to defaults");
    fs::remove_file(missing_sheet_config).expect("remove temporary configuration");
    assert_eq!(
        missing_sheet[0].range_names,
        vec!["条件格式区域".to_owned()]
    );
    assert_eq!(missing_sheet[1].range_names, vec!["表结构区域".to_owned()]);

    let empty_applicable_config = write_mapping_config(&rows, "模块化功能");
    let empty_applicable = load_s1_conditional_mappings(Some(&empty_applicable_config), template)
        .expect("empty applicable mapping table falls back to defaults");
    fs::remove_file(empty_applicable_config).expect("remove temporary configuration");
    assert_eq!(
        empty_applicable[0].range_names,
        vec!["条件格式区域".to_owned()]
    );
    assert_eq!(
        empty_applicable[1].range_names,
        vec!["表结构区域".to_owned()]
    );
}

#[test]
fn mapping_sheet_missing_required_column_fails_explicitly() {
    let rows = vec![vec!["功能名", "执行模块", "命名区域名", "输入", "输出"]];
    let config = write_mapping_config(&rows, "模块化功能");
    let result = load_s1_conditional_mappings(Some(&config), Path::new("template.xlsx"));
    fs::remove_file(config).expect("remove temporary configuration");

    let error = result.expect_err("configuration schema errors must be explicit");
    assert!(format!("{error:#}").contains("缺少列：备注"));
}
