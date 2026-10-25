use excel_reader::{read_workbook, CellValue};
use serde_json::Value;
use std::{
    fs,
    path::{Path, PathBuf},
    process::{Command, Output},
    sync::atomic::{AtomicU64, Ordering},
    time::{SystemTime, UNIX_EPOCH},
};

static TEMPORARY_DIRECTORY_SEQUENCE: AtomicU64 = AtomicU64::new(0);

fn temporary_directory() -> PathBuf {
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system clock after Unix epoch")
        .as_nanos();
    let sequence = TEMPORARY_DIRECTORY_SEQUENCE.fetch_add(1, Ordering::Relaxed);
    let path = std::env::temp_dir().join(format!(
        "audit_s5_previous_source_cli_{}_{}_{}",
        std::process::id(),
        nonce,
        sequence
    ));
    fs::create_dir_all(&path).expect("create temporary S5 CLI directory");
    path
}

fn write_source(path: &Path, value: f64) {
    let mut workbook = excel_writer::Workbook::new();
    let info = workbook
        .add_worksheet()
        .set_name("基本信息")
        .expect("name source identity sheet");
    info.write_string(0, 0, "机构代码")
        .expect("write institution label");
    info.write_string(0, 1, "I001")
        .expect("write institution code");
    info.write_string(1, 0, "统计期间")
        .expect("write period label");

    let data = workbook
        .add_worksheet()
        .set_name("贷款表")
        .expect("name source data sheet");
    data.write_string(0, 0, "模拟单位贷款数据")
        .expect("write title");
    data.write_string(1, 2, "余额").expect("write metric group");
    data.write_string(2, 0, "一级")
        .expect("write first row header");
    data.write_string(2, 1, "二级")
        .expect("write second row header");
    data.write_string(2, 2, "本期")
        .expect("write current column header");
    data.write_string(2, 3, "上期")
        .expect("write previous column header");
    data.write_string(3, 0, "贷款产品类别")
        .expect("write row path");
    data.write_string(3, 1, "经营贷款")
        .expect("write row path detail");
    data.write_number(3, 2, value).expect("write source amount");
    data.write_number(3, 3, 100.0)
        .expect("write prior-column control amount");
    workbook.save(path).expect("save source workbook");
}

fn write_numeric_scope_template(path: &Path) {
    let mut workbook = excel_writer::Workbook::new();
    let sheet = workbook
        .add_worksheet()
        .set_name("贷款表")
        .expect("name numeric-scope template sheet");
    for (row, column, value) in [
        (1, 2, "余额"),
        (1, 3, "余额"),
        (2, 2, "本期"),
        (2, 3, "上期"),
        (3, 0, "贷款产品类别"),
        (3, 1, "经营贷款"),
    ] {
        sheet
            .write_string(row, column, value)
            .expect("write numeric-scope template labels");
    }
    sheet
        .write_number(3, 2, 120.0)
        .expect("write template current value");
    sheet
        .write_number(3, 3, 100.0)
        .expect("write template previous value");
    workbook
        .define_name("表结构区域_1", "='贷款表'!$A$4:$B$4")
        .expect("define template row axis");
    workbook
        .define_name("列索引区域_2", "='贷款表'!$C$2:$D$3")
        .expect("define template column axis");
    workbook
        .define_name("数字区域_1", "='贷款表'!$C$4:$D$4")
        .expect("define template numeric region");
    workbook.save(path).expect("save numeric-scope template");
}

fn write_numeric_scope_source(path: &Path, current_value: f64) {
    let mut workbook = excel_writer::Workbook::new();
    let sheet = workbook
        .add_worksheet()
        .set_name("贷款表")
        .expect("name numeric-scope source sheet");
    for (row, column, value) in [
        (3, 3, "余额"),
        (3, 4, "余额"),
        (3, 6, "余额"),
        (4, 3, "本期"),
        (4, 4, "上期"),
        (4, 6, "本期"),
        (5, 1, "贷款产品类别"),
        (5, 2, "经营贷款"),
        (7, 1, "贷款产品类别"),
        (7, 2, "经营贷款"),
    ] {
        sheet
            .write_string(row, column, value)
            .expect("write numeric-scope source labels");
    }
    sheet
        .write_number(5, 3, current_value)
        .expect("write source current value");
    sheet
        .write_number(5, 4, 100.0)
        .expect("write source previous value");
    sheet
        .write_number(5, 6, 999.0)
        .expect("write decoy numeric value outside template region");
    sheet
        .write_number(7, 3, 777.0)
        .expect("write decoy duplicate row-path value");
    workbook.save(path).expect("save numeric-scope source");
}

fn run_numeric_scope_cli(
    config: &Path,
    current: &Path,
    previous: &Path,
    template_path: Option<&Path>,
    template_dir: Option<&Path>,
    base_data_dir: Option<&Path>,
    output_dir: &Path,
) -> Output {
    let mut command = Command::new(env!("CARGO_BIN_EXE_open-data-audit"));
    command
        .arg("s5-audit-run")
        .arg("--config")
        .arg(config)
        .arg("--source")
        .arg(current)
        .arg("--source-id")
        .arg("CUR")
        .arg("--previous-source")
        .arg(previous)
        .arg("--previous-period")
        .arg("2025-06-30")
        .arg("--period")
        .arg("2026-08-31")
        .arg("--org-code")
        .arg("I001")
        .arg("--org-name")
        .arg("合成机构")
        .arg("--business")
        .arg("单位贷款")
        .arg("--output-dir")
        .arg(output_dir)
        .arg("--run-id")
        .arg("numeric-scope")
        .arg("--developer-view");
    if let Some(template_path) = template_path {
        command.arg("--template-path").arg(template_path);
    }
    if let Some(template_dir) = template_dir {
        command.arg("--template-dir").arg(template_dir);
    }
    if let Some(base_data_dir) = base_data_dir {
        command.env("BASE_AUDIT_DATA_DIR", base_data_dir);
    }
    command.output().expect("run numeric-scope S5 audit CLI")
}

#[test]
fn s5_scan_stringifies_numeric_boolean_and_formula_labels() {
    let directory = temporary_directory();
    let path = directory.join("labels.xlsx");
    let mut workbook = excel_writer::Workbook::new();
    let sheet = workbook
        .add_worksheet()
        .set_name("路径")
        .expect("name path labels sheet");
    sheet.write_number(0, 0, 1.5).expect("write numeric label");
    sheet
        .write_boolean(0, 1, true)
        .expect("write boolean label");
    sheet
        .write_formula(0, 2, "=1+1")
        .expect("write formula label");
    sheet
        .write_number(1, 0, 2.0)
        .expect("write integer numeric label");
    workbook.save(&path).expect("save path labels workbook");

    let book = audit_core::s5_scan::ScanBook::open(&path).expect("read path labels workbook");
    assert_eq!(
        book.row_sequence("路径", 1, 1, 3),
        vec![
            ("1.5".to_owned(), 0),
            ("True".to_owned(), 1),
            ("=1+1".to_owned(), 2),
        ]
    );
    assert_eq!(
        book.row_sequence("路径", 2, 1, 1),
        vec![("2".to_owned(), 0)]
    );
    assert_eq!(book.cell_raw_text("路径", 2, 1), "2");

    fs::remove_dir_all(directory).expect("remove path labels fixture directory");
}

fn write_external_source(path: &Path, value: f64) {
    write_external_source_variant(path, value, false, false, false, "机构代码");
}

fn write_external_source_with_python_code_alias(path: &Path, value: f64) {
    write_external_source_variant(path, value, false, false, false, "社会信用代码");
}

fn write_external_source_with_layout_warning(path: &Path, value: f64, warning: bool) {
    write_external_source_variant(path, value, warning, false, false, "机构代码");
}

fn write_external_source_with_invalid_numeric(path: &Path) {
    write_external_source_variant(path, 0.0, false, true, false, "机构代码");
}

fn write_external_time_series_source(path: &Path, value: f64) {
    write_external_time_series_source_variant(path, value, true);
}

fn write_external_time_series_source_without_current_date(path: &Path, value: f64) {
    write_external_time_series_source_variant(path, value, false);
}

fn write_external_time_series_source_variant(path: &Path, value: f64, include_current_date: bool) {
    let mut workbook = excel_writer::Workbook::new();
    let mapping = workbook
        .add_worksheet()
        .set_name("机构参照表")
        .expect("name external mapping sheet");
    for (column, header) in ["机构代码", "报表项目"].iter().enumerate() {
        mapping
            .write_string(0, column as u16, *header)
            .expect("write mapping header");
    }
    mapping
        .write_string(1, 0, "I001")
        .expect("write mapping identity");
    mapping
        .write_string(1, 1, "单位贷款")
        .expect("write mapping report item");

    let data = workbook
        .add_worksheet()
        .set_name("集中系统数据")
        .expect("name external data sheet");
    for (column, header) in ["数据日期", "行名称", "贷款余额(亿元)", "备用指标"]
        .iter()
        .enumerate()
    {
        data.write_string(1, column as u16, *header)
            .expect("write time-series header");
    }
    data.write_string(2, 0, "2026-07-31")
        .expect("write prior report date");
    data.write_string(2, 1, "单位贷款")
        .expect("write prior report item");
    data.write_number(2, 2, value / 2.0)
        .expect("write prior metric value");
    if include_current_date {
        let current_date = excel_writer::ExcelDateTime::parse_from_str("2026-08-31")
            .expect("parse current report date");
        let date_format = excel_writer::Format::new().set_num_format("yyyy-mm-dd");
        data.write_datetime_with_format(3, 0, current_date, &date_format)
            .expect("write Excel date-formatted current report date");
        data.write_string(3, 1, "单位贷款")
            .expect("write current report item");
        data.write_number(3, 2, value)
            .expect("write current metric value");
    }
    let invalid_row = if include_current_date { 4 } else { 3 };
    data.write_string(invalid_row, 0, "not-a-date")
        .expect("write invalid report date");
    data.write_string(invalid_row, 1, "单位贷款")
        .expect("write invalid report item");
    data.write_number(invalid_row, 2, value * 2.0)
        .expect("write invalid-date metric value");
    workbook
        .save(path)
        .expect("save time-series source workbook");
}

fn write_external_source_variant(
    path: &Path,
    value: f64,
    warning: bool,
    invalid_numeric: bool,
    time_series: bool,
    mapping_code_header: &str,
) {
    let mut workbook = excel_writer::Workbook::new();
    let mapping = workbook
        .add_worksheet()
        .set_name("机构参照表")
        .expect("name external mapping sheet");
    mapping
        .write_string(0, 0, mapping_code_header)
        .expect("write mapping identity header");
    mapping
        .write_string(0, 1, "报表项目")
        .expect("write mapping report item header");
    mapping
        .write_string(1, 0, "I001")
        .expect("write mapping identity");
    mapping
        .write_string(1, 1, " 单位贷款 ")
        .expect("write mapping report item");

    let data = workbook
        .add_worksheet()
        .set_name("集中系统数据")
        .expect("name external data sheet");
    let headers = if time_series {
        ["数据日期", "行名称", "贷款余额(亿元)", "备用指标"]
    } else {
        ["序号", "报表项目", " 贷款余额(亿元) ", "备用指标"]
    };
    for (column, value) in headers.iter().enumerate() {
        data.write_string(1, column as u16, *value)
            .expect("write external metric headers");
    }
    if time_series {
        data.write_string(2, 0, "2026-08-31")
            .expect("write external report date");
    } else {
        data.write_number(2, 0, 1)
            .expect("write external row number");
    }
    data.write_string(2, 1, " 单位贷款 ")
        .expect("write external report item");
    if invalid_numeric {
        data.write_string(2, 2, "误填文字")
            .expect("write invalid external metric value");
    } else {
        data.write_number(2, 2, value)
            .expect("write external metric value");
    }
    data.write_number(2, 3, 1)
        .expect("write external control value");
    if warning {
        data.write_number(3, 2, 0.01)
            .expect("write unlabelled external metric value");
    }
    workbook.save(path).expect("save external source workbook");
}

fn write_config(path: &Path, include_external: bool, include_external_contract: bool) {
    write_config_variant(path, include_external, include_external_contract, false);
}

fn write_multi_current_source_config(path: &Path) {
    write_config_variant(path, false, false, true);
}

fn write_config_variant(
    path: &Path,
    include_external: bool,
    include_external_contract: bool,
    multi_current_sources: bool,
) {
    let mut workbook = excel_writer::Workbook::new();
    let rule_headers = [
        "规则ID",
        "规则修订",
        "是否启用",
        "业务类型",
        "默认数据源",
        "默认Sheet",
        "规则名称",
        "审核级别",
        "校验类型",
        "校验指标",
        "审核表达式",
        "表达式布尔值取反",
        "规则描述",
        "当前值表达式",
        "对比值表达式",
        "差值表达式",
        "取值字典",
        "生效起始日",
        "生效截止日",
        "适用范围",
        "适用表达式",
    ];
    let rule_values: Vec<&str> = if multi_current_sources {
        vec![
            "MULTI.CURRENT.A",
            "1",
            "是",
            "其他业务",
            "CUR_A",
            "贷款表",
            "本期来源 A",
            "错误",
            "表间核对",
            "经营贷款余额",
            "ABS({本期}) < 0",
            "否",
            "本期值={本期}",
            "{本期}",
            "",
            "",
            r#"{"本期": V_CUR(("贷款产品类别", "经营贷款"), ("余额", "本期"))}"#,
            "",
            "",
            "",
            "",
        ]
    } else if include_external {
        vec![
            "EXTERNAL.PARITY",
            "1",
            "是",
            "单位贷款",
            "CUR",
            "贷款表",
            "外部来源对拍",
            "错误",
            "表间核对",
            "经营贷款余额",
            "ABS({差值}) > 0",
            "否",
            "差异={差值}",
            "{本期}",
            "{大集中}",
            "{差值}",
            r#"{"本期": V_CUR(("贷款产品类别", "经营贷款"), ("余额", "本期")), "大集中": V_EXT("贷款余额(亿元)"), "差值": {本期} - {大集中}}"#,
            "",
            "",
            "",
            "",
        ]
    } else {
        vec![
            "PREVIOUS.PARITY",
            "1",
            "是",
            "单位贷款",
            "CUR",
            "贷款表",
            "上期来源对拍",
            "错误",
            "表间核对",
            "经营贷款余额",
            "ABS({本期} - {上期}) = 20",
            "否",
            "差异={差值}",
            "{本期}",
            "{上期}",
            "{差值}",
            r#"{"本期": V_CUR(("贷款产品类别", "经营贷款"), ("余额", "本期")), "上期": V_PRE(("贷款产品类别", "经营贷款"), ("余额", "本期")), "差值": {本期} - {上期}}"#,
            "",
            "",
            "",
            "",
        ]
    };
    {
        let sheet = workbook
            .add_worksheet()
            .set_name("审核规则")
            .expect("name rules sheet");
        for (column, header) in rule_headers.iter().enumerate() {
            sheet
                .write_string(0, column as u16, *header)
                .expect("write rule header");
        }
        for (column, value) in rule_values.iter().enumerate() {
            sheet
                .write_string(1, column as u16, *value)
                .expect("write rule value");
        }
        if multi_current_sources {
            let second_rule_values = [
                "MULTI.CURRENT.B",
                "1",
                "是",
                "单位贷款",
                "",
                "贷款表",
                "本期来源 B",
                "错误",
                "表间核对",
                "经营贷款余额",
                "ABS({本期}) < 0",
                "否",
                "本期值={本期}",
                "{本期}",
                "",
                "",
                r#"{"本期": V_CUR(("贷款产品类别", "经营贷款"), ("余额", "本期"))}"#,
                "",
                "",
                "",
                "",
            ];
            for (column, value) in second_rule_values.iter().enumerate() {
                sheet
                    .write_string(2, column as u16, *value)
                    .expect("write second current-source rule");
            }
        }
    }

    let metric_headers = [
        "业务类型",
        "工作表名称",
        "行路径",
        "列路径",
        "行表头标题",
        "数据类型",
        "指标类型",
        "来源单位",
        "标准单位",
        "存储尺度",
        "允许空值",
        "精度",
        "口径依据",
        "确认状态",
        "布局版本",
    ];
    let metric_values = [
        "单位贷款",
        "贷款表",
        "贷款产品类别
经营贷款",
        "余额
本期",
        "一级
二级",
        "number",
        "amount",
        "万元",
        "万元",
        "native",
        "否",
        "2",
        "合成对拍口径",
        "是",
        "synthetic-v1",
    ];
    {
        let sheet = workbook
            .add_worksheet()
            .set_name("模板指标目录")
            .expect("name metric catalog");
        for (column, header) in metric_headers.iter().enumerate() {
            sheet
                .write_string(0, column as u16, *header)
                .expect("write metric header");
        }
        for (column, value) in metric_values.iter().enumerate() {
            sheet
                .write_string(1, column as u16, *value)
                .expect("write metric value");
        }
        if multi_current_sources {
            for (column, value) in metric_values.iter().enumerate() {
                sheet
                    .write_string(
                        2,
                        column as u16,
                        if column == 0 { "其他业务" } else { value },
                    )
                    .expect("write metric for primary source business");
            }
        }
    }
    if include_external {
        let external_headers = [
            "数据源ID",
            "指标表头",
            "数据类型",
            "指标类型",
            "来源单位",
            "标准单位",
            "存储尺度",
            "允许空值",
            "精度",
            "口径依据",
            "确认状态",
        ];
        let external_values = [
            "EXT",
            "贷款余额(亿元)",
            "number",
            "amount",
            "亿元",
            "万元",
            "native",
            "否",
            "2",
            "合成对拍口径",
            "是",
        ];
        let sheet = workbook
            .add_worksheet()
            .set_name("外部指标目录")
            .expect("name external metric catalog");
        for (column, header) in external_headers.iter().enumerate() {
            sheet
                .write_string(0, column as u16, *header)
                .expect("write external metric header");
        }
        for (column, value) in external_values.iter().enumerate() {
            sheet
                .write_string(1, column as u16, *value)
                .expect("write external metric value");
        }
    }

    if include_external_contract {
        let contract_headers = [
            "数据源ID",
            "来源类型",
            "默认Sheet",
            "机构Sheet",
            "机构代码标签",
            "机构名称标签",
            "期间标签",
            "映射Sheet",
            "映射表头行",
            "映射数据起始行",
            "映射机构字段",
            "映射目标字段",
            "数据表头行",
            "数据起始行",
            "数据键列",
            "指标表头行",
            "指标起始列",
            "报表项目标签",
            "期望期间",
            "契约版本",
            "期间必须",
            "机构必须",
        ];
        let contract_values = [
            "EXT",
            "external",
            "集中系统数据",
            "",
            "",
            "",
            "",
            "机构参照表",
            "1",
            "2",
            "机构代码",
            "报表项目",
            "2",
            "3",
            "2",
            "2",
            "3",
            "",
            "",
            "synthetic-v1",
            "否",
            "否",
        ];
        let sheet = workbook
            .add_worksheet()
            .set_name("数据源契约")
            .expect("name external source contract");
        for (column, header) in contract_headers.iter().enumerate() {
            sheet
                .write_string(0, column as u16, *header)
                .expect("write source contract header");
        }
        for (column, value) in contract_values.iter().enumerate() {
            sheet
                .write_string(1, column as u16, *value)
                .expect("write source contract value");
        }
    }
    {
        let sheet = workbook
            .add_worksheet()
            .set_name("配置版本")
            .expect("name version sheet");
        for (column, header) in [
            "配置版本号",
            "状态",
            "发布日期",
            "生效起始日",
            "生效截止日",
            "发布说明",
        ]
        .iter()
        .enumerate()
        {
            sheet
                .write_string(0, column as u16, *header)
                .expect("write version header");
        }
        let version_note = if multi_current_sources {
            "Rust/Python multi current-source parity"
        } else {
            "Rust/Python V_PRE parity"
        };
        for (column, value) in ["synthetic-v1", "测试", "2026-10-25", "", "", version_note]
            .iter()
            .enumerate()
        {
            sheet
                .write_string(1, column as u16, *value)
                .expect("write version value");
        }
    }
    workbook.save(path).expect("save S5 config workbook");
}

fn run_cli(
    config: &Path,
    current: &Path,
    previous: Option<&Path>,
    external: Option<&Path>,
    output_dir: &Path,
    developer_view: bool,
) -> Output {
    let empty_template_dir = current
        .parent()
        .expect("current source parent directory")
        .join("empty-s5-layout-templates");
    fs::create_dir_all(&empty_template_dir).expect("create empty S5 layout template directory");
    let mut command = Command::new(env!("CARGO_BIN_EXE_open-data-audit"));
    command
        .arg("s5-audit-run")
        .arg("--config")
        .arg(config)
        .arg("--source")
        .arg(current)
        .arg("--template-dir")
        .arg(empty_template_dir)
        .arg("--source-id")
        .arg("CUR")
        .arg("--period")
        .arg("2026-08-31")
        .arg("--org-code")
        .arg("I001")
        .arg("--org-name")
        .arg("合成机构")
        .arg("--business")
        .arg("单位贷款")
        .arg("--output-dir")
        .arg(output_dir)
        .arg("--run-id")
        .arg(if external.is_some() {
            "with-external"
        } else if previous.is_some() {
            "with-previous"
        } else {
            "without-previous"
        });
    if let Some(previous) = previous {
        command
            .arg("--previous-source")
            .arg(previous)
            .arg("--previous-period")
            .arg("2025-06-30");
    }
    if let Some(external) = external {
        command
            .arg("--external-source")
            .arg(external)
            .arg("--external-source-id")
            .arg("EXT");
    }
    if developer_view {
        command.arg("--developer-view");
    }
    command.output().expect("run S5 audit CLI")
}

fn run_multi_current_cli(
    config: &Path,
    primary: &Path,
    additional_sources: &[(&str, &str, &Path)],
    template_dir: &Path,
    output_dir: &Path,
    run_id: &str,
) -> Output {
    let mut command = Command::new(env!("CARGO_BIN_EXE_open-data-audit"));
    command
        .arg("s5-audit-run")
        .arg("--config")
        .arg(config)
        .arg("--source")
        .arg(primary)
        .arg("--source-id")
        .arg("CUR_A")
        .arg("--source-business")
        .arg("其他业务")
        .arg("--template-dir")
        .arg(template_dir);
    for (source_id, business, path) in additional_sources {
        command
            .arg("--current-source")
            .arg(format!("{source_id}={business}={}", path.display()));
    }
    command
        .arg("--period")
        .arg("2026-08-31")
        .arg("--org-code")
        .arg("I001")
        .arg("--org-name")
        .arg("合成机构")
        .arg("--business")
        .arg("其他业务")
        .arg("--business")
        .arg("单位贷款")
        .arg("--output-dir")
        .arg(output_dir)
        .arg("--run-id")
        .arg(run_id)
        .arg("--developer-view");
    command.output().expect("run S5 multi-current CLI")
}

fn run_multi_source_python_rust_parity(
    python_source_root: &Path,
    config: &Path,
    primary: &Path,
    additional_sources: &[(&str, &str, &Path)],
    template_dir: &Path,
    report_path: &Path,
) -> Value {
    assert!(
        !report_path.exists(),
        "refusing to overwrite parity report: {}",
        report_path.display()
    );
    let comparator = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("compare")
        .join("compare_s5_outputs.py");
    let rust_binary = env!("CARGO_BIN_EXE_open-data-audit");
    let python = std::env::var_os("PYTHON").unwrap_or_else(|| "python".into());
    let mut command = Command::new(python);
    command
        .arg(comparator)
        .arg("--source-root")
        .arg(python_source_root)
        .arg("--config-workbook")
        .arg(config)
        .arg("--audit-copy")
        .arg(primary)
        .arg("--source-id")
        .arg("CUR_A")
        .arg("--source-business")
        .arg("其他业务")
        .arg("--period")
        .arg("2026-08-31")
        .arg("--institution-code")
        .arg("I001")
        .arg("--institution-name")
        .arg("合成机构");
    for (source_id, business, path) in additional_sources {
        command
            .arg("--current-source")
            .arg(format!("{source_id}={business}={}", path.display()));
    }
    command
        .arg("--business")
        .arg("其他业务")
        .arg("--business")
        .arg("单位贷款")
        .arg("--template-dir")
        .arg(template_dir)
        .arg("--rust-bin")
        .arg(rust_binary)
        .arg("--report")
        .arg(report_path);
    let output = command
        .output()
        .expect("run Python/Rust multi-source parity comparator");
    assert!(
        output.status.success(),
        "{}\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );
    serde_json::from_slice(&fs::read(report_path).expect("read multi-source parity report"))
        .expect("valid multi-source parity report")
}

fn run_config_check_cli(config: &Path, external: Option<&Path>, developer_view: bool) -> Output {
    let mut command = Command::new(env!("CARGO_BIN_EXE_open-data-audit"));
    command
        .arg("s5-config-check")
        .arg("--config")
        .arg(config)
        .arg("--business")
        .arg("单位贷款");
    if let Some(external) = external {
        command.arg("--external-path").arg(external);
    }
    if developer_view {
        command.arg("--developer-view");
    }
    command.output().expect("run S5 config-check CLI")
}

fn run_full_scope_config_check_cli(config: &Path, template_dir: &Path, external: &Path) -> Output {
    Command::new(env!("CARGO_BIN_EXE_open-data-audit"))
        .arg("s5-config-check")
        .arg("--config")
        .arg(config)
        .arg("--template-dir")
        .arg(template_dir)
        .arg("--external-path")
        .arg(external)
        .arg("--scope")
        .arg("all")
        .arg("--developer-view")
        .output()
        .expect("run full-scope S5 config-check CLI")
}

#[test]
fn s5_config_check_cli_scans_external_metric_headers() {
    let directory = temporary_directory();
    let config = directory.join("config.xlsx");
    let external = directory.join("external.xlsx");
    write_config(&config, true, false);
    write_external_source(&external, 0.15);

    let output = run_config_check_cli(&config, Some(&external), true);
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report: Value = serde_json::from_slice(&output.stdout).expect("valid config-check JSON");
    assert_eq!(report["valid"], true);
    assert_eq!(report["complete"], false);
    assert_eq!(report["checks"]["external_addresses"]["status"], "passed");
    assert_eq!(report["checks"]["external_addresses"]["checked"], 1);
    assert_eq!(report["checks"]["external_addresses"]["resolved"], 1);
    assert_eq!(
        report["checks"]["template_addresses"]["status"],
        "not_checked"
    );
    let developer_log = String::from_utf8_lossy(&output.stderr);
    assert!(developer_log.contains("s5_config_check_started"));
    assert!(developer_log.contains("s5_config_check_completed"));
    assert!(developer_log.contains("Calamine via excel-reader"));

    fs::remove_dir_all(directory).expect("remove config-check fixture directory");
}

#[test]
fn s5_config_check_cli_accepts_python_external_identity_alias() {
    let directory = temporary_directory();
    let config = directory.join("config.xlsx");
    let external = directory.join("external.xlsx");
    write_config(&config, true, false);
    write_external_source_with_python_code_alias(&external, 0.15);

    let output = run_config_check_cli(&config, Some(&external), false);
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report: Value = serde_json::from_slice(&output.stdout).expect("valid config-check JSON");
    assert_eq!(report["valid"], true);
    assert_eq!(report["checks"]["external_addresses"]["status"], "passed");
    assert_eq!(report["checks"]["external_addresses"]["checked"], 1);
    assert_eq!(report["checks"]["external_addresses"]["resolved"], 1);

    fs::remove_dir_all(directory).expect("remove external alias fixture directory");
}

#[test]
fn s5_config_check_cli_marks_missing_external_file_invalid() {
    let directory = temporary_directory();
    let config = directory.join("config.xlsx");
    let missing_external = directory.join("missing.xlsx");
    write_config(&config, true, false);

    let output = run_config_check_cli(&config, Some(&missing_external), false);
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report: Value = serde_json::from_slice(&output.stdout).expect("valid config-check JSON");
    assert_eq!(report["valid"], false);
    assert_eq!(report["checks"]["external_addresses"]["status"], "failed");
    assert_eq!(report["configuration_errors"].as_array().unwrap().len(), 1);
    assert!(
        report["checks"]["external_addresses"]["errors"][0]["message"]
            .as_str()
            .unwrap()
            .contains("外部文件不存在")
    );

    fs::remove_dir_all(directory).expect("remove config-check failure directory");
}

#[test]
fn s5_config_check_cli_scans_full_scope_template_catalog() {
    let directory = temporary_directory();
    let config = directory.join("config.xlsx");
    let external = directory.join("external.xlsx");
    let template_dir = directory.join("templates");
    fs::create_dir_all(&template_dir).expect("create template catalog directory");
    write_config(&config, true, false);
    write_external_source(&external, 0.15);
    write_source(&template_dir.join("单位贷款模板.xlsx"), 115.0);
    write_source(&template_dir.join("备用模板.xlsx"), 115.0);

    let output = run_full_scope_config_check_cli(&config, &template_dir, &external);
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report: Value = serde_json::from_slice(&output.stdout).expect("valid config-check JSON");
    assert_eq!(report["valid"], true);
    assert_eq!(report["complete"], true);
    assert_eq!(report["scope"], "all");
    assert_eq!(report["business"], "");
    assert_eq!(report["checks"]["template_addresses"]["status"], "passed");
    assert_eq!(report["checks"]["template_addresses"]["checked"], 1);
    assert_eq!(report["checks"]["template_addresses"]["resolved"], 1);
    assert_eq!(
        report["checks"]["template_addresses"]["businesses"]
            .as_array()
            .unwrap()
            .len(),
        1
    );
    assert_eq!(
        report["checks"]["template_addresses"]["businesses"][0]["template_name"],
        "单位贷款模板.xlsx"
    );
    assert_eq!(report["checks"]["external_addresses"]["status"], "passed");
    let events = String::from_utf8_lossy(&output.stderr)
        .lines()
        .map(|line| serde_json::from_str::<Value>(line).expect("progress JSON"))
        .collect::<Vec<_>>();
    assert_eq!(
        events[0]["template_catalog_engine"],
        "OOXML workbook.xml via excel-reader"
    );
    assert!(events[1]["elapsed_ms"].as_u64().is_some());

    fs::remove_dir_all(directory).expect("remove full-scope config-check directory");
}

#[test]
fn s5_config_check_cli_marks_full_scope_without_catalog_incomplete() {
    let directory = temporary_directory();
    let config = directory.join("config.xlsx");
    let external = directory.join("external.xlsx");
    let missing_template_dir = directory.join("missing_templates");
    write_config(&config, true, false);
    write_external_source(&external, 0.15);

    let output = run_full_scope_config_check_cli(&config, &missing_template_dir, &external);
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report: Value = serde_json::from_slice(&output.stdout).expect("valid config-check JSON");
    assert_eq!(report["valid"], true);
    assert_eq!(report["complete"], false);
    assert_eq!(
        report["checks"]["template_addresses"]["status"],
        "not_checked"
    );
    assert!(report["warnings"]
        .as_array()
        .unwrap()
        .iter()
        .any(|warning| { warning.as_str().unwrap_or_default().contains("模板目录") }));
    assert_eq!(report["checks"]["external_addresses"]["status"], "passed");

    fs::remove_dir_all(directory).expect("remove full-scope missing-catalog directory");
}

#[test]
#[ignore = "manual parity fixture exporter; requires AUDIT_S5_CONFIG_FIXTURE_DIR"]
fn exports_s5_config_check_fixture_for_python_parity() {
    let directory = PathBuf::from(
        std::env::var_os("AUDIT_S5_CONFIG_FIXTURE_DIR")
            .expect("AUDIT_S5_CONFIG_FIXTURE_DIR must point to a temporary output directory"),
    );
    fs::create_dir_all(&directory).expect("create parity fixture directory");
    let template_dir = directory.join("templates");
    fs::create_dir_all(&template_dir).expect("create parity template directory");
    write_config(&directory.join("config.xlsx"), true, false);
    write_external_source(&directory.join("external.xlsx"), 0.15);
    write_source(&template_dir.join("单位贷款模板.xlsx"), 115.0);
    write_source(&template_dir.join("备用模板.xlsx"), 115.0);
}

#[test]
fn s5_audit_cli_reads_and_copies_independent_previous_source() {
    let directory = temporary_directory();
    let current = directory.join("current.xlsx");
    let previous = directory.join("previous.xlsx");
    let config = directory.join("config.xlsx");
    let output_dir = directory.join("out");
    write_source(&current, 120.0);
    write_source(&previous, 90.0);
    write_config(&config, false, false);

    let output = run_cli(&config, &current, Some(&previous), None, &output_dir, true);
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    let summary: Value = serde_json::from_slice(&output.stdout).expect("valid CLI JSON");
    assert_eq!(summary["status"], "completed");
    assert_eq!(summary["status_counts"]["TRIGGERED"], 1);
    assert_eq!(
        summary["source_roles"],
        serde_json::json!(["current", "previous"])
    );
    let result_path = PathBuf::from(
        summary["output_files"]["result_workbook"]
            .as_str()
            .expect("result workbook path"),
    );
    let result_workbook = read_workbook(&result_path).expect("read S5 result workbook");
    let result_sheet = result_workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "审核结果")
        .expect("result sheet");
    assert_eq!(
        CellValue::from_probe_json(&result_sheet.values["J2"]).expect("current value"),
        CellValue::Number(120.0)
    );
    assert_eq!(
        CellValue::from_probe_json(&result_sheet.values["K2"]).expect("previous value"),
        CellValue::Number(90.0)
    );
    assert_eq!(
        CellValue::from_probe_json(&result_sheet.values["L2"]).expect("difference value"),
        CellValue::Number(30.0)
    );

    let events = String::from_utf8_lossy(&output.stderr)
        .lines()
        .map(|line| serde_json::from_str::<Value>(line).expect("progress JSON"))
        .collect::<Vec<_>>();
    assert_eq!(
        events.first().expect("start event")["source_roles"],
        serde_json::json!(["current", "previous"])
    );
    assert!(events.last().expect("completion event")["elapsed_ms"]
        .as_u64()
        .is_some());

    let ledger_path = PathBuf::from(
        summary["output_files"]["run_ledger"]
            .as_str()
            .expect("run ledger path"),
    );
    let ledger: Value = serde_json::from_slice(&fs::read(&ledger_path).expect("read run ledger"))
        .expect("valid run ledger");
    let previous_entry = ledger["audit_copy_manifest"]
        .as_array()
        .expect("audit copy manifest")
        .iter()
        .find(|entry| entry["role"] == "previous")
        .expect("previous source manifest entry");
    assert_eq!(previous_entry["source_id"], "CUR");
    assert_eq!(previous_entry["matches_snapshot"], true);
    let previous_copy = PathBuf::from(
        previous_entry["copy_path"]
            .as_str()
            .expect("previous source copy path"),
    );
    assert_eq!(
        fs::read(previous_copy).expect("read previous copy"),
        fs::read(&previous).expect("read original previous")
    );

    let output_without_previous = run_cli(
        &config,
        &current,
        None,
        None,
        &directory.join("out-without-previous"),
        false,
    );
    assert!(
        output_without_previous.status.success(),
        "{}",
        String::from_utf8_lossy(&output_without_previous.stderr)
    );
    let missing_summary: Value =
        serde_json::from_slice(&output_without_previous.stdout).expect("valid no-previous JSON");
    assert_eq!(missing_summary["status_counts"]["RETRIEVAL_FAILED"], 1);

    fs::remove_dir_all(directory).expect("remove synthetic S5 CLI files");
}

#[test]
fn s5_audit_cli_selects_multiple_current_sources_by_rule_binding() {
    let directory = temporary_directory();
    let config = directory.join("multi-current-config.xlsx");
    let primary = directory.join("primary.xlsx");
    let additional = directory.join("additional.xlsx");
    let output_dir = directory.join("out");
    let template_dir = directory.join("empty-templates");
    fs::create_dir_all(&template_dir).expect("create empty template directory");
    write_multi_current_source_config(&config);
    write_source(&primary, 111.0);
    write_source(&additional, 222.0);

    let additional_sources = [("CUR_B", "单位贷款", additional.as_path())];
    let output = run_multi_current_cli(
        &config,
        &primary,
        &additional_sources,
        &template_dir,
        &output_dir,
        "multi-current",
    );
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    let summary: Value = serde_json::from_slice(&output.stdout).expect("valid multi-source JSON");
    assert_eq!(summary["status"], "completed");
    assert_eq!(summary["status_counts"]["TRIGGERED"], 2);
    assert_eq!(
        summary["source_roles"],
        serde_json::json!(["current", "additional current"])
    );

    let result_path = PathBuf::from(
        summary["output_files"]["result_workbook"]
            .as_str()
            .expect("result workbook path"),
    );
    let result_workbook = read_workbook(&result_path).expect("read multi-source result workbook");
    let result_sheet = result_workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "审核结果")
        .expect("result sheet");
    assert_eq!(
        CellValue::from_probe_json(&result_sheet.values["J2"]).expect("primary source value"),
        CellValue::Number(111.0)
    );
    assert_eq!(
        CellValue::from_probe_json(&result_sheet.values["J3"]).expect("additional source value"),
        CellValue::Number(222.0)
    );

    let ledger_path = PathBuf::from(
        summary["output_files"]["run_ledger"]
            .as_str()
            .expect("run ledger path"),
    );
    let ledger: Value = serde_json::from_slice(&fs::read(&ledger_path).expect("read ledger"))
        .expect("valid ledger");
    let manifest = ledger["audit_copy_manifest"]
        .as_array()
        .expect("audit-copy manifest");
    for (source_id, original) in [("CUR_A", &primary), ("CUR_B", &additional)] {
        let entry = manifest
            .iter()
            .find(|entry| entry["role"] == "current" && entry["source_id"] == source_id)
            .expect("current source manifest entry");
        assert_eq!(entry["matches_snapshot"], true);
        let copy_path = PathBuf::from(entry["copy_path"].as_str().expect("source copy path"));
        assert_eq!(
            fs::read(copy_path).expect("read audit copy"),
            fs::read(original).expect("read source input")
        );
    }
    let events = String::from_utf8_lossy(&output.stderr)
        .lines()
        .map(|line| serde_json::from_str::<Value>(line).expect("progress JSON"))
        .collect::<Vec<_>>();
    assert_eq!(
        events.first().expect("start event")["current_source_ids"],
        serde_json::json!(["CUR_A", "CUR_B"])
    );

    if let Some(python_source_root) = std::env::var_os("AUDIT_PYTHON_SOURCE_ROOT") {
        let report_path = std::env::var_os("AUDIT_S5_MULTI_SOURCE_REPORT")
            .map(PathBuf::from)
            .unwrap_or_else(|| directory.join("multi-source-parity.json"));
        let python_source_root = PathBuf::from(python_source_root);
        let report = run_multi_source_python_rust_parity(
            &python_source_root,
            &config,
            &primary,
            &additional_sources,
            &template_dir,
            &report_path,
        );
        assert_eq!(report["status"], "completed");
        assert_eq!(report["comparison"]["diff_count"], 0);
        assert_eq!(report["verification"]["inputs_unchanged"], true);
        assert_eq!(
            report["comparison"]["status_counts"]["rust"]["TRIGGERED"],
            2
        );
    }

    let ambiguous = directory.join("ambiguous.xlsx");
    write_source(&ambiguous, 333.0);
    let ambiguous_sources = [
        ("CUR_B", "单位贷款", additional.as_path()),
        ("CUR_C", "单位贷款", ambiguous.as_path()),
    ];
    let ambiguous_output = run_multi_current_cli(
        &config,
        &primary,
        &ambiguous_sources,
        &template_dir,
        &directory.join("out-ambiguous"),
        "ambiguous-current",
    );
    assert!(
        ambiguous_output.status.success(),
        "{}",
        String::from_utf8_lossy(&ambiguous_output.stderr)
    );
    let ambiguous_summary: Value =
        serde_json::from_slice(&ambiguous_output.stdout).expect("valid ambiguous-source JSON");
    assert_eq!(ambiguous_summary["status"], "completed");
    assert_eq!(ambiguous_summary["status_counts"]["TRIGGERED"], 1);
    assert_eq!(ambiguous_summary["status_counts"]["CONFIG_ERROR"], 1);
    let ambiguous_result = read_workbook(
        ambiguous_summary["output_files"]["result_workbook"]
            .as_str()
            .expect("ambiguous result workbook path"),
    )
    .expect("read ambiguous-source result workbook");
    let ambiguous_sheet = ambiguous_result
        .sheets
        .iter()
        .find(|sheet| sheet.name == "审核结果")
        .expect("ambiguous result sheet");
    assert_eq!(
        CellValue::from_probe_json(&ambiguous_sheet.values["I3"])
            .expect("ambiguous-source error description"),
        CellValue::Text("业务 单位贷款 对应多个current数据源，请显式选择逻辑来源".to_owned())
    );
    if let Some(python_source_root) = std::env::var_os("AUDIT_PYTHON_SOURCE_ROOT") {
        let python_source_root = PathBuf::from(python_source_root);
        let ambiguous_report = std::env::var_os("AUDIT_S5_MULTI_SOURCE_AMBIGUOUS_REPORT")
            .map(PathBuf::from)
            .unwrap_or_else(|| directory.join("ambiguous-multi-source-parity.json"));
        let report = run_multi_source_python_rust_parity(
            &python_source_root,
            &config,
            &primary,
            &ambiguous_sources,
            &template_dir,
            &ambiguous_report,
        );
        assert_eq!(report["status"], "completed");
        assert_eq!(report["comparison"]["diff_count"], 0);
        assert_eq!(report["verification"]["inputs_unchanged"], true);
        assert_eq!(
            report["comparison"]["status_counts"]["python"]["CONFIG_ERROR"],
            1
        );
        assert_eq!(
            report["comparison"]["status_counts"]["rust"]["CONFIG_ERROR"],
            1
        );
    }

    fs::remove_dir_all(directory).expect("remove multi-source S5 CLI files");
}

#[test]
fn s5_audit_cli_reads_and_copies_configured_external_source() {
    let directory = temporary_directory();
    let current = directory.join("current.xlsx");
    let external = directory.join("external.xlsx");
    let config = directory.join("config.xlsx");
    let output_dir = directory.join("out");
    write_source(&current, 1500.0);
    write_external_source(&external, 0.15);
    write_config(&config, true, true);

    let output = run_cli(&config, &current, None, Some(&external), &output_dir, true);
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    let summary: Value = serde_json::from_slice(&output.stdout).expect("valid CLI JSON");
    assert_eq!(summary["status"], "completed");
    assert_eq!(summary["status_counts"]["TRIGGERED"], 1);
    assert_eq!(
        summary["source_roles"],
        serde_json::json!(["current", "external"])
    );
    let result_path = PathBuf::from(
        summary["output_files"]["result_workbook"]
            .as_str()
            .expect("result workbook path"),
    );
    let result_workbook = read_workbook(&result_path).expect("read external S5 result workbook");
    let result_sheet = result_workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "审核结果")
        .expect("result sheet");
    assert_eq!(
        CellValue::from_probe_json(&result_sheet.values["J2"]).expect("current value"),
        CellValue::Number(1500.0)
    );
    assert_eq!(
        CellValue::from_probe_json(&result_sheet.values["K2"]).expect("external value"),
        CellValue::Number(1500.0)
    );
    assert_eq!(
        CellValue::from_probe_json(&result_sheet.values["L2"]).expect("difference value"),
        CellValue::Number(0.0)
    );
    assert_eq!(
        CellValue::from_probe_json(&result_sheet.values["I2"]).expect("difference description"),
        CellValue::Text("差异=0.00".to_owned())
    );

    let ledger_path = PathBuf::from(
        summary["output_files"]["run_ledger"]
            .as_str()
            .expect("run ledger path"),
    );
    let ledger: Value = serde_json::from_slice(&fs::read(&ledger_path).expect("read run ledger"))
        .expect("valid run ledger");
    let external_entry = ledger["audit_copy_manifest"]
        .as_array()
        .expect("audit copy manifest")
        .iter()
        .find(|entry| entry["role"] == "external")
        .expect("external source manifest entry");
    assert_eq!(external_entry["source_id"], "EXT");
    assert_eq!(external_entry["matches_snapshot"], true);
    let external_copy = PathBuf::from(
        external_entry["copy_path"]
            .as_str()
            .expect("external source copy path"),
    );
    assert_eq!(
        fs::read(external_copy).expect("read external copy"),
        fs::read(&external).expect("read external source")
    );

    let output_without_external = run_cli(
        &config,
        &current,
        None,
        None,
        &directory.join("out-without-external"),
        false,
    );
    assert!(
        output_without_external.status.success(),
        "{}",
        String::from_utf8_lossy(&output_without_external.stderr)
    );
    let missing_summary: Value =
        serde_json::from_slice(&output_without_external.stdout).expect("valid no-external JSON");
    assert_eq!(missing_summary["status_counts"]["RETRIEVAL_FAILED"], 1);

    fs::remove_dir_all(directory).expect("remove synthetic external CLI files");
}

#[test]
fn s5_audit_cli_projects_numeric_regions_from_the_selected_template_catalog() {
    let directory = temporary_directory();
    let template_dir = directory.join("templates");
    let empty_template_dir = directory.join("empty-templates");
    fs::create_dir_all(&template_dir).expect("create numeric-scope template directory");
    fs::create_dir_all(&empty_template_dir).expect("create empty template directory");
    let current = directory.join("current.xlsx");
    let previous = directory.join("previous.xlsx");
    let config = directory.join("config.xlsx");
    let template = template_dir.join("loan-template.xlsx");
    write_numeric_scope_source(&current, 121.0);
    write_numeric_scope_source(&previous, 100.0);
    write_config(&config, false, false);
    write_numeric_scope_template(&template);

    let unbound = run_numeric_scope_cli(
        &config,
        &current,
        &previous,
        None,
        Some(&empty_template_dir),
        None,
        &directory.join("out-unbound"),
    );
    assert!(
        unbound.status.success(),
        "{}",
        String::from_utf8_lossy(&unbound.stderr)
    );
    let unbound_summary: Value =
        serde_json::from_slice(&unbound.stdout).expect("valid unbound CLI JSON");
    assert_eq!(unbound_summary["status_counts"]["STRUCTURE_ERROR"], 1);
    assert_eq!(unbound_summary["failure_rows"], 1);

    let bound = run_numeric_scope_cli(
        &config,
        &current,
        &previous,
        None,
        Some(&template_dir),
        None,
        &directory.join("out-bound"),
    );
    assert!(
        bound.status.success(),
        "{}",
        String::from_utf8_lossy(&bound.stderr)
    );
    let summary: Value = serde_json::from_slice(&bound.stdout).expect("valid bound CLI JSON");
    assert_eq!(summary["status_counts"]["TRIGGERED"], 1);
    assert_eq!(summary["triggered_rows"], 1);
    assert_eq!(summary["failure_rows"], 0);
    let result_path = PathBuf::from(
        summary["output_files"]["result_workbook"]
            .as_str()
            .expect("result workbook path"),
    );
    let result_workbook = read_workbook(&result_path).expect("read numeric-scope result workbook");
    let result_sheet = result_workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "审核结果")
        .expect("result sheet");
    assert_eq!(
        CellValue::from_probe_json(&result_sheet.values["J2"]).expect("current value"),
        CellValue::Number(121.0)
    );
    assert_eq!(
        CellValue::from_probe_json(&result_sheet.values["K2"]).expect("previous value"),
        CellValue::Number(100.0)
    );
    assert_eq!(
        CellValue::from_probe_json(&result_sheet.values["F2"]).expect("retrieval location"),
        CellValue::Text("D6".to_owned())
    );
    let events: Vec<Value> = String::from_utf8_lossy(&bound.stderr)
        .lines()
        .map(|line| serde_json::from_str(line).expect("valid progress JSON"))
        .collect();
    let template_selections = events
        .iter()
        .find(|event| event["status"] == "completed")
        .expect("completed progress event")["numeric_region_templates"]
        .as_array()
        .expect("numeric-region template selections");
    assert_eq!(template_selections.len(), 2);
    assert!(template_selections
        .iter()
        .all(|selection| selection["template"] == "loan-template.xlsx"));
    assert!(template_selections.iter().all(|selection| {
        selection["engine"] == "Calamine via excel-reader + namespace-aware defined names"
    }));

    fs::copy(&template, template_dir.join("duplicate-template.xlsx"))
        .expect("create tied template candidate");
    let ambiguous = run_numeric_scope_cli(
        &config,
        &current,
        &previous,
        None,
        Some(&template_dir),
        None,
        &directory.join("out-ambiguous"),
    );
    assert!(
        ambiguous.status.success(),
        "{}",
        String::from_utf8_lossy(&ambiguous.stderr)
    );
    let ambiguous_summary: Value =
        serde_json::from_slice(&ambiguous.stdout).expect("valid ambiguous CLI JSON");
    assert_eq!(ambiguous_summary["status_counts"]["STRUCTURE_ERROR"], 1);
    assert_eq!(ambiguous_summary["failure_rows"], 1);

    let explicit = run_numeric_scope_cli(
        &config,
        &current,
        &previous,
        Some(&template),
        Some(&template_dir),
        None,
        &directory.join("out-explicit"),
    );
    assert!(
        explicit.status.success(),
        "{}",
        String::from_utf8_lossy(&explicit.stderr)
    );
    let explicit_summary: Value =
        serde_json::from_slice(&explicit.stdout).expect("valid explicit-template CLI JSON");
    assert_eq!(explicit_summary["status_counts"]["TRIGGERED"], 1);

    let base_data_dir = directory.join("default-data");
    let default_template_dir = base_data_dir.join("s5").join("templates");
    fs::create_dir_all(&default_template_dir).expect("create default S5 template directory");
    fs::copy(&template, default_template_dir.join("loan-template.xlsx"))
        .expect("copy template to default catalog fixture");
    let default_catalog = run_numeric_scope_cli(
        &config,
        &current,
        &previous,
        None,
        None,
        Some(&base_data_dir),
        &directory.join("out-default-catalog"),
    );
    assert!(
        default_catalog.status.success(),
        "{}",
        String::from_utf8_lossy(&default_catalog.stderr)
    );
    let default_summary: Value = serde_json::from_slice(&default_catalog.stdout)
        .expect("valid default-template-catalog CLI JSON");
    assert_eq!(default_summary["status_counts"]["TRIGGERED"], 1);

    fs::remove_dir_all(directory).expect("remove numeric-scope CLI files");
}

#[test]
fn s5_audit_cli_retrieves_ordinary_external_layout_without_contract() {
    let directory = temporary_directory();
    let current = directory.join("current.xlsx");
    let external = directory.join("external.xlsx");
    let config = directory.join("config.xlsx");
    let output_dir = directory.join("out");
    write_source(&current, 1500.0);
    write_external_source(&external, 0.15);
    write_config(&config, true, false);

    let output = run_cli(&config, &current, None, Some(&external), &output_dir, true);
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    let summary: Value = serde_json::from_slice(&output.stdout).expect("valid CLI JSON");
    assert_eq!(summary["status"], "completed");
    assert_eq!(summary["status_counts"]["TRIGGERED"], 1);
    let result_path = PathBuf::from(
        summary["output_files"]["result_workbook"]
            .as_str()
            .expect("result workbook path"),
    );
    let result_workbook = read_workbook(&result_path).expect("read dynamic-layout result workbook");
    let result_sheet = result_workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "审核结果")
        .expect("result sheet");
    assert_eq!(
        CellValue::from_probe_json(&result_sheet.values["K2"]).expect("external value"),
        CellValue::Number(1500.0)
    );
    assert_eq!(
        CellValue::from_probe_json(&result_sheet.values["L2"]).expect("difference value"),
        CellValue::Number(0.0)
    );

    fs::remove_dir_all(directory).expect("remove synthetic dynamic-layout files");
}

#[test]
fn s5_audit_cli_records_dynamic_external_layout_warnings() {
    let directory = temporary_directory();
    let current = directory.join("current.xlsx");
    let external = directory.join("external.xlsx");
    let config = directory.join("config.xlsx");
    let output_dir = directory.join("out");
    write_source(&current, 1500.0);
    write_external_source_with_layout_warning(&external, 0.15, true);
    write_config(&config, true, false);

    let output = run_cli(&config, &current, None, Some(&external), &output_dir, true);
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    let summary: Value = serde_json::from_slice(&output.stdout).expect("valid CLI JSON");
    assert_eq!(summary["status"], "completed");
    assert_eq!(summary["status_counts"]["TRIGGERED"], 1);
    assert_eq!(summary["data_quality_warning_count"], 1);
    let result_path = PathBuf::from(
        summary["output_files"]["result_workbook"]
            .as_str()
            .expect("result workbook path"),
    );
    let result_workbook = read_workbook(&result_path).expect("read layout-warning result");
    let warning_sheet = result_workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "数据质量提示")
        .expect("data-quality warning sheet");
    assert_eq!(
        CellValue::from_probe_json(&warning_sheet.values["C2"]).expect("source file"),
        CellValue::Text("external.xlsx".to_owned())
    );
    assert_eq!(
        CellValue::from_probe_json(&warning_sheet.values["E2"]).expect("warning location"),
        CellValue::Text("第4行".to_owned())
    );
    let ledger_path = PathBuf::from(
        summary["output_files"]["run_ledger"]
            .as_str()
            .expect("run ledger path"),
    );
    let ledger: Value =
        serde_json::from_slice(&fs::read(ledger_path).expect("read ledger")).expect("valid ledger");
    assert_eq!(ledger["data_quality_warnings"].as_array().unwrap().len(), 1);

    fs::remove_dir_all(directory).expect("remove synthetic warning-layout files");
}

#[test]
fn s5_audit_cli_retrieves_dynamic_external_time_series_for_current_period() {
    let directory = temporary_directory();
    let current = directory.join("current.xlsx");
    let external = directory.join("external.xlsx");
    let config = directory.join("config.xlsx");
    let output_dir = directory.join("out");
    write_source(&current, 1500.0);
    write_external_time_series_source(&external, 0.15);
    write_config(&config, true, false);

    let output = run_cli(&config, &current, None, Some(&external), &output_dir, true);
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    let summary: Value = serde_json::from_slice(&output.stdout).expect("valid CLI JSON");
    assert_eq!(summary["status"], "completed");
    assert_eq!(summary["status_counts"]["TRIGGERED"], 1);
    assert_eq!(summary["data_quality_warning_count"], 1);
    let result_path = PathBuf::from(
        summary["output_files"]["result_workbook"]
            .as_str()
            .expect("result workbook path"),
    );
    let result_workbook = read_workbook(&result_path).expect("read time-series result");
    let result_sheet = result_workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "审核结果")
        .expect("result sheet");
    assert_eq!(
        CellValue::from_probe_json(&result_sheet.values["K2"]).expect("selected external value"),
        CellValue::Number(1500.0)
    );
    assert_eq!(
        CellValue::from_probe_json(&result_sheet.values["L2"]).expect("difference value"),
        CellValue::Number(0.0)
    );
    let warning_sheet = result_workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "数据质量提示")
        .expect("invalid-date warning sheet");
    assert!(matches!(
        CellValue::from_probe_json(&warning_sheet.values["F2"]).expect("warning message"),
        CellValue::Text(message) if message.contains("数据日期“not-a-date”为空或无效")
    ));
    let ledger_path = PathBuf::from(
        summary["output_files"]["run_ledger"]
            .as_str()
            .expect("run ledger path"),
    );
    let ledger: Value =
        serde_json::from_slice(&fs::read(ledger_path).expect("read time-series ledger"))
            .expect("valid time-series ledger");
    assert_eq!(ledger["data_quality_warnings"][0]["institution_name"], "");
    assert_eq!(ledger["data_quality_warnings"][0]["institution_code"], "");

    fs::remove_dir_all(directory).expect("remove synthetic time-series files");
}

#[test]
fn s5_audit_cli_uses_dynamic_external_time_series_when_contract_exists() {
    let directory = temporary_directory();
    let current = directory.join("current.xlsx");
    let external = directory.join("external.xlsx");
    let config = directory.join("config.xlsx");
    let output_dir = directory.join("out");
    write_source(&current, 1500.0);
    write_external_time_series_source(&external, 0.15);
    write_config(&config, true, true);

    let output = run_cli(&config, &current, None, Some(&external), &output_dir, true);
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    let summary: Value = serde_json::from_slice(&output.stdout).expect("valid CLI JSON");
    assert_eq!(summary["status"], "completed");
    assert_eq!(summary["status_counts"]["TRIGGERED"], 1);
    let result_path = PathBuf::from(
        summary["output_files"]["result_workbook"]
            .as_str()
            .expect("result workbook path"),
    );
    let result_workbook = read_workbook(&result_path).expect("read contract time-series result");
    let result_sheet = result_workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "审核结果")
        .expect("result sheet");
    assert_eq!(
        CellValue::from_probe_json(&result_sheet.values["K2"]).expect("selected external value"),
        CellValue::Number(1500.0)
    );
    fs::remove_dir_all(directory).expect("remove contract time-series files");
}

#[test]
fn s5_audit_cli_reports_missing_current_external_time_series_date() {
    let directory = temporary_directory();
    let current = directory.join("current.xlsx");
    let external = directory.join("external.xlsx");
    let config = directory.join("config.xlsx");
    let output_dir = directory.join("out");
    write_source(&current, 1500.0);
    write_external_time_series_source_without_current_date(&external, 0.15);
    write_config(&config, true, false);

    let output = run_cli(&config, &current, None, Some(&external), &output_dir, true);
    assert!(output.status.success());
    let summary: Value = serde_json::from_slice(&output.stdout).expect("valid CLI JSON");
    assert_eq!(summary["status_counts"]["RETRIEVAL_FAILED"], 1);
    assert_eq!(summary["data_quality_warning_count"], 1);
    let result_path = PathBuf::from(
        summary["output_files"]["result_workbook"]
            .as_str()
            .expect("result workbook path"),
    );
    let result_workbook = read_workbook(&result_path).expect("read missing-date result");
    let result_sheet = result_workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "审核结果")
        .expect("result sheet");
    assert!(matches!(
        CellValue::from_probe_json(&result_sheet.values["I2"]).expect("error message"),
        CellValue::Text(message) if message.contains("没有本期日期 2026-08-31 的记录")
    ));
    fs::remove_dir_all(directory).expect("remove missing-date files");
}

#[test]
fn s5_audit_cli_uses_zero_and_records_invalid_dynamic_external_numeric_values() {
    let directory = temporary_directory();
    let current = directory.join("current.xlsx");
    let external = directory.join("external.xlsx");
    let config = directory.join("config.xlsx");
    let output_dir = directory.join("out");
    write_source(&current, 1500.0);
    write_external_source_with_invalid_numeric(&external);
    write_config(&config, true, false);

    let output = run_cli(&config, &current, None, Some(&external), &output_dir, true);
    assert!(output.status.success());
    let summary: Value = serde_json::from_slice(&output.stdout).expect("valid CLI JSON");
    assert_eq!(summary["status_counts"]["PASSED"], 1);
    assert_eq!(summary["data_quality_warning_count"], 1);
    let result_path = PathBuf::from(
        summary["output_files"]["result_workbook"]
            .as_str()
            .expect("result workbook path"),
    );
    let result_workbook = read_workbook(&result_path).expect("read invalid-value result");
    let result_sheet = result_workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "审核结果")
        .expect("result sheet");
    assert!(!result_sheet.values.contains_key("J2"));
    let warning_sheet = result_workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "数据质量提示")
        .expect("data-quality warning sheet");
    assert!(matches!(
        CellValue::from_probe_json(&warning_sheet.values["F2"]).expect("warning message"),
        CellValue::Text(message) if message.contains("已按 0 参与审核")
    ));
    let ledger_path = PathBuf::from(
        summary["output_files"]["run_ledger"]
            .as_str()
            .expect("run ledger path"),
    );
    let ledger: Value =
        serde_json::from_slice(&fs::read(ledger_path).expect("read ledger")).expect("valid ledger");
    assert_eq!(ledger["data_quality_warnings"].as_array().unwrap().len(), 1);

    fs::remove_dir_all(directory).expect("remove synthetic invalid-value files");
}

#[test]
fn s5_audit_cli_accepts_whole_number_dynamic_external_values() {
    let directory = temporary_directory();
    let current = directory.join("current.xlsx");
    let external = directory.join("external.xlsx");
    let config = directory.join("config.xlsx");
    let output_dir = directory.join("out");
    write_source(&current, 1500.0);
    write_external_source(&external, 1.0);
    write_config(&config, true, false);

    let input_workbook = read_workbook(&external).expect("read integer external source");
    let input_sheet = input_workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "集中系统数据")
        .expect("external data sheet");
    assert_eq!(
        CellValue::from_probe_json(&input_sheet.values["C3"]).expect("whole-number metric"),
        CellValue::Number(1.0),
        "test input must be a whole-number amount"
    );

    let output = run_cli(&config, &current, None, Some(&external), &output_dir, true);
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    let summary: Value = serde_json::from_slice(&output.stdout).expect("valid CLI JSON");
    assert_eq!(summary["status_counts"]["PASSED"], 1);
    assert_eq!(summary["data_quality_warning_count"], 0);
    let result_path = PathBuf::from(
        summary["output_files"]["result_workbook"]
            .as_str()
            .expect("result workbook path"),
    );
    let result_workbook = read_workbook(&result_path).expect("read integer-value result");
    assert!(result_workbook
        .sheets
        .iter()
        .all(|sheet| sheet.name != "数据质量提示"));

    let ledger_path = PathBuf::from(
        summary["output_files"]["run_ledger"]
            .as_str()
            .expect("run ledger path"),
    );
    let ledger: Value =
        serde_json::from_slice(&fs::read(ledger_path).expect("read ledger")).expect("valid ledger");
    assert_eq!(ledger["data_quality_warnings"].as_array().unwrap().len(), 0);

    fs::remove_dir_all(directory).expect("remove synthetic integer-value files");
}
