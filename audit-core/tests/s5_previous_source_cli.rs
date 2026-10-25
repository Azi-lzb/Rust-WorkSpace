use excel_reader::{read_workbook, CellValue};
use serde_json::Value;
use std::{
    fs,
    path::{Path, PathBuf},
    process::{Command, Output},
    time::{SystemTime, UNIX_EPOCH},
};

fn temporary_directory() -> PathBuf {
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system clock after Unix epoch")
        .as_nanos();
    let path = std::env::temp_dir().join(format!(
        "audit_s5_previous_source_cli_{}_{}",
        std::process::id(),
        nonce
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

fn write_external_source(path: &Path, value: f64) {
    write_external_source_variant(path, value, false, false, false);
}

fn write_external_source_with_layout_warning(path: &Path, value: f64, warning: bool) {
    write_external_source_variant(path, value, warning, false, false);
}

fn write_external_source_with_invalid_numeric(path: &Path) {
    write_external_source_variant(path, 0.0, false, true, false);
}

fn write_external_time_series_source(path: &Path, value: f64) {
    write_external_source_variant(path, value, false, false, true);
}

fn write_external_source_variant(
    path: &Path,
    value: f64,
    warning: bool,
    invalid_numeric: bool,
    time_series: bool,
) {
    let mut workbook = excel_writer::Workbook::new();
    let mapping = workbook
        .add_worksheet()
        .set_name("机构参照表")
        .expect("name external mapping sheet");
    mapping
        .write_string(0, 0, "机构代码")
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
    let rule_values: Vec<&str> = if include_external {
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
        for (column, value) in [
            "synthetic-v1",
            "测试",
            "2026-10-25",
            "",
            "",
            "Rust/Python V_PRE 对拍",
        ]
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
    let mut command = Command::new(env!("CARGO_BIN_EXE_open-data-audit"));
    command
        .arg("s5-audit-run")
        .arg("--config")
        .arg(config)
        .arg("--source")
        .arg(current)
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
fn s5_audit_cli_fails_closed_when_dynamic_external_layout_has_unmigrated_warnings() {
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
    assert_eq!(summary["status_counts"]["RETRIEVAL_FAILED"], 1);

    fs::remove_dir_all(directory).expect("remove synthetic warning-layout files");
}

#[test]
fn s5_audit_cli_rejects_unmigrated_dynamic_external_time_series_layout() {
    let directory = temporary_directory();
    let current = directory.join("current.xlsx");
    let external = directory.join("external.xlsx");
    let config = directory.join("config.xlsx");
    let output_dir = directory.join("out");
    write_source(&current, 1500.0);
    write_external_time_series_source(&external, 0.15);
    write_config(&config, true, false);

    let output = run_cli(&config, &current, None, Some(&external), &output_dir, true);
    assert!(output.status.success());
    let summary: Value = serde_json::from_slice(&output.stdout).expect("valid CLI JSON");
    assert_eq!(summary["status_counts"]["RETRIEVAL_FAILED"], 1);
    let result_path = PathBuf::from(
        summary["output_files"]["result_workbook"]
            .as_str()
            .expect("result workbook path"),
    );
    let result_workbook = read_workbook(&result_path).expect("read time-series rejection result");
    let result_sheet = result_workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "审核结果")
        .expect("result sheet");
    let description = CellValue::from_probe_json(&result_sheet.values["I2"])
        .expect("time-series rejection message");
    assert!(matches!(description, CellValue::Text(text) if text.contains("时序布局尚未迁移")));

    fs::remove_dir_all(directory).expect("remove synthetic time-series files");
}

#[test]
fn s5_audit_cli_fails_closed_for_invalid_dynamic_external_numeric_values() {
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
    assert_eq!(summary["status_counts"]["RETRIEVAL_FAILED"], 1);
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
    let description = CellValue::from_probe_json(&result_sheet.values["I2"])
        .expect("invalid-value rejection message");
    assert!(
        matches!(description, CellValue::Text(text) if text.contains("动态取数的数据质量告警尚未迁移"))
    );

    fs::remove_dir_all(directory).expect("remove synthetic invalid-value files");
}
