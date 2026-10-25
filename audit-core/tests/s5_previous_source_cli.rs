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

fn write_config(path: &Path) {
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
    let rule_values = [
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
    ];
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
        .arg(if previous.is_some() {
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
    write_config(&config);

    let output = run_cli(&config, &current, Some(&previous), &output_dir, true);
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
