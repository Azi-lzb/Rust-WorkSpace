use excel_reader::read_workbook;
use excel_writer::Workbook;
use serde_json::Value;
use std::{
    fs,
    path::{Path, PathBuf},
    process::{Command, Output},
    time::{SystemTime, UNIX_EPOCH},
};

struct TemporaryRoot(PathBuf);

impl TemporaryRoot {
    fn new() -> Self {
        let nonce = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("system clock after Unix epoch")
            .as_nanos();
        let path = std::env::temp_dir().join(format!(
            "audit_s3_execution_compare_cli_{}_{}",
            std::process::id(),
            nonce
        ));
        fs::create_dir_all(&path).expect("create temporary test directory");
        Self(path)
    }

    fn path(&self, name: &str) -> PathBuf {
        self.0.join(name)
    }
}

impl Drop for TemporaryRoot {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

fn write_sheet(workbook: &mut Workbook, name: &str, headers: &[&str], rows: &[&[&str]]) {
    let worksheet = workbook.add_worksheet();
    worksheet.set_name(name).expect("set config worksheet name");
    for (column, header) in headers.iter().enumerate() {
        worksheet
            .write_string(0, column as u16, *header)
            .expect("write config header");
    }
    for (row_index, row) in rows.iter().enumerate() {
        for (column, value) in row.iter().enumerate() {
            worksheet
                .write_string(row_index as u32 + 1, column as u16, *value)
                .expect("write config value");
        }
    }
}

fn write_configs(common_path: &Path, feature_path: &Path) {
    let mut common = Workbook::new();
    write_sheet(
        &mut common,
        "运行参数",
        &["参数", "值"],
        &[&["默认源数据单位", "元"], &["默认目标单位", "亿元"]],
    );
    write_sheet(&mut common, "单位换算例外", &["指标代码", "启用"], &[]);
    write_sheet(
        &mut common,
        "环比警戒",
        &[
            "下限",
            "上限",
            "备注文字",
            "填充颜色",
            "是否整行填充",
            "金额变动阈值",
            "是否说明",
        ],
        &[],
    );
    common.save(common_path).expect("save synthetic 3.0 config");

    let mut feature = Workbook::new();
    write_sheet(
        &mut feature,
        "指标参照",
        &["指标代码", "表单代码", "行序号", "指标别名"],
        &[&["I001", "F001", "1", "合成指标"]],
    );
    write_sheet(
        &mut feature,
        "规则动作",
        &[
            "规则编号",
            "规则说明",
            "规则动作",
            "指标代码",
            "数据属性",
            "币种",
            "比较指标代码",
            "人民币阈值(亿元)",
            "美元阈值(亿美元)",
            "变幅阈值(%)",
            "来源分组",
            "适用表单",
            "适用频度",
            "适用批次",
            "适用场景",
            "详细说明",
            "启用",
            "停用原因",
            "备注",
        ],
        &[],
    );
    write_sheet(
        &mut feature,
        "表达式校验",
        &[
            "规则编号",
            "规则说明",
            "校验表达式",
            "是否取反",
            "容差值(万元)",
            "启用",
            "停用原因",
        ],
        &[],
    );
    feature
        .save(feature_path)
        .expect("save synthetic 3.1 config");
}

fn write_expression_configs(common_path: &Path, feature_path: &Path) {
    write_configs(common_path, feature_path);

    let mut feature = Workbook::new();
    write_sheet(
        &mut feature,
        "指标参照",
        &["指标代码", "表单代码", "行序号", "指标别名"],
        &[&["I001", "F001", "1", "合成指标"]],
    );
    write_sheet(
        &mut feature,
        "规则动作",
        &[
            "规则编号",
            "规则说明",
            "规则动作",
            "指标代码",
            "数据属性",
            "币种",
            "比较指标代码",
            "人民币阈值(亿元)",
            "美元阈值(亿美元)",
            "变幅阈值(%)",
            "来源分组",
            "适用表单",
            "适用频度",
            "适用批次",
            "适用场景",
            "详细说明",
            "启用",
            "停用原因",
            "备注",
        ],
        &[],
    );
    let legacy_expression = "[C001,,,I001,余额,CNY,月报,第1批]>{C001,,,I001,余额,CNY,月报,第1批}";
    write_sheet(
        &mut feature,
        "表达式校验（兼容）",
        &[
            "规则编号",
            "规则说明",
            "校验表达式",
            "是否取反",
            "容差值(万元)",
            "启用",
            "停用原因",
        ],
        &[&[
            "R-EXPR-001",
            "本期值大于上期值时提示",
            legacy_expression,
            "否",
            "",
            "是",
            "",
        ]],
    );
    write_sheet(
        &mut feature,
        "表达式校验",
        &[
            "规则编号",
            "规则说明",
            "校验表达式",
            "是否取反",
            "容差值(万元)",
            "启用",
            "停用原因",
            "机构类代码",
            "地区代码",
        ],
        &[&[
            "R-EXPR-001",
            "本期值大于上期值时提示",
            "[I001,余额,CNY,月报,第1批]>{I001,余额,CNY,月报,第1批}",
            "否",
            "",
            "是",
            "",
            "^ORG1$",
            "^R001$",
        ]],
    );
    feature
        .save(feature_path)
        .expect("save expression comparison config");
}

fn write_data(path: &Path, record_date: &str, value: &str) {
    let csv = format!(
        "业务类,数据日期,机构类代码,机构类名称,地区代码,地区名称,指标顺序码,指标代码,指标名称,数据属性,币种,频度,批次,数据值\nC001,{record_date},ORG1,示例机构,R001,示例地区,1,I001,合成指标,余额,CNY,月报,第1批,{value}\n"
    );
    fs::write(path, csv).expect("write synthetic central CSV");
}

fn run_cli(
    current: &Path,
    previous: &Path,
    common_config: &Path,
    feature_config: &Path,
    output: &Path,
) -> Output {
    run_cli_with_schema(
        current,
        previous,
        common_config,
        feature_config,
        output,
        Some("five-segment"),
    )
}

fn run_cli_with_schema(
    current: &Path,
    previous: &Path,
    common_config: &Path,
    feature_config: &Path,
    output: &Path,
    schema: Option<&str>,
) -> Output {
    let mut command = Command::new(env!("CARGO_BIN_EXE_open-data-audit"));
    command
        .arg("s3-execution-compare")
        .arg("--current-data")
        .arg(current)
        .arg("--previous-data")
        .arg(previous)
        .arg("--common-config")
        .arg(common_config)
        .arg("--feature-config")
        .arg(feature_config)
        .arg("--output-workbook")
        .arg(output);
    if let Some(schema) = schema {
        command.arg("--expression-schema").arg(schema);
    }
    command
        .arg("--developer-view")
        .output()
        .expect("run S3-F01 CLI")
}

#[test]
fn cli_runs_configured_comparison_writes_result_and_refuses_overwrite() {
    let root = TemporaryRoot::new();
    let current = root.path("current.csv");
    let previous = root.path("previous.csv");
    let common_config = root.path("common.xlsx");
    let feature_config = root.path("feature.xlsx");
    let output = root.path("comparison.xlsx");
    write_data(&current, "2026-08-31", "1000");
    write_data(&previous, "2026-07-31", "900");
    write_configs(&common_config, &feature_config);

    let inputs_before = [
        fs::read(&current).expect("read current input"),
        fs::read(&previous).expect("read previous input"),
        fs::read(&common_config).expect("read common config"),
        fs::read(&feature_config).expect("read feature config"),
    ];
    let result = run_cli(
        &current,
        &previous,
        &common_config,
        &feature_config,
        &output,
    );
    assert!(
        result.status.success(),
        "{}",
        String::from_utf8_lossy(&result.stderr)
    );
    assert!(output.is_file(), "S3-F01 should create its result workbook");
    assert!(
        fs::metadata(&output).expect("stat result workbook").len() > 0,
        "result workbook should not be empty"
    );
    let summary: Value = serde_json::from_slice(&result.stdout).expect("parse CLI summary");
    assert_eq!(summary["status"], "completed");
    assert_eq!(summary["feature"], "S3-F01");
    assert_eq!(summary["rows"], 1);
    assert_eq!(summary["engines"]["config_reader"], "calamine");
    assert_eq!(summary["engines"]["writer"], "rust_xlsxwriter");

    let events = String::from_utf8_lossy(&result.stderr)
        .lines()
        .map(|line| serde_json::from_str::<Value>(line).expect("parse progress event"))
        .collect::<Vec<_>>();
    assert!(events.iter().any(|event| event["status"] == "started"));
    assert!(events.iter().any(|event| event["status"] == "completed"));
    assert!(
        events
            .iter()
            .any(|event| event["step"] == "读取本期/上期导数")
    );

    let inputs_after = [
        fs::read(&current).expect("reread current input"),
        fs::read(&previous).expect("reread previous input"),
        fs::read(&common_config).expect("reread common config"),
        fs::read(&feature_config).expect("reread feature config"),
    ];
    assert_eq!(inputs_after, inputs_before, "inputs must remain unchanged");

    let output_before = fs::read(&output).expect("read published result");
    let overwrite = run_cli(
        &current,
        &previous,
        &common_config,
        &feature_config,
        &output,
    );
    assert!(
        !overwrite.status.success(),
        "existing output must be rejected"
    );
    assert_eq!(
        fs::read(&output).expect("reread published result"),
        output_before,
        "refused rerun must preserve the existing output"
    );
}

#[test]
fn cli_evaluates_legacy_eight_and_five_segment_rules_and_writes_trigger_result() {
    let root = TemporaryRoot::new();
    let current_normal = root.path("current_normal.csv");
    let current_exception = root.path("current_exception.csv");
    let previous = root.path("previous.csv");
    let common_config = root.path("common.xlsx");
    let feature_config = root.path("feature.xlsx");
    write_data(&current_normal, "2026-08-31", "100");
    write_data(&current_exception, "2026-08-31", "120");
    write_data(&previous, "2026-07-31", "100");
    write_expression_configs(&common_config, &feature_config);

    for (schema, explicit_schema) in [
        ("legacy-eight-segment", None),
        ("five-segment", Some("five-segment")),
    ] {
        for (case, current, expected_triggers) in [
            ("normal", &current_normal, 0),
            ("exception", &current_exception, 1),
        ] {
            let output = root.path(&format!("comparison_{schema}_{case}.xlsx"));
            let result = run_cli_with_schema(
                current,
                &previous,
                &common_config,
                &feature_config,
                &output,
                explicit_schema,
            );
            assert!(
                result.status.success(),
                "schema={schema}, case={case}: {}",
                String::from_utf8_lossy(&result.stderr)
            );
            let summary: Value = serde_json::from_slice(&result.stdout)
                .expect("parse S3-F01 expression CLI summary");
            assert_eq!(summary["status"], "completed");
            assert_eq!(summary["expression_schema"], schema);
            assert_eq!(summary["rows"], 1);
            assert_eq!(summary["enabled_expression_rules"], 1);
            assert_eq!(summary["triggered_expression_rules"], expected_triggers);
            assert_eq!(summary["expression_errors"], 0);

            let workbook = read_workbook(&output).expect("reopen expression result workbook");
            assert_eq!(workbook.sheets.len(), 1);
            assert_eq!(workbook.sheets[0].name, "2026年8月对比结果");
            let explanation = workbook.sheets[0]
                .values
                .get("S2")
                .and_then(|value| value.get("value"))
                .and_then(Value::as_str)
                .unwrap_or_default();
            if expected_triggers == 1 {
                assert!(explanation.contains("本期值大于上期值时提示"));
            } else {
                assert!(explanation.is_empty());
            }
        }
    }
}
