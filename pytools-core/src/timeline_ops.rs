//! Rust implementations of the selected timeline menu operations (1-1, 1-3,
//! 1-4 and 1-5). Timeline extraction (1-2) intentionally remains outside this
//! migration slice.

use crate::{
    config::{load_global, GlobalConfig, PathMapRule, TimelineRule},
    dedup_ops::{
        write_append_workbook_payload, AppendCell, AppendTargetPlan, NewAppendSheet,
        NewAppendWorkbook,
    },
    rule_engine::{pick_rules_for_workbook, SheetValue},
    workbook_extract::{matrix_from_dump, WorkbookCell},
};
use anyhow::{anyhow, bail, Context, Result};
use chrono::Local;
use excel_editor::{
    inspect_xlsx_profile, CellValue, EditOperation, ExcelEditPlan, PreservationRequirement,
    RangeStyleMode, WorkbookFormat,
};
use excel_reader::{read_workbook, SheetDump};
use excel_writer::{Format, FormatAlign, FormatBorder, Workbook};
use serde::Serialize;
use std::{
    collections::{BTreeMap, BTreeSet, HashMap, HashSet},
    fs,
    path::{Path, PathBuf},
    time::{Duration, Instant},
};

#[derive(Clone, Debug, Serialize)]
pub struct PrecheckIssue {
    pub level: String,
    pub location: String,
    pub message: String,
}

#[derive(Clone, Debug, Serialize)]
pub struct PrecheckReport {
    pub errors: usize,
    pub warnings: usize,
    pub issues: Vec<PrecheckIssue>,
}

#[derive(Clone, Debug, Default, Serialize)]
pub struct WideSummaryProgress {
    pub total_inputs: usize,
    pub processed_inputs: usize,
    pub completed_inputs: usize,
    pub failed_inputs: usize,
    pub current_input: Option<usize>,
    pub matched_rule_sheets: usize,
    pub extracted_cells: usize,
    pub output_rows: usize,
    pub target_append_plans: usize,
}

#[derive(Clone, Debug, Serialize)]
pub struct WideSummaryRunEvent {
    pub sequence: usize,
    pub timestamp: String,
    pub feature_id: String,
    pub phase: String,
    pub status: String,
    pub user_message: String,
    pub progress: WideSummaryProgress,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub input_index: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stage_elapsed_ms: Option<u128>,
    pub total_elapsed_ms: u128,
    pub components: Vec<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub actual_engine: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error_type: Option<String>,
}

struct WideSummaryProgressEmitter<'a> {
    started: Instant,
    sequence: usize,
    feature_id: String,
    progress: WideSummaryProgress,
    callback: Option<&'a mut dyn FnMut(WideSummaryRunEvent)>,
}

impl WideSummaryProgressEmitter<'_> {
    fn emit(
        &mut self,
        phase: &str,
        status: &str,
        user_message: impl Into<String>,
        input_index: Option<usize>,
        stage_elapsed: Option<Duration>,
        components: &[&str],
        actual_engine: Option<&str>,
        error_type: Option<&str>,
    ) {
        if self.callback.is_none() {
            return;
        }
        self.sequence += 1;
        let mut recorded_components = Vec::with_capacity(components.len() + 1);
        recorded_components.push("Rust".to_owned());
        recorded_components.extend(components.iter().map(|component| (*component).to_owned()));
        let event = WideSummaryRunEvent {
            sequence: self.sequence,
            timestamp: Local::now().to_rfc3339(),
            feature_id: self.feature_id.clone(),
            phase: phase.to_owned(),
            status: status.to_owned(),
            user_message: user_message.into(),
            progress: self.progress.clone(),
            input_index,
            stage_elapsed_ms: stage_elapsed.map(|elapsed| elapsed.as_millis()),
            total_elapsed_ms: self.started.elapsed().as_millis(),
            components: recorded_components,
            actual_engine: actual_engine.map(str::to_owned),
            error_type: error_type.map(str::to_owned),
        };
        if let Some(callback) = self.callback.as_mut() {
            callback(event);
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
struct WideRowKey {
    workbook: String,
    sheet: String,
    data_date: String,
    row_path: String,
    row_identity: String,
}

#[derive(Default)]
struct WideBucket {
    key: Option<WideRowKey>,
    row_number: String,
    set_values: Vec<Option<WorkbookCell>>,
    values: HashMap<String, WorkbookCell>,
}

#[derive(Default)]
struct WideRuleState {
    set_initialized: bool,
    set_headers: Vec<String>,
    include_row_path: bool,
    include_row_number: bool,
    columns: Vec<String>,
    seen_columns: HashSet<String>,
    buckets: Vec<WideBucket>,
    bucket_indexes: HashMap<WideRowKey, usize>,
}

struct WideTargetSheetState {
    headers: Vec<String>,
    initialized: bool,
    existed: bool,
    seen_keys: HashMap<Vec<String>, HashSet<Vec<String>>>,
    rows: Vec<Vec<AppendCell>>,
    appended_rows: usize,
}

const PRECHECK_COLUMN_REQUIREMENTS: &[(&str, &[&str])] = &[
    (
        "时序提取规则",
        &[
            "是否启用",
            "规则名称",
            "工作簿关键字",
            "工作表关键字",
            "行头列",
            "列表头行",
            "必含列头",
            "必含行头",
            "数据起始行",
            "数据结束行",
            "数据起始列",
            "数据结束列",
            "跳过关键字",
            "目标工作簿路径",
            "目标工作表",
        ],
    ),
    (
        "路径标准化映射",
        &[
            "是否启用",
            "映射名称",
            "适用规则名",
            "工作簿关键字",
            "工作表关键字",
            "作用对象",
            "匹配方式",
            "原始路径",
            "标准路径",
            "备注",
        ],
    ),
    (
        "打印配置",
        &[
            "是否启用",
            "打印模式",
            "源工作簿",
            "源工作表",
            "源工作表打印区域",
            "目标工作簿",
            "目标工作表",
            "FitToPagesWide",
            "FitToPagesTall",
            "打印方向",
            "水平居中",
            "垂直居中",
            "不输出批注",
            "零值不输出",
            "备注",
        ],
    ),
    (
        "重命名配置",
        &[
            "简称",
            "全称",
            "占位C",
            "代码",
            "全称(代码映射)",
            "占位F",
            "键",
            "值",
            "占位I",
            "原表名",
            "新表名",
            "占位L",
            "原文件名片段",
            "新文件名片段",
            "内容重命名启用",
            "内容重命名取值规则",
            "备注",
        ],
    ),
    (
        "归档类型配置",
        &[
            "是否启用",
            "匹配关键词",
            "归档文件夹",
            "备注",
            "需要归档的后缀",
            "排除归档的后缀",
        ],
    ),
    (
        "提交检查配置",
        &[
            "是否启用",
            "日期关键词",
            "县区关键词",
            "机构关键词",
            "备用关键词",
            "命中次数",
            "命中的文件名",
            "备注",
        ],
    ),
    (
        "文件排序配置",
        &["是否启用", "排序前缀", "匹配关键词", "备注"],
    ),
    (
        "机构映射表",
        &["原始机构名称", "映射后机构名称", "是否为外资行"],
    ),
    (
        "工作表提取",
        &[
            "是否启用",
            "工作簿关键字",
            "工作表关键字",
            "排除关键字",
            "提取区域",
            "多区域拼接方式",
            "输出文件夹",
            "文件名",
            "备注",
        ],
    ),
];

/// Validate the configuration sheets and all enabled timeline/path-map rules.
/// This is a read-only operation; it only creates configured output/log folders
/// after required configuration checks have passed.
pub fn run_timeline_precheck(config_path: impl AsRef<Path>) -> Result<PrecheckReport> {
    let config_path = config_path.as_ref();
    let mut report = PrecheckReport {
        errors: 0,
        warnings: 0,
        issues: Vec::new(),
    };
    macro_rules! add {
        ($level:expr, $location:expr, $message:expr $(,)?) => {{
            let level = $level;
            if level == "error" {
                report.errors += 1;
            } else {
                report.warnings += 1;
            }
            report.issues.push(PrecheckIssue {
                level: level.to_owned(),
                location: $location,
                message: $message,
            });
        }};
    }
    if !config_path.is_file() {
        add!(
            "error",
            config_path.display().to_string(),
            "配置文件不存在".to_owned(),
        );
        return Ok(report);
    }
    let workbook = match read_workbook(config_path) {
        Ok(value) => value,
        Err(error) => {
            add!(
                "error",
                config_path.display().to_string(),
                format!("无法打开配置: {error:#}"),
            );
            return Ok(report);
        }
    };
    let sheet_names = workbook
        .sheets
        .iter()
        .map(|sheet| sheet.name.as_str())
        .collect::<BTreeSet<_>>();
    for required in ["全局配置", "时序提取规则", "路径标准化映射"] {
        if !sheet_names.contains(required) {
            add!("error", required.to_owned(), "缺少 Sheet".to_owned());
        }
    }
    for (sheet_name, required_headers) in PRECHECK_COLUMN_REQUIREMENTS {
        if let Some(sheet) = workbook
            .sheets
            .iter()
            .find(|sheet| sheet.name == *sheet_name)
        {
            let headers = precheck_headers(sheet);
            for required in *required_headers {
                if !headers.contains_key(*required) {
                    add!(
                        "error",
                        sheet_name.to_string(),
                        format!("缺少列: {required}"),
                    );
                }
            }
        }
    }
    if let Some(global) = workbook
        .sheets
        .iter()
        .find(|sheet| sheet.name == "全局配置")
    {
        let max_column = global
            .values
            .keys()
            .filter_map(|address| parse_a1(address).map(|(_, column)| column))
            .max()
            .unwrap_or(0);
        if max_column < 2 {
            add!(
                "error",
                "全局配置".to_owned(),
                "全局配置 至少需要 键/值 两列".to_owned(),
            );
        } else {
            let keys = global
                .values
                .iter()
                .filter_map(|(cell, value)| {
                    let (row, column) = parse_a1(cell)?;
                    (column == 1 && row > 1).then(|| value_text(value))
                })
                .filter(|value| !value.is_empty())
                .collect::<BTreeSet<_>>();
            for key in ["输入目录", "输出目录", "日志目录"] {
                if !keys.contains(key) {
                    add!("error", "全局配置".to_owned(), format!("缺少键: {key}"));
                }
            }
        }
    }
    if report.errors == 0 {
        match load_global(config_path) {
            Ok(global) => {
                if !global.input_dir.exists() {
                    add!(
                        "error",
                        "全局配置/输入目录".to_owned(),
                        format!("输入目录不存在: {}", global.input_dir.display()),
                    );
                }
                if let Err(error) = fs::create_dir_all(&global.output_dir) {
                    add!(
                        "error",
                        "全局配置/输出目录".to_owned(),
                        format!("无法创建目录: {error}"),
                    );
                }
                if let Err(error) = fs::create_dir_all(&global.log_dir) {
                    add!(
                        "error",
                        "全局配置/日志目录".to_owned(),
                        format!("无法创建目录: {error}"),
                    );
                }
            }
            Err(error) => add!(
                "error",
                "全局配置".to_owned(),
                format!("解析失败: {error:#}"),
            ),
        }
        if let Some(sheet) = workbook
            .sheets
            .iter()
            .find(|sheet| sheet.name == "时序提取规则")
        {
            let headers = precheck_headers(sheet);
            let write_flag_column_exists = headers.contains_key("启用目标写入");
            let set_column_exists = headers.contains_key("set区域");
            let mut enabled_names = BTreeSet::new();
            let mut seen_names = BTreeSet::new();
            let mut duplicate_names = BTreeSet::new();
            for row in 2..=precheck_max_row(sheet) {
                if !precheck_truthy(&precheck_field(sheet, &headers, row, "是否启用")) {
                    continue;
                }
                let location = format!("时序提取规则/第{row}行");
                let name = precheck_field(sheet, &headers, row, "规则名称");
                let workbook_keyword = precheck_field(sheet, &headers, row, "工作簿关键字");
                let worksheet_keyword = precheck_field(sheet, &headers, row, "工作表关键字");
                let row_header_raw = precheck_field(sheet, &headers, row, "行头列");
                let row_header_column = crate::path_map::parse_col_spec(&row_header_raw);
                let column_header_raw = precheck_field(sheet, &headers, row, "列表头行");
                let data_row_start = precheck_field(sheet, &headers, row, "数据起始行");
                let data_row_end = precheck_field(sheet, &headers, row, "数据结束行");
                let data_column_start = precheck_field(sheet, &headers, row, "数据起始列");
                let data_column_end = precheck_field(sheet, &headers, row, "数据结束列");

                if name.is_empty() {
                    add!("error", location.clone(), "规则名称为空".to_owned());
                } else {
                    if !seen_names.insert(name.clone()) {
                        duplicate_names.insert(name.clone());
                    }
                    enabled_names.insert(name.clone());
                }
                if workbook_keyword.is_empty() {
                    add!("error", location.clone(), "工作簿关键字为空".to_owned());
                }
                if worksheet_keyword.is_empty() {
                    add!("error", location.clone(), "工作表关键字为空".to_owned());
                }
                if !row_header_raw.is_empty() {
                    if row_header_column == 0 {
                        add!("error", location.clone(), "行头列非法".to_owned());
                    }
                } else {
                    if data_row_start.is_empty() {
                        add!(
                            "error",
                            location.clone(),
                            "行头列为空时，数据起始行必填".to_owned()
                        );
                    }
                    if crate::path_map::parse_col_spec(&data_column_start) <= 0 {
                        add!(
                            "error",
                            location.clone(),
                            "行头列为空时，数据起始列必填且合法".to_owned(),
                        );
                    }
                    if !precheck_field(sheet, &headers, row, "必含行头").is_empty() {
                        add!(
                            "warning",
                            location.clone(),
                            "行头列为空时将忽略“必含行头”".to_owned(),
                        );
                    }
                }
                let header_tokens = precheck_list_header_tokens(&column_header_raw);
                if column_header_raw.is_empty() {
                    add!("error", location.clone(), "列表头行为空".to_owned());
                } else if header_tokens.is_empty()
                    || header_tokens.iter().any(|token| {
                        precheck_int_from_float(token).map_or(true, |value| value <= 0)
                    })
                {
                    add!(
                        "error",
                        location.clone(),
                        "列表头行必须为正整数或逗号列表".to_owned(),
                    );
                }
                if !data_row_start.is_empty()
                    && precheck_int_from_float(&data_row_start).map_or(true, |value| value <= 0)
                {
                    add!("error", location.clone(), "数据起始行非法".to_owned());
                }
                if !data_row_start.is_empty() && !data_row_end.is_empty() {
                    match (
                        precheck_int_from_float(&data_row_start),
                        precheck_int_from_float(&data_row_end),
                    ) {
                        (Some(start), Some(end)) if end < start => add!(
                            "error",
                            location.clone(),
                            "数据结束行 < 数据起始行".to_owned(),
                        ),
                        (Some(_), Some(_)) => {}
                        _ => add!("error", location.clone(), "数据结束行非法".to_owned()),
                    }
                }
                let data_column_start_value = crate::path_map::parse_col_spec(&data_column_start);
                let data_column_end_value = crate::path_map::parse_col_spec(&data_column_end);
                if !data_column_start.is_empty() && data_column_start_value == 0 {
                    add!("error", location.clone(), "数据起始列非法".to_owned());
                }
                if !data_column_end.is_empty() && data_column_end_value == 0 {
                    add!("error", location.clone(), "数据结束列非法".to_owned());
                }
                if data_column_start_value > 0
                    && data_column_end_value > 0
                    && data_column_end_value < data_column_start_value
                {
                    add!(
                        "error",
                        location.clone(),
                        "数据结束列 < 数据起始列".to_owned(),
                    );
                }
                let write_enabled = if write_flag_column_exists {
                    precheck_truthy(&precheck_field(sheet, &headers, row, "启用目标写入"))
                } else {
                    true
                };
                if write_enabled
                    && (precheck_field(sheet, &headers, row, "目标工作簿路径").is_empty()
                        || precheck_field(sheet, &headers, row, "目标工作表").is_empty())
                {
                    add!(
                        "warning",
                        location.clone(),
                        "启用目标写入=是，但目标工作簿路径/目标工作表未完整填写".to_owned(),
                    );
                }
                if set_column_exists {
                    let set_raw = precheck_field(sheet, &headers, row, "set区域");
                    if let Some(error) = precheck_set_error(&set_raw) {
                        add!("error", location, error);
                    }
                }
            }
            for duplicate in duplicate_names {
                add!(
                    "warning",
                    "时序提取规则".to_owned(),
                    format!("启用规则名称重复: {duplicate}"),
                );
            }

            if let Some(sheet) = workbook
                .sheets
                .iter()
                .find(|sheet| sheet.name == "路径标准化映射")
            {
                let headers = precheck_headers(sheet);
                for row in 2..=precheck_max_row(sheet) {
                    if !precheck_truthy(&precheck_field(sheet, &headers, row, "是否启用")) {
                        continue;
                    }
                    let location = format!("路径标准化映射/第{row}行");
                    let name = precheck_field(sheet, &headers, row, "映射名称");
                    let applicable_rule = precheck_field(sheet, &headers, row, "适用规则名");
                    let target = precheck_field(sheet, &headers, row, "作用对象");
                    let target = if target.is_empty() { "两者" } else { &target };
                    let mode = precheck_field(sheet, &headers, row, "匹配方式");
                    let mode = if mode.is_empty() { "精确" } else { &mode };
                    let original = precheck_field(sheet, &headers, row, "原始路径");
                    let standard = precheck_field(sheet, &headers, row, "标准路径");
                    if name.is_empty() {
                        add!("error", location.clone(), "映射名称为空".to_owned());
                    }
                    if !matches!(target, "行头" | "列头" | "两者") {
                        add!(
                            "error",
                            location.clone(),
                            "作用对象必须为 行头/列头/两者".to_owned(),
                        );
                    }
                    if !matches!(mode, "精确" | "包含") {
                        add!(
                            "error",
                            location.clone(),
                            "匹配方式必须为 精确/包含".to_owned(),
                        );
                    }
                    if original.is_empty() {
                        add!("error", location.clone(), "原始路径为空".to_owned());
                    }
                    if standard.is_empty() {
                        add!("error", location.clone(), "标准路径为空".to_owned());
                    }
                    if !applicable_rule.is_empty()
                        && applicable_rule != "全部"
                        && applicable_rule != "*"
                        && !enabled_names.contains(&applicable_rule)
                    {
                        add!(
                            "error",
                            location,
                            format!("适用规则名[{applicable_rule}]未在启用时序规则中找到"),
                        );
                    }
                }
            }
        }
    }
    Ok(report)
}

/// Compare configured row/column header paths against a template. Output
/// contains the same two data sheets and 12-column difference contract as the
/// legacy 3.9.3 operation. The source files are never modified.
pub fn run_header_path_compare(
    template_path: impl AsRef<Path>,
    source_paths: &[PathBuf],
    rules: &[TimelineRule],
    global: &GlobalConfig,
) -> Result<PathBuf> {
    let template_path = template_path.as_ref();
    if !template_path.is_file() {
        bail!("template does not exist: {}", template_path.display());
    }
    let template = read_workbook(template_path)?;
    let mut diff_rows: Vec<Vec<String>> = Vec::new();
    let mut summary_rows: Vec<Vec<String>> = Vec::new();
    for source_path in source_paths {
        if same_path(source_path, template_path) {
            continue;
        }
        let source = match read_workbook(source_path) {
            Ok(value) => value,
            Err(error) => {
                summary_rows.push(vec![
                    file_name(source_path),
                    "".into(),
                    "0".into(),
                    "0".into(),
                    format!("打开失败: {error:#}"),
                ]);
                continue;
            }
        };
        let mut pos_count = 0usize;
        let mut path_count = 0usize;
        for rule in pick_rules_for_workbook(
            rules,
            &file_name(source_path),
            &global.timeline_rule_match_mode,
        ) {
            let source_sheets = source
                .sheets
                .iter()
                .filter(|sheet| keyword_match(&sheet.name, &rule.sheet_keyword))
                .collect::<Vec<_>>();
            let template_sheets = template
                .sheets
                .iter()
                .filter(|sheet| keyword_match(&sheet.name, &rule.sheet_keyword))
                .collect::<Vec<_>>();
            for source_sheet in source_sheets {
                let template_sheet = template
                    .sheets
                    .iter()
                    .find(|sheet| sheet.name == source_sheet.name)
                    .or_else(|| (template_sheets.len() == 1).then(|| template_sheets[0]));
                let Some(template_sheet) = template_sheet else {
                    continue;
                };
                let (template_matrix, _) = matrix_from_dump(template_sheet)?;
                let (source_matrix, _) = matrix_from_dump(source_sheet)?;
                let positional = position_differences(&template_matrix, &source_matrix, rule);
                for (kind, source_value, template_value) in &positional {
                    diff_rows.push(vec![
                        "".into(),
                        "".into(),
                        rule.name.clone(),
                        rule.wb_keyword.clone(),
                        rule.sheet_keyword.clone(),
                        kind.clone(),
                        "精确".into(),
                        source_value.clone(),
                        template_value.clone(),
                        "按位置".into(),
                        file_name(source_path),
                        source_sheet.name.clone(),
                    ]);
                }
                pos_count += positional.len();
                if !positional.is_empty() {
                    for (kind, source_value, template_value) in
                        path_differences(&template_matrix, &source_matrix, rule)
                    {
                        diff_rows.push(vec![
                            "".into(),
                            "".into(),
                            rule.name.clone(),
                            rule.wb_keyword.clone(),
                            rule.sheet_keyword.clone(),
                            kind,
                            "精确".into(),
                            source_value,
                            template_value,
                            "按路径".into(),
                            file_name(source_path),
                            source_sheet.name.clone(),
                        ]);
                        path_count += 1;
                    }
                }
            }
        }
        let status = if pos_count == 0 && path_count == 0 {
            "通过"
        } else {
            "存在差异"
        };
        summary_rows.push(vec![
            file_name(source_path),
            status.into(),
            pos_count.to_string(),
            path_count.to_string(),
            String::new(),
        ]);
    }
    let path = unique_output(
        &global.output_dir,
        &format!("compare_3_9_3_{}", Local::now().format("%Y%m%d_%H%M%S")),
        "xlsx",
    )?;
    let mut workbook = Workbook::new();
    write_string_table(
        &mut workbook,
        "表头比对结果",
        &[
            "是否启用",
            "映射名称",
            "适用规则名",
            "工作簿关键字",
            "工作表关键字",
            "作用对象",
            "匹配方式",
            "原始路径",
            "标准路径",
            "备注",
            "源文件",
            "工作表",
        ],
        &diff_rows,
    )?;
    write_compare_summary_table(&mut workbook, &summary_rows)?;
    workbook
        .save(&path)
        .with_context(|| format!("save comparison report {}", path.display()))?;
    Ok(path)
}

/// Execute selected wide-summary variants (timeline menu 1-3 and 1-4).
pub fn run_wide_summary(
    source_paths: &[PathBuf],
    rules: &[TimelineRule],
    path_maps: &[PathMapRule],
    global: &GlobalConfig,
    row_suffix_enabled: bool,
) -> Result<Option<PathBuf>> {
    run_wide_summary_inner(
        source_paths,
        rules,
        path_maps,
        global,
        row_suffix_enabled,
        None,
    )
}

pub fn run_wide_summary_with_progress<F>(
    source_paths: &[PathBuf],
    rules: &[TimelineRule],
    path_maps: &[PathMapRule],
    global: &GlobalConfig,
    row_suffix_enabled: bool,
    mut on_event: F,
) -> Result<Option<PathBuf>>
where
    F: FnMut(WideSummaryRunEvent),
{
    run_wide_summary_inner(
        source_paths,
        rules,
        path_maps,
        global,
        row_suffix_enabled,
        Some(&mut on_event),
    )
}

fn run_wide_summary_inner(
    source_paths: &[PathBuf],
    rules: &[TimelineRule],
    path_maps: &[PathMapRule],
    global: &GlobalConfig,
    row_suffix_enabled: bool,
    callback: Option<&mut dyn FnMut(WideSummaryRunEvent)>,
) -> Result<Option<PathBuf>> {
    let feature_id = if row_suffix_enabled {
        "PT-1-3"
    } else {
        "PT-1-4"
    };
    let mut emitter = WideSummaryProgressEmitter {
        started: Instant::now(),
        sequence: 0,
        feature_id: feature_id.to_owned(),
        progress: WideSummaryProgress {
            total_inputs: source_paths.len(),
            ..WideSummaryProgress::default()
        },
        callback,
    };
    let mut rule_states = HashMap::<String, WideRuleState>::new();
    let mut rule_order = Vec::<String>::new();
    for (source_index, path) in source_paths.iter().enumerate() {
        let input_index = source_index + 1;
        emitter.progress.processed_inputs += 1;
        emitter.progress.current_input = Some(input_index);
        let read_started = Instant::now();
        emitter.emit(
            "source_read",
            "started",
            format!("正在读取第 {input_index}/{} 个工作簿", source_paths.len()),
            Some(input_index),
            None,
            &["pytools-core", "excel-reader", "calamine"],
            None,
            None,
        );
        let filename = file_name(path);
        let workbook = match read_workbook(path) {
            Ok(value) => {
                emitter.emit(
                    "source_read",
                    "completed",
                    format!("已读取第 {input_index}/{} 个工作簿", source_paths.len()),
                    Some(input_index),
                    Some(read_started.elapsed()),
                    &["pytools-core", "excel-reader", "calamine"],
                    Some("calamine"),
                    None,
                );
                value
            }
            Err(error) if global.error_policy == "continue" => {
                emitter.progress.failed_inputs += 1;
                emitter.emit(
                    "source_read",
                    "failed",
                    format!(
                        "第 {input_index}/{} 个工作簿读取失败，继续处理",
                        source_paths.len()
                    ),
                    Some(input_index),
                    Some(read_started.elapsed()),
                    &["pytools-core", "excel-reader", "calamine"],
                    Some("calamine"),
                    Some("workbook_read_error"),
                );
                let _ = error;
                continue;
            }
            Err(error) => {
                emitter.progress.failed_inputs += 1;
                emitter.emit(
                    "source_read",
                    "failed",
                    format!("第 {input_index}/{} 个工作簿读取失败", source_paths.len()),
                    Some(input_index),
                    Some(read_started.elapsed()),
                    &["pytools-core", "excel-reader", "calamine"],
                    Some("calamine"),
                    Some("workbook_read_error"),
                );
                return Err(error).with_context(|| format!("read {}", path.display()));
            }
        };
        let aggregation_started = Instant::now();
        emitter.emit(
            "rule_aggregation",
            "started",
            format!("正在汇总第 {input_index}/{} 个工作簿", source_paths.len()),
            Some(input_index),
            None,
            &["pytools-core", "rule_engine"],
            None,
            None,
        );
        for rule in pick_rules_for_workbook(rules, &filename, &global.timeline_rule_match_mode) {
            if !rule.set_parse_error.is_empty() {
                continue;
            }
            if !wide_rule_compatible(rule) {
                continue;
            }
            for sheet in workbook
                .sheets
                .iter()
                .filter(|sheet| keyword_match(&sheet.name, &rule.sheet_keyword))
            {
                emitter.progress.matched_rule_sheets += 1;
                let (matrix, merges) = match matrix_from_dump(sheet) {
                    Ok(value) => value,
                    Err(error) => {
                        emitter.progress.failed_inputs += 1;
                        emitter.emit(
                            "rule_aggregation",
                            "failed",
                            format!("第 {input_index}/{} 个工作簿汇总失败", source_paths.len()),
                            Some(input_index),
                            Some(aggregation_started.elapsed()),
                            &["pytools-core", "rule_engine"],
                            None,
                            Some("worksheet_projection_error"),
                        );
                        return Err(error).with_context(|| {
                            format!("project worksheet {:?} in input {input_index}", sheet.name)
                        });
                    }
                };
                let rule_name = if rule.name.trim().is_empty() {
                    "未命名规则".to_owned()
                } else {
                    rule.name.clone()
                };
                let state = rule_states.entry(rule_name.clone()).or_default();
                if !state.set_initialized {
                    let aliases = rule
                        .set_items
                        .iter()
                        .map(|(alias, _)| alias.clone())
                        .collect::<Vec<_>>();
                    state.set_headers = disambiguate_headers(&aliases);
                    state.set_initialized = true;
                }

                let set_values = rule
                    .set_items
                    .iter()
                    .map(|(_, address)| value_at_merged_address(&matrix, &merges, address))
                    .collect::<Vec<_>>();
                let (row_count, column_count) = matrix_shape(&matrix);
                if row_count > 0 && column_count > 0 {
                    let header_rows = rule
                        .col_header_rows
                        .iter()
                        .filter_map(|row| row.checked_sub(1))
                        .filter(|row| *row < row_count)
                        .collect::<Vec<_>>();
                    if !header_rows.is_empty() {
                        let row_header_columns = if rule.row_header_col_specified {
                            if rule.row_header_cols.is_empty() {
                                vec![rule.row_header_col]
                            } else {
                                rule.row_header_cols.clone()
                            }
                        } else {
                            Vec::new()
                        };
                        let row_start = if rule.data_row_start > 0 {
                            rule.data_row_start - 1
                        } else {
                            header_rows.iter().copied().max().unwrap_or(0) + 1
                        };
                        let col_start = if rule.data_col_start > 0 {
                            rule.data_col_start - 1
                        } else if rule.row_header_col_specified {
                            row_header_columns
                                .iter()
                                .filter_map(|column| column.checked_sub(1))
                                .max()
                                .unwrap_or(0)
                                + 1
                        } else {
                            0
                        };
                        let col_end = rule
                            .data_col_end
                            .filter(|column| *column > 0)
                            .map(|column| column - 1)
                            .unwrap_or_else(|| column_count - 1)
                            .min(column_count - 1);
                        if row_start < row_count && col_start <= col_end {
                            let (_, ordered_columns) = crate::rule_engine::compute_wide_col_paths(
                                &matrix,
                                &merges,
                                rule,
                                &sheet.name,
                                path_maps,
                                &filename,
                                col_start,
                                col_end,
                                &header_rows,
                                &HashMap::new(),
                            );
                            for column in ordered_columns {
                                if state.seen_columns.insert(column.clone()) {
                                    state.columns.push(column);
                                }
                            }
                        }
                    }
                }
                let cells = match crate::rule_engine::extract_cells_from_matrix(
                    &matrix,
                    &merges,
                    &filename,
                    rule,
                    &sheet.name,
                    path_maps,
                    row_suffix_enabled,
                    &HashMap::new(),
                ) {
                    Ok(cells) => cells,
                    Err(error) => {
                        emitter.progress.failed_inputs += 1;
                        emitter.emit(
                            "rule_aggregation",
                            "failed",
                            format!("第 {input_index}/{} 个工作簿汇总失败", source_paths.len()),
                            Some(input_index),
                            Some(aggregation_started.elapsed()),
                            &["pytools-core", "rule_engine"],
                            None,
                            Some("wide_rule_error"),
                        );
                        return Err(anyhow!(error)).with_context(|| {
                            format!(
                                "aggregate worksheet {:?} from input {input_index}",
                                sheet.name
                            )
                        });
                    }
                };
                emitter.progress.extracted_cells += cells.len();
                for cell in cells {
                    let row_number = cell
                        .cell_addr
                        .chars()
                        .rev()
                        .take_while(|ch| ch.is_ascii_digit())
                        .collect::<String>()
                        .chars()
                        .rev()
                        .collect::<String>();
                    let row_unique = if row_suffix_enabled {
                        cell.row_path.clone()
                    } else {
                        format!("{}#R{}", cell.row_path, row_number)
                    };
                    let row_path_out = if rule.row_header_col_specified {
                        cell.row_path.clone()
                    } else {
                        String::new()
                    };
                    let key = WideRowKey {
                        workbook: cell.source_file.clone(),
                        sheet: cell.sheet_name.clone(),
                        data_date: cell.data_date.clone(),
                        row_path: row_path_out,
                        row_identity: row_unique,
                    };
                    let state = rule_states
                        .get_mut(&rule_name)
                        .expect("wide rule state initialized before extraction");
                    if state.buckets.is_empty() && !rule_order.contains(&rule_name) {
                        rule_order.push(rule_name.clone());
                        state.include_row_path = rule.row_header_col_specified;
                        state.include_row_number = rule.output_row_no;
                    }
                    let bucket_index = if let Some(index) = state.bucket_indexes.get(&key) {
                        *index
                    } else {
                        let index = state.buckets.len();
                        state.bucket_indexes.insert(key.clone(), index);
                        state.buckets.push(WideBucket {
                            key: Some(key),
                            row_number: row_number.clone(),
                            set_values: set_values.clone(),
                            values: HashMap::new(),
                        });
                        index
                    };
                    let bucket = &mut state.buckets[bucket_index];
                    bucket.row_number = row_number;
                    if !bucket.values.contains_key(&cell.col_path) {
                        if state.seen_columns.insert(cell.col_path.clone()) {
                            state.columns.push(cell.col_path.clone());
                        }
                        bucket.values.insert(cell.col_path, cell.value);
                    }
                }
            }
        }
        emitter.progress.completed_inputs += 1;
        emitter.emit(
            "rule_aggregation",
            "completed",
            format!(
                "已处理 {}/{} 个工作簿，匹配 {} 个工作表",
                emitter.progress.completed_inputs,
                source_paths.len(),
                emitter.progress.matched_rule_sheets
            ),
            Some(input_index),
            Some(aggregation_started.elapsed()),
            &["pytools-core", "rule_engine"],
            None,
            None,
        );
    }
    if rule_order.is_empty() {
        emitter.emit(
            "summary_write",
            "skipped",
            "没有匹配的宽表数据，未生成汇总工作簿",
            None,
            None,
            &["pytools-core"],
            None,
            None,
        );
        return Ok(None);
    }
    emitter.progress.output_rows = rule_order
        .iter()
        .filter_map(|rule_name| rule_states.get(rule_name))
        .map(|state| state.buckets.len())
        .sum();
    let stamp = Local::now().format("%Y%m%d_%H%M%S");
    let label = if row_suffix_enabled {
        "宽表规则汇总（行头加后缀）"
    } else {
        "宽表规则汇总（行头不加后缀）"
    };
    let path = match unique_output(&global.output_dir, &format!("{stamp}_{label}"), "xlsx") {
        Ok(path) => path,
        Err(error) => {
            emitter.emit(
                "summary_write",
                "failed",
                "无法准备宽表汇总输出文件",
                None,
                None,
                &["pytools-core", "excel-writer"],
                Some("rust_xlsxwriter"),
                Some("output_path_error"),
            );
            return Err(error);
        }
    };
    emitter.emit(
        "summary_write",
        "started",
        "正在写入宽表汇总工作簿",
        None,
        None,
        &["pytools-core", "excel-writer", "rust_xlsxwriter"],
        None,
        None,
    );
    let write_started = Instant::now();
    let write_result = (|| -> Result<()> {
        let mut workbook = Workbook::new();
        let mut used_sheet_names = HashSet::new();
        for rule_name in &rule_order {
            let state = rule_states
                .get(rule_name)
                .expect("ordered wide rule state exists");
            let mut headers = vec![
                "工作簿名".to_owned(),
                "工作表名".to_owned(),
                "数据日期".to_owned(),
            ];
            headers.extend(state.set_headers.iter().cloned());
            if state.include_row_path {
                headers.push("行头路径".into());
            }
            if state.include_row_number {
                headers.push("行序号".into());
            }
            headers.extend(state.columns.iter().cloned());
            let sheet = workbook.add_worksheet();
            sheet.set_name(&unique_wide_sheet_name(rule_name, &mut used_sheet_names))?;
            let header_format = Format::new()
                .set_bold()
                .set_border(FormatBorder::Thin)
                .set_align(FormatAlign::Center)
                .set_align(FormatAlign::Top);
            for (column, header) in headers.iter().enumerate() {
                sheet.write_string_with_format(0, column as u16, header, &header_format)?;
            }
            for (row_index, bucket) in state.buckets.iter().enumerate() {
                let key = bucket.key.as_ref().expect("wide bucket has a row key");
                let row = (row_index + 1) as u32;
                let mut column = 0u16;
                sheet.write_string(row, column, &key.workbook)?;
                column += 1;
                sheet.write_string(row, column, &key.sheet)?;
                column += 1;
                write_wide_data_date(sheet, row, column, &key.data_date)?;
                column += 1;
                for value in &bucket.set_values {
                    write_wide_cell(sheet, row, column, value.as_ref())?;
                    column += 1;
                }
                if state.include_row_path {
                    sheet.write_string(row, column, &key.row_path)?;
                    column += 1;
                }
                if state.include_row_number {
                    sheet.write_string(row, column, &bucket.row_number)?;
                    column += 1;
                }
                for path in &state.columns {
                    write_wide_cell(sheet, row, column, bucket.values.get(path))?;
                    column += 1;
                }
            }
        }
        workbook
            .save(&path)
            .with_context(|| format!("save wide summary {}", path.display()))?;
        Ok(())
    })();
    if let Err(error) = write_result {
        emitter.emit(
            "summary_write",
            "failed",
            "写入宽表汇总工作簿失败",
            None,
            Some(write_started.elapsed()),
            &["pytools-core", "excel-writer", "rust_xlsxwriter"],
            Some("rust_xlsxwriter"),
            Some("summary_write_error"),
        );
        return Err(error);
    }
    emitter.emit(
        "summary_write",
        "completed",
        format!(
            "宽表汇总工作簿已写入，共 {} 行",
            emitter.progress.output_rows
        ),
        None,
        Some(write_started.elapsed()),
        &["pytools-core", "excel-writer", "rust_xlsxwriter"],
        Some("rust_xlsxwriter"),
        None,
    );
    Ok(Some(path))
}

fn wide_rule_compatible(rule: &TimelineRule) -> bool {
    rule.row_header_col_specified || (rule.data_row_start > 0 && rule.data_col_start > 0)
}

fn disambiguate_headers(names: &[String]) -> Vec<String> {
    let mut counts = HashMap::<String, usize>::new();
    names
        .iter()
        .map(|name| {
            let count = counts.entry(name.clone()).or_default();
            *count += 1;
            if *count == 1 {
                name.clone()
            } else {
                format!("{name}_{count}")
            }
        })
        .collect()
}

fn unique_wide_sheet_name(name: &str, used: &mut HashSet<String>) -> String {
    let base = crate::sheet_extract::sanitize_sheet_name(name);
    if used.insert(base.clone()) {
        return base;
    }
    for suffix in 2usize.. {
        let tail = format!("_{suffix}");
        let keep = 31usize.saturating_sub(tail.chars().count());
        let candidate = format!("{}{}", base.chars().take(keep).collect::<String>(), tail);
        if used.insert(candidate.clone()) {
            return candidate;
        }
    }
    unreachable!("wide summary worksheet names are bounded by Excel's sheet limit")
}

fn write_wide_data_date(
    sheet: &mut excel_writer::Worksheet,
    row: u32,
    column: u16,
    text: &str,
) -> Result<()> {
    if let Some(serial) = wide_datetime_serial(text) {
        let format = excel_writer::Format::new().set_num_format("yyyy/m/d");
        sheet.write_number_with_format(row, column, serial, &format)?;
    } else {
        sheet.write_string(row, column, text)?;
    }
    Ok(())
}

fn matrix_shape(matrix: &[Vec<Option<WorkbookCell>>]) -> (usize, usize) {
    (matrix.len(), matrix.iter().map(Vec::len).max().unwrap_or(0))
}

fn value_at_merged_address(
    matrix: &[Vec<Option<WorkbookCell>>],
    merges: &[crate::rule_engine::MergeRange],
    address: &str,
) -> Option<WorkbookCell> {
    let (row, column) = parse_a1(address)?;
    let (row, column) = (row.checked_sub(1)?, column.checked_sub(1)?);
    if let Some(value) = matrix
        .get(row)
        .and_then(|cells| cells.get(column))
        .and_then(Option::as_ref)
    {
        return Some(value.clone());
    }
    let merge = merges.iter().find(|merge| {
        merge.min_row <= row
            && row <= merge.max_row
            && merge.min_col <= column
            && column <= merge.max_col
    })?;
    matrix
        .get(merge.min_row)
        .and_then(|cells| cells.get(merge.min_col))
        .and_then(Option::as_ref)
        .cloned()
}

fn write_wide_cell(
    sheet: &mut excel_writer::Worksheet,
    row: u32,
    column: u16,
    value: Option<&WorkbookCell>,
) -> Result<()> {
    use WorkbookCell as Cell;
    let Some(value) = value else {
        return Ok(());
    };
    match value {
        Cell::Text(text) | Cell::Error(text) => {
            sheet.write_string(row, column, text)?;
        }
        Cell::Integer(number) => {
            sheet.write_number(row, column, *number as f64)?;
        }
        Cell::Number(number) => {
            sheet.write_number(row, column, *number)?;
        }
        Cell::Boolean(boolean) => {
            sheet.write_boolean(row, column, *boolean)?;
        }
        Cell::DateTimeSerial(serial) => {
            let format = wide_date_format(*serial);
            sheet.write_number_with_format(row, column, *serial, &format)?;
        }
        Cell::DateTimeIso(text) => {
            if let Some(serial) = wide_datetime_serial(text) {
                let format = wide_date_format(serial);
                sheet.write_number_with_format(row, column, serial, &format)?;
            } else {
                sheet.write_string(row, column, text)?;
            }
        }
        Cell::DurationIso(text) => {
            if let Some(serial) = wide_duration_serial(text) {
                let format = excel_writer::Format::new().set_num_format("[h]:mm:ss");
                sheet.write_number_with_format(row, column, serial, &format)?;
            } else {
                sheet.write_string(row, column, text)?;
            }
        }
    }
    Ok(())
}

fn wide_date_format(serial: f64) -> excel_writer::Format {
    excel_writer::Format::new().set_num_format(wide_date_format_code(serial))
}

fn wide_date_format_code(serial: f64) -> &'static str {
    if serial.fract().abs() < f64::EPSILON {
        "yyyy-mm-dd"
    } else {
        "yyyy-mm-dd hh:mm:ss"
    }
}

fn wide_datetime_serial(text: &str) -> Option<f64> {
    use chrono::{DateTime, NaiveDate, NaiveDateTime};
    let datetime = DateTime::parse_from_rfc3339(text)
        .map(|value| value.naive_utc())
        .or_else(|_| NaiveDateTime::parse_from_str(text, "%Y-%m-%dT%H:%M:%S%.f"))
        .or_else(|_| NaiveDateTime::parse_from_str(text, "%Y-%m-%d %H:%M:%S%.f"))
        .or_else(|_| {
            NaiveDate::parse_from_str(text, "%Y-%m-%d")
                .map(|date| date.and_hms_opt(0, 0, 0).expect("valid midnight"))
        })
        .ok()?;
    let epoch = NaiveDate::from_ymd_opt(1899, 12, 30)?.and_hms_opt(0, 0, 0)?;
    Some(datetime.signed_duration_since(epoch).num_milliseconds() as f64 / 86_400_000.0)
}

fn wide_duration_serial(text: &str) -> Option<f64> {
    let text = text.strip_prefix('P')?;
    let mut in_time = false;
    let mut number = String::new();
    let mut seconds = 0.0;
    for character in text.chars() {
        if character == 'T' {
            in_time = true;
            continue;
        }
        if character.is_ascii_digit() || character == '.' || character == '-' {
            number.push(character);
            continue;
        }
        let amount = number.parse::<f64>().ok()?;
        number.clear();
        let multiplier = match (in_time, character) {
            (false, 'D') => 86_400.0,
            (true, 'H') => 3_600.0,
            (true, 'M') => 60.0,
            (true, 'S') => 1.0,
            _ => return None,
        };
        seconds += amount * multiplier;
    }
    if number.is_empty() {
        Some(seconds / 86_400.0)
    } else {
        None
    }
}

/// Build the configured target-append work as one file-level plan per target.
/// Existing targets remain untouched; their plans must go through
/// `dedup-append-execute` with matching ExcelModifier capability evidence.
pub fn build_wide_target_append_plans(
    summary_path: &Path,
    rules: &[TimelineRule],
    output_dir: &Path,
) -> Result<Vec<AppendTargetPlan>> {
    let summary = read_workbook(summary_path)
        .with_context(|| format!("read wide-summary output {}", summary_path.display()))?;
    let mut grouped = BTreeMap::<PathBuf, Vec<&TimelineRule>>::new();
    for rule in rules.iter().filter(|rule| rule.target_write_enabled) {
        let (Some(target), Some(target_sheet)) = (&rule.target_wb_path, &rule.target_sheet) else {
            continue;
        };
        if target_sheet.trim().is_empty() {
            continue;
        }
        grouped.entry(target.clone()).or_default().push(rule);
    }

    let mut reports = Vec::new();
    for (target, target_rules) in grouped {
        let target_existed = target.is_file();
        let target_sha256 = if target_existed {
            Some(crate::file_commit::sha256_file(&target)?)
        } else {
            None
        };
        let mut report = AppendTargetPlan {
            target: target.clone(),
            candidate: wide_append_candidate_path(&target, output_dir),
            target_existed,
            target_sha256: target_sha256.clone(),
            task_ok: 0,
            task_skipped: 0,
            appended_rows: 0,
            duplicate_rows_filtered: 0,
            merged_tasks: 0,
            issues: Vec::new(),
            edit_plan: None,
            new_workbook: None,
            replacement_workbook: None,
        };
        if target.exists() && !target.is_file() {
            report.task_skipped = target_rules.len();
            report.issues.push(format!(
                "target path exists but is not a regular file: {}",
                target.display()
            ));
            reports.push(report);
            continue;
        }
        if !target
            .extension()
            .and_then(|value| value.to_str())
            .is_some_and(|value| value.eq_ignore_ascii_case("xlsx"))
        {
            report.task_skipped = target_rules.len();
            report.issues.push(format!(
                "wide target append supports only xlsx targets through ExcelModifier: {}",
                target.display()
            ));
            reports.push(report);
            continue;
        }
        let target_dump = if report.target_existed {
            match read_workbook(&target) {
                Ok(dump) => Some(dump),
                Err(error) => {
                    report.task_skipped = target_rules.len();
                    report.issues.push(format!("read target failed: {error:#}"));
                    reports.push(report);
                    continue;
                }
            }
        } else {
            None
        };
        let mut sheet_states = HashMap::<String, WideTargetSheetState>::new();
        let mut sheet_order = Vec::<String>::new();

        for rule in target_rules {
            let target_sheet_name = rule.target_sheet.as_deref().unwrap_or_default();
            let summary_sheet_name =
                crate::sheet_extract::sanitize_sheet_name(if rule.name.trim().is_empty() {
                    "未命名规则"
                } else {
                    &rule.name
                });
            let Some(summary_sheet) = summary
                .sheets
                .iter()
                .find(|sheet| sheet.name == summary_sheet_name)
            else {
                report.task_skipped += 1;
                report.issues.push(format!(
                    "rule {:?}: no summary worksheet was produced",
                    rule.name
                ));
                continue;
            };
            let (headers, data_rows) = wide_output_table(summary_sheet);
            if headers.is_empty() || data_rows.is_empty() {
                report.task_skipped += 1;
                report.issues.push(format!(
                    "rule {:?}: summary has no rows to append",
                    rule.name
                ));
                continue;
            }

            if !sheet_states.contains_key(target_sheet_name) {
                let existing_sheet = target_dump.as_ref().and_then(|dump| {
                    dump.sheets
                        .iter()
                        .find(|sheet| sheet.name == target_sheet_name)
                });
                let existing_headers = existing_sheet
                    .map(wide_existing_headers)
                    .unwrap_or_default();
                let rows = match existing_sheet {
                    Some(sheet) => match wide_existing_append_rows(sheet, &existing_headers) {
                        Ok(rows) => rows,
                        Err(error) => {
                            report.task_skipped += 1;
                            report.issues.push(format!(
                                "read existing target worksheet {target_sheet_name:?} failed: {error:#}"
                            ));
                            continue;
                        }
                    },
                    None => Vec::new(),
                };
                sheet_order.push(target_sheet_name.to_owned());
                sheet_states.insert(
                    target_sheet_name.to_owned(),
                    WideTargetSheetState {
                        initialized: existing_sheet.is_some(),
                        existed: existing_sheet.is_some(),
                        headers: existing_headers,
                        seen_keys: HashMap::new(),
                        rows,
                        appended_rows: 0,
                    },
                );
            }
            let state = sheet_states
                .get_mut(target_sheet_name)
                .expect("wide append sheet state was inserted");

            let set_aliases = disambiguate_headers(
                &rule
                    .set_items
                    .iter()
                    .map(|(alias, _)| alias.clone())
                    .collect::<Vec<_>>(),
            );
            let mut required_prefix = vec![
                "工作簿名".to_owned(),
                "工作表名".to_owned(),
                "数据日期".to_owned(),
            ];
            required_prefix.extend(set_aliases);
            if rule.row_header_col_specified {
                required_prefix.push("行头路径".to_owned());
            }
            if state.existed
                && required_prefix
                    .iter()
                    .any(|header| !state.headers.contains(header))
            {
                report.task_skipped += 1;
                report.issues.push(format!(
                    "rule {:?}: header_mismatch on target worksheet {:?}",
                    rule.name, target_sheet_name
                ));
                continue;
            }

            let key_names = match wide_target_key_names(rule, &headers) {
                Ok(names) if !names.is_empty() => names,
                Ok(_) => {
                    report.task_skipped += 1;
                    report
                        .issues
                        .push(format!("rule {:?}: target dedup key is empty", rule.name));
                    continue;
                }
                Err(error) => {
                    report.task_skipped += 1;
                    report.issues.push(format!("rule {:?}: {error}", rule.name));
                    continue;
                }
            };
            if !state.seen_keys.contains_key(&key_names) {
                let mut keys = HashSet::new();
                if let Some(existing_sheet) = target_dump.as_ref().and_then(|dump| {
                    dump.sheets
                        .iter()
                        .find(|sheet| sheet.name == target_sheet_name)
                }) {
                    for row in 2..=wide_used_row_count(existing_sheet) {
                        let key = key_names
                            .iter()
                            .map(|name| {
                                state
                                    .headers
                                    .iter()
                                    .position(|header| header == name)
                                    .and_then(|column| {
                                        wide_dump_cell(existing_sheet, row, column + 1)
                                    })
                                    .as_ref()
                                    .map(wide_cell_key)
                                    .unwrap_or_default()
                            })
                            .collect::<Vec<_>>();
                        keys.insert(key);
                    }
                }
                state.seen_keys.insert(key_names.clone(), keys);
            }
            let mut seen = state.seen_keys.get(&key_names).cloned().unwrap_or_default();

            let mut batch_keys = HashSet::<Vec<String>>::new();
            for input_row in &data_rows {
                let key = key_names
                    .iter()
                    .map(|name| {
                        headers
                            .iter()
                            .position(|header| header == name)
                            .and_then(|index| input_row.get(index))
                            .and_then(Option::as_ref)
                            .map(wide_cell_key)
                            .unwrap_or_default()
                    })
                    .collect::<Vec<_>>();
                if !batch_keys.insert(key.clone()) {
                    report.duplicate_rows_filtered += 1;
                    continue;
                }
                if seen.contains(&key) {
                    report.duplicate_rows_filtered += 1;
                    continue;
                }
                seen.insert(key);
                let mut aligned = vec![AppendCell::Value(CellValue::Blank); state.headers.len()];
                for (source_column, header) in headers.iter().enumerate() {
                    let target_column = if let Some(index) =
                        state.headers.iter().position(|value| value == header)
                    {
                        index
                    } else {
                        state.headers.push(header.clone());
                        let index = state.headers.len() - 1;
                        if state.initialized {
                            if let Some(header_row) = state.rows.first_mut() {
                                header_row.push(AppendCell::Value(CellValue::Text(header.clone())));
                            }
                            for previous_row in state.rows.iter_mut().skip(1) {
                                previous_row.push(AppendCell::Value(CellValue::Blank));
                            }
                        }
                        index
                    };
                    if aligned.len() < state.headers.len() {
                        aligned.resize(state.headers.len(), AppendCell::Value(CellValue::Blank));
                    }
                    if let Some(value) = input_row.get(source_column) {
                        aligned[target_column] = wide_append_cell(value.as_ref(), header)?;
                    }
                }
                if !state.initialized {
                    state.initialized = true;
                    state.rows.push(
                        state
                            .headers
                            .iter()
                            .cloned()
                            .map(|header| AppendCell::Value(CellValue::Text(header)))
                            .collect(),
                    );
                }
                state.rows.push(aligned);
                state.appended_rows += 1;
                report.appended_rows += 1;
            }
            state.seen_keys.insert(key_names, seen);
            if !data_rows.is_empty() || !headers.is_empty() {
                report.task_ok += 1;
            }
        }

        let replacement_sheets = sheet_order
            .iter()
            .filter_map(|name| {
                let state = sheet_states.get(name)?;
                (state.appended_rows > 0).then(|| NewAppendSheet {
                    name: name.clone(),
                    rows: state.rows.clone(),
                    style_header_row: true,
                    data_date_column: state.headers.iter().position(|header| header == "数据日期"),
                })
            })
            .collect::<Vec<_>>();

        if report.target_existed && !replacement_sheets.is_empty() {
            let profile = match inspect_xlsx_profile(&target) {
                Ok(profile) if profile.inventory_complete => profile,
                Ok(_) => {
                    report.issues.push(
                        "target OOXML inventory is incomplete; refusing to create an edit plan"
                            .to_owned(),
                    );
                    report.task_skipped += report.task_ok;
                    report.task_ok = 0;
                    reports.push(report);
                    continue;
                }
                Err(error) => {
                    report
                        .issues
                        .push(format!("inspect target failed: {error:#}"));
                    report.task_skipped += report.task_ok;
                    report.task_ok = 0;
                    reports.push(report);
                    continue;
                }
            };
            let source_workbook = wide_replacement_source_path(&report.candidate);
            let mut operations = Vec::<EditOperation>::new();
            let original_sheets = target_dump
                .as_ref()
                .expect("existing target dump was read above");
            let mut used_names = original_sheets
                .sheets
                .iter()
                .map(|sheet| sheet.name.to_lowercase())
                .collect::<HashSet<_>>();
            let mut append_index = original_sheets.sheets.len();
            for replacement in &replacement_sheets {
                let range = wide_append_sheet_range(replacement)?;
                if let Some(index) = original_sheets
                    .sheets
                    .iter()
                    .position(|sheet| sheet.name == replacement.name)
                {
                    let temporary_name = wide_temporary_sheet_name(&mut used_names)?;
                    operations.push(EditOperation::CreateSheetAt {
                        name: temporary_name.clone(),
                        index,
                    });
                    operations.push(wide_import_operation(
                        &source_workbook,
                        &replacement.name,
                        &range,
                        &temporary_name,
                    ));
                    operations.push(EditOperation::DeleteSheet {
                        name: replacement.name.clone(),
                    });
                    operations.push(EditOperation::RenameSheet {
                        old_name: temporary_name,
                        new_name: replacement.name.clone(),
                    });
                } else {
                    operations.push(EditOperation::CreateSheetAt {
                        name: replacement.name.clone(),
                        index: append_index,
                    });
                    operations.push(wide_import_operation(
                        &source_workbook,
                        &replacement.name,
                        &range,
                        &replacement.name,
                    ));
                    append_index += 1;
                }
            }
            let preserve = BTreeSet::from([
                PreservationRequirement::CellValuesAndTypes,
                PreservationRequirement::FormulaText,
                PreservationRequirement::Styles,
                PreservationRequirement::CommentsAndPositions,
                PreservationRequirement::MergedRanges,
                PreservationRequirement::ConditionalFormatting,
                PreservationRequirement::DataValidations,
                PreservationRequirement::DefinedNames,
                PreservationRequirement::RowAndColumnDimensions,
                PreservationRequirement::PageSettings,
                PreservationRequirement::HeaderFooterText,
                PreservationRequirement::Relationships,
                PreservationRequirement::SheetIdentityAndOrder,
            ]);
            let mut drop_allowed = BTreeMap::new();
            for replacement in &replacement_sheets {
                if let Some(old_sheet) = original_sheets
                    .sheets
                    .iter()
                    .find(|sheet| sheet.name == replacement.name)
                {
                    let mut allowed = BTreeSet::from([
                        PreservationRequirement::FormulaText,
                        PreservationRequirement::CommentsAndPositions,
                        PreservationRequirement::Styles,
                        PreservationRequirement::MergedRanges,
                        PreservationRequirement::RowAndColumnDimensions,
                        PreservationRequirement::PageSettings,
                        PreservationRequirement::HeaderFooterText,
                        PreservationRequirement::Relationships,
                    ]);
                    if !old_sheet.conditional_formatting.is_empty() {
                        allowed.insert(PreservationRequirement::ConditionalFormatting);
                    }
                    if !old_sheet.data_validations.is_empty() {
                        allowed.insert(PreservationRequirement::DataValidations);
                    }
                    drop_allowed.insert(replacement.name.clone(), allowed);
                }
            }
            let plan = ExcelEditPlan {
                request_id: format!(
                    "pytools-wide-append-{}",
                    target.file_stem().unwrap_or_default().to_string_lossy()
                ),
                input: target.clone(),
                output: report.candidate.clone(),
                format: WorkbookFormat::Xlsx,
                workbook: profile,
                operations,
                preserve,
                drop_allowed,
            };
            plan.validate()
                .context("validate wide-summary target append plan")?;
            report.edit_plan = Some(plan);
            report.replacement_workbook = Some(NewAppendWorkbook {
                sheets: replacement_sheets,
            });
        } else if !report.target_existed {
            if !replacement_sheets.is_empty() {
                report.new_workbook = Some(NewAppendWorkbook {
                    sheets: replacement_sheets,
                });
            }
        }
        crate::file_commit::verify_planning_snapshot(&target, target_sha256.as_deref())
            .with_context(|| format!("verify wide-summary target snapshot {}", target.display()))?;
        reports.push(report);
    }
    Ok(reports)
}

/// Create the isolated source workbook required by PT-1-3/PT-1-4 replacement
/// plans. The source is always a fresh sibling of the candidate and is removed
/// by the CLI after execution.
pub fn prepare_wide_target_replacement_source(plan: &AppendTargetPlan) -> Result<PathBuf> {
    let payload = plan
        .replacement_workbook
        .as_ref()
        .context("wide target plan has no replacement workbook payload")?;
    let edit_plan = plan
        .edit_plan
        .as_ref()
        .context("wide target plan has no ExcelEditPlan")?;
    anyhow::ensure!(
        plan.target_existed,
        "replacement requires an existing target workbook"
    );
    anyhow::ensure!(
        edit_plan.input == plan.target && edit_plan.output == plan.candidate,
        "wide target edit plan paths do not match the append plan"
    );
    let mut imports = edit_plan
        .operations
        .iter()
        .filter_map(|operation| match operation {
            EditOperation::ImportRange {
                source_workbook, ..
            } => Some(source_workbook),
            _ => None,
        });
    let source = imports
        .next()
        .context("wide target edit plan has no ImportRange operation")?;
    anyhow::ensure!(
        imports.all(|path| path == source),
        "wide target replacement imports must use one staging workbook"
    );
    anyhow::ensure!(
        source == &wide_replacement_source_path(&plan.candidate),
        "wide target replacement source path is not the planned isolated staging path"
    );
    anyhow::ensure!(
        source != &plan.target && source != &plan.candidate,
        "wide target replacement source must not alias input or output"
    );
    let source = write_append_workbook_payload(payload, source)?;
    Ok(source)
}

/// Verify the edited candidate against the staged replacement sheets and the
/// original workbook's untouched-sheet data/metadata before reporting success.
pub fn validate_wide_target_replacement_candidate(
    plan: &AppendTargetPlan,
    staging_workbook: &Path,
) -> Result<()> {
    let replacements = plan
        .replacement_workbook
        .as_ref()
        .context("wide target plan has no replacement workbook payload")?;
    let original = read_workbook(&plan.target)
        .with_context(|| format!("reopen wide target input {}", plan.target.display()))?;
    let staged = read_workbook(staging_workbook).with_context(|| {
        format!(
            "reopen wide replacement source {}",
            staging_workbook.display()
        )
    })?;
    let output = read_workbook(&plan.candidate)
        .with_context(|| format!("reopen wide target candidate {}", plan.candidate.display()))?;
    let replacement_names = replacements
        .sheets
        .iter()
        .map(|sheet| sheet.name.as_str())
        .collect::<HashSet<_>>();
    let mut expected_names = original
        .sheets
        .iter()
        .map(|sheet| sheet.name.clone())
        .collect::<Vec<_>>();
    for replacement in &replacements.sheets {
        if !expected_names.iter().any(|name| name == &replacement.name) {
            expected_names.push(replacement.name.clone());
        }
        let expected_sheet = staged
            .sheets
            .iter()
            .find(|sheet| sheet.name == replacement.name)
            .with_context(|| format!("staged replacement sheet {:?} missing", replacement.name))?;
        let actual_sheet = output
            .sheets
            .iter()
            .find(|sheet| sheet.name == replacement.name)
            .with_context(|| {
                format!("candidate replacement sheet {:?} missing", replacement.name)
            })?;
        ensure_sheet_dump_matches(expected_sheet, actual_sheet, &replacement.name)?;
    }
    let actual_names = output
        .sheets
        .iter()
        .map(|sheet| sheet.name.clone())
        .collect::<Vec<_>>();
    anyhow::ensure!(
        actual_names == expected_names,
        "wide target candidate sheet order differs: expected {expected_names:?}, found {actual_names:?}"
    );
    for source_sheet in &original.sheets {
        if replacement_names.contains(source_sheet.name.as_str()) {
            continue;
        }
        let output_sheet = output
            .sheets
            .iter()
            .find(|sheet| sheet.name == source_sheet.name)
            .with_context(|| format!("untouched worksheet {:?} missing", source_sheet.name))?;
        ensure_sheet_dump_matches(source_sheet, output_sheet, &source_sheet.name)?;
    }
    Ok(())
}

fn ensure_sheet_dump_matches(
    expected: &excel_reader::SheetDump,
    actual: &excel_reader::SheetDump,
    name: &str,
) -> Result<()> {
    anyhow::ensure!(
        expected.values == actual.values,
        "worksheet {name:?} cell values/types differ after replacement"
    );
    anyhow::ensure!(
        expected.formulas == actual.formulas,
        "worksheet {name:?} formula text differs after replacement"
    );
    anyhow::ensure!(
        expected.comments == actual.comments,
        "worksheet {name:?} comments differ after replacement"
    );
    anyhow::ensure!(
        expected.merges == actual.merges,
        "worksheet {name:?} merged ranges differ after replacement"
    );
    anyhow::ensure!(
        expected.conditional_formatting == actual.conditional_formatting,
        "worksheet {name:?} conditional formatting differs after replacement"
    );
    anyhow::ensure!(
        expected.data_validations == actual.data_validations,
        "worksheet {name:?} data validations differ after replacement"
    );
    Ok(())
}

fn wide_output_table(
    sheet: &excel_reader::SheetDump,
) -> (Vec<String>, Vec<Vec<Option<WorkbookCell>>>) {
    let rows = wide_used_row_count(sheet);
    let columns = wide_used_column_count(sheet);
    if rows == 0 || columns == 0 {
        return (Vec::new(), Vec::new());
    }
    let headers = (1..=columns)
        .map(|column| {
            wide_dump_cell(sheet, 1, column)
                .map(|value| value.to_sheet_text())
                .unwrap_or_default()
        })
        .collect::<Vec<_>>();
    let data = (2..=rows)
        .map(|row| {
            (1..=columns)
                .map(|column| wide_dump_cell(sheet, row, column))
                .collect::<Vec<_>>()
        })
        .collect();
    (headers, data)
}

fn wide_existing_headers(sheet: &excel_reader::SheetDump) -> Vec<String> {
    (1..=wide_used_column_count(sheet))
        .map(|column| {
            wide_dump_cell(sheet, 1, column)
                .map(|value| value.to_sheet_text())
                .filter(|value| !value.is_empty())
                .unwrap_or_else(|| format!("Unnamed: {}", column - 1))
        })
        .collect()
}

fn wide_used_row_count(sheet: &excel_reader::SheetDump) -> usize {
    sheet
        .values
        .keys()
        .chain(sheet.formulas.keys())
        .filter_map(|address| parse_a1(address).map(|(row, _)| row))
        .max()
        .unwrap_or(0)
}

fn wide_used_column_count(sheet: &excel_reader::SheetDump) -> usize {
    sheet
        .values
        .keys()
        .chain(sheet.formulas.keys())
        .filter_map(|address| parse_a1(address).map(|(_, column)| column))
        .max()
        .unwrap_or(0)
}

fn wide_dump_cell(
    sheet: &excel_reader::SheetDump,
    row: usize,
    column: usize,
) -> Option<WorkbookCell> {
    let address = crate::rule_engine::cell_address(row, column);
    sheet.values.get(&address).and_then(WorkbookCell::from_json)
}

fn wide_dump_cell_strict(
    sheet: &excel_reader::SheetDump,
    row: usize,
    column: usize,
) -> Result<Option<WorkbookCell>> {
    let address = crate::rule_engine::cell_address(row, column);
    sheet
        .values
        .get(&address)
        .map(|value| {
            WorkbookCell::from_json(value)
                .with_context(|| format!("decode existing target cell {address}"))
        })
        .transpose()
}

fn wide_existing_append_rows(
    sheet: &excel_reader::SheetDump,
    headers: &[String],
) -> Result<Vec<Vec<AppendCell>>> {
    let mut rows = vec![headers
        .iter()
        .cloned()
        .map(|header| AppendCell::Value(CellValue::Text(header)))
        .collect::<Vec<_>>()];
    let max_row = wide_used_row_count(sheet);
    for row in 2..=max_row {
        let mut values = Vec::with_capacity(headers.len());
        for (column, header) in headers.iter().enumerate() {
            let value = wide_dump_cell_strict(sheet, row, column + 1)?;
            values.push(wide_append_cell(value.as_ref(), header)?);
        }
        rows.push(values);
    }
    Ok(rows)
}

fn wide_append_sheet_range(sheet: &NewAppendSheet) -> Result<String> {
    anyhow::ensure!(
        !sheet.rows.is_empty(),
        "replacement worksheet has no header row"
    );
    let columns = sheet.rows.iter().map(Vec::len).max().unwrap_or(0);
    anyhow::ensure!(columns > 0, "replacement worksheet has no columns");
    anyhow::ensure!(
        sheet.rows.iter().all(|row| row.len() == columns),
        "replacement worksheet rows must have a consistent column count"
    );
    let end = crate::rule_engine::cell_address(sheet.rows.len(), columns);
    Ok(format!("A1:{end}"))
}

fn wide_temporary_sheet_name(used_names: &mut HashSet<String>) -> Result<String> {
    for suffix in 1..=10_000usize {
        let candidate = format!("__pt_replace_{suffix}");
        if used_names.insert(candidate.to_lowercase()) {
            return Ok(candidate);
        }
    }
    bail!("could not allocate a temporary worksheet name for replacement")
}

fn wide_import_operation(
    source_workbook: &Path,
    source_sheet: &str,
    source_range: &str,
    target_sheet: &str,
) -> EditOperation {
    EditOperation::ImportRange {
        source_workbook: source_workbook.to_path_buf(),
        source_sheet: source_sheet.to_owned(),
        source_range: source_range.to_owned(),
        target_sheet: target_sheet.to_owned(),
        target_start: "A1".to_owned(),
        copy_values: true,
        copy_formulas: true,
        style_mode: RangeStyleMode::Full,
        copy_comments: false,
        copy_merges: false,
        copy_row_heights: false,
        copy_column_widths: false,
        copy_row_hidden: false,
        copy_column_hidden: false,
        copy_outline_levels: false,
        copy_all_row_dimensions: false,
        copy_all_column_dimensions: false,
        materialize_blank_cells: true,
        hide_zero: false,
    }
}

fn wide_target_key_names(rule: &TimelineRule, headers: &[String]) -> Result<Vec<String>> {
    let mut names = Vec::new();
    if !rule.target_dedup_cols.is_empty() {
        for spec in &rule.target_dedup_cols {
            let index = headers
                .iter()
                .position(|header| header == spec)
                .or_else(|| {
                    let column = crate::path_map::parse_col_spec(spec);
                    (column > 0 && column <= headers.len()).then_some(column - 1)
                })
                .with_context(|| format!("target dedup column {spec:?} does not exist"))?;
            let name = headers[index].clone();
            if !names.contains(&name) {
                names.push(name);
            }
        }
    } else {
        names.push("数据日期".to_owned());
        let set_headers = disambiguate_headers(
            &rule
                .set_items
                .iter()
                .map(|(alias, _)| alias.clone())
                .collect::<Vec<_>>(),
        );
        for alias in &set_headers {
            if headers.contains(alias) && !names.contains(alias) {
                names.push(alias.clone());
            }
        }
        if headers.iter().any(|header| header == "行头路径") {
            names.push("行头路径".to_owned());
        } else {
            let fixed = ["工作簿名", "工作表名", "数据日期"]
                .into_iter()
                .map(str::to_owned)
                .chain(set_headers)
                .collect::<HashSet<_>>();
            if let Some(column) = headers.iter().find(|header| !fixed.contains(*header)) {
                names.push(column.clone());
            }
        }
        names.dedup();
    }
    Ok(names)
}

fn wide_cell_key(value: &WorkbookCell) -> String {
    match value {
        WorkbookCell::Text(value) | WorkbookCell::Error(value) => value.trim().to_owned(),
        WorkbookCell::Integer(value) => value.to_string(),
        WorkbookCell::Number(value) => {
            if value.fract() == 0.0 {
                format!("{value:.0}")
            } else {
                format!("{value:.15}")
                    .trim_end_matches('0')
                    .trim_end_matches('.')
                    .to_owned()
            }
        }
        WorkbookCell::Boolean(value) => (if *value { "True" } else { "False" }).to_owned(),
        WorkbookCell::DateTimeSerial(value) => wide_serial_key(*value),
        WorkbookCell::DateTimeIso(value) => wide_iso_datetime_key(value),
        WorkbookCell::DurationIso(value) => value.clone(),
    }
}

fn wide_append_cell(value: Option<&WorkbookCell>, header: &str) -> Result<AppendCell> {
    use WorkbookCell as Cell;
    let Some(value) = value else {
        return Ok(AppendCell::Value(CellValue::Blank));
    };
    let number_format = |serial: f64| {
        if header == "数据日期" {
            "yyyy/m/d".to_owned()
        } else {
            wide_date_format_code(serial).to_owned()
        }
    };
    Ok(match value {
        Cell::Text(value) if value.is_empty() => AppendCell::Value(CellValue::Blank),
        Cell::Text(value) | Cell::Error(value) => AppendCell::Value(CellValue::Text(value.clone())),
        Cell::Integer(value) => AppendCell::Value(CellValue::Decimal(value.to_string())),
        Cell::Number(value) => AppendCell::Value(CellValue::Decimal(value.to_string())),
        Cell::Boolean(value) => AppendCell::Value(CellValue::Boolean(*value)),
        Cell::DateTimeSerial(serial) => AppendCell::DateTime {
            serial: serial.to_string(),
            number_format: number_format(*serial),
        },
        Cell::DateTimeIso(value) => {
            let serial = wide_datetime_serial(value)
                .with_context(|| format!("unsupported date value {value:?}"))?;
            AppendCell::DateTime {
                serial: serial.to_string(),
                number_format: number_format(serial),
            }
        }
        Cell::DurationIso(value) => {
            let serial = wide_duration_serial(value)
                .with_context(|| format!("unsupported duration value {value:?}"))?;
            AppendCell::DateTime {
                serial: serial.to_string(),
                number_format: "[h]:mm:ss".to_owned(),
            }
        }
    })
}

fn wide_serial_key(serial: f64) -> String {
    use chrono::{Duration, NaiveDate, Timelike};
    let Some(epoch) =
        NaiveDate::from_ymd_opt(1899, 12, 30).and_then(|date| date.and_hms_opt(0, 0, 0))
    else {
        return serial.to_string();
    };
    let millis = (serial * 86_400_000.0).round() as i64;
    let value = epoch + Duration::milliseconds(millis);
    if value.time().num_seconds_from_midnight() == 0 {
        value.format("%Y-%m-%d").to_string()
    } else {
        value.format("%Y-%m-%d %H:%M:%S").to_string()
    }
}

fn wide_iso_datetime_key(text: &str) -> String {
    use chrono::{DateTime, NaiveDate, NaiveDateTime, Timelike};
    let datetime = DateTime::parse_from_rfc3339(text)
        .map(|value| value.naive_utc())
        .or_else(|_| NaiveDateTime::parse_from_str(text, "%Y-%m-%dT%H:%M:%S%.f"))
        .or_else(|_| NaiveDateTime::parse_from_str(text, "%Y-%m-%d %H:%M:%S%.f"));
    if let Ok(datetime) = datetime {
        if datetime.time().num_seconds_from_midnight() == 0 {
            datetime.format("%Y-%m-%d").to_string()
        } else {
            datetime.format("%Y-%m-%d %H:%M:%S").to_string()
        }
    } else if let Ok(date) = NaiveDate::parse_from_str(text, "%Y-%m-%d") {
        date.format("%Y-%m-%d").to_string()
    } else {
        text.trim().to_owned()
    }
}

fn wide_append_candidate_path(target: &Path, output_dir: &Path) -> PathBuf {
    let stamp = Local::now().format("%Y%m%d_%H%M%S");
    let stem = target.file_stem().unwrap_or_default().to_string_lossy();
    for suffix in 0usize..10000 {
        let tail = if suffix == 0 {
            String::new()
        } else {
            format!("_{suffix}")
        };
        let path = output_dir.join(format!("{stem}_wide_append_{stamp}{tail}.xlsx"));
        if !path.exists() {
            return path;
        }
    }
    output_dir.join(format!(
        "{stem}_wide_append_{}.xlsx",
        chrono::Utc::now().timestamp_millis()
    ))
}

fn wide_replacement_source_path(candidate: &Path) -> PathBuf {
    let stem = candidate.file_stem().unwrap_or_default().to_string_lossy();
    candidate.with_file_name(format!(".{stem}.replacement-source.xlsx"))
}

fn position_differences(
    template: &[Vec<Option<WorkbookCell>>],
    source: &[Vec<Option<WorkbookCell>>],
    rule: &TimelineRule,
) -> Vec<(String, String, String)> {
    let (tr, tc) = shape(template);
    let (sr, sc) = shape(source);
    let row_header_col = rule.row_header_col.saturating_sub(1);
    let header_rows = rule
        .col_header_rows
        .iter()
        .filter_map(|row| row.checked_sub(1))
        .collect::<Vec<_>>();
    let row_start = if rule.data_row_start > 0 {
        rule.data_row_start - 1
    } else {
        header_rows.iter().max().map_or(0, |row| row + 1)
    };
    let row_end = rule
        .data_row_end
        .map(|row| row.saturating_sub(1))
        .unwrap_or_else(|| tr.max(sr).saturating_sub(1))
        .min(tr.max(sr).saturating_sub(1));
    let col_start = if rule.data_col_start > 0 {
        rule.data_col_start - 1
    } else {
        row_header_col + 1
    };
    let col_end = rule
        .data_col_end
        .map(|column| column.saturating_sub(1))
        .unwrap_or_else(|| tc.max(sc).saturating_sub(1))
        .min(tc.max(sc).saturating_sub(1));
    if tr.max(sr) == 0 || tc.max(sc) == 0 {
        return Vec::new();
    }
    let mut diffs = Vec::new();
    if col_start <= col_end {
        for col in col_start..=col_end {
            let expected = header_rows
                .iter()
                .map(|row| text_at(template, *row, col))
                .filter(|text| !text.is_empty())
                .collect::<Vec<_>>()
                .join("_");
            let actual = header_rows
                .iter()
                .map(|row| text_at(source, *row, col))
                .filter(|text| !text.is_empty())
                .collect::<Vec<_>>()
                .join("_");
            if expected != actual {
                diffs.push(("列头".into(), actual, expected));
            }
        }
    }
    if row_start <= row_end {
        for row in row_start..=row_end {
            let expected = text_at(template, row, row_header_col);
            let actual = text_at(source, row, row_header_col);
            if expected != actual {
                diffs.push(("行头".into(), actual, expected));
            }
        }
    }
    diffs
}

fn path_differences(
    template: &[Vec<Option<WorkbookCell>>],
    source: &[Vec<Option<WorkbookCell>>],
    rule: &TimelineRule,
) -> Vec<(String, String, String)> {
    let (tr, tc) = shape(template);
    let (sr, sc) = shape(source);
    if tr.max(sr) == 0 || tc.max(sc) == 0 {
        return Vec::new();
    }
    let header_rows = rule
        .col_header_rows
        .iter()
        .filter_map(|row| row.checked_sub(1))
        .collect::<Vec<_>>();
    let row_start = if rule.data_row_start > 0 {
        rule.data_row_start - 1
    } else {
        header_rows.iter().max().map_or(0, |r| r + 1)
    };
    let row_end = rule
        .data_row_end
        .map(|x| x.saturating_sub(1))
        .unwrap_or_else(|| tr.max(sr) - 1)
        .min(tr.max(sr) - 1);
    let col_start = if rule.data_col_start > 0 {
        rule.data_col_start - 1
    } else {
        rule.row_header_col
    };
    let col_end = rule
        .data_col_end
        .map(|x| x.saturating_sub(1))
        .unwrap_or_else(|| tc.max(sc) - 1)
        .min(tc.max(sc) - 1);
    let rows = |matrix: &[Vec<Option<WorkbookCell>>]| -> BTreeSet<String> {
        if row_start > row_end {
            return BTreeSet::new();
        }
        (row_start..=row_end)
            .map(|r| text_at(matrix, r, rule.row_header_col.saturating_sub(1)))
            .filter(|v| !v.is_empty())
            .collect()
    };
    let cols = |matrix: &[Vec<Option<WorkbookCell>>]| -> BTreeSet<String> {
        if col_start > col_end {
            return BTreeSet::new();
        }
        (col_start..=col_end)
            .map(|c| {
                header_rows
                    .iter()
                    .map(|r| text_at(matrix, *r, c))
                    .filter(|v| !v.is_empty())
                    .collect::<Vec<_>>()
                    .join("_")
            })
            .filter(|v| !v.is_empty())
            .collect()
    };
    let trp = rows(template);
    let srp = rows(source);
    let tcp = cols(template);
    let scp = cols(source);
    let mut out = Vec::new();
    for value in srp.difference(&trp) {
        out.push(("行头".into(), value.clone(), String::new()));
    }
    for value in trp.difference(&srp) {
        out.push(("行头".into(), String::new(), value.clone()));
    }
    for value in scp.difference(&tcp) {
        out.push(("列头".into(), value.clone(), String::new()));
    }
    for value in tcp.difference(&scp) {
        out.push(("列头".into(), String::new(), value.clone()));
    }
    out
}

fn shape(matrix: &[Vec<Option<WorkbookCell>>]) -> (usize, usize) {
    (matrix.len(), matrix.iter().map(Vec::len).max().unwrap_or(0))
}
fn text_at(matrix: &[Vec<Option<WorkbookCell>>], row: usize, col: usize) -> String {
    matrix
        .get(row)
        .and_then(|r| r.get(col))
        .and_then(Option::as_ref)
        .map(SheetValue::to_sheet_text)
        .unwrap_or_default()
        .trim()
        .to_owned()
}
fn keyword_match(text: &str, keywords: &str) -> bool {
    crate::path_map::match_all_keywords(text, keywords)
}
fn file_name(path: &Path) -> String {
    path.file_name()
        .map(|name| name.to_string_lossy().into_owned())
        .unwrap_or_default()
}
fn same_path(left: &Path, right: &Path) -> bool {
    left.canonicalize()
        .ok()
        .zip(right.canonicalize().ok())
        .is_some_and(|(l, r)| l == r)
}
fn precheck_headers(sheet: &SheetDump) -> BTreeMap<String, usize> {
    sheet
        .values
        .iter()
        .filter_map(|(address, value)| {
            let (row, column) = parse_a1(address)?;
            if row != 1 {
                return None;
            }
            let header = value_text(value).trim().to_owned();
            (!header.is_empty()).then_some((header, column))
        })
        .collect()
}
fn precheck_field(
    sheet: &SheetDump,
    headers: &BTreeMap<String, usize>,
    row: usize,
    name: &str,
) -> String {
    headers
        .get(name)
        .and_then(|column| {
            sheet
                .values
                .get(&format!("{}{}", column_letters(*column), row))
        })
        .map(value_text)
        .unwrap_or_default()
        .trim()
        .to_owned()
}
fn precheck_max_row(sheet: &SheetDump) -> usize {
    sheet
        .values
        .keys()
        .filter_map(|address| parse_a1(address).map(|(row, _)| row))
        .max()
        .unwrap_or(1)
}
fn column_letters(mut column: usize) -> String {
    let mut letters = String::new();
    while column > 0 {
        let remainder = (column - 1) % 26;
        letters.insert(0, (b'A' + remainder as u8) as char);
        column = (column - 1) / 26;
    }
    letters
}
fn precheck_truthy(value: &str) -> bool {
    matches!(
        value.trim().to_lowercase().as_str(),
        "是" | "1" | "true" | "y" | "yes" | "on"
    )
}
fn precheck_int_from_float(value: &str) -> Option<i64> {
    let value = value.trim().parse::<f64>().ok()?;
    value.is_finite().then_some(value as i64)
}
fn precheck_list_header_tokens(value: &str) -> Vec<String> {
    value
        .replace('，', ",")
        .replace('；', ";")
        .replace(' ', ",")
        .replace(';', ",")
        .split(',')
        .map(str::trim)
        .filter(|token| !token.is_empty())
        .map(str::to_owned)
        .collect()
}
fn precheck_set_error(value: &str) -> Option<String> {
    let normalized = value.replace('；', ";");
    for part in normalized
        .split(';')
        .map(str::trim)
        .filter(|part| !part.is_empty())
    {
        let Some((alias, address)) = part.split_once('@') else {
            return Some(format!("set区域项格式非法: {part}（应为 别名@地址）"));
        };
        if alias.trim().is_empty() {
            return Some(format!("set区域项别名为空: {part}"));
        }
        let address = address.trim();
        let split = address.find(|ch: char| ch.is_ascii_digit());
        let valid = split.is_some_and(|split| {
            let (column, row) = address.split_at(split);
            !column.is_empty()
                && column.chars().all(|ch| ch.is_ascii_alphabetic())
                && !row.is_empty()
                && row.as_bytes()[0].is_ascii_digit()
                && row.as_bytes()[0] != b'0'
                && row.bytes().all(|ch| ch.is_ascii_digit())
                && row.parse::<usize>().is_ok_and(|number| number > 0)
        });
        if !valid {
            return Some(format!("set区域项地址非法: {part}"));
        }
    }
    None
}
fn parse_a1(address: &str) -> Option<(usize, usize)> {
    let split = address.find(|ch: char| ch.is_ascii_digit())?;
    let col = address[..split].chars().try_fold(0usize, |acc, ch| {
        acc.checked_mul(26)?
            .checked_add(ch.to_ascii_uppercase() as usize - 'A' as usize + 1)
    })?;
    let row = address[split..].parse::<usize>().ok()?;
    Some((row, col))
}
fn value_text(value: &serde_json::Value) -> String {
    match value.get("type").and_then(serde_json::Value::as_str) {
        Some("string") | Some("datetime_iso") | Some("duration_iso") | Some("error") => value
            .get("value")
            .and_then(serde_json::Value::as_str)
            .unwrap_or_default()
            .to_owned(),
        Some("number") | Some("integer") => value
            .get("value")
            .map(ToString::to_string)
            .unwrap_or_default(),
        Some("boolean") => value
            .get("value")
            .and_then(serde_json::Value::as_bool)
            .map(|b| if b { "True" } else { "False" })
            .unwrap_or_default()
            .to_owned(),
        _ => String::new(),
    }
}
fn unique_output(dir: &Path, stem: &str, extension: &str) -> Result<PathBuf> {
    fs::create_dir_all(dir)?;
    for idx in 0..10000 {
        let suffix = if idx == 0 {
            String::new()
        } else {
            format!("_{idx}")
        };
        let path = dir.join(format!("{stem}{suffix}.{extension}"));
        if !path.exists() {
            return Ok(path);
        }
    }
    bail!("cannot allocate unique output file in {}", dir.display())
}

fn write_string_table(
    workbook: &mut Workbook,
    name: &str,
    headers: &[&str],
    rows: &[Vec<String>],
) -> Result<()> {
    let sheet = workbook.add_worksheet();
    sheet.set_name(name)?;
    let header_format = Format::new()
        .set_bold()
        .set_border(FormatBorder::Thin)
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::Top);
    for (col, header) in headers.iter().enumerate() {
        sheet.write_string_with_format(0, col as u16, *header, &header_format)?;
    }
    for (row, values) in rows.iter().enumerate() {
        for (col, value) in values.iter().enumerate() {
            sheet.write_string((row + 1) as u32, col as u16, value)?;
        }
    }
    Ok(())
}

fn write_compare_summary_table(workbook: &mut Workbook, rows: &[Vec<String>]) -> Result<()> {
    let sheet = workbook.add_worksheet();
    sheet.set_name("汇总")?;
    let header_format = Format::new()
        .set_bold()
        .set_border(FormatBorder::Thin)
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::Top);
    for (column, header) in ["源文件", "状态", "按位置差异数", "按路径差异数", "备注"]
        .iter()
        .enumerate()
    {
        sheet.write_string_with_format(0, column as u16, *header, &header_format)?;
    }
    for (row_index, values) in rows.iter().enumerate() {
        for (column, value) in values.iter().enumerate() {
            let excel_row = (row_index + 1) as u32;
            let excel_column = column as u16;
            if matches!(column, 2 | 3) {
                if let Ok(count) = value.parse::<f64>() {
                    sheet.write_number(excel_row, excel_column, count)?;
                    continue;
                }
            }
            sheet.write_string(excel_row, excel_column, value)?;
        }
    }
    Ok(())
}

#[cfg(test)]
mod keyword_match_tests {
    use super::keyword_match;

    #[test]
    fn sheet_keywords_match_python_case_and_delimiter_semantics() {
        assert!(keyword_match("REPORT Data", "report,data"));
        assert!(keyword_match("REPORT Data", "report\u{ff0c}data"));
        assert!(keyword_match("REPORT Data", "report\u{ff1b}data"));
        assert!(!keyword_match("REPORT Data", "report,missing"));
    }
}
