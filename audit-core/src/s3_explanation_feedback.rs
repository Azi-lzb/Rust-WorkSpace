//! Read-only S3-F03 planning for importing institution feedback.
//!
//! The Python workflow mutates an existing explanation workbook. This Rust
//! slice reads the target and feedback workbooks, applies the business identity
//! and conflict rules, and emits an ExcelEditPlan plus a conflict workbook.
//! It deliberately does not execute an unverified ExcelModifier capability.

use anyhow::{bail, ensure, Context, Result};
use chrono::Local;
use excel_editor::{
    inspect_xlsx_profile, CellValue as EditCellValue, EditOperation, ExcelEditPlan,
    PreservationRequirement, WorkbookFormat,
};
use excel_reader::{read_value_columns_with_iso_dates, CellValue, SheetValues};
use excel_writer::{Color, Format, Workbook};
use serde::Serialize;
use sha2::{Digest, Sha256};
use std::{
    collections::{BTreeMap, BTreeSet, HashMap},
    fs::{self, File, OpenOptions},
    io::{Read, Write},
    path::{Path, PathBuf},
    sync::atomic::{AtomicU64, Ordering},
    time::Instant,
};

const REQUIRED_ID_HEADERS: [&str; 4] = ["机构代码", "机构名称", "指标代码", "指标名称"];
const DATA_VALUE_HEADER: &str = "数据值";
const CHANGE_AMOUNT_PREFIX: &str = "增减额";
const REASON_HEADER: &str = "变动原因";
const READER: &str = "calamine";
const BUSINESS_ENGINE: &str = "audit-core (Rust)";
const CONFLICT_WRITER: &str = "rust_xlsxwriter";
const CONFLICT_HEADERS: [&str; 10] = [
    "状态",
    "机构代码",
    "机构名称",
    "指标代码",
    "指标名称",
    "数据值",
    "增减额",
    "目标说明文件行号",
    "反馈文件及原因",
    "冲突原因",
];
const CONFLICT_WIDTHS: [f64; 10] = [10.0, 14.0, 26.0, 14.0, 32.0, 14.0, 14.0, 20.0, 72.0, 38.0];

static TEMP_COUNTER: AtomicU64 = AtomicU64::new(0);
static PLAN_COUNTER: AtomicU64 = AtomicU64::new(0);

type Identity = (String, String, String, String);

#[derive(Clone, Debug, Serialize)]
pub struct S3ExplanationFeedbackPlanSummary {
    pub status: &'static str,
    pub applied: bool,
    pub target_workbook: PathBuf,
    pub plan_file: PathBuf,
    pub candidate_output_workbook: PathBuf,
    pub imported: usize,
    pub unmatched: usize,
    pub conflicts: usize,
    pub planned_cell_updates: usize,
    pub conflict_report_path: Option<PathBuf>,
    pub reader: &'static str,
    pub business_engine: &'static str,
    pub conflict_writer: Option<&'static str>,
    pub target_read_time_ms: u128,
    pub feedback_read_time_ms: u128,
    pub conflict_write_time_ms: u128,
    pub plan_write_time_ms: u128,
    pub total_time_ms: u128,
}

#[derive(Clone, Debug, Serialize)]
struct InputSnapshot {
    path: PathBuf,
    sha256: String,
}

#[derive(Clone, Debug, Serialize)]
struct PlanCounts {
    imported: usize,
    unmatched: usize,
    conflicts: usize,
    planned_cell_updates: usize,
}

#[derive(Clone, Debug, Serialize)]
struct PlanDocument {
    schema_version: u32,
    feature: &'static str,
    status: &'static str,
    created_at: String,
    applied: bool,
    target: InputSnapshot,
    feedback_sources: Vec<InputSnapshot>,
    candidate_output_workbook: PathBuf,
    counts: PlanCounts,
    conflict_report_path: Option<PathBuf>,
    edit_plan: Option<ExcelEditPlan>,
}

#[derive(Clone, Debug, Serialize)]
struct ConflictRow([String; 10]);

#[derive(Clone, Debug)]
struct TargetRow {
    row_number: usize,
    organization_name: CellValue,
    indicator_name: CellValue,
}

#[derive(Clone, Debug)]
struct HeaderColumns {
    row_number: usize,
    indexes: HashMap<String, usize>,
}

#[derive(Clone, Debug)]
struct ValueGrid {
    name: String,
    max_row: usize,
    rows: BTreeMap<usize, BTreeMap<usize, CellValue>>,
}

impl From<SheetValues> for ValueGrid {
    fn from(sheet: SheetValues) -> Self {
        let mut rows = BTreeMap::<usize, BTreeMap<usize, CellValue>>::new();
        for (address, value) in sheet.values {
            if let Some((row, column)) = parse_cell_address(&address) {
                rows.entry(row).or_default().insert(column, value);
            }
        }
        Self {
            name: sheet.name,
            max_row: sheet.max_row,
            rows,
        }
    }
}

impl ValueGrid {
    fn cell(&self, row: usize, column: usize) -> CellValue {
        self.rows
            .get(&row)
            .and_then(|cells| cells.get(&column))
            .cloned()
            .unwrap_or(CellValue::Empty)
    }

    fn has_header_text(&self, expected: &str) -> bool {
        (1..=self.max_row.min(10)).any(|row| {
            self.rows.get(&row).is_some_and(|cells| {
                cells.values().any(|value| {
                    !matches!(value, CellValue::Empty) && python_string(value).trim() == expected
                })
            })
        })
    }

    fn headers(&self) -> Result<HeaderColumns> {
        for row_number in 1..=self.max_row.min(10) {
            let mut indexes = HashMap::new();
            let mut ordered_names = Vec::new();
            if let Some(cells) = self.rows.get(&row_number) {
                for (column, value) in cells {
                    if matches!(value, CellValue::Empty) {
                        continue;
                    }
                    let name = python_string(value).trim().to_owned();
                    if name.is_empty() {
                        continue;
                    }
                    if !indexes.contains_key(&name) {
                        ordered_names.push(name.clone());
                    }
                    // Match Python's dict comprehension: a duplicate header
                    // retains its first insertion position but points to the
                    // last column carrying that text.
                    indexes.insert(name, *column);
                }
            }
            if !REQUIRED_ID_HEADERS
                .iter()
                .all(|name| indexes.contains_key(*name))
            {
                continue;
            }
            let value_header = ordered_names
                .iter()
                .find(|name| name.as_str() == DATA_VALUE_HEADER || name.starts_with("本期金额"));
            let change_header = ordered_names
                .iter()
                .find(|name| name.starts_with(CHANGE_AMOUNT_PREFIX));
            if let (Some(value_header), Some(change_header)) = (value_header, change_header) {
                if indexes.contains_key(REASON_HEADER) {
                    let value_column = indexes[value_header];
                    let change_column = indexes[change_header];
                    indexes.insert(DATA_VALUE_HEADER.to_owned(), value_column);
                    indexes.insert(CHANGE_AMOUNT_PREFIX.to_owned(), change_column);
                    return Ok(HeaderColumns {
                        row_number,
                        indexes,
                    });
                }
            }
        }
        bail!("说明文件缺少机构代码、机构名称、指标代码、指标名称、数据值、增减额或变动原因列")
    }
}

/// Compute S3-F03 matching and conflict outcomes without changing the target.
/// The JSON plan contains the proposed cell updates and source hashes so a
/// later, separately validated ExcelModifier step can reject stale inputs.
pub fn plan_explanation_feedback_import(
    target_path: impl AsRef<Path>,
    feedback_paths: &[PathBuf],
    candidate_output_path: impl AsRef<Path>,
    plan_output_path: impl AsRef<Path>,
) -> Result<S3ExplanationFeedbackPlanSummary> {
    let started = Instant::now();
    let target_path = target_path.as_ref();
    let candidate_output_path = candidate_output_path.as_ref();
    let plan_output_path = plan_output_path.as_ref();
    ensure!(!feedback_paths.is_empty(), "请至少选择一个机构反馈说明文件");
    ensure_workbook_path(target_path, true)?;
    ensure!(
        target_path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx")),
        "S3-F03 目标说明文件目前仅接受 .xlsx；为保护宏和其他工作簿对象，未尝试写入 .xlsm"
    );
    ensure_workbook_path(candidate_output_path, false)?;
    ensure!(
        candidate_output_path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx")),
        "S3-F03 候选输出必须使用 .xlsx"
    );
    ensure!(
        !candidate_output_path.exists(),
        "候选输出路径已存在，未覆盖：{}",
        candidate_output_path.display()
    );
    ensure!(
        plan_output_path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("json")),
        "S3-F03 计划输出必须使用 .json"
    );
    ensure!(
        !plan_output_path.exists(),
        "导入计划路径已存在，未覆盖：{}",
        plan_output_path.display()
    );
    ensure_distinct_paths(&[target_path, candidate_output_path, plan_output_path])?;
    for feedback_path in feedback_paths {
        ensure_workbook_path(feedback_path, true)?;
        ensure!(
            feedback_path
                .extension()
                .and_then(|extension| extension.to_str())
                .is_some_and(|extension| {
                    extension.eq_ignore_ascii_case("xlsx") || extension.eq_ignore_ascii_case("xlsm")
                }),
            "S3-F03 反馈工作簿目前只支持 .xlsx 和只读 .xlsm：{}",
            feedback_path.display()
        );
    }

    let snapshot_paths = std::iter::once(target_path.to_path_buf())
        .chain(feedback_paths.iter().cloned())
        .collect::<Vec<_>>();
    let before_hashes = snapshot_paths
        .iter()
        .map(|path| sha256_file(path))
        .collect::<Result<Vec<_>>>()?;

    let target_read_started = Instant::now();
    let target_sheets = read_value_columns_with_iso_dates(target_path, usize::MAX)
        .with_context(|| format!("读取目标说明工作簿：{}", target_path.display()))?;
    let target_read_time_ms = target_read_started.elapsed().as_millis();
    let target_sheet = find_explanation_sheet(target_sheets)?;
    let target_headers = target_sheet.headers()?;

    let mut targets = HashMap::<Identity, Vec<TargetRow>>::new();
    for row_number in target_headers.row_number + 1..=target_sheet.max_row {
        let key = identity(&target_sheet, row_number, &target_headers);
        if identity_is_nonempty(&key) {
            targets.entry(key).or_default().push(TargetRow {
                row_number,
                organization_name: target_sheet
                    .cell(row_number, target_headers.indexes["机构名称"]),
                indicator_name: target_sheet.cell(row_number, target_headers.indexes["指标名称"]),
            });
        }
    }

    let feedback_read_started = Instant::now();
    let mut feedback_order = Vec::<Identity>::new();
    let mut feedback_by_key = HashMap::<Identity, BTreeMap<String, BTreeSet<String>>>::new();
    for feedback_path in feedback_paths {
        let sheets = read_value_columns_with_iso_dates(feedback_path, usize::MAX)
            .with_context(|| format!("读取机构反馈工作簿：{}", feedback_path.display()))?;
        let sheet = find_explanation_sheet(sheets)?;
        let headers = sheet.headers()?;
        let source_name = feedback_path
            .file_name()
            .unwrap_or_default()
            .to_string_lossy()
            .into_owned();
        let reason_column = headers.indexes[REASON_HEADER];
        for row_number in headers.row_number + 1..=sheet.max_row {
            let key = identity(&sheet, row_number, &headers);
            let reason_cell = sheet.cell(row_number, reason_column);
            let reason = if python_truthy(&reason_cell) {
                python_string(&reason_cell).trim().to_owned()
            } else {
                String::new()
            };
            if !identity_is_nonempty(&key) || reason.is_empty() {
                continue;
            }
            let reasons = feedback_by_key.entry(key.clone()).or_insert_with(|| {
                feedback_order.push(key.clone());
                BTreeMap::new()
            });
            reasons
                .entry(reason)
                .or_default()
                .insert(source_name.clone());
        }
    }
    let feedback_read_time_ms = feedback_read_started.elapsed().as_millis();

    let mut imported = 0usize;
    let mut unmatched = 0usize;
    let mut conflicts = 0usize;
    let mut operations = Vec::new();
    let mut conflict_rows = Vec::new();
    let target_reason_column = target_headers.indexes[REASON_HEADER];
    for key in feedback_order {
        let reasons = feedback_by_key
            .get(&key)
            .expect("feedback order and index are populated together");
        let Some(target_rows) = targets.get(&key) else {
            unmatched += 1;
            continue;
        };
        if reasons.len() != 1 || target_rows.len() != 1 {
            conflicts += 1;
            let reasons_text = reasons
                .iter()
                .map(|(reason, sources)| {
                    format!(
                        "{reason}（来源：{}）",
                        sources.iter().cloned().collect::<Vec<_>>().join("、")
                    )
                })
                .collect::<Vec<_>>()
                .join("\n");
            let mut causes = Vec::new();
            if target_rows.len() != 1 {
                causes.push(format!("目标说明文件存在 {} 行相同标识", target_rows.len()));
            }
            if reasons.len() != 1 {
                causes.push("多个反馈文件给出不同变动原因".to_owned());
            }
            let sample = &target_rows[0];
            conflict_rows.push(ConflictRow([
                "冲突".to_owned(),
                key.0.clone(),
                python_str_for_conflict(&sample.organization_name),
                key.1.clone(),
                python_str_for_conflict(&sample.indicator_name),
                key.2.trim_start_matches('#').to_owned(),
                key.3.trim_start_matches('#').to_owned(),
                target_rows
                    .iter()
                    .map(|row| row.row_number.to_string())
                    .collect::<Vec<_>>()
                    .join("、"),
                reasons_text,
                causes.join("；"),
            ]));
            continue;
        }
        let reason = reasons.keys().next().expect("one reason exists").clone();
        operations.push(EditOperation::SetCellValue {
            sheet: target_sheet.name.clone(),
            cell: cell_address(target_rows[0].row_number, target_reason_column),
            value: EditCellValue::Text(reason),
        });
        imported += 1;
    }

    let after_hashes = snapshot_paths
        .iter()
        .map(|path| sha256_file(path))
        .collect::<Result<Vec<_>>>()?;
    ensure!(
        before_hashes == after_hashes,
        "规划期间输入工作簿发生变化，未写出导入计划"
    );

    let candidate_path = normalize_new_path(candidate_output_path)?;
    let plan_path = normalize_new_path(plan_output_path)?;
    ensure_distinct_paths(&[target_path, &candidate_path, &plan_path])?;
    let conflict_report_path = if conflict_rows.is_empty() {
        None
    } else {
        Some(conflict_report_path(target_path)?)
    };

    let edit_plan = if operations.is_empty() {
        None
    } else {
        let profile = inspect_xlsx_profile(target_path)
            .with_context(|| format!("检查目标工作簿结构：{}", target_path.display()))?;
        ensure!(
            profile.format == WorkbookFormat::Xlsx,
            "S3-F03 ExcelEditPlan 目前仅支持 XLSX 目标"
        );
        let request_id = format!(
            "s3-f03-{}-{}",
            Local::now().format("%Y%m%d%H%M%S"),
            PLAN_COUNTER.fetch_add(1, Ordering::Relaxed)
        );
        let plan = ExcelEditPlan {
            request_id,
            input: target_path.to_path_buf(),
            output: candidate_path.clone(),
            format: WorkbookFormat::Xlsx,
            workbook: profile,
            operations,
            preserve: required_preservation(),
            drop_allowed: BTreeMap::new(),
        };
        plan.validate().context("校验 S3-F03 ExcelEditPlan")?;
        Some(plan)
    };

    let snapshots = snapshot_paths
        .iter()
        .zip(after_hashes)
        .map(|(path, sha256)| InputSnapshot {
            path: path.clone(),
            sha256,
        })
        .collect::<Vec<_>>();
    let mut snapshots = snapshots.into_iter();
    let target_snapshot = snapshots.next().expect("target snapshot exists");
    let feedback_snapshots = snapshots.collect::<Vec<_>>();
    let document = PlanDocument {
        schema_version: 1,
        feature: "S3-F03",
        status: "planned_not_applied",
        created_at: Local::now().to_rfc3339(),
        applied: false,
        target: target_snapshot,
        feedback_sources: feedback_snapshots,
        candidate_output_workbook: candidate_path.clone(),
        counts: PlanCounts {
            imported,
            unmatched,
            conflicts,
            planned_cell_updates: imported,
        },
        conflict_report_path: conflict_report_path.clone(),
        edit_plan,
    };

    let conflict_write_started = Instant::now();
    let conflict_temp = if let Some(path) = &conflict_report_path {
        let temporary = reserve_temporary_path(path, "xlsx")?;
        if let Err(error) = write_conflict_workbook(&temporary, &conflict_rows) {
            let _ = fs::remove_file(&temporary);
            return Err(error);
        }
        Some(temporary)
    } else {
        None
    };
    let conflict_write_time_ms = if conflict_temp.is_some() {
        conflict_write_started.elapsed().as_millis()
    } else {
        0
    };

    let plan_write_started = Instant::now();
    let plan_temp = reserve_temporary_path(&plan_path, "json")?;
    let plan_write_result = (|| -> Result<()> {
        let mut file = OpenOptions::new()
            .write(true)
            .truncate(true)
            .open(&plan_temp)
            .with_context(|| format!("打开临时导入计划：{}", plan_temp.display()))?;
        serde_json::to_writer_pretty(&mut file, &document).context("序列化 S3-F03 导入计划")?;
        file.write_all(b"\n").context("写入导入计划换行")?;
        file.sync_all().context("刷新导入计划到磁盘")?;
        Ok(())
    })();
    if let Err(error) = plan_write_result {
        let _ = fs::remove_file(&plan_temp);
        if let Some(temporary) = &conflict_temp {
            let _ = fs::remove_file(temporary);
        }
        return Err(error);
    }

    if let (Some(temporary), Some(final_path)) = (&conflict_temp, &conflict_report_path) {
        if let Err(error) = publish_without_overwrite(temporary, final_path) {
            let _ = fs::remove_file(&plan_temp);
            return Err(error);
        }
    }
    if let Err(error) = publish_without_overwrite(&plan_temp, &plan_path) {
        if let Some(final_path) = &conflict_report_path {
            let _ = fs::remove_file(final_path);
        }
        return Err(error);
    }
    let plan_write_time_ms = plan_write_started.elapsed().as_millis();

    Ok(S3ExplanationFeedbackPlanSummary {
        status: "planned_not_applied",
        applied: false,
        target_workbook: target_path.to_path_buf(),
        plan_file: plan_path,
        candidate_output_workbook: candidate_path,
        imported,
        unmatched,
        conflicts,
        planned_cell_updates: imported,
        conflict_report_path,
        reader: READER,
        business_engine: BUSINESS_ENGINE,
        conflict_writer: (!conflict_rows.is_empty()).then_some(CONFLICT_WRITER),
        target_read_time_ms,
        feedback_read_time_ms,
        conflict_write_time_ms,
        plan_write_time_ms,
        total_time_ms: started.elapsed().as_millis(),
    })
}

fn ensure_workbook_path(path: &Path, must_exist: bool) -> Result<()> {
    if must_exist {
        ensure!(path.is_file(), "工作簿不存在：{}", path.display());
    } else if path.exists() {
        bail!("目标路径已存在：{}", path.display());
    }
    Ok(())
}

fn ensure_distinct_paths(paths: &[&Path]) -> Result<()> {
    let identities = paths
        .iter()
        .map(|path| normalized_identity(path))
        .collect::<Result<BTreeSet<_>>>()?;
    ensure!(
        identities.len() == paths.len(),
        "输入、候选输出和计划输出必须使用不同路径"
    );
    Ok(())
}

fn normalized_identity(path: &Path) -> Result<String> {
    let absolute = if path.exists() {
        strip_windows_verbatim_prefix(
            fs::canonicalize(path).with_context(|| format!("规范化路径 {}", path.display()))?,
        )
    } else {
        normalize_new_path(path)?
    };
    let mut value = absolute.to_string_lossy().replace('/', "\\");
    if cfg!(windows) {
        value.make_ascii_lowercase();
    }
    Ok(value)
}

fn normalize_new_path(path: &Path) -> Result<PathBuf> {
    let parent = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or(Path::new("."));
    let name = path.file_name().context("输出路径必须包含文件名")?;
    let absolute_parent = if parent.exists() {
        strip_windows_verbatim_prefix(
            fs::canonicalize(parent)
                .with_context(|| format!("规范化输出目录 {}", parent.display()))?,
        )
    } else {
        std::env::current_dir()
            .context("获取当前目录")?
            .join(parent)
    };
    Ok(absolute_parent.join(name))
}

fn strip_windows_verbatim_prefix(path: PathBuf) -> PathBuf {
    #[cfg(windows)]
    {
        let value = path.to_string_lossy();
        if let Some(rest) = value.strip_prefix("\\\\?\\") {
            if let Some(unc) = rest.strip_prefix("UNC\\") {
                return PathBuf::from(format!("\\\\{unc}"));
            }
            return PathBuf::from(rest);
        }
    }
    path
}

fn sha256_file(path: &Path) -> Result<String> {
    let mut file =
        File::open(path).with_context(|| format!("打开输入工作簿 {}", path.display()))?;
    let mut digest = Sha256::new();
    let mut buffer = [0u8; 64 * 1024];
    loop {
        let read = file
            .read(&mut buffer)
            .with_context(|| format!("读取输入工作簿 {}", path.display()))?;
        if read == 0 {
            break;
        }
        digest.update(&buffer[..read]);
    }
    Ok(format!("{:x}", digest.finalize()))
}

fn find_explanation_sheet(sheets: Vec<SheetValues>) -> Result<ValueGrid> {
    let grids = sheets.into_iter().map(ValueGrid::from).collect::<Vec<_>>();
    if let Some(sheet) = grids.iter().find(|sheet| sheet.name == "说明文件") {
        return Ok(sheet.clone());
    }
    grids
        .into_iter()
        .find(ValueGrid::has_institution_code_header)
        .context("未找到说明文件工作表或“机构代码”表头")
}

impl ValueGrid {
    fn has_institution_code_header(&self) -> bool {
        self.has_header_text("机构代码")
    }
}

fn identity(sheet: &ValueGrid, row: usize, headers: &HeaderColumns) -> Identity {
    (
        code_value(&sheet.cell(row, headers.indexes["机构代码"])),
        code_value(&sheet.cell(row, headers.indexes["指标代码"])),
        amount_value(&sheet.cell(row, headers.indexes["数据值"])),
        amount_value(&sheet.cell(row, headers.indexes["增减额"])),
    )
}

fn identity_is_nonempty(key: &Identity) -> bool {
    key.0.len() > 0 || key.1.len() > 0 || key.2.len() > 0 || key.3.len() > 0
}

fn code_value(value: &CellValue) -> String {
    let mut text = python_string(value).trim().replace('\'', "");
    if let Some(digits) = text.strip_suffix(".0") {
        if !digits.is_empty() && digits.bytes().all(|byte| byte.is_ascii_digit()) {
            text = digits.to_owned();
        }
    }
    text.to_uppercase()
}

fn amount_value(value: &CellValue) -> String {
    if matches!(value, CellValue::Empty) {
        return String::new();
    }
    let text = python_string(value).trim().to_owned();
    match quantize_half_up_four_places(&text) {
        DecimalQuantization::Rounded(value) => format!("#{value}"),
        DecimalQuantization::Invalid(value) => value,
        DecimalQuantization::NotDecimal => text,
    }
}

enum DecimalQuantization {
    Rounded(String),
    Invalid(String),
    NotDecimal,
}

fn quantize_half_up_four_places(text: &str) -> DecimalQuantization {
    if text.eq_ignore_ascii_case("nan") || text.eq_ignore_ascii_case("+nan") {
        return DecimalQuantization::Rounded("NaN".to_owned());
    }
    if text.eq_ignore_ascii_case("-nan") {
        return DecimalQuantization::Rounded("-NaN".to_owned());
    }
    if text.eq_ignore_ascii_case("infinity") || text.eq_ignore_ascii_case("+infinity") {
        return DecimalQuantization::Invalid("Infinity".to_owned());
    }
    if text.eq_ignore_ascii_case("-infinity") {
        return DecimalQuantization::Invalid("-Infinity".to_owned());
    }

    let (negative, unsigned) = if let Some(rest) = text.strip_prefix('-') {
        (true, rest)
    } else if let Some(rest) = text.strip_prefix('+') {
        (false, rest)
    } else {
        (false, text)
    };
    if unsigned.is_empty() {
        return DecimalQuantization::NotDecimal;
    }
    let mut exponent_parts = unsigned.split(['e', 'E']);
    let mantissa = exponent_parts.next().unwrap_or_default();
    let exponent_text = exponent_parts.next();
    if exponent_parts.next().is_some() {
        return DecimalQuantization::NotDecimal;
    }
    let exponent = match exponent_text {
        Some(value) => match value.parse::<i32>() {
            Ok(value) => value,
            Err(_) => return DecimalQuantization::NotDecimal,
        },
        None => 0,
    };
    let mut decimal_parts = mantissa.split('.');
    let integer_part = decimal_parts.next().unwrap_or_default();
    let fractional_part = decimal_parts.next().unwrap_or_default();
    if decimal_parts.next().is_some()
        || (!integer_part.bytes().all(|byte| byte.is_ascii_digit()))
        || (!fractional_part.bytes().all(|byte| byte.is_ascii_digit()))
        || (integer_part.is_empty() && fractional_part.is_empty())
    {
        return DecimalQuantization::NotDecimal;
    }
    let coefficient = format!("{integer_part}{fractional_part}");
    let digits = coefficient.trim_start_matches('0');
    if digits.is_empty() {
        return DecimalQuantization::Rounded(if negative {
            "-0.0000".to_owned()
        } else {
            "0.0000".to_owned()
        });
    }

    let shift = exponent - fractional_part.len() as i32 + 4;
    if shift > 10_000 || shift < -10_000 {
        return DecimalQuantization::Invalid(decimal_to_string(text));
    }
    let mut scaled = if shift >= 0 {
        format!("{digits}{}", "0".repeat(shift as usize))
    } else {
        let discarded = (-shift) as usize;
        let keep = digits.len().saturating_sub(discarded);
        let mut retained = if keep == 0 {
            "0".to_owned()
        } else {
            digits[..keep].to_owned()
        };
        let first_discarded = if keep < digits.len() && discarded <= digits.len() {
            digits.as_bytes()[keep]
        } else {
            b'0'
        };
        if first_discarded >= b'5' {
            increment_decimal_digits(&mut retained);
        }
        retained
    };
    let precision = scaled.trim_start_matches('0').len().max(1);
    if precision > 28 {
        return DecimalQuantization::Invalid(decimal_to_string(text));
    }
    if scaled.len() <= 4 {
        scaled = format!("{}{}", "0".repeat(5 - scaled.len()), scaled);
    }
    let split = scaled.len() - 4;
    let integer = scaled[..split].trim_start_matches('0');
    let integer = if integer.is_empty() { "0" } else { integer };
    let sign = if negative { "-" } else { "" };
    DecimalQuantization::Rounded(format!("{sign}{integer}.{}", &scaled[split..]))
}

fn increment_decimal_digits(value: &mut String) {
    let mut bytes = value.as_bytes().to_vec();
    let mut index = bytes.len();
    while index > 0 {
        index -= 1;
        if bytes[index] < b'9' {
            bytes[index] += 1;
            *value = String::from_utf8(bytes).expect("decimal digits remain UTF-8");
            return;
        }
        bytes[index] = b'0';
    }
    bytes.insert(0, b'1');
    *value = String::from_utf8(bytes).expect("decimal digits remain UTF-8");
}

fn decimal_to_string(text: &str) -> String {
    if let Some(index) = text.find(['e', 'E']) {
        let (mantissa, exponent) = text.split_at(index);
        let exponent = exponent[1..].parse::<i32>().unwrap_or(0);
        let sign = if exponent >= 0 { "+" } else { "" };
        format!("{}E{sign}{exponent}", mantissa.trim_start_matches('+'))
    } else {
        text.to_owned()
    }
}

fn python_string(value: &CellValue) -> String {
    match value {
        CellValue::Empty => String::new(),
        CellValue::Text(value) => value.clone(),
        CellValue::Number(value) | CellValue::DateTimeSerial(value) => python_float_string(*value),
        CellValue::Integer(value) => value.to_string(),
        CellValue::Boolean(value) => if *value { "True" } else { "False" }.to_owned(),
        CellValue::DateTimeIso(value) => value.replace('T', " "),
        CellValue::DurationIso(value) => value.clone(),
        CellValue::Error(value) => format_excel_error(value).to_owned(),
    }
}

fn python_str_for_conflict(value: &CellValue) -> String {
    if matches!(value, CellValue::Empty) {
        "None".to_owned()
    } else {
        python_string(value)
    }
}

fn python_float_string(value: f64) -> String {
    if value.is_nan() {
        return "nan".to_owned();
    }
    if value == f64::INFINITY {
        return "inf".to_owned();
    }
    if value == f64::NEG_INFINITY {
        return "-inf".to_owned();
    }
    let value = format!("{value:?}");
    let Some((mantissa, exponent)) = value.split_once('e') else {
        return value;
    };
    let mantissa = mantissa.strip_suffix(".0").unwrap_or(mantissa);
    let exponent = exponent.parse::<i32>().unwrap_or(0);
    let sign = if exponent >= 0 { "+" } else { "" };
    let exponent_digits = exponent.abs().to_string();
    let exponent_digits = if exponent_digits.len() < 2 {
        format!("0{exponent_digits}")
    } else {
        exponent_digits
    };
    format!(
        "{mantissa}e{sign}{}{exponent_digits}",
        if exponent < 0 { "-" } else { "" }
    )
}

fn python_truthy(value: &CellValue) -> bool {
    match value {
        CellValue::Empty => false,
        CellValue::Text(value) => !value.is_empty(),
        CellValue::Number(value) | CellValue::DateTimeSerial(value) => *value != 0.0,
        CellValue::Integer(value) => *value != 0,
        CellValue::Boolean(value) => *value,
        CellValue::DateTimeIso(value) | CellValue::DurationIso(value) => !value.is_empty(),
        CellValue::Error(value) => !value.is_empty(),
    }
}

fn format_excel_error(value: &str) -> &str {
    let normalized = value
        .trim()
        .trim_start_matches('#')
        .trim_end_matches(['!', '?'])
        .replace('_', "")
        .replace('-', "")
        .replace(' ', "")
        .to_ascii_lowercase();
    match normalized.as_str() {
        "ref" => "#REF!",
        "value" => "#VALUE!",
        "na" | "notavailable" => "#N/A",
        "div0" | "divzero" => "#DIV/0!",
        "name" => "#NAME?",
        "num" => "#NUM!",
        "null" => "#NULL!",
        "spill" => "#SPILL!",
        "calc" => "#CALC!",
        _ => value,
    }
}

fn parse_cell_address(address: &str) -> Option<(usize, usize)> {
    let bytes = address.as_bytes();
    let split = bytes.iter().position(|byte| byte.is_ascii_digit())?;
    if split == 0
        || !bytes[..split].iter().all(u8::is_ascii_alphabetic)
        || !bytes[split..].iter().all(|byte| byte.is_ascii_digit())
    {
        return None;
    }
    let column = bytes[..split].iter().try_fold(0usize, |column, byte| {
        column
            .checked_mul(26)?
            .checked_add((byte.to_ascii_uppercase() - b'A' + 1) as usize)
    })?;
    let row = address[split..].parse::<usize>().ok()?;
    (row > 0 && column > 0).then_some((row, column))
}

fn cell_address(row: usize, column: usize) -> String {
    let mut column = column;
    let mut label = String::new();
    while column > 0 {
        let remainder = (column - 1) % 26;
        label.insert(0, char::from(b'A' + remainder as u8));
        column = (column - 1) / 26;
    }
    format!("{label}{row}")
}

fn required_preservation() -> BTreeSet<PreservationRequirement> {
    [
        PreservationRequirement::CellValuesAndTypes,
        PreservationRequirement::FormulaText,
        PreservationRequirement::CommentsAndPositions,
        PreservationRequirement::Styles,
        PreservationRequirement::MergedRanges,
        PreservationRequirement::ConditionalFormatting,
        PreservationRequirement::DataValidations,
        PreservationRequirement::DefinedNames,
        PreservationRequirement::RowAndColumnDimensions,
        PreservationRequirement::PageSettings,
        PreservationRequirement::HeaderFooterText,
        PreservationRequirement::Relationships,
        PreservationRequirement::UnmodifiedPackageParts,
        PreservationRequirement::SheetIdentityAndOrder,
        PreservationRequirement::FreezePanes,
        PreservationRequirement::AutoFilter,
    ]
    .into_iter()
    .collect()
}

fn conflict_report_path(target_path: &Path) -> Result<PathBuf> {
    let parent = target_path
        .parent()
        .filter(|path| !path.as_os_str().is_empty())
        .unwrap_or(Path::new("."));
    let stem = target_path
        .file_stem()
        .and_then(|value| value.to_str())
        .unwrap_or("S3-F03");
    Ok(parent.join(format!(
        "{stem}_导入冲突清单_{}.xlsx",
        Local::now().format("%Y%m%d_%H%M%S")
    )))
}

fn write_conflict_workbook(path: &Path, rows: &[ConflictRow]) -> Result<()> {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();
    sheet
        .set_name("冲突明细")
        .context("设置 S3-F03 冲突表名称")?;
    let header_format = Format::new()
        .set_bold()
        .set_font_color(Color::RGB(0xFF_FF_FF))
        .set_background_color(Color::RGB(0xC0_00_00));
    for (column, header) in CONFLICT_HEADERS.iter().enumerate() {
        sheet.write_string_with_format(0, column as u16, *header, &header_format)?;
    }
    for (index, row) in rows.iter().enumerate() {
        for (column, value) in row.0.iter().enumerate() {
            sheet.write_string((index + 1) as u32, column as u16, value)?;
        }
    }
    sheet
        .set_freeze_panes(1, 0)
        .context("冻结 S3-F03 冲突表表头")?;
    sheet
        .autofilter(0, 0, rows.len() as u32, 9)
        .context("设置 S3-F03 冲突表筛选")?;
    for (column, width) in CONFLICT_WIDTHS.iter().enumerate() {
        sheet
            .set_column_width_pixels(column as u16, (width * 7.0).round() as u32)
            .with_context(|| format!("设置 S3-F03 冲突清单第 {} 列宽", column + 1))?;
    }
    workbook.save(path).context("保存 S3-F03 冲突清单")
}

fn reserve_temporary_path(output: &Path, extension: &str) -> Result<PathBuf> {
    let parent = output
        .parent()
        .filter(|path| !path.as_os_str().is_empty())
        .unwrap_or(Path::new("."));
    fs::create_dir_all(parent).with_context(|| format!("创建输出目录：{}", parent.display()))?;
    let stem = output
        .file_stem()
        .and_then(|value| value.to_str())
        .unwrap_or("s3-f03");
    for _ in 0..100 {
        let sequence = TEMP_COUNTER.fetch_add(1, Ordering::Relaxed);
        let candidate = parent.join(format!(
            ".{stem}.{}.{}.tmp.{extension}",
            std::process::id(),
            sequence
        ));
        match OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&candidate)
        {
            Ok(_) => return Ok(candidate),
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(error) => {
                return Err(error).with_context(|| format!("创建临时输出 {}", candidate.display()))
            }
        }
    }
    bail!("无法为输出保留唯一临时路径：{}", output.display())
}

fn publish_without_overwrite(temporary: &Path, output: &Path) -> Result<()> {
    fs::hard_link(temporary, output).with_context(|| {
        format!(
            "安全发布输出失败（目标已存在或文件系统不支持硬链接）：{}",
            output.display()
        )
    })?;
    fs::remove_file(temporary).with_context(|| format!("清理临时文件：{}", temporary.display()))
}

#[cfg(test)]
mod tests {
    use super::{amount_value, code_value, quantize_half_up_four_places, DecimalQuantization};
    use excel_reader::CellValue;

    #[test]
    fn s3_feedback_identity_normalizes_codes_and_half_up_amounts() {
        assert_eq!(
            code_value(&CellValue::Text(" '6k0u.0 ".to_owned())),
            "6K0U.0"
        );
        assert_eq!(code_value(&CellValue::Number(7020.0)), "7020");
        assert_eq!(amount_value(&CellValue::Number(0.418462)), "#0.4185");
        assert_eq!(
            amount_value(&CellValue::Text("-0.00005".to_owned())),
            "#-0.0001"
        );
        assert_eq!(
            amount_value(&CellValue::Text("0.000049".to_owned())),
            "#0.0000"
        );
        assert_eq!(
            amount_value(&CellValue::Text("-0.000049".to_owned())),
            "#-0.0000"
        );
        assert_eq!(
            amount_value(&CellValue::Text("description".to_owned())),
            "description"
        );
    }

    #[test]
    fn decimal_quantization_handles_leading_zeroes_and_exponents() {
        assert!(matches!(
            quantize_half_up_four_places("5e-5"),
            DecimalQuantization::Rounded(value) if value == "0.0001"
        ));
        assert!(matches!(
            quantize_half_up_four_places("5e-6"),
            DecimalQuantization::Rounded(value) if value == "0.0000"
        ));
        assert!(matches!(
            quantize_half_up_four_places("1.23455e2"),
            DecimalQuantization::Rounded(value) if value == "123.4550"
        ));
    }
}
